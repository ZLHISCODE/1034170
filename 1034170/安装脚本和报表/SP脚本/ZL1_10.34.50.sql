--[连续升级]1
--[管理工具版本号]10.34.30
--本脚本支持从ZLHIS+ v10.34.40 升级到 v10.34.50
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--91029:黄捷,2015-11-26,修改影像报告操作记录
alter table 影像报告操作记录 add 文档标题 VARCHAR2(60);

alter table 影像报告操作记录 drop constraint 影像报告操作记录_FK_医嘱ID;
alter table 影像报告操作记录 add constraint 影像报告操作记录_FK_医嘱ID foreign key (医嘱ID) references 病人医嘱记录 (ID) On Delete Cascade Enable Novalidate;
alter table 影像报告操作记录 drop constraint 影像报告操作记录_FK_报告ID;

--89736:张德婷,2015-11-24,输液配置中心收费标准
alter table  输液配药附费
       add 病人id number(18);  

--89736:张德婷,2015-11-24,输液配置中心收费标准
Create table 配置收费方案( 
  序号  NUMBER(18) not null,
  配药类型  varchar2(50),
  项目id  NUMBER(18),
  收费项目 varchar2(100))
  tablespace ZL9BASEITEM;


Alter Table 配置收费方案 Add Constraint 配置收费方案_UQ_配药类型 Unique (配药类型,项目id) Using Index Tablespace zl9Indexhis;

--90603:许华峰,2015-11-18,同一个原型能否书写多份报告控制
Alter Table 影像报告原型清单 Add 可否书写多份 Number(1);

--90466:刘尔旋,2015-11-12,支付宝结帐修改
CREATE TABLE 三方退款信息(
  结帐ID Number(18),
  记录ID Number(18),
  金额 Number(16,5),
  卡号 Varchar2(50),
  交易流水号 Varchar2(50),
  交易说明 Varchar2(500),
  待转出 Number(3))
  TABLESPACE zl9Expense;

Alter Table 三方退款信息 Add Constraint 三方退款信息_PK Primary Key (结帐ID,记录ID) Using Index Tablespace zl9Indexhis;

Create Index 三方退款信息_IX_待转出 On 三方退款信息(待转出) Tablespace zl9Indexhis;

--90308:冉俊明,2015-11-06,性能问题，补加索引。
Create Index 医保结算明细_IX_NO On 医保结算明细(NO) Pctfree 5 Tablespace zl9Indexhis;

--90338:许华峰,2015-11-06,排队叫号数据过滤慢
Create Index 排队叫号队列_IX_队列名称 On 排队叫号队列(队列名称) Tablespace zl9Indexhis;

--88036:李小东,2015-11-02,新增药品特性属性字段(是否原研药、是否专利药、是否单独定价)
Alter Table 药品特性 Add 是否原研药 Number(1);

Alter Table 药品特性 Add 是否专利药 Number(1);

Alter Table 药品特性 Add 是否单独定价 Number(1);

--89077:许华锋,2015-11-02,报告片段适应条件
Alter Table 影像报告片段清单 Add 适应条件 Xmltype;

---90823:刘硕,2015-11-23,结构化地址增加虚拟地址识别
--89238:刘硕,2015-10-23,结构化地址支持第四级字典输入
Alter Table 区域 Modify 名称 Varchar2(100); 
Alter Table 区域 Modify 简码 Varchar2(100); 
Alter Table 区域 Modify 编码 Varchar2(15); 
Alter Table 区域 Modify 上级编码 Varchar2(15); 
alter table 区域 add 是否虚拟 number(1);
alter table 区域 add 是否不显示 number(1);
alter table 区域 add 五笔码 varchar2(100);
Alter Table 区域 Drop Constraint 区域_UQ_名称 Cascade Drop Index ; 
Alter Table 区域 Add Constraint 区域_UQ_名称 Unique (名称,上级编码) Using Index Tablespace zl9Indexhis;
Alter Table 区域 Drop Constraint 区域_FK_上级编码;
Alter Table 区域 Add Constraint 区域_FK_上级编码 Foreign Key (上级编码) References 区域(编码) On Delete Cascade;
Alter Table 区域 Modify 名称 Constraint 区域_NN_名称 Not Null;
Create Index 区域_IX_上级编码 On 区域(上级编码) Tablespace zl9Indexhis;
---90823:刘硕,2015-11-23,结构化地址增加虚拟地址识别
--90021:刘硕,2015-11-06,病人地址结构化录入支持乡镇级
alter table 病人地址信息 add(乡镇 Varchar2(100),区划代码 Varchar2(15));
Alter Table 病人地址信息 Drop Constraint 病人地址信息_PK Cascade Drop Index ; 
Alter Table 病人地址信息 Add Constraint 病人地址信息_UQ_地址类别 Unique (病人ID,主页ID,地址类别) Using Index Tablespace zl9Indexhis;
Alter Table 病人地址信息 Modify 病人ID Constraint 病人地址信息_NN_病人ID Not Null;
Alter Table 病人地址信息 Modify 地址类别 Constraint 病人地址信息_NN_地址类别 Not Null;
Alter Table 病人地址信息 Add Constraint 病人地址信息_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID) On Delete Cascade;
-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--89736:张德婷,2015-11-24,输液配置中心收费标准
Declare

  n_Count      Number(4);
Begin
  Select Count(特定项目) Into n_Count From 收费特定项目 Where 特定项目 = '肿瘤配置费' Or 特定项目 = '普通配置费';

  If n_Count <> 0 Then
    Select Count(编码) Into n_Count From 输液配药类型;
    Insert Into 输液配药类型
      (编码, 名称, 简码)
    Values
      (Substr('0000', Length(n_Count) + 1) || n_Count + 1, '肿瘤药物', 'ZLYW');
    Insert Into 输液配药类型
      (编码, 名称, 简码)
    Values
      (Substr('0000', Length(n_Count) + 1) || n_Count + 2, '普通药物', 'PTYW');
  
    Update 输液药品属性
    Set 配药类型 =
         ((Substr('0000', Length(n_Count) + 1) || n_Count + 1) || '-肿瘤药物')
    Where 药品id In(select 药品id From 药品规格
    Where 药名id In (Select 药名id From 药品特性 Where 是否肿瘤药 = 1));
    Update 输液药品属性
    Set 配药类型 =
         ((Substr('0000', Length(n_Count) + 1) || n_Count + 2) || '-普通药物')
    Where 配药类型 Is Null;
    
    n_Count:=0;
    For r_Item In (Select Decode(a.特定项目, '肿瘤配置费', '肿瘤药物', '普通药物') 配药类型, a.收费细目id, b.名称
                   From 收费特定项目 A, 收费项目目录 B
                   Where A.收费细目id=B.id and (特定项目 = '肿瘤配置费' Or 特定项目 = '普通配置费')) Loop
      n_Count:=n_Count+1;
      Insert Into 配置收费方案 Values (n_Count, r_Item.配药类型, r_Item.收费细目id, r_Item.名称);
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--89305:张德婷,2015-11-25,扫两次瓶签自动发药
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 1, 0, 0, 25, '扫两次瓶签号自动发送', Null, '0', '当瓶签号扫描两次之后自动发送'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '扫两次瓶签号自动发送');

--89736:张德婷,2015-11-24,输液配置中心收费标准
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 26,'配置费按病人收取', Null, '0', '配置费按病人收取，一个病人一天只收一次'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '配置费按病人收取');
  
--76482:李南春,2015-11-17,自助挂号三方卡自动签约
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1802, 0, 0, 0, 0, 38, '自动签约', Null, Null, '三方卡是否自动签约'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 参数号 = 38 And 系统 = &n_System And 模块 = 1802);

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1803, 0, 0, 0, 0, 38, '自动签约', Null, Null, '三方卡是否自动签约'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 参数号 = 38 And 系统 = &n_System And 模块 = 1803);

--90466:刘尔旋,2015-11-16,结帐管理支付宝修改
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1137, 0, 0, 0, 0, 47,'预交票据打印方式', '', '', '预交票据打印方式,0-不打印 1-自动打印 2-选择打印'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1137 And 参数名 = '预交票据打印方式');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1137, 0, 1, 0, 0, 46,'共用预交票据批次', '', '', '操作员共用的预交票据的批次,格式为:领用ID1,预交类型ID1|领用ID2,预交类型2|.....'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1137 And 参数名 = '共用预交票据批次');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1137, 1, 1, 0, 0, 48,'当前预交票据号', '', '', '当前预交票据号'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1137 And 参数名 = '当前预交票据号');

--90466:刘尔旋,2015-11-12,支付宝结帐修改
Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
Select &n_System,1, '三方退款信息',8,1,-NULL From Dual;

--90021:刘硕,2015-11-05,病人地址结构化录入支持乡镇级
Insert Into Zlparameters
  (Id, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, Null, 私有, 本机, 授权, 固定, 251, '病人地址结构化录入', 参数值, '0',
         '如果启用此参数,病人可以录入病人出生地点、籍贯、现住址、户口地址、联系人地址、单位地址的地方将采用结构化录入，即：分别录入省(直辖市)、市(县)、镇(区）、详细。'
  From Zlparameters
  Where 参数名 = '病人地址结构化录入' And Nvl(模块, 0) = 1261 And Nvl(系统, 0) = &n_System And Not Exists
   (Select 1
         From Zlparameters
         Where 参数名 = '病人地址结构化录入' And Nvl(模块, 0) = 0 And Nvl(系统, 0) = &n_System);

Delete Zlparameters Where 参数名 = '病人地址结构化录入' And Nvl(模块, 0) = 1261 And Nvl(系统, 0) = &n_System;

Insert Into Zlparameters
  (Id, 系统, 模块, 私有, 本机, 授权, 固定,  参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, Null, 0, 0, 0, 0, 252, '乡镇地址结构化录入', Null, Null,
         '地址可以结构录入省(直辖市)、市(县)、镇(区）、详细。启用该参数后，将展示为省(直辖市)、市(县)、镇(区）、乡镇、详细。当前地址输入为省(直辖市)、市(县)、镇(区）、乡镇、详细时，不启用该参数，则展示为：省(直辖市)、市(县)、镇(区）、详细'
  From Dual
  Where Not Exists (Select 1
         From Zlparameters
         Where 参数名 = '乡镇地址结构化录入' And Nvl(模块, 0) = 0 And Nvl(系统, 0) = &n_System);

--90217:黄捷,2015-11-03,PACS报告文档编辑器嵌入式报告增加tab
--报告主窗体参数
Insert Into 影像参数说明(ID,PID,系统,模块,分组,参数序号,参数名,默认值,参数级别,取值范围,启用条件,说明)
Values(sys_guid(),'',100,'1290',null,1,'是否显示报告列表','1',2,null,null,'根据此参数配置嵌入式报告中是否显示报告列表。');
Insert Into 影像参数说明(ID,PID,系统,模块,分组,参数序号,参数名,默认值,参数级别,取值范围,启用条件,说明)
Values(sys_guid(),'',100,'1291',null,1,'是否显示报告列表','1',2,null,null,'根据此参数配置嵌入式报告中是否显示报告列表。');
Insert Into 影像参数说明(ID,PID,系统,模块,分组,参数序号,参数名,默认值,参数级别,取值范围,启用条件,说明)
Values(sys_guid(),'',100,'1294',null,1,'是否显示报告列表','1',2,null,null,'根据此参数配置嵌入式报告中是否显示报告列表。');

--89602:陈振原,2015-10-27,医技工作站过滤病人方式 按执行时间或发送时间
Insert Into Zlparameters(ID,系统,模块,私有,本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1263, 0, 0, 0, 0, 23, '病人过滤方式','0', '0', '医技工作站过滤病人时的过滤方式按医嘱：0-执行时间，1-发送时间'
  From Dual Where Not Exists (Select 1 From zlParameters Where 参数号 = 23 And Nvl(模块, 0) = 1263 And Nvl(系统,0) = &n_System);

--86905:张险华,2015-10-12,增加对新版电子病历的查看分类控制
Update Zlparameters set 私有=0 Where 系统 = &n_System And 模块=2250 And 参数号 =1;
Insert Into Zlparameters(ID,系统,模块,私有,本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 2250, 0, 0, 1, 1, 2, '查看文件种类', '01|02,04,05|03,04,05', '01|02,04,05|03,04,05', '以竖线分隔表示门诊|住院|护站三种场合需要查看的文件种类'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块=2250 And 参数号 = 2);

--90346:梁经伙,2015-11-09,住院医生工作站，增加参数，控制对拥有全院病人权限的操作者在住院医生工作站科室和病区的显示
Insert Into Zlparameters
  (Id, 系统, 模块, 私有, 本机, 授权, 固定,  参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1261, 0, 0, 0, 0, 41, '不显示无床位的病区科室', '0', '0',
         '该参数控制对拥有全院病人权限的操作者在住院医生工作站科室和病区的显示，启用该参数后，当住院病人列表按科室显示时，不显示科室对应的病区没有床位的科室；当住院病人列表按病区显示时，不显示没有床位的病区。:0-显示无床位的病区或科室，1-不显示无床位的病区或科室'
  From Dual
  Where Not Exists (Select 1
         From Zlparameters
         Where 系统 = &n_System And 模块 = 1261 And 参数名 = '不显示无床位的病区科室');
---90823:刘硕,2015-11-23,结构化地址增加虚拟地址识别		 
--89238:刘硕,2015-10-23,结构化地址支持第四级字典输入数据升级
Update 区域 Set 编码 = Rpad(编码, 15, '0'), 上级编码 = Rpad(上级编码, 15, '0');
update 区域 set 是否虚拟=1,是否不显示=1  where 名称 like '%直辖县%' and 级数=1 ;
update 区域 set 是否虚拟=1,是否不显示=1  where 名称 like '县(%' and 级数=1 ;
update 区域 set 是否虚拟=1,是否不显示=0  where 名称 like '市辖区(%' and 级数=1 ;

--91265:张德婷,2015-12-10
update 诊疗检查部位 set 方法=substr(方法,1,instr(方法,chr(9),1)) || replace(substr(方法,instr(方法,chr(9),1)+1),';',chr(9))   where  instr(方法,chr(9),1)>0;
-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--89736:张德婷,2015-11-24,输液配置中心收费标准
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1345, '基本', User, '配置收费方案', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1345 And 功能 = '基本' And Upper(对象) = Upper('配置收费方案'));
         
--89736:张德婷,2015-11-24,输液配置中心收费标准                  
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1345, '基本', User, 'Zl_配置收费方案_设置', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1345 And 功能 = '基本' And Upper(对象) = Upper('Zl_配置收费方案_设置'));
---90823:刘硕,2015-11-23,结构化地址增加虚拟地址识别
Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select a.系统, a.序号, a.功能, a.所有者, 'Zl_Adderss_Structure', 'EXECUTE'
  From Zlprogprivs a
  Where Upper(a.对象) = Upper('Zl_病人地址信息_Update') And Not Exists
   (Select 1
         From Zlprogprivs b
         Where b.系统 = a.系统 And b.序号 = a.序号 And b.功能 = a.功能 And Upper(b.对象) = Upper('Zl_Adderss_Structure'));

--89723:许华峰,2015-11-16,对于未完成的报告，医生站人员需要有此权限才能观片
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select 100,9004,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '审核前观片',2,'对于未完成的报告，医生站人员需要有此权限才能观片',0 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

--90466:刘尔旋,2015-11-12,支付宝结帐修改
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1137, '基本', User, '三方退款信息', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1137 And 功能 = '基本' And Upper(对象) = Upper('三方退款信息'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1137, '基本', User, 'Zl_三方退款信息_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1137 And 功能 = '基本' And Upper(对象) = Upper('Zl_三方退款信息_Insert'));

--90488:刘尔旋,2015-11-11,存在计划时修改挂号安排
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1110, '安排', User, 'Zl_挂号安排_Modify', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1110 And 功能 = '安排' And Upper(对象) = Upper('Zl_挂号安排_Modify'));


--89297:梁唐彬,2015-10-20,医嘱发送自定义收费
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1252, '基本', User, 'zl_fun_CustomExpenses', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1252 And 功能 = '基本' And Upper(对象) = Upper('zl_fun_CustomExpenses'));
         
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1253, '基本', User, 'zl_fun_CustomExpenses', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1253 And 功能 = '基本' And Upper(对象) = Upper('zl_fun_CustomExpenses'));
         
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1254, '基本', User, 'zl_fun_CustomExpenses', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1254 And 功能 = '基本' And Upper(对象) = Upper('zl_fun_CustomExpenses'));

--82859:李南春,2015-10-15,挂号基本信息调整
Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select &n_System,1111,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0 Union All
Select NULL,&n_System,9003,0,'基本信息调整',1 From Dual Union All
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0) A;

Insert Into Zlrolegrant
  (系统, 序号, 角色, 功能)
  Select &n_System, 9003, 角色, 功能
  From Zlrolegrant a
  Where 系统 = &n_System And 序号 = 1111 And 功能 = '基本信息调整' And Not Exists
   (Select 1 From Zlrolegrant Where 系统 = &n_System And 序号 = 9003 And 功能 = '基本信息调整' And 角色 = a.角色);
   
Delete zlProgFuncs where 系统=&n_System and 序号=1111 and 功能='基本信息调整';

--89196:马政,2015-10-10,药品结存权限添加
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1332, '基本', User, 'Zl_药品结存记录_Cancel', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1332 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品结存记录_Cancel'));         

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1332, '基本', User, 'Zl_药品结存记录_Delete', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1332 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品结存记录_Delete'));


--80880:余伟节,2015-11-24,电子病案打印
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1566,'电子病案打印','病人的电子病案的打印。',&n_System,'zl9CISAudit');

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1566,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '基本',-NULL,NULL,1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1566,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select 'Zl_病案打印记录_Insert','EXECUTE' From Dual Union All
Select 'zl_电子病历打印_insert','EXECUTE' From Dual Union All
Select '病案打印记录','SELECT' From Dual Union All
Select '电子病历记录','SELECT' From Dual Union All
Select '病历文件列表','SELECT' From Dual Union All
Select '病历页面格式','SELECT' From Dual Union All
Select '病历应用科室','SELECT' From Dual Union All
Select '病人护理数据','SELECT' From Dual Union All
Select '病人护理文件','SELECT' From Dual Union All
Select '病人临床路径','SELECT' From Dual Union All
Select '变异常见原因','SELECT' From Dual Union All
Select '临床路径阶段','SELECT' From Dual Union All
Select '临床路径目录','SELECT' From Dual Union All
Select '临床路径版本','SELECT' From Dual Union All
Select '临床路径分支','SELECT' From Dual Union All
Select '病人合并路径','SELECT' From Dual Union All
Select '临床路径分类','SELECT' From Dual Union All
Select '临床路径项目','SELECT' From Dual Union All
Select '病人路径评估','SELECT' From Dual Union All
Select '病人路径变异','SELECT' From Dual Union All
Select '病人诊断记录','SELECT' From Dual Union All
Select '病人手麻记录','SELECT' From Dual Union All
Select '疾病编码目录','SELECT' From Dual Union All
Select '诊疗项目目录','SELECT' From Dual Union All
Select '病人信息','SELECT' From Dual Union All
Select '病案主页','SELECT' From Dual Union All
Select '部门表','SELECT' From Dual Union All
Select '部门性质说明','SELECT' From Dual Union All
Select '病人照片','SELECT' From Dual Union All
Select '病人医嘱报告','SELECT' From Dual Union All
Select '病人医嘱记录','SELECT' From Dual Union All
Select '病人护理记录','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlMenus(组别,ID,上级ID,标题,快键,说明,系统,模块,短标题,图标) Select A.组别,ZlMenus_ID.Nextval,A.ID,B.* From (
Select 组别,ID From zlMenus Where 标题 = '病案质控与评分系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null) A,
(Select 标题,快键,说明,系统,模块,短标题,图标 From zlMenus Where 1 = 0 Union All
Select '电子病案打印' ,'P' ,'对病人的病案进行预览、打印输出。' ,&n_System,1566,'电子病案打印' ,175 From Dual Union All
Select 标题,快键,说明,系统,模块,短标题,图标 From zlMenus Where 1 = 0) B;


-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------
--89305:张德婷,2015-11-25,扫两次瓶签自动发药
--报表：ZL1_INSIDE_1345_4/输液药品发药汇总单
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1345_4','输液药品发药汇总单','打印输液药品发药单',']}!h p`gc(.[xhv;S,EW',Null,15,0,0,100,Null,Null,Sysdate,Sysdate,To_Date('2015-11-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2015-11-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(zlReports_ID.CurrVal,100,1345,'输液药品发药汇总单');
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'输液瓶/袋汇总清单',9788,9626,256,1,0,0);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,2,'输液药品发药单1',11888,9626,256,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'输液数量统计','病区,202|科室,202|床号,202|姓名,202|住院号,131|打包,139|配液,139',User||'.输液配药记录,'||User||'.部门表',1,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select 病区, 科室, 床号, 姓名, 住院号, Sum(打包) As 打包, Sum(配液) As 配液' From Dual Union All
  Select 2,'From (Select b.名称 As 病区, c.名称 As 科室, a.床号, a.姓名 || ''('' || a.性别 || '' '' || a.年龄 || '')'' As 姓名, a.住院号, 0 As 打包, 1 As 配液' From Dual Union All
  Select 3,'       From 输液配药记录 A, 部门表 B, 部门表 C' From Dual Union All
  Select 4,'       Where a.病人病区id = b.Id And a.病人科室id = c.Id And A.操作状态=5 And a.部门id = [0] And' From Dual Union All
  Select 5,'             A.配药批次 = [1] And Nvl(是否打包, 0) = 0 and A.操作时间 between  Trunc(sysdate) and Trunc(sysdate + 1) - 1 / 24 / 60 / 60' From Dual Union All
  Select 6,'       Union All' From Dual Union All
  Select 7,'       Select b.名称 As 病区, c.名称 As 科室, a.床号, a.姓名 || ''('' || a.性别 || '' '' || a.年龄 || '')'' As 姓名, a.住院号, 1 As 打包, 0 As 配液' From Dual Union All
  Select 8,'       From 输液配药记录 A, 部门表 B, 部门表 C' From Dual Union All
  Select 9,'       Where a.病人病区id = b.Id And a.病人科室id = c.Id And A.操作状态=5 And a.部门id = [0] And Nvl(是否打包, 0) = 1 And' From Dual Union All
  Select 10,'             A.配药批次 = [1] and A.操作时间 between  Trunc(sysdate) and Trunc(sysdate + 1) - 1 / 24 / 60 / 60)' From Dual Union All
  Select 11,'Group By 病区, 科室, 床号, 姓名, 住院号' From Dual Union All
  Select 12,'Order By 病区, 科室, 床号, 姓名, 住院号' From Dual Union All
  Select 13,Null From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'部门',1,'选择器定义…',0,Null,Null,'Select Distinct B.ID, B.编码, B.名称 From 部门性质说明 A, 部门表 B Where A.部门id = B.ID And A.工作性质 = ''配制中心''',Null,'ID,131,'||CHR(38)||'B|编码,202,'||CHR(38)||'S|名称,202,'||CHR(38)||'S'||CHR(38)||'D',User||'.部门性质说明,'||User||'.部门表|',0);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,1,'配药批次',1,'选择器定义…',0,Null,Null,'Select Distinct 批次 from 配药工作批次 order by 批次',Null,'批次,131,'||CHR(38)||'S'||CHR(38)||'D'||CHR(38)||'B',User||'.配药工作批次|',0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'药品收发记录_数据','部门名称,202|病区名称,202|药品名称,202|规格,202|产地,202|批号,202|数量,139|单位,202|发送人,202|发送时间,202',User||'.药品收发记录,'||User||'.收费项目目录,'||User||'.药品规格,'||User||'.输液配药内容,'||User||'.部门表,'||User||'.输液配药记录',1,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select E.名称 As 部门名称, G.名称 As 病区名称, ''('' || B.编码 || '')'' || B.名称 As 药品名称, B.规格, A.产地, A.批号,' From Dual Union All
  Select 2,'       Sum(A.单量 / C.剂量系数 / C.住院包装) As 数量, C.住院单位 As 单位, F.操作人员 As 发送人, To_Char(F.操作时间, ''yyyy-mm-dd hh24:mi:ss'') As 发送时间' From Dual Union All
  Select 3,'From 药品收发记录 A, 收费项目目录 B, 药品规格 C, 输液配药内容 D, 部门表 E, 输液配药记录 F, 部门表 G' From Dual Union All
  Select 4,'Where A.药品id = B.ID And A.药品id = C.药品id And A.ID = D.收发id And A.库房id = E.ID And D.记录id = F.ID And F.病人病区id = G.ID And' From Dual Union All
  Select 5,'      F.部门id = [0] And F.配药批次 = [1] And F.操作状态=5 and F.操作时间 between Trunc(sysdate) and Trunc(sysdate + 1) - 1 / 24 / 60 / 60' From Dual Union All
  Select 6,'Group By E.名称, G.名称, B.编码, B.名称, B.规格, A.产地, A.批号, C.住院单位, F.操作人员, F.操作时间' From Dual Union All
  Select 7,'Order By E.名称, G.名称, B.编码, B.名称, B.规格, A.产地, A.批号, C.住院单位, F.操作人员, F.操作时间' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'部门',1,'选择器定义…',0,Null,Null,'Select Distinct B.ID, B.编码, B.名称 From 部门性质说明 A, 部门表 B Where A.部门id = B.ID And A.工作性质 = ''配制中心''',Null,'ID,131,'||CHR(38)||'B|编码,202,'||CHR(38)||'S|名称,202,'||CHR(38)||'S'||CHR(38)||'D',User||'.部门性质说明,'||User||'.部门表|',0);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,1,'配药批次',1,'选择器定义…',0,Null,Null,'Select Distinct 批次 from 配药工作批次 order by 批次',Null,'批次,131,'||CHR(38)||'S'||CHR(38)||'D'||CHR(38)||'B',User||'.配药工作批次|',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'[单位名称]输液药品发送清单',Null,30,135,9645,315,0,1,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,1,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,Null,0,'[输液数量统计.病区]',Null,270,795,2670,255,0,0,1,'宋体',12,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,Null,0,'发送人:[药品收发记录_数据.发送人]',Null,285,8280,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,Null,0,Null,0,'发送时间:[药品收发记录_数据.发送时间]',Null,3345,8280,3330,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,Null,0,'打印时间：[yyyy-mm-dd hh:mm:ss]',Null,6810,8280,2790,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表3',4,Null,0,Null,0,Null,Null,305,1110,9330,7070,225,0,0,'宋体',10.5,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[输液数量统计.病区]','4^255^病区',0,0,0,0,0,0,0,'宋体',0,0,0,0,0,0,1,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[输液数量统计.科室]','4^255^科室',0,0,2010,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[输液数量统计.床号]','4^255^床号',0,0,1170,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[输液数量统计.姓名]','4^255^姓名',0,0,1680,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[输液数量统计.住院号]','4^255^住院号',0,0,1140,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[输液数量统计.打包]','4^255^打包',0,0,1020,0,0,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[输液数量统计.配液]','4^255^配液',0,0,975,0,0,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签3',2,Null,0,Null,0,'发送人:[药品收发记录_数据.发送人]',Null,255,6330,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签2',2,Null,0,Null,0,'部门名称:[药品收发记录_数据.部门名称]',Null,270,900,3885,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签1',2,Null,0,Null,0,'[单位名称]输液药品发送清单',Null,3030,285,4320,330,0,0,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,1,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签4',2,Null,0,Null,0,'发送时间:[药品收发记录_数据.发送时间]',Null,3885,6315,3330,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签5',2,Null,0,Null,0,'打印时间：[yyyy-mm-dd hh:mm:ss]',Null,8385,6330,2790,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'任意表1',4,Null,0,Null,0,Null,Null,260,1260,10925,5000,255,0,0,'宋体',10.5,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[药品收发记录_数据.病区名称]','4^255^病区',0,0,1260,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[药品收发记录_数据.药品名称]','4^255^药品名称',0,0,2490,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[药品收发记录_数据.规格]','4^255^规格',0,0,1695,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[药品收发记录_数据.产地]','4^255^产地',0,0,1605,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[药品收发记录_数据.批号]','4^255^批号',0,0,1695,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[药品收发记录_数据.数量]','4^255^数量',0,0,1095,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[药品收发记录_数据.单位]','4^255^单位',0,0,1275,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);

--报表：ZL1_INSIDE_1345_4/输液药品发药汇总单
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1345,'输液药品发药汇总单','打印输液药品发药单');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select 100,1345,'输液药品发药汇总单',User,'部门表','SELECT' From Dual Union All
  Select 100,1345,'输液药品发药汇总单',User,'部门性质说明','SELECT' From Dual Union All
  Select 100,1345,'输液药品发药汇总单',User,'配药工作批次','SELECT' From Dual Union All
  Select 100,1345,'输液药品发药汇总单',User,'收费项目目录','SELECT' From Dual Union All
  Select 100,1345,'输液药品发药汇总单',User,'输液配药记录','SELECT' From Dual Union All
  Select 100,1345,'输液药品发药汇总单',User,'输液配药内容','SELECT' From Dual Union All
  Select 100,1345,'输液药品发药汇总单',User,'药品规格','SELECT' From Dual Union All
  Select 100,1345,'输液药品发药汇总单',User,'药品收发记录','SELECT' From Dual;


--90065:刘尔旋,2015-11-02,新增报表发布模块
Insert Into zlRPTPuts
  (报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1261, '预约挂号单'
  From zlReports
  Where 编号 = 'ZL1_BILL_1111_1' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1261 And 系统 = &n_System And 功能 = '预约挂号单');

--80880:余伟节,2015-11-24,电子病案打印
Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '打印首页'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1261_1' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '打印首页');

Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '中医病案首页'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1261_4' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '中医病案首页');
   
Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '四川省西医首页'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1261_5' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '四川省西医首页');

Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '四川省中医首页'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1261_6' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '四川省中医首页');
    
Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '云南省西医首页'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1261_7' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '云南省西医首页');

Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '云南省中医首页'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1261_8' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '云南省中医首页');

Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '湖南省病案首页'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1261_9' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '湖南省病案首页');
   
Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '湖南省中医病案首页'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1261_10' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '湖南省中医病案首页');

Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '长期医嘱单'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1254_1' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '长期医嘱单');

Insert Into zlRPTPuts(报表id, 系统, 程序id, 功能)
  Select ID, &n_System, 1566, '临时医嘱单'
  From zlReports
  Where 编号 = 'ZL1_INSIDE_1254_2' And 系统 = &n_System And Not Exists
   (Select 1 From zlRPTPuts Where 程序id = 1566 And 系统 = &n_System And 功能 = '临时医嘱单');

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'打印首页',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select 100,1566,'打印首页',User,'病案主页','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'病案主页从表','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'病人过敏记录','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'病人临床路径','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'病人手麻记录','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'病人信息','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'病人诊断记录','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'部门表','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'出院方式','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'婚姻状况','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'疾病编码目录','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'入院方式','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'性别','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'血型','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'医疗付款方式','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'诊断符合情况','SELECT' From Dual Union All
  Select 100,1566,'打印首页',User,'诊疗项目目录','SELECT' From Dual;

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'中医病案首页','中医病案首页');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1566,'中医病案首页',User,'病案主页','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'病案主页从表','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'病人过敏记录','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'病人临床路径','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'病人手麻记录','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'病人诊断记录','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'出院方式','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'婚姻状况','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'疾病编码目录','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'入院方式','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'性别','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'血型','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'医疗付款方式','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'诊断符合情况','SELECT' From Dual Union All
  Select &n_System,1566,'中医病案首页',User,'诊疗项目目录','SELECT' From Dual;

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'四川省西医首页','四川省西医首页');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1566,'四川省西医首页',User,'病案重症监护情况','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'病案主页','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'病案主页从表','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'病人感染记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'病人过敏记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'病人抗生素记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'病人临床路径','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'病人手麻记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'病人诊断记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'不良事件','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'出院方式','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'感染部位','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'婚姻状况','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'疾病编码目录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'器械导管目录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'器械导管使用情况','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'入院方式','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'性别','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'血型','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'医疗付款方式','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'医院感染目录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'诊断符合情况','SELECT' From Dual Union All
  Select &n_System,1566,'四川省西医首页',User,'诊疗项目目录','SELECT' From Dual;

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'四川省中医首页','四川省中医首页');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1566,'四川省中医首页',User,'病案重症监护情况','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'病案主页','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'病案主页从表','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'病人感染记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'病人过敏记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'病人抗生素记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'病人临床路径','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'病人手麻记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'病人诊断记录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'出院方式','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'婚姻状况','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'疾病编码目录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'入院方式','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'性别','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'血型','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'医疗付款方式','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'医院感染目录','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'诊断符合情况','SELECT' From Dual Union All
  Select &n_System,1566,'四川省中医首页',User,'诊疗项目目录','SELECT' From Dual;

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'云南省西医首页','云南省西医首页');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1566,'云南省西医首页',User,'Zl_电子病历打印_Insert','EXECUTE' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'病案重症监护情况','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'病案主页','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'病案主页从表','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'病人过敏记录','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'病人抗生素记录','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'病人临床路径','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'病人手麻记录','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'病人诊断记录','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'不良事件','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'出院方式','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'婚姻状况','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'疾病编码目录','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'入院方式','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'性别','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'血型','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'医疗付款方式','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'诊断符合情况','SELECT' From Dual Union All
  Select &n_System,1566,'云南省西医首页',User,'诊疗项目目录','SELECT' From Dual;

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'云南省中医首页','云南省中医首页');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1566,'云南省中医首页',User,'Zl_电子病历打印_Insert','EXECUTE' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'病案主页','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'病案主页从表','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'病人过敏记录','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'病人临床路径','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'病人手麻记录','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'病人诊断记录','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'出院方式','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'婚姻状况','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'疾病编码目录','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'入院方式','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'性别','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'血型','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'医疗付款方式','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'诊断符合情况','SELECT' From Dual Union All
  Select &n_System,1566,'云南省中医首页',User,'诊疗项目目录','SELECT' From Dual;

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'湖南省病案首页','湖南省病案首页');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1566,'湖南省病案首页',User,'Zl_电子病历打印_Insert','EXECUTE' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'病案主页','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'病案主页从表','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'病人过敏记录','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'病人临床路径','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'病人手麻记录','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'病人诊断记录','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'出院方式','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'婚姻状况','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'疾病编码目录','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'入院方式','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'性别','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'血型','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'医疗付款方式','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'诊断符合情况','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省病案首页',User,'诊疗项目目录','SELECT' From Dual;

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'湖南省中医病案首页','湖南省中医病案首页');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1566,'湖南省中医病案首页',User,'Zl_电子病历打印_Insert','EXECUTE' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'病案主页','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'病案主页从表','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'病人过敏记录','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'病人临床路径','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'病人手麻记录','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'病人诊断记录','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'出院方式','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'婚姻状况','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'疾病编码目录','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'入院方式','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'性别','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'血型','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'医疗付款方式','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'诊断符合情况','SELECT' From Dual Union All
  Select &n_System,1566,'湖南省中医病案首页',User,'诊疗项目目录','SELECT' From Dual;

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'长期医嘱单',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1566,'长期医嘱单',User,'病案主页','SELECT' From Dual Union All
  Select &n_System,1566,'长期医嘱单',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1566,'长期医嘱单',User,'病人医嘱打印','SELECT' From Dual Union All
  Select &n_System,1566,'长期医嘱单',User,'病人医嘱记录','SELECT' From Dual Union All
  Select &n_System,1566,'长期医嘱单',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1566,'长期医嘱单',User,'诊疗项目目录','SELECT' From Dual;

Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1566,'临时医嘱单',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1566,'临时医嘱单',User,'病案主页','SELECT' From Dual Union All
  Select &n_System,1566,'临时医嘱单',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1566,'临时医嘱单',User,'病人医嘱打印','SELECT' From Dual Union All
  Select &n_System,1566,'临时医嘱单',User,'病人医嘱记录','SELECT' From Dual Union All
  Select &n_System,1566,'临时医嘱单',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1566,'临时医嘱单',User,'收费项目目录','SELECT' From Dual Union All
  Select &n_System,1566,'临时医嘱单',User,'诊疗项目目录','SELECT' From Dual;


-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--91109:李小东,2015-11-27,人员所属部门过多错误修正
Create Or Replace Procedure Zl_人员表_修改
(
  Id_In           In 人员表.Id%Type,
  编号_In         In 人员表.编号%Type,
  姓名_In         In 人员表.姓名%Type,
  简码_In         In 人员表.简码%Type,
  身份证号_In     In 人员表.身份证号%Type,
  出生日期_In     In 人员表.出生日期%Type,
  性别_In         In 人员表.性别%Type,
  民族_In         In 人员表.民族%Type,
  工作日期_In     In 人员表.工作日期%Type,
  办公室电话_In   In 人员表.办公室电话%Type,
  电子邮件_In     In 人员表.电子邮件%Type,
  执业类别_In     In 人员表.执业类别%Type,
  执业范围_In     In 人员表.执业范围%Type,
  管理职务_In     In 人员表.管理职务%Type,
  专业技术职务_In In 人员表.专业技术职务%Type,
  聘任技术职务_In In 人员表.聘任技术职务%Type,
  学历_In         In 人员表.学历%Type,
  所学专业_In     In 人员表.所学专业%Type,
  留学时间_In     In 人员表.留学时间%Type,
  留学渠道_In     In 人员表.留学渠道%Type,
  接受培训_In     In 人员表.接受培训%Type,
  科研课题_In     In 人员表.科研课题%Type,
  个人简介_In     In 人员表.个人简介%Type,
  部门列表_In     In Varchar2, --部门列表_IN参数的填写方式如下："12:1;23:0;"
  人员性质_In     In Varchar2, --人员性质_IN参数的填写方式如下："门诊挂号员;医生;护士;"
  别名_In         In 人员表.别名%Type := Null,
  站点_In         In 人员表.站点%Type := Null,
  签名_In         In 人员表.签名%Type := Null,
  执业证号_In     In 人员表.执业证号%Type := Null,
  资格证书号_In   In 人员表.资格证书号%Type := Null,
  执业开始日期_In In 人员表.执业开始日期%Type := Null,
  处方权标志_In   In 人员表.处方权标志%Type := Null,
  手术等级_In     In 人员表.手术等级%Type := Null,
  移动电话_In     In 人员表.移动电话%Type := Null
) Is
  Intpos    Pls_Integer;
  Int缺省   Number(1);
  Strtemp   Varchar2(2000);
  Str性质   Varchar2(10);
  Lng部门id 部门表.Id%Type;
Begin
  --首先插入修改记录
  Update 人员表
  Set 编号 = 编号_In, 姓名 = 姓名_In, 简码 = 简码_In, 身份证号 = 身份证号_In, 出生日期 = 出生日期_In, 性别 = 性别_In, 民族 = 民族_In, 工作日期 = 工作日期_In,
      办公室电话 = 办公室电话_In, 电子邮件 = 电子邮件_In, 执业类别 = 执业类别_In, 执业范围 = 执业范围_In, 管理职务 = 管理职务_In, 专业技术职务 = 专业技术职务_In,
      聘任技术职务 = 聘任技术职务_In, 学历 = 学历_In, 所学专业 = 所学专业_In, 留学时间 = 留学时间_In, 留学渠道 = 留学渠道_In, 接受培训 = 接受培训_In, 科研课题 = 科研课题_In,
      个人简介 = 个人简介_In, 站点 = 站点_In, 别名 = 别名_In, 签名 = 签名_In, 执业证号 = 执业证号_In, 资格证书号 = 资格证书号_In, 执业开始日期 = 执业开始日期_In,
      处方权标志 = 处方权标志_In, 手术等级 = 手术等级_In, 移动电话 = 移动电话_In
  Where ID = Id_In;

  --接着删除已有的所属部门
  Delete From 部门人员 Where 人员id = Id_In;

  --接着修改所属部门
  Strtemp := 部门列表_In;

  While Strtemp Is Not Null Loop
    Intpos := Instr(Strtemp, ':');
  
    If Intpos = 0 Then
      Strtemp := '';
    Else
      --得到部门ID
      Str性质   := Substr(Strtemp, 1, Intpos - 1);
      Lng部门id := To_Number(Str性质);
      Strtemp   := Substr(Strtemp, Intpos + 1);
      --得到是否缺省
      Intpos  := Instr(Strtemp, ';');
      Int缺省 := To_Number(Substr(Strtemp, 1, Intpos - 1));
      Strtemp := Substr(Strtemp, Intpos + 1);
    
      Insert Into 部门人员 (部门id, 人员id, 缺省) Values (Lng部门id, Id_In, Int缺省);
    End If;
  End Loop;

  --接着删除已有的性质说明
  Delete From 人员性质说明 Where 人员id = Id_In;

  --最后修改人员性质说明
  Strtemp := 人员性质_In;

  While Strtemp Is Not Null Loop
    Intpos := Instr(Strtemp, ';');
  
    If Intpos = 0 Then
      Strtemp := '';
    Else
      --得到人员性质
      Str性质 := Substr(Strtemp, 1, Intpos - 1);
      Strtemp := Substr(Strtemp, Intpos + 1);
    
      Insert Into 人员性质说明 (人员性质, 人员id) Values (Str性质, Id_In);
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_人员表_修改;
/

--90806:胡俊勇,2015-11-25,回退出院医嘱
Create Or Replace Procedure Zl_病案主页_首页整理
(
  病人id_In       病案主页.病人id%Type,
  主页id_In       病案主页.主页id%Type,
  年龄_In         病案主页.年龄%Type,
  国籍_In         病案主页.国籍%Type,
  区域_In         病案主页.区域%Type,
  职业_In         病案主页.职业%Type,
  身高_In         病案主页.身高%Type,
  体重_In         病案主页.体重%Type,
  血型_In         病案主页.血型%Type,
  婚姻状况_In     病案主页.婚姻状况%Type,
  医疗付款方式_In 病案主页.医疗付款方式%Type,
  家庭地址_In     病案主页.家庭地址%Type,
  家庭电话_In     病案主页.家庭电话%Type,
  家庭邮编_In     病案主页.家庭地址邮编%Type,
  户口地址_In     病案主页.户口地址%Type,
  户口邮编_In     病案主页.户口地址邮编%Type,
  单位地址_In     病案主页.单位地址%Type,
  单位电话_In     病案主页.单位电话%Type,
  单位邮编_In     病案主页.单位邮编%Type,
  联系人姓名_In   病案主页.联系人姓名%Type,
  联系人关系_In   病案主页.联系人关系%Type,
  联系人电话_In   病案主页.联系人电话%Type,
  联系人地址_In   病案主页.联系人地址%Type,
  入院病况_In     病案主页.入院病况%Type,
  入院方式_In     病案主页.入院方式%Type,
  出院方式_In     病案主页.出院方式%Type,
  再入院_In       病案主页.再入院%Type,
  是否确诊_In     病案主页.是否确诊%Type,
  确诊日期_In     病案主页.确诊日期%Type,
  尸检标志_In     病案主页.尸检标志%Type,
  随诊标志_In     病案主页.随诊标志%Type,
  随诊期限_In     病案主页.随诊期限%Type,
  新发肿瘤_In     病案主页.新发肿瘤%Type,
  中医治疗类别_In 病案主页.中医治疗类别%Type,
  抢救次数_In     病案主页.抢救次数%Type,
  成功次数_In     病案主页.成功次数%Type,
  门诊医师_In     病案主页.门诊医师%Type,
  住院医师_In     病案主页.住院医师%Type,
  主治医师_In     病案主页.住院医师%Type,
  主任医师_In     病案主页.住院医师%Type,
  责任护士_In     病案主页.责任护士%Type,
  操作员编号_In   病案主页.编目员编号%Type := Null,
  操作员姓名_In   病案主页.编目员姓名%Type := Null
) As
  --功能：用于住院医护工作站对病人进行首页整理
  v_住院医师 病案主页.住院医师%Type;
  v_主治医师 病案主页.住院医师%Type;
  v_主任医师 病案主页.住院医师%Type;
  v_责任护士 病案主页.责任护士%Type;
  v_病人性质 病案主页.病人性质%Type;
  v_原因     病人变动记录.开始原因%Type;
  v_出院科室 病案主页.出院科室id%Type;
  v_Curdate  Date;
  v_Count    Number;
  v_Change   Varchar2(500);
  Cursor c_Oldinfo Is
    Select b.*
    From (Select c.*
           From 病人变动记录 C
           Where c.病人id = 病人id_In And c.主页id = 主页id_In And
                 c.开始时间 = (Select Min(开始时间)
                           From 病人变动记录
                           Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > v_Curdate)) A, 病人变动记录 B
    Where b.病人id = 病人id_In And b.主页id = 主页id_In And a.开始时间 = b.终止时间 And a.开始原因 = b.终止原因 And a.附加床位 = b.附加床位
    Union
    Select *
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And 开始时间 <= v_Curdate;

  Cursor c_Endinfo Is
    Select * From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
  r_Oldinfo  c_Oldinfo%RowType;
  r_Endinfo  c_Endinfo%RowType;
  v_终止原因 病人变动记录.终止原因%Type;
  v_终止时间 病人变动记录.终止时间%Type;
  v_终止人员 病人变动记录.终止人员%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;

Begin
  --取更改前的内容(用NoneData和新的比较)
  Select 病人性质, Nvl(住院医师, 'NoneData'), Nvl(出院科室id, 入院科室id), Nvl(责任护士, 'NoneData')
  Into v_病人性质, v_住院医师, v_出院科室, v_责任护士
  From 病案主页
  Where 病人id = 病人id_In And 主页id = 主页id_In;

  Begin
    Select Nvl(信息值, 'NoneData')
    Into v_主治医师
    From 病案主页从表
    Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主治医师';
  
  Exception
    When Others Then
      v_主治医师 := 'NoneData';
    
  End;

  Begin
    Select Nvl(信息值, 'NoneData')
    Into v_主任医师
    From 病案主页从表
    Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主任医师';
  
  Exception
    When Others Then
      v_主任医师 := 'NoneData';
    
  End;

  Update 病案主页
  Set 婚姻状况 = 婚姻状况_In, 年龄 = 年龄_In, 职业 = 职业_In, 国籍 = 国籍_In, 区域 = 区域_In, 医疗付款方式 = 医疗付款方式_In, 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In,
      家庭地址邮编 = 家庭邮编_In, 单位地址 = 单位地址_In, 单位电话 = 单位电话_In, 单位邮编 = 单位邮编_In, 联系人姓名 = 联系人姓名_In, 联系人关系 = 联系人关系_In,
      联系人电话 = 联系人电话_In, 联系人地址 = 联系人地址_In, 再入院 = 再入院_In, 入院病况 = 入院病况_In, 是否确诊 = 是否确诊_In, 确诊日期 = 确诊日期_In, 抢救次数 = 抢救次数_In,
      成功次数 = 成功次数_In, 尸检标志 = 尸检标志_In, 随诊标志 = 随诊标志_In, 随诊期限 = 随诊期限_In, 血型 = 血型_In, 门诊医师 = 门诊医师_In, 住院医师 = 住院医师_In,
      新发肿瘤 = 新发肿瘤_In, 中医治疗类别 = 中医治疗类别_In, 身高 = 身高_In, 体重 = 体重_In, 出院方式 = 出院方式_In, 入院方式 = 入院方式_In, 责任护士 = 责任护士_In,
      户口地址 = 户口地址_In, 户口地址邮编 = 户口邮编_In
  Where 病人id = 病人id_In And 主页id = 主页id_In;

  If v_住院医师 <> Nvl(住院医师_In, 'NoneData') Or v_主治医师 <> Nvl(主治医师_In, 'NoneData') Or v_主任医师 <> Nvl(主任医师_In, 'NoneData') Or
     v_责任护士 <> Nvl(责任护士_In, 'NoneData') Then
  
    v_原因 := Null;
    Select Count(*) Into v_Count From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 10;
    If v_Count > 0 Then
      v_原因 := 10;
    Else
      Select Count(*)
      Into v_Count
      From 病人变动记录
      Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 Is Null And 终止时间 Is Null;
      If v_Count > 0 Then
        v_原因 := 3;
      End If;
    End If;
  
    If v_原因 > 0 Then
      Select f_List2str(Cast(Collect(a.信息) As t_Strlist))
      Into v_Change
      From (Select 姓名 信息
             From 人员表
             Where 0 = 1
             Union All
             Select '住院医师(原:' || v_住院医师 || ')' 信息
             From Dual
             Where v_住院医师 <> Nvl(住院医师_In, 'NoneData')
             Union All
             Select '主治医师(原:' || v_主治医师 || ')'
             From Dual
             Where v_主治医师 <> Nvl(主治医师_In, 'NoneData')
             Union All
             Select '主任医师(原:' || v_主任医师 || ')'
             From Dual
             Where v_主任医师 <> Nvl(主任医师_In, 'NoneData')
             Union All
             Select '责任护士(原:' || v_责任护士 || ')' From Dual Where v_责任护士 <> Nvl(责任护士_In, 'NoneData')) A;
      If v_原因 = 3 Then
        v_Error := '该病人正在转科或转病区，不能进行如下变动：' || v_Change || '！';
      Else
        v_Error := '该病人正处于预出院状态，不能进行如下变动：' || v_Change || '！';
      End If;
      Raise Err_Custom;
    End If;
    Select Sysdate Into v_Curdate From Dual;
    Open c_Oldinfo;
    Fetch c_Oldinfo
      Into r_Oldinfo;
    Open c_Endinfo;
    Fetch c_Endinfo
      Into r_Endinfo;
    If c_Endinfo%RowCount = 0 Then
      --出院病人不进行变动处理
      Close c_Endinfo;
    Else
      --如果终止时间<>NULL ，就记录下终止时间和终止原因。
      If r_Oldinfo.终止时间 Is Not Null Then
        v_终止时间 := r_Oldinfo.终止时间;
        v_终止原因 := r_Oldinfo.终止原因;
        v_终止人员 := r_Oldinfo.终止人员;
      End If;
      --如果是待入住的病人，则不产生变动，直接修改已有变动
      If r_Oldinfo.开始原因 = 1 And r_Oldinfo.终止时间 Is Null Then
        Update 病人变动记录
        Set 经治医师 = 住院医师_In, 主治医师 = 主治医师_In, 责任护士 = 责任护士_In, 主任医师 = 主任医师_In
        Where ID = r_Oldinfo.Id;
      Else
        If v_住院医师 <> Nvl(住院医师_In, 'NoneData') Then
          v_原因 := 7;
          If v_终止时间 Is Null Then
            Update 病人变动记录
            Set 终止时间 = v_Curdate, 终止原因 = v_原因, 终止人员 = 操作员姓名_In
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
          Else
            Update 病人变动记录
            Set 终止时间 = v_Curdate, 终止原因 = v_原因, 终止人员 = 操作员姓名_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = v_终止时间 And 终止原因 = v_终止原因;
            --更新将来的记录如果有停止到将来的则删除上次计算时间
            Update 病人变动记录
            Set 经治医师 = 住院医师_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > v_Curdate;
          End If;
        
          --产生病历书写时机
          Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '交班', r_Oldinfo.科室id, 住院医师_In, v_Curdate, v_Curdate);
        
          While c_Oldinfo%Found Loop
            --注意:有附加床位时有多条记录
            Insert Into 病人变动记录
              (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情,
               操作员编号, 操作员姓名, 终止时间, 终止原因, 终止人员)
            Values
              (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, v_Curdate, v_原因, r_Oldinfo.附加床位, r_Oldinfo.病区id, r_Oldinfo.科室id,
               r_Oldinfo.医疗小组id, r_Oldinfo.护理等级id, r_Oldinfo.床位等级id, r_Oldinfo.床号, r_Oldinfo.责任护士, 住院医师_In,
               r_Oldinfo.主治医师, r_Oldinfo.主任医师, r_Oldinfo.病情, 操作员编号_In, 操作员姓名_In, v_终止时间, v_终止原因, v_终止人员);
          
            Fetch c_Oldinfo
              Into r_Oldinfo;
          End Loop;
        
          --如果存在停止到将来的变动就更新终止原因
          If v_终止时间 Is Not Null Then
            v_终止原因 := v_原因;
            v_终止时间 := v_Curdate;
            v_终止人员 := 操作员姓名_In;
          End If;
        
          Close c_Oldinfo;
          Open c_Oldinfo; --重新打开,以便取最新信息
          Fetch c_Oldinfo
            Into r_Oldinfo;
        End If;
      
        If v_主治医师 <> Nvl(主治医师_In, 'NoneData') Then
          Update 病案主页从表
          Set 信息值 = 主治医师_In
          Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主治医师';
        
          If Sql%RowCount = 0 Then
            Insert Into 病案主页从表
              (病人id, 主页id, 信息名, 信息值)
            Values
              (病人id_In, 主页id_In, '主治医师', 主治医师_In);
          
          End If;
        
          v_原因 := 11;
          If v_终止时间 Is Null Then
            Update 病人变动记录
            Set 终止时间 = v_Curdate, 终止原因 = v_原因, 终止人员 = 操作员姓名_In
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
          
          Else
            Update 病人变动记录
            Set 终止时间 = v_Curdate, 终止原因 = v_原因, 终止人员 = 操作员姓名_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = v_终止时间 And 终止原因 = v_终止原因;
            --更新将来的记录如果有停止到将来的则删除上次计算时间
            Update 病人变动记录
            Set 主治医师 = 主治医师_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > v_Curdate;
          End If;
        
          While c_Oldinfo%Found Loop
            Insert Into 病人变动记录
              (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情,
               操作员编号, 操作员姓名, 终止时间, 终止原因, 终止人员)
            Values
              (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, v_Curdate, v_原因, r_Oldinfo.附加床位, r_Oldinfo.病区id, r_Oldinfo.科室id,
               r_Oldinfo.医疗小组id, r_Oldinfo.护理等级id, r_Oldinfo.床位等级id, r_Oldinfo.床号, r_Oldinfo.责任护士, r_Oldinfo.经治医师,
               主治医师_In, r_Oldinfo.主任医师, r_Oldinfo.病情, 操作员编号_In, 操作员姓名_In, v_终止时间, v_终止原因, v_终止人员);
          
            Fetch c_Oldinfo
              Into r_Oldinfo;
          End Loop;
        
          --如果存在停止到将来的变动就更新终止原因
          If v_终止时间 Is Not Null Then
            v_终止原因 := v_原因;
            v_终止时间 := v_Curdate;
            v_终止人员 := 操作员姓名_In;
          End If;
        
          Close c_Oldinfo;
          Open c_Oldinfo;
          Fetch c_Oldinfo
            Into r_Oldinfo;
        End If;
      
        If v_责任护士 <> Nvl(责任护士_In, 'NoneData') Then
          v_原因 := 8;
          If v_终止时间 Is Null Then
            Update 病人变动记录
            Set 终止时间 = v_Curdate, 终止原因 = v_原因, 终止人员 = 操作员姓名_In
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
          Else
            Update 病人变动记录
            Set 终止时间 = v_Curdate, 终止原因 = v_原因, 终止人员 = 操作员姓名_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = v_终止时间 And 终止原因 = v_终止原因;
            --更新将来的记录，如果有停止到将来的则删除上次计算时间
            Update 病人变动记录
            Set 责任护士 = 责任护士_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > v_Curdate;
          End If;
          While c_Oldinfo%Found Loop
            Insert Into 病人变动记录
              (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情,
               操作员编号, 操作员姓名, 终止时间, 终止原因, 终止人员)
            Values
              (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, v_Curdate, v_原因, r_Oldinfo.附加床位, r_Oldinfo.病区id, r_Oldinfo.科室id,
               r_Oldinfo.医疗小组id, r_Oldinfo.护理等级id, r_Oldinfo.床位等级id, r_Oldinfo.床号, 责任护士_In, r_Oldinfo.经治医师,
               r_Oldinfo.主治医师, r_Oldinfo.主任医师, r_Oldinfo.病情, 操作员编号_In, 操作员姓名_In, v_终止时间, v_终止原因, v_终止人员);
            Fetch c_Oldinfo
              Into r_Oldinfo;
          End Loop;
          --如果存在停止到将来的变动就更新终止原因
          If v_终止时间 Is Not Null Then
            v_终止原因 := v_原因;
            v_终止时间 := v_Curdate;
            v_终止人员 := 操作员姓名_In;
          End If;
          Close c_Oldinfo;
          Open c_Oldinfo;
          Fetch c_Oldinfo
            Into r_Oldinfo;
        End If;
      
        If v_主任医师 <> Nvl(主任医师_In, 'NoneData') Then
          Update 病案主页从表
          Set 信息值 = 主任医师_In
          Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主任医师';
        
          If Sql%RowCount = 0 Then
            Insert Into 病案主页从表
              (病人id, 主页id, 信息名, 信息值)
            Values
              (病人id_In, 主页id_In, '主任医师', 主任医师_In);
          
          End If;
        
          v_原因 := 12;
          If v_终止时间 Is Null Then
            Update 病人变动记录
            Set 终止时间 = v_Curdate, 终止原因 = v_原因, 终止人员 = 操作员姓名_In
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
          
          Else
            Update 病人变动记录
            Set 终止时间 = v_Curdate, 终止原因 = v_原因, 终止人员 = 操作员姓名_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = v_终止时间 And 终止原因 = v_终止原因;
            --更新将来的记录如果有停止到将来的则删除上次计算时间
            Update 病人变动记录
            Set 主任医师 = 主任医师_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > v_Curdate;
          End If;
        
          While c_Oldinfo%Found Loop
            Insert Into 病人变动记录
              (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情,
               操作员编号, 操作员姓名, 终止时间, 终止原因, 终止人员)
            Values
              (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, v_Curdate, v_原因, r_Oldinfo.附加床位, r_Oldinfo.病区id, r_Oldinfo.科室id,
               r_Oldinfo.医疗小组id, r_Oldinfo.护理等级id, r_Oldinfo.床位等级id, r_Oldinfo.床号, r_Oldinfo.责任护士, r_Oldinfo.经治医师,
               r_Oldinfo.主治医师, 主任医师_In, r_Oldinfo.病情, 操作员编号_In, 操作员姓名_In, v_终止时间, v_终止原因, v_终止人员);
          
            Fetch c_Oldinfo
              Into r_Oldinfo;
          End Loop;
        
          Close c_Oldinfo;
          Open c_Oldinfo;
          Fetch c_Oldinfo
            Into r_Oldinfo;
        End If;
      
        Close c_Oldinfo;
        Close c_Endinfo;
        Select Count(*)
        Into v_Count
        From 病人变动记录
        Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(附加床位, 0) = 0 And 开始时间 Is Not Null And 终止时间 Is Null;
      
        If v_Count > 1 Then
          v_Error := '发现病人存在非法的变动记录,当前操作不能继续！' || Chr(13) || Chr(10) || '这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
          Raise Err_Custom;
        End If;
      
      End If;
    
    End If;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
  
End Zl_病案主页_首页整理;
/

--89736:张德婷,2015-11-24,输液配置中心收费标准
CREATE OR REPLACE Procedure Zl_输液配药记录_取消配药(配药id_In In Varchar2 --ID串:ID1,ID2....
                                           ) Is
  v_Tansid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_No       Varchar2(20);
  v_Usercode Varchar2(100);
  n_打包     输液配药记录.是否打包%Type := 0;
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;

  v_Error    Varchar2(255);
  n_操作状态 输液配药记录.操作状态%Type;
  Err_Custom Exception;

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');

    --检查当前输液单的状态是否为待摆药状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;

      if n_操作状态!=4 then
        v_Error := '该数据当前不是配药状态，不能进行取消配药！';
        Raise Err_Custom;
      end if;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;

    Select 操作人员, 操作时间
    Into v_操作人员, d_操作时间
    From (Select 操作人员, 操作时间 From 输液配药状态 Where 配药id = v_Tansid And 操作类型 = 2 Order By 操作时间 Desc)
    Where Rownum = 1;

    Update 输液配药记录 Set 操作状态 = 2, 操作人员 = v_操作人员, 操作时间 = d_操作时间 Where ID = v_Tansid;

    Select 是否打包 Into n_打包 From 输液配药记录 Where ID = v_Tansid;
    If n_打包 <> 1 Then
      for r_item in (Select NO From 输液配药附费 Where 配药id = v_Tansid) loop
        if r_item.NO is not null then
          Zl_住院记帐记录_Delete(r_item.NO, 1, v_Usercode, Zl_Username);
        end if;
      end loop;
    Else
      Zl_输液配药记录_取消摆药(v_Tansid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_取消配药;
/

--89736:张德婷,2015-11-24,输液配置中心收费标准
CREATE OR REPLACE Procedure Zl_配置收费方案_设置
( 
  序号_In In 配置收费方案.序号%Type, 
  配药类型_In   In 配置收费方案.配药类型%Type, 
  项目id_In   In 配置收费方案.项目id%Type, 
  收费项目_In   In 配置收费方案.收费项目%Type, 
  n_First   In Number 
) Is 
Begin 
  If n_First = 1 Then 
    Delete From 配置收费方案; 
  End If; 
 
  Insert Into 配置收费方案 (序号, 配药类型,项目id,收费项目) Values (序号_In, 配药类型_In,项目id_In,收费项目_In); 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_配置收费方案_设置;
/

--89736:张德婷,2015-11-24,输液配置中心收费标准
CREATE OR REPLACE Procedure Zl_输液配药记录_配药
(
  配药id_In   In Varchar2, --ID串:ID1,ID2....
  操作人员_In In 输液配药记录.操作人员%Type := Null,
  操作时间_In In 输液配药记录.操作时间%Type := Null
) Is
  v_Tansid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_No       Varchar2(20);
  v_Usercode Varchar2(100);
  n_操作状态 输液配药记录.操作状态%Type;
  v_Error    Varchar2(255);
  n_People      number(1);
  n_row      number(2);
  d_执行时间 date;
  v_配药类型 varchar2(50);
  n_项目id   number(18);
  Err_Custom Exception;

  Cursor c_Bill Is
    Select a.病人id, a.主页id, a.标识号, a.姓名, a.性别, a.年龄, a.床号, a.费别, a.病人病区id, a.病人科室id, a.婴儿费, e.药品id, b.库房id,f.配药类型
    From 住院费用记录 A, 药品收发记录 B, 输液配药记录 C, 输液配药内容 D, 药品规格 E, 输液药品属性 F
    Where a.Id = b.费用id And b.Id = d.收发id And d.记录id = c.Id And b.药品id = e.药品id And b.药品id = f.药品id And Nvl(c.是否打包, 0) <> 1 And c.Id = v_Tansid;

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_People:=Nvl(zl_GetSysParameter('配置费按病人收取', 1345), 0);

  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');

    --检查当前输液单的状态是否为待摆药状态
    Begin
      Select 操作状态,执行时间 Into n_操作状态,d_执行时间 From 输液配药记录 Where ID = v_Tansid;

      if n_操作状态>3 then
        v_Error := '该数据已被操作，不能进行发药！';
        Raise Err_Custom;
      end if;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;

    Update 输液配药记录 Set 操作状态 = 4, 操作人员 = 操作人员_In, 操作时间 = 操作时间_In Where ID = v_Tansid;
    Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (v_Tansid, 4, 操作人员_In, 操作时间_In);


    For r_Bill In c_Bill Loop
      select count(项目id) into n_项目id from 配置收费方案 where 配药类型=substr(r_Bill.配药类型,INSTR(r_Bill.配药类型,'-',1,1)+1);
      if n_项目id<>0 then
        n_row:=0;
        select nvl(项目id,0) into n_项目id from 配置收费方案 where 配药类型=substr(r_Bill.配药类型,INSTR(r_Bill.配药类型,'-',1,1)+1);
        if n_People=1 then
          select count(配药id) into n_row from 输液配药附费 A,住院费用记录 B,输液配药记录 C where A.No=b.no and A.配药ID=C.id and b.病人id=r_Bill.病人id And B.记录状态=1 and B.收费细目id=n_项目id and d_执行时间 Between Trunc(c.执行时间) And Trunc(c.执行时间+1) - 1 / 24 / 60 / 60;
        end if;
      else
        n_row:=1;
      end if;

      if n_row=0 then
        For r_Item In (Select a.Id 收费细目id, a.类别 收费类别, a.计算单位, a.加班加价 加班标志, d.Id 收入项目id, d.收据费目, b.现价
                       From 收费项目目录 A, 收费价目 B, 收入项目 D
                       Where a.Id = b.收费细目id And b.收入项目id = d.Id And a.id=n_项目id and
                             b.执行日期 <= Sysdate And
                             (b.终止日期 >= Sysdate Or b.终止日期 Is Null)) Loop
          Select Nextno(14) Into v_No From Dual;
          Insert Into 输液配药附费 (配药id, NO,病人id) Values (v_Tansid, v_No, r_Bill.病人id);

          Zl_住院记帐记录_Insert(v_No, 1, r_Bill.病人id, r_Bill.主页id, r_Bill.标识号, r_Bill.姓名, r_Bill.性别, r_Bill.年龄, r_Bill.床号,
                           r_Bill.费别, r_Bill.病人病区id, r_Bill.病人科室id, r_Item.加班标志, r_Bill.婴儿费, r_Bill.库房id, 操作人员_In, Null,
                           r_Item.收费细目id, r_Item.收费类别, r_Item.计算单位, Null, Null, Null, 1, 1, Null, r_Bill.库房id, Null,
                           r_Item.收入项目id, r_Item.收据费目, r_Item.现价, r_Item.现价, r_Item.现价, Null, Sysdate, Sysdate, Null, Null,
                           v_Usercode, 操作人员_In);
        End Loop;
      end if;
      if n_People<>1 and n_row=0 then
        Exit;
      end if;
    End Loop;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_配药;
/

--90760:蔡青松,2015-11-23,将LPad的第二个参数,原为13,现改为23
Create Or Replace Procedure Zl_检验报告单_Update
(
  检验标本id_In   In 检验普通结果.检验标本id%Type,
  显示隐私项目_In In Number := 0
) Is

  V_主页id     病人挂号记录.Id%Type;
  V_医嘱id     病人医嘱记录.Id%Type;
  V_开嘱科室id 病人医嘱记录.开嘱科室id%Type;
  V_病人来源   检验标本记录.病人来源%Type;
  V_病人id     检验标本记录.病人id%Type;
  V_婴儿       检验标本记录.婴儿%Type;
  V_审核人     检验标本记录.审核人%Type;
  V_病历文件id 病历单据应用.病历文件id%Type;
  V_病历文件名 病历文件列表.名称%Type;
  V_当前父id   电子病历内容.父id%Type;
  V_父id_In    电子病历内容.父id%Type;
  V_写入格式   Number := 0; --1=检验结果
  V_文件id     电子病历内容.文件id%Type;
  V_Nextid     电子病历内容.Id%Type;
  V_Loop       Number := 0;
  V_替换内容   电子病历内容.内容文本%Type;
  V_检验结果   Varchar2(5000);
  V_微生物     检验标本记录.微生物标本%Type;
  V_结果长度   Number := 32;
  V_报告       检验标本记录.一级报告%Type;

  Cursor V_Source Is
    Select ID, 文件id, Nvl(父id, 0) As 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id 复用提纲, 使用时机, 诊治要素id, 替换域,
           要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域
    From 病历文件结构
    Where 文件id = V_病历文件id
    Order By 对象序号;

  Cursor V_Result Is
    Select RPad('检验项目', V_结果长度) As 检验项目, RPad('检验结果', 10) As 检验结果, LPad('单位', 8) As 单位, LPad('结果标志', 10) As 结果标志,
           LPad('结果参考', 23) As 结果参考, 0 As 诊疗项目id, 0 As 排列序号, 0 As 隐私项目, '0' As 编码
    From Dual
    Union All
    Select RPad(B.中文名 || '(' || B.英文名 || ')', V_结果长度) As 检验项目, RPad(Nvl(A.检验结果, ' '), 10) As 检验结果,
           LPad(Nvl(B.单位, ' '), 8) As 单位,
           LPad(Decode(A.结果标志, 3, '↑', 2, '↓', 1, ' ', 4, '异常', 5, '↓↓', 6, '↑↑', ' '), 10) As 结果标志,
           LPad(Nvl(Zlgetreference(B.Id, C.标本类型, Decode(C.性别, '男', 1, '女', 2, 0), C.出生日期, C.仪器id, C.年龄, C.申请科室id), ' '),
                 23) As 结果参考, 诊疗项目id, A.排列序号, Nvl(D.隐私项目, 0) As 隐私项目,
           LPad(Decode(D.排列序号, Null, Nvl(H.编码, B.编码), D.排列序号), 4, '0') As 编码
    From 检验普通结果 A, 诊治所见项目 B, 检验标本记录 C, 检验项目 D, 诊疗项目目录 H
    Where A.检验项目id = B.Id And A.检验标本id = C.Id And C.医嘱id = V_医嘱id And B.Id = D.诊治项目id And C.报告结果 = A.记录类型 And
          A.诊疗项目id = H.Id(+)
    Order By 编码, 排列序号;

  Cursor V_Mresult Is
    Select *
    From (Select Null As 上级id, B.Id, A.标本序号, '       鉴定结果：' As 项目, RPad(C.中文名, 30, ' ') As 耐药, B.检验结果 As 用法用量1,
                  '' As 用法用量2, '' As 血药浓度1, '' As 血药浓度2, '' As 尿药浓度1, '' As 尿药浓度2, C.编码, '' As 检验备注
           From 检验标本记录 A, 检验普通结果 B, 检验细菌 C
           Where A.Id = B.检验标本id And A.报告结果 = B.记录类型 And B.细菌id = C.Id And A.医嘱id = V_医嘱id And A.样本状态 = 2
           Union All
           Select Distinct D.细菌结果id, 0, A.标本序号, '       抗生素          ',
                           RPad(Decode(药敏方法, 1, '  MIC', 2, '  Disk', 3, '  K-B', ' '), 5, ' ') || '    耐药性 ',
                           '             用法用量                ', '用法用量2', '血药浓度     ', '血药浓度2', '  尿药浓度', '尿药浓度2',
                           '000' As 编码, ''
           From 检验标本记录 A, 检验普通结果 B, 检验药敏结果 D
           Where A.Id = B.检验标本id And A.报告结果 = B.记录类型 And B.Id = D.细菌结果id And A.医嘱id = V_医嘱id And A.样本状态 = 2
           Union All
           Select Distinct D.细菌结果id, 0, A.标本序号, '￣￣￣￣￣￣￣', '￣￣￣￣￣￣￣', '￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣', '￣￣￣￣￣', '￣￣￣￣￣￣￣￣',
                           '￣￣￣￣￣', '￣￣￣￣￣￣￣', '￣￣￣￣￣￣￣￣', '001' As 编码, ''
           From 检验标本记录 A, 检验普通结果 B, 检验药敏结果 D
           Where A.Id = B.检验标本id And A.报告结果 = B.记录类型 And B.Id = D.细菌结果id And A.医嘱id = V_医嘱id And A.样本状态 = 2
           Union All
           Select 细菌结果id, ID, 标本序号, 项目, 耐药, 用法用量1, 用法用量2, 血药浓度1, 血药浓度2, 尿药浓度1, 尿药浓度2, 编码, ''
           From (Select *
                  From (Select Distinct D.细菌结果id, 0 As ID, A.标本序号, '￣￣￣￣￣￣￣' As 项目, '￣￣￣￣￣￣￣' As 耐药,
                                         '￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣' As 用法用量1, '￣￣￣￣￣' As 用法用量2, '￣￣￣￣￣￣￣￣' As 血药浓度1,
                                         '￣￣￣￣￣￣￣￣' As 血药浓度2, '￣￣￣￣￣￣￣' As 尿药浓度1, '￣￣￣￣￣￣￣￣' As 尿药浓度2, 'AAA' As 编码, '',
                                         C.编码 As 排序
                         From 检验标本记录 A, 检验普通结果 B, 检验细菌 C, 检验药敏结果 D
                         Where A.Id = B.检验标本id And A.报告结果 = B.记录类型 And B.细菌id = C.Id And B.Id = D.细菌结果id(+) And
                               A.医嘱id = V_医嘱id And A.样本状态 = 2
                         Order By 排序) A,
                       (Select Count(Distinct B.细菌id) As Num
                         From 检验标本记录 A, 检验普通结果 B
                         Where A.Id = B.检验标本id And A.报告结果 = B.记录类型 And A.医嘱id = V_医嘱id And A.样本状态 = 2) E
                  Where Rownum < E.Num)
           Where 细菌结果id Is Not Null
           Union All
           Select D.细菌结果id, 0, A.标本序号, RPad('       ' || E.中文名, 23, ' ') As 中文名,
                  RPad('  ' || D.结果, 9, ' ') ||
                   RPad(Decode(D.结果类型, 'R', '耐药(R)', 'I', '中介(I)', 'S', '敏感(S)', ' '), 10, ' '),
                  RPad(Decode(Instr(C.中文名, '支原体'), 0, E.用法用量1, ' '), 30, ' '),
                  Decode(Instr(C.中文名, '支原体'), 0, E.用法用量2, ' '),
                  RPad(Decode(Instr(C.中文名, '支原体'), 0, '    ' || E.血药浓度1, ' '), 20),
                  Decode(Instr(C.中文名, '支原体'), 0, E.血药浓度2, ' '), RPad(Decode(Instr(C.中文名, '支原体'), 0, E.尿药浓度1, ' '), 15),
                  Decode(Instr(C.中文名, '支原体'), 0, E.尿药浓度2, ' '), E.编码, ' '
           From 检验标本记录 A, 检验普通结果 B, 检验细菌 C, 检验药敏结果 D, 检验用抗生素 E
           Where A.Id = B.检验标本id And A.报告结果 = B.记录类型 And B.细菌id = C.Id And B.Id = D.细菌结果id And D.抗生素id = E.Id And
                 A.医嘱id = V_医嘱id And A.样本状态 = 2
           Order By 标本序号, 编码)
    Connect By Prior ID = 上级id
    Start With 上级id Is Null;
Begin

  Select Nvl(B.主页id, 0), Nvl(A.医嘱id, 0), Decode(A.病人来源, 2, 2, 4, 4, 1), Nvl(A.病人id, 0), Nvl(B.开嘱科室id, 0), Nvl(A.婴儿, 0),
         A.审核人, Nvl(A.微生物标本, 0)
  Into V_主页id, V_医嘱id, V_病人来源, V_病人id, V_开嘱科室id, V_婴儿, V_审核人, V_微生物
  From 检验标本记录 A, 病人医嘱记录 B
  Where A.医嘱id = B.Id(+) And A.Id = 检验标本id_In;
  If V_病人来源 = 1 Then
    --主页ID： 门诊病人填挂号ID
    Select Nvl(B.Id, 0)
    Into V_主页id
    From 病人挂号记录 B, 检验标本记录 A
    Where A.挂号单 = B.No(+) And A.Id = 检验标本id_In;
  End If;

  Select Max(Lengthb(B.中文名 || '(' || B.英文名 || ')')) + 5
  Into V_结果长度
  From 检验普通结果 A, 诊治所见项目 B, 检验标本记录 C
  Where A.检验项目id = B.Id And A.检验标本id = C.Id And C.医嘱id = V_医嘱id;

  If Nvl(V_结果长度, 0) <= 32 Then
    V_结果长度 := 32;
  End If;

  Begin
    Select 病历文件id, C.名称
    Into V_病历文件id, V_病历文件名
    From 病人医嘱记录 A, 病历单据应用 B, 病历文件列表 C
    Where A.诊疗项目id = B.诊疗项目id And B.病历文件id = C.Id And A.相关id = V_医嘱id And B.应用场合 = V_病人来源 And Rownum <= 1;
  Exception
    When Others Then
      Return;
  End;

  --删除以前的报告记录
  Begin
    Select 病历id Into V_文件id From 病人医嘱报告 Where 医嘱id = V_医嘱id And Rownum <= 1;
    Delete 电子病历内容 Where 文件id = V_文件id;
  Exception
    When Others Then
      Select 电子病历记录_Id.Nextval Into V_文件id From Dual;

      Insert Into 电子病历记录
        (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 保存人, 保存时间, 最后版本, 签名级别)
      Values
        (V_文件id, V_病人来源, V_病人id, V_主页id, V_婴儿, V_开嘱科室id, 7, V_病历文件id, V_病历文件名, V_审核人, Sysdate, V_审核人, Sysdate, 1, 0);

      Insert Into 病人医嘱报告
        (医嘱id, 病历id)
        Select Distinct 医嘱id, V_文件id
        From (Select Distinct Decode(B.医嘱id, Null, A.医嘱id, B.医嘱id) As 医嘱id
               From 检验标本记录 A, 检验项目分布 B
               Where A.Id = B.标本id(+) And A.Id = 检验标本id_In)
        Where 医嘱id Is Not Null;
  End;

  For R_Source In V_Source Loop

    V_Loop := V_Loop + 1;
    Select 电子病历内容_Id.Nextval Into V_Nextid From Dual;
    V_写入格式 := 0;

    If R_Source.父id = 0 Then
      V_当前父id := V_Nextid;
      V_父id_In  := Null;
    Else
      V_父id_In := V_当前父id;
    End If;

    If R_Source.对象类型 = 4 And R_Source.替换域 = 1 Then
      V_替换内容 := zl_Replace_Element_Value(R_Source.要素名称, V_病人id, V_主页id, V_病人来源, V_医嘱id);
      Insert Into 电子病历内容
        (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
      Values
        (V_Nextid, V_文件id, 1, 0, V_父id_In, V_Loop, R_Source.对象类型, R_Source.对象标记, R_Source.保留对象, R_Source.对象属性,
         R_Source.内容行次, V_替换内容, R_Source.是否换行);
    Elsif R_Source.对象类型 = 1 And R_Source.内容文本 = '检验结果' Then
      Insert Into 电子病历内容
        (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
      Values
        (V_Nextid, V_文件id, 1, 0, V_父id_In, V_Loop, R_Source.对象类型, R_Source.对象标记, R_Source.保留对象, R_Source.对象属性,
         R_Source.内容行次, R_Source.内容文本, R_Source.是否换行);
      V_写入格式 := 1;
    Elsif R_Source.对象类型 = 1 And (R_Source.内容文本 = '一级报告' Or R_Source.内容文本 = '二级报告' Or R_Source.内容文本 = '三级报告') Then
      Insert Into 电子病历内容
        (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
      Values
        (V_Nextid, V_文件id, 1, 0, V_父id_In, V_Loop, R_Source.对象类型, R_Source.对象标记, R_Source.保留对象, R_Source.对象属性,
         R_Source.内容行次, R_Source.内容文本, R_Source.是否换行);
      V_写入格式 := 1;
    Else
      Insert Into 电子病历内容
        (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
      Values
        (V_Nextid, V_文件id, 1, 0, V_父id_In, V_Loop, R_Source.对象类型, R_Source.对象标记, R_Source.保留对象, R_Source.对象属性,
         R_Source.内容行次, R_Source.内容文本, R_Source.是否换行);
    End If;

    If V_写入格式 = 1 Then
      Select 电子病历内容_Id.Nextval Into V_Nextid From Dual;
      --在循环中插入检验结果格式
      If V_微生物 = 0 Then
        --普通标本
        For R_Result In V_Result Loop
          If (R_Result.隐私项目 = 1 And 显示隐私项目_In = 1) Or R_Result.隐私项目 = 0 Then
            V_Loop := V_Loop + 1;
            Select 电子病历内容_Id.Nextval Into V_Nextid From Dual;
            If Instr(R_Result.结果参考, Chr(13) || Chr(10)) > 0 Then
              V_检验结果 := R_Result.检验项目 || R_Result.检验结果 || R_Result.单位 || R_Result.结果标志 ||
                        Substr(R_Result.结果参考, 1, Instr(R_Result.结果参考, Chr(13) || Chr(10)) - 1);
            Else
              V_检验结果 := R_Result.检验项目 || R_Result.检验结果 || R_Result.单位 || R_Result.结果标志 || R_Result.结果参考;
            End If;
            Insert Into 电子病历内容
              (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
            Values
              (V_Nextid, V_文件id, 1, 0, V_当前父id, V_Loop, 2, V_Loop, Null, 0, Null, V_检验结果, 1);
          End If;
        End Loop;
      Else
        --微生物标本
        If R_Source.内容文本 = '检验结果' Then
          For R_Result In V_Mresult Loop
            V_Loop := V_Loop + 1;
            Select 电子病历内容_Id.Nextval Into V_Nextid From Dual;
            V_检验结果 := R_Result.项目 || R_Result.耐药 || R_Result.用法用量1 || R_Result.血药浓度1 || R_Result.尿药浓度1;
            Insert Into 电子病历内容
              (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
            Values
              (V_Nextid, V_文件id, 1, 0, V_当前父id, V_Loop, 2, V_Loop, Null, 0, Null, V_检验结果, 1);
          End Loop;
        Elsif R_Source.内容文本 = '一级报告' Then
          Select 一级报告 Into V_报告 From 检验标本记录 Where ID = 检验标本id_In;
          Insert Into 电子病历内容
            (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
          Values
            (V_Nextid, V_文件id, 1, 0, V_当前父id, V_Loop, 2, V_Loop, Null, 0, Null, V_报告, 1);
        Elsif R_Source.内容文本 = '二级报告' Then
          Select 二级报告 Into V_报告 From 检验标本记录 Where ID = 检验标本id_In;
          Insert Into 电子病历内容
            (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
          Values
            (V_Nextid, V_文件id, 1, 0, V_当前父id, V_Loop, 2, V_Loop, Null, 0, Null, V_报告, 1);
        Elsif R_Source.内容文本 = '三级报告' Then
          Select 三级报告 Into V_报告 From 检验标本记录 Where ID = 检验标本id_In;
          Insert Into 电子病历内容
            (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
          Values
            (V_Nextid, V_文件id, 1, 0, V_当前父id, V_Loop, 2, V_Loop, Null, 0, Null, V_报告, 1);
        End If;
      End If;
    End If;

  End Loop;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_检验报告单_Update;
/

--90780:胡俊勇,2015-11-20,回退输液药品医嘱发送输液配药记录状不对
Create Or Replace Procedure Zl_病人医嘱记录_回退
(
  医嘱id_In     病人医嘱记录.Id%Type,
  Flag_In       Number := 0,
  医嘱内容_In   病人医嘱记录.医嘱内容%Type := Null,
  操作类型_In   病人医嘱状态.操作类型%Type := Null,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null
  --功能：回退住院医嘱的状态操作或发送操作(回退重整操作通过调用Zl_病人医嘱记录_批量回退来进行)
  --参数：医嘱ID_IN=一组医嘱ID
  --      FLAG_IN=附加数据。回退停止：0=清除执行终止时间,1=保留现有的执行终止时间。
  --      医嘱内容_IN=该过程被批量回退调用时才用，用于错误提示。
  --      操作类型_IN=该过程被批量回退调用时才用，用于核对回退数据。0-回退发送,n=回退具体医嘱操作
) Is
  --包含指定医嘱的操作记录,第一条为要回退的内容(状态操作优先)
  --临嘱不回退发送后的自动停止,在回退发送时自动回退停止操作
  Cursor c_Rolladvice Is
    Select b.操作人员, b.操作时间, 0 As 发送号, b.操作类型, 0 As 执行状态, Sysdate + Null As 首次时间, Sysdate + Null As 末次时间, a.上次执行时间, a.医嘱期效,
           a.诊疗类别 As 类别, a.诊疗项目id, Null As 类型, a.病人id, a.主页id, a.婴儿, 0 As 记录性质, 0 As 门诊记帐, 0 As 开嘱科室id, a.审核标记, a.开嘱医生,
           a.执行科室id
    From 病人医嘱记录 A, 病人医嘱状态 B
    Where a.Id = b.医嘱id And (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) And
          (Nvl(a.医嘱期效, 0) = 0 And b.操作类型 Not In (1, 2, 3) Or Nvl(a.医嘱期效, 0) = 1 And b.操作类型 Not In (1, 2, 3, 8))
    Union
    Select b.发送人 As 操作人员, b.发送时间 As 操作时间, b.发送号, -null As 操作类型, b.执行状态, b.首次时间, b.末次时间, a.上次执行时间, a.医嘱期效, c.类别, a.诊疗项目id,
           c.操作类型 As 类型, a.病人id, a.主页id, a.婴儿, b.记录性质, b.门诊记帐, a.开嘱科室id, a.审核标记, a.开嘱医生, a.执行科室id
    From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C
    Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In)
    Order By 操作时间 Desc, 发送号;
  r_Rolladvice c_Rolladvice%RowType;

  --方式同c_Rolladvice，只取发送部份用了自动回退处理
  Cursor c_Rollsend(v_发送号 病人医嘱发送.发送号%Type) Is
    Select Distinct b.医嘱id, b.发送时间 As 操作时间, b.发送号, b.执行状态, a.诊疗类别 As 类别, c.当前病区id As 病人病区id, a.病人科室id,
                    b.执行部门id As 执行科室id
    From 病人医嘱记录 A, 病人医嘱发送 B, 病案主页 C
    Where a.Id = b.医嘱id And b.发送号 = v_发送号 And a.病人id = c.病人id And a.主页id = c.主页id And
          (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In)
    Order By b.发送时间 Desc, b.发送号;

  --根据医嘱及发送NO求出本次回退要销帐的费用记录
  --一组医嘱并不是都填写了发送记录,且可能NO不同(药品有,用法煎法不一定有)
  --不管发送记录的计费状态(可能无需计费),有费用记录自然关联出来
  --费用只求价格父号为空的,以便取序号销帐
  --只管记录状态为1的费用,对于已销帐或部份销帐的记录,不再处理；其中"记录状态=3"的读取出来仅用于判断，不处理。
  Cursor c_Rollmoneyout
  (
    v_发送号    病人医嘱发送.发送号%Type,
    v_医嘱id    病人医嘱记录.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.记录状态, a.No, a.序号, a.收费类别, a.执行状态, d.跟踪在用, a.执行部门id, a.记录性质
    From 门诊费用记录 A, Table(t_Adviceids) B, 病人医嘱发送 C, 材料特性 D
    Where c.医嘱id = b.Column_Value And c.发送号 = v_发送号 And a.医嘱序号 = b.Column_Value And
          (a.医嘱序号 = v_医嘱id Or Nvl(v_医嘱id, 0) = 0) And a.记录状态 In (0, 1, 3) And a.No = c.No And a.记录性质 = c.记录性质 And
          a.价格父号 Is Null And a.收费细目id = d.材料id(+)
    Order By a.No, a.序号;

  Cursor c_Rollmoneyin
  (
    v_发送号    病人医嘱发送.发送号%Type,
    v_医嘱id    病人医嘱记录.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.记录状态, a.No, a.序号, a.收费类别, a.执行状态, d.跟踪在用, a.执行部门id, a.记录性质
    From 住院费用记录 A, Table(t_Adviceids) B, 病人医嘱发送 C, 材料特性 D
    Where c.医嘱id = b.Column_Value And c.发送号 = v_发送号 And a.医嘱序号 = b.Column_Value And
          (a.医嘱序号 = v_医嘱id Or Nvl(v_医嘱id, 0) = 0) And a.记录状态 In (0, 1, 3) And a.No = c.No And a.记录性质 = c.记录性质 And
          a.价格父号 Is Null And a.收费细目id = d.材料id(+)
    Order By a.No, a.序号;

  --取发送住院记帐时自动发放的卫材(还没有退料的)
  Cursor c_Stuff_Drug(v_费用id 药品收发记录.费用id%Type) Is
    Select ID
    From 药品收发记录
    Where 费用id = v_费用id And (记录状态 = 1 Or Mod(记录状态, 3) = 0) And 审核人 Is Not Null
    Order By 药品id;

  --用于处理特殊医嘱的回退
  Cursor c_Patilog
  (
    v_病人id 病人变动记录.病人id%Type,
    v_主页id 病人变动记录.主页id%Type
  ) Is
    Select *
    From 病人变动记录
    Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null
    Order By 开始时间 Desc;
  r_Patilog c_Patilog%RowType;

  Cursor c_Adviceids Is
    Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
  t_Adviceids t_Numlist;

  v_医嘱状态     病人医嘱记录.医嘱状态%Type;
  v_医嘱期效     病人医嘱记录.医嘱期效%Type;
  v_费用no       病人医嘱发送.No%Type;
  v_费用序号     Varchar2(255);
  v_末次时间     病人医嘱发送.末次时间%Type;
  v_重整时间     病人医嘱状态.操作时间%Type;
  v_操作类型     诊疗项目目录.操作类型%Type;
  v_执行频率     诊疗项目目录.执行频率%Type;
  v_上次时间     病人医嘱记录.上次执行时间%Type;
  v_执行时间     病人医嘱记录.执行时间方案%Type;
  v_开始执行时间 病人医嘱记录.开始执行时间%Type;
  v_上次打印时间 病人医嘱记录.上次打印时间%Type;
  v_频率间隔     病人医嘱记录.频率间隔%Type;
  v_间隔单位     病人医嘱记录.间隔单位%Type;
  v_发送号       病人医嘱发送.发送号%Type;
  n_护理等级id   病人变动记录.护理等级id%Type;
  d_开始时间     病人变动记录.开始时间%Type;
  d_操作时间     病人医嘱状态.操作时间%Type;
  v_Tmp发送号    病人医嘱发送.发送号%Type;
  n_执行         Number;

  Intdigit   Number(3);
  v_Update   Number(1);
  v_Count    Number(5);
  v_Temp     Varchar2(2000);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_Time     Varchar2(4000);
  n_Blndo    Number;

  v_Error Varchar2(2000);
  Err_Custom Exception;

  Function Checkmoneyundo
  (
    v_No       住院费用记录.No%Type,
    v_记录性质 住院费用记录.记录性质%Type,
    v_序号     住院费用记录.序号%Type,
    n_场合     Number := 0 --0住院，1门诊
  ) Return Number Is
    n_Num      Number;
    n_执行状态 Number;
  Begin
    n_Num := 0;
    If n_场合 = 0 Then
      Select Nvl(Sum(Nvl(付数, 1) * 数次), 0) As 数量
      Into n_Num
      From 住院费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 In (2, 3);
      Select Nvl(执行状态, 0)
      Into n_执行状态
      From 住院费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 = 3;
    Else
      Select Nvl(Sum(Nvl(付数, 1) * 数次), 0) As 数量
      Into n_Num
      From 门诊费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 In (2, 3);
      Select Nvl(执行状态, 0)
      Into n_执行状态
      From 门诊费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 = 3;
    End If;
    If n_Num <> 0 Then
      n_Num := 1;
    End If;
    --如果主记录是已执行（部分执行的）则不自动退。
    If n_执行状态 <> 0 Then
      n_Num := 0;
    End If;
    Return(n_Num);
  End;
Begin
  v_Tmp发送号 := -1;
  Open c_Rolladvice;
  Loop
    Fetch c_Rolladvice
      Into r_Rolladvice;
    If c_Rolladvice%RowCount = 0 Then
      Close c_Rolladvice;
      v_Error := Nvl(医嘱内容_In, '该医嘱') || '当前没有可以回退的内容。';
      Raise Err_Custom;
    End If;
    Exit When c_Rolladvice%NotFound;
    Exit When d_操作时间 <> r_Rolladvice.操作时间 And d_操作时间 Is Not Null;
    d_操作时间 := r_Rolladvice.操作时间;
  
    --批量回退调用时判断
    If 医嘱内容_In Is Not Null Then
      If Nvl(r_Rolladvice.操作类型, 0) <> Nvl(操作类型_In, 0) Then
        v_Error := Nvl(医嘱内容_In, '该医嘱') || '不能与当前医嘱一起回退，可能该医嘱已经执行了其他操作。';
        Raise Err_Custom;
      End If;
    End If;
  
    --一组发送号只执行一次
    If v_Tmp发送号 <> r_Rolladvice.发送号 Then
      v_Tmp发送号 := r_Rolladvice.发送号;
      n_执行      := 1;
    Else
      n_执行 := 0;
    End If;
  
    If n_执行 = 1 Then
      Open c_Adviceids;
      Fetch c_Adviceids Bulk Collect
        Into t_Adviceids;
      Close c_Adviceids;
    
      If r_Rolladvice.发送号 = 0 Then
        --回退医嘱状态操作(以时间关键字)
        --4-作废；5-重整；6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果;13-停嘱申请
        ------------------------------------------------------------------
        --最多只能退回到校对状态
        If r_Rolladvice.操作类型 = 3 Then
          v_Error := Nvl(医嘱内容_In, '该医嘱') || '当前处于通过校对状态，不能再回退。';
          Raise Err_Custom;
        Elsif r_Rolladvice.操作类型 = 4 And Nvl(r_Rolladvice.婴儿, 0) = 0 Then
          If r_Rolladvice.类别 = 'H' Then
            Select 操作类型, 执行频率 Into v_操作类型, v_执行频率 From 诊疗项目目录 Where ID = r_Rolladvice.诊疗项目id;
            If v_操作类型 = '1' And v_执行频率 = '2' Then
              v_Error := '护理等级作废后不能再回退。';
              Raise Err_Custom;
            End If;
          End If;
        End If;
      
        --检查是否回退最近次重整之前的操作
        If r_Rolladvice.操作类型 <> 5 Then
          --取最后重整时间
          Select Nvl(医嘱重整时间, To_Date('1900-01-01', 'YYYY-MM-DD'))
          Into v_重整时间
          From 病案主页
          Where 病人id = r_Rolladvice.病人id And 主页id = r_Rolladvice.主页id;
        
          If r_Rolladvice.操作时间 < v_重整时间 Then
            v_Error := '该病人最近次重整之前的操作不能再回退。';
            Raise Err_Custom;
          End If;
        End If;
      
        --删除(该组医嘱)最近的状态操作记录
        Delete /*+ Rule*/
        From 病人医嘱状态
        Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And 操作时间 = r_Rolladvice.操作时间;
      
        --取删除后应恢复的医嘱状态
        Select 操作类型
        Into v_医嘱状态
        From 病人医嘱状态
        Where 操作时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = 医嘱id_In) And 医嘱id = 医嘱id_In;
      
        --恢复(该组医嘱)回退后的状态
        Update 病人医嘱记录 Set 医嘱状态 = v_医嘱状态 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
      
        --其它额外的处理
        If r_Rolladvice.操作类型 = 8 Then
          --被超期发送收回过的医嘱 ，如果是销帐申请模式，则判断对应的“病人费用销帐”申请是否取消，是则允许回退，否则不允许，
          --                       如果是产生负数费用模式，则不允许再回退。
          --可能超期发送收回时被全部收回(无上次执行时间)
          Select /*+ Rule*/
           Nvl(Count(*), 0)
          Into v_Count
          From 病人医嘱记录 A, 病人医嘱发送 B
          Where b.医嘱id = a.Id And (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) And
                b.发送号 =
                (Select Max(发送号) From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(t_Adviceids))) And
                a.执行终止时间 Is Not Null And ((a.上次执行时间 < b.末次时间) Or (a.上次执行时间 Is Null And b.末次时间 Is Not Null));
          If v_Count > 0 Then
            If zl_GetSysParameter('超期收回产生负数费用', 1254) = '1' Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '已被超期发送收回，不能再撤消停止操作。';
              Raise Err_Custom;
            Else
              --如果已经取消销帐申请，则允许回退.
              Select Count(1)
              Into v_Count
              From 病人费用销帐 A, 住院费用记录 B, 病人医嘱记录 C
              Where a.费用id = b.Id And c.Id = b.医嘱序号 And (c.Id = 医嘱id_In Or c.相关id = 医嘱id_In);
              If v_Count > 0 Then
                v_Error := Nvl(医嘱内容_In, '该医嘱') || '已被超期发送收回，不能再撤消停止操作。';
                Raise Err_Custom;
              Else
                --得到上次执行时间等信息
                Select 上次执行时间, 执行时间方案, 开始执行时间, 上次打印时间, 频率间隔, 间隔单位
                Into v_上次时间, v_执行时间, v_开始执行时间, v_上次打印时间, v_频率间隔, v_间隔单位
                From 病人医嘱记录
                Where ID = 医嘱id_In;
                v_上次时间 := To_Date(To_Char(v_上次时间 + 1 / 24 / 60 / 60, 'yyyy-MM-dd hh24:mi:ss'), 'yyyy-MM-dd hh24:mi:ss');
              
                --修改上次执行时间为收回后的末次执行时间。
                v_末次时间 := Null;
                Begin
                  --一组医嘱的发送首末时间相同,一并给药是取最小的
                  --取相关ID为NULL的医嘱的发送记录的时间
                  --但给药途径或中药用法可能未填写发送记录
                  Select /*+ Rule*/
                   末次时间, 发送号
                  Into v_末次时间, v_发送号
                  From 病人医嘱发送
                  Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And
                        发送号 = (Select Max(发送号)
                               From 病人医嘱发送
                               Where 医嘱id In (Select Column_Value From Table(t_Adviceids))) And Rownum = 1;
                Exception
                  When Others Then
                    Null;
                End;
                Update 病人医嘱记录 Set 上次执行时间 = v_末次时间 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
              
                --还原医嘱执行时间
                Select Zl_Adviceexetimes(医嘱id_In, v_上次时间, v_末次时间, v_执行时间, v_开始执行时间, v_上次打印时间, v_频率间隔, v_间隔单位, 0)
                Into v_Time
                From Dual;
                Insert Into 医嘱执行时间
                  (要求时间, 医嘱id, 发送号)
                  Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), 医嘱id_In, v_发送号
                  From Table(f_Str2list(v_Time));
              End If;
            End If;
          End If;
        
          --护理等级变动，后续有其他变动时，不允许回退
          If r_Rolladvice.类别 = 'H' And Nvl(r_Rolladvice.婴儿, 0) = 0 Then
            Select 操作类型, 执行频率 Into v_操作类型, v_执行频率 From 诊疗项目目录 Where ID = r_Rolladvice.诊疗项目id;
            If v_操作类型 = '1' And v_执行频率 = '2' Then
              Select Count(*), Max(a.护理等级id), Max(a.开始时间)
              Into v_Count, n_护理等级id, d_开始时间
              From 病人变动记录 A
              Where a.病人id = r_Rolladvice.病人id And a.主页id = r_Rolladvice.主页id And a.开始原因 = 6 And a.终止时间 Is Null And
                    a.附加床位 = 0;
              --如果没有找到最后一条是护理等级变动则禁止
              If v_Count = 0 Then
                --医嘱护理等级和入住时候的护理等级一致时要单独判断
                Select Count(*)
                Into v_Count
                From 病人变动记录 A
                Where a.病人id = r_Rolladvice.病人id And a.主页id = r_Rolladvice.主页id And a.开始原因 = 6;
                If v_Count > 0 Then
                  v_Error := '由于护理等级医嘱停止后该病人已经产生了其他变动记录,不能回退该医嘱的停止操作。';
                  Raise Err_Custom;
                End If;
              Else
                --如果n_护理等级ID为Null，则检查是否是当前回退的医嘱对应的变动记录,目的是有多个护理等级医嘱时要求按顺序回退。
                --如果n_护理等级ID不为Null，则有可能是校对下一条护理等级时，自动停止的，未产生变动记录，
                --     则需要检查当前最后一条变动的护理等级ID是否是当前医嘱的护理等级ID,目的是有多个护理等级医嘱时要求按顺序回退，如果是则不需要再撤销最后一次变动，直接回退医嘱即可。
                If n_护理等级id Is Null Then
                  Select Count(*)
                  Into v_Count
                  From 病人变动记录 B, 病人医嘱计价 C
                  Where b.病人id = r_Rolladvice.病人id And b.主页id = r_Rolladvice.主页id And c.医嘱id = 医嘱id_In And
                        c.收费细目id = b.护理等级id And b.终止时间 = d_开始时间 And b.终止原因 = 6 And b.附加床位 = 0;
                Else
                  --开始时间只取分钟对比，校对的时候护理等级的开始时间是医嘱开始时间+当前时间的秒钟
                  Select Count(*)
                  Into v_Count
                  From 病人医嘱计价 C, 病人医嘱记录 A
                  Where a.Id = c.医嘱id And a.Id = 医嘱id_In And c.收费细目id = n_护理等级id And
                        a.开始执行时间 = To_Date(To_Char(d_开始时间, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi');
                End If;
                If v_Count = 0 Then
                  v_Error := '您回退的医嘱不是最后一条护理等级医嘱，请将后面的护理等级医嘱作废后再回退本条医嘱。';
                  Raise Err_Custom;
                End If;
              
                If n_护理等级id Is Null Then
                  --当前操作人员
                  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
                    v_人员编号 := 操作员编号_In;
                    v_人员姓名 := 操作员姓名_In;
                  Else
                    v_Temp     := Zl_Identity;
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                    v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
                    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                  End If;
                
                  Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, '1', Null, Null, '护理等级变动');
                End If;
              End If;
            End If;
          End If;
        
          --回退医嘱停止时,清空停嘱医生和时间,如果是实习医师申请后审核的，则恢复待审核状态
          Update 病人医嘱记录
          Set 执行终止时间 = Decode(Flag_In, 1, 执行终止时间, Null), 停嘱医生 = Null, 停嘱时间 = Null,
              审核标记 = Decode(r_Rolladvice.审核标记, 3, 2, r_Rolladvice.审核标记)
          Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        Elsif r_Rolladvice.操作类型 = 9 Then
          --回退医嘱确认停止时,检查是否已打印停嘱时间
          Select /*+ Rule*/
           Count(*)
          Into v_Count
          From 病人医嘱打印
          Where 打印标记 = 1 And 医嘱id In (Select Column_Value From Table(t_Adviceids));
          If v_Count > 0 Then
            v_Error := Nvl(医嘱内容_In, '该医嘱') || '的停嘱时间已经打印，不能再撤消确认停止操作。';
            Raise Err_Custom;
          End If;
        
          --回退医嘱确认停止时,清空停嘱医生和时间
          Update 病人医嘱记录 Set 确认停嘱时间 = Null, 确认停嘱护士 = Null Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        Elsif r_Rolladvice.操作类型 = 10 Then
          --回退标注皮试结果,同时删除过敏登记(+)或(-),根据记录时间
          Delete From 病人过敏记录
          Where 病人id = r_Rolladvice.病人id And Nvl(主页id, 0) = Nvl(r_Rolladvice.主页id, 0) And 记录时间 = r_Rolladvice.操作时间;
        
          Update 病人医嘱记录 Set 皮试结果 = Null Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        Elsif r_Rolladvice.操作类型 = 13 Then
          If Instr(r_Rolladvice.开嘱医生, '/') > 0 Then
            Update 病人医嘱记录 Set 审核标记 = 1 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
          Else
            Update 病人医嘱记录 Set 审核标记 = Null Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
          End If;
        End If;
      Else
        --回退医嘱发送(以发送号关键字)
        ------------------------------------------------------------------
        --当前操作人员
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      
        --检查是否是输液配液记录，并是否已经锁定，如果查询有数据说明是配液记录
        Begin
          Select Decode(max(是否锁定), 1, 1, 0)
          Into v_Count
          From 输液配药记录
          Where 医嘱id = 医嘱id_In And 发送号 = r_Rolladvice.发送号;
        Exception
          When Others Then
            v_Count := -1;
        End;
      
        If v_Count = 1 Then
          v_Error := '医嘱"' || 医嘱内容_In || '"是输液药品，已经被输液配置中心锁定，不能回退发送。';
          Raise Err_Custom;
        Elsif v_Count = 0 Then
          Zl_输液配药记录_医嘱回退(医嘱id_In, r_Rolladvice.发送号, v_人员姓名, Sysdate);
        End If;
      
        --本科发送自动执行时，回退也自动回退执行(仅护士站有此功能)
        --非跟踪在用的卫材医嘱，同普通医嘱执行处理
        Select 医嘱期效 Into v_医嘱期效 From 病人医嘱记录 Where ID = 医嘱id_In;
        If Substr(zl_GetSysParameter('本科执行自动完成', 1254), v_医嘱期效 + 1, 1) = '1' Then
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
        
          For r_Rollsend In c_Rollsend(r_Rolladvice.发送号) Loop
            If Nvl(r_Rollsend.执行状态, 0) = 1 And
               (Nvl(r_Rollsend.执行科室id, 0) = Nvl(r_Rollsend.病人病区id, 0) Or
                Nvl(r_Rollsend.执行科室id, 0) = Nvl(r_Rollsend.病人科室id, 0)) Then
            
              --医嘱的执行状态
              Update 病人医嘱发送 Set 执行状态 = 0 Where 发送号 = r_Rollsend.发送号 And 医嘱id = r_Rollsend.医嘱id;
              v_Update := 1;
            
              If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 0 Then
                --费用的执行状态
                For r_Rollmoney In c_Rollmoneyin(r_Rollsend.发送号, r_Rollsend.医嘱id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.记录状态 <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1) And
                       Not r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --普通费用直接取消执行状态，不含药品和跟踪在用的卫材
                      Update 住院费用记录
                      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
                      Where NO = r_Rollmoney.No And 记录性质 = r_Rolladvice.记录性质 And 记录状态 = r_Rollmoney.记录状态 And
                            Nvl(价格父号, 序号) = r_Rollmoney.序号 And 医嘱序号 = r_Rollsend.医嘱id;
                    Elsif r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1 Then
                      --跟踪在用的卫材，当系统参数为自动发料时，才自动退料
                      If Nvl(zl_GetSysParameter(33), '0') = '1' Then
                        For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_材料收发记录_部门退料(r_Stuff.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, 0, v_人员姓名);
                        End Loop;
                      End If;
                    Elsif r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --住院科室发药的药品自动退药
                      If r_Rollmoney.执行部门id = r_Rollsend.病人病区id Or r_Rollmoney.执行部门id = r_Rollsend.病人科室id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_药品收发记录_部门退药(r_Drug.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 2);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              Else
                --住院病人费用发送到门诊的情况，病人来源都是住院的
                For r_Rollmoney In c_Rollmoneyout(r_Rollsend.发送号, r_Rollsend.医嘱id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.记录状态 <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号, 1);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1) And
                       Not r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --普通费用直接取消执行状态，不含药品和跟踪在用的卫材
                      Update 门诊费用记录
                      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
                      Where NO = r_Rollmoney.No And 记录性质 = r_Rolladvice.记录性质 And 记录状态 = r_Rollmoney.记录状态 And
                            Nvl(价格父号, 序号) = r_Rollmoney.序号 And 医嘱序号 = r_Rollsend.医嘱id;
                    Elsif r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1 Then
                      If Nvl(zl_GetSysParameter(33), '0') = '1' Then
                        --跟踪在用的卫材，当系统参数为自动发料时，才自动退料
                        For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_材料收发记录_部门退料(r_Stuff.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, 0, v_人员姓名);
                        End Loop;
                      End If;
                    Elsif r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --本科室发药的药品自动退药
                      If r_Rollmoney.执行部门id = r_Rollsend.病人病区id Or r_Rollmoney.执行部门id = r_Rollsend.病人科室id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_药品收发记录_部门退药(r_Drug.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 1);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              End If;
            End If;
          End Loop;
        End If;
        ------------------------------------------------------------------
        --被超期收回的长期药品医嘱不允许回退(再退费用就多退了)
        If Nvl(r_Rolladvice.医嘱期效, 0) = 0 Then
          If r_Rolladvice.上次执行时间 Is Not Null And r_Rolladvice.末次时间 Is Not Null Then
            If r_Rolladvice.上次执行时间 < r_Rolladvice.末次时间 Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '最近超期发送的内容已被收回，不能再回退。';
              Raise Err_Custom;
            End If;
          Elsif r_Rolladvice.上次执行时间 Is Null And r_Rolladvice.末次时间 Is Not Null Then
            --长嘱可能被全部超期收回
            v_Error := Nvl(医嘱内容_In, '该医嘱') || '未被发送，或发送的内容已被全部超期收回，不能再回退。';
            Raise Err_Custom;
          End If;
        End If;
      
        If Nvl(r_Rolladvice.执行状态, 0) In (1, 3) And v_Update <> 1 Then
          --1-完全执行;3-正在执行
          v_Error := Nvl(医嘱内容_In, '该医嘱') || '最近发送的内容已经执行或正在执行，不能回退。';
          Raise Err_Custom;
        Else
          --如果相关医嘱已执行，则也要限制回退（例如：检验的采集方式）
          Select /*+ Rule*/
           Count(1)
          Into v_Count
          From 病人医嘱发送
          Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And 执行状态 In (1, 3) And
                发送号 =
                (Select Max(发送号) From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(t_Adviceids)));
          If v_Count > 0 Then
            v_Error := Nvl(医嘱内容_In, '该医嘱') || '最近发送的内容已经执行或正在执行，不能回退。';
            Raise Err_Custom;
          End If;
        End If;
      
        ------------------------------------------------------------------
        --将该组医嘱的费用销帐(按一组医嘱可能有不同NO处理)
        --如果原始费用已被销帐(或部分销帐),调用过程中有判断
        v_费用no   := Null;
        v_费用序号 := Null;
        If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 0 Then
          For r_Rollmoney In c_Rollmoneyin(r_Rolladvice.发送号, Null, t_Adviceids) Loop
            --对应的费用已执行
            If Nvl(r_Rollmoney.执行状态, 0) <> 0 Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '发送的费用单据"' || r_Rollmoney.No || '"中的内容已被部分或完全执行，不能回退。';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.记录状态 <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号);
            End If;
            If n_Blndo > 0 Then
              --这种仅用于判断部分退药
              If v_费用no <> r_Rollmoney.No And v_费用序号 Is Not Null Then
                Zl_住院记帐记录_Delete(v_费用no, Substr(v_费用序号, 2), v_人员编号, v_人员姓名, 2, 0, 0);
                v_费用序号 := Null;
              End If;
              v_费用no   := r_Rollmoney.No;
              v_费用序号 := v_费用序号 || ',' || r_Rollmoney.序号;
            End If;
          End Loop;
        Else
          For r_Rollmoney In c_Rollmoneyout(r_Rolladvice.发送号, Null, t_Adviceids) Loop
            --对应的费用已执行
            If Nvl(r_Rollmoney.执行状态, 0) <> 0 And Not (Nvl(r_Rollmoney.执行状态, 0) = -1 And Nvl(r_Rollmoney.记录状态, 0) = 0) Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '发送的费用单据"' || r_Rollmoney.No || '"中的内容已被部分或完全执行，不能回退。';
              Raise Err_Custom;
            End If;
            --收费单据已收费
            If r_Rollmoney.记录状态 = 1 And Not (r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 1) Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '发送的门诊单据"' || r_Rollmoney.No || '"已收费，不能回退。';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.记录状态 <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号, 1);
            End If;
            If n_Blndo > 0 Then
              --这种仅用于判断部分退药
              If v_费用no <> r_Rollmoney.No And v_费用序号 Is Not Null Then
                If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 1 Then
                  --住院发送为门诊记帐(如果是门诊医生发送为门诊记帐，门诊医嘱没有回退功能)
                  Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
                Else
                  Zl_门诊划价记录_Delete(v_费用no, Substr(v_费用序号, 2));
                End If;
                v_费用序号 := Null;
              End If;
              v_费用no   := r_Rollmoney.No;
              v_费用序号 := v_费用序号 || ',' || r_Rollmoney.序号;
            End If;
          End Loop;
        End If;
        If v_费用序号 Is Not Null And v_费用no Is Not Null Then
          v_费用序号 := Substr(v_费用序号, 2);
          If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 0 Then
            Zl_住院记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名, 2, 0, 0);
          Elsif r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 1 Then
            --住院发送为门诊记帐(如果是门诊医生发送为门诊记帐，门诊医嘱没有回退功能)
            Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
          Else
            Zl_门诊划价记录_Delete(v_费用no, v_费用序号);
          End If;
        End If;
        --输血医嘱先删除病人医嘱附费
        Delete From 病人医嘱附费 Where 发送号 = r_Rolladvice.发送号 And 医嘱id = 医嘱id_In;
      
        --删除医嘱执行时间 (仅主医嘱ID才产生了记录)
        Delete From 医嘱执行时间 Where 发送号 = r_Rolladvice.发送号 And 医嘱id = 医嘱id_In;
      
        --删除发送记录(该组医嘱的)
        Delete /*+ Rule*/
        From 病人医嘱发送
        Where 发送号 = r_Rolladvice.发送号 And 医嘱id In (Select Column_Value From Table(t_Adviceids));
      
        --标记(该组医嘱)上次执行时间(以上次发送的末次执行时间)
        --所有长嘱(包括持续性长嘱)发送时都填写了末次时间
        --临嘱可能没有，且只可能发送了一次。
        v_末次时间 := Null;
        Begin
          --一组医嘱的发送首末时间相同,一并给药是取最小的
          --取相关ID为NULL的医嘱的发送记录的时间
          --但给药途径或中药用法可能未填写发送记录
          Select /*+ Rule*/
           末次时间
          Into v_末次时间
          From 病人医嘱发送
          Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And
                发送号 =
                (Select Max(发送号) From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(t_Adviceids))) And
                Rownum = 1;
        Exception
          When Others Then
            Null;
        End;
        Update 病人医嘱记录 Set 上次执行时间 = v_末次时间 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
      
        --回退临嘱发送时，同时自动回退停止
        If Nvl(r_Rolladvice.医嘱期效, 0) = 1 Then
          --删除(该组临嘱)最近的停止状态操作记录
          Delete /*+ Rule*/
          From 病人医嘱状态
          Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And
                操作时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = 医嘱id_In) And 操作类型 = 8;
          --r_RollAdvice.操作时间:因发送时间可能不与自动停止时间相同。
        
          --取删除后应恢复的医嘱状态
          Select 操作类型
          Into v_医嘱状态
          From 病人医嘱状态
          Where 操作时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = 医嘱id_In) And 医嘱id = 医嘱id_In;
        
          --恢复(该组医嘱)回退后的状态
          Update 病人医嘱记录
          Set 医嘱状态 = v_医嘱状态, 执行终止时间 = Null, 停嘱医生 = Null, 停嘱时间 = Null
          Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        End If;
      
        --住院特殊医嘱发送后的回退(3-转科;5-出院;6-转院,11-死亡)
        If r_Rolladvice.类别 = 'Z' And Instr(',3,5,6,11,', Nvl(r_Rolladvice.类型, '0')) > 0 And Nvl(r_Rolladvice.婴儿, 0) = 0 Then
          Open c_Patilog(r_Rolladvice.病人id, r_Rolladvice.主页id);
          Fetch c_Patilog
            Into r_Patilog;
          If c_Patilog%Found Then
            If r_Rolladvice.类型 = '3' And r_Patilog.开始原因 = 3 Then
              --取消病人转科状态
              If r_Patilog.开始时间 Is Null Then
                --转科医嘱的特殊处理，当一个病人有两条转科医嘱时，只能回退最近的一条,70443
                Select Count(1)
                Into v_Count
                From 病人医嘱记录 A, 诊疗项目目录 B
                Where a.诊疗项目id = b.Id And a.病人id = r_Rolladvice.病人id And a.主页id = r_Rolladvice.主页id And a.诊疗类别 = 'Z' And
                      b.操作类型 = '3' And a.医嘱状态 = 8 And
                      a.开始执行时间 > (Select 开始执行时间 From 病人医嘱记录 Where ID = 医嘱id_In);
                If v_Count = 0 Then
                  Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, Null, Null, Null, '转科');
                Else
                  v_Error := '病人转科已经入科，不能再回退。';
                  Raise Err_Custom;
                End If;
              Else
                v_Error := '病人转科已经入科，不能再回退。';
                Raise Err_Custom;
              End If;
            Elsif r_Rolladvice.类型 In ('5', '6', '11') And r_Patilog.开始原因 = 10 Then
              --取消病人预出院状态
              Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, Null, Null, Null, '预出院');
            End If;
          End If;
          Close c_Patilog;
        End If;
      
        --回退病历时机
        --1.特殊事件(只有一条医嘱记录)：手术，7-会诊,8-抢救,11-死亡
        If r_Rolladvice.类别 = 'F' Or r_Rolladvice.类别 = 'Z' And Instr('7,8,11', r_Rolladvice.类型) > 0 Then
          Zl_电子病历时机_Delete(r_Rolladvice.病人id, r_Rolladvice.主页id, '医嘱', r_Rolladvice.开嘱科室id, 医嘱id_In);
        End If;
      
        --2.额外处理：知情同意书(手术相关的知情同意需再次调用，因为附加手术或麻醉项目可能有关联的知情同意书)
        If Instr('C,D,E,F,G,K,L', r_Rolladvice.类别) > 0 Then
          For R In (Select a.Id, a.诊疗类别 From 病人医嘱记录 A Where a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) Loop
            --相关id的一组医嘱不一定是这个类别的，所以要再判断一次类别
            If Instr('C,D,E,F,G,K,L', r.诊疗类别) > 0 Then
              Zl_电子病历时机_Delete(r_Rolladvice.病人id, r_Rolladvice.主页id, '医嘱', r_Rolladvice.开嘱科室id, r.Id);
            End If;
          End Loop;
        End If;
      End If;
    End If;
    Exit When r_Rolladvice.发送号 = 0;
  End Loop;
  Close c_Rolladvice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_回退;
/

--90466:刘尔旋,2015-11-19,支付宝结帐修改
Create Or Replace Procedure Zl_病人预交记录_Insert
(
  Id_In         病人预交记录.Id%Type,
  单据号_In     病人预交记录.No%Type,
  票据号_In     票据使用明细.号码%Type,
  病人id_In     病人预交记录.病人id%Type,
  主页id_In     病人预交记录.主页id%Type,
  科室id_In     病人预交记录.科室id%Type,
  金额_In       病人预交记录.金额%Type,
  结算方式_In   病人预交记录.结算方式%Type,
  结算号码_In   病人预交记录.结算号码%Type,
  缴款单位_In   病人预交记录.缴款单位%Type,
  单位开户行_In 病人预交记录.单位开户行%Type,
  单位帐号_In   病人预交记录.单位帐号%Type,
  摘要_In       病人预交记录.摘要%Type,
  操作员编号_In 病人预交记录.操作员编号%Type,
  操作员姓名_In 病人预交记录.操作员姓名%Type,
  领用id_In     票据使用明细.领用id%Type,
  预交类别_In   病人预交记录.预交类别%Type := Null,
  卡类别id_In   病人预交记录.卡类别id%Type := Null,
  结算卡序号_In 病人预交记录.结算卡序号%Type := Null,
  卡号_In       病人预交记录.卡号%Type := Null,
  交易流水号_In 病人预交记录.交易流水号%Type := Null,
  交易说明_In   病人预交记录.交易说明%Type := Null,
  合作单位_In   病人预交记录.合作单位%Type := Null,
  收款时间_In   病人预交记录.收款时间%Type := Null,
  操作类型_In   Integer := 0,
  结算性质_In   病人预交记录.结算性质%Type := Null
) As
  ----------------------------------------------
  --操作类型_In:0-正常缴预交;1-存为划价单
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  v_性质   结算方式.性质%Type;
  v_打印id 票据打印内容.Id%Type;
  v_担保   病人信息.担保性质%Type;
  v_Date   Date;
  n_返回值 病人余额.预交余额%Type;
  n_组id   财务缴款分组.Id%Type;
Begin
  v_Date := 收款时间_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  n_组id := Zl_Get组id(操作员姓名_In);

  --插入预交缴款记录
  Insert Into 病人预交记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别,
     卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
  Values
    (Id_In, 单据号_In, 票据号_In, 1, Decode(操作类型_In, 1, 0, 1), 病人id_In, Decode(主页id_In, 0, Null, 主页id_In),
     Decode(科室id_In, 0, Null, 科室id_In), 金额_In, 结算方式_In, 结算号码_In, v_Date, 缴款单位_In, 单位开户行_In, 单位帐号_In, 操作员编号_In, 操作员姓名_In,
     摘要_In, n_组id, 预交类别_In, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 结算性质_In);

  If 操作类型_In = 1 Then
    --暂不处理汇总表
    Return;
  End If;

  --处理票据
  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 2, 单据号_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, 2, 票据号_In, 1, 1, 领用id_In, v_打印id, v_Date, 操作员姓名_In);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  End If;

  --相关汇总表处理

  --病人余额(预交余额现收)
  Begin
    Select 性质 Into v_性质 From 结算方式 Where 名称 = 结算方式_In;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(v_性质, 1) <> 5 Then
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 0) = Nvl(预交类别_In, 0)
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 类型, 预交余额, 费用余额)
      Values
        (病人id_In, 1, Nvl(预交类别_In, 0), 金额_In, 0);
      n_返回值 := 金额_In;
    End If;
    If Nvl(金额_In, 0) = 0 Then
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
    End If;
  End If;

  --人员缴款余额(现收)
  Update 人员缴款余额
  Set 余额 = Nvl(余额, 0) + 金额_In
  Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = 结算方式_In
  Returning 余额 Into n_返回值;

  If Sql%RowCount = 0 Then
    Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, 结算方式_In, 1, 金额_In);
    n_返回值 := 金额_In;
  End If;
  If Nvl(n_返回值, 0) = 0 Then
    Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
  End If;
  --对临时担保的处理
  Select Nvl(担保性质, 0) Into v_担保 From 病人信息 Where 病人id = 病人id_In;
  If v_担保 = 1 And Nvl(金额_In, 0) > 0 Then
    Update 病人信息
    Set 担保额 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, Nvl(担保额, 0) - Nvl(金额_In, 0), Null),
        担保人 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, 担保人, Null),
        担保性质 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, 担保性质, Null)
    Where 病人id = 病人id_In;
  End If;
  If 操作类型_In = 0 Then
    --消息推送;
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_Insert;
/

--90603:许华峰,2015-11-18,同一个原型能否书写多份报告控制
--影像报告原型管理(---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_RptAntetype Is
  --Create By Hwei;
  --2014/11/25
  Type t_Refcur Is Ref Cursor;

  --1.获取文件原型类别
  Procedure p_Get_Antetypelistkind(
    Val Out t_Refcur
	);
  --2.根据文档类型获取文档信息
  Procedure p_Get_Antetypelis_By_Kind(
	Val           Out t_Refcur,
	种类_In      影像报告原型清单.种类%Type,
	Stop_Flag    Number,
	Condition_In Varchar2
	);
  --3.添加一个文档原型
  Procedure p_Add_Antetypelist(
    Id_In           影像报告原型清单.ID%Type,
	种类_In         影像报告原型清单.种类%Type,
	编码_In         影像报告原型清单.编码%Type,
	名称_In         影像报告原型清单.名称%Type,
    设备号_In		影像设备目录.设备号%Type,
	说明_In         影像报告原型清单.说明%Type,
	可否重置页面_In 影像报告原型清单.可否重置页面%Type,
	可否重置格式_In 影像报告原型清单.可否重置格式%Type,
    可否书写多份_In 影像报告原型清单.可否书写多份%Type,
	是否禁用_In     影像报告原型清单.是否禁用%Type,
	创建人_In       影像报告原型清单.创建人%Type,
	内容_In         影像报告原型清单.内容%Type,
	控制选项_In     影像报告原型清单.控制选项%Type,
	专用插件_In     影像报告原型清单.专用插件%Type,
	Copy_Id_In      影像报告原型清单.ID%Type,
	Only_Head_In    Varchar2,
	分组_In         影像报告原型清单.分组%Type
	);
  --4.修改一个文档原型
  Procedure p_Edit_Antetypelist(
    Id_In           影像报告原型清单.ID%Type,
    种类_In         影像报告原型清单.种类%Type,
    编码_In         影像报告原型清单.编码%Type,
    名称_In         影像报告原型清单.名称%Type,
    设备号_In		影像设备目录.设备号%Type,
    说明_In         影像报告原型清单.说明%Type,
    可否重置页面_In 影像报告原型清单.可否重置页面%Type,
    可否重置格式_In 影像报告原型清单.可否重置格式%Type,
    可否书写多份_In 影像报告原型清单.可否书写多份%Type,
    是否禁用_In     影像报告原型清单.是否禁用%Type,
    修改人_In       影像报告原型清单.修改人%Type,
    内容_In         影像报告原型清单.内容%Type,
    控制选项_In     影像报告原型清单.控制选项%Type,
    专用插件_In     影像报告原型清单.专用插件%Type,
    Copy_Id_In      影像报告原型清单.ID%Type,
    Only_Head_In    Varchar2,
    分组_In         影像报告原型清单.分组%Type
	);
  --5.删除一个文件原型
  Procedure p_Del_Antetypelist(
    Id_In 影像报告原型清单.Id%Type
	);
  --6.根据ID获取文件原型
  Procedure p_Get_Antetypelist_By_Id(
	Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	);
  --7.获取原型XML内容
  Procedure p_Get_Antetypelist_Content(
	Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	);
  --8.停用或启用文件原型
  Procedure p_Stop_Antetypelist(
    Id_In 影像报告原型清单.Id%Type
	);

  --9.新增文档种类信息
  Procedure p_Add_Doc_Kind(
    编码_In 影像报告种类.编码%Type,
    名称_In 影像报告种类.名称%Type,
    说明_In 影像报告种类.说明%Type
	);
  --10.删除文档种类信息
  Procedure p_Del_Doc_Kind;
  --11.获取预备提纲信息
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	);
  --12.添加预备提纲信息
  Procedure p_Add_Pre_Outline(
    ID_In   影像报告预备提纲.ID%Type,
	编码_In 影像报告预备提纲.编码%Type,
	名称_In 影像报告预备提纲.名称%Type,
	说明_In 影像报告预备提纲.说明%Type
	);
  --13.删除预备提纲信息
  Procedure p_Del_Pre_Outline;
  --14.获取导出的文档原型信息
  Procedure p_Output_Antetypelist(
    Val Out t_Refcur
	);
  --15.添加原型片段
  Procedure p_Add_Antetype_Fragments(
    原型ID_In 影像报告原型片段.原型ID%Type,
	片段ID_In 影像报告原型片段.片段ID%Type
	);
  --16.删除原型片段
  Procedure p_Del_Antetype_Fragments(
    原型ID_In 影像报告原型片段.原型ID%Type
	);
  --17.获取原型片段
  Procedure p_Get_Antetype_Fragments(
    Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	);
  --18.获取某个原型关联的某个片段分类
  Procedure p_Get_Antetype_f_Byaidfid(
    Val           Out t_Refcur,		
	原型ID_In 影像报告原型片段.原型ID%Type,
    片段ID_In 影像报告原型片段.片段ID%Type
	);
  --19.插入文档原型XML内容
  Procedure p_Edit_Antetypelist_Content(
    Id_In     影像报告原型清单.Id%Type,
	内容_In   影像报告原型清单.内容%Type,
	修改人_In 影像报告原型清单.修改人%Type
	);
  --20.获取所有原型
  Procedure p_Get_All_Antetype_Lists(
    Val Out t_Refcur
	);
  --21.获取已经设置了关联的原型片段类别的信息

  Procedure p_Get_Antetype_Fragments_Info(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	);
  --22.获取选择的类别下面的短语名称
  Procedure p_Get_Selected_Fragments(
	Val           Out t_Refcur,
	原型id_In Varchar2
	);
  --23.获取能复制的原型名称

  Procedure p_Get_Copy_Antetype(
	Val           Out t_Refcur,
	种类_In 影像报告原型清单.种类%Type
	);
  --24.获取原型的分组信息
  Procedure p_Get_Antetype_Category(
	Val           Out t_Refcur,
	种类_In 影像报告原型清单.种类%Type
	);
  --25.根据原型同步范文提纲
  Procedure p_Synchronous_Sample(
    原型id_In 影像报告原型清单.Id%Type
	);
  --26.获取导出的原型列表
  Procedure p_Get_Out_Antetypelist(
    Val Out t_Refcur
	);
  --27.通过编码获取原型种类信息
  Procedure p_Get_Antetype_Kind_By_Code(
	Val           Out t_Refcur,
	编码_In 影像报告种类.编码%Type
	);
  --28.获取事件信息，不包含固定事件
  Procedure p_Get_Doc_Event(
    Val Out t_Refcur
	);
  --29.获取关于原型导出的重复信息

  Procedure p_Get_Antetypelist_Same_Info(
	Val           Out t_Refcur,
	Tablename_In Varchar2,
    Id_In        影像报告原型清单.Id%Type,
    编码_In      Varchar2,
    名称_In      Varchar2
	);
  --30.获取事件重复的信息
  Procedure p_Event_Same_Info(
	Val           Out t_Refcur,	
	Id_In      影像报告事件.Id%Type,
    原型ID_In  影像报告事件.原型ID%Type,
    元素IID_In 影像报告事件.元素IID%Type,
    种类_In    影像报告事件.种类%Type,
    名称_In    影像报告事件.名称%Type,
    编号_In    影像报告事件.编号%Type
	);
  --31.获取原型校验的类别集合
  Procedure p_Get_Process_Kind(
    Val Out t_Refcur
	);

  ----32.获取元素或者提纲的名称集合
  --Procedure p_Get_Antetype_Ele_Section(
  --原型ID_In  影像报告原型清单.Id%Type,
  --Val     Out t_Refcur);

  --33.获取指定原型的文档处理
  Procedure p_Get_Doc_Process_Of_Antetype(
	Val           Out t_Refcur,
	原型id_In 影像报告动作.原型id%Type
	);

  --34. 根据字典名称获取相应子项
  Procedure p_Get_Dictitems_By_Title(
	Val           Out t_Refcur,
	名称_In 影像字典清单.名称%Type
	);
  --35.获得所有的预备提纲
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	);
  --36.获取所有词句信息
  Procedure p_Get_All_Fragment(
	Val           Out t_Refcur,
	学科_In Varchar2
	);

  --37.获取词句信息
  Procedure p_Get_Fragment_Filter(
	Val           Out t_Refcur,
	原型id_In 影像报告原型片段.原型ID%Type,
    作者_In   影像报告片段清单.作者%Type,
    学科_In   影像报告片段清单.学科%Type,
    Type_In   Varchar2
	);
  --38.根据原型获取关联的片段标签值
  Procedure p_Get_Label_By_Aid(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	);
  --39.获取所有词句分类
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	);
  --40.获取表名对应的最后编辑时间
  Procedure p_Get_Data_Last_Edit_Time(
	Val           Out t_Refcur,
	Table_Name_In Varchar2
	);
  --41.添加文档事件
  Procedure p_Add_Doc_Event(
    ID_In       影像报告事件.ID%Type,
    种类_In     影像报告事件.种类%Type,
    原型ID_In   影像报告事件.原型ID%Type,
    编号_In     影像报告事件.编号%Type,
    名称_In     影像报告事件.名称%Type,
    说明_In     影像报告事件.说明%Type,
    元素IID_In  影像报告事件.元素IID%Type,
    扩展标记_In 影像报告事件.扩展标记%Type);
  --42.修改文档事件
  Procedure p_Update_Doc_Event(
    Id_In       影像报告事件.Id%Type,
    种类_In     影像报告事件.种类%Type,
    名称_In     影像报告事件.名称%Type,
    说明_In     影像报告事件.说明%Type,
    元素IID_In  影像报告事件.元素IID%Type,
    扩展标记_In 影像报告事件.扩展标记%Type);
  --43.删除文档事件
  Procedure p_Delete_Doc_Event(
    Id_In 影像报告事件.Id%Type
	);
  --44.删除所有未被使用的文档事件
  Procedure p_Delete_Unused_Doc_Events(
    Count_Out Out Number
	);
  --45.获取指定原型的文档事件
  Procedure p_Get_Doc_Event_Of_Antetype(
	Val           Out t_Refcur,
	原型ID_In       影像报告事件.原型ID%Type,
	Include_Base_In Number
	);
  --46.修改文档处理编号
  Procedure p_Update_Doc_Process_Seqnum(
    Id_In   影像报告动作.Id%Type,
	序号_In 影像报告动作.序号%Type
	);
  --47.添加文档处理
  Procedure p_Add_Doc_Process(
    Id_In           影像报告动作.Id%Type,
    原型ID_In       影像报告动作.原型ID%Type,
    事件ID_In       影像报告动作.事件ID%Type,
    动作类型_In     影像报告动作.动作类型%Type,
    名称_In         影像报告动作.名称%Type,
    说明_In         影像报告动作.说明%Type,
    可否手工执行_In 影像报告动作.可否手工执行%Type,
    序号_In         影像报告动作.序号%Type,
    内容_In         影像报告动作.内容%Type
	);
  --48.修改文档处理
  Procedure p_Update_Doc_Process(
    Id_In           影像报告动作.Id%Type,
    事件ID_In       影像报告动作.事件ID%Type,
    动作类型_In     影像报告动作.动作类型%Type,
    名称_In         影像报告动作.名称%Type,
    说明_In         影像报告动作.说明%Type,
    可否手工执行_In 影像报告动作.可否手工执行%Type,
    内容_In         影像报告动作.内容%Type
	);
  --49.获取元素或者提纲的名称集合
  Procedure p_Get_Antetype_Ele_Section(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型清单.Id%Type,
	Type_In   Varchar2
	);
  --50.删除文档处理
  Procedure p_Del_Doc_Process(
    Id_In        影像报告动作.ID%Type,
	Del_Event_In Number
	);

  --51.查询元素值域类别的覆盖情况
  Procedure p_Get_Ele_Same_Info(
	Val           Out t_Refcur,
	Id_In    影像报告值域清单.Id%Type,
	Code_In  Varchar2,
	Title_In Varchar2,
	Flag_In  Varchar2
	);
  --52.获得所有的插件信息
  Procedure p_Get_DocPluginList(
    Val Out t_Refcur
	);
  --53.该ID的插件是否被原型使用过
  Procedure p_IsExit_DocPluginByID(
	Val           Out t_Refcur,
	ID_In Varchar2
	);
  --54.新增报告插件信息
  Procedure p_AddDocPlugin(
    ID_In       影像报告插件.ID%Type,
    编码_In     影像报告插件.编码%Type,
    名称_In     影像报告插件.名称%Type,
    说明_In     影像报告插件.说明%Type,
    显示样式_In 影像报告插件.显示样式%Type,
    种类_In     影像报告插件.种类%Type,
    类名_In     影像报告插件.类名%Type,
    库名_In     影像报告插件.库名%Type,
    是否禁用_In 影像报告插件.是否禁用%Type
	);
  --55.修改报告插件信息
  Procedure p_EditDocPlugin(
    ID_In       影像报告插件.ID%Type,
    编码_In     影像报告插件.编码%Type,
    名称_In     影像报告插件.名称%Type,
    说明_In     影像报告插件.说明%Type,
    显示样式_In 影像报告插件.显示样式%Type,
    种类_In     影像报告插件.种类%Type,
    类名_In     影像报告插件.类名%Type,
    库名_In     影像报告插件.库名%Type,
    是否禁用_In 影像报告插件.是否禁用%Type
	);
  --56.删除报告插件信息
  Procedure p_DelDocPlugin(
    ID_In 影像报告插件.ID%Type
	);
  --57.改变插件的可用状态
  Procedure p_IsEnableDocPlugin(
    ID_In 影像报告插件.ID%Type
	);
  --58.通过ID获得对应的插件信息
  Procedure p_GetDocPluginByID(
	Val           Out t_Refcur,
	ID_In 影像报告插件.ID%type
	);
  --59.判断编码和名称是否已存在
  Procedure p_IsExitDocPlugin(
	Val           Out t_Refcur,
	ID_In   影像报告插件.ID%Type,
    编码_In 影像报告插件.编码%Type,
    名称_In 影像报告插件.名称%Type
	);
  --60.通过ID获得对应的专用插件信息
  Procedure p_GetDocSpecPluginByID(
	Val           Out t_Refcur,
	ID_In 影像报告插件.ID%Type
	);
  --61.获得诊疗列表信息
  Procedure p_GetDiagnosisList(
	Val           Out t_Refcur,
	类别_In Varchar2,
    条件_In Varchar2
	);
  --62.获得诊疗类别列表
  Procedure p_GetDiagnosisClass(
    Val Out t_Refcur
	);
  --63.添加影像报告原型应用信息
  Procedure p_AddMedicalAntetype(
    诊疗项目ID_In 影像报告原型应用.诊疗项目ID%Type,
	应用场合_In   影像报告原型应用.应用场合%Type,
	报告原型ID_In 影像报告原型应用.报告原型ID%Type
	);
  --64.删除原型ID对应的病历单据应用信息
  Procedure p_DelMedicalAntetype(
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	);
  --65.通过原型ID获得对应的病历单据应用信息
  Procedure p_GetMedicalByAID(
	Val           Out t_Refcur,
	报告原型ID_In 影像报告原型应用.报告原型ID%Type
	);
  --66.根据原型ID删除动作信息
  Procedure p_DelDocProcessByAid(
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	);
  --67.获取ID对应的原型的树形结构
  Procedure p_GetAntetypeTreeByID(
	Val           Out t_Refcur,
	ID_In 影像报告原型清单.ID%Type
	);
  --68.原型是否存在对应的编码或名称
  procedure p_IsExitAntetype(
	Val           Out t_Refcur,
	编码_In 影像报告原型清单.编码%Type,
    名称_In 影像报告原型清单.名称%Type,
    ID_In  影像报告原型清单.ID%Type
	);

  --69  获取影像存储设备
  Procedure p_GetStorageDevice(
		Val           Out t_Refcur);

End b_PACS_RptAntetype;
/

--影像报告原型管理(---实现部分---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptAntetype Is
  --Create By Hwei;
  --2014/11/25

  --1.获取文件原型类别
  Procedure p_Get_Antetypelistkind(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select 编码, a.名称, a.编码 || '-' || a.名称 As 标题
        From 影像报告种类 A
       Order By 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelistkind;

  --2.根据文档类型获取文档信息
  Procedure p_Get_Antetypelis_By_Kind(
	Val           Out t_Refcur,
	种类_In      影像报告原型清单.种类%Type,
    Stop_Flag    Number,
    Condition_In Varchar2
	) As
  Begin
    Open Val For
      Select ID, 编码, 名称, 标题, 分组, 是否禁用, 说明, Imageindex
        From (Select Distinct 分组 As ID,
                              (Select Min(b.编码)
                                 From 影像报告原型清单 B
                                Where b.分组 = a.分组) As 编码,
                              a.分组 As 名称,
                              a.分组 As 标题,
                              null As 分组,
                              0 As 是否禁用,
                              null As 说明,
                              0 As Imageindex
                From 影像报告原型清单 A
               Where a.种类 = 种类_In
                 And ((a.是否禁用 <> 1 And Stop_Flag = 1) Or (Stop_Flag = 0))
                 And ((a.名称 Like '%' || Condition_In || '%' And
                     Condition_In Is Not Null) Or Condition_In Is Null)
                 And a.分组 Is Not Null
              Union
              Select RawtoHex(ID) ID,
                     a.编码,
                     名称 As 名称,
                     编码 || '-' || 名称 As 标题,
                     分组,
                     a.是否禁用,
                     a.说明,
                     Decode(a.是否禁用, 1, 2, 1) Imageindex
                From 影像报告原型清单 A
               Where a.
               种类 = 种类_In
                 And ((a.是否禁用 <> 1 And Stop_Flag = 1) Or (Stop_Flag = 0))
                 And ((a.名称 Like '%' || Condition_In || '%' And
                     Condition_In Is Not Null) Or Condition_In Is Null)) A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelis_By_Kind;

  --3.添加一个文档原型
  Procedure p_Add_Antetypelist(
    ID_In           影像报告原型清单.ID%Type,
    种类_In         影像报告原型清单.种类%Type,
    编码_In         影像报告原型清单.编码%Type,
    名称_In         影像报告原型清单.名称%Type,
	设备号_In		影像设备目录.设备号%Type,
    说明_In         影像报告原型清单.说明%Type,
    可否重置页面_In 影像报告原型清单.可否重置页面%Type,
    可否重置格式_In 影像报告原型清单.可否重置格式%Type,
	可否书写多份_In 影像报告原型清单.可否书写多份%Type,
    是否禁用_In     影像报告原型清单.是否禁用%Type,
    创建人_In       影像报告原型清单.创建人%Type,
    内容_In         影像报告原型清单.内容%Type,
    控制选项_In     影像报告原型清单.控制选项%Type,
    专用插件_In     影像报告原型清单.专用插件%Type,
    Copy_ID_In      影像报告原型清单.ID%Type,
    Only_Head_In    Varchar2,
    分组_In         影像报告原型清单.分组%Type
	) As
    x_Str Xmltype;
  Begin
    Begin
      If Copy_ID_In Is Null or Copy_ID_In = 0 Then
        x_Str := 内容_In;
      Else
        Select Decode(Only_Head_In,
                      1,
                      Deletexml(a.内容, '/zlxml/document/node()'),
                      a.内容)
          Into x_Str
          From 影像报告原型清单 A
         Where a.id = Copy_ID_In;
      End If;
    Exception
      When Others Then
        x_Str := 内容_In;
    End;
  
    Insert Into 影像报告原型清单
      (ID,
       种类,
       编码,
       名称,
	   设备号,
       说明,
       可否重置页面,
       可否重置格式,
	   可否书写多份,
       是否禁用,
       创建人,
       创建时间,
       内容,
       控制选项,
       专用插件,
       分组)
    Values
      (ID_In,
       种类_In,
       编码_In,
       名称_In,
	   设备号_In,
       说明_In,
       可否重置页面_In,
       可否重置格式_In,
	   可否书写多份_In,
       是否禁用_In,
       创建人_In,
       sysdate,
       x_Str,
       控制选项_In,
       专用插件_In,
       分组_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Antetypelist;

  --4.修改一个文档原型
  Procedure p_Edit_Antetypelist(
    ID_In           影像报告原型清单.ID%Type,
    种类_In         影像报告原型清单.种类%Type,
    编码_In         影像报告原型清单.编码%Type,
    名称_In         影像报告原型清单.名称%Type,
	设备号_In		影像设备目录.设备号%Type,
    说明_In         影像报告原型清单.说明%Type,
    可否重置页面_In 影像报告原型清单.可否重置页面%Type,
    可否重置格式_In 影像报告原型清单.可否重置格式%Type,
	可否书写多份_In 影像报告原型清单.可否书写多份%Type,
    是否禁用_In     影像报告原型清单.是否禁用%Type,
    修改人_In       影像报告原型清单.修改人%Type,
    内容_In         影像报告原型清单.内容%Type,
    控制选项_In     影像报告原型清单.控制选项%Type,
    专用插件_In     影像报告原型清单.专用插件%Type,
    Copy_ID_In      影像报告原型清单.ID%Type,
    Only_Head_In    Varchar2,
    分组_In         影像报告原型清单.分组%Type
	) As
    x_Str     Xmltype;
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(a.Id)
      Into n_Num
      From 影像报告原型清单 A
     Where (a.编码 = 编码_In Or a.名称 = 名称_In)
       And ID <> ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]存在相同的文档编码或者名称，请重新填写！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If Copy_ID_In Is Null or Copy_ID_In = 0 Then
      x_Str := 内容_In;
    Else
      Select Decode(Only_Head_In,
                    1,
                    Deletexml(a.内容, '/zlxml/document/node()'),
                    a.内容)
        Into x_Str
        From 影像报告原型清单 A
       Where a.id = Copy_ID_In;
    End If;
  
    Update 影像报告原型清单
       Set 种类         = 种类_In,
           编码         = 编码_In,
           名称         = 名称_In,
		   设备号		= 设备号_In,
           说明         = 说明_In,
           可否重置页面 = 可否重置页面_In,
           可否重置格式 = 可否重置格式_In,
		   可否书写多份 = 可否书写多份_In,
           是否禁用     = NVL(是否禁用_In, 是否禁用),
           修改人       = 修改人_In,
           修改时间     = sysdate,
           内容         = x_Str,
           控制选项     = 控制选项_In,
           专用插件     = 专用插件_In,
           分组         = 分组_In
     Where ID = ID_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Antetypelist;

  --5.删除一个文件原型
  Procedure p_Del_Antetypelist(
    ID_In 影像报告原型清单.Id%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(ID) Into n_Num From 影像报告记录 A Where a.原型id = ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]该原型已经被文档使用，不允许删除！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Count(原型ID)
      Into n_Num
      From 影像报告原型片段
     Where 影像报告原型片段.原型ID = ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]该文档下存在词句关联，不允许删除！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Count(ID)
      Into n_Num
      From 影像报告范文清单
     Where 影像报告范文清单.原型ID = ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]存在以此原型建立的范文信息，不允许删除！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Delete From 影像报告原型清单 C Where c.Id = ID_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Antetypelist;

  --6.根据ID获取文件原型
  Procedure p_Get_Antetypelist_By_Id(
	Val           Out t_Refcur,
	ID_In 影像报告原型清单.Id%Type
	) As
  Begin
    Open Val For
      Select rawtohex(a.ID) ID,
             a.种类,
             a.编码,
             a.名称,
			 a.设备号,
             a.说明,
             a.可否重置页面,
             a.可否重置格式,
			 a.可否书写多份,
             Extractvalue(b.Column_Value, '/root/print_hf_mode') Printhfmode,
             Extractvalue(b.Column_Value, '/root/print_follow_pages') Printfollowpages,
             Nvl(Extractvalue(b.Column_Value, '/root/print_limit'), 0) Printlimit,
             Nvl(a.控制选项.GetClobVal(), '<NULL/>') as 控制选项,
             a.是否禁用,
             Nvl(a.专用插件.GetClobVal(), '<NULL/>') as 专用插件,
             a.创建人,
             a.创建时间,
             a.修改人,
             a.修改时间,
             a.分组
        From 影像报告原型清单 A,
             Table(Xmlsequence(Extract(a.控制选项, '/root'))) B
       Where a.Id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelist_By_Id;

  --7.获取原型XML内容
  Procedure p_Get_Antetypelist_Content(
	Val           Out t_Refcur,
	ID_In 影像报告原型清单.Id%Type
	) As
  Begin
    Open Val For
      Select Nvl(a.内容.GetClobVal(), '<ZLXML/>') As 内容
        From 影像报告原型清单 A
       Where a.Id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelist_Content;

  --8.停用或启用文件原型
  Procedure p_Stop_Antetypelist(
    ID_In 影像报告原型清单.Id%Type
	) As
  Begin
    Update 影像报告原型清单
       Set 是否禁用 = Decode(是否禁用, 1, 0, 0, 1)
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Stop_Antetypelist;

  --9.新增文档种类信息
  Procedure p_Add_Doc_Kind(
    编码_In 影像报告种类.编码%Type,
    名称_In 影像报告种类.名称%Type,
    说明_In 影像报告种类.说明%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(a.编码)
      Into n_Num
      From 影像报告种类 A
     Where a.编码 = 编码_In
        Or a.名称 = 名称_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]种类的编码或者名称不能相同！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If 编码_In Is Null Or 编码_In Is Null Then
      v_Err_Msg := '[ZLSOFT]种类的编码或者名称不能为空！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Insert Into 影像报告种类
      (编码, 名称, 说明)
    Values
      (编码_In, 名称_In, 说明_In);
  
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Doc_Kind;

  --10.删除文档种类信息
  Procedure p_Del_Doc_Kind As
  Begin
    Delete From 影像报告种类;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Doc_Kind;

  --11.获取预备提纲信息
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawtoHex(ID) ID, 编码, 名称, 说明, 最后编辑时间
        From 影像报告预备提纲 A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Pre_Outline;

  --12.添加预备提纲信息
  Procedure p_Add_Pre_Outline(
    ID_In   影像报告预备提纲.ID%Type,
    编码_In 影像报告预备提纲.编码%Type,
    名称_In 影像报告预备提纲.名称%Type,
    说明_In 影像报告预备提纲.说明%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(a.编码)
      Into n_Num
      From 影像报告预备提纲 A
     Where a.编码 = 编码_In
        Or a.名称 = 名称_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]种类的编码或者名称不能相同！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If 编码_In Is Null Or 名称_In Is Null Then
      v_Err_Msg := '[ZLSOFT]种类的编码或者名称不能为空！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Insert Into 影像报告预备提纲
      (ID, 编码, 名称, 说明, 最后编辑时间)
    Values
      (ID_In, 编码_In, 名称_In, 说明_In, sysdate);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Pre_Outline;

  --13.删除预备提纲信息
  Procedure p_Del_Pre_Outline As
  Begin
    Delete From 影像报告预备提纲;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Pre_Outline;

  --14.获取导出的文档原型信息
  Procedure p_Output_Antetypelist(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select '类别' As 类别,
             b.编码 As ID,
             Null As 种类,
             b.名称 As 种类名称,
             b.编码 As 编码,
             b.名称 As 名称,
             b.说明 As 说明,
             Null As 可否重置页面,
             Null As 可否重置格式,
             Null As 是否禁用,
             Null As 创建人,
             Null As 创建时间,
             Null As 修改人,
             Null As 修改时间,
             Null As 内容
        From 影像报告种类 B
      Union All
      Select '原型' 类别,
             RawToHex(a.Id) ID,
             a.种类,
             b.名称 种类名称,
             a.编码,
             a.名称,
             a.说明,
             a.可否重置页面,
             a.可否重置格式,
             a.是否禁用,
             a.创建人,
             a.创建时间,
             a.修改人,
             a.修改时间,
             Null As 内容
        From 影像报告原型清单 A, 影像报告种类 B
       Where a.种类 = b.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Output_Antetypelist;

  --15.添加原型片段
  Procedure p_Add_Antetype_Fragments(
    原型ID_In 影像报告原型片段.原型ID%Type,
	片段ID_In 影像报告原型片段.片段ID%Type) As
  Begin
    Insert Into 影像报告原型片段
      (原型ID, 片段ID)
    Values
      (原型ID_In, 片段ID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Antetype_Fragments;

  --16.删除原型片段
  Procedure p_Del_Antetype_Fragments(
    原型ID_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Delete From 影像报告原型片段 Where 原型ID = 原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Antetype_Fragments;

  --17.获取原型片段
  Procedure p_Get_Antetype_Fragments(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(片段ID) 片段ID
        From 影像报告原型片段 A
       Where a.原型ID = 原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Fragments;

  --18.获取某个原型关联的某个片段分类
  Procedure p_Get_Antetype_f_Byaidfid(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type,
	片段ID_In 影像报告原型片段.片段ID%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(片段ID) 片段ID
        From 影像报告原型片段 A
       Where a.原型ID = 原型ID_In
         And a.片段ID = 片段ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_f_Byaidfid;

  --19.插入文档原型XML内容
  Procedure p_Edit_Antetypelist_Content(
    ID_In     影像报告原型清单.Id%Type,
	内容_In   影像报告原型清单.内容%Type,
	修改人_In 影像报告原型清单.修改人%Type
	) As
  Begin
    Update 影像报告原型清单
       Set 内容 = 内容_In, 修改人 = 修改人_In, 修改时间 = Sysdate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Antetypelist_Content;

  --20.获取所有原型
  Procedure p_Get_All_Antetype_Lists(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawtoHex(ID) ID,
             a.编码,
             编码 || '-' || 名称 As 名称,
             分组,
             a.种类,
             a.是否禁用,
             a.说明,
             Decode(a.是否禁用, 1, 2, 1) Imageindex,
             Nvl(a.内容.GetClobVal(), '<ZLXML/>') As 内容
        From 影像报告原型清单 A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Antetype_Lists;

  --21.获取已经设置了关联的原型片段类别的信息
  Procedure p_Get_Antetype_Fragments_Info(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(ID) ID,
             a.编码,
             a.名称,
             a.编码 || '-' || a.名称 标题,
             a.说明
        From 影像报告片段清单 A
       Where a.Id In (Select b.片段id
                        From 影像报告原型片段 B
                       Where b.原型id = 原型ID_In)
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Fragments_Info;

  --22.获取选择的类别下面的短语名称
  Procedure p_Get_Selected_Fragments(
	Val           Out t_Refcur,
	原型ID_In Varchar2
	) As
    v_Sql  Varchar2(4000);
    v_Aids Varchar2(4000);
    v_Msg  Varchar2(4000);
    Err Exception;
  Begin
    For Myrow In (Select RawtoHex(a.片段id) ID
                    From 影像报告原型片段 A
                   Where a.原型id = 原型ID_In) Loop
      If v_Aids Is Null Then
        v_Aids := '''' || Myrow.Id || '''';
      Else
        v_Aids := v_Aids || ',''' || Myrow.Id || '''';
      End If;
    End Loop;
  
    If v_Aids Is Null Then
      If Substr(原型ID_In, 0, 1) <> '''' Then
        v_Aids := '''' || 原型ID_In || '''';
      Else
        v_Aids := 原型ID_In;
      End If;
    End If;
  
    v_Sql := 'Select Distinct  RawtoHex(a.id) ID,  RawtoHex(a.上级ID) 上级ID , a.编码, a.编码 || ''-'' || a.名称 标题,Decode(a.节点类型, 0, 0, 1) 节点类型
      From 影像报告片段清单 A
      Start With a.Id In (' || v_Aids || ')
      Connect By Prior a.Id = a.上级ID
      Order By a.编码';
  
    Open Val For v_Sql;
  Exception
    When Err Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Selected_Fragments;

  --23.获取能复制的原型名称
  Procedure p_Get_Copy_Antetype(
	Val           Out t_Refcur,
	种类_In 影像报告原型清单.种类%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(a.id) ID, a.编码 || '-' || a.名称 标题
        From 影像报告原型清单 A
       Where a.种类 = 种类_In
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Copy_Antetype;

  --24.获取原型的分组信息
  Procedure p_Get_Antetype_Category(
	Val           Out t_Refcur,
	种类_In 影像报告原型清单.种类%Type
	) As
  Begin
    Open Val For
      Select Distinct a.分组 As 分组
        From 影像报告原型清单 A
       Where a.种类 = 种类_In
         and a.分组 Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Category;

  --25.根据原型同步范文提纲
  Procedure p_Synchronous_Sample(
    原型ID_In 影像报告原型清单.Id%Type
	) As
    x_Content Xmltype;
    x_Result  Xmltype;
    Cursor c_Antetype Is
      Select Extractvalue(c.Column_Value, '/section/@iid') Iid,
             Extractvalue(c.Column_Value, '/section/@title') Title,
             c.Column_Value As Content
        From 影像报告原型清单 A,
             Table(Xmlsequence(Extract(a.内容, '/zlxml//section'))) C
       Where a.Id = 原型ID_In;
    n_i               Number;
    n_j               Number;
    n_Count           Number;
    x_Subdocuments    Xmltype;
    x_Docparameters   Xmltype;
    x_Antetypecontent Xmltype;
    v_Textstyleno     Varchar2(10);
    v_Parastyleno     Varchar2(10);
    x_Acontent        Xmltype;
  Begin
    For Mysample In (Select b.id, b.内容
                       From 影像报告范文清单 B
                      Where b.原型id = 原型ID_In) Loop
      x_Content := Mysample.内容;
      n_i       := 1;
      If x_Content Is Null Then
        Select a.内容
          Into x_Result
          From 影像报告原型清单 A
         Where a.Id = 原型ID_In;
      Else
        Begin
          Select Extractvalue(c.Column_Value, '/section/@textstyleno') Textstyleno,
                 Extractvalue(c.Column_Value, '/section/@parastyleno') Parastyleno
            Into v_Textstyleno, v_Parastyleno
            From Table(Xmlsequence(Extract(x_Content, '/zlxml//section'))) C
           Where Rownum = 1;
        Exception
          When Others Then
            v_Textstyleno := '1';
            v_Parastyleno := '1';
        End;
      
        For Myantetype In c_Antetype Loop
          For I In 1 .. 1 Loop
            If n_i <> 1 Or n_Count <> 0 Or n_Count Is Null Then
              Select Count(*)
                Into n_Count
                From Table(Xmlsequence(Extract(x_Content, '/zlxml//section'))) C;
            End If;
            If n_Count < n_i Then
              Select Updatexml(Myantetype.Content,
                               '//section/@textstyleno',
                               v_Textstyleno)
                Into x_Acontent
                From Dual;
              Select Updatexml(x_Acontent,
                               '//section/@parastyleno',
                               v_Parastyleno)
                Into x_Acontent
                From Dual;
              Select Appendchildxml(x_Content,
                                    '/zlxml/document',
                                    x_Acontent)
                Into x_Content
                From Dual;
              Exit;
            End If;
            n_j := 1;
            For Mysample In (Select Extractvalue(c.Column_Value,
                                                 '/section/@iid') Iid,
                                    Extractvalue(c.Column_Value,
                                                 '/section/@title') Title
                               From Table(Xmlsequence(Extract(x_Content,
                                                              '/zlxml//section'))) C) Loop
              If n_i = n_j Then
                If Myantetype.Iid <> Mysample.Iid Then
                  Select Updatexml(Myantetype.Content,
                                   '//section/@textstyleno',
                                   v_Textstyleno)
                    Into x_Acontent
                    From Dual;
                  Select Updatexml(x_Acontent,
                                   '//section/@parastyleno',
                                   v_Parastyleno)
                    Into x_Acontent
                    From Dual;
                  Select Deletexml(x_Content,
                                   '//section[@iid="' || Myantetype.Iid || '"]')
                    Into x_Content
                    From Dual;
                  Select Insertxmlbefore(x_Content,
                                         '//section[@iid="' || Mysample.Iid || '"]',
                                         x_Acontent)
                    Into x_Content
                    From Dual;
                  n_j := n_j + 1;
                  Exit;
                Else
                  n_j := n_j + 1;
                  Exit;
                End If;
              End If;
              n_j := n_j + 1;
            End Loop;
            n_i := n_i + 1;
          End Loop;
        End Loop;
        x_Result := x_Content;
        For Mysample2 In (Select Iid
                            From (Select Extractvalue(c.Column_Value,
                                                      '/section/@iid') Iid
                                    From Table(Xmlsequence(Extract(x_Content,
                                                                   '/zlxml//section'))) C) C
                           Where c.Iid Not In
                                 (Select Extractvalue(c.Column_Value,
                                                      '/section/@iid') Iid
                                    From 影像报告原型清单 A,
                                         Table(Xmlsequence(Extract(a.内容,
                                                                   '/zlxml//section'))) C
                                   Where a.Id = 原型ID_In)) Loop
          Select Deletexml(x_Result,
                           '//section[@iid="' || Mysample2.Iid || '"]')
            Into x_Result
            From Dual;
        End Loop;
      End If;
    
      Update 影像报告范文清单 X
         Set x.内容 = x_Result
       Where x.Id = Mysample.Id;
    End Loop;
  
    Select a.内容
      Into x_Antetypecontent
      From 影像报告原型清单 A
     Where a.Id = 原型ID_In;
    Select Extract(x_Antetypecontent, 'zlxml/subdocuments')
      Into x_Subdocuments
      From Dual;
    Select Extract(x_Antetypecontent, 'zlxml/docparameters')
      Into x_Docparameters
      From Dual;
    Update 影像报告范文清单
       Set 内容 = Updatexml(内容, '/zlxml/subdocuments', x_Subdocuments)
     Where 原型ID = 原型ID_In;
    Update 影像报告范文清单
       Set 内容 = Updatexml(内容, '/zlxml/docparameters', x_Docparameters)
     Where 原型ID = 原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Synchronous_Sample;

  --26.获取导出的原型列表
  Procedure p_Get_Out_Antetypelist(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select ID,
             编码,
             标题,
             Parentid,
             种类,
             是否禁用,
             说明,
             Imageindex,
             名称
        From (Select a.编码 As ID,
                     a.编码 As 编码,
                     a.名称 As 标题,
                     '' As Parentid,
                     '-1' As 种类,
                     0 As 是否禁用,
                     a.说明 As 说明,
                     4 As Imageindex,
                     a.名称 名称
                From 影像报告种类 A
              Union
              Select Distinct a.种类 || '-' || a.分组 As ID,
                              (Select Min(编码)
                                 From 影像报告原型清单 B
                                Where b.分组 = a.分组) As 编码,
                              Max(a.分组) As 名称,
                              a.种类 As Parentid,
                              '0' As 种类,
                              0 As 是否禁用,
                              '' As 说明,
                              4 As Imageindex,
                              a.分组
                From 影像报告原型清单 A
               Where a.分组 Is Not Null
               Group By a.种类, a.分组
              Union
              Select RawTohex(ID),
                     a.编码,
                     编码 || '-' || 名称 As 标题,
                     Decode(a.分组, Null, a.种类, a.种类 || '-' || a.分组) Parentid,
                     a.种类 As 种类,
                     a.是否禁用,
                     a.说明,
                     Decode(a.是否禁用, 1, 1, 0, 2),
                     a.名称
                From 影像报告原型清单 A) A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Out_Antetypelist;

  --27.通过编码获取原型种类信息
  Procedure p_Get_Antetype_Kind_By_Code(
	Val           Out t_Refcur,
	编码_In 影像报告种类.编码%Type
	) As
  Begin
    Open Val For
      Select a.编码, a.名称, a.说明
        From 影像报告种类 A
       Where a.编码 = 编码_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Kind_By_Code;
  --28.获取事件信息，不包含固定事件
  Procedure p_Get_Doc_Event(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawtoHex(a.id) ID,
             a.种类,
             a.原型id,
             a.编号,
             a.名称,
             a.说明,
             a.元素iid,
             a.扩展标记
        From 影像报告事件 A
       Where a.种类 <> 1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Event;

  --29.获取关于原型导出的重复信息
  Procedure p_Get_Antetypelist_Same_Info(
	Val           Out t_Refcur,
	Tablename_In Varchar2,
	ID_In        影像报告原型清单.Id%Type,
	编码_In      Varchar2,
	名称_In      Varchar2
	) As
    n_Num    Number;
    v_Result Varchar2(100);
    v_Sql    Varchar2(4000);
  Begin
    If ID_In Is Not Null Then
      v_Sql := 'select count(*) from ' || Tablename_In || ' where id=' ||
               ID_In;
      Execute Immediate v_Sql
        Into n_Num;
      If n_Num > 0 Then
        v_Result := 'ID重复';
      End If;
    End If;
    If 编码_In Is Not Null Then
      v_Sql := 'select count(*) from ' || Tablename_In || ' where 编码=''' ||
               编码_In || '''';
      Execute Immediate v_Sql
        Into n_Num;
      If n_Num > 0 Then
        If v_Result Is Not Null Then
          v_Result := v_Result || ',编码重复';
        Else
          v_Result := '编码重复';
        End If;
      End If;
    End If;
    If 名称_In Is Not Null Then
      v_Sql := 'select count(*) from ' || Tablename_In || ' where 名称=''' ||
               名称_In || '''';
      Execute Immediate v_Sql
        Into n_Num;
      If n_Num > 0 Then
        If v_Result Is Not Null Then
          v_Result := v_Result || ',名称重复';
        Else
          v_Result := '名称重复';
        End If;
      End If;
    End If;
    Open Val For
      Select v_Result Result From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelist_Same_Info;

  --30.获取事件重复的信息
  Procedure p_Event_Same_Info(
	Val           Out t_Refcur,
	ID_In      影像报告事件.Id%Type,
    原型ID_In  影像报告事件.原型ID%Type,
    元素IID_In 影像报告事件.元素IID%Type,
    种类_In    影像报告事件.种类%Type,
    名称_In    影像报告事件.名称%Type,
    编号_In    影像报告事件.编号%Type
	) As
    v_Same_Antetype Varchar2(50);
    n_Same_Id       Number;
    n_Same_Title    Number;
    n_Same_Seqnum   Number;
    n_Maxnum        Number;
  Begin
    Select Count(*)
      Into n_Same_Title
      From 影像报告事件 A
     Where a.原型ID = 原型ID_In
       And a.种类 = 种类_In
       And a.名称 = 名称_In;
    Select Count(*)
      Into n_Same_Seqnum
      From 影像报告事件 A
     Where a.原型ID = 原型ID_In
       And a.种类 = 种类_In
       And a.编号 = 编号_In;
    Begin
      Select a.Id
        Into v_Same_Antetype
        From 影像报告事件 A
       Where a.原型ID = 原型ID_In
         And a.元素IID = 元素IID_In;
    Exception
      When Others Then
        v_Same_Antetype := '';
    End;
  
    Select Count(*) Into n_Same_Id From 影像报告事件 A Where a.Id = ID_In;
    Select Max(a.编号) Into n_Maxnum From 影像报告事件 A;
  
    Open Val For
      Select v_Same_Antetype As Sameaid,
             n_Same_Id       As Sameid,
             n_Same_Title    As Sametitle,
             n_Same_Seqnum   As Sameseqnum,
             n_Maxnum        As Maxnum
        From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Event_Same_Info;

  --31.获取原型校验的类别集合
  Procedure p_Get_Process_Kind(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Distinct 动作类型
        From (Select Extractvalue(c.Column_Value, '/step/kind') As 动作类型
                From 影像报告动作 A,
                     Table(Xmlsequence(Extract(a.内容, '/root/step'))) C) B
       Where b.动作类型 Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Process_Kind;


  --33.获取指定原型的文档处理
  Procedure p_Get_Doc_Process_Of_Antetype(
	Val           Out t_Refcur,
	原型ID_In 影像报告动作.原型id%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(p.id) ID,
             p.名称,
             p.动作类型,
             p.序号,
             p.说明,
             p.可否手工执行,
             To_Clob(Nvl(p.内容.GetClobVal(), '<NULL/>')) As 内容, --Nvl(p.内容,'<NULL/>') As 内容,
             RawtoHex(p.事件ID) 事件ID,
             0 Is_Event
        From 影像报告动作 P
       Where p.原型ID = 原型ID_In
      Union All
      Select RawtoHex(e.id) ID,
             e.名称,
             e.种类,
             e.编号,
             e.说明,
             Null,
             To_CLOB('<Null/>') As 内容, --(Null,'<NULL/>') As 内容,
             Null,
             1
        From 影像报告事件 E
       Where e.Id In (Select RawtoHex(事件ID) 事件ID
                        From 影像报告动作
                       Where 原型ID = 原型ID_In)
       Order By Is_Event, 动作类型, 序号;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Process_Of_Antetype;

  --34. 根据字典名称获取相应子项
  Procedure p_Get_Dictitems_By_Title(
	Val           Out t_Refcur,
	名称_In 影像字典清单.名称%Type
	) As
  Begin
    Open Val For
      Select a.编号, a.名称, Rawtohex(a.字典id) As 字典ID
        From 影像字典内容 A
       Where a.字典id In (Select id From 影像字典清单 b Where b.名称 = 名称_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Dictitems_By_Title;

  --35.获得所有的预备提纲
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(ID) ID, a.编码, a.名称
        From 影像报告预备提纲 a
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Phr_Onlines;

  --36.获取所有词句信息
  Procedure p_Get_All_Fragment(
	Val           Out t_Refcur,
	学科_In Varchar2
	) As
  Begin
    If 学科_In <> '' Then
      Open Val For
        Select RawToHex(a.id) ID,
               RawToHex(a.上级id) 上级id,
               a.编码,
               a.名称,
               a.说明,
               a.节点类型,
               Nvl(a.组成.GetClobVal(), '<NULL/>') As 组成,
               a.学科,
               a.标签,
               a.是否私有,
               a.作者
          From 影像报告片段清单 A
         Where (a.学科 In
               (Select /*+rule*/
                  Column_Value As Lable
                   From Table(b_PACS_RptPublic.f_Str2list(学科_In, ','))
                 Intersect
                 Select /*+rule*/
                  Column_Value As Lable
                   From Table(b_PACS_RptPublic.f_Str2list(a.学科, ','))) And
               a.节点类型 <> 0)
            Or a.节点类型 = 0
            Or a.学科 Is Null
         Order By a.编码, a.上级id;
    Else
      Open Val For
        Select RawToHex(a.id) ID,
               RawToHex(a.上级id) 上级id,
               a.编码,
               a.名称,
               a.说明,
               a.节点类型,
               Nvl(a.组成.GetClobVal(), '<NULL/>') As 组成,
               a.学科,
               a.标签,
               a.是否私有,
               a.作者
          From 影像报告片段清单 A
         Order By a.上级id, a.节点类型, a.编码, a.名称;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Fragment;

  --37. 获取词句信息
  Procedure p_Get_Fragment_Filter(
	Val           Out t_Refcur,
	原型id_In 影像报告原型片段.原型ID%Type,
    作者_In   影像报告片段清单.作者%Type,
    学科_In   影像报告片段清单.学科%Type,
    Type_In   Varchar2
	) As
  Begin
    If Type_In = '1' Then
      Open Val For
        Select Rawtohex(b.Id) ID,
               Rawtohex(b.上级id) 上级id,
               b.编码,
               b.名称,
               b.说明,
               b.节点类型,
               Nvl(b.组成.GetClobVal(), '<NULL/>') As 组成,
               b.学科,
               b.标签,
               b.是否私有,
               b.作者,
               b.最后编辑时间
          From 影像报告原型片段 A, 影像报告片段清单 B
         Where a.片段id = b.id
           And a.原型id = 原型id_In;
    Else
      Open Val For
        Select /*+ rule*/
         Rawtohex(b.Id) ID,
         Rawtohex(b.上级id) 上级id,
         b.编码,
         b.名称,
         b.说明,
         b.节点类型,
         Nvl(b.组成.GetClobVal(), '<NULL/>') As 组成,
         b.学科,
         b.标签,
         b.是否私有,
         b.作者,
         b.最后编辑时间
          From 影像报告片段清单 B
         Where b.上级id = 原型id_In
           And (b.是否私有 = 0 Or (b.是否私有 = 1 And b.作者 = 作者_In))
           And (b.学科 Is Null Or
               (b.学科 Is Not Null And
               b_PACS_RptPublic.f_If_Intersect(b.学科, 学科_In) > 0));
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Fragment_Filter;

  --38.根据原型获取关联的片段标签值
  Procedure p_Get_Label_By_Aid(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select Distinct b.标签
        From 影像报告片段清单 B
       Start With b.上级id In (Select a.片段id
                               From 影像报告原型片段 A
                              Where a.原型id = 原型ID_In)
      Connect By Prior b.Id = b.上级id;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Label_By_Aid;

  --39.获取所有词句分类
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) ID,
             Rawtohex(a.上级id) 上级id,
             a.编码,
             a.名称,
             a.说明,
             a.节点类型
        From 影像报告片段清单 A
       Where a.节点类型 = 0
       Start With 上级id Is Null
      Connect By Prior id = 上级id
       Order By 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Fragment_Class;

  --40.获取表名对应的最后编辑时间
  Procedure p_Get_Data_Last_Edit_Time(
	Val           Out t_Refcur,
	Table_Name_In Varchar2
	) As
    v_sql Varchar2(4000);
  Begin
    v_sql := 'select max(最后编辑时间) maxvalue from ' || Table_Name_In;
    Open val For v_sql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Data_Last_Edit_Time;

  --41.添加文档事件
  Procedure p_Add_Doc_Event(
    ID_In       影像报告事件.ID%Type,
    种类_In     影像报告事件.种类%Type,
    原型ID_In   影像报告事件.原型ID%Type,
    编号_In     影像报告事件.编号%Type,
    名称_In     影像报告事件.名称%Type,
    说明_In     影像报告事件.说明%Type,
    元素IID_In  影像报告事件.元素IID%Type,
    扩展标记_In 影像报告事件.扩展标记%Type
	) As
    n_Seq_Num  影像报告事件.编号%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(*)
      Into n_Is_Exist
      From 影像报告事件
     Where 原型ID = 原型ID_In
       And 种类 = 种类_In
       And 名称 = 名称_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]原型上已存在相同命名的事件[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If (编号_In Is Null Or 编号_In = 0) Then
      Select Nvl(Max(编号), 0) + 1 Into n_Seq_Num From 影像报告事件;
    Else
      Select Count(*)
        Into n_Is_Exist
        From 影像报告事件
       Where 原型ID = 原型ID_In
         And 种类 = 种类_In
         And 编号 = 编号_In;
      If n_Is_Exist > 0 Then
        v_Err_Msg := '[ZLSOFT]原型上已存在相同编号的事件[ZLSOFT]';
        Raise Err_Item;
      End If;
      n_Seq_Num := 编号_In;
    End If;
  
    Insert Into 影像报告事件
      (ID, 种类, 原型ID, 编号, 名称, 说明, 元素IID, 扩展标记)
    Values
      (ID_In,
       种类_In,
       原型ID_In,
       n_Seq_Num,
       名称_In,
       说明_In,
       元素IID_In,
       扩展标记_In);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Doc_Event;

  --42.修改文档事件
  Procedure p_Update_Doc_Event(
    Id_In       影像报告事件.Id%Type,
    种类_In     影像报告事件.种类%Type,
    名称_In     影像报告事件.名称%Type,
    说明_In     影像报告事件.说明%Type,
    元素IID_In  影像报告事件.元素IID%Type,
    扩展标记_In 影像报告事件.扩展标记%Type
	) As
    r_Aid      影像报告事件.原型ID%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select 原型ID Into r_Aid From 影像报告事件 Where ID = Id_In;
  
    Select Count(*)
      Into n_Is_Exist
      From 影像报告事件
     Where 原型ID = r_Aid
       And 种类 = 种类_In
       And 名称 = 名称_In
       And ID <> Id_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]原型上存在相同命名的事件[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Update 影像报告事件
       Set 种类     = 种类_In,
           名称     = 名称_In,
           说明     = 说明_In,
           元素IID  = 元素IID_In,
           扩展标记 = 扩展标记_In
     Where ID = Id_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Doc_Event;

  --43.删除文档事件
  Procedure p_Delete_Doc_Event(
    Id_In 影像报告事件.Id%Type
	) As
    n_Kind     影像报告事件.种类%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select 种类 Into n_Kind From 影像报告事件 Where ID = Id_In;
  
    If n_Kind = 1 Then
      v_Err_Msg := '[ZLSOFT]不允许删除固定事件[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Count(*) Into n_Is_Exist From 影像报告动作 Where 事件ID = Id_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]事件已经被使用,不能被删除！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Delete From 影像报告事件 Where ID = Id_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Delete_Doc_Event;

  --44.删除所有未被使用的文档事件
  Procedure p_Delete_Unused_Doc_Events(
    Count_Out Out Number
	) As
  Begin
    Delete From 影像报告事件
     Where 种类 <> 1
       And ID Not In
           (Select 事件ID From 影像报告动作 Where 事件ID Is Not Null);
    Count_Out := Sql%RowCount;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Delete_Unused_Doc_Events;

  --45.获取指定原型的文档事件
  Procedure p_Get_Doc_Event_Of_Antetype(
	Val           Out t_Refcur,
	原型ID_In       影像报告事件.原型ID%Type,
	Include_Base_In Number
	) As
  Begin
    If Include_Base_In = 1 Then
      Open Val For
        Select Rawtohex(t.Id) ID,
               t.种类,
               t.名称,
               t.说明,
               t.元素iid,
               t.扩展标记,
               Nvl(p.Used_Count, 0) Used_Count
          From 影像报告事件 T,
               (Select Count(*) Used_Count, Max(事件ID) 事件ID
                  From 影像报告动作
                 Where 事件ID Is Not Null
                 Group By 事件ID) P
         Where (t.种类 = 1 Or t.原型id = 原型ID_In)
           And t.Id = p.事件ID(+)
         Order By t.编号;
    Else
      Open Val For
        Select Rawtohex(t.Id) ID,
               t.种类,
               t.名称,
               t.说明,
               t.元素iid,
               t.扩展标记,
               Nvl(p.Used_Count, 0) Used_Count
          From 影像报告事件 T,
               (Select Count(*) Used_Count, Max(事件ID) 事件ID
                  From 影像报告动作
                 Where 事件ID Is Not Null
                 Group By 事件ID) P
         Where t.原型id = 原型ID_In
           And t.种类 <> 1
           And t.Id = p.事件ID(+)
         Order By t.编号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Event_Of_Antetype;

  --46.修改文档处理编号
  Procedure p_Update_Doc_Process_Seqnum(
    Id_In   影像报告动作.Id%Type,
	序号_In 影像报告动作.序号%Type) As
  Begin
    Update 影像报告动作 Set 序号 = 序号_In Where ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Doc_Process_Seqnum;

  --47.添加文档处理
  Procedure p_Add_Doc_Process(
    Id_In           影像报告动作.Id%Type,
    原型ID_In       影像报告动作.原型ID%Type,
    事件ID_In       影像报告动作.事件ID%Type,
    动作类型_In     影像报告动作.动作类型%Type,
    名称_In         影像报告动作.名称%Type,
    说明_In         影像报告动作.说明%Type,
    可否手工执行_In 影像报告动作.可否手工执行%Type,
    序号_In         影像报告动作.序号%Type,
    内容_In         影像报告动作.内容%Type
	) As
    n_Seq_Num  影像报告动作.序号%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(*)
      Into n_Is_Exist
      From 影像报告动作
     Where 原型ID = 原型ID_In
       And 名称 = 名称_In;
    If (序号_In Is Null Or 序号_In = 0) Then
      If (事件ID_In Is Null) Then
        Select Nvl(Max(序号), 0) + 1
          Into n_Seq_Num
          From 影像报告动作
         Where 原型ID = 原型ID_In
           And 事件ID Is Null;
      Else
        Select Nvl(Max(序号), 0) + 1
          Into n_Seq_Num
          From 影像报告动作
         Where 原型ID = 原型ID_In
           And 事件ID = 事件ID_In;
      End If;
    Else
      n_Seq_Num := 序号_In;
    End If;
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]原型上存在相同命名的动作[ZLSOFT]';
      Raise Err_Item;
    End If;
    Insert Into 影像报告动作
      (ID, 原型ID, 事件ID, 动作类型, 名称, 说明, 可否手工执行, 序号, 内容)
    Values
      (Id_In,
       原型ID_In,
       事件ID_In,
       动作类型_In,
       名称_In,
       说明_In,
       可否手工执行_In,
       n_Seq_Num,
       内容_In);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Doc_Process;

  --48.修改文档处理
  Procedure p_Update_Doc_Process(
    Id_In           影像报告动作.Id%Type,
    事件ID_In       影像报告动作.事件ID%Type,
    动作类型_In     影像报告动作.动作类型%Type,
    名称_In         影像报告动作.名称%Type,
    说明_In         影像报告动作.说明%Type,
    可否手工执行_In 影像报告动作.可否手工执行%Type,
    内容_In         影像报告动作.内容%Type
	) As
    r_Aid          影像报告事件.原型ID%Type;
    r_Old_Event_Id 影像报告动作.事件ID%Type;
    n_Seq_Num      影像报告事件.编号%Type;
    n_Is_Exist     Number(1) := 0;
    v_Err_Msg      Varchar2(100);
    Err_Item Exception;
  Begin
    Select 原型ID Into r_Aid From 影像报告动作 Where ID = Id_In;
    If (事件ID_In Is Not Null) Then
      Select Count(*)
        Into n_Is_Exist
        From 影像报告事件
       Where (原型ID Is Null Or 原型ID = r_Aid)
         And ID = 事件ID_In;
    
      If n_Is_Exist = 0 Then
        v_Err_Msg := '[ZLSOFT]关联的事件不存在[ZLSOFT]';
        Raise Err_Item;
      End If;
    
    End If;
  
    Select Count(*)
      Into n_Is_Exist
      From 影像报告动作
     Where 原型ID = r_Aid
       And 名称 = 名称_In
       And ID <> Id_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]原型上存在相同命名的动作[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If (r_Old_Event_Id <> 事件ID_In Or
       (事件ID_In Is Null And r_Old_Event_Id Is Not Null)) Then
      If (事件ID_In Is Null) Then
        Select Nvl(Max(序号), 0) + 1
          Into n_Seq_Num
          From 影像报告动作
         Where 原型ID = r_Aid
           And 事件ID Is Null;
      Else
        Select Nvl(Max(序号), 0) + 1
          Into n_Seq_Num
          From 影像报告动作
         Where 原型ID = r_Aid
           And 事件ID = 事件ID_In;
      End If;
    Else
      n_Seq_Num := 0;
    End If;
  
    If n_Seq_Num > 0 Then
      Update 影像报告动作
         Set 事件id       = 事件ID_In,
             动作类型     = 动作类型_In,
             名称         = 名称_In,
             说明         = 说明_In,
             可否手工执行 = 可否手工执行_In,
             内容         = 内容_In,
             序号         = n_Seq_Num
       Where ID = Id_In;
    Else
      Update 影像报告动作
         Set 事件id       = 事件ID_In,
             动作类型     = 动作类型_In,
             名称         = 名称_In,
             说明         = 说明_In,
             可否手工执行 = 可否手工执行_In,
             内容         = 内容_In
       Where ID = Id_In;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Doc_Process;

  --49.获取元素或者提纲的名称集合
  Procedure p_Get_Antetype_Ele_Section(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型清单.Id%Type,
	Type_In   Varchar2
	) As
    c_Content Clob;
  Begin
    /*Select To_Clob(a.内容)*/
    Select a.内容.getclobval()
      Into c_Content
      From 影像报告原型清单 A
     Where a.Id = 原型ID_In;
  
    If Type_In = '1' Then
      Open Val For
        Select Distinct Name
          From (Select Extractvalue(c.Column_Value, '/*/@title') As Name
                  From Table(Xmlsequence(Extract(Xmltype(c_Content),
                                                 '/zlxml/document//element[@sid and @title]|/zlxml/document//e_list[@sid and @title]|/zlxml/document//e_enum[@sid and @title]|/zlxml/document//e_etree[@sid and @title]|/zlxml/document//e_utree[@sid and @title]'))) C) A
         Where a.Name Is Not Null;
    Else
      Open Val For
        Select Distinct Name
          From (Select Extractvalue(c.Column_Value, '/section/@title') As Name
                  From Table(Xmlsequence(Extract(Xmltype(c_Content),
                                                 '//section'))) C) A
         Where a.Name Is Not Null;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Ele_Section;

  --50.删除文档处理
  Procedure p_Del_Doc_Process(Id_In        影像报告动作.ID%Type,
                              Del_Event_In Number) As
    r_Event_Id   影像报告动作.事件ID%Type := Null;
    n_Event_Kind 影像报告事件.种类%Type;
    n_Is_Exist   Number(1) := 0;
  Begin
    If Del_Event_In = 1 Then
      Select Max(e.Id), Max(e.种类)
        Into r_Event_Id, n_Event_Kind
        From 影像报告动作 P, 影像报告事件 E
       Where p.Id = Id_In
         And p.事件id = e.Id;
    End If;
  
    Delete From 影像报告动作 Where ID = Id_In;
  
    If Del_Event_In = 1 Then
      If (r_Event_Id Is Not Null And n_Event_Kind <> 1) Then
        Select Count(*)
          Into n_Is_Exist
          From 影像报告动作
         Where 事件id = r_Event_Id;
        If n_Is_Exist = 0 Then
          Delete From 影像报告事件
           Where ID = r_Event_Id
             And 种类 <> 1;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Doc_Process;

  --51.查询元素值域类别的覆盖情况
  Procedure p_Get_Ele_Same_Info(
	Val           Out t_Refcur,
	Id_In    影像报告值域清单.Id%Type,
	Code_In  Varchar2,
	Title_In Varchar2,
	Flag_In  Varchar2
	) As
    v_Result  Varchar2(50);
    v_Id      Varchar2(50);
    v_Code_Id Varchar2(50);
    n_Num     Number;
  Begin
    If Flag_In = 1 Then
      Select Count(ID)
        Into n_Num
        From 影像报告元素分类 A
       Where a.Id = Id_In;
    
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告元素分类 A
         Where a.Id = Id_In;
        If v_Id Is Not Null Then
          v_Result := 'ID重复';
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告元素分类 A
       Where a.编码 = Code_In;
    
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Code_Id
          From 影像报告元素分类 A
         Where a.编码 = Code_In;
        If v_Code_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',编码重复';
          Else
            v_Result := '编码重复';
          End If;
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告元素分类 A
       Where a.名称 = Title_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告元素分类 A
         Where a.名称 = Title_In;
        If v_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',名称重复';
          Else
            v_Result := '名称重复';
          End If;
        End If;
      End If;
    
    End If;
  
    If Flag_In = 2 Then
      Select Count(ID)
        Into n_Num
        From 影像报告元素清单 A
       Where a.Id = Id_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告元素清单 A
         Where a.Id = Id_In;
        If v_Id Is Not Null Then
          v_Result := 'ID重复';
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告元素清单 A
       Where a.编码 = Code_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Code_Id
          From 影像报告元素清单 A
         Where a.编码 = Code_In;
        If v_Code_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',编码重复';
          Else
            v_Result := '编码重复';
          End If;
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告元素清单 A
       Where a.名称 = Title_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告元素清单 A
         Where a.名称 = Title_In;
        If v_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',名称重复';
          Else
            v_Result := '名称重复';
          End If;
        End If;
      End If;
    
    End If;
  
    If Flag_In = 3 Then
      Select Count(ID)
        Into n_Num
        From 影像报告值域清单 A
       Where a.Id = Id_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告值域清单 A
         Where a.Id = Id_In;
        If v_Id Is Not Null Then
          v_Result := 'ID重复';
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告值域清单 A
       Where a.编码 = Code_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Code_Id
          From 影像报告值域清单 A
         Where a.编码 = Code_In;
        If v_Code_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',编码重复';
          Else
            v_Result := '编码重复';
          End If;
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告值域清单 A
       Where a.名称 = Title_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告值域清单 A
         Where a.名称 = Title_In;
        If v_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',名称重复';
          Else
            v_Result := '名称重复';
          End If;
        End If;
      End If;
    
    End If;
  
    Open Val For
      Select v_Result As Result, v_Id As ID, v_Code_Id As Codesameid
        From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Ele_Same_Info;

  --52.获得所有的插件信息
  Procedure p_Get_DocPluginList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(Id) ID,
             编码,
             名称,
             说明,
             显示样式,
             种类,
             Decode(显示样式, '1', '嵌入式', '弹出式') 显示样式II,
             Decode(种类, '1', '专用插件', '共享插件') 种类II,
             类名,
             库名,
             是否禁用,
             Decode(是否禁用, '1', '停用', '启用') IsEnable
        From 影像报告插件;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_DocPluginList;

  --53.该ID的插件是否被原型使用过
  Procedure p_IsExit_DocPluginByID(
    Val           Out t_Refcur,
	ID_In Varchar2
	) As
    CURSOR C_EVENT Is
      Select t.专用插件.getclobval() 专用插件 From 影像报告原型清单 t;
    anum Int := 0;
    sult Varchar2(6666);
  Begin
    For temp In C_EVENT Loop
      If instr(temp.专用插件, ID_In) > 0 Then
        anum := anum + 1;
      End If;
    End Loop;
    Open Val For
      Select anum From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_IsExit_DocPluginByID;

  --54.新增报告插件信息
  Procedure p_AddDocPlugin(
    ID_In       影像报告插件.ID%Type,
    编码_In     影像报告插件.编码%Type,
    名称_In     影像报告插件.名称%Type,
    说明_In     影像报告插件.说明%Type,
    显示样式_In 影像报告插件.显示样式%Type,
    种类_In     影像报告插件.种类%Type,
    类名_In     影像报告插件.类名%Type,
    库名_In     影像报告插件.库名%Type,
    是否禁用_In 影像报告插件.是否禁用%Type
	) As
  Begin
    Insert Into 影像报告插件
      (ID, 编码, 名称, 说明, 显示样式, 种类, 类名, 库名, 是否禁用)
    Values
      (ID_In,
       编码_In,
       名称_In,
       说明_In,
       显示样式_In,
       种类_In,
       类名_In,
       库名_In,
       是否禁用_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddDocPlugin;

  --55.修改报告插件信息
  Procedure p_EditDocPlugin(
    ID_In       影像报告插件.ID%Type,
    编码_In     影像报告插件.编码%Type,
    名称_In     影像报告插件.名称%Type,
    说明_In     影像报告插件.说明%Type,
    显示样式_In 影像报告插件.显示样式%Type,
    种类_In     影像报告插件.种类%Type,
    类名_In     影像报告插件.类名%Type,
    库名_In     影像报告插件.库名%Type,
    是否禁用_In 影像报告插件.是否禁用%Type
	) As
  Begin
    Update 影像报告插件
       Set 编码     = 编码_In,
           名称     = 名称_In,
           说明     = 说明_In,
           显示样式 = 显示样式_In,
           种类     = 种类_In,
           类名     = 类名_In,
           库名     = 库名_In,
           是否禁用 = 是否禁用_In
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditDocPlugin;

  --56.删除报告插件信息
  Procedure p_DelDocPlugin(
    ID_In 影像报告插件.ID%Type
	) As
  Begin
    Delete From 影像报告插件 Where id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelDocPlugin;

  --57.改变插件的可用状态
  Procedure p_IsEnableDocPlugin(
    ID_In 影像报告插件.ID%Type
	) As
  Begin
    Update 影像报告插件 a
       Set 是否禁用 = Decode(a.是否禁用, 1, 0, 1)
     Where id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_IsEnableDocPlugin;

  --58.通过ID获得对应的插件信息
  Procedure p_GetDocPluginByID(
	Val           Out t_Refcur,
	ID_In 影像报告插件.ID%type
	) As
  Begin
    Open Val For
      Select Rawtohex(Id) ID,
             编码,
             名称,
             说明,
             显示样式,
             种类,
             类名,
             库名,
             是否禁用
        From 影像报告插件
       Where id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDocPluginByID;

  --59.判断编码和名称是否已存在
  Procedure p_IsExitDocPlugin(
	Val           Out t_Refcur,
	ID_In   影像报告插件.ID%Type,
	编码_In 影像报告插件.编码%Type,
	名称_In 影像报告插件.名称%Type
	) As
  Begin
    Open Val For
      Select Count(id)
        From 影像报告插件 a
       Where (a.编码 = 编码_In Or a.名称 = 名称_In)
         and a.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_IsExitDocPlugin;

  --60.通过ID获得对应的专用插件信息
  Procedure p_GetDocSpecPluginByID(
	Val           Out t_Refcur,
	ID_In 影像报告插件.ID%Type
	) As
  Begin
    Open Val For
      Select Rawtohex(Id) ID,
             编码,
             名称,
             说明,
             显示样式,
             种类,
             类名,
             库名,
             是否禁用
        From 影像报告插件
       Where id = ID_In
         And 种类 = 1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDocSpecPluginByID;

  --61.获得诊疗列表信息
  Procedure p_GetDiagnosisList(
	Val           Out t_Refcur,
	类别_In Varchar2,
	条件_In Varchar2
	) As
  Begin
    Open Val For
      Select to_char(a.id) ID,
             a.编码,
             a.名称,
             (Select b.名称 From 诊疗项目类别 b Where b.编码 = a.类别) 类别
        From 诊疗项目目录 a
       Where (a.id In (Select t.诊疗项目id From 影像检查项目 t) And a.类别 = 类别_In)
         And (a.编码 Like 条件_In || '%' Or a.名称 Like 条件_In || '%');
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDiagnosisList;

  --62.获得诊疗类别列表
  Procedure p_GetDiagnosisClass(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select t.编码, t.名称, t.简码 From 诊疗项目类别 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDiagnosisClass;

  --63.添加影像报告原型应用信息
  Procedure p_AddMedicalAntetype(
    诊疗项目ID_In 影像报告原型应用.诊疗项目ID%Type,
    应用场合_In   影像报告原型应用.应用场合%Type,
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	) As
  Begin
    Insert Into 影像报告原型应用
      (诊疗项目ID, 应用场合, 报告原型ID)
    Values
      (诊疗项目ID_In, 应用场合_In, 报告原型ID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddMedicalAntetype;

  --64.删除原型ID对应的病历单据应用信息
  Procedure p_DelMedicalAntetype(
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	) As
  Begin
    Delete From 影像报告原型应用 Where 报告原型ID = 报告原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelMedicalAntetype;

  --65.通过原型ID获得对应的病历单据应用信息
  Procedure p_GetMedicalByAID(
	Val           Out t_Refcur,
	报告原型ID_In 影像报告原型应用.报告原型ID%Type
	) As
  Begin
    Open Val For
      Select id,
             x.编码,
             x.名称,
             x.类别,
             Sum(x.门诊) 门诊,
             Sum(x.住院) 住院,
             Sum(x.外诊) 外诊,
             Sum(x.体检) 体检
        From (Select id,
                     编码,
                     名称,
                     类别,
                     Decode(应用场合, '1', 1, 0) as 门诊,
                     Decode(应用场合, '2', 1, 0) as 住院,
                     Decode(应用场合, '3', 1, 0) as 外诊,
                     Decode(应用场合, '4', 1, 0) as 体检
                From (Select to_Char(a.诊疗项目id) ID,
                             (Select b.编码
                                From 诊疗项目目录 b
                               Where b.id = a.诊疗项目id) as 编码,
                             (Select b.名称
                                From 诊疗项目目录 b
                               Where b.id = a.诊疗项目id) as 名称,
                             (Select c.名称
                                From 诊疗项目类别 c
                               Where c.编码 = (Select b.类别
                                               From 诊疗项目目录 b
                                              Where b.id = a.诊疗项目id)) As 类别,
                             a.应用场合
                        From 影像报告原型应用 a
                       Where a.报告原型id = 报告原型ID_In)) x
       Group By x.id, x.编码, x.名称, x.类别;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetMedicalByAID;

  --66.根据原型ID删除动作信息
  Procedure p_DelDocProcessByAid(
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	) As
  Begin
    Delete From 影像报告动作 t Where t.原型id = 报告原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;
  --67.获取ID对应的原型的树形结构
  Procedure p_GetAntetypeTreeByID(
	Val           Out t_Refcur,
	ID_In 影像报告原型清单.ID%Type
	) As
  Begin
    Open Val For
      Select ID, 编码, 名称, 标题, 分组, 是否禁用, 说明, Imageindex
        From (Select Distinct 分组 As ID,
                              (Select Min(b.编码)
                                 From 影像报告原型清单 B
                                Where b.分组 = a.分组) As 编码,
                              a.分组 As 名称,
                              a.分组 As 标题,
                              null As 分组,
                              0 As 是否禁用,
                              null As 说明,
                              0 As Imageindex
                From 影像报告原型清单 A
               Where a.id = ID_In
              Union
              Select RawtoHex(ID) ID,
                     a.编码,
                     a.名称 As 名称,
                     编码 || '-' || 名称 As 标题,
                     分组,
                     a.是否禁用,
                     a.说明,
                     Decode(a.是否禁用, 1, 2, 1) Imageindex
                From 影像报告原型清单 A
               Where a.id = ID_In) A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetAntetypeTreeByID;

  --68.原型是否存在对应的编码或名称
  procedure p_IsExitAntetype(
	Val           Out t_Refcur,
	编码_In 影像报告原型清单.编码%Type,
	名称_In 影像报告原型清单.名称%Type,
	ID_In  影像报告原型清单.ID%Type
	) As
  begin
    Open Val For
      Select Count(*) AS num
        From 影像报告原型清单 t
       where (t.编码 = 编码_In
          or t.名称 = 名称_In) and t.id<>ID_In;
  End p_IsExitAntetype;

  --69.获取影像存储设备
  Procedure p_GetStorageDevice(
	Val           Out t_Refcur
	) Is 
  Begin 
	Open Val For
		Select 设备号||' - '||设备名 As 存储设备, 设备号, IP地址, FTP目录, FTP用户名, FTP密码, 共享目录用户名, 共享目录密码, 共享目录  
		From 影像设备目录 Where 类型 = 1;
	Exception
	  When Others Then
	  Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStorageDevice;
End b_PACS_RptAntetype;
/

--90299:张德婷,2015-11-18,已结账病人不产生未结费用记录
Create Or Replace Procedure Zl_药品收发记录_更改库房
(
  Partid_In       In 药品收发记录.库房id%Type,
  Bill_In         In 药品收发记录.单据%Type,
  No_In           In 药品收发记录.No%Type,
  Otherstockid_In In 药品收发记录.库房id%Type,
  门诊_In         In Number := 1,
  Date_In         In 药品收发记录.填制日期%Type :=Null
) Is
  --重新计算用
  Cursor c_Modifybillout Is
    Select a.Id, a.药品id, a.入出类别id, a.入出系数, a.费用id, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) 数量, Nvl(a.零售金额, 0) 金额, Nvl(a.批次, 0) 批次,
           a.供药单位id, a.成本价, a.批号, a.产地, a.效期, a.生产日期, a.批准文号
    From 药品收发记录 a, 门诊费用记录 b
    Where a.No = No_In And a.单据 = Bill_In And (a.库房id + 0 = Otherstockid_In Or a.库房id Is Null) And
          Nvl(a.摘要, '小宝') <> '拒发' And a.费用id = b.Id And b.执行状态 <> 1 And a.审核人 Is Null;

  Cursor c_Modifybillin Is
    Select a.Id, a.药品id, a.入出类别id, a.入出系数, a.费用id, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) 数量, Nvl(a.零售金额, 0) 金额, Nvl(a.批次, 0) 批次,
           a.供药单位id, a.成本价, a.批号, a.产地, a.效期, a.生产日期, a.批准文号
    From 药品收发记录 a, 住院费用记录 b
    Where a.No = No_In And a.单据 = Bill_In And (a.库房id + 0 = Otherstockid_In Or a.库房id Is Null) And
          Nvl(a.摘要, '小宝') <> '拒发' And a.费用id = b.Id And b.执行状态 <> 1 And a.审核人 Is Null;

  --用于修正病人未结费用
  Cursor c_Billout Is
    Select b.实收金额, b.病人id, 0 主页id, 0 病人病区id, b.病人科室id, b.开单部门id, b.执行部门id, b.收入项目id, b.门诊标志
    From 药品收发记录 a, 门诊费用记录 b
    Where a.费用id = b.Id And b.执行状态 <> 1 And a.No = No_In And a.单据 = Bill_In And
          (a.库房id + 0 = Otherstockid_In Or a.库房id Is Null) And Nvl(a.摘要, '小宝') <> '拒发' And a.审核人 Is Null And b.记录性质 = 2 And
          b.记录状态 = 1;

  Cursor c_Billin Is
    Select b.实收金额, b.病人id, b.主页id, b.病人病区id, b.病人科室id, b.开单部门id, b.执行部门id, b.收入项目id, b.门诊标志
    From 药品收发记录 a, 住院费用记录 b
    Where a.费用id = b.Id And b.执行状态 <> 1 And a.No = No_In And a.单据 = Bill_In And
          (a.库房id + 0 = Otherstockid_In Or a.库房id Is Null) And Nvl(a.摘要, '小宝') <> '拒发' And a.审核人 Is Null And b.记录性质 = 2 And
          b.记录状态 = 1;

  r_Modifybillout   c_Modifybillout%Rowtype;
  r_Modifybillin    c_Modifybillin%Rowtype;
  r_Billout         c_Billout%Rowtype;
  r_Billin          c_Billin%Rowtype;
  Bln收费与发药分离 Number(1);
  v_Count           Number;
Begin
  Begin
    Select 0
    Into Bln收费与发药分离
    From 未发药品记录
    Where 单据 = Bill_In And No = No_In And 库房id + 0 = Otherstockid_In;
  Exception
    When Others Then
      Bln收费与发药分离 := 1;
  End;

  --增加原库房的可以库存，减现库房的可用库存
  If 门诊_In = 1 Then
    --处理门诊
    For r_Modifybillout In c_Modifybillout Loop
      If Bln收费与发药分离 = 0 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Nvl(r_Modifybillout.数量, 0)
        Where 库房id + 0 = Otherstockid_In And 药品id = r_Modifybillout.药品id And 性质 = 1 And Nvl(批次, 0) = r_Modifybillout.批次;
      
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - Nvl(r_Modifybillout.数量, 0)
        Where 库房id + 0 = Partid_In And 药品id = r_Modifybillout.药品id And 性质 = 1 And Nvl(批次, 0) = r_Modifybillout.批次;
        
        If Sql%Rowcount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号)
          Values
            (Partid_In, r_Modifybillout.药品id, r_Modifybillout.批次, 1, 0 - Nvl(r_Modifybillout.数量, 0), 0, 0,
             r_Modifybillout.供药单位id, r_Modifybillout.成本价, r_Modifybillout.批号, r_Modifybillout.产地, r_Modifybillout.效期,
             r_Modifybillout.生产日期, r_Modifybillout.批准文号);
        End If;
      End If;
    End Loop;
  Else
    --处理住院
    For r_Modifybillin In c_Modifybillin Loop
      If Bln收费与发药分离 = 0 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Nvl(r_Modifybillin.数量, 0)
        Where 库房id + 0 = Otherstockid_In And 药品id = r_Modifybillin.药品id And 性质 = 1 And Nvl(批次, 0) = r_Modifybillin.批次;
      
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - Nvl(r_Modifybillin.数量, 0)
        Where 库房id + 0 = Partid_In And 药品id = r_Modifybillin.药品id And 性质 = 1 And Nvl(批次, 0) = r_Modifybillin.批次;
      
        If Sql%Rowcount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号)
          Values
            (Partid_In, r_Modifybillin.药品id, r_Modifybillin.批次, 1, 0 - Nvl(r_Modifybillin.数量, 0), 0, 0,
             r_Modifybillin.供药单位id, r_Modifybillin.成本价, r_Modifybillin.批号, r_Modifybillin.产地, r_Modifybillin.效期,
             r_Modifybillin.生产日期, r_Modifybillin.批准文号);
        End If;
      End If;
    End Loop;
  End If;

  --处理发其它药房处方情况，改变库房ID
  If 门诊_In = 1 Then
    --处理门诊
    For r_Billout In c_Billout Loop
      --减原库房的未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - Nvl(r_Billout.实收金额, 0)
      Where 病人id = r_Billout.病人id And Nvl(主页id, 0) = Nvl(r_Billout.主页id, 0) And
            Nvl(病人病区id, 0) = Nvl(r_Billout.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Billout.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(r_Billout.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(r_Billout.执行部门id, 0) And
            收入项目id + 0 = r_Billout.收入项目id And 来源途径 + 0 = r_Billout.门诊标志;
            
      If Sql%Rowcount <> 0 Then 
        --增加现库房的未结费用
        Update 病人未结费用
        Set 金额 = Nvl(金额, 0) + Nvl(r_Billout.实收金额, 0)
        Where 病人id = r_Billout.病人id And Nvl(主页id, 0) = Nvl(r_Billout.主页id, 0) And
              Nvl(病人病区id, 0) = Nvl(r_Billout.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Billout.病人科室id, 0) And
              Nvl(开单部门id, 0) = Nvl(r_Billout.开单部门id, 0) And Nvl(执行部门id, 0) = Partid_In And 收入项目id + 0 = r_Billout.收入项目id And
              来源途径 + 0 = r_Billout.门诊标志;
        
        If Sql%Rowcount = 0 Then 
          Insert Into 病人未结费用 
            (病人id,病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额) 
          Values 
            (r_Billout.病人id, r_Billout.病人科室id, r_Billout.开单部门id, Partid_In, 
             r_Billout.收入项目id, r_Billout.门诊标志, Nvl(r_Billout.实收金额, 0)); 
        End If;
      end if;
    End Loop;
  Else
    --处理住院
    For r_Billin In c_Billin Loop
      --减原库房的未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - Nvl(r_Billin.实收金额, 0)
      Where 病人id = r_Billin.病人id And Nvl(主页id, 0) = Nvl(r_Billin.主页id, 0) And
            Nvl(病人病区id, 0) = Nvl(r_Billin.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Billin.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(r_Billin.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(r_Billin.执行部门id, 0) And
            收入项目id + 0 = r_Billin.收入项目id And 来源途径 + 0 = r_Billin.门诊标志;
      
      If Sql%Rowcount <> 0 Then 
        --增加现库房的未结费用
        Update 病人未结费用
        Set 金额 = Nvl(金额, 0) + Nvl(r_Billin.实收金额, 0)
        Where 病人id = r_Billin.病人id And Nvl(主页id, 0) = Nvl(r_Billin.主页id, 0) And
              Nvl(病人病区id, 0) = Nvl(r_Billin.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Billin.病人科室id, 0) And
              Nvl(开单部门id, 0) = Nvl(r_Billin.开单部门id, 0) And Nvl(执行部门id, 0) = Partid_In And 收入项目id + 0 = r_Billin.收入项目id And
              来源途径 + 0 = r_Billin.门诊标志;
             
        If Sql%Rowcount = 0 Then 
          Insert Into 病人未结费用 
            (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额) 
          Values 
            (r_Billin.病人id, r_Billin.主页id, r_Billin.病人病区id, r_Billin.病人科室id, r_Billin.开单部门id, Partid_In, r_Billin.收入项目id, 
             r_Billin.门诊标志, Nvl(r_Billin.实收金额, 0)); 
        End If; 
      end if;
    End Loop;
  End If;
  
  delete from 病人未结费用 Where 金额=0;

  If 门诊_In = 1 Then
    Update 门诊费用记录
    Set 执行部门id = Partid_In
    Where Id In
          (Select Distinct 费用id From 药品收发记录 Where No = No_In And 单据 = Bill_In And 库房id + 0 = Otherstockid_In);
  Else
    Update 住院费用记录
    Set 执行部门id = Partid_In
    Where Id In
          (Select Distinct 费用id From 药品收发记录 Where No = No_In And 单据 = Bill_In And 库房id + 0 = Otherstockid_In);
  End If;

  --修改该单据所有记录(退药后再代发的情况)
  Update 药品收发记录 Set 库房id = Partid_In Where No = No_In And 单据 = Bill_In And 库房id + 0 = Otherstockid_In;

  --修改未发药品记录
  Begin
    Select 1 Into v_Count From 未发药品记录 Where 库房id + 0 = Partid_In And No = No_In And 单据 = Bill_In;
  Exception
    When Others Then
      v_Count := 0;
  End;

  If v_Count = 0 Then
    Update 未发药品记录 Set 库房id = Partid_In Where No = No_In And 单据 = Bill_In And 库房id + 0 = Otherstockid_In;
  Else
    Delete 未发药品记录 Where No = No_In And 单据 = Bill_In And 库房id + 0 = Otherstockid_In;
  End If;
  
  If Date_In Is Not Null Then
     Delete From  病人费用汇总 Where 日期>=Date_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_药品收发记录_更改库房;
/

--90046:刘尔旋,2015-11-17,支付宝结帐修改
Create Or Replace Procedure Zl_Retu_Exes
(
  v_No   In Varchar2,
  n_Type In Number
) As
  --------------------------------------------
  --参数:v_No,单据号码
  --     n_Type,单据类型:1-收费,2-记帐,3-自动记帐,4-挂号,5-就诊卡,6-预交,7-结帐
  --------------------------------------------
  n_Allow  Number(1); --是否能够单据返回
  n_Patiid Number(18);
  Err_Item Exception;
  v_Err_Msg Varchar2(100);
  n_System  Number(5);
  n_只读    Number(2);

  v_Table  Varchar2(100);
  v_Field  Varchar2(100);
  v_Sql    Varchar2(4000);
  v_Fields Varchar2(4000);

  --功能：获取表的字段字符串
  Function Getfields(v_Table In Varchar2) Return Varchar2 As
    v_Colstr Varchar2(4000);
  Begin
    Select f_List2str(Cast(Collect(Column_Name) As t_Strlist)) As Colsstr
    Into v_Colstr
    From (Select Column_Name From User_Tab_Columns Where Table_Name = v_Table Order By Column_Id);
  
    Return v_Colstr;
  End Getfields;

  --------------------------------------------
  --返回指定ID的病人预交记录子过程
  --------------------------------------------
  Procedure Zl_Retu_Prepay(n_Settle_Id H病人预交记录.结帐id%Type) As
  Begin
    For r_Rec In (Select * From H病人预交记录 Where 结帐id = n_Settle_Id) Loop
      v_Table  := '病人卡结算对照';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where 预交id = :1';
      Execute Immediate v_Sql
        Using r_Rec.Id;
    
      v_Table  := '三方结算交易';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where 交易ID = :1';
      Execute Immediate v_Sql
        Using r_Rec.Id;
        
      v_Table  := '三方退款信息';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where 记录ID = :1 And 结帐ID = :2';
      Execute Immediate v_Sql
        Using r_Rec.Id,n_Settle_Id;
    
      v_Table  := '病人卡结算记录';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where ID In(Select 卡结算id From H病人卡结算对照 Where 预交id = :1)';
      Execute Immediate v_Sql
        Using r_Rec.Id;
    
      Delete H病人卡结算记录 Where ID In (Select Distinct 卡结算id From H病人卡结算对照 Where 预交id = r_Rec.Id);
      Delete From H病人卡结算对照 Where 预交id = r_Rec.Id;
      Delete From H三方结算交易 Where 交易id = r_Rec.Id;
      Delete From H三方退款信息 Where 记录id = r_Rec.Id And 结帐ID=n_Settle_Id;
    End Loop;
  
    v_Table  := '病人预交记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where 结帐id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    Delete H病人预交记录 Where 结帐id = n_Settle_Id;
  End Zl_Retu_Prepay;

  --------------------------------------------
  --返回指定ID的病人费用记录子过程
  --------------------------------------------
  Procedure Zl_Retu_Fee(n_Settle_Id H住院费用记录.结帐id%Type) As
  Begin
    --返回病人费用记录
    v_Table  := '住院费用记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where 结帐id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    v_Table  := '门诊费用记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where 结帐id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    v_Table  := '费用补充记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where 结算id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    v_Table  := '医保结算明细';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where 结帐id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    --删除已返回的费用记录
    Delete H门诊费用记录 Where 结帐id = n_Settle_Id;
    Delete H住院费用记录 Where 结帐id = n_Settle_Id;
    Delete H费用补充记录 Where 结算id = n_Settle_Id;
    Delete H医保结算明细 Where 结帐id = n_Settle_Id;
  End Zl_Retu_Fee;

  --------------------------------------------
  --返回指定ID的药品收发记录子过程
  --------------------------------------------
  Procedure Zl_Retu_Medilist(n_Rec_Id H药品收发记录.Id%Type) As
  Begin
    --按外键引用顺序返回药品收发相关表的数据     
    v_Table  := '药品收发记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where ID = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    v_Table  := '输液配药记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where ID In(Select 记录ID From H输液配药内容 Where 收发ID =:1)';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    For P In (Select ID From H输液配药记录 Where ID In (Select 记录id From H输液配药内容 Where 收发id = n_Rec_Id)) Loop
      For R In (Select Column_Value From Table(f_Str2list('输液配药附费,输液配药状态'))) Loop
        v_Table := r.Column_Value;
        v_Field := 'ID';
      
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                    ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.Id;
      
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.Id;
      End Loop;
    End Loop;
  
    Delete H输液配药记录 Where ID In (Select 记录id From H输液配药内容 Where 收发id = n_Rec_Id);
  
    v_Table  := '药品签名记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where ID In(Select 签名ID From H药品签名明细 Where 收发ID =:1)';
    Execute Immediate v_Sql
      Using n_Rec_Id;
    Delete H药品签名记录 Where ID In (Select 签名id From H药品签名明细 Where 收发id = n_Rec_Id);
  
    For R In (Select Column_Value From Table(f_Str2list('收发记录补充信息,输液配药内容,药品签名明细,药品留存计划'))) Loop
      v_Table := r.Column_Value;
      If v_Table = '药品留存计划' Then
        v_Field := '留存ID';
      Else
        v_Field := '收发ID';
      End If;
    
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      v_Sql := 'Delete From H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    End Loop;
  
    --删除已返回的药品收发记录
    Delete H药品收发记录 Where ID = n_Rec_Id;
  End Zl_Retu_Medilist;

  --------------------------------------------
  --以下为主程序体
  --------------------------------------------
Begin
  ----------------------------------------------------------------------------------------------------------
  --刘兴宏:主要是对基于视图的视图的转储方案进行了只读判断.
  Select 编号 Into n_System From zlSystems Where Upper(所有者) = Zl_Owner And 编号 Like '1%';
  Begin
    Select Nvl(只读, 0) Into n_只读 From zlBakSpaces Where 系统 = n_System And 当前 = 1;
  Exception
    When Others Then
      v_Err_Msg := '[ZLSOFT]当前没有可用的历史数据空间,不能继续![ZLSOFT]';
      Raise Err_Item;
  End;
  If n_只读 = 1 Then
    v_Err_Msg := '[ZLSOFT]历史数据空间目前的状态为只读,不能继续![ZLSOFT]';
    Raise Err_Item;
  End If;

  --判断是否能按照单据返回
  Select Decode(Sum(Nvl(p.金额, 0)) - Sum(Nvl(p.冲预交, 0)), Null, 1, 0, 1, 0)
  Into n_Allow
  From H病人预交记录 P,
       (Select 结帐id
         From H门诊费用记录
         Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
               4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5)
         Union
         Select 结帐id
         From H住院费用记录
         Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
               4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5)
         Union
         Select 结帐id
         From H病人预交记录
         Where NO = v_No And 6 = n_Type And 记录性质 In (1, 11)
         Union
         Select ID From H病人结帐记录 Where NO = v_No And 7 = n_Type) L
  Where p.结帐id = l.结帐id And p.记录性质 In (1, 11);
  If n_Allow = 1 Then
    Select Decode(Sum(Nvl(e.实收金额, 0)) - Sum(Nvl(e.结帐金额, 0)), Null, 1, 0, 1, 0)
    Into n_Allow
    From (Select e.实收金额, e.结帐金额
           From H门诊费用记录 E,
                (Select 结帐id
                  From H门诊费用记录
                  Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
                        4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5)
                  Union
                  Select 结帐id
                  From H病人预交记录
                  Where NO = v_No And 6 = n_Type And 记录性质 In (1, 11)
                  Union
                  Select ID From H病人结帐记录 Where NO = v_No And 7 = n_Type) L
           Where e.结帐id = l.结帐id
           Union All
           Select e.实收金额, e.结帐金额
           From H住院费用记录 E,
                (Select 结帐id
                  From H住院费用记录
                  Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
                        4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5)
                  Union
                  Select 结帐id
                  From H病人预交记录
                  Where NO = v_No And 6 = n_Type And 记录性质 In (1, 11)
                  Union
                  Select ID From H病人结帐记录 Where NO = v_No And 7 = n_Type) L
           Where e.结帐id = l.结帐id) E;
  End If;

  --按照单据或病人获取结帐游标返回
  If n_Allow = 1 Then
    For r_Settle In (Select 结帐id
                     From H门诊费用记录
                     Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
                           4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5)
                     Union
                     Select 结帐id
                     From H住院费用记录
                     Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
                           4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5)
                     Union
                     Select 结帐id
                     From H病人预交记录
                     Where NO = v_No And 6 = n_Type And 记录性质 In (1, 11)
                     Union All
                     Select ID From H病人结帐记录 Where NO = v_No And 7 = n_Type) Loop
    
      v_Table  := '病人结帐记录';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where id = :1';
      Execute Immediate v_Sql
        Using r_Settle.结帐id;
    
      Zl_Retu_Prepay(r_Settle.结帐id);
      For r_Rxlist In (Select m.Id
                       From H药品收发记录 M,
                            (Select ID, NO, 序号, 记录性质
                              From H门诊费用记录
                              Where 结帐id = r_Settle.结帐id And 收费类别 In ('4', '5', '6', '7') And 记录性质 In (1, 2)) E
                       Where m.No = e.No And m.费用id = e.Id And
                             (e.记录性质 = 1 And m.单据 In (8, 24) Or e.记录性质 <> 1 And m.单据 In (9, 10, 25, 26))
                       Union All
                       Select m.Id
                       From H药品收发记录 M,
                            (Select ID, NO, 序号, 记录性质
                              From H住院费用记录
                              Where 结帐id = r_Settle.结帐id And 收费类别 In ('4', '5', '6', '7') And 记录性质 In (1, 2)) E
                       Where m.No = e.No And m.费用id = e.Id And
                             (e.记录性质 = 1 And m.单据 In (8, 24) Or e.记录性质 <> 1 And m.单据 In (9, 10, 25, 26))) Loop
        Zl_Retu_Medilist(r_Rxlist.Id);
      End Loop;
      Zl_Retu_Fee(r_Settle.结帐id);
    
      Delete H病人结帐记录 Where ID = r_Settle.结帐id;
    End Loop;
  Else
    Begin
      --n_Type,单据类型:1-收费,2-记帐,3-自动记帐,4-挂号,5-就诊卡,6-预交,7-结帐
      If n_Type = 7 Then
        Select Distinct 病人id Into n_Patiid From H病人结帐记录 Where NO = v_No;
      Elsif n_Type = 6 Then
        Select Distinct 病人id Into n_Patiid From H病人预交记录 Where NO = v_No And 6 = n_Type And 记录性质 In (1, 11);
      Elsif n_Type = 5 Or n_Type = 3 Then
        Select Distinct 病人id
        Into n_Patiid
        From H住院费用记录
        Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
              4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5);
      Elsif n_Type = 4 Or n_Type = 1 Then
        If n_Type = 1 Then
          Select Distinct 病人id
          Into n_Patiid
          From (Select Distinct 病人id
                 From H门诊费用记录
                 Where NO = v_No And 记录性质 = 1
                 Union All
                 Select Distinct 病人id From H费用补充记录 Where NO = v_No And 记录性质 = 1)
          Where Rownum < 2;
        
        Else
          Select Distinct 病人id
          Into n_Patiid
          From H门诊费用记录
          Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
                4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5);
        End If;
      Else
        Begin
          Select Distinct 病人id
          Into n_Patiid
          From H住院费用记录
          Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
                4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5);
        Exception
          When Others Then
            n_Patiid := -1;
        End;
        If Nvl(n_Patiid, 0) <= 0 Then
          Select Distinct 病人id
          Into n_Patiid
          From H门诊费用记录
          Where NO = v_No And (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
                4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5);
        End If;
      End If;
    Exception
      When Others Then
        n_Patiid := Null;
    End Zl_Patiid;
  
    For r_Settle In (Select Distinct 结帐id
                     From H门诊费用记录
                     Where 病人id = n_Patiid
                     Union
                     Select Distinct 结帐id
                     From H住院费用记录
                     Where 病人id = n_Patiid
                     Union
                     Select Distinct 结算id From H费用补充记录 Where 病人id = n_Patiid) Loop
    
      v_Table  := '病人结帐记录';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where id = :1';
      Execute Immediate v_Sql
        Using r_Settle.结帐id;
    
      Zl_Retu_Prepay(r_Settle.结帐id);
      For r_Rxlist In (Select m.Id
                       From H药品收发记录 M,
                            (Select ID, NO, 序号, 记录性质
                              From H门诊费用记录
                              Where 结帐id = r_Settle.结帐id And 收费类别 In ('4', '5', '6', '7') And 记录性质 In (1, 2)) E
                       Where m.No = e.No And m.费用id = e.Id And
                             (e.记录性质 = 1 And m.单据 In (8, 24) Or e.记录性质 <> 1 And m.单据 In (9, 10, 25, 26))
                       Union All
                       Select m.Id
                       From H药品收发记录 M,
                            (Select ID, NO, 序号, 记录性质
                              From H住院费用记录
                              Where 结帐id = r_Settle.结帐id And 收费类别 In ('4', '5', '6', '7') And 记录性质 In (1, 2)) E
                       Where m.No = e.No And m.费用id = e.Id And
                             (e.记录性质 = 1 And m.单据 In (8, 24) Or e.记录性质 <> 1 And m.单据 In (9, 10, 25, 26))
                       
                       ) Loop
        Zl_Retu_Medilist(r_Rxlist.Id);
      End Loop;
    
      Zl_Retu_Fee(r_Settle.结帐id);
      Delete H病人结帐记录 Where ID = r_Settle.结帐id;
    
    End Loop;
  End If;

  Begin
    Execute Immediate 'Update zlBakInfo Set 最后转储日期=Sysdate Where 系统=' || n_System;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM || ':' || v_Sql);
End Zl_Retu_Exes;
/

--90531:刘尔旋,2015-11-17,退号性能问题
Create Or Replace Procedure Zl_病人挂号补结算_Delete
(
  单据号_In     门诊费用记录.No%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  冲销id_In     门诊费用记录.结帐id%Type := Null,
  结算序号_In   病人预交记录.结算序号%Type := Null,
  退号时间_In   门诊费用记录.登记时间%Type := Null,
  删除门诊号_In Number := 0
) As
  --该游标用于判断是否单独收病历费,及挂号汇总表处理

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_结帐id   病人预交记录.结帐id%Type;
  n_冲销id   门诊费用记录.结帐id%Type;
  n_结算序号 病人预交记录.结算序号%Type;

  n_病人id   病人信息.病人id%Type;
  n_退费金额 病人预交记录.冲预交%Type;
  n_挂号id   病人挂号记录.Id%Type;
  n_组id     财务缴款分组.Id%Type;

  n_分诊台签到排队 Number;
  n_挂号生成队列   Number;
  d_Date           Date;
  n_病人id1        病人信息.病人id%Type;
  d_Temp           Date;
  n_就诊病人id     病人信息.病人id%Type;
  d_就诊时间       就诊登记记录.就诊时间%Type;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  --首先判断要退号/取消预约的记录是否存在

  Begin
    Select a.结帐id, a.病人id
    Into n_结帐id, n_病人id
    From 门诊费用记录 A
    Where a.记录性质 = 4 And a.No = 单据号_In And a.记录状态 = 1 And Rownum < 2;
  Exception
  
    When Others Then
      n_病人id := -1;
  End;
  If Nvl(n_病人id, 0) = -1 Then
    v_Err_Msg := '未找到指定的挂号单:' || 单据号_In || ',可能已经被人退号,不允许再次退号。';
    Raise Err_Item;
  End If;

  --2.挂号处理

  d_Date     := 退号时间_In;
  n_冲销id   := 冲销id_In;
  n_结算序号 := 结算序号_In;

  If d_Date Is Null Then
    d_Date := Sysdate;
  End If;
  If n_冲销id Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_冲销id From Dual;
  End If;

  If n_结算序号 Is Null Then
    n_结算序号 := -1 * n_冲销id;
  End If;
  --更新挂号序号状态
  If Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111)) = 1 Then
    Delete 挂号序号状态
    Where 状态 = 1 And
          (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
          (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
  Else
    Update 挂号序号状态
    Set 状态 = 4
    Where 状态 = 1 And
          (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
          (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
  End If;

  --病人就诊状态
  If n_病人id Is Not Null Then
    Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
    --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理(界面要根据提示来删除)
    If 删除门诊号_In = 1 Then
      Delete 门诊病案记录 Where 病人id = n_病人id;
      Update 病人信息 Set 门诊号 = Null Where 病人id = n_病人id;
      --费用记录包括挂号及病案、就诊卡费用,以及病人交费后退费或销帐的费用,挂号记录在最后处理
      Update 门诊费用记录 Set 标识号 = Null Where 门诊标志 = 1 And 病人id = n_病人id;
    End If;
  End If;

  --如果挂时收了就诊卡费,退费时清除就诊卡号,在非光退病历费时
  n_病人id1 := Null;
  Begin
    Select 病人id
    Into n_病人id1
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 附加标志 = 2 And Rownum < 2;
  Exception
    When Others Then
      Null;
  End;
  If n_病人id1 Is Not Null Then
    Update 病人信息
    Set 就诊卡号 = Null, 卡验证码 = Null, Ic卡号 = Decode(Ic卡号, 就诊卡号, Null, Ic卡号)
    Where 病人id = n_病人id1;
  End If;

  --门诊费用记录
  --冲销记录
  Insert Into 门诊费用记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
     数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
     结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id, 费用状态, 执行状态)
    Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
           收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
           操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_冲销id, -1 * 结帐金额, 保险项目否, 保险大类id, -1 * 统筹金额, 摘要 As 摘要, 附加标志, 保险编码, 费用类型,
           n_组id, 1, 1
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;

  Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
  Select Sum(实收金额) Into n_退费金额 From 门诊费用记录 Where 结帐id = n_冲销id;
  Insert Into 病人预交记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 结算序号, 缴款组id, 预交类别, 卡类别id,
     结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算性质)
    Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, Null, d_Date, 操作员编号_In, 操作员姓名_In,
           n_退费金额, n_冲销id, n_结算序号, n_组id, 预交类别, Null, Null, Null, Null, Null, Null, 1, 4
    From 病人预交记录 A
    Where a.记录性质 = 4 And a.结帐id = n_结帐id And Rownum = 1;
  If Sql%NotFound Then
    v_Err_Msg := '未找到挂号单为【' || 单据号_In || '】的原始挂号记录!';
    Raise Err_Item;
  End If;
  Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id;

  --病人挂号汇总
  For c_挂号 In (Select a.收费细目id, a.发生时间, a.登记时间, c.接收时间, c.执行部门id, c.执行人, m.Id As 医生id, Nvl(c.号别, b.号码) As 号码, a.结帐id,
                      a.病人id, Decode(c.预约, Null, 0, 0, 0, 1) As 预约
               From 门诊费用记录 A, 病人挂号记录 C, 挂号安排 B, 人员表 M
               Where a.记录性质 = 4 And a.结帐id = n_冲销id And a.从属父号 Is Null And c.执行人 = m.姓名(+) And a.No = c.No And
                     Nvl(c.号别, Nvl(a.计算单位, '-')) = b.号码 And Nvl(a.附加标志, 0) = 0 And Rownum < 2) Loop
    --退非挂号费用,则不处理汇总表数据
  
    If Nvl(c_挂号.预约, 0) <> 0 Then
      d_Temp := Trunc(c_挂号.接收时间);
    Else
      d_Temp := Trunc(c_挂号.发生时间);
    End If;
  
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - Nvl(c_挂号.预约, 0), 已约数 = Nvl(已约数, 0) - Nvl(c_挂号.预约, 0)
    Where 日期 = d_Temp And 科室id = c_挂号.执行部门id And 项目id = c_挂号.收费细目id And Nvl(医生姓名, '医生') = Nvl(c_挂号.执行人, '医生') And
          Nvl(医生id, 0) = Nvl(c_挂号.医生id, 0) And (号码 = c_挂号.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收, 已约数)
      Values
        (d_Temp, c_挂号.执行部门id, c_挂号.收费细目id, c_挂号.执行人, Decode(c_挂号.医生id, 0, Null, c_挂号.医生id), c_挂号.号码, -1,
         -1 * Nvl(c_挂号.预约, 0), -1 * Nvl(c_挂号.预约, 0));
    End If;
  End Loop;

  n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
  If n_挂号生成队列 <> 0 Then
    n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
    If Nvl(n_分诊台签到排队, 0) = 0 Then
      --要删除队列
      For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
        Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
      End Loop;
    End If;
  End If;

  --医保产生的就诊登记记录
  Begin
    Select 病人id, 发生时间 Into n_就诊病人id, d_就诊时间 From 病人挂号记录 Where NO = 单据号_In;
    Delete From 就诊登记记录 Where 病人id = n_就诊病人id And 就诊时间 = d_就诊时间 And 主页id Is Null;
  Exception
    When Others Then
      Null;
  End;
  --病人挂号记录
  Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;

  Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 1;
  If Sql%NotFound Then
    v_Err_Msg := '挂号单【' || 单据号_In || '】不存在或由于并发原因已经被退号';
    Raise Err_Item;
  End If;

  Insert Into 病人挂号记录
    (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名, 复诊,
     号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式)
    Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间, 操作员编号_In,
           操作员姓名_In, 复诊, 号序, 社区, 预约, 摘要 As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式
    From 病人挂号记录
    Where NO = 单据号_In And 记录状态 = 3;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号补结算_Delete;
/

--90531:刘尔旋,2015-11-17,退号性能问题
Create Or Replace Procedure Zl_病人挂号记录_Delete
(
  单据号_In       门诊费用记录.No%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
  删除门诊号_In   Number := 0,
  非原样退结算_In Varchar2 := Null,
  退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费
  退指定结算_In   病人预交记录.结算方式%Type := Null,
  退号重用_In     Number := 1
) As
  --退费类型_In,在一下几种情况下不准进行部分退费
  --    2.三方接口,暂时不支持
  -- 挂号费病历费分开退,规则
  --    普通结算方式:原结算方式退部分费用
  --    预交款:预交款,退部分
  --    预交款与普通结算方式混合:退款按照普通结算方式部分退
  --    消费卡:原样将费用部分退入消费卡
  --非原样退结算_In:指不能退还给原样结算方式(如医保的个人账户,三方账户的退现等),多个用逗分离
  --退指定结算_IN:指非原样退结算部分,应该退给哪种结算方式,为空时缺省退给现金,否则退给指定的结算方式

  --该游标用于判断是否单独收病历费,及挂号汇总表处理
  Cursor c_Registinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select a.发生时间, a.登记时间, c.接收时间, a.收费细目id As 项目id, c.执行部门id As 科室id, c.执行人 As 医生姓名, d.Id As 医生id, c.号别 As 号码
    From 门诊费用记录 A, 挂号安排 B, 病人挂号记录 C, 人员表 D
    Where a.记录性质 = 4 And a.记录状态 = v_状态 And c.No = a.No And c.执行人 = d.姓名(+) And a.No = 单据号_In And
          Nvl(a.计算单位, '号别') = c.号别 And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --该游标用于判断记录是否存在,及费用汇总表处理
  Cursor c_Moneyinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(应收金额), 0) As 应收, Nvl(Sum(实收金额), 0) As 实收, Nvl(Sum(结帐金额), 0) As 结帐
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = v_状态 And NO = 单据号_In
    Group By 病人科室id, 开单部门id, 执行部门id, 收入项目id;
  r_Moneyrow c_Moneyinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Opermoney Is
    Select Distinct b.结算方式, -1 * Nvl(b.冲预交, 0) As 冲预交
    From 门诊费用记录 A, 病人预交记录 B
    Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 2 And b.记录性质 = 4 And b.记录状态 = 2 And
          Nvl(b.冲预交, 0) <> 0 And
          Nvl(a.附加标志, 0) =
          Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(a.附加标志, 0), 1, -1, Nvl(a.附加标志, 0)), Nvl(a.附加标志, 0));

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_结帐id 病人预交记录.结帐id%Type;
  n_销帐id 门诊费用记录.结帐id%Type;

  v_退指定结算方式 病人预交记录.结算方式%Type;
  n_退款金额       病人预交记录.冲预交%Type;
  n_打印id         票据打印内容.Id%Type;
  n_病人id         病人信息.病人id%Type;
  n_退费金额       病人预交记录.冲预交%Type;
  n_预交金额       病人预交记录.冲预交%Type; --原记录 预交缴款金额
  n_返回值         病人余额.预交余额%Type;
  n_挂号id         病人挂号记录.Id%Type;
  n_组id           财务缴款分组.Id%Type;

  n_二次退费       Number; --记录是否是此单据的第二次退费
  n_分诊台签到排队 Number;
  n_预约生成队列   Number;
  n_预约挂号       Number;
  n_挂号生成队列   Number;
  d_Date           Date;
  n_记帐           门诊费用记录.记帐费用%Type;
  n_病人id1        病人信息.病人id%Type;
  n_返回额         门诊费用记录.实收金额%Type;
  n_已结帐         Number;
  n_就诊病人id     病人信息.病人id%Type;
  d_就诊时间       就诊登记记录.就诊时间%Type;
Begin
  n_组id           := Zl_Get组id(操作员姓名_In);
  v_退指定结算方式 := 退指定结算_In;

  --首先判断要退号/取消预约的记录是否存在
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := '要处理的单据不存在。';
      Raise Err_Item;
    End If;
    n_预约挂号 := 1;
  End If;
  Close c_Moneyinfo;

  --1.预约处理
  If Nvl(n_预约挂号, 0) = 1 Then
    --减少已约数
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) - 1
    Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
          (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1);
    End If;
    Close c_Registinfo;
  
    --更新挂号序号状态
    Delete 挂号序号状态
    Where 状态 = 2 And
          (号码, 序号, 日期) = (Select 计算单位, 发药窗口, Trunc(发生时间)
                          From 门诊费用记录
                          Where 记录性质 = 4 And 记录状态 = 0 And 序号 = 1 And Rownum = 1 And NO = 单据号_In) Or
          (号码, 序号, 日期) = (Select 计算单位, 发药窗口, 发生时间
                          From 门诊费用记录
                          Where 记录性质 = 4 And 记录状态 = 0 And 序号 = 1 And Rownum = 1 And NO = 单据号_In);
  
    --添加病人挂号记录的 冲销记录
    Select 病人挂号记录_Id.Nextval, Sysdate Into n_挂号id, d_Date From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1 And 记录性质 = 2;
    If Sql%NotFound Then
      v_Err_Msg := '预约单【' || 单据号_In || '】不存在或由于并发原因已经被取消预约';
      Raise Err_Item;
    End If;
  
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式
      From 病人挂号记录
      Where NO = 单据号_In;
  
    --Update 病人挂号记录 set 摘要=nvl(摘要_IN,摘要) where NO=单据号_IN;
    --删除门诊费用记录
    Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
    --如果预约生成队列时需要清除队列
  
    n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    If Nvl(n_预约生成队列, 0) = 1 Then
      --要删除队列
      For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
        Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(记帐费用, 0), 病人id, Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
  Into n_记帐, n_病人id, n_已结帐
  From 门诊费用记录
  Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;

  --2.挂号处理
  n_已结帐 := Nvl(n_已结帐, 0);

  If n_已结帐 = 1 And n_记帐 = 1 Then
    Select Sysdate, Null Into d_Date, n_销帐id From Dual;
  Else
    Select Sysdate, 病人结帐记录_Id.Nextval Into d_Date, n_销帐id From Dual;
  End If;

  ----0-全退 1-退挂号费 2-退病历费
  If Nvl(退费类型_In, 0) <> 2 Then
    --不是光退病历费时处理
    --更新挂号序号状态
    If 退号重用_In = 1 Then
      Delete 挂号序号状态
      Where 状态 = 1 And
            (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
            (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
    Else
      Update 挂号序号状态
      Set 状态 = 4
      Where 状态 = 1 And
            (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
            (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
    End If;
  
    --病人就诊状态
    If n_病人id Is Not Null Then
      Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
    
      --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      If 删除门诊号_In = 1 Then
        Delete 门诊病案记录 Where 病人id = n_病人id;
        Update 病人信息 Set 门诊号 = Null Where 病人id = n_病人id;
        --费用记录包括挂号及病案、就诊卡费用,以及病人交费后退费或销帐的费用,挂号记录在最后处理
        Update 门诊费用记录 Set 标识号 = Null Where 门诊标志 = 1 And 病人id = n_病人id;
      End If;
    End If;
  
    --如果挂时收了就诊卡费,退费时清除就诊卡号,在非光退病历费时
    n_病人id1 := Null;
    Begin
      Select 病人id
      Into n_病人id1
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 附加标志 = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_病人id1 Is Not Null And Nvl(退费类型_In, 0) <> 2 Then
      Update 病人信息
      Set 就诊卡号 = Null, 卡验证码 = Null, Ic卡号 = Decode(Ic卡号, 就诊卡号, Null, Ic卡号)
      Where 病人id = n_病人id1;
    End If;
  
  End If;

  --检查前面是否已经部分退过费用
  Begin
    Select 1 Into n_二次退费 From 门诊费用记录 Where 记录性质 = 4 And NO = 单据号_In And 记录状态 = 3 And Rownum < 2;
  Exception
    When Others Then
      n_二次退费 := 0;
  End;

  --门诊费用记录
  --冲销记录
  Insert Into 门诊费用记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
     数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
     结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
    Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
           收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
           操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
           Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
           Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
          Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));

  --原始记录
  If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
    Update 门诊费用记录
    Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
          Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  Else
    Update 门诊费用记录
    Set 记录状态 = 3
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
          Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  End If;

  n_结帐id := 0;
  If n_记帐 = 0 Then
    --获取结帐ID
    Select Nvl(结帐id, 0)
    Into n_结帐id
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And
          Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0)) And
          Rownum = 1;
  End If;

  If n_记帐 = 1 Then
    --记帐
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And
                       Nvl(附加标志, 0) =
                       Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0)) And
                       Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) - Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1
      Returning 费用余额 Into n_返回额;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (n_病人id, 1, 1, -1 * Nvl(c_费用.实收金额, 0), 0);
        n_返回额 := Nvl(c_费用.实收金额, 0);
      End If;
      If Nvl(n_返回额, 0) = 0 Then
        Delete 病人余额
        Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - Nvl(c_费用.实收金额, 0)
      Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, -1 * Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
    Delete 病人未结费用
    Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(金额, 0) = 0 And 来源途径 + 0 = 1;
  End If;

  If n_记帐 = 0 Then
    --1.退费
    --病人挂号结算:现金和个人帐户部份
    If 非原样退结算_In Is Not Null Then
      --退款金额获取
      If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
        --如果是单独退病历费,或者只退挂号费,先获取退费金额
        Begin
          --获取本次退款金额
          Select Sum(Nvl(实收金额, 0)) As 收款金额
          Into n_退款金额
          From 门诊费用记录
          Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 3 And
                Nvl(附加标志, 0) =
                Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
        
        Exception
          When Others Then
            v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                           When 1 Then
                            '挂号费用'
                           When 2 Then
                            '病历费'
                         End || '可能由于并发原因已经进行了退费或者单据不存在!';
            Raise Err_Item;
        End;
        Begin
          Select 冲预交
          Into n_退费金额
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
        Exception
          When Others Then
            n_退费金额 := 0;
        End;
      
        --a.允许的结算方式
      
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -n_退款金额,
                 n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
      
        If n_退费金额 = 0 Then
          --b.不允许的退现金
          If n_退款金额 <> 0 Then
            If v_退指定结算方式 Is Null Then
              --退给现金
              Begin
                Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
              Exception
                When Others Then
                  v_退指定结算方式 := '现金';
              End;
            End If;
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * n_退款金额)
            Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                       操作员编号_In, 操作员姓名_In, -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
                From 病人预交记录 A
                Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --a.允许的结算方式原样退
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -冲预交,
                 n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
      
        --b.不允许的退现金
        Begin
          Select Sum(冲预交)
          Into n_退费金额
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') > 0;
        Exception
          When Others Then
            n_退费金额 := 0;
        End;
        If n_退费金额 <> 0 Then
          If v_退指定结算方式 Is Null Then
            --退给现金
            Begin
              Select 结算方式
              Into v_退指定结算方式
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
            
            Exception
              When Others Then
                Begin
                  Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
                Exception
                  When Others Then
                    v_退指定结算方式 := '现金';
                End;
            End;
          End If;
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * n_退费金额)
          Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                     操作员编号_In, 操作员姓名_In, -1 * n_退费金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录 A
              Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
          End If;
        End If;
      End If;
    Else
      --退款金额获取
      If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
        --如果是单独退病历费,或者只退挂号费,先获取退费金额
        Begin
          --获取本次退款金额
          Select Sum(Nvl(实收金额, 0)) As 收款金额
          Into n_退款金额
          From 门诊费用记录
          Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 3 And
                Nvl(附加标志, 0) =
                Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
        Exception
          When Others Then
            v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                           When 1 Then
                            '挂号费用'
                           When 2 Then
                            '病历费'
                         End || '可能由于并发原因已经进行了退费或者单据不存在!';
            Raise Err_Item;
        End;
      End If;
      If Nvl(n_二次退费, 0) = 0 And Nvl(退费类型_In, 0) = 0 Then
        --首次全退
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -1 * 冲预交,
                 n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id;
      Else
        --二次退费,或者本次单退一部分
        --二次退费时,记录状态=3 ,首次部分退,记录状态为1
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                 -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 摘要 = '医保挂号' And 冲预交 = n_退款金额 And
                Rownum < 2;
        If Sql%RowCount = 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 冲预交 = n_退款金额 And Rownum < 2;
        End If;
        If Sql%RowCount = 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And Rownum < 2;
          If Sql%RowCount = 0 Then
            --部分退费,并且全部使用预交款缴费时才存在此种情况
            n_预交金额 := n_退款金额;
          End If;
        End If;
      
      End If;
    End If;
    --首次退费时,记录状态便调整为了3
    Update 病人预交记录
    Set 记录状态 = 3
    Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id;
  
    --冲预交 1-全退 2-部分退,部分退时当全部使用预交进行缴款
    If Nvl(退费类型_In, 0) = 0 Or (Nvl(退费类型_In, 0) <> 0 And n_预交金额 <> 0) Then
      --病人挂号结算:冲预交款部份
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date,
               操作员姓名_In, 操作员编号_In, -1 * Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, 冲预交, n_预交金额), n_销帐id, n_组id, 预交类别,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
        From 病人预交记录
        Where 记录性质 In (1, 11) And 结帐id = n_结帐id And Nvl(冲预交, 0) <> 0 And
              Rownum = Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, Rownum, 1);
    End If;
  
    --处理病人预交余额
    For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                 From 病人预交记录
                 Where 记录性质 In (1, 11) And 结帐id = n_销帐id
                 Group By 病人id, 预交类别) Loop
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.冲预交, 0)
      Where 病人id = c_预交.病人id And 类型 = Nvl(c_预交.预交类别, 2) And 性质 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 预交余额, 性质, 类型)
        Values
          (c_预交.病人id, Nvl(c_预交.冲预交, 0), 1, Nvl(c_预交.预交类别, 2));
        n_返回值 := Nvl(c_预交.冲预交, 0);
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    End Loop;
  
    If Nvl(退费类型_In, 0) <> 2 Then
      --光退挂号费,不回收票据
      --退卡收回票据(可能上次挂号使用票据,不能收回)
      Begin
        --从最后一次的打印内容中取
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    
      If n_打印id Is Not Null Then
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
          From 票据使用明细
          Where 打印id = n_打印id And 性质 = 1;
      End If;
    End If;
  End If;

  --单独退病历费用,不处理汇总记录
  --相关汇总表的处理

  --病人挂号汇总
  Open c_Registinfo(3);
  Fetch c_Registinfo
    Into r_Registrow;

  If c_Registinfo%RowCount = 0 Then
    --只收病历费时无号别,不处理
    Close c_Registinfo;
  Else
  
    --需要确定是否预约挂号
    --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
    --2.如果是正常挂号,则只减已挂数
  
    Begin
      Select Decode(预约, Null, 0, 0, 0, 1) Into n_预约挂号 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1;
    Exception
      When Others Then
        n_预约挂号 := 0;
    End;
  
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
    Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
          (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收, 已约数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号, -1 * n_预约挂号);
    End If;
  
    Close c_Registinfo;
  End If;

  If n_记帐 = 0 Then
    --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
    For r_Opermoney In c_Opermoney Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
      Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
        n_返回值 := r_Opermoney.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(退费类型_In, 0) <> 2 Then
    n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
    If n_挂号生成队列 <> 0 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
      If Nvl(n_分诊台签到排队, 0) = 0 Then
      
        --要删除队列
        For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
          Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
        End Loop;
      End If;
    End If;
  
    --医保产生的就诊登记记录
    Begin
      Select 病人id, 发生时间 Into n_就诊病人id, d_就诊时间 From 病人挂号记录 Where NO = 单据号_In;
      Delete From 就诊登记记录 Where 病人id = n_就诊病人id And 就诊时间 = d_就诊时间 And 主页id Is Null;
    Exception
      When Others Then
        Null;
    End;
    
    --病人挂号记录
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 1;
    If Sql%NotFound Then
      v_Err_Msg := '挂号单【' || 单据号_In || '】不存在或由于并发原因已经被退号';
      Raise Err_Item;
    End If;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式
      From 病人挂号记录
      Where NO = 单据号_In;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 2, 单据号_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_Delete;
/

--90389:冉俊明,2015-11-16,医嘱附费项目退费数量错误。
Create Or Replace Procedure Zl_门诊收费记录_销帐
(
  No_In         门诊费用记录.No%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  序号_In       Varchar2 := Null,
  退费时间_In   门诊费用记录.登记时间%Type := Null,
  退费摘要_In   门诊费用记录.摘要%Type := Null,
  结帐id_In     病人预交记录.结帐id%Type := Null,
  回收票据_In   Number := 0
) As
  --功能：删除一张门诊收费单据 
  --参数： 
  --        序号_IN           =要退费的项目序号,格式为"1,3,5,6...",缺省NULL表示退"未退的"所有行。 
  --        回收票据_In       =0:全退或最后一次全退时,收回票据。 
  --                           1:部份退费不处理票据,通过重打调用单独处理。 
  --该游标为要退费单据的所有原始记录 

  --医保全退但某种结算退现金从而产生了新的误差时,排开此处的误差处理,执行完本过程后,界面程序中单独处理新误差 
  Cursor c_Bill Is
    Select a.Id, a.No, a.附加标志, a.收费细目id, a.序号, a.价格父号, a.执行状态, a.收费类别, a.付数, a.数次, a.医嘱序号, j.诊疗类别, m.跟踪在用,
           Nvl(a.附加标志, 0) As 误差
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.No = No_In And a.记录性质 = 1 And a.记录状态 In (1, 3) And a.收费细目id + 0 = m.材料id(+)
    Order By a.收费细目id, a.序号;

  --:不管原始单据误差,都应该根据当前退费产生的误差项进行处理
  -- Decode(Sign(误差_In), 0, 999, 9)

  --该游标用于处理药品库存可用数量 
  --不要管费用的执行状态,因为先于此步处理 
  Cursor c_Stock Is
    Select ID, 药品id, 库房id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
    From 药品收发记录
    Where NO = No_In And 单据 In (8, 24) --@@@ 
          And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 门诊费用记录
                   Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And 收费类别 In ('4', '5', '6', '7') --@@@ 
                         And (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
    Order By 药品id;

  --该游标用于处理未发药品记录 
  Cursor c_Spare Is
    Select NO, 库房id, 单据 From 未发药品记录 Where NO = No_In And 单据 In (8, 24); --@@@ 

  --该光标用于处理人员缴款余额中退的不同结算方式的金额 

  n_结帐id 门诊费用记录.结帐id%Type;
  n_打印id 票据打印内容.Id%Type;

  --部分退费计算变量 
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;
  n_备货卫材 Number;
  n_准退数量 Number;
  n_退费次数 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;
  n_总金额   Number;
  n_正常退费 Number; --是否第一次退费且全部退费,在每行退费过程中判断得到。 
  n_组id     财务缴款分组.Id%Type;

  l_费用id   t_Numlist := t_Numlist();
  l_药品收发 t_Numlist := t_Numlist();
  l_使用id   t_Numlist := t_Numlist();

  l_序号     t_Numlist := t_Numlist();
  l_执行状态 t_Numlist := t_Numlist();

  n_Dec   Number;
  d_Date  Date;
  n_Count Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_启用模式     Number(3);
  v_Para         Varchar2(1000);
  n_医属执行计价 Number;

Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  --是否已经全部完全执行(只是该单据整张单据的检查) 
  Select Nvl(Count(*), 0)
  Into n_Count
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查) 
  --执行状态在原始记录上判断 
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量, 结帐id
                From 门诊费用记录
                Where NO = No_In And Mod(记录性质, 10) = 1 And Nvl(附加标志, 0) <> 9 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号), 结帐id)
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以退费的费用！';
    Raise Err_Item;
  End If;

  --确定是否在医嘱执行计价中存在数据,如果存在数据,则根据医嘱执行计价进行退费,否则按旧方式进行处理
  Select Count(1)
  Into n_医属执行计价
  From 门诊费用记录 A, 医嘱执行计价 B
  Where a.医嘱序号 = b.医嘱id And Mod(a.记录性质, 10) = 1 And a.No = No_In And a.记录状态 In (1, 3) And Rownum = 1;

  --------------------------------------------------------------------------------- 
  --公用变量 
  If 退费时间_In Is Not Null Then
    d_Date := 退费时间_In;
  Else
    Select Sysdate Into d_Date From Dual;
  End If;

  If 结帐id_In Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  Else
    n_结帐id := 结帐id_In;
  End If;

  --金额小数位数 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --循环处理每行费用(收入项目行) 
  n_总金额 := 0;
  For r_Bill In c_Bill Loop
    If Instr(',' || 序号_In || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or 序号_In Is Null Then
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        --求剩余数量,剩余应收,剩余实收 
        Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
        Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
        From 门诊费用记录
        Where NO = No_In And Mod(记录性质, 10) = 1 And 序号 = r_Bill.序号;
      
        If n_剩余数量 = 0 Then
          If 序号_In Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部退费！';
            Raise Err_Item;
          End If;
        Else
          --准退数量(非药品项目为剩余数量,原始数量) 
          If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
            --@@@ 
            --非药品部分(以具体医嘱执行为准进行检查) 
            --: 1.存在医嘱发送的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血) 
            --: 2.不存在医嘱的,则以剩余数量为准 
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null Then
              If n_医属执行计价 = 1 Then
                Select Decode(Sign(Sum(数量)), -1, 0, Sum(数量)), Count(*)
                Into n_准退数量, n_Count
                From (Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, Max(a.医嘱序号) As 医嘱id, Max(a.收费细目id) As 收费细目id,
                              Sum(Nvl(a.付数, 1) * Nvl(a.数次, 1)) As 数量,
                              Sum(Decode(a.记录状态, 2, 0, Nvl(a.付数, 1) * Nvl(a.数次, 1))) As 原始数量
                       From 门诊费用记录 A, 病人医嘱记录 M
                       Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             Instr('5,6,7', a.收费类别) = 0 And a.No = No_In And a.序号 = r_Bill.序号 And Mod(a.记录性质, 10) = 1 And
                             a.记录状态 In (1, 2, 3) And a.价格父号 Is Null
                       Group By a.序号
                       Union All
                       Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, -1 * b.数量 As 已执行, 0 原始数量
                       From 门诊费用记录 A, 医嘱执行计价 B, 病人医嘱记录 M
                       Where a.医嘱序号 = b.医嘱id And a.收费细目id = b.收费细目id + 0 And a.医嘱序号 = m.Id And
                             Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Instr('5,6,7', a.收费类别) = 0 And
                             (Exists
                              (Select 1
                               From 病人医嘱执行
                               Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And b.要求时间 = 要求时间 And Nvl(执行结果, 0) = 1) Or Exists
                              (Select 1
                               From 病人医嘱发送
                               Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And Nvl(执行状态, 0) = 1)) And Not Exists
                        (Select 1
                              From 病人医嘱附费
                              Where a.医嘱序号 = 医嘱id And a.No = NO And Mod(a.记录性质, 10) = 记录性质) And a.No = No_In And
                             a.序号 = r_Bill.序号 And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) 　and a.价格父号 Is Null) Q1
                Where Not Exists (Select 1
                       From 药品收发记录
                       Where 费用id = Q1.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) Having
                 Max(ID) <> 0;
              Else
              
                Select Nvl(Sum(数量), 0), Count(*)
                Into n_准退数量, n_Count
                From (Select a.医嘱id, a.收费细目id, Nvl(a.数量, 1) * Nvl(b.发送数次, 1) As 数量
                       From 病人医嘱计价 A, 病人医嘱发送 B, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = b.医嘱id And a.医嘱id = m.Id And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And
                             a.收费细目id = j.收费细目id And j.No = No_In And Mod(j.记录性质, 10) = 1 And j.序号 = r_Bill.序号 And
                             j.记录状态 In (1, 3) And j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             Exists
                        (Select 1
                              From 病人医嘱计价 A
                              Where a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) = 0) And Not Exists
                        (Select 1
                              From 药品收发记录
                              Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0)
                       Union All
                       Select a.医嘱id, a.收费细目id, -1 * Nvl(a.数量, 1) * Nvl(c.本次数次, 1) As 数量
                       From 病人医嘱计价 A, 病人医嘱发送 B, 病人医嘱执行 C, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = b.医嘱id And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And a.医嘱id = m.Id And
                             Nvl(c.执行结果, 1) = 1 And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And
                             j.No = No_In And Mod(j.记录性质, 10) = 1 And Nvl(a.收费方式, 0) = 0 And j.序号 = r_Bill.序号 And
                             j.记录状态 In (1, 3) And j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             Not Exists
                        (Select 1
                              From 药品收发记录
                              Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) And Not Exists
                        (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)
                       Union All
                       Select a.医嘱序号 As 医嘱id, a.收费细目id, Nvl(a.付数, 1) * a.数次 As 数量
                       From 门诊费用记录 A, 病人医嘱记录 M
                       Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And a.No = No_In And
                             Mod(a.记录性质, 10) = 1 And a.序号 = r_Bill.序号 And a.记录状态 = 2 And a.价格父号 Is Null And Not Exists
                        (Select 1 From 药品收发记录 Where NO = No_In And 单据 In (8, 24) And 药品id = a.收费细目id));
              End If;
            End If;
            If Nvl(n_Count, 0) <> 0 And n_准退数量 = 0 Then
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已执行,不允许退费！';
              Raise Err_Item;
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_准退数量 := n_剩余数量;
            End If;
          
          Else
            Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0), Count(*)
            Into n_准退数量, n_Count
            From 药品收发记录
            Where NO = No_In And 单据 In (8, 24) And Mod(记录状态, 3) = 1 --@@@ 
                  And 审核人 Is Null And 费用id = r_Bill.Id;
          
            --有剩余数量无准退数量的有两种情况： 
            --1.不跟踪在用的卫材无对应的收发记录,这时使用剩余数量 
            --2.并发操作,此时已发药或发料 
            If n_准退数量 = 0 Then
              If r_Bill.收费类别 = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发料,须退料后再退费！';
                  Raise Err_Item;
                Else
                  n_准退数量 := n_剩余数量;
                End If;
              Else
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发药,须退药后再退费！';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          If n_准退数量 > n_剩余数量 Then
            v_Err_Msg := '单据[' || No_In || '] 中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用的退费数量(' || n_准退数量 ||
                         ')大于了剩余数量(' || n_剩余数量 || ')，不允许退费！';
            Raise Err_Item;
          End If;
          If n_准退数量 < 0 Then
            v_Err_Msg := '单据[' || No_In || '] 中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用的退费数量(' || n_准退数量 ||
                         ')小于了零，不允许退费！';
            Raise Err_Item;
          End If;
        
          --是否部分退费 
          If r_Bill.执行状态 = 2 Or n_准退数量 <> Nvl(r_Bill.付数, 1) * r_Bill.数次 Then
            n_正常退费 := 0;
          End If;
        
          --该笔项目第几次退费 
          Select Nvl(Max(Abs(执行状态)), 0) + 1
          Into n_退费次数
          From 门诊费用记录
          Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 2 And Nvl(执行状态, 0) < 0 And 序号 = r_Bill.序号;
        
          --金额=剩余金额*(准退数/剩余数) 
          n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), n_Dec);
          n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), n_Dec);
          n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), n_Dec);
          n_总金额   := n_总金额 + n_实收金额;
        
          --插入退费记录 
          Insert Into 门诊费用记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别,
             收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人,
             执行状态, 费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论,
             缴款组id)
            Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                   病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                   Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价,
                   -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, 1, 执行时间, 操作员编号_In, 操作员姓名_In,
                   发生时间, d_Date, n_结帐id, -1 * n_实收金额, 保险项目否, 保险大类id, -1 * n_统筹金额, Nvl(退费摘要_In, 摘要),
                   Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, 结论, n_组id
            From 门诊费用记录
            Where ID = r_Bill.Id;
        
          --标记原费用记录 
          l_序号.Extend;
          l_序号(l_序号.Count) := r_Bill.序号;
          l_执行状态.Extend;
          l_执行状态(l_执行状态.Count) := Case
                                    When Sign(n_准退数量 - n_剩余数量) = 0 Then
                                     0
                                    Else
                                     1
                                  End;
        
          --          Update 门诊费用记录 Set 记录状态 = 3 Where ID = r_Bill.Id;
        
        End If;
      Else
        If 序号_In Is Not Null Then
          v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能退费！';
          Raise Err_Item;
        End If;
        --情况:没限定行号,原始单据中包括已经完全执行的 
        n_正常退费 := 0;
      End If;
    Else
      n_正常退费 := 0; --未指定该笔,属于部分退费 
    End If;
  End Loop;
  --标记原费用记录 
  Forall I In 1 .. l_序号.Count
    Update 门诊费用记录
    Set 记录状态 = 3, 执行状态 = l_执行状态(I)
    Where Mod(记录性质, 10) = 1 And NO = No_In And 序号 = l_序号(I) And 记录状态 In (1, 3);

  l_序号.Delete;
  For c_结帐 In (Select Distinct b.结帐id
               From 门诊费用记录 A, 病人预交记录 B
               Where a.结帐id = b.结帐id And a.No = No_In And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And
                     Nvl(b.记录状态, 0) = 1) Loop
    l_序号.Extend;
    l_序号(l_序号.Count) := c_结帐.结帐id;
  End Loop;

  Forall I In 1 .. l_序号.Count
    Update 病人预交记录 Set 记录状态 = 3 Where 结帐id = l_序号(I) And Mod(记录性质, 10) <> 1;

  --------------------------------------------------------------------------------- 
  --退费票据回收(仅全退时才回退,部分退是在重打过程中回收) 
  If 回收票据_In = 1 Then
  
    --启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
    v_Para     := Nvl(zl_GetSysParameter('票据分配规则', 1121), '0||0;0;0,0;0');
    n_启用模式 := Zl_To_Number(Substr(v_Para, 1, 1));
    If n_启用模式 <> 0 Then
      --收回票据
      Select 使用id Bulk Collect
      Into l_使用id
      From (Select Distinct b.使用id From 票据打印明细 B Where b.No = No_In And Nvl(b.票种, 0) = 1);
    
      n_启用模式 := l_使用id.Count;
      If l_使用id.Count <> 0 Then
        --插入回收记录
        Forall I In 1 .. l_使用id.Count
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 操作员姓名_In, d_Date
            From 票据使用明细 A
            Where ID = l_使用id(I) And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 Where 号码 = a.号码 And 票种 = a.票种 And Nvl(性质, 0) <> 1);
      
        Forall I In 1 .. l_使用id.Count
          Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I) And Nvl(是否回收, 0) = 0;
      
      End If;
    End If;
    If n_启用模式 = 0 Then
      --获取单据最后一次的打印ID(可能是多张单据收费打印) 
      Begin
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 1 And b.No = No_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --可能以前没有打印,无收回 
      If n_打印id Is Not Null Then
        --a.多张单据循环调用时只能收回一次 
        Select Count(*) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
        If n_Count = 0 Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
            From 票据使用明细
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
        Else
          --b.部分退费多次收回时,最后一次全退收回要排开已收回的 
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
            From 票据使用明细 A
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And 打印id = n_打印id And 票种 = 1 And 性质 = 2);
        End If;
      End If;
    End If;
  End If;

  --------------------------------------------------------------------------------- 
  --卫生材料 
  For v_出库 In (Select ID, 药品id, 库房id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
               From 药品收发记录
               Where 单据 = 21 --@@@ 
                     And Mod(记录状态, 3) = 1 And 审核人 Is Null And
                     费用id In (Select ID
                              From 门诊费用记录
                              Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And 收费类别 = '4' --@@@ 
                                    And (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
               Order By 药品id) Loop
    --处理药品库存 
    If v_出库.库房id Is Not Null Then
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0)
      Where 库房id = v_出库.库房id And 药品id = v_出库.药品id And Nvl(批次, 0) = Nvl(v_出库.批次, 0) And 性质 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 药品库存
          (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码) --@@@ 
        Values
          (v_出库.库房id, v_出库.药品id, 1, v_出库.批次, v_出库.效期,
           Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0), v_出库.批号, v_出库.产地, v_出库.灭菌效期,
           v_出库.商品条码, v_出库.内部条码);
      End If;
    End If;
  
    l_费用id.Extend;
    l_费用id(l_费用id.Count) := v_出库.费用id;
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := v_出库.Id;
  End Loop;

  --药品相关内容 
  For r_Stock In c_Stock Loop
    --处理药品库存 
    If r_Stock.库房id Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_备货卫材
      From Table(l_费用id)
      Where Column_Value = r_Stock.费用id;
      If Nvl(n_备货卫材, 0) = 0 Then
      
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0)
        Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码) --@@@ 
          Values
            (r_Stock.库房id, r_Stock.药品id, 1, r_Stock.批次, r_Stock.效期,
             Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0), r_Stock.批号, r_Stock.产地,
             r_Stock.灭菌效期, r_Stock.商品条码, r_Stock.内部条码);
        End If;
      End If;
    End If;
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := r_Stock.Id;
  End Loop;

  --删除药品收发记录 
  Forall I In 1 .. l_药品收发.Count
    Delete From 药品收发记录 Where ID = l_药品收发(I) And 审核人 Is Null;

  --未发药品记录 
  For r_Spare In c_Spare Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From 药品收发记录
    Where NO = No_In And 单据 = r_Spare.单据 --@@@ 
          And Mod(记录状态, 3) = 1 And 审核人 Is Null And Nvl(库房id, 0) = Nvl(r_Spare.库房id, 0);
  
    If n_Count = 0 Then
      Delete From 未发药品记录
      Where 单据 = r_Spare.单据 --@@@ 
            And NO = No_In And Nvl(库房id, 0) = Nvl(r_Spare.库房id, 0);
    End If;
  End Loop;
  --医嘱处理 
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 3 And 医嘱序号 Is Not Null And
                     (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null)) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, 执行状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where Mod(记录性质, 10) = 1 And Nvl(附加标志, 0) <> 9 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, 执行状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      --更新计费状态为1:已计费 
      Update 病人医嘱发送 A Set 计费状态 = 1 Where NO = No_In And 医嘱id = c_医嘱.医嘱序号 And Mod(记录性质, 10) = 1;
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And Mod(记录性质, 10) = 1 And NO = No_In;
    Else
      --更新计费状态为2:部分退费 
      Update 病人医嘱发送 A Set 计费状态 = 2 Where NO = No_In And 医嘱id = c_医嘱.医嘱序号 And Mod(记录性质, 10) = 1;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_销帐;
/

--90389:冉俊明,2015-11-16,医嘱附费项目退费数量错误。
Create Or Replace Procedure Zl_门诊收费记录_Delete
(
  No_In           门诊费用记录.No%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  医保结算方式_In Varchar2 := Null,
  序号_In         Varchar2 := Null,
  结算方式_In     病人预交记录.结算方式%Type := Null,
  误差_In         门诊费用记录.实收金额%Type := 0,
  退费时间_In     门诊费用记录.登记时间%Type := Null,
  回收票据_In     Number := 0,
  退费摘要_In     门诊费用记录.摘要%Type := Null,
  校对标志_In     Number := 0,
  结帐id_In       病人预交记录.结帐id%Type := Null,
  结算序号_In     病人预交记录.结算序号%Type := Null,
  一卡通结算_In   Varchar2 := Null,
  退款操作_In     Number := 0,
  多单据全退_In   Number := 0
) As
  --功能：删除一张门诊收费单据 
  --参数： 
  --        医保结算方式_IN   =医保退费时,不支持结算作废的结算方式,如果为空表示非医保退费或医保退费全部结算允许作废。 
  --        序号_IN           =要退费的项目序号,格式为"1,3,5,6...",缺省NULL表示退"未退的"所有行。 
  --        结算方式_IN       =当为部分退费时,退费金额的结算方式。 
  --        误差_IN           =指退费时新产生的误差金额,部份退费或医保全退但某种结算退现金时才会产生新的误差。 
  --                           此时传入仅用于计算本次退费的结算金额,误差费用记录的处理在本过程执行完后调用Zl_门诊收费误差_Insert产生 
  --        回收票据_In       =0:单张全退或多张一起全退时收回票据,注意,多张单据退费循环调本过程时只收回一次。 
  --                           1:部份退费不处理票据,通过重打调用单独处理。 
  --        校对标志_IN:0-不需要较对;1-需较对(不处理人员缴款余额,不回收票据,不处理预交余额) 
  --        退款操作_In:1-进行部分退(将退款方式退到指定的结算方式<结算方式_In>中,0-不指定退款方式) 
  --        多单据全退_IN=1-多单据全退(多张单据全退,原样退);0-非原样退
  --该游标为要退费单据的所有原始记录 

  --医保全退但某种结算退现金从而产生了新的误差时,排开此处的误差处理,执行完本过程后,界面程序中单独处理新误差 
  Cursor c_Bill Is
    Select a.Id, a.No, a.附加标志, a.收费细目id, a.序号, a.价格父号, a.执行状态, a.收费类别, a.付数, a.数次, a.医嘱序号, j.诊疗类别, m.跟踪在用,
           Nvl(a.附加标志, 0) As 误差
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.No = No_In And a.记录性质 = 1 And a.记录状态 In (1, 3) And a.收费细目id + 0 = m.材料id(+) And
          Nvl(a.附加标志, 0) <> Decode(多单据全退_In, 1, 999, 9)
    Order By a.收费细目id, a.序号;
  --:不管原始单据误差,都应该根据当前退费产生的误差项进行处理
  -- Decode(Sign(误差_In), 0, 999, 9)

  --该游标用于处理药品库存可用数量 
  --不要管费用的执行状态,因为先于此步处理 
  Cursor c_Stock Is
    Select ID, 药品id, 库房id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
    From 药品收发记录
    Where NO = No_In And 单据 In (8, 24) --@@@ 
          And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 门诊费用记录
                   Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And 收费类别 In ('4', '5', '6', '7') --@@@ 
                         And (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
    Order By 药品id;

  --该游标用于处理未发药品记录 
  Cursor c_Spare Is
    Select NO, 库房id, 单据 From 未发药品记录 Where NO = No_In And 单据 In (8, 24); --@@@ 

  --该光标用于处理人员缴款余额中退的不同结算方式的金额 
  Cursor c_Money(冲销id_In 病人预交记录.结帐id%Type) Is
    Select 结算方式, 冲预交
    From 病人预交记录
    Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = 冲销id_In And 结算方式 Is Not Null And Nvl(冲预交, 0) <> 0 And Nvl(校对标志, 0) = 0;

  --该游标用于查找收费时使用过的冲预交款记录 
  Cursor c_Deposit(V结帐id 病人预交记录.结帐id%Type) Is
    Select ID, 冲预交 As 金额, 预交类别
    From 病人预交记录
    Where 记录性质 In (1, 11) And 记录状态 In (1, 3) And 结帐id = V结帐id And Nvl(冲预交, 0) <> 0
    Order By ID Desc;

  n_病人id   病人信息.病人id%Type;
  n_结帐id   门诊费用记录.结帐id%Type;
  n_结算序号 病人预交记录.结算序号%Type;
  n_打印id   票据打印内容.Id%Type;

  n_已退金额 病人预交记录.冲预交%Type;
  n_预交金额 病人预交记录.冲预交%Type;
  n_返回值   病人预交记录.冲预交%Type;
  n_原误差费 门诊费用记录.实收金额%Type;
  --部分退费计算变量 
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;
  n_备货卫材 Number;
  n_准退数量 Number;
  n_退费次数 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;
  n_总金额   Number;
  n_费用状态 门诊费用记录.费用状态%Type;
  n_正常退费 Number; --是否第一次退费且全部退费,在每行退费过程中判断得到。 
  n_组id     财务缴款分组.Id%Type;

  v_退费结算 结算方式.名称%Type;
  v_结算内容 Varchar2(500);
  n_部分退   Number(2);

  l_费用id   t_Numlist := t_Numlist();
  l_药品收发 t_Numlist := t_Numlist();
  l_使用id   t_Numlist := t_Numlist();
  n_Dec      Number;
  d_Date     Date;
  n_Count    Number;
  n_原结帐id Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_启用模式     Number(3);
  v_Para         Varchar2(1000);
  n_医属执行计价 Number;

  Procedure Zl_Square_Update
  (
    原结帐id_In 病人预交记录.结帐id%Type,
    现结帐id_In 病人预交记录.结帐id%Type,
    缴款组id_In 病人预交记录.缴款组id%Type,
    退款时间_In 病人预交记录.收款时间%Type,
    结算序号_In 病人预交记录.结算序号%Type,
    结算内容_In Varchar2 := Null
  ) As
    n_记录状态 病人卡结算记录.记录状态%Type;
    n_预交id   病人预交记录.Id%Type;
    v_卡号     病人卡结算记录.卡号%Type;
    n_存在卡片 Number;
    d_停用日期 消费卡目录.停用日期%Type;
    n_最大序号 病人卡结算记录.序号%Type;
    n_序号     病人卡结算记录.序号%Type;
    n_余额     消费卡目录.余额%Type;
    n_接口编号 病人卡结算记录.接口编号%Type;
    d_回收时间 消费卡目录.回收时间%Type;
    n_Id       病人预交记录.Id%Type;
  Begin
    n_预交id := 0;
    --处理消费卡,结算卡在上面就已经处理了 
    For v_校对 In (Select a.Id As 预交id, c.消费卡id, c.结算金额, c.接口编号, c.卡号, c.序号, c.Id
                 From 病人预交记录 A, 病人卡结算对照 B, 病人卡结算记录 C
                 Where a.Id = b.预交id And b.卡结算id = c.Id And a.记录性质 = 3 And a.记录状态 = 1 And
                       Instr(Nvl(结算内容_In, '_LXH'), ',' || a.结算方式 || ',') = 0 And a.结帐id = 原结帐id_In) Loop
    
      If Nvl(v_校对.消费卡id, 0) <> 0 Then
        Select Max(记录状态)
        Into n_记录状态
        From 病人卡结算记录
        Where 接口编号 = v_校对.接口编号 And 消费卡id = Nvl(v_校对.消费卡id, 0) And 卡号 = v_校对.卡号 And Nvl(序号, 0) = Nvl(v_校对.序号, 0);
      Else
        Select Max(记录状态)
        Into n_记录状态
        From 病人卡结算记录
        Where 接口编号 = v_校对.接口编号 And 消费卡id Is Null And 卡号 = v_校对.卡号 And Nvl(序号, 0) = Nvl(v_校对.序号, 0);
      End If;
    
      If n_记录状态 = 1 Then
        n_记录状态 := 2;
      Else
        n_记录状态 := n_记录状态 + 2;
      End If;
      --多条时,只更新一条
      If n_预交id = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id,
           预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
          Select n_预交id, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退款时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
                 -1 * 冲预交, 现结帐id_In, 缴款组id_In, 预交类别, 卡类别id, Nvl(结算卡序号, v_校对.接口编号), 卡号, 交易流水号, 交易说明, 合作单位,
                 Decode(Nvl(校对标志_In, 0), 0, 0, Decode(Nvl(v_校对.消费卡id, 0), 0, 1, 2)), 结算序号_In, 3
          From 病人预交记录 A
          Where ID = v_校对.预交id;
      End If;
    
      If Nvl(v_校对.消费卡id, 0) <> 0 Then
        --消费卡,直接退回卡数据中 
        Begin
          Select 卡号, 1, 停用日期, (Select Max(序号) From 消费卡目录 B Where a.卡号 = b.卡号 And a.接口编号 = b.接口编号), 序号, 余额, 接口编号, 回收时间
          Into v_卡号, n_存在卡片, d_停用日期, n_最大序号, n_序号, n_余额, n_接口编号, d_回收时间
          From 消费卡目录 A
          Where ID = v_校对.消费卡id;
        Exception
          When Others Then
            n_存在卡片 := 0;
        End;
      
        --取消停用 
        If n_存在卡片 = 0 Then
          v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡被他人删除，不能再启用该卡片,请检查！';
          Raise Err_Item;
        End If;
        If Nvl(n_序号, 0) < Nvl(n_最大序号, 0) Then
          v_Err_Msg := '不能启用历史发卡记录(卡号为"' || v_卡号 || '"),请检查！';
          Raise Err_Item;
        End If;
        If Nvl(d_停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡已经被他人停用，不能再进行退费,请检查！';
          Raise Err_Item;
        End If;
      
        If d_回收时间 < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡已经回收，不能退费,请检查！';
          Raise Err_Item;
        End If;
        Update 消费卡目录 Set 余额 = Nvl(余额, 0) + v_校对.结算金额 Where ID = Nvl(v_校对.消费卡id, 0);
      End If;
    
      Select 病人卡结算记录_Id.Nextval Into n_Id From Dual;
      Insert Into 病人卡结算记录
        (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
        Select n_Id, 接口编号, 消费卡id, 序号, n_记录状态, 结算方式, -1 * v_校对.结算金额, 卡号, 交易流水号, 交易时间, 备注,
               Decode(消费卡id, Null, 0, 0, 0, 1) As 标志
        From 病人卡结算记录
        Where ID = v_校对.Id;
      Insert Into 病人卡结算对照 (预交id, 卡结算id) Values (n_预交id, n_Id);
    
      If n_记录状态 <> 2 And n_记录状态 <> 1 Then
        Update 病人卡结算记录 Set 记录状态 = 3 Where ID = v_校对.Id;
      End If;
    End Loop;
  End;

Begin
  n_组id   := Zl_Get组id(操作员姓名_In);
  n_部分退 := 0;
  --是否已经全部完全执行(只是该单据整张单据的检查) 
  Select Nvl(Count(*), 0)
  Into n_Count
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查) 
  --执行状态在原始记录上判断 
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 门诊费用记录
                Where NO = No_In And 记录性质 = 1 And Nvl(附加标志, 0) <> 9 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以退费的费用！';
    Raise Err_Item;
  End If;
  --确定是否在医嘱执行计价中存在数据,如果存在数据,则根据医嘱执行计价进行退费,否则按旧方式进行处理
  Select Count(1)
  Into n_医属执行计价
  From 门诊费用记录 A, 医嘱执行计价 B
  Where a.医嘱序号 = b.医嘱id And a.记录性质 = 1 And a.No = No_In And a.记录状态 In (1, 3) And Rownum = 1;

  --------------------------------------------------------------------------------- 
  --公用变量 
  If 退费时间_In Is Not Null Then
    d_Date := 退费时间_In;
  Else
    Select Sysdate Into d_Date From Dual;
  End If;
  If 结帐id_In Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  Else
    n_结帐id := 结帐id_In;
  End If;
  n_结算序号 := 结算序号_In;
  If n_结算序号 Is Null Then
    n_结算序号 := 结帐id_In;
  End If;
  --金额小数位数 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --获取结算方式名称 
  v_退费结算 := 结算方式_In;
  If v_退费结算 Is Null Then
    Begin
      Select 名称 Into v_退费结算 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_退费结算 := '现金';
    End;
  End If;
  --循环处理每行费用(收入项目行) 
  n_总金额   := 0;
  n_正常退费 := 1;
  For r_Bill In c_Bill Loop
    If Instr(',' || 序号_In || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or 序号_In Is Null Then
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        --求剩余数量,剩余应收,剩余实收 
        Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
        Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 1 And 序号 = r_Bill.序号;
      
        If n_剩余数量 = 0 Then
          If 序号_In Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部退费！';
            Raise Err_Item;
          End If;
          --情况：未限定行号,原始单据中的该笔已经全部退费(执行状态=0的一种可能) 
          n_正常退费 := 0;
        Else
          --准退数量(非药品项目为剩余数量,原始数量) 
          If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
            --@@@ 
            --非药品部分(以具体医嘱执行为准进行检查) 
            --: 1.存在医嘱发送的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血) 
            --: 2.不存在医嘱的,则以剩余数量为准 
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null Then
              If n_医属执行计价 = 1 Then
                Select Decode(Sign(Sum(数量)), -1, 0, Sum(数量)), Count(*)
                Into n_准退数量, n_Count
                From (Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, Max(a.医嘱序号) As 医嘱id, Max(a.收费细目id) As 收费细目id,
                              Sum(Nvl(a.付数, 1) * Nvl(a.数次, 1)) As 数量,
                              Sum(Decode(a.记录状态, 2, 0, Nvl(a.付数, 1) * Nvl(a.数次, 1))) As 原始数量
                       From 门诊费用记录 A, 病人医嘱记录 M
                       Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             Instr('5,6,7', a.收费类别) = 0 And a.No = No_In And a.序号 = r_Bill.序号 And a.记录性质 = 1 And
                             a.记录状态 In (1, 2, 3) And a.价格父号 Is Null
                       Group By a.序号
                       Union All
                       Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, -1 * b.数量 As 已执行, 0 原始数量
                       From 门诊费用记录 A, 医嘱执行计价 B, 病人医嘱记录 M
                       Where a.医嘱序号 = b.医嘱id And a.收费细目id = b.收费细目id + 0 And a.医嘱序号 = m.Id And
                             Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Instr('5,6,7', a.收费类别) = 0 And
                             (Exists
                              (Select 1
                               From 病人医嘱执行
                               Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And b.要求时间 = 要求时间 And Nvl(执行结果, 0) = 1) Or Exists
                              (Select 1
                               From 病人医嘱发送
                               Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And Nvl(执行状态, 0) = 1)) And Not Exists
                        (Select 1
                              From 病人医嘱附费
                              Where a.医嘱序号 = 医嘱id And a.No = NO And Mod(a.记录性质, 10) = 记录性质) And a.No = No_In And
                             a.序号 = r_Bill.序号 And a.记录性质 = 1 And a.记录状态 In (1, 3) 　and a.价格父号 Is Null) Q1
                Where Not Exists (Select 1
                       From 药品收发记录
                       Where 费用id = Q1.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) Having
                 Max(ID) <> 0;
              Else
              
                Select Nvl(Sum(数量), 0), Count(*)
                Into n_准退数量, n_Count
                From (Select a.医嘱id, a.收费细目id, Nvl(a.数量, 1) * Nvl(b.发送数次, 1) As 数量
                       From 病人医嘱计价 A, 病人医嘱发送 B, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = b.医嘱id And a.医嘱id = m.Id And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And
                             a.收费细目id = j.收费细目id And j.No = No_In And j.记录性质 = 1 And j.序号 = r_Bill.序号 And
                             j.记录状态 In (1, 3) And j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             Exists
                        (Select 1
                              From 病人医嘱计价 A
                              Where a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) = 0) And Not Exists
                        (Select 1
                              From 药品收发记录
                              Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0)
                       Union All
                       Select a.医嘱id, a.收费细目id, -1 * Nvl(a.数量, 1) * Nvl(c.本次数次, 1) As 数量
                       From 病人医嘱计价 A, 病人医嘱发送 B, 病人医嘱执行 C, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = b.医嘱id And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And a.医嘱id = m.Id And
                             Nvl(c.执行结果, 1) = 1 And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And
                             j.No = No_In And j.记录性质 = 1 And Nvl(a.收费方式, 0) = 0 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And
                             j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Not Exists
                        (Select 1
                              From 药品收发记录
                              Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) And Not Exists
                        (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)
                       Union All
                       Select a.医嘱序号 As 医嘱id, a.收费细目id, Nvl(a.付数, 1) * a.数次 As 数量
                       From 门诊费用记录 A, 病人医嘱记录 M
                       Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And a.No = No_In And
                             a.记录性质 = 1 And a.序号 = r_Bill.序号 And a.记录状态 = 2 And a.价格父号 Is Null And Not Exists
                        (Select 1 From 药品收发记录 Where NO = No_In And 单据 In (8, 24) And 药品id = a.收费细目id));
              End If;
            End If;
            If Nvl(n_Count, 0) <> 0 And n_准退数量 = 0 Then
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已执行,不允许退费！';
              Raise Err_Item;
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_准退数量 := n_剩余数量;
            End If;
          
          Else
            Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0), Count(*)
            Into n_准退数量, n_Count
            From 药品收发记录
            Where NO = No_In And 单据 In (8, 24) And Mod(记录状态, 3) = 1 --@@@ 
                  And 审核人 Is Null And 费用id = r_Bill.Id;
          
            --有剩余数量无准退数量的有两种情况： 
            --1.不跟踪在用的卫材无对应的收发记录,这时使用剩余数量 
            --2.并发操作,此时已发药或发料 
            If n_准退数量 = 0 Then
              If r_Bill.收费类别 = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发料,须退料后再退费！';
                  Raise Err_Item;
                Else
                  n_准退数量 := n_剩余数量;
                End If;
              Else
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发药,须退药后再退费！';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --是否部分退费 
          If r_Bill.执行状态 = 2 Or n_准退数量 <> Nvl(r_Bill.付数, 1) * r_Bill.数次 Then
            n_正常退费 := 0;
          End If;
        
          --处理门诊费用记录 
          n_费用状态 := 0;
          --该笔项目第几次退费 
          If Nvl(校对标志_In, 0) <> 0 Then
            n_退费次数 := -9; --先标明,固定为9 
            n_费用状态 := 1;
          Else
            Select Nvl(Max(Abs(执行状态)), 0) + 1
            Into n_退费次数
            From 门诊费用记录
            Where NO = No_In And 记录性质 = 1 And 记录状态 = 2 And Nvl(执行状态, 0) < 0 And 序号 = r_Bill.序号;
          End If;
        
          --金额=剩余金额*(准退数/剩余数) 
          If Nvl(r_Bill.误差, 0) = 9 Then
            --误差可以超过设置的小数位(比如:医保结算超过小数位后,误差就可能超过小数位
            n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), 5);
            n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), 5);
            n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), 5);
          Else
            n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), n_Dec);
            n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), n_Dec);
            n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), n_Dec);
          End If;
          n_总金额 := n_总金额 + n_实收金额;
        
          --插入退费记录 
          Insert Into 门诊费用记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别,
             收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人,
             执行状态, 费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论,
             缴款组id)
            Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                   病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                   Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价,
                   -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, n_费用状态, 执行时间, 操作员编号_In,
                   操作员姓名_In, 发生时间, d_Date, n_结帐id, -1 * n_实收金额, 保险项目否, 保险大类id, -1 * n_统筹金额, Nvl(退费摘要_In, 摘要),
                   Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, 结论, n_组id
            From 门诊费用记录
            Where ID = r_Bill.Id;
        
          --标记原费用记录 
          --执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1,异常收费单,还是标明9 
          Update 门诊费用记录
          Set 记录状态 = 3, 执行状态 = Decode(Nvl(执行状态, 0), 9, 9, Decode(Sign(n_准退数量 - n_剩余数量), 0, 0, 1))
          Where ID = r_Bill.Id;
        End If;
      Else
        If 序号_In Is Not Null Then
          v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能退费！';
          Raise Err_Item;
        End If;
        --情况:没限定行号,原始单据中包括已经完全执行的 
        n_正常退费 := 0;
      End If;
    Else
      n_正常退费 := 0; --未指定该笔,属于部分退费 
    End If;
  End Loop;
  --------------------------------------------------------------------------------- 
  --处理病人预交记录 

  --原单据的结帐ID 
  Select 结帐id, 病人id
  Into n_原结帐id, n_病人id
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Rownum = 1;

  If n_正常退费 = 1 And Nvl(退款操作_In, 0) = 0 Then
    --单据第一次退费且全部退完 
    --冲预交部分记录 
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date,
             操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
             Decode(校对标志_In, 1, 2, 校对标志_In), n_结算序号, 3
      From 病人预交记录
      Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
    If Nvl(校对标志_In, 0) = 0 Then
      --处理病人预交余额 
      For v_预交 In (Select 预交类别, Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交金额
                   From 病人预交记录
                   Where 记录性质 In (1, 11) And 结帐id = n_原结帐id
                   Group By 预交类别
                   Having Sum(Nvl(冲预交, 0)) <> 0) Loop
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_预交.预交金额, 0)
        Where 病人id = n_病人id And 性质 = 1 And 类型 = Nvl(v_预交.预交类别, 2)
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 类型, 预交余额, 性质)
          Values
            (n_病人id, Nvl(v_预交.预交类别, 2), Nvl(v_预交.预交金额, 0), 1);
          n_返回值 := n_预交金额;
        End If;
        If n_返回值 = 0 Then
          Delete From 病人余额 Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      End Loop;
    End If;
    --非医保全退,和医保所有结算方式都允许回退,原样退回(冲预交在前面已处理) 
    If 医保结算方式_In Is Null Then
      v_结算内容 := ',' || Nvl(一卡通结算_In, '-Lxh') || ',' || Nvl(一卡通结算_In, 'Lxh') || ',';
    
      --一卡通或消费卡或银行卡的相关数据需要特殊处理,需要最后较对. 
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
         卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, d_Date, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
               -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
               Case
                 When Nvl(卡类别id, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(结算卡序号, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(q.预交id, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(j.名称, '-') <> '-' Then
                  Decode(校对标志_In, 1, 1, 0)
                 Else
                  Decode(校对标志_In, 1, 2, 0)
               End As 校对标志, n_结算序号, 3
        From 病人预交记录 A, (Select 名称 From 结算方式 Where 性质 In (3, 4)) J,
             (Select m.Id As 预交id
               From 病人预交记录 M, 一卡通目录 C
               Where m.结帐id = n_原结帐id And m.结算方式 = c.结算方式 And m.记录性质 = 3 And m.记录状态 = 1) Q
        Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id = n_原结帐id And a.Id = q.预交id(+) And a.结算方式 = j.名称(+) And
              Instr(v_结算内容, ',' || 结算方式 || ',') = 0 And
              (Not Exists (Select 1 From 病人卡结算对照 Where a.Id = 预交id) Or Nvl(结算卡序号, 0) = 0);
    
      --处理消费卡,结算卡在上面就已经处理了 
      Zl_Square_Update(n_原结帐id, n_结帐id, n_组id, d_Date, n_结算序号, v_结算内容);
      --b.余下的就是三方接口支持的退现了,不允许作废的结算方式,加上到指定的结算方式上,加上误差(因为界面程序会在这之后退误差) 
      If 一卡通结算_In Is Not Null Then
        Begin
          Select -1 * Nvl(Sum(冲预交), 0) Into n_已退金额 From 病人预交记录 Where 结帐id = n_结帐id;
        Exception
          When Others Then
            n_已退金额 := 0;
        End;
      
        If (n_总金额 - n_已退金额) <> 0 Then
          --此时的总金额还没有包含误差,因为界面程序中在调用本过程后才产生误差费用记录 
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号,
             交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
            Select 病人预交记录_Id.Nextval, 3, NO, 2, 病人id, 主页id, '门诊退费结算', v_退费结算, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * (n_总金额 - n_已退金额), n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
                   Decode(校对标志_In, 1, 2, 0), n_结算序号, 3
            From 病人预交记录
            Where 记录性质 = 3 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum = 1;
          n_部分退 := 1;
        End If;
      End If;
      --医保按允许作废的结算方式退,不允许的,退到指定的结算方式上 
      --需要处理误差金额
    Else
      --a.原样退回 
      v_结算内容 := ',' || 医保结算方式_In || ',' || Nvl(一卡通结算_In, '-Lxh') || ',' || v_退费结算 || ',';
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, d_Date, 操作员编号_In, 操作员姓名_In, -1 * 冲预交, n_结帐id,
               n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
               
               Case
                 When Nvl(卡类别id, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(结算卡序号, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(q.预交id, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(j.名称, '-') <> '-' Then
                  Decode(校对标志_In, 1, 1, 0)
                 Else
                  Decode(校对标志_In, 1, 2, 0)
               End As 校对标志, n_结算序号, 3
        From 病人预交记录 A, (Select 名称 From 结算方式 Where 性质 In (3, 4)) J,
             (Select m.Id As 预交id
               From 病人预交记录 M, 一卡通目录 C
               Where m.结帐id = n_原结帐id And m.结算方式 = c.结算方式 And m.记录性质 = 3 And m.记录状态 = 1) Q
        Where a.记录性质 = 3 And a.记录状态 = 1 And a.结算方式 = j.名称(+) And a.结帐id = n_原结帐id And
              Instr(v_结算内容, ',' || a.结算方式 || ',') = 0 And a.Id = q.预交id(+) And
              (Not Exists (Select 1 From 病人卡结算对照 Where a.Id = 预交id) Or Nvl(结算卡序号, 0) = 0);
    
      --处理消费卡,结算卡在上面就已经处理了 
      Zl_Square_Update(n_原结帐id, n_结帐id, n_组id, d_Date, n_结算序号, v_结算内容);
    
      --b.余下的就是医保不允许作废的结算方式,加上到指定的结算方式上,加上误差(因为界面程序会在这之后退误差) 
      Begin
        Select -1 * Nvl(Sum(冲预交), 0) Into n_已退金额 From 病人预交记录 Where 结帐id = n_结帐id;
      Exception
        When Others Then
          n_已退金额 := 0;
      End;
    
      If (n_总金额 - n_已退金额) <> 0 Then
        --此时的总金额还没有包含误差,因为界面程序中在调用本过程后才产生误差费用记录 
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号,
           交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
          Select 病人预交记录_Id.Nextval, 3, NO, 2, 病人id, 主页id, Decode(一卡通结算_In, Null, '门诊医保接口退费', '门诊医保接口和三方接口退费'), v_退费结算,
                 d_Date, 操作员编号_In, 操作员姓名_In, -1 * (n_总金额 - n_已退金额), n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
                 合作单位, Decode(校对标志_In, 1, 2, 0), n_结算序号, 3
          From 病人预交记录
          Where 记录性质 = 3 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum = 1;
        n_部分退 := 1;
      End If;
    
    End If;
  Else
    ------------------------------------------------- 
    --部分退费直接退为指定结算方式 
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
       结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, 3, No_In, 2, 病人id, 主页id, '部分退费结算', v_退费结算, d_Date, Null, Null, Null, 操作员编号_In, 操作员姓名_In,
             -1 * (n_总金额 + Nvl(误差_In, 0)), n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
             Decode(校对标志, 1, 2, 0), n_结算序号, 3
      From 病人预交记录
      Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id And Rownum = 1;
  
    --如果收费时只使用了预交款,则要退预交,并且可能有多笔冲预交 
    If Sql%RowCount = 0 Then
      n_预交金额 := n_总金额 + Nvl(误差_In, 0);
    
      For r_Deposit In c_Deposit(n_原结帐id) Loop
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 d_Date, 操作员姓名_In, 操作员编号_In, Decode(Sign(r_Deposit.金额 - n_预交金额), -1, -1 * r_Deposit.金额, -1 * n_预交金额),
                 n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, Decode(校对标志_In, 1, 2, 0), n_结算序号, 3
          From 病人预交记录
          Where ID = r_Deposit.Id;
      
        --检查是否已经处理完 
        If r_Deposit.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
      If Nvl(校对标志_In, 0) = 0 Then
        --更新病人预交余额 
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) + n_总金额 + Nvl(误差_In, 0)
        Where 病人id = n_病人id And 性质 = 1 And 类型 = 1
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, n_总金额 + Nvl(误差_In, 0), 1);
          n_返回值 := n_总金额 + Nvl(误差_In, 0);
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 病人余额 Where 病人id = n_病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --更新原记录
  Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id;

  If 多单据全退_In <> 1 Then
    --处理误差项,多单据全退时 ,按原样退无误差处理
    --将误差项的记录状态调整为3
    If Nvl(误差_In, 0) <> 0 Then
      n_Count := 1;
      If n_正常退费 = 1 And Nvl(退款操作_In, 0) = 0 Then
        n_原误差费 := 0;
        --原样退,但存在误差
        If n_部分退 = 0 Then
          Select -1 * Nvl(Sum(实收金额), 0)
          Into n_原误差费
          From 门诊费用记录 A
          Where NO = No_In And a.记录性质 = 1 And a.记录状态 In (1, 3) And Nvl(a.附加标志, 0) = 9;
        End If;
        If Nvl(n_原误差费, 0) <> 0 Or Nvl(误差_In, 0) <> 0 Then
          Update 病人预交记录
          Set 冲预交 = 冲预交 - n_原误差费 - Nvl(误差_In, 0)
          Where 结算方式 = v_退费结算 And 结帐id = n_结帐id;
          If Sql%NotFound Then
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
              Select 病人预交记录_Id.Nextval, 3, No_In, 2, 病人id, 主页id, '误差费', v_退费结算, d_Date, Null, Null, Null, 操作员编号_In,
                     操作员姓名_In, -1 * n_原误差费 - Nvl(误差_In, 0), n_结帐id, n_组id, 预交类别, Null, Null, Null, Null, Null, Null, 0,
                     n_结算序号, 3
              From 病人预交记录
              Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id And Rownum = 1;
          End If;
        End If;
      End If;
    Elsif n_正常退费 = 1 And Nvl(退款操作_In, 0) = 0 Then
      --原样退时,需要处理预交记录不足的情况
      Select Nvl(Sum(Nvl(结帐金额, 0)), 0) Into n_实收金额 From 门诊费用记录 Where 结帐id = n_结帐id;
      Select Nvl(Sum(Nvl(冲预交, 0)), 0) Into n_返回值 From 病人预交记录 Where 结帐id = n_结帐id;
      If Abs(n_实收金额) <> Abs(n_返回值) Then
        n_实收金额 := n_实收金额 - n_返回值;
        Update 病人预交记录 Set 冲预交 = 冲预交 + Nvl(n_实收金额, 0) Where 结算方式 = v_退费结算 And 结帐id = n_结帐id;
        If Sql%NotFound Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
             卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
            Select 病人预交记录_Id.Nextval, 3, No_In, 2, 病人id, 主页id, '误差费', v_退费结算, d_Date, Null, Null, Null, 操作员编号_In,
                   操作员姓名_In, Nvl(n_实收金额, 0), n_结帐id, n_组id, 预交类别, Null, Null, Null, Null, Null, Null, 0, n_结算序号, 3
            From 病人预交记录
            Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id And Rownum = 1;
        End If;
      End If;
    End If;
  
    Select Nvl(Sum(Nvl(结帐金额, 0)), 0) Into n_实收金额 From 门诊费用记录 Where 结帐id = n_结帐id;
    Select Nvl(Sum(Nvl(冲预交, 0)), 0) Into n_返回值 From 病人预交记录 Where 结帐id = n_结帐id;
  
    n_实收金额 := n_实收金额 - n_返回值;
  
    If n_实收金额 <> 0 Then
      --未找到，新产生误差项
      Zl_门诊收费误差_Insert(No_In, n_实收金额, 1);
    End If;
  End If;
  --------------------------------------------------------------------------------- 
  --人员缴款余额(注意是预交记录处理后才处理，包括个人帐户等的结算金额,不含退冲预交款) 
  --如果是需要校对的,暂不处理人员缴款余额 
  If Nvl(校对标志_In, 0) = 0 Then
    For r_Moneyrow In c_Money(n_结帐id) Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + r_Moneyrow.冲预交
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Moneyrow.结算方式
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Moneyrow.结算方式, 1, r_Moneyrow.冲预交);
        n_返回值 := r_Moneyrow.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Moneyrow.结算方式 And Nvl(余额, 0) = 0;
      End If;
    End Loop;
  End If;

  --------------------------------------------------------------------------------- 
  --退费票据回收(仅全退时才回退,部分退是在重打过程中回收) 
  If 回收票据_In = 0 Then
  
    --启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
    v_Para     := Nvl(zl_GetSysParameter('票据分配规则', 1121), '0||0;0;0,0;0');
    n_启用模式 := Zl_To_Number(Substr(v_Para, 1, 1));
    If n_启用模式 <> 0 Then
      --收回票据
      Select 使用id Bulk Collect
      Into l_使用id
      From (Select Distinct b.使用id From 票据打印明细 B Where b.No = No_In And Nvl(b.票种, 0) = 1);
    
      n_启用模式 := l_使用id.Count;
      If l_使用id.Count <> 0 Then
        --插入回收记录
        Forall I In 1 .. l_使用id.Count
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 操作员姓名_In, d_Date
            From 票据使用明细 A
            Where ID = l_使用id(I) And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 Where 号码 = a.号码 And 票种 = a.票种 And Nvl(性质, 0) <> 1);
      
        Forall I In 1 .. l_使用id.Count
          Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I) And Nvl(是否回收, 0) = 0;
      
      End If;
    End If;
    If n_启用模式 = 0 Then
      --获取单据最后一次的打印ID(可能是多张单据收费打印) 
      Begin
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 1 And b.No = No_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --可能以前没有打印,无收回 
      If n_打印id Is Not Null Then
        --a.多张单据循环调用时只能收回一次 
        Select Count(*) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
        If n_Count = 0 Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
            From 票据使用明细
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
        Else
          --b.部分退费多次收回时,最后一次全退收回要排开已收回的 
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
            From 票据使用明细 A
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And 打印id = n_打印id And 票种 = 1 And 性质 = 2);
        End If;
      End If;
    End If;
  End If;

  --------------------------------------------------------------------------------- 
  --卫生材料 
  For v_出库 In (Select ID, 药品id, 库房id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
               From 药品收发记录
               Where 单据 = 21 --@@@ 
                     And Mod(记录状态, 3) = 1 And 审核人 Is Null And
                     费用id In (Select ID
                              From 门诊费用记录
                              Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And 收费类别 = '4' --@@@ 
                                    And (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
               Order By 药品id) Loop
    --处理药品库存 
    If v_出库.库房id Is Not Null Then
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0)
      Where 库房id = v_出库.库房id And 药品id = v_出库.药品id And Nvl(批次, 0) = Nvl(v_出库.批次, 0) And 性质 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 药品库存
          (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码) --@@@ 
        Values
          (v_出库.库房id, v_出库.药品id, 1, v_出库.批次, v_出库.效期,
           Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0), v_出库.批号, v_出库.产地, v_出库.灭菌效期,
           v_出库.商品条码, v_出库.内部条码);
      End If;
    End If;
  
    l_费用id.Extend;
    l_费用id(l_费用id.Count) := v_出库.费用id;
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := v_出库.Id;
  End Loop;

  --药品相关内容 
  For r_Stock In c_Stock Loop
    --处理药品库存 
    If r_Stock.库房id Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_备货卫材
      From Table(l_费用id)
      Where Column_Value = r_Stock.费用id;
      If Nvl(n_备货卫材, 0) = 0 Then
      
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0)
        Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码) --@@@ 
          Values
            (r_Stock.库房id, r_Stock.药品id, 1, r_Stock.批次, r_Stock.效期,
             Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0), r_Stock.批号, r_Stock.产地,
             r_Stock.灭菌效期, r_Stock.商品条码, r_Stock.内部条码);
        End If;
      End If;
    End If;
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := r_Stock.Id;
  End Loop;

  --删除药品收发记录 
  Forall I In 1 .. l_药品收发.Count
    Delete From 药品收发记录 Where ID = l_药品收发(I) And 审核人 Is Null;

  --未发药品记录 
  For r_Spare In c_Spare Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From 药品收发记录
    Where NO = No_In And 单据 = r_Spare.单据 --@@@ 
          And Mod(记录状态, 3) = 1 And 审核人 Is Null And Nvl(库房id, 0) = Nvl(r_Spare.库房id, 0);
  
    If n_Count = 0 Then
      Delete From 未发药品记录
      Where 单据 = r_Spare.单据 --@@@ 
            And NO = No_In And Nvl(库房id, 0) = Nvl(r_Spare.库房id, 0);
    End If;
  End Loop;
  --医嘱处理 
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 1 And 记录状态 = 3 And 医嘱序号 Is Not Null And
                     (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null)) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, 执行状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where 记录性质 = 1 And Nvl(附加标志, 0) <> 9 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, 执行状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      --更新计费状态为1:已计费 
      Update 病人医嘱发送 A Set 计费状态 = 1 Where NO = No_In And 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1;
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1 And NO = No_In;
    Else
      --更新计费状态为2:部分退费 
      Update 病人医嘱发送 A Set 计费状态 = 2 Where NO = No_In And 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_Delete;
/

--90389:冉俊明,2015-11-16,医嘱附费项目退费数量错误。
Create Or Replace Procedure Zl_门诊简单收费_Delete
(
  No_In         门诊费用记录.No%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type
) As
  --功能：删除一张门诊简单收费单据

  --该游标为要退费单据的所有原始记录
  Cursor c_Bill Is
    Select a.Id, a.No, a.附加标志, a.收费细目id, a.序号, a.价格父号, a.执行状态, a.收费类别, a.付数, a.数次, a.医嘱序号, j.诊疗类别, m.跟踪在用
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.No = No_In And a.记录性质 = 1 And a.记录状态 In (1, 3) And a.收费细目id + 0 = m.材料id(+)
    Order By a.收费细目id, a.序号;

  --该游标用于处理药品库存可用数量
  --不要管费用的执行状态,因为先于此步处理
  Cursor c_Stock Is
    Select ID, 药品id, 库房id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
    From 药品收发记录
    Where NO = No_In And 单据 In (8, 24) --@@@
          And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 门诊费用记录
                   Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And 收费类别 In ('4', '5', '6', '7'))
    Order By 药品id;

  --该游标用于处理未发药品记录
  Cursor c_Spare Is
    Select NO, 库房id, 单据 From 未发药品记录 Where NO = No_In And 单据 In (8, 24); --@@@

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Money(冲销id_In 病人预交记录.结帐id%Type) Is
    Select 结算方式, 冲预交
    From 病人预交记录
    Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = 冲销id_In And 结算方式 Is Not Null And Nvl(冲预交, 0) <> 0 And Nvl(校对标志, 0) = 0;

  --该游标用于查找收费时使用过的冲预交款记录
  Cursor c_Deposit(V结帐id 病人预交记录.结帐id%Type) Is
    Select ID, 冲预交 As 金额, 预交类别
    From 病人预交记录
    Where 记录性质 In (1, 11) And 记录状态 In (1, 3) And 结帐id = V结帐id And Nvl(冲预交, 0) <> 0
    Order By ID Desc;

  n_病人id   病人信息.病人id%Type;
  n_结帐id   门诊费用记录.结帐id%Type;
  n_结算序号 病人预交记录.结算序号%Type;
  n_打印id   票据打印内容.Id%Type;

  n_预交金额 病人预交记录.冲预交%Type;
  n_返回值   病人预交记录.冲预交%Type;
  --部分退费计算变量
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;
  n_备货卫材 Number;
  n_准退数量 Number;
  n_退费次数 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;
  n_总金额   Number;
  n_费用状态 门诊费用记录.费用状态%Type;
  n_正常退费 Number; --是否第一次退费且全部退费,在每行退费过程中判断得到。
  n_组id     财务缴款分组.Id%Type;

  v_退费结算 结算方式.名称%Type;

  l_费用id   t_Numlist := t_Numlist();
  l_药品收发 t_Numlist := t_Numlist();
  l_使用id   t_Numlist := t_Numlist();
  n_Dec      Number;
  d_Date     Date;
  n_Count    Number;
  n_原结帐id Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_启用模式     Number(3);
  v_Para         Varchar2(1000);
  n_医属执行计价 Number;

Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  --是否已经全部完全执行(只是该单据整张单据的检查)
  Select Nvl(Count(*), 0)
  Into n_Count
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  --执行状态在原始记录上判断
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 门诊费用记录
                Where NO = No_In And 记录性质 = 1 And Nvl(附加标志, 0) <> 9 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以退费的费用！';
    Raise Err_Item;
  End If;
  --确定是否在医嘱执行计价中存在数据,如果存在数据,则根据医嘱执行计价进行退费,否则按旧方式进行处理
  Select Count(1)
  Into n_医属执行计价
  From 门诊费用记录 A, 医嘱执行计价 B
  Where a.医嘱序号 = b.医嘱id And a.记录性质 = 1 And a.No = No_In And a.记录状态 In (1, 3) And Rownum = 1;

  ---------------------------------------------------------------------------------
  --公用变量
  Select Sysdate Into d_Date From Dual;
  Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  n_结算序号 := Null;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --获取结算方式名称
  Begin
    Select 名称 Into v_退费结算 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_退费结算 := '现金';
  End;

  ---------------------------------------------------------------------------------
  --循环处理每行费用(收入项目行)
  n_总金额   := 0;
  n_正常退费 := 1;
  For r_Bill In c_Bill Loop
    If Nvl(r_Bill.执行状态, 0) <> 1 Then
      --求剩余数量,剩余应收,剩余实收
      Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
      Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
      From 门诊费用记录
      Where NO = No_In And 记录性质 = 1 And 序号 = r_Bill.序号;
    
      If n_剩余数量 = 0 Then
        --情况：未限定行号,原始单据中的该笔已经全部退费(执行状态=0的一种可能)
        n_正常退费 := 0;
      Else
        --准退数量(非药品项目为剩余数量,原始数量)
        If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
          --@@@
          --非药品部分(以具体医嘱执行为准进行检查)
          --: 1.存在医嘱发送的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血)
          --: 2.不存在医嘱的,则以剩余数量为准
          n_Count := 0;
          If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null Then
            If n_医属执行计价 = 1 Then
              Select Decode(Sign(Sum(数量)), -1, 0, Sum(数量)), Count(*)
              Into n_准退数量, n_Count
              From (Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, Max(a.医嘱序号) As 医嘱id, Max(a.收费细目id) As 收费细目id,
                            Sum(Nvl(a.付数, 1) * Nvl(a.数次, 1)) As 数量,
                            Sum(Decode(a.记录状态, 2, 0, Nvl(a.付数, 1) * Nvl(a.数次, 1))) As 原始数量
                     From 门诊费用记录 A, 病人医嘱记录 M
                     Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                           Instr('5,6,7', a.收费类别) = 0 And a.No = No_In And a.序号 = r_Bill.序号 And a.记录性质 = 1 And
                           a.记录状态 In (1, 2, 3) And a.价格父号 Is Null
                     Group By a.序号
                     Union All
                     Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, -1 * b.数量 As 已执行, 0 原始数量
                     From 门诊费用记录 A, 医嘱执行计价 B, 病人医嘱记录 M
                     Where a.医嘱序号 = b.医嘱id And a.收费细目id = b.收费细目id + 0 And a.医嘱序号 = m.Id And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Instr('5,6,7', a.收费类别) = 0 And
                           (Exists
                            (Select 1
                             From 病人医嘱执行
                             Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And b.要求时间 = 要求时间 And Nvl(执行结果, 0) = 1) Or Exists
                            (Select 1
                             From 病人医嘱发送
                             Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And Nvl(执行状态, 0) = 1)) And Not Exists
                      (Select 1
                            From 病人医嘱附费
                            Where a.医嘱序号 = 医嘱id And a.No = NO And Mod(a.记录性质, 10) = 记录性质) And a.No = No_In And
                           a.序号 = r_Bill.序号 And a.记录性质 = 1 And a.记录状态 In (1, 3) 　and a.价格父号 Is Null) Q1
              Where Not Exists (Select 1 From 药品收发记录 Where 费用id = Q1.Id) Having Max(ID) <> 0;
            Else
              Select Nvl(Sum(数量), 0), Count(*)
              Into n_准退数量, n_Count
              From (Select a.医嘱id, a.收费细目id, Nvl(a.数量, 1) * Nvl(b.发送数次, 1) As 数量
                     From 病人医嘱计价 A, 病人医嘱发送 B, 门诊费用记录 J, 病人医嘱记录 M
                     Where a.医嘱id = b.医嘱id And a.医嘱id = m.Id And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And
                           a.收费细目id = j.收费细目id And j.No = No_In And j.记录性质 = 1 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And
                           j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Exists
                      (Select 1
                            From 病人医嘱计价 A
                            Where a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) = 0) And Not Exists
                      (Select 1 From 药品收发记录 Where 费用id = j.Id)
                     Union All
                     Select a.医嘱id, a.收费细目id, -1 * Nvl(a.数量, 1) * Nvl(c.本次数次, 1) As 数量
                     From 病人医嘱计价 A, 病人医嘱发送 B, 病人医嘱执行 C, 门诊费用记录 J, 病人医嘱记录 M
                     Where a.医嘱id = b.医嘱id And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And a.医嘱id = m.Id And Nvl(c.执行结果, 1) = 1 And
                           Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And j.No = No_In And
                           j.记录性质 = 1 And Nvl(a.收费方式, 0) = 0 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And j.价格父号 Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Not Exists
                      (Select 1 From 药品收发记录 Where 费用id = j.Id) And Not Exists
                      (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)
                     Union All
                     Select a.医嘱序号 As 医嘱id, a.收费细目id, Nvl(a.付数, 1) * a.数次 As 数量
                     From 门诊费用记录 A, 病人医嘱记录 M
                     Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And a.No = No_In And
                           a.记录性质 = 1 And a.序号 = r_Bill.序号 And a.记录状态 = 2 And a.价格父号 Is Null And Not Exists
                      (Select 1 From 药品收发记录 Where NO = No_In And 单据 In (8, 24) And 药品id = a.收费细目id));
            End If;
          End If;
          If Nvl(n_Count, 0) <> 0 And n_准退数量 = 0 Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已执行,不允许退费！';
            Raise Err_Item;
          End If;
        
          If Nvl(n_Count, 0) = 0 Then
            n_准退数量 := n_剩余数量;
          End If;
        
        Else
          Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0), Count(*)
          Into n_准退数量, n_Count
          From 药品收发记录
          Where NO = No_In And 单据 In (8, 24) And Mod(记录状态, 3) = 1 --@@@
                And 审核人 Is Null And 费用id = r_Bill.Id;
        
          --有剩余数量无准退数量的有两种情况：
          --1.不跟踪在用的卫材无对应的收发记录,这时使用剩余数量
          --2.并发操作,此时已发药或发料
          If n_准退数量 = 0 Then
            If r_Bill.收费类别 = '4' Then
              If n_Count > 0 Then
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发料,须退料后再退费！';
                Raise Err_Item;
              Else
                n_准退数量 := n_剩余数量;
              End If;
            Else
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发药,须退药后再退费！';
              Raise Err_Item;
            End If;
          End If;
        End If;
      
        --是否部分退费
        If r_Bill.执行状态 = 2 Or n_准退数量 <> Nvl(r_Bill.付数, 1) * r_Bill.数次 Then
          n_正常退费 := 0;
        End If;
      
        --处理门诊费用记录
        n_费用状态 := 0;
        --该笔项目第几次退费
        Select Nvl(Max(Abs(执行状态)), 0) + 1
        Into n_退费次数
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 1 And 记录状态 = 2 And Nvl(执行状态, 0) < 0 And 序号 = r_Bill.序号;
      
        n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), n_Dec);
        n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), n_Dec);
        n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), n_Dec);
        n_总金额   := n_总金额 + n_实收金额;
      
        --插入退费记录
        Insert Into 门诊费用记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
           计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态,
           费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论, 缴款组id)
          Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                 病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                 Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价,
                 -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, n_费用状态, 执行时间, 操作员编号_In, 操作员姓名_In,
                 发生时间, d_Date, n_结帐id, -1 * n_实收金额, 保险项目否, 保险大类id, -1 * n_统筹金额, 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码,
                 费用类型, 结论, n_组id
          From 门诊费用记录
          Where ID = r_Bill.Id;
      
        --标记原费用记录
        --执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1,异常收费单,还是标明9
        Update 门诊费用记录
        Set 记录状态 = 3, 执行状态 = Decode(Nvl(执行状态, 0), 9, 9, Decode(Sign(n_准退数量 - n_剩余数量), 0, 0, 1))
        Where ID = r_Bill.Id;
      End If;
    Else
      --情况:没限定行号,原始单据中包括已经完全执行的
      n_正常退费 := 0;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --处理病人预交记录
  --自动产生误差费,默认保留一位
  n_总金额 := Round(n_总金额, 1);
  --原单据的结帐ID
  Select 结帐id, 病人id
  Into n_原结帐id, n_病人id
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Rownum = 1;

  If n_正常退费 = 1 Then
    --单据第一次退费且全部退完
    --冲预交部分记录
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date,
             操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 0, n_结算序号, 3
      From 病人预交记录
      Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
    --处理病人预交余额
    For v_预交 In (Select 预交类别, Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交金额
                 From 病人预交记录
                 Where 记录性质 In (1, 11) And 结帐id = n_原结帐id
                 Group By 预交类别
                 Having Sum(Nvl(冲预交, 0)) <> 0) Loop
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_预交.预交金额, 0)
      Where 病人id = n_病人id And 性质 = 1 And 类型 = Nvl(v_预交.预交类别, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 类型, 预交余额, 性质)
        Values
          (n_病人id, Nvl(v_预交.预交类别, 2), Nvl(v_预交.预交金额, 0), 1);
        n_返回值 := n_预交金额;
      End If;
      If n_返回值 = 0 Then
        Delete From 病人余额 Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    End Loop;
  
    --原样退回(冲预交在前面已处理)
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, d_Date, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
             -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 0, n_结算序号, 3
      From 病人预交记录 A, (Select 名称 From 结算方式 Where 性质 In (3, 4)) J,
           (Select m.Id As 预交id From 病人预交记录 M Where m.结帐id = n_原结帐id And m.记录性质 = 3 And m.记录状态 = 1) Q
      Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id = n_原结帐id And a.Id = q.预交id(+) And a.结算方式 = j.名称(+);
  Else
    --部分退费直接退为指定结算方式
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
       结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, 3, No_In, 2, 病人id, 主页id, '部分退费结算', v_退费结算, d_Date, Null, Null, Null, 操作员编号_In, 操作员姓名_In,
             -1 * n_总金额, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 0, n_结算序号, 3
      
      From 病人预交记录
      Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id And Rownum = 1;
  
    --如果收费时只使用了预交款,则要退预交,并且可能有多笔冲预交
    If Sql%RowCount = 0 Then
      n_预交金额 := n_总金额;
    
      For r_Deposit In c_Deposit(n_原结帐id) Loop
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 d_Date, 操作员姓名_In, 操作员编号_In, Decode(Sign(r_Deposit.金额 - n_预交金额), -1, -1 * r_Deposit.金额, -1 * n_预交金额),
                 n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 0, n_结算序号, 3
          From 病人预交记录
          Where ID = r_Deposit.Id;
        --检查是否已经处理完
        If r_Deposit.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + n_总金额
      Where 病人id = n_病人id And 性质 = 1 And 类型 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, n_总金额, 1);
        n_返回值 := n_总金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额 Where 病人id = n_病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    End If;
  End If;
  --更新原记录
  Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id;

  Select Nvl(Sum(Nvl(结帐金额, 0)), 0) Into n_实收金额 From 门诊费用记录 Where 结帐id = n_结帐id;
  Select Nvl(Sum(Nvl(冲预交, 0)), 0) Into n_返回值 From 病人预交记录 Where 结帐id = n_结帐id;

  n_实收金额 := n_实收金额 - n_返回值;

  If n_实收金额 <> 0 Then
    --未找到，新产生误差项
    Zl_简单收费误差_Insert(No_In, n_病人id, n_结帐id, n_实收金额, d_Date, 操作员编号_In, 操作员姓名_In, 1);
  End If;

  --人员缴款余额(注意是预交记录处理后才处理，包括个人帐户等的结算金额,不含退冲预交款)
  For r_Moneyrow In c_Money(n_结帐id) Loop
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + r_Moneyrow.冲预交
    Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Moneyrow.结算方式
    Returning 余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (操作员姓名_In, r_Moneyrow.结算方式, 1, r_Moneyrow.冲预交);
      n_返回值 := r_Moneyrow.冲预交;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Moneyrow.结算方式 And Nvl(余额, 0) = 0;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --退费票据回收(仅全退时才回退,部分退是在重打过程中回收)
  --启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
  v_Para     := Nvl(zl_GetSysParameter('票据分配规则', 1121), '0||0;0;0,0;0');
  n_启用模式 := Zl_To_Number(Substr(v_Para, 1, 1));
  If n_启用模式 <> 0 Then
    --收回票据
    Select 使用id Bulk Collect
    Into l_使用id
    From (Select Distinct b.使用id From 票据打印明细 B Where b.No = No_In And Nvl(b.票种, 0) = 1);
  
    n_启用模式 := l_使用id.Count;
    If l_使用id.Count <> 0 Then
      --插入回收记录
      Forall I In 1 .. l_使用id.Count
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 操作员姓名_In, d_Date
          From 票据使用明细 A
          Where ID = l_使用id(I) And 性质 = 1 And Not Exists
           (Select 1 From 票据使用明细 Where 号码 = a.号码 And 票种 = a.票种 And Nvl(性质, 0) <> 1);
    
      Forall I In 1 .. l_使用id.Count
        Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I) And Nvl(是否回收, 0) = 0;
    
    End If;
  End If;
  If n_启用模式 = 0 Then
    --获取单据最后一次的打印ID(可能是多张单据收费打印)
    Begin
      Select ID
      Into n_打印id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 1 And b.No = No_In
             Order By a.使用时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    --可能以前没有打印,无收回
    If n_打印id Is Not Null Then
      --a.多张单据循环调用时只能收回一次
      Select Count(*) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
      If n_Count = 0 Then
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
          From 票据使用明细
          Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
      Else
        --b.部分退费多次收回时,最后一次全退收回要排开已收回的
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
          From 票据使用明细 A
          Where 打印id = n_打印id And 票种 = 1 And 性质 = 1 And Not Exists
           (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And 打印id = n_打印id And 票种 = 1 And 性质 = 2);
      End If;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --卫生材料
  For v_出库 In (Select ID, 药品id, 库房id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
               From 药品收发记录
               Where 单据 = 21 --@@@
                     And Mod(记录状态, 3) = 1 And 审核人 Is Null And
                     费用id In (Select ID
                              From 门诊费用记录
                              Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And 收费类别 = '4')
               Order By 药品id) Loop
    --处理药品库存
    If v_出库.库房id Is Not Null Then
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0)
      Where 库房id = v_出库.库房id And 药品id = v_出库.药品id And Nvl(批次, 0) = Nvl(v_出库.批次, 0) And 性质 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 药品库存
          (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码) --@@@
        Values
          (v_出库.库房id, v_出库.药品id, 1, v_出库.批次, v_出库.效期,
           Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0), v_出库.批号, v_出库.产地, v_出库.灭菌效期,
           v_出库.商品条码, v_出库.内部条码);
      End If;
    End If;
  
    l_费用id.Extend;
    l_费用id(l_费用id.Count) := v_出库.费用id;
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := v_出库.Id;
  End Loop;

  ---------------------------------------------------------------------------------
  --药品相关内容
  For r_Stock In c_Stock Loop
    --处理药品库存
    If r_Stock.库房id Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_备货卫材
      From Table(l_费用id)
      Where Column_Value = r_Stock.费用id;
      If Nvl(n_备货卫材, 0) = 0 Then
      
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0)
        Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码) --@@@
          Values
            (r_Stock.库房id, r_Stock.药品id, 1, r_Stock.批次, r_Stock.效期,
             Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0), r_Stock.批号, r_Stock.产地,
             r_Stock.灭菌效期, r_Stock.商品条码, r_Stock.内部条码);
        End If;
      End If;
    End If;
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := r_Stock.Id;
  End Loop;

  --删除药品收发记录
  Forall I In 1 .. l_药品收发.Count
    Delete From 药品收发记录 Where ID = l_药品收发(I) And 审核人 Is Null;

  --未发药品记录
  For r_Spare In c_Spare Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From 药品收发记录
    Where NO = No_In And 单据 = r_Spare.单据 --@@@
          And Mod(记录状态, 3) = 1 And 审核人 Is Null And Nvl(库房id, 0) = Nvl(r_Spare.库房id, 0);
  
    If n_Count = 0 Then
      Delete From 未发药品记录
      Where 单据 = r_Spare.单据 --@@@
            And NO = No_In And Nvl(库房id, 0) = Nvl(r_Spare.库房id, 0);
    End If;
  End Loop;
  --医嘱处理
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 1 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, 执行状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where 记录性质 = 1 And Nvl(附加标志, 0) <> 9 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, 执行状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      --更新计费状态为1:已计费
      Update 病人医嘱发送 A Set 计费状态 = 1 Where NO = No_In And 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1;
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1 And NO = No_In;
    Else
      --更新计费状态为2:部分退费
      Update 病人医嘱发送 A Set 计费状态 = 2 Where NO = No_In And 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊简单收费_Delete;
/

--90466:刘尔旋,2015-11-12,支付宝结帐修改
Create Or Replace Procedure Zl_三方退款信息_Insert
(
  结帐id_In     三方退款信息.结帐id%Type,
  记录id_In     三方退款信息.记录id%Type,
  金额_In       三方退款信息.金额%Type,
  卡号_In       三方退款信息.卡号%Type,
  交易流水号_In 三方退款信息.交易流水号%Type,
  交易说明_In   三方退款信息.交易说明%Type,
  操作类型_In   Number := 0
) As
  --功能：用于填制多笔退款交易扩展的结算信息
  --操作类型_In:0=新增,1=更新信息
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
Begin
  If 操作类型_In = 0 Then
    Insert Into 三方退款信息
      (结帐id, 记录id, 金额, 卡号, 交易流水号, 交易说明)
    Values
      (结帐id_In, 记录id_In, 金额_In, 卡号_In, 交易流水号_In, 交易说明_In);
  Else
    Update 三方退款信息
    Set 卡号 = Nvl(卡号_In, 卡号), 交易流水号 = Nvl(交易流水号_In, 交易流水号), 交易说明 = Nvl(交易说明_In, 交易说明)
    Where 结帐id = 结帐id_In And 记录id = 记录id_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方退款信息_Insert;
/

--90488:刘尔旋,2015-11-11,存在计划时修改挂号安排
Create Or Replace Procedure Zl_挂号安排_Modify
(
  Id_In       挂号安排.Id%Type,
  诊室_In     Varchar2 := Null,
  预约天数_In 挂号安排.预约天数%Type := Null
) As
  v_诊室  Varchar2(1000);
  n_Count Number(18);
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

Begin
  If 诊室_In Is Not Null Then
    Delete From 挂号安排诊室 Where 号表id = Id_In;
    v_诊室 := 诊室_In || ';';
    While v_诊室 Is Not Null Loop
      Insert Into 挂号安排诊室 (号表id, 门诊诊室) Values (Id_In, Substr(v_诊室, 1, Instr(v_诊室, ';') - 1));
      v_诊室 := Substr(v_诊室, Instr(v_诊室, ';') + 1);
    End Loop;
  End If;
  Update 挂号安排 Set 预约天数 = 预约天数_In Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号安排_Modify;
/

--90436:刘尔旋,2015-11-10,获取病人标示变动
Create Or Replace Procedure Zl_Third_Getpati_Unique
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:获取病人的唯一标识(病人ID) 
  --入参:Xml_In: 
  --   <IN> 
  --     <ZJH>证件号</ZJH>     //卡号 
  --     <ZJLX>证件类型</ZJLX>  //医疗卡类别.名称 
  --     <XM>姓名</XM> 
  --     <KH>卡号</KH> 
  --     <KLB>卡类别</KLB> 
  --    </IN> 
  --出参:Xml_Out 
  -- <OUTPUT> 
  --   <BRID>病人ID</BRID> 
  --   <MZH>门诊号</MZH> 
  --   <ERROR><MSG>错误信息</MSG></ERROR> 
  --  </OUTPUT> 

  -------------------------------------------------------------------------------------------------- 
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  x_Templet Xmltype; --模板XML 

  v_证件号       Varchar2(50);
  v_证件类型     医疗卡类别.名称%Type;
  v_姓名         Varchar2(100);
  v_验证姓名     Varchar2(100);
  v_卡号         Varchar2(100);
  v_卡类别       Varchar2(100);
  n_存在         Number(3);
  n_卡类别id     病人医疗卡信息.卡类别id%Type;
  v_操作员       人员表.姓名%Type;
  v_验证身份证号 病人信息.身份证号%Type;

  n_病人id     病人信息.病人id%Type;
  n_验证病人id 病人信息.病人id%Type;
  n_门诊号     病人信息.门诊号%Type;

  v_Temp Varchar2(32767); --临时XML 
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZJH'), Extractvalue(Value(A), 'IN/ZJLX'), Extractvalue(Value(A), 'IN/XM'),
         Extractvalue(Value(A), 'IN/KH'), Extractvalue(Value(A), 'IN/KLB')
  Into v_证件号, v_证件类型, v_姓名, v_卡号, v_卡类别
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  --先从用户自定义过程获取病人ID 
  n_病人id := Zl_Third_Custom_Getpati(v_卡类别, v_卡号);

  If Nvl(n_病人id, 0) = 0 Then
    If Nvl(v_卡类别, '-') <> '-' And Nvl(v_卡号, '-') <> '-' Then
      Select Max(a.病人id)
      Into n_病人id
      From 病人医疗卡信息 A
      Where a.卡类别id = (Select Max(ID) From 医疗卡类别 Where 名称 = v_卡类别 And Nvl(是否启用, 0) = 1) And 卡号 = v_卡号 And
            (a.状态 = 0 Or Nvl(a.状态, 0) = 1 And Exists
             (Select 1 From 医疗卡挂失方式 Where a.挂失方式 = 名称 And a.挂失时间 + 有效天数 < Trunc(Sysdate)));
    End If;
  
    If Nvl(n_病人id, 0) = 0 Then
      v_Err_Msg := '未找到该登记号的病人信息，请检查输入的登记号是否正确!';
      Raise Err_Item;
    End If;
  
  End If;

  If Nvl(n_病人id, 0) <> 0 Then
    v_Temp := '<BRID>' || n_病人id || '</BRID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Select 门诊号, 姓名, 身份证号 Into n_门诊号, v_验证姓名, v_验证身份证号 From 病人信息 Where 病人id = n_病人id;
    If v_姓名 <> v_验证姓名 Then
      v_Err_Msg := '传入的病人姓名与已有的病人姓名不符,请检查!';
      Raise Err_Item;
    End If;
    If Not v_证件号 Is Null Then
      If v_证件类型 Like '%身份证%' Then
        If v_证件号 <> v_验证身份证号 Then
          v_Err_Msg := '传入的病人身份证号与已有的病人身份证号不符,请检查!';
          Raise Err_Item;
        End If;
      Else
        Select Max(病人id)
        Into n_验证病人id
        From 病人医疗卡信息 A
        Where 卡类别id = (Select Max(ID) From 医疗卡类别 Where 名称 = v_证件类型 And Nvl(是否启用, 0) = 1) And 卡号 = v_证件号 And
              (a.状态 = 0 Or Nvl(a.状态, 0) = 1 And Exists
               (Select 1 From 医疗卡挂失方式 Where a.挂失方式 = 名称 And a.挂失时间 + 有效天数 < Trunc(Sysdate)));
        If n_验证病人id <> n_病人id Then
          v_Err_Msg := '传入的病人证件类型与已有的病人证件类型不符,请检查!';
          Raise Err_Item;
        End If;
      End If;
    End If;
    If n_门诊号 Is Null Then
      n_门诊号 := Nextno(3);
      Update 病人信息 Set 门诊号 = n_门诊号 Where 病人id = n_病人id;
    End If;
    v_Temp := '<MZH>' || n_门诊号 || '</MZH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getpati_Unique;
/

--90415:冉俊明,2015-11-09,清除划价单时没有清除高值卫材对应的材料其它出库单记录。
Create Or Replace Procedure Zl_门诊划价记录_Clear(Day_In Number) As
  --功能：自动清除划价单
  --参数：Day_IN=删除划价后超过Day_IN天未收费的单据
  Cursor c_Price Is
    Select Distinct a.No
    From 门诊费用记录 A, 未发药品记录 B
    Where a.记录性质 = 1 And a.记录状态 = 0 And a.执行状态 Not In (1, 2) And a.划价人 Is Not Null And a.操作员姓名 Is Null And
          b.单据 In (8, 24) And Nvl(b.已收费, 0) = 0 And a.No = b.No And Nvl(a.执行部门id, 0) = Nvl(b.库房id, 0) And
          Sysdate - b.填制日期 >= Day_In;
Begin
  For r_Price In c_Price Loop
    Zl_门诊划价记录_Delete(r_Price.No);
    Commit;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊划价记录_Clear;
/

--89624:刘尔旋,2015-11-09,挂号零费用消费卡问题
Create Or Replace Procedure Zl_病人挂号记录_Insert
(
  病人id_In       门诊费用记录.病人id%Type,
  门诊号_In       门诊费用记录.标识号%Type,
  姓名_In         门诊费用记录.姓名%Type,
  性别_In         门诊费用记录.性别%Type,
  年龄_In         门诊费用记录.年龄%Type,
  付款方式_In     门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In         门诊费用记录.费别%Type,
  单据号_In       门诊费用记录.No%Type,
  票据号_In       门诊费用记录.实际票号%Type,
  序号_In         门诊费用记录.序号%Type,
  价格父号_In     门诊费用记录.价格父号%Type,
  从属父号_In     门诊费用记录.从属父号%Type,
  收费类别_In     门诊费用记录.收费类别%Type,
  收费细目id_In   门诊费用记录.收费细目id%Type,
  数次_In         门诊费用记录.数次%Type,
  标准单价_In     门诊费用记录.标准单价%Type,
  收入项目id_In   门诊费用记录.收入项目id%Type,
  收据费目_In     门诊费用记录.收据费目%Type,
  结算方式_In     病人预交记录.结算方式%Type, --现金的结算名称
  应收金额_In     门诊费用记录.应收金额%Type,
  实收金额_In     门诊费用记录.实收金额%Type,
  病人科室id_In   门诊费用记录.病人科室id%Type,
  开单部门id_In   门诊费用记录.开单部门id%Type,
  执行部门id_In   门诊费用记录.执行部门id%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  发生时间_In     门诊费用记录.发生时间%Type,
  登记时间_In     门诊费用记录.登记时间%Type,
  医生姓名_In     挂号安排.医生姓名%Type,
  医生id_In       挂号安排.医生id%Type,
  病历费_In       Number, --该条记录是否病历工本费
  急诊_In         Number,
  号别_In         挂号安排.号码%Type,
  诊室_In         门诊费用记录.发药窗口%Type,
  结帐id_In       门诊费用记录.结帐id%Type,
  领用id_In       票据使用明细.领用id%Type,
  预交支付_In     病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In     病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In     病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In   门诊费用记录.保险大类id%Type,
  保险项目否_In   门诊费用记录.保险项目否%Type,
  统筹金额_In     门诊费用记录.统筹金额%Type,
  摘要_In         门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In     Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In     Number := 0, --挂号是否使用收费票据
  保险编码_In     门诊费用记录.保险编码%Type,
  复诊_In         病人挂号记录.复诊%Type := 0,
  号序_In         挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In         病人挂号记录.社区%Type := Null,
  预约接收_In     Number := 0,
  预约方式_In     预约方式.名称%Type := Null,
  生成队列_In     Number := 0,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  结算卡序号_In   病人预交记录.结算卡序号%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  合作单位_In     病人预交记录.合作单位%Type := Null,
  操作类型_In     Number := 0,
  险类_In         病人挂号记录.险类%Type := Null,
  结算模式_In     Number := 0,
  记帐费用_In     Number := 0,
  退号重用_In     Number := 1,
  修正病人费别_In Number := 0
) As
  ---------------------------------------------------------------------------
  --
  --参数:
  --     操作类型_in:0-正常挂号或者预约 1-操作员拥有加号权限加号
  --     修正病人费别_In:0-不修改病人费别 1-修改病人费别
  ----------------------------------------------------------------------------
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit(v_病人id 病人信息.病人id%Type) Is
    Select *
    From (Select a.Id, a.记录状态, a.预交类别, a.No, Nvl(a.金额, 0) As 金额
           From 病人预交记录 A,
                (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                  From 病人预交记录 A
                  Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.病人id = v_病人id And Nvl(a.预交类别, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.金额, 0)) <> 0) B
           Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                 a.No = b.No And a.病人id = v_病人id And Nvl(a.预交类别, 2) = 1
           Union All
           Select 0 As ID, 记录状态, 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
           From 病人预交记录
           Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And 病人id = v_病人id And
                 Nvl(预交类别, 2) = 1 Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
           Group By 记录状态, NO, 预交类别)
    Order By ID, NO;
  --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录
  --       同时汇总相关的汇总表(病人挂号汇总、费用汇总)
  --       第一行费用处理票据使用情况(领用ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_分时段       Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况 
  n_已约数       病人挂号汇总.已约数%Type;
  n_已接收       病人挂号汇总.其中已接收%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_打印id   票据打印内容.Id%Type;
  n_费用id   门诊费用记录.Id%Type;
  n_预交金额 病人预交记录.金额%Type;
  n_当前金额 病人预交记录.金额%Type;
  n_返回值   病人预交记录.金额%Type;
  n_预交id   病人预交记录.Id%Type;
  n_消费卡id 消费卡目录.Id%Type;
  n_挂号id   病人挂号记录.Id%Type;

  n_组id           财务缴款分组.Id%Type;
  n_序号           挂号序号状态.序号%Type;
  n_已用序号       挂号序号状态.序号%Type;
  n_序号控制       挂号安排.序号控制%Type;
  n_分诊台签到排队 Number;
  n_Count          Number;
  n_限号数         Number(18);
  n_自制卡         Number;
  d_排队时间       Date;
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type := 0;
  v_星期           挂号安排限制.限制项目%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;

  n_挂出的最大序号 Number(4) := 0;
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
Begin
  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id := Zl_Get组id(操作员姓名_In);
  If 费别_In Is Null Then
    Begin
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
        Raise Err_Item;
    End;
  Else
    v_费别 := 费别_In;
    If Nvl(修正病人费别_In, 0) = 1 Then
      Begin
        Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '没有找到对应的病人！';
          Raise Err_Item;
      End;
    End If;
  End If;
  Begin
    Delete From 挂号序号状态
    Where 号码 = 号别_In And 日期 = 发生时间_In And 序号 = 号序_In And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(预约挂号_In, 0) = 0 Then
    --挂号或者预约接收
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*)
    Into n_Count
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (1, 3) And 序号 = 序号_In And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    --获取结算方式名称
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
    Begin
      Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
    Exception
      When Others Then
        v_个人帐户 := '个人帐户';
    End;
  End If;

  n_序号 := 号序_In;
  Select Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;

  --挂号获取安排
  Begin
    Select a.Id, a.序号控制, Nvl(b.限号数, 0), Nvl(b.限约数, 0)
    Into n_安排id, n_序号控制, n_限号数, n_限约数
    From 挂号安排 A, 挂号安排限制 B
    Where a.Id = b.安排id(+) And b.限制项目(+) = v_星期 And a.号码 = 号别_In;
  
  Exception
    When Others Then
      n_安排id := -1;
  End;

  --如果是病历费或者号别为空时不检查
  If Nvl(病历费_In, 0) = 0 Or 号别_In Is Not Null Then
    If n_安排id = -1 Then
      v_Err_Msg := '不存在相应的挂号安排数据,请检查';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 1 Then
    --首先获取计划
    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 安排id = n_安排id And 审核时间 Is Not Null And
            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.生效时间) As 生效
             From 挂号安排计划 A
             Where a.审核时间 Is Not Null And 发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.安排id = n_安排id) And
            发生时间_In Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'));
    
    Exception
      When Others Then
        n_计划id := 0;
    End;
    If Nvl(n_计划id, 0) <> 0 Then
      Begin
        --获取计划的限制
        Select a.Id, a.序号控制, Nvl(b.限号数, 0) As 限号数, Nvl(b.限约数, 0) As 限约数
        Into n_计划id, n_序号控制, n_限号数, n_限约数
        From 挂号安排计划 A, 挂号计划限制 B
        Where a.号码 = 号别_In And a.Id = n_计划id And a.审核时间 Is Not Null And a.Id = b.计划id(+) And b.限制项目(+) = v_星期;
      Exception
        When Others Then
          v_Err_Msg := '不存相应的挂号安排或计划数据,请检查';
          Raise Err_Item;
      End;
    End If;
  End If;

  --获取是否分时段
  If Nvl(n_计划id, 0) = 0 Then
    Select Count(Rownum) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum <= 1;
  Else
    Select Count(Rownum) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum <= 1;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    --发生时间_in>Sysdate 发生时间>最大的时段时间--号序_in is null
    Begin
      Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And Nvl(限制数量, 0) <> 0;
      n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_追加号 := 0;
    End;
  End If;
  d_时段时间 := 发生时间_In;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 Then
    --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
    Begin
      Select Nvl(序号, 0),
             To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
      Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And
            (序号, 安排id, 星期) In (Select Nvl(Max(序号), -1), 安排id, 星期
                               From 挂号安排时段
                               Where 安排id = n_安排id And 星期 = v_星期 And
                                     Decode(操作类型_In + n_追加号, 0, To_Char(发生时间_In, 'hh24:mi'),
                                            To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By 安排id, 星期);
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 1 And n_分时段 > 0 Then
    --预约号,取计划
    Begin
      If Nvl(n_计划id, 0) = 0 Then
        --没计划生效,取安排的数据
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号安排时段 C
        Where 安排id = n_安排id And 星期 = v_星期 And
              (序号, 安排id, 星期) In
              (Select Nvl(Max(c.序号), -1), 安排id, 星期
               From 挂号安排时段 C
               Where 安排id = n_安排id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 安排id, 星期);
      Else
        --有计划生效取计划
        --没生效，代表是从挂号计划时段查询      
        Select Nvl(序号, -1),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号计划时段 C
        Where 计划id = n_计划id And 星期 = v_星期 And
              (序号, 计划id, 星期) In
              (Select Nvl(Max(c.序号), -1), 计划id, 星期
               From 挂号计划时段 C
               Where 计划id = n_计划id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 计划id, 星期);
      End If;
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 Then
  
    --获取当前未使用的序号
    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>     
      Begin
        --最大序号
        If 退号重用_In = 1 Then
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
      End;
      If n_序号 Is Null Then
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end>
    
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查       
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已用数量 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd') || '已达到最大限约数！';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --启用序号控制,未分时段 加号情况   不处理,如果以后有限制条件以后补充
      End If;
    
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 0 Then
      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 Then
        --<正常预约挂号-->
        Begin
          Select Count(0) As 已挂数, Nvl(Sum(Decode(Nvl(Sign(a.日期 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
          Into n_已挂数, n_已约数
          From 挂号序号状态 A
          Where a.号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And
                状态 Not In (4, 5);
        Exception
          When Others Then
            n_已挂数 := 0;
            n_已约数 := 0;
        End;
      
        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量  
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
        If n_限号数 <= n_已挂数 Then
          v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
          Raise Err_Item;
        End If;
      End If;
    
      --没有达到时段的限号数 号码在当前时段往后追加
    
      --获取当天挂出的最大号序
      If 退号重用_In = 1 Then
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 Not In (4, 5);
      Else
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 <> 5;
      End If;
    
      --设置序号
      n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_已约数 + 1;
      If n_序号 <= Nvl(n_挂出的最大序号, 0) Then
        n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
      End If;
    
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        If 退号重用_In = 1 Then
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
          n_已挂数   := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        n_预约有效时间 := Zl_To_Number(zl_GetSysParameter('预约有效时间', 1111));
        n_失约挂号     := Zl_To_Number(zl_GetSysParameter('失约用于挂号', 1111));
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 日期), 1, 1, 0))
            Into n_失效数
            From 挂号序号状态
            Where 号码 = 号别_In And 日期 Between Trunc(Sysdate) And Sysdate And Nvl(预约, 0) = 1 And 状态 = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
        
          --挂号 追加号码时不检查时段限号数
          If n_时段限号 <= n_已用数量 And Nvl(n_追加号, 0) = 0 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 - n_失效数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --挂号
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
          If n_限约数 <= n_已用数量 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_序号 Is Null Then
        --设置序号
        If Nvl(n_已用序号, 0) < Nvl(n_时段序号, 0) Then
          n_已用序号 := Nvl(n_时段序号, 0);
        End If;
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
    Elsif Nvl(n_分时段, 0) = 0 And Nvl(n_序号控制, 0) = 0 And Nvl(病历费_In, 0) = 0 And Nvl(号别_In, 0) > 0 Then
      ---<--普通号  -->
      Begin
        Select 已挂数, 已约数
        Into n_已用数量, n_已约数
        From 病人挂号汇总
        Where 日期 = Trunc(发生时间_In) And 号码 = 号别_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;
  End If;

  --更新挂号序号状态
  If 序号_In = 1 And Not n_序号 Is Null Then
    If n_分时段 = 1 Then
      d_序号时间 := 发生时间_In;
    Else
      d_序号时间 := Trunc(发生时间_In);
    End If;
    --锁定序号的处理
    Begin
      Select 操作员姓名, 机器名
      Into v_序号操作员, v_序号机器名
      From 挂号序号状态
      Where 状态 = 5 And 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
      n_锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_锁定       := 0;
    End;
    If n_锁定 = 0 Then
      Update 挂号序号状态
      Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
      Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_分时段, 0) = 0 Or Nvl(预约挂号_In, 0) = 1 Or (Nvl(n_序号控制, 0) = 0 And Nvl(号序_In, 0) = 0) Then
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
            Values
              (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
          Elsif Nvl(n_分时段, 0) > 0 Then
            --分时段后专家号 失约的预约号允许挂号
            Update 挂号序号状态
            Set 状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In, 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
            Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 2;
            If Sql%NotFound Then
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
              Values
                (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If 操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_序号 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        Update 挂号序号状态
        Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
        Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 5 And 操作员姓名 = 操作员姓名_In And 机器名 = v_机器名;
      End If;
    End If;
  End If;

  --产生病人挂号费用(可能单独是或包括病历费用)
  Select 病人费用记录_Id.Nextval Into n_费用id From Dual; --应该通过程序得到

  Insert Into 门诊费用记录
    (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
     计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名,
     发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
  Values
    (n_费用id, 4, Decode(预约挂号_In, 1, 0, 1), 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, 票据号_In, 1, 急诊_In,
     病历费_In, Decode(预约挂号_In, 1, To_Char(n_序号), 诊室_In), Decode(病人id_In, 0, Null, 病人id_In),
     Decode(门诊号_In, 0, Null, 门诊号_In), 付款方式_In, 姓名_In, Decode(姓名_In, Null, Null, 性别_In), Decode(姓名_In, Null, Null, 年龄_In),
     v_费别, 病人科室id_In, 收费类别_In, 号别_In, 收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In,
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额_In)),
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In)), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 开单部门id_In,
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, 发生时间_In, 登记时间_In, 保险大类id_In,
     保险项目否_In, 保险编码_In, 统筹金额_In, 摘要_In, 预约方式_In, Decode(预约挂号_In, 1, Null, n_组id));

  --汇总结算到病人预交记录
  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 0 Then
  
    If (Nvl(现金支付_In, 0) <> 0 Or (Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0)) And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    
      If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
      
        n_消费卡id := Null;
        Begin
          Select Nvl(自制卡, 0), 1 Into n_自制卡, n_Count From 卡消费接口目录 Where 编号 = 结算卡序号_In;
        Exception
          When Others Then
            n_Count := 0;
        End;
        If n_Count = 0 Then
          v_Err_Msg := '没有发现原结算卡的相应类别,不能继续操作！';
          Raise Err_Item;
        End If;
        If n_自制卡 = 1 Then
          Select ID
          Into n_消费卡id
          From 消费卡目录
          Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In And
                序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In);
        End If;
        Zl_病人卡结算记录_Insert(结算卡序号_In, n_消费卡id, 结算方式_In, 现金支付_In, 卡号_In, Null, 登记时间_In, Null, 结帐id_In, n_预交id);
      End If;
    
    End If;
  
    --对于医保挂号
    If Nvl(个帐支付_In, 0) <> 0 And 序号_In = 1 Then
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_个人帐户, 个帐支付_In, 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '医保挂号', n_组id, 4);
    End If;
  
    --对于就诊卡通过预交金挂号
    If Nvl(预交支付_In, 0) <> 0 And 序号_In = 1 Then
      n_预交金额 := 预交支付_In;
      For r_Deposit In c_Deposit(病人id_In) Loop
        n_当前金额 := Case
                    When r_Deposit.金额 - n_预交金额 < 0 Then
                     r_Deposit.金额
                    Else
                     n_预交金额
                  End;
      
        If r_Deposit.Id <> 0 Then
          --第一次冲预交(82592,将第一次标上结帐ID,冲预交标记为0)
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.Id;
        End If;
        --冲上次剩余额
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 登记时间_In, 操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
        Where 病人id = 病人id_In And 性质 = 1 And 类型 = Nvl(r_Deposit.预交类别, 2);
        --检查是否已经处理完
        If r_Deposit.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
      
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      
      End If;
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --相关汇总表的处理
    --人员缴款余额
    If 序号_In = 1 Then
      If Nvl(现金支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 现金支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金)
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, Nvl(结算方式_In, v_现金), 1, 现金支付_In);
          n_返回值 := 现金支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      End If;
    
      If Nvl(个帐支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 个帐支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
          n_返回值 := 个帐支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_个人帐户 And Nvl(余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If Nvl(预约挂号_In, 0) = 0 Then
    If Nvl(记帐费用_In, 0) = 0 Then
      --处理票据使用情况
      If 序号_In = 1 And 票据号_In Is Not Null Then
        Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
      
        --发出票据
        Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, 单据号_In);
      
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
        Values
          (票据使用明细_Id.Nextval, Decode(收费票据_In, 1, 1, 4), 票据号_In, 1, 1, 领用id_In, n_打印id, 登记时间_In, 操作员姓名_In);
      
        --状态改动
        Update 票据领用记录
        Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
        Where ID = Nvl(领用id_In, 0);
      End If;
    End If;
    --病人本次就诊(以发生时间为准)
    If Nvl(病人id_In, 0) <> 0 And 序号_In = 1 Then
      Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
  
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, Nvl(实收金额_In, 0), 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 1, Nvl(实收金额_In, 0));
    End If;
  End If;

  --病人挂号记录
  If 号别_In Is Not Null And 序号_In = 1 Then
    --And Nvl(预约挂号_In, 0) = 0
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式);
    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
      
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         Null, v_排队序号);
      End If;
    End If;
  End If;
  --病人担保信息
  If 病人id_In Is Not Null And 序号_In = 1 Then
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的病人信息,不允许挂号';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Exists (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
  
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) > Sysdate;
    End If;
  End If;
  If 序号_In = 1 Then
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
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
End Zl_病人挂号记录_Insert;
/

--90021:刘硕,2015-11-06,病人地址结构化录入支持乡镇级
Create Or Replace Procedure Zl_病人地址信息_Update
(
  功能_In     Number,
  病人id_In   病人地址信息.病人id%Type,
  主页id_In   病人地址信息.主页id%Type,
  地址类别_In 病人地址信息.地址类别%Type,
  省_In       病人地址信息.省%Type := Null,
  市_In       病人地址信息.市%Type := Null,
  县_In       病人地址信息.县%Type := Null,
  乡镇_In     病人地址信息.乡镇%Type := Null,
  其他_In     病人地址信息.其他%Type := Null,
  区划代码_In 病人地址信息.区划代码%Type := Null
) Is
  --功能：首页整理中结构化病人地址信息管理 
  --参数：功能_In 1-新增,修改   2-删除 
Begin
  If 功能_In = 1 Then
    Update 病人地址信息
    Set 省 = 省_In, 市 = 市_In, 县 = 县_In, 乡镇 = 乡镇_In, 其他 = 其他_In, 区划代码 = 区划代码_In
    Where 病人id = 病人id_In And Nvl(主页id,0) = Nvl(主页id_In,0) And 地址类别 = 地址类别_In;
    If Sql%Rowcount = 0 Then
      Insert Into 病人地址信息
        (病人id, 主页id, 地址类别, 省, 市, 县, 乡镇, 其他, 区划代码)
      Values
        (病人id_In, 主页id_In, 地址类别_In, 省_In, 市_In, 县_In, 乡镇_In, 其他_In, 区划代码_In);
    End If;
  Else
    Delete From 病人地址信息 Where 病人id = 病人id_In And Nvl(主页id,0) = Nvl(主页id_In,0) And 地址类别 = 地址类别_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人地址信息_Update;
/
---90823:刘硕,2015-11-23,结构化地址增加虚拟地址识别
Create Or Replace Function Zl_Adderss_Structure(v_Addressinfo Varchar2) Return Varchar2 Is
  --返回结构：省,省编码,是否虚拟,是否不显示,是否只有虚拟级|市,市编码,是否虚拟,是否不显示,是否只有虚拟级
  --          |区县,区县编码,是否虚拟,是否不显示,是否只有虚拟级|乡镇,乡镇编码,是否虚拟,是否不显示,是否只有虚拟级
  --          |街道,街道编码,是否虚拟,是否不显示,是否只有虚拟级
  v_省       Varchar2(100);
  v_Code省   Varchar2(100);
  v_Info省   Varchar2(100);
  v_市       Varchar2(100);
  v_Code市   Varchar2(100);
  v_Info市   Varchar2(100);
  v_区县     Varchar2(100);
  v_Code区县 Varchar2(100);
  v_Info区县 Varchar2(100);
  v_乡镇     Varchar2(100);
  v_Code乡镇 Varchar2(100);
  v_Info乡镇 Varchar2(100);
  v_街道     Varchar2(100);
  v_Code街道 Varchar2(100);
  v_Info街道 Varchar2(100);
  v_Tmp      Varchar2(100);
  v_Adrstmp  Varchar2(300);
  n_Pos      Number(5);
  n_虚拟     Number(1);
  n_不显示   Number(1);
  n_Count    Number(3);
  v_Return   Varchar2(500);
Begin
  --传入结构化的地址，不用进行地址标准化分割解析
  v_Adrstmp := v_Addressinfo;
  If v_Addressinfo Like '%,%,%,%,%' Then
    n_Pos     := Instr(v_Adrstmp, ',');
    v_省      := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_市      := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_区县    := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_乡镇    := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_街道    := Substr(v_Adrstmp, n_Pos + 1);
    Select Max(编码) Into v_Code省 From 区域 Where 名称 = v_省 And Nvl(级数, 0) = 0;
    --省级地址都没有，就不做处理
    If v_Code省 Is Not Null Then
      Select Max(编码), Max(是否虚拟), Max(是否不显示)
      Into v_Code市, n_虚拟, n_不显示
      From 区域
      Where 名称 = v_市 And Nvl(级数, 0) = 1 And 上级编码 = v_Code省;
      If v_Code市 Is Not Null Then
        v_Info市 := v_市 || ',' || v_Code市 || ',' || n_虚拟 || ',' || n_不显示;
        Select Max(编码), Max(是否虚拟), Max(是否不显示)
        Into v_Code区县, n_虚拟, n_不显示
        From 区域
        Where 名称 = v_区县 And Nvl(级数, 0) = 2 And 上级编码 = v_Code市;
        --可能是虚拟地址
      Else
        Select Max(编码), Max(上级编码)
        Into v_Code区县, v_Code市
        From 区域
        Where 名称 = v_区县 And Nvl(级数, 0) = 2 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code省);
        If v_Code市 Is Not Null Then
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
          Into v_市, v_Code市, n_虚拟, n_不显示
          From 区域
          Where 编码 = v_Code市;
        End If;
        v_Info市 := v_市 || ',' || v_Code市 || ',' || n_虚拟 || ',' || n_不显示;
        Select Max(编码), Max(是否虚拟), Max(是否不显示)
        Into v_Code区县, n_虚拟, n_不显示
        From 区域
        Where 编码 = v_Code区县;
      End If;
      v_Info区县 := v_区县 || ',' || v_Code区县 || ',' || n_虚拟 || ',' || n_不显示;
      If v_Code区县 Is Not Null Then
        --可能乡镇在详细地址中，关联参数乡镇地址结构化录入
        If v_乡镇 Is Null And Not v_街道 Is Null Then
          --先截取乡镇级的两个字做关键字，来匹配
          v_Tmp := Substr(v_街道, 1, 2);
          Select Max(名称)
          Into v_乡镇
          From 区域
          Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 = v_Code区县;
          --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配
          If n_Count > 1 Then
            v_Tmp := Substr(v_街道, 1, 3);
            Select Max(名称)
            Into v_乡镇
            From 区域
            Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 = v_Code区县;
          End If;
          If Not v_乡镇 Is Null Then
            v_街道 := Substr(v_街道, Length(v_乡镇) + 1);
          End If;
        End If;
        Select Max(编码), Max(是否虚拟), Max(是否不显示)
        Into v_Code乡镇, n_虚拟, n_不显示
        From 区域
        Where 名称 = v_乡镇 And Nvl(级数, 0) = 3 And 上级编码 = v_Code区县;
        --可能是虚拟地址
        If v_Code乡镇 Is Null Then
          Select Max(编码), Max(上级编码)
          Into v_Code街道, v_Code乡镇
          From 区域
          Where 名称 = v_街道 And Nvl(级数, 0) = 4 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code区县);
          If v_Code乡镇 Is Not Null Then
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
            Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示
            From 区域
            Where 编码 = v_Code乡镇;
          End If;
        End If;
        v_Info乡镇 := v_乡镇 || ',' || v_Code乡镇 || ',' || n_虚拟 || ',' || n_不显示;
        If v_Code乡镇 Is Not Null Then
          Select Max(编码), Max(是否虚拟), Max(是否不显示)
          Into v_Code街道, n_虚拟, n_不显示
          From 区域
          Where 名称 = v_街道 And Nvl(级数, 0) = 4 And 上级编码 = v_Code乡镇;
          v_Info街道 := v_街道 || ',' || v_Code街道 || ',' || n_虚拟 || ',' || n_不显示;
        End If;
      End If;
    End If;
    --非标准地址，是完整地址，需要分割省，市，县,
  Else
    v_Adrstmp := v_Addressinfo;
    v_Tmp     := Substr(v_Adrstmp, 1, 2);
    Select Max(名称), Max(编码) Into v_省, v_Code省 From 区域 Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 0;
    --有省级地址，说明可以结构化
    If v_Code省 Is Not Null Then
      --省级地址是标准的
      If Substr(v_Adrstmp, 1, Length(v_省)) = v_省 Then
        v_Adrstmp := Substr(v_Adrstmp, Length(v_省) + 1);
        --省级地址不标准,可能新疆省略自治区等,此时，市级地址可能是标准化的。
      Else
        --先判断二级地址是否存在虚拟地址与不显示的地址
        If v_Tmp = '内蒙' Then
          v_Tmp := '内蒙古';
        Elsif v_Tmp = '黑龙' Then
          v_Tmp := '黑龙江';
        End If;
        v_Adrstmp := Substr(v_Adrstmp, Length(v_Tmp) + 1);
      End If;
      --先截取市级的两个字做关键字，来匹配
      v_Tmp := Substr(v_Adrstmp, 1, 2);
      Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1)
      Into v_市, v_Code市, n_虚拟, n_不显示, n_Count
      From 区域
      Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 1 And 上级编码 = v_Code省;
      --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配
      If n_Count > 1 Then
        v_Tmp := Substr(v_Adrstmp, 1, 3);
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
        Into v_市, v_Code市, n_虚拟, n_不显示
        From 区域
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 1 And 上级编码 = v_Code省;
      End If;
      --判断是否存在虚拟地址或不显示的地址导致的,如果存在，则根据第三级地址来确定虚拟地址
      If v_Code市 Is Null Then
        Select Max(是否虚拟), Max(是否不显示) Into n_虚拟, n_不显示 From 区域 Where 上级编码 = v_Code省;
        If Nvl(n_虚拟, 0) = 1 Or Nvl(n_不显示, 0) = 1 Then
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1), Max(上级编码)
          Into v_区县, v_Code区县, n_虚拟, n_不显示, n_Count, v_Code市
          From 区域
          Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 2 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code省);
          --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配
          If n_Count > 1 Then
            v_Tmp := Substr(v_Adrstmp, 1, 3);
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Max(上级编码)
            Into v_区县, v_Code区县, n_虚拟, n_不显示, v_Code市
            From 区域
            Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 2 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code省);
          End If;
          If v_Code市 Is Not Null Then
            v_Info区县 := v_区县 || ',' || v_Code区县 || ',' || n_虚拟 || ',' || n_不显示;
            v_Adrstmp  := Substr(v_Adrstmp, Length(v_区县) + 1);
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
            Into v_市, v_Code市, n_虚拟, n_不显示
            From 区域
            Where 编码 = v_Code市;
            v_Info市 := v_市 || ',' || v_Code市 || ',' || n_虚拟 || ',' || n_不显示;
          End If;
        End If;
      Else
        v_Info市  := v_市 || ',' || v_Code市 || ',' || n_虚拟 || ',' || n_不显示;
        v_Adrstmp := Substr(v_Adrstmp, Length(v_市) + 1);
      End If;
      --没有区县，则解析区县
      If Not v_Code市 Is Null And v_Code区县 Is Null Then
        --先截取县级的两个字做关键字，来匹配
        v_Tmp := Substr(v_Adrstmp, 1, 2);
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1)
        Into v_区县, v_Code区县, n_虚拟, n_不显示, n_Count
        From 区域
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 2 And 上级编码 = v_Code市;
        --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
          Into v_区县, v_Code区县, n_虚拟, n_不显示
          From 区域
          Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 2 And 上级编码 = v_Code市;
        End If;
        If v_Code区县 Is Null Then
          Select Max(是否虚拟), Max(是否不显示) Into n_虚拟, n_不显示 From 区域 Where 上级编码 = v_Code市;
          If Nvl(n_虚拟, 0) = 1 Or Nvl(n_不显示, 0) = 1 Then
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1), Max(上级编码)
            Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示, n_Count, v_Code区县
            From 区域
            Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code市);
            --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配
            If n_Count > 1 Then
              v_Tmp := Substr(v_Adrstmp, 1, 3);
              Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Max(上级编码)
              Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示, v_Code区县
              From 区域
              Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code市);
            End If;
          
            If v_Code乡镇 Is Not Null Then
              v_Info乡镇 := v_乡镇 || ',' || v_Code乡镇 || ',' || n_虚拟 || ',' || n_不显示;
              v_Adrstmp  := Substr(v_Adrstmp, Length(v_乡镇) + 1);
              Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
              Into v_区县, v_Code区县, n_虚拟, n_不显示
              From 区域
              Where 编码 = v_Code区县;
              v_Info区县 := v_区县 || ',' || v_Code区县 || ',' || n_虚拟 || ',' || n_不显示;
            End If;
          End If;
        Else
          v_Info区县 := v_区县 || ',' || v_Code区县 || ',' || n_虚拟 || ',' || n_不显示;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_区县) + 1);
        End If;
      End If;
      If v_Code区县 Is Not Null And v_Code乡镇 Is Null Then
        --先截取乡镇级的两个字做关键字，来匹配
        v_Tmp := Substr(v_Adrstmp, 1, 2);
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1)
        Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示, n_Count
        From 区域
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 = v_Code区县;
        --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1)
          Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示, n_Count
          From 区域
          Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 = v_Code区县;
        End If;
        If v_Code乡镇 Is Null Then
          Select Max(是否虚拟), Max(是否不显示) Into n_虚拟, n_不显示 From 区域 Where 上级编码 = v_Code区县;
          If Nvl(n_虚拟, 0) = 1 Or Nvl(n_不显示, 0) = 1 Then
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1), Max(上级编码)
            Into v_街道, v_Code街道, n_虚拟, n_不显示, n_Count, v_Code乡镇
            From 区域
            Where 名称 = v_Adrstmp And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code区县);
          End If;
          If v_Code乡镇 Is Not Null Then
            v_Info街道 := v_街道 || ',' || v_Code街道 || ',' || n_虚拟 || ',' || n_不显示;
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
            Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示
            From 区域
            Where 编码 = v_Code乡镇;
            v_Info乡镇 := v_乡镇 || ',' || v_Code乡镇 || ',' || n_虚拟 || ',' || n_不显示;
          End If;
        Else
          v_Info乡镇 := v_乡镇 || ',' || v_Code乡镇 || ',' || n_虚拟 || ',' || n_不显示;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_乡镇) + 1);
        End If;
        If v_Code乡镇 Is Not Null And v_Code街道 Is Null Then
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
          Into v_街道, v_Code街道, n_虚拟, n_不显示
          From 区域
          Where 名称 = v_Adrstmp And Nvl(级数, 0) = 4 And 上级编码 = v_Code乡镇;
          If v_Code街道 Is Not Null Then
            v_Info街道 := v_街道 || ',' || v_Code街道 || ',' || n_虚拟 || ',' || n_不显示;
          End If;
        End If;
      End If;
    End If;
    If v_街道 Is Null Then
      v_街道 := v_Adrstmp;
    End If;
  End If;
  v_Info省 := v_省 || ',' || v_Code省 || ',,,';
  If v_Info市 Is Null Then
    v_Info市 := v_市 || ',,,';
  End If;
  --只有省没有市，判断市是否只有虚拟级
  If Not v_Code省 Is Null And v_市 Is Null Then
    Select Count(1)
    Into n_Count
    From 区域
    Where 上级编码 = v_Code省 And Nvl(是否虚拟, 0) = 0 And Nvl(是否不显示, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From 区域 Where 上级编码 = v_Code省 And Rownum < 2;
      If n_Count = 0 Then
        v_Info市 := v_Info市 || ',';
      Else
        v_Info市 := v_Info市 || ',1';
      End If;
    Else
      v_Info市 := v_Info市 || ',';
    End If;
  Else
    v_Info市 := v_Info市 || ',';
  End If;
  If v_Info区县 Is Null Then
    v_Info区县 := v_区县 || ',,,';
  End If;
  --只有市没有区县，判断区县只有虚拟级
  If Not v_Code市 Is Null And v_区县 Is Null Then
    Select Count(1)
    Into n_Count
    From 区域
    Where 上级编码 = v_Code市 And Nvl(是否虚拟, 0) = 0 And Nvl(是否不显示, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From 区域 Where 上级编码 = v_Code市 And Rownum < 2;
      If n_Count = 0 Then
        v_Info区县 := v_Info区县 || ',';
      Else
        v_Info区县 := v_Info区县 || ',1';
      End If;
    Else
      v_Info区县 := v_Info区县 || ',';
    End If;
  Else
    v_Info区县 := v_Info区县 || ',';
  End If;
  If v_Info乡镇 Is Null Then
    v_Info乡镇 := v_乡镇 || ',,,';
  End If;
  --只有区县没有乡镇，判断乡镇是否只有虚拟的下级
  If Not v_Code区县 Is Null And v_乡镇 Is Null Then
    Select Count(1)
    Into n_Count
    From 区域
    Where 上级编码 = v_Code区县 And Nvl(是否虚拟, 0) = 0 And Nvl(是否不显示, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From 区域 Where 上级编码 = v_Code区县 And Rownum < 2;
      If n_Count = 0 Then
        v_Info乡镇 := v_Info乡镇 || ',';
      Else
        v_Info乡镇 := v_Info乡镇 || ',1';
      End If;
    Else
      v_Info乡镇 := v_Info乡镇 || ',';
    End If;
  Else
    v_Info乡镇 := v_Info乡镇 || ',';
  End If;
  If v_Info街道 Is Null Then
    v_Info街道 := v_街道 || ',,,,';
  Else
    v_Info街道 := v_Info街道 || ',';
  End If;
  v_Return := v_Info省 || '|' || v_Info市 || '|' || v_Info区县 || '|' || v_Info乡镇 || '|' || v_Info街道;
  Return(v_Return);
End;
/
--90046:刘尔旋,2015-11-17,支付宝结帐修改
--90317:张永康,2015-11-05,历史数据转出对零费记帐数据的转出修正
Create Or Replace Procedure Zl1_Datamove_Tag
(
  d_End    In Date,
  n_批次   In Number,
  n_System In Number
) As
  --功能：标记待转出的数据 
  --说明：为避免Undo表空间膨胀过大，分段提交 
Begin
  --1.经济核算（费用,药品,收款和票据等）  
  --新加子查询注意性能优化，把能够将数据过滤到最小的条件放到最后，Exists类条件放前面
  Update /*+ rule*/ 病人预交记录 L
  Set 待转出 = n_批次
  Where 结帐id In
        (Select Distinct a.结帐id --1.门诊收费和挂号的收费结算记录(排除之后退号和退费的,一张单据中只要其中一行退了) 
         From 门诊费用记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 门诊费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_End))
     And a.待转出 Is Null And a.记录性质 In (1, 4) And a.登记时间 < d_End
         Union All
         Select Distinct a.结算id --2.医保补结算 
         From 费用补充记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 费用补充记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 In (1, 2) And b.登记时间 >= d_End))
     And a.待转出 Is Null And a.记录性质 = 1 And a.登记时间 < d_End
         Union All
         Select Distinct a.结帐id --3.就诊卡的收费结算记录(排除之后退卡费的,一张单据中只要其中一行退了) 
         From 住院费用记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 住院费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_End))
    And a.待转出 Is Null And a.记帐费用 = 0 And a.记录性质 = 5 And a.登记时间 < d_End
         Union All --4.门诊(记帐单)和住院的结帐结算记录 
         Select 结帐id
         From (With Settle As (Select Distinct a.Id As 结帐id, a.病人id --3.门诊(记帐单)和住院的结帐结算记录(排除之后结帐作废的) 
                               From 病人结帐记录 A
                               Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                                      (Select 1 From 病人结帐记录 B Where a.No = b.No And b.记录状态 = 2 And b.收费时间 >= d_End))
              And a.待转出 Is Null And a.收费时间 < d_End)
                Select 结帐id
                From Settle
                Minus
                --以下结帐ID要整体排除,避免部分费用明细被转出后影响后续的计算是否冲完 
                --1.一张预交款被多笔结帐冲完（结帐ID不同）
                --2.费用单据的结帐ID相关的可能还有其他NO的其他结帐ID(结帐作废后分多次结帐结清，可能部分在转出时间之后)
                --考虑到这情况的复杂性，为简化逻辑，提升查询性能，按病人ID来排除 
                Select Distinct d.Id
                From 病人结帐记录 D,
                     (Select Distinct c.病人id --多次住院可以一起结，以及门诊记帐和住院记帐可以一起结且冲同一笔预交，所以这里不加主页ID 
                       From 住院费用记录 C,
                            (Select Distinct d.No, d.序号, Mod(d.记录性质, 10) As 记录性质
                              From 住院费用记录 D,
                                   (Select s.结帐id From Settle S, 病人结帐记录 E --没有结清且该病人之后没有再结过就成了呆帐，这种就不排除 
                                     Where s.病人id = e.病人id And (e.收费时间 > d_End Or Exists (Select 1 From 在院病人 F Where s.病人id = f.病人id))) S 
                              Where d.结帐id = s.结帐id) D
                       Where c.No = d.No And Mod(c.记录性质, 10) = d.记录性质 And c.序号 = d.序号 --结帐后作废后，再对包含记帐单销帐的结帐ID为空的记录,一起汇总计算是否结清,这种结帐ID为空的数据转出在后面单独转出 
                       Group By c.No, Mod(c.记录性质, 10), c.病人id --一张单据中的一行可部分结帐，以单据为对象来判断，避免一张单据的其中一部分被转出 
                       Having Nvl(Sum(c.实收金额), 0) <> Nvl(Sum(c.结帐金额), 0) Or Exists (Select 1 --排除转出时间之后再次结帐的(作废后再次结帐)，避免原始单据转走后，后续结帐时无法正确判断 
                                                                                   From 住院费用记录 E, 病人结帐记录 S
                                                                                   Where e.No = c.No And Mod(e.记录性质, 10) = Mod(c.记录性质, 10) And
                                                                                         e.记录性质 In (12, 13, 15) And e.结帐id = s.Id  And s.待转出 Is Null And s.收费时间 >= d_End)
                       Union All
                       Select Distinct c.病人id
                       From 门诊费用记录 C,
                            (Select Distinct d.No, d.序号, Mod(d.记录性质, 10) As 记录性质
                              From 门诊费用记录 D, Settle S
                              Where d.结帐id = s.结帐id) D --因为是门诊病人，所以，只要没有结清,该病人的都不转出 
                       Where c.No = d.No And Mod(c.记录性质, 10) = d.记录性质 And c.序号 = d.序号
                       Group By c.No, Mod(c.记录性质, 10), c.病人id
                       Having Nvl(Sum(c.实收金额), 0) <> Nvl(Sum(c.结帐金额), 0) Or Exists (Select 1
                                                                                   From 门诊费用记录 E, 病人结帐记录 S
                                                                                   Where e.No = c.No And Mod(e.记录性质, 10) = Mod(c.记录性质, 10) And
                                                                                         e.记录性质 In (12, 13, 15) And e.结帐id = s.Id And s.待转出 Is Null And s.收费时间 >= d_End)) N
                Where d.病人id = n.病人id)
         );

  --排除预交款未冲完的
  --为了降低逻辑的复杂性，不排除在转出时间之后发药或未发药的费用记录对应的结帐ID，将这种情况的结算数据和费用数据强制转走 
  --因为前面的SQL查出的结帐ID可能不全是冲预交的(门诊收费和住院结帐补费等)，所以，需要单独一个SQL来排除 
  --由于可能存在数据异常(住院费用结帐冲预交类别为1的门诊预交)，所以没有加预交类别条件限定 
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = Null
  Where 待转出 = n_批次 And
        结帐id In (Select Distinct d.结帐id
                 From 病人预交记录 D,
                      --连接D表是为了查冲同一预交单据的其他结帐ID（退预交款，冲预交作废的，再次冲同一预交单据） 
                      --该预交或冲预交单据涉及的所有结帐ID的都不转出，避免部分冲预交的结帐ID被排除后，原始预交单被转走，或者其他结帐ID将费用单据的一部分(原始结帐、结帐作废、再次结一部分、再次结全部)转走 
                      (Select Distinct l.No
                        From 病人预交记录 L, 病人预交记录 P --可能本次结帐冲的只是剩余款，所以需要连接L表，查原始交预交的单据，以及记录性质为11的可能还有转出时间之后其他冲剩余款的结帐ID 
                        Where l.记录性质 = p.记录性质 And l.No = p.No And p.记录性质 In (1, 11) And p.待转出 = n_批次
                        Group By l.No, l.病人id
                        Having Nvl(Sum(l.金额), 0) <> Nvl(Sum(l.冲预交), 0) And (Exists (Select 1
                                                                                  From 病人预交记录 E --没有冲完且之后没有再冲过或结算过就成了呆帐（以及存在用负的结帐补款来表示冲预交当成冲完的清况），这种就不排除
                                                                                  Where l.病人id = e.病人id And e.待转出 Is Null And e.收款时间 > d_End)
                                                                                  Or Exists (Select 1 From 在院病人 E Where l.病人id =e.病人id)
                                                                                  Or Exists (Select 1 From 病人未结费用 E Where l.病人id =e.病人id))  
                        Or Nvl(Sum(l.金额), 0) = Nvl(Sum(l.冲预交), 0) And Exists (Select 1
                                                                                  From 病人预交记录 E --排除转出时间之后的其他结帐ID冲的,10.34.20后，冲预交全部单独增加了一条记录，收费时间就是冲预交时间(以前是在原始交预交款的记录上填冲预交字段，不能直接查到冲预交款的时间)
                                                                                  Where e.No = l.No And e.记录性质 = 11 And e.待转出 Is Null And e.收款时间 >= d_End)) N
                 Where d.No = n.No And d.记录性质 In (1, 11));

  --预交款没有使用就直接退了的记录(结帐ID为空) 
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 记录性质 = 1 And
        NO In (Select a.No
               From 病人预交记录 A
               Where a.结帐id Is Null And a.记录性质 = 1 And a.记录状态 In (2, 3) And a.待转出 Is Null And a.收款时间 < d_End
               Group By a.No
               Having Sum(a.金额) = 0);

  --冲预交款作废的记录（记录性质为2），没有结帐ID 
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 结帐id Is Null And 记录性质 = 2 And NO In (Select a.No From 病人预交记录 A Where a.待转出 = n_批次 And a.记录性质 = 3);

  Update Zldatamovelog
  Set 当前进度 = '(1/10)结算数据标记完成，正在标记费用数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 病人结帐记录
  Set 待转出 = n_批次
  Where ID In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  --结帐无结算的记录(为了提升性能，不判断费用，只要结了帐且无预交记录就当成是零费用结帐) 
  Update /*+ rule*/ 病人结帐记录 L
  Set 待转出 = n_批次
  Where 收费时间 < d_End And 待转出 Is Null And Not Exists (Select 1 From 病人预交记录 P Where l.Id = p.结帐id);

  Update /*+ rule*/ 病人卡结算对照
  Set 待转出 = n_批次
  Where 预交id In (Select a.Id From 病人预交记录 A Where 待转出 = n_批次);

  Update /*+ rule*/ 病人卡结算记录
  Set 待转出 = n_批次
  Where ID In (Select 卡结算id From 病人卡结算对照 Where 待转出 = n_批次);

  Update /*+ rule*/ 三方结算交易
  Set 待转出 = n_批次
  Where 交易id In (Select a.Id From 病人预交记录 A Where 待转出 = n_批次);

  Update /*+ rule*/ 三方退款信息
  Set 待转出 = n_批次
  Where (记录id,结帐ID) In (Select a.Id,A.结帐ID From 病人预交记录 A Where 待转出 = n_批次);

  --1.挂号打折后实收金额为0的(没有对应的预交记录),即使之后有退号费用也不管，因为金额为零不影响计算),而卡费即使为零也有预交记录 
  --结帐ID为空的是异常数据（德阳医院仅有3笔此类数据）
  --根据挂号记录再找门诊费用，比直接按时间查门诊费用要快 
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where NO In (Select NO From 病人挂号记录 Where 待转出 Is Null And 登记时间 < d_End) And 记录性质 = 4 And (实收金额 = 0 Or 结帐id Is Null);

  --2.直接收费的和结帐无结算（预交）记录的，Union不加all去掉重复以减少in的数量 
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where 结帐id In
        (Select 结帐id From 病人预交记录 Where 待转出 = n_批次 Union Select ID From 病人结帐记录 Where 待转出 = n_批次);

  --3.没有结帐id的数据(按登记时间)
  --1)未结帐的门诊记帐费用(赖账)，该病人没有预交记录或冲预交记录，并且该时间之后无门诊费用发生
  --2)未结帐的划价记录
  --3)未收费（也没有冲预交）的零费用
  --加条件"待转出 Is Null"是为了处理连续多次标记转出的情况 
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (Not Exists (Select 1 From 病人预交记录 B Where a.病人id = b.病人id And b.待转出 Is Null And 记录性质 In (1, 11)) And Not Exists
         (Select 1 From 门诊费用记录 B Where a.病人id = b.病人id And b.待转出 Is Null And 登记时间 > d_End) And 记录性质 = 2 Or 记录状态 = 0 Or
         记录性质 = 1 And 实收金额 = 0 And 结帐金额 = 0) And 结帐id Is Null And 待转出 Is Null And 登记时间 < d_End;

  --4.没有结帐id的数据(按发生时间)
  --冲销产生的记帐记录（记录状态为2），登记时间可能在当前指定转出时间之后，而原始记帐记录（记录状态为3），登记时间在指定转出时间之前。前后两者的发生时间是相同的。
  --1)未结帐的零记帐费用或打折后实收金额为零的（结帐模块参数没有勾选对零费用结帐）
  --2)结帐作废后，记帐单销帐的记录（结帐ID为空且记录状态为2的），记录状态为3的且有结帐ID的在最前面已转出. 
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (Exists (Select 1
                 From 门诊费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null And
                       b.待转出 + 0 = n_批次) And 记录状态 = 2 Or Exists
         (Select 1
          From 门诊费用记录 B
          Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.结帐id Is Null
          Group By b.No, b.记录性质, b.序号
          Having Nvl(Sum(b.实收金额), 0) = 0)) And 记录性质 = 2 And 结帐id Is Null And 待转出 Is Null And 发生时间 < d_End;

  --5.有结帐id的零费用(按发生时间)
  --按费别打折后结帐金额为零的收费记录,或者一张单据相同结帐ID的结帐金额之和为0(冲销后为零)
  --即使在转出时间之后发药的，也强制转出（为了减少逻辑复杂性，提高查询性能）
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (结帐金额 = 0 Or Exists
         (Select 1 From 门诊费用记录 C Where a.结帐id = c.结帐id Group By c.结帐id, c.No Having Sum(c.结帐金额) = 0)) And Not Exists
   (Select 1 From 病人预交记录 B Where a.结帐id = b.结帐id And b.待转出 Is Null) And 记录性质 = 1 And 结帐id Is Not Null And
        待转出 Is Null And 发生时间 < d_End;


  Update /*+ rule*/ 医保结算明细
  Set 待转出 = n_批次
  Where 结帐id In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 费用补充记录
  Set 待转出 = n_批次
  Where 结算id In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 凭条打印记录
  Set 待转出 = n_批次
  Where (NO, 记录性质) In (Select NO, 记录性质 From 门诊费用记录 Where 待转出 = n_批次);


  --1.从预交记录读是为了取就诊卡直接收费的（无结帐ID）,再加结帐记录是为了取结帐无结算（预交）记录的 
  Update /*+ rule*/ 住院费用记录
  Set 待转出 = n_批次
  Where 结帐id In
        (Select 结帐id From 病人预交记录 Where 待转出 = n_批次 Union Select ID From 病人结帐记录 Where 待转出 = n_批次);

  --2.没有结帐id的数据(按发生时间)
  --冲销产生的记帐记录（记录状态为2），登记时间可能在当前指定转出时间之后，而原始记帐记录（记录状态为3），登记时间在指定转出时间之前。前后两者的发生时间是相同的。
  --1)转出结帐作废后，记帐单销帐的记录（记帐状态为2且没有结帐ID且(记录状态为3的有结帐ID的)在最前面已转出） 
  --2)未结帐的零费用(已冲销的记帐单或打折后实收金额为零) 
  --3)没有结帐ID的划价记录处理为转出 
  
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where ((Exists (Select 1
                  From 住院费用记录 B
                  Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null And
                        b.待转出 + 0 = n_批次) And 记录状态 = 2 Or Exists
         (Select 1
           From 住院费用记录 B
           Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.结帐id Is Null
           Group By b.No, b.记录性质, b.序号
           Having Nvl(Sum(b.实收金额), 0) = 0)) And 记录性质 = 2 Or 记录状态 = 0) And 结帐id Is Null And 待转出 Is Null And 发生时间 < d_End;

  --3.离院未结帐的（赖帐病人），因为是很久以前的这些数据，如果预交已冲完，则处理为要转出 
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where 待转出 Is Null And 结帐id Is Null And
        (病人id, 主页id) In (Select 病人id, 主页id
                         From 病案主页 C
                         Where 出院日期 < d_End And 待转出 Is Null And 数据转出 Is Null And Not Exists
                          (Select 1
                                From 病人预交记录 B
                                Where b.病人id = c.病人id And b.待转出 Is Null And b.预交类别 = 2 And b.记录性质 In (1, 11)
                                Having Nvl(Sum(b.金额), 0) - Nvl(Sum(b.冲预交), 0) <> 0));

  Update Zldatamovelog
  Set 当前进度 = '(2/10)费用数据标记完成，正在标记药品数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ Rule*/ 药品收发记录 A
  Set 待转出 = n_批次
  Where Rowid In (Select m.Rowid
                  From 药品收发记录 M, 门诊费用记录 E
                  Where m.费用id = e.Id And (e.记录性质 = 1 And m.单据 In (8, 24) Or e.记录性质 = 2 And m.单据 In (9, 25)) And
                        e.收费类别 In ('4', '5', '6', '7') And e.待转出 = n_批次
                  Union All
                  Select m.Rowid
                  From 药品收发记录 M, 住院费用记录 E
                  Where m.费用id = e.Id And m.单据 In (9, 10, 25, 26) And e.记录性质 = 2 And e.收费类别 In ('4', '5', '6', '7') And
                        e.待转出 = n_批次);

  Update /*+ rule*/ 收发记录补充信息
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 输液配药内容
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 输液配药记录
  Set 待转出 = n_批次
  Where ID In (Select 记录id From 输液配药内容 Where 待转出 = n_批次);

  Update /*+ rule*/ 输液配药附费
  Set 待转出 = n_批次
  Where 配药id In (Select ID From 输液配药记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 输液配药状态
  Set 待转出 = n_批次
  Where 配药id In (Select ID From 输液配药记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品留存计划
  Set 待转出 = n_批次
  Where 留存id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品签名明细
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品签名记录
  Set 待转出 = n_批次
  Where ID In (Select 签名id From 药品签名明细 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(3/10)药品数据标记完成，正在标记缴款与票据数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 人员借款记录 Set 待转出 = n_批次 Where 待转出 Is Null And 借出时间 < d_End;

  Update /*+ rule*/ 人员收缴记录 Set 待转出 = n_批次 Where 待转出 Is Null And 登记时间 < d_End;

  Update /*+ rule*/ 人员收缴对照
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员收缴明细
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员收缴票据
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员暂存记录
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员暂存记录 Set 待转出 = n_批次 Where 待转出 Is Null And 记录性质 = 1 And 登记时间 < d_End;

  Update /*+ rule*/ 票据领用记录 A
  Set 待转出 = n_批次
  Where Not Exists
   (Select 1 From 票据使用明细 B Where b.领用id = a.Id And b.使用时间 >= d_End) And 待转出 Is Null And 剩余数量 = 0 And 登记时间 < d_End;

  Update /*+ rule*/ 票据使用明细
  Set 待转出 = n_批次
  Where 领用id In (Select ID From 票据领用记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 票据打印内容
  Set 待转出 = n_批次
  Where ID In (Select 打印id From 票据使用明细 Where 待转出 = n_批次);

  Update /*+ rule*/ 票据打印明细
  Set 待转出 = n_批次
  Where 使用id In (Select ID From 票据使用明细 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(4/10)缴款与票据数据标记完成，正在标记就诊及诊治数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --2.就诊及诊治数据 
  --不转出的条件：挂号费用未转出的，转出时间之后存在医嘱，医嘱对应的费用未转出的 
  --即使正在就诊(r.执行状态 <> 2 )的也强制转出 
  Update /*+ rule*/ 病人挂号记录 T
  Set 待转出 = n_批次
  Where Rowid In
        (Select Rowid
         From 病人挂号记录 R
         Where Not Exists (Select 1
                From 门诊费用记录 A
                Where r.No = a.No And a.登记时间 < d_End And a.记录性质 = 4 And a.待转出 Is Null) And Not Exists
          (Select 1
                From 病人医嘱记录 A
                Where a.挂号单 = r.No And a.待转出 Is Null And a.病人来源 <> 4 And Nvl(a.停嘱时间, a.开嘱时间) >= d_End) And Not Exists
          (Select 1
                From 门诊费用记录 E, 病人医嘱记录 A
                Where r.No = a.挂号单 And a.Id = e.医嘱序号 And a.病人来源 <> 4 And e.待转出 Is Null) And r.待转出 Is Null And
               r.登记时间 < d_End);

  --由于有一部分挂号数据未转出，所以，汇总表的数据可能与挂号数据不匹配 
  Update 病人挂号汇总 Set 待转出 = n_批次 Where 待转出 Is Null And 日期 < d_End;
  Update /*+ rule*/ 病人转诊记录 Set 待转出 = n_批次 Where NO In (Select NO From 病人挂号记录 Where 待转出 = n_批次);

  --通过"住院费用记录"来查询，而不是"病人结帐记录",因为离院未结的赖帐病人也转出了费用 
  --出院日期条件仍然需要，因为可能某次结帐转出了，但病人当时并未出院(一次住院多次结帐)。 
  --通过指定索引方式进行特殊优化（缺省采用"病案主页IX_出院日期"索引的效率太低） 
  Update /*+ rule*/ 病案主页 P
  Set 待转出 = n_批次
  Where Not Exists (Select 1 From 住院费用记录 A Where a.病人id = p.病人id And a.主页id = p.主页id And a.待转出 Is Null) And 待转出 Is Null And
        数据转出 Is Null And 出院日期 < d_End And
        (病人id, 主页id) In (Select Distinct 病人id, 主页id From 住院费用记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人过敏记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, ID
                         From 病人挂号记录
                         Where 待转出 = n_批次
                         Union All
                         Select 病人id, 主页id
                         From 病案主页
                         Where 待转出 = n_批次);

  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, ID
                         From 病人挂号记录
                         Where 待转出 = n_批次
                         Union All
                         Select 病人id, 主页id
                         From 病案主页
                         Where 待转出 = n_批次);

  Update /*+ rule*/ 病人手麻记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, ID
                         From 病人挂号记录
                         Where 待转出 = n_批次
                         Union All
                         Select 病人id, 主页id
                         From 病案主页
                         Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(5/10)就诊及诊治数据标记完成，正在标记护理数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --3.护理数据 
  Update /*+ rule*/ 病人护理文件
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理数据
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理明细
  Set 待转出 = n_批次
  Where 记录id In (Select ID From 病人护理数据 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人护理打印
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理活动项目
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理要素内容
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 产程要素内容
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);

  --老版护理系统数据 
  Update /*+ rule*/ 病人护理记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理内容
  Set 待转出 = n_批次
  Where 记录id In (Select ID From 病人护理记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(6/10)护理数据标记完成，正在标记病历数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --4.病历数据 
  Update /*+ rule*/ 电子病历记录
  Set 待转出 = n_批次
  Where 病人来源 <> 4 And (病人id, 主页id) In (Select 病人id, ID
                                       From 病人挂号记录
                                       Where 待转出 = n_批次
                                       Union All
                                       Select 病人id, 主页id
                                       From 病案主页
                                       Where 待转出 = n_批次);

  --自登记类病人(无挂号单号) 
  --病历ID可能重复是因为检验报告之类的，如肝功、肾功共打一张报告，即在病人医嘱报告表中，多个医嘱id对应同一报告ID 
  --为提升性能，不从医嘱发送记录的发送时间查询，不采用精确的时间，因为直接登记的检验医嘱，一般开嘱时间与发送时间相差不大
  Update /*+ rule*/ 电子病历记录
  Set 待转出 = N_批次
  Where ID In (Select C.病历id
             From 病人医嘱记录 B, 病人医嘱报告 C
             Where C.医嘱id = B.Id And Nvl(B.主页id, 0) = 0 And B.挂号单 Is Null And B.相关id Is Null And B.待转出 Is Null And
                   B.开嘱时间 < d_End);

  Update /*+ rule*/ 电子病历附件
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 电子病历格式
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 电子病历内容
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 电子病历图形
  Set 待转出 = n_批次
  Where 对象id In (Select ID From 电子病历内容 Where 待转出 = n_批次 And 对象类型 = 5);

  Update /*+ rule*/ 病人医嘱报告
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像报告驳回
  Set 待转出 = n_批次
  Where (医嘱id, 病历id) In (Select 医嘱id, 病历id From 病人医嘱报告 Where 待转出 = n_批次);

  Update /*+ rule*/ 报告查阅记录
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 疾病申报记录
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(7/10)病历数据标记完成，正在标记临床路径数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --5.临床路径 
  Update /*+ rule*/ 病人临床路径
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人合并路径
  Set 待转出 = n_批次
  Where 首要路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人合并路径评估
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人出径记录
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人路径执行
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径评估
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径变异
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径指标
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径医嘱
  Set 待转出 = n_批次
  Where 路径执行id In (Select ID From 病人路径执行 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(8/10)临床路径数据标记完成，正在标记医嘱数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --6.医嘱，检验，检查 
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where 挂号单 In (Select NO From 病人挂号记录 Where 待转出 = n_批次) And 病人来源 <> 4;
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  --自登记类病人(无挂号单)，病人医嘱报告在前面转病历时已转出 
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where Rowid In (Select b.Rowid
                  From 病人医嘱记录 B, 病人医嘱报告 C
                  Where (b.相关id = c.医嘱id Or b.Id = c.医嘱id) And c.待转出 = n_批次);

  Update /*+ rule*/ 病人医嘱计价
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人医嘱附费
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人医嘱附件
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 输血申请记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 输血检验结果
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人医嘱执行
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人医嘱打印
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 医嘱执行打印
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人诊断医嘱
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where ID In (Select 诊断id From 病人诊断医嘱 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人医嘱状态
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 医嘱签名记录
  Set 待转出 = n_批次
  Where ID In (Select 签名id From 病人医嘱状态 Where 待转出 = n_批次 And 签名id Is Not Null);

  Update /*+ rule*/ 病人医嘱发送
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 诊疗单据打印
  Set 待转出 = n_批次
  Where (NO, 记录性质) In (Select NO, 记录性质 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 医嘱执行时间
  Set 待转出 = n_批次
  Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 医嘱执行计价
  Set 待转出 = n_批次
  Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 执行打印记录
  Set 待转出 = n_批次
  Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(9/10)医嘱数据标记完成，正在标记检查检验数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 影像检查记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像报告记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像报告操作记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像检查序列
  Set 待转出 = n_批次
  Where 检查uid In (Select 检查uid From 影像检查记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像检查图象
  Set 待转出 = n_批次
  Where 序列uid In (Select 序列uid From 影像检查序列 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像申请单图像
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像收藏内容
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像危急值记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(10/10)影像数据标记完成，正在标记检验数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 检验标本记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验申请项目
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验项目分布
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验分析记录
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验质控记录
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验操作记录
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验签名记录
  Set 待转出 = n_批次
  Where 检验标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验图像结果
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验试剂记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验拒收记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验普通结果
  Set 待转出 = n_批次
  Where 检验标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验质控报告
  Set 待转出 = n_批次
  Where 结果id In (Select ID From 检验普通结果 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验药敏结果
  Set 待转出 = n_批次
  Where 细菌结果id In (Select ID From 检验普通结果 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验流水线标本
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验流水线指标
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/

--90301:许华峰,2015-11-05,历史报告信息自定义分组显示
--影像报告范文管理(---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_RptPluginOriginal Is
  Type t_Refcur Is Ref Cursor;

  -- 功    能：获取历史报告记录
  Procedure p_GetReportHistory(
    Val                   Out t_Refcur,
    医嘱id_In             In 病人医嘱记录.ID%Type,
    人员id_In             In 部门人员.人员id%Type,
    当前科室id_In         In 部门人员.部门ID%Type,
    查看其他科历史报告_In In number := 0
	);

  --功    能：获取对应报告内容
  Procedure p_GetReportContent(
    Val           Out t_Refcur,
    报告ID_In     In varchar2,
    EditorType_In Number := 0 --0:PACS报告编辑器，1--电子病历编辑器，2--报告文档编辑器
    );
  --功    能：根据医嘱ID获取检查UID
  Procedure p_GetStudyUid(
    Val       Out t_Refcur,
	医嘱id_In In 影像检查记录.医嘱id%Type
	);
  --功    能：获取预览图像总数
  Procedure p_GetStudyImageCount(
    Val Out t_Refcur, 
    查询条件_In In varchar2,
    是否临时_In In number:=0
	);
  --功    能：获取预览图像数据
  Procedure p_GetStudyImageData(
    Val         Out t_Refcur,
    查询方式_In In varchar2,
    查询条件_In In varchar2,
    开始位置_In In number,
    结束位置_In In number,
    是否临时_In In number
	);
  --功能：获取临时图像序列
  Procedure p_Get_TempImageSeries(
    Val         Out t_Refcur,
    时间范围_In In Number,
    姓名_In In 影像临时记录.姓名%Type:=null
	);
  
  --功能;获取图像备注
  procedure P_Get_NormalNote(
    Val         Out t_Refcur
	);
  --功能：插入常用图像备注
  Procedure p_Insert_Normalnote(
    note_in in 影像字典内容.名称%Type,
	code_In 影像字典内容.简码%Type
	);
  --功能：修改常用图像备注
  Procedure p_Edit_Normalnote(
    note_in In 影像字典内容.名称%Type,
	num_In  影像字典内容.编号%Type
	);
  --功能：删除常用图像备注
  Procedure p_Del_Normalnote(
    num_In 影像字典内容.编号%Type
	);

  --功能：获取备注的下一个编码
  Procedure p_Get_NormalNum(
    Val Out t_Refcur
	);
  --功能：获取插件ID
  Procedure p_Get_PlugID(
    Val     Out t_Refcur,
	类名_In In 影像报告插件.类名%Type
	);
  --功能：插入编辑器字体参数
  Procedure p_SetFontParam(
    font_In nvarchar2, 
	user_In nvarchar2
	);
  --功能：获取编辑器字体参数
  Procedure p_GetFontParam(
    Val Out t_Refcur,
    user_In nvarchar2
	);
  --功能：插入编辑器窗体参数
  Procedure p_SetFormParam(
    form_In nvarchar2, 
	user_In nvarchar2
	);
  --功能：获取编辑器字体参数
  Procedure p_GetFormParam(
    Val Out t_Refcur,
	user_In nvarchar2
	);
End b_PACS_RptPluginOriginal;
/

--影像报告范文管理(---实现部分---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptPluginOriginal Is

  --功    能：获取历史报告记录
  Procedure p_GetReportHistory(
    Val                   Out t_Refcur,
    医嘱id_In             In 病人医嘱记录.ID%Type,
    人员id_In             In 部门人员.人员id%Type,
    当前科室id_In         In 部门人员.部门ID%Type,
    查看其他科历史报告_In In number := 0
	) Is
    strSql     varchar2(4000);
    strSqlBack varchar2(4000);
    strFilter  varchar2(400);
  Begin
    If 查看其他科历史报告_In = 1 Then
      strFilter := ' ';
    Else
      strFilter := ' And c.执行科室id+0 in (select 部门id from 部门人员 where 人员id = '|| 人员id_In || 
                   ' union all select to_Number(' || 当前科室id_In || ') from dual) ';
    End If;

    strSql := 'Select 2 as 报告类型, f.编码'||'||''-''||'||'f.名称 As 科室名称, c.Id As 医嘱id, a.影像类别 as 类别,b.创建人 as 报告人,' ||
              'to_char(b.创建时间,''yyyy-mm-dd hh24:mi:ss'') as 创建时间,b.文档标题 报告名称, c.医嘱内容, TO_CHAR(RAWTOHEX(b.id)) 报告ID ' ||
              'From 影像检查记录 A, 影像报告记录 B, 病人医嘱记录 C, 影像检查记录 D, 病人医嘱记录 E, 部门表 F ' ||
              'Where a.医嘱id = b.医嘱id And d.医嘱id = e.Id And e.Id =' ||
              医嘱id_In || ' And e.执行科室ID = F.ID And b.医嘱id = c.Id And ' ||
              '(c.病人id = e.病人id Or a.关联id = d.关联id) And c.相关id Is Null ' || strFilter ||
              ' union all ' ||
              'Select 1 as 报告类型, g.编码'||'||''-''||'||'g.名称 As 科室名称, c.Id As 医嘱id, a.影像类别 as 类别, a.报告人, ' ||
              'to_char(f.创建时间,''yyyy-mm-dd hh24:mi:ss'') as 创建时间, a.影像类别||''报告'' 报告名称, c.医嘱内容,TO_CHAR( b.病历id) as 报告ID ' ||
              'From 影像检查记录 A, 病人医嘱报告 B, 病人医嘱记录 C, 影像检查记录 D, 病人医嘱记录 E, 电子病历记录 F, 部门表 G ' ||
              'Where a.医嘱id = b.医嘱id And d.医嘱id = e.Id And e.Id = ' ||
              医嘱id_In || ' And e.执行科室ID = g.ID And b.医嘱id = c.Id And b.病历ID Is Not Null And ' ||
              '(c.病人id = e.病人id Or a.关联id = d.关联id) And c.相关id Is Null And b.病历id = f.id ' || strFilter;
              
    strSqlBack := strSql;
    strSqlBack := replace(strSqlBack, '影像检查记录', 'H影像检查记录');
    strSqlBack := replace(strSqlBack, '病人医嘱报告', 'H病人医嘱报告');
    strSqlBack := replace(strSqlBack, '病人医嘱记录', 'H病人医嘱记录');

    strSql := strSql || ' UNION ALL ' || strSQLBack || ' Order By 创建时间 Asc';

    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetReportHistory;

  --功    能：获取对应报告内容
  Procedure p_GetReportContent(
    Val           Out t_Refcur,
    报告ID_In     varchar2,
    EditorType_In Number := 0 --0:电子病历编辑器，1--PACS报告编辑器，2--报告文档编辑器
    ) Is
    strSql varchar2(1000);
  Begin
    If EditorType_In = 1 Then
      strSql := 'Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文,b.开始版 as 版本 From 电子病历内容 a,电子病历内容 b ' ||
                'Where a.文件id = ' || 报告ID_In ||
                ' And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 and b.终止版=0 ';
    ElsIf EditorType_In = 0 Then
      strSql := 'select 内容 from 电子病历格式 where 文件ID=' || 报告ID_In;
    Else
      strSql := 'Select 报告内容 As 内容 From 影像报告记录 Where ID=HexToRaw(''' ||
                报告ID_In || ''')';
    End If;
  
    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetReportContent;

  --功    能：根据医嘱ID获取检查UID
  Procedure p_GetStudyUid(
    Val       Out t_Refcur,
    医嘱id_In In 影像检查记录.医嘱id%Type
	) Is
    strSql varchar2(100);
  Begin
    strSql := 'Select 检查UID from 影像检查记录 where 医嘱ID =' || 医嘱id_In;
    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyUid;

  --功    能：获取预览图像总数
  Procedure p_GetStudyImageCount(
    Val Out t_Refcur, 
    查询条件_In In varchar2,
    是否临时_In In number:=0 
	) Is
    strSql varchar2(2000);
  Begin
    if 是否临时_In = 0 then
      strSql := 'select T1.返回值+T2.返回值 as 返回值 from ' ||
              '(select count(1) as 返回值 from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c ' ||
              'where a.序列UID=b.序列UID and b.检查UID=c.检查UID and c.医嘱ID=''' ||
              查询条件_In || ''') T1,' ||
              '(select count(1) as 返回值 from H影像检查图象 a, H影像检查序列 b, 影像检查记录 c ' ||
              'where a.序列UID=b.序列UID and b.检查UID=c.检查UID and c.医嘱ID=''' ||
              查询条件_In || ''') T2';
    else
      strSql := 'select count(1)  as 返回值 from 影像临时图象  where  序列UID='''||查询条件_In || '''';
    end if;
  
    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyImageCount;

  --功    能：获取预览图像数据
  Procedure p_GetStudyImageData(
    Val         Out t_Refcur,
    查询方式_In In varchar2,
    查询条件_In In varchar2,
    开始位置_In In number,
    结束位置_In In number,
    是否临时_In In number
	) Is
    strSql    varchar2(2000);
    strFilter varchar2(100);
  Begin
    if 查询方式_In = 0 then
      strFilter := 'and c.医嘱ID=''' || 查询条件_In || '''';
    elsif 查询方式_In = 1 then
      strFilter := 'and B.序列UID=''' || 查询条件_In || '''';
    else
      strFilter := 'and A.图像UID=''' || 查询条件_In || '''';
    end if;
    
    strSql := 'Select * from (Select rownum as 顺序号, T.* from(' ||
              'Select A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1,D.IP地址 As Host1,''/''||D.Ftp目录||''/'' As Root1,' ||
              'Decode(C.接收日期,Null,'''',to_Char(C.接收日期,''YYYYMMDD'')||''/'')||C.检查UID||''/''||A.图像UID As URL,d.设备号 as 设备号1,' ||
              'E.FTP用户名 As User2,E.FTP密码 As Pwd2,E.IP地址 As Host2,''/''||E.Ftp目录||''/'' As Root2,' ||
              'e.设备号 as 设备号2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度 ' ||
              'From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E ' ||
              'Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) ' ||
              strFilter || ' '|| 'Order by 序列UID, 图像号) T ) ' ||
              'Where 顺序号>=' || 开始位置_In || ' and 顺序号<=' || 结束位置_In || '';
              
    if 是否临时_In = 1 then
      strSql:= replace(strSql,'影像检查','影像临时');
    end if;
    
    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyImageData;
  
  --功能：获取临时图像序列
  Procedure p_Get_TempImageSeries(
    Val Out t_Refcur,
    时间范围_In In Number,
    姓名_In In 影像临时记录.姓名%Type:=null
	) As
  Begin
    If 姓名_In Is Null Then
      Open Val For
        select B.序列UID,A.姓名,A.检查号 As 序号, A.接收日期 from 影像临时记录 A,影像临时序列 B
        where A.检查uid = B.检查uid And A.接收日期 Between Sysdate-时间范围_In And Sysdate
        order by 序号;
    Else
      Open Val For
        select B.序列UID,A.姓名,A.检查号 As 序号, A.接收日期 from 影像临时记录 A,影像临时序列 B
        where A.检查uid = B.检查uid And A.接收日期 Between Sysdate-时间范围_In And Sysdate and a.姓名 = 姓名_In
        order by 序号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --功能：获取图像备注
  Procedure p_Get_Normalnote(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select b.编号 As 编号, b.名称 As 名称
        From 影像字典清单 A, 影像字典内容 B
       Where a.Id = b.字典id
         And a.名称 = '影像图像备注';
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --功能：插入常用图像备注
  Procedure p_Insert_Normalnote(
    note_in In 影像字典内容.名称%Type,
	code_In 影像字典内容.简码%Type
	) As
    n_Num         Number;
    dictionary_id Varchar2(36);
  Begin
    Select id
      Into dictionary_id
      From 影像字典清单
     Where 说明 = '影像图像备注';
    Select Decode(Max(to_number(编号)), Null, 0, Max(to_number(编号)))
      Into n_Num
      From 影像字典内容
     Where 字典id = dictionary_id;
    n_Num := n_Num + 1;
    Insert Into 影像字典内容
      (字典id, 编号, 名称, 说明)
    Values
      (dictionary_id, to_char(n_Num), note_in, '影像图像备注');
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Insert_Normalnote;

  --功能：修改常用图像备注
  Procedure p_Edit_Normalnote(
    note_in In 影像字典内容.名称%Type,
	num_In  影像字典内容.编号%Type
	) As
    dictionary_id Varchar2(36);
  Begin
    Select id
      Into dictionary_id
      From 影像字典清单
     Where 说明 = '影像图像备注';
    Update 影像字典内容 t
       Set t.名称 = note_in
     Where t.字典id = dictionary_id
       And t.编号 = num_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Normalnote;

  --功能：删除常用图像备注
  Procedure p_Del_Normalnote(
    num_In 影像字典内容.编号%Type
	) As
    dictionary_id Varchar2(36);
  Begin
    Select id
      Into dictionary_id
      From 影像字典清单
     Where 说明 = '影像图像备注';
    Delete 影像字典内容 t
     Where t.字典id = dictionary_id
       And t.编号 = num_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Normalnote;

  --功能：获取备注的下一个编码
  Procedure p_Get_NormalNum(
    Val Out t_Refcur
	) As
    n_Num         Number;
    dictionary_id Varchar2(36);
  Begin
    Select id
      Into dictionary_id
      From 影像字典清单
     Where 说明 = '影像图像备注';
    Open Val For
      Select Decode(Max(to_number(编号)), Null, 1, Max(to_number(编号) + 1)) 编号
        From 影像字典内容 t
       Where t.字典id = dictionary_id;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_NormalNum;

  --功能：获取插件ID
  Procedure p_Get_PlugID(
    Val     Out t_Refcur,
    类名_In In 影像报告插件.类名%Type
	) Is
  Begin
    Open Val For
      Select RawToHex(ID) ID From 影像报告插件 Where 类名 = 类名_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_PlugID;

  --功能：插入编辑器字体参数
  Procedure p_SetFontParam(
    font_In nvarchar2, 
	user_In nvarchar2
	) As
    m_ID     nvarchar2(36);
    numcount int;
  Begin
    Select RawToHex(ID)
      Into m_ID
      From 影像参数说明
     Where 模块 = 'ImageEditor'
       And 参数名 = '字体设置';
    Select Count(*)
      Into numcount
      From 影像参数取值 t
     Where t.参数id = m_ID
       And t.参数标识 = user_In;
    If numcount > 0 then
      Update 影像参数取值 a
         Set a.参数值 = font_In
       Where a.参数标识 = user_In
         And a.参数id = m_ID;
    Else
      Insert Into 影像参数取值 a
        (ID, 参数ID, 参数标识, 参数值)
      Values
        (sys_Guid(), m_ID, user_In, font_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_SetFontParam;

  --功能：获取编辑器字体参数
  Procedure p_GetFontParam(
    Val Out t_Refcur,
    user_In nvarchar2	
	) As
    m_ID nvarchar2(36);
  Begin
    Select RawToHex(ID)
      Into m_ID
      From 影像参数说明
     Where 模块 = 'ImageEditor'
       And 参数名 = '字体设置';
    Open Val For
      Select a.参数值
        From 影像参数取值 a
       Where a.参数id = m_ID
         And a.参数标识 = user_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFontParam;

  --功能：插入编辑器窗体参数
  Procedure p_SetFormParam(
    form_In nvarchar2, 
	user_In nvarchar2
	) As
    m_ID     nvarchar2(36);
    numcount int;
  Begin
    Select RawToHex(ID)
      Into m_ID
      From 影像参数说明
     Where 模块 = 'ImageEditor'
       And 参数名 = '窗口设置';
    Select Count(*)
      Into numcount
      From 影像参数取值 t
     Where t.参数id = m_ID
       And t.参数标识 = user_In;
    If numcount > 0 then
      Update 影像参数取值 a
         Set a.参数值 = form_In
       Where a.参数标识 = user_In
         And a.参数id = m_ID;
    Else
      Insert Into 影像参数取值 a
        (ID, 参数ID, 参数标识, 参数值)
      Values
        (sys_Guid(), m_ID, user_In, form_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_SetFormParam;

  --功能：获取编辑器字体参数
  Procedure p_GetFormParam(
    Val Out t_Refcur,
	user_In nvarchar2
	) As
    m_ID nvarchar2(36);
  Begin
    Select RawToHex(ID)
      Into m_ID
      From 影像参数说明
     Where 模块 = 'ImageEditor'
       And 参数名 = '窗口设置';
    Open Val For
      Select a.参数值
        From 影像参数取值 a
       Where a.参数id = m_ID
         And a.参数标识 = user_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFormParam;
End b_PACS_RptPluginOriginal;
/

--89706:刘尔旋,2015-11-04,接口返回登记时间
Create Or Replace Procedure Zl_Third_Settlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --功能:三方接口结帐
  --入参:Xml_In:
  --<IN>
  --        <BRID>病人ID</BRID>         //病人ID
  --        <ZYID>主页ID</ZYID>         //主页ID
  --        <JSLX>2</JSLX>         //结算类型,1-门诊,2-住院.目前固定传2
  --        <JE></JE>         //本次结算总金额
  --       <JSLIST>
  --         <JS>
  --              <JSKLB>支付卡类别</JSKLB >
  --              <JSKH>支付卡号</ JSKH >
  --              <JSFS>支付方式</JSFS> //支付方式:现金;支票,如果是三方卡,可以传空
  --              <JSJE>结算金额</JSJE> //结算金额(正金额：个人补款，负金额：医院退款)<SFCYJ>为1时为冲预交金额
  --              <JYLSH>交易流水号</JYLSH>
  --              <ZY>摘要</ZY>
  --              <SFCYJ>是否冲预交</SFCYJ>  //是否冲预交，0-结算，1-冲预交.允冲预交时,只填JSJE节点
  --              <SFXFK>是否消费卡</SFXFK>  //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点
  --              <EXPENDLIST>  //扩展交易信息
  --                  <EXPEND>
  --                        <JYMC>交易名称</JYMC> //交易名称   退款时,传入冲预交的流水号
  --                        <JYLR>交易内容</JYLR> //交易内容   退款时,传入冲预交的金额
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --出参:Xml_Out
  --  <OUT>
  --       <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --    ――如无下列错误结点则说明正确执行
  --    <ERROR>
  --      <MSG>错误信息</MSG>
  --    </ERROR>
  --  </OUT>
  --------------------------------------------------------------------------------------------------
  n_主页id     病案主页.主页id%Type;
  n_病人id     病案主页.病人id%Type;
  n_结帐总额   病人预交记录.冲预交%Type;
  n_待结帐金额 病人预交记录.冲预交%Type;
  n_结算类型   Number(3);
  v_操作员编码 病人结帐记录.操作员编号%Type;
  v_操作员姓名 病人结帐记录.操作员姓名%Type;
  n_结帐id     病人结帐记录.Id%Type;
  n_冲预交金额 病人预交记录.冲预交%Type;
  d_结帐时间   Date;
  n_预交充值   病人预交记录.金额%Type;
  d_开始日期   Date;
  n_存在       Number(3);
  n_预交id     病人预交记录.Id%Type;
  n_科室id     病案主页.入院科室id%Type;
  d_结束日期   Date;
  n_结算卡序号 卡消费接口目录.编号%Type;
  n_时间类型   Number(3);
  v_Ids        Varchar2(20000);
  v_消费卡结算 Varchar2(5000);
  v_No         病人结帐记录.No%Type;
  v_预交no     病人预交记录.No%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_结算卡号   病人预交记录.卡号%Type;
  v_结算方式   病人预交记录.结算方式%Type;
  v_Temp       Varchar2(500);
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
  n_Count Number(18);

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/JSLX'))
  Into n_主页id, n_病人id, n_结帐总额, n_结算类型
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_结算类型 := Nvl(n_结算类型, 2);

  --0.相关检查
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份,不允许缴费!';
    Raise Err_Item;
  End If;

  --人员id,人员编号,人员姓名
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '系统不能认别有效的操作员,不允许结算!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员姓名 := v_Temp;
  v_Err_Msg    := Null;
  If n_结算类型 = 2 Then
    Begin
      Select Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0))
      Into n_待结帐金额
      From 住院费用记录
      Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1;
    Exception
      When Others Then
        n_待结帐金额 := 0;
    End;
  
    If n_待结帐金额 <> n_结帐总额 Then
      v_Err_Msg := '传入的结帐金额与实际结帐金额不符,不允许结算!';
      Raise Err_Item;
    End If;
    Begin
      Select 入院科室id Into n_科室id From 病案主页 Where 病人id = n_病人id And 主页id = n_主页id;
    Exception
      When Others Then
        n_科室id := Null;
    End;
  
    Select 病人结帐记录_Id.Nextval, Sysdate, Nextno(15) Into n_结帐id, d_结帐时间, v_No From Dual;
  
    n_时间类型 := zl_GetSysParameter('结帐费用时间', 1137);
    If n_时间类型 = 0 Then
      --按登记时间
      Select Trunc(Min(登记时间)), Trunc(Max(登记时间))
      Into d_开始日期, d_结束日期
      From (Select NO, 序号, 登记时间, 发生时间
             From 住院费用记录
             Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1
             Group By NO, 序号, 登记时间, 发生时间, Mod(记录性质, 10)
             Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0);
    Else
      --按发生时间  
      Select Trunc(Min(发生时间)), Trunc(Max(发生时间))
      Into d_开始日期, d_结束日期
      From (Select NO, 序号, 登记时间, 发生时间
             From 住院费用记录
             Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1
             Group By NO, 序号, 登记时间, 发生时间, Mod(记录性质, 10)
             Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0);
    End If;
  
    Zl_病人结帐记录_Insert(n_结帐id, v_No, n_病人id, d_结帐时间, d_开始日期, d_结束日期, 0, 0, n_主页id, Null, 2, Null, 2);
  
    For r_费用 In (Select Min(ID) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
                        Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
                 From 住院费用记录
                 Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1
                 Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
                 Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0
                 Order By NO, 序号) Loop
      If Nvl(r_费用.结帐金额, 0) = 0 Then
        Begin
          Select 1 Into n_存在 From 住院费用记录 Where ID = r_费用.Id And 结帐id Is Null;
        Exception
          When Others Then
            n_存在 := 0;
        End;
        If n_存在 = 1 Then
          v_Ids := v_Ids || ',' || r_费用.Id;
        Else
          Zl_结帐费用记录_Insert(0, r_费用.No, r_费用.记录性质, r_费用.记录状态, r_费用.执行状态, r_费用.序号, r_费用.金额, n_结帐id);
        End If;
      Else
        Zl_结帐费用记录_Insert(0, r_费用.No, r_费用.记录性质, r_费用.记录状态, r_费用.执行状态, r_费用.序号, r_费用.金额, n_结帐id);
      End If;
    End Loop;
  
    If v_Ids Is Not Null Then
      v_Ids := Substr(v_Ids, 2);
      Zl_结帐费用记录_Batch(v_Ids, n_病人id, n_结帐id);
    End If;
  
    n_Count := 0;
    For r_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_结算方式.是否冲预交, 0) = 0 Then
        --付款
        If n_Count = 1 Then
          v_Err_Msg := '结帐结算暂不支持多种结算方式!';
          Raise Err_Item;
        End If;
        If Nvl(r_结算方式.是否消费卡, 0) = 1 Then
          Begin
            n_结算卡序号 := To_Number(r_结算方式.结算卡类别);
          Exception
            When Others Then
              n_结算卡序号 := 0;
          End;
          If n_结算卡序号 = 0 Then
            Begin
              Select 编号
              Into n_结算卡序号
              From 卡消费接口目录
              Where 名称 = r_结算方式.结算卡类别 And Nvl(启用, 0) = 1;
            Exception
              When Others Then
                v_Err_Msg := '未找到对应的消费卡!';
                Raise Err_Item;
            End;
          End If;
          If v_结算方式 Is Null Then
            Select 结算方式 Into v_结算方式 From 卡消费接口目录 Where 编号 = n_结算卡序号;
          End If;
        Else
          Begin
            n_卡类别id := To_Number(r_结算方式.结算卡类别);
          Exception
            When Others Then
              n_卡类别id := 0;
          End;
          If n_卡类别id = 0 Then
            Begin
              Select ID Into n_卡类别id From 医疗卡类别 Where 名称 = r_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
            Exception
              When Others Then
                v_Err_Msg := '未找到对应的医疗卡!';
                Raise Err_Item;
            End;
          End If;
          If v_结算方式 Is Null Then
            Select 结算方式 Into v_结算方式 From 医疗卡类别 Where ID = n_卡类别id;
          End If;
        End If;
      
        If n_卡类别id Is Not Null Then
          --三方卡,生成住院预交款
          v_结算卡号 := r_结算方式.结算卡号;
          If r_结算方式.结算金额 > 0 Then
            Select 病人预交记录_Id.Nextval, Nextno(11) Into n_预交id, v_预交no From Dual;
            Zl_病人预交记录_Insert(n_预交id, v_预交no, Null, n_病人id, n_主页id, n_科室id, r_结算方式.结算金额, v_结算方式, '', '', '', '', '',
                             v_操作员编码, v_操作员姓名, Null, 2, n_卡类别id, Null, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明, Null,
                             d_结帐时间, 0);
            n_预交充值 := Nvl(n_预交充值, 0) + r_结算方式.结算金额;
          Else
          
            Zl_结帐缴款记录_Insert(v_No, n_病人id, n_主页id, n_科室id, v_结算方式, Null, r_结算方式.结算金额, n_结帐id, v_操作员编码, v_操作员姓名, d_结帐时间,
                             Null, Null, Null, Null, Null, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明);
          
          End If;
        Else
          If n_结算卡序号 Is Not Null Then
            --消费卡
            v_结算卡号   := r_结算方式.结算卡号;
            v_消费卡结算 := n_结算卡序号 || '|' || r_结算方式.结算卡号 || '|0|' || r_结算方式.结算金额 || '||';
            Zl_结帐缴款记录_Insert(v_No, n_病人id, n_主页id, n_科室id, v_结算方式, Null, r_结算方式.结算金额, n_结帐id, v_操作员编码, v_操作员姓名, d_结帐时间,
                             Null, Null, Null, Null, Null, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明, v_消费卡结算);
          Else
            --其他结算
            Zl_结帐缴款记录_Insert(v_No, n_病人id, n_主页id, n_科室id, v_结算方式, Null, r_结算方式.结算金额, n_结帐id, v_操作员编码, v_操作员姓名, d_结帐时间,
                             Null, Null, Null, Null, Null, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明);
          End If;
        End If;
      
        n_Count := 1;
      End If;
    End Loop;
  
    For r_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_结算方式.是否冲预交, 0) = 1 Then
        --冲预交,目前默认全冲
        n_冲预交金额 := r_结算方式.结算金额 + Nvl(n_预交充值, 0);
        For r_预交 In (Select Min(ID) As ID, NO, 结算方式, Sum(Nvl(金额, 0)) - Sum(Nvl(冲预交, 0)) As 金额, 交易流水号
                     From 病人预交记录
                     Where 病人id = n_病人id And Mod(记录性质, 10) = 1 And Nvl(预交类别, 2) = 2 And (主页id = n_主页id Or 主页id Is Null)
                     Group By NO, 结算方式, 交易流水号
                     Having Sum(Nvl(金额, 0)) - Sum(Nvl(冲预交, 0)) <> 0) Loop
          Zl_结帐预交记录_Insert(r_预交.Id, r_预交.No, 1, r_预交.金额, n_结帐id, n_病人id);
          n_冲预交金额 := n_冲预交金额 - Nvl(r_预交.金额, 0);
        End Loop;
        If n_冲预交金额 <> 0 Then
          v_Err_Msg := '传入的预交冲销金额与实际不符,请检查!';
          Raise Err_Item;
        End If;
      End If;
    End Loop;
  
    --处理扩展信息
    If Nvl(n_卡类别id, 0) <> 0 Then
      For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                            Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
        Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_预交id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 1);
      End Loop;
    End If;
    If Nvl(n_结算卡序号, 0) <> 0 Then
      For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                            Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
        Zl_三方结算交易_Insert(n_结算卡序号, 1, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
      End Loop;
    End If;
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_结帐时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Settlement;
/

--89706:刘尔旋,2015-11-04,接口返回登记时间
Create Or Replace Procedure Zl_Third_Deposit_Recharge
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --功能:预交款充值 
  --入参:Xml_In: 
  --    <IN>
  --        <BRID>病人ID</BRID>
  --        <ZYID>主页ID</ZYID>
  --        <SFMZ>是否门诊</SFMZ> //1-是门诊,0-住院
  --        <JSLIST>
  --            <JS>
  --              <JSKLB>支付卡类别</JSKLB >
  --              <JSKH>支付卡号</JSKH>
  --              <JYLSH>交易流水号</JYLSH>
  --              <JYSM>交易说明</JYSM>
  --              <JSFS>支付方式</JSFS> //充值方式:现金;支票,如果是三方卡,可以传空
  --              <JSJE>交易金额</JSJE> //充值金额
  --              <ZY>摘要</ZY> 
  --              <SFXFK>是否消费卡</SFXFK> 
  --              <JSHM>结算号码(可以不传)</JSHM> 
  --              <JKDW>缴款单位(可以不传)</JKDW> 
  --              <DWKFH>单位开户行(可以不传)</DWKFH> 
  --              <DWZH>单位帐号(可以不传)</DWZH> 
  --              <HZDW>合作单位(可以不传)</HZDW> 
  --              <EXPENDLIST>  //扩展交易信息
  --                   <EXPEND>
  --                        <JYMC>交易名称</JYMC>
  --                        <JYLR>交易内容</JYLR>
  --                   </EXPEND>
  --              </EXPENDLIST >
  --            </JS>
  --         </JSLIST>
  --    </IN>
  --出参:Xml_Out 
  --  <OUTPUT> 
  --     <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --     <YJDH>预交单号(多个逗号分隔)</YJDH>
  --    ――如无下列错误结点则说明正确执行 
  --    <ERROR> 
  --      <MSG>错误信息</MSG> 
  --    </ERROR> 
  --  </OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  v_结算方式   Varchar2(2000);
  v_Nos        Varchar2(4000);
  v_No         病人预交记录.No%Type;
  v_操作员编码 病人预交记录.操作员编号%Type;
  v_操作员姓名 病人预交记录.操作员姓名%Type;

  n_卡类别id   医疗卡类别.Id%Type;
  n_病人id     门诊费用记录.病人id%Type;
  n_主页id     病人预交记录.主页id%Type;
  n_科室id     病人预交记录.科室id%Type;
  n_预交id     病人预交记录.Id%Type;
  n_结算卡序号 病人预交记录.结算卡序号%Type;
  n_预交类别   病人预交记录.预交类别%Type;
  n_消费卡     Number(2);
  n_门诊预存   Number(2);

  d_登记时间 Date;

  v_Temp    Varchar2(32767); --临时XML 
  x_Templet Xmltype; --模板XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_Count Number(18);

  Function Zl_结算方式_Get
  (
    卡类别_In   Varchar2,
    消费卡_In   Number,
    卡类别id_In Out 病人预交记录.卡类别id%Type
  ) Return Varchar2 As
    v_名称 Varchar2(200);
  
  Begin
  
    v_结算方式 := Null;
    v_Err_Msg  := Null;
  
    Begin
      n_卡类别id := To_Number(卡类别_In);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
  
    If Nvl(消费卡_In, 0) = 1 Then
    
      If n_卡类别id = 0 Then
        Begin
          Select 编号, 结算方式, Decode(Nvl(启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!'), 名称
          Into n_卡类别id, v_结算方式, v_Err_Msg, v_名称
          From 卡消费接口目录
          Where 名称 = 卡类别_In;
        Exception
          When Others Then
            v_Err_Msg := '消费:' || 卡类别_In || '不存在!';
        End;
      
      Else
      
        Begin
          Select 编号, 结算方式, Decode(Nvl(启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!'), 名称
          Into n_卡类别id, v_结算方式, v_Err_Msg, v_名称
          From 卡消费接口目录
          Where 编号 = 卡类别_In;
        Exception
          When Others Then
            v_Err_Msg := '未找到指定的结算支付信息!';
        End;
      
      End If;
    
      If Not v_Err_Msg Is Null Then
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := Nvl(v_名称, '') || '未设置结算方式,请在消费卡管理中设置结算方式';
        Raise Err_Item;
      End If;
      卡类别id_In := n_卡类别id;
    
      Return v_结算方式;
    
    End If;
  
    If n_卡类别id = 0 Then
      Begin
        Select ID, 结算方式, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!'), 名称
        Into n_卡类别id, v_结算方式, v_Err_Msg, v_名称
        From 医疗卡类别
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          n_卡类别id := -1;
          v_Err_Msg  := 卡类别_In || '不存在!';
      End;
    Else
      Begin
        Select ID, 结算方式, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!'), 名称
        Into n_卡类别id, v_结算方式, v_Err_Msg, v_名称
        From 医疗卡类别
        Where ID = n_卡类别id;
      Exception
        When Others Then
          n_卡类别id := -1;
          v_Err_Msg  := '未找到指定的结算支付信息!';
      End;
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
    If v_结算方式 Is Null Then
      v_Err_Msg := Nvl(v_名称, '') || '未设置结算方式,请在医疗卡类别中设置结算方式';
      Raise Err_Item;
    End If;
  
    卡类别id_In := n_卡类别id;
    Return v_结算方式;
  End Zl_结算方式_Get;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), To_Number(Extractvalue(Value(A), 'IN/ZYID')),
         To_Number(Extractvalue(Value(A), 'IN/SFMZ'))
  Into n_病人id, n_主页id, n_门诊预存
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  --0.相关检查
  If Nvl(n_门诊预存, 0) = 0 Then
    n_预交类别 := 2;
  Else
    n_预交类别 := 1;
  End If;
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份,不允许充值!';
    Raise Err_Item;
  
  End If;

  --操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
  v_Temp := Zl_Identity;
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '系统不能认别有效的操作员,不允许缴费!';
    Raise Err_Item;
  End If;

  Select Nvl(a.当前科室id, b.出院科室id)
  Into n_科室id
  From 病人信息 A, 病案主页 B
  Where b.病人id(+) = a.病人id And a.主页id = b.主页id(+) And a.病人id = n_病人id;

  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_操作员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  v_Err_Msg := Null;

  If Not v_Err_Msg Is Null Then
    Raise Err_Item;
  End If;

  If Nvl(n_主页id, 0) = 0 Then
    n_主页id := Null;
  End If;

  If Nvl(n_科室id, 0) = 0 Then
    n_科室id := Null;
  End If;

  d_登记时间 := Sysdate;

  --2.确定支付方式
  n_Count := 0;
  v_Nos   := Null;
  For c_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                        Extractvalue(b.Column_Value, '/JS/JSHM') As 结算号码,
                        Extractvalue(b.Column_Value, '/JS/JKDW') As 缴款单位,
                        Extractvalue(b.Column_Value, '/JS/DWKFH') As 单位开户行,
                        Extractvalue(b.Column_Value, '/JS/DWZH') As 单位帐号,
                        Extractvalue(b.Column_Value, '/JS/HZDW') As 合作单位,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    v_No := Nextno(11);
    If v_Nos Is Null Then
      v_Nos := v_No;
    Else
      v_Nos := v_Nos || ',' || v_No;
    End If;
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
  
    If Nvl(c_结算方式.结算金额, 0) = 0 Then
      v_Err_Msg := '传入的充值金额为零,没必要进行充值处理,请检查充值金额是否传入错误!';
      Raise Err_Item;
    End If;
  
    n_结算卡序号 := Null;
    n_卡类别id   := Null;
    n_消费卡     := Nvl(c_结算方式.是否消费卡, 0);
  
    If c_结算方式.结算卡类别 Is Not Null Then
      --三方卡结算
      v_结算方式 := Zl_结算方式_Get(c_结算方式.结算卡类别, n_消费卡, n_卡类别id);
      If Nvl(n_消费卡, 0) = 1 Then
        n_结算卡序号 := n_卡类别id;
        n_卡类别id   := Null;
      End If;
    
    Else
      v_结算方式 := c_结算方式.结算方式;
      If v_结算方式 Is Null Then
        v_Err_Msg := '未确定本次充值的支付方式,请检查支付方式是否传入错误!';
        Raise Err_Item;
      End If;
    End If;
    Zl_病人预交记录_Insert(n_预交id, v_No, Null, n_病人id, n_主页id, n_科室id, c_结算方式.结算金额, v_结算方式, c_结算方式.结算号码, c_结算方式.缴款单位,
                     c_结算方式.单位开户行, c_结算方式.单位帐号, c_结算方式.摘要, v_操作员编码, v_操作员姓名, Null, n_预交类别, n_卡类别id, n_结算卡序号, c_结算方式.结算卡号,
                     c_结算方式.交易流水号, c_结算方式.交易说明, c_结算方式.合作单位, d_登记时间, 0);
  
    --保存扩展结算信息
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(c_结算方式.Expend, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(n_卡类别id, n_消费卡, c_结算方式.结算卡号, n_预交id, c_扩展.Jymc || '|' || c_扩展.Jylr, 1);
    End Loop;
    n_Count := n_Count + 1;
  End Loop;

  If n_Count = 0 Then
    v_Err_Msg := '不能有效确认当前充值的支付方式!';
    Raise Err_Item;
  End If;

  v_Temp := '<YJDH>' || v_Nos || '</YJDH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Deposit_Recharge;
/

--89706:刘尔旋,2015-11-04,接口返回登记时间
Create Or Replace Procedure Zl_Third_Charge_Del
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --功能:三方退费交易 
  --入参:Xml_In: 
  --<IN>
  --    <BRID>病人ID</BRID>
  --    <JE></JE> //退款总金额
  --    <JSKLB></JSKLB>     //结算卡类别
  --    <TFZY>退费摘要</TFZY>
  --    <JCFP>1</JCFP>      //检查发票
  --    <FYLIST>
  --        <FY>
  --           <DJH>退款单据号</DJH>
  --           <XH>退款序号(格式:1,2,3..为空代表退剩余数量)</DJH>
  --        <FY>
  --    </FYLIST>
  --    <TKLIST>
  --        <TK>
  --            <TKKLB>退款卡类别</TKKLB>
  --            <TKKH>退款卡号</TKKH>
  --            <TKFS>退款方式</TKFS> //退款方式:现金;支票,如果是三方卡,可以传空
  --            <TKJE>支付金额</TKJE>
  --            <JYLSH>交易流水号</JYLSH>
  --            <TKZY>摘要</TKZY>
  --            <TYJK>退回预交款</TYJK> //允冲预交时,只填JSJE节点:1-冲预交
  --            <SFXFK>是否消费卡</SFXFK>   //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点
  --            <EXPENDLIST>  //扩展交易信息
  --                <EXPEND>
  --                    <JYMC>交易名称</JYMC>
  --                    <JYLR>交易内容</JYLR>
  --                </EXPEND>
  --            </EXPENDLIST>
  --        </TK>
  --    </TKLIST>
  --</IN>

  --出参:Xml_Out 
  --  <OUT> 
  --    <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --    ――如无下列错误结点则说明正确执行 
  --    <ERROR> 
  --      <MSG>错误信息</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  n_退款总额 门诊费用记录.实收金额%Type;
  n_卡类别id 医疗卡类别.Id%Type;
  v_结算方式 Varchar2(2000);

  n_病人id     门诊费用记录.病人id%Type;
  n_单据病人id 门诊费用记录.病人id%Type;
  v_操作员编码 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  n_冲销id     门诊费用记录.结帐id%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_结帐金额   门诊费用记录.结帐金额%Type;
  n_误差额     病人预交记录.冲预交%Type;
  n_原结算序号 病人预交记录.结算序号%Type;
  l_挂号单     t_Strlist := t_Strlist();
  v_挂号单     门诊费用记录.No%Type;
  n_检查发票   Number(3);
  n_是否打印   Number(3);
  v_结算卡类别 Varchar2(100);

  n_消费卡id 消费卡目录.Id%Type;
  v_摘要     门诊费用记录.摘要%Type;
  n_Count    Number(18);

  d_退费时间 病人预交记录.收款时间%Type;

  v_退费结算 Varchar2(2000);
  v_普通结算 Varchar2(4000);
  n_Temp     Number(18);

  v_Temp    Varchar2(32767); --临时XML 
  x_Templet Xmltype; --模板XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  Procedure Third_Cardbalance_Modfiy
  (
    冲销id_In     病人预交记录.结帐id%Type,
    卡类别_In     Varchar2,
    卡号_In       病人预交记录.卡号%Type,
    退款金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    摘要_In       病人预交记录.摘要%Type,
    Xmlexpned_In  Xmltype
  ) Is
    n_卡类别id 医疗卡类别.Id%Type;
    v_结算方式 病人预交记录.结算方式%Type;
  
  Begin
    v_Err_Msg := Null;
    Begin
      n_卡类别id := To_Number(卡类别_In);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
    If n_卡类别id = 0 Then
      Begin
        Select ID, 结算方式, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_结算方式, v_Err_Msg
        From 医疗卡类别
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          n_卡类别id := -1;
          v_Err_Msg  := 卡类别_In || '不存在!';
      End;
    Else
      Begin
        Select ID, 结算方式, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_结算方式, v_Err_Msg
        From 医疗卡类别
        Where ID = n_卡类别id;
      Exception
        When Others Then
          n_卡类别id := -1;
          v_Err_Msg  := '未找到指定的结算支付信息!';
      End;
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  
    v_退费结算 := v_结算方式 || '|' || 退款金额_In || '|' || ' |' || Nvl(摘要_In, ' ');
    --   2.三方卡退费结算:
    --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:结算方式|结算金额|结算号码|结算摘要
    --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    --结算方式|结算金额|结算号码|结算摘要 
    Zl_门诊退费结算_Modify(2, n_病人id, 冲销id_In, v_退费结算, 0, n_卡类别id, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0);
  
    --保存扩展结算信息
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 0, 卡号_In, 冲销id_In, c_扩展.Jymc || '|' || c_扩展.Jylr, 0);
    End Loop;
  End Third_Cardbalance_Modfiy;

  Procedure Square_Cardbalance_Modfiy
  (
    冲销id_In     病人预交记录.结帐id%Type,
    卡类别_In     Varchar2,
    卡号_In       病人预交记录.卡号%Type,
    退款金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    摘要_In       病人预交记录.摘要%Type,
    
    Xmlexpned_In Xmltype
  ) Is
    n_卡类别id 医疗卡类别.Id%Type;
    v_结算方式 病人预交记录.结算方式%Type;
  
  Begin
    v_Err_Msg := Null;
    Begin
      n_卡类别id := To_Number(卡类别_In);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
  
    If n_卡类别id = 0 Then
      Begin
        Select 编号, 结算方式, Decode(Nvl(启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_结算方式, v_Err_Msg
        From 卡消费接口目录
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := '消费:' || 卡类别_In || '不存在!';
      End;
    
    Else
    
      Begin
        Select 编号, 结算方式, Decode(Nvl(启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_结算方式, v_Err_Msg
        From 卡消费接口目录
        Where 编号 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  
    v_退费结算 := v_结算方式 || '|' || 退款金额_In || '|' || ' |' || Nvl(摘要_In, ' ');
    --   4-消费卡结算:
    --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||
    --     ②退支票额_In:传入零
    Select ID
    Into n_消费卡id
    From 消费卡目录
    Where 接口编号 = n_卡类别id And 卡号 = 卡号_In And
          序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = n_卡类别id And 卡号 = 卡号_In);
  
    --卡类别ID|卡号|消费卡ID|消费金额||.
    v_退费结算 := n_卡类别id || '|' || 卡号_In || '|' || n_消费卡id || '|' || 退款金额_In;
    Zl_门诊退费结算_Modify(4, n_病人id, 冲销id_In, v_退费结算, 0, Null, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0);
  
    --保存扩展结算信息
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 1, 卡号_In, 冲销id_In, c_扩展.Jymc || '|' || c_扩展.Jylr, 0);
    End Loop;
  End Square_Cardbalance_Modfiy;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  --0.获取入参中的病人ID等信息
  Select To_Number(Extractvalue(Value(A), 'IN/BRID')), To_Number(Extractvalue(Value(A), 'IN/JE')),
         Extractvalue(Value(A), 'IN/TFZY'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Extractvalue(Value(A), 'IN/JSKLB')
  Into n_病人id, n_退款总额, v_摘要, n_检查发票, v_结算卡类别
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  --人员id,人员编号,人员姓名 
  v_Temp       := Zl_Identity(1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员姓名 := v_Temp;
  v_Err_Msg    := Null;

  If v_结算卡类别 Is Not Null Then
    Begin
      n_卡类别id := To_Number(v_结算卡类别);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
    If n_卡类别id = 0 Then
      Begin
        Select ID Into n_卡类别id From 医疗卡类别 Where 名称 = v_结算卡类别;
      Exception
        When Others Then
          v_Err_Msg := '无法确认传入的结算卡！';
          Raise Err_Item;
      End;
    End If;
  Else
    n_卡类别id := 0;
  End If;

  If Nvl(n_卡类别id, 0) <> 0 Then
    Select 结算方式 Into v_结算方式 From 医疗卡类别 Where ID = n_卡类别id;
  End If;

  --1.先进行退费

  Select 病人结帐记录_Id.Nextval, Sysdate Into n_冲销id, d_退费时间 From Dual;

  n_Count      := 0;
  n_原结算序号 := 0;
  For c_费用 In (Select Extractvalue(b.Column_Value, '/FY/DJH') As 单据号, Extractvalue(b.Column_Value, '/FY/XH') As 退款序号
               From Table(Xmlsequence(Extract(Xml_In, '/IN/FYLIST/FY'))) B) Loop
    Begin
      Select 结算序号, 结帐id, 病人id
      Into n_Temp, n_结帐id, n_单据病人id
      From 病人预交记录
      Where 结帐id In (Select 结帐id
                     From 门诊费用记录
                     Where NO = c_费用.单据号 And 记录性质 = 1 And Nvl(费用状态, 0) = 0 And 记录状态 In (1, 3) And Rownum < 2) And
            Rownum < 2;
    Exception
      When Others Then
        n_Temp := Null;
    End;
  
    If n_Temp Is Null Then
      v_Err_Msg := '指定的单据号:' || c_费用.单据号 || '未找到,不能退费!';
      Raise Err_Item;
    End If;
    Begin
      Select Max(Decode(Instr(摘要, '挂号:'), 0, '', Replace(摘要, '挂号:', '')))
      Into v_挂号单
      From 门诊费用记录
      Where NO = c_费用.单据号 And 记录性质 = 1 And Nvl(费用状态, 0) = 0 And 记录状态 In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        v_挂号单 := Null;
    End;
    If Not v_挂号单 Is Null Then
      l_挂号单.Extend;
      l_挂号单(l_挂号单.Count) := v_挂号单;
    End If;
  
    If Nvl(n_单据病人id, 0) = 0 Then
      Begin
        Select 病人id
        Into n_单据病人id
        From 门诊费用记录
        Where NO = c_费用.单据号 And 记录性质 = 1 And Nvl(费用状态, 0) = 0 And 记录状态 In (1, 3) And Rownum < 2;
      Exception
        When Others Then
          n_单据病人id := 0;
      End;
    End If;
  
    If Nvl(n_病人id, 0) <> Nvl(n_单据病人id, 0) Then
      v_Err_Msg := '本次退费的收费单:' || c_费用.单据号 || '不是当前病人的收费单,不能退费!';
      Raise Err_Item;
    End If;
  
    If n_原结算序号 <> 0 And n_原结算序号 <> n_Temp Then
      v_Err_Msg := '本次退费的单据号不是一次收费结算,不能退费!';
      Raise Err_Item;
    End If;
    n_原结算序号 := n_Temp;
  
    Select Count(*) Into n_Temp From 费用补充记录 Where 收费结帐id = n_结帐id;
    If Nvl(n_Temp, 0) <> 0 Then
      v_Err_Msg := '本次退费的单据号已经进行了保险补充结算,不能退费!';
      Raise Err_Item;
    End If;
  
    If v_结算卡类别 Is Not Null Then
      Select Count(*) Into n_Temp From 病人预交记录 Where 结帐id = n_结帐id And 结算方式 = v_结算方式;
      If Nvl(n_Temp, 0) = 0 Then
        v_Err_Msg := '本次退费的单据不是' || v_结算方式 || '结算的,不能退费!';
        Raise Err_Item;
      End If;
    End If;
  
    If Nvl(n_检查发票, 0) = 1 Then
      Select Max(Decode(a.实际票号, Null, 0, 1))
      Into n_是否打印
      From 门诊费用记录 A
      Where NO = c_费用.单据号 And 记录性质 = 1;
      If Nvl(n_是否打印, 0) = 1 Then
        v_Err_Msg := '本次退费的单据号已开发票,不能退费!';
        Raise Err_Item;
      End If;
    End If;
  
    Zl_门诊收费记录_销帐(c_费用.单据号, v_操作员编码, v_操作员姓名, c_费用.退款序号, d_退费时间, v_摘要, n_冲销id);
    n_Count := n_Count + 1;
  End Loop;
  If n_Count = 0 Then
    v_Err_Msg := '未确定本次需要退费的单据,不能退费!';
    Raise Err_Item;
  End If;

  --2.处理退费的结算信息

  n_结帐金额 := 0;

  --检查总金额是否正确 
  Select Sum(结帐金额) Into n_结帐金额 From 门诊费用记录 Where 结帐id = n_冲销id;

  n_误差额 := -1 * Nvl(n_结帐金额, 0) - Nvl(n_退款总额, 0);
  If Abs(n_误差额) > 1.00 Then
    v_Err_Msg := '单据缴款金额与实际结算的误差太大!';
    Raise Err_Item;
  End If;

  --2.确定支付方式
  n_Count := 0;
  For c_结算方式 In (Select Extractvalue(b.Column_Value, '/TK/TKKLB') As 卡类别, Extractvalue(b.Column_Value, '/TK/TKKH') As 卡号,
                        Extractvalue(b.Column_Value, '/TK/TKFS') As 结算方式,
                        -1 * To_Number(Extractvalue(b.Column_Value, '/TK/TKJE')) As 退款金额,
                        Extractvalue(b.Column_Value, '/TK/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/TK/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/TK/TKZY') As 摘要,
                        Extractvalue(b.Column_Value, '/TK/TYJK') As 是否退预交,
                        Extractvalue(b.Column_Value, '/TK/SFXFK') As 是否消费卡,
                        Extract(b.Column_Value, '/TK/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/TKLIST/TK'))) B) Loop
  
    --1.退回三方卡
    If c_结算方式.卡类别 Is Not Null And Nvl(c_结算方式.是否消费卡, 0) = 0 Then
      --1.三方卡结算
      Third_Cardbalance_Modfiy(n_冲销id, c_结算方式.卡类别, c_结算方式.卡号, c_结算方式.退款金额, c_结算方式.交易流水号, c_结算方式.交易说明, c_结算方式.摘要,
                               c_结算方式.Expend);
    Elsif c_结算方式.卡类别 Is Not Null And Nvl(c_结算方式.是否消费卡, 0) = 1 Then
      --2.消费卡结算
      Square_Cardbalance_Modfiy(n_冲销id, c_结算方式.卡类别, c_结算方式.卡号, c_结算方式.退款金额, c_结算方式.交易流水号, c_结算方式.交易说明, c_结算方式.摘要,
                                c_结算方式.Expend);
    Elsif Nvl(c_结算方式.是否退预交, 0) = 1 Then
      --3.退预交款
      Zl_门诊退费结算_Modify(4, n_病人id, n_冲销id, Null, c_结算方式.退款金额, Null, Null, Null, Null, 0, 0, 0, 0);
    Else
      --4.普通结算
      If c_结算方式.结算方式 Is Null Then
        v_Err_Msg := '未指定指付方式，不允缴款!';
        Raise Err_Item;
      End If;
      --结算方式|结算金额|结算号码|结算摘要||..
      v_退费结算 := c_结算方式.结算方式 || '|' || c_结算方式.退款金额 || '| |' || Nvl(c_结算方式.摘要, '  ');
      v_普通结算 := Nvl(v_普通结算, '') || '||' || v_退费结算;
    End If;
    n_Count := n_Count + 1;
  End Loop;

  --   0-原样退
  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
  --   1-普通退费方式:
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||..
  --     ②退支票额_In:传入零
  If n_Count = 0 Then
    v_Err_Msg := '不能有效确认当前的支付方式!';
    Raise Err_Item;
  End If;

  --5.普通结算及完成结
  If v_普通结算 Is Not Null Then
    v_普通结算 := Substr(v_普通结算, 3);
  End If;
  Zl_门诊退费结算_Modify(1, n_病人id, n_冲销id, v_普通结算, 0, Null, Null, Null, Null, 0, 0, n_误差额, 2);

  If l_挂号单.Count <> 0 Then
  
    For I In 0 .. l_挂号单.Count Loop
      x_Templet := Xmltype('<IN></IN>');
      v_Temp    := '<GHDH>' || l_挂号单(I) || '</GHDH>';
      v_Temp    := v_Temp || '<JSKLB>' || 4 || '</JSKLB>';
      v_Temp    := v_Temp || '<GHJE>' || 0 || '</GHJE>';
    
      Select Appendchildxml(x_Templet, '/IN', Xmltype(v_Temp)) Into x_Templet From Dual;
      Zl_Third_Registdel(x_Templet, Xml_Out);
    End Loop;
  Else
    v_Temp := '<CZSJ>' || To_Char(d_退费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
  End If;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Charge_Del;
/

--89706:刘尔旋,2015-11-04,接口返回登记时间
Create Or Replace Procedure Zl_Third_Payment
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --功能:三方接口支付 
  --入参:Xml_In: 
  --<IN>
  --        <NO></NO>                       //收费单据号串,逗号分隔多个单据号
  --        <JE></JE>                       //总金额
  --        <BRID>病人ID</BRID>
  --        <SFGH></SFGH>                   //是否挂号单
  --        <WCJE>误差额</WCJE>             //误差项不传时,以总金额-本次结算费用总额为准
  --       <JSLIST>
  --         <JS>
  --              <JSKLB>支付卡类别</JSKLB >
  --              <JSKH>支付卡号</ JSKH >
  --              <JSFS>支付方式</JSFS> //支付方式:现金;支票,如果是三方卡,可以传空
  --              <JSJE>支付金额</JSJE>
  --              <JYLSH>交易流水号</JYLSH>
  --              <ZY>摘要</ZY>
  --              <SFCYJ>是否冲预交</SFCYJ>  //允冲预交时,只填JSJE节点:1-冲预交
  --              <SFXFK>是否消费卡</SFXFK>  //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点
  --              <EXPENDLIST>  //扩展交易信息
  --                  <EXPEND>
  --                        <JYMC >交易名称</交易名称>
  --                        <JYLR>交易内容</JYLR>
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --出参:Xml_Out 
  --  <OUT> 
  --    ――如无下列错误结点则说明正确执行 
  --    <ERROR> 
  --    <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --      <MSG>错误信息</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  v_Nos      Varchar2(4000);
  n_收费总额 门诊费用记录.实收金额%Type;

  n_卡类别id 医疗卡类别.Id%Type;
  v_结算方式 Varchar2(2000);
  n_病人id   门诊费用记录.病人id%Type;
  v_姓名     门诊费用记录.姓名%Type;
  v_性别     门诊费用记录.性别%Type;
  v_年龄     门诊费用记录.年龄%Type;

  v_医疗付款方式编码 医疗付款方式.编码%Type;
  v_操作员编码       门诊费用记录.操作员编号%Type;
  v_操作员姓名       门诊费用记录.操作员姓名%Type;
  n_结帐id           门诊费用记录.结帐id%Type;
  n_结帐金额         门诊费用记录.结帐金额%Type;
  d_收费时间         病人预交记录.收款时间%Type;
  n_消费卡id         消费卡目录.Id%Type;
  v_收费结算         Varchar2(2000);
  v_普通结算         Varchar2(4000);
  n_是否挂号         Number(3);
  n_预交支付         门诊费用记录.实收金额%Type;
  n_普通支付         门诊费用记录.实收金额%Type;
  v_结算卡号         病人预交记录.卡号%Type;
  n_结算卡序号       病人预交记录.结算卡序号%Type;
  v_交易流水号       病人预交记录.交易流水号%Type;
  v_交易说明         病人预交记录.交易说明%Type;
  v_摘要             病人预交记录.摘要%Type;
  n_科室id           挂号安排.科室id%Type;
  n_项目id           挂号安排.项目id%Type;
  n_医生id           挂号安排.医生id%Type;
  v_医生姓名         挂号安排.医生姓名%Type;
  v_号码             挂号安排.号码%Type;
  n_门诊号           病人信息.门诊号%Type;
  d_发生时间         病人挂号记录.发生时间%Type;
  v_费别             病人信息.费别%Type;
  n_号序             病人挂号记录.号序%Type;
  n_生成队列         Number(3);

  v_Temp    Varchar2(32767); --临时XML 
  x_Templet Xmltype; --模板XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_Count    Number(18);
  v_发药窗口 Varchar2(4000);
  n_误差额   病人预交记录.冲预交%Type;
  Function Zl_诊室(号码_In 挂号安排.号码%Type) Return Varchar2 As
    n_分诊方式 挂号安排.分诊方式%Type;
    n_安排id   挂号安排.Id%Type;
    v_诊室     病人挂号记录.诊室%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
  
    Begin
      Select ID, Nvl(分诊方式, 0) Into n_安排id, n_分诊方式 From 挂号安排 Where 号码 = 号码_In;
    Exception
      When Others Then
        n_安排id := -1;
    End;
  
    If n_安排id = -1 Then
      v_Err_Msg := '号码(' || 号码_In || ')未找到!';
      Raise Err_Item;
    End If;
    --0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
    v_诊室 := Null;
    If n_分诊方式 = 1 Then
      --1-指定诊室
      Begin
        Select 门诊诊室 Into v_诊室 From 挂号安排诊室 Where 号表id = n_安排id;
      Exception
        When Others Then
          v_诊室 := Null;
      End;
    End If;
    If n_分诊方式 = 2 Then
      --2-动态分诊:该个号别当天挂号未诊数最少的诊室
      For c_诊室 In (Select 门诊诊室, Sum(Num) As Num
                   From (Select 门诊诊室, 0 As Num
                          From 挂号安排诊室
                          Where 号表id = n_安排id
                          Union All
                          Select 诊室, Count(诊室) As Num
                          From 病人挂号记录
                          Where Nvl(执行状态, 0) = 0 And 发生时间 Between Trunc(Sysdate) And Sysdate And 号别 = 号码_In And
                                诊室 In (Select 门诊诊室 From 挂号安排诊室 Where 号表id = n_安排id)
                          Group By 诊室)
                   Group By 门诊诊室
                   Order By Num) Loop
        v_诊室 := c_诊室.门诊诊室;
        Exit;
      End Loop;
    End If;
    If n_分诊方式 = 3 Then
    
      --平均分诊：当前分配=1表示下次应取的当前诊室
      n_Next  := 0;
      n_First := 1;
      For c_诊室 In (Select Rowid As Rid, 号表id, 门诊诊室, 当前分配 From 挂号安排诊室 Where 号表id = n_安排id) Loop
        If n_First = 1 Then
          v_Rowid := c_诊室.Rid;
        End If;
        If n_Next = 1 Then
          v_诊室 := c_诊室.门诊诊室;
          Update 挂号安排诊室 Set 当前分配 = 1 Where Rowid = c_诊室.Rid;
          Exit;
        End If;
        If Nvl(c_诊室.当前分配, 0) = 1 Then
          Update 挂号安排诊室 Set 当前分配 = 0 Where Rowid = c_诊室.Rid;
          n_Next := 1;
        End If;
      End Loop;
      If v_诊室 Is Null Then
        Update 挂号安排诊室 Set 当前分配 = 1 Where Rowid = v_Rowid Returning 门诊诊室 Into v_诊室;
      End If;
    End If;
  
    Return v_诊室;
  End;
  Procedure Third_Cardbalance_Modfiy
  (
    结帐id_In     病人预交记录.结帐id%Type,
    卡类别_In     Varchar2,
    卡号_In       病人预交记录.卡号%Type,
    支付金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    Xmlexpned_In  Xmltype
  ) Is
    n_卡类别id 医疗卡类别.Id%Type;
    v_结算方式 病人预交记录.结算方式%Type;
    v_名称     医疗卡类别.名称%Type;
  Begin
    v_Err_Msg := Null;
    Begin
      n_卡类别id := To_Number(卡类别_In);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
    If n_卡类别id = 0 Then
      Begin
        Select ID, 结算方式, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!'), 名称
        Into n_卡类别id, v_结算方式, v_Err_Msg, v_名称
        From 医疗卡类别
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          n_卡类别id := -1;
          v_Err_Msg  := 卡类别_In || '不存在!';
      End;
    Else
      Begin
        Select ID, 结算方式, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!'), 名称
        Into n_卡类别id, v_结算方式, v_Err_Msg, v_名称
        From 医疗卡类别
        Where ID = n_卡类别id;
      Exception
        When Others Then
          n_卡类别id := -1;
          v_Err_Msg  := '未找到指定的结算支付信息!';
      End;
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
    If v_结算方式 Is Null Then
      v_Err_Msg := Nvl(v_名称, '') || '未设置结算方式,请在医疗卡类别中设置结算方式';
      Raise Err_Item;
    End If;
  
    v_收费结算 := v_结算方式 || '|' || 支付金额_In || '|' || ' |' || ' ';
    --结算方式|结算金额|结算号码|结算摘要 
    Zl_门诊收费结算_Modify(1, n_病人id, 结帐id_In, v_收费结算, 0, 0, n_卡类别id, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0);
  
    --保存扩展结算信息
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 0, 卡号_In, 结帐id_In, c_扩展.Jymc || '|' || c_扩展.Jylr, 0);
    End Loop;
  End Third_Cardbalance_Modfiy;

  Procedure Square_Cardbalance_Modfiy
  (
    结帐id_In     病人预交记录.结帐id%Type,
    卡类别_In     Varchar2,
    卡号_In       病人预交记录.卡号%Type,
    支付金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    Xmlexpned_In  Xmltype
  ) Is
    n_卡类别id 医疗卡类别.Id%Type;
    v_结算方式 病人预交记录.结算方式%Type;
    v_名称     卡消费接口目录.名称%Type;
  Begin
    v_Err_Msg := Null;
    Begin
      n_卡类别id := To_Number(卡类别_In);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
  
    If n_卡类别id = 0 Then
      Begin
        Select 编号, 结算方式, Decode(Nvl(启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!'), 名称
        Into n_卡类别id, v_结算方式, v_Err_Msg, v_名称
        From 卡消费接口目录
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := '消费:' || 卡类别_In || '不存在!';
      End;
    
    Else
    
      Begin
        Select 编号, 结算方式, Decode(Nvl(启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!'), 名称
        Into n_卡类别id, v_结算方式, v_Err_Msg, v_名称
        From 卡消费接口目录
        Where 编号 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
    If v_结算方式 Is Null Then
      v_Err_Msg := Nvl(v_名称, '') || '未设置结算方式,请在医疗卡类别中设置结算方式';
      Raise Err_Item;
    End If;
  
    Select ID
    Into n_消费卡id
    From 消费卡目录
    Where 接口编号 = n_卡类别id And 卡号 = 卡号_In And
          序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = n_卡类别id And 卡号 = 卡号_In);
  
    --结算方式_IN格式为:卡类别ID|卡号|消费卡ID|消费金额||.... 
    v_收费结算 := n_卡类别id || '|' || 卡号_In || '|' || n_消费卡id || '|' || 支付金额_In;
    Zl_门诊收费结算_Modify(3, n_病人id, 结帐id_In, v_收费结算, 0, 0, n_卡类别id, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0);
    --保存扩展结算信息
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 1, 卡号_In, 结帐id_In, c_扩展.Jymc || '|' || c_扩展.Jylr, 0);
    End Loop;
  End Square_Cardbalance_Modfiy;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/NO'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/WCJE')),
         To_Number(Extractvalue(Value(A), 'IN/SFGH'))
  Into v_Nos, n_病人id, n_收费总额, n_误差额, n_是否挂号
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  --0.相关检查

  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份,不允许缴费!';
    Raise Err_Item;
  
  End If;

  If v_Nos Is Null Then
    v_Err_Msg := '没有指定相关的收费单据,不允许缴费!';
    Raise Err_Item;
  
  End If;

  --人员id,人员编号,人员姓名 
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '系统不能认别有效的操作员,不允许缴费!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员姓名 := v_Temp;
  v_Err_Msg    := Null;
  Begin
    Select b.编码, a.姓名, a.性别, a.年龄
    Into v_医疗付款方式编码, v_姓名, v_性别, v_年龄
    From 病人信息 A, 医疗付款方式 B
    Where a.医疗付款方式 = b.名称(+) And a.病人id = n_病人id;
  Exception
    When Others Then
      v_Err_Msg := '指定的缴费单据中不能有效识别病人,不允许缴费!';
  End;
  If Not v_Err_Msg Is Null Then
    Raise Err_Item;
  End If;
  Select 病人结帐记录_Id.Nextval, Sysdate Into n_结帐id, d_收费时间 From Dual;

  If Nvl(n_是否挂号, 0) = 0 Then
    --费用单据
    v_发药窗口 := Zl_Getclinicchargepaywins(v_Nos);
  
    --1.进行费用收费处理
    --获取发药窗口
  
    n_结帐金额 := 0;
    For c_缴费单 In (Select /*+ rule */
                   a.No, Max(a.开单部门id) As 开单部门id, Max(a.病人科室id) As 病人科室id, Max(a.病人id) As 病人id, Sum(实收金额) As 实收金额,
                   Max(a.开单人) As 开单人
                  From 门诊费用记录 A, Table(f_Str2list(v_Nos)) J
                  Where a.记录性质 = 1 And a.No = j.Column_Value And a.记录状态 = 0
                  Group By a.No) Loop
      If Nvl(c_缴费单.病人id, 0) <> n_病人id Then
        v_Err_Msg := '缴费单据:' || c_缴费单.No || '与当前病人身份不符,不允许缴费!';
        Raise Err_Item;
      End If;
    
      n_结帐金额 := n_结帐金额 + Nvl(c_缴费单.实收金额, 0);
      Zl_病人划价收费_Insert(c_缴费单.No, n_病人id, 1, v_医疗付款方式编码, v_姓名, v_性别, v_年龄, c_缴费单.病人科室id, c_缴费单.开单部门id, c_缴费单.开单人, n_结帐id,
                       d_收费时间, v_操作员编码, v_操作员姓名, v_发药窗口, 0, d_收费时间);
    
    End Loop;
  
    --检查总金额是否正确 
    If Nvl(n_误差额, 0) = 0 Then
      n_误差额 := Nvl(n_收费总额, 0) - Nvl(n_结帐金额, 0);
      If Abs(n_误差额) > 1.00 Then
        v_Err_Msg := '单据缴款金额与实际结算的误差太大!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_结帐金额, 0) + Nvl(n_误差额, 0) <> Nvl(n_收费总额, 0) Then
      v_Err_Msg := '指定的缴费单据的总金额不对,请重新选择缴费单据!';
      Raise Err_Item;
    End If;
  
    --2.确定支付方式
    n_Count := 0;
    For c_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                          Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      --1.三方卡结算
    
      If c_结算方式.结算卡类别 Is Not Null And Nvl(c_结算方式.是否消费卡, 0) = 0 Then
        --1.三方卡结算
        Third_Cardbalance_Modfiy(n_结帐id, c_结算方式.结算卡类别, c_结算方式.结算卡号, c_结算方式.结算金额, c_结算方式.交易流水号, c_结算方式.交易说明,
                                 c_结算方式.Expend);
      Elsif c_结算方式.结算卡类别 Is Not Null And Nvl(c_结算方式.是否消费卡, 0) = 1 Then
        --2.消费卡结算
        Square_Cardbalance_Modfiy(n_结帐id, c_结算方式.结算卡类别, c_结算方式.结算卡号, c_结算方式.结算金额, c_结算方式.交易流水号, c_结算方式.交易说明,
                                  c_结算方式.Expend);
      Elsif Nvl(c_结算方式.是否冲预交, 0) = 1 Then
        --3.冲预交款
        Zl_门诊收费结算_Modify(0, n_病人id, n_结帐id, Null, c_结算方式.结算金额, 0, Null, Null, Null, Null, 0, 0, 0, 0);
      Else
        --4.普通结算
        If c_结算方式.结算方式 Is Null Then
          v_Err_Msg := '未指定指付方式，不允缴款!';
          Raise Err_Item;
        End If;
        --结算方式|结算金额|结算号码|结算摘要||..
        v_收费结算 := c_结算方式.结算方式 || '|' || c_结算方式.结算金额 || '| | ';
        v_普通结算 := Nvl(v_普通结算, '') || '||' || v_收费结算;
      End If;
      n_Count := n_Count + 1;
    End Loop;
    If n_Count = 0 Then
      v_Err_Msg := '不能有效确认当前的支付方式!';
      Raise Err_Item;
    End If;
    --5.普通结算及完成结
    If v_普通结算 Is Not Null Then
      v_普通结算 := Substr(v_普通结算, 3);
    End If;
    Zl_门诊收费结算_Modify(0, n_病人id, n_结帐id, v_普通结算, Null, 0, Null, Null, Null, Null, 0, 0, n_误差额, 1);
  Else
    n_结帐金额 := 0;
    --挂号单据
    For c_费用 In (Select 1 As 顺序号, b.No, b.收据费目, b.结帐id, b.执行部门id, b.病人科室ID, b.开单人, b.收费类别, b.收入项目id,
                        To_Char(b.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.价格父号, b.从属父号, b.序号, b.收费细目id, b.计算单位,
                        Max(m.名称) As 名称, Max(m.规格) As 规格, Sum(b.标准单价) As 单价, Avg(Nvl(b.付数, 1) * b.数次) As 数量,
                        Sum(b.应收金额) As 应收金额, Sum(b.实收金额) As 实收金额, Max(j.名称) As 开单科室, Max(q.名称) As 执行科室
                 From 门诊费用记录 B, 收费项目目录 M, 部门表 J, 部门表 Q
                 Where b.No = v_Nos And b.记录性质 = 4 And
                       b.收费细目id Not In (Select 收费细目id From 收费特定项目 Where 特定项目 = '病历费') And Nvl(b.费用状态, 0) = 0 And
                       b.记录状态 = 0 And b.收费细目id = m.Id And b.开单部门id = j.Id(+) And b.执行部门id = q.Id(+)
                 Group By b.No, b.收据费目, b.结帐id, b.执行部门id, b.病人科室ID, b.开单人, b.收入项目id, b.收费类别, b.登记时间, b.价格父号, b.从属父号, b.序号,
                          b.收费细目id, b.计算单位
                 Order By 序号) Loop
      Zl_病人预约挂号记录_Update(c_费用.No, c_费用.序号, c_费用.价格父号, c_费用.从属父号, c_费用.收费类别, c_费用.收费细目id, c_费用.数量, c_费用.单价, c_费用.收入项目id,
                         c_费用.收据费目, c_费用.应收金额, c_费用.实收金额, 0, Null, Null, Null, Null, c_费用.病人科室ID, c_费用.执行部门id);
      n_结帐金额 := n_结帐金额 + c_费用.实收金额;
    End Loop;
  
    --检查总金额是否正确 
    If Nvl(n_误差额, 0) = 0 Then
      n_误差额 := Nvl(n_收费总额, 0) - Nvl(n_结帐金额, 0);
      If Abs(n_误差额) > 1.00 Then
        v_Err_Msg := '单据缴款金额与实际结算的误差太大!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_结帐金额, 0) + Nvl(n_误差额, 0) <> Nvl(n_收费总额, 0) Then
      v_Err_Msg := '指定的缴费单据的总金额不对,请重新选择缴费单据!';
      Raise Err_Item;
    End If;
  
    For c_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(c_结算方式.是否冲预交, 0) = 1 Then
        n_预交支付 := c_结算方式.结算金额;
      Else
        If Nvl(n_普通支付, 0) = 0 Then
          n_普通支付 := c_结算方式.结算金额;
          v_结算方式 := c_结算方式.结算方式;
          If Nvl(c_结算方式.是否消费卡, 0) = 1 Then
            Begin
              n_结算卡序号 := To_Number(c_结算方式.结算卡类别);
            Exception
              When Others Then
                n_结算卡序号 := 0;
            End;
            If n_结算卡序号 = 0 Then
              Begin
                Select 编号
                Into n_结算卡序号
                From 卡消费接口目录
                Where 名称 = c_结算方式.结算卡类别 And Nvl(启用, 0) = 1;
              Exception
                When Others Then
                  v_Err_Msg := '未找到对应的消费卡!';
                  Raise Err_Item;
              End;
            End If;
            If v_结算方式 Is Null Then
              Select 结算方式 Into v_结算方式 From 卡消费接口目录 Where 编号 = n_结算卡序号;
            End If;
          Else
            Begin
              n_卡类别id := To_Number(c_结算方式.结算卡类别);
            Exception
              When Others Then
                n_卡类别id := 0;
            End;
            If n_卡类别id = 0 Then
              Begin
                Select ID Into n_卡类别id From 医疗卡类别 Where 名称 = c_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
              Exception
                When Others Then
                  v_Err_Msg := '未找到对应的医疗卡!';
                  Raise Err_Item;
              End;
            End If;
            If v_结算方式 Is Null Then
              Select 结算方式 Into v_结算方式 From 医疗卡类别 Where ID = n_卡类别id;
            End If;
          End If;
          v_结算卡号   := c_结算方式.结算卡号;
          v_交易流水号 := c_结算方式.交易流水号;
          v_交易说明   := c_结算方式.交易说明;
          v_摘要       := c_结算方式.摘要;
        Else
          v_Err_Msg := '挂号结算暂不支持多种结算方式!';
          Raise Err_Item;
        End If;
      End If;
    End Loop;
  
    --预约接收
    Select a.执行部门id, a.收费细目id, c.Id, a.执行人, b.号别, b.门诊号, b.发生时间, a.费别, b.号序
    Into n_科室id, n_项目id, n_医生id, v_医生姓名, v_号码, n_门诊号, d_发生时间, v_费别, n_号序
    From 门诊费用记录 A, 病人挂号记录 B, 人员表 C
    Where a.No = v_Nos And a.记录性质 = 4 And a.序号 = 1 And a.No = b.No And a.执行人 = c.姓名(+);
    Select Decode(To_Number(zl_GetSysParameter('排队叫号模式', 1113, 100)), 0, 0, 1) Into n_生成队列 From Dual;
  
    Zl_预约挂号接收_Insert(v_Nos, Null, Null, n_结帐id, Zl_诊室(v_号码), n_病人id, n_门诊号, v_姓名, v_性别, v_年龄, v_医疗付款方式编码, v_费别, v_结算方式,
                     n_普通支付, n_预交支付, Null, d_发生时间, n_号序, v_操作员编码, v_操作员姓名, n_生成队列, d_收费时间, n_卡类别id, n_结算卡序号, v_结算卡号,
                     v_交易流水号, v_交易说明, Null, 0, 0, Null, 1);
    --处理扩展信息
    If Nvl(n_卡类别id, 0) <> 0 Then
      For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                            Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
        Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
      End Loop;
    End If;
    If Nvl(n_结算卡序号, 0) <> 0 Then
      For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                            Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
        Zl_三方结算交易_Insert(n_结算卡序号, 1, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
      End Loop;
    End If;
    --处理汇总
    Zl_病人挂号汇总_Update(v_医生姓名, n_医生id, n_项目id, n_科室id, d_发生时间, 2, v_号码, 1);
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_收费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Payment;
/

--89706:刘尔旋,2015-11-04,接口返回登记时间
Create Or Replace Procedure Zl_Third_Registdel
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS退号
  --入参:Xml_In:
  --<IN>
  --  <GHDH>A000001</GHDH>    //挂号单号
  --  <JSKLB>支付宝</JSKLB>      //结算卡类别
  --  <JCFP>1</JCFP>            //检查发票
  --  <GHJE>20</GHJE>            //挂号金额
  --  <LSH>34563</LSH>           //交易流水号
  --  <JKFS>0</JKFS>             //缴款方式,0-挂号或预约缴款;1-预约不缴款
  --  <YYFS></YYFS>              //缴款方式=1时传入，预约的预约方式
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  -- <ERROR><MSG></MSG></ERROR> //为空表示取消挂号成功
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_卡类别     Varchar2(100);
  v_No         病人挂号记录.No%Type;
  n_挂号金额   门诊费用记录.实收金额%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  n_存在       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --临时XML
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  n_已开医嘱   Number(2);
  n_检查发票   Number(3);
  n_是否打印   Number(3);
  n_缴款方式   Number(3);
  d_登记时间   Date;
  v_预约方式   病人挂号记录.预约方式%Type;
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS')
  Into v_No, v_卡类别, n_挂号金额, v_交易流水号, n_检查发票, n_缴款方式, v_预约方式
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  n_缴款方式 := Nvl(n_缴款方式, 0);

  If v_卡类别 Is Not Null And n_缴款方式 = 0 Then
    Select Nvl2(Translate(v_卡类别, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
    If v_Type = 'Num' Then
      --传入的是卡类别ID
      Select 结算方式 Into v_结算方式 From 医疗卡类别 Where ID = To_Number(v_卡类别);
    Else
      --传入的是卡类别名称
      Select 结算方式 Into v_结算方式 From 医疗卡类别 Where 名称 = v_卡类别;
    End If;
  
    Select Sum(实收金额) Into n_实收金额 From 门诊费用记录 Where NO = v_No And 记录性质 = 4;
  
    If Nvl(n_缴款方式, 0) = 0 Then
      --要退的单据不是以该结算卡结算的，则禁止退号
      Begin
        Select 1
        Into n_存在
        From 病人预交记录 A,
             (Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_No And 记录性质 = 4
               Union
               Select Distinct 结帐id From 住院费用记录 Where NO = v_No And 记录性质 = 5) B
        Where a.结帐id = b.结帐id And 结算方式 = v_结算方式 And Rownum < 2;
      Exception
        When Others Then
          n_存在 := 0;
      End;
      If n_存在 = 0 Then
        v_Err_Msg := '传入的挂号单据不是' || v_结算方式 || '结算的,无法退号!';
        Raise Err_Item;
      End If;
    Else
      Begin
        Select 1 Into n_存在 From 病人挂号记录 A Where a.No = v_No And a.预约方式 = v_预约方式 And Rownum < 2;
      Exception
        When Others Then
          n_存在 := 0;
      End;
      If n_存在 = 0 Then
        v_Err_Msg := '传入的挂号单据不是' || v_预约方式 || '预约的,无法退号!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --补充结算检查，已存在补结算数据的，不能退号
  Begin
    Select 1
    Into n_存在
    From 费用补充记录 A,
         (Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_No And 记录性质 = 4
           Union
           Select Distinct 结帐id From 住院费用记录 Where NO = v_No And 记录性质 = 5) B
    Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 1 And Nvl(a.费用状态, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_存在 := 0;
  End;
  If n_存在 = 1 Then
    v_Err_Msg := '传入的挂号单据已经进行了二次结算,无法退号!';
    Raise Err_Item;
  End If;
  --医嘱检查，已经开过医嘱的，不能退号
  Begin
    Select Distinct 1 Into n_已开医嘱 From 病人医嘱记录 Where 挂号单 = v_No;
  Exception
    When Others Then
      n_已开医嘱 := 0;
  End;
  If n_已开医嘱 = 1 Then
    v_Err_Msg := '传入的挂号单据已经开过医嘱,无法退号!';
    Raise Err_Item;
  End If;
  If Nvl(n_检查发票, 0) = 1 Then
    Select Max(Decode(a.实际票号, Null, 0, 1)) Into n_是否打印 From 门诊费用记录 A Where NO = v_No And 记录性质 = 4;
    If Nvl(n_是否打印, 0) = 1 Then
      v_Err_Msg := '本次退号的单据已开发票,不能退费!';
      Raise Err_Item;
    End If;
  End If;
  --获取操作员信息
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
  d_登记时间 := Sysdate;
  Zl_三方机构挂号_Delete(v_No, v_交易流水号, '移动平台退号', d_登记时间);

  v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
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

--90240:李南春,2015-11-04,f_Str2list联表查询加Rule标志
CREATE OR REPLACE Procedure Zl_凭条打印记录_Update
(
    记录性质_In 凭条打印记录.记录性质%Type,
    Nos_In      Varchar2,
    打印类型_In 凭条打印记录.打印类型%Type := Null,
    打印人_In   凭条打印记录.打印人%Type := Null,
    备注_In     凭条打印记录.备注%Type := Null
) As
    n_Count    Number(18);
    n_打印类型 Number(3);
Begin
    --如果已经存在记录则打印类型变为2
    Select /*+ Rule*/ Count(*) Into n_Count From 凭条打印记录 A,(Select Column_Value As No From Table(f_Str2list(Nos_In))) B Where 记录性质 = 记录性质_In And A.NO=B.No ;
    n_打印类型 := 打印类型_In;
    If n_Count > 0 Then
        n_打印类型 := 2;
    End If;
    Insert Into 凭条打印记录
        (记录性质, NO, 打印时间, 打印类型, 打印人, 机器名, Ip地址, 备注)
        Select 记录性质_In, Column_Value As NO, Sysdate, n_打印类型, 打印人_In, Sys_Context('userenv', 'HOST'),
               Sys_Context('USERENV', 'IP_ADDRESS'), 备注_In
        From Table(f_Str2list(Nos_In));
Exception
    When Others Then
        zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_凭条打印记录_Update;
/

--89706:刘尔旋,2015-11-04,接口返回登记时间
Create Or Replace Procedure Zl_Third_Regist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS挂号
  --入参:Xml_In:
  --<IN>
  --   <CZFS>3</CZFS>    //操作方式
  --   <HM>号码</HM>    //号码
  --   <HX>号序</HX>     //号序
  --   <JKFS>0</JKFS>  //缴款方式,0-挂号或预约缴款;1-预约不缴款
  --   <YYSJ>2014-10-21 </YYSJ>    //预约日期 YYYY-MM-DD,分时段非序号控制需要传入时间
  --   <JE>金额</JE>     //金额
  --   <JSLIST>
  --     <JS>            //结算信息，挂号目前仅支持一个，结构与收费一致，以后可扩展
  --       <JSKLB>结算卡类别</JSKLB>    //结算卡类别
  --       <JSKH>支付宝帐号</JSKH>           //结算卡号(支付宝帐号)
  --       <JYSM>交易说明</JYSM>            //说明，固定传支付宝
  --       <JYLSH>流水号</JYLSH>           //流水号，传订单号
  --       <JSFS>结算方式</JSFS>            //结算方式:现金、支票，如果是三方卡,可以传空
  --       <JSJE>结算金额</JSJE>            //结算金额
  --       <ZY>摘要</ZY>                  //摘要
  --       <SFCYJ></SFCYJ>              //是否冲预交，挂号目前不传
  --       <SFXFK></SFXFK>              //是否消费卡,挂号目前不传
  --       <EXPENDLIST>                 //扩展信息
  --         <EXPEND>
  --           <JYMC>交易名称</JYMC>        //交易名称
  --           <JYLR>交易内容<JYLR>         //交易内容
  --         </EXPEND>
  --         <EXPEND>
  --           ...
  --         </EXPEND>
  --       </EXPENDLIST>
  --     </JS>
  --   </JSLIST>
  --   <HZDW>合作单位</HZDW>        //合作单位名称
  --   <YYFS>支付宝<YYFS>    //预约方式,如自助机，支付宝
  --   <BRID>病人ID</BRID>     //病人ID
  --   <BRLX></BRLX>             //医保病人类型
  --   <FB>普通</FB>               //病人费别，可以不传
  --   <JQM>机器名</JQM>            //机器名
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <GHDH>挂号单号</GHDH>          //挂号单号
  -- <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  -- <ERROR><MSG>错误信息</MSG></ERROR>  //出错时返回
  --</ OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_号码       挂号安排.号码%Type;
  d_发生时间   Date;
  d_原始时间   Date;
  d_登记时间   Date;
  v_金额       Varchar2(200);
  n_应收金额   门诊费用记录.应收金额%Type;
  v_流水号     病人预交记录.交易流水号%Type;
  v_说明       门诊费用记录.摘要%Type;
  n_病人id     病人信息.病人id%Type;
  v_预约方式   预约方式.名称%Type;
  v_卡类别名称 医疗卡类别.名称%Type;
  v_结算卡号   病人预交记录.卡号%Type;
  n_门诊号     门诊费用记录.标识号%Type;
  v_姓名       门诊费用记录.姓名%Type;
  v_性别       门诊费用记录.性别%Type;
  v_年龄       门诊费用记录.年龄%Type;
  v_付款方式   门诊费用记录.付款方式%Type;
  v_费别       门诊费用记录.费别%Type;
  v_No         病人挂号记录.No%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  v_收费类别   门诊费用记录.收费类别%Type;
  n_收费细目id 门诊费用记录.收费细目id%Type;
  n_标准单价   门诊费用记录.标准单价%Type;
  n_收入项目id 门诊费用记录.收入项目id%Type;
  n_屏蔽费别   收费项目目录.屏蔽费别%Type;
  v_收据费目   门诊费用记录.收据费目%Type;
  n_病人科室id 门诊费用记录.病人科室id%Type;
  n_开单部门id 门诊费用记录.开单部门id%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_医生姓名   挂号安排.医生姓名%Type;
  n_医生id     挂号安排.医生id%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_排班       挂号安排.周日%Type;
  n_安排id     挂号安排.Id%Type;
  n_计划id     挂号安排计划.Id%Type;
  n_预交id     病人预交记录.Id%Type;
  n_序号控制   挂号安排.序号控制%Type;
  n_号序       挂号序号状态.序号%Type;
  v_星期       挂号安排限制.限制项目%Type;
  v_病人类型   病人信息.病人类型%Type;
  n_存在       Number(3);
  n_分时段     Number(3);
  v_合作单位   病人挂号记录.合作单位%Type;
  v_机器名     挂号序号状态.机器名%Type;
  n_缴款方式   Number(3);
  v_Temp       Varchar2(32767); --临时XML
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS')
  Into v_号码, n_号序, d_原始时间, n_应收金额, v_预约方式, v_合作单位, n_病人id, v_病人类型, v_费别, v_机器名, n_缴款方式
  
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Extractvalue(b.Column_Value, '/JS/JSKLB'), Extractvalue(b.Column_Value, '/JS/JSKH'),
         Extractvalue(b.Column_Value, '/JS/JSFS'), Extractvalue(b.Column_Value, '/JS/JYLSH'),
         Extractvalue(b.Column_Value, '/JS/JYSM')
  Into v_卡类别名称, v_结算卡号, v_结算方式, v_流水号, v_说明
  From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B;

  d_登记时间 := Sysdate;
  d_发生时间 := Trunc(d_原始时间);
  If v_病人类型 Is Not Null Then
    Begin
      Select 1 Into n_存在 From 病人类型 Where 名称 = v_病人类型;
    Exception
      When Others Then
        v_Err_Msg := '没有发现为(' || v_病人类型 || ')的病人类型';
        Raise Err_Item;
    End;
    Update 病人信息 Set 病人类型 = Nvl(病人类型, v_病人类型) Where 病人id = n_病人id;
  End If;
  Begin
    Select b.结算方式, b.Id Into v_结算方式, n_卡类别id From 医疗卡类别 B Where b.名称 = v_卡类别名称 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := '没有发现该结算卡的相关信息';
      Raise Err_Item;
  End;

  Select a.门诊号, a.姓名, a.性别, a.年龄, Nvl(b.编码, c.编码)
  Into n_门诊号, v_姓名, v_性别, v_年龄, v_付款方式
  From 病人信息 A, 医疗付款方式 B, (Select 编码 From 医疗付款方式 Where 缺省标志 = '1' And Rownum < 2) C
  Where a.病人id = n_病人id And a.医疗付款方式 = b.名称(+);
  v_No   := Nextno(12);
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
  v_Temp := Zl_Identity(2);
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into n_开单部门id From Dual;
  Select Decode(To_Char(d_发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;
  Begin
    Select ID
    Into n_计划id
    From (Select ID
           From 挂号安排计划
           Where 号码 = v_号码 And d_发生时间 Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                 Nvl(失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And 审核时间 Is Not Null
           Order By 生效时间 Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      Select ID Into n_安排id From 挂号安排 Where 号码 = v_号码;
  End;
  If Nvl(n_计划id, 0) <> 0 Then
    --从计划读取信息
    Select a.项目id, b.科室id, a.医生姓名, a.医生id,
           Decode(To_Char(d_发生时间, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7', a.周六,
                   Null), Nvl(a.序号控制, 0)
    Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, v_排班, n_序号控制
    From 挂号安排计划 A, 挂号安排 B
    Where a.Id = n_计划id And b.Id = a.安排id;
    Select Count(Rownum) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum < 2;
    --合作单位检查
    If v_合作单位 Is Not Null Then
      Begin
        Select 1 Into n_存在 From 合作单位计划控制 Where 计划id = n_计划id And 数量 = 0 And 合作单位 = v_合作单位;
      Exception
        When Others Then
          n_存在 := 0;
      End;
    End If;
    If n_存在 = 1 Then
      v_Err_Msg := '传入的合作单位在此号码上被禁用！';
      Raise Err_Item;
    End If;
    If n_分时段 = 1 And n_序号控制 = 0 Then
      d_发生时间 := d_原始时间;
      Select 序号
      Into n_号序
      From 挂号计划时段
      Where 计划id = n_计划id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi:ss') = To_Char(d_发生时间, 'hh24:mi:ss');
    Else
      Begin
        Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
        Into d_发生时间
        From 挂号计划时段
        Where 计划id = n_计划id And 星期 = v_星期 And 序号 = Nvl(n_号序, 0);
      Exception
        When Others Then
          If n_分时段 = 1 And n_序号控制 = 1 Then
            Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(Max(结束时间), 'hh24:mi:ss'),
                            'YYYY-MM-DD hh24:mi:ss')
            Into d_发生时间
            From 挂号计划时段
            Where 计划id = n_计划id And 星期 = v_星期;
          Else
            Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
            Into d_发生时间
            From 时间段
            Where 时间段 = v_排班;
          End If;
          If d_发生时间 < d_登记时间 Then
            d_发生时间 := d_登记时间;
          End If;
      End;
    End If;
  Else
    --从安排读取信息
    Select b.项目id, b.科室id, b.医生姓名, b.医生id,
           Decode(To_Char(d_发生时间, 'D'), '1', b.周日, '2', b.周一, '3', b.周二, '4', b.周三, '5', b.周四, '6', b.周五, '7', b.周六,
                   Null), Nvl(b.序号控制, 0)
    Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, v_排班, n_序号控制
    From 挂号安排 B
    Where b.Id = n_安排id;
    Select Count(Rownum) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum < 2;
    --合作单位检查
    If v_合作单位 Is Not Null Then
      Begin
        Select 1 Into n_存在 From 合作单位安排控制 Where 安排id = n_安排id And 数量 = 0 And 合作单位 = v_合作单位;
      Exception
        When Others Then
          n_存在 := 0;
      End;
    End If;
    If n_存在 = 1 Then
      v_Err_Msg := '传入的合作单位在此号码上被禁用！';
      Raise Err_Item;
    End If;
    If n_分时段 = 1 And n_序号控制 = 0 Then
      d_发生时间 := d_原始时间;
      Select 序号
      Into n_号序
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi:ss') = To_Char(d_发生时间, 'hh24:mi:ss');
    Else
      Begin
        Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
        Into d_发生时间
        From 挂号安排时段
        Where 安排id = n_安排id And 星期 = v_星期 And 序号 = Nvl(n_号序, 0);
      Exception
        When Others Then
          If n_分时段 = 1 And n_序号控制 = 1 Then
            Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(Max(结束时间), 'hh24:mi:ss'),
                            'YYYY-MM-DD hh24:mi:ss')
            Into d_发生时间
            From 挂号安排时段
            Where 安排id = n_安排id And 星期 = v_星期;
          Else
            Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
            Into d_发生时间
            From 时间段
            Where 时间段 = v_排班;
          End If;
          If d_发生时间 < d_登记时间 Then
            d_发生时间 := d_登记时间;
          End If;
      End;
    End If;
  End If;

  Select a.类别, b.现价, b.收入项目id, c.收据费目, a.屏蔽费别
  Into v_收费类别, n_标准单价, n_收入项目id, v_收据费目, n_屏蔽费别
  From 收费项目目录 A, 收费价目 B, 收入项目 C
  Where a.Id = n_收费细目id And b.收费细目id = a.Id And b.收入项目id = c.Id And Sysdate Between b.执行日期 And
        Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Rownum < 2;

  Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  Select 病人预交记录_Id.Nextval Into n_预交id From Dual;

  If Trunc(d_发生时间) <> Trunc(Sysdate) Then
    If Nvl(n_缴款方式, 0) = 0 Then
      Zl_三方机构挂号_Insert(3, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                       v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, Null, v_结算卡号, 1, v_费别, v_机器名, 1);
    Else
      Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                       v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, Null, v_结算卡号, 1, v_费别, v_机器名, 1);
    End If;
  Else
    Zl_三方机构挂号_Insert(1, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null, v_流水号,
                     v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, Null, v_结算卡号, 1, v_费别, v_机器名, 1);
  End If;

  For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                        Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
    Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
  End Loop;

  v_Temp := '<GHDH>' || v_No || '</GHDH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/

--91029:黄捷,2015-11-26,修改影像报告操作记录
--90217:黄捷,2015-11-03,PACS报告文档编辑器嵌入式报告增加tab
--影像报告业务(---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_RptManage Is
  Type t_Refcur Is Ref Cursor;

  --1、锁定报告人
  Procedure p_Edit_Doc_Lockinfo(
    报告_Id_In 影像报告记录.Id%Type,
	锁定人_In  影像报告记录.锁定人%Type
	);

  --2、评定报告质量
  Procedure p_Edit_Doc_EvaluatRptQuality(
    报告Id_In 影像报告记录.Id%Type,
	质量等级_In 影像报告记录.报告质量%Type
	);
                                
  --3、评定阴阳性
  Procedure p_Edit_Doc_EvaluatResult(
    报告Id_In 影像报告记录.Id%Type,
	检查结果_In 影像报告记录.结果阳性%Type
	);
                                
  --4、报告发放/回收
  Procedure p_Edit_Doc_ReportRelease(
    报告Id_In 影像报告记录.Id%Type,
	当前操作人_In 影像报告记录.报告发放人%Type
	);

 --5、新增，修改报告
  Procedure p_影像报告记录_新增(
    原型ID_In     影像报告记录.原型ID%Type,
    报告内容_In   影像报告记录.报告内容%Type,
    记录人_In     影像报告记录.记录人%Type,
    最后编辑人_In 影像报告记录.最后编辑人%Type,
    Id_In         影像报告记录.Id%Type,
    医嘱ID_In     影像报告记录.医嘱ID%Type 
	);

  --6、获取书写的文档内容
  Procedure p_Get_Doc_Content(
    Val           Out t_Refcur,
	DocID_In 影像报告记录.Id%Type
	);

  --7、设置报告打印作废信息
  Procedure p_Checkrejectsignature(Signdate_In Date,
                                   报告ID_In   影像报告操作记录.报告Id%Type,
                                   作废人_In   影像报告操作记录.作废人%Type,
                                   作废说明_In 影像报告操作记录.作废说明%Type,
                                   Val         Out Sys_Refcursor);

  --8、查询相应原型下的最大序号
  Procedure p_Get_Samplelist_Maxseqnum(
    Val           Out t_Refcur,
	原型ID_In 影像报告范文清单.原型ID%Type
	);

  --9、删除文档范文
  Procedure p_Del_影像报告范文清单(
    Id_In 影像报告范文清单.Id%Type
	);

 --10、添加文档的操作日志
 Procedure p_影像报告操作记录_Add(Id_In       影像报告操作记录.Id%Type,
                               报告ID_In   影像报告操作记录.报告ID%Type,                               
                               操作人_In   影像报告操作记录.操作人%Type,                               
                               操作类型_In 影像报告操作记录.操作类型%Type);

  --11、删除报告
  Procedure p_影像报告记录_删除(
    报告_Id_In 影像报告记录.Id%Type
	);

  --12、获取签名类型
  Procedure p_Get_SysConfigSignature(
    Val           Out t_Refcur,
	科室ID_In		In 部门表.ID%Type
	);

--13、获取账户签名印章
Procedure p_Get_PersonSignImg(
  Val           Out t_Refcur,
  ID_In		In 人员表.ID%Type
  );


--14、获取签名的证书信息
Procedure p_Get_SignCertInfo(
  Val           Out t_Refcur,
  证书ID_In		人员证书记录.ID%Type
  );

--15、更新报告状态
Procedure p_Update_ReportState(
  报告Id_In  影像报告记录.ID%Type,
  报告状态_In  影像报告记录.报告状态%Type,
  审核人_In   影像报告记录.最后审核人%Type
  );

--16、获取报告状态
Procedure p_Get_ReportState(
  Val           Out t_Refcur,
  报告Id_In	影像报告记录.ID%Type
  );

--17、报告驳回
Procedure p_Reject_Report(
  医嘱ID_In	影像报告驳回.医嘱ID%Type, 
  报告ID_In	影像报告驳回.检查报告ID%Type, 
  驳回理由_In 影像报告驳回.驳回理由%Type, 
  驳回时间_In 影像报告驳回.驳回时间%Type, 
  驳回人_In   影像报告驳回.驳回人%Type,
  待处理人_In  影像报告记录.待处理人%Type,
  报告状态_In 影像报告记录.报告状态%Type
  );

--17.1、撤销报告驳回
Procedure p_Reject_Cancel(
  ID_In       影像报告驳回.ID%Type,
  报告ID_In    影像报告驳回.检查报告ID%Type,
  报告状态_In   影像报告记录.报告状态%Type
  );

--18、获取报告驳回信息
Procedure p_Get_RejectInfo(
  Val           Out t_Refcur,
  报告ID_In	影像报告驳回.检查报告ID%Type
  );

--19、获取原型动作
Procedure p_Get_Doc_Process(
  Val           Out t_Refcur,
  原型id_In 影像报告动作.原型id%Type
  );

--20、通过学科筛选获得相应的范文信息
  Procedure p_Get_Samplelist_By_Conditions(
    Val           Out t_Refcur,
    原型id_In       Varchar2,
    学科_In  Varchar2,
    Condition_In Varchar2, --过滤筛选
    作者_In    Varchar2
    );

  --21、通过部门ID获取部门名称
  Procedure p_Get_部门名称_By_ID(
    Val           Out t_Refcur,
    ID_IN 部门表.ID%TYPE
    );

  --22、提取所有预备提纲
  Procedure p_Get_AllPreOutlines(
    Val           Out t_Refcur
    );

  --23、提取文档标题
  Procedure p_Get_reportTitle_By_ID(
    Val           Out t_Refcur,
    ID_IN  影像报告记录.id%TYPE   
    );
  
  --24、提取报告锁定人
  Procedure p_Get_报告锁定人_By_ID(
    Val           Out t_Refcur,
    ID_IN  影像报告记录.id%TYPE   
    );

  --25、通过医嘱ID获取报告列表
  Procedure p_Get_影像报告记录_By_医嘱ID(
    Val           Out t_Refcur,
    医嘱ID_IN  影像报告记录.医嘱ID%TYPE   
    );
  
  --26、查询影像流程参数值
  Procedure p_Get_影像流程参数值(
    Val           Out t_Refcur,  
    科室ID_IN  影像流程参数.科室ID%TYPE
    );

  --27、根据医嘱ID，查询对应的原型列表
  Procedure p_Get_影像原型列表_By_医嘱ID(
    Val           Out t_Refcur,
    医嘱_IN  影像检查记录.医嘱ID%TYPE   
    );

  --28、根据报告ID查询打印记录
  procedure p_Get_ReportPrintLog_By_报告ID
  (
       val out sys_refcursor  ,
       报告_IN  影像报告操作记录.报告ID%TYPE
  );

  --29、根据医嘱ID查询报告发放列表
  Procedure p_Get_ReportReleaseList(
    Val           Out t_Refcur,
    医嘱_IN  影像报告记录.医嘱ID%TYPE   
    );

  --30、根据报告ID查询驳回记录数量
  Procedure p_Get_RejectedCount(
    Val           Out t_Refcur,
    报告_IN  影像报告驳回.检查报告ID%TYPE
    );

  --31、根据医嘱ID查询报告动作需要的一些ID们
  Procedure p_Get_DocProcess_IDs(
    Val           Out t_Refcur,
    医嘱_IN  病人医嘱记录.ID%TYPE
    );

  --32、根据医嘱ID和报告ID查询报告的一些参数
  Procedure p_Get_DocInfo(
    Val           Out t_Refcur,
    医嘱ID_IN  影像检查记录.医嘱ID%TYPE,
    报告ID_IN  影像报告记录.ID%TYPE
    );
  
  --33、查询一个检查中相同原型ID的报告数量
   Procedure p_Get_SameAntetypeDocCounts(
       Val           Out t_Refcur,
       医嘱ID_IN  影像报告记录.医嘱ID%TYPE,
       原型ID_IN  影像报告记录.原型ID%TYPE
  );

  --34、提取报告图存储信息
  Procedure p_Get_DocImageSaveInof_By_ID(
    Val           Out t_Refcur,
	  ID_IN  影像报告记录.id%TYPE
    );

end b_PACS_RptManage;
/

--影像报告业务(---实现部分---)***************************************************

CREATE OR REPLACE Package Body b_PACS_RptManage Is

  --1、锁定报告人
  Procedure p_Edit_Doc_Lockinfo(
    报告_Id_In 影像报告记录.Id%Type,
	锁定人_In  影像报告记录.锁定人%Type
	) Is
  Begin
  
    --  报告ID为空，则清空所有“锁定人_In”正在锁定的标记
    If 报告_Id_In Is Null Then
      Update 影像报告记录 A Set a.锁定人 = '' Where a.锁定人 = 锁定人_In;
    Else
      Update 影像报告记录 A
         Set a.锁定人 = 锁定人_In
       Where a.Id = 报告_Id_In;
    End If;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Lockinfo;

  --2、评定报告质量
  Procedure p_Edit_Doc_EvaluatRptQuality(
    报告Id_In 影像报告记录.Id%Type,
	质量等级_In  影像报告记录.报告质量%Type
	) Is
  Begin
    Update 影像报告记录 Set 报告质量 = 质量等级_In Where Id = 报告Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_EvaluatRptQuality;
  
  --3、评定阴阳性
  Procedure p_Edit_Doc_EvaluatResult(
    报告Id_In 影像报告记录.Id%Type,
	检查结果_In 影像报告记录.结果阳性%Type
	) Is
  Begin
     Update 影像报告记录 Set 结果阳性 = 检查结果_In Where Id = 报告Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_EvaluatResult;
  
  --4、报告发放/回收
  Procedure p_Edit_Doc_ReportRelease(
    报告Id_In 影像报告记录.Id%Type,
	当前操作人_In 影像报告记录.报告发放人%Type
	) Is
    v_报告发放     影像报告记录.报告发放%Type; 
  Begin
    
    Begin 
		  Select nvl(报告发放,0) Into v_报告发放 From 影像报告记录 where ID=报告Id_In; 
    Exception 
      When Others Then 
        v_报告发放 :=0; 
    End; 
     
    Update 影像报告记录 Set 报告发放 =decode(v_报告发放,0,1,0),报告发放人=decode(v_报告发放,0,当前操作人_In,'') Where ID=报告Id_In; 
     
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_ReportRelease; 

 --5、新增，修改报告
  Procedure p_影像报告记录_新增(
    原型ID_In     影像报告记录.原型ID%Type,
    报告内容_In   影像报告记录.报告内容%Type,
    记录人_In     影像报告记录.记录人%Type,
    最后编辑人_In 影像报告记录.最后编辑人%Type,
    Id_In         影像报告记录.Id%Type,
    医嘱ID_In     影像报告记录.医嘱ID%Type
  ) As
    --原型ID_In 原型ID
    --保存文档书写记录
    --1 处理匿名数据
    --2 保存文档书写记录、状态
    --3 处理编辑日志
    --4 更新文档任务
    v_报告id    影像报告记录.Id%Type;
    v_原型名称  影像报告原型清单.名称%Type;
    v_设备号    影像报告原型清单.设备号%Type;
    v_报告序号  number;
    x_Editlog   Xmltype;
    Cur_Time    Date;
    To_Editlist t_Editlist;
    Tn_Editlist t_Editlist;
    v_Msg       Varchar2(200);
    v_New       number;
    Err_Custom  Exception;
    v_Result    影像报告记录.诊断意见%Type;
    v_操作ID    影像报告操作记录.ID%Type;

    Function Elist_Filter(
    Source_t t_Editlist
    ) Return t_Editlist Is
      Target_t t_Editlist := t_Editlist();
    Begin

      --对独立文档来说，这个函数只是将 Source_t按照编辑时间排序后输出
      For Rs In (Select /*+rule*/
                  *
                   From Table(Cast(Source_t As t_Editlist)) A
                  Order By a.编辑时间) Loop
        Target_t.Extend;
        Target_t(Target_t.Count) := t_Edits(Rs.编辑人,
                                            Rs.编辑时间,
                                            Rs.签名,
                                            Rs.审订签名);
      End Loop;
      Return Target_t;
    End;

    Function Build_Editlog(
    Tn_Edit t_Editlist,
    To_Edit t_Editlist,
    v_Did   影像报告记录.Id%Type) Return Xmltype Is
      --Tn_Edit 本次保存的新编辑记录；To_Edit上次保存的旧编辑记录
      --将两次编辑记录，组合成一个编辑记录

      x_Return Xmltype;
      r_Saveid Raw(16);
      n_Class  Number;
      --n_Class 编辑日志中的操作类别： 1-创建、2-删除、3-编辑、4-签名、5-审订、6-审签、7-撤签
      v_Signor  影像报告记录.创建人%Type;
      v_Adjunct 影像报告记录.创建人%Type;
      Tns_Edit  t_Editlist;
      Tos_Edit  t_Editlist;

      Function Atitle(原型ID 影像报告原型清单.Id%Type) Return Varchar2 Is
        v_原型名称 影像报告原型清单.名称%Type;
      Begin
        --根据原型ID，返回原型名称
        If 原型ID Is Null Then
          Return Null;
        Else
          Select 名称 Into v_原型名称 From 影像报告原型清单 Where ID = 原型ID;
          Return v_原型名称;
        End If;
      End;

    Begin
      x_Return := Xmltype('<root></root>');
      If v_Did Is Null Then
        --表明是新增文档，新增文档传null进来
        Select Sys_Guid() Into r_Saveid From Dual;

        --PACS报告没有子文档，但是下面构造XML的语句保留成跟EMR相同，这里的v_Subiid赋值为空
        Tns_Edit := Elist_Filter(Tn_Edit);
        Select Decode(Tns_Edit(Tns_Edit.Count).签名, 0, 1, 4)
          Into n_Class
          From Dual;
        Select Appendchildxml(x_Return,
                              '/root',
                              Xmlelement("operate",
                                         Xmlforest(r_Saveid As "saving_id",
                                                   n_Class As "class",
                                                   To_Char(Cur_Time,
                                                           'yyyy-mm-dd hh24:mi:ss') As
                                                   "cur_time",
                                                   最后编辑人_In As "operator",
                                                   Decode(n_Class,
                                                          4,
                                                          Tns_Edit(Tns_Edit.Count).编辑人,
                                                          '') As "signer",
                                                   '' As Adjunct)))
          Into x_Return
          From Dual;
      Else
        --不是新增的文档？
        Select Sys_Guid() Into r_Saveid From Dual;

        v_Signor  := '';
        v_Adjunct := '';
        Tns_Edit  := Elist_Filter(Tn_Edit);
        Tos_Edit  := Elist_Filter(To_Edit);
        If Tns_Edit(Tns_Edit.Count)
         .签名 = 1 And Tns_Edit(Tns_Edit.Count).审订签名 = 0 Then
          --最近一次是签名
          If Tos_Edit.Count = 0 Then
            --新增子文档直接签名
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count).编辑人 Is Null Then
            --之前没签名
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count)
           .签名 = 1 And Tns_Edit(Tns_Edit.Count)
                .编辑时间 > Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --多次普通签名
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count)
           .签名 = 1 And Tns_Edit(Tns_Edit.Count)
                .编辑时间 < Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --撤消多次签名
            n_Class   := 7;
            v_Adjunct := Tos_Edit(Tos_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count)
           .签名 = 1 And Tns_Edit(Tns_Edit.Count)
                .编辑时间 = Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --无变化
            n_Class := -1;
          End If;
        Elsif Tns_Edit(Tns_Edit.Count)
         .签名 = 1 And Tns_Edit(Tns_Edit.Count).审订签名 = 1 Then
          --审订签名
          If Tos_Edit(Tos_Edit.Count).审订签名 = 0 Then
            --之前没审签，可能是已签名或已审订
            n_Class  := 6;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count)
           .审订签名 = 1 And Tns_Edit(Tns_Edit.Count)
                .编辑时间 > Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --多次审签
            n_Class  := 6;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count)
           .审订签名 = 1 And Tns_Edit(Tns_Edit.Count)
                .编辑时间 < Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --撤消多次审签
            n_Class   := 7;
            v_Adjunct := Tos_Edit(Tos_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count)
           .审订签名 = 1 And Tns_Edit(Tns_Edit.Count)
                .编辑时间 = Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --无变化
            n_Class := -1;
          End If;
        Elsif Tns_Edit(Tns_Edit.Count).编辑人 Is Null And Tos_Edit.Count = 0 Then
          n_Class := 1;
        Elsif Tns_Edit(Tns_Edit.Count)
         .编辑人 Is Null And Tos_Edit(Tos_Edit.Count).签名 = 1 Then
          n_Class   := 7;
          v_Adjunct := Tos_Edit(Tos_Edit.Count).编辑人;
        Elsif Tns_Edit(Tns_Edit.Count)
         .编辑人 Is Null And Tos_Edit(Tos_Edit.Count).编辑人 Is Null Then
          n_Class := 3;
        Elsif Tns_Edit(Tns_Edit.Count)
         .审订签名 = 0 And Tos_Edit(Tos_Edit.Count).审订签名 = 0 Then
          n_Class := 5;
        Elsif Tns_Edit(Tns_Edit.Count)
         .审订签名 = 0 And Tos_Edit(Tos_Edit.Count).审订签名 = 1 Then
          n_Class   := 7;
          v_Adjunct := Tos_Edit(Tos_Edit.Count).编辑人;
        End If;

        If n_Class <> -1 Then
          Select Appendchildxml(x_Return,
                                '/root',
                                Xmlelement("operate",
                                           Xmlforest(r_Saveid As "saving_id",
                                                     n_Class As "class",
                                                     To_Char(Cur_Time,
                                                             'yyyy-mm-dd hh24:mi:ss') As
                                                     "cur_time",
                                                     最后编辑人_In As "operator",
                                                     Decode(n_Class,
                                                            4,
                                                            v_Signor,
                                                            6,
                                                            v_Signor,
                                                            '') As "signer",
                                                     v_Adjunct As Adjunct)))
            Into x_Return
            From Dual;
        End If;

      End If;
      Return x_Return;
    End Build_Editlog;

    Function Get_NextRPTNum(
    AntetypeName 影像报告原型清单.名称%Type,
    Order_ID 影像报告记录.医嘱Id%Type
    )
      Return Number Is
        v_序号 Number;
        v_count Number;
        v_num Number;
      Begin

        v_count :=0;
        v_num :=1;
        loop
             select count(*)+v_num into v_序号 from 影像报告记录 where 医嘱ID=Order_ID;
             select count(*) into v_count from 影像报告记录 where 医嘱ID=Order_ID and 文档标题=AntetypeName||'_'||v_序号;

             if v_count =0 then
               exit;
             end if;

             v_num := v_num +1;
         end loop;

         return v_序号;
     End;

  Begin

    Select 名称, 设备号,Sysdate
      Into v_原型名称,v_设备号, Cur_Time
      From 影像报告原型清单
     Where ID = 原型ID_In;

    --------------------1 保存文档书写记录、状态--------------------
    --提取文档的签名和编辑（新增、修改）记录
    Tn_Editlist := b_PACS_RptPublic.f_Geteditlist(报告内容_In);

    --------------------2 处理编辑日志--------------------
    select count(*) into v_New from 影像报告记录 where ID=Id_In;

    v_报告id := Id_In;
    select zlpub_pacs_取提纲内容byxml (报告内容_In,'诊断意见') into v_Result from dual;
    If v_New=0 Then
      --新增报告
      To_Editlist := t_Editlist();
      x_Editlog   := Build_Editlog(Tn_Editlist, To_Editlist, Null);

      --取报告序号
      v_报告序号 := Get_NextRPTNum(v_原型名称,医嘱ID_In);

      Insert Into 影像报告记录
        (ID,
         原型ID,
         文档标题,
         报告内容,
         创建时间,
         创建人,
         报告状态,
         最后编辑时间,
         最后编辑人,
         编辑日志,
         医嘱ID,
         记录人,
         诊断意见,
         设备号)
      Values
        (v_报告id,
         原型ID_In,
         v_原型名称||'_'||v_报告序号,
         报告内容_In,
         Cur_Time,
         最后编辑人_In,
         1,
         Cur_Time,
         最后编辑人_In,
         x_Editlog,
         医嘱ID_In,
         记录人_In,
         v_Result,
         v_设备号);
      Insert Into 病人医嘱报告(医嘱ID,检查报告ID)Values(医嘱ID_In,v_报告id);
      
      Select Sys_Guid() Into v_操作ID From Dual;
      Insert Into 影像报告操作记录(ID, 报告ID,医嘱ID,文档标题,操作人,操作时间,操作类型) 
             Values(v_操作ID,v_报告id,医嘱ID_In,v_原型名称||'_'||v_报告序号,最后编辑人_In,sysdate,6);

    Else
      --提取文件原始编辑记录,必需在更新之前提取
      Select b_PACS_RptPublic.f_Geteditlist(报告内容)
        Into To_Editlist
        From 影像报告记录
       Where ID = v_报告id;

      x_Editlog := Build_Editlog(Tn_Editlist, To_Editlist, v_报告id);
      Select Appendchildxml(编辑日志,
                            '/root',
                            Extract(x_Editlog, '/root/*'))
             Into x_Editlog From 影像报告记录 Where ID = v_报告id;

       Update 影像报告记录
                Set 报告内容     = 报告内容_In,
                最后编辑时间 = Cur_Time,
                最后编辑人   = 最后编辑人_In,
                编辑日志     = x_Editlog,
                记录人       =记录人_In,
                诊断意见     =v_Result
                Where ID = v_报告id;
       end if;

  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      Zl_Errorcenter(SQLCode, SQLErrM);
  End p_影像报告记录_新增;

  --6、获取书写的文档内容
  Procedure p_Get_Doc_Content(
    Val           Out t_Refcur,
	DocID_In 影像报告记录.Id%Type
	) As
  Begin
    Open Val For
      Select  Nvl(a.报告内容.GetClobVal(), '<ZLXML/>') As 报告内容 From 影像报告记录 A Where a.Id = DocID_In;
  End;

  --7、设置报告打印作废信息
  Procedure p_Checkrejectsignature(Signdate_In Date,
                                   报告ID_In   影像报告操作记录.报告Id%Type,
                                   作废人_In   影像报告操作记录.作废人%Type,
                                   作废说明_In 影像报告操作记录.作废说明%Type,
                                   Val         Out Sys_Refcursor) As
  Begin
    Open Val For
      Select 操作人, 操作时间
        From 影像报告操作记录
       Where 报告ID = 报告ID_In
         And 操作类型=1
         And 操作时间 >= Signdate_In
         And 作废时间 Is Null
       Order By 操作时间 Asc;
    --作废打印记录
    Update 影像报告操作记录 B
       Set 作废人 = 作废人_In, 作废时间 = Sysdate, b.作废说明 = 作废说明_In
     Where 报告ID = 报告ID_In And 操作类型=1
       And 操作时间 >= Signdate_In;

  End p_Checkrejectsignature;

  --8、查询相应原型下的最大序号
  Procedure p_Get_Samplelist_Maxseqnum(
    Val           Out t_Refcur,
	原型ID_In 影像报告范文清单.原型ID%Type
	) As
  Begin
    Open Val For
      Select Nvl(Max(a.编号), 0) + 1 As Num
        From 影像报告范文清单 A
       Where a.原型ID = 原型ID_In;
  End;

  --9、删除文档范文
  Procedure p_Del_影像报告范文清单(
    Id_In 影像报告范文清单.Id%Type
	) As
  Begin
    Delete From 影像报告范文清单 Where Id = Id_In;
  End;
  
 --10、添加文档的操作日志
  Procedure p_影像报告操作记录_Add(Id_In       影像报告操作记录.Id%Type,
                               报告ID_In   影像报告操作记录.报告ID%Type,
                               操作人_In   影像报告操作记录.操作人%Type,
                               操作类型_In 影像报告操作记录.操作类型%Type) As
  n_医嘱ID 影像报告操作记录.医嘱ID%Type;
  n_文档标题 影像报告记录.文档标题%Type;
  Begin

  Begin
    Select 医嘱ID,文档标题 Into n_医嘱ID,n_文档标题 From 影像报告记录 Where ID = 报告ID_In;
  Exception
    When Others Then
      null;
  End;
  if n_医嘱ID is not null then
    Insert Into 影像报告操作记录
      (ID, 报告ID,医嘱ID,文档标题,操作人,操作时间,操作类型)
    Values
      (Id_In, 报告ID_In, n_医嘱ID,n_文档标题,操作人_In, sysdate,操作类型_In);
    if 操作类型_In=1 then
        update 影像报告记录 set 报告打印=1 where ID=报告ID_In;
    end if;
  end if;
  Exception
    When Others Then
      Zl_Errorcenter(SQLCode, SQLErrM);
  End;

  --11、删除报告
  Procedure p_影像报告记录_删除(
    报告_Id_In 影像报告记录.Id%Type
	) As
  Begin    

    Delete From 影像报告记录 Where 影像报告记录.Id = Hextoraw(报告_Id_In);

    Delete From 病人医嘱报告 Where 检查报告ID =hextoraw(报告_Id_In);

  Exception   
    When Others Then
      Zl_Errorcenter(SQLCode, SQLErrM);
  End p_影像报告记录_删除;


--12、获取签名类型
Procedure p_Get_SysConfigSignature(
  Val           Out t_Refcur,
  科室ID_In		In 部门表.ID%Type
  )Is
Begin
    --返回用户, 模块号,功能
	Open  Val For 
	    select Zl_Fun_Getsignpar(7, 科室ID_In) as 签名类型 from dual;
Exception
  When Others Then
  Zl_Errorcenter(Sqlcode, Sqlerrm); 
End;


--13、获取账户签名印章
Procedure p_Get_PersonSignImg(
  Val           Out t_Refcur,
  ID_In		In 人员表.ID%Type
  )Is
  v_sql Varchar2(1000);
  n_count Number(5);
Begin                
  Select Count(*) Into n_Count From user_tables Where table_name =Upper('影像签名图片');
  
  If n_Count > 0 Then
     v_sql := 'Truncate Table 影像签名图片';
     Execute Immediate v_sql;   
     
     v_sql := 'Insert Into 影像签名图片 Select a.id, to_lob(a.签名图片) as 签名图片 From 人员表 a Where a.ID=' || ID_In;
     Execute Immediate v_sql;  
  Else
     v_sql := 'Create GLOBAL TEMPORARY TABLE 影像签名图片 ON COMMIT PRESERVE ROWS AS Select a.id, to_lob(a.签名图片) as 签名图片 From 人员表 a Where a.ID=' || ID_In;  
     Execute Immediate v_sql;    
  End If; 
    
  v_sql := 'Select 签名图片 From 影像签名图片 Where Id=:ID';
    --返回用户, 模块号,功能
	Open  Val For v_sql Using ID_In;  

Exception
  When Others Then
  Zl_Errorcenter(Sqlcode, Sqlerrm); 
End;


--14、获取签名的证书信息
Procedure p_Get_SignCertInfo(
  Val           Out t_Refcur,
  证书ID_In		人员证书记录.ID%Type
  )Is
Begin
	Open  Val For 
	    Select ID, CertDN,CertSN,SignCert,EncCert From 人员证书记录 Where ID=证书ID_In;
Exception
  When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm); 
End;

--15、更新报告状态
Procedure p_Update_ReportState(
  报告Id_In  影像报告记录.ID%Type,
  报告状态_In  影像报告记录.报告状态%Type,
  审核人_In   影像报告记录.最后审核人%Type
  )Is
Begin
  --报告状态1-未签名；2-已诊断；3-已审核；4-已终审；5-诊断驳回；6-审核驳回
  --如果报告状态是1-未签名；2-已诊断;5-诊断驳回，此时是没有审核人的
  if (报告状态_In=1) or (报告状态_In=2) or (报告状态_In=5) then 
    Update 影像报告记录 Set 报告状态=报告状态_In,最后审核人=null,最后审核时间=null Where ID=报告Id_In;
  elsif (报告状态_In=3) or (报告状态_In=4) then 
    Update 影像报告记录 Set 报告状态=报告状态_In,最后审核人=审核人_In,最后审核时间=sysdate Where ID=报告Id_In;
  else
    Update 影像报告记录 Set 报告状态=报告状态_In Where ID=报告Id_In;
  end if;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End;

--16、获取报告状态
Procedure p_Get_ReportState(
  Val           Out t_Refcur,
  报告Id_In	影像报告记录.ID%Type
  )Is
Begin
	Open  Val For 
	    Select 报告状态 From 影像报告记录 Where ID=报告Id_In;
Exception
  When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm); 
End;



--17、报告驳回
Procedure p_Reject_Report(
  医嘱ID_In  影像报告驳回.医嘱ID%Type,
  报告ID_In  影像报告驳回.检查报告ID%Type,
  驳回理由_In 影像报告驳回.驳回理由%Type,
  驳回时间_In 影像报告驳回.驳回时间%Type,
  驳回人_In   影像报告驳回.驳回人%Type,
  待处理人_In  影像报告记录.待处理人%Type,
  报告状态_In 影像报告记录.报告状态%Type
  )Is
Begin
  Insert Into 影像报告驳回(ID, 医嘱ID,检查报告ID,驳回理由,驳回时间,驳回人)
  Values(影像报告驳回_ID.NEXTVAL, 医嘱ID_IN, 报告ID_In, 驳回理由_IN, 驳回时间_IN, 驳回人_IN);

  Update 影像报告记录 Set 报告状态=报告状态_In,待处理人=待处理人_In Where ID=报告ID_In;

  --Update 病人医嘱发送 Set 执行过程=-1 Where 医嘱ID= 医嘱ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End;

--17.1、撤销报告驳回
Procedure p_Reject_Cancel(
  ID_In       影像报告驳回.ID%Type,
  报告ID_In    影像报告驳回.检查报告ID%Type,
  报告状态_In   影像报告记录.报告状态%Type
  )Is
Begin
  Update 影像报告驳回 Set 是否撤销=1 Where ID=ID_In;
  Update 影像报告记录 Set 报告状态=报告状态_In,待处理人='' Where ID=报告ID_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End;

--18、获取报告驳回信息
Procedure p_Get_RejectInfo(
  Val           Out t_Refcur,
  报告ID_In  影像报告驳回.检查报告ID%Type
  )Is
Begin
  Open  Val For
    Select A.ID, A.驳回理由, A.驳回时间, A.驳回人, Nvl( A.是否撤销,0) As 驳回状态, B.报告状态
    From 影像报告驳回 A, 影像报告记录 B Where A.检查报告ID=报告Id_In And A.检查报告ID = B.ID Order by 驳回时间;
End;

--19、获取原型动作
Procedure p_Get_Doc_Process(
  Val           Out t_Refcur,
  原型ID_In 影像报告动作.原型id%Type
  ) As
  Begin
    Open Val For
      Select RawtoHex(p.id) ID,
             p.名称 As 动作名称,
			 e.名称 As 事件名称,
			 e.种类 As 事件种类,
			 e.元素IID As 元素IID,
             p.动作类型,
             p.序号,
             p.说明,
             p.可否手工执行,
             To_Clob(Nvl(p.内容.GetClobVal(),'<NULL/>')) As 内容, 
             RawtoHex(p.事件ID) 事件ID
        From 影像报告动作 P, 影像报告事件 E
       Where p.事件ID = e.Id(+) And p.原型ID=原型ID_In
       Order By 动作类型, 序号;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Process;

  --20、通过学科筛选获得相应的范文信息
  Procedure p_Get_Samplelist_By_Conditions(
    Val           Out t_Refcur,
    原型id_In       Varchar2,
    学科_In          Varchar2,
    Condition_In Varchar2, --过滤筛选
    作者_In          Varchar2
  ) As
  Begin

    Open Val For
      Select /*+ rule*/ Rawtohex(a.Id) ID, a.名称, a.作者, a.说明,
             Nvl2(a.说明, a.说明 || '作者:' || a.作者, '作者:' || a.作者) Content, a.标签, a.学科
      From 影像报告范文清单 A
      Where a.原型ID = Hextoraw(原型id_In) And
            ((a.学科 Is Null And a.是否私有 = 0) Or 学科_In Is Null Or a.作者 = 作者_In Or
            (a.学科 Is Not Null And  b_PACS_RptPublic.f_If_Intersect(a.学科, 学科_In) > 0 And a.是否私有 = 0)) And
            (Condition_In Is Null Or
            (a.标签 Is Not Null And Condition_In Is Not Null And b_PACS_RptPublic.f_If_Intersect(a.标签, Condition_In) > 0))
      Order By a.编号;

  End p_Get_Samplelist_By_Conditions;

  --21、通过部门ID获取部门名称
  Procedure p_Get_部门名称_By_ID(
    Val           Out t_Refcur,
    ID_IN 部门表.ID%TYPE
    )Is
  begin
       open val for
       select 名称 from 部门表 where id=ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_部门名称_By_ID;
    

 --22、提取所有预备提纲
  Procedure p_Get_AllPreOutlines(
    Val           Out t_Refcur
  )Is
  begin
       open val for
       Select Rawtohex(ID) ID, a.编码, a.名称 From 影像报告预备提纲 a Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_AllPreOutlines;

  --23、提取文档标题
  Procedure p_Get_reportTitle_By_ID(
    Val           Out t_Refcur,
	ID_IN  影像报告记录.id%TYPE   
    )Is
  begin
       open val for
       select 文档标题 from 影像报告记录 where id=ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_reportTitle_By_ID;

  --24、提取报告锁定人
  Procedure p_Get_报告锁定人_By_ID(
    Val           Out t_Refcur,
	ID_IN  影像报告记录.id%TYPE   
    )Is
  Begin
       Open Val For
         Select 锁定人 From 影像报告记录 Where id =ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_报告锁定人_By_ID;

 --25、通过医嘱ID获取报告列表
  Procedure p_Get_影像报告记录_By_医嘱ID(
    Val           Out t_Refcur,
	医嘱ID_IN  影像报告记录.医嘱ID%TYPE
    )Is
  Begin
       Open Val For
       Select RawToHex(ID) As REPORTID, RawToHex(原型ID) As ANTETYPEID, 医嘱ID As ORDERID,文档标题 As REPORTNAME,
              创建时间 As REPORTDATE, Decode(Nvl(报告状态,0),1,'编辑中',2,'已诊断',3,'已审核',4,'已终审',5,'诊断驳回','审核驳回') As REPORTSTATE,
              创建人 As CreateUser,最后审核人 As ExamineyUser,Decode(Nvl(结果阳性,0),1,'阳性','') As RESULTPOSITIVE,
              Nvl(报告质量,0) As INNERQUALITY,' ' As REPORTQUALITY, Decode(Nvl(报告打印,0),0,'未打印','已打印') As ReportPrint,
              Decode(Nvl(报告发放,0),0,'未发放','已发放') As REPORTRELEASE ,记录人 as RECDOCTOR From 影像报告记录 Where 医嘱ID =医嘱ID_IN
              order by REPORTDATE desc;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_影像报告记录_By_医嘱ID;
  
  --26、查询影像流程参数值
  Procedure p_Get_影像流程参数值(
    Val           Out t_Refcur,
	科室ID_IN  影像流程参数.科室ID%TYPE
    )Is
  Begin
       Open val For
       Select 参数名,参数值 From 影像流程参数 Where 科室ID=科室ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_影像流程参数值;

  --27、根据医嘱ID，查询对应的原型列表
  Procedure p_Get_影像原型列表_By_医嘱ID(
    Val           Out t_Refcur,
    医嘱_IN  影像检查记录.医嘱ID%TYPE   
    )Is
  Begin
       Open Val For
       Select rawtohex(c.id) As ANTETYPEID , c.名称 As ANTETYPENAME,c.说明 
       From 病人医嘱记录 a,影像报告原型应用 b,影像报告原型清单 c 
       Where a.id=医嘱_IN And a.诊疗项目id=b.诊疗项目ID And b.报告原型ID=c.id And a.病人来源 =b.应用场合;
       
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_影像原型列表_By_医嘱ID;

  --28、根据报告ID查询打印记录
  procedure p_Get_ReportPrintLog_By_报告ID
  (
       val out sys_refcursor  ,
       报告_IN  影像报告操作记录.报告ID%TYPE
  )is
  begin
       open val for
       Select  c.文档标题 , b.操作人, To_Char(b.操作时间, 'yyyy-MM-dd HH24:mi') 打印时间, b.作废人,
               To_Char(b.作废时间, 'yyyy-MM-dd HH24:mi') 作废时间, b.作废说明
               From 影像报告操作记录 B, 影像报告记录 C
               Where c.Id = 报告_IN And b.报告ID = c.Id And 操作类型=1 Order By b.操作时间;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_ReportPrintLog_By_报告ID;

  --29、根据医嘱ID查询报告发放列表
  Procedure p_Get_ReportReleaseList(
    Val           Out t_Refcur,
    医嘱_IN  影像报告记录.医嘱ID%TYPE   
    )Is
  Begin
       Open val For
       Select rawtohex(ID) As 报告ID, 文档标题 As 报告名称,最后编辑时间 as 报告日期,
              decode(nvl(报告发放,0),0,'未发放','已发放') As 报告发放 
              From 影像报告记录 Where 报告状态 Between 2 And 4 And 医嘱ID =医嘱_IN;
       
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ReportReleaseList;

  --30、根据报告ID查询驳回记录数量
  Procedure p_Get_RejectedCount(
    Val           Out t_Refcur,
    报告_IN  影像报告驳回.检查报告ID%TYPE
    )Is
  Begin
       Open val For
       Select count(*) As 驳回数量 From 影像报告驳回 Where 检查报告ID=报告_IN;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_RejectedCount;

   --31、根据医嘱ID查询报告动作需要的一些ID们
  Procedure p_Get_DocProcess_IDs(
    Val           Out t_Refcur,
    医嘱_IN  病人医嘱记录.ID%TYPE
    )Is
  Begin
       open val for
       select ID as 医嘱ID,主页ID,挂号单 from 病人医嘱记录 where ID=医嘱_IN;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_DocProcess_IDs;

  --32、根据医嘱ID和报告ID查询报告的一些参数
  Procedure p_Get_DocInfo(
       Val           Out t_Refcur,
       医嘱ID_IN  影像检查记录.医嘱ID%TYPE,
       报告ID_IN  影像报告记录.ID%TYPE
  )Is
  Begin
      If 报告ID_IN Is Null Then 
        Open Val For 
        Select 执行科室ID,'创建人' As 创建人 From 影像检查记录 Where 医嘱ID=医嘱ID_IN;
      Else
        Open Val For
        Select 执行科室ID,创建人 From 影像检查记录 A,影像报告记录 b Where a.医嘱ID=B.医嘱ID and b.id=报告ID_IN;
      End if;
       

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_DocInfo;

  --33、查询一个检查中相同原型ID的报告数量
   Procedure p_Get_SameAntetypeDocCounts(
       Val           Out t_Refcur,
       医嘱ID_IN  影像报告记录.医嘱ID%TYPE,
       原型ID_IN  影像报告记录.原型ID%TYPE
  )Is
  Begin      
        Open Val For
        Select count(id) as DocCounts From 影像报告记录 Where 医嘱ID=医嘱ID_IN and 原型ID=原型ID_IN;    
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_SameAntetypeDocCounts;

  --34、提取报告图存储信息
  Procedure p_Get_DocImageSaveInof_By_ID(
    Val           Out t_Refcur,
	  ID_IN  影像报告记录.id%TYPE
    )Is
  begin
       open val for
       select 设备号,创建时间 from 影像报告记录 where id=ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_DocImageSaveInof_By_ID;

End b_PACS_RptManage;
/



--88036:李小东,2015-11-02,新增药品特性属性字段(是否原研药、是否专利药、是否单独定价)
Create Or Replace Procedure Zl_成药品种_Insert
(
  类别_In       In 诊疗项目目录.类别%Type := Null,
  分类id_In     In 诊疗项目目录.分类id%Type := Null,
  Id_In         In 诊疗项目目录.Id%Type,
  编码_In       In 诊疗项目目录.编码%Type := Null,
  名称_In       In 诊疗项目目录.名称%Type := Null,
  拼音_In       In 诊疗项目别名.简码%Type := Null,
  五笔_In       In 诊疗项目别名.简码%Type := Null,
  英文_In       In 诊疗项目别名.名称%Type := Null,
  单位_In       In 诊疗项目目录.计算单位%Type := Null,
  药品剂型_In   In 药品特性.药品剂型%Type := Null,
  毒理分类_In   In 药品特性.毒理分类%Type := Null,
  价值分类_In   In 药品特性.价值分类%Type := Null,
  货源情况_In   In 药品特性.货源情况%Type := Null,
  用药梯次_In   In 药品特性.用药梯次%Type := Null,
  药品类型_In   In 药品特性.药品类型%Type := Null,
  处方职务_In   In 药品特性.处方职务%Type := '00',
  处方限量_In   In 药品特性.处方限量%Type := Null,
  急救药否_In   In 药品特性.急救药否%Type := 0,
  是否新药_In   In 药品特性.是否新药%Type := 0,
  是否原料_In   In 药品特性.是否原料%Type := 0,
  是否皮试_In   In 药品特性.是否皮试%Type := 0,
  抗生素_In     In 药品特性.抗生素%Type := 0,
  参考目录id_In In 诊疗项目目录.参考目录id%Type := Null,
  品种医嘱_In   In 药品特性.品种医嘱%Type := 0,
  适用性别_In   In 诊疗项目目录.适用性别%Type := 0,
  其他别名_In   In Varchar2 := Null, --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织 
  自管药_In     In Number := Null,
  Atccode_In    In Varchar2 := Null,
  肿瘤药_In     In 药品特性.是否肿瘤药%Type := 0,
  溶媒_In       In 药品特性.溶媒%Type := 0,
  是否原研药_In In 药品特性.是否原研药%Type := 0,
  是否专利药_In In 药品特性.是否专利药%Type := 0,
  是否单独定价_In   In 药品特性.是否单独定价%Type := 0
) Is
  v_Records Varchar2(4000); --临时记录别名数据的字符串 
  v_Currrec Varchar2(1000); --包含在别名记录中的一条别名 
  v_Fields  Varchar2(1000); --临时记录一条别名的字符串 
  v_名称    诊疗项目目录.名称%Type;
  v_拼音    诊疗项目别名.简码%Type;
  v_五笔    诊疗项目别名.简码%Type;
Begin
  Insert Into 诊疗项目目录
    (类别, 分类id, ID, 编码, 名称, 计算单位, 计算方式, 执行频率, 单独应用, 组合项目, 执行安排, 计价性质, 服务对象, 建档时间, 撤档时间, 参考目录id, 适用性别)
  Values
    (类别_In, 分类id_In, Id_In, 编码_In, 名称_In, 单位_In, 1, 0, 1, 0, 0, 0, 3, Sysdate, To_Date('3000-01-01', 'YYYY-MM-DD'),
     参考目录id_In, 适用性别_In);

  Insert Into 药品特性
    (药名id, 药品剂型, 毒理分类, 价值分类, 货源情况, 用药梯次, 药品类型, 处方职务, 处方限量, 急救药否, 是否新药, 抗生素, 是否原料, 是否皮试, 品种医嘱, 临床自管药, Atccode, 是否肿瘤药, 溶媒,
     是否原研药, 是否专利药, 是否单独定价)
  Values
    (Id_In, 药品剂型_In, 毒理分类_In, 价值分类_In, 货源情况_In, 用药梯次_In, 药品类型_In, 处方职务_In, 处方限量_In, 急救药否_In, 是否新药_In, 抗生素_In, 是否原料_In,
     是否皮试_In, 品种医嘱_In, 自管药_In, Atccode_In, 肿瘤药_In, 溶媒_In, 是否原研药_In, 是否专利药_In, 是否单独定价_In);

  If 拼音_In Is Not Null Then
    Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 名称_In, 1, 拼音_In, 1);
  End If;
  If 五笔_In Is Not Null Then
    Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 名称_In, 1, 五笔_In, 2);
  End If;
  If 英文_In Is Not Null Then
    Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 英文_In, 2, Null, 0);
  End If;

  If 其他别名_In Is Null Then
    v_Records := Null;
  Else
    v_Records := 其他别名_In || '|';
  End If;
  While v_Records Is Not Null Loop
    v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields  := v_Currrec;
    v_名称    := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
    v_Fields  := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_拼音    := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
    v_Fields  := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_五笔    := v_Fields;
    If v_拼音 Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, v_名称, 9, v_拼音, 1);
    End If;
    If v_五笔 Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, v_名称, 9, v_五笔, 2);
    End If;
    v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
  End Loop;
  --添加缺省的对应输出单据 
  Insert Into 病历单据应用
    (病历文件id, 应用场合, 诊疗项目id)
    Select a.病历文件id, 1, Id_In
    From 病历单据应用 A, 诊疗项目目录 I
    Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 应用场合 = 1 And Rownum < 2;
  Insert Into 病历单据应用
    (病历文件id, 应用场合, 诊疗项目id)
    Select a.病历文件id, 2, Id_In
    From 病历单据应用 A, 诊疗项目目录 I
    Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 应用场合 = 2 And Rownum < 2;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_成药品种_Insert;
/

--88036:李小东,2015-11-02,新增药品特性属性字段(是否原研药、是否专利药、是否单独定价)
Create Or Replace Procedure Zl_成药品种_Update
(
  分类id_In     In 诊疗项目目录.分类id%Type := Null,
  Id_In         In 诊疗项目目录.Id%Type,
  编码_In       In 诊疗项目目录.编码%Type := Null,
  名称_In       In 诊疗项目目录.名称%Type := Null,
  拼音_In       In 诊疗项目别名.简码%Type := Null,
  五笔_In       In 诊疗项目别名.简码%Type := Null,
  英文_In       In 诊疗项目别名.名称%Type := Null,
  单位_In       In 诊疗项目目录.计算单位%Type := Null,
  药品剂型_In   In 药品特性.药品剂型%Type := Null,
  毒理分类_In   In 药品特性.毒理分类%Type := Null,
  价值分类_In   In 药品特性.价值分类%Type := Null,
  货源情况_In   In 药品特性.货源情况%Type := Null,
  用药梯次_In   In 药品特性.用药梯次%Type := Null,
  药品类型_In   In 药品特性.药品类型%Type := Null,
  处方职务_In   In 药品特性.处方职务%Type := '00',
  处方限量_In   In 药品特性.处方限量%Type := Null,
  急救药否_In   In 药品特性.急救药否%Type := 0,
  是否新药_In   In 药品特性.是否新药%Type := 0,
  是否原料_In   In 药品特性.是否原料%Type := 0,
  是否皮试_In   In 药品特性.是否皮试%Type := 0,
  抗生素_In     In 药品特性.抗生素%Type := 0,
  参考目录id_In In 诊疗项目目录.参考目录id%Type := Null,
  品种医嘱_In   In 药品特性.品种医嘱%Type := 0,
  适用性别_In   In 诊疗项目目录.适用性别%Type := 0,
  其他别名_In   In Varchar2 := Null, --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织 
  自管药_In     In Number := Null,
  Atccode_In    In Varchar2 := Null,
  肿瘤药_In     In 药品特性.是否肿瘤药%Type := 0,
  溶媒_In       In 药品特性.溶媒%Type := 0,
  是否原研药_In In 药品特性.是否原研药%Type := 0,
  是否专利药_In In 药品特性.是否专利药%Type := 0,
  是否单独定价_In   In 药品特性.是否单独定价%Type := 0
) Is
  v_Records Varchar2(4000); --临时记录别名数据的字符串 
  v_Currrec Varchar2(1000); --包含在别名记录中的一条别名 
  v_Fields  Varchar2(1000); --临时记录一条别名的字符串 
  v_名称    诊疗项目目录.名称%Type;
  v_拼音    诊疗项目别名.简码%Type;
  v_五笔    诊疗项目别名.简码%Type;
  Err_Notfind Exception;
Begin
  Update 诊疗项目目录
  Set 分类id = 分类id_In, 编码 = 编码_In, 名称 = 名称_In, 计算单位 = 单位_In, 参考目录id = 参考目录id_In, 适用性别 = 适用性别_In
  Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;

  Update 药品特性
  Set 药品剂型 = 药品剂型_In, 毒理分类 = 毒理分类_In, 价值分类 = 价值分类_In, 货源情况 = 货源情况_In, 用药梯次 = 用药梯次_In, 药品类型 = 药品类型_In, 处方职务 = 处方职务_In,
      处方限量 = 处方限量_In, 抗生素 = 抗生素_In, 急救药否 = 急救药否_In, 是否新药 = 是否新药_In, 是否原料 = 是否原料_In, 是否皮试 = 是否皮试_In, 品种医嘱 = 品种医嘱_In,
      临床自管药 = 自管药_In, Atccode = Atccode_In, 是否肿瘤药 = 肿瘤药_In, 溶媒 = 溶媒_In, 是否原研药 = 是否原研药_In, 是否专利药 = 是否专利药_In,
      是否单独定价 = 是否单独定价_In
  Where 药名id = Id_In;

  Update 收费项目目录 Set 名称 = 名称_In Where ID In (Select 药品id From 药品规格 Where 药名id = Id_In);

  If 拼音_In Is Null Then
    Delete From 诊疗项目别名 Where 诊疗项目id = Id_In And 性质 = 1 And 码类 = 1;
    Delete From 收费项目别名
    Where 收费细目id In (Select 药品id From 药品规格 Where 药名id = Id_In) And 性质 = 1 And 码类 = 1;
  Else
    Update 诊疗项目别名 Set 名称 = 名称_In, 简码 = 拼音_In Where 诊疗项目id = Id_In And 性质 = 1 And 码类 = 1;
    If Sql%RowCount = 0 Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 名称_In, 1, 拼音_In, 1);
    End If;
    For r_Spec In (Select 药品id From 药品规格 Where 药名id = Id_In) Loop
      Update 收费项目别名
      Set 名称 = 名称_In, 简码 = 拼音_In
      Where 收费细目id = r_Spec.药品id And 性质 = 1 And 码类 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (r_Spec.药品id, 名称_In, 1, 拼音_In, 1);
      End If;
    End Loop;
  End If;
  If 五笔_In Is Null Then
    Delete From 诊疗项目别名 Where 诊疗项目id = Id_In And 性质 = 1 And 码类 = 2;
    Delete From 收费项目别名
    Where 收费细目id In (Select 药品id From 药品规格 Where 药名id = Id_In) And 性质 = 1 And 码类 = 2;
  Else
    Update 诊疗项目别名 Set 名称 = 名称_In, 简码 = 五笔_In Where 诊疗项目id = Id_In And 性质 = 1 And 码类 = 2;
    If Sql%RowCount = 0 Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 名称_In, 1, 五笔_In, 2);
    End If;
    For r_Spec In (Select 药品id From 药品规格 Where 药名id = Id_In) Loop
      Update 收费项目别名
      Set 名称 = 名称_In, 简码 = 五笔_In
      Where 收费细目id = r_Spec.药品id And 性质 = 1 And 码类 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (r_Spec.药品id, 名称_In, 1, 五笔_In, 2);
      End If;
    End Loop;
  End If;
  If 英文_In Is Null Then
    Delete From 诊疗项目别名 Where 诊疗项目id = Id_In And 性质 = 2;
    Delete From 收费项目别名 Where 收费细目id In (Select 药品id From 药品规格 Where 药名id = Id_In) And 性质 = 2;
  Else
    Update 诊疗项目别名 Set 名称 = 英文_In Where 诊疗项目id = Id_In And 性质 = 2;
    If Sql%RowCount = 0 Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 英文_In, 2, Null, 0);
    End If;
    For r_Spec In (Select 药品id From 药品规格 Where 药名id = Id_In) Loop
      Update 收费项目别名 Set 名称 = 英文_In Where 收费细目id = r_Spec.药品id And 性质 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (r_Spec.药品id, 英文_In, 2, Null, 0);
      End If;
    End Loop;
  End If;

  Delete From 诊疗项目别名 Where 诊疗项目id = Id_In And 性质 = 9;
  Delete From 收费项目别名 Where 收费细目id In (Select 药品id From 药品规格 Where 药名id = Id_In) And 性质 = 9;
  If 其他别名_In Is Null Then
    v_Records := Null;
  Else
    v_Records := 其他别名_In || '|';
  End If;
  While v_Records Is Not Null Loop
    v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields  := v_Currrec;
    v_名称    := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
    v_Fields  := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_拼音    := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
    v_Fields  := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_五笔    := v_Fields;
    If v_拼音 Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, v_名称, 9, v_拼音, 1);
      Insert Into 收费项目别名
        (收费细目id, 名称, 性质, 简码, 码类)
        Select 药品id, v_名称, 9, v_拼音, 1 From 药品规格 Where 药名id = Id_In;
    End If;
    If v_五笔 Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, v_名称, 9, v_五笔, 2);
      Insert Into 收费项目别名
        (收费细目id, 名称, 性质, 简码, 码类)
        Select 药品id, v_名称, 9, v_五笔, 2 From 药品规格 Where 药名id = Id_In;
    End If;
    v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
  End Loop;

Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该品种不存在，可能已被其他用户删除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_成药品种_Update;
/

--89077:涂建华,2015-11-11,报告片段加载调整
--89077:许华锋,2015-11-02,报告片段适应条件
CREATE OR REPLACE Package b_PACS_RptCommon Is
  Type t_Refcur Is Ref Cursor;

  --获取预备提纲>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	);

  --元素分类>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Eleclass(
    Val Out t_Refcur
	);

  --原型片段>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetype_Fragments(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	);

  --原型列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_By_Id(
    Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	);

  --原型内容>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_Content(
    Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	);

  --范文清单>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Samplelist_By_Aid(
    Val           Out t_Refcur,
	Antetypelist_Id_In Varchar2,
	Condition_In       影像报告范文清单.名称%Type,
	Author_In          影像报告范文清单.作者%Type,
	Subjects_In        影像报告范文清单.学科%Type
	);

  --获取插件配置根据插件ID获取>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigById(
    Val           Out t_Refcur,
	Id_In 影像报告插件.ID%Type
	);

  --获取插件配置根据原型清单获取>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigByAId(
    Val           Out t_Refcur,
	Aid_In 影像报告原型清单.ID%Type
	);

  --获取元素>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Element(
    Val Out t_Refcur
	);

  --获取片段列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	);

  --获取值域列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Range_List(
    Val Out t_Refcur
	);

  --获取值域列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Combo_List(
    Val Out t_Refcur
	);

  --获取原型片段根据原型ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_FragmentDirectory_ByAid(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	);

  --根据原型id获取片段数据
  Procedure p_Get_FragmentData_ByAid(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	);

  --获取数据表的最后更新时间>>>>>>>>>>>>>>>>>>>>>>>>>>>
  procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
	Table_Name_In Varchar2
	);

  --获取片段列表根据上级ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Bypid(
    Val           Out t_Refcur,
	Pid_In 影像报告片段清单.上级ID%Type
	);

  --获取片段列表根据节点类型>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Byleaf(
    Val           Out t_Refcur,
	Leaf_In 影像报告片段清单.节点类型%Type
	);

  --获取值域信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Range_List_Byid(
    Val           Out t_Refcur,
	Id_In 影像报告值域清单.Id%Type
	);

  --根据元素ID获取值域ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Getelementrid_By_Eid(
    Val           Out t_Refcur,
	Eid_In 影像报告元素清单.Id%Type
	);

  --获取计量单位列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_GetMasure_UnitList(
    Val Out t_Refcur
	);

  --获取文档种类信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Doc_Kind(
    Val Out t_Refcur
	);

  --功能：获取所有学科信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Subjects(
    Val Out t_Refcur
	);

  --查看是否存在相应的编码或者名称(用于导入导出)>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exits_Doc_Kinds(
    Val           Out t_Refcur,
	编码_In      Varchar2,
	名称_In      Varchar2,
	Tablename_In Varchar2
	);

  --是否存在相同的ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exist_Id(
    Val           Out t_Refcur,
	Id_In        Number,
	Tablename_In Varchar2
	);

  --通过名称获取ID信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Id_By_Title(
    Val           Out t_Refcur,
	名称_In      Varchar2,
	Tablename_In Varchar2,
	Type_In      Varchar2
	);

  --通过简称片段清单
  procedure p_Get_FragmentSampleName(
    Val           Out t_Refcur,
	简称_In Varchar2
	);

  --更新ID对应的片段内容
  procedure p_Update_PhraseContent(
    Id_In      影像报告片段清单.ID%type,
	Name_In		影像报告片段清单.名称%Type,
	Content_In Varchar2
	);
  --获取原型ID对应的第一层片段节点
  procedure p_Get_FragmentData_LevelOne(
    Val           Out t_Refcur,
	AId_In 影像报告原型清单.ID%type
	);

  -- 获取片段的下层节点
  procedure p_GetFragmentDataListByFID(
    Val           Out t_Refcur,
	FId_In 影像报告片段清单.ID%type
	);
end b_PACS_RptCommon;
/

CREATE OR REPLACE Package Body b_PACS_RptCommon Is
  -- 功    能：该方法只用于演示...

  --获取预备提纲>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) As ID, a.编码, a.名称, a.说明
        From 影像报告预备提纲 A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --元素分类>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Eleclass(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(A.ID) As ID,
             A.编码,
             A.名称,
             A.说明,
             RawToHex(A.上级ID) 上级ID
        From 影像报告元素分类 A
       Order By 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --原型片段>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetype_Fragments(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(A.片段ID) As 片段ID
        From 影像报告原型片段 A
       Where a.原型ID = Aid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --原型清单>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_By_Id(
    Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	) As
  Begin
    Open Val For
      Select /*+rule*/
       RawToHex(A.ID) As ID,
       a.种类,
       a.编码,
       a.名称,
       a.说明,
       a.可否重置页面 As 页面重置,
       a.可否重置格式 As 格式重置,
       Extractvalue(b.Column_Value, '/root/print_hf_mode') Printhfmode,
       Extractvalue(b.Column_Value, '/root/print_follow_pages') Printfollowpages,
       Nvl(Extractvalue(b.Column_Value, '/root/print_limit'), 0) Printlimit,
       Nvl(a.控制选项, '<NULL/>') As 控制选项,
       a.创建人,
       a.创建时间,
       a.修改人,
       a.修改时间,
       a.是否禁用,
       A.分组
        From 影像报告原型清单 A,
             Table(Xmlsequence(Extract(a.控制选项, '/root'))) B
       Where a.Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --原型内容>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_Content(
    Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	) As
  Begin
    Open Val For
      Select Nvl(a.内容, '<ZLXML/>') As 内容
        From 影像报告原型清单 A
       Where a.Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --通过原型ID获得相应的范文信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Samplelist_By_Aid(
    Val           Out t_Refcur,
	Antetypelist_Id_In Varchar2,
	Condition_In       影像报告范文清单.名称%Type,
	Author_In          影像报告范文清单.作者%Type,
	Subjects_In        影像报告范文清单.学科%Type
	) As
  Begin
    --直接获取该原型下的范文列表
    If Length(Antetypelist_Id_In) > 30 Then
      Open Val For
        Select /*+rule*/
         RawToHex(A.ID) as ID,
         a.名称,
         a.作者,
         a.说明,
         a.学科,
         a.编号,
         a.标签,
         a.是否私有
          From 影像报告范文清单 A
         Where a.原型ID = Hextoraw(Antetypelist_Id_In)
           And (a.作者 = Author_In Or (a.学科 Is Null And a.是否私有 = 0) Or
               Subjects_In Is Null Or
               (a.学科 Is Not Null And
               b_PACS_RptPublic.f_If_Intersect(a.学科, Subjects_In) > 0 And
               a.是否私有 = 0));
    Else
      --获得一个存在原型信息的范文树形结构
      Open Val For
        Select Distinct a.分组 As ID,
                        a.分组 As 名称,
                        '' as 说明,
                        '' As 原型ID,
                        'category' As 类型,
                        '' As 作者,
                        '' As 学科,
                        Null as 最后编辑时间,
                        '' As 标签,
                        0 As 是否私有,
                        0 As Imgindex
          From 影像报告原型清单 A
         Where a.种类 = Antetypelist_Id_In
           And Exists
         (Select ID From 影像报告范文清单 C Where c.原型ID = a.Id)
           And a.分组 Is Not Null
        Union
        Select m.*
          From (Select RawToHex(B.ID) As ID,
                       b.名称,
                       b.说明,
                       b.分组 As 原型ID,
                       'antetype' As 类型,
                       '' As 作者,
                       '' As 学科,
                       Null as 最后编辑时间,
                       '' As 标签,
                       0 As 是否私有,
                       0 As Imgindex
                  From 影像报告原型清单 B
                 Where b.种类 = Antetypelist_Id_In
                   And Exists (Select ID
                          From 影像报告范文清单 C
                         Where c.原型ID = b.Id)
                 Order By b.编码) M
        
        Union All
        Select n.*
          From (Select /*+rule*/
                 RawToHex(A.ID) As ID,
                 a.名称,
                 a.说明,
                 RawToHex(A.原型ID) As 原型ID,
                 'sample' As 类型,
                 a.作者,
                 a.学科,
                 a.最后编辑时间,
                 a.标签,
                 a.是否私有,
                 Decode(a.是否私有, 1, 2, 1) As Imgindex
                  From 影像报告范文清单 A, 影像报告原型清单 C
                 Where a.原型ID = c.Id
                   And c.种类 = Antetypelist_Id_In
                   And ((a.名称 Like '%' || Condition_In || '%' And
                       Condition_In Is Not Null) Or Condition_In Is Null)
                   And (a.作者 = Author_In Or (a.学科 Is Null And a.是否私有 = 0) Or
                       Subjects_In Is Null Or
                       (a.学科 Is Not Null And
                       b_PACS_RptPublic.f_If_Intersect(a.学科, Subjects_In) > 0 And
                       a.是否私有 = 0))
                 Order By a.编号, a.名称) N;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取插件配置根据插件ID获取>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigById(
    Val           Out t_Refcur,
	Id_In 影像报告插件.ID%Type
	) As
    v_Sql Varchar2(1000);
  Begin
    If (Id_In Is Not Null) Then
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式  From 影像报告插件 T Where t.Id =:Id_In And Rownum = 1';
      
        Open Val For v_Sql
          Using Id_In;
      End;
    Else
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式 From 影像报告插件 T where 是否禁用 = 0 order by t.编码';
      
        Open Val For v_Sql;
      End;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取插件配置根据原型清单获取>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigByAId(
    Val           Out t_Refcur,
	Aid_In 影像报告原型清单.ID%Type
	) As
    v_Sql Varchar2(1000);
  Begin
    If (Aid_In Is Not Null) Then
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式 ' ||
                 ' From 影像报告插件 T ' || ' Where t.Id in ( ' ||
                 ' Select X.pluginid from 影像报告原型清单 K, ' ||
                 '  (XMLTable(''*//pluginid''  Passing K.专用插件 Columns pluginid varchar2(32) Path ''/pluginid''))  X ' ||
                 ' Where K.id=:Aid_In) And 是否禁用 = 0' || ' Union All ' ||
                 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式 ' ||
                 ' From 影像报告插件 T ' || ' Where 是否禁用 = 0 And t.种类=0 ';
      
        Open Val For v_Sql
          Using Aid_In;
      End;
    Else
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式 From 影像报告插件 T where 是否禁用 = 0 order by t.编码';
      
        Open Val For v_Sql;
      End;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取所有元素>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Element(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) as ID,
             RawToHex(T.分类ID) as 分类ID,
             T.编码,
             T.名称,
             T.前缀,
             T.后缀,
             T.说明,
             T.数据类型,
             T.数值形态,
             T.最小长度,
             T.最大长度,
             T.最小小数位,
             T.最大小数位,
             T.计量单位,
             Nvl(T.扩展描述, '<NULL/>') As 扩展描述,
             T.值域ID,
             T.值域种类
        From 影像报告元素清单 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取片段列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.上级ID) As 上级ID,
             t.编码,
             t.名称,
             t.说明,
             t.节点类型,
             Nvl(t.组成, '<NULL/>') As 组成,
             t.学科,
             t.标签,
             t.是否私有,
             t.作者,
             t.最后编辑时间,
             Nvl(t.适应条件.GetClobVal(), '<NULL/>') As 适应条件
        From 影像报告片段清单 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取值域列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Range_List(
    Val Out t_Refcur
	) as
  begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.分类ID) As 分类ID,
             T.编码,
             T.名称,
             T.说明,
             T.数据类型,
             T.值域种类,
             Nvl(t.值域描述, '<NULL/>') As 值域描述,
             t.最后编辑时间
        From 影像报告值域清单 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end;

  --获取组句列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Combo_List(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             T.编码,
             T.名称,
             T.说明,
             T.多组,
             Nvl(t.组成, '<NULL/>') As 组成,
             T.编辑人,
             T.最后编辑时间,
             T.分组
        From 影像报告组句清单 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取原型片段目录根据原型ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_FragmentDirectory_ByAid(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select 名称,
             RawToHex(ID) As ID,
             4 as ImageIndex,
             RawToHex(上级ID) As 上级ID,
             '<NULL/>' As 组成,
             编码,
             节点类型,
             是否私有,
             作者,
             标签,
			 说明,
             学科
        From 影像报告片段清单
       Where ID In (Select ID
                      From 影像报告片段清单
                     Start With ID In (Select 片段ID
                                         From 影像报告原型片段
                                        Where 原型ID = Aid_In)
                    Connect By Prior 上级ID = ID)
       order by 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取原型片段数据根据原型ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_FragmentData_ByAid(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      With TabFragmentId As
       (Select 片段ID From 影像报告原型片段 Where 原型ID = Aid_In)
      Select RawToHex(ID) As ID,
             RawToHex(上级ID) As 上级ID,
             编码,
             名称,
             说明,
             节点类型,
             Nvl(组成, '<NULL/>') As 组成,
             学科,
             标签,
             是否私有,
             作者,
             最后编辑时间,
             EXTRACTValue( 适应条件, '/Root/Rad') 项目,
             EXTRACTValue( 适应条件, '/Root/Part') 部位, 
             EXTRACTValue( 适应条件, '/Root/Kind') 类别,
             EXTRACTValue( 适应条件, '/Root/Sex') 性别,
             0 as 提纲状态,
             0 as 适应状态
        From 影像报告片段清单
       Where Id Not In (Select 片段ID From TabFragmentId)
       Start With ID In (Select 片段ID From TabFragmentId)
      Connect By Prior 上级ID = ID
      Union All
      Select RawToHex(ID) As ID,
             RawToHex(上级ID) As 上级ID,
             编码,
             名称,
             说明,
             节点类型,
             Nvl(组成, '<NULL/>') As 组成,
             学科,
             标签,
             是否私有,
             作者,
             最后编辑时间,
             EXTRACTValue( 适应条件, '/Root/Rad') 项目,
             EXTRACTValue( 适应条件, '/Root/Part') 部位, 
             EXTRACTValue( 适应条件, '/Root/Kind') 类别,
             EXTRACTValue( 适应条件, '/Root/Sex') 性别,
             0 as 提纲状态,
             0 as 适应状态
        From 影像报告片段清单
       Start With ID In (Select 片段ID From TabFragmentId)
      Connect By Prior ID = 上级ID
       order by 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取数据表的最后更新时间>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
	Table_Name_In Varchar2
	) As
    v_sql Varchar2(4000);
  Begin
    v_sql := 'select max(最后编辑时间) maxvalue from ' || Table_Name_In;
    Open val For v_sql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取片段列表根据上级ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Bypid(
    Val           Out t_Refcur,
	Pid_In 影像报告片段清单.上级ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.上级ID) As 上级ID,
             T.编码,
             T.名称,
             T.说明,
             T.节点类型,
             Nvl(T.组成, '<NULL/>') As 组成,
             T.学科,
             T.标签,
             T.是否私有,
             T.作者,
             Nvl(T.适应条件.GetClobVal(), '<NULL/>') As 适应条件,
             T.最后编辑时间
        From 影像报告片段清单 T
       Where T.上级ID = Pid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --通过节点类型获取词句列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Byleaf(
    Val           Out t_Refcur,
	Leaf_In 影像报告片段清单.节点类型%Type
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.上级ID) As 上级ID,
             T.编码,
             T.名称,
             T.说明,
             T.节点类型,
             Nvl(T.组成, '<NULL/>') As 组成,
             T.学科,
             T.标签,
             T.是否私有,
             T.作者,
             T.最后编辑时间
        From 影像报告片段清单 T
       Where t.节点类型 = Leaf_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取值域信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Range_List_Byid(
    Val           Out t_Refcur,
	Id_In 影像报告值域清单.Id%Type
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.分类ID) As 分类ID,
             T.编码,
             T.名称,
             T.说明,
             T.数据类型,
             T.值域种类,
             Nvl(T.值域描述, '<NULL/>') As 值域描述
        From 影像报告值域清单 T
       Where ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --根据元素ID获取值域ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Getelementrid_By_Eid(
    Val           Out t_Refcur,
	Eid_In 影像报告元素清单.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(A.值域ID) As 值域ID
        From 影像报告元素清单 A
       Where a.Id = Eid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取计量单位列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_GetMasure_UnitList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select 编码, 名称, 说明, 前缀 From 影像报告计量单位;
  End p_GetMasure_UnitList;

  --获取文档种类信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Doc_Kind(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select a.编码, a.名称, a.说明 From 影像报告种类 A Order By a.编码;
  End;

  --功能：获取所有学科信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Subjects(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select rawtohex(b.字典id) ID, b.编号 As 编码, b.名称, b.简码, b.说明
        From 影像字典清单 A, 影像字典内容 B
       Where a.名称 = '专业学科'
         And a.Id = b.字典id
       Order By 编码;
  End;

  --查看是否存在相应的编码或者名称(用于导入导出)>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exits_Doc_Kinds(
    Val           Out t_Refcur,
	编码_In      Varchar2,
	名称_In      Varchar2,
	Tablename_In Varchar2
	) As
    v_Type Varchar2(50);
    n_Num  Number;
    v_Sql  Varchar2(100);
  Begin
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE 编码 =''' ||
             编码_In || ''' AND 名称 =''' || 名称_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num > 0 Then
      v_Type := '1';
    End If;
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE 编码 <>''' ||
             编码_In || ''' AND 名称 = ''' || 名称_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num > 0 Then
      v_Type := '2';
    End If;
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE 编码 =''' ||
             编码_In || ''' or 名称 =''' || 名称_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num = 0 Then
      v_Type := '3';
    End If;
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE 编码 =''' ||
             编码_In || ''' AND 名称 <>''' || 名称_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num > 0 Then
      v_Type := '4';
    End If;
    Open Val For
      Select v_Type As Type From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --是否存在相同的ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exist_Id(
    Val           Out t_Refcur,
	Id_In        Number,
	Tablename_In Varchar2
	) As
    v_Sql Varchar2(100);
    n_Num Number;
  Begin
    v_Sql := 'select count(id) from ' || Tablename_In || ' where id=''' ||
             Id_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    Open Val For
      Select Decode(n_Num, 0, 0, 1) Num From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --通过名称获取ID信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Id_By_Title(
    Val           Out t_Refcur,
	名称_In      Varchar2,
	Tablename_In Varchar2,
	Type_In      Varchar2
	) As
    v_Id  Varchar2(50);
    v_Sql Varchar2(100);
  Begin
    If Type_In = '1' Then
      v_Sql := 'select id from ' || Tablename_In || ' where 名称=''' || 名称_In || '''';
    Else
      v_Sql := 'select 编码 from ' || Tablename_In || ' where 名称=''' || 名称_In || '''';
    End If;
    v_Id := '';
    Execute Immediate v_Sql
      Into v_Id;
    Open Val For
      Select v_Id ID From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --通过简称片段清单
  Procedure p_Get_FragmentSampleName(
	Val           Out t_Refcur,
	简称_In Varchar2
	) as
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.上级ID) As 上级ID,
             t.编码,
             t.名称,
             t.说明,
             t.节点类型,
             Nvl(t.组成, '<NULL/>') As 组成,
             t.学科,
             t.标签,
             t.是否私有,
             t.作者,
             t.最后编辑时间
        From 影像报告片段清单 T
      Where t.名称 LIKE '%' || 简称_In || '%';
      --Where  F_TRANS_PINYIN_CAPITAL(t.名称) LIKE '%' || 简称_In || '%';
  End p_Get_FragmentSampleName;

  --更新ID对应的片段内容
  Procedure p_Update_PhraseContent(
	Id_In      影像报告片段清单.ID%Type,
	Name_In		影像报告片段清单.名称%Type,
	Content_In Varchar2
	) as
  Begin
    Update 影像报告片段清单 t 
	Set t.组成 = Content_In, t.名称=Name_In 
	Where t.id = Id_In;
  End p_Update_PhraseContent;

  --获取原型ID对应的第一层片段节点
  Procedure p_Get_FragmentData_LevelOne(
	Val           Out t_Refcur,
	AId_In 影像报告原型清单.ID%Type
	) as
  Begin
    Open Val For
      With TabFragmentId As
       (Select 片段ID From 影像报告原型片段 Where 原型ID = Aid_In)
      Select RawToHex(ID) As ID,
             RawToHex(上级ID) As 上级ID,
             编码,
             名称,
             说明,
             节点类型,
             Nvl(组成, '<NULL/>') As 组成,
             学科,
             标签,
             是否私有,
             作者,
             最后编辑时间
        From 影像报告片段清单
       Where Id In (Select 片段ID From TabFragmentId);
  
  End p_Get_FragmentData_LevelOne;

  -- 获取片段的下层节点
  Procedure p_GetFragmentDataListByFID(
	Val           Out t_Refcur,
	FId_In 影像报告片段清单.ID%Type
	) as
  Begin
    Open Val For
      Select RawToHex(ID) As ID,
             RawToHex(上级ID) As 上级ID,
             编码,
             名称,
             说明,
             节点类型,
             Nvl(组成, '<NULL/>') As 组成,
             学科,
             标签,
             是否私有,
             作者,
             最后编辑时间
        From 影像报告片段清单
       Where 上级ID = FId_In;
  End p_GetFragmentDataListByFID;

End b_PACS_RptCommon;
/

--90492:涂建华,2015-11-16,增加片段编码或名称是否重复的判断方法
--89077:许华锋,2015-11-24,报告片段适应条件
CREATE OR REPLACE Package b_PACS_RptFragments Is
  Type t_Refcur Is Ref Cursor;


  --功能：获取所有预备提纲
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	);
  --功能：获取所有短语分类
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	);
  --功能：获取当前用户学科所有短语包括父节点
  Procedure p_Get_All_Fragment(
    Val           Out t_Refcur,
    Subjects_In 影像报告片段清单.学科%Type
	) ;


  --功能：根据分类ID查找短语
  Procedure p_Get_Fragment_By_Typeid(
    Val           Out t_Refcur,
    Id_In 影像报告片段清单.ID%Type
	) ;


   Procedure p_Get_Label_By_Typeid(
     Val           Out t_Refcur,
	 Id_In 影像报告片段清单.ID%Type
	 ) ;

  --功能：新增短语分类
  Procedure p_Add_Fragmenttype(
    Id_In     影像报告片段清单.ID%Type,
    Pid_In    影像报告片段清单.上级ID%Type,
    Code_In   影像报告片段清单.编码%Type,
    Title_In  影像报告片段清单.名称%Type,
    Note_In   影像报告片段清单.说明%Type,
    Leaf_In   影像报告片段清单.节点类型%Type,
    Author_In 影像报告片段清单.作者%Type
    ) ;

  --功能：修改短语分类
  Procedure p_Edit_Fragmenttype(
    Id_In     影像报告片段清单.ID%Type,
    Pid_In    影像报告片段清单.上级ID%Type,
    Code_In   影像报告片段清单.编码%Type,
    Title_In  影像报告片段清单.名称%Type,
    Note_In   影像报告片段清单.说明%Type,
    Leaf_In   影像报告片段清单.节点类型%Type,
    Author_In 影像报告片段清单.作者%Type
    ) ;

  --功能：删除短语分类
   Procedure p_Del_Fragmenttype(
     Id_In 影像报告片段清单.ID%Type
	 );

    --功能：添加短语
  Procedure p_Add_Fragment(
     Id_In      影像报告片段清单.ID%Type,
    Pid_In      影像报告片段清单.上级ID%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) ;

   --功能：修改短语
  Procedure p_Edit_Fragment(
    Id_In       影像报告片段清单.ID%Type,
    Pid_In      影像报告片段清单.上级ID%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    );
   --功能：删除短语
  Procedure p_Del_Fragment(
    Id_In 影像报告片段清单.ID%Type
	);

  procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	);

  --功能：导入短语
  Procedure p_Import_Fragment(
    Id_In       影像报告片段清单.ID%Type,
    Pid_In      影像报告片段清单.上级ID%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) ;

procedure p_Get_Data_Last_Edit_Time(
  Val           Out t_Refcur,
  Table_Name_In varchar2
  );

   --功能：判断片段分类能否删除
  Procedure p_IsCanDel_FragmentType(
    Val           Out t_Refcur,
	Id_In 影像报告片段清单.Id%Type
	);

  --功能：根据片段ID，设置当前片段的适应条件
  Procedure p_Edit_FragmentConditionById
  (
    ID_In      In 影像报告片段清单.ID%Type,
    适应条件_In In 影像报告片段清单.适应条件%Type
  );
  
  --功能：根据片段的父ID，设置整个目录或子目录片段的适应条件
  Procedure p_Edit_FragmentConditionByPid
  (
    上级ID_In      In 影像报告片段清单.ID%Type,
    适应条件_In    In 影像报告片段清单.适应条件%Type
  );

  --功能：获取当前检查的片段适应条件
  Procedure p_Get_FraConditionByOrderId
  (
    Val           Out t_Refcur,
	医嘱ID_In    影像检查记录.医嘱ID%Type
  );

  --功能：获取影像检查类别
  Procedure p_Get_CheckLueKind
  (
    Val           Out t_Refcur
  );
  
  --功能：根据类别获取诊疗检查部位
  Procedure p_Get_CheckPartList
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  );
  
  --功能：根据类别获取影像检查项目
  Procedure p_Get_CheckRadListByKind
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  );
  
  --功能：根据诊疗编码获取影像检查项目
  Procedure p_Get_CheckRadListByCode
  (
    Val           Out t_Refcur,
    Code_In       Varchar2
  );

  --判断是否有相同的代码
  Procedure p_Get_HasSameCode
  (
  Val      Out t_Refcur,
  ID_In      In 影像报告片段清单.ID%Type,
  Code_In  影像报告片段清单.编码%Type
  );

  --判断是否有相同的名称
  Procedure p_Get_HasSameName
  (
  Val      Out t_Refcur,
  ID_In    In 影像报告片段清单.ID%Type,
  PID_In    In 影像报告片段清单.上级ID%Type,
  Name_In  In 影像报告片段清单.名称%Type,
  Author_In In  影像报告片段清单.作者%Type
  );

  End  b_PACS_RptFragments;
/
CREATE OR REPLACE Package Body b_PACS_RptFragments Is

  ------------------------------------------------------------------------
  --片段模块
  ------------------------------------------------------------------------

  --功能：获取所有预备提纲
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(ID) ID, 编码, 名称 From 影像报告预备提纲 Order By 编码;
  End p_Get_All_Phr_Onlines;

  --功能：获取所有短语分类
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID, Rawtohex(a.上级id) As 上级id, a.编码, a.名称, a.说明, a.节点类型
      From 影像报告片段清单 A
      Where a.节点类型 = 0
      Start With 上级id Is Null
      Connect By Prior ID = 上级id
      Order By 编码;
  End p_Get_All_Fragment_Class;

  --功能：获取当前用户学科所有短语包括父节点
  Procedure p_Get_All_Fragment(
    Val           Out t_Refcur,
    Subjects_In 影像报告片段清单.学科%Type
	) As
  Begin
    If Subjects_In <> '' Then
      Open Val For
        Select Rawtohex(a.Id) As ID, Rawtohex(a.上级id) As 上级id, a.编码, a.名称, a.说明, a.节点类型, Nvl(a.组成.GetClobVal(), '<NULL/>') As 组成, 
			a.学科, a.标签, a.是否私有, a.作者, Nvl(a.适应条件.GetClobVal(), '<NULL/>') As 适应条件,a.最后编辑时间, a.节点类型 As Image
        From 影像报告片段清单 A
        Where (a.学科 In (Select /*+rule*/
                         Column_Value As Lable
                        From Table(b_Pacs_Common.f_Str2list(Subjects_In, ','))
                        Intersect
                        Select /*+rule*/
                         Column_Value As Lable
                        From Table(b_Pacs_Common.f_Str2list(a.学科, ','))) And a.节点类型 <> 0) Or a.节点类型 = 0 Or a.学科 Is Null
        Order By 编码, 上级id;
    Else
      Open Val For
        Select Rawtohex(a.Id) As ID, Rawtohex(a.上级id) As 上级id, a.编码, a.名称, a.说明, a.节点类型, Nvl(a.组成.GetClobVal(), '<NULL/>') As 组成, 
			a.学科, a.标签, a.是否私有, a.作者, Nvl(a.适应条件.GetClobVal(), '<NULL/>') As 适应条件,a.最后编辑时间, a.节点类型 As Image
        From 影像报告片段清单 A
        Order By 上级id, 节点类型, 编码, 名称;
    End If;
  End p_Get_All_Fragment;

  --功能：根据分类ID查找短语
  Procedure p_Get_Fragment_By_Typeid(
    Val           Out t_Refcur,
    Id_In 影像报告片段清单.Id%Type
  ) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID, a.上级ID,a.编码, a.名称, a.说明, a.节点类型, Nvl(a.组成.GetClobVal(), '<NULL/>') As 组成, 
				a.学科, a.标签, a.是否私有, a.作者, Nvl(a.适应条件.GetClobVal(), '<NULL/>') As 适应条件, a.最后编辑时间,a.节点类型 As Image
      From 影像报告片段清单 A
      Where a.上级id = Hextoraw(Id_In) And a.节点类型 <> 0;
  End p_Get_Fragment_By_Typeid;

  --功能：查找某分类下所有短语标签
  Procedure p_Get_Label_By_Typeid(
    Val           Out t_Refcur,
    Id_In 影像报告片段清单.Id%Type
    ) As
  Begin
    Open Val For
      Select Distinct 标签 From 影像报告片段清单 Where 上级id = Hextoraw(Id_In) And 标签 Is Not Null;
  End p_Get_Label_By_Typeid;

  --功能：新增短语分类
  Procedure p_Add_Fragmenttype(
    Id_In     影像报告片段清单.Id%Type,
    Pid_In    影像报告片段清单.上级id%Type,
    Code_In   影像报告片段清单.编码%Type,
    Title_In  影像报告片段清单.名称%Type,
    Note_In   影像报告片段清单.说明%Type,
    Leaf_In   影像报告片段清单.节点类型%Type,
    Author_In 影像报告片段清单.作者%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From 影像报告片段清单
    Where 编码 = Code_In Or 名称 = Title_In And 节点类型 = 0 And 上级id = Hextoraw(Pid_In);

    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]分类名称或编码已经存在！[ZLSOFT]';
      Raise Err_Item;
    Else
      Insert Into 影像报告片段清单
        (ID, 上级id, 编码, 名称, 说明, 节点类型, 作者, 最后编辑时间)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Author_In, Sysdate);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Fragmenttype;

  --功能：修改短语分类
  Procedure p_Edit_Fragmenttype(
    Id_In     影像报告片段清单.Id%Type,
    Pid_In    影像报告片段清单.上级id%Type,
    Code_In   影像报告片段清单.编码%Type,
    Title_In  影像报告片段清单.名称%Type,
    Note_In   影像报告片段清单.说明%Type,
    Leaf_In   影像报告片段清单.节点类型%Type,
    Author_In 影像报告片段清单.作者%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From 影像报告片段清单
    Where (编码 = Code_In Or 名称 = Title_In) And 节点类型 = 0 And 上级id = Hextoraw(Pid_In) And ID <> Hextoraw(Id_In);

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]分类名称或编码已经存在！[ZLSOFT]';
      Raise Err_Item;
    Else
      Update 影像报告片段清单
      Set 上级id = Hextoraw(Pid_In), 编码 = Code_In, 名称 = Title_In, 说明 = Note_In, 节点类型 = Leaf_In, 作者 = Author_In,
          最后编辑时间 = Sysdate
      Where ID = Hextoraw(Id_In);
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Fragmenttype;

  --功能：删除短语分类
  Procedure p_Del_Fragmenttype(
    Id_In 影像报告片段清单.Id%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From 影像报告片段清单
    Where 节点类型 <> 0 And
          ID In (Select ID From 影像报告片段清单 Connect By Prior ID = 上级id Start With ID = Hextoraw(Id_In));

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]该分类下存在短语，暂不能删除！[ZLSOFT]';
      Raise Err_Item;
    Else
      Delete 影像报告片段清单 Where ID = Hextoraw(Id_In);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Del_Fragmenttype;

  --功能：添加短语
  Procedure p_Add_Fragment(
    Id_In       影像报告片段清单.Id%Type,
    Pid_In      影像报告片段清单.上级id%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) As
  Begin

      Insert Into 影像报告片段清单
        (ID, 上级id, 编码, 名称, 说明, 节点类型, 组成, 学科, 标签, 是否私有, 作者, 最后编辑时间)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Content_In, Subjects_In, Label_In,
         Private_In, Author_In, Sysdate);

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Fragment;

  --功能：修改短语
  Procedure p_Edit_Fragment(
    Id_In       影像报告片段清单.Id%Type,
    Pid_In      影像报告片段清单.上级id%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From 影像报告片段清单
    Where (编码 = Code_In Or 名称 = Title_In) And 节点类型 <> 0 And 上级id = Hextoraw(Pid_In) And ID <> Hextoraw(Id_In);

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]短语的名称或编码已经存在！[ZLSOFT]';
      Raise Err_Item;
    Else
      Update 影像报告片段清单
      Set 上级id = Hextoraw(Pid_In), 编码 = Code_In, 名称 = Title_In, 说明 = Note_In, 节点类型 = Leaf_In, 组成 = Content_In,
          学科 = Subjects_In, 标签 = Label_In, 是否私有 = Private_In, 作者 = Author_In, 最后编辑时间 = Sysdate
      Where ID = Hextoraw(Id_In);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Fragment;

  --
  Procedure p_Get_All_Fragment_List(Val Out t_Refcur) As
  Begin
    Open Val For
      Select Rawtohex(t.Id) As ID, Rawtohex(t.上级id) As 上级id, t.编码, t.名称, t.说明, t.节点类型, Nvl(t.组成.GetClobVal(), '<NULL/>') As 组成, t.学科, t.标签, t.是否私有, t.作者,
             t.最后编辑时间
      From 影像报告片段清单 T;
  End p_Get_All_Fragment_List;

  --功能：删除短语
  Procedure p_Del_Fragment(
    Id_In 影像报告片段清单.Id%Type
	) As
  Begin
    Delete 影像报告片段清单 Where ID = Hextoraw(Id_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Del_Fragment;

  --功能：导入短语
  Procedure p_Import_Fragment(
    Id_In       影像报告片段清单.Id%Type,
    Pid_In      影像报告片段清单.上级id%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) As
    v_Num Number(2);
  Begin
    Select Count(ID)
    Into v_Num
    From 影像报告片段清单
    Where ((编码 = Code_In Or 名称 = Title_In) And 上级id = Hextoraw(Pid_In)) Or
          (上级id Is Null And (编码 = Code_In Or 名称 = Title_In));

    If v_Num > 0 Then
      Update 影像报告片段清单
      Set 组成 = Content_In, 最后编辑时间 = Sysdate, 是否私有 = 0
      Where ((编码 = Code_In Or 名称 = Title_In) And 上级id = Hextoraw(Pid_In)) Or
            (上级id Is Null And (编码 = Code_In Or 名称 = Title_In));
    Else
      Insert Into 影像报告片段清单
        (ID, 上级id, 编码, 名称, 说明, 节点类型, 组成, 学科, 标签, 是否私有, 作者, 最后编辑时间)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Content_In, Subjects_In, Label_In,
         Private_In, Author_In, Sysdate);
    End If;

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Import_Fragment;

  --
  Procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
    Table_Name_In Varchar2
    ) As
    v_Sql Varchar2(4000);
  Begin
    v_Sql := 'select max(最后编辑时间) maxvalue from ' || Table_Name_In;
    Open Val For v_Sql;
  End p_Get_Data_Last_Edit_Time;
  
   --功能：判断片段分类能否删除
  Procedure p_IsCanDel_FragmentType(
    Val           Out t_Refcur,
	Id_In 影像报告片段清单.Id%Type
	) As
  Begin
    Open Val For
      Select Count(t.id) Count
        From 影像报告片段清单 t
       Where 上级id = Hextoraw(Id_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_IsCanDel_FragmentType;
  
  --功能：根据片段ID，设置当前片段的适应条件
  Procedure p_Edit_FragmentConditionById
  (
    ID_In      In 影像报告片段清单.ID%Type,
    适应条件_In In 影像报告片段清单.适应条件%Type
  )As
  Begin
    Update 影像报告片段清单 Set 适应条件 = 适应条件_In Where ID = Hextoraw(ID_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_FragmentConditionById;
  
  --功能：根据片段的父ID，设置整个目录或子目录片段的适应条件
  Procedure p_Edit_FragmentConditionByPid
  (
    上级ID_In      In 影像报告片段清单.ID%Type,
    适应条件_In In 影像报告片段清单.适应条件%Type
  )As
  Begin
    Update 影像报告片段清单 Set 适应条件 = 适应条件_In Where 上级ID = Hextoraw(上级ID_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_FragmentConditionByPid;

  --功能：获取当前检查的片段适应条件
  Procedure p_Get_FraConditionByOrderId(
    Val           Out t_Refcur,
	  医嘱ID_In    影像检查记录.医嘱ID%Type
	) As
  Begin
    Open Val For
      Select a.id, a.性别,c.影像类别, d.编码||' - '||d.名称 检查类别, c.影像类别||' - '||e.编码||' - '||e.名称 检查项目, A.医嘱内容
      From 病人医嘱记录 a, 病人医嘱发送 b, 影像检查记录 c, 影像检查类别 d, 诊疗项目目录 e
      Where a.id = b.医嘱id and b.医嘱id=c.医嘱id and c.影像类别 = d.编码 and a.诊疗项目id = e.id and a.id = 医嘱ID_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_FraConditionByOrderId;

  --功能：获取影像检查类别
  Procedure p_Get_CheckLueKind
  (
    Val           Out t_Refcur
  ) As
  Begin
    Open Val For
      Select 编码||' - '||名称 检查类别 From 影像检查类别;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckLueKind;
  
  --功能：根据类别获取诊疗检查部位
  Procedure p_Get_CheckPartList
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  ) As
  Begin
    Open Val For
      Select Distinct 类型||分组 IID, '' 上级ID, 类型||' - '||分组 诊疗部位 From 诊疗检查部位 a,
      Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) b Where a.类型 = b.Column_Value
      Union Select 类型||分组||名称 IID, 类型||分组 上级ID, 类型||' - '||名称 诊疗部位 From 诊疗检查部位 c,
      Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) d Where c.类型 = d.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckPartList;
  
  --功能：根据类别获取影像检查项目
  Procedure p_Get_CheckRadListByKind
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  ) As
  Begin
    Open Val For
      Select I.编码, r.影像类别||' - '||I.编码||' - '||I.名称 检查项目
      From 诊疗项目目录 I, 影像检查项目 R, Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) M
      Where I.ID = R.诊疗项目id And R.影像类别=M.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckRadListByKind;
  
  --功能：根据诊疗编码获取影像检查项目
  Procedure p_Get_CheckRadListByCode
  (
    Val           Out t_Refcur,
    Code_In       Varchar2
  ) As
  Begin
    Open Val For
      Select I.编码, r.影像类别||' - '||I.编码||' - '||I.名称 检查项目
      From 诊疗项目目录 I, 影像检查项目 R, Table(Cast(f_Str2list(''||Code_In||'') As zlTools.t_Strlist)) M
      Where I.ID = R.诊疗项目id And I.编码=M.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckRadListByCode;

  --判断是否有相同的代码
  Procedure p_Get_HasSameCode
  (
  Val      Out t_Refcur,
  ID_In      In 影像报告片段清单.ID%Type,
  Code_In  影像报告片段清单.编码%Type
  ) As
  Begin
  Open Val For
    Select Count(1) From 影像报告片段清单 Where ID<>ID_In And 编码=Code_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_HasSameCode;

  --判断是否有相同的名称
  Procedure p_Get_HasSameName
  (
  Val      Out t_Refcur,
  ID_In    In 影像报告片段清单.ID%Type,
  PID_In    In 影像报告片段清单.上级ID%Type,
  Name_In  In 影像报告片段清单.名称%Type,
  Author_In In  影像报告片段清单.作者%Type
  ) As
  Begin
  Open Val For
    Select Count(1) From 影像报告片段清单 Where 上级ID=PID_In And 作者=Author_In And ID<>ID_In And 名称=Name_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_HasSameName;

End  b_PACS_RptFragments;
/


--88395:刘尔旋,2015-10-30,就诊详情返回退费审核状态
--88874:刘尔旋,2015-10-28,发药窗口
Create Or Replace Procedure Zl_Third_Getvisitinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:根据挂号单号获取该次就诊详情(医嘱为主要显示)
  --入参:Xml_In:
  --<IN>
  --    <GHDH>挂号单号</GHDH>
  --    <JSKLB>结算卡类别</JSKLB>
  --    <MXGL>明细过滤</MXGL> 0-不过滤,明细包含治疗 1-过滤,明细不包含治疗,默认为1
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  --  <GH>
  --     <YYSJ>预约时间</YYSJ> //yyyy-mm-dd hh24:mi:ss
  --     <JZSJ></JZSJ>      //实际就诊时间
  --     <DJH></DJH>        //单据号
  --     <JE></JE>          //金额
  --     <DJLX></DJLX>      //单据类型,1-收费单，4-挂号单
  --     <KDSJ></KDSJ>      //开单时间
  --     <JKFS></JKFS>      //缴款方式,0-挂号或预约缴款;1-预约不缴款
  --     <ZFZT></ZFZT>  //支付状态,0-待支付，1-已支付，2-已退费
  --     <SFJSK></SFJSK>    //是否结算卡支付，0-否，1-是
  --  </GH>
  --  <YZLIST>
  --     <YZ>                   //医嘱返回与HIS中显示的内容相同
  --        <YZID><YZID>        //医嘱ID，返回组医嘱ID
  --        <YZLX><YZLX>        //医嘱类型,如处方、检查、检验
  --         <YZMC></YZMC>        //医嘱名称
  --        <ZXKS></ZXKS>       //执行科室
  --        <ZXKSID></ZXKSID>   //执行科室ID
  --        <FYCK></FYCK>       //发药窗口
  --        <YZMX>
  --           <MX>
  --              <YZNR></YZNR>        //医嘱内容
  --              <ZXZT></ZXZT>        //医嘱执行状态
  --              <SFFY>是否发药</SFFY> // 0-否 ，1-是
  --              <GG>规格</GG>
  --              <SL>数量</SL>
  --              <DW>计算单位</DW>
  --              <BZDJ>标准单价</BZDJ>
  --              <YSJE>应收金额</YSJE>
  --              <SSJE>实收金额</SSJE>
  --           </MX>
  --           <MX/>
  --        </YZMX>
  --        <BG></BG>                   //是否已出报告，是否签名
  --        <BGLY></BGLY>               //是否外检项目,1-院内项目，2-外检项目
  --        <BGLYSM></BGLYSM>           //外检项目说明
  --        <JZBG></JZBG>                //禁止显示报告。0-允许，1-禁止
  --        <JZTS></JZTS>                 //提示文字。对于禁止查看的报告，可返回用于提示病人的信息
  --        <BLID></BLID>              //病历ID，如果<BG>字段为1，该值不为空
  --        <DJLIST>
  --           <DJ>                //费用单据信息
  --              <DJH></DJH>      //费用单据号
  --              <DJLX></DJLX>    //单据类型
  --              <JE></JE>        //单据总金额
  --              <KDSJ></KDSJ>    //开单时间
  --              <ZFZT></ZFZT>    //支付状态,0-待支付，1-已支付，2-已退费,3-退费申请中,4-审核通过,5-审核未通过
  --              <SHSM></SHSM>    //审核说明,审核未通过原因
  --              <SFJSK></SFJSK>  //是否结算卡支付，0-否，1-是
  --           </DJ>
  --           <DJ/>
  --        </DJLIST>
  --     </YZ>
  --  </YZLIST>
  --    <ERROR><MSG></MSG></ERROR>                      //如果错误返回
  --</OUTPUT>

  --------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  x_Templet Xmltype; --模板XML

  v_卡类别   Varchar2(100);
  n_卡类别id Number(18);
  v_挂号单   Varchar2(10);
  v_排队号码 Varchar2(10);
  n_Temp     Number(18);
  v_队列名称 排队叫号队列.队列名称%Type;

  n_Count Number(18);

  v_Temp       Varchar2(32767); --临时XML
  v_队列       Varchar2(32767);
  v_No         Varchar2(50);
  n_Add_Djlist Number(1); --是否增加了DJLIST的
  n_性质       Number(2);
  n_组医嘱id   Number(18);
  n_独立医嘱   Number(8);
  n_执行科室id Number(18);
  v_执行科室   Varchar2(50);
  n_退款金额   病人预交记录.冲预交%Type;
  n_明细过滤   Number(3);
  n_退费状态   病人退费申请.状态%Type;
  v_申请原因   病人退费申请.申请原因%Type;
  v_审核原因   病人退费申请.审核原因%Type;
  v_发药窗口   门诊费用记录.发药窗口%Type;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/MXGL')
  Into v_挂号单, v_卡类别, n_明细过滤
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_挂号单 Is Null Then
    v_Err_Msg := '不能找到指定的挂号单号(当前挂号单号为空)';
    Raise Err_Item;
  End If;
  If n_明细过滤 Is Null Then
    n_明细过滤 := 1;
  End If;
  n_Add_Djlist := 0;

  v_Err_Msg := Null;
  If v_卡类别 Is Not Null Then
    Begin
      n_卡类别id := To_Number(v_卡类别);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
  
    If n_卡类别id = 0 Then
      Begin
        Select ID, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_Err_Msg
        From 医疗卡类别
        Where 名称 = v_卡类别;
      Exception
        When Others Then
          v_Err_Msg := '卡类别:' || v_卡类别 || '不存在!';
      End;
    
    Else
    
      Begin
        Select ID, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_Err_Msg
        From 医疗卡类别
        Where ID = n_卡类别id;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  End If;
  n_性质 := 4;
  --1.获取挂号数据
  Begin
    Select Replace(摘要, '划价:', '') Into v_No From 病人挂号记录 Where NO = v_挂号单;
  Exception
    When Others Then
      v_No := Null;
  End;

  If v_No Is Not Null Then
    Select Count(*) Into n_Count From 门诊费用记录 Where NO = v_No And 记录性质 = 1;
    If n_Count <> 0 Then
      n_性质 := 1;
    End If;
  End If;
  If n_性质 = 4 Then
    v_No := v_挂号单;
  End If;

  n_Count := 0;
  For c_挂号 In (Select a.Id, v_No As NO, n_性质 As 记录性质, a.执行部门id, c.名称 As 执行部门,
                      To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, To_Char(a.预约时间, 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                      a.接收时间, To_Char(a.发生时间, 'yyyy-mm-dd HH24:mi:ss') As 就诊时间, a.号别, a.号序, b.金额, a.记录状态,
                      Decode(Nvl(a.执行状态, 0), 0, '等待接诊', 1, '完成就诊', 2, '正在就诊', -1, '取消就诊') As 执行状态,
                      Decode(Nvl(b.结帐id, 0), 0, 0, 1) As 支付标志, Decode(Nvl(a.记录性质, 0), 2, 1, 0) As 缴款方式, b.结帐id As 结帐id
               From 病人挂号记录 A,
                    (Select Max(Decode(记录状态, 0, 0, 2, 0, Nvl(结帐id, 0))) As 结帐id, Sum(实收金额) As 金额
                      From 门诊费用记录 B
                      Where 记录性质 = n_性质 And NO = v_No) B, 部门表 C
               Where a.No = v_挂号单 And a.执行部门id = c.Id(+)) Loop
  
    If Nvl(c_挂号.记录状态, 0) <> 1 Then
      v_Err_Msg := '单据号:' || v_挂号单 || '已经被退号!';
      Raise Err_Item;
    End If;
  
    Begin
      Select 排队号码, 队列名称
      Into v_排队号码, v_队列名称
      From 排队叫号队列
      Where 业务id = c_挂号.Id And Nvl(业务类型, 0) = 0;
    Exception
      When Others Then
        v_排队号码 := Null;
    End;
    If v_排队号码 Is Not Null Then
      --业务id_In ,业务类型_In 排队号码_In Number := Null
      n_Temp := Zl_Getsequencebeforperons(c_挂号.Id, 0, v_排队号码, v_队列名称);
      v_队列 := v_队列 || '<DL><XH>' || v_排队号码 || '</XH><QMRS>' || n_Temp || '</QMRS></DL>';
    End If;
    n_Temp := 0;
    If Nvl(n_卡类别id, 0) <> 0 Then
      Begin
        Select 1
        Into n_Temp
        From 病人预交记录
        Where 结帐id = c_挂号.结帐id And 记录性质 = 4 And 记录状态 In (1, 3) And 卡类别id = n_卡类别id And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    End If;
  
    v_Temp := '<DJH>' || c_挂号.No || '</DJH>';
    v_Temp := v_Temp || '<YYSJ>' || c_挂号.预约时间 || '</YYSJ>';
    v_Temp := v_Temp || '<JZSJ>' || c_挂号.就诊时间 || '</JZSJ>';
    v_Temp := v_Temp || '<KDSJ>' || c_挂号.登记时间 || '</KDSJ>';
    v_Temp := v_Temp || '<JKFS>' || c_挂号.缴款方式 || '</JKFS>';
    v_Temp := v_Temp || '<JE>' || c_挂号.金额 || '</JE>';
    v_Temp := v_Temp || '<DJLX>' || n_性质 || '</DJLX>';
    v_Temp := v_Temp || '<ZFZT>' || c_挂号.支付标志 || '</ZFZT>';
    v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
    If v_队列 Is Not Null Then
      v_Temp := v_Temp || v_队列;
    End If;
    v_Temp := '<GH>' || v_Temp || '</GH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    n_Count := n_Count + 1;
  End Loop;

  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '未找到指定的挂号单据:' || v_挂号单 || '!';
    Raise Err_Item;
  End If;

  --2.组建医嘱及费用相关数据
  n_组医嘱id := 0;

  For c_医嘱 In (With 医嘱费用 As
                  (Select 医嘱id, 发送号, 记录性质, NO, Max(Nvl(执行状态, 0)) As 执行状态
                  From (Select b.医嘱id, b.发送号, b.记录性质, b.No, Nvl(b.执行状态, 0) As 执行状态
                         From 病人医嘱记录 A, 病人医嘱发送 B
                         Where a.挂号单 = v_挂号单 And a.Id = b.医嘱id(+)
                         Union All
                         Select b.医嘱id, b.发送号, b.记录性质, b.No, Nvl(c.执行状态, 0) As 执行状态
                         From 病人医嘱记录 A, 病人医嘱附费 B, 病人医嘱发送 C
                         Where a.挂号单 = v_挂号单 And a.Id = b.医嘱id(+) And b.医嘱id = c.医嘱id(+) And b.发送号 = c.发送号(+))
                  Group By 医嘱id, 发送号, 记录性质, NO)
                 
                 Select Nvl(a.相关id, a.Id) As 组id, Decode(a.相关id, Null, 0, 1) As 附医嘱, a.Id, a.相关id, e.发药窗口,
                        Max(Decode(a.诊疗类别, 'E', Decode(q.操作类型, '2', '处方', '4', '处方', '6', '检验', m.名称), m.名称)) As 医嘱类型,
                        a.执行科室id, d.名称 As 执行科室, Decode(a.相关id, Null, a.医嘱内容, Null) As 组医嘱内容,
                        Max(Decode(a.诊疗类别, '5', 1, '6', 1, '7', 1, 0) * Decode(Nvl(e.执行状态, 0), 1, 1, 3, 1, 0)) As 发药状态,
                        Decode(a.相关id, Null, Null, q.名称) As 明细医嘱内容, s.规格, (e.数次 * e.付数) As 数量, e.计算单位 As 单位,
                        Decode(Nvl(b.执行状态, 0), 0, '未执行', 1, '完全执行', 2, '拒绝执行', 3, '正在执行', '正在执行') As 执行状态,
                        Max(Decode(p.审核时间, Null, Decode(C1.完成时间, Null, 0, 1), 1)) As 是否已出报告, c.病历id, e.No, e.记录性质 As 单据类型,
                        Max(e.标准单价) As 标准单价, Sum(e.应收金额) As 应收金额, Sum(e.实收金额) As 实收金额,
                        To_Char(e.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 开单时间, Decode(Nvl(e.记录状态, 0), 0, 0, 3, 2, 1) As 支付状态,
                        a.病人id
                 
                 From 病人医嘱记录 A, 医嘱费用 B, 病人医嘱报告 C, 电子病历记录 C1, 部门表 D, 门诊费用记录 E, 诊疗项目类别 M, 诊疗项目目录 Q, 收费项目目录 S, 检验标本记录 P
                 Where a.Id = b.医嘱id(+) And a.执行科室id = d.Id(+) And c.病历id = C1.Id(+) And a.Id = c.医嘱id(+) And
                       a.Id = p.医嘱id(+) And b.医嘱id = e.医嘱序号(+) And e.收费细目id = s.Id(+) And b.No = e.No(+) And
                       b.记录性质 = e.记录性质(+) And e.记录状态(+) <> 2 And a.挂号单 = v_挂号单 And a.诊疗类别 = m.编码(+) And
                       a.诊疗项目id = q.Id(+) And a.医嘱状态 In (3, 8)
                 Group By a.Id, a.婴儿, a.序号, a.相关id, e.发药窗口, a.诊疗类别, a.执行科室id, d.名称, a.医嘱内容, q.名称, s.规格, e.数次 * e.付数,
                          e.计算单位, Decode(Nvl(b.执行状态, 0), 0, '未执行', 1, '完全执行', 2, '拒绝执行', 3, '正在执行', '正在执行'), C1.完成时间,
                          Decode(c.病历id, Null, 0, 1), c.病历id, e.No, e.记录性质, e.登记时间, Decode(Nvl(e.记录状态, 0), 0, 0, 3, 2, 1),
                          p.审核时间, a.病人id
                 Order By 组id, 附医嘱, Nvl(a.婴儿, 0), a.序号) Loop
    If Nvl(n_Add_Djlist, 0) = 0 Then
      --增加DJList节点
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<YZLIST></YZLIST>')) Into x_Templet From Dual;
      n_Add_Djlist := 1;
    End If;
  
    If n_组医嘱id <> Nvl(c_医嘱.组id, 0) Then
      n_组医嘱id := Nvl(c_医嘱.组id, 0);
    
      Zl_Third_Custom_Getdeptinfo(n_组医嘱id, n_执行科室id, v_执行科室);
    
      If Nvl(n_执行科室id, 0) = 0 Then
        If c_医嘱.医嘱类型 = '检验' Then
          --检验医嘱以显示采集科室
          n_执行科室id := c_医嘱.执行科室id;
          v_执行科室   := c_医嘱.执行科室;
        Else
          Begin
            Select b.Id, b.名称, c.发药窗口
            Into n_执行科室id, v_执行科室, v_发药窗口
            From 病人医嘱记录 A, 部门表 B, 门诊费用记录 C
            Where a.Id = c.医嘱序号 And a.相关id = n_组医嘱id And a.执行科室id = b.Id And Rownum <= 1;
          Exception
            When Others Then
              n_执行科室id := c_医嘱.执行科室id;
              v_执行科室   := c_医嘱.执行科室;
              v_发药窗口   := c_医嘱.发药窗口;
          End;
        End If;
      End If;
    
      v_Temp := '<YZID>' || n_组医嘱id || '</YZID>';
      v_Temp := v_Temp || '<YZLX>' || c_医嘱.医嘱类型 || '</YZLX>';
      v_Temp := v_Temp || '<YZMC>' || c_医嘱.组医嘱内容 || '</YZMC>';
      v_Temp := v_Temp || '<ZXKS>' || v_执行科室 || '</ZXKS>';
      v_Temp := v_Temp || '<ZXKSID>' || n_执行科室id || '</ZXKSID>';
      v_Temp := v_Temp || '<FYCK>' || v_发药窗口 || '</FYCK>';
      v_Temp := v_Temp || '<BG>' || c_医嘱.是否已出报告 || '</BG>';
      v_Temp := v_Temp || Zl_Third_Custom_Getrptfrom(n_组医嘱id);
      v_Temp := v_Temp || Zl_Third_Custom_Rptlimit(c_医嘱.病人id, n_组医嘱id);
      If Nvl(c_医嘱.是否已出报告, 0) = 1 And c_医嘱.病历id Is Not Null Then
        v_Temp := v_Temp || '<BLID>' || c_医嘱.病历id || '</BLID>';
      End If;
      v_Temp := '<YZ 医嘱ID="' || n_组医嘱id || '">' || v_Temp || '<YZMX></YZMX><DJLIST></DJLIST></YZ>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      For v_费用 In (
                   
                   Select a.No, Mod(a.记录性质, 10) As 单据类型, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 开单时间,
                           Max(Decode(Nvl(a.记录状态, 0), 0, 0, 3, 2, 1)) As 支付状态, Sum(a.实收金额) As 单据金额, Max(a.结帐id) As 结算卡支付
                   From 门诊费用记录 A
                   Where (a.No, a.记录性质) In
                         (Select Distinct q.No, q.记录性质
                          From 病人医嘱记录 M, 病人医嘱发送 Q
                          Where m.Id = q.医嘱id(+) And (m.Id = n_组医嘱id Or m.相关id = n_组医嘱id)
                          Union All
                          Select Distinct q.No, q.记录性质
                          From 病人医嘱记录 M, 病人医嘱附费 Q
                          Where m.Id = q.医嘱id(+) And (m.Id = n_组医嘱id Or m.相关id = n_组医嘱id)) And
                         Nvl(a.记录状态, 0) In (0, 1, 3)
                   Group By a.No, Mod(a.记录性质, 10), To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss')) Loop
        Begin
          Select 1
          Into n_Temp
          From 病人预交记录 A, 门诊费用记录 B
          Where a.结帐id = b.结帐id And b.No = v_费用.No And Mod(b.记录性质, 10) = 1 And b.记录状态 In (1, 3) And a.卡类别id = n_卡类别id And
                Rownum < 2;
        Exception
          When Others Then
            n_Temp := 0;
        End;
        Begin
          Select -1 * Sum(结帐金额)
          Into n_退款金额
          From 门诊费用记录 B
          Where b.No = v_费用.No And Mod(b.记录性质, 10) = 1 And b.记录状态 = 2;
        Exception
          When Others Then
            n_退款金额 := 0;
        End;
        Begin
          Select 状态, 申请原因, 审核原因
          Into n_退费状态, v_申请原因, v_审核原因
          From 病人退费申请
          Where NO = v_费用.No And Mod(记录性质, 10) = Mod(v_费用.单据类型, 10);
        Exception
          When Others Then
            n_退费状态 := -1;
            v_申请原因 := '';
            v_审核原因 := '';
        End;
      
        v_Temp := '<DJH>' || v_费用.No || '</DJH>';
        v_Temp := v_Temp || '<DJLX>' || v_费用.单据类型 || '</DJLX>';
        v_Temp := v_Temp || '<JE>' || v_费用.单据金额 || '</JE>';
        v_Temp := v_Temp || '<KDSJ>' || v_费用.开单时间 || '</KDSJ>';
        If n_退费状态 = -1 Then
          v_Temp := v_Temp || '<ZFZT>' || v_费用.支付状态 || '</ZFZT>';
        Else
          If n_退费状态 = 0 Then
            v_Temp := v_Temp || '<ZFZT>3</ZFZT>';
          End If;
          If n_退费状态 = 1 Then
            If v_费用.支付状态 = 2 Then
              v_Temp := v_Temp || '<ZFZT>2</ZFZT>';
            Else
              v_Temp := v_Temp || '<ZFZT>4</ZFZT>';
            End If;
          End If;
          If n_退费状态 = 2 Then
            v_Temp := v_Temp || '<ZFZT>5</ZFZT>';
          End If;
        End If;
      
        If n_退费状态 = -1 Then
          v_Temp := v_Temp || '<SHSM>' || '' || '</SHSM>';
        Else
          If n_退费状态 = 0 Then
            v_Temp := v_Temp || '<SHSM>' || v_申请原因 || '</SHSM>';
          End If;
          If n_退费状态 = 1 Then
            v_Temp := v_Temp || '<SHSM>' || v_审核原因 || '</SHSM>';
          End If;
          If n_退费状态 = 2 Then
            v_Temp := v_Temp || '<SHSM>' || v_审核原因 || '</SHSM>';
          End If;
        End If;
      
        v_Temp := v_Temp || '<YTJE>' || Nvl(n_退款金额, 0) || '</YTJE>';
        v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
        v_Temp := '<DJ>' || v_Temp || '</DJ>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@医嘱ID="' || n_组医嘱id || '"]/DJLIST', Xmltype(v_Temp))
        Into x_Templet
        From Dual;
      End Loop;
    End If;
  
    --只有一条记录的医嘱，在明细中增加该条医嘱，以获取执行状态
    Select Decode(Count(*), 0, 1, 0) Into n_独立医嘱 From 病人医嘱记录 Where 相关id = n_组医嘱id;
    If n_独立医嘱 = 1 Then
      v_Temp := '<YZNR>' || c_医嘱.组医嘱内容 || '</YZNR>';
      v_Temp := v_Temp || '<GG>' || c_医嘱.规格 || '</GG>';
      v_Temp := v_Temp || '<SFFY>' || c_医嘱.发药状态 || '</SFFY>';
      v_Temp := v_Temp || '<SL>' || c_医嘱.数量 || '</SL>';
      v_Temp := v_Temp || '<DW>' || c_医嘱.单位 || '</DW>';
      v_Temp := v_Temp || '<BZDJ>' || Nvl(c_医嘱.标准单价, 0) || '</BZDJ>';
      v_Temp := v_Temp || '<YSJE>' || Nvl(c_医嘱.应收金额, 0) || '</YSJE>';
      v_Temp := v_Temp || '<SSJE>' || Nvl(c_医嘱.实收金额, 0) || '</SSJE>';
      v_Temp := v_Temp || '<ZXZT>' || c_医嘱.执行状态 || '</ZXZT>';
      v_Temp := '<MX>' || v_Temp || '</MX>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@医嘱ID="' || n_组医嘱id || '"]/YZMX', Xmltype(v_Temp))
      Into x_Templet
      From Dual;
    End If;
  
    If Nvl(c_医嘱.附医嘱, 0) = 1 Then
      If n_明细过滤 = 0 Or (n_明细过滤 = 1 And c_医嘱.医嘱类型 <> '治疗') Then
        v_Temp := '<YZNR>' || c_医嘱.明细医嘱内容 || '</YZNR>';
        v_Temp := v_Temp || '<GG>' || c_医嘱.规格 || '</GG>';
        v_Temp := v_Temp || '<SL>' || c_医嘱.数量 || '</SL>';
        v_Temp := v_Temp || '<DW>' || c_医嘱.单位 || '</DW>';
        v_Temp := v_Temp || '<SFFY>' || c_医嘱.发药状态 || '</SFFY>';
        v_Temp := v_Temp || '<ZXZT>' || c_医嘱.执行状态 || '</ZXZT>';
        v_Temp := v_Temp || '<BZDJ>' || Nvl(c_医嘱.标准单价, 0) || '</BZDJ>';
        v_Temp := v_Temp || '<YSJE>' || Nvl(c_医嘱.应收金额, 0) || '</YSJE>';
        v_Temp := v_Temp || '<SSJE>' || Nvl(c_医嘱.实收金额, 0) || '</SSJE>';
        v_Temp := '<MX>' || v_Temp || '</MX>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@医嘱ID="' || n_组医嘱id || '"]/YZMX', Xmltype(v_Temp))
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

--89873:刘尔旋,2015-10-27,取号支持关联过滤
Create Or Replace Procedure Zl_Third_Getnolist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取号源列表
  --入参:Xml_In:
  --<IN>
  --  <RQ>日期</RQ>
  --  <KSID>科室ID</KSID>
  --  <YSID>医生ID</YSID>
  --  <YSXM>医生姓名</YSXM>
  --  <HZDW>支付宝</HZDW>    //合作单位，传入了的时候，只取合作单位的号;为空时，只取非合作单位的号
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  --  <GROUP>
  --    <RQ>日期</RQ>
  --    <HBLIST>
  --     <HB>
  --        <HM>235</HM>       //号码
  --        <YSID>549</YSID>      //医生ID
  --        <YS>张锐</YS>       //医生姓名
  --        <KSID>123</KSID>   //科室ID
  --        <KSMC>内科</KSMC>   //科室名称
  --        <ZC>主治医师</ZC> //职称
  --        <XMID>10086<XMID> //挂号项目的ID
  --        <XMMC>挂号费</XMMC> //挂号项目的名称
  --        <YGHS>0</YGHS>      //已挂号数
  --        <SYHS>99</SYHS>   //剩余号数
  --        <PRICE>15</PRICE>      //价格
  --        <HL>普通</HL>       //挂号类型
  --        <HCXH>1</HCXH>    //是否存在缓冲序号时间段，1-存在 0或者空-不存在
  --        <FSD>0</FSD>      //是否分时段
  --        <FWMC>白天</FWMC>     //号别时段
  --        <HBTIME>(08:00-17:59)</HBTIME> //可挂时间
  --     <SPANLIST>
  --            <SPAN>
  --                  <SJD/>      //时间段
  --                  <SL/>      //数量
  --            </SPAN>
  --            ……
  --          </SPANLIST>
  --      </HB>
  --      <HB>
  --      ……
  --      </HB>
  --    </HBLIST>
  --  </GROUP>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  d_日期         Date;
  n_科室id       病人挂号记录.执行部门id%Type;
  n_医生id       人员表.Id%Type;
  v_医生姓名     人员表.姓名%Type;
  v_星期         挂号安排限制.限制项目%Type;
  v_时间段       Varchar2(100);
  v_合作单位     挂号合作单位.名称%Type;
  n_分时段       Number(3);
  n_单个剩余     Number(5);
  n_已挂数       Number(5);
  n_合约已挂数   Number(5);
  n_合计金额     收费价目.现价%Type;
  n_合约总数量   Number(5);
  n_合约剩余数量 Number(5);
  n_最大可用数量 Number(5);
  n_合约模式     Number(3); --合约模式:1-合约单位限数量模式 0-合约单位指定序号模式
  n_非合约       Number(3);
  n_是否预留     Number(3);
  d_加号时间     Date;
  n_缓冲序号     Number(3);
  n_时段数量     Number(5);
  n_预留数量     Number(5);
  n_特殊预约     Number(3);
  n_禁用         Number(3);
  v_剩余数量     Varchar2(100);
  v_Timetemp     Varchar2(100);
  v_Temp         Varchar2(32767); --临时XML
  v_Xmlmain      Clob; --临时XML
  c_Xmlmain      Clob; --临时XML
  x_Templet      Xmltype; --模板XML
  v_Err_Msg      Varchar2(200);
  n_Exists       Number(2);
  v_Sql          Varchar2(20000);
  Type c_Main Is Ref Cursor;
  r_科室id   挂号安排.科室id%Type;
  r_号类     挂号安排.号类%Type;
  r_科室名称 部门表.名称%Type;
  r_医生姓名 挂号安排.医生姓名%Type;
  r_医生id   挂号安排.医生id%Type;
  r_职称     人员表.专业技术职务%Type;
  r_号码     挂号安排.号码%Type;
  r_安排id   挂号安排.Id%Type;
  r_计划id   挂号安排计划.Id%Type;
  r_排班     挂号安排.周日%Type;
  r_项目id   挂号安排.项目id%Type;
  r_项目名称 收费项目目录.名称%Type;
  r_序号控制 挂号安排.序号控制%Type;
  r_限号数   挂号安排限制.限号数%Type;
  r_限约数   挂号安排限制.限约数%Type;
  r_已挂数   病人挂号汇总.已挂数%Type;
  r_已约数   病人挂号汇总.已约数%Type;
  r_已接收   病人挂号汇总.其中已接收%Type;
  r_价格     收费价目.现价%Type;
  r_No       c_Main;
  n_Curcount Number(3);

  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/KSID'),
         Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/YSXM'), Extractvalue(Value(A), 'IN/HZDW')
  Into d_日期, n_科室id, n_医生id, v_医生姓名, v_合作单位
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  --日期节点为空的情况
  If d_日期 Is Null Then
    d_日期 := Trunc(Sysdate);
  End If;

  Select Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;
  n_合约剩余数量 := 0;

  v_Sql := 'Select a.*, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数, Nvl(Hz.其中已接收, 0) As 已接收, b.现价 As 价格 ';
  v_Sql := v_Sql ||
           'From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码, ';
  v_Sql := v_Sql || ' Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数 ';
  v_Sql := v_Sql || 'From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id, ';
  v_Sql := v_Sql || 'Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制, ';
  v_Sql := v_Sql || 'Decode(To_Char(:1, ''D''), ''1'', Ap.周日, ''2'', Ap.周一, ''3'', Ap.周二, ''4'', Ap.周三, ''5'', Ap.周四, ';
  v_Sql := v_Sql || ' ''6'', Ap.周五, ''7'', Ap.周六, Null) As 排班, Xz.限约数, Xz.限号数 ';
  v_Sql := v_Sql || 'From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz ';
  v_Sql := v_Sql || 'Where Ap.科室id = Bm.Id(+) ';

  n_Curcount := 2;
  If Nvl(n_科室id, 0) <> 0 Then
    v_Sql      := v_Sql || 'And Ap.科室id = :2 ';
    n_Curcount := n_Curcount + 1;
  End If;
  If Nvl(n_医生id, 0) <> 0 Then
    If n_Curcount = 2 Then
      v_Sql := v_Sql || 'And Ap.医生id = :2 ';
    Else
      v_Sql := v_Sql || 'And Ap.医生id = :3 ';
    End If;
    n_Curcount := n_Curcount + 1;
  End If;
  If Nvl(v_医生姓名, '_') <> '_' Then
    If n_Curcount = 2 Then
      v_Sql := v_Sql || 'And Ap.医生姓名 = :2 ';
    End If;
    If n_Curcount = 3 Then
      v_Sql := v_Sql || 'And Ap.医生姓名 = :3 ';
    End If;
    If n_Curcount = 4 Then
      v_Sql := v_Sql || 'And Ap.医生姓名 = :4 ';
    End If;
    n_Curcount := n_Curcount + 1;
  End If;

  v_Sql      := v_Sql || 'And Ap.停用日期 Is Null And :' || n_Curcount ||
                ' Between Nvl(Ap.开始时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Nvl(Ap.终止时间, To_Date(''3000 - 01 - 01'', ''YYYY-MM-DD'')) And Xz.安排id(+) = Ap.Id And ';
  v_Sql      := v_Sql || ' Xz.限制项目(+) = Decode(To_Char(:' || n_Curcount ||
                ', ''D''), ''1'', ''周日'', ''2'', ''周一'', ''3'', ''周二'', ''4'', ''周三'', ''5'', ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || ' ''周四'', ''6'', ''周五'', ''7'', ''周六'', Null) And Not Exists ';
  v_Sql      := v_Sql || '(Select Rownum ';
  v_Sql      := v_Sql || 'From 挂号安排停用状态 Ty ';
  v_Sql      := v_Sql || 'Where Ty.安排id = Ap.Id And :' || n_Curcount ||
                ' Between Ty.开始停止时间 And Ty.结束停止时间) And Not Exists ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || '(Select Rownum ';
  v_Sql      := v_Sql || 'From 挂号安排计划 Jh Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And ';
  v_Sql      := v_Sql || ':' || n_Curcount ||
                ' Between Nvl(Jh.生效时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.失效时间, To_Date(''3000-01-01'', ''YYYY-MM-DD''))) ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Union All ';
  v_Sql      := v_Sql ||
                'Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id, Jh.Id As 计划id, ';
  v_Sql      := v_Sql || 'Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,Decode(To_Char(:' || n_Curcount ||
                ', ''D''), ''1'', Jh.周日, ''2'', Jh.周一, ''3'', Jh.周二, ''4'', Jh.周三, ''5'', Jh.周四, ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || ' ''6'', Jh.周五, ''7'', Jh.周六, Null) As 排班, Xz.限约数, Xz.限号数 ';
  v_Sql      := v_Sql || 'From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz ';
  v_Sql      := v_Sql || 'Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null ';

  If Nvl(n_科室id, 0) <> 0 Then
    v_Sql      := v_Sql || 'And Ap.科室id = :' || n_Curcount || ' ';
    n_Curcount := n_Curcount + 1;
  End If;
  If Nvl(n_医生id, 0) <> 0 Then
    v_Sql      := v_Sql || 'And Ap.医生id = :' || n_Curcount || ' ';
    n_Curcount := n_Curcount + 1;
  End If;
  If Nvl(v_医生姓名, '_') <> '_' Then
    v_Sql      := v_Sql || 'And Ap.医生姓名 = :' || n_Curcount || ' ';
    n_Curcount := n_Curcount + 1;
  End If;

  v_Sql      := v_Sql || ' And :' || n_Curcount ||
                ' Between Nvl(Jh.生效时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.失效时间, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Xz.计划id(+) = Jh.Id And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Xz.限制项目(+) = Decode(To_Char(:' || n_Curcount ||
                ', ''D''), ''1'', ''周日'', ''2'', ''周一'', ''3'', ''周二'', ''4'', ''周三'', ''5'', ''周四'', ''6'', ''周五'', ''7'', ''周六'', Null) And Not Exists ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || '(Select Rownum From 挂号安排停用状态 Ty Where Ty.安排id = Ap.Id And :' || n_Curcount ||
                ' Between Ty.开始停止时间 And Ty.结束停止时间) And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || '(Jh.生效时间, Jh.安排id) = (Select Max(Sxjh.生效时间) As 生效时间, 安排id From 挂号安排计划 Sxjh ';
  v_Sql      := v_Sql || ' Where Sxjh.审核时间 Is Not Null And :' || n_Curcount ||
                ' Between Nvl(Sxjh.生效时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Nvl(Sxjh.失效时间, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Sxjh.安排id = Jh.安排id ';
  v_Sql      := v_Sql || 'Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy ';
  v_Sql      := v_Sql || 'Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A, ';
  v_Sql      := v_Sql || '病人挂号汇总 Hz, 收费价目 B ';
  v_Sql      := v_Sql || 'Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(:' || n_Curcount || ') And a.项目id = b.收费细目id And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Nvl(b.终止日期, To_Date(''3000-1-1'', ''YYYY-Mm-DD'')) > :' || n_Curcount || ' ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'And b.执行日期 <= :' || n_Curcount || ' ';
  If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') <> '_' Then
    Open r_No For v_Sql
      Using d_日期, n_科室id, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
  End If;
  If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') = '_' Then
    Open r_No For v_Sql
      Using d_日期, n_科室id, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
  End If;
  If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') = '_' Then
    Open r_No For v_Sql
      Using d_日期, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
  End If;
  If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') <> '_' Then
    Open r_No For v_Sql
      Using d_日期, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
  End If;
  If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') = '_' Then
    Open r_No For v_Sql
      Using d_日期, n_科室id, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
  End If;
  If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') <> '_' Then
    Open r_No For v_Sql
      Using d_日期, n_科室id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
  End If;
  If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') <> '_' Then
    Open r_No For v_Sql
      Using d_日期, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
  End If;
  Loop
    Fetch r_No
      Into r_科室id, r_号类, r_科室名称, r_医生姓名, r_医生id, r_职称, r_号码, r_安排id, r_计划id, r_排班, r_项目id, r_项目名称, r_序号控制, r_限号数, r_限约数,
           r_已挂数, r_已约数, r_已接收, r_价格;
    Exit When r_No%NotFound;
    If r_计划id <> 0 Then
      Select Sign(Count(Rownum))
      Into n_分时段
      From 挂号安排计划 Jh, 挂号计划时段 Sd
      Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
            Sd.星期 =
            Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And
            Rownum < 2;
    Else
      Select Sign(Count(Rownum))
      Into n_分时段
      From 挂号安排 Ap, 挂号安排时段 Sd
      Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
            Sd.星期 =
            Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And
            Rownum < 2;
    End If;
    If n_分时段 = 0 Then
      v_Temp := '';
      If v_合作单位 Is Not Null And r_序号控制 = 1 Then
        If r_计划id <> 0 Then
          Select Nvl(Sum(数量), 0)
          Into n_合约总数量
          From 合作单位计划控制
          Where 计划id = r_计划id And 合作单位 = v_合作单位 And
                限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                              '周六', Null);
          Select Count(1)
          Into n_合约模式
          From 合作单位计划控制
          Where 计划id = r_计划id And 合作单位 = v_合作单位 And
                限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                              '周六', Null) And 序号 = 0;
        Else
          Select Nvl(Sum(数量), 0)
          Into n_合约总数量
          From 合作单位安排控制
          Where 安排id = r_安排id And 合作单位 = v_合作单位 And
                限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                              '周六', Null);
          Select Count(1)
          Into n_合约模式
          From 合作单位安排控制
          Where 安排id = r_安排id And 合作单位 = v_合作单位 And
                限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                              '周六', Null) And 序号 = 0;
        End If;
        If n_合约模式 = 0 Then
          If r_计划id <> 0 Then
            Select Count(1)
            Into n_合约已挂数
            From 病人挂号记录 A
            Where 号别 = r_号码 And 记录状态 = 1 And 发生时间 Between Trunc(d_日期) And Trunc(d_日期 + 1) - 1 / 60 / 60 / 24 And Exists
             (Select 1
                   From 合作单位计划控制
                   Where 计划id = r_计划id And 合作单位 = v_合作单位 And
                         限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                       '周五', '7', '周六', Null) And 序号 = a.号序 And 数量 <> 0);
          Else
            Select Count(1)
            Into n_合约已挂数
            From 病人挂号记录 A
            Where 号别 = r_号码 And 记录状态 = 1 And 发生时间 Between Trunc(d_日期) And Trunc(d_日期 + 1) - 1 / 60 / 60 / 24 And Exists
             (Select 1
                   From 合作单位安排控制
                   Where 安排id = r_安排id And 合作单位 = v_合作单位 And
                         限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                       '周五', '7', '周六', Null) And 序号 = a.号序 And 数量 <> 0);
          End If;
        Else
          Begin
            Select Count(1)
            Into n_合约已挂数
            From 病人挂号记录
            Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                  Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
          Exception
            When Others Then
              n_合约已挂数 := 0;
          End;
        End If;
        If n_合约总数量 = 0 Then
          n_合约剩余数量 := 0;
        Else
          n_合约剩余数量 := n_合约总数量 - n_合约已挂数;
          If n_合约剩余数量 > (Nvl(r_限号数, 0) - r_已挂数) Then
            n_合约剩余数量 := Nvl(r_限号数, 0) - r_已挂数;
          End If;
        End If;
      End If;
    Else
      v_Temp := '<SPANLIST>';
      If r_计划id <> 0 Then
        Select Max(结束时间)
        Into d_加号时间
        From 挂号计划时段
        Where 计划id = r_计划id And 星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                            '6', '周五', '7', '周六', Null);
        If r_序号控制 = 1 Then
          If Trunc(d_日期) = Trunc(Sysdate) Then
            n_特殊预约 := 0;
          Else
            Select Nvl(Max(Jh.是否预约), 0)
            Into n_特殊预约
            From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                          To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                          To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                   From 挂号安排计划 Jh, 挂号计划时段 Sd
                   Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                         Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                        '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
            Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号 And
                  Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1;
          End If;
        
          For r_Time In (Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约, 0 As 已约数,
                                Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数, Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段

                         From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                                       To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                       To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                From 挂号安排计划 Jh, 挂号计划时段 Sd
                                Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                                      Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                     '周四', '6', '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                         Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号 And
                               Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1
                         Order By 序号) Loop
            If v_合作单位 Is Not Null Then
              Begin
                Select 1
                Into n_合约模式
                From 合作单位计划控制
                Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = 0 And 合作单位 = v_合作单位;
              Exception
                When Others Then
                  n_合约模式 := 0;
              End;
            Else
              n_合约模式 := 0;
            End If;
            If r_Time.剩余数 = 0 Then
              n_单个剩余 := 0;
            Else
              n_单个剩余 := r_Time.限制数量;
            End If;
            If v_合作单位 Is Null Or n_合约模式 = 1 Then
              Begin
                Select 1
                Into n_Exists
                From 合作单位计划控制
                Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              If n_Exists = 0 Then
                If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                  Null;
                Else
                  Begin
                    Select 1
                    Into n_是否预留
                    From 挂号序号状态
                    Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                  Exception
                    When Others Then
                      n_是否预留 := 0;
                  End;
                  If n_是否预留 = 0 Then
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                  '</SPAN>';
                    n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                  End If;
                End If;
              End If;
            Else
              Begin
                Select 1
                Into n_Exists
                From 合作单位计划控制
                Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              Begin
                Select 0
                Into n_非合约
                From 合作单位计划控制
                Where 计划id = r_计划id And 合作单位 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_非合约 := 1;
              End;
              If n_Exists = 1 Or n_非合约 = 1 Then
                If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                  Null;
                Else
                  Begin
                    Select 1
                    Into n_是否预留
                    From 挂号序号状态
                    Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                  Exception
                    When Others Then
                      n_是否预留 := 0;
                  End;
                  If n_是否预留 = 0 Then
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                      '</SPAN>';
                    n_合约剩余数量 := n_合约剩余数量 + 1;
                  End If;
                End If;
              End If;
            End If;
          End Loop;
        Else
          n_最大可用数量 := Nvl(r_限约数, Nvl(r_限号数, 0)) - Nvl(r_已约数, 0);
          For r_Time In (Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约,
                                Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 已约数,
                                Jh.限制数量 - Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 剩余数,
                                Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段
                         From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                                       To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                       To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                From 挂号安排计划 Jh, 挂号计划时段 Sd
                                Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                                      Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                     '周四', '6', '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                         Where Jh.号码 = Zt.号码(+) And Jh.开始时间 = Zt.日期(+) And
                               Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1
                         Group By Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约
                         Order By Jh.序号) Loop
            If v_合作单位 Is Not Null Then
              Begin
                Select 1
                Into n_合约模式
                From 合作单位计划控制
                Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = 0 And 合作单位 = v_合作单位;
              Exception
                When Others Then
                  n_合约模式 := 0;
              End;
            Else
              n_合约模式 := 0;
            End If;
            n_单个剩余 := r_Time.剩余数;
            If v_合作单位 Is Null Or n_合约模式 = 1 Then
              Begin
                Select 1
                Into n_Exists
                From 合作单位计划控制
                Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              If n_Exists = 0 Then
                If n_最大可用数量 < n_单个剩余 Then
                  v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                '</SPAN>';
                  n_时段数量 := Nvl(n_时段数量, 0) + n_最大可用数量;
                Else
                  v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                '</SPAN>';
                  n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                End If;
              End If;
            Else
              Begin
                Select 1
                Into n_Exists
                From 合作单位计划控制
                Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              Begin
                Select 0
                Into n_非合约
                From 合作单位计划控制
                Where 计划id = r_计划id And 合作单位 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_非合约 := 1;
              End;
              If n_Exists = 1 Or n_非合约 = 1 Then
                If n_最大可用数量 < n_单个剩余 Then
                  v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                    '</SPAN>';
                  n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                Else
                  v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                  n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                End If;
              End If;
            End If;
          End Loop;
        End If;
      Else
        Select Max(结束时间)
        Into d_加号时间
        From 挂号安排时段
        Where 安排id = r_安排id And 星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                            '6', '周五', '7', '周六', Null);
        If r_序号控制 = 1 Then
          If Trunc(d_日期) = Trunc(Sysdate) Then
            n_特殊预约 := 0;
          Else
            Select Nvl(Max(Ap.是否预约), 0)
            Into n_特殊预约
            From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                          To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                          To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                   From 挂号安排 Ap, 挂号安排时段 Sd
                   Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                         Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                        '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
            Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号 And
                  Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1;
          End If;
          For r_Time In (Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约, 0 As 已约数,
                                Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数, Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段

                         From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                                       To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                       To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                From 挂号安排 Ap, 挂号安排时段 Sd
                                Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                                      Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                     '周四', '6', '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                         Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号 And
                               Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1
                         Order By 序号) Loop
            If v_合作单位 Is Not Null Then
              Begin
                Select 1
                Into n_合约模式
                From 合作单位安排控制
                Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = 0 And 合作单位 = v_合作单位;
              Exception
                When Others Then
                  n_合约模式 := 0;
              End;
            Else
              n_合约模式 := 0;
            End If;
            If r_Time.剩余数 = 0 Then
              n_单个剩余 := 0;
            Else
              n_单个剩余 := r_Time.限制数量;
            End If;
            If v_合作单位 Is Null Or n_合约模式 = 1 Then
              Begin
                Select 1
                Into n_Exists
                From 合作单位安排控制
                Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              If n_Exists = 0 Then
                If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                  Null;
                Else
                  Begin
                    Select 1
                    Into n_是否预留
                    From 挂号序号状态
                    Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                  Exception
                    When Others Then
                      n_是否预留 := 0;
                  End;
                  If n_是否预留 = 0 Then
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                  '</SPAN>';
                    n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                  End If;
                End If;
              End If;
            Else
              Begin
                Select 1
                Into n_Exists
                From 合作单位安排控制
                Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              Begin
                Select 0
                Into n_非合约
                From 合作单位安排控制
                Where 安排id = r_安排id And 合作单位 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_非合约 := 1;
              End;
              If n_Exists = 1 Or n_非合约 = 1 Then
                If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                  Null;
                Else
                  Begin
                    Select 1
                    Into n_是否预留
                    From 挂号序号状态
                    Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                  Exception
                    When Others Then
                      n_是否预留 := 0;
                  End;
                  If n_是否预留 = 0 Then
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                      '</SPAN>';
                    n_合约剩余数量 := n_合约剩余数量 + 1;
                  End If;
                End If;
              End If;
            End If;
          End Loop;
        Else
          n_最大可用数量 := Nvl(r_限约数, Nvl(r_限号数, 0)) - Nvl(r_已约数, 0);
          For r_Time In (Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约,
                                Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 已约数,
                                Ap.限制数量 - Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 剩余数,
                                Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段
                         From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                                       To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                       To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                From 挂号安排 Ap, 挂号安排时段 Sd
                                Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                                      Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                     '周四', '6', '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                         Where Ap.号码 = Zt.号码(+) And Ap.开始时间 = Zt.日期(+) And
                               Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1
                         Group By Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约
                         Order By Ap.序号) Loop
            If v_合作单位 Is Not Null Then
              Begin
                Select 1
                Into n_合约模式
                From 合作单位安排控制
                Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = 0 And 合作单位 = v_合作单位;
              Exception
                When Others Then
                  n_合约模式 := 0;
              End;
            Else
              n_合约模式 := 0;
            End If;
            n_单个剩余 := r_Time.剩余数;
            If v_合作单位 Is Null Or n_合约模式 = 1 Then
              Begin
                Select 1
                Into n_Exists
                From 合作单位安排控制
                Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              If n_Exists = 0 Then
                If n_最大可用数量 < n_单个剩余 Then
                  v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                '</SPAN>';
                  n_时段数量 := Nvl(n_时段数量, 0) + n_最大可用数量;
                Else
                  v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                '</SPAN>';
                  n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                End If;
              End If;
            Else
              Begin
                Select 1
                Into n_Exists
                From 合作单位安排控制
                Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              Begin
                Select 0
                Into n_非合约
                From 合作单位安排控制
                Where 安排id = r_安排id And 合作单位 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_非合约 := 1;
              End;
              If n_Exists = 1 Or n_非合约 = 1 Then
                If n_最大可用数量 < n_单个剩余 Then
                  v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                    '</SPAN>';
                  n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                Else
                  v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                  n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                End If;
              End If;
            End If;
          End Loop;
        End If;
      End If;
    End If;
    If v_合作单位 Is Not Null Then
      If Nvl(r_计划id, 0) <> 0 Then
        Begin
          Select 0
          Into n_非合约
          From 合作单位计划控制
          Where 计划id = r_计划id And 合作单位 = v_合作单位 And Rownum < 2;
        Exception
          When Others Then
            n_非合约 := 1;
        End;
      Else
        Begin
          Select 0
          Into n_非合约
          From 合作单位安排控制
          Where 安排id = r_安排id And 合作单位 = v_合作单位 And Rownum < 2;
        Exception
          When Others Then
            n_非合约 := 1;
        End;
      End If;
    End If;
    If v_合作单位 Is Null Or n_非合约 = 1 Then
      If r_限号数 = 0 Then
        v_剩余数量 := '';
      Else
        If Nvl(r_计划id, 0) <> 0 Then
          Select Sum(数量)
          Into n_合约总数量
          From 合作单位计划控制
          Where 计划id = r_计划id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                '周四', '6', '周五', '7', '周六', Null);
        Else
          Select Sum(数量)
          Into n_合约总数量
          From 合作单位安排控制
          Where 安排id = r_安排id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                '周四', '6', '周五', '7', '周六', Null);
        End If;
        Begin
          Select Count(1)
          Into n_合约已挂数
          From 病人挂号记录
          Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_日期) And
                Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
        Exception
          When Others Then
            n_合约已挂数 := 0;
        End;
        Select Count(1)
        Into n_预留数量
        From 挂号序号状态
        Where 状态 = 3 And 号码 = r_号码 And Trunc(日期) = Trunc(d_日期);
        If Trunc(d_日期) = Trunc(Sysdate) Then
          If Nvl(n_合约总数量, 0) = 0 Then
            v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_预留数量;
          Else
            v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_合约总数量 + n_合约已挂数 - n_预留数量;
          End If;
          n_已挂数 := r_已挂数;
          If Nvl(n_时段数量, 0) < v_剩余数量 And n_分时段 <> 0 Then
            n_缓冲序号 := 1;
            v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_加号时间, 'hh24:mi:ss') || '-' || '</SJD>' || '<SL>' ||
                          To_Number(v_剩余数量 - Nvl(n_时段数量, 0)) || '</SL>' || '</SPAN>';
          Else
            n_缓冲序号 := 0;
          End If;
        Else
          If Nvl(n_合约总数量, 0) = 0 Then
            v_剩余数量 := r_限约数 - r_已约数 - n_预留数量;
            If v_剩余数量 Is Null Then
              v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_预留数量;
            End If;
          Else
            v_剩余数量 := r_限约数 - r_已约数 - n_合约总数量 + n_合约已挂数 - n_预留数量;
            If v_剩余数量 Is Null Then
              v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_合约总数量 + n_合约已挂数 - n_预留数量;
            End If;
          End If;
          n_已挂数 := r_已挂数;
        End If;
      End If;
    Else
      If Nvl(r_计划id, 0) <> 0 Then
        If v_合作单位 Is Not Null Then
          Begin
            Select 1
            Into n_合约模式
            From 合作单位计划控制
            Where 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                '7', '周六', Null) And 计划id = r_计划id And 序号 = 0 And 合作单位 = v_合作单位;
          Exception
            When Others Then
              n_合约模式 := 0;
          End;
        Else
          n_合约模式 := 0;
        End If;
        Select Sum(数量)
        Into n_合约总数量
        From 合作单位计划控制
        Where 计划id = r_计划id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                              '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
      Else
        If v_合作单位 Is Not Null Then
          Begin
            Select 1
            Into n_合约模式
            From 合作单位安排控制
            Where 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                '7', '周六', Null) And 安排id = r_安排id And 序号 = 0 And 合作单位 = v_合作单位;
          Exception
            When Others Then
              n_合约模式 := 0;
          End;
        Else
          n_合约模式 := 0;
        End If;
        Select Sum(数量)
        Into n_合约总数量
        From 合作单位安排控制
        Where 安排id = r_安排id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                              '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
      End If;
      If n_合约模式 = 0 Then
        v_剩余数量   := n_合约剩余数量;
        n_已挂数     := r_已挂数;
        n_合约已挂数 := Nvl(n_合约总数量, 0) - n_合约剩余数量;
      Else
        n_已挂数 := r_已挂数;
        Begin
          Select Count(1)
          Into n_合约已挂数
          From 病人挂号记录
          Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
        Exception
          When Others Then
            n_合约已挂数 := 0;
        End;
        If Nvl(n_合约总数量, 0) = 0 Then
          v_剩余数量 := '0';
        Else
          v_剩余数量 := n_合约总数量 - n_合约已挂数;
        End If;
      End If;
    End If;
    Select To_Char(开始时间, 'hh24:mi') Into v_Timetemp From 时间段 Where 时间段 = r_排班;
    v_时间段 := v_Timetemp || '-';
    Select To_Char(终止时间, 'hh24:mi') Into v_Timetemp From 时间段 Where 时间段 = r_排班;
    v_时间段 := v_时间段 || v_Timetemp;
    If v_Temp Is Not Null Then
      v_Temp := v_Temp || '</SPANLIST>';
    End If;
    If v_合作单位 Is Not Null Then
      If Nvl(r_计划id, 0) <> 0 Then
        Begin
          Select 1
          Into n_禁用
          From 合作单位计划控制
          Where 计划id = r_计划id And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
        Exception
          When Others Then
            n_禁用 := 0;
        End;
      Else
        Begin
          Select 1
          Into n_禁用
          From 合作单位安排控制
          Where 安排id = r_安排id And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
        Exception
          When Others Then
            n_禁用 := 0;
        End;
      End If;
    End If;
    If Nvl(n_禁用, 0) = 0 Then
      --从项金额计算
      n_合计金额 := r_价格;
      For r_Subfee In (Select 现价, 从项数次
                       From 收费从属项目 A, 收费价目 B
                       Where a.主项id = r_项目id And a.从项id = b.收费细目id And Sysdate Between b.执行日期 And
                             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
        n_合计金额 := n_合计金额 + r_Subfee.现价 * r_Subfee.从项数次;
      End Loop;
      If Trunc(Sysdate) = Trunc(d_日期) Then
        Begin
          Select 1
          Into n_Exists
          From (Select 时间段
                 From 时间段
                 Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') < '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')) Or
                       ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                       Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS'),
                               '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS'))))
          Where 时间段 = r_排班;
        Exception
          When Others Then
            n_Exists := 0;
        End;
      Else
        n_Exists := 1;
      End If;
      If n_Exists = 1 Then
        If v_剩余数量 > 0 Then
          c_Xmlmain := '<HB>' || '<HM>' || r_号码 || '</HM>' || '<YSID>' || r_医生id || '</YSID>' || '<YS>' || r_医生姓名 ||
                       '</YS>' || '<KSID>' || r_科室id || '</KSID>' || '<KSMC>' || r_科室名称 || '</KSMC>' || '<ZC>' || r_职称 ||
                       '</ZC>' || '<XMID>' || r_项目id || '</XMID>' || '<XMMC>' || r_项目名称 || '</XMMC>' || '<YGHS>' ||
                       n_已挂数 || '</YGHS>' || '<SYHS>' || v_剩余数量 || '</SYHS>' || '<PRICE>' || n_合计金额 || '</PRICE>' ||
                       '<HCXH>' || n_缓冲序号 || '</HCXH>' || '<HL>' || r_号类 || '</HL>' || '<FSD>' || n_分时段 || '</FSD>' ||
                       '<HBTIME>' || v_时间段 || '</HBTIME>' || '<FWMC>' || r_排班 || '</FWMC>' || v_Temp || '</HB>';
          v_Xmlmain := v_Xmlmain || c_Xmlmain;
        Else
          c_Xmlmain := '<HB>' || '<HM>' || r_号码 || '</HM>' || '<YSID>' || r_医生id || '</YSID>' || '<YS>' || r_医生姓名 ||
                       '</YS>' || '<KSID>' || r_科室id || '</KSID>' || '<KSMC>' || r_科室名称 || '</KSMC>' || '<ZC>' || r_职称 ||
                       '</ZC>' || '<XMID>' || r_项目id || '</XMID>' || '<XMMC>' || r_项目名称 || '</XMMC>' || '<YGHS>' ||
                       n_已挂数 || '</YGHS>' || '<SYHS>' || v_剩余数量 || '</SYHS>' || '<PRICE>' || n_合计金额 || '</PRICE>' ||
                       '<HL>' || r_号类 || '</HL>' || '<FSD>' || n_分时段 || '</FSD>' || '<HBTIME>' || v_时间段 || '</HBTIME>' ||
                       '<FWMC>' || r_排班 || '</FWMC>' || '</HB>';
          v_Xmlmain := v_Xmlmain || c_Xmlmain;
        End If;
      End If;
    End If;
    n_合约剩余数量 := 0;
    n_合约总数量   := 0;
    n_时段数量     := 0;
    n_禁用         := 0;
    n_非合约       := 0;
  End Loop;
  Close r_No;
  v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_日期, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain || '</HBLIST>' ||
               '</GROUP>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getnolist;
/

--90046:刘尔旋,2015-11-17,支付宝结帐修改
--0000:张永康,2015-10-26,保留3个索引，提供重建其他索引的功能
--90874:张永康,2015-11-24,为级联删除外键引用的主表创建临时触发器
Create Or Replace Procedure Zl1_Datamove_Reb
(
  System_In    In Number,
  Speedmode_In In Number,
  Func_In      In Number,
  Enable_In    In Number := 0,
  Parallel_In  In Number := 0,
  Rebscope_In  In Number := 0
) As
  --功能：在历史数据转出之前，禁用触发器、自动作业、约束、索引，转出之后启用这些对象，以及重建待转出索引，收回标记转出所用索引的空间 
  --参数： 
  --System_In:    应用系统编号,100=标准版 
  --speedmode_in：数据转出模式，0-在线模式，1-离线模式（在客户端停用时，转出期间禁用转出表的主键、唯一键、外键约束和索引，以加快已转数据的删除操作） 
  --func_in:      1=触发器，2=自动作业，3=约束，4=索引，5=重建待转出索引，6-收回标记转出所用索引的空间，7-重组表的存储空间（move），并恢复被禁用的约束和索引 ,8-重建标记转出查询所需索引以外的其他索引 
  --Enable_in:    0-禁用，1=启用，对func_in值为1-4有效 
  --rebScope_in:   Func_In=6时，指重建索引的范围(0-经济核算类,1-经济核算类及医嘱类,2-全部)，Func_In=7时指Move表的范围(0-经济核算类，1-全部) 

  v_Sql      Varchar2(4000);
  n_Do       Number(1);
  n_Parallel Number(1);
  v_Tbs      Varchar2(100);

 --转出标记中的SQL查询所需的索引
  v_Indexeswithtag Varchar2(4000) := '门诊费用记录_IX_结帐ID,住院费用记录_IX_结帐ID,费用补充记录_IX_结算ID,费用补充记录_IX_登记时间,病人预交记录_IX_主页ID,病人预交记录_IX_结帐ID,病人预交记录_IX_收款时间,门诊费用记录_IX_登记时间,门诊费用记录_IX_医嘱序号,住院费用记录_IX_登记时间,病人结帐记录_IX_收费时间,病人结帐记录_IX_病人id' ||
                                     ',药品收发记录_IX_费用ID,收发记录补充信息_IX_收发ID,输液配药内容_IX_收发ID,药品留存计划_IX_留存ID,药品签名明细_IX_收发ID' ||
                                     ',人员借款记录_IX_借出时间,人员收缴记录_IX_登记时间,人员暂存记录_IX_收缴ID,人员暂存记录_IX_登记时间,票据领用记录_IX_登记时间,票据使用明细_IX_领用ID,票据打印明细_IX_使用ID' ||
                                     ',病人挂号记录_IX_登记时间,病人医嘱发送_IX_发送时间,病人医嘱记录_IX_挂号单,病人医嘱记录_IX_主页ID,病人医嘱记录_IX_相关ID' ||
                                     ',病案主页_IX_出院日期,住院费用记录_IX_病人ID,病人过敏记录_IX_病人ID,病人诊断记录_IX_病人ID,病人手麻记录_IX_主页ID' ||
                                     ',病人护理记录_IX_主页ID,病人护理内容_IX_记录id,病人护理文件_IX_主页ID,病人护理数据_IX_文件ID,病人护理明细_IX_记录ID,病人护理打印_IX_文件ID' ||
                                     ',电子病历记录_IX_病人ID,病人医嘱报告_IX_病历ID,影像报告驳回_IX_医嘱ID,报告查阅记录_IX_病历ID,病人诊断记录_IX_病历ID' ||
                                     ',病人临床路径_IX_病人ID,病人合并路径_IX_首要路径记录ID,病人路径执行_IX_路径记录ID,病人出径记录_IX_路径记录ID,病人诊断医嘱_IX_医嘱ID' ||
                                     ',影像报告记录_IX_医嘱ID,影像报告操作记录_IX_医嘱ID,影像申请单图像_IX_医嘱ID,影像收藏内容_IX_医嘱ID,检验标本记录_IX_医嘱ID,检验项目分布_IX_标本ID,检验分析记录_IX_标本ID' ||
                                     ',检验操作记录_IX_标本ID,检验图像结果_IX_标本ID,检验拒收记录_IX_医嘱ID,检验普通结果_IX_检验标本ID'; 

  --转出标记中的SQL查询所需的索引(主键及唯一键对应的索引)
  v_Constraintswithtag Varchar2(4000) := '病人预交记录_UQ_NO,病人结帐记录_UQ_NO,病人结帐记录_PK,门诊费用记录_UQ_NO,住院费用记录_UQ_NO,医保结算明细_PK' ||
                                         ',病人卡结算对照_PK,费用补充记录_PK,病人卡结算记录_PK,三方结算交易_PK,三方退款信息_PK,输液配药记录_PK,药品签名记录_PK,票据打印内容_PK,病人挂号记录_PK,病人挂号汇总_UQ_日期,病人转诊记录_UQ_NO' ||
                                         ',病人护理活动项目_UQ_页号,病人护理要素内容_UQ_页号,产程要素内容_PK,电子病历记录_PK,电子病历附件_PK,电子病历格式_PK,电子病历内容_UQ_对象序号,电子病历图形_PK,疾病申报记录_PK' ||
                                         ',病人合并路径评估_PK,病人路径评估_PK,病人路径变异_PK,病人路径指标_UQ_评估指标,病人路径医嘱_PK' ||
                                         ',病人医嘱记录_PK,病人医嘱报告_PK,病人医嘱计价_UQ_收费细目ID,病人医嘱附费_PK,病人医嘱附件_PK,病人医嘱执行_PK,医嘱执行时间_PK,医嘱执行打印_PK,病人医嘱打印_UQ_医嘱ID,输血申请记录_PK,输血检验结果_PK' ||
                                         ',病人诊断记录_PK,病人医嘱状态_PK,医嘱签名记录_PK,病人医嘱发送_PK,诊疗单据打印_PK,医嘱执行计价_PK,执行打印记录_PK' ||
                                         ',影像检查记录_PK,影像检查序列_UQ_序列号,影像检查图象_UQ_图像号,影像危急值记录_UQ_医嘱ID' ||
                                         ',检验申请项目_PK,检验质控记录_PK,检验签名记录_PK,检验试剂记录_PK,检验质控报告_PK,检验药敏结果_PK,人员收缴记录_PK,人员收缴明细_PK,人员收缴票据_PK,人员收缴对照_PK';

  --功能：1.禁用或启用引用转出表主键的他表外键,避免删除主表记录时对子表每行记录执行一次SQL查询或删除 
  --      2.禁用或启用主键或唯一键约束（禁用时会自动删除对应的索引，启用时自动创建），以提高数据删除性能 
  --例如：病人医嘱发送_FK_医嘱ID，如果这些外键所在的表，数据未转出（未在zlbaktables表中定义），执行前会检查并限制转出。 
  Procedure Setconstraintstatus As
    v_Pcol Varchar2(50);
    v_Fcol Varchar2(50);
    v_Del  Varchar2(4000);
  Begin
    --禁用时，先禁用引用转出表主键的他表外键，再禁用转出表的主键 
    If Enable_In = 0 Then
      --1.在线模式转出时，由于有业务产生删除操作，所以，对于级联删除的外键，用触发器来替代对子表数据的删除操作
      If Speedmode_In = 0 Then
        For Rp In (Select Distinct a.Table_Name As Ptable_Name, a.Constraint_Name
                   From User_Constraints A, User_Constraints C, zlBakTables B
                   Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
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
        
          v_Sql := 'Create Or Replace Trigger ' || Rp.Ptable_Name || '_Cascade_Del' || Chr(10) ||
                   '    After Delete On ' || Rp.Ptable_Name || Chr(10) || '    For Each Row' || Chr(10) || 'Begin' ||
                   Chr(10) || '    If :Old.待转出 Is Null Then ' || v_Del || Chr(10) || '    End If; ' || Chr(10) ||
                   'End ' || Rp.Ptable_Name || '_Cascade_Del;';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.禁用引用转出表主键的他表外键
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'ENABLED'
                Order By a.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.禁用主键或唯一键索引(离线转出时)
      If Speedmode_In = 1 Then
        --必须删除索引，否则即使skip_unusable_indexes为true，也无法删除存在Unusable状态的唯一性索引的表中的记录
        --保留转出标记中的SQL查询所需的索引(主键和唯一键对应的索引) 
        For R In (Select a.Table_Name, a.Constraint_Name
                  From User_Constraints A, zlBakTables T, User_Tables B
                  Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And a.Status = 'ENABLED' And
                        a.Constraint_Type In ('P', 'U') And a.Table_Name = b.Table_Name And b.Iot_Type Is Null And
                        a.Constraint_Name Not In
                        (Select Upper(Column_Value) As Constraint_Name From Table(f_Str2list(v_Constraintswithtag)))
                  Order By Constraint_Name) Loop
          v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name ||
                   ' Cascade Drop Index';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    Else
      --启用时
      --1.先启用主键和唯一键，再启用引用转出表主键的他表外键 
      If Speedmode_In = 1 Then
        --先重建索引，再启用约束，以便重建索引时利用并行执行缩短时间，并且启用约束时也可以采用novalidate方式 
        For R In (Select d.Table_Name, d.Constraint_Name,
                         f_List2str(Cast(Collect(d.Column_Name Order By d.Position) As t_Strlist)) Colstr
                  From User_Cons_Columns D,
                       (Select a.Table_Name, a.Constraint_Name
                         From User_Constraints A, zlBakTables T
                         Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And a.Status = 'DISABLED' And
                               a.Constraint_Type In ('P', 'U')) A
                  Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name
                  Group By d.Table_Name, d.Constraint_Name
                  Order By Constraint_Name) Loop
          Update Zldatamovelog
          Set 当前进度 = '正在恢复约束:' || r.Constraint_Name
          Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
        
          Select Tablespace_Name Into v_Tbs From User_Indexes Where Table_Name = r.Table_Name And Rownum < 2;
        
          --禁用主键或唯一键时，索引是被删除了的，所以这里要用Create 
          v_Sql := 'Create Unique Index ' || r.Constraint_Name || ' On ' || r.Table_Name || '(' || r.Colstr ||
                   ') Tablespace ' || v_Tbs || ' Nologging';
          Begin
            Execute Immediate v_Sql;
          Exception
            When Others Then
              Null; --可能有些主键或唯一键不是本次转出期间被禁用的，之前就存在不唯一数据，创建唯一索引会出错 
          End;
        
          --会自动建立约束与索引的关联 
          v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.启用引用转出表主键的他表外键 
      For R In (Select c.Table_Name, c.Constraint_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'DISABLED'
                Order By a.Table_Name) Loop
        --为了加快速度，采用novalidate，不验证已有数据 
        --可能引用转出表主键的他表，在zlbaktables中定义了，但没有编写对应的数据转出脚本，未验证的数据可能有违反约束的情况。 
        v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.在线模式转出时，删除之前创建的用来替代级联删除外键的触发器
      If Speedmode_In = 0 Then
        For R In (Select a.Trigger_Name
                  From User_Triggers A, zlBakTables B
                  Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And
                        Trigger_Name = Table_Name || '_CASCADE_DEL' And Triggering_Event = 'DELETE') Loop
          v_Sql := 'Drop Trigger ' || r.Trigger_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    End If;
  End Setconstraintstatus;

  --功能：高速模式时禁用LOB以外的所有索引，在线模式时仅禁用转出表引用非转出表的外键索引(例如：病人医嘱计价_IX_收费细目ID) 
  --说明：禁用索引是为了提高删除数据的性能 
  Procedure Setindexstatus As
  Begin
    If Speedmode_In = 1 Then
      --保留转出标记中的SQL查询所需的索引 
      For R In (Select /*+ rule*/
                 a.Index_Name
                From User_Indexes A, zlBakTables T
                Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And t.直接转出 = 1 And
                      a.Index_Name <> a.Table_Name || '_IX_待转出' And
                      a.Index_Name Not In
                      (Select Upper(Column_Value) As Index_Name From Table(f_Str2list(v_Indexeswithtag))) And
                      a.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And a.Index_Type = 'NORMAL' And Not Exists
                 (Select 1
                       From User_Constraints C
                       Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
          Execute Immediate v_Sql;
        Else
          Update Zldatamovelog
          Set 当前进度 = '正在重建索引:' || r.Index_Name
          Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
          Begin
            Execute Immediate v_Sql;
            --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
          
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
                       Where c.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And c.Uniqueness = 'NONUNIQUE' And
                             c.Index_Type = 'NORMAL' And c.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And
                             c.Index_Name = d.Index_Name And c.Table_Name = d.Table_Name
                       Group By d.Table_Name, d.Index_Name) A,
                     (Select e.Table_Name,
                              f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr
                       From User_Cons_Columns E, User_Constraints F, zlBakTables T, User_Constraints C
                       Where e.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And
                             e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And
                             c.Constraint_Name = f.r_Constraint_Name And c.Table_Name Not In ('病案主页', '病人信息') And
                             Not Exists
                        (Select 1 From zlBakTables G Where g.表名 = c.Table_Name And g.系统 = System_In)
                       Group By e.Table_Name, e.Constraint_Name) B
                Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          --特殊处理：以下两个索引不禁用，是由于药品目录修改规格，财务缴款需要使用 
          If r.Index_Name Not In ('病人医嘱记录_IX_收费细目ID', '药品收发记录_IX_药品ID', '药品收发记录_IX_价格ID') Then
            v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
            Execute Immediate v_Sql;
          End If;
        Else
          Update Zldatamovelog
          Set 当前进度 = '正在重建索引:' || r.Index_Name
          Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
          Execute Immediate v_Sql;
          --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
        End If;
      End Loop;
    End If;
  End Setindexstatus;

  --功能：转出数据期间，停用转出表上的所有触发器，转出后再恢复 
  Procedure Settriggerstatus As
  Begin
    For R In (Select Distinct a.Table_Name, t.停用触发器
              From User_Triggers A, zlBakTables T
              Where a.Status = Decode(Enable_In, 0, 'ENABLED', 'DISABLED') And a.Table_Name = t.表名 And t.直接转出 = 1 And
                    t.系统 = System_In) Loop
      If Enable_In = 0 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' DISABLE ALL TRIGGERS';
        Update zlBakTables Set 停用触发器 = 1 Where 系统 = System_In And 表名 = r.Table_Name;
      Elsif Nvl(r.停用触发器, 0) = 1 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' ENABLE ALL TRIGGERS';
        Update zlBakTables Set 停用触发器 = Null Where 系统 = System_In And 表名 = r.Table_Name;
      End If;
      Execute Immediate v_Sql;
    End Loop;
    Commit;
  End Settriggerstatus;

  --功能：转出数据期间，停用当前所有者的所有自动作业，转出后再启用 
  Procedure Setjobstatus As
    v_Jobs Varchar2(4000);
  Begin
    --停用 
    If Enable_In = 0 Then
      For R In (Select Job From User_Jobs Where Broken = 'N') Loop
        Dbms_Job.Broken(r.Job, True);
        v_Jobs := v_Jobs || ',' || r.Job;
      End Loop;
    
      If v_Jobs Is Not Null Then
        v_Jobs := Substr(v_Jobs, 2);
        Update zlDataMove Set 停用作业号 = v_Jobs Where 系统 = System_In And 组号 = 1;
      End If;
    Else
      --启用 
      Select 停用作业号 Into v_Jobs From zlDataMove Where 系统 = System_In And 组号 = 1;
      If v_Jobs Is Not Null Then
        For R In (Select Job
                  From User_Jobs
                  Where Broken = 'Y' And Job In (Select Column_Value From Table(f_Num2list(v_Jobs)))) Loop
          Dbms_Job.Broken(r.Job, False);
        End Loop;
        Update zlDataMove Set 停用作业号 = Null Where 系统 = System_In And 组号 = 1;
      End If;
    End If;
    --作业设置后必须提交事务才生效 
    Commit;
  End Setjobstatus;
Begin
  If Parallel_In < 2 Then
    Execute Immediate 'Alter Session DISABLE PARALLEL DDL';
  Else
    If Func_In In (6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
      --为重建索引设置并行执行（由于通常受限于IO设备的性能，设置太高的并行度反而会降低性能，如有高性能存储设备，可加大并行度） 
      --执行重建索引后会自动为索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢),在后面取消索引的并行度 
      --恢复在线库的约束和索引时，不管是不是在线模式，都加上并行，否则太慢
      Execute Immediate 'Alter Session FORCE PARALLEL DDL PARALLEL ' || Parallel_In;
      n_Parallel := 1;
    End If;
  End If;

  If Func_In = 1 Then
    --1.设置触发器 
    Settriggerstatus;
  Elsif Func_In = 2 Then
    --2.设置自动作业 
    Setjobstatus;
  Elsif Func_In = 3 Then
    --3.设置约束状态 
    Setconstraintstatus;
  Elsif Func_In = 4 Then
    --4.设置索引状态 
    Setindexstatus;
  Elsif Func_In = 5 Then
    --5.重建"待转出"索引 
    For R In (Select b.Index_Name
              From zlBakTables A, User_Indexes B
              Where a.表名 = b.Table_Name And a.直接转出 = 1 And a.系统 = System_In And b.Index_Name = b.Table_Name || '_IX_待转出'
              Union All
              Select '病案主页_IX_待转出'
              From Dual
              Where System_In = 100) Loop
      Update Zldatamovelog
      Set 当前进度 = '正在重建索引:' || r.Index_Name
      Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
    
      --耗时太短，无须并行DDL 
      --在线转出时如果重建索引会锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
      --在线重建索引太慢，所以，即使在线转出模式也不用在线重建
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
  
  Elsif Func_In = 6 Then
    --6.重建标记转出查询所用到的索引（测试表明重建后最多可缩短一半的查询时间） 
    --根据业务的启用阶段来决定重建哪些索引，以避免一些不必要的重建耗时 
    For R In (Select b.Index_Name, a.组号
              From User_Indexes B, zlBakTables A
              Where a.系统 = System_In And a.表名 = b.Table_Name And
                    b.Index_Name In (Select Upper(Column_Value)
                                     From Table(f_Str2list(v_Indexeswithtag))
                                     Union
                                     Select Upper(Column_Value)
                                     From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      n_Do := 0;
      If Rebscope_In = 0 Then
        If r.组号 < 5 Then
          n_Do := 1; --仅经济核算类 
        End If;
      Elsif Rebscope_In = 1 Then
        If r.组号 < 5 Or r.组号 = 8 Then
          n_Do := 1; --仅经济核算类、医嘱类 
        End If;
      Else
        n_Do := 1;
      End If;
    
      If n_Do = 1 Then
        Update Zldatamovelog
        Set 当前进度 = '正在重建索引:' || r.Index_Name
        Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
      
        --v_Sql := 'Alter Index ' || r.Index_Name || ' shrink Space'; 
        --使用shrink方式不能并行执行,试验表明速度比rebuild PARALLEL 8 慢6倍 
        If Speedmode_In = 1 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
        Else
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
        End If;
        Begin
          Execute Immediate v_Sql;
          --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
        
        Exception
          When Others Then
            If SQLErrM Like 'ORA-00054%' Then
              v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
              Execute Immediate v_Sql;
            End If;
        End;
      End If;
    End Loop;
  
    --重组表的数据
  Elsif Func_In = 7 Then
    --rebScope_in=0,只重组组号小于5的经济核算类表（费用、药品、票据），否则全部重组 
    For R In (Select a.表名 As Table_Name
              From zlBakTables A
              Where a.直接转出 = 1 And (组号 < Decode(Rebscope_In, 0, 5, 100))
              Order By 组号, 序号) Loop
    
      Update Zldatamovelog
      Set 当前进度 = '正在重组表:' || r.Table_Name
      Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
    
      --如果有空闲的空间，最好移到其他表空间，只有这样才能绝对移动文件尾部的数据块，以便进行表空间文件的收缩 
      --在前面设置了会话级的强制并行 
      v_Sql := 'Alter Table ' || r.Table_Name || ' Move Nologging';
      Execute Immediate v_Sql;
    
      --单独移动Lob对象 
      For L In (Select Column_Name, Tablespace_Name From User_Lobs Where Table_Name = r.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Move Lob(' || l.Column_Name || ') Store as (Tablespace ' ||
                 l.Tablespace_Name || ') Nologging';
        Execute Immediate v_Sql;
      End Loop;
    
      v_Sql := 'Alter Table ' || r.Table_Name || ' Noparallel';
      Execute Immediate v_Sql;
    
      --move后，表相关的索引会全部失效，需要全部重建 
      For S In (Select Index_Name
                From User_Indexes
                Where Table_Name = r.Table_Name And Status = 'UNUSABLE'
                Order By Index_Name) Loop
        Update Zldatamovelog
        Set 当前进度 = '正在恢复失效索引:' || s.Index_Name
        Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
      
        --在前面设置了会话级的强制并行 
        v_Sql := 'Alter Index ' || s.Index_Name || ' Rebuild Nologging';
        Execute Immediate v_Sql;
      End Loop;
    End Loop;
    --重建转出表上标记转出以外的其他索引（用于转出完成后收回空闲空间）
    --失效的索引不重建，因为转出完后有单独的重建功能
  Elsif Func_In = 8 Then
    For R In (Select b.Index_Name, a.组号
              From User_Indexes B, zlBakTables A
              Where a.系统 = System_In And a.表名 = b.Table_Name And b.Status = 'VALID' And b.Index_Type = 'NORMAL' And
                    b.Index_Name Not Like 'BIN$%' And
                    b.Index_Name Not In (Select Upper(Column_Value)
                                         From Table(f_Str2list(v_Indexeswithtag))
                                         Union
                                         Select Upper(Column_Value)
                                         From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      Update Zldatamovelog
      Set 当前进度 = '正在重建索引:' || r.Index_Name
      Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
    
      If Speedmode_In = 1 Then
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Else
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
      End If;
      Begin
        Execute Immediate v_Sql;
        --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源    
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
  End If;

  --执行重建索引后会自动为索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢) 
  --------------------------------------------------------------------------------------------------- 
  If n_Parallel = 1 Then
    Execute Immediate 'ALTER Session DISABLE PARALLEL DDL';
  
    For R In (Select Index_Name From User_Indexes Where Degree Not In ('1', '0')) Loop
      v_Sql := 'Alter Index ' || r.Index_Name || ' Noparallel';
      Execute Immediate v_Sql;
    End Loop;
  End If;

  Update Zldatamovelog
  Set 当前进度 = '重建完成'
  Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
  Commit;
  --本过程不进行错误处理，错误由调用过程处理 
End Zl1_Datamove_Reb;
/


--83368:马政,2015-10-26,材料领用能通过申购单领用
Create Or Replace Procedure Zl_材料领用_Insert
(
  入出类别id_In In 药品收发记录.入出类别id%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  对方部门id_In In 药品收发记录.对方部门id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  批次_In       In 药品收发记录.批次%Type,
  填写数量_In   In 药品收发记录.填写数量%Type,
  成本价_In     In 药品收发记录.成本价%Type,
  成本金额_In   In 药品收发记录.成本金额%Type,
  零售价_In     In 药品收发记录.零售价%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  差价_In       In 药品收发记录.差价%Type,
  领用人_In     In 药品收发记录.领用人%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  产地_In       In 药品收发记录.产地%Type := Null,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  病人id_In     In 材料领用信息.病人id%Type := Null,
  使用时间_In   In 材料领用信息.使用时间%Type := Null,
  条码_In       In 材料领用信息.条码%Type := Null,
  申购数量_in     in 药品收发记录.单量%type :=null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  v_编码     收费项目目录.编码%Type;
  v_批准文号 药品库存.批准文号%Type;

  n_实价卫材     收费项目目录.是否变价%Type;
  n_入出系数     药品收发记录.入出系数%Type; --收发ID
  n_可用数量     药品库存.可用数量%Type;
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_收发id       药品收发记录.Id%Type;
  v_商品条码     药品库存.商品条码%Type;
  v_内部条码     药品库存.内部条码%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;

Begin
  n_入出系数 := -1;
  Begin
    Select 可用数量, Decode(上次供应商id, 0, Null, 上次供应商id), 上次生产日期, 批准文号, 商品条码, 内部条码
    Into n_可用数量, n_上次供应商id, d_上次生产日期, v_批准文号, v_商品条码, v_内部条码
    From 药品库存
    Where 药品id = 材料id_In And Nvl(批次, 0) = 批次_In And 库房id = 库房id_In And 性质 = 1 And Rownum = 1;
  Exception
    When Others Then
      n_可用数量     := 0;
      n_上次供应商id := Null;
      d_上次生产日期 := Null;
      v_批准文号     := Null;
  End;

  If 批次_In > 0 Then
    If n_可用数量 - 填写数量_In < 0 Then
      Select 编码 Into v_编码 From 收费项目目录 Where ID = 材料id_In;

      v_Err_Msg := '[ZLSOFT]编码为' || v_编码 || ',批号为' || 批号_In || '的分批核算材料' || Chr(10) || Chr(13) || '可用库存数量不够！[ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;
  Select 药品收发记录_Id.Nextval Into n_收发id From Dual;

  --插入类别为出的那一笔
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价,
     摘要, 填制人, 填制日期, 供药单位id, 生产日期, 批准文号, 领用人, 商品条码, 内部条码,单量)
  Values
    (n_收发id, 1, 20, No_In, 序号_In, 库房id_In, 对方部门id_In, 入出类别id_In, n_入出系数, 材料id_In, 批次_In, 产地_In, 批号_In, 效期_In, 灭菌效期_In,
     填写数量_In, 填写数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, n_上次供应商id, d_上次生产日期, v_批准文号,
     领用人_In, v_商品条码, v_内部条码,申购数量_in);

  --同时更新库存数
  Update 药品库存
  Set 可用数量 = Nvl(可用数量, 0) - 填写数量_In
  Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 性质 = 1;

  --不插入批次是因为批次材料不够，不准出库
  If Sql%NotFound Then
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;

    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
    Values
      (库房id_In, 材料id_In, Nvl(批次_In, 0), 1, -填写数量_In, 效期_In, 灭菌效期_In, n_上次供应商id, 成本价_In, 批号_In, d_上次生产日期, 产地_In, v_批准文号,
       Decode(n_实价卫材, 1, Decode(Nvl(批次_In, 0), 0, Null, 零售价_In), Null));
  End If;

  Delete From 药品库存
  Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;
  If Nvl(病人id_In, 0) <> 0 Then

    --处理病人信息
    Insert Into 材料领用信息
      (收发id, 材料id, 病人id, 主页id, 姓名, 性别, 年龄, 床号, 医疗付款方式, 当前科室id, 当前病区id, 使用时间, 条码)
      Select n_收发id, 材料id_In, 病人id_In, 住院次数, 姓名, 性别, 年龄, 当前床号, 医疗付款方式, 当前科室id, 当前病区id, 使用时间_In, 条码_In
      From 病人信息
      Where 病人id = 病人id_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料领用_Insert;
/

--89570:冉俊明,2015-10-26,医保病人门诊划价管理划价后收费没有填写保险编码。
Create Or Replace Procedure Zl_门诊划价记录_Update
(
  --功能：用于采用划价收费模式时,在收费界面输入划价单后,
  --        医保病人身份验证成功后,更新划价单中的病人身份信息,
  --        以及相关的保险项目、金额信息。
  险类_In   保险帐户.险类%Type,
  病人id_In 病人信息.病人id%Type,
  No_In     门诊费用记录.No%Type,
  病人_In   Number := 0 --是否只更新病人信息(用于医保病人信息更新)
) As
  --病人相关信息
  Cursor c_Patiinfo Is
    Select a.病人id, a.住院次数, a.门诊号, a.住院号, a.姓名, a.性别, a.年龄, a.当前病区id, a.当前科室id, a.医疗付款方式, Nvl(b.全额统筹, 0) As 全额统筹
    From 病人信息 A,
         (Select a.病人id, Nvl(b.全额统筹, 0) As 全额统筹
           From 保险帐户 A, 保险年龄段 B
           Where a.险类 = b.险类 And Nvl(a.中心, 0) = Nvl(b.中心, 0) And Nvl(a.在职, 0) = Nvl(b.在职, 0) And b.下限 <= Nvl(a.年龄段, 0) And
                 (a.年龄段 <= b.上限 Or b.上限 = 0) And a.病人id = 病人id_In And a.险类 = 险类_In) B
    Where a.病人id = b.病人id(+) And a.病人id = 病人id_In;
  r_Patiinfo c_Patiinfo%RowType;

  --划价单据原内容
  Cursor c_Bill Is
    Select ID, 收费细目id, 实收金额 From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In Order By 序号;

  v_保险项目否 门诊费用记录.保险项目否%Type;
  v_统筹金额   门诊费用记录.统筹金额%Type;
  v_保险大类id 门诊费用记录.保险大类id%Type;
  v_保险编码   门诊费用记录.保险编码%Type;

  v_收费细目id   保险支付项目.收费细目id%Type;
  v_项目是否医保 保险支付项目.是否医保%Type;
  v_大类是否医保 保险支付大类.是否医保%Type;
  v_算法         保险支付大类.算法%Type;
  v_统筹比额     保险支付大类.统筹比额%Type;
  v_服务对象     保险支付大类.服务对象%Type;
  v_医疗付款方式 门诊费用记录.付款方式%Type;
  v_Dec          Number;
Begin
  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;

  Open c_Patiinfo;
  Fetch c_Patiinfo
    Into r_Patiinfo;
  If c_Patiinfo%RowCount > 0 Then
    If Nvl(病人_In, 0) = 0 Then
      For r_Bill In c_Bill Loop
        --计算保险字段
        Begin
          Select a.收费细目id, a.大类id, Nvl(a.是否医保, 0), Nvl(b.是否医保, 0), Nvl(b.算法, 0), Nvl(b.统筹比额, 0), Nvl(b.服务对象, 3), a.项目编码
          Into v_收费细目id, v_保险大类id, v_项目是否医保, v_大类是否医保, v_算法, v_统筹比额, v_服务对象, v_保险编码
          From 保险支付项目 A, 保险支付大类 B
          Where a.大类id = b.Id(+) And b.险类(+) = 险类_In And a.险类 = 险类_In And a.收费细目id = r_Bill.收费细目id;
        Exception
          When Others Then
            v_收费细目id := Null;
            Null;
        End;
      
        If v_收费细目id Is Not Null Then
          v_保险项目否 := 1;
          If v_保险大类id Is Not Null Then
            If v_保险项目否 = 1 And v_大类是否医保 = 1 Then
              v_保险项目否 := 1;
            Else
              v_保险项目否 := 0;
            End If;
            If v_算法 = 1 Then
              v_统筹金额 := Round(r_Bill.实收金额 * v_统筹比额 / 100, v_Dec);
              If v_保险项目否 = 1 And v_统筹比额 > 0 Then
                v_保险项目否 := 1;
              Else
                v_保险项目否 := 0;
              End If;
            End If;
          
            --如果服务对象不对，也认为它不是医保项目
            If v_保险项目否 = 1 And v_服务对象 <> 1 And v_服务对象 <> 3 Then
              v_保险项目否 := 0;
            End If;
          End If;
        
          If v_保险项目否 = 1 And v_项目是否医保 = 1 Then
            v_保险项目否 := 1;
          Else
            v_保险项目否 := 0;
          End If;
        
          --全额统筹病人
          --                IF r_PatiInfo.全额统筹=1 And v_保险项目否=1 Then
          --                    v_统筹金额:=r_Bill.实收金额;
          --                End IF;
        Else
          v_保险项目否 := 0;
          v_保险大类id := Null;
          v_统筹金额   := 0;
        End If;
      
        --更新费用记录
        Update 门诊费用记录
        Set 保险项目否 = Nvl(v_保险项目否, 0), 保险大类id = Decode(保险大类id, Null, v_保险大类id, 保险大类id), 统筹金额 = Nvl(v_统筹金额, 0),
            保险编码 = Decode(保险编码, Null, v_保险编码, 保险编码)
        Where ID = r_Bill.Id;
      End Loop;
    Else
      --更新费用记录:因为未产生对应的汇总费用,所以可以更新科室等字段
      Begin
        Select 编码 Into v_医疗付款方式 From 医疗付款方式 Where 名称 = r_Patiinfo.医疗付款方式;
      Exception
        When Others Then
          v_医疗付款方式 := Null;
      End;
      Update 门诊费用记录
      Set 姓名 = r_Patiinfo.姓名, 性别 = r_Patiinfo.性别, 年龄 = r_Patiinfo.年龄, 病人id = r_Patiinfo.病人id,
          标识号 = Decode(门诊标志, 2, r_Patiinfo.住院号, r_Patiinfo.门诊号), 付款方式 = Decode(门诊标志, 2, Nvl(v_医疗付款方式, 付款方式), 付款方式),
          病人科室id = Decode(门诊标志, 2, r_Patiinfo.当前科室id, 病人科室id)
      Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In;
    End If;
  End If;
  Close c_Patiinfo;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊划价记录_Update;
/

--82768:王振涛,2015-10-22,分钟修改
CREATE OR REPLACE Function Zl_Get_Reference
(
  Type_In       In Number, --0=参考 1=参考ID 2=危急参考 3=危急参考下限 4=危急参考上限
  项目id_In     In Number,
  标本类型_In   In Varchar2,
  性别_In       In Number,
  出生日期_In   In Date,
  仪器id_In     In Number := Null,
  年龄_In       In Varchar2 := Null,
  申请科室id_In In Number := Null
) Return Varchar2 As

  Cursor V_Reference_Type Is
    Select A.Id,
           Trim(To_Char(A.参考低值, C.格式)) || '～' || Trim(To_Char(A.参考高值, C.格式)) ||
            Decode(A.临床特征, Null, '', '成人', '', '婴儿', '', ' ' || A.临床特征) As 结果参考, B.结果类型, B.取值序列,
           Trim(To_Char(A.警示下限, C.格式)) || '～' || Trim(To_Char(A.警示上限, C.格式)) ||
            Decode(A.临床特征, Null, '', '成人', '', '婴儿', '', ' ' || A.临床特征) As 危急参考, A.警示下限, A.警示上限, Nvl(B.多参考, 0) 多参考
    From 检验项目参考 A, 检验项目 B,
         (Select '9999990' ||
                   Decode(Max(Nvl(C.小数位数, -1)), 0, '', -1, '.00', Substr('.000000', 1, 1 + Max(Nvl(C.小数位数, -1)))) As 格式



           From 检验仪器项目 C, 检验项目 D
           Where D.诊治项目id = 项目id_In And D.诊治项目id = C.项目id(+)) C
    Where A.项目id = 项目id_In And A.项目id = B.诊治项目id;

  V_Return Varchar2(4000);
  V_Sql    Varchar2(4000);

  Type C_Type Is Ref Cursor; --声明REF游标类型
  R_Emp V_Reference_Type%Rowtype; --声明一个行类型变量
  Cur   C_Type; --声明REF游标类型的变量

  V_结果类型 Number(1);

  V_年数     Number(18);
  V_月数     Number(18);
  V_日数     Number(18);
  V_小时     Number(18);
  V_出生日期 Date;
  V_Pos      Number(4);
  V_多参考   Number(4);
  V_Value    Number(18);
  V_Valuerec Varchar2(255);
  V_年龄     Varchar2(50);
  V_结果参考 Varchar2(1000);
  V_参考id   Number(18);
  V_危紧参考 Varchar2(1000);
  V_警示下限 Varchar2(1000);
  V_警示上限 Varchar2(1000);
  Function Sub_Is_Number(V_In In Varchar2) Return Boolean Is
    N_Tmp Number;
  Begin
    N_Tmp := To_Number(V_In);
    If N_Tmp Is Not Null Then
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
    V_Str       In Varchar2,
    V_Delimiter In Varchar2,
    V_Number    In Number
  ) Return Varchar2 Is
    V_Record     Varchar2(1000);
    V_Currrecord Varchar2(1000);
    V_Currnum    Number;
  Begin
    V_Record  := V_Str || V_Delimiter;
    V_Currnum := 0;
    While V_Record Is Not Null Loop
      V_Currrecord := Substr(V_Record, 1, Instr(V_Record, V_Delimiter) - 1);
      If V_Currnum = V_Number Then
        Return(V_Currrecord);
      End If;

      V_Currnum := V_Currnum + 1;
      V_Record  := Replace(V_Delimiter || V_Record, V_Delimiter || V_Currrecord || V_Delimiter);
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

  V_Sql := ' Select a.id,Trim(To_Char(A.参考低值, C.格式)) || ''～'' || Trim(To_Char(A.参考高值, C.格式)) || ' ||
           ' Decode(A.临床特征, Null, '''', ''成人'', '''', ''婴儿'','''', '' '' || A.临床特征) As 结果参考, B.结果类型, B.取值序列, ' ||
           ' Trim(To_Char(A.警示下限, C.格式)) || ''～'' || Trim(To_Char(A.警示上限, C.格式)) || ' || ' Decode(A.临床特征, Null, '''', ''成人'', '''', ''婴儿'','''', '' '' || A.临床特征) As 危急参考,a.警示下限,a.警示上限,
             nvl(b.多参考,0) 多参考 ' || ' From 检验项目参考 A, 检验项目 B, ' || ' (Select ''9999990'' || ' ||
           ' Decode(Max(Nvl(C.小数位数, -1)), 0, '''', -1, ''.00'', Substr(''.000000'', 1, 1 + Max(Nvl(C.小数位数, -1)))) As 格式 ' ||
           ' From 检验仪器项目 C, 检验项目 D ' || ' Where D.诊治项目ID = ' || 项目id_In || ' And D.诊治项目ID = C.项目ID(+)) C ' ||
           ' Where A.项目ID = ' || 项目id_In || ' And A.项目ID = B.诊治项目ID ';

  V_年龄 := 年龄_In;
  If V_年龄 = '岁' Then
    V_年龄 := Null;
  End If;

  If V_年龄 = '月' Then
    V_年龄 := Null;
  End If;

  If V_年龄 = '小时' Then
    V_年龄 := Null;
  End If;

  If V_年龄 = '天' Then
    V_年龄 := Null;
  End If;

  If Nvl(标本类型_In, '') <> '' Or 标本类型_In Is Not Null Then
    V_Sql := V_Sql || ' And A.标本类型 = ' || Chr(39) || 标本类型_In || Chr(39);
  End If;

  If Nvl(性别_In, '') <> '' Or 性别_In Is Not Null Then
    --V_Sql := V_Sql || ' And A.性别域 = Nvl(' || 性别_In || ', 1) ';
    V_Sql := V_Sql || ' And decode(A.性别域,null,' || 性别_In || ',0,' || 性别_In || ',A.性别域) = Nvl(' || 性别_In || ', 1) ';
  End If;

  If Nvl(仪器id_In, '') <> '' Or 仪器id_In Is Not Null Then
    V_Sql := V_Sql || ' And (A.仪器id = ' || 仪器id_In || ' Or A.仪器id Is Null) ';
  End If;

  If Nvl(V_年龄, '') <> '' Or V_年龄 Is Not Null Then
    If Instr(V_年龄, '岁') > 0 Or Instr(V_年龄, '月') > 0 Or Instr(V_年龄, '天') > 0 Or Instr(V_年龄, '小时') > 0 Or
       Sub_Is_Number(V_年龄) Then
      --处理日期
      V_出生日期 := 出生日期_In;
      V_年龄     := V_年龄;
      If Instr(V_年龄, '岁') > 0 Then
        V_年龄 := Substr(V_年龄, 1, Instr(V_年龄, '岁'));
      Elsif Instr(V_年龄, '月') > 0 Then
        V_年龄 := Substr(V_年龄, 1, Instr(V_年龄, '月'));
      Elsif Instr(V_年龄, '小时') > 0 Then
        V_年龄 := Substr(V_年龄, 1, Instr(V_年龄, '小时') + 1);
        if  V_年龄 = '0小时' or  V_年龄 = '0时' then
            V_年龄 :=' ';
        end if ;
      End If;
      If V_年龄 Is Not Null And (V_年龄 = '成人' Or V_年龄 = '婴儿' Or V_年龄 = '岁') = False Then
        If Substr(V_年龄, 1, 1) = '*' Then
          V_出生日期 := Add_Months(Sysdate, -216);
        Else
          If Substr(V_年龄, Length(V_年龄)) = '月' Then
            V_出生日期 := Add_Months(Sysdate, -1 * Nvl(Zlval(V_年龄), 0));
          Else
            If Substr(V_年龄, Length(V_年龄)) = '天' Then
              V_出生日期 := Sysdate - Nvl(Zlval(V_年龄), 0);
            Else
              If Substr(V_年龄, Length(V_年龄) - 1) = '小时' Then
                If Nvl(Zlval(V_年龄), 0) <> 0 Then
                  V_出生日期 := Sysdate - Nvl(Zlval(V_年龄), 0) / 24;
                End If;
              Else
                V_出生日期 := Add_Months(Sysdate, -12 * Nvl(Zlval(V_年龄), 0)) - 1;
              End If;
            End If;
          End If;
        End If;
      End If;
      If Not (V_出生日期 Is Null) Then
        Select Round(Months_Between(Sysdate, V_出生日期) / 12 - 0.5) Into V_年数 From Dual;
        Select Round(Months_Between(Sysdate, V_出生日期) - 0.5) Into V_月数 From Dual;
        Select Round(Sysdate - V_出生日期 - 0.5) Into V_日数 From Dual;
        Select Round((Sysdate - (V_出生日期 - 1 / 24)) * 24 - 1) Into V_小时 From Dual;
      End If;

      V_Sql := V_Sql || 'And (Decode(A.年龄单位, ''日'',' || V_日数 || ', ''月'',' || V_月数 || ',''小时'',' || V_小时 || ',' || V_年数 || ') ' ||
               ' Between Nvl(A.年龄下限, -9999) And Nvl(A.年龄上限, 9999) )';

    End If;

  End If;
  If Instr(V_年龄, '成人') > 0 Or Instr(V_年龄, '婴儿') > 0 Then
    --处理成人和婴儿
    V_Sql := V_Sql || ' And A.临床特征 =' || Chr(39) || V_年龄 || Chr(39);
  Else
    V_Sql := V_Sql || ' And instr(''婴儿,成人'',nvl(临床特征,'' '')) <= 0  ';
  End If;

  If Nvl(申请科室id_In, '') <> '' Or 申请科室id_In Is Not Null Then
    V_Sql := V_Sql || ' And (A.申请科室ID = ' || 申请科室id_In || ' Or nvl(A.申请科室ID,0) = 0) ';
  End If;

  If (Nvl(V_年龄, '') = '' Or V_年龄 Is Null) And (出生日期_In <> '' Or 出生日期_In Is Not Null) Then
    --按出生日期查询
    If Not (出生日期_In Is Null) Then
      Select Round(Months_Between(Sysdate, 出生日期_In) / 12 - 0.5) Into V_年数 From Dual;
      Select Round(Months_Between(Sysdate, 出生日期_In) - 0.5) Into V_月数 From Dual;
      Select Round(Sysdate - 出生日期_In - 0.5) Into V_日数 From Dual;
      Select Round((Sysdate - (出生日期_In - 1 / 24)) * 24 - 1) Into V_小时 From Dual;

      V_Sql := V_Sql || 'And (Decode(A.年龄单位, ''日'',' || V_日数 || ', ''月'',' || V_月数 || ',''小时'',' || V_小时 || ',' || V_年数 || ') ' ||
               ' Between Nvl(A.年龄下限, -9999) And Nvl(A.年龄上限, 9999) )';

    End If;
  End If;

  --加上排序
  V_Sql    := V_Sql || ' Order By a.默认 desc,A.临床特征, A.性别域 desc,a.id';
  V_Return := '';
  Open Cur For V_Sql;

  Loop
    Fetch Cur
      Into R_Emp;
    Exit When Cur%NotFound;
    If Cur%Rowcount > 0 Then

      V_结果类型 := R_Emp.结果类型;
      V_Valuerec := R_Emp.取值序列;
      V_参考id   := R_Emp.Id;
      V_多参考   := R_Emp.多参考;

      If Nvl(V_Return, '') = '' Or V_Return Is Null Then
        If Type_In = 2 Then
          V_Return := R_Emp.危急参考;
        Else
          V_Return := R_Emp.结果参考;
        End If;
      Else
        If Type_In = 2 Then
          V_Return := V_Return || Chr(13) || Chr(10) || R_Emp.危急参考;
        Else
          If V_多参考 = 1 Then
            V_Return := V_Return || Chr(13) || Chr(10) || R_Emp.结果参考;
          End If;
        End If;
      End If;

      --只增加第一个选出的警示参考
      If V_警示下限 = '' Or V_警示下限 Is Null Then
        V_警示下限 := R_Emp.警示下限;
      End If;
      If V_警示上限 = '' Or V_警示上限 Is Null Then
        V_警示上限 := R_Emp.警示上限;
      End If;
    End If;
  End Loop;

  If V_Return = '' Or V_Return Is Null Then
    Begin
      Select 结果参考, 结果类型, 取值序列, ID, 危急参考, 警示下限, 警示上限
      Into V_结果参考, V_结果类型, V_Valuerec, V_参考id, V_危紧参考, V_警示下限, V_警示上限
      From (Select A.Id,
                    Trim(To_Char(A.参考低值, C.格式)) || '～' || Trim(To_Char(A.参考高值, C.格式)) ||
                     Decode(A.临床特征, Null, '', '成人', '', '婴儿', '', ' ' || A.临床特征) As 结果参考, B.结果类型, B.取值序列,
                    Trim(To_Char(A.警示下限, C.格式)) || '～' || Trim(To_Char(A.警示上限, C.格式)) ||
                     Decode(A.临床特征, Null, '', '成人', '', '婴儿', '', ' ' || A.临床特征) As 危急参考, A.警示下限, A.警示上限
             From 检验项目参考 A, 检验项目 B,
                  (Select '9999990' ||
                            Decode(Max(Nvl(C.小数位数, -1)), 0, '', -1, '.00', Substr('.000000', 1, 1 + Max(Nvl(C.小数位数, -1)))) As 格式
                    From 检验仪器项目 C, 检验项目 D
                    Where D.诊治项目id = 项目id_In And D.诊治项目id = C.项目id(+)) C
             Where A.项目id = 项目id_In And A.项目id = B.诊治项目id
             Order By A.默认 Desc, A.临床特征, A.性别域)
      Where Rownum = 1;
      If Type_In = 2 Then
        V_Return := V_危紧参考;
      Else
        V_Return := V_结果参考;
      End If;
      --只增加第一个选出的警示参考
      If V_警示下限 = '' Or V_警示下限 Is Null Then
        V_警示下限 := R_Emp.警示下限;
      End If;
      If V_警示上限 = '' Or V_警示上限 Is Null Then
        V_警示上限 := R_Emp.警示上限;
      End If;
    Exception
      When Others Then
        V_Return := Null;
    End;
  End If;
  If V_Return <> '' Or V_Return Is Not Null Then

    If V_Return = '～' Then
      V_Return := '';
    Else
      If V_结果类型 = 2 Then
        V_Pos := Instr(V_Return, '～');

        Begin
          Select To_Number(Substr(V_Return, 1, V_Pos - 1)) Into V_Value From Dual;
        Exception
          When Others Then
            V_Value := 0;
        End;
        V_Return := Zlsplit(V_Valuerec, ';', V_Value);
      End If;
    End If;
    If Type_In = 0 Then
      Return V_Return;
    Elsif Type_In = 1 Then
      Return V_参考id;
    Elsif Type_In = 2 Then
      Return V_Return;
    Elsif Type_In = 3 Then
      Return V_警示下限;
    Elsif Type_In = 4 Then
      Return V_警示上限;
    End If;
  End If;
  Close Cur; --关闭游标
  Return V_Return;
End Zl_Get_Reference;
/

--89640:刘尔旋,2015-10-21,自动费用计算问题
Create or Replace View 病人自动费用 as
Select p.病人id, p.主页id, i.姓名, i.性别, i.年龄, i.住院号, a.费别, p.科室id, p.病区id, p.床号, p.附加床位, p.收费细目id, p.收入项目id, 1 As 标志,
       p.现价 As 标准单价, p.开始日期, p.终止日期, p.终止日期 - p.开始日期 As 天数, p.数量, p.经治医师, p.责任护士, p.操作员编号, p.操作员姓名
From 病人信息 I, 病案主页 A,
     (Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量
       From 自动计价项目 A,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 床位等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间
              From 病人变动记录
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名,
                     上次计算时间
              From 病人变动记录 B, 收费从属项目 I
              Where b.床位等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B, 收费价目 P
       Where a.病区id = b.病区id And Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.计算标志 = 1 And b.床位等级id = p.收费细目id And Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))
       Union All
       Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量
       From 自动计价项目 A,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 护理等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间
              From 病人变动记录
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 护理等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名,
                     上次计算时间
              From 病人变动记录 B, 收费从属项目 I
              Where b.护理等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B, 收费价目 P, 收费项目目录 C
       Where a.病区id = b.病区id And b.附加床位 <> 1 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And a.计算标志 = 2 And
             b.护理等级id = p.收费细目id And b.护理等级id = c.Id And Nvl(c.计算方式, 0) <> 1 And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))
       Union All
       Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, A.数量
       From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期
              From 自动计价项目
              Union All
              Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期
              From 自动计价项目 A, 收费从属项目 I
              Where a.收费细目id = i.主项id And i.固有从属 > 0) A, 病人变动记录 B, 收费价目 P
       Where a.病区id = b.病区id And b.附加床位 <> 1 And b.开始原因 <> 10 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.收费细目id = p.收费细目id And (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志 = 7) And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))) P
Where i.病人id = p.病人id And a.病人id = p.病人id And a.主页id = p.主页id
/

--89640:刘尔旋,2015-10-23,自动记帐费用计算问题
Create Or Replace Procedure Zl1_Autocptone
(
  病人id_In In Number,
  主页id_In In Number,
  期间_In   In Varchar2
) As

  -------------------------------------------------------------------------
  --功能说明：完成指定病人指定期间自动计价项目表设置自动计算的项目进行记帐处理
  --          1、系统首先根据系统参数"修正上期自动计费"，修改以往该病人自动记帐记录标志;
  --          2、综合病人的床位变化、入出转情况、调价情况等多项因素，结合期间跨度、病人费
  --             别等完成费用的正确计算：
  --             如果发现已经计算，则修改标志为正常;如果未计算，则插入新的自动记帐记录;
  --             作废以前的错误计算的记录;
  --             统计本次变动(新增和作废)，填写余额表和汇总表;
  --入口参数：
  --       病人ID_IN  number    病人身份ID
  --       主页ID_IN  number    病案主页ID，两个参数共同确定需要计算的病人
  --       期间_IN  varchar2     需要计算的最小期间
  --调用关系：zl1_AutoCptPati/zl1_AutoCptWard/zl1_AutoCptAll 调用本过程

  Cursor v_Autocur
  (
    期间_In Varchar2,
    Insure  病案主页.险类%Type
  ) Is
    Select l.病人id, l.主页id, l.姓名, l.性别, l.年龄, l.住院号, l.费别, l.科室id, l.病区id, l.床号, l.附加床位, l.收费细目id, l.收入项目id, l.标志, l.标准单价,
           Greatest(l.开始日期, Trunc(p.开始日期)) As 开始日期, l.终止日期, l.天数, l.数量, l.经治医师, l.责任护士, l.操作员编号, l.操作员姓名, i.险类, i.大类id,
           k.算法, k.统筹比额
    From (Select * From 病人自动费用 Where 病人id = 病人id_In And 主页id = 主页id_In) L,
         (Select Min(开始日期) As 开始日期 From 期间表 Where 期间 >= 期间_In) P, 保险支付项目 I, 保险支付大类 K
    Where Trunc(l.终止日期) >= Trunc(p.开始日期) And l.收费细目id = i.收费细目id(+) And i.险类(+) = Insure And i.大类id = k.Id(+)
    Order By l.开始日期;

  n_Insure       病案主页.险类%Type;
  v_Billno       Varchar2(8); --费用表实际的自动记帐号码
  n_Datecount    Integer; --日期计数器
  d_Datefrom     Date; --开始计算日期
  d_Dateto       Date; --终止计算日期
  d_Datelast     Date;
  n_Billcount    Number(5) := 0; --单据序号计数器
  n_Exsetax      Number(16, 2) := 0; --费用收取比率
  n_Exsetax_Temp Number(16, 2) := 0; --费用收取比率
  n_Summoney     Number(16, 2) := 0; --金额

  Cursor v_Sumcur
  (
    Billno    Varchar2,
    Datestart Date
  ) Is
    Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(Decode(附加标志, 0, 1, -1) * 应收金额) As 应收金额,
           Sum(Decode(附加标志, 0, 1, -1) * 实收金额) As 实收金额
    From 住院费用记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And
          (NO = Billno Or 附加标志 = 5 And 发生时间 >= Datestart)
    Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id;

  n_Dec            Number; --金额小数位数
  d_登记时间       Date; --登记时间
  d_发生时间       Date; --发生时间
  n_Dates          Number(3, 1); --当前记录的天数，全天为1
  n_Do             Number(1);
  n_返回值         病人余额.预交余额%Type;
  n_Delete         Number;
  n_医疗小组id     住院费用记录.医疗小组id%Type;
  n_护理计算标准   Number(2); --护理费计算标准
  n_收费细目id     Number(18);
  n_Temp           Number(18);
  l_护理id         t_Numlist := t_Numlist();
  l_护理等级       t_Numlist := t_Numlist();
  n_护理项目       Number(2); --1:是护理项目;0-非不护理
  n_价格           收费价目.现价%Type;
  n_护理已处理     Number(2); --1-护理费已经处理,;0-未处理
  n_收入项目id     Number(18);
  n_从属项目       Number(2);
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);

  n_病人病区id 住院费用记录.病人病区id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;

  --已经计算了的护理类型
  Type t_护理_Rec Is Record(
    收费细目id 收费项目目录.Id%Type,
    日期       Date);
  Type t_护理 Is Table Of t_护理_Rec;
  c_护理 t_护理 := t_护理();

Begin
  Begin
    Select 险类, Nvl(审核标志, 0), Nvl(状态, 0)
    Into n_Insure, n_审核标志, n_住院状态
    From 病案主页
    Where 病人id = 病人id_In And 主页id = 主页id_In;
  Exception
    When Others Then
      Return;
  End;

  n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
  n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
  If n_病人审核方式 = 1 And Nvl(n_审核标志, 0) >= 1 Then
    Return;
  End If;
  If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then
    Return;
  End If;

  v_Billno := Nextno(17);

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(160), '0'))
  Into n_Dec, n_护理计算标准
  From Dual;

  --每天5点以前，将记录时间登记为昨天，否则登记为当时
  Select Decode(Sign(To_Number(To_Char(Sysdate, 'HH24')) - 5), -1, Trunc(Sysdate) - 1 / 24 / 60, Sysdate)
  Into d_登记时间
  From Dual;

  --锁定该病人的记录,以免重复计算
  Update 病案主页 Set 状态 = 状态 Where 病人id = 病人id_In And 主页id = 主页id_In;

  -----------------------------------------------------------------
  d_Datefrom := Sysdate + 1000;
  d_Dateto   := Sysdate - 1000;
  n_Do       := 0;
  --------------------------------------------------------------------
  If n_护理计算标准 = 1 Then
    --同天以最高价位的护理费为准,先将其护理等级记住,
    For v_护理 In (Select Distinct 护理等级id
                 From (Select 护理等级id
                        From 病人变动记录
                        Where 开始原因 <> 10 And 病人id = 病人id_In And 主页id = 主页id_In
                        Union All
                        Select i.从项id As 护理等级id
                        From 病人变动记录 B, 收费从属项目 I
                        Where b.护理等级id = i.主项id And 病人id = 病人id_In And 主页id = 主页id_In And b.开始原因 <> 10 And i.固有从属 > 0)) Loop
      If Nvl(v_护理.护理等级id, 0) <> 0 Then
        l_护理id.Extend;
        l_护理id(l_护理id.Count) := v_护理.护理等级id;
      End If;
    End Loop;
  End If;
  -----------------------------------------------------------------
  --循环检查计算情况，并增加正确和新计算的记录
  -----------------------------------------------------------------
  For v_Currrow In v_Autocur(期间_In, n_Insure) Loop
  
    n_医疗小组id := Zl_医疗小组_Get(v_Currrow.科室id, v_Currrow.操作员姓名, v_Currrow.病人id, v_Currrow.主页id, d_发生时间);
  
    If d_Datefrom > v_Currrow.开始日期 Then
      d_Datefrom := v_Currrow.开始日期;
      n_Do       := 1;
      --将本次开始计算时间以后的已计算记录标志修改
      Update 住院费用记录
      Set 附加标志 = 5
      Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 <> 8 And Nvl(医嘱序号, 0) = 0 And
            发生时间 >= v_Currrow.开始日期;
    End If;
  
    If d_Dateto < v_Currrow.终止日期 Then
      d_Dateto := v_Currrow.终止日期;
    End If;
    n_收费细目id := v_Currrow.收费细目id;
    n_护理项目   := 0;
    --护理费计算标准:0-按最后一次护理计算;1-按价格最高的护理等级计算。
    If n_护理计算标准 = 1 Then
      --先确定是否护理项目,如果是,则需要重新进行计算
      Select Count(*) Into n_护理项目 From Table(l_护理id) Where Column_Value = n_收费细目id;
    End If;
  
    --提取当前收入项目的收费比率
    Begin
      Select 实收比率
      Into n_Exsetax
      From (Select 实收比率
             From 费别明细
             Where 费别 = v_Currrow.费别 And 收费细目id = v_Currrow.收费细目id And
                   (Abs(v_Currrow.标准单价 * v_Currrow.数量) Between 应收段首值 And 应收段尾值)
             Union All
             Select 实收比率
             From 费别明细
             Where 费别 = v_Currrow.费别 And 收入项目id = v_Currrow.收入项目id And
                   (Abs(v_Currrow.标准单价 * v_Currrow.数量) Between 应收段首值 And 应收段尾值) And Not Exists
              (Select 1 From 费别明细 Where 费别 = v_Currrow.费别 And 收费细目id = v_Currrow.收费细目id));
    Exception
      When Others Then
        n_Exsetax := 100.00;
    End;
  
    n_Exsetax := Nvl(n_Exsetax, 100);
    For n_Datecount In 0 .. (Trunc(v_Currrow.终止日期 + 0.5) - Trunc(v_Currrow.开始日期)) - 1 Loop
      d_发生时间   := Greatest(v_Currrow.开始日期, Trunc(v_Currrow.开始日期 + n_Datecount));
      n_Dates      := Least(Trunc(v_Currrow.开始日期 + n_Datecount + 1), v_Currrow.终止日期) -
                      Greatest(v_Currrow.开始日期, Trunc(v_Currrow.开始日期 + n_Datecount));
      n_护理已处理 := 0;
      If n_护理项目 = 1 Then
        --1.先检查当天是否存在护理变动,只有存在多个护理变动的,才会去处理(以主项目为准)
        n_从属项目 := 1;
        If l_护理等级.Count > 0 Then
          l_护理等级.Delete;
        End If;
        For v_护理 In (Select Distinct 护理等级id
                     From 病人变动记录
                     Where 开始原因 <> 10 And 病人id = 病人id_In And 主页id = 主页id_In And
                           (Trunc(开始时间) = Trunc(d_发生时间) Or Trunc(Nvl(终止时间, Sysdate)) = Trunc(d_发生时间))) Loop
          If Nvl(v_护理.护理等级id, 0) <> 0 Then
            l_护理等级.Extend;
            l_护理等级(l_护理等级.Count) := v_护理.护理等级id;
            If Nvl(v_护理.护理等级id, 0) = Nvl(v_Currrow.收费细目id, 0) Then
              n_从属项目 := 0;
            End If;
          End If;
        End Loop;
        If l_护理等级.Count > 1 Then
          --2. 存在两个以上变动,则取价位最高的
          n_Temp       := v_Currrow.收费细目id;
          n_价格       := Nvl(v_Currrow.标准单价, 0);
          n_收入项目id := v_Currrow.收入项目id;
          --本身是从属项目时,由于主项目计算时,已经计算了的,所以就不再计算
          If Nvl(n_从属项目, 0) = 1 Then
            n_护理已处理 := 1;
          End If;
          --因为可能存在多个收入项目,但收费细目相同的情况,因此,必须先检查该项目是否已经参与计算过的
          For I In 1 .. c_护理.Count Loop
            If c_护理(I).收费细目id = v_Currrow.收费细目id And c_护理(I).日期 = Trunc(d_发生时间) Then
              n_护理已处理 := 1;
              Exit;
            End If;
          End Loop;
          If Nvl(n_护理已处理, 0) = 0 Then
            c_护理.Extend;
            c_护理(c_护理.Count).收费细目id := v_Currrow.收费细目id;
            c_护理(c_护理.Count).日期 := Trunc(d_发生时间);
          End If;
          If Nvl(n_从属项目, 0) = 0 And Nvl(n_护理已处理, 0) = 0 Then
            --3.处理最高价位
            For v_价位 In (Select /*+ rule */
                          a.Column_Value As 收费细目id, p.现价, p.收入项目id
                         From Table(l_护理等级) A, 收费价目 P, 收费项目目录 C
                         Where a.Column_Value = p.收费细目id And a.Column_Value = c.Id And d_发生时间 Between p.执行日期 And
                               Nvl(p.终止日期, Sysdate) And Nvl(c.计算方式, 0) <> 1) Loop
              If Nvl(v_价位.现价, 0) > n_价格 Then
                n_价格       := Nvl(v_价位.现价, 0);
                n_Temp       := v_价位.收费细目id;
                n_收入项目id := v_价位.收入项目id;
              End If;
            End Loop;
          
            If n_Temp <> v_Currrow.收费细目id And Nvl(n_护理已处理, 0) = 0 Then
            
              n_开单部门id := v_Currrow.科室id;
              n_病人病区id := v_Currrow.病区id;
            
              For c_变动记录 In (Select 病区id, 科室id
                             From 病人变动记录
                             Where 开始原因 <> 10 And 病人id = 病人id_In And 主页id = 主页id_In And 护理等级id + 0 = n_Temp And
                                   (Trunc(开始时间) = Trunc(d_发生时间) Or Trunc(Nvl(终止时间, Sysdate)) = Trunc(d_发生时间))
                             Order By 开始时间 Desc) Loop
                n_开单部门id := c_变动记录.科室id;
                n_病人病区id := c_变动记录.病区id;
                Exit;
              End Loop;
            
              --4. 不等的话,需要重新处理相关费用
              For v_费用 In (Select n_Temp As 收费细目id, v_Currrow.数量 As 数量, n_价格 As 单价, n_收入项目id As 收入项目id
                           From Dual
                           Union All
                           Select 从项id As 收费细目id, a.从项数次 As 数量, p.现价 As 单价, p.收入项目id
                           From 收费从属项目 A, 收费价目 P, 收费项目目录 C
                           Where a.从项id = p.收费细目id And a.从项id = c.Id And Nvl(c.计算方式, 0) <> 1 And a.主项id = n_Temp And
                                 d_发生时间 Between p.执行日期 And Nvl(p.终止日期, Sysdate)) Loop
                --确定比例
                Begin
                  Select 实收比率
                  Into n_Exsetax_Temp
                  From (Select 实收比率
                         From 费别明细
                         Where 费别 = v_Currrow.费别 And 收费细目id = v_费用.收费细目id And
                               (Abs(v_费用.单价 * v_费用.数量) Between 应收段首值 And 应收段尾值)
                         Union All
                         Select 实收比率
                         From 费别明细
                         Where 费别 = v_Currrow.费别 And 收入项目id = v_费用.收入项目id And
                               (Abs(v_费用.单价 * v_费用.数量) Between 应收段首值 And 应收段尾值) And Not Exists
                          (Select 1 From 费别明细 Where 费别 = v_Currrow.费别 And 收费细目id = v_费用.收费细目id));
                Exception
                  When Others Then
                    n_Exsetax_Temp := 100.00;
                End;
                n_Exsetax_Temp := Nvl(n_Exsetax_Temp, 100);
                --如果已经计算，原记录计算完全正确，则直接修改将标志改正
                Update 住院费用记录
                Set 附加标志 = 0
                Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And Nvl(加班标志, 0) = v_Currrow.附加床位 And
                      病人科室id = Nvl(n_开单部门id, 0) And 病人病区id = Nvl(n_病人病区id, 0) And Nvl(床号, 0) = Nvl(v_Currrow.床号, 0) And
                      收费细目id = v_费用.收费细目id And 收入项目id = v_费用.收入项目id And 发生时间 = d_发生时间 And 数次 = v_费用.数量 * n_Dates And
                      标准单价 = v_费用.单价 And 应收金额 = Round(v_费用.单价 * v_费用.数量 * n_Dates, n_Dec) And
                      实收金额 = Round(v_费用.单价 * v_费用.数量 * n_Dates * n_Exsetax / 100, n_Dec);
              
                If Sql%RowCount = 0 Then
                  --如果未计算或计算错误，则增加正确的计算记录
                  Insert Into 住院费用记录
                    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id,
                     姓名, 性别, 年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志,
                     收据费目, 开单人, 划价人, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id)
                    Select 病人费用记录_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null,
                           Decode(v_Currrow.主页id, Null, 1, 2), v_Currrow.病人id, v_Currrow.主页id, n_病人病区id, n_开单部门id,
                           n_开单部门id, n_病人病区id, v_Currrow.姓名, v_Currrow.性别, v_Currrow.年龄, v_Currrow.住院号, v_Currrow.床号,
                           v_Currrow.费别, 1, v_费用.收费细目id, v_费用.收入项目id, 0, v_费用.单价, 1, v_费用.数量 * n_Dates,
                           Round(v_费用.单价 * v_费用.数量 * n_Dates, n_Dec),
                           Round(v_费用.单价 * v_费用.数量 * n_Dates * n_Exsetax / 100, n_Dec), i.类别, i.计算单位, v_Currrow.附加床位,
                           j.收据费目, v_Currrow.经治医师, v_Currrow.责任护士, v_Currrow.操作员编号, v_Currrow.操作员姓名, d_发生时间, d_登记时间,
                           Decode(v_Currrow.险类, Null, 0, 1), v_Currrow.大类id,
                           Decode(v_Currrow.算法, 1,
                                   Round(v_费用.单价 * v_费用.数量 * n_Dates * n_Exsetax / 100 * v_Currrow.统筹比额 / 100, n_Dec), 2,
                                   v_Currrow.统筹比额, 0), n_医疗小组id
                    From (Select 类别, 计算单位
                           From 收费细目
                           Where ID = v_费用.收费细目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) I,
                         (Select 收据费目
                           From 收入项目
                           Where ID = v_费用.收入项目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) J;
                  n_Billcount := n_Billcount + Sql%RowCount;
                End If;
                n_护理已处理 := 1;
              End Loop;
            End If;
          End If;
        End If;
      End If;
    
      If Nvl(n_护理已处理, 0) = 0 Then
        --如果已经计算，原记录计算完全正确，则直接修改将标志改正
        Update 住院费用记录
        Set 附加标志 = 0
        Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And Nvl(加班标志, 0) = v_Currrow.附加床位 And
              病人科室id = v_Currrow.科室id And 病人病区id = Nvl(v_Currrow.病区id, 0) And Nvl(床号, 0) = Nvl(v_Currrow.床号, 0) And
              收费细目id = v_Currrow.收费细目id And 收入项目id = v_Currrow.收入项目id And 发生时间 = d_发生时间 And 数次 = v_Currrow.数量 * n_Dates And
              标准单价 = v_Currrow.标准单价 And 应收金额 = Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates, n_Dec) And
              实收金额 = Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates * n_Exsetax / 100, n_Dec);
      
        If Sql%RowCount = 0 Then
          --如果未计算或计算错误，则增加正确的计算记录\
          Insert Into 住院费用记录
            (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 姓名, 性别,
             年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 划价人,
             操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id)
            Select 病人费用记录_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null,
                   Decode(v_Currrow.主页id, Null, 1, 2), v_Currrow.病人id, v_Currrow.主页id, v_Currrow.病区id, v_Currrow.科室id,
                   v_Currrow.科室id, v_Currrow.病区id, v_Currrow.姓名, v_Currrow.性别, v_Currrow.年龄, v_Currrow.住院号, v_Currrow.床号,
                   v_Currrow.费别, 1, v_Currrow.收费细目id, v_Currrow.收入项目id, 0, v_Currrow.标准单价, 1, v_Currrow.数量 * n_Dates,
                   Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates, n_Dec),
                   Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates * n_Exsetax / 100, n_Dec), i.类别, i.计算单位, v_Currrow.附加床位,
                   j.收据费目, v_Currrow.经治医师, v_Currrow.责任护士, v_Currrow.操作员编号, v_Currrow.操作员姓名, d_发生时间, d_登记时间,
                   Decode(v_Currrow.险类, Null, 0, 1), v_Currrow.大类id,
                   Decode(v_Currrow.算法, 1,
                           Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates * n_Exsetax / 100 * v_Currrow.统筹比额 / 100, n_Dec),
                           2, v_Currrow.统筹比额, 0), n_医疗小组id
            From (Select 类别, 计算单位
                   From 收费细目
                   Where ID = v_Currrow.收费细目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) I,
                 (Select 收据费目
                   From 收入项目
                   Where ID = v_Currrow.收入项目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) J;
        
          n_Billcount := n_Billcount + Sql%RowCount;
        End If;
      End If;
    End Loop;
  End Loop;

  If n_Do = 0 Then
    --撤销出院后,如果修改出院时间为入院当天则不产生新费用,但以前的费用要冲销
    Begin
      Select Trunc(b.上次计算时间)
      Into d_Datelast
      From 病人变动记录 A, 病人变动记录 B
      Where a.病人id = 病人id_In And a.主页id = 主页id_In And a.终止原因 = 1 And a.病人id = b.病人id And a.主页id = b.主页id And b.开始原因 = 1 And
            Trunc(b.开始时间) = Trunc(a.终止时间) And a.附加床位 = 0 And b.附加床位 = 0;
    Exception
      When Others Then
        Null;
    End;
    If d_Datelast Is Not Null Then
      d_Datefrom := d_Datelast;
      d_Dateto   := Sysdate;
      Update 住院费用记录
      Set 附加标志 = 5
      Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 <> 8 And Nvl(医嘱序号, 0) = 0 And
            发生时间 >= d_Datefrom;
    End If;
  End If;

  -----------------------------------------------------------------
  --作废以前计算的错误记录
  -----------------------------------------------------------------
  Insert Into 住院费用记录
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 姓名, 性别, 年龄, 标识号,
     床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 划价人, 操作员编号, 操作员姓名, 发生时间,
     登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id)
    Select 病人费用记录_Id.Nextval, 记录性质, NO, 2, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id,
           姓名, 性别, 年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 0, 标准单价, 付数, -数次, -应收金额, -实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人,
           划价人, 操作员编号, 操作员姓名, 发生时间, d_登记时间, 保险项目否, 保险大类id, -统筹金额, 医疗小组id
    From 住院费用记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 5 And 发生时间 >= d_Datefrom;

  -----------------------------------------------------------------
  --填写病人余额
  -----------------------------------------------------------------
  Select Sum(Decode(附加标志, 0, 1, -1) * 实收金额) As 实收金额
  Into n_Summoney
  From 住院费用记录
  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And
        (NO = v_Billno Or 附加标志 = 5 And 发生时间 >= d_Datefrom);

  Update 病人余额
  Set 费用余额 = Nvl(费用余额, 0) + Nvl(n_Summoney, 0)
  Where 病人id = 病人id_In And 性质 = 1 And 类型 = 2
  Returning 费用余额 Into n_返回值;

  If Sql%RowCount = 0 Then
    Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 2, n_Summoney, 0);
    n_返回值 := n_Summoney;
  End If;

  If Nvl(n_返回值, 0) = 0 Then
    Delete From 病人余额 Where 性质 = 1 And 病人id = 病人id_In And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
  End If;

  -----------------------------------------------------------------
  --填写病人汇总费用
  -----------------------------------------------------------------
  n_Delete := 0;
  For v_Currrow In v_Sumcur(v_Billno, d_Datefrom) Loop
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(v_Currrow.实收金额, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(v_Currrow.病人病区id, 0) And
          Nvl(病人科室id, 0) = Nvl(v_Currrow.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(v_Currrow.开单部门id, 0) And
          Nvl(执行部门id, 0) = Nvl(v_Currrow.执行部门id, 0) And Nvl(收入项目id, 0) = Nvl(v_Currrow.收入项目id, 0) And 来源途径 + 0 = 2
    Returning 金额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, 主页id_In, v_Currrow.病人病区id, v_Currrow.病人科室id, v_Currrow.开单部门id, v_Currrow.执行部门id, v_Currrow.收入项目id, 2,
         v_Currrow.实收金额);
      n_返回值 := v_Currrow.实收金额;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      n_Delete := 1;
    End If;
  End Loop;

  If Nvl(n_Delete, 0) = 1 Then
    Delete From 病人未结费用 Where 病人id = 病人id_In And 金额 = 0;
  End If;

  -----------------------------------------------------------------
  --将所有修改的附加标志还原为正常标志
  -----------------------------------------------------------------
  Update 住院费用记录
  Set 附加标志 = 0, 记录状态 = 3
  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 5 And 发生时间 >= d_Datefrom;

  -----------------------------------------------------------------
  --修改计算时间标志
  -----------------------------------------------------------------
  Update 病人变动记录
  Set 上次计算时间 = Least(d_Dateto, Nvl(终止时间, Greatest(开始时间, Sysdate)))
  Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(终止时间, Sysdate) > d_Datefrom;
  Commit; --单个病人提交
End Zl1_Autocptone;
/

--89624:刘尔旋,2015-10-20,消费卡发零费用卡
Create Or Replace Procedure Zl_医疗卡记录_Insert
(
  --参数：发卡类型=0-发卡,1-补卡,2-换卡(相当于重打)
  --      换卡时,单据号_IN传入的是原发/补卡的单据号。
  --      补卡/换卡后,再换卡时是以最后一次卡号为准。
  发卡类型_In   Number,
  单据号_In     住院费用记录.No%Type,
  病人id_In     住院费用记录.病人id%Type,
  主页id_In     住院费用记录.主页id%Type,
  标识号_In     住院费用记录.标识号%Type,
  费别_In       住院费用记录.费别%Type,
  卡类别id_In   医疗卡类别.Id%Type,
  原卡号_In     病人医疗卡信息.卡号%Type,
  医疗卡号_In   病人医疗卡信息.卡号%Type,
  变动原因_In   病人医疗卡变动.变动原因%Type,
  密码_In       病人信息.卡验证码%Type,
  姓名_In       住院费用记录.姓名%Type,
  性别_In       住院费用记录.性别%Type,
  年龄_In       住院费用记录.年龄%Type,
  病人病区id_In 住院费用记录.病人病区id%Type,
  病人科室id_In 住院费用记录.病人科室id%Type,
  收费细目id_In 住院费用记录.收费细目id%Type,
  收费类别_In   住院费用记录.收费类别%Type,
  计算单位_In   住院费用记录.计算单位%Type,
  收入项目id_In 住院费用记录.收入项目id%Type,
  收据费目_In   住院费用记录.收据费目%Type,
  标准单价_In   住院费用记录.标准单价%Type,
  执行部门id_In 住院费用记录.执行部门id%Type,
  开单部门id_In 住院费用记录.开单部门id%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  加班标志_In   住院费用记录.加班标志%Type,
  发卡时间_In   住院费用记录.登记时间%Type,
  领用id_In     票据使用明细.领用id%Type,
  Ic卡号_In     病人信息.Ic卡号%Type := Null,
  应收金额_In   住院费用记录.应收金额%Type,
  实收金额_In   住院费用记录.实收金额%Type,
  结算方式_In   病人预交记录.结算方式%Type,
  刷卡类别id_In 病人预交记录.卡类别id%Type,
  消费卡_In     Integer := 0,
  刷卡卡号_In   病人医疗卡信息.卡号%Type,
  结帐id_In     病人预交记录.结帐id%Type,
  交易流水号_In 病人预交记录.交易流水号%Type := Null,
  交易说明_In   病人预交记录.交易说明%Type := Null,
  合作单位_In   病人预交记录.合作单位%Type := Null
  
) As

  Cursor c_Precard Is
    Select ID As 费用id From 住院费用记录 Where 记录性质 = 5 And 实际票号 = 原卡号_In And 病人id = 病人id_In;
  r_Cardrow c_Precard%RowType;

  Cursor c_医疗卡 Is
    Select ID, 编码, 名称, 短名, 前缀文本, 卡号长度, 缺省标志, 是否固定, 是否严格控制, Nvl(是否刷卡, 0) As 是否刷卡, Nvl(是否自制, 0) As 是否自制,
           Nvl(是否存在帐户, 0) As 是否存在帐户, Nvl(是否全退, 0) As 是否全退, 部件, 备注, 特定项目, 结算方式, 是否启用, 卡号密文, Nvl(是否重复使用, 0) As 是否重复使用
    From 医疗卡类别
    Where ID = 卡类别id_In;
  r_医疗卡 c_医疗卡%RowType;

  v_费用id         住院费用记录.Id%Type;
  v_结帐id         住院费用记录.结帐id%Type;
  v_收回id         票据打印内容.Id%Type;
  v_打印id         票据打印内容.Id%Type;
  n_回收次数       票据使用明细.回收次数%Type;
  n_性质           票据使用明细.性质%Type;
  n_返回值         病人余额.费用余额%Type;
  n_Count          Number(18);
  n_预交id         病人预交记录.Id%Type;
  n_消费卡id       消费卡目录.Id%Type;
  n_自制卡         Number;
  n_医疗卡重复使用 Number(3);
  Err_Item Exception;
  v_Err_Msg  Varchar2(500);
  n_组id     财务缴款分组.Id%Type;
  n_变动类型 Number;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  Open c_医疗卡;
  Fetch c_医疗卡
    Into r_医疗卡;
  If c_医疗卡%RowCount = 0 Then
    Close c_医疗卡;
    v_Err_Msg := '[ZLSOFT]没有发现原医疗卡的相应类别,不能继续操作！[ZLSOFT]';
    Raise Err_Item;
  End If;

  n_医疗卡重复使用 := Nvl(r_医疗卡.是否重复使用, 0);
  Close c_医疗卡;
  If Not 结算方式_In Is Null Then
    If Nvl(结帐id_In, 0) <> 0 Then
      v_结帐id := 结帐id_In;
    Else
      Select 病人结帐记录_Id.Nextval Into v_结帐id From Dual;
    End If;
  End If;
  If 发卡类型_In <> 2 Then
    --发卡和补卡
    Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
  
    Insert Into 住院费用记录
      (ID, 记录性质, 记录状态, NO, 实际票号, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 标识号, 姓名, 性别, 年龄, 费别, 记帐费用, 门诊标志, 加班标志, 开单部门id, 开单人,
       操作员编号, 操作员姓名, 发生时间, 登记时间, 收费细目id, 收费类别, 计算单位, 付数, 数次, 发药窗口, 附加标志, 执行部门id, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 结帐id,
       结帐金额, 缴款组id, 结论)
    Values
      (v_费用id, 5, 1, 单据号_In, 医疗卡号_In, 1, 病人id_In, Decode(主页id_In, 0, Null, 主页id_In),
       Decode(病人病区id_In, 0, Null, 病人病区id_In), Decode(病人科室id_In, 0, Null, 病人科室id_In), Decode(标识号_In, 0, Null, 标识号_In),
       姓名_In, 性别_In, 年龄_In, 费别_In, Decode(结算方式_In, Null, 1, 0), 3, 加班标志_In, 开单部门id_In, 操作员姓名_In, 操作员编号_In, 操作员姓名_In,
       发卡时间_In, 发卡时间_In, 收费细目id_In, 收费类别_In, 计算单位_In, 1, 1, 医疗卡号_In, 发卡类型_In, 执行部门id_In, 收入项目id_In, 收据费目_In, 标准单价_In,
       应收金额_In, 实收金额_In, v_结帐id, Decode(结算方式_In, Null, Null, 实收金额_In), n_组id, 卡类别id_In);
  
    --如果是现收医疗卡费用，则将结算填入病人预交记录
    If Not 结算方式_In Is Null Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, NO, 记录性质, 记录状态, 病人id, 主页id, 科室id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 摘要, 缴款组id, 卡类别id, 卡号, 结算卡序号, 交易流水号,
         交易说明, 结算序号, 合作单位, 结算性质)
      Values
        (n_预交id, 单据号_In, 5, 1, 病人id_In, Decode(主页id_In, 0, Null, 主页id_In), Decode(病人科室id_In, 0, Null, 病人科室id_In),
         结算方式_In, 发卡时间_In, 操作员编号_In, 操作员姓名_In, 实收金额_In, v_结帐id, '医疗卡费用', n_组id, Decode(消费卡_In, 0, 刷卡类别id_In, Null),
         刷卡卡号_In, Decode(消费卡_In, 0, Null, 刷卡类别id_In), 交易流水号_In, 交易说明_In, v_结帐id, 合作单位_In, 5);
    
      If 消费卡_In = 1 And 刷卡卡号_In Is Not Null Then
      
        n_消费卡id := Null;
        Begin
          Select Nvl(自制卡, 0), 1 Into n_自制卡, n_Count From 卡消费接口目录 Where 编号 = 刷卡类别id_In;
        Exception
          When Others Then
            n_Count := 0;
        End;
        If n_Count = 0 Then
          v_Err_Msg := '[ZLSOFT]没有发现原结算卡的相应类别,不能继续操作！[ZLSOFT]';
          Raise Err_Item;
        End If;
        If n_自制卡 = 1 Then
          Select ID
          Into n_消费卡id
          From 消费卡目录
          Where 接口编号 = 刷卡类别id_In And 卡号 = 刷卡卡号_In And
                序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = 刷卡类别id_In And 卡号 = 刷卡卡号_In);
        End If;
        Zl_病人卡结算记录_Insert(刷卡类别id_In, n_消费卡id, 结算方式_In, 实收金额_In, 刷卡卡号_In, Null, Null, Null, v_结帐id, n_预交id);
      End If;
    End If;
  
    --发卡使用票据
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 5, 单据号_In);
    n_回收次数 := 0;
    If n_医疗卡重复使用 = 1 Then
      Select Nvl(Max(回收次数), 0), Nvl(Max(性质), 0)
      Into n_回收次数, n_性质
      From 票据使用明细
      Where 票种 = 5 And 号码 = 医疗卡号_In;
      If n_回收次数 > 0 Or n_性质 > 0 Then
        n_回收次数 := n_回收次数 + 1;
      End If;
    Else
      --需要检查是否存在票据使用明细，如果存在，肯定会发生错误
      Select Nvl(Max(性质), 0)
      Into n_性质
      From 票据使用明细 A, 票据领用记录 B
      Where a.票种 = 5 And a.号码 = 医疗卡号_In And Nvl(a.领用id, 0) = Nvl(领用id_In, 0) And a.领用id = b.Id;
      If n_性质 <> 0 Then
        v_Err_Msg := '[ZLSOFT]卡号:' || 医疗卡号_In || ' 已经使用，不能再进行发卡操作,请检查![ZLSOFT]';
        Raise Err_Item;
      End If;
    End If;
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 回收次数, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, 5, 医疗卡号_In, 1, 1, 领用id_In, Decode(n_回收次数, 0, Null, n_回收次数), v_打印id, 发卡时间_In, 操作员姓名_In);
    --如果是回收,再发的,则不减剩余数量
    If Nvl(n_回收次数, 0) = 0 Then
      --该批领用状态变化
      Update 票据领用记录
      Set 当前号码 = 医疗卡号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
      Where ID = Nvl(领用id_In, 0);
    End If;
  
    --相关汇总表的处理
    If 结算方式_In Is Null Then
      --汇总'病人余额'
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + 实收金额_In
      Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 2) = Decode(Nvl(主页id_In, 0), 0, 1, 2)
      Returning 费用余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (病人id_In, 1, Decode(Nvl(主页id_In, 0), 0, 1, 2), 0, 实收金额_In);
        n_返回值 := 实收金额_In;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --汇总'病人未结费用'
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + 实收金额_In
      Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(病人病区id_In, 0) And
            Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And
            Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And 来源途径 = 3;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Decode(主页id_In, 0, Null, 主页id_In), Decode(病人病区id_In, 0, Null, 病人病区id_In),
           Decode(病人科室id_In, 0, Null, 病人科室id_In), 开单部门id_In, 执行部门id_In, 收入项目id_In, 3, 实收金额_In);
      End If;
    
    Else
      --汇总"人员缴款余额"
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + 实收金额_In
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = 结算方式_In
      Returning 余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, 结算方式_In, 1, 实收金额_In);
        n_返回值 := 实收金额_In;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
      End If;
    End If;
  
  Else
    --处理换卡方式
    --首先查找需要换卡的原医疗卡费用记录
    Open c_Precard;
    Fetch c_Precard
      Into r_Cardrow;
  
    If c_Precard%RowCount = 0 Then
      Close c_Precard;
      v_Err_Msg := '[ZLSOFT]没有发现原医疗卡发放记录,换卡操作失败！[ZLSOFT]';
      Raise Err_Item;
    Else
      --仅当有原费用记录时才处理
      --重打收回票据
      Begin
        Select ID
        Into v_收回id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 5 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    
      If v_收回id Is Not Null Then
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 回收次数, 打印id, 使用时间, 使用人)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 4, 领用id, 回收次数, 打印id, 发卡时间_In, 操作员姓名_In
          From 票据使用明细
          Where 打印id = v_收回id And 票种 = 5 And 性质 = 1;
      End If;
    
      --重打发出票据
      Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
    
      Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 5, 单据号_In);
      n_回收次数 := 0;
      If n_医疗卡重复使用 = 1 Then
        Select Nvl(Max(回收次数), 0), Nvl(Max(性质), 0)
        Into n_回收次数, n_性质
        From 票据使用明细
        Where 票种 = 5 And 号码 = 医疗卡号_In;
        If n_回收次数 > 0 Or n_性质 > 0 Then
          n_回收次数 := n_回收次数 + 1;
        End If;
      Else
        --需要检查是否存在票据使用明细，如果存在，肯定会发生错误
        Select Nvl(Max(性质), 0)
        Into n_性质
        From 票据使用明细 A, 票据领用记录 B
        Where a.票种 = 5 And a.号码 = 医疗卡号_In And Nvl(a.领用id, 0) = Nvl(领用id_In, 0) And a.领用id = b.Id;
        If n_性质 <> 0 Then
          v_Err_Msg := '[ZLSOFT]新卡号:' || 医疗卡号_In || ' 已经使用，请换一张新卡,请检查![ZLSOFT]';
          Raise Err_Item;
        End If;
      
      End If;
    
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 回收次数, 打印id, 使用时间, 使用人)
      Values
        (票据使用明细_Id.Nextval, 5, 医疗卡号_In, 1, Decode(v_收回id, Null, 1, 3), 领用id_In, Decode(n_回收次数, 0, Null, n_回收次数), v_打印id,
         发卡时间_In, 操作员姓名_In);
      --如果是回收,再发的,则不减剩余数量
      If Nvl(n_回收次数, 0) = 0 Then
        --领用状态变化
        Update 票据领用记录
        Set 当前号码 = 医疗卡号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
        Where ID = Nvl(领用id_In, 0);
      End If;
      --更改原发卡记录状态
      Update 住院费用记录
      Set 实际票号 = 医疗卡号_In, 发药窗口 = 医疗卡号_In, 附加标志 = 2, 结论 = 卡类别id_In
      Where ID = r_Cardrow.费用id;
      Close c_Precard;
    End If;
  End If;

  --处理相关的变动信息
  --Zl_医疗卡变动_Insert (变动类型_In/病人id_In ,卡类别id_In, 原卡号_In, 医疗卡号_In, 变动原因_In, 密码_In, 操作员姓名_In, 变动时间_In
  --Ic卡号_In, 挂失方式_In)
  --变动类型_In:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
  n_变动类型 := Case
              When 发卡类型_In = 0 Then
               1
              When 发卡类型_In = 1 Then
               3
              Else
               2
            End;
  Zl_医疗卡变动_Insert(n_变动类型, 病人id_In, 卡类别id_In, 原卡号_In, 医疗卡号_In, 变动原因_In, 密码_In, 操作员姓名_In, 发卡时间_In, Ic卡号_In, Null);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_医疗卡记录_Insert;
/

--88607:王振涛,2015-10-19,解决4000个字符上传的问题
Create Or Replace Procedure Zl_检验报告单_Apply
(
  v_Strval In Varchar2,
  n_Type   In Number,
  N_Begin  In Number  := 0--开始为0
  --功能           把检验结果回写入HIS病历中,
  --参数           v_Strval 专入的标本结果内容，单条申请结果
  --               格式:
  --               类型(1=普通)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2>婴儿序号<split2>
  --                   指标1<split4>检验结果1<split4>单位1<split4>结果标志1<split4>结果参数1<split4>排列序号1<split4>隐私项目1<split4>指标代码1<split3>
  --                   指标2<split4>检验结果2<split4>单位2<split4>结果标志2<split4>结果参数2<split4>排列序号2<split4>隐私项目2<split4>指标代码2<split3>
  --                   指标3<split4>检验结果3<split4>单位3<split4>结果标志3<split4>结果参数3<split4>排列序号3<split4>隐私项目3<split4>指标代码3
  --
  --               类型(2=微生物)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2>
  --               细菌名1<split3>描述1<split3>耐药机制1<split3>
  --                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
  --                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split2>
  --               细菌名2<split3>描述2<split3>耐药机制2<split3>
  --                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
  --                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22
  --               intType 0=审核 1=取消审核
) Is
  n_Patiid     病人医嘱记录.病人id%Type;
  n_Pageid     电子病历记录.主页id%Type;
  n_Orderid    病人医嘱记录.Id%Type;
  n_Deptid     病人医嘱记录.开嘱科室id%Type;
  n_Patifrom   病人医嘱记录.病人来源%Type;
  n_Babytag    病人医嘱记录.婴儿%Type;
  v_Creator    电子病历记录.创建人%Type;
  d_Creator    电子病历记录.创建时间%Type;
  v_Speaker    电子病历记录.保存人%Type;
  d_Speaker    电子病历记录.完成时间%Type;
  n_Fileid     病历文件列表.Id%Type;
  v_Filename   病历文件列表.名称%Type;
  n_父id       电子病历内容.父id%Type;
  n_Recordid   电子病历内容.文件id%Type;
  n_Nextid     电子病历内容.Id%Type;
  n_l          Number := 0;
  n_No         Number := 0;
  v_Content    电子病历内容.内容文本%Type;
  v_Reporttag  Number;
  n_Reporttype Number;
  v_Stritems   Varchar2(4000);
  v_List       Varchar2(4000);
  n_i          Number;
  v_Listtype   Varchar2(6);
  Cursor v_Source Is
    Select ID, 文件id, Nvl(父id, 0) As 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 替换域, 要素名称
    From 病历文件结构
    Where 文件id = n_Fileid
    Order By 对象序号;
  Function v_Split
  (
    p_String    Varchar2,
    p_Separator Varchar2,
    p_Element   Integer
  ) Return Varchar2 As
    --实现VB的Split功能
    --返回在p_String中以p_Separator为分隔的第p_Element个元素串
    --第N个是从第一个开始计
    v_String Varchar2(32767);
  Begin
    v_String := p_String || p_Separator;
    For I In 1 .. p_Element - 1 Loop
      v_String := Substr(v_String, Instr(v_String, p_Separator) + Length(p_Separator));
    End Loop;
    Return Substr(v_String, 1, Instr(v_String, p_Separator) - 1);
  Exception
    When Others Then
      Return Null;
  End v_Split;
Begin
  --类型(1=普通)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split>婴儿序号
  --类型(2=微生物)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型
  n_Reporttype := v_Split(v_Strval, '<split2>', 1);
  n_Orderid    := v_Split(v_Strval, '<split2>', 2);
  n_Patifrom   := v_Split(v_Strval, '<split2>', 3);
  v_Creator    := v_Split(v_Strval, '<split2>', 5);
  v_Speaker    := v_Split(v_Strval, '<split2>', 6);
  d_Creator    := To_Date(v_Split(v_Strval, '<split2>', 4), 'yyyy-mm-dd HH24:mi:ss');
  d_Speaker    := To_Date(v_Split(v_Strval, '<split2>', 7), 'yyyy-mm-dd HH24:mi:ss');
  
  --读取病人相关信息
  Select Nvl(主页id, 0), Nvl(病人id, 0), Nvl(开嘱科室id, 0), Nvl(婴儿, 0)
  Into n_Pageid, n_Patiid, n_Deptid, n_Babytag
  From 病人医嘱记录
  Where ID = n_Orderid;
  If n_Patifrom = 1 Then
    --主页ID： 门诊病人填挂号ID
    Select Nvl(b.Id, 0)
    Into n_Pageid
    From 病人挂号记录 B, 病人医嘱记录 A
    Where a.挂号单 = b.No(+) And a.Id = n_Orderid;
  End If;

  Begin
    Select 病历文件id, c.名称
    Into n_Fileid, v_Filename
    From 病人医嘱记录 A, 病历单据应用 B, 病历文件列表 C
    Where a.诊疗项目id = b.诊疗项目id And b.病历文件id = c.Id And a.相关id = n_Orderid And b.应用场合 = n_Patifrom And Rownum <= 1;
  Exception
    When Others Then
      Return;
  End;

  Begin
    --删除以前的报告记录
    --删除以前的报告记录 
    Select 病历id Into N_Recordid From 病人医嘱报告 Where 医嘱id = N_Orderid;
  
    If N_Begin = 0 Then
      Delete 报告查阅记录 Where 医嘱id = N_Orderid;
      Delete 病人医嘱报告 Where 医嘱id = N_Orderid;
      Delete 电子病历记录 Where ID = N_Recordid;
      Delete 电子病历内容 Where 文件id = N_Recordid;
    Else
	For R_Source In V_Source Loop
      
        V_Reporttag := 0;
      
        If R_Source.对象类型 = 1 And R_Source.内容文本 = '检验结果' Then
          Select ID
          Into N_Nextid
          From 电子病历内容
          Where 文件id = N_Recordid And 内容文本 = '检验结果' And 对象标记 = R_Source.对象标记 And 对象类型 = R_Source.对象类型;
        
          Select Max(对象序号) Into N_No From 电子病历内容 Where 文件id = N_Recordid;
        
          V_Reporttag := 1;
        
        End If;
      
        If V_Reporttag = 1 Then
          --在 '检验结果' 提纲下插入检验结果 
          N_父id := N_Nextid;
          If N_Reporttype = 1 Then
            --取出每组指标集,直到为空时退出循环 
            V_Stritems := V_Split(V_Strval, '<split2>', 11);
            For N_l In 1 .. 999 Loop
              V_List := V_Split(V_Stritems, '<split3>', N_l);
              Exit When V_List Is Null;
              N_No := N_No + 1;
              Select 电子病历内容_Id.Nextval Into N_Nextid From Dual;
              v_Content := RPad(Nvl(v_Split(v_List, '<split4>', 1), ' '), 35) ||
                       RPad(Nvl(v_Split(v_List, '<split4>', 2), ' '), 10) ||
                       LPad(Nvl(v_Split(v_List, '<split4>', 3), ' '), 8) ||
                       LPad(Nvl(v_Split(v_List, '<split4>', 4), ' '), 10) ||
                       LPad(Nvl(v_Split(v_List, '<split4>', 5), ' '), 13);
              Insert Into 电子病历内容
                (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
              Values
                (N_Nextid, N_Recordid, 1, 0, N_父id, N_No, 2, N_No, Null, 0, Null, V_Content, 1);
            End Loop;
          Else
            --微生物标本 
            N_父id := N_Nextid;
          
            --取得每组细菌指标结果集,直到为空时退出循环,位于第10个开始 
            For N_i In 10 .. 999 Loop
              V_Stritems := V_Split(V_Strval, '<split2>', N_i);
              Exit When V_Stritems Is Null;
              --取出每组抗生素指标结果串,直到为空时退出循环,位于第4个开始 
              For N_l In 4 .. 999 Loop
                V_List := V_Split(V_Stritems, '<split3>', N_l);
                Exit When V_List Is Null;
                N_No      := N_No + 1;
                v_Content := RPad(Nvl(v_Split(v_List, '<split4>', 1), ' '), 20) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 2), ' '), 5) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 3), ' '), 12) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 5), ' '), 12) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 7), ' '), 7) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 9), ' '), 7);
                Select 电子病历内容_Id.Nextval Into N_Nextid From Dual;
                Insert Into 电子病历内容
                  (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
                Values
                  (N_Nextid, N_Recordid, 1, 0, N_父id, N_No, 2, N_No, Null, 0, Null, V_Content, 1);
              End Loop;
            End Loop;
          End If;
        End If;
      End Loop;
    END IF;
  Exception
    When Others Then
      Null;
  End;

  If N_Type = 1 Or N_Begin > 0 Then
    --取消审核 
    Return;
  End If;

  

  ----生成新电子病历记录、医嘱报告关联
  Select 电子病历记录_Id.Nextval Into n_Recordid From Dual;
  Insert Into 电子病历记录
    (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 完成时间, 保存人, 保存时间, 最后版本, 签名级别)
  Values
    (n_Recordid, n_Patifrom, n_Patiid, n_Pageid, n_Babytag, n_Deptid, 7, n_Fileid, v_Filename, v_Creator, d_Creator,
     d_Speaker, v_Speaker, Sysdate, 1, 0);
  Insert Into 病人医嘱报告 (医嘱id, 病历id) Values (n_Orderid, n_Recordid);

  For r_Source In v_Source Loop
    n_No := n_No + 1;
    Select 电子病历内容_Id.Nextval Into n_Nextid From Dual;
    v_Reporttag := 0;
  
    If r_Source.对象类型 = 4 And r_Source.替换域 = 1 Then
      Select Zl_Replace_Element_Value(r_Source.要素名称, n_Patiid, n_Pageid, n_Patifrom, n_Orderid)
      Into v_Content
      From Dual;
      Insert Into 电子病历内容
        (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
      Values
        (n_Nextid, n_Recordid, 1, 0, Null, n_No, r_Source.对象类型, r_Source.对象标记, r_Source.保留对象, r_Source.对象属性,
         r_Source.内容行次, v_Content, r_Source.是否换行);
    Elsif r_Source.对象类型 = 1 And r_Source.内容文本 = '检验结果' Then
      Insert Into 电子病历内容
        (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
      Values
        (n_Nextid, n_Recordid, 1, 0, Null, n_No, r_Source.对象类型, r_Source.对象标记, r_Source.保留对象, r_Source.对象属性,
         r_Source.内容行次, r_Source.内容文本, r_Source.是否换行);
      v_Reporttag := 1;
    Else
      Insert Into 电子病历内容
        (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
      Values
        (n_Nextid, n_Recordid, 1, 0, Null, n_No, r_Source.对象类型, r_Source.对象标记, r_Source.保留对象, r_Source.对象属性,
         r_Source.内容行次, r_Source.内容文本, r_Source.是否换行);
    End If;
  
    If v_Reporttag = 1 Then
      --在 '检验结果' 提纲下插入检验结果
      n_父id := n_Nextid;
      If n_Reporttype = 1 Then
        --普通标本
        n_No := n_No + 1;
        Select 电子病历内容_Id.Nextval Into n_Nextid From Dual;
        v_Content := RPad('检验项目', 35) || RPad('检验结果', 10) || LPad('单位', 8) || LPad('结果标志', 10) || LPad('结果参考', 13);
        Insert Into 电子病历内容
          (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
        Values
          (n_Nextid, n_Recordid, 1, 0, n_父id, n_No, 2, n_No, Null, 0, Null, v_Content, 1);
        --取出每组指标集,直到为空时退出循环
        v_Stritems := v_Split(v_Strval, '<split2>', 11);
        For n_l In 1 .. 999 Loop
          v_List := v_Split(v_Stritems, '<split3>', n_l);
          Exit When v_List Is Null;
          n_No := n_No + 1;
          Select 电子病历内容_Id.Nextval Into n_Nextid From Dual;
          v_Content := RPad(Nvl(v_Split(v_List, '<split4>', 1), ' '), 35) ||
                       RPad(Nvl(v_Split(v_List, '<split4>', 2), ' '), 10) ||
                       LPad(Nvl(v_Split(v_List, '<split4>', 3), ' '), 8) ||
                       LPad(Nvl(v_Split(v_List, '<split4>', 4), ' '), 10) ||
                       LPad(Nvl(v_Split(v_List, '<split4>', 5), ' '), 13);
          Insert Into 电子病历内容
            (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
          Values
            (n_Nextid, n_Recordid, 1, 0, n_父id, n_No, 2, n_No, Null, 0, Null, v_Content, 1);
        End Loop;
      Else
        --微生物标本
        n_父id := n_Nextid;
      
        --取得每组细菌指标结果集,直到为空时退出循环,位于第10个开始
        For n_i In 10 .. 999 Loop
          v_Stritems := v_Split(v_Strval, '<split2>', n_i);
          Exit When v_Stritems Is Null;
          n_No := n_No + 1;
        
          v_Content  := '       鉴定结果：' || v_Split(v_Stritems, '<split3>', 1) || v_Split(v_Stritems, '<split3>', 2) ||
                        v_Split(v_Stritems, '<split3>', 3) || Chr(13) || Chr(10);
          v_Content  := v_Content || '       抗生素          ';
          v_Listtype := v_Split(v_Stritems, '<split4>', 4);
          Select v_Content || RPad(Decode(v_Listtype, '1-MIC', '  MIC', '2-Disk', ' Disk', '3-K-B', '  K-B', ' '), 5)
          Into v_Content
          From Dual;
          v_Content := v_Content || RPad('耐药性', 12) || RPad('用法用量', 12) || RPad('血药浓度', 7) || RPad('尿药浓度', 7) ||
                       Chr(13) || Chr(10);
          v_Content := v_Content || Replace(RPad(' ', 47, '￣'), ' ', '￣');
          Select 电子病历内容_Id.Nextval Into n_Nextid From Dual;
          Insert Into 电子病历内容
            (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
          Values
            (n_Nextid, n_Recordid, 1, 0, n_父id, n_No, 2, n_No, Null, 0, Null, v_Content, 1);
        
          --取出每组抗生素指标结果串,直到为空时退出循环,位于第4个开始
          For n_l In 4 .. 999 Loop
            v_List := v_Split(v_Stritems, '<split3>', n_l);
            Exit When v_List Is Null;
            n_No      := n_No + 1;
            v_Content := RPad(Nvl(v_Split(v_List, '<split4>', 1), ' '), 20) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 2), ' '), 5) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 3), ' '), 12) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 5), ' '), 12) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 7), ' '), 7) ||
                         RPad(Nvl(v_Split(v_List, '<split4>', 9), ' '), 7);
            Select 电子病历内容_Id.Nextval Into n_Nextid From Dual;
            Insert Into 电子病历内容
              (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
            Values
              (n_Nextid, n_Recordid, 1, 0, n_父id, n_No, 2, n_No, Null, 0, Null, v_Content, 1);
          End Loop;
        End Loop;
      End If;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验报告单_Apply;
/
--89530:刘尔旋,2015-10-15,锁号接口错误处理消息
Create Or Replace Procedure Zl_Third_Lockno
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS锁号
  --入参:Xml_In:
  --<IN>
  --  <HM>5</HM>           //号码
  --  <RQ>2013-11-21 09:00</RQ>     //预约时间
  --  <CZ>1</CZ>           //操作
  --  <HX></HX>          //号序
  --  <HZDW>支付宝</HZDW>   //合作单位
  --  <JQM>机器名</JQM>        //机器名
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <HX>号序</HX>          //锁号操作并且成功时返回
  -- 错误信息  //出错时返回
  --</ OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_号码         挂号安排.号码%Type;
  d_日期         Date;
  n_操作类型     Number(3);
  n_序号控制     Number(3);
  n_存在         Number(3);
  n_分时段       Number(3);
  n_限号数       挂号安排限制.限号数%Type;
  n_安排id       挂号安排.Id%Type;
  n_计划id       挂号安排计划.Id%Type;
  n_号序         挂号序号状态.序号%Type;
  v_星期         挂号安排限制.限制项目%Type;
  v_操作员姓名   挂号序号状态.操作员姓名%Type;
  v_机器名       挂号序号状态.机器名%Type;
  v_验证姓名     挂号序号状态.操作员姓名%Type;
  v_验证机器名   挂号序号状态.机器名%Type;
  n_状态         挂号序号状态.状态%Type;
  v_合作单位     合作单位安排控制.合作单位%Type;
  n_合约模式     Number(3);
  n_启用合作单位 Number(3);
  v_Temp         Varchar2(32767); --临时XML
  v_Optemp       Varchar2(300);
  x_Templet      Xmltype; --模板XML
  v_Err_Msg      Varchar2(200);
  n_Exists       Number(2);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), To_Date(Extractvalue(Value(A), 'IN/RQ'), 'YYYY-MM-DD hh24:mi:ss'),
         Extractvalue(Value(A), 'IN/CZ'), Extractvalue(Value(A), 'IN/HX'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/JQM')
  Into v_号码, d_日期, n_操作类型, n_号序, v_合作单位, v_机器名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If v_机器名 Is Null Then
    Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  End If;
  v_Optemp := Zl_Identity(1);
  Select Substr(v_Optemp, Instr(v_Optemp, ',') + 1) Into v_Optemp From Dual;
  Select Substr(v_Optemp, Instr(v_Optemp, ',') + 1) Into v_操作员姓名 From Dual;
  If n_操作类型 = 0 Then
    --解锁
    Begin
      Select 1
      Into n_Exists
      From 挂号序号状态
      Where 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 状态 = 5 And 序号 = n_号序 And Trunc(日期) = Trunc(d_日期) And 号码 = v_号码 And
            Rownum < 2;
    Exception
      When Others Then
        n_Exists := 0;
    End;
    If n_Exists = 1 Then
      Delete 挂号序号状态
      Where 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 状态 = 5 And 序号 = n_号序 And Trunc(日期) = Trunc(d_日期) And 号码 = v_号码;
      v_Temp := '<HX>' || n_号序 || '</HX>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Else
      v_Temp := '没有发现需要解锁的序号';
      Raise Err_Item;
    End If;
  End If;

  If n_操作类型 = 1 Then
    --锁号
    Select Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
    Into v_星期
    From Dual;
    Begin
      Select 序号控制, ID
      Into n_序号控制, n_计划id
      From (Select 序号控制, ID
             From 挂号安排计划
             Where 号码 = v_号码 And d_日期 Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                   Nvl(失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And 审核时间 Is Not Null
             Order By 生效时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Select 序号控制, ID Into n_序号控制, n_安排id From 挂号安排 Where 号码 = v_号码;
    End;
    If n_序号控制 = 1 Then
      If Nvl(n_计划id, 0) <> 0 Then
        Begin
          Select 1 Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum < 2;
        Exception
          When Others Then
            n_分时段 := 0;
        End;
        Begin
          Select 1
          Into n_启用合作单位
          From 合作单位计划控制
          Where 限制项目 = v_星期 And 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
        Exception
          When Others Then
            n_启用合作单位 := 0;
        End;
        Begin
          Select 1, a.状态, a.操作员姓名, a.机器名
          Into n_存在, n_状态, v_验证姓名, v_验证机器名
          From 挂号序号状态 A, 挂号计划时段 B
          Where a.号码 = v_号码 And Trunc(a.日期) = Trunc(d_日期) And a.序号 = b.序号 And b.计划id = n_计划id And b.星期 = v_星期 And
                To_Char(b.开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And Rownum < 2;
        Exception
          When Others Then
            n_存在 := 0;
        End;
        If n_存在 = 1 Then
          If n_状态 = 5 And v_验证姓名 = v_操作员姓名 And v_机器名 = v_验证机器名 Then
            Null;
          Else
            --传入时间的序号已经被使用
            v_Temp := '传入时间' || d_日期 || '的序号已被使用';
            Raise Err_Item;
          End If;
        Else
          If n_分时段 = 1 Then
            Begin
              Select 序号
              Into n_号序
              From 挂号计划时段
              Where 计划id = n_计划id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And Rownum < 2;
            Exception
              When Others Then
                Select Max(序号) + 1
                Into n_号序
                From (Select Distinct 序号
                       From 挂号计划时段
                       Where 计划id = n_计划id And 星期 = v_星期
                       Union
                       Select Distinct 序号 From 挂号序号状态 Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期));
              
            End;
            Begin
              Select 1 Into n_存在 From 挂号序号状态 Where 号码 = v_号码 And 日期 = d_日期 And 序号 = n_号序;
            Exception
              When Others Then
                Insert Into 挂号序号状态
                  (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                Values
                  (v_号码, d_日期, n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
            End;
            v_Temp := '<HX>' || n_号序 || '</HX>';
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Else
            If v_合作单位 Is Null Or n_启用合作单位 = 0 Then
              If Trunc(d_日期) = Trunc(Sysdate) Then
                n_号序 := 1;
                Select 限号数 Into n_限号数 From 挂号计划限制 Where 计划id = n_计划id And 限制项目 = v_星期;
                For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                             From 挂号序号状态
                             Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                             Order By 序号) Loop
                  Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                  If r_序号.序号 = n_号序 Then
                    n_号序 := n_号序 + 1;
                  End If;
                End Loop;
                If n_号序 > n_限号数 Then
                  v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                  Raise Err_Item;
                Else
                  Begin
                    Select 1
                    Into n_存在
                    From 挂号序号状态
                    Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                  Exception
                    When Others Then
                      Insert Into 挂号序号状态
                        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                      Values
                        (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                  End;
                  v_Temp := '<HX>' || n_号序 || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              Else
                n_号序 := 1;
                Select 限号数 Into n_限号数 From 挂号计划限制 Where 计划id = n_计划id And 限制项目 = v_星期;
                For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                             From 挂号序号状态
                             Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                             
                             Union
                             Select 序号, Null, Null, Null
                             From 合作单位计划控制
                             Where 计划id = n_计划id And 限制项目 = v_星期 And 数量 <> 0
                             Order By 序号) Loop
                  Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                  If r_序号.序号 = n_号序 Then
                    n_号序 := n_号序 + 1;
                  End If;
                End Loop;
                If n_号序 > n_限号数 Then
                  v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                  Raise Err_Item;
                Else
                  Begin
                    Select 1
                    Into n_存在
                    From 挂号序号状态
                    Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                  Exception
                    When Others Then
                      Insert Into 挂号序号状态
                        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                      Values
                        (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                  End;
                  v_Temp := '<HX>' || n_号序 || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              End If;
            Else
              Select Count(1)
              Into n_合约模式
              From 合作单位计划控制
              Where 序号 = 0 And 计划id = n_计划id And 合作单位 = v_合作单位 And 限制项目 = v_星期 And 数量 <> 0;
              If n_合约模式 = 0 Then
                Begin
                  Select 序号
                  Into n_号序
                  From (Select 序号
                         From 合作单位计划控制 A
                         Where 计划id = n_计划id And 合作单位 = v_合作单位 And 限制项目 = v_星期 And 数量 <> 0 And
                               (Not Exists
                                (Select 1
                                 From 挂号序号状态
                                 Where 号码 = v_号码 And 序号 = a.序号 And Trunc(日期) = Trunc(d_日期) And 状态 <> 5) Or Exists
                                (Select 1
                                 From 挂号序号状态
                                 Where 号码 = v_号码 And 序号 = a.序号 And Trunc(日期) = Trunc(d_日期) And 状态 = 5 And 操作员姓名 = v_操作员姓名 And
                                       机器名 = v_机器名))
                         Order By 序号)
                  Where Rownum < 2;
                Exception
                  When Others Then
                    n_号序 := 0;
                End;
                If n_号序 = 0 Then
                  v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                  Raise Err_Item;
                Else
                  Begin
                    Select 1
                    Into n_存在
                    From 挂号序号状态
                    Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                  Exception
                    When Others Then
                      Insert Into 挂号序号状态
                        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                      Values
                        (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                  End;
                  v_Temp := '<HX>' || n_号序 || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              Else
                n_号序 := 1;
                Select 限号数 Into n_限号数 From 挂号计划限制 Where 计划id = n_计划id And 限制项目 = v_星期;
                For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                             From 挂号序号状态
                             Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                             
                             Union
                             Select 序号, Null, Null, Null
                             From 合作单位计划控制
                             Where 计划id = n_计划id And 限制项目 = v_星期 And 数量 <> 0
                             Order By 序号) Loop
                  Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                  If r_序号.序号 = n_号序 Then
                    n_号序 := n_号序 + 1;
                  End If;
                End Loop;
                If n_号序 > n_限号数 Then
                  v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                  Raise Err_Item;
                Else
                  Begin
                    Select 1
                    Into n_存在
                    From 挂号序号状态
                    Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                  Exception
                    When Others Then
                      Insert Into 挂号序号状态
                        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                      Values
                        (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                  End;
                  v_Temp := '<HX>' || n_号序 || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              End If;
            End If;
          End If;
        End If;
      Else
        Begin
          Select 1 Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum < 2;
        Exception
          When Others Then
            n_分时段 := 0;
        End;
        Begin
          Select 1
          Into n_启用合作单位
          From 合作单位安排控制
          Where 限制项目 = v_星期 And 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
        Exception
          When Others Then
            n_启用合作单位 := 0;
        End;
        Begin
          Select 1, a.状态, a.操作员姓名, a.机器名
          Into n_存在, n_状态, v_验证姓名, v_验证机器名
          From 挂号序号状态 A, 挂号安排时段 B
          Where a.号码 = v_号码 And Trunc(a.日期) = Trunc(d_日期) And a.序号 = b.序号 And b.安排id = n_安排id And b.星期 = v_星期 And
                To_Char(b.开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And Rownum < 2;
        Exception
          When Others Then
            n_存在 := 0;
        End;
        If n_存在 = 1 Then
          If n_状态 = 5 And v_验证姓名 = v_操作员姓名 And v_机器名 = v_验证机器名 Then
            Null;
          Else
            --传入时间的序号已经被使用
            v_Temp := '传入时间' || d_日期 || '的序号已被使用';
            Raise Err_Item;
          End If;
        Else
          If n_分时段 = 1 Then
            Begin
              Select 序号
              Into n_号序
              From 挂号安排时段
              Where 安排id = n_安排id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And Rownum < 2;
            Exception
              When Others Then
                Select Max(序号) + 1
                Into n_号序
                From (Select Distinct 序号
                       From 挂号安排时段
                       Where 安排id = n_安排id And 星期 = v_星期
                       Union
                       Select Distinct 序号 From 挂号序号状态 Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期));
            End;
            Begin
              Select 1 Into n_存在 From 挂号序号状态 Where 号码 = v_号码 And 日期 = d_日期 And 序号 = n_号序;
            Exception
              When Others Then
                Insert Into 挂号序号状态
                  (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                Values
                  (v_号码, d_日期, n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
            End;
            v_Temp := '<HX>' || n_号序 || '</HX>';
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Else
            If v_合作单位 Is Null Or n_启用合作单位 = 0 Then
              If Trunc(d_日期) = Trunc(Sysdate) Then
                n_号序 := 1;
                Select 限号数 Into n_限号数 From 挂号安排限制 Where 安排id = n_安排id And 限制项目 = v_星期;
                For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                             From 挂号序号状态
                             Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                             Order By 序号) Loop
                  Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                  If r_序号.序号 = n_号序 Then
                    n_号序 := n_号序 + 1;
                  End If;
                End Loop;
                If n_号序 > n_限号数 Then
                  v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                  Raise Err_Item;
                Else
                  Begin
                    Select 1
                    Into n_存在
                    From 挂号序号状态
                    Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                  Exception
                    When Others Then
                      Insert Into 挂号序号状态
                        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                      Values
                        (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                  End;
                  v_Temp := '<HX>' || n_号序 || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              Else
                n_号序 := 1;
                Select 限号数 Into n_限号数 From 挂号安排限制 Where 安排id = n_安排id And 限制项目 = v_星期;
                For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                             From 挂号序号状态
                             Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                             
                             Union
                             Select 序号, Null, Null, Null
                             From 合作单位安排控制
                             Where 安排id = n_安排id And 限制项目 = v_星期 And 数量 <> 0
                             Order By 序号) Loop
                  Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                  If r_序号.序号 = n_号序 Then
                    n_号序 := n_号序 + 1;
                  End If;
                End Loop;
                If n_号序 > n_限号数 Then
                  v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                  Raise Err_Item;
                Else
                  Begin
                    Select 1
                    Into n_存在
                    From 挂号序号状态
                    Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                  Exception
                    When Others Then
                      Insert Into 挂号序号状态
                        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                      Values
                        (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                  End;
                  v_Temp := '<HX>' || n_号序 || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              End If;
            Else
              Select Count(1)
              Into n_合约模式
              From 合作单位安排控制
              Where 序号 = 0 And 安排id = n_安排id And 合作单位 = v_合作单位 And 限制项目 = v_星期 And 数量 <> 0;
              If n_合约模式 = 0 Then
                Begin
                  Select 序号
                  Into n_号序
                  From (Select 序号
                         From 合作单位安排控制 A
                         Where 安排id = n_安排id And 合作单位 = v_合作单位 And 限制项目 = v_星期 And 数量 <> 0 And
                               (Not Exists
                                (Select 1
                                 From 挂号序号状态
                                 Where 号码 = v_号码 And 序号 = a.序号 And Trunc(日期) = Trunc(d_日期) And 状态 <> 5) Or Exists
                                (Select 1
                                 From 挂号序号状态
                                 Where 号码 = v_号码 And 序号 = a.序号 And Trunc(日期) = Trunc(d_日期) And 状态 = 5 And 操作员姓名 = v_操作员姓名 And
                                       机器名 = v_机器名))
                         Order By 序号)
                  Where Rownum < 2;
                Exception
                  When Others Then
                    n_号序 := 0;
                End;
                If n_号序 = 0 Then
                  v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                  Raise Err_Item;
                Else
                  Begin
                    Select 1
                    Into n_存在
                    From 挂号序号状态
                    Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                  Exception
                    When Others Then
                      Insert Into 挂号序号状态
                        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                      Values
                        (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                  End;
                  v_Temp := '<HX>' || n_号序 || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              Else
                n_号序 := 1;
                Select 限号数 Into n_限号数 From 挂号安排限制 Where 安排id = n_安排id And 限制项目 = v_星期;
                For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                             From 挂号序号状态
                             Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                             
                             Union
                             Select 序号, Null, Null, Null
                             From 合作单位安排控制
                             Where 安排id = n_安排id And 限制项目 = v_星期 And 数量 <> 0
                             Order By 序号) Loop
                  Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                  If r_序号.序号 = n_号序 Then
                    n_号序 := n_号序 + 1;
                  End If;
                End Loop;
                If n_号序 > n_限号数 Then
                  v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                  Raise Err_Item;
                Else
                  Begin
                    Select 1
                    Into n_存在
                    From 挂号序号状态
                    Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                  Exception
                    When Others Then
                      Insert Into 挂号序号状态
                        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                      Values
                        (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                  End;
                  v_Temp := '<HX>' || n_号序 || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Temp || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Lockno;
/

--89530:刘尔旋,2015-10-15,服务窗接口增加站点处理
Create Or Replace Procedure Zl_Third_Getdeptlist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取可挂号科室

  --入参:Xml_In:
  --<IN>
  --  <CXTS>14</CXTS>        //查询天数
  --  <HZDW>支付宝</HZDW>    //合作单位
  --  <ZD></ZD>              //站点
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <KSLIST>
  --  <KS>
  --    <ID>科室ID</ID>       //科室ID
  --    <MC>科室名称</MC>     //科室名称
  --  </KS>
  --  <KS>
  --    ...
  --  </KS>
  -- </KSLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_Temp      Varchar(5000); --临时XML
  x_Templet   Xmltype; --模板XML
  n_查询天数  Number(5);
  n_Add_Lists Number(3);
  v_合作单位  合作单位安排控制.合作单位%Type;
  n_站点      部门表.站点%Type;
  v_Err_Msg   Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  Select Extractvalue(Value(A), 'IN/CXTS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/ZD')
  Into n_查询天数, v_合作单位, n_站点
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_查询天数 := Nvl(n_查询天数, 14);
  If v_合作单位 Is Null Then
    For r_Dept In (Select Distinct a.科室id, b.名称
                   From 挂号安排 A, 部门表 B
                   Where a.停用日期 Is Null And a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null)) Loop
    
      If Nvl(n_Add_Lists, 0) = 0 Then
        --增加DJList节点
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
        n_Add_Lists := 1;
      End If;
      v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
      Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    End Loop;
  Else
    For r_Dept In (Select Distinct 科室id, 名称
                   From (Select b.科室id, d.名称
                          From (Select a.Id
                                 From 挂号安排 A
                                 Where a.停用日期 Is Null And Not Exists
                                  (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                 Union All
                                 Select a.Id
                                 From 挂号安排 A, 合作单位安排控制 B
                                 Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B,
                               挂号安排计划 C, 部门表 D
                          Where a.Id = b.Id And c.安排id = a.Id And c.审核时间 Is Not Null And
                                ((c.生效时间 < Sysdate And c.失效时间 > Sysdate + n_查询天数) Or
                                (c.生效时间 Between Sysdate And Sysdate + n_查询天数) Or
                                (c.失效时间 Between Sysdate And Sysdate + n_查询天数)) And Not Exists
                           (Select 1 From 合作单位计划控制 Where 计划id = c.Id And 合作单位 = v_合作单位 And 数量 = 0) And b.科室id = d.Id And
                                (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null)
                          Union All
                          Select b.科室id, d.名称
                          From (Select a.Id
                                 From 挂号安排 A
                                 Where a.停用日期 Is Null And Not Exists
                                  (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                 Union All
                                 Select a.Id
                                 From 挂号安排 A, 合作单位安排控制 B
                                 Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B, 部门表 D
                          Where a.Id = b.Id And Not Exists
                           (Select 1 From 挂号安排计划 Where 安排id = a.Id And Rownum < 2) And b.科室id = d.Id And
                                (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null))) Loop
      If Nvl(n_Add_Lists, 0) = 0 Then
        --增加DJList节点
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
        n_Add_Lists := 1;
      End If;
      v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
      Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    End Loop;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdeptlist;
/

--89530:刘尔旋,2015-10-15,服务窗接口增加站点处理
Create Or Replace Procedure Zl_Third_Reghistory
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取历史挂号数据
  --入参:Xml_In:
  --<IN>
  --   <BRID>88393</BRID>    //卡号
  --   <JLS>5</JLS >       //记录条数，按日期由近到远
  --   <JSKLB></JSKLB>     //结算卡类别
  --   <ZD></ZD>           //站点
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  --   <GHLIST>       //如果为空表示没有找到数据
  --    <GH>
  --     <GHDH>N001</GHDH>   //挂号单号
  --     <KS>门诊外科</KS>     //挂号科室
  --     <KSID>42</KSID>    //科室id
  --     <DJSJ>2014-10-21 14:12:44</DJSJ>    //登记时间
  --     <YYSJ>2014-10-21 14:10</YYSJ>    //预约时间
  --     <ZXZT>1</ZXZT>     //状态(预约中等待付款、已挂号、候诊等)
  --     <DDFK>1</DDFK>     //是否付款
  --     <GHFS>支付宝</GHFS>    //挂号方式(支付宝、自助机、窗口)
  --     <YSXM>LEX</YSXM>    //医生姓名
  --    </GH>
  --   </GHLIST>
  --   <ERROR><MSG>错误信息</MSG></ERROR>     //如果有错误返回
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_费别         病人信息.费别%Type;
  v_付款方式     病人信息.医疗付款方式%Type;
  v_姓名         病人信息.姓名%Type;
  v_性别         病人信息.性别%Type;
  v_年龄         病人信息.年龄%Type;
  d_出生日期     病人信息.出生日期%Type;
  v_身份证号     病人信息.身份证号%Type;
  v_卡类别       医疗卡类别.名称%Type;
  v_卡号         病人医疗卡信息.卡号%Type;
  n_卡类别id     医疗卡类别.Id%Type;
  v_结算方式     医疗卡类别.结算方式%Type;
  v_验证姓名     病人信息.姓名%Type;
  n_病人id       病人信息.病人id%Type;
  n_卡病人id     病人信息.病人id%Type;
  v_医疗卡编码   医疗卡类别.编码%Type;
  n_是否缺省密码 医疗卡类别.是否缺省密码%Type;
  v_密码         病人医疗卡信息.密码%Type;
  n_密码长度     医疗卡类别.密码长度%Type;
  n_记录数       Number(4);
  v_结算卡类别   Varchar2(100);
  n_是否付款     Number(3);
  v_Temp         Varchar2(32767); --临时XML
  x_Templet      Xmltype; --模板XML
  v_Err_Msg      Varchar2(200);
  n_Exists       Number(2);
  v_状态         Varchar2(100);
  n_站点         Number(1);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT><GHLIST></GHLIST></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/JLS'), Extractvalue(Value(A), 'IN/JSKLB'),
         Extractvalue(Value(A), 'IN/ZD')
  Into n_病人id, n_记录数, v_结算卡类别, n_站点
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_结算卡类别 Is Not Null Then
    Begin
      n_卡类别id := To_Number(v_结算卡类别);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
    If n_卡类别id = 0 Then
      Begin
        Select ID Into n_卡类别id From 医疗卡类别 Where 名称 = v_结算卡类别;
      Exception
        When Others Then
          v_Err_Msg := '无法确认传入的结算卡！';
          Raise Err_Item;
      End;
    End If;
  Else
    n_卡类别id := 0;
  End If;

  n_记录数 := Nvl(n_记录数, 0);

  If Nvl(n_卡类别id, 0) = 0 Then
    For r_Reg In (Select a.No, a.执行部门id As 科室id, b.名称 As 科室, a.登记时间 As 创建时间, a.发生时间 As 预约时间,
                         Decode(a.记录性质, 2, '预约中', Decode(a.执行状态, -1, '不就诊', 0, '等待就诊', 1, '完成就诊', 2, '正在就诊', '已挂号')) As 状态,
                         Decode(c.记录状态, 0, 0, 1) As 是否付款, a.预约方式 As 挂号方式, a.执行人 As 医生姓名
                  From 病人挂号记录 A, 部门表 B, 门诊费用记录 C
                  Where a.病人id = n_病人id And a.执行部门id = b.Id And c.No = a.No And c.序号 = 1 And c.记录性质 = 4 And a.记录状态 = 1 And
                        (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null)
                  Order By a.登记时间 Desc) Loop
      If n_记录数 <> 0 Then
        Begin
          Select Decode(记录状态, 0, 0, 1)
          Into n_是否付款
          From 门诊费用记录
          Where 记录性质 = 1 And 病人id = n_病人id And Instr(摘要, '挂号:' || r_Reg.No) > 0 And Rownum < 2;
        Exception
          When Others Then
            n_是否付款 := r_Reg.是否付款;
        End;
        If n_是否付款 = 0 And r_Reg.状态 = '预约中' And Sysdate > r_Reg.预约时间 Then
          v_状态 := '已失效';
        Else
          v_状态 := r_Reg.状态;
        End If;
        v_Temp := '<GH>' || '<GHDH>' || r_Reg.No || '</GHDH>' || '<KS>' || r_Reg.科室 || '</KS>' || '<KSID>' ||
                  r_Reg.科室id || '</KSID>' || '<DJSJ>' || To_Char(r_Reg.创建时间, 'YYYY-MM-DD hh24:mi:ss') || '</DJSJ>' ||
                  '<YYSJ>' || To_Char(r_Reg.预约时间, 'YYYY-MM-DD hh24:mi:ss') || '</YYSJ>' || '<ZXZT>' || v_状态 ||
                  '</ZXZT>' || '<DDFK>' || n_是否付款 || '</DDFK>' || '<GHFS>' || r_Reg.挂号方式 || '</GHFS>' || '<YSXM>' ||
                  r_Reg.医生姓名 || '</YSXM>' || '</GH>';
        Select Appendchildxml(x_Templet, '/OUTPUT/GHLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        n_记录数 := n_记录数 - 1;
      End If;
    End Loop;
  Else
    For r_Reg In (Select a.No, a.执行部门id As 科室id, b.名称 As 科室, a.登记时间 As 创建时间, a.发生时间 As 预约时间,
                         Decode(a.记录性质, 2, '预约中', Decode(a.执行状态, -1, '不就诊', 0, '等待就诊', 1, '完成就诊', 2, '正在就诊', '已挂号')) As 状态,
                         Decode(c.记录状态, 0, 0, 1) As 是否付款, a.预约方式 As 挂号方式, a.执行人 As 医生姓名
                  From 病人挂号记录 A, 部门表 B, 门诊费用记录 C, 病人预交记录 D
                  Where a.病人id = n_病人id And a.执行部门id = b.Id And c.No = a.No And c.序号 = 1 And c.记录性质 = 4 And a.记录状态 = 1 And
                        c.结帐id = d.结帐id And d.卡类别id = n_卡类别id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null)
                  Order By a.登记时间 Desc) Loop
      If n_记录数 <> 0 Then
        Begin
          Select Decode(记录状态, 0, 0, 1)
          Into n_是否付款
          From 门诊费用记录
          Where 记录性质 = 1 And 病人id = n_病人id And Instr(摘要, '挂号:' || r_Reg.No) > 0 And Rownum < 2;
        Exception
          When Others Then
            n_是否付款 := r_Reg.是否付款;
        End;
        If n_是否付款 = 0 And r_Reg.状态 = '预约中' And Sysdate > r_Reg.预约时间 Then
          v_状态 := '已失效';
        Else
          v_状态 := r_Reg.状态;
        End If;
        v_Temp := '<GH>' || '<GHDH>' || r_Reg.No || '</GHDH>' || '<KS>' || r_Reg.科室 || '</KS>' || '<KSID>' ||
                  r_Reg.科室id || '</KSID>' || '<DJSJ>' || To_Char(r_Reg.创建时间, 'YYYY-MM-DD hh24:mi:ss') || '</DJSJ>' ||
                  '<YYSJ>' || To_Char(r_Reg.预约时间, 'YYYY-MM-DD hh24:mi:ss') || '</YYSJ>' || '<ZXZT>' || v_状态 ||
                  '</ZXZT>' || '<DDFK>' || n_是否付款 || '</DDFK>' || '<GHFS>' || r_Reg.挂号方式 || '</GHFS>' || '<YSXM>' ||
                  r_Reg.医生姓名 || '</YSXM>' || '</GH>';
        Select Appendchildxml(x_Templet, '/OUTPUT/GHLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        n_记录数 := n_记录数 - 1;
      End If;
    End Loop;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Reghistory;
/

--89572:李南春,2015-10-26,调整变量大小
--82859:李南春,2015-10-15,挂号基本信息调整
Create Or Replace Procedure Zl_挂号病人病案_Insert
(
  处理类型_In       Number,
  病人id_In         病人信息.病人id%Type,
  门诊号_In         病人信息.门诊号%Type,
  就诊卡号_In       病人信息.就诊卡号%Type,
  卡验证码_In       病人信息.卡验证码%Type,
  姓名_In           病人信息.姓名%Type,
  性别_In           病人信息.性别%Type,
  年龄_In           病人信息.年龄%Type,
  费别_In           病人信息.费别%Type,
  医疗付款方式_In   病人信息.医疗付款方式%Type,
  国籍_In           病人信息.国籍%Type,
  民族_In           病人信息.民族%Type,
  婚姻_In           病人信息.婚姻状况%Type,
  职业_In           病人信息.职业%Type,
  身份证号_In       病人信息.身份证号%Type,
  工作单位_In       病人信息.工作单位%Type,
  合同单位id_In     病人信息.合同单位id%Type,
  单位电话_In       病人信息.单位电话%Type,
  单位邮编_In       病人信息.单位邮编%Type,
  家庭地址_In       病人信息.家庭地址%Type,
  家庭电话_In       病人信息.家庭电话%Type,
  家庭地址邮编_In   病人信息.家庭地址邮编%Type,
  登记时间_In       病人信息.登记时间%Type,
  挂号单_In         病人挂号记录.No%Type := Null,
  出生日期_In       病人信息.出生日期%Type := Null,
  医保号_In         病人信息.医保号%Type := Null,
  Ic卡号_In         病人信息.Ic卡号%Type := Null,
  险类_In           病人信息.险类%Type := Null,
  区域_In           病人信息.区域%Type := Null,
  户口地址_In       病人信息.户口地址%Type := Null,
  户口地址邮编_In   病人信息.户口地址邮编%Type := Null,
  联系人身份证号_In In 病人信息.联系人身份证号%Type := Null,
  联系人姓名_In     In 病人信息.联系人姓名%Type := Null,
  联系人电话_In     In 病人信息.联系人电话%Type := Null,
  联系人关系_In     In 病人信息.联系人关系%Type := Null,
  监护人_In         In 病人信息.监护人%Type := Null,
  出生地点_In       In 病人信息.出生地点%Type := Null
) As
  --功能：处理挂号病人病案信息
  --参数：
  --处理类型：
  --             1=新建病人信息及门诊病案(用于新挂号病人)
  --             2=修改病人信息，新建门诊病案(用于无病案的病人)
  --             3=修改病人信息，不处理门诊病案(用于有病案的病人,但可能修改了病案的门诊号)
  v_年龄     Varchar2(20);
  v_年龄单位 Varchar2(20);
  v_出生日期 Date;
  n_一卡通   Number(1);
  v_Username 人员表.姓名%Type;
  v_姓名信息     病人信息.姓名%Type;
  v_年龄信息     病人信息.年龄%Type;
  v_性别信息     病人信息.性别%Type;
  v_身份证号     病人信息.身份证号%Type;
  d_出生日期信息 Date;
  d_变动时间 Date;
Begin
  If 出生日期_In Is Null And 年龄_In Is Not Null Then
    --根据年龄求出生日期
    v_年龄单位 := Substr(年龄_In, Length(年龄_In), 1);
    If Instr('岁,月,天', v_年龄单位) <= 0 Then
      v_年龄单位 := Null;
    Else
      v_年龄 := Replace(年龄_In, v_年龄单位, '');
    End If;
    Begin
      v_年龄 := To_Number(v_年龄);
    Exception
      When Others Then
        v_年龄 := Null;
    End;
    If v_年龄 Is Not Null And v_年龄单位 Is Not Null Then
      Select Decode(v_年龄单位, '岁', Add_Months(Sysdate, -12 * v_年龄), '月', Add_Months(Sysdate, -1 * v_年龄), '天',
                     Sysdate - v_年龄)
      Into v_出生日期
      From Dual;
    End If;
  Else
    v_出生日期 := 出生日期_In;
  End If;
  
  Begin
      Select 姓名, 性别, 年龄, 出生日期,身份证号 Into v_姓名信息, v_性别信息, v_年龄信息, d_出生日期信息,v_身份证号 From 病人信息 Where 病人id = 病人id_In;
  Exception
      When Others Then
      v_姓名信息:=姓名_In;
      v_性别信息:=性别_In;
      v_年龄信息:=年龄_In;
      d_出生日期信息:=v_出生日期;
      v_身份证号:=身份证号_In;
  End;
  
  Begin
    Select 1 Into n_一卡通 From 一卡通目录 Where 启用 = 2 And Rownum <= 1;
  Exception
    When Others Then
      n_一卡通 := 0;
  End;

  If Not 就诊卡号_In Is Null And n_一卡通 = 0 Then
    Update 病人信息
    Set 就诊卡号 = Null, Ic卡号 = Decode(Ic卡号, 就诊卡号, Null, Ic卡号)
    Where 病人id <> 病人id_In And 就诊卡号 = 就诊卡号_In;
  End If;

  If 处理类型_In = 1 Then
    --新病人信息
    Insert Into 病人信息
      (病人id, 门诊号, 就诊卡号, 卡验证码, Ic卡号, 姓名, 性别, 年龄, 出生日期, 费别, 医疗付款方式, 国籍, 民族, 婚姻状况, 职业, 身份证号, 工作单位, 合同单位id, 单位电话, 单位邮编,
       家庭地址, 家庭电话, 家庭地址邮编, 户口地址, 户口地址邮编, 登记时间, 医保号, 区域, 联系人身份证号, 联系人姓名, 联系人电话, 联系人关系, 监护人, 出生地点)
    Values
      (病人id_In, 门诊号_In, 就诊卡号_In, 卡验证码_In, Ic卡号_In, 姓名_In, 性别_In, 年龄_In, v_出生日期, 费别_In, 医疗付款方式_In, 国籍_In, 民族_In, 婚姻_In,
       职业_In, 身份证号_In, 工作单位_In, Decode(合同单位id_In, 0, Null, 合同单位id_In), 单位电话_In, 单位邮编_In, 家庭地址_In, 家庭电话_In, 家庭地址邮编_In,
       户口地址_In, 户口地址邮编_In, 登记时间_In, 医保号_In, 区域_In, 联系人身份证号_In, 联系人姓名_In, 联系人电话_In, 联系人关系_In, 监护人_In, 出生地点_In);
  Elsif 处理类型_In In (2, 3) Then
    --宁波一卡通,档案号以IC卡号传入,此时不更新就诊卡号
    Update 病人信息
    Set 门诊号 = 门诊号_In, 就诊卡号 = Decode(n_一卡通, 1, Decode(Ic卡号_In, Null, Nvl(就诊卡号_In, 就诊卡号), 就诊卡号), Nvl(就诊卡号_In, 就诊卡号)),
        Ic卡号 = Decode(n_一卡通, 1, Decode(Ic卡号_In, Null, Ic卡号, Nvl(Ic卡号_In, Ic卡号)), Nvl(Ic卡号_In, Ic卡号)),
        卡验证码 = Decode(n_一卡通, 1, Decode(Ic卡号_In, Null, Nvl(卡验证码_In, 卡验证码), 卡验证码),
                       Decode(就诊卡号_In, Null, 卡验证码, Nvl(卡验证码_In, 卡验证码))), 姓名 = 姓名_In, 性别 = Nvl(性别_In, 性别), 年龄 = 年龄_In,
        出生日期 = v_出生日期, 费别 = Nvl(费别_In, 费别), 医疗付款方式 = Nvl(医疗付款方式_In, 医疗付款方式), 国籍 = Nvl(国籍_In, 国籍), 民族 = Nvl(民族_In, 民族),
        婚姻状况 = Nvl(婚姻_In, 婚姻状况), 职业 = Nvl(职业_In, 职业), 身份证号 = 身份证号_In, 工作单位 = 工作单位_In,
        合同单位id = Decode(合同单位id_In, 0, Null, 合同单位id_In), 单位电话 = 单位电话_In, 单位邮编 = 单位邮编_In, 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In,
        户口地址 = Nvl(户口地址_In,户口地址), 户口地址邮编 = Nvl(户口地址邮编_In,户口地址邮编), 家庭地址邮编 = 家庭地址邮编_In, 医保号 = 医保号_In, 险类 = 险类_In, 区域 = Nvl(区域_In, 区域),
        联系人身份证号 = 联系人身份证号_In, 联系人姓名 = 联系人姓名_In, 联系人电话 = 联系人电话_In, 联系人关系 = 联系人关系_In, 监护人 = 监护人_In, 出生地点 = 出生地点_In
    Where 病人id = 病人id_In;
    
    v_Username := Zl_Username;
    d_变动时间 := Sysdate;
    if Nvl(病人id_In,0) > 0 then
      If Nvl(姓名_In, '') <> Nvl(v_姓名信息, '') Then
        Insert Into 病人信息变动
          (病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明)
        Values
          (病人id_In, '姓名', v_姓名信息, 姓名_In, d_变动时间, v_Username, '挂号', '病人基本信息调整');
      End If;
      If Nvl(身份证号_In, '') <> Nvl(v_身份证号, '') Then
        Insert Into 病人信息变动
          (病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明)
        Values
          (病人id_In, '身份证号', v_身份证号, 身份证号_In, d_变动时间, v_Username, '挂号', '病人基本信息调整');
      End If;
      If Nvl(性别_In, '') <> Nvl(v_性别信息, '') Then
        Insert Into 病人信息变动
          (病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明)
        Values
          (病人id_In, '性别', v_性别信息, 性别_In, d_变动时间, v_Username, '挂号', '病人基本信息调整');
      End If;
      If Nvl(v_出生日期, Sysdate) <> Nvl(d_出生日期信息, Sysdate) Then
        If Nvl(年龄_In, '') <> Nvl(v_年龄信息, '') Then
          Insert Into 病人信息变动
            (病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明)
          Values
            (病人id_In, '年龄', v_年龄信息, 年龄_In, d_变动时间, v_Username, '挂号', '病人基本信息调整');
        End If;
        Insert Into 病人信息变动
          (病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明)
        Values
          (病人id_In, '出生日期',  to_char(d_出生日期信息,'YYYY-MM-DD hh24:mi'),  to_char(v_出生日期,'YYYY-MM-DD hh24:mi'), d_变动时间, v_Username, '挂号', '病人基本信息调整');
      End If;
     end if;
     
    --北京医保:问题:26982
    If Nvl(险类_In, 0) = 920 And Not 医保号_In Is Null Then
      --需要反更新
      Update 医保病人档案 B
      Set 医保号 = 医保号_In
      Where (险类, 中心, 医保号) = (Select 险类, 中心, 医保号
                             From 医保病人关联表 A
                             Where 险类 = 险类_In And a.病人id = 病人id_In And a.医保号 <> 医保号_In And Rownum = 1);
      Update 医保病人关联表 Set 医保号 = 医保号_In Where 病人id = 病人id_In And 医保号 <> 医保号_In And 险类 = 险类_In;
    End If;

    If 挂号单_In Is Not Null Then

      Update 病人挂号记录
      Set 门诊号 = 门诊号_In, 姓名 = 姓名_In, 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
      Where NO = 挂号单_In;

      Update 门诊费用记录
      Set 标识号 = 门诊号_In, 姓名 = 姓名_In, 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
      Where NO = 挂号单_In And 记录性质 = 4;
    End If;
  End If;

  --门诊病案
  If 门诊号_In Is Not Null Then
    If 处理类型_In In (1, 2) Then
      Update 门诊病案记录 Set 病案号 = 门诊号_In Where 病人id = 病人id_In;
      If Sql%RowCount = 0 Then
        Insert Into 门诊病案记录
          (病人id, 病案号, 建立日期, 病案类别, 存储状态, 存放位置)
        Values
          (病人id_In, 门诊号_In, 登记时间_In, '一般', Null, Null);
      End If;
    Elsif 处理类型_In = 3 Then
      Update 门诊病案记录 Set 病案号 = 门诊号_In Where 病人id = 病人id_In;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号病人病案_Insert;
/

--89427:冉俊明,2015-10-13,退费时，收回票据信息显示为重打收回，实际应为作废收回。
Create Or Replace Procedure Zl_Custom_Invoice_Autoallot
(
  操作类型_In   Number,
  模拟计算_In   Number,
  票种_In       票据使用明细.票种%Type,
  领用id_In     票据使用明细.领用id%Type,
  病人id_In     门诊费用记录.病人id%Type,
  Nos_In        Varchar2,
  起始发票号_In 门诊费用记录.实际票号%Type,
  使用人_In     票据使用明细.使用人%Type,
  使用时间_In   票据使用明细.使用时间%Type,
  发票号_In     In Out Varchar2,
  发票张数_In   Out Number
) As
  -------------------------------------------------------------------------------------------------------------
  --功能：根据票据分配规则,自动分配票据明细数据
  --入参：
  --     操作类型_In :1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
  --     模拟计算_IN :0-不进行模拟计算;1-进行模拟计算,模拟计算时不保存数据
  --     票种_IN     :1-门诊收费;暂无其他类型票据
  --     病人ID_IN   :病人ID,如果Nos和发票号_In为空时,表示针对该病人的所有未打印的票据进行打印
  --     NOs_IN      :单据号,多个用逗号分离,最多有400张单据,格式为:A00001,A00002.....
  --     退费NOs:退费所涉及的单据
  --     启始发票号_IN:重打票据或发出票据的启始票号;
  --     发票号_In   :可以为多个,用逗号分隔,当操作类型为3-重打票据时和4-退费回收票据有效,表示本次需要回收的票据
  --出参:
  --     发票号_In   :可以为多个,用逗号分隔,当操作类型为3-重打票据时和4-退费回收票据有效,表示重打或退费重新发出的票据
  --     发票张数_IN :返回本次收费所需要的发票张数
  -------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_分单据打印 Number(3);
  n_执行科室   Number(3);
  n_收据费目   Number(3);
  n_汇总条件   Number(3);
  n_收费细目   Number(3);

  --------------------------------------------------------
  --定义内部票据处理的数据集
  Type Ty_Rec_Bill Is Record(
    票号     票据打印明细.票号%Type,
    NO       票据打印明细.No%Type,
    序号     票据打印明细.序号%Type,
    关联序号 票据打印明细.关联票号序号%Type,
    修改标志 Number(1));
  Type Ty_Tb_Bill Is Table Of Ty_Rec_Bill;
  c_Invoce Ty_Tb_Bill := Ty_Tb_Bill();
  --------------------------------------------------------
  --按元素1,元素2,元素3,元素4,分别统计各单据的序号
  Type Ty_Rec_No Is Record(
    NO   门诊费用记录.No%Type,
    序号 Varchar2(1000));
  Type Ty_Tb_No Is Table Of Ty_Rec_No;
  c_No Ty_Tb_No := Ty_Tb_No();
  --------------------------------------------------------
  Cursor c_Fact Is
    Select 前缀文本, 剩余数量, 开始号码, 终止号码, 当前号码 From 票据领用记录 Where ID = Nvl(领用id_In, 0);
  r_Factrow c_Fact%RowType;

  v_Nos        Varchar2(4000);
  v_发票号     票据打印明细.票号%Type;
  v_开始发票号 票据打印明细.票号%Type;
  v_当前发票号 票据打印明细.票号%Type;
  v_回收票据号 Varchar2(4000);
  n_Find       Number(3);

  n_元素1_Count Number(3);
  n_元素2_Count Number(3);
  n_元素3_Count Number(3);
  n_元素4_Count Number(3);

  v_元素1    门诊费用记录.No%Type;
  n_元素2    门诊费用记录.执行部门id%Type;
  v_元素3    门诊费用记录.收据费目%Type;
  n_元素4    门诊费用记录.收费细目id%Type;
  v_发票信息 Varchar2(4000);
  n_误差项   Number(1);
  n_打印id   票据使用明细.打印id%Type;
  n_使用id   票据使用明细.Id%Type;
  n_返回数   Number(18);
  n_关联序号 Number(18);
  r_单据号   t_Strlist := t_Strlist();
  r_单据序号 t_Strlist := t_Strlist();
  l_使用id   t_Numlist := t_Numlist();
  l_关联序号 t_Numlist := t_Numlist();

  v_打印内容 Varchar2(4000);
  v_Temp     Varchar2(4000);
  Procedure Invoice_Split_Notgroup
  (
    收费nos_In       Varchar2,
    回收发票_In      Varchar2,
    本次打印发票_Out In Out Varchar2,
    本次发票张数_Out In Out Number,
    Invoce_Out       In Out Ty_Tb_Bill
  ) As
    ----------------------------------------------------------------------------
    --入参:
    --   收费收费NOs_IN:本次需要处理的发票所涉及的单据,多个用逗号分离
    --   回收发票_IN-退费时有效,多个用逗号分离，表示本次需要回收的发票号 
    --出参:
    -- 本次打印发票_Out-本次需要的发票号,多个用逗号分离
    -- 本次发票张数_Out-本次需要的发票数
    -- Invoce_Out:本次返回的发票号与单据的对应关系
    n_Count Number(18);
    n_分页  Number(18);
  
    Cursor Cr_Bill Is
      Select NO As 元素1, 执行部门id As 元素2, 收据费目 As 元素3, NO As 元素4, NO As 单据, 序号, 0 As 个数
      From 门诊费用记录
      Where Rownum <= 1;
    c_Bill Cr_Bill%RowType;
    --------------------------------------------------------------------------------------------
    --根据相关传入的数据,取对应的数据集
    Type Ty_费用明细 Is Ref Cursor;
    c_费用明细 Ty_费用明细; --游标变量 
  
  Begin
    --按单据分配票据
    If 操作类型_In = 3 Or 操作类型_In = 4 Then
      --1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
      Open c_费用明细 For
        With c_费用 As
         (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, '-', a.No) As 元素4, a.No As 单据,
                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
          From 门诊费用记录 A,
               (Select NO, 序号 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J Where m.票号 = j.Column_Value) B
          Where Mod(a.记录性质, 10) = 1 And a.No = b.No And Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
          Having Sum(Nvl(a.实收金额, 0)) <> 0)
        Select 元素1, 元素2, 元素3, 元素4, 单据, 序号, Count(*) As 个数
        From c_费用
        Group By 元素1, 元素2, 元素3, 元素4, 单据, 序号
        Order By 元素1, 元素2, 元素3, 元素4, 单据, 序号;
    Else
      Open c_费用明细 For
        With c_费用 As
         (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, '-', a.No) As 元素4, a.No As 单据,
                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
          From 门诊费用记录 A, Table(f_Str2list(收费nos_In)) B
          Where Mod(a.记录性质, 10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
          Having Sum(Nvl(a.实收金额, 0)) <> 0)
        Select 元素1, 元素2, 元素3, 元素4, 单据, 序号, Count(*) As 个数
        From c_费用
        Group By 元素1, 元素2, 元素3, 元素4, 单据, 序号
        Order By 元素1, 元素2, 元素3, 元素4, 单据, 序号;
    End If;
  
    v_元素1          := '+';
    n_元素2          := 0;
    v_元素3          := '+';
    n_元素4          := 0;
    n_元素1_Count    := 0;
    n_元素2_Count    := 0;
    n_元素3_Count    := 0;
    n_元素4_Count    := 0;
    本次发票张数_Out := 0;
    If n_汇总条件 <> 0 Then
      n_关联序号 := 1;
    Else
      n_关联序号 := 0;
    End If;
    n_Count := 0;
    c_No.Delete;
    Loop
      Fetch c_费用明细
        Into c_Bill;
      Exit When c_费用明细%NotFound;
      n_Count := 1;
    
      n_分页 := 0;
      If (v_元素1 <> c_Bill.元素1) Or (n_元素2 <> c_Bill.元素2 And n_元素2_Count >= n_执行科室 And n_执行科室 <> 0) Or
         (v_元素3 <> c_Bill.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Or (n_元素4_Count >= n_收费细目 And n_收费细目 <> 0) Then
      
        If (v_元素1 <> '+' Or n_元素2 <> 0 Or v_元素3 <> '+' Or n_元素4 <> 0) Then
          n_分页 := 1;
        End If;
        n_元素2_Count := 0;
        n_元素3_Count := 0;
        n_元素4_Count := 0;
        n_元素1_Count := 0;
        v_元素1       := '+';
        n_元素2       := 0;
        v_元素3       := '+';
      End If;
    
      If n_分页 = 1 Then
        --分页:计算发票号及相关的
        For I In 1 .. c_No.Count Loop
          Invoce_Out.Extend;
          Invoce_Out(Invoce_Out.Count).票号 := v_发票号;
          Invoce_Out(Invoce_Out.Count).No := c_No(I).No;
          Invoce_Out(Invoce_Out.Count).序号 := Case
                                               When Instr(c_No(I).序号, ',') > 0 Then
                                                Substr(c_No(I).序号, 2)
                                               Else
                                                c_No(I).序号
                                             End;
          Invoce_Out(Invoce_Out.Count).关联序号 := n_关联序号;
        End Loop;
      
        本次发票张数_Out := 本次发票张数_Out + 1;
        本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
        v_发票号         := Zl_Incstr(v_发票号);
        c_No.Delete;
      End If;
      If (v_元素1 <> c_Bill.元素1) Then
        n_元素1_Count := n_元素1_Count + 1;
        v_元素1       := c_Bill.元素1;
      End If;
      If (n_元素2 <> c_Bill.元素2) Then
        n_元素2_Count := n_元素2_Count + 1;
        n_元素2       := c_Bill.元素2;
      End If;
      If (v_元素3 <> c_Bill.元素3) Then
        n_元素3_Count := n_元素3_Count + 1;
        v_元素3       := c_Bill.元素3;
      End If;
      If n_收费细目 <> 0 Then
        n_元素4_Count := n_元素4_Count + 1;
      End If;
    
      -------------------------------------------
      --分配单据号及序号
      n_Find := 0;
      For J In 1 .. c_No.Count Loop
        If c_No(J).No = c_Bill.单据 Then
          --单据号相同,将序号合并
          c_No(J).序号 := c_No(J).序号 || ',' || c_Bill.序号;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If n_Find = 0 Then
        c_No.Extend;
        c_No(c_No.Count).No := c_Bill.单据;
        c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_Bill.序号;
      End If;
    End Loop;
  
    --是否有发票数据
    If n_Count >= 1 Then
      --最后一个发票分配
      本次发票张数_Out := 本次发票张数_Out + 1;
      本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
    Else
      本次发票张数_Out := 0;
      本次打印发票_Out := '';
    End If;
    If c_No.Count <> 0 Then
      For I In 1 .. c_No.Count Loop
        Invoce_Out.Extend;
        Invoce_Out(Invoce_Out.Count).票号 := v_发票号;
        Invoce_Out(Invoce_Out.Count).No := c_No(I).No;
        If Instr(c_No(I).序号, ',') > 0 Then
          c_No(I).序号 := Substr(c_No(I).序号, 2);
        End If;
        Invoce_Out(Invoce_Out.Count).序号 := c_No(I).序号;
        Invoce_Out(Invoce_Out.Count).关联序号 := n_关联序号;
      End Loop;
      c_No.Delete;
    End If;
    If Instr(Nvl(本次打印发票_Out, '-'), ',') > 0 Then
      本次打印发票_Out := Substr(本次打印发票_Out, 2);
    End If;
  End Invoice_Split_Notgroup;

Begin
  --处理票据数据
  If 票种_In <> 1 Then
    --暂不支持其他,只支持门诊收费
    Return;
  End If;
  v_发票号 := 起始发票号_In;
  v_Nos    := Nos_In;
  -----------------------------------------------------------------------------------------------------------------------------
  --一、获取发票分配的相关规则
  --**开始
  --1.确定是否分单据分配票号,缺省不按单据分号
  n_分单据打印 := 0;
  --2.确定是否按执行科室分单据号,缺省为按1个执行科室分号
  n_执行科室 := 1;

  --3.确定是否按收据费目分单据号,缺省为按3个收据费目分号
  n_收据费目 := 3;

  --4.确定是否按收费细目分单据号,缺省为不按收费细目分号
  n_收费细目 := 0;

  --5.决定是否首页汇总,缺省为不汇总
  n_汇总条件 := 0;

  v_回收票据号 := 发票号_In;
  发票张数_In  := 0;
  --**结束
  -----------------------------------------------------------------------------------------------------------------------------
  --二、进行发票分配
  Invoice_Split_Notgroup(Nos_In, 发票号_In, v_发票信息, 发票张数_In, c_Invoce);

  -----------------------------------------------------------------------------------------------------------------------------
  --*****************************************************************************************************************************
  --注意:
  --以下代码，不轻意更改,在上面的代码中需要确定两个变量的值:一是v_发票信息;二是c_Invoce集合中的值
  --  v_发票信息:本次所涉及的发票信息,多个用逗号分离,最好按升序排序
  --  c_Invoce:为集合数据，为发票号和单据的对应关系

  发票号_In := v_发票信息;
  If 模拟计算_In = 1 Then
    --模拟计算,只返回票据张数和使用的票据号,直接退出
    Return;
  End If;
  -----------------------------------------------------------------------------------------------------------------------------
  --四、退费时，需要先处理回收发票
  v_开始发票号 := Null;
  v_当前发票号 := Null;
  --1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
  If 操作类型_In = 3 Or 操作类型_In = 4 Then
    --收回票据
    Select 使用id Bulk Collect
    Into l_使用id
    From (Select Distinct b.使用id
           From 票据使用明细 A, 票据打印明细 B, Table(f_Str2list(v_回收票据号)) J
           Where a.Id = b.使用id And b.票号 = j.Column_Value And Nvl(b.票种, 0) = 1);
  
    --插入回收记录
    Forall I In 1 .. l_使用id.Count
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
        Select 票据使用明细_Id.Nextval, 票种, 号码, 2, Decode(操作类型_In, 3, 4, 2), 领用id, 打印id, 使用人_In, 使用时间_In
        From 票据使用明细
        Where ID = l_使用id(I);
    Forall I In 1 .. l_使用id.Count
      Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I);
  End If;

  If c_Invoce.Count = 0 Then
    --无发票数据,则直接返回,退费时，表示只收回票据
    Return;
  End If;

  -----------------------------------------------------------------------------------------------------------------------------
  --五、重新处理发出的票据(含退费重新发出的票据处理)
  If 起始发票号_In Is Null Then
    v_Err_Msg := '未传入起始发票号,不能进行票据分配处理';
    Raise Err_Item;
  End If;

  If Nvl(领用id_In, 0) <> 0 Then
    Open c_Fact;
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Err_Msg := '无效的票据领用批次，无法完成收费票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    Elsif Nvl(r_Factrow.剩余数量, 0) < 发票张数_In Then
      v_Err_Msg := '当前批次的剩余数量不足' || 发票张数_In || '张，无法完成收费票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    End If;
  End If;

  --1.实际处理票据信息
  If Nvl(n_分单据打印, 0) <> 1 Then
    --不分单据打印时,表示一次打印,打印ID填成一致
    Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  End If;

  发票张数_In := 0;
  v_打印内容  := '';
  For c_Invoce_No In (Select Column_Value As 发票号 From Table(f_Str2list(v_发票信息)) Order By 发票号) Loop
    --2.检查票据范围是否正确
    If Nvl(领用id_In, 0) <> 0 Then
      If Not (Upper(c_Invoce_No.发票号) >= Upper(r_Factrow.开始号码) And Upper(c_Invoce_No.发票号) <= Upper(r_Factrow.终止号码) And
          Length(c_Invoce_No.发票号) = Length(r_Factrow.终止号码)) Then
        v_Err_Msg := '该单据需要打印多张票据,但票据号"' || c_Invoce_No.发票号 || '"超出票据领用的号码范围！';
        Close c_Fact;
        Raise Err_Item;
      End If;
    End If;
  
    --3.处理票据打印明细
    r_单据号.Delete;
    r_单据序号.Delete;
    l_关联序号.Delete;
  
    Select 票据使用明细_Id.Nextval Into n_使用id From Dual;
  
    n_关联序号 := 0;
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).票号 = c_Invoce_No.发票号 Then
        n_关联序号 := c_Invoce(I).关联序号;
        Exit;
      End If;
    End Loop;
  
    --处理关联票据,以便回收票据
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).关联序号 = n_关联序号 And Nvl(c_Invoce(I).修改标志, 0) = 0 Then
        If n_关联序号 <> 0 Then
          c_Invoce(I).关联序号 := n_使用id;
        End If;
        c_Invoce(I).修改标志 := 1;
      End If;
    End Loop;
  
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).票号 = c_Invoce_No.发票号 Then
        r_单据号.Extend;
        r_单据号(r_单据号.Count) := c_Invoce(I).No;
        r_单据序号.Extend;
        r_单据序号(r_单据序号.Count) := c_Invoce(I).序号;
        l_关联序号.Extend;
        If Nvl(c_Invoce(I).关联序号, 0) <> 0 Then
          --检查是否存在其他的票据
          n_Find := 0;
          For J In 1 .. c_Invoce.Count Loop
            If c_Invoce(I).关联序号 = c_Invoce(J).关联序号 And c_Invoce(I).票号 <> c_Invoce(J).票号 Then
              n_Find := 1;
              Exit;
            End If;
          End Loop;
        
          If n_Find = 0 Then
            l_关联序号(l_关联序号.Count) := Null;
            c_Invoce(I).关联序号 := 0;
          Else
            l_关联序号(l_关联序号.Count) := c_Invoce(I).关联序号;
          End If;
        Else
          l_关联序号(l_关联序号.Count) := Null;
        End If;
      End If;
    End Loop;
  
    --1.处理门打印内容
    If n_分单据打印 = 1 Then
      --分单据打印,需按单据进行处理
      --票据打印内容
      n_Find := 0;
      v_Temp := '';
      For I In 1 .. r_单据号.Count Loop
        v_Temp := v_Temp || ',' || r_单据号(I);
        If Instr(Nvl(v_打印内容, '-') || ',', ',' || r_单据号(I) || ',') > 0 Then
          --已经找到
          n_Find := 1;
        End If;
      End Loop;
      v_打印内容 := v_打印内容 || Nvl(v_Temp, '+');
    
      If Nvl(n_Find, 0) = 0 Then
        Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
        Forall I In 1 .. r_单据号.Count
          Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 1, r_单据号(I));
        --以便更新门诊费用记录中的实际票号
        v_开始发票号 := c_Invoce_No.发票号;
        Forall I In 1 .. r_单据号.Count
          Update 门诊费用记录 Set 实际票号 = v_开始发票号 Where Mod(记录性质, 10) = 1 And NO = r_单据号(I);
      End If;
    Else
    
      If v_开始发票号 Is Null Then
        --以便更新门诊费用记录中的实际票号
        v_开始发票号 := c_Invoce_No.发票号;
      
        --票据打印内容
        Insert Into 票据打印内容
          (ID, 数据性质, NO)
          Select n_打印id, 1, Column_Value From Table(f_Str2list(v_Nos));
      
        Update 门诊费用记录
        Set 实际票号 = v_开始发票号
        Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(f_Str2list(v_Nos)));
      End If;
    End If;
    --2.处理票据打印明细
  
    发票张数_In := 发票张数_In + 1;
    --处理票据使用明细
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
    Values
      (n_使用id, 1, c_Invoce_No.发票号, 1, Decode(操作类型_In, 3, 3, 1), Decode(Nvl(领用id_In, 0), 0, Null, 领用id_In), n_打印id, 使用人_In, 使用时间_In);
  
    Forall I In 1 .. r_单据号.Count
      Insert Into 票据打印明细
        (使用id, 票种, 是否回收, NO, 票号, 序号, 关联票号序号)
      Values
        (n_使用id, 1, 0, r_单据号(I), c_Invoce_No.发票号, r_单据序号(I), l_关联序号(I));
  
    v_当前发票号 := c_Invoce_No.发票号;
  End Loop;

  If Nvl(领用id_In, 0) <> 0 Then
    Close c_Fact;
    Update 票据领用记录
    Set 使用时间 = 使用时间_In, 当前号码 = v_当前发票号, 剩余数量 = Nvl(剩余数量, 0) - 发票张数_In
    Where ID = 领用id_In
    Returning 剩余数量 Into n_返回数;
    If n_返回数 < 0 Then
      v_Err_Msg := '当前批次的剩余数量不足' || 发票张数_In || '张，无法完成收费票据分配操作。';
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

--89427:冉俊明,2015-10-13,退费时，收回票据信息显示为重打收回，实际应为作废收回。
Create Or Replace Procedure Zl_Invoice_Autoallot
(
  操作类型_In   Number,
  模拟计算_In   Number,
  票种_In       票据使用明细.票种%Type,
  领用id_In     票据使用明细.领用id%Type,
  病人id_In     门诊费用记录.病人id%Type,
  Nos_In        Varchar2,
  起始发票号_In 门诊费用记录.实际票号%Type,
  使用人_In     票据使用明细.使用人%Type,
  使用时间_In   票据使用明细.使用时间%Type,
  发票号_In     In Out Varchar2,
  发票张数_In   Out Number
) As
  ---------------------------------------------------------------------------------------------
  --功能：根据票据分配规则,自动分配票据明细数据
  --入参：
  --     操作类型_In :1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
  --     票种_IN     :1-门诊收费;暂无其他类型票据
  --     病人ID_IN   :病人ID,如果Nos和发票号_In为空时,表示针对该病人的所有未打印的票据进行打印
  --     NOs_IN      :单据号,多个用逗号分离,最多;有400张单据,格式为:A00001,A00002.....
  --     启始发票号_IN:重打票据或发出票据的启始票号;
  --     发票号_In   :可以为多个,当操作类型为3-重打票据时,有效
  --     模拟计算_IN :0-不进行模拟计算;1-进行模拟计算,模拟计算时不保存数据
  --出参:
  --     发票张数_IN :返回本次收费所需要的发票张数
  ---------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  v_Para       Varchar2(1000);
  v_Temp       Varchar2(4000);
  n_启用模式   Number(3);
  n_分单据打印 Number(3);
  n_执行科室   Number(3);
  n_收据费目   Number(3);
  n_汇总条件   Number(3);
  n_收费细目   Number(3);

  ---------------------------------------------------------
  Type Ty_Rec_Splitno Is Record(
    元素1    票据打印明细.No%Type,
    元素2集  Varchar2(4000),
    元素3集  Varchar2(4000),
    关联序号 Number(18));

  Type Ty_Tb_Splitno Is Table Of Ty_Rec_Splitno;
  c_Split_No   Ty_Tb_Splitno := Ty_Tb_Splitno();
  c_Split_费目 Ty_Tb_Splitno := Ty_Tb_Splitno();

  --------------------------------------------------------
  --定义内部票据处理的数据集
  Type Ty_Rec_Bill Is Record(
    票号     票据打印明细.票号%Type,
    NO       票据打印明细.No%Type,
    序号     票据打印明细.序号%Type,
    关联序号 票据打印明细.关联票号序号%Type,
    修改标志 Number(1));
  Type Ty_Tb_Bill Is Table Of Ty_Rec_Bill;
  c_Invoce Ty_Tb_Bill := Ty_Tb_Bill();
  --------------------------------------------------------
  --按元素1,元素2,元素3,元素4,分别统计各单据的序号
  Type Ty_Rec_No Is Record(
    NO   门诊费用记录.No%Type,
    序号 Varchar2(1000));
  Type Ty_Tb_No Is Table Of Ty_Rec_No;
  c_No Ty_Tb_No := Ty_Tb_No();
  --------------------------------------------------------
  Cursor c_Fact Is
    Select 前缀文本, 剩余数量, 开始号码, 终止号码, 当前号码 From 票据领用记录 Where ID = Nvl(领用id_In, 0);
  r_Factrow c_Fact%RowType;

  --------------------------------------------------------------------------------------------
  --根据相关传入的数据,取对应的数据集

  v_Nos        Varchar2(4000);
  v_发票号     票据打印明细.票号%Type;
  v_开始发票号 票据打印明细.票号%Type;
  v_当前发票号 票据打印明细.票号%Type;
  v_回收票据号 Varchar2(4000);
  n_Find       Number(3);

  n_元素1_Count Number(3);
  n_元素2_Count Number(3);
  n_元素3_Count Number(3);
  n_元素4_Count Number(3);

  v_元素1    门诊费用记录.No%Type;
  n_元素2    门诊费用记录.执行部门id%Type;
  v_元素3    门诊费用记录.收据费目%Type;
  n_元素4    门诊费用记录.收费细目id%Type;
  v_发票信息 Varchar2(4000);
  n_误差项   Number(1);
  n_打印id   票据使用明细.打印id%Type;
  n_使用id   票据使用明细.Id%Type;
  n_返回数   Number(18);
  n_关联序号 Number(18);
  r_单据号   t_Strlist := t_Strlist();
  r_单据序号 t_Strlist := t_Strlist();
  l_使用id   t_Numlist := t_Numlist();
  l_关联序号 t_Numlist := t_Numlist();

  v_打印内容   Varchar2(4000);
  l_元素2      t_Numlist := t_Numlist();
  l_元素3      t_Strlist := t_Strlist();
  v_起始发票号 票据领用记录.开始号码%Type;
  -------------------------------------------------------------------------------------------------------------------
  --Invoice_Split_Notgroup:不进行分组汇总或首页汇总时调用此过程
  Procedure Invoice_Split_Notgroup
  (
    收费nos_In       Varchar2,
    回收发票_In      Varchar2,
    本次打印发票_Out Out Varchar2,
    本次发票张数_Out Out Number
  ) As
    ----------------------------------------------------------------------------
    --入参:
    --   收费收费NOs_IN:本次需要处理的发票所涉及的单据,多个用逗号分离
    --   回收发票_IN-退费时有效,多个用逗号分离，表示本次需要回收的发票号 
    --出参:
    -- 本次打印发票_Out-本次需要的发票号,多个用逗号分离
    -- 本次发票张数_Out-本次需要的发票数
    -- 本次退费单据_Out-退费回收所涉及的NO号,多个用逗号分离
  
    n_Count Number(18);
    n_分页  Number(18);
  
    Cursor Cr_Bill Is
      Select NO As 元素1, 执行部门id As 元素2, 收据费目 As 元素3, NO As 元素4, NO As 单据, 序号, 0 As 个数
      From 门诊费用记录
      Where Rownum <= 1;
    c_Bill Cr_Bill%RowType;
    --------------------------------------------------------------------------------------------
    --根据相关传入的数据,取对应的数据集
    Type Ty_费用明细 Is Ref Cursor;
    c_费用明细 Ty_费用明细; --游标变量 
  
  Begin
    --按单据分配票据
    If 操作类型_In = 3 Or 操作类型_In = 4 Then
      Open c_费用明细 For
        With c_费用 As
         (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, '-', a.No) As 元素4, a.No As 单据,
                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
          From 门诊费用记录 A,
               (Select NO, 序号 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J Where m.票号 = j.Column_Value) B
          Where MOD(记录性质,10) = 1 And a.No = b.No And Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
          Having Sum(Nvl(a.实收金额, 0)) <> 0)
        Select 元素1, 元素2, 元素3, 元素4, 单据, 序号, Count(*) As 个数
        From c_费用
        Group By 元素1, 元素2, 元素3, 元素4, 单据, 序号
        Order By 元素1, 元素2, 元素3, 元素4, 单据, 序号;
    Else
      Open c_费用明细 For
        With c_费用 As
         (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, '-', a.No) As 元素4, a.No As 单据,
                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
          From 门诊费用记录 A, Table(f_Str2list(收费nos_In)) B
          Where MOD(记录性质,10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
          Having Sum(Nvl(a.实收金额, 0)) <> 0)
        Select 元素1, 元素2, 元素3, 元素4, 单据, 序号, Count(*) As 个数
        From c_费用
        Group By 元素1, 元素2, 元素3, 元素4, 单据, 序号
        Order By 元素1, 元素2, 元素3, 元素4, 单据, 序号;
    End If;
  
    v_元素1          := '+';
    n_元素2          := 0;
    v_元素3          := '+';
    n_元素4          := 0;
    n_元素1_Count    := 0;
    n_元素2_Count    := 0;
    n_元素3_Count    := 0;
    n_元素4_Count    := 0;
    本次发票张数_Out := 0;
    If n_汇总条件 <> 0 Then
      n_关联序号 := 1;
    Else
      n_关联序号 := 0;
    End If;
    n_Count := 0;
    c_No.Delete;
    Loop
      Fetch c_费用明细
        Into c_Bill;
      Exit When c_费用明细%NotFound;
      n_Count := 1;
    
      n_分页 := 0;
      If (v_元素1 <> c_Bill.元素1) Or (n_元素2 <> c_Bill.元素2 And n_元素2_Count >= n_执行科室 And n_执行科室 <> 0) Or
         (v_元素3 <> c_Bill.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Or (n_元素4_Count >= n_收费细目 And n_收费细目 <> 0) Then
      
        If (v_元素1 <> '+' Or n_元素2 <> 0 Or v_元素3 <> '+' Or n_元素4 <> 0) Then
          n_分页 := 1;
        End If;
        n_元素2_Count := 0;
        n_元素3_Count := 0;
        n_元素4_Count := 0;
        n_元素1_Count := 0;
        v_元素1       := '+';
        n_元素2       := 0;
        v_元素3       := '+';
      End If;
    
      If n_分页 = 1 Then
        --分页:计算发票号及相关的
        For I In 1 .. c_No.Count Loop
          c_Invoce.Extend;
          c_Invoce(c_Invoce.Count).票号 := v_发票号;
          c_Invoce(c_Invoce.Count).No := c_No(I).No;
          c_Invoce(c_Invoce.Count).序号 := Case
                                           When Instr(c_No(I).序号, ',') > 0 Then
                                            Substr(c_No(I).序号, 2)
                                           Else
                                            c_No(I).序号
                                         End;
          c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
        End Loop;
      
        本次发票张数_Out := 本次发票张数_Out + 1;
        本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
        v_发票号         := Zl_Incstr(v_发票号);
        c_No.Delete;
      End If;
      If (v_元素1 <> c_Bill.元素1) Then
        n_元素1_Count := n_元素1_Count + 1;
        v_元素1       := c_Bill.元素1;
      End If;
      If (n_元素2 <> c_Bill.元素2) Then
        n_元素2_Count := n_元素2_Count + 1;
        n_元素2       := c_Bill.元素2;
      End If;
      If (v_元素3 <> c_Bill.元素3) Then
        n_元素3_Count := n_元素3_Count + 1;
        v_元素3       := c_Bill.元素3;
      End If;
      If n_收费细目 <> 0 Then
        n_元素4_Count := n_元素4_Count + 1;
      End If;
    
      -------------------------------------------
      --分配单据号及序号
      n_Find := 0;
      For J In 1 .. c_No.Count Loop
        If c_No(J).No = c_Bill.单据 Then
          --单据号相同,将序号合并
          c_No(J).序号 := c_No(J).序号 || ',' || c_Bill.序号;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If n_Find = 0 Then
        c_No.Extend;
        c_No(c_No.Count).No := c_Bill.单据;
        c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_Bill.序号;
      End If;
    End Loop;
  
    --是否有发票数据
    If n_Count >= 1 Then
      --最后一个发票分配
      本次发票张数_Out := 本次发票张数_Out + 1;
      本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
    Else
      本次发票张数_Out := 0;
      本次打印发票_Out := '';
    End If;
    If c_No.Count <> 0 Then
      For I In 1 .. c_No.Count Loop
        c_Invoce.Extend;
        c_Invoce(c_Invoce.Count).票号 := v_发票号;
        c_Invoce(c_Invoce.Count).No := c_No(I).No;
        If Instr(c_No(I).序号, ',') > 0 Then
          c_No(I).序号 := Substr(c_No(I).序号, 2);
        End If;
        c_Invoce(c_Invoce.Count).序号 := c_No(I).序号;
        c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
      End Loop;
      c_No.Delete;
    End If;
    If Instr(Nvl(本次打印发票_Out, '-'), ',') > 0 Then
      本次打印发票_Out := Substr(本次打印发票_Out, 2);
    End If;
  End Invoice_Split_Notgroup;
  --结束:不进行分组汇总或首页汇总时调用此过程
  -------------------------------------------------------------------------------------------------------------------
  --按组汇总
  Procedure Invoice_Split_Group
  (
    收费nos_In       Varchar2,
    回收发票_In      Varchar2,
    本次打印发票_Out Out Varchar2,
    本次发票张数_Out Out Number
  ) As
  Begin
    v_元素1          := '+';
    n_元素2          := 0;
    v_元素3          := '+';
    n_元素4          := 0;
    n_元素1_Count    := 0;
    n_元素2_Count    := 0;
    n_元素3_Count    := 0;
    n_元素4_Count    := 0;
    本次发票张数_Out := 0;
  
    c_No.Delete;
    l_元素2.Delete;
  
    --按单据分配票据
    If 操作类型_In = 3 Or 操作类型_In = 4 Then
      --******************************************************************************************************************************
      --退费和重打按发票号处理(开始)  
      --4.收据费目+收费细目
      If n_分单据打印 = 0 And n_执行科室 = 0 And n_收据费目 <> 0 And n_收费细目 <> 0 Then
        v_元素3 := '+';
        c_Split_费目.Delete;
        For c_分页 In (With c_费用 As
                        (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                               Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                               Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A,
                             (Select NO, 序号
                               From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                               Where m.票号 = j.Column_Value) B
                        Where MOD(记录性质,10) = 1 And a.No = b.No And
                              Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                              Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select /*+ RULE */
                        a.元素3, Count(*) As 个数
                       From c_费用 A
                       Group By 元素3
                       Order By 元素3) Loop
          If (v_元素3 <> c_分页.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Then
            If v_元素3 <> '+' Then
              c_Split_费目.Extend;
              For J In 1 .. l_元素3.Count Loop
                --单据号相同,将序号合并
                c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
              End Loop;
              v_元素3       := '+';
              n_元素3_Count := 0;
              l_元素3.Delete;
            End If;
          End If;
          If (v_元素3 <> c_分页.元素3) Then
            n_元素3_Count := n_元素3_Count + 1;
            v_元素3       := c_分页.元素3;
            l_元素3.Extend;
            l_元素3(l_元素3.Count) := v_元素3;
          End If;
        End Loop;
        If l_元素3.Count <> 0 Then
          c_Split_费目.Extend;
          For J In 1 .. l_元素3.Count Loop
            --单据号相同,将序号合并
            c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
          End Loop;
        End If;
        n_关联序号 := 0;
        For I In 1 .. c_Split_费目.Count Loop
          c_No.Delete;
          n_关联序号    := n_关联序号 + 1;
          n_元素4_Count := 0;
          For c_分页 In (With c_费用 As
                          (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                          From 门诊费用记录 A,
                               (Select NO, 序号
                                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                                 Where m.票号 = j.Column_Value) B
                          Where MOD(记录性质,10) = 1 And a.No = b.No And
                                Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                          Having Sum(Nvl(a.实收金额, 0)) <> 0)
                         Select /*+ RULE */
                          m.元素1, 元素2, 元素3, m.元素4, m.单据, m.序号, Count(*) As 个数
                         From c_费用 M
                         Where Instr(',' || c_Split_费目(I).元素3集 || ',', ',' || m.元素3 || ',') > 0
                         Group By m.元素1, 元素2, m.元素4, 元素3, m.单据, m.序号
                         Order By m.元素1, 元素2, m.元素4, 元素3, m.单据, m.序号) Loop
            If n_元素4_Count >= n_收费细目 And n_收费细目 <> 0 Then
              --分页:计算发票号及相关的
              For J In 1 .. c_No.Count Loop
                c_Invoce.Extend;
                c_Invoce(c_Invoce.Count).票号 := v_发票号;
                c_Invoce(c_Invoce.Count).No := c_No(J).No;
                c_Invoce(c_Invoce.Count).序号 := Case
                                                 When Instr(c_No(J).序号, ',') > 0 Then
                                                  Substr(c_No(J).序号, 2)
                                                 Else
                                                  c_No(J).序号
                                               End;
                c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
              End Loop;
              本次发票张数_Out := 本次发票张数_Out + 1;
              本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
              v_发票号         := Zl_Incstr(v_发票号);
              c_No.Delete;
              n_元素4_Count := 0;
              --分页
            End If;
            n_元素4_Count := n_元素4_Count + 1;
            -------------------------------------------
            --分配单据号及序号
            n_Find := 0;
            For J In 1 .. c_No.Count Loop
              If c_No(J).No = c_分页.单据 Then
                --单据号相同,将序号合并
                c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
                n_Find := 1;
                Exit;
              End If;
            End Loop;
            If n_Find = 0 Then
              c_No.Extend;
              c_No(c_No.Count).No := c_分页.单据;
              c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
            End If;
          End Loop;
          If c_No.Count <> 0 Then
            --分页:计算发票号及相关的
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).票号 := v_发票号;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).序号 := Case
                                               When Instr(c_No(J).序号, ',') > 0 Then
                                                Substr(c_No(J).序号, 2)
                                               Else
                                                c_No(J).序号
                                             End;
              c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
            End Loop;
            本次发票张数_Out := 本次发票张数_Out + 1;
            本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
            v_发票号         := Zl_Incstr(v_发票号);
            c_No.Delete;
          End If;
        End Loop;
      End If;
    
      If (n_分单据打印 = 1 Or n_执行科室 > 0) And (n_收据费目 <> 0 Or n_收费细目 <> 0) Then
        n_元素2_Count := 0;
        v_元素1       := '+';
        n_元素2       := 0;
        c_Split_No.Delete;
        For c_分页 In (With c_费用 As
                        (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                               Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                               Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A,
                             (Select NO, 序号
                               From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                               Where m.票号 = j.Column_Value) B
                        Where MOD(记录性质,10) = 1 And a.No = b.No And
                              Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                             
                              Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select /*+ RULE */
                        a.元素1, a.元素2, b.编码, Count(*) As 个数
                       From c_费用 A, 部门表 B
                       Where a.元素2 = b.Id(+)
                       Group By 元素1, b.编码, 元素2
                       Order By 元素1, b.编码, 元素2) Loop
          If (v_元素1 <> c_分页.元素1) Or (n_元素2 <> c_分页.元素2 And n_元素2_Count >= n_执行科室 And n_执行科室 <> 0) Then
            c_Split_No.Extend;
            n_元素2_Count := 0;
            v_元素1       := '+';
            n_元素2       := 0;
          End If;
          If (v_元素1 <> c_分页.元素1) Then
            v_元素1 := c_分页.元素1;
            c_Split_No(c_Split_No.Count).元素1 := v_元素1;
          End If;
          If (n_元素2 <> c_分页.元素2) Then
            n_元素2_Count := n_元素2_Count + 1;
            n_元素2 := c_分页.元素2;
            c_Split_No(c_Split_No.Count).元素2集 := c_Split_No(c_Split_No.Count).元素2集 || ',' || n_元素2;
          End If;
        End Loop;
      End If;
    
      --6.(no Or 执行科室)+收费细目
      If (n_分单据打印 = 0 Or n_执行科室 > 0) And n_收据费目 = 0 And n_收费细目 <> 0 Then
      
        For I In 1 .. c_Split_No.Count Loop
          v_元素3 := '+';
          --只有首页汇总的,才有关联序号
          n_关联序号    := n_关联序号 + 1;
          n_元素4_Count := 0;
          For c_分页 In (With c_费用 As
                          (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                          From 门诊费用记录 A,
                               (Select NO, 序号
                                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                                 Where m.票号 = j.Column_Value) B
                          Where MOD(记录性质,10) = 1 And a.No = b.No And
                                Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                               
                                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                          Having Sum(Nvl(a.实收金额, 0)) <> 0)
                         Select /*+ RULE */
                          元素1, 元素2, 元素3, a.元素4, a.单据, a.序号, Count(*) As 个数
                         From c_费用 A
                         Where a.元素1 = c_Split_No(I).元素1 And
                               Instr(',' || c_Split_No(I).元素2集 || ',', ',' || a.元素2 || ',') > 0
                         Group By 元素1, 元素2, 元素4, 元素3, 单据, a.序号
                         Order By 元素1, 元素2, 元素4, 单据, 序号) Loop
            If n_元素4_Count >= n_收费细目 And n_收费细目 <> 0 Then
              --分配单据
              If c_No.Count <> 0 Then
                --分页:计算发票号及相关的
                For J In 1 .. c_No.Count Loop
                  c_Invoce.Extend;
                  c_Invoce(c_Invoce.Count).票号 := v_发票号;
                  c_Invoce(c_Invoce.Count).No := c_No(J).No;
                  c_Invoce(c_Invoce.Count).序号 := Case
                                                   When Instr(c_No(J).序号, ',') > 0 Then
                                                    Substr(c_No(J).序号, 2)
                                                   Else
                                                    c_No(J).序号
                                                 End;
                  c_Invoce(c_Invoce.Count).关联序号 := n_元素4_Count;
                End Loop;
                本次发票张数_Out := 本次发票张数_Out + 1;
                本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
                v_发票号         := Zl_Incstr(v_发票号);
                c_No.Delete;
              End If;
              n_元素4_Count := 0;
            End If;
            n_元素4_Count := n_元素4_Count + 1;
          
            -------------------------------------------
            --分配单据号及序号
            n_Find := 0;
            For J In 1 .. c_No.Count Loop
              If c_No(J).No = c_分页.单据 Then
                --单据号相同,将序号合并
                c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
                n_Find := 1;
                Exit;
              End If;
            End Loop;
            If n_Find = 0 Then
              c_No.Extend;
              c_No(c_No.Count).No := c_分页.单据;
              c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
            End If;
          End Loop;
          --分配单据
          If c_No.Count <> 0 Then
            --分页:计算发票号及相关的
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).票号 := v_发票号;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).序号 := Case
                                               When Instr(c_No(J).序号, ',') > 0 Then
                                                Substr(c_No(J).序号, 2)
                                               Else
                                                c_No(J).序号
                                             End;
              c_Invoce(c_Invoce.Count).关联序号 := n_元素4_Count;
            End Loop;
            本次发票张数_Out := 本次发票张数_Out + 1;
            本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
            v_发票号         := Zl_Incstr(v_发票号);
            c_No.Delete;
          End If;
        End Loop;
      End If;
    
      --7.(no Or 执行科室)+收据费目+收费细目
      n_关联序号 := 0;
      If (n_分单据打印 = 0 Or n_执行科室 > 0) And n_收据费目 <> 0 And n_收费细目 <> 0 Then
        c_Split_费目.Delete;
        For I In 1 .. c_Split_No.Count Loop
        
          n_关联序号    := n_关联序号 + 1;
          v_元素3       := '+';
          n_元素3_Count := 0;
          l_元素3.Delete;
          For c_分页 In (With c_费用 As
                          (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                          From 门诊费用记录 A,
                               (Select NO, 序号
                                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                                 Where m.票号 = j.Column_Value) B
                          Where MOD(记录性质,10) = 1 And a.No = b.No And
                                Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                               
                                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                          Having Sum(Nvl(a.实收金额, 0)) <> 0)
                         Select /*+ RULE */
                          a.元素3, Count(*) As 个数
                         From c_费用 A
                         Where a.元素1 = c_Split_No(I).元素1 And
                               Instr(',' || c_Split_No(I).元素2集 || ',', ',' || a.元素2 || ',') > 0
                         Group By 元素3
                         Order By 元素3) Loop
          
            If (v_元素3 <> c_分页.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Then
              If v_元素3 <> '+' Then
                c_Split_费目.Extend;
                c_Split_费目(c_Split_费目.Count).元素1 := c_Split_No(I).元素1;
                c_Split_费目(c_Split_费目.Count).元素2集 := c_Split_No(I).元素2集;
                c_Split_费目(c_Split_费目.Count).关联序号 := n_关联序号;
                For J In 1 .. l_元素3.Count Loop
                  --单据号相同,将序号合并
                  c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
                End Loop;
              End If;
              v_元素3       := '+';
              n_元素3_Count := 0;
              l_元素3.Delete;
            End If;
            If (v_元素3 <> c_分页.元素3) Then
              n_元素3_Count := n_元素3_Count + 1;
              v_元素3       := c_分页.元素3;
              l_元素3.Extend;
              l_元素3(l_元素3.Count) := v_元素3;
            End If;
          End Loop;
        
          If l_元素3.Count <> 0 Then
            c_Split_费目.Extend;
            c_Split_费目(c_Split_费目.Count).元素1 := c_Split_No(I).元素1;
            c_Split_费目(c_Split_费目.Count).元素2集 := c_Split_No(I).元素2集;
            c_Split_费目(c_Split_费目.Count).关联序号 := n_关联序号;
            For J In 1 .. l_元素3.Count Loop
              --单据号相同,将序号合并
              c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
            End Loop;
          End If;
        End Loop;
      
        For I In 1 .. c_Split_费目.Count Loop
          c_No.Delete;
          n_元素4_Count := 0;
          For c_分页 In (With c_费用 As
                          (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                          From 门诊费用记录 A,
                               (Select NO, 序号
                                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                                 Where m.票号 = j.Column_Value) B
                          Where MOD(记录性质,10) = 1 And a.No = b.No And
                                Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                               
                                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                          Having Sum(Nvl(a.实收金额, 0)) <> 0)
                         Select /*+ RULE */
                          元素1, 元素2, 元素3, a.元素4, a.单据, a.序号, Count(*) As 个数
                         From c_费用 A
                         Where a.元素1 = c_Split_费目(I).元素1 And
                               Instr(',' || c_Split_费目(I).元素2集 || ',', ',' || a.元素2 || ',') > 0 And
                               Instr(',' || c_Split_费目(I).元素3集 || ',', ',' || a.元素3 || ',') > 0
                         Group By 元素1, 元素2, 元素4, 元素3, a.单据, a.序号
                         Order By 元素1, 元素2, 元素4, 元素3, 单据, 序号) Loop
            If (n_元素4_Count >= n_收费细目 And n_收费细目 <> 0) Then
              --分配单据
              If c_No.Count <> 0 Then
                --分页:计算发票号及相关的
                For J In 1 .. c_No.Count Loop
                  c_Invoce.Extend;
                  c_Invoce(c_Invoce.Count).票号 := v_发票号;
                  c_Invoce(c_Invoce.Count).No := c_No(J).No;
                  c_Invoce(c_Invoce.Count).序号 := Case
                                                   When Instr(c_No(J).序号, ',') > 0 Then
                                                    Substr(c_No(J).序号, 2)
                                                   Else
                                                    c_No(J).序号
                                                 End;
                  c_Invoce(c_Invoce.Count).关联序号 := c_Split_费目(I).关联序号;
                End Loop;
                本次发票张数_Out := 本次发票张数_Out + 1;
                本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
                v_发票号         := Zl_Incstr(v_发票号);
                c_No.Delete;
              End If;
              n_元素4_Count := 0;
            End If;
            n_元素4_Count := n_元素4_Count + 1;
            -------------------------------------------
            --分配单据号及序号
            n_Find := 0;
            For J In 1 .. c_No.Count Loop
              If c_No(J).No = c_分页.单据 Then
                --单据号相同,将序号合并
                c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
                n_Find := 1;
                Exit;
              End If;
            End Loop;
            If n_Find = 0 Then
              c_No.Extend;
              c_No(c_No.Count).No := c_分页.单据;
              c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
            End If;
          End Loop;
        
          --分配单据
          If c_No.Count <> 0 Then
            --分页:计算发票号及相关的
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).票号 := v_发票号;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).序号 := Case
                                               When Instr(c_No(J).序号, ',') > 0 Then
                                                Substr(c_No(J).序号, 2)
                                               Else
                                                c_No(J).序号
                                             End;
              c_Invoce(c_Invoce.Count).关联序号 := c_Split_费目(I).关联序号;
            End Loop;
            本次发票张数_Out := 本次发票张数_Out + 1;
            本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
            v_发票号         := Zl_Incstr(v_发票号);
            c_No.Delete;
          End If;
        End Loop;
      End If;
    
      --退费和重打按发票号处理(结束)
      --******************************************************************************************************************************
      If Instr(Nvl(本次打印发票_Out, '-'), ',') > 0 Then
        本次打印发票_Out := Substr(本次打印发票_Out, 2);
      End If;
      Return;
    
    End If;
  
    --******************************************************************************************************************************
    --以下是按正常分配单据(开始)
    --4.收据费目+收费细目
    If n_分单据打印 = 0 And n_执行科室 = 0 And n_收据费目 <> 0 And n_收费细目 <> 0 Then
      v_元素3 := '+';
      c_Split_费目.Delete;
    
      For c_分页 In (With c_费用 As
                      (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                             Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                             Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                      From 门诊费用记录 A, Table(f_Str2list(收费nos_In)) B
                      Where MOD(记录性质,10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                      Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                      Having Sum(Nvl(a.实收金额, 0)) <> 0)
                     Select /*+ RULE */
                      a.元素3, Count(*) As 个数
                     From c_费用 A
                     Group By 元素3
                     Order By 元素3) Loop
        If (v_元素3 <> c_分页.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Then
          If v_元素3 <> '+' Then
            c_Split_费目.Extend;
            For J In 1 .. l_元素3.Count Loop
              --单据号相同,将序号合并
              c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
            End Loop;
            v_元素3       := '+';
            n_元素3_Count := 0;
            l_元素3.Delete;
          End If;
        End If;
        If (v_元素3 <> c_分页.元素3) Then
          n_元素3_Count := n_元素3_Count + 1;
          v_元素3       := c_分页.元素3;
          l_元素3.Extend;
          l_元素3(l_元素3.Count) := v_元素3;
        End If;
      End Loop;
      If l_元素3.Count <> 0 Then
        c_Split_费目.Extend;
        For J In 1 .. l_元素3.Count Loop
          --单据号相同,将序号合并
          c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
        End Loop;
      End If;
      n_关联序号 := 0;
      For I In 1 .. c_Split_费目.Count Loop
        c_No.Delete;
        n_关联序号    := n_关联序号 + 1;
        n_元素4_Count := 0;
        For c_分页 In (With c_费用 As
                        (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                               Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                               Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A, Table(f_Str2list(收费nos_In)) B
                        Where MOD(记录性质,10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select /*+ RULE */
                        m.元素1, 元素2, 元素3, m.元素4, m.单据, m.序号, Count(*) As 个数
                       From c_费用 M
                       Where Instr(',' || c_Split_费目(I).元素3集 || ',', ',' || m.元素3 || ',') > 0
                       Group By m.元素1, 元素2, m.元素4, 元素3, m.单据, m.序号
                       Order By m.元素1, 元素2, m.元素4, 元素3, m.单据, m.序号) Loop
          If n_元素4_Count >= n_收费细目 And n_收费细目 <> 0 Then
            --分页:计算发票号及相关的
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).票号 := v_发票号;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).序号 := Case
                                               When Instr(c_No(J).序号, ',') > 0 Then
                                                Substr(c_No(J).序号, 2)
                                               Else
                                                c_No(J).序号
                                             End;
              c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
            End Loop;
            本次发票张数_Out := 本次发票张数_Out + 1;
            本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
            v_发票号         := Zl_Incstr(v_发票号);
            c_No.Delete;
            n_元素4_Count := 0;
            --分页
          End If;
          n_元素4_Count := n_元素4_Count + 1;
          -------------------------------------------
          --分配单据号及序号
          n_Find := 0;
          For J In 1 .. c_No.Count Loop
            If c_No(J).No = c_分页.单据 Then
              --单据号相同,将序号合并
              c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
              n_Find := 1;
              Exit;
            End If;
          End Loop;
          If n_Find = 0 Then
            c_No.Extend;
            c_No(c_No.Count).No := c_分页.单据;
            c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
          End If;
        End Loop;
        If c_No.Count <> 0 Then
          --分页:计算发票号及相关的
          For J In 1 .. c_No.Count Loop
            c_Invoce.Extend;
            c_Invoce(c_Invoce.Count).票号 := v_发票号;
            c_Invoce(c_Invoce.Count).No := c_No(J).No;
            c_Invoce(c_Invoce.Count).序号 := Case
                                             When Instr(c_No(J).序号, ',') > 0 Then
                                              Substr(c_No(J).序号, 2)
                                             Else
                                              c_No(J).序号
                                           End;
            c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
          End Loop;
          本次发票张数_Out := 本次发票张数_Out + 1;
          本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
          v_发票号         := Zl_Incstr(v_发票号);
          c_No.Delete;
        End If;
      End Loop;
    End If;
  
    If (n_分单据打印 = 1 Or n_执行科室 > 0) And (n_收据费目 <> 0 Or n_收费细目 <> 0) Then
      n_元素2_Count := 0;
      v_元素1       := '+';
      n_元素2       := 0;
      c_Split_No.Delete;
      For c_分页 In (With c_费用 As
                      (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                             Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                             Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                      From 门诊费用记录 A, Table(f_Str2list(收费nos_In)) B
                      Where MOD(记录性质,10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                      Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                      Having Sum(Nvl(a.实收金额, 0)) <> 0)
                     Select /*+ RULE */
                      a.元素1, a.元素2, b.编码, Count(*) As 个数
                     From c_费用 A, 部门表 B
                     Where a.元素2 = b.Id(+)
                     Group By 元素1, b.编码, 元素2
                     Order By 元素1, b.编码, 元素2) Loop
        If (v_元素1 <> c_分页.元素1) Or (n_元素2 <> c_分页.元素2 And n_元素2_Count >= n_执行科室 And n_执行科室 <> 0) Then
          c_Split_No.Extend;
          n_元素2_Count := 0;
          v_元素1       := '+';
          n_元素2       := 0;
        End If;
        If (v_元素1 <> c_分页.元素1) Then
          v_元素1 := c_分页.元素1;
          c_Split_No(c_Split_No.Count).元素1 := v_元素1;
        End If;
        If (n_元素2 <> c_分页.元素2) Then
          n_元素2_Count := n_元素2_Count + 1;
          n_元素2 := c_分页.元素2;
          c_Split_No(c_Split_No.Count).元素2集 := c_Split_No(c_Split_No.Count).元素2集 || ',' || n_元素2;
        End If;
      End Loop;
    End If;
  
    --3.(no Or 执行科室)+收费细目
    If (n_分单据打印 = 0 Or n_执行科室 > 0) And n_收据费目 = 0 And n_收费细目 <> 0 Then
    
      For I In 1 .. c_Split_No.Count Loop
        v_元素3 := '+';
        --只有首页汇总的,才有关联序号
        n_关联序号    := Nvl(n_关联序号, 0) + 1;
        n_元素4_Count := 0;
        For c_分页 In (With c_费用 As
                        (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                               Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                               Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A, Table(f_Str2list(收费nos_In)) B
                        Where MOD(记录性质,10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select /*+ RULE */
                        元素1, 元素2, 元素3, a.元素4, a.单据, a.序号, Count(*) As 个数
                       From c_费用 A
                       Where a.元素1 = c_Split_No(I).元素1 And
                             Instr(',' || c_Split_No(I).元素2集 || ',', ',' || a.元素2 || ',') > 0
                       Group By 元素1, 元素2, 元素4, 元素3, 单据, a.序号
                       Order By 元素1, 元素2, 元素4, 单据, 序号) Loop
          If n_元素4_Count >= n_收费细目 And n_收费细目 <> 0 Then
            --分配单据
            If c_No.Count <> 0 Then
              --分页:计算发票号及相关的
              For J In 1 .. c_No.Count Loop
                c_Invoce.Extend;
                c_Invoce(c_Invoce.Count).票号 := v_发票号;
                c_Invoce(c_Invoce.Count).No := c_No(J).No;
                c_Invoce(c_Invoce.Count).序号 := Case
                                                 When Instr(c_No(J).序号, ',') > 0 Then
                                                  Substr(c_No(J).序号, 2)
                                                 Else
                                                  c_No(J).序号
                                               End;
                c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
              End Loop;
              本次发票张数_Out := 本次发票张数_Out + 1;
              本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
              v_发票号         := Zl_Incstr(v_发票号);
              c_No.Delete;
            End If;
            n_元素4_Count := 0;
          End If;
          n_元素4_Count := n_元素4_Count + 1;
        
          -------------------------------------------
          --分配单据号及序号
          n_Find := 0;
          For J In 1 .. c_No.Count Loop
            If c_No(J).No = c_分页.单据 Then
              --单据号相同,将序号合并
              c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
              n_Find := 1;
              Exit;
            End If;
          End Loop;
          If n_Find = 0 Then
            c_No.Extend;
            c_No(c_No.Count).No := c_分页.单据;
            c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
          End If;
        End Loop;
        --分配单据
        If c_No.Count <> 0 Then
          --分页:计算发票号及相关的
          For J In 1 .. c_No.Count Loop
            c_Invoce.Extend;
            c_Invoce(c_Invoce.Count).票号 := v_发票号;
            c_Invoce(c_Invoce.Count).No := c_No(J).No;
            c_Invoce(c_Invoce.Count).序号 := Case
                                             When Instr(c_No(J).序号, ',') > 0 Then
                                              Substr(c_No(J).序号, 2)
                                             Else
                                              c_No(J).序号
                                           End;
            c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
          End Loop;
          本次发票张数_Out := 本次发票张数_Out + 1;
          本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
          v_发票号         := Zl_Incstr(v_发票号);
          c_No.Delete;
        End If;
      End Loop;
    End If;
  
    --7.(no Or 执行科室)+收据费目+收费细目
    n_关联序号 := 0;
    If (n_分单据打印 = 0 Or n_执行科室 > 0) And n_收据费目 <> 0 And n_收费细目 <> 0 Then
      c_Split_费目.Delete;
    
      For I In 1 .. c_Split_No.Count Loop
        
        n_关联序号    := n_关联序号 + 1;
        v_元素3       := '+';
        n_元素3_Count := 0;
        l_元素3.Delete;
        For c_分页 In (With c_费用 As
                        (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                               Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                               Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A, Table(f_Str2list(收费nos_In)) B
                        Where MOD(记录性质,10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select /*+ RULE */
                        a.元素3, Count(*) As 个数
                       From c_费用 A
                       Where a.元素1 = c_Split_No(I).元素1 And
                             Instr(',' || c_Split_No(I).元素2集 || ',', ',' || a.元素2 || ',') > 0
                       Group By 元素3
                       Order By 元素3) Loop
          If (v_元素3 <> c_分页.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Then
            If v_元素3 <> '+' Then
              c_Split_费目.Extend;
              c_Split_费目(c_Split_费目.Count).元素1 := c_Split_No(I).元素1;
              c_Split_费目(c_Split_费目.Count).元素2集 := c_Split_No(I).元素2集;
              c_Split_费目(c_Split_费目.Count).关联序号 := n_关联序号;
              For J In 1 .. l_元素3.Count Loop
                --单据号相同,将序号合并
                c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
              End Loop;
            End If;
            v_元素3       := '+';
            n_元素3_Count := 0;
            l_元素3.Delete;
          End If;
          If (v_元素3 <> c_分页.元素3) Then
            n_元素3_Count := n_元素3_Count + 1;
            v_元素3       := c_分页.元素3;
            l_元素3.Extend;
            l_元素3(l_元素3.Count) := v_元素3;
          End If;
        End Loop;
      
        If l_元素3.Count <> 0 Then
          c_Split_费目.Extend;
          c_Split_费目(c_Split_费目.Count).元素1 := c_Split_No(I).元素1;
          c_Split_费目(c_Split_费目.Count).元素2集 := c_Split_No(I).元素2集;
          c_Split_费目(c_Split_费目.Count).关联序号 := n_关联序号;
          For J In 1 .. l_元素3.Count Loop
            --单据号相同,将序号合并
            c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
          End Loop;
        End If;
      End Loop;
    
      For I In 1 .. c_Split_费目.Count Loop
        c_No.Delete;
        n_元素4_Count := 0;
        --收费细目,按条数计数,还是要按执行科室+收据费目
        For c_分页 In (With c_费用 As
                        (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                               Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                               Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A, Table(f_Str2list(收费nos_In)) B
                        Where MOD(记录性质,10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select /*+ RULE */
                        元素1, 元素2, 元素3, a.元素4, a.单据, a.序号, Count(*) As 个数
                       From c_费用 A
                       Where a.元素1 = c_Split_费目(I).元素1 And
                             Instr(',' || c_Split_费目(I).元素2集 || ',', ',' || a.元素2 || ',') > 0 And
                             Instr(',' || c_Split_费目(I).元素3集 || ',', ',' || a.元素3 || ',') > 0
                       Group By 元素1, 元素2, 元素4, 元素3, a.单据, a.序号
                       Order By 元素1, 元素2, 元素4, 元素3, 单据, 序号) Loop
          If (n_元素4_Count >= n_收费细目 And n_收费细目 <> 0) Then
            --分配单据
            If c_No.Count <> 0 Then
              --分页:计算发票号及相关的
              For J In 1 .. c_No.Count Loop
                c_Invoce.Extend;
                c_Invoce(c_Invoce.Count).票号 := v_发票号;
                c_Invoce(c_Invoce.Count).No := c_No(J).No;
                c_Invoce(c_Invoce.Count).序号 := Case
                                                 When Instr(c_No(J).序号, ',') > 0 Then
                                                  Substr(c_No(J).序号, 2)
                                                 Else
                                                  c_No(J).序号
                                               End;
                c_Invoce(c_Invoce.Count).关联序号 := c_Split_费目(I).关联序号;
              End Loop;
              本次发票张数_Out := 本次发票张数_Out + 1;
              本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
              v_发票号         := Zl_Incstr(v_发票号);
              c_No.Delete;
            End If;
            n_元素4_Count := 0;
          End If;
          n_元素4_Count := n_元素4_Count + 1;
          -------------------------------------------
          --分配单据号及序号
          n_Find := 0;
          For J In 1 .. c_No.Count Loop
            If c_No(J).No = c_分页.单据 Then
              --单据号相同,将序号合并
              c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
              n_Find := 1;
              Exit;
            End If;
          End Loop;
          If n_Find = 0 Then
            c_No.Extend;
            c_No(c_No.Count).No := c_分页.单据;
            c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
          End If;
        End Loop;
        --分配单据
        If c_No.Count <> 0 Then
          --分页:计算发票号及相关的
          For J In 1 .. c_No.Count Loop
            c_Invoce.Extend;
            c_Invoce(c_Invoce.Count).票号 := v_发票号;
            c_Invoce(c_Invoce.Count).No := c_No(J).No;
            c_Invoce(c_Invoce.Count).序号 := Case
                                             When Instr(c_No(J).序号, ',') > 0 Then
                                              Substr(c_No(J).序号, 2)
                                             Else
                                              c_No(J).序号
                                           End;
            c_Invoce(c_Invoce.Count).关联序号 := c_Split_费目(I).关联序号;
          End Loop;
          本次发票张数_Out := 本次发票张数_Out + 1;
          本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
          v_发票号         := Zl_Incstr(v_发票号);
          c_No.Delete;
        End If;
      End Loop;
    End If;
    --正常分配单据结束
    --******************************************************************************************************************************
    If Instr(Nvl(本次打印发票_Out, '-'), ',') > 0 Then
      本次打印发票_Out := Substr(本次打印发票_Out, 2);
    End If;
  End Invoice_Split_Group;
  -------------------------------------------------------------------------------------------------------------------
Begin

  --启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
  v_Para := Nvl(zl_GetSysParameter('票据分配规则', 1121), '0||0;0;0,0;0');
  If Instr(v_Para, '||') = 0 Then
    v_Para := v_Para || '||';
  End If;
  v_Temp := Substr(v_Para, 1, Instr(v_Para, '||') - 1);
  If v_Temp Is Null Then
    --无设置值,代表无启用,直接返回
    Return;
  End If;

  --0-根据实际打印分配票号;1-根据预定规则分配票号;2-.根据自定义规则分配票号
  n_启用模式 := Zl_To_Number(v_Temp);
  If Nvl(n_启用模式, 0) = 0 Then
    --0-根据实际打印分配票号:按原来的处理方式分配票据,直接退出
    Return;
  End If;
  v_Temp       := Nvl(zl_GetSysParameter('误差项不使用票据', 1121), '0');
  n_误差项     := Zl_To_Number(v_Temp);
  v_起始发票号 := 起始发票号_In;

  If v_起始发票号 Is Null Then
    --模拟计算时,可以不传入起始发票号
    If Nvl(领用id_In, 0) <> 0 Then
      Open c_Fact;
      Fetch c_Fact
        Into r_Factrow;
    
      If c_Fact%RowCount <> 0 Then
        If Nvl(r_Factrow.当前号码, '-') <> '-' Then
          v_起始发票号 := Zl_Incstr(r_Factrow.当前号码);
        Else
          v_起始发票号 := r_Factrow.开始号码;
        End If;
      End If;
    End If;
    If v_起始发票号 Is Null Then
      v_起始发票号 := 'J0000001';
    End If;
  End If;
  
  v_发票号   := v_起始发票号;
  v_发票信息 := Null;

  --按单据分配票据
  If 操作类型_In = 3 Or 操作类型_In = 4 Then
    --1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
    If 发票号_In Is Null Then
      v_Err_Msg := '未传入指定的回收票据,不允许' || Case
                     When 操作类型_In = 1 Then
                      '重打票据。'
                     Else
                      '补打票据。'
                   End;
      Raise Err_Item;
    End If;
  
    --读取当前票据所涉及的所有票据
    v_Nos := Null;
    For c_票据 In (Select Distinct c.No As 单据号
                 From 票据打印明细 A, 票据使用明细 B, 票据打印内容 C, Table(f_Str2list(发票号_In)) J
                 Where a.使用id = b.Id And b.打印id = c.Id And a.票号 = j.Column_Value
                 Order By 单据号) Loop
      v_Nos := Nvl(v_Nos, '') || ',' || c_票据.单据号;
    End Loop;
  
    If v_Nos Is Null Then
      v_Err_Msg := '未找到指定发票(' || 发票号_In || '所对应的收费单据!';
      Raise Err_Item;
    End If;
    v_Nos := Substr(v_Nos, 2);
  Else
    v_Nos := Nos_In;
  End If;

  If n_启用模式 = 2 Then
    --根据自定义规则分配票号,调用:Zl_Custom_Invoice_Autoallot过程
    Zl_Custom_Invoice_Autoallot(操作类型_In, 模拟计算_In, 票种_In, 领用id_In, 病人id_In, v_Nos, 起始发票号_In, 使用人_In, 使用时间_In, 发票号_In,
                                发票张数_In);
    Return;
  End If;

  --参数获取:
  --1.根据预定规则分配票号
  --   NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
  v_Para := Substr(v_Para, Instr(v_Para, '||') + 2);
  If Instr(v_Para, ';') > 0 Then
    --NO:票据是否按单据进行分别打印,1表示按单据打印;0-不按单据打印
    v_Temp       := Substr(v_Para, 1, Instr(v_Para, ';') - 1);
    n_分单据打印 := Zl_To_Number(v_Temp);
    v_Para       := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --执行科室
    v_Temp     := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
    n_执行科室 := Zl_To_Number(v_Temp);
    v_Para     := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --收据费目
    v_Temp     := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
    n_收据费目 := Zl_To_Number(v_Temp);
    v_Para     := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --收据费目
    v_Temp     := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
    n_收费细目 := Zl_To_Number(v_Temp);
    v_Para     := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --执行科室
    v_Temp := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
  Else
    v_Temp := Nvl(v_Para, '0');
  End If;
  n_汇总条件 := Zl_To_Number(v_Temp);

  v_回收票据号 := 发票号_In;
  发票张数_In  := 0;
  --一、首页汇总或不汇总
  If n_汇总条件 <> 2 Then
    Invoice_Split_Notgroup(Nos_In, 发票号_In, v_发票信息, 发票张数_In);
  Else
    --二、分组汇总
    Invoice_Split_Group(Nos_In, 发票号_In, v_发票信息, 发票张数_In);
  End If;
  发票号_In := v_发票信息;
  If 模拟计算_In = 1 Then
    --模拟计算,只返回票据张数和使用的票据号,直接退出
    Return;
  End If;

  v_开始发票号 := Null;
  v_当前发票号 := Null;
  --1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
  If 操作类型_In = 3 Or 操作类型_In = 4 Then
    --收回票据
    Select 使用id Bulk Collect
    Into l_使用id
    From (Select Distinct b.使用id
           From 票据使用明细 A, 票据打印明细 B, Table(f_Str2list(v_回收票据号)) J
           Where a.Id = b.使用id And b.票号 = j.Column_Value And Nvl(b.票种, 0) = 1);
  
    --插入回收记录
    Forall I In 1 .. l_使用id.Count
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
        Select 票据使用明细_Id.Nextval, 票种, 号码, 2, Decode(操作类型_In, 3, 4, 2), 领用id, 打印id, 使用人_In, 使用时间_In
        From 票据使用明细
        Where ID = l_使用id(I);
    Forall I In 1 .. l_使用id.Count
      Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I);
  End If;

  If c_Invoce.Count = 0 Then
    --无数据,直接返回
    Return;
  End If;

  If 起始发票号_In Is Null Then
    v_Err_Msg := '未传入起始发票号,不能进行票据分配处理';
    Raise Err_Item;
  End If;

  If Nvl(领用id_In, 0) <> 0 Then
    Open c_Fact;
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Err_Msg := '无效的票据领用批次，无法完成收费票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    Elsif Nvl(r_Factrow.剩余数量, 0) < 发票张数_In Then
      v_Err_Msg := '当前批次的剩余数量不足' || 发票张数_In || '张，无法完成收费票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    End If;
  End If;
  --实际处理票据信息
  If Nvl(n_分单据打印, 0) <> 1 Then
    --不分单据打印时,表示一次打印,打印ID填成一致
    Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  End If;

  发票张数_In := 0;
  v_打印内容  := '';
  For c_Invoce_No In (Select Column_Value As 发票号 From Table(f_Str2list(v_发票信息)) Order By 发票号) Loop
    --检查票据范围是否正确
    If Nvl(领用id_In, 0) <> 0 Then
      If Not (Upper(c_Invoce_No.发票号) >= Upper(r_Factrow.开始号码) And Upper(c_Invoce_No.发票号) <= Upper(r_Factrow.终止号码) And
          Length(c_Invoce_No.发票号) = Length(r_Factrow.终止号码)) Then
        v_Err_Msg := '该单据需要打印多张票据,但票据号"' || c_Invoce_No.发票号 || '"超出票据领用的号码范围！';
        Close c_Fact;
        Raise Err_Item;
      End If;
    End If;
  
    --处理票据打印明细
    r_单据号.Delete;
    r_单据序号.Delete;
    l_关联序号.Delete;
  
    Select 票据使用明细_Id.Nextval Into n_使用id From Dual;
  
    n_关联序号 := 0;
  
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).票号 = c_Invoce_No.发票号 Then
        n_关联序号 := c_Invoce(I).关联序号;
        Exit;
      End If;
    End Loop;
    --处理关联票据,以便回收票据
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).关联序号 = n_关联序号 And Nvl(c_Invoce(I).修改标志, 0) = 0 Then
        If n_关联序号 <> 0 Then
          c_Invoce(I).关联序号 := n_使用id;
        End If;
        c_Invoce(I).修改标志 := 1;
      End If;
    End Loop;
  
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).票号 = c_Invoce_No.发票号 Then
        r_单据号.Extend;
        r_单据号(r_单据号.Count) := c_Invoce(I).No;
        r_单据序号.Extend;
        r_单据序号(r_单据序号.Count) := c_Invoce(I).序号;
        l_关联序号.Extend;
        If Nvl(c_Invoce(I).关联序号, 0) <> 0 Then
          --检查是否存在其他的票据
          n_Find := 0;
          For J In 1 .. c_Invoce.Count Loop
            If c_Invoce(I).关联序号 = c_Invoce(J).关联序号 And c_Invoce(I).票号 <> c_Invoce(J).票号 Then
              n_Find := 1;
              Exit;
            End If;
          End Loop;
        
          If n_Find = 0 Then
            l_关联序号(l_关联序号.Count) := Null;
            c_Invoce(I).关联序号 := 0;
          Else
            l_关联序号(l_关联序号.Count) := c_Invoce(I).关联序号;
          End If;
        Else
          l_关联序号(l_关联序号.Count) := Null;
        End If;
      End If;
    End Loop;
  
    --1.处理门打印内容
    If n_分单据打印 = 1 Then
      --分单据打印,需按单据进行处理
      --票据打印内容
      n_Find := 0;
      v_Temp := '';
      For I In 1 .. r_单据号.Count Loop
        v_Temp := v_Temp || ',' || r_单据号(I);
        If Instr(Nvl(v_打印内容, '-') || ',', ',' || r_单据号(I) || ',') > 0 Then
          --已经找到
          n_Find := 1;
        End If;
      End Loop;
      v_打印内容 := v_打印内容 || Nvl(v_Temp, '+');
    
      If Nvl(n_Find, 0) = 0 Then
        Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
        Forall I In 1 .. r_单据号.Count
          Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 1, r_单据号(I));
        --以便更新门诊费用记录中的实际票号
        v_开始发票号 := c_Invoce_No.发票号;
        Forall I In 1 .. r_单据号.Count
          Update 门诊费用记录 Set 实际票号 = v_开始发票号 Where MOD(记录性质,10) = 1 And NO = r_单据号(I);
      End If;
    Else
    
      If v_开始发票号 Is Null Then
        --以便更新门诊费用记录中的实际票号
        v_开始发票号 := c_Invoce_No.发票号;
      
        --票据打印内容
        Insert Into 票据打印内容
          (ID, 数据性质, NO)
          Select n_打印id, 1, Column_Value From Table(f_Str2list(v_Nos));
      
        Update 门诊费用记录
        Set 实际票号 = v_开始发票号
        Where MOD(记录性质,10) = 1 And NO In (Select Column_Value From Table(f_Str2list(v_Nos)));
      End If;
    End If;
  
    --2.处理票据打印明细
  
    发票张数_In := 发票张数_In + 1;
    --处理票据使用明细
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
    Values
      (n_使用id, 1, c_Invoce_No.发票号, 1, Decode(操作类型_In, 3, 3, 1), Decode(Nvl(领用id_In, 0), 0, Null, 领用id_In), n_打印id, 使用人_In, 使用时间_In);
  
    Forall I In 1 .. r_单据号.Count
      Insert Into 票据打印明细
        (使用id, 票种, 是否回收, NO, 票号, 序号, 关联票号序号)
      Values
        (n_使用id, 1, 0, r_单据号(I), c_Invoce_No.发票号, r_单据序号(I), l_关联序号(I));
  
    v_当前发票号 := c_Invoce_No.发票号;
  End Loop;

  If Nvl(领用id_In, 0) <> 0 Then
    Close c_Fact;
  
    Update 票据领用记录
    Set 使用时间 = 使用时间_In, 当前号码 = v_当前发票号, 剩余数量 = Nvl(剩余数量, 0) - 发票张数_In
    Where ID = 领用id_In
    Returning 剩余数量 Into n_返回数;
    If n_返回数 < 0 Then
      v_Err_Msg := '当前批次的剩余数量不足' || 发票张数_In || '张，无法完成收费票据分配操作。';
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

--89020:梁经伙,2015-10-10,添加输血医嘱并发处理
CREATE OR REPLACE Procedure Zl_病人医嘱记录_Delete
(
  --功能：删除指定医嘱,适用于门诊和住院。
  --参数：
  --      医嘱ID_IN：当前要删除的医嘱的ID(是可见行的单条医嘱ID，不是组ID)
  --      删相关_IN=0时,只删除指定ID的医嘱(医嘱编辑程序调用)。
  --          1.相关医嘱的同步删除及删除之后的序号调整由程序处理后调用对应过程。
  --          2.被删除的医嘱应该未校对,程序应已控制。
  --          3.病人医嘱状态的内容会自动删除；病人医嘱计价，病人医嘱发送未校对的没有记录。
  --      删相关_IN=1时,删除整条医嘱(管理界面调用)，如给药途径，检查组合，手术附项，中药配方。
  --          1.需要在过程中同时调整相关记录的序号。
  --          2.一并给药的只删除当前药品记录(不包括给药途径)。
  医嘱id_In 病人医嘱记录.Id%Type,
  删相关_In Number := 0
) Is
  v_状态            病人医嘱记录.医嘱状态%Type;
  v_相关id          病人医嘱记录.相关id%Type;
  v_病人id          病人医嘱记录.病人id%Type;
  v_挂号单          病人医嘱记录.挂号单%Type;
  v_主页id          病人医嘱记录.主页id%Type;
  v_婴儿            病人医嘱记录.婴儿%Type;
  v_序号            病人医嘱记录.序号%Type;
  v_内容            病人医嘱记录.医嘱内容%Type;
  v_路径执行id      病人路径执行.Id%Type;
  v_Other路径执行id 病人路径执行.Id%Type;
  v_路径执行方式    临床路径项目.执行方式%Type;
  v_内容要求        临床路径项目.内容要求%Type;
  v_Count           Number(5);
  v_路径记录id      病人临床路径.Id%Type;
  v_变异原因        病人路径执行.变异原因%Type;
  n_是否评估        Number(5);
  n_路径项目id      病人路径执行.项目id%Type;
  n_Islast          Number(5);
  n_Del_Count       Number(5);
  n_Del类型         Number(2); --0-只删除指定ID的医嘱，1-删除整条医嘱
  v_诊疗类别        病人医嘱记录.诊疗类别%Type;
  v_审核状态        病人医嘱记录.审核状态%Type;
  v_诊疗项目ID      病人医嘱记录.诊疗项目ID%Type;
  v_启用血库        zlParameters.参数值%Type;
  v_执行分类        诊疗项目目录.执行分类%type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --检查医嘱状态:并发操作
  Begin
    Select 病人id, 挂号单, 主页id, 婴儿, 医嘱状态, 相关id, 医嘱内容, 诊疗类别, 审核状态,诊疗项目ID
    Into v_病人id, v_挂号单, v_主页id, v_婴儿, v_状态, v_相关id, v_内容, v_诊疗类别, v_审核状态,v_诊疗项目ID
    From 病人医嘱记录
    Where ID = 医嘱id_In;
  Exception
    When Others Then
      Begin
        v_Error := '未发现要删除的医嘱记录，可能已被其他人删除。';
        Raise Err_Custom;
      End;
  End;
  If v_挂号单 Is Null Then
    If Not v_状态 In (1, 2, -1) Then
      v_Error := '医嘱"' || v_内容 || '"已经过校对，不能再删除。';
      Raise Err_Custom;
    End If;
  Else
    If v_状态 <> 1 Then
      v_Error := '医嘱"' || v_内容 || '"已经被发送或作废，不能删除。';
      Raise Err_Custom;
    End If;
  End If;

  --输血医嘱并发处理
  if v_诊疗类别 = 'K' and v_审核状态 in (2,5) then 
    --是否安装了血库系统
    Select count(1) into v_Count  From zlSystems Where 编号=2200;
    If Nvl(v_Count, 0) > 0 Then
      select 执行分类 into v_执行分类 from 诊疗项目目录 where  ID = v_诊疗项目ID;
      if not (nvl(v_执行分类,0) = 1) then
        --是否启用了血库管理系统
        Select Zl_Getsysparameter(236) into v_启用血库 From Dual;
        if Nvl(v_启用血库,'0') <> '0' then
          if v_审核状态 = 5 then
             v_Error := '正在配血，';
          else
             v_Error := '并且已完成配血，';
          end if;
          v_Error := '医嘱"' || v_内容 ||'"已被血库接收，' || v_Error || '不能删除，若需删除请与输血科联系。';
          Raise Err_Custom;
        end if;
      end if;
    end if;
  end if;

  Select Count(*)
  Into v_Count
  From 病人医嘱状态
  Where 医嘱id = 医嘱id_In And 操作类型 In (1, 11) And 签名id Is Not Null;
  If Nvl(v_Count, 0) > 0 Then
    v_Error := '医嘱"' || v_内容 || '"已经电子签名,不能删除。';
    Raise Err_Custom;
  End If;

  --判断删整组还是指定ID的医嘱
  If Nvl(删相关_In, 0) = 0 Then
    n_Del类型 := 0;
  Else
    If v_相关id Is Null Then
      --检查组合,手术及附加,中药配方,检验组合,以及独立医嘱
      Select Max(序号), Count(*) Into v_序号, n_Del_Count From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
      n_Del类型 := 1;
    Else
      --成药一并给药的情况(无申请)
      --先判断是否一并给药
      Select Count(*) Into n_Del_Count From 病人医嘱记录 Where 相关id = v_相关id;
      If n_Del_Count = 1 Then
        --单独给药:同时删除其给药途径
        Select Max(序号), Count(*) Into v_序号, n_Del_Count From 病人医嘱记录 Where ID = 医嘱id_In Or ID = v_相关id;
        n_Del类型 := 1;
      Else
        --一并给药:只删除当前药品
        n_Del_Count := 1;
        Select 序号 Into v_序号 From 病人医嘱记录 Where ID = 医嘱id_In;
        n_Del类型 := 0;
      End If;
    End If;
  End If;
  --
  Begin
    --如果不是路径的医嘱，则不查询路径执行表
    Select Count(1) Into v_Count From 病人路径医嘱 Where 病人医嘱id = 医嘱id_In;

    If v_Count > 0 Then
      --外连接是因为路径外项目的项目id是空
      --游标循环处理：存在同一天，同一条医嘱对应多个路径项目的情况。
      --必须生成的项目需要填写变异原因，不删项目，只删医嘱；非必须生成的项目，直接删除医嘱和对应项目
      For Rs In (Select a.Id, d.执行方式, d.内容要求, a.路径记录id, a.变异原因, a.项目id
                 From 病人路径执行 A, 病人路径医嘱 B, 临床路径项目 D
                 Where b.病人医嘱id = 医嘱id_In And b.路径执行id = a.Id And a.项目id = d.Id(+)) Loop
        v_路径执行id   := Rs.Id;
        v_路径执行方式 := Rs.执行方式;
        v_内容要求     := Rs.内容要求;
        v_路径记录id   := Rs.路径记录id;
        v_变异原因     := Rs.变异原因;
        n_路径项目id   := Rs.项目id;

        Select Count(1)
        Into n_是否评估
        From 病人路径执行 A, 病人路径评估 B
        Where a.路径记录id = b.路径记录id And a.阶段id = b.阶段id And a.天数 = b.天数 And a.Id = v_路径执行id;

        If n_是否评估 > 0 Then
          v_Error := '该医嘱对应的临床路径项目已经评估，请取消评估再删除。';
          Raise Err_Custom;
        End If;
        --生成时填了变异原因的必须适用的项目允许删除
        If Not v_路径执行方式 Is Null And v_变异原因 Is Null Then
          If v_路径执行方式 <> 3 Then
            --如果必须生成的项目，选择生成的医嘱还剩最后一个，则不允许删除
            If v_内容要求 = 1 Then
              --路径内外的医嘱进行一并给药时，可以删除原有的给药途径
              If Nvl(删相关_In, 0) = 0 And v_相关id Is Null Then
                Select Count(*)
                Into v_Count
                From 病人路径医嘱 A
                Where a.路径执行id = v_路径执行id And a.病人医嘱id <> 医嘱id_In;
              Else
                Select Count(*)
                Into v_Count
                From 病人路径医嘱 A
                Where a.路径执行id = v_路径执行id And
                      a.病人医嘱id Not In
                      (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In Or ID = v_相关id);
              End If;
              If v_Count = 0 Then
                v_Error := '该医嘱对应的临床路径项目不是必要时生成的，不能删除。';
                Raise Err_Custom;
              End If;
            Else
              --执行方式：0-无须执行(也可用于实现标签)，1-每天执行，2-至少执行一次，3-必要时执行,4-必须执行一次（所在阶段必须且仅执行一次）
              If v_路径执行方式 = 2 Or v_路径执行方式 = 4 Then
                --程序界面已经处理（当阶段为多天的情况下,当前天数不是阶段最后一天时,执行方式为2或4时，是允许不添加变异原因就删除医嘱的）
        Null;
        Else
                v_Error := '该医嘱对应的临床路径项目是必须生成的，需要添加变异原因才能删除。';
                Raise Err_Custom;
              End If;
            End If;
          End If;
        End If;

        --判断是否是最后一条医嘱，分几种情况，一组医嘱，一条医嘱
        If n_Del类型 = 0 Then
          Select Count(1)
          Into n_Islast
          From 病人临床路径 A, 病人路径执行 B, 病人路径医嘱 C
          Where a.Id = b.路径记录id And a.当前阶段id = b.阶段id And a.当前天数 = b.天数 And b.Id = c.路径执行id And a.Id = v_路径记录id;
        Else
          --n_Del类型=1的都是删整组医嘱，有可能传入的是相关ID=null的也有可能是传入的相关ID<>null的
          If v_相关id Is Null Then
            Select Decode(Count(1), 0, 1, 0)
            Into n_Islast
            From 病人临床路径 A, 病人路径执行 B, 病人路径医嘱 C, 病人医嘱记录 D
            Where a.Id = b.路径记录id And a.当前阶段id = b.阶段id And a.当前天数 = b.天数 And a.Id = v_路径记录id And c.病人医嘱id = d.Id And
                  c.路径执行id = b.Id And (d.Id <> 医嘱id_In And d.相关id <> 医嘱id_In);
          Else
            Select Decode(Count(1), 0, 1, 0)
            Into n_Islast
            From 病人临床路径 A, 病人路径执行 B, 病人路径医嘱 C
            Where a.Id = b.路径记录id And a.当前阶段id = b.阶段id And a.当前天数 = b.天数 And c.路径执行id = b.Id And a.Id = v_路径记录id And
                  (c.病人医嘱id <> 医嘱id_In And c.病人医嘱id <> v_相关id);
          End If;
        End If;

        If n_Islast = 1 Then
          --是最后一条项目的最后一条医嘱，就调用路径项目删除的过程
          If v_变异原因 Is Null Or Nvl(n_路径项目id, 0) = 0 Then
            Zl_病人路径生成_Delete(v_路径执行id);
          Else
            --必须生成但没有生成填写过变异原因的不删除项目
            Zl_病人路径生成_Delete(v_路径执行id, 2);
          End If;
        Else
          If n_Del类型 = 0 Then
            Delete From 病人路径医嘱 Where 病人医嘱id = 医嘱id_In And 路径执行id = v_路径执行id;

            --如果当前药品删除后，该执行id下只剩给药途径，则要该执行ID更改为一并给药中其他药品的执行ID
            If v_相关id Is Not Null Then
              Select Nvl(Max(c.Id), 0)
              Into v_Other路径执行id
              From 病人路径医嘱 A, 病人医嘱记录 B, 病人路径执行 C
              Where c.Id = a.路径执行id And b.相关id = v_相关id And b.Id <> 医嘱id_In And a.病人医嘱id = b.Id And
                    a.路径执行id <> v_路径执行id And c.登记时间 = (Select Max(d.登记时间)
                                                       From 病人路径执行 D, 病人医嘱记录 E, 病人路径医嘱 F
                                                       Where e.Id = f.病人医嘱id And d.Id = f.路径执行id And e.相关id = v_相关id And
                                                             f.路径执行id <> v_路径执行id);

              If v_Other路径执行id <> 0 Then
                Select Count(1)
                Into v_Count
                From 病人路径执行 A, 病人路径执行 B
                Where a.Id = v_Other路径执行id And b.Id = v_路径执行id And a.阶段id = b.阶段id And a.天数 = b.天数;

                If v_Count > 0 Then
                  Update 病人路径医嘱
                  Set 路径执行id = v_Other路径执行id
                  Where 病人医嘱id = v_相关id And 路径执行id = v_路径执行id And Not Exists
                   (Select 1 From 病人路径医嘱 C Where 路径执行id = v_路径执行id And 病人医嘱id <> v_相关id);
                End If;
              End If;
            End If;
          Else
            If v_相关id Is Null Then
              Delete From 病人路径医嘱
              Where 路径执行id = v_路径执行id And
                    病人医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In);
            Else
              --单独给药:同时删除其给药途径
              Delete From 病人路径医嘱
              Where 路径执行id = v_路径执行id And 病人医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or ID = v_相关id);
            End If;
          End If;
          --该项目对应的医嘱删除完后，则删除路径项目，如果变异原因不为空且不是路径外项目则不删除项目，只删医嘱
          If v_变异原因 Is Null Or Nvl(n_路径项目id, 0) = 0 Then
            Delete From 病人路径执行
            Where ID = v_路径执行id And Not Exists (Select 1 From 病人路径医嘱 Where 路径执行id = v_路径执行id);
          End If;
        End If;
      End Loop;
    End If;
  End;

  --删除关联诊断后删除医嘱
  If n_Del类型 = 0 Then
    Delete From 病人诊断医嘱 Where 医嘱id = 医嘱id_In;
    Delete From 病人医嘱记录 Where ID = 医嘱id_In;
  Else
    If v_相关id Is Null Then
      Delete From 病人诊断医嘱
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In);
      Delete From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
    Else
      --单独给药:同时删除其给药途径
      Delete From 病人诊断医嘱 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or ID = v_相关id);
      Delete From 病人医嘱记录 Where ID = 医嘱id_In Or ID = v_相关id;
    End If;
  End If;

  If Nvl(删相关_In, 0) <> 0 Then
    --调整序号
    Update 病人医嘱记录
    Set 序号 = 序号 - n_Del_Count
    Where 病人id = v_病人id And Nvl(主页id, 0) = Nvl(v_主页id, 0) And Nvl(挂号单, '空') = Nvl(v_挂号单, '空') And
          Nvl(婴儿, 0) = Nvl(v_婴儿, 0) And 序号 > v_序号;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_Delete;
/

--89297:梁唐彬,2015-10-20,医嘱发送自定义收费
CREATE OR REPLACE Function zl_fun_CustomExpenses
( 
  病人ID_In     In 病案主页.病人ID%Type, 
  主页ID_In     In 病案主页.主页ID%Type, 
  来源_In       In Number,
  医嘱ID_In     In 病人医嘱记录.ID%Type, 
  相关ID_In     In 病人医嘱记录.相关ID%Type, 
  期效_in       In 病人医嘱记录.医嘱期效%Type,
  医嘱频率_In   In 病人医嘱记录.执行频次%Type,
  诊疗项目ID_In In 诊疗项目目录.id%Type,
  收费细目ID_In In 收费项目目录.ID%Type,
  执行部门id_In In 部门表.ID%Type,
  诊疗类别_In   In 诊疗项目目录.类别%Type,
  收费类别_in   In 收费项目目录.类别%Type,
  医嘱总量_In   In 病人医嘱记录.总给予量%Type,
  单量_in       In 病人医嘱记录.单次用量%Type,
  计价数量_In   In 诊疗收费关系.收费数量%Type,
  费用性质_In   In 诊疗收费关系.费用性质%Type,
  计算方式_in   In 诊疗项目目录.计算方式%Type
) Return Varchar2 Is 
  ---------------------------------- 
  --功能：医嘱发送时，如果收费对照为自定义的，则会调用此过程，决定改收费项目是否收取，收取的次数；
  --规则： 
  --      1、医嘱发送时诊疗项目的收费方式为9-自定义时，当判断收费项目是否要收取时，会调用一次此过程；获取是否收取，以及收取的数量；
  --      2、如果多条医嘱的多个收费方式都是自定义，则会循环调用 
  --参数： 
  --      主页ID  :门诊传入0；住院=主页ID
  --      医嘱ID_In；当前循环的医嘱ID
  --      相关ID_In：当前循环的医嘱ID对应的组ID，如果本身就是组ID，则为空
  --      诊疗项目ID_In：当前循环到的医嘱对应的诊疗项目ID
  --      期效_in：1=临嘱，0-长嘱
  --      收费细目ID_In：当前循环到的医嘱对应的收费项目ID
  --      执行部门id_In=当前医嘱的执行科室；
  --      医嘱总量_In=当前医嘱的总量；
  --      费用性质_In=0-基础费用；1-床旁或术中加收；2-加班加收(下班时间和节假日)
  --      计算方式_in=0-不明确；1-计量执行(如药品、材料)；2-计时执行(如理疗)；3-计次(这种方式下，计算单位通常为"次")
  --返回：
  --      是否收取(1,0):收取数次   例如：1:1  即要收取，收取一个数次；0:0即代表不收取(注意最后的收费数量会乘以计价数量)
  --      收取数次可不指定；则按收费对照中默认的数量来收取，可返回是否收取：1  /  0  即可；
  ----------------------------------- 
  Err_Custom Exception; 
  v_Error Varchar2(255); 
  
Begin  
   
  Return '1'; 
 
Exception 
  When Err_Custom Then 
    Return Null; 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End zl_fun_CustomExpenses;
/
--89885:余伟节,2015-11-06,病人信息合并
CREATE OR REPLACE Procedure Zl_病人信息_Merge
(
  A病人id_In    病人信息.病人id%Type, --要合并的病人信息
  B病人id_In    病人信息.病人id%Type, --要保留的病人信息
  合并原因_In   病人合并记录.合并原因%Type,
  操作员姓名_In 人员表.姓名%Type,
  强制保留_In   Number := 0
  --标准版
  ----------------------------------------------------------------------------
  --病人信息,病案主页,病案主页从表,病人变动记录,特殊病人
  --门诊病案记录,住院病案记录,床位状况记录
  --医保病人档案,保险模拟结算,保险结算记录,帐户年度信息
  --病人余额,病人未结费用,住院费用记录,门诊费用记录,病人预交记录,病人结帐记录,未发药品记录
  --病人挂号记录,病人过敏药物,病人过敏记录,病人诊断记录,诊断情况
  --病人医嘱记录,病人手麻记录
  --病人社区信息
  
  --后备表：
  --H病人结帐记录,H病人预交记录,H住院费用记录,H门诊费用记录
  --H病人医嘱记录,H病人诊断记录,H病人过敏记录
  --H病人病历记录,H病人手麻记录
  
  --病案系统
  ----------------------------------------------------------------------------
  --病人费用,随诊记录,借阅记录
  --新生儿诊断记录,病人分娩信息
  --诊断符合情况,病案评分结果
  
) As
  --病人相关表
  Cursor c_Patitable Is
    Select a.Table_Name, Max(Decode(b.Column_Name, '病人ID', 1, 0)) As 病人id,
           Max(Decode(b.Column_Name, '主页ID', 1, 0)) As 主页id
    From User_Tables A, User_Tab_Columns B
    Where a.Table_Name = b.Table_Name And b.Column_Name In ('病人ID', '主页ID') And
          a.Table_Name Not In
          ('病人信息', '病案主页', '病案主页从表', '病人变动记录', '特殊病人', '门诊病案记录', '住院病案记录', '床位状况记录', '医保病人档案', '医保病人关联表', '保险模拟结算',
           '帐户年度信息', '病人余额', '病人未结费用', '住院费用记录', '门诊费用记录', '病人预交记录', '病人结帐记录', '未发药品记录', '病人挂号记录', '病人过敏药物', '病人过敏记录',
           '病人诊断记录', '诊断情况', '病人医嘱记录', '病人手麻记录', '病人费用', '随诊记录', '借阅记录', '病人分娩信息', '诊断符合情况', '病案评分结果', '病人担保记录', '病人社区信息',
           '病人免疫记录', '病人信息从表', '病人医疗卡属性') Having Max(Decode(b.Column_Name, '病人ID', 1, 0)) <> 0
    Group By a.Table_Name;

  --数组定义
  Type Array_Patitable Is Table Of Varchar2(100) Index By Binary_Integer;
  Arronbase Array_Patitable;
  Arronpage Array_Patitable;
  v_Loop    Number;
  n_Have    Number;

  -------------------------------------------------------
  --被合并的病人(住院号可能每次新产生,多次住院取最近一次)
  Cursor c_Infoa Is
    Select a.病人id, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.其他证件, a.身份,
           a.职业, a.民族, a.国籍, a.籍贯, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.监护人, a.联系人姓名, a.联系人关系, a.联系人地址,
           a.联系人电话, a.户口地址, a.户口地址邮编, a.Email, a.Qq, a.合同单位id, a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人, a.担保额,
           a.担保性质, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.当前床号, a.入院时间, a.出院时间, a.在院, a.Ic卡号, a.健康号,
           a.医保号, a.险类, a.查询密码, a.登记时间, a.停用时间, a.锁定, a.联系人身份证号, b.主页id, b.入院日期, b.出院日期
    From 病人信息 A, 病案主页 B
    Where a.病人id = b.病人id(+) And a.病人id = A病人id_In
    Order By 出院日期 Desc, 主页id Desc;
  r_Infoa c_Infoa%RowType;

  --要保留的病人(住院号可能每次新产生,多次住院取最近一次)
  Cursor c_Infob Is
    Select a.病人id, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.其他证件, a.身份,
           a.职业, a.民族, a.国籍, a.籍贯, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.监护人, a.联系人姓名, a.联系人关系, a.联系人地址,
           a.联系人电话, a.户口地址, a.户口地址邮编, a.Email, a.Qq, a.合同单位id, a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人, a.担保额,
           a.担保性质, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.当前床号, a.入院时间, a.出院时间, a.在院, a.Ic卡号, a.健康号,
           a.医保号, a.险类, a.查询密码, a.登记时间, a.停用时间, a.锁定, a.联系人身份证号, b.主页id, b.入院日期, b.出院日期
    From 病人信息 A, 病案主页 B
    Where a.病人id = b.病人id(+) And a.病人id = B病人id_In
    Order By 出院日期 Desc, 主页id Desc;
  r_Infob c_Infob%RowType;

  --合并后的信息
  Cursor c_Info(v_病人id 病人信息.病人id%Type) Is
    Select 病人id, 主页id, (Select Nvl(Max(主页id), 0) From 病案主页 Where 病人id = v_病人id) 最大主页id, 住院号, 病人性质, 医疗付款方式, 费别, 再入院,
           入院病区id, 入院科室id, 医疗小组id, 入院日期, 入院病况, 入院方式, 入院属性, 二级院转入, 住院目的, 入院病床, 是否陪伴, 当前病况, 当前病区id, 护理等级id, 出院科室id, 出院病床,
           出院日期, 住院天数, 出院方式, 是否确诊, 确诊日期, 新发肿瘤, 血型, 抢救次数, 成功次数, 随诊标志, 随诊期限, 尸检标志, 门诊医师, 责任护士, 住院医师, 病案号, 编目员编号, 编目员姓名,
           编目日期, 状态, 费用和, 年龄, 身高, 体重, 婚姻状况, 职业, 国籍, 学历, 单位电话, 单位邮编, 单位地址, 区域, 家庭地址, 家庭电话, 家庭地址邮编, 联系人姓名, 联系人关系, 联系人地址,
           联系人电话, 联系人身份证号, 户口地址, 户口地址邮编, 中医治疗类别, 险类, 社区, 审核标志, 审核人, 审核日期, 是否上传, 数据转出, 登记人, 登记时间, 备注, 病案状态, 病人类型
    From 病案主页
    Where 主页id = (Select Nvl(Max(主页id), 0)
                  From 病案主页
                  Where 病人id = v_病人id And Not Exists (Select 主页id From 病案主页 Where 病人id = v_病人id And 主页id = 0)) And
          病人id = v_病人id;
  r_Info c_Info%RowType;

  --合并两个住院病人
  Cursor c_Mergepati Is
    Select a.姓名, a.门诊号, a.住院号 当前住院号, b.病人id, b.主页id, b.住院号, b.病人性质, b.医疗付款方式, b.费别, b.再入院, b.入院病区id, b.入院科室id,
           b.医疗小组id, b.入院日期, b.入院病况, b.入院方式, b.入院属性, b.二级院转入, b.住院目的, b.入院病床, b.是否陪伴, b.当前病况, b.当前病区id, b.护理等级id,
           b.出院科室id, b.出院病床, b.出院日期, b.住院天数, b.出院方式, b.是否确诊, b.确诊日期, b.新发肿瘤, b.血型, b.抢救次数, b.成功次数, b.随诊标志, b.随诊期限,
           b.尸检标志, b.门诊医师, b.责任护士, b.住院医师, b.病案号, b.编目员编号, b.编目员姓名, b.编目日期, b.状态, b.费用和, b.性别, b.年龄, b.身高, b.体重, b.婚姻状况,
           b.职业, b.国籍, b.学历, b.单位电话, b.单位邮编, b.单位地址, b.区域, b.家庭地址, b.家庭电话, b.家庭地址邮编, b.联系人姓名, b.联系人关系, b.联系人地址, b.联系人电话,
           b.联系人身份证号, b.户口地址, b.户口地址邮编, b.中医治疗类别, b.险类, b.社区, b.审核标志, b.审核人, b.审核日期, b.是否上传, b.数据转出, b.登记人, b.登记时间, b.备注,
           b.病案状态, b.病人类型, b.封存时间, b.路径状态, b.单病种, b.婴儿科室id, b.婴儿病区id, b.母婴转科标志, b.医嘱重整时间
    From 病人信息 A, 病案主页 B
    Where a.病人id = b.病人id And a.病人id In (A病人id_In, B病人id_In)
    Order By b.入院日期 Desc, Nvl(b.出院日期, Sysdate) Desc;

  v_保留id 病人信息.病人id%Type;
  v_合并id 病人信息.病人id%Type;
  v_门诊号 病人信息.门诊号%Type;
  v_住院号 病人信息.住院号%Type;
  --病人未结费用(门诊部份)
  Cursor c_Owe(v_病人id 病人信息.病人id%Type) Is
    Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, Sum(金额) As 金额
    From 病人未结费用
    Where 主页id Is Null And 病人id = v_病人id
    Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径;

  --病人余额
  Cursor c_Spare(v_病人id 病人信息.病人id%Type) Is
    Select 性质, 类型, 预交余额, 费用余额 From 病人余额 Where 病人id = v_病人id;

  --医保病人档案
  Cursor c_Insure(v_病人id 病人信息.病人id%Type) Is
    Select * From 保险帐户 Where 病人id = v_病人id Order By 险类;

  --要保留的医保病人档案
  Cursor c_Keepinsure
  (
    v_病人id 病人信息.病人id%Type,
    v_险类   医保病人档案.险类%Type
  ) Is
    Select * From 保险帐户 Where 病人id = v_病人id And 险类 = v_险类;
  r_Keepinsure c_Keepinsure%RowType;

  Cursor c_Year
  (
    v_病人id 病人信息.病人id%Type,
    v_险类   医保病人档案.险类%Type
  ) Is
    Select * From 帐户年度信息 Where 病人id = v_病人id And 险类 = v_险类;

  v_原信息   病人合并记录.原信息%Type;
  v_Count    Number;
  n_Readonly Number;
  v_Sql      Varchar2(1000);

  n_主页id       病人信息.主页id%Type;
  v_Error        Varchar2(255);
  n_担保额       病人担保记录.担保额%Type;
  v_担保人       病人信息.担保人%Type;
  n_担保性质     病人担保记录.担保性质%Type;
  n_Row          Number;
  n_独立病案     Number;
  n_每次新住院号 Number;
  n_Max主页id    Number;
  n_Cnt主页id    Number;
  n_Cur主页id    Number;
  n_Cnt住院次数  Number;
  n_Cur住院次数  Number;
  n_Max住院次数  Number;
  n_Loop主页id   病人信息.主页id%Type;

  n_Lengthb Number;
  Err_Custom Exception;
Begin
  Begin
    Select 只读 Into n_Readonly From zlBakSpaces Where 当前 = 1;
  Exception
    When Others Then
      Null;
  End;
  If n_Readonly = 1 Then
    n_Readonly := 0;
    For r_Bak In (Select a.表名 Table_Name
                  From Zltools.Zlbaktables A, User_Constraints B
                  Where a.表名 = b.Table_Name And b.r_Constraint_Name = '病人信息_PK' And b.Constraint_Type = 'R') Loop
      v_Sql := 'Select Count(病人Id) From ' || r_Bak.Table_Name || ' Where 病人Id In(:1,:2)';
      Execute Immediate v_Sql
        Into n_Readonly
        Using A病人id_In, B病人id_In;
      If n_Readonly > 0 Then
        v_Error := '病人在只读的当前转储空间存在数据,不能进行合并!';
        Raise Err_Custom;
      End If;
    End Loop;
  End If;

  --程序中已检查：
  --1.选择了同一个病人
  --2.两个住院病人先入院的却在院(包括两个都在院)。
  --3.两个住院病人的住院期间存在交叉的情况
  --4.医保病人存在未结费用

  --先锁定病人不允许进行其他业务
  Zl_病人信息_锁定(A病人id_In, 1);
  Zl_病人信息_锁定(B病人id_In, 1);

  Open c_Infoa;
  Fetch c_Infoa
    Into r_Infoa;
  If c_Infoa%RowCount = 0 Then
    Close c_Infoa;
    v_Error := '没有发现被合并的病人信息！';
    Raise Err_Custom;
  End If;

  Open c_Infob;
  Fetch c_Infob
    Into r_Infob;
  If c_Infob%RowCount = 0 Then
    Close c_Infob;
    v_Error := '没有发现要保留的病人信息！';
    Raise Err_Custom;
  End If;

  --读取其它相关病人表到数组
  For r_Patitable In c_Patitable Loop
    If r_Patitable.主页id = 0 Then
      Arronbase(Arronbase.Count + 1) := r_Patitable.Table_Name;
    Else
      Arronpage(Arronpage.Count + 1) := r_Patitable.Table_Name;
    End If;
  End Loop;

  --以先住院或先登记的病人ID作为实际上要保留的病人ID
  If Nvl(强制保留_In, 0) = 1 Then
    v_保留id := B病人id_In;
  Else
    Select 病人id
    Into v_保留id
    From (Select /*+ CHOOSE */
            a.病人id
           From 病人信息 A, 病案主页 B
           Where a.病人id = b.病人id(+) And a.病人id In (A病人id_In, B病人id_In)
           Order By Nvl(b.入院日期, To_Date('3000-01-01', 'YYYY-MM-DD')), Nvl(b.出院日期, To_Date('3000-01-01', 'YYYY-MM-DD')),
                    a.登记时间, a.病人id --住院病人优先
           )
    Where Rownum = 1;
  End If;

  --先确定病案号的模式
  Select Zl_To_Number(Nvl(zl_GetSysParameter(39), '0')) Into n_独立病案 From Dual;
  --住院号模式
  Select Zl_To_Number(Nvl(zl_GetSysParameter(145), '0')) Into n_每次新住院号 From Dual;

  --另外一个就是实际最后要删除的病人ID
  If v_保留id = A病人id_In Then
    v_合并id := B病人id_In;
    --问题27445 保留指定病人的门诊号、住院号、医保号
    v_门诊号 := Nvl(r_Infob.门诊号, r_Infoa.门诊号);
    v_住院号 := Nvl(r_Infob.住院号, r_Infoa.住院号);
  Else
    v_合并id := A病人id_In;
    v_门诊号 := Nvl(r_Infob.门诊号, r_Infoa.门诊号);
    v_住院号 := Nvl(r_Infob.住院号, r_Infoa.住院号);
  End If;

  ---记录合并操作,在后面会根据r_PatiTable把合并病人的合并记录更新为保留病人的
  v_原信息 := v_合并id || ',' || r_Infoa.门诊号 || ',' || r_Infoa.住院号 || ',' || r_Infoa.就诊卡号 || ',' || r_Infoa.姓名 || ',' ||
           r_Infoa.性别 || ',' || r_Infoa.年龄 || ',' || To_Char(r_Infoa.出生日期, 'yyyy-mm-dd') || ',' || r_Infoa.身份证号 || ',' ||
           r_Infoa.婚姻状况 || ',' || r_Infoa.职业 || ',' || r_Infoa.家庭地址;
  Insert Into 病人合并记录
    (病人id, 原信息, 合并原因, 操作员姓名, 合并时间)
  Values
    (v_保留id, v_原信息, 合并原因_In, 操作员姓名_In, Sysdate);

  --开始合并
  --84398修改将住院次数计算放在外面，因需要考虑门诊和住院病人合并
  --10.34开始,住院次数不包含留关病人,合并后的住院次数=保留病人住院次数+合并病人正常入院的次数
  Select Nvl(住院次数, 0) Into n_Cur住院次数 From 病人信息 Where 病人id = v_保留id;
  Select Count(*) Into n_Cnt住院次数 From 病案主页 Where 病人id = v_合并id And 主页id <> 0 And 病人性质 = 0;
  n_Max住院次数 := n_Cur住院次数 + n_Cnt住院次数;
  --处理病案主页部份(涉及病人ID,主页ID字段的表)
  If (r_Infoa.主页id Is Not Null And r_Infob.主页id Is Not Null) Or (强制保留_In = 1 And r_Infoa.主页id Is Not Null) Then
    If r_Infoa.主页id = 0 And r_Infob.主页id = 0 Then
      Close c_Infoa;
      Close c_Infob;
      v_Error := '两个预约病人不能进行病人合并操作！';
      Raise Err_Custom;
    Elsif r_Infoa.主页id = 0 Then
      If r_Infob.入院日期 Is Not Null And r_Infob.出院日期 Is Null Then
        Close c_Infoa;
        Close c_Infob;
        v_Error := '预约病人和在院病人不能进行病人合并操作！';
        Raise Err_Custom;
      End If;
    Elsif r_Infob.主页id = 0 Then
      If r_Infoa.入院日期 Is Not Null And r_Infoa.出院日期 Is Null Then
        Close c_Infoa;
        Close c_Infob;
        v_Error := '预约病人和在院病人不能进行病人合并操作！';
        Raise Err_Custom;
      End If;
    End If;
    --求两个病人总共的住院就诊次数
    Select Count(*) Into v_Count From 病案主页 Where 病人id In (A病人id_In, B病人id_In) And 主页id <> 0;
    --因为10.19开始，入院时允许修改主页id，所以最大主页ID可能大于总的住院就诊次数
    Select Max(主页id) Into n_Max主页id From 病案主页 Where 病人id = v_保留id And 主页id <> 0;
    Select Count(*) Into n_Cnt主页id From 病案主页 Where 病人id = v_合并id And 主页id <> 0;
    If n_Max主页id + n_Cnt主页id > v_Count Then
      v_Count := n_Max主页id + n_Cnt主页id;
    End If;
    --求实际要更新的主页截至值,以前用v_Count >= n_Max主页id判断存在一个问题（对于两个病人多次交叉入院，可能导致A,B病人部分就诊次数没有更新）
    Select Nvl(Max(主页id), 0)
    Into n_Loop主页id
    From 病案主页 A, (Select Min(入院日期) 入院日期 From 病案主页 Where 病人id = v_合并id) B
    Where a.病人id = v_保留id And a.入院日期 < b.入院日期;
  
    For r_Merge In c_Mergepati Loop
      If Not (r_Merge.病人id = v_保留id And r_Merge.主页id = v_Count) And v_Count <> 0 Then
        --该病案主页要删除时,不能是已编目了的。
        If r_Merge.编目日期 Is Not Null Then
          Close c_Infoa;
          Close c_Infob;
          If r_Merge.当前住院号 Is Null Then
            v_Error := '病人' || r_Merge.姓名 || '(病人ID=' || r_Merge.病人id || ')存在已编目的病案,不允许合并该病人。';
          Else
            v_Error := '病人' || r_Merge.姓名 || '(病人ID=' || r_Merge.病人id || ',住院号=' || r_Merge.当前住院号 ||
                       ')存在已编目的病案,不允许合并该病人。';
          End If;
          Raise Err_Custom;
        End If;
        If v_Count >= Nvl(n_Loop主页id, 0) Then
          If r_Merge.主页id = 0 Then
            n_Cur主页id := 0;
            Update 病案主页
            Set 病人性质 = r_Merge.病人性质, 医疗付款方式 = r_Merge.医疗付款方式, 费别 = r_Merge.费别, 再入院 = r_Merge.再入院,
                入院病区id = r_Merge.入院病区id, 入院科室id = r_Merge.入院科室id, 入院日期 = r_Merge.入院日期, 入院病况 = r_Merge.入院病况,
                入院方式 = r_Merge.入院方式, 二级院转入 = r_Merge.二级院转入, 住院目的 = r_Merge.住院目的, 入院病床 = r_Merge.入院病床,
                是否陪伴 = r_Merge.是否陪伴, 当前病况 = r_Merge.当前病况, 当前病区id = r_Merge.当前病区id, 护理等级id = r_Merge.护理等级id,
                出院科室id = r_Merge.出院科室id, 出院病床 = r_Merge.出院病床, 出院日期 = r_Merge.出院日期, 住院天数 = r_Merge.住院天数,
                出院方式 = r_Merge.出院方式, 是否确诊 = r_Merge.是否确诊, 确诊日期 = r_Merge.确诊日期, 新发肿瘤 = r_Merge.新发肿瘤, 血型 = r_Merge.血型,
                抢救次数 = r_Merge.抢救次数, 成功次数 = r_Merge.成功次数, 随诊标志 = r_Merge.随诊标志, 随诊期限 = r_Merge.随诊期限, 尸检标志 = r_Merge.尸检标志,
                门诊医师 = r_Merge.门诊医师, 责任护士 = r_Merge.责任护士, 住院医师 = r_Merge.住院医师, 编目员编号 = r_Merge.编目员编号,
                编目员姓名 = r_Merge.编目员姓名, 编目日期 = r_Merge.编目日期, 状态 = r_Merge.状态, 费用和 = r_Merge.费用和, 姓名 = r_Merge.姓名,
                性别 = r_Merge.性别, 年龄 = r_Merge.年龄, 婚姻状况 = r_Merge.婚姻状况, 职业 = r_Merge.职业, 国籍 = r_Merge.国籍, 学历 = r_Merge.学历,
                单位电话 = r_Merge.单位电话, 单位邮编 = r_Merge.单位邮编, 单位地址 = r_Merge.单位地址, 区域 = r_Merge.区域, 家庭地址 = r_Merge.家庭地址,
                家庭电话 = r_Merge.家庭电话, 家庭地址邮编 = r_Merge.家庭地址邮编, 户口地址 = r_Merge.户口地址, 户口地址邮编 = r_Merge.户口地址邮编,
                联系人姓名 = r_Merge.联系人姓名, 联系人关系 = r_Merge.联系人关系, 联系人地址 = r_Merge.联系人地址, 联系人电话 = r_Merge.联系人电话,
                中医治疗类别 = r_Merge.中医治疗类别, 登记人 = r_Merge.登记人, 登记时间 = r_Merge.登记时间, 险类 = r_Merge.险类, 审核标志 = r_Merge.审核标志,
                是否上传 = r_Merge.是否上传, 备注 = r_Merge.备注, 数据转出 = r_Merge.数据转出, 病案号 = r_Merge.病案号,
                住院号 = Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号),病人类型 = r_Merge.病人类型,
                封存时间 = r_Merge.封存时间, 路径状态 = r_Merge.路径状态, 单病种 = r_Merge.单病种, 婴儿科室id = r_Merge.婴儿科室id,
                婴儿病区id = r_Merge.婴儿病区id, 母婴转科标志 = r_Merge.母婴转科标志, 医嘱重整时间 = r_Merge.医嘱重整时间
            Where 病人id = v_保留id And 主页id = n_Cur主页id;
            If Sql%RowCount = 0 Then
              Insert Into 病案主页
                (病人id, 主页id, 病人性质, 医疗付款方式, 费别, 再入院, 入院病区id, 入院科室id, 入院日期, 入院病况, 入院方式, 二级院转入, 住院目的, 入院病床, 是否陪伴, 当前病况,
                 当前病区id, 护理等级id, 出院科室id, 出院病床, 出院日期, 住院天数, 出院方式, 是否确诊, 确诊日期, 新发肿瘤, 血型, 抢救次数, 成功次数, 随诊标志, 随诊期限, 尸检标志,
                 门诊医师, 责任护士, 住院医师, 编目员编号, 编目员姓名, 编目日期, 状态, 费用和, 姓名, 性别, 年龄, 婚姻状况, 职业, 国籍, 学历, 单位电话, 单位邮编, 单位地址, 区域, 家庭地址,
                 家庭电话, 家庭地址邮编, 户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 中医治疗类别, 登记人, 登记时间, 险类, 审核标志, 是否上传, 备注, 数据转出,
                 病案号, 住院号,病人类型, 封存时间, 路径状态, 单病种, 婴儿科室id, 婴儿病区id, 母婴转科标志, 医嘱重整时间)
              Values
                (v_保留id, n_Cur主页id, r_Merge.病人性质, r_Merge.医疗付款方式, r_Merge.费别, r_Merge.再入院, r_Merge.入院病区id,
                 r_Merge.入院科室id, r_Merge.入院日期, r_Merge.入院病况, r_Merge.入院方式, r_Merge.二级院转入, r_Merge.住院目的, r_Merge.入院病床,
                 r_Merge.是否陪伴, r_Merge.当前病况, r_Merge.当前病区id, r_Merge.护理等级id, r_Merge.出院科室id, r_Merge.出院病床, r_Merge.出院日期,
                 r_Merge.住院天数, r_Merge.出院方式, r_Merge.是否确诊, r_Merge.确诊日期, r_Merge.新发肿瘤, r_Merge.血型, r_Merge.抢救次数,
                 r_Merge.成功次数, r_Merge.随诊标志, r_Merge.随诊期限, r_Merge.尸检标志, r_Merge.门诊医师, r_Merge.责任护士, r_Merge.住院医师,
                 r_Merge.编目员编号, r_Merge.编目员姓名, r_Merge.编目日期, r_Merge.状态, r_Merge.费用和, r_Merge.姓名, r_Merge.性别, r_Merge.年龄,
                 r_Merge.婚姻状况, r_Merge.职业, r_Merge.国籍, r_Merge.学历, r_Merge.单位电话, r_Merge.单位邮编, r_Merge.单位地址, r_Merge.区域,
                 r_Merge.家庭地址, r_Merge.家庭电话, r_Merge.家庭地址邮编, r_Merge.户口地址, r_Merge.户口地址邮编, r_Merge.联系人姓名, r_Merge.联系人关系,
                 r_Merge.联系人地址, r_Merge.联系人电话, r_Merge.中医治疗类别, r_Merge.登记人, r_Merge.登记时间, r_Merge.险类, r_Merge.审核标志,
                 r_Merge.是否上传, r_Merge.备注, r_Merge.数据转出, r_Merge.病案号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号),
                 r_Merge.病人类型, r_Merge.封存时间, r_Merge.路径状态, r_Merge.单病种, r_Merge.婴儿科室id, r_Merge.婴儿病区id,
                 r_Merge.母婴转科标志, r_Merge.医嘱重整时间);
            End If;
          Else
            n_Cur主页id := v_Count;
            Insert Into 病案主页
              (病人id, 主页id, 病人性质, 医疗付款方式, 费别, 再入院, 入院病区id, 入院科室id, 入院日期, 入院病况, 入院方式, 二级院转入, 住院目的, 入院病床, 是否陪伴, 当前病况,
               当前病区id, 护理等级id, 出院科室id, 出院病床, 出院日期, 住院天数, 出院方式, 是否确诊, 确诊日期, 新发肿瘤, 血型, 抢救次数, 成功次数, 随诊标志, 随诊期限, 尸检标志, 门诊医师,
               责任护士, 住院医师, 编目员编号, 编目员姓名, 编目日期, 状态, 费用和, 姓名, 性别, 年龄, 婚姻状况, 职业, 国籍, 学历, 单位电话, 单位邮编, 单位地址, 区域, 家庭地址, 家庭电话,
               家庭地址邮编, 户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 中医治疗类别, 登记人, 登记时间, 险类, 审核标志, 是否上传, 备注, 数据转出, 病案号, 住院号,
               病人类型, 封存时间, 路径状态, 单病种, 婴儿科室id, 婴儿病区id, 母婴转科标志, 医嘱重整时间)
            Values
              (v_保留id, n_Cur主页id, r_Merge.病人性质, r_Merge.医疗付款方式, r_Merge.费别, r_Merge.再入院, r_Merge.入院病区id, r_Merge.入院科室id,
               r_Merge.入院日期, r_Merge.入院病况, r_Merge.入院方式, r_Merge.二级院转入, r_Merge.住院目的, r_Merge.入院病床, r_Merge.是否陪伴,
               r_Merge.当前病况, r_Merge.当前病区id, r_Merge.护理等级id, r_Merge.出院科室id, r_Merge.出院病床, r_Merge.出院日期, r_Merge.住院天数,
               r_Merge.出院方式, r_Merge.是否确诊, r_Merge.确诊日期, r_Merge.新发肿瘤, r_Merge.血型, r_Merge.抢救次数, r_Merge.成功次数,
               r_Merge.随诊标志, r_Merge.随诊期限, r_Merge.尸检标志, r_Merge.门诊医师, r_Merge.责任护士, r_Merge.住院医师, r_Merge.编目员编号,
               r_Merge.编目员姓名, r_Merge.编目日期, r_Merge.状态, r_Merge.费用和, r_Merge.姓名, r_Merge.性别, r_Merge.年龄, r_Merge.婚姻状况,
               r_Merge.职业, r_Merge.国籍, r_Merge.学历, r_Merge.单位电话, r_Merge.单位邮编, r_Merge.单位地址, r_Merge.区域, r_Merge.家庭地址,
               r_Merge.家庭电话, r_Merge.家庭地址邮编, r_Merge.户口地址, r_Merge.户口地址邮编, r_Merge.联系人姓名, r_Merge.联系人关系, r_Merge.联系人地址,
               r_Merge.联系人电话, r_Merge.中医治疗类别, r_Merge.登记人, r_Merge.登记时间, r_Merge.险类, r_Merge.审核标志, r_Merge.是否上传,
               r_Merge.备注, r_Merge.数据转出, r_Merge.病案号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号),r_Merge.病人类型,
               r_Merge.封存时间, r_Merge.路径状态, r_Merge.单病种, r_Merge.婴儿科室id, r_Merge.婴儿病区id, r_Merge.母婴转科标志, r_Merge.医嘱重整时间);
          End If;
        Else
          Exit;
        End If;
      
        --更新病人相关表的病人指向
        ---------------------------------------------------------------
        --病人变动记录
        Update 病人变动记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病案主页从表
        Update 病案主页从表
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --住院费用记录
        Update 住院费用记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id,
            标识号 = Nvl(Decode(门诊标志, 1, v_门诊号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号)), 标识号)
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H住院费用记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id,
            标识号 = Nvl(Decode(门诊标志, 1, v_门诊号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号)), 标识号)
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        --门诊费用记录
        --Update 门诊费用记录
        --Set 病人id = v_保留id,
        --    标识号 = Nvl(Decode(门诊标志, 1, v_门诊号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号)), 标识号)
        --Where 病人id = r_Merge.病人id;
        --Update H门诊费用记录
        --Set 病人id = v_保留id,
        --    标识号 = Nvl(Decode(门诊标志, 1, v_门诊号, Decode(n_每次新住院号, 1, r_Merge.住院号, v_住院号)), 标识号)
        --Where 病人id = r_Merge.病人id;
      
        --病人预交记录
        Update 病人预交记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人预交记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人未结费用
        Update 病人未结费用
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --未发药品记录
        Update 未发药品记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --诊断情况
        Update 诊断情况
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --保险结算记录(病人ID和非住院病人一起在后面处理)
        Update 保险结算记录 Set 主页id = n_Cur主页id Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --保险模拟结算
        Update 保险模拟结算
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人医嘱记录(ZLHIS+)
        Update 病人医嘱记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人医嘱记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人过敏记录(ZLHIS+)
        Update 病人过敏记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人过敏记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人诊断记录(ZLHIS+)
        Update 病人诊断记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人诊断记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人手麻记录(ZLHIS+)
        Update 病人手麻记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
        Update H病人手麻记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病人担保记录(zlhis+)
        Update 病人担保记录
        Set 病人id = v_保留id, 主页id = n_Cur主页id
        Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      
        --病案系统的表
        Begin
          v_Sql := 'Update 病人费用 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update 随诊记录 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update 诊断符合情况 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update 病案评分结果 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Insert Into 病人分娩信息(病人ID,主页ID,胎儿次序,分娩方式,出生胎位,分娩情况,出生缺陷,婴儿性别,婴儿体重,Apgar评分) ' ||
                   'Select :1,:2,胎儿次序,分娩方式,出生胎位,分娩情况,出生缺陷,婴儿性别,婴儿体重,Apgar评分 From 病人分娩信息 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Delete From 病人分娩信息 Where 病人ID=:1 And 主页ID=:2';
          Execute Immediate v_Sql
            Using r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update 借阅记录 Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        Exception
          When Others Then
            Null;
        End;
      
        --其它病案主页相关表
        For v_Loop In 1 .. Arronpage.Count Loop
          v_Sql := 'Update ' || Arronpage(v_Loop) || ' Set 病人ID=:1,主页ID=:2 Where 病人ID=:3 And 主页ID=:4';
          Execute Immediate v_Sql
            Using v_保留id, n_Cur主页id, r_Merge.病人id, r_Merge.主页id;
        End Loop;
      
        --删除已调整后的病案主页
        Delete From 病案主页 Where 病人id = r_Merge.病人id And 主页id = r_Merge.主页id;
      End If;
      If r_Merge.主页id <> 0 Then
        v_Count := v_Count - 1;
      End If;
    End Loop;
  End If;

  --不涉及主页ID部份的更改(无主页ID或主页ID可能为空)
  ---------------------------------------------------------------
  --住院费用记录
  Update 住院费用记录
  Set 病人id = v_保留id, 标识号 = Nvl(Decode(门诊标志, 2, v_住院号, v_门诊号), 标识号)
  Where 病人id = v_合并id;
  Update H住院费用记录
  Set 病人id = v_保留id, 标识号 = Nvl(Decode(门诊标志, 2, v_住院号, v_门诊号), 标识号)
  Where 病人id = v_合并id;
  --门诊费用记录
  Update 门诊费用记录
  Set 病人id = v_保留id, 标识号 = Nvl(Decode(门诊标志, 2, v_住院号, v_门诊号), 标识号)
  Where 病人id = v_合并id;
  Update H门诊费用记录
  Set 病人id = v_保留id, 标识号 = Nvl(Decode(门诊标志, 2, v_住院号, v_门诊号), 标识号)
  Where 病人id = v_合并id;

  --病人预交记录
  Update 病人预交记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;
  Update H病人预交记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --未发药品记录
  Update 未发药品记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --诊断情况
  Update 诊断情况 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --病人医嘱记录(ZLHIS+)
  Update 病人医嘱记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;
  Update H病人医嘱记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --病人过敏记录(ZLHIS+):主页ID可能是挂号ID
  Update 病人过敏记录 Set 病人id = v_保留id Where 病人id = v_合并id;
  Update H病人过敏记录 Set 病人id = v_保留id Where 病人id = v_合并id;

  --病人诊断记录(ZLHIS+):主页ID可能是挂号ID
  Update 病人诊断记录 Set 病人id = v_保留id Where 病人id = v_合并id;
  Update H病人诊断记录 Set 病人id = v_保留id Where 病人id = v_合并id;

  --病人手麻记录(ZLHIS+)
  Update 病人手麻记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;
  Update H病人手麻记录 Set 病人id = v_保留id Where 病人id = v_合并id And 主页id Is Null;

  --病人挂号记录(ZLHIS+)
  Update 病人挂号记录 Set 病人id = v_保留id, 门诊号 = Nvl(v_门诊号, 门诊号) Where 病人id = v_合并id;

  --病人结帐记录
  Update 病人结帐记录 Set 病人id = v_保留id Where 病人id = v_合并id;
  Update H病人结帐记录 Set 病人id = v_保留id Where 病人id = v_合并id;

  --床位状况记录
  Update 床位状况记录 Set 病人id = v_保留id Where 病人id = v_合并id;

  --病人担保记录
  Update 病人担保记录 Set 病人id = v_保留id Where 病人id = v_合并id;
  --特殊病人
  Select Count(*) Into v_Count From 特殊病人 Where 病人id = v_保留id;
  If v_Count = 0 Then
    Update 特殊病人 Set 病人id = v_保留id Where 病人id = v_合并id;
  Else
    Delete From 特殊病人 Where 病人id = v_合并id;
  End If;

  --病人未结费用
  For r_Owe In c_Owe(v_合并id) Loop
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(r_Owe.金额, 0)
    Where 主页id Is Null And 病人id = v_保留id And Nvl(病人病区id, 0) = Nvl(r_Owe.病人病区id, 0) And
          Nvl(病人科室id, 0) = Nvl(r_Owe.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Owe.开单部门id, 0) And
          Nvl(执行部门id, 0) = Nvl(r_Owe.执行部门id, 0) And Nvl(收入项目id, 0) = Nvl(r_Owe.收入项目id, 0) And
          Nvl(来源途径, 0) = Nvl(r_Owe.来源途径, 0);
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (v_保留id, Null, r_Owe.病人病区id, r_Owe.病人科室id, r_Owe.开单部门id, r_Owe.执行部门id, r_Owe.收入项目id, r_Owe.来源途径, r_Owe.金额);
    End If;
  End Loop;
  Delete From 病人未结费用 Where 病人id = v_合并id;
  Delete From 病人未结费用 Where 病人id = v_保留id And Nvl(金额, 0) = 0;

  --病人余额
  For r_Spare In c_Spare(v_合并id) Loop
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Spare.预交余额, 0), 费用余额 = Nvl(费用余额, 0) + Nvl(r_Spare.费用余额, 0)
    Where Nvl(性质, 0) = Nvl(r_Spare.性质, 0) And 病人id = v_保留id And 类型 = Nvl(r_Spare.类型, 2);
    If Sql%RowCount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 类型, 预交余额, 费用余额)
      Values
        (v_保留id, r_Spare.性质, Nvl(r_Spare.类型, 2), r_Spare.预交余额, r_Spare.费用余额);
    End If;
  End Loop;
  Delete From 病人余额 Where 病人id = v_合并id;
  Delete From 病人余额 Where 病人id = v_保留id And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0 And 性质 = 1;

  --病人过敏药物
  Insert Into 病人过敏药物
    (病人id, 过敏药物id, 过敏药物)
    Select v_保留id, 过敏药物id, 过敏药物
    From 病人过敏药物
    Where 病人id = v_合并id And 过敏药物id Not In (Select 过敏药物id From 病人过敏药物 Where 病人id = v_保留id);
  Delete From 病人过敏药物 Where 病人id = v_合并id;

  --病人社区信息
  Insert Into 病人社区信息
    (病人id, 社区, 社区号, 标志, 就诊类型, 就诊时间)
    Select v_保留id, 社区, 社区号, 标志, 就诊类型, 就诊时间
    From 病人社区信息
    Where 病人id = v_合并id And 社区 Not In (Select 社区 From 病人社区信息 Where 病人id = v_保留id);
  Delete From 病人社区信息 Where 病人id = v_合并id;

  --病人免疫记录
  Insert Into 病人免疫记录
    (病人id, 接种时间, 接种名称)
    Select v_保留id, a.接种时间, a.接种名称
    From 病人免疫记录 A
    Where a.病人id = v_合并id And Not Exists (Select 1 From 病人免疫记录 Where 病人id = v_保留id And 接种时间 = a.接种时间);
  Delete From 病人免疫记录 Where 病人id = v_合并id;

  --病人信息从表
  Insert Into 病人信息从表
    (病人id, 信息名, 信息值, 就诊id)
    Select v_保留id, a.信息名, a.信息值, a.就诊id
    From 病人信息从表 A
    Where a.病人id = v_合并id And Not Exists (Select 1
           From 病人信息从表
           Where 病人id = v_保留id And 信息名 = a.信息名 And Nvl(就诊id, 0) = Nvl(a.就诊id, 0));
  Delete From 病人信息从表 Where 病人id = v_合并id;

  --病人医疗卡属性
  Insert Into 病人医疗卡属性
    (病人id, 卡类别id, 卡号, 信息名, 信息值)
    Select v_保留id, a.卡类别id, a.卡号, a.信息名, a.信息值
    From 病人医疗卡属性 A
    Where a.病人id = v_合并id And Not Exists (Select 1
           From 病人医疗卡属性
           Where 病人id = v_保留id And 卡类别id = a.卡类别id And 卡号 = a.卡号 And 信息名 = a.信息名);
  Delete From 病人医疗卡属性 Where 病人id = v_合并id;

  --门诊病案记录
  Select Count(*) Into v_Count From 门诊病案记录 Where 病人id = v_保留id;
  If v_Count = 0 Then
    Select Count(*) Into v_Count From 门诊病案记录 Where 病人id = v_合并id;
    If v_Count > 0 Then
      Update 门诊病案记录 Set 病人id = v_保留id Where 病人id = v_合并id;
    End If;
  Else
    Delete From 门诊病案记录 Where 病人id = v_合并id;
  End If;

  --住院病案记录
  Select Count(*) Into v_Count From 住院病案记录 Where 病人id = v_保留id;

  If v_Count = 0 Then
    Select Count(*) Into v_Count From 住院病案记录 Where 病人id = v_合并id;
    If v_Count > 0 Then
      Update 住院病案记录 Set 病人id = v_保留id Where 病人id = v_合并id;
    End If;
  Else
    Begin
      v_Sql := 'Delete From 借阅记录 Where 病人ID=:1';
      Execute Immediate v_Sql
        Using v_合并id;
    Exception
      When Others Then
        Null;
    End;
  
    Delete From 住院病案记录 Where 病人id = v_合并id;
  End If;

  --医保病人相关处理
  --即使合病或保留的病人当前不是医保帐户,只要曾是医保帐户,险类不同也不能合并
  Select Count(Distinct 险类) Into v_Count From 医保病人关联表 Where 病人id In (v_合并id, v_保留id);
  If v_Count = 2 Then
    Close c_Infoa;
    Close c_Infob;
    v_Error := '两个病人分别属于不同的保险类别，不允许合并。';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From 医保病人关联表 Where 病人id = v_合并id And 标志 = 0;
  --a.合并的病人以前是医保帐户,现在不是
  If v_Count > 0 Then
    Select Count(*) Into v_Count From 医保病人关联表 Where 病人id = v_保留id;
    --a.1保留的病人现在是医保帐户
    --a.2.1保留的病人现在不是医保帐户,以前是,与a.1相同处理
    If v_Count > 0 Then
      Delete From 帐户年度信息 Where 病人id = v_合并id;
    
      Select Count(Distinct 医保号) Into v_Count From 医保病人关联表 Where 病人id In (v_合并id, v_保留id);
      If v_Count <> 2 Then
        --两个病人医保号相同时,不用处理医保病人档案
        For r_Insure In c_Insure(v_合并id) Loop
          --被合并的病人可能关联了多个医保病人,改为关联到保留的病人上
          --问题27445 保留指定病人的门诊号、住院号、医保号
          If v_合并id = B病人id_In Then
            Update 医保病人关联表
            Set 医保号 =
                 (Select 医保号 From 医保病人关联表 Where 病人id = v_合并id), 标志 = 0
            Where 险类 = r_Insure.险类 And 中心 = r_Insure.中心 And 医保号 = r_Insure.医保号 And 病人id <> v_合并id;
          Else
            Update 医保病人关联表
            Set 医保号 =
                 (Select 医保号 From 医保病人关联表 Where 病人id = v_保留id), 标志 = 0
            Where 险类 = r_Insure.险类 And 中心 = r_Insure.中心 And 医保号 = r_Insure.医保号 And 病人id <> v_合并id;
          End If;
          --合并的病人现在不是医保,即使是用户指定要保留该病人,也不保留它的帐户信息
          Delete From 医保病人档案 Where 险类 = r_Insure.险类 And 医保号 = r_Insure.医保号;
        End Loop;
      End If;
      Delete From 医保病人关联表 Where 病人id = v_合并id;
    Else
      --a.2.2保留的病人现在和以前都不是医保帐户
      Update 帐户年度信息 Set 病人id = v_保留id Where 病人id = v_合并id;
      Update 医保病人关联表 Set 病人id = v_保留id Where 病人id = v_合并id;
      --医保病人档案表不用处理,因为通过医保号关联<医保病人关联表>
    End If;
  Else
    Select Count(*) Into v_Count From 医保病人关联表 Where 病人id = v_合并id And 标志 = 1;
    --b.合并的病人现在是医保帐户
    If v_Count > 0 Then
      Select Count(*) Into v_Count From 医保病人关联表 Where 病人id = v_保留id;
      --b.1保留的病人现在也是医保帐户
      --b.2.1保留的病人现在不是医保帐户,以前是,与b.1相同处理
      If v_Count > 0 Then
        For r_Insure In c_Insure(v_合并id) Loop
          --转移帐户年度信息
          For r_Year In c_Year(v_合并id, r_Insure.险类) Loop
            Update 帐户年度信息
            Set 帐户增加累计 = Nvl(帐户增加累计, 0) + Nvl(r_Year.帐户增加累计, 0), 帐户支出累计 = Nvl(帐户支出累计, 0) + Nvl(r_Year.帐户支出累计, 0),
                进入统筹累计 = Nvl(进入统筹累计, 0) + Nvl(r_Year.进入统筹累计, 0), 统筹报销累计 = Nvl(统筹报销累计, 0) + Nvl(r_Year.统筹报销累计, 0),
                住院次数累计 = Nvl(住院次数累计, 0) + Nvl(r_Year.住院次数累计, 0), 大额统筹累计 = Nvl(大额统筹累计, 0) + Nvl(r_Year.大额统筹累计, 0),
                起付线累计 = Nvl(起付线累计, 0) + Nvl(r_Year.起付线累计, 0), 本次起付线 = Nvl(本次起付线, r_Year.本次起付线),
                基本统筹限额 = Nvl(基本统筹限额, r_Year.基本统筹限额), 大额统筹限额 = Nvl(大额统筹限额, r_Year.大额统筹限额), 封销信息 = Nvl(封销信息, r_Year.封销信息)
            Where 病人id = v_保留id And 险类 = r_Insure.险类 And 年度 = r_Year.年度;
            If Sql%RowCount = 0 Then
              Insert Into 帐户年度信息
                (病人id, 险类, 年度, 帐户增加累计, 帐户支出累计, 进入统筹累计, 统筹报销累计, 住院次数累计, 本次起付线, 基本统筹限额, 大额统筹限额, 起付线累计, 大额统筹累计, 封销信息)
              Values
                (v_保留id, r_Insure.险类, r_Year.年度, r_Year.帐户增加累计, r_Year.帐户支出累计, r_Year.进入统筹累计, r_Year.统筹报销累计,
                 r_Year.住院次数累计, r_Year.本次起付线, r_Year.基本统筹限额, r_Year.大额统筹限额, r_Year.起付线累计, r_Year.大额统筹累计, r_Year.封销信息);
            End If;
          End Loop;
          Delete From 帐户年度信息 Where 病人id = v_合并id;
        
          Select Count(Distinct 医保号) Into v_Count From 医保病人关联表 Where 病人id In (v_合并id, v_保留id);
          If v_Count <> 2 Then
            --两个病人医保号相同时,不用处理医保病人档案
            If v_合并id = B病人id_In Then
              Update 医保病人关联表
              Set 标志 = 0
              Where (险类, 中心, 医保号) In (Select 险类, 中心, 医保号 From 医保病人关联表 Where 病人id = v_保留id);
              Update 医保病人关联表 Set 标志 = 1 Where 病人id = v_保留id;
            End If;
            Delete From 医保病人关联表 Where 病人id = v_合并id;
          Else
            --被合并的病人可能关联了多个医保病人,改为关联到保留的病人上
            --问题27445 保留指定病人的门诊号、住院号、医保号
            If v_合并id = B病人id_In Then
              Update 医保病人关联表
              Set 医保号 =
                   (Select 医保号 From 医保病人关联表 Where 病人id = v_合并id), 标志 = 0
              Where 险类 = r_Insure.险类 And 中心 = r_Insure.中心 And 医保号 = r_Insure.医保号 And 病人id <> v_合并id;
            Else
              Update 医保病人关联表
              Set 医保号 =
                   (Select 医保号 From 医保病人关联表 Where 病人id = v_保留id), 标志 = 0
              Where 险类 = r_Insure.险类 And 中心 = r_Insure.中心 And 医保号 = r_Insure.医保号 And 病人id <> v_合并id;
            End If;
            --暂存用户指定要保留病人的帐户信息
            If v_合并id = B病人id_In Then
              Open c_Keepinsure(B病人id_In, r_Insure.险类);
              Fetch c_Keepinsure
                Into r_Keepinsure;
            End If;
          
            Delete From 医保病人关联表 Where 病人id = v_合并id;
            Delete From 医保病人档案 Where 险类 = r_Insure.险类 And 医保号 = r_Insure.医保号;
          
            --保留用户指定要保留病人的帐户信息
            If v_合并id = B病人id_In Then
              If c_Keepinsure%RowCount > 0 Then
                Update 医保病人档案
                Set 卡号 = r_Keepinsure.卡号, 医保号 = r_Keepinsure.医保号, 密码 = r_Keepinsure.密码, 人员身份 = r_Keepinsure.人员身份,
                    单位编码 = r_Keepinsure.单位编码, 顺序号 = r_Keepinsure.顺序号, 退休证号 = r_Keepinsure.退休证号, 帐户余额 = r_Keepinsure.帐户余额,
                    当前状态 = r_Keepinsure.当前状态, 病种id = r_Keepinsure.病种id, 在职 = r_Keepinsure.在职, 年龄段 = r_Keepinsure.年龄段,
                    灰度级 = r_Keepinsure.灰度级, 就诊时间 = r_Keepinsure.就诊时间
                Where (险类, 中心, 医保号) In (Select 险类, 中心, 医保号 From 医保病人关联表 Where 病人id = v_保留id);
                --保留病人可能关联了多个医保病人,都要更改医保号
                Update 医保病人关联表
                Set 医保号 = r_Keepinsure.医保号, 标志 = 0
                Where (险类, 中心, 医保号) In (Select 险类, 中心, 医保号 From 医保病人关联表 Where 病人id = v_保留id);
                Update 医保病人关联表 Set 标志 = 1 Where 病人id = v_保留id;
              End If;
              Close c_Keepinsure;
            End If;
          End If;
        End Loop;
      Else
        --b.2.2保留的病人现在和以前都不是医保帐户
        Update 帐户年度信息 Set 病人id = v_保留id Where 病人id = v_合并id;
        Update 医保病人关联表 Set 病人id = v_保留id Where 病人id = v_合并id;
        --医保病人档案表不用处理,因为通过医保号关联<医保病人关联表>
      End If;
    Else
      --c.合并的病人以前和现在都不是医保帐户,不作任何处理
      Null;
    End If;
  End If;

  --处理体检子系统的病人合并
  n_Have := 0;
  Begin
    Select 1 Into n_Have From zlSystems Where Floor(编号 / 100) = 21;
  Exception
    When Others Then
      Null;
  End;
  If n_Have = 1 Then
    v_Sql := 'Begin zl21_病人信息_Merge(:1,:2); End;';
    Execute Immediate v_Sql
      Using v_合并id, v_保留id;
  End If;

  --其它病人,病案主页相关表
  For v_Loop In 1 .. Arronpage.Count Loop
    --Executesql('Update ' || Arronpage(v_Loop) || ' Set 病人ID=' || v_保留id || ' Where 病人ID=' || v_合并id || ' And Nvl(主页ID,0) = 0');
    --"主页=0，主页ID is NULL，主页ID=挂号ID"都有可能，前面部分与主页ID关联都没处理到，因此不加条件
    v_Sql := 'Update ' || Arronpage(v_Loop) || ' Set 病人ID=:1 Where 病人ID=:2';
    Execute Immediate v_Sql
      Using v_保留id, v_合并id;
  End Loop;
  For v_Loop In 1 .. Arronbase.Count Loop
    v_Sql := 'Update ' || Arronbase(v_Loop) || ' Set 病人ID=:1 Where 病人ID=:2';
    Execute Immediate v_Sql
      Using v_保留id, v_合并id;
  End Loop;

  --删除实际不保留的病人信息
  Delete From 病人信息 Where 病人id = v_合并id;

  --根据界面选择保留病人信息
  Update 病人信息
  Set 姓名 = Nvl(r_Infob.姓名, r_Infoa.姓名), 性别 = Nvl(r_Infob.性别, r_Infoa.性别), 年龄 = Nvl(r_Infob.年龄, r_Infoa.年龄), 门诊号 = v_门诊号,
      住院号 = v_住院号, 就诊卡号 = Nvl(r_Infob.就诊卡号, r_Infoa.就诊卡号), 卡验证码 = Decode(r_Infob.就诊卡号, Null, r_Infoa.卡验证码, r_Infob.卡验证码),
      费别 = Nvl(r_Infob.费别, r_Infoa.费别), 医疗付款方式 = Nvl(r_Infob.医疗付款方式, r_Infoa.医疗付款方式),
      出生日期 = Nvl(r_Infob.出生日期, r_Infoa.出生日期), 出生地点 = Nvl(r_Infob.出生地点, r_Infoa.出生地点),
      身份证号 = Nvl(r_Infob.身份证号, r_Infoa.身份证号), 身份 = Nvl(r_Infob.身份, r_Infoa.身份), 职业 = Nvl(r_Infob.职业, r_Infoa.职业),
      民族 = Nvl(r_Infob.民族, r_Infoa.民族), 国籍 = Nvl(r_Infob.国籍, r_Infoa.国籍), 学历 = Nvl(r_Infob.学历, r_Infoa.学历),
      籍贯 = Nvl(r_Infob.籍贯, r_Infoa.籍贯), 区域 = Nvl(r_Infob.区域, r_Infoa.区域), 婚姻状况 = Nvl(r_Infob.婚姻状况, r_Infoa.婚姻状况),
      家庭地址 = Nvl(r_Infob.家庭地址, r_Infoa.家庭地址), 家庭电话 = Nvl(r_Infob.家庭电话, r_Infoa.家庭电话),
      家庭地址邮编 = Nvl(r_Infob.家庭地址邮编, r_Infoa.家庭地址邮编), 户口地址 = Nvl(r_Infob.户口地址, r_Infoa.户口地址),
      户口地址邮编 = Nvl(r_Infob.户口地址邮编, r_Infoa.户口地址邮编), 联系人姓名 = Nvl(r_Infob.联系人姓名, r_Infoa.联系人姓名),
      联系人关系 = Nvl(r_Infob.联系人关系, r_Infoa.联系人关系), 联系人地址 = Nvl(r_Infob.联系人地址, r_Infoa.联系人地址),
      联系人电话 = Nvl(r_Infob.联系人电话, r_Infoa.联系人电话), 合同单位id = Nvl(r_Infob.合同单位id, r_Infoa.合同单位id),
      工作单位 = Nvl(r_Infob.工作单位, r_Infoa.工作单位), 单位电话 = Nvl(r_Infob.单位电话, r_Infoa.单位电话),
      单位邮编 = Nvl(r_Infob.单位邮编, r_Infoa.单位邮编), 单位开户行 = Nvl(r_Infob.单位开户行, r_Infoa.单位开户行),
      单位帐号 = Nvl(r_Infob.单位帐号, r_Infoa.单位帐号), 就诊时间 = Nvl(r_Infob.就诊时间, r_Infoa.就诊时间),
      就诊状态 = Nvl(r_Infob.就诊状态, r_Infoa.就诊状态), 就诊诊室 = Nvl(r_Infob.就诊诊室, r_Infoa.就诊诊室), 险类 = Nvl(r_Infob.险类, r_Infoa.险类),
      登记时间 = Nvl(r_Infob.登记时间, r_Infoa.登记时间), 住院次数 = Null, 主页id = Null, 当前床号 = Null, 当前科室id = Null, 当前病区id = Null,
      入院时间 = Null, 出院时间 = Null, 在院 = Decode(Nvl(r_Infob.在院, 0), 1, 1, Null)
  Where 病人id = v_保留id;

  Open c_Info(v_保留id);
  Fetch c_Info
    Into r_Info;
  If c_Info%RowCount > 0 Then
    --最后一次为预约病人,只需要更改住院次数和入院时间
    If r_Info.主页id = 0 Then
      Update 病人信息
      Set 主页id = Decode(r_Info.最大主页id, 0, Null, r_Info.最大主页id), 住院次数 = Decode(n_Max住院次数, 0, Null, n_Max住院次数)
      Where 病人id = v_保留id;
    Else
      Update 病人信息
      Set 主页id = Decode(r_Info.最大主页id, 0, Null, r_Info.最大主页id), 住院次数 = Decode(n_Max住院次数, 0, Null, n_Max住院次数),
          当前床号 = Decode(r_Info.出院日期, Null, r_Info.出院病床, Null), 当前病区id = Decode(r_Info.出院日期, Null, r_Info.当前病区id, Null),
          当前科室id = Decode(r_Info.出院日期, Null, r_Info.出院科室id, Null), 入院时间 = r_Info.入院日期, 出院时间 = r_Info.出院日期
      
      Where 病人id = v_保留id;
    End If;
    --处理担保信息
    Select Nvl(主页id, -1) Into n_主页id From 病人信息 Where 病人id = v_保留id;
    --提取当前有效的正常担保记录,确保正常担保与临时担保不并存
    Select Nvl(Sum(担保额), 0), Count(病人id)
    Into n_担保额, n_Row
    From 病人担保记录
    Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And (到期时间 Is Null Or 到期时间 > Sysdate) And 担保性质 = 0 And 删除标志 = 1;
    If n_Row = 0 Then
      --保留最后一条临时担保记录,其余到期
      Update 病人担保记录
      Set 到期时间 = Sysdate - 1 / 24 / 60 / 60
      Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And 担保性质 = 1 And (到期时间 Is Null Or 到期时间 > Sysdate) And 删除标志 = 1 And
            登记时间 <> (Select Max(登记时间)
                     From 病人担保记录
                     Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And 担保性质 = 1 And (到期时间 Is Null Or 到期时间 > Sysdate) And
                           删除标志 = 1);
    Else
      --有正常担保就让临时担保失效
      Update 病人担保记录
      Set 到期时间 = Sysdate - 1 / 24 / 60 / 60
      Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And 担保性质 = 1 And (到期时间 Is Null Or 到期时间 > Sysdate) And 删除标志 = 1;
    End If;
  
    --提取当前有效担保额及有效担保记录数
    n_Row    := 0;
    n_担保额 := 0;
    v_担保人 := '';
    For r_提保信息 In (Select 担保人, 担保额
                   From 病人担保记录
                   Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And (到期时间 Is Null Or 到期时间 > Sysdate) And 删除标志 = 1) Loop
      n_Row     := n_Row + 1;
      n_担保额  := n_担保额 + r_提保信息.担保额;
      v_担保人  := v_担保人 || ',' || r_提保信息.担保人;
      n_Lengthb := Lengthb(v_担保人);
      If n_Lengthb >= 101 Then
        v_Error := '不能合并担保记录，在病人信息保存时超过担保人字段长度！';
        Raise Err_Custom;
      End If;
    End Loop;
    v_担保人 := Substr(v_担保人, 2, 100);
  
    If n_Row = 0 Then
      Update 病人信息 Set 担保人 = Null, 担保额 = Null, 担保性质 = Null Where 病人id = v_保留id;
    Else
      --提取最后一条有效担保人和担保性质
      Select 担保性质
      Into n_担保性质
      From 病人担保记录
      Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And 删除标志 = 1 And
            登记时间 =
            (Select Max(登记时间)
             From 病人担保记录
             Where 病人id = v_保留id And Nvl(主页id, -1) = n_主页id And (到期时间 Is Null Or 到期时间 > Sysdate) And 删除标志 = 1);
    
      Update 病人信息 Set 担保人 = v_担保人, 担保额 = n_担保额, 担保性质 = n_担保性质 Where 病人id = v_保留id;
    End If;
  End If;

  Close c_Info;
  Close c_Infoa;
  Close c_Infob;

  --对病人进行解锁
  Update 病人信息 Set 锁定 = 0 Where 病人id In (A病人id_In, B病人id_In);
Exception
  When Err_Custom Then
    Begin
      Rollback; --不然会死锁
      Zl_病人信息_锁定(A病人id_In, 0);
      Zl_病人信息_锁定(B病人id_In, 0);
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    End;
  When Others Then
    Begin
      Rollback; --不然会死锁
      Zl_病人信息_锁定(A病人id_In, 0);
      Zl_病人信息_锁定(B病人id_In, 0);
      zl_ErrorCenter(SQLCode, SQLErrM);
    End;
End Zl_病人信息_Merge;
/


--90658:梁经伙,2015-11-20,当前传入的医嘱是中服法医嘱或者中药煎法医嘱过程内部单独处理
Create Or Replace Procedure Zl_病人医嘱执行_Finish
(
  医嘱id_In       病人医嘱执行.医嘱id%Type,
  发送号_In       病人医嘱执行.发送号%Type,
  阳性_In         病人医嘱发送.结果阳性%Type := Null,
  单独执行_In     Number := 0,
  操作员编号_In   人员表.编号%Type := Null,
  操作员姓名_In   人员表.姓名%Type := Null,
  执行部门id_In   门诊费用记录.执行部门id%Type := 0,
  检验项目记帐_In Number := 0
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --      单独执行_In=检验或检查医嘱组合是否采用对每个项目分散单独执行的方式
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门
  --检验项目记帐_In=如果是检验项目时，需要记帐但不完成医嘱发送状态
) Is
  Cursor c_Advice Is
    Select a.Id, a.相关id, Nvl(a.相关id, a.Id) As 组id, a.诊疗类别, a.病人来源, a.标本部位, a.开始执行时间, a.病人id, a.主页id, a.执行科室id, b.操作类型
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.Id = 医嘱id_In And a.诊疗项目id = b.Id;
  r_Advice c_Advice%RowType;

  v_Date     Date;
  v_开始时间 Date;
  v_诊疗类别 诊疗项目目录.类别%Type;
  v_操作类型 诊疗项目目录.操作类型%Type;
  v_Temp     Varchar2(255);
  v_费用性质 病人医嘱发送.记录性质%Type;
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  n_Cnt      Number;

Begin

  --如果启用了参数：输血和皮试医嘱执行后需要核对，则输血和皮试医嘱不自动完成
  Select b.类别, b.操作类型
  Into v_诊疗类别, v_操作类型
  From 病人医嘱记录 A, 诊疗项目目录 B
  Where a.诊疗项目id = b.Id And a.Id = 医嘱id_In;

  v_Temp := zl_GetSysParameter(186);
  If v_Temp = '11' Then
    Select Count(1)
    Into n_Cnt
    From 病人医嘱执行
    Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 核对人 Is Null And (v_诊疗类别 = 'E' And v_操作类型 In ('1', '8') Or v_诊疗类别 = 'K');
  Elsif v_Temp = '01' Then
    Select Count(1)
    Into n_Cnt
    From 病人医嘱执行
    Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 核对人 Is Null And v_诊疗类别 = 'E' And v_操作类型 = '1';
  Elsif v_Temp = '10' Then
    Select Count(1)
    Into n_Cnt
    From 病人医嘱执行
    Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 核对人 Is Null And (v_诊疗类别 = 'E' And v_操作类型 = '8' Or v_诊疗类别 = 'K');
  End If;

  If n_Cnt > 0 Then
    Return;
  End If;

  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员编号 := 操作员编号_In;
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;
  Select Sysdate Into v_Date From Dual;

  --执行状态
  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;
  Close c_Advice;

  If 检验项目记帐_In = 0 Then
    If (r_Advice.诊疗类别 = 'C' And r_Advice.相关id Is Not Null) Or r_Advice.诊疗类别 = 'D' Then
      If Nvl(单独执行_In, 0) = 1 Then
        --单个检验或检查项目
        Update 病人医嘱发送
        Set 执行状态 = 1, 完成人 = v_人员姓名, 完成时间 = v_Date, 结果阳性 = Decode(阳性_In, Null, 结果阳性, 阳性_In)
        Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In;
      Else
        --一并的检验项目或多部位的检查项目
        If r_Advice.诊疗类别 = 'D' Then
          Update 病人医嘱发送
          Set 执行状态 = 1, 完成人 = v_人员姓名, 完成时间 = v_Date, 结果阳性 = Decode(阳性_In, Null, 结果阳性, 阳性_In)
          Where 发送号 + 0 = 发送号_In And
                医嘱id In (Select ID From 病人医嘱记录 Where (ID = r_Advice.组id Or 相关id = r_Advice.组id));
        Else
          Update 病人医嘱发送
          Set 执行状态 = 1, 完成人 = v_人员姓名, 完成时间 = v_Date, 结果阳性 = Decode(阳性_In, Null, 结果阳性, 阳性_In)
          Where 发送号 + 0 = 发送号_In And 医嘱id In (Select ID From 病人医嘱记录 Where 相关id = r_Advice.相关id);
        End If;
      End If;
    Else
      --包含附加手术,检查部位,以及其它独立医嘱;麻醉和中药煎法是单独安排 
      If r_Advice.诊疗类别 = 'E' And (r_Advice.操作类型 = '4' Or r_Advice.操作类型 = '3') Then
        Update 病人医嘱发送
        Set 执行状态 = 1, 完成人 = v_人员姓名, 完成时间 = v_Date, 结果阳性 = Decode(阳性_In, Null, 结果阳性, 阳性_In)
        Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In;
      Else
        Update 病人医嘱发送
        Set 执行状态 = 1, 完成人 = v_人员姓名, 完成时间 = v_Date, 结果阳性 = Decode(阳性_In, Null, 结果阳性, 阳性_In)
        Where 发送号 + 0 = 发送号_In And
              医嘱id In (Select ID
                       From 病人医嘱记录
                       Where (ID = r_Advice.组id Or 相关id = r_Advice.组id) And 诊疗类别 = r_Advice.诊疗类别);
      End If;
    End If;
  End If;
  If r_Advice.病人来源 = 2 Then
    Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
    Into v_费用性质
    From 病人医嘱发送
    Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In;
  Else
    v_费用性质 := 1;
  End If;
  --检验自动完成采集
  If v_诊疗类别 = 'E' And v_操作类型 = '6' Then
    Update 病人医嘱发送 A
    Set a.采样人 = Nvl(a.采样人, v_人员姓名), a.采样时间 = Nvl(a.采样时间, v_Date)
    Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = r_Advice.组id Or 相关id = r_Advice.组id)) And 发送号 = 发送号_In;
  End If;

  If v_费用性质 = 1 Then
    Zl_门诊医嘱执行_Finish(医嘱id_In, 发送号_In, 单独执行_In, v_人员编号, v_人员姓名, r_Advice.组id, r_Advice.诊疗类别, 执行部门id_In);
  Else
    Zl_住院医嘱执行_Finish(医嘱id_In, 发送号_In, 单独执行_In, v_人员编号, v_人员姓名, r_Advice.组id, r_Advice.诊疗类别, 执行部门id_In);
  End If;

  If r_Advice.诊疗类别 = 'F' Then
    If Not r_Advice.标本部位 Is Null Then
      v_开始时间 := To_Date(r_Advice.标本部位, 'yyyy-mm-dd hh24:mi:ss');
    Else
      v_开始时间 := r_Advice.开始执行时间;
    End If;
    Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '手术', r_Advice.执行科室id, v_人员姓名, v_开始时间, v_Date);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱执行_Finish;
/


--90658:梁经伙,2015-11-20,当前传入的医嘱是中服法医嘱或者中药煎法医嘱过程内部单独处理
Create Or Replace Procedure Zl_病人医嘱执行_Cancel
(
  医嘱id_In     病人医嘱执行.医嘱id%Type,
  发送号_In     病人医嘱执行.发送号%Type,
  取消皮试_In   Number := Null,
  单独执行_In   Number := 0,
  执行部门id_In 门诊费用记录.执行部门id%Type := 0,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --      单独执行_In=检验或检查医嘱组合是否采用对每个项目分散单独执行的方式
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门
) Is
  Cursor C_Advice Is
    Select A.Id, A.相关id, Nvl(A.相关id, A.Id) As 组id, A.病人id, Decode(A.病人来源, 1, C.Id, A.主页id) As 就诊id, A.诊疗类别, B.操作类型,
           A.病人来源
    From 病人医嘱记录 A, 诊疗项目目录 B, 病人挂号记录 C
    Where A.诊疗项目id = B.Id And A.Id = 医嘱id_In And A.挂号单 = C.No(+);
  R_Advice C_Advice%RowType;

  V_Temp     Varchar2(255);
  V_人员编号 人员表.编号%Type;
  V_人员姓名 人员表.姓名%Type;
  V_费用性质 病人医嘱发送.记录性质%Type;
  V_操作人员 人员表.姓名%Type;
  D_完成时间 病人医嘱发送.完成时间%Type;
  N_取消执行 Number;
  N_Diffday  Number(18, 3);
  V_Date     Date;
  V_Count    Number;
  Err_Custom Exception;
  V_Error Varchar2(2000);
Begin
  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    V_人员编号 := 操作员编号_In;
    V_人员姓名 := 操作员姓名_In;
  Else
    V_Temp     := Zl_Identity;
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1);
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1);
    V_人员编号 := Substr(V_Temp, 1, Instr(V_Temp, ',') - 1);
    V_人员姓名 := Substr(V_Temp, Instr(V_Temp, ',') + 1);
  End If;

  Open C_Advice;
  Fetch C_Advice
    Into R_Advice;
  Close C_Advice;
  --医嘱取消执行天数限制，会诊医嘱不做限制
  If Not (R_Advice.诊疗类别 = 'Z' And R_Advice.操作类型 = 7) Then
    --父医嘱不需要发送，没有发送记录，此时取子医嘱的完成时间
    If Nvl(单独执行_In, 0) <> 1 Then
      Select Count(1) Into V_Count From 病人医嘱发送 Where 发送号 + 0 = 发送号_In And 医嘱id = R_Advice.组id;
    End If;
  
    If Nvl(单独执行_In, 0) = 1 Or V_Count = 0 Then
      Select 完成时间 Into D_完成时间 From 病人医嘱发送 Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In;
    Else
      Select 完成时间 Into D_完成时间 From 病人医嘱发送 Where 发送号 + 0 = 发送号_In And 医嘱id = R_Advice.组id;
    End If;
    If Not D_完成时间 Is Null Then
      Select Zl_To_Number(Nvl(zl_GetSysParameter(220), '999')) Into N_取消执行 From Dual;
      Select Sysdate - D_完成时间 Into N_Diffday From Dual;
      --完成时间超过取消执行天数的记录，不允许取消执行
      If N_Diffday > N_取消执行 Then
        V_Error := '医嘱执行完成时间超过了取消执行有效天数，不能取消执行！';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  If (R_Advice.诊疗类别 = 'C' And R_Advice.相关id Is Not Null) Or R_Advice.诊疗类别 = 'D' Then
    If Nvl(单独执行_In, 0) = 1 Then
      Select Count(*)
      Into V_Count
      From 病人医嘱执行
      Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In And Nvl(执行结果, 1) <> 0;
      Update 病人医嘱发送
      Set 执行状态 = Decode(V_Count, 0, 0, 3), 完成人 = Null, 完成时间 = Null
      Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In;
    Else
      Select Count(*)
      Into V_Count
      From 病人医嘱执行
      Where 发送号 + 0 = 发送号_In And 医嘱id In (Select ID From 病人医嘱记录 Where 相关id = R_Advice.相关id) And Nvl(执行结果, 1) <> 0;
    
      If R_Advice.诊疗类别 = 'D' Then
        Update 病人医嘱发送
        Set 执行状态 = Decode(V_Count, 0, 0, 3), 完成人 = Null, 完成时间 = Null
        Where 发送号 + 0 = 发送号_In And
              医嘱id In (Select ID From 病人医嘱记录 Where (ID = R_Advice.组id Or 相关id = R_Advice.组id));
      Else
        Update 病人医嘱发送
        Set 执行状态 = Decode(V_Count, 0, 0, 3), 完成人 = Null, 完成时间 = Null
        Where 发送号 + 0 = 发送号_In And 医嘱id In (Select ID From 病人医嘱记录 Where 相关id = R_Advice.相关id);
      End If;
    End If;
  Else
    Select Count(*) Into V_Count From 病人医嘱执行 Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;
  
   --包含附加手术,检验部位,以及其它独立医嘱;麻醉和中药煎法是单独安排
    If r_Advice.诊疗类别 = 'E' And (r_Advice.操作类型 = '4' Or r_Advice.操作类型 = '3') Then
      Update 病人医嘱发送
      Set 执行状态 = Decode(v_Count, 0, 0, 3), 完成人 = Null, 完成时间 = Null
      Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In;
    Else
      Update 病人医嘱发送
      Set 执行状态 = Decode(v_Count, 0, 0, 3), 完成人 = Null, 完成时间 = Null
      Where 发送号 + 0 = 发送号_In And
            医嘱id In (Select ID
                     From 病人医嘱记录
                     Where (ID = r_Advice.组id Or 相关id = r_Advice.组id) And 诊疗类别 = r_Advice.诊疗类别);
    End If;
  End If;
  --更新对应的费用执行状态为未执行，在Zl_门诊医嘱执行_Cancel和Zl_住院医嘱执行_Cancel中进行

  --删除过敏登记记录(当前人员登记的)
  If R_Advice.诊疗类别 = 'E' And R_Advice.操作类型 = '1' Then
    Select Max(操作时间), Max(操作人员)
    Into V_Date, V_操作人员
    From (Select 操作时间, 操作人员 From 病人医嘱状态 Where 医嘱id = 医嘱id_In And 操作类型 = 10 Order By 操作时间 Desc)
    Where Rownum < 2;
    If V_Date Is Not Null And (V_操作人员 = V_人员姓名 Or Nvl(取消皮试_In, 0) = 1) Then
      --可能因为未设置对应过敏药物而未填写过敏记录
      --删除状态，以便标记未用时回退发送记录
      Update 病人医嘱记录 Set 皮试结果 = Null Where ID = 医嘱id_In;
      For R_Date In (Select 操作时间 From 病人医嘱状态 Where 医嘱id = 医嘱id_In And 操作类型 = 10) Loop
        Delete From 病人过敏记录
        Where 病人id = R_Advice.病人id And 记录来源 = 2 And Nvl(主页id, 0) = Nvl(R_Advice.就诊id, 0) And 记录时间 = R_Date.操作时间;
        Delete 病人医嘱状态 Where 医嘱id = 医嘱id_In And 操作类型 = 10 And 操作时间 = R_Date.操作时间;
      End Loop;
    End If;
  End If;
  --检验自动采集，没有采集方式的医嘱执行记录，则清空采样人与采样时间
  Select Count(*) Into V_Count From 病人医嘱执行 Where 发送号 = 发送号_In And 医嘱id = R_Advice.组id;
  If V_Count = 0 And R_Advice.诊疗类别 = 'E' And R_Advice.操作类型 = '6' Then
    Update 病人医嘱发送 A
    Set A.采样人 = Null, A.采样时间 = Null
    Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = R_Advice.组id Or 相关id = R_Advice.组id)) And 发送号 = 发送号_In;
  End If;

  If R_Advice.病人来源 = 2 Then
    Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
    Into V_费用性质
    From 病人医嘱发送
    Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In;
  Else
    V_费用性质 := 1;
  End If;

  If V_费用性质 = 1 Then
    Zl_门诊医嘱执行_Cancel(医嘱id_In, 发送号_In, 单独执行_In, V_人员编号, V_人员姓名, R_Advice.组id, R_Advice.诊疗类别, 执行部门id_In);
  Else
    Zl_住院医嘱执行_Cancel(医嘱id_In, 发送号_In, 单独执行_In, V_人员编号, V_人员姓名, R_Advice.组id, R_Advice.诊疗类别, 执行部门id_In);
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱执行_Cancel;
/


---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.34.50' Where 编号=&n_System;
--部件版本号
Commit;