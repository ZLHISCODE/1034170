--[连续升级]1
--[管理工具版本号]10.34.30
--本脚本支持从ZLHIS+ v10.34.50 升级到 v10.34.60
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--91780:刘尔旋,2015-12-28,相同发票号轧帐
Alter Table 人员收缴票据 Add 批次 Varchar2(20);

--91427:马政,2015-12-22,新增药品验收系统
create table 药品验收记录
(
id number(18),
NO varchar2(8),
库房id number(18),
供药单位id number(18),
验收人 varchar2(200),
验收日期 date,
复核人 varchar2(200),
复核日期 date,
是否合格 number(1),
备注  varchar2(1000)
) TABLESPACE zl9MedLst
    initrans 20;

create table 药品验收明细 
(
验收id number(18),
药品id number(18),
成本价 number(16,7),
零售价 number(16,7),
进药数量 number(16,5),
批号 varchar2(20),
生产日期 date,
效期 date,
产地 varchar2(60),
批准文号 varchar2(40),
进药日期 date,
是否合格 number(1)
) TABLESPACE zl9MedLst
    initrans 20;

Create Sequence 药品验收记录_ID Start With 1; 

Alter Table 药品验收记录 Add Constraint 药品验收记录_PK Primary Key (ID) Using Index Tablespace zl9indexhis;

Alter Table 药品验收明细 Add Constraint 药品验收明细_UQ_验收ID Unique (验收id,药品id) Using Index Tablespace zl9indexhis;

Alter Table 药品验收明细 Modify 验收id Constraint 药品验收明细_NN_验收id Not Null;   

Create Index 药品验收记录_IX_库房id On 药品验收记录(库房id) Tablespace zl9Indexhis;

Create Index 药品验收记录_IX_供药单位id On 药品验收记录(供药单位id) Tablespace zl9Indexhis;   

Create Index 药品验收记录_IX_NO On 药品验收记录(NO) Tablespace zl9Indexhis;  

Create Index 药品验收明细_IX_药品id On 药品验收明细(药品id) Tablespace zl9Indexhis;  

--91225:梁经伙,2015-12-16,传染病管理系统 基本数据
create table 传染病目录(
   编码 VARCHAR2(10),
   名称 VARCHAR2(200), 
   简码 VARCHAR2(200), 
   说明 VARCHAR2(500)
) TABLESPACE zl9EprDat;

create table 疾病阳性记录(
   ID    Number(18),
   病人ID number(18), 
   主页id NUMBER(5),
   挂号单 VARCHAR2(8),
   送检时间 date,
   送检科室ID number(18), 
   送检医生 VARCHAR2(201), 
   标本名称 VARCHAR2(60),
   反馈结果 VARCHAR2(1000),
   传染病名称 VARCHAR2(200),
   检查时间 date,
   登记时间 date,
   登记人 VARCHAR2(100),
   登记科室ID number(18), 
   记录状态 number(2),
   处理人 VARCHAR2(100),
   处理时间 date,
   处理情况说明 VARCHAR2(1000),
   文件ID number(18),
   待转出 Number(3)
) TABLESPACE zl9EprDat;

create table 疾病报告反馈(
   文件ID NUMBER(18),
   登记时间 date, 
   登记人 VARCHAR2(100),
   记录状态 NUMBER(3),
   反馈内容 VARCHAR2 (500),
   处理人 VARCHAR2(100),
   处理时间 date,
   处理情况说明 VARCHAR2(500),
   待转出 Number(3)
) TABLESPACE zl9EprDat;

alter table 疾病申报记录 Add(报卡类型 VARCHAR2(50),报告医生 VARCHAR2(100),撤档人 VARCHAR2(100),撤档时间 Date,病人ID NUMBER(18),主页ID NUMBER(18),病人来源 NUMBER(3));

Create Sequence 疾病阳性记录_ID Start With 1;

Alter Table 传染病目录 Add Constraint 传染病目录_PK Primary Key (编码) Using Index Tablespace zl9Indexcis;

Alter Table 传染病目录 Add Constraint 传染病目录_UQ_名称 Unique (名称) Using Index Tablespace zl9Indexhis;

Alter Table 疾病阳性记录 Add Constraint 疾病阳性记录_PK Primary Key (ID) Using Index Tablespace zl9Indexcis;

Alter Table 疾病报告反馈 Add Constraint 疾病报告反馈_PK Primary Key (文件ID,登记时间) Using Index Tablespace zl9Indexcis;

Create Index 疾病阳性记录_IX_病人ID On 疾病阳性记录(病人ID,主页ID)  Tablespace zl9Indexcis;

Create Index 疾病阳性记录_IX_登记时间 On 疾病阳性记录(登记时间)  Tablespace zl9Indexcis;

Create Index 疾病阳性记录_IX_挂号单 On 疾病阳性记录(挂号单)  Tablespace zl9Indexcis;

Create Index 疾病阳性记录_IX_待转出 On 疾病阳性记录(待转出) Tablespace zl9Indexcis;

Create Index 疾病阳性记录_IX_文件ID On 疾病阳性记录(文件ID) Tablespace zl9Indexcis;

Create Index 疾病申报记录_IX_姓名 On 疾病申报记录(姓名) Tablespace zl9Indexcis;

Create Index 疾病申报记录_IX_病人ID On 疾病申报记录(病人ID,主页ID)  Tablespace zl9Indexcis;

Create Index 疾病报告反馈_IX_待转出 On 疾病报告反馈(待转出) Tablespace zl9Indexcis;

Create Index 疾病报告反馈_IX_登记时间 On 疾病报告反馈(登记时间) Tablespace zl9Indexcis;

--91687:余智勇,2015-12-15,增加索引
Create Index 门诊穿刺台_Ix_待穿病人id On 门诊穿刺台(待穿病人id) Pctfree 5 Tablespace Zl9indexcis Nologging;

--90666:陈刘,2015-12-07,按江苏省病例规范要求,新增体温部位:耳温
Alter Table 体温记录项目 Modify 记录符 Varchar2(20);

--92493:梁经伙,2016-01-08,疾病报告前提增加字段 报告病种
Alter Table 疾病报告前提 Add 报告病种 Varchar2(80);

--91712:涂建华,2015-12-16,病理玻片信息外键索引修正
Create Index 病理玻片信息_IX_病理医嘱ID On 病理玻片信息(病理医嘱ID) Tablespace zl9Indexcis nologging;

--91225:梁经伙,2015-12-16,传染病管理系统 基本数据
Alter Table 疾病阳性记录 Add Constraint 疾病阳性记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);
Alter Table 疾病阳性记录 Add Constraint 疾病阳性记录_FK_主页ID Foreign Key (病人ID,主页ID) References 病案主页(病人ID,主页ID);
Alter Table 疾病阳性记录 Add Constraint 疾病阳性记录_FK_送检科室ID Foreign Key (送检科室ID) References 部门表(ID);
Alter Table 疾病阳性记录 Add Constraint 疾病阳性记录_FK_登记科室ID Foreign Key (登记科室ID) References 部门表(ID);
Alter Table 疾病报告反馈 Add Constraint 疾病报告反馈_FK_文件ID Foreign Key (文件ID) References 疾病申报记录 (文件ID) On Delete Cascade;
Alter Table 疾病申报记录 Add Constraint 疾病申报记录_FK_病人ID Foreign Key (病人ID) References 病人信息 (病人ID);

--91427:马政,2015-12-22,新增药品验收系统
Alter Table 药品验收记录 Add Constraint 药品验收记录_FK_库房id Foreign Key (库房id) References 部门表(ID) On Delete Cascade;
Alter Table 药品验收记录 Add Constraint 药品验收记录_FK_供药单位id Foreign Key (供药单位id) References 供应商(ID) On Delete Cascade;
Alter Table 药品验收明细 Add Constraint 药品验收明细_FK_验收id Foreign Key (验收id) References 药品验收记录(ID) On Delete Cascade;
Alter Table 药品验收明细 Add Constraint 药品验收明细_FK_药品id Foreign Key (药品id) References 收费项目目录(ID) On Delete Cascade;


-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--89717:余伟节,2016-01-14,出院后不允许取消完成路径
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1256, 0, 0, 0, 0, 13, '出院后不允许取消完成路径', '0', '0','如果启用此参数，出院的病人不允许取消完成的路径。'
  From Dual
  Where Not Exists (Select 1
         From zlParameters
         Where 参数名 = '出院后不允许取消完成路径' And Nvl(模块, 0) = 1256 And Nvl(系统, 0) = &n_System);

--92321:胡俊勇,2015-01-13,申请单启用环节
Insert Into Zlparameters
  (Id, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 238, '申请单启用环节', '111', '111',    
  '控制申请单启用环节，如果启用对应的申请单，则医嘱下达时只能通过申请单方式下达；如果在医嘱下达界面输入查找项目，查找后也会自动弹出申请单界面进行填写；'|| Chr(13) ||'第一位依次是：检查、检验、输血'
  From Dual Where Not Exists (Select 1
         From Zlparameters  Where 参数名 = '申请单启用环节' And Nvl(模块, 0) = 0 And Nvl(系统, 0) = &n_System); 

--89419:张德婷,2015-01-05,出院病人不收配置费
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 1, 0, 0, 24,'出院病人不收配置费', Null, '0', '已经出院的病人不收配置费'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '出院病人不收配置费');

--91671:胡俊勇,2015-08-30,主刀医师达到手术等级无需审核
Insert Into Zlparameters
  (Id, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 254, '主刀医师达到手术等级无需审核', '0', '0',        
  '启用参数后控制：下达手术医嘱时，如果主刀医师满足手术等级要求（手术授权管理未启动就是按医生手术等级来）；则无需审核，可直接校对；'
  || Chr(13) ||'如果主刀医师不满足(低于)手术项目等级时，才需要审核；' 
  From Dual Where Not Exists (Select 1
         From Zlparameters  Where 参数名 = '主刀医师达到手术等级无需审核' And Nvl(模块, 0) = 0 And Nvl(系统, 0) = &n_System);

--91665:冉俊明,2015-12-29,增加多单据分单据结算时医保结算失败时只对结算成功单据收费的模式。
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1121, 0, 0, 0, 0, 104, '只对医保结算成功单据收费', '0', '0',
         '单据分单据结算模式下，当医保结算失败，但部分单据结算成功时是否对结算成功的单据进行收费。0-只有所有单据都进行医保结算成功后才能继续收费，1-医保结算失败，但部分单据结算成功时只对结算成功的单据进行收费。'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1121 And 参数名 = '只对医保结算成功单据收费');

--91427:马政,2015-12-22,新增药品验收系统
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values( 1348,'药品入库验收管理','药品入库前检查入库单药品基本信息是否合格',&n_System,'zl9MediStore'); 

Insert Into zlMenus(组别,ID,上级ID,标题,快键,系统,模块,短标题,图标,说明)
  Select 组别,Zlmenus_Id.Nextval,id,'药品入库验收管理' ,'I' ,&n_System,1348 ,'入库验收' ,114 ,'药品入库前检查入库单药品基本信息是否合格' 
         From zlMenus Where 标题 = '药库管理与药品会计系统' And 组别 = '缺省' And 系统 = &n_System And 模块 is null;          

Insert Into zlMenus(组别,ID,上级ID,标题,快键,系统,模块,短标题,图标,说明)
  Select 组别,Zlmenus_Id.Nextval,id,'药品入库验收管理' ,'I' ,&n_System,1348 ,'入库验收' ,114 ,'药品入库前检查入库单药品基本信息是否合格' 
         From zlMenus Where 标题 = '门诊中西药房管理系统' And 组别 = '缺省' And 系统 = &n_System And 模块 is null;          

Insert Into 号码控制表
  (项目序号, 项目名称, 自动补缺, 编号规则)
  Select 148, '药品入库验收', 0, 0
  From Dual
  Where Not Exists (Select 1 From 号码控制表 Where 项目序号 = 148 And 项目名称 = '药品入库验收');

--89983:陈刘,2015-12-22,呼吸表格呼吸机输出方式,新增参数值
Update zlParameters
Set 参数说明 = '当呼吸设置为表格项目，此参数决定呼吸机的输出方式。1.参数值为0，呼吸表格栏内显示R符号。2.参数值为1，在每段连续呼吸机开始对应的呼吸表格栏上方纵向输出"呼吸机"，且用"↑"标识开始，在对应结束的呼吸表格栏上方用"↓"标识终止。3.参数值为2,呼吸表格栏内显示A+呼吸值'
Where 系统 = &n_System And 模块 = 1255 And 参数号 = 85;

--91225:梁经伙,2015-12-22,传染病管理系统添加模块号和菜单
Insert Into zlComponent(部件,名称,主版本,次版本,附版本,系统,注册产品名称,注册产品简名,注册产品版本) Values('zl9Disease','传染病系统部件',10,34,0,&n_System,'中联医院信息系统','ZLHIS+','10');
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values( 1278,'传染病管理工作站','用于对传染病报告单的审核、上报等管理工作',&n_System,'zl9Disease');

Insert Into zlMenus
  (组别, ID, 上级id, 标题, 快键, 说明, 系统, 模块, 短标题, 图标)
  Select 组别, Zlmenus_Id.Nextval, ID, '传染病管理系统', 'D', ' 用于对传染病报告单的审核、上报等管理工作 ', &n_System, -null, '传染病管理', 99
  From zlMenus
  Where 标题 = '临床信息系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null;

Insert Into zlMenus
  (组别, ID, 上级id, 标题, 快键, 说明, 系统, 模块, 短标题, 图标)
  Select a.组别, Zlmenus_Id.Nextval, a.Id, b.*
  From (Select 组别, ID From zlMenus Where 标题 = '传染病管理系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null) A,
       (Select 标题, 快键, 说明, 系统, 模块, 短标题, 图标
         From zlMenus
         Where 1 = 0
         Union All
         Select '传染病管理工作站', 'D', ' 用于对传染病报告的接收、审核、上报等管理工作', &n_System, 1278, '传染病管理', 130
         From Dual
         Union All
         Select 标题, 快键, 说明, 系统, 模块, 短标题, 图标
         From zlMenus
         Where 1 = 0) B;

--91225:胡俊勇,2015-12-21,传染病管理系统
Update zlParameters
Set 参数说明 = '每位数分别代表不同消息类型：1病历审阅、2医嘱安排、3危急值、4报告撤销、5医嘱审核、6传染病报告'
Where 参数名 = '自动刷新内容' And 模块 = 1261 and 系统 = &n_System;

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1260, 1, 0, 0, 0, 28, '自动刷新病历审阅间隔', '', '0',
         '设置每多少分钟自动刷新病历审阅提醒区域中的内容，为0表示不自动刷新(可手工刷新)。'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1260 And 参数名 = '自动刷新病历审阅间隔');

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1260, 1, 0, 0, 0, 29, '自动刷新内容', '', '0', '每位数分别代表不同消息类型：1传染病报告'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1260 And 参数名 = '自动刷新内容');

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1260, 1, 0, 0, 0, 30, '自动刷新病历审阅天数', '', '1', '设置将多少天内完成的病历显示在审阅提醒区域。'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1260 And 参数名 = '自动刷新病历审阅天数');

--91225:梁经伙,2015-12-22,传染病管理系统添加参数
Insert Into zlParameters(ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
      Select Zlparameters_Id.Nextval, &n_System, 1278, a.* From (Select 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明 From zlParameters Where 1 = 0 Union All
      Select 1, 0, 0, 0, 1, '本工作站可管理文件', Null, Null, '传染病工作站可管理的文件,是病历文件列表中的文件ID' From Dual Union All
      Select 1, 0, 0, 0, 2, '审核与上报工作状态下查看最近天数的报告', '7', '7', '审核与上报工作状态下查看最近多少天的报告' From Dual Union All
      Select 1, 0, 0, 0, 3, '审核与上报工作状态下查看指定天数的报告的起始天数', Null, Null, '审核与上报工作状态下查看指定时间段的报告的起始日期' From Dual Union All
      Select 1, 0, 0, 0, 4, '审核与上报工作状态下查看指定天数的报告的结束天数', Null, Null, '审核与上报工作状态下查看指定时间段的报告的结束日期' From Dual Union All
      Select 1, 0, 0, 0, 5, '传染病系统查看状态范围', '1,1,1,0,1,0', '1,1,1,0,1,0','选择查看状态范围的的报告,所在的位为1的话代表启用查看该状态都报告，为0的话代表不查看；第1位-待审核,第2位-待返修,第3位-待上报,第4位-已上报,第5位-待填写报告卡,第6位-非传染病' From Dual Union All
      Select 1, 0, 0, 0, 6, '未填写状态下查看最近天数的报告', '0', '0', '未填写状态下查看最近多少天的报告' From Dual Union All
      Select 1, 0, 0, 0, 7, '未填写状态下查看指定天数的报告的起始天数', Null, Null, '未填写状态下查看指定时间段的报告的起始日期' From Dual Union All
      Select 1, 0, 0, 0, 8, '未填写状态下查看指定天数的报告的结束天数', Null, Null, '未填写状态下查看指定时间段的报告的结束日期' From Dual Union All
      Select 1, 0, 0, 0, 9, '已删除状态下查看最近天数的报告', '7', '7', '已删除状态下查看最近多少天的报告' From Dual Union All
      Select 1, 0, 0, 0, 10, '已删除状态下查看指定天数的报告的起始天数', Null, Null, '已删除状态下查看指定时间段的报告的起始日期' From Dual Union All
      Select 1, 0, 0, 0, 11, '已删除状态下查看指定天数的报告的结束天数', Null, Null, '已删除状态下查看指定时间段的报告的结束日期' From Dual Union All
      Select 1, 0, 0, 0, 12, '当前查看报告的工作状态', '1', '1', '0-待填写,1-审核工作，2-上报工作，3-已删除，4-查重工作' From Dual Union All
      Select 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明 From zlParameters Where 1 = 0) A;


Insert Into 业务消息类型(编码,名称,说明) 
Select 'ZLHIS_CIS_032','传染病阳性结果提醒','技师站填写传染病阳性录时，产生的一个通知消息。' From Dual Union All
Select 'ZLHIS_CIS_033','传染病报告返修提醒','传染病报告填写不符合要求，产生的一个返修通知消息。' From Dual;

--91225:梁经伙,2015-12-16,传染病管理系统 基本数据
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( &n_sysTem,'传染病目录',0,'法定传染病目录','医疗工作' ); 

Insert Into 传染病目录(编码,名称,简码) 
Select '01','鼠疫','SY' From Dual Union All
Select '02','霍乱','HL' From Dual Union All
Select '03','传染性非典型肺炎','CRXFDXFY' From Dual Union All
Select '04','艾滋病(HIV)','AZBHIV' From Dual Union All
Select '05','艾滋病(AIDS)','AZBAIDS' From Dual Union All
Select '06','病毒性肝炎(甲型)','BDXGYJX' From Dual Union All
Select '07','病毒性肝炎(乙型)','BDXGYYX' From Dual Union All
Select '08','病毒性肝炎(丙型)','BDXGYBX' From Dual Union All
Select '09','病毒性肝炎(戊型)','BDXGYWX' From Dual Union All
Select '10','病毒性肝炎(未分型)','BDXGYWFX' From Dual Union All
Select '11','脊髓灰质炎','GSHZY' From Dual Union All
Select '12','人感染高致病性禽流感','RGRGZBXQLG' From Dual Union All
Select '13','甲型H1N1流感','JXH1N1LG' From Dual Union All
Select '14','麻疹','MZ' From Dual Union All
Select '15','流行性出血热','LXXCXR' From Dual Union All
Select '16','狂犬病','KQB' From Dual Union All
Select '17','流行性乙型脑炎','LXXYXGY' From Dual Union All
Select '18','登革热','DGR' From Dual Union All
Select '19','炭疽(肺炭疽)','TJFTJ' From Dual Union All
Select '20','炭疽(未分型)','TJWFX' From Dual Union All
Select '21','痢疾(细菌性)','LJXJX' From Dual Union All
Select '22','痢疾(阿米巴性)','LJAMBX' From Dual Union All
Select '23','肺结核(涂阳)','FJHTY' From Dual Union All
Select '24','肺结核(仅培阳)','FJHJPY' From Dual Union All
Select '25','肺结核(菌阴)','FJHJY' From Dual Union All
Select '26','肺结核(未痰检)','FJHWTJ' From Dual Union All
Select '27','伤寒(伤寒)','SHSH' From Dual Union All
Select '28','伤寒(副伤寒)','SHFSH' From Dual Union All
Select '29','流行性脑脊髓膜炎','LXXLJSMY' From Dual Union All
Select '30','百日咳','BRK' From Dual Union All
Select '31','白喉','BH' From Dual Union All
Select '32','新生儿破伤风','XSEPSF' From Dual Union All
Select '33','猩红热','XHR' From Dual Union All
Select '34','布鲁氏菌病','BLSJB' From Dual Union All
Select '35','淋病、梅毒(Ⅰ期)','LBMDYQ' From Dual Union All
Select '36','淋病、梅毒(Ⅱ期)','LBMDEQ' From Dual Union All
Select '37','淋病、梅毒(Ⅲ期)','LBMDSQ' From Dual Union All
Select '38','淋病、梅毒(胎传)','LBMDTC' From Dual Union All
Select '39','淋病、梅毒(隐性)','LBMDYX' From Dual Union All
Select '40','钩端螺旋体病','GDLXTB' From Dual Union All
Select '41','血吸虫病','XXCB' From Dual Union All
Select '42','疟疾(间日疟)','LJJRL' From Dual Union All
Select '43','疟疾(恶性疟)','LJEXL' From Dual Union All
Select '44','疟疾(未分型)','LJWFX' From Dual Union All
Select '45','流行性感冒','LXXGM' From Dual Union All
Select '46','流行性腮腺炎','LXXSXY' From Dual Union All
Select '47','风疹','FZ' From Dual Union All
Select '48','急性出血性结膜炎','JXCXXJMY' From Dual Union All
Select '49','麻风病','MFB' From Dual Union All
Select '50','流行性和地方性斑疹伤寒','LXXHDFXBZSH' From Dual Union All
Select '51','黑热病','HRB' From Dual Union All
Select '52','包虫病','BCB' From Dual Union All
Select '53','丝虫病','SCB' From Dual Union All
Select '54','除霍乱、细菌性和阿米巴性痢疾、伤寒和副伤寒以外的感染性腹泻病','CHLXJXHAMBX' From Dual Union All
Select '55','手足口病','SZKB' From Dual;

--91225:梁经伙,2015-12-16,传染病管理系统
Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
Select &n_System,6,A.* From (
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0 Union All
Select '疾病报告反馈',4,1,-NULL From Dual Union All
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0) A;

--91064:刘硕,2015-12-08,外院医生输入控制
Insert Into Zlparameters
  (Id, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 253, '外院医生必须先建档', '0', '0',
         '在选择医生的信息域是否可以自由录入外院医生。0-可以自由录入外院医生；1-不能自由录入外院医生，外院医生必须先建档'
  From Dual
  Where Not Exists (Select 1
         From Zlparameters
         Where 参数名 = '外院医生必须先建档' And Nvl(模块, 0) = 0 And Nvl(系统, 0) = &n_System);

--90666:陈刘,2015-12-07,按江苏省病例规范要求,新增体温部位:耳温
Insert Into 体温部位 (项目序号, 部位, 缺省项, 固定项) Values (1, '耳温', 0, 1);

Update 体温记录项目 Set 记录符 = '·,×,○,△' Where 项目序号 = 1;

--78413:胡俊勇,2015-12-01,医嘱清单打印
Delete From zlProgPrivs Where Upper(对象) = 'ZL_医嘱打印记录_INSERT' or 对象='医嘱打印记录';

--91641:余伟节,2015-12-25,路径匹配时加入期效
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1256, 0, 0, 0, 0, 12, '匹配时期效不同算路径外项目', '0', '0',
         '0-诊疗项目相同时,期效不相同当作路径内项目,但优先匹配相同期效,1-诊疗项目和期效都相同时才算作路径内项目'
  From Dual
  Where Not Exists (Select 1
         From zlParameters
         Where 参数名 = '匹配时期效不同算路径外项目' And Nvl(模块, 0) = 1256 And Nvl(系统, 0) = &n_System);



-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--92335:李南春,2016-01-18,三方支付新模式及过程拆分
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1801, '基本', User, 'zl_人员缴款余额_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1801 And 功能 = '基本' And Upper(对象) = Upper('zl_人员缴款余额_Update'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1802, '基本', User, 'zl_人员缴款余额_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1802 And 功能 = '基本' And Upper(对象) = Upper('zl_人员缴款余额_Update'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1803, '基本', User, 'zl_人员缴款余额_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1803 And 功能 = '基本' And Upper(对象) = Upper('zl_人员缴款余额_Update'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1804, '基本', User, 'zl_人员缴款余额_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1804 And 功能 = '基本' And Upper(对象) = Upper('zl_人员缴款余额_Update'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1805, '基本', User, 'zl_人员缴款余额_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1805 And 功能 = '基本' And Upper(对象) = Upper('zl_人员缴款余额_Update'));

--89620:余伟节,2016-01-15,提前完成临床路径执行的权限
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1256,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '提前完成',11,'提前完成临床路径执行的权限。',1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

--91487:冉俊明,2016-01-05,保险补充结算多笔退费。
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1124, '医保结算', User, '三方退款信息', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1124 And 功能 = '医保结算' And Upper(对象) = Upper('三方退款信息'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1124, '结算退费', User, '三方退款信息', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1124 And 功能 = '结算退费' And Upper(对象) = Upper('三方退款信息'));        

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1124, '医保结算', User, 'Zl_三方退款信息_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1124 And 功能 = '医保结算' And Upper(对象) = Upper('Zl_三方退款信息_Insert'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1124, '结算退费', User, 'Zl_三方退款信息_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1124 And 功能 = '结算退费' And Upper(对象) = Upper('Zl_三方退款信息_Insert'));

--91427:马政,2015-12-22,新增药品验收系统
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1348,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All 
    Select '基本',-Null,NULL,1 From Dual Union All      
    Select '新增',2,'增加药品入库验收管理的操作权限。有该权限时，允许增加入库验收单',1 From Dual Union All 
    Select '修改',4,'对未审核的药品进行修改的操作权限。有该权限时，允许对未审核的入库验收单进行修改',1 From Dual Union All 
    Select '删除',5,'删除药品入库验收管理记录的操作权限。有该权限时，允许对未审核的入库验收单进行删除',1 From Dual Union All 
    Select '审核',6,'增加药品外购放库记录审核的操作权限。有该权限时，允许对入库验收单进行审核',1 From Dual Union All     
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1348,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select 'NextNO','EXECUTE' From Dual Union All 
    Select '号码控制表','SELECT' From Dual Union All 
    Select '号码控制表','UPDATE' From Dual Union All 
    Select '科室号码表','SELECT' From Dual Union All 
    Select '科室号码表','UPDATE' From Dual Union All
    Select '部门表','SELECT' From Dual Union All 
    Select '部门人员','SELECT' From Dual Union All 
    Select '部门性质分类','SELECT' From Dual Union All 
    Select '部门性质说明','SELECT' From Dual Union All 
    Select '供应商','SELECT' From Dual Union All 
    Select '药品生产商','SELECT' From Dual Union All 
    Select '人员表','SELECT' From Dual Union All 
    Select '上机人员表','SELECT' From Dual Union All 
    Select '收费价目','SELECT' From Dual Union All 
    Select '收费细目','SELECT' From Dual Union All 
    Select '收费项目别名','SELECT' From Dual Union All 
    Select '收费项目目录','SELECT' From Dual Union All 
    Select '收费执行科室','SELECT' From Dual Union All 
    Select '药品别名','SELECT' From Dual Union All 
    Select '药品材质分类','SELECT' From Dual Union All 
    Select '药品出库检查','SELECT' From Dual Union All 
    Select '药品单据性质','SELECT' From Dual Union All 
    Select '药品规格','SELECT' From Dual Union All 
    Select '药品剂型','SELECT' From Dual Union All     
    Select '药品目录','SELECT' From Dual Union All 
    Select '药品入出类别','SELECT' From Dual Union All    
    Select '药品特性','SELECT' From Dual Union All 
    Select '药品外观','SELECT' From Dual Union All 
    Select '药品卫材精度','SELECT' From Dual Union All     
    Select '诊疗分类目录','SELECT' From Dual Union All 
    Select '诊疗项目类别','SELECT' From Dual Union All 
    Select '诊疗项目目录','SELECT' From Dual Union All 
    Select '诊疗执行科室','SELECT' From Dual Union All 
    Select '药品验收记录','SELECT' From Dual Union All 
    Select '药品验收明细','SELECT' From Dual Union All 
    Select '药品储备限额','SELECT' From Dual Union All
    Select '药品验收记录_ID','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1348,'新增',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_药品验收记录_Insert','EXECUTE' From Dual Union All
Select 'Zl_药品验收明细_Insert','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1348,'修改',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_药品验收记录_Insert','EXECUTE' From Dual Union All
Select 'ZL_药品验收记录_Delete','EXECUTE' From Dual Union All
Select 'Zl_药品验收明细_Insert','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1348,'删除',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'ZL_药品验收记录_Delete','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1348,'审核',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'ZL_药品验收记录_Verify','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--91225:梁经伙,2015-12-22,传染病管理系统添加权限
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1278,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '基本',-NULL,NULL,1 From Dual Union All
Select '范围设置',1,'设置工作站可管理的疾病报告范围。有该权限时，允许设置本工作站可管理的文件',1 From Dual Union All
Select '报送',2,'疾病报告的对外报送信息登记。有该权限时，允许对疾病报告的对外报送信息进行登记',1 From Dual Union All
Select '回退',3,'取消错误的报送登记或接收拒绝操作。有该权限时，允许对疾病报告的登记接收拒绝操作进行回退',1 From Dual Union All
Select '审核',4,'审核填写了的疾病报告。有该权限时，允许对疾病报告的进行审核',1 From Dual Union All
Select '删除',5,'删除重复的疾病报告。有该权限时，允许对疾病报告进行删除',1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1278,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'ZL_Replace_Element_Value','EXECUTE' From Dual Union All
Select 'Zl_Lob_Read','EXECUTE' From Dual Union All
Select 'Zl_电子病历打印_Insert','EXECUTE' From Dual Union All
Select 'Zl_疾病申报对应_UPDATE','EXECUTE' From Dual Union All
Select '病案主页','SELECT' From Dual Union All
Select '病历常用样式','SELECT' From Dual Union All
Select '病历文件结构','SELECT' From Dual Union All
Select '病历文件列表','SELECT' From Dual Union All
Select '病历页面格式','SELECT' From Dual Union All
Select '病人信息','SELECT' From Dual Union All
Select '病人医疗卡属性','SELECT' From Dual Union All
Select '病人医疗卡信息','SELECT' From Dual Union All
Select '部门性质说明','SELECT' From Dual Union All
Select '电子病历附件','SELECT' From Dual Union All
Select '电子病历记录','SELECT' From Dual Union All
Select '电子病历内容','SELECT' From Dual Union All
Select '疾病报送单位','SELECT' From Dual Union All
Select '疾病申报对应','SELECT' From Dual Union All
Select '疾病申报记录','SELECT' From Dual Union All
Select '卡消费接口目录','SELECT' From Dual Union All
Select '人员性质说明','SELECT' From Dual Union All
Select '消费卡目录','SELECT' From Dual Union All
Select '医疗卡挂失方式','SELECT' From Dual Union All
Select '医疗卡类别','SELECT' From Dual Union All
Select '疾病报告反馈','SELECT' From Dual Union All
Select '病人照片','SELECT' From Dual Union All
Select '疾病阳性记录','SELECT' From Dual Union All
Select '病人挂号记录','SELECT' From Dual Union All
Select '部门表','SELECT' From Dual Union All
Select '人员表','SELECT' From Dual Union All
Select '部门人员','SELECT' From Dual Union All
Select '职业','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1278,'报送',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_疾病申报记录_Send','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1278,'审核',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_疾病申报记录_Update','EXECUTE' From Dual Union All
Select 'Zl_疾病申报记录_Incept','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1278,'回退',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_疾病申报记录_Untread','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1278,'删除',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_疾病申报记录_Delete','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;


--91866:许华峰,2015-12-21,传染病管理系统阳性结果反馈单登记、查询功能
Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select &n_System,1290,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0 Union All
Select NULL,&n_System,9001,1,'基本',1 From Dual Union All
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0) A;

Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select &n_System,1291,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0 Union All
Select NULL,&n_System,9001,1,'基本',1 From Dual Union All
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0) A;

Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select &n_System,1294,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0 Union All
Select NULL,&n_System,9001,1,'基本',1 From Dual Union All
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0) A;

--91225:梁经伙,2015-12-21,传染病管理系统新增加表
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select  &n_System,9001,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '传染病阳性结果登记',2,'有此权限时，允许调用接口对传染病阳性结果进行登记',1 From Dual Union All
Select '传染病阳性结果查询',3,'有此权限时，允许调用接口对传染病阳性结果进行查询',1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select  &n_System,9001,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '疾病阳性记录','SELECT' From Dual Union All 
    Select '传染病目录','SELECT' From Dual Union All 
    Select '疾病报告前提','SELECT' From Dual Union All     
    Select '诊疗检验标本','SELECT' From Dual Union All 
    Select 'Zl_疾病阳性检测记录_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_疾病阳性检测记录_Update','EXECUTE' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--91225:梁经伙,2015-12-21,传染病阳性结果反馈单查询
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,9001,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '病人照片','SELECT' From Dual Union All    
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--91738:张险华,2015-12-17,电子病案审查访问新版PACS报告
Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select &n_System,1560,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0 Union All
Select '基本',&n_System,9004,1,'基本',1 From Dual Union All
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0) A;

Insert Into Zltools.Zlrolegrant
 (系统, 序号, 角色, 功能)
 Select Distinct &n_System, 9004, 角色, '基本'
 From Zltools.Zlrolegrant A
 Where 序号 = 1560 And Not Exists (Select 1 From Zltools.Zlrolegrant Where 序号 = 9004 And 角色 = a.角色);

--89242:李南春,2015-12-08,使用结构化地址控件
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1107, '基本', User, 'zl_病人地址信息_update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1107 And 功能 = '基本' And Upper(对象) = Upper('zl_病人地址信息_update'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1111, '基本', User, 'zl_病人地址信息_update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1111 And 功能 = '基本' And Upper(对象) = Upper('zl_病人地址信息_update'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1113, '病案修改', User, 'zl_病人地址信息_update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1113 And 功能 = '病案修改' And Upper(对象) = Upper('zl_病人地址信息_update'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1107, '基本', User, 'Zl_Adderss_Structure', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1107 And 功能 = '基本' And Upper(对象) = Upper('Zl_Adderss_Structure'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1111, '基本', User, 'Zl_Adderss_Structure', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1111 And 功能 = '基本' And Upper(对象) = Upper('Zl_Adderss_Structure'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1113, '病案修改', User, 'Zl_Adderss_Structure', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1113 And 功能 = '病案修改' And Upper(对象) = Upper('Zl_Adderss_Structure'));

--89620:余伟节,2016-01-15,提前完成临床路径执行的权限
Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1256,2,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0 Union All
Select '结束路径',2,1,0 From Dual Union All
Select '提前完成',2,0,0 From Dual Union All
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0) A;






-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------

--91225:梁唐彬,2016-01-19,传染病管理系统
--报表：ZL1_REPORT_1280/法定传染病报告登记表
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1280','法定传染病报告登记表','查询并展示传染病报告登记记录','H`;~@e`~{( PlscuZ,\L','Microsoft XPS Document Writer',15,0,0,100,1280,'基本',Sysdate,Sysdate,Null,Null);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'法定传染病报告登记表',11906,16838,9,2,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'登记记录','ROWNUM,139|姓名,202|性别,202|出生日期,202|职业,202|家庭地址,202|电话,202|发病日期,202|确诊日期,202|实验,130|临床,130|携带,130|疑似,130|诊断,202|报告科室,202|填报日期,202|填报人,202|收卡日期,202|收卡人,202|网络报告日期,202|备注,202',User||'.电子病历记录,'||User||'.疾病申报记录,'||User||'.病人信息,'||User||'.部门表',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select Rownum, 姓名, 性别, 出生日期, 职业, 家庭地址, 电话, 发病日期, 确诊日期, 实验, 临床, 携带, 疑似, 诊断, 报告科室, 填报日期, 填报人, 收卡日期, 收卡人, 网络报告日期, 备注' From Dual Union All
  Select 2,'From (Select p.姓名, p.性别, to_char(p.出生日期,''yyyy-mm-dd'') 出生日期, p.职业, p.家庭地址, p.联系人电话 电话, to_char(s.发病日期,''yyyy-mm-dd'') 发病日期, to_char(s.确诊日期,''yyyy-mm-dd'')确诊日期, '''' 实验, '''' 临床, '''' 携带, '''' 疑似,' From Dual Union All
  Select 3,'              s.诊断描述1 || s.诊断描述2 诊断, d.名称 报告科室, to_char(l.完成时间,''yyyy-mm-dd'') 填报日期, l.保存人 填报人, to_char(Trunc(s.收拒时间),''yyyy-mm-dd'') 收卡日期, s.收拒人 收卡人,' From Dual Union All
  Select 4,'              to_char(Trunc(s.报送时间),''yyyy-mm-dd'') 网络报告日期, s.填报备注 备注' From Dual Union All
  Select 5,'       From 电子病历记录 L, 疾病申报记录 S, 病人信息 P, 部门表 D' From Dual Union All
  Select 6,'       Where l.病历种类 = 5 And l.文件id In ([0]) And l.完成时间 Between [1] And [2] And' From Dual Union All
  Select 7,'             l.Id = s.文件id(+) And s.处理状态(+) <> -1 And l.病人id = p.病人id And l.科室id = d.Id' From Dual Union All
  Select 8,'       Order By p.姓名)' From Dual Union All
  Select 9,Null From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'文件',0,'选择器定义…',0,Null,Null,'select id ,名称 From 病历文件列表 where 种类=5',Null,'ID,131,'||CHR(38)||'S'||CHR(38)||'B|名称,202,'||CHR(38)||'D'||CHR(38)||'S',User||'.病历文件列表|',Null);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,1,'开始时间',2,CHR(38)||'上月初时间',0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,2,'结束时间',2,CHR(38)||'上月末时间',0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,'任意表1',11,'统计时间:[=开始时间]  -  [=结束时间]',Null,90,615,3780,210,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,'任意表1',12,'法定传染病报告登记表',Null,6720,120,3300,330,0,0,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,'登记记录',Null,90,975,16560,10470,450,0,1,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[登记记录.ROWNUM]','4^450^编号',0,0,315,0,0,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[登记记录.姓名]','4^450^姓名',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[登记记录.性别]','4^450^性别',0,0,315,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[登记记录.出生日期]','4^450^出生日期',0,0,990,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[登记记录.职业]','4^450^职业',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[登记记录.家庭地址]','4^450^家庭地址',0,0,1245,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[登记记录.电话]','4^450^电话',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[登记记录.发病日期]','4^450^发病日期',0,0,1005,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[登记记录.确诊日期]','4^450^确诊日期',0,0,1065,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[登记记录.实验]','4^450^实验',0,0,285,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,10,Null,Null,'[登记记录.临床]','4^450^临床',0,0,270,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,11,Null,Null,'[登记记录.携带]','4^450^携带',0,0,255,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-13,12,Null,Null,'[登记记录.疑似]','4^450^疑似',0,0,285,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-14,13,Null,Null,'[登记记录.诊断]','4^450^诊断',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-15,14,Null,Null,'[登记记录.报告科室]','4^450^报告科室',0,0,825,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-16,15,Null,Null,'[登记记录.填报日期]','4^450^填报日期',0,0,990,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-17,16,Null,Null,'[登记记录.填报人]','4^450^填报人',0,0,840,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-18,17,Null,Null,'[登记记录.收卡日期]','4^450^收卡日期',0,0,990,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-19,18,Null,Null,'[登记记录.收卡人]','4^450^收卡人',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-20,19,Null,Null,'[登记记录.网络报告日期]','4^450^网络报告日期',0,0,810,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-21,20,Null,Null,'[登记记录.备注]','4^450^备注',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);

--报表：ZL1_REPORT_1280/法定传染病报告登记表
Insert into zlPrograms(序号,标题,说明,系统,部件) Values(1280,'法定传染病报告登记表','查询并展示传染病报告登记记录',100,'zl9Report');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1280,'基本',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select 100,1280,'基本',User,'病历文件列表','SELECT' From Dual Union All
  Select 100,1280,'基本',User,'病人信息','SELECT' From Dual Union All
  Select 100,1280,'基本',User,'部门表','SELECT' From Dual Union All
  Select 100,1280,'基本',User,'电子病历记录','SELECT' From Dual Union All
  Select 100,1280,'基本',User,'疾病申报记录','SELECT' From Dual;
Insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'法定传染病报告登记表','法定传染病报告登记表',Null,105,'查询并展示传染病报告登记记录',100,1280 From zlMenus Where 系统=100 And 组别='缺省' And 标题='传染病管理系统' And 模块 is NULL;

--报表：ZL1_REPORT_1281/传染病阳性检测结果一览表
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1281','传染病阳性检测结果一览表','查询传染病阳性检测结果','Mv:uZldpv3%Fmxx}^"QW',Null,15,0,0,100,1281,'基本',Sysdate,Sysdate,To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'传染病阳性检测结果一览表1',11904,16832,256,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'疾病阳性记录_数据','ID,131|来源,130|病人ID,131|姓名,202|性别,202|年龄,202|科室,202|标识号,131|送检时间,202|送检医生,202|送检科室,202|标本名称,202|反馈结果,202|疑似疾病,202|登记人,202|登记时间,202|处理人,202|处理时间,202|处理情况说明,202',User||'.疾病阳性记录,'||User||'.病案主页,'||User||'.病人挂号记录,'||User||'.部门表',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'select A.Id, A.来源, 病人id, A.姓名, A.性别,A.年龄,e.名称 as 科室, A.标识号,A.送检时间, A.送检医生,f.名称 As 送检科室,A.标本名称, A.反馈结果,  A.疑似疾病, A.登记人,A.登记时间,A.处理人, A.处理时间, A.处理情况说明' From Dual Union All
  Select 2,'from ' From Dual Union All
  Select 3,'(Select a.Id,  ''住院'' As 来源, a.病人id, c.姓名, c.性别,c.年龄,' From Dual Union All
  Select 4,'      C.出院科室id as 科室ID, c.住院号 As 标识号,To_Char(a.送检时间, ''yyyy-MM-dd hh24:mi'')  送检时间, a.送检医生, a.送检科室id, a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, a.登记人,' From Dual Union All
  Select 5,'       To_Char(a.登记时间, ''yyyy-MM-dd hh24:mi'') 登记时间, a.处理人, To_Char(a.处理时间, ''yyyy-MM-dd hh24:mi'')  处理时间, a.处理情况说明' From Dual Union All
  Select 6,'From 疾病阳性记录 A, 病案主页 C' From Dual Union All
  Select 7,'Where a.病人id = c.病人id And a.主页id = c.主页id and a.登记时间 Between [0] And [1]' From Dual Union All
  Select 8,'union all' From Dual Union All
  Select 9,'Select a.Id,  ''门诊'' As 来源, a.病人id,  b.姓名 , b.性别 , b.年龄,' From Dual Union All
  Select 10,'        b.执行部门id as 科室ID , b.门诊号 As 标识号,To_Char(a.送检时间, ''yyyy-MM-dd hh24:mi'')  送检时间, a.送检医生, a.送检科室ID, a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, a.登记人,' From Dual Union All
  Select 11,'       To_Char(a.登记时间, ''yyyy-MM-dd hh24:mi'') 登记时间, a.处理人, To_Char(a.处理时间, ''yyyy-MM-dd hh24:mi'')  处理时间, a.处理情况说明' From Dual Union All
  Select 12,'From 疾病阳性记录 A, 病人挂号记录 B' From Dual Union All
  Select 13,'Where  a.病人id = b.病人id And a.挂号单 = b.No And a.登记时间 Between [0] And [1]) A ,部门表 E, 部门表 F' From Dual Union All
  Select 14,'where  a.送检科室id = f.Id(+) And  A.科室ID = e.Id(+)' From Dual Union All
  Select 15,'Order By a.Id' From Dual Union All
  Select 16,Null From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'开始日期',2,CHR(38)||'本月初时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,1,'结束日期',2,CHR(38)||'本月末时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'操作员',2,Null,0,'任意表1',21,'统计人：[操作员姓名]',Null,435,14685,2100,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'报表名称',2,Null,0,'任意表1',12,'[单位名称]传染病阳性检测结果一览表',Null,2778,390,6120,375,0,0,1,'宋体',18,0,0,0,0,16777215,0,Null,Null,Null,1,0,1,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'打印时间',2,Null,0,'任意表1',23,'打印时间：[yyyy-MM-dd hh:mm:ss]',Null,7985,14670,3255,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,'任意表1',13,'日期：[=开始日期]至[=结束日期]',Null,8090,1080,3150,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,'疾病阳性记录_数据',Null,435,1500,10805,12960,255,0,0,'宋体',10.5,0,0,0,0,16777215,1,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[疾病阳性记录_数据.来源]','4^225^来源',0,0,1005,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[疾病阳性记录_数据.姓名]','4^225^姓名',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[疾病阳性记录_数据.性别]','4^225^性别',0,0,1005,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[疾病阳性记录_数据.年龄]','4^225^年龄',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[疾病阳性记录_数据.科室]','4^225^科室',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[疾病阳性记录_数据.标识号]','4^225^标识号',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[疾病阳性记录_数据.送检时间]','4^225^送检时间',0,0,2040,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[疾病阳性记录_数据.送检医生]','4^225^送检医生',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[疾病阳性记录_数据.送检科室]','4^225^送检科室',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[疾病阳性记录_数据.标本名称]','4^225^标本名称',0,0,1155,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,10,Null,Null,'[疾病阳性记录_数据.反馈结果]','4^225^反馈结果',0,0,1785,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,11,Null,Null,'[疾病阳性记录_数据.疑似疾病]','4^225^疑似疾病',0,0,1365,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-13,12,Null,Null,'[疾病阳性记录_数据.登记人]','4^225^登记人',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-14,13,Null,Null,'[疾病阳性记录_数据.登记时间]','4^225^登记时间',0,0,2010,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-15,14,Null,Null,'[疾病阳性记录_数据.处理人]','4^225^处理人',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-16,15,Null,Null,'[疾病阳性记录_数据.处理时间]','4^225^处理时间',0,0,2010,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-17,16,Null,Null,'[疾病阳性记录_数据.处理情况说明]','4^225^处理情况说明',0,0,2205,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);

--报表：ZL1_REPORT_1281/传染病阳性检测结果一览表
Insert into zlPrograms(序号,标题,说明,系统,部件) Values(1281,'传染病阳性检测结果一览表','查询传染病阳性检测结果',100,'zl9Report');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1281,'基本',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select 100,1281,'基本',User,'病案主页','SELECT' From Dual Union All
  Select 100,1281,'基本',User,'病人挂号记录','SELECT' From Dual Union All
  Select 100,1281,'基本',User,'部门表','SELECT' From Dual Union All
  Select 100,1281,'基本',User,'疾病阳性记录','SELECT' From Dual;
Insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'传染病阳性检测结果一览表','传染病阳性检测结果一览表',Null,105,'查询传染病阳性检测结果',100,1281 From zlMenus Where 系统=100 And 组别='缺省' And 标题='传染病管理系统' And 模块 is NULL;

--报表：ZL1_REPORT_1282/传染病分年龄汇总表
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1282','传染病分年龄汇总表','分年龄对传染病进行汇总','Mv:jLio`s)4ViooG*U\',Null,15,0,0,100,1282,'基本',Sysdate,Sysdate,To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'传染病分年龄性别汇总表1',11904,16832,9,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'疾病申报记录_数据','年龄,202|传染病名称,202|男,139|女,139|计,139',User||'.疾病申报记录,'||User||'.疾病阳性记录,'||User||'.疾病报告反馈',1,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select nvl(a.年龄, ''未知年龄'') as 年龄,  b.传染病名称, sum(decode(A.性别, ''男'',1,0)) as 男,sum(decode(A.性别, ''女'',1,0)) as 女,sum(decode(A.性别,''男'',1,1)) as 计 ' From Dual Union All
  Select 2,'From 疾病申报记录 A, 疾病阳性记录 B,疾病报告反馈 C ' From Dual Union All
  Select 3,'Where a.文件id = b.文件id  and A.文件ID = C.文件ID  and c.登记时间 Between [0] And [1]' From Dual Union All
  Select 4,'Group By nvl(a.年龄, ''未知年龄''),b.传染病名称' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'开始日期',2,CHR(38)||'本月初时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,1,'结束日期',2,CHR(38)||'本月末时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'统计人',2,Null,0,'汇总表1',21,'统计人：[操作员姓名]',Null,825,15790,2100,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标题',2,Null,0,'汇总表1',12,'[单位名称]传染病分年龄汇总表',Null,2975,615,6105,450,0,0,1,'宋体',22,0,0,0,0,16777215,0,Null,Null,Null,1,0,1,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'统计时间',2,Null,0,'汇总表1',23,'统计时间：[yyyy-mm-dd HH:MM:SS]',Null,7975,15790,3255,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'日期',2,Null,0,'汇总表1',13,'日期：[=开始日期]至[=结束日期]',Null,8080,1575,3150,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'汇总表1',5,Null,0,Null,0,'疾病申报记录_数据',Null,825,1965,10405,13605,255,0,0,'宋体',10.5,0,0,0,0,16777215,1,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,7,zlRPTItems_ID.CurrVal-1,0,Null,Null,'年龄',Null,0,0,1005,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,8,zlRPTItems_ID.CurrVal-2,0,Null,Null,'传染病名称',Null,0,0,1000,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-3,0,Null,Null,'男',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-4,1,Null,Null,'女',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-5,2,Null,Null,'计',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);

--报表：ZL1_REPORT_1282/传染病分年龄汇总表
Insert into zlPrograms(序号,标题,说明,系统,部件) Values(1282,'传染病分年龄汇总表','分年龄对传染病进行汇总',100,'zl9Report');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1282,'基本',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select 100,1282,'基本',User,'疾病报告反馈','SELECT' From Dual Union All
  Select 100,1282,'基本',User,'疾病申报记录','SELECT' From Dual Union All
  Select 100,1282,'基本',User,'疾病阳性记录','SELECT' From Dual;
Insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'传染病分年龄汇总表','传染病分年龄汇总表',Null,105,'分年龄对传染病进行汇总',100,1282 From zlMenus Where 系统=100 And 组别='缺省' And 标题='传染病管理系统' And 模块 is NULL;

--报表：ZL1_REPORT_1283/传染病分职业汇总表
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1283','传染病分职业汇总表','分职业对传染病进行汇总','Mv:jX}o`s)4Vi{{G*U\',Null,15,0,0,100,1283,'基本',Sysdate,Sysdate,To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'传染病分职业汇总表1',11904,16832,9,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病人医嘱记录_数据','职业,202|传染病名称,202|男,139|女,139|计,139',User||'.疾病申报记录,'||User||'.疾病阳性记录,'||User||'.疾病报告反馈',1,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select nvl(a.职业,''其他'') as 职业,  b.传染病名称, sum(decode(A.性别, ''男'',1,0)) as 男,sum(decode(A.性别, ''女'',1,0)) as 女,sum(decode(A.性别,''男'',1,1)) as 计 ' From Dual Union All
  Select 2,'From 疾病申报记录 A, 疾病阳性记录 B,疾病报告反馈 C ' From Dual Union All
  Select 3,'Where a.文件id = b.文件id  and A.文件ID = C.文件ID  and c.登记时间 Between [0] And [1]' From Dual Union All
  Select 4,'Group By nvl(a.职业,''其他''),b.传染病名称' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'开始日期',2,CHR(38)||'本月初时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,1,'结束日期',2,CHR(38)||'本月末时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'操作员',2,Null,0,'汇总表1',21,'统计人：[操作员姓名]',Null,645,15600,2100,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,'汇总表1',12,'[单位名称]传染病分职业汇总表',Null,2705,570,6105,450,0,0,1,'宋体',22,0,0,0,0,16777215,0,Null,Null,Null,1,0,1,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'日期',2,Null,0,'汇总表1',13,'日期：[=开始日期]至[=结束日期]',Null,7720,1350,3150,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,'汇总表1',23,'打印时间:[yyyy-mm-dd HH:MM]',Null,8035,15585,2835,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'汇总表1',5,Null,0,Null,0,'病人医嘱记录_数据',Null,645,1740,10225,13470,255,0,0,'宋体',10.5,0,0,0,0,16777215,1,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,7,zlRPTItems_ID.CurrVal-1,0,Null,Null,'职业',Null,0,0,1005,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,'SUM',1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,8,zlRPTItems_ID.CurrVal-2,0,Null,Null,'传染病名称',Null,0,0,1000,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,'SUM',1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-3,0,Null,Null,'男',Null,0,0,1335,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-4,1,Null,Null,'女',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-5,2,Null,Null,'计',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);

--报表：ZL1_REPORT_1283/传染病分职业汇总表
Insert into zlPrograms(序号,标题,说明,系统,部件) Values(1283,'传染病分职业汇总表','分职业对传染病进行汇总',100,'zl9Report');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1283,'基本',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select 100,1283,'基本',User,'疾病报告反馈','SELECT' From Dual Union All
  Select 100,1283,'基本',User,'疾病申报记录','SELECT' From Dual Union All
  Select 100,1283,'基本',User,'疾病阳性记录','SELECT' From Dual;
Insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'传染病分职业汇总表','传染病分职业汇总表',Null,105,'分职业对传染病进行汇总',100,1283 From zlMenus Where 系统=100 And 组别='缺省' And 标题='传染病管理系统' And 模块 is NULL;






-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--92736:刘尔旋,2016-01-19,结帐接口修改
Create Or Replace Procedure Zl_Third_Getsettlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取HIS结帐数据
  --入参:Xml_In:
  --<IN>
  -- <BRID></BRID>       //病人ID 
  -- <ZYID></ZYID>         //主页ID
  -- <JSLX></JSLX>       //结算类型。1-门诊,2-住院。固定传2
  -- <JSKLB></JSKLB>       //结算卡类别
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  --<JBXX>              //基本信息
  --   <XM></XM>           //姓名
  --   <XB></XB>           //性别
  --   <NL></NL>         //年龄
  --   <ZYH></ ZYH>        //住院号
  --   <ZYKS></ ZYKS>          //住院科室  
  --   <KSID></KSID>         //科室ID
  --   <ZZYS></ ZZYS>          //主治医生  
  --   <RYSJ></ RYSJ>          //入院时间
  --   <CYSJ></ CYSJ >         //出院时间 
  --   <JZSJ></JZSJ>         //结帐时间(未结帐为空)
  --   <DJH></DJH>         //单据号(未结帐为空)
  --   <JSZFY></JSZFY>         //结算总费用
  --</JBXX>
  --<YJKLIST>              //冲抵预缴款集合
  --   <ITEM>
  --     <DJH><DJH>        //预交款单据号
  --     <JSFS></JSFS>     //结算方式（为名称，返回什么就取什么）
  --     <JE></JE>           //预缴款金额
  --     <JYLSH></JYLSH>       //交易流水号（便于冲销使用）
  --     <SFJSK></SFJSK>       //是否结算卡，1-是，0-否。如果是由传入的卡类别缴费，返回1，否则返回0
  --   </ITEM>
  --</YJKLIST >
  --<TBQK>               //退补情况
  --   <TBLX></TBLX>         //退补类型(1:个人补款，2:医院退款)
  --   <TBJE></TBJE>         //退补金额
  --</TBQK>
  -- <ERROR><MSG></MSG></ERROR>    //出现错误时返回具体原因，error节点为空表示成功
  --</OUTPUT>  

  --------------------------------------------------------------------------------------------------
  n_病人id     病人信息.病人id%Type;
  n_主页id     病案主页.主页id%Type;
  n_结算类型   Number(3);
  v_结算卡类别 Varchar2(200);
  n_卡类别id   医疗卡类别.Id%Type;
  n_是否结清   Number(3); -- 1-未结清,0-结清
  n_结帐金额   住院费用记录.结帐金额%Type;
  v_Temp       Varchar2(32767); --临时XML
  v_Subtemp    Varchar2(32767);
  v_结帐ids    Varchar2(5000);
  n_退补金额   病人预交记录.冲预交%Type;
  n_病人余额   病人预交记录.金额%Type;
  n_结帐id     病人预交记录.结帐id%Type;
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/ZYID'), Extractvalue(Value(A), 'IN/JSLX'),
         Extractvalue(Value(A), 'IN/JSKLB')
  Into n_病人id, n_主页id, n_结算类型, v_结算卡类别
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  --默认住院结帐
  n_结算类型 := Nvl(n_结算类型, 2);
  Begin
    Select ID Into n_卡类别id From 医疗卡类别 Where 名称 = v_结算卡类别;
  Exception
    When Others Then
      v_Err_Msg := '无法确认传入的结算卡,请检查!';
      Raise Err_Item;
  End;
  If n_结算类型 = 2 Then
    Begin
      Select Distinct 1
      Into n_是否结清
      From 住院费用记录
      Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1 Having
       Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0;
    Exception
      When Others Then
        n_是否结清 := 0;
    End;
    If n_是否结清 = 0 Then
      --结清,读取结帐数据
      For r_结帐 In (Select 姓名, 性别, 年龄, 住院号, 住院科室, 科室id, 主治医生, To_Char(入院时间, 'yyyy-mm-dd') As 入院时间,
                          To_Char(出院时间, 'yyyy-mm-dd') As 出院时间, To_Char(结帐时间, 'yyyy-mm-dd') As 结帐时间, 单据号, 结算总费用, 结帐id
                   From (Select c.姓名, c.性别, c.年龄, c.住院号, e.名称 As 住院科室, c.入院科室id As 科室id, c.住院医师 As 主治医生, c.入院日期 As 入院时间,
                                 c.出院日期 As 出院时间, a.收费时间 As 结帐时间, a.No As 单据号, Sum(d.冲预交) As 结算总费用, a.Id As 结帐id
                          From 病人结帐记录 A, 病人信息 B, 病案主页 C, 病人预交记录 D, 部门表 E
                          Where a.记录状态 = 1 And a.病人id = c.病人id And a.病人id = b.病人id And b.主页id = c.主页id And a.病人id = n_病人id And
                                d.结帐id = a.Id And c.入院科室id = e.Id(+) And Exists
                           (Select 1 From 病人预交记录 Where 结帐id = a.Id And 结算方式 = v_结算卡类别)
                          Group By c.姓名, c.性别, c.年龄, c.住院号, e.名称, c.入院科室id, c.住院医师, c.入院日期, c.出院日期, a.收费时间, a.No, a.Id
                          Order By 结帐时间 Desc)
                   Where Rownum < 2) Loop
        v_Temp := '<XM>' || r_结帐.姓名 || '</XM>';
        v_Temp := v_Temp || '<XB>' || r_结帐.性别 || '</XB>';
        v_Temp := v_Temp || '<NL>' || r_结帐.年龄 || '</NL>';
        v_Temp := v_Temp || '<ZYH>' || r_结帐.住院号 || '</ZYH>';
        v_Temp := v_Temp || '<ZYKS>' || r_结帐.住院科室 || '</ZYKS>';
        v_Temp := v_Temp || '<KSID>' || r_结帐.科室id || '</KSID>';
        v_Temp := v_Temp || '<ZZYS>' || r_结帐.主治医生 || '</ZZYS>';
        v_Temp := v_Temp || '<RYSJ>' || r_结帐.入院时间 || '</RYSJ>';
        v_Temp := v_Temp || '<CYSJ>' || r_结帐.出院时间 || '</CYSJ>';
        v_Temp := v_Temp || '<JZSJ>' || r_结帐.结帐时间 || '</JZSJ>';
        v_Temp := v_Temp || '<DJH>' || r_结帐.单据号 || '</DJH>';
        v_Temp := v_Temp || '<JSZFY>' || r_结帐.结算总费用 || '</JSZFY>';
        v_Temp := '<JBXX>' || v_Temp || '</JBXX>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        n_结帐id := r_结帐.结帐id;
      End Loop;
      If n_结帐id Is Null Then
        v_Err_Msg := '该病人没有结帐数据!';
        Raise Err_Item;
      End If;
      v_Temp := '';
      For r_预交 In (Select NO As 单据号, 结算方式, Sum(冲预交) As 金额, 交易流水号, Max(卡类别id) As 卡类别id
                   From 病人预交记录
                   Where 结帐id = n_结帐id And Mod(记录性质, 10) = 1
                   Group By NO, 结算方式, 交易流水号
                   Order By 单据号 Desc) Loop
        v_Temp := '<DJH>' || r_预交.单据号 || '</DJH>';
        v_Temp := v_Temp || '<JSFS>' || r_预交.结算方式 || '</JSFS>';
        v_Temp := v_Temp || '<JE>' || r_预交.金额 || '</JE>';
        v_Temp := v_Temp || '<JYLSH>' || r_预交.交易流水号 || '</JYLSH>';
        If n_卡类别id = r_预交.卡类别id Then
          v_Temp := v_Temp || '<SFJSK>' || 1 || '</SFJSK>';
        Else
          v_Temp := v_Temp || '<SFJSK>' || 0 || '</SFJSK>';
        End If;
        v_Temp    := '<ITEM>' || v_Temp || '</ITEM>';
        v_Subtemp := v_Subtemp || v_Temp;
      End Loop;
      v_Subtemp := '<YJKLIST>' || v_Subtemp || '</YJKLIST>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Subtemp)) Into x_Templet From Dual;
      Select Nvl(Sum(冲预交), 0)
      Into n_退补金额
      From 病人预交记录
      Where 结帐id = n_结帐id And Mod(记录性质, 10) = 2 And Nvl(校对标志, 0) = 0;
      If n_退补金额 < 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(n_退补金额) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Else
      --未结清，读取未结数据
      For r_Info In (Select c.姓名, c.性别, c.年龄, c.住院号, d.名称 As 住院科室, c.入院科室id As 科室id, c.住院医师 As 主治医生,
                            To_Char(c.入院日期, 'yyyy-mm-dd') As 入院时间, To_Char(c.出院日期, 'yyyy-mm-dd') As 出院时间
                     From 病案主页 C, 部门表 D
                     Where c.病人id = n_病人id And c.入院科室id = d.Id(+) And c.主页id = n_主页id And Rownum < 2) Loop
        v_Temp := '<XM>' || r_Info.姓名 || '</XM>';
        v_Temp := v_Temp || '<XB>' || r_Info.性别 || '</XB>';
        v_Temp := v_Temp || '<NL>' || r_Info.年龄 || '</NL>';
        v_Temp := v_Temp || '<ZYH>' || r_Info.住院号 || '</ZYH>';
        v_Temp := v_Temp || '<ZYKS>' || r_Info.住院科室 || '</ZYKS>';
        v_Temp := v_Temp || '<KSID>' || r_Info.科室id || '</KSID>';
        v_Temp := v_Temp || '<ZZYS>' || r_Info.主治医生 || '</ZZYS>';
        v_Temp := v_Temp || '<RYSJ>' || r_Info.入院时间 || '</RYSJ>';
        v_Temp := v_Temp || '<CYSJ>' || r_Info.出院时间 || '</CYSJ>';
        v_Temp := v_Temp || '<JZSJ>' || '' || '</JZSJ>';
        v_Temp := v_Temp || '<DJH>' || '' || '</DJH>';
      End Loop;
      Begin
        Select Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0))
        Into n_结帐金额
        From 住院费用记录
        Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1;
      Exception
        When Others Then
          n_结帐金额 := 0;
      End;
      v_Temp := v_Temp || '<JSZFY>' || n_结帐金额 || '</JSZFY>';
      v_Temp := '<JBXX>' || v_Temp || '</JBXX>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Subtemp := '';
      For r_预交 In (Select NO As 单据号, 结算方式, Sum(Nvl(金额, 0)) - Sum(Nvl(冲预交, 0)) As 金额, 交易流水号, Max(卡类别id) As 卡类别id
                   From 病人预交记录
                   Where 病人id = n_病人id And Mod(记录性质, 10) = 1 And Nvl(预交类别, 2) = 2 And (主页id = n_主页id Or 主页id Is Null)
                   Group By NO, 结算方式, 交易流水号
                   Having Sum(Nvl(金额, 0)) - Sum(Nvl(冲预交, 0)) <> 0) Loop
        v_Temp := '<DJH>' || r_预交.单据号 || '</DJH>';
        v_Temp := v_Temp || '<JSFS>' || r_预交.结算方式 || '</JSFS>';
        v_Temp := v_Temp || '<JE>' || r_预交.金额 || '</JE>';
        v_Temp := v_Temp || '<JYLSH>' || r_预交.交易流水号 || '</JYLSH>';
        If n_卡类别id = r_预交.卡类别id Then
          v_Temp := v_Temp || '<SFJSK>' || 1 || '</SFJSK>';
        Else
          v_Temp := v_Temp || '<SFJSK>' || 0 || '</SFJSK>';
        End If;
        v_Temp     := '<ITEM>' || v_Temp || '</ITEM>';
        v_Subtemp  := v_Subtemp || v_Temp;
        n_病人余额 := Nvl(n_病人余额, 0) + Nvl(r_预交.金额, 0);
      End Loop;
      v_Subtemp := '<YJKLIST>' || v_Subtemp || '</YJKLIST>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Subtemp)) Into x_Templet From Dual;
      If n_病人余额 - n_结帐金额 > 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(n_病人余额 - n_结帐金额) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End If;
  End If;
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getsettlement;
/

--92818:刘硕,2016-01-18,最后一次入院更新主页ID=NUll的结构化地址
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
  d_出院日期 病案主页.出院日期%Type;
  n_Count    Number(3);
Begin
  If 功能_In = 1 Then
    Update 病人地址信息
    Set 省 = 省_In, 市 = 市_In, 县 = 县_In, 乡镇 = 乡镇_In, 其他 = 其他_In, 区划代码 = 区划代码_In
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And 地址类别 = 地址类别_In;
    If Sql%Rowcount = 0 Then
      Insert Into 病人地址信息
        (病人id, 主页id, 地址类别, 省, 市, 县, 乡镇, 其他, 区划代码)
      Values
        (病人id_In, 主页id_In, 地址类别_In, 省_In, 市_In, 县_In, 乡镇_In, 其他_In, 区划代码_In);
    End If;
    --若主页ID是病人最后一次在该院就诊，则更新主页ID=Null的数据
    If Not 主页id_In Is Null Then
      Select 出院日期 Into d_出院日期 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
      --存在出院时间，则判断该出院后是否存在就诊或住院数据
      If Not d_出院日期 Is Null Then
        --先判断住院
        Select Count(1) Into n_Count From 病案主页 Where 病人id = 病人id_In And 入院日期 >= d_出院日期;
        If n_Count = 0 Then
          Begin
            --该过程病案、标准版均有。病案系统若单独安装没有病人挂号记录
            Execute Immediate 'Select Count(1) From 病人挂号记录 Where 病人id =:1  And 登记时间 >=:2 '
              Into n_Count
              Using 病人id_In, d_出院日期;
          Exception
            When Others Then
              Null;
          End;
        End If;
      End If;
      If d_出院日期 Is Null Or Nvl(n_Count, 0) = 0 Then
        Update 病人地址信息
        Set 省 = 省_In, 市 = 市_In, 县 = 县_In, 乡镇 = 乡镇_In, 其他 = 其他_In, 区划代码 = 区划代码_In
        Where 病人id = 病人id_In And 主页id Is Null And 地址类别 = 地址类别_In;
      End If;
    End If;
  Else
    Delete From 病人地址信息
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And 地址类别 = 地址类别_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人地址信息_Update;
/

--92335:李南春,2016-01-18,三方支付新模式及过程拆分
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
  合作单位_In   病人预交记录.合作单位%Type := Null,
  更新交款余额_In  Number := 0--是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况。
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
      if Nvl(更新交款余额_In,0)=0 then
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
      End if;
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
  修正病人费别_In Number := 0,
  更新交款余额_In  Number := 0--是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
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
    If 序号_In = 1 And Nvl(更新交款余额_In,0)=0 Then
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

CREATE OR REPLACE Procedure Zl_病人预交记录_Insert
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
  结算性质_In   病人预交记录.结算性质%Type := Null,
  更新交款余额_In  Number := 0--是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况。
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
  If Nvl(更新交款余额_In,0)=0 then
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
  End if;
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

CREATE OR REPLACE Procedure zl_人员缴款余额_Update(
费用模块_In      Number,
结帐id_In        病人预交记录.结帐id%Type,
结算方式_In      病人预交记录.结算方式%Type,
现金支付_In      病人预交记录.冲预交%Type,
个帐支付_In      病人预交记录.冲预交%Type,
操作员姓名_In    病人预交记录.操作员姓名%Type
) as 
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
  n_返回值 病人余额.预交余额%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
begin
  ---自助模块更新人员交款余额
  ---费用模块：1-预交款,2-结帐补款,3-收费收款,4-挂号收款,5-就诊卡收款
 if 费用模块_In=1 or 费用模块_In=5 then
    --人员缴款余额(现收)
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 现金支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = 结算方式_In
    Returning 余额 Into n_返回值;

    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, 结算方式_In, 1, 现金支付_In);
      n_返回值 := 现金支付_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
    End If;
 elsif 费用模块_In=3 then
    For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
                 From 病人预交记录 A
                 Where a.结帐id = 结帐id_In And Mod(a.记录性质, 10) <> 1
                 Group By 结算方式, 操作员姓名) Loop

      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
      Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
      End If;
    End Loop;
 elsif 费用模块_In=4 then
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
 End if;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_人员缴款余额_Update;
/

--92776:梁经伙,2016-01-15,删除传染病报告卡时，清除报告卡关联的反馈单的文件ID
CREATE OR REPLACE Procedure Zl_电子病历记录_Delete(Id_In In 电子病历记录.Id%Type) Is
  n_处理状态 电子病历记录.处理状态%Type; 
  e_Submit Exception; 
Begin 
  Select Nvl(处理状态, 0) Into n_处理状态 From 电子病历记录 Where ID = Id_In; 
  If n_处理状态 > 0 Then 
    Raise e_Submit; 
  End If; 
  Delete 病人诊断记录 T 
  Where t.Id In (Select a.Id 
                 From 病人诊断记录 A, 电子病历记录 C 
                 Where a.病历id = c.Id And a.病人id = c.病人id And a.主页id = c.主页id And c.Id = Id_In); 
  Update 电子病历时机 
  Set 完成记录id = Null, 完成时间 = Null 
  Where (病人id, 主页id, 文件id) = (Select 病人id, 主页id, 文件id From 电子病历记录 Where ID = Id_In) And 完成记录id = Id_In; 
  update 疾病阳性记录 set 文件ID = NULL where 文件ID = Id_In;  --传染病管理系统清除关联的反馈单的文件ID
  Delete 电子病历打印 Where 文件id = Id_In; 
  Delete 电子病历记录 Where ID = Id_In; 
  Delete 疾病申报记录 Where 文件id = Id_In; --为支持新版病历，删除了外键 
Exception 
  When e_Submit Then 
    Raise_Application_Error(-20101, '[ZLSOFT]不能删除被后续接收的病历！[ZLSOFT]'); 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_电子病历记录_Delete;
/

--92527:刘尔旋,2016-01-15,重算费用病人余额问题
Create Or Replace Procedure Zl_病人未结门诊费用_Recalc(病人id_In 住院费用记录.病人id%Type) As
  v_费别     费别.名称%Type;
  v_No       门诊费用记录.NO%Type;
  n_实收金额 门诊费用记录.实收金额%Type;
  n_费用余额 病人余额.费用余额%Type;
  n_小数位数 Number(2);
  v_Counter  Number(5);
  d_Sysdate  Date;
  v_Thisinfo Varchar(100);
  v_Lastinfo Varchar(100);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  Select 费别 Into v_费别 From 病人信息 Where 病人id = 病人id_In;

  --条件判断
  --a.当前不是按主从项汇总计算折扣模式
  v_Counter := To_Number(Nvl(Zl_Getsysparameter(93), 0));
  If v_Counter = 1 Then
    v_Error := '当前费别使用主从项汇总计算折扣模式,不支持费用重算!';
    Raise Err_Custom;
  End If;

  --b.当前费别不是使用药品按成本价加收打折的费别
  v_Counter := 0;
  Select Count(费别) Into v_Counter From 费别明细 Where 费别 = v_费别 And 计算方法 = 1;
  If v_Counter > 0 Then
    v_Error := '当前费别使用药品按成本价加收打折模式,不支持费用重算!';
    Raise Err_Custom;
  End If;

  --c.没有未结费用
  Begin
    Select 费用余额 Into n_费用余额 From 病人余额 Where 病人id = 病人id_In and 类型=1 and 性质=1;
  Exception
    When Others Then
      n_费用余额 := 0;
  End;
  --可能有未结费用，但不是本次住院发生的，在后面执行时再判断本次是否有未结明细
  If n_费用余额 = 0 Then
    v_Error := '病人不存在未结费用,不用进行费用重算!';
    Raise Err_Custom;
  End If;

  --d.不存在与本次住院费别不同的费用明细
  v_Counter := 0;
  Select Count(ID) Into v_Counter From 门诊费用记录 Where 病人id = 病人id_In And 费别 <> v_费别;
  If v_Counter = 0 Then
    v_Error := '病人不存在与本次住院费别不同的费用明细 ,不用进行费用重算!';
    Raise Err_Custom;
  End If;

  --执行
  v_Counter  := 0;
  d_Sysdate  := Sysdate;
  n_小数位数 := To_Number(Nvl(Zl_Getsysparameter(9), 2));
  For r_Fee In (Select 病人id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 病人科室id, 收费类别, 收费细目id, 计算单位,
                       加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 开单部门id, 开单人, 执行部门id, 发生时间,
                       操作员编号, 操作员姓名, Nvl(Sum(应收金额), 0) 应收金额, Nvl(Sum(实收金额), 0) 实收金额
                From (Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号,
                              门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 0 As 病人病区id, 病人科室id, 费别, 收费类别,
                              收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目,
                              标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id,
                              执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id,
                              保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊
                       From 门诊费用记录
                       Union All
                       Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号,
                              门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 0 As 病人病区id, 病人科室id, 费别, 收费类别,
                              收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目,
                              标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id,
                              执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id,
                              保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊
                       From H门诊费用记录)
                Where 病人id = 病人id_In And 记录状态 <> 0 And 记帐费用 = 1
                Group By 病人id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 病人科室id, 收费类别, 收费细目id, 计算单位,
                         加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 开单部门id, 开单人, 执行部门id, 发生时间,
                         操作员编号, 操作员姓名
                Having(Nvl(Sum(实收金额), 0) <> Nvl(Sum(结帐金额), 0) Or Nvl(Sum(结帐金额), 0) = 0) And Not(Nvl(Sum(应收金额), 0) = 0 And Nvl(Sum(实收金额), 0) = 0)
                Order By 开单部门id, 开单人, 操作员姓名) Loop
    --          包括从未结的费用,费用明细部分结帐,以及结帐后作废,这些记录有可能已转入后备表
    --          1.排开了已全部结帐的记录(Sum(应收金额)=Sum(应收金额))
    --          2.排开了无打折冲减的记帐后已销帐的记录(Sum(应收金额)=0,Sum(应收金额)=0)
    --          3.不排开打折冲减后发生了单据销帐的记录，要将原冲减记录一并汇总重算(Sum(应收金额)=0,Sum(应收金额)<>0)
    --          4.不排开打折冲减后产生的实收和结帐都为零的记录，因为改回原来的费别时，要重算回去
    If r_Fee.应收金额 <> 0 Then
      Begin
        Select 实收金额
        Into n_实收金额
        From (Select Round(r_Fee.应收金额 * Nvl(实收比率, 0) / 100, n_小数位数) 实收金额
               From 费别明细
               Where 收费细目id = r_Fee.收费细目id And 费别 = v_费别 And Abs(r_Fee.应收金额) Between 应收段首值 And
                     应收段尾值 And Nvl(计算方法, 0) = 0
               Union All
               Select Round(r_Fee.应收金额 * Nvl(实收比率, 0) / 100, n_小数位数) 实收金额
               From 费别明细 A
               Where 收入项目id = r_Fee.收入项目id And 费别 = v_费别 And Abs(r_Fee.应收金额) Between 应收段首值 And
                     应收段尾值 And Nvl(计算方法, 0) = 0 And Not Exists
                (Select 1 From 费别明细 B Where B.费别 = A.费别 And B.收费细目id = r_Fee.收费细目id));
      Exception
        When Others Then
          n_实收金额 := r_Fee.应收金额;
      End;
    Else
      n_实收金额 := 0;
    End If;
    --计算用来冲减原实收的差额
    n_实收金额 := -1 * (r_Fee.实收金额 - n_实收金额);
  
    If n_实收金额 <> 0 Then
      --一张单据的开单部门id,开单人,操作员姓名,床号要求相同，如果其中之一变了则产生新单据，如果都没有变，一张单据最多100条明细
      v_Thisinfo := r_Fee.开单部门id || r_Fee.开单人 || r_Fee.操作员姓名 || ' ';
      If v_Counter = 0 Or v_Counter = 100 Or v_Thisinfo <> v_Lastinfo Then
        v_No       := Nextno(14);
        v_Counter  := 1;
        v_Lastinfo := v_Thisinfo;
      Else
        v_Counter := v_Counter + 1;
      End If;
    
      Insert Into 门诊费用记录
        (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄,
         病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 发药窗口, 加班标志,
         附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 划价人, 开单部门id,
         开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 摘要, 是否急诊,
         医嘱序号)
      Values
        (病人费用记录_Id.Nextval, 2, v_No, 1, v_Counter, Null, Null, Null, r_Fee.门诊标志, r_Fee.病人id, r_Fee.标识号,
         r_Fee.姓名, r_Fee.性别, r_Fee.年龄, r_Fee.病人科室id, v_费别, r_Fee.收费类别, r_Fee.收费细目id, r_Fee.计算单位,
         Null, Null, 0, 0, Null, r_Fee.加班标志, r_Fee.附加标志, r_Fee.婴儿费, r_Fee.收入项目id, r_Fee.收据费目, 0, 0,
         n_实收金额, Null, 1, Null, r_Fee.开单部门id, r_Fee.开单人, r_Fee.发生时间, d_Sysdate, r_Fee.执行部门id, 0, Null,
         Null, r_Fee.操作员编号, r_Fee.操作员姓名, Decode(v_Counter, 1, '实收重算冲减', ''), 0, Null);
    End If;
  End Loop;

  If v_Counter = 0 Then
    v_Error := '由于以下原因之一,没有进行费用重算:' || Chr(13) || Chr(13) || 'a.没有发现病人本次住院的未结费用.' ||
               Chr(13) || 'b.所有未结费用已进行了费用重算.' || Chr(13) || 'c.按当前费别重算的实收冲减金额都为零.';
    Raise Err_Custom;
  Else
    --病人余额
    n_实收金额 := 0;
    Select Sum(实收金额)
    Into n_实收金额
    From 门诊费用记录
    Where 病人id = 病人id_In And 记录性质 = 2 And 登记时间 = d_Sysdate;
    Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) + n_实收金额 Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1;
    If Sql%Rowcount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 费用余额, 预交余额, 类型) Values (病人id_In, 1, n_实收金额, 0, 1);
    End If;
  
    --病人未结费用
    For r_Fee In (Select Null As 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(实收金额) 实收金额
                  From 门诊费用记录
                  Where 病人id = 病人id_In And 记录性质 = 2 And 登记时间 = d_Sysdate
                  Group By 病人科室id, 开单部门id, 执行部门id, 收入项目id) Loop
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + r_Fee.实收金额
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = Nvl(r_Fee.病人病区id, 0) And
            Nvl(病人科室id, 0) = Nvl(r_Fee.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Fee.开单部门id, 0) And
            Nvl(执行部门id, 0) = Nvl(r_Fee.执行部门id, 0) And 收入项目id + 0 = Nvl(r_Fee.收入项目id, 0) And
            来源途径 + 0 = 2;
      If Sql%Rowcount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, r_Fee.病人病区id, r_Fee.病人科室id, r_Fee.开单部门id, r_Fee.执行部门id, r_Fee.收入项目id, 2,
           r_Fee.实收金额);
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人未结门诊费用_Recalc;
/

--92527:刘尔旋,2016-01-14,费别重算后病人余额问题
Create Or Replace Procedure Zl_病人未结费用_Recalc
(
  病人id_In 住院费用记录.病人id%Type,
  主页id_In 住院费用记录.主页id%Type
) As
  v_费别     费别.名称%Type;
  v_No       住院费用记录.No%Type;
  n_实收金额 住院费用记录.实收金额%Type;
  n_费用余额 病人余额.费用余额%Type;
  n_小数位数 Number(2);
  v_Counter  Number(5);
  d_Sysdate  Date;
  v_Thisinfo Varchar(100);
  v_Lastinfo Varchar(100);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  Select 费别 Into v_费别 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;

  --条件判断
  --a.当前不是按主从项汇总计算折扣模式
  v_Counter := To_Number(Nvl(zl_GetSysParameter(93), 0));
  If v_Counter = 1 Then
    v_Error := '当前费别使用主从项汇总计算折扣模式,不支持费用重算!';
    Raise Err_Custom;
  End If;

  --b.当前费别不是使用药品按成本价加收打折的费别
  v_Counter := 0;
  Select Count(费别) Into v_Counter From 费别明细 Where 费别 = v_费别 And 计算方法 = 1;
  If v_Counter > 0 Then
    v_Error := '当前费别使用药品按成本价加收打折模式,不支持费用重算!';
    Raise Err_Custom;
  End If;

  --c.没有未结费用
  Begin
    Select 费用余额 Into n_费用余额 From 病人余额 Where 病人id = 病人id_In And 类型 = 2 And 性质 = 1;
  Exception
    When Others Then
      n_费用余额 := 0;
  End;
  --可能有未结费用，但不是本次住院发生的，在后面执行时再判断本次是否有未结明细
  If n_费用余额 = 0 Then
    v_Error := '病人不存在未结费用,不用进行费用重算!';
    Raise Err_Custom;
  End If;

  --d.不存在与本次住院费别不同的费用明细
  v_Counter := 0;
  Select Count(ID) Into v_Counter From 住院费用记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 费别 <> v_费别;
  If v_Counter = 0 Then
    v_Error := '病人不存在与本次住院费别不同的费用明细 ,不用进行费用重算!';
    Raise Err_Custom;
  End If;

  --执行
  v_Counter  := 0;
  d_Sysdate  := Sysdate;
  n_小数位数 := To_Number(Nvl(zl_GetSysParameter(9), 2));
  For r_Fee In (Select 病人id, 主页id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 收费类别, 收费细目id, 计算单位, 加班标志, 附加标志, 婴儿费,
                       收入项目id, 收据费目, 开单部门id, 开单人, 执行部门id, 发生时间, 操作员编号, 操作员姓名, 医疗小组id, Nvl(Sum(应收金额), 0) 应收金额,
                       Nvl(Sum(实收金额), 0) 实收金额
                From (Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别,
                              年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id,
                              收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号,
                              操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 医疗小组id
                       From 住院费用记录
                       Union All
                       Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别,
                              年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id,
                              收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号,
                              操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 医疗小组id
                       From H住院费用记录)
                Where 病人id = 病人id_In And 主页id = 主页id_In And 记录状态 <> 0 And 记帐费用 = 1
                Group By 病人id, 主页id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 收费类别, 收费细目id, 计算单位, 加班标志, 附加标志, 婴儿费,
                         收入项目id, 收据费目, 开单部门id, 开单人, 执行部门id, 发生时间, 操作员编号, 操作员姓名, 医疗小组id
                Having(Nvl(Sum(实收金额), 0) <> Nvl(Sum(结帐金额), 0) Or Nvl(Sum(结帐金额), 0) = 0) And Not(Nvl(Sum(应收金额), 0) = 0 And Nvl(Sum(实收金额), 0) = 0)
                Order By 开单部门id, 开单人, 操作员姓名) Loop
    --          包括从未结的费用,费用明细部分结帐,以及结帐后作废,这些记录有可能已转入后备表
    --          1.排开了已全部结帐的记录(Sum(应收金额)=Sum(应收金额))
    --          2.排开了无打折冲减的记帐后已销帐的记录(Sum(应收金额)=0,Sum(应收金额)=0)
    --          3.不排开打折冲减后发生了单据销帐的记录，要将原冲减记录一并汇总重算(Sum(应收金额)=0,Sum(应收金额)<>0)
    --          4.不排开打折冲减后产生的实收和结帐都为零的记录，因为改回原来的费别时，要重算回去
    If r_Fee.应收金额 <> 0 Then
      Begin
        Select 实收金额
        Into n_实收金额
        From (Select Round(r_Fee.应收金额 * Nvl(实收比率, 0) / 100, n_小数位数) 实收金额
               From 费别明细
               Where 收费细目id = r_Fee.收费细目id And 费别 = v_费别 And Abs(r_Fee.应收金额) Between 应收段首值 And 应收段尾值 And Nvl(计算方法, 0) = 0
               Union All
               Select Round(r_Fee.应收金额 * Nvl(实收比率, 0) / 100, n_小数位数) 实收金额
               From 费别明细 A
               Where 收入项目id = r_Fee.收入项目id And 费别 = v_费别 And Abs(r_Fee.应收金额) Between 应收段首值 And 应收段尾值 And Nvl(计算方法, 0) = 0 And
                     Not Exists (Select 1 From 费别明细 B Where b.费别 = a.费别 And b.收费细目id = r_Fee.收费细目id));
      Exception
        When Others Then
          n_实收金额 := r_Fee.应收金额;
      End;
    Else
      n_实收金额 := 0;
    End If;
    --计算用来冲减原实收的差额
    n_实收金额 := -1 * (r_Fee.实收金额 - n_实收金额);
  
    If n_实收金额 <> 0 Then
      --一张单据的开单部门id,开单人,操作员姓名,床号要求相同，如果其中之一变了则产生新单据，如果都没有变，一张单据最多100条明细
      v_Thisinfo := r_Fee.开单部门id || r_Fee.开单人 || r_Fee.操作员姓名 || r_Fee.床号;
      If v_Counter = 0 Or v_Counter = 100 Or v_Thisinfo <> v_Lastinfo Then
        v_No       := Nextno(14);
        v_Counter  := 1;
        v_Lastinfo := v_Thisinfo;
      Else
        v_Counter := v_Counter + 1;
      End If;
    
      Insert Into 住院费用记录
        (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 门诊标志, 病人id, 主页id, 标识号, 床号, 姓名, 性别, 年龄, 病人病区id, 病人科室id, 费别,
         收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 发药窗口, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用,
         划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 摘要, 是否急诊, 医嘱序号, 医疗小组id)
      Values
        (病人费用记录_Id.Nextval, 2, v_No, 1, v_Counter, Null, Null, 0, Null, r_Fee.门诊标志, r_Fee.病人id, r_Fee.主页id, r_Fee.标识号,
         r_Fee.床号, r_Fee.姓名, r_Fee.性别, r_Fee.年龄, r_Fee.病人病区id, r_Fee.病人科室id, v_费别, r_Fee.收费类别, r_Fee.收费细目id, r_Fee.计算单位,
         Null, Null, 0, 0, Null, r_Fee.加班标志, r_Fee.附加标志, r_Fee.婴儿费, r_Fee.收入项目id, r_Fee.收据费目, 0, 0, n_实收金额, Null, 1,
         Null, r_Fee.开单部门id, r_Fee.开单人, r_Fee.发生时间, d_Sysdate, r_Fee.执行部门id, 0, Null, Null, r_Fee.操作员编号, r_Fee.操作员姓名,
         Decode(v_Counter, 1, '实收重算冲减', ''), 0, Null, r_Fee.医疗小组id);
    End If;
  End Loop;

  If v_Counter = 0 Then
    v_Error := '由于以下原因之一,没有进行费用重算:' || Chr(13) || Chr(13) || 'a.没有发现病人本次住院的未结费用.' || Chr(13) || 'b.所有未结费用已进行了费用重算.' ||
               Chr(13) || 'c.按当前费别重算的实收冲减金额都为零.';
    Raise Err_Custom;
  Else
    --病人余额
    n_实收金额 := 0;
    Select Sum(实收金额)
    Into n_实收金额
    From 住院费用记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 2 And 登记时间 = d_Sysdate;
    Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) + n_实收金额 Where 病人id = 病人id_In And 性质 = 1 And 类型 = 2;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 费用余额, 预交余额, 类型) Values (病人id_In, 1, n_实收金额, 0, 2);
    End If;
  
    --病人未结费用
    For r_Fee In (Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(实收金额) 实收金额
                  From 住院费用记录
                  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 2 And 登记时间 = d_Sysdate
                  Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id) Loop
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + r_Fee.实收金额
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 主页id_In And Nvl(病人病区id, 0) = r_Fee.病人病区id And
            Nvl(病人科室id, 0) = r_Fee.病人科室id And Nvl(开单部门id, 0) = r_Fee.开单部门id And Nvl(执行部门id, 0) = r_Fee.执行部门id And
            收入项目id + 0 = r_Fee.收入项目id And 来源途径 + 0 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, 主页id_In, r_Fee.病人病区id, r_Fee.病人科室id, r_Fee.开单部门id, r_Fee.执行部门id, r_Fee.收入项目id, 2, r_Fee.实收金额);
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人未结费用_Recalc;
/

--89717:余伟节,2016-01-14,出院后不允许取消完成路径
CREATE OR REPLACE Procedure Zl_病人路径结束_Delete
(
  路径记录id_In 病人临床路径.Id%Type,
  结束类型_In   病人临床路径.状态%Type
) Is
  v_阶段id     病人路径评估.阶段id%Type;
  v_前一阶段id 病人路径评估.阶段id%Type;
  v_日期       病人路径评估.日期%Type;
  v_天数       病人路径评估.天数%Type;

  v_病人id       病人临床路径.病人id%Type;
  v_主页id       病人临床路径.主页id%Type;
  d_登记时间     病人路径评估.登记时间%Type;
  d_出院日期     病案主页.出院日期%Type;
  n_当前阶段id   病人合并路径.当前阶段id%Type;
  v_是否检查出院 Varchar2(20);
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  --出院病人不允许取消完成路径
  Select zl_GetSysParameter('出院后不允许取消完成路径', 1256) Into v_是否检查出院 From Dual;
  If v_是否检查出院 = '1' Then
    Select b.出院日期
    Into d_出院日期
    From 病人临床路径 A, 病案主页 B
    Where a.病人id = b.病人id And a.主页id = b.主页id And a.Id = 路径记录id_In;
    If d_出院日期 Is Not Null Then
      If d_出院日期 <= Sysdate Then
        v_Error := '该病人已经出院,不允许取消完成路径！';
        Raise Err_Custom;
      End If;
    End If;
  End If;
  --
  Select 前一阶段id Into v_阶段id From 病人临床路径 Where ID = 路径记录id_In;
  Select Max(日期), Max(天数)
  Into v_日期, v_天数
  From 病人路径执行
  Where 路径记录id = 路径记录id_In And 阶段id = v_阶段id;

  Select 结束时间 Into d_登记时间 From 病人临床路径 Where ID = 路径记录id_In;

  --如果取消结束登记时间结束的合并路径
  For c_Merge In (Select ID
                  From 病人合并路径
                  Where 结束时间 = d_登记时间 And 首要路径记录id = 路径记录id_In And 结束时间 Is Not Null) Loop
    Select b.合并路径阶段id
    Into n_当前阶段id
    From 病人合并路径评估 B
    Where b.登记时间 = (Select Max(c.登记时间)
                    From 病人合并路径评估 C
                    Where c.路径记录id = b.路径记录id And c.合并路径记录id = b.合并路径记录id) And b.合并路径记录id = c_Merge.Id And
          b.路径记录id = 路径记录id_In;
  
    Update 病人合并路径
    Set 结束时间 = Null, 前一阶段id = n_当前阶段id, 当前阶段id = 前一阶段id
    Where 结束时间 = d_登记时间 And 首要路径记录id = 路径记录id_In And 结束时间 Is Not Null;
  End Loop;

  If 结束类型_In = 3 Then
    --评估结果为变异时自动结束的,取消结束自动取消评估
    Delete 病人路径评估 Where 路径记录id = 路径记录id_In And 阶段id = v_阶段id And 日期 = v_日期;
    Delete 病人路径指标 Where 路径记录id = 路径记录id_In And 阶段id = v_阶段id And 日期 = v_日期;
  End If;

  --b.回退到前一个阶段
  Select Max(阶段id)
  Into v_前一阶段id
  From 病人路径执行
  Where 路径记录id = 路径记录id_In And
        登记时间 = (Select Max(登记时间) From 病人路径执行 Where 路径记录id = 路径记录id_In And 阶段id <> v_阶段id);

  Update 病人临床路径
  Set 结束时间 = Null, 状态 = 1, 前一阶段id = v_前一阶段id, 当前阶段id = v_阶段id, 当前天数 = v_天数
  Where ID = 路径记录id_In
  Returning 病人id, 主页id Into v_病人id, v_主页id;

  --更新病案主页当前路径的状态
  Update 病案主页 Set 路径状态 = 1 Where 病人id = v_病人id And 主页id = v_主页id;

  Delete 病人出径记录 Where 病人id = v_病人id And 主页id = v_主页id;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人路径结束_Delete;
/

--92518:胡俊勇,2016-01-11,传染病阳性结果消息
Create Or Replace Procedure Zl_业务消息清单_Read
(
  病人id_In     In 业务消息清单.病人id%Type,
  就诊id_In     In 业务消息清单.就诊id%Type,
  类型编码_In   In 业务消息清单.类型编码%Type,
  阅读场合_In   In 业务消息状态.阅读场合%Type,
  阅读人_In     In 业务消息状态.阅读人%Type,
  阅读部门id_In In 业务消息状态.阅读部门id%Type,
  阅读时间_In   In 业务消息状态.阅读时间%Type := Null,
  消息id_In     In 业务消息状态.消息id%Type := Null
) Is
  d_Cur   Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  If 阅读时间_In Is Null Then
    Select Sysdate Into d_Cur From Dual;
  Else
    d_Cur := 阅读时间_In;
  End If;
  If Nvl(消息id_In, 0) <> 0 Then
    Insert Into 业务消息状态
      (消息id, 阅读场合, 阅读人, 阅读时间, 阅读部门id)
    Values
      (消息id_In, 阅读场合_In, 阅读人_In, d_Cur, 阅读部门id_In);
    Update 业务消息清单 Set 是否已阅 = 1 Where ID = 消息id_In;
  Else
    For R In (Select a.Id
              From 业务消息清单 A
              Where a.病人id = 病人id_In And a.就诊id = 就诊id_In And a.类型编码 = 类型编码_In And Nvl(a.是否已阅, 0) = 0) Loop
      Insert Into 业务消息状态
        (消息id, 阅读场合, 阅读人, 阅读时间, 阅读部门id)
      Values
        (r.Id, 阅读场合_In, 阅读人_In, d_Cur, 阅读部门id_In);
    End Loop;
    Update 业务消息清单
    Set 是否已阅 = 1
    Where 病人id = 病人id_In And 就诊id = 就诊id_In And 类型编码 = 类型编码_In And Nvl(是否已阅, 0) = 0;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_业务消息清单_Read;
/

--91752:刘鹏飞,2016-01-09,记录签名后新增数据在签名数据错误处理
Create Or Replace Procedure Zl_电子护理记录_Update
(
  病人id_In   In 病人护理记录.病人id%Type,
  主页id_In   In 病人护理记录.主页id%Type,
  婴儿_In     In 病人护理记录.婴儿%Type,
  开始时间_In In 病人护理记录.发生时间%Type, --本记录有效跨度的开始时间 
  结束时间_In In 病人护理记录.发生时间%Type, --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除 
  记录类型_In In 病人护理内容.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,签名记录=5,下标说明=6 
  项目序号_In In 病人护理内容.项目序号%Type, --护理项目的序号，非护理项目固定为0 
  记录标记_In In 病人护理内容.记录标记%Type, --记录内容的特殊标志 
  记录内容_In In 病人护理内容.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容 
  体温部位_In In 病人护理内容.体温部位%Type := Null,
  他人记录_In In Number := 1,
  项目首次_In In Number := 1,
  复试合格_In In Number := 0,
  是否说明_In In Number := 0, --是说明,则不填写单位 
  发生时间_In In 病人护理记录.发生时间%Type := Null, --护理记录的发生时间 
  未记说明_In In 病人护理内容.未记说明%Type := Null,	--未记说明
  操作员_IN	  IN 病人护理记录.保存人%Type:=null
) Is
  v_保存人   病人护理记录.保存人%Type;
  v_记录人   病人护理内容.记录人%Type;
  v_记录内容 病人护理内容.记录内容%Type;
  n_护理级别 病人护理记录.护理级别%Type;
  d_结束时间 病人护理记录.发生时间%Type;
  d_发生时间 病人护理记录.发生时间%Type;
  n_记录id   病人护理内容.记录id%Type;
  v_科室id   病人护理记录.科室id%Type;
  v_组号     病人护理内容.记录组号%Type;
  v_活动项目 护理记录项目.项目性质%Type;
  n_项目类型 护理记录项目.项目类型%Type;
  n_项目表示 护理记录项目.项目表示%Type;
  n_开始版本 病人护理内容.开始版本%Type;
  n_当前版本 病人护理内容.开始版本%Type;
  v_Records  Number;
  n_Add      Number;
  --主过程 

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Begin
    Select p.姓名 Into v_保存人 From 上机人员表 O, 人员表 P Where o.人员id = p.Id And 用户名 = User;
  Exception
    When Others Then
      v_保存人 := User;
  End;
  if 操作员_IN is not null then 
	v_保存人 := 操作员_IN;
  end if ;

  d_发生时间 := 发生时间_In;

  If d_发生时间 Is Null Then
    d_发生时间 := 开始时间_In;
  End If;

  If 结束时间_In Is Null Then
    d_结束时间 := 开始时间_In;
  Else
    d_结束时间 := 结束时间_In;
  End If;

  n_项目类型 := 1;
  Begin
    Select 项目类型, 项目表示, 项目性质
    Into n_项目类型, n_项目表示, v_活动项目
    From 护理记录项目
    Where 项目序号 = 项目序号_In;
  Exception
    When Others Then
      v_活动项目 := 1;
  End;
  --检查病人在本次记录时间跨度内，包含相同记录项目，但发生时间不相同的护理记录，进行清理 
  --------------------------------------------------------------------------------------------------------------------- 
  If (项目首次_In = 1) Or (记录内容_In Is Null And 未记说明_In Is Null) Then
    For r_List In (Select l.Id, Count(*) As 记录数
                   From 病人护理记录 L, 病人护理内容 D
                   Where l.Id = d.记录id And l.病人id = 病人id_In And l.主页id = 主页id_In And Nvl(l.婴儿, 0) = Nvl(婴儿_In, 0) And
                         l.病人来源 = 2 And d.终止版本 Is Null And d.项目序号 = 项目序号_In And d.记录类型 <> 5 And
                         (记录内容_In Is Null And l.发生时间 >= 开始时间_In Or 记录内容_In Is Not Null And l.发生时间 >= 开始时间_In) And
                         l.发生时间 <= d_结束时间
                   Group By l.Id) Loop
      n_当前版本 := 0;
      n_记录id   := r_List.Id;
      Begin
        Select Nvl(开始版本, 1)
        Into n_当前版本
        From 病人护理内容
        Where 记录id = n_记录id And 项目序号 = 项目序号_In And Nvl(记录标记, 0) = 记录标记_In And 记录类型 = 记录类型_In And
              Decode(v_活动项目, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And 终止版本 Is Null;
      Exception
        When Others Then
          n_当前版本 := 0;
      End;
    
      If 记录类型_In = 2 Or 记录类型_In = 6 Then
        Delete 病人护理内容
        Where 记录id = n_记录id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And 终止版本 Is Null;
      Else
        If 体温部位_In Is Not Null Then
          Delete 病人护理内容
          Where 记录id = n_记录id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And Nvl(体温部位, '无') = Nvl(体温部位_In, '无') And
                Nvl(记录标记, 0) = Nvl(记录标记_In, 0) And 终止版本 Is Null;
        Else
          Delete 病人护理内容
          Where 记录id = n_记录id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And Nvl(记录标记, 0) = Nvl(记录标记_In, 0) And 终止版本 Is Null;
        End If;
      End If;
    
      --处理版本 
      Update 病人护理内容
      Set 终止版本 = Null
      Where 终止版本 = n_当前版本 And 记录id = n_记录id And 项目序号 = 项目序号_In And 记录类型 = 记录类型_In And Nvl(记录标记, 0) = 记录标记_In And
            Decode(v_活动项目, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无');
    
      --检查是否还存在上次签名后修改的记录,如果不存在,则将签名记录的终止版本清为空 
      Begin
        Select 1
        Into v_Records
        From 病人护理内容
        Where 终止版本 = n_当前版本 And 记录id = n_记录id And 记录类型 <> 5 And Rownum < 2;
      Exception
        When Others Then
          v_Records := 0;
      End;
    
      If v_Records = 0 Then
        Update 病人护理内容 Set 终止版本 = Null Where 终止版本 = n_当前版本 And 记录类型 = 5 And 记录id = n_记录id;
      End If;
    
      Update 病人护理记录
      Set 最后版本 = 最后版本 - 1
      Where ID = n_记录id And 最后版本 Not In (Select 终止版本 From 病人护理内容 Where 记录类型 <> 5 And 记录id = n_记录id);
    
      Delete From 病人护理内容
      Where 记录id = n_记录id And 记录类型 = 5 And
            Nvl(开始版本, 1) Not In
            (Select Nvl(开始版本, 1) From 病人护理内容 A Where a.记录类型 <> 5 And a.记录id = n_记录id);
    
      Delete From 病人护理记录 A
      Where a.Id = n_记录id And Not Exists (Select 1 From 病人护理内容 B Where b.记录id = a.Id);
    End Loop;
  End If;

  If 记录内容_In Is Null And 未记说明_In Is Null Then
    Return;
  End If;
  --------------------------------------------------------------------------------------------------------------------- 
  n_记录id := 0;
  Begin
    Select ID
    Into n_记录id
    From 病人护理记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(婴儿, 0) = Nvl(婴儿_In, 0) And 病人来源 = 2 And 发生时间 = d_发生时间 And
          Rownum < 2;
  Exception
    When Others Then
      n_记录id := 0;
  End;

  n_护理级别 := Zl_Patittendgrade(病人id_In, 主页id_In, 开始时间_In);

  --------------------------------------------------------------------------------------------------------------------- 
  v_科室id := 0;
  Begin
    Select 科室id
    Into v_科室id
    From 病人变动记录
    Where 科室id Is Not Null And 病人id = 病人id_In And 主页id = 主页id_In And
          (开始时间 Between 开始时间_In And d_结束时间 Or 开始时间 <= 开始时间_In) And (开始时间_In <= 终止时间 Or 终止时间 Is Null) And Rownum < 2;
  Exception
    When Others Then
      v_科室id := 0;
  End;
  If v_科室id = 0 Then
    v_Error := '在' || To_Char(开始时间_In, 'yyyy-mm-dd hh24:mi:ss') || '至' || To_Char(d_结束时间, 'yyyy-mm-dd hh24:mi:ss') ||
               '段内无对应科室，不能操作！';
    Raise Err_Custom;
  End If;

  --确认开始版本号 
  --------------------------------------------------------------------------------------------------------------------- 
  Select Nvl(Max(Nvl(a.开始版本, 1)), 0) + 1
  Into n_开始版本
  From 病人护理内容 A, 病人护理记录 B
  Where b.病人id = 病人id_In And b.主页id = 主页id_In And Nvl(b.婴儿, 0) = Nvl(婴儿_In, 0) And b.病人来源 = 2 And b.发生时间 = d_发生时间 And
        a.记录id = b.Id And a.记录类型 = 5;

  n_当前版本 := n_开始版本;

  --检查是不是本人的记录 
  n_Add      := 1;
  v_记录人   := '';
  v_记录内容 := '';
  Begin
    Select 记录人, 记录内容, Nvl(开始版本, 1)
    Into v_记录人, v_记录内容, n_当前版本
    From 病人护理内容
    Where 记录id = n_记录id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And Nvl(记录标记, 0) = Nvl(记录标记_In, 0) And
          Decode(v_活动项目, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And 终止版本 Is Null;
  Exception
    When Others Then
      v_记录人 := '';
      n_Add    := 1;
  End;
  --------------------------------------------------------------------------------------------------------------------- 
  If 他人记录_In = 0 Then
    If v_记录人 Is Not Null And v_记录人 <> v_保存人 Then
      v_Error := '在' || To_Char(开始时间_In, 'yyyy-mm-dd hh24:mi:ss') || '至' || To_Char(d_结束时间, 'yyyy-mm-dd hh24:mi:ss') ||
                 '段内记录人不是当前人，你无权修改！';
      Raise Err_Custom;
    End If;
  End If;
  --改写病人护理记录：如果已经存在与病人、科室和发生时间相同的记录则修改，否则增加新的记录 
  --------------------------------------------------------------------------------------------------------------------- 
  If n_记录id = 0 Then
    Select 病人护理记录_Id.Nextval Into n_记录id From Dual;
    n_Add := 1;
  Else
    If n_项目类型 = 0 And n_项目表示 = 0 Then
      If n_Add = 1 And Zl_To_Number(v_记录内容) = Zl_To_Number(记录内容_In) Then
        n_Add := 0;
      End If;
    Else
      If n_Add = 1 And v_记录内容 = 记录内容_In Then
        n_Add := 0;
      End If;
    End If;
  End If;

  If n_Add = 0 And n_开始版本 > n_当前版本 And n_开始版本 > 1 Then
    n_开始版本 := n_开始版本 - 1;
  End If;

  Update 病人护理记录 Set 保存人 = v_保存人, 保存时间 = Sysdate, 最后版本 = n_开始版本 Where ID = n_记录id;

  If Sql%RowCount = 0 Then
    Insert Into 病人护理记录
      (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 护理级别, 发生时间, 保存人, 保存时间, 最后版本)
    Values
      (n_记录id, 2, 病人id_In, 主页id_In, 婴儿_In, v_科室id, n_护理级别, d_发生时间, v_保存人, Sysdate, n_开始版本);
  End If;

  --处理版本问题 
  --------------------------------------------------------------------------------------------------------------------- 
  Update 病人护理内容
  Set 终止版本 = n_开始版本, 开始版本 = Nvl(开始版本, 1)
  Where 记录id = n_记录id And 记录类型 = 5 And 终止版本 Is Null And n_Add = 1;

  If 记录类型_In = 1 Then
    Update 病人护理内容
    Set 终止版本 = n_开始版本, 开始版本 = Nvl(开始版本, 1)
    Where 记录id = n_记录id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And 终止版本 Is Null And n_Add = 1 And
          Decode(v_活动项目, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And Nvl(开始版本, 1) <> n_开始版本;
  Else
    Update 病人护理内容
    Set 终止版本 = n_开始版本, 开始版本 = Nvl(开始版本, 1)
    Where 记录id = n_记录id And 记录类型 = 记录类型_In And n_Add = 1 And
          项目名称 = Decode(记录类型_In, 2, '上标说明', 6, '下标说明', 3, '入出转', 4, 记录内容_In) And 终止版本 Is Null And
          Nvl(开始版本, 1) <> n_开始版本;
  End If;

  --删除已经登记的该区间的病人护理内容 
  --------------------------------------------------------------------------------------------------------------------- 
  If 记录类型_In = 2 Or 记录类型_In = 6 Then
    Delete 病人护理内容
    Where 记录id = n_记录id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And 终止版本 Is Null;
  Else
    Delete 病人护理内容
    Where 记录id = n_记录id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And
          Decode(v_活动项目, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And Nvl(记录标记, 0) = Nvl(记录标记_In, 0) And
          终止版本 Is Null;
  End If;

  --插入本次登记的病人护理内容 
  If 记录类型_In = 1 Then
    --如果是活动项目则根据当前记录的项目序号,取最大组号(活动项目存在不同部位的数据,需要自动更新组号以便保存多条数据) 
    v_组号 := 1;
    If v_活动项目 = 2 Then
      Begin
        Select Nvl(Max(记录组号), 0) + 1
        Into v_组号
        From 病人护理内容
        Where 记录id = n_记录id And 项目序号 = 项目序号_In;
      Exception
        When Others Then
          v_组号 := 1;
      End;
    End If;
  
    Insert Into 病人护理内容
      (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 记录人, 体温部位, 复试合格, 开始版本, 终止版本, 记录组号, 未记说明)
      Select 病人护理内容_Id.Nextval, n_记录id, 记录类型_In, 分组名, 项目id, 项目序号, 项目名称, 项目类型, 记录内容_In, Decode(是否说明_In, 1, Null, 项目单位),
             记录标记_In, v_保存人, 体温部位_In, 复试合格_In, n_开始版本, Null, v_组号, 未记说明_In
      From 护理记录项目
      Where 项目序号 = 项目序号_In;
  Else
    Insert Into 病人护理内容
      (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 记录人, 体温部位, 复试合格, 开始版本, 终止版本, 记录组号, 未记说明)
    Values
      (病人护理内容_Id.Nextval, n_记录id, 记录类型_In, Null, Null, 0, Decode(记录类型_In, 2, '上标说明', 6, '下标说明', 3, '入出转', 4, 记录内容_In),
       Decode(记录类型_In, 3, 0, 1), Decode(记录类型_In, 4, '1', 记录内容_In), '', 记录标记_In, v_保存人, 体温部位_In, 复试合格_In, n_开始版本, Null, 1,
       未记说明_In);
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子护理记录_Update;
/

--91225:胡俊勇,2016-01-11,传染病管理系统
Create Or Replace Procedure Zl_电子病历记录_Update
(
  Id_In       In 电子病历记录.Id%Type,
  病人来源_In In 电子病历记录.病人来源%Type,
  病人id_In   In 电子病历记录.病人id%Type,
  主页id_In   In 电子病历记录.主页id%Type,
  婴儿_In     In 电子病历记录.婴儿%Type,
  科室id_In   In 电子病历记录.科室id%Type,
  文件id_In   In 电子病历记录.文件id%Type,
  医嘱id_In   In 病人医嘱报告.医嘱id%Type := Null,
  创建时间_In In 电子病历记录.创建时间%Type := Null
) Is
  v_保存人     电子病历记录.保存人%Type;
  d_创建时间   电子病历记录.创建时间%Type;
  d_保存时间   电子病历记录.保存时间%Type;
  d_完成时间   电子病历记录.完成时间%Type := Null;
  n_最后版本   电子病历记录.最后版本%Type := 1;
  n_预制提纲id 电子病历内容.预制提纲id%Type;
  n_定义提纲id 电子病历内容.定义提纲id%Type;
  v_对象属性   电子病历内容.对象属性%Type;
  n_处理状态   电子病历记录.处理状态%Type;
  e_Submit Exception;
  e_Nofile Exception;
  e_Repeat Exception;

  n_种类 病历文件列表.种类%Type;
  v_名称 病历文件列表.名称%Type;
  v_事件 病历时限要求.事件%Type;
  n_唯一 病历时限要求.唯一%Type;
  n_表格 Number(1);
  n_Num  Number;
  n_Lab  Number;

  --传送病人诊断记录
  Procedure Put_Pati_Diag
  (
    v_Kind_Emr  In Varchar2,
    n_Kind_Base In 病人诊断记录.诊断类型%Type,
    n_Del_Old   In Number
  ) Is
    n_类型      病人诊断记录.诊断类型%Type;
    n_中医      Number(1); --是否中医：0-西医;1-中医
    n_疾病id    病人诊断记录.疾病id%Type; --对应疾病编码目录(ICD或中医疾病)的ID
    n_诊断id    病人诊断记录.诊断id%Type; --对应疾病诊断目录的ID
    n_证候id    病人诊断记录.证候id%Type; --对应疾病诊断目录的ID
    n_疑诊      病人诊断记录.是否疑诊%Type; --是否疑诊：0-确诊;1-疑诊
    d_日期      病人诊断记录.记录日期%Type; --诊断次序
    n_次序      病人诊断记录.诊断次序%Type; --诊断次序
    v_入院病情  病人诊断记录.入院病情%Type;
    v_出院情况  病人诊断记录.出院情况%Type;
    n_Syncpage  Number(1); --是否同步更新病案首页 0-不同步 1-同步
    n_西医order Number(2); --首页诊断次序
    n_中医order Number(2); --首页诊断次序
  Begin
    --取得是否同步更新病案首页参数
    n_Syncpage := Nvl(zl_GetSysParameter('SyncPage', 1070), 0);
  
    If n_Del_Old = 1 Then
      n_次序      := 0;
      n_西医order := 0;
      n_中医order := 0;
    Else
      Select Nvl(Max(诊断次序), 0)
      Into n_次序
      From 病人诊断记录
      Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = 1 And 病历id + 0 = Id_In And Nvl(医嘱id, 0) = Nvl(医嘱id_In, 0);
      If n_Syncpage = 1 Then
        Select Nvl(Max(诊断次序), 0)
        Into n_西医order
        From 病人诊断记录
        Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = 3;
      End If;
    End If;
  
    For r_Temp In (Select Rownum As 次序, 对象属性 As 属性, 内容文本 As 描述
                   From 临时病历内容
                   Where 对象类型 = 7 And Substr(对象属性, 1, 2) = v_Kind_Emr And Nvl(终止版, 0) = 0) Loop
      If n_Del_Old = 1 And r_Temp.次序 = 1 Then
        Delete 病人诊断记录
        Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = 1 And 病历id + 0 = Id_In And
              诊断类型 In (n_Kind_Base, n_Kind_Base + 10) And Nvl(医嘱id, 0) = Nvl(医嘱id_In, 0);
        If n_Syncpage = 1 And (n_Kind_Base = 2 Or n_Kind_Base = 3) Then
          --只处理入院诊断和出院诊断
          Delete 病人诊断记录
          Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = 3 And 诊断类型 In (n_Kind_Base, n_Kind_Base + 10);
        End If;
      End If;
      n_中医   := To_Number(Substr(r_Temp.属性, Instr(r_Temp.属性, ';', 1, 1) + 1,
                                 Instr(r_Temp.属性, ';', 1, 2) - Instr(r_Temp.属性, ';', 1, 1) - 1));
      n_疾病id := To_Number(Substr(r_Temp.属性, Instr(r_Temp.属性, ';', 1, 2) + 1,
                                 Instr(r_Temp.属性, ';', 1, 3) - Instr(r_Temp.属性, ';', 1, 2) - 1));
      n_诊断id := To_Number(Substr(r_Temp.属性, Instr(r_Temp.属性, ';', 1, 3) + 1,
                                 Instr(r_Temp.属性, ';', 1, 4) - Instr(r_Temp.属性, ';', 1, 3) - 1));
      n_证候id := To_Number(Substr(r_Temp.属性, Instr(r_Temp.属性, ';', 1, 4) + 1,
                                 Instr(r_Temp.属性, ';', 1, 5) - Instr(r_Temp.属性, ';', 1, 4) - 1));
      n_疑诊   := To_Number(Substr(r_Temp.属性, Instr(r_Temp.属性, ';', 1, 5) + 1,
                                 Instr(r_Temp.属性, ';', 1, 6) - Instr(r_Temp.属性, ';', 1, 5) - 1));
      d_日期   := To_Date(Substr(r_Temp.属性, Instr(r_Temp.属性, ';', 1, 6) + 1,
                               Instr(r_Temp.属性, ';', 1, 7) - Instr(r_Temp.属性, ';', 1, 6) - 1), 'yyyy-mm-dd hh24:mi:ss');
      If n_Kind_Base <> 1 And n_Kind_Base <> 2 And n_Kind_Base <> 3 Then
        n_中医 := 0;
      End If;
      If n_中医 = 1 Then
        n_类型 := n_Kind_Base + 10;
      Else
        n_类型 := n_Kind_Base;
      End If;
      Insert Into 病人诊断记录
        (ID, 病人id, 主页id, 医嘱id, 记录来源, 诊断次序, 病历id, 诊断类型, 疾病id, 诊断id, 证候id, 诊断描述, 是否疑诊, 记录日期, 记录人)
      Values
        (病人诊断记录_Id.Nextval, 病人id_In, 主页id_In, 医嘱id_In, 1, r_Temp.次序 + n_次序, Id_In, n_类型,
         Decode(n_疾病id, 0, Null, n_疾病id), Decode(n_诊断id, 0, Null, n_诊断id), Decode(n_证候id, 0, Null, n_证候id), r_Temp.描述,
         n_疑诊, d_日期, v_保存人);
      If n_Syncpage = 1 And (n_Kind_Base = 2 Or n_Kind_Base = 3) Then
        If n_Kind_Base = 3 Then
          v_入院病情 := Substr(r_Temp.属性, Instr(r_Temp.属性, ';', 1, 7) + 1,
                           Instr(r_Temp.属性, ';', 1, 8) - Instr(r_Temp.属性, ';', 1, 7) - 1);
          v_出院情况 := Substr(r_Temp.属性, Instr(r_Temp.属性, ';', 1, 8) + 1);
        End If;
        --如果需要同步首页诊断，只处理入院诊断和出院诊断
        If n_中医 = 1 Then
          n_中医order := n_中医order + 1;
        Else
          n_西医order := n_西医order + 1;
        End If;
        Insert Into 病人诊断记录
          (ID, 病人id, 主页id, 记录来源, 诊断次序, 编码序号, 诊断类型, 疾病id, 诊断id, 证候id, 诊断描述, 入院病情, 出院情况, 是否疑诊, 记录日期, 记录人)
        Values
          (病人诊断记录_Id.Nextval, 病人id_In, 主页id_In, 3, Decode(n_中医, 1, n_中医order, n_西医order), 1, n_类型,
           Decode(n_疾病id, 0, Null, n_疾病id), Decode(n_诊断id, 0, Null, n_诊断id), Decode(n_证候id, 0, Null, n_证候id),
           Replace(r_Temp.描述, '(?)', ''), v_入院病情, v_出院情况, n_疑诊, d_日期, v_保存人);
      End If;
    End Loop;
  End Put_Pati_Diag;

Begin
  Begin
    Select p.姓名 Into v_保存人 From 上机人员表 O, 人员表 P Where o.人员id = p.Id And 用户名 = User;
  Exception
    When Others Then
      v_保存人 := User;
  End;
  d_保存时间 := Sysdate;
  d_创建时间 := Nvl(创建时间_In, Sysdate);

  Select Greatest(Nvl(Max(开始版), 1), Nvl(Max(终止版), 1) + 1) Into n_最后版本 From 临时病历内容;
  If n_最后版本 <= 0 Then
    n_最后版本 := 1;
  End If;

  Select Count(*) Into n_Num From 病历文件列表 Where ID = 文件id_In;
  If n_Num = 0 Then
    Raise e_Nofile;
  End If;

  Select l.种类, l.名称, q.事件, q.唯一
  Into n_种类, v_名称, v_事件, n_唯一
  From 病历文件列表 L, 病历时限要求 Q
  Where l.Id = q.文件id(+) And l.Id = 文件id_In;

  Update 电子病历记录
  Set 病人来源 = 病人来源_In, 病人id = 病人id_In, 主页id = 主页id_In, 婴儿 = 婴儿_In, 科室id = 科室id_In, 文件id = 文件id_In, 保存时间 = d_保存时间
  Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Insert Into 电子病历记录
      (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 最后版本, 创建人, 创建时间, 保存人, 保存时间)
    Values
      (Id_In, 病人来源_In, 病人id_In, 主页id_In, 婴儿_In, 科室id_In, n_种类, 文件id_In, v_名称, n_最后版本, v_保存人, d_创建时间, v_保存人, d_保存时间);
    If n_种类 = 7 And Nvl(医嘱id_In, 0) <> 0 Then
      --检查报告的重复性
      Select Count(*)
      Into n_Num
      From 电子病历记录 L, 病人医嘱报告 R
      Where l.Id = r.病历id And r.医嘱id = 医嘱id_In And l.文件id = 文件id_In;
      If n_Num > 0 Then
        Raise e_Repeat;
      End If;
      --单独处理检验有多个单独下的医嘱合并为一个核收的情况
      Begin
        Select a.Id
        Into n_Lab
        From 检验标本记录 A, 病人医嘱记录 B
        Where a.医嘱id = b.相关id And Rownum <= 1 And a.医嘱id = 医嘱id_In;
      Exception
        When Others Then
          n_Lab := 0;
      End;
      If n_Lab = 0 Then
        --其他项目还是正常处理
        Insert Into 病人医嘱报告 (医嘱id, 病历id) Values (医嘱id_In, Id_In);
      Else
        --单独处理检验项目
        Insert Into 病人医嘱报告
          (医嘱id, 病历id)
          Select Distinct b.医嘱id, Id_In
          From 检验标本记录 A, 检验项目分布 B
          Where a.Id = b.标本id And a.医嘱id = 医嘱id_In And b.医嘱id Is Not Null;
      End If;
    End If;
  Else
    Select Nvl(处理状态, 0) Into n_处理状态 From 电子病历记录 Where ID = Id_In;
    Select Max(处理状态) Into n_Num From 疾病申报记录 Where 文件id = Id_In;
    If Nvl(n_Num, 0) <> 4 and Nvl(n_Num, 0) <> 5 Then
      If n_处理状态 > 0 Then
        Raise e_Submit;
      End If;
    End If;
  End If;

  Update 电子病历内容
  Set 对象序号 = -1 * 对象序号, 内容行次 = -1 * 内容行次, 终止版 = Decode(Nvl(终止版, 0), 0, n_最后版本, 终止版)
  Where 文件id = Id_In;
  For r_Temp In (Select ID, 父id, 开始版, 终止版, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 定义提纲id, 预制提纲id, 复用提纲, 使用时机,
                        诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域
                 From 临时病历内容
                 Order By ID) Loop
  
    --清理预制提纲id：对以前文件(XML或历史文件)，可能预制提纲与当前系统不符合。
    n_预制提纲id := r_Temp.预制提纲id;
    If r_Temp.对象类型 = 1 And Nvl(n_预制提纲id, 0) <> 0 Then
      Select Max(ID) Into n_预制提纲id From 病历文件结构 Where ID = n_预制提纲id And 文件id Is Null;
      If n_预制提纲id = 0 Then
        n_预制提纲id := Null;
      End If;
    End If;
    --修复定义提纲id：如果定义提纲id不存在，则根据提纲名称查找对应的定义提纲id
    n_定义提纲id := r_Temp.定义提纲id;
    If r_Temp.对象类型 = 1 Then
      If Nvl(n_定义提纲id, 0) <> 0 Then
        Select Max(ID) Into n_定义提纲id From 病历文件结构 Where ID = n_定义提纲id And 文件id = 文件id_In;
      End If;
      If Nvl(n_定义提纲id, 0) = 0 Then
        Select Max(ID)
        Into n_定义提纲id
        From 病历文件结构
        Where 文件id = 文件id_In And 内容文本 || 预制提纲id = r_Temp.内容文本 || n_预制提纲id;
      End If;
      If n_定义提纲id = 0 Then
        n_定义提纲id := Null;
      End If;
    End If;
  
    v_对象属性 := r_Temp.对象属性;
    --从签名对象获得保存人和完成时间
    If r_Temp.对象类型 = 8 Then
      If Instr(v_对象属性, ';', 1, 5) = 0 Then
        v_对象属性 := v_对象属性 || ';';
      End If;
      If Instr(v_对象属性, ';', 1, 5) - Instr(v_对象属性, ';', 1, 4) = 1 Then
        v_对象属性 := Substr(v_对象属性, 1, Instr(v_对象属性, ';', 1, 4) - 1) || ';' || To_Char(d_保存时间, 'yyyy-mm-dd hh24:mi:ss') ||
                  Substr(v_对象属性, Instr(v_对象属性, ';', 1, 5));
      End If;
      If r_Temp.开始版 >= n_最后版本 Then
        If Nvl(Instr(r_Temp.内容文本, ';'), 0) = 0 Then
          v_保存人 := r_Temp.内容文本;
        Else
          --内容文本中存放有签名人;ID,有可能签名同名所以必须使用ID,同时确保历史数据的回退正常。
          Begin
            Select 姓名 Into v_保存人 From 人员表 Where ID = Substr(r_Temp.内容文本, Instr(r_Temp.内容文本, ';') + 1);
          Exception
            When Others Then
              v_保存人 := Substr(r_Temp.内容文本, 1, Instr(r_Temp.内容文本, ';') - 1);
          End;
        End If;
      End If;
      If d_完成时间 Is Null And r_Temp.开始版 = 1 Then
        d_完成时间 := To_Date(Substr(v_对象属性, Instr(v_对象属性, ';', 1, 4) + 1,
                                 Instr(v_对象属性, ';', 1, 5) - Instr(v_对象属性, ';', 1, 4) - 1), 'yyyy-mm-dd hh24:mi:ss');
      End If;
    End If;
  
    Update 电子病历内容
    Set 父id = r_Temp.父id, 开始版 = r_Temp.开始版, 终止版 = r_Temp.终止版, 对象序号 = r_Temp.对象序号, 对象类型 = r_Temp.对象类型, 对象标记 = r_Temp.对象标记,
        保留对象 = r_Temp.保留对象, 对象属性 = v_对象属性, 内容行次 = r_Temp.内容行次, 内容文本 = r_Temp.内容文本, 是否换行 = r_Temp.是否换行, 定义提纲id = n_定义提纲id,
        预制提纲id = n_预制提纲id, 复用提纲 = r_Temp.复用提纲, 使用时机 = r_Temp.使用时机, 诊治要素id = r_Temp.诊治要素id, 替换域 = r_Temp.替换域,
        要素名称 = r_Temp.要素名称, 要素类型 = r_Temp.要素类型, 要素长度 = r_Temp.要素长度, 要素小数 = r_Temp.要素小数, 要素单位 = r_Temp.要素单位,
        要素表示 = r_Temp.要素表示, 输入形态 = r_Temp.输入形态, 要素值域 = r_Temp.要素值域
    Where ID = r_Temp.Id And 文件id + 0 = Id_In;
    If Sql%RowCount = 0 Then
      Insert Into 电子病历内容
        (ID, 文件id, 父id, 开始版, 终止版, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 定义提纲id, 预制提纲id, 复用提纲, 使用时机, 诊治要素id,
         替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域)
      Values
        (r_Temp.Id, Id_In, r_Temp.父id, r_Temp.开始版, r_Temp.终止版, r_Temp.对象序号, r_Temp.对象类型, r_Temp.对象标记, r_Temp.保留对象,
         v_对象属性, r_Temp.内容行次, r_Temp.内容文本, r_Temp.是否换行, n_定义提纲id, n_预制提纲id, r_Temp.复用提纲, r_Temp.使用时机, r_Temp.诊治要素id,
         r_Temp.替换域, r_Temp.要素名称, r_Temp.要素类型, r_Temp.要素长度, r_Temp.要素小数, r_Temp.要素单位, r_Temp.要素表示, r_Temp.输入形态,
         r_Temp.要素值域);
    Else
      --普通表格：由于编辑时没有痕迹，按单元保存；因此需要恢复子单元，保证版本记录
      If r_Temp.对象类型 = 3 Then
        n_表格 := 0;
        If Instr(v_对象属性, ';', 1, 18) = 0 Then
          n_表格 := 1;
        Elsif Substr(v_对象属性, Instr(v_对象属性, ';', 1, 18) + 1, 1) = '0' Then
          n_表格 := 1;
        End If;
        If n_表格 = 1 Then
          Update 电子病历内容
          Set 对象序号 = Abs(对象序号), 内容行次 = Abs(内容行次)
          Where 文件id = Id_In And 父id = r_Temp.Id And 开始版 <= n_最后版本 And 对象类型 <> 5;
        End If;
      End If;
    End If;
  End Loop;
  Delete 电子病历内容
  Where (Nvl(对象序号, 0) < 0 Or Nvl(内容行次, 0) < 0 Or Nvl(开始版, 1) > n_最后版本) And 文件id = Id_In;

  Update 电子病历记录
  Set 完成时间 = d_完成时间, 保存人 = v_保存人, 最后版本 = n_最后版本,
      签名级别 =
       (Select Nvl(Sum(Power(2, 要素表示 - 1)), 0)
        From (Select Distinct 要素表示 From 临时病历内容 Where 对象类型 = 8 And 开始版 >= n_最后版本))
  Where ID = Id_In;

  --先删除原有诊断，因为有可能原有诊断被删除或更改
  Delete 病人诊断记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 病历id + 0 = Id_In;
  --填写病人诊断记录
  If n_种类 = 1 Then
    Put_Pati_Diag('11', 1, 1);
  Elsif n_种类 = 2 And (v_事件 = '入院' Or v_事件 = '首次入院' Or v_事件 = '再次入院') And n_唯一 = 1 Then
    Put_Pati_Diag('21', 2, 1);
    Put_Pati_Diag('22', 2, 1);
    Put_Pati_Diag('23', 2, 1);
    Put_Pati_Diag('24', 2, 0);
  Elsif n_种类 = 2 And (v_事件 = '24小时出院' Or v_事件 = '24小时死亡') Then
    Put_Pati_Diag('21', 2, 1);
    Put_Pati_Diag('22', 2, 1);
    Put_Pati_Diag('23', 2, 1);
    Put_Pati_Diag('24', 2, 0);
    Put_Pati_Diag('31', 3, 1);
  Elsif n_种类 = 2 And (v_事件 = '出院' Or v_事件 = '死亡') Then
    Put_Pati_Diag('31', 3, 1);
  Elsif n_种类 = 2 And v_事件 = '手术' Then
    Put_Pati_Diag('41', 8, 1);
    Put_Pati_Diag('42', 9, 1);
  Elsif n_种类 = 7 And (医嘱id_In Is Not Null) Then
    Put_Pati_Diag('51', 6, 1);
    Put_Pati_Diag('52', 22, 1);
    --只处理阳性标志
    --Update 病人医嘱发送 Set 结果阳性 = 0 Where 医嘱id = 医嘱id_In;
    Update 病人医嘱发送
    Set 结果阳性 = 1
    Where 医嘱id = 医嘱id_In And Exists
     (Select 内容文本
           From 临时病历内容
           Where 对象类型 = 7 And (Substr(对象属性, 1, 2) = '51' Or Substr(对象属性, 1, 2) = '52') And Nvl(终止版, 0) = 0);
  End If;

  --处理电子病历时机
  If d_完成时间 Is Null Then
    Update 电子病历时机 Set 完成时间 = Null Where 病人id = 病人id_In And 主页id = 主页id_In And 完成记录id = Id_In;
    If Sql%RowCount = 0 Then
      Zl_电子病历时机_Update(病人id_In, 主页id_In, 病人来源_In, 科室id_In, 文件id_In, Id_In, Null, v_保存人);
    End If;
  Else
    Zl_电子病历时机_Update(病人id_In, 主页id_In, 病人来源_In, 科室id_In, 文件id_In, Id_In, d_完成时间, v_保存人);
  End If;
Exception
  When e_Submit Then
    Raise_Application_Error(-20101, '[ZLSOFT]不能更改被后续接收的病历！[ZLSOFT]');
  When e_Nofile Then
    Raise_Application_Error(-20101, '[ZLSOFT]病历文件定义丢失，请联系系统管理员！[ZLSOFT]');
  When e_Repeat Then
    Raise_Application_Error(-20101, '[ZLSOFT]其他人已经书写并保存了报告，不能再保存！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子病历记录_Update;
/

--91844:陈刘,2016-01-07,体温单没有护理明细,护理文件可以正确删除
CREATE OR REPLACE Procedure ZL_病人护理文件_DELETE(
	ID_IN IN 病人护理文件.ID%Type 
) 
IS 
	ERR_ITEM Exception; 
	V_ERR_MSG  VARCHAR2(500); 
	LNGSIGNED NUMBER ; 
Begin 
	--如果有数据则不允许删除 
	Begin 
		SELECT 1 INTO LNGSIGNED 
		FROM 病人护理数据 A,病人护理文件 B,病人护理明细 C
		WHERE B.ID=ID_IN And A.文件ID=B.ID And C.记录ID = a.id And RowNum<2; 
	Exception 
		When Others Then LNGSIGNED:=0; 
	End ; 
 
	IF LNGSIGNED=1 THEN 
		V_ERR_MSG := '该文件已经产生护理数据不允许删除,请检查！'; 
		RAISE ERR_ITEM; 
	End IF ; 
 
	--删除打印解析 
	DELETE 病人护理打印 WHERE 文件ID=ID_IN; 
	--删除明细数据 
	DELETE 病人护理明细 WHERE 记录ID IN (SELECT ID FROM 病人护理数据 WHERE 文件ID=ID_IN); 
	--删除行记录 
	DELETE 病人护理数据 WHERE 文件ID=ID_IN; 
	--删除护理文件 
	DELETE 病人护理文件 WHERE ID=ID_IN; 
	--将上级数据的续打ID设置为空 
	UPDATE 病人护理文件 SET 续打ID=NULL WHERE 续打ID=ID_IN; 
Exception 
	WHEN ERR_ITEM THEN 
		RAISE_APPLICATION_ERROR(-20101, '[ZLSOFT]' || V_ERR_MSG || '[ZLSOFT]'); 
	When Others Then 
		ZL_ERRORCENTER (SQLCODE, SQLERRM); 
End ZL_病人护理文件_DELETE;
/

--92469:蔡青松,2016-01-07,去掉一个参数 送检人_In
Create Or Replace Procedure Zl_病人医嘱发送_Sampleinput
(
  医嘱id      In Varchar2,
  接收人_In   In 病人医嘱发送.接收人%Type := Null,
  接收批次_In In 病人医嘱发送.接收批次%Type := 0,
  人员编号_In In 人员表.编号%Type := Null,
  人员姓名_In In 人员表.姓名%Type := Null
) Is
  --未审核的费用行(不包含药品)
  Cursor c_Verify(v_医嘱id In Number) Is
    Select Distinct 记录性质, NO, 序号
    From 住院费用记录
    Where 收费类别 Not In ('5', '6', '7') And 医嘱序号 + 0 = v_医嘱id And 记帐费用 = 1 And 记录状态 = 0 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id = v_医嘱id
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id)))
    Union All
    Select Distinct 记录性质, NO, 序号
    From 门诊费用记录
    Where 收费类别 Not In ('5', '6', '7') And 医嘱序号 + 0 = v_医嘱id And 记帐费用 = 1 And 记录状态 = 0 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id = v_医嘱id
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id)))
    Order By 记录性质, NO, 序号;

  --查找当前标本的相关申请
  Cursor c_Samplequest(v_医嘱id In Number) Is
    Select Distinct ID As 医嘱id, 病人来源 From 病人医嘱记录 Where v_医嘱id In (ID, 相关id);

  v_执行 Number(1);
  v_No   病人医嘱发送.No%Type;
  v_性质 病人医嘱发送.记录性质%Type;
  v_序号 Varchar2(1000);

  v_医嘱id   病人医嘱发送.医嘱id%Type;
  v_相关id   病人医嘱记录.相关id%Type;
  v_费用性质 病人医嘱发送.记录性质%Type;
  v_样本条码 病人医嘱发送.样本条码%Type;
  v_Records  Varchar2(2000);
  v_Currrec  Varchar2(50);
  v_Fields   Varchar2(50);
  v_Count    Number(18);
  v_病人id   病人医嘱记录.病人id%Type;
  v_主页id   病人医嘱记录.主页id%Type;
  v_是否出院 Number; --0=出院,1=在院
  v_记录状态 Number;
  v_病人来源 病人医嘱记录.病人来源%Type;
  v_Date     Date;
  Err_Custom Exception;
  v_Error Varchar2(100);
Begin
  Select Sysdate Into v_Date From Dual;
  --执行后自动审核对应的记帐划价单(不包含药品)
  Select Zl_To_Number(Nvl(Zl_Getsysparameter(81), '0')) Into v_执行 From Dual;

  v_Records := 医嘱id || '|';

  While v_Records Is Not Null Loop
  
    v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields  := v_Currrec;
    v_医嘱id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    v_相关id  := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    If 接收人_In Is Null Then
      Update 病人医嘱发送
      Set 接收人 = Null, 接收时间 = Null, 接收批次 = Null
      Where 医嘱id In (v_医嘱id, v_相关id);
      Update 病人医嘱发送
      Set 执行状态 = Decode(样本条码, Null, 0, 1)
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID In (v_医嘱id, v_相关id) And 相关id Is Null);
      For r_Samplequest In c_Samplequest(v_相关id) Loop
        If r_Samplequest.病人来源 = 2 Then
          Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
          Into v_费用性质
          From 病人医嘱发送
          Where 医嘱id = r_Samplequest.医嘱id;
        Else
          v_费用性质 := 1;
        End If;
        If v_费用性质 = 2 Then
          --2.费用执行处理
          Update 住院费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = 接收人_In
          Where 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Samplequest.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null)
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Null) And 采样人 Is Null);
        Else
          Update 门诊费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = 接收人_In
          Where 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Samplequest.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null)
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Null) And 采样人 Is Null);
        End If;
      End Loop;
    Else
      --判断是否已出院，如果已出院负不完成登记
      Begin
        If v_主页id Is Null Then
          Select a.病人id, a.主页id, a.病人来源
          Into v_病人id, v_主页id, v_病人来源
          From 病人医嘱记录 A, 病案主页 B
          Where a.病人id = b.病人id And a.主页id = b.主页id(+) And a.Id = v_医嘱id;
        End If;
      Exception
        When Others Then
          v_病人来源 := 1;
      End;
      If v_病人来源 = 2 Then
        If Nvl(v_主页id, 0) > 0 Then
          Select Decode(出院日期, Null, 1, 0)
          Into v_是否出院
          From 病案主页
          Where 病人id = v_病人id And 主页id = v_主页id;
        Else
          v_是否出院 := 0;
        End If;
      
        If v_是否出院 = 0 Then
          --出院的才处理
          Begin
            Select Nvl(记录状态, 0)
            Into v_记录状态
            From 住院费用记录
            Where 医嘱序号 = v_医嘱id And Nvl(记录状态, 0) = 0 And Rownum = 1;
          Exception
            When Others Then
              v_记录状态 := 1;
          End;
        
          Select Nvl(样本条码, 0) Into v_样本条码 From 病人医嘱发送 Where 医嘱id = v_医嘱id;
          If v_样本条码 = 0 Then
            v_Error := '病人已出院不能完成登记!';
            Raise Err_Custom;
          End If;
        
        End If;
      End If;
    
      Update 病人医嘱发送
      Set 接收人 = 接收人_In, 接收时间 = v_Date, 接收批次 = 接收批次_In,  重采标本 = Null
      Where 医嘱id In (v_医嘱id, v_相关id);
      Update 病人医嘱发送
      Set 执行状态 = 1
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID In (v_医嘱id, v_相关id) And 相关id Is Null);
      --记帐划价单是否转为记帐单
      --2.检查当前标本相关的申请的相关标本是否完成审核
      For r_Samplequest In c_Samplequest(v_相关id) Loop
        v_Count := 0;
        --r_SampleQuest.医嘱id申请已经完成,处理后续环节
        If v_Count = 0 Then
          If r_Samplequest.病人来源 = 2 Then
            Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
            Into v_费用性质
            From 病人医嘱发送
            Where 医嘱id = r_Samplequest.医嘱id;
          Else
            v_费用性质 := 1;
          End If;
          If v_费用性质 = 2 Then
            --2.费用执行处理
            Update 住院费用记录
            Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
            Where 收费类别 Not In ('5', '6', '7') And
                  (医嘱序号, 记录性质, NO) In
                  (Select 医嘱id, 记录性质, NO
                   From 病人医嘱附费
                   Where 医嘱id = r_Samplequest.医嘱id
                   Union All
                   Select 医嘱id, 记录性质, NO
                   From 病人医嘱发送
                   Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null)
                   Union All
                   Select 医嘱id, 记录性质, NO
                   From 病人医嘱发送
                   Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Null) And 采样人 Is Null);
          Else
            Update 门诊费用记录
            Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
            Where 收费类别 Not In ('5', '6', '7') And
                  (医嘱序号, 记录性质, NO) In
                  (Select 医嘱id, 记录性质, NO
                   From 病人医嘱附费
                   Where 医嘱id = r_Samplequest.医嘱id
                   Union All
                   Select 医嘱id, 记录性质, NO
                   From 病人医嘱发送
                   Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null)
                   Union All
                   Select 医嘱id, 记录性质, NO
                   From 病人医嘱发送
                   Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Null) And 采样人 Is Null);
          End If;
          --3.自动审核记帐
          If v_执行 = 1 Then
            For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
              If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
                If v_序号 Is Not Null Then
                  If v_费用性质 = 1 Then
                    Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
                  Elsif v_费用性质 = 2 Then
                    Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
                  End If;
                End If;
                v_序号 := Null;
              End If;
              v_No   := r_Verify.No;
              v_性质 := r_Verify.记录性质;
              v_序号 := v_序号 || ',' || r_Verify.序号;
            End Loop;
            If v_序号 Is Not Null Then
              If v_费用性质 = 1 Then
                Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
              Elsif v_费用性质 = 2 Then
                Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
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
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人医嘱发送_Sampleinput;
/

--92410:刘尔旋,2016-01-05,费用状态调整
Create Or Replace Procedure Zl_门诊转住院_三方卡结算
(
  No_In         住院费用记录.No%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  退费时间_In   住院费用记录.发生时间%Type,
  门诊退费_In   Number := 0,
  入院科室id_In 住院费用记录.开单部门id%Type := Null,
  主页id_In     住院费用记录.主页id%Type := Null,
  三方退费_In   Number := 0,
  结帐id_In     病人预交记录.结帐id%Type := Null
) As
  v_结帐ids    Varchar2(3000);
  n_组id       财务缴款分组.Id%Type;
  n_退现       Number;
  v_预交no     病人预交记录.No%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  v_Nos        Varchar2(3000);
  v_Info       Varchar2(5000);
  v_当前结算   Varchar2(3000);
  v_原结帐ids  Varchar2(5000);
  n_Tempid     病人预交记录.Id%Type;
  v_流水号     病人预交记录.交易流水号%Type;
  v_说明       病人预交记录.交易说明%Type;
  n_预交id     病人预交记录.Id%Type;
  n_原预交id   病人预交记录.Id%Type;
  n_病人id     病人信息.病人id%Type;
  n_原结帐id   病人预交记录.结帐id%Type;
  n_冲销金额   病人预交记录.冲预交%Type;
  n_卡序号     病人预交记录.卡类别id%Type;
  n_三方卡     Number;
  n_返回值     人员缴款余额.余额%Type;
  v_结算方式   病人预交记录.结算方式%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_卡号       病人预交记录.卡号%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  v_交易说明   病人预交记录.交易说明%Type;
  n_原样退     Number;
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  Procedure Zl_Square_Update
  (
    结帐ids_In    Varchar2,
    现结帐id_In   病人预交记录.结帐id%Type,
    缴款组id_In   病人预交记录.缴款组id%Type,
    退款时间_In   病人预交记录.收款时间%Type,
    结算序号_In   病人预交记录.结算序号%Type,
    结算内容_In   Varchar2 := Null,
    退费金额_In   病人预交记录.冲预交%Type := Null,
    结算卡序号_In 病人预交记录.结算卡序号%Type := Null
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
    For v_校对 In (Select Min(a.Id) As 预交id, c.消费卡id, Sum(c.结算金额) As 结算金额, c.接口编号, c.卡号, Min(c.序号) As 序号, Min(c.Id) As ID
                 From 病人预交记录 A, 病人卡结算对照 B, 病人卡结算记录 C
                 Where a.Id = b.预交id And a.结算卡序号 = 结算卡序号_In And b.卡结算id = c.Id And a.记录性质 = 3 And
                       Instr(Nvl(结算内容_In, '_LXH'), ',' || a.结算方式 || ',') = 0 And
                       a.结帐id In (Select Column_Value From Table(f_Str2list(结帐ids_In)))
                 Group By c.消费卡id, c.接口编号, c.卡号) Loop
    
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
                 -1 * 退费金额_In, 现结帐id_In, 缴款组id_In, 预交类别, 卡类别id, Nvl(结算卡序号, v_校对.接口编号), 卡号, 交易流水号, 交易说明, 合作单位, 2, 结算序号_In,
                 结算性质
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
        Update 消费卡目录 Set 余额 = Nvl(余额, 0) + 退费金额_In Where ID = Nvl(v_校对.消费卡id, 0);
      End If;
    
      Select 病人卡结算记录_Id.Nextval Into n_Id From Dual;
      Insert Into 病人卡结算记录
        (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
        Select n_Id, 接口编号, 消费卡id, 序号, n_记录状态, 结算方式, -1 * 退费金额_In, 卡号, 交易流水号, 交易时间, 备注,
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
  n_组id := Zl_Get组id(操作员姓名_In);
  If 结帐id_In Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  Else
    n_结帐id := 结帐id_In;
  End If;

  Select 结帐id, 病人id
  Into n_原结帐id, n_病人id
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Rownum < 2;

  For r_结账id In (Select Distinct 结帐id
                 From 门诊费用记录
                 Where NO In (Select Distinct NO
                              From 门诊费用记录
                              Where 结帐id In (Select 结帐id
                                             From 病人预交记录
                                             Where 结算序号 In (Select b.结算序号
                                                            From 门诊费用记录 A, 病人预交记录 B
                                                            Where a.No = No_In And b.结算序号 < 0 And Mod(a.记录性质, 10) = 1 And
                                                                  a.记录状态 <> 0 And a.结帐id = b.结帐id))) And
                       Mod(记录性质, 10) = 1 And 记录状态 <> 0
                 Union
                 Select Distinct 结帐id
                 From 门诊费用记录
                 Where NO In (Select Distinct NO
                              From 门诊费用记录
                              Where 结帐id In (Select a.结帐id
                                             From 门诊费用记录 A, 病人预交记录 B
                                             Where a.No = No_In And b.结算序号 > 0 And Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And
                                                   a.结帐id = b.结帐id))) Loop
    v_原结帐ids := v_原结帐ids || ',' || r_结账id.结帐id;
  End Loop;
  v_原结帐ids := Substr(v_原结帐ids, 2);

  Begin
    Select 摘要
    Into v_Info
    From 病人预交记录
    Where 结算方式 Is Null And 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id;
  Exception
    When Others Then
      v_Info := '';
  End;
  --处理卡结算信息
  If v_Info Is Not Null Then
    While v_Info Is Not Null Loop
      v_当前结算 := Substr(v_Info, 1, Instr(v_Info, '|') - 1);
      n_三方卡   := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
      n_卡序号   := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
      n_冲销金额 := -1 * To_Number(v_当前结算);
    
      If n_三方卡 = 0 Then
        --消费卡
        Select 结算方式
        Into v_结算方式
        From 病人预交记录
        Where 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And 结算卡序号 = n_卡序号 And Rownum < 2;
        Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, n_冲销金额, n_卡序号);
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) - n_冲销金额
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
        Returning 余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, v_结算方式, 1, -1 * n_冲销金额);
          n_返回值 := n_冲销金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
        End If;
      Else
        --结算卡
        Select 结算方式, 卡类别id, 卡号, 交易流水号, 交易说明
        Into v_结算方式, n_卡类别id, v_卡号, v_交易流水号, v_交易说明
        From 病人预交记录
        Where 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And 卡类别id = n_卡序号 And Rownum < 2;
        If Nvl(门诊退费_In, 0) = 1 Then
          If 三方退费_In = 0 Then
            v_Err_Msg := '存在无法退现的三方账户,无法进行退费!';
            Raise Err_Item;
          End If;
          Update 病人预交记录
          Set 冲预交 = 冲预交 - n_冲销金额
          Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
               缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
            Values
              (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲销金额, v_结算方式, Null, 退费时间_In,
               Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, n_卡类别id, Null, v_卡号, v_交易流水号, v_交易说明, Null, n_结帐id,
               -1 * n_结帐id, 0, 3);
          End If;
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) - n_冲销金额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
          Returning 余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, v_结算方式, 1, -1 * n_冲销金额);
            n_返回值 := -1 * n_冲销金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
          End If;
        Else
          Begin
            Select 1 Into n_退现 From 医疗卡类别 Where ID = n_卡类别id And 是否退现 = 1;
          Exception
            When Others Then
              n_退现 := 0;
          End;
        
          If 三方退费_In = 1 Or n_退现 = 0 Then
            v_结算方式 := v_结算方式;
            n_原样退   := 1;
          Else
            n_原样退 := 0;
            Begin
              Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
            Exception
              When Others Then
                Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
            End;
          End If;
        
          If 三方退费_In = 0 Then
            If n_原样退 = 1 Then
              Select 交易流水号, 交易说明, ID
              Into v_流水号, v_说明, n_原预交id
              From 病人预交记录
              Where 结帐id = n_原结帐id And 结算方式 = v_结算方式 And Rownum < 2;
            
              Update 病人预交记录
              Set 冲预交 = 冲预交 - n_冲销金额
              Where 记录性质 = 3 And 记录状态 = 2 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式 And 结帐id = n_结帐id;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲销金额, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, n_卡类别id, Null, v_卡号, v_交易流水号, v_交易说明, Null, n_结帐id,
                   -1 * n_结帐id, 0, 3);
              End If;
            
              Update 病人预交记录
              Set 金额 = 金额 + n_冲销金额
              Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                v_预交no := Nextno(11);
                Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位)
                Values
                  (n_预交id, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, n_冲销金额, v_结算方式, Null, 退费时间_In, Null, Null,
                   Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, 2, n_卡类别id, Null, v_卡号, v_流水号, v_说明, Null);
                Update 三方结算交易 Set 交易id = n_预交id Where 交易id = n_原预交id;
              End If;
            Else
              Update 病人预交记录
              Set 冲预交 = 冲预交 - n_冲销金额
              Where 记录性质 = 3 And 记录状态 = 2 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式 And 结帐id = n_结帐id;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲销金额, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id,
                   -1 * n_结帐id, 0, 3);
              End If;
            
              Update 病人预交记录
              Set 金额 = 金额 + n_冲销金额
              Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                v_预交no := Nextno(11);
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 预交类别)
                Values
                  (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, n_冲销金额, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, 2);
              End If;
            End If;
          
            --病人余额
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + n_冲销金额
            Where 性质 = 1 And 病人id = n_病人id And 类型 = 2
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, 2, n_冲销金额, 0);
              n_返回值 := n_冲销金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End If;
          --4.2缴款数据处理
          --   因为没有实际收病人的钱,所以不处理
          --部分退费情况，退原预交记录
          If 三方退费_In = 1 Then
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) - n_冲销金额
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
            Returning 余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 人员缴款余额
                (收款员, 结算方式, 性质, 余额)
              Values
                (操作员姓名_In, v_结算方式, 1, -1 * n_冲销金额);
              n_返回值 := -1 * n_冲销金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 人员缴款余额
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
            End If;
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * n_冲销金额)
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲销金额, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, n_卡类别id, Null, v_卡号, v_交易流水号, v_交易说明, Null, n_结帐id,
                 -1 * n_结帐id, 0, 3);
            End If;
          End If;
        End If;
      End If;
      v_Info := Substr(v_Info, Instr(v_Info, '|') + 1);
    End Loop;
  End If;

  Delete From 病人预交记录 Where 结帐id = n_结帐id And 记录状态 = 2 And 结算方式 Is Null;
  Update 门诊费用记录 Set 费用状态 = 0 Where 结帐id = n_结帐id;
  Update 门诊费用记录 Set 费用状态 = 0 Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 2;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_三方卡结算;
/

--91225:梁经伙,2016-01-04,传染病管理系统在疾病申报记录表中添加了字段
CREATE OR REPLACE Procedure Zl_疾病申报记录_Incept
(
  文件id_In     In 疾病申报记录.文件id%Type,
  Incept_In     In Number, --接收还是拒绝
  说明_In       In 疾病申报记录.收拒说明%Type,
  文档id_In     In Varchar2,
  病人id_In     In 疾病申报记录.病人ID%Type,
  主页ID_In     In 疾病申报记录.主页ID%Type,
  病人来源_In   In 疾病申报记录.病人来源%Type,
  Emrcontent_In In Varchar2  --新病历诊断串
) Is
  v_收拒人   人员表.姓名%Type;

  v_姓名      疾病申报记录.姓名%Type;
  v_性别      疾病申报记录.性别%Type;
  v_年龄      疾病申报记录.年龄%Type;
  v_职业      疾病申报记录.职业%Type;
  v_家庭地址  疾病申报记录.家庭地址%Type;
  v_家庭电话  疾病申报记录.家庭电话%Type;
  v_发病日期  疾病申报记录.发病日期%Type;
  v_确诊日期  疾病申报记录.确诊日期%Type;
  v_诊断描述1 疾病申报记录.诊断描述1%Type;
  v_诊断描述2 疾病申报记录.诊断描述2%Type;
  v_填报备注  疾病申报记录.填报备注%Type;
  v_内容文本  电子病历内容.内容文本%Type;
  v_报卡类型  疾病申报记录.报卡类型%Type;
  v_报告医生  疾病申报记录.报告医生%Type;

  v_Count Number;
  e_Changed Exception;

  Function Trimlen
  (
    Str_In Varchar2,
    Len_In Number
  ) Return Varchar2 Is
    v_Temp Varchar2(4000);
  Begin
    If Str_In Is Not Null Then
      For I In 1 .. Length(Str_In) Loop
        If Lengthb(v_Temp || Substr(Str_In, I, 1)) <= Len_In Then
          v_Temp := v_Temp || Substr(Str_In, I, 1);
        Else
          Exit;
        End If;
      End Loop;
    End If;
    Return v_Temp;
  End Trimlen;
Begin

  Select 姓名 Into v_收拒人 From 人员表 P, 上机人员表 U Where p.Id = u.人员id And u.用户名 = User And Rownum < 2;

  If Length(文档id_In) <> 32 Then
    --新病历ID是32位GUID
    Update 电子病历记录 Set 处理状态 = Decode(Incept_In, 1, 1, -1) Where ID = 文件id_In;
    If Sql%RowCount = 0 Then
      Raise e_Changed;
    End If;
  End If;

  --自动提取申报病历中的项目内容
  If Incept_In = 1 Then
    If Length(文档id_In) <> 32 Then
      --固定对应要素
      v_Count := 0;
      For r_Item In (Select 要素名称, 要素类型, 内容行次, 内容文本
                     From 电子病历内容
                     Where (对象类型 = 4 or 对象类型 = 8 )And 文件id = 文件id_In
                     Order By 对象序号, 内容行次) Loop

        If r_Item.要素名称 = '姓名' Then
          v_姓名 := Trimlen(r_Item.内容文本, 20);
        Elsif r_Item.要素名称 = '性别' Then
          v_性别 := Trimlen(r_Item.内容文本, 4);
        Elsif r_Item.要素名称 = '年龄' Then
          v_年龄 := Trimlen(r_Item.内容文本, 10);
        Elsif r_Item.要素名称 = '职业'  Then
          v_职业 := Trimlen(r_Item.内容文本, 80);
        Elsif r_Item.要素名称 = '家庭地址' Then
          v_家庭地址 := Trimlen(r_Item.内容文本, 100);
        Elsif r_Item.要素名称 = '家庭电话'  Then
          v_家庭电话 := Trimlen(r_Item.内容文本, 20);
        Elsif r_Item.要素名称 = '当前日期'  Then
          v_Count := v_Count + 1;
          If v_Count = 1  Then
            --病历中第1个"当前日期"作为发病日期
            Begin
              v_发病日期 := To_Date(Replace(Replace(Replace(r_Item.内容文本, '年', '-'), '月', '-'), '日', ''), 'YYYY-MM-DD');
            Exception
              When Others Then
                Null;
            End;
          Elsif v_Count = 2  Then
            --病历中第2个"当前日期"作为确诊日期
            Begin
              v_确诊日期 := To_Date(Replace(Replace(Replace(r_Item.内容文本, '年', '-'), '月', '-'), '日', ''), 'YYYY-MM-DD');
            Exception
              When Others Then
                Null;
            End;
          End If;
        Elsif r_Item.要素名称 = '常见传染病' Then
          v_诊断描述1 := Trimlen(r_Item.内容文本, 150);
        End If;
      End Loop;

        --其他临时要素对应
      For r_Item In (Select 申报项目, 对应要素 From 疾病申报对应) Loop
        Begin
          Select 内容文本
          Into v_内容文本
          From 电子病历内容
          Where 对象类型 = 4 And 诊治要素id Is Null And 要素名称 = r_Item.对应要素 And 文件id = 文件id_In;
        Exception
          When Others Then
            v_内容文本 := Null;
        End;

        If r_Item.申报项目 = '诊断描述2' Then
          v_诊断描述2 := Trimlen(v_内容文本, 150);
        Elsif r_Item.申报项目 = '填报备注' Then
          v_填报备注 := Trimlen(v_内容文本, 100);
        End If;
      End Loop;
    Else
      Select 姓名, 性别, 年龄, 职业, 家庭地址, 家庭电话, 家庭电话
      Into v_姓名, v_性别, v_年龄, v_职业, v_家庭地址, v_家庭电话, v_家庭电话
      From 病人信息
      Where 病人id = 病人id_In;
      v_发病日期  := '';
      v_确诊日期  := '';
      v_诊断描述1 := Substr(Emrcontent_In, 1, Instr(Emrcontent_In, '|') - 1);
      v_诊断描述2 := '';
      v_填报备注  := Substr(Emrcontent_In, Instr(Emrcontent_In, '|') + 1,Instr(Emrcontent_In, '|',1,2)-1-Instr(Emrcontent_In, '|'));
      v_报卡类型  := '1 初次报告';
      v_报告医生  := Substr(Emrcontent_In, Instr(Emrcontent_In, '|',1,2) + 1);
    End If;
  End If;

  --接收数据
  Update 疾病申报记录
  Set 处理状态 = Decode(Incept_In, 1, 1, -1), 收拒人 = v_收拒人, 收拒时间 = Sysdate, 收拒说明 = 说明_In, 姓名 = v_姓名, 性别 = v_性别, 年龄 = v_年龄,
      职业 = v_职业, 家庭地址 = v_家庭地址, 家庭电话 = v_家庭电话, 发病日期 = v_发病日期, 确诊日期 = v_确诊日期, 诊断描述1 = v_诊断描述1, 诊断描述2 = v_诊断描述1,
      填报备注 = v_填报备注, 报告医生 = v_报告医生,报卡类型 = v_报卡类型,病人id= 病人id_In,主页ID = 主页ID_In,病人来源 = 病人来源_In
  Where 文件id = 文件id_In;
  If Sql%RowCount = 0 Then
    Insert Into 疾病申报记录
      (文件id, 处理状态, 收拒人, 收拒时间, 收拒说明, 姓名, 性别, 年龄, 职业, 家庭地址, 家庭电话, 发病日期, 确诊日期, 诊断描述1, 诊断描述2, 填报备注, 文档id, 报告医生, 报卡类型,病人id,主页ID,病人来源)
    Values
      (文件id_In, Decode(Incept_In, 1, 1, -1), v_收拒人, Sysdate, 说明_In, v_姓名, v_性别, v_年龄, v_职业, v_家庭地址, v_家庭电话, v_发病日期,
       v_确诊日期, v_诊断描述1, v_诊断描述2, v_填报备注, 文档id_In, v_报告医生, v_报卡类型,病人id_In,主页ID_In,病人来源_In);
  End If;
 
Exception
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]用户身份不明确！[ZLSOFT]');
  When e_Changed Then
    Raise_Application_Error(-20101, '[ZLSOFT]疾病报告已经被其他用户改变！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病申报记录_Incept;
/

--92258:许华峰,2015-12-31,缩略图插件中显示报告图象
--92392:许华峰,2016-01-05,将缩略图插件中的sql查询语句写入包中，同一管理
--影像报告插件管理(---定义部分---)***************************************************
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

  --功    能：根据医嘱ID获取检查信息
  Procedure p_GetStudyInfoByAdviceId(
    Val       Out t_Refcur,
    医嘱id_In In 影像检查记录.医嘱id%Type
  );

  --功    能：获取报告图像总数
  Procedure p_GetReportImageCount(
    Val Out t_Refcur,
    查询条件_In In varchar2
  );

  --功    能：获取报告图像数据
  Procedure p_GetReportImageData(
    Val         Out t_Refcur,
    查询条件_In In varchar2,
    开始位置_In In number,
    结束位置_In In number
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
  
  --功能：根据图像UID获取检查信息
  Procedure p_GetStudyInfoByImageUID(
    Val Out t_Refcur,
    医嘱ID_In In 影像检查记录.医嘱ID%Type,
    图像UID_In In 影像检查图象.图像UID%Type
  );
  
  --功能：根据检查UID获取FTP信息
  Procedure p_GetFtpinfoByStudyUID(
    Val Out t_Refcur,
    检查UID_In In 影像检查记录.检查UID%Type
  );
  
  --功能：根据科室ID获取FTP信息
  Procedure p_GetFtpinfoByDeptId(
    Val Out t_Refcur,
    科室ID_In In 影像流程参数.科室ID%Type
  );
  
  --功能：根据医嘱ID获取FTP信息
  Procedure p_GetFtpinfoByAdvicetId(
    Val Out t_Refcur,
    医嘱ID_In In 影像检查记录.医嘱ID%Type
  );
  
  --功能：获取检查UID
  Procedure p_GetStudyUID(
    Val Out t_Refcur,
    检查UID_In In 影像检查记录.检查UID%Type
  );
  
  --功能：获取序列UID
  Procedure p_GetSeriesUID(
    Val Out t_Refcur,
    序列UID_In In 影像检查序列.序列UID%Type
  );
  
  --功能：根据设备号获取设备信息
  Procedure p_GetDeviceInfo(
    Val Out t_Refcur,
    设备号_In In 影像设备目录.设备号%Type
  );
  
  --获取医技站存储设备号
  Procedure p_GetDeviceIdByAdviceId(
    Val Out t_Refcur,
    医嘱ID_In In 病人医嘱发送.医嘱ID%Type
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

  --功    能：根据医嘱ID获取检查信息
  Procedure p_GetStudyInfoByAdviceId(
    Val       Out t_Refcur,
    医嘱id_In In 影像检查记录.医嘱id%Type
  ) Is
    strSql varchar2(100);
  Begin
    strSql := 'Select 检查UID,报告图象,接收日期,检查号,姓名,性别,年龄 from 影像检查记录 where 医嘱ID =' || 医嘱id_In;
    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyInfoByAdviceId;

  --功    能：获取报告图像总数
  Procedure p_GetReportImageCount(
    Val Out t_Refcur,
    查询条件_In In varchar2
  ) Is
  Begin
    Open Val For
      Select Count(B.Column_Value) 返回值
      From 影像检查记录 A, Table(Cast(f_Str2list(Replace(A.报告图象,';',',')) As zlTools.t_Strlist)) B Where 医嘱ID = 查询条件_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetReportImageCount;

  --功    能：获取报告图像数据
  Procedure p_GetReportImageData(
    Val         Out t_Refcur,
    查询条件_In In varchar2,
    开始位置_In In number,
    结束位置_In In number
  ) Is
  Begin
    Open Val For
         Select * from (Select rownum as 顺序号, rownum as 图像号, B.FTP用户名 As User1,B.FTP密码 As Pwd1,B.IP地址 As Host1,'/'||B.Ftp目录||'/' As Root1,
          Decode(A.接收日期,Null,'',to_Char(A.接收日期,'YYYYMMDD')||'/')||A.检查UID||'/'||Replace(D.Column_Value,'.jpg','') As URL,B.设备号 as 设备号1,
          C.FTP用户名 As User2,C.FTP密码 As Pwd2,C.IP地址 As Host2,'/'||C.Ftp目录||'/' As Root2,
          C.设备号 as 设备号2,Replace(D.Column_Value,'.jpg','') AS 图像UID,A.检查UID,'' 序列UID,0 动态图,'' 编码名称,'' 采集时间, '' 录制长度
          From 影像检查记录 A, 影像设备目录 B, 影像设备目录 C, Table(Cast(f_Str2list(Replace(A.报告图象,';',',')) As zlTools.t_Strlist)) D
          Where A.位置一 = B.设备号(+) And A.位置二 = C.设备号(+) And A.医嘱id = 查询条件_In)
          Where 顺序号 >= 开始位置_In and 顺序号<=结束位置_In;

  End p_GetReportImageData;

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
  
  --功能：根据图像UID获取检查信息
  Procedure p_GetStudyInfoByImageUID(
    Val Out t_Refcur,
    医嘱ID_In In 影像检查记录.医嘱ID%Type,
    图像UID_In In 影像检查图象.图像UID%Type
  )As
  Begin
    Open Val For
      Select D.检查UID From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像临时序列 D
      Where C.医嘱ID=医嘱ID_In And A.图像UID=图像UID_In And A.序列UID=B.序列UID And B.检查UID=C.检查UID And A.序列UID = D.序列UID;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyInfoByImageUID;
  
  --功能：根据检查UID获取FTP信息
  Procedure p_GetFtpinfoByStudyUID(
    Val Out t_Refcur,
    检查UID_In In 影像检查记录.检查UID%Type
  )As
  Begin
    Open Val For
      Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd,C.位置一,C.位置二,C.位置三,C.接收日期,
      D.IP地址 As Host,'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')||C.检查UID As URL
      From 影像检查记录 C,影像设备目录 D Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+) And C.检查UID= 检查UID_In Union All
      Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd,C.位置一,C.位置二,C.位置三,C.接收日期,
      D.IP地址 As Host,'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')||C.检查UID As URL
      From 影像临时记录 C,影像设备目录 D Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+) And C.检查UID= 检查UID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFtpinfoByStudyUID;
  
  --功能：根据科室ID获取FTP信息
  Procedure p_GetFtpinfoByDeptId(
    Val Out t_Refcur,
    科室ID_In In 影像流程参数.科室ID%Type
  )As
  Begin
    Open Val For
      Select a.设备号, a.ip地址, a.ftp用户名, a.ftp密码 From 影像设备目录 a, 影像流程参数 b
      Where a.设备号 = b.参数值 And b.参数名 = '存储设备号' And b.科室id=科室ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFtpinfoByDeptId;
  
  --功能：根据医嘱ID获取FTP信息
  Procedure p_GetFtpinfoByAdvicetId(
    Val Out t_Refcur,
    医嘱ID_In In 影像检查记录.医嘱ID%Type
  )As
  Begin
    Open Val For
      Select a.设备号, a.ip地址, a.ftp用户名, a.ftp密码 From 影像设备目录 a, 影像检查记录 b 
      Where b.位置一 = a.设备号(+) And b.医嘱id =医嘱ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFtpinfoByAdvicetId;
  
  --功能：获取检查UID
  Procedure p_GetStudyUID(
    Val Out t_Refcur,
    检查UID_In In 影像检查记录.检查UID%Type
  )As
  Begin
    Open Val For
      Select 检查UID from 影像检查记录 where 检查UID = 检查UID_In Union All Select 检查UID from 影像临时记录 where 检查UID = 检查UID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyUID;
  
  --功能：获取序列UID
  Procedure p_GetSeriesUID(
    Val Out t_Refcur,
    序列UID_In In 影像检查序列.序列UID%Type
  )As
  Begin
    Open Val For
      Select 序列UID from 影像检查序列 where 序列UID = 序列UID_In Union All Select 序列UID from 影像临时序列 where 序列UID = 序列UID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetSeriesUID;
  
  --功能：根据设备号获取设备信息
  Procedure p_GetDeviceInfo(
    Val Out t_Refcur,
    设备号_In In 影像设备目录.设备号%Type
  )As
  Begin
    Open Val For
      Select 设备号,设备名,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL,FTP用户名,FTP密码,IP地址
      From 影像设备目录 Where 类型=1 and 设备号=设备号_In and NVL(状态,0)=1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDeviceInfo;
  
  --获取医技站存储设备号
  Procedure p_GetDeviceIdByAdviceId(
    Val Out t_Refcur,
    医嘱ID_In In 病人医嘱发送.医嘱ID%Type
  )As
  Begin
    Open Val For
      Select d.参数值 From 医技执行房间 a, 病人医嘱发送 b, 影像DICOM服务对 c, 影像DICOM服务参数 d
      Where a.科室ID = b.执行部门id And a.执行间 = b.执行间 And a.检查设备 = c.设备号
      And c.服务功能='图像接收' And c.服务ID=d.服务ID And d.参数名称='存储设备' And b.医嘱id=医嘱ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDeviceIdByAdviceId;
End b_PACS_RptPluginOriginal;
/


--89676:陈刘,2015-12-30,记录单同步文字项目到体温单
--91458:刘鹏飞,2016-01-04,入量导入处理
Create Or Replace Procedure Zl_病人护理数据_Update
(
  文件id_In   In 病人护理数据.文件id%Type,
  发生时间_In In 病人护理数据.发生时间%Type,
  记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，签名记录=5，审签记录=15 
  项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0 
  记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容；37或38/37 
  体温部位_In In 病人护理明细.体温部位%Type := Null,
  他人记录_In In Number := 1,
  数据来源_In In 病人护理明细.数据来源%Type := 0,
  审签_In     In Number := 0,
  操作员_In   In 病人护理数据.保存人%Type := Null,
  记录组号_In In 病人护理明细.记录组号%Type := Null, --适用分类汇总(一条数据对应多条相同项目的明细) 
  相关序号_In In 病人护理明细.相关序号%Type := Null, --适用分类汇总(记录汇总项目关联的名称项目序号) 
  未记说明_In In 病人护理明细.未记说明%Type := Null --入量导入存储医嘱ID:发送号
) Is
  Intins      Number(18);
  Int共用     Number(1);
  n_Newid     病人护理数据.Id%Type;
  n_Oldid     病人护理数据.Id%Type;
  n_行数      病人护理打印.行数%Type;
  n_Mutilbill Number(1);
  n_Synchro   Number(1);

  n_汇总类别     病人护理数据.汇总类别%Type;
  v_科室id       部门表.Id%Type;
  v_保存人       人员表.姓名%Type;
  v_记录人       人员表.姓名%Type;
  n_文件id       病人护理数据.文件id%Type;
  n_记录id       病人护理数据.Id%Type;
  n_明细id       病人护理明细.Id%Type;
  n_来源id       病人护理明细.来源id%Type;
  v_数据来源     病人护理明细.数据来源%Type;
  n_最高版本     病人护理明细.开始版本%Type;
  n_项目性质     护理记录项目.项目性质%Type;
  n_病人id       病人护理文件.病人id%Type;
  n_主页id       病人护理文件.主页id%Type;
  n_婴儿         病人护理文件.婴儿%Type;
  d_婴儿出院时间 病人医嘱记录.开始执行时间%Type;
  d_文件开始时间 病人护理文件.开始时间%Type;
  --提取该病人当前科室所有未结束的护理文件，且文件开始时间小于等于记录发生时间的文件列表供同步数据使用 
  Cursor Cur_Fileformats Is
    Select a.Id As 格式id, b.Id As 文件id, a.保留, a.子类, b.婴儿
    From 病历文件列表 a, 病人护理文件 b, 病人护理文件 c, 病人护理数据 d
    Where a.种类 = 3 And a.保留 <> 1 And a.Id = b.格式id And b.Id <> c.Id And b.结束时间 Is Null And b.开始时间 <= d.发生时间 And
          (a.通用 = 1 Or (a.通用 = 2 And b.科室id = c.科室id)) And c.病人id = b.病人id And c.主页id = b.主页id And c.婴儿 = b.婴儿 And
          c.Id = d.文件id And d.Id = n_记录id And c.Id = 文件id_In
    Order By a.编号;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --取记录ID 
  Int共用     := 0;
  n_记录id    := 0;
  n_Mutilbill := 0;
  If 操作员_In Is Null Then
    v_保存人 := Zl_Username;
  Else
    v_保存人 := 操作员_In;
  End If;

  --如果是对应多份护理文件值为1，表示需同步其它护理文件；否则不处理文件同步 
  n_Mutilbill := Zl_To_Number(Zl_Getsysparameter('对应多份护理文件', 1255));

  Begin
    Select Id, 汇总类别
    Into n_记录id, n_汇总类别
    From 病人护理数据
    Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
  Exception
    When Others Then
      n_记录id := 0;
  End;

  --检查是不是本人的记录 
  --------------------------------------------------------------------------------------------------------------------- 
  If 他人记录_In = 0 And n_记录id > 0 And 审签_In = 0 Then
    v_记录人 := '';
    Begin
      Select 记录人
      Into v_记录人
      From 病人护理明细
      Where 记录id = n_记录id And 项目序号 = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
            Nvl(记录组号, 0) = Nvl(记录组号_In, 0) And 终止版本 Is Null;
    Exception
      When Others Then
        v_记录人 := '';
    End;
    If v_记录人 Is Not Null And v_记录人 <> v_保存人 Then
      v_Error := '你无权修改他人登记的护理数据！';
      Raise Err_Custom;
    End If;
  End If;

  --检查是否入科 
  Select 病人id, 主页id, Nvl(婴儿, 0), 开始时间
  Into n_病人id, n_主页id, n_婴儿, d_文件开始时间
  From 病人护理文件
  Where Id = 文件id_In;
  d_婴儿出院时间 := Null;
  If n_婴儿 <> 0 Then
    Begin
      Select 开始执行时间
      Into d_婴儿出院时间
      From 病人医嘱记录 b, 诊疗项目目录 c
      Where b.诊疗项目id + 0 = c.Id And b.医嘱状态 = 8 And Nvl(b.婴儿, 0) <> 0 And c.类别 = 'Z' And
            Instr(',3,5,11,', ',' || c.操作类型 || ',', 1) > 0 And b.病人id = n_病人id And b.主页id = n_主页id And b.婴儿 = n_婴儿;
    Exception
      When Others Then
        d_婴儿出院时间 := Null;
    End;
  End If;
  If d_婴儿出院时间 Is Null Then
    v_科室id := 0;
    Begin
      Select a.科室id
      Into v_科室id
      From 病人变动记录 a, 病人护理文件 b
      Where a.科室id Is Not Null And a.病人id = b.病人id And a.主页id = b.主页id And b.Id = 文件id_In And
            (To_Date(To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI') || '59', 'YYYY-MM-DD HH24:MI:SS') >= a.开始时间 And
            (To_Date(To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI') || '00', 'YYYY-MM-DD HH24:MI:SS') < =
            Nvl(a.终止时间, Sysdate) Or a.终止时间 Is Null)) And Rownum < 2;
    Exception
      When Others Then
        v_科室id := 0;
    End;
    If v_科室id = 0 Then
      v_Error := '数据发生时间 ' || To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI:SS') || ' 不在病人有效变动时间范围内，不能操作！';
      Raise Err_Custom;
    End If;
  Else
    If 发生时间_In < d_文件开始时间 Or 发生时间_In > d_婴儿出院时间 Then
      v_Error := '数据发生时间 ' || To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI:SS') || ' 不在病人有效变动时间范围内，不能操作！';
      Raise Err_Custom;
    End If;
  End If;

  --如果数据来源<>0则退出 
  n_来源id := 0;
  If n_记录id > 0 Then
    Begin
      Select 数据来源, Nvl(来源id, 0)
      Into v_数据来源, n_来源id
      From 病人护理明细
      Where 记录id = n_记录id And Nvl(项目序号, 0) = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
            Nvl(记录组号, 0) = Nvl(记录组号_In, 0);
    Exception
      When Others Then
        v_数据来源 := 0;
    End;
    If v_数据来源 > 0 And n_来源id > 0 Then
      Return;
    End If;
  End If;

  --取最高版本 
  Select Nvl(Max(Nvl(a.开始版本, 1)), 0) + 1, Count(b.Id)
  Into n_最高版本, Intins
  From 病人护理明细 a, 病人护理数据 b
  Where b.Id = n_记录id And a.记录id = b.Id And Mod(a.记录类型, 10) = 5;

  --目前已经签名的数据不能修改，只有在审签模式下进行修改，即审签_In=1 
  If 审签_In <> 1 And Intins > 0 Then
    v_Error := '发生时间 ' || To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI:SS') || ' 所对应的数据已经签名或审签，不能继续操作！' || Chr(13) || Chr(10) ||
               '这可能是由于网络并发操作引起的，请刷新后再试！';
    Raise Err_Custom;
  End If;
  Intins := 0;

  --无内容时,要清除数据（审签回退时会自动清除审签过程中修改的数据，所以此处只需考虑普签即可） 
  If 记录内容_In Is Null Then
    Begin
      Select Id
      Into n_明细id
      From 病人护理明细
      Where 记录id = n_记录id And Nvl(项目序号, 0) = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
            Nvl(记录组号, 0) = Nvl(记录组号_In, 0) And 终止版本 Is Null;
    Exception
      --无数据退出 
      When Others Then
        Return;
    End;
  
    --查找除了本条要删除的数据，是否还存其他有效的数据，如果存在只删除本条数据，否则删除此发生时间对应的所有数据。 
    Select Count(Id)
    Into Intins
    From 病人护理明细
    Where 记录id = n_记录id And Mod(记录类型, 10) <> 5 And 终止版本 Is Null And Id <> n_明细id;
    If Intins = 0 Then
      Delete From 病人护理明细 Where 记录id = n_记录id;
    Else
      Delete From 病人护理明细 Where Id = n_明细id;
    End If;
  
    Delete From 病人护理数据 a
    Where a.Id = n_记录id And Not Exists (Select 1 From 病人护理明细 b Where b.记录id = a.Id);
  
    --如果是删除签名后修改产生的最后一条数据,则应将签名记录的终止版本清为空 
    Begin
      Select 1
      Into Intins
      From 病人护理明细
      Where 开始版本 = n_最高版本 And 终止版本 Is Null And 记录类型 = 1 And 记录id = n_记录id;
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Update 病人护理明细 Set 终止版本 = Null Where 记录类型 = 5 And 开始版本 = n_最高版本 - 1 And 记录id = n_记录id;
    End If;
    If Nvl(n_汇总类别, 0) <> 0 Then
      Return;
    End If;
  
    --############ 
    --清除共用数据 
    --############ 
    For Rsdel In (Select Distinct 记录id From 病人护理明细 Where 来源id = n_明细id) Loop
    
      Delete 病人护理明细 Where 来源id = n_明细id And 记录id = Rsdel.记录id;
      --删除对应的打印数据 
      Begin
        Select Count(*) Into Intins From 病人护理明细 Where 记录id = Rsdel.记录id;
      Exception
        When Others Then
          Intins := 0;
      End;
      If Intins = 0 Then
        --提取清除数据对应的文件ID 
        Begin
          Select b.Id, a.保留
          Into n_文件id, Intins
          From 病历文件列表 a, 病人护理文件 b, 病人护理数据 c
          Where a.Id = b.格式id And b.Id = c.文件id And c.Id = Rsdel.记录id;
        Exception
          When Others Then
            n_文件id := 0;
        End;
        Delete 病人护理数据 Where Id = Rsdel.记录id;
        If Intins <> -1 Then
          Zl_病人护理打印_Update(n_文件id, 发生时间_In, 1, 1);
        End If;
      End If;
    End Loop;
  Else
    --检查录入的项目是否属于该记录单 
    Begin
      Select 1
      Into Intins
      From (Select b.项目序号
             From 病历文件结构 a, 护理记录项目 b
             Where a.要素名称 = b.项目名称 And b.项目序号 = 项目序号_In And
                   父id = (Select b.Id
                          From 病人护理文件 a, 病历文件结构 b
                          Where a.Id = 文件id_In And a.格式id = b.文件id And b.父id Is Null And b.对象序号 = 4)
             Union
             Select 项目序号
             From 护理记录项目
             Where 项目性质 = 2 And 项目序号 = 项目序号_In);
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Return;
    End If;
    If n_记录id = 0 Then
      Select 病人护理数据_Id.Nextval Into n_记录id From Dual;
    
      Insert Into 病人护理数据
        (Id, 文件id, 发生时间, 最后版本, 保存人, 保存时间)
      Values
        (n_记录id, 文件id_In, 发生时间_In, n_最高版本, v_保存人, Sysdate);
    End If;
  
    --插入本次登记的病人护理明细 
    Update 病人护理明细
    Set 记录内容 = 记录内容_In, 数据来源 = 数据来源_In, 未记说明 = 未记说明_In, 记录人 = v_保存人, 记录时间 = Sysdate
    Where 记录id = n_记录id And 项目序号 = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
          Nvl(记录组号, 0) = Nvl(记录组号_In, 0) And 开始版本 = n_最高版本 And 终止版本 Is Null;
    If Sql%Rowcount = 0 Then
      Select 病人护理明细_Id.Nextval Into n_明细id From Dual;
      Insert Into 病人护理明细
        (Id, 记录id, 记录类型, 项目分组, 项目id, 相关序号, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 记录组号, 体温部位, 数据来源, 共用, 未记说明, 开始版本, 终止版本,
         记录人, 记录时间)
        Select n_明细id, n_记录id, 记录类型_In, a.分组名, a.项目id, 相关序号_In, a.项目序号, Upper(a.项目名称), a.项目类型, 记录内容_In, a.项目单位, 0,
               记录组号_In, 体温部位_In, 数据来源_In, Nvl(b.共用, 0), 未记说明_In, n_最高版本, Null, v_保存人, Sysdate
        From 护理记录项目 a, 病人护理明细 b
        Where a.项目序号 = b.项目序号(+) And b.终止版本(+) Is Null And b.记录id(+) = n_记录id And a.项目序号 = 项目序号_In And Rownum < 2;
    End If;
    Select Id
    Into n_明细id
    From 病人护理明细
    Where 记录id = n_记录id And 项目序号 = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
          Nvl(记录组号, 0) = Nvl(记录组号_In, 0) And 开始版本 = n_最高版本 And 终止版本 Is Null;
    --填写历史数据及签名记录的终止版本 
    Update 病人护理明细
    Set 终止版本 = n_最高版本
    Where 记录id = n_记录id And ((Mod(记录类型, 10) <> 5 And 项目序号 = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
          Nvl(记录组号, 0) = Nvl(记录组号_In, 0)) Or 记录类型 = Decode(审签_In, 1, 15, 5)) And 开始版本 <= n_最高版本 - 1 And 终止版本 Is Null;
  
    --如果是未签名数据，最后修改操作员做为该记录的保存人更新 
    If n_最高版本 = 1 Then
      Update 病人护理数据 Set 保存人 = v_保存人, 保存时间 = Sysdate Where Id = n_记录id;
    End If;
  
    If Nvl(n_汇总类别, 0) <> 0 Then
      Return;
    End If;
  
    --############ 
    --同步共用数据 
    --############ 
    --1\先处理体温单（一个病人始终只存在一份有效的体温单文件） 
    --如果体温表存在相同发生时间的数据，使用它的ID 
    --CL,2015-12-30,记录单同步文字项目到体温单
    For Row_Format In Cur_Fileformats Loop
      If Row_Format.保留 = -1 Then
        If Row_Format.子类 = '1' Then
          Begin
            Select 1, h.项目性质
            Into Intins, n_项目性质
            From (Select To_Char(f.项目序号) As 序号, g.项目性质
                   From 体温记录项目 f, 护理记录项目 g
                   Where f.项目序号 = g.项目序号 And g.项目性质 = 2 And
                         (g.适用科室 = 1 Or
                         (g.适用科室 = 2 And Exists
                          (Select 1 From 护理适用科室 d Where g.项目序号 = d.项目序号 And d.科室id = v_科室id))) And Nvl(g.应用方式, 0) <> 0 And
                         (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(Nvl(Row_Format.婴儿, 0), 0, 1, 2))
                   Union All
                   Select b.内容文本 As 序号, 1 As 项目性质
                   From 病历文件结构 a, 病历文件结构 b
                   Where a.文件id = Row_Format.格式id And a.父id Is Null And a.对象序号 In (2, 3) And b.父id = a.Id) h
            Where Instr(',' || h.序号 || ',', ',' || 项目序号_In || ',', 1) > 0;
          Exception
            When Others Then
              Intins := 0;
          End;
        Else
          Begin
            Select 1, g.项目性质
            Into Intins, n_项目性质
            From 体温记录项目 f, 护理记录项目 g
            Where f.项目序号 = g.项目序号 And Nvl(g.应用方式, 0) = 1 And g.护理等级 >= 0 And
                  (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(Nvl(Row_Format.婴儿, 0), 0, 1, 2)) And f.项目序号 = 项目序号_In And
                  (g.适用科室 = 1 Or (g.适用科室 = 2 And Exists
                   (Select 1 From 护理适用科室 d Where g.项目序号 = d.项目序号 And d.科室id = v_科室id)));
          Exception
            When Others Then
              Intins := 0;
          End;
        End If;
      
        If Intins > 0 Then
          --LPF,2013-01-23,检查此项目是否需要进行同步(对于以前已经同步过的数据，为了保证记录单和体温单数据一直将不根据此函数判断。) 
          n_Synchro := Zl_Temperatureprogram(文件id_In, v_科室id, 项目序号_In, 发生时间_In);
          Begin
            Select b.Id
            Into n_Newid
            From 病人护理文件 a, 病人护理数据 b
            Where a.Id = Row_Format.文件id And b.文件id = a.Id And b.发生时间 = 发生时间_In;
          Exception
            When Others Then
              n_Newid := 0;
          End;
          n_Oldid := n_Newid;
          If n_Newid = 0 And n_Synchro = 1 Then
            Select 病人护理数据_Id.Nextval Into n_Newid From Dual;
            --产生体温单主记录 
            Insert Into 病人护理数据
              (Id, 文件id, 保存人, 保存时间, 发生时间, 最后版本)
            Values
              (n_Newid, Row_Format.文件id, v_保存人, Sysdate, 发生时间_In, 1);
          End If;
        
          If n_Newid > 0 Then
            --插入未同步的体温单数据(仍然要联接多表查询) 
            Select Count(*)
            Into v_数据来源
            From 病人护理明细
            Where 记录id = n_Newid And 项目序号 = 项目序号_In And
                  Decode(n_项目性质, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无');
            If v_数据来源 = 0 Then
              --说明在同步开始已经进行过检查 
              If n_Synchro = 1 Then
                --没有检查此项目是否需要同步
                Insert Into 病人护理明细
                  (Id, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 体温部位, 数据来源, 来源id, 开始版本, 终止版本, 记录人,
                   记录时间, 记录组号)
                  Select 病人护理明细_Id.Nextval, n_Newid, b.记录类型, b.项目分组, b.项目id, b.项目序号, b.项目名称, b.项目类型, b.记录内容, b.项目单位,
                         b.记录标记, b.体温部位, 1, b.Id, 1, Null, b.记录人, Sysdate, 1
                  From (Select 项目序号_In As 项目序号, Nvl(体温部位_In, '无') As 体温部位
                         From Dual
                         Minus
                         Select f.项目序号, Decode(Nvl(f.项目性质, 1), 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无'))
                         From 病人护理明细 e, 护理记录项目 f
                         Where e.记录id = n_Newid And e.项目序号 = f.项目序号) a, 病人护理明细 b
                  Where a.项目序号 = b.项目序号 And b.记录id = n_记录id And b.Id = n_明细id;
                If Sql%Rowcount > 0 Then
                  Int共用 := 1;
                End If;
              End If;
            Else
              Update 病人护理明细
              Set 记录内容 = 记录内容_In, 来源id = n_明细id
              Where 记录id = n_Newid And 项目序号 = 项目序号_In And
                    Decode(n_项目性质, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And 数据来源 > 0;
              If Sql%Rowcount > 0 Then
                Int共用 := 1;
              End If;
            End If;
          End If;
        End If;
        --2\再循环处理记录单 
      Else
        If n_Mutilbill = 1 Then
          --提取记录单与当前记录单存在重叠的且有数据的固定项目 
          Select Count(*)
          Into Intins
          From (Select b.项目序号
                 From 病历文件结构 a, 护理记录项目 b
                 Where a.要素名称 = b.项目名称 And b.项目表示 In (0, 4, 5) And
                       父id =
                       (Select Id From 病历文件结构 Where 文件id = Row_Format.格式id And 父id Is Null And 对象序号 = 4)
                 Intersect
                 Select b.项目序号
                 From 病历文件结构 a, 护理记录项目 b, 病人护理文件 c, 病人护理数据 d, 病人护理明细 g
                 Where c.Id = d.文件id And a.文件id = c.格式id And d.Id = g.记录id And d.Id = n_记录id And g.Id = n_明细id And
                       b.项目序号 = g.项目序号 And b.项目表示 In (0, 4, 5) And g.记录类型 = 1 And a.要素名称 = b.项目名称 And
                       a.父id = (Select Id From 病历文件结构 e Where e.文件id = c.格式id And 父id Is Null And 对象序号 = 4));
        
          If Intins > 0 Then
            n_Newid := 0;
            --可能指定文件已经存在相同发生时间的数据，直接用它的ID即可 
            Begin
              Select c.Id
              Into n_Newid
              From 病人护理数据 c
              Where c.文件id = Row_Format.文件id And c.发生时间 = 发生时间_In;
            Exception
              When Others Then
                n_Newid := 0;
            End;
          
            If n_Newid = 0 Then
              --产生记录单主记录 
              Select 病人护理数据_Id.Nextval Into n_Newid From Dual;
            
              Insert Into 病人护理数据
                (Id, 文件id, 保存人, 保存时间, 发生时间, 最后版本)
                Select n_Newid, Row_Format.文件id, c.保存人, c.保存时间, c.发生时间, 1
                From 病人护理数据 c
                Where c.Id = n_记录id;
            End If;
          
            If n_Newid > 0 Then
              --插入未同步的记录单数据 
              Select Count(*) Into v_数据来源 From 病人护理明细 Where 记录id = n_Newid And 项目序号 = 项目序号_In;
              If v_数据来源 = 0 Then
                Insert Into 病人护理明细
                  (Id, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 体温部位, 数据来源, 来源id, 未记说明, 开始版本, 终止版本,
                   记录人, 记录时间)
                  Select 病人护理明细_Id.Nextval, n_Newid, b.记录类型, b.项目分组, b.项目id, b.项目序号, b.项目名称, b.项目类型, b.记录内容, b.项目单位,
                         b.记录标记, b.体温部位, 1, b.Id, b.未记说明, 1, Null, b.记录人, Sysdate
                  From (Select b.项目序号
                         From 病历文件结构 a, 护理记录项目 b
                         Where a.要素名称 = b.项目名称 And b.项目表示 In (0, 4, 5) And
                               父id = (Select Id
                                      From 病历文件结构
                                      Where 文件id = Row_Format.格式id And 父id Is Null And 对象序号 = 4)
                         Intersect
                         Select b.项目序号
                         From 病历文件结构 a, 护理记录项目 b, 病人护理文件 c, 病人护理数据 d, 病人护理明细 g
                         Where c.Id = d.文件id And a.文件id = c.格式id And d.Id = g.记录id And d.Id = n_记录id And g.Id = n_明细id And
                               b.项目序号 = g.项目序号 And b.项目表示 In (0, 4, 5) And g.记录类型 = 1 And a.要素名称 = b.项目名称 And
                               a.父id =
                               (Select Id From 病历文件结构 e Where e.文件id = c.格式id And 父id Is Null And 对象序号 = 4)) a, 病人护理明细 b
                  Where a.项目序号 = b.项目序号 And b.记录id = n_记录id And b.Id = n_明细id;
                If Sql%Rowcount > 0 Then
                  Int共用 := 1;
                  --原行数不要动 
                  Begin
                    Select 行数 Into n_行数 From 病人护理打印 Where 文件id = Row_Format.文件id And 记录id = n_Newid;
                  Exception
                    When Others Then
                      n_行数 := 1;
                  End;
                  Zl_病人护理打印_Update(Row_Format.文件id, 发生时间_In, n_行数, 0);
                End If;
              Else
                Update 病人护理明细
                Set 记录内容 = 记录内容_In, 未记说明 = 未记说明_In, 来源id = n_明细id
                Where 记录id = n_Newid And 项目序号 = 项目序号_In And 数据来源 > 0;
                If Sql%Rowcount > 0 Then
                  Int共用 := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End Loop;
  
    If Int共用 = 1 Then
      Update 病人护理明细 Set 共用 = 1 Where Id = n_明细id;
      --将历史数据的共用标志设置为NULL 
      Update 病人护理明细 Set 共用 = Null Where 记录id = n_记录id And 项目序号 = 项目序号_In And Id <> n_明细id;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人护理数据_Update;
/

--91458:刘鹏飞,2016-01-04,入量导入处理
Create Or Replace Procedure Zl_护理文件样式_Update
(
  文件id_In     In 病历文件结构.文件id%Type,
  表头层数_In   Number,
  总列数_In     Number,
  最小行高_In   Number,
  文本字体_In   Varchar2,
  文本颜色_In   Number,
  表格颜色_In   Number,
  标题文本_In   Varchar2,
  标题字体_In   Varchar2,
  开始时间_In   Number,
  终止时间_In   Number,
  条件字体_In   Varchar2,
  条件颜色_In   Number,
  有效数据行_In Number,
  首列合并_In   Number,
  时间隐藏_In   Number, --记录单预览、打印时隐藏时间列(如：血糖记录单不需要显示具体的时间)
  页面格式_In   病历页面格式.格式%Type, --按.PaperKind;.PaperOrient;.PaperHeight;.PaperWidth;.MarginLeft;.MarginRight;.MarginTop;.MarginBottom组织
  页眉文本_In   病历页面格式.页眉%Type,
  页脚文本_In   病历页面格式.页脚%Type,
  表上标签_In   Varchar2, --按照"前缀{项目}"组织，以"|"为分隔的表上标签集合
  表头单元_In   Varchar2, --按照"列号,层号,文本"组织，以"|"为分隔的表头单元集合
  表列集合_In   Varchar2, --按照"列号,列宽,项目集合"组织，以"|"为分隔的表列集合；其中项目集合组织为"前缀{项目}后缀`是否汇总", 空格分隔。
  汇总时段_In   Varchar2 := Null, --按照"时段名称,开始时间点,结束时间点"组织，以"|"为分隔的集合。
  表下标签_In   Varchar2 := Null, --按照"前缀{项目}"组织，以"|"为分隔的表上标签集合
  分类汇总_In   Number := Null
) Is
  v_Items    Varchar2(4000); --项目集合
  v_Subitems Varchar2(4000); --项目集合
  v_Fields   Varchar2(4000); --一个项目的属性组合
  v_Colno    Varchar2(100); --项目列号
  n_父id     病历文件结构.父id%Type;
  n_对象序号 病历文件结构.对象序号%Type;
  n_对象标记 病历文件结构.对象标记%Type;
  v_对象属性 病历文件结构.对象属性%Type;
  n_内容行次 病历文件结构.内容行次%Type;
  v_内容文本 病历文件结构.内容文本%Type;
  v_是否换行 病历文件结构.是否换行%Type;
  v_要素名称 病历文件结构.要素名称%Type;
  v_要素单位 病历文件结构.要素单位%Type;
  n_要素表示 病历文件结构.要素表示%Type;
Begin
  Delete 病历文件结构 Where 文件id = 文件id_In;

  Update 病历页面格式
  Set 格式 = 页面格式_In, 页眉 = 页眉文本_In, 页脚 = 页脚文本_In
  Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = 文件id_In);

  Select 病历文件结构_Id.Nextval Into n_父id From Dual;
  Insert Into 病历文件结构
    (Id, 文件id, 对象序号, 对象类型, 对象属性, 内容文本)
  Values
    (n_父id, 文件id_In, 1, 1, '表格基本属性和样式的说明', '表格样式');
  Insert Into 病历文件结构
    (Id, 文件id, 父id, 对象类型, 对象序号, 对象属性, 内容文本, 要素名称)
    Select 病历文件结构_Id.Nextval, 文件id_In, n_父id, 4, 序号, 属性, 文本, 名称
    From (Select 1 As 序号, '目前支持单层(1)和多层次(2)' As 属性, To_Char(表头层数_In) As 文本, '表头层数' As 名称
           From Dual
           Union All
           Select 2, '表格总共有的列数', To_Char(总列数_In), '总列数'
           From Dual
           Union All
           Select 3, '每行的最小高度(缇)', To_Char(最小行高_In), '最小行高'
           From Dual
           Union All
           Select 4, '表格的默认字体', 文本字体_In, '文本字体'
           From Dual
           Union All
           Select 5, '表格文本颜色RGB值', To_Char(文本颜色_In), '文本颜色'
           From Dual
           Union All
           Select 6, '表格线的基本颜色', To_Char(表格颜色_In), '表格颜色'
           From Dual
           Union All
           Select 7, '标题的文字内容', 标题文本_In, '标题文本'
           From Dual
           Union All
           Select 8, '标题的字体', 标题字体_In, '标题字体'
           From Dual
           Union All
           Select 9, '按24小时表示的条件开始范围', To_Char(开始时间_In), '开始时间'
           From Dual
           Union All
           Select 10, '小于开始时间表示次日终止', To_Char(终止时间_In), '终止时间'
           From Dual
           Union All
           Select 11, '符合条件的内容记录的字体', 条件字体_In, '条件字体'
           From Dual
           Union All
           Select 13, '有效数据行', To_Char(有效数据行_In), '有效数据行'
           From Dual
           Union All
           Select 14, '日期时间合并', To_Char(首列合并_In), '日期时间合并'
           From Dual
           Union All
           Select 12, '符合表件的内容记录的颜色', To_Char(条件颜色_In), '条件颜色'
           From Dual
           Union All
           Select 15, '时间列隐藏', To_Char(时间隐藏_In), '时间列隐藏'
           From Dual
           Union All
           Select 16, '分类汇总', To_Char(分类汇总_In), '分类汇总'
           From Dual);

  Select 病历文件结构_Id.Nextval Into n_父id From Dual;
  Insert Into 病历文件结构
    (Id, 文件id, 对象序号, 对象类型, 对象属性, 内容文本)
  Values
    (n_父id, 文件id_In, 2, 1, '由替换项组成的表上项目', '表上标签');
  If 表头单元_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(表上标签_In) || '|';
  End If;
  n_对象序号 := 0;
  While v_Items Is Not Null Loop
    v_Fields   := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    n_对象序号 := n_对象序号 + 1;
    v_内容文本 := Substr(v_Fields, 1, Instr(v_Fields, '{') - 1);
    v_要素名称 := Substr(v_Fields, Instr(v_Fields, '{') + 1, Instr(v_Fields, '}') - Instr(v_Fields, '{') - 1);
    If Substr(v_内容文本, 1, 2) = Chr(13) || Chr(10) Then
      v_是否换行 := 1;
      v_内容文本 := Substr(v_内容文本, 3);
    Else
      v_是否换行 := 0;
    End If;
    Insert Into 病历文件结构
      (Id, 文件id, 父id, 对象类型, 对象序号, 内容文本, 要素名称, 是否换行)
    Values
      (病历文件结构_Id.Nextval, 文件id_In, n_父id, 4, n_对象序号, v_内容文本, v_要素名称, v_是否换行);
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;

  Select 病历文件结构_Id.Nextval Into n_父id From Dual;
  Insert Into 病历文件结构
    (Id, 文件id, 对象序号, 对象类型, 对象属性, 内容文本)
  Values
    (n_父id, 文件id_In, 3, 1, '组成表头的各单元内容', '表头单元');
  If 表头单元_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(表头单元_In) || '|';
  End If;
  While v_Items Is Not Null Loop
    v_Fields   := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    n_对象序号 := To_Number(Substr(v_Fields, 1, Instr(v_Fields, ',', 1, 1) - 1));
    n_内容行次 := To_Number(Substr(v_Fields,
                               Instr(v_Fields, ',', 1, 1) + 1,
                               Instr(v_Fields, ',', 1, 2) - Instr(v_Fields, ',', 1, 1) - 1));
    v_内容文本 := Substr(v_Fields, Instr(v_Fields, ',', 1, 2) + 1);
    Insert Into 病历文件结构
      (Id, 文件id, 父id, 对象类型, 对象序号, 内容行次, 内容文本)
    Values
      (病历文件结构_Id.Nextval, 文件id_In, n_父id, 2, n_对象序号, n_内容行次, v_内容文本);
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;

  Select 病历文件结构_Id.Nextval Into n_父id From Dual;
  Insert Into 病历文件结构
    (Id, 文件id, 对象序号, 对象类型, 对象属性, 内容文本)
  Values
    (n_父id, 文件id_In, 4, 1, '表体各数据列的定义设置', '表列集合');
  If 表列集合_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(表列集合_In) || '|';
  End If;
  While v_Items Is Not Null Loop
    v_Fields := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    --汇总列设置了对照关系列号为:项目列号`对照列号
    v_Colno := Substr(v_Fields, 1, Instr(v_Fields, ',', 1, 1) - 1);
    If Instr(v_Colno, '`', 1, 1) > 0 Then
      n_对象序号 := To_Number(Substr(v_Colno, 1, Instr(v_Colno, '`', 1, 1) - 1));
      n_对象标记 := To_Number(Substr(v_Colno, Instr(v_Colno, '`', 1, 1) + 1));
    Else
      n_对象序号 := To_Number(v_Colno);
      n_对象标记 := Null;
    End If;
    v_对象属性 := Substr(v_Fields,
                     Instr(v_Fields, ',', 1, 1) + 1,
                     Instr(v_Fields, ',', 1, 2) - Instr(v_Fields, ',', 1, 1) - 1);
    v_Subitems := Substr(v_Fields, Instr(v_Fields, ',', 1, 2) + 1);
    If v_Subitems Is Null Then
      Insert Into 病历文件结构
        (Id, 文件id, 父id, 对象类型, 对象序号, 对象标记, 对象属性, 内容行次, 内容文本, 要素名称, 要素单位)
      Values
        (病历文件结构_Id.Nextval, 文件id_In, n_父id, 4, n_对象序号, n_对象标记, v_对象属性, 1, '', '', '');
    Else
      v_Subitems := Rtrim(v_Subitems) || ' ';
    End If;
    n_内容行次 := 0;
    While v_Subitems Is Not Null Loop
      n_内容行次 := n_内容行次 + 1;
      v_Fields   := Substr(v_Subitems, 1, Instr(v_Subitems, ' ') - 1);
      v_内容文本 := Substr(v_Fields, 1, Instr(v_Fields, '{') - 1);
      v_要素名称 := Substr(v_Fields, Instr(v_Fields, '{') + 1, Instr(v_Fields, '}') - Instr(v_Fields, '{') - 1);
      If Instr(v_Fields, '`') > 0 Then
        v_要素单位 := Substr(v_Fields, Instr(v_Fields, '}') + 1, Instr(v_Fields, '`') - Instr(v_Fields, '}') - 1);
        n_要素表示 := To_Number(Substr(v_Fields, Instr(v_Fields, '`', 1, 1) + 1));
      Else
        v_要素单位 := Substr(v_Fields, Instr(v_Fields, '}') + 1);
        n_要素表示 := 0;
      End If;
      If n_内容行次 > 1 Then
        n_对象标记 := Null;
      End If;
      Insert Into 病历文件结构
        (Id, 文件id, 父id, 对象类型, 对象序号, 对象标记, 对象属性, 内容行次, 内容文本, 要素名称, 要素单位, 要素表示)
      Values
        (病历文件结构_Id.Nextval, 文件id_In, n_父id, 4, n_对象序号, n_对象标记, v_对象属性, n_内容行次, v_内容文本, v_要素名称, v_要素单位, n_要素表示);
      v_Subitems := Substr(v_Subitems, Instr(v_Subitems, ' ') + 1);
    End Loop;
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;

  Select 病历文件结构_Id.Nextval Into n_父id From Dual;
  Insert Into 病历文件结构
    (Id, 文件id, 对象序号, 对象类型, 对象属性, 内容文本)
  Values
    (n_父id, 文件id_In, 5, 1, '汇总时的分组依据', '汇总时段');
  If 汇总时段_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(汇总时段_In) || '|';
  End If;

  n_内容行次 := 0;
  While v_Items Is Not Null Loop
    v_Fields   := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    n_内容行次 := n_内容行次 + 1;
    Insert Into 病历文件结构
      (Id, 文件id, 父id, 对象类型, 对象序号, 内容行次, 内容文本)
    Values
      (病历文件结构_Id.Nextval, 文件id_In, n_父id, 2, n_内容行次, n_内容行次, v_Fields);
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;

  Select 病历文件结构_Id.Nextval Into n_父id From Dual;
  Insert Into 病历文件结构
    (Id, 文件id, 对象序号, 对象类型, 对象属性, 内容文本)
  Values
    (n_父id, 文件id_In, 6, 1, '由替换项组成的表上项目', '表下标签');
  If 表头单元_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(表下标签_In) || '|';
  End If;
  n_对象序号 := 0;
  While v_Items Is Not Null Loop
    v_Fields   := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    n_对象序号 := n_对象序号 + 1;
    v_内容文本 := Substr(v_Fields, 1, Instr(v_Fields, '{') - 1);
    v_要素名称 := Substr(v_Fields, Instr(v_Fields, '{') + 1, Instr(v_Fields, '}') - Instr(v_Fields, '{') - 1);
    If Substr(v_内容文本, 1, 2) = Chr(13) || Chr(10) Then
      v_是否换行 := 1;
      v_内容文本 := Substr(v_内容文本, 3);
    Else
      v_是否换行 := 0;
    End If;
    Insert Into 病历文件结构
      (Id, 文件id, 父id, 对象类型, 对象序号, 内容文本, 要素名称, 是否换行)
    Values
      (病历文件结构_Id.Nextval, 文件id_In, n_父id, 4, n_对象序号, v_内容文本, v_要素名称, v_是否换行);
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_护理文件样式_Update;
/

--92208:刘尔旋,2015-12-29,病人结帐产生三方结算信息
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
      If n_预交id Is Not Null Then
        For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                              Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                       From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
          Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_预交id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 1);
        End Loop;
      Else
        For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                              Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                       From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
          Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
        End Loop;
      End If;
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

--92157:余伟节,2015-12-29,避免住院医嘱编辑界面生成路径医嘱时将病人临床路径的当前阶段和当前天数置为空
Create Or Replace Procedure Zl_病人路径生成_Delete
(
  执行记录id_In 病人路径执行.Id%Type,
  调用模式_In   Number := 0,
  调用场合_In   Number := 0
) Is
  --参数:调用模式_in=0:取消路径项目时调用,=1:重新生成医嘱时调用,=2：取消生成必须生成的项目时,
  --               =3:ZL_病人医嘱记录_Delete调用,防止住院医嘱编辑界面删除路径医嘱时将病人临床路径的当前阶段和当前天数置为空。
  --     调用场合_In =0:医生站  ;1-护士站
  t_Id   t_Numlist;
  t_时间 t_Strlist;
  --长期医嘱,其它阶段存在时不删除,未校对时才删除医嘱(界面已限制已校对但未作废的不允许删除路径项目)
  Cursor c_Advice(导入时间_In 病人临床路径.导入时间%Type) Is
    Select a.病人医嘱id
    From 病人路径医嘱 A, 病人医嘱记录 C
    Where 路径执行id = 执行记录id_In And a.病人医嘱id = c.Id And c.医嘱状态 = 1 And
          To_Date(To_Char(c.开嘱时间 + 59 / 24 / 60 / 60, 'yyyy-mm-dd hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') > 导入时间_In And
          Not Exists
     (Select 1 From 病人路径医嘱 B Where a.病人医嘱id = b.病人医嘱id And a.路径执行id <> b.路径执行id);

  Cursor c_Doc Is
    Select ID, To_Char(创建时间, 'yyyy-MM-dd hh24:mi:ss') From 电子病历记录 Where 路径执行id = 执行记录id_In;

  --删除最后一个项目时，检查前面是否有提前跳过的阶段。
  Cursor c_Turn
  (
    路径记录id_In 病人路径执行.路径记录id%Type,
    天数_In       病人路径执行.天数%Type,
    阶段id_In     病人路径执行.阶段id%Type
  ) Is
    Select a.阶段id
    From 病人路径执行 A
    Where a.路径记录id = 路径记录id_In And a.天数 = 天数_In And a.阶段id <> 阶段id_In And a.项目内容 = '未生成任何项目' And Exists
     (Select 1
           From 病人路径评估 B
           Where a.路径记录id = b.路径记录id And a.阶段id = b.阶段id And a.天数 = b.天数 And b.时间进度 = 1)
    Order By a.登记时间 Desc;
  t_阶段id t_Numlist;

  Cursor c_Merge
  (
    路径记录id_In 病人路径执行.路径记录id%Type,
    阶段id_In     病人路径执行.阶段id%Type
  ) Is
    Select a.Id, Max(b.合并路径阶段id) As 阶段id
    From 病人合并路径 A, 病人合并路径评估 B
    Where a.Id = b.合并路径记录id(+) And a.首要路径记录id = b.路径记录id(+) And b.阶段id(+) <> 阶段id_In And a.首要路径记录id = 路径记录id_In And
          (b.登记时间 = (Select Max(登记时间)
                     From 病人合并路径评估 C
                     Where c.路径记录id = b.路径记录id And c.合并路径记录id = b.合并路径记录id And c.阶段id = b.阶段id) Or b.登记时间 Is Null)
    Group By a.Id;
  t_合并路径阶段id t_Numlist;
  t_合并路径记录id t_Numlist;

  r_Pp_Item 病人路径执行%RowType;

  v_阶段id          病人路径执行.阶段id%Type;
  v_前一阶段id      病人路径执行.阶段id%Type;
  v_路径记录id      病人路径执行.路径记录id%Type;
  v_天数            病人路径执行.天数%Type;
  v_Last天数        病人路径执行.天数%Type;
  v_相关id          病人医嘱记录.相关id%Type;
  v_Other路径执行id 病人路径执行.Id%Type;
  n_Count           Number(5);
  n_变化天数        Number(5);
  d_导入时间        病人临床路径.导入时间%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --如果路径内外项目是一并给药的，删除其中一个路径项目后，要该将其中的给药途径的执行ID更改为一并给药中其他药品的执行ID
  Select Nvl(Max(b.相关id), 0)
  Into v_相关id
  From 病人路径医嘱 A, 病人医嘱记录 B
  Where a.路径执行id = 执行记录id_In And a.病人医嘱id = b.Id And b.诊疗类别 In ('5', '6');

  If v_相关id <> 0 Then
    Select Nvl(Max(c.Id), 0)
    Into v_Other路径执行id
    From 病人路径医嘱 A, 病人医嘱记录 B, 病人路径执行 C
    Where c.Id = a.路径执行id And b.相关id = v_相关id And a.病人医嘱id = b.Id And a.路径执行id <> 执行记录id_In And
          c.登记时间 = (Select Max(d.登记时间)
                    From 病人路径执行 D, 病人医嘱记录 E, 病人路径医嘱 F
                    Where e.Id = f.病人医嘱id And d.Id = f.路径执行id And e.相关id = v_相关id And f.路径执行id <> 执行记录id_In);
  
    If v_Other路径执行id <> 0 Then
      Select Count(1)
      Into n_Count
      From 病人路径执行 A, 病人路径执行 B
      Where a.Id = v_Other路径执行id And b.Id = 执行记录id_In And a.阶段id = b.阶段id And a.天数 = b.天数;
    
      If n_Count > 0 Then
        --如果是两个项目有相同医嘱，对应不同的路径项目（合并路径和首要路径不重复生成的医嘱）时，不修改路径执行ID
        Select Count(1) Into n_Count From 病人路径医嘱 Where 路径执行id = v_Other路径执行id And 病人医嘱id = v_相关id;
        If n_Count = 0 Then
          Update 病人路径医嘱
          Set 路径执行id = v_Other路径执行id
          Where 病人医嘱id = v_相关id And 路径执行id = 执行记录id_In;
        End If;
      End If;
    End If;
  End If;

  --导入时间之前的路径项目对应的病历医嘱不删除
  Select b.导入时间
  Into d_导入时间
  From 病人路径执行 A, 病人临床路径 B
  Where a.路径记录id = b.Id And a.Id = 执行记录id_In;
  If d_导入时间 Is Not Null Then
    --是否允许取消的逻辑规则在界面程序中检查
    Open c_Advice(d_导入时间);
    Fetch c_Advice Bulk Collect
      Into t_Id;
    Close c_Advice;
  
    Delete 病人路径医嘱 Where 路径执行id = 执行记录id_In;
    If t_Id.Count > 0 Then
      Forall I In 1 .. t_Id.Count
        Delete From 病人医嘱记录 Where ID = t_Id(I) And 医嘱状态 = 1;
    End If;
  
    If 调用模式_In = 0 Or 调用模式_In = 2 Then
      Open c_Doc;
      Fetch c_Doc Bulk Collect
        Into t_Id, t_时间;
      Close c_Doc;
      If t_Id.Count > 0 Then
        For I In 1 .. t_Id.Count Loop
          If To_Date(t_时间(I), 'yyyy-MM-dd hh24:mi:ss') > d_导入时间 Then
            Zl_电子病历记录_Delete(t_Id(I));
          Else
            Update 电子病历记录 Set 路径执行id = Null Where ID = t_Id(I);
          End If;
        End Loop;
      End If;
    End If;
  End If;

  --如果是取消生成必须生成的项目时，不删除执行记录
  If 调用模式_In = 3 Then
    Select * Into r_Pp_Item From 病人路径执行 T Where ID = 执行记录id_In;
    Delete 病人路径执行 Where ID = 执行记录id_In;
    Select Count(1)
    Into n_Count
    From 病人路径执行
    Where 路径记录id = r_Pp_Item.路径记录id And 阶段id = r_Pp_Item.阶段id And 天数 = r_Pp_Item.天数;
    If n_Count = 0 Then
      --增加一个特殊项目[未生成任何项目]
      Insert Into 病人路径执行
        (ID, 路径记录id, 阶段id, 日期, 天数, 分类, 项目id, 登记人, 登记时间, 项目序号, 项目内容, 执行者, 生成者, 项目结果)
      Values
        (病人路径执行_Id.Nextval, r_Pp_Item.路径记录id, r_Pp_Item.阶段id, r_Pp_Item.日期, r_Pp_Item.天数, r_Pp_Item.分类, Null,
         Zl_Username, Sysdate, Null, '未生成任何项目', Null, 1, '已经执行|1' || Chr(9) || '已经执行');
    End If;
  Elsif 调用模式_In <> 2 Then
    Delete 病人路径执行
    Where ID = 执行记录id_In
    Returning 路径记录id, 阶段id, 天数 Into v_路径记录id, v_阶段id, v_天数;
  End If;

  If 调用模式_In = 0 And 调用场合_In = 0 Then
    Select Count(1)
    Into n_Count
    From 病人路径执行
    Where 路径记录id = v_路径记录id And 阶段id = v_阶段id And 天数 = v_天数;
    If n_Count = 0 Then
      Select Max(天数) Into v_天数 From 病人路径执行 Where 路径记录id = v_路径记录id And 阶段id = v_阶段id;
      Select Max(天数) Into v_Last天数 From 病人路径执行 Where 路径记录id = v_路径记录id;
      --记录变化的天数
      Select 当前天数 Into n_变化天数 From 病人临床路径 Where ID = v_路径记录id;
      --如果当前阶段的最后一个执行记录被删除(全部都是非必须执行的情况下)
      --由于路径跳转，一个阶段的天数可能与另一个路径的阶段交叉（例如：a路径第3阶段:3-5天,先执行第3天，跳转到其他路径后跳回来执行第5天）
      If v_天数 Is Null Or v_天数 <> v_Last天数 Then
        --a.如果当前没有任何执行记录
        If v_Last天数 Is Null Then
          Update 病人临床路径
          Set 前一阶段id = Null, 当前阶段id = Null, 当前天数 = Null, 状态 = 1
          Where ID = v_路径记录id;
          Update 病人合并路径
          Set 前一阶段id = Null, 当前阶段id = Null, 当前天数 = Null
          Where 首要路径记录id = v_路径记录id;
        Else
          --b.回退到前一个阶段
          --如果前一阶段是跳过的阶段，则直接删除
          Open c_Turn(v_路径记录id, v_Last天数, v_阶段id);
          Fetch c_Turn Bulk Collect
            Into t_阶段id;
          Close c_Turn;
          If t_阶段id.Count > 0 Then
            Forall I In 1 .. t_阶段id.Count
              Delete From 病人路径评估 Where 路径记录id = v_路径记录id And 阶段id = t_阶段id(I) And 天数 = v_Last天数;
            Forall I In 1 .. t_阶段id.Count
              Delete From 病人路径执行 Where 路径记录id = v_路径记录id And 阶段id = t_阶段id(I) And 天数 = v_Last天数;
            --删除后取最后一个阶段为前一阶段ID
            Select Max(阶段id)
            Into v_前一阶段id
            From 病人路径执行
            Where 路径记录id = v_路径记录id And 登记时间 = (Select Max(登记时间) From 病人路径执行 Where 路径记录id = v_路径记录id);
            Update 病人临床路径 Set 前一阶段id = v_前一阶段id Where ID = v_路径记录id;
          End If;
          --修改病人临床路径信息
          Select Max(阶段id)
          Into v_阶段id
          From 病人路径执行
          Where 路径记录id = v_路径记录id And
                登记时间 = (Select Max(登记时间)
                        From 病人路径执行
                        Where 路径记录id = v_路径记录id And 阶段id <> (Select 前一阶段id From 病人临床路径 Where ID = v_路径记录id));
          --重新获取当前天数
          --当前一阶段评估为：下一阶段提前至明天（时间进度=2）且第二天生成时在可选阶段中又跳过中间阶段生成后面的阶段时,
          --这种场景生成的路径表单执行取消本次生成时，v_Last天数需要重新获取）
          Select Max(天数) Into v_Last天数 From 病人路径执行 Where 路径记录id = v_路径记录id;
          --
          Update 病人临床路径
          Set 当前阶段id = 前一阶段id, 前一阶段id = v_阶段id, 当前天数 = v_Last天数, 状态 = 1
          Where ID = v_路径记录id;
        
          n_变化天数 := n_变化天数 - v_Last天数;
          --修改病人临床合并路径信息
          Select Nvl(当前阶段id, 0) Into v_阶段id From 病人临床路径 Where ID = v_路径记录id;
          Open c_Merge(v_路径记录id, v_阶段id);
          Fetch c_Merge Bulk Collect
            Into t_合并路径记录id, t_合并路径阶段id;
          Close c_Merge;
          If t_合并路径阶段id.Count > 0 Then
            Forall I In 1 .. t_合并路径阶段id.Count
              Update 病人合并路径
              Set 当前天数 = Decode(前一阶段id, Null, Null, Nvl(当前天数, 0) - Nvl(n_变化天数, 0)), 当前阶段id = 前一阶段id,
                  前一阶段id = t_合并路径阶段id(I)
              Where 首要路径记录id = v_路径记录id And ID = t_合并路径记录id(I);
          End If;
        End If;
      Else
        --如果一个阶段有多天，取消最后一个项目时，只更新天数
        Update 病人临床路径 Set 当前天数 = v_天数 Where ID = v_路径记录id And 当前天数 <> v_天数;
        n_变化天数 := n_变化天数 - v_天数;
        If n_变化天数 <> 0 Then
          --修改病人临床合并路径信息
          Select Nvl(当前阶段id, 0) Into v_阶段id From 病人临床路径 Where ID = v_路径记录id;
          Open c_Merge(v_路径记录id, v_阶段id);
          Fetch c_Merge Bulk Collect
            Into t_合并路径记录id, t_合并路径阶段id;
          Close c_Merge;
          If t_合并路径阶段id.Count > 0 Then
            Forall I In 1 .. t_合并路径阶段id.Count
              Update 病人合并路径
              Set 当前天数 = Decode(前一阶段id, Null, Null, Nvl(当前天数, 0) - Nvl(n_变化天数, 0)), 当前阶段id = 前一阶段id,
                  前一阶段id = t_合并路径阶段id(I)
              Where 首要路径记录id = v_路径记录id And ID = t_合并路径记录id(I);
          End If;
        End If;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人路径生成_Delete;
/

--92157:余伟节,2015-12-29,避免住院医嘱编辑界面生成路径医嘱时将病人临床路径的当前阶段和当前天数置为空
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
            Zl_病人路径生成_Delete(v_路径执行id,3);
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

--91385:冉俊明,2015-12-29,修改“角币二舍八入，三七做五”的误差计算规则.
Create Or Replace Procedure Zl_病人结算记录_Update
(
  结帐id_In       病人预交记录.结帐id%Type,
  保险结算_In     Varchar2, --"结算方式|结算金额||....."
  结帐_In         Number := 0,
  缺省结算方式_In Varchar2 := Null,
  缺省冲预交_In   Number := 0, --0-用现金缴款,1:剩于款项用冲预交支付(门诊预交),2-剩于款项用冲预交支付(住院预交)
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  结算卡序号_In   病人预交记录.结算卡序号%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  合作单位_In     病人预交记录.合作单位%Type := Null
  
) As
  --该游标为要删除的由费用记录产生的结算记录

  Cursor c_Del Is
    Select a.Id, a.记录性质, a.冲预交, a.结算方式, b.性质, a.预交类别
    From 病人预交记录 A, 结算方式 B
    Where a.结算方式 = b.名称 And a.结帐id = 结帐id_In;

  --相关信息
  v_No         病人预交记录.No%Type;
  v_病人id     住院费用记录.病人id%Type;
  v_主页id     住院费用记录.主页id%Type;
  v_发生时间   住院费用记录.发生时间%Type;
  v_登记时间   住院费用记录.登记时间%Type;
  v_操作员编号 住院费用记录.操作员编号%Type;
  v_操作员姓名 住院费用记录.操作员姓名%Type;

  --本次结算变量
  v_金额合计 病人预交记录.冲预交%Type;

  --保险结算
  v_保险结算 Varchar2(255);
  v_当前结算 Varchar2(50);
  v_结算方式 病人预交记录.结算方式%Type;
  v_结算金额 病人预交记录.冲预交%Type;

  v_记录性质 病人预交记录.记录性质%Type;
  v_缺省     病人预交记录.结算方式%Type;

  --分币处理及误差变量
  v_现金金额   病人预交记录.冲预交%Type;
  v_Cashcented 病人预交记录.冲预交%Type;
  v_误差金额   病人预交记录.冲预交%Type;
  v_费用id     住院费用记录.Id%Type;
  v_序号       住院费用记录.序号%Type;
  v_收费类别   住院费用记录.收费类别%Type;
  v_收费细目id 住院费用记录.收费细目id%Type;
  v_收入项目id 住院费用记录.收入项目id%Type;
  v_收据费目   住院费用记录.收据费目%Type;
  n_Noexists   Number(3);
  n_医疗小组id 住院费用记录.医疗小组id%Type;
  n_结算序号   病人预交记录.结算序号%Type;
  n_费用状态   门诊费用记录.费用状态%Type;
  n_预交金额   病人预交记录.金额%Type;
  n_当前金额   病人预交记录.金额%Type;
  v_误差项     结算方式.名称%Type;

  --临时变量
  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_组id     财务缴款分组.Id%Type;
  n_执行状态 门诊费用记录.执行状态%Type;
Begin
  --如果缺省结算方式为空，则取现金结算方式
  If 缺省结算方式_In Is Null Then
    Begin
      Select 名称 Into v_缺省 From 结算方式 Where 性质 = 1 And Rownum < 2;
    Exception
      When Others Then
        v_缺省 := '现金';
    End;
  Else
    v_缺省 := 缺省结算方式_In;
  End If;

  --取得本次结算的相关信息
  If Nvl(结帐_In, 0) = 1 Then
    Select NO, 病人id, 收费时间, 操作员编号, 操作员姓名, 缴款组id, 0
    Into v_No, v_病人id, v_登记时间, v_操作员编号, v_操作员姓名, n_组id, n_执行状态
    From 病人结帐记录
    Where ID = 结帐id_In;
  Else
    Begin
      n_Noexists := 0;
      Select NO, 病人id, 登记时间, 操作员编号, 操作员姓名, 缴款组id, 执行状态, 费用状态
      Into v_No, v_病人id, v_登记时间, v_操作员编号, v_操作员姓名, n_组id, n_执行状态, n_费用状态
      From 门诊费用记录
      Where 结帐id = 结帐id_In And Rownum < 2;
    Exception
      When Others Then
        n_Noexists := 1;
    End;
    If n_Noexists = 1 Then
      --费用记录不存在，从补充记录中找
      Select NO, 病人id, 登记时间, 操作员编号, 操作员姓名, 缴款组id, 费用状态
      Into v_No, v_病人id, v_登记时间, v_操作员编号, v_操作员姓名, n_组id, n_费用状态
      From 费用补充记录
      Where 结算id = 结帐id_In And Rownum < 2;
    End If;
    If Nvl(n_费用状态, 0) = 1 Then
      --异常单据为空:
      v_缺省 := Null;
    End If;
  
    Begin
      --20051027 陈东
      Select 记录性质
      Into v_记录性质
      From 病人预交记录
      Where 结帐id = 结帐id_In And Rownum = 1 And Mod(记录性质, 10) <> 1;
    Exception
      When Others Then
        v_记录性质 := -1;
    End;
    If v_记录性质 = -1 Then
      Begin
        Select Decode(记录性质, 1, 3, 11, 3, 4, 4, 记录性质)
        Into v_记录性质
        From 门诊费用记录
        Where 结帐id = 结帐id_In And Rownum = 1;
      Exception
        When Others Then
          --可能是卡费
          Select 记录性质 Into v_记录性质 From 住院费用记录 Where 结帐id = 结帐id_In And Rownum = 1;
      End;
    End If;
  End If;

  If Nvl(v_病人id, 0) <> 0 And Nvl(结帐_In, 0) = 1 Then
    Select 主页id Into v_主页id From 病人信息 Where 病人id = v_病人id;
  End If;
  Select 结算序号 Into n_结算序号 From 病人预交记录 Where 结帐id = 结帐id_In And Rownum = 1;

  ----回退缴款,预交不动,因为没有改冲预交的
  --收费未最未最终完成的,代表按异常单据修正,不处理人员缴款余额
  v_金额合计 := 0;
  For r_Del In c_Del Loop
    If r_Del.记录性质 Not In (1, 11) Then
      If Nvl(n_费用状态, 0) <> 1 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) - r_Del.冲预交
        Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = r_Del.结算方式;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (v_操作员姓名, r_Del.结算方式, 1, -1 * r_Del.冲预交);
        End If;
      End If;
      v_金额合计 := v_金额合计 + r_Del.冲预交;
      Delete From 病人预交记录 Where ID = r_Del.Id;
    Else
      --检查是否冲预交
      If Nvl(缺省冲预交_In, 0) <> 0 Then
        v_金额合计 := v_金额合计 + r_Del.冲预交;
        If Nvl(n_费用状态, 0) <> 1 Then
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Del.冲预交, 0)
          Where 病人id = v_病人id And 类型 = Nvl(r_Del.预交类别, 2);
          If Sql%NotFound Then
            Insert Into 病人余额
              (病人id, 性质, 预交余额, 费用余额, 类型)
            Values
              (v_病人id, 1, Nvl(r_Del.冲预交, 0), 0, Nvl(r_Del.预交类别, 2));
          End If;
        End If;
        If r_Del.记录性质 = 1 Then
          Update 病人预交记录 Set 冲预交 = 0 Where ID = r_Del.Id;
        Else
          Delete 病人预交记录 Where ID = r_Del.Id;
        End If;
      End If;
    End If;
  End Loop;

  --------------------------------------------------------------------------------------------------------------
  --------------------------------------------------------------------------------------------------------------
  --产生医保支付结算
  If 保险结算_In Is Not Null Then
    --各个保险结算
    v_保险结算 := 保险结算_In || '||';
    While v_保险结算 Is Not Null Loop
      v_当前结算 := Substr(v_保险结算, 1, Instr(v_保险结算, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
      Values
        (病人预交记录_Id.Nextval, Decode(结帐_In, 1, 2, v_记录性质), v_No, 1, v_病人id, v_主页id, '保险部份', v_结算方式, v_登记时间, v_操作员编号,
         v_操作员姓名, v_结算金额, 结帐id_In, n_组id, n_结算序号, Mod(Decode(结帐_In, 1, 2, v_记录性质), 10));
    
      v_金额合计 := v_金额合计 - v_结算金额;
    
      v_保险结算 := Substr(v_保险结算, Instr(v_保险结算, '||') + 2);
    End Loop;
  End If;
  --剩余部分用预交
  If Nvl(缺省冲预交_In, 0) <> 0 And v_金额合计 <> 0 Then
  
    n_预交金额 := v_金额合计;
    For c_预交 In (Select *
                 From (Select a.Id, a.记录状态, a.预交类别, a.No, Nvl(a.金额, 0) As 金额
                        From 病人预交记录 A,
                             (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                               From 病人预交记录 A
                               Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.病人id = v_病人id And Nvl(a.预交类别, 2) = 缺省冲预交_In
                               Group By NO
                               Having Sum(Nvl(a.金额, 0)) <> 0) B
                        Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And
                              a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And a.No = b.No And a.病人id = v_病人id And
                              Nvl(a.预交类别, 2) = 缺省冲预交_In
                        Union All
                        Select 0 As ID, 记录状态, 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
                        From 病人预交记录
                        Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And 病人id = v_病人id And
                              Nvl(预交类别, 2) = 缺省冲预交_In Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
                        Group By 记录状态, NO, 预交类别)
                 Order By ID, NO) Loop
    
      n_当前金额 := Case
                  When c_预交.金额 - n_预交金额 < 0 Then
                   c_预交.金额
                  Else
                   n_预交金额
                End;
    
      If c_预交.Id <> 0 Then
        --第一次冲预交(82592,将第一次标上结帐ID,冲预交标记为0)
        Update 病人预交记录
        Set 冲预交 = 0, 结帐id = 结帐id_In, 结算序号 = n_结算序号, 结算性质 = Mod(Decode(结帐_In, 1, 2, v_记录性质), 10)
        Where ID = c_预交.Id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
         冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
               v_登记时间, v_操作员姓名, v_操作员编号, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结算序号,
               Mod(Decode(结帐_In, 1, 2, v_记录性质), 10)
        From 病人预交记录
        Where NO = c_预交.No And 记录状态 = c_预交.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = v_病人id And 性质 = 1 And 类型 = Nvl(c_预交.预交类别, 2);
      --检查是否已经处理完
      If c_预交.金额 < n_预交金额 Then
        n_预交金额 := n_预交金额 - c_预交.金额;
      Else
        n_预交金额 := 0;
      End If;
    
      If n_预交金额 = 0 Then
        Exit;
      End If;
    End Loop;
    If n_预交金额 <> 0 Then
      v_Err_Msg := '[ZLSOFT]预交余不够支付本次支付金额,不能继续操作！[ZLSOFT]';
      Raise Err_Item;
    End If;
    Delete From 病人余额 Where 病人id = v_病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    v_金额合计 := n_预交金额;
  End If;

  --剩余部份全部用缺省结算方式结算，(小于零也不进行额外处理)
  If v_金额合计 <> 0 Then
    Update 病人预交记录
    Set 冲预交 = 冲预交 + v_金额合计, 卡类别id = 卡类别id_In, 结算卡序号 = 结算卡序号_In, 卡号 = 卡号_In, 交易流水号 = 交易流水号_In, 交易说明 = 交易说明_In,
        合作单位 = 合作单位_In, 结算序号 = n_结算序号
    
    Where 结帐id = 结帐id_In And Nvl(结算方式, 'LXH_Test') = Nvl(v_缺省, 'LXH_Test') And 记录性质 = Decode(结帐_In, 1, 2, v_记录性质);
    If Sql%RowCount = 0 Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 卡类别id, 结算卡序号, 卡号, 交易流水号,
         交易说明, 合作单位, 结算性质)
      Values
        (病人预交记录_Id.Nextval, Decode(结帐_In, 1, 2, v_记录性质), v_No, 1, v_病人id, v_主页id, '保险结算修正', v_缺省, v_登记时间, v_操作员编号,
         v_操作员姓名, v_金额合计, 结帐id_In, n_组id, n_结算序号, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In,
         Mod(Decode(结帐_In, 1, 2, v_记录性质), 10));
    End If;
  
    --挂号结算,分币处理(由于挂号界面没有预结算,所以在此过程中根据分币处理规则来修正)
    If v_记录性质 = 4 Then
    
      Begin
        Select a.冲预交
        Into v_现金金额
        From 病人预交记录 A, 结算方式 B
        Where a.结算方式 = b.名称 And b.性质 = 1 And a.结帐id = 结帐id_In;
      Exception
        When Others Then
          v_现金金额 := 0;
      End;
      If Floor(Abs(v_现金金额) * 10) <> Abs(v_现金金额) * 10 Then
        --误差处理
        v_Cashcented := Zl_Cent_Money(v_现金金额);
        v_误差金额   := v_Cashcented - v_现金金额;
        If v_误差金额 <> 0 Then
          If n_结算序号 < 0 Then
            --10.34之后误差数据
            Begin
              Select 名称 Into v_误差项 From 结算方式 Where 性质 = 9;
            Exception
              When Others Then
                v_Err_Msg := '不能正确读取误差项的信息，请先检查结算方式中误差费是否设置正确。';
                Raise Err_Item;
            End;
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
            Values
              (病人预交记录_Id.Nextval, Decode(结帐_In, 1, 2, v_记录性质), v_No, 1, v_病人id, v_主页id, '误差费', v_误差项, v_登记时间, v_操作员编号,
               v_操作员姓名, v_误差金额, 结帐id_In, n_组id, n_结算序号, Mod(Decode(结帐_In, 1, 2, v_记录性质), 10));
          Else
            --1.更新预交记录(一定存在记录)
            Update 病人预交记录
            Set 冲预交 = v_Cashcented
            Where 结算方式 = (Select 名称 From 结算方式 Where 性质 = 1 And Rownum = 1) And 结帐id = 结帐id_In;
          
            --2.生成误差费用记录(注:计算单位记录的是号别,所以不取误差项的)
            Begin
              Select a.类别, a.Id, c.Id, c.收据费目
              Into v_收费类别, v_收费细目id, v_收入项目id, v_收据费目
              From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费特定项目 D
              Where d.特定项目 = '误差项' And d.收费细目id = a.Id And a.Id = b.收费细目id And b.收入项目id = c.Id And
                    Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'));
            Exception
              When Others Then
                v_Err_Msg := '不能正确读取收费误差项的信息，请先检查该项目是否设置正确。';
                Raise Err_Item;
            End;
            If Nvl(结帐_In, 0) = 1 Then
              Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
              Select Max(序号) + 1, Max(发生时间) Into v_序号, v_发生时间 From 住院费用记录 Where 结帐id = 结帐id_In;
              n_医疗小组id := Zl_医疗小组_Get(0, v_操作员姓名, v_病人id, v_主页id, v_发生时间);
            
              Insert Into 住院费用记录
                (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 床号, 姓名, 性别, 年龄, 病人病区id, 病人科室id, 费别, 收费类别,
                 收费细目id, 计算单位, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间,
                 登记时间, 执行部门id, 执行人, 执行状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 是否上传, 缴款组id, 医疗小组id)
                Select v_费用id, 记录性质, NO, 实际票号, 记录状态, v_序号, Null, Null, 门诊标志, 病人id, 标识号, 床号, 姓名, 性别, 年龄, 病人病区id, 病人科室id,
                       费别, v_收费类别, v_收费细目id, 计算单位, 发药窗口, 1, 1, 加班标志, 9, v_收入项目id, v_收据费目, v_误差金额, v_误差金额, v_误差金额, 记帐费用,
                       划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 结帐id_In, v_误差金额, 操作员编号, 操作员姓名, 1, 缴款组id,
                       Decode(n_医疗小组id, Null, 医疗小组id, n_医疗小组id)
                From 住院费用记录
                Where 结帐id = 结帐id_In And Rownum = 1;
            Else
              Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
              Select Max(序号) + 1 Into v_序号 From 门诊费用记录 Where 结帐id = 结帐id_In;
              Insert Into 门诊费用记录
                (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id,
                 计算单位, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
                 执行部门id, 执行人, 执行状态, 费用状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 是否上传, 缴款组id)
                Select v_费用id, 记录性质, NO, 实际票号, 记录状态, v_序号, Null, Null, 门诊标志, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 病人科室id, 费别,
                       v_收费类别, v_收费细目id, 计算单位, 发药窗口, 1, 1, 加班标志, 9, v_收入项目id, v_收据费目, v_误差金额, v_误差金额, v_误差金额, 记帐费用, 划价人,
                       开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 费用状态, 结帐id_In, v_误差金额, 操作员编号, 操作员姓名, 1, 缴款组id
                From 门诊费用记录
                Where 结帐id = 结帐id_In And Rownum = 1;
            End If;
          End If;
          --3.更新汇总表
          --只可能产生误差金额的变化.仅为了变量处理方便而用游标
        End If;
      End If;
    End If;
  End If;

  --最后再处理"人员缴款余额"(没有动冲预交那部分,所以"病人余额"的预交余额不用更新)
  For r_Del In c_Del Loop
    If r_Del.记录性质 Not In (1, 11) Then
      If Nvl(n_费用状态, 0) <> 1 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + r_Del.冲预交
        Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = r_Del.结算方式;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (v_操作员姓名, r_Del.结算方式, 1, r_Del.冲预交);
        End If;
      End If;
    End If;
  End Loop;
  Delete From 人员缴款余额 Where 性质 = 1 And 收款员 = v_操作员姓名 And Nvl(余额, 0) = 0;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结算记录_Update;
/

--91385:冉俊明,2015-12-29,修改“角币二舍八入，三七做五”的误差计算规则.
Create Or Replace Function Zl_Cent_Money
(
  Money_In In Number,
  Type_In  In Number := 2
) Return Number As
  n_Sign Integer;
  n_Temp Number(16, 5);
  n_金额 Number(16, 5);
  n_Mode Number(1);
Begin
  --         0.不处理
  --         1.采取四舍五入法,eg:0.51=0.50;0.56=0.60
  --         2.补整收法,eg:0.51=0.60,0.56=0.60
  --         3.舍分收法,eg:0.51=0.50,0.56=0.50
  --        4.四舍六入五成双,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
  --           四舍六入五成双,详见我国科学技术委员会正式颁布的《数字修约规则》,但根据vb的Round函数,若被舍弃的数字包括几位数字时，不对该数字进行连续修约 
  --           即银行家舍入法:四舍六入五考虑，五后非零就进一，五后皆零看奇偶，五前为偶应舍去，五前为奇要进一
  --         5.三七作五、二舍八入,对角进行处理，不需要先对分币进行舍入,即0.24(含)以下都舍掉角，0.75(含)以上都进角，0.25-0.74处理为0.5。
  --         6.五舍六入:eg:0.15=0.10:0.16=0.2:   刘兴洪 问题:34519  日期:2010-12-06 09:58:02

  n_Mode := To_Number(Substr(Nvl(zl_GetSysParameter(14) || '000', '000'), Type_In, 1));
  n_Sign := Sign(Money_In);
  n_金额 := Abs(Money_In);
  If n_Mode = 1 Then
    --1.四舍五入法,eg:0.51=0.50;0.56=0.60
    n_Temp := n_Sign * Round(n_金额, 1);
    Return n_Temp;
  End If;
  If n_Mode = 2 Then
    ----2.补整收法,eg:0.51=0.60,0.56=0.60
    n_Temp := n_Sign * Ceil(n_金额 * 10) / 10;
    Return n_Temp;
  End If;
  If n_Mode = 3 Then
    ----3.舍分收法,eg:0.51=0.50,0.56=0.50
    n_Temp := n_Sign * Floor(n_金额 * 10) / 10;
    Return n_Temp;
  End If;
  If n_Mode = 4 Then
    ----4.四舍六入五成双,由于Oracle没有相关函数,算法复杂,暂不支持
    n_Temp := n_Sign * n_金额;
    Return n_Temp;
  End If;
  If n_Mode = 5 Then
    ----5.三七作五、二舍八入,eg:0.29=0,0.30=0.50,0.79=0.50,0.80=1.00
    n_Temp := Round(n_金额 - Floor(n_金额), 1);
    If n_Temp >= 0.8 Then
      n_Temp := 1;
    Elsif n_Temp < 0.3 Then
      n_Temp := 0;
    Else
      n_Temp := 0.5;
    End If;
    n_Temp := Floor(n_金额) + n_Temp; --5.三七作五、二舍八入,eg:0.24=0,0.25=0.50,0.74=0.50,0.75=1.00
    n_Temp := n_Sign * n_Temp;
    Return n_Temp;
  End If;
  If n_Mode = 6 Then
    ----6.五舍六入
    n_Temp := n_Sign * Round(n_金额 - 0.01, 1);
    Return n_Temp;
  End If;
  Return Money_In;
Exception
  When Others Then
    Return Null;
End Zl_Cent_Money;
/

--91665:冉俊明,2015-12-29,增加多单据分单据结算时医保结算失败时只对结算成功单据收费的模式。
Create Or Replace Procedure Zl_门诊收费票据_Insert
(
  No_In           Varchar2,
  票据号_In       票据使用明细.号码%Type,
  领用id_In       票据使用明细.领用id%Type,
  使用人_In       票据使用明细.使用人%Type,
  使用时间_In     票据使用明细.使用时间%Type,
  打印id_In       票据打印内容.Id%Type := 0,
  票据张数_In     Number := 1,
  医保接口打印_In Number := 0
) As
  --功能：处理门诊收费票据的发出
  --参数：
  --      NO_IN       =     收费的单据号,可能是多张单据同时收费。格式为：A0000001,A0000002,....
  --      票据号_IN   =     要使用的开始票据号。该票据号应该不为空，否则不用处理票据，也不能区分多张一起收费的单据。
  --      领用ID_IN   =     严格控制票据时，为使用票据的领用批次。非严格控制时，为NULL。
  --      打印ID_IN   =     当修改多单据中的一张时,为了便于整体重打,将该单据的打印内容填写为与退费单据相同,不单独新发出票据,由退费重打发出
  --      票据张数_In =     实际所需的票据打印张数
  --      医保接口打印_In = 是否医保接口打印先存入票据数据，若是将传入打印id_In

  --该游标用于票据范围判断
  Cursor c_Fact Is
    Select * From 票据领用记录 Where ID = Nvl(领用id_In, 0);
  r_Factrow c_Fact%RowType;

  v_票据号     票据使用明细.号码%Type;
  v_当前票据号 票据使用明细.号码%Type;
  v_打印id     票据打印内容.Id%Type;

  v_当前号 门诊费用记录.No%Type;
  v_单据号 Varchar2(1000);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(打印id_In, 0) = 0 Or Nvl(医保接口打印_In, 0) = 1 Then
    --无票据号时,不用处理票据
    If 票据号_In Is Null Then
      Return;
    End If;
    v_打印id := Nvl(打印id_In, 0);
    If Nvl(v_打印id, 0) = 0 Then
      Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
    End If;
  
    --生成单据的票据打印内容
    v_单据号 := No_In || ',';
    While v_单据号 Is Not Null Loop
      v_当前号 := Substr(v_单据号, 1, Instr(v_单据号, ',') - 1);
      --票据打印内容
      Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 1, v_当前号);
      --门诊费用记录中填写开始票据号以便显示
      Update 门诊费用记录 Set 实际票号 = 票据号_In Where 记录性质 = 1 And NO = v_当前号;
      v_单据号 := Substr(v_单据号, Instr(v_单据号, ',') + 1);
    End Loop;
  
    --并发出票据
    v_票据号 := 票据号_In;
    If Nvl(领用id_In, 0) <> 0 Then
      Open c_Fact;
      Fetch c_Fact
        Into r_Factrow;
      If c_Fact%RowCount = 0 Then
        v_Error := '无效的票据领用批次，无法完成收费票据分配操作。';
        Close c_Fact;
        Raise Err_Custom;
      Elsif Nvl(r_Factrow.剩余数量, 0) < 票据张数_In Then
        v_Error := '当前批次的剩余数量不足' || 票据张数_In || '张，无法完成收费票据分配操作。';
        Close c_Fact;
        Raise Err_Custom;
      End If;
    End If;
    For I In 1 .. 票据张数_In Loop
      --检查票据范围是否正确
      If Nvl(领用id_In, 0) <> 0 Then
        If Not (Upper(v_票据号) >= Upper(r_Factrow.开始号码) And Upper(v_票据号) <= Upper(r_Factrow.终止号码) And
            Length(v_票据号) = Length(r_Factrow.终止号码)) Then
          v_Error := '该单据需要打印多张票据,但票据号"' || v_票据号 || '"超出票据领用的号码范围！';
          Close c_Fact;
          Raise Err_Custom;
        End If;
      End If;
    
      --发出票据
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
      Values
        (票据使用明细_Id.Nextval, 1, v_票据号, 1, 1, 领用id_In, v_打印id, 使用人_In, 使用时间_In);
    
      v_当前票据号 := v_票据号;
      --下一个票据号
      v_票据号 := Zl_Incstr(v_票据号);
    End Loop;
  
    If Nvl(领用id_In, 0) <> 0 Then
      Update 票据领用记录
      Set 使用时间 = 使用时间_In, 当前号码 = v_当前票据号, 剩余数量 = Nvl(剩余数量, 0) - 票据张数_In
      Where ID = 领用id_In;
    
      Close c_Fact;
    End If;
  Else
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (打印id_In, 1, No_In);
    If 票据号_In Is Null Then
      Return;
    End If;
    --门诊费用记录中填写开始票据号以便显示
    Update 门诊费用记录 Set 实际票号 = 票据号_In Where 记录性质 = 1 And NO = No_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费票据_Insert;
/

--92335:李南春,2016-01-18,三方支付新模式及过程拆分
--91665:冉俊明,2015-12-29,增加多单据分单据结算时医保结算失败时只对结算成功单据收费的模式。
Create Or Replace Procedure Zl_门诊收费结算_Modify
(
  操作类型_In     Number,
  病人id_In       门诊费用记录.病人id%Type,
  结帐id_In       病人预交记录.结帐id%Type,
  结算方式_In     Varchar2,
  冲预交_In       病人预交记录.冲预交%Type := Null,
  退支票额_In     病人预交记录.冲预交%Type := Null,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  缴款_In         病人预交记录.缴款%Type := Null,
  找补_In         病人预交记录.找补%Type := Null,
  误差金额_In     门诊费用记录.实收金额%Type := Null,
  完成结算_In     Number := 0,
  缺省结算方式_In 结算方式.名称%Type := Null,
  更新交款余额_In  Number := 0--是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:收费结算时,修改结算的相关信息
  --操作类型_In:
  --   0-普通收费方式:
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
  --     ②退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
  --   1.三方卡结算:
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
  --     ②退支票额_In:传入零
  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.." 
  --     ②退支票额_In:传入零
  --   3-消费卡结算:
  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."  消费卡ID:为零时,根据卡号自动定位
  --     ②冲预交_In: 传入零
  --     ②退支票额_In:传入零
  -- 冲预交_In: 存在冲预交时,传入
  -- 误差金额_In:存在误差费时,传入
  -- 完成结算_In:1-完成收费;0-未完成收费
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_结算内容 Varchar2(500);
  v_当前结算 Varchar2(50);
  v_卡号     病人医疗卡信息.卡号%Type;
  n_消费卡id 消费卡目录.Id%Type;
  n_卡类别id 病人预交记录.结算卡序号%Type;
  v_名称     Varchar2(100);
  n_自制卡   卡消费接口目录.自制卡%Type;
  n_序号     病人卡结算记录.序号%Type;
  n_Id       病人卡结算记录.Id%Type;
  n_预交id   病人预交记录.Id%Type;
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  n_返回值   人员缴款余额.余额%Type;
  n_预交金额 病人预交记录.冲预交%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  v_退支票   病人预交记录.结算方式%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;
  v_误差费   结算方式.名称%Type;
  n_Count    Number;
  n_Havenull Number;
  l_预交id   t_Numlist := t_Numlist();

  Cursor c_Feedata Is
    Select Max(m.病人id) As 病人id, Max(m.登记时间) As 登记时间, Max(m.操作员编号) As 操作员编号, Max(m.操作员姓名) As 操作员姓名, Sum(结帐金额) As 结算金额,
           Max(m.缴款组id) As 缴款组id
    From 门诊费用记录 M
    Where m.结帐id = 结帐id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号
    From 病人预交记录
    Where 结帐id = 结帐id_In And 结算方式 Is Null;
  r_Balancedata c_Balancedata%RowType;

Begin
  Begin
    Select 名称 Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
  Exception
    When Others Then
      v_误差费 := '误差费';
  End;

  --0.正式结算
  Select Count(1), Max(Decode(结算方式, Null, 1, 0))
  Into n_Count, n_Havenull
  From 病人预交记录
  Where 结帐id = 结帐id_In;

  --1.增加结算方式为空的结算数据
  n_结算金额 := 0;
  n_Count    := 0;
  Open c_Feedata;
  Begin
    Fetch c_Feedata
      Into r_Feedata;
    --修正或新增结算方式为null的记录
    Select Nvl(Sum(冲预交), 0) Into n_结算金额 From 病人预交记录 Where 结帐id = 结帐id_In;
    If n_Havenull = 0 Or Round(Nvl(r_Feedata.结算金额, 0), 6) <> Round(Nvl(n_结算金额, 0), 6) Then
      --先删除存在的结算方式为null的记录
      Delete From 病人预交记录 Where 结帐id = 结帐id_In And 结算方式 Is Null;
      Select Nvl(Sum(冲预交), 0) Into n_结算金额 From 病人预交记录 Where 结帐id = 结帐id_In;
    
      n_结算金额 := Round(Nvl(r_Feedata.结算金额, 0) - n_结算金额, 6);
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, Decode(病人id_In, 0, Null, 病人id_In), Null, r_Feedata.登记时间, r_Feedata.操作员编号,
         r_Feedata.操作员姓名, n_结算金额, 结帐id_In, r_Feedata.缴款组id, -1 * 结帐id_In, 1, 3);
    End If;
  Exception
    When Others Then
      n_Count := 1;
  End;
  Close c_Feedata;
  If n_Count = 1 Then
    v_Err_Msg := '未找到指定的收费明细数据,结算操作失败！';
    Raise Err_Item;
  End If;

  If 操作类型_In = 0 And Nvl(退支票额_In, 0) <> 0 Then
    Begin
      Select b.名称
      Into v_退支票
      From 结算方式应用 A, 结算方式 B
      Where a.应用场合 = '收费' And b.名称 = a.结算方式 And Nvl(b.应付款, 0) = 1 And Rownum <= 1;
    Exception
      When Others Then
        v_退支票 := '无';
    End;
    If v_退支票 = '无' Then
      v_Err_Msg := '在结算场合中,不存在结算性质为应付款的结算方式,请在[结算方式]中设置！';
      Raise Err_Item;
    End If;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If Nvl(误差金额_In, 0) <> 0 Then
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + Nvl(误差金额_In, 0)
    Where 结帐id = 结帐id_In And 结算方式 = v_误差费;
    If Sql%NotFound Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_误差费, r_Balancedata.收款时间, r_Balancedata.操作员编号,
         r_Balancedata.操作员姓名, 误差金额_In, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null, Null, 卡号_In,
         交易流水号_In, 交易说明_In, Null, 3);
    End If;
    Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(误差金额_In, 0) Where 结帐id = 结帐id_In And 结算方式 Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
      Raise Err_Item;
    End If;
  End If;

  --预交款处理
  If Nvl(冲预交_In, 0) <> 0 Then
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '不能确定病人的病人ID,收费不能使用预交款结算,结算操作失败！';
      Raise Err_Item;
    End If;
  
    --病人余额检查
    Begin
      Select Nvl(预交余额, 0) - Nvl(费用余额, 0)
      Into n_预交金额
      From 病人余额
      Where 病人id = 病人id_In And Nvl(性质, 0) = 1 And 类型 = 1;
    Exception
      When Others Then
        n_预交金额 := 0;
    End;
    If n_预交金额 < 冲预交_In Then
      v_Err_Msg := '病人的当前预交余额为 ' || LTrim(To_Char(n_预交金额, '9999999990.00')) || '，小于本次支付金额 ' ||
                   LTrim(To_Char(冲预交_In, '9999999990.00')) || '，支付失败！';
      Raise Err_Item;
    End If;
  
    n_预交金额 := 冲预交_In;
  
    For c_冲预交 In (Select *
                  From (Select a.Id, a.记录状态, a.No, Nvl(a.金额, 0) As 金额
                         From 病人预交记录 A,
                              (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                                From 病人预交记录 A
                                Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.病人id = 病人id_In And 预交类别 = 1
                                Group By NO
                                Having Sum(Nvl(a.金额, 0)) <> 0) B
                         Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.No = b.No And a.病人id = 病人id_In And a.预交类别 = 1
                         Union All
                         Select 0 As ID, 记录状态, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
                         From 病人预交记录
                         Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And 病人id = 病人id_In And
                               预交类别 = 1 Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
                         Group By 记录状态, NO)
                  Order By ID, NO) Loop
    
      If c_冲预交.金额 - n_预交金额 < 0 Then
        n_冲预交 := c_冲预交.金额;
      Else
        n_冲预交 := n_预交金额;
      End If;
    
      If c_冲预交.Id <> 0 Then
        --第一次冲预交(将第一次标上结帐ID,冲预交标记为0)
        Update 病人预交记录
        Set 冲预交 = 0, 结帐id = 结帐id_In, 结算序号 = -1 * 结帐id_In, 结算性质 = 3
        Where ID = c_冲预交.Id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
               r_Balancedata.收款时间, r_Balancedata.操作员姓名, r_Balancedata.操作员编号, n_冲预交, 结帐id_In, r_Balancedata.缴款组id, 预交类别,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * 结帐id_In, 3
        From 病人预交记录
        Where NO = c_冲预交.No And 记录状态 = c_冲预交.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新数据(结算方式为NULL的)
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_冲预交
      Where 结帐id = 结帐id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
      --检查是否已经处理完
      If c_冲预交.金额 < n_预交金额 Then
        n_预交金额 := n_预交金额 - c_冲预交.金额;
      Else
        n_预交金额 := 0;
      End If;
      If n_预交金额 = 0 Then
        Exit;
      End If;
    
    End Loop;
    --检查金额是否足够
    If Abs(n_预交金额) > 0 Then
      v_Err_Msg := '病人的当前预交余额小于本次支付金额 ' || LTrim(To_Char(冲预交_In, '9999999990.00')) || '，支付失败！';
      Raise Err_Item;
    End If;
  
    --更新病人预交余额
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) - 冲预交_In
    Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (病人id_In, 1, -1 * 冲预交_In, 1);
      n_返回值 := -1 * 冲预交_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  End If;
  If 操作类型_In = 0 Then
  
    If Nvl(退支票额_In, 0) <> 0 Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_退支票, r_Balancedata.收款时间, r_Balancedata.操作员编号,
         r_Balancedata.操作员姓名, 退支票额_In, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null, Null, 卡号_In,
         交易流水号_In, 交易说明_In, Null, 3);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - 退支票额_In Where 结帐id = 结帐id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    End If;
  
    --各个收费结算 :格式为:"结算方式|结算金额|结算号码|结算摘要||.."
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
           交易流水号, 交易说明, 结算号码, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
           r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号,
           2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 3);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
          Raise Err_Item;
        End If;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 操作类型_In = 1 Then
    --三方卡结算交易
  
    v_当前结算 := 结算方式_In;
  
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    If Nvl(n_结算金额, 0) <> 0 Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
         r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号,
         2, 卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 3);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --2.医保结算(调用此过程,采取平均分摊的方式分摊结算情况):这种情况医保结处后,必须全退
  If 操作类型_In = 2 Then
    --2.1检查是否已经存在医保结算数据,存在先删除
    n_结算金额 := 0;
    For v_医保 In (Select ID, 结算方式, 冲预交
                 From 病人预交记录 A
                 Where 结帐id = 结帐id_In And Exists
                  (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4)) And Mod(记录性质, 10) <> 1)
    
     Loop
      n_结算金额 := n_结算金额 + Nvl(v_医保.冲预交, 0);
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := v_医保.Id;
    End Loop;
    If Nvl(n_结算金额, 0) <> 0 Then
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 结帐id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    End If;
    If l_预交id.Count <> 0 Then
      Forall I In 1 .. l_预交id.Count
        Delete 病人预交记录 Where ID = l_预交id(I);
    End If;
  
    If 结算方式_In Is Not Null Then
      v_结算内容 := 结算方式_In || '||';
    End If;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      For c_结算信息 In (Select 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号
                     From 病人预交记录
                     Where 结帐id = 结帐id_In And 结算方式 Is Null) Loop
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 1, c_结算信息.病人id, Null, '保险结算', v_结算方式, c_结算信息.收款时间, c_结算信息.操作员编号, c_结算信息.操作员姓名,
           n_结算金额, c_结算信息.结帐id, c_结算信息.缴款组id, c_结算信息.结算序号, 1, 3);
      End Loop;
      --更新数据(结算方式为NULL的)
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_结算金额
      Where 结帐id = 结帐id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --3-消费卡批量结算
  If 操作类型_In = 3 Then
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      --卡类别ID|卡号|消费卡ID|消费金额
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(v_当前结算);
      Begin
        Select 名称, 自制卡, 结算方式 Into v_名称, n_自制卡, v_结算方式 From 卡消费接口目录 Where 编号 = 卡类别id_In;
      Exception
        When Others Then
          v_名称 := Null;
      End;
      If v_名称 Is Null Then
        v_Err_Msg := '未找到对应的结算卡接口,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := v_名称 || '未设置对应的结算方式,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If Nvl(n_消费卡id, 0) = 0 Then
        --未传入消费卡ID时,以卡号为准进行查找(卡号的合法性,在程序中有判断)
        Begin
          Select ID
          Into n_消费卡id
          From 消费卡目录
          Where 接口编号 = n_卡类别id And 卡号 = v_卡号 And
                序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = n_卡类别id And 卡号 = v_卡号);
        Exception
          When Others Then
            n_消费卡id := 0;
        End;
        If Nvl(n_消费卡id, 0) = 0 Then
          v_Err_Msg := '未找到卡号为:' || v_卡号 || '的' || v_名称 || '.,本次刷卡消费失败!';
          Raise Err_Item;
        End If;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Then
      
        Update 病人预交记录
        Set 冲预交 = Nvl(冲预交, 0) + n_结算金额
        Where 结帐id = r_Balancedata. 结帐id And 结算方式 = v_结算方式 And 结算卡序号 = n_卡类别id
        Returning ID Into n_预交id;
        If Sql%NotFound Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算卡序号, 校对标志, 结算性质)
          Values
            (n_预交id, 3, Null, 1, r_Balancedata. 病人id, Null, Null, v_结算方式, r_Balancedata. 收款时间, r_Balancedata. 操作员编号,
             r_Balancedata. 操作员姓名, n_结算金额, r_Balancedata. 结帐id, r_Balancedata. 缴款组id, r_Balancedata. 结算序号, n_卡类别id, 2, 3);
        End If;
      
        --插入卡结算记录
        Begin
          Select Nvl(Max(Nvl(序号, 0)), 0) + 1
          Into n_序号
          From 病人卡结算记录
          Where 接口编号 = n_卡类别id And Nvl(消费卡id, 0) = Nvl(n_消费卡id, 0) And 卡号 = v_卡号;
        Exception
          When Others Then
            n_序号 := 1;
        End;
      
        Select 病人卡结算记录_Id.Nextval Into n_Id From Dual;
      
        Insert Into 病人卡结算记录
          (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
        Values
          (n_Id, n_卡类别id, n_消费卡id, n_序号, 1, v_结算方式, n_结算金额, v_卡号, Null, r_Balancedata. 收款时间, Null, 0);
        --如果消费卡,需同时更改其余额
        If Nvl(n_消费卡id, 0) <> 0 Then
          Update 消费卡目录 Set 余额 = 余额 - n_结算金额 Where ID = n_消费卡id;
          If Sql%NotFound Then
            v_Err_Msg := '卡号为' || v_卡号 || '的' || v_名称 || '未找到!';
            Raise Err_Item;
          End If;
        End If;
        Insert Into 病人卡结算对照 (预交id, 卡结算id) Values (n_预交id, n_Id);
      
        --更新数据(结算方式为NULL的)
        Update 病人预交记录
        Set 冲预交 = 冲预交 - n_结算金额
        Where 结帐id = r_Balancedata. 结帐id And 结算方式 Is Null And Nvl(校对标志, 0) = 1
        Returning Nvl(冲预交, 0) Into n_返回值;
      
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If Nvl(完成结算_In, 0) = 0 Then
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --完成收费,需要处理人员缴款余额,预交记录(结算方式=NULL)

  --1.删除结算方式为NULL的预交记录
  Delete 病人预交记录 Where 结帐id = 结帐id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未缴款的数据,不能完成结算!';
    Else
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！!';
    End If;
    Raise Err_Item;
  End If;

  --结算金额为零时，增加一条金额为0的病人预交记录
  Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In;

  If n_Count = 0 Then
    v_结算方式 := 缺省结算方式_In;
    If v_结算方式 Is Null Then
      Begin
        Select 结算方式 Into v_结算方式 From 结算方式应用 Where 应用场合 = '收费' And Nvl(缺省标志, 0) = 1;
      Exception
        When Others Then
          v_结算方式 := Null;
      End;
      If v_结算方式 Is Null Then
        Begin
          Select 名称 Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1;
        Exception
          When Others Then
            v_结算方式 := '现金';
        End;
      End If;
    End If;
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
       交易流水号, 交易说明, 结算号码, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
       r_Balancedata.操作员姓名, 0, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null, Null, Null, Null,
       交易说明_In, Null, 3);
  End If;

  --2.处理缴款数据和找补数据及校对标志更新为0
  Update 病人预交记录 Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0 Where 结帐id = 结帐id_In;

  --3.更新费用状态
  Update 门诊费用记录 Set 费用状态 = 0 Where 结帐id = 结帐id_In;

  --4.更新人员缴款数据
  If Nvl(更新交款余额_In,0)=0 then
    For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
                 From 病人预交记录 A
                 Where a.结帐id = 结帐id_In And Mod(a.记录性质, 10) <> 1
                 Group By 结算方式, 操作员姓名) Loop
    
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
      Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
      End If;
    End Loop;
  End if;
  --收费后产生导引
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 4, 结帐id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费结算_Modify;
/

--92177:刘尔旋,2015-12-29,医保校对时主页ID问题
Create Or Replace Procedure Zl_住院收费结算_Update
(
  结帐id_In       住院费用记录.结帐id%Type,
  结帐结算_In     Varchar2, --结帐结算_IN-非医保时:结算方式|结算金额|结算号码||.....医保时:结算方式|结算金额|保险类别,保险密码,保险帐号||.....
  冲预交_In       Varchar2, --冲预交_IN= ID|单据号|金额|记录状态||.....
  缴款_In         病人预交记录.缴款%Type := Null,
  找补_In         病人预交记录.找补%Type := Null,
  三方帐户结算_In Varchar2 := Null --:结算方式|结算金额|卡类别ID|卡号|交易流水号|交易说明||...
) As
  --功能:处理结帐时和医保正式结算后,相关结算信息的调整
  --     因为虚拟结帐后,生成的医保结算金额总额及分摊可能会与正式结算时有差异,所以提供了校对功能,
  --   操作员在结算校对时,可以调整非医保结算方式的各种结算金额及方式,重新生成结算串,并且可能产生误差金额.

  --病人信息
  Cursor c_Pati(v_病人id 病人信息.病人id%Type) Is
    Select a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, b.主页id, b.出院病床, b.当前病区id, b.出院科室id, Nvl(b.费别, a.费别) As 费别, c.编码 As 付款方式
    From 病人信息 A, 病案主页 B, 医疗付款方式 C, (Select Max(主页id) As 主页id From 住院费用记录 Where 结帐id = 结帐id_In) D
    Where a.病人id = v_病人id And a.病人id = b.病人id(+) And b.主页id = Nvl(d.主页id, 0) And a.医疗付款方式 = c.名称(+);
  r_Pati c_Pati%RowType;

  --过程变量
  v_结算内容 Varchar2(4000);
  v_当前结算 Varchar2(100);
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 Varchar2(100); --保险结算记录时,存入:保险类别,保险密码,保险帐号

  n_卡类别id   病人预交记录.卡类别id%Type;
  v_卡号       病人预交记录.卡号%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  v_交易说明   病人预交记录.交易说明%Type;

  v_No         病人预交记录.No%Type;
  v_病人id     病人预交记录.病人id%Type;
  v_收款时间   病人预交记录.收款时间%Type;
  v_操作员编号 病人预交记录.操作员编号%Type;
  v_操作员姓名 病人预交记录.操作员姓名%Type;
  v_误差结算   病人预交记录.结算方式%Type;

  v_预交id   病人预交记录.Id%Type;
  v_记录状态 病人预交记录.记录状态%Type;

  v_保险类别 病人预交记录.缴款单位%Type;
  v_保险帐号 病人预交记录.单位开户行%Type;
  v_保险密码 病人预交记录.单位帐号%Type;
  v_付款方式 门诊费用记录.付款方式%Type;
  n_返回值   病人余额.预交余额%Type;
  n_Dele     Number; --0-不删除,1-删除
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_门诊标志   Number; --1.门诊,2-住院,3-门诊和住院
  n_组id       财务缴款分组.Id%Type;
  n_类别       Number;
  n_门诊冲预交 病人预交记录.冲预交%Type;
  n_住院冲预交 病人预交记录.冲预交%Type;

Begin

  --1.取预交记录中的需要的相关信息
  Select NO, 病人id, 收费时间, 操作员编号, 操作员姓名, 缴款组id
  Into v_No, v_病人id, v_收款时间, v_操作员编号, v_操作员姓名, n_组id
  From 病人结帐记录
  Where ID = 结帐id_In;

  Open c_Pati(v_病人id);
  Fetch c_Pati
    Into r_Pati;

  --误差相关信息
  Begin
    Select 名称 Into v_误差结算 From 结算方式 Where 性质 = 9;
  Exception
    When Others Then
      Begin
        v_Error := '不能正确读取收费误差项的信息，请先检查该项目是否设置正确。';
        Raise Err_Custom;
      End;
  End;

  --2.删除旧的记录,回退汇总数据
  --回退人员缴款余额,病人余额,
  For c_Del In (Select 结算方式, 操作员姓名, 冲预交 From 病人预交记录 Where 结帐id = 结帐id_In And 记录性质 = 2) Loop
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) - Nvl(c_Del.冲预交, 0)
    Where 结算方式 = c_Del.结算方式 And 收款员 = v_操作员姓名 And 性质 = 1;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (v_操作员姓名, c_Del.结算方式, 1, -1 * c_Del.冲预交);
    End If;
  End Loop;

  If v_病人id > 0 Then
    For v_预交 In (Select 预交类别, Sum(Nvl(冲预交, 0)) As 预交金额
                 From 病人预交记录
                 Where 结帐id = 结帐id_In And 记录性质 In (1, 11)
                 Group By 预交类别
                 Having Sum(Nvl(冲预交, 0)) <> 0) Loop
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_预交.预交金额, 0)
      Where 病人id = v_病人id And 类型 = Nvl(v_预交.预交类别, 2) And 性质 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 类型, 预交余额, 性质)
        Values
          (v_病人id, Nvl(v_预交.预交类别, 2), Nvl(v_预交.预交金额, 0), 1);
      End If;
    End Loop;
  End If;

  --取门诊标志
  Select Case
           When 门诊标志 = 1 And 住院标志 = 1 Then
            3
           When 门诊标志 = 1 Then
            1
           Else
            2
         End, 付款方式
  Into n_门诊标志, v_付款方式
  From (Select Nvl(Max(门诊标志), 0) As 门诊标志, Nvl(Max(住院标志), 0) As 住院标志, Max(付款方式) As 付款方式
         From (Select 1 As 门诊标志, 0 As 住院标志, 付款方式
                From 门诊费用记录
                Where 结帐id = 结帐id_In And Rownum = 1
                Union All
                Select 0 As 门诊标志, 1 As 住院标志, '' As 付款方式
                From 住院费用记录
                Where 结帐id = 结帐id_In And Rownum = 1));

  --回退汇总表.         病人未结费用(因为新误差将立即结帐,所以不处理)
  --只可能产生误差金额的变化. 旧误差只可能存在一行,仅为了变量处理方便而用游标

  --删除结帐缴款,保险结算记录
  Delete 三方结算交易 Where 交易id In (Select ID From 病人预交记录 Where 结帐id = 结帐id_In And 记录性质 = 2);

  Delete 病人预交记录 Where 结帐id = 结帐id_In And 记录性质 = 2;

  --第一次冲预交的,清空冲减额
  Update 病人预交记录 Set 冲预交 = Null, 结帐id = Null, 结算性质 = Null Where 结帐id = 结帐id_In And 记录性质 = 1;
  --删除冲余款
  Delete 病人预交记录 Where 结帐id = 结帐id_In And 记录性质 = 11;

  --删除误差记录
  If n_门诊标志 = 1 Then
    Delete 门诊费用记录 Where 结帐id = 结帐id_In And 附加标志 = 9;
  Else
    Delete 住院费用记录 Where 结帐id = 结帐id_In And 附加标志 = 9;
  End If;

  --4.重新生成病人预交记录相关数据
  --4.1.补款结算,保险结算
  If 结帐结算_In Is Not Null Then
    v_结算内容 := 结帐结算_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算号码 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
      If Instr(v_结算号码, ',') > 0 Then
        --医保结算:保险类别,保险密码,保险帐号
        v_结算号码 := v_结算号码 || ',';
        v_保险类别 := Substr(v_结算号码, 1, Instr(v_结算号码, ',') - 1);
        v_结算号码 := Substr(v_结算号码, Instr(v_结算号码, ',') + 1);
        v_保险密码 := Substr(v_结算号码, 1, Instr(v_结算号码, ',') - 1);
        v_结算号码 := Substr(v_结算号码, Instr(v_结算号码, ',') + 1);
        v_保险帐号 := Substr(v_结算号码, 1, Instr(v_结算号码, ',') - 1);
        v_结算号码 := Null;
      Else
        v_保险类别 := Null;
        v_保险密码 := Null;
        v_保险帐号 := Null;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员编号, 操作员姓名, 冲预交,
           结帐id, 缴款, 找补, 缴款组id, 结算性质)
        Values
          (病人预交记录_Id.Nextval, v_No, Null, 2, 1, v_病人id, r_Pati.主页id, r_Pati.出院科室id, Null, v_结算方式, v_结算号码, '结帐缴款',
           v_保险类别, v_保险密码, v_保险帐号, v_收款时间, v_操作员编号, v_操作员姓名, n_结算金额, 结帐id_In,
           Decode(v_结算内容, 结帐结算_In || '||', 缴款_In, Null), Decode(v_结算内容, 结帐结算_In || '||', 找补_In, Null), n_组id, 2);
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 三方帐户结算_In Is Not Null Then
    --结算方式|结算金额|卡类别ID|卡号|交易流水号|交易说明||...
    v_结算内容 := 三方帐户结算_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    
      v_卡号     := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    
      v_交易流水号 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_交易说明   := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员编号, 操作员姓名, 冲预交,
           结帐id, 缴款, 找补, 缴款组id, 结算性质, 卡类别id, 卡号, 交易流水号, 交易说明)
        Values
          (病人预交记录_Id.Nextval, v_No, Null, 2, 1, v_病人id, r_Pati.主页id, r_Pati.出院科室id, Null, v_结算方式, v_结算号码, '结帐缴款', Null,
           Null, Null, v_收款时间, v_操作员编号, v_操作员姓名, n_结算金额, 结帐id_In, Null, Null, n_组id, 2, n_卡类别id, v_卡号, v_交易流水号, v_交易说明);
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --4.2.预交结算
  If 冲预交_In Is Not Null Then
    v_结算内容   := 冲预交_In || '||';
    n_门诊冲预交 := 0;
    n_住院冲预交 := 0;
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_预交id   := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1)); --是记录冲预交的ID
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1); --是记录冲预交的NO号
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_记录状态 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If v_预交id <> 0 Then
        --第一次冲预交(将第一次标上结帐ID,冲预交标记为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 2 Where ID = v_预交id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, v_记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
               v_收款时间, v_操作员姓名, v_操作员编号, n_结算金额, 结帐id_In, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2
        From 病人预交记录
        Where NO = v_结算号码 And 记录性质 In (1, 11) And 记录状态 = v_记录状态 And Rownum = 1;
    
      Begin
        Select Nvl(预交类别, 2)
        Into n_类别
        From 病人预交记录
        Where NO = v_结算号码 And 记录性质 In (1, 11) And 记录状态 = v_记录状态 And Rownum = 1;
      Exception
        When Others Then
          n_类别 := 2;
      End;
      If Nvl(n_类别, 0) = 1 Then
        n_门诊冲预交 := n_门诊冲预交 + Nvl(n_结算金额, 0);
      Else
        n_住院冲预交 := n_住院冲预交 + Nvl(n_结算金额, 0);
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  
    --更新病人余额
    If n_门诊冲预交 <> 0 Then
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_门诊冲预交
      Where 病人id = v_病人id And 性质 = 1 And 类型 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (v_病人id, 1, -1 * n_门诊冲预交, 1);
        n_返回值 := -1 * n_门诊冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额 Where 病人id = v_病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    End If;
    If n_住院冲预交 <> 0 Then
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_住院冲预交
      Where 病人id = v_病人id And 性质 = 1 And 类型 = 2
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (v_病人id, 2, -1 * n_住院冲预交, 1);
        n_返回值 := -1 * n_住院冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额 Where 病人id = v_病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    End If;
  
  End If;

  --5.相关汇总表的处理
  --汇总人员缴款余额
  --缴款结算,保险结算
  n_Dele := 0;
  For c_结帐 In (Select 结算方式, 冲预交 From 病人预交记录 Where 结帐id = 结帐id_In And 记录性质 = 2) Loop
  
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + Nvl(c_结帐.冲预交, 0)
    Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = c_结帐.结算方式
    Returning 余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (v_操作员姓名, c_结帐.结算方式, 1, Nvl(c_结帐.冲预交, 0));
      n_返回值 := Nvl(Nvl(c_结帐.冲预交, 0), 0);
    End If;
  
    If Nvl(n_返回值, 0) = 0 Then
      n_Dele := 1;
    End If;
  
  End Loop;

  If n_Dele = 1 Then
    Delete From 人员缴款余额 Where 性质 = 1 And 收款员 = v_操作员姓名 And Nvl(余额, 0) = 0;
  End If;
  --汇总表,只需重汇误差行,因为其它项不会变,未结费用不变(新产生的误差项已结帐),只有一行误差记录,仅为使用变量方便而用游标

  --6.医保相关表的处理
  --Delete 医保核对表 Where 结帐Id=结帐Id_IN;
  Update 保险结算明细 Set 标志 = 2 Where 结帐id = 结帐id_In;

  Close c_Pati;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_住院收费结算_Update;
/


--91903:蔡青松,2015-12-28,将已审核但未采样的微生物医嘱的执行状态改为已执行
CREATE OR REPLACE Procedure Zl_检验标本记录_报告审核
(
  Id_In       检验标本记录.Id%Type,
  审核人_In   检验标本记录.审核人%Type := Null,
  人员编号_In 人员表.编号%Type := Null,
  人员姓名_In 人员表.姓名%Type := Null
) Is

  --未审核的费用行(不包含药品)
  Cursor c_Verify(v_医嘱id In Number) Is
    Select Distinct 记录性质, NO, 序号
    From 住院费用记录
    Where 收费类别 Not In ('5', '6', '7') And 记帐费用 = 1 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id = v_医嘱id
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id))) And 医嘱序号 = v_医嘱id
    Union All
    Select Distinct 记录性质, NO, 序号
    From 门诊费用记录
    Where 收费类别 Not In ('5', '6', '7') And 记帐费用 = 1 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id = v_医嘱id
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id))) And 医嘱序号 = v_医嘱id
    Order By 记录性质, NO, 序号;

  --查找当前标本的相关申请
  Cursor c_Samplequest(v_微生物 In Number) Is
    Select Distinct 医嘱id, 病人来源
    From (Select a.医嘱id, b.病人来源
           From 检验项目分布 A, 检验标本记录 B
           Where 0 = v_微生物 And a.标本id = Id_In And a.医嘱id Is Not Null And a.标本id = b.Id
           Union
           Select a.医嘱id, b.病人来源
           From 检验项目分布 A, 检验标本记录 B
           Where 1 = v_微生物 And a.标本id = Id_In And a.医嘱id Is Not Null And a.标本id = b.Id
           Union
           Select b.Id As 医嘱id, a.病人来源
           From 检验标本记录 A, 病人医嘱记录 B
           Where a.Id = Id_In And a.医嘱id = b.相关id);

  Cursor c_Stuff
  (
    v_No     Varchar2,
    v_主页id Number
  ) Is
    Select NO, 单据, 库房id
    From 未发药品记录
    Where NO = v_No And 单据 In (24, 25, 26) And 库房id Is Not Null And Not Exists
     (Select 1 From Dual Where zl_GetSysParameter(Decode(v_主页id, Null, 92, 63)) = '1') And Exists
     (Select a.序号
           From 住院费用记录 A, 材料特性 B
           Where a.记录性质 = 2 And a.记录状态 = 1 And a.No = v_No And a.收费细目id = b.材料id And b.跟踪在用 = 1
           Union All
           Select a.序号
           From 门诊费用记录 A, 材料特性 B
           Where a.记录性质 = 2 And a.记录状态 = 1 And a.No = v_No And a.收费细目id = b.材料id And b.跟踪在用 = 1)
    Order By 库房id;

  v_执行 Number(1);
  v_No   病人医嘱发送.No%Type;
  v_性质 病人医嘱发送.记录性质%Type;
  v_序号 Varchar2(1000);

  v_Count Number(18);

  v_微生物标本 Number(1) := 0;
  v_主页id     Number(18);
  v_婴儿       Number(1);
  v_年龄       Varchar2(100);
  v_仪器       Number(18);
  v_Intloop    Number;
  Err_Custom Exception;
  v_Error Varchar2(100);
Begin
  Select Nvl(婴儿, 0), 年龄 Into v_婴儿, v_年龄 From 检验标本记录 Where ID = Id_In;

  --执行后自动审核对应的记帐划价单(不包含药品)
  Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_执行 From Dual;

  v_微生物标本 := 0;
  Begin
    Select 1 Into v_微生物标本 From 检验标本记录 Where 微生物标本 = 1 And ID = Id_In;
  Exception
    When Others Then
      v_微生物标本 := 0;
  End;

  --1.置本标本的状态及审核人和时间
  Update 检验标本记录
  Set 审核人 = Decode(审核人_In, Null, 人员姓名_In, 审核人_In), 审核时间 = Sysdate, 样本状态 = 2
  Where ID = Id_In;

  --记录审核过程
  Insert Into 检验操作记录
    (ID, 标本id, 操作类型, 操作员, 操作时间)
  Values
    (检验操作记录_Id.Nextval, Id_In, 0, Decode(审核人_In, Null, 人员姓名_In, 审核人_In), Sysdate);

  --2.检查当前标本相关的申请的相关标本是否完成审核
  For r_Samplequest In c_Samplequest(v_微生物标本) Loop

    v_Count := 0;

    If v_微生物标本 = 0 Then
      Begin
        Select Nvl(Count(1), 0)
        Into v_Count
        From 检验标本记录
        Where 样本状态 < 2 And ID In (Select 标本id From 检验项目分布 Where 医嘱id = r_Samplequest.医嘱id);
      Exception
        When Others Then
          v_Count := 0;
      End;
    End If;

    --r_SampleQuest.医嘱id申请已经完成,处理后续环节
    If v_Count = 0 Then

      --1.置申请单的执行状态
      Update 病人医嘱发送
      Set 执行状态 = 1, 完成人 = 人员姓名_In, 完成时间 = Sysdate
      Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id));
      
      update 病人医嘱发送
      Set 执行状态 = 1, 完成人 = 人员姓名_In, 完成时间 = Sysdate
      Where 医嘱id In (select 相关ID from 病人医嘱记录 where ID in(Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));

      If r_Samplequest.病人来源 = 2 Then
        --2.费用执行处理
        Update 住院费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
      Else
        Update 门诊费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
      End If;

      --3.自动审核记帐
      If v_执行 = 1 Then
        For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
          If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
            If v_序号 Is Not Null Then
              If v_性质 = 1 Then
                Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
              Elsif v_性质 = 2 Then
                Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
              End If;
            End If;
            v_序号 := Null;
          End If;
          v_No   := r_Verify.No;
          v_性质 := r_Verify.记录性质;
          v_序号 := v_序号 || ',' || r_Verify.序号;
        End Loop;
        If v_序号 Is Not Null Then
          If v_性质 = 1 Then
            Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          Elsif v_性质 = 2 Then
            Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          End If;
          v_序号 := Null;
        End If;
      End If;

      --审核试剂消耗单
      v_Intloop := 1;
      v_No      := Null;
      Select 仪器id Into v_仪器 From 检验标本记录 Where ID = Id_In;
      For r_检验试剂 In (Select c.材料id, c.数量
                     From 病人医嘱记录 A, 检验报告项目 B, 检验试剂关系 C
                     Where a.相关id = r_Samplequest.医嘱id And a.诊疗项目id = b.诊疗项目id And b.报告项目id = c.项目id And c.仪器id = v_仪器) Loop
        Zl_检验试剂记录_Insert(r_Samplequest.医嘱id, v_Intloop, r_检验试剂.材料id, r_检验试剂.数量);
        v_Intloop := v_Intloop + 1;
      End Loop;
      Select Count(*) Into v_Intloop From 检验试剂记录 Where 医嘱id = r_Samplequest.医嘱id And NO Is Null;
      If v_Intloop > 1 Then
        v_No := Nextno(14);
        Update 检验试剂记录 Set NO = v_No Where 医嘱id = r_Samplequest.医嘱id;
      End If;
      If v_No Is Not Null Then

        Zl_检验试剂记录_Bill(r_Samplequest.医嘱id, v_No);

        v_主页id := Null;
        Select 主页id Into v_主页id From 病人医嘱记录 A Where ID = r_Samplequest.医嘱id;

        If v_主页id Is Null Then
          Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In);
        Else
          Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In);
        End If;

        --如果记帐没有自动发料,则自动发料,否则不处理
        For r_Stuff In c_Stuff(v_No, v_主页id) Loop
          Zl_材料收发记录_处方发料(r_Stuff.库房id, 25, v_No, 人员姓名_In, 人员姓名_In, 人员姓名_In, 1, Sysdate);
        End Loop;
      End If;
    End If;
  End Loop;
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
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
End Zl_检验标本记录_报告审核;
/

--91225:梁经伙,2015-12-28,在疾病申报记录里面处理状态是1才可报送，但是在传染病管理系统里面处理状态是3才能报送
Create Or Replace Procedure Zl_疾病申报记录_Send
(
  文件id_In   In Varchar2,
  报送人_In   In 疾病申报记录.报送人%Type,
  报送时间_In In 疾病申报记录.报送时间%Type,
  报送单位_In In 疾病申报记录.报送单位%Type,
  报送备注_In In 疾病申报记录.报送备注%Type
) Is
  v_姓名   人员表.姓名%Type;
  n_文件id Number;
  e_Changed Exception;
Begin

  If Length(文件id_In) <> 32 Then
    n_文件id := To_Number(文件id_In); --新病历ID是32位GUID
  End If;

  Select 姓名 Into v_姓名 From 人员表 P, 上机人员表 U Where p.Id = u.人员id And u.用户名 = User And Rownum < 2;
  If Length(文件id_In) <> 32 Then
    --如果没有归档，则将其归档 
    Update 电子病历记录 Set 归档人 = v_姓名, 归档日期 = Sysdate Where ID = 文件id_In And 归档人 Is Null;
  End If;

  Update 疾病申报记录
  Set 处理状态 = 2, 报送人 = 报送人_In, 报送时间 = 报送时间_In, 报送单位 = 报送单位_In, 报送备注 = 报送备注_In, 登记人 = v_姓名, 登记时间 = Sysdate
  Where Decode(Length(文件id_In), 32, 文档id, 文件id) = Decode(Length(文件id_In), 32, 文件id_In, n_文件id) And (处理状态 = 1 or 处理状态 = 3);
  If Sql%RowCount = 0 Then
    Raise e_Changed;
  End If;
Exception
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]用户身份不明确！[ZLSOFT]');
  When e_Changed Then
    Raise_Application_Error(-20101, '[ZLSOFT]疾病报告已经被其他用户改变！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病申报记录_Send;
/

--91780:刘尔旋,2015-12-28,相同发票号轧帐
Create Or Replace Procedure Zl_收费员轧帐票据_Insert
(
  收缴id_In   In 人员收缴票据.收缴id%Type,
  票据信息_In Varchar2
) Is
  --------------------------------------------------------------------------------------------------------------------
  --功能:收费员轧帐明细写入
  --参数:结算信息_IN:票种,性质,序号,票据张数,开始票号,终止票号,金额,发生时间|票种,性质,序号,票据张数,开始票号,终止票号,金额,发生时间|...
  --                 票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --         性质:1-正常票据;2-退费收回票据;3-重打收回票据
  --                 发生时间:yyyy-mm-dd hh24:mi:ss
  --
  --------------------------------------------------------------------------------------------------------------------
  v_结算内容 Varchar2(4000);
  v_当前结算 Varchar2(500);

  n_票种     人员收缴票据.票种%Type;
  n_性质     人员收缴票据.性质%Type;
  n_序号     人员收缴票据.序号%Type;
  n_票据张数 人员收缴票据.票据张数%Type;
  v_开始票号 人员收缴票据.开始票号%Type;
  v_终止票号 人员收缴票据.终止票号%Type;
  n_金额     人员收缴票据.金额%Type;
  v_发生时间 Varchar2(20);
  v_批次     人员收缴票据.批次%Type;

  t_开始票号 t_Strlist := t_Strlist();
  t_终止票号 t_Strlist := t_Strlist();
  t_发生时间 t_Strlist := t_Strlist();
  t_批次     t_Strlist := t_Strlist();
  t_票种     t_Numlist := t_Numlist();
  t_性质     t_Numlist := t_Numlist();
  t_序号     t_Numlist := t_Numlist();
  t_金额     t_Numlist := t_Numlist();
  t_票据张数 t_Numlist := t_Numlist();
Begin

  v_结算内容 := 票据信息_In || '|'; --以空格分开以|结尾,没有结算号码的
  While v_结算内容 Is Not Null Loop
  
    v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
    n_票种     := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    t_票种.Extend;
    t_票种(t_票种.Count) := n_票种;
  
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
    n_性质     := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    t_性质.Extend;
    t_性质(t_性质.Count) := n_性质;
  
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
    n_序号     := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    t_序号.Extend;
    t_序号(t_序号.Count) := n_序号;
  
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
    n_票据张数 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    t_票据张数.Extend;
    t_票据张数(t_票据张数.Count) := n_票据张数;
  
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
    v_开始票号 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
    t_开始票号.Extend;
    t_开始票号(t_开始票号.Count) := v_开始票号;
  
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
    v_终止票号 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
    t_终止票号.Extend;
    t_终止票号(t_终止票号.Count) := v_终止票号;
  
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
    n_金额     := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    t_金额.Extend;
    t_金额(t_金额.Count) := n_金额;
  
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
    v_发生时间 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    t_发生时间.Extend;
    If Nvl(v_发生时间, '-') = '-' Then
      t_发生时间(t_发生时间.Count) := Null;
    Else
      t_发生时间(t_发生时间.Count) := v_发生时间;
    End If;
  
    v_批次 := LTrim(Substr(v_当前结算, Instr(v_当前结算, ',') + 1));
    t_批次.Extend;
    If Nvl(v_批次, '-') = '-' Then
      t_批次(t_批次.Count) := Null;
    Else
      t_批次(t_批次.Count) := v_批次;
    End If;
  
    v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
  End Loop;
  --批量插入数据
  Forall I In 1 .. t_票种.Count
    Insert Into 人员收缴票据
      (收缴id, 票种, 性质, 序号, 票据张数, 开始票号, 终止票号, 金额, 发生时间, 批次)
    Values
      (收缴id_In, t_票种(I), t_性质(I), t_序号(I), t_票据张数(I), t_开始票号(I), t_终止票号(I), t_金额(I),
       To_Date(t_发生时间(I), 'yyyy-mm-dd hh24:mi:ss'), t_批次(I));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_收费员轧帐票据_Insert;
/

--92454:刘尔旋,2016-01-07,三方挂号排队问题
--92007:刘尔旋,2015-12-25,三方预约填写预约操作员
Create Or Replace Procedure Zl_三方机构挂号_Insert
(
  操作方式_In     Integer,
  病人id_In       门诊费用记录.病人id%Type,
  号码_In         挂号安排.号码%Type,
  号序_In         挂号序号状态.序号%Type,
  单据号_In       门诊费用记录.No%Type,
  票据号_In       门诊费用记录.实际票号%Type,
  结算方式_In     病人预交记录.结算方式%Type, --现金的结算名称
  摘要_In         门诊费用记录.摘要%Type, --预约挂号摘要信息
  发生时间_In     门诊费用记录.发生时间%Type,
  登记时间_In     门诊费用记录.登记时间%Type,
  合作单位_In     挂号合作单位.名称%Type,
  挂号金额合计_In 门诊费用记录.实收金额%Type,
  领用id_In       票据使用明细.领用id%Type,
  收费票据_In     Number := 0, --挂号是否使用收费票据
  交易流水号_In   病人预交记录.交易流水号%Type,
  交易说明_In     病人预交记录.交易说明%Type,
  预约方式_In     预约方式.名称%Type := Null,
  预交id_In       病人预交记录.Id%Type := Null,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  加入序号状态_In Number := 0,
  是否自助设备_In Number := 0,
  结帐id_In       门诊费用记录.结帐id%Type := Null,
  锁定类型_In     Number := 0,
  保险结算_In     Varchar2 := Null,
  冲预交_In       Number := Null,
  支付卡号_In     病人预交记录.卡号%Type := Null,
  退号重用_In     Number := 1,
  费别_In         门诊费用记录.费别%Type := Null,
  机器名_In       挂号序号状态.机器名%Type := Null,
  更新年龄_In     Number := 0
) As
  --功能：三方机构进行挂号(包含预约;预约挂号不扣款;预约挂号扣款)
  --入参:操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
  --      加入序号状态_In:1-表示强制加入挂号序号状态表中;否则根据启用序号或启用时段时才加入.
  --      是否自助设备_In:1-表示是医院的自助设备进行此函数的调用,自助设备调用此函数 允许加号,否则不允许
  --      锁定类型_In :0-产生正常数据 1-表示对单据进行锁定,产生未生效的单据信息;2-对锁定的记录进行解锁-生成正常数据:未生效的单据在银行扣款完成后进行解锁
  --      保险结算_IN:格式="结算方式|结算金额||....."
  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_打印id   票据打印内容.Id%Type;
  n_返回值   病人预交记录.金额%Type;
  v_排队号码 Varchar2(20);
  v_队列名称 排队叫号队列.队列名称%Type;
  n_预交id   病人预交记录.Id%Type;
  n_挂号id   病人挂号记录.Id%Type;
  v_结算内容 Varchar2(3000);
  v_当前结算 Varchar2(150);

  v_结算方式       病人预交记录.结算方式%Type;
  n_结算金额       病人预交记录.冲预交%Type;
  n_结算合计       Number(16, 5);
  n_预交金额       病人预交记录.冲预交%Type;
  n_组id           财务缴款分组.Id%Type;
  d_排队时间       Date;
  n_锁定           Number;
  n_同科限约一个号 Number(18);
  n_病人预约科室数 Number(18);
  n_已约科室       Number(18);

  n_合作单位限制       Number(18);
  n_是否开放           Number(1);
  n_Count              Number(18);
  n_行号               Number(18);
  n_序号               病人挂号记录.号序%Type;
  n_费用id             门诊费用记录.Id%Type;
  n_价格父号           Number(18);
  n_原项目id           收费项目目录.Id%Type;
  n_原收入项目id       收费项目目录.Id%Type;
  v_诊室               病人挂号记录.诊室%Type;
  n_安排id             挂号安排.Id%Type;
  n_实收金额合计       门诊费用记录.实收金额%Type;
  n_开单部门id         门诊费用记录.开单部门id%Type;
  n_实收金额           门诊费用记录.实收金额%Type;
  n_应收金额           门诊费用记录.实收金额%Type;
  n_结帐id             病人结帐记录.Id%Type;
  v_Temp               Varchar2(500);
  n_预约时段序号       Number;
  n_预约总数           Number;
  d_时段开始时间       Date;
  n_预约数量           合作单位挂号汇总.已约数%Type;
  n_号序               病人挂号记录.号序%Type;
  d_登记时间           Date;
  v_操作员编号         人员表.编号%Type;
  v_操作员姓名         人员表.姓名%Type;
  n_预约               Integer;
  v_星期               挂号安排时段.星期%Type;
  n_启用分时段         Integer;
  n_已挂数             病人挂号汇总.已挂数%Type;
  n_已约数             病人挂号汇总.已约数%Type;
  n_其中已接收         病人挂号汇总.已约数%Type;
  n_预约生成队列       Number;
  d_Date               Date;
  n_挂号序号           Number;
  v_排队标记           排队叫号队列.排队标记%Type;
  v_排队序号           排队叫号队列.排队序号%Type;
  v_机器名             挂号序号状态.机器名%Type;
  v_序号操作员         挂号序号状态.操作员姓名%Type;
  v_序号机器名         挂号序号状态.机器名%Type;
  n_序号锁定           Number := 0;
  v_付款方式           病人挂号记录.医疗付款方式%Type;
  v_费别               门诊费用记录.费别%Type;
  n_屏蔽费别           Number(3) := 0;
  n_Tmp安排id          挂号安排.Id%Type;
  n_计划id             挂号安排计划.Id%Type;
  v_年龄               病人信息.年龄%Type;
  n_合作单位限数量模式 Number;

  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.费别, a.险类, c.编码 As 付款方式
    From 病人信息 A, 医疗付款方式 C
    Where a.病人id = n_病人id And a.医疗付款方式 = c.名称(+);

  r_Pati c_Pati%RowType;

  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit(n_病人id 病人信息.病人id%Type) Is
    Select *
    From (Select a.Id, a.记录状态, a.预交类别, a.No, Nvl(a.金额, 0) As 金额
           From 病人预交记录 A,
                (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                  From 病人预交记录 A
                  Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.病人id = n_病人id And Nvl(a.预交类别, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.金额, 0)) <> 0) B
           Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                 a.No = b.No And a.病人id = n_病人id And Nvl(a.预交类别, 2) = 1
           Union All
           Select 0 As ID, 记录状态, 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
           From 病人预交记录
           Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And 病人id = n_病人id And
                 Nvl(预交类别, 2) = 1 Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
           Group By 记录状态, NO, 预交类别)
    Order By ID, NO;

  Cursor c_安排
  (
    v_号码        挂号安排.号码%Type,
    d_发生时间_In Date
  ) Is
    Select *
    From (With 安排时间段 As (Select 时间段
                         From (Select 时间段,
                                       To_Date(Decode(Sign(开始时间 - 终止时间), 1, '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                       To_Date(Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As 终止时间,
                                       To_Date('3000-01-10 ' || To_Char(d_发生时间_In, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As 当前时间,
                                       To_Date('3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间1,
                                       To_Date('3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As 终止时间1
                                From 时间段)
                         Where 当前时间 Between 开始时间 And 终止时间1 Or 当前时间 Between 开始时间1 And 终止时间)
           Select Distinct p.Id, p.号类, p.号码, p.科室id, b.编码 As 科室编码, b.名称 As 科室名称, p.项目id, c.编码 As 项目编码, c.名称 As 项目名称,
                           p.医生id, d.编号 As 医生编号, p.医生姓名, p.限号数, p.限约数, p.周日 As 日, p.周一 As 一, p.周二 As 二, p.周三 As 三,
                           p.周四 As 四, p.周五 As 五, p.周六 As 六, p.序号控制
           From (Select p.Id, p.号码, p.号类, p.科室id, p.项目id, p.医生id, p.医生姓名, b.限号数, b.限约数, Nvl(p.病案必须, 0) As 病案必须, p.周日, p.周一,
                         p.周二, p.周三, p.周四, p.周五, p.周六, p.分诊方式, p.序号控制,
                         Decode(To_Char(d_发生时间_In, 'D'), '1', p.周日, '2', p.周一, '3', p.周二, '4', p.周三, '5', p.周四, '6', p.周五,
                                 '7', p.周六, Null) As 排班
                  From 挂号安排 P, 挂号安排限制 B
                  Where p.停用日期 Is Null And p.Id = b.安排id(+) And
                        b.限制项目(+) = Decode(To_Char(d_发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                           '6', '周五', '7', '周六', Null) And
                        d_发生时间_In Between Nvl(p.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                        Nvl(p.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                   (Select 1
                         From 挂号安排计划
                         Where 安排id = p.Id And (d_发生时间_In Between 生效时间 And 失效时间) And 审核时间 Is Not Null) And Not Exists
                   (Select 1
                         From 挂号安排停用状态
                         Where 安排id = p.Id And d_发生时间_In Between 开始停止时间 And 结束停止时间) And p.号码 = v_号码
                  Union All
                  Select c.Id, c.号码, c.号类, c.科室id, p.项目id, p.医生id, p.医生姓名, b.限号数, b.限约数, Nvl(c.病案必须, 0) As 病案必须, p.周日, p.周一,
                         p.周二, p.周三, p.周四, p.周五, p.周六, p.分诊方式, p.序号控制,
                         Decode(To_Char(d_发生时间_In, 'D'), '1', p.周日, '2', p.周一, '3', p.周二, '4', p.周三, '5', p.周四, '6', p.周五,
                                 '7', p.周六, Null) As 排班
                  From 挂号安排计划 P, 挂号安排 C, 挂号计划限制 B,
                       (Select Max(a.生效时间) As 生效, 安排id
                         From 挂号安排计划 A, 挂号安排 B
                         Where a.安排id = b.Id And a.审核时间 Is Not Null And
                               发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                               Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.号码 = 号码_In
                         Group By 安排id) E
                  Where p.安排id = c.Id And p.Id = b.计划id(+) And p.生效时间 = e.生效 And p.安排id = e.安排id And
                        Nvl(p.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) = Nvl(e.生效, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                        b.限制项目(+) = Decode(To_Char(d_发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                           '6', '周五', '7', '周六', Null) And (d_发生时间_In Between p.生效时间 And p.失效时间) And
                        p.审核时间 Is Not Null And Not Exists
                   (Select 1
                         From 挂号安排停用状态
                         Where 安排id = c.Id And d_发生时间_In Between 开始停止时间 And 结束停止时间) And p.号码 = v_号码) P, 部门表 B, 收费项目目录 C,
                人员表 D
           Where p.科室id = b.Id And p.医生id = d.Id(+) And p.项目id = c.Id And
                 (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                 (Nvl(p.医生id, 0) = 0 Or Exists
                  (Select 1
                   From 人员表 Q
                   Where p.医生id = q.Id And (q.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or q.撤档时间 Is Null))) And Exists
            (Select 1 From 安排时间段 Where 时间段 = p.排班))
           Order By 号码;


  r_安排 c_安排%RowType;

  Function Zl_诊室(号码_In 挂号安排.号码%Type) Return Varchar2 As
    n_分诊方式 挂号安排.分诊方式%Type;
    n_安排id   挂号安排.Id%Type;
    v_诊室     病人挂号记录.诊室%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
  
    If 锁定类型_In = 2 Then
      --对单据进行解锁,首先检查是否存在锁定
      Select Count(Rowid) Into n_锁定 From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      If n_锁定 = 0 Then
        v_Err_Msg := '单据号为(' || 单据号_In || ')的单据,不存在或者已经被解锁!';
        Raise Err_Item;
      End If;
      Select Max(号序) Into n_号序 From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
    End If;
  
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

  Function Zl_操作员
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
    -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
  Begin
    If Type_In = 0 Then
      --缺省部门
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --操作员编码
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --操作员姓名
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
  If 费别_In Is Null Then
    Select Zl_Custom_Getpatifeetype(1, 病人id_In) Into v_费别 From Dual;
  Else
    v_费别 := 费别_In;
  End If;
  If v_费别 Is Null Then
    n_屏蔽费别 := 1;
    Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
  End If;
  Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
  If 更新年龄_In = 1 Then
    Select Zl_Age_Calc(病人id_In) Into v_年龄 From Dual;
    If v_年龄 Is Not Null Then
      Update 病人信息 Set 年龄 = v_年龄 Where 病人id = 病人id_In;
    End If;
  End If;
  --获取当前机器名称
  If 机器名_In Is Not Null Then
    v_机器名 := 机器名_In;
  Else
    Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  End If;
  n_实收金额合计 := 0;
  Select Count(*) + 1
  Into n_挂号序号
  From 病人挂号记录
  Where 号别 = 号码_In And 登记时间 Between Trunc(发生时间_In) And Trunc(发生时间_In + 1) - 1 / 24 / 60 / 60;
  --Begin
  --部门ID,部门名称;人员ID,人员编号,人员姓名
  v_Temp := Zl_Identity(0);
  If Nvl(v_Temp, ' ') = ' ' Then
    v_Err_Msg := '当前操作人员未设置对应的人员关系,不能继续。';
    Raise Err_Item;
  End If;

  If 登记时间_In Is Null Then
    d_登记时间 := Sysdate;
  Else
    d_登记时间 := 登记时间_In;
  End If;
  If Trunc(Sysdate) > Trunc(发生时间_In) Then
    v_Err_Msg := '不能挂以前的号(' || To_Char(发生时间_In, 'yyyy-mm-dd') || ')。';
    Raise Err_Item;
  End If;
  n_同科限约一个号 := Nvl(zl_GetSysParameter('病人同科限约一个号', 1111), 0);
  n_病人预约科室数 := Nvl(zl_GetSysParameter('病人预约科室数', 1111), 0);
  n_开单部门id     := To_Number(Zl_操作员(0, v_Temp));
  v_操作员编号     := Zl_操作员(1, v_Temp);
  v_操作员姓名     := Zl_操作员(2, v_Temp);
  n_组id           := Zl_Get组id(v_操作员姓名);

  If 操作方式_In <> 1 Then
    --预约检查是否添加合作单位控制
    --如果设置了合作单位控制 则
    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 号码 = 号码_In And 发生时间_In Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
            Nvl(失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Rownum < 2
      Order By 生效时间 Desc;
    Exception
      When Others Then
        Select ID Into n_Tmp安排id From 挂号安排 Where 号码 = 号码_In;
    End;
    If Nvl(n_计划id, 0) <> 0 Then
      Select Count(0)
      Into n_合作单位限制
      From 合作单位计划控制
      Where 合作单位 = 合作单位_In And 计划id = n_计划id And Rownum < 2;
    Else
      Select Count(0)
      Into n_合作单位限制
      From 合作单位安排控制
      Where 合作单位 = 合作单位_In And 安排id = n_Tmp安排id And Rownum < 2;
    End If;
  End If;

  If 操作方式_In <> 2 Then
    v_诊室 := Zl_诊室(号码_In);
  End If;
  If 操作方式_In <> 2 And 结算方式_In Is Not Null Then
    --检查结算方式是否完备
    Select Count(*) Into n_Count From 结算方式 Where 名称 = Nvl(结算方式_In, 'Lxh') And 性质 In (2, 7, 8);
    If Nvl(卡类别id_In, 0) <> 0 And n_Count = 0 Then
      Select Count(1)
      Into n_Count
      From 医疗卡类别
      Where ID = Nvl(卡类别id_In, 0) And 结算方式 = Nvl(结算方式_In, 'lxh');
    End If;
    If n_Count = 0 Then
      v_Err_Msg := '结算方式(' || 结算方式_In || ')未设置,请在结算方式管理中设置。';
      Raise Err_Item;
    End If;
  End If;

  --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
  Select Count(*) Into n_Count From 门诊费用记录 Where 记录性质 = 4 And 记录状态 In (1, 3) And NO = 单据号_In;
  If n_Count <> 0 Then
    v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
    Raise Err_Item;
  End If;

  Open c_Pati(病人id_In);
  n_Count := 0;
  Begin
    Fetch c_Pati
      Into r_Pati;
  Exception
    When Others Then
      n_Count := -1;
  End;
  If n_Count = -1 Then
    v_Err_Msg := '病人未找到，不能继续。';
    Raise Err_Item;
  End If;

  Open c_安排(号码_In, 发生时间_In);
  Begin
    Fetch c_安排
      Into r_安排;
  Exception
    When Others Then
      n_Count := -1;
  End;
  If n_Count = -1 Then
    v_Err_Msg := '该号别没有在' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '中进行安排。';
    Raise Err_Item;
  End If;

  Select Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', '周日')
  Into v_星期
  From Dual;
  Begin
    Select 1 Into n_启用分时段 From 挂号安排时段 Where 安排id = r_安排.Id And 星期 = v_星期 And Rownum <= 1;
  Exception
    When Others Then
      n_启用分时段 := 0;
  End;

  --对参数控制进行检查
  --仅在预约不扣款时进行检查
  If 操作方式_In = 2 Then
    If Nvl(n_同科限约一个号, 0) <> 0 Or Nvl(n_病人预约科室数, 0) <> 0 Then
      n_已约科室 := 0;
      For c_Chkitem In (Select Count(1) As 已约, a.执行部门id As 科室id, Nvl(k.名称, '') As 科室
                        From 病人挂号记录 A, 病人信息 B, 部门表 K
                        Where a.病人id = b.病人id And a.病人id = 病人id_In And a.执行部门id = k.Id(+) And a.记录性质 = 2 And 记录状态 = 1 And
                              a.预约时间 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60
                        Group By a.执行部门id, k.名称) Loop
        If Nvl(n_同科限约一个号, 0) <> 0 And c_Chkitem.科室id = r_安排.科室id Then
        
          v_Err_Msg := '该病人已经在科室[' || c_Chkitem.科室 || ']进行了预约,不能再预约！';
          Raise Err_Item;
        
          If Nvl(n_病人预约科室数, 0) > 0 And c_Chkitem.科室id <> r_安排.科室id Then
            n_已约科室 := n_已约科室 + 1;
          End If;
        End If;
      End Loop;
      If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
        v_Err_Msg := '同一病人在最多同时预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
        Raise Err_Item;
      End If;
    End If;
  End If;

  d_Date         := Null;
  d_时段开始时间 := Null;

  If Nvl(r_安排.限号数, 0) >= 0 Or r_安排.限号数 Is Null Then
  
    Select Nvl(Sum(Nvl(b.已挂数, 0)), 0), Nvl(Sum(Nvl(b.其中已接收, 0)), 0), Nvl(Sum(Nvl(b.已约数, 0)), 0)
    Into n_已挂数, n_其中已接收, n_已约数
    From 挂号安排 A, 病人挂号汇总 B
    Where a.科室id = b.科室id And a.项目id = b.项目id And a.号码 = 号码_In And b.日期 Between Trunc(发生时间_In) And
          Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And (a.号码 = b.号码 Or b.号码 Is Null) And Nvl(a.医生id, 0) = Nvl(b.医生id, 0) And
          Nvl(a.医生姓名, '医生') = Nvl(b.医生姓名, '医生');
  
    If n_启用分时段 = 1 Then
      If Nvl(r_安排.序号控制, 0) = 1 Then
        If Nvl(是否自助设备_In, 0) = 0 Then
          Select Count(*), Max(开始时间)
          Into n_Count, d_时段开始时间
          From 挂号安排时段
          Where 安排id = r_安排.Id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0);
          v_Temp := '挂号';
          If 操作方式_In > 1 Then
            v_Temp := '预约挂号';
          End If;
        
          If n_Count = 0 Then
            v_Err_Msg := '号别为' || 号码_In || '的挂号安排中不存在序号为' || Nvl(号序_In, 0) || '的安排,不能再' || v_Temp || '！';
            Raise Err_Item;
          End If;
        End If;
        --过点的,不能选择挂号
        If Trunc(Sysdate) = Trunc(发生时间_In) Then
          --挂当天的号
          v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
          For v_时段 In (Select To_Date(v_Temp || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                              To_Date(To_Char(Sysdate + Decode(Sign(开始时间 - 结束时间), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                       To_Char(结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间, 限制数量, 是否预约
                       From 挂号安排时段
                       Where 安排id = r_安排.Id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0)) Loop
            If Sysdate > v_时段.结束时间 Then
              v_Err_Msg := '号别为' || 号码_In || '及号序为' || Nvl(号序_In, 0) || '的安排,已经超过时点,不能再' || v_Temp || '！';
              Raise Err_Item;
            End If;
          End Loop;
        End If;
      Elsif 操作方式_In > 1 Then
        --未启用序号的,需要检查预约的情况
      
        n_Count := 0;
        For v_时段 In (Select 序号, 开始时间, 结束时间, 限制数量, 是否预约
                     From 挂号安排时段
                     Where 安排id = r_安排.Id And 星期 = v_星期 And
                           (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                           Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1, '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                    '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')) And
                           '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                           ('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                           '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS') And
                           Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                    '3000-01-11 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                    '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
          n_预约时段序号 := v_时段.序号;
          d_时段开始时间 := v_时段.开始时间;
        
          Select Count(*), Max(序号)
          Into n_Count, n_预约总数
          From 挂号序号状态
          Where 号码 = 号码_In And 日期 = 发生时间_In And 状态 Not In (4, 5);
        
          If Nvl(n_Count, 0) > Nvl(v_时段.限制数量, 0) And 锁定类型_In <> 2 Then
            v_Err_Msg := '号别为' || 号码_In || '的挂号安排中在' || To_Char(v_时段.开始时间, 'hh24:mi:ss') || '至' ||
                         To_Char(v_时段.结束时间, 'hh24:mi:ss') || '最多只能预约' || Nvl(v_时段.限制数量, 0) || '人,不能再进行预约挂号！';
            Raise Err_Item;
          End If;
          n_Count := 1;
        End Loop;
      
        If n_Count = 0 Then
          v_Err_Msg := '号别为' || 号码_In || '的挂号安排中没有相关的安排计划(' || To_Char(发生时间_In, 'YYYY-mm-dd HH24:MI:SS') ||
                       '),不能进行预约挂号！';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  If 操作方式_In = 1 And 锁定类型_In <> 2 Then
    --挂号规则:
    --  已挂数不能大于限号数
    If n_已挂数 >= Nvl(r_安排.限号数, 0) And r_安排.限号数 Is Not Null Then
      v_Err_Msg := '该号别今天已达到限号数 ' || Nvl(r_安排.限号数, 0) || '不能再挂号！';
      Raise Err_Item;
    End If;
  End If;

  If 操作方式_In > 1 Then
    --预约的相关检查
    --规则:
    --   1.已限约不能超过限约数
    --   2.检查是否启用时段的
    If n_已约数 >= Nvl(r_安排.限约数, 0) And Nvl(r_安排.限约数, 0) <> 0 And r_安排.限约数 Is Not Null And 锁定类型_In <> 2 Then
      v_Err_Msg := '该号别已达到限约数 ' || Nvl(r_安排.限约数, 0) || '不能再预约挂号！';
      Raise Err_Item;
    End If;
  End If;
  If n_合作单位限制 > 0 And 操作方式_In <> 1 And 合作单位_In Is Not Null Then
  
    If Nvl(r_安排.序号控制, 0) = 1 And Nvl(号序_In, 0) = 0 Then
      v_Err_Msg := '当前安排使用了序号控制,请确认所需要预约的序号,不能继续。';
      Raise Err_Item;
    End If; --Nvl(r_安排.序号控制, 0) =0
  
    n_序号 := Case
              When Nvl(r_安排.序号控制, 0) = 1 Or n_启用分时段 = 1 And 操作方式_In > 1 Then
               Nvl(号序_In, 0)
              Else
               0
            End;
    --合作单位限数量模式
    Begin
      If Nvl(n_计划id, 0) <> 0 Then
        Select 0
        Into n_序号
        From 合作单位计划控制
        Where 合作单位 = 合作单位_In And 计划id = n_计划id And
              限制项目 = Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                            '周六', Null) And 数量 <> 0 And 序号 = 0 And Rownum < 2;
      Else
        Select 0
        Into n_序号
        From 合作单位安排控制
        Where 合作单位 = 合作单位_In And 安排id = n_Tmp安排id And
              限制项目 = Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                            '周六', Null) And 数量 <> 0 And 序号 = 0 And Rownum < 2;
      End If;
      n_合作单位限数量模式 := 1;
    Exception
      When Others Then
        n_合作单位限数量模式 := 0;
    End;
    --开放序号检查
    For c_合作单位 In (Select c.序号, 数量
                   From 挂号安排 A, 合作单位安排控制 C
                   Where a.号码 = 号码_In And Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                 '周四', '6', '周五', '7', '周六', Null) = c.限制项目(+) And a.Id = c.安排id And
                         c.合作单位 = 合作单位_In And c.序号 = n_序号 And Not Exists
                    (Select 1
                          From 挂号安排计划 D
                          Where d.安排id = a.Id And d.审核时间 Is Not Null And
                                发生时间_In Between Nvl(d.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                Nvl(d.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')))
                   Union All
                   Select c.序号, 数量
                   From 挂号安排计划 A, 挂号安排 D, 合作单位计划控制 C,
                        (Select Max(a.生效时间) As 生效, 安排id
                          From 挂号安排计划 A, 挂号安排 B
                          Where a.安排id = b.Id And a.审核时间 Is Not Null And
                                发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.号码 = 号码_In
                          Group By 安排id) E
                   Where a.安排id = d.Id And a.审核时间 Is Not Null And d.号码 = 号码_In And a.安排id = e.安排id And
                         Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                         Nvl(e.生效, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                         Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                '7', '周六', Null) = c.限制项目(+) And a.Id = c.计划id And c.合作单位 = 合作单位_In And c.序号 = n_序号 And
                         发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                         Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'))) Loop
    
      If Nvl(r_安排.序号控制, 0) = 1 And c_合作单位.序号 = n_序号 And n_合作单位限数量模式 = 0 Then
        n_是否开放 := 1;
        Exit;
      Elsif (Nvl(r_安排.序号控制, 0) = 0 And c_合作单位.序号 = n_序号) Or n_合作单位限数量模式 = 1 Then
        Begin
          Select Nvl(已约数, 0)
          Into n_预约数量
          From 合作单位挂号汇总
          Where 合作单位 = 合作单位_In And 日期 = Trunc(发生时间_In) And 号码 = 号码_In;
        Exception
          When Others Then
            n_预约数量 := 0;
        End;
        If c_合作单位.数量 <= n_预约数量 And Nvl(c_合作单位.数量, 0) > 0 And 锁定类型_In <> 2 Then
          v_Err_Msg := '该号别已达到限约数 ' || Nvl(c_合作单位.数量, 0) || '不能再预约挂号！';
          Raise Err_Item;
        End If;
        n_是否开放 := 1;
        Exit;
      End If;
    
    End Loop;
  
    If Nvl(n_是否开放, 0) = 0 Then
      v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '未开放,不能继续。';
      Raise Err_Item;
    End If;
  End If;

  --检查限号数和限约数
  n_行号         := 1;
  n_原项目id     := 0;
  n_原收入项目id := 0;
  n_实收金额合计 := 0;
  If 锁定类型_In <> 1 Then
    If 操作方式_In <> 2 Then
      If Nvl(结帐id_In, 0) = 0 Then
        --这里应该程序传入
        Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
      Else
        n_结帐id := 结帐id_In;
      End If;
    Else
      n_结帐id := Null;
    End If;
  End If;
  For c_Item In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次, c.Id As 收入项目id,
                        c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                 From 收费项目目录 A, 收费价目 B, 收入项目 C
                 Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = r_安排.项目id And Sysdate Between b.执行日期 And
                       Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Union All
                 Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                        c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, 1 As 从属父号
                 From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
                 Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And d.主项id = r_安排.项目id And
                       Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Order By 性质, 项目编码, 收入编码) Loop
    n_价格父号 := Null;
    If n_原项目id = c_Item.项目id Then
      If n_原收入项目id <> c_Item.收入项目id Then
        n_价格父号 := n_行号;
      End If;
      n_原收入项目id := c_Item.收入项目id;
    End If;
    n_原项目id := c_Item.项目id;
    n_应收金额 := Round(c_Item.数次 * c_Item.单价, 5);
    n_实收金额 := n_应收金额;
    If Nvl(c_Item.屏蔽费别, 0) <> 1 And n_屏蔽费别 = 0 Then
      --打折:
      v_Temp     := Zl_Actualmoney(r_Pati.费别, c_Item.项目id, c_Item.收入项目id, n_应收金额);
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
      n_实收金额 := Zl_To_Number(v_Temp);
    End If;
    n_实收金额合计 := Nvl(n_实收金额合计, 0) + n_实收金额;
  
    --锁定单据不产生费用
    If 锁定类型_In <> 1 Then
      --产生病人挂号费用(可能单独是或包括病历费用)
      Select 病人费用记录_Id.Nextval Into n_费用id From Dual;
      --:操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
      Insert Into 门诊费用记录
        (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id,
         收费类别, 计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人,
         操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
      Values
        (n_费用id, 4, Decode(操作方式_In, 2, 0, 1), n_行号, n_价格父号, c_Item.从属父号, 单据号_In, 票据号_In, 1, Null, Null,
         Decode(操作方式_In, 2, To_Char(号序_In), v_诊室), r_Pati.病人id, r_Pati.门诊号, r_Pati.付款方式, r_Pati.姓名, r_Pati.性别, r_Pati.年龄,
         r_Pati.费别, r_安排.科室id, c_Item.类别, 号码_In, c_Item.项目id, c_Item.收入项目id, c_Item.收据费目, 1, c_Item.数次, c_Item.单价,
         n_应收金额, n_实收金额, Decode(操作方式_In, 2, Null, n_实收金额), n_结帐id, 0, n_开单部门id, v_操作员姓名,
         Decode(操作方式_In, 2, v_操作员姓名, Null), r_安排.科室id, r_安排.医生姓名, v_操作员编号, v_操作员姓名, 发生时间_In, d_登记时间, Null, 0, Null, Null,
         摘要_In, 预约方式_In, Decode(操作方式_In, 2, Null, n_组id));
    End If;
    n_行号 := n_行号 + 1;
  
  End Loop;

  If Round(Nvl(挂号金额合计_In, 0), 5) <> Round(Nvl(n_实收金额合计, 0), 5) Then
    v_Err_Msg := '本次挂号金额不正确,可能是因为医院调整了价格,请重新获取挂号收费项目的价格,不能继续。';
    Raise Err_Item;
  End If;

  If n_启用分时段 = 1 Then
    d_Date := To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(d_时段开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_Date := Trunc(发生时间_In);
  End If;

  --更新挂号序号状态
  If 锁定类型_In <> 2 Then
    n_号序 := 号序_In;
  End If;
  Begin
    Select 1
    Into n_Count
    From 挂号序号状态
    Where Trunc(日期) = Trunc(发生时间_In) And 号码 = 号码_In And 序号 = n_号序 And 状态 <> 5;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 1 Then
    If n_启用分时段 = 0 And Nvl(r_安排.序号控制, 0) = 1 Then
      n_号序 := Null;
    End If;
    If n_启用分时段 = 1 And Nvl(r_安排.序号控制, 0) = 1 Then
      v_Err_Msg := '当前序号已被使用，请重新选择一个序号！';
      Raise Err_Item;
    End If;
  End If;
  n_Count := 0;
  If n_启用分时段 = 0 And Nvl(r_安排.序号控制, 0) = 1 And n_号序 Is Null And 锁定类型_In <> 2 Then
    If 退号重用_In = 1 Then
      Select Nvl(Max(序号), 0) + 1
      Into n_号序
      From 挂号序号状态
      Where 日期 = Trunc(发生时间_In) And 号码 = r_安排.号码 And 状态 Not In (4, 5);
    Else
      Select Nvl(Max(序号), 0) + 1
      Into n_号序
      From 挂号序号状态
      Where 日期 = Trunc(发生时间_In) And 号码 = r_安排.号码 And 状态 <> 5;
    End If;
  End If;
  If n_启用分时段 = 1 And 锁定类型_In <> 2 Then
  
    If 操作方式_In > 1 And Nvl(r_安排.序号控制, 0) = 0 Then
      --规则:预约时段序号||预约数
      If Nvl(n_预约总数, 0) = 0 Then
        v_Temp := Nvl(r_安排.限约数, 0);
        v_Temp := LTrim(RTrim(v_Temp));
        v_Temp := LPad(Nvl(n_预约总数, 0) + 1, Length(v_Temp), '0');
        v_Temp := n_预约时段序号 || v_Temp;
        n_号序 := To_Number(v_Temp);
      Else
        n_号序 := n_预约总数 + 1;
      End If;
    End If;
  End If;

  If Nvl(r_安排.序号控制, 0) = 1 Or (操作方式_In > 1 And n_启用分时段 = 1) Or 加入序号状态_In = 1 Then
    --锁定序号的处理
    Begin
      Select 操作员姓名, 机器名
      Into v_序号操作员, v_序号机器名
      From 挂号序号状态
      Where 状态 = 5 And 号码 = 号码_In And Trunc(日期) = Trunc(d_Date) And 序号 = n_号序;
      n_序号锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_序号锁定   := 0;
    End;
    If n_序号锁定 = 0 Then
      Update 挂号序号状态
      Set 状态 = Decode(操作方式_In, 2, 2, 1), 预约 = Decode(操作方式_In, 1, 0, 1), 登记时间 = Sysdate
      Where 号码 = 号码_In And 日期 = d_Date And 序号 = n_号序 And 操作员姓名 = v_操作员姓名;
      If Sql%RowCount = 0 Then
        Begin
          Insert Into 挂号序号状态
            (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
          Values
            (号码_In, d_Date, n_号序, Decode(操作方式_In, 2, 2, 1), v_操作员姓名, Decode(操作方式_In, 1, 0, 1), Sysdate);
        
          If n_合作单位限制 > 0 And 操作方式_In > 1 And Nvl(n_是否开放, 0) = 1 Then
            Update 合作单位挂号汇总
            Set 已约数 = 已约数 + Decode(操作方式_In, 2, 1, 0), 已接数 = 已接数 + Decode(操作方式_In, 3, 1, 0)
            Where 号码 = 号码_In And 日期 = d_Date And 序号 = n_号序 And 合作单位 = 合作单位_In;
            If Sql%NotFound Then
              Insert Into 合作单位挂号汇总
                (号码, 日期, 序号, 合作单位, 已约数, 已接数)
              Values
                (号码_In, d_Date, n_号序, 合作单位_In, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 3, 1, 0));
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_号序 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If v_操作员姓名 <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_号序 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        Update 挂号序号状态
        Set 状态 = Decode(操作方式_In, 2, 2, 1), 预约 = Decode(操作方式_In, 1, 0, 1), 登记时间 = Sysdate
        Where 号码 = 号码_In And Trunc(日期) = Trunc(d_Date) And 序号 = n_号序 And 状态 = 5 And 操作员姓名 = v_操作员姓名 And 机器名 = v_机器名;
      End If;
    End If;
  End If;

  --锁定单据不产生任何 费用
  If 操作方式_In <> 2 And 锁定类型_In <> 1 Then
    --挂号,预约挂号已经扣款部分
    n_预交id := 预交id_In;
    If Nvl(n_预交id, 0) = 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    End If;
    n_结算合计 := 0;
    If 保险结算_In Is Not Null Then
      --各个保险结算
      v_结算内容 := 保险结算_In || '||';
      n_结算合计 := 0;
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
        n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
        If Nvl(n_结算金额, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
          Values
            (n_预交id, 4, 单据号_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, d_登记时间, v_操作员编号, v_操作员姓名, n_结算金额,
             n_结帐id, n_组id, n_结帐id, 4);
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        End If;
        n_结算合计 := Nvl(n_结算合计, 0) + Nvl(n_结算金额, 0);
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
      End Loop;
    End If;
  
    If Nvl(冲预交_In, 0) <> 0 Then
      --处理总预交
      n_结算合计 := n_结算合计 + Nvl(冲预交_In, 0);
      n_预交金额 := 冲预交_In;
      For r_Deposit In c_Deposit(病人id_In) Loop
        n_结算金额 := Case
                    When r_Deposit.金额 - n_预交金额 < 0 Then
                     r_Deposit.金额
                    Else
                     n_预交金额
                  End;
        If r_Deposit.Id <> 0 Then
          --第一次冲预交(82592,将第一次标上结帐ID,冲预交标记为0)
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = n_结帐id, 结算性质 = 4 Where ID = r_Deposit.Id;
        End If;
        --冲上次剩余额
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 d_登记时间, v_操作员姓名, v_操作员编号, n_结算金额, n_结帐id, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_结算金额
        Where 病人id = 病人id_In And 性质 = 1 And 类型 = Nvl(r_Deposit.预交类别, 2);
        --检查是否已经处理完
        If r_Deposit.金额 <= n_结算金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余额不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      End If;
    End If;
    --剩余款项,用指定结算方支付
    n_结算金额 := Nvl(n_实收金额合计, 0) - Nvl(n_结算合计, 0);
    If Nvl(n_结算金额, 0) < 0 Then
      v_Err_Msg := '挂号的相关结算金额超出了当前实结金额,不能继续操作！';
      Raise Err_Item;
    End If;
    If Nvl(n_结算金额, 0) <> 0 Or (Nvl(n_结算金额, 0) = 0 And Nvl(冲预交_In, 0) = 0) Then
      If 结算方式_In Is Null Then
        v_Err_Msg := '未传入指定的结算方式,不能继续操作！';
        Raise Err_Item;
      End If;
    
      If Nvl(预交id_In, 0) <> 0 Then
        --传入的预交ID_In主要是为了解决三方交易,如果医保结算站用了该ID,需要用新的ID进行更新,三方交易用转入的ID
        Update 病人预交记录 Set ID = n_预交id Where ID = Nvl(预交id_In, 0);
        n_预交id := Nvl(预交id_In, 0);
      End If;
    
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 交易流水号, 交易说明, 结算序号, 合作单位, 卡类别id, 卡号,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, r_Pati.病人id, 结算方式_In, Nvl(n_结算金额, 0), d_登记时间, v_操作员编号, v_操作员姓名, n_结帐id, 合作单位_In || '缴款',
         n_组id, 交易流水号_In, 交易说明_In, n_结帐id, 合作单位_In, 卡类别id_In, 支付卡号_In, 4);
    End If;
  
    --更新人员缴款数据
  
    For v_缴款 In (Select 结算方式, Sum(Nvl(a.冲预交, 0)) As 冲预交
                 From 病人预交记录 A
                 Where a.结帐id = n_结帐id And Mod(a.记录性质, 10) <> 1 And Nvl(病人id, 0) = Nvl(病人id_In, 0)
                 Group By 结算方式) Loop
    
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + Nvl(v_缴款.冲预交, 0)
      Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = v_缴款.结算方式
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (v_操作员姓名, v_缴款.结算方式, 1, Nvl(v_缴款.冲预交, 0));
        n_返回值 := Nvl(v_缴款.冲预交, 0);
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = v_操作员姓名 And 结算方式 = 结算方式_In And 性质 = 1 And Nvl(余额, 0) = 0;
      End If;
    
    End Loop;
  
  End If;

  --处理挂号记录
  If 锁定类型_In = 2 Then
    Begin
      Select ID Into n_挂号id From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
    Exception
      When Others Then
        Null;
    End;
  Else
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
  End If;

  Update 病人挂号记录
  Set 记录性质 = Decode(操作方式_In, 2, 2, 1), 记录状态 = Decode(锁定类型_In, 1, 0, 1), 门诊号 = r_Pati.门诊号, 操作员姓名 = v_操作员姓名,
      操作员编号 = v_操作员编号, 预约 = Decode(操作方式_In, 1, 0, 1), 接收人 = Decode(锁定类型_In, 1, Null, Decode(操作方式_In, 2, Null, v_操作员姓名)),
      接收时间 = Case 锁定类型_In
                When 1 Then
                 Null
                Else
                 Case 操作方式_In
                   When 2 Then
                    Null
                   Else
                    d_登记时间
                 End
              End, 交易流水号 = Nvl(交易流水号_In, 交易流水号), 交易说明 = Nvl(交易说明_In, 交易说明), 合作单位 = Nvl(合作单位_In, 合作单位),
      预约操作员 = Decode(操作方式_In, 1, Nvl(预约操作员, Null), Nvl(预约操作员, v_操作员姓名)),
      预约操作员编号 = Decode(操作方式_In, 1, Nvl(预约操作员编号, Null), Nvl(预约操作员编号, v_操作员编号))
  Where ID = n_挂号id;
  If Sql%NotFound Then
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = r_Pati.付款方式 And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式, 预约操作员, 预约操作员编号)
    Values
      (n_挂号id, 单据号_In, Decode(操作方式_In, 2, 2, 1), Decode(锁定类型_In, 1, 0, 1), r_Pati.病人id, r_Pati.门诊号, r_Pati.姓名,
       r_Pati.性别, r_Pati.年龄, 号码_In, 0, v_诊室, Null, r_安排.科室id, r_安排.医生姓名, 0, Null, d_登记时间, 发生时间_In,
       Case When(Nvl(操作方式_In, 0)) = 1 Then Null Else 发生时间_In End, v_操作员编号, v_操作员姓名, 0, n_号序, Decode(操作方式_In, 1, 0, 1),
       Decode(操作方式_In, 2, Null, v_操作员姓名), Decode(操作方式_In, 2, To_Date(Null), d_登记时间), 交易流水号_In, 交易说明_In, 合作单位_In, v_付款方式,
       Decode(操作方式_In, 1, Null, v_操作员姓名), Decode(操作方式_In, 1, Null, v_操作员编号));
  End If;
  --锁定单据不能产生队列
  If 锁定类型_In <> 1 Then
    n_预约生成队列 := 0;
    If 操作方式_In > 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
    --挂号和收费的预约都直接进入队列(收费预约缺少接收过程,所以直接和挂号一样直接进入队列)
    If 操作方式_In <> 2 Or n_预约生成队列 = 1 Then
      If Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113)) <> 0 Then
        --排队叫号模式:-0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站      
        If Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113)) = 0 Or n_预约生成队列 = 1 Then
          --产生队列
          --.按”执行部门” 的方式生成队列
          v_队列名称 := r_安排.科室id;
          v_排队号码 := Zlgetnextqueue(r_安排.科室id, n_挂号id, 号码_In || '|' || 号序_In);
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
          v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)  
          d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号码_In, 号序_In, d_Date);
          --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,v_排队标记,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, r_安排.科室id, v_排队号码, Null, r_Pati.姓名, r_Pati.病人id, v_诊室, r_安排.医生姓名,
                           d_排队时间, 预约方式_In, n_启用分时段, v_排队序号);
        End If;
      End If;
    End If;
  
    If Nvl(操作方式_In, 0) = 1 Then
      --处理票据使用情况
      If 票据号_In Is Not Null Then
        Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
        --发出票据
        Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, 单据号_In);
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
        Values
          (票据使用明细_Id.Nextval, Decode(收费票据_In, 1, 1, 4), 票据号_In, 1, 1, 领用id_In, n_打印id, d_登记时间, v_操作员姓名);
        --状态改动
        Update 票据领用记录
        Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
        Where ID = Nvl(领用id_In, 0);
      End If;
      --病人本次就诊(以发生时间为准)
      If Nvl(r_Pati.病人id, 0) <> 0 Then
        Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = v_诊室 Where 病人id = r_Pati.病人id;
      End If;
    End If;
  End If;
  --病人挂号汇总
  --解锁单据时不用再对汇总单据进行统计了 在锁定单据时已经进行了汇总
  If 锁定类型_In <> 2 Then
    --操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
    --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收;3-收费预约
    n_预约 := Case
              When Nvl(操作方式_In, 0) = 1 Then
               0
              When Nvl(操作方式_In, 0) = 2 Then
               1
              When Nvl(操作方式_In, 0) = 3 Then
               3
              Else
               0
            End;
    Zl_病人挂号汇总_Update(r_安排.医生姓名, r_安排.医生id, r_安排.项目id, r_安排.科室id, 发生时间_In, n_预约, 号码_In);
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 1, n_挂号id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方机构挂号_Insert;
/

--91633:余伟节,2015-12-25,临床路径取消特殊项目【未生成任何项目】
Create Or Replace Procedure Zl_病人路径生成_Insert
(
  序号_In           Number, --医嘱界面产生路径项目时，序号为0
  病人id_In         病人临床路径.病人id%Type,
  主页id_In         病人临床路径.主页id%Type,
  婴儿_In           电子病历记录.婴儿%Type,
  科室id_In         病人临床路径.科室id%Type,
  路径记录id_In     病人路径执行.路径记录id%Type,
  阶段id_In         病人路径执行.阶段id%Type,
  日期_In           病人路径执行.日期%Type,
  天数_In           病人路径执行.天数%Type,
  分类_In           病人路径执行.分类%Type,
  项目id_In         病人路径执行.项目id%Type,
  医嘱ids_In        Varchar2,
  病历文件ids_In    Varchar2,
  病人病历ids_In    Varchar2,
  登记人_In         病人路径执行.登记人%Type,
  登记时间_In       病人路径执行.登记时间%Type,
  项目内容_In       病人路径执行.项目内容%Type := Null,
  执行者_In         病人路径执行.执行者%Type := Null,
  项目结果_In       病人路径执行.项目结果%Type := Null,
  图标id_In         病人路径执行.图标id%Type := Null,
  添加原因_In       病人路径执行.添加原因%Type := Null,
  变异原因_In       病人路径执行.变异原因%Type := Null,
  自动执行_In       Number := 0,
  电子病历id_In     电子病历记录.Id%Type := Null,
  合并路径阶段s_In  Varchar2 := Null, --用于修改合并路径的当前阶段ID，格式：合并路径记录ID:阶段ID,合并路径记录ID:阶段ID。。。。
  合并路径记录id_In 病人路径执行.合并路径记录id%Type := Null,
  合并路径阶段id_In 病人路径执行.合并路径阶段id%Type := Null,
  插入位置id_In     病人路径执行.Id%Type := 0,
  生成者_In         病人路径执行.生成者 %Type := 1,
  任务ids_In        Varchar2 := Null,
  生成时间性质_In   病人路径执行.生成时间性质%Type := Null --1-补录,2-暂存
) Is
  v_当前阶段id 病人临床路径.当前阶段id%Type;
  v_路径执行id 病人路径执行.Id%Type;
  v_病历id     电子病历记录.Id%Type;
  t_Advice     t_Numlist;
  t_File       t_Numlist;
  t_Doc        t_Numlist;

  v_Id             电子病历内容.Id%Type;
  v_父id           电子病历内容.父id%Type;
  v_当前父id       电子病历内容.父id%Type;
  v_原对象序号     电子病历内容.父id%Type;
  v_内容文本       电子病历内容.内容文本%Type;
  v_执行环节       Varchar2(20);
  n_当前天数       病人临床路径.当前天数%Type;
  n_合并路径记录id 病人路径执行.合并路径记录id%Type;
  n_合并路径阶段id 病人路径执行.合并路径阶段id%Type;
  n_天数           病人临床路径.当前天数%Type;
  v_合并路径阶段s  Varchar2(255);

  v_项目序号 病人路径执行.项目序号%Type;
  n_Count    Number;
  n_Minnum   Number;
  v_Error    Varchar2(255);
  Err_Custom Exception;

  --项目序号处理
  Procedure p_Sort_项目序号
  (
    项目序号_In In 病人路径执行.项目序号%Type,
    执行id_In   In 病人路径执行.Id%Type
  ) Is
    n_Num Number;
  Begin
    n_Num := 项目序号_In;
    For r_Outpathitem In (Select a.Id, Nvl(a.项目序号, b.项目序号) As 项目序号
                          From 病人路径执行 A, 临床路径项目 B
                          Where a.路径记录id = 路径记录id_In And a.阶段id = 阶段id_In And a.天数 = 天数_In And a.分类 = 分类_In And
                                a.项目id = b.Id(+) And Nvl(a.项目序号, b.项目序号) >= 项目序号_In
                          Order By Nvl(a.项目序号, b.项目序号)) Loop
      n_Num := n_Num + 1;
      --1-从插入位置处之后的所有路径外项目序号加 1
      Update 病人路径执行 A Set a.项目序号 = n_Num Where a.Id = r_Outpathitem.Id;
    End Loop;
    Update 病人路径执行 A Set a.项目序号 = 项目序号_In Where a.Id = 执行id_In;
  Exception
    When Others Then
      Null;
  End p_Sort_项目序号;
Begin
  If 序号_In = 1 And (项目内容_In Is Null Or 项目内容_In = '未生成任何项目' Or 项目内容_In = '路径外项目') Then
    --合并路径
    If 合并路径阶段s_In Is Not Null Then
      Select Nvl(当前天数, 1) Into n_当前天数 From 病人临床路径 Where ID = 路径记录id_In;
      --求出增量(首要路径提前合并路径就提前，首要路径延后，合并路径就延后)
      n_天数          := 天数_In - n_当前天数;
      v_合并路径阶段s := 合并路径阶段s_In || ',';
      While v_合并路径阶段s Is Not Null Loop
        n_合并路径记录id := To_Number(Substr(v_合并路径阶段s, 1, Instr(v_合并路径阶段s, ':') - 1));
        n_合并路径阶段id := To_Number(Substr(v_合并路径阶段s, Instr(v_合并路径阶段s, ':') + 1,
                                       Instr(v_合并路径阶段s, ',') - Instr(v_合并路径阶段s, ':') - 1));
        Select Nvl(当前阶段id, 0) Into v_当前阶段id From 病人合并路径 Where ID = n_合并路径记录id;
        If v_当前阶段id <> n_合并路径阶段id Then
          Update 病人合并路径 Set 前一阶段id = 当前阶段id, 当前阶段id = n_合并路径阶段id Where ID = n_合并路径记录id;
        End If;
        Update 病人合并路径 Set 当前天数 = Nvl(当前天数, 1) + n_天数 Where ID = n_合并路径记录id;
      
        v_合并路径阶段s := Substr(v_合并路径阶段s, Instr(v_合并路径阶段s, ',') + 1);
      End Loop;
    End If;
    --首要路径
    If 生成者_In = 1 Then
      Select Nvl(当前阶段id, 0) Into v_当前阶段id From 病人临床路径 Where ID = 路径记录id_In;
      If v_当前阶段id <> 阶段id_In Then
        Update 病人临床路径 Set 前一阶段id = 当前阶段id, 当前阶段id = 阶段id_In Where ID = 路径记录id_In;
      End If;
      Update 病人临床路径 Set 当前天数 = 天数_In Where ID = 路径记录id_In;
    End If;
  End If;

  --添加的路径外项目:即使有可选的项目可能还未生成,序号占用见后面补充项目调序处理
  If 项目内容_In Is Not Null Then
    Select Max(Nvl(a.项目序号, b.项目序号)) + 1
    Into v_项目序号
    From 病人路径执行 A, 临床路径项目 B
    Where a.路径记录id = 路径记录id_In And a.阶段id = 阶段id_In And a.天数 = 天数_In And a.分类 = 分类_In And a.项目id = b.Id(+);
  End If;

  v_路径执行id := 0;
  If 序号_In = 0 And 项目内容_In Is Null Then
    --加max是为了容错以前的数据，实际上同一项目在当天只有一条执行记录
    Select Nvl(Max(ID), 0)
    Into v_路径执行id
    From 病人路径执行
    Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 天数 = 天数_In And 项目id = 项目id_In;
  End If;

  --医嘱界面添加的非路径外项目
  If v_路径执行id = 0 Then
    Select Count(1)
    Into n_Count
    From 病人路径执行
    Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In;
    If n_Count = 0 Then
      --首次生成路径项目，转存前一天的暂存项目
      Update 病人路径执行
      Set 阶段id = 阶段id_In, 日期 = 日期_In, 天数 = 天数_In, 项目序号 = Null
      Where ID In (Select ID From 病人路径执行 Where 路径记录id = 路径记录id_In And 生成时间性质 = 2);
      --修改暂存标识
      Update 病人路径执行
      Set 生成时间性质 = Null
      Where ID In (Select a.Id
                   From 病人路径执行 A, 病人路径医嘱 B, 病人医嘱记录 C
                   Where a.Id = b.路径执行id And b.病人医嘱id = c.Id And a.路径记录id = 路径记录id_In And a.生成时间性质 = 2 And
                         a.日期 = Trunc(c.开始执行时间));
    End If;
    Select 病人路径执行_Id.Nextval Into v_路径执行id From Dual;
    Insert Into 病人路径执行
      (ID, 路径记录id, 阶段id, 日期, 天数, 分类, 项目id, 登记人, 登记时间, 项目序号, 项目内容, 执行者, 生成者, 项目结果, 图标id, 添加原因, 变异原因, 合并路径记录id, 合并路径阶段id,
       生成时间性质)
    Values
      (v_路径执行id, 路径记录id_In, 阶段id_In, 日期_In, 天数_In, 分类_In, 项目id_In, 登记人_In, 登记时间_In, v_项目序号, 项目内容_In, 执行者_In, 生成者_In,
       项目结果_In, 图标id_In, 添加原因_In, 变异原因_In, 合并路径记录id_In, 合并路径阶段id_In, 生成时间性质_In);
  
    --路径外项目序号插入 排序
    If 插入位置id_In <> 0 Then
      --获取要插入的序号
      Select Nvl(a.项目序号, b.项目序号)
      Into v_项目序号
      From 病人路径执行 A, 临床路径项目 B
      Where a.Id = 插入位置id_In And a.项目id = b.Id(+);
      --序号调整
      p_Sort_项目序号(v_项目序号, v_路径执行id);
    End If;
    --路径项目补充生成时,序号重整:假如临床路径项目存在A1,A2,A3这3个项目,首次生成A1,A2后,再生成路径外项目B1,B2,同时将B1,B2插入到A1的位置
    --         那么此时病人路径执行中的序号变为:B1(1),B2(2),A1(3),A2(4),如果再补充生成A3时,路径显示顺序变为：B1(1),B2(2),A1(3),A3(3),A2(4)
    --         这样就会出现路径项目中补充生成的A3不能按照临床路径项目的顺序A1,A2,A3 正确排序。
  
    --当前阶段，当前天数，当前分类下，存在路径内项目且路径内的项目序号被重新调整过。（未添加路径外项目时，路径内项目的序号为空）
    Select Nvl(Count(ID), 0)
    Into n_Count
    From 病人路径执行
    Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 天数 = 天数_In And 分类 = 分类_In And 项目id Is Not Null And 项目序号 Is Not Null;
    --补充生成的路径项目序号重整
    If n_Count > 0 And 项目id_In Is Not Null Then
      --查找补充生成的路径项目,应该插入的位置
      Select Min(b.项目序号)
      Into n_Minnum
      From 病人路径执行 A, 临床路径项目 B
      Where a.路径记录id = 路径记录id_In And a.阶段id = 阶段id_In And a.天数 = 天数_In And a.分类 = 分类_In And a.项目id = b.Id;
    
      Select 项目序号 Into v_项目序号 From 临床路径项目 Where ID = 项目id_In;
      --确定该路径项目序号插入的位置：
      If v_项目序号 = n_Minnum Then
        --v_项目序号 = n_Minnum：病人路径执行记录已在此排序前插入到数据库，插入的这天数据就是最小的这条数据
        Select 项目序号
        Into v_项目序号
        From (Select Nvl(a.项目序号, b.项目序号) As 项目序号
               From 病人路径执行 A, 临床路径项目 B
               Where a.路径记录id = 路径记录id_In And a.阶段id = 阶段id_In And a.天数 = 天数_In And a.分类 = 分类_In And a.项目id = b.Id And
                     b.项目序号 > n_Minnum
               Order By b.项目序号)
        Where Rownum = 1;
      Else
        Select 项目序号
        Into v_项目序号
        From (Select Nvl(a.项目序号, b.项目序号) As 项目序号
               
               From 病人路径执行 A, 临床路径项目 B
               Where a.路径记录id = 路径记录id_In And a.阶段id = 阶段id_In And a.天数 = 天数_In And a.分类 = 分类_In And a.项目id = b.Id And
                     b.项目序号 < v_项目序号
               Order By b.项目序号 Desc)
        Where Rownum = 1;
        v_项目序号 := v_项目序号 + 1;
      End If;
      p_Sort_项目序号(v_项目序号, v_路径执行id);
    End If;
  
    --如果是自动执行模式（连续提前多个阶段时调用）;补录路径外项目
    If 自动执行_In = 1 Then
      Select zl_GetSysParameter('是否启用路径执行环节', 1256) Into v_执行环节 From Dual;
      If v_执行环节 = '1' Then
        Select zl_GetSysParameter('路径执行环节启用场合', 1256) Into v_执行环节 From Dual;
      
        Select Nvl(Nvl(a.执行者, b.执行者), 0)
        Into n_Count
        From 病人路径执行 A, 临床路径项目 B
        Where a.项目id = b.Id(+) And a.Id = v_路径执行id;
        --当前执行者符合启用场合自动执行,当执行者取不到值时,统一处理。
        If n_Count = 0 Or Substr(v_执行环节, n_Count, 1) = '1' Then
          Update 病人路径执行
          Set 执行人 = 登记人_In, 执行时间 = 登记时间_In, 执行结果 = '已经执行', 执行说明 = '自动执行。'
          Where ID = v_路径执行id;
        End If;
      End If;
    End If;
  End If;
  --删除特殊项目：未生成任何项目（如果当前阶段，当前日期存在其他项目，需删除“未生成任何项目”）
  Select Count(ID)
  Into n_Count
  From 病人路径执行 T
  Where t.路径记录id = 路径记录id_In And t.阶段id = 阶段id_In And t.天数 = 天数_In And NVL(t.项目内容,'路径内项目') = '未生成任何项目';

  If n_Count > 0 Then
    Select Count(ID)
    Into n_Count
    From 病人路径执行 T
    Where t.路径记录id = 路径记录id_In And t.阶段id = 阶段id_In And t.天数 = 天数_In And NVL(t.项目内容,'路径内项目') <> '未生成任何项目';
    If n_Count > 0 Then
      Delete From 病人路径执行 T
      Where t.路径记录id = 路径记录id_In And t.阶段id = 阶段id_In And t.天数 = 天数_In And NVL(t.项目内容,'路径内项目') = '未生成任何项目';
    End If;
  End If;

  If 医嘱ids_In Is Not Null Then
    Select Column_Value Bulk Collect Into t_Advice From Table(f_Num2list(医嘱ids_In));
    Forall I In 1 .. t_Advice.Count
      Insert Into 病人路径医嘱 (路径执行id, 病人医嘱id) Values (v_路径执行id, t_Advice(I));
  End If;

  If 病人病历ids_In Is Not Null Then
    Select Column_Value Bulk Collect Into t_Doc From Table(f_Num2list(病人病历ids_In));
    Select Column_Value Bulk Collect Into t_File From Table(f_Num2list(病历文件ids_In));
    For I In 1 .. t_Doc.Count Loop
      v_病历id := t_Doc(I);
    
      Insert Into 电子病历记录
        (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 保存人, 保存时间, 最后版本, 签名级别, 编辑方式, 路径执行id)
        Select v_病历id, 2, 病人id_In, 主页id_In, 婴儿_In, 科室id_In, 种类, ID, 名称, 登记人_In, 登记时间_In, 登记人_In, 登记时间_In, 1, 0, Decode(保留,2,1,0),
               v_路径执行id
        From 病历文件列表
        Where ID = t_File(I);

      For Rs In (Select ID, 文件id, Nvl(父id, 0) As 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机,
                        诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域
                 From 病历文件结构
                 Where 文件id = t_File(I)
                 Order By 对象序号) Loop

        Select 电子病历内容_Id.Nextval Into v_Id From Dual;
      
        If Rs.父id = 0 Then
          v_当前父id := v_Id;
          v_父id     := Null;
        Else
          --对象序号为空的时候，父ID就不是按照顺序的了，需要重新查找
          If Rs.对象序号 Is Null Then
            Select 对象序号 Into v_原对象序号 From 病历文件结构 Where ID = Rs.父id;
            If v_原对象序号 Is Null Then
              v_父id := Null;
            Else
              Select ID Into v_父id From 电子病历内容 Where 文件id = v_病历id And 对象序号 = v_原对象序号;
            End If;
          Else
            v_父id := v_当前父id;
          End If;
        End If;
      
        If Rs.对象类型 = 4 And Rs.替换域 = 1 Then
          v_内容文本 := Zl_Replace_Element_Value(Rs.要素名称, 病人id_In, 主页id_In, 2, Null, 婴儿_In);
        Else
          v_内容文本 := Rs.内容文本;
        End If;
      
        Insert Into 电子病历内容
          (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 定义提纲id, 复用提纲, 使用时机, 诊治要素id,
           替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域)
        Values
          (v_Id, v_病历id, 1, 0, v_父id, Rs.对象序号, Rs.对象类型, Rs.对象标记, Rs.保留对象, Rs.对象属性, Rs.内容行次, v_内容文本, Rs.是否换行, Rs.预制提纲id,
           Decode(Rs.父id, 0, Rs.Id, Null), Rs.复用提纲, Rs.使用时机, Rs.诊治要素id, Rs.替换域, Rs.要素名称, Rs.要素类型, Rs.要素长度, Rs.要素小数,
           Rs.要素单位, Rs.要素表示, Rs.输入形态, Rs.要素值域);
      
        If Rs.对象类型 = 5 Then
          Insert Into 电子病历图形 (对象id, 图形) Values (v_Id, (Select 图形 From 病历文件图形 Where 对象id = Rs.Id));
        End If;
      
      End Loop;
    
      Insert Into 电子病历格式
        (文件id, 内容)
      Values
        (v_病历id, (Select 内容 From 病历文件格式 Where 文件id = t_File(I)));
    End Loop;
  End If;

  If Nvl(电子病历id_In, 0) <> 0 Then
    Update 电子病历记录 Set 路径执行id = v_路径执行id Where ID = 电子病历id_In;
  End If;
  If 任务ids_In Is Not Null Then
    For Rs In (Select /*+ Rule*/
                Column_Value As 任务id
               From Table(Cast(f_Str2list(任务ids_In, ',') As Zltools.t_Strlist))) Loop
      Insert Into 病人路径病历 (路径执行id, 任务id) Values (v_路径执行id, Rs.任务id);
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人路径生成_Insert;
/

--92410:刘尔旋,2016-01-05,费用状态调整
--91842:刘尔旋,2016-01-18,门诊转住院处理
Create Or Replace Procedure Zl_门诊转住院_收费转出
(
  No_In         住院费用记录.No%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  退费时间_In   住院费用记录.发生时间%Type,
  门诊退费_In   Number := 0,
  入院科室id_In 住院费用记录.开单部门id%Type := Null,
  主页id_In     住院费用记录.主页id%Type := Null,
  结算方式_In   病人预交记录.结算方式%Type := Null,
  结帐id_In     病人预交记录.结帐id%Type := Null,
  原结帐id_In   病人预交记录.结帐id%Type := Null,
  误差费_In     病人预交记录.冲预交%Type := Null
) As
  --门诊退费_In:0-门诊转住院立即销帐;1-门诊退费模式
  -- 门诊退费_In为1时:入院科室id_In和主页ID_IN可以不传入
  n_Count      Number(5);
  n_原结帐id   住院费用记录.结帐id%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  n_预交使用额 病人预交记录.冲预交%Type;
  n_实际冲销   病人预交记录.冲预交%Type;
  n_组id       财务缴款分组.Id%Type;
  n_病人id     病人信息.病人id%Type;
  v_预交no     病人预交记录.No%Type;
  n_预交金额   病人预交记录.冲预交%Type;
  n_打印id     票据使用明细.打印id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  v_开单人     门诊费用记录.开单人%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_误差费     病人预交记录.冲预交%Type;
  v_误差费     结算方式.名称%Type;
  n_返回值     病人余额.费用余额%Type;
  v_结算方式   结算方式.名称%Type;
  v_Nos        Varchar2(3000);
  v_结帐ids    Varchar2(3000);
  v_原结帐ids  Varchar2(3000);
  n_Tempid     病人预交记录.Id%Type;
  n_预交id     病人预交记录.Id%Type;
  n_医保       Number;
  n_存在       Number;
  n_退现       Number;
  n_部分退费   Number;
  n_退费条数   Number;
  n_异常标志   Number;
  n_计算误差   Number;
  n_费用状态   门诊费用记录.费用状态%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  Procedure Zl_Square_Update
  (
    结帐ids_In    Varchar2,
    现结帐id_In   病人预交记录.结帐id%Type,
    缴款组id_In   病人预交记录.缴款组id%Type,
    退款时间_In   病人预交记录.收款时间%Type,
    结算序号_In   病人预交记录.结算序号%Type,
    结算内容_In   Varchar2 := Null,
    退费金额_In   病人预交记录.冲预交%Type := Null,
    结算卡序号_In 病人预交记录.结算卡序号%Type := Null
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
    For v_校对 In (Select Min(a.Id) As 预交id, c.消费卡id, Sum(c.结算金额) As 结算金额, c.接口编号, c.卡号, Max(c.序号) As 序号, Max(c.Id) As ID
                 From 病人预交记录 A, 病人卡结算对照 B, 病人卡结算记录 C
                 Where a.Id = b.预交id And a.结算卡序号 = 结算卡序号_In And b.卡结算id = c.Id And a.记录性质 = 3 And
                       Instr(Nvl(结算内容_In, '_LXH'), ',' || a.结算方式 || ',') = 0 And
                       a.结帐id In (Select Column_Value From Table(f_Str2list(结帐ids_In)))
                 Group By c.消费卡id, c.接口编号, c.卡号) Loop
    
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
                 -1 * 退费金额_In, 现结帐id_In, 缴款组id_In, 预交类别, 卡类别id, Nvl(结算卡序号, v_校对.接口编号), 卡号, 交易流水号, 交易说明, 合作单位, 2, 结算序号_In,
                 结算性质
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
        Update 消费卡目录 Set 余额 = Nvl(余额, 0) + 退费金额_In Where ID = Nvl(v_校对.消费卡id, 0);
      End If;
    
      Select 病人卡结算记录_Id.Nextval Into n_Id From Dual;
      Insert Into 病人卡结算记录
        (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
        Select n_Id, 接口编号, 消费卡id, 序号, n_记录状态, 结算方式, -1 * 退费金额_In, 卡号, 交易流水号, 交易时间, 备注,
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
  n_组id := Zl_Get组id(操作员姓名_In);
  --误差费
  Begin
    Select 名称 Into v_误差费 From 结算方式 Where 性质 = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := '没有发现误差结算方式，请检查是否正确设置！';
      Raise Err_Item;
  End;

  If 原结帐id_In Is Null Then
  
    Select Count(NO), Sum(实收金额) Into n_Count, n_实收金额 From 门诊费用记录 Where NO = No_In And 记录性质 = 1;
    If n_Count = 0 Or n_实收金额 = 0 Then
      v_Err_Msg := '单据' || No_In || '不是收费单据或因并发原因他人操作了该单据,不能转为住院费用.';
      Raise Err_Item;
    End If;
  
    Select 结帐id, 病人id, 开单部门id, 开单人
    Into n_原结帐id, n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Rownum < 2;
  
    --1.1作废费用记录
    If 结帐id_In Is Null Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Else
      n_结帐id := 结帐id_In;
    End If;
  
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
       计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 执行时间,
       操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id, 费用状态)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id,
             收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, -1 * 应收金额, -1 * 实收金额, 开单部门id,
             开单人, 执行部门id, 划价人, 执行人, -1, 执行时间, 操作员编号_In, 操作员姓名_In, 发生时间, 退费时间_In, n_结帐id, -1 * 结帐金额, 保险项目否, 保险大类id, 统筹金额,
             摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id, 0
      From 门诊费用记录
      Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
  
    --Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
  
    --1.2作废预交记录
    --作废冲预交部分
    For r_结账id In (Select Distinct 结帐id
                   From 门诊费用记录
                   Where NO In (Select Distinct NO
                                From 门诊费用记录
                                Where 结帐id In (Select 结帐id
                                               From 病人预交记录
                                               Where 结算序号 In (Select b.结算序号
                                                              From 门诊费用记录 A, 病人预交记录 B
                                                              Where a.No = No_In And b.结算序号 < 0 And Mod(a.记录性质, 10) = 1 And
                                                                    a.记录状态 <> 0 And a.结帐id = b.结帐id))) And
                         Mod(记录性质, 10) = 1 And 记录状态 <> 0
                   Union
                   Select Distinct 结帐id
                   From 门诊费用记录
                   Where NO In (Select Distinct NO
                                From 门诊费用记录
                                Where 结帐id In (Select a.结帐id
                                               From 门诊费用记录 A, 病人预交记录 B
                                               Where a.No = No_In And b.结算序号 > 0 And Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And
                                                     a.结帐id = b.结帐id))) Loop
      v_原结帐ids := v_原结帐ids || ',' || r_结账id.结帐id;
    End Loop;
    v_原结帐ids := Substr(v_原结帐ids, 2);
  
    Begin
      Select 1
      Into n_医保
      From 保险结算记录
      Where 记录id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And Rownum < 2;
    Exception
      When Others Then
        n_医保 := 0;
    End;
  
    If n_医保 = 1 Then
      Begin
        Select 1
        Into n_存在
        From 医保结算明细
        Where NO = No_In And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And Rownum < 2;
      Exception
        When Others Then
          v_Err_Msg := '当前单据' || No_In || '不存在医保结算明细,无法进行门诊转住院!';
          Raise Err_Item;
      End;
    End If;
  
    --医保退款
    For r_医保 In (Select 结帐id, NO, 结算方式, 金额, 备注
                 From 医保结算明细
                 Where NO = No_In And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids)))) Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) - r_医保.金额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_医保.结算方式
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_医保.结算方式, 1, -1 * r_医保.金额);
        n_返回值 := r_医保.金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_医保.结算方式 And Nvl(余额, 0) = 0;
      End If;
    
      Update 病人预交记录
      Set 冲预交 = 冲预交 + (-1 * r_医保.金额)
      Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = r_医保.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
           缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
        Values
          (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_医保.金额, r_医保.结算方式, Null, 退费时间_In,
           Null, Null, Null, 操作员编号_In, 操作员姓名_In, r_医保.备注, n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id,
           0, 3);
      End If;
    
      Update 病人预交记录
      Set 记录状态 = 3
      Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
            结算方式 = r_医保.结算方式;
    
      Update 医保结算明细
      Set 金额 = 金额 + (-1 * r_医保.金额)
      Where NO = No_In And 结帐id = n_结帐id And 结算方式 = r_医保.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额)
        Values
          (n_结帐id, No_In, r_医保.结算方式, -1 * r_医保.金额);
      End If;
      n_实收金额 := n_实收金额 - r_医保.金额;
    End Loop;
  
    Begin
      Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
    Exception
      When Others Then
        Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
    End;
  
    If n_实收金额 <> 0 Then
      For r_Prepay In (Select NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号,
                              卡号, 交易流水号, 交易说明, 合作单位
                       From 病人预交记录 A
                       Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids)))
                       Group By n_Tempid, NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 卡类别id, 结算卡序号, 卡号,
                                交易流水号, 交易说明, 合作单位) Loop
        If n_实收金额 <> 0 Then
          If r_Prepay.冲预交 >= n_实收金额 Then
            Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 缴款组id)
              Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                     r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                     操作员编号_In, -1 * n_实收金额, n_结帐id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                     r_Prepay.交易说明, r_Prepay.合作单位, 1, -1 * n_结帐id, n_组id
              From Dual;
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(n_实收金额, 0)
            Where 病人id = n_病人id And 类型 = 1 And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, n_实收金额, 1);
              n_返回值 := n_实收金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
            n_实收金额 := 0;
          Else
            Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 缴款组id)
              Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                     r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                     操作员编号_In, -1 * r_Prepay.冲预交, n_结帐id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                     r_Prepay.交易说明, r_Prepay.合作单位, 1, -1 * n_结帐id, n_组id
              From Dual;
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Prepay.冲预交, 0)
            Where 病人id = n_病人id And 类型 = 1 And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, r_Prepay.冲预交, 1);
              n_返回值 := r_Prepay.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
            n_实收金额 := n_实收金额 - r_Prepay.冲预交;
          End If;
        End If;
      End Loop;
    End If;
    --2.票据收回
    --可能以前没有打印,无收回
    Select Nvl(Max(ID), 0)
    Into n_打印id
    From (Select b.Id
           From 票据使用明细 A, 票据打印内容 B
           Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = No_In
           Order By a.使用时间 Desc)
    Where Rownum < 2;
    If n_打印id > 0 Then
      --多张单据循环调用时只能收回一次
      Select Count(打印id) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
      If n_Count = 0 Then
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 退费时间_In, 操作员姓名_In
          From 票据使用明细
          Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
      End If;
    End If;
  
    --3.缴款数据处理(
    --   现有两种情况:
    --    1. 转出过程直接销帐的,则缴款数据不增加;
    --    2. 先转出,再到门诊退款退票,则需要进行缴款数据处理
    If Nvl(门诊退费_In, 0) = 1 Then
      For c_预交 In (Select a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, Min(a.交易流水号) As 交易流水号,
                          Min(a.交易说明) As 交易说明, Min(a.合作单位) As 合作单位, b.性质
                   From 病人预交记录 A, 结算方式 B
                   Where a.记录性质 = 3 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                         a.结算方式 = b.名称 And b.性质 In (1, 2, 7, 8) And a.结算方式 Is Not Null
                   Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质
                   Having Sum(a.冲预交) <> 0
                   Order By a.卡类别id, 性质 Desc) Loop
        If n_实收金额 <> 0 Then
          Begin
            Select 是否退现 Into n_退现 From 医疗卡类别 Where ID = c_预交.卡类别id;
          Exception
            When Others Then
              n_退现 := 0;
          End;
          If (c_预交.性质 = 7 Or (c_预交.性质 = 8 And c_预交.卡类别id Is Not Null)) And n_退现 = 0 Then
            If c_预交.冲预交 > n_实收金额 Then
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '1' || ',' || c_预交.卡类别id || ',' || -1 * n_实收金额 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || c_预交.卡类别id || ',' || -1 * n_实收金额 || '|', n_组id,
                   Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = c_预交.结算方式;
              n_费用状态 := 1;
              n_实收金额 := 0;
            Else
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|', n_组id,
                   Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
            
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = c_预交.结算方式;
              n_费用状态 := 1;
              n_实收金额 := n_实收金额 - c_预交.冲预交;
            End If;
          Else
            n_实际冲销 := 0;
            If c_预交.性质 In (3, 4) Or (c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null) Then
              v_结算方式 := c_预交.结算方式;
            Else
              If 结算方式_In Is Null Then
                Begin
                  Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
                Exception
                  When Others Then
                    Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
                End;
              Else
                v_结算方式 := 结算方式_In;
              End If;
            End If;
          
            If c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null Then
              If n_实收金额 >= c_预交.冲预交 Then
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, c_预交.冲预交, c_预交.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null,
                     退费时间_In, Null, Null, Null, 操作员编号_In, 操作员姓名_In,
                     '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|', n_组id, Null, Null, Null, Null, Null, Null,
                     n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := c_预交.冲预交;
              Else
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, n_实收金额, c_预交.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '0' || ',' || c_预交.结算卡序号 || ',' || -1 * n_实收金额 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                     Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || c_预交.结算卡序号 || ',' || -1 * n_实收金额 || '|', n_组id,
                     Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := n_实收金额;
              End If;
            Else
              If c_预交.冲预交 > n_实收金额 Then
                n_实际冲销 := n_实收金额;
              Else
                n_实际冲销 := c_预交.冲预交;
              End If;
            End If;
          
            If c_预交.结算卡序号 Is Null Then
              Update 人员缴款余额
              Set 余额 = Nvl(余额, 0) - n_实际冲销
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
              Returning 余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 人员缴款余额
                  (收款员, 结算方式, 性质, 余额)
                Values
                  (操作员姓名_In, v_结算方式, 1, -1 * n_实际冲销);
                n_返回值 := n_实际冲销;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 人员缴款余额
                Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
              End If;
            
              --退原预交记录
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实际冲销)
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实际冲销, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, c_预交.合作单位, n_结帐id,
                   -1 * n_结帐id, 0, 3);
              End If;
            End If;
            Update 病人预交记录
            Set 记录状态 = 3
            Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                  结算方式 = c_预交.结算方式;
            n_实收金额 := n_实收金额 - n_实际冲销;
          End If;
        End If;
      End Loop;
    
      --更新费用审核记录
      Update 费用审核记录
      Set 记录状态 = 2
      Where 费用id In (Select ID From 门诊费用记录 Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3)) And 性质 = 1;
      --作废门诊记录
      Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
      For r_Clinic In (Select 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
                              发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, Sum(应收金额) As 应收金额,
                              Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id, 划价人, Max(记帐单id) As 记帐单id, 发生时间,
                              实际票号
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 1 And 记录状态 In (2, 3) And Nvl(附加标志, 0) Not In (8, 9)
                       Group By 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码,
                                费用类型, 发药窗口, 付数, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 开单部门id, 开单人, 执行部门id, 划价人, 发生时间, 实际票号
                       Having Sum(数次) <> 0) Loop
        Insert Into 门诊费用记录
          (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
           保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间,
           登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 缴款组id, 结帐id, 结帐金额, 费用状态)
        Values
          (病人费用记录_Id.Nextval, 1, No_In, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 1, r_Clinic.病人id,
           '', r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别, r_Clinic.收费细目id,
           r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数,
           -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价,
           -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 0, r_Clinic.开单部门id, r_Clinic.开单人, r_Clinic.发生时间,
           退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', n_组id, n_结帐id,
           -1 * r_Clinic.实收金额, 0);
      End Loop;
    Else
      --4.退款转预交(不产生票据,由操作员通过重打进行)
      For r_Pay In (Select Min(a.Id) As 预交id, a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号,
                           a.交易说明, a.合作单位, b.性质
                    From 病人预交记录 A, 结算方式 B
                    Where a.记录性质 = 3 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                          a.结算方式 = b.名称 And (b.性质 In (1, 2, 7, 8)) And a.结算方式 Is Not Null
                    Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质, a.交易流水号, a.交易说明, a.合作单位


                    
                    Having Sum(a.冲预交) <> 0
                    Order By a.卡类别id, 性质 Desc) Loop
        --4.1产生预交款单据 (不存在部分退费的情况)
        --所有单据,按规则生成预交款单据
        --因为收款后立即缴款,所以人员缴款余额无变化
        If n_实收金额 <> 0 Then
          If r_Pay.性质 = 7 Or (r_Pay.性质 = 8 And r_Pay.卡类别id Is Not Null) Then
            If r_Pay.冲预交 > n_实收金额 Then
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '1' || ',' || r_Pay.卡类别id || ',' || -1 * n_实收金额 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || r_Pay.卡类别id || ',' || -1 * n_实收金额 || '|', n_组id,
                   Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
            
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = r_Pay.结算方式;
              n_费用状态 := 1;
              n_实收金额 := 0;
            Else
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|',
                   n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
            
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = r_Pay.结算方式;
              n_费用状态 := 1;
              n_实收金额 := n_实收金额 - r_Pay.冲预交;
            End If;
          Else
            n_实际冲销 := 0;
            If r_Pay.性质 In (3, 4) Or (r_Pay.性质 = 8 And r_Pay.结算卡序号 Is Not Null) Then
              v_结算方式 := r_Pay.结算方式;
            Else
              If 结算方式_In Is Null Then
                Begin
                  Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
                Exception
                  When Others Then
                    Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
                End;
              Else
                v_结算方式 := 结算方式_In;
              End If;
            End If;
          
            If r_Pay.性质 = 8 And r_Pay.结算卡序号 Is Not Null Then
              If n_实收金额 >= r_Pay.冲预交 Then
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, r_Pay.冲预交, r_Pay.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null,
                     退费时间_In, Null, Null, Null, 操作员编号_In, 操作员姓名_In,
                     '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|', n_组id, Null, Null, Null, Null, Null,
                     Null, n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := r_Pay.冲预交;
              Else
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, n_实收金额, r_Pay.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * n_实收金额 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                     Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * n_实收金额 || '|', n_组id,
                     Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := n_实收金额;
              End If;
            Else
              If r_Pay.冲预交 > n_实收金额 Then
                n_实际冲销 := n_实收金额;
              Else
                n_实际冲销 := r_Pay.冲预交;
              End If;
            End If;
          
            If r_Pay.性质 Not In (3, 4, 7, 8) Then
              Update 病人预交记录
              Set 金额 = 金额 + n_实际冲销
              Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                v_预交no := Nextno(11);
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 预交类别)
                Values
                  (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, n_实际冲销, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, r_Pay.预交类别);
              End If;
            
              --病人余额
              Update 病人余额
              Set 预交余额 = Nvl(预交余额, 0) + n_实际冲销
              Where 性质 = 1 And 病人id = n_病人id And 类型 = 2
              Returning 预交余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, 2, n_实际冲销, 0);
                n_返回值 := n_实际冲销;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 病人余额
                Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
              End If;
            End If;
            --4.2缴款数据处理
            --   因为没有实际收病人的钱,所以不处理
            --部分退费情况，退原预交记录
            If r_Pay.性质 In (3, 4) Then
              Update 人员缴款余额
              Set 余额 = Nvl(余额, 0) - n_实际冲销
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式
              Returning 余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 人员缴款余额
                  (收款员, 结算方式, 性质, 余额)
                Values
                  (操作员姓名_In, r_Pay.结算方式, 1, -1 * n_实际冲销);
                n_返回值 := n_实际冲销;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 人员缴款余额
                Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式 And Nvl(余额, 0) = 0;
              End If;
            End If;
          
            If r_Pay.性质 <> 8 Then
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实际冲销)
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实际冲销, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, r_Pay.卡类别id, r_Pay.结算卡序号, r_Pay.卡号, r_Pay.交易流水号,
                   r_Pay.交易说明, r_Pay.合作单位, n_结帐id, -1 * n_结帐id, 0, 3);
              End If;
            End If;
          
            Update 病人预交记录
            Set 记录状态 = 3
            Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                  结算方式 = v_结算方式;
            n_实收金额 := n_实收金额 - n_实际冲销;
          
          End If;
        End If;
      End Loop;
    End If;
  
    If 误差费_In Is Not Null Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
         卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
      Values
        (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, 误差费_In, v_误差费, Null, 退费时间_In, Null, Null,
         Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 0, 3);
    End If;
    Delete From 病人预交记录
    Where 结帐id = n_结帐id And 记录性质 = 3 And 记录状态 = 2 And 冲预交 = 0 And 结算方式 Is Not Null;
    Delete From 病人预交记录 Where 结帐id = n_原结帐id And 摘要 = '预交临时记录' And 记录性质 = 3;
    Update 门诊费用记录 Set 费用状态 = Nvl(n_费用状态, 0) Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 2;
  Else
    --医保按结算转出
    For r_Nos In (Select Distinct a.No
                  From 门诊费用记录 A
                  Where Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And a.结帐id = 原结帐id_In) Loop
      v_Nos := v_Nos || ',' || r_Nos.No;
    End Loop;
    v_Nos := Substr(v_Nos, 2);
  
    For r_结帐ids In (Select Distinct a.结帐id
                    From 门诊费用记录 A
                    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                          a.记录状态 <> 0) Loop
      v_结帐ids := v_结帐ids || ',' || r_结帐ids.结帐id;
    End Loop;
    v_结帐ids := Substr(v_结帐ids, 2);
    Select Count(a.No), Sum(a.实收金额)
    Into n_Count, n_实收金额
    From 门诊费用记录 A
    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1;
    If n_Count = 0 Or n_实收金额 = 0 Then
      v_Err_Msg := '本次结算不是收费或因并发原因他人操作了该结算,不能转为住院费用.';
      Raise Err_Item;
    End If;
  
    Select 结帐id, 病人id, 开单部门id, 开单人
    Into n_原结帐id, n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where 结帐id = 原结帐id_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And Rownum < 2;
  
    Begin
      Select 1
      Into n_部分退费
      From 门诊费用记录 A
      Where Mod(a.记录性质, 10) = 1 And a.记录状态 = 2 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
            Rownum < 2;
    Exception
      When Others Then
        n_部分退费 := 0;
    End;
  
    Begin
      Select 0
      Into n_部分退费
      From 门诊费用记录 A
      Where 记录性质 = 11 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select Count(Avg(1))
      Into n_退费条数
      From 病人预交记录 A
      Where a.记录性质 = 3 And a.记录状态 <> 0 And 结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))
      Group By a.结算方式;
    Exception
      When Others Then
        n_退费条数 := 0;
    End;
    --1.1作废费用记录
    If 结帐id_In Is Null Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Else
      n_结帐id := 结帐id_In;
    End If;
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
       计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 执行时间,
       操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id, 费用状态)
      Select 病人费用记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄,
             a.标识号, a.付款方式, a.费别, a.病人科室id, a.收费类别, a.收费细目id, a.计算单位, a.付数, a.发药窗口, -1 * a.数次, a.加班标志, a.附加标志, a.收入项目id,
             a.收据费目, a.记帐费用, a.标准单价, -1 * a.应收金额, -1 * a.实收金额, a.开单部门id, a.开单人, a.执行部门id, a.划价人, a.执行人, -1, a.执行时间,
             操作员编号_In, 操作员姓名_In, a.发生时间, 退费时间_In, n_结帐id, -1 * a.结帐金额, a.保险项目否, a.保险大类id, a.统筹金额, a.摘要,
             Decode(Nvl(a.附加标志, 0), 9, 1, 0), a.保险编码, a.费用类型, n_组id, 0
      From 门诊费用记录 A
      Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And a.记录状态 = 1;
  
    --作废医保
    For r_医保 In (Select 结帐id, NO, 结算方式, 金额, 备注
                 From 医保结算明细
                 Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And
                       结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))) Loop
      Update 医保结算明细
      Set 金额 = 金额 + (-1 * r_医保.金额)
      Where NO = r_医保.No And 结帐id = r_医保.结帐id And 结算方式 = r_医保.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额)
        Values
          (r_医保.结帐id, r_医保.No, r_医保.结算方式, -1 * r_医保.金额);
      End If;
    End Loop;
  
    --Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
    --1.2作废预交记录
    --作废冲预交部分
    If n_部分退费 = 0 And Nvl(门诊退费_In, 0) = 0 Then
      For r_Prepay In (Select NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, -1 * Sum(冲预交) As 冲预交,
                              卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质
                       From 病人预交记录 A
                       Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                             Nvl(冲预交, 0) <> 0
                       Group By n_Tempid, NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, 卡类别id, 结算卡序号,
                                卡号, 交易流水号, 交易说明, 合作单位, 结算性质) Loop
        Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 预交类别, 结算性质)
          Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                 r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                 操作员编号_In, r_Prepay.冲预交, n_结帐id, n_组id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                 r_Prepay.交易说明, r_Prepay.合作单位, -1 * n_结帐id, 1, r_Prepay.结算性质
          From Dual;
      End Loop;
    
      For v_预交 In (Select 病人id, Nvl(预交类别, 2) As 预交类别, Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交金额
                   From 病人预交记录 A
                   Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                         a.结帐id <> n_结帐id
                   Group By 病人id, Nvl(预交类别, 2)
                   Having Sum(Nvl(冲预交, 0)) <> 0) Loop
      
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_预交.预交金额, 0)
        Where 病人id = v_预交.病人id And 类型 = Nvl(v_预交.预交类别, 2) And 性质 = 1
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 类型, 预交余额, 性质)
          Values
            (v_预交.病人id, Nvl(v_预交.预交类别, 2), v_预交.预交金额, 1);
          n_返回值 := v_预交.预交金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 病人余额
          Where 病人id = v_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      End Loop;
    Else
      If n_退费条数 = 0 And Nvl(门诊退费_In, 0) = 0 Then
        --只使用了预交，原样退回预交
        For r_Prepay In (Select NO, 实际票号, 病人id, 主页id, 科室id, Max(结算方式) As 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间,
                                -1 * Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质
                         From 病人预交记录 A
                         Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                               Nvl(冲预交, 0) <> 0
                         Group By n_Tempid, NO, 实际票号, 病人id, 主页id, 科室id, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, 卡类别id, 结算卡序号, 卡号,
                                  交易流水号, 交易说明, 合作单位, 结算性质) Loop
          Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
             结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 预交类别, 结算性质)
            Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                   r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                   操作员编号_In, r_Prepay.冲预交, n_结帐id, n_组id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                   r_Prepay.交易说明, r_Prepay.合作单位, -1 * n_结帐id, 1, r_Prepay.结算性质
            From Dual;
          Select -1 * 冲预交 Into n_预交金额 From 病人预交记录 Where ID = n_Tempid;
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + Nvl(n_预交金额, 0)
          Where 病人id = r_Prepay.病人id And 类型 = 1 And 性质 = 1
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, n_预交金额, 1);
            n_返回值 := n_预交金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = r_Prepay.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
          End If;
        End Loop;
      Else
        Begin
          Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
        Exception
          When Others Then
            Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
        End;
        Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
          Select n_Tempid, Max(NO), Max(实际票号), 3, 3, 病人id, 主页id, 科室id, Null, v_结算方式, Max(结算号码), '预交临时记录', Null, Null,
                 Null, Max(收款时间), 操作员姓名_In, 操作员编号_In, Sum(冲预交), n_原结帐id, Null, Null, Null, Null, Null, Null,
                 -1 * n_原结帐id, 3
          From 病人预交记录 A
          Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                Nvl(冲预交, 0) <> 0
          Group By n_Tempid, 3, 3, 病人id, 主页id, 科室id, Null, v_结算方式, '预交临时记录', 操作员姓名_In, 操作员编号_In, n_原结帐id;
      End If;
    End If;
  
    --作废门诊缴费及医保部分
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退费时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
             0, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * n_结帐id, 结算性质
      From 病人预交记录 A, 结算方式 B
      Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
            a.结算方式 = b.名称 And b.性质 Not In (7, 8);
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质, 校对标志)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退费时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
             0, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * n_结帐id, 结算性质, 1
      From 病人预交记录 A, 结算方式 B
      Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
            a.结算方式 = b.名称 And b.性质 = 7;
    If Sql%RowCount <> 0 Then
      n_费用状态 := 1;
    End If;
  
    Update 病人预交记录
    Set 记录状态 = 3
    Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)));
  
    --2.票据收回
    --可能以前没有打印,无收回
    For r_Nos In (Select Distinct a.No
                  From 门诊费用记录 A
                  Where Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And
                        a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))) Loop
    
      Select Nvl(Max(ID), 0)
      Into n_打印id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = r_Nos.No
             Order By a.使用时间 Desc)
      Where Rownum < 2;
      If n_打印id > 0 Then
        --多张单据循环调用时只能收回一次
        Select Count(打印id) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
        If n_Count = 0 Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 退费时间_In, 操作员姓名_In
            From 票据使用明细
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
        End If;
      End If;
    End Loop;
  
    --3.缴款数据处理(
    --   现有两种情况:
    --    1. 转出过程直接销帐的,则缴款数据不增加;
    --    2. 先转出,再到门诊退款退票,则需要进行缴款数据处理
    If Nvl(门诊退费_In, 0) = 1 Then
      For c_预交 In (Select a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, Min(a.交易流水号) As 交易流水号,
                          Min(a.交易说明) As 交易说明, Min(a.合作单位) As 合作单位, b.性质
                   From 病人预交记录 A, 结算方式 B
                   Where a.记录性质 = 3 And a.记录状态 In (2, 3) And
                         a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And a.结算方式 = b.名称 And
                         b.性质 In (1, 2, 3, 4, 7, 8) And a.结算方式 Is Not Null
                   Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质
                   Having Sum(a.冲预交) <> 0) Loop
        Begin
          Select 是否退现 Into n_退现 From 医疗卡类别 Where ID = c_预交.卡类别id;
        Exception
          When Others Then
            n_退现 := 0;
        End;
        If (c_预交.性质 = 7 Or (c_预交.性质 = 8 And c_预交.卡类别id Is Not Null)) And n_退现 = 0 Then
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|'
          Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
               缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
            Values
              (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null, 退费时间_In,
               Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|', n_组id,
               Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
          End If;
          n_费用状态 := 1;
        Else
          If c_预交.性质 In (3, 4) Or (c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null) Then
            v_结算方式 := c_预交.结算方式;
          Else
            If 结算方式_In Is Null Then
              Begin
                Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
              Exception
                When Others Then
                  Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
              End;
            Else
              v_结算方式 := 结算方式_In;
            End If;
          End If;
        
          If c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null Then
            --Zl_Square_Update(v_结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, c_预交.冲预交, c_预交.结算卡序号);
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|'
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|', n_组id,
                 Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
            End If;
            n_费用状态 := 1;
          End If;
          If c_预交.结算卡序号 Is Null Then
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) - c_预交.冲预交
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
            Returning 余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 人员缴款余额
                (收款员, 结算方式, 性质, 余额)
              Values
                (操作员姓名_In, v_结算方式, 1, -1 * c_预交.冲预交);
              n_返回值 := c_预交.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 人员缴款余额
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
            End If;
            --部分退费情况，退原预交记录
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交)
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, c_预交.合作单位, n_结帐id,
                 -1 * n_结帐id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    
      --更新费用审核记录
      Update 费用审核记录
      Set 记录状态 = 2
      Where 费用id In (Select a.Id
                     From 门诊费用记录 A
                     Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                           a.记录状态 In (1, 3)) And 性质 = 1;
      --作废门诊记录
      For r_Nos In (Select Distinct NO
                    From 门诊费用记录
                    Where Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And
                          结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))) Loop
        Update 门诊费用记录 Set 记录状态 = 3 Where NO = r_Nos.No And Mod(记录性质, 10) = 1 And 记录状态 = 1;
      End Loop;
      For r_Clinic In (Select Min(a.记录性质) As 记录性质, a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.姓名, a.性别, a.年龄, a.病人科室id, a.费别,
                              a.收费类别, a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, a.保险编码, a.费用类型, a.发药窗口, a.付数, Sum(a.数次) As 数次,
                              a.加班标志, a.附加标志, a.收入项目id, a.收据费目, a.标准单价, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额,
                              Sum(a.统筹金额) As 统筹金额, a.开单部门id, a.开单人, a.执行部门id, a.划价人, Max(a.记帐单id) As 记帐单id,
                              Max(a.是否急诊) As 是否急诊, a.发生时间, Min(a.实际票号) As 实际票号
                       From 门诊费用记录 A
                       Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                             a.记录状态 In (2, 3) And Nvl(a.附加标志, 0) Not In (8, 9)
                       Group By a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.姓名, a.性别, a.年龄, a.病人科室id, a.费别, a.收费类别, a.收费细目id,
                                a.计算单位, a.保险项目否, a.保险大类id, a.保险编码, a.费用类型, a.发药窗口, a.付数, a.加班标志, a.附加标志, a.收入项目id, a.收据费目,
                                a.标准单价, a.开单部门id, a.开单人, a.执行部门id, a.划价人, a.发生时间
                       Having Sum(a.数次) <> 0) Loop
        Insert Into 门诊费用记录
          (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
           保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间,
           登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id, 结帐id, 结帐金额, 执行状态, 费用状态)
        Values
          (病人费用记录_Id.Nextval, r_Clinic.记录性质, r_Clinic.No, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号,
           1, r_Clinic.病人id, '', r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别,
           r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口,
           r_Clinic.付数, -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价,
           -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 0, r_Clinic.开单部门id, r_Clinic.开单人, r_Clinic.发生时间,
           退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', r_Clinic.是否急诊, n_组id, n_结帐id,
           -1 * r_Clinic.实收金额, -1, 0);
      End Loop;
    Else
      --4.退款转预交(不产生票据,由操作员通过重打进行)
    
      For r_Pay In (Select Min(a.Id) As 预交id, a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号,
                           a.交易说明, a.合作单位, b.性质
                    From 病人预交记录 A, 结算方式 B
                    Where a.记录性质 = 3 And a.记录状态 In (2, 3) And
                          a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And a.结算方式 = b.名称 And
                          b.性质 In (1, 2, 3, 4, 7, 8) And a.结算方式 Is Not Null
                    Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质, a.交易流水号, a.交易说明, a.合作单位


                    
                    Having Sum(a.冲预交) <> 0) Loop
        --4.1产生预交款单据 (不存在部分退费的情况)
        --所有单据,按规则生成预交款单据
        --因为收款后立即缴款,所以人员缴款余额无变化
        If r_Pay.性质 = 7 Or (r_Pay.性质 = 8 And r_Pay.卡类别id Is Not Null) Then
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|'
          Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
               缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
            Values
              (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null, 退费时间_In,
               Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|', n_组id,
               Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
          End If;
          n_费用状态 := 1;
        Else
          If r_Pay.性质 In (3, 4) Or (r_Pay.性质 = 8 And r_Pay.结算卡序号 Is Not Null) Then
            v_结算方式 := r_Pay.结算方式;
          Else
            Begin
              Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
            Exception
              When Others Then
                Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
            End;
          End If;
        
          If r_Pay.性质 = 8 Then
            --Zl_Square_Update(v_结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, r_Pay.冲预交, r_Pay.结算卡序号);
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|'
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|', n_组id,
                 Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
            End If;
            n_费用状态 := 1;
          End If;
          If r_Pay.性质 Not In (3, 4, 7, 8) Then
            Update 病人预交记录
            Set 金额 = 金额 + r_Pay.冲预交
            Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              v_预交no := Nextno(11);
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 预交类别)
              Values
                (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, r_Pay.冲预交, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, r_Pay.预交类别);
            End If;
          
            --病人余额
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + r_Pay.冲预交
            Where 性质 = 1 And 病人id = n_病人id And 类型 = 2
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, 2, r_Pay.冲预交, 0);
              n_返回值 := r_Pay.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End If;
          --4.2缴款数据处理
          --   因为没有实际收病人的钱,所以不处理
          --部分退费情况，退原预交记录
          If r_Pay.性质 In (3, 4) Then
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) - r_Pay.冲预交
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式
            Returning 余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 人员缴款余额
                (收款员, 结算方式, 性质, 余额)
              Values
                (操作员姓名_In, r_Pay.结算方式, 1, -1 * r_Pay.冲预交);
              n_返回值 := r_Pay.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 人员缴款余额
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式 And Nvl(余额, 0) = 0;
            End If;
          End If;
        
          If r_Pay.结算卡序号 Is Null Then
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交)
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, r_Pay.卡类别id, r_Pay.结算卡序号, r_Pay.卡号, r_Pay.交易流水号,
                 r_Pay.交易说明, r_Pay.合作单位, n_结帐id, -1 * n_结帐id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    End If;
    If 误差费_In Is Not Null Then
      Begin
        Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
      Exception
        When Others Then
          Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
      End;
      Update 病人预交记录
      Set 冲预交 = 冲预交 - 误差费_In
      Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
      Update 病人预交记录
      Set 冲预交 = 冲预交 + 误差费_In
      Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_误差费;
      If Sql%RowCount = 0 Then
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
           缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
        Values
          (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, 误差费_In, v_误差费, Null, 退费时间_In, Null, Null,
           Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 0, 3);
      End If;
    End If;
    Delete From 病人预交记录 Where 结帐id = n_原结帐id And 摘要 = '预交临时记录' And 记录性质 = 3;
    Delete From 病人预交记录
    Where 结帐id = n_结帐id And 记录性质 = 3 And 记录状态 = 2 And 冲预交 = 0 And 结算方式 Is Not Null;
    Update 门诊费用记录
    Set 费用状态 = Nvl(n_费用状态, 0)
    Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(记录性质, 10) = 1 And 记录状态 = 2;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_收费转出;
/

--91427:马政,2015-12-22,新增药品验收系统
Create Or Replace Procedure Zl_药品验收记录_Insert
(
  Id_In         In 药品验收记录.Id%Type,
  No_In         In 药品验收记录.No%Type,
  库房id_In     In 药品验收记录.库房id%Type,
  供药单位id_In In 药品验收记录.供药单位id%Type,
  验收人_In     In 药品验收记录.验收人%Type,
  验收日期_In   In 药品验收记录.验收日期%Type,
  是否合格_In   In 药品验收记录.是否合格%Type :=0,
  备注_in     in 药品验收记录.备注%type :=null
) Is
Begin
  Insert Into 药品验收记录
    (ID, NO, 库房id, 供药单位id, 验收人, 验收日期,  是否合格,备注)
  Values
    (Id_In, No_In, 库房id_In, 供药单位id_In, 验收人_In, 验收日期_In, 是否合格_In,备注_in);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--91427:马政,2015-12-22,新增药品验收系统
Create Or Replace Procedure Zl_药品验收明细_Insert
(
  验收id_In   In 药品验收明细.验收id%Type,
  药品id_In   In 药品验收明细.药品id%Type,
  成本价_In   In 药品验收明细.成本价%Type :=null,
  零售价_In   In 药品验收明细.零售价%Type :=null,
  进药数量_In In 药品验收明细.进药数量%Type:=null,
  批号_In     In 药品验收明细.批号%Type:=null,
  生产日期_In In 药品验收明细.生产日期%Type:=null,
  效期_In     In 药品验收明细.效期%Type:=null,
  产地_In     In 药品验收明细.产地%Type:=null,
  批准文号_In In 药品验收明细.批准文号%Type:=null,
  进药日期_In In 药品验收明细.进药日期%Type:=null,
  是否合格_In In 药品验收明细.是否合格%Type:=0
) Is
Begin
  Insert Into 药品验收明细
    (验收id, 药品id, 成本价, 零售价, 进药数量, 批号, 生产日期, 效期, 产地, 批准文号, 进药日期, 是否合格)
  Values
    (验收id_In, 药品id_In, 成本价_In, 零售价_In, 进药数量_In, 批号_In, 生产日期_In, 效期_In, 产地_In, 批准文号_In, 进药日期_In, 是否合格_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--91427:马政,2015-12-22,新增药品验收系统
Create Or Replace Procedure Zl_药品验收记录_Delete(验收id_In In 药品验收记录.Id%Type) Is
  Err_Isverified Exception;
Begin
  Delete From 药品验收记录 Where ID = 验收id_In And 复核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--91427:马政,2015-12-22,新增药品验收系统
Create Or Replace Procedure Zl_药品验收记录_Verify
(
  验收id_In   In 药品验收记录.Id%Type,
  复核人_In   In 药品验收记录.复核人%Type,
  复核日期_In In 药品验收记录.复核日期%Type
) Is
  Err_Isverified Exception;
Begin
  Update 药品验收记录 Set 复核人 = 复核人_In, 复核日期 = 复核日期_In Where ID = 验收id_In And 复核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--92705:梁经伙,2016-01-15,传染病工作站回退已上报的报告，应该会退到待上报状态
--91225:梁经伙,2016-01-11,接收和审核一起进行后，回退时回退接收和审核
CREATE OR REPLACE Procedure Zl_疾病申报记录_Untread
(
  文件id_In   In Varchar2,
  IsStation_In   in Number:=NULL      --是否是传染病工作站调用，0，不是，1是传染病工作站调用
) Is
  n_处理状态 疾病申报记录.处理状态%Type;
  n_文件id   Number;
  n_Count    Number;
Begin
  If Length(文件id_In) <> 32 Then
    n_文件id := To_Number(文件id_In); --新病历ID是32位GUID
  End If;

  Select count(1)
  Into n_Count
  From 疾病申报记录
  Where 撤档人 Is Not Null And 撤档时间 Is Not Null And
        Decode(Length(文件id_In), 32, 文档id, 文件id) = Decode(Length(文件id_In), 32, 文件id_In, n_文件id);

  If n_Count > 0 Then      --取消删除
    Update 疾病申报记录
    Set 撤档人 = Null, 撤档时间 = Null
    Where 撤档人 Is Not Null And 撤档时间 Is Not Null And
          Decode(Length(文件id_In), 32, 文档id, 文件id) = Decode(Length(文件id_In), 32, 文件id_In, n_文件id);
  Else
    Select 处理状态
    Into n_处理状态
    From 疾病申报记录
    Where Decode(Length(文件id_In), 32, 文档id, 文件id) = Decode(Length(文件id_In), 32, 文件id_In, n_文件id);
    If n_处理状态 = 2 Then         --取消上报
      --如果是申报登记时进行的归档（归档人和申报登记人是否相同），则取消归档
      If Length(文件id_In) <> 32 Then
        Update 电子病历记录
        Set 归档人 = Null, 归档日期 = Null
        Where ID = n_文件id And 归档人 = (Select 登记人 From 疾病申报记录 Where 文件id = n_文件id);
      End If;
      if IsStation_In =1 then
        Update 疾病申报记录
        Set 处理状态 = 3, 报送人 = '', 报送时间 = Null, 报送单位 = Null, 报送备注 = '', 登记人 = '', 登记时间 = ''
        Where Decode(Length(文件id_In), 32, 文档id, 文件id) = Decode(Length(文件id_In), 32, 文件id_In, n_文件id);
      else
        Update 疾病申报记录
        Set 处理状态 = 1, 报送人 = '', 报送时间 = Null, 报送单位 = Null, 报送备注 = '', 登记人 = '', 登记时间 = ''
        Where Decode(Length(文件id_In), 32, 文档id, 文件id) = Decode(Length(文件id_In), 32, 文件id_In, n_文件id);
      end if;
    Elsif n_处理状态 = 1 Or n_处理状态 = -1 Then   --取消接收和拒绝
      If Length(文件id_In) <> 32 Then
        Update 电子病历记录 Set 处理状态 = 0 Where ID = n_文件id;
      End If;
      Delete 疾病申报记录
      Where Decode(Length(文件id_In), 32, 文档id, 文件id) = Decode(Length(文件id_In), 32, 文件id_In, n_文件id);
    Elsif n_处理状态 = 3 Or n_处理状态 = 4 Then  --取消审核
      If Length(文件id_In) <> 32 Then
        Update 电子病历记录 Set 处理状态 = 0 Where ID = n_文件id;
		Delete 疾病报告反馈 Where 文件id = n_文件id And 登记时间 = (Select Max(登记时间) From 疾病报告反馈 Where 文件id = n_文件id);
      End If;
      
      Update 疾病申报记录 Set 处理状态 = 1
      Where Decode(Length(文件id_In), 32, 文档id, 文件id) = Decode(Length(文件id_In), 32, 文件id_In, n_文件id);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病申报记录_Untread;
/

--91225:梁经伙,2015-12-22,传染病管理系统新增过程
CREATE OR REPLACE Procedure Zl_疾病申报记录_Delete(Id_In In Varchar2) Is
  v_撤档人 疾病申报记录.撤档人%Type;
  n_文件id Number;
  e_Changed Exception;
Begin

  If Length(Id_In) <> 32 Then
    n_文件id := To_Number(Id_In); --新病历ID是32位GUID 
  End If;
  Select b.姓名 Into v_撤档人 From 上机人员表 A, 人员表 B Where a.人员id = b.Id And a.用户名 = User And Rownum < 2;

  Update 疾病申报记录
  Set 撤档人 = v_撤档人, 撤档时间 = Sysdate
  Where Decode(Length(Id_In), 32, 文档id, 文件id) = Decode(Length(Id_In), 32, Id_In, n_文件id);
  If Sql%RowCount = 0 Then
    Raise e_Changed;
  End If;

Exception
  When e_Changed Then
    Raise_Application_Error(-20101, '[ZLSOFT]疾病报告已经被其他用户改变！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病申报记录_Delete;
/

--91225:梁经伙,2015-12-24,传染病管理系统
CREATE OR REPLACE Procedure Zl_疾病申报记录_Update
(
  文件id_In     In 疾病申报记录.文件id%Type,
  Aduitstate_In In Number,
  登记时间_In   In 疾病报告反馈.登记时间%Type,
  登记人_In     In 疾病报告反馈.登记人%Type,
  反馈内容_In   In 疾病报告反馈.反馈内容%Type,
  处理人_In     In 疾病报告反馈.处理人%Type,
  处理时间_In   In 疾病报告反馈.处理时间%Type,
  处理内容_In   In 疾病报告反馈.处理情况说明%Type
) Is
Begin
  If Aduitstate_In = 3 Then
    Update 疾病申报记录 Set 处理状态 = Aduitstate_In Where 文件id = 文件id_In;
    Insert Into 疾病报告反馈
      (文件id, 登记时间, 登记人, 记录状态, 反馈内容)
    Values
      (文件id_In, 登记时间_In, 登记人_In, 3, 反馈内容_In);
  Elsif Aduitstate_In = 4 Then
    Update 疾病申报记录 Set 处理状态 = Aduitstate_In Where 文件id = 文件id_In;
    Insert Into 疾病报告反馈
      (文件id, 登记时间, 登记人, 记录状态, 反馈内容)
    Values
      (文件id_In, 登记时间_In, 登记人_In, 1, 反馈内容_In);
  Elsif Aduitstate_In = 5 Then
      Update 疾病申报记录 Set 处理状态 = Aduitstate_In, 报卡类型='2 订正报告' Where 文件id = 文件id_In;
    
      Update 疾病报告反馈
      Set 记录状态 = 2, 处理人 = 处理人_In, 处理时间 = 处理时间_In, 处理情况说明 = 处理内容_In
      Where 文件id = 文件id_In And 登记时间 = (Select Max(登记时间) From 疾病报告反馈 Where 文件id = 文件id_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病申报记录_Update;
/

--91225:梁经伙,2015-12-18,阳性检测结果插入和更新的过程
CREATE OR REPLACE Procedure Zl_疾病阳性检测记录_Insert
(
  Id_In         In 疾病阳性记录.Id%Type,
  病人id_In     In 疾病阳性记录.病人id%Type,
  主页id_In     In 疾病阳性记录.主页id%Type,
  挂号单_In     In 疾病阳性记录.挂号单%Type,
  送检时间_In   In 疾病阳性记录.送检时间%Type,
  送检科室id_In In 疾病阳性记录.送检科室id%Type,
  送检医生_In   In 疾病阳性记录.送检医生%Type,
  标本名称_In   In 疾病阳性记录.标本名称%Type,
  反馈结果_In   In 疾病阳性记录.反馈结果%Type,
  传染病_In     In 疾病阳性记录.传染病名称%Type,
  检查时间_In   In 疾病阳性记录.检查时间%Type,
  登记时间_In   In 疾病阳性记录.登记时间%Type,
  登记人_In     In 疾病阳性记录.登记人%Type,
  登记科室id_In In 疾病阳性记录.登记科室id%Type,
  记录状态_In   In 疾病阳性记录.记录状态%Type
) Is
Begin
  Insert Into 疾病阳性记录
    (ID, 病人id, 主页id, 挂号单, 送检时间, 送检科室id, 送检医生, 标本名称, 反馈结果, 传染病名称, 检查时间, 登记时间, 登记人, 登记科室id, 记录状态)
  Values
    (Id_In, 病人id_In, 主页id_In, 挂号单_In, 送检时间_In, 送检科室id_In, 送检医生_In, 标本名称_In, 反馈结果_In, 传染病_In, 检查时间_In, 登记时间_In, 登记人_In,
     登记科室id_In, 记录状态_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病阳性检测记录_Insert;
/

--91225:梁经伙,2016-01-11,阳性检测结果，关联和取消关联反馈单
CREATE OR REPLACE Procedure Zl_疾病阳性检测记录_Update
(
  Operate_in         In  number,
  Id_In           In 疾病阳性记录.Id%Type,
  文件ID_In       In 疾病阳性记录.文件ID%Type,
  记录状态_In     In 疾病阳性记录.记录状态%Type,
  处理人_In       In 疾病阳性记录.处理人%Type,
  处理时间_In     In 疾病阳性记录.处理时间%Type,
  处理情况说明_In In 疾病阳性记录.处理情况说明%Type
) Is
Begin
  if Operate_in = 1 then      /*设置处理说明 */
      Update 疾病阳性记录
      Set 处理人 = 处理人_In, 处理时间 = 处理时间_In, 处理情况说明 = 处理情况说明_In,记录状态 = 记录状态_In,文件ID = 文件ID_In
      Where ID = Id_In;
  elsif Operate_in = 2 then   /*关联报告单和阳性结果反馈单*/
    if  文件ID_In is not null then
        Update 疾病阳性记录 Set 文件ID = 文件ID_In Where ID = Id_In;
    end if;
  elsif Operate_in = 3 then   /*取消报告单和阳性结果反馈单的关联*/
    if  文件ID_In is not null then
        Update 疾病阳性记录 Set 文件ID = NULL Where 文件ID = 文件ID_In;
    end if;
  end if;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病阳性检测记录_Update;
/

--91709:冉俊明,2015-12-17,异常收费单据作废未产生病人预交记录。
Create Or Replace Procedure Zl_门诊退费结算_Modify
(
  操作类型_In     Number,
  病人id_In       门诊费用记录.病人id%Type,
  冲销id_In       病人预交记录.结帐id%Type,
  结算方式_In     Varchar2,
  冲预交_In       病人预交记录.冲预交%Type := Null,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  缴款_In         病人预交记录.缴款%Type := Null,
  找补_In         病人预交记录.找补%Type := Null,
  误差金额_In     门诊费用记录.实收金额%Type := Null,
  完成退费_In     Number := 0,
  原结帐id_In     病人预交记录.结帐id%Type := Null,
  剩余转预交_In   Number := 0,
  缺省结算方式_In 结算方式.名称%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:收费结算时,修改结算的相关信息
  --操作类型_In:
  --   0-原样退
  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
  --   1-普通退费方式:
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
  --   2.三方卡退费结算:
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
  --     ②退支票额_In:传入零
  --   4-消费卡结算:
  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
  --     ②退支票额_In:传入零

  -- 冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款
  -- 剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交
  -- 误差金额_In:存在误差费时,传入
  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费
  -- 原结帐ID_IN:原样退时,传入(如果原样退未传入时,则以最后一次结帐为准)
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_结算内容 Varchar2(500);
  v_当前结算 Varchar2(50);
  v_卡号     病人医疗卡信息.卡号%Type;
  n_消费卡id 消费卡目录.Id%Type;
  n_卡类别id 病人预交记录.结算卡序号%Type;
  v_名称     Varchar2(100);
  n_自制卡   卡消费接口目录.自制卡%Type;
  n_序号     病人卡结算记录.序号%Type;
  n_Id       病人卡结算记录.Id%Type;
  n_预交id   病人预交记录.Id%Type;
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  n_返回值   人员缴款余额.余额%Type;
  n_预交金额 病人预交记录.冲预交%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;
  v_误差费   结算方式.名称%Type;

  v_退费结算 结算方式.名称%Type;
  v_No       病人预交记录.No%Type;
  n_Dec      Number; --金额小数位数 

  n_Count    Number;
  n_Havenull Number;
  l_预交id   t_Numlist := t_Numlist();
  n_原结帐id 病人预交记录.结帐id%Type;
  n_重结id   病人预交记录.结帐id%Type;
  n_结帐id   病人预交记录.结帐id%Type;
  n_结算序号 病人预交记录.结帐id%Type;
  v_Msg      Varchar2(5000);
  Cursor c_Feedata Is
    Select Max(NO) As NO, Max(m.病人id) As 病人id, Max(m.登记时间) As 登记时间, Max(m.操作员编号) As 操作员编号, Max(m.操作员姓名) As 操作员姓名,
           Sum(结帐金额) As 结算金额, Max(m.缴款组id) As 缴款组id
    From 门诊费用记录 M
    Where m.结帐id = 冲销id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号
    From 病人预交记录
    Where 结帐id = 冲销id_In And 结算方式 Is Null;
  r_Balancedata c_Balancedata%RowType;

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
                 Where a.Id = b.预交id And b.卡结算id = c.Id And a.记录性质 = 3 And a.记录状态 In (1, 3) And
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
          Select n_预交id, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退款时间_In, 缴款单位, 单位开户行, 单位帐号, r_Balancedata.操作员编号,
                 r_Balancedata.操作员姓名, -1 * 冲预交, 现结帐id_In, 缴款组id_In, 预交类别, 卡类别id, Nvl(结算卡序号, v_校对.接口编号), 卡号, 交易流水号, 交易说明,
                 合作单位, 2, 结算序号_In, Mod(记录性质, 10)
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

  Begin
    Select 名称 Into v_退费结算 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_退费结算 := '现金';
  End;

  Open c_Feedata;
  Fetch c_Feedata
    Into r_Feedata;

  If r_Feedata.No Is Null Then
    v_Err_Msg := '未找到指定的退费记录！';
    Raise Err_Item;
  End If;

  --金额小数位数 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --0.正式结算
  Select Count(1), Max(Decode(结算方式, Null, 1, 0)), Max(结算序号)
  Into n_Count, n_Havenull, n_结算序号
  From 病人预交记录
  Where 结帐id = 冲销id_In;

  If Nvl(n_Count, 0) = 0 Or Nvl(误差金额_In, 0) <> 0 Then
    --增加结算方式为NULL的记录
    Begin
      Select 名称 Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
    Exception
      When Others Then
        v_误差费 := '误差费';
    End;
  End If;

  --1.增加结算方式为空的结算数据
  If Nvl(n_Havenull, 0) = 0 Then
    n_Count := 0;
    Begin
      n_结算金额 := Round(Nvl(r_Feedata.结算金额, 0), n_Dec);
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 2, Decode(病人id_In, 0, Null, 病人id_In), Null, r_Feedata.登记时间, r_Feedata.操作员编号,
         r_Feedata.操作员姓名, n_结算金额, 冲销id_In, r_Feedata.缴款组id, -1 * 冲销id_In, 1, 3);
      --误差费(先汇总后生成误差费
      If n_结算金额 <> Nvl(r_Feedata.结算金额, 0) Then
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 1, Decode(病人id_In, 0, Null, 病人id_In), v_误差费, r_Feedata.登记时间, r_Feedata.操作员编号,
           r_Feedata.操作员姓名, Nvl(r_Feedata.结算金额, 0) - n_结算金额, 冲销id_In, r_Feedata.缴款组id, -1 * 冲销id_In, 1, 3);
      End If;
      n_结算序号 := -1 * 冲销id_In;
    Exception
      When Others Then
        n_Count := 1;
    End;
    If n_Count = 1 Then
      v_Err_Msg := '未找到指定的退费明细数据,结算操作失败！';
      Raise Err_Item;
    End If;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If 操作类型_In = 0 Then
    --0.原样退
    n_原结帐id := 原结帐id_In;
    If Nvl(n_原结帐id, 0) = 0 Then
      Select Max(结帐id)
      Into n_原结帐id
      From 门诊费用记录 A,
           (Select 登记时间 From 门诊费用记录 Where NO = r_Feedata.No And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3)) B
      Where a.No = r_Feedata.No And Mod(a.记录性质, 10) = 1 And a.登记时间 = b.登记时间;
    End If;
    If Nvl(n_原结帐id, 0) = 0 Then
      v_Err_Msg := '未找到原结帐数据,不能原样退！';
      Raise Err_Item;
    End If;
  
    --1.先处理预交款
    n_结算金额 := 0;
    For v_退预交 In (Select a.Id, Nvl(a.冲预交, 0) As 金额
                  From 病人预交记录 A
                  Where Mod(记录性质, 10) = 1 And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0
                  Order By 收款时间 Desc) Loop
    
      n_结算金额 := n_结算金额 + v_退预交.金额;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号,
         卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质)
        Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, Null, 摘要, 结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
               r_Balancedata.操作员姓名, -1 * v_退预交.金额, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null,
               Null, 卡号, 交易流水号, 交易说明, Null, 预交类别, 3
        From 病人预交记录
        Where ID = v_退预交.Id;
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - v_退预交.金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    End Loop;
    If Nvl(n_结算金额, 0) <> 0 Then
    
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_结算金额)
      Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (病人id_In, 1, (-1 * n_结算金额), 1);
        n_返回值 := (-1 * n_结算金额);
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Balancedata.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    End If;
    --2.处理消费卡部分
    Zl_Square_Update(n_原结帐id, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.收款时间, r_Balancedata.结算序号, v_结算内容);
    --3.处理其他结算部分
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号,
       交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, r_Balancedata.收款时间, r_Balancedata.操作员编号,
             r_Balancedata.操作员姓名, -1 * 冲预交, r_Balancedata.结帐id, r_Balancedata.缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位,
             Case
               When Nvl(a.卡类别id, 0) <> 0 Then
                1
               When Nvl(a.结算卡序号, 0) <> 0 Then
                1
               When Nvl(q.预交id, 0) <> 0 Then
                1
               When Nvl(j.名称, '-') <> '-' Then
               --医保
                1
               Else
                2
             End As 校对标志, r_Balancedata.结算序号, 3
      From 病人预交记录 A, (Select 名称 From 结算方式 Where 性质 In (3, 4)) J,
           (Select m.Id As 预交id
             From 病人预交记录 M, 一卡通目录 C
             Where m.结帐id = n_原结帐id And m.结算方式 = c.结算方式 And m.记录性质 = 3 And m.记录状态 In (1, 3)) Q
      Where Mod(a.记录性质, 10) <> 1 And a.记录状态 In (1, 3) And a.结算方式 = j.名称(+) And a.结算方式 Is Not Null And a.结帐id = n_原结帐id And
            a.Id = q.预交id(+) And (Not Exists (Select 1 From 病人卡结算对照 Where a.Id = 预交id) Or Nvl(结算卡序号, 0) = 0);
  
    --更新结算方式为NULL 的记录
    Select Sum(冲预交) Into n_返回值 From 病人预交记录 Where 结帐id = r_Balancedata.结帐id And 结算方式 Is Not Null;
    Select Sum(结帐金额)
    Into n_结算金额
    From 门诊费用记录
    Where 结帐id = r_Balancedata.结帐id And Mod(记录性质, 10) = 1;
  
    Update 病人预交记录
    Set 冲预交 = Nvl(n_结算金额, 0) - Nvl(n_返回值, 0)
    Where 结帐id = 冲销id_In And 结算方式 Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
      Raise Err_Item;
    End If;
  
  End If;

  n_重结id := 0;
  If 操作类型_In <> 0 Then
    --不是全退时,检查是否产生了重新收费数据的
    Begin
      Select 结帐id Into n_重结id From 病人预交记录 Where 结算序号 = n_结算序号 And 结帐id <> 冲销id_In And Rownum < 2;
    Exception
      When Others Then
        n_重结id := 0;
    End;
  End If;

  --需要处理误差金额
  If Nvl(误差金额_In, 0) <> 0 Then
    --误差费放在重收的结算记录中
    n_结帐id := 冲销id_In;
    If Nvl(n_重结id, 0) <> 0 Then
      n_结帐id := n_重结id;
    End If;
  
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + Nvl(误差金额_In, 0)
    Where 结帐id = n_结帐id And 结算方式 = v_误差费;
    If Sql%NotFound Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, Null, v_误差费, r_Balancedata.收款时间, r_Balancedata.操作员编号,
         r_Balancedata.操作员姓名, 误差金额_In, n_结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In,
         交易说明_In, Null, 3);
    End If;
  
    Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(误差金额_In, 0) Where 结帐id = n_结帐id And 结算方式 Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
      Raise Err_Item;
    End If;
  End If;

  --预交款处理:如果是冲预交,需要先处理冲预交款
  If Nvl(冲预交_In, 0) <> 0 Then
    If Nvl(r_Balancedata.病人id, 0) = 0 Then
      v_Err_Msg := '不能确定病人信息,不能使用预交款结算！';
      Raise Err_Item;
    End If;
  
    n_预交金额 := 冲预交_In;
    If n_预交金额 < 0 And Nvl(剩余转预交_In, 0) = 1 Then
    
      --     ②冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款
      --     ③剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交
    
      --1.先生成冲值预交:
      v_No := Nextno(11);
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 金额, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 预交类别, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 1, v_No, 1, r_Balancedata.病人id, Null, '退费生成预交', v_退费结算, r_Balancedata.收款时间,
         r_Balancedata.操作员编号, r_Balancedata.操作员姓名, -1 * n_预交金额, r_Balancedata.结帐id, r_Balancedata.缴款组id,
         r_Balancedata.结算序号, 0, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 1, Null);
    
      --更新病人余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + n_预交金额
      Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (病人id_In, 1, n_预交金额, 1);
        n_返回值 := n_预交金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Balancedata.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --2.生成退费记录
      If Nvl(n_重结id, 0) <> 0 Then
        Select Sum(冲预交)
        Into n_返回值
        From 病人预交记录
        Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
        If Nvl(n_返回值, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 结算号码, 结算性质)
          Values
            (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_退费结算, r_Balancedata.收款时间,
             r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_返回值, 冲销id_In, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null,
             Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3);
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_返回值, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
        End If;
        n_结算金额 := n_结算金额 - Nvl(n_返回值, 0);
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
           交易流水号, 交易说明, 结算号码, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_退费结算, r_Balancedata.收款时间,
           r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, n_重结id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null,
           Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3);
        Update 病人预交记录 Set 冲预交 = 冲预交 + n_预交金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
          Raise Err_Item;
        End If;
      Else
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
           交易流水号, 交易说明, 结算号码, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
           r_Balancedata.操作员编号, r_Balancedata.操作员姓名, -1 * n_预交金额, r_Balancedata.结帐id, r_Balancedata.缴款组id,
           r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3);
        Update 病人预交记录 Set 冲预交 = 冲预交 + n_预交金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    If Nvl(n_预交金额, 0) < 0 And Nvl(剩余转预交_In, 0) = 0 Then
    
      n_原结帐id := 原结帐id_In;
      If Nvl(n_原结帐id, 0) = 0 Then
        Select Max(b.结帐id)
        Into n_原结帐id
        From 门诊费用记录 A, 门诊费用记录 B
        Where a.结帐id = 冲销id_In And a.No = b.No And b.记录性质 = 1 And b.记录状态 In (1, 3);
      End If;
    
      If Nvl(n_原结帐id, 0) = 0 Then
        v_Err_Msg := '未找到原结帐数据,不能原样退！';
        Raise Err_Item;
      End If;
      If Nvl(n_重结id, 0) <> 0 Then
        Select Sum(冲预交)
        Into n_返回值
        From 病人预交记录
        Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
        If Nvl(n_返回值, 0) <> 0 Then
          For v_退预交 In (Select a.Id, Nvl(a.冲预交, 0) As 金额
                        From 病人预交记录 A
                        Where Mod(记录性质, 10) = 1 And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0
                        Order By 收款时间 Desc) Loop
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质)
              Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, Null, 摘要, 结算方式, r_Balancedata.收款时间,
                     r_Balancedata.操作员编号, r_Balancedata.操作员姓名, Nvl(n_返回值, 0), r_Balancedata.结帐id, r_Balancedata.缴款组id,
                     r_Balancedata.结算序号, 2, Null, Null, 卡号, 交易流水号, 交易说明, Null, 预交类别, 3
              From 病人预交记录
              Where ID = v_退预交.Id;
            Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_返回值, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
            If Sql%NotFound Then
              v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
              Raise Err_Item;
            End If;
            n_预交金额 := n_预交金额 - Nvl(n_返回值, 0);
          End Loop;
        End If;
        n_返回值 := 0;
        --2.退预交款
        For v_退预交 In (Select Max(a.Id) As ID, Max(a.收款时间) As 收款时间, Sum(Nvl(a.冲预交, 0)) As 金额
                      From 病人预交记录 A,
                           (Select Distinct a.结帐id
                             From 门诊费用记录 A, 门诊费用记录 B
                             Where a.No = b.No And Mod(b.记录性质, 10) = 1 And b.结帐id = n_原结帐id) B
                      Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) = 1 And Nvl(a.预交类别, 0) = 1 And a.结帐id <> 冲销id_In
                      Group By NO
                      Order By 收款时间 Desc) Loop
        
          If v_退预交.金额 + n_预交金额 < 0 Then
            n_结算金额 := v_退预交.金额;
            n_预交金额 := n_预交金额 + v_退预交.金额;
          Else
            n_结算金额 := n_预交金额;
            n_预交金额 := 0;
          End If;
        
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质)
            Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, Null, 摘要, 结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
                   r_Balancedata.操作员姓名, n_结算金额, n_重结id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null, Null, 卡号,
                   交易流水号, 交易说明, Null, 预交类别, 3
            From 病人预交记录
            Where ID = v_退预交.Id;
          Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = n_重结id And 结算方式 Is Null;
          n_返回值 := 1;
          If Sql%NotFound Then
            v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
            Raise Err_Item;
          End If;
          If n_预交金额 = 0 Then
            Exit;
          End If;
        End Loop;
      Else
        --退预交款
        n_返回值   := 0;
        n_预交金额 := -1 * n_预交金额;
      
        For v_退预交 In (Select Max(a.Id) As ID, Max(a.收款时间) As 收款时间, Sum(Nvl(a.冲预交, 0)) As 金额
                      From 病人预交记录 A,
                           (Select Distinct a.结帐id
                             From 门诊费用记录 A, 门诊费用记录 B
                             Where a.No = b.No And Mod(b.记录性质, 10) = 1 And b.结帐id = n_原结帐id) B
                      Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) = 1 And Nvl(a.预交类别, 0) = 1
                      Group By NO
                      Having Sum(Nvl(a.冲预交, 0)) > 0
                      Order By 收款时间 Desc) Loop
        
          If v_退预交.金额 - n_预交金额 < 0 Then
            n_结算金额 := -1 * v_退预交.金额;
            n_预交金额 := n_预交金额 - v_退预交.金额;
          Else
            n_结算金额 := -1 * n_预交金额;
            n_预交金额 := 0;
          End If;
        
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质)
            Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
                   r_Balancedata.操作员姓名, n_结算金额, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null,
                   Null, 卡号, 交易流水号, 交易说明, Null, 预交类别, 3
            From 病人预交记录
            Where ID = v_退预交.Id;
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
          n_返回值 := 1;
          If Sql%NotFound Then
            v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
            Raise Err_Item;
          End If;
          If n_预交金额 = 0 Then
            Exit;
          End If;
        End Loop;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        v_Err_Msg := '未找到原始的冲预交记录,不能回退预交款！';
        Raise Err_Item;
      End If;
    
      If Nvl(n_预交金额, 0) <> 0 Then
        v_Err_Msg := '当前退预交超过了收费结算中的冲预交款,不能回退预交款！';
        Raise Err_Item;
      End If;
    
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + (-1 * 冲预交_In)
      Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (病人id_In, 1, (-1 * 冲预交_In), 1);
        n_返回值 := (-1 * 冲预交_In);
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Balancedata.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
    End If;
  
    n_预交金额 := 冲预交_In;
    If Nvl(n_预交金额, 0) > 0 Then
      --冲预交款
      --病人余额检查
      Begin
        Select Nvl(预交余额, 0) - Nvl(费用余额, 0)
        Into n_预交金额
        From 病人余额
        Where 病人id = 病人id_In And Nvl(性质, 0) = 1 And 类型 = 1;
      Exception
        When Others Then
          n_预交金额 := 0;
      End;
      If n_预交金额 < 冲预交_In Then
        v_Err_Msg := '病人的当前预交余额为 ' || LTrim(To_Char(n_预交金额, '9999999990.00')) || '，小于本次支付金额 ' ||
                     LTrim(To_Char(冲预交_In, '9999999990.00')) || ' ！';
        Raise Err_Item;
      End If;
    
      n_预交金额 := 冲预交_In;
      n_结帐id   := 冲销id_In;
      If Nvl(n_重结id, 0) <> 0 Then
        n_结帐id := n_重结id;
        --总的冲预交金额 = 本次冲预交金额 + 未冲销金额
        --因为在后面会将未冲销金额全部退为预交款
        Select Sum(冲预交)
        Into n_返回值
        From 病人预交记录
        Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
        If Nvl(n_返回值, 0) <> 0 Then
          n_预交金额 := n_预交金额 - Nvl(n_返回值, 0);
        End If;
      End If;
    
      For c_冲预交 In (Select *
                    From (Select a.Id, a.记录状态, a.No, Nvl(a.金额, 0) As 金额
                           From 病人预交记录 A,
                                (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                                  From 病人预交记录 A
                                  Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.病人id = 病人id_In And 预交类别 = 1
                                  Group By NO
                                  Having Sum(Nvl(a.金额, 0)) <> 0) B
                           Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.No = b.No And a.病人id = 病人id_In And a.预交类别 = 1
                           Union All
                           Select 0 As ID, 记录状态, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
                           From 病人预交记录
                           Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And 病人id = 病人id_In And
                                 预交类别 = 1 Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
                           Group By 记录状态, NO)
                    Order By ID, NO) Loop
      
        If c_冲预交.金额 - n_预交金额 < 0 Then
          n_冲预交 := c_冲预交.金额;
        Else
          n_冲预交 := n_预交金额;
        End If;
      
        If c_冲预交.Id <> 0 Then
          --第一次冲预交(将第一次标上结帐ID,冲预交标记为0)
          Update 病人预交记录
          Set 冲预交 = 0, 结帐id = n_结帐id, 结算序号 = n_结算序号, 结算性质 = 3
          Where ID = c_冲预交.Id;
        End If;
        --冲上次剩余额
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 r_Balancedata.收款时间, r_Balancedata.操作员姓名, r_Balancedata.操作员编号, n_冲预交, n_结帐id, r_Balancedata.缴款组id, 预交类别,
                 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结算序号, 3
          From 病人预交记录
          Where NO = c_冲预交.No And 记录状态 = c_冲预交.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新数据(结算方式为NULL的)
        Update 病人预交记录
        Set 冲预交 = 冲预交 - n_冲预交
        Where 结帐id = n_结帐id And 结算方式 Is Null
        Returning Nvl(冲预交, 0) Into n_返回值;
      
        --检查是否已经处理完
        If c_冲预交.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - c_冲预交.金额;
        Else
          n_预交金额 := 0;
        End If;
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
    
      --检查金额是否足够
      If Abs(n_预交金额) > 0 Then
        v_Err_Msg := '病人的当前预交余额小于本次支付金额 ' || LTrim(To_Char(冲预交_In, '9999999990.00')) || ' ！';
        Raise Err_Item;
      End If;
    
      If Nvl(n_重结id, 0) <> 0 Then
        Select Sum(冲预交)
        Into n_返回值
        From 病人预交记录
        Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
        If Nvl(n_返回值, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
             结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                   r_Balancedata.收款时间, r_Balancedata.操作员姓名, r_Balancedata.操作员编号, n_返回值, 冲销id_In, r_Balancedata.缴款组id,
                   预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结算序号, 3
            From 病人预交记录
            Where 结帐id = n_重结id And 记录性质 In (1, 11) And Rownum = 1;
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_返回值, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
          If Sql%NotFound Then
            v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
            Raise Err_Item;
          End If;
        End If;
      
      End If;
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - 冲预交_In
      Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (病人id_In, 1, -1 * 冲预交_In, 1);
        n_返回值 := -1 * 冲预交_In;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    End If;
  End If;

  If 操作类型_In = 1 Then
    --   1-普通退费方式:
    --各个收费结算 :格式为:"结算方式|结算金额|结算号码|结算摘要||.."
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      --不判断“结算金额”是否为零，有可能已经退完，但这时结算方式为空的重结和冲销记录的冲预交之和为零
      If v_结算方式 Is Null Then
        v_结算方式 := 缺省结算方式_In;
      End If;
      --If Nvl(n_结算金额, 0) <> 0 Then
      n_结算金额 := Nvl(n_结算金额, 0);
      If Nvl(n_重结id, 0) <> 0 Then
        --肯定是收款
        --1.先按此种方式全退
        --2.再按此种方式收款
        --3:1+2=本次退款
        --1.先将退费的全部作废掉
        Select Sum(冲预交)
        Into n_返回值
        From 病人预交记录
        Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
      
        If Nvl(n_返回值, 0) <> 0 Then
        
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 结算号码, 结算性质)
          Values
            (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
             r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_返回值, 冲销id_In, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null,
             Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3);
          Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_返回值, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
        End If;
        n_结算金额 := n_结算金额 - Nvl(n_返回值, 0);
        --2.退款
        If Nvl(n_结算金额, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 结算号码, 结算性质)
          Values
            (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
             r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, n_重结id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null,
             Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3);
          Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = n_重结id And 结算方式 Is Null;
        End If;
      Else
        --:>退款
        If Nvl(n_结算金额, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 结算号码, 结算性质)
          Values
            (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
             r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, r_Balancedata.结帐id, r_Balancedata.缴款组id,
             r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3);
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
          If Sql%NotFound Then
            v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  
  End If;

  If 操作类型_In = 2 Then
    --   2.三方卡退费结算:
  
    v_当前结算 := 结算方式_In;
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    If Nvl(n_结算金额, 0) <> 0 Then
    
      If Nvl(n_重结id, 0) <> 0 Then
        --1.先按此种方式全退
        --2.再按此种方式收款
        --3:1+2=本次退款
        --1.先将退费的全部作废掉
        Select Sum(冲预交)
        Into n_返回值
        From 病人预交记录
        Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
        If Nvl(n_返回值, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 结算号码, 结算性质)
          Values
            (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
             r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_返回值, 冲销id_In, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2,
             卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3);
          Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_返回值, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
        End If;
        n_结算金额 := n_结算金额 - Nvl(n_返回值, 0);
      
        --2.退款
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
           交易流水号, 交易说明, 结算号码, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
           r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, n_重结id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2,
           卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3);
        Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = n_重结id And 结算方式 Is Null;
      
      Else
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
           交易流水号, 交易说明, 结算号码, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
           r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号,
           2, 卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 3);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  If 操作类型_In = 3 Then
    --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    --3.1检查是否已经存在医保结算数据,存在先删除
    n_结算金额 := 0;
    For v_医保 In (Select ID, 结算方式, 冲预交
                 From 病人预交记录 A
                 Where 结帐id = 冲销id_In And Exists
                  (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4)) And Mod(记录性质, 10) <> 1) Loop
      n_结算金额 := n_结算金额 + Nvl(v_医保.冲预交, 0);
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := v_医保.Id;
    End Loop;
    If Nvl(n_结算金额, 0) <> 0 Then
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 冲销id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    End If;
    If l_预交id.Count <> 0 Then
      Forall I In 1 .. l_预交id.Count
        Delete 病人预交记录 Where ID = l_预交id(I);
    End If;
  
    If 结算方式_In Is Not Null Then
      v_结算内容 := 结算方式_In || '||';
    End If;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      For c_结算信息 In (Select 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号
                     From 病人预交记录
                     Where 结帐id = 冲销id_In And 结算方式 Is Null) Loop
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 2, c_结算信息.病人id, Null, '保险结算', v_结算方式, c_结算信息.收款时间, c_结算信息.操作员编号, c_结算信息.操作员姓名,
           n_结算金额, c_结算信息.结帐id, c_结算信息.缴款组id, c_结算信息.结算序号, 1, 3);
      End Loop;
      --更新数据(结算方式为NULL的)
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_结算金额
      Where 结帐id = 冲销id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --4-消费卡批量结算
  If 操作类型_In = 4 Then
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      --卡类别ID|卡号|消费卡ID|消费金额
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(v_当前结算);
      Begin
        Select 名称, 自制卡, 结算方式 Into v_名称, n_自制卡, v_结算方式 From 卡消费接口目录 Where 编号 = n_卡类别id;
      Exception
        When Others Then
          v_名称 := Null;
      End;
      If v_名称 Is Null Then
        v_Err_Msg := '未找到对应的结算卡接口,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := v_名称 || '未设置对应的结算方式,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If Nvl(n_结算金额, 0) <> 0 Then
        n_结帐id := 冲销id_In;
      
        If Nvl(n_重结id, 0) <> 0 Then
        
          Select Sum(冲预交)
          Into n_返回值
          From 病人预交记录
          Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
          If Nvl(n_返回值, 0) <> 0 Then
          
            Update 病人预交记录
            Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_返回值, 0)
            Where 结帐id = 冲销id_In And 结算方式 = v_结算方式 And 结算卡序号 = n_卡类别id
            Returning ID Into n_预交id;
          
            If Sql%NotFound Then
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              Insert Into 病人预交记录
                (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算卡序号, 校对标志,
                 结算性质)
              Values
                (n_预交id, 3, Null, 2, r_Balancedata. 病人id, Null, Null, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
                 r_Balancedata. 操作员姓名, n_返回值, r_Balancedata.结帐id, r_Balancedata. 缴款组id, r_Balancedata.结算序号, n_卡类别id, 2,
                 3);
            End If;
            Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_返回值, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
          
            --插入卡结算记录
            --检查消费卡是否正确
            n_序号 := Zl_消费卡目录_Check(n_卡类别id, v_卡号, n_消费卡id, v_Err_Msg);
            If Nvl(n_序号, 0) = 0 Then
              Raise Err_Item;
            End If;
            Begin
              Select Nvl(Max(Nvl(序号, 0)), 0) + 1
              Into n_序号
              From 病人卡结算记录
              Where 接口编号 = n_卡类别id And Nvl(消费卡id, 0) = Nvl(n_消费卡id, 0) And 卡号 = v_卡号;
            Exception
              When Others Then
                n_序号 := 1;
            End;
          
            Select 病人卡结算记录_Id.Nextval Into n_Id From Dual;
          
            Insert Into 病人卡结算记录
              (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
            Values
              (n_Id, n_卡类别id, n_消费卡id, n_序号, 1, v_结算方式, n_返回值, v_卡号, Null, r_Balancedata.收款时间, Null, 0);
          
            --如果消费卡,需同时更改其余额
            If Nvl(n_消费卡id, 0) <> 0 Then
              Update 消费卡目录 Set 余额 = 余额 - n_返回值 Where ID = n_消费卡id;
              If Sql%NotFound Then
                v_Err_Msg := '卡号为' || v_卡号 || '的' || v_名称 || '未找到!';
                Raise Err_Item;
              End If;
            End If;
            Insert Into 病人卡结算对照 (预交id, 卡结算id) Values (n_预交id, n_Id);
            n_结算金额 := n_结算金额 - Nvl(n_返回值, 0);
          End If;
          n_结帐id := n_重结id;
        
        End If;
      
        Update 病人预交记录
        Set 冲预交 = Nvl(冲预交, 0) + n_结算金额
        Where 结帐id = n_结帐id And 结算方式 = v_结算方式 And 结算卡序号 = n_卡类别id
        Returning ID Into n_预交id;
      
        If Sql%NotFound Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算卡序号, 校对标志, 结算性质)
          Values
            (n_预交id, 3, Null, Decode(Nvl(n_重结id, 0), 0, 2, 1), r_Balancedata. 病人id, Null, Null, v_结算方式,
             r_Balancedata. 收款时间, r_Balancedata. 操作员编号, r_Balancedata. 操作员姓名, n_结算金额, n_结帐id, r_Balancedata. 缴款组id,
             r_Balancedata. 结算序号, n_卡类别id, 2, 3);
        End If;
      
        --插入卡结算记录
        n_序号 := Zl_消费卡目录_Check(n_卡类别id, v_卡号, n_消费卡id, v_Err_Msg);
        If Nvl(n_序号, 0) = 0 Then
          Raise Err_Item;
        End If;
      
        Begin
          Select Nvl(Max(Nvl(序号, 0)), 0) + 1
          Into n_序号
          From 病人卡结算记录
          Where 接口编号 = n_卡类别id And Nvl(消费卡id, 0) = Nvl(n_消费卡id, 0) And 卡号 = v_卡号;
        Exception
          When Others Then
            n_序号 := 1;
        End;
      
        Select 病人卡结算记录_Id.Nextval Into n_Id From Dual;
      
        Insert Into 病人卡结算记录
          (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
        Values
          (n_Id, n_卡类别id, n_消费卡id, n_序号, 1, v_结算方式, n_结算金额, v_卡号, Null, r_Balancedata.收款时间, Null, 0);
        --如果消费卡,需同时更改其余额
        If Nvl(n_消费卡id, 0) <> 0 Then
          Update 消费卡目录 Set 余额 = 余额 - n_结算金额 Where ID = n_消费卡id;
          If Sql%NotFound Then
            v_Err_Msg := '卡号为' || v_卡号 || '的' || v_名称 || '未找到!';
            Raise Err_Item;
          End If;
        End If;
        Insert Into 病人卡结算对照 (预交id, 卡结算id) Values (n_预交id, n_Id);
      
        --更新数据(结算方式为NULL的)
        Update 病人预交记录
        Set 冲预交 = 冲预交 - n_结算金额
        Where 结帐id = n_结帐id And 结算方式 Is Null
        Returning Nvl(冲预交, 0) Into n_返回值;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If Nvl(完成退费_In, 0) = 0 Then
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --完成收费,需要处理人员缴款余额,预交记录(结算方式=NULL)
  If Nvl(完成退费_In, 0) = 1 Then
    Update 病人预交记录 Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0 Where 结帐id = 冲销id_In;
    Return;
  End If;

  --1.删除结算方式为NULL的预交记录
  Delete 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 冲销id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未缴款的数据,不能完成结算!';
    Else
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！!';
    End If;
    Raise Err_Item;
  End If;
  If Nvl(n_重结id, 0) <> 0 Then
    Delete 病人预交记录 Where 结帐id = n_重结id And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
    If Sql%NotFound Then
      Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = n_重结id And 结算方式 Is Null;
      If n_Count <> 0 Then
        v_Err_Msg := '还存在未缴款的数据,不能完成结算!';
      Else
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！!';
      End If;
      Raise Err_Item;
    End If;
    Update 病人预交记录 Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0 Where 结帐id = n_重结id;
  
  End If;

  --结算金额为零时，增加一条金额为0的病人预交记录
  Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 冲销id_In;

  If n_Count = 0 Then
    v_结算方式 := 缺省结算方式_In;
    If v_结算方式 Is Null Then
      Begin
        Select 结算方式 Into v_结算方式 From 结算方式应用 Where 应用场合 = '收费' And Nvl(缺省标志, 0) = 1;
      Exception
        When Others Then
          v_结算方式 := Null;
      End;
      If v_结算方式 Is Null Then
        Begin
          Select 名称 Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1;
        Exception
          When Others Then
            v_结算方式 := '现金';
        End;
      End If;
    End If;
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
       交易流水号, 交易说明, 结算号码, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
       r_Balancedata.操作员姓名, 0, r_Balancedata.结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null, Null, Null, Null,
       交易说明_In, Null, 3);
  End If;

  --2.处理缴款数据和找补数据及校对标志更新为0
  Update 病人预交记录 Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0 Where 结帐id = 冲销id_In;

  --3.更新费用状态
  Update 门诊费用记录 Set 费用状态 = 0 Where 结帐id = 冲销id_In;
  If Nvl(n_重结id, 0) <> 0 Then
    Update 门诊费用记录 Set 费用状态 = 0 Where 结帐id = n_重结id;
  End If;

  --4.更新人员缴款数据
  If n_重结id <> 0 Then
    For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
                 From 病人预交记录 A
                 Where a.结帐id In (冲销id_In, n_重结id) And Mod(a.记录性质, 10) <> 1
                 Group By 结算方式, 操作员姓名) Loop
    
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
      Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
      End If;
    End Loop;
  
  Else
    For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
                 From 病人预交记录 A
                 Where a.结帐id = 冲销id_In And Mod(a.记录性质, 10) <> 1
                 Group By 结算方式, 操作员姓名) Loop
    
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
      Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
      End If;
    End Loop;
  End If;
  --消息推送
  Select 病人id_In || ',' || 冲销id_In || ',' || Decode(完成退费_In, 2, 0, 0, 0, 1) Into v_Msg From Dual;
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 5, v_Msg;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊退费结算_Modify;
/

--91225:梁经伙,2015-12-16,传染病管理系统新增加 疾病报告反馈 表
Create Or Replace Procedure Zl_Retu_Clinic
(
  n_Patiid In Number,
  v_Times  In Varchar2,
  n_Flag   In Number
) As
  --------------------------------------------
  --参数:n_Patiid,病人id
  --     v_Times,挂号单号或住院主页id（体检时，挂号单是体检单号）
  --     n_Flag,门诊或住院标志:0-门诊,1-住院,2-体检（此时，只有n_Patiid参数无效）
  --------------------------------------------
  Err_Item Exception;
  v_Err_Msg    Varchar2(100);
  n_System     Number(5);
  n_Opersystem Number(5);
  n_只读       Number(2);

  v_Table    Varchar2(100);
  v_Subtable Varchar2(100);
  v_Field    Varchar2(100);
  v_Subfield Varchar2(100);
  v_Sql      Varchar2(4000);
  v_Fields   Varchar2(4000);

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
  --返回指定病人ID和主页的相关表的子过程
  --------------------------------------------
  Procedure Zl_Retu_Other
  (
    n_Pati_Id 病案主页.病人id%Type,
    n_Page_Id 病案主页.主页id%Type
  ) As
  
  Begin
  
    For R In (Select Column_Value From Table(f_Str2list('病人过敏记录,病人诊断记录,病人手麻记录'))) Loop
      v_Table  := r.Column_Value;
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where 病人id = :1 And 主页id = :2';
      Execute Immediate v_Sql
        Using n_Pati_Id, n_Page_Id;
    
      v_Sql := 'Delete From H' || v_Table || ' Where 病人id = :1 And 主页id = :2';
      Execute Immediate v_Sql
        Using n_Pati_Id, n_Page_Id;
    End Loop;
  End Zl_Retu_Other;

  --------------------------------------------
  --返回指定病人ID和主页的临床路径相关表的子过程
  --------------------------------------------
  Procedure Zl_Retu_Path
  (
    n_Pati_Id 病案主页.病人id%Type,
    n_Page_Id 病案主页.主页id%Type
  ) As
  Begin
    v_Table  := '病人临床路径';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    --病人路径医嘱，在病人医嘱记录转出之后执行
    For P In (Select ID As 路径记录id From H病人临床路径 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop
      For R In (Select Column_Value
                From Table(f_Str2list('病人路径执行,病人合并路径,病人路径评估,病人路径变异,病人路径指标,病人合并路径评估,病人出径记录'))) Loop
        v_Table := r.Column_Value;
        If v_Table = '病人合并路径' Then
          v_Field := '首要路径记录id';
        Else
          v_Field := '路径记录id';
        End If;
     
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                    ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.路径记录id;
      
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.路径记录id;
      End Loop;
    End Loop;
  
    Delete H病人临床路径 Where 病人id = n_Pati_Id And 主页id = n_Page_Id;
  End Zl_Retu_Path;

  --------------------------------------------
  --返回指定病人ID和主页的护理相关表的子过程
  --------------------------------------------
  Procedure Zl_Retu_Tend
  (
    n_Pati_Id 病案主页.病人id%Type,
    n_Page_Id 病案主页.主页id%Type
  ) As
  Begin
  
    v_Table  := '病人护理文件';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    For P In (Select ID As 文件id From H病人护理文件 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop
      For R In (Select Column_Value From Table(f_Str2list('病人护理数据,病人护理打印,病人护理活动项目,病人护理要素内容,产程要素内容'))) Loop
        v_Table  := r.Column_Value;
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                    ' Where 文件id = :1';
        Execute Immediate v_Sql
          Using p.文件id;
      
        If v_Table = '病人护理数据' Then
          v_Fields := Getfields('病人护理明细');
          v_Sql    := 'Insert Into 病人护理明细(' || v_Fields || ') Select ' || v_Fields ||
                      ' From H病人护理明细 Where 记录id In (Select ID From H病人护理数据 Where 文件id = :1)';
          Execute Immediate v_Sql
            Using p.文件id;
        
          v_Sql := 'Delete H病人护理明细 Where 记录id In (Select ID From H病人护理数据 Where 文件id = :1)';
          Execute Immediate v_Sql
            Using p.文件id;
        End If;
      
        v_Sql := 'Delete H' || v_Table || ' Where 文件id = :1';
        Execute Immediate v_Sql
          Using p.文件id;
      End Loop;
    End Loop;
  
    Delete H病人护理文件 Where 病人id = n_Pati_Id And 主页id = n_Page_Id;


                   --老版护理系统数据
                   ------------------------------------------------------------------------
                   v_Table  := '病人护理记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
                   
                   For P In (Select ID From H病人护理记录 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop      
        v_Table  := '病人护理内容';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                    ' Where 记录ID = :1';
        Execute Immediate v_Sql
          Using p.ID;
      
        v_Sql := 'Delete H' || v_Table || ' Where 记录ID = :1';
        Execute Immediate v_Sql
          Using p.ID;     
    End Loop;
  
    Delete H病人护理记录 Where 病人id = n_Pati_Id And 主页id = n_Page_Id;
  End Zl_Retu_Tend;

  --------------------------------------------
  --返回指定ID的病人新版电子病历记录子过程
  --------------------------------------------
  Procedure Zl_Retu_Epr(n_Rec_Id H电子病历记录.Id%Type) As
    v_Field Varchar(100);
  Begin
    v_Table  := '电子病历记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where id = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    --病人诊断记录在Zl_Retu_Other中已转回（无病历ID外键）
    --影像报告驳回,病人医嘱报告,报告查阅记录,这几张表的数据在Zl_Retu_Order中转回医嘱后再处理
    For R In (Select Column_Value From Table(f_Str2list('电子病历附件,电子病历格式,电子病历内容,疾病申报记录,疾病报告反馈'))) Loop
      v_Table := r.Column_Value;
      If v_Table = '电子病历附件' Then
        v_Field := '病历id';
      Else
        v_Field := '文件id';
      End If;
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      If v_Table = '电子病历内容' Then
        v_Fields := Getfields('电子病历图形');
        v_Sql    := 'Insert Into 电子病历图形(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H电子病历图形 Where 对象id In (Select ID From H电子病历内容 Where 文件id = :1 And 对象类型 = 5)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H电子病历图形 Where 对象id In (Select ID From H电子病历内容 Where 文件id = n_Rec_Id And 对象类型 = 5);
      End If;
    
      v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    End Loop;
  
    Delete H电子病历记录 Where ID = n_Rec_Id;
  End Zl_Retu_Epr;
  --------------------------------------------
  --返回指定ID的病人医嘱记录子过程，必须在病历、临床路径转出之后执行(病人医嘱报告,影像报告驳回，病人路径医嘱)
  --在Zl_Retu_Other中已转回了"病人诊断记录",转回"病人诊断医嘱"时不用再转
  --------------------------------------------
  Procedure Zl_Retu_Order(n_Rec_Id H病人医嘱记录.Id%Type) As
  Begin
    v_Table  := '病人医嘱记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where id = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    --以"医嘱ID,发送号"为外键的，都按医嘱ID直接转回，只需要排在"病人医嘱发送"之后即可
    --由于外键关系，"报告查阅记录"须在"病人医嘱报告"后面
    For P In (Select Column_Value
              From Table(f_Str2list('病人医嘱计价,病人医嘱状态,病人医嘱发送,病人医嘱附费,病人医嘱附件,病人医嘱执行,病人医嘱打印,输血申请记录,输血检验结果,' ||
                                     '医嘱执行打印,医嘱执行时间,医嘱执行计价,执行打印记录,病人诊断医嘱,病人路径医嘱,病人医嘱报告,报告查阅记录,' ||
                                     '影像报告驳回,影像报告记录,影像报告操作记录,影像检查记录,影像申请单图像,影像收藏内容,影像危急值记录,检验标本记录,检验试剂记录,检验拒收记录'))) Loop
      v_Table := p.Column_Value;
      If Instr('病人路径医嘱', v_Table) > 0 Then
        v_Field := '病人医嘱ID';
      Else
        v_Field := '医嘱ID';
      End If;
    
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      If v_Table = '病人医嘱状态' Then
        v_Fields := Getfields('医嘱签名记录');
        v_Sql    := 'Insert Into 医嘱签名记录(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H医嘱签名记录 Where ID In (Select 签名id From H病人医嘱状态 Where 医嘱id = :1 And 签名id Is Not Null)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete H医嘱签名记录
        Where ID In (Select 签名id From H病人医嘱状态 Where 医嘱id = n_Rec_Id And 签名id Is Not Null);
      
      Elsif v_Table = '病人医嘱发送' Then
        v_Fields := Getfields('诊疗单据打印');
        v_Sql    := 'Insert Into 诊疗单据打印(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H诊疗单据打印 Where (NO, 记录性质) In (Select NO, 记录性质 From H病人医嘱发送 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete H诊疗单据打印 Where (NO, 记录性质) In (Select NO, 记录性质 From H病人医嘱发送 Where 医嘱id = n_Rec_Id);
      
      Elsif v_Table = '影像检查记录' Then
        v_Fields := Getfields('影像检查序列');
        v_Sql    := 'Insert Into 影像检查序列(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H影像检查序列 Where 检查uid In (Select 检查uid From H影像检查记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('影像检查图象');
        v_Sql    := 'Insert Into 影像检查图象(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H影像检查图象 Where 序列uid In (Select b.序列uid From H影像检查记录 A, H影像检查序列 B Where a.医嘱id = :1 And a.检查uid = b.检查uid)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H影像检查图象
        Where 序列uid In (Select b.序列uid
                        From H影像检查记录 A, H影像检查序列 B
                        Where a.医嘱id = n_Rec_Id And a.检查uid = b.检查uid);
        Delete H影像检查序列 Where 检查uid In (Select 检查uid From H影像检查记录 Where 医嘱id = n_Rec_Id);
      
      Elsif v_Table = '检验标本记录' Then
        For R In (Select Column_Value
                  From Table(f_Str2list('检验申请项目,检验分析记录,检验项目分布,检验质控记录,检验操作记录,检验签名记录,检验图像结果'))) Loop
          v_Subtable := r.Column_Value;
          If v_Subtable = '检验签名记录' Then
            v_Subfield := '检验标本ID';
          Else
            v_Subfield := '标本ID';
          End If;
          v_Fields := Getfields(v_Subtable);
          v_Sql    := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' || v_Fields || ' From H' ||
                      v_Subtable || ' Where ' || v_Subfield || ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        
          v_Sql := 'Delete H' || v_Subtable || ' Where ' || v_Subfield ||
                   ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        End Loop;
      
        v_Fields := Getfields('检验普通结果');
        v_Sql    := 'Insert Into 检验普通结果(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('检验药敏结果');
        v_Sql    := 'Insert Into 检验药敏结果(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H检验药敏结果 Where 细菌结果id In (Select ID From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = :1))';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('检验质控报告');
        v_Sql    := 'Insert Into 检验质控报告(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H检验质控报告 Where 结果ID In (Select ID From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = :1))';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H检验药敏结果
        Where 细菌结果id In
              (Select ID From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = n_Rec_Id));
        Delete H检验质控报告
        Where 结果id In
              (Select ID From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = n_Rec_Id));
      
        Delete H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = n_Rec_Id);
      End If;
    
      v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    End Loop;
  
    --手麻数据
    If n_Opersystem > 0 Then
      Execute Immediate 'zl24_Retu_Oper(:1)'
        Using n_Rec_Id;
    End If;
  
    Delete H病人医嘱记录 Where ID = n_Rec_Id;
  End Zl_Retu_Order;

  --------------------------------------------
  --以下为主程序体
  --------------------------------------------
Begin
  ----------------------------------------------------------------------------------------------------------
  --对基于视图的转储方案进行了只读判断.
  Select 编号 Into n_System From zlSystems Where Upper(所有者) = Zl_Owner And 编号 Like '1%';
  Begin
    Select Nvl(只读, 0) Into n_只读 From Zlbakspaces Where 系统 = n_System And 当前 = 1;
  Exception
    When Others Then
      v_Err_Msg := '[ZLSOFT]当前没有可用的历史数据空间,不能继续![ZLSOFT]';
      Raise Err_Item;
  End;
  If n_只读 = 1 Then
    v_Err_Msg := '[ZLSOFT]历史数据空间目前的状态为只读,不能继续![ZLSOFT]';
    Raise Err_Item;
  End If;

  --对基于视图的转储方案进行了只读判断.
  n_Opersystem := 0;
  Select 编号 Into n_Opersystem From zlSystems Where Upper(所有者) = Zl_Owner And 编号 Like '24%';
  If n_Opersystem > 0 Then
    Begin
      Select Nvl(只读, 0) Into n_只读 From Zlbakspaces Where 系统 = n_Opersystem And 当前 = 1;
    Exception
      When Others Then
        v_Err_Msg := '[ZLSOFT]当前没有可用的手麻子系统历史数据空间,不能继续![ZLSOFT]';
        Raise Err_Item;
    End;
    If n_只读 = 1 Then
      v_Err_Msg := '[ZLSOFT]手麻子系统历史数据空间目前的状态为只读,不能继续![ZLSOFT]';
     Raise Err_Item;
    End If;
  End If;

  --1.门诊病人，按挂号单抽回
  If n_Flag = 0 Then
    v_Table  := '病人挂号记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where NO =:1 ';
    Execute Immediate v_Sql
      Using v_Times;
  
    For r_Other In (Select ID, 病人id From H病人挂号记录 Where NO = v_Times) Loop
      Zl_Retu_Other(r_Other.病人id, r_Other.Id);
    End Loop;
  
    For r_Epr In (Select /*+ Rule*/
                   b.Id
                  From H病人挂号记录 A, H电子病历记录 B
                  Where a.No = v_Times And a.病人id = n_Patiid And b.病人id = a.病人id And b.主页id = a.Id) Loop
      Zl_Retu_Epr(r_Epr.Id);
    End Loop;
  
    For r_Order In (Select ID From H病人医嘱记录 Where 病人来源 <> 4 And 病人id = n_Patiid And 挂号单 = v_Times) Loop
      Zl_Retu_Order(r_Order.Id);
    End Loop;
  
    --转诊记录
    v_Table  := '病人转诊记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where NO =:1';
   Execute Immediate v_Sql
      Using v_Times;
  
    Delete H病人转诊记录 Where NO = v_Times;
    Delete H病人挂号记录 Where NO = v_Times;
  
    --2.住院病人，按病人ID和主页ID抽回
  Elsif n_Flag = 1 Then
    Zl_Retu_Other(n_Patiid, To_Number(v_Times));
    Zl_Retu_Path(n_Patiid, To_Number(v_Times));
  
    --先转病历，再转医嘱（影像报告驳回，病人医嘱报告这类又有病历又有医嘱的子表，在医嘱转回后处理）
    For r_Epr In (Select ID From H电子病历记录 Where 病人id = n_Patiid And 主页id = To_Number(v_Times)) Loop
      Zl_Retu_Epr(r_Epr.Id);
    End Loop;
  
    Zl_Retu_Tend(n_Patiid, To_Number(v_Times));
  
    For r_Order In (Select ID From H病人医嘱记录 Where 病人id = n_Patiid And 主页id = To_Number(v_Times)) Loop
      Zl_Retu_Order(r_Order.Id);
    End Loop;
    Update 病案主页 Set 数据转出 = 0 Where 病人id = n_Patiid And 主页id = To_Number(v_Times);
  
    --3.体检病人
  Elsif n_Flag = 2 Then
    Zl_Retu_Other(n_Patiid, v_Times);
  
    For r_Cpr In (Select ID From H病人医嘱记录 Where 病人来源 = 4 And 挂号单 = v_Times) Loop
      Zl_Retu_Order(r_Cpr.Id);
    End Loop;
  
  End If;

  Begin
    Execute Immediate 'Update zlbakInfo  set 最后转储日期=sysdate where 系统=' || n_System;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm || ':' || v_Sql);
End Zl_Retu_Clinic;
/

--91225:梁经伙,2015-12-16,传染病管理系统新增表 疾病报告反馈
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
                                         ',病人护理活动项目_UQ_页号,病人护理要素内容_UQ_页号,产程要素内容_PK,电子病历记录_PK,电子病历附件_PK,电子病历格式_PK,电子病历内容_UQ_对象序号,电子病历图形_PK,疾病申报记录_PK,疾病报告反馈_PK' ||
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

--91225:梁经伙,2015-12-16,传染病管理系统新增 疾病报告反馈 表
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
  
  Update /*+ rule*/ 疾病报告反馈
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

--92335:李南春,2016-01-18,三方支付新模式及过程拆分
--91561:刘尔旋,2015-12-14,预约接收免费号产生预交记录
Create Or Replace Procedure Zl_预约挂号接收_Insert
(
  No_In            门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  领用id_In        票据使用明细.领用id%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  结算方式_In      病人预交记录.结算方式%Type, --现金的结算名称
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额
  预交支付_In      病人预交记录.冲预交%Type, --挂号时使用的预交金额
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额
  发生时间_In      门诊费用记录.发生时间%Type,
  号序_In          挂号序号状态.序号%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  生成队列_In      Number := 0,
  登记时间_In      门诊费用记录.登记时间%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  冲预交病人ids_In Varchar2 := Null,
  三方调用_In      Number := 0,
  更新交款余额_In  Number := 0--是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.病人id, a.记录状态, a.预交类别, a.No, Nvl(a.金额, 0) As 金额
           From 病人预交记录 A,
                (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                  From 病人预交记录 A
                  Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And
                        a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(a.预交类别, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.金额, 0)) <> 0) B
           Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                 a.No = b.No And a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And
                 Nvl(a.预交类别, 2) = 1
           Union All
           Select 0 As ID, Max(病人id) As 病人id, 记录状态, 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
           From 病人预交记录
           Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And
                 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1 Having
            Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
           Group By 记录状态, 预交类别, NO)
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), ID, NO, 预交类别;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;
  v_号别     门诊费用记录.计算单位%Type;
  v_号序     门诊费用记录.发药窗口%Type;
  v_排队号码 排队叫号队列.排队号码 %Type;
  v_预约方式 病人挂号记录.预约方式 %Type;

  n_打印id        票据打印内容.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  v_冲预交病人ids Varchar2(4000);

  n_挂号id         病人挂号记录.Id%Type;
  n_分诊台签到排队 Number;
  n_组id           财务缴款分组.Id%Type;
  n_Count          Number(18);
  n_排队           Number;
  n_当天排队       Number;
  n_当前金额       病人预交记录.金额%Type;
  n_预交id         病人预交记录.Id%Type;
  n_消费卡id       消费卡目录.Id%Type;
  n_自制卡         Number;

  d_Date     Date;
  d_预约时间 门诊费用记录.发生时间%Type;
  d_发生时间 Date;
  d_排队时间 Date;
  n_时段     Number := 0;
  n_存在     Number := 0;
  v_排队序号 排队叫号队列.排队序号%Type;
  n_结算模式 病人信息.结算模式%Type;
  n_票种     票据使用明细.票种%Type;
  v_付款方式 病人挂号记录.医疗付款方式%Type;
  n_接收模式 Number := 0;
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  n_接收模式      := Nvl(zl_GetSysParameter(64, 1111), 0);

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
  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式
    Into v_号别, v_号序, d_预约时间, d_发生时间, v_预约方式
    From 病人挂号记录
    Where 记录性质 = 2 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      v_Err_Msg := '当前预约挂号单已被其它人接收';
      Raise Err_Item;
  End;

  --判断是否分时段
  Begin
    Select 1
    Into n_时段
    From Dual
    Where Exists (Select 1
           From 挂号安排时段 A, 挂号安排 B
           Where a.安排id = b.Id And b.号码 = v_号别 And Rownum < 2
           Union All
           Select 1
           From 挂号计划时段 C, 挂号安排计划 D 　
           Where c.计划id = d.Id And d.号码 = v_号别 And d.生效时间 > Sysdate And Rownum < 2);
  Exception
    When Others Then
      n_时段 := 0;
  End;

  If n_时段 = 0 And 三方调用_In = 0 Then
    If n_接收模式 = 0 Then
      If Trunc(发生时间_In) = Trunc(Sysdate) Then
        d_发生时间 := 发生时间_In;
      Else
        d_发生时间 := Sysdate;
      End If;
    Else
      d_发生时间 := 发生时间_In;
    End If;
  Else
    If Not 发生时间_In Is Null Then
      d_发生时间 := 发生时间_In;
    End If;
  End If;
  If Not v_号序 Is Null Then
    If 号序_In Is Null Then
      Delete 挂号序号状态 Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
    Else
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
      
        If n_时段 = 0 And 三方调用_In = 0 Then
          --提前接收或延迟接收
          Delete 挂号序号状态 Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
          Begin
            Select 1 Into n_存在 From 挂号序号状态 Where 号码 = v_号别 And 日期 = Trunc(Sysdate) And 序号 = v_号序;
          Exception
            When Others Then
              n_存在 := 0;
          End;
          If n_存在 = 0 Then
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
            Values
              (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
          Else
            --号码已被使用的情况
            Begin
              v_号序 := 1;
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
              Values
                (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            Exception
              When Others Then
                Select Min(序号 + 1)
                Into v_号序
                From 挂号序号状态 A
                Where 号码 = v_号别 And 日期 = Trunc(Sysdate) And Not Exists
                 (Select 1 From 挂号序号状态 Where 号码 = a.号码 And 日期 = a.日期 And 序号 = a.序号 + 1);
                Insert Into 挂号序号状态
                  (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
                Values
                  (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            End;
          End If;
        Else
          Update 挂号序号状态
          Set 状态 = 1, 登记时间 = Sysdate
          Where Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序 And 号码 = v_号别 And 状态 = 2;
          If Sql% NotFound Then
            Begin
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
              Values
                (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            Exception
              When Others Then
                v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
                Raise Err_Item;
            End;
          End If;
        
        End If;
      
      Else
        Update 挂号序号状态
        Set 序号 = 号序_In, 状态 = 1, 登记时间 = Sysdate
        Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
            Values
              (v_号别, Trunc(d_发生时间), v_号序, 1, 操作员姓名_In, Sysdate);
          Exception
            When Others Then
              v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  Else
    If Not 号序_In Is Null Then
      Begin
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
        Values
          (v_号别, Trunc(Sysdate), 号序_In, 1, 操作员姓名_In, Sysdate);
      Exception
        When Others Then
          v_Err_Msg := '序号' || 号序_In || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
      End;
      v_号序 := 号序_In;
    Else
      v_号序 := Null;
    End If;
  End If;

  --更新门诊费用记录
  Update 门诊费用记录
  Set 记录状态 = 1, 实际票号 = Decode(Nvl(记帐费用_In, 0), 1, Null, 票据号_In), 结帐id = Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In),
      结帐金额 = Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额), 发药窗口 = 诊室_In, 病人id = 病人id_In, 标识号 = 门诊号_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
      性别 = 性别_In, 付款方式 = 付款方式_In, 费别 = 费别_In, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In,
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In
  Where 记录状态 = 1 And NO = No_In And 记录性质 = 2
  Returning ID Into n_挂号id;
  If Sql%NotFound Then
    Begin
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         摘要, 号序, 预约, 预约方式, 接收人, 接收时间, 预约时间, 险类, 医疗付款方式)
        Select n_挂号id, No_In, 1, 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, 执行人, 0, Null,
               登记时间, 发生时间, 操作员编号, 操作员姓名, 摘要, v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In, Nvl(登记时间_In, Sysdate), 发生时间,
               Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '由于并发原因,单据号为【' || No_In || '】的病人' || 姓名_In || '已经被接收';
        Raise Err_Item;
    End;
  End If;

  --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
  If Nvl(生成队列_In, 0) <> 0 Then
    n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
    If Nvl(n_分诊台签到排队, 0) = 0 Then
      For v_挂号 In (Select ID, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序 From 病人挂号记录 Where NO = No_In) Loop
      
        Begin
          Select 1,
                 Case
                   When 排队时间 < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_排队, n_当天排队
          From 排队叫号队列
          Where 业务类型 = 0 And 业务id = v_挂号.Id And Rownum <= 1;
        Exception
          When Others Then
            n_排队 := 0;
        End;
        If n_排队 = 0 Then
          --产生队列
          --按”执行部门”产生队列
          n_挂号id   := v_挂号.Id;
          v_队列名称 := v_挂号.执行部门id;
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, n_挂号id, v_挂号.号别 || '|' || v_挂号.号序);
          v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
          d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_挂号.号别, v_挂号.号序, v_挂号.发生时间);
          --   队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In,
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, v_挂号.执行部门id, v_排队号码, Null, 姓名_In, 病人id_In, v_挂号.诊室, v_挂号.执行人, d_排队时间,
                           v_预约方式, Null, v_排队序号);
        Elsif Nvl(n_当天排队, 0) = 1 Then
          --更新队列号
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, v_挂号.Id, v_挂号.号别 || '|' || Nvl(v_挂号.号序, 0));
          v_排队序号 := Zlgetsequencenum(0, v_挂号.Id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人, v_排队号码, v_排队序号);
        
        Else
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人);
        End If;
      End Loop;
    End If;
  End If;

  --汇总结算到病人预交记录
  If (Nvl(现金支付_In, 0) <> 0 Or (Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0)) And
     Nvl(记帐费用_In, 0) = 0 Then
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算序号,
       结算性质)
    Values
      (n_预交id, 4, 1, No_In, 病人id_In, Nvl(结算方式_In, v_现金), 现金支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id,
       卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 结帐id_In, 4);
  
    If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
    
      n_消费卡id := Null;
      Begin
        Select Nvl(自制卡, 0), 1 Into n_自制卡, n_Count From 卡消费接口目录 Where 编号 = 结算卡序号_In;
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
        Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In And
              序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In);
      End If;
      Zl_病人卡结算记录_Insert(结算卡序号_In, n_消费卡id, 结算方式_In, 现金支付_In, 卡号_In, Null, 登记时间_In, Null, 结帐id_In, n_预交id);
    End If;
  
  End If;

  --对于就诊卡通过预交金挂号
  If Nvl(预交支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.Id <> 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.Id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 登记时间_In,
               操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结帐id_In, 4
        From 病人预交记录
        Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(r_Deposit.预交类别, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 类型, 预交余额, 性质)
        Values
          (r_Deposit.病人id, Nvl(r_Deposit.预交类别, 2), -1 * n_当前金额, 1);
        n_返回值 := -1 * n_当前金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
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
  End If;

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(现金支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In,0)=0 Then
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
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = Nvl(结算方式_In, v_现金) And Nvl(余额, 0) = 0;
    End If;
  End If;

  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In,0)=0 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 个帐支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
      n_返回值 := 个帐支付_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And Nvl(余额, 0) = 0;
    End If;
  End If;

  --处理票据使用情况
  If 票据号_In Is Not Null And Nvl(记帐费用_In, 0) = 0 Then
    Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  
    --当前票据的票种
    Select 票种 Into n_票种 From 票据领用记录 Where ID = Nvl(领用id_In, 0);
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, No_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, n_票种, 票据号_In, 1, 1, 领用id_In, n_打印id, d_Date, 操作员姓名_In);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = d_Date
    Where ID = Nvl(领用id_In, 0);
  End If;

  If Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In And Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, 1, Nvl(c_费用.实收金额, 0), 0);
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    n_结算模式 := 0;
    Update 病人信息
    Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In
    Where 病人id = 病人id_In
    Returning Nvl(结算模式, 0) Into n_结算模式;
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
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
  End If;

  --病人担保信息
  If 病人id_In Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Exists (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) > d_Date;
    End If;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 1, n_挂号id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_Insert;
/

--91085:王振涛,2015-12-08,取消审核，同步取消提示
Create Or Replace Procedure Zl_检验标本记录_审核取消(Id_In 检验标本记录.Id%Type) Is
  --查找当前标本的相关申请
  Cursor c_Samplequest Is
    Select Distinct 医嘱id
    From (Select 医嘱id
           From 检验标本记录
           Where ID = Id_In
           Union
           Select 医嘱id From 检验项目分布 Where 标本id = Id_In);

  v_主页id Number(18);
  v_No     Varchar2(20);
  v_Temp   Varchar2(255);
  v_Fileid Number(18);

  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_当前时间 Date;
Begin
  --操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_当前时间 := Sysdate;


  --1.取消标本审核
  Update 检验标本记录
  Set 审核人 = Null, 审核时间 = Null, 打印次数 = Null, 审核未通过 = Null, 样本状态 = 1
  Where ID = Id_In;
  --Delete 检验签名记录 Where 检验标本id = Id_In;
  --记录审核过程
  Insert Into 检验操作记录
    (ID, 标本id, 操作类型, 操作员, 操作时间)
  Values
    (检验操作记录_Id.Nextval, Id_In, 1, v_人员姓名, Sysdate);

  --2.检查当前标本相关的申请的相关标本
  For r_Samplequest In c_Samplequest Loop
  
    --1.置申请单的正在执行状态
    Update 病人医嘱发送
    Set 执行状态 = 3, 完成人 = Null, 完成时间 = Null
    Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id = 相关id);
  
    Update 病人医嘱发送
    Set 执行状态 = 3, 完成人 = Null, 完成时间 = Null
    Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id = ID) And Nvl(采样人, '空空') = '空空';
  
    Begin
      Select 病历id Into v_Fileid From 病人医嘱报告 Where 医嘱id = r_Samplequest.医嘱id;
      Zl_报告查阅记录_Cancel(r_Samplequest.医嘱id, v_Fileid, Null);
      Delete 病人医嘱报告 Where 医嘱id = r_Samplequest.医嘱id;
    Exception
      When Others Then
        v_Fileid := 0;
    End;
    If v_Fileid <> 0 Then
      Delete 电子病历记录 Where ID = v_Fileid;
      Delete 电子病历内容 Where 文件id = v_Fileid;
    End If;
  
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验标本记录_审核取消;
/

--91259:刘硕,2015-12-08,地址复制，缓冲区过小问题处理
Create Or Replace Function Zl_Adderss_Structure(v_Addressinfo Varchar2) Return Varchar2 Is
  --返回结构：省,省编码,是否虚拟,是否不显示,是否只有虚拟级|市,市编码,是否虚拟,是否不显示,是否只有虚拟级
  --          |区县,区县编码,是否虚拟,是否不显示,是否只有虚拟级|乡镇,乡镇编码,是否虚拟,是否不显示,是否只有虚拟级
  --          |街道,街道编码,是否虚拟,是否不显示,是否只有虚拟级
  v_省       Varchar2(100);
  v_Code省   Varchar2(15);
  v_Info省   Varchar2(150);
  v_市       Varchar2(100);
  v_Code市   Varchar2(15);
  v_Info市   Varchar2(150);
  v_区县     Varchar2(100);
  v_Code区县 Varchar2(15);
  v_Info区县 Varchar2(150);
  v_乡镇     Varchar2(100);
  v_Code乡镇 Varchar2(15);
  v_Info乡镇 Varchar2(150);
  v_街道     Varchar2(500);
  v_Code街道 Varchar2(15);
  v_Info街道 Varchar2(550);
  v_Tmp      Varchar2(100);
  v_Adrstmp  Varchar2(500);
  n_Pos      Number(5);
  n_虚拟     Number(1);
  n_不显示   Number(1);
  n_Count    Number(3);
  v_Return   Varchar2(700);
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

--91090:王振涛,2015-12-07,报告时间更新
Create Or Replace Procedure Zl_检验普通结果_Batchupdate
(
  检验标本id_In In 检验普通结果.检验标本id%Type,
  仪器id_In     In 检验普通结果.仪器id%Type := Null,
  标本类型_In   In Varchar2,
  性别_In       In Number,
  出生日期_In   In Date,
  检验指标_In   In Varchar2, --格式：项目ID^值|。。。
  微生物_In     In Number := 0, --1=微生物
  酶标板id_In   In 检验酶标记录.Id%Type := Null
) Is
  v_记录类型 Number(2);

  v_Temp           Varchar2(255);
  v_人员姓名       人员表.姓名%Type;
  v_Count          Number;
  v_医嘱id         检验标本记录.医嘱id%Type;
  v_药敏结果       检验药敏结果.结果%Type;
  v_细菌id         检验细菌.Id%Type;
  v_检验普通结果id 检验普通结果.Id%Type;
  v_药敏方法       检验药敏结果.药敏方法%Type;
  v_Od             检验普通结果.Od%Type;
  v_Cutoff         检验普通结果.Cutoff%Type;
  v_Sco            检验普通结果.Sco%Type;

  v_Records   Varchar2(4000);
  v_Currrec   Varchar2(100);
  v_年龄      检验标本记录.年龄%Type;
  v_项目id    检验普通结果.检验项目id%Type;
  v_检验结果  检验普通结果.检验结果%Type;
  v_检验结果1 检验普通结果.检验结果%Type;
  v_临时结果  检验普通结果.检验结果%Type;

  v_审核字串     Varchar2(4000);
  v_仪器审核     Number;
  v_仪器审核完成 Number;
  v_仪器审核内容 Varchar2(4000);

  v_Resultref  Varchar2(1000);
  v_参考值     Varchar2(1000);
  v_参考值1    Varchar2(1000);
  v_危急参考   Varchar2(1000);
  v_结果标志   Number;
  v_加算值     Number;
  v_换算比     Number;
  v_小数       Number;
  v_警戒下限   Number;
  v_警戒上限   Number;
  v_多参考     Number;
  v_申请科室id Number;

  v_Lower Number;
  v_Upper Number;

  Function Zlval(Vstr In Varchar2) Return Number Is
    Result Number(16, 6);
    Intbit Number(8);
    Strnum Varchar(10);
    Function Sub_Is_Number(v_In In Varchar2) Return Boolean Is
      n_Tmp Number;
    Begin
      n_Tmp := To_Number(v_In);
      If n_Tmp Is Not Null Then
        Return True;
      End If;
    Exception
      When Others Then
        Return False;
    End Sub_Is_Number;
  Begin
    Strnum := '';
    If Sub_Is_Number(Vstr) = True Then
      Result := To_Number(Nvl(Vstr, 0));
      Return(Result);
    Else
      For Intbit In 1 .. 10 Loop
        If Instr('0123456789.', Substr(Vstr, Intbit, 1)) = 0 Then
          Exit;
        End If;
        Strnum := Strnum || Substr(Vstr, Intbit, 1);
        Null;
      End Loop;
      Result := To_Number(Nvl(Strnum, 0));
      Return(Result);
    End If;
  End Zlval;
  -- >>>>>>>>>>>>>>>>>>  检查是否数字的函数  <<<<<<<<<<<<<<<<<<
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
Begin
  --操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Nvl(报告结果, 0), Nvl(医嘱id, 0), 年龄, 申请科室id
  Into v_记录类型, v_医嘱id, v_年龄, v_申请科室id
  From 检验标本记录
  Where ID = 检验标本id_In;
  If Sql%Rowcount > 0 Then
    v_Records := 检验指标_In || '|';
    While v_Records Is Not Null Loop
      If Nvl(微生物_In, 0) = 0 Then
        --普通标本处理
        v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
        If Instr(v_Currrec, '<Split>') > 0 Then
          v_审核字串 := Substr(v_Currrec, Instr(v_Currrec, '<Split>') + 7);
          v_Currrec  := Substr(v_Currrec, 1, Instr(v_Currrec, '<Split>') - 1);
        End If;
        v_项目id := To_Number(Substr(v_Currrec, 1, Instr(v_Currrec, '^') - 1));
        v_Temp   := Substr(v_Currrec, Instr(v_Currrec, '^') + 1);
        If Instr(v_Temp, '^') > 0 Then
          v_检验结果 := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
          v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
          v_Od       := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
          v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
          v_Cutoff   := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
          v_Sco      := Substr(v_Temp, Instr(v_Temp, '^') + 1);
        Else
          v_检验结果 := Substr(v_Currrec, Instr(v_Currrec, '^') + 1);
          v_Od       := Null;
          v_Cutoff   := Null;
          v_Sco      := Null;
        End If;
        If v_审核字串 Is Not Null Then
          If Instr(v_审核字串, '^') > 0 Then
            v_仪器审核     := Substr(v_审核字串, 1, Instr(v_审核字串, '^') - 1);
            v_审核字串     := Substr(v_审核字串, Instr(v_审核字串, '^') + 1);
            v_仪器审核内容 := v_审核字串;
          Else
            v_仪器审核     := v_审核字串;
            v_仪器审核内容 := Null;
          End If;
        End If;
        v_小数 := 2;
        Select b.警戒下限, b.警戒上限, Max(a.加算值), Max(a.换算比), Max(Nvl(a.小数位数, 2))
        Into v_警戒下限, v_警戒上限, v_加算值, v_换算比, v_小数
        From 检验仪器项目 A, 检验项目 B
        Where a.项目id = b.诊治项目id And a.仪器id = 仪器id_In And a.项目id = v_项目id
        Group By b.警戒下限, b.警戒上限;
        If v_加算值 Is Null Then
          v_加算值 := 0;
        End If;
        If v_换算比 Is Null Then
          v_换算比 := 1;
        End If;
      
        If Instr(v_检验结果, '+') = 0 Then
          v_检验结果1 := v_检验结果;
          Begin
            If v_加算值 <> 0 Or v_换算比 <> 1 Then
              v_检验结果 := (v_检验结果 + v_加算值) * v_换算比;
            
            End If;
            If Instr(v_检验结果, 'E') = 0 Then
              If Zlval(v_检验结果) = v_检验结果 Then
                If v_小数 = 0 Then
                  v_检验结果 := Trim(To_Char(To_Number(Nvl(Trim(v_检验结果), 0)), '999999999'));
                Else
                  v_检验结果 := Trim(To_Char(To_Number(Nvl(Trim(v_检验结果), 0)), '999999990' || Substr('.000000', 1, 1 + v_小数)));
                End If;
              End If;
            End If;
          Exception
            When Others Then
              v_检验结果 := v_检验结果1;
          End;
        End If;
      
        --获取参考并判断结果标志
        v_Resultref := Zlgetreference(v_项目id, 标本类型_In, 性别_In, 出生日期_In, 仪器id_In, v_年龄, v_申请科室id);
        v_参考值1   := v_Resultref;
        Select Nvl(多参考, 0) Into v_多参考 From 检验项目 Where 诊治项目id = v_项目id;
        If Instr(v_Resultref, Chr(13) || Chr(10)) > 0 Then
          v_Resultref := Substr(v_Resultref, 1, Instr(v_Resultref, Chr(13) || Chr(10)) - 1);
        Else
          v_多参考 := 0;
        End If;
        v_危急参考 := Zl_Get_Reference(2, v_项目id, 标本类型_In, 性别_In, 出生日期_In, 仪器id_In, v_年龄, v_申请科室id);
      
        v_结果标志 := 1;
        v_临时结果 := v_检验结果;
        v_检验结果 := Replace(Replace(v_检验结果, '>', ''), '<', '');
        --处理">"和"<"的判断
        --If Instr(v_检验结果, '>') > 0 Then
        --  v_参考值   := v_Resultref;
        --  v_结果标志 := 3;
        --End If;
        --If Instr(v_检验结果, '<') > 0 Then
        --  v_参考值   := v_Resultref;
        --  v_结果标志 := 2;
        --End If;
      
        If v_结果标志 = 1 Then
          If (Instr(v_检验结果, '+') > 0 Or Instr(v_检验结果, '*') > 0) And Sub_Is_Number(v_检验结果) = False Then
            v_参考值   := v_Resultref;
            v_结果标志 := 4;
          Else
            If v_Resultref Is Null Or Sub_Is_Number(v_检验结果) = False Then
              v_参考值   := Nvl(v_Resultref, '');
              v_结果标志 := 1;
            Else
              v_参考值 := Nvl(v_Resultref, '');
              If Length(v_Resultref) > 0 Then
                If Instr(v_Resultref, '～') > 0 Then
                  If Instr(v_Resultref, '～') < Length(v_Resultref) Then
                    v_Upper := Zlval(Nvl(Substr(v_Resultref, Instr(v_Resultref, '～') + 1), 0));
                  Else
                    v_Upper := 0;
                  End If;
                  v_Lower := Zlval(Nvl(Substr(v_Resultref, 1, Instr(v_Resultref, '～') - 1), 0));
                Else
                  v_Upper := Zlval(v_Resultref);
                  v_Lower := Zlval(v_Resultref);
                End If;
                If Nvl(v_检验结果, 0) > v_Upper And v_Upper <> 0 Then
                  v_结果标志 := 3;
                Else
                  If Nvl(v_检验结果, 0) < v_Lower And v_Lower <> 0 Then
                    v_结果标志 := 2;
                  Else
                    v_结果标志 := 1;
                  End If;
                End If;
                If v_结果标志 <> 1 Then
                  If Sub_Is_Number(v_检验结果) = True Then
                    If Instr(v_危急参考, '～') > 0 Then
                    
                      If Nvl(Zlval(v_检验结果), 0) < To_Number(Substr(v_危急参考, 1, Instr(v_危急参考, '～') - 1)) Then
                        v_结果标志 := 5;
                      End If;
                    
                      If Nvl(Zlval(v_检验结果), 0) > To_Number(Substr(v_危急参考, 1, Instr(v_危急参考, '～') - 1)) Then
                        v_结果标志 := 6;
                      End If;
                    
                    End If;
                  End If;
                End If;
              Else
                v_结果标志 := 1;
                If Sub_Is_Number(v_检验结果) = True Then
                  If Instr(v_危急参考, '～') > 0 Then
                  
                    If Nvl(Zlval(v_检验结果), 0) < To_Number(Substr(v_危急参考, 1, Instr(v_危急参考, '～') - 1)) Then
                      v_结果标志 := 5;
                    End If;
                  
                    If Nvl(Zlval(v_检验结果), 0) > To_Number(Substr(v_危急参考, 1, Instr(v_危急参考, '～') - 1)) Then
                      v_结果标志 := 6;
                    End If;
                  
                  End If;
                End If;
              End If;
            End If;
          End If;
        End If;
        v_检验结果 := v_临时结果;
        Update 检验普通结果
        Set 检验结果 = v_检验结果, 结果标志 = Decode(v_多参考, 0, v_结果标志, 1), 修改者 = Decode(原始结果, Null, Null, v_人员姓名),
            修改时间 = Decode(原始结果, Null, Null, Sysdate), 原始结果 = Decode(原始结果, Null, v_检验结果, 原始结果),
            原始记录时间 = Decode(原始结果, Null, Sysdate, 原始记录时间), 记录者 = Decode(原始结果, Null, v_人员姓名, 记录者), 仪器id = 仪器id_In,
            结果参考 = Decode(v_多参考, 0, v_参考值, v_参考值1), Od = v_Od, Cutoff = v_Cutoff, Sco = v_Sco, 酶标板id = 酶标板id_In
        Where 检验标本id = 检验标本id_In And 检验项目id = v_项目id And 记录类型 = v_记录类型;
      
        If Sql%Rowcount = 0 Then
          Insert Into 检验普通结果
            (ID, 检验标本id, 检验项目id, 检验结果, 结果标志, 记录类型, 原始结果, 原始记录时间, 记录者, 仪器id, 结果参考, Od, Cutoff, Sco, 酶标板id)
          Values
            (检验普通结果_Id.Nextval, 检验标本id_In, v_项目id, v_检验结果, Decode(v_多参考, 0, v_结果标志, 1), 0, v_检验结果, Sysdate, v_人员姓名,
             仪器id_In, Decode(v_多参考, 0, v_参考值, v_参考值1), v_Od, v_Cutoff, v_Sco, 酶标板id_In);
        End If;
      
        Update 检验流水线指标
        Set 仪器是否审核 = v_仪器审核, 审核内容 = v_仪器审核内容
        Where 标本id = 检验标本id_In And 项目id = v_项目id;
      
        If Sql%Rowcount = 0 Then
          Insert Into 检验流水线指标
            (ID, 标本id, 项目id, 仪器是否审核, 审核内容)
          Values
            (检验流水线指标_Id.Nextval, 检验标本id_In, v_项目id, v_仪器审核, v_仪器审核内容);
        End If;
      
        Select Count(*) Into v_Count From 检验项目分布 Where 标本id = 检验标本id_In And 项目id + 0 = v_项目id;
        If v_Count = 0 Then
          Insert Into 检验项目分布
            (ID, 标本id, 项目id, 医嘱id, 范围)
          Values
            (检验项目分布_Id.Nextval, 检验标本id_In, v_项目id, Decode(v_医嘱id, 0, Null, v_医嘱id), 1);
        End If;
      Else
        --处理微生物
        v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
        If Instr(v_Currrec, '<Split>') > 0 Then
          v_审核字串 := Substr(v_Currrec, Instr(v_Currrec, '<Split>') + 7);
          v_Currrec  := Substr(v_Currrec, 1, Instr(v_Currrec, '<Split>') - 1);
        End If;
        v_Temp     := v_Currrec;
        v_项目id   := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '^') - 1));
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
        v_检验结果 := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
        v_药敏方法 := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
        v_药敏结果 := v_Temp;
        If v_审核字串 Is Not Null Then
          If Instr(v_审核字串, '^') > 0 Then
            v_仪器审核     := Substr(v_审核字串, 1, Instr(v_审核字串, '^') - 1);
            v_审核字串     := Substr(v_审核字串, Instr(v_审核字串, '^') + 1);
            v_仪器审核内容 := v_审核字串;
          Else
            v_仪器审核     := v_审核字串;
            v_仪器审核内容 := Null;
          End If;
        End If;
      
        Begin
          Select Distinct ID
          Into v_细菌id
          From 检验细菌 A, 仪器细菌对照 B
          Where a.Id = b.细菌id(+) And (a.中文名 = 标本类型_In Or a.英文名 = 标本类型_In Or a.简码 = 标本类型_In Or b.通道编码 = 标本类型_In);
        Exception
          When Others Then
            Return;
        End;
        If Sql%Rowcount > 0 Then
          Update 检验普通结果
          Set 修改者 = Decode(原始结果, Null, Null, v_人员姓名), 修改时间 = Decode(原始结果, Null, Null, Sysdate),
              记录者 = Decode(原始结果, Null, v_人员姓名, 记录者), 仪器id = 仪器id_In, 记录类型 = v_记录类型
          Where 检验标本id = 检验标本id_In And 细菌id = v_细菌id;
        
          If Sql%Rowcount = 0 Then
            Select 检验普通结果_Id.Nextval Into v_检验普通结果id From Dual;
            Insert Into 检验普通结果
              (ID, 检验标本id, 细菌id, 原始记录时间, 记录者, 仪器id, 记录类型)
            Values
              (v_检验普通结果id, 检验标本id_In, v_细菌id, Sysdate, v_人员姓名, 仪器id_In, v_记录类型);
          Else
            Select ID Into v_检验普通结果id From 检验普通结果 Where 检验标本id = 检验标本id_In And 细菌id = v_细菌id;
          End If;
          --------------暂不处理微生物流水线问题-----------------------------------------------------------       
          --         Update 检验流水线指标
          --          Set 仪器是否审核 = v_仪器审核, 审核内容 = v_仪器审核内容
          --          Where 标本id = 检验标本id_In And 项目id = v_项目id;
          --       
          --          If Sql%Rowcount = 0 Then
          --            Insert Into 检验流水线指标
          --              (ID, 标本id, 项目id, 仪器是否审核, 审核内容)
          --            Values
          --             (检验流水线指标_Id.Nextval, 检验标本id_In, v_项目id, v_仪器审核, v_仪器审核内容);
          --          End If;
          --------------------------------------------------------------------------------------------------       
          Select Count(*) Into v_Count From 检验项目分布 Where 标本id = 检验标本id_In And 项目id + 0 = v_细菌id;
          If v_Count = 0 Then
            Insert Into 检验项目分布
              (ID, 标本id, 细菌id, 医嘱id, 范围)
            Values
              (检验项目分布_Id.Nextval, 检验标本id_In, v_细菌id, Decode(v_医嘱id, 0, Null, v_医嘱id), 1);
          End If;
          If Nvl(v_项目id, 0) <> 0 Then
            Update 检验药敏结果
            Set 修改者 = v_人员姓名, 修改时间 = Sysdate, 结果 = v_药敏结果, 结果类型 = v_检验结果, 仪器id = 仪器id_In, 药敏方法 = v_药敏方法
            Where 细菌结果id = v_检验普通结果id And 抗生素id = v_项目id;
          
            If Sql%Rowcount = 0 Then
              Insert Into 检验药敏结果
                (细菌结果id, 抗生素id, 修改者, 修改时间, 结果, 结果类型, 记录类型, 仪器id, 药敏方法)
              Values
                (v_检验普通结果id, v_项目id, v_人员姓名, Sysdate, v_药敏结果, v_检验结果, 0, 仪器id_In, v_药敏方法);
            End If;
          End If;
        End If;
      End If;
      --v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');;
      v_Records := Substr(v_Records, Instr(v_Records, '|') + 1);
    End Loop;
    If 检验指标_In Is Not Null Then
      Update 检验标本记录 Set 检验时间 = Sysdate Where ID = 检验标本id_In;
    End If;
    If Nvl(微生物_In, 0) = 0 Then
      Select Count(*)
      Into v_仪器审核完成
      From 检验流水线指标
      Where 标本id = 检验标本id_In And Nvl(仪器是否审核, 0) = 0;
      If v_仪器审核完成 = 0 Then
        Update 检验流水线标本 Set 仪器是否审核 = 1 Where 标本id = 检验标本id_In;
        If Sql%Rowcount = 0 Then
          Insert Into 检验流水线标本 (ID, 标本id, 仪器是否审核) Values (检验流水线指标_Id.Nextval, 检验标本id_In, 1);
        End If;
      Else
        Update 检验流水线标本 Set 仪器是否审核 = 0 Where 标本id = 检验标本id_In;
        If Sql%Rowcount = 0 Then
          Insert Into 检验流水线标本 (ID, 标本id, 仪器是否审核) Values (检验流水线指标_Id.Nextval, 检验标本id_In, 0);
        End If;
      End If;
    End If;
  End If;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_检验普通结果_Batchupdate;
/


--89666:许华峰,2015-12-03,显示报告创建时间和审核时间
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
              创建人 As CreateUser,最后审核时间 As ExamineyDate,最后审核人 As ExamineyUser,Decode(Nvl(结果阳性,0),1,'阳性','') As RESULTPOSITIVE,
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
    Update 影像报告片段清单 Set 适应条件 = 适应条件_In Where ID = Hextoraw(ID_In) And 节点类型 != 0;
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
    Update 影像报告片段清单 Set 适应条件 = 适应条件_In Where 上级ID = Hextoraw(上级ID_In) And 节点类型 != 0;
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

--89419:张德婷,2015-01-05,出院病人不收配置费
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
  n_row      number(10);
  n_Out      number(1);

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_Out:=Nvl(zl_GetSysParameter('出院病人不收配置费', 1345), 0);

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
      Select Max(NO) Into v_No From 输液配药附费 Where 配药id = v_Tansid;
      if v_No is not null then
        select count(no) into n_row from 住院费用记录 where NO=v_No and 序号=1 and 记录状态=1;
        if n_row<>0 then
           Zl_住院记帐记录_Delete(v_No, 1, v_Usercode, Zl_Username);
        end if;
      end if;
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

--89419:张德婷,2015-01-05,出院病人不收配置费
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
  v_收费项目id varchar2(200);
  v_info    varchar2(200);
  v_id varchar2(20);
  n_count number(18);
  n_Out number(10);
  n_OutNum number(10);
  n_打包状态 number(1);
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
  n_Out:=Nvl(zl_GetSysParameter('出院病人不收配置费', 1345), 0);

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
      Select 操作状态,执行时间,nvl(是否打包,0)  Into n_操作状态,d_执行时间,n_打包状态 From 输液配药记录 Where ID = v_Tansid;

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

    if n_打包状态=0 then
      n_count:=0;
      Select Nextno(14) Into v_No From Dual;
      For r_Bill In c_Bill Loop
        Select count(病人id) into n_OutNum From 病案主页 where 主页ID=r_Bill.主页id And 病人ID=r_Bill.病人id  And (Nvl(状态,0)=3 Or 出院日期 Is Not NULL);

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

        if n_row=0 and (n_OutNum=0 or n_out=0) then
          For r_Item In (Select a.Id 收费细目id, a.类别 收费类别, a.计算单位, a.加班加价 加班标志, d.Id 收入项目id, d.收据费目, b.现价
                         From 收费项目目录 A, 收费价目 B, 收入项目 D
                         Where a.Id = b.收费细目id And b.收入项目id = d.Id And a.id=n_项目id and
                               b.执行日期 <= Sysdate And
                               (b.终止日期 >= Sysdate Or b.终止日期 Is Null)) Loop
            if n_count=0 then
              Insert Into 输液配药附费 (配药id, NO,病人id) Values (v_Tansid, v_No, r_Bill.病人id);
            end if;

            n_count:=n_count+1;
            Zl_住院记帐记录_Insert(v_No, n_count, r_Bill.病人id, r_Bill.主页id, r_Bill.标识号, r_Bill.姓名, r_Bill.性别, r_Bill.年龄, r_Bill.床号,
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
    end if;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_配药;
/

--91083:张德婷,2015-12-02,输液配置中心不配置药品正常使用
CREATE OR REPLACE Procedure Zl_输液配药记录_核查
(
  部门id_In   In 输液配药记录.部门id%Type,
  医嘱id_In   In Varchar2, --输液医嘱给药途径对应的医嘱ID:医嘱ID1,医嘱ID2...
  发送号_In   In 病人医嘱发送.发送号%Type,
  核查人_In   In 输液配药状态.操作人员%Type,
  核查时间_In In 输液配药状态.操作时间%Type
) Is
  v_Count    Number;
  v_序号     Number;
  v_执行时间 Date;

  v_相关id      Number;
  v_New相关id   Number;
  v_Old相关id   Number;
  v_发送号      Number;
  v_Tmp         Varchar2(200);
  I             Number;
  v_配药id      Number;
  v_批次        Number;
  v_Maxno       Varchar2(4000);
  v_Lableno     Varchar2(200);
  v_Maxbatch    Number;
  v_Curdose     Number;
  v_Sumdose     Number;
  v_Drugcount   Number;
  v_Currdate    Date;
  n_Needcheck   Number;
  n_Lngid       药品收发记录.Id%Type;
  n_Count       Number(3);
  n_单据        药品收发记录.单据%Type;
  v_No          药品收发记录.No%Type;
  n_发送次数    Number(5);
  n_病人id      病人信息.病人id%Type := 0;
  b_Change      Boolean;
  n_Sum         Number(8);
  n_调整批次    Number(1);
  n_Cur         Number(5);
  v_上次发送号  病人医嘱发送.发送号%Type;
  v_医嘱ids     Varchar2(4000);
  v_Tansid      Varchar2(12);
  v_当前病人    Varchar2(20);
  n_Num         Number(8);
  d_Old执行时间 Date;
  n_是否打包    Number(1);
  n_打包        Number(1);
  n_摆药单      Number(2);
  --控制参数
  v_医嘱类型       Number;
  v_输液总量       Number;
  v_大输液剂型     Varchar2(2000);
  v_大输液给药途径 Varchar2(2000);
  v_来源科室       Varchar2(4000);
  v_Continue       Number := 1;
  v_Nodosage       Number := 0;
  v_保持上次批次   Number := 0;
  d_手工打包时间   Date;
  n_Tpn处置方式    Number := 0;
  Err_Item Exception;

  Cursor c_医嘱记录 Is
    Select /*+rule */
    Distinct e.医嘱id As 相关id, e.发送号, b.频率间隔, b.间隔单位, b.执行时间方案, a.姓名, a.性别, a.年龄, a.住院号, a.当前床号 As 床号, a.当前病区id As 病人病区id,
             a.当前科室id As 病人科室id, e.首次时间, e.末次时间, b.开始执行时间, Nvl(e.发送数次, 0) As 次数, e.发送时间, Decode(b.医嘱期效, 0, 1, 2) As 医嘱类型,
             b.诊疗项目id As 给药途径, b.病人id, Nvl(c.执行标记, 0) As 是否tpn
    From 病人医嘱发送 E, 病人医嘱记录 B, 病人信息 A, 诊疗项目目录 C, Table(f_Num2list(医嘱id_In)) D
    Where e.医嘱id = b.Id And b.病人id = a.病人id And c.类别 = 'E' And c.操作类型 = '2' And c.执行分类 = 1 And b.诊疗项目id = c.Id And
          e.医嘱id = d.Column_Value And e.发送号 = 发送号_In
    Order By b.病人id, e.医嘱id, e.发送号;

  Cursor c_单个医嘱记录 Is
    Select /*+rule */
    Distinct e.医嘱id As 相关id, e.发送号, b.频率间隔, b.间隔单位, b.执行时间方案, a.姓名, a.性别, a.年龄, a.住院号, a.当前床号 As 床号, a.当前病区id As 病人病区id,
             a.当前科室id As 病人科室id, e.首次时间, e.末次时间, b.开始执行时间, Nvl(e.发送数次, 0) As 次数, e.发送时间, Decode(b.医嘱期效, 0, 1, 2) As 医嘱类型,
             b.诊疗项目id As 给药途径, b.病人id
    From 病人医嘱发送 E, 病人医嘱记录 B, 病人信息 A, 诊疗项目目录 C, Table(f_Num2list(医嘱id_In)) D
    Where e.医嘱id = b.Id And b.病人id = a.病人id And c.类别 = 'E' And c.操作类型 = '2' And c.执行分类 = 1 And b.诊疗项目id = c.Id And
          e.医嘱id = d.Column_Value And e.发送号 = 发送号_In And b.病人id = n_病人id
    Order By e.医嘱id, e.发送号;

  Cursor c_收发记录 Is
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, 输液药品属性 E
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.相关id = v_相关id And b.发送号 = 发送号_In And c.序号 < 1000
    Order By c.No, c.序号;

  v_医嘱记录     c_医嘱记录%RowType;
  v_收发记录     c_收发记录%RowType;
  v_单个医嘱记录 c_单个医嘱记录%RowType;
  Function Zl_Getpivaworkbatch
  (
    执行时间_In   In Date,
    配置中心id_In In 输液配药记录.部门id%Type
  ) Return Number As
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_配药批次 Is
      Select 批次, 配药时间, 给药时间, 打包
      From 配药工作批次
      Where 启用 = 1 And 配置中心id = 配置中心id_In
      Order By 批次;

    v_配药批次 c_配药批次%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');

    Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次 Where 启用 = 1 And 配置中心id = 配置中心id_In;

    For v_配药批次 In c_配药批次 Loop
      v_Batch     := 0;
      v_Starttime := To_Date(Substr(v_配药批次.给药时间, 1, Instr(v_配药批次.给药时间, '-') - 1), 'hh24:mi');
      v_Endtime   := To_Date(Substr(v_配药批次.给药时间, Instr(v_配药批次.给药时间, '-') + 1), 'hh24:mi');

      If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
        v_Batch := v_配药批次.批次;
        n_打包  := v_配药批次.打包;
        Exit When v_Batch > 0;
      End If;
    End Loop;

    If v_Batch = 0 Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;
Begin
  n_Count          := 0;
  v_医嘱类型       := Zl_To_Number(Nvl(zl_GetSysParameter('医嘱类型', 1345), 1));
  v_输液总量       := Zl_To_Number(Nvl(zl_GetSysParameter('同批次输液总量', 1345), 0));
  v_大输液剂型     := Nvl(zl_GetSysParameter('大输液药品剂型', 1345), '');
  v_大输液给药途径 := Nvl(zl_GetSysParameter('输液给药途径', 1345), '');
  v_来源科室       := Nvl(zl_GetSysParameter('来源病区', 1345), '');
  v_保持上次批次   := Zl_To_Number(Nvl(zl_GetSysParameter('保持上次批次', 1345), 0));

  v_医嘱ids  := 医嘱id_In;
  v_当前病人 := '';
  v_New相关id:=0;

  Select Trunc(Sysdate) Into v_Currdate From Dual;

  Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次 Where 启用 = 1 And 配置中心id = 部门id_In;

  --检查当前病人的医嘱是否有今天需要执行的输液单是锁定状态的
  If Instr(v_医嘱ids, ',') = 0 Then
    v_Tansid := v_医嘱ids;
  Else
    v_Tansid := Substr(v_Tmp, 1, Instr(v_医嘱ids, ',') - 1);
  End If;

  Select Count(ID)
  Into n_Num
  From 输液配药记录
  Where 是否锁定 = 1 And 执行时间 Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
        医嘱id In
        (Select 相关id
         From 病人医嘱记录
         Where 病人id = (Select 病人id From 病人医嘱记录 Where 相关id = v_Tansid And Rownum < 2) And (诊疗类别 = '5' Or 诊疗类别 = '6')) And
        Rownum < 2;

  If n_Num > 0 Then
    Select 姓名
    Into v_当前病人
    From 输液配药记录
    Where 是否锁定 = 1 And 执行时间 Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
          医嘱id In
          (Select 相关id
           From 病人医嘱记录
           Where 病人id = (Select 病人id From 病人医嘱记录 Where 相关id = v_Tansid And Rownum < 2) And (诊疗类别 = '5' Or 诊疗类别 = '6')) And
          Rownum < 2;
    Raise Err_Item;
  End If;

  For v_医嘱记录 In c_医嘱记录 Loop
    v_Continue := 1;

    Select Count(1) into v_Continue
    From 病人医嘱记录 A, 输液不配置药品 B,住院费用记录 C
    Where c.收费细目id = b.药品id and A.id=C.医嘱序号 And a.相关id = v_医嘱记录.相关id and C.记录状态=1;
    If v_Continue = 0 Then
      v_Continue := 1;
    Else
      v_Continue := 0;
    End If;

    --参数控制产生输液单
    If (v_医嘱类型 = 1 And v_医嘱记录.医嘱类型 <> 1) Or (v_医嘱类型 = 2 And v_医嘱记录.医嘱类型 <> 2) Then
      v_Continue := 0;
    End If;

    If Not v_大输液给药途径 Is Null Then
      If Instr(',' || v_大输液给药途径 || ',', ',' || v_医嘱记录.给药途径 || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;

    If Not v_来源科室 Is Null Then
      If Instr(',' || v_来源科室 || ',', ',' || v_医嘱记录.病人病区id || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;

    If v_医嘱记录.是否tpn = 2 Then
      v_Continue := 1;
    end if;

    If v_Continue = 1 Then
      v_Old相关id := v_New相关id;
      v_相关id    := v_医嘱记录.相关id;
      v_New相关id := v_相关id;
      v_发送号    := v_医嘱记录.发送号;
      v_序号      := 0;

      If v_Continue = 1 Then
        Select Count(医嘱id)
        Into v_Count
        From 医嘱执行时间
        Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号;

        v_Nodosage := 0;

        For I In 1 .. v_Count Loop
          Select 输液配药记录_Id.Nextval Into v_配药id From Dual;
          v_序号 := v_序号 + 1;

          If I > 1 Then
            --从医嘱执行时间表中取医嘱的执行时间
            Select 要求时间
            Into v_执行时间
            From 医嘱执行时间
            Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And 要求时间 > v_执行时间 And Rownum = 1
            Order By 要求时间;
          Else
            Select Min(要求时间)
            Into v_执行时间
            From 医嘱执行时间
            Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And Rownum = 1
            Order By 要求时间;
          End If;

          v_批次 := 0;

          If d_Old执行时间 <> Trunc(v_执行时间) Or d_Old执行时间 Is Null Then
            b_Change      := True;
            d_Old执行时间 := v_执行时间;

            Select /*+ rule*/
             Count(a.要求时间)
            Into n_Cur
            From 医嘱执行时间 A
            Where a.医嘱id In (Select ID
                             From 病人医嘱记录
                             Where 病人id = (Select 病人id From 病人医嘱记录 Where ID = v_医嘱记录.相关id And Rownum < 2)) And
                  a.要求时间 Between Trunc(v_执行时间) And Trunc(v_执行时间+1) - 1 / 24 / 60 / 60;

            Select Count(a.要求时间)
            Into n_Sum
            From 医嘱执行时间 A
            Where a.医嘱id In (Select ID
                             From 病人医嘱记录
                             Where 病人id = (Select 病人id From 病人医嘱记录 Where ID = v_医嘱记录.相关id And Rownum < 2)) And
                  a.要求时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;

            Select Count(Distinct a.摆药单号)
            Into n_摆药单
            From 输液配药记录 A
            Where a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = v_医嘱记录.病人id And 相关id Is Null) And
                  a.执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;

            If n_Cur <> n_Sum Or  n_摆药单 > 1 Then
              Update 输液配药记录
              Set 是否调整批次 = 1
              Where 医嘱id In (Select ID
                             From 病人医嘱记录
                             Where 病人id = (Select 病人id From 病人医嘱记录 Where ID = v_医嘱记录.相关id And Rownum < 2)) And
                    执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
              b_Change := False;

              For v_输液记录 In (Select ID, 执行时间
                             From 输液配药记录
                             Where 医嘱id In
                                   (Select ID
                                    From 病人医嘱记录
                                    Where 病人id =
                                          (Select 病人id From 病人医嘱记录 Where ID = v_医嘱记录.相关id And Rownum < 2)) And
                                   执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间+1) - 1 / 24 / 60 / 60 And 操作状态 < 2) Loop
                v_批次 := Zl_Getpivaworkbatch(v_输液记录.执行时间, 部门id_In);
                Update 输液配药记录 Set 配药批次 = v_批次 Where ID = v_输液记录.Id;
                v_批次 := 0;
              End Loop;
            End If;
          End If;

          If b_Change = True Then
            b_Change := True;
            n_病人id := v_医嘱记录.病人id;
            Select Count(ID)
            Into n_Sum
            From 输液配药记录
            Where 医嘱id = v_医嘱记录.相关id And 执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;
            If n_Sum = 0 Then
              Update 输液配药记录
              Set 是否调整批次 = 1
              Where 医嘱id In (Select ID
                             From 病人医嘱记录
                             Where 病人id = (Select 病人id From 病人医嘱记录 Where ID = v_医嘱记录.相关id And Rownum < 2)) And
                    执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
              b_Change := False;
            End If;

            If b_Change = True Then
              For v_单个医嘱记录 In c_单个医嘱记录 Loop
                --检查输液单是否调整到打包状态
                Select Count(ID)
                Into n_Sum
                From 输液配药记录
                Where 医嘱id = v_单个医嘱记录.相关id And 执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60 And
                      打包时间 Is Not Null;
                If n_Sum <> 0 Then
                  Update 输液配药记录
                  Set 是否调整批次 = 1
                  Where 医嘱id In
                        (Select ID
                         From 病人医嘱记录
                         Where 病人id = (Select 病人id From 病人医嘱记录 Where ID = v_单个医嘱记录.相关id And Rownum < 2)) And
                        执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                  b_Change := False;
                  Exit;
                End If;

                Select Count(医嘱id)
                Into n_Cur
                From 医嘱执行时间
                Where 医嘱id = v_单个医嘱记录.相关id And 要求时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60;
                Select Count(医嘱id)
                Into n_Sum
                From 医嘱执行时间
                Where 医嘱id = v_单个医嘱记录.相关id And 要求时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;
                If n_Sum <> n_Cur Then
                  Update 输液配药记录
                  Set 是否调整批次 = 1
                  Where 医嘱id In
                        (Select ID
                         From 病人医嘱记录
                         Where 病人id = (Select 病人id From 病人医嘱记录 Where ID = v_单个医嘱记录.相关id And Rownum < 2)) And
                        执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                  b_Change := False;
                  Exit;
                End If;
              End Loop;
            End If;
          End If;

          If v_保持上次批次 = 1 Or b_Change = True Then
            --取上次的批次
            Begin
              Select Distinct 配药批次
              Into v_批次
              From 输液配药记录 A
              Where 医嘱id = v_医嘱记录.相关id And
                    发送号 = (Select Distinct Max(发送号)
                           From 输液配药记录
                           Where 医嘱id = v_医嘱记录.相关id And 发送号 <> v_医嘱记录.发送号 And 执行时间<v_执行时间 and To_Char(执行时间, 'hh24:mi:ss') = To_Char(v_执行时间, 'hh24:mi:ss')) And
                    To_Char(a.执行时间, 'hh24:mi:ss') = To_Char(v_执行时间, 'hh24:mi:ss');
            Exception
              When Others Then
                v_批次 := 0;
            End;
          End If;

          If v_批次 = 0 Then
            v_批次 := Zl_Getpivaworkbatch(v_执行时间, 部门id_In);

            --同病人同批次总输液量控制，超过则分配到下个批次
            If v_输液总量 > 0 And Not v_大输液剂型 Is Null And v_批次 < v_Maxbatch Then
              Begin
                Select /*+rule */
                 Sum(单量) As 单量
                Into v_Curdose
                From (Select Distinct c.Id, c.单量
                       From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, 药品规格 E, 药品特性 F, Table(f_Str2list(v_大输液剂型)) G
                       Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id And
                             e.药名id = f.药名id And b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And
                             f.药品剂型 = g.Column_Value And a.相关id = v_相关id And b.发送号 = v_发送号);
              Exception
                When Others Then
                  v_Curdose := 0;
              End;

              Begin
                Select /*+rule */
                 Sum(单量) As 单量
                Into v_Sumdose
                From (Select Distinct a.Id, a.单量
                       From 药品收发记录 A, 病人医嘱记录 B, 输液配药记录 C, 输液配药内容 D, 药品规格 E, 药品特性 F, Table(f_Str2list(v_大输液剂型)) G
                       Where c.Id = d.记录id And a.Id = d.收发id And c.医嘱id = b.Id And a.药品id + 0 = e.药品id And
                             e.药名id = f.药名id And b.病人id + 0 = v_医嘱记录.病人id And f.药品剂型 = g.Column_Value And
                             c.执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And c.配药批次 = v_批次);
              Exception
                When Others Then
                  v_Sumdose := 0;
              End;

              If v_Sumdose > 0 And v_Sumdose + v_Curdose > v_输液总量 Then
                v_批次 := v_批次 + 1;
              End If;
            End If;
          End If;

          if v_Old相关id<>v_医嘱记录.相关id then
            Select Count(医嘱id)
            Into n_发送次数
            From 输液配药记录
            Where 医嘱id = v_医嘱记录.相关id
            Order By 执行时间;
          else
            n_发送次数:=n_发送次数+1;
          end if;

          If n_发送次数 > 99 Then
            n_发送次数 := Mod(n_发送次数, 99);
          End If;

          If Length(v_医嘱记录.相关id) > 9 Then
            If n_发送次数 < 10 Then
              Select '91' || Substr(To_Char(v_医嘱记录.相关id), Length(v_医嘱记录.相关id) - 8) || To_Char(v_医嘱记录.相关id) || '0' ||
                      To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr(To_Char(v_医嘱记录.相关id), Length(v_医嘱记录.相关id) - 8) || To_Char(v_医嘱记录.相关id) ||
                      To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            End If;
          Else
            If n_发送次数 < 10 Then
              Select '91' || Substr('000000000', Length(v_医嘱记录.相关id) + 1) || To_Char(v_医嘱记录.相关id) || '0' ||
                      To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr('000000000', Length(v_医嘱记录.相关id) + 1) || To_Char(v_医嘱记录.相关id) || To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            End If;
          End If;
          n_调整批次 := 0;
          If b_Change = False Then
            n_调整批次 := 1;
          End If;

          If v_批次 <> 0 Then
            Select Nvl(Max(打包), 0)
            Into n_打包
            From 配药工作批次
            Where 启用 = 1 And 配置中心id = 部门id_In And 批次 = v_批次;
          End If;

          If Trunc(v_执行时间) <= v_Currdate Or n_打包 <> 0 Then
            n_是否打包     := 1;
            d_手工打包时间 := Sysdate;
          Else
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;

           --如果是TPN：如果指定了要打包或配置，则不管其他条件如何都设置为打包或配置
          If v_医嘱记录.是否tpn = 2 Then
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;

          --产生配药记录
          Insert Into 输液配药记录
            (ID, 部门id, 序号, 姓名, 性别, 年龄, 住院号, 床号, 病人病区id, 病人科室id, 执行时间, 医嘱id, 发送号, 配药批次, 瓶签号, 是否调整批次, 是否打包, 打包时间, 操作状态,
             操作人员, 操作时间)
          Values
            (v_配药id, 部门id_In, v_序号, v_医嘱记录.姓名, v_医嘱记录.性别, v_医嘱记录.年龄, v_医嘱记录.住院号, v_医嘱记录.床号, v_医嘱记录.病人病区id,
             v_医嘱记录.病人科室id, v_执行时间, v_医嘱记录.相关id, v_医嘱记录.发送号, Decode(v_批次, 0, Null, v_批次), v_Maxno, n_调整批次, n_是否打包,
             d_手工打包时间, 1, 核查人_In, 核查时间_In);

          Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (v_配药id, 1, 核查人_In, 核查时间_In);

          --产生配药记录对应的药品记录
          For v_收发记录 In c_收发记录 Loop
            If v_收发记录.是否不予配置 = 1 Then
              v_Nodosage := 1;
            End If;

            n_Count := n_Count + 1;

            Select 药品收发记录_Id.Nextval Into n_Lngid From Dual;

            Insert Into 药品收发记录
              (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 付数, 填写数量, 实际数量,
               成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期, 灭菌效期,
               产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间)
              Select n_Lngid, 记录状态, 单据, NO, n_Count + 1000, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期,
                     效期, 付数, 填写数量 / v_Count, 实际数量 / v_Count, 成本价, 成本金额 / v_Count, 扣率, 零售价, 零售金额 / v_Count, 差价 / v_Count,
                     '复制', 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期, 灭菌效期, 产品合格证, 发药方式, 发药窗口,
                     领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间


        From 药品收发记录
              Where ID = v_收发记录.收发id;

            Insert Into 输液配药内容 (记录id, 收发id, 数量) Values (v_配药id, n_Lngid, v_收发记录.数量 / v_Count);
          End Loop;

        End Loop;

        For v_收发记录 In c_收发记录 Loop
          n_单据 := v_收发记录.单据;

          v_No := v_收发记录.No;
          Delete From 药品收发记录 Where ID = v_收发记录.收发id;
        End Loop;

        --如果存在“不予配置”属性的药品，也设置为打包
        If v_Nodosage = 1 Then
          Update 输液配药记录 Set 是否打包 = 1 Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号;
        End If;
      End If;
    End If;
  End Loop;

  For v_收发记录 In (Select ID From 药品收发记录 Where 序号 < 1000 And 单据 = n_单据 And NO = v_No) Loop
    n_Count := n_Count + 1;
    Update 药品收发记录 Set 序号 = n_Count + 1000, 摘要 = '复制' Where ID = v_收发记录.Id;
  End Loop;

  Update 药品收发记录
  Set 序号 = 序号 - 1000, 摘要 = '医嘱发送'
  Where 摘要 = '复制' And 序号 > 1000 And 单据 = n_单据 And NO = v_No;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]病人' || v_当前病人 || '在输液配置中心有被锁定的输液单，发送失败！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_核查;
/

--90943:刘尔旋,2015-12-02,支付宝预约病历费问题
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
    For c_费用 In (Select 1 As 顺序号, b.No, b.收据费目, b.结帐id, b.执行部门id, b.病人科室ID, b.开单人, b.收费类别, b.收入项目id, b.附加标志,
                        To_Char(b.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.价格父号, b.从属父号, b.序号, b.收费细目id, b.计算单位,
                        Max(m.名称) As 名称, Max(m.规格) As 规格, Sum(b.标准单价) As 单价, Avg(Nvl(b.付数, 1) * b.数次) As 数量,
                        Sum(b.应收金额) As 应收金额, Sum(b.实收金额) As 实收金额, Max(j.名称) As 开单科室, Max(q.名称) As 执行科室
                 From 门诊费用记录 B, 收费项目目录 M, 部门表 J, 部门表 Q
                 Where b.No = v_Nos And b.记录性质 = 4 And Nvl(b.费用状态, 0) = 0 And
                       b.记录状态 = 0 And b.收费细目id = m.Id And b.开单部门id = j.Id(+) And b.执行部门id = q.Id(+)
                 Group By b.No, b.收据费目, b.结帐id, b.执行部门id, b.病人科室ID, b.开单人, b.收入项目id, b.收费类别, b.登记时间, b.价格父号, b.从属父号, b.序号,
                          b.收费细目id, b.计算单位, b.附加标志
                 Order By 序号) Loop
      Zl_病人预约挂号记录_Update(c_费用.No, c_费用.序号, c_费用.价格父号, c_费用.从属父号, c_费用.收费类别, c_费用.收费细目id, c_费用.数量, c_费用.单价, c_费用.收入项目id,
                         c_费用.收据费目, c_费用.应收金额, c_费用.实收金额, c_费用.附加标志, Null, Null, Null, Null, c_费用.病人科室ID, c_费用.执行部门id);
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

--90943:刘尔旋,2015-11-30,挂号取计划ID问题
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
  Select Decode(To_Char(d_原始时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;
  Begin
    Select ID
    Into n_计划id
    From (Select ID
           From 挂号安排计划
           Where 号码 = v_号码 And d_原始时间 Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
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

--90512:马政,2015-11-17,上次采购价信息处理
Create Or Replace Procedure Zl_药品其他入库_Verify
(
  No_In     In 药品收发记录.No%Type := Null,
  审核人_In In 药品收发记录.审核人%Type := Null
) Is
  Err_Isverified Exception;
  Err_Isbatch Exception;
  v_Druginf Varchar2(50); --原不分批现在分批的药品信息

  Cursor c_药品收发记录 Is
    Select a.Id, a.实际数量, a.零售金额, a.零售价, a.差价, a.库房id, a.药品id, a.批次, a.成本价, a.批号, a.效期, a.产地, a.入出类别id, a.生产日期, a.批准文号,
           a.供药单位id, Nvl(b.是否变价, 0) As 时价
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.No = No_In And a.单据 = 4 And a.记录状态 = 1
    Order By a.药品id;
Begin
  Update 药品收发记录
  Set 审核人 = 审核人_In, 审核日期 = Sysdate
  Where NO = No_In And 单据 = 4 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  --主要针对原不分批现在分批的药品，不能对其审核
  Begin
    Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
    Into v_Druginf
    From 药品收发记录 A, 药品规格 B, 收费项目目录 I, 收费项目别名 N
    Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 4 And
          a.记录状态 = 1 And Nvl(a.批次, 0) = 0 And
          ((Nvl(b.药库分批, 0) = 1 And
          a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or Nvl(b.药房分批, 0) = 1) And
          Rownum = 1;
  Exception
    When Others Then
      v_Druginf := '';
  End;

  If v_Druginf Is Not Null Then
    Raise Err_Isbatch;
  End If;

  --原分批现不分批的药品,在审核时，要处理他
  Update 药品收发记录
  Set 批次 = 0
  Where ID In
        (Select ID
         From 药品收发记录 A, 药品规格 B
         Where b.药品id = a.药品id And a.No = No_In And a.单据 = 4 And a.记录状态 = 1 And Nvl(a.批次, 0) > 0 And
               (Nvl(b.药库分批, 0) = 0 Or
               (Nvl(b.药房分批, 0) = 0 And
               a.库房id In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室')))));

  For v_药品收发记录 In c_药品收发记录 Loop
    --更改药品库存表的相应数据
    Zl_药品库存_Update(v_药品收发记录.Id, 0);
    Update 药品规格
    Set 成本价 = v_药品收发记录.成本价, 上次售价 = Decode(v_药品收发记录.时价, 1, v_药品收发记录.零售价, Null), 上次供应商id = v_药品收发记录.供药单位id,
        上次批号 = v_药品收发记录.批号, 上次生产日期 = v_药品收发记录.生产日期, 上次产地 = v_药品收发记录.产地, 上次批准文号 = v_药品收发记录.批准文号
    Where 药品id = v_药品收发记录.药品id;
  End Loop;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品[' || v_Druginf || ']，不能审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品其他入库_Verify;
/

--92493:梁经伙,2016-01-08,疾病报告前提西新添加了字段 报告病种
Create Or Replace Procedure Zl_疾病报告前提_Append
(
  文件id_In   In 疾病报告前提.文件id%Type,
  疾病_In     In Varchar2 := Null, --按分号分隔的疾病id串
  诊断_In     In Varchar2 := Null, --按分号分隔的疾病id串
  报告病种_In In Varchar2 := Null
) Is
  v_Disease Varchar2(4000);
  n_疾病id  疾病报告前提.疾病id%Type;
  n_诊断id  疾病报告前提.诊断id%Type;
Begin

  Update 病历文件列表 Set 名称 = 名称 Where ID = 文件id_In;

  If Sql%RowCount = 0 Then
    Raise No_Data_Found;
  End If;

  If 疾病_In Is Not Null Then
    v_Disease := 疾病_In || ';';
    While v_Disease Is Not Null Loop
      n_疾病id  := To_Number(Substr(v_Disease, 1, Instr(v_Disease, ';') - 1));
      v_Disease := Substr(v_Disease, Instr(v_Disease, ';') + 1);
      Update 疾病报告前提 Set 文件id = 文件id, 报告病种 = 报告病种_In Where 文件id = 文件id_In And 疾病id = n_疾病id;
      If Sql%RowCount = 0 Then
        Insert Into 疾病报告前提 (文件id, 疾病id, 报告病种) Values (文件id_In, n_疾病id, 报告病种_In);
      End If;
    End Loop;
  End If;

  If 诊断_In Is Not Null Then
    v_Disease := 诊断_In || ';';
    While v_Disease Is Not Null Loop
      n_诊断id  := To_Number(Substr(v_Disease, 1, Instr(v_Disease, ';') - 1));
      v_Disease := Substr(v_Disease, Instr(v_Disease, ';') + 1);
      Update 疾病报告前提 Set 文件id = 文件id, 报告病种 = 报告病种_In Where 文件id = 文件id_In And 诊断id = n_诊断id;
      If Sql%RowCount = 0 Then
        Insert Into 疾病报告前提 (文件id, 诊断id, 报告病种) Values (文件id_In, n_诊断id, 报告病种_In);
      End If;
    End Loop;
  End If;
Exception
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]没有找到文件，可能已经被删除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病报告前提_Append;
/




Insert Into zlFilesUpgrade (文件类型,文件名,版本号,修改日期,所属系统,业务部件,安装路径,文件说明,强制覆盖,自动注册,加入日期,序号) select 1,'zl9Disease.dll','', Null ,'1','zl9Cisjob','[APPSOFT]\APPLY','疾病报告部件','0','1',sysdate,序号 from Dual a,(Select max(to_number(序号))+1 序号 from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(文件名)='ZL9DISEASE.DLL');
Insert Into zlFilesUpgrade (文件类型,文件名,版本号,修改日期,所属系统,业务部件,安装路径,文件说明,强制覆盖,自动注册,加入日期,序号) select 1,'zlDisReportCard.dll','', Null ,'1','zl9Cisjob','[APPSOFT]\APPLY','疾病报告设置部件','0','1',sysdate,序号 from Dual a,(Select max(to_number(序号))+1 序号 from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(文件名)='ZLDISREPORTCARD.DLL');
Insert Into zlFilesUpgrade (文件类型,文件名,版本号,修改日期,所属系统,业务部件,安装路径,文件说明,强制覆盖,自动注册,加入日期,序号) select 1,'zl9PacsImageCap.dll','', Null ,'1','zl9PacsWork','[APPSOFT]\APPLY','视频采集部件','0','1',sysdate,序号 from Dual a,(Select max(to_number(序号))+1 序号 from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(文件名)='ZL9PACSIMAGECAP.DLL');

---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.34.60' Where 编号=&n_System;
--部件版本号
Commit;