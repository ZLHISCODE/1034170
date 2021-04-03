--[连续升级]1
--[管理工具版本号]10.34.170
--本脚本支持从ZLHIS+ v10.34.160 升级到 v10.34.170
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--129869:蒋廷中,2019-01-11,新增病人用药清单表和病人用药配方表
Create Table 病人用药清单(
  ID  NUMBER(18),
  病人ID  NUMBER(18), 
  主页ID  NUMBER(5),  
  组号  NUMBER(18),
  用药来源  NUMBER(1),
  药品类别  VARCHAR2(1),
  用药内容  VARCHAR2(1000),
  诊疗项目ID  NUMBER(18),    
  收费细目ID  NUMBER(18),    
  天数  Number(16,5),
  开始时间  DATE,
  终止时间  DATE,
  登记时间  DATE,
  登记人  VARCHAR2(20),
  总给予量   NUMBER(16,5),
  单次用量  NUMBER(16,5),   
  执行频次  VARCHAR2(20),   
  频率次数  NUMBER(3),   
  频率间隔  NUMBER(3),    
  间隔单位  VARCHAR2(4),
  用法ID  NUMBER(18),
  煎法ID  NUMBER(18),
  备注  VARCHAR2(1000),
  待转出     NUMBER(3)
)TABLESPACE zl9CisRec;

Create Table 病人用药配方(
  配方ID  NUMBER(18),
  序号  NUMBER(3),  
  诊疗项目ID  NUMBER(18),    
  收费细目ID  NUMBER(18),  
  单量  NUMBER(16,5),    
  脚注  VARCHAR2(100),
  待转出     NUMBER(3)
)TABLESPACE zl9CisRec;

Create Sequence 病人用药清单_ID Start With 1;

Alter Table 病人用药清单 Add Constraint 病人用药清单_PK Primary Key (ID) Using Index Tablespace zl9Indexhis;

Alter Table 病人用药配方 Add Constraint 病人用药配方_PK Primary Key (配方id,序号) Using Index Tablespace zl9Indexhis;

Create Index 病人用药清单_IX_病人ID on 病人用药清单(病人ID, 主页ID) Tablespace zl9indexhis;

Create Index 病人用药清单_IX_收费细目ID on 病人用药清单(收费细目ID) Tablespace zl9indexhis;

Create Index 病人用药清单_IX_诊疗项目ID on 病人用药清单(诊疗项目ID) Tablespace zl9indexhis;

Create Index 病人用药清单_IX_开始时间 on 病人用药清单(开始时间) Tablespace zl9indexhis;

Create Index 病人用药清单_IX_待转出 on 病人用药清单(待转出) Tablespace zl9indexhis;

Create Index 病人用药配方_IX_收费细目ID on 病人用药配方(收费细目ID) Tablespace zl9indexhis;

Create Index 病人用药配方_IX_诊疗项目ID on 病人用药配方(诊疗项目ID) Tablespace zl9indexhis;

Create Index 病人用药配方_IX_待转出 on 病人用药配方(待转出) Tablespace zl9indexhis;

Alter Table 病人用药清单 Add Constraint 病人用药清单_FK_收费细目ID Foreign Key (收费细目ID) References 收费项目目录 (id);
Alter Table 病人用药清单 Add Constraint 病人用药清单_FK_诊疗项目ID Foreign Key (诊疗项目ID) References 诊疗项目目录 (id);
Alter Table 病人用药清单 Add Constraint 病人用药清单_FK_病人ID Foreign Key (病人ID, 主页ID) References 病案主页 (病人ID, 主页ID);
Alter Table 病人用药配方 Add Constraint 病人用药配方_FK_配方id Foreign Key (配方id) References 病人用药清单 (id) ON DELETE CASCADE;
Alter Table 病人用药配方 Add Constraint 病人用药配方_FK_收费细目ID Foreign Key (收费细目ID) References 收费项目目录 (id);
Alter Table 病人用药配方 Add Constraint 病人用药配方_FK_诊疗项目ID Foreign Key (诊疗项目ID) References 诊疗项目目录 (id);

Alter Table 病人用药清单 Add Constraint 病人用药清单_FK_用法ID Foreign Key (用法ID) References 诊疗项目目录 (id);
Alter Table 病人用药清单 Add Constraint 病人用药清单_FK_煎法ID Foreign Key (煎法ID) References 诊疗项目目录 (id);
Create Index 病人用药清单_IX_用法ID on 病人用药清单(用法ID) Tablespace zl9indexhis;
Create Index 病人用药清单_IX_煎法ID on 病人用药清单(煎法ID) Tablespace zl9indexhis;


--136111:冉俊明,2019-01-08,修正挂号合作单位与预约方式存在相同名称时临床出诊安排报错的问题
Alter Table 临床出诊变动明细 Drop Constraint 临床出诊变动明细_PK Cascade Drop Index;

Alter Table 临床出诊变动明细 Add Constraint 临床出诊变动明细_UQ_变动id Unique(变动id,变动性质,序号,类型,名称) Using Index Tablespace Zl9indexhis;

Alter Table 临床出诊变动明细 Modify 变动id Constraint 临床出诊变动明细_NN_变动id Not Null;

Alter Table 临床出诊变动明细 Modify 变动性质 Constraint 临床出诊变动明细_NN_变动性质 Not Null;

--128511:殷瑞,2018-10-19,修正药品部门发药实际留存发药数不一致的问题
alter table 药品留存计划 add 实际数量 number(18);

--101765:秦龙,2018-08-27,增加字段是否辅助用药
Alter table 药品特性 add 是否辅助用药 number(1);

--00000:刘硕,2018-08-23,上机人员表补充索引
Create Index 上机人员表_IX_人员ID On 上机人员表(人员id)   Tablespace zl9indexhis;
--128124:蒋敏,2018-07-13,疾病编码管理新增手术操作类型
Create Table 手术操作类型(
    编码 Number(1), 
    名称 Varchar2(20)
)TABLESPACE zl9BaseItem;

Alter Table 手术操作类型 Add Constraint 手术操作类型_PK Primary Key (编码) Using Index Tablespace zl9Indexhis;
Alter Table 手术操作类型 Add Constraint 手术操作类型_UQ_名称 Unique (名称) Using Index Tablespace zl9Indexhis;

Alter Table 疾病编码目录 Add 手术操作类型 Number(1);

-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--130397:廖思奇,2018-01-24,报到打印前选择格式
Insert Into zlParameters
(ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
Select Zlparameters_Id.Nextval, &n_System, 1290, 0, 1, 1, 1, 59, '报到打印前选择格式', Null, '0', 
'报到后自动打印前选择要打印的报表格式。'
From Dual;

Insert Into zlParameters
(ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
Select Zlparameters_Id.Nextval, &n_System, 1291, 0, 1, 1, 1, 61, '报到打印前选择格式', Null, '0', 
'报到后自动打印前选择要打印的报表格式。'
From Dual;

Insert Into zlParameters
(ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
Select Zlparameters_Id.Nextval, &n_System, 1294, 0, 1, 1, 1, 117, '报到打印前选择格式', Null, '0', 
'报到后自动打印前选择要打印的报表格式。'
From Dual;


--100871:殷瑞,2019-01-16,门诊处方审查合格自动发送处方
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 267, '门诊审方处方自动发送', Null, '0',
         '门诊审方合格确认后，启用自动发送处方功能。0-不启用,1-启用'
  From Dual;

--135633:胡俊勇,2019-01-14,分化程序和最高诊断依据
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
Select  zlParameters_ID.Nextval,&n_System,-Null,-Null,-Null,-Null,-Null,318,'出院肿瘤诊断填写分化程度和最高诊断依据','0','0',
'西医诊断中出院诊断中诊断编码为为肿瘤诊断时才填写分化程度和最高诊断依据,0-任意诊断均可填写；1-疾病编码为肿瘤诊断时才填写,适用于住院医生首页整理和病案整理' from dual;
  
 
--129869:蒋廷中,2019-01-11,新增用药清单用药配方的转出脚本
Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
Select &n_System,12,A.* From (
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0
Union All Select '病人用药配方',1,1,-NULL From Dual
Union All Select '病人用药清单',2,1,-NULL From Dual) A;


Insert Into zlBakTableindex(系统,表名,索引名)
Select &n_System,A.* From (
Select 表名,索引名 From zlBakTableindex Where 1 = 0
Union All Select '病人用药配方','病人用药配方_PK' From Dual
Union All Select '病人用药清单','病人用药清单_PK' From Dual
Union All Select 表名,索引名 From zlBakTableindex Where 1 = 0) A;

--129869:蒋廷中,2019-01-11,新增病人用药清单表和病人用药配方表
Insert into zlTables(系统,表名,表空间,分类) Values(100,'病人用药清单','zl9CisRec','B1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'病人用药配方','zl9CisRec','B1');

--115085:胡俊勇,2018-01-07,医嘱单打印参数拆分  
Insert Into zlParameters
(ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
Select Zlparameters_Id.Nextval, &n_System, 1254, 0, 0, 0, 0, 77, '长嘱单重整换页', '0', '1', 
'打印长期医嘱单时，如果启用了此参数，医嘱重整后打印重整后的医嘱会换页并在首行打印重整标记。'
From Dual;

Insert Into zlParameters
(ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
Select Zlparameters_Id.Nextval, &n_System, 1254, 0, 0, 0, 0, 78, '长嘱单转科换页', '0', '1', 
'打印长期医嘱单时，如果启用了此参数，下达转科医嘱后的医嘱打印时会换页。'
From Dual;

Insert Into zlParameters
(ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
Select Zlparameters_Id.Nextval, &n_System, 1254, 0, 0, 0, 0, 79, '长嘱单术后换页', '0', '1', 
'打印长期医嘱单时，如果启用了此参数，下术后医嘱后的医嘱打印时会换页。'
From Dual;


--115085:胡俊勇,2018-01-07,医嘱单打印参数拆分  
Declare
  v_Par Zlparameters.参数值%Type;
Begin
  Begin
    Select 参数值 Into v_Par From zlParameters Where 参数名 = '重整和术后医嘱换页打印' And 模块 = 1254;
    Update zlParameters
    Set 参数值 = v_Par
    Where 模块 = 1254 And 参数名 In ('长嘱单重整换页', '长嘱单转科换页', '长嘱单术后换页');
    Delete zlParameters Where 参数名 = '重整和术后医嘱换页打印' And 模块 = 1254;
    Update zlParameters
    Set 参数说明 = '必须启用参数：长嘱单转科换页，此参数才生效。打印医嘱单时，如果启用了此参数，并且启用了参数：长嘱单转科换页，则另起一页打印的首行打印“重开医嘱”字样。'
    Where 参数名 = '转科换页后在首行打印重开医嘱' And 模块 = 1254;
  Exception
    When Others Then
      Null;
  End;
End;
/

--79893:殷瑞,2018-09-26,处方发药新增参数
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1341, 1, 1, 0, 0, 53, '发药后检查卫材发放情况', Null, '0',
         '勾选该参数时在发药后检查当前病人在该药房是否有卫生材料的未发单据，有则提示进行发放'
  From Dual;

--119905:冉俊明,2018-09-12,病人收费管理按病人补打票据含挂号费
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1121, 1, 0, 0, 0, 115, '按病人补打票据含挂号费', Null, '0',
         '按病人补打票据时，是否缺省提取挂号费。0-不缺省提取挂号费，1-缺省提取挂号费'
  From Dual;

--129322:殷瑞,2018-08-01,部门发药新增参数控制服药时间和用药次数显示
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1342, 0, 1, 0, 0, 34, '加载服药时间和用药次数', Null, '0',
         '0-不启用;1-启用。控制界面是否显示服药时间和用药次数'
  From Dual;

--120645:秦龙,2018-07-18,增加本机参数“计算库存数量时方式”
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1330, 1, 1, 0, 0, 13, '计算库存数量时方式', '0', '0',
         '当选择0-采用实际数量时，根据统计库存的实际数量来计算库房库存数量，当选择1-采用可用数量时，根据统计库存的可用数量来计算库房库存数量，0-采用实际数量，1-采用可用数量'
  From Dual;

--128118:秦龙,2018-07-16,增加系统参数“已失效卫材禁止入库”
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, 0, 0, 0, 0, 309, '已失效卫材禁止入库', '0', '0',
         '如果启用了该参数，在卫材外购入库，卫材其他入库，保存或审核单据时，则禁止已失效卫材的单据保存或审核。'
  From Dual;

--128590:廖思奇,2018-07-13,修正一个ID错误
Update 快捷功能信息 Set 菜单ID = 8127 Where 菜单说明 = '影像观片' And 项目 = 'ZL9PACSWORK' And 模块号 = 1290;

--128124:蒋敏,2018-07-13,疾病编码管理新增手术操作类型
Insert Into 手术操作类型(编码, 名称) Values (1, '治疗性操作');
Insert Into 手术操作类型(编码, 名称) Values (2, '诊断性操作');
Insert Into 手术操作类型(编码, 名称) Values (3, '介入治疗');
Insert Into 手术操作类型(编码, 名称) Values (4, '手术');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'手术操作类型','ZL9BASEITEM','A1');
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( 100,'手术操作类型',0,'用于保存手术的操作类型(手术、治疗、操作)','医疗工作' );

--127396:秦龙,2018-07-13,增加系统参数“已失效药品禁止入库”
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, 0, 0, 0, 0, 306, '已失效药品禁止入库', '0', '0',
         '如果启用了该参数，在药品外购入库，药品其他入库，保存或审核单据时，则禁止已失效药品的单据保存或审核。'
  From Dual;




-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--129869:蒋廷中,2019-01-11,新增病人用药清单表和病人用药配方表
Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select &n_System, 1253, '病人用药清单', 39, '有权限可对病人进行用药清单登记和查阅。', 1
  From Dual;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1253,'病人用药清单',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select '病人用药清单','SELECT' From Dual
Union All Select '病人用药配方','SELECT' From Dual
Union All Select 'Zl_病人用药清单_Insert','EXECUTE' From Dual
Union All Select 'Zl_病人用药清单_Update','EXECUTE' From Dual
Union All Select 'Zl_病人用药清单_Delete','EXECUTE' From Dual
Union All Select 'Zl_病人用药配方_Insert','EXECUTE' From Dual
Union All Select 'Zl_病人用药配方_Delete','EXECUTE' From Dual) A;

--136392:秦龙,2019-01-09,部门属性中增加权限
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1001,'增删改',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select '药品规格','SELECT' From Dual
Union All Select '材料特性','SELECT' From Dual
Union All Select '收费项目目录','SELECT' From Dual) A;

--126385:冉俊明,2018-10-29,LED语音报价权限控制调整
Update zlProgFuncs
Set 说明 = 'LED设备使用的权限。有该权限时，允许使用LED语音报价器'
Where 系统 = &n_System And 序号 = 1121 And 功能 = 'LED与语音';

--132712:余伟节,2018-10-25,刷身份证时姓名不一致时调整基本信息
 Insert Into zlProgPrivs(系统, 序号, 功能, 所有者, 对象, 权限)
   Select &n_System, 1101, '基本', User, 'Zl_病人信息_基本信息调整', 'EXECUTE' From Dual;

 Insert Into zlProgPrivs(系统, 序号, 功能, 所有者, 对象, 权限)
   Select &n_System, 1131, '基本', User, 'Zl_病人信息_基本信息调整', 'EXECUTE' From Dual;

--132528:涂建华,2018-10-10,身份证唯一识别处理
Insert Into zlProgPrivs
(系统, 序号, 功能, 所有者, 对象, 权限)
Select &n_System, 1290, '基本', User, a.* From (
Select 对象, 权限 From zlProgPrivs Where 1 = 0 
Union All Select 'Zl_Custom_Patiids_Get', 'EXECUTE' From Dual) A;

Insert Into zlProgPrivs
(系统, 序号, 功能, 所有者, 对象, 权限)
Select &n_System, 1291, '基本', User, a.* From (
Select 对象, 权限 From zlProgPrivs Where 1 = 0 
Union All Select 'Zl_Custom_Patiids_Get', 'EXECUTE' From Dual) A;

Insert Into zlProgPrivs
(系统, 序号, 功能, 所有者, 对象, 权限)
Select &n_System, 1294, '基本', User, a.* From (
Select 对象, 权限 From zlProgPrivs Where 1 = 0 
Union All Select 'Zl_Custom_Patiids_Get', 'EXECUTE' From Dual) A;

--131276:焦博,2018-09-27,根据自定义函数zl_Custom_PatiIDs_Get来获取病人ID
Insert Into zlProgPrivs
(系统, 序号, 功能, 所有者, 对象, 权限)
Select &n_System, 1107, '基本', User, a.* From (
Select 对象, 权限 From zlProgPrivs Where 1 = 0 
Union All Select 'Zl_Custom_Patiids_Get', 'EXECUTE' From Dual) A;

Insert Into zlProgPrivs
(系统, 序号, 功能, 所有者, 对象, 权限)
Select &n_System, 1111, '基本', User, a.* From (
Select 对象, 权限 From zlProgPrivs Where 1 = 0 
Union All Select 'Zl_Custom_Patiids_Get', 'EXECUTE' From Dual) A;

Insert Into zlProgPrivs
(系统, 序号, 功能, 所有者, 对象, 权限)
Select &n_System, 1113, '基本', User, a.* From (
Select 对象, 权限 From zlProgPrivs Where 1 = 0 
Union All Select 'Zl_Custom_Patiids_Get', 'EXECUTE' From Dual) A;

--131549:冉俊明,2018-09-26,号源由固定排班方式转为月排班方式，无法添加到已发布月出诊表中
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1114, '调整安排', User, 'ZL_临床出诊记录_DELETE', 'EXECUTE' From Dual;

--131413:余伟节,2018-09-25,根据自定义函数zl_Custom_PatiIDs_Get来获取病人ID
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_system,1101,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'zl_Custom_PatiIDs_Get','EXECUTE' From Dual) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_system,1131,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'zl_Custom_PatiIDs_Get','EXECUTE' From Dual) A;

--128989:焦博,2018-07-19,针对门诊分诊管理模块的基本权限,增加对门诊诊室适用科室表的查询权限
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1113,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select '门诊诊室适用科室','SELECT' From Dual) A;

--128124:蒋敏,2018-07-13,疾病编码管理新增手术操作类型
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select 100,1013,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select '手术操作类型','SELECT' From Dual) A;

--101944:余伟节,2019-01-17,担保信息权限说明修正
Update zlProgFuncs
Set 说明 = '允许操作病人担保信息的权限。有该权限时，在病人入院管理、入院登记时允许操作病人担保信息;否则,不允许操作。'
Where 系统 = &n_System And 序号 = 1131 And 功能 = '担保信息';

-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--137311:刘涛,2019-01-22,出库原料只提取可用库存大于0
Create Or Replace Procedure Zl_自制材料入库_Insert
(
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  对方部门id_In In 药品收发记录.对方部门id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  实际数量_In   In 药品收发记录.实际数量%Type,
  零售价_In     In 药品收发记录.零售价%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  填制日期_In   In 药品收发记录.填制日期%Type := Null,
  记录数_In     In Integer := 0
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  n_Id         药品收发记录.Id%Type; --收发ID
  n_出库类别id 药品收发记录.入出类别id%Type; --入出类别ID
  n_入的类别id 药品收发记录.入出类别id%Type; --入出类别ID
  n_剩余数量   药品库存.可用数量%Type;
  n_当前数量   药品收发记录.实际数量%Type;
  n_售价       药品收发记录.零售价%Type;
  n_现价       收费价目.现价%Type;
  n_零售金额   药品收发记录.零售金额%Type;
  n_成本价     药品收发记录.成本价%Type;
  n_成本金额   药品收发记录.成本金额%Type;
  n_总出库成本 药品收发记录.成本金额%Type;
  n_差价       药品收发记录.差价%Type;
  n_出序号     药品收发记录.序号%Type;
  n_最后出库id 药品收发记录.Id%Type;
  n_实价卫材   收费项目目录.是否变价%Type;
  n_库房分批   Integer; --是否分批核算   1:分批;0：不分批
  n_在用分批   Integer; --在用分批
  n_批次       药品收发记录.批次%Type := Null; --批次
  v_负成本计算 Zlparameters.参数值%Type;
Begin
  -------------------------------------------------------------------------------------------
  --1.先处理原料出库部分
  Select b.Id
  Into n_出库类别id
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 31 And b.系数 = -1 And Rownum < 2;

  Select Zl_Getsysparameter(120) Into v_负成本计算 From Dual;

  Select Max(序号) Into n_出序号 From 药品收发记录 Where NO = No_In And 单据 = 16 And 入出系数 = -1;
  If Nvl(n_出序号, 0) < 记录数_In Then
    n_出序号 := 记录数_In;
  End If;
  n_总出库成本 := 0;
  For v_组成 In (Select a.*, b.是否变价, c.指导差价率, c.成本价, c.在用分批
               From 自制材料构成 A, 收费项目目录 B, 材料特性 C
               Where a.原料材料id = b.Id And a.自制材料id = 材料id_In And a.原料材料id = c.材料id
			   Order By a.原料材料id) Loop
    n_剩余数量 := Round(实际数量_In * v_组成.分子 / v_组成.分母, 7);
  
    If n_剩余数量 = 0 Then
      v_Err_Msg := '[ZLSOFT]该单据中存在一笔以上原料的数量为零了！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If Nvl(v_组成.是否变价, 0) = 0 Then
      --定价处理
      Begin
        Select Nvl(现价, 0)
        Into n_现价
        From 收费价目
        Where 收费细目id = v_组成.原料材料id And ((Sysdate Between 执行日期 And 终止日期) Or (Sysdate >= 执行日期 And 终止日期 Is Null));
      Exception
        When Others Then
          v_Err_Msg := 'Err';
      End;
      If Nvl(v_Err_Msg, ' ') = 'Err' Then
        v_Err_Msg := '[ZLSOFT]该单据中存在一笔以上原料还未进行定价！[ZLSOFT]';
        Raise Err_Item;
      End If;
    Else
      n_现价 := 0;
    End If;
    n_最后出库id := -1;
    --按先进先出法原则出库
    For v_库存 In (Select Nvl(批次, 0) As 批次, Max(零售价) As 零售价, Sum(Nvl(可用数量, 0)) As 可用数量, Sum(Nvl(实际数量, 0)) As 实际数量,
                        Sum(Nvl(实际差价, 0)) As 实际差价, Sum(Nvl(实际金额, 0)) As 实际金额, Max(上次产地) As 上次产地, Max(上次批号) As 上次批号,
                        Max(上次生产日期) As 上次生产日期, Max(效期) As 效期, Max(灭菌效期) As 灭菌效期, Max(批准文号) As 批准文号
                 From 药品库存
                 Where 药品id = v_组成.原料材料id And 性质 = 1 And 库房id = 对方部门id_In
                 Group By Nvl(批次, 0)
				 Having Sum(Nvl(可用数量, 0)) > 0
                 Order By Nvl(批次, 0)) Loop
    
      Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = v_组成.原料材料id;
    
      If Nvl(v_组成.是否变价, 0) = 1 Then
        --实价处理
        If Nvl(v_库存.实际数量, 0) > 0 Then
          If Nvl(v_库存.批次, 0) <> 0 And Nvl(v_库存.零售价, 0) <> 0 Then
            --分批实价，如果库存有零售价，则只能以零售价为准.
            n_售价 := Nvl(v_库存.零售价, 0);
          Else
            n_售价 := Nvl(v_库存.实际金额, 0) / v_库存.实际数量;
          End If;
        Else
          --无库数:需提示
          v_Err_Msg := '[ZLSOFT]该单据中存在一笔以上原料的实际数量不足[ZLSOFT]';
          Raise Err_Item;
        End If;
      Else
        n_售价 := n_现价;
      End If;
      If Nvl(v_库存.可用数量, 0) >= n_剩余数量 Then
        n_当前数量 := n_剩余数量;
      Else
        n_当前数量 := Nvl(v_库存.可用数量, 0);
      End If;
      n_零售金额 := Round(n_当前数量 * n_售价, 7);
    
      --算成本价
      If Nvl(v_库存.实际金额, 0) <= 0 Then
        If v_负成本计算 = '1' And Nvl(v_组成.成本价, 0) > 0 Then
          n_成本价 := v_组成.成本价;
          n_差价   := Round(n_零售金额 - n_当前数量 * n_成本价, 7);
        Else
          n_差价   := n_零售金额 * v_组成.指导差价率 / 100;
          n_成本价 := (n_零售金额 - n_差价) / n_当前数量;
        End If;
      Else
        n_差价   := n_零售金额 * (v_库存.实际差价 / v_库存.实际金额);
        n_成本价 := (n_零售金额 - n_差价) / n_当前数量;
      End If;
      n_成本价     := Nvl(n_成本价, 0);
      n_成本金额   := n_成本价 * n_当前数量;
      n_总出库成本 := n_总出库成本 + n_成本金额;
    
      n_出序号 := n_出序号 + 1;
      Select 药品收发记录_Id.Nextval Into n_最后出库id From Dual;
    
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价,
         摘要, 填制人, 填制日期, 费用id, 扣率)
      Values
        (n_最后出库id, 1, 16, No_In, n_出序号, 对方部门id_In, 库房id_In, n_出库类别id, -1, v_组成.原料材料id,
         Decode(v_库存.批次, 0, Null, v_库存.批次), v_库存.上次批号, v_库存.效期, v_库存.灭菌效期, Nvl(n_当前数量, 0), n_当前数量, n_成本价, n_成本金额, n_售价,
         n_零售金额, n_差价, 摘要_In, 填制人_In, 填制日期_In, 材料id_In, 序号_In);
    
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) - n_当前数量
      Where 库房id = 对方部门id_In And 药品id = v_组成.原料材料id And Nvl(批次, 0) = Nvl(v_库存.批次, 0) And 性质 = 1;
    
      If Sql%NotFound Then
        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 可用数量, 零售价)
        Values
          (对方部门id_In, v_组成.原料材料id, Decode(Nvl(v_库存.批次, 0), 0, Null, v_库存.批次), 1, -n_当前数量,
           Decode(n_实价卫材, 1, Decode(Nvl(v_库存.批次, 0), 0, Null, n_售价), Null));
      End If;
    
      Delete From 药品库存
      Where 库房id = 对方部门id_In And 药品id = v_组成.原料材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
            Nvl(实际差价, 0) = 0;
      n_剩余数量 := n_剩余数量 - Nvl(v_库存.可用数量, 0);
      If n_剩余数量 <= 0 Then
        Exit;
      End If;
    End Loop;
  
    If n_剩余数量 > 0 Then
      --比库存数还多,需要将剩余数量加入
      If Nvl(v_组成.是否变价, 0) = 1 Or Nvl(v_组成.在用分批, 0) = 1 Then
        --实价或在用分批必需要有库存
        v_Err_Msg := '[ZLSOFT]该单据中存在一笔以上原料的可用数量不足，请检查[ZLSOFT]';
        Raise Err_Item;
      End If;
    
      If n_最后出库id = -1 Then
        --表示根本没有库存，需要处理相关的数量
        n_售价     := n_现价;
        n_零售金额 := Round(n_剩余数量 * n_售价, 7);
        If v_负成本计算 = '1' And Nvl(v_组成.成本价, 0) > 0 Then
          n_成本价 := v_组成.成本价;
          n_差价   := Round(n_零售金额 - n_剩余数量 * n_成本价, 7);
        Else
          n_差价   := n_零售金额 * v_组成.指导差价率 / 100;
          n_成本价 := (n_零售金额 - n_差价) / n_剩余数量;
        End If;
        n_成本价     := Nvl(n_成本价, 0);
        n_成本金额   := n_成本价 * n_剩余数量;
        n_总出库成本 := n_总出库成本 + n_成本金额;
      
        n_出序号 := n_出序号 + 1;
        Select 药品收发记录_Id.Nextval Into n_最后出库id From Dual;
      
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额,
           差价, 摘要, 填制人, 填制日期, 费用id, 扣率)
        Values
          (n_最后出库id, 1, 16, No_In, n_出序号, 对方部门id_In, 库房id_In, n_出库类别id, -1, v_组成.原料材料id, Null, Null, Null, Null, n_剩余数量,
           n_剩余数量, n_成本价, n_成本金额, n_售价, n_零售金额, n_差价, 摘要_In, 填制人_In, 填制日期_In, 材料id_In, 序号_In);
      Else
        --还存在剩余
        Select 成本价, 零售价, 零售金额, 差价, 填写数量, 成本金额
        Into n_成本价, n_售价, n_零售金额, n_差价, n_当前数量, n_成本金额
        From 药品收发记录
        Where ID = n_最后出库id;
      
        Update 药品收发记录
        Set 填写数量 = Nvl(填写数量, 0) + n_剩余数量, 实际数量 = Nvl(实际数量, 0) + n_剩余数量, 成本价 = Nvl(n_成本价, 0),
            成本金额 = Nvl(n_成本价, 0) * (n_当前数量 + n_剩余数量), 零售价 = n_售价, 零售金额 = n_售价 * (n_当前数量 + n_剩余数量),
            差价 = Round((n_售价 * (n_当前数量 + n_剩余数量)) - (Nvl(n_成本价, 0) * (n_当前数量 + n_剩余数量)), 7)
        Where ID = n_最后出库id;
        n_总出库成本 := (n_总出库成本 - n_成本金额) + Nvl(n_成本价, 0) * (n_当前数量 + n_剩余数量);
      
      End If;
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) - n_剩余数量
      Where 库房id = 对方部门id_In And 药品id = v_组成.原料材料id And Nvl(批次, 0) = 0 And 性质 = 1;
    
      If Sql%NotFound Then
        Insert Into 药品库存
          (库房id, 药品id, 性质, 可用数量, 零售价)
        Values
          (对方部门id_In, v_组成.原料材料id, 1, -n_剩余数量, Null);
      End If;
    
      Delete From 药品库存
      Where 库房id = 对方部门id_In And 药品id = v_组成.原料材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
            Nvl(实际差价, 0) = 0;
    End If;
  End Loop;

  n_成本价 := n_总出库成本 / 实际数量_In;
  n_差价   := 零售金额_In - n_总出库成本;

  Select 药品收发记录_Id.Nextval Into n_Id From Dual;

  --确定是否分批  
  Select Nvl(库房分批, 0), Nvl(在用分批, 0) Into n_库房分批, n_在用分批 From 材料特性 Where 材料id = 材料id_In;

  If n_在用分批 = 0 Then
    If n_库房分批 = 1 Then
      Begin
        Select Distinct 0
        Into n_库房分批
        From 部门性质说明
        Where (工作性质 = '发料部门' Or 工作性质 Like '制剂室') And 部门id = 库房id_In;
      Exception
        When Others Then
          n_库房分批 := 1;
      End;
    
      If n_库房分批 = 1 Then
        n_批次 := n_Id;
      End If;
    End If;
  Else
    n_批次 := n_Id;
  End If;

  Select b.Id
  Into n_入的类别id
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 31 And b.系数 = 1 And Rownum < 2;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 批号, 效期, 灭菌日期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额,
     差价, 摘要, 填制人, 填制日期)
  Values
    (n_Id, 1, 16, No_In, 序号_In, 库房id_In, 对方部门id_In, n_入的类别id, 1, 材料id_In, n_批次, 批号_In, 效期_In, 灭菌日期_In, 灭菌效期_In, 实际数量_In,
     实际数量_In, n_成本价, n_总出库成本, 零售价_In, 零售金额_In, n_差价, 摘要_In, 填制人_In, 填制日期_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_自制材料入库_Insert;
/

--137357:刘涛,2019-01-21,过程取消库存检查
Create Or Replace Procedure Zl_材料外购_Insert
(
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  供药单位id_In In 药品收发记录.供药单位id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  产地_In       In 药品收发记录.产地%Type := Null,
  批号_In       In 药品收发记录.批号%Type := Null,
  生产日期_In   In 药品收发记录.生产日期%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
  实际数量_In   In 药品收发记录.实际数量%Type := Null,
  成本价_In     In 药品收发记录.成本价%Type := Null,
  成本金额_In   In 药品收发记录.成本金额%Type := Null,
  扣率_In       In 药品收发记录.扣率%Type := Null,
  零售价_In     In 药品收发记录.零售价%Type := Null,
  零售金额_In   In 药品收发记录.零售金额%Type := Null,
  差价_In       In 药品收发记录.差价%Type := Null,
  零售差价_In   In 药品收发记录.差价%Type := Null,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  注册证号_In   In 药品收发记录.注册证号%Type := Null,
  填制人_In     In 药品收发记录.填制人%Type := Null,
  随货单号_In   In 应付记录.随货单号%Type := Null,
  发票号_In     In 应付记录.发票号%Type := Null,
  发票日期_In   In 应付记录.发票日期%Type := Null,
  发票金额_In   In 应付记录.发票金额%Type := Null,
  填制日期_In   In 药品收发记录.填制日期%Type := Null,
  核查人_In     In 药品收发记录.配药人%Type := Null,
  核查日期_In   In 药品收发记录.配药日期%Type := Null,
  批次_In       In 药品收发记录.批次%Type := 0,
  退货_In       In Number := 1,
  高值材料_In   In Varchar2 := Null,
  商品条码_In   In 药品收发记录.商品条码%Type := Null,
  内部条码_In   In 药品收发记录.内部条码%Type := Null,
  费用id_In     In 药品收发记录.费用id%Type := 0,
  发票代码_In   In 应付记录.发票代码%Type := Null,
  财务审核_In   In Number := 0,
  批准文号_In   In 药品收发记录.批准文号%Type := Null,
  验收结论_In   In 药品收发记录.验收结论%Type := Null,
  加成率_In     In 药品收发记录.频次%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  v_No         应付记录.No%Type; --应付记录的NO 
  v_商品名     收费项目目录.名称%Type; --通用名称 
  v_规格       收费项目目录.规格%Type;
  v_产地       收费项目目录.规格%Type;
  v_单位       收费项目目录.计算单位%Type;
  v_Lngid      药品收发记录.Id%Type; --收发ID 
  n_应付id     应付记录.Id%Type; --应付记录的ID 
  n_入出类别id 药品收发记录.入出类别id%Type; --入出类别ID 
  n_入出系数   药品收发记录.入出系数%Type; --入出系数 
  n_批次       药品收发记录.批次%Type := Null; --批次 
  n_库房分批   Integer; --是否分批核算    1:分批；0：不分批 
  n_在用分批   Integer; --是否在用分批       1:分批；0：不分批 
  v_可用数量   药品库存.可用数量%Type;
Begin

  If Not 批准文号_In Is Null And Not 产地_In Is Null Then
    Update 药品生产商对照 Set 批准文号 = 批准文号_In Where 药品id = 材料id_In And 厂家名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 药品生产商对照 (药品id, 厂家名称, 批准文号) Values (材料id_In, 产地_In, 批准文号_In);
    End If;
  End If;

  --取该材料的名称 
  v_产地 := '';
  Select 名称, 规格, 计算单位 Into v_商品名, v_规格, v_单位 From 收费项目目录 Where ID = 材料id_In;

  If v_规格 Is Not Null Then
    If Instr(v_规格, '|') <> 0 Then
      v_产地 := Substr(v_规格, Instr(v_规格, '|'));
      v_规格 := Substr(v_规格, Instr(v_规格, '|') - 1);
    End If;
  End If;

  Select 药品收发记录_Id.Nextval Into v_Lngid From Dual;

  Select Nvl(库房分批, 0), Nvl(在用分批, 0) Into n_库房分批, n_在用分批 From 材料特性 Where 材料id = 材料id_In;

  --财务审核直接用传过来的批次
  If 财务审核_In = 0 Then
    If 费用id_In > 0 And 批次_In > 0 Then
      n_批次 := 批次_In;
    Else
      If n_在用分批 = 0 Then
        If n_库房分批 = 1 Then
          Begin
            Select Distinct 0
            Into n_库房分批
            From 部门性质说明
            Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = 库房id_In;
          Exception
            When Others Then
              n_库房分批 := 1;
          End;
        
          If n_库房分批 = 1 Then
            n_批次 := v_Lngid;
          End If;
        Else
          n_批次 := 0;
        End If;
      Else
        n_批次 := v_Lngid;
      End If;
    End If;
  Else
    n_批次 := 批次_In;
  End If;

  Select b.Id, b.系数
  Into n_入出类别id, n_入出系数
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 30 And Rownum < 2;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 灭菌日期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额,
     扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 发药方式, 配药人, 配药日期, 注册证号, 用法, 商品条码, 内部条码, 费用id, 批准文号, 验收结论, 频次)
  Values
    (v_Lngid, 1, 15, No_In, 序号_In, 库房id_In, 供药单位id_In, n_入出类别id, n_入出系数, 材料id_In, Decode(退货_In, -1, 批次_In, n_批次), 产地_In,
     批号_In, 生产日期_In, 效期_In, 灭菌日期_In, 灭菌效期_In, 退货_In * 实际数量_In, 退货_In * 实际数量_In, 成本价_In, 退货_In * 成本金额_In, 扣率_In, 零售价_In,
     退货_In * 零售金额_In, 退货_In * 差价_In, 摘要_In, 填制人_In, 填制日期_In, Decode(退货_In, -1, 1, 0), 核查人_In, 核查日期_In, 注册证号_In, 零售差价_In,
     商品条码_In, 内部条码_In, 费用id_In, 批准文号_In, 验收结论_In, 加成率_In);

  --高值材料信息 
  If Length(高值材料_In) > 0 Then
    Insert Into 收发记录补充信息
      (收发id, 科室, 病人姓名, 住院号, 床号)
    Values
      (v_Lngid, Substr(高值材料_In, 1, Instr(高值材料_In, ',', 1, 1) - 1),
       Substr(高值材料_In, Instr(高值材料_In, ',', 1, 1) + 1, Instr(高值材料_In, ',', 1, 2) - Instr(高值材料_In, ',', 1, 1) - 1),
       Substr(高值材料_In, Instr(高值材料_In, ',', 1, 2) + 1, Instr(高值材料_In, ',', 1, 3) - Instr(高值材料_In, ',', 1, 2) - 1),
       Substr(高值材料_In, Instr(高值材料_In, ',', 1, 3) + 1, Length(高值材料_In)));
  End If;

  If 发票号_In Is Not Null Or 随货单号_In Is Not Null Then
  
    Select 应付记录_Id.Nextval Into n_应付id From Dual;
  
    --如果是第一笔明细,则产生应付记录的NO 
    Begin
      Select NO
      Into v_No
      From 应付记录
      Where 系统标识 = 5 And 记录性质 = 0 And 记录状态 = 1 And 入库单据号 = No_In And Rownum < 2;
    Exception
      When Others Then
        v_No := Nextno(67);
    End;
  
    Insert Into 应付记录
      (ID, 记录性质, 记录状态, 项目id, 序号, 单位id, NO, 系统标识, 收发id, 入库单据号, 单据金额, 随货单号, 发票号, 发票日期, 发票金额, 品名, 规格, 产地, 批号, 计量单位, 数量,
       采购价, 采购金额, 填制人, 填制日期, 审核人, 审核日期, 摘要, 库房id, 发票代码)
    Values
      (n_应付id, 0, 1, 材料id_In, 序号_In, 供药单位id_In, v_No, 5, v_Lngid, No_In, 退货_In * 零售金额_In, 随货单号_In, 发票号_In, 发票日期_In,
       退货_In * Decode(发票号_In, Null, 成本金额_In, 发票金额_In), v_商品名, v_规格, v_产地, 批号_In, v_单位, 退货_In * 实际数量_In, 成本价_In,
       退货_In * 成本金额_In, 填制人_In, 填制日期_In, Null, Null, 摘要_In, 库房id_In, 发票代码_In);
  End If;

  --退货时下可用数量 
  If 退货_In = -1 And Nvl(费用id_In, 0) <> 2 Then
    --检查库存 
    Begin
      Select Nvl(可用数量, 0)
      Into v_可用数量
      From 药品库存
      Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 性质 = 1;
    Exception
      When Others Then
        v_可用数量 := 0;
    End;
	
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) - 实际数量_In
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = 批次_In And 性质 = 1;
    Delete From 药品库存
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料外购_Insert;
/

--137357:刘涛,2019-01-21,过程取消可用库存检查
Create Or Replace Procedure Zl_材料申领_Insert
(
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  对方部门id_In In 药品收发记录.对方部门id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  批次_In       In 药品收发记录.批次%Type,
  填写数量_In   In 药品收发记录.填写数量%Type,
  实际数量_In   In 药品收发记录.实际数量%Type,
  成本价_In     In 药品收发记录.成本价%Type,
  成本金额_In   In 药品收发记录.成本金额%Type,
  零售价_In     In 药品收发记录.零售价%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  差价_In       In 药品收发记录.差价%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  产地_In       In 药品收发记录.产地%Type := Null,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  填制日期_In   In 药品收发记录.填制日期%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(500);
  v_批准文号 药品库存.批准文号%Type;

  n_Id         药品收发记录.Id%Type; --收发ID
  n_入的类别id 药品收发记录.入出类别id%Type; --入出类别ID
  n_出的类别id 药品收发记录.入出类别id%Type; --入出类别ID
  v_下库存     Zlparameters.参数值%Type;
  v_明确批次   Zlparameters.参数值%Type;

  v_编码         收费项目目录.编码%Type;
  n_可用数量     药品库存.可用数量%Type;
  n_批次         药品收发记录.批次%Type := Null; --主要针对入库中实行分批核算的材料
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_上次生产日期 药品库存.上次生产日期%Type;
  n_实价卫材     收费项目目录.是否变价%Type;
  n_零售价       药品收发记录.零售价%Type;
  v_商品条码     药品库存.商品条码%Type;
  v_内部条码     药品库存.内部条码%Type;

  n_是否分批 Integer; --判断入库是否分批核算   1:分批；0：不分批
  n_库房分批 Integer; --判断入库是否分批核算   1:分批；0：不分批
  n_在用分批 Integer; --判断入库是否分批核算   1:分批；0：不分批
  v_Records  Number;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_下库存 From Dual;

  --只有在明确批次的情况下才能下可用库存
  Select Nvl(zl_GetSysParameter(83), '0') Into v_明确批次 From Dual;

  --首先找出入和出的类别ID
  Select b.Id
  Into n_入的类别id
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 34 And b.系数 = 1 And Rownum < 2;

  Select b.Id
  Into n_出的类别id
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 34 And b.系数 = -1 And Rownum < 2;

  Select 药品收发记录_Id.Nextval Into n_Id From Dual;

  Begin
    Select 可用数量, 零售价, Decode(上次供应商id, 0, Null, 上次供应商id), 上次生产日期, 批准文号, 商品条码, 内部条码
    Into n_可用数量, n_零售价, n_上次供应商id, n_上次生产日期, v_批准文号, v_商品条码, v_内部条码
    From 药品库存
    Where 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 库房id = 库房id_In And 性质 = 1 And Rownum = 1;
  Exception
    When Others Then
      n_可用数量     := 0;
      n_上次供应商id := Null;
      n_上次生产日期 := Null;
      v_批准文号     := Null;
  End;

  --插入类别为出的那一笔
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价,
     摘要, 填制人, 填制日期, 发药方式, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
  Values
    (药品收发记录_Id.Nextval, 1, 19, No_In, 序号_In, 库房id_In, 对方部门id_In, n_出的类别id, -1, 材料id_In, 批次_In, 产地_In, 批号_In, 效期_In,
     灭菌效期_In, 填写数量_In, 实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, 1, n_上次供应商id, n_上次生产日期,
     v_批准文号, v_商品条码, v_内部条码);

  Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;

  If To_Number(v_下库存, '99999') = 1 And To_Number(v_明确批次, '99999') = 1 Then

    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) - 实际数量_In
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 性质 = 1;

    If Sql%RowCount = 0 Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
      Values
        (库房id_In, 材料id_In, Nvl(批次_In, 0), 1, -实际数量_In, 效期_In, 灭菌效期_In, n_上次供应商id, 成本价_In, 批号_In, n_上次生产日期, 产地_In,
         v_批准文号, Decode(n_实价卫材, 1, Decode(Nvl(批次_In, 0), 0, Null, 零售价_In), Null));
    End If;

    --同时更新库存数
    Delete From 药品库存
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;

    --下面是判断入库的材料是否是分批核算材料
    Select Nvl(库房分批, 0), Nvl(在用分批, 0) Into n_库房分批, n_在用分批 From 材料特性 Where 材料id = 材料id_In;

    n_是否分批 := 0;
    If n_在用分批 = 0 Then
      If n_库房分批 = 1 Then
        Begin
          Select Distinct 0
          Into n_是否分批
          From 部门性质说明
          Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = 对方部门id_In;
        Exception
          When Others Then
            n_是否分批 := 1;
        End;
      End If;
    Else
      n_是否分批 := 1;
    End If;

    If n_是否分批 = 1 And Nvl(批次_In, 0) = 0 Then
      --入库分批且出库不分批
      n_批次 := n_Id;
    Elsif n_是否分批 = 0 Then
      --入库不分批
      n_批次 := 0;
    Elsif Nvl(批次_In, 0) <> 0 Then
      --入库分批且出库也分批
      n_批次 := 批次_In;
    End If;

    --插入类别为入的那一笔
    Insert Into 药品收发记录
      (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额,
       差价, 摘要, 填制人, 填制日期, 发药方式, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
    Values
      (n_Id, 1, 19, No_In, 序号_In + 1, 对方部门id_In, 库房id_In, n_入的类别id, 1, 材料id_In, n_批次, 产地_In, 批号_In, 效期_In, 灭菌效期_In,
       填写数量_In, 实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, 1, n_上次供应商id, n_上次生产日期, v_批准文号,
       v_商品条码, v_内部条码);

    --检查是否存在相同材料相同批次的数据，如果存在不允许保存
    Select Count(*)
    Into v_Records
    From 药品收发记录
    Where 单据 = 19 And NO = No_In And 入出系数 = -1 And 药品id + 0 = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0);

    If v_Records > 1 Then
      Select 编码 Into v_编码 From 收费项目目录 Where ID = 材料id_In;
      v_Err_Msg := '[ZLSOFT]编码为' || v_编码 || '的药品，存在多条重复的记录，请合并为一条记录！[ZLSOFT]';
      Raise Err_Item;
    End If;
  Else
    --插入类别为入的那一笔
    Insert Into 药品收发记录
      (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额,
       差价, 摘要, 填制人, 填制日期, 发药方式, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
    Values
      (n_Id, 1, 19, No_In, 序号_In + 1, 对方部门id_In, 库房id_In, n_入的类别id, 1, 材料id_In, 批次_In, 产地_In, 批号_In, 效期_In, 灭菌效期_In,
       填写数量_In, 实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, 1, n_上次供应商id, n_上次生产日期, v_批准文号,
       v_商品条码, v_内部条码);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料申领_Insert;
/
--137476:蒋敏,2019-02-13,诊疗部位修改时，关联的诊疗项目会没有勾选
--137245:蒋敏,2019-01-21,过程中解析方法解析不对，导致方法前面多0或者多1
CREATE OR REPLACE Procedure Zl_诊疗检查部位_Edit
(
  操作_In     In Number, --1:增加;2:修改;3:删除
  类型_In     In 诊疗检查部位.类型%Type,
  原编码_In   In 诊疗检查部位.编码%Type,
  新编码_In   In 诊疗检查部位.编码%Type := Null,
  名称_In     In 诊疗检查部位.名称%Type := Null,
  分组_In     In 诊疗检查部位.分组%Type := Null,
  备注_In     In 诊疗检查部位.备注%Type := Null,
  方法_In     In 诊疗检查部位.方法%Type := Null,
  适用性别_In In 诊疗检查部位.适用性别%Type := Null,
  上级方法_In In 诊疗检查部位.方法%Type := Null --格式：上级方法|方法;上级方法|方法...(若上级方法为空，则为|方法，若为同一个上级，则用逗号分隔)
) Is
  v_原名称 诊疗检查部位.名称%Type := Null;
  e_Notfind Exception;
  v_方法   Varchar2(1000);
  v_Fields Varchar2(1000);
  v_Tmp    Varchar2(1000);
  n_Count  Number;
  n_记录id 诊疗项目部位.Id%Type;
  v_上级 Varchar(100);
  v_格式方法 Varchar2(1000);
Begin
  If 操作_In = 1 Then
    Insert Into 诊疗检查部位
      (类型, 编码, 名称, 分组, 备注, 方法, 适用性别)
    Values
      (类型_In, 新编码_In, 名称_In, 分组_In, 备注_In, 方法_In, 适用性别_In);
  Elsif 操作_In = 2 Then
    Begin
      Select 名称 Into v_原名称 From 诊疗检查部位 Where 编码 = 原编码_In And 类型 = 类型_In;
    Exception
      When Others Then
        Null;
    End;
    If v_原名称 Is Null Then
      Raise e_Notfind;
    End If;
    Update 诊疗检查部位
    Set 编码 = 新编码_In, 名称 = 名称_In, 分组 = 分组_In, 备注 = 备注_In, 方法 = 方法_In, 适用性别 = 适用性别_In
    Where 编码 = 原编码_In And 类型 = 类型_In;
    --级联修改
    --原来没有的方法现在新增
    v_方法 := ';' || Replace(方法_In,Chr(32),';');
    v_方法 := Replace(v_方法, ';;', ';');  
    v_方法 := Replace(v_方法, ';;', ';');      
    v_方法 := Replace(v_方法, ',', Chr(10));
    v_方法 := Replace(v_方法, Chr(9), ';');
    v_方法 := Replace(v_方法, ';0', Chr(10) || '(上级)0');
    v_方法 := Replace(v_方法, ';1', Chr(10) || '(上级)1');
    v_方法 := Replace(v_方法, Chr(10), ';');
    v_方法 := Replace(v_方法, ';;', ';');
    v_方法 := v_方法 || ';';
    v_方法 := Substr(v_方法, 2);
    While v_方法 Is Not Null Loop
      --依次取每个项目v_Tmp
      v_Fields := Substr(v_方法, 1, Instr(v_方法, ';') - 1);
      v_方法    := Substr(v_方法, Instr(v_方法, ';') + 1);
      If Substr(v_Fields, 1, 4) = '(上级)' Then
        v_Fields := Substr(v_Fields, 5);
        v_Fields:=Substr(v_Fields, 2);
        v_Tmp:=v_Fields;
        v_上级:=NULL;
      Else
        v_Fields:=Substr(v_Fields, 2);
        v_上级   := v_Tmp;
      End If;
      If v_Fields Is Not Null Then
        v_格式方法:=v_格式方法 ||';'|| v_Fields ||','|| Nvl(v_上级,' ');
        For r_Used In (Select Distinct 项目id From 诊疗项目部位 Where 部位 = 名称_In And 类型 = 类型_In) Loop
          Select Count(ID)
          Into n_Count
          From 诊疗项目部位
          Where 项目id = r_Used.项目id And 部位 = 名称_In And 类型 = 类型_In And 方法 = v_Fields And Nvl(上级方法,' ')=Nvl(v_上级,' ');
          If n_Count = 0 Then            
            Select 诊疗项目部位_Id.Nextval Into n_记录id From Dual;
            Insert Into 诊疗项目部位
              (ID, 项目id, 类型, 部位, 方法,上级方法)
            Values
              (n_记录id, r_Used.项目id, 类型_In, 名称_In, v_Fields,v_上级);
          End If;
        End Loop;
      End If;
    End Loop;
    --原有的方法，现在已经删除了或原有的部位的名称已经改变了
    v_格式方法 :=v_格式方法|| ';';
    For r_Used In (Select ID, 项目id, 部位, 方法, 类型, 默认,上级方法 From 诊疗项目部位 Where 部位 = v_原名称 And 类型 = 类型_In) Loop
      If Instr(v_格式方法, ';' || r_Used.方法 ||','|| Nvl(r_Used.上级方法,' ') || ';') = 0 Then
        Delete 诊疗项目部位
        Where id=r_Used.id;
      Else
        Update 诊疗项目部位
        Set 部位 = 名称_In
        Where id=r_Used.id;
      End If;
    End Loop;    
  Elsif 操作_In = 3 Then
    Delete 诊疗检查部位 Where 编码 = 原编码_In And 类型 = 类型_In;
  End If;
Exception
  When e_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该部位不存在，可能已被其他用户删除修改！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗检查部位_Edit;
/

--134969:李南春,2019-01-14,预交支付检查
CREATE OR REPLACE Procedure Zl_病人挂号记录_Insert
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
  更新交款余额_In Number := 0, --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况 
  修正病人年龄_In Number := 0, 
  收费单_In       病人挂号记录.收费单%Type := Null 
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
    Select NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id, 
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id, Min(Decode(记录性质, 1, 收款时间, NULL)) as 收款时间
    From 病人预交记录 
    Where 记录性质 In (1, 11) And 病人id = v_病人id And Nvl(预交类别, 2) = 1 Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0 
    Group By NO 
    Order By 收款时间; 
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
  n_病人余额 病人预交记录.金额%Type;
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
  v_Temp           Varchar2(3000); 
  n_分时点显示     Number(3); 
  n_号序使用否     Number(3); 
  d_启用时间       Date; 
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
 
  If Nvl(修正病人年龄_In, 0) = 1 Then 
    Begin 
      Update 病人信息 Set 年龄 = 年龄_In Where 病人id = 病人id_In; 
    Exception 
      When Others Then 
        v_Err_Msg := '没有找到对应的病人！'; 
        Raise Err_Item; 
    End; 
  End If; 
 
  Begin 
    Delete From 挂号序号状态 
    Where 号码 = 号别_In And 日期 = 发生时间_In And 序号 = 号序_In And 状态 = 3 And 操作员姓名 = 操作员姓名_In; 
  Exception 
    When Others Then 
      Null; 
  End; 
  v_Temp := zl_GetSysParameter(256); 
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then 
    Null; 
  Else 
    Begin 
      d_启用时间 := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss'); 
    Exception 
      When Others Then 
        d_启用时间 := Null; 
    End; 
    If d_启用时间 Is Not Null Then 
      If 发生时间_In > d_启用时间 Then 
        v_Err_Msg := '当前挂号的发生时间' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '已经启用了出诊表排班模式,不能再使用计划排班模式挂号!'; 
        Raise Err_Item; 
      End If; 
    End If; 
  End If; 
 
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
          --挂号时检查 ，非加号的情况下，检测当前号序是否被使用，防止并发情况下导致序号重复的可能。
          --启用序号控制未分时段 达到了限制 
          --检测挂号记录中当前序号是否已经使用，若未使用则不检测挂号数量
          Select Count(1)
          Into n_号序使用否
          From 挂号序号状态 
          Where 号码 = 号别_In And 序号= nvl(号序_In,0) And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 操作员姓名 <> 操作员姓名_In; 
          if nvl(号序_In,0)>0 And n_号序使用否=1 then
              v_Err_Msg := '号别' || 号别_In || '号序' || to_char(n_号序使用否) || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已被其它用户使用！'; 
              Raise Err_Item; 
          end if;
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
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then 
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号; 
            End If; 
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
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then 
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号; 
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
        If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then 
          Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号; 
        End If; 
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
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), 登记时间_In, 
         操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4); 
 
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
      Select Nvl(Sum(Nvl(预交余额, 0) - Nvl(费用余额, 0)), 0)
      Into n_病人余额
      From 病人余额
      Where 病人id = 病人id_In And Nvl(性质, 0) = 1 And Nvl(类型, 0) = 1;
      if n_病人余额 < 预交支付_In Then
        v_Err_Msg := '病人的当前预交余额为 ' || Ltrim(To_Char(n_病人余额, '9999999990.00')) || '，小于本次支付金额 ' ||
                     Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，支付失败！';
        Raise Err_Item;
      End if;
      
      n_预交金额 := 预交支付_In; 
      For r_Deposit In c_Deposit(病人id_In) Loop 
        n_当前金额 := Case 
                    When r_Deposit.金额 - n_预交金额 < 0 Then 
                     r_Deposit.金额 
                    Else 
                     n_预交金额 
                  End; 
 
        If r_Deposit.结帐id = 0 Then 
          --第一次冲预交(82592,将第一次标上结帐ID,冲预交标记为0) 
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id; 
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
        Where 病人id = 病人id_In And 性质 = 1 And 类型 = Nvl(1, 2); 
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
    If 序号_In = 1 And Nvl(更新交款余额_In, 0) = 0 Then 
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
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 收费单) 
    Values 
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In, 
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In, 
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In, 
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null), 
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null), 
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 收费单_In); 
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then 
      Update 病人挂号记录 
      Set 预约 = 1, 预约时间 = 发生时间_In, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In 
      Where ID = n_挂号id; 
    End If; 
    n_预约生成队列 := 0; 
    If Nvl(预约挂号_In, 0) = 1 Then 
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113)); 
    End If; 
 
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站 
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then 
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113)); 
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then 
        n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0); 
        If Nvl(预约挂号_In, 0) = 1 And n_分时点显示 = 1 And n_分时段 = 1 Then 
          n_分时点显示 := 1; 
        Else 
          n_分时点显示 := Null; 
        End If; 
        --产生队列 
        --.按”执行部门” 的方式生成队列 
        v_队列名称 := 执行部门id_In; 
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号); 
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0); 
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In); 
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In 
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In, 
                         n_分时点显示, v_排队序号); 
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
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists 
     (Select 1 
           From 病人担保记录 
           Where 病人id = 病人id_In And 主页id Is Not Null And 
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In)); 
 
    If Sql%RowCount > 0 Then 
      Update 病人担保记录 
      Set 到期时间 = Sysdate 
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) >= Sysdate; 
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

--136685:焦博,2019-01-14,当传入的参数大于4000个字符时，分次调用zl_记帐报警线_Modify保存数据
Create Or Replace Procedure Zl_记帐报警线_Modify
(
  报警线_In In Varchar2,
  站点_In   Varchar2 := Null,
  删除_In   Number := 0
) Is
  ---------------------------------------------------------------------------------------------------------------------
  --报警线_IN 参数的填写方式如下：
  --        "适用病人|病区ID,报警方法,报警值,报警标志1,报警标志2,报警标志3,催款下限,催款标准...."
  --站点_In:NULL为不分站点设置;否则是分站点设置
  --删除_In:1-先删除数据再插入;0-不删除数据,仅插入
  ---------------------------------------------------------------------------------------------------------------------
  n_Pos       Number;
  v_Temp      Varchar2(4000);
  n_病区id    记帐报警线.病区id%Type;
  v_适用病人  记帐报警线.适用病人%Type;
  v_报警方法  记帐报警线.报警方法%Type;
  n_报警值    记帐报警线.报警值%Type;
  v_报警标志1 记帐报警线.报警标志1%Type;
  v_报警标志2 记帐报警线.报警标志2%Type;
  v_报警标志3 记帐报警线.报警标志3%Type;
  n_催款下限  记帐报警线.催款下限%Type;
  n_催款标准  记帐报警线.催款标准%Type;
Begin
  v_Temp := 报警线_In;

  v_适用病人 := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1);
  If 删除_In = 1 Then
    If 站点_In Is Null Then
      Delete From 记帐报警线 J Where 适用病人 = v_适用病人;
    Else
      Delete From 记帐报警线 J
      Where 适用病人 = v_适用病人 And (Exists (Select 1 From 部门表 Where ID = j.病区id And 站点 = 站点_In) Or 病区id Is Null);
    End If;
  End If;
  --循环处理
  v_Temp := Substr(v_Temp, Instr(v_Temp, '|') + 1);
  While v_Temp Is Not Null Loop
    n_Pos := Instr(v_Temp, ',');
  
    If n_Pos = 0 Then
      v_Temp := '';
    Else
      --得到病区ID
      If Substr(v_Temp, 1, n_Pos - 1) Is Null Then
        n_病区id := Null;
      Else
        n_病区id := To_Number(Substr(v_Temp, 1, n_Pos - 1));
      End If;
      v_Temp := Substr(v_Temp, n_Pos + 1);
    
      --得到报警方法
      n_Pos      := Instr(v_Temp, ',');
      v_报警方法 := To_Number(Substr(v_Temp, 1, n_Pos - 1));
      v_Temp     := Substr(v_Temp, n_Pos + 1);
      --得到报警值
      n_Pos    := Instr(v_Temp, ',');
      n_报警值 := To_Number(Substr(v_Temp, 1, n_Pos - 1));
      v_Temp   := Substr(v_Temp, n_Pos + 1);
    
      --得到报警标志1
      n_Pos       := Instr(v_Temp, ',');
      v_报警标志1 := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp      := Substr(v_Temp, n_Pos + 1);
    
      --得到报警标志2
      n_Pos       := Instr(v_Temp, ',');
      v_报警标志2 := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp      := Substr(v_Temp, n_Pos + 1);
    
      --得到报警标志3
      n_Pos       := Instr(v_Temp, ',');
      v_报警标志3 := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp      := Substr(v_Temp, n_Pos + 1);
    
      --得到催款下限
      n_Pos      := Instr(v_Temp, ',');
      n_催款下限 := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp     := Substr(v_Temp, n_Pos + 1);
      --催款标准
      n_Pos      := Instr(v_Temp, ',');
      n_催款标准 := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp     := Substr(v_Temp, n_Pos + 1);
    
      Insert Into 记帐报警线
        (病区id, 适用病人, 报警方法, 报警值, 报警标志1, 报警标志2, 报警标志3, 催款下限, 催款标准)
      Values
        (n_病区id, v_适用病人, v_报警方法, n_报警值, v_报警标志1, v_报警标志2, v_报警标志3, n_催款下限, n_催款标准);
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_记帐报警线_Modify;
/

--132624:蒋廷中,2019-01-14,医嘱的生效时间都相同时不停止
Create Or Replace Procedure Zl_病人医嘱记录_校对
(
  --功能：校对指定的医嘱
  --参数：医嘱ID_IN=Nvl(相关ID,ID)
  --      状态_IN=校对通过3或校对疑问2
  --      自动校对_IN=保存之后调用自动校对,自动填写计价内容
  --说明：一组医嘱只能调用一次,过程同时完成处理一组医嘱的校对
  医嘱id_In     病人医嘱记录.Id%Type,
  状态_In       病人医嘱记录.医嘱状态%Type,
  校对时间_In   病人医嘱状态.操作时间%Type,
  校对说明_In   病人医嘱状态.操作说明%Type := Null,
  自动校对_In   Number := Null,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null
) Is
  --用于医嘱检查
  v_状态       病人医嘱记录.医嘱状态%Type;
  v_期效       病人医嘱记录.医嘱期效%Type;
  v_病人id     病人医嘱记录.病人id%Type;
  v_主页id     病人医嘱记录.主页id%Type;
  v_婴儿       病人医嘱记录.婴儿%Type;
  v_医嘱内容   病人医嘱记录.医嘱内容%Type;
  v_开嘱时间   病人医嘱记录.开嘱时间%Type;
  v_开始时间   病人医嘱记录.开始执行时间%Type;
  v_开嘱医生   病人医嘱记录.开嘱医生%Type;
  v_前提id     病人医嘱记录.前提id%Type;
  v_执行标记   病人医嘱记录.执行标记%Type;
  v_执行科室id 病人医嘱记录.执行科室id%Type;
  v_标本部位   病人医嘱记录.标本部位%Type;
  v_停止时间   病人医嘱记录.开嘱时间%Type;
  v_开嘱科室id 病人医嘱记录.开嘱科室id%Type;

  --用于变更护理等级
  v_诊疗类别   病人医嘱记录.诊疗类别%Type;
  v_诊疗项目id 病人医嘱记录.诊疗项目id%Type;
  v_操作类型   诊疗项目目录.操作类型%Type;
  v_护理等级id 病案主页.护理等级id%Type;
  v_紧急标志   病人医嘱记录.紧急标志%Type;
  v_入院方式   入院方式.名称%Type;

  v_Stopadviceids 病人医嘱记录.医嘱内容%Type;
  n_Adviceid      病人医嘱记录.病人id%Type;
  n_标记          Number(18);
  --与该项目同一自动停止互斥组的项目:组中应该都是长嘱(包括当前医嘱),程序应已检查。
  --注意应加婴儿条件,同时也应停止除当前医嘱外的其它相同诊疗项目的医嘱。
  Cursor c_Exclude Is
    Select Distinct b.Id As 医嘱id, b.开始执行时间, b.执行终止时间, b.上次执行时间, b.开嘱医生, b.执行时间方案, b.频率间隔, b.频率次数, b.间隔单位
    From 诊疗互斥项目 A, 病人医嘱记录 B
    Where a.类型 = 3 And a.项目id = b.诊疗项目id And b.Id <> 医嘱id_In And Nvl(b.医嘱期效, 0) = 0 And b.医嘱状态 In (3, 5, 6, 7) And
          b.病人id = v_病人id And Nvl(b.主页id, 0) = Nvl(v_主页id, 0) And Nvl(b.婴儿, 0) = Nvl(v_婴儿, 0) And
          a.组编号 In (Select Distinct 组编号 From 诊疗互斥项目 Where 类型 = 3 And 项目id = v_诊疗项目id)
    Order By b.Id;
  v_终止时间 病人医嘱记录.执行终止时间%Type;

  --护理等级互斥
  Cursor c_Nurse Is
    Select a.Id As 医嘱id, a.开始执行时间, a.执行终止时间, a.上次执行时间, a.开嘱医生
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id And a.诊疗类别 = 'H' And b.操作类型 = '1' And a.病人id = v_病人id And Nvl(a.主页id, 0) = Nvl(v_主页id, 0) And
          Nvl(a.婴儿, 0) = Nvl(v_婴儿, 0) And Nvl(a.医嘱期效, 0) = 0 And a.医嘱状态 In (3, 5, 6, 7) And a.Id <> 医嘱id_In;

  --记录入出量互斥
  Cursor c_Patiio Is
    Select a.Id As 医嘱id, a.开始执行时间, a.执行终止时间, a.上次执行时间, a.开嘱医生
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id And a.诊疗类别 = 'Z' And b.操作类型 = '12' And a.病人id = v_病人id And Nvl(a.主页id, 0) = Nvl(v_主页id, 0) And
          Nvl(a.婴儿, 0) = Nvl(v_婴儿, 0) And Nvl(a.医嘱期效, 0) = 0 And a.医嘱状态 In (3, 5, 6, 7) And a.Id <> 医嘱id_In;

  --记录病情互斥
  Cursor c_Patistate Is
    Select a.Id As 医嘱id, a.开始执行时间, a.执行终止时间, a.上次执行时间, a.开嘱医生
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id And a.诊疗类别 = 'Z' And b.操作类型 In ('9', '10') And a.病人id = v_病人id And
          Nvl(a.主页id, 0) = Nvl(v_主页id, 0) And Nvl(a.婴儿, 0) = Nvl(v_婴儿, 0) And Nvl(a.医嘱期效, 0) = 0 And
          a.医嘱状态 In (3, 5, 6, 7) And a.Id <> 医嘱id_In;
  --变动有效记录
  Cursor c_Oldinfo Is
    Select b.*
    From (Select c.*
           From 病人变动记录 C
           Where c.病人id = v_病人id And c.主页id = v_主页id And
                 c.开始时间 = (Select Min(开始时间)
                           From 病人变动记录
                           Where 病人id = v_病人id And 主页id = v_主页id And 开始时间 > v_开始时间) And
                 NVL(c.终止时间|| '','空') = (Select  NVL(Min(终止时间)|| '','空')
                           From 病人变动记录
                           Where 病人id = v_病人id And 主页id = v_主页id And 开始时间 > v_开始时间)) A, 病人变动记录 B
    
    Where b.病人id = v_病人id And b.主页id = v_主页id And a.开始时间 = b.终止时间 And a.开始原因 = b.终止原因 And a.附加床位 = b.附加床位
    Union
    Select *
    From 病人变动记录
    Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null And 开始时间 <= v_开始时间;

  Cursor c_Endinfo Is
    Select * From 病人变动记录 Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null;
  r_Oldinfo      c_Oldinfo%RowType;
  r_Endinfo      c_Endinfo%RowType;
  v_变动终止原因 病人变动记录.终止原因%Type;
  v_变动终止时间 病人变动记录.终止时间%Type;
  v_变动终止人员 病人变动记录.终止人员%Type;

  --包含病人(婴儿)的所有未停长嘱(含配方长嘱)
  Cursor c_Needstop
  (
    v_病人id   病人医嘱记录.病人id%Type,
    v_主页id   病人医嘱记录.主页id%Type,
    v_婴儿     病人医嘱记录.婴儿%Type,
    v_Stoptime Date
  ) Is
    Select a.Id, a.诊疗类别, b.操作类型, b.执行频率
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id(+) And a.病人id = v_病人id And a.主页id = v_主页id And (v_婴儿 = -1 Or Nvl(a.婴儿, 0) = Nvl(v_婴儿, 0)) And
          Nvl(a.医嘱期效, 0) = 0 And a.医嘱状态 Not In (1, 2, 4, 8, 9) And a.开始执行时间 < v_Stoptime
    Order By a.序号;
  --包含病人(婴儿)的已停但未确认的长嘱,终止执行时间在指定时间之后
  Cursor c_Havestop
  (
    v_病人id   病人医嘱记录.病人id%Type,
    v_主页id   病人医嘱记录.主页id%Type,
    v_婴儿     病人医嘱记录.婴儿%Type,
    v_Stoptime Date
  ) Is
    Select ID
    From 病人医嘱记录
    Where 病人id = v_病人id And 主页id = v_主页id And (v_婴儿 = -1 Or Nvl(婴儿, 0) = Nvl(v_婴儿, 0)) And Nvl(医嘱期效, 0) = 0 And
          医嘱状态 = 8 And 执行终止时间 > v_Stoptime And 开始执行时间 < v_Stoptime
    Order By 序号;

  --取一组医嘱的计价内容
  Cursor c_Price Is
    Select a.Id, b.收费项目id, b.收费数量, b.从属项目, b.费用性质, b.收费方式, c.类别 As 收费类别, a.诊疗类别, e.操作类型, e.试管编码,
           Sum(Decode(Nvl(c.是否变价, 0), 1, Nvl(d.缺省价格, d.原价), Null)) As 单价
    From 病人医嘱记录 A, 诊疗收费关系 B, 收费项目目录 C, 收费价目 D, 诊疗项目目录 E
    Where a.诊疗项目id = b.诊疗项目id And b.收费项目id = c.Id And c.Id = d.收费细目id And
          (a.相关id Is Null And a.执行标记 In (1, 2) And b.费用性质 = 1 Or
          a.标本部位 = b.检查部位 And a.检查方法 = b.检查方法 And Nvl(b.费用性质, 0) = 0 Or
          a.检查方法 Is Null And Nvl(b.费用性质, 0) = 0 And b.检查部位 Is Null And b.检查方法 Is Null) And
          a.诊疗类别 Not In ('5', '6', '7') And Nvl(a.计价特性, 0) = 0 And Nvl(a.执行性质, 0) Not In (0, 5) And c.服务对象 In (2, 3) And
          (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And Sysdate Between d.执行日期 And
          Nvl(d.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(b.收费数量, 0) <> 0 And
          Not (Nvl(c.是否变价, 0) = 1 And Nvl(Nvl(d.缺省价格, d.原价), 0) = 0) And a.诊疗项目id = e.Id And
          (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In)
    Group By a.Id, b.收费项目id, b.收费数量, b.从属项目, b.费用性质, b.收费方式, c.类别, a.诊疗类别, e.操作类型, e.试管编码;

  Cursor c_Pati(v_病人id 病人信息.病人id%Type) Is
    Select * From 病人信息 Where 病人id = v_病人id;
  r_Pati c_Pati%RowType;

  v_材料id 采血管类型.材料id%Type;

  --其它临时变量
  v_Count    Number;
  v_Date     Date;
  v_Temp     Varchar2(255);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_叮嘱执行 Varchar2(5);

  v_Error Varchar2(255);
  Err_Custom Exception;

  Function Getadvicetext(v_医嘱id 病人医嘱记录.Id%Type) Return Varchar2 Is
    v_Text 病人医嘱记录.医嘱内容%Type;
    v_类别 病人医嘱记录.诊疗类别%Type;
    v_配方 Number;
  Begin
    Select 诊疗类别, 医嘱内容 Into v_类别, v_Text From 病人医嘱记录 Where ID = v_医嘱id;
    If v_类别 = 'E' Then
      --西药，中成药的医嘱内容
      Begin
        Select 诊疗类别, Decode(诊疗类别, '7', v_Text, 医嘱内容)
        Into v_类别, v_Text
        From 病人医嘱记录
        Where 相关id = v_医嘱id And 诊疗类别 In ('5', '6', '7') And Rownum = 1;
      Exception
        When Others Then
          Null;
      End;
      If v_类别 = '7' Then
        v_配方 := 1;
      End If;
    End If;
    If Length(v_Text) > 30 Then
      v_Text := Substr(v_Text, 1, 30) || '...';
    End If;
    If Length(v_Text) > 20 Then
      v_Text := '"' || v_Text || '"' || Chr(13) || Chr(10);
    Else
      v_Text := '"' || v_Text || '"';
    End If;
    If v_配方 = 1 Then
      v_Text := '中药配方' || v_Text;
    End If;
    Return(v_Text);
  End;
Begin
  --检查医嘱状态是否正确:并发操作
  Begin
    Select a.医嘱期效, a.医嘱状态, a.开嘱时间, a.开嘱医生, a.开始执行时间, a.病人id, a.主页id, a.婴儿, a.医嘱内容, a.诊疗类别, a.诊疗项目id, a.前提id,
           Nvl(b.操作类型, '0'), Nvl(a.执行标记, 0), a.执行科室id, a.标本部位, a.开嘱科室id, Nvl(a.紧急标志, 0) As 紧急标志
    Into v_期效, v_状态, v_开嘱时间, v_开嘱医生, v_开始时间, v_病人id, v_主页id, v_婴儿, v_医嘱内容, v_诊疗类别, v_诊疗项目id, v_前提id, v_操作类型, v_执行标记,
         v_执行科室id, v_标本部位, v_开嘱科室id, v_紧急标志
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id(+) And a.Id = 医嘱id_In;
  Exception
    When Others Then
      Begin
        v_Error := '医嘱已被删除，不能进行校对。' || Chr(13) || Chr(10) || '这可能是并发操作引起的，请重新读取校对数据。';
        Raise Err_Custom;
      End;
  End;
  If v_状态 <> 1 Then
    v_Error := '医嘱"' || Getadvicetext(医嘱id_In) || '"不是新开的医嘱，不能通过校对。' || Chr(13) || Chr(10) || '这可能是并发操作引起的，请重新读取校对数据。';
    Raise Err_Custom;
  End If;
  --再次检查校对时间的有效性:并发操作
  If To_Char(v_开嘱时间, 'YYYY-MM-DD HH24:MI') <= To_Char(v_开始时间, 'YYYY-MM-DD HH24:MI') Then
    If To_Char(校对时间_In, 'YYYY-MM-DD HH24:MI') < To_Char(v_开嘱时间, 'YYYY-MM-DD HH24:MI') Then
      v_Error := '医嘱"' || Getadvicetext(医嘱id_In) || '"的校对时间不能小于开嘱时间 ' || To_Char(v_开嘱时间, 'YYYY-MM-DD HH24:MI') || '。' ||
                 Chr(13) || Chr(10) || '这可能是并发操作引起的，请重新读取校对数据。';
      Raise Err_Custom;
    End If;
  Else
    If To_Char(校对时间_In, 'YYYY-MM-DD HH24:MI') < To_Char(v_开始时间, 'YYYY-MM-DD HH24:MI') Then
      v_Error := '医嘱"' || Getadvicetext(医嘱id_In) || '"的校对时间不能小于开始执行时间 ' || To_Char(v_开始时间, 'YYYY-MM-DD HH24:MI') || '。' ||
                 Chr(13) || Chr(10) || '这可能是并发操作引起的，请重新读取校对数据。';
      Raise Err_Custom;
    End If;
  End If;

  --如果要求签名，检查校对时是否有签名(并发取消签名)
  If 状态_In = 3 Then
    Select Zl_Fun_Getsignpar(Decode(v_前提id, Null, 1, 3), v_开嘱科室id) Into v_Count From Dual;
    If v_Count = 1 Then
      --证书停用或未注册证书不进入签名环节只判断一条数据即可
      For C In (Select a.是否停用
                From 人员证书记录 A, 人员表 B
                Where a.人员id = b.Id And b.姓名 = v_开嘱医生
                Order By a.注册时间 Desc) Loop
        If Nvl(c.是否停用, 0) = 0 Then
          Select Count(*)
          Into v_Count
          From 病人医嘱状态 A
          Where 操作类型 = 1 And 医嘱id = 医嘱id_In And
                (签名id Is Null And Exists
                 (Select 1
                  From 人员表 R, 人员性质说明 X
                  Where r.Id = x.人员id And r.姓名 = a.操作人员 And x.人员性质 = '护士') And Not Exists
                 (Select 1
                  From 人员表 R, 人员性质说明 Y
                  Where r.Id = y.人员id And r.姓名 = a.操作人员 And y.人员性质 = '医生') Or 签名id Is Not Null Or a.操作人员 <> v_开嘱医生);
          If Nvl(v_Count, 0) = 0 Then
            v_Error := '医嘱"' || Getadvicetext(医嘱id_In) || '"还没有电子签名，不能通过校对。';
            Raise Err_Custom;
          End If;
        End If;
        Exit;
      End Loop;
    End If;
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

  --因为可能同时：新开->自动校对->互斥自动停止,因此分别-2,-1秒
  Select Sysdate - 1 / 60 / 60 / 24 Into v_Date From Dual;

  Update 病人医嘱记录
  Set 医嘱状态 = 状态_In, 校对护士 = v_人员姓名, 校对时间 = 校对时间_In
  Where ID = 医嘱id_In Or 相关id = 医嘱id_In;

  Insert Into 病人医嘱状态
    (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
    Select ID, 状态_In, v_人员姓名, v_Date, 校对说明_In From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;

  --校对通过时的其它处理
  If 状态_In = 3 Then
    --自动校对时，自动填写缺省的计价内容
    If Nvl(自动校对_In, 0) = 1 Then
      --1.变价的计价项目,如果最低限价不为0,则缺省为最低限价,否则不加入;可再手工计价.
      --2.对于非药嘱药品和在用卫材未定执行科室,发送时会取缺省的,可再手工设置。
      For r_Price In c_Price Loop
        --取(检验)医嘱的管码和材料,采集方式以检验项目的为准
        v_材料id := Null;
        If r_Price.诊疗类别 = 'E' And r_Price.操作类型 = '6' Then
          Begin
            Select c.材料id
            Into v_材料id
            From 病人医嘱记录 A, 诊疗项目目录 B, 采血管类型 C
            Where a.诊疗项目id = b.Id And b.试管编码 = c.编码 And a.相关id = r_Price.Id And Rownum = 1;
          Exception
            When Others Then
              Null;
          End;
        Elsif r_Price.诊疗类别 = 'C' And r_Price.试管编码 Is Not Null Then
          Begin
            Select 材料id Into v_材料id From 采血管类型 Where 编码 = r_Price.试管编码;
          Exception
            When Others Then
              Null;
          End;
        End If;
      
        --判断处理检验试管费用的收取
        If (Nvl(r_Price.收费方式, 0) = 1 And r_Price.收费类别 = '4' And r_Price.收费项目id = Nvl(v_材料id, 0) Or
           Not (Nvl(r_Price.收费方式, 0) = 1 And r_Price.收费类别 = '4' And Nvl(v_材料id, 0) <> 0)) Then
          Insert Into 病人医嘱计价
            (医嘱id, 收费细目id, 数量, 单价, 从项, 执行科室id, 费用性质, 收费方式)
          Values
            (r_Price.Id, r_Price.收费项目id, r_Price.收费数量, r_Price.单价, r_Price.从属项目, Null, r_Price.费用性质, r_Price.收费方式);
        End If;
      End Loop;
    End If;
  
    --自由录入的临嘱医嘱标记为停止
    If Nvl(v_期效, 0) = 1 And v_诊疗项目id Is Null Then
      Update 病人医嘱记录
      Set 医嘱状态 = 8, 停嘱时间 = 校对时间_In, 停嘱医生 = v_开嘱医生
      Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
    
      Insert Into 病人医嘱状态
        (医嘱id, 操作类型, 操作人员, 操作时间)
        Select ID, 8, v_人员姓名, Sysdate From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
      End If;  
      
    --判断是否开启叮嘱需要执行
    v_叮嘱执行:= zl_GetSysParameter(288);
    if v_叮嘱执行=1 and v_诊疗项目id Is Null then
        Insert Into 病人医嘱发送
          (医嘱id, 发送号, 记录性质, NO, 记录序号, 发送数次, 发送人, 发送时间, 执行状态, 执行部门id, 计费状态, 首次时间, 末次时间)
        Values
          (医嘱id_In, NextNO('10','0','','1'), '2',NextNO('14','0','','1'), '1', '1', v_人员姓名, sysdate, '0', v_执行科室id,'0',sysdate,sysdate);    
     End If;

    --将同一自动停止互斥组中的病人其它医嘱停止(如果尚未停止)
    For r_Exclude In c_Exclude Loop
      Select Decode(Sign(r_Exclude.开始执行时间 - v_开始时间), 1, r_Exclude.开始执行时间, v_开始时间)
      Into v_终止时间
      From Dual;
      Select Decode(Sign(r_Exclude.执行终止时间 - v_开始时间), -1, r_Exclude.执行终止时间, v_开始时间)
      Into v_终止时间
      From Dual;
      Zl_病人医嘱记录_停止(r_Exclude.医嘱id, v_终止时间, v_开嘱医生, 1);
      v_Stopadviceids := v_Stopadviceids || ',' || r_Exclude.医嘱id;
    End Loop;
  
    --对一些特殊医嘱的处理
    If v_诊疗类别 = 'H' And v_操作类型 = '1' And Nvl(v_期效, 0) = 0 Then
      --校对护理等级时,同步更改病人护理等级
      If Nvl(v_婴儿, 0) = 0 Then
        --病人当前应处于正常住院状态
        v_Temp := Null;
        Begin
          Select Decode(状态, 1, '等待入科', 2, '正在转科', 3, '已预出院', Null)
          Into v_Temp
          From 病案主页
          Where 病人id = v_病人id And 主页id = v_主页id;
        Exception
          When Others Then
            Null;
        End;
        If v_Temp Is Not Null Then
          v_Error := '病人当前处于' || v_Temp || '状态,医嘱"' || v_医嘱内容 || '"不能通过校对。';
          Raise Err_Custom;
        End If;
      
        Begin
          --根据收费对照处理，当前医嘱计价表还没有填写
          --未设置时,不处理；相同时,不处理；有多个时,只取一个。
          Select a.收费项目id
          Into v_护理等级id
          From 诊疗收费关系 A, 收费项目目录 B
          Where a.收费项目id = b.Id And b.类别 = 'H' And Nvl(b.项目特性, 0) <> 0 And a.诊疗项目id = v_诊疗项目id And Rownum = 1 And
                Not Exists
           (Select 1 From 病案主页 Where 病人id = v_病人id And 主页id = v_主页id And 护理等级id = a.收费项目id);
        Exception
          When Others Then
            Null;
        End;
      End If;
    
      --变动记录的时间加上秒，以便回退操作时区分同一分种的校对、停止等操作
      v_开始时间 := To_Date(To_Char(v_开始时间, 'yyyy-mm-dd hh24:mi') || To_Char(Sysdate, 'ss'), 'yyyy-mm-dd hh24:mi:ss');
      If v_护理等级id Is Not Null Then
        Zl_病人变动记录_Nurse(v_病人id, v_主页id, v_护理等级id, v_开始时间, v_人员编号, v_人员姓名);
      End If;
    
      --并停止其它护理等级医嘱(护理等级应该都为"持续性"长嘱,且只有一个未停)
      For r_Nurse In c_Nurse Loop
        Select Decode(Sign(r_Nurse.开始执行时间 - v_开始时间), 1, r_Nurse.开始执行时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Select Decode(Sign(r_Nurse.执行终止时间 - v_开始时间), -1, r_Nurse.执行终止时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Zl_病人医嘱记录_停止(r_Nurse.医嘱id, v_终止时间, v_开嘱医生, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Nurse.医嘱id;
      End Loop;
    Elsif v_诊疗类别 = 'Z' And v_操作类型 In ('9', '10') And Nvl(v_期效, 0) = 0 And Nvl(v_婴儿, 0) = 0 Then
      --病重病危医嘱：9-病重;10-病危
      --停止相同医嘱
      For r_Patistate In c_Patistate Loop
        Select Decode(Sign(r_Patistate.开始执行时间 - v_开始时间), 1, r_Patistate.开始执行时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Select Decode(Sign(r_Patistate.执行终止时间 - v_开始时间), -1, r_Patistate.执行终止时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Zl_病人医嘱记录_停止(r_Patistate.医嘱id, v_终止时间, v_开嘱医生, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Patistate.医嘱id;
      End Loop;
    
      --产生病情变动
      Open c_Oldinfo; --必须在处理之前先打开
      Fetch c_Oldinfo
        Into r_Oldinfo;
      Open c_Endinfo;
      Fetch c_Endinfo
        Into r_Endinfo;
      If c_Endinfo%RowCount = 0 Then
        Close c_Endinfo;
        v_Error := '未发现该病人当前有效的变动记录！';
        Raise Err_Custom;
      End If;
      Select Count(*)
      Into v_Count
      From 病人变动记录
      Where 病人id = v_病人id And 主页id = v_主页id And 开始时间 Is Null And 终止时间 Is Null;
      If v_Count > 0 Then
        v_Error := '病人当前处于转科状态，请先办理转科确认或者取消转科状态。';
        Raise Err_Custom;
      End If;
    
      Update 病案主页
      Set 当前病况 = Decode(v_操作类型, '9', '重', '10', '危')
      Where 病人id = v_病人id And 主页id = v_主页id;
    
      --取消上次变动
      If r_Oldinfo.终止时间 Is Not Null Then
        v_变动终止时间 := r_Oldinfo.终止时间;
        v_变动终止原因 := r_Oldinfo.终止原因;
        v_变动终止人员 := r_Oldinfo.终止人员;
        --取消上次变动
        Update 病人变动记录
        Set 终止时间 = v_开始时间, 终止原因 = 13, 终止人员 = v_人员姓名, 上次计算时间 = Null
        Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 = v_变动终止时间 And 终止原因 = v_变动终止原因;
        --更新将来的记录如果有停止到将来的则删除上次计算时间
        Update 病人变动记录
        Set 病情 = Decode(v_操作类型, '9', '重', '10', '危'), 上次计算时间 = Null
        Where 病人id = v_病人id And 主页id = v_主页id And 开始时间 > v_开始时间;
      Else
        Update 病人变动记录
        Set 终止时间 = v_开始时间, 终止原因 = 13, 终止人员 = v_人员姓名,
            上次计算时间 = Decode(Sign(Nvl(上次计算时间, v_开始时间) - v_开始时间), 1, Null, 上次计算时间)
        Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null;
      End If;
    
      While c_Oldinfo%Found Loop
        Insert Into 病人变动记录
          (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情, 操作员编号, 操作员姓名,
           终止时间, 终止原因, 终止人员)
        Values
          (病人变动记录_Id.Nextval, v_病人id, v_主页id, v_开始时间, 13, r_Oldinfo.附加床位, r_Oldinfo.病区id, r_Oldinfo.科室id,
           r_Oldinfo.护理等级id, r_Oldinfo.床位等级id, r_Oldinfo.床号, r_Oldinfo.责任护士, r_Oldinfo.经治医师, r_Oldinfo.主治医师,
           r_Oldinfo.主任医师, Decode(v_操作类型, '9', '重', '10', '危'), v_人员编号, v_人员姓名, v_变动终止时间, v_变动终止原因, v_变动终止人员);
      
        Fetch c_Oldinfo
          Into r_Oldinfo;
      End Loop;
    
      Close c_Oldinfo;
      Close c_Endinfo;
    Elsif v_诊疗类别 = 'Z' And v_操作类型 = '12' And Nvl(v_期效, 0) = 0 And Nvl(v_婴儿, 0) = 0 Then
      --记录入出量的医嘱，互斥
      For r_Patiio In c_Patiio Loop
        Select Decode(Sign(r_Patiio.开始执行时间 - v_开始时间), 1, r_Patiio.开始执行时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Select Decode(Sign(r_Patiio.执行终止时间 - v_开始时间), -1, r_Patiio.执行终止时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Zl_病人医嘱记录_停止(r_Patiio.医嘱id, v_终止时间, v_开嘱医生, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Patiio.医嘱id;
      End Loop;
    Elsif (v_诊疗类别 = 'Z' And v_操作类型 In ('3', '4', '5', '6', '11', '14') And
          (v_操作类型 <> '14' Or v_操作类型 = '14' And v_执行标记 = 1)) Or (v_诊疗类别 = 'F' And v_执行标记 = 1) Then
      v_Count := 0;
      If v_操作类型 = '4' Or v_操作类型 = '14' Or v_诊疗类别 = 'F' Then
        --保持与以前校对时相同的处理
        If Nvl(v_婴儿, 0) = 0 Then
          v_Count := 1;
        End If;
      Else
        --这几个特殊医嘱在校对中停止医嘱是新加的内容，保持与发送中相同的处理
        v_Count := 1;
        If Nvl(v_婴儿, 0) = 0 Then
          v_婴儿 := -1;
        Else
          v_婴儿 := Nvl(v_婴儿, 0);
        End If;
      End If;
      If v_Count = 1 Then
        If v_诊疗类别 = 'F' And v_执行标记 = 1 Then
          --在手术当天(取整)停止
          v_开始时间 := Trunc(To_Date(v_标本部位, 'yyyy-mm-dd hh24:mi:ss'));
        End If;
      
        --几个特殊医嘱校对时停止前面的长嘱,在医嘱开始时终止：3-转科;4-术后;5-出院;6-转院,11-死亡,14-术前
        For r_Needstop In c_Needstop(v_病人id, v_主页id, v_婴儿, v_开始时间) Loop
          Select Decode(Sign(开始执行时间 - v_开始时间), 1, 开始执行时间, v_开始时间)
          Into v_停止时间
          From 病人医嘱记录
          Where ID = r_Needstop.Id;
          Update 病人医嘱记录
          Set 医嘱状态 = 8, 执行终止时间 = v_停止时间, 停嘱时间 = 校对时间_In, 停嘱医生 = v_开嘱医生
          Where ID = r_Needstop.Id;
        
          Insert Into 病人医嘱状态
            (医嘱id, 操作类型, 操作人员, 操作时间)
            Select ID, 8, v_人员姓名, 校对时间_In From 病人医嘱记录 Where ID = r_Needstop.Id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Needstop.Id;
        End Loop;
        --已停止未确认的长嘱,终止时间在医嘱开始后的,调前其终止时间(同时多个特殊医嘱的情况)
        For r_Havestop In c_Havestop(v_病人id, v_主页id, v_婴儿, v_开始时间) Loop
          Update 病人医嘱记录
          Set 执行终止时间 = Decode(Sign(开始执行时间 - v_开始时间), 1, 开始执行时间, v_开始时间), 停嘱时间 = 校对时间_In, 停嘱医生 = v_开嘱医生
          Where ID = r_Havestop.Id;
        
          --不修改停止医嘱的操作人员，因为停止时，医生可能已进行电子签名
          Update 病人医嘱状态 Set 操作时间 = 校对时间_In Where 医嘱id = r_Havestop.Id And 操作类型 = 8;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Havestop.Id;
        End Loop;
        --处理长期备用医嘱(没有执行（发送）过的标记未用）
        Update 病人医嘱记录
        Set 执行标记 = -1
        Where 病人id = v_病人id And 主页id = v_主页id And 医嘱期效 = 0 And 执行频次 = '必要时' And 上次执行时间 Is Null And 医嘱状态 In (3, 5, 6, 7) And
              执行标记 <> -1;
        --如果是转院转科死亡出院医嘱同时处理临时备用医嘱。
        If v_操作类型 In ('3', '5', '6', '11') Then
          Update 病人医嘱记录
          Set 执行标记 = -1
          Where 病人id = v_病人id And 主页id = v_主页id And 医嘱期效 = 1 And 执行频次 = '需要时' And 医嘱状态 = 3 And 执行标记 <> -1;
        End If;
      End If;
    Elsif v_诊疗类别 = 'Z' And v_操作类型 = '2' Then
      --对留观病人下达入院通知;
      --预约登记的条件：1.当前无预约,2.当前是门诊留观病人（在院时也允许，因为需要先预约,入院接收时检查了必须出院后才能接收）
      Select Count(*) Into v_Count From 病案主页 Where 病人id = v_病人id And Nvl(主页id, 0) = 0;
      If v_Count = 0 Then
        Select Count(*) Into v_Count From 病案主页 Where 病人id = v_病人id And 主页id = v_主页id And 病人性质 <> 1;
      End If;
      If v_Count = 0 Then
        Open c_Pati(v_病人id);
        Fetch c_Pati
          Into r_Pati;
        Close c_Pati;
      
        v_入院方式 := Null;
        If v_紧急标志 = 1 Then
          v_入院方式 := '急诊';
        End If;
      
        Zl_入院病案主页_Insert(1, 0, r_Pati.病人id, r_Pati.住院号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别, r_Pati.出生日期,
                         r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份, r_Pati.身份证号, r_Pati.出生地点,
                         r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址, r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系,
                         r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位, r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行,
                         r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额, r_Pati.担保性质, v_执行科室id, Null, Null, v_入院方式, Null, Null,
                         v_开嘱医生, r_Pati.籍贯, r_Pati.区域, v_开始时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null, Null, Null,
                         Null, r_Pati.险类, v_人员编号, v_人员姓名, 0, Null, Null, 0);
      End If;
    End If;
    --医嘱停止消息的处理
    If v_Stopadviceids Is Not Null Then
      v_Stopadviceids := Substr(v_Stopadviceids, 2);
      Select Max(a.Id)
      Into n_标记
      From 病人医嘱记录 A
      Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.医嘱期效 = 0 And a.医嘱状态 = 8 And
            Nvl(a.执行标记, 0) <> -1 And a.病人来源 <> 3;
      If n_标记 Is Not Null Then
        Select Max(a.Id)
        Into n_Adviceid
        From 病人医嘱记录 A
        Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.紧急标志 = 1 And a.医嘱期效 = 0 And
              a.医嘱状态 = 8 And Nvl(a.执行标记, 0) <> -1 And a.病人来源 <> 3;
        If n_Adviceid Is Not Null Then
          n_Adviceid := n_标记;
          Select Nvl(Max(0), 2)
          Into n_标记
          From 业务消息清单 A
          Where a.病人id = v_病人id And a.就诊id = v_主页id And a.类型编码 = 'ZLHIS_CIS_002' And a.优先程度 = 2 And a.是否已阅 = 0;
        Else
          Select Nvl(Max(0), 1)
          Into n_标记
          From 业务消息清单 A
          Where a.病人id = v_病人id And a.就诊id = v_主页id And a.类型编码 = 'ZLHIS_CIS_002' And a.是否已阅 = 0;
        End If;
        If n_标记 > 0 Then
          For R In (Select a.病人性质 As 性质, a.出院科室id As 科室id, a.当前病区id As 病区id
                    From 病案主页 A
                    Where a.病人id = v_病人id And a.主页id = v_主页id) Loop
            Zl_业务消息清单_Insert(v_病人id, v_主页id, r.科室id, r.病区id, r.性质, '有新停止医嘱。', '0010', 'ZLHIS_CIS_002', n_Adviceid, n_标记,
                             0, Null, r.病区id);
          End Loop;
        End If;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_校对;
/

--136758:焦博,2019-01-11,调整Oracle过程Zl_住院记帐记录_Insert,在保存住院记帐数据时检查价格父号
CREATE OR REPLACE Procedure Zl_住院记帐记录_Insert
(
  No_In             住院费用记录.No%Type,
  序号_In           住院费用记录.序号%Type,
  病人id_In         住院费用记录.病人id%Type,
  主页id_In         住院费用记录.主页id%Type,
  标识号_In         住院费用记录.标识号%Type,
  姓名_In           住院费用记录.姓名%Type,
  性别_In           住院费用记录.性别%Type,
  年龄_In           住院费用记录.年龄%Type,
  床号_In           住院费用记录.床号%Type,
  费别_In           住院费用记录.费别%Type,
  病区id_In         住院费用记录.病人病区id%Type,
  科室id_In         住院费用记录.病人科室id%Type,
  加班标志_In       住院费用记录.加班标志%Type,
  婴儿费_In         住院费用记录.婴儿费%Type,
  开单部门id_In     住院费用记录.开单部门id%Type,
  开单人_In         住院费用记录.开单人%Type,
  从属父号_In       住院费用记录.从属父号%Type,
  收费细目id_In     住院费用记录.收费细目id%Type,
  收费类别_In       住院费用记录.收费类别%Type,
  计算单位_In       住院费用记录.计算单位%Type,
  保险项目否_In     住院费用记录.保险项目否%Type,
  保险大类id_In     住院费用记录.保险大类id%Type,
  保险编码_In       住院费用记录.保险编码%Type,
  付数_In           住院费用记录.付数%Type,
  数次_In           住院费用记录.数次%Type,
  附加标志_In       住院费用记录.附加标志%Type,
  执行部门id_In     住院费用记录.执行部门id%Type,
  价格父号_In       住院费用记录.价格父号%Type,
  收入项目id_In     住院费用记录.收入项目id%Type,
  收据费目_In       住院费用记录.收据费目%Type,
  标准单价_In       住院费用记录.标准单价%Type,
  应收金额_In       住院费用记录.应收金额%Type,
  实收金额_In       住院费用记录.实收金额%Type,
  统筹金额_In       住院费用记录.统筹金额%Type,
  发生时间_In       住院费用记录.发生时间%Type,
  登记时间_In       住院费用记录.登记时间%Type,
  药品摘要_In       药品收发记录.摘要%Type,
  划价_In           Number,
  操作员编号_In     住院费用记录.操作员编号%Type,
  操作员姓名_In     住院费用记录.操作员姓名%Type,
  多病人单_In       Number := 0,
  类别id_In         药品单据性质.类别id%Type := Null,
  记帐单id_In       住院费用记录.记帐单id%Type := Null,
  费用摘要_In       住院费用记录.摘要%Type := Null,
  是否急诊_In       住院费用记录.是否急诊%Type := 0,
  医嘱序号_In       住院费用记录.医嘱序号%Type := Null,
  频次_In           药品收发记录.频次%Type := Null,
  单量_In           药品收发记录.单量%Type := Null,
  用法_In           药品收发记录.用法%Type := Null, --用法[|煎法]
  期效_In           药品收发记录.扣率%Type := Null,
  计价特性_In       药品收发记录.扣率%Type := Null,
  简单记帐_In       Number := 0,
  费用类型_In       住院费用记录.费用类型%Type := Null,
  医技补临床费用_In Number := 0,
  领药部门id_In     药品收发记录.对方部门id%Type := Null,
  中药形态_In       住院费用记录.结论%Type := Null,
  医疗小组id_In     住院费用记录.医疗小组id%Type := -1,
  备货材料_In       Number := 0,
  批次_In           药品收发记录.批次%Type := Null
) As
  --功能：新收一张住院记帐单据
  --参数：
  --   药品摘要_IN:存放医嘱中的附加说明或修改保存新单据时用。目前仅用于存放于药品收发记录的摘要中。
  --         原单据(记录状态=2)记录修改产生的新单据号。
  --         新单据(记录状态=1)记录所修改的原单据号。
  --   划价-是否属于住院划价。
  --   医技补临床费用_in:医技站补费时,如果开单科室为临床科室则划价人和记帐人填写为不同,用于销帐申请时区分填写审核科室
  --   备货材料_IN: 医技站的卫生材料备货处理方式:0-正常记帐单;1-备货材料记帐
  --   批次_In:当备货材料_IN=1时有效.传入指定的卫生材料的批次
  v_费用id 住院费用记录.Id%Type;
  v_优先级 未发药品记录.优先级%Type;

  --药房分批、时价药品--
  ------------------------------------------------------------
  --该游标用于分批药品数量分解
  Cursor c_Stock
  (
    n_Outmode Number,
    n_库房id  药品收发记录.库房id%Type
  ) Is
    Select 库房id, 药品id, 批次, 上次批号, 可用数量, 实际数量, 实际金额, 上次供应商id, 批准文号, 上次产地, 上次生产日期, 灭菌效期, 效期, 零售价, 商品条码, 内部条码
    From 药品库存
    Where 药品id = 收费细目id_In And 库房id = n_库房id And 性质 = 1 And (Nvl(批次, 0) = Nvl(批次_In, 0) Or Nvl(批次_In, 0) = 0) And
          (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And Nvl(可用数量, 0) > 0
    Order By Decode(n_Outmode, 1, 效期, Null), Decode(n_Outmode, 2, 上次批号, Null), Nvl(批次, 0);
  r_Stock c_Stock%RowType;

  --属性
  n_分批 药品规格.药房分批%Type;
  n_时价 收费项目目录.是否变价%Type;
  v_名称 收费项目目录.名称%Type;
  --临时变量
  n_总数量   Number;
  n_当前数量 Number;
  n_总金额   Number;
  n_当前单价 Number;
  --药品收发记录
  n_批次       药品收发记录.批次%Type;
  n_序号       药品收发记录.序号%Type;
  n_扣率       药品收发记录.扣率%Type;
  n_领药部门id 药品收发记录.对方部门id%Type;
  n_供药单位id 药品收发记录.供药单位id%Type;
  v_商品条码   药品收发记录.商品条码%Type;
  v_内部条码   药品收发记录.内部条码%Type;
  d_效期       药品收发记录.效期%Type;
  d_灭菌效期   药品收发记录.灭菌效期%Type;
  d_灭菌日期   药品收发记录.灭菌日期%Type;
  d_生产日期   药品收发记录.生产日期%Type;
  n_虚拟库房id 药品收发记录.库房id%Type;
  v_产地       药品收发记录.产地%Type;
  v_批号       药品收发记录.批号%Type;
  v_批准文号   药品收发记录.批准文号%Type;
  v_其他出库no 药品收发记录.No%Type;
  n_Aval       药品库存.可用数量%Type;
  v_部门名称   部门表.名称%Type;
  ------------------------------------------------------------
  v_用法       药品收发记录.用法%Type;
  v_煎法       药品收发记录.外观%Type;
  n_单价小数   Number;
  n_医疗小组id 住院费用记录.医疗小组id%Type;
  n_库房id     药品库存.库房id%Type;
  n_修正库房id 药品库存.库房id%Type;
  n_Outmode    Number(1);
  v_Dec        Number;
  v_Count      Number;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  n_出库序号       药品收发记录.序号%Type;
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);
  v_操作员编号     人员表.编号%Type;
  v_操作员姓名     人员表.姓名%Type;
  v_Temp           Varchar2(255);
Begin
  v_操作员编号 := 操作员编号_In;
  v_操作员姓名 := 操作员姓名_In;
  If v_操作员编号 Is Null Then
    v_Temp := Zl_Identity(1);
    If Not (Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_') Then
      v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
      v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      v_操作员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
      v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      v_操作员姓名 := v_Temp;
    End If;
  End If;
  --根据执行库房取虚拟库房ID
  Begin
    Select 虚拟库房id Into n_虚拟库房id From 虚拟库房对照 Where 科室id = 执行部门id_In And Rownum <= 1;
  Exception
    When Others Then
      n_虚拟库房id := 0;
  End;
  If Nvl(批次_In, 0) <> 0 Then
    Select Nvl(Sum(可用数量), 0)
    Into n_Aval
    From 药品库存
    Where 药品id = 收费细目id_In And 批次 = 批次_In And 库房id = 执行部门id_In;
    If n_Aval <= 0 Then
      n_修正库房id := n_虚拟库房id;
    End If;
  End If;

  If Nvl(备货材料_In, 0) = 1 And 收费类别_In = '4' Then
    If Nvl(n_虚拟库房id, 0) = 0 Then
      Begin
        Select 名称 Into v_Err_Msg From 部门表 Where ID = 执行部门id_In;
      Exception
        When Others Then
          v_Err_Msg := '';
      End;
      v_Err_Msg := '执行部门"' || Nvl(v_Err_Msg, '') || '"未设置虚拟部门,请在卫材参数设置中设置.';
      Raise Err_Item;
    End If;
  End If;
  If Nvl(多病人单_In, 0) = 1 Or Nvl(序号_In, 0) = 1 Then
    --记帐表,全部检查,如果是记帐单只检查第一条,其他不检查
  
    n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
    n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
    If n_病人审核方式 = 1 Or n_未入科禁止记账 = 1 Then
      Begin
        Select 审核标志, 状态
        Into n_审核标志, n_住院状态
        From 病案主页
        Where 病人id = Nvl(病人id_In, 0) And 主页id = Nvl(主页id_In, 0);
      Exception
        When Others Then
          n_审核标志 := 0;
          n_住院状态 := 0;
      End;
      If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then
        v_Err_Msg := '病人未入科,禁止对病人相关费用的操作!';
        Raise Err_Item;
      End If;
    
      If n_病人审核方式 = 1 Then
      
        If Nvl(n_审核标志, 0) = 1 Then
          v_Err_Msg := '该病人目前正在审核费用,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
        If Nvl(n_审核标志, 0) = 2 Then
          v_Err_Msg := '该病人目前已经完成了费用审核,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;
  ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
  --领药部门确认规则:
  -- 1.传入 :领药部门ID_IN,直接以传入的为准
  -- 2.领药部门ID_IN=NULL的情况, 如果”开单科室=病人科室”，则填为”病人病区”，如果”开单科室<>病人科室”，则填写为”开单科室”
  If Nvl(领药部门id_In, 0) = 0 Then
    If Nvl(科室id_In, 0) = Nvl(开单部门id_In, 0) Then
      --如果”开单科室=病人科室”，则填为”病人病区”(如果没有入科,即病匹为空这种情况,则以病人科室为准,由于一般这种情况较少(护土开单),因此,这种情况应该不会存在)
      n_领药部门id := Nvl(病区id_In, 0);
      If Nvl(n_领药部门id, 0) = 0 Then
        n_领药部门id := 科室id_In;
      End If;
    Else
      --如果”开单科室<>病人科室”，则填写为”开单科室”
      n_领药部门id := 开单部门id_In;
    End If;
  Else
    n_领药部门id := 领药部门id_In;
  End If;
  --需要检查0这种情况,回为有关联
  If Nvl(n_领药部门id, 0) = 0 Then
    n_领药部门id := Null;
  End If;
  ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

  --药品用法煎法分解
  If 用法_In Is Not Null Then
    If Instr(用法_In, '|') > 0 Then
      v_用法 := Substr(用法_In, 1, Instr(用法_In, '|') - 1);
      v_煎法 := Substr(用法_In, Instr(用法_In, '|') + 1);
    Else
      v_用法 := 用法_In;
    End If;
  End If;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into v_Dec, n_单价小数
  From Dual;

  --住院费用记录
  Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
  n_医疗小组id := 医疗小组id_In;

  If Nvl(医疗小组id_In, 0) < 0 Then
    n_医疗小组id := Zl_医疗小组_Get(开单部门id_In, 开单人_In, 病人id_In, 主页id_In, 发生时间_In);
  End If;

  --目前有很多用户反馈，开医嘱时，价格父号传入错误，因此特加入了以下限制，随后医嘱这边查出问题后，再取消该限制.
  If Nvl(价格父号_In, 0) <> 0 Then
    v_Count := 0;
    Select Count(1)
    Into v_Count
    From 住院费用记录
    Where 记录性质 = 2 And NO = No_In And (序号 = 价格父号_In Or 价格父号 = 价格父号_In) And 收入项目id = 收入项目id_In;
    If Nvl(v_Count, 0) <> 0 Then
      v_Err_Msg := '第 ' || 序号_In || ' 行的收费项目的价格父号存在异常，请与管理员联系.';
      Raise Err_Item;
    End If;
  End If;

  Insert Into 住院费用记录
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id,
     计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id,
     开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 医嘱序号, 结论, 医疗小组id)
  Values
    (v_费用id, 2, No_In, Decode(划价_In, 1, 0, 1), 序号_In, Decode(从属父号_In, 0, Null, 从属父号_In),
     Decode(价格父号_In, 0, Null, 价格父号_In), 多病人单_In, 2, 病人id_In, 主页id_In, Decode(标识号_In, 0, Null, 标识号_In), 姓名_In, 性别_In,
     年龄_In, 床号_In, Decode(病区id_In, 0, Null, 病区id_In), Decode(科室id_In, 0, Null, 科室id_In), 费别_In, 收费类别_In, 收费细目id_In,
     计算单位_In, 保险项目否_In, 保险大类id_In, 保险编码_In, 费用类型_In, Decode(Nvl(简单记帐_In, 0), 0, Null, 收费类别_In), 付数_In, 数次_In, 加班标志_In,
     附加标志_In, 婴儿费_In, 收入项目id_In, 收据费目_In, 标准单价_In, 应收金额_In, 实收金额_In, 统筹金额_In, 1, 开单部门id_In, 开单人_In, 发生时间_In, 登记时间_In,
     执行部门id_In, 0, Decode(划价_In, 1, v_操作员姓名, Decode(医技补临床费用_In, 1, '补临床费', Null)), Decode(划价_In, 1, Null, v_操作员编号),
     Decode(划价_In, 1, Null, v_操作员姓名), 记帐单id_In, 费用摘要_In, 是否急诊_In, 医嘱序号_In, 中药形态_In, n_医疗小组id);

  Select Max(使用限量 - Nvl(已用数量, 0))
  Into n_当前数量
  From 病人审批项目
  Where 病人id = 病人id_In And 主页id = 主页id_In And 项目id = 收费细目id_In And Nvl(使用限量, 0) <> 0;
  If 付数_In * 数次_In <= Nvl(n_当前数量, 0) Then
    Update 病人审批项目
    Set 已用数量 = Nvl(已用数量, 0) + 付数_In * 数次_In
    Where 病人id = 病人id_In And 主页id = 主页id_In And 项目id = 收费细目id_In And Nvl(使用限量, 0) <> 0;
  Elsif Not n_当前数量 Is Null Then
    v_Err_Msg := '第 ' || 序号_In || ' 行输入的数次超过了批准的可用数量' || n_当前数量 || '.'; --简化为不转换,直接以售价单位提示.
    Raise Err_Item;
  End If;

  --相关汇总表的处理
  If Nvl(划价_In, 0) = 0 Then
    --病人余额
    Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) + 实收金额_In Where 病人id = 病人id_In And 类型 = 2 And 性质 = 1;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 类型, 性质, 费用余额, 预交余额) Values (病人id_In, 2, 1, 实收金额_In, 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + 实收金额_In
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(病区id_In, 0) And
          Nvl(病人科室id, 0) = Nvl(科室id_In, 0) And Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And
          Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And 来源途径 + 0 = 2;
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, 主页id_In, 病区id_In, 科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 2, 实收金额_In);
    End If;
  End If;

  --药品和卫生材料部分
  v_Count := 0; --@@@
  If 收费类别_In = '4' Then
    --跟踪在用的卫材才处理
    Select 跟踪在用 Into v_Count From 材料特性 Where 材料id = 收费细目id_In;
  End If;
  If 收费类别_In In ('5', '6', '7') Or (收费类别_In = '4' And Nvl(v_Count, 0) = 1) Then
    If 收费类别_In = '4' Then
      Select Nvl(a.在用分批, 0), Nvl(b.是否变价, 0), b.名称
      Into n_分批, n_时价, v_名称
      From 材料特性 A, 收费项目目录 B
      Where a.材料id = b.Id And b.Id = 收费细目id_In;
    
      --卫材分批出库方式
      Select Zl_To_Number(Nvl(zl_GetSysParameter(156), 0)) Into n_Outmode From Dual;
    Else
      Select Nvl(a.药房分批, 0), Nvl(b.是否变价, 0), b.名称
      Into n_分批, n_时价, v_名称
      From 药品规格 A, 收费项目目录 B
      Where a.药品id = b.Id And b.Id = 收费细目id_In;
    
      --药品分批出库方式
      Select Zl_To_Number(Nvl(zl_GetSysParameter(150), 0)) Into n_Outmode From Dual;
    End If;
  
    n_总数量 := 付数_In * 数次_In;
    n_总金额 := 0;
    If Nvl(备货材料_In, 0) = 1 And 收费类别_In = '4' Then
      n_库房id := n_虚拟库房id;
      Open c_Stock(n_Outmode, n_虚拟库房id);
    Else
      If Nvl(n_修正库房id, 0) <> 0 Then
        n_库房id := n_修正库房id;
        Open c_Stock(n_Outmode, n_修正库房id);
      Else
        n_库房id := 执行部门id_In;
        Open c_Stock(n_Outmode, 执行部门id_In);
      End If;
    End If;
  
    While n_总数量 <> 0 Loop
      Fetch c_Stock
        Into r_Stock;
      If c_Stock%NotFound Then
        --第一次就没有库存,分批或时价都不允许(包含备货卫材)。
        --分批药品数量分解不完,也就是库存不足。
        If n_分批 = 1 Or n_时价 = 1 Or (Nvl(备货材料_In, 0) = 1 And 收费类别_In = '4') Then
          Close c_Stock;
          If 医嘱序号_In Is Null Then
            If 收费类别_In = '4' Then
              If Nvl(备货材料_In, 0) = 1 And Not (n_分批 = 1 Or n_时价 = 1) Then
                v_Err_Msg := '第 ' || 序号_In || ' 行的卫生材料"' || v_名称 || '"没有足够的材料库存,不能进行备货记帐！';
              Else
                v_Err_Msg := '第 ' || 序号_In || ' 行的分批或时价卫生材料"' || v_名称 || '"没有足够的材料库存' || Case
                               When Nvl(备货材料_In, 0) = 0 Then
                                '！'
                               Else
                                ',不能进行备货记帐！'
                             End;
              End If;
            Else
              v_Err_Msg := '第 ' || 序号_In || ' 行的分批或时价药品"' || v_名称 || '"没有足够的药品库存！';
            End If;
          Else
            If 收费类别_In = '4' Then
              If Nvl(备货材料_In, 0) = 1 And Not (n_分批 = 1 Or n_时价 = 1) Then
                v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现卫生材料"' || v_名称 || '"没有足够的材料库存,不能进行备货记帐！';
              Else
                v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现分批或时价卫生材料"' || v_名称 || '"没有足够的材料库存' || Case
                               When Nvl(备货材料_In, 0) = 0 Then
                                '！'
                               Else
                                ',不能进行备货记帐！'
                             End;
              End If;
            Else
              v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现分批或时价药品"' || v_名称 || '"没有足够的药品库存！';
            End If;
          End If;
          Raise Err_Item;
        End If;
      Elsif (n_分批 = 1 And Nvl(r_Stock.批次, 0) = 0) Or (n_分批 = 0 And Nvl(r_Stock.批次, 0) <> 0) Then
        Close c_Stock;
        If 医嘱序号_In Is Null Then
          If 收费类别_In = '4' Then
            v_Err_Msg := '第 ' || 序号_In || ' 行卫生材料"' || v_名称 || '"的分批属性与库存记录不相符,请检查材料数据的正确性！';
          Else
            v_Err_Msg := '第 ' || 序号_In || ' 行药品"' || v_名称 || '"的分批属性与库存记录不相符,请检查药品数据的正确性！';
          End If;
        Else
          If 收费类别_In = '4' Then
            v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现卫生材料"' || v_名称 || '"的分批属性与库存记录不相符,请检查材料数据的正确性！';
          Else
            v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现药品"' || v_名称 || '"的分批属性与库存记录不相符,请检查药品数据的正确性！';
          End If;
        End If;
        Raise Err_Item;
      End If;
    
      If c_Stock%Found Then
        If Nvl(r_Stock.实际数量, 0) = 0 And (n_总数量 > 0 Or n_时价 = 1) Then
          --实际数量为零时，不正常，不允许出库
          --实际数量不为零，金额为零，可能是正常的零价格管理。
          --负数的情况相当于入库,这种情况应是允许的；但时价需要计算价格，必须要有实际数量。
          Close c_Stock;
          If 医嘱序号_In Is Null Then
            If 收费类别_In = '4' Then
              v_Err_Msg := '第 ' || 序号_In || ' 行的卫生材料"' || v_名称 || '"当前无库存实际数量，可能存在尚未退料的记录，当前不能出库。';
            Else
              v_Err_Msg := '第 ' || 序号_In || ' 行药品"' || v_名称 || '"当前无库存实际数量，可能存在尚未退药的记录，当前不能出库。';
            End If;
          Else
            If 收费类别_In = '4' Then
              v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现卫生材料"' || v_名称 || '"当前无库存实际数量，可能存在尚未退料的记录，当前不能出库。';
            Else
              v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现药品"' || v_名称 || '"当前无库存实际数量，可能存在尚未退药的记录，当前不能出库。';
            End If;
          End If;
          Raise Err_Item;
        End If;
      End If;
    
      --确定本次分解数量
      If n_分批 = 1 Or n_时价 = 1 Then
        --对于不分批的时价只可能分解一次,分解不完上面判断了.它分解是为了计算单价.
        --每次分解取小者,库存不够分解不完在上面判断.
        If n_总数量 <= Nvl(r_Stock.可用数量, 0) Then
          n_当前数量 := n_总数量;
        Else
          n_当前数量 := Nvl(r_Stock.可用数量, 0);
        End If;
        If n_时价 = 1 Then
          n_当前单价 := Round(Nvl(r_Stock.零售价, Nvl(r_Stock.实际金额 / r_Stock.实际数量, 0)), n_单价小数);
        Elsif n_分批 = 1 Then
          n_当前单价 := 标准单价_In;
        End If;
      Else
        --普通药品
        --不管够不够,程序中已根据参数判断
        If Nvl(备货材料_In, 0) = 1 And 收费类别_In = '4' Then
          If n_总数量 > Nvl(r_Stock.可用数量, 0) Then
            --不分批, 但又是备货卫材方式出库的,则需要检查当前库存是否充足.
            v_Err_Msg := '第 ' || 序号_In || ' 行的卫生材料"' || v_名称 || '"没有足够的材料库存,不能进行备货记帐！';
            Raise Err_Item;
          End If;
        End If;
        n_当前数量 := n_总数量;
        n_当前单价 := 标准单价_In;
      End If;
    
      --药品库存(普通情况可能没有记录)
      If c_Stock%Found Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - n_当前数量
        Where 库房id = n_库房id And 药品id = 收费细目id_In And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1
        Returning 可用数量 Into v_Count;
      
        Zl_药品库存_可用数量异常处理(n_库房id, 收费细目id_In, Nvl(r_Stock.批次, 0));
      
        If n_分批 = 1 Or n_时价 = 1 Then
          If v_Count < 0 Then
            If 收费类别_In = '4' Then
              v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现卫生材料"' || v_名称 || '"当前库存实际数量不足，当前不能出库，可能是由于网络并发操作引起，请刷新后再试！。';
            Else
              v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现药品"' || v_名称 || '"当前库存实际数量不足，当前不能出库，可能是由于网络并发操作引起，请刷新后再试！。';
            End If;
            Raise Err_Item;
          End If;
        End If;
      Elsif n_库房id Is Not Null Then
        --只有不分批非时价药品可能库存不足出库
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - n_当前数量
        Where 库房id = n_库房id And 药品id = 收费细目id_In And Nvl(批次, 0) = 0 And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 可用数量, 商品条码, 内部条码)
          Values
            (n_库房id, 收费细目id_In, 1, -1 * n_当前数量, r_Stock.商品条码, r_Stock.内部条码);
        End If;
      
        Zl_药品库存_可用数量异常处理(n_库房id, 收费细目id_In, 0);
      End If;
    
      If Nvl(备货材料_In, 0) = 1 And 收费类别_In = '4' Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - n_当前数量
        Where 库房id = 执行部门id_In And 药品id = 收费细目id_In And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
      
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 商品条码, 内部条码, 效期, 上次批号, 上次产地, 上次供应商id, 上次生产日期, 批准文号)
          Values
            (执行部门id_In, 收费细目id_In, Nvl(r_Stock.批次, 0), 1, -1 * n_当前数量, r_Stock.商品条码, r_Stock.内部条码, r_Stock.效期,
             r_Stock.上次批号, r_Stock.上次产地, r_Stock.上次供应商id, r_Stock.上次生产日期, r_Stock.批准文号);
        End If;
      End If;
    
      --药品收发记录
      n_批次       := Null;
      v_批号       := Null;
      d_效期       := Null;
      v_产地       := Null;
      d_灭菌效期   := Null;
      d_灭菌日期   := Null;
      n_供药单位id := Null;
      d_生产日期   := Null;
      v_批准文号   := Null;
      If c_Stock%Found Then
        n_批次       := r_Stock.批次;
        v_批号       := r_Stock.上次批号;
        d_效期       := r_Stock.效期;
        v_产地       := r_Stock.上次产地;
        n_供药单位id := r_Stock.上次供应商id;
        d_生产日期   := r_Stock.上次生产日期;
        v_批准文号   := r_Stock.批准文号;
        v_商品条码   := r_Stock.商品条码;
        v_内部条码   := r_Stock.内部条码;
      
        --卫材灭菌效期:一次性材料且有效期
        If 收费类别_In = '4' Then
          v_Count := 0;
          Begin
            Select 灭菌效期 Into v_Count From 材料特性 Where Nvl(一次性材料, 0) = 1 And 材料id = 收费细目id_In;
          Exception
            When Others Then
              Null;
          End;
          If Nvl(v_Count, 0) > 0 Then
            d_灭菌效期 := r_Stock.灭菌效期;
            d_灭菌日期 := d_灭菌效期 - v_Count * 30;
          End If;
        End If;
      End If;
    
      Select Nvl(Max(序号), 0) + 1
      Into n_序号
      From 药品收发记录
      Where NO = No_In And 记录状态 = 1 And 单据 = Decode(多病人单_In, 1, 10, 9) + Decode(收费类别_In, '4', 16, 0);
    
      n_扣率 := Null;
      If 期效_In Is Not Null Or 计价特性_In Is Not Null Then
        n_扣率 := Nvl(期效_In, 0) || Nvl(计价特性_In, 0);
      End If;
    
      --分批药品,如果是只使用了一个批次,则要填写付数
      If n_分批 = 1 And n_当前数量 <> 付数_In * 数次_In Then
        v_Count := 1;
      Else
        v_Count := 0;
      End If;
    
      --修改的原单据号存放在摘要中
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 零售价, 零售金额, 摘要, 填制人,
         填制日期, 费用id, 频次, 单量, 用法, 外观, 扣率, 灭菌效期, 灭菌日期, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
      Values
        (药品收发记录_Id.Nextval, 1, Decode(多病人单_In, 1, 10, 9) + Decode(收费类别_In, '4', 16, 0), No_In, n_序号, 执行部门id_In,
         n_领药部门id, 类别id_In, -1, 收费细目id_In, n_批次, v_产地, v_批号, d_效期, Decode(v_Count, 1, 1, 付数_In),
         Decode(v_Count, 1, n_当前数量, n_当前数量 / 付数_In), Decode(v_Count, 1, n_当前数量, n_当前数量 / 付数_In), n_当前单价,
         Round(n_当前单价 * n_当前数量, v_Dec), 药品摘要_In, v_操作员姓名, 登记时间_In, v_费用id, 频次_In, 单量_In, v_用法, v_煎法, n_扣率, d_灭菌效期,
         d_灭菌日期, n_供药单位id, d_生产日期, v_批准文号, v_商品条码, v_内部条码);
    
      --产生其他出库单
      If 收费类别_In = '4' And (Nvl(备货材料_In, 0) = 1 Or Nvl(n_修正库房id, 0) <> 0) Then
        Begin
          Select Max(a.No), Max(a.序号)
          Into v_其他出库no, n_出库序号
          From 药品收发记录 A, 住院费用记录 B
          Where a.费用id = b.Id And a.单据 = 21 And b.No = No_In And b.记录性质 = 2;
        Exception
          When Others Then
            v_其他出库no := Null;
        End;
        If v_其他出库no Is Null Then
          v_其他出库no := Nextno(74, n_虚拟库房id, Null, 1);
        End If;
        If v_其他出库no Is Null Then
          v_Err_Msg := '在生成卫生材料的其他出库单时,获取相关的出库NO有误,请检查出库单的规则是否有误!';
          Raise Err_Item;
        End If;
        If Nvl(科室id_In, 0) <> 0 Then
          Select 名称 Into v_部门名称 From 部门表 Where ID = 科室id_In;
        End If;
        v_Err_Msg := LPad(' ', 4);
        v_Err_Msg := Substr('病人姓名:' || 姓名_In || v_Err_Msg || '性别:' || 性别_In || v_Err_Msg || '年龄' || 年龄_In || v_Err_Msg ||
                            '床号:' || 床号_In || v_Err_Msg || '住院号:' || Nvl(标识号_In, '') || v_Err_Msg || '病人科室:' || v_部门名称, 1,
                            100);
      
        n_出库序号 := Nvl(n_出库序号, 0) + 1;
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 零售价, 零售金额, 摘要, 填制人,
           填制日期, 费用id, 频次, 单量, 用法, 外观, 扣率, 灭菌效期, 灭菌日期, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
        Values
          (药品收发记录_Id.Nextval, 1, 21, v_其他出库no, n_出库序号, n_虚拟库房id, n_领药部门id, 类别id_In, -1, 收费细目id_In, n_批次, v_产地, v_批号,
           d_效期, Decode(v_Count, 1, 1, 付数_In), Decode(v_Count, 1, n_当前数量, n_当前数量 / 付数_In),
           Decode(v_Count, 1, n_当前数量, n_当前数量 / 付数_In), n_当前单价, Round(n_当前单价 * n_当前数量, v_Dec), v_Err_Msg, v_操作员姓名,
           登记时间_In, v_费用id, 频次_In, 单量_In, v_用法, v_煎法, n_扣率, d_灭菌效期, d_灭菌日期, n_供药单位id, d_生产日期, v_批准文号, v_商品条码, v_内部条码);
      End If;
      v_Err_Msg := '';
      n_总数量  := n_总数量 - n_当前数量;
      n_总金额  := n_总金额 + Round(n_当前数量 * n_当前单价, v_Dec);
    End Loop;
  
    --未发药品记录
    Update 未发药品记录
    Set 病人id = 病人id_In, 主页id = 主页id_In, 姓名 = 姓名_In
    Where 单据 = Decode(多病人单_In, 1, 10, 9) + Decode(收费类别_In, '4', 16, 0) And NO = No_In And
          Nvl(库房id, 0) = Nvl(执行部门id_In, 0);
    If Sql%RowCount = 0 Then
      --取身份优先级
      Begin
        Select b.优先级 Into v_优先级 From 病人信息 A, 身份 B Where a.身份 = b.名称(+) And a.病人id = 病人id_In;
      Exception
        When Others Then
          Null;
      End;
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 填制日期, 已收费, 打印状态)
      Values
        (Decode(多病人单_In, 1, 10, 9) + Decode(收费类别_In, '4', 16, 0), No_In, 病人id_In, 主页id_In, 姓名_In, v_优先级, n_领药部门id,
         执行部门id_In, 登记时间_In, Decode(划价_In, 1, 0, 1), 0);
    End If;
    Zl_Prescription_Type_Zy_Update(No_In, 2, 收费细目id_In, 收费类别_In);
    --可能时价药品的库存金额和数量变化了
    If n_时价 = 1 Then
      --只有一个批次时,直接取该批次的单价
      If n_当前数量 <> 付数_In * 数次_In Then
        n_当前单价 := Round(n_总金额 / (付数_In * 数次_In), n_单价小数);
      End If;
      If n_当前单价 <> 标准单价_In Then
        Close c_Stock;
        If 医嘱序号_In Is Null Then
          If 收费类别_In = '4' Then
            v_Err_Msg := '第 ' || 序号_In || ' 行的时价卫生材料"' || v_名称 || '"当前计算单价不一致,请重新输入数量计算！';
          Else
            v_Err_Msg := '第 ' || 序号_In || ' 行的时价药品"' || v_名称 || '"当前计算单价不一致,请重新输入数量计算！';
          End If;
        Else
          --医嘱摆药时是按病人分次计算并提交数据库,因此不同病人使用相同实价药品没有问题。
          --但同一病人同时使用两笔以上相同实价药品则会有问题。
          If 收费类别_In = '4' Then
            v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现时价卫生材料"' || v_名称 || '"当前计算的单价发生变化。' || Chr(13) || Chr(10) ||
                         '请检查该病人是否同时使用了两笔相同的"' || v_名称 || '"！';
          Else
            v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现时价药品"' || v_名称 || '"当前计算的单价发生变化。' || Chr(13) || Chr(10) ||
                         '请检查该病人是否同时使用了两笔相同的"' || v_名称 || '"！';
          End If;
        End If;
        Raise Err_Item;
      End If;
    End If;
    Close c_Stock;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_住院记帐记录_Insert;
/

--129869:蒋廷中,2019-01-11,新增病人用药清单涉及的过程
Create Or Replace Procedure Zl_病人用药清单_Insert
(
  Id_In         In 病人用药清单.Id%Type,
  病人id_In     In 病人用药清单.病人id%Type,
  主页id_In     In 病人用药清单.主页id%Type,
  组号_In       In 病人用药清单.组号%Type,
  用药来源_In   In 病人用药清单.用药来源%Type,
  药品类别_In   In 病人用药清单.药品类别%Type,
  用药内容_In   In 病人用药清单.用药内容%Type,
  诊疗项目id_In In 病人用药清单.诊疗项目id%Type,
  收费细目id_In In 病人用药清单.收费细目id%Type,
  天数_In       In 病人用药清单.天数%Type,
  开始时间_In   In 病人用药清单.开始时间%Type,
  终止时间_In   In 病人用药清单.终止时间%Type,
  登记时间_In   In 病人用药清单.登记时间%Type,
  登记人_In     In 病人用药清单.登记人%Type,
  总给予量_In   In 病人用药清单.总给予量%Type,
  单次用量_In   In 病人用药清单.单次用量%Type,
  执行频次_In   In 病人用药清单.执行频次%Type,
  频率次数_In   In 病人用药清单.频率次数%Type,
  频率间隔_In   In 病人用药清单.频率间隔%Type,
  间隔单位_In   In 病人用药清单.间隔单位%Type,
  用法id_In     In 病人用药清单.用法id%Type,
  煎法id_In     In 病人用药清单.煎法id%Type,
  备注_In       In 病人用药清单.备注%Type
) Is
Begin
  Insert Into 病人用药清单
    (ID, 病人id, 主页id, 组号, 用药来源, 药品类别, 用药内容, 诊疗项目id, 收费细目id, 天数, 开始时间, 终止时间, 登记时间, 登记人, 总给予量, 单次用量, 执行频次, 频率次数, 频率间隔,
     间隔单位, 用法id, 煎法id, 备注)
  Values
    (Id_In, 病人id_In, 主页id_In, 组号_In, 用药来源_In, 药品类别_In, 用药内容_In, 诊疗项目id_In, 收费细目id_In, 天数_In, 开始时间_In,
     终止时间_In, 登记时间_In, 登记人_In, 总给予量_In, 单次用量_In, 执行频次_In, 频率次数_In, 频率间隔_In, 间隔单位_In, 用法id_In, 煎法id_In, 备注_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人用药清单_Insert;
/

--129869:蒋廷中,2019-01-11,新增病人用药清单涉及的过程
CREATE OR REPLACE Procedure Zl_病人用药清单_Update
(
  id_In     In 病人用药清单.id%Type,
  组号_In       In 病人用药清单.组号%Type,
  用药来源_In   In 病人用药清单.用药来源%Type,
  药品类别_In   In 病人用药清单.药品类别%Type,
  用药内容_In   In 病人用药清单.用药内容%Type,
  诊疗项目id_In In 病人用药清单.诊疗项目id%Type,
  收费细目id_In In 病人用药清单.收费细目id%Type,
  天数_In       In 病人用药清单.天数%Type,
  开始时间_In   In 病人用药清单.开始时间%Type,
  终止时间_In   In 病人用药清单.终止时间%Type,
  登记时间_In   In 病人用药清单.登记时间%Type,
  登记人_In     In 病人用药清单.登记人%Type,
  总给予量_In   In 病人用药清单.总给予量%Type,
  单次用量_In   In 病人用药清单.单次用量%Type,
  执行频次_In   In 病人用药清单.执行频次%Type,
  频率次数_In   In 病人用药清单.频率次数%Type,
  频率间隔_In   In 病人用药清单.频率间隔%Type,
  间隔单位_In   In 病人用药清单.间隔单位%Type,
  用法id_In     In 病人用药清单.用法id%Type,
  煎法id_In     In 病人用药清单.煎法id%Type,
  备注_In       In 病人用药清单.备注%Type
) Is
Begin
  Update 病人用药清单
  Set 组号 = 组号_In, 用药来源 = 用药来源_In, 药品类别 = 药品类别_In, 用药内容 = 用药内容_In, 诊疗项目id = 诊疗项目id_In, 收费细目id = 收费细目id_In, 天数 = 天数_In,
      开始时间 = 开始时间_In, 终止时间 = 终止时间_In, 登记时间 = 登记时间_In, 登记人 = 登记人_In, 总给予量 = 总给予量_In, 单次用量 = 单次用量_In, 执行频次 = 执行频次_In,
      频率次数 = 频率次数_In, 频率间隔 = 频率间隔_In, 间隔单位 = 间隔单位_In, 用法id = 用法id_In, 煎法id = 煎法id_In, 备注 = 备注_In
  Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人用药清单_Update;
/

--129869:蒋廷中,2019-01-11,新增病人用药清单涉及的过程
Create Or Replace Procedure Zl_病人用药清单_Delete(Id_In In 病人用药清单.Id%Type) Is
Begin
  Delete From 病人用药清单 Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人用药清单_Delete;
/

--129869:蒋廷中,2019-01-11,新增病人用药清单涉及的过程
Create Or Replace Procedure Zl_病人用药配方_Insert
(
  配方id_In     In 病人用药配方.配方id%Type,
  序号_In       In 病人用药配方.序号%Type,
  诊疗项目id_In In 病人用药配方.诊疗项目id%Type,
  收费细目id_In In 病人用药配方.收费细目id%Type,
  单量_In       In 病人用药配方.单量%Type,
  脚注_In       In 病人用药配方.脚注%Type
) Is
Begin
  Insert Into 病人用药配方
    (配方id, 序号, 诊疗项目id, 收费细目id, 单量, 脚注)
  Values
    (配方id_In, 序号_In, 诊疗项目id_In, 收费细目id_In, 单量_In, 脚注_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人用药配方_Insert;
/

--129869:蒋廷中,2019-01-11,新增病人用药清单涉及的过程
CREATE OR REPLACE Procedure Zl_病人用药配方_Delete(配方Id_In In 病人用药配方.配方Id%Type) Is
Begin
  Delete From 病人用药配方 Where 配方Id = 配方Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人用药配方_Delete;
/

--128654:陈刘,2019-01-10,分组汇总汇总数据误差修正
CREATE OR REPLACE Procedure Zl_护理二次汇总_Update
(
  文件id_In   In 病人护理数据.文件id%Type,
  发生时间_In In 病人护理数据.发生时间%Type,
  汇总时间_In In 病人护理数据.发生时间%Type, --汇总记录的发生时间
  记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，签名记录=5，审签记录=15
  项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
  记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容；37或38/37
  体温部位_In In 病人护理明细.体温部位%Type := Null,
  操作员_In   In 病人护理数据.保存人%Type := Null,
  记录组号_In In 病人护理明细.记录组号%Type := Null, --适用分类汇总(一条数据对应多条相同项目的明细)
  数据来源_In In 病人护理明细.来源id%Type := 0,
  删除_In     In Number := 0
) Is
  Intins     Number(18);
  Int共用    Number(1);
  n_Exists   Number(1);
  n_Newid    病人护理数据.Id%Type;
  n_Oldid    病人护理数据.Id%Type;
  n_Synchro  Number(1);
  n_汇总id   病人护理数据.Id%Type;
  n_文件id   病人护理数据.文件id%Type;
  v_汇总文本 病人护理数据.汇总文本%Type;

  n_汇总类别 病人护理数据.汇总类别%Type;
  v_科室id   部门表.Id%Type;
  v_保存人   人员表.姓名%Type;
  n_记录id   病人护理数据.Id%Type;
  n_明细id   病人护理明细.Id%Type;
  v_数据来源 病人护理明细.数据来源%Type;
  n_项目性质 护理记录项目.项目性质%Type;
  --提取该病人当前科室所有未结束的护理文件，且文件开始时间小于等于记录发生时间的文件列表供同步数据使用
  Cursor Cur_Fileformats Is
    Select a.Id As 格式id, b.Id As 文件id, a.保留, a.子类, b.婴儿
    From 病历文件列表 A, 病人护理文件 B, 病人护理文件 C, 病人护理数据 D
    Where a.种类 = 3 And a.保留 <> 1 And a.Id = b.格式id And b.Id <> c.Id And b.结束时间 Is Null And b.开始时间 <= d.发生时间 And
          (a.通用 = 1 Or (a.通用 = 2 And b.科室id = c.科室id)) And c.病人id = b.病人id And c.主页id = b.主页id And c.婴儿 = b.婴儿 And
          c.Id = d.文件id And d.Id = n_记录id And c.Id = 文件id_In
    Order By a.编号;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --取记录ID
  Int共用  := 0;
  n_记录id := 0;
  Intins   := 0;
  If 操作员_In Is Null Then
    v_保存人 := Zl_Username;
  Else
    v_保存人 := 操作员_In;
  End If;

  Begin
    Select Max(ID) Into n_汇总id From 病人护理数据 Where 文件id = 文件id_In And 发生时间 = 汇总时间_In;
  End;
  Begin
    Select ID, 汇总类别
    Into n_记录id, n_汇总类别
    From 病人护理数据
    Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
  Exception
    When Others Then
      n_记录id := 0;
  End;

  Intins := 0;
  --查找除了本条要删除的数据，是否还存其他有效的数据，如果存在只删除本条数据，否则删除此发生时间对应的所有数据。
  If 删除_In = 1 Then

    Select a.Id
    Into n_文件id
    From 病人护理文件 A, 病人护理文件 B, 病历文件列表 C
    Where b.Id = 文件id_In And a.病人id = b.病人id And a.主页id = b.主页id And a.婴儿 = b.婴儿 And a.格式id = c.Id And c.种类 = 3 And
          c.保留 = -1 And a.开始时间 < 发生时间_In And (a.结束时间 > 发生时间_In Or a.结束时间 Is Null);
    Select Max(1)
    Into Intins
    From 病人护理数据
    Where Instr(汇总文本 || ',', ',' || n_汇总id || ',') > 0 And 文件id = n_文件id;
    Begin
      Select a.Id, a.汇总类别
      Into n_记录id, n_汇总类别
      From 病人护理数据 A, 病人护理明细 B
      Where a.文件id = n_文件id And a.发生时间 = 发生时间_In And a.Id = b.记录id And b.来源id = 数据来源_In And b.项目序号 = 项目序号_In;
    Exception
      When Others Then
        n_记录id := 0;
    End;

    Begin
      Select ID
      Into n_明细id
      From 病人护理明细
      Where 记录id = n_记录id And Nvl(项目序号, 0) = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
            Nvl(记录组号, 1) = Nvl(记录组号_In, 1) And 终止版本 Is Null;
    Exception
      --无数据退出
      When Others Then
        Return;
    End;

    If Intins = 0 Then
      Select Count(ID)
      Into Intins
      From 病人护理明细
      Where 记录id = n_记录id And Mod(记录类型, 10) <> 5 And 终止版本 Is Null And ID <> n_明细id;
      If Intins = 0 Then
        Delete From 病人护理明细 Where 记录id = n_记录id;
      Else
        Delete From 病人护理明细 Where ID = n_明细id;
      End If;
      Delete From 病人护理数据 A
      Where a.Id = n_记录id And Not Exists (Select 1 From 病人护理明细 B Where b.记录id = a.Id);
    Else
      Select 汇总文本 Into v_汇总文本 From 病人护理数据 Where ID = n_记录id;
      v_汇总文本 := Replace(v_汇总文本 || ',', ',' || n_汇总id || ',', ',');
      v_汇总文本 := Substr(v_汇总文本, 1, Length(v_汇总文本) - 1);
      Update 病人护理数据 Set 汇总文本 = v_汇总文本 Where ID = n_记录id;
      If v_汇总文本 Is Null Then
        Select Count(ID)
        Into Intins
        From 病人护理明细
        Where 记录id = n_记录id And Mod(记录类型, 10) <> 5 And 终止版本 Is Null And ID <> n_明细id;
        If Intins = 0 Then
          Delete From 病人护理明细 Where 记录id = n_记录id;
        Else
          Delete From 病人护理明细 Where ID = n_明细id;
        End If;
        Delete From 病人护理数据 A
        Where a.Id = n_记录id And Not Exists (Select 1 From 病人护理明细 B Where b.记录id = a.Id);
      End If;
    End If;
    Return; --删除就直接退出
  End If;

  --############
  --二次汇总保存数据
  --############
  --处理体温单（一个病人始终只存在一份有效的体温单文件）
  --如果体温表存在相同发生时间的数据，使用它的ID
  For Row_Format In Cur_Fileformats Loop
    If Row_Format.保留 = -1 Then
      If Row_Format.子类 = '1' Then
        Begin
          Select 1, h.项目性质
          Into Intins, n_项目性质
          From (Select To_Char(f.项目序号) As 序号, g.项目性质
                 From 体温记录项目 F, 护理记录项目 G
                 Where f.项目序号 = g.项目序号 And g.项目性质 = 2 And
                       (g.适用科室 = 1 Or (g.适用科室 = 2 And Exists
                        (Select 1 From 护理适用科室 D Where g.项目序号 = d.项目序号 And d.科室id = v_科室id))) And Nvl(g.应用方式, 0) <> 0 And
                       (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(Nvl(Row_Format.婴儿, 0), 0, 1, 2))
                 Union All
                 Select b.内容文本 As 序号, 1 As 项目性质
                 From 病历文件结构 A, 病历文件结构 B
                 Where a.文件id = Row_Format.格式id And a.父id Is Null And a.对象序号 In (2, 3) And b.父id = a.Id) H
          Where Instr(',' || h.序号 || ',', ',' || 项目序号_In || ',', 1) > 0;
        Exception
          When Others Then
            Intins := 0;
        End;
      Else
        Begin
          Select 1, g.项目性质
          Into Intins, n_项目性质
          From 体温记录项目 F, 护理记录项目 G
          Where f.项目序号 = g.项目序号 And Nvl(g.应用方式, 0) <> 0 And g.护理等级 >= 0 And
                (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(Nvl(Row_Format.婴儿, 0), 0, 1, 2)) And f.项目序号 = 项目序号_In And
                (g.适用科室 = 1 Or (g.适用科室 = 2 And Exists
                 (Select 1 From 护理适用科室 D Where g.项目序号 = d.项目序号 And d.科室id = v_科室id)));
        Exception
          When Others Then
            Intins := 0;
        End;
      End If;

      If Intins > 0 Then

        --LPF,2013-01-23,检查此项目是否需要进行同步(对于以前已经同步过的数据，为了保证记录单和体温单数据一直将不根据此函数判断。)
        n_Synchro := Zl_Temperatureprogram(文件id_In, v_科室id, 项目序号_In, 发生时间_In);
        Begin
          Select b.Id Into n_Newid From 病人护理数据 B Where b.文件id = Row_Format.文件id And b.发生时间 = 发生时间_In;
        Exception
          When Others Then
            n_Newid := 0;
        End;
        n_Oldid := n_Newid;
        If n_Newid = 0 And n_Synchro = 1 Then
          Select 病人护理数据_Id.Nextval Into n_Newid From Dual;
          --产生体温单主记录
          Insert Into 病人护理数据
            (ID, 文件id, 保存人, 保存时间, 发生时间, 最后版本, 汇总文本)
          Values
            (n_Newid, Row_Format.文件id, v_保存人, Sysdate, 发生时间_In, 1, ',' || n_汇总id);
        Else
          If n_Oldid <> 0 Then
            Select Max(1)
            Into n_Exists
            From 病人护理数据
            Where ID = n_Oldid And Instr(汇总文本 || ',', ',' || n_汇总id || ',') > 0;
            If n_Exists Is Null Then
              Update 病人护理数据 Set 汇总文本 = 汇总文本 || ',' || n_汇总id Where ID = n_Oldid;
            End If;
          End If;
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
                (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 体温部位, 数据来源, 来源id, 开始版本, 终止版本, 记录人,
                 记录时间, 记录组号)
                Select 病人护理明细_Id.Nextval, n_Newid, 记录类型_In, a.分组名, a.项目id, a.项目序号, a.项目名称, a.项目类型, 记录内容_In, a.项目单位, 0,
                       体温部位_In, 1, 数据来源_In, 1, Null, b.记录人, Sysdate, 1
                From 护理记录项目 A, 病人护理明细 B
                Where a.项目序号 = b.项目序号(+) And b.终止版本(+) Is Null And b.记录id(+) = n_记录id And a.项目序号 = 项目序号_In And
                      Rownum < 2;
              If Sql%RowCount > 0 Then
                Int共用 := 1;
              End If;
            End If;
          Else
            For r_Twd In (Select b.Id
                          From 病人护理数据 B
                          Where b.文件id = Row_Format.文件id And
                                b.发生时间 Between
                                To_Date(To_Char(发生时间_In, 'YYYY-MM-DD hh24:mi') || ':00', 'YYYY-MM-DD hh24:mi:ss') And
                                To_Date(To_Char(发生时间_In, 'YYYY-MM-DD hh24:mi') || ':59', 'YYYY-MM-DD hh24:mi:ss')) Loop
              Select Max(1)
              Into n_Exists
              From 病人护理明细
              Where 记录id = r_Twd.Id And 项目序号 = 项目序号_In And 来源id = 数据来源_In;
              If n_Exists = 1 Then
                n_Oldid := r_Twd.Id;
              End If;
              Exit When n_Exists Is Not Null;
            End Loop;
            If n_Exists Is Null Then
              Select 病人护理数据_Id.Nextval Into n_Newid From Dual;
              --产生体温单主记录
              Insert Into 病人护理数据
                (ID, 文件id, 保存人, 保存时间, 发生时间, 最后版本, 汇总文本)
              Values
                (n_Newid, Row_Format.文件id, v_保存人, Sysdate, 发生时间_In + 1 / 24 / 60 / 60, 1, ',' || n_汇总id);

              Insert Into 病人护理明细
                (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 体温部位, 数据来源, 来源id, 开始版本, 终止版本, 记录人,
                 记录时间, 记录组号)
                Select 病人护理明细_Id.Nextval, n_Newid, 记录类型_In, a.分组名, a.项目id, a.项目序号, a.项目名称, a.项目类型, 记录内容_In, a.项目单位, 0,
                       体温部位_In, 1, 数据来源_In, 1, Null, b.记录人, Sysdate, 1
                From 护理记录项目 A, 病人护理明细 B
                Where a.项目序号 = b.项目序号(+) And b.终止版本(+) Is Null And b.记录id(+) = n_记录id And a.项目序号 = 项目序号_In And
                      Rownum < 2;
              If Sql%RowCount > 0 Then
                Int共用 := 1;
              End If;
            Else
              Select Max(1)
              Into n_Exists
              From 病人护理数据
              Where ID = n_Oldid And Instr(汇总文本 || ',', ',' || n_汇总id || ',') > 0;
              If n_Exists Is Null Then
                Update 病人护理数据 Set 汇总文本 = 汇总文本 || ',' || n_汇总id Where ID = n_Oldid;
              End If;
            End If;
          End If;
        End If;
      End If;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_护理二次汇总_Update;
/
--128654:陈刘,2019-01-10,分组汇总汇总数据误差修正
CREATE OR REPLACE Procedure Zl_病人护理数据_Update
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
  n_Syntend   Number(1);
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
    From 病历文件列表 A, 病人护理文件 B, 病人护理文件 C, 病人护理数据 D
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
  n_Syntend   := 0;
  If 操作员_In Is Null Then
    v_保存人 := Zl_Username;
  Else
    v_保存人 := 操作员_In;
  End If;

  --如果是对应多份护理文件值为1，表示需同步其它护理文件；否则不处理文件同步
  n_Mutilbill := Zl_To_Number(zl_GetSysParameter('对应多份护理文件', 1255));
  --如果允许多份护理文件之间数据同步,则自动同步,否则不同步
  n_Syntend := Zl_To_Number(zl_GetSysParameter('允许数据同步', 1255));

  Begin
    Select ID, 汇总类别
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
  Where ID = 文件id_In;
  d_婴儿出院时间 := Null;
  If n_婴儿 <> 0 Then
    Begin
      Select 开始执行时间
      Into d_婴儿出院时间
      From 病人医嘱记录 B, 诊疗项目目录 C
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
      From 病人变动记录 A, 病人护理文件 B
      Where a.科室id Is Not Null And a.病人id = b.病人id And a.主页id = b.主页id And b.Id = 文件id_In And
            (To_Date(To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI') || '59', 'YYYY-MM-DD HH24:MI:SS') >= a.开始时间 And
            (To_Date(To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI') || '00', 'YYYY-MM-DD HH24:MI:SS') < = Nvl(a.终止时间, Sysdate) Or
            a.终止时间 Is Null)) And Rownum < 2;
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
  From 病人护理明细 A, 病人护理数据 B
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
      Select ID
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
    Select Count(ID)
    Into Intins
    From 病人护理明细
    Where 记录id = n_记录id And Mod(记录类型, 10) <> 5 And 终止版本 Is Null And ID <> n_明细id;
    If Intins = 0 Then
      Delete From 病人护理明细 Where 记录id = n_记录id;
    Else
      Delete From 病人护理明细 Where ID = n_明细id;
    End If;

    Delete From 病人护理数据 A
    Where a.Id = n_记录id And Not Exists (Select 1 From 病人护理明细 B Where b.记录id = a.Id);

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
          From 病历文件列表 A, 病人护理文件 B, 病人护理数据 C
          Where a.Id = b.格式id And b.Id = c.文件id And c.Id = Rsdel.记录id;
        Exception
          When Others Then
            n_文件id := 0;
        End;
        Delete 病人护理数据 Where ID = Rsdel.记录id;
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
             From 病历文件结构 A, 护理记录项目 B
             Where a.要素名称 = b.项目名称 And b.项目序号 = 项目序号_In And
                   父id = (Select b.Id
                          From 病人护理文件 A, 病历文件结构 B
                          Where a.Id = 文件id_In And a.格式id = b.文件id And b.父id Is Null And b.对象序号 = 4)
             Union
             Select 项目序号
             From 护理记录项目
             Where 项目性质 = 2 And 项目序号 = 项目序号_In
             Union
             Select 项目序号
             From 护理记录项目
             Where 项目表示 = 4 And 项目序号 = 项目序号_In);
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
        (ID, 文件id, 发生时间, 最后版本, 保存人, 保存时间)
      Values
        (n_记录id, 文件id_In, 发生时间_In, n_最高版本, v_保存人, Sysdate);
    End If;

    --插入本次登记的病人护理明细
    Update 病人护理明细
    Set 记录内容 = 记录内容_In, 未记说明 = 未记说明_In, 记录人 = v_保存人, 记录时间 = Sysdate
    Where 记录id = n_记录id And 项目序号 = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
          Nvl(记录组号, 0) = Nvl(记录组号_In, 0) And 开始版本 = n_最高版本 And 终止版本 Is Null;
    If Sql%RowCount = 0 Then
      Select 病人护理明细_Id.Nextval Into n_明细id From Dual;
      Insert Into 病人护理明细
        (ID, 记录id, 记录类型, 项目分组, 项目id, 相关序号, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 记录组号, 体温部位, 数据来源, 共用, 未记说明, 开始版本, 终止版本,
         记录人, 记录时间)
        Select n_明细id, n_记录id, 记录类型_In, a.分组名, a.项目id, 相关序号_In, a.项目序号, Upper(a.项目名称), a.项目类型, 记录内容_In, a.项目单位, 0,
               记录组号_In, 体温部位_In, 数据来源_In, Nvl(b.共用, 0), 未记说明_In, n_最高版本, Null, v_保存人, Sysdate
        From 护理记录项目 A, 病人护理明细 B
        Where a.项目序号 = b.项目序号(+) And b.终止版本(+) Is Null And b.记录id(+) = n_记录id And a.项目序号 = 项目序号_In And Rownum < 2;
    End If;
    Select ID
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
      Update 病人护理数据 Set 保存人 = v_保存人, 保存时间 = Sysdate Where ID = n_记录id;
    End If;

    If Nvl(n_汇总类别, 0) <> 0   Then
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
                   From 体温记录项目 F, 护理记录项目 G
                   Where f.项目序号 = g.项目序号 And g.项目性质 = 2 And
                         (g.适用科室 = 1 Or
                         (g.适用科室 = 2 And Exists
                          (Select 1 From 护理适用科室 D Where g.项目序号 = d.项目序号 And d.科室id = v_科室id))) And Nvl(g.应用方式, 0) <> 0 And
                         (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(Nvl(Row_Format.婴儿, 0), 0, 1, 2))
                   Union All
                   Select b.内容文本 As 序号, 1 As 项目性质
                   From 病历文件结构 A, 病历文件结构 B
                   Where a.文件id = Row_Format.格式id And a.父id Is Null And a.对象序号 In (2, 3) And b.父id = a.Id) H
            Where Instr(',' || h.序号 || ',', ',' || 项目序号_In || ',', 1) > 0;
          Exception
            When Others Then
              Intins := 0;
          End;
        Else
          Begin
            Select 1, g.项目性质
            Into Intins, n_项目性质
            From 体温记录项目 F, 护理记录项目 G
            Where f.项目序号 = g.项目序号 And Nvl(g.应用方式, 0) <> 0 And g.护理等级 >= 0 And
                  (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(Nvl(Row_Format.婴儿, 0), 0, 1, 2)) And f.项目序号 = 项目序号_In And
                  (g.适用科室 = 1 Or (g.适用科室 = 2 And Exists
                   (Select 1 From 护理适用科室 D Where g.项目序号 = d.项目序号 And d.科室id = v_科室id)));
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
            From 病人护理文件 A, 病人护理数据 B
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
              (ID, 文件id, 保存人, 保存时间, 发生时间, 最后版本)
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
                  (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 体温部位, 数据来源, 来源id, 开始版本, 终止版本, 记录人,
                   记录时间, 记录组号)
                  Select 病人护理明细_Id.Nextval, n_Newid, b.记录类型, b.项目分组, b.项目id, b.项目序号, b.项目名称, b.项目类型, b.记录内容, b.项目单位,
                         b.记录标记, b.体温部位, 1, b.Id, 1, Null, b.记录人, Sysdate, 1
                  From (Select 项目序号_In As 项目序号, Nvl(体温部位_In, '无') As 体温部位
                         From Dual
                         Minus
                         Select f.项目序号, Decode(Nvl(f.项目性质, 1), 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无'))
                         From 病人护理明细 E, 护理记录项目 F
                         Where e.记录id = n_Newid And e.项目序号 = f.项目序号) A, 病人护理明细 B
                  Where a.项目序号 = b.项目序号 And b.记录id = n_记录id And b.Id = n_明细id;
                If Sql%RowCount > 0 Then
                  Int共用 := 1;
                End If;
              End If;
            Else
              Update 病人护理明细
              Set 记录内容 = 记录内容_In, 来源id = n_明细id
              Where 记录id = n_Newid And 项目序号 = 项目序号_In And
                    Decode(n_项目性质, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And 数据来源 > 0;
              If Sql%RowCount > 0 Then
                Int共用 := 1;
              End If;
            End If;
          End If;
        End If;
        --2\再循环处理记录单
      Else
        If n_Mutilbill = 1 And n_Syntend = 1 Then
          --提取记录单与当前记录单存在重叠的且有数据的固定项目
          Select Count(*)
          Into Intins
          From (Select b.项目序号
                 From 病历文件结构 A, 护理记录项目 B
                 Where a.要素名称 = b.项目名称 And b.项目表示 In (0, 4, 5) And
                       父id =
                       (Select ID From 病历文件结构 Where 文件id = Row_Format.格式id And 父id Is Null And 对象序号 = 4)
                 Intersect
                 Select b.项目序号
                 From 病历文件结构 A, 护理记录项目 B, 病人护理文件 C, 病人护理数据 D, 病人护理明细 G
                 Where c.Id = d.文件id And a.文件id = c.格式id And d.Id = g.记录id And d.Id = n_记录id And g.Id = n_明细id And
                       b.项目序号 = g.项目序号 And b.项目表示 In (0, 4, 5) And g.记录类型 = 1 And a.要素名称 = b.项目名称 And
                       a.父id = (Select ID From 病历文件结构 E Where e.文件id = c.格式id And 父id Is Null And 对象序号 = 4));

          If Intins > 0 Then
            n_Newid := 0;
            --可能指定文件已经存在相同发生时间的数据，直接用它的ID即可
            Begin
              Select c.Id
              Into n_Newid
              From 病人护理数据 C
              Where c.文件id = Row_Format.文件id And c.发生时间 = 发生时间_In;
            Exception
              When Others Then
                n_Newid := 0;
            End;

            If n_Newid = 0 Then
              --产生记录单主记录
              Select 病人护理数据_Id.Nextval Into n_Newid From Dual;

              Insert Into 病人护理数据
                (ID, 文件id, 保存人, 保存时间, 发生时间, 最后版本)
                Select n_Newid, Row_Format.文件id, c.保存人, c.保存时间, c.发生时间, 1
                From 病人护理数据 C
                Where c.Id = n_记录id;
            End If;

            If n_Newid > 0 Then
              --插入未同步的记录单数据
              Select Count(*) Into v_数据来源 From 病人护理明细 Where 记录id = n_Newid And 项目序号 = 项目序号_In;
              If v_数据来源 = 0 Then
                Insert Into 病人护理明细
                  (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 体温部位, 数据来源, 来源id, 未记说明, 开始版本, 终止版本,
                   记录人, 记录时间)
                  Select 病人护理明细_Id.Nextval, n_Newid, b.记录类型, b.项目分组, b.项目id, b.项目序号, b.项目名称, b.项目类型, b.记录内容, b.项目单位,
                         b.记录标记, b.体温部位, 1, b.Id, b.未记说明, 1, Null, b.记录人, Sysdate
                  From (Select b.项目序号
                         From 病历文件结构 A, 护理记录项目 B
                         Where a.要素名称 = b.项目名称 And b.项目表示 In (0, 4, 5) And
                               父id = (Select ID
                                      From 病历文件结构
                                      Where 文件id = Row_Format.格式id And 父id Is Null And 对象序号 = 4)
                         Intersect
                         Select b.项目序号
                         From 病历文件结构 A, 护理记录项目 B, 病人护理文件 C, 病人护理数据 D, 病人护理明细 G
                         Where c.Id = d.文件id And a.文件id = c.格式id And d.Id = g.记录id And d.Id = n_记录id And g.Id = n_明细id And
                               b.项目序号 = g.项目序号 And b.项目表示 In (0, 4, 5) And g.记录类型 = 1 And a.要素名称 = b.项目名称 And
                               a.父id =
                               (Select ID From 病历文件结构 E Where e.文件id = c.格式id And 父id Is Null And 对象序号 = 4)) A, 病人护理明细 B
                  Where a.项目序号 = b.项目序号 And b.记录id = n_记录id And b.Id = n_明细id;
                If Sql%RowCount > 0 Then
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
                Where 记录id = n_Newid And 项目序号 = 项目序号_In And (数据来源 > 0 Or 数据来源 <> 3);
                If Sql%RowCount > 0 Then
                  Int共用 := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End Loop;

    If Int共用 = 1 Then
      Update 病人护理明细 Set 共用 = 1 Where ID = n_明细id;
      --将历史数据的共用标志设置为NULL
      Update 病人护理明细 Set 共用 = Null Where 记录id = n_记录id And 项目序号 = 项目序号_In And ID <> n_明细id;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人护理数据_Update;
/
--128654:陈刘,2019-01-10,分组汇总汇总数据误差修正

CREATE OR REPLACE Procedure Zl_病人护理数据_Collect
(
  文件id_In   In 病人护理数据.文件id%Type,
  发生时间_In In 病人护理数据.发生时间%Type,
  汇总类别_In In 病人护理数据.汇总类别%Type,
  汇总文本_In In 病人护理数据.汇总文本%Type,
  汇总标记_In In 病人护理数据.汇总标记%Type,
  开始时点_In In 病人护理数据.开始时点%Type,
  结束时点_In In 病人护理数据.结束时点%Type,
  删除_In     Number := 0
) Is
  n_Exist  Number(1);
  v_记录id 病人护理数据.Id%Type;
  v_来源id 病人护理数据.Id%Type;
  v_User   人员表.姓名%Type;
  n_文件id 病人护理数据.文件id%Type;
Begin
  If 删除_In = 0 Then
    v_User := Zl_Username;
    Begin
      Select 1 Into n_Exist From 病人护理数据 Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
    Exception
      When Others Then
        n_Exist := 0;
    End;

    If n_Exist = 0 Then
      Insert Into 病人护理数据
        (ID, 文件id, 发生时间, 最后版本, 保存人, 保存时间, 汇总类别, 汇总文本, 汇总标记, 开始时点, 结束时点)
      Values
        (病人护理数据_Id.Nextval, 文件id_In, 发生时间_In, 1, v_User, Sysdate, 汇总类别_In, 汇总文本_In, 汇总标记_In, 开始时点_In, 结束时点_In);
    End If;
  Else
    Select ID Into v_记录id From 病人护理数据 Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
    Select a.Id
    Into n_文件id
    From 病人护理文件 A, 病人护理文件 B, 病历文件列表 C
    Where b.Id = 文件id_In And a.病人id = b.病人id And a.主页id = b.主页id And a.婴儿 = b.婴儿 And a.格式id = c.Id And c.种类 = 3 And
          c.保留 = -1 And a.开始时间 < 发生时间_In And (a.结束时间 > 发生时间_In Or a.结束时间 Is Null);
    Select Max(a.记录id)
    Into v_来源id
    From 病人护理明细 A, 病人护理明细 B
    Where a.来源id = b.Id(+) And b.记录id = v_记录id;

    For r_List In (Select a.发生时间, b.项目序号, b.项目名称,B.来源id
                   From 病人护理数据 A, 病人护理明细 B
                   Where 文件id = n_文件id And a.Id = b.记录id And Instr(汇总文本, v_记录id) > 0 And b.数据来源 = 1 ) Loop
      Zl_护理二次汇总_Update(文件id_In, r_List.发生时间, 发生时间_In, 1, r_List.项目序号, Null, Null, Null, Null, r_List.来源id, 1);
    End Loop;
    Delete 病人护理明细 Where 记录id = v_记录id;
    Delete 病人护理数据 Where 文件id = 文件id_In And 发生时间 = 发生时间_In;

  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人护理数据_Collect;
/

--134969:李南春,2019-01-14,预交支付检查
--134441:李南春,2019-01-09,挂号检查项目是否一致
CREATE OR REPLACE Procedure Zl_病人挂号记录_出诊_Insert
(
  出诊记录id_In    临床出诊记录.Id%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  序号_In          门诊费用记录.序号%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  数次_In          门诊费用记录.数次%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  结算方式_In      Varchar2,
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  医生姓名_In      挂号安排.医生姓名%Type,
  医生id_In        挂号安排.医生id%Type,
  病历费_In        Number, --该条记录是否病历工本费
  急诊_In          Number,
  号别_In          挂号安排.号码%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  领用id_In        票据使用明细.领用id%Type,
  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In    门诊费用记录.保险大类id%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In      Number := 0, --挂号是否使用收费票据
  保险编码_In      门诊费用记录.保险编码%Type,
  复诊_In          病人挂号记录.复诊%Type := 0,
  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In          病人挂号记录.社区%Type := Null,
  预约接收_In      Number := 0,
  预约方式_In      预约方式.名称%Type := Null,
  生成队列_In      Number := 0,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  合作单位_In      病人预交记录.合作单位%Type := Null,
  操作类型_In      Number := 0,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  退号重用_In      Number := 1,
  冲预交病人ids_In Varchar2 := Null,
  修正病人费别_In  Number := 0,
  更新交款余额_In  Number := 0, --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
  预约顺序号_In    临床出诊序号控制.预约顺序号%Type := Null,
  修正病人年龄_In  Number := 0,
  收费单_In        病人挂号记录.收费单%Type := Null
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
    Select NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id, Min(Decode(记录性质, 1, 收款时间, NULL)) as 收款时间
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id = v_病人id And Nvl(预交类别, 2) = 1 Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO
    Order By 收款时间;
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
  n_原始分时段   Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况
  n_已约数       病人挂号汇总.已约数%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_费用id   门诊费用记录.Id%Type;
  n_病人余额 病人预交记录.金额%Type;
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
  v_结算方式记录   Varchar2(1000);
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  v_结算方式       结算方式.名称%Type;
  v_结算内容       Varchar2(1000);
  v_当前结算       Varchar2(200);
  v_结算号码       病人预交记录.结算号码%Type;
  n_结算金额       病人预交记录.冲预交%Type;
  n_三方卡标志     Number(2);
  n_预约顺序号     临床出诊序号控制.预约顺序号%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;
  n_Exists         Number;
  n_挂出的最大序号 Number(4) := 0;
  n_分时点显示     Number(3);
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
  n_状态           临床出诊序号控制.挂号状态%Type;
Begin

  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id := Zl_Get组id(操作员姓名_In);

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
  v_费别 := 费别_In;
  d_时段时间 := 发生时间_In;

  If Nvl(序号_In, 0) = 1 Then
    If 出诊记录id_In Is Not Null Then
      Begin
        Select 1
        Into n_Exists
        From 临床出诊记录 a, 临床出诊号源 b
        Where a.Id = 出诊记录id_In And a.号源id = b.Id And b.号码 = 号别_In And a.科室id = 执行部门id_In And a.项目id = 收费细目id_In And
              Nvl(a.是否发布, 0) = 1 And Nvl(a.是否锁定, 0) = 0;
      Exception
        When Others Then
          v_Err_Msg := '无法确定出诊记录，请检查出诊记录是否存在或被锁定！';
          Raise Err_Item;
      End;
    End If;

    If 费别_In Is Null Then
      Begin
        Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Exception
        When Others Then
          v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
          Raise Err_Item;
      End;
    End If;
    If Nvl(修正病人费别_In, 0) = 1 And v_费别 Is Not Null Then
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    End If;

    If Nvl(修正病人年龄_In, 0) = 1 Then
      Update 病人信息 Set 年龄 = 年龄_In Where 病人id = 病人id_In;
    End If;

    If 门诊号_In Is Not Null Then
      Update 病人信息 Set 门诊号 = 门诊号_In Where 病人id = 病人id_In And Nvl(门诊号, 0) = 0;
    End If;

    Update 临床出诊序号控制
    Set 挂号状态 = 0
    Where 记录id = 出诊记录id_In And 序号 = 号序_In And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;

    --获取是否分时段
    Begin
      Select Nvl(是否分时段, 0), Nvl(是否序号控制, 0), 限号数, 限约数, 已挂数, 已约数
      Into n_分时段, n_序号控制, n_限号数, n_限约数, n_已挂数, n_已约数
      From 临床出诊记录
      Where ID = 出诊记录id_In;
      n_原始分时段 := n_分时段;
    Exception
      When Others Then
        n_分时段     := 0;
        n_原始分时段 := n_分时段;
        n_序号控制   := 0;
        n_限号数     := Null;
        n_限约数     := Null;
    End;

    --获取当前未使用的序号
    If Nvl(预约挂号_In, 0) = 0 Then
      n_预约有效时间 := Zl_To_Number(Zl_Getsysparameter('预约有效时间', 1111));
      n_失约挂号     := Zl_To_Number(Zl_Getsysparameter('失约用于挂号', 1111));
    End If;
    n_失效数 := 0;

    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>
      If Nvl(预约挂号_In, 0) = 0 And Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
        Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 预约时间), 1, 1, 0))
        Into n_失效数
        From 病人挂号记录
        Where 出诊记录id = 出诊记录id_In And 记录状态 = 1 And 记录性质 = 2;
      End If;
      If n_序号 Is Null Then
        n_已用序号 := Null;
        If n_原始分时段 = 0 Then
          Select Min(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And Nvl(挂号状态, 0) = 0;
        End If;
        If n_已用序号 Is Null Then
          Select Nvl(Max(序号), 0) + 1 Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
        End If;
        n_序号 := Nvl(n_已用序号, 0);
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已挂数 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限号数！';
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
      --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
      If n_序号 Is Null And Nvl(预约挂号_In, 0) = 1 Then
        Begin
          Select 序号
          Into n_序号
          From 临床出诊序号控制
          Where 记录id = 出诊记录id_In And 开始时间 = 发生时间_In And Rownum < 2;
        Exception
          When Others Then
            n_序号 := Null;
        End;
      End If;

      If Nvl(预约挂号_In, 0) = 1 Then
        Begin
          Select Nvl(序号, 0),
                 To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                 数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
          Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
          From 临床出诊序号控制
          Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 Is Null;
        Exception
          When Others Then
            n_时段序号 := -1;
            n_分时段   := 0;
            d_时段时间 := 发生时间_In;
            n_时段限号 := 0;
            n_时段限约 := 0;
        End;
      End If;

      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 And Nvl(预约挂号_In, 0) = 1 Then
        --<正常预约挂号-->

        Select Nvl(Sum(Decode(Nvl(Sign(a.开始时间 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
        Into n_已约数
        From 临床出诊序号控制 A
        Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4, 5);

        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
      Elsif 操作类型_In = 0 And n_限号数 <= n_已挂数 - n_失效数 Then
        v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
        Raise Err_Item;
      End If;
      If Nvl(预约挂号_In, 0) = 1 Then
        --获取当天挂出的最大号序
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 临床出诊序号控制 A
        Where 记录id = 出诊记录id_In And 预约顺序号 Is Null And 挂号状态 Not In (0, 5);
        If 预约顺序号_In Is Not Null Then
          n_预约顺序号 := 预约顺序号_In;
        Else

          Select Nvl(Max(预约顺序号), 0) + 1
          Into n_预约顺序号
          From 临床出诊序号控制
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Not Null;

        End If;
        --设置序号
        n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_预约顺序号;
        If n_预约顺序号 Is Null Then
          n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
        End If;
      End If;
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := 发生时间_In;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;

      Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(开始时间 - d_时段时间), 0, 1, 0))
      Into n_已用序号, n_已挂数, n_已用数量
      From 临床出诊序号控制
      Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4, 5);

      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then

          Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 开始时间), 1, 1, 0))
          Into n_失效数
          From 临床出诊序号控制
          Where 记录id = 出诊记录id_In And 开始时间 Between Trunc(Sysdate) And Sysdate And Nvl(挂号状态, 0) = 2;

        End If;
      End If;

      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then

          --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
          If 号序_In Is Null Then

            Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'),
                                'yyyy-mm-dd hh24:mi:ss'))
            Into d_最大序号时间
            From 临床出诊序号控制
            Where 记录id = 出诊记录id_In And Nvl(数量, 0) <> 0;

            n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                       When -1 Then
                        0
                       Else
                        1
                     End;

          End If;

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
          --预约挂号
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
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已挂数, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已挂数, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;

    --更新挂号序号状态
    If Not n_序号 Is Null Then
      If n_分时段 = 1 Then
        d_序号时间 := 发生时间_In;
      Else
        d_序号时间 := Trunc(发生时间_In);
      End If;

      --锁定序号的处理
      Begin
        If n_预约顺序号 Is Null Then
          Select 操作员姓名, 工作站名称
          Into v_序号操作员, v_序号机器名
          From 临床出诊序号控制
          Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_序号;
        Else
          Select 操作员姓名, 工作站名称
          Into v_序号操作员, v_序号机器名
          From 临床出诊序号控制
          Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号;
        End If;
        n_锁定 := 1;
      Exception
        When Others Then
          v_序号操作员 := Null;
          v_序号机器名 := Null;
          n_锁定       := 0;
      End;

      If n_锁定 = 0 Then
        If n_预约顺序号 Is Null Then
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
        End If;

        If Sql%Rowcount = 0 Then
          Begin
            If Nvl(n_分时段, 0) > 0 Then
              If Nvl(n_序号控制, 0) = 1 Then
                --分时段后专家号 失约的预约号允许挂号
                Update 临床出诊序号控制
                Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
                Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) In (0, 2);
                If Sql%NotFound Then
                  Begin
                    Select 挂号状态 Into n_状态 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 序号 = n_序号;
                  Exception
                    When Others Then
                      n_状态 := -1;
                  End;

                  If n_状态 <> -1 Then
                    v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
                    Raise Err_Item;
                  End If;

                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                    Select 出诊记录id_In, n_序号, d_序号时间, d_序号时间, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1), Null,
                           Null, Null, 操作员姓名_In, '追加号'
                    From Dual;
                End If;
              Else
                If Nvl(预约接收_In, 0) = 1 Then
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注, 预约顺序号)
                    Select 记录id, 序号, 开始时间, 终止时间, 1, 1, Decode(预约挂号_In, 1, 2, 1), Null, Null, Null, 操作员姓名_In, n_序号,
                           n_预约顺序号
                    From 临床出诊序号控制
                    Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Null;
                End If;
              End If;
            Else
              If Nvl(n_序号控制, 0) = 1 Then
                Update 临床出诊序号控制
                Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
                Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 0;

                If Sql%Rowcount = 0 Then
                  Begin
                    Select 挂号状态 Into n_状态 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 序号 = n_序号;
                  Exception
                    When Others Then
                      n_状态 := -1;
                  End;
                  If n_状态 <> -1 Then
                    v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
                    Raise Err_Item;
                  End If;
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                    Select 出诊记录id_In, n_序号, 发生时间_In, 发生时间_In, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1),
                           Null, Null, Null, 操作员姓名_In, '追加号'
                    From Dual;

                End If;
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
          v_Err_Msg := '序号' || n_序号 || '已被其他站点(' || v_机器名 || ')锁定,请重新选择一个序号.';
          Raise Err_Item;
        End If;
        If n_预约顺序号 Is Null Then
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And 工作站名称 = v_机器名;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And
                工作站名称 = v_机器名;
        End If;
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
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 And 序号_In = 1 Then
      v_结算内容     := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      v_结算方式记录 := '';
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);

        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));

        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);

        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);

        If Instr('|' || v_结算方式记录 || '|', '|' || Nvl(v_结算方式, v_现金) || '|') <> 0 Then
          v_Err_Msg := '使用了重复的结算方式,请检查!';
          Raise Err_Item;
        Else
          v_结算方式记录 := v_结算方式记录 || '|' || Nvl(v_结算方式, v_现金);
        End If;

        If n_三方卡标志 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, 合作单位_In, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4,
             v_结算号码);
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
            Zl_病人卡结算记录_Insert(结算卡序号_In, n_消费卡id, Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 卡号_In, Null, 登记时间_In, Null, 结帐id_In,
                              n_预交id);
          End If;
        End If;

        If Nvl(更新交款余额_In, 0) = 0 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + n_结算金额
          Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金)
          Returning 余额 Into n_返回值;

          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, Nvl(v_结算方式, v_现金), 1, n_结算金额);
            n_返回值 := n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End If;

        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
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
      Select Nvl(Sum(Nvl(预交余额, 0) - Nvl(费用余额, 0)), 0)
      Into n_病人余额
      From 病人余额
      Where 病人id = 病人id_In And Nvl(性质, 0) = 1 And Nvl(类型, 0) = 1;
      if n_病人余额 < 预交支付_In Then
        v_Err_Msg := '病人的当前预交余额为 ' || Ltrim(To_Char(n_病人余额, '9999999990.00')) || '，小于本次支付金额 ' ||
                     Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，支付失败！';
        Raise Err_Item;
      End if;
      
      n_预交金额 := 预交支付_In;
      For r_Deposit In c_Deposit(病人id_In) Loop
        n_当前金额 := Case
                    When r_Deposit.金额 - n_预交金额 < 0 Then
                     r_Deposit.金额
                    Else
                     n_预交金额
                  End;

        If r_Deposit.结帐id = 0 Then
          --第一次冲预交(填上结帐ID,金额为0)
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;

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
        Where 病人id = 病人id_In And 性质 = 1 And 类型 = Nvl(1, 2);
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
    If 序号_In = 1 And Nvl(更新交款余额_In, 0) = 0 Then
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
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 出诊记录id, 收费单)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 出诊记录id_In, 收费单_In);

    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = 发生时间_In, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
      Where ID = n_挂号id;
    End If;

    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;

    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
        n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(预约挂号_In, 0) = 1 And n_分时点显示 = 1 And n_分时段 = 1 Then
          n_分时点显示 := 1;
        Else
          n_分时点显示 := Null;
        End If;
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         n_分时点显示, v_排队序号);
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
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));

    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) >= Sysdate;
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
End Zl_病人挂号记录_出诊_Insert;
/

--134441:李南春,2019-01-09,挂号检查项目是否一致
CREATE OR REPLACE Procedure Zl_Third_Registercheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS挂号检查
  --入参:Xml_In:
  --<IN>
  --  <BRID>1</BRID>                    //病人ID
  --  <XM>姓名</XM>                     //姓名
  --  <SFZH>510221197008184710</SFZH>   //身份证号
  --  <HM>0100</HM>                     //号码
  --  <CZJLID>100</CZJLID>              //出诊记录ID,计划排班模式可以不传
  --  <GHSJ>2016-08-10 09:52:00</GHSJ>  //挂号时间
  --  <KSID>1</KSID>                    //科室ID
  --  <YSXM>张震</YSXM>                 //医生姓名
  --  <GHXMID>1</GHXMID>                 //挂号主项目，不传时不检查
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <ERROR><MSG></MSG></ERROR> //为空表示检查成功
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_病人id         病人信息.病人id%Type;
  v_姓名           病人信息.姓名%Type;
  v_身份证号       病人信息.身份证号%Type;
  v_号码           挂号安排.号码%Type;
  n_项目ID         挂号安排.项目ID%Type;
  n_出诊记录id     Number(18);
  d_发生时间       病人挂号记录.发生时间%Type;
  v_Para           Varchar2(500);
  d_启用时间       Date;
  n_挂号模式       Number(3);
  n_同科限号数     Number;
  n_同科限约数     Number;
  n_病人挂号科室数 Number;
  n_病人预约科室数 Number;
  n_专家号挂号限制 Number;
  n_专家号预约限制 Number;
  n_Exists         Number;
  n_Count          Number;
  n_科室id         病人挂号记录.执行部门id%Type;
  v_医生姓名       病人挂号记录.执行人%Type;
  v_性别           病人信息.性别%Type;
  v_年龄           病人信息.年龄%Type;
  n_已约科室       Number;
  v_Checkresult    Varchar2(500);
  v_Temp           Varchar2(32767); --临时XML
  x_Templet        Xmltype; --模板XML
  v_Err_Msg        Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/CZJLID'),
         To_Date(Extractvalue(Value(A), 'IN/GHSJ'), 'yyyy-mm-dd hh24:mi:ss'),
         To_Number(Extractvalue(Value(A), 'IN/KSID')), Extractvalue(Value(A), 'IN/YSXM'),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM'),
         To_Number(Extractvalue(Value(A), 'IN/GHXMID'))
  Into n_病人id, v_号码, n_出诊记录id, d_发生时间, n_科室id, v_医生姓名, v_身份证号, v_姓名, n_项目ID
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := zl_third_GetPatiID(v_身份证号,v_姓名);
  End If;
  If nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查';
    Raise Err_Item;
  End If;

  v_Para := zl_GetSysParameter(256);
  If v_Para Is Not Null Then
    n_挂号模式 := Substr(v_Para, 1, 1);
    Begin
      d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;

    If Sysdate - 10 > Nvl(d_启用时间, Sysdate - 30) Then
      If n_挂号模式 = 1 And Nvl(d_发生时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_出诊记录id Is Null Then
        v_Temp := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    Else
      If n_挂号模式 = 1 And Nvl(d_发生时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_出诊记录id Is Null Then
        Begin
          Select a.Id
          Into n_出诊记录id
          From 临床出诊记录 A, 临床出诊号源 B
          Where a.号源id = b.Id And b.号码 = v_号码 And Nvl(d_发生时间, Sysdate) Between a.开始时间 And a.终止时间;
        Exception
          When Others Then
            v_Temp := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
            v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
            Xml_Out := x_Templet;
            Return;
        End;
      End If;
    End If;
  End If;

  If n_出诊记录id Is Not Null Then
    Select 性别, 年龄 Into v_性别, v_年龄 From 病人信息 Where 病人id = n_病人id And Rownum < 2;
    v_Checkresult := Zl_临床出诊限制_Check(n_出诊记录id, v_年龄, v_性别);
    If Substr(Nvl(v_Checkresult, '0'), 1, 1) <> '0' Then
      v_Temp := '病人不适用该本号别,请检查！';
      v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Xml_Out := x_Templet;
      Return;
    End If;

    for C_排班 In (Select a.号码, b.科室id, a.项目id From 临床出诊号源 a, 临床出诊记录 b
                    where a.id = b.号源id And b.Id = n_出诊记录id) loop
      v_Temp  := Null;
      n_Count := 1;
      if v_号码 <> C_排班.号码 then
        v_Temp := '挂号信息的号码错误,请检查！';
      Elsif n_科室id <> C_排班.科室id then
        v_Temp := '挂号信息的科室错误,请检查！';
      Elsif n_项目id <> C_排班.项目id And Nvl(n_项目id, 0) <> 0 then

        v_Temp := '挂号信息的收费项目错误,请检查！';
      end IF;
      IF v_Temp is not null Then
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End if;
    end loop;

    IF NVL(n_Count, 0) <> 1 Then
      v_Temp := '挂号信息错误,请重试！';
      v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Xml_Out := x_Templet;
      Return;
    End IF;
  End If;

  If Trunc(Sysdate) > Trunc(d_发生时间) Then
    v_Temp := '不能挂以前的号(' || To_Char(d_发生时间, 'yyyy-mm-dd') || ')。';
    v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  v_Temp := Zl_Identity(0);
  If Nvl(v_Temp, ' ') = ' ' Then
    v_Temp := '当前操作人员未设置对应的人员关系,不能继续。';
    v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  v_Temp           := Nvl(zl_GetSysParameter('病人同科限挂N个号', 1111), '0|0') || '|';
  n_同科限号数     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
  n_同科限约数     := To_Number(Nvl(zl_GetSysParameter('病人同科限约N个号', 1111), '0'));
  n_病人预约科室数 := To_Number(Nvl(zl_GetSysParameter('病人预约科室数', 1111), '0'));
  n_病人挂号科室数 := To_Number(Nvl(zl_GetSysParameter('病人挂号科室限制', 1111), '0'));
  n_专家号挂号限制 := To_Number(Nvl(zl_GetSysParameter('专家号挂号限制'), '0'));
  n_专家号预约限制 := To_Number(Nvl(zl_GetSysParameter('专家号预约限制'), '0'));

  If Trunc(Sysdate) <> Trunc(d_发生时间) Then
    If Nvl(n_病人预约科室数, 0) <> 0 Then
      n_已约科室 := 0;
      For c_Chkitem In (Select Distinct 执行部门id
                        From 病人挂号记录
                        Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(d_发生时间) And
                              Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
        n_已约科室 := n_已约科室 + 1;
      End Loop;
      If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
        v_Temp := '同一病人最多同时能预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
    If Nvl(n_同科限约数, 0) <> 0 Then
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(d_发生时间) And
            Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
      If n_Count >= Nvl(n_同科限约数, 0) And Nvl(n_同科限约数, 0) > 0 Then
        v_Temp := '该病人已经在该科室预约了' || n_Count || '次,不能再预约！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
  Else
    If Nvl(n_病人挂号科室数, 0) <> 0 Then
      n_已约科室 := 0;
      For c_Chkitem In (Select Distinct 执行部门id
                        From 病人挂号记录
                        Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(d_发生时间) And
                              Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
        n_已约科室 := n_已约科室 + 1;
      End Loop;
      If n_已约科室 >= Nvl(n_病人挂号科室数, 0) And Nvl(n_病人挂号科室数, 0) > 0 Then
        v_Temp := '同一病人最多同时能挂号[' || Nvl(n_病人挂号科室数, 0) || ']个科室,不能再挂号！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
    If Nvl(n_同科限号数, 0) <> 0 Then
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(d_发生时间) And
            Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
      If n_Count >= Nvl(n_同科限号数, 0) And Nvl(n_同科限号数, 0) > 0 Then
        v_Temp := '该病人已经在该科室挂号了' || n_Count || '次,不能再挂号！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
  End If;

  If Trunc(Sysdate) = Trunc(d_发生时间) Then
    --挂号
    If Nvl(n_专家号挂号限制, 0) <> 0 And v_医生姓名 Is Not Null Then
      If n_出诊记录id Is Null Then
        --无出诊记录对应
        Begin
          Select Count(1)
          Into n_Exists
          From 病人挂号记录
          Where 病人id = n_病人id And 号别 = v_号码 And 发生时间 Between Trunc(d_发生时间) And Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And
                记录状态 = 1 And 记录性质 = 1;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_专家号挂号限制 Then
          v_Temp := '该病人已经超过本号挂号限制,不能再次挂号！';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      Else
        --对应出诊记录
        Begin
          Select Count(1)
          Into n_Exists
          From 病人挂号记录
          Where 病人id = n_病人id And 出诊记录id = n_出诊记录id And 记录状态 = 1 And 记录性质 = 1;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_专家号挂号限制 Then
          v_Temp := '该病人已经超过本号挂号限制,不能再次挂号！';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      End If;
    End If;
  Else
    --预约
    If Nvl(n_专家号预约限制, 0) <> 0 And v_医生姓名 Is Not Null Then
      If n_出诊记录id Is Null Then
        --无出诊记录对应
        Begin
          Select Count(1)
          Into n_Exists
          From 病人挂号记录
          Where 病人id = n_病人id And 号别 = v_号码 And 发生时间 Between Trunc(d_发生时间) And Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And
                记录状态 = 1 And 记录性质 = 2;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_专家号预约限制 Then
          v_Temp := '该病人已经超过本号预约限制,不能再次预约！';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      Else
        --对应出诊记录
        Begin
          Select Count(1)
          Into n_Exists
          From 病人挂号记录
          Where 病人id = n_病人id And 出诊记录id = n_出诊记录id And 记录状态 = 1 And 记录性质 = 2;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_专家号预约限制 Then
          v_Temp := '该病人已经超过本号预约限制,不能再次预约！';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      End If;
    End If;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registercheck;
/

--134969:李南春,2019-01-11,预交支付检查
--134441:李南春,2019-01-09,挂号检查项目是否一致
CREATE OR REPLACE Procedure Zl_三方机构挂号_Insert
(
  操作方式_In     Integer,
  病人id_In       门诊费用记录.病人id%Type,
  号码_In         挂号安排.号码%Type,
  号序_In         挂号序号状态.序号%Type,
  单据号_In       门诊费用记录.No%Type,
  票据号_In       门诊费用记录.实际票号%Type,
  结算方式_In     Varchar2,
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
  更新年龄_In     Number := 0,
  购买病历_In     Number := 0,
  出诊记录id_In   临床出诊记录.Id%Type := Null
) As
  --功能：三方机构进行挂号(包含预约;预约挂号不扣款;预约挂号扣款)
  --入参:操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
  --      结算方式_IN:支持多种结算方式,多种结算方式时，传入格式如下:结算方式名称1,金额,结算号码,三方卡标志|结算方式名称2,金额,结算号码,三方卡标志|...
  --      加入序号状态_In:1-表示强制加入挂号序号状态表中;否则根据启用序号或启用时段时才加入.
  --      是否自助设备_In:1-表示是医院的自助设备进行此函数的调用,自助设备调用此函数 允许加号,否则不允许
  --      锁定类型_In :0-产生正常数据 1-表示对单据进行锁定,产生未生效的单据信息;2-对锁定的记录进行解锁-生成正常数据:未生效的单据在银行扣款完成后进行解锁
  --      保险结算_IN:格式="结算方式|结算金额||....."
  Err_Item Exception;
  Err_Special Exception;
  v_Err_Msg            Varchar2(255);
  n_打印id             票据打印内容.Id%Type;
  n_返回值             病人预交记录.金额%Type;
  v_排队号码           Varchar2(20);
  v_队列名称           排队叫号队列.队列名称%Type;
  n_预交id             病人预交记录.Id%Type;
  n_挂号id             病人挂号记录.Id%Type;
  v_结算内容           Varchar2(3000);
  v_当前结算           Varchar2(150);
  d_发生时间           Date;
  v_结算方式           病人预交记录.结算方式%Type;
  n_结算金额           病人预交记录.冲预交%Type;
  n_结算合计           Number(16, 5);
  n_预交金额           病人预交记录.冲预交%Type;
  n_病人余额           病人余额.预交余额%Type;
  n_组id               财务缴款分组.Id%Type;
  d_排队时间           Date;
  n_锁定               Number;
  n_病人预约科室数     Number(18);
  n_已约科室           Number(18);
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
  v_收费项目ids        Varchar2(300);
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
  v_排队序号           排队叫号队列.排队序号%Type;
  v_机器名             挂号序号状态.机器名%Type;
  v_序号操作员         挂号序号状态.操作员姓名%Type;
  v_序号机器名         挂号序号状态.机器名%Type;
  n_序号锁定           Number := 0;
  n_病历费id           收费特定项目.收费细目id%Type;
  v_付款方式           病人挂号记录.医疗付款方式%Type;
  v_费别               门诊费用记录.费别%Type;
  n_屏蔽费别           Number(3) := 0;
  n_Tmp安排id          挂号安排.Id%Type;
  n_计划id             挂号安排计划.Id%Type;
  v_年龄               病人信息.年龄%Type;
  n_合作单位限数量模式 Number;
  n_出诊记录id         临床出诊记录.Id%Type;
  n_挂号模式           Number(3);
  n_同科限号数         Number;
  n_同科限约数         Number;
  n_病人挂号科室数     Number;
  n_分时点显示         Number;
  d_启用时间           Date;
  n_Exists             Number;
  v_Exists             Varchar2(4000);
  v_Para               Varchar2(2000);
  n_专家号挂号限制     Number;
  n_专家号预约限制     Number;
  v_时间段             时间段.时间段%Type;
  d_检查开始时间       时间段.开始时间%Type;
  d_检查结束时间       时间段.终止时间%Type;

  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.费别, a.险类, c.编码 As 付款方式
    From 病人信息 A, 医疗付款方式 C
    Where a.病人id = n_病人id And a.医疗付款方式 = c.名称(+);

  r_Pati c_Pati%RowType;

  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit(v_病人id 病人信息.病人id%Type) Is
    Select NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id = v_病人id And Nvl(预交类别, 2) = 1 Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO
    Order By 结帐id, NO;

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
                           p.周四 As 四, p.周五 As 五, p.周六 As 六, p.序号控制, p.计划id
           From (Select p.Id, p.号码, p.号类, p.科室id, p.项目id, p.医生id, p.医生姓名, b.限号数, b.限约数, Nvl(p.病案必须, 0) As 病案必须, p.周日, p.周一,
                         p.周二, p.周三, p.周四, p.周五, p.周六, p.分诊方式, p.序号控制,
                         Decode(To_Char(d_发生时间_In, 'D'), '1', p.周日, '2', p.周一, '3', p.周二, '4', p.周三, '5', p.周四, '6', p.周五,
                                 '7', p.周六, Null) As 排班, Null As 计划id
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
                                 '7', p.周六, Null) As 排班, p.Id As 计划id
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

  Procedure Zl_三方机构挂号_出诊_Insert
  (
    记录id_In       临床出诊记录.Id%Type,
    操作方式_In     Integer,
    病人id_In       门诊费用记录.病人id%Type,
    号码_In         挂号安排.号码%Type,
    号序_In         挂号序号状态.序号%Type,
    单据号_In       门诊费用记录.No%Type,
    票据号_In       门诊费用记录.实际票号%Type,
    结算方式_In     Varchar2,
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
    更新年龄_In     Number := 0,
    购买病历_In     Number := 0
  ) As
    --功能：三方机构进行挂号(包含预约;预约挂号不扣款;预约挂号扣款),出诊表排班模式下使用
    --入参: 操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
    --      加入序号状态_In:1-表示强制加入挂号序号状态表中;否则根据启用序号或启用时段时才加入.
    --      是否自助设备_In:1-表示是医院的自助设备进行此函数的调用,自助设备调用此函数 允许加号,否则不允许
    --      锁定类型_In :0-产生正常数据 1-表示对单据进行锁定,产生未生效的单据信息;2-对锁定的记录进行解锁-生成正常数据:未生效的单据在银行扣款完成后进行解锁
    --      保险结算_IN:格式="结算方式|结算金额||....."
    --      冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
    Err_Item Exception;
    Err_Special Exception;
    v_Err_Msg            Varchar2(255);
    n_打印id             票据打印内容.Id%Type;
    n_返回值             病人预交记录.金额%Type;
    v_排队号码           Varchar2(20);
    v_队列名称           排队叫号队列.队列名称%Type;
    n_预交id             病人预交记录.Id%Type;
    n_挂号id             病人挂号记录.Id%Type;
    v_结算内容           Varchar2(3000);
    v_当前结算           Varchar2(150);
    v_结算方式           病人预交记录.结算方式%Type;
    n_结算金额           病人预交记录.冲预交%Type;
    n_结算合计           Number(16, 5);
    n_预交金额           病人预交记录.冲预交%Type;
    n_病人余额           病人余额.预交余额%Type;
    n_组id               财务缴款分组.Id%Type;
    d_排队时间           Date;
    n_锁定               Number;
    n_病人预约科室数     Number(18);
    n_已约科室           Number(18);
    d_发生时间           Date;
    n_合作单位限制       Number(18);
    n_是否开放           Number(1);
    n_Count              Number(18);
    n_行号               Number(18);
    v_号别               病人挂号记录.号别%Type;
    n_序号               病人挂号记录.号序%Type;
    n_费用id             门诊费用记录.Id%Type;
    n_价格父号           Number(18);
    n_原项目id           收费项目目录.Id%Type;
    n_原收入项目id       收费项目目录.Id%Type;
    v_诊室               病人挂号记录.诊室%Type;
    n_实收金额合计       门诊费用记录.实收金额%Type;
    n_开单部门id         门诊费用记录.开单部门id%Type;
    n_实收金额           门诊费用记录.实收金额%Type;
    n_应收金额           门诊费用记录.实收金额%Type;
    n_结帐id             病人结帐记录.Id%Type;
    v_Temp               Varchar2(500);
    v_结算方式记录       Varchar2(1000);
    n_预约时段序号       Number;
    n_序号控制           临床出诊记录.是否序号控制%Type;
    n_限约数             临床出诊记录.限约数%Type;
    n_项目id             临床出诊记录.项目id%Type;
    n_科室id             临床出诊记录.科室id%Type;
    d_终止时间           临床出诊记录.终止时间%Type;
    v_医生姓名           临床出诊记录.医生姓名%Type;
    n_医生id             临床出诊记录.医生id%Type;
    n_预约顺序号         临床出诊序号控制.预约顺序号%Type;
    n_预约总数           Number;
    d_时段开始时间       Date;
    d_时段终止时间       Date;
    v_收费项目ids        Varchar2(300);
    n_三方卡标志         Number;
    n_号序               病人挂号记录.号序%Type;
    d_登记时间           Date;
    n_单笔金额           病人预交记录.冲预交%Type;
    v_结算号码           病人预交记录.结算号码%Type;
    v_操作员编号         人员表.编号%Type;
    v_操作员姓名         人员表.姓名%Type;
    n_预约               Integer;
    v_现金               病人预交记录.结算方式%Type;
    n_启用分时段         Integer;
    n_已挂数             病人挂号汇总.已挂数%Type;
    n_已约数             病人挂号汇总.已约数%Type;
    n_其中已接收         病人挂号汇总.已约数%Type;
    n_预约生成队列       Number;
    n_限号数             临床出诊记录.限号数%Type;
    d_Date               Date;
    n_挂号序号           Number;
    v_排队序号           排队叫号队列.排队序号%Type;
    v_机器名             挂号序号状态.机器名%Type;
    v_序号操作员         挂号序号状态.操作员姓名%Type;
    v_序号机器名         挂号序号状态.机器名%Type;
    n_序号锁定           Number := 0;
    n_病历费id           收费特定项目.收费细目id%Type;
    v_付款方式           病人挂号记录.医疗付款方式%Type;
    v_费别               门诊费用记录.费别%Type;
    n_屏蔽费别           Number(3) := 0;
    v_年龄               病人信息.年龄%Type;
    n_合作单位限数量模式 Number;
    n_同科限号数         Number;
    n_分时点显示         Number;
    n_同科限约数         Number;
    n_病人挂号科室数     Number;
    n_Exists             Number(5);
    n_替诊医生id         临床出诊记录.替诊医生id%Type;
    v_替诊医生姓名       临床出诊记录.替诊医生姓名%Type;
    d_替诊开始时间       临床出诊记录.替诊开始时间%Type;
    d_替诊终止时间       临床出诊记录.替诊终止时间%Type;
    n_专家号挂号限制     Number;
    n_专家号预约限制     Number;

    Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
      Select a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.费别, a.险类, c.编码 As 付款方式
      From 病人信息 A, 医疗付款方式 C
      Where a.病人id = n_病人id And a.医疗付款方式 = c.名称(+);

    r_Pati c_Pati%RowType;

    --该游标用于收费冲预交的可用预交列表
    --以ID排序，优先冲上次未冲完的。
    Cursor c_Deposit(v_病人id 病人信息.病人id%Type) Is
      Select NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
             Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
      From 病人预交记录
      Where 记录性质 In (1, 11) And 病人id = v_病人id And Nvl(预交类别, 2) = 1 Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
      Group By NO
      Order By 结帐id, NO;

    Function Zl_诊室(记录id_In 临床出诊记录.Id%Type) Return Varchar2 As
      n_分诊方式 临床出诊记录.分诊方式%Type;
      v_诊室     病人挂号记录.诊室%Type;
      v_Rowid    Varchar2(500);
      n_Next     Integer;
      n_First    Integer;
    Begin

      If 锁定类型_In = 2 Then
        --对单据进行解锁,首先检查是否存在锁定
        Select Count(Rowid)
        Into n_锁定
        From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
        If n_锁定 = 0 Then
          v_Err_Msg := '单据号为(' || 单据号_In || ')的单据,不存在或者已经被解锁!';
          Raise Err_Item;
        End If;
        Select Max(号序) Into n_号序 From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      End If;

      Begin
        Select Nvl(分诊方式, 0) Into n_分诊方式 From 临床出诊记录 Where ID = 记录id_In;
      Exception
        When Others Then
          v_Err_Msg := '出诊记录(' || 记录id_In || ')未找到!';
          Raise Err_Item;
      End;

      --0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
      v_诊室 := Null;
      If n_分诊方式 = 1 Then
        --1-指定诊室
        Begin
          Select b.名称 Into v_诊室 From 临床出诊诊室记录 A, 门诊诊室 B Where a.诊室id = b.Id And a.记录id = 记录id_In;
        Exception
          When Others Then
            v_诊室 := Null;
        End;
      End If;
      If n_分诊方式 = 2 Then
        --2-动态分诊:该个号别当天挂号未诊数最少的诊室
        For c_诊室 In (Select 门诊诊室, Sum(Num) As Num
                     From (Select b.名称 As 门诊诊室, 0 As Num
                            From 临床出诊诊室记录 A, 门诊诊室 B
                            Where a.诊室id = b.Id And a.记录id = 记录id_In
                            Union All
                            Select 诊室, Count(诊室) As Num
                            From 病人挂号记录
                            Where Nvl(执行状态, 0) = 0 And 发生时间 Between Trunc(Sysdate) And Sysdate And 号别 = 号码_In And
                                  诊室 In (Select d.名称
                                         From 临床出诊诊室记录 C, 门诊诊室 D
                                         Where c.诊室id = d.Id And c.记录id = 记录id_In)
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
        For c_诊室 In (Select a.Rowid As Rid, b.名称 As 门诊诊室, a.当前分配
                     From 临床出诊诊室记录 A, 门诊诊室 B
                     Where a.诊室id = b.Id And a.记录id = 记录id_In) Loop
          If n_First = 1 Then
            v_Rowid := c_诊室.Rid;
          End If;
          If n_Next = 1 Then
            v_诊室 := c_诊室.门诊诊室;
            Update 临床出诊诊室记录 Set 当前分配 = 1 Where Rowid = c_诊室.Rid;
            Exit;
          End If;
          If Nvl(c_诊室.当前分配, 0) = 1 Then
            Update 临床出诊诊室记录 Set 当前分配 = 0 Where Rowid = c_诊室.Rid;
            n_Next := 1;
          End If;
        End Loop;
        If v_诊室 Is Null Then
          Update 临床出诊诊室记录 Set 当前分配 = 1 Where Rowid = v_Rowid Returning 诊室id Into v_诊室;
          Select 名称 Into v_诊室 From 门诊诊室 Where ID = v_诊室;
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
    d_发生时间 := 发生时间_In;

    If d_发生时间 Is Null Then
      d_发生时间 := Sysdate;
    End If;

    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;

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
    Where 出诊记录id = 记录id_In And 登记时间 Between Trunc(发生时间_In) And Trunc(发生时间_In + 1) - 1 / 24 / 60 / 60;

    If 登记时间_In Is Null Then
      d_登记时间 := Sysdate;
    Else
      d_登记时间 := 登记时间_In;
    End If;
    If Trunc(Sysdate) > Trunc(发生时间_In) Then
      v_Err_Msg := '不能挂以前的号(' || To_Char(发生时间_In, 'yyyy-mm-dd') || ')。';
      Raise Err_Item;
    End If;

    v_Temp           := Nvl(zl_GetSysParameter('病人同科限挂N个号', 1111), '0|0') || '|';
    n_同科限号数     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_同科限约数     := To_Number(Nvl(zl_GetSysParameter('病人同科限约N个号', 1111), '0'));
    n_病人预约科室数 := To_Number(Nvl(zl_GetSysParameter('病人预约科室数', 1111), '0'));
    n_病人挂号科室数 := To_Number(Nvl(zl_GetSysParameter('病人挂号科室限制', 1111), '0'));
    n_专家号挂号限制 := To_Number(Nvl(zl_GetSysParameter('专家号挂号限制'), '0'));
    n_专家号预约限制 := To_Number(Nvl(zl_GetSysParameter('专家号预约限制'), '0'));

    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Item;
    End If;
    n_开单部门id := To_Number(Zl_操作员(0, v_Temp));
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
    n_组id       := Zl_Get组id(v_操作员姓名);

    --支付宝并发提交检查
    Select Nvl(Max(1), 0)
    Into n_Exists
    From 病人挂号记录
    Where 病人id = 病人id_In And 号别 = 号码_In And 号序 = 号序_In And 操作员姓名 = v_操作员姓名 And Nvl(记录id_In, 0) = Nvl(出诊记录id, 0) And
          登记时间 > Sysdate - 0.01 And 记录状态 = 1 And 发生时间 = 发生时间_In;
    If n_Exists = 1 Then
      v_Err_Msg := '病人已经挂号,不能重复挂相同的号！';
      Raise Err_Special;
    End If;

    If 操作方式_In <> 1 Then
      --预约检查是否添加合作单位控制
      --如果设置了合作单位控制 则
      Begin
        Select 1
        Into n_合作单位限制
        From 临床出诊挂号控制记录
        Where 记录id = 记录id_In And 类型 = 1 And 性质 = 1 And 控制方式 <> 4 And Rownum < 2;
      Exception
        When Others Then
          n_合作单位限制 := 0;
      End;
    End If;

    If 操作方式_In <> 2 Then
      v_诊室 := Zl_诊室(记录id_In);
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

    Begin
      Select Nvl(a.是否分时段, 0), a.限号数, a.已挂数, a.其中已接收, a.已约数, a.是否序号控制, a.限约数, a.项目id, a.科室id, a.医生id, a.医生姓名, a.替诊医生id,
             a.替诊医生姓名, a.替诊开始时间, a.替诊终止时间, b.号码
      Into n_启用分时段, n_限号数, n_已挂数, n_其中已接收, n_已约数, n_序号控制, n_限约数, n_项目id, n_科室id, n_医生id, v_医生姓名, n_替诊医生id, v_替诊医生姓名,
           d_替诊开始时间, d_替诊终止时间, v_号别
      From 临床出诊记录 a, 临床出诊号源 b
      Where a.ID = 记录id_In and a.号源id = b.id And Nvl(a.是否锁定, 0) = 0;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '该号别没有在' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '中进行安排。';
      Raise Err_Item;
    End If;

    IF v_号别 <> 号码_In Then
      v_Err_Msg := '号别和出诊记录不一致，请检查。';
      Raise Err_Item;
    End IF;

    If 发生时间_In Between Nvl(d_替诊开始时间, Sysdate) And Nvl(d_替诊终止时间, Sysdate - 1) And v_替诊医生姓名 Is Not Null Then
      n_医生id   := n_替诊医生id;
      v_医生姓名 := v_替诊医生姓名;
    End If;

    --对参数控制进行检查
    --仅在预约不扣款时进行检查
    If 操作方式_In = 2 Then
      If Nvl(n_同科限约数, 0) <> 0 Or Nvl(n_病人预约科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
          Raise Err_Item;
        End If;

        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
        If n_Count >= Nvl(n_同科限约数, 0) And Nvl(n_同科限约数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室预约了' || n_Count || '次,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_专家号预约限制, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 出诊记录id = 记录id_In;
        If n_Count >= Nvl(n_专家号预约限制, 0) And Nvl(n_专家号预约限制, 0) > 0 Then
          v_Err_Msg := '该病人已经超过本号预约限制,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_同科限号数, 0) <> 0 Or Nvl(n_病人挂号科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人挂号科室数, 0) And Nvl(n_病人挂号科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能挂号[' || Nvl(n_病人挂号科室数, 0) || ']个科室,不能再挂号！';
          Raise Err_Item;
        End If;

        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
        If n_Count >= Nvl(n_同科限号数, 0) And Nvl(n_同科限号数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室挂号了' || n_Count || '次,不能再挂号！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_专家号挂号限制, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 出诊记录id = 记录id_In;
        If n_Count >= Nvl(n_专家号挂号限制, 0) And Nvl(n_专家号挂号限制, 0) > 0 Then
          v_Err_Msg := '该病人已经超过本号挂号限制,不能再挂号！';
          Raise Err_Item;
        End If;
      End If;
    End If;

    d_Date         := Null;
    d_时段开始时间 := Null;

    If Nvl(n_限号数, 0) >= 0 Or n_限号数 Is Null Then
      If n_启用分时段 = 1 Then
        If Nvl(n_序号控制, 0) = 1 Then
          If Nvl(是否自助设备_In, 0) = 0 Then
            Select Count(*), Max(开始时间)
            Into n_Count, d_时段开始时间
            From 临床出诊序号控制
            Where 记录id = 记录id_In And 序号 = Nvl(号序_In, 0);

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
                                To_Date(To_Char(Sysdate + Decode(Sign(开始时间 - 终止时间), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                         To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 终止时间, 数量, 是否预约
                         From 临床出诊序号控制
                         Where 记录id = 记录id_In And 序号 = Nvl(号序_In, 0)) Loop
              If Sysdate > v_时段.终止时间 Then
                v_Err_Msg := '号别为' || 号码_In || '及号序为' || Nvl(号序_In, 0) || '的安排,已经超过时点,不能再' || v_Temp || '！';
                Raise Err_Item;
              End If;
            End Loop;
          End If;
        Elsif 操作方式_In > 1 Then
          --未启用序号的,需要检查预约的情况
          n_Count := 0;
          For v_时段 In (Select 序号, 开始时间, 终止时间, 数量, 是否预约
                       From 临床出诊序号控制
                       Where 记录id = 记录id_In And
                             (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                             Decode(Sign(开始时间 - 终止时间 - 1 / 24 / 60 / 60), 1,
                                      '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')) And
                             '3000-01-10 ' || To_Char(终止时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                             ('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                             '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS') And
                             Decode(Sign(开始时间 - 终止时间 - 1 / 24 / 60 / 60), 1,
                                      '3000-01-11 ' || To_Char(终止时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(终止时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
            n_预约时段序号 := v_时段.序号;
            d_时段开始时间 := v_时段.开始时间;
            d_时段终止时间 := v_时段.终止时间;

            Select Count(*), Max(序号), Max(预约顺序号) + 1
            Into n_Count, n_预约总数, n_预约顺序号
            From 临床出诊序号控制
            Where 记录id = 记录id_In And Nvl(挂号状态, 0) Not In (0, 4, 5);

            If Nvl(n_Count, 0) > Nvl(v_时段.数量, 0) And 锁定类型_In <> 2 Then
              v_Err_Msg := '号别为' || 号码_In || '的挂号安排中在' || To_Char(v_时段.开始时间, 'hh24:mi:ss') || '至' ||
                           To_Char(v_时段.终止时间, 'hh24:mi:ss') || '最多只能预约' || Nvl(v_时段.数量, 0) || '人,不能再进行预约挂号！';
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
      If n_已挂数 >= Nvl(n_限号数, 0) And n_限号数 Is Not Null Then
        v_Err_Msg := '该号别今天已达到限号数 ' || Nvl(n_限号数, 0) || '不能再挂号！';
        Raise Err_Item;
      End If;
    End If;

    If 操作方式_In > 1 Then
      --预约的相关检查
      --规则:
      --   1.已限约不能超过限约数
      --   2.检查是否启用时段的
      If n_已约数 >= Nvl(n_限约数, 0) And Nvl(n_限约数, 0) <> 0 And n_限约数 Is Not Null And 锁定类型_In <> 2 Then
        v_Err_Msg := '该号别已达到限约数 ' || Nvl(n_限约数, 0) || '不能再预约挂号！';
        Raise Err_Item;
      End If;
      If 预约方式_In Is Not Null Then
        Select Zl_Fun_Get临床出诊预约状态(记录id_In, 发生时间_In, 号序_In, 预约方式_In, Null, 0, v_操作员姓名, v_机器名)
        Into v_Exists
        From Dual;
        If To_Number(Substr(v_Exists, 1, 1)) <> 0 Then
          v_Err_Msg := '传入的预约方式' || 预约方式_In || '不可用,原因:' || Substr(v_Exists, 3);
          Raise Err_Item;
        End If;
      End If;
    End If;
    If n_合作单位限制 > 0 And 操作方式_In <> 1 And 合作单位_In Is Not Null Then
      If Nvl(n_序号控制, 0) = 1 And Nvl(号序_In, 0) = 0 Then
        v_Err_Msg := '当前安排使用了序号控制,请确认所需要预约的序号,不能继续。';
        Raise Err_Item;
      End If; --Nvl(r_安排.序号控制, 0) =0

      n_序号 := Case
                When Nvl(n_序号控制, 0) = 1 Or n_启用分时段 = 1 And 操作方式_In > 1 Then
                 Nvl(号序_In, 0)
                Else
                 0
              End;

      --合作单位控制模式
      Begin
        Select Nvl(控制方式, 0)
        Into n_合作单位限数量模式
        From 临床出诊挂号控制记录
        Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1 And Rownum < 2;
      Exception
        When Others Then
          n_合作单位限数量模式 := 4;
      End;

      If n_合作单位限数量模式 = 0 Then
        v_Err_Msg := '当前号码(' || Nvl(号码_In, 0) || '未开放' || 合作单位_In || '的预约,不能继续。';
        Raise Err_Item;
      End If;
      If n_合作单位限数量模式 = 1 Or n_合作单位限数量模式 = 2 Then
        Select 数量
        Into n_Count
        From 临床出诊挂号控制记录
        Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1;
        If n_合作单位限数量模式 = 1 Then
          n_Count := Round(Nvl(n_限约数, n_限号数) * n_Count / 100);
        End If;
        Select Count(1)
        Into n_Exists
        From 病人挂号记录
        Where 记录状态 = 1 And 出诊记录id = 记录id_In And 合作单位 = 合作单位_In;
        If n_Exists >= n_Count Then
          v_Err_Msg := '当前号码(' || Nvl(号码_In, 0) || '已经超过' || 合作单位_In || '的预约限制,不能继续。';
          Raise Err_Item;
        End If;
      End If;
      --开放序号检查
      If n_合作单位限数量模式 = 3 Then
        For c_合作单位 In (Select 序号, 数量
                       From 临床出诊挂号控制记录
                       Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1 And 序号 = 号序_In) Loop
          If n_序号控制 = 1 Then
            Begin
              Select 1
              Into n_Count
              From 临床出诊序号控制
              Where 记录id = 记录id_In And 序号 = 号序_In And Nvl(挂号状态, 0) = 0;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 1 Then
              n_是否开放 := 1;
            Else
              v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '已经超过' || 合作单位_In || '的预约限制,不能继续。';
              Raise Err_Item;
            End If;
          Else
            Select Count(1)
            Into n_Count
            From 临床出诊序号控制
            Where 记录id = 记录id_In And 序号 = 号序_In And 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0;
            If n_Count >= c_合作单位.数量 Then
              v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '已经超过' || 合作单位_In || '的预约限制,不能继续。';
              Raise Err_Item;
            Else
              n_是否开放 := 1;
            End If;
          End If;
        End Loop;

        If Nvl(n_是否开放, 0) = 0 Then
          v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '未开放,不能继续。';
          Raise Err_Item;
        End If;
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

    If Nvl(购买病历_In, 0) = 1 Then
      Begin
        Select 收费细目id Into n_病历费id From 收费特定项目 Where 特定项目 = '病历费';
        v_收费项目ids := n_项目id || ',' || n_病历费id;
      Exception
        When Others Then
          v_Err_Msg := '不能确定病历费,挂号失败!';
          Raise Err_Item;
      End;
    Else
      v_收费项目ids := n_项目id;
    End If;

    For c_Item In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_项目id And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_病历费id And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 3 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, 1 As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And
                         d.主项id In (Select Column_Value From Table(f_Str2list(v_收费项目ids))) And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
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
           Decode(操作方式_In, 2, To_Char(号序_In), v_诊室), r_Pati.病人id, r_Pati.门诊号, r_Pati.付款方式, r_Pati.姓名, r_Pati.性别,
           r_Pati.年龄, r_Pati.费别, n_科室id, c_Item.类别, 号码_In, c_Item.项目id, c_Item.收入项目id, c_Item.收据费目, 1, c_Item.数次,
           c_Item.单价, n_应收金额, n_实收金额, Decode(操作方式_In, 2, Null, n_实收金额), n_结帐id, 0, n_开单部门id, v_操作员姓名,
           Decode(操作方式_In, 2, v_操作员姓名, Null), n_科室id, v_医生姓名, v_操作员编号, v_操作员姓名, 发生时间_In, d_登记时间, Null, 0, Null, Null,
           摘要_In, 预约方式_In, Decode(操作方式_In, 2, Null, n_组id));
      End If;
      n_行号 := n_行号 + 1;

    End Loop;

    If Round(Nvl(挂号金额合计_In, 0), 5) <> Round(Nvl(n_实收金额合计, 0), 5) Then
      v_Err_Msg := '本次挂号金额不正确,可能是因为医院调整了价格,请重新获取挂号收费项目的价格,不能继续。';
      Raise Err_Item;
    End If;

    If n_启用分时段 = 1 Then
      d_Date := To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(d_时段开始时间, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
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
      From 临床出诊序号控制
      Where 记录id = 记录id_In And 序号 = n_号序 And Nvl(挂号状态, 0) Not In (0, 5);
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_启用分时段 = 0 And Nvl(n_序号控制, 0) = 1 Then
        n_号序 := Null;
      End If;
      If n_启用分时段 = 1 And Nvl(n_序号控制, 0) = 1 Then
        v_Err_Msg := '当前序号已被使用，请重新选择一个序号！';
        Raise Err_Item;
      End If;
    End If;

    If n_启用分时段 = 0 And Nvl(n_序号控制, 0) = 1 And n_号序 Is Null And 锁定类型_In <> 2 Then
      Select Nvl(Min(序号), 0)
      Into n_号序
      From 临床出诊序号控制
      Where 记录id = 记录id_In And Nvl(挂号状态, 0) = 5 And 操作员姓名 = v_操作员姓名 And 工作站名称 = v_机器名;
      If n_号序 = 0 Then
        Select Nvl(Min(序号), 0) Into n_号序 From 临床出诊序号控制 Where 记录id = 记录id_In And Nvl(挂号状态, 0) = 0;
        If n_号序 = 0 Then
          Select Nvl(Max(序号), 0) + 1 Into n_号序 From 临床出诊序号控制 Where 记录id = 记录id_In;
        End If;
      End If;
    End If;

    If n_启用分时段 = 1 And 锁定类型_In <> 2 Then
      If 操作方式_In > 1 And Nvl(n_序号控制, 0) = 0 Then
        --规则:预约时段序号||预约数
        If Nvl(n_预约总数, 0) = 0 Then
          v_Temp := Nvl(n_限约数, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_预约总数, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_预约时段序号 || v_Temp;
          n_号序 := To_Number(v_Temp);
        Else
          n_号序 := n_预约总数 + 1;
        End If;
      End If;
    End If;

    If Nvl(n_序号控制, 0) = 1 Or (操作方式_In > 1 And n_启用分时段 = 1) Or 加入序号状态_In = 1 Then
      --锁定序号的处理
      Begin
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where 挂号状态 = 5 And 记录id = 记录id_In And 序号 = n_号序;
        n_序号锁定 := 1;
      Exception
        When Others Then
          v_序号操作员 := Null;
          v_序号机器名 := Null;
          n_序号锁定   := 0;
      End;
      If n_序号锁定 = 0 Then
        If n_启用分时段 = 1 And n_序号控制 = 0 Then
          Insert Into 临床出诊序号控制
            (记录id, 序号, 预约顺序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 类型, 名称, 操作员姓名, 备注)
            Select 记录id_In, n_预约时段序号, n_预约顺序号, d_时段开始时间, d_时段终止时间, 1, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 2, 2, 1),
                   1, 合作单位_In, v_操作员姓名, n_号序
            From Dual;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(操作方式_In, 2, 2, 1), 操作员姓名 = v_操作员姓名
          Where 记录id = 记录id_In And 序号 = n_号序;
        End If;
        If Sql%RowCount = 0 Then
          Begin
            If n_启用分时段 = 1 Then
              --分时段
              If n_序号控制 = 1 Then
                --序号控制
                Select Max(终止时间) Into d_终止时间 From 临床出诊序号控制 Where 记录id = 记录id_In;
                If Sysdate > d_终止时间 Then
                  d_终止时间 := Sysdate;
                End If;
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 类型, 名称, 操作员姓名)
                  Select 记录id_In, n_号序, d_终止时间, d_终止时间, 1, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 2, 2, 1), 1,
                         合作单位_In, v_操作员姓名
                  From Dual;
              Else
                --分时段,非序号控制
                Null;
              End If;
            Else
              --不分时段
              Insert Into 临床出诊序号控制
                (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 类型, 名称, 操作员姓名)
                Select 记录id_In, n_号序, 开始时间, 终止时间, 1, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 2, 2, 1), 1, 合作单位_In,
                       v_操作员姓名
                From 临床出诊序号控制
                Where 记录id = 记录id_In And 序号 = 1;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '序号' || n_号序 || '已被使用,请重新选择一个序号.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_操作员姓名 <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
          v_Err_Msg := '序号' || n_号序 || '已被机器' || v_机器名 || '锁定,请重新选择一个序号.';
          Raise Err_Item;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(操作方式_In, 2, 2, 1), 锁号时间 = Null
          Where 记录id = 记录id_In And 序号 = n_号序 And 挂号状态 = 5 And 操作员姓名 = v_操作员姓名 And 工作站名称 = v_机器名;
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
              (n_预交id, 4, 单据号_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, d_登记时间, v_操作员编号, v_操作员姓名,
               n_结算金额, n_结帐id, n_组id, n_结帐id, 4);
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          End If;
          n_结算合计 := Nvl(n_结算合计, 0) + Nvl(n_结算金额, 0);
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
        End Loop;
      End If;

      If Nvl(冲预交_In, 0) <> 0 Then
        Select Nvl(Sum(Nvl(预交余额, 0) - Nvl(费用余额, 0)), 0)
        Into n_病人余额
        From 病人余额
        Where 病人id = 病人id_In And Nvl(性质, 0) = 1 And Nvl(类型, 0) = 1;
        if n_病人余额 < 冲预交_In Then
          v_Err_Msg := '病人的当前预交余额为 ' || Ltrim(To_Char(n_病人余额, '9999999990.00')) || '，小于本次支付金额 ' ||
                       Ltrim(To_Char(冲预交_In, '9999999990.00')) || '，支付失败！';
          Raise Err_Item;
        End if;

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
          If r_Deposit.结帐id = 0 Then
            --第一次冲预交(填上结帐ID,金额为0)
            Update 病人预交记录 Set 冲预交 = 0, 结帐id = n_结帐id, 结算性质 = 4 Where ID = r_Deposit.原预交id;
          End If;
          --冲上次剩余额
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名,
             操作员编号, 冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行,
                   单位帐号, d_登记时间, v_操作员姓名, v_操作员编号, n_结算金额, n_结帐id, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, 4
            From 病人预交记录
            Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;

          --更新病人预交余额
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) - n_结算金额
          Where 病人id = 病人id_In And 性质 = 1 And 类型 = Nvl(1, 2)
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 预交余额, 性质, 类型) Values (病人id_In, -1 * n_结算金额, 1, 1);
            n_返回值 := -1 * n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;

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
        If Instr(结算方式_In, ',') = 0 Then
          --只传入一种结算方式的
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(结算方式_In, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             v_操作员编号, v_操作员姓名, n_结帐id, '挂号收费', n_组id, 卡类别id_In, Null, 支付卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4, v_结算号码);
        Else
          v_结算内容     := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
          n_Exists       := 0;
          v_结算方式记录 := '';
          While v_结算内容 Is Not Null Loop
            v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
            v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);

            v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            n_单笔金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));

            v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);

            v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            n_三方卡标志 := To_Number(v_当前结算);

            If Instr('|' || v_结算方式记录 || '|', '|' || Nvl(v_结算方式, v_现金) || '|') <> 0 Then
              v_Err_Msg := '使用了重复的结算方式,请检查!';
              Raise Err_Item;
            Else
              v_结算方式记录 := v_结算方式记录 || '|' || Nvl(v_结算方式, v_现金);
            End If;

            If n_三方卡标志 = 0 Then
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              Insert Into 病人预交记录
                (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号,
                 交易说明, 合作单位, 结算性质, 结算号码)
              Values
                (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_单笔金额, 0), 登记时间_In,
                 v_操作员编号, v_操作员姓名, n_结帐id, '挂号收费', n_组id, Null, Null, Null, Null, Null, 合作单位_In, 4, v_结算号码);
            Else
              If n_Exists = 1 Then
                v_Err_Msg := '目前挂号仅支持一种三方结算方式,不能继续操作！';
                Raise Err_Item;
              End If;
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              Insert Into 病人预交记录
                (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号,
                 交易说明, 合作单位, 结算性质, 结算号码)
              Values
                (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_单笔金额, 0), 登记时间_In,
                 v_操作员编号, v_操作员姓名, n_结帐id, '挂号收费', n_组id, 卡类别id_In, Null, 支付卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4, v_结算号码);
              n_Exists := 1;
            End If;

            v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
          End Loop;
        End If;
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
          Where 收款员 = v_操作员姓名 And 结算方式 = v_缴款.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
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
        操作员编号 = v_操作员编号, 预约 = Decode(操作方式_In, 1, 0, 1),
        接收人 = Decode(锁定类型_In, 1, Null, Decode(操作方式_In, 2, Null, v_操作员姓名)),
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
        预约操作员编号 = Decode(操作方式_In, 1, Nvl(预约操作员编号, Null), Nvl(预约操作员编号, v_操作员编号)), 出诊记录id = 记录id_In
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
         操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式, 预约操作员, 预约操作员编号, 出诊记录id)
      Values
        (n_挂号id, 单据号_In, Decode(操作方式_In, 2, 2, 1), Decode(锁定类型_In, 1, 0, 1), r_Pati.病人id, r_Pati.门诊号, r_Pati.姓名,
         r_Pati.性别, r_Pati.年龄, 号码_In, 0, v_诊室, Null, n_科室id, v_医生姓名, 0, Null, d_登记时间, 发生时间_In,
         Case When(Nvl(操作方式_In, 0)) = 1 Then Null Else 发生时间_In End, v_操作员编号, v_操作员姓名, 0, n_号序, Decode(操作方式_In, 1, 0, 1),
         Decode(操作方式_In, 2, Null, v_操作员姓名), Decode(操作方式_In, 2, To_Date(Null), d_登记时间), 交易流水号_In, 交易说明_In, 合作单位_In,
         v_付款方式, Decode(操作方式_In, 1, Null, v_操作员姓名), Decode(操作方式_In, 1, Null, v_操作员编号), 记录id_In);
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
            n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(操作方式_In, 0) > 1 And n_分时点显示 = 1 And n_启用分时段 = 1 Then
              n_分时点显示 := 1;
            Else
              n_分时点显示 := Null;
            End If;
            --产生队列
            --.按”执行部门” 的方式生成队列
            v_队列名称 := n_科室id;
            v_排队号码 := Zlgetnextqueue(n_科室id, n_挂号id, 号码_In || '|' || 号序_In);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号码_In, 号序_In, d_Date);
            --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,v_排队标记,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
            Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, n_科室id, v_排队号码, Null, r_Pati.姓名, r_Pati.病人id, v_诊室, v_医生姓名, d_排队时间,
                             预约方式_In, n_分时点显示, v_排队序号);
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
      Zl_病人挂号汇总_Update(v_医生姓名, n_医生id, n_项目id, n_科室id, 发生时间_In, n_预约, 号码_In, 0, 记录id_In);
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
    When Err_Special Then
      Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_出诊记录id := 出诊记录id_In;
  v_Para       := zl_GetSysParameter(256);
  n_挂号模式   := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  d_发生时间 := 发生时间_In;
  If d_发生时间 Is Null Then
    d_发生时间 := Sysdate;
  End If;

  If Sysdate - 10 > Nvl(d_启用时间, Sysdate - 30) Then
    If n_挂号模式 = 1 And Nvl(发生时间_In, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_出诊记录id Is Null Then
      v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
      Raise Err_Item;
    End If;
  Else
    If n_挂号模式 = 1 And Nvl(发生时间_In, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_出诊记录id Is Null Then
      Begin
        Select a.Id
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = 号码_In And Nvl(发生时间_In, Sysdate) Between a.开始时间 And a.终止时间;
      Exception
        When Others Then
          v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If n_出诊记录id Is Not Null Then
    --出诊表排班模式
    Zl_三方机构挂号_出诊_Insert(n_出诊记录id, 操作方式_In, 病人id_In, 号码_In, 号序_In, 单据号_In, 票据号_In, 结算方式_In, 摘要_In, 发生时间_In, 登记时间_In,
                        合作单位_In, 挂号金额合计_In, 领用id_In, 收费票据_In, 交易流水号_In, 交易说明_In, 预约方式_In, 预交id_In, 卡类别id_In, 加入序号状态_In,
                        是否自助设备_In, 结帐id_In, 锁定类型_In, 保险结算_In, 冲预交_In, 支付卡号_In, 退号重用_In, 费别_In, 机器名_In, 更新年龄_In, 购买病历_In);
  Else
    v_Temp := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_启用时间 := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      Exception
        When Others Then
          Null;
      End;
      If 发生时间_In > d_启用时间 Then
        v_Err_Msg := '当前挂号的发生时间' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '已经启用了出诊表排班模式,不能再使用计划排班模式挂号!';
        Raise Err_Item;
      End If;
    End If;
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
    v_Temp           := Nvl(zl_GetSysParameter('病人同科限挂N个号', 1111), '0|0') || '|';
    n_同科限号数     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_同科限约数     := To_Number(Nvl(zl_GetSysParameter('病人同科限约N个号', 1111), '0'));
    n_病人预约科室数 := To_Number(Nvl(zl_GetSysParameter('病人预约科室数', 1111), '0'));
    n_病人挂号科室数 := To_Number(Nvl(zl_GetSysParameter('病人挂号科室限制', 1111), '0'));
    n_专家号挂号限制 := To_Number(Nvl(zl_GetSysParameter('专家号挂号限制'), '0'));
    n_专家号预约限制 := To_Number(Nvl(zl_GetSysParameter('专家号预约限制'), '0'));

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
    n_开单部门id := To_Number(Zl_操作员(0, v_Temp));
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
    n_组id       := Zl_Get组id(v_操作员姓名);

    --支付宝并发提交检查
    Select Nvl(Max(1), 0)
    Into n_Exists
    From 病人挂号记录
    Where 病人id = 病人id_In And 号别 = 号码_In And 号序 = 号序_In And 操作员姓名 = v_操作员姓名 And Nvl(n_出诊记录id, 0) = Nvl(出诊记录id, 0) And
          登记时间 > Sysdate - 0.01 And 记录状态 = 1 And 发生时间 = 发生时间_In;
    If n_Exists = 1 Then
      v_Err_Msg := '病人已经挂号,不能重复挂相同的号！';
      Raise Err_Special;
    End If;

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

    Select Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                   '周日')
    Into v_星期
    From Dual;
    Begin
      If r_安排.计划id Is Null Then
        Select Max(1) Into n_启用分时段 From 挂号安排时段 Where 安排id = r_安排.Id And 星期 = v_星期 And Rownum < 2;
        Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
        Into v_时间段
        From 挂号安排
        Where ID = r_安排.Id;
      Else
        Select Max(1)
        Into n_启用分时段
        From 挂号计划时段
        Where 计划id = r_安排.计划id And 星期 = v_星期 And Rownum < 2;
        Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
        Into v_时间段
        From 挂号安排计划
        Where ID = r_安排.计划id;
      End If;
    Exception
      When Others Then
        n_启用分时段 := 0;
    End;

    If v_时间段 Is Not Null And d_启用时间 Is Not Null Then
      --检查是否跨模式挂号安排
      Select To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_检查开始时间, d_检查结束时间
      From 时间段
      Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
      If d_检查开始时间 > d_检查结束时间 Then
        d_检查结束时间 := d_检查结束时间 + 1;
      End If;
      If d_检查结束时间 > d_启用时间 Then
        --获取出诊记录id
        Begin
          Select a.Id
          Into n_出诊记录id
          From 临床出诊记录 A, 临床出诊号源 B
          Where a.号源id = b.Id And b.号码 = 号码_In And 上班时段 = v_时间段 And 发生时间_In Between 开始时间 And 终止时间;
        Exception
          When Others Then
            n_出诊记录id := Null;
        End;
      End If;
    End If;

    --对参数控制进行检查
    --仅在预约不扣款时进行检查
    If 操作方式_In = 2 Then
      If Nvl(n_同科限约数, 0) <> 0 Or Nvl(n_病人预约科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> r_安排.科室id) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
          Raise Err_Item;
        End If;

        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = r_安排.科室id;
        If n_Count >= Nvl(n_同科限约数, 0) And Nvl(n_同科限约数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室预约了' || n_Count || '次,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_专家号预约限制, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = r_安排.号码;
        If n_Count >= Nvl(n_专家号预约限制, 0) And Nvl(n_专家号预约限制, 0) > 0 Then
          v_Err_Msg := '该病人已经超过本号预约限制,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_同科限号数, 0) <> 0 Or Nvl(n_病人挂号科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> r_安排.科室id) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人挂号科室数, 0) And Nvl(n_病人挂号科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能挂号[' || Nvl(n_病人挂号科室数, 0) || ']个科室,不能再挂号！';
          Raise Err_Item;
        End If;

        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = r_安排.科室id;
        If n_Count >= Nvl(n_同科限号数, 0) And Nvl(n_同科限号数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室挂号了' || n_Count || '次,不能再挂号！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_专家号挂号限制, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = r_安排.号码;
        If n_Count >= Nvl(n_专家号挂号限制, 0) And Nvl(n_专家号挂号限制, 0) > 0 Then
          v_Err_Msg := '该病人已经超过本号挂号限制,不能再挂号！';
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
            If r_安排.计划id Is Null Then
              Select Count(*), Max(开始时间)
              Into n_Count, d_时段开始时间
              From 挂号安排时段
              Where 安排id = r_安排.Id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0);
            Else
              Select Count(*), Max(开始时间)
              Into n_Count, d_时段开始时间
              From 挂号计划时段
              Where 计划id = r_安排.计划id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0);
            End If;
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
            If r_安排.计划id Is Null Then
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
            Else
              For v_时段 In (Select To_Date(v_Temp || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                  To_Date(To_Char(Sysdate + Decode(Sign(开始时间 - 结束时间), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间, 限制数量, 是否预约
                           From 挂号计划时段
                           Where 计划id = r_安排.计划id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0)) Loop
                If Sysdate > v_时段.结束时间 Then
                  v_Err_Msg := '号别为' || 号码_In || '及号序为' || Nvl(号序_In, 0) || '的安排,已经超过时点,不能再' || v_Temp || '！';
                  Raise Err_Item;
                End If;
              End Loop;
            End If;
          End If;
        Elsif 操作方式_In > 1 Then
          --未启用序号的,需要检查预约的情况
          n_Count := 0;
          If r_安排.计划id Is Null Then
            For v_时段 In (Select 序号, 开始时间, 结束时间, 限制数量, 是否预约
                         From 挂号安排时段
                         Where 安排id = r_安排.Id And 星期 = v_星期 And
                               (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
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
          Else
            For v_时段 In (Select 序号, 开始时间, 结束时间, 限制数量, 是否预约
                         From 挂号计划时段
                         Where 计划id = r_安排.计划id And 星期 = v_星期 And
                               (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
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
          End If;

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
                限制项目 = Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                              '7', '周六', Null) And 数量 <> 0 And 序号 = 0 And Rownum < 2;
        Else
          Select 0
          Into n_序号
          From 合作单位安排控制
          Where 合作单位 = 合作单位_In And 安排id = n_Tmp安排id And
                限制项目 = Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                              '7', '周六', Null) And 数量 <> 0 And 序号 = 0 And Rownum < 2;
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

    If Nvl(购买病历_In, 0) = 1 Then
      Begin
        Select 收费细目id Into n_病历费id From 收费特定项目 Where 特定项目 = '病历费';
        v_收费项目ids := r_安排.项目id || ',' || n_病历费id;
      Exception
        When Others Then
          v_Err_Msg := '不能确定病历费,挂号失败!';
          Raise Err_Item;
      End;
    Else
      v_收费项目ids := r_安排.项目id;
    End If;

    For c_Item In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = r_安排.项目id And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_病历费id And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 3 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, 1 As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And
                         d.主项id In (Select Column_Value From Table(f_Str2list(v_收费项目ids))) And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
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
           Decode(操作方式_In, 2, To_Char(号序_In), v_诊室), r_Pati.病人id, r_Pati.门诊号, r_Pati.付款方式, r_Pati.姓名, r_Pati.性别,
           r_Pati.年龄, r_Pati.费别, r_安排.科室id, c_Item.类别, 号码_In, c_Item.项目id, c_Item.收入项目id, c_Item.收据费目, 1, c_Item.数次,
           c_Item.单价, n_应收金额, n_实收金额, Decode(操作方式_In, 2, Null, n_实收金额), n_结帐id, 0, n_开单部门id, v_操作员姓名,
           Decode(操作方式_In, 2, v_操作员姓名, Null), r_安排.科室id, r_安排.医生姓名, v_操作员编号, v_操作员姓名, 发生时间_In, d_登记时间, Null, 0, Null,
           Null, 摘要_In, 预约方式_In, Decode(操作方式_In, 2, Null, n_组id));
      End If;
      n_行号 := n_行号 + 1;

    End Loop;

    If Round(Nvl(挂号金额合计_In, 0), 5) <> Round(Nvl(n_实收金额合计, 0), 5) Then
      v_Err_Msg := '本次挂号金额不正确,可能是因为医院调整了价格,请重新获取挂号收费项目的价格,不能继续。';
      Raise Err_Item;
    End If;

    If n_启用分时段 = 1 Then
      d_Date := To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(d_时段开始时间, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
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

    If n_出诊记录id Is Not Null Then
      Update 临床出诊序号控制
      Set 挂号状态 = Decode(操作方式_In, 2, 2, 1), 操作员姓名 = v_操作员姓名
      Where 记录id = n_出诊记录id And 序号 = n_序号;
      If 操作方式_In = 2 Then
        Update 临床出诊记录 Set 已约数 = 已约数 + 1 Where ID = n_出诊记录id;
      Else
        If 操作方式_In <> 1 Then
          Update 临床出诊记录
          Set 已约数 = 已约数 + 1, 已挂数 = 已挂数 + 1, 其中已接收 = 其中已接收 + 1
          Where ID = n_出诊记录id;
        Else
          Update 临床出诊记录 Set 已挂数 = 已挂数 + 1 Where ID = n_出诊记录id;
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
              (n_预交id, 4, 单据号_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, d_登记时间, v_操作员编号, v_操作员姓名,
               n_结算金额, n_结帐id, n_组id, n_结帐id, 4);
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          End If;
          n_结算合计 := Nvl(n_结算合计, 0) + Nvl(n_结算金额, 0);
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
        End Loop;
      End If;

      If Nvl(冲预交_In, 0) <> 0 Then
        Select Nvl(Sum(Nvl(预交余额, 0) - Nvl(费用余额, 0)), 0)
        Into n_病人余额
        From 病人余额
        Where 病人id = 病人id_In And Nvl(性质, 0) = 1 And Nvl(类型, 0) = 1;
        if n_病人余额 < 冲预交_In Then
          v_Err_Msg := '病人的当前预交余额为 ' || Ltrim(To_Char(n_病人余额, '9999999990.00')) || '，小于本次支付金额 ' ||
                       Ltrim(To_Char(冲预交_In, '9999999990.00')) || '，支付失败！';
          Raise Err_Item;
        End if;
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
          If r_Deposit.结帐id = 0 Then
            --第一次冲预交(填上结帐ID,金额为0)
            Update 病人预交记录 Set 冲预交 = 0, 结帐id = n_结帐id, 结算性质 = 4 Where ID = r_Deposit.原预交id;
          End If;
          --冲上次剩余额
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名,
             操作员编号, 冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行,
                   单位帐号, d_登记时间, v_操作员姓名, v_操作员编号, n_结算金额, n_结帐id, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, 4
            From 病人预交记录
            Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;

          --更新病人预交余额
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) - n_结算金额
          Where 病人id = 病人id_In And 性质 = 1 And 类型 = Nvl(1, 2)
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 预交余额, 性质, 类型) Values (病人id_In, -1 * n_结算金额, 1, 1);
            n_返回值 := -1 * n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;

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
          (n_预交id, 4, 1, 单据号_In, r_Pati.病人id, 结算方式_In, Nvl(n_结算金额, 0), d_登记时间, v_操作员编号, v_操作员姓名, n_结帐id,
           合作单位_In || '缴款', n_组id, 交易流水号_In, 交易说明_In, n_结帐id, 合作单位_In, 卡类别id_In, 支付卡号_In, 4);
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
        操作员编号 = v_操作员编号, 预约 = Decode(操作方式_In, 1, 0, 1),
        接收人 = Decode(锁定类型_In, 1, Null, Decode(操作方式_In, 2, Null, v_操作员姓名)),
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
         Decode(操作方式_In, 2, Null, v_操作员姓名), Decode(操作方式_In, 2, To_Date(Null), d_登记时间), 交易流水号_In, 交易说明_In, 合作单位_In,
         v_付款方式, Decode(操作方式_In, 1, Null, v_操作员姓名), Decode(操作方式_In, 1, Null, v_操作员编号));
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
            n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(操作方式_In, 0) > 1 And n_分时点显示 = 1 And n_启用分时段 = 1 Then
              n_分时点显示 := 1;
            Else
              n_分时点显示 := Null;
            End If;
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
                             d_排队时间, 预约方式_In, n_分时点显示, v_排队序号);
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
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方机构挂号_Insert;
/

--136111:冉俊明,2019-01-08,修正挂号合作单位与预约方式存在相同名称时临床出诊安排报错的问题
Create Or Replace Procedure Zl_临床出诊诊室_Update
(
  Id_In       临床出诊限制.Id%Type,
  分诊方式_In 临床出诊限制.分诊方式%Type := Null,
  诊室_In     Varchar2 := Null,
  出诊记录_In Number := 0
) As
  --功能：更新临床出诊诊室 
  --参数： 
  --     诊室_In:诊室1,诊室2,... 
  --     出诊记录_In:是否是对出诊记录进行删除 
  n_Count  Number;
  n_变动id 临床出诊变动记录.Id%Type;
  v_诊室   临床出诊变动记录.现门诊诊室%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  If Nvl(出诊记录_In, 0) = 0 Then
    Update 临床出诊限制 Set 分诊方式 = 分诊方式_In Where ID = Id_In;
  
    Delete From 临床出诊诊室 Where 限制id = Id_In;
    --出诊诊室 
    If 诊室_In Is Not Null Then
    
      Insert Into 临床出诊诊室
        (限制id, 诊室id)
        Select Id_In, Column_Value From Table(f_Str2list(诊室_In, ','));
    
      If Nvl(分诊方式_In, 0) = 1 Then
        Update 临床出诊限制 Set 诊室id = To_Number(诊室_In) Where ID = Id_In;
      End If;
    End If;
    Return;
  End If;

  --临床出诊变动信息 
  Select Count(1)
  Into n_Count
  From 临床出诊表 A, 临床出诊安排 B, 临床出诊记录 C
  Where a.Id = b.出诊id And b.Id = c.安排id And a.发布人 Is Not Null And c.Id = Id_In;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '出诊记录不存在！';
    Raise Err_Item;
  End If;

  Select 临床出诊变动记录_Id.Nextval Into n_变动id From Dual;
  Insert Into 临床出诊变动记录
    (ID, 记录id, 变动类型, 原分诊方式, 原诊室id, 原门诊诊室, 现分诊方式, 操作员姓名, 登记时间)
    Select n_变动id, a.Id, 3, a.分诊方式, a.诊室id, b.名称, 分诊方式_In, zl_UserName, Sysdate
    From 临床出诊记录 A, 门诊诊室 B
    Where a.诊室id = b.Id(+) And a.Id = Id_In;

  Insert Into 临床出诊变动明细
    (变动id, 变动性质, 序号, 诊室id, 门诊诊室)
    Select n_变动id, 1, 序号, 诊室id, 名称
    From (Select Rownum As 序号, a.诊室id, b.名称
           From 临床出诊诊室记录 A, 门诊诊室 B
           Where a.诊室id = b.Id(+) And a.记录id = Id_In);

  --保存原始临床出诊记录 
  Select Count(1) Into n_Count From 临床出诊记录 Where 相关id = Id_In;
  If Nvl(n_Count, 0) = 0 Then
    For c_记录 In (Select ID, 安排id, To_Date('1900-01-01', 'yyyy-mm-dd') As 出诊日期, 登记人, 登记时间, 是否发布
                 From 临床出诊记录
                 Where ID = Id_In) Loop
      Zl_临床出诊记录_Copy(c_记录.Id, c_记录.安排id, c_记录.出诊日期, c_记录.登记人, c_记录.登记时间, c_记录.是否发布, c_记录.Id);
    End Loop;
  End If;

  Update 临床出诊记录 Set 分诊方式 = 分诊方式_In Where ID = Id_In;
  Delete From 临床出诊诊室记录 Where 记录id = Id_In;

  --临床出诊变动后信息 
  If 诊室_In Is Not Null Then
    Insert Into 临床出诊诊室记录
      (记录id, 诊室id)
      Select Id_In, Column_Value From Table(f_Str2list(诊室_In, ','));
  
    Insert Into 临床出诊变动明细
      (变动id, 变动性质, 序号, 诊室id, 门诊诊室)
      Select n_变动id, 2, Rownum, a.Id, a.名称
      From 门诊诊室 A, (Select Column_Value As 诊室id From Table(f_Str2list(诊室_In, ','))) B
      Where a.Id = b.诊室id;
  
    If Nvl(分诊方式_In, 0) = 1 Then
      Update 临床出诊记录 Set 诊室id = To_Number(诊室_In) Where ID = Id_In;
    
      Update 临床出诊变动记录
      Set 现诊室id = To_Number(诊室_In),
          现门诊诊室 =
           (Select 名称 From 门诊诊室 Where ID = To_Number(诊室_In))
      Where ID = n_变动id
      Returning 现门诊诊室 Into v_诊室;
      --病人挂号记录 
      Update 病人挂号记录 Set 诊室 = v_诊室 Where 出诊记录id = Id_In;
      --门诊费用记录 
      Update 门诊费用记录
      Set 发药窗口 = v_诊室
      Where 记录性质 = 4 And NO In (Select NO From 病人挂号记录 Where 出诊记录id = Id_In);
    End If;
  
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊诊室_Update;
/

--115085:胡俊勇,2018-01-07,医嘱单打印参数拆分
Create Or Replace Procedure Zl_病人医嘱打印_Insert
(
  病人id_In 病人医嘱记录.病人id%Type,
  主页id_In 病人医嘱记录.主页id%Type,
  婴儿_In   病人医嘱记录.婴儿%Type,
  期效_In   病人医嘱记录.医嘱期效%Type,
  行数_In   Number
  --功能：将病人没有打印过的医嘱插入 病人医嘱打印
  --参数：行数_In：报表医嘱单一页可以打多少行
  --      行数_In医嘱单报表的行数，通常是28行。
) Is
  n_序号     病人医嘱记录.序号%Type;
  n_医嘱id   病人医嘱记录.Id%Type;
  n_重整标记 Number;
  v_Max_Date Date;
  d_重整     Date;
  d_Pdate    Date;
  n_打重开     Number;
  n_转科       Number;
  n_页号       Number;
  n_行号       Number;
  n_位置       Number;
  n_打印模式   Number;
  n_打给药方式 Number;
  n_Lzzkhy     Number; --临嘱单  
  n_重整换页   Number;
  n_转科换页   Number; --长期医嘱单
  n_术后换页   Number;
  b_重整换页   Boolean;
  n_Cnt        Number;
  v_Tmp        Varchar2(200);

  --c_Advice 取出待打印的医嘱，在打印临嘱时转科医嘱都会读取出来，后面要判断是不是要生成打印记录
  Cursor c_Advice Is
    Select 医嘱id, 顺序, 打印标记, 特殊医嘱, 换页
    From (With Printtable As (Select a.Id As 医嘱id, a.序号 As 顺序, 0 As 打印标记, Null As 特殊医嘱,
                                     Decode(a.诊疗类别, 'Z', Decode(b.操作类型, '3', 3, '4', 4, 0), 0) As 换页, a.诊疗项目id, a.相关id,
                                     b.操作类型, a.诊疗类别
                              From 病人医嘱记录 A, 诊疗项目目录 B
                              Where a.病人id = 病人id_In And a.主页id = 主页id_In And Nvl(a.婴儿, 0) = 婴儿_In And a.诊疗项目id = b.Id(+) And
                                    (期效_In = 0 And (a.医嘱期效 = 0 Or n_位置 In (-1, 0, 2) And a.医嘱期效 = 1 And a.诊疗类别 = 'Z' And
                                    b.操作类型 In ('5', '3', '11')) Or
                                    期效_In = 1 And a.医嘱期效 = 1 And
                                    Not (n_位置 = 0 And Nvl(a.诊疗类别, 'X') = 'Z' And Nvl(b.操作类型, 'X') In ('5', '3', '11')) Or
                                    期效_In = 1 And a.医嘱期效 = 1 And n_位置 = 0 And a.诊疗类别 = 'Z' And b.操作类型 = '3') And
                                    a.医嘱状态 Not In (-1, 2) And (n_打印模式 = 1 And a.医嘱状态 = 1 Or a.医嘱状态 <> 1) And
                                    Nvl(a.屏蔽打印, 0) = 0 And a.序号 > n_序号 And a.病人来源 = 2)
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, 诊疗项目目录 I, Printtable P
           Where l.Id = p.医嘱id And l.诊疗项目id = i.Id And
                 (l.诊疗类别 Not In ('5', '6', '7', 'E') Or l.诊疗类别 = 'E' And Nvl(i.操作类型, '0') Not In ('2', '3') Or
                 i.Id Is Null) And l.相关id Is Null
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, Printtable P
           Where l.Id = p.医嘱id And l.诊疗类别 In ('5', '6')
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From Printtable P
           Where p.诊疗类别 = 'E' And p.操作类型 = '2' And p.相关id Is Null And
                 (n_打给药方式 = 1 Or n_打给药方式 = 2 And Exists
                  (Select 1 From 病人医嘱记录 L Where l.相关id = p.医嘱id Having Count(1) > 1))
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From Printtable P
           Where p.诊疗项目id Is Null
           Order By 顺序);


  Cursor c_Advice_Redo Is
    Select 医嘱id, 顺序, 打印标记, 特殊医嘱, 换页
    From (With Printtable As (Select a.Id As 医嘱id, a.序号 As 顺序, 0 As 打印标记, Null As 特殊医嘱,
                                     Decode(a.诊疗类别, 'Z', Decode(b.操作类型, '3', 3, '4', 4, 0), 0) As 换页, a.诊疗项目id, a.相关id,
                                     b.操作类型, a.诊疗类别
                              From 病人医嘱记录 A, 诊疗项目目录 B
                              Where a.病人id = 病人id_In And a.主页id = 主页id_In And Nvl(a.婴儿, 0) = 婴儿_In And a.诊疗项目id = b.Id(+) And
                                    (期效_In = 0 And (a.医嘱期效 = 0 Or n_位置 In (-1, 0, 2) And a.医嘱期效 = 1 And a.诊疗类别 = 'Z' And
                                    b.操作类型 In ('5', '3', '11')) Or
                                    期效_In = 1 And a.医嘱期效 = 1 And
                                    Not (n_位置 = 0 And Nvl(a.诊疗类别, 'X') = 'Z' And Nvl(b.操作类型, 'X') In ('5', '3', '11'))) And
                                    a.医嘱状态 Not In (-1, 2) And (n_打印模式 = 1 And a.医嘱状态 = 1 Or a.医嘱状态 <> 1) And
                                    Nvl(a.屏蔽打印, 0) = 0 And a.序号 > n_序号 And Exists
                               (Select 1 From 病人医嘱状态 C Where a.Id = c.医嘱id And c.操作时间 >= v_Max_Date) And a.病人来源 = 2)
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, 诊疗项目目录 I, Printtable P
           Where l.Id = p.医嘱id And l.诊疗项目id = i.Id And
                 (l.诊疗类别 Not In ('5', '6', '7', 'E') Or l.诊疗类别 = 'E' And Nvl(i.操作类型, '0') Not In ('2', '3') Or
                 i.Id Is Null) And l.相关id Is Null
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, Printtable P
           Where l.Id = p.医嘱id And l.诊疗类别 In ('5', '6')
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From Printtable P
           Where p.诊疗类别 = 'E' And p.操作类型 = '2' And p.相关id Is Null And
                 (n_打给药方式 = 1 Or n_打给药方式 = 2 And Exists
                  (Select 1 From 病人医嘱记录 L Where l.相关id = p.医嘱id Having Count(1) > 1))
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From Printtable P
           Where p.诊疗项目id Is Null
           Order By 顺序);


  --获取下一个或用的行号和页号
  Function Getnextpos
  (
    v_页号 病人医嘱打印.页号%Type,
    v_行号 病人医嘱打印.行号%Type,
    v_行数 Number
  ) Return Varchar2 Is
    n_p Number;
    n_r Number;
  Begin
    If v_行号 = 0 Then
      n_p := 1;
      n_r := 1;
    Elsif v_行号 = v_行数 Then
      n_p := v_页号 + 1;
      n_r := 1;
    Else
      n_p := v_页号;
      n_r := v_行号 + 1;
    End If;
    Return(n_p || ',' || n_r);
  End;

Begin
  n_位置       := Zl_To_Number(Nvl(zl_GetSysParameter('转科和出院打印', 1254), 0));
  n_打印模式   := Zl_To_Number(Nvl(zl_GetSysParameter('医嘱单打印模式', 1253), 0));
  n_打给药方式 := Zl_To_Number(Nvl(zl_GetSysParameter('药品用法单独打印一行', 1254), 0));
  n_Lzzkhy     := Zl_To_Number(Nvl(zl_GetSysParameter('临嘱单转科换页', 1254), 0));
  n_重整换页 := Zl_To_Number(Nvl(zl_GetSysParameter('长嘱单重整换页', 1254), 0));
  n_转科换页 := Zl_To_Number(Nvl(zl_GetSysParameter('长嘱单转科换页', 1254), 0)); 
  n_术后换页 := Zl_To_Number(Nvl(zl_GetSysParameter('长嘱单术后换页', 1254), 0)); 
  n_打重开 := Zl_To_Number(Nvl(zl_GetSysParameter('转科换页后在首行打印重开医嘱', 1254), 0));

  --判断是不是重整后打印医嘱
  If 期效_In = 1 Then
    d_重整 := To_Date('1900-01-01', 'YYYY-MM-DD');
  Else
    Select 医嘱重整时间 Into d_重整 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
    If d_重整 Is Null Then
      d_重整 := To_Date('1900-01-01', 'YYYY-MM-DD');
    End If;
  End If;
  v_Max_Date := d_重整;
  Begin
    Select 医嘱id, 打印时间, 页号, 行号
    Into n_医嘱id, d_Pdate, n_页号, n_行号
    From (Select 医嘱id, 打印时间, 页号, 行号
           From 病人医嘱打印
           Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(婴儿, 0) = 婴儿_In And 期效 = 期效_In And 医嘱id Is Not Null
           Order By 页号 Desc, 行号 Desc)
    Where Rownum < 2;
  
    Select Nvl(Max(序号), 0)
    Into n_序号
    From 病人医嘱记录
    Where ID = (Select Nvl(a.相关id, a.Id) From 病人医嘱记录 A Where a.Id = n_医嘱id);
  
    If 期效_In = 0 Then
      If d_Pdate Is Not Null Then
        If d_Pdate < d_重整 And d_重整 <> To_Date('1900-01-01', 'YYYY-MM-DD') Then
          n_重整标记 := 1;
          n_序号     := 0;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      n_页号 := 0;
      n_行号 := 0;
      n_序号 := 0;
  End;

  If n_医嘱id Is Not Null Then
    Select Max(b.操作类型)
    Into v_Tmp
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id(+) And a.Id = n_医嘱id And a.诊疗类别 = 'Z';
  End If;
  If v_Tmp = '3' Then
    n_Cnt := 3;
  Elsif v_Tmp = '4' Then
    n_Cnt := 4;
  End If;

  v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
  n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
  n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);

  If n_Cnt = 3 And n_Lzzkhy = 1 And 期效_In = 1 Then
    --临时医嘱转科换页
    If n_行号 <> 1 Then
      n_行号 := 1;
      n_页号 := n_页号 + 1;
    End If;
  Elsif 期效_In = 0 Then
    b_重整换页 := False;
    --重整，术后，转科重开，这些只针对于长期医嘱单
    --重整标记
    If n_重整标记 = 1 Then
      If n_重整换页 = 1 Then
        --重整换页
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
      End If;
      --重整标记
      Insert Into 病人医嘱打印
        (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
      Values
        (-1 * Null, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, 0, Null);
      v_Tmp      := Getnextpos(n_页号, n_行号, 行数_In);
      n_页号     := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
      n_行号     := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      b_重整换页 := True;
    End If;
  
    If n_Cnt = 3 Then
      --转科
      If n_转科换页 = 1 Then
        --如果重整已经换了页就不用换了。
        If Not b_重整换页 Then
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
        End If;
        If n_打重开 = 1 Then
          Insert Into 病人医嘱打印
            (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
          Values
            (-1 * Null, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, 0, 1);
          v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
          n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
          n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
        End If;
      End If;
    Elsif n_Cnt = 4 Then
      --术后
      If n_术后换页 = 1 Then
        If Not b_重整换页 Then
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
        End If;
      End If;
    End If;
  End If;
  n_转科 := 0;

  --最近次重整后,需要打印的医嘱，考虑换页打印情况转科术后
  ---r_Print.换页 对特殊医嘱标记，4－术后，3－转科
  If v_Max_Date = To_Date('1900-01-01', 'YYYY-MM-DD') Then
    For r_Print In c_Advice Loop
      ----换页或者打医嘱重开字样
    
      --长期医嘱单
      If n_转科换页 = 1 And n_转科 = 1 And 期效_In = 0 Then
        If n_打重开 = 1 Then
          --打重开字样
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
          Insert Into 病人医嘱打印
            (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
          Values
            (-1 * Null, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, 0, 1);
          v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
          n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
          n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
        Else
          --只是单纯换一页
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
        End If;
        n_转科 := 0;
      End If;
    
      --临时医嘱单
      If 期效_In = 1 And n_转科 = 1 And n_Lzzkhy = 1 Then
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
        n_转科 := 0;
      End If;
    
      If r_Print.换页 = 4 And n_术后换页 = 1 Then
        --术后医嘱换页
        --如果行号为1说明已经是新的一页的第一行,否则换页
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
      End If;
    
      If 期效_In = 0 Or 期效_In = 1 And (n_位置 = 2 Or n_位置 = 1 Or r_Print.换页 <> 3) Then
        Insert Into 病人医嘱打印
          (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
        Values
          (r_Print.医嘱id, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, r_Print.打印标记, r_Print.特殊医嘱);
        v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
        n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
        n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End If;
    
      --启用了转科换页打重开字样，则插入一重开医嘱标记，此处一定换页，因为转科换页前要先打出转科医嘱，
      --这里不插入数据，只进行标记，再下一次循时才插入。如果转科医嘱是最后一条是不用打印新开字样的。
      If r_Print.换页 = 3 Then
        n_转科 := 1;
      End If;
    End Loop;
  Else
    For r_Print In c_Advice_Redo Loop
      ----换页或者打医嘱重开字样
      If n_转科换页 = 1 And n_转科 = 1 And 期效_In = 0 Then
        If n_打重开 = 1 Then
          --打重开字样
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
          Insert Into 病人医嘱打印
            (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
          Values
            (-1 * Null, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, 0, 1);
          v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
          n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
          n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
        Else
          --只是单纯换一页
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
        End If;
        n_转科 := 0;
      End If;
    
      --临时医嘱单
      If 期效_In = 1 And n_转科 = 1 And n_Lzzkhy = 1 Then
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
        n_转科 := 0;
      End If;
    
      If r_Print.换页 = 4 And n_术后换页 = 1 Then
        --术后医嘱换页
        --如果行号为1说明已经是新的一页的第一行,否则换页
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
      End If;
      Insert Into 病人医嘱打印
        (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
      Values
        (r_Print.医嘱id, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, r_Print.打印标记, r_Print.特殊医嘱);
      v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
      n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
      n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      --启用了转科换页打重开字样，则插入一重开医嘱标记，此处一定换页，因为转科换页前要先打出转科医嘱
      --这里不插入数据，只进行标记，再下一次循时才插入。如果转科医嘱是最后一条是不用打印新开字样的。
    
      If r_Print.换页 = 3 Then
        n_转科 := 1;
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱打印_Insert;
/
--133819:李南春,2019-01-07,退病历费不用判断挂号记录
Create Or Replace Procedure Zl_病人挂号记录_Delete
(
  单据号_In       门诊费用记录.No%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
  删除门诊号_In   Number := 0,
  非原样退结算_In Varchar2 := Null,
  退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费 3-退附加费 4-退挂号与病历 5-退挂号与附加
  退指定结算_In   病人预交记录.结算方式%Type := Null,
  退号重用_In     Number := 1,
  收回票据号_In   Varchar2 := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  结算方式_In     Varchar2 := Null,
  退预交_In       病人预交记录.冲预交%Type := Null
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
  Cursor c_Opermoney(n_Id 病人预交记录.结帐id%Type) Is
    Select Distinct b.结算方式, -1 * Nvl(b.冲预交, 0) As 冲预交
    From 病人预交记录 B
    Where b.结帐id = n_Id And b.记录性质 = 4 And b.记录状态 = 2 And Nvl(b.冲预交, 0) <> 0;

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
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type;
  d_启用时间       Date;
  d_发生时间       病人挂号记录.发生时间%Type;
  n_出诊记录id     临床出诊记录.Id%Type;
  v_号码           挂号安排.号码%Type;
  n_序号           病人挂号记录.号序%Type;
  v_时间段         时间段.时间段%Type;
  d_检查开始时间   Date;
  d_检查结束时间   Date;
  v_Temp           Varchar2(500);
  v_附加ids        Varchar2(500);
  n_就诊病人id     病人信息.病人id%Type;
  d_就诊时间       就诊登记记录.就诊时间%Type;
  v_结算内容       Varchar2(5000);
  v_当前结算       Varchar2(1000);
  v_结算方式       病人预交记录.结算方式%Type;
  n_三方卡标志     Number;
  n_结算金额       病人预交记录.冲预交%Type;
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

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_附加ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_附加ids := Null;
  End;

  If Nvl(退费类型_In, 0) Not In (2, 3) Then
    v_Temp := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_启用时间 := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      Exception
        When Others Then
          d_启用时间 := Null;
      End;
    End If;

    Select 号别, 号序, 发生时间 Into v_号码, n_序号, d_发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum < 2;

    Begin
      Select a.Id Into n_安排id From 挂号安排 A Where a.号码 = v_号码;
    Exception
      When Others Then
        n_安排id := -1;
    End;

    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 安排id = n_安排id And 审核时间 Is Not Null And
            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.生效时间) As 生效
             From 挂号安排计划 A
             Where a.审核时间 Is Not Null And d_发生时间 Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.安排id = n_安排id) And
            d_发生时间 Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'));
    Exception
      When Others Then
        n_计划id := 0;
    End;

    Begin
      If Nvl(n_计划id, 0) = 0 Then
        Select Decode(To_Char(d_发生时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
        Into v_时间段
        From 挂号安排
        Where ID = n_安排id;
      Else
        Select Decode(To_Char(d_发生时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
        Into v_时间段
        From 挂号安排计划
        Where ID = n_计划id;
      End If;
    Exception
      When Others Then
        v_时间段 := Null;
    End;

    If v_时间段 Is Not Null And d_启用时间 Is Not Null Then
      --检查是否跨模式挂号安排
      Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_检查开始时间, d_检查结束时间
      From 时间段
      Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
      If d_检查开始时间 > d_检查结束时间 Then
        d_检查结束时间 := d_检查结束时间 + 1;
      End If;
      If d_检查开始时间 < d_启用时间 And d_检查结束时间 > d_启用时间 Then
        --获取出诊记录id
        Begin
          Select a.Id
          Into n_出诊记录id
          From 临床出诊记录 A, 临床出诊号源 B
          Where a.号源id = b.Id And b.号码 = v_号码 And 上班时段 = v_时间段 And d_发生时间 Between 开始时间 And 终止时间;
        Exception
          When Others Then
            n_出诊记录id := Null;
        End;
      End If;
    End If;
  End if;
  
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
  
    If n_出诊记录id Is Not Null Then
      Update 临床出诊记录 Set 已约数 = 已约数 - 1 Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
      Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_序号;
    End If;
  
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
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
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
  
    If n_病人id1 Is Not Null And Nvl(退费类型_In, 0) Not In (2, 3) Then
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

  If Nvl(退费类型_In, 0) = 0 Or Nvl(退费类型_In, 0) = 2 Then
    --全退,退病历费
    --门诊费用记录，冲销记录
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
  Elsif Nvl(退费类型_In, 0) = 1 Or Nvl(退费类型_In, 0) = 4 Then
    --退挂号费,退挂号与病历费
    --门诊费用记录，冲销记录
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
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 3 Then
    --退附加费
    --门诊费用记录，冲销记录
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
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 5 Then
    --退挂号与附加费
    --门诊费用记录，冲销记录
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
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    End If;
  End If;

  n_结帐id := 0;
  If n_记帐 = 0 Then
    --获取结帐ID
    Select Nvl(结帐id, 0)
    Into n_结帐id
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum < 2;
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
    If 结算方式_In Is Null And Nvl(退预交_In, 0) = 0 Then
      If 非原样退结算_In Is Not Null Then
        --退款金额获取
        If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
          --如果是单独退病历费,或者只退挂号费,先获取退费金额
          Begin
            --获取本次退款金额
            Select -1 * Sum(Nvl(实收金额, 0)) As 收款金额
            Into n_退款金额
            From 门诊费用记录
            Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id;
          
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
              Set 冲预交 = 冲预交 + (-1 * n_退款金额), 交易说明 = Nvl(交易说明_In, 交易说明)
              Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                   卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                  Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                         操作员编号_In, 操作员姓名_In, -1 * n_退款金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                         Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位, 4
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
            Set 冲预交 = 冲预交 + (-1 * n_退费金额), 交易说明 = Nvl(交易说明_In, 交易说明)
            Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
                 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                       操作员编号_In, 操作员姓名_In, -1 * n_退费金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                       Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位, 4
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
            Select -1 * Sum(Nvl(实收金额, 0)) As 收款金额
            Into n_退款金额
            From 门诊费用记录
            Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id;
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
    Else
      --按结算方式退
      If 结算方式_In is Not Null then
         v_结算内容 := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
        While v_结算内容 Is Not Null Loop
          v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
          v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
        
          v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
          n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
        
          v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
          n_三方卡标志 := To_Number(v_当前结算);
        
          If n_三方卡标志 = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, v_结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                     -1 * n_结算金额, n_销帐id, n_组id, 预交类别, Null, Null, Null, Null, 交易说明_In, 合作单位, 4
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And Rownum < 2;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, v_结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                     -1 * n_结算金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, Nvl(交易说明_In, 交易说明), 合作单位, 4
              
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And
                    (卡类别id Is Not Null Or 结算卡序号 Is Not Null) And Rownum < 2;
          End If;
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
        End Loop;
      End IF;
      n_预交金额 := Nvl(退预交_In, 0);
    End IF;
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
  
    If 收回票据号_In Is Not Null Then
      --光退挂号费,不回收票据
      --退卡收回票据(可能上次挂号使用票据,不能收回)
      Begin
        --从最后一次打印的内容中取
        --81907
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_打印id := Null;
      End;
    
      --先收回原票据
      If n_打印id Is Not Null Then
        Begin
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        Exception
          When Others Then
            Delete From 票据使用明细 Where 打印id = n_打印id And 性质 = 2 And 原因 = 2;
            Insert Into 票据使用明细
              (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
              Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
              From 票据使用明细
              Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --单独退病历费用,不处理汇总记录
  --相关汇总表的处理

  --病人挂号汇总
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
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
    
      If n_出诊记录id Is Not Null Then
        Update 临床出诊记录
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
        Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
        Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_序号;
      End If;
    
      Close c_Registinfo;
    End If;
  End If;

  If n_记帐 = 0 Then
    --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
    For r_Opermoney In c_Opermoney(n_销帐id) Loop
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
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
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

--133819:李南春,2019-01-07,退病历费不用判断挂号记录
Create Or Replace Procedure Zl_病人挂号记录_出诊_Delete
(
  单据号_In       门诊费用记录.No%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因 
  删除门诊号_In   Number := 0,
  非原样退结算_In Varchar2 := Null,
  退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费 3-退附加费 4-退挂号与病历 5-退挂号与附加 
  退指定结算_In   病人预交记录.结算方式%Type := Null,
  退号重用_In     Number := 1,
  结算方式_In     Varchar2 := Null,
  退预交_In       病人预交记录.冲预交%Type := Null,
  收回票据号_In   Varchar2 := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null
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
    From 门诊费用记录 A, 病人挂号记录 C, 人员表 D
    Where a.记录性质 = 4 And a.No = 单据号_In And a.No = c.No And a.记录状态 = v_状态 And c.执行人 = d.姓名(+) And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --该游标用于判断记录是否存在,及费用汇总表处理 
  Cursor c_Moneyinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(应收金额), 0) As 应收, Nvl(Sum(实收金额), 0) As 实收, Nvl(Sum(结帐金额), 0) As 结帐
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = v_状态 And NO = 单据号_In
    Group By 病人科室id, 开单部门id, 执行部门id, 收入项目id;
  r_Moneyrow c_Moneyinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额 
  Cursor c_Opermoney(n_Id 病人预交记录.结帐id%Type) Is
    Select Distinct b.结算方式, -1 * Nvl(b.冲预交, 0) As 冲预交
    From 病人预交记录 B
    Where b.结帐id = n_Id And b.记录性质 = 4 And b.记录状态 = 2 And Nvl(b.冲预交, 0) <> 0;

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
  n_序号           病人挂号记录.号序%Type;
  n_就诊病人id     病人信息.病人id%Type;
  d_就诊时间       就诊登记记录.就诊时间%Type;
  n_出诊记录id     临床出诊记录.Id%Type;
  v_结算内容       Varchar2(5000);
  v_当前结算       Varchar2(1000);
  v_附加ids        Varchar2(500);
  v_Temp           Varchar2(500);
  v_结算方式       病人预交记录.结算方式%Type;
  n_三方卡标志     Number;
  n_结算金额       病人预交记录.冲预交%Type;
  n_检查数         Number;
Begin
  n_组id           := Zl_Get组id(操作员姓名_In);
  v_退指定结算方式 := 退指定结算_In;

  If Nvl(退费类型_In, 0) Not In (2, 3) Then
    Select 出诊记录id, 号序 Into n_出诊记录id, n_序号 From 病人挂号记录 Where NO = 单据号_In And Rownum < 2;
  End if;

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

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_附加ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_附加ids := Null;
  End;

  --1.预约处理 
  If Nvl(n_预约挂号, 0) = 1 Then
    --减少已约数 
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    n_检查数 := Null;
    Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) - 1 Where ID = n_出诊记录id Returning 已约数 Into n_检查数;
    If Nvl(n_检查数, 0) < 0 Then
      Update 临床出诊记录 Set 已约数 = 0 Where ID = n_出诊记录id;
    End If;
  
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) - 1
    Where 日期 = Trunc(r_Registrow.发生时间) And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
          Nvl(医生姓名, '-') = Nvl(r_Registrow.医生姓名, '-') And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
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
    Update 临床出诊序号控制
    Set 挂号状态 = 0, 操作员姓名 = Null
    Where 挂号状态 = 2 And 记录id = n_出诊记录id And 序号 = n_序号;
  
    Update 临床出诊序号控制
    Set 挂号状态 = 4, 操作员姓名 = Null
    Where 挂号状态 = 2 And 记录id = n_出诊记录id And 备注 = To_Char(n_序号);
  
    --添加病人挂号记录的 冲销记录 
    Select 病人挂号记录_Id.Nextval, Sysdate Into n_挂号id, d_Date From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1 And 记录性质 = 2;
    If Sql%NotFound Then
      v_Err_Msg := '预约单【' || 单据号_In || '】不存在或由于并发原因已经被取消预约';
      Raise Err_Item;
    End If;
  
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式, 出诊记录id, 预约操作员, 预约操作员编号)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式,
             n_出诊记录id, 预约操作员, 预约操作员编号
      From 病人挂号记录
      Where NO = 单据号_In;
  
    Update 门诊费用记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
    Insert Into 门诊费用记录
      (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别,
       收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id,
       执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id, 费用状态, 待转出)
      Select 病人费用记录_Id.Nextval, 记录性质, NO, 实际票号, 2, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式,
             病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, -1 * 应收金额,
             -1 * 实收金额, 划价人, 开单部门id, 开单人, 发生时间, d_Date, 执行部门id, 执行人, -1, 执行时间, 结论, 操作员编号_In, 操作员姓名_In, Null, Null,
             保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id, 费用状态, 待转出
      From 门诊费用记录
      Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 3;
  
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
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
    --不是光退病历费时处理 
    --更新挂号序号状态 
    If 退号重用_In = 1 Then
      Update 临床出诊序号控制
      Set 挂号状态 = 0, 操作员姓名 = Null
      Where 挂号状态 = 1 And 记录id = n_出诊记录id And 序号 = n_序号;
    
      Update 临床出诊序号控制
      Set 挂号状态 = 4, 操作员姓名 = Null
      Where 挂号状态 = 1 And 记录id = n_出诊记录id And 备注 = To_Char(n_序号);
    Else
      Update 临床出诊序号控制
      Set 挂号状态 = 4, 操作员姓名 = 操作员姓名_In
      Where 挂号状态 = 1 And 记录id = n_出诊记录id And (序号 = n_序号 Or 备注 = To_Char(n_序号));
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
  If Nvl(退费类型_In, 0) = 0 Or Nvl(退费类型_In, 0) = 2 Then
    --全退,退病历费 
    --门诊费用记录，冲销记录 
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
  Elsif Nvl(退费类型_In, 0) = 1 Or Nvl(退费类型_In, 0) = 4 Then
    --退挂号费,退挂号与病历费 
    --门诊费用记录，冲销记录 
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
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录 
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 3 Then
    --退附加费 
    --门诊费用记录，冲销记录 
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
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录 
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 5 Then
    --退挂号与附加费 
    --门诊费用记录，冲销记录 
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
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
  
    --原始记录 
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    End If;
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
    If 结算方式_In Is Null And Nvl(退预交_In, 0) = 0 Then
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
              Set 冲预交 = 冲预交 + (-1 * n_退款金额), 交易说明 = Nvl(交易说明_In, 交易说明)
              Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                   卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                  Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                         操作员编号_In, 操作员姓名_In, -1 * n_退款金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                         Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位,
                         4
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
                Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And
                      Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
              
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
            Set 冲预交 = 冲预交 + (-1 * n_退费金额), 交易说明 = Nvl(交易说明_In, 交易说明)
            Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                       操作员编号_In, 操作员姓名_In, -1 * n_退费金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                       Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位, 4
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
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * 冲预交, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
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
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 摘要 = '医保挂号' And
                  冲预交 = n_退款金额 And Rownum < 2;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                     -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 冲预交 = n_退款金额 And
                    Rownum < 2;
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
    Else
      --按结算方式退 
      If 结算方式_In Is Not Null Then
        v_结算内容 := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的 
        While v_结算内容 Is Not Null Loop
          v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
          v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
        
          v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
          n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
        
          v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
          n_三方卡标志 := To_Number(v_当前结算);
        
          If n_三方卡标志 = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 结算号码)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, v_结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                     -1 * n_结算金额, n_销帐id, n_组id, 预交类别, Null, Null, Null, Null, 交易说明_In, 合作单位, 4, 结算号码
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And Rownum < 2;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 结算号码)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, v_结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                     -1 * n_结算金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, Nvl(交易说明_In, 交易说明), 合作单位, 4, 结算号码
              
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And
                    (卡类别id Is Not Null Or 结算卡序号 Is Not Null) And Rownum < 2;
          End If;
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
        End Loop;
      End If;
      n_预交金额 := Nvl(退预交_In, 0);
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
  
    If 收回票据号_In Is Not Null Then
      --光退挂号费,不回收票据 
      --退卡收回票据(可能上次挂号使用票据,不能收回) 
      Begin
        --从最后一次打印的内容中取 
        --81907 
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_打印id := Null;
      End;
    
      --先收回原票据 
      If n_打印id Is Not Null Then
        Begin
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        Exception
          When Others Then
            Delete From 票据使用明细 Where 打印id = n_打印id And 性质 = 2 And 原因 = 2;
            Insert Into 票据使用明细
              (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
              Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
              From 票据使用明细
              Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --单独退病历费用,不处理汇总记录 
  --相关汇总表的处理 
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
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
      n_检查数 := Null;
      Update 临床出诊记录
      Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
      Where ID = n_出诊记录id
      Returning 已挂数 Into n_检查数;
    
      If Nvl(n_检查数, 0) < 0 Then
        Update 临床出诊记录 Set 已挂数 = 0 Where ID = n_出诊记录id;
      End If;
    
      Update 病人挂号汇总
      Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
      Where 日期 = Trunc(r_Registrow.发生时间) And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
            Nvl(医生姓名, '-') = Nvl(r_Registrow.医生姓名, '-') And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
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
  End If;

  If n_记帐 = 0 Then
    --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款) 
    For r_Opermoney In c_Opermoney(n_销帐id) Loop
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
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
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
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式, 出诊记录id)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式,
             n_出诊记录id
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
End Zl_病人挂号记录_出诊_Delete;
/

--134969:李南春,2019-01-11,预交支付检查
CREATE OR REPLACE Procedure Zl_预约挂号接收_Insert
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
  更新交款余额_In  Number := 0, --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
  摘要_In          病人挂号记录.摘要%Type := Null
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, No, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, Id, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(预交类别, 2) = 1 Having
     Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By No, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, No;

  v_Err_Msg Varchar2(255);
  Err_Item    Exception;
  Err_Special Exception;

  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;
  v_号别     门诊费用记录.计算单位%Type;
  v_号序     门诊费用记录.发药窗口%Type;
  v_排队号码 排队叫号队列.排队号码 %Type;
  v_预约方式 病人挂号记录.预约方式 %Type;

  n_预交金额      病人预交记录.金额%Type;
  n_病人余额      病人余额.预交余额%Type;
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

  d_Date       Date;
  d_预约时间   门诊费用记录.发生时间%Type;
  d_发生时间   Date;
  d_排队时间   Date;
  n_时段       Number := 0;
  n_存在       Number := 0;
  v_排队序号   排队叫号队列.排队序号%Type;
  n_结算模式   病人信息.结算模式%Type;
  v_付款方式   病人挂号记录.医疗付款方式%Type;
  v_操作员姓名 病人挂号记录.接收人%Type;
  n_接收模式   Number := 0;
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  n_接收模式      := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);

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
      Select Max(接收人) Into v_操作员姓名 From 病人挂号记录 Where 记录性质 = 2 And 记录状态 In (1, 3) And NO = No_In;
      If v_操作员姓名 Is Null Then
        v_Err_Msg := '当前预约挂号单已被取消';
        Raise Err_Item;
      Else
        If v_操作员姓名 = 操作员姓名_In Then
          v_Err_Msg := '当前预约挂号单已被接收';
          Raise Err_Special;
        Else
          v_Err_Msg := '当前预约挂号单已被其它人接收';
          Raise Err_Item;
        End If;
      End If;
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
  --分时段的号别，只能当天接收
  If n_时段 = 1 And 三方调用_In = 0 And n_接收模式 = 0 Then
    If Trunc(发生时间_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '分时段的预约挂号单只能当天接收！';
      Raise Err_Item;
    End If;
  End If;

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
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 摘要 = Nvl(摘要_In, 摘要)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In,
      摘要 = Nvl(摘要_In, 摘要)
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
               登记时间, 发生时间, 操作员编号, 操作员姓名, Nvl(摘要_In, 摘要), v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In,
               Nvl(登记时间_In, Sysdate), 发生时间, Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式
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
        --预约接收时，改变记录标志
        Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
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
      (n_预交id, 4, 1, No_In, 病人id_In, Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
       n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 结帐id_In, 4);

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
    Select Nvl(Sum(Nvl(预交余额, 0) - Nvl(费用余额, 0)), 0)
    Into n_病人余额
    From 病人余额
    Where 病人id In (Select /*+cardinality(d,10)*/
                    d.Column_Value
                   From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(性质, 0) = 1 And Nvl(类型, 0) = 1;
    if n_病人余额 < 预交支付_In Then
      v_Err_Msg := '病人的当前预交余额为 ' || Ltrim(To_Char(n_病人余额, '9999999990.00')) || '，小于本次支付金额 ' ||
                   Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，支付失败！';
      Raise Err_Item;
    End if;

    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
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
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (r_Deposit.病人id, Nvl(1, 2), -1 * n_当前金额, 1);
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
    IF n_预交金额 > 0 Then
      v_Err_Msg := '病人的当前预交余额小于本次支付金额 ' || Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，不能继续操作！';
      Raise Err_Item;
    End IF;
  End If;

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       Null, Null, Null, Null, Null, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(现金支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 0) = 0 Then
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

  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 0) = 0 Then
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
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) >= d_Date;
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
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_Insert;
/

--134441:李南春,2019-01-16,挂号检查项目是否一致
--134969:李南春,2019-01-11,预交支付检查
CREATE OR REPLACE Procedure Zl_预约挂号接收_出诊_Insert
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
  结算方式_In      Varchar2, --现金的结算名称
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
  更新交款余额_In  Number := 0, --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
  摘要_In          病人挂号记录.摘要%Type := Null
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, No, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, Id, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(预交类别, 2) = 1 Having
     Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By No, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, No;

  v_Err_Msg Varchar2(255);
  Err_Item    Exception;
  Err_Special Exception;
  v_操作员姓名 病人挂号记录.接收人%Type;
  v_现金       结算方式.名称%Type;
  v_个人帐户   结算方式.名称%Type;
  v_队列名称   排队叫号队列.队列名称%Type;
  v_号别       门诊费用记录.计算单位%Type;
  v_号序       门诊费用记录.发药窗口%Type;
  v_排队号码   排队叫号队列.排队号码 %Type;
  v_预约方式   病人挂号记录.预约方式 %Type;

  n_预交金额      病人预交记录.金额%Type;
  n_病人余额      病人余额.预交余额%Type;
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

  d_Date         Date;
  d_预约时间     门诊费用记录.发生时间%Type;
  d_发生时间     Date;
  d_排队时间     Date;
  n_时段         Number := 0;
  n_存在         Number := 0;
  v_结算内容     Varchar2(2000);
  v_当前结算     Varchar2(500);
  n_结算金额     病人预交记录.冲预交%Type;
  v_结算号码     病人预交记录.结算号码%Type;
  v_结算方式     病人预交记录.结算方式%Type;
  n_三方卡标志   Number(3);
  v_排队序号     排队叫号队列.排队序号%Type;
  n_结算模式     病人信息.结算模式%Type;
  v_付款方式     病人挂号记录.医疗付款方式%Type;
  n_接收模式     Number := 0;
  n_出诊记录id   病人挂号记录.出诊记录id%Type;
  n_新出诊记录id 病人挂号记录.出诊记录id%Type;
  n_号源id       临床出诊记录.号源id%Type;
  n_预约顺序号   临床出诊序号控制.预约顺序号%Type;
  n_旧分时段     临床出诊记录.是否分时段%Type;
  n_旧序号控制   临床出诊记录.是否序号控制%Type;
  n_旧科室id     临床出诊记录.科室id%Type;
  n_旧项目id     临床出诊记录.项目id%Type;
  n_旧医生id     临床出诊记录.医生id%Type;
  n_挂号模式     Number(3);
  d_启用时间     Date;
  v_Paratemp     Varchar2(500);
  v_Registtemp   Varchar2(500);
  n_检查         Number(3);
  n_序号控制     临床出诊记录.是否序号控制%Type;
  v_旧上班时段   临床出诊记录.上班时段%Type;
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  v_Paratemp      := Nvl(zl_GetSysParameter('挂号排班模式'), 0);
  n_接收模式      := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);
  n_挂号模式      := To_Number(Substr(v_Paratemp, 1, 1));
  If n_挂号模式 = 1 Then
    Begin
      d_启用时间 := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
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
  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式, 出诊记录id
    Into v_号别, v_号序, d_预约时间, d_发生时间, v_预约方式, n_出诊记录id
    From 病人挂号记录
    Where 记录性质 = 2 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(接收人) Into v_操作员姓名 From 病人挂号记录 Where 记录性质 = 2 And 记录状态 In (1, 3) And NO = No_In;
      If v_操作员姓名 Is Null Then
        v_Err_Msg := '当前预约挂号单已被取消';
        Raise Err_Item;
      Else
        If v_操作员姓名 = 操作员姓名_In Then
          v_Err_Msg := '当前预约挂号单已被接收';
          Raise Err_Special;
        Else
          v_Err_Msg := '当前预约挂号单已被其它人接收';
          Raise Err_Item;
        End If;
      End If;
  End;

  --判断是否分时段
  Select Nvl(是否分时段, 0), 号源id, Nvl(是否序号控制, 0)
  Into n_时段, n_号源id, n_序号控制
  From 临床出诊记录
  Where ID = n_出诊记录id;

  If n_时段 = 1 And 三方调用_In = 0 And n_接收模式 = 0 Then
    If Trunc(发生时间_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '分时段的预约挂号单只能当天接收！';
      Raise Err_Item;
    End If;
  End If;

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

  If d_启用时间 Is Not Null Then
    If d_发生时间 < d_启用时间 Then
      v_Err_Msg := '当前预约挂号单属于出诊表排班模式安排，不能在' || To_Char(d_启用时间, 'yyyy-mm-dd hh24:mi:ss') || '之前接收!';
      Raise Err_Item;
    End If;
  End If;

  If Not v_号序 Is Null Then
    If 号序_In Is Null Then
      Update 临床出诊序号控制 Set 挂号状态 = 0 Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
    Else
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        If n_时段 = 0 And 三方调用_In = 0 Then
          --提前接收或延迟接收
          Update 临床出诊序号控制 Set 挂号状态 = 0 Where 序号 = v_号序 And 记录id = n_出诊记录id;

          Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
          Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
          From 临床出诊记录
          Where ID = n_出诊记录id;
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                  Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
              Raise Err_Item;
          End;

          Begin
            Select 1
            Into n_存在
            From 临床出诊序号控制
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          Exception
            When Others Then
              n_存在 := 0;
          End;

          If n_存在 = 1 Then
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          Else
            --号码已被使用的情况
            Select Min(序号) Into v_号序 From 临床出诊序号控制 Where 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
            If v_号序 Is Null Then
              v_Err_Msg := '接收当天没有可用序号,无法接收!';
              Raise Err_Item;
            End If;
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          End If;
        Else
          Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
          Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
          From 临床出诊记录
          Where ID = n_出诊记录id;
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                  Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
              Raise Err_Item;
          End;
          Update 临床出诊序号控制
          Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
          Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
          Returning 预约顺序号 Into n_预约顺序号;

          Update 临床出诊序号控制
          Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
          Where 序号 = v_号序 And 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
          If Sql% RowCount = 0 Then
            v_Err_Msg := '接收当天序号' || v_号序 || '已被其它人使用,无法接收.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not 号序_In Is Null Then
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
        Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
        From 临床出诊记录
        Where ID = n_出诊记录id;
        Begin
          Select ID
          Into n_新出诊记录id
          From 临床出诊记录
          Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
            Raise Err_Item;
        End;
        Update 临床出诊序号控制
        Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
        Returning 预约顺序号 Into n_预约顺序号;
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
        Where 序号 = 号序_In And 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '接收当天序号' || 号序_In || '已被其它人使用,无法接收.';
          Raise Err_Item;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id;

      End If;
      v_号序 := 号序_In;
    Else
      v_号序 := Null;
    End If;
  End If;

  --检查挂号项目
  Select Count(1)
  Into n_Count
  From 临床出诊记录 a, 门诊费用记录 b
  Where a.Id = Nvl(n_新出诊记录id, n_出诊记录id) And b.No = No_In And b.序号 = 1 And a.项目id = b.收费细目id And a.科室id = b.执行部门id;
  If n_Count = 0 Then
    v_Err_Msg := '挂号项目、科室不一致，无法接收！';
    Raise Err_Item;
  End If;
  
  --更新门诊费用记录
  Update 门诊费用记录
  Set 记录状态 = 1, 实际票号 = Decode(Nvl(记帐费用_In, 0), 1, Null, 票据号_In), 结帐id = Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In),
      结帐金额 = Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额), 发药窗口 = 诊室_In, 病人id = 病人id_In, 标识号 = 门诊号_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
      性别 = 性别_In, 付款方式 = 付款方式_In, 费别 = 费别_In, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In,
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 摘要 = Nvl(摘要_In, 摘要)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('挂号排班模式');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_发生时间 Then
        v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '未启用出诊表排班模式,目前无法接收!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_检查
      From 临床出诊记录
      Where ID = Nvl(n_新出诊记录id, n_出诊记录id) And d_发生时间 Between 停诊开始时间 And 停诊终止时间;
    Exception
      When Others Then
        n_检查 := 0;
    End;
    If n_检查 = 1 And Not (n_时段 = 1 And n_序号控制 = 1) Then
      v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '的安排已经被停诊,无法接收!';
      Raise Err_Item;
    End If;
  End If;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In,
      出诊记录id = Nvl(n_新出诊记录id, n_出诊记录id), 摘要 = Nvl(摘要_In, 摘要)
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
         摘要, 号序, 预约, 预约方式, 接收人, 接收时间, 预约时间, 险类, 医疗付款方式, 出诊记录id)
        Select n_挂号id, No_In, 1, 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, 执行人, 0, Null,
               登记时间, 发生时间, 操作员编号, 操作员姓名, Nvl(摘要_In, 摘要), v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In,
               Nvl(登记时间_In, Sysdate), 发生时间, Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, Nvl(n_新出诊记录id, n_出诊记录id)
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
  If Nvl(记帐费用_In, 0) = 0 Then
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 Then
      v_结算内容 := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);

        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));

        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);

        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);

        If n_三方卡标志 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, Null, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4, v_结算号码);
          If Nvl(结算卡序号_In, 0) <> 0 Then
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
            Zl_病人卡结算记录_Insert(结算卡序号_In, n_消费卡id, v_结算方式, n_结算金额, 卡号_In, Null, 登记时间_In, Null, 结帐id_In, n_预交id);
          End If;
        End If;

        If Nvl(更新交款余额_In, 0) = 0 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + n_结算金额
          Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金)
          Returning 余额 Into n_返回值;

          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, Nvl(v_结算方式, v_现金), 1, n_结算金额);
            n_返回值 := n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End If;

        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
    End If;
  End If;

  --对于就诊卡通过预交金挂号
  If Nvl(预交支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Select Nvl(Sum(Nvl(预交余额, 0) - Nvl(费用余额, 0)), 0)
    Into n_病人余额
    From 病人余额
    Where 病人id In (Select /*+cardinality(d,10)*/
                    d.Column_Value
                   From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(性质, 0) = 1 And Nvl(类型, 0) = 1;
    if n_病人余额 < 预交支付_In Then
      v_Err_Msg := '病人的当前预交余额为 ' || Ltrim(To_Char(n_病人余额, '9999999990.00')) || '，小于本次支付金额 ' ||
                   Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，支付失败！';
      Raise Err_Item;
    End if;

    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
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
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (r_Deposit.病人id, Nvl(1, 2), -1 * n_当前金额, 1);
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
    IF n_预交金额 > 0 Then
      v_Err_Msg := '病人的当前预交余额小于本次支付金额 ' || Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，不能继续操作！';
      Raise Err_Item;
    End IF;
  End If;

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       Null, Null, Null, Null, Null, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 0) = 0 Then
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
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) >= d_Date;
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
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_出诊_Insert;
/

--135930:殷瑞,2018-12-24,修正收费单和记账单判断的问题
Create Or Replace Procedure Zl_未发药品记录_分配发药窗口
(
  No_In       药品收发记录.No%Type,
  单据_In     药品收发记录.单据%Type,
  药房id_In   药品收发记录.库房id%Type,
  发药窗口_In 药品收发记录.发药窗口%Type := Null
) Is
Begin
  If 发药窗口_In Is Not Null Then
    Update 药品收发记录 Set 发药窗口 = 发药窗口_In Where NO = No_In And 单据 = 单据_In And 库房id = 药房id_In;
  
    Update 未发药品记录 Set 发药窗口 = 发药窗口_In Where NO = No_In And 单据 = 单据_In And 库房id = 药房id_In;
  
    Update 门诊费用记录
    Set 发药窗口 = 发药窗口_In
    Where NO = No_In And 记录性质 = Decode(单据_In, 8, 1, 2) And 执行部门id = 药房id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_未发药品记录_分配发药窗口;
/

--135348:余伟节,2018-12-17,病人信息合并排除病人地址信息
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
           '病人免疫记录', '病人信息从表', '病人医疗卡属性','病人地址信息') Having Max(Decode(b.Column_Name, '病人ID', 1, 0)) <> 0
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
      v_Sql := 'Select Count(病人Id) From H' || r_Bak.Table_Name || ' Where 病人Id In(:1,:2)';
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
    If Arronbase(v_Loop) = '病人照片' Then
      Select Count(1) Into n_Have From 病人照片 Where 病人id = v_保留id;
      If n_Have = 1 Then
        Delete From 病人照片 Where 病人id = v_合并id;
      End If;
    End If;
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
      入院时间 = Null, 出院时间 = Null, 在院 = Decode(Nvl(r_Infob.在院, 0), 1, 1, Null), 健康号 = Nvl(r_Infob.健康号, r_Infoa.健康号)
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

--130781:胡俊勇,2018-12-13,检验危急值记录删除
Create Or Replace Procedure Zl_病人危急值记录_Delete(Id_In In 病人危急值记录.Id%Type) Is
Begin
  Delete 业务消息清单
  Where 类型编码 = 'ZLHIS_LIS_003' And 业务标识 = (Select To_Char(医嘱id) From 病人危急值记录 Where ID = Id_In);
  Delete 病人危急值记录 Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人危急值记录_Delete;
/

--134581:李小东,2018-11-28,标本登记-医嘱id传参超长4000拆分
Create Or Replace Procedure Zl_病人医嘱发送_Sampleinput
(
  医嘱id      In Varchar2,
  接收人_In   In 病人医嘱发送.接收人%Type := Null,
  接收批次_In In 病人医嘱发送.接收批次%Type := 0,
  人员编号_In In 人员表.编号%Type := Null,
  人员姓名_In In 人员表.姓名%Type := Null,
  送检人_In   In 病人医嘱发送.送检人%Type := Null
) Is
  --未审核的费用行(不包含药品)
  Cursor c_Verify(v_医嘱id In Number) Is
    Select Distinct 记录性质, NO, 序号, 记录状态, 门诊标志
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
    Select Distinct 记录性质, NO, 序号, 记录状态, 门诊标志
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
  v_Records  Varchar2(4000);
  v_Currrec  Varchar2(50);
  v_Fields   Varchar2(50);
  v_Count    Number(18);
  v_病人id   病人医嘱记录.病人id%Type;
  v_主页id   病人医嘱记录.主页id%Type;
  v_是否出院 Number; --0=出院,1=在院
  v_病人来源 病人医嘱记录.病人来源%Type;
  v_Date     Date;
  Err_Custom Exception;
  v_Error Varchar2(100);
  n_Par   Number;
Begin
  Select Sysdate Into v_Date From Dual;
  --执行后自动审核对应的记帐划价单(不包含药品)
  Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_执行 From Dual;

  v_Records := 医嘱id || '|';

  While v_Records Is Not Null Loop
  
    v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields  := v_Currrec;
    v_医嘱id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    v_相关id  := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    If 接收人_In Is Null Then
      Update 病人医嘱发送 Set 接收人 = Null, 接收时间 = Null, 接收批次 = Null Where 医嘱id In (v_医嘱id, v_相关id);
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
      --判断是否已出院，如果已出院则不完成登记
      Begin
        Select a.病人id, a.主页id, a.病人来源
        Into v_病人id, v_主页id, v_病人来源
        From 病人医嘱记录 A, 病案主页 B
        Where a.病人id = b.病人id And a.主页id = b.主页id(+) And a.Id = v_医嘱id;
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
          Select Nvl(样本条码, 0) Into v_样本条码 From 病人医嘱发送 Where 医嘱id = v_医嘱id;
          If v_样本条码 = 0 Then
            v_Error := '病人已出院不能完成登记!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      --检查医嘱是否收费
      n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
      If n_Par = 1 Then
        For r_Samplequest In c_Samplequest(v_相关id) Loop
          For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
            If r_Verify.记录状态 = 0 Then
              If r_Verify.门诊标志 = 1 Then
                v_Error := '标本未收费，不允许执行，请联系管理员！';
                Raise Err_Custom;
              Elsif r_Verify.门诊标志 = 2 Then
                v_Error := '标本未记账，不允许执行，请联系管理员！';
                Raise Err_Custom;
              End If;
            End If;
          End Loop;
        End Loop;
      End If;
    
      Update 病人医嘱发送
      Set 接收人 = 接收人_In, 接收时间 = v_Date, 接收批次 = 接收批次_In, 重采标本 = Null, 送检人 = 送检人_In
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
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱发送_Sampleinput;
/

--134626:胡俊勇,2018-11-26,卫材医嘱超期收回
CREATE OR REPLACE Procedure Zl_病人医嘱记录_收回
(
  --功能：将指定医嘱超期发送部分收回。如果上次发送没有产生费用，则仅收回医嘱的上次执行时间。
  --参数：
  --      收回量_IN=对西药、中成药为按住院单位的收回量,对中药为收回付数,对其它医嘱为收回总量或次数。
  --      医嘱ID_IN=每条要收回的医嘱记录的ID(明细存储的ID),对成药或配方,不一定包含给药途径或用法煎法(可能为叮嘱而未读取)
  --      上次时间_IN=医嘱超期发送部分收回后应该还原的上次执行时间(严格按频率计算得来),为空时表示被全部收回了。
  --      NO_IN=当收回要产生负数费用记录时，为新生成记录的单据号(供费用及药品使用),当前处理的只是新NO的一部份。
  --            因为药品可能分批,所以序号在处理时取。
  --            如果全是划价单（传入值为：调整划价单），则不产生负数单据，直接修改或删除划价单
  收回量_In     病人医嘱发送.发送数次%Type,
  医嘱id_In     病人医嘱记录.Id%Type,
  上次时间_In   病人医嘱记录.上次执行时间%Type,
  收回时间_In   Date,
  No_In         住院费用记录.No%Type := Null,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null
) Is
  --收回医嘱对应的发送费用明细的剩余数量,按后产生的费用先收回
  --剩余数量没有排开已申请的数量部份，在产生新申请时覆盖原来的申请
  --对药品和卫材，对一个数量，可能存在未执行和已执行部分，需分别填写申请记录，且以未执行优先
  --执行标志=0-未执行,1-已执行；药品的有部分执行，以收发记录中的明细量区分为准；非药品的只优先处理未执行的
  Cursor c_Detail Is
    Select *
    From (With 医嘱费用记录 As (Select Max(Decode(b.记录状态, 2, 0, b.Id)) As 费用id, b.No, Nvl(b.价格父号, b.序号) As 序号, b.收费细目id,
                                 b.病人病区id, Sum(Nvl(b.付数, 1) * b.数次) As 剩余数量, b.收费类别, Max(Nvl(b.执行状态, 0)) As 执行状态, d.跟踪在用,
                                 c.诊疗类别, c.医嘱内容, c.单次用量, Max(b.记录状态) As 记录状态, Max(b.登记时间) As 登记时间, Nvl(e.收费方式, 0) As 收费方式
                          From 病人医嘱发送 A, 住院费用记录 B, 病人医嘱记录 C, 材料特性 D, 病人医嘱计价 E
                          Where a.医嘱id = 医嘱id_In And a.No = b.No And a.记录性质 = b.记录性质 And a.医嘱id = b.医嘱序号 And
                                b.价格父号 Is Null And b.收费细目id = d.材料id(+) And c.Id = 医嘱id_In And e.医嘱id(+) = b.医嘱序号 And
                                e.收费细目id(+) = b.收费细目id
                          Group By b.No, b.记录性质, Nvl(b.价格父号, b.序号), b.收费细目id, b.病人病区id, b.收费类别, d.跟踪在用, c.诊疗类别, c.医嘱内容,
                                   c.单次用量, e.收费方式
                          Having Sum(Nvl(b.付数, 1) * b.数次) > 0)
           Select 费用id, NO, 序号, 收费细目id, 病人病区id, 收费类别, 跟踪在用, 诊疗类别, 医嘱内容, 单次用量, 剩余数量, Null As 已执行量, Null As 未执行量,
                  执行状态 As 执行标志, 记录状态, 登记时间, 收费方式
           From 医嘱费用记录
           Where 收费类别 Not In ('5', '6', '7') And Not (收费类别 = '4' And Nvl(跟踪在用, 0) = 1)
           Union All
           Select a.费用id, a.No, a.序号, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, a.诊疗类别, a.医嘱内容, a.单次用量, a.剩余数量, 0 As 已执行量,
                  Sum(Nvl(b.付数, 1) * b.实际数量) As 未执行量, 0 As 执行标志, a.记录状态, Max(a.登记时间) As 登记时间, 收费方式
           From 医嘱费用记录 A, 药品收发记录 B
           Where (a.收费类别 In ('5', '6', '7') Or (a.收费类别 = '4' And Nvl(a.跟踪在用, 0) = 1)) And a.费用id = b.费用id And
                 a.No = b.No And b.单据 In (9, 10, 25, 26) And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null
           Group By a.费用id, a.No, a.序号, a.记录状态, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, a.诊疗类别, a.医嘱内容, a.单次用量, a.剩余数量, 收费方式
           Having Sum(Nvl(b.付数, 1) * b.实际数量) > 0
           Union All
           Select a.费用id, a.No, a.序号, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, a.诊疗类别, a.医嘱内容, a.单次用量, a.剩余数量,
                  Sum(Nvl(b.付数, 1) * b.实际数量) As 已执行量, 0 As 未执行量, 1 As 执行标志, a.记录状态, Max(a.登记时间) As 登记时间, 收费方式
           From 医嘱费用记录 A, 药品收发记录 B
           Where (a.收费类别 In ('5', '6', '7') Or (a.收费类别 = '4' And Nvl(a.跟踪在用, 0) = 1)) And a.费用id = b.费用id And
                 a.No = b.No And b.单据 In (9, 10, 25, 26) And Not (Mod(b.记录状态, 3) = 1 And b.审核人 Is Null)
           Group By a.费用id, a.No, a.序号, a.记录状态, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, a.诊疗类别, a.医嘱内容, a.单次用量, a.剩余数量, 收费方式
           Having Sum(Nvl(b.付数, 1) * b.实际数量) > 0)
           Order By Decode(诊疗类别, '5', 0, '6', 0, '7', 0, 收费细目id), 执行标志, 登记时间 Desc;


  Cursor c_Applay(v_费用ids Varchar2) Is
    Select a.费用id, b.No, b.序号, a.数量, a.申请时间, a.申请类别
    From 病人费用销帐 A, 住院费用记录 B
    Where a.费用id = b.Id And a.申请部门id = a.审核部门id And a.申请时间 = 收回时间_In And
          a.费用id In (Select * From Table(Cast(f_Num2list(v_费用ids) As Zltools.t_Numlist)))
    Order By NO, 序号;

  --包含指定药品长嘱发送时产生的相关费用及药品/卫材记录信息(因多次发送有多条记录,分批的已在界面禁止)
  --药品医嘱填写了"病人医嘱发送"记录,对应的给药途径不一定填写了的(可能为叮嘱),且NO不同。
  --因为要收回的次数可能包含了多次发送的内容,所以要将多次发送的收发记录都取出来，多次发送时，划价的先收回（修改或删除）
  Cursor c_Drug Is
    Select a.病人id, a.主页id, d.姓名, Nvl(Nvl(x.剂量系数, y.换算系数), 1) As 剂量系数, Nvl(x.住院包装, 1) As 住院包装,
           Nvl(x.最大效期, y.最大效期) As 最大效期, Nvl(b.付数, 1) * b.实际数量 As 数量, b.Id As 收发id, b.单据, b.药品id, b.对方部门id, b.库房id,
           b.费用id, Nvl(Nvl(x.药房分批, y.在用分批), 0) As 分批, b.批次, b.批号, b.效期, a.记录状态, a.No, a.序号, a.收费细目id, a.执行状态 As 执行标志
    From 住院费用记录 A, 药品收发记录 B, 病人医嘱发送 C, 病人信息 D, 药品规格 X, 材料特性 Y
    Where c.医嘱id = 医嘱id_In And a.No = c.No And a.记录性质 = c.记录性质 And a.记录状态 In (0, 1, 3) And a.医嘱序号 + 0 = 医嘱id_In And
          a.No = b.No And a.Id = b.费用id + 0 And b.单据 In (9, 10, 25, 26) And (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And
          a.病人id = d.病人id And b.药品id = x.药品id(+) And b.药品id = y.材料id(+)
    Order By a.记录状态, b.No Desc, b.Id Desc;

  --包含非药长嘱(含给药途径)发送时所产生的费用(因多个收入而有多条记录)
  --对非药医嘱,直接收回指定量,不管多次发送(如果多次发送价格不同,则收回的价格是以最后次的；不然就要根据多个收入依次减收回量)。
  --卫材本身是售价单位，无需住院单位转换
  --非药长嘱都填写了发送记录(除开了叮嘱及护理等级)
  --一天只收一次或一次发送只收一次的项目暂时不支持负数申请
  Cursor c_Other Is
    With 医嘱费用记录 As
     (Select a.No, a.序号, a.记录状态, a.收费细目id, a.Id As 费用id, a.数次 As 剩余数量, Nvl(a.执行状态, 0) As 执行状态, a.医嘱序号, b.发送号,
             c.数量 As 对照数量, Nvl(c.收费方式, 0) As 收费方式, a.收费类别
      From 住院费用记录 A, 病人医嘱发送 B, 病人医嘱计价 C
      Where a.No = b.No And a.记录性质 = b.记录性质 And a.医嘱序号 + 0 = b.医嘱id And b.医嘱id = 医嘱id_In And a.医嘱序号 = c.医嘱id(+) And
            a.收费细目id = c.收费细目id(+))
    Select a.No, a.序号, a.费用id, a.剩余数量, a.收费细目id, a.记录状态, a.执行状态, a.对照数量, a.收费方式, a.收费类别
    From (Select a.No, a.序号, a.记录状态, a.收费细目id, a.费用id, a.剩余数量, a.对照数量, a.执行状态, a.医嘱序号, a.收费方式, a.收费类别
           From 医嘱费用记录 A
           Where a.记录状态 In (1, 3) And a.发送号 = (Select Max(发送号) From 医嘱费用记录 Where 记录状态 In (1, 3))
           Union All
           Select a.No, a.序号, a.记录状态, a.收费细目id, a.费用id, a.剩余数量, a.对照数量, a.执行状态, a.医嘱序号, a.收费方式, a.收费类别
           From 医嘱费用记录 A
           Where a.记录状态 = 0) A
    Order By a.收费细目id, a.序号, a.记录状态;

  --按序号排序是为了产生新记录时,填写同一收费细目的不同收入项目的价格父号

  --该游标用于处理费用相关汇总表
  Cursor c_Money
  (
    v_Start 住院费用记录.序号%Type,
    v_End   住院费用记录.序号%Type
  ) Is
    Select 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(Nvl(应收金额, 0)) As 应收金额, Sum(Nvl(实收金额, 0)) As 实收金额
    From 住院费用记录
    Where 记录性质 = 2 And 记录状态 = 1 And NO = No_In And 序号 Between v_Start And v_End
    Group By 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id;

  --系统参数指定执行后需要自动审核的划价费用：用于非药医嘱，包含对应的药品及卫材费用
  Cursor c_Verify
  (
    v_Start 住院费用记录.序号%Type,
    v_End   住院费用记录.序号%Type
  ) Is
    Select NO, 序号
    From 住院费用记录
    Where 记录性质 = 2 And 记录状态 = 0 And NO = No_In And 价格父号 Is Null And 序号 Between v_Start And v_End;

  Cursor c_Compound
  (
    相关id_In       病人医嘱记录.相关id%Type,
    执行终止时间_In 病人医嘱记录.执行终止时间%Type,
    配药id_In       输液配药记录.Id%Type
  ) Is
    Select b.费用id, b.药品id As 收费细目id, Sum(a.数量) As 数量, c.住院包装, c.住院单位, d.名称, e.病人病区id, e.操作状态, e.Id As 配药id, f.No,
           Nvl(f.价格父号, f.序号) As 序号, f.记录状态 As 记录状态, f.执行状态 As 执行标志
    From 输液配药内容 A, 药品收发记录 B, 药品规格 C, 收费项目目录 D, 输液配药记录 E, 住院费用记录 F
    Where a.收发id = b.Id And b.药品id = c.药品id And c.药品id = d.Id And e.Id = a.记录id And f.No = b.No And f.Id = b.费用id And
          e.医嘱id = 相关id_In And e.执行时间 > 执行终止时间_In And e.Id = 配药id_In
    Group By b.费用id, b.药品id, c.住院包装, c.住院单位, d.名称, e.病人病区id, e.操作状态, e.Id, f.No, f.价格父号, f.序号, f.记录状态, f.执行状态;

  v_Dec      Number;
  v_First    Number;
  v_划价类别 Varchar2(255);

  v_诊疗类别 病人医嘱记录.诊疗类别%Type;
  v_单次用量 病人医嘱记录.单次用量%Type;
  v_跟踪在用 材料特性.跟踪在用%Type;

  v_费用序号 住院费用记录.序号%Type;
  v_收发序号 药品收发记录.序号%Type;
  v_费用id   住院费用记录.Id%Type;
  v_实收金额 住院费用记录.实收金额%Type;

  v_开始序号 住院费用记录.序号%Type;
  v_结束序号 住院费用记录.序号%Type;

  v_医嘱执行 病人医嘱发送.执行状态%Type;

  v_剂量系数 药品规格.剂量系数%Type;
  v_住院包装 药品规格.住院包装%Type;
  v_医嘱内容 病人医嘱记录.医嘱内容%Type;

  v_结帐参数       Zlparameters.参数值%Type;
  v_配液药销帐申请 Zlparameters.参数值%Type;
  v_结帐金额       住院费用记录.结帐金额%Type;

  v_收费细目id   住院费用记录.收费细目id%Type;
  v_剩余数量     住院费用记录.数次%Type;
  v_收回数量     住院费用记录.数次%Type;
  v_当前数量     住院费用记录.数次%Type;
  v_当前付数     住院费用记录.付数%Type;
  v_费用ids      Varchar2(4000);
  v_组id         病人医嘱记录.Id%Type;
  v_对照数量     病人医嘱计价.数量%Type;

  v_Delno    Varchar2(4000);
  v_Temp     Varchar2(4000);
  v_收费内容 Varchar2(4000);
  v_No       住院费用记录.No%Type;
  v_人员编号 住院费用记录.操作员编号%Type;
  v_人员姓名 住院费用记录.操作员姓名%Type;

  n_相关id       病人医嘱记录.相关id%Type;
  d_执行终止时间 病人医嘱记录.执行终止时间%Type;
  n_药品id       病人医嘱记录.收费细目id%Type;
  b_输液配药记录 Boolean;
  d_收回时间     病人医嘱记录.执行终止时间%Type;
  n_申请类别     病人费用销帐.申请类别%Type;
  n_Count        Number;

  v_Error Varchar2(255);
  Err_Custom Exception;

  Procedure 负数收发记录_Insert
  (
    费用id_In     Number,
    批次_In       药品收发记录.批次%Type,
    分批_In       药品规格.药房分批%Type,
    批号_In       药品收发记录.批号%Type,
    效期_In       药品收发记录.效期%Type,
    最大效期_In   药品规格.最大效期%Type,
    收发id_In     药品收发记录.Id%Type,
    病人id_In     住院费用记录.病人id%Type,
    主页id_In     住院费用记录.主页id%Type,
    药品id_In     药品收发记录.药品id%Type,
    库房id_In     药品收发记录.库房id%Type,
    单据_In       药品收发记录.单据%Type,
    姓名_In       病人信息.姓名%Type,
    对方部门id_In 药品收发记录.对方部门id%Type,
    收费类别_In   住院费用记录.收费类别%Type,
    划价类别_In   Varchar
  ) Is
    v_批次   药品收发记录.批次%Type;
    v_效期   药品收发记录.效期%Type;
    v_批号   药品收发记录.批号%Type;
    v_优先级 身份.优先级%Type;
  Begin
    --确定批次
    If Nvl(批次_In, 0) <> 0 And 分批_In = 0 Then
      --原分批,现不分批
      v_批次 := Null;
      v_批号 := 批号_In;
      v_效期 := 效期_In;
    Elsif Nvl(批次_In, 0) = 0 And 分批_In = 1 Then
      --原不分批,现分批
      Select 药品收发记录_Id.Nextval Into v_批次 From Dual;
      Select To_Char(Sysdate, 'YYYYMMDD') Into v_批号 From Dual;
      If 最大效期_In Is Not Null Then
        v_效期 := Trunc(Sysdate + 最大效期_In * 30);
      Else
        v_效期 := Null;
      End If;
    Else
      v_批次 := 批次_In;
      v_批号 := 批号_In;
      v_效期 := 效期_In;
    End If;

    Insert Into 药品收发记录
      (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 零售价, 零售金额, 摘要, 填制人, 填制日期,
       费用id, 单量, 频次, 用法, 供药单位id, 生产日期, 批准文号, 灭菌效期)
      Select 药品收发记录_Id.Nextval, 1, 单据, No_In, v_收发序号, 库房id, 对方部门id, 入出类别id, -1, 药品id, v_批次, 产地, v_批号, v_效期, v_当前付数,
             -1 * v_当前数量, -1 * v_当前数量, 零售价, Round(-1 * v_当前付数 * v_当前数量 * 零售价, v_Dec), '超期发送收回', v_人员姓名, 收回时间_In, 费用id_In,
             单量, 频次, 用法, 供药单位id, 生产日期, 批准文号, 灭菌效期
      From 药品收发记录
      Where ID = 收发id_In;

    --药品库存
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) - (-1 * v_当前付数 * v_当前数量)
    Where 库房id = 库房id_In And 药品id = 药品id_In And Nvl(批次, 0) = Nvl(v_批次, 0) And 性质 = 1;
    If Sql%RowCount = 0 Then
      Insert Into 药品库存
        (库房id, 药品id, 性质, 可用数量, 批次, 效期)
      Values
        (库房id_In, 药品id_In, 1, v_当前付数 * v_当前数量, v_批次, v_效期);
    End If;

    Zl_药品库存_可用数量异常处理(库房id_In, 药品id_In, v_批次);

    --未发药品记录
    Update 未发药品记录
    Set 病人id = 病人id_In, 主页id = 主页id_In, 姓名 = 姓名_In
    Where 单据 = 单据_In And NO = No_In And 库房id + 0 = 库房id_In;

    If Sql%RowCount = 0 Then
      --取身份优先级
      Begin
        Select b.优先级 Into v_优先级 From 病人信息 A, 身份 B Where a.身份 = b.名称(+) And a.病人id = 病人id_In;
      Exception
        When Others Then
          Null;
      End;

      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 填制日期, 已收费, 打印状态)
      Values
        (单据_In, No_In, 病人id_In, 主页id_In, 姓名_In, v_优先级, 对方部门id_In, 库房id_In, 收回时间_In,
         Decode(Nvl(Instr(划价类别_In, Decode(收费类别_In, '4', '4', '5')), 0), 0, 1, 0), 0);
    End If;

    v_收发序号 := v_收发序号 + 1;
  End;
Begin
  --取操作员信息(部门ID,部门名称;人员ID,人员编号,人员姓名)
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
  --检查是否是输液配液记录，并是否已经锁定
  Select 医嘱内容 Into v_医嘱内容 From 病人医嘱记录 Where ID = 医嘱id_In;
  Select Count(1)
  Into n_Count
  From 输液配药记录 A, 病人医嘱记录 B
  Where a.医嘱id = b.Id And 医嘱id = 医嘱id_In And a.执行时间 > b.执行终止时间 And a.是否锁定 = 1;

  If n_Count > 0 Then
    v_Error := '医嘱"' || v_医嘱内容 || '"是输液药品，已经被输液配置中心锁定，不能超期收回。';
    Raise Err_Custom;
  End If;

  If Nvl(收回量_In, 0) > 0 Then
    --判断是否是输液配药药品(输液配制中心药品统一走销帐申请)
    b_输液配药记录 := False;
    Select a.相关id, a.执行终止时间, Max(b.收费细目id)
    Into n_相关id, d_执行终止时间, n_药品id
    From 病人医嘱记录 A, 住院费用记录 B
    Where a.Id = 医嘱id_In And a.Id = b.医嘱序号(+)
    Group By a.相关id, a.执行终止时间;
    If n_相关id Is Not Null Then
      If d_执行终止时间 Is Not Null Then
        d_收回时间       := 收回时间_In;
        v_配液药销帐申请 := zl_GetSysParameter('配液输液单配药后允许销帐申请', 1345);
        Select Count(1) Into n_Count From 输液配药记录 E Where e.医嘱id = n_相关id And e.执行时间 > d_执行终止时间;
        If n_Count > 0 Then
          b_输液配药记录 := True;
          For X In (Select e.Id As 配药id, e.操作状态
                    From 输液配药记录 E
                    Where e.医嘱id = n_相关id And e.执行时间 > d_执行终止时间 And Nvl(e.操作状态, 0) In (1, 2, 3, 4, 5, 6, 7, 8)) Loop
            If Not (x.操作状态 In (4, 5, 6, 7, 8) And Nvl(v_配液药销帐申请, '0') = '0') Then
              For r_Compound In c_Compound(n_相关id, d_执行终止时间, x.配药id) Loop
                If x.操作状态 = 1 Then
                  n_申请类别 := 0;
                Else
                  n_申请类别 := 1;
                End If;
                Select Count(1)
                Into n_Count
                From 病人费用销帐
                Where 费用id = r_Compound.费用id And 收费细目id = r_Compound.收费细目id And
                      申请时间 =
                      (Select Max(操作时间) From 输液配药状态 A Where a.配药id = r_Compound.配药id And a.操作类型 = 9);
                If n_Count = 0 Then
                  Zl_病人费用销帐_Insert(r_Compound.费用id, r_Compound.收费细目id, r_Compound.病人病区id, r_Compound.数量, v_人员姓名, d_收回时间,
                                   n_申请类别, Null, r_Compound.配药id);
                  If x.操作状态 = 1 Then
                    --未发药的，自动审核。
                    Zl_病人费用销帐_Audit(r_Compound.费用id, d_收回时间, v_人员姓名, d_收回时间, 1, 1, n_申请类别);
                    Zl_住院记帐记录_Delete(r_Compound.No, r_Compound.序号 || ':' || r_Compound.数量 || ':' || r_Compound.配药id,
                                     v_人员编号, v_人员姓名, 2, Null, Null, d_收回时间);
                  End If;
                End If;
              End Loop;
              --由于不同批次（执行时间）申请时，申请时间和费用ID有唯一约束，所以同时销帐多个批次时，依次加一秒
              d_收回时间 := d_收回时间 + 1 / 24 / 60 / 60;
            End If;
          End Loop;
        End If;
      End If;
    End If;
    --a.销帐申请收回模式
    --输液配药记录单独进行销帐
    If b_输液配药记录 = False Then
      If No_In Is Null Then
        v_结帐参数 := zl_GetSysParameter(23);
        --根据收回数量对照原始费用进行分摊申请
        For r_Detail In c_Detail Loop
          --确定该收费细目ID的收回总数量
          If Nvl(v_收费细目id, 0) <> r_Detail.收费细目id And (r_Detail.诊疗类别 Not In ('5', '6', '7') Or Nvl(v_收费细目id, 0) = 0) Then
            --数量未分摊完成
            If v_收费细目id Is Not Null And v_收回数量 > 0 Then
              v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能存在手工销帐或已结帐的费用，或相应的划价单已被删除。';
              Raise Err_Custom;
            End If;
            --药品收回总量是以最后发送规格为准计算的，以此计算出收回售价数量
            Begin
              Select 剂量系数, 住院包装 Into v_剂量系数, v_住院包装 From 药品规格 Where 药品id = r_Detail.收费细目id;
            Exception
              When Others Then
                Null;
            End;
            --计算收回数量，如果收费方式不是0正常收取，则使用特殊的方法进行计算
            If r_Detail.收费方式 = 0 Then
              If r_Detail.诊疗类别 = '7' Then
                --中药配方药品：付数*单量
                v_收回数量 := Round(收回量_In * r_Detail.单次用量 / Nvl(v_剂量系数, 1), 5);
              Else
                If r_Detail.诊疗类别 Not In ('5', '6') Then
                  Select Nvl(Max(数量), 1)
                  Into v_对照数量
                  From 病人医嘱计价
                  Where 医嘱id = 医嘱id_In And 收费细目id = r_Detail.收费细目id;
                Else
                  v_对照数量 := 1;
                End If;
                v_收回数量 := Round(收回量_In * Nvl(v_住院包装, 1), 5) * v_对照数量;
              End If;
            Else
              Select Nvl(Sum(数量), 0)
              Into v_收回数量
              From 医嘱执行计价
              Where 医嘱id = 医嘱id_In And 收费细目id = r_Detail.收费细目id And
                    要求时间 > Nvl(上次时间_In, To_Date('1900-01-01', 'yyyy-MM-dd'));

              v_收回数量 := Round(v_收回数量, 5);

            End If;
            v_医嘱内容 := r_Detail.医嘱内容;
          End If;

          --该收费细目的每个费用明细分摊收回
          If v_收回数量 > 0 Then
            --检查对应费用是否已结帐，当禁止时
            v_结帐金额 := 0;
            If v_结帐参数 = '2' And r_Detail.记录状态 <> 0 Then
              Select Sum(结帐金额)
              Into v_结帐金额
              From 住院费用记录
              Where NO = r_Detail.No And 记录性质 In (2, 12) And Nvl(价格父号, 序号) = r_Detail.序号;
            End If;

            If Nvl(v_结帐金额, 0) = 0 Then
              If r_Detail.收费类别 In ('5', '6', '7') Or r_Detail.收费类别 = '4' And r_Detail.跟踪在用 = 1 Then
                --药品和跟踪在用的卫材
                If r_Detail.执行标志 = 0 Then
                  v_剩余数量 := r_Detail.未执行量;
                Elsif r_Detail.执行标志 = 1 Then
                  v_剩余数量 := r_Detail.已执行量;
                End If;
              Else
                --普通费用
                v_剩余数量 := r_Detail.剩余数量;
              End If;
              If v_收回数量 > v_剩余数量 Then
                v_当前数量 := v_剩余数量;
              Else
                v_当前数量 := v_收回数量;
              End If;
              v_收回数量 := v_收回数量 - v_当前数量;
              --系统参数决定执行后是否审核划价单，所以，已执行的仍然可能是划价单
              If r_Detail.执行标志 = 0 And r_Detail.记录状态 = 0 Then
                v_Delno := v_Delno || '|' || r_Detail.No || ',' || r_Detail.序号 || ':' || v_当前数量;
              Else
                If Not (r_Detail.收费类别 = '7' And r_Detail.执行标志 <> 0) Then
                  Zl_病人费用销帐_Insert(r_Detail.费用id, r_Detail.收费细目id, r_Detail.病人病区id, v_当前数量, v_人员姓名, 收回时间_In,
                                   r_Detail.执行标志);
                End If;
              End If;
              v_费用ids := v_费用ids || ',' || r_Detail.费用id;
            End If;
          End If;
          v_收费细目id := r_Detail.收费细目id;
        End Loop;

        --数量未分摊完成
        If v_收回数量 > 0 Then
          v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能存在手工销帐或已结帐的费用，或相应的划价单已被删除。';
          Raise Err_Custom;
        End If;
        --本科的销帐申请自动审核
        If zl_GetSysParameter('超期收回费用本科自动审核', 1254) = '1' And v_费用ids Is Not Null Then
          For r_Applay In c_Applay(Substr(v_费用ids, 2)) Loop
            Zl_病人费用销帐_Audit(r_Applay.费用id, r_Applay.申请时间, v_人员姓名, 收回时间_In, 1, 1, r_Applay.申请类别);
            v_Delno := v_Delno || '|' || r_Applay.No || ',' || r_Applay.序号 || ':' || r_Applay.数量;
          End Loop;
        End If;
      Else
        ---b.负数收回模式-------------------------------------------------------------------------------------------------------
        --如果全是划价单，就不用产生负数冲销单据
        If No_In = '调整划价单' Then
          --未审核的划价单，先进行修改或删除，可能多次发送为不同的NO,为了计算每次的收回量，需要按收费细目ID排序
          For r_Price In (Select c.诊疗类别, b.No, b.序号, b.收费细目id, Nvl(b.付数, 1) * b.数次 As 剩余数量, c.单次用量, d.剂量系数, d.住院包装,
                                 c.医嘱内容, Nvl(e.收费方式, 0) As 收费方式
                          From 病人医嘱发送 A, 住院费用记录 B, 病人医嘱记录 C, 药品规格 D, 病人医嘱计价 E
                          Where a.医嘱id = 医嘱id_In And a.No = b.No And a.记录性质 = b.记录性质 And a.医嘱id = b.医嘱序号 And
                                b.价格父号 Is Null And b.收费细目id = d.药品id(+) And b.记录状态 = 0 And c.Id = a.医嘱id And
                                b.医嘱序号 = e.医嘱id(+) And b.收费细目id = e.收费细目id(+)
                          Order By 收费细目id, NO Desc) Loop
            If Nvl(v_收费细目id, 0) <> r_Price.收费细目id Then
              --数量未分摊完成
              If v_收费细目id Is Not Null And v_收回数量 > 0 Then
                v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能相关划价单已被删除或审核。';
                Raise Err_Custom;
              End If;
              --计算收回数量，如果收费方式不是0正常收取，则使用特殊的方法进行计算
              If r_Price.收费方式 = 0 Then
                If r_Price.诊疗类别 = '7' Then
                  --中药配方药品：付数*单量
                  v_收回数量 := Round(收回量_In * r_Price.单次用量 / Nvl(r_Price.剂量系数, 1), 5);
                Else
                  If r_Price.诊疗类别 Not In ('5', '6') Then
                    Select Nvl(Max(数量), 1)
                    Into v_对照数量
                    From 病人医嘱计价
                    Where 医嘱id = 医嘱id_In And 收费细目id = r_Price.收费细目id;
                  Else
                    v_对照数量 := 1;
                  End If;
                  v_收回数量 := Round(收回量_In * Nvl(r_Price.住院包装, 1), 5) * v_对照数量;
                End If;
              Else
                Select Nvl(Sum(数量), 0)
                Into v_收回数量
                From 医嘱执行计价
                Where 医嘱id = 医嘱id_In And 收费细目id = r_Price.收费细目id And
                      要求时间 > Nvl(上次时间_In, To_Date('1900-01-01', 'yyyy-MM-dd'));

                v_收回数量 := Round(v_收回数量, 5);
              End If;
              v_医嘱内容 := r_Price.医嘱内容;
            End If;
            If v_收回数量 > 0 Then
              If v_收回数量 > r_Price.剩余数量 Then
                v_当前数量 := r_Price.剩余数量;
              Else
                v_当前数量 := v_收回数量;
              End If;
              v_收回数量 := v_收回数量 - v_当前数量;
              v_Delno    := v_Delno || '|' || r_Price.No || ',' || r_Price.序号 || ':' || v_当前数量;
            End If;
            v_收费细目id := r_Price.收费细目id;
          End Loop;
          --数量未分摊完成
          If v_收回数量 > 0 Then
            v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能相关划价单已被删除或审核。';
            Raise Err_Custom;
          End If;
        Else
          --负数冲销，可能存在划价单与记帐单混合的情况
          --金额小数位数
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;
          --生成划价单系统参数
          Select zl_GetSysParameter(80) Into v_划价类别 From Dual;
          v_开始序号 := Null;
          v_结束序号 := Null;

          Select a.诊疗类别, a.单次用量, b.跟踪在用
          Into v_诊疗类别, v_单次用量, v_跟踪在用
          From 病人医嘱记录 A, 材料特性 B
          Where ID = 医嘱id_In And a.收费细目id = b.材料id(+);

          If v_诊疗类别 In ('5', '6', '7') Or (v_诊疗类别 = '4' And Nvl(v_跟踪在用, 0) = 1) Then
            --药品、卫材
            -----------------------------------------------------------------------------------------------------
            v_收回数量 := Null;
            Select Nvl(Max(序号), 0) + 1
            Into v_收发序号
            From 药品收发记录
            Where 单据 In (9, 10, 25, 26) And 记录状态 = 1 And NO = No_In;
            Select Nvl(Max(序号), 0) + 1
            Into v_费用序号
            From 住院费用记录
            Where 记录性质 = 2 And 记录状态 In (0, 1) And NO = No_In;

            --一条医嘱的药品只有一行，这里的循环是为了处理多次发送的情况，分批药品在界面已禁用负数收回
            For r_Drug In c_Drug Loop
              --初始化要收回的总数量(零售数量)
              v_First := 0;
              If v_收回数量 Is Null Then
                If v_诊疗类别 = '7' Then
                  v_收回数量 := Round(收回量_In * v_单次用量 / r_Drug.剂量系数, 5);
                Else
                  If v_诊疗类别 Not In ('5', '6') Then
                    Select Nvl(Max(数量), 1)
                    Into v_对照数量
                    From 病人医嘱计价
                    Where 医嘱id = 医嘱id_In And 收费细目id = r_Drug.收费细目id;
                  Else
                    v_对照数量 := 1;
                  End If;
                  v_收回数量 := Round(收回量_In * r_Drug.住院包装, 5) * v_对照数量;
                End If;
                v_First := 1;
              End If;

              --如果第一次数量就足够，则按付数处理，否则付数不好处理
              If v_收回数量 > r_Drug.数量 Then
                v_当前付数 := 1;
                v_当前数量 := r_Drug.数量;
                v_收回数量 := v_收回数量 - r_Drug.数量;
              Else
                If v_First = 1 And v_诊疗类别 = '7' Then
                  v_当前付数 := 收回量_In;
                  v_当前数量 := Round(v_单次用量 / r_Drug.剂量系数, 5);
                Else
                  v_当前付数 := 1;
                  v_当前数量 := v_收回数量;
                End If;
                v_收回数量 := 0;
              End If;

              If r_Drug.记录状态 = 0 Then
                v_Delno := v_Delno || '|' || r_Drug.No || ',' || r_Drug.序号 || ':' || v_当前数量 * v_当前付数;
              Else
                If Not (v_诊疗类别 = '7' And r_Drug.执行标志 <> 0) Then

                  Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
                  负数收发记录_Insert(v_费用id, r_Drug.批次, r_Drug.分批, r_Drug.批号, r_Drug.效期, r_Drug.最大效期, r_Drug.收发id,
                                r_Drug.病人id, r_Drug.主页id, r_Drug.药品id, r_Drug.库房id, r_Drug.单据, r_Drug.姓名, r_Drug.对方部门id,
                                v_诊疗类别, v_划价类别);

                  --住院费用记录
                  -------------------------------------------------------------------------------------
                  --记录序号范围以处理汇总表
                  If v_开始序号 Is Null Then
                    v_开始序号 := v_费用序号;
                  End If;
                  v_结束序号 := v_费用序号;

                  Insert Into 住院费用记录
                    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id,
                     费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额,
                     统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 医嘱序号, 划价人, 操作员编号, 操作员姓名)
                    Select v_费用id, 2, No_In, Decode(Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0), 0, 1, 0),
                           v_费用序号, Null, Null, 多病人单, 2, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
                           收费细目id, 计算单位, 保险项目否, 保险大类id, v_当前付数, -1 * v_当前数量, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价,
                           Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec), Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec), Null, 1,
                           开单部门id, 开单人, 收回时间_In, 收回时间_In, 执行部门id, 0, 医嘱序号, v_人员姓名,
                           Decode(Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0), 0, v_人员编号, Null),
                           Decode(Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0), 0, v_人员姓名, Null)
                    From 住院费用记录
                    Where ID = r_Drug.费用id;

                  Select Zl_Actualmoney(费别, 收费细目id, 收入项目id, 应收金额, 数次, 执行部门id)
                  Into v_Temp
                  From 住院费用记录
                  Where ID = v_费用id;
                  v_实收金额 := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                  Update 住院费用记录 A Set 实收金额 = v_实收金额 Where ID = v_费用id;

                  v_费用序号 := v_费用序号 + 1;
                End If;
                If v_收回数量 <= 0 Then
                  Exit;
                End If;
              End If;
            End Loop;

            If v_收回数量 <> 0 Then
              --没有收回所有数量,收发记录本身有问题(如记录不全或数量为负)
              Null;
            End If;
          Else
            --其它非药医嘱(包括给药途径，及绑定的卫材等)
            -----------------------------------------------------------------------------------------------------
            Select Nvl(Max(序号), 0) + 1
            Into v_收发序号
            From 药品收发记录
            Where 单据 In (9, 10, 25, 26) And 记录状态 = 1 And NO = No_In;
            --取费用序号
            Select Nvl(Max(序号), 0) + 1
            Into v_费用序号
            From 住院费用记录
            Where 记录性质 = 2 And 记录状态 In (0, 1) And NO = No_In;

            For r_Other In c_Other Loop
              If Nvl(v_收费内容, '0') <> r_Other.收费细目id || ',' || r_Other.序号 Then
                --根据最近一次发送的费用记录，按需要收回的数量全部收回
                --计算收回数量，如果收费方式不是0正常收取，则使用特殊的方法进行计算
                If r_Other.收费方式 = 0 Then
                  v_收回数量 := 收回量_In * Nvl(r_Other.对照数量, 1);
                Else
                  Select Nvl(Sum(数量), 0)
                  Into v_收回数量
                  From 医嘱执行计价
                  Where 医嘱id = 医嘱id_In And 收费细目id = r_Other.收费细目id And
                        要求时间 > Nvl(上次时间_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
                End If;
              End If;

              If v_收回数量 > 0 Then
                If r_Other.记录状态 = 0 Then
                  If v_收回数量 > r_Other.剩余数量 Then
                    v_当前数量 := r_Other.剩余数量;
                  Else
                    v_当前数量 := v_收回数量;
                  End If;
                Else
                  v_当前数量 := v_收回数量;
                End If;
                v_收回数量 := v_收回数量 - v_当前数量;
                v_当前付数 := 1;

                If r_Other.记录状态 = 0 Then
                  v_Delno := v_Delno || '|' || r_Other.No || ',' || r_Other.序号 || ':' || v_当前数量;
                Else
                  --记录序号范围以处理汇总表
                  If v_开始序号 Is Null Then
                    v_开始序号 := v_费用序号;
                  End If;
                  v_结束序号 := v_费用序号;

                  --住院费用记录:按理如果收回量大于了上次发送量,则不正确
                  Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
                  If r_Other.收费类别 In ('4', '5', '6', '7') Then
                    For r_Otherdrug In (Select a.病人id, a.主页id, d.姓名, Nvl(Nvl(x.剂量系数, y.换算系数), 1) As 剂量系数,
                                                 Nvl(x.住院包装, 1) As 住院包装, Nvl(x.最大效期, y.最大效期) As 最大效期,
                                                 Nvl(b.付数, 1) * b.实际数量 As 数量, b.Id As 收发id, b.单据, b.药品id, b.对方部门id,
                                                 b.库房id, b.费用id, Nvl(Nvl(x.药房分批, y.在用分批), 0) As 分批, b.批次, b.批号, b.效期,
                                                 a.记录状态, a.No, a.序号, a.收费细目id
                                          From 住院费用记录 A, 药品收发记录 B, 病人信息 D, 药品规格 X, 材料特性 Y
                                          Where a.Id = r_Other.费用id And a.记录状态 In (0, 1, 3) And a.No = b.No And
                                                a.Id = b.费用id + 0 And b.单据 In (9, 10, 25, 26) And
                                                (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And a.病人id = d.病人id And
                                                b.药品id = x.药品id(+) And b.药品id = y.材料id(+)
                                        Order By a.记录状态, b.No Desc, b.Id Desc) Loop
                      负数收发记录_Insert(v_费用id, r_Otherdrug.批次, r_Otherdrug.分批, r_Otherdrug.批号, r_Otherdrug.效期,
                                    r_Otherdrug.最大效期, r_Otherdrug.收发id, r_Otherdrug.病人id, r_Otherdrug.主页id,
                                    r_Otherdrug.药品id, r_Otherdrug.库房id, r_Otherdrug.单据, r_Otherdrug.姓名,
                                    r_Otherdrug.对方部门id, r_Other.收费类别, v_划价类别);
                    End Loop;
                  End If;
                  --医嘱已执行，收回的费用也填为已执行：不包含药品和跟踪在用的卫材，因为实际发放表示执行
                  Insert Into 住院费用记录
                    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id,
                     费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额,
                     统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行时间, 执行人, 医嘱序号, 划价人, 操作员编号, 操作员姓名)
                    Select v_费用id, 2, No_In, Decode(Nvl(Instr(v_划价类别, r_Other.收费类别), 0), 0, 1, 0), v_费用序号, Null,
                           Decode(a.价格父号, Null, Null, v_费用序号 + a.价格父号 - a.序号), a.多病人单, 2, a.病人id, a.主页id, a.标识号, a.姓名,
                           a.性别, a.年龄, a.床号, a.病人病区id, a.病人科室id, a.费别, a.收费类别, a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, 1,
                           -1 * v_当前数量, a.加班标志, a.附加标志, a.婴儿费, a.收入项目id, a.收据费目, a.标准单价,
                           Round(-1 * v_当前数量 * a.标准单价, v_Dec), Round(-1 * v_当前数量 * a.标准单价, v_Dec), Null, 1, a.开单部门id,
                           a.开单人, 收回时间_In, 收回时间_In, a.执行部门id,
                           Decode(r_Other.执行状态, 1,
                                   Decode(a.收费类别, '4', Decode(b.跟踪在用, 1, 0, 1), Decode(Instr(',5,6,7,', a.收费类别), 0, 1, 0)),
                                   0),
                           Decode(r_Other.执行状态, 1,
                                   Decode(a.收费类别, '4', Decode(b.跟踪在用, 1, Null, 收回时间_In),
                                           Decode(Instr(',5,6,7,', a.收费类别), 0, 收回时间_In, Null)), Null),
                           Decode(r_Other.执行状态, 1,
                                   Decode(a.收费类别, '4', Decode(b.跟踪在用, 1, Null, v_人员姓名),
                                           Decode(Instr(',5,6,7,', a.收费类别), 0, v_人员姓名, Null)), Null), a.医嘱序号, v_人员姓名,
                           Decode(Nvl(Instr(v_划价类别, r_Other.收费类别), 0), 0, v_人员编号, Null),
                           Decode(Nvl(Instr(v_划价类别, r_Other.收费类别), 0), 0, v_人员姓名, Null)
                    From 住院费用记录 A, 材料特性 B
                    Where a.Id = r_Other.费用id And a.收费细目id = b.材料id(+);

                  Select Zl_Actualmoney(费别, 收费细目id, 收入项目id, 应收金额, 数次, 执行部门id)
                  Into v_Temp
                  From 住院费用记录
                  Where ID = v_费用id;
                  v_实收金额 := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                  Update 住院费用记录 A Set 实收金额 = v_实收金额 Where ID = v_费用id;

                  v_费用序号 := v_费用序号 + 1;
                  v_医嘱执行 := r_Other.执行状态; --多个收费项目的执行状态是一样的
                End If;

                v_收费内容 := r_Other.收费细目id || ',' || r_Other.序号;
              End If;
            End Loop;

            --如果医嘱已执行，则按系统参数执行后自动审核费用：包含已执行医嘱对应的药品和卫材费用。
            -----------------------------------------------------------------------------------------------------
            If Nvl(v_医嘱执行, 0) = 1 And v_开始序号 Is Not Null And v_结束序号 Is Not Null Then
              If zl_GetSysParameter(81) = '1' Then
                For r_Verify In c_Verify(v_开始序号, v_结束序号) Loop
                  Zl_住院记帐记录_Verify(r_Verify.No, v_人员编号, v_人员姓名, r_Verify.序号, Null, 收回时间_In);
                End Loop;
              End If;
            End If;
          End If;

          --处理费用汇总表
          -----------------------------------------------------------------------------------------------------
          If v_开始序号 Is Not Null And v_结束序号 Is Not Null Then
            --最后统一处理费用相关汇总表
            For r_Money In c_Money(v_开始序号, v_结束序号) Loop
              --病人余额
              Update 病人余额
              Set 费用余额 = Nvl(费用余额, 0) + r_Money.实收金额
              Where 病人id = r_Money.病人id And 性质 = 1 And 类型 = 2;

              If Sql%RowCount = 0 Then
                Insert Into 病人余额
                  (病人id, 性质, 类型, 费用余额, 预交余额)
                Values
                  (r_Money.病人id, 1, 2, r_Money.实收金额, 0);
              End If;

              --病人未结费用
              Update 病人未结费用
              Set 金额 = Nvl(金额, 0) + r_Money.实收金额
              Where 病人id = r_Money.病人id And 主页id = r_Money.主页id And Nvl(病人病区id, 0) = Nvl(r_Money.病人病区id, 0) And
                    Nvl(病人科室id, 0) = Nvl(r_Money.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Money.开单部门id, 0) And
                    Nvl(执行部门id, 0) = Nvl(r_Money.执行部门id, 0) And 收入项目id + 0 = r_Money.收入项目id And 来源途径 + 0 = 2;

              If Sql%RowCount = 0 Then
                Insert Into 病人未结费用
                  (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
                Values
                  (r_Money.病人id, r_Money.主页id, r_Money.病人病区id, r_Money.病人科室id, r_Money.开单部门id, r_Money.执行部门id,
                   r_Money.收入项目id, 2, r_Money.实收金额);
              End If;
            End Loop;
          End If;
        End If;
      End If;
    End If;
  End If;

  --过程Zl_住院记帐记录_Delete，不支持每次删除一行的循环处理（序号重整），必须把一个单据要删除的序号一次性传入
  If Not v_Delno Is Null Then
    v_Temp := '';
    v_No   := '';
    For r_Price In (Select /*+ rule*/
                     C1 As NO, C2 As 序号数量
                    From Table(f_Str2list2(Substr(v_Delno, 2), '|', ','))
                    Order By NO) Loop
      If v_No Is Not Null And v_No <> r_Price.No Then
        Zl_住院记帐记录_Delete(v_No, v_Temp, v_人员编号, v_人员姓名, 2);
        v_No := '';
      End If;
      If v_No Is Null Then
        v_No   := r_Price.No;
        v_Temp := r_Price.序号数量;
      Else
        v_Temp := v_Temp || ',' || r_Price.序号数量;
      End If;
    End Loop;
    If Not v_No Is Null Then
      Zl_住院记帐记录_Delete(v_No, v_Temp, v_人员编号, v_人员姓名, 2);
    End If;
  End If;

  --处理医嘱的上次执行时间:给药途径等可能因为未发送而没调用收回过程。
  -----------------------------------------------------------------------------------------------------
  Select Nvl(相关id, ID) Into v_组id From 病人医嘱记录 Where ID = 医嘱id_In;
  Update 病人医嘱记录 Set 上次执行时间 = 上次时间_In Where ID = v_组id Or 相关id = v_组id;

  --删除医嘱执行时间
  If 上次时间_In Is Null Then
    --全部收回
    Delete From 医嘱执行时间 Where 医嘱id = v_组id;
    Delete From 医嘱执行计价 Where 医嘱id = 医嘱id_In;
  Else
    --可能收回多次发送的数据
    Delete From 医嘱执行时间 Where 医嘱id = v_组id And 要求时间 > 上次时间_In;
    Delete From 医嘱执行计价 Where 医嘱id = 医嘱id_In And 要求时间 > 上次时间_In;
  End If;
  --处理输液配液记录的批次问题，每个医嘱都进行调用，在过程里面只处理了输液配液的医嘱
  Zl_输液配药记录_批次调整(医嘱id_In);

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_收回;
/

--126851:李小东,2018-11-28,标本拒收后改变费用执行状态为0
Create Or Replace Procedure Zl_检验申请拒收_Update
(
  医嘱ids_In  In Varchar2, --多个医嘱ID用逗号分隔
  执行说明_In In 病人医嘱发送.执行说明%Type
) Is
  v_费用性质 病人医嘱发送.记录性质%Type;

  Cursor c_Samplequest Is
    Select Distinct ID As 医嘱id, 病人来源
    From 病人医嘱记录 A
    Where a.Id In (Select * From Table(Cast(f_Num2list(医嘱ids_In) As Zltools.t_Numlist)));
Begin
  --处理医嘱执行状态
  Update 病人医嘱发送
  Set 执行状态 = 2, 执行说明 = 执行说明_In, 采样人 = Null, 采样时间 = Null, 送检人 = Null, 标本送出时间 = Null, 标本发送批号 = Null, 接收人 = Null,
      接收时间 = Null
  Where 医嘱id In (Select * From Table(Cast(f_Num2list(医嘱ids_In) As Zltools.t_Numlist)));

  --处理费用执行状态
  For r_Samplequest In c_Samplequest Loop
    If r_Samplequest.病人来源 = 2 Then
      Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
      Into v_费用性质
      From 病人医嘱发送
      Where 医嘱id = r_Samplequest.医嘱id;
    Else
      v_费用性质 := 1;
    End If;
  
    If v_费用性质 = 2 Then
      Update 住院费用记录
      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
      Where 收费类别 Not In ('5', '6', '7') And
            (医嘱序号, 记录性质, NO) In
            (Select 医嘱id, 记录性质, NO
             From 病人医嘱附费
             Where 医嘱id = r_Samplequest.医嘱id
             Union All
             Select 医嘱id, 记录性质, NO
             From 病人医嘱发送
             Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null));
    Else
      Update 门诊费用记录
      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
      Where 收费类别 Not In ('5', '6', '7') And
            (医嘱序号, 记录性质, NO) In
            (Select 医嘱id, 记录性质, NO
             From 病人医嘱附费
             Where 医嘱id = r_Samplequest.医嘱id
             Union All
             Select 医嘱id, 记录性质, NO
             From 病人医嘱发送
             Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null));
    End If;
  
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验申请拒收_Update;
/

--134222:焦博,2018-11-21,调整Oracle过程Zl_Third_Getregistalter,分段创建返回的XML字符串
Create Or Replace Procedure Zl_Third_Getregistalter
(
  Xml_In  Xmltype,
  Xml_Out Out Xmltype
) Is
  -----------------------------------------------
  --功能：获取当天操作的停换诊安排
  --入参：XML_IN
  --<IN>
  --  <JSKLB>结算卡类别</JSKLB>
  --  <RQ>日期</RQ>
  --</IN>
  --出参:XML_OUT
  --<OUTPUT>
  --  <TZLISTS>          //停诊列表
  --    <ITEM>
  --      <HM>号码</HM>
  --      <YSID>医生ID</YSID>
  --      <YS>医生姓名</YS>
  --      <KSSJ>停诊开始时间</KSSJ>
  --      <JSSJ>停诊结束时间</JSSJ>
  --      <BRLIST>
  --        <INFO>
  --          <YYNO>预约单据号</YYNO>
  --          <BRID>病人ID</BRID>
  --          <YYSJ>预约时间</YYSJ>
  --          <CZSJ>操作时间</CZSJ>
  --          <YYKS>预约科室</YYKS>
  --          <GHLX>号类</GHLX>
  --          <YSXM>医生姓名</YSXM>
  --        </INFO>
  --      </BRLIST>
  --    </ITEM>
  --  </TZLISTS>
  --  <HZLISTS>          //换诊列表
  --    <ITEM>
  --      <BRID>病人ID</BRID>
  --      <YYSJ>预约的操作时间</YYSJ>
  --      <YSJ>原预约时间</YSJ>
  --      <YHM>原号码</YHM>
  --      <YYS>原医生</YYS>
  --      <YZC>原医生的职称</YZC>
  --      <XSJ>现预约时间</XSJ>
  --      <XHM>现号码</XHM>
  --      <XYS>现医生</XYS>
  --      <XZC>现医生的职称</XZC>
  --    </ITEM>
  --  </HZLIST>
  --</OUTPUT>
  -----------------------------------------------------

  d_Date     Date;
  v_Jsklb    Varchar2(100);
  n_卡类别id 医疗卡类别.Id%Type;
  n_Cnt      Number(3);
  v_Temp     Clob;
  v_Brinfo   Varchar2(4000);
  d_启用时间 Date;
  v_Para     Varchar2(2000);
  n_Exists   Number(3);
  n_挂号模式 Number(3);
  x_Templet  Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/JSKLB') Into v_Jsklb From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'yyyy-mm-dd')
  Into d_Date
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select b.Id Into n_卡类别id From 医疗卡类别 B Where b.名称 = v_Jsklb And Rownum < 2;
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  If n_挂号模式 = 1 And Nvl(d_Date, Sysdate) > Nvl(d_启用时间, Sysdate - 30) Then
    --出诊表排班模式
    --获取停诊安排
    For r_停诊 In (Select a.Id As 记录id, b.号码, a.医生id, a.医生姓名, a.停诊开始时间, a.停诊终止时间
                 From 临床出诊记录 A, 临床出诊号源 B, 临床出诊停诊记录 C
                 Where a.Id = c.记录id And a.号源id = b.Id And a.停诊开始时间 Is Not Null And c.审批时间 Between d_Date And
                       d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || r_停诊.号码 || '</HM><YSID>' || r_停诊.医生id || '</YSID><YS>' || r_停诊.医生姓名 ||
                '</YS><KSSJ>' || r_停诊.停诊开始时间 || '</KSSJ><JSSJ>' || r_停诊.停诊终止时间 || '</JSSJ><BRLIST>';
      For r_停诊病人 In (Select a.记录性质, a.No, a.病人id, To_Char(a.发生时间, 'yyyy-mm-dd') As 发生时间,
                            To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.名称, d.号类, c.医生姓名 As 医生姓名
                     From 病人挂号记录 A, 部门表 B, 临床出诊记录 C, 临床出诊号源 D
                     Where a.执行部门id = b.Id And a.出诊记录id = c.Id And c.号源id = d.Id And 记录状态 = 1 And
                           发生时间 Between r_停诊.停诊开始时间 And r_停诊.停诊终止时间 And a.出诊记录id = r_停诊.记录id And Not Exists
                      (Select 1 From 就诊变动记录 Where 挂号单 = a.No)) Loop
        --停诊病人列表，不包含已经换诊和取消了的病人
        If r_停诊病人.记录性质 = 2 Then
          v_Brinfo := '<INFO><YYNO>' || r_停诊病人.No || '</YYNO><BRID>' || r_停诊病人.病人id || '</BRID><YYSJ>' || r_停诊病人.发生时间 ||
                      '</YYSJ><CZSJ>' || r_停诊病人.登记时间 || '</CZSJ>' || '<YYKS>' || r_停诊病人.名称 || '</YYKS><GHLX>' ||
                      r_停诊病人.号类 || '</GHLX><YSXM>' || r_停诊病人.医生姓名 || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        Else
          Begin
            Select 1
            Into n_Exists
            From 病人预交记录
            Where NO = r_停诊病人.No And 记录性质 = 4 And 卡类别id = n_卡类别id;
          Exception
            When Others Then
              n_Exists := 0;
          End;
          If n_Exists = 1 Then
            v_Brinfo := '<INFO><YYNO>' || r_停诊病人.No || '</YYNO><BRID>' || r_停诊病人.病人id || '</BRID><YYSJ>' || r_停诊病人.发生时间 ||
                        '</YYSJ><CZSJ>' || r_停诊病人.登记时间 || '</CZSJ>' || '<YYKS>' || r_停诊病人.名称 || '</YYKS><GHLX>' ||
                        r_停诊病人.号类 || '</GHLX><YSXM>' || r_停诊病人.医生姓名 || '</YSXM></INFO>';
            v_Temp   := v_Temp || v_Brinfo;
          End If;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    --换取换诊列表
    v_Temp := '';
    For r_换诊 In (Select d.记录性质, d.No, a.病人id, To_Char(d.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                        To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.原号码, a.原医生姓名, b.专业技术职务 As 原职务, a.现号码, a.现医生姓名,
                        c.专业技术职务 As 现职务
                 From 就诊变动记录 A, 人员表 B, 人员表 C, 病人挂号记录 D
                 Where a.登记时间 Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.原医生id = b.Id And a.现医生id = c.Id And
                       a.挂号单 = d.No) Loop
      --只返回该卡类别挂号的病人         
      If r_换诊.记录性质 = 2 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || r_换诊.病人id || '</BRID><YYSJ>' || r_换诊.登记时间 || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || r_换诊.预约时间 || '</YSJ><YHM>' || r_换诊.原号码 || '</YHM><YYS>' || r_换诊.原医生姓名 ||
                  '</YYS><YZC>' || r_换诊.原职务 || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || r_换诊.预约时间 || '</XSJ><XHM>' || r_换诊.现号码 || '</XHM><XYS>' || r_换诊.现医生姓名 ||
                  '</XYS><XZC>' || r_换诊.现职务 || '</XZC></ITEM>';
      Else
        Begin
          Select 1 Into n_Exists From 病人预交记录 Where NO = r_换诊.No And 记录性质 = 4 And 卡类别id = n_卡类别id;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists = 1 Then
          v_Temp := v_Temp || '<ITEM><BRID>' || r_换诊.病人id || '</BRID><YYSJ>' || r_换诊.登记时间 || '</YYSJ>';
          v_Temp := v_Temp || '<YSJ>' || r_换诊.预约时间 || '</YSJ><YHM>' || r_换诊.原号码 || '</YHM><YYS>' || r_换诊.原医生姓名 ||
                    '</YYS><YZC>' || r_换诊.原职务 || '</YZC>';
          v_Temp := v_Temp || '<XSJ>' || r_换诊.预约时间 || '</XSJ><XHM>' || r_换诊.现号码 || '</XHM><XYS>' || r_换诊.现医生姓名 ||
                    '</XYS><XZC>' || r_换诊.现职务 || '</XZC></ITEM>';
        End If;
      End If;
    End Loop;
    v_Temp := '<HZLISTS>' || v_Temp || '</HZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --计划排班模式
    --获取停诊安排
    For Rs In (Select b.号码, b.医生id, b.医生姓名, To_Char(a.开始停止时间, 'yyyy-mm-dd hh24:mi:ss') As 开始停止时间,
                      To_Char(a.结束停止时间, 'yyyy-mm-dd hh24:mi:ss') As 结束停止时间
               From 挂号安排停用状态 A, 挂号安排 B
               Where a.安排id = b.Id And a.制订日期 Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || Rs.号码 || '</HM><YSID>' || Rs.医生id || '</YSID><YS>' || Rs.医生姓名 ||
                '</YS><KSSJ>' || Rs.开始停止时间 || '</KSSJ><JSSJ>' || Rs.结束停止时间 || '</JSSJ><BRLIST>';
      ----2015/7/28
      For Rs_Br In (Select a.No, a.病人id, To_Char(a.发生时间, 'yyyy-mm-dd') As 发生时间,
                           To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.名称, c.号类, a.执行人 As 医生姓名
                    From 病人挂号记录 A, 部门表 B, 挂号安排 C
                    Where a.号别 = Rs.号码 And a.执行状态 = 0 And a.执行部门id = b.Id And b.Id = c.科室id And a.号别 = c.号码 And
                          Trunc(发生时间) Between Trunc(To_Date(Rs.开始停止时间, 'yyyy-mm-dd hh24:mi:ss')) And
                          Trunc(To_Date(Rs.结束停止时间, 'yyyy-mm-dd hh24:mi:ss'))) Loop
        --只返回该卡类别挂号的病人
        Select Count(*)
        Into n_Cnt
        From (Select 1
               From 病人预交记录 A
               Where a.No = Rs_Br.No And a.记录性质 = 4 And a.记录状态 = 1 And a.病人id = Rs_Br.病人id And 卡类别id = n_卡类别id
               Union All
               Select 1 From 病人挂号记录 Where NO = Rs_Br.No And 记录状态 = 1 And 交易说明 = v_Jsklb);
        If n_Cnt > 0 Then
          v_Brinfo := '<INFO><YYNO>' || Rs_Br.No || '</YYNO><BRID>' || Rs_Br.病人id || '</BRID><YYSJ>' || Rs_Br.发生时间 ||
                      '</YYSJ><CZSJ>' || Rs_Br.登记时间 || '</CZSJ>' || '<YYKS>' || Rs_Br.名称 || '</YYKS><GHLX>' || Rs_Br.号类 ||
                      '</GHLX><YSXM>' || Rs_Br.医生姓名 || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    --获取换诊记录
    v_Temp := '';
    For Rs In (Select d.No, a.病人id, To_Char(d.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                      To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.原号码, a.原医生姓名, b.专业技术职务 As 原职务, a.现号码, a.现医生姓名,
                      c.专业技术职务 As 现职务
               From 就诊变动记录 A, 人员表 B, 人员表 C, 病人挂号记录 D
               Where a.登记时间 Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.原医生id = b.Id And a.现医生id = c.Id And
                     a.挂号单 = d.No) Loop
      --只返回该卡类别挂号的病人         
      Select Count(*)
      Into n_Cnt
      From (Select 1
             From 病人预交记录 A
             Where a.No = Rs.No And a.记录性质 = 4 And a.记录状态 = 1 And a.病人id = Rs.病人id And 卡类别id = n_卡类别id
             Union All
             Select 1 From 病人挂号记录 Where NO = Rs.No And 记录状态 = 1 And 交易说明 = v_Jsklb);
      If n_Cnt > 0 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || Rs.病人id || '</BRID><YYSJ>' || Rs.登记时间 || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || Rs.预约时间 || '</YSJ><YHM>' || Rs.原号码 || '</YHM><YYS>' || Rs.原医生姓名 || '</YYS><YZC>' ||
                  Rs.原职务 || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || Rs.预约时间 || '</XSJ><XHM>' || Rs.现号码 || '</XHM><XYS>' || Rs.现医生姓名 || '</XYS><XZC>' ||
                  Rs.现职务 || '</XZC></ITEM>';
      End If;
    End Loop;
    v_Temp := '<HZLISTS>' || v_Temp || '</HZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getregistalter;
/


--133906:刘涛,2018-11-09,处理不分批入库收发记录批次为空
Create Or Replace Procedure Zl_药品外购_Insert
(
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  对方部门id_In In 药品收发记录.对方部门id%Type,
  供药单位id_In In 药品收发记录.供药单位id%Type,
  药品id_In     In 药品收发记录.药品id%Type,
  产地_In       In 药品收发记录.产地%Type := Null,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  实际数量_In   In 药品收发记录.实际数量%Type := Null,
  成本价_In     In 药品收发记录.成本价%Type := Null,
  成本金额_In   In 药品收发记录.成本金额%Type := Null,
  扣率_In       In 药品收发记录.扣率%Type := Null,
  零售价_In     In 药品收发记录.零售价%Type := Null,
  零售金额_In   In 药品收发记录.零售金额%Type := Null,
  差价_In       In 药品收发记录.差价%Type := Null,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  填制人_In     In 药品收发记录.填制人%Type := Null,
  发票号_In     In 应付记录.发票号%Type := Null,
  发票日期_In   In 应付记录.发票日期%Type := Null,
  发票金额_In   In 应付记录.发票金额%Type := Null,
  填制日期_In   In 药品收发记录.填制日期%Type := Null,
  外观_In       In 药品收发记录.外观%Type := Null,
  产品合格证_In In 药品收发记录.产品合格证%Type := Null,
  核查人_In     In 药品收发记录.配药人%Type := Null,
  核查日期_In   In 药品收发记录.配药日期%Type := Null,
  批次_In       In 药品收发记录.批次%Type := 0,
  退货_In       In Number := 1,
  生产日期_In   In 药品收发记录.生产日期%Type := Null,
  批准文号_In   In 药品收发记录.批准文号%Type := Null,
  随货单号_In   In 应付记录.随货单号%Type := Null,
  金额差_In     In 药品收发记录.零售金额%Type := Null,
  加成率_In     In 药品收发记录.频次%Type := Null,
  发票代码_In   In 应付记录.发票代码%Type := Null,
  计划id_In     In 药品收发记录.计划id%Type := Null,
  财务审核_In   In Number := 0,
  验收结论_In   In 药品收发记录.验收结论%Type := Null
) Is
  v_No         应付记录.No%Type; --应付记录的NO
  v_商品名     收费项目目录.名称%Type; --通用名称
  v_规格       收费项目目录.规格%Type;
  v_产地       收费项目目录.规格%Type;
  v_单位       收费项目目录.计算单位%Type;
  v_Lngid      药品收发记录.Id%Type; --收发ID
  v_应付id     应付记录.Id%Type; --应付记录的ID
  v_入出类别id 药品收发记录.入出类别id%Type; --入出类别ID
  v_入出系数   药品收发记录.入出系数%Type; --入出系数
  v_批次       药品收发记录.批次%Type := Null; --批次
  v_药库分批   Integer; --是否药库分批    1:分批；0：不分批
  v_药房分批   Integer; --是否药房分批       1:分批；0：不分批
  v_指导批价   药品规格.指导批发价%Type;
  v_时价分批   Number(1);

  Err_Msg Varchar2(255);
  Err_Noenough Exception;
Begin

  If Not 批准文号_In Is Null And Not 产地_In Is Null Then
    Update 药品生产商对照 Set 批准文号 = 批准文号_In Where 药品id = 药品id_In And 厂家名称 = 产地_In;
  End If;
  If Sql%RowCount = 0 And Not 产地_In Is Null And Not 批准文号_In Is Null Then
    Insert Into 药品生产商对照 (药品id, 厂家名称, 批准文号) Values (药品id_In, 产地_In, 批准文号_In);
  End If;

  --取该药品的商品名
  v_产地 := '';
  Select 名称, 规格, 计算单位, Nvl(是否变价, 0)
  Into v_商品名, v_规格, v_单位, v_时价分批
  From 收费项目目录
  Where ID = 药品id_In;
  If v_规格 Is Not Null Then
    If Instr(v_规格, '|') <> 0 Then
      v_产地 := Substr(v_规格, Instr(v_规格, '|'));
      v_规格 := Substr(v_规格, Instr(v_规格, '|') - 1);
    End If;
  End If;

  Select 药品收发记录_Id.Nextval Into v_Lngid From Dual;
  Select Nvl(药库分批, 0), Nvl(药房分批, 0), Nvl(指导批发价, 0)
  Into v_药库分批, v_药房分批, v_指导批价
  From 药品规格
  Where 药品id = 药品id_In;

  --财务审核_in=0表示普通入库，财务审核_in=1表示是财务审核产生新单据，如果是财务审核模式不需要重新产生批次
  If 财务审核_In = 0 Then
  	v_批次 := 0;
    If v_药房分批 = 0 Then
      If v_药库分批 = 1 Then
        Begin
          Select Distinct 0
          Into v_药库分批
          From 部门性质说明
          Where ((工作性质 Like '%药房') Or (工作性质 Like '制剂室')) And 部门id = 库房id_In;
        Exception
          When Others Then
            v_药库分批 := 1;
        End;
      
        If v_药库分批 = 1 Then
          v_批次 := Zl_Fun_Getbatchnum(药品id_In, 产地_In, 批号_In, 成本价_In, 零售价_In, v_Lngid, 供药单位id_In);
        End If;
      End If;
    Else
      v_批次 := Zl_Fun_Getbatchnum(药品id_In, 产地_In, 批号_In, 成本价_In, 零售价_In, v_Lngid, 供药单位id_In);
    End If;
  Else
    v_批次 := 批次_In;
  End If;

  If v_时价分批 = 1 And v_批次 > 0 Then
    v_时价分批 := 1;
  Else
    v_时价分批 := 0;
  End If;

  Select b.Id, b.系数
  Into v_入出类别id, v_入出系数
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 1 And Rownum < 2;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价,
     零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 配药日期, 发药方式, 单量, 外观, 产品合格证, 生产日期, 批准文号, 用法, 频次, 计划id, 验收结论)
  Values
    (v_Lngid, 1, 1, No_In, 序号_In, 库房id_In, 供药单位id_In, v_入出类别id, 对方部门id_In, v_入出系数, 药品id_In,
     Decode(退货_In, -1, 批次_In, v_批次), 产地_In, 批号_In, 效期_In, 退货_In * 实际数量_In, 退货_In * 实际数量_In, 成本价_In, 退货_In * 成本金额_In,
     扣率_In, 零售价_In, 退货_In * 零售金额_In, 退货_In * 差价_In, 摘要_In, 填制人_In, 填制日期_In, 核查人_In, 核查日期_In, Decode(退货_In, -1, 1, 0),
     v_指导批价, 外观_In, 产品合格证_In, 生产日期_In, 批准文号_In, Decode(退货_In, -1, Null, Decode(v_时价分批, 1, 金额差_In, Null)), 加成率_In,
     计划id_In, 验收结论_In);
  If 发票号_In Is Not Null Or 随货单号_In Is Not Null Then
    --如果是第一笔明细,则产生应付记录的NO
    Begin
      Select NO
      Into v_No
      From 应付记录
      Where 系统标识 = 1 And 记录性质 = 0 And 记录状态 = 1 And 入库单据号 = No_In And Rownum < 2;
    Exception
      When Others Then
        v_No := Nextno(67);
    End;
    Select 应付记录_Id.Nextval Into v_应付id From Dual;
    Insert Into 应付记录
      (ID, 记录性质, 记录状态, 单位id, NO, 系统标识, 收发id, 入库单据号, 单据金额, 发票号, 发票日期, 发票金额, 品名, 规格, 产地, 批号, 计量单位, 数量, 采购价, 采购金额, 填制人,
       填制日期, 审核人, 审核日期, 摘要, 项目id, 序号, 随货单号, 库房id, 发票修改时间, 发票代码)
    Values
      (v_应付id, 0, 1, 供药单位id_In, v_No, 1, v_Lngid, No_In, 退货_In * 零售金额_In, 发票号_In, 发票日期_In,
       退货_In * Decode(Nvl(发票金额_In, 0), 0, 成本金额_In, 发票金额_In), v_商品名, v_规格, v_产地, 批号_In, v_单位, 退货_In * 实际数量_In, 成本价_In,
       退货_In * 成本金额_In, 填制人_In, 填制日期_In, Null, Null, 摘要_In, 药品id_In, 序号_In, 随货单号_In, 库房id_In, Sysdate, 发票代码_In);
  End If;
  --调用库存更新过程
  Zl_药品库存_Update(v_Lngid, 0);
Exception
  When Err_Noenough Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品外购_Insert;
/

--125595:李南春,2018-10-23,预约接收不更新挂号科室
Create Or Replace Procedure Zl_病人预约挂号记录_Update
(
  单据号_In     门诊费用记录.No%Type,
  序号_In       门诊费用记录.序号%Type,
  价格父号_In   门诊费用记录.价格父号%Type,
  从属父号_In   门诊费用记录.从属父号%Type,
  收费类别_In   门诊费用记录.收费类别%Type,
  收费细目id_In 门诊费用记录.收费细目id%Type,
  数次_In       门诊费用记录.数次%Type,
  标准单价_In   门诊费用记录.标准单价%Type,
  收入项目id_In 门诊费用记录.收入项目id%Type,
  收据费目_In   门诊费用记录.收据费目%Type,
  应收金额_In   门诊费用记录.应收金额%Type,
  实收金额_In   门诊费用记录.实收金额%Type,
  病历费_In     Number, --该条记录是否病历工本费
  保险大类id_In 门诊费用记录.保险大类id%Type,
  保险项目否_In 门诊费用记录.保险项目否%Type,
  统筹金额_In   门诊费用记录.统筹金额%Type,
  保险编码_In   门诊费用记录.保险编码%Type,
  病人科室id_In 门诊费用记录.病人科室id%Type,
  执行部门id_In 门诊费用记录.执行部门id%Type,
  摘要_In       门诊费用记录.摘要%Type := Null,
  是否挂号项_In Number := 0
) As
  v_费用id 门诊费用记录.Id%Type;
  v_Error  Varchar2(255);
  Err_Custom Exception;
  Cursor c_费用 Is
    Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别,
           收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间,
           登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊
    From 门诊费用记录
    Where NO = 单据号_In And 记录性质 = 4 And 序号 = 1 And 记录状态 = 0;
Begin

  If Nvl(序号_In, 1) = 1 Then
    --第一条记录,只更新数据
    Update 门诊费用记录
    Set 价格父号 = Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号 = Decode(从属父号_In, 0, Null, 从属父号_In), 附加标志 = 病历费_In,
        收费类别 = 收费类别_In, 收费细目id = 收费细目id_In, 收入项目id = 收入项目id_In, 收据费目 = 收据费目_In, 付数 = 1, 数次 = 数次_In, 标准单价 = 标准单价_In,
        应收金额 = 应收金额_In, 实收金额 = 实收金额_In, 保险大类id = 保险大类id_In, 保险项目否 = 保险项目否_In, 保险编码 = 保险编码_In, 统筹金额 = 统筹金额_In,
        病人科室id =  Decode(是否挂号项_In, 1, 病人科室id, 病人科室id_In), 执行部门id = Decode(是否挂号项_In, 1, 执行部门id, 执行部门id_In), 摘要 = Nvl(摘要_In, 摘要)
    Where NO = 单据号_In And 序号 = 1 And 记录状态 = 0 And 记录性质 = 4;
    --删除序号大于1的数据;
    Delete 门诊费用记录 Where NO = 单据号_In And 序号 > 1 And 记录性质 = 4;
  Else
    --插入数据
    If 病历费_In <> 3 Then
      Select 病人费用记录_Id.Nextval Into v_费用id From Dual; --应该通过程序得到
      For r_费用 In c_费用 Loop
        Insert Into 门诊费用记录
          (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id,
           收费类别, 计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人,
           操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论)
        Values
          (v_费用id, 4, 0, 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, r_费用.实际票号, 1, r_费用.加班标志, 病历费_In,
           r_费用.发药窗口, r_费用.病人id, r_费用.标识号, r_费用.付款方式, r_费用.姓名, r_费用.性别, r_费用.年龄, r_费用.费别, 病人科室id_In, 收费类别_In, r_费用.计算单位,
           收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In, Null, Null, 0, r_费用.开单部门id, r_费用.操作员姓名,
           r_费用.操作员姓名, 执行部门id_In, r_费用.执行人, r_费用.操作员编号, r_费用.操作员姓名, r_费用.发生时间, r_费用.登记时间, 保险大类id_In, 保险项目否_In, 保险编码_In,
           统筹金额_In, Nvl(摘要_In, r_费用.摘要), r_费用.结论);
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预约挂号记录_Update;
/

--128511:殷瑞,2018-10-19,修正药品部门发药实际留存发药数不一致的问题
CREATE OR REPLACE Procedure Zl_药品留存记录_Insert
(
  期间_In       In 药品留存.期间%Type,
  汇总发药号_In In 药品收发记录.汇总发药号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  药品id_In     In 药品收发记录.药品id%Type,
  批次_In       In 药品收发记录.批次%Type := Null,
  实际数量_In   药品收发记录.实际数量%Type,
  零售价_In     In 药品收发记录.零售价%Type,
  填制人_In     In 药品收发记录.填制人%Type := Null,
  Date_In       In 药品收发记录.审核日期%Type,
  领药部门id_In In 药品收发记录.对方部门id%Type,
  登记时间_In   In 药品留存计划.登记时间%Type := Null
) Is
  Intdigit       Number(1);
  Nextid         Number;
  v_实际数量     药品收发记录.实际数量%Type;
  v_成本价       药品收发记录.成本价%Type;
  v_成本金额     药品收发记录.成本金额%Type;
  v_零售金额     药品收发记录.零售金额%Type;
  v_差价         药品收发记录.差价%Type;
  v_批号         药品收发记录.批号%Type;
  v_效期         药品收发记录.效期%Type;
  v_产地         药品收发记录.产地%Type;
  v_批准文号     药品收发记录.批准文号%Type;
  v_入出类别id   药品收发记录.入出类别id%Type; --入出类别ID
  v_入出系数     药品收发记录.入出系数%Type; --入出系数
  v_零售价       药品收发记录.零售价%Type; --零售价
  v_实际留存数量 药品留存计划.实际数量%Type;
Begin
  --获取金额小数位数
  Select Nvl(精度, 2) Into Intdigit From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;

  Select Sum(实际数量 * 付数) 实际数量, Avg(成本价) 成本价, Sum(零售金额) 零售金额, Sum(成本金额) 成本金额, Sum(差价) 差价, Max(批号) 批号, Max(效期) 效期,
         Max(产地) 产地, Max(批准文号) 批准文号
  Into v_实际数量, v_成本价, v_零售金额, v_成本金额, v_差价, v_批号, v_效期, v_产地, v_批准文号
  From 药品收发记录
  Where 汇总发药号 = 汇总发药号_In And 单据 In (8, 9, 10) And 库房id = 库房id_In And 对方部门id = 领药部门id_In And 药品id = 药品id_In And
        Nvl(批次, 0) = 批次_In;

  Select 药品收发记录_Id.Nextval Into Nextid From Dual;

  v_差价     := Round((实际数量_In / v_实际数量) * v_差价, Intdigit);
  v_成本金额 := Round((实际数量_In / v_实际数量) * v_成本金额, Intdigit);
  v_零售金额 := Round((实际数量_In / v_实际数量) * v_零售金额, Intdigit);
  v_零售价   := Round(v_零售金额 / 实际数量_In, Intdigit);
  --取药品留存记录的入出类别
  Select b.Id, b.系数
  Into v_入出类别id, v_入出系数
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 27 And Rownum < 2;

  --增加药品留存记录
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 汇总发药号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 填写数量, 实际数量, 零售价, 零售金额, 成本价, 成本金额, 差价, 填制日期, 填制人,
     审核日期, 审核人, 批号, 效期, 产地, 批准文号, 核查日期)
  Values
    (Nextid, 1, 27, Nextid, 汇总发药号_In, 库房id_In, 领药部门id_In, v_入出类别id, v_入出系数, 药品id_In, 批次_In, 实际数量_In, 实际数量_In, v_零售价,
     v_零售金额, v_成本价, v_成本金额, v_差价, Date_In, 填制人_In, Date_In, 填制人_In, v_批号, v_效期, v_产地, v_批准文号, 登记时间_In);

  --查询当前实际留存数量
  Select Nvl(实际数量, 留存数量)
  Into v_实际留存数量
  From 药品留存计划
  Where 部门id = 领药部门id_In And 库房id = 库房id_In And 药品id = 药品id_In And 状态 <> 1;

  --修改药品留存计划
  If v_实际留存数量 > 实际数量_In Then
    --部分执行 
    Update 药品留存计划
    Set 留存id = Nextid, 状态 = 2, 实际数量 = v_实际留存数量 - 实际数量_In
    Where 部门id = 领药部门id_In And 库房id = 库房id_In And 药品id = 药品id_In And 状态 <> 1;
  Else
    --全部执行
    Update 药品留存计划
    Set 留存id = Nextid, 状态 = 1, 实际数量 = 0
    Where 部门id = 领药部门id_In And 库房id = 库房id_In And 药品id = 药品id_In And 状态 <> 1;
  End If;

  --更新药品留存
  Update 药品留存
  Set 可用数量 = 可用数量 + 实际数量_In, 实际数量 = 实际数量 + 实际数量_In, 实际金额 = 实际金额 + v_零售金额
  Where 期间 = 期间_In And 科室id = 领药部门id_In And 库房id = 库房id_In And 药品id = 药品id_In;

  --如果没有库存则增加
  If Sql%NotFound Then
    Insert Into 药品留存
      (期间, 科室id, 库房id, 药品id, 可用数量, 实际数量, 实际金额)
    Values
      (期间_In, 领药部门id_In, 库房id_In, 药品id_In, 实际数量_In, 实际数量_In, v_零售金额);
  End If;

  --清除数量金额为零的记录
  Delete From 药品留存 Where 实际数量 = 0 And 实际金额 = 0;

  --增加药品库存
  Update 药品库存
  Set 可用数量 = Nvl(可用数量, 0) + Nvl(实际数量_In, 0), 实际数量 = Nvl(实际数量, 0) + Nvl(实际数量_In, 0), 实际金额 = Nvl(实际金额, 0) + Nvl(v_零售金额, 0),
      实际差价 = Nvl(实际差价, 0) + Nvl(v_差价, 0)
  Where 库房id = 库房id_In And 药品id = 药品id_In And Nvl(批次, 0) = 批次_In And 性质 = 1;

  If Sql%NotFound Then
    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次批号, 效期, 上次产地, 批准文号)
    Values
      (库房id_In, 药品id_In, 批次_In, 1, 实际数量_In, 实际数量_In, v_零售金额, v_差价, v_批号, v_效期, v_产地, v_批准文号);
  End If;

  Zl_药品库存_可用数量异常处理(库房id_In, 药品id_In, 批次_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品留存记录_Insert;
/

--132602:焦博,2018-10-19,修正OracleZl_小组轧帐记录_Cancel,获取上次轧帐时间时,加上判断财务缴款组ID进行判断
Create Or Replace Procedure Zl_小组轧帐记录_Cancel
(
  Id_In       In 人员收缴记录.Id%Type,
  作废人_In   In 人员收缴记录.作废人%Type,
  作废人id_In In 人员表.Id%Type,
  作废时间_In In 人员收缴记录.作废时间%Type,
  缴款组id_In In 财务缴款分组.Id%Type
) Is
  ---------------------------------------------------------------------------------------- 
  --功能:财务组轧帐记录作废 
  --参数:ID_IN:要作废记录的ID 
  ---------------------------------------------------------------------------------------- 
  v_Err_Msg Varchar2(100);
  Err_Item Exception;
  n_Exists       Number(2);
  n_Count        Number(18);
  v_No           人员收缴记录.No%Type;
  d_截止时间     人员收缴记录.终止时间%Type;
  v_收款员       人员收缴记录.收款员%Type;
  d_作废时间     人员收缴记录.作废时间%Type;
  d_财务收款时间 人员收缴记录.财务收款时间%Type;
Begin
  ---保存前并发检查 
  n_Exists := 0;

  Select Count(1), Max(NO), Max(收款员), Max(作废时间), Max(财务收款时间)
  Into n_Exists, v_No, v_收款员, d_作废时间, d_财务收款时间
  From 人员收缴记录
  Where ID = Id_In And 缴款组id = 缴款组id_In;

  If n_Exists = 0 Then
    v_Err_Msg := '记录未被找到，可能已被删除，无法进行轧帐作废操作！';
    Raise Err_Item;
  End If;

  If d_财务收款时间 Is Not Null Then
    v_Err_Msg := v_收款员 || '收款员的轧帐单号为[' || v_No || ']的记录已被财务科收款，不允许作废！';
    Raise Err_Item;
  End If;

  If d_作废时间 Is Not Null Then
    v_Err_Msg := v_收款员 || '收款员的轧帐单号为[' || v_No || ']的记录已被作废，不允许再次作废！';
    Raise Err_Item;
  End If;

  --检查是否最后一次轧帐记录 
  Select Count(*)
  Into n_Count
  From 人员收缴记录
  Where 登记时间 > (Select 登记时间 From 人员收缴记录 Where ID = Id_In) And 记录性质 = 3 And ID + 0 <> Id_In And Rownum < 2 And
        收款员 || '' = v_收款员 And 作废时间 Is Null And 缴款组id = 缴款组id_In;

  If n_Count >= 1 Then
    --是不是最后一次的轧帐记录 
    v_Err_Msg := '轧帐单号为:' || v_No || '的轧帐记录不是你最后一次的轧帐记录,不允许作废!';
    Raise Err_Item;
  End If;

  --作废轧帐操作 
  Update 人员收缴记录 Set 作废人 = 作废人_In, 作废时间 = 作废时间_In Where ID = Id_In And 记录性质 = 3;
  Insert Into 人员收缴对照
    (收缴id, 性质, 记录id)
    Select Id_In, 8, ID From 人员收缴记录 Where 小组轧账id = Id_In And 记录性质 = 2;
  Update 人员收缴记录 Set 小组轧账id = Null Where 小组轧账id = Id_In;

  --恢复最后一次有效的轧帐时间 
  Select Max(终止时间)
  Into d_截止时间
  From 人员收缴记录
  Where 登记时间 <= (Select 登记时间 From 人员收缴记录 Where ID = Id_In) And ID + 0 <> Id_In And 作废时间 Is Null And 财务收款时间 Is Null And
        收款员 || '' = v_收款员 And 记录性质 = 3 And 缴款组id = 缴款组id_In;
  If d_截止时间 Is Null Then
    --取组里最小一次收款记录的登记时间 
    Select Min(登记时间)
    Into d_截止时间
    From 人员收缴记录
    Where 记录性质 = 2 And 作废时间 Is Null And 财务收款时间 Is Null And 缴款组id = 缴款组id_In;
  End If;
  Update 财务缴款分组 Set 上次轧帐时间 = d_截止时间 Where ID = 缴款组id_In;
  Update 财务组组长构成 Set 上次轧帐时间 = d_截止时间 Where 组id = 缴款组id_In And 组长id = 作废人id_In;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_小组轧帐记录_Cancel;
/

--132941:殷瑞,2018-10-18,修正肿瘤药、抗生素及营养药不能正常发送到输液配置中心的问题
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
  v_Maxbatch    Number;
  v_Currdate    Date;
  n_Lngid       药品收发记录.Id%Type;
  n_Count       Number(3);
  n_单据        药品收发记录.单据%Type;
  v_No          药品收发记录.No%Type;
  n_发送次数    Number(5);
  n_病人id      病人信息.病人id%Type := 0;
  b_Change      Boolean := True;
  n_Sum         Number;
  n_调整批次    Number(1);
  n_Cur         Number(5);
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
  v_大输液给药途径 Varchar2(2000);
  v_来源科室       Varchar2(4000);
  v_Continue       Number := 1;
  v_Nodosage       Number := 0;
  v_保持上次批次   Number := 0;
  d_手工打包时间   Date;
  v_药品类型       Varchar2(20);
  n_打包药品批次   Number(1);
  n_特殊药品批次   Number(1);
  n_优先级         Number := 999;
  n_自动排批       Number := 0;
  n_科室id         Number := 0;
  n_Row            Number(2);
  n_备用批次       Number := 0;
  n_剩余数量       Number := 0;
  n_单次数量       Number := 0;
  n_累计数量       Number := 0;
  n_医嘱id         Number := 0;
  n_填写数量       Number := 0;
  v_配药类型       Varchar2(20);
  v_时间串         Varchar2(100);
  v_时间值         Date;
  v_Fields         Varchar2(100);
  v_是否改变       Varchar2(20);
  v_时间串1        Varchar2(100);
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
    Distinct e.医嘱id, e.发送号, b.频率间隔, b.间隔单位, b.执行时间方案, a.姓名, a.性别, a.年龄, a.住院号, a.当前床号 As 床号, a.当前病区id As 病人病区id,
             a.当前科室id As 病人科室id, e.首次时间, e.末次时间, b.开始执行时间, Nvl(e.发送数次, 0) As 次数, e.发送时间, Decode(b.医嘱期效, 0, 1, 2) As 医嘱类型,
             b.诊疗项目id As 给药途径, b.病人id
    From 病人医嘱发送 E, 病人医嘱记录 B, 病人信息 A, 诊疗项目目录 C
    Where e.医嘱id = b.Id And b.病人id = a.病人id And b.诊疗项目id = c.Id And b.相关id = v_相关id And e.发送号 = 发送号_In
    Order By e.医嘱id, e.发送号;

  Cursor c_收发记录 Is
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, 输液药品属性 E
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.Id = n_医嘱id And b.发送号 = 发送号_In And c.序号 < 1000
    Order By c.No, c.序号;

  Cursor c_原始收发记录 Is
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, 输液药品属性 E
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.相关id = v_相关id And b.发送号 = 发送号_In And c.序号 < 1000
    Order By c.No, c.序号;

  Cursor c_输液单记录 Is
    Select a.Id, a.执行时间, a.配药批次, a.医嘱id, d.发送时间
    From 输液配药记录 A, 病人医嘱记录 B, 配药工作批次 C, 病人医嘱发送 D
    Where a.医嘱id = b.Id And a.配药批次 = c.批次 And d.医嘱id = a.医嘱id And a.发送号 = d.发送号 And c.批次 <> 0 And c.药品类型 Is Null And
          b.病人id = n_病人id And a.操作状态 < 2 And a.执行时间 Between Trunc(v_时间值) And Trunc(v_时间值 + 1) - 1 / 24 / 60 / 60;

  v_输液单记录   c_输液单记录%RowType;
  v_医嘱记录     c_医嘱记录%RowType;
  v_收发记录     c_收发记录%RowType;
  v_单个医嘱记录 c_单个医嘱记录%RowType;

  Function Zl_Getpivaworkbatch
  (
    执行时间_In In Date,
    发送时间_In In Date,
    药品类型_In In Varchar2 := Null
  ) Return Number As

    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_配药批次 Is
      Select 批次, 配药时间, 给药时间, 打包, 药品类型
      From 配药工作批次
      Where 启用 = 1 And 配置中心id = 部门id_In
      Order By 药品类型, 批次;

    v_配药批次 c_配药批次%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');

    Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次 Where 启用 = 1 And 配置中心id = 部门id_In;

    For v_配药批次 In c_配药批次 Loop
      v_Batch := 0;

      --当天发送的医嘱发送到备用批次
      If (Trunc(执行时间_In) >= Trunc(v_Currdate) And Trunc(发送时间_In) < Trunc(执行时间_In)) Or n_备用批次 = 0 Then
        If v_配药批次.批次 <> '0' And
           ((Nvl(v_配药批次.药品类型, '0') <> '0' And v_配药批次.药品类型 = 药品类型_In) Or Nvl(v_配药批次.药品类型, '0') = '0') Then
          v_Starttime := To_Date(Substr(v_配药批次.给药时间, 1, Instr(v_配药批次.给药时间, '-') - 1), 'hh24:mi');
          v_Endtime   := To_Date(Substr(v_配药批次.给药时间, Instr(v_配药批次.给药时间, '-') + 1), 'hh24:mi');

          If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
            v_Batch := v_配药批次.批次;
            n_打包  := v_配药批次.打包;
            Exit When v_Batch > 0;
          End If;
        End If;
      End If;
    End Loop;

    If v_Batch = 0 And (n_打包药品批次 <> 1 Or n_备用批次 = 1) Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;

  Function Zl_Getfirst
  (
    配药id_In In Number,
    科室id_In In Number
  ) Return Number As
    n_First  Number;
    n_科室id Number;
    Cursor c_优先级 Is
      Select 科室id, 配药类型, 优先级, 频次
      From 输液药品优先级
      Where (科室id = 科室id_In Or 科室id = 0)
      Order By 科室id, 优先级 Desc;

    r_优先级 c_优先级%RowType;
  Begin
    n_First := 0;
    For r_优先级 In c_优先级 Loop
      If n_科室id <> 0 And r_优先级.科室id = 0 Then
        Exit;
      End If;
      n_科室id := r_优先级.科室id;

      For r_配药记录 In (Select Distinct d.配药类型, e.执行频次
                     From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C, 输液药品属性 D, 病人医嘱记录 E
                     Where a.医嘱id = e.Id And a.Id = b.记录id And b.收发id = c.Id And c.药品id = d.药品id And a.Id = 配药id_In) Loop
        If Instr(r_配药记录.配药类型, r_优先级.配药类型, 1) > 0 And (Instr(r_优先级.频次, r_配药记录.执行频次, 1) > 0 Or r_优先级.频次 = '所有频次') Then
          n_First := r_优先级.优先级;
          Exit;
        End If;
      End Loop;
    End Loop;

    If n_First = 0 Then
      n_First := 999;
    End If;
    Return(n_First);
  End;
Begin
  n_Count          := 0;
  v_医嘱类型       := Zl_To_Number(Nvl(zl_GetSysParameter('医嘱类型', 1345), 1));
  v_大输液给药途径 := Nvl(zl_GetSysParameter('输液给药途径', 1345), '');
  v_来源科室       := Nvl(zl_GetSysParameter('来源科室', 1345), '');
  v_保持上次批次   := Zl_To_Number(Nvl(zl_GetSysParameter('保持上次批次', 1345), 0));
  n_打包药品批次   := Zl_To_Number(Nvl(zl_GetSysParameter('单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包', 1345), 0));
  n_特殊药品批次   := Zl_To_Number(Nvl(zl_GetSysParameter('特殊药品按药品类型指定批次', 1345), 0));
  n_自动排批       := Zl_To_Number(Nvl(zl_GetSysParameter('启动自动排批', 1345), 0));
  n_备用批次       := Zl_To_Number(Nvl(zl_GetSysParameter('当天发送的医嘱产生的输液单全部到备用批次', 1345), 0));
  v_医嘱ids        := 医嘱id_In;
  v_当前病人       := '';
  n_发送次数       := 0;

  Select Trunc(Sysdate) Into v_Currdate From Dual;

  Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次;

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

  --先将原收发记录的序号增大，新的收发记录产生后再删除
  --Update 药品收发记录
  --Set 序号 = 序号 + 10000
  --Where ID In (Select \*+rule *\
  --             Distinct c.Id
  --             From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, Table(f_Num2list(医嘱id_In)) F
  --             Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And b.执行部门id + 0 = 部门id_In And
  --                   c.单据 = 9 And c.审核日期 Is Null And a.相关id = f.Column_Value And b.发送号 = 发送号_In And c.序号 < 10000);

  For v_医嘱记录 In c_医嘱记录 Loop
    v_Continue := 1;
    n_病人id   := v_医嘱记录.病人id;
    n_科室id   := v_医嘱记录.病人科室id;

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
      If Instr(',' || v_来源科室 || ',', ',' || v_医嘱记录.病人科室id || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;

    v_药品类型 := Null;
    For r_药品类型 In (Select Decode(Nvl(d.抗生素, 0), 0, Decode(Nvl(d.是否肿瘤药, 0), 0, '', '肿瘤药'), '抗生素') 药品类型
                   From 病人医嘱记录 A, 药品规格 B, 住院费用记录 C, 药品特性 D
                   Where c.收费细目id = b.药品id And b.药名id = d.药名id And c.医嘱序号 = a.Id And a.相关id = v_医嘱记录.相关id) Loop
      If r_药品类型.药品类型 Is Not Null Then
        v_药品类型 := r_药品类型.药品类型;
        v_Continue := 1;
      End If;
    End Loop;

    If v_药品类型 Is Null Then
      If v_医嘱记录.是否tpn = 2 Then
        v_药品类型 := '营养药';
        v_Continue := 1;
      End If;
    End If;
    
    --输液不配置药品限制
    Select Count(1)
    Into v_Continue
    From 病人医嘱记录 A, 输液不配置药品 B, 住院费用记录 C
    Where c.收费细目id = b.药品id And c.医嘱序号 = a.Id And a.相关id = v_医嘱记录.相关id;
    If v_Continue = 0 Then
      v_Continue := 1;
    Else
      v_Continue := 0;
    End If;
    
    If v_Continue = 1 Then
      v_Old相关id := v_New相关id;
      v_相关id    := v_医嘱记录.相关id;
      v_New相关id := v_相关id;
      v_发送号    := v_医嘱记录.发送号;
      v_序号      := 0;

      If v_Continue = 1 Then
        --v_Count := Zl_Gettransexenumber(v_医嘱记录.开始执行时间, v_医嘱记录.首次时间, v_医嘱记录.末次时间, v_医嘱记录.频率间隔, v_医嘱记录.间隔单位, v_医嘱记录.执行时间方案);
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
            b_Change := True;
          End If;

          If b_Change = True Then
            If d_Old执行时间 <> Trunc(v_执行时间) Or d_Old执行时间 Is Null Then
              d_Old执行时间 := v_执行时间;

              Select Count(Distinct a.摆药单号)
              Into n_摆药单
              From 输液配药记录 A
              Where a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = v_医嘱记录.病人id And 相关id Is Null) And
                    a.执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60 And 操作状态 >= 2 And 操作状态 < 9;

              If n_摆药单 > 1 Then
                Update 输液配药记录
                Set 是否调整批次 = 1
                Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And

                      执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                b_Change := False;

              End If;
            End If;
          End If;

          If b_Change = True Then
            n_病人id := v_医嘱记录.病人id;
            Select Count(ID)

            Into n_Sum
            From 输液配药记录
            Where 医嘱id = v_医嘱记录.相关id And 执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;
            If n_Sum = 0 Then
              Update 输液配药记录
              Set 是否调整批次 = 1
              Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And

                    执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
              b_Change := False;

            End If;

            If b_Change = True Then
              --检查输液单是否调整到打包状态
              Select Count(a.Id)
              Into n_Sum
              From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C
              Where a.Id = b.记录id And b.收发id = c.Id And
                    a.医嘱id In (Select ID
                               From 病人医嘱记录
                               Where 病人id = (Select 病人id From 病人医嘱记录 Where ID = v_医嘱记录.相关id And Rownum < 2)) And
                    a.执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60 And a.打包时间 Is Not Null;
              If n_Sum <> 0 Then
                Update 输液配药记录
                Set 是否调整批次 = 1
                Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_执行时间) And
                      Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                b_Change := False;
              End If;

              Select Count(医嘱id)
              Into n_Cur
              From 医嘱执行时间
              Where 医嘱id = v_医嘱记录.相关id And 要求时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60;

              Select Count(医嘱id)
              Into n_Sum
              From 医嘱执行时间
              Where 医嘱id = v_医嘱记录.相关id And 要求时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;

              If n_Sum <> n_Cur Then
                Update 输液配药记录
                Set 是否调整批次 = 1
                Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_执行时间) And
                      Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                b_Change := False;
              End If;
            End If;
          End If;

          If v_时间串 <> Trunc(Sysdate) || ';false\' Or v_时间串 Is Null Then
            If Trunc(v_执行时间) = Trunc(Sysdate) Then
              If b_Change = False Then
                v_时间串 := Trunc(v_执行时间) || ';false\';
              Else
                v_时间串 := Trunc(v_执行时间) || ';true\';
              End If;
            End If;
          End If;

          If v_时间串1 <> Trunc(Sysdate + 1) || ';false\' Or v_时间串1 Is Null Then
            If Trunc(v_执行时间) = Trunc(Sysdate + 1) Then
              If b_Change = False Then
                v_时间串1 := Trunc(v_执行时间) || ';false\';
              Else
                v_时间串1 := Trunc(v_执行时间) || ';true\';
              End If;
            End If;
          End If;

          If v_药品类型 Is Null Or n_特殊药品批次 = 0then v_批次 := Zl_Getpivaworkbatch(v_执行时间, Sysdate) ; Else
          --药品类型不为空，直接根据药品类型匹配批次
           v_批次 := Zl_Getpivaworkbatch(v_执行时间, Sysdate, v_药品类型) ; End If ;

            Select Count(医嘱id)
              Into n_发送次数
              From 医嘱执行时间
              Where 医嘱id = v_医嘱记录.相关id And 要求时间 <= v_执行时间
              Order By 要求时间;

              If n_发送次数 > 99
           Then
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
            Select Nvl(Max(打包), 0), Max(药品类型)
            Into n_打包, v_配药类型
            From 配药工作批次
            Where 批次 = v_批次 And 配置中心id = 部门id_In;
          End If;

          If (Trunc(v_执行时间) <= v_Currdate Or n_打包 <> 0) And v_配药类型 Is Null Then
            n_是否打包     := 1;
            d_手工打包时间 := Null;
          Else
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;

          --如果是TPN不管其他条件如何都设置为配置
          If v_医嘱记录.是否tpn = 2 Then
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;

          If v_批次 = 0 Then
            n_是否打包 := 1;
          End If;
          --产生配药记录
          Insert Into 输液配药记录
            (ID, 部门id, 序号, 姓名, 性别, 年龄, 住院号, 床号, 病人病区id, 病人科室id, 执行时间, 医嘱id, 发送号, 配药批次, 瓶签号, 是否调整批次, 是否打包, 打包时间, 操作状态,
             操作人员, 操作时间)
          Values
            (v_配药id, 部门id_In, v_序号, v_医嘱记录.姓名, v_医嘱记录.性别, v_医嘱记录.年龄, v_医嘱记录.住院号, v_医嘱记录.床号, v_医嘱记录.病人病区id,
             v_医嘱记录.病人科室id, v_执行时间, v_医嘱记录.相关id, v_医嘱记录.发送号, v_批次, v_Maxno, n_调整批次, n_是否打包, d_手工打包时间, 1, 核查人_In, 核查时间_In);

          Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (v_配药id, 1, 核查人_In, 核查时间_In);

          For v_单个医嘱记录 In c_单个医嘱记录 Loop
            n_医嘱id   := v_单个医嘱记录.医嘱id;
            n_累计数量 := 0;
            n_剩余数量 := 0;

            Select Sum(c.实际数量)
            Into n_Sum
            From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D
            Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And b.执行部门id + 0 = 部门id_In And
                  c.单据 = 9 And c.审核日期 Is Null And a.Id = n_医嘱id And b.发送号 = v_医嘱记录.发送号 And c.序号 < 1000;

            --产生配药记录对应的药品记录
            For v_收发记录 In c_收发记录 Loop
              If v_收发记录.是否不予配置 = 1 Then
                v_Nodosage := 1;
              End If;

              Select 药品收发记录_Id.Nextval Into n_Lngid From Dual;
              n_累计数量 := n_累计数量 + v_收发记录.数量;

              If n_剩余数量 = 0 Then
                n_剩余数量 := n_Sum / v_Count;
              End If;
              n_单次数量 := n_Sum / v_Count;

              If n_累计数量 >= n_Sum / v_Count * I Then
                n_Count := n_Count + 1;
                Insert Into 药品收发记录
                  (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 付数, 填写数量, 实际数量,
                   成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期,
                   灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间)
                  Select n_Lngid, 记录状态, 单据, NO, n_Count + 1000, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号,
                         生产日期, 效期, 付数, n_剩余数量, n_剩余数量, 成本价, 成本价 * n_剩余数量, 扣率, 零售价, 零售价 * n_剩余数量, 差价 * (实际数量 / n_剩余数量),
                         '复制', 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期, 灭菌效期, 产品合格证, 发药方式,
                         发药窗口, 领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间

                  From 药品收发记录
                  Where ID = v_收发记录.收发id;

                Insert Into 输液配药内容 (记录id, 收发id, 数量) Values (v_配药id, n_Lngid, n_剩余数量);

                n_剩余数量 := 0;
                Exit;
              Elsif n_累计数量 > (n_Sum / v_Count * (I - 1)) Then
                n_Count    := n_Count + 1;
                n_填写数量 := n_累计数量 - (n_Sum / v_Count * (I - 1)) - (n_单次数量 - n_剩余数量);
                Insert Into 药品收发记录
                  (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 付数, 填写数量, 实际数量,
                   成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期,
                   灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间)
                  Select n_Lngid, 记录状态, 单据, NO, n_Count + 1000, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号,
                         生产日期, 效期, 付数, n_填写数量, n_填写数量, 成本价, 成本价 * n_填写数量, 扣率, 零售价, 零售价 * n_填写数量, 差价 * (实际数量 / n_填写数量),
                         '复制', 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期, 灭菌效期, 产品合格证, 发药方式,
                         发药窗口, 领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间

                  From 药品收发记录
                  Where ID = v_收发记录.收发id;

                Insert Into 输液配药内容 (记录id, 收发id, 数量) Values (v_配药id, n_Lngid, n_填写数量);

                n_剩余数量 := n_剩余数量 - n_填写数量;
              End If;
            End Loop;
          End Loop;
          n_优先级 := Zl_Getfirst(v_配药id, v_医嘱记录.病人科室id);
          Update 输液配药记录 Set 优先级 = n_优先级 Where ID = v_配药id;

        End Loop;

        For v_收发记录 In c_原始收发记录 Loop
          n_单据 := v_收发记录.单据;

          v_No := v_收发记录.No;
          Delete From 药品收发记录 Where ID = v_收发记录.收发id;
        End Loop;

        --单个药品或者不予配置的药品默认为0批次
        Select Count(收发id) Into n_Row From 输液配药内容 Where 记录id = v_配药id;
        If (v_Nodosage = 1 Or n_Row = 1) And n_打包药品批次 = 1 Then
          Update 输液配药记录
          Set 配药批次 = 0, 是否打包 = 1
          Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And 操作状态 < 2;
        End If;
        --如果存在“不予配置”属性的药品，也设置为打包
        If v_Nodosage = 1 Then
          Update 输液配药记录
          Set 是否打包 = 1
          Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And 操作状态 < 2;
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

  If n_备用批次 = 1 Then

    Select Count(a.Id)
    Into n_Sum
    From 输液配药记录 A, 病人医嘱发送 B
    Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And
          a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id And 相关id Is Null) And b.发送时间 Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And a.执行时间 Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And 操作状态 < 9;
    If n_Sum <> 0 Then
      b_Change  := False;
      v_时间串1 := Trunc(Sysdate + 1) || ';false\';

      Update 输液配药记录
      Set 是否调整批次 = 1
      Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(Sysdate + 1) And
            Trunc(Sysdate + 2) - 1 / 24 / 60 / 60 And 操作状态 < 2;
    End If;
  End If;
  If v_时间串 Is Null Then
    v_时间串 := v_时间串1;
  Else
    v_时间串 := v_时间串 || v_时间串1;
  End If;

  While v_时间串 Is Not Null Loop
    --分解单据ID串
    v_Fields   := Substr(v_时间串, 1, Instr(v_时间串, '\') - 1);
    v_时间值   := Substr(v_Fields, 1, Instr(v_Fields, ';') - 1);
    v_是否改变 := Substr(v_Fields, Instr(v_Fields, ';') + 1);

    v_时间串 := Replace('\' || v_时间串, '\' || v_Fields || '\');

    If v_是否改变 = 'true' Then
      b_Change := True;
    Else
      b_Change := False;
    End If;

    If b_Change = True Then
      Select Count(医嘱id)
      Into n_Cur
      From (Select Distinct a.要求时间, a.医嘱id
             From 医嘱执行时间 A, 输液配药记录 B
             Where a.要求时间 = b.执行时间 And a.医嘱id = b.医嘱id And a.要求时间 Between Trunc(v_时间值) And
                   Trunc(v_时间值 + 1) - 1 / 24 / 60 / 60 And
                   a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id And 相关id Is Null));
      Select Count(医嘱id)
      Into n_Sum
      From (Select Distinct a.要求时间, a.医嘱id
             From 医嘱执行时间 A, 输液配药记录 B
             Where a.要求时间 = b.执行时间 And a.医嘱id = b.医嘱id And a.要求时间 Between Trunc(v_时间值 - 1) And
                   Trunc(v_时间值) - 1 / 24 / 60 / 60 And
                   a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id And 相关id Is Null));

      If n_Cur <> n_Sum Then
        Update 输液配药记录
        Set 是否调整批次 = 1
        Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_时间值) And
              Trunc(v_时间值 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
        b_Change := False;
      End If;
    End If;

    If v_保持上次批次 = 1 And b_Change = True Then
      For v_输液单记录 In c_输液单记录 Loop
        Begin
          Select Distinct 配药批次
          Into v_批次
          From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C
          Where a.Id = b.记录id And b.收发id = c.Id And a.医嘱id = v_输液单记录.医嘱id And
                To_Char(a.执行时间, 'hh24:mi:ss') = To_Char(v_输液单记录.执行时间, 'hh24:mi:ss') And
                a.执行时间 Between Trunc(v_输液单记录.执行时间 - 1) And Trunc(v_输液单记录.执行时间) - 1 / 24 / 60 / 60 And Rownum = 1;
        Exception
          When Others Then
            Begin
              Select Distinct 配药批次
              Into v_批次
              From 输液配药记录 A
              Where a.医嘱id = v_输液单记录.医嘱id And To_Char(a.执行时间, 'hh24:mi:ss') = To_Char(v_输液单记录.执行时间, 'hh24:mi:ss') And
                    a.操作状态 <> 12 And a.执行时间 Between Trunc(v_输液单记录.执行时间 - 1) And Trunc(v_输液单记录.执行时间) - 1 / 24 / 60 / 60 And
                    Rownum = 1;
            Exception
              When Others Then
                v_批次 := v_输液单记录.配药批次;
            End;
        End;

        Update 输液配药记录
        Set 是否确认调整 = 0, 是否调整批次 = 0
        Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_输液单记录.执行时间) And
              Trunc(v_输液单记录.执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;

        If v_输液单记录.配药批次 <> v_批次 Then
          Update 输液配药记录 Set 配药批次 = v_批次 Where ID = v_输液单记录.Id;
          Select Nvl(Max(打包), 0) Into n_打包 From 配药工作批次 Where 批次 = v_批次 And 配置中心id = 部门id_In;
          If n_打包 <> 0 Then
            Update 输液配药记录 Set 是否打包 = n_打包 Where ID = v_输液单记录.Id;
          Else
            Select Nvl(Max(打包), 0)
            Into n_打包
            From 配药工作批次
            Where 批次 = v_输液单记录.配药批次 And 配置中心id = 部门id_In;

            If n_打包 <> 0 Then
              Update 输液配药记录 Set 是否打包 = 0 Where ID = v_输液单记录.Id;
            End If;
          End If;
        End If;
      End Loop;
    End If;

    If n_自动排批 = 1 And (b_Change = False Or v_保持上次批次 = 0) Then
      For v_输液单记录 In c_输液单记录 Loop
        v_批次 := Zl_Getpivaworkbatch(v_输液单记录.执行时间, v_输液单记录.发送时间);
        Update 输液配药记录 Set 配药批次 = v_批次 Where ID = v_输液单记录.Id;
      End Loop;
      Zl_输液配药记录_自动排批(n_病人id, n_科室id, 部门id_In, v_时间值);
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]病人' || v_当前病人 || '在输液配置中心有被锁定的输液单，发送失败！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_核查;
/

--132638:胡俊勇,2018-10-11,身份证唯一性处理判断
Create Or Replace Function Zl_Pati_Is_Inhospital(病人id_In In 病人信息.病人id%Type) Return Number Is
  n_Count   Number;
  v_Sql     Varchar2(200);
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  --功能说明判断指定病人是否是处于在院就医状态
  --返回：0-不是在院就医状态，1-是处于在院就医状态
  --住院
  Select Count(1) Into n_Count From 病案主页 A Where a.病人id = 病人id_In And a.出院日期 Is Null;
  If n_Count > 0 Then
    Return(1);
  End If;

  --挂号
  Select Count(1)
  Into n_Count
  From 病人挂号记录 A
  Where a.病人id = 病人id_In And Nvl(a.执行状态, 0) <> 1 And a.记录状态 = 1;
  If n_Count > 0 Then
    Return(1);
  End If;

  Begin
    n_Count := 0;
    --老版体检：系统号是2100
    v_Sql := 'Select Count(1) From 体检任务人员 A, 体检任务记录 B  Where a.任务id = b.Id And a.体检状态 = 2 And b.任务状态 <> 4 And a.病人id= :1';
    Execute Immediate v_Sql
      Into n_Count
      Using 病人id_In;
    If n_Count > 0 Then
      Return(1);
    End If;
  Exception
    When Others Then
      Null;
  End;

  Begin
    n_Count := 0;
    --新版体检：系统号是2700
    v_Sql := 'Select Count(1)  From 体检登记人员 A, 体检任务登记 B  Where a.任务登记id = b.Id And a.体检状态 = 1 And b.任务状态 <> 3 And a.体检人员id =:1';
    Execute Immediate v_Sql
      Into n_Count
      Using 病人id_In;
    If n_Count > 0 Then
      Return(1);
    End If;
  Exception
    When Others Then
      Null;
  End;
  Return(0);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pati_Is_Inhospital;
/

--129825:刘涛,2018-10-08,增加传入结存日期
Create Or Replace Procedure Zl_材料结存记录_Insert
(
  库房id_In In 材料结存记录.库房id%Type := Null,
  填制人_In In 材料结存记录.填制人%Type := Null,
  转结_In   In Number := 1,
  结存日期_In In 材料结存记录.期末日期%Type := Null
) Is
  v_Lngid      材料结存记录.Id%Type;
  d_开始日期   材料结存记录.期初日期%Type;
  d_结束日期   材料结存记录.期末日期%Type;
  n_结存时点   Number(2);
  n_上次结存id 材料结存记录.Id%Type;
  v_上次期间   材料结存记录.期间%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;
  n_未审核记录 Number(1) := 0;
Begin
  If 转结_In = 0 Then
    --初始结存，期初日期 = 期末日期= 当前系统日期 
    d_开始日期 := Sysdate;
    d_结束日期 := d_开始日期;
  Else
    --检查是否存在未审核的结存记录，如果存在则不能结存 
    Select Count(ID) Into n_未审核记录 From 材料结存记录 Where 库房id = 库房id_In And 审核日期 Is Null;
  
    If n_未审核记录 > 0 Then
      v_Error := '上次结存未审核，不能再次结存';
      Raise Err_Custom;
    End If;
  
    --取结存时点，默认每月最后一日结存 
    n_结存时点 := Nvl(zl_GetSysParameter(281), 0);
  
    If n_结存时点 <> -1 Then
      --自动结存 
      --取上次结存ID,上期期末日期作为本期的期初日期,上次期间 
      Select Max(ID), Trunc(Max(期末日期)) + 1, Max(期间)
      Into n_上次结存id, d_开始日期, v_上次期间
      From 材料结存记录
      Where 库房id = 库房id_In And 取消人 Is Null;
    
      --自动结存 
      If n_结存时点 = 0 Or n_结存时点 > To_Number(To_Char(Trunc(Last_Day(d_开始日期 - 1)), 'dd')) Then
        --指定按每月最后一天结存；或者结存时点大于了本月最大天数，也按本月最后一天结存 
        d_结束日期 := Trunc(Last_Day(Sysdate - 1)) + 1 - 1 / 24 / 60 / 60;
      Else
        d_结束日期 := Trunc(Sysdate - 1, 'MONTH') + n_结存时点 - 1 / 24 / 60 / 60;
      End If;
      --检查日期，在结存时点后才能进行结存 
      If Sysdate - d_结束日期 < 0 Then
        v_Error := '本月结存时点未到，不能提前结存！';
        Raise Err_Custom;
      End If;
    
      --检查期间 
      If v_上次期间 = To_Char(Trunc(d_结束日期), 'yyyymm') Then
        v_Error := '本月已经结存，不能再次结存！';
        Raise Err_Custom;
      End If;
    Else
      --手工结存 
      --取上次结存ID,上期期末日期作为本期的期初日期,上次期间 
      Select Max(ID), Max(期末日期) + 1 / 24 / 60 / 60, Max(期间)
      Into n_上次结存id, d_开始日期, v_上次期间
      From 材料结存记录
      Where 库房id = 库房id_In And 取消人 Is Null;
      
	  If 结存日期_In Is Null Then
        d_结束日期 := Sysdate;
      Else
        d_结束日期 := 结存日期_In;
      End If;

    End If;
  End If;

  Select 材料结存记录_Id.Nextval Into v_Lngid From Dual;

  --产生材料结存主表 
  Insert Into 材料结存记录
    (ID, 库房id, 期初日期, 期末日期, 填制人, 填制日期, 审核人, 审核日期, 上次结存id, 期间, 性质)
  Values
    (v_Lngid, 库房id_In, d_开始日期, d_结束日期, Nvl(填制人_In, Zl_Username), Sysdate,
     Decode(d_开始日期, Null, Nvl(填制人_In, Zl_Username), Null), Decode(d_开始日期, Null, Sysdate, Null), n_上次结存id,
     To_Char(Trunc(d_结束日期), 'yyyymm'), Decode(转结_In, 0, 0, 1));

  If 转结_In = 0 Then
    --初始结存，以当前库存为准期末 = 期初=当前库存数据 
    Insert Into 材料结存明细
      (结存id, 库房id, 材料id, 批次, 期初数量, 期初金额, 期初差价, 期末数量, 期末金额, 期末差价)
      Select v_Lngid, 库房id, 药品id, 批次, Sum(期初数量), Sum(期初金额), Sum(期初差价), Sum(期末数量), Sum(期末金额), Sum(期末差价)
      From (Select a.库房id, a.药品id, Nvl(a.批次, 0) 批次, Nvl(a.实际数量, 0) As 期初数量, Nvl(a.实际金额, 0) As 期初金额,
                    Nvl(a.实际差价, 0) As 期初差价, Nvl(a.实际数量, 0) As 期末数量, Nvl(a.实际金额, 0) As 期末金额, Nvl(a.实际差价, 0) As 期末差价
             From 药品库存 A,材料特性 B 
             Where a.药品id = b.材料id And a.性质 = 1 And a.库房id = 库房id_In
             Union All
             Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, -1 * a.入出系数 * a.实际数量 * Nvl(a.付数, 1) As 期初数量,
                    -1 * a.入出系数 * a.零售金额 As 期初金额, -1 * a.入出系数 * a.差价 As 期初差价, -1 * a.入出系数 * a.实际数量 * Nvl(a.付数, 1) As 期末数量,
                    -1 * a.入出系数 * a.零售金额 As 期末金额, -1 * a.入出系数 * a.差价 As 期末差价
             From 药品收发记录 A,材料特性 B 
             Where a.药品id = b.材料id And a.库房id + 0 = 库房id_In And a.审核日期 > d_结束日期)
      Group By 库房id, 药品id, 批次
      Order By 库房id, 药品id, 批次;
  Else
    --产生药品结存明细表，本期期末=本期期初(上期期末)+期间发生 
    Insert Into 材料结存明细
      (结存id, 库房id, 材料id, 批次, 期初数量, 期初金额, 期初差价, 期末数量, 期末金额, 期末差价)
      Select v_Lngid, 库房id, 材料id, 批次, Sum(期初数量), Sum(期初金额), Sum(期初差价), Sum(期末数量), Sum(期末金额), Sum(期末差价)
      From (Select 库房id, 材料id, Nvl(批次, 0) As 批次, 期末数量 As 期初数量, 期末金额 As 期初金额, 期末差价 As 期初差价, 期末数量, 期末金额, 期末差价
             From 材料结存明细
             Where 结存id = n_上次结存id
             Union All
             Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, 0 As 期初数量, 0 As 期初金额, 0 As 期初差价,
                    a.入出系数 * a.实际数量 * Nvl(a.付数, 1) As 期末数量, a.入出系数 * a.零售金额 As 期末金额, a.入出系数 * a.差价 As 期末差价
             From 药品收发记录 A,材料特性 B 
             Where a.药品id = b.材料id And a.库房id + 0 = 库房id_In And a.审核日期 Between d_开始日期 And d_结束日期)
      Group By 库房id, 材料id, 批次
      Order By 库房id, 材料id, 批次;
  
    --计算误差：本期期末-库存记录(减去本期期末时间后发生的数据) 
    Insert Into 材料结存误差
      (ID, 结存id, 库房id, 材料id, 批次, 数量差, 金额差, 差价差)
      Select 材料结存误差_Id.Nextval, v_Lngid, a.库房id, a.材料id, a.批次, a.实际数量 As 数量差, a.实际金额 As 金额差, a.实际差价 As 差价差
      From (Select 库房id, 材料id, 批次, Sum(实际数量) As 实际数量, Sum(实际金额) As 实际金额, Sum(实际差价) As 实际差价
             From (Select a.库房id, a.材料id, Nvl(a.批次, 0) As 批次, Nvl(a.期末数量, 0) As 实际数量, Nvl(a.期末金额, 0) As 实际金额,
                           Nvl(a.期末差价, 0) As 实际差价
                    From 材料结存明细 A
                    Where a.结存id = v_Lngid
                    Union All
                    Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, -1 * Nvl(a.实际数量, 0) As 实际数量, -1 * Nvl(a.实际金额, 0) As 实际金额,
                           -1 * Nvl(a.实际差价, 0) As 实际差价
                    From 药品库存 A,材料特性 B
                    Where a.药品id = b.材料id And a.性质 = 1 And a.库房id = 库房id_In
                    Union All
                    Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, a.入出系数 * a.实际数量 * Nvl(a.付数, 1) As 实际数量,
                           a.入出系数 * a.零售金额 As 实际金额, a.入出系数 * a.差价 As 实际差价
                    From 药品收发记录 A,材料特性 B
                    Where a.药品id = b.材料id And a.库房id = 库房id_In And a.审核日期 > d_结束日期) A
             Group By 库房id, 材料id, 批次) A
      Where a.实际数量 <> 0 Or a.实际金额 <> 0 Or a.实际差价 <> 0;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料结存记录_Insert;
/

--119346:冉俊明,2019-01-09,启用临床出诊安排，医保补充结算退号时没有更新临床出诊序号控制及临床出诊记录
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
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_病人id     病人信息.病人id%Type;
  n_出诊记录id 病人挂号记录.出诊记录id%Type;
  d_发生时间   就诊登记记录.就诊时间%Type;

  n_结帐id   病人预交记录.结帐id%Type;
  n_冲销id   门诊费用记录.结帐id%Type;
  n_结算序号 病人预交记录.结算序号%Type;
  n_退费金额 病人预交记录.冲预交%Type;
  n_组id     财务缴款分组.Id%Type;
  d_退号时间 病人预交记录.收款时间%Type;

  n_退号重用 Number;
  v_号别     病人挂号记录.号别%Type;
  n_号序     病人挂号记录.号序%Type;

  n_挂号id         病人挂号记录.Id%Type;
  n_分诊台签到排队 Number;
  n_挂号生成队列   Number;
  n_病人id1        病人信息.病人id%Type;
  d_Temp           Date;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  --首先判断要退号/取消预约的记录是否存在
  Begin
    Select 病人id, 发生时间, 出诊记录id, 号别, 号序
    Into n_病人id, d_发生时间, n_出诊记录id, v_号别, n_号序
    From 病人挂号记录
    Where NO = 单据号_In;
  Exception
    When Others Then
      v_Err_Msg := '未找到指定的挂号单:' || 单据号_In || ',可能已经被人退号,不允许再次退号。';
      Raise Err_Item;
  End;

  Begin
    Select a.结帐id
    Into n_结帐id
    From 门诊费用记录 A
    Where a.记录性质 = 4 And a.No = 单据号_In And a.记录状态 = 1 And 病人id = n_病人id And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := '未找到指定的挂号单:' || 单据号_In || ',可能已经被人退号,不允许再次退号。';
      Raise Err_Item;
  End;

  --2.挂号处理
  d_退号时间 := Nvl(退号时间_In, Sysdate);
  n_冲销id   := 冲销id_In;
  n_结算序号 := 结算序号_In;
  If n_冲销id Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_冲销id From Dual;
  End If;
  If n_结算序号 Is Null Then
    n_结算序号 := -1 * n_冲销id;
  End If;

  --病人挂号记录
  Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 1;
  If Sql%NotFound Then
    v_Err_Msg := '挂号单【' || 单据号_In || '】不存在或由于并发原因已经被退号。';
    Raise Err_Item;
  End If;

  Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
  Insert Into 病人挂号记录
    (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名, 复诊,
     号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式)
    Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_退号时间, 发生时间, 操作员编号_In,
           操作员姓名_In, 复诊, 号序, 社区, 预约, 摘要 As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式
    From 病人挂号记录
    Where NO = 单据号_In And 记录状态 = 3;

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
  Select Max(病人id)
  Into n_病人id1
  From 门诊费用记录
  Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 附加标志 = 2 And Rownum < 2;
  If n_病人id1 Is Not Null Then
    Update 病人信息
    Set 就诊卡号 = Null, 卡验证码 = Null, Ic卡号 = Decode(Ic卡号, 就诊卡号, Null, Ic卡号)
    Where 病人id = n_病人id1;
  End If;

  --门诊费用记录
  Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
  Insert Into 门诊费用记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
     数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
     结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id, 费用状态, 执行状态)
    Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
           收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
           操作员编号_In, 操作员姓名_In, 发生时间, d_退号时间, n_冲销id, -1 * 结帐金额, 保险项目否, 保险大类id, -1 * 统筹金额, 摘要 As 摘要, 附加标志, 保险编码, 费用类型,
           n_组id, 1, -1
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In;

  --产生结算方式为NULL的记录
  Update 病人预交记录 Set 记录状态 = 3 Where Mod(记录性质, 10) <> 1 And 结帐id = n_结帐id;
  Select Sum(实收金额) Into n_退费金额 From 门诊费用记录 Where 结帐id = n_冲销id;
  Insert Into 病人预交记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 结算序号, 缴款组id, 预交类别, 卡类别id,
     结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算性质)
    Select 病人预交记录_Id.Nextval, a.No, a.实际票号, 4, 2, a.病人id, a.主页id, a.科室id, a.摘要, Null, d_退号时间, 操作员编号_In, 操作员姓名_In, n_退费金额,
           n_冲销id, n_结算序号, n_组id, 预交类别, Null, Null, Null, Null, Null, Null, 1, 4
    From 病人预交记录 A
    Where a.结帐id = n_结帐id And Rownum < 2;
  If Sql%NotFound Then
    v_Err_Msg := '未找到挂号单为【' || 单据号_In || '】的原始挂号记录!';
    Raise Err_Item;
  End If;

  --更新挂号序号状态
  n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111));
  If n_出诊记录id Is Null Then
    If Nvl(n_退号重用, 0) = 0 Then
      Update 挂号序号状态
      Set 状态 = 4
      Where 状态 = 1 And 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_发生时间) And Trunc(d_发生时间 + 1) - 1 / 24 / 60 / 60;
    Else
      Delete 挂号序号状态
      Where 状态 = 1 And 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_发生时间) And Trunc(d_发生时间 + 1) - 1 / 24 / 60 / 60;
    End If;
  Else
    If Nvl(n_退号重用, 0) = 0 Then
      Update 临床出诊序号控制
      Set 挂号状态 = 4
      Where 记录id = n_出诊记录id And (序号 = n_号序 Or 备注 = To_Char(n_号序));
    Else
      Update 临床出诊序号控制
      Set 挂号状态 = 0, 类型 = Null, 名称 = Null, 操作员姓名 = Null, 工作站名称 = Null
      Where 记录id = n_出诊记录id And (序号 = n_号序 Or 备注 = To_Char(n_号序));
    End If;
  End If;

  --病人挂号汇总
  For c_挂号 In (Select a.收费细目id, a.发生时间, c.接收时间, c.执行部门id, c.执行人, m.Id As 医生id, c.号别 As 号码,
                      Decode(Nvl(c.预约, 0), 0, 0, 1) As 预约
               From 门诊费用记录 A, 病人挂号记录 C, 人员表 M
               Where a.记录性质 = 4 And a.结帐id = n_冲销id And a.从属父号 Is Null And c.执行人 = m.姓名(+) And a.No = c.No And
                     Nvl(a.附加标志, 0) = 0 And Rownum < 2) Loop
  
    If c_挂号.预约 <> 0 Then
      d_Temp := Trunc(c_挂号.接收时间);
    Else
      d_Temp := Trunc(c_挂号.发生时间);
    End If;
  
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - c_挂号.预约, 已约数 = Nvl(已约数, 0) - c_挂号.预约
    Where 日期 = d_Temp And 科室id = c_挂号.执行部门id And 项目id = c_挂号.收费细目id And Nvl(医生姓名, '医生') = Nvl(c_挂号.执行人, '医生') And
          Nvl(医生id, 0) = Nvl(c_挂号.医生id, 0) And (号码 = c_挂号.号码 Or 号码 Is Null);
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收, 已约数)
      Values
        (d_Temp, c_挂号.执行部门id, c_挂号.收费细目id, c_挂号.执行人, Decode(c_挂号.医生id, 0, Null, c_挂号.医生id), c_挂号.号码, -1, -1 * c_挂号.预约,
         -1 * c_挂号.预约);
    End If;
  
    If n_出诊记录id Is Not Null Then
      Update 临床出诊记录
      Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - c_挂号.预约, 已约数 = Nvl(已约数, 0) - c_挂号.预约
      Where ID = n_出诊记录id And Nvl(已挂数, 0) > 0;
    End If;
  End Loop;

  --要删除队列
  n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
  If n_挂号生成队列 <> 0 Then
    n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
    If Nvl(n_分诊台签到排队, 0) = 0 Then
      For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
        Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
      End Loop;
    End If;
  End If;

  --医保产生的就诊登记记录
  Delete From 就诊登记记录 Where 病人id = n_病人id And 就诊时间 = d_发生时间 And 主页id Is Null;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号补结算_Delete;
/

--129255:刘涛,2018-09-27,处理出库房分批入库房不分批的批次问题
Create Or Replace Procedure Zl_药品申领_Insert
(
  No_In           In 药品收发记录.No%Type,
  序号_In         In 药品收发记录.序号%Type,
  库房id_In       In 药品收发记录.库房id%Type,
  对方部门id_In   In 药品收发记录.对方部门id%Type,
  药品id_In       In 药品收发记录.药品id%Type,
  批次_In         In 药品收发记录.批次%Type,
  填写数量_In     In 药品收发记录.填写数量%Type,
  实际数量_In     In 药品收发记录.实际数量%Type,
  成本价_In       In 药品收发记录.成本价%Type,
  成本金额_In     In 药品收发记录.成本金额%Type,
  零售价_In       In 药品收发记录.零售价%Type,
  零售金额_In     In 药品收发记录.零售金额%Type,
  差价_In         In 药品收发记录.差价%Type,
  填制人_In       In 药品收发记录.填制人%Type,
  产地_In         In 药品收发记录.产地%Type := Null,
  批号_In         In 药品收发记录.批号%Type := Null,
  效期_In         In 药品收发记录.效期%Type := Null,
  摘要_In         In 药品收发记录.摘要%Type := Null,
  填制日期_In     In 药品收发记录.填制日期%Type := Null,
  上次供应商id_In In 药品收发记录.供药单位id%Type := Null,
  批准文号_In     In 药品收发记录.批准文号%Type := Null,
  申领方式_In     In 药品收发记录.单量%Type := 0,
  结束时间_In     In 药品收发记录.频次%Type := Null
) Is
  v_Lngid        药品收发记录.Id%Type; --收发ID
  n_出库收发id   药品收发记录.Id%Type; --出库库房收发id
  v_入的类别id   药品收发记录.入出类别id%Type; --入出类别ID
  v_出的类别id   药品收发记录.入出类别id%Type; --入出类别ID
  d_上次生产日期 药品库存.上次生产日期%Type;

  v_是否分批 Integer; --判断入库是否药库分批   1:分批；0：不分批
  v_药库分批 Integer; --判断入库是否药库分批   1:分批；0：不分批
  v_药房分批 Integer; --判断入库是否药库分批   1:分批；0：不分批
  v_批次     药品收发记录.批次%Type := Null; --主要针对入库中实行药库分批的药品
Begin
  --首先找出入和出的类别ID
  Select b.Id
  Into v_入的类别id
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 6 And b.系数 = 1 And Rownum < 2;
  Select b.Id
  Into v_出的类别id
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 6 And b.系数 = -1 And Rownum < 2;

  Select 药品收发记录_Id.Nextval Into v_Lngid From Dual;

  Begin
    Select 上次生产日期
    Into d_上次生产日期
    From 药品库存
    Where 性质 = 1 And 库房id = 库房id_In And 药品id = 药品id_In And Nvl(批次, 0) = Nvl(批次_In, 0);
  Exception
    When Others Then
      d_上次生产日期 := Null;
  End;

  Select 药品收发记录_Id.Nextval Into n_出库收发id From Dual;
  --插入类别为出的那一笔
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 发药方式, 供药单位id, 批准文号, 生产日期, 单量, 频次)
  Values
    (n_出库收发id, 1, 6, No_In, 序号_In, 库房id_In, 对方部门id_In, v_出的类别id, -1, 药品id_In, 批次_In, 产地_In, 批号_In, 效期_In, 填写数量_In,
     实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, 1, 上次供应商id_In, 批准文号_In, d_上次生产日期, 申领方式_In,
     结束时间_In);

  Select Nvl(药库分批, 0), Nvl(药房分批, 0) Into v_药库分批, v_药房分批 From 药品规格 Where 药品id = 药品id_In;

  v_是否分批 := 0;
  If v_药房分批 = 0 Then
    If v_药库分批 = 1 Then
      Begin
        Select Distinct 0
        Into v_是否分批
        From 部门性质说明
        Where ((工作性质 Like '%药房') Or (工作性质 Like '制剂室')) And 部门id = 对方部门id_In;
      Exception
        When Others Then
          v_是否分批 := 1;
      End;
    End If;
  Else
    v_是否分批 := 1;
  End If;

  If v_是否分批 = 1 And Nvl(批次_In, 0) = 0 Then
    --入库分批且出库不分批
    v_批次 := Zl_Fun_Getbatchnum(药品id_In, 产地_In, 批号_In, 成本价_In, 零售价_In, v_Lngid, 上次供应商id_In);
  Elsif v_是否分批 = 0 Then
    --入库不分批
    v_批次 := 0;
  Elsif Nvl(批次_In, 0) <> 0 Then
    --入库分批且出库也分批
    v_批次 := 批次_In;
  End If;

  --插入类别为入的那一笔
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 发药方式, 供药单位id, 批准文号, 生产日期, 单量, 频次)
  Values
    (v_Lngid, 1, 6, No_In, 序号_In + 1, 对方部门id_In, 库房id_In, v_入的类别id, 1, 药品id_In, v_批次, 产地_In, 批号_In, 效期_In, 填写数量_In,
     实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, 1, 上次供应商id_In, 批准文号_In, d_上次生产日期, 申领方式_In,
     结束时间_In);

  --下库存数据
  Zl_药品库存_Update(n_出库收发id, 0);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品申领_Insert;
/

--131276:焦博,2018-09-27,根据自定义函数zl_Custom_PatiIDs_Get来获取病人ID
CREATE OR REPLACE Function Zl_Custom_Patiids_Get
(
  模块号_In   Number,
  身份证号_In 病人信息.身份证号%Type,
  姓名_In  病人信息.姓名%Type:= Null,
  性别_In  病人信息.性别%Type:= Null
) Return Varchar2 Is
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  Return NULL;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Custom_Patiids_Get;
/

--131549:冉俊明,2018-09-26,号源由固定排班方式转为月排班方式，无法添加到已发布月出诊表中
Create Or Replace Procedure Zl_临床出诊记录_Delete
(
  号源id_In   临床出诊记录.号源id%Type,
  开始日期_In 临床出诊记录.出诊日期%Type
) As
  --删除固定出诊表中某个日期之后未使用的临床出诊记录
  l_记录id t_Numlist := t_Numlist();
  n_Count  Number;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  Select Count(1)
  Into n_Count
  From 临床出诊记录 A, 病人挂号记录 D, 临床出诊安排 B, 临床出诊表 C
  Where a.Id = d.出诊记录id And a.安排id = b.Id And b.出诊id = c.Id And c.排班方式 = 0 And b.号源id = 号源id_In And a.出诊日期 >= 开始日期_In And
        Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '当前号源在' || To_Char(开始日期_In, 'yyyy-mm-dd') || '之后存在预约挂号记录，不能删除！';
    Raise Err_Item;
  End If;

  Select a.Id
  Bulk Collect
  Into l_记录id
  From 临床出诊记录 A, 临床出诊安排 B, 临床出诊表 C
  Where a.安排id = b.Id And b.出诊id = c.Id And c.排班方式 = 0 And b.号源id = 号源id_In And a.出诊日期 >= 开始日期_In;

  Zl_临床出诊记录_Batchdelete(l_记录id);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊记录_Delete;
/


--131243:余伟节,2018-09-14,身份证号校验增加台湾地区码
Create Or Replace Procedure Zl_Third_Buildpatient
(
  Patiinfo_In  In Xmltype,
  Patiinfo_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------
  --参数说明:
  -- 入参 Patiinfo_In:
  --<IN>
  --  <ZJH></ZJH>                 //证件号，目前仅支持身份证号
  --  <ZJLX></ZJLX>                       //证件类型(目前仅支持身份证,为空时默认为身份证)
  --  <XM></XM>                       //姓名
  --  <SJH></SJH>                      //手机号
  --</IN>

  --出参 Patiinfo_Out：
  --<OUTPUT>
  --       <BRID></BRID>                //病人ID
  --       <MZH></MZH>                  //门诊号
  --     <ERROR></ERROR>         //如果有错误返回该节点
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Pati_Id      病人信息.病人id%Type;
  n_Card_Type_Id 医疗卡类别.Id%Type;
  n_Count        Number(5);
  n_Sum          Number(5);
  v_校验位       Varchar2(50);

  v_姓名         病人信息.姓名%Type;
  v_身份证号     病人信息.身份证号%Type;
  v_手机号       病人信息.家庭电话%Type;
  v_性别         病人信息.性别%Type;
  v_年龄         病人信息.年龄%Type;
  v_操作员       人员表.姓名%Type;
  v_医疗付款方式 病人信息.医疗付款方式%Type;
  n_门诊号       病人信息.门诊号%Type;
  v_证件类型     医疗卡类别.名称%Type;
  v_证件号       病人医疗卡信息.卡号%Type;

  v_Pattern Varchar2(500);
  v_Temp    Varchar2(32767); --临时XML
  v_Err_Msg Varchar2(2000);
  n_存在    Number(2);

  d_出生日期  病人信息.出生日期%Type;
  d_Curr_Time Date;

  Err_Item Exception;
Begin
  Patiinfo_Out := Xmltype('<OUTPUT></OUTPUT>');
  Select Sysdate Into d_Curr_Time From Dual;

  --新建病人：姓名、身份证号、手机号（存在家庭电话中）、出生日期、性别、年龄(后面三项可从身份证中获取)。
  Select Extractvalue(Value(I), 'IN/XM'), Extractvalue(Value(I), 'IN/ZJH'), Extractvalue(Value(I), 'IN/SJH'),
         Extractvalue(Value(I), 'IN/ZJLX')
  Into v_姓名, v_证件号, v_手机号, v_证件类型
  From Table(Xmlsequence(Extract(Patiinfo_In, 'IN'))) I;

  Begin
    If v_证件类型 Is Null Then
      Select 病人id
      Into n_Pati_Id
      From 病人医疗卡信息
      Where 卡号 = v_证件号 And 卡类别id In (Select ID From 医疗卡类别 Where 名称 Like '%身份证%') And Rownum < 2;
    Else
      Select 病人id
      Into n_Pati_Id
      From 病人医疗卡信息
      Where 卡号 = v_证件号 And 卡类别id In (Select ID From 医疗卡类别 Where 名称 = v_证件类型) And Rownum < 2;
    End If;
    n_存在 := 1;
  Exception
    When Others Then
      n_存在 := 0;
  End;

  If Nvl(n_存在, 0) = 1 Then
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    Select 门诊号 Into n_门诊号 From 病人信息 Where 病人id = n_Pati_Id;
    If n_门诊号 Is Null Then
      n_门诊号 := Nextno(3);
      Update 病人信息 Set 门诊号 = n_门诊号 Where 病人id = n_Pati_Id;
    End If;
    v_Temp := '<MZH>' || n_门诊号 || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
  Else
    If v_姓名 Is Null Then
      v_Err_Msg := '传入姓名为空!';
      Raise Err_Item;
    End If;
    If v_证件类型 Like '%身份证%' Or v_证件类型 Is Null Then
      v_身份证号 := v_证件号;
    Else
      v_Err_Msg := '目前不支持身份证以外的方式建档！';
      Raise Err_Item;
    End If;
  
    If v_身份证号 Is Null Then
      v_Err_Msg := '传入身份证号为空!';
      Raise Err_Item;
    Else
      --身份证合法验证
      v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    
      --地区检验
      If Instr(v_Pattern, Substr(v_身份证号, 1, 2)) = 0 Then
        v_Err_Msg := '身份证前两位地区码不正确!';
        Raise Err_Item;
      End If;
      --身份证长度检查
      If Length(v_身份证号) = 15 Then
        --检查身份证号:15位身份证号要求全部为数字
        v_Pattern := '^\d{15}$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_身份证号, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中包含非法字符，请检查!';
          Raise Err_Item;
        End If;
        --获取性别
        If Mod(To_Number(Substr(v_身份证号, 15, 1)), 2) = 1 Then
          v_性别 := '男';
        Else
          v_性别 := '女';
        End If;
        --出生日期的合法性检查
      
        v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(v_身份证号, 7, 6), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Raise Err_Item;
        Else
          --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
          If Instr(',0229,0230,', ',' || Substr(v_身份证号, 9, 4) || ',') > 0 Then
            v_Temp     := '19' || Substr(v_身份证号, 7, 2) || '0301';
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_出生日期 := To_Date('19' || Substr(v_身份证号, 7, 6), 'yyyy-mm-dd');
          End If;
          If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '身份证中的出生日期无效，请检查!';
            Raise Err_Item;
          End If;
        End If;
      Elsif Length(v_身份证号) = 18 Then
        -- 18 位身份证号前17 位全部为数字，最后1位可为数字或x
        v_Pattern := '^\d{17}[0-9Xx]$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_身份证号, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中包含非法字符!';
          Raise Err_Item;
        End If;
        --获取性别
        If Mod(To_Number(Substr(v_身份证号, 17, 1)), 2) = 1 Then
          v_性别 := '男';
        Else
          v_性别 := '女';
        End If;
        --出生日期的合法性检查
        v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(v_身份证号, 7, 8), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Raise Err_Item;
        Else
          --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
          If Instr(',0229,0230,', ',' || Substr(v_身份证号, 11, 4) || ',') > 0 Then
            v_Temp     := Substr(v_身份证号, 7, 4) || '0301';
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_出生日期 := To_Date(Substr(v_身份证号, 7, 8), 'yyyy-mm-dd');
          End If;
          If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '身份证中的出生日期无效，请检查!';
            Raise Err_Item;
          End If;
          --计算校验位
          n_Sum     := (To_Number(Substr(v_身份证号, 1, 1)) + To_Number(Substr(v_身份证号, 11, 1))) * 7 +
                       (To_Number(Substr(v_身份证号, 2, 1)) + To_Number(Substr(v_身份证号, 12, 1))) * 9 +
                       (To_Number(Substr(v_身份证号, 3, 1)) + To_Number(Substr(v_身份证号, 13, 1))) * 10 +
                       (To_Number(Substr(v_身份证号, 4, 1)) + To_Number(Substr(v_身份证号, 14, 1))) * 5 +
                       (To_Number(Substr(v_身份证号, 5, 1)) + To_Number(Substr(v_身份证号, 15, 1))) * 8 +
                       (To_Number(Substr(v_身份证号, 6, 1)) + To_Number(Substr(v_身份证号, 16, 1))) * 4 +
                       (To_Number(Substr(v_身份证号, 7, 1)) + To_Number(Substr(v_身份证号, 17, 1))) * 2 +
                       To_Number(Substr(v_身份证号, 8, 1)) * 1 + To_Number(Substr(v_身份证号, 9, 1)) * 6 +
                       To_Number(Substr(v_身份证号, 10, 1)) * 3;
          n_Count   := Mod(n_Sum, 11);
          v_Pattern := '10X98765432';
          v_校验位  := Substr(v_Pattern, n_Count + 1, 1);
          If v_校验位 <> Upper(Substr(v_身份证号, 18, 1)) Then
            v_Err_Msg := '身份证号码不正确，请检查。';
            Raise Err_Item;
          End If;
        End If;
      Else
        v_Err_Msg := '身份证长度不对,请检查。';
        Raise Err_Item;
      End If;
    
      If Nvl(v_年龄, '_') = '_' Then
        v_年龄 := Zl_Age_Calc(0, d_出生日期, d_Curr_Time);
      End If;
    End If;
  
    Select 名称 Into v_医疗付款方式 From 医疗付款方式 Where 缺省标志 = 1;
    n_Pati_Id := Nextno(1);
    n_门诊号  := Nextno(3);
    Insert Into 病人信息
      (病人id, 姓名, 身份证号, 家庭电话, 出生日期, 性别, 年龄, 登记时间, 门诊号, 医疗付款方式, 手机号)
      Select n_Pati_Id, v_姓名, v_身份证号, v_手机号, d_出生日期, v_性别, v_年龄, d_Curr_Time, n_门诊号, v_医疗付款方式, v_手机号


      From Dual;
    --病人信息保存完后，完成医疗卡绑定（二代身份证卡类别的绑定）
    Begin
      If v_证件类型 Is Null Then
        Select ID Into n_Card_Type_Id From 医疗卡类别 Where 名称 Like '%身份证%' And Rownum < 2;
      Else
        Select ID Into n_Card_Type_Id From 医疗卡类别 Where 名称 = v_证件类型 And Rownum < 2;
      End If;
    Exception
      When No_Data_Found Then
        v_Err_Msg := '身份证卡类别不存在！';
        Raise Err_Item;
    End;
    Select b.姓名 Into v_操作员 From 上机人员表 A, 人员表 B Where a.人员id = b.Id And a.用户名 = User;
  
    Zl_医疗卡变动_Insert(11, n_Pati_Id, n_Card_Type_Id, Null, v_身份证号, '创建虚拟卡', Null, v_操作员, d_Curr_Time);
  
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    v_Temp := '<MZH>' || n_门诊号 || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
  End If;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Buildpatient;
/

--131243:余伟节,2018-09-14,身份证校验增加台湾地区码
Create Or Replace Function Zl_Fun_Checkidcard
(
  Idcard_In   In Varchar2,
  Calcdate_In In Date := Null
) Return Varchar2 Is
  -------------------------------------------------------------------------------
  --功能：身份证号码合法性校验,并返回身份证号的出生日期、性别、年龄
  --参数说明:
  -- 入参 IDcard_In:身份证号码
  --    Calcdate_In:计算日期,缺省时按系统时间
  -- 返回值：固定格式XML串
  --<OUTPUT>
  --       <BIRTHDAY></BIRTHDAY>                //出生日期
  --       <SEX></SEX>                  //性别
  --       <AGE></AGE>                //年龄
  --     <MSG></MSG>         //空串-身份证号有效(可从身份证号中获取出生日期和性别)，非空串-返回错误信息
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Count     Number(5);
  n_Sum       Number(5);
  v_校验位    Varchar2(50);
  v_Pattern   Varchar2(500);
  v_Err_Msg   Varchar2(2000);
  v_性别      Varchar2(100);
  v_年龄      Varchar2(100);
  d_Curr_Time Date;
  d_出生日期  Date;
  v_Temp      Varchar2(20);

Begin
  Select Sysdate Into d_Curr_Time From Dual;

  If Idcard_In Is Null Then
    v_Err_Msg := '传入身份证号为空!';
    Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
  Else
    --身份证合法验证
    v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    --地区检验
    If Instr(v_Pattern, Substr(Idcard_In, 1, 2)) = 0 Then
      v_Err_Msg := '身份证前两位地区码不正确!';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    --身份证长度检查
    If Length(Idcard_In) = 15 Then
      --检查身份证号:15位身份证号要求全部为数字
      v_Pattern := '^\d{15}$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中包含非法字符，请检查!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --获取性别
      If Mod(To_Number(Substr(Idcard_In, 15, 1)), 2) = 1 Then
        v_性别 := '男';
      Else
        v_性别 := '女';
      End If;
      --出生日期的合法性检查
      v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(Idcard_In, 7, 6), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中的出生日期无效，请检查!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
        If Instr(',0229,0230,', ',' || Substr(Idcard_In, 9, 4) || ',') > 0 Then
          v_Temp     := '19' || Substr(Idcard_In, 7, 2) || '0301';
          d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
        Else
          d_出生日期 := To_Date('19' || Substr(Idcard_In, 7, 6), 'yyyy-mm-dd');
        End If;
        If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Elsif Length(Idcard_In) = 18 Then
      -- 18 位身份证号前17 位全部为数字，最后1位可为数字或x
      v_Pattern := '^\d{17}[0-9Xx]$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中包含非法字符!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --获取性别
      If Mod(To_Number(Substr(Idcard_In, 17, 1)), 2) = 1 Then
        v_性别 := '男';
      Else
        v_性别 := '女';
      End If;
      --出生日期的合法性检查
      v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(Idcard_In, 7, 8), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中的出生日期无效，请检查!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
        If Instr(',0229,0230,', ',' || Substr(Idcard_In, 11, 4) || ',') > 0 Then
          v_Temp     := Substr(Idcard_In, 7, 4) || '0301';
          d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
        Else
          d_出生日期 := To_Date(Substr(Idcard_In, 7, 8), 'yyyy-mm-dd');
        End If;
        If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
        --计算校验位
        n_Sum     := (To_Number(Substr(Idcard_In, 1, 1)) + To_Number(Substr(Idcard_In, 11, 1))) * 7 +
                     (To_Number(Substr(Idcard_In, 2, 1)) + To_Number(Substr(Idcard_In, 12, 1))) * 9 +
                     (To_Number(Substr(Idcard_In, 3, 1)) + To_Number(Substr(Idcard_In, 13, 1))) * 10 +
                     (To_Number(Substr(Idcard_In, 4, 1)) + To_Number(Substr(Idcard_In, 14, 1))) * 5 +
                     (To_Number(Substr(Idcard_In, 5, 1)) + To_Number(Substr(Idcard_In, 15, 1))) * 8 +
                     (To_Number(Substr(Idcard_In, 6, 1)) + To_Number(Substr(Idcard_In, 16, 1))) * 4 +
                     (To_Number(Substr(Idcard_In, 7, 1)) + To_Number(Substr(Idcard_In, 17, 1))) * 2 +
                     To_Number(Substr(Idcard_In, 8, 1)) * 1 + To_Number(Substr(Idcard_In, 9, 1)) * 6 +
                     To_Number(Substr(Idcard_In, 10, 1)) * 3;
        n_Count   := Mod(n_Sum, 11);
        v_Pattern := '10X98765432';
        v_校验位  := Substr(v_Pattern, n_Count + 1, 1);
        If v_校验位 <> Upper(Substr(Idcard_In, 18, 1)) Then
          v_Err_Msg := '身份证号码不正确，请检查。';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Else
      v_Err_Msg := '身份证长度不对,请检查。';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    v_年龄 := Zl_Age_Calc(0, d_出生日期, Calcdate_In);
  End If;

  Return '<OUTPUT><BIRTHDAY>' || To_Char(d_出生日期, 'YYYY-MM-DD') || '</BIRTHDAY><SEX>' || v_性别 || '</SEX><AGE>' || v_年龄 || '</AGE><MSG></MSG></OUTPUT>';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Checkidcard;
/

--129869:蒋廷中,2019-01-11,新增用药清单用药配方的转出脚本
--128929:胡俊勇,2018-07-17,疾病申报反馈表字段弄错
--128920:蔡青松,2018-07-30,提取检验项目分布数据时，只提取对应标本的数据
--131161:张永康,2018-09-05,补充4张表对含LOB字段的表从远程历史库抽回
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
  v_Sqlchild Varchar2(4000);
  v_Fields   Varchar2(4000);

  v_Dblink Varchar2(30);

  Type t_Tab_Col Is Table Of Varchar2(4000) Index By Varchar2(32);
  Arr_Tab_Col t_Tab_Col;

  ---------------------------------------------
  --功能：获取表的字段字符串
  Function Getfields(v_Table In Varchar2) Return Varchar2 As
    v_Colstr Varchar2(4000);
  Begin
    If Arr_Tab_Col.Exists(v_Table) Then
      v_Colstr := Arr_Tab_Col(v_Table);
    Else
      Select f_List2str(Cast(Collect(Column_Name) As t_Strlist)) As Colsstr
      Into v_Colstr
      From (Select Column_Name From User_Tab_Columns Where Table_Name = v_Table Order By Column_Id);
    
      Arr_Tab_Col(v_Table) := v_Colstr;
    End If;
  
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
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2';
      Execute Immediate v_Sql
        Using n_Pati_Id, n_Page_Id;
    
      v_Sql := 'Delete From H' || v_Table || ' Where 病人id = :1 And 主页id = :2';
      Execute Immediate v_Sql
        Using n_Pati_Id, n_Page_Id;
    End Loop;
  End Zl_Retu_Other;

  --------------------------------------------
  --返回指定病人ID和主页的用药清单表的子过程
  --------------------------------------------
    Procedure Zl_Retu_Drug
  (
    n_Pati_Id 病案主页.病人id%Type,
    n_Page_Id 病案主页.主页id%Type
  ) As
  Begin
    v_Table  := '病人用药清单';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    --病人用药配方，在病人用药清单转出之后执行
    For P In (Select ID From H病人用药清单 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop

        v_Table := '病人用药配方';
        v_Field := '配方id';
      
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.id;
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.id;
    End Loop;
  
    Delete H病人用药清单 Where 病人id = n_Pati_Id And 主页id = n_Page_Id;
  End Zl_Retu_Drug;

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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2 ';
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
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    For P In (Select ID As 文件id From H病人护理文件 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop
      For R In (Select Column_Value
                From Table(f_Str2list('病人护理数据,病人护理打印,病人护理活动项目,病人护理要素内容,产程要素内容'))) Loop
        v_Table  := r.Column_Value;
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where 文件id = :1';
        Execute Immediate v_Sql
          Using p.文件id;
      
        If v_Table = '病人护理数据' Then
          v_Fields := Getfields('病人护理明细');
          v_Sql    := 'Insert Into 病人护理明细(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    For P In (Select ID From H病人护理记录 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop
      v_Table  := '病人护理内容';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where 记录ID = :1';
      Execute Immediate v_Sql
        Using p.Id;
    
      v_Sql := 'Delete H' || v_Table || ' Where 记录ID = :1';
      Execute Immediate v_Sql
        Using p.Id;
    End Loop;
  
    Delete H病人护理记录 Where 病人id = n_Pati_Id And 主页id = n_Page_Id;
  End Zl_Retu_Tend;

  --------------------------------------------
  --返回指定ID的病人新版电子病历记录子过程
  --------------------------------------------
  Procedure Zl_Retu_Epr(n_Rec_Id H电子病历记录.Id%Type) As
    v_Field Varchar2(100);
  Begin
    v_Table  := '电子病历记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where id = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    --病人诊断记录在Zl_Retu_Other中已转回（无病历ID外键）
    --影像报告驳回,病人医嘱报告,报告查阅记录,这几张表的数据在Zl_Retu_Order中转回医嘱后再处理
    For R In (Select Column_Value
              From Table(f_Str2list('电子病历附件,电子病历格式,电子病历内容,疾病申报记录,疾病报告反馈,疾病申报反馈'))) Loop
      v_Table := r.Column_Value;
      If v_Table = '电子病历附件' Then
        v_Field := '病历id';
      Elsif v_Table = '疾病申报反馈' Then
        v_Field := '申报id';
      Else
        v_Field := '文件id';
      End If;
      v_Fields := Getfields(v_Table);
    
      --含LOB字段的表(电子病历图形,电子病历格式,电子病历附件)，其H表是临时表，所以需直接指定dblink
      If v_Dblink Is Not Null And (v_Table = '电子病历附件' Or v_Table = '电子病历格式') Then
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                 ' From ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                 ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      If v_Table = '电子病历内容' Then
        v_Fields := Getfields('电子病历图形');
      
        If v_Dblink Is Not Null Then
          v_Sql := 'Insert Into 电子病历图形(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                   ' From 电子病历图形@' || v_Dblink ||
                   ' a Where 对象id In (Select ID From H电子病历内容 Where 文件id = :1 And 对象类型 = 5)';
        Else
          v_Sql := 'Insert Into 电子病历图形(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                   ' From H电子病历图形 Where 对象id In (Select ID From H电子病历内容 Where 文件id = :1 And 对象类型 = 5)';
        End If;
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        If v_Dblink Is Not Null Then
          v_Sql := 'Delete 电子病历图形@' || v_Dblink ||
                   ' Where 对象id In (Select ID From H电子病历内容 Where 文件id = :1 And 对象类型 = 5)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        Else
          Delete H电子病历图形 Where 对象id In (Select ID From H电子病历内容 Where 文件id = n_Rec_Id And 对象类型 = 5);
        End If;
      End If;
    
      If v_Dblink Is Not Null And (v_Table = '电子病历附件' Or v_Table = '电子病历格式') Then
        v_Sql := 'Delete ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where id = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    --以"医嘱ID,发送号"为外键的，都按医嘱ID直接转回，只需要排在"病人医嘱发送"之后即可
    --由于外键关系，"报告查阅记录"须在"病人医嘱报告"后面
    For P In (Select Column_Value
              From Table(f_Str2list('病人医嘱计价,病人医嘱状态,病人医嘱发送,病人医嘱附费,病人医嘱附件,病人医嘱执行,病人医嘱打印,输血申请记录,输血检验结果,输血申请项目,' ||
                                     '医嘱执行打印,医嘱执行时间,医嘱执行计价,执行打印记录,病人诊断医嘱,病人路径医嘱,病人医嘱报告,报告查阅记录,' ||
                                     '影像报告驳回,影像报告记录,影像报告操作记录,影像检查记录,影像申请单图像,影像收藏内容,影像危急值记录,检验标本记录,检验试剂记录,检验拒收记录,疾病阳性记录,医嘱申请单文件,病人危急值记录'))) Loop
      v_Table := p.Column_Value;
      If Instr('病人路径医嘱', v_Table) > 0 Then
        v_Field := '病人医嘱ID';
      Else
        v_Field := '医嘱ID';
      End If;
    
      v_Fields := Getfields(v_Table);
    
      If v_Dblink Is Not Null And (v_Table = '影像报告记录' Or v_Table = '医嘱申请单文件') Then
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                 ' From ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                 ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
    
      If v_Table = '病人医嘱状态' Or v_Table = '病人医嘱报告' Then
        v_Sqlchild := v_Sql;
      Else
        Execute Immediate v_Sql
          Using n_Rec_Id;
      End If;    
    
      If v_Table = '病人医嘱状态' Then
        v_Fields := Getfields('医嘱签名记录');
        v_Sql    := 'Insert Into 医嘱签名记录(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H医嘱签名记录 Where ID In (Select 签名id From H病人医嘱状态 Where 医嘱id = :1 And 签名id Is Not Null)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete H医嘱签名记录
        Where ID In (Select 签名id From H病人医嘱状态 Where 医嘱id = n_Rec_Id And 签名id Is Not Null);
      
        Execute Immediate v_Sqlchild
          Using n_Rec_Id;
      
      Elsif v_Table = '病人医嘱发送' Then
        v_Fields := Getfields('诊疗单据打印');
        v_Sql    := 'Insert Into 诊疗单据打印(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H诊疗单据打印 Where (NO, 记录性质) In (Select NO, 记录性质 From H病人医嘱发送 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete H诊疗单据打印 Where (NO, 记录性质) In (Select NO, 记录性质 From H病人医嘱发送 Where 医嘱id = n_Rec_Id);
      
      Elsif v_Table = '影像检查记录' Then
        v_Fields := Getfields('影像检查序列');
        v_Sql    := 'Insert Into 影像检查序列(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H影像检查序列 Where 检查uid In (Select 检查uid From H影像检查记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('影像检查图象');
        v_Sql    := 'Insert Into 影像检查图象(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
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
          If v_Subtable = '检验项目分布' Then
            v_Sql := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' ||
                     Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Subtable || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1) And 医嘱id=:2';
            Execute Immediate v_Sql
              Using n_Rec_Id, n_Rec_Id;
          
            v_Sql := 'Delete H' || v_Subtable || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1)  And 医嘱id=:2';
            Execute Immediate v_Sql
              Using n_Rec_Id, n_Rec_Id;
          Elsif v_Dblink Is Not Null And v_Subtable = '检验图像结果' Then
            v_Sql := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' ||
                     Replace(v_Fields, '待转出', 'Null as 待转出') || ' From ' || v_Subtable || '@' || v_Dblink || ' Where ' ||
                     v_Subfield || ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
            Execute Immediate v_Sql
              Using n_Rec_Id;
          
            v_Sql := 'Delete ' || v_Subtable || '@' || v_Dblink || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
            Execute Immediate v_Sql
              Using n_Rec_Id;
          Else
            v_Sql := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' ||
                     Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Subtable || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
            Execute Immediate v_Sql
              Using n_Rec_Id;
          
            v_Sql := 'Delete H' || v_Subtable || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
            Execute Immediate v_Sql
              Using n_Rec_Id;
          End If;
        End Loop;
      
        v_Fields := Getfields('检验普通结果');
        v_Sql    := 'Insert Into 检验普通结果(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('检验药敏结果');
        v_Sql    := 'Insert Into 检验药敏结果(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H检验药敏结果 Where 细菌结果id In (Select ID From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = :1))';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('检验质控报告');
        v_Sql    := 'Insert Into 检验质控报告(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
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
      Elsif v_Table = '病人医嘱报告' Then
        v_Fields := Getfields('医嘱报告内容');
        If v_Dblink Is Not Null Then
          v_Sql := 'Insert Into 医嘱报告内容(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                   ' From 医嘱报告内容@' || v_Dblink || ' Where ID In (Select 报告id From 病人医嘱报告@' || v_Dblink ||
                   ' Where 医嘱id = :1 And 报告id Is Not Null)';
        Else
          v_Sql := 'Insert Into 医嘱报告内容(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                   ' From H医嘱报告内容 Where ID In (Select 报告id From H病人医嘱报告 Where 医嘱id = :1 And 报告id Is Not Null)';
        End If;
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        If v_Dblink Is Not Null Then
          v_Sql := 'Delete 医嘱报告内容@' || v_Dblink || ' Where ID In (Select 报告id From 病人医嘱报告@' || v_Dblink ||
                   ' Where 医嘱id = :1 And 报告id Is Not Null);';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        Else
          Delete H医嘱报告内容
          Where ID In (Select 报告id From H病人医嘱报告 Where 医嘱id = n_Rec_Id And 报告id Is Not Null);
        End If;
      
        Execute Immediate v_Sqlchild
          Using n_Rec_Id;
      Elsif v_Table = '病人危急值记录' Then
      
        v_Fields := Getfields('病人危急值病历');
        v_Sql    := 'Insert Into 病人危急值病历(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H病人危急值病历 Where 危急值id In (Select ID From H病人危急值记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete H病人危急值病历 Where 危急值id In (Select ID From H病人危急值记录 Where 医嘱id = n_Rec_Id);
      
        v_Fields := Getfields('病人危急值医嘱');
        v_Sql    := 'Insert Into 病人危急值医嘱(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H病人危急值医嘱 Where 危急值id In (Select ID From H病人危急值记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H病人危急值医嘱 Where 危急值id In (Select ID From H病人危急值记录 Where 医嘱id = n_Rec_Id);
      
      End If;
    
      If v_Dblink Is Not Null And (v_Table = '影像报告记录' Or v_Table = '医嘱申请单文件') Then
        v_Sql := 'Delete ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
    
      Execute Immediate v_Sql
        Using n_Rec_Id;
    End Loop;
  
    --手麻数据
    If n_Opersystem > 0 Then
      Execute Immediate 'Call zl24_Retu_Oper(:1)'
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

  Select Max(Db连接) Into v_Dblink From zlBakSpaces Where 系统 = 100 And 当前 = 1;

  --对基于视图的转储方案进行了只读判断.
  n_Opersystem := 0;
  Select 编号 Into n_Opersystem From zlSystems Where Upper(所有者) = Zl_Owner And 编号 Like '24%';
  If n_Opersystem > 0 Then
    Begin
      Select Nvl(只读, 0) Into n_只读 From zlBakSpaces Where 系统 = n_Opersystem And 当前 = 1;
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
    --抽回未结记帐费用
    Zl_Retu_Exes(n_Patiid, 8);
  
    v_Table  := '病人挂号记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where NO =:1 ';
    Execute Immediate v_Sql
      Using v_Times;
  
    For r_Other In (Select ID, 病人id From H病人挂号记录 Where NO = v_Times) Loop
      Zl_Retu_Other(r_Other.病人id, r_Other.Id);
    End Loop;
  
    For r_Epr In (Select b.Id
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where NO =:1';
    Execute Immediate v_Sql
      Using v_Times;
  
    Delete H病人转诊记录 Where NO = v_Times;
    Delete H病人挂号记录 Where NO = v_Times;
  
    --2.住院病人，按病人ID和主页ID抽回
  Elsif n_Flag = 1 Then
    --抽回未结记帐费用
    Zl_Retu_Exes(n_Patiid || ',' || v_Times, 8);
  
    Zl_Retu_Other(n_Patiid, To_Number(v_Times));
    Zl_Retu_Path(n_Patiid, To_Number(v_Times));
    Zl_Retu_Drug(n_Patiid, To_Number(v_Times));
  
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

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM || ':' || v_Sql);
End Zl_Retu_Clinic;
/

--101765:秦龙,2018-08-27,增加传参“是否辅助用药”
Create Or Replace Procedure Zl_草药品种_Update
(
  分类id_In       In 诊疗项目目录.分类id%Type := Null,
  Id_In           In 诊疗项目目录.Id%Type,
  编码_In         In 诊疗项目目录.编码%Type := Null,
  名称_In         In 诊疗项目目录.名称%Type := Null,
  拼音_In         In 诊疗项目别名.简码%Type := Null,
  五笔_In         In 诊疗项目别名.简码%Type := Null,
  英文_In         In 诊疗项目别名.名称%Type := Null,
  单位_In         In 诊疗项目目录.计算单位%Type := Null,
  毒理分类_In     In 药品特性.毒理分类%Type := Null,
  价值分类_In     In 药品特性.价值分类%Type := Null,
  货源情况_In     In 药品特性.货源情况%Type := Null,
  用药梯次_In     In 药品特性.用药梯次%Type := Null,
  药品类型_In     In 药品特性.药品类型%Type := Null,
  处方职务_In     In 药品特性.处方职务%Type := '00',
  处方限量_In     In 药品特性.处方限量%Type := Null,
  单独应用_In     In 诊疗项目目录.单独应用%Type := Null,
  是否原料_In     In 药品特性.是否原料%Type := 0,
  适用性别_In     In 诊疗项目目录.适用性别%Type := 0,
  参考目录id_In   In 诊疗项目目录.参考目录id%Type := Null,
  其他别名_In     In Varchar2 := Null, --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织
  自管药_In       In Number := Null,
  是否辅助用药_In In 药品特性.是否辅助用药%Type := 0
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
  Set 分类id = 分类id_In, 编码 = 编码_In, 名称 = 名称_In, 计算单位 = 单位_In, 参考目录id = 参考目录id_In, 适用性别 = 适用性别_In, 单独应用 = 单独应用_In
  Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;

  Update 药品特性
  Set 毒理分类 = 毒理分类_In, 价值分类 = 价值分类_In, 货源情况 = 货源情况_In, 用药梯次 = 用药梯次_In, 药品类型 = 药品类型_In, 处方职务 = 处方职务_In, 处方限量 = 处方限量_In,
      是否原料 = 是否原料_In, 临床自管药 = 自管药_In, 是否辅助用药 = 是否辅助用药_In
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
End Zl_草药品种_Update;
/

--101765:秦龙,2018-08-27,增加传参“是否辅助用药”
Create Or Replace Procedure Zl_草药品种_Insert
(
  类别_In         In 诊疗项目目录.类别%Type := Null,
  分类id_In       In 诊疗项目目录.分类id%Type := Null,
  Id_In           In 诊疗项目目录.Id%Type,
  编码_In         In 诊疗项目目录.编码%Type := Null,
  名称_In         In 诊疗项目目录.名称%Type := Null,
  拼音_In         In 诊疗项目别名.简码%Type := Null,
  五笔_In         In 诊疗项目别名.简码%Type := Null,
  英文_In         In 诊疗项目别名.名称%Type := Null,
  单位_In         In 诊疗项目目录.计算单位%Type := Null,
  毒理分类_In     In 药品特性.毒理分类%Type := Null,
  价值分类_In     In 药品特性.价值分类%Type := Null,
  货源情况_In     In 药品特性.货源情况%Type := Null,
  用药梯次_In     In 药品特性.用药梯次%Type := Null,
  药品类型_In     In 药品特性.药品类型%Type := Null,
  处方职务_In     In 药品特性.处方职务%Type := '00',
  处方限量_In     In 药品特性.处方限量%Type := Null,
  单独应用_In     In 诊疗项目目录.单独应用%Type := Null,
  是否原料_In     In 药品特性.是否原料%Type := 0,
  适用性别_In     In 诊疗项目目录.适用性别%Type := 0,
  参考目录id_In   In 诊疗项目目录.参考目录id%Type := Null,
  其他别名_In     In Varchar2 := Null, --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织
  自管药_In       In Number := Null,
  是否辅助用药_In In 药品特性.是否辅助用药%Type := 0
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
    (类别_In, 分类id_In, Id_In, 编码_In, 名称_In, 单位_In, 1, 0, 单独应用_In, 0, 0, 0, 3, Sysdate,
     To_Date('3000-01-01', 'YYYY-MM-DD'), 参考目录id_In, 适用性别_In);

  Insert Into 药品特性
    (药名id, 药品剂型, 毒理分类, 价值分类, 货源情况, 用药梯次, 药品类型, 处方职务, 处方限量, 急救药否, 是否新药, 是否原料, 是否皮试, 临床自管药, 是否辅助用药)
  Values
    (Id_In, '方剂', 毒理分类_In, 价值分类_In, 货源情况_In, 用药梯次_In, 药品类型_In, 处方职务_In, 处方限量_In, 0, 0, 是否原料_In, 0, 自管药_In, 是否辅助用药_In);

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
End Zl_草药品种_Insert;
/

--101765:秦龙,2018-08-27,增加传参“是否辅助用药”
Create Or Replace Procedure Zl_成药品种_Update
(
  分类id_In       In 诊疗项目目录.分类id%Type := Null,
  Id_In           In 诊疗项目目录.Id%Type,
  编码_In         In 诊疗项目目录.编码%Type := Null,
  名称_In         In 诊疗项目目录.名称%Type := Null,
  拼音_In         In 诊疗项目别名.简码%Type := Null,
  五笔_In         In 诊疗项目别名.简码%Type := Null,
  英文_In         In 诊疗项目别名.名称%Type := Null,
  单位_In         In 诊疗项目目录.计算单位%Type := Null,
  药品剂型_In     In 药品特性.药品剂型%Type := Null,
  毒理分类_In     In 药品特性.毒理分类%Type := Null,
  价值分类_In     In 药品特性.价值分类%Type := Null,
  货源情况_In     In 药品特性.货源情况%Type := Null,
  用药梯次_In     In 药品特性.用药梯次%Type := Null,
  药品类型_In     In 药品特性.药品类型%Type := Null,
  处方职务_In     In 药品特性.处方职务%Type := '00',
  处方限量_In     In 药品特性.处方限量%Type := Null,
  急救药否_In     In 药品特性.急救药否%Type := 0,
  是否新药_In     In 药品特性.是否新药%Type := 0,
  是否原料_In     In 药品特性.是否原料%Type := 0,
  是否皮试_In     In 药品特性.是否皮试%Type := 0,
  抗生素_In       In 药品特性.抗生素%Type := 0,
  参考目录id_In   In 诊疗项目目录.参考目录id%Type := Null,
  品种医嘱_In     In 药品特性.品种医嘱%Type := 0,
  适用性别_In     In 诊疗项目目录.适用性别%Type := 0,
  其他别名_In     In Varchar2 := Null, --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织 
  自管药_In       In Number := Null,
  Atccode_In      In Varchar2 := Null,
  肿瘤药_In       In 药品特性.是否肿瘤药%Type := 0,
  溶媒_In         In 药品特性.溶媒%Type := 0,
  是否原研药_In   In 药品特性.是否原研药%Type := 0,
  是否专利药_In   In 药品特性.是否专利药%Type := 0,
  是否单独定价_In In 药品特性.是否单独定价%Type := 0,
  是否辅助用药_In In 药品特性.是否辅助用药%Type := 0
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
      是否单独定价 = 是否单独定价_In, 是否辅助用药 = 是否辅助用药_In
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

--101765:秦龙,2018-08-27,增加传参“是否辅助用药”
Create Or Replace Procedure Zl_成药品种_Insert
(
  类别_In         In 诊疗项目目录.类别%Type := Null,
  分类id_In       In 诊疗项目目录.分类id%Type := Null,
  Id_In           In 诊疗项目目录.Id%Type,
  编码_In         In 诊疗项目目录.编码%Type := Null,
  名称_In         In 诊疗项目目录.名称%Type := Null,
  拼音_In         In 诊疗项目别名.简码%Type := Null,
  五笔_In         In 诊疗项目别名.简码%Type := Null,
  英文_In         In 诊疗项目别名.名称%Type := Null,
  单位_In         In 诊疗项目目录.计算单位%Type := Null,
  药品剂型_In     In 药品特性.药品剂型%Type := Null,
  毒理分类_In     In 药品特性.毒理分类%Type := Null,
  价值分类_In     In 药品特性.价值分类%Type := Null,
  货源情况_In     In 药品特性.货源情况%Type := Null,
  用药梯次_In     In 药品特性.用药梯次%Type := Null,
  药品类型_In     In 药品特性.药品类型%Type := Null,
  处方职务_In     In 药品特性.处方职务%Type := '00',
  处方限量_In     In 药品特性.处方限量%Type := Null,
  急救药否_In     In 药品特性.急救药否%Type := 0,
  是否新药_In     In 药品特性.是否新药%Type := 0,
  是否原料_In     In 药品特性.是否原料%Type := 0,
  是否皮试_In     In 药品特性.是否皮试%Type := 0,
  抗生素_In       In 药品特性.抗生素%Type := 0,
  参考目录id_In   In 诊疗项目目录.参考目录id%Type := Null,
  品种医嘱_In     In 药品特性.品种医嘱%Type := 0,
  适用性别_In     In 诊疗项目目录.适用性别%Type := 0,
  其他别名_In     In Varchar2 := Null, --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织 
  自管药_In       In Number := Null,
  Atccode_In      In Varchar2 := Null,
  肿瘤药_In       In 药品特性.是否肿瘤药%Type := 0,
  溶媒_In         In 药品特性.溶媒%Type := 0,
  是否原研药_In   In 药品特性.是否原研药%Type := 0,
  是否专利药_In   In 药品特性.是否专利药%Type := 0,
  是否单独定价_In In 药品特性.是否单独定价%Type := 0,
  是否辅助用药_In In 药品特性.是否辅助用药%Type := 0
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
     是否原研药, 是否专利药, 是否单独定价, 是否辅助用药)
  Values
    (Id_In, 药品剂型_In, 毒理分类_In, 价值分类_In, 货源情况_In, 用药梯次_In, 药品类型_In, 处方职务_In, 处方限量_In, 急救药否_In, 是否新药_In, 抗生素_In, 是否原料_In,
     是否皮试_In, 品种医嘱_In, 自管药_In, Atccode_In, 肿瘤药_In, 溶媒_In, 是否原研药_In, 是否专利药_In, 是否单独定价_In, 是否辅助用药_In);

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

--73706:冉俊明,2018-08-01,体检费用不更新病人余额.费用余额
Create Or Replace Procedure Zl_记帐记录_发料审核
(
  Billid_In     药品收发记录.Id%Type, --药品收发记录ID串,格式:id
  Id_In         住院费用记录.Id%Type,
  No_In         住院费用记录.No%Type,
  病人id_In     住院费用记录.病人id%Type,
  主页id_In     住院费用记录.主页id%Type,
  病人病区id_In 住院费用记录.病人病区id%Type,
  病人科室id_In 住院费用记录.病人科室id%Type,
  开单部门id_In 住院费用记录.开单部门id%Type,
  执行部门id_In 住院费用记录.执行部门id%Type,
  收入项目id_In 住院费用记录.收入项目id%Type,
  实收金额_In   住院费用记录.实收金额%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  门诊标志_In   住院费用记录.门诊标志%Type,
  审核时间_In   住院费用记录.登记时间%Type := Null
) As
  --功能：审核一张记帐划价单
  --参数：
  --    审核时间_IN：用于部份需要统一控制或返回时间的地方
  n_类别 Number;
  v_Date Date;
Begin
  If 审核时间_In Is Null Then
    Select Sysdate Into v_Date From Dual;
  Else
    v_Date := 审核时间_In;
  End If;

  If 门诊标志_In = 1 Or 门诊标志_In = 4 Then
    Update 门诊费用记录
    Set 记录状态 = 1, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 登记时间 = v_Date --已产生的药品记录的时间不变
    Where ID = Id_In;
  Else
    Update 住院费用记录
    Set 记录状态 = 1, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 登记时间 = v_Date --已产生的药品记录的时间不变
    Where ID = Id_In;
  End If;

  --药品收发记录.填制日期
  Update 药品收发记录
  Set 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date)
  Where ID = Billid_In;

  If Nvl(门诊标志_In, 0) = 1 Or Nvl(门诊标志_In, 0) = 2 Then
    n_类别 := 门诊标志_In;
  Elsif Nvl(主页id_In, 0) = 0 Or Nvl(门诊标志_In, 0) = 4 Then
    n_类别 := 1;
  Else
    n_类别 := 2;
  End If;

  --病人余额
  If Nvl(门诊标志_In, 0) <> 4 Then
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And 性质 = 1 And 类型 = n_类别;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, n_类别, 实收金额_In, 0);
    End If;
  End If;

  --病人未结费用
  Update 病人未结费用
  Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
  Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(病人病区id_In, 0) And
        Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And
        Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And 来源途径 + 0 = 门诊标志_In;

  If Sql%RowCount = 0 Then
    Insert Into 病人未结费用
      (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
    Values
      (病人id_In, 主页id_In, 病人病区id_In, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 门诊标志_In, Nvl(实收金额_In, 0));
  End If;

  --库房中的药品已全部审核则标为已收费
  If 门诊标志_In = 1 Or 门诊标志_In = 4 Then
    Update 未发药品记录
    Set 已收费 = 1, 填制日期 = v_Date
    Where NO = No_In And 单据 = 25 And Nvl(已收费, 0) = 0 And
          Nvl(库房id, 0) Not In (Select Distinct Nvl(执行部门id, 0)
                               From 门诊费用记录
                               Where 记录性质 = 2 And NO = No_In And 收费类别 = '4' And 记录状态 = 0);
  Else
    Update 未发药品记录
    Set 已收费 = 1, 填制日期 = v_Date
    Where NO = No_In And 单据 = 25 And Nvl(已收费, 0) = 0 And
          Nvl(库房id, 0) Not In (Select Distinct Nvl(执行部门id, 0)
                               From 住院费用记录
                               Where 记录性质 = 2 And NO = No_In And 收费类别 = '4' And 记录状态 = 0);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_记帐记录_发料审核;
/

--135887:刘兴洪,2018-12-20,解决结帐作废时三方卡退回了预交的问是
--73706:冉俊明,2018-08-01,体检费用不更新病人余额.费用余额
Create Or Replace Procedure Zl_病人结帐记录_Delete
(
  No_In           病人结帐记录.No%Type,
  操作员编号_In   病人结帐记录.操作员编号%Type,
  操作员姓名_In   病人结帐记录.操作员姓名%Type,
  误差金额_In     病人预交记录.冲预交%Type := 0, --医保或预交退现金产生的误差
  结帐作废结算_In Varchar2 := Null, --结算方式|结算金额|结算号码||......
  预交退现金_In   Number := 0, --当预交款退现金时，结算方式及金额通过参数结帐作废结算_In传入
  冲销id_In       病人预交记录.结帐id%Type := Null,
  冲销时间_In     Date := Null,
  缴预交id_In     病人预交记录.Id%Type := Null, --在作废时将相关的金额充值到预交款时填写
  票据号_In       病人结帐记录.实际票号%Type := Null,
  领用id_In       票据领用记录.Id%Type := Null,
  票种_In         票据使用明细.票种%Type := Null
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  --该游标用于预交记录相关信息
  Cursor c_Deposit(v_Id 病人预交记录.结帐id%Type) Is
    Select 病人id, 记录性质, 结算方式, 冲预交, 预交类别 From 病人预交记录 Where 结帐id = v_Id;
  r_Depositrow c_Deposit%RowType;

  --该游标用于处理费用相关汇总表
  Cursor c_Money(v_Id 病人预交记录.结帐id%Type) Is
    Select NO, 开单部门id, 病人科室id, 执行部门id, 病人病区id, 病人id, 主页id, 收入项目id, 门诊标志, 结帐金额
    From 住院费用记录
    Where 结帐id = v_Id
    Union All
    Select NO, 开单部门id, 病人科室id, 执行部门id, 0 As 病人病区id, 病人id, 0 As 主页id, 收入项目id, 门诊标志, 结帐金额
    From 门诊费用记录
    Where 结帐id = v_Id;

  r_Moneyrow c_Money%RowType;

  --该游标包含病人的相关信息
  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, b.主页id, b.出院病床, b.当前病区id, b.出院科室id, Nvl(b.费别, a.费别) As 费别, a.险类, c.编码 As 付款方式
    From 病人信息 A, 病案主页 B, 医疗付款方式 C
    Where a.病人id = n_病人id And a.病人id = b.病人id(+) And Nvl(a.主页id, 0) = b.主页id(+) And a.医疗付款方式 = c.名称(+);
  r_Pati c_Pati%RowType;

  --过程变量
  v_结算内容 Varchar2(500);
  v_当前结算 Varchar2(50);
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_实际票号 病人预交记录.实际票号%Type;
  v_误差no   住院费用记录.No%Type;
  v_误差     结算方式.名称%Type;
  n_病人id   病人信息.病人id%Type;

  n_原id   病人结帐记录.Id%Type;
  n_结帐id 病人结帐记录.Id%Type;
  n_打印id 票据打印内容.Id%Type;

  n_来源     Number; --1-门诊;2-住院;3-门诊和住院
  n_返回值   病人余额.预交余额%Type;
  n_组id     财务缴款分组.Id%Type;
  n_预交类别 Number;
  d_Date     Date;
  n_预交id   病人预交记录.Id%Type;
  n_卡结算id 病人结帐记录.Id%Type;
  v_打印id   票据打印内容.Id%Type;

  n_预交合计 病人预交记录.冲预交%Type;
  n_结帐合计 住院费用记录.结帐金额%Type;

Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  Select 名称 Into v_误差 From 结算方式 Where 性质 = 9 And Rownum = 1;

  Begin
    Select ID, 病人id, 实际票号 Into n_原id, n_病人id, v_实际票号 From 病人结帐记录 Where 记录状态 = 1 And NO = No_In;
    --最后一次打印的内容
    Select Max(ID)
    Into n_打印id
    From (Select b.Id
           From 票据使用明细 A, 票据打印内容 B
           Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 3 And b.No = No_In
           Order By a.使用时间 Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      Begin
        v_Err_Msg := '没有发现要作废的结帐单据,可能已经作废！';
        Raise Err_Item;
      End;
  End;

  Open c_Pati(n_病人id);
  Fetch c_Pati
    Into r_Pati; --体检系统调用此过程,团体结帐时没有病人信息

  d_Date := 冲销时间_In;
  If d_Date Is Null Then
    Select Sysdate Into d_Date From Dual;
  End If;

  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 3, No_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, 票种_In, 票据号_In, 1, 6, 领用id_In, v_打印id, d_Date, 操作员姓名_In);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  End If;

  n_结帐id := 冲销id_In;
  If Nvl(n_结帐id, 0) = 0 Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  End If;

  --病人结帐记录
  Insert Into 病人结帐记录
    (ID, NO, 实际票号, 记录状态, 中途结帐, 病人id, 操作员编号, 操作员姓名, 开始日期, 结束日期, 收费时间, 备注, 原因, 缴款组id, 结帐类型)
    Select n_结帐id, NO, 实际票号, 2, 中途结帐, 病人id, 操作员编号_In, 操作员姓名_In, 开始日期, 结束日期, d_Date, 备注, 原因, n_组id, 结帐类型
    From 病人结帐记录
    Where ID = n_原id;

  Update 病人结帐记录 Set 记录状态 = 3 Where ID = n_原id;

  --作废收回票据(可能以前没有使用票据,无法收回)
  If n_打印id Is Not Null Then
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
      Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
      From 票据使用明细
      Where 打印id = n_打印id And 票种 In (1, 3) And 性质 = 1;
  End If;

  --病人预交记录(冲预交及缴款)
  If 结帐作废结算_In Is Null Then
    --插入普通的结算信息(不包含三方退款
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
      Select 病人预交记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 病人id, 主页id, 科室id, Null,
             结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date, 操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 合作单位, 2
      From 病人预交记录
      Where 结帐id = n_原id And Mod(记录性质, 10) <> 1 And
            ((卡类别id Is Not Null And Nvl(冲预交, 0) > 0) Or (Mod(记录性质, 10) <> 1 And 卡类别id Is Null));
  
    --插入预交，但不包含已经退的预交
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
      Select 病人预交记录_Id.Nextval, NO, 实际票号, 11 As 记录性质, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
             d_Date, 操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2
      From (Select a.No, Sum(a.冲预交) As 冲预交, Max(a.实际票号) As 实际票号, Max(a.记录状态) As 记录状态, Max(a.病人id) As 病人id,
                    Max(a.主页id) As 主页id, Max(a.科室id) As 科室id, Max(a.结算方式) As 结算方式, Max(a.结算号码) As 结算号码, Max(a.摘要) As 摘要,
                    Max(a.缴款单位) As 缴款单位, Max(a.单位开户行) As 单位开户行, Max(a.单位帐号) As 单位帐号, Max(a.预交类别) As 预交类别,
                    Max(a.卡类别id) As 卡类别id, Max(a.结算卡序号) As 结算卡序号, Max(a.卡号) As 卡号, Max(a.交易流水号) As 交易流水号,
                    Max(a.交易说明) As 交易说明, Max(a.合作单位) As 合作单位
             From (Select a.No, a.冲预交, a.实际票号, a.记录状态, a.病人id, a.主页id, a.科室id, a.结算方式, a.结算号码, a.摘要, a.缴款单位, a.单位开户行,
                           a.单位帐号, a.预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号, a.交易说明, a.合作单位
                    From 病人预交记录 A
                    Where a.结帐id = n_原id And Mod(记录性质, 10) = 1
                    Union All
                    Select b.No, -1 * a.金额 As 冲预交, b.实际票号, b.记录状态, b.病人id, b.主页id, b.科室id, b.结算方式, b.结算号码, b.摘要, b.缴款单位,
                           b.单位开户行, b.单位帐号, b.预交类别, b.卡类别id, b.结算卡序号, b.卡号, b.交易流水号, b.交易说明, b.合作单位

                    From 三方退款信息 A, 病人预交记录 B
                    Where a.结帐id = n_原id And a.记录id = b.Id) A
             Group By a.No) A
      Where Nvl(a.冲预交, 0) <> 0;
  
    --消费卡处理
    For c_消费卡结算 In (Select a.Id, a.结算卡序号, Nvl(b.名称, '消费卡') As 卡名称
                    From 病人预交记录 A, 卡消费接口目录 B
                    Where a.结算卡序号 = b.编号(+) And a.结帐id = n_原id And Nvl(a.结算卡序号, 0) <> 0) Loop
      Select ID
      Into n_预交id
      From 病人预交记录
      Where 结帐id = n_结帐id And 结算卡序号 = Nvl(c_消费卡结算.结算卡序号, 0);
    
      For c_消费卡 In (Select a.Id, a.接口编号, a.消费卡id, a.序号, a.记录状态, a.结算方式, a.结算金额, a.卡号, a.交易流水号, a.交易时间, a.备注, a.结算标志,
                           b.停用日期, b.回收时间
                    From 病人卡结算记录 A, 消费卡目录 B
                    Where a.消费卡id = b.Id(+) And
                          a.Id In (Select 卡结算id From 病人卡结算对照 Where 预交id = c_消费卡结算.Id)) Loop
      
        If Nvl(c_消费卡.停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || c_消费卡.卡号 || '"的' || c_消费卡结算.卡名称 || '已经被他人停用，不能再进行结帐作废,请检查！';
          Raise Err_Item;
        End If;
      
        If c_消费卡.回收时间 < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || c_消费卡.卡号 || '"的' || c_消费卡结算.卡名称 || '已经回收，不能退费,请检查！';
          Raise Err_Item;
        End If;
        Select 病人卡结算记录_Id.Nextval Into n_卡结算id From Dual;
      
        Insert Into 病人卡结算记录
          (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
        Values
          (n_卡结算id, c_消费卡结算.结算卡序号, c_消费卡.消费卡id, c_消费卡.序号, 2, c_消费卡.结算方式, -1 * c_消费卡.结算金额, c_消费卡.卡号, c_消费卡.交易流水号, d_Date,
           Null, 0);
      
        Insert Into 病人卡结算对照 (预交id, 卡结算id) Values (n_预交id, n_卡结算id);
      
        Update 消费卡目录 Set 余额 = 余额 + c_消费卡.结算金额 Where ID = c_消费卡.消费卡id;
        If Sql%NotFound Then
          v_Err_Msg := '卡号为' || c_消费卡.卡号 || '的' || c_消费卡结算.卡名称 || '未找到!';
          Raise Err_Item;
        End If;
      End Loop;
    End Loop;
  
  Else
    --1.先处理冲预交部分
    If 预交退现金_In = 0 Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 病人id, 主页id, 科室id,
               Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date, 操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2
        From 病人预交记录
        Where 结帐id = n_原id And 记录性质 In (1, 11) And Nvl(冲预交, 0) <> 0;
    End If;
  
    --2.再处理结帐结算,包括医保和非医保
    v_结算内容 := 结帐作废结算_In || ' ||'; --以空格分开以|结尾,没有结算号码的
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算号码 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, No_In, v_实际票号, 12, 1, n_病人id, r_Pati.主页id, r_Pati.出院科室id, Null, v_结算方式, v_结算号码, '结帐作废退款',
         Null, Null, Null, d_Date, 操作员姓名_In, 操作员编号_In, -1 * n_结算金额, n_结帐id, n_组id, 2);
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;
  --确定结帐的费用记录来源
  Begin
    Select Case
             When Nvl(Max(住院), 0) = 1 And Nvl(Max(门诊), 0) = 1 Then
              3
             When Nvl(Max(住院), 0) = 1 Then
              2
             Else
              1
           End
    Into n_来源
    From (Select 1 As 住院, 0 As 门诊
           From 住院费用记录
           Where 结帐id = n_原id And Rownum = 1
           Union All
           Select 0 As 住院, 1 As 门诊
           From 门诊费用记录
           Where 结帐id = n_原id And Rownum = 1);
  
  Exception
    When Others Then
      n_来源 := 3;
  End;

  If 误差金额_In <> 0 Then
    Update 病人预交记录
    Set 冲预交 = 冲预交 + 误差金额_In
    Where NO = No_In And 记录性质 = 12 And 记录状态 = 1 And 结帐id = n_结帐id And 结算方式 = v_误差;
    If Sql%RowCount = 0 Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, No_In, v_实际票号, 12, 1, n_病人id, r_Pati.主页id, r_Pati.出院科室id, Null, v_误差, Null, '结帐作废退款', Null,
         Null, Null, d_Date, 操作员姓名_In, 操作员编号_In, 误差金额_In, n_结帐id, n_组id, 2);
    End If;
  End If;

  If n_来源 = 2 Or n_来源 = 3 Then
    --作废结帐对应的费用记录:不包含原始结帐产生的误差项目
    Insert Into 住院费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id,
       病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id,
       开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要,
       缴款组id, 医疗小组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 序号, 从属父号, 价格父号, 多病人单,
             记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次,
             加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人,
             执行时间, 操作员姓名, 操作员编号, -1 * 结帐金额, n_结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id, 医疗小组id
      From 住院费用记录
      Where 结帐id = n_原id;
  End If;

  If n_来源 = 1 Or n_来源 = 3 Then
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别,
       收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
       执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 序号, 从属父号, 价格父号, 记帐单id,
             病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费,
             记帐费用, 收入项目id, 收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号,
             -1 * 结帐金额, n_结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id
      From 门诊费用记录
      Where 结帐id = n_原id;
  End If;
  --相关汇总表处理
  For r_Depositrow In c_Deposit(n_结帐id) Loop
    If r_Depositrow.记录性质 In (1, 11) Then
    
      --病人余额(预交)
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - r_Depositrow.冲预交 --注:新的结帐ID产生的是负数金额
      Where 病人id = r_Depositrow.病人id And 类型 = Nvl(r_Depositrow.预交类别, 2) And 性质 = 1
      Returning 预交余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (r_Depositrow.病人id, 1, Nvl(r_Depositrow.预交类别, 2), -1 * r_Depositrow.冲预交, 0);
        n_返回值 := -1 * r_Depositrow.冲预交;
      End If;
    
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额
        Where 性质 = 1 And 病人id = r_Depositrow.病人id And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    
    Else
      --人员缴款余额,医保不支持作废的结算方式在新的预交结算中已被处理为了退现金,
      --此处用加,表示收回退给病人的现金(结帐时,退款是负,作废时是正)
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + r_Depositrow.冲预交
      Where 收款员 = 操作员姓名_In And 结算方式 = r_Depositrow.结算方式 And 性质 = 1
      Returning 余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Depositrow.结算方式, 1, r_Depositrow.冲预交);
        n_返回值 := -1 * r_Depositrow.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 结算方式 = r_Depositrow.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
      End If;
    End If;
  End Loop;

  For r_Moneyrow In c_Money(n_结帐id) Loop
    --病人余额 ,误差项已结帐,所以不需要更新这两个汇总表
    If Nvl(v_误差no, 'sc') <> Nvl(r_Moneyrow.No, 'sc') Then
      If Nvl(r_Moneyrow.门诊标志, 0) = 1 Or Nvl(r_Moneyrow.门诊标志, 0) = 2 Then
        n_预交类别 := r_Moneyrow.门诊标志;
      Elsif Nvl(r_Moneyrow.主页id, 0) = 0 Or Nvl(r_Moneyrow.门诊标志, 0) = 4 Then
        --体检:门诊病人
        n_预交类别 := 1;
      Else
        n_预交类别 := 2;
      End If;
    
      If Nvl(r_Moneyrow.门诊标志, 0) <> 4 Then
        Update 病人余额
        Set 费用余额 = Nvl(费用余额, 0) - r_Moneyrow.结帐金额 --注:新的结帐ID产生的是负数金额
        Where 病人id = r_Moneyrow.病人id And 类型 = n_预交类别 And 性质 = 1
        Returning 费用余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 性质, 类型, 预交余额, 费用余额)
          Values
            (r_Moneyrow.病人id, 1, n_预交类别, 0, -1 * r_Moneyrow.结帐金额);
          n_返回值 := -1 * r_Moneyrow.结帐金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete 病人余额
          Where 病人id = r_Moneyrow.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - r_Moneyrow.结帐金额
      Where 病人id = r_Moneyrow.病人id And Nvl(主页id, 0) = Nvl(r_Moneyrow.主页id, 0) And
            Nvl(病人病区id, 0) = Nvl(r_Moneyrow.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Moneyrow.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(r_Moneyrow.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(r_Moneyrow.执行部门id, 0) And
            收入项目id + 0 = r_Moneyrow.收入项目id And 来源途径 + 0 = r_Moneyrow.门诊标志;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (r_Moneyrow.病人id, Decode(r_Moneyrow.主页id, Null, Null, 0, Null, r_Moneyrow.主页id),
           Decode(r_Moneyrow.病人病区id, Null, Null, 0, Null, r_Moneyrow.病人病区id), r_Moneyrow.病人科室id, r_Moneyrow.开单部门id,
           r_Moneyrow.执行部门id, r_Moneyrow.收入项目id, r_Moneyrow.门诊标志, -1 * r_Moneyrow.结帐金额);
      End If;
    End If;
  End Loop;

  If Nvl(缴预交id_In, 0) <> 0 Then
    --作废时将退款金额充值到预交款帐户,这里标明是本次结帐缴存的
    Update 病人预交记录 Set 结帐id = 冲销id_In Where ID = 缴预交id_In And 结帐id Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '未找到对应的预交款记录！';
      Raise Err_Item;
    End If;
  End If;

  --需要检查退款合计是否相等
  Select Round(Sum(冲预交), 5), Round(Sum(结帐合计), 5)
  Into n_预交合计, n_结帐合计
  From (Select Sum(冲预交) As 冲预交, 0 As 结帐合计
         From 病人预交记录
         Where 结帐id = n_结帐id
         Union All
         Select 0 As 冲预交, Sum(结帐金额) As 结帐合计
         From 住院费用记录
         Where 结帐id = n_结帐id
         Union All
         Select 0 As 冲预交, Sum(结帐金额) As 结帐合计
         From 门诊费用记录
         Where 结帐id = n_结帐id);

  If Nvl(n_预交合计, 0) <> Nvl(n_结帐合计, 0) Then
    v_Err_Msg := '本次结算作废合计(' || Trim(To_Char(Nvl(n_预交合计, 0), '99999999999.99999')) || ')与本次结帐作废费用合计(' ||
                 Trim(To_Char(Nvl(n_结帐合计, 0), '99999999999.99999')) || ')不等，请检查当前结算与费用明细合计是否一致！';
    Raise Err_Item;
  End If;

  Close c_Pati;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结帐记录_Delete;
/

--73706:冉俊明,2018-08-01,体检费用不更新病人余额.费用余额
Create Or Replace Procedure Zl_结帐费用记录_Insert
(
  Id_In       住院费用记录.Id%Type,
  No_In       住院费用记录.No%Type,
  记录性质_In 住院费用记录.记录性质%Type,
  记录状态_In 住院费用记录.记录状态%Type,
  执行状态_In 住院费用记录.执行状态%Type,
  序号_In     住院费用记录.序号%Type,
  结帐金额_In 住院费用记录.结帐金额%Type,
  结帐id_In   住院费用记录.结帐id%Type
) As
  n_Next_Id    住院费用记录.Id%Type;
  n_病人id     住院费用记录.病人id%Type;
  n_主页id     住院费用记录.主页id%Type;
  n_病人病区id 住院费用记录.病人病区id%Type;
  n_病人科室id 住院费用记录.病人科室id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  n_执行部门id 住院费用记录.执行部门id%Type;
  n_收入项目id 住院费用记录.收入项目id%Type;
  n_门诊标志   住院费用记录.门诊标志%Type;
  n_记帐费用   住院费用记录.记帐费用%Type;
  v_操作员     住院费用记录.操作员姓名%Type;
  v_操作员姓名 住院费用记录.操作员姓名%Type;

  n_结帐金额 住院费用记录.结帐金额%Type;
  n_实收金额 住院费用记录.实收金额%Type;
  n_返回值   病人余额.预交余额%Type;
  n_类别     Number(18);
  v_Temp     Varchar2(500);
  Err_Custom Exception;
  Err_Special Exception;
  v_Error Varchar2(255);
  n_来源  Number;
Begin
  --人员id,人员编号,人员姓名
  v_Temp := Zl_Identity(1);
  If Not (Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_') Then
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_操作员姓名 := v_Temp;
  End If;

  If Id_In <> 0 Then
    Begin
      Select 2 Into n_来源 From 住院费用记录 Where ID = Id_In;
    Exception
      When Others Then
        n_来源 := 1;
    End;
  
    --第一次结帐但部分结
    If n_来源 = 1 Then
      Update 门诊费用记录 Set 结帐金额 = 结帐金额_In, 结帐id = 结帐id_In Where ID = Id_In And 结帐id Is Null;
    Else
      Update 住院费用记录 Set 结帐金额 = 结帐金额_In, 结帐id = 结帐id_In Where ID = Id_In And 结帐id Is Null;
    End If;
  
    If Sql%RowCount = 0 Then
      If n_来源 = 1 Then
        Select Max(b.操作员姓名)
        Into v_操作员
        From 门诊费用记录 A, 病人结帐记录 B
        Where a.Id = Id_In And b.Id = a.结帐id;
      Else
        Select Max(b.操作员姓名)
        Into v_操作员
        From 住院费用记录 A, 病人结帐记录 B
        Where a.Id = Id_In And b.Id = a.结帐id;
      End If;
      If v_操作员 Is Null Then
        v_Error := '未发现结帐的费用,当前结帐操作不能继续。';
        Raise Err_Custom;
      Else
        If v_操作员姓名 = v_操作员 Then
          v_Error := '发现已经被结帐的费用,当前结帐操作不能继续。';
          Raise Err_Special;
        Else
          v_Error := '发现已经被其他人结帐的费用,当前结帐操作不能继续。';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  
    n_Next_Id := Id_In;
  Else
    --结以前的余帐
    Select 病人费用记录_Id.Nextval Into n_Next_Id From Dual;
  
    If Mod(记录性质_In, 10) = 3 Or Mod(记录性质_In, 10) = 5 Then
      --自动记帐或就诊卡;肯定是住院
      n_来源 := 2;
    Else
      Begin
        Select 2
        Into n_来源
        From 住院费用记录
        Where NO = No_In And 序号 = 序号_In And 记录状态 = 记录状态_In And Nvl(执行状态, 0) = Nvl(执行状态_In, 0) And
              Substr(记录性质, Length(记录性质), 1) = 记录性质_In And Rownum = 1;
      Exception
        When Others Then
          n_来源 := 1;
      End;
    End If;
  
    If n_来源 = 1 Then
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别,
         收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
         执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 保险编码, 费用类型, 是否急诊, 摘要)
        Select n_Next_Id, NO, 实际票号, To_Number('1' || 记录性质_In), 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄,
               标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, Null,
               Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额_In, 结帐id_In, 保险项目否,
               保险大类id, 统筹金额, 保险编码, 费用类型, 是否急诊, 摘要
        From 门诊费用记录
        Where NO = No_In And 序号 = 序号_In And 记录状态 = 记录状态_In And Nvl(执行状态, 0) = Nvl(执行状态_In, 0) And
              Substr(记录性质, Length(记录性质), 1) = 记录性质_In And Rownum = 1;
    
      --检查多次结帐后结帐金额是否高于原金额
      Select Nvl(Sum(实收金额), 0), Nvl(Sum(结帐金额), 0)
      Into n_实收金额, n_结帐金额
      From 门诊费用记录
      Where NO = No_In And 序号 = 序号_In And 记录状态 = 记录状态_In And Substr(记录性质, Length(记录性质), 1) = 记录性质_In And
            Nvl(执行状态, 0) = 执行状态_In;
    
    Else
      Insert Into 住院费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id,
         病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人,
         开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 保险编码, 费用类型,
         是否急诊, 摘要, 医疗小组id)
        Select n_Next_Id, NO, 实际票号, To_Number('1' || 记录性质_In), 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志,
               姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id,
               收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额_In,
               结帐id_In, 保险项目否, 保险大类id, 统筹金额, 保险编码, 费用类型, 是否急诊, 摘要, 医疗小组id
        From 住院费用记录
        Where NO = No_In And 序号 = 序号_In And 记录状态 = 记录状态_In And Nvl(执行状态, 0) = Nvl(执行状态_In, 0) And
              Substr(记录性质, Length(记录性质), 1) = 记录性质_In And Rownum = 1;
      --检查多次结帐后结帐金额是否高于原金额
      Select Nvl(Sum(实收金额), 0), Nvl(Sum(结帐金额), 0)
      Into n_实收金额, n_结帐金额
      From 住院费用记录
      Where NO = No_In And 序号 = 序号_In And 记录状态 = 记录状态_In And Substr(记录性质, Length(记录性质), 1) = 记录性质_In And
            Nvl(执行状态, 0) = 执行状态_In;
    End If;
    If n_结帐金额 > n_实收金额 Then
      If n_来源 = 1 Then
        Select Max(b.操作员姓名)
        Into v_操作员
        From 门诊费用记录 A, 病人结帐记录 B
        Where a.Id = Id_In And b.Id = a.结帐id;
      Else
        Select Max(b.操作员姓名)
        Into v_操作员
        From 住院费用记录 A, 病人结帐记录 B
        Where a.Id = Id_In And b.Id = a.结帐id;
      End If;
      If v_操作员 Is Null Then
        v_Error := '未发现结帐的费用,当前结帐操作不能继续。';
        Raise Err_Custom;
      Else
        If v_操作员姓名 = v_操作员 Then
          v_Error := '发现已经被结帐的费用,当前结帐操作不能继续。';
          Raise Err_Special;
        Else
          v_Error := '发现已经被其他人结帐的费用,当前结帐操作不能继续。';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  End If;
  If n_来源 = 1 Then
    Select 病人id, Null, Null, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 门诊标志, 记帐费用
    Into n_病人id, n_主页id, n_病人病区id, n_病人科室id, n_开单部门id, n_执行部门id, n_收入项目id, n_门诊标志, n_记帐费用
    From 门诊费用记录
    Where ID = n_Next_Id;
    n_类别 := 1;
  Else
    Select 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 门诊标志, 记帐费用
    Into n_病人id, n_主页id, n_病人病区id, n_病人科室id, n_开单部门id, n_执行部门id, n_收入项目id, n_门诊标志, n_记帐费用
    From 住院费用记录
    Where ID = n_Next_Id;
  
    If Nvl(n_门诊标志, 0) = 1 Or Nvl(n_门诊标志, 0) = 2 Then
      n_类别 := n_门诊标志;
    Elsif Nvl(n_主页id, 0) = 0 Or Nvl(n_门诊标志, 0) = 4 Then
      n_类别 := 1;
    Else
      n_类别 := 2;
    End If;
  End If;

  --病人余额
  If Nvl(n_门诊标志, 0) <> 4 Then
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) - 结帐金额_In
    Where 病人id = n_病人id And 性质 = 1 And 类型 = n_类别
    Returning 费用余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, n_类别, 0, -1 * 结帐金额_In);
      n_返回值 := -1 * 结帐金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0 And 病人id = n_病人id;
    End If;
  End If;

  --病人未结费用
  Update 病人未结费用
  Set 金额 = Nvl(金额, 0) - 结帐金额_In
  Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(n_主页id, 0) And Nvl(病人病区id, 0) = Nvl(n_病人病区id, 0) And
        Nvl(病人科室id, 0) = Nvl(n_病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(n_开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(n_执行部门id, 0) And
        收入项目id + 0 = n_收入项目id And 来源途径 + 0 = n_门诊标志
  Returning 金额 Into n_返回值;
  If Sql%RowCount = 0 Then
    Insert Into 病人未结费用
      (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
    Values
      (n_病人id, Decode(n_主页id, 0, Null, n_主页id), Decode(n_病人病区id, 0, Null, n_病人病区id), n_病人科室id, n_开单部门id, n_执行部门id,
       n_收入项目id, n_门诊标志, -1 * 结帐金额_In);
    n_返回值 := -1 * 结帐金额_In;
  End If;
  If Nvl(n_返回值, 0) = 0 Then
    Delete From 病人未结费用 Where 病人id = n_病人id And Nvl(金额, 0) = 0;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_结帐费用记录_Insert;
/

--73706:冉俊明,2018-08-01,体检费用不更新病人余额.费用余额
Create Or Replace Procedure Zl_结帐费用记录_Batch
(
  Ids_In    Varchar2,
  病人id_In 住院费用记录.病人id%Type,
  结帐id_In 住院费用记录.结帐id%Type
  
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
  v_Temp    Varchar2(500);
  Err_Special Exception;
  v_操作员     住院费用记录.操作员姓名%Type;
  v_操作员姓名 住院费用记录.操作员姓名%Type;

  n_门诊结算合计 住院费用记录.结帐金额%Type;
  n_住院结算合计 住院费用记录.结帐金额%Type;
  n_类别         Number;

  n_来源   Number; --1门诊;2-住院;3-门诊和住院
  n_返回值 病人余额.费用余额%Type;
Begin
  --人员id,人员编号,人员姓名
  v_Temp := Zl_Identity(1);
  If Not (Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_') Then
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_操作员姓名 := v_Temp;
  End If;
  n_门诊结算合计 := 0;
  n_住院结算合计 := 0;
  Begin
    Select Case
             When Max(住院来源) = 1 And Max(门诊来源) = 1 Then
              3
             When Max(门诊来源) = 1 Then
              1
             Else
              2
           End
    Into n_来源
    From (Select /*+ Rule*/
            1 As 住院来源, 0 As 门诊来源
           From 住院费用记录 A, (Select Column_Value From Table(Cast(f_Num2list(Ids_In) As Zltools.t_Numlist))) J
           Where a.Id = j.Column_Value And 病人id + 0 = 病人id_In And 结帐id Is Null And Rownum = 1
           Union All
           Select /*+ Rule*/
            0 As 住院来源, 1 As 门诊来源
           From 门诊费用记录 A, (Select Column_Value From Table(Cast(f_Num2list(Ids_In) As Zltools.t_Numlist))) J
           Where a.Id = j.Column_Value And 病人id + 0 = 病人id_In And 结帐id Is Null And Rownum = 1);
  Exception
    When Others Then
      n_来源 := 2;
  End;

  --第一次结帐并且全结,Ids_In最大长度限制3998
  If n_来源 = 1 Then
    Update 门诊费用记录
    Set 结帐金额 = 实收金额, 结帐id = 结帐id_In
    Where 病人id = 病人id_In And Instr(',' || Ids_In || ',', ',' || ID || ',') > 0 And 结帐id Is Null;
  Elsif n_来源 = 2 Then
    Update 住院费用记录
    Set 结帐金额 = 实收金额, 结帐id = 结帐id_In
    Where 病人id = 病人id_In And Instr(',' || Ids_In || ',', ',' || ID || ',') > 0 And 结帐id Is Null;
  Else
    Update 门诊费用记录
    Set 结帐金额 = 实收金额, 结帐id = 结帐id_In
    Where 病人id = 病人id_In And Instr(',' || Ids_In || ',', ',' || ID || ',') > 0 And 结帐id Is Null;
  
    Update 住院费用记录
    Set 结帐金额 = 实收金额, 结帐id = 结帐id_In
    Where 病人id = 病人id_In And Instr(',' || Ids_In || ',', ',' || ID || ',') > 0 And 结帐id Is Null;
  End If;

  If Sql%RowCount = 0 Then
    If n_来源 = 1 Then
      Select Max(b.操作员姓名)
      Into v_操作员
      From 门诊费用记录 A, 病人结帐记录 B
      Where Instr(',' || Ids_In || ',', ',' || a.Id || ',') > 0 And b.Id = a.结帐id;
    Else
      Select Max(b.操作员姓名)
      Into v_操作员
      From 住院费用记录 A, 病人结帐记录 B
      Where Instr(',' || Ids_In || ',', ',' || a.Id || ',') > 0 And b.Id = a.结帐id;
    End If;
    If v_操作员 Is Null Then
      v_Err_Msg := '未发现结帐的费用,当前结帐操作不能继续。';
      Raise Err_Item;
    Else
      If v_操作员姓名 = v_操作员 Then
        v_Err_Msg := '发现已经被结帐的费用,当前结帐操作不能继续。';
        Raise Err_Special;
      Else
        v_Err_Msg := '发现已经被其他人结帐的费用,当前结帐操作不能继续。';
        Raise Err_Item;
      End If;
    End If;
  End If;
  If n_来源 = 2 Or n_来源 = 3 Then
    For r_f In (Select 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 门诊标志, 记帐费用, Sum(结帐金额) 结帐金额

                
                From 住院费用记录
                Where 病人id = 病人id_In And Instr(',' || Ids_In || ',', ',' || ID || ',') > 0
                Group By 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 门诊标志, 记帐费用) Loop
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - r_f.结帐金额
      Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(r_f.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_f.病人病区id, 0) And
            Nvl(病人科室id, 0) = Nvl(r_f.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_f.开单部门id, 0) And
            Nvl(执行部门id, 0) = Nvl(r_f.执行部门id, 0) And 收入项目id + 0 = r_f.收入项目id And 来源途径 + 0 = r_f.门诊标志;
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, r_f.主页id, r_f.病人病区id, r_f.病人科室id, r_f.开单部门id, r_f.执行部门id, r_f.收入项目id, r_f.门诊标志, -1 * r_f.结帐金额);
      End If;
    
      If Nvl(r_f.门诊标志, 0) = 1 Or Nvl(r_f.门诊标志, 0) = 2 Then
        n_类别 := r_f.门诊标志;
      Elsif Nvl(r_f.主页id, 0) = 0 Or Nvl(r_f.门诊标志, 0) = 4 Then
        n_类别 := 1;
      Else
        n_类别 := 2;
      End If;
    
      If Nvl(r_f.门诊标志, 0) <> 4 Then
        If n_类别 = 1 Then
          n_门诊结算合计 := Nvl(n_门诊结算合计, 0) + Nvl(r_f.结帐金额, 0);
        Else
          n_住院结算合计 := Nvl(n_住院结算合计, 0) + Nvl(r_f.结帐金额, 0);
        End If;
      End If;
    End Loop;
  End If;
  If n_来源 = 3 Or n_来源 = 1 Then
    For r_f In (Select Null As 主页id, Null As 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 门诊标志, 记帐费用, Sum(结帐金额) 结帐金额
                From 门诊费用记录
                Where 病人id = 病人id_In And Instr(',' || Ids_In || ',', ',' || ID || ',') > 0
                Group By 病人科室id, 开单部门id, 执行部门id, 收入项目id, 门诊标志, 记帐费用) Loop
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - r_f.结帐金额
      Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(r_f.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_f.病人病区id, 0) And
            Nvl(病人科室id, 0) = Nvl(r_f.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_f.开单部门id, 0) And
            Nvl(执行部门id, 0) = Nvl(r_f.执行部门id, 0) And 收入项目id + 0 = r_f.收入项目id And 来源途径 + 0 = r_f.门诊标志;
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, r_f.主页id, r_f.病人病区id, r_f.病人科室id, r_f.开单部门id, r_f.执行部门id, r_f.收入项目id, r_f.门诊标志, -1 * r_f.结帐金额);
      End If;
      If Nvl(r_f.门诊标志, 0) = 1 Or Nvl(r_f.门诊标志, 0) = 2 Then
        n_类别 := r_f.门诊标志;
      Elsif Nvl(r_f.主页id, 0) = 0 Or Nvl(r_f.门诊标志, 0) = 4 Then
        n_类别 := 1;
      Else
        n_类别 := 2;
      End If;
    
      If Nvl(r_f.门诊标志, 0) <> 4 Then
        If n_类别 = 1 Then
          n_门诊结算合计 := Nvl(n_门诊结算合计, 0) + Nvl(r_f.结帐金额, 0);
        Else
          n_住院结算合计 := Nvl(n_住院结算合计, 0) + Nvl(r_f.结帐金额, 0);
        End If;
      End If;
    End Loop;
  End If;

  Delete From 病人未结费用 Where 病人id = 病人id_In And Nvl(金额, 0) = 0;

  --病人余额
  If Nvl(n_门诊结算合计, 0) <> 0 Then
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) - n_门诊结算合计
    Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1
    Returning 费用余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (病人id_In, 1, 1, 0, -1 * n_门诊结算合计);
      n_返回值 := -1 * n_门诊结算合计;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0 And 病人id = 病人id_In And 类型 = 1;
    End If;
  End If;
  If Nvl(n_住院结算合计, 0) <> 0 Then
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) - n_住院结算合计
    Where 病人id = 病人id_In And 性质 = 1 And 类型 = 2
    Returning 费用余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (病人id_In, 1, 2, 0, -1 * n_住院结算合计);
      n_返回值 := -1 * n_住院结算合计;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0 And 病人id = 病人id_In And 类型 = 2;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_结帐费用记录_Batch;
/

--73706:冉俊明,2018-08-01,体检费用不更新病人余额.费用余额
Create Or Replace Procedure Zl_门诊记帐记录_Delete
(
  No_In         门诊费用记录.No%Type,
  序号_In       Varchar2,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type
) As
  --功能：冲销一张门诊记帐单据中指定序号行
  --序号：格式如"1,3,5,7,8",为空表示冲销所有可冲销行
  --该光标用于销帐指定费用行

  --该游标为要退费单据的所有原始记录
  Cursor c_Bill(n_标志 Number) Is
    Select a.Id, a.价格父号, a.序号, a.执行状态, a.收费类别, a.医嘱序号, a.病人id, a.收入项目id, a.开单部门id, a.执行部门id, a.病人科室id, a.实收金额,
           Decode(a.记录状态, 0, 1, 0) As 划价, j.诊疗类别, m.跟踪在用
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.收费细目id + 0 = m.材料id(+) And a.No = No_In And a.记录性质 = 2 And a.记录状态 In (0, 1, 3) And
          a.门诊标志 = n_标志
    Order By a.收费细目id, a.序号;

  --该游标用于处理药品库存可用数量
  --不要管费用的执行状态,因为先于此步处理
  Cursor c_Stock(n_标志 Number) Is
    Select ID, 库房id, 药品id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
    From 药品收发记录
    Where NO = No_In And 单据 In (9, 25) And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 门诊费用记录
                   Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = n_标志 And
                         (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
    Order By 药品id;

  --该游标用于处理费用记录序号
  Cursor c_Serial Is
    Select 序号, 价格父号 From 门诊费用记录 Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) Order By 序号;
  l_药品收发 t_Numlist := t_Numlist();
  l_划价     t_Numlist := t_Numlist();
  l_费用id   t_Numlist := t_Numlist();

  v_医嘱ids Varchar2(4000);

  n_医嘱id   病人医嘱记录.Id%Type;
  n_父号     门诊费用记录.价格父号%Type;
  n_门诊标志 门诊费用记录.门诊标志%Type;

  --部分退费计算变量
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;

  n_准退数量 Number;
  n_退费次数 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;

  n_Dec Number;

  n_Count   Number;
  d_Curdate Date;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --是否已经全部完全执行(只是整张单据的检查)
  Select Nvl(Count(*), 0), Max(Nvl(门诊标志, 1))
  Into n_Count, n_门诊标志
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  If Nvl(n_门诊标志, 0) = 0 Then
    n_门诊标志 := 1;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 门诊费用记录
                Where NO = No_In And 记录性质 = 2 And 门诊标志 = n_门诊标志 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 2 And 门诊标志 = n_门诊标志 And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！';
    Raise Err_Item;
  End If;

  ---------------------------------------------------------------------------------
  --公用变量
  Select Sysdate Into d_Curdate From Dual;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --循环处理每行费用(收入项目行)
  For r_Bill In c_Bill(n_门诊标志) Loop
    If Instr(',' || 序号_In || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or 序号_In Is Null Then
    
      If r_Bill.划价 = 0 Then
        If Nvl(r_Bill.执行状态, 0) <> 1 Then
          --求剩余数量,剩余应收,剩余实收
          Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
          Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
          From 门诊费用记录
          Where NO = No_In And 记录性质 = 2 And 序号 = r_Bill.序号;
        
          If n_剩余数量 = 0 Then
            If 序号_In Is Not Null Then
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部销帐！';
              Raise Err_Item;
            End If;
            --情况：未限定行号,原始单据中的该笔已经全部销帐(执行状态=0的一种可能)
          Else
            --准销数量(非药品项目为剩余数量,原始数量)
            If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
            
              --@@@
              --非药品部分(以具体医嘱执行为准进行检查)
              --: 1.存在医嘱发送的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血)
              --: 2.对于病人医吃计价中的收费方式为:0-正常收取 的,才支持部分退;如果是其他的,则只能全退
              --: 3.不存在医嘱的,则以剩余数量为准
              n_Count := 0;
              If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null Then
              
                Select Nvl(Sum(数量), 0), Count(*)
                Into n_准退数量, n_Count
                From (Select j.医嘱序号 As 医嘱id, j.收费细目id, Nvl(j.付数, 1) * Nvl(j.数次, 1) As 数量
                       From 门诊费用记录 J, 病人医嘱记录 M
                       Where j.医嘱序号 = m.Id And j.No = No_In And j.记录性质 = 2 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And
                             Exists
                        (Select 1
                              From 病人医嘱发送 A
                              Where a.医嘱id = j.医嘱序号 And Nvl(a.执行状态, 0) <> 1 And a.No || '' = No_In) And Exists
                        (Select 1
                              From 病人医嘱计价 A
                              Where a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) = 0) And j.价格父号 Is Null And
                             Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             (j.记录状态 In (1, 3) And Not Exists
                              (Select 1
                               From 药品收发记录
                               Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) Or
                              j.记录状态 = 2 And Not Exists
                              (Select 1 From 药品收发记录 Where NO = No_In And 单据 In (8, 24) And 药品id = j.收费细目id))
                       Union All
                       Select a.医嘱id, a.收费细目id, -1 * Nvl(a.数量, 1) * Nvl(c.本次数次, 1) As 数量
                       From 病人医嘱计价 A, 病人医嘱发送 B, 病人医嘱执行 C, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = b.医嘱id And b.医嘱id = c.医嘱id And Nvl(a.收费方式, 0) = 0 And b.发送号 = c.发送号 And
                             a.医嘱id = m.Id And Nvl(c.执行结果, 1) = 1 And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And
                             a.收费细目id = j.收费细目id And j.No = No_In And j.记录性质 = 2 And j.序号 = r_Bill.序号 And
                             j.记录状态 In (1, 3) And j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             Not Exists
                        (Select 1
                              From 药品收发记录
                              Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) And Not Exists
                        (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)
                       Union All
                       Select a.医嘱id, a.收费细目id, 0 As 数量
                       From 病人医嘱计价 A, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = m.Id And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) <> 0 And
                             j.No = No_In And j.记录性质 = 2 And Nvl(j.执行状态, 0) = 2 And Not Exists
                        (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1) And
                             Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0);
              
              End If;
            
              If Nvl(n_Count, 0) = 0 Then
                n_准退数量 := n_剩余数量;
              End If;
            
            Else
              Select Sum(Nvl(付数, 1) * 实际数量)
              Into n_准退数量
              From 药品收发记录
              Where NO = No_In And 单据 In (9, 25) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = r_Bill.Id;
            
              --不跟踪在用的卫生材料
              If r_Bill.收费类别 = '4' And Nvl(n_准退数量, 0) = 0 Then
                n_准退数量 := n_剩余数量;
              End If;
            End If;
          
            --处理门诊费用记录
          
            --该笔项目第几次销帐
            Select Nvl(Max(Abs(执行状态)), 0) + 1
            Into n_退费次数
            From 门诊费用记录
            Where NO = No_In And 记录性质 = 2 And 记录状态 = 2 And 序号 = r_Bill.序号;
          
            --金额=剩余金额*(准退数/剩余数)
            n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), n_Dec);
            n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), n_Dec);
            n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), n_Dec);
          
            --插入退费记录
            Insert Into 门诊费用记录
              (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别,
               收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人,
               执行人, 执行状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论)
              Select 病人费用记录_Id.Nextval, NO, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                     病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                     Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用,
                     标准单价, -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, 执行时间, 操作员编号_In,
                     操作员姓名_In, 发生时间, d_Curdate, 保险项目否, 保险大类id, -1 * n_统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论
              From 门诊费用记录
              Where ID = r_Bill.Id;
          
            --记录病人医嘱附费对应的医嘱ID(不是主费用)
            If n_医嘱id Is Null And r_Bill.医嘱序号 Is Not Null Then
              n_医嘱id := r_Bill.医嘱序号;
            End If;
          
            --病人余额
            If n_门诊标志 <> 4 Then
              Update 病人余额
              Set 费用余额 = Nvl(费用余额, 0) - n_实收金额
              Where 病人id = r_Bill.病人id And 性质 = 1 And 类型 = 1;
              If Sql%RowCount = 0 Then
                Insert Into 病人余额
                  (病人id, 性质, 类型, 费用余额, 预交余额)
                Values
                  (r_Bill.病人id, 1, 1, -1 * n_实收金额, 0);
              End If;
            End If;
          
            --病人未结费用
            Update 病人未结费用
            Set 金额 = Nvl(金额, 0) - n_实收金额
            Where 病人id = r_Bill.病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And
                  Nvl(病人科室id, 0) = Nvl(r_Bill.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Bill.开单部门id, 0) And
                  Nvl(执行部门id, 0) = Nvl(r_Bill.执行部门id, 0) And 收入项目id + 0 = r_Bill.收入项目id And 来源途径 + 0 = n_门诊标志;
            If Sql%RowCount = 0 Then
              Insert Into 病人未结费用
                (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
              Values
                (r_Bill.病人id, Null, Null, r_Bill.病人科室id, r_Bill.开单部门id, r_Bill.执行部门id, r_Bill.收入项目id, n_门诊标志,
                 -1 * n_实收金额);
            End If;
          
            --标记原费用记录
            --执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1
            Update 门诊费用记录
            Set 记录状态 = 3, 执行状态 = Decode(Sign(n_准退数量 - n_剩余数量), 0, 0, 1)
            Where ID = r_Bill.Id;
          End If;
        Else
          If 序号_In Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能销帐！';
            Raise Err_Item;
          End If;
          --情况:没限定行号,原始单据中包括已经完全执行的
        End If;
      End If;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --药品相关内容
  ------------------------------------------------------------------------------------------------------------------------
  --先处理备货材料
  For v_出库 In (Select ID, 库房id, 药品id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
               From 药品收发记录
               Where 单据 = 21 And Mod(记录状态, 3) = 1 And 审核人 Is Null And
                     费用id In (Select ID
                              From 门诊费用记录
                              Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And 收费类别 = '4' And 门诊标志 = n_门诊标志 And
                                    (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
               Order By 药品id) Loop
    --处理药品库存
    If v_出库.库房id Is Not Null Then
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0)
      Where 库房id = v_出库.库房id And 药品id = v_出库.药品id And Nvl(批次, 0) = Nvl(v_出库.批次, 0) And 性质 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 药品库存
          (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
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

  For r_Stock In c_Stock(n_门诊标志) Loop
  
    --处理药品库存
    If r_Stock.库房id Is Not Null Then
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0)
      Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 药品库存
          (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
        Values
          (r_Stock.库房id, r_Stock.药品id, 1, r_Stock.批次, r_Stock.效期,
           Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0), r_Stock.批号, r_Stock.产地,
           r_Stock.灭菌效期, r_Stock.商品条码, r_Stock.内部条码);
      End If;
    
      Zl_药品库存_可用数量异常处理(r_Stock.库房id, r_Stock.药品id, Nvl(r_Stock.批次, 0));
    End If;
  
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := r_Stock.Id;
  End Loop;

  --删除药品收发记录
  Forall I In 1 .. l_药品收发.Count
    Delete From 药品收发记录 Where ID = l_药品收发(I);

  ------------------------------------------------------------------------------------------------------------------------
  --批量删未发药品记录

  Delete From 未发药品记录 A
  Where NO = No_In And 单据 In (9, 25) And Not Exists
   (Select 1
         From 药品收发记录
         Where 单据 = a.单据 And Nvl(库房id, 0) = Nvl(a.库房id, 0) And NO = No_In And Mod(记录状态, 3) = 1 And 审核人 Is Null);
  ------------------------------------------------------------------------------------------------------------------------

  ---------------------------------------------------------------------------------
  --如果是划价,直接删除费用记录(药品处理后)
  n_Count   := 0;
  v_医嘱ids := Null;
  For r_Bill In c_Bill(n_门诊标志) Loop
    If Instr(',' || 序号_In || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or 序号_In Is Null Then
      If r_Bill.划价 = 1 Then
        If Nvl(r_Bill.执行状态, 0) <> 1 Then
          l_划价.Extend;
          l_划价(l_划价.Count) := r_Bill.Id;
        
          --Delete From 门诊费用记录 Where ID = r_Bill.ID;
          n_Count := n_Count + 1; --记录是否有删除行
        
          If r_Bill.医嘱序号 Is Not Null Then
            If Instr(',' || Nvl(v_医嘱ids, '') || ',', ',' || r_Bill.医嘱序号 || ',') = 0 Then
              v_医嘱ids := Nvl(v_医嘱ids, '') || ',' || r_Bill.医嘱序号;
            End If;
          
            --记录病人医嘱附费对应的医嘱ID(不是主费用)
            If n_医嘱id Is Null Then
              n_医嘱id := r_Bill.医嘱序号;
            End If;
          End If;
        Else
          If 序号_In Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能销帐！';
            Raise Err_Item;
          End If;
          --情况:没限定行号,原始单据中包括已经完全执行的
        End If;
      End If;
    End If;
  End Loop;

  --删除划价记录
  Forall I In 1 .. l_划价.Count
    Delete From 门诊费用记录 Where ID = l_划价(I);

  --删除之后再统一调整序号
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.价格父号 Is Null Then
        n_父号 := n_Count;
      End If;
    
      Update 门诊费用记录
      Set 序号 = n_Count, 价格父号 = Decode(价格父号, Null, Null, n_父号)
      Where NO = No_In And 记录性质 = 2 And 序号 = r_Serial.序号;
    
      Update 门诊费用记录 Set 从属父号 = n_Count Where NO = No_In And 记录性质 = 2 And 从属父号 = r_Serial.序号;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;

  --整张单据全部冲完时，删除病人医嘱附费
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where 记录性质 = 2 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 2 And NO = No_In;
    End If;
  End Loop;

  If v_医嘱ids Is Not Null Then
    --医嘱处理
    --场合_In    Integer:=0, --0:门诊;1-住院
    --性质_In    Integer:=1, --1-收费单;2-记帐单
    --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2 := Null
    v_医嘱ids := Substr(v_医嘱ids, 2);
    Zl_医嘱发送_计费状态_Update(0, 2, 0, No_In, v_医嘱ids);
  Else
    Zl_医嘱发送_计费状态_Update(0, 2, 2, No_In, v_医嘱ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊记帐记录_Delete;
/

--73706:冉俊明,2018-08-01,体检费用不更新病人余额.费用余额
CREATE OR REPLACE Procedure Zl_门诊记帐记录_Insert
(
  No_In         门诊费用记录.No%Type,
  序号_In       门诊费用记录.序号%Type,
  病人id_In     门诊费用记录.病人id%Type,
  标识号_In     门诊费用记录.标识号%Type,
  姓名_In       门诊费用记录.姓名%Type,
  性别_In       门诊费用记录.性别%Type,
  年龄_In       门诊费用记录.年龄%Type,
  费别_In       门诊费用记录.费别%Type,
  加班标志_In   门诊费用记录.加班标志%Type,
  婴儿费_In     门诊费用记录.婴儿费%Type,
  病人科室id_In 门诊费用记录.病人科室id%Type,
  开单部门id_In 门诊费用记录.开单部门id%Type,
  开单人_In     门诊费用记录.开单人%Type,
  从属父号_In   门诊费用记录.从属父号%Type,
  收费细目id_In 门诊费用记录.收费细目id%Type,
  收费类别_In   门诊费用记录.收费类别%Type,
  计算单位_In   门诊费用记录.计算单位%Type,
  付数_In       门诊费用记录.付数%Type,
  数次_In       门诊费用记录.数次%Type,
  附加标志_In   门诊费用记录.附加标志%Type,
  执行部门id_In 门诊费用记录.执行部门id%Type,
  价格父号_In   门诊费用记录.价格父号%Type,
  收入项目id_In 门诊费用记录.收入项目id%Type,
  收据费目_In   门诊费用记录.收据费目%Type,
  标准单价_In   门诊费用记录.标准单价%Type,
  应收金额_In   门诊费用记录.应收金额%Type,
  实收金额_In   门诊费用记录.实收金额%Type,
  发生时间_In   门诊费用记录.发生时间%Type,
  登记时间_In   门诊费用记录.登记时间%Type,
  药品摘要_In   药品收发记录.摘要%Type,
  划价_In       Number,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  类别id_In     药品单据性质.类别id%Type := Null,
  记帐单id_In   门诊费用记录.记帐单id%Type := Null,
  费用摘要_In   门诊费用记录.摘要%Type := Null,
  医嘱序号_In   门诊费用记录.医嘱序号%Type := Null,
  频次_In       药品收发记录.频次%Type := Null,
  单量_In       药品收发记录.单量%Type := Null,
  用法_In       药品收发记录.用法%Type := Null, --用法[|煎法]
  期效_In       药品收发记录.扣率%Type := Null,
  计价特性_In   药品收发记录.扣率%Type := Null,
  门诊标志_In   门诊费用记录.门诊标志%Type := 1,
  中药形态_In   门诊费用记录.结论%Type := Null,
  备货材料_In   Number := 0,
  批次_In       药品收发记录.批次%Type := Null
) As
  --功能：新收一张门诊记帐单据
  --参数：
  --   药品摘要_IN:修改保存新单据时用。目前仅用于存放于药品收发记录的摘要中。
  --         原单据(记录状态=2)记录修改产生的新单据号。
  --         新单据(记录状态=1)记录所修改的原单据号。
  v_费用id 门诊费用记录.Id%Type;
  v_优先级 未发药品记录.优先级%Type;
  n_急诊   病人挂号记录.急诊%Type;

  --药房分批、时价药品--
  ------------------------------------------------------------
  --该游标用于分批药品数量分解
  Cursor c_Stock
  (
    n_Outmode Number,
    n_库房id  药品收发记录.库房id%Type
  ) Is
    Select 库房id, 药品id, 批次, 效期, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 灭菌效期, 批准文号, 零售价, 上次扣率,
           商品条码, 内部条码
    From 药品库存
    Where 药品id = 收费细目id_In And 库房id = n_库房id And 性质 = 1 And (Nvl(批次, 0) = Nvl(批次_In, 0) Or Nvl(批次_In, 0) = 0) And
          (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And Nvl(可用数量, 0) > 0
    Order By Decode(n_Outmode, 1, 效期, Null), Decode(n_Outmode, 2, 上次批号, Null), Nvl(批次, 0);
  r_Stock      c_Stock%RowType;
  v_商品条码   药品收发记录.商品条码%Type;
  v_内部条码   药品收发记录.内部条码%Type;
  n_虚拟库房id 药品收发记录.库房id%Type;
  v_其他出库no 药品收发记录.No%Type;
  n_库房id     药品库存.库房id%Type;
  n_出库序号   药品收发记录.序号%Type;
  v_部门名称   部门表.名称%Type;
  --属性
  n_分批 药品规格.药房分批%Type;
  n_时价 收费项目目录.是否变价%Type;
  v_名称 收费项目目录.名称%Type;
  --临时变量
  n_总数量   Number;
  n_当前数量 Number;
  n_总金额   Number;
  n_当前单价 Number;
  --药品收发记录
  n_批次       药品收发记录.批次%Type;
  v_产地       药品收发记录.产地%Type;
  v_批号       药品收发记录.批号%Type;
  d_效期       药品收发记录.效期%Type;
  n_序号       药品收发记录.序号%Type;
  n_扣率       药品收发记录.扣率%Type;
  d_灭菌效期   药品收发记录.灭菌效期%Type;
  d_灭菌日期   药品收发记录.灭菌日期%Type;
  n_供药单位id 药品收发记录.供药单位id%Type;
  d_生产日期   药品收发记录.生产日期%Type;
  v_批准文号   药品收发记录.批准文号%Type;
  ------------------------------------------------------------
  v_用法       药品收发记录.用法%Type;
  v_煎法       药品收发记录.外观%Type;
  n_Aval       药品库存.可用数量%Type;
  n_修正库房id 药品库存.库房id%Type;
  n_单价小数   Number;

  n_Outmode Number(1);
  n_Dec     Number;
  n_Count   Number;
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_发药窗口 药品收发记录.发药窗口%Type;
  n_跟踪在用 材料特性.跟踪在用%Type;
  n_出库检查 Number(1);
Begin
  n_跟踪在用 := 0;
  If 收费类别_In = '4' Then
    --跟踪在用的卫材才处理
    Select Nvl(跟踪在用, 0) Into n_跟踪在用 From 材料特性 Where 材料id = 收费细目id_In;
  End If;

  --根据执行库房取虚拟库房ID
  Begin
    Select 虚拟库房id Into n_虚拟库房id From 虚拟库房对照 Where 科室id = 执行部门id_In And Rownum <= 1;
  Exception
    When Others Then
      n_虚拟库房id := 0;
  End;
  If Nvl(批次_In, 0) <> 0 Then
    Select Nvl(Sum(可用数量), 0)
    Into n_Aval
    From 药品库存
    Where 药品id = 收费细目id_In And 批次 = 批次_In And 库房id = 执行部门id_In;
    If n_Aval <= 0 Then
      n_修正库房id := n_虚拟库房id;
    End If;
  End If;

  If Nvl(备货材料_In, 0) = 1 And 收费类别_In = '4' Then
    If Nvl(n_虚拟库房id, 0) = 0 Then
      Begin
        Select 名称 Into v_Err_Msg From 部门表 Where ID = 执行部门id_In;
      Exception
        When Others Then
          v_Err_Msg := '';
      End;
      v_Err_Msg := '执行部门"' || Nvl(v_Err_Msg, '') || '"未设置虚拟部门,请在卫材参数设置中设置.';
      Raise Err_Item;
    End If;
  End If;

  --药品用法煎法分解
  If 用法_In Is Not Null Then
    If Instr(用法_In, '|') > 0 Then
      v_用法 := Substr(用法_In, 1, Instr(用法_In, '|') - 1);
      v_煎法 := Substr(用法_In, Instr(用法_In, '|') + 1);
    Else
      v_用法 := 用法_In;
    End If;
  End If;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into n_Dec, n_单价小数
  From Dual;

  If (收费类别_In In ('5', '6', '7') Or 收费类别_In = '4' And n_跟踪在用 = 1) And Nvl(划价_In, 0) = 0 Then
    --同一张单据,满足同一药房同一窗口
    Begin
      Select 发药窗口
      Into v_发药窗口
      From 门诊费用记录
      Where 收费类别 In ('5', '6', '7', '4') And NO = No_In And 记录性质 = 2 And 执行部门id = 执行部门id_In And 发药窗口 Is Not Null And
            Rownum <= 1;
    Exception
      When Others Then
        v_发药窗口 := Null;
    End;
    If v_发药窗口 Is Null Then
      --同一病人在普通号挂号有效挂号天数内且未发药的且上班的,以最近一次记账窗口为准
      n_Count := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
      If n_Count = 0 Then
        n_Count := 1;
      End If;
    
      Begin
        Select 发药窗口
        Into v_发药窗口
        From (Select 登记时间, 发药窗口
               From 门诊费用记录 A
               Where 收费类别 In ('5', '6', '7', '4') And 病人id = 病人id_In And 登记时间 Between Sysdate - n_Count And Sysdate And
                     记录性质 = 2 And 执行部门id = 执行部门id_In And 发药窗口 Is Not Null And Exists
                (Select 1
                      From 未发药品记录
                      Where a.No = NO And 单据 In (9, 26) And 库房id + 0 = 执行部门id_In And 病人id + 0 = 病人id_In) And Exists
                (Select 1
                      From 发药窗口
                      Where Nvl(上班否, 0) = 1 And 名称 = a.发药窗口 And Nvl(专家, 0) = 0 And 药房id = 执行部门id_In)
               Order By 登记时间 Desc)
        Where Rownum <= 1;
      
      Exception
        When Others Then
          v_发药窗口 := Null;
      End;
      If v_发药窗口 Is Null Then
        v_发药窗口 := Zl_Get发药窗口(执行部门id_In);
      End If;
    End If;
  End If;
  --门诊费用记录
  Select 病人费用记录_Id.Nextval Into v_费用id From Dual;

  --是否是急诊挂号单
  If Nvl(医嘱序号_In, 0) <> 0 Then
    Select Nvl(Max(急诊), 0)
    Into n_急诊
    From 病人挂号记录
    Where NO In (Select 挂号单 From 病人医嘱记录 Where ID = Nvl(医嘱序号_In, 0)) And 病人id = 病人id_In;
  End If;

  Insert Into 门诊费用记录
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次, 加班标志,
     附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 操作员编号, 操作员姓名, 婴儿费, 记帐单id,
     摘要, 医嘱序号, 结论, 发药窗口, 是否急诊)
  Values
    (v_费用id, 2, No_In, Decode(划价_In, 1, 0, 1), 序号_In, Decode(从属父号_In, 0, Null, 从属父号_In),
     Decode(价格父号_In, 0, Null, 价格父号_In), 门诊标志_In, 病人id_In, Decode(标识号_In, 0, Null, 标识号_In), 姓名_In, 性别_In, 年龄_In,
     病人科室id_In, 费别_In, 收费类别_In, 收费细目id_In, 计算单位_In, 付数_In, 数次_In, 加班标志_In, 附加标志_In, 收入项目id_In, 收据费目_In, 标准单价_In, 应收金额_In,
     实收金额_In, 1, 操作员姓名_In, 开单部门id_In, 开单人_In, 发生时间_In, 登记时间_In, 执行部门id_In, 0, Decode(划价_In, 1, Null, 操作员编号_In),
     Decode(划价_In, 1, Null, 操作员姓名_In), 婴儿费_In, 记帐单id_In, 费用摘要_In, 医嘱序号_In, 中药形态_In, v_发药窗口, Nvl(n_急诊, 0));

  --相关汇总表的处理
  If Nvl(划价_In, 0) = 0 Then
    If Nvl(门诊标志_In, 0) <> 4 Then
      --病人余额
      Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) + 实收金额_In Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, 实收金额_In, 0);
      End If;
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + 实收金额_In
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 门诊标志_In;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 门诊标志_In, 实收金额_In);
    End If;
  
  End If;

  --药品和卫生材料部分

  If 收费类别_In In ('5', '6', '7') Or (收费类别_In = '4' And Nvl(n_跟踪在用, 0) = 1) Then
    If 收费类别_In = '4' Then
      Select Nvl(a.在用分批, 0), Nvl(b.是否变价, 0), b.名称
      Into n_分批, n_时价, v_名称
      From 材料特性 A, 收费项目目录 B
      Where a.材料id = b.Id And b.Id = 收费细目id_In;
    
      --卫材分批出库方式
      Select Zl_To_Number(Nvl(zl_GetSysParameter(156), 0)) Into n_Outmode From Dual;
    Else
      Select Nvl(a.药房分批, 0), Nvl(b.是否变价, 0), b.名称
      Into n_分批, n_时价, v_名称
      From 药品规格 A, 收费项目目录 B
      Where a.药品id = b.Id And b.Id = 收费细目id_In;
    
      --药品分批出库方式
      Select Zl_To_Number(Nvl(zl_GetSysParameter(150), 0)) Into n_Outmode From Dual;
    End If;
  
    n_总数量 := 付数_In * 数次_In;
    n_总金额 := 0;
    If Nvl(备货材料_In, 0) = 1 And 收费类别_In = '4' Then
      n_库房id := n_虚拟库房id;
    Else
      If Nvl(n_修正库房id, 0) <> 0 Then
        n_库房id := n_修正库房id;
      Else
        n_库房id := 执行部门id_In;
      End If;
    End If;
    Open c_Stock(n_Outmode, n_库房id);
  
    Begin
      If 收费类别_In In ('5', '6', '7') Then
        Select 检查方式 Into n_出库检查 From 药品出库检查 Where 库房id = n_库房id;
      Else
        Select 检查方式 Into n_出库检查 From 材料出库检查 Where 库房id = n_库房id;
      End If;
    Exception
      When Others Then
        n_出库检查 := 0;
    End;
  
    While n_总数量 <> 0 Loop
      Fetch c_Stock
        Into r_Stock;
      If c_Stock%NotFound Then
        --第一次就没有库存,分批或时价都不允许。
        --分批药品数量分解不完,也就是库存不足。
        If n_分批 = 1 Or n_时价 = 1 Then
          Close c_Stock;
          If 医嘱序号_In Is Null Then
            If 收费类别_In = '4' Then
              If Nvl(备货材料_In, 0) = 1 And Not (n_分批 = 1 Or n_时价 = 1) Then
                v_Err_Msg := '第 ' || 序号_In || ' 行的卫生材料"' || v_名称 || '"没有足够的材料库存,不能进行备货记帐！';
              Else
                v_Err_Msg := '第 ' || 序号_In || ' 行的分批或时价卫生材料"' || v_名称 || '"没有足够的材料库存' || Case
                               When Nvl(备货材料_In, 0) = 0 Then
                                '！'
                               Else
                                ',不能进行备货记帐！'
                             End;
              End If;
            Else
              v_Err_Msg := '第 ' || 序号_In || ' 行的分批或时价药品"' || v_名称 || '"没有足够的库存！';
            End If;
          Else
            If 收费类别_In = '4' Then
              If Nvl(备货材料_In, 0) = 1 And Not (n_分批 = 1 Or n_时价 = 1) Then
                v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现卫生材料"' || v_名称 || '"没有足够的材料库存,不能进行备货记帐！';
              Else
                v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现分批或时价卫生材料"' || v_名称 || '"没有足够的材料库存' || Case
                               When Nvl(备货材料_In, 0) = 0 Then
                                '！'
                               Else
                                ',不能进行备货记帐！'
                             End;
              End If;
            Else
              v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现分批或时价药品"' || v_名称 || '"没有足够的库存！';
            End If;
          End If;
          Raise Err_Item;
        End If;
      Elsif (n_分批 = 1 And Nvl(r_Stock.批次, 0) = 0) Or (n_分批 = 0 And Nvl(r_Stock.批次, 0) <> 0) Then
        Close c_Stock;
        If 医嘱序号_In Is Null Then
          If 收费类别_In = '4' Then
            v_Err_Msg := '第 ' || 序号_In || ' 行卫生材料"' || v_名称 || '"的分批属性与库存记录不相符,请检查材料数据的正确性！';
          Else
            v_Err_Msg := '第 ' || 序号_In || ' 行药品"' || v_名称 || '"的分批属性与库存记录不相符,请检查药品数据的正确性！';
          End If;
        Else
          If 收费类别_In = '4' Then
            v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现卫生材料"' || v_名称 || '"的分批属性与库存记录不相符,请检查材料数据的正确性！';
          Else
            v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现药品"' || v_名称 || '"的分批属性与库存记录不相符,请检查药品数据的正确性！';
          End If;
        End If;
        Raise Err_Item;
      End If;
    
      If c_Stock%Found Then
        If Nvl(r_Stock.实际数量, 0) = 0 And (n_总数量 > 0 Or n_时价 = 1) And n_出库检查 = 2 Then
          --实际数量为零时，如果严格控制库存，不允许出库
          --实际数量不为零，金额为零，可能是正常的零价格管理。
          --负数的情况相当于入库,这种情况应是允许的；但时价需要计算价格，必须要有实际数量。
          Close c_Stock;
          If 医嘱序号_In Is Null Then
            If 收费类别_In = '4' Then
              v_Err_Msg := '第 ' || 序号_In || ' 行的卫生材料"' || v_名称 || '"当前无库存实际数量，可能存在尚未退料的记录，当前不能出库。';
            Else
              v_Err_Msg := '第 ' || 序号_In || ' 行药品"' || v_名称 || '"当前无库存实际数量，可能存在尚未退药的记录，当前不能出库。';
            End If;
          Else
            If 收费类别_In = '4' Then
              v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现卫生材料"' || v_名称 || '"当前无库存实际数量，可能存在尚未退料的记录，当前不能出库。';
            Else
              v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现药品"' || v_名称 || '"当前无库存实际数量，可能存在尚未退药的记录，当前不能出库。';
            End If;
          End If;
          Raise Err_Item;
        End If;
      End If;
    
      --确定本次分解数量
      If n_分批 = 1 Or n_时价 = 1 Then
        --对于不分批的时价只可能分解一次,分解不完上面判断了.它分解是为了计算单价.
        --每次分解取小者,库存不够分解不完在上面判断.
        If n_总数量 <= Nvl(r_Stock.可用数量, 0) Then
          n_当前数量 := n_总数量;
        Else
          n_当前数量 := Nvl(r_Stock.可用数量, 0);
        End If;
        If n_时价 = 1 Then
          n_当前单价 := Round(Nvl(r_Stock.零售价, Nvl(r_Stock.实际金额 / r_Stock.实际数量, 0)), n_单价小数);
        Elsif n_分批 = 1 Then
          n_当前单价 := 标准单价_In;
        End If;
      Else
        --普通药品
        --不管够不够,程序中已根据参数判断
        If Nvl(备货材料_In, 0) = 1 And 收费类别_In = '4' Then
          If n_总数量 > Nvl(r_Stock.可用数量, 0) Then
            --不分批, 但又是备货卫材方式出库的,则需要检查当前库存是否充足.
            v_Err_Msg := '第 ' || 序号_In || ' 行的卫生材料"' || v_名称 || '"没有足够的材料库存,不能进行备货记帐！';
            Raise Err_Item;
          End If;
        End If;
        n_当前数量 := n_总数量;
        n_当前单价 := 标准单价_In;
      End If;
    
      --药品库存(普通情况可能没有记录)
      If c_Stock%Found Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - n_当前数量
        Where 库房id = n_库房id And 药品id = 收费细目id_In And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
      
        Zl_药品库存_可用数量异常处理(n_库房id, 收费细目id_In, Nvl(r_Stock.批次, 0));
      Elsif 执行部门id_In Is Not Null Then
        --只有不分批非时价药品可能库存不足出库
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - n_当前数量
        Where 库房id = n_库房id And 药品id = 收费细目id_In And Nvl(批次, 0) = 0 And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 可用数量, 商品条码, 内部条码)
          Values
            (执行部门id_In, 收费细目id_In, 1, -1 * n_当前数量, r_Stock.商品条码, r_Stock.内部条码);
        End If;
      
        Zl_药品库存_可用数量异常处理(n_库房id, 收费细目id_In, 0);
      End If;
    
      --高值卫材模式减少发料部门可用数量
      If Nvl(备货材料_In, 0) = 1 And 收费类别_In = '4' Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - n_当前数量
        Where 库房id = 执行部门id_In And 药品id = 收费细目id_In And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
      
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 商品条码, 内部条码, 效期, 上次批号, 上次产地, 上次供应商id, 上次生产日期, 批准文号)
          Values
            (执行部门id_In, 收费细目id_In, Nvl(r_Stock.批次, 0), 1, -1 * n_当前数量, r_Stock.商品条码, r_Stock.内部条码, r_Stock.效期,
             r_Stock.上次批号, r_Stock.上次产地, r_Stock.上次供应商id, r_Stock.上次生产日期, r_Stock.批准文号);
        End If;
      End If;
    
      --药品收发记录
      n_批次       := Null;
      v_批号       := Null;
      d_效期       := Null;
      v_产地       := Null;
      d_灭菌效期   := Null;
      d_灭菌日期   := Null;
      n_供药单位id := Null;
      d_生产日期   := Null;
      v_批准文号   := Null;
    
      If c_Stock%Found Then
        n_批次       := r_Stock.批次;
        v_批号       := r_Stock.上次批号;
        d_效期       := r_Stock.效期;
        v_产地       := r_Stock.上次产地;
        n_供药单位id := r_Stock.上次供应商id;
        d_生产日期   := r_Stock.上次生产日期;
        v_批准文号   := r_Stock.批准文号;
        v_商品条码   := r_Stock.商品条码;
        v_内部条码   := r_Stock.内部条码;
      
        --卫材灭菌效期:一次性材料且有效期
        If 收费类别_In = '4' Then
          n_Count := 0;
          Begin
            Select 灭菌效期 Into n_Count From 材料特性 Where Nvl(一次性材料, 0) = 1 And 材料id = 收费细目id_In;
          Exception
            When Others Then
              Null;
          End;
          If Nvl(n_Count, 0) > 0 Then
            d_灭菌效期 := r_Stock.灭菌效期;
            d_灭菌日期 := d_灭菌效期 - n_Count * 30;
          End If;
        End If;
      End If;
    
      Select Nvl(Max(序号), 0) + 1
      Into n_序号
      From 药品收发记录
      Where 单据 = Decode(收费类别_In, '4', 25, 9) And 记录状态 = 1 And NO = No_In;
    
      n_扣率 := Null;
      If 期效_In Is Not Null Or 计价特性_In Is Not Null Then
        n_扣率 := Nvl(期效_In, 0) || Nvl(计价特性_In, 0);
      End If;
    
      --分批药品,如果是只使用了一个批次,则要填写付数
      If n_分批 = 1 And n_当前数量 <> 付数_In * 数次_In Then
        n_Count := 1;
      Else
        n_Count := 0;
      End If;
    
      --修改的原单据号存放在摘要中
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 零售价, 零售金额, 摘要, 填制人,
         填制日期, 费用id, 频次, 单量, 用法, 外观, 扣率, 灭菌效期, 灭菌日期, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
      Values
        (药品收发记录_Id.Nextval, 1, Decode(收费类别_In, '4', 25, 9), No_In, n_序号, 执行部门id_In, 开单部门id_In, 类别id_In, -1, 收费细目id_In,
         n_批次, v_产地, v_批号, d_效期, Decode(n_Count, 1, 1, 付数_In), Decode(n_Count, 1, n_当前数量, n_当前数量 / 付数_In),
         Decode(n_Count, 1, n_当前数量, n_当前数量 / 付数_In), n_当前单价, Round(n_当前单价 * n_当前数量, n_Dec), 药品摘要_In, 操作员姓名_In, 登记时间_In,
         v_费用id, 频次_In, 单量_In, v_用法, v_煎法, n_扣率, d_灭菌效期, d_灭菌日期, n_供药单位id, d_生产日期, v_批准文号, v_商品条码, v_内部条码);
    
      --产生其他出库单
      If 收费类别_In = '4' And (Nvl(备货材料_In, 0) = 1 Or Nvl(n_修正库房id, 0) <> 0) Then
        Begin
          Select Max(a.No), Max(a.序号)
          Into v_其他出库no, n_出库序号
          From 药品收发记录 A, 住院费用记录 B
          Where a.费用id = b.Id And b.No = No_In And 记录性质 = 2 And b.门诊标志 = 门诊标志_In And
                Instr(',8,9,10,21,24,25,26,', ',' || a.单据 || ',') > 0;
        Exception
          When Others Then
            v_其他出库no := Null;
        End;
        If v_其他出库no Is Null Then
          v_其他出库no := Nextno(74, n_虚拟库房id, Null, 1);
        End If;
        If v_其他出库no Is Null Then
          v_Err_Msg := '在生成卫生材料的其他出库单时,获取相关的出库NO有误,请检查出库单的规则是否有误!';
          Raise Err_Item;
        End If;
        If Nvl(病人科室id_In, 0) <> 0 Then
          Select 名称 Into v_部门名称 From 部门表 Where ID = 病人科室id_In;
        End If;
        v_Err_Msg := LPad(' ', 4);
        v_Err_Msg := Substr('病人姓名:' || 姓名_In || v_Err_Msg || '性别:' || 性别_In || v_Err_Msg || '年龄' || 年龄_In || v_Err_Msg ||
                            '门诊号:' || Nvl(标识号_In, '') || v_Err_Msg || '病人科室:' || v_部门名称, 1, 100);
      
        n_出库序号 := Nvl(n_出库序号, 0) + 1;
      
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 零售价, 零售金额, 摘要, 填制人,
           填制日期, 费用id, 频次, 单量, 用法, 外观, 扣率, 灭菌效期, 灭菌日期, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
        Values
          (药品收发记录_Id.Nextval, 1, 21, v_其他出库no, n_出库序号, n_虚拟库房id, 开单部门id_In, 类别id_In, -1, 收费细目id_In, n_批次, v_产地, v_批号,
           d_效期, Decode(n_Count, 1, 1, 付数_In), Decode(n_Count, 1, n_当前数量, n_当前数量 / 付数_In),
           Decode(n_Count, 1, n_当前数量, n_当前数量 / 付数_In), n_当前单价, Round(n_当前单价 * n_当前数量, n_Dec), v_Err_Msg, 操作员姓名_In,
           登记时间_In, v_费用id, 频次_In, 单量_In, v_用法, v_煎法, n_扣率, d_灭菌效期, d_灭菌日期, n_供药单位id, d_生产日期, v_批准文号, v_商品条码, v_内部条码);
      End If;
      v_Err_Msg := '';
      n_总数量  := n_总数量 - n_当前数量;
      n_总金额  := n_总金额 + Round(n_当前数量 * n_当前单价, n_Dec);
    End Loop;
  
    --未发药品记录
    Update 未发药品记录
    Set 病人id = 病人id_In, 姓名 = 姓名_In, 发药窗口 = v_发药窗口
    Where 单据 = Decode(收费类别_In, '4', 25, 9) And NO = No_In And Nvl(库房id, 0) = Nvl(执行部门id_In, 0);
    If Sql%RowCount = 0 Then
      --取身份优先级
      Begin
        Select b.优先级 Into v_优先级 From 病人信息 A, 身份 B Where a.身份 = b.名称(+) And a.病人id = 病人id_In;
      Exception
        When Others Then
          Null;
      End;
      Insert Into 未发药品记录
        (单据, NO, 病人id, 姓名, 优先级, 对方部门id, 库房id, 填制日期, 已收费, 打印状态, 发药窗口)
      Values
        (Decode(收费类别_In, '4', 25, 9), No_In, 病人id_In, 姓名_In, v_优先级, 开单部门id_In, 执行部门id_In, 登记时间_In,
         Decode(划价_In, 1, 0, 1), 0, v_发药窗口);
    End If;
    Zl_Prescription_Type_Update(No_In, 2, 收费细目id_In, 收费类别_In);
  
    --可能分批时价药品分解的批次变了
    If n_时价 = 1 Then
      --只有一个批次时,直接取该批次的单价
      If n_当前数量 <> 付数_In * 数次_In Then
        n_当前单价 := Round(n_总金额 / (付数_In * 数次_In), n_单价小数);
      End If;
      If n_当前单价 <> 标准单价_In Then
        Close c_Stock;
        If 医嘱序号_In Is Null Then
          If 收费类别_In = '4' Then
            v_Err_Msg := '第 ' || 序号_In || ' 行的时价卫生材料"' || v_名称 || '"当前计算单价不一致,请重新输入数量计算！';
          Else
            v_Err_Msg := '第 ' || 序号_In || ' 行的时价药品"' || v_名称 || '"当前计算单价不一致,请重新输入数量计算！';
          End If;
        Else
          If 收费类别_In = '4' Then
            v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现时价卫生材料"' || v_名称 || '"当前计算的单价发生变化。' || Chr(13) || Chr(10) ||
                         '请检查该病人是否同时使用了两笔相同的"' || v_名称 || '"！';
          Else
            v_Err_Msg := '在处理病人"' || 姓名_In || '"时发现时价药品"' || v_名称 || '"当前计算的单价发生变化。' || Chr(13) || Chr(10) ||
                         '请检查该病人是否同时使用了两笔相同的"' || v_名称 || '"！';
          End If;
        End If;
        Raise Err_Item;
      End If;
    End If;
  
    Close c_Stock;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊记帐记录_Insert;
/

--73706:冉俊明,2018-08-01,体检费用不更新病人余额.费用余额
Create Or Replace Procedure Zl_门诊记帐记录_Verify
(
  No_In         门诊费用记录.No%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  序号_In       Varchar2 := Null,
  审核时间_In   门诊费用记录.登记时间%Type := Null
) As
  --功能：审核一张门诊记帐划价单
  --参数：
  --    序号_IN：格式如"1,3,5,7,8",为空表示审核所有未审核的行
  --    审核时间_IN：用于部份需要统一控制或返回时间的地方
  --只读取指定序号的,未审核的部份进行处理
  Cursor c_Bill Is
    Select a.Id, a.病人id, a.实收金额, a.门诊标志, a.收入项目id, a.执行部门id, a.开单部门id, a.病人科室id, a.发药窗口, a.收费类别, Nvl(b.跟踪在用, 0) As 跟踪在用,
           a.医嘱序号
    From 门诊费用记录 A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And a.记录性质 = 2 And a.记录状态 = 0 And a.No = No_In And
          (Instr(',' || 序号_In || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 Or 序号_In Is Null)
    Order By a.序号;

  --审核中包含跟踪在用的未发卫料时，根据参数设置是否自动发料
  Cursor c_Stuff Is
    Select ID, 库房id
    From 药品收发记录 M
    Where NO = No_In And 单据 = 25 And 库房id Is Not Null And 记录状态 = 1 And 审核人 Is Null And Exists
     (Select 1
           From 门诊费用记录 A, 材料特性 B
           Where a.Id = m.费用id + 0 And a.记录性质 = 2 And a.记录状态 = 1 And a.No = No_In And
                 (Instr(',' || 序号_In || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 Or 序号_In Is Null) And
                 a.收费细目id = b.材料id And b.跟踪在用 = 1)
    Order By 库房id, 药品id;

  --
  n_发料号   药品收发记录.汇总发药号%Type;
  n_库房id   药品收发记录.库房id%Type;
  v_收发ids  Varchar2(4000);
  d_Date     Date;
  v_医嘱ids  Varchar2(4000);
  v_发药窗口 药品收发记录.发药窗口%Type;

  Type t_Record Is Record(
    药房id   Number(18),
    发药窗口 Varchar2(10));

  Type t_发药窗口 Is Table Of t_Record;
  c_发药窗口 t_发药窗口 := t_发药窗口();
  n_Step     Number(18);
  n_Havedata Number(2);
  n_Count    Number(18);

Begin
  If 审核时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 审核时间_In;
  End If;

  For r_Bill In c_Bill Loop
  
    --处理发药窗口
    If (r_Bill.收费类别 In ('5', '6', '7') Or r_Bill.收费类别 = '4' And r_Bill.跟踪在用 = 1) Then
      --同一张单据,满足同一药房同一窗口
      v_发药窗口 := Null;
      n_Havedata := 0;
      For n_Step In 1 .. c_发药窗口.Count Loop
        If c_发药窗口(n_Step).药房id = Nvl(r_Bill.执行部门id, 0) Then
          v_发药窗口 := c_发药窗口(n_Step).发药窗口;
          n_Havedata := 1;
          Exit;
        End If;
      End Loop;
    
      If v_发药窗口 Is Null Then
        --同一病人在普通号挂号有效挂号天数内且未发药的且上班的,以最近一次记账窗口为准
        n_Count := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
        If n_Count = 0 Then
          n_Count := 1;
        End If;
      
        Begin
          Select 发药窗口
          Into v_发药窗口
          From (Select 登记时间, 发药窗口
                 From 门诊费用记录 A
                 Where 收费类别 In ('5', '6', '7', '4') And 病人id = r_Bill.病人id And 登记时间 Between Sysdate - n_Count And Sysdate And
                       记录性质 = 2 And 执行部门id = r_Bill.执行部门id And 发药窗口 Is Not Null And Exists
                  (Select 1
                        From 未发药品记录
                        Where a.No = NO And 单据 In (9, 25) And 库房id + 0 = r_Bill.执行部门id And 病人id + 0 = r_Bill.病人id) And
                       Exists
                  (Select 1
                        From 发药窗口
                        Where Nvl(上班否, 0) = 1 And 名称 = a.发药窗口 And Nvl(专家, 0) = 0 And 药房id = r_Bill.执行部门id)
                 Order By 登记时间 Desc)
          Where Rownum <= 1;
        Exception
          When Others Then
            v_发药窗口 := Null;
        End;
        If v_发药窗口 Is Null Then
          v_发药窗口 := Zl_Get发药窗口(r_Bill.执行部门id);
        End If;
      
      End If;
      If n_Havedata = 0 Then
        c_发药窗口.Extend;
        c_发药窗口(c_发药窗口.Count).药房id := r_Bill.执行部门id;
        c_发药窗口(c_发药窗口.Count).发药窗口 := v_发药窗口;
      End If;
    End If;
  
    Update 门诊费用记录
    Set 记录状态 = 1, 发药窗口 = v_发药窗口, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 登记时间 = d_Date --已产生的药品记录的时间不变 
    Where ID = r_Bill.Id;
  
    --药品收发记录.填制日期
    Update 药品收发记录
    Set 填制日期 = Decode(Sign(Nvl(审核日期, d_Date) - d_Date), -1, 填制日期, d_Date)
    Where NO = No_In And 单据 In (9, 25) And 费用id = r_Bill.Id;
  
    --病人余额 
    If Nvl(r_Bill.门诊标志, 0) <> 4 Then
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(r_Bill.实收金额, 0)
      Where 病人id = r_Bill.病人id And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (r_Bill.病人id, 1, 1, r_Bill.实收金额, 0);
      End If;
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(r_Bill.实收金额, 0)
    Where 病人id = r_Bill.病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(r_Bill.病人科室id, 0) And
          Nvl(开单部门id, 0) = Nvl(r_Bill.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(r_Bill.执行部门id, 0) And
          收入项目id + 0 = r_Bill.收入项目id And 来源途径 + 0 = r_Bill.门诊标志;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (r_Bill.病人id, Null, Null, r_Bill.病人科室id, r_Bill.开单部门id, r_Bill.执行部门id, r_Bill.收入项目id, r_Bill.门诊标志,
         Nvl(r_Bill.实收金额, 0));
    End If;
  
    If r_Bill.医嘱序号 Is Not Null Then
      v_医嘱ids := v_医嘱ids || ',' || r_Bill.医嘱序号;
    End If;
  
  End Loop;

  --处理医嘱发送计费状态
  If v_医嘱ids Is Not Null Then
    v_医嘱ids := Substr(v_医嘱ids, 2);
    Zl_医嘱发送_计费状态_Update(0, 2, 1, No_In, v_医嘱ids);
  End If;
  --更新发药窗口
  For n_Step In 1 .. c_发药窗口.Count Loop
    Update 药品收发记录
    Set 发药窗口 = c_发药窗口(n_Step).发药窗口
    Where 库房id = c_发药窗口(n_Step).药房id And 单据 In (9, 25) And NO = No_In And
          费用id + 0 In (Select ID
                       From 门诊费用记录
                       Where 记录性质 = 2 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));
  
    Update 未发药品记录
    Set 发药窗口 = c_发药窗口(n_Step).发药窗口
    Where 库房id = c_发药窗口(n_Step).药房id And 单据 In (9, 25) And NO = No_In;
  End Loop;

  --库房中的药品已全部审核则标为已收费
  Update 未发药品记录
  Set 已收费 = 1, 填制日期 = d_Date
  Where NO = No_In And 单据 = 9 And Nvl(已收费, 0) = 0 And
        Nvl(库房id, 0) Not In
        (Select Distinct Nvl(执行部门id, 0)
         From 门诊费用记录
         Where 记录性质 = 2 And NO = No_In And 收费类别 In ('5', '6', '7') And 记录状态 = 0);

  Update 未发药品记录
  Set 已收费 = 1, 填制日期 = d_Date
  Where NO = No_In And 单据 = 25 And Nvl(已收费, 0) = 0 And
        Nvl(库房id, 0) Not In (Select Distinct Nvl(执行部门id, 0)
                             From 门诊费用记录
                             Where 记录性质 = 2 And NO = No_In And 收费类别 = '4' And 记录状态 = 0);

  --处理跟踪在用卫料自动发料
  If zl_GetSysParameter(92) = '1' Then
    For r_Stuff In c_Stuff Loop
      If n_发料号 Is Null Then
        n_发料号 := Nextno(20);
      End If;
    
      If r_Stuff.库房id <> Nvl(n_库房id, 0) Then
        If Nvl(n_库房id, 0) <> 0 And v_收发ids Is Not Null Then
          v_收发ids := Substr(v_收发ids, 2);
          Zl_药品收发记录_批量发料(v_收发ids, n_库房id, 操作员姓名_In, Sysdate, 1, 操作员姓名_In, n_发料号, 操作员姓名_In);
        End If;
      
        n_库房id  := r_Stuff.库房id;
        v_收发ids := Null;
      End If;
    
      v_收发ids := v_收发ids || '|' || r_Stuff.Id || ',0';
    End Loop;
    If Nvl(n_库房id, 0) <> 0 And v_收发ids Is Not Null Then
      v_收发ids := Substr(v_收发ids, 2);
      Zl_药品收发记录_批量发料(v_收发ids, n_库房id, 操作员姓名_In, Sysdate, 1, 操作员姓名_In, n_发料号, 操作员姓名_In);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊记帐记录_Verify;
/

--73706:冉俊明,2018-08-01,体检费用不更新病人余额.费用余额
Create Or Replace Procedure Zl_门诊转住院_记帐转出
(
  No_In         住院费用记录.No%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  退费时间_In   住院费用记录.发生时间%Type,
  门诊销帐_In   Number := 0
) As
  --门诊销帐_In:0-门诊转住院立即销帐;1-门诊记帐退费模式
  n_Count      Number(5);
  n_实收金额   住院费用记录.实收金额%Type;
  n_病人id     住院费用记录.病人id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  v_开单人     门诊费用记录.开单人%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);
Begin

  Select Count(NO), Sum(实收金额) Into n_Count, n_实收金额 From 门诊费用记录 Where NO = No_In And 记录性质 = 2;
  If n_Count = 0 Then
    v_Err_Msg := '单据' || No_In || '不是记帐单据或因并发原因他人操作了该单据,不能转为住院费用.';
    Raise Err_Item;
  End If;

  Select 病人id, 开单部门id, 开单人
  Into n_病人id, n_开单部门id, v_开单人
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 2 And 记录状态 In (1, 3) And Rownum = 1;

  --处理病人余额
  Begin
    Select Nvl(Sum(实收金额), 0)
    Into n_实收金额
    From 门诊费用记录
    Where NO = No_In And 记录性质 = 2 And 记录状态 In (1, 2, 3) And Nvl(门诊标志, 0) <> 4 And 结帐id Is Null
    Group By NO, 记录性质;
  Exception
    When Others Then
      n_实收金额 := 0;
  End;

  Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) - n_实收金额 Where 病人id = n_病人id And 类型 = 1 And 性质 = 1;
  If Sql%RowCount = 0 And n_实收金额 <> 0 Then
    Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (n_病人id, 1, 1, -1 * n_实收金额, 0);
  End If;

  --处理未结费用
  For v_未结 In (Select 开单部门id, 病人id, 病人科室id, 执行部门id, 收入项目id, 门诊标志, -1 * Nvl(Sum(实收金额), 0) As 实收金额
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 In (1, 2, 3)
               Group By 开单部门id, 病人id, 病人科室id, 执行部门id, 收入项目id, 门诊标志) Loop
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + v_未结.实收金额
    Where 病人id = v_未结.病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(v_未结.病人科室id, 0) And
          Nvl(开单部门id, 0) = Nvl(v_未结.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(v_未结.执行部门id, 0) And 收入项目id + 0 = v_未结.收入项目id And
          来源途径 + 0 = v_未结.门诊标志;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (v_未结.病人id, Null, Null, v_未结.病人科室id, v_未结.开单部门id, v_未结.执行部门id, v_未结.收入项目id, v_未结.门诊标志, v_未结.实收金额);
    End If;
  End Loop;

  --作废费用记录
  Insert Into 门诊费用记录
    (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id, 计算单位,
     付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 执行时间, 操作员编号,
     操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论)
    Select 病人费用记录_Id.Nextval, NO, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id,
           收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, -1 * 应收金额, -1 * 实收金额, 开单部门id,
           开单人, 执行部门id, 划价人, 执行人, -1, 执行时间, 操作员编号_In, 操作员姓名_In, 发生时间, 退费时间_In, 保险项目否, 保险大类id, -1 * 统筹金额, 记帐单id, 摘要, 保险编码,
           是否急诊, 结论
    From 门诊费用记录
    Where NO = No_In And 记录性质 = 2 And 记录状态 = 1;

  --Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 2 And 记录状态 = 1;

  --药品处理(未处理,主要是因为直接转换成相关的药房即可.)
  If Nvl(门诊销帐_In, 0) = 1 Then
    Update 费用审核记录
    Set 记录状态 = 2
    Where 费用id In (Select ID From 门诊费用记录 Where NO = No_In And 记录性质 = 2 And 记录状态 In (1, 3)) And 性质 = 1;
    --作废门诊记录
    Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 2 And 记录状态 = 1;
    For r_Clinic In (Select 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
                            发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, Sum(应收金额) As 应收金额,
                            Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id, 划价人, 记帐单id, 是否急诊, 缴款组id, 发生时间,
                            实际票号
                     From 门诊费用记录
                     Where NO = No_In And 记录性质 = 2 And 记录状态 In (2, 3) And 附加标志 Not In (8, 9)
                     Group By 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
                              发药窗口, 付数, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 开单部门id, 开单人, 执行部门id, 划价人, 记帐单id, 是否急诊, 缴款组id,
                              发生时间, 实际票号
                     Having Sum(数次) <> 0) Loop
      Insert Into 门诊费用记录
        (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
         保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人,
         发生时间, 登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id)
      Values
        (病人费用记录_Id.Nextval, 2, No_In, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 1, r_Clinic.病人id, '',
         r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别, r_Clinic.收费细目id,
         r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数,
         -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.婴儿费, r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价,
         -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 1, r_Clinic.开单部门id, r_Clinic.开单人, r_Clinic.发生时间,
         退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', r_Clinic.是否急诊, r_Clinic.缴款组id);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_记帐转出;
/

--73706:冉俊明,2018-08-01,体检费用不更新病人余额.费用余额
Create Or Replace Procedure Zl_病人未结门诊费用_Recalc(病人id_In 住院费用记录.病人id%Type) As
  v_费别     费别.名称%Type;
  v_No       门诊费用记录.No%Type;
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
    Select 费用余额 Into n_费用余额 From 病人余额 Where 病人id = 病人id_In And 类型 = 1 And 性质 = 1;
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
  n_小数位数 := To_Number(Nvl(zl_GetSysParameter(9), 2));
  For r_Fee In (Select 病人id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 病人科室id, 收费类别, 收费细目id, 计算单位, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目,
                       开单部门id, 开单人, 执行部门id, 发生时间, 操作员编号, 操作员姓名, Nvl(Sum(应收金额), 0) 应收金额, Nvl(Sum(实收金额), 0) 实收金额
                From (Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号,
                              0 As 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目,
                              标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名,
                              结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊
                       From 门诊费用记录
                       Union All
                       Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号,
                              0 As 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目,
                              标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名,
                              结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊
                       From H门诊费用记录)
                Where 病人id = 病人id_In And 记录状态 <> 0 And 记帐费用 = 1
                Group By 病人id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 病人科室id, 收费类别, 收费细目id, 计算单位, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目,
                         开单部门id, 开单人, 执行部门id, 发生时间, 操作员编号, 操作员姓名
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
      v_Thisinfo := r_Fee.开单部门id || r_Fee.开单人 || r_Fee.操作员姓名 || ' ';
      If v_Counter = 0 Or v_Counter = 100 Or v_Thisinfo <> v_Lastinfo Then
        v_No       := Nextno(14);
        v_Counter  := 1;
        v_Lastinfo := v_Thisinfo;
      Else
        v_Counter := v_Counter + 1;
      End If;
    
      Insert Into 门诊费用记录
        (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
         保险大类id, 付数, 数次, 发药窗口, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
         执行部门id, 执行状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 摘要, 是否急诊, 医嘱序号)
      Values
        (病人费用记录_Id.Nextval, 2, v_No, 1, v_Counter, Null, Null, Null, r_Fee.门诊标志, r_Fee.病人id, r_Fee.标识号, r_Fee.姓名,
         r_Fee.性别, r_Fee.年龄, r_Fee.病人科室id, v_费别, r_Fee.收费类别, r_Fee.收费细目id, r_Fee.计算单位, Null, Null, 0, 0, Null,
         r_Fee.加班标志, r_Fee.附加标志, r_Fee.婴儿费, r_Fee.收入项目id, r_Fee.收据费目, 0, 0, n_实收金额, Null, 1, Null, r_Fee.开单部门id,
         r_Fee.开单人, r_Fee.发生时间, d_Sysdate, r_Fee.执行部门id, 0, Null, Null, r_Fee.操作员编号, r_Fee.操作员姓名,
         Decode(v_Counter, 1, '实收重算冲减', ''), 0, Null);
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
    From 门诊费用记录
    Where 病人id = 病人id_In And 记录性质 = 2 And Nvl(门诊标志, 0) <> 4 And 登记时间 = d_Sysdate;
    Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) + n_实收金额 Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1;
    If Sql%RowCount = 0 Then
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
            Nvl(执行部门id, 0) = Nvl(r_Fee.执行部门id, 0) And 收入项目id + 0 = Nvl(r_Fee.收入项目id, 0) And 来源途径 + 0 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, r_Fee.病人病区id, r_Fee.病人科室id, r_Fee.开单部门id, r_Fee.执行部门id, r_Fee.收入项目id, 2, r_Fee.实收金额);
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人未结门诊费用_Recalc;
/

--122793:冉俊明,2018-08-06,修正临床出诊固定安排，新增临床安排并审核后，临时出诊的安排被覆盖掉了的问题
Create Or Replace Procedure Zl1_Auto_Buildingregisterplan
(
  挂号时间_In In Date := Null,
  号源id_In   临床出诊号源.Id%Type := Null
) As
  Pragma Autonomous_Transaction;
  -------------------------------------------------------------------------
  --功能说明：自动生成临床出诊记录
  --          1、根据号源自动生成预约数内的临床出诊记录;
  --          2、预约天数的确定:号源预约天数-->预约方式的天数（取最大)-->系统预约天数
  --入参:挂号时间_IN:NULL时，自动生成;否则只检查指定日期是否生成了出诊记录没有
  --    号源id_In:NULL时处理所有号源，否则只处理指定号源
  -------------------------------------------------------------------------
  n_缺省预约天数 临床出诊号源.预约天数%Type;
  v_操作员姓名   临床出诊安排.操作员姓名%Type;
  d_登记日期     临床出诊安排.登记时间%Type;
  n_安排id       临床出诊安排.Id%Type;
  n_项目id       临床出诊安排.项目id %Type;

  n_记录id   临床出诊记录.Id%Type;
  d_当前日期 临床出诊记录.出诊日期%Type;

  l_固定时段 t_Strlist := t_Strlist();
  n_Count    Number(18);

  n_加预约天数 Number := 0;
  d_开始时间   临床出诊记录.开始时间%Type;
Begin

  Select Max(预约天数) Into n_缺省预约天数 From 预约方式;
  If Nvl(n_缺省预约天数, 0) = 0 Then
    n_缺省预约天数 := To_Number(Nvl(zl_GetSysParameter('挂号允许预约天数'), '0'));
  End If;
  If Nvl(n_缺省预约天数, 0) = 0 Then
    n_缺省预约天数 := 7;
  End If;

  --以半天为单位,如果参数“号源开放时间”在12:00:00-23:59:59期间的，则开放预约天数+1天
  n_加预约天数 := Zl_Fun_Getappointmentdays;

  d_当前日期   := Trunc(Nvl(挂号时间_In, Sysdate));
  d_登记日期   := Sysdate;
  v_操作员姓名 := Zl_Username;

  --第一层循环，号源信息
  For c_号源 In (Select c.Id, c.号类, c.号码, c.项目id, c.科室id, c.医生姓名,
                      Decode(Nvl(c.预约天数, 0), 0, n_缺省预约天数, c.预约天数) + n_加预约天数 As 预约天数, Nvl(b.站点, '-') As 站点,
                      Nvl(c.是否假日换休, 0) As 是否假日换休, Nvl(c.假日控制状态, 0) As 假日控制状态, Nvl(c.排班方式, 0) As 排班方式
               From 临床出诊号源 C, 部门表 B, 人员表 A, 收费项目目录 D
               Where c.科室id = b.Id And c.医生id = a.Id(+) And c.项目id = d.Id And Nvl(c.是否删除, 0) = 0 And
                     Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(a.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (号源id_In Is Null Or c.Id = 号源id_In)
                    --
                     And Exists (Select 1
                      From 临床出诊安排 M, 临床出诊表 N
                      Where m.出诊id = n.Id And m.号源id = c.Id And Nvl(n.排班方式, 0) = 0 And n.发布时间 Is Not Null And
                            m.审核时间 Is Not Null And d_当前日期 <= m.终止时间)) Loop
  
    --检查当前日期所在的安排的收费项目是否为号源中的收费项目，如果不是，则更新号源中的收费项目
    Begin
      Select 项目id
      Into n_项目id
      From (Select a.项目id
             From 临床出诊安排 A, 临床出诊表 B
             Where a.出诊id = b.Id And a.号源id = c_号源.Id And a.审核时间 Is Not Null And d_当前日期 Between a.开始时间 And a.终止时间 And
                   Nvl(b.排班方式, 0) = 0 And b.发布时间 Is Not Null
             Order By a.登记时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        n_项目id := Null;
    End;
    If Nvl(n_项目id, 0) <> 0 Then
      If Nvl(c_号源.项目id, 0) <> n_项目id Then
        Update 临床出诊号源 Set 项目id = n_项目id Where ID = c_号源.Id;
        Commit;
      End If;
    End If;
  
    --第二层循环，出诊日期
    --从头一天开始生成，避免如全日(8:00-7:59)在0:00-7:59没有出诊记录
    --1.未指定号源ID，则是正常生成出诊记录，有出诊记录的日期将不再处理
    --2.指定了号源ID，肯定是发布后新增了临时安排重新生成出诊记录
    For c_日期 In (Select m.日期,
                        Decode(To_Char(m.日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                                '周六', Null) As 星期
                 From (Select Trunc(d_当前日期) + 天数 As 日期
                        From (Select Level - 1 As 天数 From Dual Connect By Level <= c_号源.预约天数 + 1)
                        Where 号源id_In Is Not Null
                        Union All
                        Select Trunc(d_当前日期 - 1) + 天数 As 日期
                        From (Select Level - 1 As 天数 From Dual Connect By Level <= c_号源.预约天数 + 2)
                        Where 号源id_In Is Null And Not Exists
                         (Select 1
                               From 临床出诊记录 A
                               Where a.号源id = c_号源.Id And a.出诊日期 = Trunc(d_当前日期 - 1) + 天数)) M
                 Where 挂号时间_In Is Null Or Trunc(挂号时间_In) = m.日期) Loop
    
      l_固定时段 := t_Strlist();
      --检查当日是否在月/周出诊表中,若在，则不生成出诊记录
      Select Count(1)
      Into n_Count
      From 临床出诊安排 A, 临床出诊表 B
      Where a.出诊id = b.Id And a.号源id = c_号源.Id And c_日期.日期 Between Trunc(a.开始时间) And Trunc(a.终止时间) And
            Nvl(b.排班方式, 0) In (1, 2) And Rownum < 2;
    
      --当前号源为按月/周排班，且当前日期之前已有按月/周排班的出诊记录就不再按固定安排生成出诊记录了
      If n_Count = 0 And Nvl(c_号源.排班方式, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 临床出诊安排 A, 临床出诊表 B
        Where a.出诊id = b.Id And Nvl(b.排班方式, 0) In (1, 2) And a.号源id = c_号源.Id And a.开始时间 < c_日期.日期 And Rownum < 2;
      End If;
    
      If n_Count = 0 Then
        If 号源id_In Is Null Then
          --出诊安排,取最后登记的一个
          Begin
            Select 安排id
            Into n_安排id
            From (Select a.Id As 安排id
                   From 临床出诊安排 A, 临床出诊表 B
                   Where a.号源id = c_号源.Id And a.出诊id = b.Id And Nvl(b.排班方式, 0) = 0 And b.发布时间 Is Not Null And
                         a.审核时间 Is Not Null And c_日期.日期 Between a.开始时间 And a.终止时间
                   Order By a.登记时间 Desc)
            Where Rownum < 2;
          Exception
            When Others Then
              n_安排id := 0;
          End;
        Else
          --如果指定了号源ID，肯定是发布后新增了临时安排重新生成出诊记录，最后登记的一个肯定是本次新增的，
          --只需要处理这个安排即可，不在这个安排有效时间范围内的就不处理
          Begin
            Select 安排id
            Into n_安排id
            From (Select a.Id As 安排id, a.开始时间, a.终止时间, Row_Number() Over(Order By a.登记时间 Desc) As 行号
                   From 临床出诊安排 A, 临床出诊表 B
                   Where a.号源id = c_号源.Id And a.出诊id = b.Id And Nvl(b.排班方式, 0) = 0 And b.发布时间 Is Not Null And
                         a.审核时间 Is Not Null)
            Where 行号 = 1 And c_日期.日期 Between 开始时间 And 终止时间;
          Exception
            When Others Then
              n_安排id := 0;
          End;
        End If;
      
        If Nvl(n_安排id, 0) <> 0 Then
          If 号源id_In Is Not Null Then
            --2.指定了号源ID，肯定是发布后新增了临时安排重新生成出诊记录
            --当日有出诊记录，需要做如下处理
            For c_记录 In (Select a.安排id, a.Id As 记录id, a.出诊日期, a.上班时段, a.是否分时段, a.是否序号控制
                         From 临床出诊记录 A
                         Where a.号源id = c_号源.Id And a.出诊日期 = c_日期.日期) Loop
            
              Select Count(1) Into n_Count From 病人挂号记录 Where 出诊记录id = c_记录.记录id;
              If n_Count = 0 Then
                --2.2.1如果时段不存在预约挂号数据，则删除重新生成
                Zl_临床出诊上班时段_Delete(c_记录.安排id, To_Char(c_记录.出诊日期, 'yyyy-mm-dd'), 1, c_记录.上班时段);
              Else
                --2.2.2如果时段存在预约挂号数据，则只需调整出诊记录的安排ID即可
                Update 临床出诊记录 Set 安排id = n_安排id Where ID = c_记录.记录id;
                l_固定时段.Extend();
                l_固定时段(l_固定时段.Count) := c_记录.上班时段;
              End If;
            End Loop;
          End If;
        
          --检查这天是否出诊
          Select Count(1) Into n_Count From 临床出诊限制 Where 安排id = n_安排id And 限制项目 = c_日期.星期;
          If n_Count = 0 Then
            --如果不存在临床出诊记录，则增加临床出诊记录(时间段为NULL 的空记录)
            Insert Into 临床出诊记录
              (ID, 安排id, 号源id, 出诊日期, 登记人, 登记时间)
              Select 临床出诊记录_Id.Nextval, n_安排id, a.Id As ID, c_日期.日期, v_操作员姓名, d_登记日期 As 登记时间
              From 临床出诊号源 A, 临床出诊安排 B
              Where a.Id = b.号源id And b.Id = n_安排id And Not Exists
               (Select 1 From 临床出诊记录 Where 号源id = a.Id And 出诊日期 = c_日期.日期);
          Else
            For c_记录 In (With c_时间段 As
                            (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间
                            From (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间,
                                          Row_Number() Over(Partition By 时间段 Order By 时间段, 站点 Asc, 号类 Asc) As 组号
                                   From 时间段
                                   Where Nvl(站点, c_号源.站点) = c_号源.站点 And Nvl(号类, c_号源.号类) = c_号源.号类)
                            Where 组号 = 1)
                           Select n_安排id As 安排id, B1.号源id, c_日期.日期 As 出诊日期, m.上班时段, m.Id As 限制id,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                           'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.终止时间, 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.终止时间 <= j.开始时间 Then
                                     1
                                    Else
                                     0
                                  End As 终止时间, Null As 停诊开始时间, Null As 停诊终止时间, Null As 停诊原因,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.缺省时间, j.开始时间), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.缺省时间 < j.开始时间 Then
                                     1
                                    Else
                                     0
                                  End As 缺省预约时间,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.提前时间, j.开始时间), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.开始时间 < j.提前时间 Then
                                     -1
                                    Else
                                     0
                                  End As 提前挂号时间, m.限号数, 0 As 已挂数, m.限约数, 0 As 已约数, 0 As 其中已接收, m.是否序号控制, m.是否分时段, m.预约控制,
                                  m.是否独占, B1.项目id, B1.医生id, B1.医生姓名, Null As 替诊医生id, Null As 替诊医生姓名, m.分诊方式, m.诊室id,
                                  0 As 是否锁定, 0 As 是否临时出诊, v_操作员姓名 As 操作员姓名, d_登记日期 As 登记时间, c_日期.星期 As 限制项目
                           From 临床出诊安排 B1, 临床出诊限制 M, c_时间段 J
                           Where B1.Id = n_安排id And B1.Id = m.安排id And m.限制项目 = c_日期.星期 And m.上班时段 = j.时间段 And
                                 To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                         'yyyy-mm-dd hh24:mi:ss') >= B1.开始时间 And Not Exists
                            (Select 1 From Table(l_固定时段) Where Column_Value = m.上班时段)) Loop
            
              Select 临床出诊记录_Id.Nextval Into n_记录id From Dual;
              Insert Into 临床出诊记录
                (ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 停诊开始时间, 停诊终止时间, 停诊原因, 缺省预约时间, 提前挂号时间, 限号数, 已挂数, 限约数, 已约数,
                 其中已接收, 是否序号控制, 是否分时段, 预约控制, 是否独占, 项目id, 科室id, 医生id, 医生姓名, 替诊医生id, 替诊医生姓名, 分诊方式, 诊室id, 是否锁定, 是否临时出诊, 登记人,
                 登记时间, 是否发布)
              Values
                (n_记录id, c_记录.安排id, c_记录.号源id, c_记录.出诊日期, c_记录.上班时段, c_记录.开始时间, c_记录.终止时间, c_记录.停诊开始时间, c_记录.停诊终止时间,
                 c_记录.停诊原因, c_记录.缺省预约时间, c_记录.提前挂号时间, c_记录.限号数, c_记录.已挂数, c_记录.限约数, c_记录.已约数, c_记录.其中已接收, c_记录.是否序号控制,
                 c_记录.是否分时段, c_记录.预约控制, c_记录.是否独占, c_记录.项目id, c_号源.科室id, c_记录.医生id, c_记录.医生姓名, c_记录.替诊医生id, c_记录.替诊医生姓名,
                 c_记录.分诊方式, c_记录.诊室id, c_记录.是否锁定, c_记录.是否临时出诊, c_记录.操作员姓名, d_登记日期, 1);
            
              d_开始时间 := c_记录.开始时间;
              --插入临床出诊序号控制
              If Nvl(c_记录.是否分时段, 0) = 1 And Nvl(c_记录.是否序号控制, 0) = 1 Then
                --分时段且启用序号控制，使用"预约顺序号"记录"是否预约"
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 预约顺序号)
                  Select n_记录id, 序号,
                         To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When d_开始时间 > To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                             1
                            Else
                             0
                          End,
                         To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When d_开始时间 >= To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                             1
                            Else
                             0
                          End, 限制数量, 是否预约, 是否预约
                  From 临床出诊时段
                  Where 限制id = c_记录.限制id;
              Else
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
                  Select n_记录id, 序号,
                         To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'),
                                 'yyyy-mm-dd hh24:mi:ss') + Case
                           When d_开始时间 > To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'),
                                                 'yyyy-mm-dd hh24:mi:ss') Then
                            1
                           Else
                            0
                         End,
                         To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'),
                                 'yyyy-mm-dd hh24:mi:ss') + Case
                           When d_开始时间 >= To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') Then
                            1
                           Else
                            0
                         End, 限制数量, 是否预约
                  From 临床出诊时段
                  Where 限制id = c_记录.限制id;
              End If;
            
              --插入合作单位挂号控制记录
              Insert Into 临床出诊挂号控制记录
                (类型, 性质, 名称, 记录id, 序号, 控制方式, 数量)
                Select 类型, 性质, 名称, n_记录id, 序号, 控制方式, 数量
                From 临床出诊挂号控制
                Where 限制id = c_记录.限制id;
            
              --插入临床出诊诊室记录
              Insert Into 临床出诊诊室记录
                (记录id, 诊室id)
                Select n_记录id, 诊室id From 临床出诊诊室 Where 限制id = c_记录.限制id;
            End Loop;
          
            --根据停诊安排和法定节假日调整出诊记录的出诊/预约情况
            Zl_Clinicvisitmodify(c_号源.Id, n_安排id, c_日期.日期, v_操作员姓名, d_登记日期);
          End If;
        End If;
      End If;
      --一天一提交
      Commit;
    End Loop;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Auto_Buildingregisterplan;
/

--128221:秦龙,2018-07-27,处理预调价后入库成本价显示异常
Create Or Replace Procedure Zl_药品收发记录_成本价调价
(
  药品id_In   In 药品收发记录.药品id%Type,
  执行时间_In In 成本价调价信息.执行日期%Type := Null
) As
  Adjustdate     Date; --调价时间
  n_序号         Number(8);
  n_Stockid      药品收发记录.库房id%Type;
  n_入出类别id   药品收发记录.入出类别id%Type;
  n_入出系数     药品收发记录.入出系数%Type;
  n_收发id       药品收发记录.Id%Type;
  n_调整额       药品收发记录.零售金额%Type;
  n_供应商id     药品收发记录.供药单位id%Type;
  v_No           药品收发记录.No%Type;
  v_应付id       应付记录.Id%Type; --应付记录的ID
  v_应付no       应付记录.No%Type;
  n_原成本价     药品收发记录.成本价%Type;
  n_新成本价     药品收发记录.成本价%Type;
  n_Run          Number(1);
  v_Count        Number(1) := 0;
  v_调价id       成本价调价信息.Id%Type;
  n_平均成本价   药品收发记录.成本价%Type;
  n_流通金额小数 Number;
  v_调价汇总号   成本价调价信息.调价汇总号%Type;

  Cursor c_Stock Is --当前库存
    Select 上次供应商id, a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, a.上次批号, a.效期, a.上次产地, a.灭菌效期,
           Decode(Sign(Nvl(a.批次, 0)), 1, a.上次采购价, a.平均成本价) As 原成本价
    From 药品库存 A
    Where a.性质 = 1 And a.药品id = 药品id_In
    Order By a.库房id;

  v_Stock c_Stock%RowType;

  Cursor c_Costadjust Is --成本价调价信息
    Select a.库房id, a.药品id, Nvl(a.批次, 0) 批次, a.上次供应商id, Nvl(a.实际数量, 0) As 实际数量, a.实际金额, a.实际差价, a.上次产地 As 产地,
           a.上次批号 As 批号, a.效期, a.上次生产日期 As 生产日期, a.批准文号, a.灭菌效期, a.平均成本价 As 原成本价, b.新成本价, b.发票号, b.发票日期, b.发票金额,
           Nvl(a.上次采购价, 0) As 上次采购价, b.Id As 调价id
    From 药品库存 A, 成本价调价信息 B
    Where a.药品id = b.药品id And Nvl(a.上次供应商id, 0) = Nvl(b.供药单位id, 0) And a.库房id = b.库房id And Nvl(a.批次, 0) = Nvl(b.批次, 0) And
          a.性质 = 1 And b.执行日期 Is Null And a.药品id = 药品id_In
    Order By a.库房id;

  v_Costadjust c_Costadjust%RowType;

  Cursor c_Pay Is --应付管理
    Select Distinct a.供药单位id, a.药品id, a.发票号, a.发票日期, a.发票金额, b.名称, b.计算单位, b.规格
    From 成本价调价信息 A, 收费项目目录 B
    Where a.药品id = b.Id And a.应付款变动 = 1 And a.药品id = 药品id_In And a.供药单位id Is Not Null
    Order By a.供药单位id;

  v_Pay c_Pay%RowType;
Begin
  Adjustdate := Sysdate;
  n_Stockid  := 0;
  n_Run      := 0;

  --判断是否存在无库存调价
  Begin
    Select ID, 新成本价, 调价汇总号
    Into v_调价id, n_新成本价, v_调价汇总号
    From 成本价调价信息
    Where 执行日期 Is Null And Nvl(库房id, 0) = 0 And 药品id = 药品id_In;
  Exception
    When Others Then
      v_调价id   := 0;
      n_新成本价 := Null;
	  v_调价汇总号 := Null;
  End;

  --取流通业务精度位数
  --类别:1-药品 2-卫材
  --内容：2-零售价 4-金额
  --单位：药品:1-售价 5-金额单位
  Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;

  --无库存调价
  If v_调价id > 0 Then
    --根据当前库存重新产生调价信息
    For v_Stock In c_Stock Loop
      Zl_成本价调价信息_Insert(v_Stock.上次供应商id, v_Stock.库房id, v_Stock.药品id, v_Stock.批次, v_Stock.上次批号, v_Stock.效期, v_Stock.上次产地,
                        v_Stock.灭菌效期, v_Stock.原成本价, n_新成本价, Null, Null, Null, 0, v_调价汇总号);
      v_Count := v_Count + 1;
    End Loop;
  
    If v_Count > 0 Then
      --如果当前有库存记录，则删除无库存调价记录
      Delete 成本价调价信息 Where ID = v_调价id;
    Else
      Update 成本价调价信息 Set 执行日期 = Decode(执行时间_In, Null, Adjustdate, 执行时间_In) Where ID = v_调价id;
    
      Update 药品规格 Set 成本价 = n_新成本价 Where 药品id = 药品id_In And 成本价 <> n_新成本价;
    End If;
  End If;

  Select b.Id, b.系数
  Into n_入出类别id, n_入出系数
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 5 And Rownum < 2;

  For v_Costadjust In c_Costadjust Loop
    n_Run := 1;
    If n_Stockid <> v_Costadjust.库房id Then
      n_序号    := 1;
      n_Stockid := v_Costadjust.库房id;
      v_No      := Nextno(25, n_Stockid);
    Else
      n_序号 := n_序号 + 1;
    End If;
  
    --产生库存差价调整单
    Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
  
    If v_Costadjust.实际数量 = 0 And (Nvl(v_Costadjust.实际金额, 0) <> 0 Or Nvl(v_Costadjust.实际差价, 0) <> 0) Then
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期, 审核人,
         审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率, 灭菌效期)
      Values
        (n_收发id, 1, 5, v_No, n_序号, v_Costadjust.库房id, n_入出类别id, v_Costadjust.上次供应商id, n_入出系数, v_Costadjust.药品id,
         v_Costadjust.批次, v_Costadjust.产地, v_Costadjust.批号, v_Costadjust.效期, 0, v_Costadjust.实际金额, v_Costadjust.实际差价, 0,
         '成本价调价', Zl_Username, Adjustdate, Zl_Username, Adjustdate, v_Costadjust.生产日期, v_Costadjust.批准文号,
         v_Costadjust.新成本价, 1, v_Costadjust.原成本价, v_Costadjust.灭菌效期);
    
      --更新库存
      Zl_药品库存_Update(n_收发id);
    
      Update 药品规格 Set 成本价 = v_Costadjust.新成本价 Where 药品id = v_Costadjust.药品id;
    
      --更新成本价调价信息
      Update 成本价调价信息
      Set 批号 = v_Costadjust.批号, 效期 = v_Costadjust.效期, 产地 = v_Costadjust.产地, 灭菌效期 = v_Costadjust.灭菌效期, 收发id = n_收发id,
          执行日期 = Decode(执行时间_In, Null, Adjustdate, 执行时间_In)
      Where ID = v_Costadjust.调价id;
    Else
      --n_调整额   := (v_Costadjust.实际金额 - v_Costadjust.实际差价) - Round(v_Costadjust.新成本价 * v_Costadjust.实际数量, 2);
      n_调整额   := Round((v_Costadjust.原成本价 - v_Costadjust.新成本价) * v_Costadjust.实际数量, n_流通金额小数);
      n_原成本价 := v_Costadjust.原成本价;
      If n_原成本价 = 0 Then
        n_原成本价 := v_Costadjust.上次采购价;
      End If;
    
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期, 审核人,
         审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率, 灭菌效期)
      Values
        (n_收发id, 1, 5, v_No, n_序号, v_Costadjust.库房id, n_入出类别id, v_Costadjust.上次供应商id, n_入出系数, v_Costadjust.药品id,
         v_Costadjust.批次, v_Costadjust.产地, v_Costadjust.批号, v_Costadjust.效期, v_Costadjust.实际数量, v_Costadjust.实际金额,
         v_Costadjust.实际差价, n_调整额, '成本价调价', Zl_Username, Adjustdate, Zl_Username, Adjustdate, v_Costadjust.生产日期,
         v_Costadjust.批准文号, v_Costadjust.新成本价, 1, n_原成本价, v_Costadjust.灭菌效期);
    
      --更新库存
      Zl_药品库存_Update(n_收发id, 0);
    
      Update 药品规格
      Set 成本价 = v_Costadjust.新成本价
      Where 药品id = v_Costadjust.药品id And 成本价 <> v_Costadjust.新成本价;
    
      --更新成本价调价信息
      Update 成本价调价信息
      Set 批号 = v_Costadjust.批号, 效期 = v_Costadjust.效期, 产地 = v_Costadjust.产地, 灭菌效期 = v_Costadjust.灭菌效期, 原成本价 = n_原成本价,
          收发id = n_收发id, 执行日期 = Decode(执行时间_In, Null, Adjustdate, 执行时间_In)
      Where ID = v_Costadjust.调价id;
    End If;
  End Loop;

  --产生应付记录及修改应付余额
  If n_Run = 1 Then
    For v_Pay In c_Pay Loop
      v_应付no := Nextno(67);
    
      Select 应付记录_Id.Nextval Into v_应付id From Dual;
    
      Insert Into 应付记录
        (ID, 记录性质, 记录状态, 单位id, NO, 系统标识, 发票号, 发票日期, 发票金额, 品名, 规格, 填制人, 填制日期, 摘要)
      Values
        (v_应付id, 1, 1, v_Pay.供药单位id, v_应付no, 1, v_Pay.发票号, v_Pay.发票日期, v_Pay.发票金额, v_Pay.名称, v_Pay.规格, Zl_Username,
         Adjustdate, '成本价调价自动产生应付款变动记录');
    
      Update 应付余额 Set 金额 = Nvl(金额, 0) + Nvl(v_Pay.发票金额, 0) Where 单位id = v_Pay.供药单位id And 性质 = 1;
    
      If Sql%NotFound Then
        Insert Into 应付余额 (单位id, 性质, 金额) Values (v_Pay.供药单位id, 1, Nvl(v_Pay.发票金额, 0));
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_成本价调价;
/

--128308:秦龙,2018-07-25,处理药品用法用量勾选应用于当前分类
CREATE OR REPLACE Procedure Zl_用法用量_Update
(
  药名id_In     In 诊疗用法用量.项目id%Type,
  过敏试验id_In In Varchar2, --以"|"分隔的过敏实验的内容
  处方限量_In   In 药品特性.处方限量%Type,
  疗程_In       In 诊疗用法用量.疗程%Type,
  用法用量_In   In Varchar2, --以"|"分隔的用法用量内容，每条记录按"用法ID^频次^成人剂量^小儿剂量^医生嘱托"组织
  方式_In       In Number := 0, --0-诊疗项目本身,1-当前类别;2-特定分类项目
  类别_In       In Varchar2 := '0',
  分类id_In     In 诊疗项目目录.分类id%Type := 0
) Is
  v_Records  Varchar2(4000);
  v_Currrec  Varchar2(1000);
  v_Fields   Varchar2(1000);
  v_用法id   诊疗用法用量.用法id%Type;
  v_频次     诊疗用法用量.频次%Type;
  v_成人剂量 诊疗用法用量.成人剂量%Type;
  v_小儿剂量 诊疗用法用量.小儿剂量%Type;
  v_医生嘱托 诊疗用法用量.医生嘱托%Type;
  v_Ddd值    诊疗用法用量.Ddd值%Type;
  v_性质     诊疗用法用量.性质%Type;
  v_是否皮试 药品特性.是否皮试%Type;

  Cursor c_Item Is
    Select i.Id
    From 诊疗项目目录 I, 药品特性 T, (Select 药品剂型 From 药品特性 Where 药名id = 药名id_In) C
    Where i.Id = t.药名id And t.药品剂型 = c.药品剂型 And i.分类id = 分类id_In And i.Id <> 药名id_In;
Begin
  For r_Item In (Select ID
                 From 诊疗项目目录
                 Where (方式_In = 0 And ID = 药名id_In) Or (方式_In = 1 And 类别 = 类别_In) Or
                       (分类id In (Select ID From 诊疗分类目录 Start With ID = 分类id_In Connect By Prior ID = 上级id))) Loop

    If 过敏试验id_In Is Not Null Then
      v_是否皮试 := 1;
    Else
      v_是否皮试 := 0;
    End If;

    Update 药品特性 Set 处方限量 = 处方限量_In, 是否皮试 = v_是否皮试 Where 药名id = r_Item.Id;

    Delete From 诊疗用法用量 Where 项目id = r_Item.Id And 性质 = 0;

    v_Records := 过敏试验id_In;

    While v_Records Is Not Null Loop
      v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields  := v_Currrec;
      v_用法id  := To_Number(v_Fields);
      Insert Into 诊疗用法用量 (项目id, 性质, 用法id) Values (r_Item.Id, 0, v_用法id);
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;

    Delete From 诊疗用法用量 Where 项目id = r_Item.Id And 性质 > 0;
    If 用法用量_In Is Null Then
      v_Records := Null;
    Else
      v_Records := 用法用量_In || '|';
    End If;
    v_性质 := 0;
    While v_Records Is Not Null Loop
      v_Currrec  := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields   := v_Currrec;
      v_用法id   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_频次     := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_成人剂量 := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_小儿剂量 := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_医生嘱托 := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_Ddd值    := To_Number(v_Fields);
      v_性质     := v_性质 + 1;
      Insert Into 诊疗用法用量
        (项目id, 性质, 用法id, 频次, 成人剂量, 小儿剂量, 医生嘱托, 疗程, Ddd值)
      Values
        (r_Item.Id, v_性质, v_用法id, v_频次, v_成人剂量, v_小儿剂量, v_医生嘱托, 疗程_In, v_Ddd值);
      If 分类id_In <> 0 Then
        For t_Item In c_Item Loop
          delete from 诊疗用法用量 where 项目id=t_item.id  and 用法id=v_用法id and 性质>0;
          Insert Into 诊疗用法用量 (项目id, 性质, 用法id, 频次) Values (t_Item.Id, v_性质, v_用法id, v_频次);
        End Loop;
      End If;
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_用法用量_Update;
/

--105711:李业庆,2019-01-11,出库库存平均成本价处理
--129306:刘涛,2018-07-24,取消过程中分批卫材库存不足禁止
Create Or Replace Procedure Zl_材料移库_Verify
(
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  对方部门id_In In 药品收发记录.对方部门id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  产地_In       In 药品收发记录.产地%Type,
  出批次_In     In 药品收发记录.批次%Type,
  填写数量_In   In 药品收发记录.填写数量%Type,
  实际数量_In   In 药品收发记录.实际数量%Type,
  成本价_In     In 药品收发记录.成本价%Type,
  成本金额_In   In 药品收发记录.成本金额%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  差价_In       In 药品收发记录.差价%Type,
  出类别id_In   In 药品收发记录.入出类别id%Type,
  入类别id_In   In 药品收发记录.入出类别id%Type,
  No_In         In 药品收发记录.No%Type,
  审核人_In     In 药品收发记录.审核人%Type,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
  审核日期_In   In 药品收发记录.审核日期%Type := Null,
  移库单_In     In Number := 1,
  零售价_In     In 药品收发记录.零售价%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg    Varchar2(500);
  v_编码       收费项目目录.编码%Type;
  v_负成本计算 Zlparameters.参数值%Type;
  v_批准文号   药品库存.批准文号%Type;
  n_实价卫材   收费项目目录.是否变价%Type;

  n_入批次       药品收发记录.批次%Type := Null;
  n_实际库存金额 药品库存.实际金额%Type;
  n_实际库存差价 药品库存.实际差价%Type;
  n_出库差价     药品库存.实际差价%Type;
  n_成本价       药品收发记录.成本价%Type;
  n_成本金额     药品收发记录.成本金额%Type;
  n_实际数量     药品库存.实际数量%Type;
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_上次生产日期 药品库存.上次生产日期%Type;
  n_零售价       药品收发记录.零售价%Type;
  n_小数         Number;
  v_上次扣率     药品库存.上次扣率%Type;
  v_商品条码     药品库存.商品条码%Type;
  v_内部条码     药品库存.内部条码%Type;
  n_平均成本价   药品库存.平均成本价%Type;
Begin
  --获取金额小数位数
  Select Nvl(精度, 2) Into n_小数 From 药品卫材精度 Where 类别 = 2 And 内容 = 4 And 单位 = 5;
  Select zl_GetSysParameter(120) Into v_负成本计算 From Dual;

  --由于移库处理允许在审核时改变实际数量，
  --所以首先对实际数量和其他相应的字段进行更新。
  Begin
    Select Nvl(实际金额, 0), Nvl(实际差价, 0), Nvl(实际数量, 0), Nvl(上次扣率, 100), 商品条码, 内部条码
    Into n_实际库存金额, n_实际库存差价, n_实际数量, v_上次扣率, v_商品条码, v_内部条码
    From 药品库存
    Where 药品id = 材料id_In And Nvl(批次, 0) = 出批次_In And 库房id = 库房id_In And 性质 = 1 And Rownum = 1;
  Exception
    When Others Then
      n_实际库存金额 := 0;
      n_实际数量     := 0;
      v_上次扣率     := 100;
      v_商品条码     := Null;
      v_内部条码     := Null;
  End;
  Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;

  If 成本价_In Is Null Then
    --成本价为空
    n_成本价   := Round(Zl_Fun_Getoutcost(材料id_In, 出批次_In, 库房id_In), 7);
    n_成本金额 := Round(n_成本价 * 实际数量_In, n_小数);
    n_出库差价 := Round(零售金额_In - n_成本金额, n_小数);
  Else
    n_成本价   := 成本价_In;
    n_成本金额 := 成本金额_In;
    n_出库差价 := 差价_In;
  End If;

  Update 药品收发记录
  Set 审核人 = Nvl(审核人_In, 审核人), 审核日期 = 审核日期_In, 实际数量 = 实际数量_In, 成本价 = n_成本价, 成本金额 = n_成本金额, 零售价 = 零售价_In, 零售金额 = 零售金额_In,
      差价 = n_出库差价, 扣率 = v_上次扣率, 商品条码 = v_商品条码, 内部条码 = v_内部条码
  Where NO = No_In And 单据 = 19 And 药品id = 材料id_In And 记录状态 = 1 And 序号 In (序号_In, 序号_In + 1) And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;

  --取入类别的批次和上次供应商
  Select 批次, 供药单位id, 生产日期, 批准文号, 零售价
  Into n_入批次, n_上次供应商id, n_上次生产日期, v_批准文号, n_零售价
  From 药品收发记录
  Where NO = No_In And 单据 = 19 And 记录状态 = 1 And 序号 = 序号_In + 1;

  --更改入类别的材料库存的相应数据

  Update 药品库存
  Set 可用数量 = Nvl(可用数量, 0) + 实际数量_In, 实际数量 = Nvl(实际数量, 0) + 实际数量_In, 实际金额 = Nvl(实际金额, 0) + 零售金额_In,
      实际差价 = Nvl(实际差价, 0) + n_出库差价, 上次采购价 = Decode(上次采购价, Null, n_成本价, 0, n_成本价, 上次采购价), 上次批号 = Nvl(批号_In, 上次批号),
      上次产地 = Nvl(产地_In, 上次产地), 灭菌效期 = Nvl(灭菌效期_In, 灭菌效期), 商品条码 = v_商品条码, 内部条码 = v_内部条码
  Where 库房id = 对方部门id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(n_入批次, 0) And 性质 = 1;

  If Sql%NotFound Then
    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次采购价, 上次批号, 上次产地, 效期, 灭菌效期, 上次供应商id, 上次生产日期, 批准文号, 零售价, 上次扣率, 商品条码,
       内部条码, 平均成本价)
    Values
      (对方部门id_In, 材料id_In, n_入批次, 1, 实际数量_In, 实际数量_In, 零售金额_In, n_出库差价, n_成本价, 批号_In, 产地_In, 效期_In, 灭菌效期_In, n_上次供应商id,
       n_上次生产日期, v_批准文号, Decode(n_实价卫材, 1, Decode(Nvl(n_入批次, 0), 0, Null, n_零售价), Null), v_上次扣率, v_商品条码, v_内部条码, n_成本价);
  End If;

  Delete From 药品库存
  Where 库房id = 对方部门id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  --入库类别：不分批的重新计算平均成本价
  If Nvl(n_入批次, 0) = 0 Then
    Update 药品库存
    Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
    Where 药品id = 材料id_In And Nvl(批次, 0) = Nvl(n_入批次, 0) And 库房id = 对方部门id_In And Nvl(实际数量, 0) <> 0 And 性质 = 1;
    If Sql%NotFound Then
      Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = 材料id_In;
      Update 药品库存
      Set 平均成本价 = n_平均成本价
      Where 药品id = 材料id_In And 库房id = 对方部门id_In And Nvl(批次, 0) = Nvl(n_入批次, 0) And 性质 = 1;
    End If;
  End If;

  --更改出类别的材料库存的相应数据

  Update 药品库存
  Set 实际数量 = Nvl(实际数量, 0) - 实际数量_In, 实际金额 = Nvl(实际金额, 0) - 零售金额_In, 实际差价 = Nvl(实际差价, 0) - n_出库差价
  Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(出批次_In, 0) And 性质 = 1;

  If Sql%NotFound Then
    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次批号, 上次产地, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次生产日期, 批准文号, 零售价, 上次扣率, 商品条码,
       内部条码, 平均成本价)
    Values
      (库房id_In, 材料id_In, 出批次_In, 1, 0, -实际数量_In, -零售金额_In, -n_出库差价, 批号_In, 产地_In, 效期_In, 灭菌效期_In, n_上次供应商id, n_成本价,
       n_上次生产日期, v_批准文号, Decode(n_实价卫材, 1, Decode(Nvl(出批次_In, 0), 0, Null, n_零售价), Null), v_上次扣率, v_商品条码, v_内部条码, n_成本价);
  End If;

  --出库房，平均成本价为空时需要重新计算库存表中的平均成本价
  Update 药品库存
  Set 平均成本价 = n_成本价
  Where 药品id = 材料id_In And Nvl(批次, 0) = Nvl(出批次_In, 0) And 库房id = 库房id_In And 性质 = 1 And Nvl(平均成本价, 0) = 0;

  Delete From 药品库存
  Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料移库_Verify;
/

--129306:刘涛,2018-07-24,取消过程在分批材料库存不足禁止
Create Or Replace Procedure Zl_材料其他出库_Insert
(
  入出类别id_In In 药品收发记录.入出类别id%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  批次_In       In 药品收发记录.批次%Type,
  填写数量_In   In 药品收发记录.填写数量%Type,
  成本价_In     In 药品收发记录.成本价%Type,
  成本金额_In   In 药品收发记录.成本金额%Type,
  零售价_In     In 药品收发记录.零售价%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  差价_In       In 药品收发记录.差价%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  产地_In       In 药品收发记录.产地%Type := Null,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  外调价_In     In 药品收发记录.单量%Type,
  外调单位_In   In 药品收发记录.发药窗口%Type,
  增值税率_In   In Number := Null
) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(500);
  v_编码     收费项目目录.编码%Type;
  v_批准文号 药品库存.批准文号%Type;

  n_可用数量     药品库存.可用数量%Type;
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_入出系数     药品收发记录.入出系数%Type; --收发ID
  n_实价卫材     收费项目目录.是否变价%Type;
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

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 供药单位id, 生产日期, 批准文号, 单量, 发药窗口, 频次, 商品条码, 内部条码)
  Values
    (药品收发记录_Id.Nextval, 1, 21, No_In, 序号_In, 库房id_In, 入出类别id_In, n_入出系数, 材料id_In, 批次_In, 产地_In, 批号_In, 效期_In, 灭菌效期_In,
     填写数量_In, 填写数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, n_上次供应商id, d_上次生产日期, v_批准文号,
     外调价_In, 外调单位_In, 增值税率_In, v_商品条码, v_内部条码);

  Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;

  --同时更新库存数
  Update 药品库存
  Set 可用数量 = Nvl(可用数量, 0) - 填写数量_In
  Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 性质 = 1;

  --不插入批次是因为批次材料不够，不准出库
  If Sql%NotFound Then
    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
    Values
      (库房id_In, 材料id_In, Nvl(批次_In, 0), 1, -填写数量_In, 效期_In, 灭菌效期_In, n_上次供应商id, 成本价_In, 批号_In, d_上次生产日期, 产地_In, v_批准文号,
       Decode(n_实价卫材, 1, Decode(Nvl(批次_In, 0), 0, Null, 零售价_In), Null));
  End If;

  Delete From 药品库存
  Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料其他出库_Insert;
/

--129306:刘涛,2018-07-24,取消过程中分批材料库存不足禁止
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
  申购数量_In   In 药品收发记录.单量%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  v_编码         收费项目目录.编码%Type;
  v_批准文号     药品库存.批准文号%Type;
  v_下库存       Zlparameters.参数值%Type;
  n_实价卫材     收费项目目录.是否变价%Type;
  n_入出系数     药品收发记录.入出系数%Type; --收发ID
  n_可用数量     药品库存.可用数量%Type;
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_收发id       药品收发记录.Id%Type;
  v_商品条码     药品库存.商品条码%Type;
  v_内部条码     药品库存.内部条码%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_下库存 From Dual;
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

  Select 药品收发记录_Id.Nextval Into n_收发id From Dual;

  --插入类别为出的那一笔
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价,
     摘要, 填制人, 填制日期, 供药单位id, 生产日期, 批准文号, 领用人, 商品条码, 内部条码, 单量)
  Values
    (n_收发id, 1, 20, No_In, 序号_In, 库房id_In, 对方部门id_In, 入出类别id_In, n_入出系数, 材料id_In, 批次_In, 产地_In, 批号_In, 效期_In, 灭菌效期_In,
     填写数量_In, 填写数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, n_上次供应商id, d_上次生产日期, v_批准文号,
     领用人_In, v_商品条码, v_内部条码, 申购数量_In);

  --同时更新库存数
  If v_下库存 = 1 Then
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) - 填写数量_In
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 性质 = 1;
  
    --不插入批次是因为批次材料不够，不准出库
    If Sql%NotFound Then
      Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;
    
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
      Values
        (库房id_In, 材料id_In, Nvl(批次_In, 0), 1, -填写数量_In, 效期_In, 灭菌效期_In, n_上次供应商id, 成本价_In, 批号_In, d_上次生产日期, 产地_In,
         v_批准文号, Decode(n_实价卫材, 1, Decode(Nvl(批次_In, 0), 0, Null, 零售价_In), Null));
    End If;
  
    Delete From 药品库存
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  End If;

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

--129306:刘涛,2018-07-24,过程取消分批材料库存不足禁止
Create Or Replace Procedure Zl_材料移库_Insert
(
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  对方部门id_In In 药品收发记录.对方部门id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  批次_In       In 药品收发记录.批次%Type,
  填写数量_In   In 药品收发记录.填写数量%Type,
  实际数量_In   In 药品收发记录.实际数量%Type,
  成本价_In     In 药品收发记录.成本价%Type,
  成本金额_In   In 药品收发记录.成本金额%Type,
  零售价_In     In 药品收发记录.零售价%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  差价_In       In 药品收发记录.差价%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  产地_In       In 药品收发记录.产地%Type := Null,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  填制日期_In   In 药品收发记录.填制日期%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(100);
  v_下库存   Zlparameters.参数值%Type;
  v_编码     收费项目目录.编码%Type;
  v_批准文号 药品库存.批准文号%Type;

  n_可用数量     药品库存.可用数量%Type;
  n_Id           药品收发记录.Id%Type; --收发ID
  n_入的类别id   药品收发记录.入出类别id%Type; --入出类别ID
  n_出的类别id   药品收发记录.入出类别id%Type; --入出类别ID
  n_批次         药品收发记录.批次%Type := Null; --主要针对入库中实行分批核算的材料
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_上次生产日期 药品库存.上次生产日期%Type;
  v_商品条码     药品库存.商品条码%Type;
  v_内部条码     药品库存.内部条码%Type;
  n_实价卫材     收费项目目录.是否变价%Type;
  n_是否分批     Integer; --判断入库是否分批核算   1:分批；0：不分批
  n_库房分批     Integer; --判断入库是否分批核算   1:分批；0：不分批
  n_在用分批     Integer; --判断入库是否分批核算   1:分批；0：不分批

  n_Records Number;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_下库存 From Dual;

  Begin
    Select 可用数量, Decode(上次供应商id, 0, Null, 上次供应商id), 上次生产日期, 批准文号, 商品条码, 内部条码
    Into n_可用数量, n_上次供应商id, n_上次生产日期, v_批准文号, v_商品条码, v_内部条码
    From 药品库存
    Where 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 库房id = 库房id_In And 性质 = 1 And Rownum = 1;
  Exception
    When Others Then
      n_可用数量     := 0;
      n_上次供应商id := Null;
  End;

  --首先找出入和出的类别ID
  Select b.Id
  Into n_入的类别id
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 34 And b.系数 = 1 And Rownum < 2;

  Select b.Id
  Into n_出的类别id
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 34 And b.系数 = -1 And Rownum < 2;

  --插入类别为出的那一笔
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价,
     摘要, 填制人, 填制日期, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
  Values
    (药品收发记录_Id.Nextval, 1, 19, No_In, 序号_In, 库房id_In, 对方部门id_In, n_出的类别id, -1, 材料id_In, 批次_In, 产地_In, 批号_In, 效期_In,
     灭菌效期_In, 填写数量_In, 实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, n_上次供应商id, n_上次生产日期,
     v_批准文号, v_商品条码, v_内部条码);

  If To_Number(v_下库存, '9999') = 1 Then

    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) - 实际数量_In
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 性质 = 1;

    If Sql%RowCount = 0 Then
      Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;

      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
      Values
        (库房id_In, 材料id_In, Nvl(批次_In, 0), 1, -实际数量_In, 效期_In, 灭菌效期_In, n_上次供应商id, 成本价_In, 批号_In, n_上次生产日期, 产地_In,
         v_批准文号, Decode(n_实价卫材, 1, Decode(Nvl(批次_In, 0), 0, Null, 零售价_In), Null));
    End If;

    --同时更新库存数
    Delete From 药品库存
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  End If;

  --下面是判断入库的材料是否是分批核算材料
  Select Nvl(库房分批, 0), Nvl(在用分批, 0) Into n_库房分批, n_在用分批 From 材料特性 Where 材料id = 材料id_In;

  n_是否分批 := 0;

  If n_在用分批 = 0 Then
    If n_库房分批 = 1 Then
      Begin
        Select Distinct 0
        Into n_是否分批
        From 部门性质说明
        Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = 对方部门id_In;
      Exception
        When Others Then
          n_是否分批 := 1;
      End;
    End If;
  Else
    n_是否分批 := 1;
  End If;

  Select 药品收发记录_Id.Nextval Into n_Id From Dual;

  If n_是否分批 = 1 And Nvl(批次_In, 0) = 0 Then
    --入库分批且出库不分批
    n_批次 := n_Id;
  Elsif n_是否分批 = 0 Then
    --入库不分批
    n_批次 := 0;
  Elsif Nvl(批次_In, 0) <> 0 Then
    --入库分批且出库也分批
    n_批次 := 批次_In;
  End If;

  --插入类别为入的那一笔
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价,
     摘要, 填制人, 填制日期, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
  Values
    (n_Id, 1, 19, No_In, 序号_In + 1, 对方部门id_In, 库房id_In, n_入的类别id, 1, 材料id_In, n_批次, 产地_In, 批号_In, 效期_In, 灭菌效期_In,
     填写数量_In, 实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, n_上次供应商id, n_上次生产日期, v_批准文号,
     v_商品条码, v_内部条码);

  --检查是否存在相同材料相同批次的数据，如果存在不允许保存
  Select Count(*)
  Into n_Records
  From 药品收发记录
  Where 单据 = 19 And NO = No_In And 入出系数 = -1 And 药品id + 0 = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0);

  If n_Records > 1 Then
    Select 编码 Into v_编码 From 收费项目目录 Where ID = 材料id_In;
    v_Err_Msg := '[ZLSOFT]编码为' || v_编码 || '的药品，存在多条重复的记录，请合并为一条记录！[ZLSOFT]';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料移库_Insert;
/

--127591:刘涛,2018-07-24,过程取消检查库存数量
Create Or Replace Procedure Zl_材料外购_Strike
(
  行次_In       In Integer,
  原记录状态_In In 药品收发记录.记录状态%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  冲销数量_In   In 药品收发记录.实际数量%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  随货单号_In   In 应付记录.随货单号%Type := Null,
  发票号_In     In 应付记录.发票号%Type := Null,
  发票日期_In   In 应付记录.发票日期%Type := Null,
  发票金额_In   In 应付记录.发票金额%Type := Null,
  全部冲销_In   In 药品收发记录.实际数量%Type := 0, --用于财务审核
  财务审核_In   In Number := 0, --财务审核标志:1-财务审核,0-冲销
  摘要_In       In 药品收发记录.摘要%Type := Null,
  发票代码_In   In 应付记录.发票代码%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  v_产地     药品收发记录.产地%Type;
  v_批号     药品收发记录.批号%Type;
  v_摘要     药品收发记录.摘要%Type;
  v_核查人   药品收发记录.配药人%Type;
  v_注册证号 药品收发记录.注册证号%Type;
  d_效期     药品收发记录.效期%Type;
  d_灭菌效期 药品收发记录.灭菌效期%Type;
  d_灭菌日期 药品收发记录.灭菌日期%Type;
  d_生产日期 药品收发记录.生产日期%Type;
  d_核查日期 药品收发记录.配药日期%Type;
  v_商品条码 药品收发记录.商品条码%Type;
  v_内部条码 药品收发记录.内部条码%Type;

  n_应付id     应付记录.Id%Type;
  n_库房id     药品收发记录.库房id%Type;
  n_供药单位id 药品收发记录.供药单位id%Type;
  n_入出类别id 药品收发记录.入出类别id%Type;
  n_批次       药品收发记录.批次%Type;
  n_成本价     药品收发记录.成本价%Type;
  n_成本金额   药品收发记录.成本金额%Type;
  n_扣率       药品收发记录.扣率%Type;
  n_零售价     药品收发记录.零售价%Type;
  n_零售金额   药品收发记录.零售金额%Type;
  n_差价       药品收发记录.差价%Type;
  n_零售差价   药品收发记录.差价%Type;

  n_剩余数量     药品收发记录.实际数量%Type;
  n_剩余成本金额 药品收发记录.成本金额%Type;
  n_剩余零售金额 药品收发记录.零售金额%Type;
  n_剩余差价金额 药品收发记录.差价%Type;
  n_平均成本价   药品库存.平均成本价%Type;
  v_批准文号     药品收发记录.批准文号%Type;

  n_发药方式 药品收发记录.发药方式%Type;
  n_入出系数 药品收发记录.入出系数%Type;
  n_冲销数量 药品收发记录.实际数量%Type;
  --对冲销数量进行检查
  n_库存数   药品库存.实际数量%Type;
  n_收发id   药品收发记录.Id%Type;
  n_记录状态 药品收发记录.记录状态%Type;
  n_实价卫材 收费项目目录.是否变价%Type;

  n_库房分批   Integer;
  n_在用分批   Integer;
  n_分批属性   Integer;
  n_库房       Integer;
  n_分批       Number;
  n_小数       Number(2);
  n_发票金额   Number(16, 5);
  n_Batchcount Integer; --原不分批现在分批的材料的数量

Begin

  Select Nvl(精度, 2) Into n_小数 From 药品卫材精度 Where 性质 = 0 And 类别 = 2 And 内容 = 4 And 单位 = 5;

  If 行次_In = 1 Then
    Update 药品收发记录
    Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3), 费用id = Decode(财务审核_In, 1, 1, Null)
    Where NO = No_In And 单据 = 15 And 记录状态 = 原记录状态_In;
  
    If Sql%RowCount = 0 Then
      v_Err_Msg := '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
    
      Raise Err_Item;
    End If;
  End If;

  --主要针对原不分批现在分批的材料，不能对其审核
  Select Count(*)
  Into n_Batchcount
  From 药品收发记录 A, 材料特性 B
  Where a.药品id = b.材料id And a.No = No_In And a.单据 = 15 And Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And
        a.药品id + 0 = 材料id_In And
        ((Nvl(b.在用分批, 0) = 1 And
        a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '发料部门') Or (工作性质 Like '制剂室'))) Or Nvl(b.在用分批, 0) = 1);

  If n_Batchcount > 0 Then
    v_Err_Msg := '[ZLSOFT]该单据中第' || 序号_In || '行的材料原来不分批，现在分批的材料，不能冲销！[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Sum(a.实际数量) As 剩余数量, Sum(a.成本金额) As 剩余成本金额, Sum(a.零售金额) As 剩余零售金额,
         Sum(To_Number(To_Char(Nvl(用法, '0'), '999999999990.9999999'), '999999999990.9999999')) As 剩余差价金额, a.库房id,
         a.供药单位id, a.入出类别id, a.入出系数, Nvl(a.批次, 0), a.产地, a.批号, a.效期, a.灭菌效期, a.灭菌日期, a.生产日期, a.成本价, a.扣率, a.零售价, a.注册证号,
         b.库房分批, b.在用分批, Max(a.发药方式), Max(a.配药人), Max(配药日期), a.商品条码, a.内部条码, a.批准文号
  Into n_剩余数量, n_剩余成本金额, n_剩余零售金额, n_剩余差价金额, n_库房id, n_供药单位id, n_入出类别id, n_入出系数, n_批次, v_产地, v_批号, d_效期, d_灭菌效期, d_灭菌日期,
       d_生产日期, n_成本价, n_扣率, n_零售价, v_注册证号, n_库房分批, n_在用分批, n_发药方式, v_核查人, d_核查日期, v_商品条码, v_内部条码, v_批准文号
  From 药品收发记录 A, 材料特性 B
  Where a.No = No_In And a.药品id = b.材料id And a.单据 = 15 And a.药品id + 0 = 材料id_In And a.序号 = 序号_In
  Group By a.库房id, a.供药单位id, a.入出类别id, a.入出系数, Nvl(a.批次, 0), a.产地, a.批号, a.效期, a.灭菌效期, a.灭菌日期, a.生产日期, a.成本价, a.扣率,
           a.零售价, a.注册证号, b.库房分批, b.在用分批, a.商品条码, a.内部条码,a.批准文号;

  --判断该部门是库房还是发料部门
  Begin
    Select Distinct 0
    Into n_库房
    From 部门性质说明
    Where (工作性质 = '发料部门' Or 工作性质 = '制剂室') And 部门id = n_库房id;
  Exception
    When Others Then
      n_库房 := 1;
  End;

  --根据部门性质,判断分批特性
  If n_库房 = 0 Then
    n_分批属性 := n_在用分批;
  
  Else
    n_分批属性 := n_库房分批;
  End If;

  --n_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
  n_分批 := 0;
  If n_分批属性 = 1 And n_批次 <> 0 Then
    n_分批 := n_批次;
  End If;

  --全部冲销或者财务审核时，冲销数量等于剩余数量；其他情况冲销数量等于传入的冲销数量 
  If 全部冲销_In = 1 Or 财务审核_In = 1 Then
    n_冲销数量 := n_剩余数量;
  Else
    n_冲销数量 := 冲销数量_In;
  End If;

  Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;

  n_成本金额 := Round(n_冲销数量 / n_剩余数量 * n_剩余成本金额, n_小数);
  n_零售金额 := Round(n_冲销数量 / n_剩余数量 * n_剩余零售金额, n_小数);
  n_差价     := Round(n_零售金额 - n_成本金额, n_小数);
  If 全部冲销_In = 1 Or 财务审核_In = 1 Then
    n_零售差价 := n_剩余差价金额;
  Else
    n_零售差价 := Round(n_冲销数量 / n_剩余数量 * n_剩余差价金额, n_小数);
  End If;

  Select 药品收发记录_Id.Nextval Into n_收发id From Dual;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 灭菌效期, 灭菌日期, 填写数量, 实际数量, 成本价, 成本金额,
     扣率, 零售价, 零售金额, 差价, 摘要, 注册证号, 填制人, 填制日期, 审核人, 审核日期, 配药人, 配药日期, 发药方式, 用法, 费用id, 商品条码, 内部条码, 批准文号)
  Values
    (n_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 15, No_In, 序号_In, n_库房id, n_供药单位id, n_入出类别id, 1, 材料id_In, n_批次, v_产地,
     v_批号, d_生产日期, d_效期, d_灭菌效期, d_灭菌日期, -n_冲销数量, -n_冲销数量, n_成本价, -n_成本金额, n_扣率, n_零售价, -n_零售金额, -n_差价, 摘要_In, v_注册证号,
     填制人_In, 填制日期_In, 填制人_In, 填制日期_In, v_核查人, d_核查日期, n_发药方式, -n_零售差价, Decode(财务审核_In, 1, 1, Null), v_商品条码, v_内部条码,
     v_批准文号);

  --对于冲销的单据也应该对应付余额表进行处理
  --只对填了发票号的记录进行处理
  n_发票金额 := Nvl(发票金额_In, 0);
  If (Nvl(发票号_In, ' ') <> ' ' Or 随货单号_In Is Not Null) And Nvl(n_发票金额, 0) <> 0 Then
    --对于财务审核的，要将剩余的发票金额全部冲销
    If 全部冲销_In = 1 Then
    
      Select Sum(b.发票金额)
      Into n_发票金额
      From (Select ID From 药品收发记录 Where 单据 = 15 And NO = No_In And 序号 = 序号_In) A, 应付记录 B
      Where a.Id = b.收发id And b.系统标识 = 5 And b.记录性质 <> -1;
    
    End If;
  
    Update 应付余额 Set 金额 = Nvl(金额, 0) - Nvl(n_发票金额, 0) Where 单位id = n_供药单位id And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 应付余额 (单位id, 性质, 金额) Values (n_供药单位id, 1, -nvl(n_发票金额, 0));
    
    End If;
    Delete 应付余额 Where 单位id = n_供药单位id And Nvl(金额, 0) = 0;
  
  End If;

  Update 药品库存
  Set 可用数量 = Nvl(可用数量, 0) - n_冲销数量, 实际数量 = Nvl(实际数量, 0) - n_冲销数量, 实际金额 = Nvl(实际金额, 0) - n_零售金额,
      实际差价 = Nvl(实际差价, 0) - n_差价, 上次供应商id = n_供药单位id, 上次采购价 = n_成本价, 上次批号 = v_批号, 上次产地 = v_产地, 灭菌效期 = d_灭菌效期,
      上次生产日期 = d_生产日期, 效期 = d_效期, 批准文号 = v_批准文号
  Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(n_分批, 0) And 性质 = 1;

  If Sql%NotFound Then
    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 灭菌效期, 上次生产日期, 零售价, 平均成本价, 批准文号)
    Values
      (n_库房id, 材料id_In, n_分批, 1, -n_冲销数量, -n_冲销数量, -n_零售金额, -n_差价, n_供药单位id, n_成本价, v_批号, v_产地, d_效期, d_灭菌效期, d_生产日期,
       Decode(n_实价卫材, 0, Null, Decode(Nvl(n_分批, 0), 0, Null, n_零售价)), n_成本价, v_批准文号);
  End If;

  --清除数量金额为零的记录
  Delete From 药品库存
  Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  --重新计算库存表中的平均成本价
  Update 药品库存
  Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
  Where 性质 = 1 And 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(n_分批, 0) And Nvl(实际数量, 0) <> 0;
  If Sql%NotFound Then
    Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = 材料id_In;
    Update 药品库存
    Set 平均成本价 = n_平均成本价
    Where 性质 = 1 And 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(n_分批, 0) And Nvl(平均成本价, 0) <> n_成本价;
  End If;

  --产生应付记录的冲销记录(先判断应付记录中是否已存在该记录对应的冲销记录,是则更新;否则新增)
  Select 应付记录_Id.Nextval Into n_应付id From Dual;

  Begin
    Select Max(记录状态) + 3
    Into n_记录状态
    From 应付记录
    Where (系统标识, 记录性质, NO, 项目id, 序号) In
          (Select 系统标识, 记录性质, NO, 项目id, 序号
           From 应付记录
           Where 收发id =
                 (Select ID From 药品收发记录 Where NO = No_In And 单据 = 15 And 序号 = 序号_In And Mod(记录状态, 3) = 0) And 系统标识 = 5 And
                 记录性质 = 0) And 记录状态 <> 1 And Mod(记录状态, 3) <> 0;
  
  Exception
    When Others Then
      n_记录状态 := 2;
  End;
  If n_记录状态 Is Null Then
    n_记录状态 := 2;
  End If;
  If Mod(n_记录状态, 3) <> 2 Then
    n_记录状态 := n_记录状态 + 1;
  End If;
  If Mod(n_记录状态, 3) <> 2 Then
    n_记录状态 := n_记录状态 + 1;
  End If;

  Insert Into 应付记录
    (ID, 记录性质, 记录状态, 项目id, 序号, 单位id, NO, 系统标识, 收发id, 入库单据号, 单据金额, 随货单号, 发票号, 发票日期, 发票金额, 品名, 规格, 产地, 批号, 计量单位, 数量, 采购价,
     采购金额, 填制人, 填制日期, 审核人, 审核日期, 摘要, 库房id, 发票代码)
    Select n_应付id, 记录性质, n_记录状态, 材料id_In, 序号_In, 单位id, NO, 5, n_收发id, 入库单据号, -1 * n_零售金额, 随货单号, 发票号, 发票日期, -n_发票金额, 品名,
           规格, 产地, 批号, 计量单位, -1 * n_冲销数量, 采购价, -1 * 采购价 * n_冲销数量, 填制人_In, 填制日期_In, 填制人_In, 填制日期_In, 摘要, 库房id, 发票代码_In
    From 应付记录
    Where 收发id = (Select ID From 药品收发记录 Where NO = No_In And 单据 = 15 And 序号 = 序号_In And Mod(记录状态, 3) = 0) And 系统标识 = 5 And
          记录性质 = 0;

  Update 应付记录
  Set 记录状态 = 3
  Where 收发id = (Select ID From 药品收发记录 Where NO = No_In And 单据 = 15 And 序号 = 序号_In And Mod(记录状态, 3) = 0) And 系统标识 = 5 And
        记录性质 = 0;

  --处理调价后冲销（财务审核时不需要处理）        
  If 财务审核_In = 0 Then
    Zl_材料收发记录_调价修正(n_收发id);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料外购_Strike;
/

--128171:刘涛,2018-09-29,处理申请冲销减可用数量处理
--129697:刘涛,2018-08-02,取消过程中冲销数量与库存数量的比较
--127591:刘涛,2018-07-24,过程取消检查库存数量
Create Or Replace Procedure Zl_材料移库_Strike
(
  行次_In       In Integer,
  原记录状态_In In 药品收发记录.记录状态%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  冲销数量_In   In 药品收发记录.实际数量%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  冲销方式_In   In Integer := 0
  --0－正常冲销方式；1－产生冲销申请单据；2－审核已产生的冲销申请单据
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_Batch_Count Integer; --原不分批现在分批的药品的数量

  n_库房id       药品收发记录.库房id%Type;
  n_对方部门id   药品收发记录.对方部门id%Type;
  n_批次         药品收发记录.批次%Type;
  n_成本价       药品收发记录.成本价%Type;
  n_成本金额     药品收发记录.成本金额%Type;
  n_零售价       药品收发记录.零售价%Type;
  n_零售金额     药品收发记录.零售金额%Type;
  n_差价率       药品收发记录.差价%Type;
  v_产地         药品收发记录.产地%Type;
  v_批号         药品收发记录.批号%Type;
  v_效期         药品收发记录.效期%Type;
  v_商品条码     药品收发记录.商品条码%Type;
  v_内部条码     药品收发记录.内部条码%Type;
  v_灭菌日期     药品收发记录.灭菌日期%Type;
  v_灭菌效期     药品收发记录.灭菌效期%Type;
  n_供药单位id   药品收发记录.供药单位id%Type;
  d_生产日期     药品收发记录.生产日期%Type;
  v_批准文号     药品收发记录.批准文号%Type;
  n_扣率         药品收发记录.扣率%Type;
  n_序号         药品收发记录.序号%Type;
  n_入出系数     药品收发记录.入出系数%Type;
  n_入出类别id   药品收发记录.入出类别id%Type;
  v_配药人       药品收发记录.配药人%Type;
  d_发送日期     药品收发记录.配药日期%Type;
  v_摘要         药品收发记录.摘要%Type;
  n_剩余数量     药品收发记录.实际数量%Type;
  n_剩余成本金额 药品收发记录.成本金额%Type;
  n_剩余零售金额 药品收发记录.零售金额%Type;
  n_收发id       药品收发记录.Id%Type;
  n_实价卫材     收费项目目录.是否变价%Type;
  --对冲销数量进行检查
  n_库房分批   Integer;
  n_在用分批   Integer;
  n_小数       Number;
  n_记录数     Number;
  n_平均成本价 药品库存.平均成本价%Type;
  v_下可用库存 Zlparameters.参数值%Type;
  n_可用数量   药品库存.可用数量%Type;

  Cursor c_药品收发记录 Is
    Select a.Id, a.序号, a.库房id, a.对方部门id, a.入出类别id, a.入出系数, a.药品id, Nvl(a.批次, 0) As 批次, a.产地, a.批号, a.效期, a.配药人,
           a.配药日期 As 发送日期, a.摘要, a.供药单位id, a.批准文号, a.生产日期, a.成本价, a.零售价, Nvl(b.是否变价, 0) As 时价, a.扣率, a.单量, a.频次, a.商品条码,
           a.内部条码, a.灭菌日期, a.灭菌效期
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.No = No_In And a.单据 = 19 And (a.序号 >= 序号_In And a.序号 <= 序号_In + 1) And
          (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id, a.批次, a.序号;

  Cursor c_冲销申请记录 Is
    Select a.Id, a.序号, a.库房id, a.对方部门id, a.入出类别id, a.入出系数, a.药品id, Nvl(a.批次, 0) As 批次, a.产地, a.批号, a.效期, a.配药人,
           a.配药日期 As 发送日期, a.摘要, a.供药单位id, a.批准文号, a.生产日期, a.成本价, a.实际数量, a.零售金额, a.差价, a.零售价, Nvl(b.是否变价, 0) As 时价,
           a.扣率, a.单量, a.频次, a.商品条码, a.内部条码, a.灭菌日期, a.灭菌效期
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.No = No_In And a.单据 = 19 And (a.序号 >= 序号_In And a.序号 <= 序号_In + 1) And
          (a.记录状态 = 原记录状态_In And Mod(a.记录状态, 3) = 2) And a.审核日期 Is Null
    Order By a.药品id, a.批次;
Begin
  --获取金额小数位数
  Select Nvl(精度, 2) Into n_小数 From 药品卫材精度 Where 性质 = 0 And 类别 = 2 And 内容 = 4 And 单位 = 5;
  Select Nvl(zl_GetSysParameter(95), '0') Into v_下可用库存 From Dual;

  If 冲销方式_In = 1 Then
    --申请冲销，只产生冲销数据，不填写审核人，审核日期，也不更新库存
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where NO = No_In And 单据 = 19 And 记录状态 = 原记录状态_In;
    
      If Sql%RowCount = 0 Then
        v_Err_Msg := '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
        Raise Err_Item;
      End If;
    End If;
  
    --原来不分批，现在分批的卫生材料，不能冲销
    Select Count(*)
    Into n_Batch_Count
    From 药品收发记录 A, 材料特性 B
    Where a.药品id = b.材料id And a.No = No_In And a.单据 = 19 And a.药品id + 0 = 材料id_In And Mod(a.记录状态, 3) = 0 And
          Nvl(a.批次, 0) = 0 And
          ((Nvl(b.库房分批, 0) = 1 And
          a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%发料部门') Or (工作性质 Like '制剂室'))) Or
          Nvl(b.在用分批, 0) = 1);
  
    If n_Batch_Count > 0 Then
      v_Err_Msg := '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的卫生材料，不能冲销！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --获取当前冲销单据剩余数量
    Select Sum(a.实际数量) As 剩余数量, Sum(a.成本金额) As 剩余成本金额, Sum(a.零售金额) As 剩余零售金额, a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0),
           b.库房分批, b.在用分批
    Into n_剩余数量, n_剩余成本金额, n_剩余零售金额, n_成本价, n_零售价, n_库房id, n_批次, n_库房分批, n_在用分批
    From 药品收发记录 A, 材料特性 B
    Where a.No = No_In And a.药品id = b.材料id And a.单据 = 19 And a.药品id + 0 = 材料id_In And a.序号 = 序号_In
    Group By a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0), b.库房分批, b.在用分批;
    --判断该部门是库房还是发料部门
    --n_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    --因为对于对方库房，相当于出库，而对于当前库房，相当于入库，所以当前库房不予检查，仅检查退库那条记录
    Select Nvl(a.批次, 0)
    Into n_批次
    From 药品收发记录 A
    Where a.No = No_In And a.单据 = 19 And a.药品id + 0 = 材料id_In And a.序号 = 序号_In + 1 And Mod(a.记录状态, 3) = 0;
  
    If Nvl(n_剩余数量, 0) = 0 Then
      v_Err_Msg := '[ZLSOFT]该单据中第' || Ceil(序号_In / 2) || '行的材料已经被冲销完成,不能再冲！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    n_成本金额 := Round(冲销数量_In / n_剩余数量 * n_剩余成本金额, n_小数);
    n_零售金额 := Round(冲销数量_In / n_剩余数量 * n_剩余零售金额, n_小数);
    n_差价率   := Round(n_零售金额 - n_成本金额, n_小数);
  
    For v_药品收发记录 In c_药品收发记录 Loop
      n_库房id     := v_药品收发记录.库房id;
      n_对方部门id := v_药品收发记录.对方部门id;
      n_批次       := v_药品收发记录.批次;
      n_零售价     := v_药品收发记录.零售价;
      n_入出系数   := v_药品收发记录.入出系数;
      n_成本价     := v_药品收发记录.成本价;
      v_产地       := v_药品收发记录.产地;
      v_批号       := v_药品收发记录.批号;
      v_效期       := v_药品收发记录.效期;
      v_商品条码   := v_药品收发记录.商品条码;
      v_内部条码   := v_药品收发记录.内部条码;
      v_灭菌效期   := v_药品收发记录.灭菌效期;
      v_灭菌日期   := v_药品收发记录.灭菌日期;
      n_供药单位id := v_药品收发记录.供药单位id;
      d_生产日期   := v_药品收发记录.生产日期;
      v_批准文号   := v_药品收发记录.批准文号;
      n_扣率       := v_药品收发记录.扣率;
      n_序号       := v_药品收发记录.序号;
      n_入出类别id := v_药品收发记录.入出类别id;
      v_配药人     := v_药品收发记录.配药人;
      d_发送日期   := v_药品收发记录.发送日期;
      v_摘要       := v_药品收发记录.摘要;
    
      Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
      Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;
    
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌日期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价,
         零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 配药日期, 供药单位id, 生产日期, 批准文号, 扣率, 商品条码, 内部条码)
      Values
        (n_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 19, No_In, n_序号, n_库房id, n_对方部门id, n_入出类别id, n_入出系数, 材料id_In,
         n_批次, v_产地, v_批号, v_效期, v_灭菌日期, v_灭菌效期, -冲销数量_In, -冲销数量_In, n_成本价, -n_成本金额, n_零售价, -n_零售金额, -n_差价率, v_摘要,
         填制人_In, 填制日期_In, v_配药人, d_发送日期, n_供药单位id, d_生产日期, v_批准文号, n_扣率, v_商品条码, v_内部条码);
    
      --参数为1表示申请冲销时下可用数量，仅对原移入库房
      If v_下可用库存 = '1' And n_入出系数 = 1 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - 冲销数量_In
        Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = n_批次 And 性质 = 1;
      
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 上次批号, 效期, 上次产地)
          Values
            (n_库房id, 材料id_In, n_批次, 1, -1 * 冲销数量_In, v_批号, v_效期, v_产地);
        End If;
      
        Delete From 药品库存
        Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
              Nvl(实际差价, 0) = 0;
      End If;
    End Loop;
  
  Elsif 冲销方式_In = 2 Then
  
    --审核申请冲销产生的单据，及填写审核人，审核日期，并更新库存
    For v_冲销申请记录 In c_冲销申请记录 Loop
      n_库房id     := v_冲销申请记录.库房id;
      n_对方部门id := v_冲销申请记录.对方部门id;
      n_批次       := v_冲销申请记录.批次;
      n_零售价     := v_冲销申请记录.零售价;
      n_零售金额   := v_冲销申请记录.零售金额;
      n_差价率     := v_冲销申请记录.差价;
      n_入出系数   := v_冲销申请记录.入出系数;
      n_成本价     := v_冲销申请记录.成本价;
      v_产地       := v_冲销申请记录.产地;
      v_批号       := v_冲销申请记录.批号;
      v_效期       := v_冲销申请记录.效期;
      v_商品条码   := v_冲销申请记录.商品条码;
      v_内部条码   := v_冲销申请记录.内部条码;
      v_灭菌效期   := v_冲销申请记录.灭菌效期;
      v_灭菌日期   := v_冲销申请记录.灭菌日期;
      n_供药单位id := v_冲销申请记录.供药单位id;
      d_生产日期   := v_冲销申请记录.生产日期;
      v_批准文号   := v_冲销申请记录.批准文号;
      n_扣率       := v_冲销申请记录.扣率;
      --原分批现不分批的材料,在冲消时，要处理他
      Begin
        Select Count(*)
        Into n_记录数
        From 药品收发记录 A, 材料特性 B
        Where b.材料id = a.药品id And a.药品id = 材料id_In And a.No = No_In And a.单据 = 19 And a.库房id = n_库房id And
              Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) > 0 And
              (Nvl(b.库房分批, 0) = 0 Or
              (Nvl(b.在用分批, 0) = 0 And
              a.库房id In (Select 部门id From 部门性质说明 Where (工作性质 Like '%发料部门') Or (工作性质 Like '制剂室'))));
      Exception
        When Others Then
          n_记录数 := 0;
      End;
      If n_记录数 > 0 Then
        n_批次 := 0;
      Else
        n_批次 := Nvl(n_批次, 0);
      End If;
      Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;
      --申请时，已经下了可用库存这个地方就不能在下了
      If v_下可用库存 = '1' And n_入出系数 = 1 Then
        n_可用数量 := 0;
      Else
        n_可用数量 := 冲销数量_In;
      End If;
    
      --更改药品库存表的相应数据
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + Nvl(n_可用数量, 0) * n_入出系数, 实际数量 = Nvl(实际数量, 0) + Nvl(冲销数量_In, 0) * n_入出系数,
          实际金额 = Nvl(实际金额, 0) + Nvl(n_零售金额, 0) * n_入出系数, 实际差价 = Nvl(实际差价, 0) + Nvl(n_差价率, 0) * n_入出系数,
          上次采购价 = Nvl(n_成本价, 上次采购价), 上次批号 = Nvl(v_批号, 上次批号), 上次产地 = Nvl(v_产地, 上次产地), 效期 = Nvl(v_效期, 效期),
          零售价 = Decode(n_实价卫材, 1, Decode(Nvl(批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, n_零售价, 零售价)), Null),
          商品条码 = Nvl(商品条码, v_商品条码), 内部条码 = Nvl(内部条码, v_内部条码)
      Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = n_批次 And 性质 = 1;
    
      If Sql%NotFound Then
        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次采购价, 上次批号, 上次产地, 效期, 灭菌效期, 上次供应商id, 上次生产日期, 批准文号, 零售价, 上次扣率,
           商品条码, 内部条码, 平均成本价)
        Values
          (n_库房id, 材料id_In, n_批次, 1, 冲销数量_In * n_入出系数, 冲销数量_In * n_入出系数, n_零售金额 * n_入出系数, n_差价率 * n_入出系数, n_成本价, v_批号,
           v_产地, v_效期, v_灭菌效期, n_供药单位id, d_生产日期, v_批准文号, Decode(n_实价卫材, 1, Decode(Nvl(n_批次, 0), 0, Null, n_零售价), Null),
           n_扣率, v_商品条码, v_内部条码, n_成本价);
      End If;
    
      Delete From 药品库存
      Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
            Nvl(实际差价, 0) = 0;
    
      --填写审核人、审核日期
      Update 药品收发记录
      Set 审核人 = 填制人_In, 审核日期 = 填制日期_In
      Where NO = No_In And 单据 = 19 And ID = v_冲销申请记录.Id;
    
      Zl_材料收发记录_调价修正(v_冲销申请记录.Id);
    End Loop;
  Else
    --正常冲销业务，产生冲销单据审核并更新库存
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where NO = No_In And 单据 = 19 And 记录状态 = 原记录状态_In;
      If Sql%RowCount = 0 Then
        v_Err_Msg := '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
        Raise Err_Item;
      End If;
    End If;
  
    Select Count(*)
    Into n_Batch_Count
    From 药品收发记录 A, 材料特性 B
    Where a.药品id = b.材料id And a.No = No_In And a.单据 = 19 And a.药品id + 0 = 材料id_In And Mod(a.记录状态, 3) = 0 And
          Nvl(a.批次, 0) = 0 And
          ((Nvl(b.库房分批, 0) = 1 And
          a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%发料部门') Or (工作性质 Like '制剂室'))) Or
          Nvl(b.在用分批, 0) = 1);
  
    If n_Batch_Count > 0 Then
      v_Err_Msg := '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的卫生材料，不能冲销！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Sum(a.实际数量) As 剩余数量, Sum(a.成本金额) As 剩余成本金额, Sum(a.零售金额) As 剩余零售金额, a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0),
           b.库房分批, b.在用分批
    Into n_剩余数量, n_剩余成本金额, n_剩余零售金额, n_成本价, n_零售价, n_库房id, n_批次, n_库房分批, n_在用分批
    From 药品收发记录 A, 材料特性 B
    Where a.No = No_In And a.药品id = b.材料id And a.单据 = 19 And a.药品id + 0 = 材料id_In And a.序号 = 序号_In
    Group By a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0), b.库房分批, b.在用分批;
  
    --判断该部门是库房还是发料部门
    --n_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    --因为对于对方库房，相当于出库，而对于当前库房，相当于入库，所以当前库房不予检查，仅检查退库那条记录
    Select Nvl(a.批次, 0)
    Into n_批次
    From 药品收发记录 A
    Where a.No = No_In And a.单据 = 19 And a.药品id + 0 = 材料id_In And a.序号 = 序号_In + 1 And Mod(a.记录状态, 3) = 0;
  
    If Nvl(n_剩余数量, 0) = 0 Then
      v_Err_Msg := '[ZLSOFT]该单据中第' || Ceil(序号_In / 2) || '行的材料已经被冲销完成,不能再冲！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    n_成本金额 := Round(冲销数量_In / n_剩余数量 * n_剩余成本金额, n_小数);
    n_零售金额 := Round(冲销数量_In / n_剩余数量 * n_剩余零售金额, n_小数);
    n_差价率   := Round(n_零售金额 - n_成本金额, n_小数);
  
    For v_药品收发记录 In c_药品收发记录 Loop
      n_库房id     := v_药品收发记录.库房id;
      n_对方部门id := v_药品收发记录.对方部门id;
      n_批次       := v_药品收发记录.批次;
      n_零售价     := v_药品收发记录.零售价;
      n_入出系数   := v_药品收发记录.入出系数;
      n_成本价     := v_药品收发记录.成本价;
      v_产地       := v_药品收发记录.产地;
      v_批号       := v_药品收发记录.批号;
      v_效期       := v_药品收发记录.效期;
      v_商品条码   := v_药品收发记录.商品条码;
      v_内部条码   := v_药品收发记录.内部条码;
      v_灭菌效期   := v_药品收发记录.灭菌效期;
      v_灭菌日期   := v_药品收发记录.灭菌日期;
      n_供药单位id := v_药品收发记录.供药单位id;
      d_生产日期   := v_药品收发记录.生产日期;
      v_批准文号   := v_药品收发记录.批准文号;
      n_扣率       := v_药品收发记录.扣率;
      n_序号       := v_药品收发记录.序号;
      n_入出类别id := v_药品收发记录.入出类别id;
      v_配药人     := v_药品收发记录.配药人;
      d_发送日期   := v_药品收发记录.发送日期;
      v_摘要       := v_药品收发记录.摘要;
    
      Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
      Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;
    
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌日期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价,
         零售金额, 差价, 摘要, 填制人, 填制日期, 审核人, 审核日期, 配药人, 配药日期, 供药单位id, 生产日期, 批准文号, 扣率, 商品条码, 内部条码)
      Values
        (n_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 19, No_In, n_序号, n_库房id, n_对方部门id, n_入出类别id, n_入出系数, 材料id_In,
         n_批次, v_产地, v_批号, v_效期, v_灭菌日期, v_灭菌效期, -冲销数量_In, -冲销数量_In, n_成本价, -n_成本金额, n_零售价, -n_零售金额, -n_差价率, v_摘要,
         填制人_In, 填制日期_In, 填制人_In, 填制日期_In, v_配药人, d_发送日期, n_供药单位id, d_生产日期, v_批准文号, n_扣率, v_商品条码, v_内部条码);
    
      --原分批现不分批的材料,在冲消时，要处理他
      Begin
        Select Count(*)
        Into n_记录数
        From 药品收发记录 A, 材料特性 B
        Where b.材料id = a.药品id And a.药品id = 材料id_In And a.No = No_In And a.单据 = 19 And a.库房id = n_库房id And
              Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) > 0 And
              (Nvl(b.库房分批, 0) = 0 Or
              (Nvl(b.在用分批, 0) = 0 And
              a.库房id In (Select 部门id From 部门性质说明 Where (工作性质 Like '%发料部门') Or (工作性质 Like '制剂室'))));
      Exception
        When Others Then
          n_记录数 := 0;
      End;
      If n_记录数 > 0 Then
        n_批次 := 0;
      Else
        n_批次 := Nvl(n_批次, 0);
      End If;
    
      --更改药品库存表的相应数据
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) - Nvl(冲销数量_In, 0) * n_入出系数, 实际数量 = Nvl(实际数量, 0) - Nvl(冲销数量_In, 0) * n_入出系数,
          实际金额 = Nvl(实际金额, 0) - Nvl(n_零售金额, 0) * n_入出系数, 实际差价 = Nvl(实际差价, 0) - Nvl(n_差价率, 0) * n_入出系数,
          上次采购价 = Nvl(n_成本价, 上次采购价), 上次批号 = Nvl(v_批号, 上次批号), 上次产地 = Nvl(v_产地, 上次产地), 效期 = Nvl(v_效期, 效期),
          零售价 = Decode(n_实价卫材, 1, Decode(Nvl(批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, n_零售价, 零售价)), Null),
          商品条码 = Nvl(商品条码, v_商品条码), 内部条码 = Nvl(内部条码, v_内部条码)
      Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = n_批次 And 性质 = 1;
    
      If Sql%NotFound Then
        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次采购价, 上次批号, 上次产地, 效期, 灭菌效期, 上次供应商id, 上次生产日期, 批准文号, 零售价, 上次扣率,
           商品条码, 内部条码, 平均成本价)
        Values
          (n_库房id, 材料id_In, n_批次, 1, -冲销数量_In * n_入出系数, -冲销数量_In * n_入出系数, -n_零售金额 * n_入出系数, -n_差价率 * n_入出系数, n_成本价,
           v_批号, v_产地, v_效期, v_灭菌效期, n_供药单位id, d_生产日期, v_批准文号,
           Decode(n_实价卫材, 1, Decode(Nvl(n_批次, 0), 0, Null, n_零售价), Null), n_扣率, v_商品条码, v_内部条码, n_成本价);
      End If;
    
      Delete From 药品库存
      Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
            Nvl(实际差价, 0) = 0;
    
      --重新计算库存表中的平均成本价
      Update 药品库存
      Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
      Where 性质 = 1 And 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = n_批次 And Nvl(实际数量, 0) <> 0;
      If Sql%NotFound Then
        Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = 材料id_In;
        Update 药品库存
        Set 平均成本价 = n_平均成本价
        Where 性质 = 1 And 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = n_批次 And Nvl(平均成本价, 0) <> n_成本价;
      End If;
      --处理调价后冲销
      Zl_材料收发记录_调价修正(n_收发id);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料移库_Strike;
/

--129697:刘涛,2018-08-02,取消过程中冲销数量与库存数量的比较
--127591:刘涛,2018-07-24,过程取消检查库存数量
Create Or Replace Procedure Zl_材料其他入库_Strike
(
  行次_In       In Integer,
  原记录状态_In In 药品收发记录.记录状态%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  冲销数量_In   In 药品收发记录.实际数量%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  全部冲销_In   In 药品收发记录.实际数量%Type := 0, --1-全部冲销,0-部分冲销,
  摘要_In       In 药品收发记录.摘要%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  v_产地 药品收发记录.产地%Type;
  v_批号 药品收发记录.批号%Type;

  n_库房id     药品收发记录.库房id%Type;
  n_入出类别id 药品收发记录.入出类别id%Type;
  n_批次       药品收发记录.批次%Type;
  n_成本价     药品收发记录.成本价%Type;
  n_成本金额   药品收发记录.成本金额%Type;
  n_扣率       药品收发记录.扣率%Type;
  n_零售价     药品收发记录.零售价%Type;
  n_零售金额   药品收发记录.零售金额%Type;
  n_差价       药品收发记录.差价%Type;
  n_零售差价   药品收发记录.差价%Type;
  v_商品条码   药品收发记录.商品条码%Type;
  v_内部条码   药品收发记录.内部条码%Type;
  v_批准文号   药品收发记录.批准文号%Type;

  n_入出系数     药品收发记录.入出系数%Type;
  n_收发id       药品收发记录.Id%Type;
  n_剩余数量     药品收发记录.实际数量%Type;
  n_剩余成本金额 药品收发记录.成本金额%Type;
  n_剩余零售金额 药品收发记录.零售金额%Type;
  n_剩余差价金额 药品收发记录.差价%Type;

  n_冲销数量 药品收发记录.实际数量%Type;
  n_实价卫材 收费项目目录.是否变价%Type;

  n_Batchcount Integer; --原不分批现在分批的材料的数量
  --对冲销数量进行检查
  n_库房分批 Integer;
  n_在用分批 Integer;
  n_分批属性 Integer;
  n_库房     Integer;
  n_分批     Number;
  n_小数     Number(2);

  d_生产日期   药品收发记录.生产日期%Type;
  d_效期       药品收发记录.效期%Type;
  d_灭菌效期   药品收发记录.灭菌效期%Type;
  d_灭菌日期   药品收发记录.灭菌日期%Type;
  n_平均成本价 药品库存.平均成本价%Type;
Begin
  n_冲销数量 := 冲销数量_In;
  Select Nvl(精度, 2) Into n_小数 From 药品卫材精度 Where 性质 = 0 And 类别 = 2 And 内容 = 4 And 单位 = 5;

  If 行次_In = 1 Then
    Update 药品收发记录
    Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
    Where NO = No_In And 单据 = 17 And 记录状态 = 原记录状态_In;
  
    If Sql%RowCount = 0 Then
      v_Err_Msg := '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;

  --主要针对原不分批现在分批的材料，不能对其审核
  Select Count(*)
  Into n_Batchcount
  From 药品收发记录 A, 材料特性 B
  Where a.药品id = b.材料id And a.No = No_In And a.单据 = 17 And Mod(a.记录状态, 3) = 0 And a.药品id + 0 = 材料id_In And
        Nvl(a.批次, 0) = 0 And
        ((Nvl(b.库房分批, 0) = 1 And
        a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '发料部门') Or (工作性质 Like '制剂室'))) Or Nvl(b.在用分批, 0) = 1);

  If n_Batchcount > 0 Then
    v_Err_Msg := '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能冲销！[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Sum(a.实际数量) As 剩余数量, Sum(a.成本金额) As 剩余成本金额, Sum(a.零售金额) As 剩余零售金额,
         Sum(To_Number(To_Char(Nvl(用法, '0'), '999999999990.9999999'), '999999999990.9999999')) As 剩余差价金额, a.库房id,
         a.入出类别id, a.入出系数, Nvl(a.批次, 0), a.产地, a.批号, a.生产日期, a.效期, a.灭菌效期, a.灭菌日期, a.成本价, a.扣率, a.零售价, b.库房分批, b.在用分批,
         a.商品条码, a.内部条码, a.批准文号
  Into n_剩余数量, n_剩余成本金额, n_剩余零售金额, n_剩余差价金额, n_库房id, n_入出类别id, n_入出系数, n_批次, v_产地, v_批号, d_生产日期, d_效期, d_灭菌效期, d_灭菌日期,
       n_成本价, n_扣率, n_零售价, n_库房分批, n_在用分批, v_商品条码, v_内部条码, v_批准文号
  From 药品收发记录 A, 材料特性 B
  Where a.No = No_In And a.单据 = 17 And a.药品id = b.材料id And a.药品id + 0 = 材料id_In And a.序号 = 序号_In
  Group By a.库房id, a.入出类别id, a.入出系数, Nvl(a.批次, 0), a.产地, a.批准文号, a.批号, a.生产日期, a.效期, a.灭菌效期, a.灭菌日期, a.成本价, a.扣率, a.零售价,
           b.库房分批, b.在用分批, a.商品条码, a.内部条码;

  --判断该部门是库房还是发料部门
  Begin
    Select Distinct 0
    Into n_库房
    From 部门性质说明
    Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = n_库房id;
  Exception
    When Others Then
      n_库房 := 1;
  End;

  --根据部门性质,判断分批特性
  If n_库房 = 0 Then
    n_分批属性 := n_在用分批;
  Else
    n_分批属性 := n_库房分批;
  End If;

  --n_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
  n_分批 := 0;
  If n_分批属性 = 1 And n_批次 <> 0 Then
    n_分批 := n_批次;
  End If;

  If Nvl(n_剩余数量, 0) = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据中第' || 序号_In || '行的材料已经被冲销完成,不能再冲！[ZLSOFT]';
    Raise Err_Item;
  End If;

  If 全部冲销_In = 1 Then
    n_冲销数量 := n_剩余数量;
  End If;

  Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;

  n_成本金额 := Round(n_冲销数量 / n_剩余数量 * n_剩余成本金额, n_小数);
  n_零售金额 := Round(n_冲销数量 / n_剩余数量 * n_剩余零售金额, n_小数);
  n_差价     := Round(n_零售金额 - n_成本金额, n_小数);
  n_零售差价 := Round(n_冲销数量 / n_剩余数量 * n_剩余差价金额, n_小数);

  Select 药品收发记录_Id.Nextval Into n_收发id From Dual;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 灭菌日期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额,
     差价, 用法, 摘要, 填制人, 填制日期, 审核人, 审核日期, 商品条码, 内部条码, 批准文号)
  Values
    (n_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 17, No_In, 序号_In, n_库房id, n_入出类别id, n_入出系数, 材料id_In, n_批次, v_产地,
     v_批号, d_生产日期, d_效期, d_灭菌日期, d_灭菌效期, -n_冲销数量, -n_冲销数量, n_成本价, -n_成本金额, n_零售价, -n_零售金额, -n_差价, -n_零售差价, 摘要_In, 填制人_In,
     填制日期_In, 填制人_In, 填制日期_In, v_商品条码, v_内部条码, v_批准文号);

  --更改药品库存表的相应数据

  Update 药品库存
  Set 可用数量 = Nvl(可用数量, 0) - Nvl(n_冲销数量, 0), 实际数量 = Nvl(实际数量, 0) - Nvl(n_冲销数量, 0), 实际金额 = Nvl(实际金额, 0) - Nvl(n_零售金额, 0),
      实际差价 = Nvl(实际差价, 0) - Nvl(n_差价, 0), 上次采购价 = Nvl(n_成本价, 上次采购价), 上次批号 = Nvl(v_批号, 上次批号), 上次产地 = Nvl(v_产地, 上次产地),
      上次生产日期 = Nvl(d_生产日期, 上次生产日期), 灭菌效期 = Nvl(d_灭菌效期, 灭菌效期), 效期 = Nvl(d_效期, 效期), 批准文号 = v_批准文号
  Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(n_分批, 0) And 性质 = 1;

  If Sql%NotFound Then
    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次采购价, 上次批号, 上次生产日期, 上次产地, 效期, 灭菌效期, 零售价, 平均成本价, 批准文号)
    Values
      (n_库房id, 材料id_In, n_分批, 1, -n_冲销数量, -n_冲销数量, -n_零售金额, -n_差价, n_成本价, v_批号, d_生产日期, v_产地, d_效期, d_灭菌效期,
       Decode(n_实价卫材, 0, Null, Decode(Nvl(n_分批, 0), 0, Null, n_零售价)), n_成本价, v_批准文号);
  End If;

  Delete From 药品库存
  Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  Zl_材料收发记录_调价修正(n_收发id);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料其他入库_Strike;
/

--128682:殷瑞,2018-07-16,修正不分批药品在退药时批号和效期为空的情况
Create Or Replace Procedure Zl_药品收发记录_部门退药
(
  Billid_In     In 药品收发记录.Id%Type,
  People_In     In 药品收发记录.审核人%Type,
  Date_In       In 药品收发记录.审核日期%Type,
  批号_In       In 药品库存.上次批号%Type := Null,
  效期_In       In 药品库存.效期%Type := Null,
  产地_In       In 药品库存.上次产地%Type := Null,
  退药数量_In   In 药品收发记录.实际数量%Type := Null,
  退药库房_In   In 药品收发记录.库房id%Type := Null,
  退药人_In     In 药品收发记录.领用人%Type := Null,
  Intdigit_In   In Number := 2,
  门诊_In       In Number := 2,
  汇总发药号_In In 药品收发记录.汇总发药号%Type := Null
) Is
  --只读变量
  Int记录状态   药品收发记录.记录状态%Type;
  Int执行状态   住院费用记录.执行状态%Type;
  Bln部分退药   Number;
  Lng入出类别id Number(18);
  Strno         药品收发记录.No%Type;
  Int单据       药品收发记录.单据%Type;
  Lng库房id     药品收发记录.库房id%Type;
  Lng药品id     药品收发记录.药品id%Type;
  Dbl实际数量   药品收发记录.实际数量%Type;
  Dbl实际金额   药品收发记录.零售金额%Type;
  Dbl实际成本   药品收发记录.成本金额%Type;
  Dbl实际差价   药品收发记录.差价%Type;
  Lng费用id     药品收发记录.费用id%Type;
  n_零售价      药品收发记录.零售价%Type;
  n_是否变价    Number;
  n_时价分批    Number;

  --20020731 Modified by zyb
  --处理退药时，分批核算性质改变后的处理
  Lng新批次 药品收发记录.批次%Type;
  Lng分批   药品规格.药房分批%Type;
  Lng批次   药品收发记录.批次%Type; --原批次

  Str批号        药品收发记录.批号%Type; --原批号
  Date效期       药品收发记录.效期%Type; --原效期
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_上次采购价   药品库存.上次采购价%Type;
  v_上次产地     药品库存.上次产地%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;
  v_批准文号     药品库存.批准文号%Type;

  n_记录性质   住院费用记录.记录性质%Type;
  v_收费类别   住院费用记录.收费类别%Type;
  n_付数       药品收发记录.付数%Type;
  n_原始数量   药品收发记录.实际数量%Type;
  v_冲销记录id 药品收发记录.Id%Type;
  Err_Custom Exception;
  v_Error    Varchar2(255);
  v_配药确认 药房配药控制.配药确认%Type;
  v_配药     药房配药控制.配药%Type;
  v_排队状态 Number(1);
  v_执行时间 药品收发记录.审核日期%Type;
Begin
  If 退药数量_In Is Not Null Then
    If 退药数量_In = 0 Then
      Return;
    End If;
  End If;

  --获取该收发记录的单据、药品ID、库房ID
  Select a.单据, a.No, a.库房id, a.药品id, a.费用id, a.入出类别id, a.记录状态, Nvl(a.批次, 0), a.批号, a.效期, a.供药单位id, a.产地, a.生产日期, a.批准文号,
         a.成本价, a.付数, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) As 实际数量, a.零售价, Nvl(b.是否变价, 0) 是否变价
  Into Int单据, Strno, Lng库房id, Lng药品id, Lng费用id, Lng入出类别id, Int记录状态, Lng批次, Str批号, Date效期, n_上次供应商id, v_上次产地, d_上次生产日期,
       v_批准文号, n_上次采购价, n_付数, n_原始数量, n_零售价, n_是否变价
  From 药品收发记录 A, 收费项目目录 B
  Where a.药品id = b.Id And a.Id = Billid_In;

  Begin
    Select Nvl(配药确认, 0), Nvl(配药, 0)
    Into v_配药确认, v_配药
    From 药房配药控制
    Where 药房id = Lng库房id And Rownum = 1;
  
  Exception
    When Others Then
      v_配药确认 := 0;
      v_配药     := 0;
      Null;
  End;

  If v_配药确认 = 0 And v_配药 = 0 Then
    v_排队状态 := 2;
  Elsif v_配药确认 = 1 Then
    v_排队状态 := 0;
  Elsif v_配药 = 1 Then
    v_排队状态 := 1;
  End If;

  --获取该笔记录剩余未退数量、金额及差价
  --尽量避免金额及差价未出完的现象
  Select Sum(Nvl(实际数量, 0) * Nvl(付数, 1)), Sum(Nvl(零售金额, 0)), Sum(Nvl(成本金额, 0)), Sum(Nvl(差价, 0))
  Into Dbl实际数量, Dbl实际金额, Dbl实际成本, Dbl实际差价
  From 药品收发记录
  Where 审核人 Is Not Null And NO = Strno And 单据 = Int单据 And 序号 = (Select 序号 From 药品收发记录 Where ID = Billid_In);

  --如果允许退药数为零，表示已退药
  If Dbl实际数量 = 0 Then
    v_Error := '该单据已被其他操作员退药，请刷新后再试！';
    Raise Err_Custom;
  End If;
  If Nvl(退药数量_In, 0) > Dbl实际数量 Then
    v_Error := '该单据已被其他操作员部分退药，请刷新后再试！';
    Raise Err_Custom;
  End If;

  --获取该药品当前是否分批的信息
  Select Nvl(药房分批, 0) Into Lng分批 From 药品规格 Where 药品id = Lng药品id;
  --如果是部分退药，则重新计算零售金额及差价
  Bln部分退药 := 0;
  If Not (退药数量_In Is Null Or Nvl(退药数量_In, 0) = Dbl实际数量) Then
    Bln部分退药 := 1;
  End If;
  If Bln部分退药 = 1 Then
    Dbl实际金额 := Round(Dbl实际金额 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际成本 := Round(Dbl实际成本 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际差价 := Round(Dbl实际差价 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际数量 := 退药数量_In;
  End If;

  If n_原始数量 = 退药数量_In Then
    Dbl实际数量 := 退药数量_In / n_付数;
  Else
    n_付数 := 1;
  End If;

  --lng分批:0-不分批;1-分批;2-原分批，现不分批，按不分批处理;3-原不分批，现分批，产生新批次
  If Lng分批 = 0 And Lng批次 <> 0 Then
    --原分批，现不分批，按不分批处理
    Lng分批 := 2;
  Elsif Lng分批 <> 0 And Lng批次 = 0 Then
    --原不分批,现分批,产生新的批次，并在新产生的发药记录中使用
    Lng分批 := 3;
  Else
    If Lng批次 = 0 Then
      Lng分批 := 0;
    Else
      Lng分批 := 1;
    End If;
  End If;
  --判断是否时价分批
  If (Lng分批 = 1 Or Lng分批 = 3) And n_是否变价 = 1 Then
    n_时价分批 := 1;
  Else
    n_时价分批 := 0;
  End If;

  --记录状态的含义有所变化
  --冲销的记录状态        :iif(int记录状态=1,0,1)+1
  --被冲销的记录状态        :iif(int记录状态=1,0,1)+2
  --等待发药的记录状态    :iif(int记录状态=1,0,1)+3

  --产生冲销记录
  Select 药品收发记录_Id.Nextval Into v_冲销记录id From Dual;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额,
     差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 外观, 领用人, 供药单位id, 生产日期, 批准文号, 汇总发药号, 发药方式, 注册证号)
    Select v_冲销记录id, Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 1, Int单据, Strno, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地,
           批号, 效期, n_付数, -dbl实际数量, -dbl实际数量, 成本价, -dbl实际成本, 扣率, 零售价, -dbl实际金额, -dbl实际差价, 摘要, People_In, Date_In, 配药人,
           People_In, Date_In, 费用id, 单量, 频次, 用法, 发药窗口, 退药库房_In, 退药人_In, 供药单位id, 生产日期, 批准文号, 汇总发药号_In, 发药方式, 注册证号
    From 药品收发记录
    Where ID = Billid_In;

  --如果是部分冲销，则付数填为1，实际数量为付数与实际数量的积
  --产生正常记录以供继续发药
  Select 药品收发记录_Id.Nextval Into Lng新批次 From Dual;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额,
     差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号)
    Select Lng新批次, Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 3, Int单据, Strno, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id,
           Decode(Lng分批, 1, 批次, 3, Lng新批次, 0), Decode(Lng分批, 3, 产地_In, 1, 产地, 产地), Decode(Lng分批, 3, 批号_In, 批号),
           Decode(Lng分批, 3, 效期_In, 效期), n_付数, Dbl实际数量, Dbl实际数量, 成本价, Dbl实际成本, 扣率, 零售价, Dbl实际金额, Dbl实际差价, 摘要, 填制人, 填制日期,
           Null, Null, Null, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号
    From 药品收发记录
    Where ID = Billid_In;

  --更新费用记录的执行状态(0-未执行;1-完全执行;2-部分执行)
  Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 0, 0, 0, 2)
  Into Int执行状态
  From 药品收发记录
  Where 单据 = Int单据 And NO = Strno And 费用id = Lng费用id And 审核人 Is Not Null;

  If 门诊_In = 1 Then
    Select 记录性质, 收费类别 Into n_记录性质, v_收费类别 From 门诊费用记录 Where ID = Lng费用id;
  Else
    Select 记录性质, 收费类别 Into n_记录性质, v_收费类别 From 住院费用记录 Where ID = Lng费用id;
  End If;

  If Int执行状态 = 0 Then
    If 门诊_In = 1 Then
      Update 门诊费用记录
      Set 执行状态 = Int执行状态, 执行人 = Null, 执行时间 = Null
      Where NO = Strno And
            序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = Billid_In)) And
            Mod(记录性质, 10) = n_记录性质 And 记录状态 <> 2 And 执行部门id = Lng库房id;
    Else
      Update 住院费用记录 Set 执行状态 = Int执行状态, 执行人 = Null, 执行时间 = Null Where ID = Lng费用id;
    End If;
  Else
    If 门诊_In = 1 Then
      Update 门诊费用记录
      Set 执行状态 = Int执行状态
      Where NO = Strno And
            序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = Billid_In)) And
            Mod(记录性质, 10) = n_记录性质 And 记录状态 <> 2 And 执行部门id = Lng库房id;
    Else
      Update 住院费用记录 Set 执行状态 = Int执行状态 Where ID = Lng费用id;
    End If;
  End If;

  --插入未发药品记录
  Begin
    If 门诊_In = 1 Then
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数, 领药号, 排队状态)
        Select a.单据, a.No, a.病人id, Null, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1,
               a.产品合格证, v_排队状态
        From (Select b.单据, b.No, a.病人id, a.姓名, Decode(a.记录状态, 0, 0, 1) 已收费, b.对方部门id, b.库房id, b.发药窗口, b.填制日期, c.身份,
                      b.产品合格证
               From 门诊费用记录 A, 药品收发记录 B, 病人信息 C
               Where b.Id = Billid_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
        Where b.名称(+) = a.身份;
    Else
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数, 领药号, 排队状态)
        Select a.单据, a.No, a.病人id, a.主页id, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1,
               a.产品合格证, v_排队状态
        From (Select b.单据, b.No, a.病人id, a.主页id, a.姓名, Decode(a.记录状态, 0, 0, 1) 已收费, b.对方部门id, b.库房id, b.发药窗口, b.填制日期,
                      c.身份, b.产品合格证
               From 住院费用记录 A, 药品收发记录 B, 病人信息 C
               Where b.Id = Billid_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
        Where b.名称(+) = a.身份;
    End If;
  Exception
    When Others Then
      Null;
  End;

  --修改处方类型
  Zl_Prescription_Type_Update(Strno, n_记录性质, Lng药品id, v_收费类别);

  --修改原记录为被冲销记录
  Update 药品收发记录 Set 记录状态 = Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 2 Where ID = Billid_In;

  --修改药品库存(反冲库存)
  If Lng分批 <> 3 Then
    Update 药品库存
    Set 实际数量 = Nvl(实际数量, 0) + Dbl实际数量 * n_付数, 实际金额 = Nvl(实际金额, 0) + Dbl实际金额, 实际差价 = Nvl(实际差价, 0) + Dbl实际差价
    Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(批次, 0) = Lng批次;
  
    If Sql%RowCount = 0 Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 零售价, 上次批号, 效期, 上次供应商id, 上次采购价, 上次产地, 上次生产日期, 批准文号, 平均成本价)
      Values
        (Lng库房id, Lng药品id, Decode(Lng分批, 2, 0, Lng批次), 1, Dbl实际数量 * n_付数, Dbl实际金额, Dbl实际差价,
         Decode(n_时价分批, 1, n_零售价, Null), Decode(Lng分批, 1, Str批号, Null), Decode(Lng分批, 1, Date效期, Null), n_上次供应商id,
         n_上次采购价, v_上次产地, d_上次生产日期, v_批准文号, n_上次采购价);
    End If;
  
    Zl_药品库存_可用数量异常处理(Lng库房id, Lng药品id, Lng批次);
  Else
    Insert Into 药品库存
      (库房id, 药品id, 批次, 效期, 性质, 实际数量, 实际金额, 实际差价, 零售价, 上次批号, 上次产地, 上次供应商id, 上次采购价, 上次生产日期, 批准文号, 平均成本价)
    Values
      (Lng库房id, Lng药品id, Lng新批次, 效期_In, 1, Dbl实际数量 * n_付数, Dbl实际金额, Dbl实际差价, Decode(n_时价分批, 1, n_零售价, Null), 批号_In,
       产地_In, n_上次供应商id, n_上次采购价, d_上次生产日期, v_批准文号, n_上次采购价);
  End If;

  Delete 药品库存
  Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  --处理调价修正
  Zl_药品收发记录_调价修正(v_冲销记录id);

  Begin
    --移动支付宝项目在发药后动态调用生成推送信息的过程
    Execute Immediate 'Begin zl_服务窗消息_发送(:1,:2); End;'
      Using 7, Billid_In || ',' || 退药数量_In || ',' || 门诊_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_部门退药;
/

--128124:蒋敏,2018-07-13,疾病编码管理新增手术操作类型
Create Or Replace Procedure Zl_疾病编码目录_Insert
(
  Id_In       In 疾病编码目录.Id%Type,
  编码_In     In 疾病编码目录.编码%Type,
  序号_In     In 疾病编码目录.序号%Type,
  附码_In     In 疾病编码目录.附码%Type,
  统计码_In   In 疾病编码目录.统计码%Type,
  名称_In     In 疾病编码目录.名称%Type,
  简码_In     In 疾病编码目录.简码%Type,
  说明_In     In 疾病编码目录.说明%Type,
  性别限制_In In 疾病编码目录.性别限制%Type,
  疗效限制_In In 疾病编码目录.疗效限制%Type,
  类别_In     In 疾病编码目录.类别%Type,
  手术类型_In In 疾病编码目录.手术类型%Type,
  分类id_In   In 疾病编码目录.分类id%Type,
  分娩_In     In 疾病编码目录.分娩%Type := Null,
  五笔码_In   In 疾病编码目录.五笔码%Type := Null,
  参数_In     In Varchar2, --科室ID串:科室ID1,科室ID2,科室ID3...
  应用_In     In Number := 0, --应用范围:0-应用于当前项目;1-应用于同级项目;2-应用于当前分类;3-应用于当前类别
  适用范围_In In 疾病编码目录.适用范围%Type := 0,
  适用应用_In In Number := 0, --适用应用范围:0-应用于当前项目;1-应用于同级项目;2-应用于当前分类;3-应用于当前类别
  手术操作类型_In In 疾病编码目录.手术操作类型%Type
) Is
  v_Infotmp Varchar2(4000);
  n_科室id  疾病编码科室.科室id%Type;
Begin
  Insert Into 疾病编码目录
    (ID, 编码, 序号, 附码, 统计码, 名称, 简码, 说明, 性别限制, 疗效限制, 手术类型, 分娩, 类别, 分类id, 五笔码, 适用范围,手术操作类型)
  Values
    (Id_In, 编码_In, 序号_In, 附码_In, 统计码_In, 名称_In, 简码_In, 说明_In, 性别限制_In, 疗效限制_In, 手术类型_In, 分娩_In, 类别_In, 分类id_In, 五笔码_In,
     适用范围_In,手术操作类型_In);

  If 参数_In Is Null Then
    v_Infotmp := Null;
  Else
    v_Infotmp := 参数_In || ',';
  End If;

  Delete 疾病编码科室 Where 疾病id = Id_In;
  While v_Infotmp Is Not Null Loop
    --分解单据ID串
    n_科室id  := Substr(v_Infotmp, 1, Instr(v_Infotmp, ',') - 1);
    v_Infotmp := Replace(',' || v_Infotmp, ',' || n_科室id || ',');
    Insert Into 疾病编码科室 (疾病id, 科室id) Values (Id_In, n_科室id);
  End Loop;

  If 应用_In <> 0 Then
    If 应用_In = 1 Then
      --应用于同级项目
      Delete 疾病编码科室 Where 疾病id In (Select ID From 疾病编码目录 Where 分类id = 分类id_In) And 疾病id <> Id_In;
    
      Insert Into 疾病编码科室
        (疾病id, 科室id)
        Select a.Id, b.科室id
        From 疾病编码目录 A, 疾病编码科室 B
        Where b.疾病id = Id_In And a.分类id = 分类id_In And a.Id <> Id_In;
    Elsif 应用_In = 2 Then
      --应用于当前分类
      Delete From 疾病编码科室
      Where 疾病id In (Select ID
                     From 疾病编码目录
                     Where 分类id In (Select ID
                                    From 疾病编码分类
                                    Where 类别 = 类别_In
                                    Start With ID = 分类id_In
                                    Connect By Prior ID = 上级id)) And 疾病id <> Id_In;
    
      Insert Into 疾病编码科室
        (疾病id, 科室id)
        Select a.Id, b.科室id
        From (Select ID
               From 疾病编码目录
               Where 分类id In (Select ID
                              From 疾病编码分类
                              Where 类别 = 类别_In
                              Start With ID = 分类id_In
                              Connect By Prior ID = 上级id)) A, 疾病编码科室 B
        Where b.疾病id = Id_In And a.Id <> Id_In;
    Elsif 应用_In = 3 Then
      --应用于当前类别      
      Delete From 疾病编码科室 Where 疾病id In (Select ID From 疾病编码目录 Where 类别 = 类别_In) And 疾病id <> Id_In;
    
      Insert Into 疾病编码科室
        (疾病id, 科室id)
        Select a.Id, b.科室id
        From (Select ID From 疾病编码目录 Where 类别 = 类别_In) A, 疾病编码科室 B
        Where b.疾病id = Id_In And a.Id <> Id_In;
    End If;
  End If;

  --适用范围应用
  If 适用应用_In = 1 Then
    --应用于同级项目
    Update 疾病编码目录 Set 适用范围 = 适用范围_In Where 类别 = 类别_In And 分类id = 分类id_In;
  Elsif 适用应用_In = 2 Then
    --应用于当前分类
    Update 疾病编码目录
    Set 适用范围 = 适用范围_In
    Where 类别 = 类别_In And
          分类id In
          (Select ID From 疾病编码分类 Where 类别 = 类别_In Start With ID = 分类id_In Connect By Prior ID = 上级id);
  Elsif 适用应用_In = 3 Then
    --应用于当前类别
    Update 疾病编码目录 Set 适用范围 = 适用范围_In Where 类别 = 类别_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病编码目录_Insert;
/
--128124:蒋敏,2018-07-13,疾病编码管理新增手术操作类型
Create Or Replace Procedure Zl_疾病编码目录_Update
(
  Id_In       In 疾病编码目录.Id%Type,
  编码_In     In 疾病编码目录.编码%Type,
  序号_In     In 疾病编码目录.序号%Type,
  附码_In     In 疾病编码目录.附码%Type,
  统计码_In   In 疾病编码目录.统计码%Type,
  名称_In     In 疾病编码目录.名称%Type,
  简码_In     In 疾病编码目录.简码%Type,
  说明_In     In 疾病编码目录.说明%Type,
  性别限制_In In 疾病编码目录.性别限制%Type,
  疗效限制_In In 疾病编码目录.疗效限制%Type,
  类别_In     In 疾病编码目录.类别%Type,
  手术类型_In In 疾病编码目录.手术类型%Type,
  分类id_In   In 疾病编码目录.分类id%Type,
  分娩_In     In 疾病编码目录.分娩%Type,
  五笔码_In   In 疾病编码目录.五笔码%Type := Null,
  参数_In     In Varchar2, --科室ID串:科室ID1,科室ID2,科室ID3...
  应用_In     In Number := 0, --应用范围:0-应用于当前项目;1-应用于同级项目;2-应用于当前分类;3-应用于当前类别
  适用范围_In In 疾病编码目录.适用范围%Type := 0,
  适用应用_In In Number := 0, --适用应用范围:0-应用于当前项目;1-应用于同级项目;2-应用于当前分类;3-应用于当前类别
  手术操作类型_In In 疾病编码目录.手术操作类型%Type
) Is
  v_Infotmp Varchar2(4000);
  n_科室id  疾病编码科室.科室id%Type;
Begin
  Update 疾病编码目录
  Set 编码 = 编码_In, 序号 = 序号_In, 附码 = 附码_In, 统计码 = 统计码_In, 名称 = 名称_In, 简码 = 简码_In, 说明 = 说明_In, 性别限制 = 性别限制_In,
      疗效限制 = 疗效限制_In, 手术类型 = 手术类型_In, 分娩 = 分娩_In, 分类id = 分类id_In, 五笔码 = 五笔码_In, 适用范围 = 适用范围_In, 手术操作类型 = 手术操作类型_In
  Where ID = Id_In;

  If 参数_In Is Null Then
    v_Infotmp := Null;
  Else
    v_Infotmp := 参数_In || ',';
  End If;

  Delete 疾病编码科室 Where 疾病id = Id_In;
  While v_Infotmp Is Not Null Loop
    --分解单据ID串
    n_科室id  := Substr(v_Infotmp, 1, Instr(v_Infotmp, ',') - 1);
    v_Infotmp := Replace(',' || v_Infotmp, ',' || n_科室id || ',');
  
    Insert Into 疾病编码科室 (疾病id, 科室id) Values (Id_In, n_科室id);
  End Loop;

  If 应用_In <> 0 Then
    If 应用_In = 1 Then
      --应用于同级项目
      Delete 疾病编码科室 Where 疾病id In (Select ID From 疾病编码目录 Where 分类id = 分类id_In) And 疾病id <> Id_In;
    
      Insert Into 疾病编码科室
        (疾病id, 科室id)
        Select a.Id, b.科室id
        From 疾病编码目录 A, 疾病编码科室 B
        Where b.疾病id = Id_In And a.分类id = 分类id_In And a.Id <> Id_In;
    Elsif 应用_In = 2 Then
       --应用于当前分类
      Delete From 疾病编码科室
      Where 疾病id In (Select ID
                     From 疾病编码目录
                     Where 分类id In (Select ID
                                    From 疾病编码分类
                                    Where 类别 = 类别_In
                                    Start With ID = 分类id_In
                                    Connect By Prior ID = 上级id)) And 疾病id <> Id_In;
    
      Insert Into 疾病编码科室
        (疾病id, 科室id)
        Select a.Id, b.科室id
        From (Select ID
               From 疾病编码目录
               Where 分类id In (Select ID
                              From 疾病编码分类
                              Where 类别 = 类别_In
                              Start With ID = 分类id_In
                              Connect By Prior ID = 上级id)) A, 疾病编码科室 B
        Where b.疾病id = Id_In And a.Id <> Id_In;
    Elsif 应用_In = 3 Then
      --应用于当前类别      
      Delete From 疾病编码科室 Where 疾病id In (Select ID From 疾病编码目录 Where 类别 = 类别_In) And 疾病id <> Id_In;
    
      Insert Into 疾病编码科室
        (疾病id, 科室id)
        Select a.Id, b.科室id
        From (Select ID From 疾病编码目录 Where 类别 = 类别_In) A, 疾病编码科室 B
        Where b.疾病id = Id_In And a.Id <> Id_In;
    End If;
  End If;

  --适用范围应用
  If 适用应用_In = 1 Then
    --应用于同级项目
    Update 疾病编码目录 Set 适用范围 = 适用范围_In Where 类别 = 类别_In And 分类id = 分类id_In;
  Elsif 适用应用_In = 2 Then
    --应用于当前分类
    Update 疾病编码目录
    Set 适用范围 = 适用范围_In
    Where 类别 = 类别_In And
          分类id In
          (Select ID From 疾病编码分类 Where 类别 = 类别_In Start With ID = 分类id_In Connect By Prior ID = 上级id);
  Elsif 适用应用_In = 3 Then
    --应用于当前类别
    Update 疾病编码目录 Set 适用范围 = 适用范围_In Where 类别 = 类别_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病编码目录_Update;
/

--129707:李业庆,2018-02-08,修正门诊标志错误
Create Or Replace Procedure Zl_药品收发记录_批量发料
(
  收发id_In     In Varchar2, --格式:id1,批次1|id2,批次2|.....
  库房id_In     In 药品收发记录.库房id%Type,
  审核人_In     In 药品收发记录.审核人%Type,
  审核日期_In   In 药品收发记录.审核日期%Type,
  发料方式_In   In 药品收发记录.发药方式%Type := 3, --1-处方发料;2-批量发料;3-部门发料;-1 停止发料
  领料人_In     In 药品收发记录.领用人%Type := Null,
  发料标识号_In In 药品收发记录.汇总发药号%Type := Null,
  配料人_In     In 药品收发记录.配药人%Type := Null,
  审核人编码_In In 人员表.编号%Type := Null
) Is
  --只读变量
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  v_批号     药品收发记录.批号%Type;
  v_上次产地 药品库存.上次产地%Type;
  v_批准文号 药品库存.批准文号%Type;

  v_Loop_Str Varchar2(4000);
  v_Fields   Varchar2(4000);

  n_Id       药品收发记录.Id%Type;
  n_批次     药品收发记录.批次%Type;
  n_成本价   药品收发记录.成本价%Type;
  n_库房id   药品收发记录.库房id%Type;
  n_库存金额 药品库存.实际金额%Type;
  n_库存差价 药品库存.实际差价%Type;
  n_未发数   未发药品记录.未发数%Type;
  --可写变量
  n_成本金额 药品收发记录.成本金额%Type;
  n_实际差价 药品收发记录.差价%Type;
  n_可用数量 药品收发记录.填写数量%Type;
  n_批次_Cur 药品收发记录.批次%Type;
  n_实价卫材 收费项目目录.是否变价%Type;

  n_上次供应商id       药品库存.上次供应商id%Type;
  n_上次采购价         药品库存.上次采购价%Type;
  n_执行状态           Number;
  n_差价率             Number;
  n_收费与发料分离     Number(1);
  n_小数               Number(2);
  n_成本价小数         Number(2);
  n_允许未审核处方发料 Number(2);
  n_序号               Number;

  d_效期                   药品收发记录.效期%Type;
  d_上次生产日期           药品库存.上次生产日期%Type;
  v_入库no                 药品收发记录.No%Type;
  v_入库库房id             药品收发记录.库房id%Type := 0;
  v_病人信息               Varchar2(200);
  n_虚拟库房               药品库存.库房id%Type;
  v_允许未审核记账单发料   Number(1);
  v_允许未收费的划价单发料 Number(1);
  v_自动审核记账单         Number(1);
  n_平均成本价             药品库存.平均成本价%Type;
  n_门诊费用               Number(1);
Begin
  --获取金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_小数 From Dual;
  --获取成本价小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(157), '2')) Into n_成本价小数 From Dual;

  Select zl_GetSysParameter('允许未审核的记账处方发料') Into v_允许未审核记账单发料 From Dual;
  Select zl_GetSysParameter('执行后自动审核划价单') Into v_自动审核记账单 From Dual;
  Select zl_GetSysParameter('允许未收费的门诊划价处方发料') Into v_允许未收费的划价单发料 From Dual;

  If 收发id_In Is Null Then
    v_Loop_Str := Null;
  Else
    v_Loop_Str := 收发id_In || '|';
  End If;

  While v_Loop_Str Is Not Null Loop
    --分解单据ID串
    v_Fields   := Substr(v_Loop_Str, 1, Instr(v_Loop_Str, '|') - 1);
    n_Id       := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    n_批次     := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    v_Loop_Str := Replace('|' || v_Loop_Str, '|' || v_Fields || '|');
  
    --检查相关操作
    v_Err_Msg := 'NO';
    For c_Check In (Select a.Id, a.单据, b.Id 费用id, a.No, b.No 费用no, b.病人id, Null 主页id, Null 病人病区id, b.病人科室id, b.开单部门id,
                           b.执行部门id, b.收入项目id, b.实收金额, b.操作员编号, b.操作员姓名, Nvl(b.记录状态, 0) As 审核标志, a.审核人,
                           Decode(Nvl(a.摘要, 'No拒发'), '拒发', 3, b.执行状态) 执行状态, b.门诊标志, 1 门诊费用
                    From 药品收发记录 A, 门诊费用记录 B
                    Where a.费用id = b.Id And a.Id = n_Id And a.审核人 Is Null And a.单据 In (24, 25, 26)
                    Union All
                    Select a.Id, a.单据, b.Id 费用id, a.No, b.No 费用no, b.病人id, b.主页id, b.病人病区id, b.病人科室id, b.开单部门id, b.执行部门id,
                           b.收入项目id, b.实收金额, b.操作员编号, b.操作员姓名, Nvl(b.记录状态, 0) As 审核标志, a.审核人,
                           Decode(Nvl(a.摘要, 'No拒发'), '拒发', 3, b.执行状态) 执行状态, b.门诊标志, 2 门诊费用
                    From 药品收发记录 A, 住院费用记录 B
                    Where a.费用id = b.Id And a.Id = n_Id And a.审核人 Is Null And a.单据 In (24, 25, 26)) Loop
      If Not (c_Check.审核人 Is Null) Then
        v_Err_Msg := '该处方[' || c_Check.No || ']已被其它操作员发料，操作被迫中止！';
        Raise Err_Item;
      End If;
      If Nvl(c_Check.执行状态, 0) = 3 Then
        v_Err_Msg := '该处方[' || c_Check.No || ']已拒发，操作被迫中止！';
        Raise Err_Item;
      End If;
    
      If Nvl(c_Check.审核标志, 0) = 0 And c_Check.单据 = 25 Then
        If v_允许未审核记账单发料 = 0 Then
          v_Err_Msg := '该处方[' || c_Check.No || ']还未审核，操作被迫中止！';
          Raise Err_Item;
        Else
          If v_自动审核记账单 = 1 Then
            --审核门诊和住院的单据
            If c_Check.操作员姓名 Is Null Then
              --审核门诊和住院的单据
              Zl_记帐记录_发料审核(c_Check.Id, c_Check.费用id, c_Check.费用no, c_Check.病人id, c_Check.主页id, c_Check.病人病区id,
                           c_Check.病人科室id, c_Check.开单部门id, c_Check.执行部门id, c_Check.收入项目id, c_Check.实收金额, 审核人编码_In,
                           审核人_In, c_Check.门诊标志, Null);
            Else
              --审核门诊和住院的单据
              Zl_记帐记录_发料审核(c_Check.Id, c_Check.费用id, c_Check.费用no, c_Check.病人id, c_Check.主页id, c_Check.病人病区id,
                           c_Check.病人科室id, c_Check.开单部门id, c_Check.执行部门id, c_Check.收入项目id, c_Check.实收金额, c_Check.操作员编号,
                           c_Check.操作员姓名, c_Check.门诊标志, Null);
            End If;
          End If;
        End If;
      End If;
    
      If Nvl(c_Check.审核标志, 0) = 0 And c_Check.单据 = 24 And v_允许未收费的划价单发料 = 0 Then
        v_Err_Msg := '该处方[' || c_Check.No || ']还未收费，操作被迫中止！';
        Raise Err_Item;
      End If;
    
      v_Err_Msg := 'Have';
    
      n_门诊费用 := c_Check.门诊费用;
    
    End Loop;
  
    If v_Err_Msg = 'NO' Then
      v_Err_Msg := '未找到指定单据,可能已经被其他操作员处理,操作被迫中止！';
      Raise Err_Item;
    End If;
  
    --获取该收发记录的单据、药品ID、库房ID,零售金额及实际数量、入出类别ID
    For c_收发 In (Select a.单据, a.No, a.药品id, a.库房id, a.费用id, a.零售价, Nvl(a.零售金额, 0) As 实际金额,
                        Nvl(a.实际数量, 0) * Nvl(a.付数, 1) As 实际数量, a.入出类别id, a.入出系数, Nvl(a.批次, 0) As 批次,
                        '[' || c.编码 || ']' || c.名称 As 名称, a.批号, a.效期, a.供药单位id, a.产地, a.生产日期, a.批准文号, a.商品条码, a.内部条码,
                        b.序号 As 费用序号
                 From 药品收发记录 A, 收费项目目录 C, 门诊费用记录 B
                 Where a.Id = n_Id And a.药品id = c.Id And a.费用id = b.Id And a.审核人 Is Null And a.单据 In (24, 25, 26)
                 Union All
                 Select a.单据, a.No, a.药品id, a.库房id, a.费用id, a.零售价, Nvl(a.零售金额, 0) As 实际金额,
                        Nvl(a.实际数量, 0) * Nvl(a.付数, 1) As 实际数量, a.入出类别id, a.入出系数, Nvl(a.批次, 0) As 批次,
                        '[' || c.编码 || ']' || c.名称 As 名称, a.批号, a.效期, a.供药单位id, a.产地, a.生产日期, a.批准文号, a.商品条码, a.内部条码,
                        b.序号 As 费用序号
                 From 药品收发记录 A, 收费项目目录 C, 住院费用记录 B
                 Where a.Id = n_Id And a.药品id = c.Id And a.费用id = b.Id And a.审核人 Is Null And a.单据 In (24, 25, 26)) Loop
      If Nvl(n_批次, 0) = 0 Then
        n_批次_Cur := c_收发.批次;
      Else
        n_批次_Cur := Nvl(n_批次, 0);
      End If;
    
      --检查是否已经填写库房
      n_收费与发料分离 := 0;
      If c_收发.库房id Is Null Then
        n_收费与发料分离 := 1;
      End If;
    
      n_库房id := 库房id_In;
      --取该批卫生材料的批号
      Begin
        Select 上次批号, 效期, Nvl(可用数量, 0), 上次供应商id, 上次产地, 上次生产日期, 批准文号, 上次采购价, Nvl(实际金额, 0) 实际金额, Nvl(实际差价, 0) 实际差价
        Into v_批号, d_效期, n_可用数量, n_上次供应商id, v_上次产地, d_上次生产日期, v_批准文号, n_上次采购价, n_库存金额, n_库存差价

        
        From 药品库存
        Where 库房id + 0 = n_库房id And 药品id = c_收发.药品id And 性质 = 1 And Nvl(批次, 0) = n_批次_Cur;
      Exception
        When Others Then
          n_库存金额   := 0;
          n_库存差价   := 0;
          n_上次采购价 := 0;
          n_可用数量   := 0;
      End;
    
      --高值卫材虚拟出库模式
      Begin
        Select 库房id
        Into n_虚拟库房
        From 药品收发记录
        Where 单据 = 21 And 审核日期 Is Null And 药品id = c_收发.药品id And Nvl(批次, 0) = c_收发.批次 And 费用id = c_收发.费用id And
              Rownum = 1;
      Exception
        When Others Then
          n_虚拟库房 := 0;
      End;
    
      --可用数量不足则退出
      If n_批次_Cur <> Nvl(c_收发.批次, 0) Then
        If n_虚拟库房 = 0 And n_可用数量 < Nvl(c_收发.实际数量, 0) And n_批次_Cur <> 0 Then
          v_Err_Msg := c_收发.名称 || '的可用数量不足，操作中止！';
          Raise Err_Item;
        End If;
      End If;
    
      If n_虚拟库房 = 0 Then
        --普通模式取发料部门价格
        n_成本价 := Round(Zl_Fun_Getoutcost(c_收发.药品id, c_收发.批次, n_库房id), n_成本价小数);
      Else
        --高值卫材虚拟出库模式取虚拟库房价格
        n_成本价 := Round(Zl_Fun_Getoutcost(c_收发.药品id, c_收发.批次, n_虚拟库房), n_成本价小数);
      End If;
      n_成本金额 := Round(n_成本价 * c_收发.实际数量, n_小数);
      n_实际差价 := Round(c_收发.实际金额 - n_成本金额, n_小数);
    
      --更新药品收发记录的零售金额、成本金额及差价
      Update 药品收发记录
      Set 成本价 = n_成本价, 成本金额 = n_成本金额, 差价 = n_实际差价, 库房id = n_库房id, 批次 = n_批次_Cur, 批号 = v_批号, 效期 = d_效期, 配药人 = 配料人_In,
          审核人 = 审核人_In, 审核日期 = 审核日期_In, 发药方式 = 发料方式_In, 领用人 = 领料人_In, 汇总发药号 = 发料标识号_In, 供药单位id = n_上次供应商id, 产地 = v_上次产地,
          生产日期 = d_上次生产日期, 批准文号 = v_批准文号
      Where ID = n_Id;
    
      --并发操作检查
      If Sql%RowCount = 0 Then
        v_Err_Msg := '不存在相关的发料记录，材料信息为:' || c_收发.名称 || '，操作中止！';
        Raise Err_Item;
      End If;
    
      --更新费用记录的执行状态(已执行)
      Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 1, 0, 1, 2)
      Into n_执行状态
      From 药品收发记录
      Where 单据 = c_收发.单据 And NO = c_收发.No And 费用id = c_收发.费用id And 审核人 Is Null;
    
      If n_门诊费用 = 1 Then
        Update 门诊费用记录
        Set 执行状态 = n_执行状态, 执行部门id = 库房id_In, 执行人 = 审核人_In, 执行时间 = 审核日期_In
        Where NO = c_收发.No And (Mod(记录性质, 10) = 1 Or Mod(记录性质, 10) = 2) And 记录状态 <> 2 And 序号 = c_收发.费用序号;
      Else
        Update 住院费用记录
        Set 执行状态 = n_执行状态, 执行部门id = 库房id_In, 执行人 = 审核人_In, 执行时间 = 审核日期_In
        Where ID = c_收发.费用id;
      End If;
    
      --更新未发药品记录(如果未发数为零则删除)
      Select Count(*)
      Into n_未发数
      From 药品收发记录
      Where 单据 = c_收发.单据 And NO = c_收发.No And 审核人 Is Null And (库房id + 0 = n_库房id Or 库房id Is Null) And
            Nvl(LTrim(RTrim(摘要)), 'No_拒发') <> '拒发';
    
      If n_未发数 = 0 Then
        Delete 未发药品记录 Where 单据 = c_收发.单据 And NO = c_收发.No And (库房id + 0 = n_库房id Or 库房id Is Null);
      End If;
    
      Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = c_收发.药品id;
    
      --更新原批次库存的可用数量
      --更新发料批次库存的可用及实际数量
      If c_收发.批次 <> n_批次_Cur Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + c_收发.实际数量,
            零售价 = Decode(n_实价卫材, 1, Decode(Nvl(c_收发.批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_收发.零售价, 零售价)), Null)
        Where 性质 = 1 And 库房id + 0 = n_库房id And 药品id = c_收发.药品id And Nvl(批次, 0) = c_收发.批次;
      
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - c_收发.实际数量,
            零售价 = Decode(n_实价卫材, 1, Decode(Nvl(n_批次_Cur, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_收发.零售价, 零售价)), Null)
        Where 性质 = 1 And 库房id + 0 = n_库房id And 药品id = c_收发.药品id And Nvl(批次, 0) = n_批次_Cur;
      End If;
    
      If n_收费与发料分离 = 1 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - c_收发.实际数量, 实际数量 = Nvl(实际数量, 0) - c_收发.实际数量, 实际金额 = Nvl(实际金额, 0) - c_收发.实际金额,
            实际差价 = Nvl(实际差价, 0) - n_实际差价,
            零售价 = Decode(n_实价卫材, 1, Decode(Nvl(n_批次_Cur, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_收发.零售价, 零售价)), Null),
            上次采购价 = Decode(上次采购价, Null, n_成本价, 上次采购价), 平均成本价 = Decode(平均成本价, Null, n_成本价, 平均成本价),
            商品条码 = Decode(商品条码, Null, c_收发.商品条码, 商品条码), 内部条码 = Decode(内部条码, Null, c_收发.内部条码, 内部条码),
            效期 = Decode(效期, Null, c_收发.效期, 效期), 上次批号 = Decode(上次批号, Null, c_收发.批号, 上次批号),
            上次生产日期 = Decode(上次生产日期, Null, c_收发.生产日期, 上次生产日期), 上次产地 = Decode(上次产地, Null, c_收发.产地, 上次产地)
        Where 库房id + 0 = n_库房id And 药品id = c_收发.药品id And 性质 = 1 And Nvl(批次, 0) = n_批次_Cur;
      Else
        Update 药品库存
        Set 实际数量 = Nvl(实际数量, 0) - c_收发.实际数量, 实际金额 = Nvl(实际金额, 0) - c_收发.实际金额, 实际差价 = Nvl(实际差价, 0) - n_实际差价,
            零售价 = Decode(n_实价卫材, 1, Decode(Nvl(n_批次_Cur, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_收发.零售价, 零售价)), Null),
            上次采购价 = Decode(上次采购价, Null, n_成本价, 上次采购价), 平均成本价 = Decode(平均成本价, Null, n_成本价, 平均成本价),
            商品条码 = Decode(商品条码, Null, c_收发.商品条码, 商品条码), 内部条码 = Decode(内部条码, Null, c_收发.内部条码, 内部条码),
            效期 = Decode(效期, Null, c_收发.效期, 效期), 上次批号 = Decode(上次批号, Null, c_收发.批号, 上次批号),
            上次生产日期 = Decode(上次生产日期, Null, c_收发.生产日期, 上次生产日期), 上次产地 = Decode(上次产地, Null, c_收发.产地, 上次产地)
        Where 库房id + 0 = n_库房id And 药品id = c_收发.药品id And 性质 = 1 And Nvl(批次, 0) = n_批次_Cur;
      End If;
    
      If Sql%RowCount = 0 Then
        If n_上次采购价 = 0 Then
          If Nvl(c_收发.实际数量, 0) = 0 Then
            n_上次采购价 := Round(n_成本金额, 5);
          Else
            n_上次采购价 := Round(n_成本金额 / c_收发.实际数量, 5);
          End If;
        
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次批号, 上次产地, 上次供应商id, 上次采购价, 上次生产日期, 批准文号, 效期, 零售价, 商品条码, 内部条码,
             平均成本价)
          Values
            (n_库房id, c_收发.药品id, n_批次_Cur, 1, 0 - c_收发.实际数量, 0 - c_收发.实际数量, 0 - c_收发.实际金额, 0 - n_实际差价, v_批号, v_上次产地,
             n_上次供应商id, n_上次采购价, d_上次生产日期, v_批准文号, d_效期,
             Decode(n_实价卫材, 1, Decode(Nvl(n_批次_Cur, 0), 0, Null, c_收发.零售价), Null), c_收发.商品条码, c_收发.内部条码, n_上次采购价);
        End If;
      
      End If;
      Delete 药品库存
      Where 性质 = 1 And 库房id + 0 = n_库房id And 药品id = c_收发.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
            Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
    
      If n_虚拟库房 > 0 Then
        --审核备货卫材在虚拟库房的其他出库单据
        For v_出库 In (Select 序号, NO, 库房id, 药品id, Nvl(批次, 0) As 批次, 实际数量, 成本价, 成本金额, 零售金额, 差价, 入出类别id
                     From 药品收发记录
                     Where 单据 = 21 And 审核日期 Is Null And 药品id = c_收发.药品id And Nvl(批次, 0) = c_收发.批次 And 费用id = c_收发.费用id) Loop
        
          Update 药品收发记录
          Set 汇总发药号 = n_Id
          Where 单据 = 21 And 审核日期 Is Null And 药品id = c_收发.药品id And Nvl(批次, 0) = c_收发.批次 And 费用id = c_收发.费用id;
        
          Zl_材料其他出库_Verify(v_出库.序号, v_出库.No, v_出库.库房id, v_出库.药品id, v_出库.批次, v_出库.实际数量, v_出库.成本价, v_出库.成本金额, v_出库.零售金额,
                           v_出库.差价, v_出库.入出类别id, 审核人_In, 审核日期_In);
        End Loop;
      
        --产生备货卫材在卫材仓库的外购入库单据
        For v_入库 In (Select NO, 序号, 供药单位id, 药品id, 产地, 批号, 生产日期, 效期, 灭菌日期, 灭菌效期, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要,
                            注册证号, Nvl(批次, 0) As 批次, 商品条码, 内部条码
                     From 药品收发记录
                     Where 单据 = 21 And 审核日期 Is Not Null And 药品id = c_收发.药品id And Nvl(批次, 0) = c_收发.批次 And
                           费用id = c_收发.费用id And 汇总发药号 = n_Id) Loop
          Begin
            Select 库房id Into v_入库库房id From 虚拟库房对照 Where 科室id = c_收发.库房id;
          Exception
            When Others Then
              v_入库库房id := 0;
          End;
        
          If v_入库库房id > 0 Then
          
            --同一张发料单产生的入库单的NO要一致
            Select Max(NO), Max(序号) + 1
            Into v_入库no, n_序号
            From 药品收发记录
            Where 单据 = 15 And 审核日期 Is Null And 供药单位id = v_入库.供药单位id And
                  费用id In
                  (Select Distinct 费用id
                   From 药品收发记录
                   Where 单据 = 21 And 审核日期 Is Not Null And
                         NO = (Select Distinct NO
                               From 药品收发记录
                               Where 单据 = 21 And 审核日期 Is Not Null And 费用id = c_收发.费用id And 汇总发药号 = n_Id));
          
            If v_入库no Is Null Or v_入库no = '' Then
              --如果入库NO为Null, 产生新的入库单NO
              v_入库no := Nextno(68, v_入库库房id);
              n_序号   := 1;
            End If;
          
            Begin
              If n_门诊费用 = 1 Then
                Select b.名称 || ',' || a.姓名 || ',' || a.标识号 || ',' || '' As 病人信息
                Into v_病人信息
                From 门诊费用记录 A, 部门表 B
                Where a.病人科室id = b.Id And a.Id = c_收发.费用id;
              Else
                Select b.名称 || ',' || a.姓名 || ',' || a.标识号 || ',' || a.床号 As 病人信息
                Into v_病人信息
                From 住院费用记录 A, 部门表 B
                Where a.病人科室id = b.Id And a.Id = c_收发.费用id;
              End If;
            Exception
              When Others Then
                v_病人信息 := '';
            End;
          
            Zl_材料外购_Insert(v_入库no, n_序号, v_入库库房id, v_入库.供药单位id, v_入库.药品id, v_入库.产地, v_入库.批号, v_入库.生产日期, v_入库.效期,
                           v_入库.灭菌日期, v_入库.灭菌效期, v_入库.实际数量, v_入库.成本价, v_入库.成本金额, v_入库.扣率, v_入库.零售价, v_入库.零售金额, v_入库.差价,
                           Null, '【自动入账】' || v_入库.摘要, v_入库.注册证号, 审核人_In, Null, Null, Null, Null, 审核日期_In, Null, Null,
                           v_入库.批次, 1, v_病人信息, v_入库.商品条码, v_入库.内部条码, c_收发.费用id);
          End If;
        End Loop;
      End If;
    
      --修正误差数据
      Zl_材料收发记录_调价修正(n_Id);
    End Loop;
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_批量发料;
/


--129869:蒋廷中,2019-01-11,新增用药清单用药配方的转出脚本
CREATE OR REPLACE Procedure Zl1_Datamove_Tag
(
  d_End            In Date,
  n_批次           In Number,
  n_System         In Number,
  n_预交剩余款上限 In 病人预交记录.金额%Type := 10 --当病人不存在未结费用，也不是在院病人时，允许未冲完的预交款在指定值以下的数据强制转出，避免大量呆帐未转出从而影响转出速度



) As
  --功能：标记待转出的数据
  --说明：为避免Undo表空间膨胀过大，分段提交
  d_Lastend Date; --最终转出截止时间（d_End为本批转出截止时间）

  --递归取消“一张预交款单据中的一部分被标记为待转出”的数据
  Procedure Datamove_Tag_Update
  (
    结帐id_In t_Numlist,
    d_End     In Date,
    n_批次    In Number
  ) As

    c_结帐id t_Numlist := t_Numlist();
    c_No     t_Strlist := t_Strlist();
  Begin
    --1.1一张预交单据被多个结帐ID冲了，找出其中的一部分被标记为待转出的数据，如：
    --   NO=A001 记录性质=11 结帐ID=10 待转出=1
    --   NO=A001 记录性质=11 结帐ID=11 待转出=NULL
    If 结帐id_In Is Null Then
      Select Distinct a.No
      Bulk Collect
      Into c_No
      From 病人预交记录 A
      Where a.记录性质 In (1, 11) And a.待转出 = n_批次 And Exists
       (Select 1 From 病人预交记录 Where NO = a.No And 记录性质 In (1, 11) And 待转出 Is Null);
    Else
      Select Distinct a.No
      Bulk Collect
      Into c_No
      From 病人预交记录 A
      Where a.结帐id In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(结帐id_In) B) And a.记录性质 In (1, 11) And a.待转出 Is Null And Exists
       (Select 1 From 病人预交记录 Where NO = a.No And 记录性质 In (1, 11) And 待转出 + 0 = n_批次);
    End If;

    If c_No.Count = 0 Then
      Return;
    End If;

    --1.2取消标记
    Forall I In 1 .. c_No.Count
      Update 病人预交记录 Set 待转出 = Null Where NO = c_No(I) And 记录性质 In (1, 11);

    --------------------------------------------------------------------------------------------------------
    --2.1一个结帐ID冲了多张预交单据，找出其中的一部分被标记为待转出的数据，如：
    --   NO=A001 记录性质=11 结帐ID=20 待转出=1
    --   NO=A002 记录性质=11 结帐ID=20 待转出=NULL
    Select Distinct a.结帐id
    Bulk Collect
    Into c_结帐id
    From 病人预交记录 A
    Where a.No In (Select /*+cardinality(b,10) */
                    Column_Value
                   From Table(c_No) B) And a.记录性质 In (1, 11) And a.待转出 Is Null And a.收款时间 + 0 < d_End And Exists
     (Select 1 From 病人预交记录 Where 结帐id = a.结帐id And 待转出 + 0 = n_批次);

    If c_结帐id.Count = 0 Then
      Return;
    End If;

    --2.2取消标记(包括一次结帐的其他结算方式的记录)
    Forall I In 1 .. c_结帐id.Count
      Update 病人预交记录 Set 待转出 = Null Where 结帐id = c_结帐id(I);

    --递归调用
    Datamove_Tag_Update(c_结帐id, d_End, n_批次);
  End Datamove_Tag_Update;
Begin
  Select 本次最终日期 Into d_Lastend From zlDataMove Where 系统 = n_System And 组号 = 1;
  If d_Lastend Is Null Then
    Return;
  End If;
  --新加子查询注意性能优化，把能够将数据过滤到最小的条件放到最后，Exists类条件放前面

  --1.经济核算（费用,药品,收款和票据等）
  --冲销业务与原始业务的发生时间相同，登记时间不同，所以要按发生时间来查询.
  --以下情况，可能有多个结帐ID，或涉及多个费用单据，这些数据要一起转出或排除转出，否则影响后续判断是否结清
  --1.一张费用单据的一行费用或多行费用可能分多次结帐（有多个不同的结帐ID）
  --2.结帐作废后也可能分多次结清(一张单据多个不同的结帐ID)
  --3.结帐作废后可能与其他费用单据一起结(一张单据的多个结帐ID，涉及多个费用NO，这些NO可能之前结帐作废过，有其他结帐ID)
  --考虑到这情况的复杂性，为简化逻辑，提升查询性能，按病人ID来排除(该病人的结帐数据都不转出)

  Update /*+ rule*/ 病人预交记录 L
  Set 待转出 = n_批次
  Where 结帐id In
        (Select Distinct a.结帐id --1.门诊收费和挂号的收费结算记录
         From 门诊费用记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 门诊费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_Lastend)) And a.待转出 Is Null And
               a.记录性质 In (1, 4) And a.发生时间 < d_End And a.登记时间 < d_Lastend
         Union All
         Select Distinct b.结算id --2.医保补结算(没有发生时间字段,作废记录的登记时间不同，为了把收费和作废的一次性转出，所以要连接B表)
         From 费用补充记录 A, 费用补充记录 B
         Where a.待转出 Is Null And a.No = b.No And a.记录性质 = b.记录性质 And a.登记时间 < d_End
         Union All
         Select Distinct a.结帐id --3.就诊卡的收费结算记录(排除之后退卡费的,一张单据中只要其中一行退了)
         From 住院费用记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 住院费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_Lastend)) And a.待转出 Is Null And
               a.记帐费用 = 0 And a.记录性质 = 5 And a.发生时间 < d_End
         Union All --4.住院记帐费用的结帐结算记录
         Select 结帐id
         From (With Settle As (Select Distinct c.结帐id
                               From (Select Distinct b.No, b.序号, Mod(b.记录性质, 10) As 记录性质
                                      From (Select Distinct b.Id
                                             From 病人结帐记录 A, 病人结帐记录 B --作废的结帐单的收费时间可能在指定时间之后，所以要连接B表
                                             Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                                                    (Select 1
                                                     From 病人结帐记录 C
                                                     Where a.No = c.No And c.记录状态 = 2 And c.收费时间 >= d_Lastend)) And
                                                   a.待转出 Is Null And a.No = b.No And (a.结帐类型 = 2 Or Nvl(a.结帐类型, 0) = 0) And
                                                   a.收费时间 < d_End) A, 住院费用记录 B
                                      Where a.Id = b.结帐id) B, 住院费用记录 C --通过C表找到这些费用单据的所有结帐ID一起转(可能在转出时间之后)
                               Where c.No = b.No And Mod(c.记录性质, 10) = b.记录性质 And c.序号 = b.序号)
                Select 结帐id
                From Settle
                Minus
                Select Distinct a.Id
                From 病人结帐记录 A,
                     (Select Distinct 病人id
                       From (Select c.病人id, c.No, Mod(c.记录性质, 10) As 记录性质, Nvl(Sum(c.实收金额), 0) As 实收金额,
                                     Nvl(Sum(c.结帐金额), 0) As 结帐金额
                              From 住院费用记录 C, Settle S
                              Where c.结帐id = s.结帐id
                              Group By c.No, Mod(c.记录性质, 10), c.病人id) C
                       Where c.实收金额 <> c.结帐金额 And Exists (Select 1 From 在院病人 F Where c.病人id = f.病人id) --出院病人没有结清的也转走（在需要时再抽回），否则排除的数据量太大
                             Or Exists (Select 1
                              From 住院费用记录 E, 病人结帐记录 S
                              Where e.No = c.No And Mod(e.记录性质, 10) = c.记录性质 And e.结帐id = s.Id And
                                    s.待转出 Is Null And s.收费时间 >= d_Lastend)) N --即使是在本批转出时间之后结清，只要不是在最终转出时间之后，就不排除



                Where a.病人id = n.病人id And (a.结帐类型 = 2 Or Nvl(a.结帐类型, 0) = 0))
                Union All --5.门诊记帐费用的结帐结算记录
                Select 结帐id
                From (With Settle As (Select Distinct c.结帐id
                                      From (Select Distinct b.No, b.序号, Mod(b.记录性质, 10) As 记录性质
                                             From (Select Distinct b.Id
                                                    From 病人结帐记录 A, 病人结帐记录 B
                                                    Where a.待转出 Is Null And a.No = b.No And (a.结帐类型 = 1 Or Nvl(a.结帐类型, 0) = 0) And
                                                          a.收费时间 < d_End) A, 门诊费用记录 B
                                             Where a.Id = b.结帐id) B, 门诊费用记录 C
                                      Where c.No = b.No And Mod(c.记录性质, 10) = b.记录性质 And c.序号 = b.序号)
                       Select 结帐id
                       From Settle
                       Minus
                       Select Distinct a.Id
                       From 病人结帐记录 A,
                            (Select Distinct c.病人id
                              From (Select c.病人id, c.No, Mod(c.记录性质, 10) As 记录性质, Nvl(Sum(c.实收金额), 0) As 实收金额,
                                            Nvl(Sum(c.结帐金额), 0) As 结帐金额
                                     From 门诊费用记录 C, Settle S
                                     Where c.结帐id = s.结帐id
                                     Group By c.No, Mod(c.记录性质, 10), c.病人id) C
                              Where c.实收金额 <> c.结帐金额 --门诊病人没有结清的不转走
                                    Or Exists (Select 1
                                     From 门诊费用记录 E, 病人结帐记录 S
                                     Where e.No = c.No And Mod(e.记录性质, 10) = c.记录性质 And e.结帐id = s.Id And
                                           s.待转出 Is Null And s.收费时间 >= d_Lastend)) N
                       Where a.病人id = n.病人id And (a.结帐类型 = 1 Or Nvl(a.结帐类型, 0) = 0))
         );

  --排除预交款未冲完的
  --为了降低逻辑的复杂性，不排除在转出时间之后发药或未发药的费用记录对应的结帐ID，将这种情况的结算数据和费用数据强制转走
  --因为前面的SQL查出的结帐ID可能不全是冲预交的(门诊收费和住院结帐补费等)，所以，需要单独一个SQL来排除
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = Null
  Where 待转出 = n_批次 And
        结帐id In
        (Select Distinct d.结帐id --该单据相关的所有冲预交的结帐ID都不转出
         From 病人预交记录 D,
              (Select Distinct l.No
                From (Select l.No, l.病人id, l.预交类别, Nvl(Sum(l.金额), 0) As 金额, Nvl(Sum(l.冲预交), 0) As 冲预交,
                              Sum(Decode(l.待转出, Null, Decode(结帐ID,Null,Decode(记录状态,2,0,1),1), 0)) As 未转出
                       From 病人预交记录 L --可能按结帐ID确认本次待转出的冲的只是剩余款，所以需要连接L表，查原始交预交的单据，以及记录性质为11的可能还有转出时间之后其他冲剩余款的结帐ID
                       Where l.记录性质 In (1, 11) And
                             l.No In
                             (Select Distinct p.No From 病人预交记录 P Where p.记录性质 In (1, 11) And p.待转出 = n_批次)
                       Group By l.No, l.病人id, l.预交类别) L --多次住院可以一次结清，所以，不能加主页ID
                Where 未转出 > 0 --只要该预交单据还有未转出的预交或冲预交记录，则不转出，避免转出一部分导致后续判断错误
                      Or
                      l.金额 <> l.冲预交 And
                      (Exists (Select 1
                               From 病人预交记录 E --剩的预交款，一般用负数交预交来退款（NO号不同），这种相当于是冲完了，不排除
                               Where e.病人id = l.病人id And e.预交类别 = l.预交类别 And e.记录性质 In (1, 11) And
                                     (e.待转出 = n_批次 Or e.待转出 Is Null And e.结帐id Is Null And e.记录性质 = 1 And 收款时间 < d_End)
                                Having abs(Nvl(Sum(e.金额), 0) - Nvl(Sum(e.冲预交), 0)) > n_预交剩余款上限) --余额小于等于n不排除，与下面第3种结帐ID为空的要保持一致
                       Or l.预交类别 = 2 And Exists (Select 1 From 在院病人 E Where l.病人id = e.病人id) Or Exists
                       (Select 1
                        From 病人未结费用 E
                        Where l.病人id = e.病人id And (l.预交类别 = 1 And e.主页id Is Null Or l.预交类别 = 2 And e.主页id Is Not Null)))) N
         Where d.No = n.No And d.记录性质 In (1, 11));

  --单独处理3种结帐ID为空的预交记录
  --1.预交款没有使用就直接退了的记录(结帐ID为空)
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 记录性质 = 1 And
        NO In (Select a.No
               From 病人预交记录 A
               Where a.结帐id Is Null And a.记录性质 = 1 And a.记录状态 In (2, 3) And a.待转出 Is Null And a.收款时间 < d_End
               Group By a.No
               Having Sum(a.金额) = 0);

  --2.交预交款后退款的记录（结帐ID为空，记录状态为2）
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 结帐id Is Null And 记录性质 = 1 And 记录状态 = 2 And
        NO In (Select a.No From 病人预交记录 A Where a.记录性质 = 1 And a.记录状态 = 3 And a.待转出 = n_批次);

  --排除同一张预交款单据部分记录被标记为转出的,只要有不转出的，则整张单据都不转出
  --跟第2种有关联影响，所以要放在它之后执行
  --要影响第3种情况的判断，所以要放在它之前执行
  Datamove_Tag_Update(Null, d_End, n_批次);

  --3.预交款未用完时用交负数预交来退款(结帐ID为空，并且跟原始的冲预交的NO没有关联关系)
  --不加条件"金额 < 0"，因为存在预交款没有使用过，就直接用交负数预交来退款的情况
  Update /*+ rule*/ 病人预交记录 L
  Set 待转出 = n_批次
  Where Exists (Select 1
         From 病人预交记录 E
         Where e.病人id = l.病人id And e.预交类别 = l.预交类别 And e.记录性质 In (1, 11) And
               (e.待转出 = n_批次 Or e.待转出 Is Null And e.结帐id Is Null And e.记录性质 = 1 And 记录状态 = 1 And 收款时间 < d_End)
         Group By e.病人id
         Having abs(Nvl(Sum(e.金额), 0) - Nvl(Sum(e.冲预交), 0)) <= n_预交剩余款上限) --余额小于等于n要转出，与前面“排除预交款未冲完的”要保持一致

        And Exists (Select 1
         From 病人预交记录 E
         Where e.病人id = l.病人id And e.预交类别 = l.预交类别 And e.记录性质 In (1, 11) And e.待转出 = n_批次) And
        待转出 Is Null And 结帐id Is Null And 记录性质 = 1 And 记录状态 = 1 And 收款时间 < d_End;

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
  Where (记录id, 结帐id) In (Select a.Id, a.结帐id From 病人预交记录 A Where 待转出 = n_批次);

  --1.挂号费用异常数据
  --a.结帐ID为空（实收金额可能不为零）
  --b.结帐ID不为空，打折后实收金额为0（应收金额正负冲销）的挂号费用，没有挂号记录，也没有预交记录
  --按发生时间转出，因为收和退的发生时间相同，登记时间不同。
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where 待转出 Is Null And 发生时间 < d_End And 记录性质 = 4 And (实收金额 = 0 Or 结帐id Is Null);

  --2.直接收费的和结帐无结算（预交）记录的，Union不加all去掉重复以减少in的数量
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where 结帐id In
        (Select 结帐id From 病人预交记录 Where 待转出 = n_批次 Union Select ID From 病人结帐记录 Where 待转出 = n_批次);

  --3.没有结帐id的数据(按发生时间)
  --a.未结帐的划价记录
  --b.未收费的零费用
  --加条件"待转出 Is Null"是为了处理连续多次标记转出的情况
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (记录状态 = 0 Or 记录性质 = 1 And 实收金额 = 0 And 结帐金额 = 0) And 结帐id Is Null And 待转出 Is Null And 发生时间 < d_End;

  --4.没有结帐id的数据(按发生时间)
  --未结帐的门诊记帐费用(赖账)，该病人没有预交余额，并且病人在最终转出时间之后无未结门诊记帐费用
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where Not Exists (Select 1
         From 病人预交记录 B
         Where b.病人id = a.病人id And b.待转出 Is Null And b.预交类别 = 1 And b.记录性质 In (1, 11) Having
          Nvl(Sum(b.金额), 0) <> Nvl(Sum(b.冲预交), 0)) And Not Exists
   (Select 1
         From 门诊费用记录 B
         Where a.病人id = b.病人id And b.记录性质 = 2 And b.结帐id Is Null And b.待转出 Is Null And b.登记时间 > = d_Lastend) And
        记录性质 = 2 And 结帐id Is Null And 待转出 Is Null And 发生时间 < d_End;

  --5.没有结帐id的数据(按发生时间)
  --冲销产生的记帐记录（记录状态为2），登记时间可能在当前指定转出时间之后，而原始记帐记录（记录状态为3），登记时间在指定转出时间之前。前后两者的发生时间是相同的。
  --a.未结帐的零记帐费用或打折后实收金额为零的（结帐模块参数没有勾选对零费用结帐）
  --b.结帐作废后，记帐单销帐的记录（结帐ID为空且记录状态为2的），记录状态为3的且有结帐ID的在最前面已转出.
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

  --6.有结帐id的零费用(按发生时间)
  --a.按费别打折后结帐金额为零的收费记录,
  --b.一张单据相同结帐ID的结帐金额之和为0(冲销后为零)
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
  --冲销产生的记帐记录（记录状态为2），原始记录和冲销记录的发生时间是相同的。
  --1)转出结帐作废后，记帐单销帐的记录（记录状态为2，且没有结帐ID，且(记录状态为3的有结帐ID的)在最前面已转出）
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
           Having Nvl(Sum(b.实收金额), 0) = 0)) And a.记录性质 In (2, 3, 5) Or a.记录状态 = 0) And a.结帐id Is Null And a.待转出 Is Null And
        a.发生时间 < d_End;

  --3.离院未结帐的（赖帐病人），因为是很久以前的这些数据，如果预交已冲完，则处理为要转出
  --去掉病案主页中的"数据转出 is null"的条件，是因为一些病人可能在之前的批次中已转出了
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where 待转出 Is Null And 结帐id Is Null And
        (病人id, 主页id) In (Select 病人id, 主页id
                         From 病案主页 C
                         Where 出院日期 < d_End And 待转出 Is Null And Not Exists
                          (Select 1
                                From 病人预交记录 B
                                Where b.病人id = c.病人id And b.待转出 Is Null And b.预交类别 = 2 And b.记录性质 In (1, 11) Having
                                 Nvl(Sum(b.金额), 0) <> Nvl(Sum(b.冲预交), 0)));

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

  Update /*+ rule*/ 药品收发门诊标志 A
  Set 待转出 = n_批次
  Where (a.处方号, a.单据) In (Select b.No, b.单据 From 药品收发记录 B Where b.待转出 = n_批次);

  Update /*+ rule*/ 药品收发住院标志
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

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
  Where Not Exists (Select 1 From 票据使用明细 B Where b.领用id = a.Id And b.使用时间 >= d_Lastend) And 待转出 Is Null And 剩余数量 = 0 And
        登记时间 < d_End;

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
  --不转出的条件：挂号费用未转出的，最终转出时间之后存在医嘱（这些医嘱因为时间没有到，不应转出），医嘱对应的费用未转出的
  --即使正在就诊(r.执行状态 <> 2 )的也强制转出(医生可能没有使用完成就诊功能)
  Update /*+ rule*/ 病人挂号记录 T
  Set 待转出 = n_批次
  Where Rowid In
        (Select Rowid
         From 病人挂号记录 R
         Where Not Exists (Select 1 From 门诊费用记录 A Where r.No = a.No And a.记录性质 = 4 And a.待转出 Is Null) And Not Exists
          (Select 1
                From 病人医嘱记录 A
                Where a.挂号单 = r.No And a.待转出 Is Null And a.病人来源 <> 4 And Nvl(a.停嘱时间, a.开嘱时间) >= d_Lastend) And
               Not Exists (Select 1
                From 门诊费用记录 E, 病人医嘱记录 A
                Where r.No = a.挂号单 And a.Id = e.医嘱序号 And a.病人来源 <> 4 And e.待转出 Is Null) And
               r.待转出 Is Null And r.发生时间 < d_End);

  --由于有一部分挂号数据未转出，所以，汇总表的数据可能与挂号数据不匹配
  Update 病人挂号汇总 Set 待转出 = n_批次 Where 待转出 Is Null And 日期 < d_End;
  Update /*+ rule*/ 病人转诊记录 Set 待转出 = n_批次 Where NO In (Select NO From 病人挂号记录 Where 待转出 = n_批次);

  --通过"住院费用记录"来查询，而不是"病人结帐记录",因为离院未结的赖帐病人也转出了费用
  --出院日期条件仍然需要，因为可能某次结帐转出了，但病人在最终转出截止时间之前并未出院(一次住院多次结帐)。
  --通过指定索引方式进行特殊优化（缺省采用"病案主页IX_出院日期"索引的效率太低）
  --不加"数据转出 is null"的条件，因为一次住院多次结帐时，如果跨不同的转出批次(转出截止时间)，该字段将会被更新多次。
  Update /*+ rule*/ 病案主页 P
  Set 待转出 = n_批次
  Where Not Exists
   (Select 1 From 住院费用记录 A Where a.病人id = p.病人id And a.主页id = p.主页id And a.待转出 Is Null) And 待转出 Is Null And
        出院日期 < d_Lastend And (病人id, 主页id) In (Select Distinct 病人id, 主页id From 住院费用记录 Where 待转出 = n_批次);

  --已出院，但没有费用的，也标记为转出，以便转出病历数据
  Update /*+ rule*/ 病案主页 P
  Set 待转出 = n_批次
  Where Not Exists (Select 1 From 住院费用记录 A Where a.病人id = p.病人id And a.主页id = p.主页id) And 待转出 Is Null And 数据转出 Is Null And
        出院日期 < d_End;

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
  Where 病人来源 = 1 And (病人id, 主页id) In (Select 病人id, ID From 病人挂号记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 电子病历记录
  Set 待转出 = n_批次
  Where 病人来源 = 2 And (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  --自登记类病人(无挂号单号)
  --病历ID可能重复是因为检验报告之类的，如肝功、肾功共打一张报告，即在病人医嘱报告表中，多个医嘱id对应同一报告ID
  --为提升性能，不从医嘱发送记录的发送时间查询，不采用精确的时间，因为直接登记的检验医嘱，一般开嘱时间与发送时间相差不大
  --有些特殊（错误）数据，挂号单为空的医嘱，除了来源为3的（直接登记的检查检验医嘱），还可能有来源为1或4的（门诊或体检医嘱），主页ID可能不是0
  Update /*+ rule*/ 电子病历记录
  Set 待转出 = n_批次
  Where 待转出 Is Null And 病历种类 = 7 And ID In (Select c.病历id
               From 病人医嘱记录 B, 病人医嘱报告 C
               Where c.医嘱id = b.Id And b.病人来源<>2 And b.挂号单 Is Null And b.相关id Is Null And b.待转出 Is Null And
                     b.开嘱时间 < d_End);

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
  Where 病历id In (Select ID From 电子病历记录 Where 病历种类 = 7 And 待转出 = n_批次);
  Update /*+ rule*/ 影像报告驳回
  Set 待转出 = n_批次
  Where (医嘱id, 病历id) In (Select 医嘱id, 病历id From 病人医嘱报告 Where 待转出 = n_批次);
  Update /*+ rule*/ 医嘱报告内容
  Set 待转出 = n_批次
  Where ID In (Select 报告id From 病人医嘱报告 Where 待转出 = n_批次);

  Update /*+ rule*/ 报告查阅记录
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 病历种类 = 7 And 待转出 = n_批次);

  Update /*+ rule*/ 疾病申报记录
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 病历种类 = 5 And 待转出 = n_批次);

  Update /*+ rule*/ 疾病报告反馈
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 病历种类 = 5 And 待转出 = n_批次);

  Update /*+ rule*/ 疾病申报反馈
  Set 待转出 = n_批次
  Where 申报id In (Select ID From 电子病历记录 Where 病历种类 = 5 And 待转出 = n_批次);

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
  --加上病人来源，避免来源为3的自登记类病人误填了挂号单后，医嘱被转走了而医嘱报告没有转出
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where 挂号单 In (Select NO From 病人挂号记录 Where 待转出 = n_批次) And 病人来源 =1;

  --加上病人来源，避免 后，医嘱被转走了而医嘱报告没有转出
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次) And 病人来源 = 2;

  --自登记类病人(无挂号单)，病人医嘱报告在前面转病历时已转出
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where Rowid In (Select b.Rowid
                  From 病人医嘱记录 B, 病人医嘱报告 C
                  Where (b.相关id = c.医嘱id Or b.Id = c.医嘱id) And c.待转出 = n_批次);

  --自登记类病人(无挂号单)，没有医嘱报告
  Update /*+ rule*/ 病人医嘱记录 A
  Set 待转出 = n_批次
  Where Not Exists (Select 1 From 病人医嘱报告 B Where a.Id = b.医嘱id) And Not Exists
     (Select 1 From 病人医嘱报告 B Where a.相关id = b.医嘱id) And 挂号单 Is Null And 病人来源 = 3 And 待转出 Is Null And 开嘱时间 < d_End;

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
  Update /*+ rule*/ 输血申请项目
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

  Update /*+ rule*/ 处方审查明细
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 处方审查记录
  Set 待转出 = n_批次
  Where ID In (Select 审方id From 处方审查明细 Where 待转出 = n_批次);

  Update /*+ rule*/ 处方审查结果
  Set 待转出 = n_批次
  Where 审方id In (Select ID From 处方审查记录 Where 待转出 = n_批次);

  Update /*+ rule*/ Ris检查预约
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 疾病阳性记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 医嘱申请单文件
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  
  Update /*+ rule*/ 病人危急值记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人危急值病历
  Set 待转出 = n_批次
  Where 危急值id In (Select ID From 病人危急值记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人危急值医嘱
  Set 待转出 = n_批次
  Where 危急值id In (Select ID From 病人危急值记录 Where 待转出 = n_批次);

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

  Update /*+ rule*/ 病人用药清单 
  Set 待转出 = n_批次 
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次); 
 
  Update /*+ rule*/ 病人用药配方
  Set 待转出 = n_批次 
  Where 配方id In (Select ID From 病人用药清单 Where 待转出 = n_批次); 

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/

--105711:李业庆,2019-01-11,出库库存平均成本价处理
Create Or Replace Procedure Zl_材料外购_Verify
(
  No_In       In 药品收发记录.No%Type := Null,
  审核人_In   In 药品收发记录.审核人%Type := Null,
  审核日期_In In 药品收发记录.审核日期%Type := Sysdate
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_单位id      药品收发记录.供药单位id%Type;
  n_发票金额    应付记录.发票金额%Type;
  n_库存金额    药品库存.实际金额%Type;
  n_库存差价    药品库存.实际差价%Type;
  n_库存数量    药品库存.实际数量%Type;
  n_实价卫材    收费项目目录.是否变价%Type;
  n_Batch_Count Integer; --原不分批现在分批的材料的数量
  v_条码前缀    Varchar2(20);
  v_内部条码    药品库存.内部条码%Type;
  v_移库no      药品收发记录.No%Type;
  v_对方库房id  药品收发记录.库房id%Type := 0;
  v_入类别id    药品收发记录.入出类别id%Type := 0;
  v_出类别id    药品收发记录.入出类别id%Type := 0;
  n_平均成本价  药品库存.平均成本价%Type;
  n_可用数量    药品收发记录.实际数量%Type;
Begin
  v_条码前缀 := Nvl(zl_GetSysParameter(159), '');

  Update 药品收发记录
  Set 审核人 = Nvl(审核人_In, 审核人), 审核日期 = 审核日期_In
  Where NO = No_In And 单据 = 15 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人审核或删除，不能进行审核！[ZLSOFT]';
    Raise Err_Item;
  End If;
  --主要针对原不分批现在分批的材料，不能对其审核
  Select Count(*)
  Into n_Batch_Count
  From 药品收发记录 A, 材料特性 B
  Where a.药品id = b.材料id And a.No = No_In And a.单据 = 15 And a.记录状态 = 1 And Nvl(a.批次, 0) = 0 And
        ((Nvl(b.库房分批, 0) = 1 And
        a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '发料部门') Or (工作性质 Like '制剂室'))) Or Nvl(b.在用分批, 0) = 1);

  If n_Batch_Count > 0 Then
    v_Err_Msg := '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能审核！[ZLSOFT]';
    Raise Err_Item;
  End If;

  --原分批现不分批的材料,在审核时，要处理他
  Update 药品收发记录
  Set 批次 = 0
  Where ID In
        (Select ID
         From 药品收发记录 A, 材料特性 B
         Where a.药品id = b.材料id And a.No = No_In And a.单据 = 15 And a.记录状态 = 1 And Nvl(a.批次, 0) > 0 And
               (Nvl(b.库房分批, 0) = 0 Or
               (Nvl(b.在用分批, 0) = 0 And
               a.库房id In (Select 部门id From 部门性质说明 Where (工作性质 Like '发料部门') Or (工作性质 Like '制剂室')))));

  For v_收发 In (Select a.Id, a.实际数量, a.发药方式, a.零售价, a.零售金额, a.差价, a.库房id, a.药品id, a.批次, a.供药单位id, a.成本价, a.批号, a.效期,
                      a.灭菌效期, a.生产日期, a.产地, a.入出类别id, a.注册证号, a.扣率, a.商品条码, a.内部条码, Nvl(b.是否条码管理, 0) As 条码管理, a.批准文号,
                      Nvl(a.费用id, 0) As 费用id, 序号, a.入出系数
               From 药品收发记录 A, 材料特性 B
               Where a.药品id = b.材料id And a.No = No_In And a.单据 = 15 And a.记录状态 = 1
               Order By a.药品id, a.批次) Loop
    v_内部条码 := Null;
    If v_收发.条码管理 = 1 Then
      If v_收发.内部条码 Is Null Then
        If Not v_条码前缀 Is Null Then
          v_内部条码 := v_条码前缀 || Nextno(126);
        Else
          v_内部条码 := Nextno(126);
        End If;
      Else
        v_内部条码 := v_收发.内部条码;
      End If;
      --处理条码打印管理数据
      Insert Into 卫材条码打印记录
        (NO, 单据, 库房id, 材料id, 序号, 商品条码, 内部条码, 入库数量, 打印数量, 入库时间)
      Values
        (No_In, 15, v_收发.库房id, v_收发.药品id, v_收发.序号, v_收发.商品条码, v_内部条码, v_收发.实际数量, 0, 审核日期_In);
    End If;
  
    --更改材料库存表的相应数据
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = v_收发.药品id;
  
    If v_收发.费用id = 2 Then
      n_可用数量 := Nvl(v_收发.实际数量, 0);
    Else
      If v_收发.发药方式 = 1 Then
        n_可用数量 := 0;
      Else
        n_可用数量 := Nvl(v_收发.实际数量, 0);
      End If;
    End If;
  
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + n_可用数量, 实际数量 = Nvl(实际数量, 0) + Nvl(v_收发.实际数量, 0), 实际金额 = Nvl(实际金额, 0) + Nvl(v_收发.零售金额, 0),
        实际差价 = Nvl(实际差价, 0) + Nvl(v_收发.差价, 0), 上次供应商id = Nvl(v_收发.供药单位id, 上次供应商id), 上次采购价 = Nvl(v_收发.成本价, 上次采购价),
        上次批号 = Nvl(v_收发.批号, 上次批号), 上次产地 = Nvl(v_收发.产地, 上次产地), 灭菌效期 = Nvl(v_收发.灭菌效期, 灭菌效期),
        上次生产日期 = Nvl(v_收发.生产日期, 上次生产日期), 效期 = Nvl(v_收发.效期, 效期),
        零售价 = Decode(Nvl(v_收发.批次, 0), 0, Null, Decode(n_实价卫材, 1, v_收发.零售价, Null)), 上次扣率 = Nvl(v_收发.扣率, 上次扣率),
        商品条码 = v_收发.商品条码, 内部条码 = v_内部条码, 批准文号 = v_收发.批准文号
    Where 库房id = v_收发.库房id And 药品id = v_收发.药品id And Nvl(批次, 0) = Nvl(v_收发.批次, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 灭菌效期, 效期, 零售价, 上次扣率, 商品条码,
         内部条码, 平均成本价, 批准文号)
      Values
        (v_收发.库房id, v_收发.药品id, v_收发.批次, 1, n_可用数量, v_收发.实际数量, v_收发.零售金额, v_收发.差价, v_收发.供药单位id, v_收发.成本价, v_收发.批号,
         v_收发.生产日期, v_收发.产地, v_收发.灭菌效期, v_收发.效期, Decode(Nvl(v_收发.批次, 0), 0, Null, Decode(n_实价卫材, 1, v_收发.零售价, Null)),
         v_收发.扣率, v_收发.商品条码, v_内部条码, v_收发.成本价, v_收发.批准文号);
    End If;
  
    If v_收发.内部条码 Is Null And Not v_内部条码 Is Null Then
      Update 药品收发记录 Set 内部条码 = v_内部条码 Where ID = v_收发.Id;
    End If;
  
    --清除数量金额为零的记录
    Delete From 药品库存
    Where 库房id = v_收发.库房id And 药品id = v_收发.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  
    --更改材料收发汇总表的相应数据
    --更新该材料的成本价
    Begin
      Select Sum(Nvl(实际金额, 0)), Sum(Nvl(实际差价, 0)), Sum(Nvl(实际数量, 0))
      Into n_库存金额, n_库存差价, n_库存数量
      From 药品库存
      Where 性质 = 1 And 药品id = v_收发.药品id;
    Exception
      When Others Then
        n_库存数量 := 0;
    End;
  
    --更新该药品的成本价
    Update 材料特性
    Set 成本价 = v_收发.成本价, 上次售价 = v_收发.零售价, 上次供应商id = v_收发.供药单位id, 上次产地 = v_收发.产地
    Where 材料id = v_收发.药品id;
  
    --更改材料特性中的注册证号:如果发现材料特性表中的注册证号没填，则直接反写给材料特性表中的注册证号
    If Nvl(v_收发.注册证号, ' ') <> ' ' Then
      Update 材料特性 Set 注册证号 = v_收发.注册证号 Where 材料id = v_收发.药品id And 注册证号 Is Null;
    End If;
  
    --不分批入库时才重新计算库存表中的平均成本价
    If Nvl(v_收发.批次, 0) = 0 And v_收发.入出系数 * Nvl(v_收发.实际数量, 0) > 0 Then
      Update 药品库存
      Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
      Where 药品id = v_收发.药品id And Nvl(批次, 0) = Nvl(v_收发.批次, 0) And 库房id = v_收发.库房id And 性质 = 1 And Nvl(实际数量, 0) <> 0;
      If Sql%NotFound Then
        Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = v_收发.药品id;
        Update 药品库存
        Set 平均成本价 = n_平均成本价
        Where 药品id = v_收发.药品id And 库房id = v_收发.库房id And Nvl(批次, 0) = Nvl(v_收发.批次, 0) And 性质 = 1;
      End If;
    End If;
  End Loop;

  --对应付余额表进行处理
  --此处用一个块，主要是解决没有对应发票号的记录
  Begin
    Update 应付记录
    Set 审核人 = 审核人_In, 审核日期 = 审核日期_In
    Where 入库单据号 = No_In And 系统标识 = 5 And 记录性质 = 0 And 记录状态 = 1;
  
    Select b.单位id, Sum(发票金额)
    Into n_单位id, n_发票金额
    From 药品收发记录 A, 应付记录 B
    Where a.Id = b.收发id And a.No = No_In And a.单据 = 15 And b.系统标识 = 5
    Group By b.单位id;
  
    If Nvl(n_单位id, 0) <> 0 Then
      Update 应付余额 Set 金额 = Nvl(金额, 0) + Nvl(n_发票金额, 0) Where 单位id = n_单位id And 性质 = 1;
    
      If Sql%NotFound Then
        Insert Into 应付余额 (单位id, 性质, 金额) Values (n_单位id, 1, n_发票金额);
      End If;
    End If;
  Exception
    When No_Data_Found Then
      Null;
  End;

  --如果是自动产生的备货卫材入库单，则产生移库单
  For v_Data In (Select ID, 序号, 实际数量, 发药方式, 零售价, 零售金额, 差价, 库房id, 药品id, 批次, 供药单位id, 成本价, 成本金额, 批号, 效期, 灭菌效期, 生产日期, 产地,
                        入出类别id, 注册证号, 扣率, 摘要, 商品条码, 内部条码, 费用id, 审核人, 审核日期
                 From 药品收发记录
                 Where NO = No_In And 单据 = 15 And 记录状态 = 1 And 审核日期 Is Not Null And 费用id > 0
                 Order By 药品id, 批次, 序号) Loop
    If v_对方库房id = 0 Then
      Begin
        Select Distinct 库房id Into v_对方库房id From 药品收发记录 Where 单据 In (24, 25) And 费用id = v_Data.费用id;
      Exception
        When Others Then
          v_对方库房id := 0;
      End;
    End If;
  
    If v_对方库房id > 0 Then
      If v_移库no Is Null Then
        v_移库no := Nextno(72, v_Data.库房id);
      End If;
    
      Zl_材料移库_Insert(v_移库no, v_Data.序号 * 2 - 1, v_Data.库房id, v_对方库房id, v_Data.药品id, v_Data.批次, v_Data.实际数量, v_Data.实际数量,
                     v_Data.成本价, v_Data.成本金额, v_Data.零售价, v_Data.零售金额, v_Data.差价, v_Data.审核人, v_Data.产地, v_Data.批号,
                     v_Data.效期, v_Data.灭菌效期, v_Data.摘要, v_Data.审核日期);
    End If;
  End Loop;

  --对新产生的移库单进行备料和审核
  If Not v_移库no Is Null Then
    Zl_材料移库_Prepare(v_移库no, 审核人_In);
    Zl_材料移库_Prepare(v_移库no);
  
    Select b.Id As 类别id
    Into v_入类别id
    From 药品单据性质 A, 药品入出类别 B
    Where a.类别id = b.Id And a.单据 = 34 And 系数 = 1 And Rownum < 2;
  
    Select b.Id As 类别id
    Into v_出类别id
    From 药品单据性质 A, 药品入出类别 B
    Where a.类别id = b.Id And a.单据 = 34 And 系数 = -1 And Rownum < 2;
  
    For v_Data In (Select 序号, 库房id, 对方部门id, 药品id, 产地, Nvl(批次, 0) As 批次, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, NO, 填制人, 批号,
                          效期, 灭菌效期, 填制日期
                   From 药品收发记录
                   Where 单据 = 19 And NO = v_移库no And 审核日期 Is Null And 入出系数 = -1
                   Order By 药品id, 批次, 序号) Loop
    
      Zl_材料移库_Verify(v_Data.序号, v_Data.库房id, v_Data.对方部门id, v_Data.药品id, v_Data.产地, v_Data.批次, v_Data.填写数量, v_Data.实际数量,
                     v_Data.成本价, v_Data.成本金额, v_Data.零售金额, v_Data.差价, v_出类别id, v_入类别id, v_Data.No, v_Data.填制人, v_Data.批号,
                     v_Data.效期, v_Data.灭菌效期, v_Data.填制日期, 1, v_Data.零售价);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料外购_Verify;
/

--105711:李业庆,2019-01-11,出库库存平均成本价处理
Create Or Replace Procedure Zl_材料其他入库_Verify
(
  No_In     In 药品收发记录.No%Type := Null,
  审核人_In In 药品收发记录.审核人%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_实价卫材    收费项目目录.是否变价%Type;
  n_Batch_Count Integer; --原不分批现在分批的材料的数量
  v_条码前缀    Varchar2(20);
  v_内部条码    药品库存.内部条码%Type;
  n_平均成本价  药品库存.平均成本价%Type;
Begin
  v_条码前缀 := Nvl(zl_GetSysParameter(159), '');

  Update 药品收发记录
  Set 审核人 = 审核人_In, 审核日期 = Sysdate
  Where NO = No_In And 单据 = 17 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;

  --主要针对原不分批现在分批的材料，不能对其审核
  Select Count(*)
  Into n_Batch_Count
  From 药品收发记录 A, 材料特性 B
  Where a.药品id = b.材料id And a.No = No_In And a.单据 = 17 And a.记录状态 = 1 And Nvl(a.批次, 0) = 0 And
        ((Nvl(b.库房分批, 0) = 1 And
        a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '发料部门') Or (工作性质 Like '制剂室'))) Or Nvl(b.在用分批, 0) = 1);

  If n_Batch_Count > 0 Then
    v_Err_Msg := '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能审核！[ZLSOFT';
    Raise Err_Item;
  End If;

  --原分批现不分批的材料,在审核时，要处理他
  Update 药品收发记录
  Set 批次 = 0
  Where ID =
        (Select ID
         From 药品收发记录 A, 材料特性 B
         Where b.材料id = a.药品id And a.No = No_In And a.单据 = 17 And a.记录状态 = 1 And Nvl(a.批次, 0) > 0 And
               (Nvl(b.库房分批, 0) = 0 Or
               (Nvl(b.在用分批, 0) = 0 And
               a.库房id In (Select 部门id From 部门性质说明 Where (工作性质 Like '发料部门') Or (工作性质 Like '制剂室')))));

  For c_收发 In (Select a.Id, a.实际数量, a.零售价, a.零售金额, a.差价, a.库房id, a.药品id, a.批次, a.成本价, a.批号, a.效期, a.灭菌效期, a.灭菌日期, a.产地,
                      a.入出类别id, a.生产日期, a.商品条码, a.内部条码, Nvl(b.是否条码管理, 0) As 条码管理, a.批准文号
               From 药品收发记录 A, 材料特性 B
               Where a.药品id = b.材料id And a.No = No_In And a.单据 = 17 And a.记录状态 = 1
               Order By a.药品id, a.批次) Loop
  
    v_内部条码 := Null;
    If c_收发.条码管理 = 1 Then
      If Not v_条码前缀 Is Null Then
        v_内部条码 := v_条码前缀 || Nextno(126);
      Else
        v_内部条码 := Nextno(126);
      End If;
    End If;
  
    --更改药品库存表的相应数据
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = c_收发.药品id;
  
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + Nvl(c_收发.实际数量, 0), 实际数量 = Nvl(实际数量, 0) + Nvl(c_收发.实际数量, 0),
        实际金额 = Nvl(实际金额, 0) + Nvl(c_收发.零售金额, 0), 实际差价 = Nvl(实际差价, 0) + Nvl(c_收发.差价, 0), 上次采购价 = Nvl(c_收发.成本价, 上次采购价),
        上次批号 = Nvl(c_收发.批号, 上次批号), 上次生产日期 = Nvl(c_收发.生产日期, 上次生产日期), 上次产地 = Nvl(c_收发.产地, 上次产地), 效期 = Nvl(c_收发.效期, 效期),
        灭菌效期 = Nvl(c_收发.灭菌效期, 灭菌效期), 零售价 = Decode(Nvl(c_收发.批次, 0), 0, Null, Decode(n_实价卫材, 1, c_收发.零售价, Null)),
        商品条码 = c_收发.商品条码, 内部条码 = v_内部条码, 批准文号 = c_收发.批准文号
    Where 库房id = c_收发.库房id And 药品id = c_收发.药品id And Nvl(批次, 0) = Nvl(c_收发.批次, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次采购价, 上次批号, 上次生产日期, 上次产地, 效期, 灭菌效期, 零售价, 商品条码, 内部条码, 平均成本价, 批准文号)
      Values
        (c_收发.库房id, c_收发.药品id, c_收发.批次, 1, c_收发.实际数量, c_收发.实际数量, c_收发.零售金额, c_收发.差价, c_收发.成本价, c_收发.批号, c_收发.生产日期,
         c_收发.产地, c_收发.效期, c_收发.灭菌效期, Decode(Nvl(c_收发.批次, 0), 0, Null, Decode(n_实价卫材, 1, c_收发.零售价, Null)), c_收发.商品条码,
         v_内部条码, c_收发.成本价, c_收发.批准文号);
    End If;
  
    Delete From 药品库存
    Where 库房id = c_收发.库房id And 药品id = c_收发.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  
    If Not v_内部条码 Is Null Then
      Update 药品收发记录 Set 内部条码 = v_内部条码 Where ID = c_收发.Id;
    End If;
  
    --更新该材料的成本价
    Update 材料特性 Set 成本价 = c_收发.成本价, 上次售价 = c_收发.零售价 Where 材料id = c_收发.药品id;
  
    --不分批入库时才重新计算库存表中的平均成本价
    If Nvl(c_收发.批次, 0) = 0 Then
      Update 药品库存
      Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
      Where 药品id = c_收发.药品id And Nvl(批次, 0) = Nvl(c_收发.批次, 0) And 库房id = c_收发.库房id And Nvl(实际数量, 0) <> 0 And 性质 = 1;
      If Sql%NotFound Then
        Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = c_收发.药品id;
        Update 药品库存
        Set 平均成本价 = n_平均成本价
        Where 药品id = c_收发.药品id And 库房id = c_收发.库房id And Nvl(批次, 0) = Nvl(c_收发.批次, 0) And 性质 = 1;
      End If;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料其他入库_Verify;
/

--105711:李业庆,2019-01-11,出库库存平均成本价处理
Create Or Replace Procedure Zl_材料其他出库_Verify
(
  序号_In       In 药品收发记录.序号%Type,
  No_In         In 药品收发记录.No%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  批次_In       In 药品收发记录.批次%Type,
  实际数量_In   In 药品收发记录.实际数量%Type,
  成本价_In     In 药品收发记录.成本价%Type,
  成本金额_In   In 药品收发记录.成本金额%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  差价_In       In 药品收发记录.差价%Type,
  入出类别id_In In 药品收发记录.入出类别id%Type,
  审核人_In     In 药品收发记录.审核人%Type,
  审核日期_In   In 药品收发记录.审核日期%Type
) Is
  Err_Item Exception;
  v_Err_Msg    Varchar2(100);
  v_负成本计算 Zlparameters.参数值%Type;
  v_批准文号   药品库存.批准文号%Type;
  v_批号       药品收发记录.批号%Type;
  v_产地       药品收发记录.产地%Type;

  n_实际库存金额 药品库存.实际金额%Type;
  n_实际库存差价 药品库存.实际差价%Type;
  n_实际库存数量 药品库存.实际数量%Type;

  n_出库差价 药品库存.实际差价%Type;
  n_成本价   药品收发记录.成本价%Type;
  n_成本金额 药品收发记录.成本金额%Type;

  n_上次供应商id 药品库存.上次供应商id%Type;
  n_实价卫材     收费项目目录.是否变价%Type;
  n_零售价       药品收发记录.零售价%Type;
  n_小数         Number(2);

  d_上次生产日期 药品库存.上次生产日期%Type;
  d_效期         药品库存.效期%Type;
  d_灭菌效期     药品库存.灭菌效期%Type;
  v_上次扣率     药品库存.上次扣率%Type;
  v_商品条码     药品库存.商品条码%Type;
  v_内部条码     药品库存.内部条码%Type;
  n_平均成本价   药品库存.平均成本价%Type;
Begin
  Select Nvl(精度, 2) Into n_小数 From 药品卫材精度 Where 类别 = 2 And 内容 = 4 And 单位 = 5;
  Select zl_GetSysParameter(120) Into v_负成本计算 From Dual;

  --由于领用处理允许在审核时改变实际数量，
  --所以首先对实际数量和其他相应的字段进行更新。
  Begin
    Select Nvl(实际金额, 0), Nvl(实际差价, 0), Nvl(实际数量, 0), Nvl(上次扣率, 100), 商品条码, 内部条码
    Into n_实际库存金额, n_实际库存差价, n_实际库存数量, v_上次扣率, v_商品条码, v_内部条码
    From 药品库存
    Where 药品id = 材料id_In And Nvl(批次, 0) = 批次_In And 库房id = 库房id_In And 性质 = 1 And Rownum = 1;
  Exception
    When Others Then
      n_实际库存金额 := 0;
      n_实际库存数量 := 0;
      v_上次扣率     := 100;
      v_商品条码     := Null;
      v_内部条码     := Null;
  End;

  If 成本价_In Is Null Then
    --成本价为空
    n_成本价   := Round(Zl_Fun_Getoutcost(材料id_In, 批次_In, 库房id_In), 7);
    n_成本金额 := Round(n_成本价 * 实际数量_In, n_小数);
    n_出库差价 := Round(零售金额_In - n_成本金额, n_小数);
  Else
    n_成本价   := 成本价_In;
    n_成本金额 := 成本金额_In;
    n_出库差价 := 差价_In;
  End If;

  --取上次供应商ID
  Begin
    Select 供药单位id, 零售价, 生产日期, 批准文号, 效期, 灭菌效期, 批号, 产地
    Into n_上次供应商id, n_零售价, d_上次生产日期, v_批准文号, d_效期, d_灭菌效期, v_批号, v_产地
    From 药品收发记录
    Where NO = No_In And 单据 = 21 And 药品id = 材料id_In And 记录状态 = 1 And 序号 = 序号_In And Rownum = 1;
  
  Exception
    When Others Then
      n_上次供应商id := Null;
  End;

  Update 药品收发记录
  Set 审核人 = Nvl(审核人_In, 审核人), 审核日期 = 审核日期_In, 成本价 = n_成本价, 成本金额 = n_成本金额, 差价 = n_出库差价, 扣率 = v_上次扣率, 商品条码 = v_商品条码,
      内部条码 = v_内部条码
  Where NO = No_In And 单据 = 21 And 药品id = 材料id_In And 序号 = 序号_In And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;
  Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;

  --更改药品库存的相应数据
  Update 药品库存
  Set 实际数量 = Nvl(实际数量, 0) - 实际数量_In, 实际金额 = Nvl(实际金额, 0) - 零售金额_In, 实际差价 = Nvl(实际差价, 0) - n_出库差价,
      零售价 = Decode(n_实价卫材, 1, Decode(Nvl(批次_In, 0), 0, Null, Decode(Nvl(零售价, 0), 0, n_零售价, 零售价)), Null), 商品条码 = v_商品条码,
      内部条码 = v_内部条码
  Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 性质 = 1;

  If Sql%NotFound Then
    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价, 上次扣率, 商品条码,
       内部条码, 平均成本价)
    Values
      (库房id_In, 材料id_In, 批次_In, 1, -实际数量_In, -实际数量_In, -零售金额_In, -n_出库差价, d_效期, d_灭菌效期, n_上次供应商id, n_成本价, v_批号,
       d_上次生产日期, v_产地, v_批准文号, Decode(n_实价卫材, 1, Decode(Nvl(批次_In, 0), 0, Null, n_零售价), Null), v_上次扣率, v_商品条码, v_内部条码,
       n_成本价);
  End If;

  Delete From 药品库存
  Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  --出库房，平均成本价为空时需要重新计算库存表中的平均成本价
  Update 药品库存
  Set 平均成本价 = n_成本价
  Where 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 库房id = 库房id_In And 性质 = 1 And Nvl(平均成本价, 0) = 0;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料其他出库_Verify;
/

--137357:刘涛,2019-01-21,过程取消可用库存检查
--105711:李业庆,2019-01-11,出库库存平均成本价处理
Create Or Replace Procedure Zl_材料领用_Verify
(
  序号_In       In 药品收发记录.序号%Type,
  No_In         In 药品收发记录.No%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  对方部门id_In In 药品收发记录.对方部门id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  产地_In       In 药品收发记录.产地%Type,
  批次_In       In 药品收发记录.批次%Type,
  填写数量_In   In 药品收发记录.填写数量%Type,
  实际数量_In   In 药品收发记录.实际数量%Type,
  成本价_In     In 药品收发记录.成本价%Type,
  成本金额_In   In 药品收发记录.成本金额%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  差价_In       In 药品收发记录.差价%Type,
  入出类别id_In In 药品收发记录.入出类别id%Type,
  审核人_In     In 药品收发记录.审核人%Type,
  审核日期_In   In 药品收发记录.审核日期%Type,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  审核方式_In   In Number := 0
) Is
  Err_Item Exception;
  v_Err_Msg    Varchar2(500);
  v_负成本计算 Zlparameters.参数值%Type;
  v_编码       收费项目目录.编码%Type;
  v_批准文号   药品库存.批准文号%Type;

  d_上次生产日期 药品库存.上次生产日期%Type;
  d_效期         药品库存.效期%Type;
  d_灭菌效期     药品库存.灭菌效期%Type;

  n_可用数量     药品库存.可用数量%Type;
  n_实际数量     药品库存.实际数量%Type;
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_零售价       药品收发记录.零售价%Type;
  n_实价卫材     收费项目目录.是否变价%Type;
  n_小数         Number(2);
  v_上次扣率     药品库存.上次扣率%Type;
  n_数量差       药品收发记录.实际数量%Type;
  v_商品条码     药品库存.商品条码%Type;
  v_内部条码     药品库存.内部条码%Type;
  n_平均成本价   药品库存.平均成本价%Type;
  v_下库存       Zlparameters.参数值%Type;
Begin
  Select zl_GetSysParameter(120) Into v_负成本计算 From Dual;
  Select Nvl(精度, 2) Into n_小数 From 药品卫材精度 Where 性质 = 0 And 类别 = 2 And 内容 = 4 And 单位 = 5;
  Select Nvl(zl_GetSysParameter(95), '0') Into v_下库存 From Dual;
  --由于领用处理允许在审核时改变实际数量，
  --所以首先对实际数量和其他相应的字段进行更新。

  Begin
    Select Nvl(可用数量, 0), Nvl(实际数量, 0), Nvl(上次扣率, 100), 商品条码, 内部条码
    Into n_可用数量, n_实际数量, v_上次扣率, v_商品条码, v_内部条码
    From 药品库存
    Where 药品id = 材料id_In And Nvl(批次, 0) = 批次_In And 库房id = 库房id_In And 性质 = 1 And Rownum = 1;
  
  Exception
    When Others Then
      n_可用数量 := 0;
      n_实际数量 := 0;
      v_上次扣率 := 100;
      v_商品条码 := Null;
      v_内部条码 := Null;
  End;

  --取上次供应商ID
  Begin
    Select 供药单位id, 生产日期, 批准文号, 效期, 灭菌效期, 零售价
    Into n_上次供应商id, d_上次生产日期, v_批准文号, d_效期, d_灭菌效期, n_零售价
    From 药品收发记录
    Where NO = No_In And 单据 = 20 And 药品id = 材料id_In And 记录状态 = 1 And 序号 = 序号_In And Rownum = 1;
  
  Exception
    When Others Then
      n_上次供应商id := Null;
      d_效期         := 效期_In;
  End;

  If 审核方式_In = 0 Then
    --出库审核
    Begin
      Select 实际数量 - 实际数量_In
      Into n_数量差
      From 药品收发记录
      Where NO = No_In And 单据 = 20 And 药品id = 材料id_In And 序号 = 序号_In And 记录状态 = 1 And 审核人 Is Null;
    Exception
      When Others Then
        n_数量差 := Null;
    End;
  
    If n_数量差 Is Null Then
      v_Err_Msg := '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Update 药品收发记录
    Set 审核人 = Nvl(审核人_In, 审核人), 审核日期 = 审核日期_In, 实际数量 = 实际数量_In, 成本价 = 成本价_In, 成本金额 = 成本金额_In, 零售金额 = 零售金额_In, 差价 = 差价_In,
        扣率 = v_上次扣率, 商品条码 = v_商品条码, 内部条码 = v_内部条码
    Where NO = No_In And 单据 = 20 And 药品id = 材料id_In And 序号 = 序号_In And 记录状态 = 1 And 审核人 Is Null;
  Elsif 审核方式_In = 1 Then
    --财务审核
    Begin
      Select 实际数量 - 实际数量_In
      Into n_数量差
      From 药品收发记录
      Where NO = No_In And 单据 = 20 And 药品id = 材料id_In And 序号 = 序号_In And 记录状态 = 1 And 配药人 Is Null;
    Exception
      When Others Then
        n_数量差 := Null;
    End;
  
    If n_数量差 Is Null Then
      v_Err_Msg := '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Update 药品收发记录
    Set 配药人 = Nvl(审核人_In, 配药人), 配药日期 = 审核日期_In, 实际数量 = 实际数量_In, 成本价 = 成本价_In, 成本金额 = 成本金额_In, 零售金额 = 零售金额_In, 差价 = 差价_In,
        扣率 = v_上次扣率
    Where NO = No_In And 单据 = 20 And 药品id = 材料id_In And 序号 = 序号_In And 记录状态 = 1 And 配药人 Is Null;
  End If;

  Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = 材料id_In;

  If 审核方式_In = 0 Then
    --审核
    If v_下库存 = 0 Then
      n_数量差 := -1 * 实际数量_In;
    End If;
  Else
    --核查
    If v_下库存 = 0 Then
      n_数量差 := 0;
    End If;
  End If;

  If 审核方式_In = 0 Then
    --出库审核处理实际库存
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + n_数量差, 实际数量 = Nvl(实际数量, 0) - 实际数量_In, 实际金额 = Nvl(实际金额, 0) - 零售金额_In,
        实际差价 = Nvl(实际差价, 0) - 差价_In,
        零售价 = Decode(n_实价卫材, 1, Decode(Nvl(批次_In, 0), 0, Null, Decode(Nvl(零售价, 0), 0, n_零售价, 零售价)), Null), 商品条码 = v_商品条码,
        内部条码 = v_内部条码
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价, 上次扣率,
         商品条码, 内部条码, 平均成本价)
      Values
        (库房id_In, 材料id_In, 批次_In, 1, -实际数量_In, -实际数量_In, -零售金额_In, -差价_In, d_效期, d_灭菌效期, n_上次供应商id, 成本价_In, 批号_In,
         d_上次生产日期, 产地_In, v_批准文号, Decode(n_实价卫材, 1, Decode(Nvl(批次_In, 0), 0, Null, n_零售价), Null), v_上次扣率, v_商品条码, v_内部条码,
         成本价_In);
    End If;
  
    Delete From 药品库存
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  
    --出库房，平均成本价为空时需要重新计算库存表中的平均成本价
    Update 药品库存
    Set 平均成本价 = 成本价_In
    Where 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 库房id = 库房id_In And 性质 = 1 And Nvl(平均成本价, 0) = 0;
  
  Elsif 审核方式_In = 1 Then
    --财务审核仅处理可用库存
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + n_数量差
    Where 库房id = 库房id_In And 药品id = 材料id_In And Nvl(批次, 0) = Nvl(批次_In, 0) And 性质 = 1;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料领用_Verify;
/

--105711:李业庆,2019-01-11,出库库存平均成本价处理
Create Or Replace Procedure Zl_材料盘点_Verify
(
  No_In     In 药品收发记录.No%Type := Null,
  审核人_In In 药品收发记录.审核人%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  n_成本价     药品收发记录.成本价%Type;
  n_材料id     药品收发记录.药品id%Type;
  n_实价卫材   收费项目目录.是否变价%Type;
  n_平均成本价 药品库存.平均成本价%Type;

  n_Batch_Count Integer; --原不分批现在分批的材料的数量

Begin
  Update 药品收发记录
  Set 审核人 = 审核人_In, 审核日期 = Sysdate
  Where NO = No_In And 单据 = 22 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;

  For c_单据 In (Select ID, 实际数量, 零售价, 零售金额, 差价, 库房id, 药品id 材料id, 批次, 批号, 效期, 灭菌效期, 产地, 入出类别id, 入出系数, 供药单位id, 生产日期, 批准文号
               From 药品收发记录
               Where NO = No_In And 单据 = 22 And 记录状态 = 1
               Order By 材料id, 批次) Loop
  
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = c_单据.材料id;
  
    If Nvl(c_单据.实际数量, 0) <> 0 Then
      n_成本价 := Round((Nvl(c_单据.零售金额, 0) - Nvl(c_单据.差价, 0)) / c_单据.实际数量, 7);
    Else
      n_成本价 := 0;
    End If;
  
    --更改药品库存表的相应数据
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + Decode(c_单据.入出系数, 1, Nvl(c_单据.实际数量, 0), 0),
        实际数量 = Nvl(实际数量, 0) + Nvl(c_单据.实际数量, 0) * c_单据.入出系数, 实际金额 = Nvl(实际金额, 0) + Nvl(c_单据.零售金额, 0) * c_单据.入出系数,
        实际差价 = Nvl(实际差价, 0) + Nvl(c_单据.差价, 0) * c_单据.入出系数,
        零售价 = Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_单据.零售价, 零售价)), Null),
        上次批号 = c_单据.批号, 上次产地 = c_单据.产地, 效期 = c_单据.效期
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价, 平均成本价)
      
      Values
        (c_单据.库房id, c_单据.材料id, c_单据.批次, 1, Decode(c_单据.入出系数, 1, Nvl(c_单据.实际数量, 0), 0), c_单据.实际数量 * c_单据.入出系数,
         c_单据.零售金额 * c_单据.入出系数, c_单据.差价 * c_单据.入出系数, c_单据.效期, c_单据.灭菌效期, c_单据.供药单位id, n_成本价, c_单据.批号, c_单据.生产日期, c_单据.产地,
         c_单据.批准文号, Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, c_单据.零售价), Null), n_成本价);
    End If;
  
    Delete From 药品库存
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  
    --平均价为空时重新计算平均成本价
    Update 药品库存
    Set 平均成本价 = n_成本价
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And 性质 = 1 And Nvl(平均成本价, 0) = 0;
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料盘点_Verify;
/

---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--129157:刘硕,2018-07-30,自动升级业务独立部件
Insert Into Zltools.Zlfilesupgrade
  (序号, 加入日期, 安装路径, 文件类型, 文件名, 版本号, 修改日期, 所属系统, 业务部件, Md5, 文件说明, 自动注册, 强制覆盖, 附加安装路径)
  Select 序号, To_Date('2018-07-30 11:35:31', 'yyyy-mm-dd hh24:mi:ss'), '[APPSOFT]', 0,
         'ZLHISCRUSTCOM.DLL', Null, Null, Null, Null, Null, '自动升级业务逻辑部件。缺失该部件在后续版本可能无法正常进行自动升级。', 1, 0, Null
  From Dual A, (Select NVL(Max(To_Number(序号)),0) + 1 序号 From zlFilesUpgrade) B
  Where Not Exists (Select 1 From Zltools.Zlfilesupgrade Where Upper(文件名) = 'ZLHISCRUSTCOM.DLL');
EXECUTE Zlfiles_Autoupdate('ZLSM4.DLL','7C9C89500C99326DE55FB06BE912332F',Null,To_Date('2018-12-26 10:20:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[SYSTEM]',Null,'0,1,3,4,6,21,22,23,24,25,26,28','国产加密算法部件。用于对系统中敏感信息进行加密。缺失该部件，将会导致包含密码的系统配置无法保存。',0,0,Null);
EXECUTE Zlfiles_Autoupdate('7Z.DLL','E3C7BC97672CDEB280DD43F2A69776BB','9.20.0.0',To_Date('2011-03-30 11:44:26', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),4,'[SYSTEM]',Null,'1','7-ZIP压缩程序',0,0,Null);
EXECUTE Zlfiles_Autoupdate('7Z.EXE','7083BA03D91F9D76CC659F973F14F839','9.20.0.0',To_Date('2011-03-30 11:44:26', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),4,'[SYSTEM]',Null,'1','7-ZIP压缩程序',0,0,Null);
EXECUTE Zlfiles_Autoupdate('AAMD532.DLL','CEFD956A1EF122CDA4D53007BAB6C694','1.0.0.1',To_Date('2011-09-27 11:15:40', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),4,'[SYSTEM]',Null,'1','部件功能:三方的MD5计算接口。使用原因:当不能使用VB进行MD5计算时，使用该三方部件进行Md5计算。缺失后果:常规无影响。但是自动升级检查该文件存在性，当不存在时不能自动升级。',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORGS.DLL','9DA8D145AE596EAF974AAC909B04115C','10.34.0',To_Date('2017-06-21 18:02:04', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:甘肃中医院专用体温单部件系统定位:处理老版护士工作站中体温单的相关业务缺失后果:缺失后会使用标准版体温部件',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORGX.DLL','B5C489885CF22547D59AC18EAD8C0513','10.32.0',To_Date('2017-06-21 18:02:16', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单(广西地区适用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准版体温部件',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORHEN.DLL','680CB972A8397819874DC052687EC7A7','10.33.0',To_Date('2017-06-21 18:02:40', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单相关功能(河南地区适用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORHN.DLL','AE646534CE848FA6F306CA2111DA26F4','10.32.0',To_Date('2017-06-21 18:01:54', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单相关功能(河南地区适用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORHUN.DLL','02F2305082A6710898C4B8DB1318EEBF','10.34.0',To_Date('2017-06-21 18:03:00', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单相关功能(湖南省通适用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORQD.DLL','4E943413771D91286A9E51AEFA11C2B9','10.32.0',To_Date('2017-06-21 18:02:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单相关功能(青岛地区适用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORSCDQ.DLL','ED2F8236A314F24863DE0B0815309FA7','10.34.0',To_Date('2017-06-21 18:01:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单相关功能(四川地区通用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORSXET.DLL','763529B90BA6E2A1C20851BABED21F52','10.34.0',To_Date('2017-06-21 18:02:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单相关功能(陕西西安儿童医院适用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORSXHZ.DLL','354FD75A6CA8B69E16A12C198E9A3FFE','10.32.0',To_Date('2017-06-21 18:01:42', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单相关功能(陕西省汉中市适用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORYDEY.DLL','94F57B95A1520DFB8CE78600A469D931','10.34.100',To_Date('2017-06-21 18:01:32', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单相关功能(医大二院适用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITORYX.DLL','C16426F08673D413E09D70077EFB7E39','10.32.0',To_Date('2017-06-21 18:01:12', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版护士工作站中体温单相关功能(云南省玉溪市人民医院适用)系统定位:处理老版护士工作站中体温单相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9CISAUDIT.DLL','B586DA484B36BBAF52A8D513970173CB','10.34.170',To_Date('2019-02-19 18:03:32', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'1','部件功能:电子病案审查归档系统定位:包含模块：病案评分标准、病案审查标准、电子病案审查、电子病案借阅、电子病案评分、病历质量查阅、电子病案接收；以电子病历质控为核心业务的应用模块集中于该部件中。缺失后果:质控相关模块窗口无法使用。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9CISBASE.DLL','A5EB7633B9510317FF292415338151F7','10.34.170',To_Date('2019-02-19 18:01:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISAudit.dll,zl9InExse.DLL,zl9LisWork.dll,zl9PACSWork.dll,zl9Oper.dll,zl9CISBase.dll,zl9Blood.dll,zl9CISJob.dll','1','部件功能:临床基础部件系统定位:设置药品目录，诊疗项目及相关，检查、检验、影像等相关基础数据缺失后果:无法设置药品目录，诊疗项目及相关，检查、检验、影像等相关基础数据',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9CISJOB.DLL','EA3C8C3866D90D6717589632E4603E47','10.34.170',To_Date('2019-02-19 18:00:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','Zl9DrugStore.dll','1','部件功能:临床工作站框架部件系统定位:住院医生站，住院护士站，新版护士站，老版护士站，老版医技站，电子病案查阅。缺失后果:以上工作站无法使用。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9CUSTACC.DLL','997AFB47C8E0345D845AF3BF4229AD31','10.34.170',To_Date('2019-02-19 17:48:54', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9BaseItem.dll,zl9InExse.DLL,zl9OutExse.dll','1','部件功能:自定义记帐单部件系统定位:专项记帐和专项记账单设置缺失后果:不能使用专项记帐单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9DISEASE.DLL','A2B3353CE0222202E54EB43D34218B2E','10.34.170',To_Date('2019-02-19 17:56:14', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:传染病管理系统相关功能系统定位:处理传染病管理系统相关业务缺失后果:无法使用传染病填写、上报相关功能',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9DRUGSTORE.DLL','236FBA91CB085A52A811A33EC0A8CDD7','10.34.170',To_Date('2019-02-19 17:50:04', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'1','部件功能:药房事务系统定位:门诊、住院药房发药管理，输液配置中心管理缺失后果:无法使用门诊、住院药房管理功能，无法使用静配功能',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9INEXSE.DLL','3E4EFF7130AC11C5325804EED904DDB5','10.34.170',To_Date('2019-02-19 17:54:36', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll,zl9InPatient.dll,zl9OutExse.dll','1','部件功能:住院费用部件系统定位:住院记帐、科室分散记帐、医技科室记帐、自动记帐计算、病人费用查询、费用审核、执行登记、病人结帐处理。缺失后果:住院费用业务不能运行。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9INPATIENT.DLL','B71E1917E9F248C00F02F0FCEEB8CA60','10.34.170',To_Date('2019-02-19 17:49:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll,zl9InExse.DLL','1','部件功能:住院病人部件系统定位:实现病人入院登记、病人入出管理、病区床位管理缺失后果:无法完成住院病人登记、病人入出管理、病区床位管理',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9LCDSHOW.DLL','EF69CD89F32120840A7145E31847F02B','10.34.0',To_Date('2014-10-30 22:33:20', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9Transfusion.dll','1','部件功能:老版排队显示业务封装系统定位:老版排队情况封装缺失后果:不能显示排队情况',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9LEDVOICE.DLL','EECBCF5F3DBB67C790FD52BB25A6D528','10.34.170',To_Date('2019-02-19 17:34:40', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9RegEvent.dll,zl9CardSquare.dll,zl9InExse.DLL,zl9OutExse.dll,zl9Patient.dll,zl9InPatient.dll','1','部件功能:LED显示、语言报价部件系统定位:向病人显示收费信息缺失后果:不能支持语言半价或者在LED屏显示信息',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9LISCOMM.EXE','497FDA89597D83E9622BC91725EA032D','10.34.170',To_Date('2019-02-19 18:00:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]','zl9LisWork.dll','1','部件功能:老版检验通讯程序系统定位:老版检验通讯程序，处理仪器回传数据，加工成检验系统能够认识的数据格式缺失后果:检验结果将不能正常回传给老版LIS系统',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9LISQUERY_BASE.DLL','5621235D774A20F9DE346488ABF5EE51','10.34.10',To_Date('2014-12-24 14:30:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9LisWork.dll','1','部件功能:检验外挂接口部件系统定位:支持检验外挂接口。缺失后果:综合查询外挂不能加载',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9LISQUERY_DFN.DLL','81DA77CF87E9F523B109FBDB6D0249A1','10.34.0',To_Date('2014-10-30 20:17:52', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9LisWork.dll','1','部件功能:老版LIS外挂部件系统定位:加载渠道开发的外挂部件缺失后果:不能正常加载外挂',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9LISWORK.DLL','38570B657EBFBDD8D0437A5581474B3A','10.34.170',To_Date('2019-02-19 18:01:34', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:老版LIS核心部件。系统定位:处理检验相关操作。包含检验技师工作站、检验采集工作站、检验登记。缺失后果:检验相关业务讲不能正常使用。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9MEDISTORE.DLL','14D0378357945CCE42659C0D32A4FA67','10.34.170',To_Date('2019-02-19 17:50:56', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'1','部件功能:zlMediStore系统定位:药品流通业务部件，如入库，出库，盘点等业务操作缺失后果:无法进行药品流通业务',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9OUTEXSE.DLL','2C6FBBF7E02020C53BCA0A2FF1D6F8BB','10.34.170',To_Date('2019-02-20 09:32:08', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','Zl9DrugStore.dll','1','部件功能:门诊费用部件系统定位:门诊划价、门诊收费、门诊记账缺失后果:缺失上述功能',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PACSCONTROL.OCX','9F12074151D3DE119C12DDBB088B104E','10.34.170',To_Date('2019-02-20 09:34:24', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9PACSWork.dll','1','部件功能:自定义控件封装系统定位:对常用的控件进行封装缺失后果:进入影像系统将产生异常',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PACSIMAGECAP.DLL','01FCB6E1AF5331CD26D15B246B971FDC','10.34.170',To_Date('2019-02-19 18:05:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9PACSWork.dll','1','部件功能:提供影像图像的采集与传输系统定位:影像检查图像采集支持缺失后果:影像采集病理系统不能运行。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PACSWORK.DLL','696D38BAA7E2CCADD6DC8A82A825729B','10.34.170',To_Date('2019-02-19 18:06:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'1','部件功能:进行影像系统基本业务处理系统定位:封装了对影像系统基本业务的处理，是业务系统的入库。缺失后果:不能进入对应的影像系统。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PARTOGRAM.DLL','23F91048B2D5A90825E17E947138D01E','10.34.170',To_Date('2019-02-19 18:07:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:产程图管理相关功能系统定位:处理产程图相关业务缺失后果:无法使用产程图展示,编辑,打印功能',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PATIENT.DLL','D820D59679F42BE8B614E72C12613136','10.34.170',To_Date('2019-02-20 09:31:38', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9RegEvent.dll,zl9CardSquare.dll,zl9InExse.DLL,zl9OutExse.dll,zl9InPatient.dll','1','部件功能:病人信息管理部件。系统定位:病人信息登记、修改、删除等操作。缺失后果:无法对病人信息进行维护。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9RECIPEAUDIT.DLL','50BE2893EE7D0A5F7BD8163DC383CC7A','10.34.170',To_Date('2019-02-19 18:14:42', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:药剂师对门诊和住院的用药处方审查系统定位:控制有问题的处方，提升处方合格率缺失后果:缺少将无法使用处方审查系统',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9REGEVENT.DLL','DD150852FCB41BDD1FA35A0D11907250','10.34.170',To_Date('2019-02-19 17:52:42', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll,zl9OutExse.dll','1','部件功能:门诊挂号部件系统定位:设置挂号安排和临床出诊，提供挂号、预约、分诊等功能，提供患者服务中心对病人预约进行管理缺失后果:与挂号相关的功能不能使用',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9STUFF.DLL','9A8A2B67B0F8B794A8C09A299CC7FDD8','10.34.170',To_Date('2019-02-19 17:52:48', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','Zl9DrugStore.dll','1','部件功能:zl9Stuff系统定位:卫材业务部件，包括卫材目录，卫材入出流通管理，卫材发放管理等缺失后果:不能开展卫材流通，发放等业务',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHART.DLL','951F1F11F540331B978948F3CCEDD50C','10.34.170',To_Date('2019-02-19 17:58:58', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(标准体温单,未启用地区性体温单时均使用此体温单)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:无法进行体温单的展示和数据编辑功能',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTGD.DLL','7036033E86A2B8318A47ADE901530FAD','10.34.170',To_Date('2019-02-19 18:12:12', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(广东省地区适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTGS.DLL','1276FEC265E85CC43A3C296705E77487','10.34.170',To_Date('2019-02-19 17:59:24', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(甘肃中医院适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTGX.DLL','E7145DC0D3B8434FB30BD5E7E18D3340','10.34.170',To_Date('2019-02-19 17:59:02', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(广西省适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTHNNX.DLL','453CF91CC82940009FCAC175AE9E0DD4','10.34.170',To_Date('2019-02-19 18:11:32', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(湖南宁乡人民医院适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTJX.DLL','4CF521EB7BDB76ADFF43F2BA22D58E0E','10.34.170',To_Date('2019-02-19 18:15:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(江西省适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTNJ.DLL','33205787668FE3BF3E966102C9CA0185','10.34.170',To_Date('2019-02-19 18:14:36', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(江苏地区适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTS3201.DLL','229A7AEA13D8A7281EBE9CF3BD53896C','10.34.170',To_Date('2019-02-19 17:59:12', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(陕西3201医院适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温部件',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTSC.DLL','83EDFAEE45A7898F8E36EDC1BF29CE76','10.34.170',To_Date('2019-02-19 17:47:46', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(四川地区通用体温部件)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTSCZG.DLL','2613C6D4431D88B388C4E0D2FAEBC5CE','10.34.170',To_Date('2019-02-19 17:48:04', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(四川自贡市适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTSX.DLL','87204AA48DE8BB2DE0565772A6218EBA','10.34.170',To_Date('2019-02-19 18:10:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(山西地区适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTYDEY.DLL','AE3E65467ABC070618A426F250B3F819','10.34.170',To_Date('2019-02-19 18:11:56', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(医大二院适用)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TEMPERATURECHARTYN.DLL','EFE500FEE48CFEA083A4166744D85218','10.34.170',To_Date('2019-02-19 18:14:00', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:新版护士工作站体温单相关功能(云南大理专用体温部件)系统定位:处理新版护士工作站中体温单的相关业务缺失后果:缺失会自动使用标准体温单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TRANSFUSION.DLL','02FAEE3B3AFC51DAE633F661570360A4','10.34.170',To_Date('2019-02-19 18:02:44', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'1','部件功能:门诊输液系统部件系统定位:处理门诊输液执行、附费相关业务缺失后果:无法使用门诊输液工作站',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9XWINTERFACE.DLL','D3375C95ECA263EC3774A54EECE7AD98','10.34.170',To_Date('2019-02-19 18:14:54', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zlWizardStart.exe,zl9CISAudit.dll,zl9CISJob.dll,zl9PACSWork.dll,zl9Patient.dll,zl9InPatient.dll,zl9CISBase.dll,zl9BaseItem.dll','1','部件功能:提供ris与his系统数据交换接口系统定位:专业ris系统支持缺失后果:不能使用专业版ris系统',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLACTMAIN.EXE','569D92F5D99BAB6BF2FEFEF7D1259BD9','10.34.170',To_Date('2019-02-19 18:12:14', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[APPSOFT]',Null,'1','部件功能:BH融合中的虚拟导航台。系统定位:BH调用各个模块均通过该程序进行导航。缺失后果:缺失时BH无法使用所有的业务模块。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLCISAUDITPRINT.EXE','C08FCEEC633A12BEE62290EE690C8334','10.34.170',To_Date('2019-02-19 18:14:10', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISAudit.dll','1','部件功能:用于电子病案审查中,文件-输出到PDF系统定位:避免连续PDF输出引起系统GDI超量，导致系统假死缺失后果:无法进行PDF输出',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLCISPATH.DLL','95AA62404D5CF2DABABCEEB166268DC7','10.34.170',To_Date('2019-02-19 17:56:40', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISAudit.dll,zl9CISJob.dll','1','部件功能:临床路径部件系统定位:临床路径应用、临床路径管理、临床路径跟踪缺失后果:临床路径应用、临床路径管理、临床路径跟踪将无法正常运行',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLDISREPORTCARD.DLL','8F65F43655C1E174C84C3306E7784A29','10.34.60',To_Date('2016-01-20 10:12:16', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9Disease.dll','1','部件功能:传染病固定格式报告卡控件系统定位:用于填写传染病报告卡缺失后果:无法使用固定格式传染病报告卡',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLGETIMAGE.EXE','051EB1E7D561B439A7335D237D5813DE','10.34.10',To_Date('2014-12-24 14:31:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9PACSWork.dll','1','部件功能:提供影像检查图像下载支持系统定位:后台下载影像检查图像缺失后果:不能观片',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLKNOWLEDGECONVERT.DLL','F0E28A1770D514238E7591038382B64C','10.34.170',To_Date('2019-02-19 18:15:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[Appsoft]\Apply','zl9CISJob.dll','1','知识库中心药品说明书部件',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLLISINTERFACE.DLL','DE8B214E0B7816BBB786ABED773C9E2D','10.34.10',To_Date('2014-12-24 14:33:46', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9LisWork.dll','1','部件功能:三方检验接口部件系统定位:支持三方检验接入到HIS系统中缺失后果:三方检验不能正常接入到his系统，该部件未正常发布。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLLISRECEIVESEND.EXE','BC251F502907D92092AF4C857476060A','10.34.170',To_Date('2019-02-19 18:00:46', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]','zl9LisWork.dll','1','部件功能:主要与检验仪器直接通讯系统定位:记录仪器回传的检验结果，并保存文本为LIS认识的检验结果。缺失后果:如果缺失,老版LIS将不能正常解析检验数据',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLLOGIN.DLL','FE4A7A80B27D88B8C3338B553E5A4754','10.34.170',To_Date('2019-02-19 18:14:44', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1','部件功能:公共登录部件，提供统一登录接口，以及登陆中的功能、授权、客户端控制、升级等处理。系统定位:各个Exe均通过该部件实现登录。缺失后果:缺失该部件将会导致启动Exe程序出错。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLNEWQUERY.EXE','119796792562AD6C008E529DE55E93C8','10.34.170',To_Date('2019-02-19 17:51:04', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]',Null,'1','部件功能:老版自助系统系统定位:自助挂号、Lis打印、费用查询缺失后果:老版自助系统不能使用',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPACSFTPTOOLS.EXE','2110B63987A4072EEFE5C2DB88A45EF7','1.0.0',To_Date('2017-10-23 09:49:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]','zl9PACSWork.dll','1','部件功能:对FTP进行测试，排查FTP相关操作错误。',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPACSIMAGEVALID.EXE','F88A5B39D554678991A1D439DF0DA3A8','10.34.170',To_Date('2019-01-16 17:27:06', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[Appsoft]','zl9PACSWork.dll','1','对Pacs图像完整性进行后台检测',0,1,Null);
EXECUTE Zlfiles_Autoupdate('ZLPACSRICHPAGES.OCX','F722E6AD924344A82BC30041A4DD6B0B','1.3520.447',To_Date('2018-05-25 16:14:04', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zlPublicPacs.dll','1','部件功能:封装pacs智能文档编辑器相关处理。系统定位:使用Pacs智能文档编辑器编辑需要。缺失后果:不能使用pacs智能文档编辑器',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPACSSRV.EXE','C1ED65F32D4C4213C04FF03F02BC398B','10.34.170',To_Date('2019-02-19 18:03:06', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]','zl9CISJob.dll','1','部件功能:接受Dicom设备发送的检查图像系统定位:PACS网关服务，监听影像DICOM设备请求并进行处理缺失后果:不能与影像DICOM设备通讯，不能接受设备图像',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPACSVBCOMMON.DLL','4BCD7BFC6F184343C2867A94267FFF0F','10.34.170',To_Date('2019-02-19 18:05:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9PACSWork.dll','1','部件功能:封装zlpacs与pacs智能报告编辑器之间的数据交换系统定位:zlpacs整合pacs智能报告编辑器缺失后果:不能使用pacs智能报告编辑器',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPASSINTERFACE.DLL','51786D14A0369A40D2B65D57DEB23922','10.34.170',To_Date('2019-02-19 17:53:04', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9CISJob.dll','1','部件功能:合理用药监测部件：集成了美康、大通、太元通、药卫士等合理用药监测接口。系统定位:供临床医生工作站、药品发药组件调用。缺失后果:合理用药监测功能无法启用。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLQUEUEOPER.OCX','05212C62063BEE8A8B79405F845C4167','10.34.170',To_Date('2019-02-19 18:04:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9Oper.dll,zl9PACSWork.dll','1','部件功能:排队叫号业务封装系统定位:pacs排队管理支持缺失后果:不能进行排队操作',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLQUEUESHOW.EXE','A49E5768FBD39B1AAA5A265C36CC822D','10.34.170',To_Date('2019-02-19 18:14:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\ZLQUEUESHOW',Null,'1','部件功能:新版排队显示系统定位:pacs排队情况显示缺失后果:不能显示pacs排队状态',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLRISDUMPTOOL.EXE','4202E79D877C69D42C3018CD6ECADB3E','10.34.150',To_Date('2018-04-04 14:38:54', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]',Null,'1','部件功能:基础数据，用户，诊疗项目，数据字典等初始化系统定位:初始化ris接口数据缺失后果:不能使用ris系统。',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLSCREENKEYBOARD.EXE','E9CB37C2D76B5E77EF23B2669DAAE40C','10.34.0',To_Date('2014-10-30 22:47:20', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1','部件功能:屏幕键盘小程序系统定位:在门诊医生工作站中用到，强制续诊，门诊医嘱下达缺失后果:强制续诊，门诊医嘱下达不能用使用键盘功能。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLSOFTSHOWHISFORMS.EXE','8C6732D55E26F59F097181771F3AE873','10.34.170',To_Date('2019-02-19 18:15:24', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]','zl9XWInterface.dll','1','部件功能:显示病历查阅,医嘱，pacs历史报告等系统定位:ris中调用查看病历内容。缺失后果:ris系统不能查看病历内容。',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLSVGPROCESS.DLL','9E00926BA33C621864432B114DF8CE93','1.0.3',To_Date('2015-09-29 16:24:42', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9PACSWork.dll','1','部件功能:svg图像转换使用原因:影像PACS智能报告编辑器检查图像转换缺失后果:不能查看报告图像',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLSVRNOTICE.EXE','939DDD0F860DEF59BDDF5A7B23AB5EE0','10.34.0',To_Date('2014-10-30 20:13:20', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[APPSOFT]',Null,'1','部件功能:自动提醒服务。系统定位:进行消息提醒的提示与阅读。缺失后果:无发处理消息提醒。',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9CASHBILL.DLL','E4B94C268A1FFD89369A0568FC78430E','10.34.170',To_Date('2019-02-19 17:48:46', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9RegEvent.dll,zl9CardSquare.dll,zl9InExse.DLL,zl9OutExse.dll,zl9Patient.dll,zl9InPatient.dll,zl9CustAcc.dll','1,21','部件功能:财务监控及票据管理系统定位:收费轧帐、财务组收款、收费财务监控、人员借款和票据的入库、领用和报损。缺失后果:相关功能不能使用',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PACSCORE.DLL','8CC5134CB4B8077B4F6B5EEDAFD063F9','10.34.170',To_Date('2019-02-19 18:04:58', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9PACSWork.dll','1,21','部件功能:PACS观片处理部件系统定位:查看PACS图像缺失后果:不能进行影像观片',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLQUEUEMANAGE.DLL','FF05CBF8C593E059F59891447F394859','10.34.170',To_Date('2019-02-19 17:51:20', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9RegEvent.dll,zlWizardStart.exe,zl9CISJob.dll,zl9Transfusion.dll','1,21','部件功能:老板排队业务封装系统定位:老板排队支持缺失后果:不能进行排队管理',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9CARDSQUARE.DLL','9AB44219BCD056B55DD88A589DC4C2BF','10.34.170',To_Date('2019-02-19 17:52:36', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9RegEvent.dll,zl9Transfusion.dll,zl9Blood.dll,zl9CISAudit.dll,zl9PACSWork.dll,zl9OutExse.dll,zl9Oper.dll,zl9LisWork.dll,ZL9LabWork.dll,zl9InPatient.dll,zl9InExse.DLL,Zl9DrugStore.dll,zl9Patient.dll,zl9CISJob.dll,zl9CardSquare.dll,zl9XWInterface.dll,zl9Stuff.dll,zl9CISBase.dll','1,21,22,24','部件功能:结算卡管理部件系统定位:医疗卡、消费卡管理缺失后果:涉及一卡通的业务无法使用',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLCISKERNEL.DLL','779108A01C890006B304367A09B2E17B','10.34.170',To_Date('2019-02-19 17:55:16', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISAudit.dll,zl9InExse.DLL,zl9LisWork.dll,zl9PACSWork.dll,zl9Oper.dll,zl9CISBase.dll,zl9Blood.dll','1,21,22,24','部件功能:临床核心部件，提供医嘱相关操作封装等接口，提供DOCK页签等。系统定位:提供医嘱核功能缺失后果:临床医嘱相关功能丢失，各大工作站的医嘱信息页签丢失。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLDSVIDEOPROCESS.OCX','915B1465C101552783EF186341BDA99A','1.2.74',To_Date('2016-08-03 14:10:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9RegEvent.dll,zl9CardSquare.dll,ZL9Peis.dll,zl9PeisManage.dll,zl9PACSWork.dll','1,21,22,24,25,26','部件功能:视频采集相关功能封装使用原因:体检人员照片采集缺失后果:不能采集图像和录像',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9COMMEVENTS.DLL','475A07ED30357C66E63300B867A7298F','10.34.170',To_Date('2019-02-19 18:09:24', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,21,22,24,26','部件功能:公共的基础事件部件：自动发卡、投币、键盘输入和自动读卡事件触发后，主程序能够响应。系统定位:三方程序触发事件,以便主程序接收数据。缺失后果:自助发卡、自助系统的现金支付功能调用就会出现错误，所有使用第三方接口自动读卡也会出现错误。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9KEYBOARD.DLL','F5AB1944768EC7CA07F026CA4E63C730','10.34.170',To_Date('2019-02-19 17:34:30', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9RegEvent.dll,zl9CardSquare.dll,zl9Patient.dll,zl9InPatient.dll','1,21,22,24,26','部件功能:密码键盘部件系统定位:使用密码键盘设备缺失后果:密码键盘无法使用',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPICTUREEDITOR.DLL','B7939177CF682C15B532E745D2E8250A','10.34.0',To_Date('2014-10-30 20:18:48', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','ZL9Peis.dll,zl9PeisManage.dll,zlWizardStart.exe,zl9Blood.dll,zl9CISJob.dll,zl9PACSWork.dll,zl9Oper.dll,zl9LisWork.dll,zl9Disease.dll,zl9CISAudit.dll','1,22,24','部件功能:用于对图片进行压缩处理系统定位:在病历标记图管理、电子病历编辑、保存过程中对图片进行压缩处理，以便优化处理后存入数据库缺失后果:因固定引用，缺少将无法使用电子病历所有模块',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLRICHEDITOR.OCX','3B05C1D3038373164D394B1F0ACC383D','10.34.170',To_Date('2019-02-19 17:53:16', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','ZL9Peis.dll,zl9PeisManage.dll,zlWizardStart.exe,zl9Blood.dll,zl9CISJob.dll,zl9PACSWork.dll,zl9Oper.dll,zl9LisWork.dll,zl9Disease.dll,zl9CISAudit.dll,zl9Disease.dll','1,22,24','部件功能:病历编辑核心部件系统定位:提供病历编辑、打印输出功能缺失后果:无法编辑、打印病历',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLRICHEPR.DLL','2AB02E6A387C449291B086FEF2FCFC02','10.34.170',To_Date('2019-02-20 09:35:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zlWizardStart.exe,zl9Blood.dll,zl9CISJob.dll,zl9PACSWork.dll,zl9Oper.dll,zl9LisWork.dll,zl9Disease.dll,zl9CISAudit.dll,zl9Disease.dll','1,22,24','部件功能:病历编辑窗口及业务处理程序系统定位:提供模块：病历标记图形管理、护理记录项目管理、病历文件管理、病历范文管理、诊疗单据设置、移动护士站基础设置、病人病历检索、疾病申报管理缺失后果:病历相关业务无法开展，医生工作站因直接引用，将无法打开。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLSUBCLASS.OCX','AD1EED43473111FBB5635454BFD6BB88','10.34.0',To_Date('2014-10-30 20:18:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','ZL9Peis.dll,zl9PeisManage.dll,zlWizardStart.exe,zl9Blood.dll,zl9CISJob.dll,zl9PACSWork.dll,zl9Oper.dll,zl9LisWork.dll,zl9Disease.dll,zl9CISAudit.dll,zl9Disease.dll','1,22,24','部件功能:鼠标、键盘勾子系统定位:用于向病历编辑相关模块提供鼠标、键盘勾子，以便在病历编辑过程中对界面内容进行鼠标控制、控件原有快捷键屏蔽。缺失后果:无法编辑病历',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLTABLE.OCX','35EEA0A1B80A67193953C8B19E3BE76C','10.34.0',To_Date('2014-10-30 20:19:00', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','ZL9Peis.dll,zl9PeisManage.dll,zlWizardStart.exe,zl9Blood.dll,zl9CISJob.dll,zl9PACSWork.dll,zl9Oper.dll,zl9LisWork.dll,zl9Disease.dll,zl9CISAudit.dll,zl9Disease.dll','1,22,24','部件功能:向病历编辑过程中提供自定义内嵌表格支持系统定位:全文式病历编辑过程中，插入表格后编辑生成对应的表格图，以及后续编辑时再次进行编辑转换、以提供检查报告图组缺失后果:无法进行内嵌表格编辑',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLTABLEEPR.DLL','EB23852D9B3CF990843B425C795C821F','10.34.170',To_Date('2019-02-19 17:53:00', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zlWizardStart.exe,zl9Blood.dll,zl9CISJob.dll,zl9PACSWork.dll,zl9Oper.dll,zl9LisWork.dll,zl9Disease.dll,zl9CISAudit.dll,zl9Disease.dll','1,22,24','部件功能:表格式病历核心编辑器系统定位:用于以表格式病历进行编辑的主窗口、表格式病历的打印缺失后果:无法进行表格式病历编辑',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BODYEDITOR.DLL','336A9D81E3F6832D5EC6AC9DF3BD9DF5','10.34.170',To_Date('2019-02-19 18:06:12', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISJob.dll','1,24','部件功能:老版护士工作站中标准体温单相关功能调用,在未使用地区性体温部件时,均使用此部件系统定位:处理老版护士工作站中标准体温单相关业务缺失后果:无法使用老版护士工作站体温单的数据展示,编辑和打印功能',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9TENDFILE.DLL','50D5E8FC794A23494EE433FA41239D3E','10.34.170',To_Date('2019-02-19 17:59:44', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9CISAudit.dll,zl9CISJob.dll,zl9Oper.dll','1,24','部件功能:处理护士工作站记录单相关业务系统定位:护士工作站中记录单相关业务流程缺失后果:无法在护士工作站中进行记录单的查看和操作',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLMEDRECPAGE.DLL','14CEAA5D2BED82432FC97A0880D4A122','10.34.170',To_Date('2019-02-19 17:36:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9CISJob.dll,zl9MedRec.dll','1,3','部件功能:住院首页、病案首页程序系统定位:处理病人住院首页、病案首页相关业务缺失后果:临床工作站、病案系统等业务模块无法使用',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPATIADDRESS.OCX','AA4431E0EFE116459F1F7F86B5E2DE3E','10.34.170',To_Date('2019-02-19 17:48:04', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9RegEvent.dll,zl9CardSquare.dll,zl9CISJob.dll,zl9InPatient.dll,zl9Patient.dll,zl9MedRec.dll,zl9InExse.DLL','1,3','部件功能:结构化地址部件系统定位:支持系统中进行结构化地址填写缺失后果:无法使用入院、挂号、首页等相关程序',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPUBLICADVICE.DLL','5A013B7A783252F89B9260B83A5660FC','10.34.170',To_Date('2019-02-19 17:56:34', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9CISJob.dll','1,3,21,22,24,26','部件功能:封装医嘱核心业务功能，提供公共接口，DOCK页签等。系统定位:封装工作站和医嘱核心业务功能缺失后果:通过该部件去使用临床功能会报错或者失效。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPUBLICEXPENSE.DLL','2FA060008DE1957D09555273C528C371','10.34.170',To_Date('2019-02-19 17:36:36', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9InExse.DLL,zl9OutExse.dll','1,3,21,22,24,26','部件功能:费用公共部件。系统定位:提供医生站预约、挂号，医嘱附费等功能，提供公共接口。缺失后果:费用相关功能无法使用',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPUBLICPACS.DLL','750508A63E1BC0C519F11530D4747215','10.34.170',To_Date('2019-02-19 17:58:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9PACSWork.dll','1,3,21,22,24,26','部件功能:封装PACS依赖业务调用接口系统定位:调用pacs相关的处理功能缺失后果:如临床不能进行pacs观片等',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPUBLICPATH.DLL','7E06E83CE08E1E3CBE63C62E12A5B3B8','10.34.170',To_Date('2019-02-19 17:58:24', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9CISJob.dll','1,3,21,22,24,26','部件功能:临床路径公共接口部件系统定位:提供临床路径开放接口缺失后果:临床路径功能无法正常运行',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPUBLICPATIENT.DLL','7B7E6343318488B6CA20FD6065293B9F','10.34.170',To_Date('2019-02-19 17:56:40', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9CISJob.dll,zl9Patient.dll,zl9InPatient.dll','1,3,21,22,24,26','部件功能:病人信息公共部件,封装了病人信息相关的公共方法：病人基本信息调整、身份证号反算年龄等。系统定位:供各个业务模块调用。如首页、病人信息管理、病人入院管理的基本信息调整功能。缺失后果:程序无法正常运行',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPUBLICPEIS.DLL','E62DA486310938E30979114ADB3029E6','10.34.170',To_Date('2019-02-19 17:36:34', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','ZL9Peis.dll,zl9PeisManage.dll','1,3,21,22,24,26','部件功能:体检公共接口系统定位:提供其他业务或三方调用体检功能的接口（如生成PDF，查阅体检报告）缺失后果:该接口不能正常工作。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BASEITEM.DLL','7617C6A1B3EC829BCBFE1755CD3DB1E1','10.34.170',To_Date('2019-02-19 17:48:38', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9BaseItem.dll,zl9InExse.DLL,zl9OutExse.dll','1,3,4,6,21','部件功能:基础数据管理系统定位:业务基础部件，包括部门，人员，收费项目，收入项目，各业务公共参数等基础设置缺失后果:无法进行基础数据设置',1,0,Null);
EXECUTE Zlfiles_Autoupdate('QRMAKER.OCX','C00A0B76BC515DAA01060F7F9A230D0D','1.31.0.0',To_Date('2013-11-19 14:13:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),4,'[SYSTEM]',Null,'1,3,4,6,21,22,23,24,26','部件功能:输出2D条码使用原因:用户需求缺失后果:体检报告输出不完整',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL7Z.DLL','B4BB5790C910278FCBD443B672620E5B','1.0.0',To_Date('2017-08-08 11:34:16', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[SYSTEM]',Null,'1,3,4,6,21,22,23,24,26','部件功能:中联7z压缩解压部件系统定位:进行文件压缩解压缺失后果:无法进行客户端部件升级。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9APPTOOL.DLL','2E76F0420889B70CD1CCA56A5AB7953A','10.34.170',To_Date('2019-02-19 17:35:38', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,3,4,6,21,22,23,24,26','部件功能:导航台的中提供的基础应用工具功能部件。系统定位:提供了个人系统级的个性化设置、系统基础字典数据管理、邮件收发管理等功能。缺失后果:缺失将会导致无法使用整个系统。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BILLEDIT.OCX','603EAB07721073D19A2D47FB33C44844','10.34.170',To_Date('2019-02-19 17:34:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,3,4,6,21,22,23,24,26','部件功能:单据编辑控件系统定位:提供表格控件缺失后果:用到此控件的界面不能正常打开，或者报错。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9COMLIB.DLL','5EB61E2EE67F3667784B2530375751EE','10.34.170',To_Date('2019-02-19 17:17:26', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,3,4,6,21,22,23,24,26','部件功能:公共的基础函数库，用来提供统一的SQL查询、一些系统常用API封装、常用控件处理、常用类型方法以及应用系统基础业务的常用查询。系统定位:ZLHIS的系统底层支持部件，一般部件均使用该部件提供的公共方法进行编码缺失后果:整个应用系统无法使用。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9ESIGN.DLL','151125F10A4C80608A157137804F6F26','10.34.170',To_Date('2019-02-19 17:35:42', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zl9CISJob.dll','1,3,4,6,21,22,23,24,26','部件功能:电子签名部件系统定位:集成不同CA厂商的电子签名接口并供各个业务模块调用。缺失后果:各个业务模块将无法启用电子签名功能',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9INSURE.DLL','7AECC1FA6C9B820702F39129411A2AAF','10.34.170',To_Date('2019-02-19 17:43:16', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,3,4,6,21,22,23,24,26','部件功能:医保接口部件。使用原因:医保项目检查，医保记帐作废上传缺失后果:重打时出错',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9REPORT.DLL','BF1C48B4B596F4ABFB87EE0646799A66','10.34.170',To_Date('2019-02-19 17:35:32', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,3,4,6,21,22,23,24,26','部件功能:支持业务系统输出自定义报表的内容，以及设计自定义报表系统定位:方便用户和技术人员缺失后果:缺少将无法输出报表和设计报表',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLBRW.DLL','F909E6E5F6555A4CFA92ABB91264C6F2','10.34.170',To_Date('2019-02-19 17:47:24', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[APPSOFT]\APPLY',Null,'1,3,4,6,21,22,23,24,26','部件功能:标准导航台样式，即双列表样式。系统定位:用来进行各个业务的导航，并提供一些基础的工具。缺失后果:缺失该部件，当选择的导航台样式为该样式时，无法进入整个系统。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLHIS+.EXE','93AFEBAAAD4F5C6C7348C27E6722D5AA','10.34.170',To_Date('2019-02-19 17:47:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[APPSOFT]',Null,'1,3,4,6,21,22,23,24,26','部件功能:ZLHIS+启动程序。系统定位:登录该程序才能进入导航台，进行业务操作。缺失后果:缺失该部件将无法进行各项业务。',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLHISCRUST.EXE','54198BA377688A1459901F0EC413B839','10.35.110',To_Date('2018-09-21 10:18:36', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[APPSOFT]',Null,'1,3,4,6,21,22,23,24,26','部件功能:客户端自动升级工具。系统定位:通过该工具对各个客户端进行文件升级。缺失后果:缺失该文件，将会导致需要升级的客户端无法进入导航台。',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLHISCRUSTCOM.DLL','497540A71D209A3683776E56EBF8B961','10.34.170',To_Date('2019-02-19 18:15:46', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[APPSOFT]',Null,'1,3,4,6,21,22,23,24,26','自动升级业务逻辑部件。缺失该部件在后续版本可能无法正常进行自动升级。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLICCARD.DLL','0BEA4C2CAC039D4509A3967D78936E03','10.34.170',To_Date('2019-02-19 17:35:48', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,3,4,6,21,22,23,24,26','部件功能:IC卡读卡的统一接口系统定位:IC卡读卡缺失后果:一卡通无法使用，部分会导致程序异常退出',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLIDCARD.DLL','3D81C62BC721C566063F8D3C1DE64976','10.34.170',To_Date('2019-02-19 17:35:44', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,3,4,6,21,22,23,24,26','部件功能:身份证读卡的统一接口系统定位:读取身份证信息缺失后果:无法读取身份证信息',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLIDKIND.OCX','63F501584C7BE355DDAAA8AAC6075409','10.34.170',To_Date('2019-02-19 17:17:06', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,3,4,6,21,22,23,24,26','部件功能:病人身份识别控件系统定位:刷卡和读卡查询病人缺失后果:使用的地方应控件丢失出错',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLMDI.DLL','C6CC418F5434D1D66FA2CE3B3BC38B4B','10.34.170',To_Date('2019-02-19 17:47:26', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[APPSOFT]\APPLY',Null,'1,3,4,6,21,22,23,24,26','部件功能:Mdi样式导航台，即父窗体里面存在各个子窗体。系统定位:用来进行各个业务的导航，并提供一些基础的工具。缺失后果:缺失该部件，当选择的导航台样式为该样式时，无法进入整个系统。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLRUNAS.EXE','CACE237C26F0699C63828FF3D79B8566','9.43.0',To_Date('2013-11-04 10:14:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),4,'[APPSOFT]',Null,'1,3,4,6,21,22,23,24,26','该文件在自动升级zlhisCrust.exe中使用。主要功能,在USER权限下可以使用管理员权限来进行登录执行管理操作',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLWIN.DLL','077AD0FDD27FCC90158C8543937C5E33','10.34.170',To_Date('2019-02-19 17:47:30', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[APPSOFT]\APPLY',Null,'1,3,4,6,21,22,23,24,26','部件功能:仿Window桌面样式导航台。系统定位:用来进行各个业务的导航，并提供一些基础的工具。缺失后果:缺失该部件，当选择的导航台样式为该样式时，无法进入整个系统。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PRINTMODE.DLL','34954C60CAC29622AFA61FDB4D4C3B21','10.34.0',To_Date('2014-10-30 20:12:32', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'1,3,4,6,21,22,23,24,26,28','部件功能:打印表格控件内容、通过命令生成输出内容等系统定位:方便用户输出数据缺失后果:可以缺少，但无法输出数据',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLUPGRADEREADER.EXE','34E43C8DA16DDAEFC286BAFCC4C083BE','10.34.0',To_Date('2014-10-30 20:17:58', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[APPSOFT]',Null,'1,3,6,22,23,24,26','部件功能:升级说明阅读器。系统定位:进行重大功能的核对以及培训事宜的处理。缺失后果:无法进行升级问题清单的阅读与核对。',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9DUE.DLL','FF0D4463C940A2C5FFA482AC71BC5C08','10.34.170',To_Date('2019-02-19 17:49:20', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'1,4,6','部件功能:付款管理部件系统定位:医院所有采购商品应付和已付的管理缺失后果:缺少将无法使用付款管理',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISBASE.DLL','D0486E3B3587C32430B0B21C47B7B305','10.34.170',To_Date('2019-02-19 18:13:46', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll','21','部件功能:体检基础业务系统定位:体检基础数据设置（增删改），包括体检相关的公共或系统参数设置，是体检产品运行必不可少的部件。缺失后果:不能维护体检基础数据；同时体检业务大部份功能也会不正常。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISCOMLIB.DLL','A01FD51C7BC2A44F5C57586378AE768D','10.34.170',To_Date('2019-02-20 09:35:06', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检公共组件系统定位:给体检基础功能和业务功能提供公共的方法函数缺失后果:整个体检产品无法使用。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISDEVANALYSE.DLL','EFEF29F2F55B7E5E979181B2534A8CCB','10.34.0',To_Date('2014-10-30 22:47:08', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检非标准串口的接口程序系统定位:定时读取非标准串口仪器所产生的指标结果数据缺失后果:不能接收到非标准串口仪器的指标结果数据',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISFLOW.DLL','204556F12B3B0F670D57DA85B63FB70D','10.34.170',To_Date('2019-02-19 18:12:54', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检执行业务系统定位:实现体检分科、总检等业务功能缺失后果:体检产品不能使用',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISGROUPRPT.DLL','AC6F55AD33A3EF94F4056912A136E7D9','10.34.170',To_Date('2019-02-19 18:06:38', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检团体报告预览及打印输出。系统定位:实现体检团体报告预览及打印输出。缺失后果:不能实现团体报告预览及打印输出',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISINNERINTERFACE.DLL','AE079FFE1A9532EB5887408ADE5F8E45','10.34.170',To_Date('2019-02-19 18:12:14', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检内部接口系统定位:实现费用、排队叫号、医嘱、票据等的数据交换缺失后果:体检产品不能正确运行',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISINSTRUMENT.DLL','BC352B7E5B4E7EC9669FF2F16F587493','10.34.0',To_Date('2014-10-30 22:39:08', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检基本仪器数据接口系统定位:完成体检基本仪器数据传输接口缺失后果:不能接收到身高体重仪等基本仪器数据',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISINTERFACE.DLL','77831E1D322F32A2129860F94A882619','10.34.0',To_Date('2014-10-30 22:37:20', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:第三方体检接口部件系统定位:实现三方体检和ZLHIS的数据交换缺失后果:不能实现三方体检和ZLHIS的接口',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISMANAGE.DLL','9238C8C4430548A7951F44F1B05363A7','10.34.170',To_Date('2019-02-19 18:13:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll','21','部件功能:体检管理业务系统定位:实现体检登记、报到、填写结果、打印报告、指引单等业务缺失后果:体检产品不能使用。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISPERSONPDF.DLL','62B6AE968619645CD79536A439B43E73','10.34.170',To_Date('2019-02-19 18:14:16', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检报告PDF输出接口系统定位:体检报告打印三方生成的PDF文件缺失后果:不能完整输出体检报告内容',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISPERSONREPORT.DLL','2A2DDC7B04713415851188A18BA136D7','10.34.170',To_Date('2019-02-19 18:14:12', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检个人报告自定义报表系统定位:体检报告中需要调用自定义报表进行打印缺失后果:体检报告中自定义报表时不能输出完整的体检报告',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISPERSONRPT.DLL','2AF095E725F497807CDF10D153D35B2D','10.34.170',To_Date('2019-02-19 18:06:30', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:个人体检报告输出系统定位:按固定的格式生成报告打印数据缺失后果:不能输出个人体检报告',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PEISRPT.DLL','495A16FDEDB5F43A108B80932581121B','10.34.170',To_Date('2019-02-19 18:04:08', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检报告格式系统定位:个人体检报告的内容生成缺失后果:不能生成报告内容',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPEISAUTOANALYSE.EXE','0D8CCFDC4DB69FEDB858CBC94FA3743B','10.34.0',To_Date('2014-10-30 22:47:24', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]','ZL9Peis.dll,zl9PeisManage.dll,zl9PeisBase.dll','21','部件功能:体检自动分析服务系统定位:实现非标准的仪器数据接口缺失后果:不能接收到非标准串口的体检仪器数据',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9DRAWREPORT.DLL','9887A1B8F338FD1DAD6FA5B6D8959B63','10.34.170',To_Date('2019-02-19 18:03:44', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','ZL9Peis.dll,zl9PeisManage.dll,zl9Oper.dll','21,24','部件功能:zl9DrawReport系统定位:实现固定报告格式的打印输出预览缺失后果:不能实现打印输出预览功能',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9BLOOD.DLL','C8B8E363B6EB52D01E9712C43F78EFE7','10.34.170',To_Date('2019-02-19 18:03:34', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'22','部件功能:血库系统核心框架部件系统定位:包含血库相关功能模块：血液目录管理、血液供应入出库、科室配血管理、科室发血管理、血袋回收、报废等缺失后果:血库系统将无法使用',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLPUBLICBLOOD.DLL','A374390EC638722891025570555F9278','10.34.170',To_Date('2019-02-19 18:01:26', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),0,'[PUBLIC]',Null,'22','血库业务封装公共部件',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9INFECT.DLL','106B748E40A093C2CD5AD93238D2908C','10.34.0',To_Date('2017-04-13 17:31:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'23','部件功能:院感管理系统核心框架部件系统定位:院感系统核心功能窗体，含：病例监测管理、病例日报管理、人员监测管理、医院感染汇总表缺失后果:无法使用院感系统',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9OPER.DLL','9A1538CFB0E0266480B894C85CF9A28D','10.34.170',To_Date('2019-02-19 18:04:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'24','部件功能:手术麻醉部件系统定位:实现手术安排及相关计费缺失后果:手麻产品不能运行',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9OPSSTARAND.DLL','605E76487A9E9E8B5235695D8512A6DD','10.34.0',To_Date('2014-10-30 22:35:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9Oper.dll','24','部件功能:手术麻醉单打印预览系统定位:生成手术麻醉单，并进行打印或预览缺失后果:不能生成手术麻醉单',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLOPERINTERFACE.DLL','9B57A5CFF5240D49E8468432CDD7127D','10.34.170',To_Date('2019-02-19 18:06:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9Oper.dll','24','部件功能:三方手麻软件接口ZLHIS系统定位:实现三方手麻产品和ZLHIS产品之间的功能接口缺失后果:无法实现和ZLHIS的数据交换。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9LABPRINTSVR.EXE','77DD418BDE71674CD5796CC0565CB273','10.34.0',To_Date('2014-10-30 22:39:12', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]','ZL9LabWork.dll','25,26','部件功能:新版LIS打印服务系统定位:主要处理 批量打印报告缺失后果:导诊和新版LIS打印报告部分将不能使用',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9COMLIBPSS.DLL','6C9B83649EFC93F3C24DF746EE5CAFFA','10.34.170',To_Date('2019-02-19 18:07:50', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zlWizardStart.exe','26','部件功能:导诊公共函数部件系统定位:提供公共方法函数（和导诊业务无关）缺失后果:系统不能运行',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDCARDS.DLL','99F23F6ECE036A250D072E04B89C578C','10.34.170',To_Date('2019-02-19 18:11:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:自助发卡与签约系统定位:提供自助设备上进行发卡和绑定卡操作缺失后果:自助发卡与签约功能缺失',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDCONTROL.OCX','6515AEF98D34542432542869D7571328','10.34.170',To_Date('2019-02-19 18:08:46', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zlWizardStart.exe','26','部件功能:提供自助系统所需要的控件系统定位:完成自助系统的控件统一效果及功能实现缺失后果:不能运行自助系统并报错',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDDEPOSIT.DLL','179279E195B0AEB1653A0F52012879C9','10.34.170',To_Date('2019-02-19 18:10:38', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:自助充值管理系统定位:提供门诊预交和住院预交以及历史充值记录查询缺失后果:自助充值及查询功能缺失',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDEMR.DLL','E6C1A4E243F80B7AAB994E835AAAC369','10.34.170',To_Date('2019-02-19 18:14:02', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:zl9WizardEMR系统定位:门诊电子病历缺失后果:无法使用自助服务系统门诊电子病历相关功能',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDFEEQUERY.DLL','2EC85CB6E18272981771F7D3F71D548C','10.34.170',To_Date('2019-02-19 18:10:42', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:自助费用查询系统定位:自助设备上查询病人门诊和住院费用缺失后果:自助费用查询功能缺失',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDINVOICE.DLL','E7B231327273AD9508F8A5B219C85C5E','10.34.170',To_Date('2019-02-19 18:10:48', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:自助票据打印系统定位:自助设备上对未打印票据的缴费和挂号单据打印票据缺失后果:挂号和收费票据打印功能缺失',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDLABCALL.DLL','BB0E4A44FAB6A363161AB40DBF7E7FB1','10.34.0',To_Date('2015-09-30 16:21:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:zl9WizardLABCall系统定位:检验叫号缺失后果:暂无',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDLABPRINT.DLL','A126883AFF6AD4B18FDED14DE96899CA','10.34.170',To_Date('2019-02-19 18:10:54', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:导诊打印老版LIS报告系统定位:在导诊系统中，打印老版LIS相关报告缺失后果:不能正常打印老版LIS报告',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDLIB.DLL','BB7173F5A12954EF3182E6DB9F2CC4B4','10.34.170',To_Date('2019-02-19 18:08:18', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[PUBLIC]','zlWizardStart.exe','26','部件功能:病人自助系统公共库系统定位:提供自助系统中需要使用的方法函数缺失后果:系统不能运行',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDMAIN.EXE','7AC6D52A169DB91E4E5C0B03C83C5CD1','10.34.170',To_Date('2019-02-19 18:11:14', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]',Null,'26','部件功能:病人自助系统后台管理程序系统定位:完成自助系统的所有后台设置，包括资源配置、动态页面设计、静态页面参数等缺失后果:系统不能正常运行',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDMANAGE.DLL','7D0F94F5B2BEA35A72A5F8FC7483DBD4','10.34.170',To_Date('2019-02-19 18:10:00', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9WizardMain.exe','26','部件功能:自助系统后台管理系统定位:配置自助系统的资源、页面、参数等缺失后果:不能启动后台管理',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDPAGE.DLL','A531F9BF5823D33E8511F75B26965944','10.34.170',To_Date('2019-02-19 18:09:34', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:自助动态页面显示系统定位:根据动态页面的设计显示最终的页面展示效果缺失后果:自助系统不能使用',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDPAYFEE.DLL','AD6EADD02261B26C2FE5ED682EA00058','10.34.170',To_Date('2019-02-19 18:10:58', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:自助缴费管理系统定位:自助设备上对划价单据进行缴费缺失后果:自助缴费功能缺失',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDPEISQUEUE.DLL','0FBB0B24D38B2DB3801B8DBD99D969B0','10.34.0',To_Date('2017-10-25 18:58:44', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:体检自助排队系统定位:体检人员通过病人自助系统提供体检自助排队功能进行自助体检排队缺失后果:体检人员无法自助排队',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDPHARMACY.DLL','A8F887EE6332C6233666173CF6461706','10.34.170',To_Date('2019-02-19 18:15:24', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:自助系统中的门诊药房病人签到管理系统定位:作为自助系统的一部分存在缺失后果:缺少将无法使用门诊药房签到',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDPRICE.DLL','9193C1D05EAF22BA0115A2FBFAECFAA8','10.34.170',To_Date('2019-02-19 18:09:40', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:提供收费项目的价格自助查询系统定位:通过简码等实现收费项目价格的自助查询缺失后果:不能自助查询价格',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDPROFICIENT.DLL','608F104DB783C178875C56D092744CBB','10.34.170',To_Date('2019-02-19 18:09:46', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:实现医院专家介绍自助查询系统定位:实现医院专家介绍自助查询缺失后果:不能实现专家介绍查询',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDPROOF.DLL','91661D7BEAFD25134468AAFD65D1C797','10.34.170',To_Date('2019-02-19 18:14:12', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:自助凭条打印系统定位:自助设备上打印缴费和挂号凭条缺失后果:缴费和挂号凭条打印功能缺失',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDREGEVENT.DLL','87F4F69FEDF67CDEB67CDF0A6B6B462C','10.34.170',To_Date('2019-02-19 18:11:16', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:自助挂号和预约系统定位:自助设备上病人进行挂号和预约以及取号缺失后果:挂号和预约以及取号功能缺失，今日就诊也不能进行挂号',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9WIZARDTODAY.DLL','4127E584DDA88A1E28B25EF727D59195','10.34.170',To_Date('2019-02-19 18:09:52', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:今日就诊系统定位:查询挂号安排，进行挂号/预约，查看挂号科室上班时间缺失后果:挂号、预约以及查询功能缺失',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLSWITCHEFFECT.DLL','8C04FB8034AE0E7F4A404ABCFAC73311','10.34.0',To_Date('2015-09-30 16:21:28', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:实现自助页面切换效果系统定位:实现自助页面切换效果缺失后果:自助页面切换时没有切换变换效果，直接切换。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLSWITCHPAGE.DLL','9F7235B2EAC8C3D08CB585AD0AA2227D','10.34.0',To_Date('2015-09-30 16:21:26', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:实现自助页面切换效果系统定位:实现自助页面切换效果缺失后果:没有切换效果',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLWIZARDHEC.DLL','77E056E1120F1F597858D4BCE892AC32','10.34.170',To_Date('2019-02-19 18:15:36', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[Appsoft]\Apply','zl9WizardManage.dll','26','自助签到',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLWIZARDHECPRINT.DLL','8AB751D471F0E678AF4E6AE65C89C713','10.34.170',To_Date('2019-02-19 18:15:42', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[Appsoft]\Apply','zl9WizardManage.dll','26','自助打印体检报告',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLWIZARDNEWLABPRINT.DLL','2587AD9F501B33F8C42B1406F4EBA65B','10.34.170',To_Date('2019-02-19 18:10:34', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:导诊新版LIS打印部件系统定位:在导诊系统中，进行新版LIS报告打印缺失后果:新版LIS报告将不能正常打印',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLWIZARDPACSPRINT.DLL','64C85BFFD460CBB583C0F00193C9D4D8','10.34.170',To_Date('2019-02-19 18:11:42', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zlWizardStart.exe','26','部件功能:pacs报告自助打印系统定位:提供患者自助服务系统支持缺失后果:不能进行pacs自助打印',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZLWIZARDSTART.EXE','0C3B1F091D9F89B6A7EA620B6FB3965C','10.34.170',To_Date('2019-02-19 18:10:02', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]',Null,'26','部件功能:自助系统前台查询启动程序系统定位:启动自助系统前台功能缺失后果:不能运行自助系统前台查询',0,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9MEDREC.DLL','BDB880798370A0E4825F0C43603D3F42','10.34.170',To_Date('2019-02-19 18:02:52', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'3','部件功能:病案管理事务系统定位:病案系统管理、门诊日报、住院日报功能缺失后果:病案系统、门诊、住院日报功能失效。',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9MATERIAL.DLL','D025BB1CFE0A24679B90BD3646FEDADC','10.34.170',To_Date('2019-02-19 18:02:26', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'4','部件功能:管理医院的物资系统定位:既可以独立存在，也可以共享标准系统存在缺失后果:缺少将无法管理物资',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9PURVEY.DLL','4862207BA0FEB874EC2AB49D3C2BC1CA','10.34.170',To_Date('2019-02-19 18:03:40', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY','zl9Material.dll','4','部件功能:管理供应室的器械包系统定位:属于物资系统的子系统缺失后果:缺少将无法管理器械包',1,0,Null);
EXECUTE Zlfiles_Autoupdate('ZL9DEVICE.DLL','CADBCF649821BEA054CA79E001865393','10.34.170',To_Date('2019-02-19 18:02:22', 'yyyy-mm-dd HH24:mi:ss'),To_Date('', 'yyyy-mm-dd HH24:mi:ss'),1,'[APPSOFT]\APPLY',Null,'6','部件功能:管理医院的设备系统定位:既可以独立存在，也可以共享标准系统存在缺失后果:缺少将无法管理设备',1,0,Null);

--系统版本号
Update zlSystems Set 版本号='10.34.170' Where 编号=&n_System;
--部件版本号
Commit;