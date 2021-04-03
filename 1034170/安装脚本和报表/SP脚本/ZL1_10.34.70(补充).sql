--[连续升级]1
--[管理工具版本号]10.34.30
--本脚本支持从ZLHIS+ v10.34.60 升级到 v10.34.70
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--90040:梁经伙,2016-03-17,增加身份证未录原因，如果存在 身份证号状态 表，则改名为 身份证未录原因   
Declare
  n_Tableexist Number := 0;
Begin
  Begin
    Execute Immediate 'select count(*)  from USER_TABLES T WHERE T.TABLE_NAME = ''身份证号状态'''
      Into n_Tableexist;
  
    If n_Tableexist = 1 Then
      Execute Immediate 'ALTER TABLE 身份证号状态 RENAME TO 身份证未录原因';
      Execute Immediate 'alter table 身份证未录原因 drop Constraint 身份证号状态_PK';
      Execute Immediate 'alter table 身份证未录原因 drop Constraint 身份证号状态_UQ_名称';
    Else
      Execute Immediate 'Create Table 身份证未录原因(
                                                    编码 VARCHAR2(2),
                                                    名称 VARCHAR2(50),
                                                    简码 VARCHAR2(10),
                                                  缺省标志 NUMBER(1) default 0,
                                                    说明 VARCHAR2(50))
                                                    TABLESPACE zl9BaseItem';
    End If;
  Exception
    When Others Then
      Null;
  End;
End;
/
Alter Table 身份证未录原因 Add Constraint 身份证未录原因_PK Primary Key (编码) Using Index Tablespace zl9Indexcis;
Alter Table 身份证未录原因 Add Constraint 身份证未录原因_UQ_名称 Unique (名称) Using Index Tablespace zl9Indexhis;

--92729:胡俊勇,2016-03-08,专业版RIS接口处理
Create Table RIS检查预约 (
医嘱ID  NUMBER(18),
预约ID   NUMBER(18),
预约日期  DATE,
检查设备ID  NUMBER(18),
检查设备名称  VARCHAR2(64),
预约开始时间  DATE,
预约结束时间  DATE,
预约开始时间段  DATE,
预约结束时间段  DATE,
待转出  NUMBER(3))     
TABLESPACE zl9CisRec;

Alter Table RIS检查预约 Add Constraint RIS检查预约_PK Primary Key(医嘱ID) Using Index Tablespace zl9Indexhis;
Alter Table RIS检查预约 Add Constraint RIS检查预约_FK_医嘱ID Foreign Key (医嘱ID) References 病人医嘱记录(ID) On Delete Cascade;
Create Index RIS检查预约_IX_待转出 On RIS检查预约(待转出) Tablespace zl9Indexcis;

--93380:刘尔旋,2016-02-24,财务组多组长轧帐
Create Table 财务组组长构成(
    组ID Number(18),
    组长ID Number(18),
    上次轧帐时间 Date)
    TABLESPACE zl9Expense
    PCTFREE 5;

Alter Table 财务组组长构成 Add Constraint 财务组组长构成_PK Primary Key (组ID,组长ID) Using Index Tablespace ZL9INDEXHIS;
Create Index 财务组组长构成_IX_组长ID On 财务组组长构成(组长ID) Pctfree 5 Tablespace zl9indexhis;

--91225:梁经伙,2016-02-16,疾病阳性记录 新增字段 医嘱ID
alter table  疾病阳性记录 add (医嘱ID number(18));
Create Index 疾病阳性记录_IX_医嘱ID On 疾病阳性记录(医嘱ID) Tablespace zl9Indexcis;
Alter Table 疾病阳性记录 Add Constraint 疾病阳性记录_FK_医嘱ID Foreign Key (医嘱ID) References 病人医嘱记录(ID);

--91316:张德婷,2016-01-21,输液配置中心批次
alter table 配药工作批次 add 药品类型 varchar2(20);

--91333:张德婷,2016-01-21,输液配置中心打印流水号
alter table 输液配药记录 modify 打印标志 number(5);

--92808:马政,2016-01-22,卫材单据打印控制
Create table 药品收发主表 
(
ID number(18),
no varchar2(8),
单据 number(2),
库房id number(18),
打印状态 number(1)
) tablespace ZL9MEDLST;

Create Sequence 药品收发主表_ID Start With 1; 
Alter Table 药品收发主表 Add Constraint 药品收发主表_PK Primary Key (ID) Using Index Tablespace zl9indexhis;
Alter Table 药品收发主表 Add Constraint 药品收发主表_UQ_NO Unique (no,单据,库房id) Using Index Tablespace zl9indexhis;

--93380:刘尔旋,2016-02-24,财务组多组长轧帐
Alter Table 财务组组长构成 Add Constraint 财务组组长构成_FK_组ID Foreign Key (组ID) References 财务缴款分组(Id);
Alter Table 财务组组长构成 Add Constraint 财务组组长构成_FK_组长ID Foreign Key (组长ID) References 人员表(Id);

--91225:梁经伙,2016-3-7,修改 疾病阳性记录，传染病记录表里面的字段长度
alter table 传染病目录 modify  编码  VARCHAR2(20);
alter table 疾病阳性记录 modify  标本名称 varchar2(64);
alter table 疾病阳性记录 modify  送检医生 VARCHAR2(100);
alter table 传染病目录 modify  (名称 varchar2(150),简码 VARCHAR2(20),说明 VARCHAR2(200));

-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--90040:梁经伙,2016-03-10,增加身份证未录原因字典表的默认数据
Insert Into 身份证未录原因
  (编码, 名称, 简码)
  Select '01', '未带', 'WD' From Dual Where Not Exists (Select 1 From 身份证未录原因 Where 编码 = '01');
Insert Into 身份证未录原因
  (编码, 名称, 简码)
  Select '02', '遗失待办', 'YSDB' From Dual Where Not Exists (Select 1 From 身份证未录原因 Where 编码 = '02');
Insert Into 身份证未录原因
  (编码, 名称, 简码)
  Select '03', '未办', 'WB' From Dual Where Not Exists (Select 1 From 身份证未录原因 Where 编码 = '03');


--90040:梁经伙,2016-03-17,如果zlBaseCode表里面已经存在身份证号状态表的数据就直接修改
update zlBaseCode set 表名 = '身份证未录原因' where 系统 = &n_sysTem and 表名 = '身份证号状态';
--90040:梁经伙,2016-03-10,增加身份证未录原因字典表
Insert Into zlBaseCode
  (系统, 表名, 固定, 说明, 分类)
  Select &n_System, '身份证未录原因', 0, '病人身份证号码未录入的原因', '人员属性'
  From Dual
  Where Not Exists (Select 1 From zlBaseCode Where 系统 =&n_System and  表名 = '身份证未录原因');


--92729:胡俊勇,2016-03-08,专业版RIS接口处理
Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
Select &n_System,8,A.* From (
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0 Union All
Select 'RIS检查预约',21,1,-NULL From Dual Union All
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0) A;

--93917:许华峰,2016-03-07,与RIS的数据交换接口脚本
Insert Into zlComponent(部件,名称,主版本,次版本,附版本,系统,注册产品名称,注册产品简名,注册产品版本) Values('zl9XWInterface','影像信息系统工作部件',10,35,10,&n_System,'中联医院信息系统','ZLHIS+','10'); 
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values( 1287,'医学影像信息系统专业版','影像RIS和PACS系统',&n_System,'zl9XWInterface'); 

Insert Into zlMenus(组别,ID,上级ID,标题,快键,系统,模块,短标题,图标,说明)
	Select 组别,Zlmenus_Id.Nextval,ID,'医学影像信息系统专业版' ,'B' ,&n_System,-NULL ,'医学影像信息系统专业版' ,99 ,'影像RIS和PACS系统' 
         From zlMenus Where 标题 = '医学影像系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null;                 

Insert Into zlMenus(组别,ID,上级ID,标题,快键,系统,模块,短标题,图标,说明) 
	Select A.组别,ZlMenus_ID.Nextval,A.ID,B.* From (
	     Select 组别,ID From zlMenus Where 标题 = '医学影像信息系统专业版' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null) A,
	    (Select 标题,快键,系统,模块,短标题,图标,说明 From zlMenus Where 1 = 0 Union All
	Select '医学影像信息系统专业版' ,'R'  ,&n_System, 1287, '影像信息工作站' ,99, '影像RIS和PACS工作站'  From Dual Union All						
                 Select 标题,快键,系统,模块,短标题,图标,说明 From zlMenus Where 1 = 0) B; 

--78768:李南春,2016-03-04,IDKind增加缺省读卡类别
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 1,'缺省读卡类别', Null, '0', '设置缺省的读卡的类别,存储的是卡类别ID'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1153 And 参数名 = '缺省读卡类别');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 2,'向前滚动-功能键', Null, Null, '设置IDkind的功能键'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1153 And 参数名 = '向前滚动-功能键');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 3,'向前滚动-快键', Null, Null, '设置IDkind的功能键'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1153 And 参数名 = '向前滚动-快键');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 4,'向后滚动-功能键', Null, Null, '设置IDkind的功能键'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1153 And 参数名 = '向后滚动-功能键');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 5,'向后滚动-快键', Null, Null, '设置IDkind的功能键'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1153 And 参数名 = '向后滚动-快键');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 6,'读卡-功能键', Null, Null, '设置IDkind的功能键'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1153 And 参数名 = '读卡-功能键');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 7,'读卡-快键', Null, Null, '设置IDkind的功能键'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1153 And 参数名 = '读卡-快键');

--91954:张德婷,2016-02-29,区别阳煤和涪陵医院的自动排批规则
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 35, '自动排批时输液单的批次只往后面批次变动', '0', '0', '当该参数启用之后，自动排批不会将后面的批次排到前面，例如当2#不够时，不会将3批的分配到2#'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '自动排批时输液单的批次只往后面批次变动');

--92718:许华峰,2016-02-26,启用影像信息系统接口控制
Insert Into Zlparameters
  (Id, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 255, '启用医学影像信息系统专业版接口', '0', '0',
         '控制是否启用影像信息系统接口，0-未启用;1-启用，必须安装影像信息系统的才生效。'
  From Dual
  Where Not Exists (Select 1
         From Zlparameters
         Where 参数名 = '启用医学影像信息系统专业版接口' And Nvl(模块, 0) = 0 And Nvl(系统, 0) = &n_System);

--93311:张德婷,2016-02-24,发送医嘱的时候不允许置换药房到静配
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 34, '不允许置换药房到输液配置中心', '0', '0', '此参数开启后，发送医嘱的时候不允许置换药房到输液配置中心'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where Nvl(系统, 0) = &n_System And  Nvl(模块, 0) =1345 And 参数名 = '不允许置换药房到输液配置中心');

--92384:胡俊勇,2016-02-24,医嘱待执行消息
Insert Into 业务消息类型(编码,名称,说明,保留天数)  Select 'ZLHIS_CIS_034','待执行医嘱提醒','医嘱发送后有需要执行登记的医嘱，产生的一个待执行通消息。',7 From Dual;

--91956:张德婷,2016-02-22,除特殊药品外首日的药品医嘱不进入静配中心
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 32,'是否按设置的常用药品进行药品过滤操作', Null, '1', '不根据设置的常用药品过滤输液单的时候，药品是根据当天输液单的汇总数量确定'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '是否按设置的常用药品进行药品过滤操作');

--92537:梁经伙,2016-02-17,增加参数 诊断手术名称自由调整 控制诊断手术的自由调整
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, 0, 0, 0, 0, 257, '诊断手术名称自由调整', '0', '0', '控制是否启用诊断和手术的自由调整功能，启用后，住院，门诊和诊断选择器的诊断描述和手术名称可以自由修改。0-未启用;1-启用'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where Nvl(系统, 0) = &n_System And  Nvl(模块, 0) = 0 And 参数名 = '诊断手术名称自由调整');


--93221:刘尔旋,2016-02-02,缺省轧帐时间
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1506, 1, 0, 0, 0, 2, '缺省轧帐时间', Null, Null,
         '每天轧帐的默认时间,当前时间超过参数设置的时间的,截止时间默认为参数设置的时间;当前时间未超过参数设置的时间的,以当前时间为准'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 参数名 = '缺省轧帐时间' And 系统 = &n_System And 模块 = 1506);


--92468:李南春,2016-01-25,挂号发卡严格控制的情况下，卡费为零是否走发票号
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1111, 0, 0, 0, 0, 67, '零卡费走票号', '0', '0', '挂号发卡严格控制的情况下，卡费为零是否走发票号'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 参数名 = '零卡费走票号' And 系统 = &n_System And 模块 = 1111);


--91584:张德婷,2016-01-22,审核静配所有医嘱
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 30,'审核该药房的所有数据', Null, '0', '审核该药房的部门发药和静配中心的数据'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '审核该药房的所有数据');

--91732:张德婷,2016-01-22,除特殊药品外首日的药品医嘱不进入静配中心
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 31,'除特殊药品外其他首日的药品医嘱均在部门发药执行', Null, '0', '除特殊药品外其他首日的药品医嘱均在部门发药执行'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '除特殊药品外其他首日的药品医嘱均在部门发药执行');

--91954:张德婷,2016-01-22,Zl_输液配药记录_自动排批
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 33,'启动自动排批', Null, '0', '启动自动排批之后将根据容量优先级自动设置输液单的批次，同时不能保持上次批次'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '启动自动排批');

--92852:刘鹏飞,2016-01-21,床位卡片排序方式控制
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1265, 0, 0, 0, 0, 9, '床位卡片排序方式', '', '1', '控制床位卡片的显示顺序:是按找床号排序显示，还是按床位编制分类+床位号排序显示'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 参数名 = '床位卡片排序方式' And 系统 = &n_System And 模块 = 1265);

--91316:张德婷,2016-01-21,输液配置中心批次
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 27,'单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包', Null, '0', '单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 28,'按批次，药品排序', Null, '0', '此参数勾选之后，界面上的所有药品按批次，药品进行排序'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '按批次，药品排序');

Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 29,'特殊药品按药品类型指定批次', Null, '0', '肿瘤药，营养药根据配药批次的药品类型匹配批次'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1345 And 参数名 = '特殊药品按药品类型指定批次');

--92725:李业庆,2016-01-21,增加配药单打印显示方式参数
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1341, 0, 1, 0, 0, 55, '待配药单据打印显示方式', '0', '0', '0-显示所有配药单,1-只显示未打印的待配药单据,2-只显示已打印的待配药单据'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1341 And 参数名 = '待配药单据打印显示方式');

--92725:李业庆,2016-01-21,增加待发药单据扫描后自动呼叫参数
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1341, 0, 1, 0, 0, 56, '待发药单据扫描后自动呼叫', '0', '0', '0-不自动呼叫,1-扫描后自动语音呼叫'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1341 And 参数名 = '待发药单据扫描后自动呼叫');


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--94102:李小东,2016-03-14,对应处方权限更正
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1023, '对应处方', User, '病历单据应用', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1023 And 功能 = '对应处方' And Upper(对象) = Upper('病历单据应用'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1023, '对应处方', User, '病历文件列表', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1023 And 功能 = '对应处方' And Upper(对象) = Upper('病历文件列表'));

--93675:梁经伙,2016-03-14,诊断选择器怎加了存储过程增加权限
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1252,'医嘱下达',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_病案主页_首页整理EX','EXECUTE' From Dual Union All 
    Select 'Zl_病案主页从表_首页整理','EXECUTE' From Dual Union All 
    Select '分化程度','SELECT' From Dual Union All 
    Select '最高诊断依据','SELECT' From Dual Union All 
	Select '住院死亡原因','SELECT' From Dual Union All 
	Select '病案项目','SELECT' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1253,'医嘱下达',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_病案主页_首页整理EX','EXECUTE' From Dual Union All 
    Select 'Zl_病案主页从表_首页整理','EXECUTE' From Dual Union All 
    Select '分化程度','SELECT' From Dual Union All 
    Select '最高诊断依据','SELECT' From Dual Union All 
	Select '住院死亡原因','SELECT' From Dual Union All 
	Select '病案项目','SELECT' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;


--90040:梁经伙,2016-03-17,如果zlProgPrivs表里面已经存在身份证号状态的数据就直接修改
update zlProgPrivs set 对象 = '身份证未录原因' where 系统 = &n_System  and 对象 = '身份证号状态';

--90040:梁经伙,2016-03-10,为表 身份证未录原因 添加权限
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1261,'首页整理',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '身份证未录原因','SELECT' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1260,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '身份证未录原因','SELECT' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--92729:胡俊勇,2016-03-08,专业版RIS接口处理
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select  &n_System,1252,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select 'RIS检查预约','SELECT' From Dual Union All    
    Select 'Zl_Ris检查预约_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_Ris检查预约_Delete','EXECUTE' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select  &n_System,1253,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select 'RIS检查预约','SELECT' From Dual Union All    
    Select 'Zl_Ris检查预约_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_Ris检查预约_Delete','EXECUTE' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select  &n_System,1254,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select 'RIS检查预约','SELECT' From Dual Union All    
    Select 'Zl_Ris检查预约_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_Ris检查预约_Delete','EXECUTE' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--93917:许华峰,2016-03-08,与RIS的数据交换接口脚本
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
	Select &n_System,1287,A.* From (
	Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
	Select '基本',-NULL,NULL,1 From Dual Union All       
	Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;
       
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
	Select &n_System,1287,'基本',User,A.* From (
	Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
	Select '费别','SELECT' From Dual Union All
	Select '民族','SELECT' From Dual Union All
	Select '职业','SELECT' From Dual Union All
	Select '性别','SELECT' From Dual Union All
	Select '部门表','SELECT' From Dual Union All
	Select '人员表','SELECT' From Dual Union All
	Select '病人信息','SELECT' From Dual Union All
	Select '病案主页','SELECT' From Dual Union All
	Select '婚姻状况','SELECT' From Dual Union All
	Select '上机人员表','SELECT' From Dual Union All
	Select '人员证书记录','SELECT' From Dual Union All
	Select '部门性质说明','SELECT' From Dual Union All
	Select '病人挂号记录','SELECT' From Dual Union All
	Select '病人医嘱记录','SELECT' From Dual Union All
	Select '病人医嘱发送','SELECT' From Dual Union All
	Select '病人医嘱附费','SELECT' From Dual Union All
	Select '病人医嘱报告','SELECT' From Dual Union All
	Select '病人诊断医嘱','SELECT' From Dual Union All
	Select '病人诊断记录','SELECT' From Dual Union All
	Select '影像检查项目','SELECT' From Dual Union All
	Select '收费项目目录','SELECT' From Dual Union All
	Select '门诊费用记录','SELECT' From Dual Union All
	Select '住院费用记录','SELECT' From Dual Union All
	Select '诊疗项目目录','SELECT' From Dual Union All
	Select '诊疗项目部位','SELECT' From Dual Union All
	Select '诊疗执行科室','SELECT' From Dual Union All
	Select '诊疗收费关系','SELECT' From Dual Union All
	Select '保险支付项目','SELECT' From Dual Union All
	Select '体检任务人员','SELECT' From Dual Union All
	Select '体检任务发送','SELECT' From Dual Union All
	Select '医疗付款方式','SELECT' From Dual Union All
	Select 'ZlComponent','SELECT' From Dual Union All
	Select 'Zl_Lob_Append','EXECUTE' From Dual Union All 
	Select 'Zl_Lob_Read','EXECUTE' From Dual Union All
	Select 'Zl_Fun_Getsignpar','EXECUTE' From Dual Union All
	Select 'zl_影像消息_XML内容获取','EXECUTE' From Dual Union All
	Select 'zl_挂号病人病案_INSERT','EXECUTE' From Dual Union All
	Select 'ZL_病人医嘱记录_Insert','EXECUTE' From Dual Union All
	Select 'ZL_病人医嘱发送_Insert','EXECUTE' From Dual Union All
	Select 'Zl_Ris检查预约_Delete','EXECUTE' From Dual Union All
	Select 'Zl_Ris检查预约_Insert','EXECUTE' From Dual Union All
	Select 'b_zlXWInterface','EXECUTE' From Dual Union All
	Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--78768:李南春,2016-03-04,IDKind增加缺省读卡类别
Insert Into zlProgFuncs(系统,序号,功能,排列,缺省值,说明)
Select &n_System,1153,A.* From (
Select 功能,排列,缺省值,说明 From zlProgFuncs Where 1 = 0 Union All
Select '参数设置',1,1,'设置参数的操作权限。有该权限时,允许进行本地参数设置。' From Dual Union All
Select 功能,排列,缺省值,说明 From zlProgFuncs Where 1 = 0) A;

--93758:陈振原,2016-03-01,脚本oracle对象权限访问错误
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1075, '基本', User, '病历范文内容', 'SELECT'
  From Dual
  Where Not Exists (Select 1 From zlProgPrivs Where 系统 = &n_System And 序号 = 1075 And 对象 = '病历范文内容');

--92033:张德婷,2016-02-24,区分是否发过药
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1342, '基本', User, 'Zl_输液配药记录_审核', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1342 And 功能 = '基本' And Upper(对象) = Upper('Zl_输液配药记录_审核'));

--90447:刘尔旋,2016-02-22,复诊的确定
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1111, '基本', User, 'Zl1_Fun_Getreturnvisit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1111 And 功能 = '基本' And Upper(对象) = Upper('Zl1_Fun_Getreturnvisit'));

--91225:梁经伙,2016-02-18,医生站使用传染病反馈单添加的授权
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select  &n_System,1260,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '疾病阳性记录','SELECT' From Dual Union All 
    Select '疾病报告反馈','SELECT' From Dual Union All
	Select '疾病申报记录','SELECT' From Dual Union All     
    Select 'Zl_疾病阳性检测记录_Update','EXECUTE' From Dual Union All 
    Select 'Zl_疾病申报记录_Update','EXECUTE' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select  &n_System,1261,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '疾病阳性记录','SELECT' From Dual Union All 
    Select '疾病报告反馈','SELECT' From Dual Union All 
	Select '疾病申报记录','SELECT' From Dual Union All    
    Select 'Zl_疾病阳性检测记录_Update','EXECUTE' From Dual Union All 
    Select 'Zl_疾病申报记录_Update','EXECUTE' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--93336:胡俊勇,2016-02-16,输血执行登记
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1252, '基本', User, 'Zl_Fun_Get输血执行登记', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1252 And 功能 = '基本' And Upper(对象) = Upper('Zl_Fun_Get输血执行登记'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1253, '基本', User, 'Zl_Fun_Get输血执行登记', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1253 And 功能 = '基本' And Upper(对象) = Upper('Zl_Fun_Get输血执行登记'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1254, '基本', User, 'Zl_Fun_Get输血执行登记', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1254 And 功能 = '基本' And Upper(对象) = Upper('Zl_Fun_Get输血执行登记'));    

--92699:刘鹏飞,2016-01-22,病区标记设置权限增加
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1265,'病区标记设置',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'zl_病区标记内容_update','EXECUTE' From Dual Union All
Select 'zl_病区标记内容_insert','EXECUTE' From Dual Union All
Select 'zl_病区标记内容_delete','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--92808:马政,2016-01-22,卫材单据打印控制
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1712, '基本', User, '药品收发主表', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1712 And 功能 = '基本' And 对象 = '药品收发主表');
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1712, '基本', User, 'Zl_药品收发主表_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1712 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品收发主表_Insert'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1713, '基本', User, '药品收发主表', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1713 And 功能 = '基本' And 对象 = '药品收发主表');
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1713, '基本', User, 'Zl_药品收发主表_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1713 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品收发主表_Insert'));    

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1714, '基本', User, '药品收发主表', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1714 And 功能 = '基本' And 对象 = '药品收发主表');
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1714, '基本', User, 'Zl_药品收发主表_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1714 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品收发主表_Insert'));  

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1716, '基本', User, '药品收发主表', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1716 And 功能 = '基本' And 对象 = '药品收发主表');
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1716, '基本', User, 'Zl_药品收发主表_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1716 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品收发主表_Insert'));  

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1717, '基本', User, '药品收发主表', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1717 And 功能 = '基本' And 对象 = '药品收发主表');
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1717, '基本', User, 'Zl_药品收发主表_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1717 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品收发主表_Insert'));  
         
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1718, '基本', User, '药品收发主表', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1718 And 功能 = '基本' And 对象 = '药品收发主表');
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1718, '基本', User, 'Zl_药品收发主表_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1718 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品收发主表_Insert'));  

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1719, '基本', User, '药品收发主表', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1719 And 功能 = '基本' And 对象 = '药品收发主表');
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1719, '基本', User, 'Zl_药品收发主表_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1719 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品收发主表_Insert'));                    
         
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1722, '基本', User, '药品收发主表', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1722 And 功能 = '基本' And 对象 = '药品收发主表');
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1722, '基本', User, 'Zl_药品收发主表_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1722 And 功能 = '基本' And Upper(对象) = Upper('Zl_药品收发主表_Insert'));




-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------
-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--93917:许华峰,2016-03-20,与RIS的数据交换接口脚本
Create Or Replace Function Zlpub_Pacs_获取提纲内容
(
  医嘱id_In   In 病人医嘱报告.医嘱id%Type,
  报告提纲_In In 电子病历内容.内容文本%Type
) Return Varchar2 Is

  v_Result        Varchar2(4000);
  v_Singleresult  Varchar2(4000);
  v_Reportcontent 影像报告记录.报告内容%Type;
  n_Count         Number(2);

  Cursor Cur_Report_Contents Is
    Select 报告内容 From 影像报告记录 Where 医嘱id = 医嘱id_In;

Begin
  v_Result       := '';
  v_Singleresult := '';

  Select Count(1) Into n_Count From 影像报告记录 Where 医嘱id = 医嘱id_In;

  If n_Count > 0 Then
    For Row_Report_Contents In Cur_Report_Contents Loop
      v_Reportcontent := Row_Report_Contents.报告内容;
      Select Zlpub_Pacs_取提纲内容byxml(v_Reportcontent, 报告提纲_In) Into v_Singleresult From Dual;
      If v_Result Is Null And Not v_Singleresult Is Null Then
        v_Result := v_Singleresult;
      Elsif Not v_Singleresult Is Null Then
        v_Result := v_Result || ';' || v_Singleresult;
      End If;
    End Loop;
    Return v_Result;
  Else
    Select Count(1)
    Into n_Count
    From 电子病历内容 A, 电子病历内容 B, 病人医嘱报告 C
    Where a.对象类型 = 1 And a.Id = b.父id And b.对象类型 = 2 And b.开始版 = 0 And a.文件id = c.病历id And c.医嘱id = 医嘱id_In And
          a.内容文本 = 报告提纲_In;
  
    If n_Count > 0 Then
      Select b.内容文本
      Into v_Result
      From 电子病历内容 A, 电子病历内容 B, 病人医嘱报告 C
      Where a.对象类型 = 1 And a.Id = b.父id And b.对象类型 = 2 And b.开始版 = 0 And a.文件id = c.病历id And c.医嘱id = 医嘱id_In And
            a.内容文本 = 报告提纲_In;
    Else
      Select b.内容文本
      Into v_Result
      From 电子病历内容 A, 电子病历内容 B, 病人医嘱报告 C
      Where a.对象类型 = 3 And a.Id = b.父id And b.对象类型 = 2 And b.终止版 = 0 And a.文件id = c.病历id And c.医嘱id = 医嘱id_In And
            a.内容文本 = 报告提纲_In;
    End If;
  
    Return v_Result;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_获取提纲内容;
/

--93675:梁经伙,2016-03-11,该存储过程只用于诊断选择器中病案首页某个字段的改变用的
CREATE OR REPLACE Procedure Zl_病案主页_首页整理EX
(
  病人id_In  In 病案主页.病人id%Type,
  主页id_In  In 病案主页.主页id%Type,
  信息名_In  In varchar2,     /*病案主页的字段名*/
  信息值_In  In varchar2      /*病案主页某字段的值*/
) Is
/*该存储过程只用于诊断选择器中病案首页某个字段的改变用的*/
Begin
  If 信息名_In = '出院方式' Then
    update 病案主页 set 出院方式 = 信息值_In where 病人id = 病人id_In And 主页id = 主页id_In;
  elsif 信息名_In = '尸检标志' Then
     update 病案主页 set 尸检标志 = to_number(信息值_In) where 病人id = 病人id_In And 主页id = 主页id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案主页_首页整理EX;
/

--92729:胡俊勇,2016-03-18,专业版RIS接口
Create Or Replace Procedure Zl_Ris检查预约_Insert
(
  医嘱id_In     In Ris检查预约.医嘱id%Type,
  预约id_In     In Ris检查预约.预约id%Type,
  预约日期_In   In Ris检查预约.预约日期%Type,
  设备id_In     In Ris检查预约.检查设备id%Type,
  设备名称_In   In Ris检查预约.检查设备名称%Type,
  开始时间_In   In Ris检查预约.预约开始时间%Type,
  结束时间_In   In Ris检查预约.预约结束时间%Type,
  开始时间段_In In Ris检查预约.预约开始时间段%Type,
  结束时间段_In In Ris检查预约.预约结束时间段%Type
) Is
Begin
  Insert Into Ris检查预约
    (医嘱id, 预约id, 预约日期, 检查设备id, 检查设备名称, 预约开始时间, 预约结束时间, 预约开始时间段, 预约结束时间段)
  Values
    (医嘱id_In, 预约id_In, 预约日期_In, 设备id_In, 设备名称_In, 开始时间_In, 结束时间_In, 开始时间段_In, 结束时间段_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ris检查预约_Insert;
/

--92729:胡俊勇,2016-03-08,专业版RIS接口处理
Create Or Replace Procedure Zl_Ris检查预约_Delete(医嘱id_In In Ris检查预约.医嘱id%Type) Is
Begin
  Delete Ris检查预约 Where 医嘱id = 医嘱id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ris检查预约_Delete;
/

--92729:胡俊勇,2016-03-08,专业版RIS接口处理
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
                                         ',病人诊断记录_PK,病人医嘱状态_PK,医嘱签名记录_PK,病人医嘱发送_PK,诊疗单据打印_PK,医嘱执行计价_PK,执行打印记录_PK,RIS检查预约_PK' ||
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

--92729:胡俊勇,2016-03-08,专业版RIS接口处理
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
  
  Update /*+ rule*/ RIS检查预约
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

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

--93964:梁唐彬,2016-03-08,变动记录错误
Create Or Replace Procedure Zl_病人医嘱记录_停止
(
  --功能：停止指定的医嘱
  --说明：一并给药的只能调用一次
  --参数：ID_IN=相关ID为NULL的医嘱的ID(给药途径,中药用法,检查项目,主要手术,及独立医嘱)
  --      内部调用_IN=是否其他过程内部在调用，主要区分是否手工调用停止护理等级
  Id_In         病人医嘱记录.Id%Type,
  终止时间_In   病人医嘱记录.执行终止时间%Type,
  停嘱医生_In   病人医嘱记录.停嘱医生%Type,
  内部调用_In   Number := 0,
  医师资格_In   Number := 0,
  停嘱审核_In   Number := 0,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null
) Is
  v_状态       病人医嘱记录.医嘱状态%Type;
  v_医嘱内容   病人医嘱记录.医嘱内容%Type;
  v_护理等级id 病案主页.护理等级id%Type;
  v_病人id     病案主页.病人id%Type;
  v_主页id     病案主页.主页id%Type;
  v_婴儿       病人医嘱记录.婴儿%Type;
  v_诊疗类别   病人医嘱记录.诊疗类别%Type;
  v_操作类型   诊疗项目目录.操作类型%Type;

  Cursor c_Oldinfo Is
    Select b.*
    From (Select c.*
           From 病人变动记录 C
           Where c.病人id = v_病人id And c.主页id = v_主页id And
                 c.开始时间 = (Select Min(开始时间)
                           From 病人变动记录
                           Where 病人id = v_病人id And 主页id = v_主页id And 开始时间 > 终止时间_In) And
                 c.终止时间 = (Select Min(终止时间)
                           From 病人变动记录
                           Where 病人id = v_病人id And 主页id = v_主页id And 开始时间 > 终止时间_In)) A, 病人变动记录 B
    
    Where b.病人id = v_病人id And b.主页id = v_主页id And a.开始时间 = b.终止时间 And a.开始原因 = b.终止原因 And a.附加床位 = b.附加床位
    Union
    Select *
    From 病人变动记录
    Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null And 开始时间 <= 终止时间_In;

  Cursor c_Endinfo Is
    Select * From 病人变动记录 Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null;
  r_Oldinfo  c_Oldinfo%Rowtype;
  r_Endinfo  c_Endinfo%Rowtype;
  v_终止原因 病人变动记录.终止原因%Type;
  v_终止时间 病人变动记录.终止时间%Type;
  v_终止人员 病人变动记录.终止人员%Type;

  v_Date     Date;
  v_Count    Number;
  v_Temp     Varchar2(255);
  v_人员编号 病人变动记录.操作员编号%Type;
  v_人员姓名 病人变动记录.操作员姓名%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --检查医嘱状态是否正确:并发操作
  Select a.医嘱状态, a.医嘱内容, a.病人id, a.主页id, Nvl(a.婴儿, 0), Nvl(a.诊疗类别, '*') As 诊疗类别, b.操作类型
  Into v_状态, v_医嘱内容, v_病人id, v_主页id, v_婴儿, v_诊疗类别, v_操作类型
  From 病人医嘱记录 A, 诊疗项目目录 B
  Where a.诊疗项目id = b.Id(+) And a.Id = Id_In;
  If v_状态 In (4, 8, 9) Then
    v_Error := '医嘱"' || v_医嘱内容 || '"已经被作废或停止，不能再停止。';
    Raise Err_Custom;
  End If;

  --检查是否是输液配液记录，并是否已经锁定
  Select Count(1) Into v_Count From 输液配药记录 Where 是否锁定 = 1 And 医嘱ID = Id_In And 执行时间 > 终止时间_In;
  If v_Count > 0 Then
    v_Error := '医嘱"' || v_医嘱内容 || '"是输液药品，已经被输液配置中心锁定，不能停止。';
    Raise Err_Custom;
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

  --判断是否具有执业医师或助理医师的资格和参数是否勾选
  If 医师资格_In > 0 Or 停嘱审核_In = 0 Then
    Update 病人医嘱记录
    Set 医嘱状态 = 8, 执行终止时间 = 终止时间_In, 停嘱医生 = 停嘱医生_In, 停嘱时间 = v_Date, 审核标记 = Decode(审核标记, 2, 3, 审核标记)
    Where ID = Id_In Or 相关id = Id_In;
  
    Insert Into 病人医嘱状态
      (医嘱id, 操作类型, 操作人员, 操作时间)
      Select ID, 8, v_人员姓名, v_Date --护士停时记录为护士
      From 病人医嘱记录
      Where ID = Id_In Or 相关id = Id_In;
  Else
    --否则只修改审核标记，产生新的状态
    Update 病人医嘱记录 Set 审核标记 = 2 Where ID = Id_In Or 相关id = Id_In;
  
    Insert Into 病人医嘱状态
      (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
      Select ID, 13, v_人员姓名, v_Date, To_Char(终止时间_In, 'YYYY-MM-DD HH24:MI:SS') --护士停时记录为护士
      From 病人医嘱记录
      Where ID = Id_In Or 相关id = Id_In;
  End If;

  --其他特殊处理
  If Nvl(内部调用_In, 0) = 0 And (医师资格_In > 0 Or 停嘱审核_In = 0) Then
    --停止病情医嘱时，变动病人病情
    If v_诊疗类别 = 'Z' And v_操作类型 In ('9', '10') Then
      Open c_Oldinfo; --必须在处理之前先打开
      Fetch c_Oldinfo
        Into r_Oldinfo;
      Open c_Endinfo;
      Fetch c_Endinfo
        Into r_Endinfo;
      If c_Endinfo%Rowcount = 0 Then
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
    
      Update 病案主页 Set 当前病况 = '一般' Where 病人id = v_病人id And 主页id = v_主页id;
    
      --取消上次变动
      If r_Oldinfo.终止时间 Is Not Null Then
        v_终止时间 := r_Oldinfo.终止时间;
        v_终止原因 := r_Oldinfo.终止原因;
        v_终止人员 := r_Oldinfo.终止人员;
        --取消上次变动
        Update 病人变动记录
        Set 终止时间 = 终止时间_In, 终止原因 = 13, 终止人员 = v_人员姓名, 上次计算时间 = Null
        Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 = v_终止时间 And 终止原因 = v_终止原因;
        --更新将来的记录如果有停止到将来的则删除上次计算时间
        Update 病人变动记录
        Set 病情 = '一般', 上次计算时间 = Null
        Where 病人id = v_病人id And 主页id = v_主页id And 开始时间 > 终止时间_In;
      Else
        Update 病人变动记录
        Set 终止时间 = 终止时间_In, 终止原因 = 13, 终止人员 = v_人员姓名
        Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null;
      End If;
    
      While c_Oldinfo%Found Loop
        Insert Into 病人变动记录
          (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情, 操作员编号, 操作员姓名,
           终止时间, 终止原因, 终止人员)
        Values
          (病人变动记录_Id.Nextval, v_病人id, v_主页id, 终止时间_In, 13, r_Oldinfo.附加床位, r_Oldinfo.病区id, r_Oldinfo.科室id,
           r_Oldinfo.护理等级id, r_Oldinfo.床位等级id, r_Oldinfo.床号, r_Oldinfo.责任护士, r_Oldinfo.经治医师, r_Oldinfo.主治医师,
           r_Oldinfo.主任医师, '一般', v_人员编号, v_人员姓名, v_终止时间, v_终止原因, v_终止人员);
      
        Fetch c_Oldinfo
          Into r_Oldinfo;
      End Loop;
    
      Close c_Oldinfo;
      Close c_Endinfo;
    Elsif v_诊疗类别 = 'H' And v_操作类型 = '1' And v_婴儿 = 0 Then
      --停止护理等级时，同时取消病人的护理等级费用
      Begin
        Select c.收费细目id
        Into v_护理等级id
        From 病人医嘱记录 A, 病人医嘱计价 C, 收费项目目录 D
        Where a.Id = c.医嘱id And c.收费细目id = d.Id And d.类别 = 'H' And Nvl(d.项目特性, 0) <> 0 And a.Id = Id_In And Rownum = 1 And
              Exists
         (Select 1 From 病案主页 Where 病人id = v_病人id And 主页id = v_主页id And 护理等级id = c.收费细目id);
      Exception
        When Others Then
          Null;
      End;
      If v_护理等级id Is Not Null Then
        --变动记录的时间加上秒，以便回退操作时区分同一分种的校对、停止等操作
        v_Date := To_Date(To_Char(终止时间_In, 'yyyy-mm-dd hh24:mi') || To_Char(v_Date, 'ss'), 'yyyy-mm-dd hh24:mi:ss');
        Zl_病人变动记录_Nurse(v_病人id, v_主页id, Null, v_Date, v_人员编号, v_人员姓名);
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人医嘱记录_停止;
/

--93964:梁唐彬,2016-03-08,变动记录错误问题
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
                 c.终止时间 = (Select Min(终止时间)
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
  v_参数值   Zlparameters.参数值%Type;
  v_Count    Number;
  v_Date     Date;
  v_Temp     Varchar2(255);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;

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
    Select zl_GetSysParameter(25) Into v_参数值 From Dual;
    If Nvl(v_参数值, '0') <> '0' Then
      Select zl_GetSysParameter(26) Into v_参数值 From Dual;
      If v_前提id Is Null Then
        Select Count(*) Into v_Count From 电子签名启用部门 Where 部门id = v_开嘱科室id And 场合 = 1;
      Else
        Select Count(*) Into v_Count From 电子签名启用部门 Where 部门id = v_开嘱科室id And 场合 = 3;
      End If;
      If Nvl(Substr(v_参数值, 2, 1), '0') = '1' And v_前提id Is Null And v_Count > 0 Or
         Nvl(Substr(v_参数值, 3, 1), '0') = '1' And v_前提id Is Not Null And v_Count > 0 Then
        Select Nvl(Max(是否停用), 0)
        Into v_Count
        From (Select a.是否停用, a.注册时间
               From 人员证书记录 A, 人员表 B
               Where a.人员id = b.Id And b.姓名 = v_开嘱医生
               Order By a.注册时间 Desc)
        Where Rownum < 2;
        If v_Count = 0 Then
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
      End If;
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
        Set 终止时间 = v_开始时间, 终止原因 = 13, 终止人员 = v_人员姓名
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

--93964:梁唐彬,2016-03-07,病况变动记录停止错误问题
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
          
          If r_Rolladvice.类别 = 'Z' And  Instr(',9,10,', Nvl(r_Rolladvice.类型, '0')) > 0 And Nvl(r_Rolladvice.婴儿, 0) = 0 Then 
            --回退病况医嘱时，调用变动记录回退
            Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, Null, Null, Null, '病况变动'); 
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

--93917:许华峰,2016-03-16,与RIS的数据交换接口脚本
Create Or Replace Package b_zlXWInterface Is
  Type t_Refcur Is Ref Cursor;

  --1、接收RIS状态改变
  Procedure ReceiveRisState
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    状态_In     Number,
    操作人员_In 病人医嘱发送.完成人%Type,
    执行时间_In 病人医嘱发送.完成时间%Type := Null,
    执行说明_In 病人医嘱发送.执行说明%Type := Null,
    单独执行_In Number := 0
  );

  --2、费用确认
  Procedure 影像费用执行
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  );

  --3、取消费用确认
  Procedure 影像费用执行_Cancel
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  );

  --4、接收RIS的报告
  Procedure ReceiveReport
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    报告所见_In 电子病历内容.内容文本%Type,
    报告意见_In 电子病历内容.内容文本%Type,
    报告建议_In 电子病历内容.内容文本%Type,
    报告医生_In 电子病历记录.创建人%Type
  );

  --5、修改申请单信息
  Procedure 影像病人信息_修改
  (
    医嘱id_In       病人医嘱记录.Id%Type,
    姓名_In         病人信息.姓名%Type,
    性别_In         病人信息.性别%Type,
    年龄_In         病人信息.年龄%Type,
    费别_In         病人信息.费别%Type,
    医疗付款方式_In 病人信息.医疗付款方式%Type,
    民族_In         病人信息.民族%Type,
    婚姻_In         病人信息.婚姻状况%Type,
    职业_In         病人信息.职业%Type,
    身份证号_In     病人信息.身份证号%Type,
    家庭地址_In     病人信息.家庭地址%Type,
    家庭电话_In     病人信息.家庭电话%Type,
    家庭地址邮编_In 病人信息.家庭地址邮编%Type,
    出生日期_In     病人信息.出生日期%Type := Null
  );

  --6、取消申请单信息
  Procedure 取消检查申请单
  (
    医嘱id_In     病人医嘱执行.医嘱id%Type,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := 0,
    拒绝原因_In   病人医嘱发送.执行说明%Type := Null
  );

End b_zlXWInterface;
/

Create Or Replace Package Body b_zlXWInterface Is

  --1、接收RIS状态改变
  Procedure ReceiveRisState
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    状态_In     Number,
    操作人员_In 病人医嘱发送.完成人%Type,
    执行时间_In 病人医嘱发送.完成时间%Type := Null,
    执行说明_In 病人医嘱发送.执行说明%Type := Null,
    单独执行_In Number := 0
  ) Is
  
    --参数：医嘱ID_IN - 单独执行的医嘱ID。
    --      状态_IN - -1-删除；0-预约；1-登记；3-检查完成；4-检查中止；9-初步报告；12-报告审核；15-发放
    --     单独执行_In -0-全部执行；1-单独执行；检查医嘱组合是否采用对每个项目分散单独执行的方式
  
    Cursor c_Adviceinfo Is
      Select ID, 相关id, Nvl(相关id, ID) As 组id, 诊疗类别, 病人来源, 执行科室id
      From 病人医嘱记录
      Where ID = 医嘱id_In;
    r_Adviceinfo c_Adviceinfo%RowType;
  
    v_Count Number;
    v_Error Varchar2(255);
    Err_Custom Exception;
    v_执行状态 病人医嘱发送.执行状态%Type;
    v_执行过程 病人医嘱发送.执行过程%Type;
  
  Begin
  
    v_执行状态 := 0;
    v_执行过程 := 0;
  
    --提取医嘱的主医嘱ID，及组ID
    Open c_Adviceinfo;
    Fetch c_Adviceinfo
      Into r_Adviceinfo;
    Close c_Adviceinfo;
  
    --先检查是否已经出院的住院病人，已经预出院或者出院的检查申请，不允许开始检查和审核费用
    Select Count(*)
    Into v_Count
    From 病人医嘱记录 A, 病案主页 B
    Where a.病人id = b.病人id And a.主页id = b.主页id And (b.出院日期 Is Not Null Or b.状态 = 3) And a.Id = r_Adviceinfo.组id;
  
    If v_Count > 0 Then
      v_Error := '住院病人已经出院或者预出院，无法开始检查。';
      Raise Err_Custom;
    End If;
  
    --根据状态_IN执行医嘱
    ---1-删除；0-预约；1-登记；3-检查完成；4-检查中止；9-初步报告；12-报告审核；15-发放
  
    If 状态_In = -1 Or 状态_In = 0 Then
      v_执行状态 := 0; --未执行
      v_执行过程 := 0;
    Elsif 状态_In = 1 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 2; --已报到
    Elsif 状态_In = 3 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 3; --已检查
    Elsif 状态_In = 4 Then
      --不改变
      v_执行状态 := v_执行状态;
    Elsif 状态_In = 9 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 4; --已报告
    Elsif 状态_In = 12 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 5; --已审核
    Elsif 状态_In = 15 Then
      v_执行状态 := 1; --完全执行
      v_执行过程 := 6; --已完成
    End If;
  
    --开始执行医嘱
    If Nvl(单独执行_In, 0) = 1 Then
      -- 单个部位医嘱单独执行
      Update 病人医嘱发送
      Set 执行状态 = v_执行状态, 执行过程 = v_执行过程, 执行说明 = 执行说明_In
      Where 医嘱id = 医嘱id_In;
    Else
      Update 病人医嘱发送
      Set 执行状态 = v_执行状态, 执行过程 = v_执行过程, 执行说明 = 执行说明_In
      Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = r_Adviceinfo.组id Or 相关id = r_Adviceinfo.组id));
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receiverisstate;

  --2、费用确认
  Procedure 影像费用执行
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  ) Is
    --参数：医嘱ID_IN=单独执行的医嘱ID。
    --      单独执行_In=检查医嘱组合是否采用对每个项目分散单独执行的方式,0-不单独执行
    Cursor c_Advice Is
      Select ID, 相关id, Nvl(相关id, ID) As 组id, 诊疗类别, 病人来源 From 病人医嘱记录 Where ID = 医嘱id_In;
    r_Advice c_Advice%RowType;
  
    v_Temp     Varchar2(255);
    v_人员编号 人员表.编号%Type;
    v_人员姓名 人员表.姓名%Type;
    v_部门id   部门表.Id%Type;
    v_费用性质 病人医嘱发送.记录性质%Type;
    v_发送号   病人医嘱发送.发送号%Type;
    v_执行过程 病人医嘱发送.执行过程%Type;
  Begin
  
    Select 发送号, 执行过程 Into v_发送号, v_执行过程 From 病人医嘱发送 Where 医嘱id = 医嘱id_In;
  
    --登记和完成才执行费用  2-登记，3-检查，4-报告，5-审核，6-完成
    If v_执行过程 >= 2 Or v_执行过程 <= 6 Then
    
      --取当前操作人员
      If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null And 执行部门id_In Is Not Null Then
        v_人员编号 := 操作员编号_In;
        v_人员姓名 := 操作员姓名_In;
        v_部门id   := 执行部门id_In;
      Else
        v_Temp     := Zl_Identity;
        v_部门id   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      End If;
    
      --取主医嘱ID
      Open c_Advice;
      Fetch c_Advice
        Into r_Advice;
      Close c_Advice;
    
      If r_Advice.病人来源 = 2 Then
        Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
        Into v_费用性质
        From 病人医嘱发送
        Where 发送号 = v_发送号 And 医嘱id = 医嘱id_In;
      Else
        v_费用性质 := 1;
      End If;
    
      --执行费用和自动发料
      If v_费用性质 = 1 Then
        Zl_门诊医嘱执行_Finish(医嘱id_In, v_发送号, 单独执行_In, v_人员编号, v_人员姓名, r_Advice.组id, r_Advice.诊疗类别, v_部门id);
      Else
        Zl_住院医嘱执行_Finish(医嘱id_In, v_发送号, 单独执行_In, v_人员编号, v_人员姓名, r_Advice.组id, r_Advice.诊疗类别, v_部门id);
      End If;
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像费用执行;

  --3、取消费用确认
  Procedure 影像费用执行_Cancel
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  ) Is
    --参数：
    --      医嘱ID_IN=单独执行的医嘱ID。
    --      单独执行_In=检查医嘱组合是否采用对每个项目分散单独执行的方式,0-不单独执行
  
    v_发送号 病人医嘱发送.发送号%Type;
  Begin
  
    Select 发送号 Into v_发送号 From 病人医嘱发送 Where 医嘱id = 医嘱id_In;
  
    --调用统一的医嘱执行Cancel过程
    Zl_病人医嘱执行_Cancel(医嘱id_In, v_发送号, Null, 单独执行_In, 执行部门id_In, 操作员编号_In, 操作员姓名_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像费用执行_Cancel;

  --4、接收RIS的报告
  Procedure ReceiveReport
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    报告所见_In 电子病历内容.内容文本%Type,
    报告意见_In 电子病历内容.内容文本%Type,
    报告建议_In 电子病历内容.内容文本%Type,
    报告医生_In 电子病历记录.创建人%Type
  ) Is
  
    --提取病人医嘱及报告的相关信息
    Cursor c_Advice(v_组id Number) Is
      Select e.Id, e.病人来源, e.病人id, e.主页id, e.婴儿, e.病人科室id, e.文件id, e.病历种类, e.病历名称, f.病历id, e.执行科室id
      From (Select c.Id, c.病人来源, c.病人id, c.主页id, c.婴儿, c.病人科室id, c.文件id, d.种类 病历种类, d.名称 病历名称, c.执行科室id
             From (Select a.Id, a.病人来源, a.病人id, a.主页id, a.婴儿, a.病人科室id, b.病历文件id 文件id, a.执行科室id
                    From 病人医嘱记录 A, 病历单据应用 B
                    Where a.Id = v_组id And a.诊疗项目id = b.诊疗项目id(+) And b.应用场合(+) = Decode(a.病人来源, 2, 2, 4, 4, 1)) C,
                  病历文件列表 D
             Where c.文件id = d.Id(+)) E, 病人医嘱报告 F
      Where e.Id = f.医嘱id(+);
  
    --查找文件的组成元素
    Cursor c_File(v_File Number) Is
      Select a.Id, a.文件id, a.父id, a.对象序号, a.对象类型, a.对象标记, a.保留对象, a.对象属性, a.内容行次, a.内容文本, a.是否换行, a.预制提纲id, a.复用提纲,
             a.使用时机, a.诊治要素id, a.替换域, a.要素名称, a.要素类型, a.要素长度, a.要素小数, a.要素单位, a.要素表示, a.输入形态, a.要素值域
      From 病历文件结构 A
      Where a.文件id = v_File
      Order By a.对象序号;
  
    Cursor c_Report(v_电子病历记录id Number) Is
      Select /*+ rule */
       b.Id, a.内容文本
      From 电子病历内容 A, 电子病历内容 B
      Where a.文件id = v_电子病历记录id And Nvl(a.定义提纲id, 0) <> 0 And (a.内容文本 Like '%所见%' Or a.内容文本 Like '%意见%' Or a.内容文本 Like '%建议%') And
            b.父id = a.Id And b.是否换行 = 1;
  
    r_Advice        c_Advice%RowType;
    v_病历id        电子病历内容.文件id%Type;
    v_病历内容id    电子病历内容.Id%Type;
    v_病历内容idnew 电子病历内容.Id%Type;
    v_对象序号      电子病历内容.对象序号%Type;
    v_父id          电子病历内容.父id%Type;
    v_内容文本      电子病历内容.内容文本%Type;
    v_定义提纲id    电子病历内容.定义提纲id%Type;
    --v_格式内容    电子病历格式.内容%Type;
    v_Error Varchar2(255);
    Err_Custom Exception;
    v_主医嘱id 病人医嘱发送.医嘱id%Type;
  Begin
  
    -- 提取主医嘱ID ，防止因为传入部位医嘱，导致报告保存出错
    Select Nvl(相关id, ID) As 组id Into v_主医嘱id From 病人医嘱记录 Where ID = 医嘱id_In;
  
    Open c_Advice(v_主医嘱id);
    Fetch c_Advice
      Into r_Advice;
  
    If Nvl(r_Advice.文件id, 0) = 0 Then
      v_Error := '本次检查项目没有对应相关的检查报告，请与管理员联系！';
      Raise Err_Custom;
    Else
      If Nvl(r_Advice.病历id, 0) > 0 Then
        ----产生过报告
        --找出检查已填写的报告提纲中含有"%所见%","%描述%","%建议%","%意见%",并用传入的参数更新
        For r_Report In c_Report(r_Advice.病历id) Loop
          If r_Report.内容文本 Like '%所见%' Then
            Update 电子病历内容 Set 内容文本 = 报告所见_In Where ID = r_Report.Id;
          Elsif r_Report.内容文本 Like '%意见%' Then
            Update 电子病历内容 Set 内容文本 = 报告意见_In Where ID = r_Report.Id;
          Elsif r_Report.内容文本 Like '%建议%' Then
            Update 电子病历内容 Set 内容文本 = 报告建议_In Where ID = r_Report.Id;
          End If;
        End Loop;
        --更新保存时间
        Update 电子病历记录
        Set 完成时间 = Sysdate, 保存人 = 报告医生_In, 保存时间 = Sysdate
        Where ID = r_Advice.病历id;
      Else
        --产生电子病历记录
        Select 电子病历记录_Id.Nextval Into v_病历id From Dual;
        Insert Into 电子病历记录
          (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 完成时间, 保存人, 保存时间, 最后版本, 签名级别)
        Values
          (v_病历id, r_Advice.病人来源, r_Advice.病人id, r_Advice.主页id, r_Advice.婴儿, r_Advice.病人科室id, r_Advice.病历种类,
           r_Advice.文件id, r_Advice.病历名称, 报告医生_In, Sysdate, Sysdate, 报告医生_In, Sysdate, 1, 2);
      
        --产生医嘱报告记录
        Insert Into 病人医嘱报告 (医嘱id, 病历id) Values (v_主医嘱id, v_病历id);
      
	    v_对象序号 := 0;
		
        --新产生报告内容
        For r_File In c_File(r_Advice.文件id) Loop
          Select 电子病历内容_Id.Nextval Into v_病历内容id From Dual;
        
          If Nvl(r_File.父id, 0) <> 0 And (r_File.内容文本 Like '%所见%') Then
            --所见定义行(非提纲)
            v_内容文本   := 报告所见_In || Chr(13) || Chr(13);
            v_定义提纲id := 0;
          Elsif Nvl(r_File.父id, 0) <> 0 And (r_File.内容文本 Like '%意见%') Then
            --意见定义行(非提纲)
            v_内容文本   := 报告意见_In || Chr(13) || Chr(13);
            v_定义提纲id := 0;
          Elsif Nvl(r_File.父id, 0) <> 0 And (r_File.内容文本 Like '%建议%') Then
            --建议定义行(非提纲)
            v_内容文本   := 报告建议_In || Chr(13) || Chr(13);
            v_定义提纲id := 0;
          Elsif Nvl(r_File.对象类型, 0) = 1 And Nvl(r_File.父id, 0) = 0 Then
            --提纲定义行
            v_父id       := v_病历内容id;
            v_内容文本   := r_File.内容文本;
            v_定义提纲id := r_File.Id;
          Elsif Nvl(r_File.对象类型, 0) = 4 And r_File.要素名称 Is Not Null Then
            --自动替换要素
            v_内容文本   := Zl_Replace_Element_Value(r_File.要素名称, r_Advice.病人id, r_Advice.主页id, r_Advice.病人来源, r_Advice.Id);
            v_定义提纲id := 0;
          Else
            v_内容文本   := r_File.内容文本;
            v_定义提纲id := 0;
          End If;
        
          --报告内容单独写一行
          If Nvl(r_File.父id, 0) <> 0 And (r_File.内容文本 Like '%所见%' Or r_File.内容文本 Like '%意见%' Or r_File.内容文本 Like '%建议%') Then
            --先写提纲显示名称，再写内容，同时对象序号发生变化
            Select 电子病历内容_Id.Nextval Into v_病历内容idnew From Dual;
            v_对象序号 := v_对象序号 + 1;
            Insert Into 电子病历内容
              (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域,
               要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
            Values
              (v_病历内容idnew, v_病历id, 0, 0, Decode(v_定义提纲id, 0, v_父id, Null), v_对象序号, r_File.对象类型, r_File.对象标记,
               r_File.保留对象, 0, Null, v_内容文本, r_File.是否换行, r_File.预制提纲id, r_File.复用提纲, r_File.使用时机, r_File.诊治要素id,
               r_File.替换域, r_File.要素名称, r_File.要素类型, r_File.要素长度, r_File.要素小数, r_File.要素单位, r_File.要素表示, r_File.输入形态,
               r_File.要素值域, Decode(v_定义提纲id, 0, Null, v_定义提纲id));
            
            v_内容文本 := r_File.内容文本;
          End If;
        
		  v_对象序号 := v_对象序号 + 1;
		  
          Insert Into 电子病历内容
            (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域,
             要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
          Values
            (v_病历内容id, v_病历id, 1, 0, Decode(v_定义提纲id, 0, v_父id, Null), v_对象序号, r_File.对象类型, r_File.对象标记, r_File.保留对象,
             r_File.对象属性, Null, v_内容文本, r_File.是否换行, r_File.预制提纲id, r_File.复用提纲, r_File.使用时机, r_File.诊治要素id, r_File.替换域,
             r_File.要素名称, r_File.要素类型, r_File.要素长度, r_File.要素小数, r_File.要素单位, r_File.要素表示, r_File.输入形态, r_File.要素值域,
             Decode(v_定义提纲id, 0, Null, v_定义提纲id));
        End Loop;
      
        --因电子病历格式中含了内容文字格式，此种方法导入之后内容文字将不可见
        --Select 内容 Into v_格式内容 From 病历文件格式 Where 文件ID=r_Advice.文件ID;
        --Insert Into 电子病历格式 (文件ID,内容) Values (v_病历id,v_格式内容);
      
      End If;
    End If;
    Close c_Advice;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receivereport;

  --5、修改申请单信息
  Procedure 影像病人信息_修改
  (
    医嘱id_In       病人医嘱记录.Id%Type,
    姓名_In         病人信息.姓名%Type,
    性别_In         病人信息.性别%Type,
    年龄_In         病人信息.年龄%Type,
    费别_In         病人信息.费别%Type,
    医疗付款方式_In 病人信息.医疗付款方式%Type,
    民族_In         病人信息.民族%Type,
    婚姻_In         病人信息.婚姻状况%Type,
    职业_In         病人信息.职业%Type,
    身份证号_In     病人信息.身份证号%Type,
    家庭地址_In     病人信息.家庭地址%Type,
    家庭电话_In     病人信息.家庭电话%Type,
    家庭地址邮编_In 病人信息.家庭地址邮编%Type,
    出生日期_In     病人信息.出生日期%Type := Null
  ) As
  
    v_年龄     Varchar2(20);
    v_年龄单位 Varchar2(20);
    v_出生日期 Date;
    v_病人来源 病人医嘱记录.病人来源%Type;
    v_病人id   病人医嘱记录.病人id%Type;
  Begin
    Begin
      Select 病人来源, 病人id Into v_病人来源, v_病人id From 病人医嘱记录 Where ID = 医嘱id_In;
    Exception
      When Others Then
        Return;
    End;
  
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
  
    If v_病人来源 = 3 Then
      Update 病人信息
      Set 姓名 = 姓名_In, 性别 = Nvl(性别_In, 性别), 年龄 = 年龄_In, 出生日期 = v_出生日期, 费别 = Nvl(费别_In, 费别),
          医疗付款方式 = Nvl(医疗付款方式_In, 医疗付款方式), 民族 = Nvl(民族_In, 民族), 婚姻状况 = Nvl(婚姻_In, 婚姻状况), 职业 = Nvl(职业_In, 职业),
          身份证号 = 身份证号_In, 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In, 家庭地址邮编 = 家庭地址邮编_In
      Where 病人id = v_病人id;
    
      --修改对应的医嘱记录
      Update 病人医嘱记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In
      Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
    Else
      Update 病人信息
      Set 民族 = Nvl(民族_In, 民族), 婚姻状况 = Nvl(婚姻_In, 婚姻状况), 职业 = Nvl(职业_In, 职业), 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In,
          家庭地址邮编 = 家庭地址邮编_In
      Where 病人id = v_病人id;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像病人信息_修改;

  --6、取消申请单信息
  Procedure 取消检查申请单
  (
    医嘱id_In     病人医嘱执行.医嘱id%Type,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := 0,
    拒绝原因_In   病人医嘱发送.执行说明%Type := Null
  ) As
    --参数：医嘱ID_IN=单独执行的医嘱ID
  
    v_发送号 病人医嘱执行.发送号%Type;
  
  Begin
  
    Begin
      Select 发送号 Into v_发送号 From 病人医嘱发送 Where 医嘱id = 医嘱id_In;
    Exception
      When Others Then
        Return;
    End;
  
    Zl_病人医嘱执行_拒绝执行(医嘱id_In, v_发送号, 操作员编号_In, 操作员姓名_In, 执行部门id_In, 拒绝原因_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 取消检查申请单;
end b_zlXWInterface;
/

--93623:梁唐彬,2016-03-04,检查医嘱计费状态调整
Create Or Replace Procedure Zl_病人医嘱发送_Insert
(
  医嘱id_In     病人医嘱发送.医嘱id%Type,
  发送号_In     病人医嘱发送.发送号%Type,
  记录性质_In   病人医嘱发送.记录性质%Type,
  No_In         病人医嘱发送.No%Type,
  记录序号_In   病人医嘱发送.记录序号%Type,
  发送数次_In   病人医嘱发送.发送数次%Type,
  首次时间_In   病人医嘱发送.首次时间%Type,
  末次时间_In   病人医嘱发送.末次时间%Type,
  发送时间_In   病人医嘱发送.发送时间%Type,
  执行状态_In   病人医嘱发送.执行状态%Type,
  执行部门id_In 病人医嘱发送.执行部门id%Type,
  计费状态_In   病人医嘱发送.计费状态%Type,
  First_In      Number := 0,
  样本条码_In   病人医嘱发送.样本条码%Type := Null,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null,
  领药号_In     未发药品记录.领药号%Type := Null,
  门诊记帐_In   病人医嘱发送.门诊记帐%Type := Null,
  分解时间_In   Varchar2 := Null
  --功能：填写病人医嘱发送记录
  --参数：
  --      医嘱id_In=要发送的每个医嘱ID
  --      First_IN=表示是否一组医嘱的第一医嘱行,以便处理医嘱相关内容(如成药,配方的第一行,因为给药途径,配方煎法,用法可能为叮嘱不发送)
  --      发送数次_IN,首次时间_IN,末次时间_IN:对"持续性"长嘱,不填写发送数次,可填写首末次时间(用于回退)。
  --      门诊记帐_In,住院临嘱发送到门诊记帐时才填写为1（因为记录性质是2，用于区分住院记帐），其余情况均填空。
) Is
  --包含病人及医嘱(一组医嘱中第一行)相关信息的游标
  Cursor c_Advice Is
    Select Nvl(a.相关id, a.Id) As 组id, a.序号, a.病人id, a.主页id, a.婴儿, a.姓名, a.病人科室id, c.操作类型, a.诊疗类别, a.医嘱期效, a.医嘱状态, a.医嘱内容,
           a.开嘱医生, a.开嘱时间, a.开始执行时间, a.上次执行时间, a.执行终止时间, a.执行时间方案, a.频率次数, a.频率间隔, a.间隔单位, a.开嘱科室id, a.标本部位, a.执行科室id,
           a.相关id, a.诊疗项目id
    From 病人医嘱记录 A, 诊疗项目目录 C
    Where a.诊疗项目id = c.Id And a.Id = 医嘱id_In;
  r_Advice c_Advice%RowType;

  --包含病人(婴儿)的所有未停长嘱(含配方长嘱),婴儿传入-1表示都处理
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
  --包含病人(婴儿)的已停但未确认的长嘱,终止执行时间在指定时间之后,婴儿传入-1表示都处理
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

  --其它临时变量
  v_婴儿     病人医嘱记录.婴儿%Type;
  v_持续性   Number(1); --是否持续性长嘱
  v_Autostop Number(1);
  v_Date     Date;
  v_Temp     Varchar2(255);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_停止时间 病人医嘱记录.开嘱时间%Type;
  n_执行状态 病人医嘱发送.执行状态%Type;
  d_开始时间 病人医嘱记录.开始执行时间%Type;

  v_Stopadviceids 病人医嘱记录.医嘱内容%Type;
  n_Adviceid      病人医嘱记录.病人id%Type;
  n_标记          Number(18);
  v_Error         Varchar2(255);
  Err_Custom Exception;
Begin
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
  --如果首次时间为空则填入开始执行时间
  If 首次时间_In Is Null Or 分解时间_In Is Null Or 末次时间_In Is Null Then
    Select 开始执行时间 Into d_开始时间 From 病人医嘱记录 Where ID = 医嘱id_In;
  End If;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;
  Close c_Advice;

  --是一组医嘱的第一行时处理医嘱内容
  If Nvl(First_In, 0) = 1 Then
    --并发操作检查
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.医嘱状态, 0) = 4 Then
      --检查要发送的医嘱是否被作废
      v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被其他人作废。' || Chr(13) || Chr(10) ||
                 '该病人的医嘱发送失败。请重新读取发送清单再试。';
      Raise Err_Custom;
    End If;
  
    If Nvl(r_Advice.医嘱期效, 0) = 0 Then
      --长嘱：含成药长嘱,配方长嘱,非药"可选频率"长嘱,非药"持续性"长嘱
    
      --检查长嘱是否已被发送
      If r_Advice.上次执行时间 Is Not Null Then
        If r_Advice.上次执行时间 >= 首次时间_In Then
          v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被其他人发送。' || Chr(13) || Chr(10) ||
                     '该病人的医嘱发送失败。请重新读取发送清单再试。';
          Raise Err_Custom;
        End If;
      End If;
    
      --检查长嘱发送前是否已被自动停止(如术后)
      If r_Advice.执行终止时间 Is Not Null Then
        If 首次时间_In > r_Advice.执行终止时间 Then
          v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被停止。' || Chr(13) || Chr(10) ||
                     '该病人的医嘱发送失败。请重新读取发送清单再试。';
          Raise Err_Custom;
        End If;
      End If;
    Elsif Nvl(r_Advice.医嘱状态, 0) In (8, 9) Then
      --临嘱：含配方临嘱
    
      --检查是否已被发送(或因其它原因自动停止)
      v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被其他人发送。' || Chr(13) || Chr(10) ||
                 '该病人的医嘱发送失败。请重新读取发送清单再试。';
      Raise Err_Custom;
    End If;
  
    --发送后的医嘱处理
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.医嘱期效, 0) = 0 Then
      --长期医嘱:更新上次执行时间
      Update 病人医嘱记录 Set 上次执行时间 = 末次时间_In Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
    
      --判断是否持续性长嘱
      v_持续性 := 0;
      If r_Advice.执行时间方案 Is Null And (Nvl(r_Advice.频率次数, 0) = 0 Or Nvl(r_Advice.频率间隔, 0) = 0 Or r_Advice.间隔单位 Is Null) Then
        v_持续性 := 1;
      End If;
    
      --预定了终止时间且未停止的自动停止
      If r_Advice.执行终止时间 Is Not Null And Nvl(r_Advice.医嘱状态, 0) Not In (8, 9) Then
        v_Autostop := 0;
        If v_持续性 = 1 Then
          --非药"持续性"长嘱
          If Trunc(末次时间_In) = Trunc(r_Advice.执行终止时间 - 1) Then
            v_Autostop := 1; --终止这天不执行
          End If;
        Elsif Zl_Advicenexttime(医嘱id_In) > r_Advice.执行终止时间 Then
          --成药长嘱或非药"可选频率"长嘱
          v_Autostop := 1; --如果是等于,还可以执行一次
        End If;
      
        If v_Autostop = 1 Then
          Update 病人医嘱记录
          Set 医嘱状态 = 8, 停嘱时间 = 末次时间_In, 停嘱医生 = r_Advice.开嘱医生
          Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
        
          Insert Into 病人医嘱状态
            (医嘱id, 操作类型, 操作人员, 操作时间)
            Select ID, 8, r_Advice.开嘱医生, 发送时间_In
            From 病人医嘱记录
            Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Advice.组id;
        End If;
      End If;
    Else
      --临嘱停止。
      --住院医生发送时自动校对、停止：校对是以Sysdate取的,为避免重复,停止时间也取Sysdate
      Select Sysdate Into v_Date From Dual;
      Update 病人医嘱记录
      Set 医嘱状态 = 8, 执行终止时间 = 末次时间_In,
          --为一次性临嘱时没有
          上次执行时间 = 末次时间_In,
          --为一次性临嘱时没有
          停嘱时间 = v_Date,
          --发送时间_IN,
          停嘱医生 = r_Advice.开嘱医生
      Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
    
      Insert Into 病人医嘱状态
        (医嘱id, 操作类型, 操作人员, 操作时间)
        Select ID, 8, v_人员姓名, v_Date --发送时间_IN
        From 病人医嘱记录
        Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
    End If;
  
    --特殊医嘱的处理
    ---------------------------------------------------------------------------------------
    If r_Advice.诊疗类别 = 'Z' And Nvl(r_Advice.操作类型, '0') <> '0' Then
      --(1-留观;2-住院;)3-转科;4-术后(不发送);5-出院;6-转院,7-会诊,11-死亡
    
      --几种特殊医嘱要自动停止病人该医嘱之前(按时间算)所有未停的长嘱
      If r_Advice.操作类型 In ('3', '5', '6', '11') Then
        If Nvl(r_Advice.婴儿, 0) = 0 Then
          v_婴儿 := -1;
        Else
          v_婴儿 := Nvl(r_Advice.婴儿, 0);
        End If;
        For r_Needstop In c_Needstop(r_Advice.病人id, r_Advice.主页id, v_婴儿, r_Advice.开始执行时间) Loop
          Select Decode(Sign(开始执行时间 - r_Advice.开始执行时间), 1, 开始执行时间, r_Advice.开始执行时间)
          Into v_停止时间
          From 病人医嘱记录
          Where ID = r_Needstop.Id;
          Update 病人医嘱记录
          Set 医嘱状态 = 8, 执行终止时间 = v_停止时间, 停嘱时间 = 发送时间_In, 停嘱医生 = r_Advice.开嘱医生
          Where ID = r_Needstop.Id;
        
          Insert Into 病人医嘱状态
            (医嘱id, 操作类型, 操作人员, 操作时间)
            Select ID, 8, v_人员姓名, 发送时间_In From 病人医嘱记录 Where ID = r_Needstop.Id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Needstop.Id;
        End Loop;
        --已停止未确认的长嘱,终止时间在医嘱开始后的,调前其终止时间(同时多个特殊医嘱的情况)
        For r_Havestop In c_Havestop(r_Advice.病人id, r_Advice.主页id, v_婴儿, r_Advice.开始执行时间) Loop
          Update 病人医嘱记录
          Set 执行终止时间 = Decode(Sign(开始执行时间 - r_Advice.开始执行时间), 1, 开始执行时间, r_Advice.开始执行时间), 停嘱时间 = 发送时间_In,
              停嘱医生 = r_Advice.开嘱医生
          Where ID = r_Havestop.Id;
        
          --不修改停止医嘱的操作人员，因为停止时，医生可能已进行电子签名
          Update 病人医嘱状态 Set 操作时间 = 发送时间_In Where 医嘱id = r_Havestop.Id And 操作类型 = 8;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Havestop.Id;
        End Loop;
        --处理长期备用医嘱(没有执行（发送）过的标记未用）,同时处理临嘱
        Update 病人医嘱记录
        Set 执行标记 = -1
        Where 病人id = r_Advice.病人id And 主页id = r_Advice.主页id And
              (医嘱期效 = 0 And 执行频次 = '必要时' And 上次执行时间 Is Null And 医嘱状态 In (3, 5, 6, 7) Or
              医嘱期效 = 1 And 执行频次 = '需要时' And 医嘱状态 = 3) And 执行标记 <> -1;
      End If;
    
      --具体的特殊处理
      If Nvl(r_Advice.婴儿, 0) = 0 Then
        If r_Advice.操作类型 = '3' And 执行部门id_In Is Not Null And r_Advice.病人科室id Is Not Null And
           Nvl(r_Advice.病人科室id, 0) <> Nvl(执行部门id_In, 0) Then
          --转科医嘱,将病人登记转科到"执行科室ID"(在院病人且当前科室与转入科室不同才处理)
          Zl_病人变动记录_Change(r_Advice.病人id, r_Advice.主页id, 执行部门id_In, v_人员编号, v_人员姓名);
        Elsif r_Advice.操作类型 In ('5', '6', '11') Then
          --出院、转院、死亡医嘱,将病人标记为预出院
          Begin
            Select 开始时间
            Into v_Date
            From 病人变动记录
            Where 开始时间 Is Not Null And 终止时间 Is Null And 病人id = r_Advice.病人id And 主页id = r_Advice.主页id;
          Exception
            When Others Then
              v_Date := To_Date('1900-01-01', 'YYYY-MM-DD');
          End;
          If r_Advice.开始执行时间 <= v_Date Then
            v_Error := '医嘱"' || r_Advice.医嘱内容 || '"的开始时间应大于该病人上次变动时间 ' || To_Char(v_Date, 'YYYY-MM-DD HH24:Mi') || ' 。';
            Raise Err_Custom;
          End If;
          Zl_病人变动记录_Preout(r_Advice.病人id, r_Advice.主页id, r_Advice.开始执行时间);
        End If;
      End If;
    End If;
    --12小时未执行的备用临嘱处理为标记未用
    If r_Advice.医嘱期效 = 1 Then
      Update 病人医嘱记录
      Set 执行标记 = -1
      Where 病人id = r_Advice.病人id And 主页id = r_Advice.主页id And 执行标记 <> -1 And 医嘱期效 = 1 And 执行频次 = '需要时' And
            Sysdate - 开始执行时间 > 0.5 And 医嘱状态 = 3;
    End If;
  End If;

  --填写发送记录
  ---------------------------------------------------------------------------------------
  n_执行状态 := 执行状态_In;
  If 执行状态_In = 1 Then
    v_Temp := zl_GetSysParameter(186);
    If v_Temp = '11' Then
      If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 in ('1','8') Or r_Advice.诊疗类别 = 'K' Then
        n_执行状态 := 0;
      End If;
    Elsif v_Temp = '01' Then
      If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '1' Then
        n_执行状态 := 0;
      End If;
    Elsif v_Temp = '10' Then
      If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '8' Or r_Advice.诊疗类别 = 'K' Then
        n_执行状态 := 0;
      End If;    
    End If;
  End If;

  Insert Into 病人医嘱发送
    (医嘱id, 发送号, 记录性质, NO, 记录序号, 发送数次, 发送人, 发送时间, 执行状态, 执行部门id, 计费状态, 首次时间, 末次时间, 样本条码, 门诊记帐)
  Values
    (医嘱id_In, 发送号_In, 记录性质_In, No_In, 记录序号_In, 发送数次_In, v_人员姓名, 发送时间_In, n_执行状态, 执行部门id_In, 计费状态_In,
     Nvl(首次时间_In, d_开始时间), Nvl(末次时间_In, d_开始时间), 样本条码_In, 门诊记帐_In);

  --手术和检查医嘱同步更新主医嘱的计费状态   
  If 计费状态_In = 1 And  r_Advice.组ID <> 医嘱id_In  And (r_Advice.诊疗类别 = 'D' Or r_Advice.诊疗类别 = 'F') Then   
     Update 病人医嘱发送 Set 计费状态 = 1 Where 医嘱ID = r_Advice.组ID And 发送号 = 发送号_In;
  End If;

  --领药号的填写
  If 领药号_In Is Not Null Then
    Update 未发药品记录 Set 领药号 = 领药号_In Where NO = No_In And 单据 = 9 And 领药号 Is Null;
    Update 药品收发记录 Set 产品合格证 = 领药号_In Where NO = No_In And 单据 = 9 And 产品合格证 Is Null;
  End If;

  --自动填为已执行时，需要同步处理费用执行状态及审核划价状态
  If 执行状态_In = 1 Then
    Zl_病人医嘱执行_Finish(医嘱id_In, 发送号_In, Null, Null, v_人员编号, v_人员姓名, 执行部门id_In);
  End If;

  --产生医嘱执行时间记录(只产生主记录的)
  If Nvl(分解时间_In, To_Char(d_开始时间, 'yyyy-mm-dd hh24:mi:ss')) Is Not Null Then
    If r_Advice.相关id Is Null Then
      Insert Into 医嘱执行时间
        (要求时间, 医嘱id, 发送号)
        Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), 医嘱id_In, 发送号_In
        From Table(f_Str2list(Nvl(分解时间_In, To_Char(d_开始时间, 'yyyy-mm-dd hh24:mi:ss'))));
    End If;
  End If;

  --病历书写时机的填写
  If r_Advice.诊疗类别 = 'F' Then
    --一组手术只调一次
    If r_Advice.相关id Is Null Then
      If Not r_Advice.标本部位 Is Null Then
        v_Date := To_Date(r_Advice.标本部位, 'yyyy-mm-dd hh24:mi:ss');
      Else
        v_Date := r_Advice.开始执行时间;
      End If;
      Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '手术', r_Advice.开嘱科室id, r_Advice.开嘱医生, v_Date, v_Date,
                       r_Advice.执行科室id);
    End If;
  Elsif r_Advice.诊疗类别 = 'Z' And r_Advice.操作类型 = '7' Then
    Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '会诊', r_Advice.开嘱科室id, r_Advice.开嘱医生, r_Advice.开始执行时间,
                     r_Advice.开始执行时间, r_Advice.执行科室id);
  Elsif r_Advice.诊疗类别 = 'Z' And r_Advice.操作类型 = '8' Then
    Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '抢救', r_Advice.开嘱科室id, r_Advice.开嘱医生, r_Advice.开始执行时间,
                     r_Advice.开始执行时间, r_Advice.执行科室id);
  Elsif r_Advice.诊疗类别 = 'Z' And r_Advice.操作类型 = '11' Then
    Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '死亡', r_Advice.开嘱科室id, r_Advice.开嘱医生, r_Advice.开始执行时间,
                     r_Advice.开始执行时间, r_Advice.执行科室id);
  End If;
  --额外调用(知情文件允许的诊疗类别才调用)
  If Instr('C,D,E,F,G,K,L', r_Advice.诊疗类别) > 0 Then
    Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '知情文书', r_Advice.开嘱科室id, r_Advice.开嘱医生, r_Advice.开始执行时间,
                     r_Advice.开始执行时间, r_Advice.执行科室id, r_Advice.诊疗项目id, r_Advice.医嘱内容);
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
        Select Nvl(Max(0), 2)
        Into n_标记
        From 业务消息清单 A
        Where a.病人id = r_Advice.病人id And a.就诊id = r_Advice.主页id And a.类型编码 = 'ZLHIS_CIS_002' And a.优先程度 = 2 And
              a.是否已阅 = 0;
      Else
        n_Adviceid := n_标记;
        Select Nvl(Max(0), 1)
        Into n_标记
        From 业务消息清单 A
        Where a.病人id = r_Advice.病人id And a.就诊id = r_Advice.主页id And a.类型编码 = 'ZLHIS_CIS_002' And a.是否已阅 = 0;
      End If;
      If n_标记 > 0 Then
        For R In (Select a.病人性质 As 性质, a.出院科室id As 科室id, a.当前病区id As 病区id
                  From 病案主页 A
                  Where a.病人id = r_Advice.病人id And a.主页id = r_Advice.主页id) Loop
          Zl_业务消息清单_Insert(r_Advice.病人id, r_Advice.主页id, r.科室id, r.病区id, r.性质, '有新停止医嘱。', '0010', 'ZLHIS_CIS_002',
                           n_Adviceid, n_标记, 0, Null, r.病区id);
        End Loop;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱发送_Insert;
/

--93623:梁唐彬,2016-03-04,检查计费状态调整
Create Or Replace Procedure Zl_门诊医嘱发送_Insert
(
  医嘱id_In     病人医嘱发送.医嘱id%Type,
  发送号_In     病人医嘱发送.发送号%Type,
  记录性质_In   病人医嘱发送.记录性质%Type,
  No_In         病人医嘱发送.No%Type,
  记录序号_In   病人医嘱发送.记录序号%Type,
  发送数次_In   病人医嘱发送.发送数次%Type,
  首次时间_In   病人医嘱发送.首次时间%Type,
  末次时间_In   病人医嘱发送.末次时间%Type,
  发送时间_In   病人医嘱发送.发送时间%Type,
  执行状态_In   病人医嘱发送.执行状态%Type,
  执行部门id_In 病人医嘱发送.执行部门id%Type,
  计费状态_In   病人医嘱发送.计费状态%Type,
  First_In      Number := 0,
  样本条码_In   病人医嘱发送.样本条码%Type := Null,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null
  --功能：填写病人医嘱发送记录 
  --参数：First_IN=表示是否一组医嘱的第一医嘱行,以便处理医嘱相关内容(如成药,配方的第一行,因为给药途径,配方煎法,用法可能为叮嘱不发送) 
) Is
  --包含病人及医嘱(一组医嘱中第一行)相关信息的游标 
  Cursor c_Advice Is
    Select Nvl(a.相关id, a.Id) As 组id, a.序号, a.病人id, a.挂号单, a.婴儿, b.姓名, c.操作类型, a.诊疗类别, a.医嘱状态, a.医嘱内容, a.开嘱医生, a.开始执行时间,
           a.执行时间方案, a.频率次数, a.频率间隔, a.间隔单位, Nvl(a.紧急标志, 0) As 紧急标志
    From 病人医嘱记录 A, 病人信息 B, 诊疗项目目录 C
    Where a.病人id = b.病人id And a.诊疗项目id = c.Id And a.Id = 医嘱id_In
    Group By Nvl(a.相关id, a.Id), a.序号, a.病人id, a.挂号单, a.婴儿, b.姓名, c.操作类型, a.诊疗类别, a.医嘱状态, a.医嘱内容, a.开嘱医生, a.开始执行时间,
             a.执行时间方案, a.频率次数, a.频率间隔, a.间隔单位, a.紧急标志;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_病人id 病人信息.病人id%Type) Is
    Select * From 病人信息 Where 病人id = v_病人id;
  r_Pati c_Pati%RowType;

  --其它临时变量 
  v_Temp     Varchar2(255);
  v_Count    Number;
  v_病人性质 病案主页.病人性质%Type;
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_入院方式 入院方式.名称%Type;
  d_开始时间 病人医嘱记录.开始执行时间%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
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
  --如果首次时间为空则填入开始执行时间
  If 首次时间_In Is Null Then
    Select 开始执行时间 Into d_开始时间 From 病人医嘱记录 Where ID = 医嘱id_In;
  End If;
  Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
  --是一组医嘱的第一行时处理医嘱内容 
  If Nvl(First_In, 0) = 1 Then
   
  
    --并发操作检查 
    --------------------------------------------------------------------------------------- 
    If Nvl(r_Advice.医嘱状态, 0) <> 1 Then
      v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被其他人发送。' || Chr(13) || Chr(10) ||
                 '该病人的医嘱发送失败。请重新读取发送清单再试。';
      Raise Err_Custom;
    End If;
  
    --发送后的医嘱处理:临嘱发送后自动停止 
    --------------------------------------------------------------------------------------- 
    Update 病人医嘱记录
    Set 医嘱状态 = 8, 执行终止时间 = 末次时间_In,
        --可能没有 
        停嘱时间 = 发送时间_In,
        --要作为发送时间显示 
        停嘱医生 = v_人员姓名 --要作为发送人显示,不同于住院,门诊医嘱无护士操作 
    Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
  
    Insert Into 病人医嘱状态
      (医嘱id, 操作类型, 操作人员, 操作时间)
      Select ID, 8, v_人员姓名, 发送时间_In From 病人医嘱记录 Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
  
    --特殊医嘱的处理 
    --------------------------------------------------------------------------------------- 
    If r_Advice.诊疗类别 = 'Z' And Nvl(r_Advice.操作类型, '0') <> '0' And Nvl(r_Advice.婴儿, 0) = 0 Then
      --1-留观;2-住院; 
      If Instr(',1,2,', r_Advice.操作类型) > 0 And 执行部门id_In Is Not Null Then
        --满足产生新的预约登记的条件：1.当前无预约,2.当前不在院,3-无要求预约时间内的住院记录 
      
        --删除超过挂号有效天数的预约登记 
        Begin
          Select Count(*) Into v_Count From 病案主页 Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = 0;
        Exception
          When Others Then
            v_Count := 0;
        End;
        If Nvl(v_Count, 0) > 0 Then
          Zl_入院病案主页_Delete(r_Advice.病人id, 0, 0, 0);
          v_Count := 0;
        End If;
      
        If v_Count = 0 Then
          Select Count(*) Into v_Count From 病案主页 Where 病人id = r_Advice.病人id And 出院日期 Is Null;
        End If;
        If v_Count = 0 Then
          Select Count(*)
          Into v_Count
          From 病案主页
          Where 病人id = r_Advice.病人id And (入院日期 >= r_Advice.开始执行时间 Or 出院日期 >= r_Advice.开始执行时间);
        End If;
        If v_Count = 0 Then
          If r_Advice.操作类型 = '1' Then
            --留观医嘱,将病人在"开始时间"留观到临床执行科室 
            Begin
              v_病人性质 := 2;
              Select Decode(服务对象, 1, 1, 2)
              Into v_病人性质
              From 部门性质说明
              Where 工作性质 = '临床' And 部门id = 执行部门id_In;
            Exception
              When Others Then
                Null;
            End;
          Elsif r_Advice.操作类型 = '2' Then
            --住院医嘱,将病人在"开始时间"登记到临床执行科室 
            v_病人性质 := 0;
          End If;
        
          Open c_Pati(r_Advice.病人id);
          Fetch c_Pati
            Into r_Pati;
        
          v_入院方式 := Null;
          If r_Advice.紧急标志 = 1 Then
            v_入院方式 := '急诊';
          Else
            Select Decode(急诊, 1, '急诊', Null)
            Into v_入院方式
            From 病人挂号记录
            Where NO = r_Advice.挂号单 And 记录性质 = 1 And 记录状态 = 1;
          End If;
        
          If v_病人性质 = 1 Then
            Zl_入院病案主页_Insert(1, v_病人性质, r_Pati.病人id, r_Pati.门诊号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别,
                             r_Pati.出生日期, r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份,
                             r_Pati.身份证号, r_Pati.出生地点, r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址,
                             r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系, r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位,
                             r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行, r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额,
                             r_Pati.担保性质, 执行部门id_In, Null, Null, v_入院方式, Null, Null, r_Advice.开嘱医生, r_Pati.籍贯, r_Pati.区域,
                             r_Advice.开始执行时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null, Null, Null, Null, r_Pati.险类,
                             v_人员编号, v_人员姓名, 0, Null, Null, 0);
          Else
            Zl_入院病案主页_Insert(1, v_病人性质, r_Pati.病人id, r_Pati.住院号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别,
                             r_Pati.出生日期, r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份,
                             r_Pati.身份证号, r_Pati.出生地点, r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址,
                             r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系, r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位,
                             r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行, r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额,
                             r_Pati.担保性质, 执行部门id_In, Null, Null, v_入院方式, Null, Null, r_Advice.开嘱医生, r_Pati.籍贯, r_Pati.区域,
                             r_Advice.开始执行时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null, Null, Null, Null, r_Pati.险类,
                             v_人员编号, v_人员姓名, 0, Null, Null, 0);
          End If;
          Close c_Pati;
        End If;
      End If;
    End If;
  
   
  End If;
  Close c_Advice;
  --填写发送记录 
  --------------------------------------------------------------------------------------- 
  Insert Into 病人医嘱发送
    (医嘱id, 发送号, 记录性质, NO, 记录序号, 发送数次, 发送人, 发送时间, 执行状态, 执行部门id, 计费状态, 首次时间, 末次时间, 样本条码, 门诊记帐)
  Values
    (医嘱id_In, 发送号_In, 记录性质_In, No_In, 记录序号_In, 发送数次_In, v_人员姓名, 发送时间_In, 执行状态_In, 执行部门id_In, 计费状态_In,
     Nvl(首次时间_In, d_开始时间), Nvl(末次时间_In, d_开始时间), 样本条码_In, Decode(记录性质_In, 2, 1, Null));

  --手术和检查医嘱同步更新主医嘱的计费状态   
  If 计费状态_In = 1 And  r_Advice.组ID <> 医嘱id_In  And (r_Advice.诊疗类别 = 'D' Or r_Advice.诊疗类别 = 'F') Then   
     Update 病人医嘱发送 Set 计费状态 = 1 Where 医嘱ID = r_Advice.组ID And 发送号 = 发送号_In;
  End If;
  
  --自动填为已执行时，需要同步处理费用执行状态及审核划价状态 
  If 执行状态_In = 1 Then
    Zl_病人医嘱执行_Finish(医嘱id_In, 发送号_In, Null, Null, v_人员编号, v_人员姓名, 执行部门id_In);
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin zl_服务窗消息_发送(:1,:2); End;'
      Using 3, 发送号_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊医嘱发送_Insert;
/

--93868:刘尔旋,2016-03-04,清除失效预约问题
Create Or Replace Procedure Zl_挂号序号状态_Delete
(
  操作方式_In Number := 0,
  号别_In     病人挂号记录.号别%Type := Null
) As
  n_预约有效时间 Number(5);
  n_失约用于挂号 Number(2);
  n_挂号有效天数 Number(5);
Begin
  If 操作方式_In = 0 Then
    --清除历史记录
    Delete 挂号序号状态 Where 日期 < Trunc(Sysdate);
  Else
    --清除失约号
    n_预约有效时间 := Nvl(zl_GetSysParameter('预约有效时间', 1111), 0);
    n_失约用于挂号 := Nvl(zl_GetSysParameter('失约用于挂号', 1111), 0);
    n_挂号有效天数 := Nvl(zl_GetSysParameter('挂号有效天数'), 7);
    If n_预约有效时间 <> 0 And n_失约用于挂号 <> 0 Then
      If 号别_In Is Null Then
        For c_失效预约 In (Select b.号码, b.日期, b.序号
                       From 病人挂号记录 A, 挂号序号状态 B
                       Where a.预约时间 - 1 / 24 / 60 * n_预约有效时间 < Sysdate And a.预约时间 > Sysdate - n_挂号有效天数 And a.记录性质 = 2 And
                             a.号别 = b.号码 And a.号序 = b.序号) Loop
          Delete From 挂号序号状态
          Where 日期 = c_失效预约.日期 And 序号 = c_失效预约.序号 And 状态 = 2 And 号码 = c_失效预约.号码;
        End Loop;
      Else
        For c_失效预约 In (Select b.号码, b.日期, b.序号
                       From 病人挂号记录 A, 挂号序号状态 B
                       Where a.预约时间 - 1 / 24 / 60 * n_预约有效时间 < Sysdate And a.预约时间 > Sysdate - n_挂号有效天数 And a.记录性质 = 2 And
                             a.号别 = b.号码 And a.号序 = b.序号 And a.号别 = 号别_In) Loop
          Delete From 挂号序号状态
          Where 日期 = c_失效预约.日期 And 序号 = c_失效预约.序号 And 状态 = 2 And 号码 = c_失效预约.号码;
        End Loop;
      End If;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号序号状态_Delete;
/

--93820:马政,2016-03-03,主动调售价后原价字段处理
Create Or Replace Procedure Zl_药品收发记录_Adjust
(
  Adjustid    In Number, --调价记录的ID
  Bln定价     In Number := 0, --是否转为定价销售（更新2004-06-08、收费细目中的变价）
  Billinfo_In In Varchar2 := Null, --用于时价药品按批次调价。格式:"批次1,现价1|批次2,现价2|....."
  药品id_In   In Number := 0 --当不为0时表示是成本价调价，不处理售价相关内容
) As
  Classid      Number(18); --入出类别
  v_Billno     药品收发记录.No%Type; --调价单号
  Rundate      Date; --调价生效时间
  Blnrun       Number(1); --调价时刻到了
  Blncurprice  Number(1); --时价药品
  Lng细目id    Number(18); --收费细目ID
  Adjustdate   Date; --调价时间
  n_批次       Number(18);
  n_现价       收费价目.现价%Type;
  n_原价       收费价目.原价%Type;
  v_Infotmp    Varchar2(4000);
  v_Fields     Varchar2(4000);
  n_序号       Number(8);
  n_原价id     收费价目.原价id%Type;
  n_收入项目id 收费价目.收入项目id%Type;
  n_零售金额   药品库存.实际金额%Type;
  n_零售价     药品库存.零售价%Type;
  n_收发id     药品收发记录.Id%Type;
  n_变动原因   收费价目.变动原因%Type;

  Cursor c_Price --普通调价
  Is
    Select 1 记录状态, 13 单据, v_Billno NO, Rownum 序号, Classid 入出类别id, m.Id As 药品id, s.批次, Null 批号, Null 效期, s.上次产地 As 产地,
           1 付数, s.上次供应商id As 供应商id, s.实际数量 填写数量, 0 实际数量, a.原价 成本价, 0 成本金额, a.现价 零售价, 0 扣率, '药品调价' 摘要, Zl_Username 填制人,
           Sysdate 填制日期, s.库房id 库房id, 1 入出系数, a.Id 价格id, Nvl(m.是否变价, 0) As 时价, s.实际金额 As 库存金额, s.实际差价 As 库存差价,
           Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, a.原价, Nvl(s.实际金额, 0) / s.实际数量)) As 原售价
    From 药品库存 S, 收费项目目录 M, 收费价目 A
    Where s.药品id = m.Id And m.Id = a.收费细目id And s.性质 = 1 And a.变动原因 = 0 And a.Id = Adjustid And a.执行日期 <= Sysdate;

  v_Data c_Price%RowType;

  Cursor c_时价按批次调价 --时价药品按批次调价
  Is
    Select 1 记录状态, 13 单据, v_Billno NO, n_序号 + Rownum 序号, Classid 入出类别id, m.药品id 药品id, s.批次 批次, Null 批号, Null 效期,
           s.上次产地 As 产地, s.上次供应商id As 供应商id, 1 付数, Nvl(s.实际数量, 0) 填写数量, 0 实际数量, a.原价 成本价, 0 成本金额, n_现价 零售价, 0 扣率,
           '药品调价' 摘要, Zl_Username 填制人, Sysdate 填制日期, s.库房id 库房id, 1 入出系数, a.Id 价格id, Nvl(b.是否变价, 0) As 时价,
           s.实际金额 As 库存金额, s.实际差价 As 库存差价, Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, a.原价, Nvl(s.实际金额, 0) / s.实际数量)) As 原售价
    From 药品库存 S, 药品目录 M, 收费价目 A, 收费项目目录 B
    Where s.药品id = b.Id And s.药品id = m.药品id And m.药品id = a.收费细目id And s.性质 = 1 And a.变动原因 = 0 And a.Id = Adjustid And
          a.执行日期 <= Sysdate And Nvl(s.批次, 0) = n_批次;

  v_时价按批次调价 c_时价按批次调价%RowType;
Begin
  If 药品id_In <> 0 Then
    --成本价调价
    Zl_药品收发记录_成本价调价(药品id_In);
  Else
    n_变动原因 := 0;
    --取调价记录生效日期
    Select 收费细目id, 执行日期, 收入项目id Into Lng细目id, Rundate, n_收入项目id From 收费价目 Where ID = Adjustid;
  
    If Sysdate >= Rundate Then
      Blnrun := 1;
    Else
      Blnrun := 0;
    End If;
  
    If Blnrun = 1 Then
      --取入出类别ID
      Select 类别id Into Classid From 药品单据性质 Where 单据 = 13;
    
      --取序列
      Select Nextno(147) Into v_Billno From Dual;
    
      --取该药品是否是时价药品
      Select Nvl(是否变价, 0) Into Blncurprice From 收费项目目录 Where ID = Lng细目id;
    
      --检查是否存在原价和现价相同的情况，相同时不执行调售价功能，并且删除这条收费价目记录，恢复原来的收费价目
      Begin
        Select 原价id Into n_原价id From 收费价目 Where ID = Adjustid And 原价 = 现价 And 原价id Is Not Null;
      Exception
        When Others Then
          n_原价id := 0;
      End;
    
      If n_原价id > 0 Then
        --如果现价=原价，这种情况下是单独调整收入项目，更新收入项目ID，删除调价记录
        Delete 收费价目 Where ID = Adjustid;
        Update 收费价目
        Set 收入项目id = n_收入项目id, 终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')
        Where ID = n_原价id;
      Else
        Adjustdate := Sysdate;
      
        Begin
          Select 变动原因 Into n_变动原因 From 收费价目 Where ID = Adjustid And 变动原因 = 1;
        Exception
          When Others Then
            n_变动原因 := 0;
        End;
        If n_变动原因 = 0 Then
          If Billinfo_In = '' Or Billinfo_In Is Null Then
            For v_Data In c_Price Loop
              Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
              If Nvl(v_Data.填写数量, 0) = 0 And (Nvl(v_Data.库存金额, 0) <> 0 Or Nvl(v_Data.库存差价, 0) <> 0) Then
                --数量=0 金额或差价<>0时只更新库存表中对应的零售价,并产生售价修正数据但是金额差=0，只记录最新售价，金额差和差价差不填数据            
                --产生调价影响记录
                Insert Into 药品收发记录
                  (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 摘要, 填制人,
                   填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
                Values
                  (n_收发id, v_Data.记录状态, v_Data.单据, v_Data.No, v_Data.序号, v_Data.入出类别id, v_Data.药品id, v_Data.批次,
                   v_Data.批号, v_Data.效期, v_Data.产地, v_Data.付数, v_Data.填写数量, v_Data.实际数量,
                   Decode(Blncurprice, 1, v_Data.原售价, v_Data.成本价), v_Data.成本金额, v_Data.零售价, v_Data.扣率, v_Data.摘要,
                   v_Data.填制人, v_Data.填制日期, v_Data.库房id, v_Data.入出系数, v_Data.价格id, Zl_Username, Adjustdate, v_Data.库存金额,
                   v_Data.库存差价, v_Data.供应商id);
              
                --更新库存零售价,只有时价分批药品才能更新零售价字段
                Zl_药品库存_Update(n_收发id);
              Else
                If Blncurprice = 1 Then
                  n_零售价 := v_Data.库存金额 / v_Data.填写数量;
                Else
                  n_零售价 := v_Data.成本价;
                End If;
                n_零售金额 := Round((v_Data.零售价 - n_零售价) * v_Data.填写数量, 2);
              
                --产生调价影响记录
                Insert Into 药品收发记录
                  (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要,
                   填制人, 填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
                Values
                  (n_收发id, v_Data.记录状态, v_Data.单据, v_Data.No, v_Data.序号, v_Data.入出类别id, v_Data.药品id, v_Data.批次,
                   v_Data.批号, v_Data.效期, v_Data.产地, v_Data.付数, v_Data.填写数量, v_Data.实际数量,
                   Decode(Blncurprice, 1, v_Data.原售价, v_Data.成本价), v_Data.成本金额, v_Data.零售价, v_Data.扣率, n_零售金额, n_零售金额,
                   v_Data.摘要, v_Data.填制人, v_Data.填制日期, v_Data.库房id, v_Data.入出系数, v_Data.价格id, Zl_Username, Adjustdate,
                   v_Data.库存金额, v_Data.库存差价, v_Data.供应商id);
              
                --更新药品库存
                Zl_药品库存_Update(n_收发id);
              End If;
            End Loop;
          Else
            n_序号 := 0;
            --时价药品按批次调价
            v_Infotmp := Billinfo_In || '|';
            While v_Infotmp Is Not Null Loop
              --分解单据ID串
              v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
              n_批次    := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
              n_现价    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
              v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
            
              For v_时价按批次调价 In c_时价按批次调价 Loop
                If v_时价按批次调价.填写数量 <> 0 Then
                  n_原价 := Nvl(v_时价按批次调价.库存金额, 0) / v_时价按批次调价.填写数量;
                Else
                  n_原价 := v_时价按批次调价.成本价;
                End If;
              
                Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
                If Nvl(v_时价按批次调价.填写数量, 0) = 0 And (Nvl(v_时价按批次调价.库存金额, 0) <> 0 Or Nvl(v_时价按批次调价.库存差价, 0) <> 0) Then
                  --数量=0 金额或差价<>0时只更新库存表中对应的零售价,并产生售价修正数据但是金额差=0，只记录最新售价，金额差和差价差不填数据              
                  --产生调价影响记录
                  Insert Into 药品收发记录
                    (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 摘要, 填制人,
                     填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
                  Values
                    (n_收发id, v_时价按批次调价.记录状态, v_时价按批次调价.单据, v_时价按批次调价.No, v_时价按批次调价.序号, v_时价按批次调价.入出类别id, v_时价按批次调价.药品id,
                     v_时价按批次调价.批次, v_时价按批次调价.批号, v_时价按批次调价.效期, v_时价按批次调价.产地, v_时价按批次调价.付数, v_时价按批次调价.填写数量,
                     v_时价按批次调价.实际数量, Decode(Blncurprice, 1, v_时价按批次调价.原售价, v_时价按批次调价.成本价), v_时价按批次调价.成本金额, v_时价按批次调价.零售价,
                     v_时价按批次调价.扣率, v_时价按批次调价.摘要, v_时价按批次调价.填制人, v_时价按批次调价.填制日期, v_时价按批次调价.库房id, v_时价按批次调价.入出系数,
                     v_时价按批次调价.价格id, Zl_Username, Adjustdate, v_时价按批次调价.库存金额, v_时价按批次调价.库存差价, v_时价按批次调价.供应商id);
                  n_序号 := n_序号 + 1;
                
                  --更新库存零售价,只有时价分批药品才能更新零售价字段
                  Zl_药品库存_Update(n_收发id);
                Else
                
                  n_零售价   := v_时价按批次调价.库存金额 / v_时价按批次调价.填写数量;
                  n_零售金额 := Round((n_现价 - n_零售价) * v_时价按批次调价.填写数量, 2);
                  --产生调价影响记录
                  Insert Into 药品收发记录
                    (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价,
                     摘要, 填制人, 填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
                  Values
                    (n_收发id, v_时价按批次调价.记录状态, v_时价按批次调价.单据, v_时价按批次调价.No, v_时价按批次调价.序号, v_时价按批次调价.入出类别id, v_时价按批次调价.药品id,
                     v_时价按批次调价.批次, v_时价按批次调价.批号, v_时价按批次调价.效期, v_时价按批次调价.产地, v_时价按批次调价.付数, v_时价按批次调价.填写数量,
                     v_时价按批次调价.实际数量, Decode(Blncurprice, 1, v_时价按批次调价.原售价, v_时价按批次调价.成本价), v_时价按批次调价.成本金额, v_时价按批次调价.零售价,
                     v_时价按批次调价.扣率, n_零售金额, n_零售金额, v_时价按批次调价.摘要, v_时价按批次调价.填制人, v_时价按批次调价.填制日期, v_时价按批次调价.库房id,
                     v_时价按批次调价.入出系数, v_时价按批次调价.价格id, Zl_Username, Adjustdate, v_时价按批次调价.库存金额, v_时价按批次调价.库存差价,
                     v_时价按批次调价.供应商id);
                  n_序号 := n_序号 + 1;
                
                  --更新库存
                  Zl_药品库存_Update(n_收发id);
                End If;
              End Loop;
            End Loop;
          End If;
        
          Update 收费价目 Set 变动原因 = 1 Where ID = Adjustid;
        
          --更新药品目录、收费细目中的变价
          If Bln定价 = 1 Then
            Update 收费项目目录 Set 是否变价 = 0 Where ID = Lng细目id;
            Update 收费细目 Set 是否变价 = 0 Where ID = Lng细目id;
          End If;
        End If;
      End If;
    
      If n_变动原因 = 0 Then
        --成本价调价
        Zl_药品收发记录_成本价调价(Lng细目id, Rundate);
      End If;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_Adjust;
/
--93853:张永康,2016-03-03,历史数据抽回错误修正
Create Or Replace Procedure Zl_Retu_Exes
(
  v_No   In Varchar2,
  n_Type In Number
) As
  --------------------------------------------
  --参数:v_No,单据号码
  --     n_Type,单据类型:1-收费,2-记帐,3-自动记帐,4-挂号,5-就诊卡,6-预交,7-结帐,8-未结费用的病人id,主页ID
  --------------------------------------------
  n_Allow  Number(1); --是否能够单据返回
  n_Patiid Number(18);
  n_Pageid Number(5);
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
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
                  ' Where 预交id = :1';
      Execute Immediate v_Sql
        Using r_Rec.Id;
    
      v_Table  := '三方结算交易';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
                  ' Where 交易ID = :1';
      Execute Immediate v_Sql
        Using r_Rec.Id;
    
      v_Table  := '三方退款信息';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
                  ' Where 记录ID = :1 And 结帐ID = :2';
      Execute Immediate v_Sql
        Using r_Rec.Id, n_Settle_Id;
    
      v_Table  := '病人卡结算记录';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
                  ' Where ID In(Select 卡结算id From H病人卡结算对照 Where 预交id = :1)';
      Execute Immediate v_Sql
        Using r_Rec.Id;
    
      Delete H病人卡结算记录 Where ID In (Select Distinct 卡结算id From H病人卡结算对照 Where 预交id = r_Rec.Id);
      Delete From H病人卡结算对照 Where 预交id = r_Rec.Id;
      Delete From H三方结算交易 Where 交易id = r_Rec.Id;
      Delete From H三方退款信息 Where 记录id = r_Rec.Id And 结帐id = n_Settle_Id;
    End Loop;
  
    v_Table  := '病人预交记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
                ' Where 结帐id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    v_Table  := '门诊费用记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
                ' Where 结帐id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    v_Table  := '费用补充记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
                ' Where 结算id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    v_Table  := '医保结算明细';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
                ' Where ID = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    v_Table  := '输液配药记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
                ' Where ID In(Select 记录ID From H输液配药内容 Where 收发ID =:1)';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    For P In (Select ID From H输液配药记录 Where ID In (Select 记录id From H输液配药内容 Where 收发id = n_Rec_Id)) Loop
      For R In (Select Column_Value From Table(f_Str2list('输液配药附费,输液配药状态'))) Loop
        v_Table := r.Column_Value;
        v_Field := 'ID';
      
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
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
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
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

  If n_Type = 8 Then
    --8-抽回指定病人的未结记帐费用
    --排除：结帐作废后，记帐单销帐的记录（记录状态为2的没有结帐ID，记录状态为3的有结帐ID的) 
    If Instr(v_No, ',') = 0 Then
      --a.按病人ID抽回门诊病人的未结记帐费用
      For Rno In (Select NO, 记录性质
                  From H门诊费用记录 A
                  Where 病人id = To_Number(v_No) And a.记帐费用 = 1 And 结帐id Is Null And Not Exists
                   (Select 1
                         From H门诊费用记录 B
                         Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null)) Loop
        v_Table  := '门诊费用记录';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where NO = :1 And 记录性质 = :2';
        Execute Immediate v_Sql
          Using Rno.No, Rno.记录性质;
      
        For r_Rxlist In (Select m.Id
                         From H药品收发记录 M, H门诊费用记录 E
                         Where e.No = Rno.No And e.记录性质 = Rno.记录性质 And e.收费类别 In ('4', '5', '6', '7') And m.费用id = e.Id And
                               m.单据 In (9, 10, 25, 26)) Loop
          Zl_Retu_Medilist(r_Rxlist.Id);
        End Loop;
        Delete H门诊费用记录 Where NO = Rno.No And 记录性质 = Rno.记录性质;
      End Loop;
    Else
      --b.按病人ID,主页ID抽回住院病人的未结记帐费用
      n_Patiid := Substr(v_No, 1, Instr(v_No, ',') - 1);
      n_Pageid := Substr(v_No, Instr(v_No, ',') + 1);
    
      For Rno In (Select NO, 记录性质
                  From H住院费用记录 A
                  Where 病人id = n_Patiid And 主页id = n_Pageid And a.记帐费用 = 1 And 结帐id Is Null And Not Exists
                   (Select 1
                         From H住院费用记录 B
                         Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null)) Loop
        v_Table  := '住院费用记录';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where NO = :1 And 记录性质 = :2';
        Execute Immediate v_Sql
          Using Rno.No, Rno.记录性质;
      
        For r_Rxlist In (Select m.Id
                         From H药品收发记录 M, H住院费用记录 E
                         Where e.No = Rno.No And e.记录性质 = Rno.记录性质 And e.收费类别 In ('4', '5', '6', '7') And m.费用id = e.Id And
                               m.单据 In (9, 10, 25, 26)) Loop
          Zl_Retu_Medilist(r_Rxlist.Id);
        End Loop;
        Delete H住院费用记录 Where NO = Rno.No And 记录性质 = Rno.记录性质;
      End Loop;
    End If;
  Else
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
           Select ID
           From H病人结帐记录
           Where NO = v_No And 7 = n_Type) L
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
                    Select ID
                    From H病人结帐记录
                    Where NO = v_No And 7 = n_Type) L
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
                    Select ID
                    From H病人结帐记录
                    Where NO = v_No And 7 = n_Type) L
             Where e.结帐id = l.结帐id) E;
    End If;
  
    --按照单据或病人获取结帐游标返回
    If n_Allow = 1 Then
      For r_Settle In (Select 结帐id
                       From H门诊费用记录
                       Where NO = v_No And
                             (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
                             4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5)
                       Union
                       Select 结帐id
                       From H住院费用记录
                       Where NO = v_No And
                             (1 = n_Type And 记录性质 = 1 Or 2 = n_Type And 记录性质 = 2 Or 3 = n_Type And 记录性质 = 3 Or
                             4 = n_Type And 记录性质 = 4 Or 5 = n_Type And 记录性质 = 5)
                       Union
                       Select 结帐id
                       From H病人预交记录
                       Where NO = v_No And 6 = n_Type And 记录性质 In (1, 11)
                       Union All
                       Select ID
                       From H病人结帐记录
                       Where NO = v_No And 7 = n_Type) Loop
      
        v_Table  := '病人结帐记录';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
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
          Select Distinct 病人id
          Into n_Patiid
          From H病人预交记录
          Where NO = v_No And 6 = n_Type And 记录性质 In (1, 11);
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
                   Select Distinct 病人id
                   From H费用补充记录
                   Where NO = v_No And 记录性质 = 1)
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
                       Select Distinct 结算id
                       From H费用补充记录
                       Where 病人id = n_Patiid) Loop
      
        v_Table  := '病人结帐记录';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'待转出','Null as 待转出') || ' From H' || v_Table ||
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
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM || ':' || v_Sql);
End Zl_Retu_Exes;
/

--93846:刘尔旋,2016-03-03,诊间支付判断
Create Or Replace Function Zl_Get_Threecardtypeid
(
  模块号_In Number,
  病人id_In 门诊费用记录.病人id%Type,
  医嘱id_In 病人医嘱记录.Id%Type
) Return Number Is
  ------------------------------------------------------------------------------------------- 
  --功能:获取当前三方接口支付的卡类别ID 
  --入参: 模块号_IN-调整的模块号 
  --      病人ID_In-当前就诊的病人ID 
  --      医嘱id_In-医嘱ID
  --返回:三方帐户支付的卡类别ID,如果返回了卡类别ID,则不再进入刷卡界面,而直接调用三方支付接口 
  --说明: 
  --  1.门诊医生站调用: 目前暂时在门诊医生工作站调用,需要配合诊间支付的参数设置才有效 
  --  2.其他执行科室:暂未调用该过程,如期需要可以扩展 
  --  3.此过程,需根据用户的具体业务需求进行返回 
  ------------------------------------------------------------------------------------------- 
Begin
  If Nvl(模块号_In, 0) = 1260 Then
    --门诊医生工作站 
    If Nvl(病人id_In, 0) = 0 And Nvl(医嘱id_In, 0) = 0 Then
      Return Null;
    End If;
    Return Null;
  End If;
  Return Null;
End Zl_Get_Threecardtypeid;
/

--93853:张永康,2016-03-03,历史数据抽回错误修正
--93801:吴涛,2016-03-02,历史数据抽回的错误
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
    v_Field Varchar(100);
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
              From Table(f_Str2list('电子病历附件,电子病历格式,电子病历内容,疾病申报记录,疾病报告反馈'))) Loop
      v_Table := r.Column_Value;
      If v_Table = '电子病历附件' Then
        v_Field := '病历id';
      Else
        v_Field := '文件id';
      End If;
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      If v_Table = '电子病历内容' Then
        v_Fields := Getfields('电子病历图形');
        v_Sql    := 'Insert Into 电子病历图形(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where id = :1';
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
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      If v_Table = '病人医嘱状态' Then
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
          v_Sql    := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' ||
                      Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Subtable || ' Where ' || v_Subfield ||
                      ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        
          v_Sql := 'Delete H' || v_Subtable || ' Where ' || v_Subfield ||
                   ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
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
      End If;
    
      v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
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

--91646:张德婷,2016-02-29,静配收取材料费
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
      for r_item in (Select A.NO,B.序号 From 输液配药附费 A,住院费用记录 B Where A.病人ID=B.病人ID and A.NO=B.No and B.记录状态=1 and A.配药id = v_Tansid) loop
        if r_item.NO is not null then
          Zl_住院记帐记录_Delete(r_item.NO,r_item.序号, v_Usercode, Zl_Username);
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

--93696:马政,2016-03-03,无库存冲销数据修正处理
Create Or Replace Procedure Zl_药品收发记录_调价修正(收发id_In In 药品收发记录.Id%Type) Is
  v_单据         药品收发记录.单据%Type;
  v_入出类别id   药品收发记录.入出类别id%Type;
  v_售价精度     Number;
  v_金额精度     Number;
  v_收发id       药品收发记录.Id%Type;
  v_原价         药品收发记录.零售价%Type;
  v_现价         药品收发记录.零售价%Type;
  v_是否变价     收费项目目录.是否变价%Type;
  v_价格id       收费价目.Id%Type;
  v_库房id       药品收发记录.库房id%Type;
  v_药品id       药品收发记录.药品id%Type;
  v_批次         药品收发记录.批次%Type;
  v_实际数量     药品收发记录.实际数量%Type;
  v_修正金额     药品收发记录.零售金额%Type;
  v_入出系数     药品收发记录.入出系数%Type;
  v_执行修正     Number;
  v_审核日期     药品收发记录.审核日期%Type;
  v_零售金额     药品收发记录.零售金额%Type;
  v_差价         药品收发记录.差价%Type;
  v_No           药品收发记录.No%Type;
  n_平均成本价   药品库存.平均成本价%Type;
  n_费用单价精度 Number;
  d_填制日期     药品收发记录.填制日期%Type;
  d_日期         药品收发记录.填制日期%Type;
  n_记录状态     药品收发记录.记录状态%Type;
  n_当前数量     药品库存.实际数量%Type;

  v_Billno 药品收发记录.No%Type;
Begin
  --修正规则
  --售价：
  --定价：以收费价目现价与冲销价格判断，如果不等产生修正
  --时价：时价无库存不修正有库存则判断当前价格与冲销价格是否不同，不同则处理否则不处理

  --成本价：
  --有调价则修正，无调价检查库存与冲销价格是否不同，不同则修正，否则不修正

  --1、售价调剂处理
  --提取单据信息，原价，现价及药价属性
  Select a.单据, a.记录状态, a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) As 实际数量, a.入出系数,
         Nvl(a.零售价, 0) 原价, b.现价, Nvl(c.是否变价, 0) 是否变价, b.Id As 价格id
  Into v_单据, n_记录状态, v_库房id, v_药品id, v_批次, v_实际数量, v_入出系数, v_原价, v_现价, v_是否变价, v_价格id
  From 药品收发记录 A, 收费价目 B, 收费项目目录 C
  Where a.药品id = b.收费细目id And a.药品id = c.Id And
        (Sysdate Between b.执行日期 And b.终止日期 Or Sysdate >= b.执行日期 And b.终止日期 Is Null) And a.Id = 收发id_In;

  --获取金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(157), '5')) Into n_费用单价精度 From Dual;
  If v_单据 = 8 Or v_单据 = 9 Or v_单据 = 10 Then
    --发药应该取费用精度，其他业务应该取药品卫材精度中精度
    Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_金额精度 From Dual;
  Else
    Select Nvl(精度, 2) Into v_金额精度 From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;
  End If;
  --定价直接取收费价目中现价
  v_执行修正 := 0;
  If v_是否变价 = 0 Then
    v_执行修正 := 1;
  Else
    --时价药品的现售价，不能从收费价目中取，因为时价调价可能是按库房和批次来调整的
    v_现价 := 0;
  
    If v_批次 > 0 Then
      --时价分批优先从库存表中取，如果没有库存则说明冲销价格与当前价格一样则不用处理了
      Begin
        Select Nvl(零售价, 0)
        Into v_现价
        From 药品库存
        Where 性质 = 1 And 库房id + 0 = v_库房id And 药品id = v_药品id And Nvl(批次, 0) = v_批次;
      Exception
        When Others Then
          v_现价 := 0;
      End;
    End If;
  
    If v_现价 > 0 Then
      v_执行修正 := 1;
    End If;
  End If;

  If v_执行修正 = 1 Then
    --发药类单据售价精度为5，其他流通类单据为7
    If v_单据 = 8 Or v_单据 = 9 Or v_单据 = 10 Then
      v_售价精度 := n_费用单价精度;
    Else
      v_售价精度 := 7;
    End If;
  
    --比较原价和现价，不同则处理
    If v_原价 <> Round(v_现价, v_售价精度) Then
      Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
      Select Nextno(147) Into v_Billno From Dual;
    
      Select 类别id Into v_入出类别id From 药品单据性质 Where 单据 = 13;
    
      v_修正金额 := Round(v_入出系数 * (Round(v_现价, v_售价精度) - v_原价) * v_实际数量, v_金额精度);
    
      --产生调价修正记录
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
         填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 费用id)
        Select v_收发id, 1, 13, v_Billno, 序号, v_入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, Abs(v_实际数量), 0, v_原价, 0,
               Round(v_现价, v_售价精度), 扣率, v_修正金额, v_修正金额, '自动修正调价变动', 审核人, 审核日期, 库房id, 1, v_价格id, 审核人, 审核日期, 收发id_In
        From 药品收发记录
        Where ID = 收发id_In;
    
      --更新药品库存
      Zl_药品库存_Update(v_收发id);
    
    End If;
  End If;

  --2、成本价调价，只有冲销业务才修正
  If Mod(n_记录状态, 3) = 2 Then
    Select 库房id, 药品id, Nvl(批次, 0) 批次, 入出系数 * Nvl(实际数量, 0) * Nvl(付数, 1) As 实际数量, 入出系数 * 零售金额, 入出系数 * 差价, 成本价
    Into v_库房id, v_药品id, v_批次, v_实际数量, v_零售金额, v_差价, v_原价
    From 药品收发记录
    Where ID = 收发id_In;
  
    --取原始单据的审核时间
    Select a.审核日期
    Into v_审核日期
    From 药品收发记录 A, 药品收发记录 B
    Where b.Id = 收发id_In And a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And a.药品id + 0 = b.药品id And a.序号 = b.序号 And
          (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0);
  
    v_执行修正 := 0;
    Begin
      Select 1, 新成本价
      Into v_执行修正, v_现价
      From 成本价调价信息
      Where 库房id + 0 = v_库房id And 药品id + 0 = v_药品id And Nvl(批次, 0) = v_批次 And 执行日期 > v_审核日期 And Rownum = 1
      Order By 执行日期 Desc;
    
    Exception
      When Others Then
        v_执行修正 := 0;
        v_现价     := 0;
      
        --可能是无库存调价，那么不需要库房id和批次判断，只需要用药品id和批次
        Begin
          Select 1, 新成本价
          Into v_执行修正, v_现价
          From 成本价调价信息
          Where 药品id + 0 = v_药品id And 执行日期 > v_审核日期 And Rownum = 1
          Order By 执行日期 Desc;
        
        Exception
          When Others Then
            --可能出现冲销或者退库很久以前的数据,这个时候价格可能已经发生变化,所以需要修正
            Begin
              Select 平均成本价, Nvl(实际数量, 0) As 实际数量
              Into v_现价, n_当前数量
              From 药品库存
              Where 库房id = v_库房id And 药品id = v_药品id And Nvl(批次, 0) = v_批次 And 性质 = 1;
            
              --无库存不处理，有库存可能有两种情况，1.原始有库存，冲销抵消后还有数量，2、原始无数量，新库存数据来源于冲销单据
              --如果是1情况那么则通过库存表原价格与冲销价格比较即可
              --如果是2情况那么则需要判断当前库存数量与冲销数量*入出系数是否相等，如果相等则说明肯定是原始无库存然后冲销根据产生的
              --这个时候需要粗略判断，只能通过规格来，另加一下价格判断，如价格为负数
              If n_当前数量 = v_实际数量 Then
                --说明是第二种情况产生的库存数据，这个时候可能产生的数据很离谱，则需要简单判断
                --93696问题需要将此语句屏蔽掉，及无库存退库时不再处理离谱数据
                /*Select 成本价 Into v_现价 From 药品规格 Where 药品id = v_药品id;
                If Abs(Abs(v_现价) - Abs(v_原价)) > 1 Then
                  v_执行修正 := 1;
                End If;*/
                Null;
              Else
                --说明是第一种情况产生的数据，直接用当前库存与冲销单据价格比较即可
                If Round(v_现价, 2) <> Round(v_原价, 2) Then
                  v_执行修正 := 1;
                End If;
              End If;
            Exception
              When Others Then
                v_执行修正 := 0;
            End;
        End;
    End;
  
    If v_执行修正 = 1 Then
      v_修正金额 := (v_零售金额 - v_差价) - Round(Round(v_现价, 3) * v_实际数量, v_金额精度);
    
      If v_修正金额 <> 0 Then
        Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
        Select b.Id, b.系数
        Into v_入出类别id, v_入出系数
        From 药品单据性质 A, 药品入出类别 B
        Where a.类别id = b.Id And a.单据 = 5 And Rownum < 2;
      
        v_No := Nextno(25, v_库房id);
      
        --产生库存差价调整单
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期,
           审核人, 审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率, 灭菌效期, 费用id)
          Select v_收发id, 1, 5, v_No, 1, 库房id, v_入出类别id, 供药单位id, v_入出系数, 药品id, 批次, 产地, 批号, 效期, v_实际数量, v_零售金额, v_差价,
                 v_修正金额, '自动修正调价变动', 审核人, 审核日期, 审核人, 审核日期, 生产日期, 批准文号, v_现价, 1, 成本价, 灭菌效期, 收发id_In
          From 药品收发记录
          Where ID = 收发id_In;
      
        --更新库存
        Zl_药品库存_Update(v_收发id);
      
      End If;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_调价修正;
/

--93696:马政,2016-03-03,无库存冲销数据修正处理
Create Or Replace Procedure Zl_药品库存_Update
(
  Id_In       In 药品收发记录.Id%Type,
  Delete_In   In Number := 0,
  冲销方式_In In Number := 0,
  发药标志_In In Number := 0,
  财务审核_In In Number := 0
) Is
  ----------------------------------------------------------------------------------------
  --功能:根据明细数据更新库存
  --关键：根据下可用库存参数决定是否处理可用数量
  --业务规则：按照模块分开处理数据，便于后期维护
  --过程适用范围：药品流通业务，涉及到产生药品收发记录明细后再更新库存表并重算库存表平均成本价的业务，该过程
  --只能由其他过程内部调用，不能作为单独过程直接执行
  --参数:
  --     Id_In:药品业务新增、删除、审核、冲销时产生收发记录明细的id
  --     Delete_in: 0--非删除操作业务（新增、审核、冲销） 1--删除操作业务
  --     冲销方式_In: 0--正常冲销方式 1-产生冲销申请单据 2-发送 3-回退 目前只有移库模块有效
  --     发药标志_in: 0--不标记  1--标记  此参数只有药品处方、部门发药模块有效
  --     财务审核_in:0,财务审核单据,1-其他业务
  ----------------------------------------------------------------------------------------
  v_下可用数量 Zlparameters.参数值%Type;
  n_可用数量   药品库存.实际数量%Type;
  n_实际数量   药品库存.实际数量%Type;
  n_零售金额   药品库存.实际金额%Type;
  n_差价       药品库存.实际差价%Type;
  n_时价分批   Number(1);
  n_成本价     药品收发记录.成本价%Type;
  n_零售价     药品库存.零售价%Type;

  n_库存数量   药品库存.实际数量%Type;
  n_库存平均价 药品库存.平均成本价%Type;
  n_总数量     药品收发记录.实际数量%Type;
  n_总成本价   药品收发记录.成本价%Type;

  --业务明细数据，把库存数据更新需要的数据都列出来
  Cursor c_Detail Is
    Select a.Id, a.记录状态, a.单据, a.No, a.序号, a.库房id, a.供药单位id, a.入出类别id, a.对方部门id, a.入出系数, Nvl(a.发药方式, 0) As 发药方式, a.药品id,
           Nvl(a.批次, 0) 批次, a.产地, a.批号, a.生产日期, a.效期, a.付数, Nvl(a.填写数量, 0) As 填写数量, a.实际数量, a.成本价, a.成本金额, a.扣率, a.零售价,
           Nvl(a.零售金额, 0) As 零售金额, Nvl(a.差价, 0) As 差价, a.配药人, a.配药日期, a.审核人, a.审核日期, a.灭菌日期, a.灭菌效期, a.批准文号, a.商品条码,
           a.内部条码, b.是否变价, a.单量, a.频次, a.摘要
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.Id = Id_In;

  v_Detail c_Detail%RowType;
Begin
  --取下可用库存参数
  Select zl_GetSysParameter(96) Into v_下可用数量 From Dual;

  For v_Detail In c_Detail Loop
    n_实际数量 := v_Detail.入出系数 * v_Detail.实际数量 * Nvl(v_Detail.付数, 1);
    If n_实际数量 Is Null Then
      n_实际数量 := 0;
    End If;
    n_零售金额 := v_Detail.入出系数 * v_Detail.零售金额;
    n_差价     := v_Detail.入出系数 * v_Detail.差价;
  
    --先取库存和单据的数量和成本价
    Begin
      Select Nvl(实际数量, 0), Nvl(平均成本价, 0)
      Into n_库存数量, n_库存平均价
      From 药品库存
      Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
    Exception
      When Others Then
        n_库存数量   := 0;
        n_库存平均价 := 0;
    End;
  
    --外购入库：正常业务是入库，在填单时不处理可用数量，在审核时处理。特殊的，退库模式时在填单时要根据参数处理可用数量，在审核则相反处理可用数量
    --删除单据时要把填单时预减的加回去
    --冲销时直接按数量加减库存可用数量
    --用数量判断是入库还是退库
    If v_Detail.单据 = 1 Then
      If v_Detail.审核日期 Is Null Then
        --未审核单据，填单或删除
        If Delete_In = 0 Then
          If n_实际数量 < 0 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          Else
            n_可用数量 := 0;
          End If;
        Else
          If n_实际数量 < 0 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := -1 * n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          Else
            n_可用数量 := 0;
          End If;
        End If;
      Else
        --已审核或已冲销
        If v_Detail.记录状态 = 1 Then
          --审核
          If n_实际数量 < 0 Then
            --退库要考虑填单时已经处理了可用数量
            If v_下可用数量 = '1' Then
              n_可用数量 := 0;
            Else
              n_可用数量 := n_实际数量;
            End If;
          Else
            --普通入库
            n_可用数量 := n_实际数量;
          End If;
        Else
          --冲销
          n_可用数量 := n_实际数量;
        End If;
      End If;
    End If;
  
    --自制入库：对于自制药品来说是入库，在填单时不处理可用数量，在审核时处理。特殊的，对于原料药来说是出库，在填单时根据参数处理可用数量，在审核则相反处理可用数量
    --删除单据时要把原料药预减的数量加回去
    --用入出系数判断是入库还是退库
    If v_Detail.单据 = 2 Then
      If v_Detail.审核日期 Is Null Then
        --填单和删除
        If Delete_In = 0 Then
          If v_Detail.入出系数 < 0 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          Else
            n_可用数量 := 0;
          End If;
        Else
          If v_Detail.入出系数 < 0 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := -1 * n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          Else
            n_可用数量 := 0;
          End If;
        End If;
      Else
        --审核和冲销
        If v_Detail.入出系数 < 0 Then
          If v_Detail.记录状态 = 1 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := 0;
            Else
              n_可用数量 := n_实际数量;
            End If;
          Else
            n_可用数量 := n_实际数量;
          End If;
        Else
          n_可用数量 := n_实际数量;
        End If;
      End If;
    End If;
  
    --协定入库：对于协定药品来说是入库，在填单时不处理可用数量，在审核时处理。特殊的，对于组成药来说是出库，在填单时根据参数处理可用数量，在审核则相反处理可用数量
    --删除单据时要把原料药预减的数量加回去
    --用入出系数判断是入库还是退库
    If v_Detail.单据 = 3 Then
      If v_Detail.审核日期 Is Null Then
        --填单和删除
        If Delete_In = 0 Then
          If v_Detail.入出系数 < 0 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          Else
            n_可用数量 := 0;
          End If;
        Else
          If v_Detail.入出系数 < 0 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := -1 * n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          Else
            n_可用数量 := 0;
          End If;
        End If;
      Else
        --审核和冲销
        If v_Detail.入出系数 < 0 Then
          If v_Detail.记录状态 = 1 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := 0;
            Else
              n_可用数量 := n_实际数量;
            End If;
          Else
            n_可用数量 := n_实际数量;
          End If;
        Else
          n_可用数量 := n_实际数量;
        End If;
      End If;
    End If;
  
    --其他入库：正常业务是入库，在填单时不处理可用数量，在审核时处理。特殊的，负数入库模式时要根据参数处理可用数量，在审核时则相反处理可用数量
    --删除单据时要把填单时预减的加回去
    --用数量判断是入库还是退库
    If v_Detail.单据 = 4 Then
      If v_Detail.审核日期 Is Null Then
        --填单和删除
        If Delete_In = 0 Then
          If n_实际数量 < 0 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          Else
            n_可用数量 := 0;
          End If;
        Else
          If n_实际数量 < 0 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := -1 * n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          Else
            n_可用数量 := 0;
          End If;
        End If;
      Else
        If v_Detail.记录状态 = 1 Then
          --审核
          If n_实际数量 < 0 Then
            If v_下可用数量 = '1' Then
              n_可用数量 := 0;
            Else
              n_可用数量 := n_实际数量;
            End If;
          Else
            --普通入库
            n_可用数量 := n_实际数量;
          End If;
        Else
          --冲销
          n_可用数量 := n_实际数量;
        End If;
      End If;
    End If;
  
    --差价调整，成本价调价：不涉及库存数量变化，在填单时不处理，在审核时只处理金额，差价等数据
    If v_Detail.单据 = 5 Then
      n_可用数量 := 0;
    End If;
  
    --移库：移库有两条单据，一条出库单据，一条入库单据；出库单据需要根据下可用库存参数决定是否下可用库存，如果是冲销则可用库存则相反处理
    --在填单不减可用数量时，在发送时预减可用数量，审核时不处理可用数量
    --申请冲销模式时也要根据参数来处理可用数量
    --填单时出库业务根据参数决定是否下库存，入库业务不下库存；删除时出库业务更加参数要把库存还回去，入库业务不还库存
    If v_Detail.单据 = 6 Then
      If v_Detail.审核日期 Is Null Then
        If Delete_In = 0 Then
          --新增、修改、发送、回退、冲销申请
          If v_Detail.记录状态 = 1 Then
            If 冲销方式_In = 2 Then
              --发送
              If v_下可用数量 = '0' And v_Detail.入出系数 = -1 Then
                n_可用数量 := n_实际数量;
              Else
                n_可用数量 := 0;
              End If;
            Elsif 冲销方式_In = 3 Then
              --回退
              If v_下可用数量 = '0' And v_Detail.入出系数 = -1 Then
                n_可用数量 := -1 * n_实际数量;
              Else
                n_可用数量 := 0;
              End If;
            Else
              --新增
              If v_下可用数量 = '1' And v_Detail.入出系数 = -1 Then
                n_可用数量 := n_实际数量;
              Else
                n_可用数量 := 0;
              End If;
            End If;
          Else
            --申请冲销
            If v_下可用数量 = '1' And v_Detail.入出系数 = 1 Then
              n_可用数量 := n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          End If;
        Else
          --删除
          If v_Detail.记录状态 = 1 Then
            If v_下可用数量 = '1' And v_Detail.入出系数 = -1 Then
              n_可用数量 := -1 * n_实际数量;
            Elsif v_Detail.配药日期 Is Not Null And v_Detail.入出系数 = -1 Then
              n_可用数量 := -1 * n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          Else
            If v_下可用数量 = '1' And v_Detail.入出系数 = 1 Then
              n_可用数量 := -1 * n_实际数量;
            Else
              n_可用数量 := 0;
            End If;
          End If;
        End If;
      Else
        If v_Detail.记录状态 = 1 Then
          --审核
          If v_Detail.入出系数 = -1 Then
            --出的那笔
            n_可用数量 := 0;
          Else
            --入的那笔
            n_可用数量 := n_实际数量;
          End If;
        Else
          If 冲销方式_In = 0 Then
            --正常冲销审核
            n_可用数量 := n_实际数量;
          Else
            --申请冲销审核
            If v_下可用数量 = '1' And v_Detail.入出系数 = 1 Then
              n_可用数量 := 0;
            Else
              n_可用数量 := n_实际数量;
            End If;
          End If;
        End If;
      End If;
    End If;
  
    --领用：正常业务是出库，在填单时根据参数处理可用数量，在审核时相反处理
    --删除单据时要把填单时预减的加回去
    If v_Detail.单据 = 7 Then
      If v_Detail.审核日期 Is Null Then
        --填单和删除
        If Delete_In = 0 Then
          If v_下可用数量 = '1' Then
            n_可用数量 := n_实际数量;
          Else
            n_可用数量 := 0;
          End If;
        Else
          If v_下可用数量 = '1' Then
            n_可用数量 := -1 * n_实际数量;
          Else
            n_可用数量 := 0;
          End If;
        End If;
      Else
        --审核和冲销
        If v_Detail.记录状态 = 1 Then
          --审核
          If v_下可用数量 = '1' Then
            n_可用数量 := 0;
          Else
            n_可用数量 := n_实际数量;
          End If;
        Else
          --冲销
          n_可用数量 := n_实际数量;
        End If;
      End If;
    End If;
  
    --发药业务：在填单时固定处理可用数量，在审核时不处理
    --删除单据时要把填单时预减的加回去
    --不再发药标记的可用数量处理等同于删除，填单操作
    If v_Detail.单据 = 8 Or v_Detail.单据 = 9 Or v_Detail.单据 = 10 Then
      If v_Detail.审核日期 Is Null Then
        If Delete_In = 0 Then
          If 发药标志_In = 0 Then
            n_可用数量 := n_实际数量;
          Else
            n_可用数量 := -1 * n_实际数量;
          End If;
        Else
          n_可用数量 := -1 * n_实际数量;
        End If;
      Else
        n_可用数量 := 0;
      End If;
    End If;
  
    --其他出库：正常业务是出库，在填单时根据参数处理可用数量，在审核时相反处理
    --删除单据时要把填单时预减的加回去
    If v_Detail.单据 = 11 Then
      If v_Detail.审核日期 Is Null Then
        --填单和删除
        If Delete_In = 0 Then
          If v_下可用数量 = '1' Then
            n_可用数量 := n_实际数量;
          Else
            n_可用数量 := 0;
          End If;
        Else
          If v_下可用数量 = '1' Then
            n_可用数量 := -1 * n_实际数量;
          Else
            n_可用数量 := 0;
          End If;
        End If;
      Else
        --冲销和审核
        If v_Detail.记录状态 = 1 Then
          --审核
          If v_下可用数量 = '1' Then
            n_可用数量 := 0;
          Else
            n_可用数量 := n_实际数量;
          End If;
        Else
          --冲销
          n_可用数量 := n_实际数量;
        End If;
      End If;
    End If;
  
    --盘点：填单时盘盈业务不处理可用数量，盘亏业务固定处理可用数量，在审核时相反处理
    --删除单据时要把填单时预减的加回去
    --用入出系数区分盘盈盘亏业务
    If v_Detail.单据 = 12 Then
      If v_Detail.审核日期 Is Null Then
        --填单和删除
        If Delete_In = 0 Then
          If v_Detail.入出系数 = -1 Then
            n_可用数量 := n_实际数量;
          Else
            n_可用数量 := 0;
          End If;
        Else
          If v_Detail.入出系数 = -1 Then
            n_可用数量 := -1 * n_实际数量;
          Else
            n_可用数量 := 0;
          End If;
        End If;
      Else
        --审核和冲销
        If v_Detail.记录状态 = 1 Then
          --审核
          If v_Detail.入出系数 = '1' Then
            n_可用数量 := n_实际数量;
          Else
            n_可用数量 := 0;
          End If;
        Else
          --冲销
          n_可用数量 := n_实际数量;
        End If;
      End If;
    End If;
  
    --售价调价：不涉及库存数量变化，在填单时不处理，在审核时只处理金额，差价等数据
    If v_Detail.单据 = 13 Then
      n_可用数量 := 0;
    End If;
  
    --药品留存：产生发药单据时，已经下了库存，部门发药时，需要将库存加回去
    If v_Detail.单据 = 27 Then
      n_可用数量 := n_实际数量;
    End If;
  
    If v_Detail.批次 > 0 And v_Detail.是否变价 = 1 Then
      n_时价分批 := 1;
    Else
      n_时价分批 := 0;
    End If;
  
    n_零售价 := v_Detail.零售价;
    --特殊单据需要处理成本价 特殊单据有单据=5 单据=12
    If v_Detail.单据 = 5 Or v_Detail.单据 = 12 Then
      If v_Detail.单据 = 5 Then
        If v_Detail.填写数量 <> 0 Then
          n_零售价 := Nvl(v_Detail.零售价, 0) / v_Detail.填写数量;
        Else
          n_零售价 := 0;
        End If;
        --审核
        If v_Detail.记录状态 = 1 Then
          --差价调整发药方式=0；主动调价、退货、发药产生的调价修正发药方式=1
          n_成本价 := v_Detail.单量;
        Else
          --冲销 还原原始成本价
          Begin
            --成本价=(金额-差价)/数量
            n_成本价 := (Nvl(v_Detail.零售价, 0) - Nvl(v_Detail.成本价, 0)) / v_Detail.填写数量;
          Exception
            When Others Then
              Select 成本价 Into n_成本价 From 药品规格 Where 药品id = v_Detail.药品id;
          End;
        End If;
      Else
        n_成本价 := v_Detail.单量;
      End If;
    Else
      If v_Detail.单据 = 13 Then
        n_成本价 := Nvl(v_Detail.单量, 0) - Nvl(v_Detail.频次, 0);
      Else
        n_成本价 := v_Detail.成本价;
      End If;
    End If;
  
    --根据业务数据更新库存记录
    If v_Detail.审核日期 Is Null Then
      If n_可用数量 <> 0 Then
        --填单，删除时只更新可用数量
        Update 药品库存
        Set 可用数量 = 可用数量 + n_可用数量
        Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
      
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 效期, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 灭菌效期, 批准文号, 零售价, 上次扣率,
             商品条码, 内部条码, 平均成本价)
          Values
            (v_Detail.库房id, v_Detail.药品id, v_Detail.批次, v_Detail.效期, 1, n_可用数量, 0, 0, 0, v_Detail.供药单位id, n_成本价,
             v_Detail.批号, v_Detail.生产日期, v_Detail.产地, v_Detail.灭菌效期, v_Detail.批准文号, Decode(n_时价分批, 1, n_零售价, Null),
             v_Detail.扣率, v_Detail.商品条码, v_Detail.内部条码, n_成本价);
        End If;
      End If;
    Else
      --审核时更新库存可用数量，实际数量，库存金额，库存差价等数据
      If v_Detail.单据 = 5 Then
        --单据=5 的成本价修正记录 平均成本价不需要重算，因为保存了最新价格的
        If v_Detail.摘要 = '外购退库差价误差自动修正' Or v_Detail.摘要 = '财务审核价格变动修正' Then
          --这一步肯定是外购退库，外购退库只更新成本价,且肯定有库存
          Update 药品库存
          Set 实际差价 = Nvl(实际差价, 0) + n_差价
          Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
        Else
          Update 药品库存
          Set 平均成本价 = n_成本价, 上次采购价 = n_成本价, 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_差价
          Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
          If Sql%NotFound Then
            Insert Into 药品库存
              (库房id, 药品id, 批次, 性质, 实际差价, 上次批号, 效期, 上次产地, 上次供应商id, 上次生产日期, 批准文号, 实际金额, 上次采购价, 平均成本价)
            Values
              (v_Detail.库房id, v_Detail.药品id, v_Detail.批次, 1, n_差价, v_Detail.批号, v_Detail.效期, v_Detail.产地,
               v_Detail.供药单位id, v_Detail.生产日期, v_Detail.批准文号, n_零售金额, n_成本价, n_成本价);
          End If;
        End If;
      Elsif v_Detail.单据 = 13 Then
        --单据=13 的售价修正记录 同步更新的金额和差价，所以不需要重算平均成本价
        Update 药品库存
        Set 零售价 = Decode(n_时价分批, 1, n_零售价, Null), 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_差价
        Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 零售价)
          Values
            (v_Detail.库房id, v_Detail.药品id, v_Detail.批次, 1, 0, 0, n_零售金额, n_零售金额, Decode(n_时价分批, 1, n_零售价, Null));
        Else
          If n_时价分批 = 1 Then
            Update 药品库存
            Set 零售价 = 实际金额 / 实际数量
            Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次 And
                  Nvl(实际数量, 0) <> 0;
          End If;
        End If;
      Else
        --按入库和出库 状态分解
        --入库业务,出库冲销，不分批多种价格入库冲销需要更新库存表所有信息
        If (v_Detail.入出系数 = 1 And v_Detail.记录状态 = 1) Or (v_Detail.入出系数 = -1 And Mod(v_Detail.记录状态, 3) = 2) Or
           (v_Detail.入出系数 = 1 And Mod(v_Detail.记录状态, 3) = 2) Then
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_可用数量, 实际数量 = Nvl(实际数量, 0) + n_实际数量, 实际金额 = Nvl(实际金额, 0) + n_零售金额,
              实际差价 = Nvl(实际差价, 0) + n_差价, 上次供应商id = v_Detail.供药单位id,
              上次采购价 = Decode(v_Detail.单据, 1, Decode(v_Detail.发药方式, 1, 上次采购价, n_成本价), n_成本价),
              上次批号 = Nvl(v_Detail.批号, 上次批号), 上次生产日期 = Nvl(v_Detail.生产日期, 上次生产日期), 上次产地 = Nvl(v_Detail.产地, 上次产地),
              灭菌效期 = Nvl(v_Detail.灭菌效期, 灭菌效期), 效期 = Nvl(v_Detail.效期, 效期), 批准文号 = Nvl(v_Detail.批准文号, 批准文号),
              上次扣率 = Decode(v_Detail.单据, 1, v_Detail.扣率, 上次扣率), 商品条码 = Nvl(v_Detail.商品条码, 商品条码),
              内部条码 = Nvl(v_Detail.内部条码, 内部条码)
          Where 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次 And 性质 = 1;
          --外购入库和其他入库审核时
          If (v_Detail.单据 = 1 And v_Detail.记录状态 = 1 And 财务审核_In = 0) Or (v_Detail.单据 = 4 And v_Detail.记录状态 = 1) Then
            Update 药品库存
            Set 零售价 = Decode(n_时价分批, 1, n_零售价, Null)
            Where 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次 And 性质 = 1;
          End If;
          --不分批入库需要重算成本价
          --外购退货、财务审核和所有冲销业务不更新平均成本价，保持当前价格
          If (v_Detail.单据 = 1 And v_Detail.发药方式 = 1) Or Mod(v_Detail.记录状态, 3) = 2 Or (v_Detail.单据 = 1 And 财务审核_In = 1) Then
            Null;
          Else
            --按总金额/总数量方式计算平均成本价而不用（金额-差价）/数量是为了数据的准确性
            n_总数量 := (n_库存数量 + n_实际数量);
            If n_总数量 <> 0 Then
              n_总成本价 := (n_库存数量 * n_库存平均价 + n_实际数量 * n_成本价) / n_总数量;
              Update 药品库存
              Set 平均成本价 = n_总成本价
              Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
            End If;
          End If;
        Else
          --出库业务只需要更新数量、金额、差价
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_可用数量, 实际数量 = Nvl(实际数量, 0) + n_实际数量, 实际金额 = Nvl(实际金额, 0) + n_零售金额,
              实际差价 = Nvl(实际差价, 0) + n_差价
          Where 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次 And 性质 = 1;
        End If;
        --库存表未找到数据则需要产生库存表所有信息
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 效期, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 灭菌效期, 批准文号, 零售价, 上次扣率,
             商品条码, 内部条码, 平均成本价)
          Values
            (v_Detail.库房id, v_Detail.药品id, v_Detail.批次, v_Detail.效期, 1, n_可用数量, n_实际数量, n_零售金额, n_差价, v_Detail.供药单位id,
             n_成本价, v_Detail.批号, v_Detail.生产日期, v_Detail.产地, v_Detail.灭菌效期, v_Detail.批准文号,
             Decode(n_时价分批, 1, n_零售价, Null), v_Detail.扣率, v_Detail.商品条码, v_Detail.内部条码, n_成本价);
        End If;       
      End If;
    End If;
  
    --删除多余的库存数据
    If 财务审核_In = 0 Then
      Delete From 药品库存
      Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次 And Nvl(可用数量, 0) = 0 And
            Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品库存_Update;
/


--93623:刘兴洪,2016-02-26,更改医嘱发送的计费状态
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
  --场合_In    Integer:=0, --0:门诊;1-住院
  --性质_In    Integer:=1, --1-收费单;2-记帐单
  --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  --No_In      门诊费用记录.No%Type,
  --医嘱ids_In Varchar2 := Null
  Zl_医嘱发送_计费状态_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_销帐;
/

--92736:刘尔旋,2016-02-25,退费申请模式服务窗处理
Create Or Replace Procedure Zl_Third_Charge_Delcheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --功能:三方退费检查 
  --入参:Xml_In: 
  --<IN>
  --    <BRID>病人ID</BRID>
  --    <JE></JE> //退款总金额
  --    <JSKLB></JSKLB>     //结算卡类别
  --    <TFZY>退费摘要</TFZY>
  --    <JCFP>1</JCFP>      //检查发票,0-不检查;1-检查;为1时，打印了发票的单据不能退费
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
  --    ――如无下列错误结点则说明通过检查
  --    <ERROR> 
  --      <MSG>错误信息</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  n_退款总额 门诊费用记录.实收金额%Type;

  n_病人id     门诊费用记录.病人id%Type;
  n_单据病人id 门诊费用记录.病人id%Type;
  v_操作员编码 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_原结算序号 病人预交记录.结算序号%Type;
  v_结算卡类别 Varchar2(100);
  v_结算方式   医疗卡类别.结算方式%Type;
  n_卡类别id   医疗卡类别.Id%Type;

  v_摘要     门诊费用记录.摘要%Type;
  n_Count    Number(18);
  n_Temp     Number(18);
  n_检查发票 Number(3);
  n_是否打印 Number(3);
  n_退费模式 Number(3);

  v_Temp    Varchar2(32767); --临时XML 
  x_Templet Xmltype; --模板XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  --0.获取入参中的病人ID等信息
  Select To_Number(Extractvalue(Value(A), 'IN/BRID')), To_Number(Extractvalue(Value(A), 'IN/JE')),
         Extractvalue(Value(A), 'IN/TFZY'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Extractvalue(Value(A), 'IN/JSKLB')
  Into n_病人id, n_退款总额, v_摘要, n_检查发票, v_结算卡类别
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  --0.相关检查
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份,不允许退费操作!';
    Raise Err_Item;
  End If;

  n_退费模式 := zl_GetSysParameter('门诊退费须先申请');
  If Nvl(n_退费模式, 0) = 1 Then
    v_Err_Msg := '当前为退费申请模式,不允许服务窗退费!';
    Raise Err_Item;
  End If;

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

  --人员id,人员编号,人员姓名 
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '系统不能认别有效的操作员,不允许退费!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员姓名 := v_Temp;
  v_Err_Msg    := Null;

  --1.退费检查

  n_Count      := 0;
  n_原结算序号 := 0;
  For c_费用 In (Select Extractvalue(b.Column_Value, '/FY/DJH') As 单据号, Extractvalue(b.Column_Value, '/FY/XH') As 退款序号
               From Table(Xmlsequence(Extract(Xml_In, '/IN/FYLIST/FY'))) B) Loop
  
    If c_费用.单据号 Is Null Then
      v_Err_Msg := '未确定指定退费的单据号,不能退费!';
      Raise Err_Item;
    End If;
    Begin
      Select a.结算序号, a.结帐id, a.病人id
      Into n_Temp, n_结帐id, n_单据病人id
      From 病人预交记录 A, 门诊费用记录 B
      Where a.结帐id = b.结帐id And b.No = c_费用.单据号 And b.记录性质 = 1 And Nvl(b.费用状态, 0) = 0 And b.记录状态 In (1, 3) And
            Rownum < 2;
    Exception
      When Others Then
        n_Temp := Null;
    End;
  
    If n_Temp Is Null Then
      v_Err_Msg := '指定的单据号:' || c_费用.单据号 || '未找到,不能退费!';
      Raise Err_Item;
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
  
    n_Count := n_Count + 1;
  End Loop;

  If n_Count = 0 Then
    v_Err_Msg := '未确定本次需要退费的单据,不能退费!';
    Raise Err_Item;
  End If;

  --2.支付方式检查
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
      Null;
    Elsif c_结算方式.卡类别 Is Not Null And Nvl(c_结算方式.是否消费卡, 0) = 1 Then
      --2.消费卡结算
      Null;
    Elsif Nvl(c_结算方式.是否退预交, 0) = 1 Then
      --3.退预交款
      Null;
    Else
      --4.普通结算
      If c_结算方式.结算方式 Is Null Then
        v_Err_Msg := '未指定支付方式,不能退费!';
        Raise Err_Item;
      End If;
    End If;
    n_Count := n_Count + 1;
  End Loop;

  If n_Count = 0 Then
    v_Err_Msg := '不能有效确认当前的支付方式,不能退费!';
    Raise Err_Item;
  End If;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Charge_Delcheck;
/

--93623:刘兴洪,2016-02-26,调整医嘱发送的计费状态
Create Or Replace Procedure Zl_医嘱发送_计费状态_Update
(
  场合_In    Integer := 0, --0:门诊;1-住院
  性质_In    Integer := 1, --1-收费单;2-记帐单
  操作_In    Integer := 0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  No_In      门诊费用记录.No%Type,
  医嘱ids_In Varchar2 := Null
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_医嘱id   t_Numlist := t_Numlist();
  l_医嘱id1  t_Numlist := t_Numlist();
  l_计费状态 t_Numlist := t_Numlist();
  n_Count    Number(18);

Begin
  If 医嘱ids_In Is Null Then
    If 场合_In = 0 Then
      Select ID Bulk Collect
      Into l_医嘱id1
      From (Select Distinct 医嘱序号 As ID
             From 门诊费用记录
             Where NO = No_In And 记录性质 = 性质_In And 记录状态 In (0, 1, 3) And 医嘱序号 Is Not Null);
    Else
      Select ID Bulk Collect
      Into l_医嘱id1
      From (Select Distinct 医嘱序号 As ID
             From 住院费用记录
             Where NO = No_In And 记录性质 = 性质_In And 记录状态 In (0, 1, 3) And 医嘱序号 Is Not Null);
    End If;
  Else
    Select ID Bulk Collect
    Into l_医嘱id1
    From (Select Distinct Column_Value As ID From Table(f_Str2list(医嘱ids_In)) B);
  End If;
  If l_医嘱id1.Count = 0 Then
    Return;
  End If;

  For c_医嘱 In (With c_医嘱信息 As
                  (Select Column_Value As 医嘱id From Table(l_医嘱id1))
                 Select ID, 相关id, 诊疗类别
                 From 病人医嘱记录 A
                 Where a.Id In (Select 医嘱id
                                From c_医嘱信息
                                Union All
                                Select Distinct 相关id
                                From 病人医嘱记录 A1, c_医嘱信息 B1
                                Where A1.Id = B1.医嘱id And 相关id Is Not Null)) Loop
    --1.划价单删除
    If Nvl(操作_In, 0) = 0 Then
      --D    检查       JC
      --F    手术       SS
      --:-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费，2-部分退费(销帐)，3-全部收费(仅门诊收费有)，4-全部退费(销帐)
      If c_医嘱.相关id Is Null And Instr(',D,F,', ',' || c_医嘱.诊疗类别 || ',') > 0 Then
        If 场合_In = 0 Then
          Select Nvl(Count(*), 0)
          Into n_Count
          From 门诊费用记录 A, (Select ID From 病人医嘱记录 A Where ID = c_医嘱.Id Or 相关id = c_医嘱.Id) B, 病人医嘱发送 C
          Where a.医嘱序号 = b.Id And a.No = No_In And a.记录性质 = 性质_In And a.No = c.No And a.记录性质 = c.记录性质 And c.计费状态 = 1 And
                Nvl(a.附加标志, 0) <> 9;
        Else
          Select Nvl(Count(*), 0)
          Into n_Count
          From 住院费用记录 A, (Select ID From 病人医嘱记录 A Where ID = c_医嘱.Id Or 相关id = c_医嘱.Id) B, 病人医嘱发送 C
          Where a.医嘱序号 = b.Id And a.No = No_In And a.记录性质 = 性质_In And a.No = c.No And a.记录性质 = c.记录性质 And c.计费状态 = 1 And
                Nvl(a.附加标志, 0) <> 9;
        End If;
      Else
        If 场合_In = 0 Then
          Select Nvl(Count(*), 0)
          Into n_Count
          From 门诊费用记录 A, (Select ID From 病人医嘱记录 A Where ID = c_医嘱.Id) B, 病人医嘱发送 C
          Where a.医嘱序号 = b.Id And a.No = No_In And a.记录性质 = 性质_In And a.No = c.No And a.记录性质 = c.记录性质 And c.计费状态 = 1 And
                Nvl(a.附加标志, 0) <> 9;
        Else
          Select Nvl(Count(*), 0)
          Into n_Count
          From 住院费用记录 A, (Select ID From 病人医嘱记录 A Where ID = c_医嘱.Id Or 相关id = c_医嘱.Id) B, 病人医嘱发送 C
          Where a.医嘱序号 = b.Id And a.No = No_In And a.记录性质 = 性质_In And a.No = c.No And a.记录性质 = c.记录性质 And c.计费状态 = 1 And
                Nvl(a.附加标志, 0) <> 9;
        End If;
      End If;
      If Nvl(n_Count, 0) = 0 Then
        l_医嘱id.Extend;
        l_医嘱id(l_医嘱id.Count) := c_医嘱.Id;
        l_计费状态.Extend;
        l_计费状态(l_计费状态.Count) := 0; --未计费
      
      End If;
    End If;
  
    --1.收费或记帐
    If Nvl(操作_In, 0) = 1 Then
      --:-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费，2-部分退费(销帐)，3-全部收费(仅门诊收费有)，4-全部退费(销帐)
      n_Count := 0;
      If c_医嘱.相关id Is Null And Instr(',D,F,', ',' || c_医嘱.诊疗类别 || ',') = 0 Then
        n_Count := 1;
      End If;
      If Nvl(n_Count, 0) = 0 Then
        l_医嘱id.Extend;
        l_医嘱id(l_医嘱id.Count) := c_医嘱.Id;
        l_计费状态.Extend;
        l_计费状态(l_计费状态.Count) := 3; --全部收费
      
      End If;
    End If;
  
    --2.退费或销帐
    If Nvl(操作_In, 0) = 2 Then
      --:-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费，2-部分退费(销帐)，3-全部收费(仅门诊收费有)，4-全部退费(销帐)
      If c_医嘱.相关id Is Null And Instr(',D,F,', ',' || c_医嘱.诊疗类别 || ',') > 0 Then
        If 场合_In = 0 Then
          Select Case
                   When Min(状态) = Max(状态) Then
                    Max(状态)
                   Else
                    3
                 End
          Into n_Count
          From (
                 
                 Select Case
                            When 剩余数量 = 原始数量 And 原始数量 <> 0 Then
                             0
                            When 剩余数量 = 0 Then
                             2
                            Else
                             3
                          End As 状态
                 
                 From (Select 序号, Sum(Nvl(Nvl(付数, 1) * 数次, 0)) As 剩余数量,
                                Nvl(Sum(Decode(记录状态, 1, 1, 3, 1, 0) * Decode(记录性质, 11, 0, 1)  * Nvl(Nvl(付数, 1) * 数次, 0)), 0) As 原始数量
                         From 门诊费用记录
                         Where Decode(性质_In, 2, 记录性质, Mod(记录性质, 10)) = 性质_In And NO = No_In And 价格父号 Is Null And
                               Nvl(附加标志, 0) <> 9 And
                               医嘱序号 + 0 In (Select ID From 病人医嘱记录 Where ID = c_医嘱.Id Or 相关id = c_医嘱.Id)
                         Group By 序号));
        Else
          Select Case
                   When Min(状态) = Max(状态) Then
                    Max(状态)
                   Else
                    3
                 End
          Into n_Count
          From (
                 
                 Select Case
                           When 剩余数量 = 原始数量 And 原始数量 <> 0 Then
                            0
                           When 剩余数量 = 0 Then
                            2
                           Else
                            3
                         End As 状态
                 
                 From (Select 序号, Sum(Nvl(Nvl(付数, 1) * 数次, 0)) As 剩余数量,
                                Nvl(Sum(Decode(记录状态, 1, 1, 3, 1, 0) * Decode(记录性质, 11, 0, 1) * Nvl(Nvl(付数, 1) * 数次, 0)), 0) As 原始数量
                         From 住院费用记录
                         Where Decode(性质_In, 2, 记录性质, Mod(记录性质, 10)) = 性质_In And NO = No_In And 价格父号 Is Null And
                               Nvl(附加标志, 0) <> 9 And
                               医嘱序号 + 0 In (Select ID From 病人医嘱记录 Where ID = c_医嘱.Id Or 相关id = c_医嘱.Id)
                         Group By 序号));
        
        End If;
      Else
        If 场合_In = 0 Then
          Select Case
                   When Min(状态) = Max(状态) Then
                    Max(状态)
                   Else
                    3
                 End
          Into n_Count
          From (
                 
                 Select Case
                            When 剩余数量 = 原始数量 And 原始数量 <> 0 Then
                             0
                            When 剩余数量 = 0 Then
                             2
                            Else
                             3
                          End As 状态
                 From (Select 序号, Sum(Nvl(Nvl(付数, 1) * 数次, 0)) As 剩余数量,
                                Nvl(Sum(Decode(记录状态, 1, 1, 3, 1, 0) * Decode(记录性质, 11, 0, 1) * Nvl(Nvl(付数, 1) * 数次, 0)), 0) As 原始数量
                         From 门诊费用记录
                         Where Decode(性质_In, 2, 记录性质, Mod(记录性质, 10)) = 性质_In And NO = No_In And 价格父号 Is Null And
                               Nvl(附加标志, 0) <> 9 And 医嘱序号 + 0 = c_医嘱.Id
                         Group By 序号));
        Else
        
          Select Case
                   When Min(状态) = Max(状态) Then
                    Max(状态)
                   Else
                    3
                 End
          Into n_Count
          From (
                 
                 Select Case
                           When 剩余数量 = 原始数量 And 原始数量 <> 0 Then
                            0
                           When 剩余数量 = 0 Then
                            2
                           Else
                            3
                         End As 状态
                 From (Select 序号, Sum(Nvl(Nvl(付数, 1) * 数次, 0)) As 剩余数量,
                                Nvl(Sum(Decode(记录状态, 1, 1, 3, 1, 0) * Decode(记录性质, 11, 0, 1) * Nvl(Nvl(付数, 1) * 数次, 0)), 0) As 原始数量
                         From 住院费用记录
                         Where Decode(性质_In, 2, 记录性质, Mod(记录性质, 10)) = 性质_In And NO = No_In And 价格父号 Is Null And
                               Nvl(附加标志, 0) <> 9 And 医嘱序号 + 0 = c_医嘱.Id
                         Group By 序号));
        
        End If;
      
      End If;
    
      If n_Count <> 0 Then
      
        l_医嘱id.Extend;
        l_医嘱id(l_医嘱id.Count) := c_医嘱.Id;
        l_计费状态.Extend;
        l_计费状态(l_计费状态.Count) := Case
                                  When Nvl(n_Count, 0) = 2 Then
                                   4
                                  Else
                                   2
                                End;
      End If;
    End If;
  End Loop;

  Forall I In 1 .. l_医嘱id.Count
    Update 病人医嘱发送 A
    Set a.计费状态 = l_计费状态(I)
    Where 医嘱id = l_医嘱id(I) And 记录性质 = 性质_In And NO = No_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_医嘱发送_计费状态_Update;
/

--93623:刘兴洪,2016-02-25,更改医嘱发送的计费状态
Create Or Replace Procedure Zl_门诊划价记录_Delete
(
  No_In   门诊费用记录.No%Type,
  序号_In Varchar2 := Null --主要用于门诊医生站作废单个药品
) As
  --功能：删除一张门诊划价单据
  --该光标用于处理药品库存可用数量
  Cursor c_Stock Is
    Select 发药方式, 库房id, 批次, 药品id, 实际数量, 付数, 灭菌效期, 产地, 批号, 效期, ID, 商品条码, 内部条码, 费用id
    From 药品收发记录
    Where 单据 In (8, 24) And NO = No_In And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 门诊费用记录
                   Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 And 收费类别 In ('4', '5', '6', '7') And
                         (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
    Order By 药品id;
  --该游标用于处理费用记录序号
  Cursor c_Serial Is
    Select ID, 价格父号 From 门诊费用记录 Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 Order By 序号;

  v_医嘱ids  Varchar2(4000);
  l_医嘱id   t_Numlist := t_Numlist();
  l_药品收发 t_Numlist := t_Numlist();
  v_医嘱id   病人医嘱记录.Id%Type;
  l_费用id   t_Numlist := t_Numlist();
  n_备货卫材 Number;

  n_父号         门诊费用记录.序号%Type;
  n_Count        Number;
  n_医嘱数       Number(5);
  n_已执行_Count Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin

  --是否已经删除或收费
  Select Nvl(Count(ID), 0), Sum(Decode(医嘱序号, Null, 0, 1)), Max(医嘱序号), Sum(Decode(Nvl(执行状态, 0), 1, 1, 2, 1, 0))
  Into n_Count, n_医嘱数, v_医嘱id, n_已执行_Count
  From 门诊费用记录
  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And
        (Instr(',' || 序号_In || ',', ',' || Nvl(价格父号, 序号) || ',') > 0 Or 序号_In Is Null);

  If n_Count = 0 Then
    v_Err_Msg := '要删除的费用记录不存在，可能已经删除或已经收费。';
    Raise Err_Item;
  End If;
  --是否已经执行
  If Nvl(n_已执行_Count, 0) > 0 Then
    v_Err_Msg := '要删除的费用记录中包含已执行的内容！';
    Raise Err_Item;
  End If;

  --医嘱费用：检查正在执行的医嘱(注意已执行的情况在下面检查,因为不传 序号_IN 这种情况费用界面已限制)
  Select Nvl(Count(*), 0)
  Into n_Count
  From 病人医嘱发送
  Where 执行状态 = 3 And (NO, 记录性质, 医嘱id) In
        (Select NO, 记录性质, 医嘱序号
                      From 门诊费用记录
                      Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 And 医嘱序号 Is Not Null And
                            (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null));
  If n_Count > 0 Then
    v_Err_Msg := '要删除的费用中存在对应的医嘱正在执行的情况，不能删除！';
    Raise Err_Item;
  End If;

  --药品相关内容
  --先处理备货材料
  For v_出库 In (Select 发药方式, 库房id, 批次, 药品id, 实际数量, 付数, 灭菌效期, 产地, 批号, 效期, ID, 商品条码, 内部条码, 费用id
               From 药品收发记录
               Where 单据 = 21 And Mod(记录状态, 3) = 1 And 审核人 Is Null And
                     费用id In (Select ID
                              From 门诊费用记录
                              Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 And 收费类别 = '4' And
                                    (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
               Order By 药品id) Loop
  
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
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := v_出库.Id;
  
    l_费用id.Extend;
    l_费用id(l_费用id.Count) := v_出库.费用id;
  End Loop;

  For r_Stock In c_Stock Loop
  
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
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
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
    Delete From 药品收发记录 Where ID = l_药品收发(I);

  ------------------------------------------------------------------------------------------------------------------------
  --批量删未发药品记录
  Delete From 未发药品记录 A
  Where NO = No_In And 单据 In (8, 24) And Not Exists
   (Select 1
         From 药品收发记录
         Where 单据 = a.单据 And Nvl(库房id, 0) = Nvl(a.库房id, 0) And NO = No_In And Mod(记录状态, 3) = 1 And 审核人 Is Null);
  ------------------------------------------------------------------------------------------------------------------------

  --删除病人医嘱附费(最后一次删除时)
  If 序号_In Is Null Then
    --Begin
    --  Select 医嘱序号
    --  Into v_医嘱id
    --  From 门诊费用记录
    --  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And Rownum = 1;
    -- Exception
    --  When Others Then
    --    Null;
    -- End;
  
    If v_医嘱id Is Not Null Then
      Delete From 病人医嘱附费 Where 医嘱id = v_医嘱id And NO = No_In And 记录性质 = 1;
    End If;
  End If;

  If n_医嘱数 > 0 Then
    If n_医嘱数 = 1 Then
      l_医嘱id.Extend;
      l_医嘱id(l_医嘱id.Count) := v_医嘱id;
    Else
      Select Distinct 医嘱序号 Bulk Collect
      Into l_医嘱id
      From 门诊费用记录
      Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And 医嘱序号 Is Not Null And
            (Instr(',' || 序号_In || ',', ',' || Nvl(价格父号, 序号) || ',') > 0 Or 序号_In Is Null);
    End If;
  End If;

  --门诊费用记录
  Delete From 门诊费用记录
  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And
        (Instr(',' || 序号_In || ',', ',' || Nvl(价格父号, 序号) || ',') > 0 Or 序号_In Is Null);
  If Sql%RowCount = 0 Then
    v_Err_Msg := '要删除的费用记录不存在，可能已经删除或已经收费。';
    Raise Err_Item;
  End If;

  If 序号_In Is Not Null Then
    --重新调整剩余费用费用记录的序号
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.价格父号 Is Null Then
        n_父号 := n_Count;
      End If;
      Update 门诊费用记录 Set 序号 = n_Count, 价格父号 = Decode(价格父号, Null, Null, n_父号) Where ID = r_Serial.Id;
    
      n_Count := n_Count + 1;
    End Loop;
  End If;
  v_医嘱ids := Null;
  For I In 1 .. l_医嘱id.Count Loop
    v_医嘱ids := Nvl(v_医嘱ids, '') || ',' || l_医嘱id(I);
  End Loop;
  If v_医嘱ids Is Not Null Then
    v_医嘱ids := Substr(v_医嘱ids, 2);
    --场合_In    Integer, --0:门诊;1-住院
    --性质_In    Integer, --1-收费单;2-记帐单
    --操作_In    Integer, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2
    Zl_医嘱发送_计费状态_Update(0, 1, 0, No_In, v_医嘱ids);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊划价记录_Delete;
/

--93623:刘兴洪,2016-02-25,更改医嘱发送的计费状态
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
  --场合_In    Integer:=0, --0:门诊;1-住院
  --性质_In    Integer:=1, --1-收费单;2-记帐单
  --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  --No_In      门诊费用记录.No%Type,
  --医嘱ids_In Varchar2 := Null
  Zl_医嘱发送_计费状态_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊简单收费_Delete;
/

--93623:刘兴洪,2016-02-25,更改医嘱发送的计费状态
Create Or Replace Procedure Zl_门诊划价记录_Clear(Day_In Number) As
  --功能：自动清除划价单
  --参数：Day_IN=删除划价后超过Day_IN天未收费的单据
  Cursor c_Price Is
    Select Distinct a.No, f_List2str(Cast(Collect(To_Char(a.序号)) As t_Strlist)) As 序号
    From 门诊费用记录 A, 未发药品记录 B
    Where a.记录性质 = 1 And a.记录状态 = 0 And a.执行状态 Not In (1, 2) And a.划价人 Is Not Null And a.操作员姓名 Is Null And
          b.单据 In (8, 24) And Nvl(b.已收费, 0) = 0 And a.No = b.No And Nvl(a.执行部门id, 0) = Nvl(b.库房id, 0) And
          Sysdate - b.填制日期 >= Day_In
    Group By a.No;
Begin
  For r_Price In c_Price Loop
    If Not r_Price.序号 Is Null Then
      Zl_门诊划价记录_Delete(r_Price.No, r_Price.序号);
      Commit;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊划价记录_Clear;
/

--93623:刘兴洪,2016-02-25,更改医嘱发送的计费状态
Create Or Replace Procedure Zl_病人划价收费_Insert
(
  No_In         门诊费用记录.No%Type,
  病人id_In     门诊费用记录.病人id%Type,
  病人来源_In   Number,
  付款方式_In   门诊费用记录.付款方式%Type,
  姓名_In       门诊费用记录.姓名%Type,
  性别_In       门诊费用记录.性别%Type,
  年龄_In       门诊费用记录.年龄%Type,
  病人科室id_In 门诊费用记录.病人科室id%Type,
  开单部门id_In 门诊费用记录.开单部门id%Type,
  开单人_In     门诊费用记录.开单人%Type,
  结帐id_In     门诊费用记录.结帐id%Type,
  发生时间_In   门诊费用记录.发生时间%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  发药窗口_In   Varchar2 := Null,
  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
  登记时间_In   门诊费用记录.登记时间%Type := Null
) As
  --功能：用于收费时收取划价单费用
  --参数：
  --      发药窗口_In:执行部门ID1|发药窗口1;...;执行部门IDn|发药窗口n

  --        病人来源_IN:1-门诊;2-住院
  --说明：
  --        1.收取划价费用时,才计算费用相关汇总,在划价时不处理;但药品相关汇总(姓名除外)划价时已经计算。
  --        2.收取划价费用时,目前界面及过程中未处理加收工本费,由划价时直接处理。
  --该游标为划价原单据内容
  Cursor c_Price Is
    Select ID
    From 门诊费用记录
    Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 And 操作员姓名 Is Null
    Order By 序号;

  n_Array_Size Number := 200;
  t_费用id     t_Numlist;
  v_部门名称   部门表.名称%Type;

  v_标识号   门诊费用记录.标识号%Type;
  v_付款方式 医疗付款方式.名称%Type;

  --临时变量
  n_Count      Number;
  n_新病人模式 Number;
  v_出库no     药品收发记录.No%Type;
  v_Date       Date;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  n_组id 财务缴款分组.Id%Type;

Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  Select Count(ID)
  Into n_Count
  From 门诊费用记录
  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And 操作员姓名 Is Null;
  If n_Count = 0 Then
    v_Err_Msg := '不能读取划价单内容,该单据可能已经删除或已经收费！';
    Raise Err_Item;
  End If;
  v_Date := 登记时间_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    Select Decode(当前科室id, Null, 门诊号, 住院号) Into v_标识号 From 病人信息 Where 病人id = 病人id_In;
  End If;

  ------------------------------------------------------------------------------------------------------------------------
  --批量更新
  Open c_Price;
  Loop
    Fetch c_Price Bulk Collect
      Into t_费用id Limit n_Array_Size;
    Exit When t_费用id.Count = 0;
  
    --循环处理门诊费用记录
    Forall I In 1 .. t_费用id.Count
    --执行状态相关字段不处理,在划价时处理;因为可能未收费发药,这种已执行的划价单是允许收费操作的。
    --为保证与预交结算记录的时间相同,重新填写登记时间,但药品部分不变动。
      Update 门诊费用记录
      Set 记录状态 = 1, 病人id = Decode(病人id_In, 0, Null, 病人id_In), 标识号 = v_标识号, 付款方式 = 付款方式_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
          性别 = 性别_In,
          --可能保持医嘱发送的内容
          病人科室id = Nvl(病人科室id_In, 病人科室id), 开单部门id = Nvl(开单部门id_In, 开单部门id), 开单人 = Nvl(开单人_In, 开单人), 结帐金额 = 实收金额,
          结帐id = 结帐id_In, 发生时间 = 发生时间_In, 登记时间 = v_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 是否急诊 = 是否急诊_In,
          缴款组id = n_组id, 费用状态 = 1, 执行状态 = Decode(Nvl(执行状态, 0), -1, Null, Nvl(执行状态, 0))
      Where ID = t_费用id(I) And 记录状态 = 0;
  
    If Sql%RowCount <> t_费用id.Count Then
      v_Err_Msg := '由于并发操作,该单据可能已经删除或已经收费！';
      Raise Err_Item;
    End If;
  End Loop;

  Close c_Price;

  --相关汇总表的处理
  --药品部分非费用信息的修改
  --药品未发记录(如果已发药则修改不到),分离发药时无库房ID
  --可能存在材料和药品库房相同，但材料无发药窗口
  Update 未发药品记录
  Set 病人id = Decode(病人id_In, 0, Null, 病人id_In), 姓名 = 姓名_In, 对方部门id = 开单部门id_In, 已收费 = 1, 填制日期 = v_Date
  Where 单据 = 24 And NO = No_In And
        Nvl(库房id, 0) In (Select Distinct Nvl(执行部门id, 0)
                         From 门诊费用记录
                         Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 = '4');

  Update 未发药品记录
  Set 病人id = Decode(病人id_In, 0, Null, 病人id_In), 姓名 = 姓名_In, 对方部门id = 开单部门id_In, 已收费 = 1, 填制日期 = v_Date
  Where 单据 = 8 And NO = No_In And
        Nvl(库房id, 0) In (Select Distinct Nvl(执行部门id, 0)
                         From 门诊费用记录
                         Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));

  --药品收发记录(可能已经发药或取消发药,所有记录更改)
  Update 药品收发记录
  Set 对方部门id = 开单部门id_In, 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date)
  Where 单据 = 24 And NO = No_In And
        费用id + 0 In (Select ID From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 = '4');

  -------------------------------------------------------------------------------------------
  --处理备货卫材
  n_Count := Null;
  Begin
    Select Count(*), Max(a.No)
    Into n_Count, v_出库no
    From 药品收发记录 A, 门诊费用记录 B
    Where a.费用id = b.Id And b.收费类别 = '4' And b.记录性质 = 1 And b.记录状态 = 1 And b.No = No_In And
          Instr(',8,9,10,21,24,25,26,', ',' || a.单据 || ',') > 0 And Rownum <= 1;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(n_Count, 0) > 0 Then
    If Nvl(病人科室id_In, 0) <> 0 Then
      Select 名称 Into v_部门名称 From 部门表 Where ID = 病人科室id_In;
    End If;
    v_Err_Msg := LPad(' ', 4);
    v_Err_Msg := Substr('病人姓名:' || 姓名_In || v_Err_Msg || '性别:' || 性别_In || v_Err_Msg || '年龄' || 年龄_In || v_Err_Msg ||
                        '门诊号:' || Nvl(v_标识号, '') || v_Err_Msg || '病人科室:' || v_部门名称, 1, 100);
  
    Update 药品收发记录
    Set 对方部门id = 开单部门id_In, 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date), 摘要 = v_Err_Msg
    Where 单据 = 21 And NO = v_出库no And
          费用id + 0 In (Select ID From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 = '4');
  End If;

  Update 药品收发记录
  Set 对方部门id = 开单部门id_In, 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date)
  Where 单据 = 8 And NO = No_In And
        费用id + 0 In
        (Select ID From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));

  If Not 发药窗口_In Is Null Then
    --更新发药窗口
    For v_窗口 In (Select To_Number(C1) As C1, C2 From Table(f_Str2list2(发药窗口_In, ';', '|'))) Loop
    
      Update 门诊费用记录
      Set 发药窗口 = Nvl(v_窗口.C2, 发药窗口)
      Where NO = No_In And 记录性质 = 1 And 记录状态 = 1 And 执行部门id = Nvl(v_窗口.C1, 执行部门id) And 收费类别 In ('5', '6', '7');
    
      Update 药品收发记录
      Set 发药窗口 = Nvl(v_窗口.C2, 发药窗口)
      Where 单据 = 8 And NO = No_In And 库房id = Nvl(v_窗口.C1, 库房id) And
            费用id + 0 In (Select ID
                         From 门诊费用记录
                         Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));
    
      Update 未发药品记录
      Set 发药窗口 = Nvl(v_窗口.C2, 发药窗口)
      Where 单据 = 8 And NO = No_In And 库房id = Nvl(v_窗口.C1, 库房id) And
            Nvl(库房id, 0) In (Select Distinct Nvl(执行部门id, 0)
                             From 门诊费用记录
                             Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));
    
    End Loop;
  End If;

  --更新部份病人信息
  If 病人id_In Is Not Null Then
    If 付款方式_In Is Not Null And 病人来源_In = 1 Then
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In;
    End If;
    --通过划价单收费时不允许改费别,因为费用不允许变
    Update 病人信息
    Set 性别 = Decode(姓名, '新病人', Nvl(性别_In, 性别), 性别), 年龄 = Decode(姓名, '新病人', Nvl(年龄_In, 年龄), 年龄),
        姓名 = Decode(姓名, '新病人', 姓名_In, 姓名), 医疗付款方式 = Nvl(v_付款方式, 医疗付款方式)
    Where 病人id = 病人id_In;
    Select Zl_To_Number(Nvl(zl_GetSysParameter('自动产生姓名', '1111'), '0')) Into n_新病人模式 From Dual;
    If n_新病人模式 = 1 Then
    
      Update 病人挂号记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
    
      Update 门诊费用记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 付款方式 = 付款方式_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
    End If;
  End If;

  --医嘱处理
  --场合_In    Integer:=0, --0:门诊;1-住院
  --性质_In    Integer:=1, --1-收费单;2-记帐单
  --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  --No_In      门诊费用记录.No%Type,
  --医嘱ids_In Varchar2 := Null
  Zl_医嘱发送_计费状态_Update(0, 1, 1, No_In);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人划价收费_Insert;
/

--93623:刘兴洪,2016-02-25,更改医嘱发送的计费状态
Create Or Replace Procedure Zl_划价收费记录_Insert
(
  No_In         门诊费用记录.No%Type,
  病人id_In     门诊费用记录.病人id%Type,
  病人来源_In   Number,
  付款方式_In   门诊费用记录.付款方式%Type,
  姓名_In       门诊费用记录.姓名%Type,
  性别_In       门诊费用记录.性别%Type,
  年龄_In       门诊费用记录.年龄%Type,
  病人科室id_In 门诊费用记录.病人科室id%Type,
  开单部门id_In 门诊费用记录.开单部门id%Type,
  开单人_In     门诊费用记录.开单人%Type,
  收费结算_In   Varchar2,
  冲预交额_In   病人预交记录.冲预交%Type,
  保险结算_In   Varchar2,
  结帐id_In     门诊费用记录.结帐id%Type,
  发生时间_In   门诊费用记录.发生时间%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  发药窗口_In   Varchar2,
  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
  缴款_In       病人预交记录.缴款%Type := Null,
  找补_In       病人预交记录.找补%Type := Null,
  三方卡结算_In Varchar2 := Null,
  登记时间_In   门诊费用记录.登记时间%Type := Null,
  结算序号_In   病人预交记录.结算序号%Type := Null,
  交易流水号_In 病人预交记录.交易流水号%Type := Null,
  交易说明_In   病人预交记录.交易说明%Type := Null,
  简单收费_In   Number := 0
) As
  --功能：用于收费时收取划价单费用
  --参数：
  --      发药窗口_In:执行部门ID1|发药窗口1;...;执行部门IDn|发药窗口n
  --        病人来源_IN:1-门诊;2-住院
  --        收费结算_IN:格式="结算方式|结算金额|结算号码|结算摘要||.....",注意无结算号码和摘要时要用空格填充
  --        保险结算_IN:格式="结算方式|结算金额||....."
  --        三方卡结算_In:格式=卡类别Id|是否消费卡|结算金额|卡号|备注||...
  --        交易流水号_In和交易说明_In:收费结算_IN时有效.
  --说明：
  --        1.收取划价费用时,才计算费用相关汇总,在划价时不处理;但药品相关汇总(姓名除外)划价时已经计算。
  --        2.收取划价费用时,目前界面及过程中未处理加收工本费,由划价时直接处理。
  --该游标为划价原单据内容

  --=================================
  --备注：该过程目前只有简单收费使用！
  --=================================

  Cursor c_Price Is
    Select ID
    From 门诊费用记录
    Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 And 操作员姓名 Is Null
    Order By 序号;

  n_Array_Size Number := 200;
  t_费用id     t_Numlist;
  v_部门名称   部门表.名称%Type;
  --该游标用于收费冲预交的可用预交列表(该SQL参考住院结帐)
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit(v_病人id 病人信息.病人id%Type) Is
    Select *
    From (Select a.Id, a.记录状态, Nvl(a.预交类别, 2) As 预交类别, a.No, Nvl(a.金额, 0) As 金额
           From 病人预交记录 A,
                (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                  From 病人预交记录 A
                  Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.病人id = v_病人id And Nvl(预交类别, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.金额, 0)) <> 0) B
           Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.No = b.No And a.病人id = v_病人id And Nvl(a.预交类别, 2) = 1
           Union All
           Select 0 As ID, 记录状态, Nvl(预交类别, 2) As 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
           From 病人预交记录
           Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(预交类别, 2) = 1 And Nvl(金额, 0) <> Nvl(冲预交, 0) And
                 病人id = v_病人id Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
           Group By 记录状态, NO, Nvl(预交类别, 2))
    Order By ID, 预交类别 Desc, NO;

  --预交与结算相关变量
  v_预交金额 病人预交记录.冲预交%Type;
  v_结算内容 Varchar2(3000);
  v_当前结算 Varchar2(150);
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;

  v_标识号   门诊费用记录.标识号%Type;
  v_付款方式 医疗付款方式.名称%Type;
  n_返回值   病人余额.预交余额%Type;

  --临时变量
  n_Count      Number;
  n_新病人模式 Number;
  v_出库no     药品收发记录.No%Type;
  v_Date       Date;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  n_组id     财务缴款分组.Id%Type;
  n_卡类别id 医疗卡类别.Id%Type;
  n_消费卡   Number;
  v_卡号     病人预交记录.卡号%Type;
  v_卡名称   Varchar2(100);
  n_自制卡   Number;
  n_预交id   病人预交记录.Id%Type;
  n_消费卡id 消费卡目录.Id%Type;
  n_金额     病人预交记录.金额%Type;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  Select Count(ID)
  Into n_Count
  From 门诊费用记录
  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And 操作员姓名 Is Null;
  If n_Count = 0 Then
    v_Err_Msg := '不能读取划价单内容,该单据可能已经删除或已经收费！';
    Raise Err_Item;
  End If;
  v_Date := 登记时间_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    Select Decode(当前科室id, Null, 门诊号, 住院号) Into v_标识号 From 病人信息 Where 病人id = 病人id_In;
  End If;

  ------------------------------------------------------------------------------------------------------------------------
  --批量更新
  Open c_Price;
  Loop
    Fetch c_Price Bulk Collect
      Into t_费用id Limit n_Array_Size;
    Exit When t_费用id.Count = 0;
  
    --循环处理门诊费用记录
    Forall I In 1 .. t_费用id.Count
    --执行状态相关字段不处理,在划价时处理;因为可能未收费发药,这种已执行的划价单是允许收费操作的。
    --为保证与预交结算记录的时间相同,重新填写登记时间,但药品部分不变动。
      Update 门诊费用记录
      Set 记录状态 = 1, 病人id = Decode(病人id_In, 0, Null, 病人id_In), 标识号 = v_标识号, 付款方式 = 付款方式_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
          性别 = 性别_In,
          --可能保持医嘱发送的内容
          病人科室id = Nvl(病人科室id_In, 病人科室id), 开单部门id = Nvl(开单部门id_In, 开单部门id), 开单人 = Nvl(开单人_In, 开单人), 结帐金额 = 实收金额,
          结帐id = 结帐id_In, 发生时间 = 发生时间_In, 登记时间 = v_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 是否急诊 = 是否急诊_In,
          缴款组id = n_组id
      Where ID = t_费用id(I) And 记录状态 = 0;
  
    If Sql%RowCount <> t_费用id.Count Then
      v_Err_Msg := '由于并发操作,该单据可能已经删除或已经收费！';
      Raise Err_Item;
    End If;
  End Loop;
  Close c_Price;
  ------------------------------------------------------------------------------------------------------------------------

  --预交款相关结算
  --收费结算
  If 收费结算_In Is Not Null Then
    v_结算内容 := 收费结算_In || '||';
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
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补, 缴款组id, 结算序号, 交易流水号,
           交易说明, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, No_In, 1, Decode(病人id_In, 0, Null, 病人id_In), Null, v_结算摘要, v_结算方式, v_结算号码, v_Date,
           操作员编号_In, 操作员姓名_In, n_结算金额, 结帐id_In, Decode(v_结算内容, 收费结算_In || '||', 缴款_In, Null),
           Decode(v_结算内容, 收费结算_In || '||', 找补_In, Null), n_组id, 结算序号_In, 交易流水号_In, 交易说明_In, 3);
      End If;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 保险结算_In Is Not Null Then
    --各个保险结算
    v_结算内容 := 保险结算_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, No_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, v_Date, 操作员编号_In,
           操作员姓名_In, n_结算金额, 结帐id_In, n_组id, 结算序号_In, 3);
      End If;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 三方卡结算_In Is Not Null Then
    v_结算内容 := 三方卡结算_In || '||';
    While v_结算内容 Is Not Null Loop
      --卡类别Id|是否消费卡|结算金额|卡号|备注||...
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡   := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算摘要 := v_当前结算;
    
      If n_消费卡 = 1 Then
        Select 结算方式, 名称, Nvl(自制卡, 0)
        Into v_结算方式, v_卡名称, n_自制卡
        From 卡消费接口目录
        Where 编号 = n_卡类别id;
        If v_结算方式 Is Null Then
          v_Err_Msg := v_卡名称 || '未设置结算方式对照,请在消费卡中进行设置,结算失败！';
          Raise Err_Item;
        End If;
      Else
        Select 结算方式, 名称 Into v_结算方式, v_卡名称 From 医疗卡类别 Where ID = n_卡类别id;
        If v_结算方式 Is Null Then
          v_Err_Msg := v_卡名称 || '未设置结算方式对照,请在医疗卡管理中进行设置,结算失败！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_结算金额, 0) <> 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 卡类别id, 结算卡序号, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补, 缴款组id,
           结算序号, 卡号, 结算性质)
        Values
          (n_预交id, 3, No_In, 1, Decode(病人id_In, 0, Null, 病人id_In), Null, v_结算摘要, Decode(n_消费卡, 1, Null, n_卡类别id),
           Decode(n_消费卡, 0, Null, n_卡类别id), v_结算方式, v_结算号码, v_Date, 操作员编号_In, 操作员姓名_In, n_结算金额, 结帐id_In, Null, Null,
           n_组id, 结算序号_In, v_卡号, 3);
      
        --卡结算对照
        If n_消费卡 = 1 Then
          n_消费卡id := Null;
          If n_自制卡 = 1 Then
            Select ID
            Into n_消费卡id
            From 消费卡目录
            Where 接口编号 = n_卡类别id And 卡号 = v_卡号 And
                  序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = n_卡类别id And 卡号 = v_卡号);
          End If;
          Zl_病人卡结算记录_Insert(n_卡类别id, n_消费卡id, v_结算方式, n_结算金额, v_卡号, Null, Null, v_结算摘要, 结帐id_In, n_预交id);
        End If;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --预交结算
  If Nvl(冲预交额_In, 0) <> 0 Then
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '不能确定病人病人ID,收费使用预交款结算失败！';
      Raise Err_Item;
    End If;
  
    v_预交金额 := 冲预交额_In;
    For r_Deposit In c_Deposit(病人id_In) Loop
    
      n_金额 := Case
                When r_Deposit.金额 - v_预交金额 < 0 Then
                 r_Deposit.金额
                Else
                 v_预交金额
              End;
      If r_Deposit.Id <> 0 Then
        --第一次冲预交(82592,将第一次标上结帐ID,冲预交标记为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 3 Where ID = r_Deposit.Id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 校对标志, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, v_Date,
               操作员姓名_In, 操作员编号_In, n_金额, 结帐id_In, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号_In, 校对标志, 3
        From 病人预交记录
        Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_金额
      Where 病人id = 病人id_In And 性质 = 1 And 类型 = Nvl(r_Deposit.预交类别, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 预交余额, 类型, 性质) Values (病人id_In, -n_金额, Nvl(r_Deposit.预交类别, 2), 1);
        n_返回值 := -冲预交额_In;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --检查是否已经处理完
      If r_Deposit.金额 < v_预交金额 Then
        v_预交金额 := v_预交金额 - r_Deposit.金额;
      Else
        v_预交金额 := 0;
      End If;
      If v_预交金额 = 0 Then
        Exit;
      End If;
    End Loop;
    --检查金额是否足够
    If v_预交金额 > 0 Then
      v_Err_Msg := '病人的当前预交余额不足金额 ' || LTrim(To_Char(冲预交额_In, '9999999990.00')) || ' ！';
      Raise Err_Item;
    End If;
  
  End If;

  --相关汇总表的处理

  --汇总"人员缴款余额"
  --收费结算
  n_返回值 := 0;
  If 收费结算_In Is Not Null Then
    v_结算内容 := 收费结算_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(n_结算金额, 0)
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
        Returning Nvl(余额, 0) + n_返回值 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, v_结算方式, 1, Nvl(n_结算金额, 0));
          n_返回值 := Nvl(n_返回值, 0) + Nvl(n_结算金额, 0);
        End If;
      
      End If;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --各个保险结算
  If 保险结算_In Is Not Null Then
    v_结算内容 := 保险结算_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(n_结算金额, 0)
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
        Returning Nvl(余额, 0) + n_返回值 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, v_结算方式, 1, Nvl(n_结算金额, 0));
          n_返回值 := Nvl(n_返回值, 0) + Nvl(n_结算金额, 0);
        End If;
      End If;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 三方卡结算_In Is Not Null Then
    v_结算内容 := 三方卡结算_In || '||';
    While v_结算内容 Is Not Null Loop
      --卡类别Id|是否消费卡|结算金额|卡号|备注||...
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡   := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := Substr(v_当前结算, 1, Instr(v_当前结算, '|') + 1);
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算摘要 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    
      If n_消费卡 = 1 Then
        Select 结算方式, 名称, Nvl(自制卡, 0)
        Into v_结算方式, v_卡名称, n_自制卡
        From 卡消费接口目录
        Where 编号 = n_卡类别id;
        If v_结算方式 Is Null Then
          v_Err_Msg := v_卡名称 || '未设置结算方式对照,请在消费卡中进行设置,结算失败！';
          Raise Err_Item;
        End If;
      Else
        Select 结算方式, 名称 Into v_结算方式, v_卡名称 From 医疗卡类别 Where ID = n_卡类别id;
        If v_结算方式 Is Null Then
          v_Err_Msg := v_卡名称 || '未设置结算方式对照,请在医疗卡管理中进行设置,结算失败！';
          Raise Err_Item;
        End If;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(n_结算金额, 0)
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
        Returning Nvl(余额, 0) + n_返回值 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, v_结算方式, 1, Nvl(n_结算金额, 0));
          n_返回值 := Nvl(n_返回值, 0) + Nvl(n_结算金额, 0);
        End If;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If Nvl(n_返回值, 0) = 0 Then
    Delete From 人员缴款余额 Where 性质 = 1 And 收款员 = 操作员姓名_In And Nvl(余额, 0) = 0;
  End If;

  --药品部分非费用信息的修改
  --药品未发记录(如果已发药则修改不到),分离发药时无库房ID
  --可能存在材料和药品库房相同，但材料无发药窗口
  Update 未发药品记录
  Set 病人id = Decode(病人id_In, 0, Null, 病人id_In), 姓名 = 姓名_In, 对方部门id = 开单部门id_In, 已收费 = 1, 填制日期 = v_Date
  Where 单据 = 24 And NO = No_In And
        Nvl(库房id, 0) In (Select Distinct Nvl(执行部门id, 0)
                         From 门诊费用记录
                         Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 = '4');

  Update 未发药品记录
  Set 病人id = Decode(病人id_In, 0, Null, 病人id_In), 姓名 = 姓名_In, 对方部门id = 开单部门id_In, 已收费 = 1, 填制日期 = v_Date
  Where 单据 = 8 And NO = No_In And
        Nvl(库房id, 0) In (Select Distinct Nvl(执行部门id, 0)
                         From 门诊费用记录
                         Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));

  --药品收发记录(可能已经发药或取消发药,所有记录更改)
  Update 药品收发记录
  Set 对方部门id = 开单部门id_In, 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date)
  Where 单据 = 24 And NO = No_In And
        费用id + 0 In (Select ID From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 = '4');

  -------------------------------------------------------------------------------------------
  --处理备货卫材
  n_Count := Null;
  Begin
    Select Count(*), Max(a.No)
    Into n_Count, v_出库no
    From 药品收发记录 A, 门诊费用记录 B
    Where a.费用id = b.Id And b.收费类别 = '4' And b.记录性质 = 1 And b.记录状态 = 1 And
          Instr(',8,9,10,21,24,25,26,', ',' || a.单据 || ',') > 0 And b.No = No_In And Rownum <= 1;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(n_Count, 0) > 0 Then
    If Nvl(病人科室id_In, 0) <> 0 Then
      Select 名称 Into v_部门名称 From 部门表 Where ID = 病人科室id_In;
    End If;
    v_Err_Msg := LPad(' ', 4);
    v_Err_Msg := Substr('病人姓名:' || 姓名_In || v_Err_Msg || '性别:' || 性别_In || v_Err_Msg || '年龄' || 年龄_In || v_Err_Msg ||
                        '门诊号:' || Nvl(v_标识号, '') || v_Err_Msg || '病人科室:' || v_部门名称, 1, 100);
  
    Update 药品收发记录
    Set 对方部门id = 开单部门id_In, 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date), 摘要 = v_Err_Msg
    Where 单据 = 21 And NO = v_出库no And
          费用id + 0 In (Select ID From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 = '4');
  End If;

  Update 药品收发记录
  Set 对方部门id = 开单部门id_In, 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date)
  Where 单据 = 8 And NO = No_In And
        费用id + 0 In
        (Select ID From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));
  If Not 发药窗口_In Is Null Then
    --更新发药窗口
    If Nvl(简单收费_In, 0) <> 0 Then
      Update 门诊费用记录
      Set 发药窗口 = 发药窗口_In
      Where NO = No_In And 记录性质 = 1 And 记录状态 = 1 And 收费类别 = 'Z';
    Else
      For v_窗口 In (Select To_Number(C1) As C1, C2 From Table(f_Str2list2(发药窗口_In, ';', '|'))) Loop
        Update 门诊费用记录
        Set 发药窗口 = Nvl(v_窗口.C2, 发药窗口)
        Where NO = No_In And 记录性质 = 1 And 记录状态 = 1 And 执行部门id = Nvl(v_窗口.C1, 执行部门id) And 收费类别 In ('5', '6', '7');
      
        Update 药品收发记录
        Set 发药窗口 = Nvl(v_窗口.C2, 发药窗口)
        Where 单据 = 8 And NO = No_In And 库房id = Nvl(v_窗口.C1, 库房id) And
              费用id + 0 In (Select ID
                           From 门诊费用记录
                           Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));
      
        Update 未发药品记录
        Set 发药窗口 = Nvl(v_窗口.C2, 发药窗口)
        Where 单据 = 8 And NO = No_In And 库房id = Nvl(v_窗口.C1, 库房id) And
              Nvl(库房id, 0) In (Select Distinct Nvl(执行部门id, 0)
                               From 门诊费用记录
                               Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));
      End Loop;
    End If;
  End If;

  --更新部份病人信息
  If 病人id_In Is Not Null Then
    If 付款方式_In Is Not Null And 病人来源_In = 1 Then
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In;
    End If;
    --通过划价单收费时不允许改费别,因为费用不允许变
    Update 病人信息
    Set 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄), 姓名 = Decode(姓名, '新病人', 姓名_In, 姓名), 医疗付款方式 = Nvl(v_付款方式, 医疗付款方式)
    Where 病人id = 病人id_In;
    Select Zl_To_Number(Nvl(zl_GetSysParameter('自动产生姓名', '1111'), '0')) Into n_新病人模式 From Dual;
    If n_新病人模式 = 1 Then
    
      Update 病人挂号记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
    
      Update 门诊费用记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 付款方式 = 付款方式_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
    End If;
  End If;
  --医嘱处理
  --场合_In    Integer:=0, --0:门诊;1-住院
  --性质_In    Integer:=1, --1-收费单;2-记帐单
  --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  --No_In      门诊费用记录.No%Type,
  --医嘱ids_In Varchar2 := Null
  Zl_医嘱发送_计费状态_Update(0, 1, 1, No_In);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_划价收费记录_Insert;
/

--93623:刘兴洪,2016-02-25,更改医嘱发送的计费状态
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
  n_备货卫材 Number;

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
            Update 病人余额
            Set 费用余额 = Nvl(费用余额, 0) - n_实收金额
            Where 病人id = r_Bill.病人id And 性质 = 1 And 类型 = 1;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 性质, 类型, 费用余额, 预交余额)
              Values
                (r_Bill.病人id, 1, 1, -1 * n_实收金额, 0);
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
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
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
  If 序号_In Is Null And n_医嘱id Is Not Null Then
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where NO = No_In And 记录性质 = 2 And 医嘱序号 + 0 = n_医嘱id
                  Group By 记录状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Nvl(Sum(数量), 0) <> 0);
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = n_医嘱id And 记录性质 = 2 And NO = No_In;
    End If;
  End If;
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

--93623:刘兴洪,2016-02-25,更改医嘱发送的计费状态
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
  --场合_In    Integer:=0, --0:门诊;1-住院
  --性质_In    Integer:=1, --1-收费单;2-记帐单
  --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  --No_In      门诊费用记录.No%Type,
  --医嘱ids_In Varchar2 := Null
  Zl_医嘱发送_计费状态_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_Delete;
/

--93623:刘兴洪,2016-02-25,更改医嘱发送的计费状态
Create Or Replace Procedure Zl_住院记帐记录_Delete
(
  No_In           住院费用记录.No%Type,
  序号_In         Varchar2,
  操作员编号_In   住院费用记录.操作员编号%Type,
  操作员姓名_In   住院费用记录.操作员姓名%Type,
  记录性质_In     住院费用记录.记录性质%Type := 2,
  操作状态_In     Number := 0,
  输液配药检查_In Number := 1,
  登记时间_In     住院费用记录.登记时间%Type := Sysdate
) As
  --功能：冲销一张住院记帐单据中指定序号行
  --序号：格式如"1,3,5,7,8",或"1:2:33456,3:2,5:2,7:2,8:2",冒号前面的数字表示行号,中间的数字表示退的数量,后面的数字表示配药记录的ID,目前仅在销帐审核时才传入
  --      为空表示冲销所有可冲销行
  --记录性质:    2-人工记帐单,3-自动记帐单
  --输液配药检查:    0-医嘱调用，不检查药品是否进入输液配药中心；1-非医嘱调用，检查药品是否进入配药中心
  --该光标用于销帐指定费用行
  --操作状态_In:0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程)
  --该游标为要退费单据的所有原始记录
  Cursor c_Bill Is
    Select ID, 价格父号, 序号, 执行状态, 记录性质, 收费类别, 医嘱序号, 收费细目id, 病人id, 主页id, 收入项目id, 开单部门id, 病人科室id, 执行部门id, 病人病区id, 付数, 数次
    From 住院费用记录
    Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 门诊标志 = 2
    Order By 收费细目id, 序号;

  --该游标用于处理药品库存可用数量
  --不要管费用的执行状态,因为先于此步处理
  Cursor c_Stock(v_序号_In Varchar2) Is
    Select ID, 单据, NO, 库房id, 药品id, 批次, 发药方式, 付数, 实际数量, 灭菌效期, 效期, 产地, 批号, 填制日期, 费用id, 商品条码, 内部条码
    From 药品收发记录
    Where NO = No_In And 单据 In (9, 10, 25, 26) And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 住院费用记录
                   Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And
                         门诊标志 = 2 And (Instr(',' || v_序号_In || ',', ',' || 序号 || ',') > 0 Or v_序号_In Is Null))
    Order By 药品id, 填制日期 Desc;

  r_Stock c_Stock%RowType;
  --该游标用于处理费用记录序号
  Cursor c_Serial Is
    Select 序号, 价格父号
    From 住院费用记录
    Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3)
    Order By 序号;

  Cursor Cr_药品 Is
    Select ID, 单据, NO, 库房id, 药品id, 批次, 发药方式, 0 As 数量, 灭菌效期, 效期, 产地, 批号, 填制日期, 费用id
    From 药品收发记录
    Where Rownum <= 1;
  v_药品 Cr_药品%RowType;

  v_医嘱id 病人医嘱记录.Id%Type;
  n_划价   Number;
  v_父号   住院费用记录.价格父号%Type;
  v_序号   Varchar2(2000);
  v_Tmp    Varchar2(4000);

  v_医嘱ids    Varchar2(4000);
  l_药品收发   t_Numlist := t_Numlist();
  l_划价       t_Numlist := t_Numlist();
  l_费用id     t_Numlist := t_Numlist();
  n_付数       Number;
  n_虚拟库房id 药品收发记录.库房id%Type;
  n_其他出库id 药品收发记录.Id%Type;
  n_库房id     药品收发记录.库房id%Type;
  n_返回值     Number;
  --部分退费计算变量
  v_剩余数量 Number;
  v_剩余应收 Number;
  v_剩余实收 Number;
  v_剩余统筹 Number;

  v_准退数量 Number;
  v_退费次数 Number;
  v_应收金额 Number;
  v_实收金额 Number;
  v_统筹金额 Number;
  n_Temp     Number;
  n_部分销帐 Number;
  v_Dec      Number;
  n_Count    Number;
  v_Curdate  Date;
  Err_Item Exception;
  v_Err_Msg        Varchar2(255);
  n_备货卫材       Number;
  n_病人id         病案主页.病人id%Type;
  n_主页id         病案主页.主页id%Type;
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);
  v_配药id         Varchar2(4000);
  Type Ty_药品 Is Ref Cursor;
  c_药品 Ty_药品; --游标变量

Begin
  --销帐审核时,非药品会传入行号的销帐数量
  If Not 序号_In Is Null Then
    If Instr(序号_In, ':') > 0 Then
      v_Tmp := 序号_In || ',';
      While Not v_Tmp Is Null Loop
        v_序号 := v_序号 || ',' || Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
        If Instr(Substr(v_Tmp, Instr(v_Tmp, ':') + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':') - 1), ':') > 0 Then
          v_配药id := v_配药id || ',' ||
                    Substr(v_Tmp, Instr(v_Tmp, ':', 1, 2) + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':', 1, 2) - 1);
        End If;
        v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End Loop;
      v_序号 := Substr(v_序号, 2);
      If v_配药id Is Not Null Then
        v_配药id := Substr(v_配药id, 2);
      End If;
    Else
      v_序号 := 序号_In;
    End If;
  End If;

  --是否已经全部完全执行(只是整张单据的检查)
  Select Nvl(Count(*), 0), Nvl(Max(病人id), 0), Nvl(Max(主页id), 0)
  Into n_Count, n_病人id, n_主页id
  From 住院费用记录
  Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1 And 门诊标志 = 2;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
  n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
  If n_病人审核方式 = 1 Or n_未入科禁止记账 = 1 Then
  
    Begin
      Select 审核标志, 状态 Into n_审核标志, n_住院状态 From 病案主页 Where 病人id = n_病人id And 主页id = n_主页id;
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

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 住院费用记录
                Where NO = No_In And 记录性质 = 记录性质_In And 门诊标志 = 2 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 住院费用记录
                       Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！';
    Raise Err_Item;
  End If;

  --医嘱费用：检查正在执行的医嘱(注意已执行的情况在下面检查,因为不传 序号_IN 这种情况费用界面已限制)
  If Nvl(操作状态_In, 0) <> 1 Then
    --走销帐申请流程的，不检查医保执行状态
    Select Nvl(Count(*), 0)
    Into n_Count
    From 病人医嘱发送
    Where 执行状态 = 3 And (NO, 记录性质, 医嘱id) In
          (Select NO, 记录性质, 医嘱序号
                        From 住院费用记录
                        Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 医嘱序号 Is Not Null And
                              (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null));
    If n_Count > 0 Then
      v_Err_Msg := '要销帐的费用中存在对应的医嘱正在执行的情况，不能销帐！';
      Raise Err_Item;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --先打开药品对应数据集,以确保当前条件下有数据,为了处理并发判断
  --不能在游标条件中取消"审核人 is Null"条件，因为多次退药可能部份又已发
  Open c_Stock(v_序号);

  --公用变量
  Select 登记时间_In Into v_Curdate From Dual;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;

  For c_编目病案 In (Select a.姓名
                 From 病人信息 A, 病案主页 B
                 Where a.病人id = b.病人id And b.编目日期 Is Not Null And
                       (b.病人id, b.主页id) In
                       (Select Distinct 病人id, 主页id
                        From 住院费用记录
                        Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 门诊标志 = 2)) Loop
    v_Err_Msg := '病人『' || c_编目病案.姓名 || '』 已经被病案编目,不能被销帐！';
    Raise Err_Item;
  End Loop;
  v_医嘱ids := Null;
  --循环处理每行费用(收入项目行)
  For r_Bill In c_Bill Loop
    --检查已经存在病案编目的,则不能进行销帐处理
    If Instr(',' || v_序号 || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or v_序号 Is Null Then
      Select Decode(记录状态, 0, 1, 0) Into n_划价 From 住院费用记录 Where ID = r_Bill.Id;
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        --求剩余数量,剩余应收,剩余实收
        Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
        Into v_剩余数量, v_剩余应收, v_剩余实收, v_剩余统筹
        From 住院费用记录
        Where NO = No_In And 记录性质 = 记录性质_In And 序号 = r_Bill.序号;
        n_部分销帐 := 0;
        If v_剩余数量 = 0 Then
          If v_序号 Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部销帐！';
            Raise Err_Item;
          End If;
          --情况：未限定行号,原始单据中的该笔已经全部销帐(执行状态=0的一种可能)
        Else
        
          If Instr(序号_In, ':') > 0 Then
            v_Tmp := ',' || 序号_In;
            v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || r_Bill.序号 || ':') + Length(',' || r_Bill.序号 || ':'));
            v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
            If Instr(v_Tmp, ':') > 0 Then
              v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            End If;
            v_准退数量 := v_Tmp;
            n_部分销帐 := 1;
          End If;
        
          --准销数量(非药品项目为剩余数量,原始数量)
          If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Then
            If Instr(序号_In, ':') = 0 Or 序号_In Is Null Then
              v_准退数量 := v_剩余数量;
            End If;
          Else
            --医嘱超期收回时,卫材可能没有发放,但申请销帐的是部分数量,所以要以申请的为准
            If Instr(序号_In, ':') = 0 Or 序号_In Is Null Then
              Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0), Count(*)
              Into v_准退数量, n_Count
              From 药品收发记录
              Where NO = No_In And 单据 In (9, 10, 25, 26) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = r_Bill.Id;
            End If;
          
            --有剩余数量无准退数量的有两种情况：
            --1.不跟踪在用的卫材无对应的收发记录,这时使用剩余数量
            --2.并发操作,此时已发药或发料
            If v_准退数量 = 0 Then
              If r_Bill.收费类别 = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发料,须退料后再退费！';
                  Raise Err_Item;
                Else
                  v_准退数量 := v_剩余数量;
                End If;
              Else
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发药,须退药后再退费！';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --处理住院费用记录
          If Nvl(n_划价, 0) = 0 Then
            --划价时,直接更改数量,所以不须查划冲销次数
            --该笔项目第几次销帐
            Select Nvl(Max(Abs(执行状态)), 0) + 1
            Into v_退费次数
            From 住院费用记录
            Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 = 2 And 序号 = r_Bill.序号 And 门诊标志 = 2;
          End If;
        
          --金额=剩余金额*(准退数/剩余数)
          v_应收金额 := Round(v_剩余应收 * (v_准退数量 / v_剩余数量), v_Dec);
          v_实收金额 := Round(v_剩余实收 * (v_准退数量 / v_剩余数量), v_Dec);
          v_统筹金额 := Round(v_剩余统筹 * (v_准退数量 / v_剩余数量), v_Dec);
          If Nvl(n_划价, 0) = 1 Then
            If Nvl(n_部分销帐, 0) = 0 Then
              l_划价.Extend;
              l_划价(l_划价.Count) := r_Bill.Id;
              n_返回值 := 0;
            Else
              --更新数量
              --划价的,先将相关的数据处理在内部表集中
              n_付数 := 0;
              If r_Bill.付数 > 1 Then
                --如果是中药,超期回收肯定是回收的付数,而不是次数.因此,需要检查准退数量是否可以整 除
                If Trunc(v_准退数量 / r_Bill.数次) <> (v_准退数量 / r_Bill.数次) Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用为中药,请按付数进行退费！';
                  Raise Err_Item;
                End If;
                n_付数 := Trunc(v_准退数量 / r_Bill.数次);
                If Nvl(r_Bill.付数, 0) - n_付数 < 0 Then
                  v_准退数量 := r_Bill.数次;
                Else
                  v_准退数量 := 0;
                End If;
              End If;
              Update 住院费用记录
              Set 付数 = 付数 - n_付数, 数次 = 数次 - v_准退数量, 应收金额 = Nvl(应收金额, 0) - v_应收金额, 实收金额 = Nvl(实收金额, 0) - v_实收金额,
                  登记时间 = v_Curdate, 统筹金额 = Nvl(统筹金额, 0) - v_统筹金额
              Where ID = r_Bill.Id
              Returning Nvl(数次, 0) * Nvl(付数, 0) Into n_返回值;
            End If;
            If Nvl(n_返回值, 0) <= 0 Then
              l_划价.Extend;
              l_划价(l_划价.Count) := r_Bill.Id;
            End If;
            If r_Bill.医嘱序号 Is Not Null Then
              If Instr(',' || Nvl(v_医嘱ids, '') || ',', ',' || r_Bill.医嘱序号 || ',') = 0 Then
                v_医嘱ids := Nvl(v_医嘱ids, '') || ',' || r_Bill.医嘱序号;
              End If;
              --记录病人医嘱附费对应的医嘱ID(不是主费用)
              If v_医嘱id Is Null Then
                v_医嘱id := r_Bill.医嘱序号;
              End If;
            End If;
          
          End If;
        
          If Nvl(n_划价, 0) = 0 Then
            --划价时,直接更改数量,所以不须查划冲销次数
            --插入退费记录
            Insert Into 住院费用记录
              (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 主页id, 病人id, 医嘱序号, 门诊标志, 多病人单, 婴儿费, 姓名, 性别, 年龄, 标识号, 床号, 费别, 病人病区id,
               病人科室id, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人,
               执行部门id, 划价人, 执行人, 执行状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 保险编码, 记帐单id, 摘要, 费用类型, 是否急诊,
               结论, 医疗小组id)
              Select 病人费用记录_Id.Nextval, NO, 记录性质, 2, 序号, 从属父号, 价格父号, 主页id, 病人id, 医嘱序号, 门诊标志, 多病人单, 婴儿费, 姓名, 性别, 年龄, 标识号,
                     床号, 费别, 病人病区id, 病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(v_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                     Decode(Sign(v_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * v_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用,
                     标准单价, -1 * v_应收金额, -1 * v_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * v_退费次数, 执行时间, 操作员编号_In,
                     操作员姓名_In, 发生时间, v_Curdate, 保险项目否, 保险大类id, -1 * v_统筹金额, 保险编码, 记帐单id, 摘要, 费用类型, 是否急诊, 结论, 医疗小组id
              From 住院费用记录
              Where ID = r_Bill.Id;
          
            --记录病人医嘱附费对应的医嘱ID(不是主费用)
            If v_医嘱id Is Null And r_Bill.医嘱序号 Is Not Null Then
              v_医嘱id := r_Bill.医嘱序号;
            End If;
          
            Update 病人审批项目
            Set 已用数量 = Nvl(已用数量, 0) - v_准退数量
            Where 病人id = r_Bill.病人id And 主页id = r_Bill.主页id And 项目id = r_Bill.收费细目id And Nvl(使用限量, 0) <> 0;
          
            --病人余额
            Update 病人余额
            Set 费用余额 = Nvl(费用余额, 0) - v_实收金额
            Where 病人id = r_Bill.病人id And 类型 = 2 And 性质 = 1;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 类型, 性质, 费用余额, 预交余额)
              Values
                (r_Bill.病人id, 2, 1, -1 * v_实收金额, 0);
            End If;
          
            --病人未结费用
            Update 病人未结费用
            Set 金额 = Nvl(金额, 0) - v_实收金额
            Where 病人id = r_Bill.病人id And Nvl(主页id, 0) = Nvl(r_Bill.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_Bill.病人病区id, 0) And
                  Nvl(病人科室id, 0) = Nvl(r_Bill.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Bill.开单部门id, 0) And
                  Nvl(执行部门id, 0) = Nvl(r_Bill.执行部门id, 0) And 收入项目id + 0 = r_Bill.收入项目id And 来源途径 + 0 = 2;
            If Sql%RowCount = 0 Then
              Insert Into 病人未结费用
                (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
              Values
                (r_Bill.病人id, r_Bill.主页id, r_Bill.病人病区id, r_Bill.病人科室id, r_Bill.开单部门id, r_Bill.执行部门id, r_Bill.收入项目id, 2,
                 -1 * v_实收金额);
            End If;
          
            --标记原费用记录
            --执行状态:全部退完(准退数=剩余数)标记为0,否则保持原状态
            If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Then
              --一般情况非药品和卫材的项目,不存在部分销帐的情况,只有销帐申请和销帐审核时,才会出现部分销帐,所以
              --执行状态只有两种:0.未执行;1已执行;
              --由于在销帐审核过程中将已执行强制改为了2部分执行,因此需要在此处改为1已执行.未执行的不变.
              Update 住院费用记录
              Set 记录状态 = 3, 执行状态 = Decode(Sign(v_准退数量 - v_剩余数量), 0, 0, Decode(执行状态, 2, 1, 执行状态))
              Where ID = r_Bill.Id;
            Else
              Update 住院费用记录
              Set 记录状态 = 3, 执行状态 = Decode(Sign(v_准退数量 - v_剩余数量), 0, 0, 执行状态)
              Where ID = r_Bill.Id;
            End If;
          End If;
        End If;
      Else
        If v_序号 Is Not Null Then
          v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能销帐！';
          Raise Err_Item;
        End If;
        --情况:没限定行号,原始单据中包括已经完全执行的
      End If;
    End If;
  End Loop;

  --不存在配药ID,检查该药品是否在输液配药中心
  If v_配药id Is Null And 输液配药检查_In = 1 Then
    For v_费用 In (Select ID
                 From 住院费用记录
                 Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = 2 And
                       (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null)) Loop
      Begin
        Select Count(1)
        Into n_Count
        From 输液配药内容 A, 药品收发记录 B
        Where a.收发id = b.Id And b.费用id = v_费用.Id And Instr(',8,9,10,21,24,25,26,', ',' || b.单据 || ',') > 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '存在已经进入输液配药中心的待销帐药品，无法完成销帐！';
        Raise Err_Item;
      End If;
    End Loop;
  End If;

  n_部分销帐 := 0;
  ---------------------------------------------------------------------------------
  --药品相关处理:主要是对销帐审核有效.(可以是部分)
  For v_费用 In (Select ID, 序号, 收费类别
               From 住院费用记录
               Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = 2 And
                     (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null)) Loop
    --根据费用ID来进行相关的处理
    v_准退数量 := 0;
    If Instr(序号_In, ':') > 0 Then
      v_Tmp := ',' || 序号_In;
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || v_费用.序号 || ':') + Length(',' || v_费用.序号 || ':'));
      v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
      If Instr(v_Tmp, ':') > 0 Then
        v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      End If;
      v_准退数量 := v_Tmp;
    End If;
    If v_准退数量 <> 0 Then
      n_部分销帐 := 1;
      n_Temp     := 0;
      --------------------------------------------------------------------------------------
      --检查是否备货记帐卫材,规则如下
      -- a.如果存在存在未审核的其他出库且部分销帐时,直接在原来的基础上更改其他出库数量
      -- b.如果存在存在未审核的其他出库且完全销帐时,直接删除
      -- c.库存处理:还原为虚拟库房的可用数量;发料部门不处理
      -- d.如果已经发了料,这个时间由于其他出库单已经审核,因此就按正常情况流转,库存恢复到发料部门中
      n_虚拟库房id := Null;
      n_其他出库id := Null;
      If v_费用.收费类别 = '4' Then
        Begin
          Select 1, 库房id, ID
          Into n_备货卫材, n_虚拟库房id, n_其他出库id
          From 药品收发记录
          Where 费用id = v_费用.Id And 审核日期 Is Null And 单据 = 21 And Rownum = 1;
        Exception
          When Others Then
            n_备货卫材 := 0;
        End;
      Else
        n_备货卫材 := 0;
      End If;
      --------------------------------------------------------------------------------------
      If v_配药id Is Not Null Then
        Open c_药品 For
          Select /*+ rule*/
           a.Id, a.单据, a.No, a.库房id, a.药品id, a.批次, a.发药方式,
           Decode(a.发药方式, Null, 1, -1, 0, 1) * Nvl(a.付数, 1) * Nvl(a.实际数量, 0) As 数量, a.灭菌效期, a.效期, a.产地, a.批号, a.填制日期,
           a.费用id
          From 药品收发记录 A, Table(f_Str2list(v_配药id)) B, 输液配药内容 C
          Where a.No = No_In And a.单据 In (9, 10, 25, 26) And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null And a.费用id = v_费用.Id And
                a.Id = c.收发id And c.记录id = b.Column_Value
          Order By 填制日期;
      Else
        Open c_药品 For
          Select ID, 单据, NO, 库房id, 药品id, 批次, 发药方式, Decode(发药方式, Null, 1, -1, 0, 1) * Nvl(付数, 1) * Nvl(实际数量, 0) As 数量,
                 灭菌效期, 效期, 产地, 批号, 填制日期, 费用id
          From 药品收发记录
          Where NO = No_In And 单据 In (9, 10, 25, 26) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = v_费用.Id
          Order By 填制日期;
      End If;
      Loop
        Fetch c_药品
          Into v_药品;
        Exit When c_药品%NotFound;
        n_Temp := v_药品.数量;
        If v_准退数量 >= n_Temp Then
          l_药品收发.Extend;
          l_药品收发(l_药品收发.Count) := v_药品.Id;
          If Nvl(n_其他出库id, 0) > 0 Then
            l_药品收发.Extend;
            l_药品收发(l_药品收发.Count) := n_其他出库id;
          End If;
          v_准退数量 := v_准退数量 - n_Temp;
        Else
          If v_费用.收费类别 = '7' Then
            --当前行的数量要大
            Update 药品收发记录
            Set 付数 = 1, 实际数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量,
                填写数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(填写数量, 0) - v_准退数量,
                成本金额 =
                 (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 成本价,
                零售金额 =
                 (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 零售价,
                差价 = Round((Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 零售价 -
                            (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 成本价, 5)
            Where ID = v_药品.Id;
          Else
            Update 药品收发记录
            Set 实际数量 = Nvl(实际数量, 0) - v_准退数量, 填写数量 = Nvl(填写数量, 0) - v_准退数量,
                成本金额 =
                 (Nvl(实际数量, 0) - v_准退数量) * 成本价,
                零售金额 =
                 (Nvl(实际数量, 0) - v_准退数量) * 零售价,
                差价 = Round((Nvl(实际数量, 0) - v_准退数量) * 零售价 - (Nvl(实际数量, 0) - v_准退数量) * 成本价, 5)
            Where ID = v_药品.Id;
          End If;
          --更新其他出库单
          If Nvl(n_其他出库id, 0) <> 0 Then
            If v_费用.收费类别 = '7' Then
              Update 药品收发记录
              Set 付数 = 1, 实际数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量,
                  填写数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量,
                  成本金额 =
                   (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 成本价,
                  零售金额 =
                   (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 零售价,
                  差价 = Round((Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 零售价 -
                              (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 成本价, 5)
              Where ID = Nvl(n_其他出库id, 0);
            Else
              Update 药品收发记录
              Set 实际数量 = Nvl(实际数量, 0) - v_准退数量, 填写数量 = Nvl(实际数量, 0) - v_准退数量,
                  成本金额 =
                   (Nvl(实际数量, 0) - v_准退数量) * 成本价,
                  零售金额 =
                   (Nvl(实际数量, 0) - v_准退数量) * 零售价,
                  差价 = Round((Nvl(实际数量, 0) - v_准退数量) * 零售价 - (Nvl(实际数量, 0) - v_准退数量) * 成本价, 5)
              Where ID = Nvl(n_其他出库id, 0);
            End If;
          End If;
          n_Temp     := v_准退数量;
          v_准退数量 := 0;
        End If;
        If Nvl(n_备货卫材, 0) = 1 Then
          n_库房id := n_虚拟库房id;
        Else
          n_库房id := v_药品.库房id;
        End If;
      
        If n_库房id Is Not Null Then
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_Temp
          Where 库房id = n_库房id And 药品id = v_药品.药品id And Nvl(批次, 0) = Nvl(v_药品.批次, 0) And 性质 = 1;
          If Sql%RowCount = 0 Then
            Insert Into 药品库存
              (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期)
            Values
              (n_库房id, v_药品.药品id, 1, v_药品.批次, v_药品.效期, n_Temp, v_药品.批号, v_药品.产地, v_药品.灭菌效期);
          End If;
          Delete 药品库存
          Where 库房id = n_库房id And 药品id = v_药品.药品id And Nvl(批次, 0) = Nvl(v_药品.批次, 0) And 性质 = 1 And Nvl(可用数量, 0) = 0 And
                Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
        End If;
      
        If Nvl(n_备货卫材, 0) = 1 Then
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_Temp
          Where 库房id = v_药品.库房id And 药品id = v_药品.药品id And Nvl(批次, 0) = Nvl(v_药品.批次, 0) And 性质 = 1;
          If Sql%RowCount = 0 Then
            Insert Into 药品库存
              (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期)
            Values
              (v_药品.库房id, v_药品.药品id, 1, v_药品.批次, v_药品.效期, n_Temp, v_药品.批号, v_药品.产地, v_药品.灭菌效期);
          End If;
        End If;
      
        If v_准退数量 = 0 Then
          Exit;
        End If;
      End Loop;
      --不跟踪卫材的,不检查:因为不跟噻的话,不会在药品收发记录中存在
      If Nvl(v_准退数量, 0) <> 0 And Not (v_费用.收费类别 = '4' And n_Temp = 0) Then
        --未分配完成,表示此药品可能已经执行.
        v_Err_Msg := '要销帐的费用中存在已发的药品或卫材，或已被其他人销帐；这可能是并发操作引起的。';
        Raise Err_Item;
      End If;
    End If;
  End Loop;

  If n_部分销帐 = 0 Then
    ------------------------------------------------------------------------------------------------------------------------
    --先处理备货材料
    For v_出库 In (Select ID, 单据, NO, 库房id, 药品id, 批次, 发药方式, 付数, 实际数量, 灭菌效期, 效期, 产地, 批号, 填制日期, 费用id, 商品条码, 内部条码
                 From 药品收发记录
                 Where 单据 = 21 And Mod(记录状态, 3) = 1 And 审核人 Is Null And
                       费用id In (Select ID
                                From 住院费用记录
                                Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 = '4' And 门诊标志 = 2 And
                                      (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null))
                 Order By 药品id, 填制日期 Desc) Loop
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
        Delete 药品库存
        Where 库房id = v_出库.库房id And 药品id = v_出库.药品id And Nvl(批次, 0) = Nvl(v_出库.批次, 0) And 性质 = 1 And Nvl(可用数量, 0) = 0 And
              Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
      End If;
      l_费用id.Extend;
      l_费用id(l_费用id.Count) := v_出库.费用id;
      l_药品收发.Extend;
      l_药品收发(l_药品收发.Count) := v_出库.Id;
    End Loop;
  
    --药品相关内容
    Fetch c_Stock
      Into r_Stock;
    While c_Stock%Found Loop
    
      --处理药品库存
      If r_Stock.库房id Is Not Null Then
      
        Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
        Into n_备货卫材
        From Table(l_费用id)
        Where Column_Value = r_Stock.费用id;
      
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0)
        Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期)
          Values
            (r_Stock.库房id, r_Stock.药品id, 1, r_Stock.批次, r_Stock.效期,
             Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0), r_Stock.批号, r_Stock.产地,
             r_Stock.灭菌效期);
        End If;
        Delete 药品库存
        Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1 And
              Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
      End If;
    
      --删除药品收发记录(加上并发操作检查:审核人 Is Null)
      --Delete From 药品收发记录 Where ID = r_Stock.ID And 审核人 Is Null;
    
      l_药品收发.Extend;
      l_药品收发(l_药品收发.Count) := r_Stock.Id;
      Fetch c_Stock
        Into r_Stock;
    End Loop;
    Close c_Stock;
  
    --删除药品收发记录
    Forall I In 1 .. l_药品收发.Count
      Delete From 药品收发记录 Where ID = l_药品收发(I) And 审核人 Is Null;
    If Sql%RowCount <> l_药品收发.Count And l_药品收发.Count <> 0 Then
      v_Err_Msg := '要销帐的费用中存在已发的药品或卫材，或已被其他人销帐；这可能是并发操作引起的。';
      Raise Err_Item;
    End If;
  Else
    --删除药品收发记录
    Forall I In 1 .. l_药品收发.Count
      Delete From 药品收发记录 Where ID = l_药品收发(I) And 审核人 Is Null;
  End If;
  --未发药品记录
  Delete From 未发药品记录 A
  Where NO = No_In And 单据 In (9, 10, 25, 26) And Not Exists
   (Select 1
         From 药品收发记录
         Where 单据 = a.单据 And Nvl(库房id, 0) = Nvl(a.库房id, 0) And NO = No_In And Mod(记录状态, 3) = 1 And 审核人 Is Null);

  ---------------------------------------------------------------------------------
  --如果是划价,直接删除费用记录(药品处理后)
  n_Count := l_划价.Count;
  --删除划价记录
  Forall I In 1 .. l_划价.Count
    Delete From 住院费用记录 Where ID = l_划价(I);

  --删除之后再统一调整序号
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.价格父号 Is Null Then
        v_父号 := n_Count;
      End If;
    
      Update 住院费用记录
      Set 序号 = n_Count, 价格父号 = Decode(价格父号, Null, Null, v_父号)
      Where NO = No_In And 记录性质 = 记录性质_In And 序号 = r_Serial.序号;
    
      Update 住院费用记录
      Set 从属父号 = n_Count
      Where NO = No_In And 记录性质 = 记录性质_In And 从属父号 = r_Serial.序号;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;
  --整张单据全部冲完时，删除病人医嘱附费
  If v_序号 Is Null And v_医嘱id Is Not Null Then
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 住院费用记录
                  Where NO = No_In And 记录性质 = 2 And 医嘱序号 + 0 = v_医嘱id
                  Group By 记录状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Nvl(Sum(数量), 0) <> 0);
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = v_医嘱id And 记录性质 = 2 And NO = No_In;
    End If;
  End If;

  If v_医嘱ids Is Not Null Then
    --医嘱处理
    --场合_In    Integer:=0, --0:门诊;1-住院
    --性质_In    Integer:=1, --1-收费单;2-记帐单
    --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2 := Null
    v_医嘱ids := Substr(v_医嘱ids, 2);
    Zl_医嘱发送_计费状态_Update(1, 2, 0, No_In, v_医嘱ids);
  Else
    Zl_医嘱发送_计费状态_Update(1, 2, 2, No_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_住院记帐记录_Delete;
/
 

--93606:冉俊明,2016-02-25,医保病人部分退费时重结费用的误差费的记录状态填为了2，应该是1
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
  n_记录状态 病人预交记录.记录状态%Type;

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
    n_结帐id   := 冲销id_In;
    n_记录状态 := 2;
    If Nvl(n_重结id, 0) <> 0 Then
      n_结帐id   := n_重结id;
      n_记录状态 := 1;
    End If;
  
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + Nvl(误差金额_In, 0)
    Where 结帐id = n_结帐id And 结算方式 = v_误差费;
    If Sql%NotFound Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, n_记录状态, r_Balancedata.病人id, Null, Null, v_误差费, r_Balancedata.收款时间,
         r_Balancedata.操作员编号, r_Balancedata.操作员姓名, 误差金额_In, n_结帐id, r_Balancedata.缴款组id, r_Balancedata.结算序号, 2, Null,
         Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3);
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

--89305:张德婷,2015-02-24,扫两次瓶签自动发药
CREATE OR REPLACE Procedure Zl_输液配药记录_发送
(
  配药id_In   In Varchar2, --ID串:ID1,ID2....
  操作人员_In In 输液配药记录.操作人员%Type := Null,
  操作时间_In In 输液配药记录.操作时间%Type := Null,
  操作说明_In In 输液配药状态.操作说明%Type := Null
) Is
  v_Tansid Varchar2(20);
  v_Tmp    Varchar2(4000);
  v_Error    Varchar2(255);
  n_操作状态 输液配药记录.操作状态%Type;
  Err_Custom Exception;
Begin
  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');

     Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;

      if n_操作状态>4 then
        v_Error := '该数据已被操作，不能进行发送操作！';
        Raise Err_Custom;
      end if;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;

    Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间,操作说明) Values (v_Tansid, 5, 操作人员_In, 操作时间_In,操作说明_In);
    Update 输液配药记录 Set 操作状态 = 5, 操作人员 = 操作人员_In, 操作时间 = 操作时间_In Where ID = v_Tansid;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_发送;
/


--93235:刘尔旋,2016-02-24,结帐推送消息
Create Or Replace Procedure Zl_病人结帐记录_Insert
(
  Id_In           病人结帐记录.Id%Type,
  单据号_In       病人结帐记录.No%Type,
  病人id_In       病人结帐记录.病人id%Type,
  收费时间_In     病人结帐记录.收费时间%Type,
  开始日期_In     病人结帐记录.开始日期%Type,
  结束日期_In     病人结帐记录.结束日期%Type,
  中途结帐_In     病人结帐记录.中途结帐%Type := 0,
  多病人结帐_In   Number := 0,
  最大结帐次数_In Number := 0,
  备注_In         病人结帐记录.备注%Type := Null,
  来源_In         Number := 1,
  原因_In         病人结帐记录.原因%Type := Null,
  结帐类型_In     病人结帐记录.结帐类型%Type := 2
  --功能：插入一条病人结帐记录
  --1.来源_In:1-门诊;2-住院
) As
  v_Temp     Varchar2(255);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_组id 财务缴款分组.Id%Type;
Begin
  --操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  n_组id     := Zl_Get组id(v_人员姓名);

  --病人结帐记录
  Insert Into 病人结帐记录
    (ID, NO, 实际票号, 记录状态, 中途结帐, 病人id, 开始日期, 结束日期, 收费时间, 操作员编号, 操作员姓名, 备注, 原因, 缴款组id, 结帐类型)
  Values
    (Id_In, 单据号_In, Null, 1, 中途结帐_In, Decode(多病人结帐_In, 1, Null, 病人id_In), 开始日期_In, 结束日期_In, 收费时间_In, v_人员编号, v_人员姓名,
     备注_In, 原因_In, n_组id, 结帐类型_In);
     
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 15, Id_In;
  Exception
    When Others Then
      Null;
  End;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结帐记录_Insert;
/

--93380:刘尔旋,2016-02-24,财务组多组长轧帐
Create Or Replace Procedure Zl_缴款成员组成_Insert
(
  组id_In     In 缴款成员组成.组id%Type,
  成员id_In   In 缴款成员组成.成员id%Type,
  操作类型_In Number := 0
) Is
  --操作类型_IN:0-新增缴款成员,1-新增缴款组长
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_Count Number(18);
Begin
  If 操作类型_In = 0 Then
    --检查是否已经存在于其他组
    Select Count(b.组名称)
    Into n_Count
    From 缴款成员组成 A, 财务缴款分组 B
    Where a.组id = b.Id And a.组id <> 组id_In And 成员id = 成员id_In And 删除日期 >= Sysdate;
    If n_Count <> 0 Then
      v_Err_Msg := '[ZLSOFT]该人员可能因并发的原因已经被分配到其他组,不能再分配到该组了,请重新读取该人员![ZLSOFT]';
      Raise Err_Item;
    End If;
    Insert Into 缴款成员组成 (组id, 成员id) Values (组id_In, 成员id_In);
  Else
    Select Count(b.组名称)
    Into n_Count
    From 缴款成员组成 A, 财务缴款分组 B
    Where a.组id = b.Id And a.组id <> 组id_In And 成员id = 成员id_In And 删除日期 >= Sysdate;
    If n_Count <> 0 Then
      v_Err_Msg := '[ZLSOFT]该人员可能因并发的原因已经被分配到其他组,不能再分配成为组长了,请重新读取该人员![ZLSOFT]';
      Raise Err_Item;
    End If;
    Insert Into 财务组组长构成 (组id, 组长id) Values (组id_In, 成员id_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_缴款成员组成_Insert;
/

--93380:刘尔旋,2016-02-24,财务组多组长轧帐
Create Or Replace Procedure Zl_缴款成员组成_Move
(
  成员id_In   Varchar2,
  原组id_In   In 缴款成员组成.组id%Type,
  新组id_In   In 缴款成员组成.组id%Type := -1,
  操作类型_In In Number := 0
) Is
  --操作类型_IN:0-处理缴款成员,1-处理缴款组长
  --成员ID_IN:多个成员时,用逗号分离
  --新组ID_IN:-1表示移除;
  n_Pos     Number;
  v_Temp    Varchar2(4000);
  n_成员id  缴款成员组成.成员id%Type;
  n_余额    人员缴款余额.余额%Type;
  v_姓名    人员表.姓名%Type;
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  l_成员id t_Numlist := t_Numlist();
Begin
  --循环处理
  v_Temp := 成员id_In;
  If 操作类型_In = 0 Then
    While v_Temp Is Not Null Loop
      n_Pos := Instr(v_Temp, ',');
      If n_Pos = 0 Then
        If Not v_Temp Is Null Then
          n_成员id := To_Number(v_Temp);
          v_Temp   := Null;
          Select Sum(Nvl(a.余额, 0)), Max(a.收款员)
          Into n_余额, v_姓名
          From 人员缴款余额 A, 人员表 B
          Where a.性质 = 1 And a.收款员 = b.姓名 And b.Id = n_成员id;
          If Nvl(n_余额, 0) <> 0 And 新组id_In > 0 Then
            v_Err_Msg := '[ZLSOFT]人员『' || v_姓名 || '』还存在暂存金,不能调动到其他组![ZLSOFT]';
            Raise Err_Item;
          End If;
          l_成员id.Extend;
          l_成员id(l_成员id.Count) := n_成员id;
        End If;
      Else
        --得到成员ID
        n_成员id := To_Number(Substr(v_Temp, 1, n_Pos - 1));
        v_Temp   := Substr(v_Temp, n_Pos + 1);
        Select Sum(Nvl(a.余额, 0)), Max(a.收款员)
        Into n_余额, v_姓名
        From 人员缴款余额 A, 人员表 B
        Where a.性质 = 1 And a.收款员 = b.姓名 And b.Id = n_成员id;
        If Nvl(n_余额, 0) <> 0 And 新组id_In > 0 Then
          v_Err_Msg := '[ZLSOFT]人员『' || v_姓名 || '』还存在暂存金,不能调动到其他组![ZLSOFT]';
          Raise Err_Item;
        End If;
        l_成员id.Extend;
        l_成员id(l_成员id.Count) := n_成员id;
      End If;
    End Loop;
  
    Forall I In 1 .. l_成员id.Count
      Delete 缴款成员组成 Where 组id = 原组id_In And 成员id = l_成员id(I);
    If 新组id_In > 0 Then
      Forall I In 1 .. l_成员id.Count
        Insert Into 缴款成员组成 (组id, 成员id) Values (新组id_In, l_成员id(I));
    End If;
  Else
    While v_Temp Is Not Null Loop
      n_Pos := Instr(v_Temp, ',');
      If n_Pos = 0 Then
        If Not v_Temp Is Null Then
          n_成员id := To_Number(v_Temp);
          v_Temp   := Null;
          l_成员id.Extend;
          l_成员id(l_成员id.Count) := n_成员id;
        End If;
      Else
        --得到成员ID
        n_成员id := To_Number(Substr(v_Temp, 1, n_Pos - 1));
        v_Temp   := Substr(v_Temp, n_Pos + 1);
        l_成员id.Extend;
        l_成员id(l_成员id.Count) := n_成员id;
      End If;
    End Loop;
  
    Forall I In 1 .. l_成员id.Count
      Delete 财务组组长构成 Where 组id = 原组id_In And 组长id = l_成员id(I);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_缴款成员组成_Move;
/

--93380:刘尔旋,2016-02-24,财务组多组长轧帐
Create Or Replace Procedure Zl_小组轧帐记录_Insert
(
  Id_In       In 人员收缴记录.Id%Type,
  单据号_In   In 人员收缴记录.No%Type,
  缴款组id_In In 人员收缴记录.缴款组id%Type,
  开始时间_In In 人员收缴记录.开始时间%Type,
  终止时间_In In 人员收缴记录.终止时间%Type,
  收款员_In   In 人员收缴记录.收款员%Type,
  收款员id_In In 人员表.Id%Type,
  收款时间_In In 人员收缴记录.小组收款时间%Type,
  轧帐信息_In In Varchar2,
  操作类型_In In Number := 0
) Is
  ----------------------------------------------------------------------------------------
  --功能:财务组轧帐记录写入
  --参数:
  --     轧帐信息_IN:记录ID1,记录ID2,...记录IDn
  --     操作类型_IN:0-非批量添加 1-批量添加第一条记录 2-批量添加中记录 3-批量添加最后一条记录
  ----------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(100);
  Err_Item Exception;
  n_Exists   Number(2);
  v_No       人员收缴记录.No%Type;
  v_收款员   人员收缴记录.收款员%Type;
  n_冲预交款 人员收缴记录.冲预交款%Type := 0;
  n_借入合计 人员收缴记录.借入合计%Type := 0;
  n_借出合计 人员收缴记录.借出合计%Type := 0;

Begin
  --保存前的并发检查  
  n_Exists := 0;
  Begin
    Select /*+ Rule*/
     a.No, a.收款员, 1
    Into v_No, v_收款员, n_Exists
    From 人员收缴记录 A, Table(f_Num2list(轧帐信息_In)) B
    Where a.Id = b.Column_Value And 小组轧账id Is Not Null And Rownum < 2;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_收款员 || '收款员的收款单号为[' || v_No || ']的记录已被轧帐，不允许再次轧帐！';
    Raise Err_Item;
  End If;
  Begin
    Select /*+ Rule*/
     a.No, a.收款员, 1
    Into v_No, v_收款员, n_Exists
    From 人员收缴记录 A, Table(f_Num2list(轧帐信息_In)) B
    Where a.Id = b.Column_Value And 财务收款时间 Is Not Null And Rownum < 2;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_收款员 || '收款员的收款单号为[' || v_No || ']的记录已被财务科收款，不允许小组进行轧帐！';
    Raise Err_Item;
  End If;
  Begin
    Select /*+ Rule*/
     a.No, a.收款员, 1
    Into v_No, v_收款员, n_Exists
    From 人员收缴记录 A, Table(f_Num2list(轧帐信息_In)) B
    Where a.Id = b.Column_Value And 作废时间 Is Not Null And Rownum < 2;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_收款员 || '收款员的收款单号为[' || v_No || ']的记录已被作废，不允许进行轧帐！';
    Raise Err_Item;
  End If;

  --汇总金额
  Select /*+ Rule*/
   Sum(a.冲预交款), Sum(a.借入合计), Sum(a.借出合计)
  Into n_冲预交款, n_借入合计, n_借出合计
  From 人员收缴记录 A, Table(f_Num2list(轧帐信息_In)) B
  Where a.Id = b.Column_Value
  Group By a.记录性质;

  If 操作类型_In = 0 Or 操作类型_In = 1 Then
    --插入轧帐记录
    Insert Into 人员收缴记录
      (ID, 记录性质, NO, 收款员, 缴款组id, 登记时间, 开始时间, 终止时间, 冲预交款, 借入合计, 借出合计, 小组轧账id, 登记人)
      Select Id_In, 3, 单据号_In, 收款员_In, 缴款组id_In, 收款时间_In, 开始时间_In, 终止时间_In, n_冲预交款, n_借入合计, n_借出合计, Id_In, 收款员_In
      From Dual;
    Update 财务缴款分组 Set 上次轧帐时间 = 终止时间_In Where ID = 缴款组id_In;
    Update 财务组组长构成 Set 上次轧帐时间 = 终止时间_In Where 组id = 缴款组id_In And 组长id = 收款员id_In;
  End If;
  --更新轧帐记录
  If 操作类型_In = 2 Or 操作类型_In = 3 Then
    Update 人员收缴记录
    Set 冲预交款 = 冲预交款 + n_冲预交款, 借入合计 = 借入合计 + n_借入合计, 借出合计 = 借出合计 + n_借出合计
    Where ID = Id_In;
  End If;

  Update 人员收缴记录
  Set 小组轧账id = Id_In
  Where 小组收款id In (Select Column_Value From Table(f_Num2list(轧帐信息_In)));

  If 操作类型_In = 0 Or 操作类型_In = 3 Then
    --更新明细
    Insert Into 人员收缴明细
      (收缴id, 结算方式, 金额)
      Select Id_In, a.结算方式, Sum(a.金额)
      From 人员收缴明细 A, 人员收缴记录 B
      Where a.收缴id = b.Id And b.记录性质 = 2 And b.小组轧账id = Id_In
      Group By Id_In, 结算方式;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_小组轧帐记录_Insert;
/

--93380:刘尔旋,2016-02-24,财务组多组长轧帐
CREATE OR REPLACE Procedure Zl_小组轧帐记录_Cancel
(
  Id_In       In 人员收缴记录.Id%Type,
  作废人_In   In 人员收缴记录.作废人%Type,
  作废人ID_IN In 人员表.ID%Type,
  作废时间_In In 人员收缴记录.作废时间%Type,
  缴款组id_In In 财务缴款分组.Id%Type
) Is
  ----------------------------------------------------------------------------------------
  --功能:财务组轧帐记录作废
  --参数:ID_IN:要作废记录的ID
  ----------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(100);
  Err_Item Exception;
  n_Exists   Number(2);
  n_Count    Number(18);
  v_No       人员收缴记录.No%Type;
  d_截止时间 人员收缴记录.终止时间%Type;
  v_收款员   人员收缴记录.收款员%Type;
Begin
  ---保存前并发检查
  n_Exists := 0;
  Begin
    Select NO, 收款员 Into v_No, v_收款员 From 人员收缴记录 Where ID = Id_In;
  Exception
    When Others Then
      n_Exists := 1;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := '记录未被找到，可能已被删除，无法进行轧帐作废操作！';
    Raise Err_Item;
  End If;
  Begin
    Select 1 Into n_Exists From 人员收缴记录 Where 财务收款时间 Is Not Null And ID = Id_In;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_收款员 || '收款员的轧帐单号为[' || v_No || ']的记录已被财务科收款，不允许作废！';
    Raise Err_Item;
  End If;
  Begin
    Select 1 Into n_Exists From 人员收缴记录 Where 作废时间 Is Not Null And ID = Id_In;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_收款员 || '收款员的轧帐单号为[' || v_No || ']的记录已被作废，不允许再次作废！';
    Raise Err_Item;
  End If;

  --检查是否最后一次轧帐记录
  Select Count(*)
  Into n_Count
  From 人员收缴记录
  Where 登记时间 > (Select 登记时间 From 人员收缴记录 Where ID = Id_In) And 记录性质 = 3 And ID + 0 <> Id_In And Rownum < 2 And
        收款员 || '' = v_收款员 And 作废时间 Is Null;

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
        收款员 || '' = v_收款员 And 记录性质 = 3;
  If d_截止时间 Is Null Then
    --取组里最小一次收款记录的登记时间
    Select Min(登记时间)
    Into d_截止时间
    From 人员收缴记录
    Where 记录性质 = 2 And 作废时间 Is Null And 财务收款时间 Is Null And 缴款组id = 缴款组id_In;
  End If;
  Update 财务缴款分组 Set 上次轧帐时间 = d_截止时间 Where ID = 缴款组id_In;
  Update 财务组组长构成 Set 上次轧帐时间 = d_截止时间 Where 组ID = 缴款组id_In And 组长ID=作废人ID_IN;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101,  '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_小组轧帐记录_Cancel;
/

--93380:刘尔旋,2016-02-24,财务组多组长轧帐
Create Or Replace Procedure Zl_收费员轧帐记录_Insert
(
  Id_In         In 人员收缴记录.Id%Type,
  No_In         In 人员收缴记录.No%Type,
  收款员_In     In 人员收缴记录.收款员%Type,
  收款部门id_In In 人员收缴记录.收款部门id%Type,
  组长id_In     In 人员表.Id%Type,
  冲预交款_In   In 人员收缴记录.冲预交款%Type,
  借入合计_In   In 人员收缴记录.借入合计%Type,
  借出合计_In   In 人员收缴记录.借出合计%Type,
  摘要_In       In 人员收缴记录.摘要%Type,
  开始时间_In   In 人员收缴记录.开始时间%Type,
  终止时间_In   In 人员收缴记录.终止时间%Type,
  登记人_In     In 人员收缴记录.登记人%Type,
  登记时间_In   In 人员收缴记录.登记时间%Type,
  收缴标志_In   In 人员收缴记录.收缴标志%Type,
  收费对照_In   In Varchar2,
  操作类别_In   In Integer := 0,
  暂存金_In     In 人员暂存记录.金额%Type := 0,
  类别_In       In 人员收缴记录.类别%Type := 0
) Is
  --------------------------------------------------------------------------------------------
  --功能:收费员轧帐记录写入
  --参数:记录性质_IN:
  --     收费对照_IN:性质1,记录ID1|性质2,记录ID2|...|性质n,记录IDn
  --     操作类别_In:0-保存轧帐记录和对照;1-只保存对照
  --------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  Err_Item Exception;

  n_组id       人员收缴记录.缴款组id%Type;
  n_Count      Number(18);
  n_暂存id     人员暂存记录.Id%Type;
  v_暂存单号   人员暂存记录.No%Type;
  v_小组收款人 人员收缴记录.小组收款人%Type;
  v_现金       结算方式.名称%Type;

Begin
  --并发检查,是否当前时间范围内已经有轧帐记录
  If 操作类别_In = 0 Then
    Select Count(1)
    Into n_Count
    From 人员收缴记录
    Where 收款员 = 收款员_In And 开始时间 > 开始时间_In And 作废时间 Is Null And Nvl(类别, 0) = Nvl(类别_In, 0) And Rownum < 2 And 记录性质 = 1;
    If n_Count <> 0 Then
      v_Err_Msg := '收费员:"' || 收款员_In || '"在当前轧帐范围内已经被他人进行轧帐,不能再进行轧帐！';
      Raise Err_Item;
    End If;
  End If;

  --检查是否已存在轧帐明细
  For c_检查 In (Select /*+ rule*/
                a.性质, a.记录id, m.收款员
               From 人员收缴对照 A, Table(f_Str2list2(收费对照_In, '|', ',')) B, 人员收缴记录 M
               Where a.性质 = b.C1 And a.记录id = b.C2 And a.收缴id = m.Id And m.作废时间 Is Null And Rownum < 2) Loop
    --1-收费(含挂号),2-结帐,3-预交,4-借款;5-消费卡充值;6--消费卡面值;7-暂存金(本次增加)；８－收款或轧帐作废对照（本次增加）
    If c_检查.性质 = 1 Then
      Select Decode(记录性质, 1, '收费', 4, '挂号', '') || '单据为:' || NO || '的' || Decode(记录性质, 1, '收费', 4, '挂号', '') ||
              '记录已经被轧帐，不能再进行轧帐处理'
      Into v_Err_Msg
      From 门诊费用记录
      Where 结帐id = c_检查.记录id And Rownum < 2;
    End If;
  
    If c_检查.性质 = 2 Then
      Select '结帐单据为:' || NO || '的结帐记录已经被轧帐，不能再进行轧帐处理'
      Into v_Err_Msg
      From 病人结帐记录
      Where ID = c_检查.记录id And Rownum < 2;
    End If;
    If c_检查.性质 = 3 Then
      Select '预交单据为:' || NO || '的预交记录已经被轧帐，不能再进行轧帐处理'
      Into v_Err_Msg
      From 病人预交记录
      Where ID = c_检查.记录id And Rownum < 2;
    End If;
    If c_检查.性质 = 4 And 收款员_In = c_检查.收款员 Then
      Select '收费员在' || To_Char(借出时间, 'yyyy-mm-dd hh24:mi:ss') || '的' || Decode(借款人, Null, '借出款', '借入款') ||
              '已经被轧帐，不能再进行轧帐处理'
      Into v_Err_Msg
      From 人员借款记录
      Where ID = c_检查.记录id And Rownum < 2;
    End If;
  
    If c_检查.性质 = 5 Then
      Select '收费员在' || To_Char(充值时间, 'yyyy-mm-dd hh24:mi:ss') || '的消费卡充值记录已经被轧帐，不能再进行轧帐处理'
      Into v_Err_Msg
      From 消费卡充值记录
      Where ID = c_检查.记录id And Rownum < 2;
    End If;
    If c_检查.性质 = 6 Then
      Select '收费员在发卡号为:' || 卡号 || '的且发卡时间为:' || To_Char(发卡时间, 'yyyy-mm-dd hh24:mi:ss') || '发卡记录已经被轧帐，不能再进行轧帐处理'
      Into v_Err_Msg
      From 消费卡目录
      Where ID = c_检查.记录id And Rownum < 2;
    End If;
    If c_检查.性质 = 7 Then
      Select '暂存单号为:' || NO || '的暂存记录已经被轧帐，不能再进行轧帐处理'
      Into v_Err_Msg
      From 人员暂存记录
      Where ID = c_检查.记录id And Rownum < 2;
    End If;
    If c_检查.性质 = 9 Then
      Select '单号为:' || NO || '的补充结算记录已经被轧帐，不能再进行轧帐处理'
      Into v_Err_Msg
      From 费用补充记录
      Where 结算id = c_检查.记录id And Rownum < 2;
    End If;
    If v_Err_Msg Is Not Null Then
      Raise Err_Item;
    End If;
  End Loop;

  If 操作类别_In = 0 Then
    n_组id := Zl_Get组id(收款员_In);
    If 组长id_In Is Not Null Then
      Select 姓名 Into v_小组收款人 From 人员表 Where ID = 组长id_In;
    End If;
    Insert Into 人员收缴记录
      (ID, 记录性质, NO, 收款员, 收款部门id, 冲预交款, 借入合计, 借出合计, 摘要, 开始时间, 终止时间, 缴款组id, 登记人, 登记时间, 收缴标志, 类别, 小组收款人)
    Values
      (Id_In, 1, No_In, 收款员_In, 收款部门id_In, 冲预交款_In, 借入合计_In, 借出合计_In, 摘要_In, 开始时间_In, 终止时间_In, n_组id, 登记人_In, 登记时间_In,
       收缴标志_In, 类别_In, v_小组收款人);
  
    Update 人员缴款余额 Set 上次轧帐时间 = 终止时间_In Where 收款员 = 收款员_In;
  End If;
  --插入收费对照
  Insert Into 人员收缴对照
    (收缴id, 性质, 记录id)
    Select Id_In, C1, C2 From Table(f_Str2list2(收费对照_In, '|', ','));
  If 暂存金_In <> 0 Then
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1 And Rownum < 2;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
    Select 人员暂存记录_Id.Nextval Into n_暂存id From Dual;
    v_暂存单号 := Nextno(141);
    Insert Into 人员暂存记录
      (ID, 收缴id, 记录性质, NO, 结算方式, 金额, 收款员, 登记人, 登记时间)
    Values
      (n_暂存id, Id_In, 2, v_暂存单号, v_现金, -1 * 暂存金_In, 收款员_In, 登记人_In, 登记时间_In);
  
    Insert Into 人员收缴对照 (收缴id, 性质, 记录id) Values (Id_In, 7, n_暂存id);
    Select 人员暂存记录_Id.Nextval Into n_暂存id From Dual;
    v_暂存单号 := Nextno(141);
    Insert Into 人员暂存记录
      (ID, 收缴id, 记录性质, NO, 结算方式, 金额, 收款员, 登记人, 登记时间)
    Values
      (n_暂存id, Id_In, 2, v_暂存单号, v_现金, 暂存金_In, 收款员_In, 登记人_In, 登记时间_In + 1 / 24 / 60 / 60);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_收费员轧帐记录_Insert;
/

--93564:冉俊明,2016-02-24,费用金额为0时未成功插入结算方式为NULL的病人预交记录。
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
    If Nvl(n_Havenull, 0) = 0 Or Round(Nvl(r_Feedata.结算金额, 0), 6) <> Round(Nvl(n_结算金额, 0), 6) Then
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

--93333:王振涛,2016-02-24,取消参数控制发料
Create Or Replace Procedure Zl_检验标本记录_标本核收
(
  Id_In         In 检验标本记录.Id%Type,
  医嘱id_In     In 检验标本记录.医嘱id%Type,
  多个医嘱_In   In Varchar2, --用于更新多个医嘱的执行状态
  复盖标本id_In In 检验标本记录.Id%Type := 0, --补填时指向另一个标本时复盖指向的标本
  标本序号_In   In 检验标本记录.标本序号%Type,
  采样时间_In   In 检验标本记录.采样时间%Type,
  采样人_In     In 检验标本记录.采样人%Type,
  仪器id_In     In 检验标本记录.仪器id%Type,
  核收时间_In   In 检验标本记录.核收时间%Type,
  标本形态_In   In 检验标本记录.标本形态%Type,
  检验人_In     In 检验标本记录.检验人%Type := Null,
  检验时间_In   In 检验标本记录.检验时间%Type := Null,
  微生物标本_In In 检验标本记录.微生物标本%Type := Null,
  标本类别_In   In 检验标本记录.标本类别%Type := 0,
  检验备注_In   In 检验标本记录.检验备注%Type := Null,
  姓名_In       In 检验标本记录.姓名%Type := Null,
  性别_In       In 检验标本记录.性别%Type := Null,
  年龄_In       In 检验标本记录.年龄%Type := Null,
  No_In         In 检验标本记录.No%Type := Null,
  标本类型_In   In 检验标本记录.标本类型%Type := Null,
  申请科室id_In In 检验标本记录.申请科室id%Type := Null,
  申请人_In     In 检验标本记录.申请人%Type := Null,
  标识号_In     In 检验标本记录.标识号%Type := Null,
  床号_In       In 检验标本记录.床号%Type := Null,
  病人科室_In   In 检验标本记录.病人科室%Type := Null,
  检验项目_In   In 检验标本记录.检验项目%Type := Null,
  申请类型_In   In 检验标本记录.申请类型%Type := Null,
  病人id_In     In 检验标本记录.病人id%Type := Null,
  执行科室_In   In 检验标本记录.执行科室id%Type := Null,
  人员编号_In   In 人员表.编号%Type := Null,
  人员姓名_In   In 人员表.姓名%Type := Null
) Is

  Cursor v_Advice Is
    Select /*+ Rule */
    Distinct a.Id, a.开嘱时间, a.标本部位, f.样本条码, a.执行科室id, a.诊疗项目id, a.开嘱科室id, a.开嘱医生, a.病人id, a.病人来源, a.婴儿, a.紧急标志 As 紧急,
             b.门诊号, b.住院号, b.出生日期, a.挂号单, Decode(c.主页id, 0, Null, c.主页id) As 主页id, d.操作类型, f.接收人, f.接收时间
    From 病人医嘱记录 A, 病人医嘱发送 F, 病人信息 B, 病案主页 C, 诊疗项目目录 D
    Where a.相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist))) And a.Id = f.医嘱id And
          a.病人id = b.病人id And a.病人id = c.病人id(+) And a.主页id = c.主页id(+) And a.诊疗项目id = d.Id(+);

  Cursor v_Advice_1 Is
    Select /*+ Rule */
    Distinct b.No As 单据号
    From 病人医嘱记录 A, 病人医嘱发送 B
    Where a.Id = b.医嘱id And a.相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
    Union All
    Select /*+ Rule */
    Distinct b.No As 单据号
    From 病人医嘱记录 A, 病人医嘱发送 B
    Where a.Id = b.医嘱id And a.Id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)));

  Cursor v_Patient Is
    Select 病人id, 住院号, 门诊号, 出生日期 From 病人信息 Where 病人id = 病人id_In;

  --未审核的费用行(不包含药品)
  Cursor c_Verify(v_医嘱id In Number) Is
    Select /*+ Rule */
    Distinct a.记录性质, a.No, a.序号, a.医嘱序号, a.门诊标志
    From 住院费用记录 A, 病人医嘱发送 B,
         (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From 病人医嘱记录
           Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))) C
    Where a.收费类别 Not In ('5', '6', '7') And a.医嘱序号 = c.Id And a.记录状态 = 0 And 价格父号 Is Null And a.医嘱序号 = b.医嘱id And
          a.记录性质 = b.记录性质 And a.No = b.No And a.记帐费用 = 1
    Union All
    Select /*+ Rule */
    Distinct a.记录性质, a.No, a.序号, a.医嘱序号, a.门诊标志
    From 住院费用记录 A, 病人医嘱附费 B,
         (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From 病人医嘱记录
           Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))) C
    Where a.收费类别 Not In ('5', '6', '7') And a.医嘱序号 = c.Id And a.记录状态 = 0 And 价格父号 Is Null And a.医嘱序号 = b.医嘱id And
          a.记录性质 = b.记录性质 And a.No = b.No And a.记帐费用 = 1
    Union All
    Select /*+ Rule */
    Distinct a.记录性质, a.No, a.序号, a.医嘱序号, a.门诊标志
    From 门诊费用记录 A, 病人医嘱发送 B,
         (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From 病人医嘱记录
           Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))) C
    Where a.收费类别 Not In ('5', '6', '7') And a.医嘱序号 = c.Id And a.记录状态 = 0 And 价格父号 Is Null And a.医嘱序号 = b.医嘱id And
          a.记录性质 = b.记录性质 And a.No = b.No And a.记帐费用 = 1
    Union All
    Select /*+ Rule */
    Distinct a.记录性质, a.No, a.序号, a.医嘱序号, a.门诊标志
    From 门诊费用记录 A, 病人医嘱附费 B,
         (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From 病人医嘱记录
           Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))) C
    Where a.收费类别 Not In ('5', '6', '7') And a.医嘱序号 = c.Id And a.记录状态 = 0 And 价格父号 Is Null And a.医嘱序号 = b.医嘱id And
          a.记录性质 = b.记录性质 And a.No = b.No And a.记帐费用 = 1
    Order By 记录性质, NO, 序号;

  --查找当前标本的相关申请
  Cursor c_Samplequest(v_微生物 In Number) Is
    Select Distinct 医嘱id, 病人来源
    From (Select Decode(a.医嘱id, Null, b.医嘱id, a.医嘱id) As 医嘱id, b.病人来源
           From 检验项目分布 A, 检验标本记录 B
           Where Nvl(v_微生物, 0) = 0 And a.标本id = b.Id And b.医嘱id In (Select 医嘱id From 检验标本记录 Where ID = Id_In) And
                 a.医嘱id Is Not Null
           Union
           Select Decode(a.医嘱id, Null, b.医嘱id, a.医嘱id) As 医嘱id, b.病人来源
           From 检验项目分布 A, 检验标本记录 B
           Where 1 = v_微生物 And b.Id = a.标本id And b.Id = Id_In
           Union
           Select b.Id As 医嘱id, b.病人来源
           From 检验标本记录 A, 病人医嘱记录 B
           Where a.Id = Id_In And a.医嘱id In (b.Id, b.相关id));

  Cursor c_Stuff
  (
    v_No     Varchar2,
    v_主页id Number
  ) Is
    Select NO As 单据号, 单据, 库房id
    From 未发药品记录
    Where NO = v_No And 单据 In (24, 25, 26) And 库房id Is Not Null And Not Exists
     (Select 1 From Dual Where Zl_Getsysparameter(Decode(v_主页id, Null, 92, 63)) = '1') And Exists
     (Select a.序号
           From 住院费用记录 A, 材料特性 B
           Where a.记录性质 = 2 And a.No = v_No And a.收费细目id = b.材料id And b.跟踪在用 = 1
           Union All
           Select a.序号
           From 门诊费用记录 A, 材料特性 B
           Where a.记录性质 = 2 And a.No = v_No And a.收费细目id = b.材料id And b.跟踪在用 = 1)
    Order By 库房id;

  r_Advice   v_Advice%Rowtype;
  r_Advice_1 v_Advice_1%Rowtype;
  r_Patient  v_Patient%Rowtype;

  Err_Custom Exception;
  v_Error Varchar2(1000);
  v_Flag  Number(18);

  v_Temp      Varchar2(255);
  v_Seq       Number;
  v_Union     Number;
  v_Patientid Number;
  v_Itemid    Number;
  v_Count     Number;
  v_执行      Number;
  v_No        病人医嘱发送.No%Type;
  v_性质      病人医嘱发送.记录性质%Type;
  v_序号      Varchar2(1000);
  v_主页id    Number(18);
  v_门诊标志  住院费用记录.门诊标志%Type;
  n_Count     Number;
  v_姓名      病人医嘱记录.姓名%Type;
  v_性别      病人医嘱记录.性别%Type;
  v_年龄      病人医嘱记录.年龄%Type;
  v_病人来源  病人医嘱记录.病人来源%Type;
  v_婴儿      病人医嘱记录.婴儿%Type;
  v_婴儿姓名  病人医嘱记录.姓名%Type;
  v_婴儿性别  病人医嘱记录.性别%Type;
Begin

  If Nvl(复盖标本id_In, 0) > 0 Then
    Begin
      Select 姓名 Into v_Temp From 检验标本记录 Where ID = 复盖标本id_In And 姓名 Is Null;
    Exception
      When Others Then
        v_Error := '指定复盖的标本已被核收或已删除，请重新指定！';
        Raise Err_Custom;
    End;
  End If;

  If Nvl(医嘱id_In, 0) > 0 Then
    Select 姓名, 性别, 年龄, 病人来源, 婴儿
    Into v_姓名, v_性别, v_年龄, v_病人来源, v_婴儿
    From 病人医嘱记录
    Where ID = 医嘱id_In;
  
    If v_病人来源 <> 3 Then
      If Nvl(v_婴儿, 0) = 0 Then
        If v_姓名 <> 姓名_In Or v_性别 <> 性别_In  Then
          v_Error := '病人姓名、性别、年龄和医嘱不符不能保存，请检查或修改病人信息后再进行保存！';
          Raise Err_Custom;
        End If;
      Else
        Select b.婴儿姓名, b.婴儿性别
        Into v_婴儿姓名, v_婴儿性别
        From 病人医嘱记录 A, 病人新生儿记录 B
        Where a.病人id = b.病人id And a.主页id = b.主页id And a.婴儿 = b.序号 And
              a.相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist))) And Rownum = 1;
      
        If v_婴儿姓名 <> 姓名_In Or v_婴儿性别 <> 性别_In Then
          v_Error := '病人姓名、性别、年龄和医嘱不符不能保存，请检查或修改病人信息后再进行保存！';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  
    Select Count(ID) Into v_Flag From 检验标本记录 Where 医嘱id = 医嘱id_In And ID <> Id_In;
    If v_Flag > 0 Then
      Select Count(Distinct b.报告项目id)
      Into v_Flag
      From 病人医嘱记录 A, 检验报告项目 B
      Where a.诊疗项目id = b.诊疗项目id And a.相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)));
    
      Select Count(a.项目id)
      Into n_Count
      From 检验项目分布 A
      Where a.医嘱id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist))) And a.标本id <> Id_In;
      If (v_Flag - n_Count) <= 0 Then
        v_Error := '当前医嘱已被核收，不能重复核收！';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  If 医嘱id_In = 0 Then
    Open v_Patient;
    Fetch v_Patient
      Into r_Patient;
  
    If v_Patient%Found Then
      Zl_病人信息_锁定检查(r_Patient.病人id);
    End If;
  
    Update 检验标本记录
    Set 采样时间 = Decode(采样时间_In, Null, 采样时间, 采样时间_In), 采样人 = Decode(采样人_In, Null, 采样人, 采样人_In), 标本类型 = Nvl(标本类型_In, 标本类型),
        检验时间 = 检验时间_In, 姓名 = Decode(姓名_In, Null, 姓名, 姓名_In), 性别 = Decode(性别_In, Null, 性别, 性别_In),
        年龄 = Decode(年龄_In, Null, 年龄, 年龄_In), 年龄数字 = Decode(年龄_In, Null, Null, Zl_Val(年龄_In)),
        年龄单位 = Decode(年龄_In, Null, 年龄单位,
                       Decode(年龄_In, Null, Null, '成人', '成人', '婴儿', '婴儿',
                               Decode(Sign(Instr(年龄_In, '岁')), 1, '岁',
                                       Decode(Sign(Instr(年龄_In, '月')), 1, '月',
                                               Decode(Sign(Instr(年龄_In, '天')), 1, '天',
                                                       Decode(Sign(Instr(年龄_In, '小时')), 1, '小时', Null)))))),
        申请科室id = Decode(申请科室id_In, Null, 申请科室id, 申请科室id_In), 申请人 = Decode(申请人_In, Null, 申请人, 申请人_In),
        标本形态 = Decode(标本形态_In, Null, 标本形态, 标本形态_In), 标识号 = Decode(标识号_In, Null, 标识号, 标识号_In),
        床号 = Decode(床号_In, Null, 床号, 床号_In), 病人科室 = Decode(病人科室_In, Null, 病人科室, 病人科室_In),
        检验项目 = Decode(检验项目_In, Null, 检验项目, 检验项目_In), 病人id = Decode(病人id_In, Null, 病人id, 病人id_In),
        医嘱id = Decode(医嘱id_In, Null, 医嘱id, 0, 医嘱id, 医嘱id_In)
    Where ID = Id_In;
    If Sql%NotFound Then
      Insert Into 检验标本记录
        (ID, 医嘱id, 标本序号, 采样时间, 采样人, 标本类型, 核收人, 核收时间, 样本状态, 申请类型, 仪器id, 样本条码, 申请时间, 标本形态, 报告结果, 执行科室id, 检验人, 检验时间, 微生物标本,
         标本类别, 检验备注, 申请科室id, 申请人, 姓名, 性别, 年龄, 年龄数字, 年龄单位, 病人id, 病人来源, 婴儿, NO, 合并id, 标识号, 床号, 病人科室, 紧急, 门诊号, 住院号, 出生日期,
         挂号单, 主页id, 检验项目, 操作类型, 接收人, 接收时间)
      Values
        (Id_In, Decode(医嘱id_In, 0, Null, 医嘱id_In), 标本序号_In, 采样时间_In, 采样人_In, 标本类型_In, 人员姓名_In, 核收时间_In, 1, 申请类型_In,
         Decode(仪器id_In, 0, Null, 仪器id_In), Null, Null, 标本形态_In, 0, 执行科室_In, 检验人_In, 检验时间_In, 微生物标本_In, 标本类别_In, 检验备注_In,
         申请科室id_In, 申请人_In, 姓名_In, 性别_In, 年龄_In, Zl_Val(年龄_In),
         Decode(年龄_In, Null, Null, '成人', '成人', '婴儿', '婴儿',
                 Decode(Sign(Instr(年龄_In, '岁')), 1, '岁',
                         Decode(Sign(Instr(年龄_In, '月')), 1, '月',
                                 Decode(Sign(Instr(年龄_In, '天')), 1, '天', Decode(Sign(Instr(年龄_In, '小时')), 1, '小时', Null))))),
         病人id_In, Decode(r_Patient.住院号, Null, Decode(r_Patient.门诊号, Null, 3, 1), 2), 0, Null, Null, 标识号_In, 床号_In,
         病人科室_In, 标本类别_In, r_Patient.门诊号, r_Patient.住院号, r_Patient.出生日期, Null, Null, 检验项目_In, Null, Null, Null);
    End If;
    If Nvl(复盖标本id_In, 0) > 0 Then
      Zl_检验标本记录_Union(Id_In, 复盖标本id_In);
    End If;
    --记录核收和补填操作
    Insert Into 检验操作记录
      (ID, 标本id, 操作类型, 操作员, 操作时间)
    Values
      (检验操作记录_Id.Nextval, Id_In, 2, 人员姓名_In, Sysdate);
    Close v_Patient;
  Else
    Open v_Advice;
    Fetch v_Advice
      Into r_Advice;
  
    If v_Advice%Found Then
      Zl_病人信息_锁定检查(r_Advice.病人id);
    End If;
  
    Update 检验标本记录
    Set 医嘱id = Decode(医嘱id_In, Null, 医嘱id, 0, 医嘱id, 医嘱id_In), 采样时间 = Decode(采样时间_In, Null, 采样时间, 采样时间_In),
        采样人 = Decode(采样人_In, Null, 采样人, 采样人_In), 标本序号 = Decode(标本序号_In, Null, 标本序号, 标本序号_In),
        标本类型 = Decode(标本类型_In, Null, Decode(标本类型, Null, r_Advice.标本部位, 标本类型), 标本类型_In),
        申请时间 = Decode(r_Advice.开嘱时间, Null, 申请时间, r_Advice.开嘱时间), 核收人 = Decode(核收人, Null, 人员姓名_In, 核收人),
        样本条码 = Decode(r_Advice.样本条码, Null, 样本条码, r_Advice.样本条码), 申请类型 = Decode(申请类型_In, Null, 申请类型, 申请类型_In),
        执行科室id = Decode(执行科室_In, Null, 执行科室id, 执行科室_In), 检验人 = Decode(检验人_In, Null, 检验人, 检验人_In),
        检验时间 = Decode(检验时间_In, Null, 检验时间, 检验时间_In), 检验备注 = Decode(检验备注_In, Null, 检验备注, 检验备注_In),
        申请科室id = Decode(申请科室id_In, Null, 申请科室id, 申请科室id_In), 申请人 = Decode(申请人_In, Null, 申请人, 申请人_In),
        姓名 = Decode(姓名_In, Null, 姓名, 姓名_In), 性别 = Decode(性别_In, Null, 性别, 性别_In), 年龄 = Decode(年龄_In, Null, 年龄, 年龄_In),
        年龄数字 = Decode(年龄_In, Null, 年龄数字, Zl_Val(年龄_In)),
        年龄单位 = Decode(年龄_In, Null, 年龄单位,
                       Decode(年龄_In, Null, Null, '成人', '成人', '婴儿', '婴儿',
                               Decode(Sign(Instr(年龄_In, '岁')), 1, '岁',
                                       Decode(Sign(Instr(年龄_In, '月')), 1, '月',
                                               Decode(Sign(Instr(年龄_In, '天')), 1, '天',
                                                       Decode(Sign(Instr(年龄_In, '小时')), 1, '小时', Null)))))),
        病人id = Decode(r_Advice.病人id, Null, 病人id, r_Advice.病人id), 病人来源 = Decode(r_Advice.病人来源, Null, 病人来源, r_Advice.病人来源),
        婴儿 = Decode(r_Advice.婴儿, 婴儿, r_Advice.婴儿), NO = Decode(No_In, Null, NO, No_In), 合并id = v_Union,
        标本形态 = Decode(标本形态_In, Null, 标本形态, 标本形态_In), 标识号 = Decode(标识号_In, Null, 标识号, 标识号_In),
        床号 = Decode(床号_In, Null, 床号, 床号_In), 病人科室 = Decode(病人科室_In, Null, 病人科室, 病人科室_In), 标本类别 = 标本类别_In,
        门诊号 = r_Advice.门诊号, 住院号 = r_Advice.住院号, 出生日期 = r_Advice.出生日期, 挂号单 = r_Advice.挂号单, 主页id = r_Advice.主页id,
        检验项目 = Decode(检验项目_In, Null, 检验项目, 检验项目_In), 操作类型 = r_Advice.操作类型, 接收人 = r_Advice.接收人, 接收时间 = r_Advice.接收时间
    Where ID = Id_In;
  
    If Sql%NotFound Then
      Insert Into 检验标本记录
        (ID, 医嘱id, 标本序号, 采样时间, 采样人, 标本类型, 核收人, 核收时间, 样本状态, 申请类型, 仪器id, 样本条码, 申请时间, 标本形态, 报告结果, 执行科室id, 检验人, 检验时间, 微生物标本,
         标本类别, 检验备注, 申请科室id, 申请人, 姓名, 性别, 年龄, 年龄数字, 年龄单位, 病人id, 病人来源, 婴儿, NO, 合并id, 标识号, 床号, 病人科室, 紧急, 门诊号, 住院号, 出生日期,
         挂号单, 主页id, 检验项目, 操作类型, 接收人, 接收时间)
      Values
        (Id_In, Decode(医嘱id_In, 0, Null, 医嘱id_In), 标本序号_In, 采样时间_In, 采样人_In, Nvl(标本类型_In, r_Advice.标本部位), 人员姓名_In,
         核收时间_In, 1, 申请类型_In, Decode(仪器id_In, 0, Null, 仪器id_In), r_Advice.样本条码, r_Advice.开嘱时间, 标本形态_In, 0, 执行科室_In,
         检验人_In, 检验时间_In, 微生物标本_In, 标本类别_In, 检验备注_In, 申请科室id_In, 申请人_In, 姓名_In, 性别_In, 年龄_In, Zl_Val(年龄_In),
         Decode(年龄_In, Null, Null, '成人', '成人', '婴儿', '婴儿',
                 Decode(Sign(Instr(年龄_In, '岁')), 1, '岁',
                         Decode(Sign(Instr(年龄_In, '月')), 1, '月',
                                 Decode(Sign(Instr(年龄_In, '天')), 1, '天', Decode(Sign(Instr(年龄_In, '小时')), 1, '小时', Null))))),
         r_Advice.病人id, r_Advice.病人来源, r_Advice.婴儿, No_In, v_Union, 标识号_In, 床号_In, 病人科室_In, r_Advice.紧急, r_Advice.门诊号,
         r_Advice.住院号, r_Advice.出生日期, r_Advice.挂号单, r_Advice.主页id, 检验项目_In, r_Advice.操作类型, r_Advice.接收人, r_Advice.接收时间);
    End If;
    If Nvl(复盖标本id_In, 0) > 0 Then
      Zl_检验标本记录_Union(Id_In, 复盖标本id_In);
    End If;
    Insert Into 检验操作记录
      (ID, 标本id, 操作类型, 操作员, 操作时间)
    Values
      (检验操作记录_Id.Nextval, Id_In, 2, 人员姓名_In, Sysdate);
  
    --查找主项目有时填写合并ID
    Begin
      Select a.Id
      Into v_Union
      From 检验标本记录 A, 检验标本记录 B, 病人医嘱记录 C, 检验合并规则 D, 病人医嘱记录 E
      Where a.病人id = b.病人id And b.Id = Id_In And a.样本状态 = 1 And Nvl(a.病人id, 0) <> 0 And a.医嘱id = c.相关id And
            d.主项目id = c.诊疗项目id And d.合并项目id = e.诊疗项目id And e.Id = r_Advice.Id And Rownum = 1
      Order By a.核收时间 Desc;
    Exception
      When Others Then
        v_Union := Null;
    End;
    If Nvl(v_Union, 0) <> 0 Then
      Update 检验标本记录 Set 合并id = v_Union Where (ID = Id_In Or 医嘱id = r_Advice.Id);
    End If;
    --查找有了主项目时填写合并项目
    Begin
      Select a.Id, a.病人id, c.主项目id
      Into v_Union, v_Patientid, v_Itemid
      From 检验标本记录 A, 病人医嘱记录 B, 检验合并规则 C
      Where a.医嘱id = b.相关id And b.诊疗项目id = c.主项目id And a.Id = Id_In And Rownum = 1;
    Exception
      When Others Then
        v_Union := Null;
    End;
    If Nvl(v_Union, 0) <> 0 Then
      Update 检验标本记录
      Set 合并id = v_Union
      Where ID In (Select a.Id
                   From 检验标本记录 A, 病人医嘱记录 B, 检验合并规则 C
                   Where a.医嘱id = b.相关id And b.诊疗项目id = c.合并项目id And c.主项目id = v_Itemid And a.病人id = v_Patientid And
                         a.样本状态 = 1);
    End If;
  
    v_Seq := 1;
    Close v_Advice;
    v_Flag := 0;
    Begin
      Select Nvl(Max(1), 0) Into v_Flag From 检验申请项目 Where 标本id = Id_In;
    Exception
      When Others Then
        v_Flag := 0;
    End;
    If v_Flag = 0 Then
      For r_Advice In v_Advice Loop
        Update 检验申请项目
        Set 标本id = Id_In, 诊疗项目id = r_Advice.诊疗项目id
        Where 标本id = Id_In And 诊疗项目id = r_Advice.诊疗项目id;
        If Sql%Rowcount = 0 Then
          Insert Into 检验申请项目 (标本id, 诊疗项目id, 序号) Values (Id_In, r_Advice.诊疗项目id, v_Seq);
        End If;
        v_Seq := v_Seq + 1;
      End Loop;
    End If;
  
  End If;

  --根据参数来判断是否发料
    For r_Advice_1 In v_Advice_1 Loop
      --如果记帐没有自动发料,则自动发料,否则不处理
      For r_Stuff In c_Stuff(r_Advice_1.单据号, v_主页id) Loop
      
        Zl_材料收发记录_处方发料(r_Stuff.库房id, r_Stuff.单据, r_Stuff.单据号, 人员姓名_In, 人员姓名_In, 人员姓名_In, 1, Sysdate);
      End Loop;
    End Loop;


  Update /*+ Rule */ 病人医嘱发送
  Set 执行状态 = 3
  Where 执行状态 = 0 And
        医嘱id In (Select ID
                 From 病人医嘱记录
                 Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
                 Union All
                 Select ID
                 From 病人医嘱记录
                 Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist))));

  --执行后自动审核对应的记帐划价单(不包含药品)
  Select Zl_To_Number(Nvl(Zl_Getsysparameter(81), '0')) Into v_执行 From Dual;
  --2.检查当前标本相关的申请的相关标本是否完成审核
  For r_Samplequest In c_Samplequest(微生物标本_In) Loop
  
    v_Count := 0;
  
    --r_SampleQuest.医嘱id申请已经完成,处理后续环节
    If v_Count = 0 Then
    
      If r_Samplequest.病人来源 = 2 Then
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
      If Nvl(v_执行, 0) = 1 Then
        For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
          If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
            If v_序号 Is Not Null Then
              If r_Verify.门诊标志 = 1 Then
                Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
              Elsif r_Verify.门诊标志 = 2 Then
                Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
              End If;
            End If;
            v_序号 := Null;
          End If;
          v_门诊标志 := r_Verify.门诊标志;
          v_No       := r_Verify.No;
          v_性质     := r_Verify.记录性质;
          v_序号     := v_序号 || ',' || r_Verify.序号;
        End Loop;
        If v_序号 Is Not Null Then
          If v_门诊标志 = 1 Then
            Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          Elsif v_门诊标志 = 2 Then
            Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          End If;
        End If;
      End If;
    
    End If;
  End Loop;

  If Nvl(申请类型_In, 0) = 1 Then
    Zl_病人医嘱记录_屏蔽打印(医嘱id_In, 1);
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_检验标本记录_标本核收;
/

--93144:李南春,2016-02-29,预交退款金额检查
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
  更新交款余额_In  Number := 0,--是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况。
  退款方式_In   Number := 0
) As
  ----------------------------------------------
  --操作类型_In:0-正常缴预交;1-存为划价单;3-余额退款
  --退款方式_In;0-提示;1-禁止；2-忽略
  v_Err_Msg Varchar2(200);
  n_Err_Num Number;
  Err_Item Exception;

  v_性质   结算方式.性质%Type;
  v_打印id 票据打印内容.Id%Type;
  v_担保   病人信息.担保性质%Type;
  v_Date   Date;
  n_返回值 病人余额.预交余额%Type;
  n_组id   财务缴款分组.Id%Type;
  n_病人余额 病人余额.预交余额%Type;
  n_三方预交 病人余额.预交余额%Type;
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
  
  If 金额_In<0 then
    Begin
      Select Nvl(预交余额,0)-Nvl(费用余额,0) into n_病人余额 From 病人余额 Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 0) = Nvl(预交类别_In, 0);
    Exception
      When Others Then
        Null;
    End;
    --余额退款要考虑三方预交是否支持退现
    IF 操作类型_In = 3 then
      For c_三方预交 In (Select A.预交ID,A.预交类别,A.卡类别ID,A.结算卡序号 as 消费接口ID,nvl(B.编码,C.编号) as 编码,nvl(B.名称,C.名称) as 名称,
                              Decode(B.编码,NULL,C.是否全退,B.是否全退) as 是否全退,Decode(B.编码,NULL,C.是否退现,B.是否退现) as 是否退现,
                              A.卡号,A.交易流水号,A.交易说明,          A.预交余额
                              From (   Select A.预交类别,nvl(A.卡类别ID,0) as 卡类别ID,nvl(A.结算卡序号,0) as 结算卡序号,
                              A.卡号,A.交易流水号,A.交易说明,           max(decode(sign(金额),-1, decode(A.记录状态,1,0,2,0,ID),ID)) as 预交ID,
                              nvl(sum(金额),0)-nvl(sum(nvl(冲预交,0)),0) as 预交余额    From 病人预交记录 A    Where   A.病人ID=病人id_In and (nvl(A.结算卡序号,0)<>0 or nvl(卡类别ID,0)<>0)
                              Group by A.预交类别,nvl(A.卡类别ID,0),nvl(A.结算卡序号,0),A.卡号,A.交易流水号,A.交易说明
                              Having nvl(sum(金额),0)-nvl(sum(nvl(冲预交,0)),0)  <>0 ) A,医疗卡类别 B,卡消费接口目录 C
                              Where A.预交类别 =Nvl(预交类别_In, 0) And  A.卡类别ID=B.ID(+)  and A.结算卡序号=C.编号(+)  and nvl(A.预交余额,0)<>0   Order by 编码,A.卡号,A.交易流水号,A.交易说明) Loop

        IF instr(',7,8,',','|| v_性质 || ',') =0 And Nvl(c_三方预交.是否退现,0) = 0 And Nvl(c_三方预交.预交余额,0) > 0 Then
          n_三方预交:= Nvl(n_三方预交,0) + Nvl(c_三方预交.预交余额,0);
        ElsIf instr(',7,8,',','|| v_性质 || ',') >0 Then
              If Nvl(c_三方预交.卡号,'0') = Nvl(卡号_In,'0') And Nvl(c_三方预交.交易流水号,'0')= Nvl(交易流水号_In,'0') And Nvl(c_三方预交.交易说明,'0') = Nvl(交易说明_In,'0') then
                n_三方预交:= Nvl(n_三方预交,0) + Nvl(c_三方预交.预交余额,0);
              End if;
        End IF;
      End Loop;
    End if;
    
    If instr(',7,8,',','|| v_性质 || ',') > 0 And Nvl(n_三方预交,0) < 0 And 操作类型_In = 3 Then
        n_Err_Num := -20101;
        v_Err_Msg :='退款金额大于病人三方预交金额，不能继续！';
        Raise Err_Item;
    ElsIf Nvl(n_病人余额,0) < 0 And 退款方式_In <> 2 Then
        n_Err_Num := -20101;
        v_Err_Msg :='退款金额大于病人剩余预交余额，不能继续！';
        If 退款方式_In = 0 Then
          n_Err_Num := -20111;
          v_Err_Msg :='退款金额大于病人剩余预交余额，是否忽略？';
        End If;
        Raise Err_Item;
    ElsIf instr(',7,8,',','|| v_性质 || ',') = 0 And Nvl(n_病人余额,0) - Nvl(n_三方预交,0) < 0 And 操作类型_In = 3 And 退款方式_In <> 2 then
        n_Err_Num := -20101;
        v_Err_Msg :='退款金额大于病人剩余预交余额，不能继续！';
        If 退款方式_In = 0 Then
          n_Err_Num := -20111;
          v_Err_Msg :='退款金额大于病人剩余预交余额，是否忽略？';
        End If;
        Raise Err_Item;
    End if;
  End if;

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
  If 操作类型_In <> 1 Then
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
    Raise_Application_Error(n_Err_Num, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_Insert;
/

--94030:胡俊勇,2016-03-16,危机值消息阅读
--92384:胡俊勇,2016-02-24,医嘱待执行消息
CREATE OR REPLACE Procedure Zl_业务消息清单_Read
(
  病人id_In     In 业务消息清单.病人id%Type,
  就诊id_In     In 业务消息清单.就诊id%Type,
  类型编码_In   In 业务消息清单.类型编码%Type,
  阅读场合_In   In 业务消息状态.阅读场合%Type,
  阅读人_In     In 业务消息状态.阅读人%Type,
  阅读部门id_In In 业务消息状态.阅读部门id%Type,
  阅读时间_In   In 业务消息状态.阅读时间%Type := Null,
  消息id_In     In 业务消息状态.消息id%Type := Null,
  业务标识_In   In 业务消息清单.业务标识 %Type := Null
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
  Elsif 业务标识_In Is Not Null Then
    For R In (Select a.Id
              From 业务消息清单 A
              Where a.病人id = 病人id_In And a.就诊id = 就诊id_In And a.类型编码 = 类型编码_In And a.业务标识 = 业务标识_In And
                    Nvl(a.是否已阅, 0) = 0) Loop
      Insert Into 业务消息状态
        (消息id, 阅读场合, 阅读人, 阅读时间, 阅读部门id)
      Values
        (r.Id, 阅读场合_In, 阅读人_In, d_Cur, 阅读部门id_In);
    End Loop;
    Update 业务消息清单
    Set 是否已阅 = 1
    Where 病人id = 病人id_In And 就诊id = 就诊id_In And 类型编码 = 类型编码_In And Nvl(是否已阅, 0) = 0 And 业务标识 = 业务标识_In;
  Elsif 类型编码_In = 'ZLHIS_CIS_034' Then
    --待执行消息特殊处理
    For R In (Select ID
              From 业务消息清单 A, 业务消息提醒部门 B
              Where a.Id = b.消息id And a.病人id = 病人id_In And a.就诊id = 就诊id_In And a.类型编码 = 类型编码_In And Nvl(a.是否已阅, 0) = 0 And
                    b.部门id = 阅读部门id_In
              Group By ID) Loop
      Insert Into 业务消息状态
        (消息id, 阅读场合, 阅读人, 阅读时间, 阅读部门id)
      Values
        (r.Id, 阅读场合_In, 阅读人_In, d_Cur, 阅读部门id_In);
      Update 业务消息清单 Set 是否已阅 = 1 Where ID = r.Id;
    End Loop;
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

--91316:张德婷,2015-02-23,输液配置中心批次
--91954:张德婷,2016-03-15,按照容量规则自动排批
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
  v_药品类型       varchar2(20);
  n_打包药品批次   number(1);
  n_特殊药品批次   number(1);
  n_优先级         number:=999;
  n_自动排批       number:=0;
  n_科室id       number:=0;
  n_row            number(2);
  Err_Item Exception;

  Cursor c_医嘱记录 Is
    Select /*+rule */
    Distinct e.医嘱id As 相关id, e.发送号, b.频率间隔, b.间隔单位, b.执行时间方案, a.姓名, a.性别, a.年龄, a.住院号, a.当前床号 As 床号, a.当前病区id As 病人病区id,
             a.当前科室id As 病人科室id, e.首次时间, e.末次时间, b.开始执行时间, Nvl(e.发送数次, 0) As 次数, e.发送时间, Decode(b.医嘱期效, 0, 1, 2) As 医嘱类型,
             b.诊疗项目id As 给药途径, b.病人id, Nvl(c.执行标记, 0) As 是否tpn
    From 病人医嘱发送 E, 病人医嘱记录 B, 病人信息 A, 诊疗项目目录 C, Table(f_Num2list(医嘱id_In)) D
    Where e.医嘱id = b.Id And b.病人id = a.病人id And  c.类别 = 'E' And c.操作类型 = '2' And c.执行分类 = 1 And b.诊疗项目id = c.Id And
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
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No,F.抗生素,F.是否肿瘤药
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, 输液药品属性 E,药品特性 F,药品规格 G
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id=G.药品ID and G.药名ID=F.药名ID And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.相关id = v_相关id And b.发送号 = 发送号_In And c.序号 < 1000
    Order By c.No, c.序号;

  v_医嘱记录     c_医嘱记录%RowType;
  v_收发记录     c_收发记录%RowType;
  v_单个医嘱记录 c_单个医嘱记录%RowType;

  Function Zl_Getpivaworkbatch(执行时间_In In Date,药品类型_In In varchar2:=null) Return Number As
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_配药批次 Is
      Select 批次, 配药时间, 给药时间, 打包,药品类型 From 配药工作批次 Where 启用 = 1 and 配置中心id=部门id_In Order By 批次;

    v_配药批次 c_配药批次%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');

    Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次 Where 启用 = 1 and 配置中心id=部门id_In;

    For v_配药批次 In c_配药批次 Loop
      v_Batch     := 0;

      if 药品类型_In is null then
        if v_配药批次.批次<>'0'and v_配药批次.药品类型 is null then
          v_Starttime := To_Date(Substr(v_配药批次.给药时间, 1, Instr(v_配药批次.给药时间, '-') - 1), 'hh24:mi');
          v_Endtime   := To_Date(Substr(v_配药批次.给药时间, Instr(v_配药批次.给药时间, '-') + 1), 'hh24:mi');

          If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
            v_Batch := v_配药批次.批次;
            n_打包  := v_配药批次.打包;
            Exit When v_Batch > 0;
          End If;
        end if;
      else
        if 药品类型_In=v_配药批次.药品类型 then
          v_Batch := v_配药批次.批次;
          n_打包  := v_配药批次.打包;
          Exit When v_Batch > 0;
        end if;
      end if;
    End Loop;

    If v_Batch = 0 and n_打包药品批次<>1 Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;

  Function Zl_GetFirst(配药id_In In number,科室id_In In number) Return Number As
    n_First     Number;
    n_科室id    number;
    Cursor c_优先级 Is
      select 科室id,配药类型,优先级,频次 from 输液药品优先级 where (科室id=科室id_In or 科室id=0) order by 科室id,优先级 desc;

    r_优先级 c_优先级%RowType;
  Begin
   n_First:=0;
   for r_优先级 in c_优先级 loop
     if n_科室id<>0 and r_优先级.科室id=0 then
       exit;
     end if;
     n_科室id:=r_优先级.科室id;

     for r_配药记录 in(select distinct D.配药类型,E.执行频次 from 输液配药记录 A,输液配药内容 B,药品收发记录 C,输液药品属性 D,病人医嘱记录 E  where A.医嘱id=E.Id and A.id=B.记录ID and B.收发ID=C.id and C.药品ID=D.药品ID and a.id=配药id_In) loop
       if instr(r_配药记录.配药类型,r_优先级.配药类型,1)>0 and instr(r_优先级.频次,r_配药记录.执行频次,1)>0 then
         n_First:= r_优先级.优先级;
         exit;
       end if;
     end loop;
   end loop;

   if n_First=0 then
     n_First:=999;
   end if;
   Return(n_First);
  End;
Begin
  n_Count          := 0;
  v_医嘱类型       := Zl_To_Number(Nvl(zl_GetSysParameter('医嘱类型', 1345), 1));
  v_输液总量       := Zl_To_Number(Nvl(zl_GetSysParameter('同批次输液总量', 1345), 0));
  v_大输液剂型     := Nvl(zl_GetSysParameter('大输液药品剂型', 1345), '');
  v_大输液给药途径 := Nvl(zl_GetSysParameter('输液给药途径', 1345), '');
  v_来源科室       := Nvl(zl_GetSysParameter('来源科室', 1345), '');
  v_保持上次批次   := Zl_To_Number(Nvl(zl_GetSysParameter('保持上次批次', 1345), 0));
  n_Tpn处置方式    := Zl_To_Number(Nvl(zl_GetSysParameter('静脉营养药物处置方式', 1345), 0));
  n_打包药品批次   := Zl_To_Number(Nvl(zl_GetSysParameter('单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包', 1345), 0));
  n_特殊药品批次   := Zl_To_Number(Nvl(zl_GetSysParameter('特殊药品按药品类型指定批次', 1345), 0));
  n_自动排批:= Zl_To_Number(Nvl(zl_GetSysParameter('启动自动排批', 1345), 0));
  v_医嘱ids  := 医嘱id_In;
  v_当前病人 := '';

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
    n_病人id := v_医嘱记录.病人id;
    n_科室id:= v_医嘱记录.病人科室id;

    Select Count(1)
    Into v_Continue
    From 病人医嘱记录 A, 输液不配置药品 B,住院费用记录 C
    Where c.收费细目id = b.药品id And c.医嘱序号 =A.id and A.相关id= v_医嘱记录.相关id;
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
      If Instr(',' || v_来源科室 || ',', ',' || v_医嘱记录.病人科室id || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;

    v_药品类型:=null;
    for r_药品类型 in (Select decode(nvl(D.是否肿瘤药,0),0,'','肿瘤药') 药品类型
    From 病人医嘱记录 A, 药品规格 B,住院费用记录 C,药品特性 D
    Where c.收费细目id = b.药品id And B.药名ID=D.药名ID And c.医嘱序号 =A.id and A.相关id= v_医嘱记录.相关id) loop
      if r_药品类型.药品类型 is not null then
        v_药品类型:=r_药品类型.药品类型;
      end if;
    end loop;

    if v_药品类型 is null then
       If v_医嘱记录.是否tpn = 2 Then
        v_药品类型:='营养药';
        v_Continue := 1;
      end if;
    end if;

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
            --v_执行时间 := Zl_Gettransexetime(v_医嘱记录.开始执行时间, v_执行时间, v_医嘱记录.频率间隔, v_医嘱记录.间隔单位, v_医嘱记录.执行时间方案);
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
          --药品类型不为空，则根据药品类型匹配批次

          if (v_药品类型 is null or n_特殊药品批次=0) and n_自动排批=0 then
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

            If (v_保持上次批次 = 1 Or b_Change = True) and n_自动排批=0 Then
              --取上次的批次
              Begin
                Select Distinct 配药批次
                Into v_批次
                From 输液配药记录 A
                Where 医嘱id = v_医嘱记录.相关id And
                      发送号 = (Select Distinct Max(发送号)
                             From 输液配药记录
                             Where 医嘱id = v_医嘱记录.相关id And 发送号 <> v_医嘱记录.发送号) And
                      To_Char(a.执行时间, 'hh24:mi:ss') = To_Char(v_执行时间, 'hh24:mi:ss');
              Exception
                When Others Then
                  v_批次 := 0;
              End;
            End If;

            If v_批次 = 0 Then
              v_批次 := Zl_Getpivaworkbatch(v_执行时间);

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

          elsif v_药品类型 is not null and n_特殊药品批次=1 then
            --药品类型不为空，直接根据药品类型匹配批次
            v_批次 := Zl_Getpivaworkbatch(sysdate,v_药品类型);
          else
            v_批次 := Zl_Getpivaworkbatch(v_执行时间);
          end if;  

          Select Count(医嘱id)
          Into n_发送次数
          From 医嘱执行时间
          Where 医嘱id = v_医嘱记录.相关id And 要求时间 <= v_执行时间
          Order By 要求时间;

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
            Select Nvl(Max(打包), 0) Into n_打包 From 配药工作批次 Where 批次 = v_批次;
          End If;

          If (Trunc(v_执行时间) <= v_Currdate Or n_打包 <> 0) and (v_药品类型 is null or n_特殊药品批次=0) Then
            n_是否打包     := 1;
            d_手工打包时间 := Sysdate;
          Else
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;

          --如果是TPN不管其他条件如何都设置为配置
          If v_医嘱记录.是否tpn = 2 Then
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;

          if v_批次=0 then
             n_是否打包:=1;
          end if;
          --产生配药记录
          Insert Into 输液配药记录
            (ID, 部门id, 序号, 姓名, 性别, 年龄, 住院号, 床号, 病人病区id, 病人科室id, 执行时间, 医嘱id, 发送号, 配药批次, 瓶签号, 是否调整批次, 是否打包, 打包时间, 操作状态,
             操作人员, 操作时间)
          Values
            (v_配药id, 部门id_In, v_序号, v_医嘱记录.姓名, v_医嘱记录.性别, v_医嘱记录.年龄, v_医嘱记录.住院号, v_医嘱记录.床号, v_医嘱记录.病人病区id,
             v_医嘱记录.病人科室id, v_执行时间, v_医嘱记录.相关id, v_医嘱记录.发送号, v_批次, v_Maxno, n_调整批次, n_是否打包,
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

            n_优先级:=Zl_GetFirst(v_配药id,v_医嘱记录.病人科室id);
            update 输液配药记录 set 优先级=n_优先级 where id=v_配药id;
          End Loop;

        End Loop;

        For v_收发记录 In c_收发记录 Loop
          n_单据 := v_收发记录.单据;

          v_No := v_收发记录.No;
          Delete From 药品收发记录 Where ID = v_收发记录.收发id;
        End Loop;

        --单个药品或者不予配置的药品默认为0批次
        select count(收发id) into n_Row from 输液配药内容 where 记录id=v_配药id;
        if (v_Nodosage = 1 or n_row=1) and n_打包药品批次=1 then
          Update 输液配药记录 Set 配药批次 = 0,是否打包 = 1 Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 and 操作状态<2;
        end if;
        --如果存在“不予配置”属性的药品，也设置为打包
        If v_Nodosage = 1 Then
          Update 输液配药记录 Set 是否打包 = 1 Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 and 操作状态<2;
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

  if n_自动排批=1 then
    Zl_输液配药记录_自动排批(n_病人id,n_科室id,部门id_In,v_执行时间);
  end if;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]病人' || v_当前病人 || '在输液配置中心有被锁定的输液单，发送失败！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_核查;
/


--90447:刘尔旋,2016-02-22,复诊的确定
Create Or Replace Function Zl1_Fun_Getreturnvisit
(
  病人id_In 病人信息.病人id%Type,
  科室id_In 部门表.Id%Type
) Return Number As
  --------------------------------------------------------------------------------------------------
  --功能:挂号时判断病人是否复诊
  --默认规则:当前挂号科室的临床性质或临床性质取两位的编码且在30天内存在挂号记录的,为复诊
  --入参:
  --  病人ID_In   : 挂号病人ID
  --  科室ID_In   : 挂号科室ID
  --返回:是否复诊,1-复诊,0-初诊
  --------------------------------------------------------------------------------------------------
  v_挂号性质 临床部门.工作性质%Type;
  n_Exists   Number(3);
  v_部门ids  Varchar2(4000);
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
Begin
  Begin
    Select 工作性质 Into v_挂号性质 From 临床部门 Where 部门id = 科室id_In And Rownum < 2;
  Exception
    When Others Then
      Return 0;
  End;
  v_挂号性质 := Substr(v_挂号性质, 1, 2) || '%';
  For r_部门 In (Select Distinct 部门id From 临床部门 Where 工作性质 Like v_挂号性质) Loop
    v_部门ids := v_部门ids || ',' || r_部门.部门id;
  End Loop;
  If v_部门ids Is Not Null Then
    v_部门ids := Substr(v_部门ids, 2);
  End If;
  n_Exists := 0;
  Select Max(1)
  Into n_Exists
  From 病人挂号记录
  Where 登记时间 > Sysdate - 30 And 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And
        执行部门id In (Select Column_Value From Table(f_Str2list(v_部门ids)));
  Return Nvl(n_Exists, 0);
Exception
  When Err_Item Then
    v_Err_Msg := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/


--91225:梁经伙,2016-02-16,疾病阳性记录新增字段 医嘱ID
CREATE OR REPLACE Procedure Zl_疾病阳性检测记录_Insert
(
  Id_In         In 疾病阳性记录.Id%Type,
  病人id_In     In 疾病阳性记录.病人id%Type,
  主页id_In     In 疾病阳性记录.主页id%Type,
  挂号单_In     In 疾病阳性记录.挂号单%Type,
  医嘱ID_In     In 疾病阳性记录.医嘱ID%Type,
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
    (ID, 病人id, 主页id, 挂号单, 医嘱ID,送检时间, 送检科室id, 送检医生, 标本名称, 反馈结果, 传染病名称, 检查时间, 登记时间, 登记人, 登记科室id, 记录状态)
  Values
    (Id_In, 病人id_In, 主页id_In, 挂号单_In, 医嘱ID_In,送检时间_In, 送检科室id_In, 送检医生_In, 标本名称_In, 反馈结果_In, 传染病_In, 检查时间_In, 登记时间_In, 登记人_In,
     登记科室id_In, 记录状态_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_疾病阳性检测记录_Insert;
/

--93336:胡俊勇,2016-02-16,输血执行登记
Create Or Replace Function Zl_Fun_Get输血执行登记(医嘱id_In In 病人医嘱记录.Id%Type) Return Varchar2 Is
  n_Count 病人医嘱执行.本次数次%Type;
  n_Boold Number; --是否用血库系统
  v_Tmp   Varchar2(4000);
  n_Bags  Number; --血库配血袋数
  --功能：对 输血途径 医嘱执行登记
  --参数：医嘱id_In 主医嘱ID
  --返回：固定格式字符串， 登记一次所填写的数量，是否启用血库，血袋总数。例：0.33333<SPLIT>1<SPLIT>3
  --说明：后台数据以5位小数四舍五入保存，界面提示和判断时要乘上血袋总数四舍五入取整进行比较
Begin
  Select Count(1) Into n_Boold From zlSystems Where 编号 = 2200;
  If n_Boold = 1 Then
    n_Boold := Nvl(zl_GetSysParameter(236), 0);
  End If;
  If n_Boold = 1 Then
    v_Tmp := 'Select Zl_Get_输血执行次数(:1) as 数量 From Dual';
    Execute Immediate v_Tmp
      Into n_Bags
      Using 医嘱id_In;
  End If;
  If n_Bags Is Null Then
    n_Count := 1;
  Else
    n_Count := Round(1 / n_Bags, 5);
  End If;
  v_Tmp := To_Char(n_Count, '0.99999') || '<SPLIT>' || Nvl(n_Boold, 0) || '<SPLIT>' || Nvl(n_Bags, 0);
  Return v_Tmp;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Get输血执行登记;
/  
--91646:张德婷,2016-02-22,配药收取材料费
CREATE OR REPLACE Function zl_fun_PIVACustom
( 
  配药id_In     In 输液配药记录.ID%type
) Return Varchar2 Is 
  ---------------------------------- 
  --功能：静配配药时，根据药品的配药类型，使用单量，给药途径确定收费的材料id；
  --返回：
  --      收费项目id1m收取数次；收费项目id2,收费数次   例如：12,1;13,2  即要收费项目id为12的收取一次费用，收费项目id为13的收取两次费用
  ----------------------------------- 
  v_Temp   varchar2(500);
  Err_Custom Exception; 
  v_Error Varchar2(255); 
  v_IsSpecial varchar2(20);
Begin
  v_Temp:='';
  for r_item in (select distinct A.药品id,A.配药类型,F.名称 给药途径,D.单量,G.剂量系数 from 输液药品属性 A,输液配药记录 B,输液配药内容 C,药品收发记录 D,病人医嘱记录 E,诊疗项目目录 F,药品规格 G where B.医嘱id=E.id and E.诊疗项目id=F.id and B.Id=C.记录ID and C.收发ID=D.ID and D.药品ID=A.药品ID and D.药品ID=G.药品ID and
    B.Id= 配药id_In) loop
    --判断是否为胰岛素
    if r_item.配药类型='胰岛素' then
      --1ml注射器,收费数次
      v_IsSpecial:='';
    end if;
      
    if r_item.配药类型='肿瘤药' then
      --根据给药途径判断
       if  r_item.给药途径='泵注' then
         --60ml注射器,收费数次
         v_Temp:='';
       else 
         --20ml注射器,收费数次
          v_Temp:='';
       end if;
       exit; 
    end if;
  
    if r_item.配药类型='营养液' then
       --各个收费项目分别连接
       v_Temp:='';   
       
       exit;   
    end if;
  
    --判断特殊药品；找出对应的收费id
    if  r_item.药品id='1' then
      --60ml注射器,收费数次
      v_Temp:='';  
    elsif v_Temp is null then
      --20ml注射器,收费数次
      v_Temp:='';
    end if;
  
  end loop;
  Return v_Temp || v_IsSpecial; 
Exception 
  When Err_Custom Then 
    Return Null; 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End zl_fun_PIVACustom;
/       

--91646:张德婷,2016-02-03,配药收取材料费
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
  n_数次 number(2);
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
        if n_count=0 and (n_OutNum=0 or n_out=0) then
          --收取材料费
          --v_收费项目id:='6970,2;6971,1;';
          select zl_fun_PIVACustom(v_Tansid) into  v_收费项目id from dual;
          While v_收费项目id Is Not Null Loop
             v_info:= Substr(v_收费项目id, 1, Instr(v_收费项目id, ';') - 1);
             v_收费项目id := Replace(';' || v_收费项目id, ';' || v_info || ';');
             

             v_id:= Substr(v_info, 1, Instr(v_info, ',') - 1);
             v_info := Replace(',' || v_info, ',' || v_id || ',');

            For r_Item In (Select a.Id 收费细目id, a.类别 收费类别, a.计算单位, a.加班加价 加班标志, d.Id 收入项目id, d.收据费目, b.现价
                           From 收费项目目录 A, 收费价目 B, 收入项目 D
                           Where a.Id = b.收费细目id And b.收入项目id = d.Id And a.id=v_id and
                                 b.执行日期 <= Sysdate And
                                 (b.终止日期 >= Sysdate Or b.终止日期 Is Null)) Loop
              if n_count=0 then
                Insert Into 输液配药附费 (配药id, NO,病人id) Values (v_Tansid, v_No, r_Bill.病人id);
              end if;

              n_count:=n_count+1;
              Zl_住院记帐记录_Insert(v_No, n_count, r_Bill.病人id, r_Bill.主页id, r_Bill.标识号, r_Bill.姓名, r_Bill.性别, r_Bill.年龄, r_Bill.床号,
                               r_Bill.费别, r_Bill.病人病区id, r_Bill.病人科室id, r_Item.加班标志, r_Bill.婴儿费, r_Bill.库房id, 操作人员_In, Null,
                               r_Item.收费细目id, r_Item.收费类别, r_Item.计算单位, Null, Null, Null, 1,v_info, Null, r_Bill.库房id, Null,
                               r_Item.收入项目id, r_Item.收据费目, r_Item.现价, r_Item.现价*v_info, r_Item.现价*v_info, Null, Sysdate, Sysdate, Null, Null,
                               v_Usercode, 操作人员_In);
            End Loop;
          end loop;
        end if;


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

--92898:梁经伙,2016-02-03,过敏时间允许为空，不输入过敏时间的话不以当前时间替换
CREATE OR REPLACE Procedure Zl_病人过敏记录_Insert
( 
  病人id_In     病人过敏记录.病人id%Type, 
  主页id_In     病人过敏记录.主页id%Type, 
  来源_In       病人过敏记录.记录来源%Type, 
  药物id_In     病人过敏记录.药物id%Type, 
  药物名_In     病人过敏记录.药物名%Type, 
  结果_In       病人过敏记录.结果%Type := 1, 
  过敏时间_In   病人过敏记录.过敏时间%Type := Null, 
  记录时间_In   病人过敏记录.记录时间%Type := Null, 
  过敏反应_In   病人过敏记录.过敏反应%Type := Null, 
  过敏源编码_In 病人过敏记录.过敏源编码%Type := Null 
) Is 
  V_Date     Date; 
  V_Temp     Varchar2(255); 
  V_人员姓名 病人过敏记录.记录人%Type; 
  N_Count    Number; 
Begin 
  V_Temp     := Zl_Identity; 
  V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1); 
  V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1); 
  V_人员姓名 := Substr(V_Temp, Instr(V_Temp, ',') + 1); 
 
  If 记录时间_In Is Not Null Then 
    V_Date := 记录时间_In; 
  Else 
    Select Sysdate Into V_Date From Dual; 
  End If; 
 
  Insert Into 病人过敏记录 
    (ID, 病人id, 主页id, 记录来源, 药物id, 药物名, 结果, 过敏时间, 记录时间, 记录人, 过敏反应,过敏源编码) 
  Values 
    (病人过敏记录_Id.Nextval, 病人id_In, 主页id_In, 来源_In, 药物id_In, 药物名_In, 结果_In, 过敏时间_In, V_Date, V_人员姓名, 过敏反应_In,过敏源编码_In); 
 
  Select Count(1) Into N_Count From 病人过敏药物 Where 病人id = 病人id_In And 过敏药物 = 药物名_In; 
  If N_Count = 0 Then 
    Insert Into 病人过敏药物 
      (病人id, 过敏药物id, 过敏药物, 过敏反应) 
    Values 
      (病人id_In, 药物id_In, 药物名_In, 过敏反应_In); 
  Else 
    Update 病人过敏药物 
    Set 过敏药物id = 药物id_In, 过敏反应 = 过敏反应_In 
    Where 病人id = 病人id_In And 过敏药物 = 药物名_In; 
  End If; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_病人过敏记录_Insert;
/

--93209:蔡青松,2016-02-01,调整条码长度,当医嘱ID长度大于8时,将条码由原来的12位调整为13位
Create Or Replace Function Nextno
(
  序号_In     In 号码控制表.项目序号%Type,
  科室id_In   In 部门表.Id%Type := Null,
  v_Tag       In Varchar2 := Null,
  号码个数_In In Integer := 1
) Return Varchar2
--    功能：根据特定规则产生新的号码,规则如下：
  --    一、项目序号：
  --       1   病人ID         数字
  --       2   住院号         数字
  --       3   门诊号         数字
  --       10  医嘱发送号     数字,顺序递增编号
  --       x   其它单据号     字符,根据编号规则顺序递增编号,不自动补缺
  --    二、年度位确定原则：
  --       以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码
  --
  --    说明：最大号码-10存入号码控制表,用于并发情况下补缺号(取了号,但未使用)
  --          For Update在并发情况下锁定行,不用Wait选项以避免向调用者返回空
  --          v_Tag 可用其它未指明的参数,目前 有影像检查类别(编码)
  --    返回：最大号码
 Is
  Pragma Autonomous_Transaction;
  v_No     号码控制表.最大号码%Type;
  v_Maxno  号码控制表.最大号码%Type;
  n_Maxno  Number;
  n_Amt    Number;
  n_Mod    号码控制表.编号规则%Type;
  v_Deptno Varchar2(20);
  v_Year   Varchar2(1);
  v_Tmp    Varchar2(10);

  v_试管编码   Number;
  v_生成条码   Varchar2(20);
  v_编码       Varchar2(10);
  v_医嘱       Varchar2(18);
  v_Error      Varchar2(255);
  n_Checkmaxno Number;

  Err_Custom Exception;
Begin

  --1.病人ID
  If 序号_In = 1 Then
    Select Nvl(编号规则, 0) Into n_Mod From 号码控制表 Where 项目序号 = 序号_In;
  
    --从序列取值，用于不要求病人ID必须连续的用户减少并发争用
    If n_Mod = 1 Then
      Select 病人信息_Id.Nextval Into v_No From Dual;
    Else
      Select Nvl(最大号码, '0') Into v_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
    
      Select Nvl(Max(病人id), 0) + 1 Into n_Maxno From 病人信息 Where 病人id >= To_Number(v_Maxno);
    
      Update 号码控制表 Set 最大号码 = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1) Where 项目序号 = 序号_In;
      v_No := To_Char(n_Maxno);
    End If;
    --2.住院号
  Elsif 序号_In = 2 Then
    Select Nvl(最大号码, '0'), Nvl(编号规则, 0)
    Into v_Maxno, n_Mod
    From 号码控制表
    Where 项目序号 = 序号_In
    For Update;
  
    If n_Mod = 0 Then
      --0.顺序编号
      Select Nvl(Max(住院号), 0) + 1 Into n_Maxno From 病人信息 Where 住院号 >= To_Number(v_Maxno);
    
      Update 号码控制表 Set 最大号码 = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1) Where 项目序号 = 序号_In;
    Elsif n_Mod = 1 Then
      --1.年月(YYMM)+顺序号(0000)
      v_Tmp := To_Char(Sysdate, 'YYMM');
    
      Select Nvl(Max(住院号), To_Number(v_Tmp || '0000')) + 1
      Into n_Maxno
      From 病人信息
      Where 住院号 Like To_Number(v_Tmp) || '%' And 住院号 >= To_Number(v_Maxno);
      Update 号码控制表
      Set 最大号码 = Decode(Sign(n_Maxno - 10 - To_Number(v_Tmp || '0000')), 1, n_Maxno - 10, To_Number(v_Tmp || '0001'))
      Where 项目序号 = 序号_In;
    
    Elsif n_Mod = 2 Then
      --2.年(YYYY)+顺序号(00000)
      v_Tmp := To_Char(Sysdate, 'YYYY');
    
      Select Nvl(Max(住院号), To_Number(v_Tmp || '00000')) + 1
      Into n_Maxno
      From 病人信息
      Where 住院号 Like To_Number(v_Tmp) || '%' And 住院号 >= To_Number(v_Maxno);
      Update 号码控制表
      Set 最大号码 = Decode(Sign(n_Maxno - 10 - To_Number(v_Tmp || '00000')), 1, n_Maxno - 10, To_Number(v_Tmp || '00001'))
      Where 项目序号 = 序号_In;
    
    End If;
    v_No := To_Char(n_Maxno);
  
    --3.门诊号
  Elsif 序号_In = 3 Then
    Select Nvl(最大号码, '0'), Nvl(编号规则, 0)
    Into v_Maxno, n_Mod
    From 号码控制表
    Where 项目序号 = 序号_In
    For Update;
  
    If n_Mod = 0 Then
      --0.顺序编号
    
      Select Nvl(Max(门诊号), 0) + 1 Into n_Maxno From 病人信息 Where 门诊号 >= To_Number(v_Maxno);
    
      Update 号码控制表 Set 最大号码 = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1) Where 项目序号 = 序号_In;
    Elsif n_Mod = 1 Then
      --1.日期编号YYMMDD
      v_Tmp := To_Char(Sysdate, 'YYMMDD');
    
      Select Nvl(Max(门诊号), To_Number(v_Tmp || '0000')) + 1
      Into n_Maxno
      From 病人信息
      Where 门诊号 Like To_Number(v_Tmp) || '%' And 门诊号 >= To_Number(v_Maxno);
      Update 号码控制表
      Set 最大号码 = Decode(Sign(n_Maxno - 10 - To_Number(v_Tmp || '0000')), 1, n_Maxno - 10, To_Number(v_Tmp || '0001'))
      Where 项目序号 = 序号_In;
    
    End If;
    v_No := To_Char(n_Maxno);
  
    --10.医嘱发送号
  Elsif 序号_In = 10 Then
    Select Nvl(最大号码, '0') Into v_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
    v_Maxno := To_Char(To_Number(v_Maxno) + 1);
    Update 号码控制表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In;
    v_No := v_Maxno;
  
    --汇总发药号
  Elsif 序号_In = 20 Then
    --YYYYMMDD+5位顺序号(00000)
    Select To_Char(Sysdate, 'yyyymmdd') Into v_Tmp From Dual;
    Select Nvl(最大号码, '0') Into v_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
  
    If v_Maxno = '0' Then
      v_Maxno := v_Tmp || '00001';
    Else
      If Substr(v_Maxno, 1, 8) = v_Tmp Then
        If To_Number(Substr(v_Maxno, 9, 5)) = 99999 Then
          v_Maxno := v_Tmp || '00001';
        Else
          v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 9, 5)) + 1, '00000'));
        End If;
      Else
        v_Maxno := v_Tmp || '00001';
      End If;
    End If;
    Update 号码控制表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In;
  
    v_No := v_Maxno;
  
  Elsif 序号_In = 123 Then
  
    Select Nvl(Max(参数值), 0) Into n_Mod From 影像流程参数 Where 科室id = 科室id_In And 参数名 = '检查号生成方式';
    Select Nvl(Max(参数值), 1)
    Into n_Checkmaxno
    From 影像流程参数
    Where 科室id = 科室id_In And 参数名 = '提取实际最大号码';
  
    If n_Mod = 1 Then
      --影像检查号按科室递增
      --从号码表提取最大号码
      Select Nvl(Max(最大号码), 0) Into n_Maxno From 科室号码表 Where 项目序号 = 序号_In And 科室id = 科室id_In;
      If n_Maxno = 0 Then
        --没有记录，则自动增加科室号码表
        Insert Into 科室号码表 (项目序号, 科室id, 编号, 最大号码) Values (序号_In, 科室id_In, 'A', '1');
      End If;
    
      --提取实际最大号码,如果有实际最大号码，这个号码会比号码表中的大10
      If n_Checkmaxno = 1 Then
      
        Select Nvl(Max(检查号), 0) + 1 Into n_Amt From 影像检查记录 Where 执行科室id = 科室id_In And 检查号 >= n_Maxno;
      
        If n_Amt > n_Maxno Then
          n_Maxno := n_Amt;
        End If;
      Else
        -- 如果不提取实际最大号码，加上10
        n_Maxno := n_Maxno + 10 + 1;
      End If;
    
      -- 回填最大号码
      If (n_Checkmaxno = 0 And 号码个数_In = n_Maxno) Or n_Checkmaxno = 1 Then
        Update 科室号码表
        Set 最大号码 = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1)
        Where 项目序号 = 序号_In And 科室id = 科室id_In;
      End If;
    Else
      --影像检查号按类别递增
      Select Nvl(Max(最大号码), 0) Into n_Maxno From 影像检查类别 Where 编码 = v_Tag;
      --提取实际最大号码，如果有实际最大号码，这个号码会比号码表中的大10
      If n_Checkmaxno = 1 Then
      
        Select Nvl(Max(检查号), 0) + 1 Into n_Amt From 影像检查记录 Where 影像类别 = v_Tag And 检查号 >= n_Maxno;
      
        If n_Amt > n_Maxno Then
          n_Maxno := n_Amt;
        End If;
      Else
        -- 如果不提取实际最大号码，加上10
        n_Maxno := n_Maxno + 10 + 1;
      End If;
    
      -- 回填最大号码
      If (n_Checkmaxno = 0 And 号码个数_In = n_Maxno) Or n_Checkmaxno = 1 Then
        Update 影像检查类别 Set 最大号码 = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1) Where 编码 = v_Tag;
      End If;
    End If;
    v_No := To_Char(n_Maxno);
  
  Elsif 序号_In = 124 Then
    ----------------------------------------------------------------------------------------------------------------------------
    --体检健康号
  
    Begin
      Select Nvl(编号规则, 0) Into n_Mod From 号码控制表 Where 项目序号 = 序号_In;
    Exception
      When Others Then
        v_Error := '号码控制表中不存在序号为' || 序号_In || '的号码';
        Raise Err_Custom;
    End;
  
    If n_Mod = 0 Then
      --顺序编号
      Select Nvl(Zl_To_Number(最大号码), 0) + 号码个数_In
      Into v_Maxno
      From 号码控制表
      Where 项目序号 = 序号_In
      For Update;
      Update 号码控制表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In;
    Else
      --前缀+顺序号
      Update 号码编号表 Set 最大号码 = 最大号码 Where 项目序号 = 序号_In And 号码前缀 = v_Tag;
      If Sql%RowCount = 0 Then
        Insert Into 号码编号表
          (项目序号, 号码前缀, 日期, 最大号码)
          Select 序号_In, v_Tag, Sysdate, '' From Dual;
      End If;
    
      Select Nvl(Zl_To_Number(最大号码), 0) + 号码个数_In
      Into n_Maxno
      From 号码编号表 A
      Where a.项目序号 = 序号_In And a.号码前缀 = v_Tag;
    
      If Substr(n_Maxno, 1, Length(v_Tag)) <> v_Tag Then
        --进位了
        Select Nvl(Zl_To_Number(Substr(最大号码, Length(v_Tag) + 1)), 0) + 号码个数_In
        Into n_Maxno
        From 号码编号表 A
        Where a.项目序号 = 序号_In And a.号码前缀 = v_Tag;
        Select v_Tag || n_Maxno Into v_Maxno From Dual;
      Else
        If n_Maxno = 号码个数_In Then
          v_Maxno := v_Tag || To_Char(号码个数_In);
        Else
          v_Maxno := n_Maxno;
        End If;
      End If;
    
      Update 号码编号表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In And 号码前缀 = v_Tag;
    End If;
    v_No := v_Maxno;
  
  Elsif 序号_In = 125 Then
    --这里是为了减少号码控制表的锁定时间
    If n_Mod = 0 Then
      Begin
        Select 试管编码
        Into v_试管编码
        From 病人医嘱记录 A, 诊疗项目目录 B
        Where a.诊疗项目id = b.Id And a.Id = 科室id_In And 试管编码 Is Not Null;
      Exception
        When Others Then
          v_Error := '没有找到检验项目对应的管码！';
          Raise Err_Custom;
      End;
    Else
      Begin
        Select 试管编码, c.编码
        Into v_试管编码, v_编码
        From 病人医嘱记录 A, 诊疗项目目录 B, 诊疗检验类型 C
        Where a.诊疗项目id = b.Id And a.Id = 科室id_In And b.操作类型 = c.名称 And 试管编码 Is Not Null;
      
      Exception
        When Others Then
          v_Error := '没有找到检验项目对应的管码和编码！';
          Raise Err_Custom;
      End;
    End If;
  
    Select Nvl(最大号码, '0'), Nvl(编号规则, 0)
    Into v_Maxno, n_Mod
    From 号码控制表
    Where 项目序号 = 序号_In
    For Update;
  
    If n_Mod = 0 Then
      --按医嘱生成
    
      v_医嘱 := 科室id_In;
      If Length(v_医嘱) > 12 Then
        v_医嘱 := Substr(v_医嘱, Length(v_医嘱) - 11);
      Else
        v_医嘱 := LPad(v_医嘱, 12, '0');
      End If;
      if Length(ltrim(v_医嘱,'0'))>8 then
         Select v_试管编码 || LPad(Substr(v_医嘱, Length(v_医嘱) - (13 - Length(v_试管编码) - 2)), (13 - Length(v_试管编码)), '0')
          Into v_生成条码
         From Dual;
      else
         Select v_试管编码 || LPad(Substr(v_医嘱, Length(v_医嘱) - (12 - Length(v_试管编码) - 2)), (12 - Length(v_试管编码)), '0')
         Into v_生成条码
         From Dual;
      end if;
      v_No := v_生成条码;
    Else
      --按"小组编号（1位）+管码(2位)+日期(6位)+顺序号(3)位"生成条码
      Begin
        Select 最大号码
        Into v_Maxno
        From 号码编号表
        Where 项目序号 = 序号_In And 号码前缀 = v_编码 || v_试管编码 And Trunc(日期) = Trunc(Sysdate)
        For Update;
        v_Maxno := v_Maxno + 1;
        If Length(v_Maxno) <= 3 Then
          v_Maxno := LPad(v_Maxno, 3, '0');
        End If;
        v_No := v_编码 || v_试管编码 || To_Char(Trunc(Sysdate), 'yymmdd') || v_Maxno;
        Update 号码编号表
        Set 日期 = Trunc(Sysdate), 最大号码 = v_Maxno
        Where 号码前缀 = v_编码 || v_试管编码 And Trunc(日期) = Trunc(Sysdate);
      Exception
        When Others Then
          Update 号码编号表 Set 日期 = Trunc(Sysdate), 最大号码 = 1 Where 号码前缀 = v_编码 || v_试管编码;
          If Sql%RowCount = 0 Then
            Insert Into 号码编号表
              (项目序号, 号码前缀, 日期, 最大号码)
            Values
              (序号_In, v_编码 || v_试管编码, Trunc(Sysdate), 1);
          End If;
          v_No := v_编码 || v_试管编码 || To_Char(Trunc(Sysdate), 'yymmdd') || '001';
      End;
    End If;
    --卫材条码序号
  Elsif 序号_In = 126 Then
    --12位顺序号
    Select Nvl(最大号码, '0') Into v_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
  
    If v_Maxno = '0' Then
      v_Maxno := '000000000001';
    Else
      v_Maxno := Trim(To_Char(To_Number(v_Maxno) + 1, '000000000000'));
    End If;
    Update 号码控制表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In;
  
    v_No := v_Maxno;
  Elsif 序号_In = 135 Then
    --药品卫材调价
    Begin
      Select Nvl(编号规则, 0) Into n_Mod From 号码控制表 Where 项目序号 = 序号_In;
    Exception
      When Others Then
        v_Error := '号码控制表中不存在序号为' || 序号_In || '的号码';
        Raise Err_Custom;
    End;
    --1.年月(YYYYMM)+顺序号(0000)
    v_Tmp := To_Char(Sysdate, 'YYYYMM');
    Select Nvl(最大号码, '0') Into v_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
  
    If v_Maxno = '0' Then
      v_Maxno := v_Tmp || '0001';
    Else
      If Substr(v_Maxno, 1, 6) = v_Tmp Then
        v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 7, 4)) + 1, '0000'));
      Else
        v_Maxno := v_Tmp || '0001';
      End If;
    End If;
    Update 号码控制表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In;
    v_No := v_Maxno;
  Elsif 序号_In = 136 Then
      --YYMMDD+5位顺序号(00000)
      Select To_Char(sysdate, 'yymmdd') Into v_Tmp From Dual;
      Select Nvl(最大号码, '0') Into v_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
    
      If v_Maxno = '0' Then
        v_Maxno := v_Tmp || '00001';
      Else
        If Substr(v_Maxno, 1, 6) = v_Tmp Then
          If To_Number(Substr(v_Maxno, 7, 5)) = 99999 Then
            v_Maxno := v_Tmp || '00001';
          Else
            v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 7, 5)) + 1, '00000'));
          End If;
        Else
          v_Maxno := v_Tmp || '00001';
        End If;
      End If;
      Update 号码控制表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In;
      v_No := v_Maxno;
  Elsif 序号_In = 131 Then
    --体检报到号
    Select Nvl(最大号码, '0') Into v_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
    v_Maxno := To_Char(To_Number(v_Maxno) + 1);
    Update 号码控制表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In;
    v_No := v_Maxno;
    --其它单据号
  Else
    Begin
      Select Nvl(编号规则, 0) Into n_Mod From 号码控制表 Where 项目序号 = 序号_In;
    Exception
      When Others Then
        v_Error := '号码控制表中不存在序号为' || 序号_In || '的号码';
        Raise Err_Custom;
    End;
  
    --如果规则是按科室编号顺序产生，但又未设置科室编码，则采取按年顺序编号
    If n_Mod = 2 And 序号_In <> 122 Then
      Begin
        Select 编号 Into v_Deptno From 科室号码表 Where 科室id = 科室id_In And 项目序号 = 序号_In;
      Exception
        When Others Then
          Null;
      End;
      If v_Deptno Is Null Then
        n_Mod := 0;
      End If;
    End If;
  
    Select Decode(Sign(Intyear - 10), -1, To_Char(Intyear, '9'), Chr(55 + Intyear))
    Into v_Year
    From (Select To_Number(To_Char(Sysdate, 'yyyy'), '9999') - 1990 As Intyear From Dual);
  
    If n_Mod = 0 Then
      --0.按年顺序编号
      Select Nvl(最大号码, '0') Into v_No From 号码控制表 Where 项目序号 = 序号_In For Update;
    
      --第2位增长方式:0-9,A-Z
      Select Bit1 || Decode(Sign(Ascii(Bit2) - Ascii('9')), -1, LPad(To_Number(Bit28) + 1, 7, '0'),
                             Decode(Bit38, '999999', Decode(Bit2, '9', 'A', Chr(Ascii(Bit2) + 1)) || '000000',
                                     Bit2 || LPad(To_Number(Bit38) + 1, 6, '0')))
      Into v_No
      From (Select Substr(Maxno, 1, 1) As Bit1, Substr(Maxno, 2, 1) As Bit2, Substr(Maxno, 2) As Bit28,
                    Substr(Maxno, 3) As Bit38
             From (Select Decode(v_No, '0', v_Year || '0000000',
                                   Decode(Sign(Ascii(Substr(v_No, 1, 1)) - Ascii(v_Year)), -1, v_Year || '0000000', v_No)) As Maxno
                    From Dual));
    
      Update 号码控制表 Set 最大号码 = v_No Where 项目序号 = 序号_In;
    
    Elsif n_Mod = 1 Then
      --1.按年+日顺序编号:YDDD0000
      Select Nvl(最大号码, '0') Into v_No From 号码控制表 Where 项目序号 = 序号_In For Update;
      Select v_Year || LPad(Trunc(Sysdate - Trunc(Sysdate, 'YYYY') + 1, 0), 3, '0') || '0000' Into v_Maxno From Dual;
      If v_No < v_Maxno Then
        v_No := v_Maxno;
      End If;
      v_No := Substr(v_No, 1, 4) || LPad(To_Number(Substr(v_No, 5, 4)) + 1, 4, '0');
      Update 号码控制表 Set 最大号码 = v_No Where 项目序号 = 序号_In;
    
    Elsif n_Mod = 2 Then
      If 序号_In = 122 Then
        --2.按科室编码+YYMMDD+3位顺序号:2201090728001
        Select Count(*) Into n_Maxno From 科室号码表 Where 项目序号 = 序号_In And 科室id = 科室id_In;
        If Nvl(n_Maxno, 0) = 0 Then
          Insert Into 科室号码表 (项目序号, 科室id, 最大号码, 编号) Values (序号_In, 科室id_In, Null, Null);
          Commit;
        End If;
      
        Select 编码 Into v_Deptno From 部门表 Where ID = 科室id_In;
        Select Nvl(最大号码, '-') Into v_No From 科室号码表 Where 项目序号 = 序号_In And 科室id = 科室id_In For Update;
        v_Tmp := To_Char(Sysdate, 'YYMMDD');
      
        If Substr(v_No, 1, Length(v_Deptno || v_Tmp)) = v_Deptno || v_Tmp Then
          v_No := v_Deptno || v_Tmp || LPad(To_Number(Substr(v_No, Length(v_Deptno || v_Tmp) + 1)) + 1, 3, '0');
        Else
          v_No := v_Deptno || v_Tmp || LPad('1', 3, '0');
        End If;
        Update 科室号码表 Set 最大号码 = v_No Where 项目序号 = 序号_In And 科室id = 科室id_In;
      
      Else
        --2.按年+科室编号+月+顺序号:YKDD0000
        Begin
          --符号-的asscii为45,用于和year比较(0的ascii为48)
          Select 编号, Nvl(最大号码, '-')
          Into v_Deptno, v_No
          From 科室号码表
          Where 项目序号 = 序号_In And Nvl(科室id, 0) = Nvl(科室id_In, 0)
          For Update;
        Exception
          When Others Then
            Null;
        End;
        If v_Deptno Is Null Then
          v_Error := '科室未设置编号，无法产生号码！';
          Raise Err_Custom;
        Else
          v_Tmp := To_Char(Sysdate, 'MM');
          Select Substr(Maxno, 1, 4) || LPad(To_Number(Substr(Maxno, 5, 4)) + 1, 4, '0')
          Into v_No
          From (Select Decode(Sign(Ascii(Substr(v_No, 1, 1)) - Ascii(v_Year)), -1, v_Year || v_Deptno || v_Tmp || '0000',
                                Decode(Sign(To_Number(Substr(v_No, 3, 2)) - To_Number(v_Tmp)), -1,
                                        v_Year || v_Deptno || v_Tmp || '0000', v_No)) As Maxno
                 From Dual);
          Update 科室号码表 Set 最大号码 = v_No Where 项目序号 = 序号_In And Nvl(科室id, 0) = Nvl(科室id_In, 0);
        
        End If;
      End If;
    Elsif n_Mod = 3 Then
    
      --按年月日+000001生成
      Select Substr(To_Char(Sysdate, 'yyyymmdd'), 3, 6) Into v_Tmp From Dual;
      Select Nvl(最大号码, '0') Into v_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
    
      If v_Maxno = '0' Then
        v_Maxno := v_Tmp || '000001';
      Else
        If Substr(v_Maxno, 1, 6) = v_Tmp Then
          v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 7, 6)) + 1, '000000'));
        Else
          v_Maxno := v_Tmp || '000001';
        End If;
      End If;
      Update 号码控制表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In;
      v_No := v_Maxno;
    
    Elsif n_Mod = 5 Then
      --1.年月(YYYYMM)+顺序号(000000)
      v_Tmp := To_Char(Sysdate, 'YYYYMM');
      Select Nvl(最大号码, '0') Into v_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
    
      If v_Maxno = '0' Then
        v_Maxno := v_Tmp || '000001';
      Else
        If Substr(v_Maxno, 1, 6) = v_Tmp Then
          v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 7, 6)) + 1, '000000'));
        Else
          v_Maxno := v_Tmp || '000001';
        End If;
      End If;
      Update 号码控制表 Set 最大号码 = v_Maxno Where 项目序号 = 序号_In;
      v_No := v_Maxno;
    Else
      v_Error := '序号为' || 序号_In || '的号码,其规则值:' || n_Mod || ',当前系统不支持！';
      Raise Err_Custom;
    End If;
  End If;

  Commit;
  Return v_No;
Exception
  When Err_Custom Then
    Rollback;
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Rollback;
    zl_ErrorCenter(SQLCode, SQLErrM);
End Nextno;
/

--92808:马政,2016-01-22,卫材单据打印控制
Create Or Replace Procedure Zl_药品收发主表_Insert
(
  No_In     In 药品收发主表.No%Type,
  单据_In   In 药品收发主表.单据%Type,
  库房id_In In 药品收发主表.库房id%Type
) Is
  n_Count Number;
  n_Id    药品收发主表.Id%Type;
Begin
  n_Count := 0;
  Begin
    Select 1 Into n_Count From 药品收发主表 Where NO = No_In And 单据 = 单据_In And 库房id = 库房id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    Select 药品收发主表_Id.Nextval Into n_Id From Dual;
    Insert Into 药品收发主表 (ID, NO, 单据, 库房id, 打印状态) Values (n_Id, No_In, 单据_In, 库房id_In, 1);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--93146:刘尔旋,2016-01-27,获取号源剩余数量问题
Create Or Replace Procedure Zl_Third_Docarrange
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:医生排班计划
  --入参:Xml_In:
  --<IN>
  --   <YSID>870</YSID>    //医生ID
  --   <KDID>870</KSID>    //科室ID
  --   <KSSJ>2014-10-29 </KSSJ>    //开始时间
  --   <CXTS>14</CXTS>    //查询天数
  --   <HZDW>支付宝</HZDW> //合作单位
  --   <HL>号类</HL>      //号类，可传多个，用逗号分隔，格式:普通,专家,...
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  --   <PBLIST>       //未返回该节点表示没有数据
  --    <PB>
  --     <RQ>2014-10-29</RQ>     //日期
  --     <SYHS>5</SYHS>    //剩余号数
  --     <SBSJ>全日</SBSJ>             //上班时间
  --     <YGS>5</YGS>    //已挂号数
  --    </PB>
  --   <PBLIST>
  --   <ERROR><MSG></MSG></ERROR> //错误情况返回
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  d_日期         Date;
  v_排班         挂号安排.周日%Type;
  n_限号数       挂号安排限制.限号数%Type;
  n_已挂数       挂号安排限制.限号数%Type;
  n_总已挂数     挂号安排限制.限号数%Type;
  n_限约数       挂号安排限制.限号数%Type;
  n_已约数       挂号安排限制.限号数%Type;
  n_剩余数       挂号安排限制.限号数%Type;
  v_上班时间     Varchar2(300);
  n_医生id       人员表.Id%Type;
  n_科室id       部门表.Id%Type;
  n_查询天数     Number(4);
  n_合作单位数量 Number(5);
  n_合约已挂数   Number(4);
  n_合约存在     Number(3);
  n_安排存在     Number(3);
  v_号码         挂号安排.号码%Type;
  n_安排id       挂号安排计划.安排id%Type;
  n_计划id       挂号安排计划.Id%Type;
  v_合作单位     挂号合作单位.名称%Type;
  n_Daycount     Number(4);
  d_开始时间     Date;
  d_原始时间     Date;
  n_禁用         Number(3);
  v_Temp         Varchar2(32767); --临时XML
  x_Templet      Xmltype; --模板XML
  v_Err_Msg      Varchar2(200);
  v_号类         Varchar2(200);
  n_Exists       Number(2);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/KSID'), Extractvalue(Value(A), 'IN/CXTS'),
         To_Date(Extractvalue(Value(A), 'IN/KSSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/HL')
  Into n_医生id, n_科室id, n_查询天数, d_开始时间, v_合作单位, v_号类
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_查询天数 := Nvl(n_查询天数, 14);
  d_原始时间 := Trunc(d_开始时间);
  d_开始时间 := Trunc(d_开始时间);
  n_Daycount := 0;
  If Nvl(n_科室id, 0) = 0 Then
    While (n_Daycount < n_查询天数) Loop
      If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
        d_开始时间 := Sysdate - n_Daycount;
      Else
        d_开始时间 := d_原始时间;
      End If;
      n_安排存在 := 0;
      v_上班时间 := Null;
      n_总已挂数 := 0;
      n_已挂数   := 0;
      n_剩余数   := 0;
      n_限号数   := 0;
      n_已约数   := 0;
      n_限约数   := 0;
      For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                          a.安排id, a.计划id, a.号码, a.号类
                   
                   From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码,
                                 Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数

                          
                          From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                        Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                        Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                        Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                 From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                 Where Ap.科室id = Bm.Id(+) And Ap.医生id = n_医生id And Ap.停用日期 Is Null And
                                       d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                       Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                       Xz.限制项目(+) = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二',
                                                           '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                  (Select Rownum
                                        From 挂号安排停用状态 Ty
                                        Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                       Not Exists
                                  (Select Rownum
                                        From 挂号安排计划 Jh
                                        Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                              d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                              Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                 Union All
                                 Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id, Jh.Id As 计划id,
                                        Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                        Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                        Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                 From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                 Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And Jh.医生id = n_医生id And
                                       d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                       Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.计划id(+) = Jh.Id And
                                       Xz.限制项目(+) = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二',
                                                           '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                  (Select Rownum
                                        From 挂号安排停用状态 Ty
                                        Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                       (Jh.生效时间, Jh.安排id) =
                                       (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                        From 挂号安排计划 Sxjh
                                        Where Sxjh.审核时间 Is Not Null And
                                              d_开始时间 + n_Daycount Between Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                              Nvl(Sxjh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.安排id = Jh.安排id
                                        Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                          Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                        病人挂号汇总 Hz, 收费价目 B
                   Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                         b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount
                   
                   ) Loop
        If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
          v_上班时间 := v_上班时间 || '+' || r_排班.排班;
          n_总已挂数 := n_总已挂数 + r_排班.已挂数;
          n_已挂数   := r_排班.已挂数;
          n_限号数   := r_排班.限号数;
          n_已约数   := r_排班.已约数;
          n_限约数   := r_排班.限约数;
          n_安排id   := Nvl(r_排班.安排id, 0);
          n_计划id   := Nvl(r_排班.计划id, 0);
          v_号码     := r_排班.号码;
          n_安排存在 := 1;
          If v_上班时间 Is Not Null Then
            If v_合作单位 Is Not Null Then
              If n_计划id <> 0 Then
                Begin
                  Select 1
                  Into n_合约存在
                  From 合作单位计划控制
                  Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约存在 := 0;
                End;
              Else
                Begin
                  Select 1
                  Into n_合约存在
                  From 合作单位安排控制
                  Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约存在 := 0;
                End;
              End If;
            End If;
          
            If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
              If n_计划id <> 0 Then
                Begin
                  Select Sum(数量)
                  Into n_合作单位数量
                  From 合作单位计划控制
                  Where 计划id = n_计划id And
                        限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                      '周四', '6', '周五', '7', '周六', Null);
                Exception
                  When Others Then
                    n_合作单位数量 := 0;
                End;
              Else
                Begin
                  Select Sum(数量)
                  Into n_合作单位数量
                  From 合作单位安排控制
                  Where 安排id = n_安排id And
                        限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                      '周四', '6', '周五', '7', '周六', Null);
                Exception
                  When Others Then
                    n_合作单位数量 := 0;
                End;
              End If;
              Begin
                Select Count(1)
                Into n_合约已挂数
                From 病人挂号记录
                Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                      Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
              Exception
                When Others Then
                  n_合约已挂数 := 0;
              End;
              If n_合作单位数量 = 0 Then
                n_合作单位数量 := Null;
              End If;
              If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
              Else
                n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
              End If;
            Else
              --合约单位
              If n_计划id <> 0 Then
                Begin
                  Select 1
                  Into n_禁用
                  From 合作单位计划控制
                  Where 计划id = n_计划id And
                        限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                      '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
                Exception
                  When Others Then
                    n_禁用 := 0;
                End;
              Else
                Begin
                  Select 1
                  Into n_禁用
                  From 合作单位安排控制
                  Where 安排id = n_安排id And
                        限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                      '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
                Exception
                  When Others Then
                    n_禁用 := 0;
                End;
              End If;
              If Nvl(n_禁用, 0) = 0 Then
                Begin
                  Select Count(1)
                  Into n_合约已挂数
                  From 病人挂号记录
                  Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                        Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_合约已挂数 := 0;
                End;
                If n_计划id <> 0 Then
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                Else
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                End If;
                If n_合作单位数量 = 0 Then
                  n_合作单位数量 := Null;
                End If;
                n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
              
              End If;
            End If;
          End If;
          n_合作单位数量 := 0;
          n_合约存在     := 0;
          n_禁用         := 0;
        End If;
      End Loop;
      v_上班时间 := Substr(v_上班时间, 2);
      If n_安排存在 = 1 Then
        v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                  '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 || '</YGS>' ||
                  '</PB>';
      End If;
      n_Daycount := n_Daycount + 1;
    End Loop;
  Else
    While (n_Daycount < n_查询天数) Loop
      If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
        d_开始时间 := Sysdate - n_Daycount;
      Else
        d_开始时间 := d_原始时间;
      End If;
      v_上班时间 := Null;
      n_总已挂数 := 0;
      n_已挂数   := 0;
      n_剩余数   := 0;
      n_限号数   := 0;
      n_已约数   := 0;
      n_限约数   := 0;
      n_安排存在 := 0;
      For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                          a.安排id, a.计划id, a.号码, a.号类
                   
                   From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码,
                                 Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数

                          
                          From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                        Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                        Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                        Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                 From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                 Where Ap.科室id = Bm.Id(+) And Ap.医生id = n_医生id And Ap.科室id = n_科室id And Ap.停用日期 Is Null And
                                       d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                       Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                       Xz.限制项目(+) = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二',
                                                           '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                  (Select Rownum
                                        From 挂号安排停用状态 Ty
                                        Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                       Not Exists
                                  (Select Rownum
                                        From 挂号安排计划 Jh
                                        Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                              d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                              Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                 Union All
                                 Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id, Jh.Id As 计划id,
                                        Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                        Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                        Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                 From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                 Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And Jh.医生id = n_医生id And
                                       Ap.科室id = n_科室id And
                                       d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                       Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.计划id(+) = Jh.Id And
                                       Xz.限制项目(+) = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二',
                                                           '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                  (Select Rownum
                                        From 挂号安排停用状态 Ty
                                        Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                       (Jh.生效时间, Jh.安排id) =
                                       (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                        From 挂号安排计划 Sxjh
                                        Where Sxjh.审核时间 Is Not Null And
                                              d_开始时间 + n_Daycount Between Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                              Nvl(Sxjh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.安排id = Jh.安排id
                                        Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                          Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                        病人挂号汇总 Hz, 收费价目 B
                   Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                         b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount
                   
                   ) Loop
        If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
          v_上班时间 := v_上班时间 || '+' || r_排班.排班;
          n_总已挂数 := n_总已挂数 + r_排班.已挂数;
          n_已挂数   := r_排班.已挂数;
          n_限号数   := r_排班.限号数;
          n_已约数   := r_排班.已约数;
          n_限约数   := r_排班.限约数;
          n_安排id   := Nvl(r_排班.安排id, 0);
          n_计划id   := Nvl(r_排班.计划id, 0);
          v_号码     := r_排班.号码;
          n_安排存在 := 1;
        
          If v_上班时间 Is Not Null Then
            If v_合作单位 Is Not Null Then
              If n_计划id <> 0 Then
                Begin
                  Select 1
                  Into n_合约存在
                  From 合作单位计划控制
                  Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约存在 := 0;
                End;
              Else
                Begin
                  Select 1
                  Into n_合约存在
                  From 合作单位安排控制
                  Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约存在 := 0;
                End;
              End If;
            End If;
          
            If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
              If n_计划id <> 0 Then
                Begin
                  Select Sum(数量)
                  Into n_合作单位数量
                  From 合作单位计划控制
                  Where 计划id = n_计划id And
                        限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                      '周四', '6', '周五', '7', '周六', Null);
                Exception
                  When Others Then
                    n_合作单位数量 := 0;
                End;
              Else
                Begin
                  Select Sum(数量)
                  Into n_合作单位数量
                  From 合作单位安排控制
                  Where 安排id = n_安排id And
                        限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                      '周四', '6', '周五', '7', '周六', Null);
                Exception
                  When Others Then
                    n_合作单位数量 := 0;
                End;
              End If;
              Begin
                Select Count(1)
                Into n_合约已挂数
                From 病人挂号记录
                Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                      Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
              Exception
                When Others Then
                  n_合约已挂数 := 0;
              End;
              If n_合作单位数量 = 0 Then
                n_合作单位数量 := Null;
              End If;
              If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
              Else
                n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
              End If;
            Else
              --合约单位
              If n_计划id <> 0 Then
                Begin
                  Select 1
                  Into n_禁用
                  From 合作单位计划控制
                  Where 计划id = n_计划id And
                        限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                      '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
                Exception
                  When Others Then
                    n_禁用 := 0;
                End;
              Else
                Begin
                  Select 1
                  Into n_禁用
                  From 合作单位安排控制
                  Where 安排id = n_安排id And
                        限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                      '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
                Exception
                  When Others Then
                    n_禁用 := 0;
                End;
              End If;
              If Nvl(n_禁用, 0) = 0 Then
                Begin
                  Select Count(1)
                  Into n_合约已挂数
                  From 病人挂号记录
                  Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                        Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_合约已挂数 := 0;
                End;
                If n_计划id <> 0 Then
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                Else
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                End If;
                If n_合作单位数量 = 0 Then
                  n_合作单位数量 := Null;
                End If;
                n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
              
              End If;
            End If;
          End If;
          n_合作单位数量 := 0;
          n_合约存在     := 0;
          n_禁用         := 0;
        End If;
      End Loop;
      v_上班时间 := Substr(v_上班时间, 2);
      If n_安排存在 = 1 Then
        v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                  '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 || '</YGS>' ||
                  '</PB>';
      End If;
      n_Daycount := n_Daycount + 1;
    End Loop;
  End If;
  v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Docarrange;
/

--93052:刘尔旋,2016-02-02,服务窗结帐问题处理
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
                          Where a.记录状态 = 1 And a.病人id = c.病人id And a.病人id = b.病人id And c.主页id = n_主页id And a.病人id = n_病人id And
                                d.结帐id = a.Id And c.入院科室id = e.Id(+) And Exists
                           (Select 1 From 病人预交记录 Where 结帐id = a.Id And 结算方式 = v_结算卡类别 And 主页id = n_主页id)
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
      If Nvl(n_病人余额,0) - Nvl(n_结帐金额,0) > 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(Nvl(n_病人余额,0) - Nvl(n_结帐金额,0)) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getsettlement;
/

--93052:刘尔旋,2016-01-25,支付宝结帐校对标志问题
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
  Update 病人预交记录 Set 校对标志 = 0 Where 结帐id = n_结帐id And Nvl(校对标志, 0) <> 0;
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

--90143:李南春,2016-01-25,卡费生成划价单时生成发卡记录
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
  更新交款余额_In  Number := 0,--是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况。
  摘要_In       住院费用记录.摘要%Type := Null
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
       结帐金额, 缴款组id, 结论,摘要)
    Values
      (v_费用id, 5, 1, 单据号_In, 医疗卡号_In, 1, 病人id_In, Decode(主页id_In, 0, Null, 主页id_In),
       Decode(病人病区id_In, 0, Null, 病人病区id_In), Decode(病人科室id_In, 0, Null, 病人科室id_In), Decode(标识号_In, 0, Null, 标识号_In),
       姓名_In, 性别_In, 年龄_In, 费别_In, Decode(结算方式_In, Null, 1, 0), 3, 加班标志_In, 开单部门id_In, 操作员姓名_In, 操作员编号_In, 操作员姓名_In,
       发卡时间_In, 发卡时间_In, 收费细目id_In, 收费类别_In, 计算单位_In, 1, 1, 医疗卡号_In, 发卡类型_In, 执行部门id_In, 收入项目id_In, 收据费目_In, 标准单价_In,
       应收金额_In, 实收金额_In, v_结帐id, Decode(结算方式_In, Null, Null, 实收金额_In), n_组id, 卡类别id_In, 摘要_In);
  
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

Create Or Replace Procedure Zl_医疗卡记录_Delete
(
  单据号_In     住院费用记录.No%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type
) As
  Cursor c_Cardinfo Is
    Select a.Id As 费用id, Nvl(a.记帐费用, 0) As 记帐, a.结帐id, a.实际票号, a.病人id, Nvl(a.主页id, 0) As 主页id,
           Nvl(a.病人病区id, 0) As 病人病区id, Nvl(a.病人科室id, 0) As 病人科室id, Nvl(a.开单部门id, 0) As 开单部门id,
           Nvl(a.执行部门id, 0) As 执行部门id, a.收入项目id, a.实收金额, b.结算方式, b.冲预交, b.卡类别id, b.卡号, b.结算卡序号, b.结算序号, a.结论,
           b.Id As 预交id, a.摘要
    From 住院费用记录 A, 病人预交记录 B
    Where a.记录性质 = 5 And a.记录状态 = 1 And a.No = 单据号_In And a.结帐id = b.结帐id(+);
  r_Cardrow c_Cardinfo%RowType;

  v_费用id   住院费用记录.Id%Type;
  v_结帐id   住院费用记录.结帐id%Type;
  v_打印id   票据打印内容.Id%Type;
  n_返回值   病人余额.费用余额%Type;
  n_卡类别id Number(18);
  v_划价状态 门诊费用记录.记录状态%Type;

  v_Date Date;

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_组id 财务缴款分组.Id%Type;

Begin
  Open c_Cardinfo;
  Fetch c_Cardinfo
    Into r_Cardrow;
  n_组id := Zl_Get组id(操作员姓名_In);

  --首先判断要退卡的记录是否存在
  If c_Cardinfo%RowCount = 0 Then
    Close c_Cardinfo;
    v_Err_Msg := '[ZLSOFT]没有发现要退卡的记录,该记录可能已经退除！[ZLSOFT]';
    Raise Err_Item;
  Else
    Select Sysdate Into v_Date From Dual;
    Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
  
    If r_Cardrow.记帐 = 0 Then
      Select 病人结帐记录_Id.Nextval Into v_结帐id From Dual;
    End If;
  
    --退除就诊卡费用记录
    Insert Into 住院费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次, 加班标志,
       附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额,
       缴款组id, 结论, 摘要)
      Select v_费用id, NO, 实际票号, 记录性质, 2, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
             -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号_In, 操作员姓名_In,
             发生时间, v_Date, v_结帐id, Decode(v_结帐id, Null, Null, -结帐金额), n_组id, 结论, 摘要
      From 住院费用记录
      Where ID = r_Cardrow.费用id;
  
    Update 住院费用记录 Set 记录状态 = 3 Where ID = r_Cardrow.费用id;
    
    --处理发卡划价单，如果划价还未收费，直接删除
    Begin
      Select Nvl(记录状态,-1) into v_划价状态 From 门诊费用记录 Where 病人ID=r_Cardrow.病人id And 记录性质=1 And NO=r_Cardrow.摘要;
    Exception
      When Others Then
        v_划价状态 := -1;
    End;
    if v_划价状态=0 then
      Zl_门诊划价记录_Delete(r_Cardrow.摘要);
    end if;
    
    --预交款里现收的结算金额
    If r_Cardrow.记帐 = 0 Then
      Insert Into 病人预交记录
        (ID, NO, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, v_Date, 操作员编号_In, 操作员姓名_In, -冲预交, v_结帐id,
               n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 5
        From 病人预交记录
        Where 记录性质 = 5 And 记录状态 = 1 And 结帐id = r_Cardrow.结帐id;
    
      Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 5 And 记录状态 = 1 And 结帐id = r_Cardrow.结帐id;
    
      --Zl_病人卡结算记录_Strike(结帐id_In In Varchar2,  预交id_In 病人预交记录.ID%Type := -1
      Zl_病人卡结算记录_Strike(r_Cardrow.结帐id, r_Cardrow.预交id);
    
    End If;
  
    --退卡收回票据
    Begin
      Select ID
      Into v_打印id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 5 And b.No = 单据号_In
             Order By a.使用时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    If v_打印id Is Not Null Then
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 回收次数, 打印id, 使用时间, 使用人)
        Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 回收次数, 打印id, v_Date, 操作员姓名_In
        From 票据使用明细
        Where 打印id = v_打印id And 票种 = 5 And 性质 = 1;
    End If;
  
    n_卡类别id := To_Number(Nvl(r_Cardrow.结论, '0'));
    If n_卡类别id = 0 Then
      --取就诊卡
      Select ID Into n_卡类别id From 医疗卡类别 Where 名称 = '就诊卡' And Nvl(是否固定, 0) = 1;
    End If;
  
    --处理相关的变动信息
    --Zl_医疗卡变动_Insert (变动类型_In/病人id_In ,卡类别id_In, 原卡号_In, 医疗卡号_In, 变动原因_In, 密码_In, 操作员姓名_In, 变动时间_In
    --Ic卡号_In, 挂失方式_In)
    --变动类型_In:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
    Zl_医疗卡变动_Insert(4, r_Cardrow.病人id, n_卡类别id, r_Cardrow.实际票号, r_Cardrow.实际票号, Null, Null, 操作员姓名_In, v_Date, Null,
                    Null);
  
    ----------------------------------------------------------------------------------------------------------------------------------------
  
    --相关汇总表的处理
    If r_Cardrow.记帐 = 1 Then
      --汇总'病人余额'
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + (-1 * r_Cardrow.实收金额)
      Where 性质 = 1 And 病人id = r_Cardrow.病人id And Nvl(类型, 2) = Decode(Nvl(r_Cardrow.主页id, 0), 0, 1, 2)
      Returning 费用余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (r_Cardrow.病人id, 1, Decode(Nvl(r_Cardrow.主页id, 0), 0, 1, 2), 0, -1 * r_Cardrow.实收金额);
        n_返回值 := -1 * r_Cardrow.实收金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额 Where 性质 = 1 And 病人id = r_Cardrow.病人id And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --汇总'病人未结费用'
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + (-1 * r_Cardrow.实收金额)
      Where 病人id = r_Cardrow.病人id And Nvl(主页id, 0) = r_Cardrow.主页id And Nvl(病人病区id, 0) = r_Cardrow.病人病区id And
            Nvl(病人科室id, 0) = r_Cardrow.病人科室id And Nvl(开单部门id, 0) = r_Cardrow.开单部门id And
            Nvl(执行部门id, 0) = r_Cardrow.执行部门id And 收入项目id + 0 = r_Cardrow.收入项目id And 来源途径 = 3;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (r_Cardrow.病人id, Decode(r_Cardrow.主页id, 0, Null, r_Cardrow.主页id),
           Decode(r_Cardrow.病人病区id, 0, Null, r_Cardrow.病人病区id), Decode(r_Cardrow.病人科室id, 0, Null, r_Cardrow.病人科室id),
           Decode(r_Cardrow.开单部门id, 0, Null, r_Cardrow.开单部门id), Decode(r_Cardrow.执行部门id, 0, Null, r_Cardrow.执行部门id),
           r_Cardrow.收入项目id, 3, -1 * r_Cardrow.实收金额);
      End If;
    
    Elsif r_Cardrow.结算方式 Is Not Null Then
      --汇总"人员缴款余额"
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + (-1 * r_Cardrow.冲预交)
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Cardrow.结算方式
      Returning 余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Cardrow.结算方式, 1, -1 * r_Cardrow.冲预交);
        n_返回值 := -1 * r_Cardrow.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = r_Cardrow.结算方式 And Nvl(余额, 0) = 0;
      End If;
    End If;
  
    Close c_Cardinfo;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20999, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_医疗卡记录_Delete;
/

--94132:刘尔旋,2016-03-14,支付宝医生换诊号源问题
--93006:刘尔旋,2016-01-22,禁止预约号源处理
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
    v_Sql      := v_Sql || 'And Jh.医生id = :' || n_Curcount || ' ';
    n_Curcount := n_Curcount + 1;
  End If;
  If Nvl(v_医生姓名, '_') <> '_' Then
    v_Sql      := v_Sql || 'And Jh.医生姓名 = :' || n_Curcount || ' ';
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
	--限约数=0的预约禁止
    If Trunc(d_日期) <> Trunc(Sysdate) Then
      If r_限约数 = 0 Then
        n_禁用 := 1;
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

--92949:马政,2016-01-22,人员所属部门太多时错误修正
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

--91316:张德婷,2016-03-10,修正自动排批问题
--91954:张德婷,2016-02-29,自动排批相关
CREATE OR REPLACE Procedure Zl_输液配药记录_自动排批
(
 病人id_In In number,
 科室id_In In number,
 部门id_In In number,
 执行日期_In in date
) Is
v_批次串 varchar2(500);
n_总量 number(5);
v_配药id  varchar2(500);
v_batch varchar2(20);
v_Fields varchar2(100);
v_Tansid varchar(18);
n_科室id number(18);
v_批次   varchar2(10);
v_Id     varchar2(200);
n_自动排批模式  number(1);
Begin
  n_自动排批模式:=Zl_To_Number(Nvl(zl_GetSysParameter('自动排批时输液单的批次只往后面批次变动', 1345), 0));
  --该科室各个批次对应的容量
  select max(批次) into v_批次 from 配药工作批次 where 配置中心id=部门id_In and 药品类型 is  null;
  for R_Batch in (select B.批次 配药批次,A.容量,A.科室id from 科室容量设置 A,配药工作批次 B where A.配置中心ID=B.配置中心ID And A.配药批次=(B.批次 || '#') and (A.科室id=科室id_In or A.科室ID=0) and A.配置中心id=部门id_In order by A.科室id desc, A.配药批次 asc) loop

    n_科室id:=R_Batch.科室id;

  --该病人按批次排序，执行时间，优先级排序,各个批次现有的容量，优先级在产生输液配药记录的时候写入
    n_总量:=0;
    for r_item in (Select a.Id 配药id, d.单量,A.瓶签号,A.配药批次
    From 输液配药记录 A, 病人医嘱记录 B, 输液配药内容 C, 药品收发记录 D, 药品规格 E, 药品特性 F,配药工作批次 G
    Where a.Id = c.记录id And c.收发id = d.Id And d.药品id = e.药品id And e.药名id = f.药名id And a.部门id = 部门id_In and G.配置中心ID=a.部门id And
          a.医嘱id = b.Id And b.病人id = 病人id_In And A.配药批次=G.批次 And G.批次<>0 and G.药品类型 is null And f.溶媒 = 1 And a.执行时间 Between Trunc(执行日期_In ) And
          Trunc(执行日期_In+1) - 1 / 24 / 60 / 60 and  A.配药批次<=decode(n_自动排批模式,1,R_Batch.配药批次,100)
    Order By a.优先级,a.配药批次, a.执行时间, d.单量 desc) loop

      if instr(','|| v_配药id,','|| r_item.配药id || ',',1)<1then
        --当该配药id首次循环的时候对其单量进行累计
        v_配药id:=v_配药id || r_item.配药id || ',';
        n_总量:=n_总量+r_item.单量;
        v_批次串:=v_批次串 || r_item.配药id || ',' || R_Batch.配药批次 || '|';
      elsif instr('|'|| v_批次串,'|'|| r_item.配药id || ',' || R_Batch.配药批次 || '|',1)>0 and n_总量<>0 then
        --当该配药id及批次出现过，则累计该单量
        n_总量:=n_总量+r_item.单量;
      end if;

      if n_总量>=R_Batch.容量 then
        exit;
      end if;

    end loop;


  end loop;

  --如果该科室设置了单独的容量信息则不考虑所有科室的模式
  for r_item in (Select a.Id 配药id, d.单量,A.瓶签号
  From 输液配药记录 A, 病人医嘱记录 B, 输液配药内容 C, 药品收发记录 D, 药品规格 E, 药品特性 F,配药工作批次 G
  Where a.Id = c.记录id And c.收发id = d.Id And d.药品id = e.药品id And e.药名id = f.药名id And a.部门id = 部门id_In  And
        a.医嘱id = b.Id And b.病人id = 病人id_In And A.配药批次=G.批次 And G.批次<>0 and G.配置中心ID=a.部门id and G.药品类型 is null And f.溶媒 = 1 And a.执行时间 Between Trunc(执行日期_In) And
        Trunc(执行日期_In+1) - 1 / 24 / 60 / 60
  Order By a.配药批次,a.优先级,a.执行时间, d.单量) loop
    if instr(','|| v_配药id,','|| r_item.配药id || ',',1)<1then
      --没有进行自动分批的输液单直接在最后一个批次
      v_配药id:=v_配药id || r_item.配药id || ',';
      v_批次串:=v_批次串 || r_item.配药id || ',' || v_批次 || '|';
    end if;
  end loop;


  --根据自动调批修正数据
  while v_批次串 is not null loop
    --分解单据ID串
    v_Fields := Substr(v_批次串, 1, Instr(v_批次串, '|') - 1);
    v_Tansid := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    v_batch   := Substr(v_Fields, Instr(v_Fields, ',') + 1);

    v_批次串 := Replace('|' || v_批次串, '|' || v_Fields || '|');

    update 输液配药记录 set 配药批次=v_batch where id=v_Tansid;
  end loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_自动排批;
/

--93011:刘硕,2016-01-22,修正病案接收管理中接收病案后在电子病案审查中不能看见
Create Or Replace Procedure Zl_病案接收记录_Insert
(
  病人id_In   In 病案接收记录.病人id%Type,
  主页id_In   In 病案接收记录.主页id%Type,
  运送人_In   In 病案接收记录.运送人%Type,
  接收人_In   In 病案接收记录.接收人%Type,
  接收时间_In In 病案接收记录.接收时间%Type,
  记录时间_In In 病案接收记录.记录时间%Type
) Is
  n_Id      病案接收记录.Id%Type;
  n_Count   Number := 0;
  v_Err_Msg Varchar2(2000);
  Err_Item Exception;
  n_提交id Number(18);
Begin
  Select Count(病人id) Into n_Count From 病案接收记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
  If Nvl(n_Count, 0) > 0 Then
    Update 病案接收记录
    Set 病人id = 病人id_In, 主页id = 主页id_In, 运送人 = 运送人_In, 接收人 = 接收人_In, 接收时间 = 接收时间_In, 记录时间 = 记录时间_In
    Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Select 病案接收记录_Id.Nextval Into n_Id From Dual;
    Insert Into 病案接收记录
      (ID, 病人id, 主页id, 运送人, 接收人, 接收时间, 记录时间)
    Values
      (n_Id, 病人id_In, 主页id_In, 运送人_In, 接收人_In, 接收时间_In, 记录时间_In);
  End If;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]病案接收管理增加及修改接收记录失败，请核查再试！[ZLSOFT]';
    Raise Err_Item;
  End If;

  Begin
    --电子病案接收需要，因为可能独立安装，所以只能动态执行
    Execute Immediate 'Select ID From 病案提交记录 Where 病人ID=:1 and 主页id=:2'
      Into n_提交id
      Using 病人id_In, 主页id_In;
    Execute Immediate 'CALL ZL_病案提交记录_RECEIVE(:1,:2)'
      Using n_提交id, 接收人_In;
  Exception
    When Others Then
      Null;
  End;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案接收记录_Insert;
/

--92918:刘尔旋,2016-01-21,就诊变动记录安装脚本修正
Create Or Replace Procedure Zl_就诊变动记录_Insert
(
  挂号单号_In   In 病人挂号记录.No%Type,
  操作类别_In   In 就诊变动记录.类别%Type,
  变动原因_In   In 就诊变动记录.变动原因%Type,
  操作员姓名_In In 就诊变动记录.操作员姓名%Type,
  操作员编号_In In 就诊变动记录.操作员编号%Type,
  号码_In       In 就诊变动记录.现号码%Type := Null,
  科室id_In     In 就诊变动记录.现科室id%Type := Null,
  项目id_In     In 就诊变动记录.现项目id%Type := Null,
  医生id_In     In 就诊变动记录.现医生id%Type := Null,
  医生姓名_In   In 就诊变动记录.现医生姓名%Type := Null,
  诊室_In       In 就诊变动记录.现诊室%Type := Null,
  号序_In       In 就诊变动记录.现号序%Type := Null,
  预约时间_In   In 就诊变动记录.现预约时间%Type := Null
) Is
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  -----------------------------------------------------------
  --说明：挂号换号前调用，记录换号变动，号码_IN以后的值若不传则代表无变动
  --      操作类别_IN = 1-批量换号;2-分诊换号;3-强制续诊换号
  -----------------------------------------------------------
Begin
  Insert Into 就诊变动记录
    (ID, 类别, 挂号单, 病人id, 变动原因, 原号码, 现号码, 原科室id, 现科室id, 原项目id, 现项目id, 原医生id, 现医生id, 原医生姓名, 现医生姓名, 原诊室, 现诊室, 原号序, 现号序,
     原预约时间, 现预约时间, 登记时间, 操作员姓名, 操作员编号)
    Select 就诊变动记录_Id.Nextval, 操作类别_In, 挂号单号_In, a.病人id, 变动原因_In, a.号别, Nvl(号码_In, a.号别), a.执行部门id,
           Nvl(科室id_In, a.执行部门id), b.收费细目id, Nvl(项目id_In, b.收费细目id), c.Id, 医生id_In, a.执行人, 医生姓名_In, a.诊室, 诊室_In, a.号序,
           Nvl(号序_In, a.号序), a.预约时间, Nvl(预约时间_In, a.预约时间), Sysdate, 操作员姓名_In, 操作员编号_In
    From 病人挂号记录 A, 门诊费用记录 B, 人员表 C
    Where a.No = 挂号单号_In And a.记录状态 = 1 And a.执行人 = c.姓名(+) And b.No = a.No And b.记录性质 = 4 And b.序号 = 1 And Rownum < 2;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_就诊变动记录_Insert;
/

--92984:马政,2016-01-21,部门管理标准版与其他系统过程偏差处理
Create Or Replace Procedure Zl_部门表_Stop(Id_In In 部门表.ID%Type) Is
Begin
  Update 部门表 Set 撤档时间 = Sysdate Where ID = Id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_部门表_Stop;
/

--92984:马政,2016-01-21,部门管理标准版与其他系统过程偏差处理
Create Or Replace Procedure Zl_部门表_Reuse(Id_In In 部门表.ID%Type) Is
Begin
  Update 部门表 Set 撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Where ID = Id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_部门表_Reuse;
/

--91316:张德婷,2016-01-21,输液配置中心批次
CREATE OR REPLACE Procedure Zl_配药工作批次_Save
( 
  批次id_In In Varchar2, --工作批次串:批次1,配药时间1,给药时间1,打包1,启用1,颜色1,药品类型1|.... 
  配置中心ID_In In 配药工作批次.配置中心ID%type 
) Is 
  v_批次     Varchar2(20); 
  v_配药时间 Varchar2(20); 
  v_给药时间 Varchar2(20); 
  n_打包     Number(1); 
  v_启用     Number(1); 
  v_颜色     number(18); 
  v_Field    Varchar2(500); 
  v_Tmp      Varchar2(500); 
  v_药品类型 varchar2(50);
Begin 
  If 批次id_In Is Null Then 
    v_Tmp := Null; 
  Else 
    v_Tmp := 批次id_In || '|'; 
  End If; 
 
  Delete From 配药工作批次 where 配置中心ID=配置中心ID_In; 
 
  While v_Tmp Is Not Null Loop 
    --分解 
    v_Field := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1); 
    v_Tmp   := Replace('|' || v_Tmp, '|' || v_Field || '|'); 
 
    v_批次  := Substr(v_Field, 1, Instr(v_Field, ',') - 1); 
    v_Field := Substr(v_Field, Instr(v_Field, ',') + 1); 
 
    v_配药时间 := Substr(v_Field, 1, Instr(v_Field, ',') - 1); 
    v_Field    := Substr(v_Field, Instr(v_Field, ',') + 1); 
 
    v_给药时间 := Substr(v_Field, 1, Instr(v_Field, ',') - 1); 
    v_Field    := Substr(v_Field, Instr(v_Field, ',') + 1); 
 
    n_打包 := Substr(v_Field, 1, Instr(v_Field, ',') - 1); 
    v_Field    := Substr(v_Field, Instr(v_Field, ',') + 1); 
 
    v_启用 := Substr(v_Field,1, Instr(v_Field, ',')- 1); 
    v_Field    := Substr(v_Field, Instr(v_Field, ',') + 1);
    
    v_颜色 :=Substr(v_Field,1, Instr(v_Field, ',')- 1); 
    v_药品类型 := Substr(v_Field, Instr(v_Field, ',')+ 1); 
 
    Insert Into 配药工作批次 
      (批次, 配药时间, 给药时间, 打包, 启用,颜色,药品类型,配置中心ID) 
    Values 
      (v_批次, v_配药时间, v_给药时间, n_打包, v_启用,v_颜色,v_药品类型,配置中心ID_In); 
  End Loop; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_配药工作批次_Save;
/

--91333:张德婷,2016-01-21,输液配置中心打印流水号
CREATE OR REPLACE Procedure Zl_输液配药记录_打印
( 
  配药id_In   In Varchar2, --ID串:ID1,ID2.... 
  打印时间_In In 输液配药记录.打印时间%Type 
) Is 
  v_Tansid Varchar2(20); 
  v_Tmp    Varchar2(4000); 
  n_Count  Number(5); 
  n_Row    NUmber(5);
Begin 
  
  n_Count := 0; 
  If 配药id_In Is Null Then 
    v_Tmp := Null; 
  Else 
    v_Tmp := 配药id_In || ','; 
  End If; 
  While v_Tmp Is Not Null Loop 
    --分解单据ID串 
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1); 
    v_Tmp    := Substr(v_Tmp, Instr(v_Tmp, ',') + 1); 
 
    n_Count := n_Count + 1; 
    select count(id) into n_Row from 输液配药记录 where Nvl(打印标志, 0)<>0 and 打印时间 between Trunc(打印时间_In ) And Trunc(打印时间_In+1) - 1 / 24 / 60 / 60;
    Update 输液配药记录 
    Set 打印标志 = decode(Nvl(打印标志, 0),0,n_Row+1,打印标志), 打印时间 = 打印时间_In, 打印序号 = n_Count 
    Where ID = v_Tansid; 
  End Loop; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_输液配药记录_打印;
/

--93964:梁唐彬,2016-03-07,两条相同时间变动记录无法回退问题
--92938:刘鹏飞,2016-01-21,病人变动记录撤销过程升级脚本和安装脚本不同修正
Create Or Replace Procedure Zl_病人变动记录_Undo
(
  病人id_In     病案主页.病人id%Type,
  主页id_In     病案主页.主页id%Type,
  操作员编号_In 病人变动记录.操作员编号%Type,
  操作员姓名_In 病人变动记录.操作员姓名%Type,
  数据_In       Varchar2 := Null, --a.转为住院时,清除住院号,b-检查自动记帐费用是否已结帐
  床号_In       Varchar2 := Null, --传入时表示撤销出院时安排到新的床号，原床位被占用在程序中判断
  主床位_In     Varchar2 := Null, --传入时表示撤销出院时安排到新的主床位，原床位被占用在程序中判断
  撤销方式_In   Varchar2 := Null --指明具体撤销操作，如撤销出院、转科等必须输入
) As
  -----------------------------------------------------------
  --说明：1.撤消病人最近一次的变动
  --        2.前提：当病人包床时,对其中一张床位作变动,则所有床位相应产生变动
  -----------------------------------------------------------
  --要撤消的变动记录(如果包床,可能多条)
  Cursor c_Curlog Is
    Select *
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And (终止时间 Is Null Or 终止原因 = 1)
    Order By 终止时间 Desc, 开始时间 Desc;
  r_Curlogrow c_Curlog%Rowtype;

  --撤消后要恢复的变动记录(如果包床,可能多条)
  Cursor c_Prelog
  (
    v_终止时间 病人变动记录.终止时间%Type,
    v_终止原因 病人变动记录.终止原因%Type
  ) Is
    Select *
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = v_终止时间 And 终止原因 = v_终止原因
    Order By 终止时间 Desc, 开始时间 Desc;
  r_Prelogrow c_Prelog%Rowtype;

  --获取病人原床位所住病人信息
  Cursor c_Prebed
  (
    v_病人id 病案主页.病人id%Type,
    v_主页id 病案主页.主页id%Type
  ) Is
    Select a.床号, c.出院病床, c.出院科室id, c.当前病区id
    From 病人变动记录 a, 床位状况记录 b, 病案主页 c
    Where a.病人id = 病人id_In And a.主页id = 主页id_In And 终止原因 = 4 And a.病区id = b.病区id And a.病人id = c.病人id And
          a.主页id = c.主页id And a.床号 = b.床号
    Order By 终止时间 Desc, 开始时间 Desc;
  r_Prebedrow c_Prebed%Rowtype;

  Cursor c_Prebedpati
  (
    v_出院科室id 病人变动记录.科室id%Type,
    v_病区id     病人变动记录.病区id%Type,
    v_原床号     病人变动记录.床号%Type
  ) Is
    Select a.病人id, c.主页id, a.床号, c.出院病床
    From 病人变动记录 a, 病案主页 c,
         (Select 病人id
           From 床位状况记录
           Where (科室id Is Null Or 科室id = v_出院科室id Or 共用 = 1) And 病区id = v_病区id And 床号 = v_原床号) d
    Where a.病人id = d.病人id And a.主页id = (Select 主页id From 病人信息 Where 病人id = d.病人id) And a.终止原因 = 4 And a.病人id = c.病人id And
          a.主页id = c.主页id
    Order By 终止时间 Desc, 开始时间 Desc;
  r_Prebedpati c_Prebedpati%Rowtype;

  v_开始时间 病人变动记录.开始时间%Type;
  v_开始原因 病人变动记录.开始原因%Type;
  v_终止人员 病人变动记录.终止人员%Type;

  v_Count       Number;
  v_Countcurlog Number;
  v_Countprelog Number;

  Err_Custom Exception;
  v_Error Varchar2(255);

  v_撤销方式     Varchar2(100);
  v_共享号       Zlsystems.共享号%Type;
  v_病案状态     Number(3);
  v_床号串       Varchar2(255);
  v_床号         病人变动记录.床号%Type;
  v_病人id       病人变动记录.病人id%Type;
  v_主页id       病人变动记录.主页id%Type;
  v_病区id       病人变动记录.病区id%Type;
  v_原床号1      病人变动记录.床号%Type;
  v_原床号2      病人变动记录.床号%Type;
  v_当前床号1    病人变动记录.床号%Type;
  v_当前床号2    病人变动记录.床号%Type;
  v_出院科室id   病人变动记录.科室id%Type;
  v_床位等级id   病人变动记录.床位等级id%Type;
  v_险类         病案主页.险类%Type;
  v_姓名         病人信息.姓名%Type;
  n_原科室id     病案主页.出院科室id%Type;
  n_原病区id     病案主页.当前病区id%Type;
  v_母婴转科标志 病案主页.母婴转科标志%Type;
  v_Tmp          Varchar2(100);
  d_开始时间     Date;
Begin
  If 撤销方式_In Is Null Then
    v_Error := '[ZLSOFT]没有指明具体的撤销操作！[ZLSOFT]';
    Raise Err_Custom;
  Else
    v_撤销方式 := 撤销方式_In;
  End If;

  Open c_Curlog;
  Fetch c_Curlog
    Into r_Curlogrow;
  If c_Curlog%Rowcount = 0 Then
    v_Error := '[ZLSOFT]病人当前没有可以撤消的操作！[ZLSOFT]';
    Close c_Curlog;
    Raise Err_Custom;
  End If;
  
  Select Count(*) Into v_Count From 病案主页 Where 病人id = 病人id_In And 主页id > 主页id_In;
  If v_Count > 0 Then
    v_Error := '[ZLSOFT]您只能对病人的最后一次住院进行撤销操作,本次撤销操作终止![ZLSOFT]';
    Raise Err_Custom;
  End If;
  
  Select Count(Id)
  Into v_Countcurlog
  From 病人变动记录
  Where 病人id = 病人id_In And 主页id = 主页id_In And (终止时间 Is Null Or 终止原因 = 1)
  Order By 终止时间 Desc, 开始时间 Desc;

  Select Count(a.床号)
  Into v_Countprelog
  From 病人变动记录 a, 床位状况记录 b, 病案主页 c
  Where a.病人id = 病人id_In And a.主页id = 主页id_In And 终止原因 = 4 And a.病区id = b.病区id And a.病人id = c.病人id And a.主页id = c.主页id And
        a.床号 = b.床号
  Order By 终止时间 Desc, 开始时间 Desc;

  --判断是否撤消床位对换
  If v_撤销方式 = '换床' And v_Countcurlog <= 1 And v_Countprelog <= 1 Then
    Open c_Prebed(病人id_In, 主页id_In);
    Fetch c_Prebed
      Into r_Prebedrow;
  
    v_出院科室id := r_Prebedrow.出院科室id;
    v_病区id     := r_Prebedrow.当前病区id;
    v_原床号1    := r_Prebedrow.床号;
    v_当前床号1  := r_Prebedrow.出院病床;
  
    For r_Prebedpati In c_Prebedpati(v_出院科室id, v_病区id, v_原床号1) Loop
      v_病人id    := r_Prebedpati.病人id;
      v_主页id    := r_Prebedpati.主页id;
      v_原床号2   := r_Prebedpati.床号;
      v_当前床号2 := r_Prebedpati.出院病床;
    
      If v_病人id <> 0 And v_主页id <> 0 And v_原床号1 = v_当前床号2 And v_原床号2 = v_当前床号1 Then
        v_撤销方式 := '床位对换';
        Select 险类 Into v_险类 From 病案主页 Where 病人id = v_病人id And 主页id = v_主页id;
        If v_险类 Is Null Then
          Zl_病人变动记录_Undo(v_病人id, v_主页id, 操作员编号_In, 操作员姓名_In, Null, 床号_In, 主床位_In, v_撤销方式);
        Else
          Zl_病人变动记录_Undo(v_病人id, v_主页id, 操作员编号_In, 操作员姓名_In, '1', 床号_In, 主床位_In, v_撤销方式);
        End If;
      End If;
    
      --只对最近一次床位是对方床位的记录进行处理
      Exit;
    End Loop;
  End If;

  If r_Curlogrow.终止时间 Is Null And r_Curlogrow.开始时间 Is Null And r_Curlogrow.开始原因 = 3 And v_撤销方式 = '转科' Then
    --撤消转科(标志)
    Delete From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 Is Null And 终止时间 Is Null And 开始原因 = 3;
  
    Update 病案主页 Set 状态 = 0 Where 病人id = 病人id_In And 主页id = 主页id_In;
  
    Close c_Curlog;
  Elsif r_Curlogrow.终止时间 Is Null And r_Curlogrow.开始时间 Is Null And r_Curlogrow.开始原因 = 15 And v_撤销方式 = '转病区' Then
    --撤消转科(标志)
    Delete From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 Is Null And 终止时间 Is Null And 开始原因 = 15;
  
    Update 病案主页 Set 状态 = 0 Where 病人id = 病人id_In And 主页id = 主页id_In;
  
    Close c_Curlog;
  Elsif r_Curlogrow.终止时间 Is Not Null And r_Curlogrow.终止原因 = 1 And v_撤销方式 = '出院' Then
    --撤消出院
    v_开始时间 := r_Curlogrow.终止时间; --新增的变动记录的开始时间
  
    Select Zl_住院日报_Count(r_Curlogrow.科室id, r_Curlogrow.终止时间) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]已产生业务时间内的住院日报,不能办理该业务![ZLSOFT]';
      Raise Err_Custom;
    End If;
    --是否进行过电子病案审查
    Select Nvl(病案状态, 0) Into v_病案状态 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
    If v_病案状态 Not In (0, 2) Then
      v_Error := '[ZLSOFT]病人的电子病案已提交审查，不能再撤消出院。[ZLSOFT]';
      Close c_Curlog;
      Raise Err_Custom;
    End If;
  
    --删除病历书写时机
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '出院', r_Curlogrow.科室id);
  
    --恢复入院
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Not Null And 终止原因 = 1;
  
    Select 开始原因
    Into v_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Update 病案主页
    Set 状态 = Decode(v_开始原因, 10, 3, 状态), 出院日期 = Null, 出院方式 = Null, 随诊标志 = Null, 随诊期限 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In;
    --处理床位
    If 床号_In Is Null Then
      --原床位没有被占用,家庭病床也不会被占用(程序中已判断被占用情况,占用会传入床号_In)
      Close c_Curlog;
      For r_Curlogrow In c_Curlog Loop
        If r_Curlogrow.床号 Is Not Null Then
          --检查床位
          Select Count(*)
          Into v_Count
          From 床位状况记录
          Where 病区id = r_Curlogrow.病区id And 床号 = r_Curlogrow.床号 And 状态 = '空床';
          If v_Count = 0 Then
            v_Error := '[ZLSOFT]操作失败,床位 ' || r_Curlogrow.床号 || ' 不是空床！[ZLSOFT]';
            Raise Err_Custom;
          End If;
          --重新占用床位
          Update 床位状况记录
          Set 状态 = '占用', 病人id = 病人id_In, 等级id = r_Curlogrow.床位等级id, 科室id = r_Curlogrow.科室id --强行恢复以前的科室,共用床也不用处理了。
          Where 病区id = r_Curlogrow.病区id And 床号 = r_Curlogrow.床号;
        End If;
      
        If Nvl(r_Curlogrow.附加床位, 0) = 0 Then
          Update 病人信息
          Set 出院时间 = Null, 当前病区id = r_Curlogrow.病区id, 当前科室id = r_Curlogrow.科室id, 当前床号 = r_Curlogrow.床号, 在院 = 1
          Where 病人id = 病人id_In;
        
          --更新在院病人
          Begin
            Update 在院病人
            Set 病区id = Nvl(r_Curlogrow.病区id, 0), 科室id = r_Curlogrow.科室id
            Where 病人id = 病人id_In;
            If Sql%Rowcount = 0 Then
              Insert Into 在院病人
                (病人id, 科室id, 病区id)
              Values
                (病人id_In, r_Curlogrow.科室id, Nvl(r_Curlogrow.病区id, 0));
            End If;
          Exception
            When Others Then
              Null;
          End;
        
        End If;
      End Loop;
    Else
      --原床位被占用，传入新安排的床位,入住一张或多张病床病床
      v_床号串 := 床号_In || ',';
      --如果病人出院前状态为预出院，则撤消预出院
      If v_开始原因 = 10 Then
        --撤消预出院
        Update 病案主页 Set 状态 = 0 Where 病人id = 病人id_In And 主页id = 主页id_In;
        --恢复变动
        Delete From 病人变动记录
        Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 10 And 终止时间 Is Null;
      
        Update 病人变动记录
        Set 终止时间 = v_开始时间, 终止原因 = 4, 终止人员 = 操作员姓名_In
        Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 10 And 终止时间 = r_Curlogrow.开始时间;
      Else
        Update 病人变动记录
        Set 终止时间 = v_开始时间, 终止原因 = 4, 终止人员 = 操作员姓名_In
        Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
      End If;
    
      While v_床号串 Is Not Null Loop
        v_床号 := Substr(v_床号串, 1, Instr(v_床号串, ',') - 1);
        --原始床位等级与新床位等级及数量在程序中判断
        --检查床位
        Select Count(*)
        Into v_Count
        From 床位状况记录
        Where 病区id = r_Curlogrow.病区id And 床号 = v_床号 And 状态 = '空床';
        If v_Count = 0 Then
          v_Error := '操作失败,床位 ' || v_床号 || ' 不是空床！';
          Close c_Curlog;
          Raise Err_Custom;
        End If;
        --更新床位状况记录
        Update 床位状况记录
        Set 状态 = '占用', 病人id = 病人id_In, 科室id = r_Curlogrow.科室id
        Where 病区id = r_Curlogrow.病区id And 床号 = v_床号;
      
        Select 等级id Into v_床位等级id From 床位状况记录 Where 病区id = r_Curlogrow.病区id And 床号 = v_床号;
        ----新增原因为4
        Insert Into 病人变动记录
          (Id, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情, 操作员编号,
           操作员姓名)
        Values
          (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, v_开始时间, 4, Decode(主床位_In, v_床号, 0, 1), r_Curlogrow.病区id,
           r_Curlogrow.科室id, r_Curlogrow.医疗小组id, r_Curlogrow.护理等级id, v_床位等级id, v_床号, r_Curlogrow.责任护士, r_Curlogrow.经治医师,
           r_Curlogrow.主治医师, r_Curlogrow.主任医师, r_Curlogrow.病情, 操作员编号_In, 操作员姓名_In);
      
        v_床号串 := Substr(v_床号串, Instr(v_床号串, ',') + 1);
      End Loop;
      --更新病人信息
      Update 病人信息
      Set 出院时间 = Null, 当前病区id = r_Curlogrow.病区id, 当前科室id = r_Curlogrow.科室id, 当前床号 = 主床位_In, 在院 = 1
      Where 病人id = 病人id_In;
    
      --更新在院病人
      Begin
        Update 在院病人 Set 病区id = Nvl(r_Curlogrow.病区id, 0), 科室id = r_Curlogrow.科室id Where 病人id = 病人id_In;
        If Sql%Rowcount = 0 Then
          Insert Into 在院病人
            (病人id, 科室id, 病区id)
          Values
            (病人id_In, r_Curlogrow.科室id, Nvl(r_Curlogrow.病区id, 0));
        End If;
      Exception
        When Others Then
          Null;
      End;
    
      --更新病案主页出院病床
      Update 病案主页 Set 出院病床 = 主床位_In Where 病人id = 病人id_In And 主页id = 主页id_In;
      Close c_Curlog;
    End If;
    --删除出院诊断 保留该诊断信息
    --Delete From 病人诊断记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 诊断类型 in (3,13) And 记录来源 = 2;
  
    Begin
      Select 共享号 Into v_共享号 From Zlsystems Where Floor(编号 / 100) = 3;
    Exception
      When Others Then
        Null;
    End;
    --删除该病人的随诊记录
    If v_共享号 = 100 Then
      Execute Immediate 'Delete From 随诊记录 Where 病人id =:1 And 主页id =:2'
        Using 病人id_In, 主页id_In;
    End If;
  Elsif r_Curlogrow.开始原因 = 1 And v_撤销方式 = '入院入住' Then
    --撤消入科(入院同时入科)
    v_开始时间 := r_Curlogrow.开始时间;
    Select Zl_住院日报_Count(r_Curlogrow.科室id, v_开始时间) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]已产生业务时间内的住院日报,不能办理该业务![ZLSOFT]';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into v_Count
    From 病人护理文件 a, 病人护理数据 b
    Where a.Id = b.文件id And 病人id = 病人id_In And 主页id = 主页id_In And 科室id = r_Curlogrow.科室id And Rownum < 2;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]病人已产生有数据的护理文件，不能办理该业务！[ZLSOFT]';
      Raise Err_Custom;
    End If;
    Delete From 病人护理文件 Where 病人id = 病人id_In And 主页id = 主页id_In And 科室id = r_Curlogrow.科室id;
  
    Close c_Curlog;
  
    If 数据_In = '1' Then
      For r_Fee In (Select No
                    From 住院费用记录
                    Where 病人id = 病人id_In And 主页id = 主页id_In And Mod(记录性质, 10) = 3 And 登记时间 >= v_开始时间
                    Group By No, 序号, Mod(记录性质, 10)
                    Having Sum(结帐金额) <> 0) Loop
        v_Error := '[ZLSOFT]该病人的自动记帐费用已结帐,不能进行撤销操作！[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --退除当前床位
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.床号 Is Not Null Then
        Update 床位状况记录
        Set 状态 = '空床', 病人id = Null, 科室id = Decode(共用, 1, Null, 科室id)
        Where 病区id = r_Curlogrow.病区id And 床号 = r_Curlogrow.床号;
      End If;
    End Loop;
    --相关信息还原
    Update 病案主页 Set 入院病床 = Null, 出院病床 = Null, 状态 = 1 Where 病人id = 病人id_In And 主页id = 主页id_In;
  
    Update 病人信息 Set 当前床号 = Null Where 病人id = 病人id_In;
  
    --恢复变动(入院同时入科不会有包床)
    --因为是同一条记录中的撤消,所以不处理人员
    Update 病人变动记录
    Set 床位等级id = Null, 床号 = Null, 责任护士 = Null, 经治医师 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 1 And 终止时间 Is Null;
  Elsif r_Curlogrow.开始原因 = 2 And v_撤销方式 = '入住' Then
    --撤消入院入科
    v_开始时间 := r_Curlogrow.开始时间;
    Select Zl_住院日报_Count(r_Curlogrow.科室id, v_开始时间) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]已产生业务时间内的住院日报,不能办理该业务![ZLSOFT]';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into v_Count
    From 病人护理文件 a, 病人护理数据 b
    Where a.Id = b.文件id And 病人id = 病人id_In And 主页id = 主页id_In And 科室id = r_Curlogrow.科室id And Rownum < 2;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]病人已产生有数据的护理文件，不能办理该业务！[ZLSOFT]';
      Raise Err_Custom;
    End If;
    Delete From 病人护理文件 Where 病人id = 病人id_In And 主页id = 主页id_In And 科室id = r_Curlogrow.科室id;
  
    If 数据_In = '1' Then
      For r_Fee In (Select No
                    From 住院费用记录
                    Where 病人id = 病人id_In And 主页id = 主页id_In And Mod(记录性质, 10) = 3 And 登记时间 >= v_开始时间
                    Group By No, 序号, Mod(记录性质, 10)
                    Having Sum(结帐金额) <> 0) Loop
        v_Error := '[ZLSOFT]该病人的自动记帐费用已结帐,不能进行撤销操作！[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --删除病历书写时机
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '入院', r_Curlogrow.科室id);
    Close c_Curlog;
    --退除当前床位
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.床号 Is Not Null Then
        Update 床位状况记录
        Set 状态 = '空床', 病人id = Null, 科室id = Decode(共用, 1, Null, 科室id)
        Where 病区id = r_Curlogrow.病区id And 床号 = r_Curlogrow.床号;
      End If;
    End Loop;
    --相关信息还原
    Open c_Prelog(v_开始时间, 2);
    Fetch c_Prelog
      Into r_Prelogrow;
    Update 病案主页
    Set 入院病床 = Null, 出院病床 = Null, 状态 = 1, 当前病况 = r_Prelogrow.病情, 入院病况 = r_Prelogrow.病情, 医疗小组id = r_Prelogrow.医疗小组id,
        护理等级id = r_Prelogrow.护理等级id
    Where 病人id = 病人id_In And 主页id = 主页id_In;
    Close c_Prelog;
  
    Update 病人信息 Set 当前床号 = Null Where 病人id = 病人id_In;
    Delete 病案主页从表
    Where 病人id = 病人id_In And 主页id = 主页id_In And (信息名 = '主治医师' Or 信息名 = '主任医师');
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 2 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 2 And 终止时间 = v_开始时间;
  Elsif r_Curlogrow.开始原因 = 3 And v_撤销方式 = '转科入住' Then
    --撤消转科入科
    v_开始时间 := r_Curlogrow.开始时间;
    Select Zl_住院日报_Count(r_Curlogrow.科室id, v_开始时间) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]已产生业务时间内的住院日报,不能办理该业务![ZLSOFT]';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into v_Count
    From 病人护理文件
    Where 病人id = 病人id_In And 主页id = 主页id_In And 科室id = r_Curlogrow.科室id And 创建时间 >= r_Curlogrow.开始时间;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]病人已产生护理文件，不能办理该业务！[ZLSOFT]';
      Raise Err_Custom;
    End If;
  
    If 数据_In = '1' Then
      For r_Fee In (Select No
                    From 住院费用记录
                    Where 病人id = 病人id_In And 主页id = 主页id_In And Mod(记录性质, 10) = 3 And 登记时间 >= v_开始时间
                    Group By No, 序号, Mod(记录性质, 10)
                    Having Sum(结帐金额) <> 0) Loop
        v_Error := '[ZLSOFT]该病人的自动记帐费用已结帐,不能进行撤销操作！[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
  
    --删除病历书写时机
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '转科', r_Curlogrow.科室id);
    Close c_Curlog;
    --退除当前床位
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.床号 Is Not Null Then
        Update 床位状况记录
        Set 状态 = '空床', 病人id = Null, 科室id = Decode(共用, 1, Null, 科室id)
        Where 病区id = r_Curlogrow.病区id And 床号 = r_Curlogrow.床号;
      End If;
    End Loop;
    --检查及还原原床位
    For r_Prelogrow In c_Prelog(v_开始时间, 3) Loop
      d_开始时间 := r_Prelogrow.开始时间;
      If r_Prelogrow.床号 Is Not Null Then
        Select Count(*)
        Into v_Count
        From 床位状况记录
        Where 病区id = r_Prelogrow.病区id And 床号 = r_Prelogrow.床号 And 状态 = '空床';
        If v_Count = 0 Then
          v_Error := '[ZLSOFT]病人转入该科室前的床位 ' || r_Prelogrow.床号 || ' 当前非空床或已经撤消！[ZLSOFT]';
          Raise Err_Custom;
        End If;
      
        Update 床位状况记录
        Set 状态 = '占用', 病人id = 病人id_In, 等级id = r_Prelogrow.床位等级id, 科室id = r_Prelogrow.科室id --强行恢复以前的科室,共用床也不用处理了。
        Where 病区id = r_Prelogrow.病区id And 床号 = r_Prelogrow.床号;
      End If;
      --相关信息还原
      If Nvl(r_Prelogrow.附加床位, 0) = 0 Then
        --判断是否有婴儿
        Select Count(1) Into v_Count From 病人新生儿记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
        If v_Count > 0 Then
          Select 母婴转科标志 Into v_母婴转科标志 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
          If v_母婴转科标志 Is Not Null Then
            If Substr(v_母婴转科标志, Length(v_母婴转科标志)) = '1' Then
              --如果是1表示母亲和婴儿未分开，则清空“病案主页.婴儿科室ID”和”病案主页.婴儿病区ID”；
              Update 病案主页 Set 婴儿科室id = Null, 婴儿病区id = Null Where 病人id = 病人id_In And 主页id = 主页id_In;
              If Length(v_母婴转科标志) > 1 Then
                v_母婴转科标志 := Substr(v_母婴转科标志, 1, Length(v_母婴转科标志) - 1);
              Else
                v_母婴转科标志 := '';
              End If;
            Else
              --如果是0，表示是上次转科是母亲单独转走的，则重新将原产科科室和病区填写到“病案主页.婴儿科室ID”和”病案主页.婴儿病区ID”（从病人变动记录中取），并清除最后一位标识
              If Length(v_母婴转科标志) > 1 Then
                v_母婴转科标志 := Substr(v_母婴转科标志, 1, Length(v_母婴转科标志) - 1);
                --查看上一次转科的标识
                If Substr(v_母婴转科标志, Length(v_母婴转科标志)) = '1' Then
                  Update 病案主页
                  Set 婴儿科室id = Null, 婴儿病区id = Null
                  Where 病人id = 病人id_In And 主页id = 主页id_In;
                Else
                  --取出转科前的婴儿科室病区ID
                  v_Tmp   := v_母婴转科标志;
                  v_Count := 1;
                  While v_Tmp Is Not Null Loop
                    v_Count := v_Count + 1;
                    If Substr(v_Tmp, Length(v_Tmp)) = '1' Then
                      Select Max(a.科室id) As 科室id, Max(a.病区id) As 科室id
                      Into n_原科室id, n_原病区id
                      From (Select 科室id, 病区id, Rownum As 序号
                             From (Select 科室id, 病区id
                                    From 病人变动记录
                                    Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 3 And 附加床位 = 0
                                    Order By 开始时间 Desc)) a
                      Where 序号 = v_Count;
                    End If;
                    If Length(v_Tmp) = 1 Then
                      v_Tmp := '';
                    Else
                      v_Tmp := Substr(v_Tmp, 1, Length(v_Tmp) - 1);
                    End If;
                  End Loop;
                  If Nvl(n_原科室id, 0) = 0 Then
                    --如果没有找到，则取入院科室
                    Select Max(b.科室id), Max(b.病区id)
                    Into n_原科室id, n_原病区id
                    From 病人变动记录 b
                    Where b.病人id = 病人id_In And b.主页id = 主页id_In And b.科室id Is Not Null And b.病区id Is Not Null And
                          b.开始时间 = (Select Min(a.开始时间)
                                    From 病人变动记录 a
                                    Where a.病人id = b.病人id And a.主页id = b.主页id And a.科室id Is Not Null And
                                          a.病区id Is Not Null And a.附加床位 = 0);
                  End If;
                
                  If Nvl(n_原科室id, 0) <> 0 Then
                    Update 病案主页
                    Set 婴儿科室id = n_原科室id, 婴儿病区id = n_原病区id
                    Where 病人id = 病人id_In And 主页id = 主页id_In;
                  Else
                    --之前没有转科记录,清空婴儿科室ID和婴儿病区ID
                    Update 病案主页
                    Set 婴儿科室id = Null, 婴儿病区id = Null
                    Where 病人id = 病人id_In And 主页id = 主页id_In;
                  End If;
                End If;
              Else
                --只有这一次转科,回退后清空婴儿科室病区ID
                v_母婴转科标志 := '';
                Update 病案主页
                Set 婴儿科室id = Null, 婴儿病区id = Null
                Where 病人id = 病人id_In And 主页id = 主页id_In;
              End If;
            End If;
            --去除最后一位标识
            Update 病案主页 Set 母婴转科标志 = v_母婴转科标志 Where 病人id = 病人id_In And 主页id = 主页id_In;
          End If;
        
        End If;
      
        Update 病案主页
        Set 状态 = 2, 当前病区id = r_Prelogrow.病区id, 出院科室id = r_Prelogrow.科室id, 医疗小组id = r_Prelogrow.医疗小组id,
            出院病床 = r_Prelogrow.床号, 护理等级id = r_Prelogrow.护理等级id, 责任护士 = r_Prelogrow.责任护士, 住院医师 = r_Prelogrow.经治医师,
            当前病况 = r_Curlogrow.病情
        Where 病人id = 病人id_In And 主页id = 主页id_In;
      
        Update 病案主页从表
        Set 信息值 = r_Prelogrow.主治医师
        Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主治医师';
        Update 病案主页从表
        Set 信息值 = r_Prelogrow.主任医师
        Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主任医师';
      
        Update 病人信息
        Set 当前病区id = r_Prelogrow.病区id, 当前科室id = r_Prelogrow.科室id, 当前床号 = r_Prelogrow.床号
        Where 病人id = 病人id_In;
      
        --更新在院病人
        Update 在院病人 Set 病区id = Nvl(r_Prelogrow.病区id, 0), 科室id = r_Prelogrow.科室id Where 病人id = 病人id_In;
      
      End If;
    End Loop;
    --恢复变动(恢复到临时转科标记状态)
    Delete From 病人变动记录
    Where 附加床位 = 1 And 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 3 And 终止时间 Is Null;
  
    Select 终止人员
    Into v_终止人员
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 3 And 终止时间 = v_开始时间 And Nvl(附加床位, 0) = 0;
    --临时记录的操作员信息记录的是终止人员,因为没有记录终止人员编号,就不恢复
    Update 病人变动记录
    Set 开始时间 = Null, 医疗小组id = Null, 护理等级id = Null, 床位等级id = Null, 床号 = Null, 责任护士 = Null, 经治医师 = Null, 操作员编号 = Null,
        操作员姓名 = v_终止人员, 上次计算时间 = Null, 病区id = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 3 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 3 And 终止时间 = v_开始时间 And 开始时间 = d_开始时间;
  
  Elsif r_Curlogrow.开始原因 = 4 And v_撤销方式 = '换床' Then
    --撤消换床
    v_开始时间 := r_Curlogrow.开始时间;
    Close c_Curlog;
    If 数据_In = '1' Then
      For r_Fee In (Select No
                    From 住院费用记录
                    Where 病人id = 病人id_In And 主页id = 主页id_In And Mod(记录性质, 10) = 3 And 登记时间 >= v_开始时间
                    Group By No, 序号, Mod(记录性质, 10)
                    Having Sum(结帐金额) <> 0) Loop
        v_Error := '[ZLSOFT]该病人的自动记帐费用已结帐,不能进行撤销操作！[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --退除当前床位
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.床号 Is Not Null Then
        Update 床位状况记录
        Set 状态 = '空床', 病人id = Null, 科室id = Decode(共用, 1, Null, 科室id)
        Where 病区id = r_Curlogrow.病区id And 床号 = r_Curlogrow.床号;
      End If;
    End Loop;
    --检查及还原原床位
    For r_Prelogrow In c_Prelog(v_开始时间, 4) Loop
      d_开始时间 := r_Prelogrow.开始时间;
      If r_Prelogrow.床号 Is Not Null Then
        Select Count(*)
        Into v_Count
        From 床位状况记录
        Where 病区id = r_Prelogrow.病区id And 床号 = r_Prelogrow.床号 And 状态 = '空床';
        If v_Count = 0 Then
          v_Error := '[ZLSOFT]病人最近一次换床前所入住的床位 ' || r_Prelogrow.床号 || ' 当前非空床或已经撤消！[ZLSOFT]';
          Raise Err_Custom;
        End If;
      
        Update 床位状况记录
        Set 状态 = '占用', 病人id = 病人id_In, 等级id = r_Prelogrow.床位等级id, 科室id = Decode(共用, 1, r_Prelogrow.科室id, 科室id)
        Where 病区id = r_Prelogrow.病区id And 床号 = r_Prelogrow.床号;
      End If;
      --病人信息、病案主页,仅当病区与科室独立时,换床才可以换病区,此处为简化判断,统一还原病区
      If Nvl(r_Prelogrow.附加床位, 0) = 0 Then
        Update 病案主页
        Set 出院病床 = r_Prelogrow.床号, 当前病区id = r_Prelogrow.病区id
        Where 病人id = 病人id_In And 主页id = 主页id_In;
      
        Update 病人信息 Set 当前床号 = r_Prelogrow.床号, 当前病区id = r_Prelogrow.病区id Where 病人id = 病人id_In;
        --更新在院病人
        Update 在院病人 Set 病区id = Nvl(r_Prelogrow.病区id, 0) Where 病人id = 病人id_In;
      End If;
    End Loop;
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 4 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 4 And 终止时间 = v_开始时间 And 开始时间 = d_开始时间;
  Elsif r_Curlogrow.开始原因 = 4 And v_撤销方式 = '床位对换' Then
    --撤消床位对换
    v_开始时间 := r_Curlogrow.开始时间;
    Select 姓名 Into v_姓名 From 病人信息 Where 病人id = 病人id_In;
    Close c_Curlog;
    If 数据_In = '1' Then
      For r_Fee In (Select No
                    From 住院费用记录
                    Where 病人id = 病人id_In And 主页id = 主页id_In And Mod(记录性质, 10) = 3 And 登记时间 >= v_开始时间
                    Group By No, 序号, Mod(记录性质, 10)
                    Having Sum(结帐金额) <> 0) Loop
        v_Error := '[ZLSOFT]病人 ' || v_姓名 || ' 的自动记帐费用已结帐,不能进行撤销操作！[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --退除当前床位
    --检查及还原原床位
    For r_Prelogrow In c_Prelog(v_开始时间, 4) Loop
      d_开始时间 := r_Prelogrow.开始时间;
      If r_Prelogrow.床号 Is Not Null Then
        Update 床位状况记录
        Set 状态 = '占用', 病人id = 病人id_In, 等级id = r_Prelogrow.床位等级id, 科室id = Decode(共用, 1, r_Prelogrow.科室id, 科室id)
        Where 病区id = r_Prelogrow.病区id And 床号 = r_Prelogrow.床号;
      End If;
      --病人信息、病案主页,仅当病区与科室独立时,换床才可以换病区,此处为简化判断,统一还原病区
      If Nvl(r_Prelogrow.附加床位, 0) = 0 Then
        Update 病案主页
        Set 出院病床 = r_Prelogrow.床号, 当前病区id = r_Prelogrow.病区id
        Where 病人id = 病人id_In And 主页id = 主页id_In;
      
        Update 病人信息 Set 当前床号 = r_Prelogrow.床号, 当前病区id = r_Prelogrow.病区id Where 病人id = 病人id_In;
        --更新在院病人
        Update 在院病人 Set 病区id = Nvl(r_Prelogrow.病区id, 0) Where 病人id = 病人id_In;
      End If;
    End Loop;
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 4 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 4 And 终止时间 = v_开始时间 And 开始时间 = d_开始时间;
  Elsif r_Curlogrow.开始原因 = 5 And v_撤销方式 = '床位等级变动' Then
    --撤消床位等级变动
    If 数据_In = '1' Then
      For r_Fee In (Select No
                    From 住院费用记录
                    Where 病人id = 病人id_In And 主页id = 主页id_In And Mod(记录性质, 10) = 3 And 收费细目id = r_Curlogrow.床位等级id And
                          登记时间 >= r_Curlogrow.开始时间
                    Group By No, 序号, Mod(记录性质, 10)
                    Having Sum(结帐金额) <> 0) Loop
        Close c_Curlog;
        v_Error := '[ZLSOFT]该病人的自动记帐费用已结帐,不能进行撤销操作！[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --还原原床位的等级
    For r_Prelogrow In c_Prelog(r_Curlogrow.开始时间, 5) Loop
      d_开始时间 := r_Prelogrow.开始时间;
      If r_Prelogrow.床号 Is Not Null Then
        Update 床位状况记录
        Set 等级id = r_Prelogrow.床位等级id
        Where 病区id = r_Prelogrow.病区id And 床号 = r_Prelogrow.床号;
      End If;
    End Loop;
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 5 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 5 And 终止时间 = r_Curlogrow.开始时间 And 开始时间 = d_开始时间;
  
    Close c_Curlog;
  Elsif r_Curlogrow.开始原因 = 6 And v_撤销方式 = '护理等级变动' Then
    --撤消护理等级变动
    If 数据_In = '1' Then
      For r_Fee In (Select No
                    From 住院费用记录
                    Where 病人id = 病人id_In And 主页id = 主页id_In And Mod(记录性质, 10) = 3 And 收费细目id = r_Curlogrow.护理等级id And
                          登记时间 >= r_Curlogrow.开始时间
                    Group By No, 序号, Mod(记录性质, 10)
                    Having Sum(结帐金额) <> 0) Loop
        Close c_Curlog;
        v_Error := '[ZLSOFT]该病人的自动记帐费用已结帐,不能进行撤销操作！[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
  
    Open c_Prelog(r_Curlogrow.开始时间, 6);
    Fetch c_Prelog
      Into r_Prelogrow;
    --恢复原护理等级
    Update 病案主页 Set 护理等级id = r_Prelogrow.护理等级id Where 病人id = 病人id_In And 主页id = 主页id_In;
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 6 And 终止时间 Is Null;
  
    --医嘱产生的护理等级变动没有记录秒，可能前一等级的停止时间与当前等级的开始时间是同一分钟，所以要取max(id)
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 6 And 终止时间 = r_Curlogrow.开始时间  And 开始时间 = r_Prelogrow.开始时间;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.开始原因 = 7 And v_撤销方式 = '经治医师改变' Then
    --删除病历书写时机
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '交班', r_Curlogrow.科室id);
    --撤消经治医师改变
    Open c_Prelog(r_Curlogrow.开始时间, 7);
    Fetch c_Prelog
      Into r_Prelogrow;
    --恢复原医师
    Update 病案主页 Set 住院医师 = r_Prelogrow.经治医师 Where 病人id = 病人id_In And 主页id = 主页id_In;
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 7 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 7 And 终止时间 = r_Curlogrow.开始时间  And 开始时间 = r_Prelogrow.开始时间;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.开始原因 = 8 And v_撤销方式 = '责任护士改变' Then
    --撤消责任护士改变
    Open c_Prelog(r_Curlogrow.开始时间, 8);
    Fetch c_Prelog
      Into r_Prelogrow;
    --恢复原责任护士
    Update 病案主页 Set 责任护士 = r_Prelogrow.责任护士 Where 病人id = 病人id_In And 主页id = 主页id_In;
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 8 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 8 And 终止时间 = r_Curlogrow.开始时间  And 开始时间 = r_Prelogrow.开始时间;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.开始原因 = 9 And v_撤销方式 = '转为住院病人' Then
    --撤消转为住院病人
    --删除病历书写时机
    Open c_Prelog(r_Curlogrow.开始时间, 9);
    Fetch c_Prelog
      Into r_Prelogrow;
      
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '入院', r_Curlogrow.科室id);
    Update 病案主页 Set 病人性质 = 2 Where 病人id = 病人id_In And 主页id = 主页id_In;
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 9 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 9 And 终止时间 = r_Curlogrow.开始时间 And 开始时间 = r_Prelogrow.开始时间;
  
    If 主页id_In = 1 And Nvl(数据_In, '0') = '1' Then
      Update 病人信息 Set 住院号 = Null Where 病人id = 病人id_In;
      Update 病案主页 Set 住院号 = Null Where 病人id = 病人id_In;
    End If;
  
    Close c_Curlog;
    Close c_Prelog;
  Elsif r_Curlogrow.开始原因 = 10 And v_撤销方式 = '预出院' Then
    --撤消预出院
    --删除病历书写时机
    Open c_Prelog(r_Curlogrow.开始时间, 10);
    Fetch c_Prelog
      Into r_Prelogrow;
      
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '出院', r_Curlogrow.科室id);
    --恢复住院状态
    Update 病案主页 Set 状态 = 0 Where 病人id = 病人id_In And 主页id = 主页id_In;
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 10 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 上次计算时间 = Null, 终止人员 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 10 And 终止时间 = r_Curlogrow.开始时间 And 开始时间 = r_Prelogrow.开始时间;
  
    Close c_Curlog;
    Close c_Prelog;
  Elsif r_Curlogrow.开始原因 = 11 And v_撤销方式 = '主治医师变动' Then
    --撤消主治医师改变
    Open c_Prelog(r_Curlogrow.开始时间, 11);
    Fetch c_Prelog
      Into r_Prelogrow;
    --恢复原主治医师
    Update 病案主页从表
    Set 信息值 = r_Prelogrow.主治医师
    Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主治医师';
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 11 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 11 And 终止时间 = r_Curlogrow.开始时间  And 开始时间 = r_Prelogrow.开始时间;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.开始原因 = 12 And v_撤销方式 = '主任医师变动' Then
    --撤消主任医师改变
    Open c_Prelog(r_Curlogrow.开始时间, 12);
    Fetch c_Prelog
      Into r_Prelogrow;
    --恢复原主任医师
    Update 病案主页从表
    Set 信息值 = r_Prelogrow.主任医师
    Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主任医师';
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 12 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 12 And 终止时间 = r_Curlogrow.开始时间  And 开始时间 = r_Prelogrow.开始时间;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.开始原因 = 13 And v_撤销方式 = '病况变动' Then
    --撤消病情改变
    Open c_Prelog(r_Curlogrow.开始时间, 13);
    Fetch c_Prelog
      Into r_Prelogrow;
    --恢复原病情
    Update 病案主页 Set 当前病况 = r_Prelogrow.病情 Where 病人id = 病人id_In And 主页id = 主页id_In;
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 13 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 13 And 终止时间 = r_Curlogrow.开始时间 And 开始时间 = r_Prelogrow.开始时间;
  
    Close c_Prelog;
    Close c_Curlog;
  
  Elsif r_Curlogrow.开始原因 = 14 And v_撤销方式 = '转医疗小组' Then
    --删除病历书写时机
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '交班', r_Curlogrow.科室id);
  
    --撤消医疗小组改变
    Open c_Prelog(r_Curlogrow.开始时间, 14);
    Fetch c_Prelog
      Into r_Prelogrow;
    --恢复原医疗小组
    Update 病案主页
    Set 医疗小组id = r_Prelogrow.医疗小组id, 住院医师 = r_Prelogrow.经治医师
    Where 病人id = 病人id_In And 主页id = 主页id_In;
    --恢复原主治医师
    Update 病案主页从表
    Set 信息值 = r_Prelogrow.主治医师
    Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主治医师';
  
    --恢复变动
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 14 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 14 And 终止时间 = r_Curlogrow.开始时间  And 开始时间 = r_Prelogrow.开始时间;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.开始原因 = 15 And v_撤销方式 = '转病区入住' Then
    --撤消入病区
    v_开始时间 := r_Curlogrow.开始时间;
  
    If 数据_In = '1' Then
      For r_Fee In (Select No
                    From 住院费用记录
                    Where 病人id = 病人id_In And 主页id = 主页id_In And Mod(记录性质, 10) = 3 And 登记时间 >= v_开始时间
                    Group By No, 序号, Mod(记录性质, 10)
                    Having Sum(结帐金额) <> 0) Loop
        v_Error := '[ZLSOFT]该病人的自动记帐费用已结帐,不能进行撤销操作！[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
  
    Open c_Prelog(v_开始时间, 15);
    Fetch c_Prelog
      Into r_Prelogrow;
  
    d_开始时间 := r_Prelogrow.开始时间;
    --对有效的医嘱(未停止、作废的长嘱，未发送的临嘱)，为原病区执行的，将执行科室自动更换为新的病区
    Update 病人医嘱记录
    Set 执行科室id = r_Prelogrow.病区id
    Where 病人id = 病人id_In And 主页id = 主页id_In And 执行科室id = r_Curlogrow.病区id And 医嘱状态 Not In (4, 8, 9) And 开嘱时间 < v_开始时间;
    --对未审核的记帐划价单，为原病区执行的，将执行科室自动更换为新的病区
    Update 住院费用记录
    Set 执行部门id = r_Prelogrow.病区id
    Where 病人id = 病人id_In And 主页id = 主页id_In And 执行部门id = r_Curlogrow.病区id And 记录状态 = 0;
    Close c_Prelog;
    Close c_Curlog;
  
    --退除当前床位
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.床号 Is Not Null Then
        Update 床位状况记录
        Set 状态 = '空床', 病人id = Null, 科室id = Decode(共用, 1, Null, 科室id)
        Where 病区id = r_Curlogrow.病区id And 床号 = r_Curlogrow.床号;
      End If;
    End Loop;
    --检查及还原原床位
    For r_Prelogrow In c_Prelog(v_开始时间, 15) Loop
      If r_Prelogrow.床号 Is Not Null Then
        Select Count(*)
        Into v_Count
        From 床位状况记录
        Where 病区id = r_Prelogrow.病区id And 床号 = r_Prelogrow.床号 And 状态 = '空床';
        If v_Count = 0 Then
          v_Error := '[ZLSOFT]病人转入该病区前的床位 ' || r_Prelogrow.床号 || ' 当前非空床或已经撤消！[ZLSOFT]';
          Raise Err_Custom;
        End If;
      
        Update 床位状况记录
        Set 状态 = '占用', 病人id = 病人id_In, 等级id = r_Prelogrow.床位等级id, 科室id = r_Prelogrow.科室id --强行恢复以前的科室,共用床也不用处理了。
        Where 病区id = r_Prelogrow.病区id And 床号 = r_Prelogrow.床号;
      End If;
      --相关信息还原
      If Nvl(r_Prelogrow.附加床位, 0) = 0 Then
        Update 病案主页
        Set 状态 = 2, 当前病区id = r_Prelogrow.病区id, 出院科室id = r_Prelogrow.科室id, 医疗小组id = r_Prelogrow.医疗小组id,
            出院病床 = r_Prelogrow.床号, 护理等级id = r_Prelogrow.护理等级id, 责任护士 = r_Prelogrow.责任护士, 住院医师 = r_Prelogrow.经治医师,
            当前病况 = r_Curlogrow.病情
        Where 病人id = 病人id_In And 主页id = 主页id_In;
      
        Update 病人信息 Set 当前病区id = r_Prelogrow.病区id, 当前床号 = r_Prelogrow.床号 Where 病人id = 病人id_In;
      
        --更新在院病人
        Update 在院病人 Set 病区id = Nvl(r_Prelogrow.病区id, 0) Where 病人id = 病人id_In;
      
      End If;
    End Loop;
  
    --恢复变动(恢复到临时转科标记状态)
    Delete From 病人变动记录
    Where 附加床位 = 1 And 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 15 And 终止时间 Is Null;
  
    Select 终止人员
    Into v_终止人员
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 15 And 终止时间 = v_开始时间 And Nvl(附加床位, 0) = 0;
    --临时记录的操作员信息记录的是终止人员,因为没有记录终止人员编号,就不恢复
    Update 病人变动记录
    Set 开始时间 = Null, 医疗小组id = Null, 护理等级id = Null, 床位等级id = Null, 床号 = Null, 责任护士 = Null, 经治医师 = Null, 操作员编号 = Null,
        操作员姓名 = v_终止人员, 上次计算时间 = Null, 附加床位 = Null, 主治医师 = Null, 病情 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 15 And 终止时间 Is Null;
  
    Update 病人变动记录
    Set 终止时间 = Null, 终止原因 = Null, 终止人员 = Null, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止原因 = 15 And 终止时间 = v_开始时间  And 开始时间 = d_开始时间;
  Else
    Close c_Curlog;
    v_Error := '[ZLSOFT]你执行的撤销' || v_撤销方式 || '操作已经被其他人执行,请刷新界面！[ZLSOFT]';
    Raise Err_Custom;
  End If;
  --并发操作检查
  Select Count(*)
  Into v_Count
  From 病人变动记录
  Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(附加床位, 0) = 0 And 开始时间 Is Not Null And 终止时间 Is Null;
  If v_Count > 1 Then
    v_Error := '发现病人存在非法的变动记录,当前操作不能继续！' || Chr(13) || Chr(10) || '这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In And 出院日期 Is Null;
  If v_Count = 0 Then
    v_Error := '操作失败,该病人已出院,不能进行当前操作' || Chr(13) || Chr(10) || '这可能是由于网络并发操作引起的,请刷新病人状态！';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, v_Error);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人变动记录_Undo;
/

--94041:马政,2016-03-10,材料售价调价成本价字段修正
--92931:马政,2016-01-21,材料调价生效过程升级脚本和安装脚本不同修正
Create Or Replace Procedure Zl_材料收发记录_Adjust
(
  调价id_In   In Number, --调价记录的ID
  定价_In     In Number := 0, --是否转为定价销售（更新材料特性、收费细目中的变价）
  材料id_In   In Number := 0, --当不为0时表示是成本价调价，不处理售价相关内容
  Billinfo_In In Varchar2 := Null --用于时价卫材按批次调价。格式:"批次1,现价1|批次2,现价2|....."
) As
  n_入出类别id 药品收发记录.入出类别id%Type; --入出类别
  v_调价单据号 药品收发记录.No%Type; --调价单号
  d_生效日期   Date; --调价生效时间
  n_执行调价   Number(1); --调价时刻到了
  n_实价材料   Number(1); --时价药品
  n_收费细目id Number(18); --收费细目ID
  d_审核日期   药品收发记录.审核日期%Type;
  n_零售金额   药品库存.实际金额%Type;
  n_零售价     药品库存.零售价%Type;
  n_序号       Integer(8);
  v_Infotmp    Varchar2(4000);
  v_Fields     Varchar2(4000);
  n_批次       Number(18);
  n_现价       收费价目.现价%Type;
  n_原价       收费价目.原价%Type;
  n_收发id     药品收发记录.Id%Type;
  n_时价分批   Number(1);

  Cursor c_Price --普通调价
  Is
    Select 1 记录状态, 13 单据, v_调价单据号 NO, Rownum 序号, n_入出类别id 入出类别id, m.材料id 药品id, s.批次 批次, Null 批号, s.效期,
           Decode(s.上次产地, Null, q.产地, s.上次产地) 产地, 1 付数, s.实际数量 填写数量, 0 实际数量, a.原价 成本价, 0 成本金额, a.现价 零售价, 0 扣率,
           Nvl(s.零售价, 0) As 库存零售价, s.实际金额 As 库存金额, s.实际差价 As 库存差价, '卫材调价' 摘要, User 填制人, Sysdate 填制日期, s.库房id 库房id,
           1 入出系数, a.Id 价格id, s.上次生产日期, s.灭菌效期, s.批准文号, s.上次供应商id,
           Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, a.原价, Nvl(s.实际金额, 0) / s.实际数量)) As 原售价
    From 药品库存 S, 材料特性 M, 收费价目 A, 收费项目目录 Q
    Where s.药品id = m.材料id And m.材料id = q.Id And m.材料id = a.收费细目id And s.性质 = 1 And a.变动原因 = 0 And a.Id = 调价id_In And
          a.执行日期 <= Sysdate;

  Cursor c_时价按批次调价 --时价卫材按批次调价
  Is
    Select 1 记录状态, 13 单据, v_调价单据号 NO, n_序号 + Rownum 序号, n_入出类别id 入出类别id, s.药品id 药品id, s.批次 批次, Null 批号, s.效期,
           Decode(s.上次产地, Null, b.产地, s.上次产地) 产地, 1 付数, Nvl(s.实际数量, 0) 填写数量, 0 实际数量, a.原价 成本价, 0 成本金额, n_现价 零售价, 0 扣率,
           '卫材调价' 摘要, User 填制人, Sysdate 填制日期, s.库房id 库房id, 1 入出系数, a.Id 价格id, Nvl(b.是否变价, 0) As 时价, s.实际金额 As 库存金额,
           s.实际差价 As 库存差价, Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, a.原价, Nvl(s.实际金额, 0) / s.实际数量)) As 原售价
    From 药品库存 S, 材料特性 M, 收费价目 A, 收费项目目录 B
    Where s.药品id = m.材料id And m.材料id = a.收费细目id And a.收费细目id = b.Id And s.性质 = 1 And a.变动原因 = 0 And a.Id = 调价id_In And
          a.执行日期 <= Sysdate And Nvl(s.批次, 0) = n_批次;
Begin

  If 材料id_In <> 0 Then
    --成本价调价
    Zl_材料收发记录_成本价调价(材料id_In);
    Return;
  End If;

  --取入出类别ID
  Select 类别id Into n_入出类别id From 药品单据性质 Where 单据 = 13;

  --取序列
  Select Nextno(147) Into v_调价单据号 From Dual;
  --取调价记录生效日期
  Select 收费细目id, 执行日期 Into n_收费细目id, d_生效日期 From 收费价目 Where ID = 调价id_In;
  --取该材料是否是时价药品
  Select Nvl(是否变价, 0) Into n_实价材料 From 收费项目目录 Where ID = n_收费细目id;

  If Sysdate >= d_生效日期 Then
    n_执行调价 := 1;
  Else
    n_执行调价 := 0;
  End If;

  If n_执行调价 = 1 Then
    d_审核日期 := Sysdate;
    --普通调价处理
    If Billinfo_In = '' Or Billinfo_In Is Null Then
      --非时价药品调价
      For c_调价 In c_Price Loop
        If Nvl(c_调价.填写数量, 0) = 0 And Nvl(c_调价.库存金额, 0) = 0 And Nvl(c_调价.库存差价, 0) = 0 Then
          Null;
        Elsif Nvl(c_调价.填写数量, 0) = 0 And (Nvl(c_调价.库存金额, 0) <> 0 Or Nvl(c_调价.库存差价, 0) <> 0) Then
          --数量=0 金额或差价<>0时只更新库存表中对应的零售价,并产生售价修正数据但是金额差=0，只记录最新售价，金额差和差价差不填数据


        
          --产生调价影响记录
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 摘要, 填制人, 填制日期,
             库房id, 入出系数, 价格id, 审核人, 审核日期, 生产日期, 灭菌效期, 批准文号, 供药单位id, 单量, 频次)
          Values
            (药品收发记录_Id.Nextval, c_调价.记录状态, c_调价.单据, c_调价.No, c_调价.序号, c_调价.入出类别id, c_调价.药品id, c_调价.批次, c_调价.批号, c_调价.效期,
             c_调价.产地, c_调价.付数, c_调价.填写数量, c_调价.实际数量, Decode(n_实价材料, 1, c_调价.原售价, c_调价.成本价), c_调价.成本金额, c_调价.零售价, c_调价.扣率,
             c_调价.摘要, c_调价.填制人, c_调价.填制日期, c_调价.库房id, c_调价.入出系数, c_调价.价格id, User, d_审核日期, c_调价.上次生产日期, c_调价.灭菌效期,
             c_调价.批准文号, c_调价.上次供应商id, c_调价.库存金额, c_调价.库存差价);
        
          --更新材料库存 ，只有时价卫材才更新零售价
          Update 药品库存
          Set 零售价 = Decode(n_实价材料, 1, Decode(Nvl(c_调价.批次, 0), 0, Null, c_调价.零售价), Null)
          Where 库房id = c_调价.库房id And 药品id = c_调价.药品id And 性质 = 1 And Nvl(批次, 0) = Nvl(c_调价.批次, 0);
        Else
          If n_实价材料 = 1 Then
            If c_调价.库存零售价 = 0 Then
              n_零售价 := c_调价.库存金额 / c_调价.填写数量;
            Else
              n_零售价 := c_调价.库存零售价;
            End If;
          Else
            n_零售价 := c_调价.成本价;
          End If;
          n_零售金额 := Round((c_调价.零售价 - n_零售价) * c_调价.填写数量, 2);
        
          --产生调价影响记录
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
             填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 生产日期, 灭菌效期, 批准文号, 供药单位id, 单量, 频次)
          Values
            (药品收发记录_Id.Nextval, c_调价.记录状态, c_调价.单据, c_调价.No, c_调价.序号, c_调价.入出类别id, c_调价.药品id, c_调价.批次, c_调价.批号, c_调价.效期,
             c_调价.产地, c_调价.付数, c_调价.填写数量, c_调价.实际数量, Decode(n_实价材料, 1, c_调价.原售价, c_调价.成本价), c_调价.成本金额, c_调价.零售价, c_调价.扣率,
             n_零售金额, n_零售金额, c_调价.摘要, c_调价.填制人, c_调价.填制日期, c_调价.库房id, c_调价.入出系数, c_调价.价格id, User, d_审核日期, c_调价.上次生产日期,
             c_调价.灭菌效期, c_调价.批准文号, c_调价.上次供应商id, c_调价.库存金额, c_调价.库存差价);
        
          --更新材料库存
          Update 药品库存
          Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_零售金额,
              零售价 = Decode(n_实价材料, 1, Decode(Nvl(c_调价.批次, 0), 0, Null, c_调价.零售价), Null)
          Where 库房id = c_调价.库房id And 药品id = c_调价.药品id And 性质 = 1 And Nvl(批次, 0) = Nvl(c_调价.批次, 0);
        
          If Sql%RowCount = 0 Then
            Insert Into 药品库存
              (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
            Values
              (c_调价.库房id, c_调价.药品id, c_调价.批次, 1, 0, 0, n_零售金额, n_零售金额, c_调价.效期, c_调价. 灭菌效期, c_调价.上次供应商id, c_调价.成本价,
               c_调价.批号, c_调价.上次生产日期, c_调价.产地, c_调价.批准文号,
               Decode(n_实价材料, 1, Decode(Nvl(c_调价.批次, 0), 0, Null, c_调价.零售价), Null));
          End If;
        End If;
      End Loop;
    Else
      --时价分批调价处理
      n_序号 := 0;
      --时价药品按批次调价
      v_Infotmp := Billinfo_In || '|';
      While v_Infotmp Is Not Null Loop
        --分解单据ID串
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        n_批次    := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        n_现价    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        For v_时价按批次调价 In c_时价按批次调价 Loop
          If v_时价按批次调价.填写数量 <> 0 Then
            n_原价 := Nvl(v_时价按批次调价.库存金额, 0) / v_时价按批次调价.填写数量;
          Else
            n_原价 := v_时价按批次调价.成本价;
          End If;
        
          Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
          If Nvl(v_时价按批次调价.填写数量, 0) = 0 And Nvl(v_时价按批次调价.库存金额, 0) = 0 And Nvl(v_时价按批次调价.库存差价, 0) = 0 Then
            Null;
          Elsif Nvl(v_时价按批次调价.填写数量, 0) = 0 And (Nvl(v_时价按批次调价.库存金额, 0) <> 0 Or Nvl(v_时价按批次调价.库存差价, 0) <> 0) Then
            --数量=0 金额或差价<>0时只更新库存表中对应的零售价,并产生售价修正数据但是金额差=0，只记录最新售价，金额差和差价差不填数据


          
            --产生调价影响记录
            Insert Into 药品收发记录
              (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 摘要, 填制人, 填制日期,
               库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次)
            Values
              (n_收发id, v_时价按批次调价.记录状态, v_时价按批次调价.单据, v_时价按批次调价.No, v_时价按批次调价.序号, v_时价按批次调价.入出类别id, v_时价按批次调价.药品id,
               v_时价按批次调价.批次, v_时价按批次调价.批号, v_时价按批次调价.效期, v_时价按批次调价.产地, v_时价按批次调价.付数, v_时价按批次调价.填写数量, v_时价按批次调价.实际数量,
               Decode(n_实价材料, 1, v_时价按批次调价.原售价, v_时价按批次调价.成本价), v_时价按批次调价.成本金额, v_时价按批次调价.零售价, v_时价按批次调价.扣率,
               v_时价按批次调价.摘要, v_时价按批次调价.填制人, v_时价按批次调价.填制日期, v_时价按批次调价.库房id, v_时价按批次调价.入出系数, v_时价按批次调价.价格id, User, d_审核日期,
               v_时价按批次调价.库存金额, v_时价按批次调价.库存差价);
            n_序号 := n_序号 + 1;
            --处理库存
            --更新库存零售价,只有时价分批药品才能更新零售价字段
            Update 药品库存
            Set 零售价 = Decode(v_时价按批次调价.时价, 1, Decode(Nvl(v_时价按批次调价.批次, 0), 0, Null, v_时价按批次调价.零售价), Null)
            Where 库房id = v_时价按批次调价.库房id And 药品id = v_时价按批次调价.药品id And 性质 = 1 And Nvl(批次, 0) = Nvl(v_时价按批次调价.批次, 0);
          Else
            n_零售价   := v_时价按批次调价.库存金额 / v_时价按批次调价.填写数量;
            n_零售金额 := Round((n_现价 - n_零售价) * v_时价按批次调价.填写数量, 2);
            --产生调价影响记录
            Insert Into 药品收发记录
              (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要,
               填制人, 填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次)
            Values
              (n_收发id, v_时价按批次调价.记录状态, v_时价按批次调价.单据, v_时价按批次调价.No, v_时价按批次调价.序号, v_时价按批次调价.入出类别id, v_时价按批次调价.药品id,
               v_时价按批次调价.批次, v_时价按批次调价.批号, v_时价按批次调价.效期, v_时价按批次调价.产地, v_时价按批次调价.付数, v_时价按批次调价.填写数量, v_时价按批次调价.实际数量,
               Decode(n_实价材料, 1, v_时价按批次调价.原售价, v_时价按批次调价.成本价), v_时价按批次调价.成本金额, v_时价按批次调价.零售价, v_时价按批次调价.扣率, n_零售金额,
               n_零售金额, v_时价按批次调价.摘要, v_时价按批次调价.填制人, v_时价按批次调价.填制日期, v_时价按批次调价.库房id, v_时价按批次调价.入出系数, v_时价按批次调价.价格id, User,
               d_审核日期, v_时价按批次调价.库存金额, v_时价按批次调价.库存差价);
            n_序号 := n_序号 + 1;
            --处理库存
            If v_时价按批次调价.时价 = 1 And Nvl(v_时价按批次调价.批次, 0) > 0 Then
              n_时价分批 := 1;
            Else
              n_时价分批 := 0;
            End If;
          
            If Nvl(v_时价按批次调价.批次, 0) = 0 Then
              Update 药品库存
              Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_零售金额
              Where 库房id = v_时价按批次调价.库房id And 药品id = v_时价按批次调价.药品id And 性质 = 1 And (批次 Is Null Or 批次 = 0);
            Else
              Update 药品库存
              Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_零售金额,
                  零售价 = Decode(n_时价分批, 1, v_时价按批次调价.零售价, 零售价)
              Where 库房id = v_时价按批次调价.库房id And 药品id = v_时价按批次调价.药品id And 性质 = 1 And 批次 = v_时价按批次调价.批次;
            End If;
          
            If Sql%RowCount = 0 Then
              Insert Into 药品库存
                (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 零售价)
              Values
                (v_时价按批次调价.库房id, v_时价按批次调价.药品id, v_时价按批次调价.批次, 1, 0, 0, n_零售金额, n_零售金额,
                 Decode(n_时价分批, 1, v_时价按批次调价.零售价, Null));
            End If;
          End If;
        End Loop;
      End Loop;
    End If;
  
    Update 药品收发记录 Set 审核人 = User, 审核日期 = Sysdate Where 价格id = 调价id_In;
    Update 收费价目 Set 变动原因 = 1 Where ID = 调价id_In;
  
    --更新药品目录、收费细目中的变价
    If 定价_In = 1 Then
      Update 收费项目目录 Set 是否变价 = 0 Where ID = n_收费细目id;
    End If;
    --成本价调价
    Zl_材料收发记录_成本价调价(n_收费细目id);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料收发记录_Adjust;
/

--93587:李业庆,2016-02-25,费用执行状态空值处理
Create Or Replace Procedure Zl_药品收发记录_处方发药
(
  Partid_In        In 药品收发记录.库房id%Type,
  Bill_In          In 药品收发记录.单据%Type,
  No_In            In 药品收发记录.No%Type,
  People_In        In 药品收发记录.审核人%Type,
  配药人_In        In 药品收发记录.配药人%Type := Null,
  校验人_In        In 药品收发记录.填制人%Type := Null,
  发药方式_In      In 药品收发记录.发药方式%Type := 1,
  发药时间_In      In 药品收发记录.审核日期%Type := Null,
  操作员编号_In    In 人员表.编号%Type := Null,
  操作员姓名_In    In 人员表.姓名%Type := Null,
  Intdigit_In      In Number := 2,
  Intautoverify_In In Number := 0,
  门诊_In          In Number := 1,
  核查人_In        In 药品收发记录.核查人%Type := Null,
  未取药_In        In 药品收发记录.是否未取药%Type := Null
) Is
  --住院病人
  Cursor c_Modifybillin Is
    Select a.Id, a.药品id, a.入出类别id, a.入出系数, a.费用id, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) 数量, Nvl(a.零售金额, 0) 金额, Nvl(a.批次, 0) 批次,
           a.供药单位id, a.成本价, a.批号, a.产地, a.效期, a.生产日期, a.批准文号, b.病人id, b.序号, Nvl(c.处方类型, Nvl(a.注册证号, 0)) 处方类型,
           Nvl(a.零售价, 0) As 零售价
    From 药品收发记录 A, 住院费用记录 B, 未发药品记录 C
    Where a.单据 = c.单据 And a.No = c.No And Nvl(a.库房id, 0) = Nvl(c.库房id, 0) And a.No = No_In And a.单据 = Bill_In And
          (a.库房id + 0 = Partid_In Or a.库房id Is Null) And Nvl(a.摘要, '小宝') <> '拒发' And a.费用id = b.Id And Nvl(b.执行状态,0) <> 1 And
          Mod(a.记录状态, 3) = 1 And a.审核人 Is Null;

  --门诊病人
  Cursor c_Modifybillout Is
    Select a.Id, a.药品id, a.入出类别id, a.入出系数, a.费用id, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) 数量, Nvl(a.零售金额, 0) 金额, Nvl(a.批次, 0) 批次,
           a.供药单位id, a.成本价, a.批号, a.产地, a.效期, a.生产日期, a.批准文号, b.病人id, b.序号, Nvl(c.处方类型, Nvl(a.注册证号, 0)) 处方类型,
           Nvl(a.零售价, 0) As 零售价, b.No, b.记录性质
    From 药品收发记录 A, 门诊费用记录 B, 未发药品记录 C
    Where a.单据 = c.单据 And a.No = c.No And Nvl(a.库房id, 0) = Nvl(c.库房id, 0) And a.No = No_In And a.单据 = Bill_In And
          (a.库房id + 0 = Partid_In Or a.库房id Is Null) And Nvl(a.摘要, '小宝') <> '拒发' And a.费用id = b.Id And Nvl(b.执行状态,0) <> 1 And
          Mod(a.记录状态, 3) = 1 And a.审核人 Is Null
    Order By 药品id;

  v_Modifybillin  c_Modifybillin%RowType;
  v_Modifybillout c_Modifybillout%RowType;

  --只读变量
  Dbl差价率  Number;
  v_核查日期 药品收发记录.核查日期%Type;
  --可写变量
  Dbl实际金额       药品收发记录.零售金额%Type;
  Dbl成本金额       药品收发记录.成本金额%Type;
  Dbl实际差价       药品收发记录.差价%Type;
  Date操作时间      药品收发记录.审核日期%Type;
  Bln收费与发药分离 Number(1);
  n_平均成本价      药品库存.平均成本价%Type;
  v_填制人          药品收发记录.填制人%Type;
  v_Error           Varchar2(4000);
  Err_Custom Exception;
Begin
  --取发药时间
  If 发药时间_In Is Null Then
    Select Sysdate Into Date操作时间 From Dual;
  Else
    Date操作时间 := 发药时间_In;
  End If;

  v_核查日期 := Date操作时间;
  Begin
    Select 0 Into Bln收费与发药分离 From 未发药品记录 Where 单据 = Bill_In And NO = No_In And 库房id + 0 = Partid_In;
  Exception
    When Others Then
      Bln收费与发药分离 := 1;
  End;

  --重写已发药处方的配药人
  Update 药品收发记录
  Set 配药人 = Decode(配药人_In, Null, 配药人, 配药人_In), 配药日期 = Decode(配药人_In, Null, 配药日期, Date操作时间)
  Where NO = No_In And 单据 = Bill_In And (库房id + 0 = Partid_In Or 库房id Is Null) And Mod(记录状态, 3) = 1 And 审核人 Is Not Null;

  Begin
    Select 填制人
    Into v_填制人
    From 药品收发记录
    Where NO = No_In And 单据 = Bill_In And 库房id + 0 = Partid_In And 审核日期 Is Null And Rownum = 1
    For Update Nowait;
  Exception
    When Others Then
      v_Error := '已有其他用户在执行发药，不能重复操作！';
      Raise Err_Custom;
  End;

  --重新计算成本价、成本金额、零售金额及差价
  If 门诊_In = 1 Then
    --处理门诊数据
    For v_Modifybillout In c_Modifybillout Loop
      n_平均成本价 := Round(Zl_Fun_Getoutcost(v_Modifybillout.药品id, v_Modifybillout.批次, Partid_In), 5);
      Dbl成本金额  := Round(n_平均成本价 * Nvl(v_Modifybillout.数量, 0), Intdigit_In);
      --零售金额
      Dbl实际金额 := Nvl(v_Modifybillout.金额, 0);
      --差价
      Dbl实际差价 := Round(Dbl实际金额 - Dbl成本金额, Intdigit_In);
    
      --更新药品收发记录的零售金额、成本金额、差价、审核人等信息
      Update 药品收发记录
      Set 成本价 = n_平均成本价, 成本金额 = Dbl成本金额, 差价 = Dbl实际差价, 库房id = Partid_In, 配药人 = Decode(配药人_In, Null, 配药人, 配药人_In),
          核查人 = 核查人_In, 核查日期 = v_核查日期, 配药日期 = Decode(配药人_In, Null, 配药日期, Date操作时间),
          填制人 = Decode(校验人_In, Null, 填制人, 校验人_In), 审核人 = Decode(People_In, Null, Zl_Username, People_In),
          审核日期 = Date操作时间, 发药方式 = 发药方式_In, 注册证号 = v_Modifybillout.处方类型, 是否未取药 = 未取药_In
      Where ID = v_Modifybillout.Id;
    
      If Bln收费与发药分离 = 1 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - Nvl(v_Modifybillout.数量, 0), 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybillout.数量, 0),
            实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybillout.金额, 0), 实际差价 = Nvl(实际差价, 0) - Dbl实际差价
        Where 库房id + 0 = Partid_In And 药品id = v_Modifybillout.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybillout.批次;
      Else
        Update 药品库存
        Set 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybillout.数量, 0), 实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybillout.金额, 0),
            实际差价 = Nvl(实际差价, 0) - Dbl实际差价
        Where 库房id + 0 = Partid_In And 药品id = v_Modifybillout.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybillout.批次;
      End If;
    
      If Sql%RowCount = 0 Then
        If Bln收费与发药分离 = 1 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号, 平均成本价)
          Values
            (Partid_In, v_Modifybillout.药品id, v_Modifybillout.批次, 1, 0 - Nvl(v_Modifybillout.数量, 0),
             0 - Nvl(v_Modifybillout.数量, 0), 0 - Nvl(v_Modifybillout.金额, 0), 0 - Dbl实际差价, v_Modifybillout.供药单位id,
             v_Modifybillout.成本价, v_Modifybillout.批号, v_Modifybillout.产地, v_Modifybillout.效期, v_Modifybillout.生产日期,
             v_Modifybillout.批准文号, v_Modifybillout.成本价);
        Else
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号, 平均成本价)
          Values
            (Partid_In, v_Modifybillout.药品id, v_Modifybillout.批次, 1, 0 - Nvl(v_Modifybillout.数量, 0),
             0 - Nvl(v_Modifybillout.金额, 0), 0 - Dbl实际差价, v_Modifybillout.供药单位id, v_Modifybillout.成本价,
             v_Modifybillout.批号, v_Modifybillout.产地, v_Modifybillout.效期, v_Modifybillout.生产日期, v_Modifybillout.批准文号,
             v_Modifybillout.成本价);
        End If;
      End If;
    
      Delete 药品库存
      Where 库房id + 0 = Partid_In And 药品id = v_Modifybillout.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
            Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0 And 性质 = 1;
    
      --更新费用记录的执行状态(已执行)
      Update 门诊费用记录
      Set 执行状态 = 1, 执行人 = Decode(People_In, Null, Zl_Username, People_In), 执行部门id = Partid_In, 执行时间 = Date操作时间
      
      Where NO = v_Modifybillout.No And Mod(记录性质, 10) = v_Modifybillout.记录性质 And 记录状态 <> 2 And 序号 = v_Modifybillout.序号;
    
      --费用审核（重复审核也没有关系）
      If Intautoverify_In = 1 Then
        If Bill_In = 9 Then
          Zl_门诊记帐记录_Verify(No_In, 操作员编号_In, 操作员姓名_In, v_Modifybillout.序号, 发药时间_In);
        End If;
      End If;
    
      --处理调价修正
      Zl_药品收发记录_调价修正(v_Modifybillout.Id);
    End Loop;
  Else
    --处理住院数据
    For v_Modifybillin In c_Modifybillin Loop
      n_平均成本价 := Round(Zl_Fun_Getoutcost(v_Modifybillin.药品id, v_Modifybillin.批次, Partid_In), 5);
      Dbl成本金额  := Round(n_平均成本价 * Nvl(v_Modifybillin.数量, 0), Intdigit_In);
      --零售金额
      Dbl实际金额 := Nvl(v_Modifybillin.金额, 0);
      --差价
      Dbl实际差价 := Round(Dbl实际金额 - Dbl成本金额, Intdigit_In);
    
      --更新药品收发记录的零售金额、成本金额、差价、审核人等信息
      Update 药品收发记录
      Set 成本价 = n_平均成本价, 成本金额 = Dbl成本金额, 差价 = Dbl实际差价, 库房id = Partid_In, 配药人 = Decode(配药人_In, Null, 配药人, 配药人_In),
          核查人 = 核查人_In, 核查日期 = v_核查日期, 配药日期 = Decode(配药人_In, Null, 配药日期, Date操作时间),
          填制人 = Decode(校验人_In, Null, 填制人, 校验人_In), 审核人 = Decode(People_In, Null, Zl_Username, People_In),
          审核日期 = Date操作时间, 发药方式 = 发药方式_In, 注册证号 = v_Modifybillout.处方类型
      Where ID = v_Modifybillin.Id;
    
      If Bln收费与发药分离 = 1 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - Nvl(v_Modifybillin.数量, 0), 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybillin.数量, 0),
            实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybillin.金额, 0), 实际差价 = Nvl(实际差价, 0) - Dbl实际差价
        Where 库房id + 0 = Partid_In And 药品id = v_Modifybillin.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybillin.批次;
      Else
        Update 药品库存
        Set 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybillin.数量, 0), 实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybillin.金额, 0),
            实际差价 = Nvl(实际差价, 0) - Dbl实际差价
        Where 库房id + 0 = Partid_In And 药品id = v_Modifybillin.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybillin.批次;
      End If;
    
      If Sql%RowCount = 0 Then
        If Bln收费与发药分离 = 1 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号, 平均成本价)
          Values
            (Partid_In, v_Modifybillin.药品id, v_Modifybillin.批次, 1, 0 - Nvl(v_Modifybillin.数量, 0),
             0 - Nvl(v_Modifybillin.数量, 0), 0 - Nvl(v_Modifybillin.金额, 0), 0 - Dbl实际差价, v_Modifybillin.供药单位id,
             v_Modifybillin.成本价, v_Modifybillin.批号, v_Modifybillin.产地, v_Modifybillin.效期, v_Modifybillin.生产日期,
             v_Modifybillin.批准文号, v_Modifybillin.成本价);
        Else
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号, 平均成本价)
          Values
            (Partid_In, v_Modifybillin.药品id, v_Modifybillin.批次, 1, 0 - Nvl(v_Modifybillin.数量, 0),
             0 - Nvl(v_Modifybillin.金额, 0), 0 - Dbl实际差价, v_Modifybillin.供药单位id, v_Modifybillin.成本价, v_Modifybillin.批号,
             v_Modifybillin.产地, v_Modifybillin.效期, v_Modifybillin.生产日期, v_Modifybillin.批准文号, v_Modifybillin.成本价);
        End If;
      End If;
    
      Delete 药品库存
      Where 库房id + 0 = Partid_In And 药品id = v_Modifybillin.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
            Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0 And 性质 = 1;
    
      --更新费用记录的执行状态(已执行)
      Update 住院费用记录
      Set 执行状态 = 1, 执行人 = Decode(People_In, Null, Zl_Username, People_In), 执行部门id = Partid_In, 执行时间 = Date操作时间
      
      Where ID = v_Modifybillin.费用id;
    
      --费用审核（重复审核也没有关系）
      If Intautoverify_In = 1 Then
        If Bill_In = 9 Then
          Zl_住院记帐记录_Verify(No_In, 操作员编号_In, 操作员姓名_In, v_Modifybillin.序号, v_Modifybillin.病人id, 发药时间_In);
        End If;
      End If;
    
      --处理调价修正
      Zl_药品收发记录_调价修正(v_Modifybillin.Id);
    End Loop;
  End If;

  --更新或删除未发药品记录
  Delete 未发药品记录 Where NO = No_In And 单据 = Bill_In And (库房id + 0 = Partid_In Or 库房id Is Null);

  If Bill_In = 8 Then
    Begin
      --移动支付宝项目在发药后动态调用生成推送信息的过程
      Execute Immediate 'Begin zl_服务窗消息_发送(:1,:2); End;'
        Using 6, No_In || ',' || Partid_In;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_处方发药;
/

--93587:李业庆,2016-02-25,费用执行状态空值处理
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
          Nvl(a.摘要, '小宝') <> '拒发' And a.费用id = b.Id And Nvl(b.执行状态,0) <> 1 And a.审核人 Is Null;

  Cursor c_Modifybillin Is
    Select a.Id, a.药品id, a.入出类别id, a.入出系数, a.费用id, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) 数量, Nvl(a.零售金额, 0) 金额, Nvl(a.批次, 0) 批次,
           a.供药单位id, a.成本价, a.批号, a.产地, a.效期, a.生产日期, a.批准文号
    From 药品收发记录 a, 住院费用记录 b
    Where a.No = No_In And a.单据 = Bill_In And (a.库房id + 0 = Otherstockid_In Or a.库房id Is Null) And
          Nvl(a.摘要, '小宝') <> '拒发' And a.费用id = b.Id And Nvl(b.执行状态,0) <> 1 And a.审核人 Is Null;

  --用于修正病人未结费用
  Cursor c_Billout Is
    Select b.实收金额, b.病人id, 0 主页id, 0 病人病区id, b.病人科室id, b.开单部门id, b.执行部门id, b.收入项目id, b.门诊标志
    From 药品收发记录 a, 门诊费用记录 b
    Where a.费用id = b.Id And Nvl(b.执行状态,0) <> 1 And a.No = No_In And a.单据 = Bill_In And
          (a.库房id + 0 = Otherstockid_In Or a.库房id Is Null) And Nvl(a.摘要, '小宝') <> '拒发' And a.审核人 Is Null And b.记录性质 = 2 And
          b.记录状态 = 1;

  Cursor c_Billin Is
    Select b.实收金额, b.病人id, b.主页id, b.病人病区id, b.病人科室id, b.开单部门id, b.执行部门id, b.收入项目id, b.门诊标志
    From 药品收发记录 a, 住院费用记录 b
    Where a.费用id = b.Id And Nvl(b.执行状态,0) <> 1 And a.No = No_In And a.单据 = Bill_In And
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
--93706:余伟节,2016-03-09,计算日期为2月最后一天时计算年龄天数超过31天的特殊处理
Create Or Replace Function Zl_Age_Calc
(
  病人id_In   病人信息.病人id%Type,
  出生日期_In Date := Null,
  计算日期_In Date := Null
) Return Varchar2
--功能:根据出生日期计算年龄.当天登记病人,保持年龄不变.
  --返回:1天以内：X小时[X分钟],1天至1月以内：X天,1月至1岁以内：X月[X天],1岁至儿童年龄上限：X岁[X月],>=儿童年龄上限：X岁
  --说明:1天以内，是指按出生日期24小时算;1月以内，是指对天计算；比如7.8日出生，8.8日才算1月;1岁以内，也是对天计算。;“以内”都是指“<”。
 As
  d_出生日期      Date;
  d_计算日期      Date;
  n_Days          Number;
  n_Months        Number;
  v_年龄          病人信息.年龄%Type;
  n_Upperagelimit Number; --参数:年龄上限

  v_Return Varchar2(20); --由于病人信息等相关表的年龄字段为10个字符，所以最大允许10个字符或5个汉字
Begin
  --当天登记的病人不用重算年龄
  If Nvl(病人id_In, 0) <> 0 Then
    Begin
      Select 年龄
      Into v_年龄
      From 病人信息
      Where 病人id = 病人id_In And Floor(Sysdate - 登记时间) = 0 And 年龄 Is Not Null;
    Exception
      When Others Then
        Null;
    End;
    If v_年龄 Is Not Null Then
      v_Return := v_年龄;
      Return v_Return;
    End If;
  End If;

  If 出生日期_In Is Null Then
    If Nvl(病人id_In, 0) <> 0 Then
      Select 出生日期 Into d_出生日期 From 病人信息 Where 病人id = 病人id_In;
    End If;
    If d_出生日期 Is Null Then
      Return Null;
    End If;
  Else
    d_出生日期 := 出生日期_In;
  End If;
  If 计算日期_In Is Null Then
    Select Sysdate Into d_计算日期 From Dual;
  Else
    d_计算日期 := 计算日期_In;
  End If;
  --如果出生日期大于计算日期,则直接为0小时
  If (d_计算日期 - d_出生日期) < 0 Then
    v_Return := '0小时';
    Return v_Return;
  End If;
  --获取儿童年龄的上限
  Begin
    Select Nvl(参数值, 14)
    Into n_Upperagelimit
    From Zlparameters
    Where 系统 = 100 And Nvl(模块, 0) = 0 And 参数号 = 147;
  Exception
    When Others Then
      n_Upperagelimit := 14;
  End;

  n_Months := Trunc(Months_Between(d_计算日期, d_出生日期));
  If n_Months < 12 * n_Upperagelimit Then
    --小于1岁的情况
    If n_Months < 12 Then
      --小于1月
      If n_Months < 1 Then
        n_Days := Trunc(d_计算日期 - d_出生日期);
        --一天以内
        If n_Days = 0 Then
          n_Days := Trunc((d_计算日期 - d_出生日期) * 24 * 60);
          If Mod(n_Days, 60) = 0 Then
            v_Return := n_Days / 60 || '小时';
          Else
            v_Return := Floor(n_Days / 60) || '小时' || Mod(n_Days, 60) || '分钟';
          End If;
        Else
          --一天至一月
          v_Return := n_Days || '天';
        End If;
      Else
        --大于1月
        n_Days := Trunc(Add_Months(d_计算日期, -1 * n_Months) - d_出生日期);
	If n_Days >= 31 Then
          --针对计算日期是2月份最后一天,出生日期刚好大于2月份最后一天且当天不是本月的最后一天
          --如：计算日期：2016-02-29   出生日期：2015-01-30
          n_Months := n_Months + 1;
          n_Days   := n_Days - 31;
        End If;
        If n_Days = 0 Then
          v_Return := n_Months || '月';
        Else
          v_Return := n_Months || '月' || n_Days || '天';
        End If;
      End If;
    Else
      --1岁到小于婴儿年龄上限的情况
      If Mod(n_Months, 12) = 0 Then
        v_Return := n_Months / 12 || '岁';
      Else
        v_Return := Floor(n_Months / 12) || '岁' || Mod(n_Months, 12) || '月';
      End If;
    End If;
  Else
    --大于等于婴儿年龄上限(直接X岁)
    v_Return := Floor(n_Months / 12) || '岁';
  End If;
  Return v_Return;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Age_Calc;
/
---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
Delete from zlFilesUpgrade Where UPPER(文件名)='REGCOM.DLL';
Insert Into zlFilesUpgrade
  (序号, 文件类型, 文件名, 版本号, 修改日期, 安装路径, 文件说明, 强制覆盖, 自动注册, 加入日期)
  Select Max(序号) + 1, 4, 'REGCOM.DLL', '', '', '[System]', '部件注册文件', 1, 1, Sysdate
  From zlFilesUpgrade
  Where Upper(文件名) = 'REGCOM.DLL';

Delete from zlFilesUpgrade Where UPPER(文件名)='ZL9PACSIMAGECAP.DLL';
Insert Into zlFilesUpgrade (文件类型,文件名,版本号,修改日期,所属系统,业务部件,安装路径,文件说明,强制覆盖,自动注册,加入日期,序号) select 1,'ZL9PACSIMAGECAP.DLL','', Null ,'1','ZL9PACSWORK','[Appsoft]\Apply','','0','1',sysdate,序号 from Dual a,(Select max(to_number(序号))+1 序号 from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(文件名)='ZL9PACSIMAGECAP.DLL');

Delete from zlFilesUpgrade Where UPPER(文件名)='ZL9XWINTERFACE.DLL';
Insert Into zlFilesUpgrade (文件类型,文件名,版本号,修改日期,所属系统,业务部件,安装路径,文件说明,强制覆盖,自动注册,加入日期,序号) select 0,'ZL9XWINTERFACE.DLL','', Null ,'1','','[Appsoft]\Apply','','0','1',sysdate,序号 from Dual a,(Select max(to_number(序号))+1 序号 from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(文件名)='ZL9XWINTERFACE.DLL');

--系统版本号
Update zlSystems Set 版本号='10.34.70' Where 编号=&n_System;







--部件版本号
Commit;