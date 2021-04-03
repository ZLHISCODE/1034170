--[连续升级]1
--[管理工具版本号]10.34.160
--本脚本支持从ZLHIS+ v10.34.150 升级到 v10.34.160
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--128157:蒋敏,2018-07-04,诊疗项目部位添加外键索引
Create Index 诊疗项目部位_IX_部位 on 诊疗项目部位(部位,类型) Tablespace zl9indexhis;

--119477:陈刘,2018-06-11,分类汇总之后,再次分组汇总到体温单

Alter Table 护理记录项目 add 分组汇总 VARCHAR2(100);

--126057:刘涛,2018-06-06,药品收发主表增加字段“对方库房id”
Alter Table 药品收发主表 Add 对方部门id Number(18);

Alter Table 药品收发主表 Drop Constraint 药品收发主表_UQ_NO Cascade drop Index;

Alter Table 药品收发主表 Add Constraint 药品收发主表_UQ_NO Unique (NO, 单据, 库房ID, 对方部门ID) Using Index Tablespace zl9Indexcis;

--126017:蒋敏,2018-05-23,诊疗项目管理检查部位方法的选择处理
Alter Table 诊疗项目部位 Add 上级方法 Varchar2(30);
Alter Table 诊疗项目部位 Drop Constraint 诊疗项目部位_UQ Cascade Drop Index;
Alter Table 诊疗项目部位 Add Constraint 诊疗项目部位_UQ_项目id Unique(项目id,部位,方法,类型,上级方法)Using Index Tablespace Zl9indexhis;

--124269:胡俊勇,2018-05-07,基于会诊申请下达医嘱标记
Alter Table 病人医嘱记录 Add 会诊医嘱ID number(18);  

--120692:刘鹏飞,2018-04-17,护理记录支持检验项目导入
create table 护理内容导入定义
(
类别 number(1),
名称 varchar2(100),
格式 varchar2(500)
)tablespace zl9BaseItem;
alter table 护理内容导入定义 add constraint 护理内容导入定义_PK primary key (类别) using index tablespace zl9Indexhis;

--111037:余伟节,2018-06-24,新生儿登记允许录入死亡时间
Alter Table 病人新生儿记录 Add 死亡时间 Date;

--124487:杨周一,2018-04-18,加大部分表和索引的事务量(执行后对于已经开辟的数据段无效,如需对已有数据(段)生效,需要重建索引或Move表)
Declare
  Cursor c_Sql Is
    Select 'Alter table ' || Table_Name || ' Initrans 20' Executesql
    From User_Tables
    Where Ini_Trans = 1 And
          Table_Name In (Select /*+ cardinality(a,10)*/
                          Upper(Column_Value) Tblname
                         From Table(f_Str2list('电子病历记录,电子病历格式,电子病历附件,电子病历内容,电子病历图形,电子病历打印,挂号序号状态', ',')) A)
    Union All
    Select Distinct 'Alter index ' || Index_Name || ' Initrans 20' Executesql
    From User_Indexes
    Where Ini_Trans = 2 And Index_Type = 'NORMAL' And
          Table_Name In (Select /*+ cardinality(a,10)*/
                          Upper(Column_Value)
                         From Table(f_Str2list('电子病历记录,电子病历格式,电子病历附件,电子病历内容,电子病历图形,电子病历打印,挂号序号状态', ',')) A);
  c_Row c_Sql%RowType;
Begin
  For c_Row In c_Sql Loop
    Execute Immediate c_Row.Executesql;
  End Loop;
End;
/

-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--127720:殷瑞,2018-06-29,处方发药定价分批药品，允许自动切换批次
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1341, 0, 0, 0, 0, 75, '发药批次更换', '0', '0',
         '0-不启用,1-启用。启用后，定价分批药品当库存为严格检查时，且发药时该药品的实际数量不足，则自动寻找库存足够的其他批次并替换更新'
  From Dual;

--127487:刘涛,2018-06-25,分批卫材入库是否检查批号产地
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, 0, 0, 0, 0, 305, '分批卫材批号产地控制', '1', '1',
         '启用该参数后,在涉及到卫材入库的地方检查分批卫材是否录入了批号和产地；0-不检查分批卫材是否录入了批号和产地；1-检查分批卫材是否录入了批号和产地。'
  From Dual;

--127480:焦博,2018-06-20,修正Oracle过程Zl_挂号安排_Autoupdate
Update 挂号安排计划 Set 实际生效 = To_Date('3000-01-01', 'yyyy-mm-dd') Where 生效时间 > Sysdate And 实际生效 < Sysdate;

--124184:陈龙,2018-04-13,新增血库业务消息
insert into 业务消息类型(编码,名称,说明,保留天数) values('ZLHIS_BLOOD_008','提交输血反应','医生填写并提交了输血反应时，产生该消息',7);

--124195:陈龙,2018-04-13,新增血库业务消息
insert into 业务消息类型(编码,名称,说明,保留天数) values('ZLHIS_BLOOD_005','血液审核完成','血液审核完成后，血液处于待发状态,如果超过预定输血日期，提示相应科室',7);

--124189:陈龙,2018-04-20,新增血库业务消息
insert into 业务消息类型(编码,名称,说明,保留天数) values('ZLHIS_BLOOD_006','出现输血反应','护士执行输血时，出现输血反应，提示相应医生站',7);

--124187:陈龙,2018-04-24,新增血库业务消息
insert into 业务消息类型(编码,名称,说明,保留天数) values('ZLHIS_BLOOD_007','血袋回收提示','护士执行完成后，提示护士站或医技站回收血袋',7);

--127049:王煜,2018-06-12,针对年龄段优惠
Update Zlparameters
Set 参数名 = '加收附加费', 参数说明 = '在自助挂号的时候，判断是否在挂号费用的基础上加收附加费'
Where 系统 = &n_System And 模块 = 1802 And 参数名 = '加收药事服务费';

Update Zlparameters
Set 参数名 = '加收附加费',参数说明 = '在自助预约的时候，判断是否在挂号费用的基础上加收附加费'
Where 系统 = &n_System And 模块 = 1803 And 参数名 = '加收药事服务费';

--126057:刘涛,2018-06-06,药品收发主表增加字段“对方库房id”移库申领数据修正
Declare
  v_No         药品收发主表.No%Type;
  n_单据       药品收发主表.单据%Type;
  n_库房id     药品收发主表.库房id%Type;
  n_对方部门id 药品收发主表.库房id%Type;

  Cursor c_药品收发主表 Is
    Select a.No, a.单据, a.库房id From 药品收发主表 A Where a.单据 = 19;
Begin
  For r_药品收发主表 In c_药品收发主表 Loop
    Begin
      Select 对方部门id
      Into n_对方部门id
      From 药品收发记录
      Where NO = r_药品收发主表.No And 库房id = r_药品收发主表.库房id And 单据 = r_药品收发主表.单据 And 入出系数 = -1 And Rownum < 2;
    Exception
      When Others Then
        n_对方部门id := 0;
    End;
  
    Update 药品收发主表
    Set 对方部门id = n_对方部门id
    Where NO = r_药品收发主表.No And 单据 = r_药品收发主表.单据 And 库房id = r_药品收发主表.库房id;

    Commit;
  End Loop;
End;
/

--124269:胡俊勇,2018-05-07,基于会诊申请下达医嘱标记
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, 0, 0, 0, 0, 302, '会诊科室下达医嘱由会诊申请科室处理', '0',
         '0', '0-表示不启用,1-表示启用,若启用参数，则会诊科室下的医嘱必须由会诊申请科室的护士进行校对或者发送' 
  From Dual;

--124273:胡俊勇,2018-05-07,存在未发送医嘱时禁止处理转科医嘱
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1254, 0, 0, 0, 0, 85, '存在未发送医嘱时禁止处理转科医嘱', '0', '0',
         '0-表示不启用,1-表示启用,若启用参数，存在可以发送的长期医嘱时就会禁止校对或者发送转科医嘱，本参数在判断的时候只判断长嘱，另外可以结合(特殊医嘱发送前检查未生效医嘱)参数配合使用'
  From Dual;    

--124467:殷瑞,2018-04-23,处方发药模块新增自动配药的规则控制
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1341, 0, 1, 0, 0, 70, '自动配药规则', '0', '0',
         '0-全部处方自动配药;1-电子处方(有医嘱的处方)自动配药;2-手工处方(无医嘱的处方)自动配药'
  From Dual;

--124699:张永康,2018-04-23,历史数据转出在离线模式下应保留的索引，统一补充
Insert Into zlBakTableindex(系统,表名,索引名) Select 100,'病人预交记录','病人预交记录_IX_主页ID' From Dual;

Insert Into zlBakTableindex(系统,表名,索引名) Select 100,'门诊费用记录','门诊费用记录_IX_病人ID' From Dual;

Insert Into zlBakTableindex(系统,表名,索引名) Select 100,'病人医嘱记录','病人医嘱记录_IX_开嘱时间' From Dual;

Insert Into zlBakTableindex(系统,表名,索引名) Select 100,'疾病阳性记录','疾病阳性记录_IX_医嘱ID' From Dual;

Insert Into zlBakTableindex(系统,表名,索引名) Select 100,'检验流水线标本','检验流水线标本_IX_标本ID' From Dual;

Insert Into zlBakTableindex(系统,表名,索引名) Select 100,'检验流水线指标','检验流水线指标_IX_标本ID' From Dual;

--120692:刘鹏飞,2018-04-17,护理记录支持检验项目导入
Insert into zlTables(系统,表名,表空间,分类) Values(&n_System,'护理内容导入定义','ZL9BASEITEM','A2');


--123971:刘鹏飞,2018-04-18,增加参数控制血液接收后才允许执行登记
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值,参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -Null, 0, 0, 0, 0, 301, '血液接收后才允许执行登记', '1', '1','启用血库系统时医护人员取血回室后，是否需要进行血液接收核对环节才允许进行输血执行情况登记：0-无需进行接收环节即可进行执行情况登记,1-必须进行血液接收核对环节才允许进行执行情况登记'
  From Dual;
-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--119477:陈刘,2018-07-09,分类汇总之后,再次分组汇总到体温单
Insert Into Zlprogprivs(系统, 序号, 功能, 所有者, 对象, 权限)Values(&n_System, 1255, '护理记录审签', User, 'Zl_护理二次汇总_Update', 'EXECUTE');

--127340:殷瑞,2018-06-29,部门发药基础权限新增对象
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1342,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'Zl_Gettransexenumber','EXECUTE' From Dual) A;

--121712:蒋敏,2018-05-21,诊疗项目管理分开授权
Insert Into Zlprogfuncs (系统, 序号, 功能, 排列, 说明, 缺省值) Values (&n_System, 1054, '诊疗项目编辑', 5, '增加、删除、修改诊疗项目的操作权限。有该权限时，允许对分类及诊疗项目进行增加、删除、修改、启用、停用，并允许设置检查部位、采集方式、标本对照、排斥关系、对应单据', 1);
Insert Into Zlprogfuncs (系统, 序号, 功能, 排列, 说明, 缺省值) Values (&n_System, 1054, '中药配方编辑', 7, '增加、删除、修改中药配方的操作权限。有该权限时，允许对分类及中药配方进行增加、删除、修改、启用、停用', 1);
Insert Into Zlprogfuncs (系统, 序号, 功能, 排列, 说明, 缺省值) Values (&n_System, 1054, '成套方案编辑', 11, '增加、删除、修改成套方案的操作权限。有该权限时，允许对分类及成套方案进行增加、删除、修改、启用、停用', 1);
Update Zlprogfuncs Set 排列 = 2 Where 系统 = &n_System And 序号 = 1054 And 功能 = '参数设置';
Update Zlprogfuncs Set 排列 = 6 Where 系统 = &n_System And 序号 = 1054 And 功能 = '管理中药配方';
Update Zlprogfuncs Set 排列 = 8 Where 系统 = &n_System And 序号 = 1054 And 功能 = '管理成套方案';
Update Zlprogfuncs Set 排列 = 9 Where 系统 = &n_System And 序号 = 1054 And 功能 = '全院成套方案';
Update Zlprogfuncs Set 排列 = 10 Where 系统 = &n_System And 序号 = 1054 And 功能 = '本科成套方案';
Update Zlprogfuncs Set 排列 = 12 Where 系统 = &n_System And 序号 = 1054 And 功能 = '修改全院成套方案';
Update Zlprogfuncs Set 排列 = 13 Where 系统 = &n_System And 序号 = 1054 And 功能 = '修改科室成套方案';
Update Zlprogfuncs Set 排列 = 14 Where 系统 = &n_System And 序号 = 1054 And 功能 = '修改个人成套方案';

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1054,'成套方案编辑',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_成套方案内容_Insert','EXECUTE' From Dual
Union All Select 'ZL_成套方案项目_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_诊疗分类目录_DELETE','EXECUTE' From Dual
Union All Select 'ZL_诊疗分类目录_Insert','EXECUTE' From Dual
Union All Select 'ZL_诊疗分类目录_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_DELETE','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_Insert','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_REUSE','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_STOP','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_所见项目_DELETE','EXECUTE' From Dual
Union All Select 'ZL_所见项目_Insert','EXECUTE' From Dual
Union All Select 'ZL_所见项目_UPDATE','EXECUTE' From Dual) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1054,'诊疗项目编辑',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_检查组合项目_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_检验报告项目_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_检验项目_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_配伍禁忌_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_所见项目_DELETE','EXECUTE' From Dual
Union All Select 'ZL_所见项目_Insert','EXECUTE' From Dual
Union All Select 'ZL_所见项目_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_用法用量_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_诊疗单据应用_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_诊疗分类目录_DELETE','EXECUTE' From Dual
Union All Select 'ZL_诊疗分类目录_Insert','EXECUTE' From Dual
Union All Select 'ZL_诊疗分类目录_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_诊疗互斥项目_SAVE','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_DELETE','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_Insert','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_REUSE','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_STOP','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_UPDATE','EXECUTE' From Dual
Union All Select 'Zl_诊疗项目部位_DELETE','EXECUTE' From Dual
Union All Select 'Zl_诊疗项目部位_Insert','EXECUTE' From Dual
Union All Select '检验项目','SELECT' From Dual
Union All Select '检验项目参考','SELECT' From Dual
Union All Select '诊疗分类目录_ID','SELECT' From Dual
Union All Select '诊疗项目目录_ID','SELECT' From Dual
Union All Select '诊治所见项目_ID','SELECT' From Dual) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1054,'中药配方编辑',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_中药配方_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_诊疗分类目录_DELETE','EXECUTE' From Dual
Union All Select 'ZL_诊疗分类目录_Insert','EXECUTE' From Dual
Union All Select 'ZL_诊疗分类目录_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_DELETE','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_Insert','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_REUSE','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_STOP','EXECUTE' From Dual
Union All Select 'ZL_诊疗项目_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_所见项目_DELETE','EXECUTE' From Dual
Union All Select 'ZL_所见项目_Insert','EXECUTE' From Dual
Union All Select 'ZL_所见项目_UPDATE','EXECUTE' From Dual) A;

Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1054,1,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '修改个人成套方案',2,0,0 From Dual) A;

Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1054,2,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '本科成套方案',2,1,0 From Dual
Union All Select '修改科室成套方案',2,0,0 From Dual) A;

Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1054,3,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '全院成套方案',2,1,0 From Dual
Union All Select '修改全院成套方案',2,0,0 From Dual) A;

Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1054,4,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '管理诊疗项目',2,1,0 From Dual
Union All Select '诊疗项目编辑',2,0,0 From Dual) A;

Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1054,5,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '管理中药配方',2,1,0 From Dual
Union All Select '中药配方编辑',2,0,0 From Dual) A;

Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1054,6,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '管理成套方案',2,1,0 From Dual
Union All Select '成套方案编辑',2,0,0 From Dual) A;

Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1054,7,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '成套方案编辑',2,1,0 From Dual
Union All Select '修改个人成套方案',2,0,0 From Dual
Union All Select '修改科室成套方案',2,0,0 From Dual
Union All Select '修改全院成套方案',2,0,0 From Dual) A;

Insert Into Zlrolegrant
  (系统, 序号, 角色, 功能)
  Select 系统, 序号, 角色, '诊疗项目编辑'
  From Zlrolegrant
  Where 系统 = &n_System And 序号 = 1054 And 功能 = '项目编辑';
Insert Into Zlrolegrant
  (系统, 序号, 角色, 功能)
  Select 系统, 序号, 角色, '中药配方编辑'
  From Zlrolegrant
  Where 系统 = &n_System And 序号 = 1054 And 功能 = '项目编辑';
Insert Into Zlrolegrant
  (系统, 序号, 角色, 功能)
  Select 系统, 序号, 角色, '成套方案编辑'
  From Zlrolegrant
  Where 系统 = &n_System And 序号 = 1054 And 功能 = '项目编辑';

Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select &n_System,1252,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0
Union All Select NULL,&n_System,1054,0,'成套方案编辑',1 From Dual) A;

Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select &n_System,1253,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0
Union All Select NULL,&n_System,1054,0,'成套方案编辑',1 From Dual) A;

Delete Zlmodulerelas Where 系统 = &n_System And 模块 = 1252 And 功能 Is Null And 相关系统 = &n_System And 相关模块 = 1054 And 相关功能 = '项目编辑';
Delete Zlmodulerelas Where 系统 = &n_System And 模块 = 1253 And 功能 Is Null And 相关系统 = &n_System And 相关模块 = 1054 And 相关功能 = '项目编辑';
Delete Zlprogfuncs Where 系统 = &n_System And 序号 = 1054 And 功能 = '项目编辑';

--120692:刘鹏飞,2018-04-17,护理记录支持检验项目导入
Insert Into Zlprogprivs(系统, 序号, 功能, 所有者, 对象, 权限)Values(&n_System, 1255, '基本', User, '护理内容导入定义', 'SELECT');
Insert Into Zlprogprivs(系统, 序号, 功能, 所有者, 对象, 权限) Values (&n_System, 1255, '护理记录登记', User, 'Zl_护理内容导入定义_Update', 'EXECUTE');
--124418:王振涛,2018-04-17,三方LIS权限
DELETE zlProgFuncs WHERE 系统 = &n_System And 序号 = 1215 ;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1215,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0
Union All Select '基本',-NULL,NULL,1 From Dual
Union All Select '核收标本',1,'核收检验申请单,并确定检验人及检验时间。',1 From Dual
Union All Select '核收撤消',2,'是否可以撤消已经核收的标本。',1 From Dual
Union All Select '审核标本',3,'对已经检验的标本进行审核确认。',1 From Dual
Union All Select '未收费审核',4,'能够审核未收取检验相关费用的检验单。',1 From Dual
Union All Select '审核取消',5,'对已经审核了的标本进行撤消处理。',1 From Dual
Union All Select '已审已打印可回滚',6,'有此权限，则可以回滚已审核并且已打印的报告。',1 From Dual
Union All Select '记帐检查余额',-NULL,NULL,1 From Dual) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1215,'核收标本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_检验标本记录_标本核收','EXECUTE' From Dual
Union All Select 'Zl_检验普通结果_Write','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1215,'核收撤消',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'Zl_检验标本记录_取消核收','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1215,'审核标本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_检验普通结果_BATCHUPDATE','EXECUTE' From Dual
Union All Select 'ZL_检验标本记录_报告审核','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1215,'未收费审核',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_检验普通结果_BATCHUPDATE','EXECUTE' From Dual
Union All Select 'ZL_检验标本记录_报告审核','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1215,'审核取消',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_检验标本记录_审核取消','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1215,'已审已打印可回滚',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_检验标本记录_审核取消','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1215,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'Zl_检验医嘱标记_Edit','EXECUTE' From Dual
Union All Select 'Zl_住院记帐记录_Verify','EXECUTE' From Dual
Union All Select 'Zl_检验报告单_Insert','EXECUTE' From Dual
Union All Select 'Zl_电子病历格式_Insert','EXECUTE' From Dual
Union All Select '部门表','SELECT' From Dual
Union All Select '药品卫材精度','SELECT' From Dual
Union All Select '药品库存','SELECT' From Dual
Union All Select '病人未结费用','SELECT' From Dual
Union All Select '病人余额','SELECT' From Dual
Union All Select '药品收发记录','SELECT' From Dual
Union All Select '检验流水线标本','SELECT' From Dual
Union All Select '检验流水线指标','SELECT' From Dual
Union All Select '检验试剂记录','SELECT' From Dual
Union All Select '检验仪器项目','SELECT' From Dual
Union All Select '检验细菌','SELECT' From Dual
Union All Select '检验抗生素用药','SELECT' From Dual
Union All Select '检验细菌抗生素','SELECT' From Dual
Union All Select '检验药敏结果','SELECT' From Dual
Union All Select '检验普通结果','SELECT' From Dual
Union All Select '检验申请项目','SELECT' From Dual
Union All Select '检验合并规则','SELECT' From Dual
Union All Select '检验操作记录','SELECT' From Dual
Union All Select '检验报告项目','SELECT' From Dual
Union All Select '病人新生儿记录','SELECT' From Dual
Union All Select '材料特性','SELECT' From Dual
Union All Select '未发药品记录','SELECT' From Dual
Union All Select '检验项目分布','SELECT' From Dual
Union All Select '病人医嘱附费','SELECT' From Dual
Union All Select '诊疗项目目录','SELECT' From Dual
Union All Select '病案主页','SELECT' From Dual
Union All Select '病人信息','SELECT' From Dual
Union All Select '住院费用记录','SELECT' From Dual
Union All Select '门诊费用记录','SELECT' From Dual
Union All Select '电子病历记录','SELECT' From Dual
Union All Select '病人医嘱报告','SELECT' From Dual
Union All Select '病历文件列表','SELECT' From Dual
Union All Select '检验标本记录','SELECT' From Dual
Union All Select '电子病历内容','SELECT' From Dual
Union All Select '人员表','SELECT' From Dual
Union All Select '部门人员','SELECT' From Dual
Union All Select '病人医嘱发送','SELECT' From Dual
Union All Select '病历单据应用','SELECT' From Dual
Union All Select '病人医嘱记录','SELECT' From Dual
Union All Select '电子病历格式','SELECT' From Dual
) A;






-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--128596:陈刘,2018-07-09,无法批量录入婴儿体温单
CREATE OR REPLACE Procedure Zl_体温单骑线设置_Update
(
  文件id_In   In 病人护理文件.Id%Type, --病人护理文件ID
  发生时间_In In 病人护理数据.发生时间%Type,
  项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
  记录id_In   In 病人护理数据.Id%Type,
  编辑_In     In Number := 0
) As
  n_子类         病历文件列表.种类%Type;
  n_开始时点     Number;
  n_监测次数     Number;
  n_时间间隔     Number;
  v_入院时间     病人变动记录.开始时间%Type;
  Ncount         Number;
  v_记录日期     病人护理数据.发生时间%Type;
  v_体温记录     病人护理明细.记录内容%Type;
  v_开始时间     Varchar2(20);
  v_结束时间     Varchar2(20);
  v_中间时间     Varchar2(20);
  v_显示时间pre  Varchar2(20);
  v_显示时间next Varchar2(20);
  v_入院开始时间 Varchar2(20);
  v_入院结束时间 Varchar2(20);
  v_数值         病人护理明细.记录内容%Type;
  v_内容时间     Varchar2(20);
  n_骑线         Number(1);
  n_明细id       病人护理明细.Id%Type;
  v_Error        Varchar2(255);
  n_入院         Number(1);
  n_Time         Number(2);
  n_p            Number(2);
  Err_Custom Exception;
  --当前时段显示的体温数据
  Function f_Nowshow
  (
    开始时间_In   In Varchar2,
    结束时间_In   In Varchar2,
    中间时间_In   In Varchar2,
    Id_In         In 病人护理明细.Id%Type,
    护理文件id_In In 病人护理文件.Id%Type
  ) Return Varchar2 Is
    n_时间差   Number;
    n_显示     Number(1);
    v_记录内容 病人护理明细.记录内容%Type;
    v_时间     Varchar2(20);
  Begin
    n_时间差 := -1;
    For r_Temp In (Select g.发生时间, f.记录内容, f.显示, f.未记说明
                   From 病人护理文件 B, 病人护理数据 G, 病人护理明细 F
                   Where b.Id = g.文件id And g.Id = f.记录id And b.Id = 护理文件id_In And f.项目序号 = 1 And f.记录类型 = 1 And
                         f.记录标记 = 0 And g.发生时间 Between To_Date(开始时间_In, 'YYYY-MM-DD hh24:mi:ss') And
                         To_Date(结束时间_In, 'YYYY-MM-DD hh24:mi:ss') And f.Id <> Id_In
                   Order By g.发生时间) Loop
      If n_时间差 = -1 Then
        n_时间差   := Abs((r_Temp.发生时间 - To_Date(中间时间_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60);
        v_记录内容 := r_Temp.记录内容;
        n_显示     := r_Temp.显示;
        v_时间     := To_Char(r_Temp.发生时间, 'YYYY-MM-DD hh24:mi:ss');
      Else
        If r_Temp.显示 = 1 Then
          If n_显示 = 1 And Abs((r_Temp.发生时间 - To_Date(中间时间_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60) < n_时间差 Then
            n_时间差   := Abs((r_Temp.发生时间 - To_Date(中间时间_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60);
            v_记录内容 := r_Temp.记录内容;
            n_显示     := r_Temp.显示;
            v_时间     := To_Char(r_Temp.发生时间, 'YYYY-MM-DD hh24:mi:ss');
          Else
            n_时间差   := Abs((r_Temp.发生时间 - To_Date(中间时间_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60);
            v_记录内容 := r_Temp.记录内容;
            n_显示     := r_Temp.显示;
            v_时间     := To_Char(r_Temp.发生时间, 'YYYY-MM-DD hh24:mi:ss');
          End If;
        Else
          If Abs((r_Temp.发生时间 - To_Date(中间时间_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60) < n_时间差 And n_显示 = 0 Then
            n_时间差   := Abs((r_Temp.发生时间 - To_Date(中间时间_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60);
            v_记录内容 := r_Temp.记录内容;
            n_显示     := r_Temp.显示;
            v_时间     := To_Char(r_Temp.发生时间, 'YYYY-MM-DD hh24:mi:ss');
          End If;
        End If;

      End If;
      If r_Temp.未记说明 Is Not Null And r_Temp.记录内容 Is Null Then
        Return Null;
      End If;
    End Loop;
    If v_时间 Is Not Null Then
      Return v_时间 || '|' || v_记录内容;
    Else
      Return Null;
    End If;
  Exception
    When Others Then
      Return Null;
  End f_Nowshow;

Begin
  n_入院 := 0;
  If 项目序号_In <> 1 Then
    Return;
  End If;

  If 编辑_In = 1 Then
    Update 病人护理明细 Set 记录类型 = 1 Where 记录id = 记录id_In And 项目序号 = 项目序号_In;
  End If;

  Begin
    Select Max(a.子类)
    Into n_子类
    From 病历文件列表 A, 病人护理文件 B
    Where a.种类 = 3 And a.保留 <> 1 And a.Id = b.格式id And b.Id = 文件id_In;
  End;

  --查询入院时间
  If n_子类 = 1 Then
    For r_List In (Select c.要素名称, c.内容文本
                   From 病人护理文件 A, 病历文件结构 C, 病历文件结构 D
                   Where c.父id = d.Id And d.父id Is Null And d.对象序号 = 1 And a.格式id = c.文件id And a.Id = 文件id_In
                   Order By c.Id) Loop
      Case r_List.要素名称
        When '开始时点' Then
          n_开始时点 := To_Number(r_List.内容文本);
        When '监测次数' Then
          n_监测次数 := To_Number(r_List.内容文本);
        When '时间间隔' Then
          n_时间间隔 := To_Number(r_List.内容文本);
    Else
          v_Error := '';
      End Case;
    End Loop;
  Else
    n_开始时点 :=  Zl_To_Number(zl_GetSysParameter('体温开始时间', 1255)) ;
    n_监测次数 := 6;
    n_时间间隔 := 4;
  End If;

  Select Min(h.开始时间)
  Into v_入院时间
  From 病人变动记录 H, 病人护理文件 B
  Where h.开始时间 Is Not Null And h.病人id = b.病人id And h.主页id = b.主页id And b.Id = 文件id_In
  Group By h.病人id, h.主页id;

  v_记录日期 := To_Date(To_Char(v_入院时间, 'YYYY-MM-DD'), 'YYYY-MM-DD hh24:mi:ss');
  Ncount     := Floor(((v_入院时间 - v_记录日期) * 24 - n_开始时点) / n_时间间隔);

  If Ncount > n_监测次数 Then
    Ncount := n_监测次数;
  End If;
  v_入院开始时间 := To_Char(v_记录日期 + ((n_开始时点 + Ncount * n_时间间隔 - (n_时间间隔 / 4)) / 24), 'YYYY-MM-DD hh24:mi:ss');
  v_入院结束时间 := To_Char(v_记录日期 + ((n_开始时点 + Ncount * n_时间间隔 + (n_时间间隔 / 4)) / 24), 'YYYY-MM-DD hh24:mi:ss');
  If v_入院时间 <= To_Date(v_入院开始时间, 'YYYY-MM-DD hh24:mi:ss') And
     v_入院时间 >= To_Date(v_入院开始时间, 'YYYY-MM-DD hh24:mi:ss') - n_时间间隔 / 4 / 24 Then
    v_入院结束时间 := v_入院开始时间;
    v_入院开始时间 := To_Char(v_记录日期 + (n_开始时点 + Ncount * n_时间间隔) / 24, 'YYYY-MM-DD hh24:mi:ss');
    n_入院         := 1;
  Elsif v_入院时间 <= To_Date(v_入院开始时间, 'YYYY-MM-DD hh24:mi:ss') + n_时间间隔 / 24 And
        v_入院时间 >= To_Date(v_入院结束时间, 'YYYY-MM-DD hh24:mi:ss') Then
    v_入院开始时间 := v_入院结束时间;
    v_入院结束时间 := To_Char(v_记录日期 + ((n_开始时点 + (Ncount + 1) * n_时间间隔) / 24), 'YYYY-MM-DD hh24:mi:ss');
    n_入院         := 1;
  End If;

  v_记录日期 := To_Date(To_Char(发生时间_In, 'YYYY-MM-DD'), 'YYYY-MM-DD hh24:mi:ss');
  Ncount     := Floor(((发生时间_In - v_记录日期) * 24 - n_开始时点) / n_时间间隔);

  If Ncount > n_监测次数 Then
    Ncount := n_监测次数;
  End If;

  --当前数据所在时间断
  v_开始时间 := To_Char(v_记录日期 + ((n_开始时点 + Ncount * n_时间间隔) / 24), 'YYYY-MM-DD hh24:mi:ss');
  v_结束时间 := To_Char(v_记录日期 + ((n_开始时点 + n_时间间隔 * (Ncount + 1)) / 24), 'YYYY-MM-DD hh24:mi:ss');
  v_中间时间 := To_Char(To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss') +
                    (To_Date(v_结束时间, 'YYYY-MM-DD hh24:mi:ss') - To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss')) / 2,
                    'YYYY-MM-DD hh24:mi:ss');

  Select Max(f.记录内容), Max(f.Id)
  Into v_体温记录, n_明细id
  From 病人护理文件 B, 病人护理数据 G, 病人护理明细 F
  Where b.Id = g.文件id And g.Id = f.记录id And b.Id = 文件id_In And f.项目序号 = 1 And f.记录标记 = 0 And g.发生时间 = 发生时间_In;

  v_数值         := f_Nowshow(v_开始时间, v_结束时间, v_中间时间, n_明细id, 文件id_In);
  v_显示时间next := '';
  While v_显示时间next Is Null Loop
    If v_数值 Is Null Then
      If v_内容时间 Is Not Null Then
        v_内容时间 := To_Char(To_Date(v_中间时间, 'YYYY-MM-DD hh24:mi:ss') + n_时间间隔 / 24, 'YYYY-MM-DD hh24:mi:ss');
      Else
        v_内容时间 := v_中间时间;
      End If;
    Else
      n_p        := Instr(v_数值, '|');
      v_内容时间 := Substr(v_数值, 1, n_p - 1);
      v_数值     := Substr(v_数值, n_p + 1);
    End If;
    If To_Date(v_内容时间, 'YYYY-MM-DD hh24:mi:ss') < 发生时间_In Then
      v_数值 := f_Nowshow(To_Char(To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss') + n_时间间隔 / 24, 'YYYY-MM-DD hh24:mi:ss'),
                        To_Char(To_Date(v_结束时间, 'YYYY-MM-DD hh24:mi:ss') + n_时间间隔 / 24, 'YYYY-MM-DD hh24:mi:ss'),
                        To_Char(To_Date(v_中间时间, 'YYYY-MM-DD hh24:mi:ss') + n_时间间隔 / 24, 'YYYY-MM-DD hh24:mi:ss'), n_明细id,
                        文件id_In);
    Else
      v_显示时间next := v_内容时间;
    End If;
  End Loop;
  v_数值     := '';
  v_内容时间 := '';

  --循环查询当前时间之前的普通数据
  n_Time := 0;
  While n_Time * n_时间间隔 <= 24 Loop
    If To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss') < Trunc(To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss')) + n_开始时点 / 24 And
       To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss') <> Trunc(To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss')) Then
      v_结束时间 := To_Char(Trunc(To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss')) + n_开始时点 / 24, 'YYYY-MM-DD hh24:mi:ss');
      v_开始时间 := To_Char(Trunc(To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss')), 'YYYY-MM-DD hh24:mi:ss');
      v_中间时间 := To_Char(To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss') +
                        (To_Date(v_结束时间, 'YYYY-MM-DD hh24:mi:ss') - To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss')) / 2,
                        'YYYY-MM-DD hh24:mi:ss');
      v_数值     := f_Nowshow(v_开始时间, v_结束时间, v_中间时间, n_明细id, 文件id_In);
      v_结束时间 := To_Char(Trunc(To_Date(v_结束时间, 'YYYY-MM-DD hh24:mi:ss')), 'YYYY-MM-DD hh24:mi:ss');
      v_开始时间 := To_Char(To_Date(v_结束时间, 'YYYY-MM-DD hh24:mi:ss') - n_时间间隔 / 24, 'YYYY-MM-DD hh24:mi:ss');
      v_中间时间 := To_Char(To_Date(v_结束时间, 'YYYY-MM-DD hh24:mi:ss') - n_时间间隔 / 2 / 24, 'YYYY-MM-DD hh24:mi:ss');
    Else
      v_数值     := f_Nowshow(v_开始时间, v_结束时间, v_中间时间, n_明细id, 文件id_In);
      v_开始时间 := To_Char(To_Date(v_开始时间, 'YYYY-MM-DD hh24:mi:ss') - n_时间间隔 / 24, 'YYYY-MM-DD hh24:mi:ss');
      v_结束时间 := To_Char(To_Date(v_结束时间, 'YYYY-MM-DD hh24:mi:ss') - n_时间间隔 / 24, 'YYYY-MM-DD hh24:mi:ss');
      v_中间时间 := To_Char(To_Date(v_中间时间, 'YYYY-MM-DD hh24:mi:ss') - n_时间间隔 / 24, 'YYYY-MM-DD hh24:mi:ss');

    End If;
    n_Time := n_Time + 1;
    If v_数值 Is Not Null Then
      n_p           := Instr(v_数值, '|');
      v_内容时间    := Substr(v_数值, 1, n_p - 1);
      v_数值        := Substr(v_数值, n_p + 1);
      v_显示时间pre := v_内容时间;
      Exit When To_Date(v_内容时间, 'YYYY-MM-DD hh24:mi:ss') < 发生时间_In;
    End If;
  End Loop;
  If v_内容时间 Is Not Null Then
    If v_数值 < v_体温记录 And v_体温记录 > 37.5 Then
      Select Count(f.Id)
      Into n_骑线
      From 病人护理文件 B, 病人护理数据 G, 病人护理明细 F
      Where b.Id = g.文件id And g.Id = f.记录id And b.Id = 文件id_In And f.项目序号 = 项目序号_In And g.发生时间 <> 发生时间_In And
            f.记录类型 = 7 And g.发生时间 Between To_Date(v_显示时间pre, 'YYYY-MM-DD hh24:mi:ss') And
            To_Date(v_显示时间next, 'YYYY-MM-DD hh24:mi:ss');

      If n_骑线 < 1 Then
        Update 病人护理明细 Set 记录类型 = 7 Where 记录id = 记录id_In And 项目序号 = 项目序号_In And 终止版本 Is Null;
      End If;

    End If;
  Else
    If 发生时间_In >= To_Date(v_入院开始时间, 'YYYY-MM-DD hh24:mi:ss') And 发生时间_In <= To_Date(v_入院结束时间, 'YYYY-MM-DD hh24:mi:ss') And
       n_入院 = 1 Then
      Select Count(f.Id)
      Into n_骑线
      From 病人护理文件 B, 病人护理数据 G, 病人护理明细 F
      Where b.Id = g.文件id And g.Id = f.记录id And b.Id = 文件id_In And f.项目序号 = 项目序号_In And g.发生时间 <> 发生时间_In And
            f.记录类型 = 7 And g.发生时间 Between To_Date(v_入院开始时间, 'YYYY-MM-DD hh24:mi:ss') And
            To_Date(v_入院结束时间, 'YYYY-MM-DD hh24:mi:ss');

      If n_骑线 < 1 Then
        Update 病人护理明细 Set 记录类型 = 7 Where 记录id = 记录id_In And 项目序号 = 项目序号_In And 终止版本 Is Null;
      End If;
    End If;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_体温单骑线设置_Update;
/
--128391:陈龙,2018-07-05,增加审核状态为2的情况
CREATE OR REPLACE PROCEDURE Zl_医嘱审核管理_Update
(
  医嘱id_In   病人医嘱状态.医嘱id%TYPE,
  操作时间_In 病人医嘱状态.操作时间%TYPE,
  操作说明_In 病人医嘱状态.操作说明%TYPE := NULL,
  审核对象_In NUMBER := 1, --1=手术医嘱，2=输血医嘱 
  操作人员_In VARCHAR2 := NULL
) IS
  --修改只适用于审核不通过的医嘱，修改其审核说明 
  Err_Item EXCEPTION;
  v_Err_Msg  VARCHAR2(200);
  n_Count    NUMBER;
  n_审核状态 NUMBER;
BEGIN
  SELECT COUNT(1) INTO n_Count FROM 病人医嘱记录 WHERE Id = 医嘱id_In;
  SELECT 审核状态 INTO n_审核状态 FROM 病人医嘱记录 WHERE Id = 医嘱id_In;
  IF n_Count = 0 THEN
    v_Err_Msg := '有医嘱已经删除,请查证。';
    RAISE Err_Item;
  END IF;

  IF 审核对象_In = 1 THEN
    UPDATE 病人医嘱状态
    SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
    WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 12 AND
          操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 12);
  ELSE
    IF n_审核状态 = 1 THEN
      UPDATE 病人医嘱状态
      SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 19 AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 19);
      IF SQL%NOTFOUND THEN
        INSERT INTO 病人医嘱状态
          (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
          SELECT Id, 19, 操作人员_In, 操作时间_In, 操作说明_In
          FROM 病人医嘱记录
          WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In;
      END IF;
    ELSIF n_审核状态 = 7 THEN
      UPDATE 病人医嘱状态
      SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 18 AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 18);
      IF SQL%NOTFOUND THEN
        INSERT INTO 病人医嘱状态
          (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
          SELECT Id, 18, 操作人员_In, 操作时间_In, 操作说明_In
          FROM 病人医嘱记录
          WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In;
      END IF;
    ELSIF n_审核状态 = 3 THEN
      UPDATE 病人医嘱状态
      SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 12 AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 12);
      IF SQL%NOTFOUND THEN
        INSERT INTO 病人医嘱状态
          (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
          SELECT Id, 12, 操作人员_In, 操作时间_In, 操作说明_In
          FROM 病人医嘱记录
          WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In;
      END IF;
	ELSIF n_审核状态 = 4 OR n_审核状态 = 2 THEN
      UPDATE 病人医嘱状态
      SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 11 AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 11);
      IF SQL%NOTFOUND THEN
        INSERT INTO 病人医嘱状态
          (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
          SELECT Id, 11, 操作人员_In, 操作时间_In, 操作说明_In
          FROM 病人医嘱记录
          WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In;
      END IF;
    END IF;
  END IF;
EXCEPTION
  WHEN Err_Item THEN
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  WHEN OTHERS THEN
    Zl_Errorcenter(SQLCODE, SQLERRM);
END Zl_医嘱审核管理_Update;
/

--128391:陈龙,2018-07-05,增加传入的审核状态为2的情况
CREATE OR REPLACE PROCEDURE Zl_医嘱审核管理_Cancel
(
  医嘱ids_In  VARCHAR2,
  审核对象_In NUMBER := 1, --1=手术医嘱，2=输血医嘱
  执行类别_In NUMBER := 0 --0=老版血库流程；不为0时，则为目标审核状态：1=待审核；7=待签发；4-已签发；3-已拒绝；
) IS
  --取消审核
  CURSOR c_Advice IS
    SELECT * FROM TABLE(CAST(f_Num2list(医嘱ids_In) AS t_Numlist));
  Err_Item EXCEPTION;
  v_Err_Msg  VARCHAR2(200);
  n_Count    NUMBER;
  n_医嘱状态 NUMBER;
  n_审核状态 NUMBER;
  n_操作类型 NUMBER;
BEGIN
  FOR r_Advice IN c_Advice LOOP
    SELECT COUNT(1), MAX(医嘱状态), Nvl(MAX(审核状态), 0)
    INTO n_Count, n_医嘱状态, n_审核状态
    FROM 病人医嘱记录
    WHERE Id = r_Advice.Column_Value;
  
    IF n_Count = 0 THEN
      v_Err_Msg := '有医嘱已经删除,请查证。';
      RAISE Err_Item;
    END IF;
  
    IF n_医嘱状态 <> 1 THEN
      v_Err_Msg := '您选择的医嘱中包含有校对的医嘱，不能取消审核。';
      RAISE Err_Item;
    END IF;
  
    IF n_审核状态 = 1 THEN
      n_操作类型 := 19;
    ELSIF n_审核状态 = 7 THEN
      n_操作类型 := 18;
    ELSIF n_审核状态 = 3 THEN
      n_操作类型 := 12;
    ELSIF n_审核状态 = 4 OR n_审核状态 = 2 THEN
      n_操作类型 := 11;
    END IF;
  
    IF 审核对象_In = 1 OR 执行类别_In = 0 THEN
      UPDATE 病人医嘱记录 SET 审核状态 = 1 WHERE Id = r_Advice.Column_Value OR 相关id = r_Advice.Column_Value;
      DELETE FROM 病人医嘱状态
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = r_Advice.Column_Value OR 相关id = r_Advice.Column_Value) AND
            操作类型 IN (11, 12) AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = r_Advice.Column_Value AND 操作类型 IN (11, 12));
    ELSIF 审核对象_In = 2 AND 执行类别_In <> 0 THEN
      UPDATE 病人医嘱记录
      SET 审核状态 = 执行类别_In
      WHERE Id = r_Advice.Column_Value OR 相关id = r_Advice.Column_Value;
      DELETE FROM 病人医嘱状态
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = r_Advice.Column_Value OR 相关id = r_Advice.Column_Value) AND
            操作类型 = n_操作类型 AND
            操作时间 =
            (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = r_Advice.Column_Value AND 操作类型 = n_操作类型);
    END IF;
  
  END LOOP;
EXCEPTION
  WHEN Err_Item THEN
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  WHEN OTHERS THEN
    Zl_Errorcenter(SQLCODE, SQLERRM);
END Zl_医嘱审核管理_Cancel;
/

--128152:胡俊勇,2018-07-05,身份证号中日期校验
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
    v_Err_Msg := '传人身份证号为空!';
    Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
  Else
    --身份证合法验证 
    v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,91';
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
        Begin
          d_出生日期 := To_Date('19' || Substr(Idcard_In, 7, 6), 'yyyy-mm-dd');
        Exception
          When Others Then
            --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
            v_Temp := '19' || Substr(Idcard_In, 7, 6);
            If Instr(v_Temp || ',', '0229,') > 0 Then
              v_Temp := '19' || Substr(Idcard_In, 7, 5) || '8';
            End If;
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd');
        End;
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
        Begin
          d_出生日期 := To_Date(Substr(Idcard_In, 7, 8), 'yyyy-mm-dd');
        Exception
          When Others Then
            --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
            v_Temp := Substr(Idcard_In, 7, 8);
            If Instr(v_Temp || ',', '0229,') > 0 Then
              v_Temp := Substr(Idcard_In, 7, 7) || '8';
            End If;
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd');
        End;
      
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

--127817:董露露,2018-07-03,处理首页提取籍贯时数据提取错误的问题
CREATE OR REPLACE Function Zl_Adderss_Structure(v_Addressinfo Varchar2,n_Type Number :=Null) Return Varchar2 Is
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
      If n_Type is Null Then
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1) 
        Into v_市, v_Code市, n_虚拟, n_不显示, n_Count 
        From 区域 
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 1 And 上级编码 = v_Code省; 
      End If;
      --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配
      If n_Count > 1 Then
        v_Tmp := Substr(v_Adrstmp, 1, 3);
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
        Into v_市, v_Code市, n_虚拟, n_不显示
        From 区域
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 1 And 上级编码 = v_Code省;
      End If;
      --判断是否存在虚拟地址或不显示的地址导致的,如果存在，则根据第三级地址来确定虚拟地址
      --可能是没有第二级，因此需要第三级判断
      If v_Code市 Is Null Then
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

--125588:秦龙,2018-07-02,处理实际数量为0的卫材调价
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
        /*If Nvl(c_调价.填写数量, 0) = 0 And Nvl(c_调价.库存金额, 0) = 0 And Nvl(c_调价.库存差价, 0) = 0 Then
        Null;*/
        If Nvl(c_调价.填写数量, 0) = 0 And (Nvl(c_调价.库存金额, 0) <> 0 Or Nvl(c_调价.库存差价, 0) <> 0) Then
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

--125588:秦龙,2018-07-02,处理卫材实际数量为0的调价
Create Or Replace Procedure Zl_材料收发记录_成本价调价(材料id_In In 药品收发记录.药品id%Type) As
  v_No         药品收发记录.No%Type;
  v_应付id     应付记录.Id%Type; --应付记录的ID 
  v_应付单据号 应付记录.No%Type;
  d_调价时间   Date;
  n_序号       Number(8);
  n_库房id     药品收发记录.库房id%Type;
  n_入出类别id 药品收发记录.入出类别id%Type;
  n_入出系数   药品收发记录.入出系数%Type;
  n_收发id     药品收发记录.Id%Type;
  n_调整额     药品收发记录.零售金额%Type;
  n_原成本价   药品收发记录.成本价%Type;
  n_新成本价   药品收发记录.成本价%Type;
  n_平均成本价 药品库存.平均成本价%Type;
  v_调价id     成本价调价信息.Id%Type;
  v_调价汇总号 成本价调价信息.调价汇总号%Type;
  n_Count      Number(1) := 0;

  Cursor c_Stock Is --当前库存 
    Select 上次供应商id, a.库房id, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, a.上次批号, a.效期, a.上次产地, a.灭菌效期,
           Decode(Sign(Nvl(a.批次, 0)), 1, a.上次采购价, a.平均成本价) As 原成本价
    From 药品库存 A
    Where a.性质 = 1 And Nvl(a.实际数量, 0) <> 0 And a.药品id = 材料id_In
    Order By a.库房id;

  v_Stock c_Stock%RowType;
Begin
  d_调价时间 := Sysdate;
  n_库房id   := 0;

  --判断是否存在无库存调价 
  Begin
    Select ID, 新成本价, 调价汇总号
    Into v_调价id, n_新成本价, v_调价汇总号
    From 成本价调价信息
    Where 执行日期 Is Null And Nvl(库房id, 0) = 0 And 药品id = 材料id_In;
  Exception
    When Others Then
      v_调价id   := 0;
      n_新成本价 := Null;
  End;

  --无库存调价 
  If v_调价id > 0 Then
    --根据当前库存重新产生调价信息 
    For v_Stock In c_Stock Loop
      Zl_材料成本调价_Insert(v_Stock.上次供应商id, v_Stock.库房id, v_Stock.材料id, v_Stock.批次, v_Stock.上次批号, v_Stock.原成本价, n_新成本价,
                       Null, Null, 0, 0, v_调价汇总号);
      n_Count := n_Count + 1;
    End Loop;
  
    If n_Count > 0 Then
      --如果当前有库存记录，则删除无库存调价记录 
      Delete 成本价调价信息 Where ID = v_调价id;
    Else
      Update 成本价调价信息 Set 执行日期 = d_调价时间 Where ID = v_调价id;
    
      Update 材料特性 Set 成本价 = n_新成本价 Where 材料id = 材料id_In And 成本价 <> n_新成本价;
    End If;
  End If;

  --取库存差价调整的入出类别ID 
  Select b.Id, b.系数
  Into n_入出类别id, n_入出系数
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 33 And Rownum < 2;

  For c_成本调整 In (Select a.库房id, a.药品id As 材料id, Nvl(a.批次, 0) 批次, a.上次供应商id, a.实际数量, a.实际金额, a.实际差价, a.上次产地 As 产地,
                        a.上次批号 As 批号, a.灭菌效期, a.效期, a.上次生产日期 As 生产日期, a.批准文号, Nvl(a.平均成本价, 0) As 原成本价, b.新成本价, b.发票号,
                        b.发票日期, b.发票金额, Nvl(a.上次采购价, 0) As 上次采购价, b.Id As 调价id
                 From 药品库存 A, 成本价调价信息 B
                 Where a.药品id = b.药品id And Nvl(a.上次供应商id, 0) = Nvl(b.供药单位id, 0) And a.库房id = b.库房id And
                       Nvl(a.批次, 0) = Nvl(b.批次, 0) And a.性质 = 1 And b.执行日期 Is Null And a.药品id = 材料id_In
                 Order By a.库房id) Loop
    If n_库房id <> c_成本调整.库房id Then
      n_序号   := 1;
      n_库房id := c_成本调整.库房id;
      v_No     := Nextno(71, n_库房id);
    Else
      n_序号 := n_序号 + 1;
    End If;
  
    Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
  
    /*If Nvl(c_成本调整.实际数量, 0) = 0 And Nvl(c_成本调整.实际金额, 0) = 0 And Nvl(c_成本调整.实际差价, 0) = 0 Then
    --数量,金额、差价都为0，则表示数据是填单下可用数量出库产生的单据，此单据还没有审核，因此只需要更新调价信息，其他不更新
    Update 材料特性 Set 成本价 = c_成本调整.新成本价 Where 材料id = c_成本调整.材料id;
    
    Update 成本价调价信息
    Set 收发id = n_收发id, 执行日期 = d_调价时间, 效期 = c_成本调整.效期, 灭菌效期 = c_成本调整.灭菌效期, 产地 = c_成本调整.产地, 批号 = c_成本调整.批号
    Where ID = c_成本调整.调价id;*/
    If Nvl(c_成本调整.实际数量, 0) = 0 And (Nvl(c_成本调整.实际金额, 0) <> 0 Or Nvl(c_成本调整.实际差价, 0) <> 0) Then
      --数量=0 金额或差价<>0时只更新库存表中对应的平均成本价和特性表中成本价，并产生成本价修正数据但是差价差=0，只记录最新成本价
      --产生调价记录，只记录最新成本价
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期, 审核人,
         审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率)
      Values
        (n_收发id, 1, 18, v_No, n_序号, c_成本调整.库房id, n_入出类别id, c_成本调整.上次供应商id, n_入出系数, c_成本调整.材料id, c_成本调整.批次, c_成本调整.产地,
         c_成本调整.批号, c_成本调整.效期, 0, c_成本调整.实际金额, c_成本调整.实际差价, 0, '卫生材料成本价调价', Zl_Username, d_调价时间, Zl_Username, d_调价时间,
         c_成本调整.生产日期, c_成本调整.批准文号, c_成本调整.新成本价, 1, c_成本调整.原成本价);
      --更新库存      
      Update 药品库存
      Set 平均成本价 = c_成本调整.新成本价, 上次采购价 = c_成本调整.新成本价
      Where 库房id = c_成本调整.库房id And 药品id = c_成本调整.材料id And Nvl(批次, 0) = c_成本调整.批次 And 性质 = 1;
      Update 材料特性 Set 成本价 = c_成本调整.新成本价 Where 材料id = c_成本调整.材料id;
    
      Update 成本价调价信息
      Set 收发id = n_收发id, 执行日期 = d_调价时间, 效期 = c_成本调整.效期, 灭菌效期 = c_成本调整.灭菌效期, 产地 = c_成本调整.产地, 批号 = c_成本调整.批号
      Where ID = c_成本调整.调价id;
    Else
      --调整相应的库存:原成本金额-实新成本金额 
      n_调整额   := (c_成本调整.实际金额 - c_成本调整.实际差价) - Round(c_成本调整.新成本价 * c_成本调整.实际数量, 2);
      n_原成本价 := c_成本调整.原成本价;
    
      If n_原成本价 <= 0 Then
        n_原成本价 := c_成本调整.上次采购价;
      End If;
    
      --目前：收发记录对应: 
      -- 扣率--> 原成本价 
      -- 单量-->新成本价 
      -- 填写数量-->库存实际数量 
      -- 零售价-->库存实际金额 
      -- 成本价-->库存实际差价 
      -- 差价-->本次调整额 
    
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期, 审核人,
         审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率)
      Values
        (n_收发id, 1, 18, v_No, n_序号, c_成本调整.库房id, n_入出类别id, c_成本调整.上次供应商id, n_入出系数, c_成本调整.材料id, c_成本调整.批次, c_成本调整.产地,
         c_成本调整.批号, c_成本调整.效期, c_成本调整.实际数量, c_成本调整.实际金额, c_成本调整.实际差价, n_调整额, '卫生材料成本价调价', Zl_Username, d_调价时间,
         Zl_Username, d_调价时间, c_成本调整.生产日期, c_成本调整.批准文号, c_成本调整.新成本价, 1, n_原成本价);
    
      --更新库存 
      Update 药品库存
      Set 实际差价 = Nvl(实际差价, 0) + n_调整额
      Where 库房id = c_成本调整.库房id And 药品id = c_成本调整.材料id And Nvl(批次, 0) = Nvl(c_成本调整.批次, 0) And 性质 = 1;
    
      If Sql%NotFound Then
        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 实际差价, 上次批号, 效期, 上次产地, 上次供应商id, 上次生产日期, 批准文号, 灭菌效期)
        Values
          (c_成本调整.库房id, c_成本调整.材料id, c_成本调整.批次, 1, n_调整额, c_成本调整.批号, c_成本调整.效期, c_成本调整.产地, c_成本调整.上次供应商id, c_成本调整.生产日期,
           c_成本调整.批准文号, c_成本调整.灭菌效期);
      End If;
    
      Update 药品库存
      Set 上次采购价 = c_成本调整.新成本价
      Where 药品id = c_成本调整.材料id And 上次采购价 <> c_成本调整.新成本价;
    
      Update 材料特性
      Set 成本价 = c_成本调整.新成本价
      Where 材料id = c_成本调整.材料id And 成本价 <> c_成本调整.新成本价;
    
      --重新计算库存表中的平均成本价 
      Update 药品库存
      Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
      Where 药品id = c_成本调整.材料id And Nvl(批次, 0) = Nvl(c_成本调整.批次, 0) And 库房id = c_成本调整.库房id And 性质 = 1 And
            Nvl(实际数量, 0) <> 0;
      If Sql%NotFound Then
        Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = c_成本调整.材料id;
        Update 药品库存
        Set 平均成本价 = n_平均成本价
        Where 药品id = c_成本调整.材料id And 库房id = c_成本调整.库房id And Nvl(批次, 0) = Nvl(c_成本调整.批次, 0) And 性质 = 1;
      End If;
    
      --更新成本价调价信息 
      Update 成本价调价信息
      Set 收发id = n_收发id, 执行日期 = d_调价时间, 原成本价 = n_原成本价, 效期 = c_成本调整.效期, 灭菌效期 = c_成本调整.灭菌效期, 产地 = c_成本调整.产地,
          批号 = c_成本调整.批号
      Where ID = c_成本调整.调价id;
    End If;
  End Loop;

  --产生应付记录 
  For c_应付 In (Select Distinct a.供药单位id, a.药品id, a.发票号, a.发票日期, a.发票金额, b.名称, b.计算单位, b.规格
               From 成本价调价信息 A, 收费项目目录 B
               Where a.药品id = b.Id And Nvl(a.应付款变动, 0) = 1 And Nvl(a.供药单位id, 0) <> 0 And a.药品id = 材料id_In
               Order By a.供药单位id) Loop
  
    v_应付单据号 := Nextno(67);
  
    Select 应付记录_Id.Nextval Into v_应付id From Dual;
  
    Insert Into 应付记录
      (ID, 记录性质, 记录状态, 单位id, NO, 系统标识, 发票号, 发票日期, 发票金额, 品名, 规格, 填制人, 填制日期, 审核人, 审核日期, 摘要)
    Values
      (v_应付id, 1, 1, c_应付.供药单位id, v_应付单据号, 5, c_应付.发票号, c_应付.发票日期, c_应付.发票金额, c_应付.名称, c_应付.规格, Zl_Username, d_调价时间,
       Zl_Username, d_调价时间, '成本价调价自动产生应付款变动记录');
  
    If Nvl(c_应付.供药单位id, 0) <> 0 Then
      Update 应付余额 Set 金额 = Nvl(金额, 0) + Nvl(c_应付.发票金额, 0) Where 单位id = c_应付.供药单位id And 性质 = 1;
      If Sql%NotFound Then
        Insert Into 应付余额 (单位id, 性质, 金额) Values (c_应付.供药单位id, 1, Nvl(c_应付.发票金额, 0));
      End If;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料收发记录_成本价调价;
/

--127450:李南春,2018-06-27,挂号按先进先出原则使用预交款
Create Or Replace Procedure Zl_病人挂号记录_出诊_Insert
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
  n_安排id         挂号安排.Id%Type;
  n_预约顺序号     临床出诊序号控制.预约顺序号%Type;
  n_计划id         挂号安排计划.Id%Type := 0;
  v_星期           挂号安排限制.限制项目%Type;
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
        From 临床出诊记录
        Where ID = 出诊记录id_In And Nvl(是否发布, 0) = 1 And Nvl(是否锁定, 0) = 0;
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

--127720:殷瑞,2018-06-29,同步处理药品收发记录数据更新问题
Create Or Replace Procedure Zl_药品收发记录_更新批次
(
  Id_In     In 药品收发记录.Id%Type,
  药品id_In In 药品收发记录.药品id%Type,
  批次_In   药品收发记录.批次%Type := Null
) Is
  Str药品     Varchar2(500);
  Lng库房id   药品收发记录.库房id%Type;
  Lngcur批次  药品收发记录.批次%Type;
  Lnglast批次 药品收发记录.批次%Type;
  Str批号     药品收发记录.批号%Type;
  Str效期     药品收发记录.效期%Type;
  Lng供应商id 药品收发记录.供药单位id%Type;
  Dat生产日期 药品收发记录.生产日期%Type;
  Str产地     药品收发记录.产地%Type;
  Dbl可用数量 药品收发记录.填写数量%Type;
  Dbl实际数量 药品收发记录.实际数量%Type;
  Str批准文号 药品收发记录.批准文号%Type;
  v_Error     Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(批次_In, 0) = 0 Then
    Return;
  End If;

  Select Nvl(a.批次, 0), a.库房id, Nvl(a.实际数量, 0) * Nvl(a.付数, 1), '[' || c.编码 || ']' || c.名称
  Into Lnglast批次, Lng库房id, Dbl实际数量, Str药品
  From 药品收发记录 A, 收费项目目录 C
  Where a.Id = Id_In And a.药品id = c.Id;
  If Nvl(批次_In, 0) = 0 Then
    Lngcur批次 := Nvl(Lnglast批次, 0);
  Else
    Lngcur批次 := Nvl(批次_In, 0);
  End If;

  Begin
    v_Error := '有一笔分批核算的药品，指定的批次已失效,不能完成操作！';
    --取该批药品的批号 
    Select 上次批号, 效期, Nvl(可用数量, 0) 可用数量, 上次供应商id, 上次生产日期, 上次产地, 批准文号
    Into Str批号, Str效期, Dbl可用数量, Lng供应商id, Dat生产日期, Str产地, Str批准文号
    From 药品库存
    Where 库房id = Lng库房id And 药品id = 药品id_In And 性质 = 1 And Nvl(批次, 0) = Lngcur批次 And
          (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate));
  
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  End;

  If Lngcur批次 <> Nvl(Lnglast批次, 0) Then
    If Dbl可用数量 < Dbl实际数量 And Lngcur批次 <> 0 Then
      v_Error := Str药品 || '的可用数量不足，操作中止！';
      Raise Err_Custom;
    End If;
  End If;
  --更新药品收发记录的批次信息 
  Update 药品收发记录
  Set 批次 = Lngcur批次, 批号 = Str批号, 效期 = Str效期, 供药单位id = Lng供应商id, 生产日期 = Dat生产日期, 产地 = Str产地, 批准文号 = Str批准文号
  Where ID = Id_In;

  --更新原批次库存的可用数量 
  --更新发药批次库存的可用数量 
  If Lnglast批次 <> Lngcur批次 Then
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + Dbl实际数量
    Where 库房id + 0 = Lng库房id And 药品id = 药品id_In And 性质 = 1 And Nvl(批次, 0) = Lnglast批次;
  
    --异常数据处理
    Zl_药品库存_可用数量异常处理(Lng库房id, 药品id_In, Lnglast批次);
  
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) - Dbl实际数量
    Where 库房id + 0 = Lng库房id And 药品id = 药品id_In And 性质 = 1 And Nvl(批次, 0) = Lngcur批次;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_更新批次;
/

--127655:胡俊勇,2018-06-22,过程中Commit语句清除
CREATE OR REPLACE Procedure Zl_门诊穿刺台_Liquid
(
  科室id_In In 排队记录.科室id%Type,
  病人id_In In 排队记录.病人id%Type
) Is

  --功能：“接单”为病人的排队记录分配等待穿刺的一个穿刺台

  n_Sn  门诊穿刺台.序号%Type;
  v_Tmp 病人医嘱记录.姓名%Type;
  d_Tmp Date;

  Err_Item Exception;
  v_Err_Msg varchar2(200);
Begin

  -- 清除24小时过期的门诊穿刺台
  Update 门诊穿刺台
  Set 待穿病人id = Null
  Where ID In
        (Select a.Id
         From 门诊穿刺台 A, 排队记录 B
         Where a.科室id = b.科室id And a.待穿病人id = b.病人id And a.有效 = 1 And a.科室id = 科室id_In And b.日期 < Sysdate - 1);

  -- 清除24小时过期，并且状态不为“1-待配液；5-待穿刺”的门诊穿刺台
  Update 门诊穿刺台
  Set 待穿病人id = Null
  Where ID In (Select a.Id
               From 门诊穿刺台 A, 排队记录 B
               Where a.科室id = b.科室id And a.待穿病人id = b.病人id And a.有效 = 1 And a.科室id = 科室id_In And Not b.状态 In (1, 5) And
                     b.日期 < Sysdate - 1);
    

  -- 为病人“排队记录”分配穿刺台
  Begin
    -- 锁定排队记录（并发控制）
    Select 日期
    Into d_Tmp
    From 排队记录
    Where 科室id = 科室id_In And 病人id = 病人id_In And 状态 = 1
    For Update Nowait;

    -- 查找科室相同，其他穿刺台未分配给排队记录
    Begin
      Select 序号
      Into n_Sn
      From 门诊穿刺台
      Where 科室id = 科室id_In And
            Not 序号 In (Select 穿刺台 From 排队记录 Where 状态 = 1 And 科室id = 科室id_In And 穿刺台 > 0) And
            (待穿病人id Is Null Or 待穿病人id = 0) And 有效 = 1 And Rownum < 2;
    Exception
      When Others Then
        n_Sn := Null;
    End;

    If n_Sn Is Not Null Then
      -- 找到后更新
      Update 排队记录 Set 穿刺台 = n_Sn, 日期 = Sysdate Where 科室id = 科室id_In And 病人id = 病人id_In And 状态 = 1;
        
    Else
      -- 未找到，就平均分配一个穿刺台给病人的排队记录
      Begin
        Select 穿刺台
        Into n_Sn
        From (Select a.穿刺台, Count(1) 数量
               From 排队记录 A, 门诊穿刺台 B
               Where a.科室id = b.科室id And a.穿刺台 = b.序号 And a.科室id = 科室id_In And a.日期 Between Sysdate - 1 And Sysdate And
                     a.状态 = 1 And b.有效 = 1
               Group By 穿刺台
               Order By 数量, 穿刺台) A
        Where Rownum < 2;
      Exception
        When Others Then
          n_Sn := Null;
      End;

      If n_Sn Is Not Null Then
        -- 找到后更新
        Update 排队记录 Set 穿刺台 = n_Sn, 日期 = Sysdate Where 科室id = 科室id_In And 病人id = 病人id_In And 状态 = 1;
          
      Else
        -- 未找到，分配一个最小号的穿刺台
        Begin
          Select Min(序号) Into n_Sn From 门诊穿刺台 Where 科室id = 科室id_In And 有效 = 1;
        Exception
          When Others Then
            n_Sn := Null;
        End;
        If n_Sn Is Not Null Then
          Update 排队记录
          Set 穿刺台 = n_Sn, 日期 = Sysdate
          Where 科室id = 科室id_In And 病人id = 病人id_In And 状态 = 1;
            
        Else
            
          v_Err_Msg := '当前科室未设置穿刺台或有效的穿刺台！';
          Raise Err_Item;
        End If;
      End If;

    End If;

  Exception
    When Err_Item Then
      Raise Err_Item;
    When Others Then
      Begin
        Select 姓名 Into v_Tmp From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Tmp := '未知';
      End;
      v_Err_Msg := '[' || v_Tmp || ']不在待配液队列中！';
      Raise Err_Item;
  End;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊穿刺台_Liquid;
/

--127651:刘兴洪,2018-06-22,删除过程中含有Commit语句，避免事务不连续. 
Create Or Replace Procedure Zl1_Autocptall(强制记帐_In In Number := 0) As
  Modilast Number(1); --是否修正上期自动计费参数
  Period   Varchar2(6); --需要计算的最小期间
  Cursor Patitab Is
    Select Distinct 病人id, 主页id
    From 在院病人自动记帐
    Where Trunc(终止日期) >= (Select Min(开始日期) From 期间表 Where 期间 >= Period);
Begin
  If f_Is_Primary_Node = 0 Then
    Return;
  End If;
  Begin
    Select 期间 Into Period From 期间表 Where Trunc(Sysdate) - 1 Between Trunc(开始日期) And Trunc(终止日期);
  Exception
    When Others Then
      Return;
  End;
  Select zl_GetSysParameter(7) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  For Patifld In Patitab Loop
    If Patifld.病人id Is Not Null And Patifld.主页id Is Not Null Then
      Zl1_Autocptone(Patifld.病人id, Patifld.主页id, Period, 1, 强制记帐_In);
    End If;
  End Loop;
End Zl1_Autocptall;
/

--127651:刘兴洪,2018-06-22,删除过程中含有Commit语句，避免事务不连续. 
Create Or Replace Procedure Zl1_Autocptpati
(
  Patiid      In Number,
  Pageid      In Number,
  Recalcbdate In 病人变动记录.上次计算时间%Type := Null,
  强制记帐_In In Number := 0
) As
  Modilast Number(1); --是否修正上期自动计费参数
  Period   Varchar2(6); --需要计算的最小期间
Begin
  Begin
    Select 期间 Into Period From 期间表 Where Trunc(Sysdate) Between Trunc(开始日期) And Trunc(终止日期);
  Exception
    When Others Then
      Return;
  End;

  Select Zl_To_Number(zl_GetSysParameter(7)) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  If Recalcbdate Is Not Null Then
    Update 病人变动记录
    Set 上次计算时间 = Null
    Where 病人id = Patiid And 主页id = Pageid And 上次计算时间 >= Recalcbdate;
  End If;

  Zl1_Autocptone(Patiid, Pageid, Period, 0, 强制记帐_In);
End Zl1_Autocptpati;
/

--127651:刘兴洪,2018-06-22,删除过程中含有Commit语句，避免事务不连续. 
Create Or Replace Procedure Zl1_Autocptward
(
  Wardid      In Number,
  Recalcbdate In 病人变动记录.上次计算时间%Type := Null,
  强制记帐_In In Number := 0
) As
  Modilast Number(1); --是否修正上期自动计费参数
  Period   Varchar2(6); --需要计算的最小期间

  Cursor Patitab Is
    Select Distinct 病人id, 主页id
    From 在院病人自动记帐
    Where 病区id = Wardid And Trunc(终止日期) >= (Select Min(开始日期) From 期间表 Where 期间 >= Period);
Begin
  Begin
    Select 期间 Into Period From 期间表 Where Trunc(Sysdate) - 1 Between Trunc(开始日期) And Trunc(终止日期);
  Exception
    When Others Then
      Return;
  End;
  Select zl_GetSysParameter(7) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  If Recalcbdate Is Not Null Then
    Update 病人变动记录
    Set 上次计算时间 = Null
    Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 当前病区id = Wardid And 出院日期 Is Null) And
          上次计算时间 >= Recalcbdate;
  End If;

  For Patifld In Patitab Loop
    If Patifld.病人id Is Not Null And Patifld.主页id Is Not Null Then
      Zl1_Autocptone(Patifld.病人id, Patifld.主页id, Period, 1, 强制记帐_In);
    End If;
  End Loop;
End Zl1_Autocptward;
/

--127651:刘兴洪,2018-06-22,删除过程中含有Commit语句，避免事务不连续. 
 Create Or Replace Procedure Zl_Third_Swapstaut
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:读取三方支付的状态
  --入参:Xml_In:
  --<IN>
  --        <JYLB>交易类别</JYLB> //支付宝，微信等
  --        <JYLSH>交易流水号</JYLSH>
  --        <BRID>病人ID</BRID>
  --</IN>
  --出参:Xml_Out
  --  <OUT>
  --    <ZT>状态</ZT>  ――0-交易失败,1-交易成功;2-交易正在进行中;3-不存在该交易记录
  --    <JYSJ>交易时间</JYSJ>  ――当状态为1时才返回,否则为空  格式为'YYYY-MM-DD hh24:mi:ss'
  --    <JYID>业务交易ID</JYID>  ――当状态为1时才返回,否则为空   针对挂号和结帐为结帐ID;针对预交为预交ID;针对收费为结算序号
  --    <YWLX>业务类型</YWLX>   ―― null-历史数据;1-预交;2-结帐;3-收费;4-挂号
  --    <DJH>单据号</DJH>  ―― 多个用逗号分隔  针对挂号为挂号单据号,针对结帐为结帐单据号,针对预交为预交单据号,针对收费为收费单据号


  --  </OUT>
  --------------------------------------------------------------------------------------------------
  v_Temp    Varchar2(32767); --临时XML
  x_Templet Xmltype; --模板XML

  v_交易类别   三方交易记录.类别%Type;
  v_交易流水号 三方交易记录.流水号%Type;
  n_病人id     病人预交记录.病人id%Type;

  n_Count    Number(18);
  n_交易id   三方交易记录.业务结算id%Type;
  n_业务类型 三方交易记录.业务类型%Type;
  v_Nos      Varchar2(3000);
  d_交易时间 Date;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Nvl(Extractvalue(Value(A), 'IN/JYLB'), '-'), Nvl(Extractvalue(Value(A), 'IN/JYLSH'), '-'),
         Nvl(Extractvalue(Value(A), 'IN/BRID'), 0)
  Into v_交易类别, v_交易流水号, n_病人id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Begin
    Select 状态, 交易时间, 业务结算id, 业务类型
    Into n_Count, d_交易时间, n_交易id, n_业务类型
    From 三方交易记录
    Where 类别 = v_交易类别 And 流水号 = v_交易流水号
    For Update Nowait;
  Exception
    When Others Then
      n_Count := -1;
  End;

  If n_Count = -1 Then
    Select Count(1) Into n_Count From 三方交易记录 Where 类别 = v_交易类别 And 流水号 = v_交易流水号;
    If n_Count = 0 Then
      Select Count(1), Max(收款时间), Max(Decode(记录性质, 1, ID, 3, 结算序号, 结帐id)), Max(Mod(记录性质, 10))
      Into n_Count, d_交易时间, n_交易id, n_业务类型
      From 病人预交记录
      Where 记录性质 <> 11 And Nvl(校对标志, 0) <> 1 And 病人id = n_病人id And 交易流水号 = v_交易流水号 And
            卡类别id In (Select ID From 医疗卡类别 Where 名称 = v_交易类别);
      If n_Count = 0 Then
        n_Count := 3;
      Else
        --不存在三方交易记录，但存在病人预交记录，也表示交易成功
        n_Count := 1;
      End If;
    Else
      n_Count := 2;
    End If;
  End If;

  If n_Count = 1 Then
    If n_业务类型 = 1 Then
      Select Max(NO) Into v_Nos From 病人预交记录 Where 记录性质 = 1 And ID = n_交易id;
    Elsif n_业务类型 = 2 Then
      Select Max(NO) Into v_Nos From 病人结帐记录 Where ID = n_交易id;
    Elsif n_业务类型 = 3 Then
      Select f_List2str(Cast(Collect(NO) As t_Strlist))
      Into v_Nos
      From (Select Distinct a.No As NO
             From 门诊费用记录 A, 病人预交记录 B
             Where a.结帐id = b.结帐id And b.结算序号 = n_交易id);
    Elsif n_业务类型 = 4 Then
      Select Max(NO) Into v_Nos From 门诊费用记录 Where 记录性质 = 4 And 结帐id = n_交易id;
    End If;
  End If;

  v_Temp := '<ZT>' || n_Count || '</ZT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JYSJ>' || To_Char(d_交易时间, 'YYYY-MM-DD hh24:mi:ss') || '</JYSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JYID>' || Nvl(n_交易id, 0) || '</JYID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YWLX>' || n_业务类型 || '</YWLX>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<DJH>' || Nvl(v_Nos, '') || '</DJH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Swapstaut;
/

--127651:刘兴洪,2018-06-22,删除过程中含有Commit语句，避免事务不连续. 
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
      Zl_门诊划价记录_Delete(r_Price.No, r_Price.序号, 1);
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊划价记录_Clear;
/

--127620:涂建华,2018-06-22,commit语句处理
CREATE OR REPLACE Procedure Zl_影像检查图象_校对
( 
  医嘱id_In   In 影像检查记录.医嘱id%Type, 
  图像uid_In  In 影像检查图象.图像uid%Type, 
  校对日期_In In 影像检查记录.校对日期%Type, 
  校对结果_In In 影像检查图象.校对结果%Type 
) Is 
  n_校对状态 Number(1); 
  n_Tag      Number(1); 
  n_Count    Number(4); 
  v_检查uid  影像检查记录.检查uid%Type; 
Begin 
 
  Select Nvl(校对状态, 0), 检查uid Into n_校对状态, v_检查uid From 影像检查记录 Where 医嘱id = 医嘱id_In; 
  Update 影像检查图象 Set 校对结果 = 校对结果_In Where 图像uid = 图像uid_In; 
 
  If 校对结果_In = 5 Or 校对结果_In = 6 Then 
    n_Tag := 1; 
  Else 
    n_Tag := 2; 
  End If; 
 
  If n_校对状态 = 0 Then 
    Update 影像检查记录 Set 校对状态 = n_Tag, 校对日期 = 校对日期_In Where 医嘱id = 医嘱id_In; 
  Elsif n_校对状态 = 1 And n_Tag = 2 Then 
    Update 影像检查记录 Set 校对状态 = n_Tag Where 医嘱id = 医嘱id_In; 
  Elsif n_校对状态 = 2 And (校对结果_In = 5 Or 校对结果_In = 6) Then 
    Select Count(1) 
    Into n_Count 
    From 影像检查序列 b, 影像检查图象 c 
    Where b.序列uid = c.序列uid And b.检查uid = v_检查uid And (c.校对结果 > 0 And c.校对结果 < 5); 
 
    If n_Count = 0 Then 
      Update 影像检查记录 Set 校对状态 = 1 Where 医嘱id = 医嘱id_In; 
    End If; 
  End If; 
Exception 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End Zl_影像检查图象_校对;
/

--126802:李南春,2018-06-21,附加费返回固定的金额信息
Create Or Replace Procedure Zl_Third_Getregfeedetail
(
  Xml_In  In Xmltype,
  Xml_Out In Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取挂号费用明细
  --入参:Xml_In:
  --<IN>
  --  <BRID></BRID> //病人ID
  --  <XM></XM>     //姓名
  --  <SFZH></SFZH> //身份证号
  --  <SFYY></SFYY> //是否仅预约不支付,1-仅预约不支付,0-挂号,预约支付,预约接收，默认为0
  --  <GHDH></GHDH> //挂号单号,预约接收时传入
  --  <GHHM></GHHM> //挂号安排号码,挂号和预约时传入
  --  <XMID></XMID> //挂号安排的项目ID,挂号和预约时传入
  --  <FB></FB>     //病人费别
  --  <FKFS></FKFS> //付款方式
  --  <RQ></RQ>     //日期
  --  <ZD></ZD>     //站点
  --</IN>
  --出参:Xml_Out
  -- <OUTPUT>
  --  <ZJE></ZJE>   //总实收金额
  --  <XMMX>        //项目明细
  --    <XM>
  --      <DJH></DJH>       //单据号
  --      <MC></MC>   //项目名称
  --      <ID></ID>   //项目ID
  --      <SL></SL>   //数量，数次*付数
  --      <YSJE></YSJE>   //应收金额
  --      <SSJE></SSJE>   //实收金额
  --      <SJFM></SJFM>       //收据费目
  --    </XM>
  --    <XM>
  --    ...
  --    </XM>
  --  </XMMX>
  -- </OUTPUT>

  --------------------------------------------------------------------------------------------------
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  v_Temp       Varchar2(4000);
  n_项目id     挂号安排.项目id%Type;
  v_No         门诊费用记录.No%Type;
  n_预约       Number(3);
  n_病人id     病人信息.病人id%Type;
  v_姓名       病人信息.姓名%Type;
  v_性别       病人信息.性别%Type;
  v_年龄       病人信息.年龄%Type;
  v_身份证号   病人信息.身份证号%Type;
  v_费别       病人信息.费别%Type;
  v_付款方式   医疗付款方式.名称%Type;
  v_方式       Varchar2(20);
  d_日期       Date;
  v_站点       部门表.站点%Type;
  v_号码       挂号安排.号码%Type;
  n_总金额     门诊费用记录.实收金额%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  v_实收       Varchar(500);
  v_附加项目id Varchar2(500);
  v_附加内容   Varchar2(500);
  v_附加值     Varchar2(100);
  n_cursor     Number(3);
  Err_Item     Exception;
  
  TYPE Price_type IS RECORD(项目ID 门诊费用记录.收费细目ID%Type,
                              数次 门诊费用记录.数次%TYPE, 
                              单价 门诊费用记录.标准单价%TYPE, 
                              应收 门诊费用记录.应收金额%TYPE, 
                              实收 门诊费用记录.实收金额%TYPE);--定义Price记录类型 
  TYPE Price_type_array IS TABLE OF Price_type INDEX BY BINARY_INTEGER;--定义存放Price记录的数组类型 
  Price_rec Price_type;--声明变量，类型：Price记录类型
  Price_rec_array Price_type_array;--声明变量，类型：存放Price记录的数组类型
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/XMID'), Extractvalue(Value(A), 'IN/GHHM'),
         Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/ZD'), Extractvalue(Value(A), 'IN/SFYY'),
         Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM'),
         To_Date(Extractvalue(Value(A), 'IN/RQ'), 'yyyy-mm-dd hh24:mi:ss'), Extractvalue(Value(A), 'IN/FKFS')
  Into n_病人id, n_项目id, v_号码, v_费别, v_站点, n_预约, v_No, v_身份证号, v_姓名, d_日期, v_方式
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_方式 Is Null Then
    Select Max(名称) Into v_付款方式 From 医疗付款方式 Where 缺省标志 = 1;
  Else
    Select Max(名称) Into v_付款方式 From 医疗付款方式 Where 编码 = v_方式;
    If v_付款方式 Is Null Then
      v_付款方式 := v_方式;
    End If;
  End If;
  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查';
    Raise Err_Item;
  End If;
  Select Max(性别), Max(年龄) Into v_性别, v_年龄 From 病人信息 Where 病人ID = n_病人id;
  
  n_总金额 := 0;
  If v_No Is Null Then
    --挂号或者预约
    For c_挂号项目 In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_项目id And d_日期 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                   From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And d.主项id = n_项目id And
                         d_日期 Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
      v_实收     := Zl_Actualmoney(v_费别, c_挂号项目.项目id, c_挂号项目.收入项目id, c_挂号项目.数次 * c_挂号项目.单价);
      n_实收金额 := To_Number(Substr(v_实收, Instr(v_实收, ':') + 1));
      n_总金额   := n_总金额 + Nvl(n_实收金额, 0);
      v_Temp     := v_Temp || '<XM><DJH></DJH><MC>' || c_挂号项目.项目名称 || '</MC>' || '<ID>' || c_挂号项目.项目id || '</ID>' ||
                    '<SL>' || c_挂号项目.数次 || '</SL>' || '<YSJE>' || c_挂号项目.数次 * c_挂号项目.单价 || '</YSJE>' || '<SSJE>' ||
                    n_实收金额 || '</SSJE>' || '<SJFM>' || c_挂号项目.收据费目 || '</SJFM></XM>';
    End Loop;
  Else
    --预约接收
    For c_挂号项目 In (Select a.收费细目id As 项目id, a.应收金额, a.实收金额, a.计算单位, a.收据费目, b.名称 As 项目名称, a.No, Nvl(a.付数, 1) As 付数, a.数次
                   From 门诊费用记录 A, 收费项目目录 B
                   Where a.收费细目id = b.Id And a.No = v_No And a.记录性质 = 4 And a.记录状态 = 0) Loop
      n_总金额 := n_总金额 + Nvl(c_挂号项目.实收金额, 0);
      v_号码   := c_挂号项目.计算单位;
      v_Temp   := v_Temp || '<XM><DJH>' || c_挂号项目.No || '</DJH><MC>' || c_挂号项目.项目名称 || '</MC>' || '<ID>' || c_挂号项目.项目id ||
                  '</ID>' || '<SL>' || c_挂号项目.付数 * c_挂号项目.数次 || '</SL>' || '<YSJE>' || c_挂号项目.应收金额 || '</YSJE>' ||
                  '<SSJE>' || c_挂号项目.实收金额 || '</SSJE>' || '<SJFM>' || c_挂号项目.收据费目 || '</SJFM></XM>';
    End Loop;
  End If;

  If Nvl(n_预约, 0) = 0 Then
    Begin
      Select Zl_Fun_Customregexpenses(n_病人id, 0, v_号码, v_姓名, v_性别, v_年龄, v_身份证号, v_费别, v_付款方式) Into v_附加项目id From Dual;
    Exception
      When Others Then
        v_附加项目id := Null;
    End;
    If v_附加项目id Is Not Null Then
      IF Instr(v_附加项目id, '|') > 0 Then
        v_附加内容 := v_附加项目id || ','; --以空格分开以|结尾,没有结算号码的
        v_附加项目id := '';
        n_cursor   := 0;
        While v_附加内容 Is Not Null Loop
          v_附加值 := Substr(v_附加内容, 1, Instr(v_附加内容, ',') - 1);
          Price_rec.项目ID := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
        
          v_附加值 := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_rec.数次 := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
        
          v_附加值 := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_rec.单价 := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
          
          v_附加值 := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_rec.应收 := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
        
          v_附加值 := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_rec.实收 := To_Number(v_附加值);
          n_cursor := n_cursor + 1;
          Price_rec_array(n_cursor):=Price_rec;
          v_附加内容 := Substr(v_附加内容, Instr(v_附加内容, ',') + 1);
          v_附加项目id := v_附加项目id || ',' || Price_rec_array(n_cursor).项目ID;
        End Loop;
        
        If v_附加项目id is not null then
          v_附加项目id := substr(v_附加项目id, 2);
        End if;
        
        For c_附加项目 In (Select /*+cardinality(D,10)*/
                        5 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次, c.Id As 收入项目id,
                        c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                       From 收费项目目录 A, 收费价目 B, 收入项目 C, Table(f_Str2list(v_附加项目id)) D
                       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.Column_Value And d_日期 Between b.执行日期 And
                             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')))  Loop
          FOR n_cursor IN 1..Price_rec_array.count LOOP        
            IF c_附加项目.项目id = Price_rec_array(n_cursor).项目ID Then
              n_实收金额 := Price_rec_array(n_cursor).实收;
              n_总金额   := n_总金额 + Nvl(n_实收金额, 0);
              v_Temp     := v_Temp || '<XM><DJH></DJH><MC>' || c_附加项目.项目名称 || '</MC>' || '<ID>' || c_附加项目.项目id || '</ID>' ||
                          '<SL>' || Price_rec_array(n_cursor).数次 || '</SL>' || '<YSJE>' || Price_rec_array(n_cursor).应收 || '</YSJE>' || '<SSJE>' ||
                          Price_rec_array(n_cursor).实收 || '</SSJE>' || '<SJFM>' || c_附加项目.收据费目 || '</SJFM></XM>';
              EXIT;
            End IF;
          End LOOP;
        End Loop;  
      Else
        For c_附加项目 In (Select /*+cardinality(D,10)*/
                        5 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次, c.Id As 收入项目id,
                        c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                       From 收费项目目录 A, 收费价目 B, 收入项目 C, Table(f_Str2list(v_附加项目id)) D
                       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.Column_Value And d_日期 Between b.执行日期 And
                             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                       Union All
                       Select /*+cardinality(E,10)*/
                        5 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                        c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                       From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D, Table(f_Str2list(v_附加项目id)) E
                       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And d.主项id = e.Column_Value And
                             d_日期 Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
          v_实收     := Zl_Actualmoney(v_费别, c_附加项目.项目id, c_附加项目.收入项目id, c_附加项目.数次 * c_附加项目.单价);
          n_实收金额 := To_Number(Substr(v_实收, Instr(v_实收, ':') + 1));
          n_总金额   := n_总金额 + Nvl(n_实收金额, 0);
          v_Temp     := v_Temp || '<XM><DJH></DJH><MC>' || c_附加项目.项目名称 || '</MC>' || '<ID>' || c_附加项目.项目id || '</ID>' ||
                        '<SL>' || c_附加项目.数次 || '</SL>' || '<YSJE>' || c_附加项目.数次 * c_附加项目.单价 || '</YSJE>' || '<SSJE>' ||
                        n_实收金额 || '</SSJE>' || '<SJFM>' || c_附加项目.收据费目 || '</SJFM></XM>';
        End Loop;
      End IF;
    End If;
  End If;

  v_Temp := '<XMMX>' || v_Temp || '</XMMX>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<ZJE>' || n_总金额 || '</ZJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getregfeedetail;
/

--126802:李南春,2018-06-21,附加费返回固定的金额信息
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
  --        <BRID>病人ID</BRID>             //病人ID
  --        <XM>姓名</XM>                   //姓名
  --        <SFZH>身份证号</SFZH>           //身份证号
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
  --    <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --    <JZID>结帐ID</JZID>          //本次结帐ID
  --    ――如无下列错误结点则说明正确执行 
  --    <ERROR> 
  --      <MSG>错误信息</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  v_Nos      Varchar2(4000);
  n_收费总额 门诊费用记录.实收金额%Type;

  n_卡类别id 医疗卡类别.Id%Type;
  v_结算方式 Varchar2(2000);
  n_病人id   门诊费用记录.病人id%Type;
  v_身份证号 病人信息.身份证号%Type;
  v_姓名     门诊费用记录.姓名%Type;
  v_性别     门诊费用记录.性别%Type;
  v_年龄     门诊费用记录.年龄%Type;

  v_医疗付款方式编码 医疗付款方式.编码%Type;
  v_付款方式         医疗付款方式.名称%Type;
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
  v_Para             Varchar2(500);
  n_挂号模式         Number(3);
  d_启用时间         Date;
  v_临时结算方式     病人预交记录.结算方式%Type;
  n_出诊记录id       临床出诊记录.Id%Type;
  n_序号             门诊费用记录.序号%Type;
  v_附加项目id       Varchar2(500);
  v_附加内容         Varchar2(500);
  v_附加值           Varchar2(100);
  n_cursor           Number(3);
  n_实收金额         门诊费用记录.实收金额%Type;
  v_实收             Varchar2(500);
  n_从属父号         门诊费用记录.从属父号%Type;
  n_病人科室id       门诊费用记录.病人科室id%Type;
  n_执行部门id       门诊费用记录.执行部门id%Type;
  v_No               门诊费用记录.No%Type;
  n_医保支付         病人预交记录.冲预交%Type;
  n_Exists           Number;
  v_卡类别           三方交易记录.类别%Type;
  n_业务类型         三方交易记录.业务类型%Type;
  n_结算序号         病人预交记录.结算序号%Type;
  v_Temp             Varchar2(32767); --临时XML 
  x_Templet          Xmltype; --模板XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  Err_Special Exception;
  n_Count    Number(18);
  v_操作员   门诊费用记录.操作员姓名%Type;
  v_发药窗口 Varchar2(4000);
  n_误差额   病人预交记录.冲预交%Type;
  
  TYPE Price_type IS RECORD(项目ID 门诊费用记录.收费细目ID%Type,
                              数次 门诊费用记录.数次%TYPE, 
                              单价 门诊费用记录.标准单价%TYPE, 
                              应收 门诊费用记录.应收金额%TYPE, 
                              实收 门诊费用记录.实收金额%TYPE);--定义Price记录类型 
  TYPE Price_type_array IS TABLE OF Price_type INDEX BY BINARY_INTEGER;--定义存放Price记录的数组类型 
  Price_rec Price_type;--声明变量，类型：Price记录类型
  Price_rec_array Price_type_array;--声明变量，类型：存放Price记录的数组类型

  Function Zl_出诊诊室(记录id_In 临床出诊记录.Id%Type) Return Varchar2 As
    n_分诊方式 临床出诊记录.分诊方式%Type;
    v_诊室     病人挂号记录.诊室%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
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
                          Where Nvl(执行状态, 0) = 0 And 发生时间 Between Trunc(Sysdate) And Sysdate And 出诊记录id = 记录id_In And
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
         To_Number(Extractvalue(Value(A), 'IN/SFGH')), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM')
  Into v_Nos, n_病人id, n_收费总额, n_误差额, n_是否挂号, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
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
    Select b.编码, b.名称, a.姓名, a.性别, a.年龄
    Into v_医疗付款方式编码, v_付款方式, v_姓名, v_性别, v_年龄
    From 病人信息 A, 医疗付款方式 B
    Where a.医疗付款方式 = b.名称(+) And a.病人id = n_病人id;
  Exception
    When Others Then
      v_Err_Msg := '指定的缴费单据中不能有效识别病人,不允许缴费!';
  End;
  If Not v_Err_Msg Is Null Then
    Raise Err_Item;
  End If;

  Select Decode(Nvl(n_是否挂号, 0), 0, 3, 4) Into n_业务类型 From Dual;

  For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If c_交易记录.结算卡类别 Is Null Then
      v_卡类别 := c_交易记录.结算方式;
    Else
      Select Decode(Translate(Nvl(c_交易记录.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
      Into n_Count
      From Dual;
    
      If Nvl(n_Count, 0) = 1 Then
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_交易记录.结算卡类别);
      Else
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_交易记录.结算卡类别;
      End If;
    End If;
    If v_卡类别 Is Null Then
      v_Err_Msg := '不支持的结算方式,请检查！';
      Raise Err_Item;
    End If;
  
    If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, n_业务类型) = 0 Then
      v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易!';
      Raise Err_Special;
    End If;
  End Loop;

  Select 病人结帐记录_Id.Nextval, Sysdate Into n_结帐id, d_收费时间 From Dual;
  n_结算序号 := -1 * n_结帐id;

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
      n_误差额 := Nvl(n_结帐金额, 0) - Nvl(n_收费总额, 0);
      If Abs(n_误差额) > 1.00 Then
        v_Err_Msg := '单据缴款金额与实际结算的误差太大!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_结帐金额, 0) <> Nvl(n_收费总额, 0) + Nvl(n_误差额, 0) Then
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
    
      If c_结算方式.结算卡类别 Is Null Then
        v_卡类别 := c_结算方式.结算方式;
      Else
        Select Decode(Translate(Nvl(c_结算方式.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Count
        From Dual;
      
        If Nvl(n_Count, 0) = 1 Then
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_结算方式.结算卡类别);
        Else
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_结算方式.结算卡类别;
        End If;
      End If;
    
      Update 三方交易记录
      Set 业务结算id = n_结算序号
      Where 流水号 = c_结算方式.交易流水号 And 类别 = v_卡类别 And 业务类型 = n_业务类型;
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
    v_Para     := zl_GetSysParameter(256);
    n_挂号模式 := Substr(v_Para, 1, 1);
    Begin
      d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
    For c_费用 In (Select 1 As 顺序号, b.No, b.收据费目, b.结帐id, b.执行部门id, b.病人科室id, b.开单人, b.收费类别, b.收入项目id, b.附加标志,
                        To_Char(b.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.价格父号, b.从属父号, b.序号, b.收费细目id, b.计算单位,
                        Max(m.名称) As 名称, Max(m.规格) As 规格, Sum(b.标准单价) As 单价, Avg(Nvl(b.付数, 1) * b.数次) As 数量,
                        Sum(b.应收金额) As 应收金额, Sum(b.实收金额) As 实收金额, Max(j.名称) As 开单科室, Max(q.名称) As 执行科室
                 From 门诊费用记录 B, 收费项目目录 M, 部门表 J, 部门表 Q
                 Where b.No = v_Nos And b.记录性质 = 4 And Nvl(b.费用状态, 0) = 0 And b.记录状态 = 0 And b.收费细目id = m.Id And
                       b.开单部门id = j.Id(+) And b.执行部门id = q.Id(+)
                 Group By b.No, b.收据费目, b.结帐id, b.执行部门id, b.病人科室id, b.开单人, b.收入项目id, b.收费类别, b.登记时间, b.价格父号, b.从属父号, b.序号,
                          b.收费细目id, b.计算单位, b.附加标志
                 Order By 序号) Loop
      Zl_病人预约挂号记录_Update(c_费用.No, c_费用.序号, c_费用.价格父号, c_费用.从属父号, c_费用.收费类别, c_费用.收费细目id, c_费用.数量, c_费用.单价, c_费用.收入项目id,
                         c_费用.收据费目, c_费用.应收金额, c_费用.实收金额, c_费用.附加标志, Null, Null, Null, Null, c_费用.病人科室id, c_费用.执行部门id);
      n_结帐金额   := n_结帐金额 + c_费用.实收金额;
      n_序号       := c_费用.序号;
      n_病人科室id := c_费用.病人科室id;
      n_执行部门id := c_费用.执行部门id;
      v_No         := c_费用.No;
    End Loop;
  
    Select a.执行部门id, a.收费细目id, c.Id, a.执行人, b.号别, b.门诊号, b.发生时间, a.费别, b.号序, b.出诊记录id
    Into n_科室id, n_项目id, n_医生id, v_医生姓名, v_号码, n_门诊号, d_发生时间, v_费别, n_号序, n_出诊记录id
    From 门诊费用记录 A, 病人挂号记录 B, 人员表 C
    Where a.No = v_Nos And a.记录性质 = 4 And a.序号 = 1 And a.No = b.No And a.执行人 = c.姓名(+);
  
    Begin
      Select Zl_Fun_Customregexpenses(n_病人id, 0, v_号码, v_姓名, v_性别, v_年龄, v_身份证号, v_费别, v_付款方式) Into v_附加项目id From Dual;
    Exception
      When Others Then
        v_附加项目id := Null;
    End;
    If v_附加项目id Is Not Null Then
      IF Instr(v_附加项目id, '|') > 0 Then
        v_附加内容 := v_附加项目id || ','; --以空格分开以|结尾,没有结算号码的
        v_附加项目id := '';
        n_cursor   := 0;
        While v_附加内容 Is Not Null Loop
          v_附加值 := Substr(v_附加内容, 1, Instr(v_附加内容, ',') - 1);
          Price_rec.项目ID := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
        
          v_附加值 := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_rec.数次 := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
        
          v_附加值 := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_rec.单价 := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
          
          v_附加值 := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_rec.应收 := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
        
          v_附加值 := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_rec.实收 := To_Number(v_附加值);
          n_cursor := n_cursor + 1;
          Price_rec_array(n_cursor):=Price_rec;
          v_附加内容 := Substr(v_附加内容, Instr(v_附加内容, ',') + 1);
          v_附加项目id := v_附加项目id || ',' || Price_rec_array(n_cursor).项目ID;
        End Loop;
        
        If v_附加项目id is not null then
          v_附加项目id := substr(v_附加项目id, 2);
        End if;
        
        For c_附加项目 In (Select /*+cardinality(D,10)*/
                        5 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次, c.Id As 收入项目id,
                        c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                       From 收费项目目录 A, 收费价目 B, 收入项目 C, Table(f_Str2list(v_附加项目id)) D
                       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.Column_Value And Sysdate Between b.执行日期 And
                             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
        
          n_序号 := n_序号 + 1;
          Zl_病人预约挂号记录_Update(v_No, n_序号, Null, Null, c_附加项目.类别, c_附加项目.项目id, Price_rec_array(n_cursor).数次, Price_rec_array(n_cursor).单价, c_附加项目.收入项目id,
                               c_附加项目.收据费目, Price_rec_array(n_cursor).应收, Price_rec_array(n_cursor).实收, Null, Null, Null, Null, Null, n_病人科室id,
                               n_执行部门id);

          n_实收金额 := Price_rec_array(n_cursor).实收;                  
          n_结帐金额 := n_结帐金额 + n_实收金额;
        End Loop;                        
      Else
        For c_附加项目 In (Select /*+cardinality(D,10)*/
                        5 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次, c.Id As 收入项目id,
                        c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                       From 收费项目目录 A, 收费价目 B, 收入项目 C, Table(f_Str2list(v_附加项目id)) D
                       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.Column_Value And Sysdate Between b.执行日期 And
                             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                       Union All
                       Select /*+cardinality(E,10)*/
                        6 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                        c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                       From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D, Table(f_Str2list(v_附加项目id)) E
                       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And d.主项id = e.Column_Value And
                             Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
          n_序号 := n_序号 + 1;
          If c_附加项目.性质 = 5 Then
            n_从属父号 := n_序号;
          End If;
        
          v_实收     := Zl_Actualmoney(v_费别, c_附加项目.项目id, c_附加项目.收入项目id, c_附加项目.数次 * c_附加项目.单价);
          n_实收金额 := To_Number(Substr(v_实收, Instr(v_实收, ':') + 1));
        
          If c_附加项目.性质 = 5 Then
            Zl_病人预约挂号记录_Update(v_No, n_序号, Null, Null, c_附加项目.类别, c_附加项目.项目id, c_附加项目.数次, c_附加项目.单价, c_附加项目.收入项目id,
                               c_附加项目.收据费目, c_附加项目.数次 * c_附加项目.单价, n_实收金额, Null, Null, Null, Null, Null, n_病人科室id,
                               n_执行部门id);
          Else
            Zl_病人预约挂号记录_Update(v_No, n_序号, Null, n_从属父号, c_附加项目.类别, c_附加项目.项目id, c_附加项目.数次, c_附加项目.单价, c_附加项目.收入项目id,
                               c_附加项目.收据费目, c_附加项目.数次 * c_附加项目.单价, n_实收金额, Null, Null, Null, Null, Null, n_病人科室id,
                               n_执行部门id);
          End If;
          n_结帐金额 := n_结帐金额 + n_实收金额;
        
        End Loop;
      End IF;
    End If;
  
    --检查总金额是否正确 
    If Nvl(n_误差额, 0) = 0 Then
      n_误差额 := Nvl(n_结帐金额, 0) - Nvl(n_收费总额, 0);
      If Abs(n_误差额) > 1.00 Then
        v_Err_Msg := '单据缴款金额与实际结算的误差太大!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_结帐金额, 0) <> Nvl(n_收费总额, 0) + Nvl(n_误差额, 0) Then
      Select Max(操作员姓名) Into v_操作员 From 门诊费用记录 Where 记录性质 = 4 And NO = v_Nos;
      If v_操作员 = v_操作员姓名 Then
        v_Err_Msg := '指定的缴费单据的总金额不对,请重新选择缴费单据!';
        Raise Err_Special;
      Else
        v_Err_Msg := '指定的缴费单据的总金额不对,请重新选择缴费单据!';
        Raise Err_Item;
      End If;
    End If;
  
    --预约接收
    If n_挂号模式 = 1 Then
      If d_启用时间 > d_发生时间 And n_出诊记录id Is Null Then
        n_挂号模式 := 0;
      End If;
    End If;
  
    Select Decode(To_Number(zl_GetSysParameter('排队叫号模式', 1113, 100)), 0, 0, 1) Into n_生成队列 From Dual;
    If n_挂号模式 = 0 Then
      For c_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                            Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                            Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                            Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                            Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                            Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明,
                            Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                            Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                            Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
        If Nvl(c_结算方式.是否冲预交, 0) = 1 Then
          n_预交支付 := c_结算方式.结算金额;
        Else
          If c_结算方式.结算方式 Is Not Null Then
            Select Nvl(Max(1), 0) Into n_Exists From 结算方式 Where 名称 = c_结算方式.结算方式 And 性质 In (3, 4);
            If n_Exists = 1 Then
              n_医保支付 := c_结算方式.结算金额;
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
                      Select ID
                      Into n_卡类别id
                      From 医疗卡类别
                      Where 名称 = c_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
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
                    Select ID
                    Into n_卡类别id
                    From 医疗卡类别
                    Where 名称 = c_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
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
        End If;
        If c_结算方式.结算卡类别 Is Null Then
          v_卡类别 := c_结算方式.结算方式;
        Else
          Select Decode(Translate(Nvl(c_结算方式.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_结算方式.结算卡类别);
          Else
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_结算方式.结算卡类别;
          End If;
        End If;
        Update 三方交易记录
        Set 业务结算id = n_结帐id
        Where 流水号 = c_结算方式.交易流水号 And 类别 = v_卡类别 And 业务类型 = n_业务类型;
      End Loop;
      Zl_预约挂号接收_Insert(v_Nos, Null, Null, n_结帐id, Zl_诊室(v_号码), n_病人id, n_门诊号, v_姓名, v_性别, v_年龄, v_医疗付款方式编码, v_费别,
                       v_结算方式, n_普通支付, n_预交支付, n_医保支付, d_发生时间, n_号序, v_操作员编码, v_操作员姓名, n_生成队列, d_收费时间, n_卡类别id, n_结算卡序号,
                       v_结算卡号, v_交易流水号, v_交易说明, Null, 0, 0, Null, 1);
    Else
      For c_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                            Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                            Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                            Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                            Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                            Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明,
                            Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                            Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                            Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
        If Nvl(c_结算方式.是否冲预交, 0) = 1 Then
          n_预交支付 := c_结算方式.结算金额;
        Else
          n_普通支付 := Nvl(n_普通支付, 0) + c_结算方式.结算金额;
          If c_结算方式.结算方式 Is Null Then
            --三方卡结算方式
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
              Select 结算方式 Into v_临时结算方式 From 卡消费接口目录 Where 编号 = n_结算卡序号;
            Else
              Begin
                n_卡类别id := To_Number(c_结算方式.结算卡类别);
              Exception
                When Others Then
                  n_卡类别id := 0;
              End;
              If n_卡类别id = 0 Then
                Begin
                  Select ID
                  Into n_卡类别id
                  From 医疗卡类别
                  Where 名称 = c_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
                Exception
                  When Others Then
                    v_Err_Msg := '未找到对应的医疗卡!';
                    Raise Err_Item;
                End;
              End If;
              Select 结算方式 Into v_临时结算方式 From 医疗卡类别 Where ID = n_卡类别id;
            End If;
            v_结算卡号   := c_结算方式.结算卡号;
            v_交易流水号 := c_结算方式.交易流水号;
            v_交易说明   := c_结算方式.交易说明;
            v_摘要       := c_结算方式.摘要;
            v_结算方式   := v_结算方式 || '|' || v_临时结算方式 || ',' || c_结算方式.结算金额 || ',,1';
          Else
            --其他结算方式
            v_结算方式 := v_结算方式 || '|' || c_结算方式.结算方式 || ',' || c_结算方式.结算金额 || ',,1';
          End If;
        End If;
        If c_结算方式.结算卡类别 Is Null Then
          v_卡类别 := c_结算方式.结算方式;
        Else
          Select Decode(Translate(Nvl(c_结算方式.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_结算方式.结算卡类别);
          Else
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_结算方式.结算卡类别;
          End If;
        End If;
        Update 三方交易记录
        Set 业务结算id = n_结帐id
        Where 流水号 = c_结算方式.交易流水号 And 类别 = v_卡类别 And 业务类型 = n_业务类型;
      End Loop;
      If v_结算方式 Is Not Null Then
        v_结算方式 := Substr(v_结算方式, 2);
      End If;
      Zl_预约挂号接收_出诊_Insert(v_Nos, Null, Null, n_结帐id, Zl_出诊诊室(n_出诊记录id), n_病人id, n_门诊号, v_姓名, v_性别, v_年龄, v_医疗付款方式编码,
                          v_费别, v_结算方式, n_普通支付, n_预交支付, Null, d_发生时间, n_号序, v_操作员编码, v_操作员姓名, n_生成队列, d_收费时间, n_卡类别id,
                          n_结算卡序号, v_结算卡号, v_交易流水号, v_交易说明, Null, 0, 0, Null, 1);
    End If;
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
    Zl_病人挂号汇总_Update(v_医生姓名, n_医生id, n_项目id, n_科室id, d_发生时间, 2, v_号码, 1, n_出诊记录id);
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_收费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JZID>' || n_结帐id || '</JZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Payment;
/

--126802:李南春,2018-06-21,附加费返回固定的金额信息
Create Or Replace Function Zl_Fun_Customregexpenses
(
  病人id_In       In 病人信息.病人id%Type,
  险类_In         In 病人信息.险类%Type,
  号码_In         In 挂号安排.号码%Type,
  姓名_In         In 病人信息.姓名%Type := Null,
  性别_In         In 病人信息.性别%Type := Null,
  年龄_In         In 病人信息.年龄%Type := Null,
  身份证号_In     In 病人信息.身份证号%Type := Null,
  费别_In         In 病人信息.费别%Type := Null,
  医疗付款方式_In In 病人信息.医疗付款方式%Type := Null
) Return Varchar2
--    功能：挂号附加费处理项目用户自定义函数
  --    参数：
  --        病人ID_In：病人信息.病人ID
  --        险类_In：病人信息.险类
  --        号码_In: 挂号安排.号码
  --    返回: 格式一：收费细目ID1|数次1|单价1|应收1|实收1,收费细目ID2|数次2....多个收费细目用逗号分隔,项目的应收、实收等信息都以返回的值为准。
  --          格式二：收费细目ID1,收费细目ID2...只返回收费细目ID时以收费价目为准。
  --    返回NULL时，不处理,不能返回相同的收费细目ID
 Is
Begin
  Return Null;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Fun_Customregexpenses;
/

--127450:李南春,2018-06-20,余额退款时增加退款记录的冲预交信息，避免被退预交再次使用
Create Or Replace Procedure Zl_病人预交记录_Insert
(
  Id_In           病人预交记录.Id%Type,
  单据号_In       病人预交记录.No%Type,
  票据号_In       票据使用明细.号码%Type,
  病人id_In       病人预交记录.病人id%Type,
  主页id_In       病人预交记录.主页id%Type,
  科室id_In       病人预交记录.科室id%Type,
  金额_In         病人预交记录.金额%Type,
  结算方式_In     病人预交记录.结算方式%Type,
  结算号码_In     病人预交记录.结算号码%Type,
  缴款单位_In     病人预交记录.缴款单位%Type,
  单位开户行_In   病人预交记录.单位开户行%Type,
  单位帐号_In     病人预交记录.单位帐号%Type,
  摘要_In         病人预交记录.摘要%Type,
  操作员编号_In   病人预交记录.操作员编号%Type,
  操作员姓名_In   病人预交记录.操作员姓名%Type,
  领用id_In       票据使用明细.领用id%Type,
  预交类别_In     病人预交记录.预交类别%Type := Null,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  结算卡序号_In   病人预交记录.结算卡序号%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  合作单位_In     病人预交记录.合作单位%Type := Null,
  收款时间_In     病人预交记录.收款时间%Type := Null,
  操作类型_In     Integer := 0,
  结算性质_In     病人预交记录.结算性质%Type := Null,
  更新交款余额_In Number := 0,
  退款检查_In     Number := 0,
  强制退现_In     Number := 0,
  是否转账_In     Number := 0
) As
  ----------------------------------------------
  --操作类型_In:0-正常缴预交;1-存为划价单;3-余额退款
  --退款检查_In;0-忽略退款金额是否大于了病人余额；1-检查退款金额
  --更新交款余额_In：0-在本过程中更新；1-在 zl_人员缴款余额_Update 中更新
  --强制退现_In:0-不强制，1-三方卡或消费卡不允许退现但强制退现金给病人
  --是否转账_In:0-原样退或退现，1-转账到支持的三方卡上
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  v_性质     结算方式.性质%Type;
  v_打印id   票据打印内容.Id%Type;
  v_担保     病人信息.担保性质%Type;
  v_Date     Date;
  n_返回值   病人余额.预交余额%Type;
  n_组id     财务缴款分组.Id%Type;
  n_病人余额 病人余额.预交余额%Type;
  n_三方预交 病人余额.预交余额%Type;
  n_退款金额        病人预交记录.金额%Type;
  n_剩余款          病人预交记录.金额%Type;
  n_结帐id          病人结帐记录.ID%Type;
  
  Cursor C_冲预交 is
    Select A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明, 0 as 序号, A.收款时间, A.金额 AS 预交金
    From 病人预交记录 A Where RowNum < 2;
  r_冲预交 C_冲预交%Rowtype;
  
  Type Ty_剩余款 Is Ref Cursor;
  C_剩余款 Ty_剩余款; --动态游标变量 
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
  Elsif 操作类型_In = 3 Then
    --生成一条原预交ID的冲销记录，同时也生成一条余额退款的冲销记录
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    IF Nvl(卡类别id_In, 0) = 0 And Nvl(结算卡序号_In, 0) =0 then
      --退现，包括普通结算方式退现、强制退现、三方卡允许退现
      Open C_剩余款 For
           Select A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明, 
                   Min(decode(sign(A.金额),-1,0,1)) AS 序号, Min(decode(A.记录性质,1,A.收款时间,null)) AS 收款时间,  
                   Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) as 预交金
              From 病人预交记录 A, 医疗卡类别 B, 卡消费接口目录 C
             Where A.病人ID = 病人id_In And A.记录性质 In (1,11) And A.预交类别 = Nvl(预交类别_In, 2)
               And A.卡类别ID = B.ID(+) And Decode(强制退现_In, 1, 1, Nvl(B.是否退现, 1)) = 1
               And A.卡类别ID = C.编号(+) And Decode(强制退现_In, 1, 1, Nvl(C.是否退现, 1)) = 1
             Group By A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明
            Having Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) <> 0
             Order By 序号,收款时间;
    ElsIF Nvl(是否转账_In, 0) = 1 Then
      --转账，三方卡允许退现或者强制退现，传入的卡号可能不是原卡号,金额由同种卡类别的预交缴款分摊
      --目前只支持同一种卡转账
      Open C_剩余款 For
           Select A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明, 
                   Min(decode(sign(A.金额),-1,0,1)) AS 序号, Min(decode(A.记录性质,1,A.收款时间,null)) AS 收款时间,  
                   Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) as 预交金
              From 病人预交记录 A, 医疗卡类别 B
             Where A.病人ID = 病人id_In And A.记录性质 In (1,11) And A.预交类别 = Nvl(预交类别_In, 2)
               And A.卡类别ID = B.ID(+)
               And Nvl(卡类别id, 0) = Nvl(卡类别id_In, 0) And Nvl(交易流水号, '-') = Nvl(交易流水号_In, '-')
             Group By A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明
            Having Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) <> 0
             Order By 序号,收款时间;
    Else
      --退三方卡或者是消费卡，根据卡类别ID、结算卡序号、卡号、交易流水号缺省原预交记录，如果不能确定唯一则进行分摊
      Open C_剩余款 For
           Select A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明, 
                   Min(decode(sign(A.金额),-1,0,1)) AS 序号, Min(decode(A.记录性质,1,A.收款时间,null)) AS 收款时间,  
                   Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) as 预交金
              From 病人预交记录 A
             Where A.病人ID = 病人id_In And A.记录性质 In (1,11) And A.预交类别 = Nvl(预交类别_In, 2)
               And Nvl(A.卡类别id, 0) = Nvl(卡类别id_In, 0) And Nvl(A.结算卡序号, 0) = Nvl(结算卡序号_In, 0) 
               And Nvl(A.卡号, '-') = Nvl(卡号_In, '-') And Nvl(交易流水号, '-') = Nvl(交易流水号_In, '-')
             Group By A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明
            Having Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) <> 0
             Order By 序号,收款时间;
    End IF;
    
    n_剩余款 := -1 * 金额_In;
    n_退款金额 := 0;
    Loop
      Fetch C_剩余款
        Into r_冲预交;
      Exit When C_剩余款%NotFound;
      IF r_冲预交.NO <> 单据号_In Then
        IF n_剩余款 > r_冲预交.预交金 then
           n_退款金额 := r_冲预交.预交金;
           n_剩余款 := n_剩余款 - n_退款金额;
        Else
           n_退款金额 := n_剩余款;
           n_剩余款 := 0;
        End IF;
          	  
        IF nvl(n_退款金额, 0) <> 0 THEN 
          UPDATE 病人预交记录  SET 结帐ID = n_结帐id WHERE NO = r_冲预交.NO AND 记录性质 = 1 AND 结帐ID IS NULL;
          Insert Into 病人预交记录
             (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 操作员编号,
             收款时间, 操作员姓名, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 1, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In,
             v_Date, 操作员姓名_In, 摘要, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, n_退款金额, NULL
          From 病人预交记录
          Where NO = r_冲预交.NO And 记录性质 In (1, 11) And RowNum < 2;
        END IF;

        IF n_剩余款 = 0 Then 
          Exit;
        End IF;
      End IF;
    END LOOP;

    IF n_剩余款 <> 0 And Nvl(退款检查_In, 0) = 1 THEN 
      v_Err_Msg := '退款金额大于病人剩余预交余额。';
      Raise Err_Item;
    END IF;
    
    n_退款金额 := -1 * (-1 * 金额_In - n_剩余款);
    IF n_退款金额 <> 0 Then
      Update 病人预交记录 Set 结帐id = n_结帐id Where ID = Id_In;
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别,
         卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 单据号_In, 票据号_In, 11, 1, 病人id_In, Decode(主页id_In, 0, Null, 主页id_In),
         Decode(科室id_In, 0, Null, 科室id_In), NULL, 结算方式_In, 结算号码_In, v_Date, 缴款单位_In, 单位开户行_In, 单位帐号_In, 操作员编号_In, 操作员姓名_In,
         摘要_In, n_组id, 预交类别_In, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, n_结帐id, n_退款金额, NULL);
    End IF;
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

  If 金额_In < 0 Then
    Begin
      Select Nvl(预交余额, 0) - Nvl(费用余额, 0)
      Into n_病人余额
      From 病人余额
      Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 0) = Nvl(预交类别_In, 0);
    Exception
      When Others Then
        Null;
    End;
    --余额退款要考虑三方预交是否支持退现
    If 操作类型_In = 3 And Nvl(强制退现_In, 0) = 0 Then
      For c_三方预交 In (Select a.预交id, a.预交类别, a.卡类别id, a.结算卡序号 As 消费接口id, Nvl(b.编码, c.编号) As 编码, Nvl(b.名称, c.名称) As 名称,
                            Decode(b.编码, Null, c.是否全退, b.是否全退) As 是否全退, Decode(b.编码, Null, c.是否退现, b.是否退现) As 是否退现, a.卡号,
                            a.交易流水号, a.交易说明, a.预交余额
                     From (Select a.预交类别, Nvl(a.卡类别id, 0) As 卡类别id, Nvl(a.结算卡序号, 0) As 结算卡序号, a.卡号, a.交易流水号, a.交易说明,
                                   Max(Decode(Sign(金额), -1, Decode(a.记录状态, 1, 0, 2, 0, ID), ID)) As 预交id,
                                   Nvl(Sum(金额), 0) - Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交余额
                            From 病人预交记录 A
                            Where a.病人id = 病人id_In And (Nvl(a.结算卡序号, 0) <> 0 Or Nvl(卡类别id, 0) <> 0)
                            Group By a.预交类别, Nvl(a.卡类别id, 0), Nvl(a.结算卡序号, 0), a.卡号, a.交易流水号, a.交易说明
                            Having Nvl(Sum(金额), 0) - Nvl(Sum(Nvl(冲预交, 0)), 0) <> 0) A, 医疗卡类别 B, 卡消费接口目录 C
                     Where a.预交类别 = Nvl(预交类别_In, 0) And a.卡类别id = b.Id(+) And a.结算卡序号 = c.编号(+) And Nvl(a.预交余额, 0) <> 0
                     Order By 编码, a.卡号, a.交易流水号, a.交易说明) Loop
      
        If Instr(',7,8,', ',' || v_性质 || ',') = 0 And Nvl(c_三方预交.是否退现, 0) = 0 And Nvl(c_三方预交.预交余额, 0) > 0 Then
          n_三方预交 := Nvl(n_三方预交, 0) + Nvl(c_三方预交.预交余额, 0);
        Elsif Instr(',7,8,', ',' || v_性质 || ',') > 0 Then
          If Nvl(c_三方预交.卡号, '0') = Nvl(卡号_In, '0') And Nvl(c_三方预交.交易流水号, '0') = Nvl(交易流水号_In, '0') And
             Nvl(c_三方预交.交易说明, '0') = Nvl(交易说明_In, '0') Then
            n_三方预交 := Nvl(n_三方预交, 0) + Nvl(c_三方预交.预交余额, 0);
          End If;
        End If;
      End Loop;
    End If;
  
    If Instr(',7,8,', ',' || v_性质 || ',') > 0 And Nvl(n_三方预交, 0) < 0 And 操作类型_In = 3 Then
      v_Err_Msg := '退款金额大于病人三方预交金额。';
      Raise Err_Item;
    Elsif Nvl(n_病人余额, 0) < 0 And 退款检查_In = 1 Then
      v_Err_Msg := '退款金额大于病人剩余预交余额。';
      Raise Err_Item;
    Elsif Instr(',7,8,', ',' || v_性质 || ',') = 0 And Nvl(n_病人余额, 0) - Nvl(n_三方预交, 0) < 0 And 操作类型_In = 3 And 退款检查_In = 1 Then
      v_Err_Msg := '退款金额大于病人剩余预交余额。';
      Raise Err_Item;
    End If;
  End If;

  --人员缴款余额(现收)
  If Nvl(更新交款余额_In, 0) = 0 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 金额_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = 结算方式_In
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, 结算方式_In, 1, 金额_In);
      n_返回值 := 金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
    End If;
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
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_Insert;
/

--127450:李南春,2018-06-20,挂号按先进先出原则使用预交款
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

--127480:焦博,2018-06-20,修正Oracle过程Zl_挂号安排_Autoupdate
CREATE OR REPLACE Procedure Zl_挂号安排_Autoupdate Is
  Err_Item Exception;
  v_Date Date;
  -- v_Err_Msg Varchar2(100);
  v_Unitscount Number;
Begin
  --n_更新执行人 ：是否更新病人挂号记录 和门诊费用记录中的执行人
  --               如果计划中更改了 挂号项目 则不允许更新 病人挂号记录和门诊费用记录中的数据
  Select Sysdate Into v_Date From Dual;
  Select Count(0) Into v_Unitscount From 合作单位安排控制 Where Rownum = 1;

  For v_生效 In (Select ID, 安排id, 号码, 生效时间, 失效时间, 周日, 周一, 周二, 周三, 周四, 周五, 周六, 分诊方式, 序号控制, 执行时间 As 上次生效时间, 项目id, 医生姓名, 医生id,
                      序号, 科室id, 是否相同
               From (Select a.Id, a.安排id, a.号码, a.生效时间, a.失效时间, a.周日, a.周一, a.周二, a.周三, a.周四, a.周五, a.周六, a.分诊方式, a.序号控制,
                             b.执行时间, a.项目id, a.医生姓名, a.医生id, Nvl(b.执行计划id, 0) As 执行计划id,
                             Row_Number() Over(Partition By a.安排id Order By a.生效时间 Desc) As 顺序号, b.序号, b.科室id,
                             Case
                               When b.项目id = a.项目id And Nvl(a.医生id, 0) = Nvl(b.医生id, 0) And
                                    Nvl(a.医生姓名, '-') = Nvl(b.医生姓名, '-') Then
                                1
                               Else
                                0
                             End As 是否相同
                      From 挂号安排计划 A, 挂号安排 B
                      Where Sysdate Between a.生效时间 And a.失效时间 And a.安排id = b.Id And
                            a.实际生效 >= To_Date('3000-01-01', 'yyyy-mm-dd') And a.生效时间 + 0 <= Sysdate And 审核人 Is Not Null And
                            b.停用日期 Is Null)
               Where 顺序号 = 1 And ID <> Nvl(执行计划id, 0)) Loop
    Update 挂号安排计划
    Set 实际生效 = v_生效.上次生效时间
    Where 安排id = v_生效.安排id And 失效时间 <= v_生效.失效时间 And 生效时间 < Sysdate And ID <> v_生效.Id And
          实际生效 >= To_Date('3000-01-01', 'yyyy-mm-dd');
  
    Update 挂号安排
    Set 周日 = v_生效.周日, 周一 = v_生效.周一, 周二 = v_生效.周二, 周三 = v_生效.周三, 周四 = v_生效.周四, 周五 = v_生效.周五, 周六 = v_生效.周六,
        分诊方式 = v_生效.分诊方式, 序号控制 = v_生效.序号控制, 开始时间 = Sysdate, 终止时间 = v_生效.失效时间, 项目id = Nvl(v_生效.项目id, 项目id), 执行时间 = v_Date,
        执行计划id = v_生效.Id, 序号 = Decode(v_生效.是否相同, 1, 序号, 9999999), 医生姓名 = v_生效.医生姓名, 医生id = v_生效.医生id
    Where ID = v_生效.安排id;
  
    --重新调整序号
    If Nvl(v_生效.是否相同, 0) <> 1 Then
    
      Update 挂号安排 A
      Set 序号 = -1 * 序号
      Where 项目id = v_生效.项目id And a.科室id = v_生效.科室id And Nvl(a.医生姓名, '-') = Nvl(v_生效.医生姓名, '-') And
            Nvl(a.医生id, 0) = Nvl(v_生效.医生id, 0);
      For v_序号 In (Select a.Id, Rownum As 序号
                   From 挂号安排 A
                   Where a.项目id = v_生效.项目id And a.科室id = v_生效.科室id And Nvl(a.医生姓名, '-') = Nvl(v_生效.医生姓名, '-') And
                         Nvl(a.医生id, 0) = Nvl(v_生效.医生id, 0)
                   Order By a.Id) Loop
        Update 挂号安排 A Set 序号 = v_序号.序号 Where ID = v_序号.Id;
      End Loop;
    End If;
    Delete 挂号安排诊室 Where 号表id = v_生效.安排id;
    Insert Into 挂号安排诊室
      (号表id, 门诊诊室)
      Select v_生效.安排id, 门诊诊室 From 挂号计划诊室 Where 计划id = v_生效.Id;
    Delete 挂号安排限制 Where 安排id = v_生效.安排id;
    Insert Into 挂号安排限制
      (安排id, 限制项目, 限号数, 限约数)
      Select v_生效.安排id, 限制项目, 限号数, 限约数 From 挂号计划限制 Where 计划id = v_生效.Id;
    Delete 挂号安排时段 Where 安排id = v_生效.安排id;
    Insert Into 挂号安排时段
      (安排id, 序号, 开始时间, 结束时间, 限制数量, 是否预约, 星期)
      Select v_生效.安排id, 序号, 开始时间, 结束时间, 限制数量, 是否预约, 星期
      From 挂号计划时段
      Where 计划id = v_生效.Id;
    If Nvl(v_Unitscount, 0) > 0 Then
      Delete 合作单位安排控制 Where 安排id = v_生效.安排id;
      Insert Into 合作单位安排控制
        (安排id, 合作单位, 限制项目, 序号, 数量)
        Select v_生效.安排id, 合作单位, 限制项目, 序号, 数量 From 合作单位计划控制 Where 计划id = v_生效.Id;
    End If;
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号安排_Autoupdate;
/

--126104:陈刘,2018-06-14,体温单住院天数按第一份体温单开始时间计算
Create Or Replace Function Zl_Calcindaysnew
(
  文件id_In   In 病人护理文件.Id%Type,
  病人id_In   In 病案主页.病人id%Type,
  主页id_In   In 病案主页.主页id%Type,
  住院日期_In In Date := Sysdate --指定某一具体日期时的住院天数 
) Return Number As

  d_入院时间 病案主页.入院日期%Type;
  d_出院时间 病案主页.出院日期%Type;
  d_开始时间 病人护理文件.开始时间%Type;
  n_Days     Number(18);
  n_Bady     Number(18);
  n_Badybill Number(18);
  n_Addday   Number(18);
Begin

  n_Days     := 0;
  n_Bady     := 0;
  n_Badybill := 0;
  n_Addday   := 1;
  d_入院时间 := Null;
  d_出院时间 := Null;
  d_开始时间 := Null;
  --提取体温单开始时间 
  Begin
    Select 婴儿 Into n_Bady From 病人护理文件 Where ID = 文件id_In;
  Exception
    When Others Then
      n_Bady := 0;
  End;
  --提取第一个体温单的开始时间,106122-CL-07-02-21 
  Begin
    Select 开始时间
    Into d_开始时间
    From (Select 开始时间
           From 病人护理文件 A, 病历文件列表 B
           Where 病人id = 病人id_In And 主页id = 主页id_In And a.格式id = b.Id And b.保留 = -1
           Order By 开始时间)
    Where Rownum < 2;
  Exception
    When Others Then
      d_开始时间 := Null;
  End;

  --如果是婴儿开始时间以出生时间为准 
  If n_Bady <> 0 Then
    Begin
      Select a.出生时间
      Into d_入院时间
      From 病人新生儿记录 A
      Where a.病人id = 病人id_In And a.主页id = 主页id_In And a.序号 = n_Bady;
    Exception
      When Others Then
        d_入院时间 := d_开始时间;
    End;
    Begin
      Select Nvl(参数值, 0) Into n_Badybill From zlParameters Where 模块 = 1255 And 参数名 = '婴儿体温单首日天数显示0'; --婴儿体温单首日天数从0开始还是从1开始 
    Exception
      When Others Then
        n_Badybill := 0;
    End;
  End If;

  If d_入院时间 Is Null Then
    d_入院时间 := d_开始时间;
  End If;

  --提取体温单的实际结束时间 
  Begin
    Select Decode(Sign(a.出院时间 - b.发生时间), 1, a.出院时间, b.发生时间)
    Into d_出院时间
    From (Select Max(Nvl(终止时间, Sysdate)) As 出院时间, Max(病人id) 病人id, Max(主页id) 主页id
           From 病人变动记录
           Where 开始时间 Is Not Null And 病人id = 病人id_In And 主页id = 主页id_In) A,
         (Select Nvl(发生时间, Sysdate) 发生时间, 病人id, 主页id
           From (Select Max(发生时间) 发生时间, Max(a.病人id) 病人id, Max(a.主页id) 主页id
                  From 病人护理文件 A, 病人护理数据 B
                  Where a.Id = b.文件id And a.Id = 文件id_In)) B
    Where a.病人id = b.病人id And a.主页id = b.主页id;
  
  Exception
    When Others Then
      d_出院时间 := Sysdate;
  End;

  If n_Badybill = 1 Then
    n_Addday := 0;
  Else
    n_Addday := 1;
  End If;

  If d_入院时间 Is Not Null Then
    If Trunc(住院日期_In) > Trunc(d_出院时间) Then
      Select Trunc(d_出院时间) - Trunc(d_入院时间) + n_Addday Into n_Days From Dual;
    Else
      Select Trunc(住院日期_In) - Trunc(d_入院时间) + n_Addday Into n_Days From Dual;
    End If;
  End If;

  Return(n_Days);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Calcindaysnew;
/


--119477:陈刘,2018-07-09,分类汇总之后,再次分组汇总到体温单
Create Or Replace Procedure Zl_护理二次汇总_Update
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
  数据来源_In In 病人护理明细.数据来源%Type := 0,
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
      Select ID, 汇总类别
      Into n_记录id, n_汇总类别
      From 病人护理数据
      Where 文件id = n_文件id And 发生时间 = 发生时间_In;
    Exception
      When Others Then
        n_记录id := 0;
    End;
    
    Begin
      Select ID
      Into n_明细id
      From 病人护理明细
      Where 记录id = n_记录id And Nvl(项目序号, 0) = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') and  Nvl(记录组号, 1) = Nvl(记录组号_In, 1)  And 终止版本 Is Null ;
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
      If v_汇总文本 is null Then
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
                       体温部位_In, 数据来源_In, Null, 1, Null, b.记录人, Sysdate, 1
                From 护理记录项目 A, 病人护理明细 B
                Where a.项目序号 = b.项目序号(+) And b.终止版本(+) Is Null And b.记录id(+) = n_记录id And a.项目序号 = 项目序号_In And
                      Rownum < 2;
            
              If Sql%RowCount > 0 Then
                Int共用 := 1;
              End If;
            End If;
          Else
            Update 病人护理明细
            Set 记录内容 = 记录内容_In
            Where 记录id = n_Newid And 项目序号 = 项目序号_In And
                  Decode(n_项目性质, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无');
            If Sql%RowCount > 0 Then
              Int共用 := 1;
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

--119477:陈刘,2018-06-11,分类汇总之后,再次分组汇总到体温单

Create Or Replace Procedure Zl_病人护理数据_Collect
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
    Where b.Id = 文件id_In And a.病人id = b.病人id And a.主页id = b.主页id And a.婴儿 = b.婴儿 And A.格式id = c.Id And c.种类 = 3 And
          c.保留 = -1 And a.开始时间 < 发生时间_In And (a.结束时间 > 发生时间_In Or a.结束时间 Is Null);
    Select Max(a.记录id)
    Into v_来源id
    From 病人护理明细 A, 病人护理明细 B
    Where a.来源id = b.Id(+) And b.记录id = v_记录id;
  
    For r_List In (Select a.发生时间, b.项目序号, b.项目名称
                   From 病人护理数据 A, 病人护理明细 B
                   Where 文件id = n_文件id And a.Id = b.记录id And Instr(汇总文本, v_记录id) > 0 And b.数据来源 = 1 And 来源id Is Null) Loop
      Zl_护理二次汇总_Update(文件id_In, r_List.发生时间, 发生时间_In, 1, r_List.项目序号, Null, Null, Null, Null, 1, 1);
    End Loop;
    Delete 病人护理明细 Where 记录id = v_记录id;
    Delete 病人护理数据 Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
  
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人护理数据_Collect;
/
--119477:陈刘,2018-06-11,分类汇总之后,再次分组汇总到体温单

Create Or Replace Procedure Zl_护理记录项目_Update
(
  项目序号_In In 护理记录项目.项目序号%Type,
  项目名称_In In 护理记录项目.项目名称%Type,
  项目类型_In In 护理记录项目.项目类型%Type,
  项目长度_In In 护理记录项目.项目长度%Type,
  项目小数_In In 护理记录项目.项目小数%Type,
  项目单位_In In 护理记录项目.项目单位%Type,
  项目表示_In In 护理记录项目.项目表示%Type,
  项目值域_In In 护理记录项目.项目值域%Type,
  护理等级_In In 护理记录项目.护理等级%Type,
  分组名_In   In 护理记录项目.分组名%Type,
  项目id_In   In 护理记录项目.项目id%Type,
  应用方式_In In 护理记录项目.应用方式%Type,
  适用病人_In In 护理记录项目.适用病人%Type,
  项目性质_In In 护理记录项目.项目性质%Type := 1,
  应用场合_In In 护理记录项目.应用场合%Type := 0,
  说明_In     In 护理记录项目.说明%Type := Null,
  缺省值_In   In 护理记录项目.缺省值%Type := Null,
  分组汇总_In In 护理记录项目.分组汇总%Type := Null
) Is
  n_汇总 Number(1);
Begin
  n_汇总 := 0;
  Select Count(项目序号) Into n_汇总 From 护理记录项目 Where 项目序号 = 项目序号_In And 项目表示 = 4;
  If 分组汇总_In Is Null Then
    Update 护理记录项目
    Set 项目名称 = 项目名称_In, 项目类型 = 项目类型_In, 项目长度 = 项目长度_In, 项目小数 = 项目小数_In, 项目单位 = 项目单位_In, 项目表示 = 项目表示_In, 项目值域 = 项目值域_In,
        护理等级 = 护理等级_In, 分组名 = 分组名_In, 项目id = 项目id_In, 应用方式 = 应用方式_In, 适用病人 = 适用病人_In, 项目性质 = 项目性质_In, 应用场合 = 应用场合_In,
        说明 = 说明_In, 缺省值 = 缺省值_In
    Where 项目序号 = 项目序号_In;
  Else
    Update 护理记录项目 Set 分组汇总 = 分组汇总_In Where 项目序号 = 项目序号_In;
  End If;
  If 项目序号_In = 2 Then
    Update 护理记录项目
    Set 项目类型 = 项目类型_In, 项目长度 = 项目长度_In, 项目小数 = 项目小数_In, 项目单位 = 项目单位_In, 项目表示 = 项目表示_In, 项目值域 = 项目值域_In, 护理等级 = 护理等级_In,
        分组名 = 分组名_In, 项目性质 = 项目性质_In, 应用场合 = 应用场合_In, 说明 = 说明_In
    Where 项目序号 = -1;
  End If;

  If 项目序号_In = 4 Or 项目序号_In = 5 Then
    Update 护理记录项目
    Set 项目类型 = 项目类型_In, 项目长度 = 项目长度_In, 项目小数 = 项目小数_In, 项目单位 = 项目单位_In, 项目表示 = 项目表示_In, 护理等级 = 护理等级_In, 分组名 = 分组名_In,
        应用方式 = 应用方式_In, 适用病人 = 适用病人_In, 项目性质 = 项目性质_In, 应用场合 = 应用场合_In, 说明 = 说明_In
    Where 项目序号 In (4, 5);
  End If;
  If 项目表示_In = 4 Then
    Insert Into 护理汇总项目
      (序号, 父序号)
      Select 项目序号_In, Null From Dual Where Not Exists (Select 1 From 护理汇总项目 Where 序号 = 项目序号_In);
  Else
    If n_汇总 = 1 Then
      Delete 护理汇总项目 Where 序号 = 项目序号_In;
      Update 护理汇总项目 Set 父序号 = Null Where 父序号 = 项目序号_In;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_护理记录项目_Update;
/
--119477:陈刘,2018-06-11,分类汇总之后,再次分组汇总到体温单

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
  未记说明_In In 病人护理明细.未记说明%Type := Null, --入量导入存储医嘱ID:发送号
  分组汇总_In In number :=0
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
    If Nvl(n_汇总类别, 0) <> 0 and  分组汇总_In=0 Then
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
  
    If Nvl(n_汇总类别, 0) <> 0 and  分组汇总_In=0  Then
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

--126662:蒋敏,2018-06-11,历史修改把之前的过程覆盖了
Create Or Replace Procedure Zl_诊疗项目_Insert
(
  类别_In             In 诊疗项目目录.类别%Type := Null,
  分类id_In           In 诊疗项目目录.分类id%Type := Null,
  Id_In               In 诊疗项目目录.Id%Type,
  编码_In             In 诊疗项目目录.编码%Type := Null,
  名称_In             In 诊疗项目目录.名称%Type := Null,
  名称拼音_In         In 诊疗项目别名.简码%Type := Null,
  名称五笔_In         In 诊疗项目别名.简码%Type := Null,
  别名_In             诊疗项目目录.名称%Type := Null,
  别名拼音_In         诊疗项目别名.简码%Type := Null,
  别名五笔_In         诊疗项目别名.简码%Type := Null,
  操作类型_In         In 诊疗项目目录.操作类型%Type := Null,
  执行频率_In         In 诊疗项目目录.执行频率%Type := Null,
  单独应用_In         In 诊疗项目目录.单独应用%Type := Null,
  计算方式_In         In 诊疗项目目录.计算方式%Type := Null,
  计算单位_In         In 诊疗项目目录.计算单位%Type := Null,
  适用性别_In         In 诊疗项目目录.适用性别%Type := Null,
  执行安排_In         In 诊疗项目目录.执行安排%Type := Null,
  服务对象_In         In 诊疗项目目录.服务对象%Type := Null,
  组合项目_In         In 诊疗项目目录.组合项目%Type := Null,
  标本部位_In         In 诊疗项目目录.标本部位%Type := Null,
  手术操作id_In       In 疾病诊断对照.疾病id%Type := Null,
  执行科室_In         In 诊疗项目目录.执行科室%Type := Null,
  门诊执行_In         In 诊疗执行科室.执行科室id%Type := Null,
  住院执行_In         In 诊疗执行科室.执行科室id%Type := Null,
  定向执行_In         In Varchar2, --开单科室定向执行的说明串，以'|'分割，每个定向按'开单科室id^执行科室id'形式组织
  参考目录id_In       In 诊疗项目目录.参考目录id%Type := Null,
  应用范围_In         In Number := 0,
  录入限量_In         In 诊疗项目目录.录入限量%Type := Null,
  限量范围_In         In Number := 0,
  执行标记_In         In Number := 0,
  执行分类_In         In 诊疗项目目录.执行分类%Type := 0,
  站点_In             In 诊疗项目目录.站点%Type := Null,
  项目频率_In         In Varchar2 := Null, --该项目的频率设置串：编码|编码......
  计算规则_In         In 诊疗项目目录.计算规则%Type := Null,
  使用科室_In         In Varchar2 := Null, --使用科室的IDs,用逗号分隔
  使用科室应用范围_In In Number := 0, --使用科室应用的范围  0-本项，1-应用于同级，2-分类下所有，3-应用于当前类别
  First_In            In Number := 1, --First：1-需要删除执行科室，再新增，0-不删除执行科室，直接新增
  计算系数_In         In 诊疗项目目录.计算系数%Type := Null,
  输血检验对照_In     In Varchar2 :=Null,
  原始id_IN           In 诊疗项目目录.Id%Type:=0,
  试管编码_In         In 诊疗项目目录.试管编码%Type := Null  
) Is
  Type t_诊疗项目 Is Ref Cursor;
  c_诊疗项目   t_诊疗项目;
  t_Id         t_Numlist;
  v_Id         诊疗项目目录.Id%Type;
  v_Records    Varchar2(4000); --临时记录开单科室定向执行科室的字符串
  v_Currrec    Varchar2(1000); --包含在定向执行科室字符串中的一个定向
  v_Fields     Varchar2(1000);
  v_开单科室id 诊疗执行科室.开单科室id%Type := Null;
  v_执行科室id 诊疗执行科室.执行科室id%Type := Null;
  n_序号       Number;
  v_编号       Varchar2(1000);
  v_Strtmp     Varchar2(1000);
  v_Strinput   Varchar2(1000);
Begin
  If First_In = 1 Then
    Insert Into 诊疗项目目录
      (类别, 分类id, ID, 编码, 名称, 操作类型, 执行频率, 单独应用, 计算方式, 计算单位, 适用性别, 执行安排, 服务对象, 执行科室, 组合项目, 标本部位, 建档时间, 撤档时间, 参考目录id, 录入限量,
       执行标记, 执行分类, 计算规则, 站点, 计算系数,试管编码)
    Values
      (类别_In, 分类id_In, Id_In, 编码_In, 名称_In, 操作类型_In, 执行频率_In, 单独应用_In, 计算方式_In, 计算单位_In, 适用性别_In, 执行安排_In, 服务对象_In,
       执行科室_In, 组合项目_In, Decode(类别_In, 'D', Decode(组合项目_In, 1, '', 标本部位_In), 标本部位_In), Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), 参考目录id_In, 录入限量_In, 执行标记_In, 执行分类_In, 计算规则_In, 站点_In, 计算系数_In,试管编码_In);
    If 手术操作id_In Is Not Null Then
      Insert Into 疾病诊断对照 (疾病id, 诊断id, 手术id) Values (手术操作id_In, Null, Id_In);
    End If;
    If 名称拼音_In Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 名称_In, 1, 名称拼音_In, 1);
    End If;
    If 名称五笔_In Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 名称_In, 1, 名称五笔_In, 2);
    End If;
    If 别名_In Is Not Null And 别名拼音_In Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 别名_In, 9, 别名拼音_In, 1);
    End If;
    If 别名_In Is Not Null And 别名五笔_In Is Not Null Then
      Insert Into 诊疗项目别名 (诊疗项目id, 名称, 性质, 简码, 码类) Values (Id_In, 别名_In, 9, 别名五笔_In, 2);
    End If;
  End If;
  If 应用范围_In = 1 Then
    If 分类id_In Is Null Then
      Open c_诊疗项目 For
        Select ID From 诊疗项目目录 Where 分类id Is Null Order By 编码;
    Else
      Open c_诊疗项目 For
        Select ID From 诊疗项目目录 Where 分类id = 分类id_In Order By 编码;
    End If;
  Elsif 应用范围_In = 2 Then
    If 分类id_In Is Null Then
      Open c_诊疗项目 For
        Select c.Id
        From 诊疗项目目录 C, (Select ID From 诊疗分类目录 Start With 上级id Is Null Connect By Prior ID = 上级id) D
        Where d.Id = c.分类id
        Order By 编码;
    Else
      Open c_诊疗项目 For
        Select c.Id
        From 诊疗项目目录 C, (Select ID From 诊疗分类目录 Start With ID = 分类id_In Connect By Prior ID = 上级id) D
        Where d.Id = c.分类id
        Order By 编码;
    End If;
  Elsif 应用范围_In = 3 Then
    Open c_诊疗项目 For
      Select ID From 诊疗项目目录 Where 类别 = 类别_In Order By 编码;
  Else
    Open c_诊疗项目 For
      Select ID From 诊疗项目目录 Where ID = Id_In;
  End If;

  Loop
    Fetch c_诊疗项目
      Into v_Id;
    Exit When c_诊疗项目%NotFound;
  
    If First_In = 1 Then
      Delete From 诊疗执行科室 Where 诊疗项目id = v_Id;
      If 执行科室_In = 4 And 门诊执行_In Is Not Null Then
        Insert Into 诊疗执行科室 (诊疗项目id, 病人来源, 开单科室id, 执行科室id) Values (v_Id, 1, Null, 门诊执行_In);
      End If;
      If 执行科室_In = 4 And 住院执行_In Is Not Null Then
        Insert Into 诊疗执行科室 (诊疗项目id, 病人来源, 开单科室id, 执行科室id) Values (v_Id, 2, Null, 住院执行_In);
      End If;
    End If;
    If 执行科室_In <> 4 Or 定向执行_In Is Null Then
      v_Records := Null;
    Else
      v_Records := 定向执行_In || '|';
    End If;
  
    While v_Records Is Not Null Loop
      v_Currrec    := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields     := v_Currrec;
      v_开单科室id := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_执行科室id := To_Number(v_Fields);
      Insert Into 诊疗执行科室
        (诊疗项目id, 病人来源, 开单科室id, 执行科室id)
      Values
        (v_Id, Null, Decode(v_开单科室id, 0, Null, v_开单科室id), v_执行科室id);
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;
    If 应用范围_In <> 0 Then
      Update 诊疗项目目录 Set 执行科室 = 执行科室_In Where ID = v_Id;
    End If;
  End Loop;
  Close c_诊疗项目;

  If First_In = 1 Then
    If 类别_In = 'C' Or 类别_In = 'F' Or 类别_In = 'K' Then
      Insert Into 病历单据应用
        (病历文件id, 应用场合, 诊疗项目id)
        Select a.病历文件id, 1, Id_In
        From 病历单据应用 A, 诊疗项目目录 I
        Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 应用场合 = 1 And (服务对象_In = 0 Or 服务对象_In = 1) And Rownum < 2;
      Insert Into 病历单据应用
        (病历文件id, 应用场合, 诊疗项目id)
        Select a.病历文件id, 2, Id_In
        From 病历单据应用 A, 诊疗项目目录 I
        Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 应用场合 = 2 And (服务对象_In = 0 Or 服务对象_In = 2) And Rownum < 2;
    Elsif 类别_In = 'D' Or 类别_In = 'E' Then
      Insert Into 病历单据应用
        (病历文件id, 应用场合, 诊疗项目id)
        Select a.病历文件id, 1, Id_In
        From 病历单据应用 A, 诊疗项目目录 I
        Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 操作类型 = 操作类型_In And 应用场合 = 1 And (服务对象_In = 0 Or 服务对象_In = 1) And
              Rownum < 2;
      Insert Into 病历单据应用
        (病历文件id, 应用场合, 诊疗项目id)
        Select a.病历文件id, 2, Id_In
        From 病历单据应用 A, 诊疗项目目录 I
        Where a.诊疗项目id = i.Id And i.类别 = 类别_In And 操作类型 = 操作类型_In And 应用场合 = 2 And (服务对象_In = 0 Or 服务对象_In = 2) And
              Rownum < 2;
    End If;
  End If;

  If 限量范围_In = 1 Then
    If 分类id_In Is Null Then
      Update 诊疗项目目录 Set 录入限量 = 录入限量_In Where 分类id Is Null;
    Else
      Update 诊疗项目目录 Set 录入限量 = 录入限量_In Where 分类id = 分类id_In;
    End If;
  Elsif 限量范围_In = 2 Then
    If 分类id_In Is Null Then
      Update 诊疗项目目录
      Set 录入限量 = 录入限量_In
      Where 分类id In (Select ID From 诊疗分类目录 Start With 上级id Is Null Connect By Prior ID = 上级id);
    Else
      Update 诊疗项目目录
      Set 录入限量 = 录入限量_In
      Where 分类id In (Select ID From 诊疗分类目录 Start With ID = 分类id_In Connect By Prior ID = 上级id);
    End If;
  Elsif 限量范围_In = 3 Then
    Update 诊疗项目目录 Set 录入限量 = 录入限量_In Where 类别 = 类别_In;
  Elsif 限量范围_In = 4 Then
    Update 诊疗项目目录 Set 录入限量 = 录入限量_In;
  End If;

  --该项目的频率设置
  If 类别_In <> 'C' Then
    Delete 诊疗用法用量 Where 项目id = Id_In;
    If 项目频率_In Is Not Null Then
      v_Strinput := 项目频率_In || '|';
      n_序号     := 0;
    
      While v_Strinput Is Not Null Loop
        v_Strtmp := Substr(v_Strinput, 1, Instr(v_Strinput, '|') - 1);
        v_编号   := v_Strtmp;
        n_序号   := n_序号 + 1;
      
        Insert Into 诊疗用法用量 (项目id, 性质, 频次) Values (Id_In, n_序号, v_编号);
        v_Strinput := Replace('|' || v_Strinput, '|' || v_Strtmp || '|');
      End Loop;
    End If;
  End If;
  --使用科室
  If 使用科室应用范围_In = 1 Then
    If 分类id_In Is Null Then
      Open c_诊疗项目 For
        Select ID From 诊疗项目目录 Where 分类id Is Null Order By 编码;
    Else
      Open c_诊疗项目 For
        Select ID From 诊疗项目目录 Where 分类id = 分类id_In Order By 编码;
    End If;
  Elsif 使用科室应用范围_In = 2 Then
    If 分类id_In Is Null Then
      Open c_诊疗项目 For
        Select c.Id
        From 诊疗项目目录 C, (Select ID From 诊疗分类目录 Start With 上级id Is Null Connect By Prior ID = 上级id) D
        Where d.Id = c.分类id
        Order By 编码;
    Else
      Open c_诊疗项目 For
        Select c.Id
        From 诊疗项目目录 C, (Select ID From 诊疗分类目录 Start With ID = 分类id_In Connect By Prior ID = 上级id) D
        Where d.Id = c.分类id
        Order By 编码;
    End If;
  Elsif 使用科室应用范围_In = 3 Then
    Open c_诊疗项目 For
      Select ID From 诊疗项目目录 Where 类别 = 类别_In Order By 编码;
  Else
    Open c_诊疗项目 For
      Select ID From 诊疗项目目录 Where ID = Id_In;
  End If;
  Fetch c_诊疗项目 Bulk Collect
    Into t_Id;
  Close c_诊疗项目;

  Forall I In 1 .. t_Id.Count
    Delete 诊疗适用科室 Where 项目id = t_Id(I) And Instr(',' || 使用科室_In || ',', ',' || 科室id || ',') = 0;

  If 使用科室_In Is Not Null Then
    Forall I In 1 .. t_Id.Count
      Insert Into 诊疗适用科室
        (项目id, 科室id)
        Select t_Id(I), Column_Value
        From Table(f_Num2list(使用科室_In)) A
        Where Not Exists (Select 1 From 诊疗适用科室 Where 科室id = Column_Value And 项目id = t_Id(I));
  End If;
  --输血检验对照
  If 类别_In = 'K' And 输血检验对照_In Is Not Null Then
    v_Strinput := 输血检验对照_In || '|';
  
    While v_Strinput Is Not Null Loop
      v_Strtmp := Substr(v_Strinput, 1, Instr(v_Strinput, '|') - 1);
      v_Id     := v_Strtmp;
    
      Insert Into 输血检验对照 (项目id, 检验项目id) Values (Id_In, v_Id);
      v_Strinput := Replace('|' || v_Strinput, '|' || v_Strtmp || '|');
    End Loop;
  End If;
  
  if 原始id_IN<>0 then
    Zl_诊疗收费_Insert(id_In,原始id_IN);
  end if;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗项目_Insert;
/
--125837:涂建华,2018-06-07,排队显示方式调整
CREATE OR REPLACE Procedure Zl_排队叫号队列_清除业务
(
       业务类型_IN 排队叫号队列.业务类型%Type,
       有效天数_IN Number := 1
)
Is
Begin
  case 业务类型_IN
    when -1 then Null;
    else
      --清除当前业务类型，而且时间在有效时间之前的排队信息
      delete from 排队语音呼叫 where 业务类型 = 业务类型_IN And 生成时间 <=  sysdate - (1 / 48);
     
      Delete From 排队叫号队列 
      Where 业务类型 = 业务类型_IN And To_Number(Trunc(Sysdate - 排队叫号队列.排队时间)) >= 有效天数_In;
  end case;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_排队叫号队列_清除业务;
/

--126057:刘涛,2018-06-07,Zl_药品收发主表_Insert修改
Create Or Replace Procedure Zl_药品收发主表_Insert
(
  No_In     In 药品收发主表.No%Type,
  单据_In   In 药品收发主表.单据%Type,
  库房id_In In 药品收发主表.库房id%Type,
  对方部门id_In In 药品收发主表.对方部门id%Type
) Is
  n_Count Number;
  n_Id    药品收发主表.Id%Type;
Begin
  n_Count := 0;
  Begin
    Select 1 Into n_Count From 药品收发主表 Where NO = No_In And 单据 = 单据_In And 库房id = 库房id_In And 对方部门id = 对方部门id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    Select 药品收发主表_Id.Nextval Into n_Id From Dual;
    Insert Into 药品收发主表 (ID, NO, 单据, 库房id, 打印状态, 对方部门id) Values (n_Id, No_In, 单据_In, 库房id_In, 1, 对方部门id_In);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发主表_Insert;
/

--126591:刘兴洪,2018-06-04,增加误差费的处理.
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
  --        <XM>姓名</XM>               //姓名
  --        <SFZH>身份证号</SFZH>       //身份证号
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
  v_姓名       病人信息.姓名%Type;
  v_身份证号   病人信息.身份证号%Type;
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
  d_结束日期   Date;
  d_最小日期   Date;
  d_最大日期   Date;

  n_预交id     病人预交记录.Id%Type;
  n_科室id     病案主页.入院科室id%Type;
  n_结算卡序号 卡消费接口目录.编号%Type;
  n_时间类型   Number(3);
  v_Ids        Varchar2(20000);
  v_No         病人结帐记录.No%Type;
  v_预交no     病人预交记录.No%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_结算卡号   病人预交记录.卡号%Type;
  v_结算方式   病人预交记录.结算方式%Type;
  v_Temp       Varchar2(500);
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;

  n_Count    Number(18);
  n_Number   Number(2);
  n_费用id   门诊费用记录.Id%Type;
  n_记录性质 门诊费用记录.记录性质%Type;
  v_费用no   门诊费用记录.No%Type;
  n_序号     门诊费用记录.序号%Type;
  n_记录状态 门诊费用记录.记录状态%Type;
  n_执行状态 门诊费用记录.执行状态%Type;
  n_未结金额 门诊费用记录.实收金额%Type;
  n_结帐金额 门诊费用记录.实收金额%Type;
  n_误差费   门诊费用记录.实收金额%Type;

  v_卡类别     三方交易记录.类别%Type;
  v_消费卡结算 Varchar2(20000);

  Type t_费用结算明细 Is Ref Cursor;
  c_费用结算明细 t_费用结算明细;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/JSLX')),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_主页id, n_病人id, n_结帐总额, n_结算类型, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_结算类型 := Nvl(n_结算类型, 2);
  If n_结算类型 = 1 And Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
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

  For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Not (c_交易记录.结算卡类别 Is Null Or Nvl(c_交易记录.是否消费卡, '0') = '1' Or Nvl(c_交易记录.是否冲预交, 0) = 1) Then
    
      Select Decode(Translate(Nvl(c_交易记录.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
      Into n_Count
      From Dual;
    
      If Nvl(n_Count, 0) = 1 Then
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_交易记录.结算卡类别);
      Else
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_交易记录.结算卡类别;
      End If;
    
      If v_卡类别 Is Null Then
        v_Err_Msg := '不支持的结算方式,请检查！';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, 2) = 0 Then
        v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易!';
        Raise Err_Special;
      End If;
    End If;
  End Loop;

  Select Max(入院科室id) Into n_科室id From 病案主页 Where 病人id = n_病人id And 主页id = n_主页id;
  Select 病人结帐记录_Id.Nextval, Sysdate, Nextno(15) Into n_结帐id, d_结帐时间, v_No From Dual;

  n_时间类型 := Zl_Getsysparameter('结帐费用时间', 1137);

  Select 病人结帐记录_Id.Nextval, Sysdate, Nextno(15) Into n_结帐id, d_结帐时间, v_No From Dual;

  If n_结算类型 = 2 Then
    Open c_费用结算明细 For
      Select Max(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
             Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
             Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
      From 住院费用记录
      Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1
      Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
      Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, 序号;
  Else
  
    Open c_费用结算明细 For
      Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
             Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
             Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
      From 门诊费用记录
      Where 病人id = n_病人id And 记录状态 <> 0 And 记帐费用 = 1
      Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
      Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
      Union All
      Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
             Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
             Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
      From 住院费用记录
      Where 病人id = n_病人id And 记录状态 <> 0 And 记帐费用 = 1 And Mod(记录性质, 10) = 5
      Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
      Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, 序号;
  End If;

  n_待结帐金额 := 0;
  Loop
    Fetch c_费用结算明细
      Into n_费用id, n_记录性质, v_费用no, n_序号, n_记录状态, n_执行状态, d_最小日期, d_最大日期, n_未结金额, n_结帐金额;
    Exit When c_费用结算明细%NotFound;
  
    n_待结帐金额 := n_待结帐金额 + Nvl(n_未结金额, 0);
    If d_开始日期 Is Null Then
      d_开始日期 := d_最小日期;
    Elsif d_开始日期 > d_最小日期 Then
      d_开始日期 := d_最小日期;
    End If;
    If d_结束日期 Is Null Then
      d_结束日期 := d_最大日期;
    Elsif d_结束日期 < d_最大日期 Then
      d_结束日期 := d_最大日期;
    End If;
  
    If Nvl(n_结帐金额, 0) = 0 Then
      If n_费用id Is Not Null Then
        If Length(v_Ids || ',' || n_费用id) > 4000 Then
          v_Ids := Substr(v_Ids, 2);
          Zl_结帐费用记录_Batch(v_Ids, n_病人id, n_结帐id);
          v_Ids := '';
        End If;
        v_Ids := v_Ids || ',' || n_费用id;
      Else
        Zl_结帐费用记录_Insert(0, v_费用no, n_记录性质, n_记录状态, n_执行状态, n_序号, n_未结金额, n_结帐id);
      End If;
    Else
      Zl_结帐费用记录_Insert(0, v_费用no, n_记录性质, n_记录状态, n_执行状态, n_序号, n_未结金额, n_结帐id);
    End If;
  
  End Loop;

  If v_Ids Is Not Null Then
    v_Ids := Substr(v_Ids, 2);
    Zl_结帐费用记录_Batch(v_Ids, n_病人id, n_结帐id);
  End If;
  n_待结帐金额 := Round(n_待结帐金额, 6);

  If n_待结帐金额 <> Nvl(n_结帐总额, 0) Then
    v_Err_Msg := '传入的结帐金额与实际结帐金额不符,不允许结算!';
    Raise Err_Item;
  End If;

  Zl_病人结帐记录_Insert(n_结帐id, v_No, n_病人id, d_结帐时间, d_开始日期, d_结束日期, 0, 0, n_主页id, Null, 2, Null, n_结算类型);

  n_结帐金额 := 0;
  n_Count    := 0;
  For r_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    v_卡类别   := r_结算方式.结算方式;
    n_结帐金额 := n_结帐金额 + Nvl(r_结算方式.结算金额, 0);
  
    If Nvl(r_结算方式.是否冲预交, 0) = 0 Then
      --付款
      If n_Count = 1 Then
        v_Err_Msg := '结帐结算暂不支持多种结算方式!';
        Raise Err_Item;
      End If;
      n_卡类别id := Null;
      If r_结算方式.结算卡类别 Is Not Null Then
        Select Decode(Translate(Nvl(r_结算方式.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
      
        If Nvl(r_结算方式.是否消费卡, 0) = 1 Then
          If Nvl(n_Number, 0) = 1 Then
            Select Max(编号), Max(结算方式), Max(名称)
            Into n_结算卡序号, v_结算方式, v_卡类别
            From 卡消费接口目录
            Where 编号 = n_卡类别id And Nvl(启用, 0) = 1;
          Else
            Select Max(编号), Max(结算方式), Max(名称)
            Into n_结算卡序号, v_结算方式, v_卡类别
            From 卡消费接口目录
            Where 名称 = r_结算方式.结算卡类别 And Nvl(启用, 0) = 1;
          
          End If;
        
          If n_结算卡序号 Is Null Then
            v_Err_Msg := '未找到对应的消费卡信息';
            Raise Err_Item;
          
          End If;
          n_卡类别id := Null;
        
        Else
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID), Max(结算方式), Max(名称)
            Into n_卡类别id, v_结算方式, v_卡类别
            From 医疗卡类别
            Where ID = n_卡类别id And Nvl(是否启用, 0) = 1;
          Else
            Select Max(ID), Max(结算方式), Max(名称)
            Into n_卡类别id, v_结算方式, v_卡类别
            From 医疗卡类别
            Where 名称 = r_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
          End If;
        
          If n_卡类别id Is Null Then
            v_Err_Msg := '未找到对应的医疗卡信息!';
            Raise Err_Item;
          End If;
        End If;
      End If;
    
      If n_卡类别id Is Not Null Then
        --三方卡,生成住院预交款
        v_结算卡号 := r_结算方式.结算卡号;
        If r_结算方式.结算金额 > 0 Then
          --充值部分不应该算作本次结帐
          n_结帐金额 := n_结帐金额 - Nvl(r_结算方式.结算金额, 0);
          Select 病人预交记录_Id.Nextval, Nextno(11) Into n_预交id, v_预交no From Dual;
          Zl_病人预交记录_Insert(n_预交id, v_预交no, Null, n_病人id, n_主页id, n_科室id, r_结算方式.结算金额, v_结算方式, '', '', '', '', '',
                           v_操作员编码, v_操作员姓名, Null, n_结算类型, n_卡类别id, Null, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明, Null,
                           d_结帐时间, 0);
          n_预交充值 := Nvl(n_预交充值, 0) + r_结算方式.结算金额;
          For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                         From Table(Xmlsequence(Extract(r_结算方式.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
            Zl_三方结算交易_Insert(n_卡类别id, 0, r_结算方式.结算卡号, n_预交id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 1);
          
          End Loop;
        
        Else
        
          Zl_结帐缴款记录_Insert(v_No, n_病人id, n_主页id, n_科室id, v_结算方式, Null, r_结算方式.结算金额, n_结帐id, v_操作员编码, v_操作员姓名, d_结帐时间,
                           Null, Null, Null, Null, Null, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明);
        
          For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                         From Table(Xmlsequence(Extract(r_结算方式.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
            Zl_三方结算交易_Insert(n_卡类别id, 0, r_结算方式.结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
          End Loop;
        
        End If;
      
      Else
        If n_结算卡序号 Is Not Null Then
          --消费卡
          v_消费卡结算 := Nvl(v_消费卡结算, '') || '||' || n_结算卡序号 || '|' || r_结算方式.结算卡号 || '|0|' || r_结算方式.结算金额;
        Else
          --其他结算
          Zl_结帐缴款记录_Insert(v_No, n_病人id, n_主页id, n_科室id, v_结算方式, Null, r_结算方式.结算金额, n_结帐id, v_操作员编码, v_操作员姓名, d_结帐时间,
                           Null, Null, Null, Null, Null, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明);
        End If;
      End If;
      n_Count := 1;
    Else
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
  
    Update 三方交易记录
    Set 业务结算id = n_结帐id
    Where 流水号 = Nvl(r_结算方式.交易流水号, '-') And 类别 = v_卡类别 And 业务类型 = 2;
  End Loop;

  --消费卡处理
  If v_消费卡结算 Is Not Null Then
    v_消费卡结算 := Substr(v_消费卡结算, 3);
  End If;

  n_误差费   := Round(Nvl(n_结帐总额, 0) - Nvl(n_结帐金额, 0), 6);
  v_结算方式 := Null;
  If Abs(Nvl(n_误差费, 0)) > 1 Then
    v_Err_Msg := '计算的误差金额大于了1.00或小于-1.00元,不允许结帐操作,请检查!';
    Raise Err_Item;
  End If;

  n_结帐总额 := n_结帐金额;

  n_结帐金额 := 0;
  If Nvl(n_误差费, 0) <> 0 Then
    Select Nvl(Max(名称), '误差费') Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 9;
    n_结帐金额 := Nvl(n_误差费, 0);
  End If;
  If Nvl(n_误差费, 0) <> 0 Or v_消费卡结算 Is Not Null Then
    Zl_结帐缴款记录_Insert(v_No, n_病人id, n_主页id, n_科室id, v_结算方式, Null, n_结帐金额, n_结帐id, v_操作员编码, v_操作员姓名, d_结帐时间, Null, Null,
                     Null, Null, Null, Null, Null, Null, Null, v_消费卡结算);
  End If;

  --检查结算信息总额与结算总额是否正确
  Select Sum(冲预交) Into n_结帐金额 From 病人预交记录 Where 结帐id = n_结帐id;
  If Round(n_结帐金额, 6) <> Round(n_结帐总额, 6) Then
  
    v_Err_Msg := '传入的结算合计金额与实际结帐金额合计不符,不允许结算!';
    Raise Err_Item;
  End If;

  Update 病人预交记录 Set 校对标志 = 0 Where 结帐id = n_结帐id And Nvl(校对标志, 0) <> 0;
  v_Temp := '<CZSJ>' || To_Char(d_结帐时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Third_Settlement;
/

--13874:李南春,2018-06-04,挂号部分退费退到指定的结算方式
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

--126862:焦博,2018-06-20,退病历费或附加费时,不应该更新病人挂号汇总
--13874:李南春,2018-06-04,挂号部分退费退到指定的结算方式
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

  Select 出诊记录id, 号序 Into n_出诊记录id, n_序号 From 病人挂号记录 Where NO = 单据号_In And Rownum < 2;

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
  If Nvl(退费类型_In, 0) <> 2 Then
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

--127548:黄捷,2018-06-21,RIS取消登记时删除预约
--125867:黄捷,2018-05-23,RIS接口出院患者有未缴费用不允许执行费用
Create Or Replace Package b_Zlxwinterface Is
  Type t_Refcur Is Ref Cursor;

  --1、接收RIS状态改变
  Procedure Receiverisstate
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
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
  Procedure Receivereport
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
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

  --7、插入医嘱操作失败记录
  Procedure Ris医嘱失败记录_Insert
  (
    病人来源_In   In Ris医嘱失败记录.病人来源%Type,
    病人id_In     In Ris医嘱失败记录.病人id%Type,
    主页id_In     In Ris医嘱失败记录.主页id%Type,
    挂号单号_In   In Ris医嘱失败记录.挂号单号%Type,
    发送号_In     In Ris医嘱失败记录.发送号%Type,
    体检任务id_In In Ris医嘱失败记录.体检任务id%Type,
    体检报到号_In In Ris医嘱失败记录.体检报到号%Type,
    发送类型_In   In Ris医嘱失败记录.发送类型%Type
  );

  --8、更新医嘱操作失败记录
  Procedure Ris医嘱失败记录_重发
  (
    Id_In       In Ris医嘱失败记录.Id%Type,
    操作类型_In In Number
  );

  --9、销账后新建住院记账单据
  Procedure 病人医嘱_重建单据
  (
    医嘱id_In In 病人医嘱发送.医嘱id%Type,
    No_In     In 病人医嘱发送.No%Type,
    Action_In In Number
  );

  --10、打印RIS检查预约通知单
  Procedure Ris检查预约_打印(医嘱id_In In Ris检查预约.医嘱id%Type);

  --11、更新RIS分科室启用参数
  Procedure Ris启用控制_Update
  (
    检查类型_In Ris启用控制.检查类型%Type,
    场合_In     Ris启用控制.场合%Type,
    部门ids_In  Varchar2,
    启用类型_In Number
  );

  --12、删除RIS分科室启用参数
  Procedure Ris启用控制_Delete;

  --13、根据元素名提取信息
  Function Ris_Replace_Element_Value
  (
    元素名_In   In 诊治所见项目.中文名%Type,
    病人id_In   In 电子病历记录.病人id%Type,
    就诊id_In   In 电子病历记录.主页id%Type,
    病人来源_In In 电子病历记录.病人来源%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type
  ) Return Varchar2;

  --14、删除RIS分院设置参数
  Procedure Ris分院设置_Delete;

  --15、更新RISRis分院设置参数
  Procedure Ris分院设置_Update
  (
    Id_In           Ris分院设置.Id%Type,
    医院名称_In     Ris分院设置.医院名称%Type,
    医院代码_In     Ris分院设置.医院代码%Type,
    用户名_In       Ris分院设置.用户名%Type,
    密码_In         Ris分院设置.密码%Type,
    数据库服务名_In Ris分院设置.数据库服务名%Type
  );
End b_Zlxwinterface;
/
Create Or Replace Package Body b_Zlxwinterface Is

  --1、接收RIS状态改变
  Procedure Receiverisstate
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
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
      Select a.Id, a.相关id, Nvl(a.相关id, a.Id) As 组id, a.诊疗类别, a.病人来源, a.执行科室id, b.执行过程
      From 病人医嘱记录 A, 病人医嘱发送 B
      Where a.Id = b.医嘱id And ID = 医嘱id_In;
    r_Adviceinfo c_Adviceinfo%RowType;
  
    v_执行状态 病人医嘱发送.执行状态%Type;
    v_执行过程 病人医嘱发送.执行过程%Type;
    n_执行     Number; --标记是否需要更新状态，1：需要更新，其他不需要更新
    v_Count    Number;
    v_完成人   病人医嘱发送.完成人%Type;
    v_完成时间 病人医嘱发送.完成时间%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    v_执行状态 := 0;
    v_执行过程 := 0;
  
    --提取医嘱的主医嘱ID，及组ID
    Open c_Adviceinfo;
    Fetch c_Adviceinfo
      Into r_Adviceinfo;
    Close c_Adviceinfo;
  
    --根据状态_IN执行医嘱
    ---1-删除；0-预约(在RIS中实际上就是删除)；1-登记；3-检查完成；4-检查中止；9-初步报告；12-报告审核；13-取消审核；14-报告删除；15-发放
  
    If 状态_In = -1 Or 状态_In = 0 Then
      v_执行状态 := 0; --未执行
      v_执行过程 := 0;
    Elsif 状态_In = 1 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 2; --已报到
    Elsif 状态_In = 3 Or 状态_In = 14 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 3; --已检查
    Elsif 状态_In = 4 Then
      --不改变
      v_执行状态 := v_执行状态;
    Elsif 状态_In = 9 Or 状态_In = 13 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 4; --已报告
    Elsif 状态_In = 12 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 5; --已审核
    Elsif 状态_In = 15 Then
      v_执行状态 := 1; --完全执行
      v_执行过程 := 6; --已完成
      v_完成人   := 操作人员_In;
      v_完成时间 := 执行时间_In;
    End If;
  
    n_执行 := 1; --默认都要更新状态
  
    If 状态_In = 13 Or 状态_In = 14 Then
      --删除对应报告数据
      Delete From 电子病历记录
      Where ID = (Select 病历id From 病人医嘱报告 Where 医嘱id = 医嘱id_In And Risid = Risid_In);
      Delete From 病人医嘱报告 Where 医嘱id = 医嘱id_In And Risid = Risid_In;
    
      --删除后判断是否还存在报告，若存在则医嘱状态保持不变，若报告全部删除则更新医嘱状态
      Select Count(1) Into v_Count From 病人医嘱报告 Where 医嘱id = 医嘱id_In;
    
      If v_Count > 0 Then
        n_执行 := 0; --若存在则医嘱状态保持不变
      End If;
    End If;
  
    --如果是删除，则删除已有的预约信息
    If 状态_In = -1 Or 状态_In = 0 Then
      Zl_Ris检查预约_Delete(医嘱id_In);
    End If;
  
    --如果是登记，先判断此检查是否未执行
    If 状态_In = 1 Then
      If r_Adviceinfo.执行过程 >= 3 Then
        v_Error := '患者已经做过检查了，不能重复登记。';
        Raise Err_Custom;
      End If;
    End If;
  
    --开始执行医嘱
    If n_执行 = 1 Then
      If Nvl(单独执行_In, 0) = 1 Then
        -- 单个部位医嘱单独执行
        Update 病人医嘱发送
        Set 执行状态 = v_执行状态, 执行过程 = v_执行过程, 执行说明 = 执行说明_In, 完成人 = v_完成人, 完成时间 = v_完成时间
        Where 医嘱id = 医嘱id_In;
      Else
        Update 病人医嘱发送
        Set 执行状态 = v_执行状态, 执行过程 = v_执行过程, 执行说明 = 执行说明_In, 完成人 = v_完成人, 完成时间 = v_完成时间
        Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = r_Adviceinfo.组id Or 相关id = r_Adviceinfo.组id));
      End If;
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
    v_Error    Varchar2(255);
    Err_Custom Exception;
  Begin
  
    --取主医嘱ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    Select 发送号, 执行过程 Into v_发送号, v_执行过程 From 病人医嘱发送 Where 医嘱id = r_Advice.组id;
  
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
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
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
  
    Cursor c_Advice Is
      Select ID, 相关id, Nvl(相关id, ID) As 组id From 病人医嘱记录 Where ID = 医嘱id_In;
    r_Advice c_Advice%RowType;
  
    v_发送号 病人医嘱发送.发送号%Type;
    v_Count  Number;
    v_Error  Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    --取主医嘱ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    --先检查是否已经出院的住院病人，已经预出院或者出院的检查申请，不允执行费用
    Select Count(*)
    Into v_Count
    From 病人医嘱记录 A, 病案主页 B
    Where a.病人id = b.病人id And a.主页id = b.主页id And (b.出院日期 Is Not Null Or b.状态 = 3) And a.Id = r_Advice.组id;
  
    If v_Count > 0 Then
      v_Error := '住院病人已经出院或者预出院，不能取消费用。';
      Raise Err_Custom;
    End If;
  
    Select 发送号 Into v_发送号 From 病人医嘱发送 Where 医嘱id = r_Advice.组id;
  
    --调用统一的医嘱执行Cancel过程
    Zl_病人医嘱执行_Cancel(医嘱id_In, v_发送号, Null, 单独执行_In, 执行部门id_In, 操作员编号_In, 操作员姓名_In);
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像费用执行_Cancel;

  --4、接收RIS的报告
  Procedure Receivereport
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
    报告所见_In 电子病历内容.内容文本%Type,
    报告意见_In 电子病历内容.内容文本%Type,
    报告建议_In 电子病历内容.内容文本%Type,
    报告医生_In 电子病历记录.创建人%Type
  ) Is
    --提取病人医嘱及报告的相关信息
    Cursor c_Advice
    (
      v_组id  Number,
      v_Risid Number
    ) Is
      Select e.Id, e.病人来源, e.病人id, e.主页id, e.婴儿, e.病人科室id, e.文件id, e.病历种类, e.病历名称, f.病历id, e.执行科室id
      From (Select c.Id, c.病人来源, c.病人id, c.主页id, c.婴儿, c.病人科室id, c.文件id, d.种类 病历种类, d.名称 病历名称, c.执行科室id
             From (Select a.Id, a.病人来源, a.病人id, a.主页id, a.婴儿, a.病人科室id, b.病历文件id 文件id, a.执行科室id
                    From 病人医嘱记录 A, 病历单据应用 B
                    Where a.Id = v_组id And a.诊疗项目id = b.诊疗项目id(+) And b.应用场合(+) = Decode(a.病人来源, 2, 2, 4, 4, 1)) C,
                  病历文件列表 D
             Where c.文件id = d.Id(+)) E, 病人医嘱报告 F
      Where e.Id = f.医嘱id(+) And f.Risid(+) = v_Risid;
  
    --查找文件的组成元素
    Cursor c_File(v_File Number) Is
      Select a.Id, a.文件id, a.父id, a.对象序号, a.对象类型, a.对象标记, a.保留对象, a.对象属性, a.内容行次, a.内容文本, a.是否换行, a.预制提纲id, a.复用提纲,
             a.使用时机, a.诊治要素id, a.替换域, a.要素名称, a.要素类型, a.要素长度, a.要素小数, a.要素单位, a.要素表示, a.输入形态, a.要素值域
      From 病历文件结构 A
      Where a.文件id = v_File
      Order By a.对象序号;
  
    Cursor c_Report(v_电子病历记录id Number) Is
      Select b.Id, a.内容文本
      From 电子病历内容 A, 电子病历内容 B
      Where a.对象类型 = 3 And a.Id = b.父id And b.对象类型 = 2 And b.终止版 = 0 And a.文件id = v_电子病历记录id;
  
    Cursor c_Content
    (
      v_文件id Number,
      v_表格id Number
    ) Is
      Select a.Id, a.文件id, a.父id, a.对象序号, a.对象类型, a.对象标记, a.保留对象, a.对象属性, a.内容行次, a.内容文本, a.是否换行, a.预制提纲id, a.复用提纲,
             a.使用时机, a.诊治要素id, a.替换域, a.要素名称, a.要素类型, a.要素长度, a.要素小数, a.要素单位, a.要素表示, a.输入形态, a.要素值域
      From 病历文件结构 A
      Where 文件id = v_文件id And 父id = v_表格id;
  
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
    v_表格     Varchar2(300);
    n_数量     Number;
    n_Rptcount Number;
    v_病历名称 电子病历记录.病历名称%Type;
    v_挂号单id 病人挂号记录.Id%Type;
  
    Function Getrptno
    (
      v_医嘱idin   病人医嘱发送.医嘱id%Type,
      v_病历名称in 电子病历记录.病历名称%Type
    ) Return Varchar As
      v_Return Number;
      v_No     Number;
      v_Count  Number;
    Begin
      Select Count(医嘱id) + 1 Into v_No From 病人医嘱报告 Where 医嘱id = v_医嘱idin;
      v_Count := 1;
      While v_Count = 1 Loop
        Select Count(ID)
        Into v_Count
        From 病人医嘱报告 A, 电子病历记录 B
        Where a.医嘱id = v_医嘱idin And a.病历id = b.Id And b.病历名称 = v_病历名称in || v_No;
        If v_Count = 1 Then
          v_No := v_No + 1;
        End If;
      End Loop;
      v_Return := v_No;
      Return v_Return;
    End Getrptno;
  
  Begin
  
    -- 提取主医嘱ID ，防止因为传入部位医嘱，导致报告保存出错
    Select Nvl(相关id, ID) As 组id Into v_主医嘱id From 病人医嘱记录 Where ID = 医嘱id_In;
  
    Open c_Advice(v_主医嘱id, Nvl(Risid_In, 0));
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
            Update 电子病历内容 Set 内容文本 = 报告所见_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.内容文本 Like '%意见%' Then
            Update 电子病历内容 Set 内容文本 = 报告意见_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.内容文本 Like '%建议%' Then
            Update 电子病历内容 Set 内容文本 = 报告建议_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          End If;
        End Loop;
        --更新保存时间
        Update 电子病历记录
        Set 完成时间 = Sysdate, 保存人 = 报告医生_In, 保存时间 = Sysdate
        Where ID = r_Advice.病历id;
      Else
        --先判断单据中是否有对应的提纲和表格
        If Nvl(报告所见_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_数量
          From 病历文件结构 A, 病历文件结构 B
          Where a.父id = b.Id And a.对象类型 = 3 And b.对象类型 = 1 And a.内容文本 Like '%所见%' And a.文件id = r_Advice.文件id;
        
          If n_数量 <= 0 Then
            v_Error := '在诊疗单据中没有找到【所见】对应的提纲或表格，请联系管理员设置！';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(报告意见_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_数量
          From 病历文件结构 A, 病历文件结构 B
          Where a.父id = b.Id And a.对象类型 = 3 And b.对象类型 = 1 And a.内容文本 Like '%意见%' And a.文件id = r_Advice.文件id;
        
          If n_数量 <= 0 Then
            v_Error := '在诊疗单据中没有找到【意见】对应的提纲或表格，请联系管理员设置！';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(报告建议_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_数量
          From 病历文件结构 A, 病历文件结构 B
          Where a.父id = b.Id And a.对象类型 = 3 And b.对象类型 = 1 And a.内容文本 Like '%建议%' And a.文件id = r_Advice.文件id;
        
          If n_数量 <= 0 Then
            v_Error := '在诊疗单据中没有找到【建议】对应的提纲或表格，请联系管理员设置！';
            Raise Err_Custom;
          End If;
        End If;
      
        If r_Advice.病人来源 = 1 Then
          --门诊，提取挂号单ID
          Select Nvl(c.Id, 0)
          Into v_挂号单id
          From 病人医嘱记录 B, 病人挂号记录 C
          Where b.挂号单 = c.No(+) And c.记录状态 In (1, 3) And b.Id = v_主医嘱id;
        Else
          --体检或者外诊，无挂号单ID，直接设置为0
          v_挂号单id := 0;
        End If;
      
        --产生电子病历记录
        Select 电子病历记录_Id.Nextval Into v_病历id From Dual;
        n_Rptcount := Getrptno(医嘱id_In, r_Advice.病历名称);
        If n_Rptcount > 1 Then
          v_病历名称 := r_Advice.病历名称 || n_Rptcount;
        Else
          v_病历名称 := r_Advice.病历名称;
        End If;
        Insert Into 电子病历记录
          (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 完成时间, 保存人, 保存时间, 最后版本, 签名级别)
        Values
          (v_病历id, r_Advice.病人来源, r_Advice.病人id, Decode(r_Advice.病人来源, 2, r_Advice.主页id, v_挂号单id), r_Advice.婴儿,
           r_Advice.病人科室id, r_Advice.病历种类, r_Advice.文件id, v_病历名称, 报告医生_In, Sysdate, Sysdate, 报告医生_In, Sysdate, 1, 2);
      
        --产生医嘱报告记录
        Insert Into 病人医嘱报告 (医嘱id, 病历id, Risid) Values (v_主医嘱id, v_病历id, Risid_In);
      
        v_对象序号 := 0;
      
        --新产生报告内容
        For r_File In c_File(r_Advice.文件id) Loop
          Select 电子病历内容_Id.Nextval Into v_病历内容id From Dual;
          v_内容文本   := r_File.内容文本;
          v_定义提纲id := 0;
        
          If Nvl(r_File.对象类型, 0) = 1 And Nvl(r_File.父id, 0) = 0 Then
            --提纲
            v_定义提纲id := r_File.Id;
            v_父id       := v_病历内容id;
          End If;
        
          If Nvl(r_File.对象类型, 0) = 4 And r_File.要素名称 Is Not Null Then
            --元素
            v_内容文本 := Zl_Replace_Element_Value(r_File.要素名称, r_Advice.病人id, r_Advice.主页id, r_Advice.病人来源, r_Advice.Id);
          End If;
        
          If Nvl(r_File.父id, 0) <> 0 Then
            v_定义提纲id := 0;
          End If;
        
          v_对象序号 := v_对象序号 + 1;
        
          If Instr(v_表格, '|' || r_File.父id || '|') > 0 Then
            Null;
          Else
            Insert Into 电子病历内容
              (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域,
               要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
            Values
              (v_病历内容id, v_病历id, 1, 0, Decode(v_定义提纲id, 0, v_父id, Null), v_对象序号, r_File.对象类型, r_File.对象标记, r_File.保留对象,
               r_File.对象属性, Null, v_内容文本, r_File.是否换行, r_File.预制提纲id, r_File.复用提纲, r_File.使用时机, r_File.诊治要素id,
               r_File.替换域, r_File.要素名称, r_File.要素类型, r_File.要素长度, r_File.要素小数, r_File.要素单位, r_File.要素表示, r_File.输入形态,
               r_File.要素值域, Decode(v_定义提纲id, 0, Null, v_定义提纲id));
          End If;
        
          --为表格时，插入文本内容
          If Nvl(r_File.对象类型, 0) = 3 And Nvl(r_File.父id, 0) <> 0 Then
            v_表格 := v_表格 || ',|' || r_File.Id || '|';
          
            If r_File.内容文本 Like '%所见%' Then
              v_内容文本 := 报告所见_In || Chr(13) || Chr(13);
            Elsif r_File.内容文本 Like '%意见%' Then
              v_内容文本 := 报告意见_In || Chr(13) || Chr(13);
            Else
              v_内容文本 := 报告建议_In || Chr(13) || Chr(13);
            End If;
          
            For r_Con In c_Content(r_Advice.文件id, r_File.Id) Loop
              Select 电子病历内容_Id.Nextval Into v_病历内容idnew From Dual;
              v_对象序号 := v_对象序号 + 1;
            
              Insert Into 电子病历内容
                (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id,
                 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
              Values
                (v_病历内容idnew, v_病历id, 1, 0, v_病历内容id, v_对象序号, 2, r_Con.对象标记, r_Con.保留对象, r_Con.对象属性, Null, v_内容文本,
                 r_Con.是否换行, r_Con.预制提纲id, r_Con.复用提纲, r_Con.使用时机, r_Con.诊治要素id, r_Con.替换域, r_Con.要素名称, r_Con.要素类型,
                 r_Con.要素长度, r_Con.要素小数, r_Con.要素单位, r_Con.要素表示, r_Con.输入形态, r_Con.要素值域,
                 Decode(v_定义提纲id, 0, Null, v_定义提纲id));
            End Loop;
          End If;
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

  --7、插入医嘱操作失败记录
  Procedure Ris医嘱失败记录_Insert
  (
    病人来源_In   In Ris医嘱失败记录.病人来源%Type,
    病人id_In     In Ris医嘱失败记录.病人id%Type,
    主页id_In     In Ris医嘱失败记录.主页id%Type,
    挂号单号_In   In Ris医嘱失败记录.挂号单号%Type,
    发送号_In     In Ris医嘱失败记录.发送号%Type,
    体检任务id_In In Ris医嘱失败记录.体检任务id%Type,
    体检报到号_In In Ris医嘱失败记录.体检报到号%Type,
    发送类型_In   In Ris医嘱失败记录.发送类型%Type
  ) Is
  Begin
    Insert Into Ris医嘱失败记录
      (ID, 病人来源, 病人id, 主页id, 挂号单号, 发送号, 体检任务id, 体检报到号, 发送类型, 发送时间, 重发次数)
    Values
      (Ris医嘱失败记录_Id.Nextval, 病人来源_In, 病人id_In, 主页id_In, 挂号单号_In, 发送号_In, 体检任务id_In, 体检报到号_In, 发送类型_In, Sysdate, 0);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris医嘱失败记录_Insert;

  --8、更新医嘱操作失败记录
  Procedure Ris医嘱失败记录_重发
  (
    Id_In       In Ris医嘱失败记录.Id%Type,
    操作类型_In In Number
  ) Is
    v_重发次数 Ris医嘱失败记录.重发次数%Type;
  Begin
    --操作类型_In -- 1 重发成功，删除记录；2--重发失败
  
    If 操作类型_In = 1 Then
      Delete From Ris医嘱失败记录 Where ID = Id_In;
    Else
      Select 重发次数 Into v_重发次数 From Ris医嘱失败记录 Where ID = Id_In;
      If v_重发次数 >= 99 Then
        v_重发次数 := 99;
      Else
        v_重发次数 := v_重发次数 + 1;
      End If;
      Update Ris医嘱失败记录 Set 发送时间 = Sysdate, 重发次数 = v_重发次数 Where ID = Id_In;
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris医嘱失败记录_重发;

  --9、销账后新建住院记账单据
  Procedure 病人医嘱_重建单据
  (
    医嘱id_In In 病人医嘱发送.医嘱id%Type,
    No_In     In 病人医嘱发送.No%Type,
    Action_In In Number
  ) Is
    -- Action_In: 1 重建单据；2 取消重建单据
    v_No 病人医嘱发送.No%Type;
  Begin
    If Action_In = 1 Then
      Select Nextno(14) Into v_No From Dual;
    
      Update 病人医嘱发送
      Set NO = v_No, 计费状态 = 0
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In);
      Update 住院费用记录 Set 医嘱序号 = Null Where NO = No_In;
    Elsif Action_In = 2 Then
      Update 住院费用记录 Set 医嘱序号 = 医嘱id_In Where NO = No_In;
      Update 病人医嘱发送
      Set NO = No_In, 计费状态 = 4
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 病人医嘱_重建单据;

  --10、打印RIS检查预约通知单
  Procedure Ris检查预约_打印(医嘱id_In In Ris检查预约.医嘱id%Type) Is
    v_Temp     Varchar2(255);
    v_人员姓名 人员表.姓名%Type;
  Begin
    --取当前操作人员
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  
    Update Ris检查预约 Set 是否打印 = 1, 打印人 = v_人员姓名, 打印时间 = Sysdate Where 医嘱id = 医嘱id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris检查预约_打印;

  --11、更新RIS分科室启用参数
  Procedure Ris启用控制_Update
  (
    检查类型_In Ris启用控制.检查类型%Type,
    场合_In     Ris启用控制.场合%Type,
    部门ids_In  Varchar2,
    启用类型_In Number
  ) Is
  
    l_部门id   t_Numlist := t_Numlist();
    v_启用ris  Ris启用控制.是否启用ris%Type;
    v_启用预约 Ris启用控制.是否启用预约%Type;
  
    Cursor c_Dept(Dept_In Varchar2) Is
      Select Column_Value From Table(f_Num2list(Dept_In));
  Begin
  
    If 启用类型_In = 1 Then
      v_启用ris  := 1;
      v_启用预约 := Null;
      Delete From Ris启用控制 Where 检查类型 = 检查类型_In And 场合 = 场合_In And 是否启用ris = 1;
    Else
      v_启用ris  := Null;
      v_启用预约 := 1;
      Delete From Ris启用控制 Where 检查类型 = 检查类型_In And 场合 = 场合_In And 是否启用预约 = 1;
    End If;
  
    If 部门ids_In Is Null Then
      Insert Into Ris启用控制
        (ID, 检查类型, 场合, 部门id, 是否启用ris, 是否启用预约)
      Values
        (Ris启用控制_Id.Nextval, 检查类型_In, 场合_In, Null, v_启用ris, v_启用预约);
    Else
      Open c_Dept(部门ids_In);
      Fetch c_Dept Bulk Collect
        Into l_部门id;
      Close c_Dept;
    
      Forall I In 1 .. l_部门id.Count
        Insert Into Ris启用控制
          (ID, 检查类型, 场合, 部门id, 是否启用ris, 是否启用预约)
        Values
          (Ris启用控制_Id.Nextval, 检查类型_In, 场合_In, l_部门id(I), v_启用ris, v_启用预约);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris启用控制_Update;

  --12、删除RIS分科室启用参数
  Procedure Ris启用控制_Delete Is
  
  Begin
    Delete From Ris启用控制;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris启用控制_Delete;

  --13、根据元素名提取信息
  Function Ris_Replace_Element_Value
  (
    元素名_In   In 诊治所见项目.中文名%Type,
    病人id_In   In 电子病历记录.病人id%Type,
    就诊id_In   In 电子病历记录.主页id%Type,
    病人来源_In In 电子病历记录.病人来源%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type
  ) Return Varchar2 Is
    v_Return Varchar2(4000) := Null;
    Cursor c_Patient Is
      Select 姓名, 性别, Decode(性别, '男', 'M', '女', 'F', 'O') As 性别编码, 出生日期, 病人id, 联系人地址, 家庭电话, 联系人电话, 婚姻状况, 身份证号, 当前科室id,
             当前病区id, 当前床号 As 床号, 就诊卡号, 入院时间, 出院时间
      From 病人信息
      Where 病人id = 病人id_In;
    r_Patient c_Patient%RowType;
  
    Cursor c_Order Is
      Select 主页id, 婴儿, Decode(病人来源, 1, 'OUTPAT', 2, 'INPAT', 'UNK') As 病人来源, 开嘱医生, 开嘱时间, 校对护士, 医嘱内容, 紧急标志, 执行科室id
      From 病人医嘱记录
      Where ID = 医嘱id_In;
    r_Order c_Order%RowType;
  
    Cursor c_Diagnose Is
      Select 诊断描述 || Decode(Nvl(是否疑诊, 0), 0, '', ' (？)') As 临床诊断
      From 病人诊断医嘱 A, 病人诊断记录 B
      Where a.医嘱id = 医嘱id_In And a.诊断id = b.Id;
    r_Diagnose c_Diagnose%RowType;
  
    --获取指定表的行类型
    Procedure p_Get_Rowtype(Table_In In Varchar2) Is
    Begin
      If Table_In = '病人信息' Then
        Open c_Patient;
        Fetch c_Patient
          Into r_Patient;
      Elsif Table_In = '病人医嘱记录' Then
        Open c_Order;
        Fetch c_Order
          Into r_Order;
      Elsif Table_In = '病人诊断记录' Then
        Open c_Diagnose;
        Fetch c_Diagnose
          Into r_Diagnose;
      End If;
    Exception
      When Others Then
        Null;
    End p_Get_Rowtype;
  
  Begin
    Case
    --直接返回的输入元素
      When 元素名_In = '医嘱ID' Then
        v_Return := 医嘱id_In;
      When 元素名_In = '病人ID' Then
        v_Return := 病人id_In;
      
    --姓名，性别单独处理，可能是婴儿
      When Instr(',姓名,性别,性别编码,出生日期,', ',' || 元素名_In || ',') > 0 Then
        p_Get_Rowtype('病人医嘱记录');
        p_Get_Rowtype('病人信息');
        If Nvl(r_Order.婴儿, 0) = 0 Then
          If 元素名_In = '姓名' Then
            v_Return := r_Patient.姓名;
          Elsif 元素名_In = '性别' Then
            v_Return := r_Patient.性别;
          Elsif 元素名_In = '性别编码' Then
            v_Return := r_Patient.性别编码;
          Elsif 元素名_In = '出生日期' Then
            v_Return := To_Char(r_Patient.出生日期, 'YYYYMMDDMISS');
          End If;
        Else
          If 元素名_In = '姓名' Then
            Select Decode(婴儿姓名, Null, r_Patient.姓名 || '之婴' || Trim(To_Char(序号, '9')), 婴儿姓名) As 婴儿姓名
            Into v_Return
            From 病人新生儿记录
            Where 病人id = 病人id_In And 主页id = r_Order.主页id And 序号 = Nvl(r_Order.婴儿, 0);
          Elsif Instr('性别', 元素名_In) > 0 Then
            Select 婴儿性别
            Into v_Return
            From 病人新生儿记录
            Where 病人id = 病人id_In And 主页id = r_Order.主页id And 序号 = Nvl(r_Order.婴儿, 0);
            If 元素名_In = '性别编码' Then
              Select Decode(v_Return, '男', 'M', '女', 'F', 'O') Into v_Return From Dual;
            End If;
          Elsif 元素名_In = '出生日期' Then
            Select 出生时间
            Into v_Return
            From 病人新生儿记录
            Where 病人id = 病人id_In And 主页id = r_Order.主页id And 序号 = Nvl(r_Order.婴儿, 0);
            v_Return := To_Char(v_Return, 'YYYYMMDDMISS');
          End If;
        End If;
      
    --查询病人信息表返回的元素
      When Instr(',联系人地址,家庭电话,联系人电话,婚姻状况,身份证号,床号,就诊卡号,入院时间,出院时间,', ',' || 元素名_In || ',') > 0 Then
        p_Get_Rowtype('病人信息');
        Case 元素名_In
          When '联系人地址' Then
            v_Return := r_Patient.联系人地址;
          When '家庭电话' Then
            v_Return := r_Patient.家庭电话;
          When '联系人电话' Then
            v_Return := r_Patient.联系人电话;
          When '婚姻状况' Then
            v_Return := r_Patient.婚姻状况;
          When '身份证号' Then
            v_Return := r_Patient.身份证号;
          When '床号' Then
            v_Return := r_Patient.床号;
          When '就诊卡号' Then
            v_Return := r_Patient.就诊卡号;
          When '入院时间' Then
            v_Return := To_Char(r_Patient.入院时间, 'YYYYMMDDMISS');
          When '出院时间' Then
            v_Return := To_Char(r_Patient.出院时间, 'YYYYMMDDMISS');
          Else
            v_Return := '';
        End Case;
        --查询医嘱表返回的元素
      When Instr(',病人来源,开嘱医生,开嘱时间,校对护士,医嘱内容,紧急标志,紧急标志对码,', ',' || 元素名_In || ',') > 0 Then
        p_Get_Rowtype('病人医嘱记录');
        Case 元素名_In
          When '病人来源' Then
            v_Return := r_Order.病人来源;
          When '开嘱医生' Then
            v_Return := r_Order.开嘱医生;
          When '开嘱时间' Then
            v_Return := To_Char(r_Order.开嘱时间, 'YYYYMMDDMISS');
          When '校对护士' Then
            v_Return := r_Order.校对护士;
          When '医嘱内容' Then
            v_Return := r_Order.医嘱内容;
          When '紧急标志' Then
            v_Return := r_Order.紧急标志;
        End Case;
        --查询诊断记录返回的元素
      When 元素名_In = '临床诊断' Then
        p_Get_Rowtype('病人诊断记录');
        v_Return := r_Diagnose.临床诊断;
      
      Else
        --自行查询SQL返回值的元素
        If 元素名_In = '执行站点' Then
          p_Get_Rowtype('病人医嘱记录');
          Select Decode(站点, 1, 'SITE0002', 2, 'SITE0001', 3, 'SITE0003', 'SITE0001')
          Into v_Return
          From 部门表
          Where ID = r_Order.执行科室id;
        End If;
        If 元素名_In = '当前科室名称' Then
          p_Get_Rowtype('病人信息');
          Select 名称 Into v_Return From 部门表 Where ID = r_Patient.当前科室id;
        End If;
        If 元素名_In = '病区名称' Then
          p_Get_Rowtype('病人信息');
          Select 名称 Into v_Return From 部门表 Where ID = r_Patient.当前病区id;
        End If;
        If 元素名_In = '标识号' Then
          Select Decode(a.病人来源, 1, c.门诊号, 2, Decode(c.住院号, Null, c.门诊号, c.住院号), 4, c.健康号, c.门诊号)
          Into v_Return
          From 病人医嘱记录 A, 病人信息 C
          Where a.病人id = c.病人id And a.Id = 医嘱id_In;
        End If;
    End Case;
  
    Return Trim(v_Return);
  Exception
    When Others Then
      Return Null;
  End Ris_Replace_Element_Value;

  --14、删除RIS分院设置参数
  Procedure Ris分院设置_Delete Is
  Begin
    Delete From Ris分院设置;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris分院设置_Delete;

  --15、更新RISRis分院设置参数
  Procedure Ris分院设置_Update
  (
    Id_In           Ris分院设置.Id%Type,
    医院名称_In     Ris分院设置.医院名称%Type,
    医院代码_In     Ris分院设置.医院代码%Type,
    用户名_In       Ris分院设置.用户名%Type,
    密码_In         Ris分院设置.密码%Type,
    数据库服务名_In Ris分院设置.数据库服务名%Type
  ) Is
  
  Begin
  
    Insert Into Ris分院设置
      (ID, 医院名称, 医院代码, 用户名, 密码, 数据库服务名)
    Values
      (Id_In, 医院名称_In, 医院代码_In, 用户名_In, 密码_In, 数据库服务名_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris分院设置_Update;

End b_Zlxwinterface;
/
--128157:蒋敏,2018-07-04,诊疗项目部位删除数据不正确
--126017:蒋敏,2018-05-23,诊疗项目管理检查部位方法的选择处理
Create Or Replace Procedure Zl_诊疗检查部位_Edit
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
  上级方法_In In 诊疗检查部位.方法%Type := Null --格式：上级方法|方法;上级方法|方法...(若上级方法为空，则为|方法)    
) Is
  v_原名称 诊疗检查部位.名称%Type := Null;
  e_Notfind Exception;
  v_方法   Varchar2(1000);
  v_Fields Varchar2(1000);
  v_Tmp    Varchar2(1000);
  n_Count  Number;
  n_记录id 诊疗项目部位.Id%Type;
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
    v_方法 := ';' || 方法_In;
    v_方法 := Replace(v_方法, ',', Chr(10));
    v_方法 := Replace(v_方法, Chr(9), ';');
    v_方法 := Replace(v_方法, ';0', Chr(10));
    v_方法 := Replace(v_方法, ';1', Chr(10));
    v_方法 := Replace(v_方法, Chr(10), ';');
    v_方法 := Replace(v_方法, ';;', ';');
    v_方法 := v_方法 || ';';
  
    v_方法 := Substr(v_方法, 2);
    --原有的方法，现在已经删除了或原有的部位的名称已经改变了
    For r_Used In (Select ID, 项目id, 部位, 方法, 类型, 默认,上级方法 From 诊疗项目部位 Where 部位 = v_原名称 And 类型 = 类型_In) Loop
      If Instr(';' || v_方法, ';' || r_Used.方法 || ';') = 0 Then
        Delete 诊疗项目部位
        Where ID=r_Used.id;
      Else
        Update 诊疗项目部位
        Set 部位 = 名称_In
        Where ID=r_Used.id;
      End If;
    End Loop;
  
    --原来没有的方法现在新增
    v_Tmp := v_方法;
    While v_Tmp Is Not Null Loop
      --依次取每个项目
      v_Fields := Substr(v_Tmp, 1, Instr(v_Tmp, ';') - 1);
      v_Tmp    := Substr(v_Tmp, Instr(v_Tmp, ';') + 1);
    
      If v_Fields Is Not Null Then
        For r_Used In (Select Distinct 项目id From 诊疗项目部位 Where 部位 = 名称_In And 类型 = 类型_In) Loop
          Select Count(ID)
          Into n_Count
          From 诊疗项目部位
          Where 项目id = r_Used.项目id And 部位 = 名称_In And 类型 = 类型_In And 方法 = v_Fields;
        
          If n_Count = 0 Then
            Select 诊疗项目部位_Id.Nextval Into n_记录id From Dual;
            Insert Into 诊疗项目部位
              (ID, 项目id, 类型, 部位, 方法)
            Values
              (n_记录id, r_Used.项目id, 类型_In, 名称_In, v_Fields);
          End If;
        End Loop;
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
--126017:蒋敏,2018-05-23,诊疗项目管理检查部位方法的选择处理
CREATE OR REPLACE Procedure Zl_诊疗项目部位_Insert
(
  项目id_In In 诊疗项目部位.项目id%Type,
  类型_In   In 诊疗项目部位.类型%Type,
  部位_In   In 诊疗项目部位.部位%Type,
  方法_In   In 诊疗项目部位.方法%Type,
  默认_In   In 诊疗项目部位.默认%Type := Null,
  上级方法_In  In 诊疗项目部位.上级方法%Type := Null
) As
  v_Code Varchar2(20); --编码
  Err_Notfind Exception;
Begin
  Select Rtrim(编码) Into v_Code From 诊疗项目目录 Where 类别 = 'D' And Id = 项目id_In;
  If v_Code Is Null Then
    Raise Err_Notfind;
  End If;
  Insert Into 诊疗项目部位 (ID, 项目id, 类型, 部位, 方法, 默认,上级方法) Values (诊疗项目部位_ID.Nextval, 项目id_In, 类型_In, 部位_In, 方法_In, 默认_In,上级方法_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该项目不存在，可能已被其他用户删除！[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_诊疗项目部位_Insert;
/
--119268:李小东,2018-05-11,去掉过程中删除相关代码
Create Or Replace Procedure Zl_检验图像结果_Update
(
  标本id_In   In 检验图像结果.标本id%Type,
  图像类型_In In 检验图像结果.图像类型%Type,
  图像点_In   In Varchar2, -- 图像名称^图像类型;结果1;结果2;结果3....   图像类型 1=直方图 2=散点图 
  图像lob_In  In Number, -- 0=小于4000  1=大于4000 需要特殊处理 
  开始_In     In Number, -- 1=开始 
  图像位置_In In 检验图像结果.图像位置%Type := Null
) Is
  l_Clob Clob;
Begin
  -- 保存到FTP 
  If 图像位置_In Is Not Null Then
    Update 检验图像结果 Set 图像位置 = 图像位置_In Where 标本id = 标本id_In And 图像类型 = 图像类型_In;
    If Sql%RowCount = 0 Then
      Insert Into 检验图像结果
        (ID, 标本id, 图像类型, 图像位置)
      Values
        (检验图像结果_Id.Nextval, 标本id_In, 图像类型_In, 图像位置_In);
    End If;
    Return;
  End If;

  -- 保存到数据库 
  If 图像点_In Is Null Then
    Return;
  End If;

  If 图像lob_In = 0 Then
    Update 检验图像结果 Set 图像点 = 图像点_In Where 标本id = 标本id_In And 图像类型 = 图像类型_In;
    If Sql%RowCount = 0 Then
      Insert Into 检验图像结果
        (ID, 标本id, 图像类型, 图像点)
      Values
        (检验图像结果_Id.Nextval, 标本id_In, 图像类型_In, 图像点_In);
    End If;
  Else
    If 开始_In = 1 Then
      Update 检验图像结果 Set 图像点 = Empty_Clob() Where 标本id = 标本id_In And 图像类型 = 图像类型_In;
      If Sql%RowCount = 0 Then
        Insert Into 检验图像结果
          (ID, 标本id, 图像类型, 图像点)
        Values
          (检验图像结果_Id.Nextval, 标本id_In, 图像类型_In, Empty_Clob());
      End If;
    End If;
    Select 图像点 Into l_Clob From 检验图像结果 Where 标本id = 标本id_In And 图像类型 = 图像类型_In For Update;
    Dbms_Lob.Writeappend(l_Clob, Length(图像点_In), 图像点_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验图像结果_Update;
/

--119268:李小东,2018-05-11,去掉过程中删除相关代码
Create Or Replace Procedure Zl_电子病历格式_Insert
(
  Id_In   In 电子病历格式.文件id%Type,
  Txt_In  In Varchar2,
  开始_In In Number -- 1=开始 
) Is
  l_Blob Blob;
Begin
  If 开始_In = 1 Then
    Update 电子病历格式 Set 内容 = Empty_Blob() Where 文件id = Id_In;
    If Sql%RowCount = 0 Then
      Insert Into 电子病历格式 (文件id, 内容) Values (Id_In, Empty_Blob());
    End If;
  End If;
  Select 内容 Into l_Blob From 电子病历格式 Where 文件id = Id_In For Update;
  Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子病历格式_Insert;
/

--125261:胡俊勇,2018-05-08,转科医嘱校对发送处理自动停长嘱
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
          Nvl(a.医嘱期效, 0) = 0 And a.医嘱状态 Not In (1, 2, 4, 8, 9) And a.开始执行时间 <= v_Stoptime
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
--111037:余伟节,2018-06-24,新生儿登记允许录入死亡时间
--124866:胡俊勇,2018-06-06,病重病危医嘱变动
--125261:胡俊勇,2018-05-08,转科医嘱校对发送处理自动停长嘱
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
          Nvl(a.医嘱期效, 0) = 0 And a.医嘱状态 Not In (1, 2, 4, 8, 9) And a.开始执行时间 <= v_Stoptime
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
          If r_Advice.开始执行时间 < v_Date Then
            v_Error := '医嘱"' || r_Advice.医嘱内容 || '"的开始时间应大于该病人上次变动时间 ' || To_Char(v_Date, 'YYYY-MM-DD HH24:Mi') || ' 。';
            Raise Err_Custom;
          End If;
          Zl_病人变动记录_Preout(r_Advice.病人id, r_Advice.主页id, r_Advice.开始执行时间);
        End If;
      Else
        If r_Advice.操作类型 = '11' Then
          Update 病人新生儿记录
          Set 死亡时间 = r_Advice.开始执行时间
          Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = Nvl(r_Advice.主页id, 0) And Nvl(序号, 0) = Nvl(r_Advice.婴儿, 0);
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

--124269:胡俊勇,2018-05-07,基于会诊申请下达医嘱标记
Create Or Replace Procedure Zl_病人医嘱记录_Insert
(
  Id_In           In 病人医嘱记录.Id%Type,
  相关id_In       In 病人医嘱记录.相关id%Type,
  序号_In         In 病人医嘱记录.序号%Type,
  病人来源_In     In 病人医嘱记录.病人来源%Type,
  病人id_In       In 病人医嘱记录.病人id%Type,
  主页id_In       In 病人医嘱记录.主页id%Type,
  婴儿_In         In 病人医嘱记录.婴儿%Type,
  医嘱状态_In     In 病人医嘱记录.医嘱状态%Type,
  医嘱期效_In     In 病人医嘱记录.医嘱期效%Type,
  诊疗类别_In     In 病人医嘱记录.诊疗类别%Type,
  诊疗项目id_In   In 病人医嘱记录.诊疗项目id%Type,
  收费细目id_In   In 病人医嘱记录.收费细目id%Type,
  天数_In         In 病人医嘱记录.天数%Type,
  单次用量_In     In 病人医嘱记录.单次用量%Type,
  总给予量_In     In 病人医嘱记录.总给予量%Type,
  医嘱内容_In     In 病人医嘱记录.医嘱内容%Type,
  医生嘱托_In     In 病人医嘱记录.医生嘱托%Type,
  标本部位_In     In 病人医嘱记录.标本部位%Type,
  执行频次_In     In 病人医嘱记录.执行频次%Type,
  频率次数_In     In 病人医嘱记录.频率次数%Type,
  频率间隔_In     In 病人医嘱记录.频率间隔%Type,
  间隔单位_In     In 病人医嘱记录.间隔单位%Type,
  执行时间方案_In In 病人医嘱记录.执行时间方案%Type,
  计价特性_In     In 病人医嘱记录.计价特性%Type,
  执行科室id_In   In 病人医嘱记录.执行科室id%Type,
  执行性质_In     In 病人医嘱记录.执行性质%Type,
  紧急标志_In     In 病人医嘱记录.紧急标志%Type,
  开始执行时间_In In 病人医嘱记录.开始执行时间%Type,
  执行终止时间_In In 病人医嘱记录.执行终止时间%Type,
  病人科室id_In   In 病人医嘱记录.病人科室id%Type,
  开嘱科室id_In   In 病人医嘱记录.开嘱科室id%Type,
  开嘱医生_In     In 病人医嘱记录.开嘱医生%Type,
  开嘱时间_In     In 病人医嘱记录.开嘱时间%Type,
  挂号单_In       In 病人医嘱记录.挂号单%Type := Null,
  前提id_In       In 病人医嘱记录.前提id%Type := Null,
  检查方法_In     In 病人医嘱记录.检查方法%Type := Null,
  执行标记_In     In 病人医嘱记录.执行标记%Type := Null,
  可否分零_In     In 病人医嘱记录.可否分零%Type := Null,
  摘要_In         In 病人医嘱记录.摘要%Type := Null,
  操作员姓名_In   In 病人医嘱状态.操作人员%Type := Null,
  零费记帐_In     In 病人医嘱记录.零费记帐%Type := Null,
  用药目的_In     In 病人医嘱记录.用药目的%Type := Null,
  用药理由_In     In 病人医嘱记录.用药理由%Type := Null,
  审核状态_In     In 病人医嘱记录.审核状态%Type := Null,
  申请序号_In     In 病人医嘱记录.申请序号%Type := Null,
  超量说明_In     In 病人医嘱记录.超量说明%Type := Null,
  首次用量_In     In 病人医嘱记录.首次用量%Type := Null,
  配方id_In       In 病人医嘱记录.配方id%Type := Null,
  手术情况_In     In 病人医嘱记录.手术情况%Type := Null,
  组合项目id_In   In 病人医嘱记录.组合项目id%Type := Null,
  皮试结果_In     In 病人医嘱记录.皮试结果%Type := Null,
  处方序号_In     In 病人医嘱记录.处方序号%Type := Null,
  会诊医嘱id_In   In 病人医嘱记录.会诊医嘱id%Type := Null
  --功能：医生或护士新开,补录医嘱时新产生的医嘱记录。可用于门诊或住院。
) Is
  v_Temp     Varchar2(255);
  v_人员姓名 病人医嘱状态.操作人员%Type;

  v_姓名     病人信息.姓名%Type;
  v_性别     病人信息.性别%Type;
  v_年龄     病人信息.年龄%Type;
  d_手术时间 病人医嘱记录.手术时间%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --当前操作人员
  If 操作员姓名_In Is Not Null Then
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  If Nvl(主页id_In, 0) <> 0 Then
    Select 姓名, 性别, 年龄 Into v_姓名, v_性别, v_年龄 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Select 姓名, 性别, 年龄 Into v_姓名, v_性别, v_年龄 From 病人信息 Where 病人id = 病人id_In;
  End If;

  If Instr(',F,K,', 诊疗类别_In) > 0 Then
    d_手术时间 := To_Date(标本部位_In, 'yyyy-mm-dd hh24:mi:ss');
  End If;

  --病人医嘱记录
  Insert Into 病人医嘱记录
    (ID, 相关id, 序号, 病人来源, 病人id, 主页id, 姓名, 性别, 年龄, 婴儿, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id, 收费细目id, 天数, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 标本部位,
     检查方法, 执行标记, 执行频次, 频率次数, 频率间隔, 间隔单位, 执行时间方案, 计价特性, 执行科室id, 执行性质, 紧急标志, 可否分零, 开始执行时间, 执行终止时间, 病人科室id, 开嘱科室id, 开嘱医生,
     开嘱时间, 挂号单, 前提id, 摘要, 零费记帐, 手术时间, 用药目的, 用药理由, 审核状态, 申请序号, 超量说明, 首次用量, 配方id, 手术情况, 组合项目id, 皮试结果, 处方序号, 会诊医嘱id)
  Values
    (Id_In, 相关id_In, 序号_In, 病人来源_In, 病人id_In, 主页id_In, v_姓名, v_性别, v_年龄, 婴儿_In, 医嘱状态_In, 医嘱期效_In, 诊疗类别_In, 诊疗项目id_In,
     收费细目id_In, 天数_In, 单次用量_In, 总给予量_In, 医嘱内容_In, 医生嘱托_In, 标本部位_In, 检查方法_In, 执行标记_In, 执行频次_In, 频率次数_In, 频率间隔_In, 间隔单位_In,
     执行时间方案_In, 计价特性_In, 执行科室id_In, 执行性质_In, 紧急标志_In, 可否分零_In, 开始执行时间_In, 执行终止时间_In, 病人科室id_In, 开嘱科室id_In, 开嘱医生_In,
     开嘱时间_In, 挂号单_In, 前提id_In, 摘要_In, 零费记帐_In, d_手术时间, 用药目的_In, 用药理由_In, 审核状态_In, 申请序号_In, 超量说明_In, 首次用量_In, 配方id_In,
     手术情况_In, 组合项目id_In, 皮试结果_In, 处方序号_In, 会诊医嘱id_In);

  --病人医嘱状态
  If 医嘱状态_In <> -1 Then
    Delete From 病人医嘱状态 Where 医嘱id = Id_In And 操作类型 = 1;
    If Sql%RowCount <> 0 Then
      v_Error := '相同ID的新开医嘱已经存在。';
      Raise Err_Custom;
    End If;
    --因为可能同时：新开->自动校对(住院医生发送)->互斥自动停止(住院医生发送临嘱停止),因此分别-2,-1秒
    Insert Into 病人医嘱状态
      (医嘱id, 操作类型, 操作人员, 操作时间)
    Values
      (Id_In, 1, v_人员姓名, Sysdate - 2 / 60 / 60 / 24);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_Insert;
/

--124269:胡俊勇,2018-05-07,基于会诊申请下达医嘱标记
Create Or Replace Procedure Zl_病人医嘱记录_Update
(
  Id_In           In 病人医嘱记录.Id%Type,
  相关id_In       In 病人医嘱记录.相关id%Type,
  序号_In         In 病人医嘱记录.序号%Type,
  医嘱状态_In     In 病人医嘱记录.医嘱状态%Type,
  医嘱期效_In     In 病人医嘱记录.医嘱期效%Type,
  诊疗项目id_In   In 病人医嘱记录.诊疗项目id%Type,
  收费细目id_In   In 病人医嘱记录.收费细目id%Type,
  天数_In         In 病人医嘱记录.天数%Type,
  单次用量_In     In 病人医嘱记录.单次用量%Type,
  总给予量_In     In 病人医嘱记录.总给予量%Type,
  医嘱内容_In     In 病人医嘱记录.医嘱内容%Type,
  医生嘱托_In     In 病人医嘱记录.医生嘱托%Type,
  标本部位_In     In 病人医嘱记录.标本部位%Type,
  执行频次_In     In 病人医嘱记录.执行频次%Type,
  频率次数_In     In 病人医嘱记录.频率次数%Type,
  频率间隔_In     In 病人医嘱记录.频率间隔%Type,
  间隔单位_In     In 病人医嘱记录.间隔单位%Type,
  执行时间方案_In In 病人医嘱记录.执行时间方案%Type,
  计价特性_In     In 病人医嘱记录.计价特性%Type,
  执行科室id_In   In 病人医嘱记录.执行科室id%Type,
  执行性质_In     In 病人医嘱记录.执行性质%Type,
  紧急标志_In     In 病人医嘱记录.紧急标志%Type,
  开始执行时间_In In 病人医嘱记录.开始执行时间%Type,
  执行终止时间_In In 病人医嘱记录.执行终止时间%Type,
  病人科室id_In   In 病人医嘱记录.病人科室id%Type,
  开嘱科室id_In   In 病人医嘱记录.开嘱科室id%Type,
  开嘱医生_In     In 病人医嘱记录.开嘱医生%Type,
  开嘱时间_In     In 病人医嘱记录.开嘱时间%Type,
  检查方法_In     In 病人医嘱记录.检查方法%Type := Null,
  执行标记_In     In 病人医嘱记录.执行标记%Type := Null,
  可否分零_In     In 病人医嘱记录.可否分零%Type := Null,
  摘要_In         In 病人医嘱记录.摘要%Type := Null,
  操员作姓名_In   In 病人医嘱状态.操作人员%Type := Null,
  零费记帐_In     In 病人医嘱记录.零费记帐%Type := Null,
  用药目的_In     In 病人医嘱记录.用药目的%Type := Null,
  用药理由_In     In 病人医嘱记录.用药理由%Type := Null,
  审核状态_In     In 病人医嘱记录.审核状态%Type := Null,
  超量说明_In     In 病人医嘱记录.超量说明%Type := Null,
  首次用量_In     In 病人医嘱记录.首次用量%Type := Null,
  手术情况_In     In 病人医嘱记录.手术情况%Type := Null,
  组合项目id_In   In 病人医嘱记录.组合项目id%Type := Null,
  皮试结果_In     In 病人医嘱记录.皮试结果%Type := Null,
  处方序号_In     In 病人医嘱记录.处方序号%Type := Null,
  会诊医嘱id_In   In 病人医嘱记录.会诊医嘱id%Type := Null
  --功能：被医生或护士修改了部分内容的医嘱记录。可用于门诊或住院。
  --说明：Update时之所以涉及诊疗项目ID,计价特性变化,是因为给药途径,用法的变化
  --      Update时之所以涉及期效变化,是因为自由录入医嘱可任意改变期效
) Is
  v_Count Number;

  v_Temp            Varchar2(255);
  v_人员姓名        病人医嘱状态.操作人员%Type;
  v_处方审查锁定ids Varchar2(4000);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --检查该医嘱状态:并发操作
  Begin
    Select 医嘱状态 Into v_Count From 病人医嘱记录 Where ID = Id_In;
  Exception
    When Others Then
      Begin
        v_Error := '医嘱"' || 医嘱内容_In || '"已经不存在,可能已被其他人删除。';
        Raise Err_Custom;
      End;
  End;
  If v_Count Not In (-1, 1, 2) Then
    v_Error := '医嘱"' || 医嘱内容_In || '"已经校对或发送,不能再修改。';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From 病人医嘱状态 Where 医嘱id = Id_In And 操作类型 = 1 And 签名id Is Not Null;
  If Nvl(v_Count, 0) > 0 Then
    v_Error := '医嘱"' || 医嘱内容_In || '"已经电子签名,不能再修改。';
    Raise Err_Custom;
  End If;

  --处方审查撤销
  If 相关id_In Is Null Then
    Zl_处方审查_Cancel(Id_In, v_处方审查锁定ids);
  End If;

  If v_处方审查锁定ids Is Not Null Then
    v_Error := '医嘱"' || 医嘱内容_In || '"已锁定，正在进行处方审查，不能再修改。';
    Raise Err_Custom;
  End If;

  --当前操作人员
  If 操员作姓名_In Is Not Null Then
    v_人员姓名 := 操员作姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  --病人医嘱记录
  Update 病人医嘱记录
  Set 相关id = 相关id_In,
      --比如一并给药，重新设置检查部位等引起的相关ID变化
      序号 = 序号_In, 医嘱状态 = 医嘱状态_In,
      --!因为只能修改未校对医嘱，所以应该为新开，校对疑问的医嘱修改后为新开
      医嘱期效 = 医嘱期效_In, 诊疗项目id = 诊疗项目id_In, 收费细目id = 收费细目id_In, 天数 = 天数_In, 单次用量 = 单次用量_In, 总给予量 = 总给予量_In, 医嘱内容 = 医嘱内容_In,
      医生嘱托 = 医生嘱托_In, 标本部位 = 标本部位_In, 检查方法 = 检查方法_In, 执行标记 = 执行标记_In, 执行频次 = 执行频次_In, 频率次数 = 频率次数_In, 频率间隔 = 频率间隔_In,
      间隔单位 = 间隔单位_In, 执行时间方案 = 执行时间方案_In, 计价特性 = 计价特性_In, 执行科室id = 执行科室id_In, 执行性质 = 执行性质_In, 可否分零 = 可否分零_In,
      --药品根据外购药,出院带药的调整时会发生变化
      紧急标志 = 紧急标志_In, 开始执行时间 = 开始执行时间_In, 执行终止时间 = 执行终止时间_In,
      --!长嘱的终止时间可以修改,临嘱应该为空
      病人科室id = 病人科室id_In,
      --修改时更新为病人的当前科室
      开嘱科室id = 开嘱科室id_In,
      --修改后会根据当前科室变化
      开嘱医生 = 开嘱医生_In, 审核标记 = Decode(Nvl(Instr(开嘱医生_In, '/'), 0), 0, Decode(审核标记, 1, Null, 审核标记), 1),
      --护士开医嘱时可以更改
      开嘱时间 = 开嘱时间_In,
      --补录的可以修改
      摘要 = 摘要_In, 零费记帐 = 零费记帐_In, 手术时间 = Decode(诊疗类别, 'F', To_Date(标本部位_In, 'yyyy-mm-dd hh24:mi:ss'), Null),
      用药目的 = 用药目的_In, 用药理由 = 用药理由_In, 审核状态 = 审核状态_In, 超量说明 = 超量说明_In, 首次用量 = 首次用量_In, 手术情况 = 手术情况_In, 组合项目id = 组合项目id_In,
      皮试结果 = 皮试结果_In,
      --合理用药监测
      处方序号 = 处方序号_In, 会诊医嘱id = 会诊医嘱id_In
  Where ID = Id_In;

  --病人医嘱状态:更新医生新开这条
  --因为可能同时：新开(修改)->自动校对(住院医生发送)->互斥自动停止(住院医生发送临嘱停止),因此分别-2,-1秒
  If 医嘱状态_In <> -1 Then
    Update 病人医嘱状态
    Set 操作人员 = v_人员姓名, 操作时间 = Sysdate - 2 / 60 / 60 / 24
    Where 医嘱id = Id_In And 操作类型 = 1; --新开这条始终有,校对疑问保留作为历史记录
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_Update;
/

--125083:胡俊勇,2018-05-03,危急值删除
Create Or Replace Procedure Zl_病人危急值记录_Delete(Id_In In 病人危急值记录.Id%Type) Is
Begin
  Delete 病人危急值记录 Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人危急值记录_Delete;
/

--113688:董露露,2018-04-28,取消再入院的门诊留观病人的登记后病人信息中的入出院信息没有更新
Create Or Replace Procedure Zl_入院病案主页_Delete
(
  病人id_In     病案主页.病人id%Type,
  主页id_In     病案主页.主页id%Type,
  转留观_In     Number := 0,
  清除住院号_In Number := 0
  --功能：取消病人入院/预约登记
  --     主页ID_IN:为0时表示取消预约登记
  --     转留观_IN:将正常入院登记病人转为住院留观病人
  --     清除住院号_In:第一次住院的病人转留观时是否清除住院号
) As
  v_入院时间   病案主页.入院日期%Type;
  v_入院科室   病案主页.入院科室id%Type;
  v_出院时间   病案主页.出院日期%Type;
  v_住院号     病案主页.住院号%Type;
  v_再入院     病案主页.再入院%Type;
  v_出院科室id 病案主页.出院科室id%Type;
  n_病人性质   病案主页.病人性质%Type;
  n_主页id     病案主页.主页id%Type;

  v_Count Number;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select Nvl(状态, 0), Nvl(病人性质, 0)
  Into v_Count, n_病人性质
  From 病案主页
  Where 病人id = 病人id_In And 主页id = 主页id_In;
  If v_Count <> 1 Then
    v_Error := '该病人已经入科,请先将病人撤消至入院状态。';
    Raise Err_Custom;
  End If;

  --删除电子病历时机
  Select 出院科室id, 再入院 Into v_出院科室id, v_再入院 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
  If v_再入院 = 0 Then
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '入院', v_出院科室id);
  Else
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '再次入院', v_出院科室id);
  End If;

  --提取最近一次不为空的住院号
  Begin
    If 主页id_In = 0 Then
      Select 住院号
      Into v_住院号
      From 病案主页
      Where 病人id = 病人id_In And
            主页id =
            (Select Max(主页id) From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) <> 0 And Nvl(住院号, 0) <> 0);
    Else
      Select 住院号
      Into v_住院号
      From 病案主页
      Where 病人id = 病人id_In And
            主页id =
            (Select Max(主页id) From 病案主页 Where 病人id = 病人id_In And 主页id < 主页id_In And Nvl(住院号, 0) <> 0);
    End If;
  Exception
    When Others Then
      Null;
  End;

  If 转留观_In = 1 And Nvl(主页id_In, 0) <> 0 Then
    Update 病案主页
    Set 病人性质 = 2, 住院号 = Decode(清除住院号_In, 1, Null, 住院号)
    Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(病人性质, 0) = 0;
  
    --调整住院次数
    Update 病人信息 Set 住院次数 = Decode(Sign(住院次数 - 1), 1, 住院次数 - 1, Null) Where 病人id = 病人id_In;
    If 清除住院号_In = 1 Then
      Update 病人信息 Set 住院号 = v_住院号 Where 病人id = 病人id_In;
    End If;
  Else
    Begin
      Select b.入院日期, b.出院日期, b.入院科室id
      Into v_入院时间, v_出院时间, v_入院科室
      From 病人信息 A, 病案主页 B
      Where a.病人id = 病人id_In And a.病人id = b.病人id And a.主页id = b.主页id And Nvl(b.主页id, 0) <> 0;
    Exception
      When Others Then
        Null;
    End;
    --撤消预约登记病人不检查住院日报
    If Nvl(主页id_In, 0) <> 0 Then
      Select Zl_住院日报_Count(v_入院科室, v_入院时间) Into v_Count From Dual;
      If v_Count > 0 Then
        v_Error := '已产生业务时间内的住院日报,不能办理该业务!';
        Raise Err_Custom;
      End If;
    End If;
    --门诊留观病人下达入院通知后存在两条有效的病案主页记录（36549）
    Select Count(*) Into v_Count From 病案主页 Where 病人id = 病人id_In And 入院日期 Is Not Null And 出院日期 Is Null;
    If Not v_Count > 1 Then
      v_Count := 0;
      If Nvl(主页id_In, 0) <> 0 And Nvl(n_病人性质, 0) = 0 Then
        v_Count := 1;
      End If;
      --再入院病人,取消入院登记时,病人信息的入院时间和出院时间应该回退到上一次入院日期和出院日期
      If v_再入院 = 1 Then
        Begin
          Select 入院日期, 出院日期
          Into v_入院时间, v_出院时间
          From 病案主页
          Where 病人id = 病人id_In And
                主页id = (Select Max(主页id)
                        From 病案主页
                        Where 病人id = 病人id_In And 主页id < 主页id_In);
    	Exception
      		When Others Then
        	Null;
        End;
      End If;    
      Update 病人信息
      Set 住院号 = v_住院号, 住院次数 = Decode(v_Count, 0, 住院次数, Decode(Sign(住院次数 - 1), 1, 住院次数 - 1, Null)), 当前科室id = Null,
          当前病区id = Null, 当前床号 = Null, 入院时间 = v_入院时间, 出院时间 = v_出院时间, 担保人 = Null, 担保额 = Null, 担保性质 = Null, 在院 = Null
      Where 病人id = 病人id_In;
      Delete From 在院病人 Where 病人id = 病人id_In;
    End If;
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
    Delete From 病人诊断记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = 2;
  
    --本次住院如果交了预交款,改为当作门诊交的
    Update 病人预交记录 Set 主页id = Null Where 病人id = 病人id_In And 主页id = 主页id_In;
  
    --本次发卡的,改变门诊发卡
    Update 住院费用记录 Set 主页id = Null Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 5;
  
    --本次住院的所有费用记录无结算且已全部冲销，则将对应费用记录中的"主页ID"清除。
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From 住院费用记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记帐费用 = 1 And 结帐id Is Not Null;
  
    If v_Count = 0 Then
      Begin
        Select Nvl(Count(*), 0)
        Into v_Count
        From 住院费用记录
        Where 病人id = 病人id_In And 主页id = 主页id_In And 记帐费用 = 1
        Group By NO, 记录性质, 序号
        Having Nvl(Sum(实收金额), 0) <> 0;
      Exception
        When Others Then
          v_Count := 0;
      End;
    
      If v_Count = 0 Then
        Delete 病人未结费用 Where 病人id = 病人id_In And 主页id = 主页id_In And 金额 = 0;
        Update 住院费用记录 Set 主页id = Null Where 病人id = 病人id_In And 主页id = 主页id_In And 记帐费用 = 1;
      End If;
    End If;
  
    --本次住院所有医嘱记录都已作废
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From 病人医嘱记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(医嘱状态, 0) <> 4;
    If v_Count = 0 Then
      Delete From 病人医嘱记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
    End If;
  
    --以下表,没有建病案主页(病人ID,主页ID)的外键,因为其主页ID可能是挂号ID
    Delete From 病人过敏记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 病人诊断记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 病人新生儿记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 电子病历记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 电子病历打印 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    --如果入院发放了就诊卡,则删除会失败(病人费用记录主页ID有外键约束)
    Delete From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    --修改病人信息的主页ID和住院次数
    Select Max(主页id) Into n_主页id From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) <> 0;
    Update 病人信息 Set 主页id = n_主页id Where 病人id = 病人id_In;
    If n_主页id Is Null Then
      Update 病人信息 Set 住院次数 = Null Where 病人id = 病人id_In;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_入院病案主页_Delete;
/
--124963:冉俊明,2018-04-28,预交款在门诊使用后重打了发票，在门诊费用转住院时报错
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
  
    Select Count(NO), Sum(实收金额)
    Into n_Count, n_实收金额
    From 门诊费用记录
    Where NO = No_In And Mod(记录性质, 10) = 1;
    If n_Count = 0 Or n_实收金额 = 0 Then
      v_Err_Msg := '单据' || No_In || '不是收费单据或因并发原因他人操作了该单据,不能转为住院费用.';
      Raise Err_Item;
    End If;
  
    Select 结帐id, 病人id, 开单部门id, 开单人
    Into n_原结帐id, n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And Rownum < 2;
  
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
      Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 1;
  
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
                                                     a.结帐id = b.结帐id)) And Mod(记录性质, 10) = 1 And 记录状态 <> 0) Loop
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
      For r_Prepay In (Select NO, Max(Decode(记录性质, 1, 实际票号, Null)) As 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行,
                              单位帐号, Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位
                       From 病人预交记录 A
                       Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids)))
                       Group By NO, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位) Loop
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
      Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 1;
      For r_Clinic In (Select 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
                              发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, Sum(应收金额) As 应收金额,
                              Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id, 划价人, Max(记帐单id) As 记帐单id, 发生时间,
                              实际票号
                       From 门诊费用记录
                       Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 In (2, 3) And Nvl(附加标志, 0) Not In (8, 9)
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
                  结算方式 = r_Pay.结算方式;
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
      For r_Prepay In (Select NO, Max(Decode(记录性质, 1, 实际票号, Null)) As 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行,
                              单位帐号, 收款时间, -1 * Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质
                       From 病人预交记录 A
                       Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                             Nvl(冲预交, 0) <> 0
                       Group By NO, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
                                合作单位, 结算性质) Loop
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
        For r_Prepay In (Select NO, Max(Decode(记录性质, 1, 实际票号, Null)) As 实际票号, 病人id, 主页id, 科室id, Max(结算方式) As 结算方式, 结算号码,
                                缴款单位, 单位开户行, 单位帐号, 收款时间, -1 * Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质
                         From 病人预交记录 A
                         Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                               Nvl(冲预交, 0) <> 0
                         Group By NO, 病人id, 主页id, 科室id, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
                                  合作单位, 结算性质) Loop
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

--123754:冉俊明,2018-04-26,医生工作站预约挂号性能问题
Create Or Replace Procedure Zl1_Auto_Buildingregisterplan
(
  挂号时间_In In Date := Null,
  号源id_In   临床出诊号源.Id%Type := Null
) As
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
                         a.审核时间 Is Not Null And c_日期.日期 Between 开始时间 And 终止时间)
            Where 行号 = 1;
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

--123726:冉俊明,2018-04-26,分批卫材费用销帐时报错
Create Or Replace Procedure Zl_病人费用销帐_Audit
(
  Id_In       病人费用销帐.费用id%Type,
  申请时间_In 病人费用销帐.申请时间%Type,
  审核人_In   病人费用销帐.审核人%Type,
  审核时间_In 病人费用销帐.审核时间%Type,
  状态_In     病人费用销帐.状态%Type,
  Int自动退料 Integer := 1,
  申请类别_In 病人费用销帐.申请类别%Type := 1 --对药品和卫材有效,缺省为已执行的药品或卫材 
) As
  n_执行状态       住院费用记录.执行状态%Type;
  n_申请类别       病人费用销帐.申请类别%Type;
  v_收费类别       住院费用记录.收费类别%Type;
  v_No             住院费用记录.No%Type;
  n_实际数量       药品收发记录.实际数量%Type;
  n_数量           病人费用销帐.数量%Type;
  n_收发id         药品收发记录.Id%Type;
  n_医嘱id         住院费用记录.Id%Type;
  v_跟踪在用       材料特性.跟踪在用%Type;
  n_收费细目id     住院费用记录.收费细目id%Type;
  n_审核部门id     病人费用销帐.审核部门id%Type;
  n_执行部门id     住院费用记录.执行部门id%Type;
  n_病人id         住院费用记录.病人id%Type;
  n_主页id         住院费用记录.主页id%Type;
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);

  n_Cnt     Number(18);
  n_Temp    Number(18);
  v_Err_Msg Varchar2(300);
  Err_Item Exception;
Begin

  n_申请类别 := 0;
  Select a.执行状态, a.收费类别, a.收费细目id, a.执行部门id, a.No, Nvl(b.跟踪在用, 0), a.医嘱序号, 病人id, 主页id
  Into n_执行状态, v_收费类别, n_收费细目id, n_执行部门id, v_No, v_跟踪在用, n_医嘱id, n_病人id, n_主页id
  From 住院费用记录 A, 材料特性 B
  Where a.Id = Id_In And a.收费细目id = b.材料id(+);

  If Nvl(n_主页id, 0) <> 0 Then
  
    n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
    n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
    If n_病人审核方式 = 1 Or n_未入科禁止记账 = 1 Then
      Begin
        Select 审核标志, 状态
        Into n_审核标志, n_住院状态
        From 病案主页
        Where 病人id = Nvl(n_病人id, 0) And 主页id = Nvl(n_主页id, 0);
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
  If Instr(',5,6,7', ',' || v_收费类别) > 0 Or (v_收费类别 = '4' And Nvl(v_跟踪在用, 0) = 1) Then
    n_申请类别 := 申请类别_In;
  End If;

  Update 病人费用销帐
  Set 审核人 = 审核人_In, 审核时间 = 审核时间_In, 状态 = 状态_In
  Where 费用id = Id_In And 申请类别 = n_申请类别 And 申请时间 = 申请时间_In And 状态 = 0
  Returning 数量, 审核部门id Into n_数量, n_审核部门id;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '销帐审核失败,当前操作的记录可能因为并发操作已经被他人处理,请先刷新信息!';
    Raise Err_Item;
  End If;

  If n_申请类别 = 0 And (Instr(',5,6,7', ',' || v_收费类别) > 0 Or (v_收费类别 = '4' And Nvl(v_跟踪在用, 0) = 1)) Then
    --需要检查未执行的数量必须全部申请,才会通过 
    Select Sum(Nvl(付数, 0) * Nvl(实际数量, 0))
    Into n_实际数量
    From 药品收发记录
    Where 审核日期 Is Null And 费用id = Id_In And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0;
    If Nvl(n_实际数量, 0) < Nvl(n_数量, 0) Then
      Select '在单据号<<' || v_No || '>>中' || Decode(v_收费类别, '4', '卫材', '药品') || '为:' || Chr(13) || 编码 || '-' || 名称 ||
              Chr(13) || '的申请数量(' || LTrim(To_Char(n_数量, '9999999990.99')) || ')大于了待发' || Decode(v_收费类别, '4', '料', '药') ||
              '数量(' || LTrim(To_Char(Nvl(n_实际数量, 0), '9999999990.99')) || '),不允许审核!'
      Into v_Err_Msg
      From 收费项目目录
      Where ID = n_收费细目id;
      Raise Err_Item;
    End If;
  
    If n_医嘱id <> 0 Then
      Select Nvl(Max(d.Id), 0)
      Into n_Cnt
      From 病人医嘱记录 A, 病人医嘱发送 B, 输液配药记录 D
      Where a.Id = n_医嘱id And a.Id = b.医嘱id And b.No = v_No And a.相关id = d.医嘱id And b.发送号 = d.发送号 And b.记录性质 = 2 And
            d.操作时间 = 申请时间_In And d.操作状态 = 9;
    
      If n_Cnt <> 0 Then
        Select Count(1)
        Into n_Temp
        From 输液配药状态
        Where 配药id = n_Cnt And 操作类型 = 10 And 操作时间 = 审核时间_In;
        If n_Temp = 0 Then
          Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (n_Cnt, 10, 审核人_In, 审核时间_In);
        End If;
        Update 输液配药记录 Set 操作人员 = 审核人_In, 操作时间 = 审核时间_In, 操作状态 = 10 Where ID = n_Cnt;
      End If;
    End If;
  End If;

  If n_执行状态 <> 0 Then
    If Instr(',5,6,7,', ',' || v_收费类别 || ',') > 0 And n_申请类别 = 1 Then
      If n_执行部门id <> n_审核部门id Then
        Begin
          Select '[' || 编码 || ']' || 名称 Into v_Err_Msg From 收费项目目录 Where ID = n_收费细目id;
        Exception
          When Others Then
            v_Err_Msg := '';
        End;
        v_Err_Msg := '在销帐审核时,药品为' || v_Err_Msg || ' 的已经被执行科室执行,不能再进行销帐审核,请取消审核!';
        Raise Err_Item;
      End If;
    End If;
  
    If v_收费类别 = '4' Then
      If v_跟踪在用 = 1 Then
        If n_执行部门id <> n_审核部门id And n_申请类别 = 1 And Int自动退料 <> 1 Then
          Begin
            Select '[' || 编码 || ']' || 名称 Into v_Err_Msg From 收费项目目录 Where ID = n_收费细目id;
          Exception
            When Others Then
              v_Err_Msg := '';
          End;
          v_Err_Msg := '在销帐审核时,卫材为' || v_Err_Msg || ' 的已经被执行科室执行,不能再进行销帐审核,请取消审核!';
          Raise Err_Item;
        End If;
      
        If n_申请类别 = 1 And Int自动退料 = 1 Then
          n_收发id := -1;
          --可能来自于多个批次 
          For c_收发记录 In (Select ID, 批号, Nvl(Sum(Nvl(付数, 1) * 实际数量), 0) As 数次
                         From 药品收发记录
                         Where 费用id = Id_In And 单据 In (25, 26) And (记录状态 = 1 Or Mod(记录状态, 3) = 0)
                         Group By ID, 批号) Loop
            n_收发id := c_收发记录.Id;
            If n_数量 = 0 Then
              Exit;
            End If;
          
            If n_数量 > c_收发记录.数次 Then
              n_Temp := c_收发记录.数次;
              n_数量 := n_数量 - c_收发记录.数次;
            Else
              n_Temp := n_数量;
              n_数量 := 0;
            End If;
            Zl_材料收发记录_部门退料(c_收发记录.Id, 审核人_In, 审核时间_In, c_收发记录.批号, Null, Null, n_Temp, 0);
          End Loop;
          If n_收发id = -1 Then
            v_Err_Msg := '在销帐审核时,卫材为' || v_Err_Msg || ' 的未找到相关的药品收发信息,可能是因为中途' || Chr(13) ||
                         '更改了卫材的跟踪属性,不能再进行销帐审核,请取消审核!';
            Raise Err_Item;
          End If;
        End If;
      Else
        --不是跟踪的卫材 
        Update 住院费用记录 Set 执行状态 = 0 Where ID = Id_In;
      End If;
    Elsif Instr(',5,6,7,', ',' || v_收费类别 || ',') = 0 Then
      --可能存在部分消帐,所以先将非药品的处理成部分执行,再在销帐审核过程(ZL_住院记帐记录_Delete)中处理,处理规则如下: 
      --在调用本过程时: 
      --   1.如果是已经执行的,则改为部分执行(执行状态=2);再在销帐过程中处理这部分数据(ZL_住院记帐记录_Delete):即:如果执行状态=2,并且部分销帐的,则改为1(已执行) 
      --      原因是因为非药品类只能存在两种状态.已执行;2-未执行 
      --   2.如果是未执行的,则执行状态还是为0,而在销帐过程中记录状态保持不变 
      Update 住院费用记录 Set 执行状态 = Decode(Nvl(执行状态, 0), 0, 0, 2) Where ID = Id_In; --非药品由于没有取消执行的操作,所以对已执行的要先改状态才能调销帐 
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人费用销帐_Audit;
/

--124675:张永康,2018-04-20,在转出3201医院的历史数据测试中发现的特殊数据的处理
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

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/

--124292:殷瑞,2018-04-19,修正医嘱销帐回退后，输液配药记录中的操作状态未改变的情况
Create Or Replace Procedure Zl_输液配药记录_医嘱回退
(
  医嘱id_In   In 输液配药记录.医嘱id%Type,
  发送号_In   In 输液配药记录.发送号%Type,
  操作人员_In In 输液配药记录.操作人员%Type := Null,
  操作时间_In In 输液配药记录.操作时间%Type := Null
) Is
  n_Count Number(5);
Begin
  --只对状态=1(未配药)的记录处理，如果已经配药了，则通过销账方式处理
  Select Count(ID) Into n_Count From 输液配药记录 Where 操作状态 in (1,10)  And 医嘱id = 医嘱id_In And 发送号 = 发送号_In;

  If n_Count > 0 Then
    Update 输液配药记录
    Set 操作状态 = 12, 操作人员 = Nvl(操作人员_In, Zl_Username), 操作时间 = Nvl(操作时间_In, Sysdate)
    Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In;
  
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间)
      Select ID, 12, Nvl(操作人员_In, Zl_Username), Nvl(操作时间_In, Sysdate)
      From 输液配药记录
      Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_医嘱回退;
/

--120692:刘鹏飞,2018-04-17,护理记录支持检验项目导入
Create Or Replace Procedure Zl_护理内容导入定义_Update
(
  类别_In 护理内容导入定义.类别%Type,
  名称_In 护理内容导入定义.名称%Type,
  格式_In 护理内容导入定义.格式%Type
) Is
Begin
  Update 护理内容导入定义 Set 名称 = 名称_In, 格式 = 格式_In Where 类别 = 类别_In;
  If Sql%Rowcount = 0 Then
    Insert Into 护理内容导入定义 (类别, 名称, 格式) Values (类别_In, 名称_In, 格式_In);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_护理内容导入定义_Update;
/
--123732:蔡青松,2018-04-09,修复审核标本之后医嘱执行人不正常问题
Create Or Replace Procedure Zl_检验标本记录_报告审核
(
  Id_In       检验标本记录.Id%Type,
  审核人_In   检验标本记录.审核人%Type := Null,
  人员编号_In 人员表.编号%Type := Null,
  人员姓名_In 人员表.姓名%Type := Null
) Is

  --未审核的费用行(不包含药品) 
  Cursor c_Verify(v_医嘱id In Number) Is
    Select Distinct 2 As 记录性质, NO, 序号, 记录状态, 门诊标志
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
    Select Distinct 1 As 记录性质, NO, 序号, 记录状态,门诊标志
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

  v_执行  Number(1);
  v_No    病人医嘱发送.No%Type;
  v_Nonew 病人医嘱发送.No%Type;
  v_性质  病人医嘱发送.记录性质%Type;
  v_序号  Varchar2(1000);

  v_Count      Number(18);
  v_Counts     Number(18);
  v_微生物标本 Number(1) := 0;
  v_主页id     Number(18);
  v_婴儿       Number(1);
  v_年龄       Varchar2(100);
  v_仪器       Number(18);
  v_Intloop    Number;
  Err_Custom Exception;
  v_Error Varchar2(100);

  n_Par Number;
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

  --先判断医嘱是否收费
  n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
  If n_Par = 1 Then
    For r_Samplequest In c_Samplequest(v_微生物标本) Loop
      For r_相关医嘱 In (Select ID As 医嘱id From 病人医嘱记录 Where 相关id = r_Samplequest.医嘱id) Loop
        For r_Verify In c_Verify(r_相关医嘱.医嘱id) Loop
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
    End Loop;
  End If;

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
      Set 执行状态 = 1, 完成人 = Decode(审核人_In, Null, 人员姓名_In, 审核人_In), 完成时间 = Sysdate
      Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id));
    
      Update 病人医嘱发送
      Set 执行状态 = 1, 完成人 = 人员姓名_In, 完成时间 = Sysdate
      Where 医嘱id In (Select 相关id
                     From 病人医嘱记录
                     Where ID In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
    
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
        Select Count(*) Into v_Counts From 病人医嘱记录 Where 相关id = r_Samplequest.医嘱id;
        If v_Counts > 0 Then
          For r_相关医嘱 In (Select ID As 医嘱id From 病人医嘱记录 Where 相关id = r_Samplequest.医嘱id) Loop
            For r_Verify In c_Verify(r_相关医嘱.医嘱id) Loop
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
          End Loop;
        Else
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
        End If;
        If v_序号 Is Not Null Then
          If v_性质 = 1 Then
            Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          Elsif v_性质 = 2 Then
            Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          End If;
          v_序号 := Null;
          --  v_性质 := null; 
        End If;
      End If;
    
      --审核试剂消耗单 
      v_Intloop := 1;
    
      Select 仪器id Into v_仪器 From 检验标本记录 Where ID = Id_In;
      For r_检验试剂 In (Select c.材料id, c.数量
                     From 病人医嘱记录 A, 检验报告项目 B, 检验试剂关系 C
                     Where a.相关id = r_Samplequest.医嘱id And a.诊疗项目id = b.诊疗项目id And b.报告项目id = c.项目id And c.仪器id = v_仪器) Loop
        Zl_检验试剂记录_Insert(r_Samplequest.医嘱id, v_Intloop, r_检验试剂.材料id, r_检验试剂.数量);
        v_Intloop := v_Intloop + 1;
      End Loop;
      Select Count(*) Into v_Intloop From 检验试剂记录 Where 医嘱id = r_Samplequest.医嘱id And NO Is Null;
      If v_Intloop > 1 Then
        v_Nonew := Nextno(14);
        Update 检验试剂记录 Set NO = v_Nonew Where 医嘱id = r_Samplequest.医嘱id;
      End If;
      If v_Nonew Is Not Null Then
      
        Zl_检验试剂记录_Bill(r_Samplequest.医嘱id, v_Nonew);
      
        v_主页id := Null;
        Select 主页id Into v_主页id From 病人医嘱记录 A Where ID = r_Samplequest.医嘱id;
      
        If v_主页id Is Null Then
          Zl_门诊记帐记录_Verify(v_Nonew, 人员编号_In, 人员姓名_In);
        Else
          Zl_住院记帐记录_Verify(v_Nonew, 人员编号_In, 人员姓名_In);
        End If;
      
        --如果记帐没有自动发料,则自动发料,否则不处理 
        For r_Stuff In c_Stuff(v_Nonew, v_主页id) Loop
          Zl_材料收发记录_处方发料(r_Stuff.库房id, 25, v_Nonew, 人员姓名_In, 人员姓名_In, 人员姓名_In, 1, Sysdate);
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

--122764:李业庆,2018-04-12,金额精度修改
Create Or Replace Procedure Zl_药品协定对照出库_Insert
(
  No_In         In 药品收发记录.No%Type,
  出库类别id_In In 药品收发记录.入出类别id%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  数量精度_In   In Number := 2,
  成本价精度_In In Number := 2,
  售价精度_In   In Number := 2,
  金额精度_In   In Number := 2
) As
  v_Maxserial 药品收发记录.序号%Type;
  n_收发id    药品收发记录.Id%Type;

  Cursor c_组成药品 Is
    Select (Rownum + v_Maxserial) As 序号, 协定药品id, 上次产地, 上次供应商id, 上次批号, 效期, 上次生产日期, 批准文号, 数量, 摘要, 填制人, 填制日期, 药品id, 药品序号,
           对方部门id, 库房id, 成本价, Round(Round(成本价, 成本价精度_In) * 数量, 金额精度_In) As 成本金额, 售价,
           Round(Round(售价, 售价精度_In) * 数量, 金额精度_In) As 售价金额, Round(售价 * 数量, 金额精度_In) - Round(成本价 * 数量, 金额精度_In) As 差价
    From (Select a.协定药品id, c.上次产地, c.上次供应商id, c.上次批号, c.效期, c.上次生产日期, c.批准文号,
                  Round(Round(e.实际数量, 数量精度_In) * (a.分子 / a.分母), 数量精度_In) As 数量, e.摘要, e.填制人, e.填制日期, e.药品id, e.序号 As 药品序号,
                  e.对方部门id, e.库房id,
                  Decode(Sign(Nvl(c.实际金额, 0)), 1, (d.现价 - d.现价 * (c.实际差价 / c.实际金额)), (d.现价 - d.现价 * (b.指导差价率 / 100))) As 成本价,
                  d.现价 As 售价
           From 协定药品对照 A, (Select b.* From 收费项目目录 A, 药品规格 B Where a.Id = b.药品id And Nvl(是否变价, 0) = 0) B,
                (Select 库房id, 药品id, 实际金额, 实际差价, 上次采购价, 上次产地, 上次供应商id, 上次批号, 效期, 上次生产日期, 批准文号
                  From 药品库存
                  Where 性质 = 1 And 库房id = 库房id_In) C,
                (Select 收费细目id, 现价
                  From 收费价目
                  Where ((Sysdate Between 执行日期 And 终止日期) Or (Sysdate >= 执行日期 And 终止日期 Is Null))) D,
                (Select * From 药品收发记录 Where NO = No_In And 单据 = 3 And 入出系数 = 1) E
           Where a.协定药品id = b.药品id And b.药品id = d.收费细目id And b.药品id = c.药品id(+) And a.药品id = e.药品id
           Union All
           Select a.协定药品id, c.上次产地, c.上次供应商id, c.上次批号, c.效期, c.上次生产日期, c.批准文号,
                  Round(Round(e.实际数量, 数量精度_In) * (a.分子 / a.分母), 数量精度_In) As 数量, e.摘要, e.填制人, e.填制日期, e.药品id, e.序号 As 药品序号,
                  e.对方部门id, e.库房id,
                  Decode(Sign(Nvl(c.实际金额, 0)), 1, (c.现价 - c.现价 * (c.实际差价 / c.实际金额)), (c.现价 - c.现价 * (b.指导差价率 / 100))) As 成本价,
                  c.现价 As 售价
           From 协定药品对照 A, (Select b.* From 收费项目目录 A, 药品规格 B Where a.Id = b.药品id And Nvl(是否变价, 0) = 1) B,
                (Select 库房id, 药品id, 实际金额, 实际差价, 上次采购价, 上次产地, 上次供应商id, 上次批号, 效期, 上次生产日期, 批准文号, 实际金额 / 实际数量 As 现价
                  From 药品库存
                  Where 性质 = 1 And 库房id = 库房id_In And 实际数量 > 0) C,
                (Select * From 药品收发记录 Where NO = No_In And 单据 = 3 And 入出系数 = 1) E
           Where a.协定药品id = b.药品id And b.药品id = c.药品id And a.药品id = e.药品id)
    Order By 协定药品id;
Begin
  Select Max(序号) Into v_Maxserial From 药品收发记录 Where NO = No_In And 单据 = 3 And 入出系数 = 1;
  For v_组成药品 In c_组成药品 Loop
    Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
  
    Insert Into 药品收发记录
      (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 产地, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期,
       费用id, 扣率, 供药单位id, 批号, 效期, 生产日期, 批准文号)
    Values
      (n_收发id, 1, 3, No_In, v_组成药品.序号, v_组成药品.对方部门id, v_组成药品.库房id, 出库类别id_In, -1, v_组成药品.协定药品id, v_组成药品.上次产地, v_组成药品.数量,
       v_组成药品.数量, v_组成药品.成本价, v_组成药品.成本金额, v_组成药品.售价, v_组成药品.售价金额, v_组成药品.差价, v_组成药品.摘要, v_组成药品.填制人, v_组成药品.填制日期,
       v_组成药品.药品id, v_组成药品.药品序号, v_组成药品.上次供应商id, v_组成药品.上次批号, v_组成药品.效期, v_组成药品.上次生产日期, v_组成药品.批准文号);
  
    --参数为1表示在填单时下可用数量
    Zl_药品库存_Update(n_收发id, 0);
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品协定对照出库_Insert;
/

--127766:殷瑞,2018-07-03,修正发药时注册证号的填写遗漏
--124583:李业庆,2018-04-20,部门发药,退药填写发药类型
CREATE OR REPLACE Procedure Zl_药品收发记录_批量发药
(
  Billinfo_In   In Varchar2, --格式:"id1,批次1|id2,批次2|....."
  Partid_In     In 药品收发记录.库房id%Type,
  People_In     In 药品收发记录.审核人%Type,
  Date_In       In 药品收发记录.审核日期%Type,
  发药方式_In   In 药品收发记录.发药方式%Type := 3,
  领药人_In     In 药品收发记录.领用人%Type := Null,
  汇总发药号_In In 药品收发记录.汇总发药号%Type := Null,
  Intdigit_In   In Number := 2,
  配药人_In     In 药品收发记录.配药人%Type := Null,
  核查人_In     In 药品收发记录.核查人%Type := Null
) Is
  --只读变量
  v_Infotmp     Varchar2(4000);
  v_Fields      Varchar2(4000);
  n_Billid      药品收发记录.Id%Type;
  n_批次        药品收发记录.批次%Type;
  Lng入出类别id Number(18);
  Int入出系数   Number;
  Int执行状态   Number;
  Int单据       药品收发记录.单据%Type;
  Strno         药品收发记录.No%Type;
  Lng库房id     药品收发记录.库房id%Type;
  Lng药品id     药品收发记录.药品id%Type;
  Lng费用id     药品收发记录.费用id%Type;
  Dbl差价率     Number;
  v_零售价      药品收发记录.零售价%Type;
  Int未发数     未发药品记录.未发数%Type;
  v_核查日期    药品收发记录.核查日期%Type;
  --可写变量
  Dbl实际数量 药品收发记录.实际数量%Type;
  Dbl实际金额 药品收发记录.零售金额%Type;
  Dbl成本金额 药品收发记录.成本金额%Type;
  Dbl实际差价 药品收发记录.差价%Type;
  --2002-07-31朱玉宝
  --LNGLAST批次 发药前确定的批次(已减可用数量)
  Str药名           Varchar2(200);
  Dbl可用数量       药品收发记录.填写数量%Type;
  Lnglast批次       药品收发记录.批次%Type;
  Lngcur批次        药品收发记录.批次%Type;
  Str批号           药品收发记录.批号%Type;
  Str效期           药品收发记录.效期%Type;
  n_上次供应商id    药品库存.上次供应商id%Type;
  n_上次采购价      药品库存.上次采购价%Type;
  v_上次产地        药品库存.上次产地%Type;
  d_上次生产日期    药品库存.上次生产日期%Type;
  v_批准文号        药品库存.批准文号%Type;
  n_记录状态        药品收发记录.记录状态%Type;
  n_平均成本价      药品库存.平均成本价%Type;
  n_发药方式        药品收发记录.发药方式%Type;
  v_摘要            药品收发记录.摘要%Type;
  Bln收费与发药分离 Number(1);
  v_Error           Varchar2(255);
  Err_Custom Exception;
  n_时价     Number(1) := 0;
  n_时价分批 Number(1) := 0;
  n_处方类型 未发药品记录.处方类型%Type;
Begin
  Select Sysdate Into v_核查日期 From Dual;
  If Billinfo_In Is Null Then
    v_Infotmp := Null;
  Else
    v_Infotmp := Billinfo_In || '|';
  End If;
  While v_Infotmp Is Not Null Loop
    --分解单据ID串
    v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
    n_Billid  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    n_批次    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');

    --获取该收发记录的单据、药品ID、库房ID,零售金额及实际数量、入出类别ID
    Begin
      Select a.单据, a.No, a.药品id, a.库房id, a.费用id, Nvl(a.零售价, 0), Nvl(a.零售金额, 0), Nvl(a.实际数量, 0) * Nvl(a.付数, 1), a.入出类别id,
             a.入出系数, Nvl(a.批次, 0), a.批号, a.效期, a.供药单位id, a.产地, a.生产日期, a.批准文号, Nvl(a.发药方式, 0), a.摘要, a.记录状态
      Into Int单据, Strno, Lng药品id, Lng库房id, Lng费用id, v_零售价, Dbl实际金额, Dbl实际数量, Lng入出类别id, Int入出系数, Lnglast批次, Str批号, Str效期,
           n_上次供应商id, v_上次产地, d_上次生产日期, v_批准文号, n_发药方式, v_摘要, n_记录状态
      From 药品收发记录 A
      Where a.Id = n_Billid And a.审核日期 Is Null
      For Update Nowait;

      Select '[' || c.编码 || ']' || c.名称, Nvl(是否变价, 0) 时价
      Into Str药名, n_时价
      From 收费项目目录 C
      Where c.Id = Lng药品id;
    Exception
      When Others Then
        Int单据 := 0;
        v_Error := '已有其他用户在执行发药，不能重复操作！';
        Raise Err_Custom;
    End;

    If n_发药方式 = -1 Or v_摘要 = '拒发' Then
      Int单据 := 0;
    End If;

    If Int单据 > 0 Then
      If Nvl(n_批次, 0) = 0 Then
        Lngcur批次 := Lnglast批次;
      Else
        Lngcur批次 := Nvl(n_批次, 0);
      End If;

      --检查是否已经填写库房
      Bln收费与发药分离 := 0;
      If Lng库房id Is Null Then
        Bln收费与发药分离 := 1;
      End If;
      Lng库房id := Partid_In;

      --取该批药品的批号
      Begin
        Select 上次批号, 效期, Nvl(可用数量, 0), 上次供应商id, 上次产地, 上次生产日期, 批准文号, 上次采购价
        Into Str批号, Str效期, Dbl可用数量, n_上次供应商id, v_上次产地, d_上次生产日期, v_批准文号, n_上次采购价
        From 药品库存
        Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(批次, 0) = Lngcur批次;
      Exception
        When Others Then
          n_上次采购价 := 0;
          Dbl可用数量  := 0;
      End;

      --可用数量不足则退出
      If Lngcur批次 <> Nvl(Lnglast批次, 0) Then
        If Dbl可用数量 < Dbl实际数量 And Lngcur批次 <> 0 Then
          v_Error := Str药名 || '的可用数量不足，操作中止！';
          Raise Err_Custom;
        End If;
      End If;

      If n_记录状态 = 1 Then
        --原始发药记录，取最新价格
        n_平均成本价 := Round(Zl_Fun_Getoutcost(Lng药品id, Lngcur批次, Lng库房id), 5);
      Else
        --退药再发记录，取原始单据价格
        Select a.成本价
        Into n_平均成本价
        From 药品收发记录 A, 药品收发记录 B
        Where b.Id = n_Billid And a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And a.药品id + 0 = b.药品id And
              a.序号 = b.序号 And Nvl(a.批次, 0) = Nvl(b.批次, 0) And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0);
      End If;

      Dbl成本金额 := Round(n_平均成本价 * Nvl(Dbl实际数量, 0), Intdigit_In);
      Dbl实际差价 := Round(Dbl实际金额 - Dbl成本金额, Intdigit_In);

      --查询处方类型
      Select 处方类型
      Into n_处方类型
      From 未发药品记录
      Where NO = Strno And 单据 = Int单据 And (库房id + 0 = Lng库房id Or 库房id Is Null);
      
      --更新药品收发记录的零售金额、成本金额及差价
      Update 药品收发记录
      Set 库房id = Lng库房id, 成本价 = n_平均成本价, 成本金额 = Dbl成本金额, 差价 = Dbl实际差价, 批次 = Lngcur批次, 批号 = Str批号, 效期 = Str效期,
          配药人 = 配药人_In, 核查人 = 核查人_In, 核查日期 = v_核查日期, 审核人 = People_In, 审核日期 = Date_In, 发药方式 = 发药方式_In, 领用人 = 领药人_In,
          汇总发药号 = 汇总发药号_In, 供药单位id = n_上次供应商id, 产地 = v_上次产地, 生产日期 = d_上次生产日期, 批准文号 = v_批准文号,注册证号 = n_处方类型
      Where ID = n_Billid;
      --并发操作检查
      If Sql%RowCount = 0 Then
        v_Error := '要发药的药品记录"' || Str药名 || '"不存在，操作中止！';
        Raise Err_Custom;
      End If;

      --更新住院费用记录的执行状态(已执行)
      Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 1, 0, 1, 2)
      Into Int执行状态
      From 药品收发记录
      Where 单据 = Int单据 And NO = Strno And 费用id = Lng费用id And 审核人 Is Null;
      Update 住院费用记录
      Set 执行状态 = Int执行状态, 执行人 = Decode(People_In, Null, Zl_Username, People_In), 执行时间 = Date_In, 执行部门id = Partid_In
      Where ID = Lng费用id;

      --更新未发药品记录(如果未发数为零则删除)
      Select Count(*)
      Into Int未发数
      From 药品收发记录
      Where 单据 = Int单据 And (库房id + 0 = Lng库房id Or 库房id Is Null) And NO = Strno And 审核人 Is Null And
            Nvl(LTrim(RTrim(摘要)), '小宝') <> '拒发';

      If Int未发数 = 0 Then
        Delete 未发药品记录 Where NO = Strno And 单据 = Int单据 And (库房id + 0 = Lng库房id Or 库房id Is Null);
      End If;

      --更新原批次库存的可用数量
      --更新发药批次库存的可用及实际数量
      If Lnglast批次 <> Lngcur批次 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Dbl实际数量
        Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(批次, 0) = Lnglast批次;

        Zl_药品库存_可用数量异常处理(Lng库房id, Lng药品id, Lnglast批次);

        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - Dbl实际数量
        Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(批次, 0) = Lngcur批次;

        Zl_药品库存_可用数量异常处理(Lng库房id, Lng药品id, Lngcur批次);
      End If;

      If n_时价 = 1 And Lngcur批次 > 0 Then
        n_时价分批 := 1;
      Else
        n_时价分批 := 0;
      End If;

      If Bln收费与发药分离 = 1 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - Dbl实际数量, 实际数量 = Nvl(实际数量, 0) - Dbl实际数量, 实际金额 = Nvl(实际金额, 0) - Dbl实际金额,
            实际差价 = Nvl(实际差价, 0) - Dbl实际差价, 平均成本价 = Decode(平均成本价, Null, n_平均成本价, 平均成本价),
            上次采购价 = Decode(上次采购价, Null, n_平均成本价, 上次采购价), 零售价 = Decode(n_时价分批, 1, Decode(零售价, Null, v_零售价, 零售价), 零售价)
        Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(批次, 0) = Lngcur批次;
      Else
        Update 药品库存
        Set 实际数量 = Nvl(实际数量, 0) - Dbl实际数量, 实际金额 = Nvl(实际金额, 0) - Dbl实际金额, 实际差价 = Nvl(实际差价, 0) - Dbl实际差价,
            平均成本价 = Decode(平均成本价, Null, n_平均成本价, 平均成本价), 上次采购价 = Decode(上次采购价, Null, n_平均成本价, 上次采购价),
            零售价 = Decode(n_时价分批, 1, Decode(零售价, Null, v_零售价, 零售价), 零售价)
        Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(批次, 0) = Lngcur批次;
      End If;

      If Sql%RowCount = 0 Then
        If n_上次采购价 = 0 Then
          If Dbl实际数量 = 0 Then
            Dbl实际数量 := 1;
          End If;
          n_上次采购价 := Round(Dbl成本金额 / Dbl实际数量, 5);
        End If;

        If Bln收费与发药分离 = 1 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次批号, 上次产地, 上次供应商id, 上次采购价, 上次生产日期, 批准文号, 效期, 平均成本价, 零售价)
          Values
            (Lng库房id, Lng药品id, Lngcur批次, 1, 0 - Dbl实际数量, 0 - Dbl实际数量, 0 - Dbl实际金额, 0 - Dbl实际差价, Str批号, v_上次产地,
             n_上次供应商id, n_平均成本价, d_上次生产日期, v_批准文号, Str效期, n_平均成本价, Decode(n_时价分批, 1, v_零售价, Null));
        Else
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 上次批号, 上次产地, 上次供应商id, 上次采购价, 上次生产日期, 批准文号, 效期, 平均成本价, 零售价)
          Values
            (Lng库房id, Lng药品id, Lngcur批次, 1, 0 - Dbl实际数量, 0 - Dbl实际金额, 0 - Dbl实际差价, Str批号, v_上次产地, n_上次供应商id, n_平均成本价,
             d_上次生产日期, v_批准文号, Str效期, n_平均成本价, Decode(n_时价分批, 1, v_零售价, Null));
        End If;
      End If;

      Zl_药品库存_可用数量异常处理(Lng库房id, Lng药品id, Lngcur批次);

      Delete 药品库存
      Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
            Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;

      --处理调价修正
      Zl_药品收发记录_调价修正(n_Billid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_批量发药;
/

--127766:殷瑞,2018-07-02,修正药品部门退药后单据类型的错误
--124583:李业庆,2018-04-20,部门发药,退药填写发药类型
CREATE OR REPLACE Procedure Zl_药品收发记录_部门退药
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
           Decode(Lng分批, 1, 批次, 3, Lng新批次, 0), Decode(Lng分批, 3, 产地_In, 1, 产地, 产地), Decode(Lng分批, 3, 批号_In, 1, 批号, Null),
           Decode(Lng分批, 3, 效期_In, 1, 效期, Null), n_付数, Dbl实际数量, Dbl实际数量, 成本价, Dbl实际成本, 扣率, 零售价, Dbl实际金额, Dbl实际差价, 摘要,
           填制人, 填制日期, Null, Null, Null, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号
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

--115765:刘涛,2018-06-28,价格精度处理
--127911,李业庆,2018-06-28,高值卫材取虚拟库房的成本价
--124862:李业庆,2018-04-26,只处理卫材未审核的发料单据
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
  n_门诊标志               Number(1);
  v_入库no                 药品收发记录.No%Type;
  v_入库库房id             药品收发记录.库房id%Type := 0;
  v_病人信息               Varchar2(200);
  n_虚拟库房               药品库存.库房id%Type;
  v_允许未审核记账单发料   Number(1);
  v_允许未收费的划价单发料 Number(1);
  v_自动审核记账单         Number(1);
  n_平均成本价             药品库存.平均成本价%Type;
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
                           Decode(Nvl(a.摘要, 'No拒发'), '拒发', 3, b.执行状态) 执行状态, 1 门诊标志
                    From 药品收发记录 A, 门诊费用记录 B
                    Where a.费用id = b.Id And a.Id = n_Id And a.审核人 Is Null And a.单据 In (24, 25, 26)
                    Union All
                    Select a.Id, a.单据, b.Id 费用id, a.No, b.No 费用no, b.病人id, b.主页id, b.病人病区id, b.病人科室id, b.开单部门id, b.执行部门id,
                           b.收入项目id, b.实收金额, b.操作员编号, b.操作员姓名, Nvl(b.记录状态, 0) As 审核标志, a.审核人,
                           Decode(Nvl(a.摘要, 'No拒发'), '拒发', 3, b.执行状态) 执行状态, 2 门诊标志
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
    
      n_门诊标志 := c_Check.门诊标志;
    
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
    
      If n_门诊标志 = 1 Then
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
              If n_门诊标志 = 1 Then
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

--125779:殷瑞,2018-05-28,退药按药品id排序处理
--125779:李业庆,2018-05-15,退药按药品id排序处理
Create Or Replace Procedure Zl_输液配药记录_销帐审核
(
  配药id_In   In Varchar2, --ID串:ID1,审核标志1,ID2,审核标志2....
  操作人员_In In 输液配药记录.操作人员%Type,
  操作时间_In In 输液配药记录.操作时间%Type
) Is
  v_Tansid     Varchar2(20);
  v_Tansids    Varchar2(4000);
  v_Tmp        Varchar2(4000);
  v_Usercode   Varchar2(100);
  v_发药id     药品收发记录.Id%Type;
  n_Count      Number(1);
  d_审核时间   药品收发记录.审核日期%Type;
  v_No         药品收发记录.No%Type;
  v_上次no     药品收发记录.No%Type;
  n_审核标志   Number(1);
  n_操作状态   Number(2);
  v_收发ids    Varchar2(4000);
  v_退药待发id 药品收发记录.Id%Type;
  v_原始id     药品收发记录.Id%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;

  Cursor c_销帐记录 Is
    Select Distinct a.费用id, b.操作时间
    From 药品收发记录 A, 输液配药记录 B, 输液配药内容 C
    Where a.Id = c.收发id And b.Id = c.记录id And b.Id = v_Tansid And b.操作状态 = 9;

  v_销帐记录 c_销帐记录%RowType;

  Cursor c_退药记录 Is
    Select /*+ rule*/
    Distinct a.Id As 退药id, c.收发id, c.数量, a.药品id, a.批次,c.记录id as 配药id
    From 药品收发记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where c.记录id = d.Column_Value And b.Id = c.收发id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And
          a.药品id + 0 = b.药品id And a.序号 = b.序号 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id, a.批次;

  v_退药记录 c_退药记录%RowType;

  Cursor c_费用销帐 Is
    Select /*+ rule*/
     a.No, a.序号 || ':' || c.数量 || ':' || c.记录id As 费用序号
    From 住院费用记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.Id = b.费用id And b.Id = c.收发id And Mod(b.记录状态, 3) = 1 And c.记录id = d.Column_Value;

  v_费用销帐 c_费用销帐%RowType;

Begin
  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;

  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid   := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    n_审核标志 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
  
    v_收发ids := Null;
  
    --统计审核确认的输液单(n_审核标志 = 1)
    If n_审核标志 = 1 Then
      If v_Tansids Is Null Then
        v_Tansids := v_Tansid;
      Else
        v_Tansids := v_Tansids || ',' || v_Tansid;
      End If;
    End If;
  
    --检查当前输液单的状态是否为待摆药状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;
    
      If n_操作状态 <> 9 Then
        v_Error := '该数据已被操作，不能进行销帐审核！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    If n_审核标志 = 1 Then
      n_操作状态 := 10;
    Elsif n_审核标志 = 2 Then
      n_操作状态 := 11;
    End If;
  
    --查找输液单对应的收发NO
    Begin
      Select NO
      Into v_No
      From 药品收发记录
      Where ID In (Select 收发id From 输液配药内容 Where 记录id In (Select ID From 输液配药记录 Where ID = v_Tansid)) And Rownum < 2;
    Exception
      When Others Then
        v_No := '';
    End;
  
    --收发NO相同的配药ID，审核时间以此设置为延长1秒
    If v_No = v_上次no Then
      d_审核时间 := d_审核时间 + 1 / 24 / 60 / 60;
    Else
      d_审核时间 := 操作时间_In;
      v_上次no   := v_No;
    End If;
  
    --销帐记录处理
    For v_销帐记录 In c_销帐记录 Loop
      Zl_病人费用销帐_Audit(v_销帐记录.费用id, v_销帐记录.操作时间, 操作人员_In, d_审核时间, n_审核标志);
    End Loop;
  
    Select Count(*) Into n_Count From 输液配药状态 Where 配药id = v_Tansid And 操作时间 = 操作时间_In;
  
    If n_Count <> 1 Then
      Insert Into 输液配药状态
        (配药id, 操作类型, 操作人员, 操作时间)
      Values
        (v_Tansid, n_操作状态, 操作人员_In, 操作时间_In);
    End If;
    Update 输液配药记录 Set 操作人员 = 操作人员_In, 操作时间 = 操作时间_In, 操作状态 = n_操作状态 Where ID = v_Tansid;
  End Loop;

  --先退药
  For v_退药记录 In c_退药记录 Loop
    Zl_药品收发记录_部门退药(v_退药记录.退药id, 操作人员_In, 操作时间_In, Null, Null, Null, v_退药记录.数量, Null, 操作人员_In);
  
    --取退药待发id
    Select a.Id
    Into v_发药id
    From 药品收发记录 A, 药品收发记录 B
    Where b.Id = v_退药记录.退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
          a.序号 = b.序号 And Mod(a.记录状态, 3) = 1 And a.审核日期 Is Null;
  
    --输液配药内容中的收发ID更新为退药待发的收发ID
    Update 输液配药内容 Set 收发id = v_发药id Where 记录id = v_退药记录.配药id And 收发id = v_退药记录.收发id;
  
    If v_收发ids Is Null Then
      v_收发ids := v_发药id;
    Else
      v_收发ids := v_收发ids || ',' || v_发药id;
    End If;
  
    --取原始id
    Select a.Id
    Into v_原始id
    From 药品收发记录 A, 药品收发记录 B
    Where b.Id = v_退药记录.退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
          a.序号 = b.序号 And Mod(a.记录状态, 3) = 0 And a.审核日期 Is Not Null;
  
    Insert Into 输液配药内容
      (记录id, 收发id, 数量)
      Select 记录id, v_原始id, 数量 From 输液配药内容 Where 记录id = v_退药记录.配药id And 收发id = v_发药id;
  
    v_收发ids := v_收发ids || ',' || v_原始id;
  End Loop;

  --费用销帐
  For v_费用销帐 In c_费用销帐 Loop
    Zl_住院记帐记录_Delete(v_费用销帐.No, v_费用销帐.费用序号, v_Usercode, Zl_Username, 2, 1, 1, d_审核时间);
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_销帐审核;
/

--127629:李业庆,2018-06-22,去掉过程中的commit
Create Or Replace Procedure Zl1_Autocloseaccount Is
  v_Lngid    药品结存记录.Id%Type;
  d_开始日期 药品结存记录.期初日期%Type;
  d_结束日期 药品结存记录.期末日期%Type;
  n_结存时点 Number(2);
  v_Error    Varchar2(255);
  Err_Custom Exception;
  d_计算日期     药品结存记录.期末日期%Type;
  n_结存id       药品结存记录.Id%Type;
  n_未审核结存id 药品结存记录.Id%Type;

  Cursor c_Stock Is
    Select Distinct b.Id
    From 部门性质说明 A, 部门表 B
    Where a.部门id = b.Id And a.工作性质 In ('西药库', '成药库', '中药库', '西药房', '成药房', '中药房', '制剂室') And
          To_Char(b.撤档时间, 'yyyy-MM-dd') = '3000-01-01'
    Order By b.Id;
  r_Stock c_Stock%RowType;
Begin
  --取结存时点，默认每月最后一日结存
  n_结存时点 := Nvl(zl_GetSysParameter(221), 0);

  --只有自动结存才走此过程，手工结存不在走此过程
  If n_结存时点 <> -1 Then
    --计算本次结存的结束日期；因为自动结存是对前一天数据进行结存，所以需要按当前日期提前一天来计算或判断，
    If n_结存时点 = 0 Or n_结存时点 > To_Number(To_Char(Trunc(Last_Day(Sysdate - 1)), 'dd')) Then
      --指定按每月最后一天结存；或者结存时点大于了本月最大天数，也按本月最后一天结存
      d_结束日期 := Trunc(Last_Day(Sysdate - 1)) + 1 - 1 / 24 / 60 / 60;
    Else
      d_结束日期 := Trunc(Sysdate - 1, 'MONTH') + n_结存时点 - 1 / 24 / 60 / 60;
    End If;
  
    --检查日期，在结存时点后才能进行自动结存
    If Sysdate - d_结束日期 > 0 Then
      For r_Stock In c_Stock Loop
        --判断期间内是否有结存(不算转结)
        --此处不再通过“期间”字段进行判断，而是通过结存时间点来判断：如2016-05-28 23：59：59，如有则不结存，无则结存
        Select Nvl(Max(ID), 0)
        Into n_结存id
        From 药品结存记录
        Where 库房id = r_Stock.Id And 期末日期 = d_结束日期 And 取消人 Is Null;
      
        If n_结存id > 0 Then
          --如果当前期间已经结存过了，就不再结存，一个期间只结存一次
          Null;
        Else
          --取库房最大的结存ID和本次结存的开始日期
          Select Nvl(Max(ID), 0), Max(期末日期) + 1 / 24 / 60 / 60
          Into n_结存id, d_开始日期
          From 药品结存记录
          Where 库房id = r_Stock.Id And 取消人 Is Null;
        
          --开始时间不能大于结束时间
          If d_开始日期 <= d_结束日期 Then
            If n_结存id > 0 Then
              --检查是否存在未审核的结存，如果存在则自动审核(通常情况都是在期间内手工审核)
              Select Nvl(Max(ID), 0)
              Into n_未审核结存id
              From 药品结存记录
              Where 库房id = r_Stock.Id And 审核日期 Is Null;
            
              If n_未审核结存id > 0 Then
                Zl_药品结存记录_Verify(n_未审核结存id, Zl_Username);
              End If;
            
              --产生新的结存记录
              Select 药品结存记录_Id.Nextval Into v_Lngid From Dual;
            
              Insert Into 药品结存记录
                (ID, 库房id, 期初日期, 期末日期, 填制人, 填制日期, 上次结存id, 期间, 性质)
              Values
                (v_Lngid, r_Stock.Id, d_开始日期, d_结束日期, Nvl(Zl_Username, 'zlhis'), Sysdate, n_结存id,
                 To_Char(Trunc(d_结束日期), 'yyyymm'), 1);
            
              --产生药品结存明细表，本期期末=本期期初(上期期末)+期间发生
              Insert Into 药品结存明细
                (结存id, 库房id, 药品id, 批次, 期初数量, 期初金额, 期初差价, 期末数量, 期末金额, 期末差价)
                Select v_Lngid, 库房id, 药品id, 批次, Sum(期初数量), Sum(期初金额), Sum(期初差价), Sum(期末数量), Sum(期末金额), Sum(期末差价)
                From (Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, a.期末数量 As 期初数量, a.期末金额 As 期初金额, a.期末差价 As 期初差价, a.期末数量,
                              a.期末金额, a.期末差价
                       From 药品结存明细 A, 药品规格 B
                       Where a.药品id = b.药品id And a.结存id = n_结存id
                       Union All
                       Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, 0 As 期初数量, 0 As 期初金额, 0 As 期初差价,
                              a.入出系数 * a.实际数量 * Nvl(a.付数, 1) As 期末数量, a.入出系数 * a.零售金额 As 期末金额, a.入出系数 * a.差价 As 期末差价
                       From 药品收发记录 A, 药品规格 B
                       Where a.药品id = b.药品id And a.库房id + 0 = r_Stock.Id And a.审核日期 Between d_开始日期 And d_结束日期)
                Group By 库房id, 药品id, 批次
                Order By 库房id, 药品id, 批次;
            
              --计算误差：本期期末-库存记录(减去本期期末时间后发生的数据)
              Insert Into 药品结存误差
                (ID, 结存id, 库房id, 药品id, 批次, 数量差, 金额差, 差价差)
                Select 药品结存误差_Id.Nextval, v_Lngid, a.库房id, a.药品id, a.批次, a.实际数量 As 数量差, a.实际金额 As 金额差, a.实际差价 As 差价差
                From (Select 库房id, 药品id, 批次, Sum(实际数量) As 实际数量, Sum(实际金额) As 实际金额, Sum(实际差价) As 实际差价
                       From (Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, Nvl(a.期末数量, 0) As 实际数量, Nvl(a.期末金额, 0) As 实际金额,
                                     Nvl(a.期末差价, 0) As 实际差价
                              From 药品结存明细 A, 药品规格 B
                              Where a.药品id = b.药品id And a.结存id = v_Lngid
                              Union All
                              Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, -1 * Nvl(a.实际数量, 0) As 实际数量,
                                     -1 * Nvl(a.实际金额, 0) As 实际金额, -1 * Nvl(实际差价, 0) As 实际差价
                              From 药品库存 A, 药品规格 B
                              Where a.药品id = b.药品id And a.性质 = 1 And a.库房id = r_Stock.Id
                              Union All
                              Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, a.入出系数 * a.实际数量 * Nvl(a.付数, 1) As 实际数量,
                                     a.入出系数 * a.零售金额 As 实际金额, a.入出系数 * a.差价 As 实际差价
                              From 药品收发记录 A, 药品规格 B
                              Where a.药品id = b.药品id And a.库房id = r_Stock.Id And a.审核日期 > d_结束日期) A
                       Group By 库房id, 药品id, 批次) A
                Where a.实际数量 <> 0 Or a.实际金额 <> 0 Or a.实际差价 <> 0;
              --自动结存后立马审核结存信息
              Zl_药品结存记录_Verify(v_Lngid, Zl_Username);
            End If;
          End If;
        End If;
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Autocloseaccount;
/

--127629:李业庆,2018-06-22,去掉过程中的commit
Create Or Replace Procedure Zl1_Autostuffcloseaccount Is
  v_Lngid    材料结存记录.Id%Type;
  d_开始日期 材料结存记录.期初日期%Type;
  d_结束日期 材料结存记录.期末日期%Type;
  n_结存时点 Number(2);
  v_Error    Varchar2(255);
  Err_Custom Exception;
  d_计算日期     材料结存记录.期末日期%Type;
  n_结存id       材料结存记录.Id%Type;
  n_未审核结存id 材料结存记录.Id%Type;

  Cursor c_Stock Is
    Select Distinct b.Id
    From 部门性质说明 A, 部门表 B
    Where a.部门id = b.Id And a.工作性质 In ('卫材库', '发料部门') And To_Char(b.撤档时间, 'yyyy-MM-dd') = '3000-01-01'
    Order By b.Id;
  r_Stock c_Stock%RowType;
Begin
  --取结存时点，默认每月最后一日结存
  n_结存时点 := Nvl(zl_GetSysParameter(281), 0);

  --只有自动结存才走此过程，手工结存不在走此过程
  If n_结存时点 <> -1 Then
    --计算本次结存的结束日期；因为自动结存是对前一天数据进行结存，所以需要按当前日期提前一天来计算或判断，
    If n_结存时点 = 0 Or n_结存时点 > To_Number(To_Char(Trunc(Last_Day(Sysdate - 1)), 'dd')) Then
      --指定按每月最后一天结存；或者结存时点大于了本月最大天数，也按本月最后一天结存
      d_结束日期 := Trunc(Last_Day(Sysdate - 1)) + 1 - 1 / 24 / 60 / 60;
    Else
      d_结束日期 := Trunc(Sysdate - 1, 'MONTH') + n_结存时点 - 1 / 24 / 60 / 60;
    End If;
  
    --检查日期，在结存时点后才能进行自动结存
    If Sysdate - d_结束日期 > 0 Then
      For r_Stock In c_Stock Loop
        --判断期间内是否有结存(不算转结)
        --此处不再通过“期间”字段进行判断，而是通过结存时间点来判断：如2016-05-28 23：59：59，如有则不结存，无则结存
        Select Nvl(Max(ID), 0)
        Into n_结存id
        From 材料结存记录
        Where 库房id = r_Stock.Id And 期末日期 = d_结束日期 And 取消人 Is Null;
      
        If n_结存id > 0 Then
          --如果当前期间已经结存过了，就不再结存，一个期间只结存一次
          Null;
        Else
          --取库房最大的结存ID和本次结存的开始日期
          Select Nvl(Max(ID), 0), Max(期末日期) + 1 / 24 / 60 / 60
          Into n_结存id, d_开始日期
          From 材料结存记录
          Where 库房id = r_Stock.Id And 取消人 Is Null;
        
          --开始时间不能大于结束时间
          If d_开始日期 <= d_结束日期 Then
            If n_结存id > 0 Then
              --检查是否存在未审核的结存，如果存在则自动审核(通常情况都是在期间内手工审核)
              Select Nvl(Max(ID), 0)
              Into n_未审核结存id
              From 材料结存记录
              Where 库房id = r_Stock.Id And 审核日期 Is Null;
            
              If n_未审核结存id > 0 Then
                Zl_材料结存记录_Verify(n_未审核结存id, Zl_Username);
              End If;
            
              --产生新的结存记录
              Select 材料结存记录_Id.Nextval Into v_Lngid From Dual;
            
              Insert Into 材料结存记录
                (ID, 库房id, 期初日期, 期末日期, 填制人, 填制日期, 上次结存id, 期间, 性质)
              Values
                (v_Lngid, r_Stock.Id, d_开始日期, d_结束日期, Nvl(Zl_Username, 'zlhis'), Sysdate, n_结存id,
                 To_Char(Trunc(d_结束日期), 'yyyymm'), 1);
            
              --产生药品结存明细表，本期期末=本期期初(上期期末)+期间发生
              Insert Into 材料结存明细
                (结存id, 库房id, 材料id, 批次, 期初数量, 期初金额, 期初差价, 期末数量, 期末金额, 期末差价)
                Select v_Lngid, 库房id, 材料id, 批次, Sum(期初数量), Sum(期初金额), Sum(期初差价), Sum(期末数量), Sum(期末金额), Sum(期末差价)
                From (Select a.库房id, a.材料id, Nvl(a.批次, 0) As 批次, a.期末数量 As 期初数量, a.期末金额 As 期初金额, a.期末差价 As 期初差价, a.期末数量,
                              a.期末金额, a.期末差价
                       From 材料结存明细 A, 材料特性 B
                       Where a.材料id = b.材料id And a.结存id = n_结存id
                       Union All
                       Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, 0 As 期初数量, 0 As 期初金额, 0 As 期初差价,
                              a.入出系数 * a.实际数量 * Nvl(a.付数, 1) As 期末数量, a.入出系数 * a.零售金额 As 期末金额, a.入出系数 * a.差价 As 期末差价
                       From 药品收发记录 A, 材料特性 B
                       Where a.药品id = b.材料id And a.库房id + 0 = r_Stock.Id And a.审核日期 Between d_开始日期 And d_结束日期)
                Group By 库房id, 材料id, 批次
                Order By 库房id, 材料id, 批次;
            
              --计算误差：本期期末-库存记录(减去本期期末时间后发生的数据)
              Insert Into 材料结存误差
                (ID, 结存id, 库房id, 材料id, 批次, 数量差, 金额差, 差价差)
                Select 材料结存误差_Id.Nextval, v_Lngid, a.库房id, a.材料id, a.批次, a.实际数量 As 数量差, a.实际金额 As 金额差, a.实际差价 As 差价差
                From (Select 库房id, 材料id, 批次, Sum(实际数量) As 实际数量, Sum(实际金额) As 实际金额, Sum(实际差价) As 实际差价
                       From (Select a.库房id, a.材料id, Nvl(a.批次, 0) As 批次, Nvl(a.期末数量, 0) As 实际数量, Nvl(a.期末金额, 0) As 实际金额,
                                     Nvl(a.期末差价, 0) As 实际差价
                              From 材料结存明细 A, 材料特性 B
                              Where a.材料id = b.材料id And a.结存id = v_Lngid
                              Union All
                              Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, -1 * Nvl(a.实际数量, 0) As 实际数量,
                                     -1 * Nvl(a.实际金额, 0) As 实际金额, -1 * Nvl(实际差价, 0) As 实际差价
                              From 药品库存 A, 材料特性 B
                              Where a.药品id = b.材料id And a.性质 = 1 And a.库房id = r_Stock.Id
                              Union All
                              Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次, a.入出系数 * a.实际数量 * Nvl(a.付数, 1) As 实际数量,
                                     a.入出系数 * a.零售金额 As 实际金额, a.入出系数 * a.差价 As 实际差价
                              From 药品收发记录 A, 材料特性 B
                              Where a.药品id = b.材料id And a.库房id = r_Stock.Id And a.审核日期 > d_结束日期) A
                       Group By 库房id, 材料id, 批次) A
                Where a.实际数量 <> 0 Or a.实际金额 <> 0 Or a.实际差价 <> 0;
              --自动结存后立马审核结存信息
              Zl_材料结存记录_Verify(v_Lngid, Zl_Username);
            End If;
          End If;
        End If;
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Autostuffcloseaccount;
/

--127629:李业庆,2018-06-22,去掉过程中的commit
CREATE OR REPLACE Procedure Zl1_Autosend As
  v_库房id       药品收发记录.库房id%Type;
  v_自动发药天数 药房配药控制.自动发药天数%Type;

  Intdigit      Number(1);
  Intautoverify Number(1);
  Str操作员编号 人员表.编号%Type;
  Str操作员姓名 人员表.姓名%Type;

  Cursor Autosenddepid Is
    Select Nvl(药房id, 0) 药房id, 自动发药天数 From 药房配药控制 Where 门诊 = 2 And Nvl(自动发药天数, 0) > 0;

  Cursor Autosendlist Is
    Select Distinct A.库房id, A.ID, Nvl(A.批次, 0) 批次, C.操作员姓名,a.药品id
    From 药品收发记录 A, 未发药品记录 B, 住院费用记录 C
    Where A.单据 = B.单据 And A.NO = B.NO And A.费用id = C.ID And Nvl(A.库房id, v_库房id) + 0 = Nvl(B.库房id, v_库房id) And
          A.单据 In (9, 10) And Mod(A.记录状态, 3) = 1 And A.审核人 Is Null And Nvl(A.库房id, 0) + 0 = v_库房id And
          B.填制日期 < Sysdate - v_自动发药天数 order by a.药品id;

  v_Autosenddepid Autosenddepid%RowType;
  v_Autosendlist  Autosendlist%RowType;
Begin
  If f_Is_Primary_Node = 0 Then
    Return;
  End If;

  --取操作员编号与姓名
  Select 编号, 姓名
  Into Str操作员编号, Str操作员姓名
  From 人员表 A, 上机人员表 B
  Where A.ID = B.人员id And B.用户名 = User;

  --获取金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
  --判断划价单发药后是否自动审核为记帐单
  Select Zl_To_Number(zl_GetSysParameter(81)) Into Intautoverify From Dual;

  For v_Autosenddepid In Autosenddepid Loop
    v_库房id       := v_Autosenddepid.药房id;
    v_自动发药天数 := v_Autosenddepid.自动发药天数;
    If v_自动发药天数 > 30 Then
      v_自动发药天数 := 30;
    End If;
    For v_Autosendlist In Autosendlist Loop
      Zl_药品收发记录_部门发药(v_Autosendlist.库房id, v_Autosendlist.ID, v_Autosendlist.操作员姓名, Sysdate, v_Autosendlist.批次, 3, Null,
                     Null, Str操作员编号, Str操作员姓名, Intdigit, Intautoverify);
    End Loop;
  End Loop;
End Zl1_Autosend;
/

--127629:李业庆,2018-06-22,去掉过程中的commit
Create Or Replace Procedure Zl_材料差价重整_Update
(
  期间_In In 期间表.期间%Type
) Is
  Cursor c_期间表 Is
    Select 开始日期, 终止日期 From 期间表 Where 期间 >= 期间_In And Sysdate >= 开始日期;

  Cursor c_平均差价率
  (
    v_开始日期 Date,
    v_终止日期 Date
  ) Is
    Select d.药库, s.库房id, s.材料id, s.批次, Decode(Sign(s.金额), 1, s.差价 / s.金额, m.指导差价率 / 100) As 差价率
    From (Select o.库房id, o.材料id, o.批次, Nvl(e.当前金额, 0) - Nvl(j.发生金额, 0) - o.出库金额 As 金额,
                  Nvl(e.当前差价, 0) - Nvl(j.发生差价, 0) - o.出库差价 As 差价
           From (Select 库房id, 药品id 材料id, Nvl(批次, 0) As 批次, Sum(入出系数 * 零售金额) As 出库金额, Sum(入出系数 * 差价) As 出库差价
                  From 药品收发记录 L
                  Where 审核日期 Between Trunc(v_开始日期) And Trunc(v_终止日期) + 1 - 1 / 24 / 60 / 60 And
                        (单据 = 19 And Exists
                         (Select 1 From 部门性质说明 C Where c.部门id = l.库房id And c.工作性质 In ('卫材库', '虚拟库房')) And Not Exists
                         (Select 1
                          From 部门性质说明 C
                          Where c.部门id = l.对方部门id And c.工作性质 In ('卫材库', '制剂室', '虚拟库房')) Or 单据 Between 8 And 11 Or
                         单据 In (20, 21))
                  Group By 库房id, 药品id, Nvl(批次, 0)) O,
                (Select 库房id, 药品id 材料id, Nvl(批次, 0) As 批次, Sum(入出系数 * 零售金额) As 发生金额, Sum(入出系数 * 差价) As 发生差价
                  From 药品收发记录
                  Where 审核日期 >= Trunc(v_终止日期) + 1
                  Group By 库房id, 药品id, Nvl(批次, 0)) J,
                (Select 库房id, 药品id 材料id, Nvl(批次, 0) As 批次, Sum(实际金额) As 当前金额, Sum(实际差价) As 当前差价
                  From 药品库存
                  Where 性质 = 1
                  Group By 库房id, 药品id, Nvl(批次, 0)) E
           Where o.库房id = j.库房id(+) And o.材料id = j.材料id(+) And o.批次 = j.批次(+) And o.库房id = e.库房id(+) And
                 o.材料id = e.材料id(+) And o.批次 = e.批次(+)) S, 材料特性 M,
         (Select 部门id, Min(Decode(工作性质, '卫材库', 1, Decode(工作性质, '虚拟库房', 1, 2))) As 药库
           From 部门性质说明
           Where 工作性质 In ('卫材库', '虚拟库房', '发料部门')
           Group By 部门id) D
    Where s.材料id = m.材料id And s.库房id = d.部门id
    Order By d.药库, s.库房id, s.材料id, s.批次;

  Cursor c_材料出库记录
  (
    v_开始日期 Date,
    v_终止日期 Date,
    v_库房     Integer,
    v_库房id   Integer,
    v_材料id   Integer,
    v_批次     Integer
  ) Is
    Select ID, 单据, NO, 审核日期, 入出类别id, 入出系数, 成本价, 实际数量 * 付数 As 实际数量, 零售金额, 差价, 产地, 批号, 效期, 灭菌效期, 对方部门id, 生产日期, 批准文号,
           供药单位id
    From 药品收发记录 L
    Where 审核日期 Between Trunc(v_开始日期) And Trunc(v_终止日期) + 1 - 1 / 24 / 60 / 60 And 库房id = v_库房id And 药品id = v_材料id And
          Nvl(批次, 0) = Nvl(v_批次, 0) And
          (v_库房 = 1 And 单据 = 19 And Not Exists
           (Select 1 From 部门性质说明 C Where c.部门id = l.对方部门id And c.工作性质 In ('卫材库', '虚拟库房')) Or 单据 Between 8 And 11 Or
           单据 In (20, 21));
  v_原差价     Number(18, 2);
  v_现差价     Number(18, 2);
  v_成本价     Number(18, 4);
  v_对方类别id Integer;
  v_小数       Number(2);
Begin
  Select nvl(精度,2) into v_小数 From 药品卫材精度 Where 性质=0 and 类别 = 2 And 内容 = 4 And 单位 = 5;

  For v_Period In c_期间表 Loop
    For v_Avgtax In c_平均差价率(v_Period.开始日期, v_Period.终止日期) Loop
      For v_Outrec In c_材料出库记录(v_Period.开始日期, v_Period.终止日期, v_Avgtax.药库, v_Avgtax.库房id, v_Avgtax.材料id, v_Avgtax.批次) Loop
        v_原差价 := v_Outrec.差价;
        v_现差价 := Round(Nvl(v_Outrec.零售金额, 0) * v_Avgtax.差价率, v_小数);
        If Nvl(v_Outrec.实际数量, 0) = 0 Then
          v_成本价 := v_Outrec.成本价;
        Else
          v_成本价 := Round((Nvl(v_Outrec.零售金额, 0) - v_现差价) / v_Outrec.实际数量, 4);
        End If;

        Update 药品收发记录
        Set 差价 = Round(v_现差价, v_小数), 成本金额 = Round(Nvl(v_Outrec.零售金额, 0) - v_现差价, v_小数), 成本价 = v_成本价
        Where ID = v_Outrec.Id;

        Update 药品库存
        Set 实际差价 = Round(Nvl(实际差价, 0) + (v_现差价 - v_原差价) * v_Outrec.入出系数, v_小数)
        Where 库房id = v_Avgtax.库房id And 药品id = v_Avgtax.材料id And Nvl(批次, 0) = Nvl(v_Avgtax.批次, 0) And 性质 = 1;
        If Sql%NotFound Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号)
          Values
            (v_Avgtax.库房id, v_Avgtax.材料id, v_Avgtax.批次, 1, 0, 0, 0, Round((v_现差价 - v_原差价) * v_Outrec.入出系数, v_小数),
             v_Outrec.效期, v_Outrec.灭菌效期, v_Outrec.供药单位id, v_成本价, v_Outrec.批号, v_Outrec.生产日期, v_Outrec.产地, v_Outrec.批准文号);
        End If;

        If v_Outrec.单据 = 19 Then
          Update 药品收发记录
          Set 差价 = Round(v_现差价, v_小数), 成本金额 = Round(Nvl(v_Outrec.零售金额, 0), v_小数) - v_现差价, 成本价 = v_成本价
          Where NO = v_Outrec.No And 单据 = 19 And 药品id + 0 = v_Avgtax.材料id And Nvl(批次, 0) = Nvl(v_Avgtax.批次, 0) And
                库房id + 0 = v_Outrec.对方部门id And 对方部门id + 0 = v_Avgtax.库房id And 入出系数 = -1 * v_Outrec.入出系数;
          If Sql%NotFound Then
            Null;
          Else
            Select 入出类别id
            Into v_对方类别id
            From 药品收发记录
            Where NO = v_Outrec.No And 单据 = 19 And 药品id + 0 = v_Avgtax.材料id And Nvl(批次, 0) = Nvl(v_Avgtax.批次, 0) And
                  库房id + 0 = v_Outrec.对方部门id And 对方部门id + 0 = v_Avgtax.库房id And 入出系数 = -1 * v_Outrec.入出系数 And
                  Rownum < 2;

            Update 药品库存
            Set 实际差价 = Round(Nvl(实际差价, 0) + (v_现差价 - v_原差价) * v_Outrec.入出系数 * -1, v_小数)
            Where 库房id = v_Outrec.对方部门id And 药品id = v_Avgtax.材料id And Nvl(批次, 0) = Nvl(v_Avgtax.批次, 0) And 性质 = 1;
            If Sql%NotFound Then
              Insert Into 药品库存
                (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号)
              Values
                (v_Outrec.对方部门id, v_Avgtax.材料id, v_Avgtax.批次, 1, 0, 0, 0,
                 Round((v_现差价 - v_原差价) * v_Outrec.入出系数 * -1, v_小数), v_Outrec.效期, v_Outrec.灭菌效期, v_Outrec.供药单位id, v_成本价,
                 v_Outrec.批号, v_Outrec.生产日期, v_Outrec.产地, v_Outrec.批准文号);
            End If;
          End If;
        End If;
      End Loop;
      Delete From 药品库存
      Where 库房id = v_Avgtax.库房id And 药品id = v_Avgtax.材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
            Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
    End Loop;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料差价重整_Update;
/

--127629:李业庆,2018-06-22,去掉过程中的commit
Create Or Replace Procedure Zl_药品差价重整_全月平均
(
  开始时间_In In Date,
  库房id_In   In 药品收发记录.库房id%Type,
  结束时间_In In Date
) Is
  Cursor c_平均差价率
  (
    v_开始日期 Date,
    v_终止日期 Date
  ) Is
    Select d.药库, s.库房id, s.药品id, s.批次, Decode(Sign(s.金额), 1, s.差价 / s.金额, m.指导差价率 / 100) As 差价率
    From (Select o.库房id, o.药品id, o.批次, Nvl(e.当前金额, 0) - Nvl(j.发生金额, 0) - o.出库金额 As 金额,
                  Nvl(e.当前差价, 0) - Nvl(j.发生差价, 0) - o.出库差价 As 差价
           From (Select 库房id, 药品id, Nvl(批次, 0) As 批次, Sum(入出系数 * 零售金额) As 出库金额, Sum(入出系数 * 差价) As 出库差价
                  From 药品收发记录 L
                  Where 库房id = 库房id_In And 审核日期 Between Trunc(v_开始日期) And Trunc(v_终止日期) + 1 - 1 / 24 / 60 / 60 And
                        (单据 = 6 And Exists
                         (Select 1
                          From 部门性质说明 C
                          Where c.部门id = l.库房id And c.工作性质 In ('西药库', '中药库', '成药库')) And Not Exists
                         (Select 1
                          From 部门性质说明 C
                          Where c.部门id = l.对方部门id And c.工作性质 In ('西药库', '中药库', '成药库', '制剂室')) Or 单据 Between 7 And 11)
                  Group By 库房id, 药品id, Nvl(批次, 0)) O,
                (Select 库房id, 药品id, Nvl(批次, 0) As 批次, Sum(入出系数 * 零售金额) As 发生金额, Sum(入出系数 * 差价) As 发生差价
                  From 药品收发记录
                  Where 库房id = 库房id_In And 审核日期 >= Trunc(v_终止日期) + 1
                  Group By 库房id, 药品id, Nvl(批次, 0)) J,
                (Select 库房id, 药品id, Nvl(批次, 0) As 批次, Sum(实际金额) As 当前金额, Sum(实际差价) As 当前差价
                  From 药品库存
                  Where 性质 = 1
                  Group By 库房id, 药品id, Nvl(批次, 0)) E
           Where o.库房id = j.库房id(+) And o.药品id = j.药品id(+) And o.批次 = j.批次(+) And o.库房id = e.库房id(+) And
                 o.药品id = e.药品id(+) And o.批次 = e.批次(+)) S, 药品规格 M,
         (Select 部门id, Min(Decode(工作性质, '西药库', 1, '中药库', 1, '成药库', 1, 2)) As 药库
           From 部门性质说明
           Where 工作性质 In ('西药库', '中药库', '成药库', '西药房', '中药房', '成药房')
           Group By 部门id) D
    Where s.药品id = m.药品id And s.库房id = d.部门id
    Order By d.药库, s.库房id, s.药品id, s.批次;

  Cursor c_药品出库记录
  (
    v_开始日期 Date,
    v_终止日期 Date,
    v_药库     Integer,
    v_库房id   Integer,
    v_药品id   Integer,
    v_批次     Integer
  ) Is
    Select ID, 单据, NO, 审核日期, 入出类别id, 入出系数, 成本价, 实际数量 * 付数 As 实际数量, 零售金额, 差价, 产地, 批号, 效期, 对方部门id
    From 药品收发记录 L
    Where 库房id = 库房id_In And 审核日期 Between Trunc(v_开始日期) And Trunc(v_终止日期) + 1 - 1 / 24 / 60 / 60 And 库房id = v_库房id And
          药品id = v_药品id And Nvl(批次, 0) = Nvl(v_批次, 0) And
          (v_药库 = 1 And 单据 = 6 And Not Exists
           (Select 1
            From 部门性质说明 C
            Where c.部门id = l.对方部门id And c.工作性质 In ('西药库', '中药库', '成药库', '制剂室')) Or 单据 Between 7 And 11);

  v_原差价     药品库存.实际差价%Type;
  v_现差价     药品库存.实际差价%Type;
  v_成本价     药品库存.上次采购价%Type;
  v_对方类别id Integer;
  Intdigit     Number;
Begin
  --获取金额小数位数
  Select Nvl(精度, 2) Into Intdigit From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;

  For v_Avgtax In c_平均差价率(开始时间_In, 结束时间_In) Loop
    For v_Outrec In c_药品出库记录(开始时间_In, 结束时间_In, v_Avgtax.药库, v_Avgtax.库房id, v_Avgtax.药品id, v_Avgtax.批次) Loop
      v_原差价 := v_Outrec.差价;
      v_现差价 := Round(Nvl(v_Outrec.零售金额, 0) * v_Avgtax.差价率, Intdigit);
      If Nvl(v_Outrec.实际数量, 0) = 0 Then
        v_成本价 := v_Outrec.成本价;
      Else
        v_成本价 := Round((Nvl(v_Outrec.零售金额, 0) - v_现差价) / v_Outrec.实际数量, 7);
      End If;
    
      Update 药品收发记录
      Set 差价 = v_现差价, 成本金额 = Nvl(v_Outrec.零售金额, 0) - v_现差价, 成本价 = v_成本价
      Where ID = v_Outrec.Id;
    
      Update 药品库存
      Set 实际差价 = Nvl(实际差价, 0) + (v_现差价 - v_原差价) * v_Outrec.入出系数
      Where 库房id = v_Avgtax.库房id And 药品id = v_Avgtax.药品id And Nvl(批次, 0) = Nvl(v_Avgtax.批次, 0) And 性质 = 1;
      If Sql%NotFound Then
        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期)
        Values
          (v_Avgtax.库房id, v_Avgtax.药品id, v_Avgtax.批次, 1, 0, 0, 0, (v_现差价 - v_原差价) * v_Outrec.入出系数, Null, v_成本价,
           v_Outrec.批号, v_Outrec.产地, v_Outrec.效期);
      End If;
    
      If v_Outrec.单据 = 6 Then
        Update 药品收发记录
        Set 差价 = v_现差价, 成本金额 = Nvl(v_Outrec.零售金额, 0) - v_现差价, 成本价 = v_成本价
        Where NO = v_Outrec.No And 单据 = 6 And 药品id + 0 = v_Avgtax.药品id And Nvl(批次, 0) = Nvl(v_Avgtax.批次, 0) And
              库房id + 0 = v_Outrec.对方部门id And 对方部门id + 0 = v_Avgtax.库房id And 入出系数 = -1 * v_Outrec.入出系数;
        If Sql%NotFound Then
          Null;
        Else
          Select 入出类别id
          Into v_对方类别id
          From 药品收发记录
          Where NO = v_Outrec.No And 单据 = 6 And 药品id + 0 = v_Avgtax.药品id And Nvl(批次, 0) = Nvl(v_Avgtax.批次, 0) And
                库房id + 0 = v_Outrec.对方部门id And 对方部门id + 0 = v_Avgtax.库房id And 入出系数 = -1 * v_Outrec.入出系数 And Rownum < 2;
        
          Update 药品库存
          Set 实际差价 = Nvl(实际差价, 0) + (v_现差价 - v_原差价) * v_Outrec.入出系数 * -1
          Where 库房id = v_Outrec.对方部门id And 药品id = v_Avgtax.药品id And Nvl(批次, 0) = Nvl(v_Avgtax.批次, 0) And 性质 = 1;
          If Sql%NotFound Then
            Insert Into 药品库存
              (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期)
            Values
              (v_Outrec.对方部门id, v_Avgtax.药品id, v_Avgtax.批次, 1, 0, 0, 0, (v_现差价 - v_原差价) * v_Outrec.入出系数 * -1, Null,
               v_成本价, v_Outrec.批号, v_Outrec.产地, v_Outrec.效期);
          End If;
        End If;
      End If;
    End Loop;
    Delete From 药品库存
    Where 库房id = v_Avgtax.库房id And 药品id = v_Avgtax.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品差价重整_全月平均;
/
--111037:余伟节,2018-06-24,新生儿登记允许录入死亡时间
Create Or Replace Procedure Zl_病人新生儿记录_Insert
(
  病人id_In   病人新生儿记录.病人id%Type,
  主页id_In   病人新生儿记录.主页id%Type,
  序号_In     病人新生儿记录.序号%Type,
  婴儿姓名_In 病人新生儿记录.婴儿姓名%Type,
  婴儿性别_In 病人新生儿记录.婴儿性别%Type,
  分娩次数_In 病人新生儿记录.分娩次数%Type,
  分娩方式_In 病人新生儿记录.分娩方式%Type,
  胎儿状况_In 病人新生儿记录.胎儿状况%Type,
  出生时间_In 病人新生儿记录.出生时间%Type,
  身长_In     病人新生儿记录.身长%Type,
  体重_In     病人新生儿记录.体重%Type,
  血型_In     病人新生儿记录.血型%Type,
  备注说明_In 病人新生儿记录.备注说明%Type := Null,
  死亡时间_In 病人新生儿记录.死亡时间%Type := Null
) Is
Begin
  Insert Into 病人新生儿记录
    (病人id, 主页id, 序号, 婴儿姓名, 婴儿性别, 分娩次数, 分娩方式, 胎儿状况, 身长, 体重, 血型, 出生时间, 死亡时间, 备注说明)
  Values
    (病人id_In, 主页id_In, 序号_In, 婴儿姓名_In, 婴儿性别_In, 分娩次数_In, 分娩方式_In, 胎儿状况_In, 身长_In, 体重_In, 血型_In, 出生时间_In, 死亡时间_In,
     备注说明_In);

  Zl_病区自动标记_Update(病人id_In, 主页id_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人新生儿记录_Insert;
/

--111037:余伟节,2018-06-24,新生儿登记允许录入死亡时间

Create Or Replace Procedure Zl_病人医嘱记录_作废
(
  Id_In         病人医嘱记录.Id%Type,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null,
  护理医嘱id_In 病人医嘱记录.Id%Type := Null,
  作废时间_In   病人医嘱状态.操作时间%Type := Null
) Is
  --功能：作废指定的医嘱(未发送的长嘱或临嘱)
  --说明：一并给药的只能调用一次(界面显示有多行)
  --参数：ID_IN=组医嘱ID
  --      护理医嘱id_In 取除开本次作废的护理等级医嘱外的最近的自动停止的护理等级医嘱id

  v_发送号       病人医嘱发送.发送号%Type;
  v_费用no       门诊费用记录.No%Type;
  v_记录性质     门诊费用记录.记录性质%Type;
  v_费用序号     Varchar2(255);
  n_自动取消执行 Number(1) := 0;
  n_先作废后退药 Number(1) := 0;

  v_Date     Date;
  v_Count    Number;
  v_Temp     Varchar2(255);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;

  --包含医嘱相关信息
  Cursor c_Advice Is
    Select a.病人id, a.挂号单, a.主页id, a.婴儿, a.医嘱状态, a.上次执行时间, a.医嘱内容, a.诊疗类别, b.操作类型, a.病人来源, a.执行科室id, b.执行频率, a.诊疗项目id,
           a.开始执行时间
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id And a.Id = Id_In;

  r_Advice c_Advice%RowType;

  --门诊医嘱作废时，取对应的费用销帐或作废(收费划价单)：
  --根据医嘱及发送NO求出本次回退要销帐或退费的记录
  --一组医嘱并不是都填写了发送记录,也不一定都计费了,且可能NO不同
  --只管记录状态为1的记录,如果已经销帐或部份销帐的记录,不再处理
  --费用只求价格父号为空的,以便取序号销帐
  --如果"门诊药嘱先作废后退药",则不对相应费用(包括给药途径的)进行检查和处理,除非是还没有执行的记帐单,或未执行、收费的划价单，可以先删了


  Cursor c_Rollmoney(v_发送号 病人医嘱发送.发送号%Type) Is
    Select Decode(a.记录性质, 11, 1, a.记录性质) As 记录性质, a.记录状态, a.No, a.序号, a.执行状态 As 费用执行, c.执行状态 As 医嘱执行, c.执行部门id, b.病人科室id,
           b.诊疗类别, i.操作类型
    From 门诊费用记录 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 I
    Where c.医嘱id = b.Id And c.发送号 = v_发送号 And (b.Id = Id_In Or b.相关id = Id_In) And a.医嘱序号 = b.Id And a.记录状态 In (0, 1) And
          a.No = c.No And (a.记录性质 = c.记录性质 Or a.记录性质 = 11 And c.记录性质 = 1) And b.诊疗项目id = i.Id And a.价格父号 Is Null And
          (n_先作废后退药 = 0 Or
          n_先作废后退药 = 1 And
          Not (Exists (Select 1
                        From 门诊费用记录 D
                        Where d.医嘱序号 = b.Id And d.记录状态 In (0, 1) And d.No = c.No And
                              (d.记录性质 = c.记录性质 Or d.记录性质 = 11 And c.记录性质 = 1) And d.收费类别 In ('5', '6', '7'))) Or
          Nvl(a.执行状态, 0) = 0 And Not (a.记录性质 = 1 And a.记录状态 <> 0))
    Order By a.记录性质, a.No, a.序号, a.收费细目id;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --检查医嘱状态是否正确:并发操作
  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;

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

  --检查是否是输液配液记录，并是否已经锁定
  Select Count(1) Into v_Count From 输液配药记录 Where 是否锁定 = 1 And 医嘱id = Id_In;
  If v_Count > 0 Then
    v_Error := '医嘱"' || r_Advice.医嘱内容 || '"是输液药品，已经被输液配置中心锁定，不能作废。';
    Raise Err_Custom;
  End If;

  If r_Advice.挂号单 Is Null And r_Advice.病人来源 <> 3 Then
    If r_Advice.医嘱状态 In (4, 8, 9) Then
      v_Error := '医嘱"' || r_Advice.医嘱内容 || '"已经被作废或停止，不能再作废。';
      Raise Err_Custom;
    Elsif r_Advice.上次执行时间 Is Not Null Then
      v_Error := '医嘱"' || r_Advice.医嘱内容 || '"已经发送，不能被作废。';
      Raise Err_Custom;
    End If;
  
    --持续性护理等级无须发送，校对后就可能已自动计费，作废及回退作废都应按停止流程处理。
    If r_Advice.诊疗类别 = 'H' And r_Advice.操作类型 = '1' And r_Advice.执行频率 = '2' And Nvl(r_Advice.婴儿, 0) = 0 Then
      --(已取消，由于存在无费退院的情况，问题号：45977)a.开始时间是当天之前的，说明已生效（自动费用计算），不允许作废。
      --医嘱的时间只精确到了分钟，所以变动记录的开始时间要去掉秒来比较。
      v_Count := 0;
      Begin
        Select b.终止时间
        Into v_Date
        From 病人变动记录 B, 病人医嘱计价 C
        Where b.病人id = r_Advice.病人id And b.主页id = r_Advice.主页id And c.医嘱id = Id_In And c.收费细目id = b.护理等级id And
              b.开始原因 = 6 And b.附加床位 = 0 And
              To_Char(b.开始时间, 'yyyy-mm-dd hh24:mi') = To_Char(r_Advice.开始执行时间, 'yyyy-mm-dd hh24:mi');
      Exception
        When Others Then
          v_Count := 1;
      End;
      If v_Count = 0 Then
        --d.后续有其他变动发生
        If v_Date Is Not Null Then
          v_Error := '由于护理等级医嘱生效后已经产生了其他变动记录,不能作废该医嘱。';
          Raise Err_Custom;
        Else
          --本次有要自动启用的护理等级，如果和原来护理等级相同则不用撤消护理变动记录
          If Nvl(护理医嘱id_In, 0) <> 0 Then
            Delete 病人医嘱状态 Where 医嘱id = 护理医嘱id_In And 操作类型 In (8, 9);
            Select 操作类型
            Into v_Count
            From (Select 操作类型 From 病人医嘱状态 Where 医嘱id = 护理医嘱id_In Order By 操作时间 Desc)
            Where Rownum < 2;
            Update 病人医嘱记录
            Set 医嘱状态 = v_Count, 执行终止时间 = Null, 停嘱医生 = Null, 停嘱时间 = Null, 确认停嘱时间 = Null, 确认停嘱护士 = Null
            Where ID = 护理医嘱id_In;
            --排除过于频繁的操作
            Select Count(a.Id)
            Into v_Count
            From 病人医嘱记录 A, 诊疗收费关系 B, 病案主页 C
            Where a.诊疗项目id = b.诊疗项目id And c.护理等级id = b.收费项目id And c.病人id = a.病人id And c.主页id = a.主页id And
                  a.Id = 护理医嘱id_In;
          End If;
          If v_Count = 0 Then
            --c.护理等级是最后一条变动
            Zl_病人变动记录_Undo(r_Advice.病人id, r_Advice.主页id, v_人员编号, v_人员姓名, '1', Null, Null, '护理等级变动');
          End If;
        End If;
      Else
        --恢复最近一次被自动停止的护理等级
        If Nvl(护理医嘱id_In, 0) <> 0 Then
          Delete 病人医嘱状态 Where 医嘱id = 护理医嘱id_In And 操作类型 In (8, 9);
          Select 操作类型
          Into v_Count
          From (Select 操作类型 From 病人医嘱状态 Where 医嘱id = 护理医嘱id_In Order By 操作时间 Desc)
          Where Rownum < 2;
          Update 病人医嘱记录
          Set 医嘱状态 = v_Count, 执行终止时间 = Null, 停嘱医生 = Null, 停嘱时间 = Null, 确认停嘱时间 = Null, 确认停嘱护士 = Null
          Where ID = 护理医嘱id_In;
        Else
          --病人入院时指定的护理级产生的变动记录和医嘱新开产生的变动记录不同，这里要先判断
          Select Count(a.Id)
          Into v_Count
          From 病人变动记录 A
          Where a.病人id = r_Advice.病人id And a.主页id = r_Advice.主页id And a.开始原因 = 6;
          If v_Count <> 0 Then
            --b.如果与以前的护理等级相同，则校对时没有产生护理等级变动,产生护理等级停止变动
            Zl_病人变动记录_Nurse(r_Advice.病人id, r_Advice.主页id, Null, Sysdate, v_人员编号, v_人员姓名);
          End If;
        End If;
      End If;
    End If;
  Else
    If r_Advice.医嘱状态 <> 8 Then
      v_Error := '医嘱"' || r_Advice.医嘱内容 || '"尚未发送或已经作废。';
      Raise Err_Custom;
    End If;
    --医嘱附费判断
    Select Count(1)
    Into v_Count
    From 病人医嘱附费 A, 病人医嘱记录 B
    Where a.医嘱id = b.Id And (b.Id = Id_In Or b.相关id = Id_In);
    If v_Count <> 0 Then
      v_Error := '医嘱"' || r_Advice.医嘱内容 || '"存在附加费用，不能作废。';
      Raise Err_Custom;
    End If;
  
    Begin
      --医嘱ID为传入值的这条医嘱不一定发送了的,甚至无发送。
      Select Distinct 发送号
      Into v_发送号
      From 病人医嘱发送
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = Id_In Or 相关id = Id_In);
    Exception
      When Others Then
        v_发送号 := Null;
    End;
  
    Select Zl_To_Number(Nvl(zl_GetSysParameter(68), 0)) Into n_先作废后退药 From Dual;
    Select Zl_To_Number(Nvl(zl_GetSysParameter('门诊本科自动执行', '1252'), 0)) Into n_自动取消执行 From Dual;
    If n_自动取消执行 = 1 And v_发送号 Is Not Null Then
      --先更新医嘱和费用的执行状态，因为后续的判断，以及过程Zl_门诊记帐记录_Delete中有检查
      For Rc In (Select a.医嘱id, a.执行部门id
                 From 病人医嘱发送 A, 病人医嘱记录 B
                 Where a.医嘱id = b.Id And (b.Id = Id_In Or b.相关id = Id_In) And a.执行部门id = b.病人科室id) Loop
        Zl_病人医嘱执行_Cancel(Rc.医嘱id, v_发送号, Null, 1, Rc.执行部门id);
      End Loop;
    End If;
  
    --门诊医嘱只可能发送一次
    --后面退费时还有检查，因为可能医嘱没有费用，所以要检查一次执行状态
    Select Count(*)
    Into v_Count
    From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 I
    Where a.医嘱id = b.Id And b.诊疗项目id = i.Id And a.执行状态 In (1, 3) And (b.Id = Id_In Or b.相关id = Id_In) And
          (n_先作废后退药 = 0 Or
          n_先作废后退药 = 1 And Not (b.诊疗类别 In ('5', '6', '7') Or b.诊疗类别 = 'E' And i.操作类型 In ('2', '3', '4')));
    If v_Count > 0 Then
      v_Error := '该医嘱已经执行或正在执行，不能作废。';
      Raise Err_Custom;
    End If;
  End If;

  If 作废时间_In Is Null Then
    Select Sysdate Into v_Date From Dual;
  Else
    v_Date := 作废时间_In;
  End If;

  Update 病人医嘱记录 Set 医嘱状态 = 4 Where ID = Id_In Or 相关id = Id_In;

  Insert Into 病人医嘱状态
    (医嘱id, 操作类型, 操作人员, 操作时间)
    Select ID, 4, v_人员姓名, v_Date From 病人医嘱记录 Where ID = Id_In Or 相关id = Id_In;

  --住院医嘱作废时,未打印的情况下,缺省设置为屏蔽打印
  If r_Advice.挂号单 Is Null Then
    Select Count(*)
    Into v_Count
    From 病人医嘱打印
    Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = Id_In Or 相关id = Id_In);
    If Nvl(v_Count, 0) = 0 Then
      Zl_病人医嘱记录_屏蔽打印(Id_In, 1);
    End If;
    If Nvl(r_Advice.婴儿, 0) > 0 And r_Advice.操作类型 = '11' Then
      Update 病人新生儿记录
      Set 死亡时间 = Null
      Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = Nvl(r_Advice.主页id, 0) And 序号 = Nvl(r_Advice.婴儿, 0);
    End If;
  Else
    --门诊医嘱(临嘱)作废时还需要回退相关内容:只有一次发送
    --回退划价或记帐费用
    If v_发送号 Is Not Null Then
      --将该组医嘱的费用删除或销帐(按一组医嘱可能有不同NO处理)
      --门诊记帐：如果原始费用已被销帐(或部分销帐),调用过程中有判断
      --门诊划价：如果已收费，则不允许删除
      v_费用no   := Null;
      v_费用序号 := Null;
      For r_Rollmoney In c_Rollmoney(v_发送号) Loop
        If Nvl(r_Rollmoney.医嘱执行, 0) In (1, 3) Then
          --1-完全执行;3-正在执行
          v_Error := '医嘱"' || r_Advice.医嘱内容 || '"已经执行或正在执行，不能作废。';
          Raise Err_Custom;
        End If;
        If Nvl(r_Rollmoney.费用执行, 0) In (1, 2) Then
          --1-完全执行;2-部份执行
          v_Error := '医嘱费用单据"' || r_Rollmoney.No || '"中的内容已经全部或部分执行，不能作废。';
          Raise Err_Custom;
        End If;
        If r_Rollmoney.费用执行 = 9 Then
          v_Error := '医嘱费用单据"' || r_Rollmoney.No || '"中的收费结算产生异常，不能作废。';
          Raise Err_Custom;
        End If;
        v_Count := 1;
        If r_Rollmoney.记录性质 = 1 And r_Rollmoney.记录状态 <> 0 Then
          If 1 = n_先作废后退药 And r_Rollmoney.诊疗类别 = 'E' And r_Rollmoney.操作类型 In ('2', '3', '4') Then
            v_Count := 0;
          Else
            v_Error := '医嘱费用单据"' || r_Rollmoney.No || '"已经收费，不能作废。';
            Raise Err_Custom;
          End If;
        End If;
        If 1 = v_Count Then
          If Nvl(v_费用no, '空') <> r_Rollmoney.No Then
            If v_费用序号 Is Not Null And v_费用no Is Not Null Then
              v_费用序号 := Substr(v_费用序号, 2);
              If v_记录性质 = 1 Then
                Zl_门诊划价记录_Delete(v_费用no, v_费用序号);
              Elsif v_记录性质 = 2 Then
                Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
              End If;
            End If;
            v_费用序号 := Null;
          End If;
          v_记录性质 := r_Rollmoney.记录性质;
          v_费用no   := r_Rollmoney.No;
          v_费用序号 := v_费用序号 || ',' || r_Rollmoney.序号;
        End If;
      End Loop;
      If v_费用序号 Is Not Null And v_费用no Is Not Null Then
        v_费用序号 := Substr(v_费用序号, 2);
        If v_记录性质 = 1 Then
          Zl_门诊划价记录_Delete(v_费用no, v_费用序号);
        Elsif v_记录性质 = 2 Then
          Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
        End If;
      End If;
    
      --如果"门诊药嘱先作废后退药"，则对应的给药途径费用设置为未执行，以便退费
      If n_先作废后退药 = 1 Then
        Update 门诊费用记录
        Set 执行状态 = 0
        Where 执行状态 = 1 And 医嘱序号 = Id_In And Exists
         (Select 1
               From 病人医嘱记录 A, 诊疗项目目录 B
               Where a.诊疗项目id = b.Id And b.类别 = 'E' And b.操作类型 In ('2', '3', '4') And a.Id = Id_In);
      End If;
    
      --回退医嘱发送记录(及执行记录)
      Delete From 病人医嘱执行 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = Id_In Or 相关id = Id_In);
      Delete From 病人医嘱发送 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = Id_In Or 相关id = Id_In);
    
      --回退特殊医嘱的处理
      If r_Advice.诊疗类别 = 'Z' And Nvl(r_Advice.操作类型, '0') <> '0' And Nvl(r_Advice.婴儿, 0) = 0 Then
        If r_Advice.操作类型 = '1' And r_Advice.执行科室id Is Not Null Then
          --留观医嘱
          Select Count(*)
          Into v_Count
          From 病案主页
          Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = 0 And 入院科室id = r_Advice.执行科室id And 病人性质 In (1, 2);
          If v_Count = 1 Then
            Zl_入院病案主页_Delete(r_Advice.病人id, 0);
          End If;
        Elsif r_Advice.操作类型 = '2' And r_Advice.执行科室id Is Not Null Then
          --住院医嘱
          Select Count(*)
          Into v_Count
          From 病案主页
          Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = 0 And 入院科室id = r_Advice.执行科室id And Nvl(病人性质, 0) = 0;
          If v_Count = 1 Then
            Zl_入院病案主页_Delete(r_Advice.病人id, 0);
          End If;
        End If;
      End If;
    End If;
  End If;

  --删除过敏登记记录
  If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '1' Then
    --Update 病人医嘱记录 Set 皮试结果=Null Where ID=ID_IN; --保留最后的皮试结果
    --删除不过敏的记录，过敏记录保留，因为不管医嘱是否作废，病人对该药过敏
    For r_Test In (Select 操作时间 From 病人医嘱状态 Where 医嘱id = Id_In And 操作类型 = 10) Loop
      Delete From 病人过敏记录
      Where 病人id = r_Advice.病人id And 记录来源 = 2 And Nvl(主页id, 0) = Nvl(r_Advice.主页id, 0) And 记录时间 = r_Test.操作时间;
    End Loop;
  End If;

  Close c_Advice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_作废;
/

--115765:刘涛,2018-06-28,成本价精度处理
--127911,李业庆,2018-06-28,高值卫材取虚拟库房的成本价
Create Or Replace Procedure Zl_材料收发记录_处方发料
(
  Partid_In   In 药品收发记录.库房id%Type,
  Bill_In     In 药品收发记录.单据%Type,
  No_In       In 药品收发记录.No%Type,
  People_In   In 药品收发记录.审核人%Type,
  配药人_In   In 药品收发记录.配药人%Type := Null,
  校验人_In   In 药品收发记录.填制人%Type := Null,
  发药方式_In In 药品收发记录.发药方式%Type := 1,
  发药时间_In In 药品收发记录.审核日期%Type := Null
) Is
  --重新计算用
  Cursor c_Modifybill Is
    Select a.Id, a.药品id, a.入出类别id, a.入出系数, a.费用id, a.供药单位id, a.生产日期, a.批准文号, a.灭菌效期, a.效期, a.产地,
           Nvl(a.实际数量, 0) * Nvl(a.付数, 1) 数量, Nvl(a.零售金额, 0) 金额, Nvl(a.批次, 0) 批次, a.批号, 2 As 病人来源, a.库房id, a.内部条码, a.商品条码
    From 药品收发记录 A, 住院费用记录 B
    Where a.No = No_In And a.单据 = Bill_In And (a.库房id + 0 = Partid_In Or a.库房id Is Null) And Nvl(a.摘要, '拒发否') <> '拒发' And
          a.费用id = b.Id And b.执行状态 <> 1 And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null
    Union All
    Select a.Id, a.药品id, a.入出类别id, a.入出系数, a.费用id, a.供药单位id, a.生产日期, a.批准文号, a.灭菌效期, a.效期, a.产地,
           Nvl(a.实际数量, 0) * Nvl(a.付数, 1) 数量, Nvl(a.零售金额, 0) 金额, Nvl(a.批次, 0) 批次, a.批号, 1 As 病人来源, a.库房id, a.内部条码, a.商品条码
    From 药品收发记录 A, 门诊费用记录 B
    Where a.No = No_In And a.单据 = Bill_In And (a.库房id + 0 = Partid_In Or a.库房id Is Null) And Nvl(a.摘要, '拒发否') <> '拒发' And
          a.费用id = b.Id And b.执行状态 <> 1 And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null
    Order By 药品id;

  v_Modifybill c_Modifybill%RowType;

  --只读变量
  n_库存金额 药品库存.实际金额%Type;
  n_库存差价 药品库存.实际差价%Type;
  n_差价率   材料特性.指导差价率%Type;

  --可写变量
  n_成本金额       药品收发记录.成本金额%Type;
  n_成本价         药品收发记录.成本价%Type;
  n_实际差价       药品收发记录.差价%Type;
  d_操作时间       药品收发记录.审核日期%Type;
  n_收费与发料分离 Number(1);
  n_小数           Number(1);
  n_成本价小数           Number(1);
  v_入库no         药品收发记录.No%Type;
  v_入库库房id     药品收发记录.库房id%Type := 0;
  v_病人信息       Varchar2(200);
  n_虚拟库房       药品库存.库房id%Type;
  n_序号           Number;
  n_平均成本价     药品库存.平均成本价%Type;
Begin
  If 发药时间_In Is Null Then
    Select Sysdate Into d_操作时间 From Dual;
  Else
    d_操作时间 := 发药时间_In;
  End If;

  Begin
    Select 0 Into n_收费与发料分离 From 未发药品记录 Where 单据 = Bill_In And NO = No_In And 库房id + 0 = Partid_In;
  Exception
    When Others Then
      n_收费与发料分离 := 1;
  End;

  --获取金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_小数 From Dual;
  --获取成本价小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(157), '2')) Into n_成本价小数 From Dual;

  --重写已发料处方的配药人
  Update 药品收发记录
  Set 配药人 = Decode(配药人_In, Null, 配药人, 配药人_In)
  Where NO = No_In And 单据 = Bill_In And (库房id + 0 = Partid_In Or 库房id Is Null) And Mod(记录状态, 3) = 1 And 审核人 Is Not Null;

  --重新计算成本价、成本金额、零售金额及差价
  For v_Modifybill In c_Modifybill Loop
    --高值卫材虚拟出库模式
    If n_虚拟库房 Is Null Then
      Begin
        Select 库房id
        Into n_虚拟库房
        From 药品收发记录
        Where 单据 = 21 And 审核日期 Is Null And 药品id = v_Modifybill.药品id And Nvl(批次, 0) = v_Modifybill.批次 And
              费用id = v_Modifybill.费用id And Rownum = 1;
      Exception
        When Others Then
          n_虚拟库房 := 0;
      End;
    End If;
  
    If n_虚拟库房 = 0 Then
      --普通模式取发料部门价格
      n_成本价 := Round(Zl_Fun_Getoutcost(v_Modifybill.药品id, v_Modifybill.批次, Partid_In), n_成本价小数);
    Else
      --高值卫材虚拟出库模式取虚拟库房价格
      n_成本价 := Round(Zl_Fun_Getoutcost(v_Modifybill.药品id, v_Modifybill.批次, n_虚拟库房), n_成本价小数);
    End If;
    n_成本金额 := Round(n_成本价 * v_Modifybill.数量, n_小数);
    n_实际差价 := Round(Nvl(v_Modifybill.金额, 0) - n_成本金额, n_小数);
  
    --更新药品收发记录的零售金额、成本金额及差价
    Update 药品收发记录 Set 成本价 = n_成本价, 成本金额 = n_成本金额, 差价 = n_实际差价 Where ID = v_Modifybill.Id;
  
    If n_收费与发料分离 = 1 Then
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) - Nvl(v_Modifybill.数量, 0), 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybill.数量, 0),
          实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybill.金额, 0), 实际差价 = Nvl(实际差价, 0) - n_实际差价,
          上次采购价 = Decode(上次采购价, Null, n_成本价, 上次采购价), 平均成本价 = Decode(平均成本价, Null, n_成本价, 平均成本价),
          商品条码 = Decode(商品条码, Null, v_Modifybill.商品条码, 商品条码), 内部条码 = Decode(内部条码, Null, v_Modifybill.内部条码, 内部条码),
          效期 = Decode(效期, Null, v_Modifybill.效期, 效期), 上次批号 = Decode(上次批号, Null, v_Modifybill.批号, 上次批号),
          上次生产日期 = Decode(上次生产日期, Null, v_Modifybill.生产日期, 上次生产日期), 上次产地 = Decode(上次产地, Null, v_Modifybill.产地, 上次产地)
      Where 库房id + 0 = Partid_In And 药品id = v_Modifybill.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybill.批次;
    Else
      Update 药品库存
      Set 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybill.数量, 0), 实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybill.金额, 0),
          实际差价 = Nvl(实际差价, 0) - n_实际差价, 上次采购价 = Decode(上次采购价, Null, n_成本价, 上次采购价),
          平均成本价 = Decode(平均成本价, Null, n_成本价, 平均成本价), 商品条码 = Decode(商品条码, Null, v_Modifybill.商品条码, 商品条码),
          内部条码 = Decode(内部条码, Null, v_Modifybill.内部条码, 内部条码), 效期 = Decode(效期, Null, v_Modifybill.效期, 效期),
          上次批号 = Decode(上次批号, Null, v_Modifybill.批号, 上次批号), 上次生产日期 = Decode(上次生产日期, Null, v_Modifybill.生产日期, 上次生产日期),
          上次产地 = Decode(上次产地, Null, v_Modifybill.产地, 上次产地)
      Where 库房id + 0 = Partid_In And 药品id = v_Modifybill.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybill.批次;
    End If;
  
    If Sql%RowCount = 0 Then
      If n_收费与发料分离 = 1 Then
        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 平均成本价)
        Values
          (Partid_In, v_Modifybill.药品id, v_Modifybill.批次, 1, 0 - Nvl(v_Modifybill.数量, 0), 0 - Nvl(v_Modifybill.金额, 0),
           0 - n_实际差价, v_Modifybill.效期, v_Modifybill.灭菌效期, v_Modifybill.供药单位id, n_成本价, v_Modifybill.批号,
           v_Modifybill.生产日期, v_Modifybill.产地, v_Modifybill.批准文号, n_成本价);
      Else
        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 平均成本价)
        Values
          (Partid_In, v_Modifybill.药品id, v_Modifybill.批次, 1, 0 - Nvl(v_Modifybill.数量, 0), 0 - Nvl(v_Modifybill.数量, 0),
           0 - Nvl(v_Modifybill.金额, 0), 0 - n_实际差价, v_Modifybill.效期, v_Modifybill.灭菌效期, v_Modifybill.供药单位id, n_成本价,
           v_Modifybill.批号, v_Modifybill.生产日期, v_Modifybill.产地, v_Modifybill.批准文号, n_成本价);
      End If;
    End If;
  
    Delete 药品库存
    Where 库房id + 0 = Partid_In And 药品id = v_Modifybill.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
          Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0 And 性质 = 1;
  
    --更新病人费用记录的执行状态(已执行)
    If v_Modifybill.病人来源 = 2 Then
      Update 住院费用记录
      Set 执行状态 = 1, 执行人 = Decode(People_In, Null, Zl_Username, People_In), 执行时间 = 发药时间_In
      Where ID = v_Modifybill.费用id;
    Else
      Update 门诊费用记录
      Set 执行状态 = 1, 执行人 = Decode(People_In, Null, Zl_Username, People_In), 执行时间 = 发药时间_In
      Where ID = v_Modifybill.费用id;
    End If;
    --写审核人
    Update 药品收发记录
    Set 库房id = Partid_In, 配药人 = Decode(配药人_In, Null, 配药人, 配药人_In), 填制人 = Decode(校验人_In, Null, 填制人, 校验人_In),
        审核人 = People_In, 审核日期 = d_操作时间, 发药方式 = 发药方式_In
    Where ID = v_Modifybill.Id;
    --修正误差数据
    Zl_材料收发记录_调价修正(v_Modifybill.Id);
  
    If n_虚拟库房 > 0 Then
      --审核备货卫材在虚拟库房的其他出库单据
      For v_出库 In (Select 序号, NO, 库房id, 药品id, Nvl(批次, 0) As 批次, 实际数量, 成本价, 成本金额, 零售金额, 差价, 入出类别id
                   From 药品收发记录
                   Where 单据 = 21 And 审核日期 Is Null And 药品id = v_Modifybill.药品id And Nvl(批次, 0) = v_Modifybill.批次 And
                         费用id = v_Modifybill.费用id) Loop
      
        Update 药品收发记录
        Set 汇总发药号 = v_Modifybill.Id
        Where 单据 = 21 And 审核日期 Is Null And 药品id = v_Modifybill.药品id And Nvl(批次, 0) = v_Modifybill.批次 And
              费用id = v_Modifybill.费用id;
      
        Zl_材料其他出库_Verify(v_出库.序号, v_出库.No, v_出库.库房id, v_出库.药品id, v_出库.批次, v_出库.实际数量, v_出库.成本价, v_出库.成本金额, v_出库.零售金额,
                         v_出库.差价, v_出库.入出类别id, People_In, d_操作时间);
      End Loop;
    
      --产生备货卫材在卫材仓库的外购入库单据
      For v_入库 In (Select NO, 序号, 供药单位id, 药品id, 产地, 批号, 生产日期, 效期, 灭菌日期, 灭菌效期, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要,
                          注册证号, Nvl(批次, 0) As 批次, 商品条码, 内部条码
                   From 药品收发记录
                   Where 单据 = 21 And 审核日期 Is Not Null And 药品id = v_Modifybill.药品id And Nvl(批次, 0) = v_Modifybill.批次 And
                         费用id = v_Modifybill.费用id And 汇总发药号 = v_Modifybill.Id) Loop
        Begin
          Select 库房id Into v_入库库房id From 虚拟库房对照 Where 科室id = v_Modifybill.库房id;
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
                费用id In (Select Distinct 费用id
                         From 药品收发记录
                         Where 单据 = 21 And 审核日期 Is Not Null And
                               NO = (Select Distinct NO
                                     From 药品收发记录
                                     Where 单据 = 21 And 审核日期 Is Not Null And 费用id = v_Modifybill.费用id));
        
          If v_入库no Is Null Or v_入库no = '' Then
            --如果入库NO为Null, 产生新的入库单NO
            v_入库no := Nextno(68, v_入库库房id);
            n_序号   := 1;
          End If;
        
          Begin
            If v_Modifybill.病人来源 = 1 Then
              Select b.名称 || ',' || a.姓名 || ',' || a.标识号 || ',' || '' As 病人信息
              Into v_病人信息
              From 门诊费用记录 A, 部门表 B
              Where a.病人科室id = b.Id And a.Id = v_Modifybill.费用id;
            Else
              Select b.名称 || ',' || a.姓名 || ',' || a.标识号 || ',' || a.床号 As 病人信息
              Into v_病人信息
              From 住院费用记录 A, 部门表 B
              Where a.病人科室id = b.Id And a.Id = v_Modifybill.费用id;
            End If;
          Exception
            When Others Then
              v_病人信息 := '';
          End;
        
          Zl_材料外购_Insert(v_入库no, n_序号, v_入库库房id, v_入库.供药单位id, v_入库.药品id, v_入库.产地, v_入库.批号, v_入库.生产日期, v_入库.效期,
                         v_入库.灭菌日期, v_入库.灭菌效期, v_入库.实际数量, v_入库.成本价, v_入库.成本金额, v_入库.扣率, v_入库.零售价, v_入库.零售金额, v_入库.差价,
                         Null, '【自动入账】' || v_入库.摘要, v_入库.注册证号, People_In, Null, Null, Null, Null, d_操作时间, Null, Null,
                         v_入库.批次, 1, v_病人信息, v_入库.商品条码, v_入库.内部条码, v_Modifybill.费用id);
        End If;
      End Loop;
    End If;
  End Loop;

  --更新或删除未发药品记录
  Delete 未发药品记录 Where NO = No_In And 单据 = Bill_In And (库房id + 0 = Partid_In Or 库房id Is Null);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料收发记录_处方发料;
/

--127738,李业庆,2018-07-02,财务审核产生数据为0库存
Create Or Replace Procedure Zl_药品外购_Verify
(
  Newno_In    In 药品收发记录.No%Type := Null,
  Oldno_In    In 药品收发记录.No%Type := Null,
  审核人_In   In 药品收发记录.审核人%Type := Null,
  审核日期_In In 药品收发记录.审核日期%Type := Sysdate
) Is
  Err_Isverified Exception;
  Err_Isbatch Exception;
  v_Druginf        Varchar2(50); --原不分批现在分批的药品信息
  v_供药单位id     药品收发记录.供药单位id%Type;
  v_发票金额       应付记录.发票金额%Type;
  v_可用数量       药品库存.可用数量%Type;
  v_时价分批       Number(1);
  n_原成本价       药品收发记录.成本价%Type;
  v_Newno          药品收发记录.No%Type;
  n_New序号        Number;
  n_收发id         药品收发记录.Id%Type;
  n_调整额         药品收发记录.零售金额%Type;
  n_入出类别id     药品收发记录.入出类别id%Type;
  n_售价入出类别id 药品收发记录.入出类别id%Type;
  n_入出系数       药品收发记录.入出系数%Type;
  n_平均成本价     药品库存.平均成本价%Type;
  n_冲销成本价     药品收发记录.成本价%Type;
  n_冲销售价       药品收发记录.零售价%Type;
  v_Billno         药品收发记录.No%Type;

  Cursor c_药品收发记录 Is
    Select a.Id, a.零售价, a.实际数量, a.零售金额, a.差价, a.库房id, a.药品id, a.批次, a.供药单位id, a.成本价, a.批号, a.效期, a.产地, a.入出类别id, a.生产日期,
           a.批准文号, Nvl(b.是否变价, 0) As 时价, Nvl(a.发药方式, 0) As 退库, a.灭菌效期, a.扣率, Nvl(a.计划id, 0) As 计划id,
           Nvl(a.费用id, 0) As 费用id, a.序号
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.No = Newno_In And a.单据 = 1 And a.记录状态 = 1
    Order By a.药品id, a.批次;
Begin

  n_New序号 := 1;

  Select b.Id, b.系数
  Into n_入出类别id, n_入出系数
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 5 And Rownum < 2;

  Select 类别id Into n_售价入出类别id From 药品单据性质 Where 单据 = 13 And Rownum < 2;

  Update 药品收发记录
  Set 审核人 = Nvl(审核人_In, 审核人), 审核日期 = 审核日期_In
  Where NO = Newno_In And 单据 = 1 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  --主要针对原不分批现在分批的药品，不能对其审核
  Begin
    Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
    Into v_Druginf
    From 药品收发记录 A, 药品规格 B, 收费项目目录 I, 收费项目别名 N
    Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = Newno_In And a.单据 = 1 And
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
         Where b.药品id = a.药品id And a.No = Newno_In And a.单据 = 1 And a.记录状态 = 1 And Nvl(a.批次, 0) > 0 And
               (Nvl(b.药库分批, 0) = 0 Or
               (Nvl(b.药房分批, 0) = 0 And
               a.库房id In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室')))));

  For v_药品收发记录 In c_药品收发记录 Loop
    --处理采购计划表中的执行数量，多次导入采用累加执行数量
    If v_药品收发记录.计划id > 0 Then
      Update 药品计划内容
      Set 执行数量 = Nvl(执行数量, 0) + v_药品收发记录.实际数量
      Where 计划id = v_药品收发记录.计划id And 药品id = v_药品收发记录.药品id;
    End If;
    --调用库存更新记录更新库存表
    Zl_药品库存_Update(v_药品收发记录.Id, 0);
  
    --如果是分批药品退库，取原来的成本价
    If v_药品收发记录.退库 = 1 Then
      Begin
        Select 平均成本价
        Into n_原成本价
        From 药品库存
        Where 性质 = 1 And 库房id = v_药品收发记录.库房id And 药品id = v_药品收发记录.药品id And Nvl(批次, 0) = Nvl(v_药品收发记录.批次, 0);
      Exception
        When Others Then
          n_原成本价 := 0;
      End;
    End If;
  
    If v_药品收发记录.时价 = 1 Then
      Update 药品规格 Set 上次售价 = v_药品收发记录.零售价 Where 药品id = v_药品收发记录.药品id;
    End If;
  
    If v_药品收发记录.退库 = 0 Then
      --更新该药品的成本价
      Update 药品规格
      Set 成本价 = v_药品收发记录.成本价, 上次供应商id = v_药品收发记录.供药单位id, 上次批号 = v_药品收发记录.批号, 上次生产日期 = v_药品收发记录.生产日期, 上次产地 = v_药品收发记录.产地,
          上次批准文号 = v_药品收发记录.批准文号
      Where 药品id = v_药品收发记录.药品id;
    End If;
  
    --如果是分批药品退库，则检查成本价是否变动，如果变动，则产生差价调整记录并修正库存差价
    If Oldno_In Is Null Then
      If v_药品收发记录.退库 = 1 Then
        If n_原成本价 <> 0 And n_原成本价 <> v_药品收发记录.成本价 Then
          If v_Newno Is Null Then
            v_Newno := Nextno(25, v_药品收发记录.库房id);
          End If;
        
          --产生库存差价调整单
          Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        
          n_调整额 := (v_药品收发记录.零售金额 - v_药品收发记录.差价) - Round(n_原成本价 * v_药品收发记录.实际数量, 2);
        
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期,
             审核人, 审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率, 灭菌效期, 费用id)
          Values
            (n_收发id, 1, 5, v_Newno, n_New序号, v_药品收发记录.库房id, n_入出类别id, v_药品收发记录.供药单位id, n_入出系数, v_药品收发记录.药品id,
             v_药品收发记录.批次, v_药品收发记录.产地, v_药品收发记录.批号, v_药品收发记录.效期, v_药品收发记录.实际数量, v_药品收发记录.零售金额, v_药品收发记录.差价, n_调整额,
             '外购退库差价误差自动修正', Nvl(审核人_In, Zl_Username), 审核日期_In, Nvl(审核人_In, Zl_Username), 审核日期_In, v_药品收发记录.生产日期,
             v_药品收发记录.批准文号, v_药品收发记录.成本价, 0, n_原成本价, v_药品收发记录.灭菌效期, v_药品收发记录.Id);
        
          n_New序号 := n_New序号 + 1;
          --更新库存
          Zl_药品库存_Update(n_收发id);
        End If;
      End If;
    Else
      Select Distinct 成本价, 零售价
      Into n_冲销成本价, n_冲销售价
      From 药品收发记录
      Where NO = Oldno_In And 药品id = v_药品收发记录.药品id And Nvl(批次, 0) = Nvl(v_药品收发记录.批次, 0) And 序号 = v_药品收发记录.序号 And 单据 = 1 And
            Mod(记录状态, 3) = 2;
      If n_冲销成本价 <> v_药品收发记录.成本价 Then
        --产生库存差价调整单
        v_Newno := Nextno(25, v_药品收发记录.库房id);
        Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
      
        n_调整额 := Round((v_药品收发记录.成本价 - n_冲销成本价) * v_药品收发记录.实际数量, 2);
      
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期,
           审核人, 审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率, 灭菌效期, 费用id)
        Values
          (n_收发id, 1, 5, v_Newno, n_New序号, v_药品收发记录.库房id, n_入出类别id, v_药品收发记录.供药单位id, n_入出系数, v_药品收发记录.药品id, v_药品收发记录.批次,
           v_药品收发记录.产地, v_药品收发记录.批号, v_药品收发记录.效期, v_药品收发记录.实际数量, v_药品收发记录.零售金额, v_药品收发记录.差价, n_调整额, '财务审核价格变动修正',
           Nvl(审核人_In, Zl_Username), 审核日期_In, Nvl(审核人_In, Zl_Username), 审核日期_In, v_药品收发记录.生产日期, v_药品收发记录.批准文号,
           v_药品收发记录.成本价, 0, n_冲销成本价, v_药品收发记录.灭菌效期, v_药品收发记录.Id);
      
        n_New序号 := n_New序号 + 1;
      
        --更新库存
        Zl_药品库存_Update(n_收发id);
      End If;
    
      --更新售价
      If n_冲销售价 <> v_药品收发记录.零售价 Then
        Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        Select Nextno(147) Into v_Billno From Dual;
      
        n_调整额 := Round((n_冲销售价 - v_药品收发记录.零售价) * v_药品收发记录.实际数量, 2);
      
        --产生调价修正记录
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
           填制日期, 库房id, 入出系数, 审核人, 审核日期, 费用id)
        Values
          (n_收发id, 1, 13, v_Billno, n_New序号, n_售价入出类别id, v_药品收发记录.药品id, v_药品收发记录.批次, v_药品收发记录.批号, v_药品收发记录.效期,
           v_药品收发记录.产地, 1, v_药品收发记录.实际数量, 0, n_冲销售价, 0, v_药品收发记录.零售价, 0, n_调整额, n_调整额, '财务审核价格变动修正',
           Nvl(审核人_In, Zl_Username), 审核日期_In, v_药品收发记录.库房id, 1, Nvl(审核人_In, Zl_Username), 审核日期_In, v_药品收发记录.Id);
      
        n_New序号 := n_New序号 + 1;
        --更新药品库存
        Zl_药品库存_Update(n_收发id);
      End If;
    End If;
  End Loop;

  --对应付余额表进行处理
  --此处用一个块，主要是解决没有对应发票号的记录
  Begin
    Update 应付记录
    Set 审核人 = 审核人_In, 审核日期 = 审核日期_In
    Where 入库单据号 = Newno_In And 系统标识 = 1 And 记录性质 = 0 And 记录状态 = 1;
  
    Select b.单位id, Sum(发票金额)
    Into v_供药单位id, v_发票金额
    From 药品收发记录 A, 应付记录 B
    Where a.Id = b.收发id And a.No = Newno_In And a.单据 = 1 And b.系统标识 = 1
    Group By b.单位id;
  
    If Nvl(v_供药单位id, 0) <> 0 Then
      Update 应付余额 Set 金额 = Nvl(金额, 0) + Nvl(v_发票金额, 0) Where 单位id = v_供药单位id And 性质 = 1;
    
      If Sql%NotFound Then
        Insert Into 应付余额 (单位id, 性质, 金额) Values (v_供药单位id, 1, v_发票金额);
      End If;
    End If;
  Exception
    When No_Data_Found Then
      Null;
  End;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品[' || v_Druginf || ']，不能审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品外购_Verify;
/

--128236,余伟节,2018-07-11,首次生成合并路径不匹配
CREATE OR REPLACE Procedure Zl_病人路径评估_Insert
( 
  功能_In            Number, --1=新增,2=修改 
  路径记录id_In      病人临床路径.Id%Type, 
  阶段id_In          临床路径阶段.Id%Type, 
  日期_In            病人路径评估.日期%Type, 
  天数_In            病人路径评估.天数%Type, 
  评估人_In          病人路径评估.评估人%Type, 
  评估结果_In        病人路径评估.评估结果%Type, --0=正常,1=变异继续,2=变异后退出,3=变异后结束（界面程序调用结束路径的过程） 
  评估说明_In        病人路径评估.评估说明%Type, 
  登记人_In          病人路径评估.登记人%Type, 
  变异审核人_In      病人路径评估.变异审核人%Type, 
  变异原因_In        Varchar2, --存在多个变异原因时：变异原因1,变异原因2 
  时间进度_In        病人路径评估.时间进度%Type, --0=正常，1=下一阶段提前至今天，2=下一阶段提前至明天，-1=延后 
  新路径id_In        病人临床路径.路径id%Type, 
  新路径版本_In      病人临床路径.版本号%Type, 
  指标评估_In        Varchar2, --指标名称|指标结果|指标类型||...,末尾带||,允许为空 
  序号_In            Number, 
  跳转审核人_In      病人路径评估.跳转审核人%Type := Null, 
  审核历史跳转_In    Number := 0, 
  结束合并路径ids_In Varchar2 := Null, --本次评估结束的合并路径记录IDs 
  生成时间性质_In    病人路径执行.生成时间性质%Type := Null --当功能_In =2,生成时间性质_IN=1-时，只修改评估结果及变异原因（存在多个变异原因的情况） 
) Is 
  Cursor c_Merge(路径记录id_In 病人路径执行.路径记录id%Type) Is 
    Select a.Id, a.当前阶段id, a.当前天数 
    From 病人合并路径 A 
    Where a.首要路径记录id = 路径记录id_In And a.结束时间 Is Null And a.当前阶段id is not Null; 
  t_合并路径记录id t_Numlist; 
  t_合并路径阶段id t_Numlist; 
  t_合并路径天数   t_Numlist; 
 
  v_Str   Varchar2(4000); 
  v_Tmp   Varchar2(1000); 
  n_Index Number; 
  I       Number(5) := 1; 
 
  l_指标名称 t_Strlist := t_Strlist(); 
  l_指标结果 t_Strlist := t_Strlist(); 
  l_指标类型 t_Numlist := t_Numlist(); 
 
  v_原路径id     病人临床路径.路径id%Type; 
  v_原路径版本   病人临床路径.版本号%Type; 
  d_跳转审核时间 病人路径评估.跳转审核人%Type; 
  d_登记时间     病人路径评估.登记时间%Type; 
  n_当前阶段id   病人合并路径.当前阶段id%Type; 
  d_Date         Date; 
  n_病人id       病人临床路径.病人id%Type; 
  n_主页id       病人临床路径.主页id%Type; 
  n_Count        Number(5); 
  v_Error        Varchar2(255); 
  Err_Custom Exception; 
 
  Procedure p_暂存项目_Delete 
  ( 
    路径记录id_In Number, 
    阶段id_In     Number, 
    日期_In       Date 
  ) Is 
    n_Count Number(5); 
  Begin 
    --变异后退出或变异后结束要删除暂存路径外项目，取消医嘱关联 
    Select Count(1) 
    Into n_Count 
    From 病人路径执行 T 
    Where t.路径记录id = 路径记录id_In And t.阶段id = 阶段id_In And t.日期 = 日期_In And t.项目id Is Null And t.生成时间性质 = 2; 
    If n_Count > 0 Then 
      --取消医嘱关联 
      Delete From 病人路径医嘱 
      Where 路径执行id In (Select a.Id 
                       From 病人路径执行 A, 病人路径医嘱 B 
                       Where a.Id = b.路径执行id And a.路径记录id = 路径记录id_In And a.阶段id = 阶段id_In And a.日期 = 日期_In And 
                             a.项目id Is Null And a.生成时间性质 = 2); 
      --删除表单项目 
      Delete From 病人路径执行 
      Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In And 项目id Is Null And 生成时间性质 = 2; 
    End If; 
  End p_暂存项目_Delete; 
 
Begin 
  Select Sysdate Into d_Date From Dual; 
  If 跳转审核人_In Is Not Null Then 
    d_跳转审核时间 := d_Date; 
  End If; 
  If 序号_In = 1 Then 
    If 功能_In = 1 Then 
      If Nvl(新路径id_In, 0) <> 0 Then 
        Select 路径id, 版本号 Into v_原路径id, v_原路径版本 From 病人临床路径 Where ID = 路径记录id_In; 
        Update 病人临床路径 Set 路径id = 新路径id_In, 版本号 = 新路径版本_In Where ID = 路径记录id_In; 
      End If; 
 
      If 评估结果_In = 2 Or 评估结果_In = 3 Then 
        p_暂存项目_Delete(路径记录id_In, 阶段id_In, 日期_In); 
      End If; 
 
      Insert Into 病人路径评估 
        (路径记录id, 阶段id, 日期, 天数, 评估人, 评估时间, 评估结果, 评估说明, 登记人, 登记时间, 变异原因, 时间进度, 变异审核人, 变异审核时间, 原路径id, 原路径版本, 跳转审核人, 跳转审核时间) 
      Values 
        (路径记录id_In, 阶段id_In, 日期_In, 天数_In, 评估人_In, d_Date, Decode(评估结果_In, 0, 1, -1), 评估说明_In, 登记人_In, d_Date, Null, 
         时间进度_In, 变异审核人_In, d_Date, v_原路径id, v_原路径版本, 跳转审核人_In, d_跳转审核时间); 
 
      If 变异原因_In Is Not Null Then 
        n_Index := 0; 
        For r_变异原因 In (Select Column_Value As 变异原因 From Table(f_Str2list(变异原因_In))) Loop 
          If n_Index = 0 Then 
            --插入一个变异原因到病人路径评估，兼容以前 
            Update 病人路径评估 T 
            Set t.变异原因 = r_变异原因.变异原因 
            Where t.路径记录id = 路径记录id_In And t.阶段id = 阶段id_In And t.日期 = 日期_In; 
            n_Index := 1; 
          End If; 
          Insert Into 病人路径变异 
            (路径记录id, 阶段id, 日期, 变异原因) 
          Values 
            (路径记录id_In, 阶段id_In, 日期_In, r_变异原因.变异原因); 
        End Loop; 
      End If; 
 
      --存储合并路径评估 
      Open c_Merge(路径记录id_In); 
      Fetch c_Merge Bulk Collect 
        Into t_合并路径记录id, t_合并路径阶段id, t_合并路径天数; 
      Close c_Merge; 
      If t_合并路径记录id.Count > 0 Then 
        Forall I In 1 .. t_合并路径记录id.Count 
          Insert Into 病人合并路径评估 
            (路径记录id, 阶段id, 日期, 合并路径记录id, 合并路径阶段id, 合并路径天数, 登记时间) 
          Values 
            (路径记录id_In, 阶段id_In, 日期_In, t_合并路径记录id(I), t_合并路径阶段id(I), t_合并路径天数(I), d_Date); 
      End If; 
    Elsif 功能_In = 2 Then 
      --功能=2 
      If Nvl(生成时间性质_In, 0) <> 1 Then 
        Select 登记时间 
        Into d_登记时间 
        From 病人路径评估 
        Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In; 
        Update 病人路径评估 
        Set 评估人 = 评估人_In, 评估时间 = d_Date, 评估结果 = Decode(评估结果_In, 0, 1, -1), 评估说明 = 评估说明_In, 登记人 = 登记人_In, 登记时间 = d_Date, 
            时间进度 = 时间进度_In, 变异审核人 = 变异审核人_In, 变异审核时间 = d_Date, 跳转审核人 = 跳转审核人_In, 跳转审核时间 = d_跳转审核时间 
        Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In; 
 
        If 评估结果_In = 2 Or 评估结果_In = 3 Then 
          p_暂存项目_Delete(路径记录id_In, 阶段id_In, 日期_In); 
        End If; 
 
        --删除后再插入（存在多个变异原因） 
        If 变异原因_In Is Not Null Then 
          Delete From 病人路径变异 Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In; 
          n_Index := 0; 
          For r_变异原因 In (Select Column_Value As 变异原因 From Table(f_Str2list(变异原因_In))) Loop 
            If n_Index = 0 Then 
              Update 病人路径评估 
              Set 变异原因 = r_变异原因.变异原因 
              Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In; 
              n_Index := 1; 
            End If; 
            Insert Into 病人路径变异 
              (路径记录id, 阶段id, 日期, 变异原因) 
            Values 
              (路径记录id_In, 阶段id_In, 日期_In, r_变异原因.变异原因); 
          End Loop; 
        End If; 
        --补录的路径外项目已经完成评估时，重新调整评估结果及变异原因 
        --1.补录前评估结果为1（正常）时，调整为-1 （变异继续) 
        --2.变异原因删除后重新录入（原因：避免病人路径变异中重复插入相同变异原因值） 
      Elsif Nvl(生成时间性质_In, 0) = 1 Then 
        Delete From 病人路径变异 Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In; 
        n_Index := 0; 
        For r_新变异原因 In (Select Distinct 变异原因 
                        From 病人路径执行 
                        Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In And 变异原因 Is Not Null And 
                              Nvl(生成时间性质, 0) < 2) Loop 
          If n_Index = 0 Then 
            Update 病人路径评估 
            Set 变异原因 = r_新变异原因.变异原因, 评估结果 = Decode(评估结果, 1, -1) 
            Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In; 
            n_Index := 1; 
          End If; 
          Insert Into 病人路径变异 
            (路径记录id, 阶段id, 日期, 变异原因) 
          Values 
            (路径记录id_In, 阶段id_In, 日期_In, r_新变异原因.变异原因); 
        End Loop; 
 
      End If; 
    End If; 
    If 审核历史跳转_In = 1 Then 
      Update 病人路径评估 
      Set 跳转审核人 = 跳转审核人_In, 跳转审核时间 = d_跳转审核时间 
      Where 路径记录id = 路径记录id_In And 原路径id Is Not Null And 跳转审核人 Is Null; 
    End If; 
  End If; 
 
  If Not 指标评估_In Is Null Then 
    v_Str := 指标评估_In; 
    Loop 
      n_Index := Instr(v_Str, '||'); 
      Exit When(Nvl(n_Index, 0) = 0); 
      l_指标名称.Extend; 
      l_指标结果.Extend; 
      l_指标类型.Extend; 
 
      v_Tmp := Substr(v_Str, 1, n_Index - 1); 
      l_指标名称(I) := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1); 
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1); 
      l_指标结果(I) := Trim(Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1)); 
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1); 
      l_指标类型(I) := To_Number(v_Tmp); 
 
      v_Str := Substr(v_Str, n_Index + 2); 
      I     := I + 1; 
    End Loop; 
 
    If 功能_In = 1 Then 
      Forall I In 1 .. l_指标名称.Count 
 
        Insert Into 病人路径指标 
          (路径记录id, 阶段id, 日期, 天数, 评估类型, 评估指标, 指标结果, 指标类型) 
        Values 
          (路径记录id_In, 阶段id_In, 日期_In, 天数_In, 2, l_指标名称(I), l_指标结果(I), l_指标类型(I)); 
    Else 
      Forall I In 1 .. l_指标名称.Count 
        Update 病人路径指标 
        Set 指标结果 = l_指标结果(I) 
        Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In And 评估指标 = l_指标名称(I); 
    End If; 
  End If; 
 
  If 序号_In = 1 And 评估结果_In = 2 Then 
    If 功能_In = 2 Then 
      n_Index := 0; 
      Select 当前阶段id Into n_Index From 病人临床路径 Where ID = 路径记录id_In; 
      If n_Index <> 阶段id_In Then 
        v_Error := '该病人已生成了次日的路径项目,不能修改评估结果来结束路径。'; 
        Raise Err_Custom; 
      End If; 
    End If; 
    --当前天数,不清除,便于统计分析 
    Update 病人临床路径 
    Set 结束时间 = d_Date, 状态 = 3, 前一阶段id = 阶段id_In, 当前阶段id = Null 
    Where ID = 路径记录id_In 
    Returning 病人id, 主页id Into n_病人id, n_主页id; 
 
    --更新病案主页当前路径的状态 
    Update 病案主页 Set 路径状态 = 3 Where 病人id = n_病人id And 主页id = n_主页id; 
 
    --结束合并路径 
    Update 病人合并路径 
    Set 结束时间 = d_Date, 前一阶段id = 当前阶段id, 当前阶段id = Null 
    Where 首要路径记录id = 路径记录id_In And 结束时间 Is Null; 
  Elsif 序号_In = 1 Then 
    --首要路径修改成正常，则取消结束合并路径 
    If 功能_In = 2 and  Nvl(生成时间性质_In, 0) <> 1Then 
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
    End If; 
  End If; 
  If 结束合并路径ids_In Is Not Null Then 
    Update /*+ Rule */ 病人合并路径 
    Set 结束时间 = d_Date, 前一阶段id = 当前阶段id, 当前阶段id = Null 
    Where ID In (Select * From Table(Cast(f_Num2list(结束合并路径ids_In) As Zltools.t_Numlist))); 
 
  End If; 
Exception 
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]'); 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_病人路径评估_Insert;
/
---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.34.160' Where 编号=&n_System;
--部件版本号
Commit;
