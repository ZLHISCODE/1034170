--[连续升级]1
--[管理工具版本号]10.34.130
--本脚本支持从ZLHIS+ v10.34.130 升级到 v10.34.140
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--116339:陈刘,2017-12-14,记录项目管理√符号设置修改
alter table 护理记录项目 add 缺省值 VARCHAR2(100);

--115026:胡俊勇,2017-12-04,病人危急值
Create Table 病人危急值记录(
  ID Number(18),
  数据来源 varchar2(100),    
  病人ID number(18),
  主页ID NUMBER(5),
  挂号单 VARCHAR2(8),
  婴儿 number(3),
  姓名 VARCHAR2(100),
  性别 VARCHAR2(4),
  年龄 varchar2(20),    
  医嘱ID number(18),
  标本ID NUMBER(18),   
  危急值描述 varchar2(2000),       
  报告时间 date,
  报告科室ID number(18),
  报告人 VARCHAR2(20),    
  处理情况 varchar2(2000),
  确认时间 date,          
  确认人 VARCHAR2(20),
  确认科室ID number(18),       
  状态 number(3),      
  是否危急值 number(1),  
  待转出 Number(3)
) TABLESPACE zl9CisRec;

Create Sequence 病人危急值记录_ID Start With 1;

Alter Table 病人危急值记录 Add Constraint 病人危急值记录_PK Primary Key (ID) Using Index Tablespace zl9Indexcis;
Alter Table 病人危急值记录 Add Constraint 病人危急值记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);
Alter Table 病人危急值记录 Add Constraint 病人危急值记录_FK_主页ID Foreign Key (病人ID,主页ID) References 病案主页(病人ID,主页ID);
Alter Table 病人危急值记录 Add Constraint 病人危急值记录_FK_报告科室ID Foreign Key (报告科室ID) References 部门表(ID);
Alter Table 病人危急值记录 Add Constraint 病人危急值记录_FK_确认科室ID Foreign Key (确认科室ID) References 部门表(ID);
Alter Table 病人危急值记录 Add Constraint 病人危急值记录_FK_医嘱ID Foreign Key (医嘱ID) References 病人医嘱记录(ID);
Create Index 病人危急值记录_IX_病人ID On 病人危急值记录(病人ID,主页ID)  Tablespace zl9Indexcis;
Create Index 病人危急值记录_IX_挂号单 On 病人危急值记录(挂号单)  Tablespace zl9Indexcis;
Create Index 病人危急值记录_IX_医嘱ID On 病人危急值记录(医嘱ID) Tablespace zl9Indexcis;
Create Index 病人危急值记录_IX_报告时间 On 病人危急值记录(报告时间)  Tablespace zl9Indexcis;
Create Index 病人危急值记录_IX_待转出 On 病人危急值记录(待转出) Tablespace zl9Indexcis;

CREATE TABLE 病人危急值医嘱(
    危急值ID NUMBER(18),
    医嘱ID NUMBER(18),
    待转出 Number(3))
    TABLESPACE zl9CisRec;

Alter Table 病人危急值医嘱 Add Constraint 病人危急值医嘱_UQ_危急值ID Unique (危急值ID,医嘱ID) Using Index Tablespace zl9Indexcis;
Alter Table 病人危急值医嘱 Add Constraint 病人危急值医嘱_FK_危急值ID Foreign Key (危急值ID) References 病人危急值记录(ID) On Delete Cascade;
Alter Table 病人危急值医嘱 Add Constraint 病人危急值医嘱_FK_医嘱ID Foreign Key (医嘱ID) References 病人医嘱记录(ID) On Delete Cascade; 
Alter Table 病人危急值医嘱 Modify 危急值ID Constraint 病人危急值医嘱_NN_危急值ID Not Null;
Create Index 病人危急值医嘱_IX_医嘱ID On 病人危急值医嘱(医嘱ID) Tablespace zl9Indexcis;
Create Index 病人危急值医嘱_IX_待转出 On 病人危急值医嘱(待转出) Tablespace zl9Indexcis;

CREATE TABLE 病人危急值病历(
    危急值ID NUMBER(18),
    文档ID VARCHAR2(32),
    子文档ID VARCHAR2(32),
    标题 varchar2(100),
    完成人 varchar2(20),
    完成时间 date,
    待转出 Number(3))
    TABLESPACE zl9EprDat;    

Alter Table 病人危急值病历 Add Constraint 病人危急值病历_UQ_危急值ID Unique (危急值ID,文档ID,子文档ID) Using Index Tablespace zl9Indexcis;
Alter Table 病人危急值病历 Add Constraint 病人危急值病历_FK_危急值ID Foreign Key (危急值ID) References 病人危急值记录(ID) On Delete Cascade;
Alter Table 病人危急值病历 Modify 危急值ID Constraint 病人危急值病历_NN_危急值ID Not Null;
Create Index 病人危急值病历_IX_待转出 On 病人危急值病历(待转出) Tablespace zl9Indexcis;

--113432:黄捷,2017-11-09,新版体检修改体检报到号为字符型
alter table RIS医嘱失败记录 rename column 体检报到号 to 体检报到号_bak;
alter table RIS医嘱失败记录 add 体检报到号 VARCHAR2(20);
update RIS医嘱失败记录 set 体检报到号=to_char(体检报到号_bak);

--115695:秦龙,2017-11-09,修改编码字段长度
Alter Table 诊疗分类目录 Modify(编码 Varchar2(20));

--111635:梁唐彬,2017-07-14,XML自定义申请单
CREATE TABLE 自定义申请单文件(
  文件ID NUMBER(18),
  文件名 VARCHAR2(200),
  类别 number(2),
  内容 CLOB,
  创建人 VARCHAR2(20),
  创建时间 DATE
  )
TABLESPACE zl9EprLob;

CREATE TABLE 医嘱申请单文件(
  医嘱ID NUMBER(18),
  文件ID NUMBER(18),
  文件名 VARCHAR2(200),
  类别 number(2),
  内容 CLOB,
  待转出 Number(3)
  )
TABLESPACE zl9EprLob;

Alter Table 病历文件列表 Add(格式 Number(5));

Alter Table 自定义申请单文件 Add Constraint 自定义申请单文件_PK Primary Key (文件ID,类别) Using Index Tablespace zl9Indexcis;

Alter Table 医嘱申请单文件 Add Constraint 医嘱申请单文件_PK Primary Key (医嘱ID,文件ID,类别) Using Index Tablespace zl9Indexcis;

Create Index 医嘱申请单文件_IX_待转出 On 医嘱申请单文件(待转出) Tablespace zl9Indexcis;

Create Index 医嘱申请单文件_IX_文件ID On 医嘱申请单文件(文件ID) Tablespace zl9Indexcis;

--111635:梁唐彬,2017-07-14,XML自定义申请单
Alter Table 自定义申请单文件 Add Constraint 自定义申请单文件_FK_文件ID Foreign Key (文件ID) References 病历文件列表(ID) On Delete Cascade;
Alter Table 医嘱申请单文件 Add Constraint 医嘱申请单文件_FK_医嘱ID Foreign Key (医嘱ID) References 病人医嘱记录(ID) On Delete Cascade;
Alter Table 医嘱申请单文件 Add Constraint 医嘱申请单文件_FK_文件ID Foreign Key (文件ID) References 病历文件列表(ID) On Delete Cascade;

--114434:余伟节,2017-11-17,杭州逸曜合理用药处方序号
Alter Table 病人医嘱记录 Add 处方序号 Number(18);
Create Sequence 病人医嘱记录_处方序号 Start With 1 Cache 100;
Create Index 病人医嘱记录_IX_处方序号 On 病人医嘱记录(处方序号) Pctfree 5 Tablespace zl9Indexcis;


-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--118463:殷瑞,2017-12-18,修正退药单据自动默认为发药状态的问题
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1342, 0, 0, 0, 0, 28, '退药待发单据默认为发药状态', '0', '0', '退药待发单据需要默认为发药状态.1-是;0-不是'
  From Dual;

--118267:杨周一,2017-12-16,LIS图片数据转出单独执行文件
Insert Into Zltools.Zlfilesupgrade
  (序号, 加入日期, 安装路径, 文件类型, 文件名, 版本号, 修改日期, 所属系统, 业务部件, Md5, 文件说明, 自动注册, 强制覆盖, 附加安装路径)
  Select 序号, To_Date('2017-12-16 02:25:54', 'yyyy-mm-dd hh24:mi:ss'), '[APPSOFT]', 0, 'ZLLISPIC2FTP.EXE', Null, Null, Null,
         Null, Null, '部件功能:LIS图片数据转出单独执行文件', 0, 0, Null
  From Dual A, (Select Nvl(Max(To_Number(序号)), 0) + 1 序号 From zlFilesUpgrade) B
  Where Not Exists (Select 1 From Zltools.Zlfilesupgrade Where Upper(文件名) = 'ZLLISPIC2FTP.EXE');

--101301:陈龙,2017-12-06,取血完成消息提示护士站
insert into 业务消息类型(编码,名称,说明,保留天数) values ('ZLHIS_BLOOD_003','取血完成提醒','护士取血完成，提醒护士站',1);

--116339:陈刘,2017-12-14,记录项目管理√符号设置修改
Declare
  Strdata Varchar2(1000);
  Strpre  Varchar2(1000);
  Strtext Varchar2(30);
  Cursor Cur_Item Is
    Select 项目序号, 项目值域, 缺省值 From 护理记录项目 Where 项目表示 In (2, 3);
Begin

  For Row_Format In Cur_Item Loop
    Strdata := Row_Format.项目值域;
    Strpre  := '';
    While Strdata Is Not Null Loop
      If Instr(Strdata, ';', 1) > 0 Then
        Strtext := Substr(Strdata, 1, Instr(Strdata, ';', 1) - 1);
      Else
        Strtext := Strdata;
        Strdata := '';
      End If;
      If Instr(Strtext, '√', 1) = 1 Then
        If Strpre Is Null Then
          Strpre := Substr(Strtext, 2);
        Else
          Strpre := Strpre || ';' || Substr(Strtext, 2);
        End If;
        Strtext := Substr(Strtext, 2);
        Strdata := Substr(Strdata, Instr(Strdata, ';', 1) + 1);
        Strdata := Strpre || ';' || Strdata;
        Update 护理记录项目 Set 缺省值 = Strtext, 项目值域 = Strdata Where 项目序号 = Row_Format.项目序号;
        Exit;
      Else
        If Strpre Is Null Then
          Strpre := Strtext;
        Else
          Strpre := Strpre || ';' || Strtext;
        End If;
        Strdata := Substr(Strdata, Instr(Strdata, ';', 1) + 1);
      End If;
    End Loop;
  End Loop;
  Update 护理记录项目 Set 缺省值 = '√', 项目值域 = '√;√(异)' Where 项目名称 = '生产' And 保留项目 = 1;
End;
/

--116846:刘鹏飞,2017-12-14,输血申请保存自定义函数检查
Insert Into Zlprocedure(Id, 类型, 名称, 状态, 所有者, 说明) Values(Zlprocedure_Id.Nextval,2,'Zl1_EX_BloodApplyCheck',3,User,'新开和修改输血申请时，保存数据之前对申请的相关内容进行检查，并返回提示及处理结果。');

--117641:李南春,2017-12-11,删除无效参数
Delete from Zlparameters Where 系统 = &n_System And  模块 = 1111 and 参数名 = '退号显示详细信息';

--94173:胡俊勇,2017-12-08,校对疑问消息
Insert Into 业务消息类型(编码,名称,说明,保留天数) 
Select 'ZLHIS_CIS_035','校对疑问提醒','护士校对医嘱时设为疑问时产生的一个通知消息。',7 From Dual;

--000000:蒋敏,2017-12-06，期间表数据添加
Insert Into 期间表(期间,开始日期,终止日期) 
Select '201802',to_date('2018-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201803',to_date('2018-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201804',to_date('2018-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201805',to_date('2018-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201806',to_date('2018-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201807',to_date('2018-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201808',to_date('2018-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201809',to_date('2018-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201810',to_date('2018-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201811',to_date('2018-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201812',to_date('2018-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201901',to_date('2019-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201902',to_date('2019-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201903',to_date('2019-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201904',to_date('2019-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201905',to_date('2019-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201906',to_date('2019-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201907',to_date('2019-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201908',to_date('2019-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201909',to_date('2019-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201910',to_date('2019-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201911',to_date('2019-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201912',to_date('2019-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202001',to_date('2020-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202002',to_date('2020-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-02-29 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202003',to_date('2020-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202004',to_date('2020-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202005',to_date('2020-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202006',to_date('2020-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202007',to_date('2020-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202008',to_date('2020-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202009',to_date('2020-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202010',to_date('2020-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202011',to_date('2020-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202012',to_date('2020-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202101',to_date('2021-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202102',to_date('2021-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202103',to_date('2021-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202104',to_date('2021-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202105',to_date('2021-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202106',to_date('2021-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202107',to_date('2021-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202108',to_date('2021-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202109',to_date('2021-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202110',to_date('2021-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202111',to_date('2021-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202112',to_date('2021-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202201',to_date('2022-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202202',to_date('2022-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202203',to_date('2022-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202204',to_date('2022-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202205',to_date('2022-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202206',to_date('2022-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202207',to_date('2022-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202208',to_date('2022-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202209',to_date('2022-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202210',to_date('2022-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202211',to_date('2022-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202212',to_date('2022-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202301',to_date('2023-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202302',to_date('2023-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202303',to_date('2023-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202304',to_date('2023-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202305',to_date('2023-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202306',to_date('2023-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202307',to_date('2023-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202308',to_date('2023-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202309',to_date('2023-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202310',to_date('2023-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202311',to_date('2023-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202312',to_date('2023-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202401',to_date('2024-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202402',to_date('2024-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-02-29 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202403',to_date('2024-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202404',to_date('2024-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202405',to_date('2024-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202406',to_date('2024-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202407',to_date('2024-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202408',to_date('2024-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202409',to_date('2024-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202410',to_date('2024-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202411',to_date('2024-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202412',to_date('2024-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202501',to_date('2025-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202502',to_date('2025-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202503',to_date('2025-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202504',to_date('2025-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202505',to_date('2025-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202506',to_date('2025-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202507',to_date('2025-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202508',to_date('2025-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202509',to_date('2025-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202510',to_date('2025-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202511',to_date('2025-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202512',to_date('2025-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202601',to_date('2026-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202602',to_date('2026-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202603',to_date('2026-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202604',to_date('2026-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202605',to_date('2026-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202606',to_date('2026-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202607',to_date('2026-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202608',to_date('2026-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202609',to_date('2026-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202610',to_date('2026-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202611',to_date('2026-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202612',to_date('2026-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202701',to_date('2027-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202702',to_date('2027-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202703',to_date('2027-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202704',to_date('2027-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202705',to_date('2027-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202706',to_date('2027-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202707',to_date('2027-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202708',to_date('2027-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202709',to_date('2027-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202710',to_date('2027-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202711',to_date('2027-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202712',to_date('2027-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202801',to_date('2028-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202802',to_date('2028-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-02-29 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202803',to_date('2028-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202804',to_date('2028-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202805',to_date('2028-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202806',to_date('2028-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202807',to_date('2028-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202808',to_date('2028-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202809',to_date('2028-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202810',to_date('2028-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202811',to_date('2028-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202812',to_date('2028-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual;

--115026:胡俊勇,2017-12-04,病人危急值
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values( 1284,'病人危急值查询','用于对病人危急值查询统计分析。',&n_System,'zl9CISJob');

Insert Into zlMenus
  (组别, ID, 上级id, 标题, 快键, 说明, 系统, 模块, 短标题, 图标)
  Select 组别, Zlmenus_Id.Nextval, ID, '危急值管理系统', 'D', '用于对病人危急值处理查询分析。', &n_System, -null, '危急值管理', 99
  From zlMenus
  Where 标题 = '临床信息系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null;

Insert Into zlMenus
  (组别, ID, 上级id, 标题, 快键, 说明, 系统, 模块, 短标题, 图标)
  Select a.组别, Zlmenus_Id.Nextval, a.Id, b.*
  From (Select 组别, ID From zlMenus Where 标题 = '危急值管理系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null) A,
       (Select 标题, 快键, 说明, 系统, 模块, 短标题, 图标
         From zlMenus
         Where 1 = 0          
         Union All         
         Select '病人危急值查询', 'D', '用于对病人危急值查询统计分析。', &n_System, 1284, '危急值管理', 99
         From Dual         
         Union All
         Select 标题, 快键, 说明, 系统, 模块, 短标题, 图标
         From zlMenus
         Where 1 = 0) B;

--115026:胡俊勇,2017-12-04,病人危急值
Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
Select &n_System,8,A.* From (
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0 Union All 
      Select '病人危急值记录',25,1,-Null From Dual Union All  
      Select '病人危急值医嘱',26,1,-Null From Dual Union All
      Select '病人危急值病历',27,1,-Null From Dual Union All
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0) A;

Insert Into zlBakTableindex(系统,表名,索引名)
Select &n_System,A.* From (
Select 表名,索引名 From zlBakTableindex Where 1 = 0 Union All
Select '病人危急值记录','病人危急值记录_IX_医嘱ID' From Dual Union All
Select '病人危急值医嘱','病人危急值医嘱_UQ_危急值ID' From Dual Union All
Select '病人危急值病历','病人危急值病历_UQ_危急值ID' From Dual Union All
Select 表名,索引名 From zlBakTableindex Where 1 = 0) A;

--115026:胡俊勇,2017-12-04,病人危急值
Insert into zlTables(系统,表名,表空间,分类) Values(100,'病人危急值记录','ZL9CISREC','B1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'病人危急值医嘱','ZL9CISREC','B1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'病人危急值病历','zl9EprDat','B1');


--115806:刘涛,2017-11-24,修改时价药品入库时取上次售价的参数说明
Update zlParameters
Set 参数说明 = '用来处理药品在外购（其他）入库的时候，售价是按什么方式来的， 0-其他方式取（默认） 1-优先取上次外购入库的售价作为本次售价.'
Where 参数名 = '时价药品入库时取上次售价';

--111635:梁唐彬,2017-07-14,XML自定义申请单
Insert into zlTables(系统,表名,表空间,分类) Values(100,'医嘱申请单文件','ZL9EPRLOB','B1');
Insert into zlTables(系统,表名,表空间,分类) Values(100,'自定义申请单文件','ZL9EPRLOB','A2');

Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
Select &n_System,8,A.* From (
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0 Union All
Select '医嘱申请单文件',24,1,-NULL From Dual Union All
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0) A;

Insert Into zlBakTableindex(系统,表名,索引名)
Select &n_System,A.* From (
Select 表名,索引名 From zlBakTableindex Where 1 = 0 Union All
Select '医嘱申请单文件','医嘱申请单文件_PK' From Dual Union All
Select 表名,索引名 From zlBakTableindex Where 1 = 0) A;

--114364:李南春,2017-11-14,自助发卡流程控制
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值,参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1801, 0, 0, 0, 0, 21, '发卡流程密码控制', NULL, '0',
         '控制发卡时录入密码还是使用缺省密码.0-发卡由病人输入密码;1-发卡使用缺省密码，不进入密码界面'
  From Dual;

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值,参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1801, 0, 0, 0, 0, 22, '缺省密码', NULL, NULL,
         '控制发卡时录入密码还是使用缺省密码.参数格式：类别,卡类别ID,密码缺省方式,缺省固定密码||...'
  From Dual;


--115481:曾杰,2017-11-2,是否自动弹出快捷叫号窗口
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值,参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1291, 0, 1, 0, 0, 57, '自动弹出快捷呼叫窗口', '1', '1','启用排队叫号后是否自动弹出快捷窗口,0-不自动弹出;1-自动弹出'
  From Dual;
   
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值,参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1290, 0, 1, 0, 0, 54, '自动弹出快捷呼叫窗口', '1', '1','启用排队叫号后是否自动弹出快捷窗口,0-不自动弹出;1-自动弹出'
  From Dual;

--114920:刘兴洪,2017-12-11,增加条码输入框,主要解决多段条码的问题
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1257, 1, 0, 0, 0, 18, '上次选择条码控制', Null, '0',
         '用于控制上次是否选择了显示条码输入框，以便再次进入时默认:1-上次选择了条码显示,0-上次未选择条码显示。'
  From Dual;

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1150, 1, 1, 0, 0, 43, '上次选择条码控制', Null, '0',
         '用于控制上次是否选择了显示条码输入框，以便再次进入时默认:1-上次选择了条码显示,0-上次未选择条码显示。'
  From Dual;
 
-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--114601:殷瑞,2017-12-21,补充药品部门发药的基本查看权限
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1342, '基本', User, '病区标记记录', 'SELECT' From Dual;

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1342, '基本', User, '病区标记内容', 'SELECT' From Dual;

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1342, '基本', User, '病人新生儿记录', 'SELECT' From Dual;

--104221:余伟节,2017-12-15,身份证号检查
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,9003,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_Fun_Checkidcard','EXECUTE' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--116846:刘鹏飞,2017-12-14,输血申请保存自定义函数检查
Insert Into zlProgPrivs(系统,序号,所有者,功能,对象,权限) values(&n_System,1252,User,'医嘱下达','Zl1_EX_BloodApplyCheck','EXECUTE');

Insert Into zlProgPrivs(系统,序号,所有者,功能,对象,权限) values(&n_System,1253,User,'医嘱下达','Zl1_EX_BloodApplyCheck','EXECUTE');

--115026:胡俊勇,2017-12-04,病人危急值
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select  &n_System,9001,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '危急值处理',4,'有此权限时，允许调用接口对危急值进行登记',1 From Dual Union All 
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select  &n_System,9001,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '病人危急值记录_ID','SELECT' From Dual Union All 
    Select '病人危急值医嘱','SELECT' From Dual Union All    
    Select '病人危急值记录','SELECT' From Dual Union All    
    Select 'Zl_病人危急值记录_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_病人危急值记录_Update','EXECUTE' From Dual Union All 
    Select 'Zl_病人危急值记录_DELETE','EXECUTE' From Dual Union All
	Select 'Zl_病人危急值记录_处理','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--115026:胡俊勇,2017-12-04,病人危急值
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select 100,1284,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '基本',-NULL,NULL,1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select 100,1284,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select '病人医嘱记录','SELECT' From Dual Union All
Select '病案主页','SELECT' From Dual Union All
Select '部门表','SELECT' From Dual Union All
Select '在院病人','SELECT' From Dual Union All
Select '病人照片','SELECT' From Dual Union All
Select '诊疗项目目录','SELECT' From Dual Union All 
Select '病人信息','SELECT' From Dual Union All
Select '病人挂号记录','SELECT' From Dual Union All
Select '病人危急值记录','SELECT' From Dual Union All
Select '病人危急值医嘱','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--115026:胡俊勇,2017-12-04,病人危急值
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1261,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All 
    Select '危急值处理',17,'有该权限时，住院医生站才处理病人危急值记录。',1 From Dual Union All 
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1260,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All 
    Select '危急值处理',23,'有该权限时，门诊医生站才处理病人危急值记录。',1 From Dual Union All 
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1260,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
    Select '病人危急值记录_ID','SELECT' From Dual Union All 
    Select '病人危急值记录','SELECT' From Dual Union All 
    Select '病人危急值医嘱','SELECT' From Dual Union All  
    Select '病人危急值病历','SELECT' From Dual Union All   
    Select 'Zl_病人危急值记录_Insert','EXECUTE' From Dual Union All 
    Select 'ZL_病人危急值记录_UPDATE','EXECUTE' From Dual Union All
    Select 'Zl_病人危急值记录_DELETE','EXECUTE' From Dual Union All
    Select 'Zl_病人危急值记录_处理','EXECUTE' From Dual Union All
    Select 'Zl_病人危急值医嘱_Update','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1261,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
    Select '病人危急值记录_ID','SELECT' From Dual Union All 
    Select '病人危急值记录','SELECT' From Dual Union All 
    Select '病人危急值医嘱','SELECT' From Dual Union All  
    Select '病人危急值病历','SELECT' From Dual Union All   
    Select 'Zl_病人危急值记录_Insert','EXECUTE' From Dual Union All 
    Select 'ZL_病人危急值记录_UPDATE','EXECUTE' From Dual Union All
    Select 'Zl_病人危急值记录_DELETE','EXECUTE' From Dual Union All
    Select 'Zl_病人危急值记录_处理','EXECUTE' From Dual Union All
    Select 'Zl_病人危急值医嘱_Update','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1252,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '病人危急值医嘱','SELECT' From Dual Union All  
    Select 'Zl_病人危急值医嘱_Update','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1253,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '病人危急值医嘱','SELECT' From Dual Union All  
    Select 'Zl_病人危急值医嘱_Update','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1254,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '病人危急值医嘱','SELECT' From Dual Union All  
    Select 'Zl_病人危急值医嘱_Update','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--117527:冉俊明,2017-12-01,预交款管理使用消费卡支付后退余额，金额未退回消费卡
Delete From zlProgPrivs
Where 系统 = &n_System And 序号 = 1103 And 功能 = '预交退款' And Upper(对象) = 'ZL_病人卡结算记录_STRIKE';

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1103, '预交退款', User, 'ZL_病人卡结算记录_退款', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1103 And 功能 = '预交退款' And Upper(对象) = 'ZL_病人卡结算记录_退款');

--111635:梁唐彬,2017-07-14,XML自定义申请单
Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1253, '基本', User, 'Zl_自定义申请单文件_Edit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1253 And 功能 = '基本' And Upper(对象) = Upper('Zl_自定义申请单文件_Edit'));

Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1253, '基本', User, 'Zl_医嘱申请单文件_Edit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1253 And 功能 = '基本' And Upper(对象) = Upper('Zl_医嘱申请单文件_Edit'));

Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1252, '基本', User, 'Zl_自定义申请单文件_Edit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1252 And 功能 = '基本' And Upper(对象) = Upper('Zl_自定义申请单文件_Edit'));

Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1252, '基本', User, 'Zl_医嘱申请单文件_Edit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1252 And 功能 = '基本' And Upper(对象) = Upper('Zl_医嘱申请单文件_Edit'));

Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1252, '基本', User, '诊治所见分类', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1252 And 功能 = '基本' And Upper(对象) = Upper('诊治所见分类'));

Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1253, '基本', User, '诊治所见分类', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1253 And 功能 = '基本' And Upper(对象) = Upper('诊治所见分类'));

Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1252, '基本', User, '医嘱申请单文件', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1252 And 功能 = '基本' And Upper(对象) = Upper('医嘱申请单文件'));

Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1253, '基本', User, '医嘱申请单文件', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1253 And 功能 = '基本' And Upper(对象) = Upper('医嘱申请单文件'));    

Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1252, '基本', User, '自定义申请单文件', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1252 And 功能 = '基本' And Upper(对象) = Upper('自定义申请单文件'));

Insert Into Zlprogprivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1253, '基本', User, '自定义申请单文件', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where 系统 = &n_System And 序号 = 1253 And 功能 = '基本' And Upper(对象) = Upper('自定义申请单文件'));   


--116034:梁唐彬,2017-11-03,路径生成病历无内容问题
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1256, '基本', User, 'Zl_Lob_ReadForPath', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1256 And 功能 = '基本' And Upper(对象) = Upper('Zl_Lob_ReadForPath'));

--112953:梁唐彬,2017-09-11,药品说明书知识库
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1252, '基本', User, 'Zl_Drugexplain_Readlob', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1252 And 功能 = '基本' And Upper(对象) = Upper('Zl_Drugexplain_Readlob'));
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1253, '基本', User, 'Zl_Drugexplain_Readlob', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1253 And 功能 = '基本' And Upper(对象) = Upper('Zl_Drugexplain_Readlob'));


--114434:余伟节,2017-11-17,杭州逸曜合理用药处方序号
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1252,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select '病人医嘱记录_处方序号','SELECT' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------
--115026:胡俊勇,2017-12-04,病人危急值
--报表：ZL1_INSIDE_1254_20/危急值记录单
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1254_20','危急值记录单','危急值记录单','Yn2t*l~v}1;F~et9C<AD',Null,15,0,0,100,Null,Null,Sysdate,Sysdate,To_Date('2017-11-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2017-11-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'危急值记录单',11904,16832,9,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病人危急值','类别,202|姓名,202|性别,202|年龄,202|科室,202|标号名,202|标号,139|床号名,202|床号,202|危急值描述,202|报告时间,202|报告科室,202|报告人,202|处理情况,202|是否是危急值,202|确认时间,202|确认科室,202|确认人,202',User||'.病人危急值记录,'||User||'.部门表,'||User||'.病案主页,'||User||'.病人挂号记录,'||User||'.病人医嘱记录',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select decode(h.诊疗类别,''D'',''检查类'',''检验类'') as 类别,a.姓名, a.性别, a.年龄,' From Dual Union All
  Select 2,'decode(a.挂号单,null,g.名称,f.名称) as 科室,decode(a.挂号单,null,''住院号'',''门诊号'') as 标号名,decode(a.挂号单,null,d.住院号,e.门诊号) as 标号' From Dual Union All
  Select 3,',decode(a.挂号单,null,''床 号'',''复 诊'') as 床号名,decode(a.挂号单,null,d.出院病床,decode(e.复诊,1,''是'',''否'')) as 床号,a.危急值描述,' From Dual Union All
  Select 4,'To_Char(a.报告时间, ''yyyy-mm-dd hh24:mi'') as 报告时间,b.名称 as 报告科室,a.报告人,a.处理情况, ' From Dual Union All
  Select 5,'decode(a.状态,1,''   是   否 '', decode(a.是否危急值,1,''   是√ 否'',''   是 否√'')) as 是否是危急值,' From Dual Union All
  Select 6,'To_Char(a.确认时间, ''yyyy-mm-dd hh24:mi'') as 确认时间,c.名称 as 确认科室, a.确认人 ' From Dual Union All
  Select 7,'From 病人危急值记录 A,部门表 b,部门表 c,病案主页 d,病人挂号记录 e,部门表 f,部门表 g,病人医嘱记录 h' From Dual Union All
  Select 8,'Where a.报告科室id=b.id and a.确认科室id=c.id(+) and a.病人id=d.病人id(+) and a.主页id=d.主页id(+)' From Dual Union All
  Select 9,'and a.挂号单=e.no(+) and e.执行部门id=f.id(+) and d.出院科室id=g.id(+) and a.医嘱id=h.id(+) and a.Id = [0]' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'记录ID',1,'22',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条14',1,Null,0,Null,0,Null,Null,3970,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条10',1,Null,0,Null,0,Null,Null,2835,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条11',1,Null,0,Null,0,Null,Null,3215,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条13',1,Null,0,Null,0,Null,Null,4335,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条1',1,Null,0,Null,0,Null,Null,4485,1230,2025,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条15',1,Null,0,Null,0,Null,Null,4705,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条16',1,Null,0,Null,0,Null,Null,5070,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条17',1,Null,0,Null,0,Null,Null,5460,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条38',1,Null,0,Null,0,Null,Null,5465,2260,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条35',1,Null,0,Null,0,Null,Null,5465,2905,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条43',1,Null,0,Null,0,Null,Null,5675,8460,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条12',1,Null,0,Null,0,Null,Null,3580,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条40',1,Null,0,Null,0,Null,Null,5720,5390,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条18',1,Null,0,Null,0,Null,Null,5825,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条22',1,Null,0,Null,0,Null,Null,6175,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条21',1,Null,0,Null,0,Null,Null,6540,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条23',1,Null,0,Null,0,Null,Null,6910,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条24',1,Null,0,Null,0,Null,Null,7275,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条25',1,Null,0,Null,0,Null,Null,7665,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条26',1,Null,0,Null,0,Null,Null,8030,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条37',1,Null,0,Null,0,Null,Null,8265,2260,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条45',1,Null,0,Null,0,Null,Null,8265,2905,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条42',1,Null,0,Null,0,Null,Null,8265,5390,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条3',1,Null,0,Null,0,Null,Null,8265,8460,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条27',1,Null,0,Null,0,Null,Null,8410,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条28',1,Null,0,Null,0,Null,Null,8775,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条30',1,Null,0,Null,0,Null,Null,9165,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条29',1,Null,0,Null,0,Null,Null,9530,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条31',1,Null,0,Null,0,Null,Null,9900,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条32',1,Null,0,Null,0,Null,Null,10265,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条4',1,Null,0,Null,0,Null,Null,590,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条6',1,Null,0,Null,0,Null,Null,980,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条5',1,Null,0,Null,0,Null,Null,1345,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条39',1,Null,0,Null,0,Null,Null,1675,2260,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条36',1,Null,0,Null,0,Null,Null,1675,2905,1575,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条41',1,Null,0,Null,0,Null,Null,1675,5390,1905,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条44',1,Null,0,Null,0,Null,Null,1675,8460,1980,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条7',1,Null,0,Null,0,Null,Null,1715,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条8',1,Null,0,Null,0,Null,Null,2080,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条9',1,Null,0,Null,0,Null,Null,2470,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'危急值记录单',Null,4500,885,1980,330,0,0,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,Null,0,'性  别',Null,4770,2010,630,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条33',1,Null,0,Null,0,Null,Null,10655,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签37',2,Null,0,Null,0,'[病人危急值.标号名]',Null,4770,2655,1995,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签29',2,Null,0,Null,0,'报告科室 [病人危急值.报告科室]',Null,4770,5115,3150,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签24',2,Null,0,Null,0,'确认科室 [病人危急值.确认科室]',Null,4770,8205,1935,210,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条34',1,Null,0,Null,0,Null,Null,11020,5565,210,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签26',2,Null,0,Null,0,'是否是危急值 [病人危急值.是否是危急值]',Null,315,7755,3990,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签17',2,Null,0,Null,0,'危急值描述',Null,525,3195,1125,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签25',2,Null,0,Null,0,'确认时间 [病人危急值.确认时间]',Null,690,8205,3150,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签18',2,Null,0,Null,0,'处理情况',Null,750,5970,900,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签28',2,Null,0,Null,0,'报告时间 [病人危急值.报告时间]',Null,810,5115,3150,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,Null,0,'姓  名',Null,1020,2010,630,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签8',2,Null,0,Null,0,'科  室',Null,1020,2655,630,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签33',2,Null,0,Null,0,'[病人危急值.姓名]',Null,1710,2010,1785,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签36',2,Null,0,Null,0,'[病人危急值.科室]',Null,1710,2655,1785,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签27',2,Null,0,Null,0,'[病人危急值.处理情况]',Null,1725,6015,2205,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签31',2,Null,0,Null,0,'[病人危急值.危急值描述]',Null,1755,3225,2415,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签34',2,Null,0,Null,0,'[病人危急值.性别]',Null,5475,2010,1785,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签38',2,Null,0,Null,0,'[病人危急值.标号]',Null,5475,2655,1785,225,0,2,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签23',2,Null,0,Null,0,'确认人 [病人危急值.确认人]',Null,7560,8190,2280,210,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签35',2,Null,0,Null,0,'[病人危急值.年龄]',Null,8310,2010,1785,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签40',2,Null,0,Null,0,'[病人危急值.床号]',Null,8310,2655,1785,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签30',2,Null,0,Null,0,'报告人 [病人危急值.报告人]',Null,7560,5115,2730,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'框线2',10,Null,0,Null,0,'框线1',Null,1715,5975,8115,1440,0,0,0,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'框线1',10,Null,0,Null,0,'框线1',Null,1740,3195,8115,1440,0,0,0,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签32',2,Null,0,Null,0,'[病人危急值.类别]',Null,7560,1470,1785,225,0,0,1,'宋体',10.5,0,0,0,255,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,Null,0,'年  龄',Null,7560,2010,630,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,纵向分栏,横向分栏,左右间距,上下间距,源ID,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签39',2,Null,0,Null,0,'[病人危急值.床号名]',Null,7560,2655,1995,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);


--报表：ZL1_INSIDE_1254_20/危急值记录单
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1252,'危急值记录单','危急值记录单');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1253,'危急值记录单','危急值记录单');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select 100,1252,'危急值记录单',User,'病人危急值记录','SELECT' From Dual Union All
  Select 100,1252,'危急值记录单',User,'病人挂号记录','SELECT' From Dual Union All
  Select 100,1252,'危急值记录单',User,'病人医嘱记录','SELECT' From Dual Union All
  Select 100,1252,'危急值记录单',User,'部门表','SELECT' From Dual Union All 
  Select 100,1252,'危急值记录单',User,'病案主页','SELECT' From Dual Union All  
  Select 100,1253,'危急值记录单',User,'病人危急值记录','SELECT' From Dual Union All
  Select 100,1253,'危急值记录单',User,'病人挂号记录','SELECT' From Dual Union All 
  Select 100,1253,'危急值记录单',User,'病人诊断医嘱','SELECT' From Dual Union All
  Select 100,1253,'危急值记录单',User,'部门表','SELECT' From Dual Union All 
  Select 100,1253,'危急值记录单',User,'病案主页','SELECT' From Dual;


-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--118566:李小东,2017-12-22,老版LIS检验项目管理修改站点为空
Create Or Replace Procedure Zl_检验项目_Edit
(
  编辑类型_In   In Number, --1-增加；2-修改；3-删除
  Id_In         In 诊疗项目目录.Id%Type,
  诊疗分类id_In In 诊疗项目目录.分类id%Type := Null,
  操作类型_In   In 诊疗项目目录.操作类型%Type := Null,
  编码_In       In 诊疗项目目录.编码%Type := Null,
  名称_In       In 诊疗项目目录.名称%Type := Null,
  名称拼音_In   In 诊疗项目别名.简码%Type := Null,
  名称五笔_In   In 诊疗项目别名.简码%Type := Null,
  别名_In       In 诊疗项目别名.名称%Type := Null,
  英文名_In     In 检验项目.缩写%Type := Null,
  计算单位_In   In 诊疗项目目录.计算单位%Type := Null,
  标本部位_In   In 诊疗项目目录.标本部位%Type := Null,
  适用性别_In   In 诊疗项目目录.适用性别%Type := Null,
  单独应用_In   In 诊疗项目目录.单独应用%Type := Null,
  组合项目_In   In 诊疗项目目录.组合项目%Type := Null,
  排列序号_In   In 检验项目.排列序号%Type := Null,
  检验方法_In   In 检验项目.检验方法%Type := Null,
  
  项目类别_In In 检验项目.项目类别%Type := Null,
  结果类型_In In 检验项目.结果类型%Type := Null,
  结果范围_In In 检验项目.结果范围%Type := Null,
  默认值_In   In 检验项目.默认值%Type := Null,
  计算公式_In In 检验项目.计算公式%Type := Null,
  取值序列_In In 检验项目.取值序列%Type := Null,
  隐私项目_In In 检验项目.隐私项目%Type := Null,
  多参考_In   In 检验项目.多参考%Type := Null,
  
  阳性公式_In   In 检验项目.阳性公式%Type := Null,
  弱阳性公式_In In 检验项目.弱阳性公式%Type := Null,
  Cutoff公式_In In 检验项目.Cutoff公式%Type := Null
  
) Is
  v_服务对象   诊疗项目目录.服务对象%Type;
  v_组合项目   诊疗项目目录.组合项目%Type;
  v_执行科室   诊疗项目目录.执行科室%Type;
  v_报告项目id 检验报告项目.报告项目id%Type := 0;
  v_站点       诊疗项目目录.站点%Type;

  Function Get_诊治项目id(诊疗项目id_In In 诊疗项目目录.Id%Type) Return Number Is
    v_诊治项目id 诊治所见项目.Id%Type;
  Begin
    Select 报告项目id Into v_诊治项目id From 检验报告项目 Where 诊疗项目id = 诊疗项目id_In And 报告项目id Is Not Null;
    Return v_诊治项目id;
  Exception
    When Others Then
      Return Null;
  End Get_诊治项目id;

Begin
  If 编辑类型_In = 1 Then
    Zl_诊疗项目_Insert('C', 诊疗分类id_In, Id_In, 编码_In, 名称_In, 名称拼音_In, 名称五笔_In, 别名_In, 名称拼音_In, 名称五笔_In, 操作类型_In, 1, 单独应用_In,
                   3, 计算单位_In, 适用性别_In, 0, 3, 组合项目_In, 标本部位_In, Null, 4, Null, Null, Null, Null, 0);
    --Update 诊疗项目目录 Set 排列序号 = 排列序号_In Where ID = Id_In;
    If 组合项目_In = 0 Then
      Select 诊治所见项目_Id.Nextval Into v_报告项目id From Dual;
    End If;
  Elsif 编辑类型_In = 2 Then
    Select 服务对象, Nvl(组合项目, 0), 执行科室, 站点
    Into v_服务对象, v_组合项目, v_执行科室, v_站点
    From 诊疗项目目录
    Where ID = Id_In;
    Zl_诊疗项目_Update('C', 诊疗分类id_In, Id_In, 编码_In, 名称_In, 名称拼音_In, 名称五笔_In, 别名_In, 名称拼音_In, 名称五笔_In, 操作类型_In, 1, 单独应用_In,
                   3, 计算单位_In, 适用性别_In, 0, v_服务对象, 组合项目_In, 标本部位_In, Null, v_执行科室, Null, Null, Null, Null, 1, 0, Null, 0,
                   0, 0, v_站点);
    --Update 诊疗项目目录 Set 排列序号 = 排列序号_In Where ID = Id_In;
    If v_组合项目 = 0 Then
      v_报告项目id := Get_诊治项目id(Id_In);
      If 组合项目_In = 1 Then
        Delete 检验报告项目 Where 诊疗项目id = Id_In And 细菌id Is Null;
        Delete 诊治所见项目 Where ID = v_报告项目id;
      End If;
    Else
      If 组合项目_In = 0 Then
        Delete 检验报告项目 Where 诊疗项目id = Id_In;
        Select 诊治所见项目_Id.Nextval Into v_报告项目id From Dual;
      End If;
    End If;
    -- 用老版程序增加的项目,可能没有报告项目id 2007-07-13
    If Nvl(v_报告项目id, 0) = 0 Then
      Select 诊治所见项目_Id.Nextval Into v_报告项目id From Dual;
    End If;
  Elsif 编辑类型_In = 3 Then
    Select Nvl(组合项目, 0) Into v_组合项目 From 诊疗项目目录 Where ID = Id_In;
    If v_组合项目 = 0 Then
      v_报告项目id := Get_诊治项目id(Id_In);
      Delete 检验报告项目 Where 诊疗项目id = Id_In;
      Delete 诊治所见项目 Where ID = v_报告项目id;
    End If;
    Delete 诊疗项目目录 Where ID = Id_In;
    Return;
  End If;

  If 组合项目_In = 0 Then
    Update 诊治所见项目
    Set 编码 = 编码_In, 中文名 = 名称_In, 英文名 = 英文名_In, 替换域 = 0, 类型 = Decode(结果类型_In, 1, 0, 2, 1, 3, 3),
        长度 = Decode(结果类型_In, 1, 10, 2, 100, 3, 10), 小数 = Decode(结果类型_In, 1, 3, 2, 0, 3, 0), 单位 = 计算单位_In, 表示法 = 0,
        性别域 = 适用性别_In
    Where ID = v_报告项目id;
    If Sql%RowCount = 0 Then
      Insert Into 诊治所见项目
        (ID, 编码, 中文名, 英文名, 替换域, 类型, 长度, 小数, 单位, 表示法, 性别域)
      Values
        (v_报告项目id, 编码_In, 名称_In, 英文名_In, 0, Decode(结果类型_In, 1, 0, 2, 1, 3, 3), Decode(结果类型_In, 1, 10, 2, 100, 3, 10),
         Decode(结果类型_In, 1, 3, 2, 0, 3, 0), 计算单位_In, 0, 适用性别_In);
      Insert Into 检验报告项目
        (ID, 诊疗项目id, 报告项目id, 检验标本)
      Values
        (检验报告项目_Id.Nextval, Id_In, v_报告项目id, 标本部位_In);
    Else
      Update 检验报告项目 Set 检验标本 = 标本部位_In Where 诊疗项目id = Id_In And 报告项目id = v_报告项目id;
    End If;
  
    Update 检验项目
    Set 缩写 = 英文名_In, 单位 = 计算单位_In, 项目类别 = 项目类别_In, 结果类型 = 结果类型_In, 结果范围 = 结果范围_In, 默认值 = 默认值_In, 计算公式 = 计算公式_In,
        取值序列 = 取值序列_In, 隐私项目 = 隐私项目_In, 阳性公式 = 阳性公式_In, 弱阳性公式 = 弱阳性公式_In, Cutoff公式 = Cutoff公式_In, 排列序号 = 排列序号_In,
        检验方法 = 检验方法_In, 多参考 = 多参考_In
    Where 诊治项目id = v_报告项目id;
    If Sql%RowCount = 0 Then
      Insert Into 检验项目
        (诊治项目id, 缩写, 单位, 项目类别, 结果类型, 结果范围, 默认值, 计算公式, 取值序列, 隐私项目, 阳性公式, 弱阳性公式, Cutoff公式, 排列序号, 检验方法, 多参考)
      Values
        (v_报告项目id, 英文名_In, 计算单位_In, 项目类别_In, 结果类型_In, 结果范围_In, 默认值_In, 计算公式_In, 取值序列_In, 隐私项目_In, 阳性公式_In, 弱阳性公式_In,
         Cutoff公式_In, 排列序号_In, 检验方法_In, 多参考_In);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验项目_Edit;
/

--118738:刘涛,2017-12-20,修改发票信息库存上次供应商处理
Create Or Replace Procedure Zl_材料外购发票信息_Update
(
  No_In         In 药品收发记录.No%Type := Null,
  记录状态_In   In 药品收发记录.记录状态%Type := Null,
  序号_In       In 药品收发记录.序号%Type := Null,
  发票号_In     In 应付记录.发票号%Type := Null,
  发票日期_In   In 应付记录.发票日期%Type := Null,
  发票金额_In   In 应付记录.发票金额%Type := Null,
  供货单位id_In In 应付记录.单位id%Type := 0,
  发票代码_In   In 应付记录.发票代码%Type := Null
) Is
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_No         应付记录.No%Type;
  n_应付id     应付记录.Id%Type;
  n_收发id     应付记录.收发id%Type;
  n_付款序号   应付记录.付款序号%Type;
  n_发票金额   应付记录.发票金额%Type; --旧发票金额
  n_供货单位id 应付记录.单位id%Type;
  n_Dec        Number;
  n_剩余数量   应付记录.发票金额%Type;
Begin
  --金额小数位数
  Select Nvl(精度, 2) Into n_Dec From 药品卫材精度 Where 性质 = 0 And 类别 = 2 And 内容 = 4 And 单位 = 5;

  --取是否付款及总额
  Begin
    Select Max(付款序号), Sum(Nvl(发票金额, 0))
    Into n_付款序号, n_发票金额
    From 应付记录
    Where 收发id In (Select ID From 药品收发记录 Where NO = No_In And 序号 = 序号_In And 单据 = 15) And 系统标识 = 5 And 记录性质 = -1;
  Exception
    When Others Then
      n_发票金额 := 0;
  End;

  n_付款序号 := Nvl(n_付款序号, 0);

  If n_付款序号 <> 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被付了款，不能再修改发票信息[ZLSOFT]';
    Raise Err_Item;
  End If;

  If 发票金额_In > n_发票金额 And n_发票金额 <> 0 Then
    v_Err_Msg := '[ZLSOFT]发票金额不能小于计划付款金额[ZLSOFT]';
    Raise Err_Item;
  End If;
  n_发票金额 := Nvl(n_发票金额, 0);

  --判断是否冲销后的记录
  If 记录状态_In <> 1 Then
    Begin
      Select Sum(Nvl(发票金额, 0))
      Into n_发票金额
      From 应付记录
      Where 收发id In (Select ID From 药品收发记录 Where NO = No_In And 序号 = 序号_In And 单据 = 15) And 系统标识 = 5 And 记录性质 = 0;
    Exception
      When Others Then
        n_发票金额 := 0;
    End;
    n_发票金额 := Nvl(n_发票金额, 0);
    If Nvl(发票号_In, ' ') = ' ' And 发票金额_In <> 0 Then
      v_Err_Msg := '[ZLSOFT]不能对冲销或被冲销记录的发票号改为空,不能保存！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select 供药单位id, Sum(Nvl(实际数量, 0))
    Into n_供货单位id, n_剩余数量
    From 药品收发记录
    Where 单据 = 15 And NO = No_In And 序号 = 序号_In
    Group By 供药单位id;
  
    --更新相关的发票信息,只更改发票号，发票日期
    For v_收发 In (Select a.Id, a.库房id, a.No, a.记录状态, a.零售金额, b.名称, b.规格, b.产地, a.批号, b.计算单位, a.实际数量, a.成本价, a.成本金额, a.填制人,
                        a.填制日期, a.审核人, a.审核日期, a.摘要, a.药品id, a.序号, a.供药单位id
                 From 药品收发记录 A, 收费项目目录 B
                 Where a.单据 = 15 And a.No = No_In And a.序号 = 序号_In And a.药品id = b.Id
                 Order By a.Id) Loop
      Update 应付记录
      Set 发票号 = 发票号_In, 发票代码 = 发票代码_In, 发票日期 = 发票日期_In, 发票金额 = Round((v_收发.实际数量 / n_剩余数量) * 发票金额_In, n_Dec),
          单位id = 供货单位id_In, 发票修改时间 = Sysdate
      Where 收发id = v_收发.Id And 系统标识 = 5 And 记录性质 = 0;
    
      If Sql%RowCount = 0 Then
        If 发票号_In Is Not Null Then
          --如果是第一笔明细,则产生应付记录的NO
          Begin
            Select NO
            Into v_No
            From 应付记录
            Where 系统标识 = 5 And 记录性质 = 0 And 入库单据号 = No_In And Rownum < 2;
          Exception
            When Others Then
              v_No := Nextno(67);
          End;
        
          Select 应付记录_Id.Nextval Into n_应付id From Dual;
          Insert Into 应付记录
            (ID, 记录性质, 记录状态, 单位id, NO, 系统标识, 收发id, 入库单据号, 单据金额, 发票号, 发票日期, 发票金额, 品名, 规格, 产地, 批号, 计量单位, 数量, 采购价, 采购金额,
             填制人, 填制日期, 审核人, 审核日期, 摘要, 项目id, 序号, 库房id, 发票代码, 发票修改时间)
          Values
            (n_应付id, 0, v_收发.记录状态, 供货单位id_In, v_No, 5, v_收发.Id, v_收发.No, v_收发.零售金额, 发票号_In, 发票日期_In,
             Round((v_收发.实际数量 / n_剩余数量) * 发票金额_In, n_Dec), v_收发.名称, v_收发.规格, v_收发.产地, v_收发.批号, v_收发.计算单位, v_收发.实际数量,
             v_收发.成本价, v_收发.成本金额, v_收发.填制人, v_收发.填制日期, v_收发.审核人, v_收发.审核日期, v_收发.摘要, v_收发.药品id, v_收发.序号, v_收发.库房id,
             发票代码_In, Sysdate);
        End If;
      End If;
    End Loop;
  Else
    --未冲销的单据
    Select a.Id, Nvl(b.发票金额, 0), a.供药单位id
    Into n_收发id, n_发票金额, n_供货单位id
    From 药品收发记录 A, (Select * From 应付记录 Where 系统标识 = 5 And 记录性质 = 0 And 记录状态 = 1 And 付款序号 Is Null) B
    Where a.Id = b.收发id(+) And a.No = No_In And a.单据 = 15 And a.记录状态 = 1 And a.序号 = 序号_In;
  
    Update 应付记录
    Set 发票号 = 发票号_In, 发票代码 = 发票代码_In, 发票日期 = 发票日期_In, 发票金额 = 发票金额_In, 单位id = 供货单位id_In, 发票修改时间 = Sysdate
    Where 收发id = n_收发id And 系统标识 = 5 And 记录状态 = 1 And 记录性质 = 0;
  
    If Sql%RowCount = 0 Then
      If 发票号_In Is Not Null Or 发票代码_In Is Not Null Then
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
      
        Select 应付记录_Id.Nextval Into n_应付id From Dual;
      
        Insert Into 应付记录
          (ID, 记录性质, 记录状态, 项目id, 序号, 单位id, NO, 系统标识, 收发id, 入库单据号, 单据金额, 发票号, 发票日期, 发票金额, 品名, 规格, 产地, 批号, 计量单位, 数量, 采购价,
           采购金额, 填制人, 填制日期, 审核人, 审核日期, 摘要, 库房id, 发票代码, 发票修改时间)
          Select n_应付id, 0, 1, a.药品id, a.序号, 供货单位id_In, v_No, 5, n_收发id, a.No, a.零售金额, 发票号_In, 发票日期_In, 发票金额_In, b.名称,
                 b.规格, b.产地, a.批号, b.计算单位, a.实际数量, a.成本价, a.成本金额, a.填制人, a.填制日期, a.审核人, a.审核日期, a.摘要, a.库房id, 发票代码_In,
                 Sysdate
          From 药品收发记录 A, 收费项目目录 B
          Where a.单据 = 15 And a.No = No_In And a.序号 = 序号_In And a.药品id = b.Id;
      End If;
    End If;
  End If;

  Update 应付余额 Set 金额 = Nvl(金额, 0) - n_发票金额 Where 单位id = n_供货单位id And 性质 = 1;
  If Sql%NotFound Then
    Insert Into 应付余额 (单位id, 性质, 金额) Values (n_供货单位id, 1, -n_发票金额);
  End If;
  Update 应付余额 Set 金额 = Nvl(金额, 0) + Nvl(发票金额_In, 0) Where 单位id = 供货单位id_In And 性质 = 1;

  If Sql%NotFound Then
    Insert Into 应付余额 (单位id, 性质, 金额) Values (供货单位id_In, 1, 发票金额_In);
  End If;

  --更新药品收发记录中的供药单位
  Update 药品收发记录 Set 供药单位id = 供货单位id_In Where NO = No_In And 单据 = 15 And 序号 = 序号_In;

  --更新药品库存里的上次供应商
  Update 药品库存
  Set 上次供应商id = 供货单位id_In
  Where 性质 = 1 And (库房id, 药品id, 批次) In (Select 库房id, 药品id, nvl(批次,0) as 批次 From 药品收发记录 Where NO = No_In And 单据 = 15);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销或已经付过款！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料外购发票信息_Update;
/

--118738:刘涛,2017-12-20,修改发票信息库存上次供应商处理
Create Or Replace Procedure Zl_药品外购发票信息_Update
(
  No_In       In 药品收发记录.No%Type := Null,
  序号_In     In 药品收发记录.序号%Type,
  发票号_In   In 应付记录.发票号%Type := Null,
  发票日期_In In 应付记录.发票日期%Type := Null,
  发票金额_In In 应付记录.发票金额%Type := Null,
  供药单位_In In 应付记录.单位id%Type := 0,
  操作标志_In Number, --1、未冲销单据修改发票信息; 2、部分冲销单据修改发票信息
  发票代码_In In 应付记录.发票代码%Type := Null
) Is
  Errinfor Varchar2(255);
  Erritem Exception;

  v_No         应付记录.No%Type;
  v_应付id     应付记录.Id%Type;
  v_收发id     应付记录.收发id%Type;
  v_付款序号   应付记录.付款序号%Type;
  v_发票金额   应付记录.发票金额%Type; --旧发票金额
  v_供药单位id 应付记录.单位id%Type;
  n_Dec        Number;
  n_剩余数量   药品收发记录.实际数量%Type;

  Cursor c_药品记录 Is
    Select a.Id, a.库房id, a.No, a.记录状态, a.零售金额, b.名称, b.规格, b.产地, a.批号, b.计算单位, a.实际数量, a.成本价, a.成本金额, a.填制人, a.填制日期,
           a.审核人, a.审核日期, a.摘要, a.药品id, a.序号, a.供药单位id
    From 药品收发记录 A, 收费项目目录 B
    Where a.单据 = 1 And a.No = No_In And a.序号 = 序号_In And a.药品id = b.Id
    Order By a.Id;
Begin
  --金额小数位数
  Select Nvl(精度, 2) Into n_Dec From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;

  --取是否付款及总额
  Begin
    Select Max(付款序号), Sum(Nvl(发票金额, 0))
    Into v_付款序号, v_发票金额
    From 应付记录
    Where 收发id In (Select ID From 药品收发记录 Where NO = No_In And 序号 = 序号_In And 单据 = 1) And 系统标识 = 1 And 记录性质 = -1;
  Exception
    When Others Then
      v_发票金额 := 0;
      Null;
  End;
  v_发票金额 := Nvl(v_发票金额, 0);
  v_付款序号 := Nvl(v_付款序号, 0);
  If v_付款序号 <> 0 Then
    Errinfor := '[ZLSOFT]该单据已经被付了款，不能再修改发票信息[ZLSOFT]';
    Raise Erritem;
  End If;
  If 发票金额_In > v_发票金额 And v_发票金额 <> 0 Then
    Errinfor := '[ZLSOFT]发票金额不能大于计划付款金额[ZLSOFT]';
    Raise Erritem;
  End If;

  If 操作标志_In = 1 Then
    --未冲销单据
    Select a.Id, Nvl(b.发票金额, 0), a.供药单位id
    Into v_收发id, v_发票金额, v_供药单位id
    From 药品收发记录 A, (Select * From 应付记录 Where 系统标识 = 1 And 记录性质 = 0 And 记录状态 = 1 And 付款序号 Is Null) B
    Where a.Id = b.收发id(+) And a.No = No_In And a.单据 = 1 And a.记录状态 = 1 And a.序号 = 序号_In;
  
    Update 应付记录
    Set 发票号 = 发票号_In, 发票代码 = 发票代码_In, 发票日期 = 发票日期_In, 发票金额 = 发票金额_In, 单位id = 供药单位_In, 发票修改时间 = Sysdate
    Where 收发id = v_收发id And 系统标识 = 1 And 记录状态 = 1 And 记录性质 = 0;
  
    If Sql%RowCount = 0 Then
      If 发票号_In Is Not Null Then
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
           填制日期, 审核人, 审核日期, 摘要, 项目id, 序号, 库房id, 发票修改时间, 发票代码)
          Select v_应付id, 0, 1, 供药单位_In, v_No, 1, v_收发id, a.No, a.零售金额, 发票号_In, 发票日期_In, 发票金额_In, b.名称, b.规格, b.产地, a.批号,
                 b.计算单位, a.实际数量, a.成本价, a.成本金额, a.填制人, a.填制日期, a.审核人, a.审核日期, a.摘要, a.药品id, a.序号, a.库房id, Sysdate,
                 发票代码_In
          From 药品收发记录 A, 收费项目目录 B
          Where a.单据 = 1 And a.No = No_In And a.序号 = 序号_In And a.药品id = b.Id;
      End If;
    End If;
  Else
    --计算原单据的发票金额
    Begin
      Select Sum(Nvl(发票金额, 0))
      Into v_发票金额
      From 应付记录
      Where 收发id In (Select ID From 药品收发记录 Where NO = No_In And 序号 = 序号_In And 单据 = 1) And 系统标识 = 1 And 记录性质 = 0;
    Exception
      When Others Then
        v_发票金额 := 0;
        Null;
    End;
  
    v_发票金额 := Nvl(v_发票金额, 0);
  
    --部分冲销单据，按数量分摊发票金额
    Select 供药单位id, Sum(实际数量)
    Into v_供药单位id, n_剩余数量
    From 药品收发记录
    Where 单据 = 1 And NO = No_In And 序号 = 序号_In
    Group By 供药单位id;
  
    For v_药品记录 In c_药品记录 Loop
      Update 应付记录
      Set 发票号 = 发票号_In, 发票代码 = 发票代码_In, 发票日期 = 发票日期_In, 发票金额 = Round((v_药品记录.实际数量 / n_剩余数量) * 发票金额_In, n_Dec),
          单位id = 供药单位_In, 发票修改时间 = Sysdate
      Where 收发id = v_药品记录.Id And 系统标识 = 1 And 记录性质 = 0;
    
      If Sql%RowCount = 0 Then
        If 发票号_In Is Not Null Then
          --如果是第一笔明细,则产生应付记录的NO
          Begin
            Select NO
            Into v_No
            From 应付记录
            Where 系统标识 = 1 And 记录性质 = 0 And 入库单据号 = No_In And Rownum < 2;
          Exception
            When Others Then
              v_No := Nextno(67);
          End;
        
          Select 应付记录_Id.Nextval Into v_应付id From Dual;
          Insert Into 应付记录
            (ID, 记录性质, 记录状态, 单位id, NO, 系统标识, 收发id, 入库单据号, 单据金额, 发票号, 发票日期, 发票金额, 品名, 规格, 产地, 批号, 计量单位, 数量, 采购价, 采购金额,
             填制人, 填制日期, 审核人, 审核日期, 摘要, 项目id, 序号, 库房id, 发票修改时间, 发票代码)
          Values
            (v_应付id, 0, v_药品记录.记录状态, 供药单位_In, v_No, 1, v_药品记录.Id, v_药品记录.No, v_药品记录.零售金额, 发票号_In, 发票日期_In,
             Round((v_药品记录.实际数量 / n_剩余数量) * 发票金额_In, n_Dec), v_药品记录.名称, v_药品记录.规格, v_药品记录.产地, v_药品记录.批号, v_药品记录.计算单位,
             v_药品记录.实际数量, v_药品记录.成本价, v_药品记录.成本金额, v_药品记录.填制人, v_药品记录.填制日期, v_药品记录.审核人, v_药品记录.审核日期, v_药品记录.摘要,
             v_药品记录.药品id, v_药品记录.序号, v_药品记录.库房id, Sysdate, 发票代码_In);
        End If;
      End If;
    
    End Loop;
  End If;

  Update 应付余额 Set 金额 = Nvl(金额, 0) - v_发票金额 Where 单位id = v_供药单位id And 性质 = 1;
  If Sql%NotFound Then
    Insert Into 应付余额 (单位id, 性质, 金额) Values (v_供药单位id, 1, -v_发票金额);
  End If;
  Update 应付余额 Set 金额 = Nvl(金额, 0) + 发票金额_In Where 单位id = 供药单位_In And 性质 = 1;
  If Sql%NotFound Then
    Insert Into 应付余额 (单位id, 性质, 金额) Values (供药单位_In, 1, 发票金额_In);
  End If;

  --更新药品收发记录中的供药单位
  Update 药品收发记录 Set 供药单位id = 供药单位_In Where NO = No_In And 单据 = 1 And 序号 = 序号_In;

  --更新药品库存里的上次供应商
  Update 药品库存
  Set 上次供应商id = 供药单位_In
  Where (库房id, 药品id, 批次) In (Select 库房id, 药品id, nvl(批次,0) as 批次 From 药品收发记录 Where NO = No_In And 单据 = 1) And 性质 = 1;
Exception
  When Erritem Then
    Raise_Application_Error(-20101, Errinfor);
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销或已经付过款！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品外购发票信息_Update;
/

--106747:李小东,2017-12-19,直接登记的标本回滚到无主状态时删除病人医嘱记录
Create Or Replace Procedure Zl_检验标本记录_转为无主
(
  医嘱id_In   In 检验标本记录.医嘱id%Type,
  删除院外_In In Number := 0
) Is
  --=0删除无主=1不删除无主

  Cursor c_Sample Is
    Select Distinct a.Id As 标本id, Decode(b.医嘱id, Null, a.医嘱id, b.医嘱id) As 医嘱id, a.申请类型, a.病人id, a.病人来源
    From 检验项目分布 B, 检验标本记录 A
    Where a.Id = b.标本id(+) And a.医嘱id = 医嘱id_In;

  Cursor c_Stuff(Vno Varchar2) Is
    Select Distinct s.Id, s.批号, s.实际数量, s.已发数量
    From (Select a.Id, a.批号, a.实际数量, b.已发数量, a.记录状态, a.审核人
           From (Select a.Id, a.药品id, a.序号, a.单据, a.批号, a.实际数量, a.记录状态, a.审核人
                  From 药品收发记录 A
                  Where a.审核人 Is Not Null And Nvl(a.发药方式, 0) <> -1 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0) And a.No = Vno And
                        a.单据 In (24, 25, 26)) A,
                (Select a.单据, a.药品id, a.序号, Sum(a.实际数量) As 已发数量
                  From 药品收发记录 A
                  Where a.审核人 Is Not Null And Nvl(a.发药方式, 0) <> -1 And a.No = Vno And
                        单据 In (Select 单据
                               From 药品收发记录
                               Where NO = Vno And 审核人 Is Not Null And Nvl(发药方式, 0) <> -1 And (记录状态 = 1 Or Mod(记录状态, 3) = 0) And
                                     单据 In (24, 25, 26))
                  Group By a.单据, a.药品id, a.序号) B
           Where a.单据 = b.单据 And a.药品id + 0 = b.药品id And a.序号 = b.序号 And b.已发数量 <> 0) S
    Where (s.记录状态 = 1 Or Mod(s.记录状态, 3) = 0) And s.实际数量 > (s.实际数量 - s.已发数量) And s.审核人 Is Not Null;

  v_Temp       Varchar2(255);
  v_人员部门id 部门人员.部门id%Type;
  v_人员编号   人员表.编号%Type;
  v_人员姓名   人员表.姓名%Type;
  Err_Custom Exception;
  v_Error Varchar2(255);
  v_Flag  Number(1) := 0;

  v_No       Varchar2(20);
  v_当前时间 Date;
  v_主页id   Number(18);
Begin
  v_Temp       := Zl_Identity;
  v_人员部门id := To_Number(Substr(v_Temp, 1, Instr(v_Temp, ',') - 1));
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员编号   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_人员姓名   := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_当前时间   := Sysdate;
  v_Flag       := 0;
  Begin
    Select Nvl(Max(1), 0) Into v_Flag From 检验标本记录 Where 微生物标本 = 1 And 医嘱id = 医嘱id_In;
  Exception
    When Others Then
      v_Flag := 0;
  End;

  If v_Flag = 1 Then
  
    For r_Sample In c_Sample Loop
      Update 检验项目分布
      Set 医嘱id = Null
      Where 标本id In (Select Distinct ID From 检验标本记录 Where 医嘱id = r_Sample.医嘱id);
      Update 检验标本记录
      Set 医嘱id = Null, 姓名 = Null, 性别 = Null, 年龄 = Null, 病人id = Null, 病人来源 = Null, 婴儿 = Null, 合并id = Null, 紧急 = Null,
          挂号单 = Null, 门诊号 = Null, 住院号 = Null, 出生日期 = Null, 主页id = Null, 检验项目 = Null, 操作类型 = Null, 年龄单位 = Null,
          年龄数字 = Null, 申请人 = Null, 申请科室id = Null, 采样人 = Null, 采样时间 = Null, 标本类型 = Null, 标本形态 = Null, 接收人 = Null,
          接收时间 = Null, 样本条码 = Null
      Where 医嘱id = r_Sample.医嘱id;
      If r_Sample.申请类型 = 1 Then
        --删除时有问题时不进行删除
        Begin
          Delete 医嘱执行时间 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id));
          Delete 病人医嘱发送 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id));
          Delete 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id);
          Null;
        Exception
          When Others Then
            Null;
        End;
      Else
        Update 病人医嘱发送
        Set 执行状态 = 0
        Where 执行状态 = 3 And 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id));
      
        If r_Sample.病人来源 = 2 Then
          Update /*+ rule */ 住院费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
          Where 病人id = r_Sample.病人id And 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Sample.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id) And 相关id Is Not Null) And
                       接收人 Is Null
                 Union All
                 Select a.医嘱id, a.记录性质, a.No
                 From 病人医嘱发送 A, 住院费用记录 B
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id) And 相关id Is Null) And
                       采样人 Is Null And a.医嘱id = b.医嘱序号 And a.记录性质 = b.记录性质 And a.No = b.No And
                       b.执行人 In (Select Distinct 姓名
                                 From 人员表 A, 部门人员 B, 部门性质说明 C
                                 Where a.Id = b.人员id And b.部门id = c.部门id And c.工作性质 = '检验'));
        Else
          Update /*+ rule */ 门诊费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
          Where 病人id = r_Sample.病人id And 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Sample.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id) And 相关id Is Not Null) And
                       接收人 Is Null
                 Union All
                 Select a.医嘱id, a.记录性质, a.No
                 From 病人医嘱发送 A, 门诊费用记录 B
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id) And 相关id Is Null) And
                       采样人 Is Null And a.医嘱id = b.医嘱序号 And a.记录性质 = b.记录性质 And a.No = b.No And
                       b.执行人 In (Select Distinct 姓名
                                 From 人员表 A, 部门人员 B, 部门性质说明 C
                                 Where a.Id = b.人员id And b.部门id = c.部门id And c.工作性质 = '检验'));
        
        End If;
        --取消试剂消耗单的审核
        v_No := '';
        Begin
          Select Distinct NO Into v_No From 检验试剂记录 Where 医嘱id = r_Sample.医嘱id;
        Exception
          When Others Then
            v_No := '';
        End;
      
        If v_No Is Not Null Then
        
          For r_Stuff In c_Stuff(v_No) Loop
            Zl_材料收发记录_部门退料(r_Stuff.Id, v_人员姓名, v_当前时间, r_Stuff.批号, Null, Null, r_Stuff.已发数量, 0, Null, 0);
          End Loop;
        
          v_主页id := Null;
          Select 主页id Into v_主页id From 病人医嘱记录 A Where ID = r_Sample.医嘱id;
        
          If v_主页id Is Null Then
            Zl_门诊记帐记录_Delete(v_No, '', v_人员编号, v_人员姓名);
          Else
            Zl_住院记帐记录_Delete(v_No, '', v_人员编号, v_人员姓名);
          End If;
          Update 检验试剂记录 Set NO = '' Where 医嘱id = r_Sample.医嘱id;
        End If;
      End If;
    
    End Loop;
  
  Else
  
    For r_Sample In c_Sample Loop
    
      --检查是否允许取消核收
      v_Flag := 0;
      Begin
        Select 1 Into v_Flag From 检验标本记录 Where ID = r_Sample.标本id And 样本状态 = 2;
      Exception
        When Others Then
          v_Flag := 0;
      End;
    
      If v_Flag = 1 Then
        v_Error := '当前申请所在的标本中有已经被审核的，请先取消审核！';
        Raise Err_Custom;
      End If;
    
      --删除合并关联项目
      Update 检验标本记录 Set 合并id = Null Where 合并id In (Select ID From 检验标本记录 Where 医嘱id = 医嘱id_In);
      --更改检验标本记录里记录的医嘱id,其实这可以不要此信息,以后考虑取消
      Update 检验标本记录
      Set 医嘱id = Null, 姓名 = Null, 性别 = Null, 年龄 = Null, 病人id = Null, 病人来源 = Null, 婴儿 = Null, 合并id = Null, 紧急 = Null,
          挂号单 = Null, 门诊号 = Null, 住院号 = Null, 出生日期 = Null, 主页id = Null, 检验项目 = Null, 操作类型 = Null, 年龄单位 = Null,
          年龄数字 = Null, 申请人 = Null, 申请科室id = Null, 采样人 = Null, 采样时间 = Null, 标本类型 = Null, 标本形态 = Null, 接收人 = Null,
          接收时间 = Null, 样本条码 = Null
      Where ID = r_Sample.标本id;
    
      Update 检验项目分布 Set 医嘱id = Null Where 标本id = r_Sample.标本id;
    
      If r_Sample.申请类型 = 1 Then
        --删除时有问题时不进行删除
        Begin
          Delete 医嘱执行时间 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id));
          Delete 病人医嘱发送 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id));
          Delete 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id);
          Null;
        Exception
          When Others Then
            Null;
        End;
      Else
        Update 病人医嘱发送
        Set 执行状态 = 0
        Where 执行状态 = 3 And 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id));
      
        If r_Sample.病人来源 = 2 Then
          Update /*+ rule */ 住院费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
          Where 病人id = r_Sample.病人id And 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Sample.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id) And 相关id Is Not Null) And
                       接收人 Is Null
                 Union All
                 Select a.医嘱id, a.记录性质, a.No
                 From 病人医嘱发送 A, 住院费用记录 B
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id) And 相关id Is Null) And
                       采样人 Is Null And a.医嘱id = b.医嘱序号 And a.记录性质 = b.记录性质 And a.No = b.No And
                       b.执行人 In (Select Distinct 姓名
                                 From 人员表 A, 部门人员 B, 部门性质说明 C
                                 Where a.Id = b.人员id And b.部门id = c.部门id And c.工作性质 = '检验'));
        Else
          Update /*+ rule */ 门诊费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
          Where 病人id = r_Sample.病人id And 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Sample.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id) And 相关id Is Not Null) And
                       接收人 Is Null
                 Union All
                 Select a.医嘱id, a.记录性质, a.No
                 From 病人医嘱发送 A, 门诊费用记录 B
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Sample.医嘱id In (ID, 相关id) And 相关id Is Null) And
                       采样人 Is Null And a.医嘱id = b.医嘱序号 And a.记录性质 = b.记录性质 And a.No = b.No And
                       b.执行人 In (Select Distinct 姓名
                                 From 人员表 A, 部门人员 B, 部门性质说明 C
                                 Where a.Id = b.人员id And b.部门id = c.部门id And c.工作性质 = '检验'));
        End If;
      End If;
      --取消试剂消耗单的审核
      v_No := '';
      Begin
        Select Distinct NO Into v_No From 检验试剂记录 Where 医嘱id = r_Sample.医嘱id;
      Exception
        When Others Then
          v_No := '';
      End;
      If v_No Is Not Null Then
        For r_Stuff In c_Stuff(v_No) Loop
          Zl_材料收发记录_部门退料(r_Stuff.Id, v_人员姓名, v_当前时间, r_Stuff.批号, Null, Null, r_Stuff.已发数量, 0, Null, 0);
        End Loop;
      
        v_主页id := Null;
        Select 主页id Into v_主页id From 病人医嘱记录 A Where ID = r_Sample.医嘱id;
      
        If v_主页id Is Null Then
          Zl_门诊记帐记录_Delete(v_No, '', v_人员编号, v_人员姓名);
        Else
          Zl_住院记帐记录_Delete(v_No, '', v_人员编号, v_人员姓名);
        End If;
        Update 检验试剂记录 Set NO = '' Where 医嘱id = r_Sample.医嘱id;
      End If;
      --删除试剂消耗单
    --Delete From 检验试剂记录 Where 医嘱id = r_Sample.医嘱id;
    End Loop;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验标本记录_转为无主;
/

--115597:焦博,2017-12-18,严格控制并支持重复使用的院内卡，当数据库中有相同卡号时，重复使用该卡号票据领用记录将不会发生数量变化
Create Or Replace Procedure Zl_医疗卡记录_Insert
(
  --参数：发卡类型=0-发卡,1-补卡,2-换卡(相当于重打)
  --      换卡时,单据号_IN传入的是原发/补卡的单据号。
  --      补卡/换卡后,再换卡时是以最后一次卡号为准。
  发卡类型_In     Number,
  单据号_In       住院费用记录.No%Type,
  病人id_In       住院费用记录.病人id%Type,
  主页id_In       住院费用记录.主页id%Type,
  标识号_In       住院费用记录.标识号%Type,
  费别_In         住院费用记录.费别%Type,
  卡类别id_In     医疗卡类别.Id%Type,
  原卡号_In       病人医疗卡信息.卡号%Type,
  医疗卡号_In     病人医疗卡信息.卡号%Type,
  变动原因_In     病人医疗卡变动.变动原因%Type,
  密码_In         病人信息.卡验证码%Type,
  姓名_In         住院费用记录.姓名%Type,
  性别_In         住院费用记录.性别%Type,
  年龄_In         住院费用记录.年龄%Type,
  病人病区id_In   住院费用记录.病人病区id%Type,
  病人科室id_In   住院费用记录.病人科室id%Type,
  收费细目id_In   住院费用记录.收费细目id%Type,
  收费类别_In     住院费用记录.收费类别%Type,
  计算单位_In     住院费用记录.计算单位%Type,
  收入项目id_In   住院费用记录.收入项目id%Type,
  收据费目_In     住院费用记录.收据费目%Type,
  标准单价_In     住院费用记录.标准单价%Type,
  执行部门id_In   住院费用记录.执行部门id%Type,
  开单部门id_In   住院费用记录.开单部门id%Type,
  操作员编号_In   住院费用记录.操作员编号%Type,
  操作员姓名_In   住院费用记录.操作员姓名%Type,
  加班标志_In     住院费用记录.加班标志%Type,
  发卡时间_In     住院费用记录.登记时间%Type,
  领用id_In       票据使用明细.领用id%Type,
  Ic卡号_In       病人信息.Ic卡号%Type := Null,
  应收金额_In     住院费用记录.应收金额%Type,
  实收金额_In     住院费用记录.实收金额%Type,
  结算方式_In     病人预交记录.结算方式%Type,
  刷卡类别id_In   病人预交记录.卡类别id%Type,
  消费卡_In       Integer := 0,
  刷卡卡号_In     病人医疗卡信息.卡号%Type,
  结帐id_In       病人预交记录.结帐id%Type,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  合作单位_In     病人预交记录.合作单位%Type := Null,
  更新交款余额_In Number := 0, --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况。
  摘要_In         住院费用记录.摘要%Type := Null
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
       结帐金额, 缴款组id, 结论, 摘要)
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
      Select Nvl(Max(a.回收次数), 0), Nvl(Max(a.性质), 0)
      Into n_回收次数, n_性质
      From 票据使用明细 A, 票据打印内容 B, 住院费用记录 C
      Where a.打印id = b.Id And b.No = c.No And a.票种 = 5 And c.结论 = 卡类别id_In And c.记录性质 = 5 And a.号码 = 医疗卡号_In;
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
      If Nvl(更新交款余额_In, 0) = 0 Then
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

--118465:廖思奇,2017-12-15,提供最新的过程
Create Or Replace Procedure Zl_影像插件功能_Update
(
  插件id_In           In 影像插件功能.插件id%Type,
  名称_In             In 影像插件功能.名称%Type,
  方法_In             In 影像插件功能.方法%Type,
  方法参数_In         In 影像插件功能.方法参数%Type,
  是否启用_In         In 影像插件功能.是否启用%Type,
  是否加入右键菜单_In In 影像插件功能.是否加入右键菜单%Type,
  是否加入工具栏_In   In 影像插件功能.是否加入工具栏%Type,
  自动执行时机_In     In 影像插件功能.自动执行时机%Type,
  Vbs脚本_In          In 影像插件功能.Vbs脚本%Type
) Is

  n_功能序号 影像插件功能.功能序号%Type;
  n_Id       Number;

Begin

  Select Nvl(Max(功能序号), 0) + 1 Into n_功能序号 From 影像插件功能 Where 插件id = 插件id_In;
  Select Nvl(Max(ID), 0) + 1 Into n_Id From 影像插件功能;

  Insert Into 影像插件功能
    (ID, 插件id, 功能序号, 名称, 方法, 方法参数, 是否启用, 是否加入右键菜单, 是否加入工具栏, 自动执行时机, Vbs脚本)
  Values
    (n_Id, 插件id_In, n_功能序号, 名称_In, 方法_In, 方法参数_In, 是否启用_In, 是否加入右键菜单_In, 是否加入工具栏_In, 自动执行时机_In, Vbs脚本_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_影像插件功能_Update;
/

--116339:陈刘,2017-12-14,记录项目管理√符号设置修改
Create Or Replace Procedure Zl_护理记录项目_Insert
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
  缺省值_In     In 护理记录项目.缺省值%Type := Null
) Is
Begin
  Insert Into 护理记录项目
    (项目序号, 项目名称, 项目类型, 项目长度, 项目小数, 项目单位, 项目表示, 项目值域, 护理等级, 分组名, 项目id, 适用科室, 应用方式, 适用病人, 项目性质, 应用场合, 说明, 缺省值)
  Values
    (项目序号_In, 项目名称_In, 项目类型_In, 项目长度_In, 项目小数_In, 项目单位_In, 项目表示_In, 项目值域_In, 护理等级_In, 分组名_In, 项目id_In, 1, 应用方式_In,
     适用病人_In, 项目性质_In, 应用场合_In, 说明_In, 缺省值_In);

  If 项目表示_In = 4 Then
    Insert Into 护理汇总项目
      (序号, 父序号)
      Select 项目序号_In, Null From Dual Where Not Exists (Select 1 From 护理汇总项目 Where 序号 = 项目序号_In);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_护理记录项目_Insert;
/

--116339:陈刘,2017-12-14,记录项目管理√符号设置修改
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
  缺省值_In     In 护理记录项目.缺省值%Type := Null
) Is
  n_汇总 Number(1);
Begin
  n_汇总 := 0;
  Select Count(项目序号) Into n_汇总 From 护理记录项目 Where 项目序号 = 项目序号_In And 项目表示 = 4;
  Update 护理记录项目
  Set 项目名称 = 项目名称_In, 项目类型 = 项目类型_In, 项目长度 = 项目长度_In, 项目小数 = 项目小数_In, 项目单位 = 项目单位_In, 项目表示 = 项目表示_In, 项目值域 = 项目值域_In,
      护理等级 = 护理等级_In, 分组名 = 分组名_In, 项目id = 项目id_In, 应用方式 = 应用方式_In, 适用病人 = 适用病人_In, 项目性质 = 项目性质_In, 应用场合 = 应用场合_In,
      说明 = 说明_In, 缺省值 = 缺省值_In
  Where 项目序号 = 项目序号_In;

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

--116848:刘鹏飞,2017-12-14,输血申请检验结果血型对照
Create Or Replace Function Zl_Fun_BloodApplyCode
(
  申请类型_In Number,
  用血安排_In Number,
  模式_In     Number := 1
) Return Varchar2 As
  v_Return Varchar2(100);
Begin
  --功能说明:下单申请时：1、强制控制是否允许修改申请单ABO和RH；2、返回输血申请单上检验指标中的ABO和RH指标代码(也可只返回ABO指标代码)。
  ----                                                        A：提取检验结果时则自动更新ABO和RH。B:保存时检查ABO和RH是否和检验结果内容一致

  --入参说明：
  ----申请类型_In=1-输血申请单;2-取血通知单(便于医院根据申请类型控制)
  ----用血安排_In=0-普通输血;1-紧急输血(便于医院根据输血紧急程度控制)
  ----模式_in:0=通过参数值控制是否允许允许申请单ABO和RH；1=根据检验结果更新ABO和RH；保存是检查ABO、RH是否和检验结果一致(输血申请单时有效)
  --函数返回：模式_in=0时，返回0(允许修改)或1(不允许修改)；模式_in=1时，返回字符串格式: ABO指标代码:0(询问)或1(禁止),RH指标代码:0(询问)或1(禁止)，
  ----        如：800001:1,表示保存时ABO和检验结果不一致时则禁止保存，也可直接写指标代码，如：800001，表示保存时ABO和检验结果不一致时则进行询问。

  If  模式_In = 0 Then
    --0表示允许修改ABO;1不允许修改ABO
    v_Return := '0';
  Else
    --返回空不自动匹配ABO和RH，且保存时不进行检查。
    v_Return := '';
  End If;
  Return v_Return;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Fun_BloodApplyCode;
/

--116846:刘鹏飞,2017-12-14,输血申请保存自定义函数检查
Create Or Replace Function Zl1_EX_BloodApplyCheck
(
  调用场合_In         Number,
  病人id_In           病人医嘱记录.病人id%Type,
  就诊id_In           Number,
  申请类型_In         Number,
  用血安排_In         Number,
  是否待诊_In         输血申请记录.是否待诊%Type,
  诊断内容_In         病人医嘱附件.内容%Type,
  诊断ids_In          Varchar2,
  输血类型_In         输血类型.名称%Type,
  输血目的_In         输血目的.名称%Type,
  输血性质_In         输血性质.名称%Type,
  预定输血日期_In     病人医嘱记录.标本部位%Type,
  血型_In             血型.名称%Type,
  Rhd_In              Varchar2,
  申请项目_In         Varchar2,
  输血执行科室id_In   病人医嘱记录.执行科室id%Type,
  途径id_In           病人医嘱记录.诊疗项目id%Type,
  途径执行科室id_In   病人医嘱记录.执行科室id%Type,
  备注_In             病人医嘱记录.医生嘱托%Type,
  婴儿序号_In         病人医嘱记录.婴儿%Type := 0,
  即往输血史_In       输血申请记录.即往输血史%Type := Null,
  既往输血反应史_In   输血申请记录.既往输血反应史%Type := Null,
  输血禁忌及过敏史_In 输血申请记录.输血禁忌及过敏史%Type := Null,
  孕产情况_In         输血申请记录.孕产情况%Type := Null,
  受血者属地_In       输血申请记录.受血者属地%Type := Null,
  检验结果_In         Varchar2 := Null
) Return Varchar2
--功能说明：新开和修改输血申请时，保存数据之前对申请的相关内容进行检查，并返回提示及处理结果。
--适用说明：如过需在数据保存之前对申请单内容进行特定检查和控制，以满足医院业务特定环节，则请调整此函数
--入参说明：
  ----调用场合_in=1-门诊,2-住院 
  ----就诊id_In=门诊时传挂号记录id,住院时传入主页id 
  ----申请类型_In=1-输血申请单;2-取血通知单
  ----用血安排_In=0-普通输血;1-紧急输血
  ----是否待诊_In=0-非待诊;1-待诊
  ----诊断内容_In=待诊时为空，否则为诊断内容信息
  ----诊断ids_In=从首页选择的诊断则传入诊断iD，多个诊断以','号分割，自由录入的诊断为空
  ----输血类型_In=对应输血类型字典表的名称
  ----输血目的_In=对应输血目的字典表的名称
  ----输血性质_In=对应输血性质字典表的名称
  ----血型_In=对应血型字典表的名称
  ----Rhd_In=;+;-
  ----申请项目_In=以诊疗项目+申请量的方式传入，如申请多个品种则以';'分割，格式如：输血诊疗项目ID,申请量;输血诊疗项目ID,申请量
  ----途径id_In=输血申请单则是采集方式的诊疗项目id，取血通知单则是输血途径的诊疗项目ID
  ----孕产情况_In=格式:孕次/产次
  ----检验结果_In=输血申请则返回申请单下方的检验结果信息（字段内容请参考:输血检验结果），取血申请则为空。单个指标的内容以<SplitCol>分割，不同指标之间以<SplitRow>分割，返回格式如下：
  ----            检验项目ID<SplitCol>指标代码<SplitCol>指标中文名<SplitCol>指标英文名<SplitCol>指标结果<SplitCol>结果单位<SplitCol>结果标志<SplitCol>结果参考
  ----            <SplitCol>取值序列<SplitCol>是否人工填写<SplitRow>检验项目ID<SplitCol>指标代码<SplitCol>指标中文名<SplitCol>指标英文名<SplitCol>指标结果
  ----            <SplitCol>结果单位<SplitCol>结果标志<SplitCol>结果参考<SplitCol>取值序列<SplitCol>是否人工填写
  ----    
--函数返回："处理结果|提示信息",处理结果=0-正常,1-询问提示,2-禁止；处理结果为0时，无需返回提示信息及分隔符。 
 As
  v_Return Varchar2(200);
Begin
  v_Return := Null;
  Return v_Return;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl1_EX_BloodApplyCheck;
/

--118154:余伟节,2017-12-12,合理用药杭州逸曜
Create Or Replace Procedure Zl_病人医嘱记录_Update
(
  Id_In           病人医嘱记录.Id%Type,
  相关id_In       病人医嘱记录.相关id%Type,
  序号_In         病人医嘱记录.序号%Type,
  医嘱状态_In     病人医嘱记录.医嘱状态%Type,
  医嘱期效_In     病人医嘱记录.医嘱期效%Type,
  诊疗项目id_In   病人医嘱记录.诊疗项目id%Type,
  收费细目id_In   病人医嘱记录.收费细目id%Type,
  天数_In         病人医嘱记录.天数%Type,
  单次用量_In     病人医嘱记录.单次用量%Type,
  总给予量_In     病人医嘱记录.总给予量%Type,
  医嘱内容_In     病人医嘱记录.医嘱内容%Type,
  医生嘱托_In     病人医嘱记录.医生嘱托%Type,
  标本部位_In     病人医嘱记录.标本部位%Type,
  执行频次_In     病人医嘱记录.执行频次%Type,
  频率次数_In     病人医嘱记录.频率次数%Type,
  频率间隔_In     病人医嘱记录.频率间隔%Type,
  间隔单位_In     病人医嘱记录.间隔单位%Type,
  执行时间方案_In 病人医嘱记录.执行时间方案%Type,
  计价特性_In     病人医嘱记录.计价特性%Type,
  执行科室id_In   病人医嘱记录.执行科室id%Type,
  执行性质_In     病人医嘱记录.执行性质%Type,
  紧急标志_In     病人医嘱记录.紧急标志%Type,
  开始执行时间_In 病人医嘱记录.开始执行时间%Type,
  执行终止时间_In 病人医嘱记录.执行终止时间%Type,
  病人科室id_In   病人医嘱记录.病人科室id%Type,
  开嘱科室id_In   病人医嘱记录.开嘱科室id%Type,
  开嘱医生_In     病人医嘱记录.开嘱医生%Type,
  开嘱时间_In     病人医嘱记录.开嘱时间%Type,
  检查方法_In     病人医嘱记录.检查方法%Type := Null,
  执行标记_In     病人医嘱记录.执行标记%Type := Null,
  可否分零_In     病人医嘱记录.可否分零%Type := Null,
  摘要_In         病人医嘱记录.摘要%Type := Null,
  操员作姓名_In   病人医嘱状态.操作人员%Type := Null,
  零费记帐_In     病人医嘱记录.零费记帐%Type := Null,
  用药目的_In     病人医嘱记录.用药目的%Type := Null,
  用药理由_In     病人医嘱记录.用药理由%Type := Null,
  审核状态_In     病人医嘱记录.审核状态%Type := Null,
  超量说明_In     病人医嘱记录.超量说明%Type := Null,
  首次用量_In     病人医嘱记录.首次用量%Type := Null,
  手术情况_In     病人医嘱记录.手术情况%Type := Null,
  组合项目id_In   病人医嘱记录.组合项目id%Type := Null,
  皮试结果_In     病人医嘱记录.皮试结果%Type := Null,
  处方序号_In     病人医嘱记录.处方序号%Type := Null
  --功能：被医生或护士修改了部分内容的医嘱记录。可用于门诊或住院。
  --说明：Update时之所以涉及诊疗项目ID,计价特性变化,是因为给药途径,用法的变化
  --      Update时之所以涉及期效变化,是因为自由录入医嘱可任意改变期效
) Is
  v_Count Number;

  v_Temp     Varchar2(255);
  v_人员姓名 病人医嘱状态.操作人员%Type;
  v_处方审查锁定IDs varChar2(4000);

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
     zl_处方审查_cancel(Id_In,v_处方审查锁定IDs);
  End If;

  If v_处方审查锁定IDs Is Not Null Then
  v_Error :=  '医嘱"' || 医嘱内容_In || '"已锁定，正在进行处方审查，不能再修改。';
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
      处方序号 = 处方序号_In
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

--107484:廖思奇,2017-12-12,调整Zl_影像消息_Xml内容获取中 婴儿病人姓名获取方式
--107484:廖思奇,2017-12-13,布局还原成以前的样式，增加过程错误处理
--117484:廖思奇,2017-12-14,调整Zl_影像消息_Xml内容获取中 婴儿病人姓名获取方式
Create Or Replace Function zl_影像消息_XML内容获取
( 
    医嘱ID_In 病人医嘱记录.id%Type, 
    消息类型_In varchar2, 
    当前用户_In varchar2,
    消息标记_In varchar2:=Null         --新版：检查报告ID 
) Return varchar2 IS 
  v_Context varchar2(4000); 
  n_婴儿序号 病人医嘱记录.婴儿%Type;
  v_姓名 病人医嘱记录.姓名%Type;
  n_主页id   病人医嘱记录.主页id%Type;
 
  --ZLHIS_CIS_005(医技执行安排完成) 
  Function Get_Zlhis_Cis_005 Return varchar2 As 
    v_Return varchar2(4000); 
  Begin 
        Select 
          '<patient_info>' || 
             '<patient_id>' || a.病人id || '</patient_id>' || 
             '<patient_name>' || v_姓名 ||'</patient_name>' || 
          '</patient_info>' || 
          '<patient_clinic>' || 
             '<patient_source>' || b.病人来源 ||'</patient_source>' || 
             '<clinic_dept_id>' || b.病人科室id || '</clinic_dept_id>' || 
          '</patient_clinic>' || 
          '<patient_order>' || 
             '<order_id>' || c.医嘱id || '</order_id>' || 
             '<order_expiry>' || b.医嘱期效 ||'</order_expiry>' || 
             '<order_kind>' || b.诊疗类别 || '</order_kind>' || 
             '<operation_kind>' || d.操作类型 ||'</operation_kind>' || 
             '<order_item_id>' || c.医嘱id || '</order_item_id>' || 
             '<order_item_title>' || b.医嘱内容 ||'</order_item_title>' || 
          '</patient_order>' || 
          '<arrange_result>' || 
             '<arrange_time>' ||To_Char(c.安排时间,'yyyy/mm/dd hh24:mi:ss')|| '</arrange_time>' || 
          '</arrange_result>'   Into v_Return 
 
      From 病人信息 A, 病人医嘱记录 B, 病人医嘱发送 C, 诊疗项目目录 D 
      Where a.病人id = b.病人id And c.医嘱id = b.Id And b.诊疗项目id = d.Id And c.安排时间 Is Not Null And b.相关id Is Null And 
          b.诊疗类别 = 'D' And b.Id = 医嘱id_In; 
    Return v_Return; 
  End Get_Zlhis_Cis_005; 
 
  --ZLHIS_CIS_017(患者检查申请) 
  Function Get_ZLHIS_CIS_017 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(d.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.当前科室id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<check_request>' || 
               '<request_id>' || b.id || '</request_id>' || 
               '<check_item_id>' || b.诊疗项目id || '</check_item_id>' || 
               '<check_item_title>' || b.医嘱内容 || '</check_item_title>' || 
               '<execute_dept_id>' || nvl(c.执行部门id,0) || '</execute_dept_id>' || 
               '<send_serial>' || c.发送号 || '</send_serial>' || 
               '<bill_no>' || c.NO || '</bill_no>' || 
               '<bill_kind>' || c.记录性质 || '</bill_kind>' || 
               '<create_doctor>' || b.开嘱医生 || '</create_doctor>' || 
               '<create_time>' || b.开嘱时间 || '</create_time>' || 
               '<create_dept_id>' || nvl(b.开嘱科室id,0) || '</create_dept_id>' || 
           '</check_request>' into v_return 
 
        From 病人信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 病人挂号记录 d 
        Where a.病人id=b.病人id And b.id=c.医嘱id And b.挂号单=d.no(+) And b.相关ID Is Null 
              And a.病人id=b.病人id And b.id=医嘱ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_CIS_017; 
  
  --ZLHIS_CIS_015(医技拒绝执行) 
  Function Get_ZLHIS_CIS_015 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
       Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(e.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.当前科室id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<refuse_order>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<order_expiry>1</order_expiry>' || 
               '<order_kind>' || b.诊疗类别 || '</order_kind>' || 
               '<operation_kind>' || d.操作类型 || '</operation_kind>' || 
               '<order_item_id>' || b.诊疗项目ID || '</order_item_id>' || 
               '<order_item_title>' || d.名称 || '</order_item_title>' || 
           '</refuse_order>' Into v_return 
 
       From 病人信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 诊疗项目目录 d, 病人挂号记录 e 
       Where a.病人id=b.病人id And b.id=c.医嘱id And b.诊疗项目id=d.id And b.挂号单=e.no(+) And b.相关ID Is Null 
              And a.病人id=b.病人id And b.id=医嘱ID_In; 
              
       Return v_return;     
  End Get_ZLHIS_CIS_015;   
   
 
  --ZLHIS_CIS_024(患者医嘱撤销) 
  Function Get_ZLHIS_CIS_024 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
       Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(e.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.当前科室id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<cancel_order>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<order_kind>' || b.诊疗类别 || '</order_kind>' || 
               '<operation_kind>' || d.操作类型 || '</operation_kind>' || 
           '</cancel_order>' Into v_return 
 
       From 病人信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 诊疗项目目录 d, 病人挂号记录 e 
       Where a.病人id=b.病人id And b.id=c.医嘱id And b.诊疗项目id=d.id And b.挂号单=e.no(+) And b.相关ID Is Null 
              And a.病人id=b.病人id And b.id=医嘱ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_CIS_024; 
 
 
  --ZLHIS_PACS_001(检查报告完成) 
  Function Get_ZLHIS_PACS_001 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
    If 消息标记_In Is Null Then
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(e.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.当前科室id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<advice_info>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<check_item_id>' || b.诊疗项目id || '</check_item_id>' || 
               '<check_item_title>' || b.医嘱内容 || '</check_item_title>' || 
               '<create_doctor>' || b.开嘱医生 || '</create_doctor>' || 
               '<create_dept_id>' || nvl(b.开嘱科室id, 0) || '</create_dept_id>' || 
               '<study_execute_id>' || nvl(b.执行科室id,0) || '</study_execute_id>' || 
           '</advice_info>' || 
           '<report_inf>' || 
               '<report_id>' || d.病历id || '</report_id>' || 
               '<report_doctor>' || c.完成人 || '</report_doctor>' || 
               '<result_positive>' || c.结果阳性 || '</result_positive>' || 
           '</report_inf>'  Into v_return 
 
        From 病人信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 病人医嘱报告 d, 病人挂号记录 e 
        Where a.病人id=b.病人id And b.id=c.医嘱id And b.id=d.医嘱id And b.挂号单=e.no(+) And b.相关ID Is Null 
              And d.检查报告id Is Null And a.病人id=b.病人id And b.id=医嘱ID_In; 
    Else
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(e.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.当前科室id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<advice_info>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<check_item_id>' || b.诊疗项目id || '</check_item_id>' || 
               '<check_item_title>' || b.医嘱内容 || '</check_item_title>' || 
               '<create_doctor>' || b.开嘱医生 || '</create_doctor>' || 
               '<create_dept_id>' || nvl(b.开嘱科室id, 0) || '</create_dept_id>' || 
               '<study_execute_id>' || nvl(b.执行科室id,0) || '</study_execute_id>' || 
           '</advice_info>' || 
           '<report_inf>' || 
               '<report_id>' || d.检查报告id || '</report_id>' || 
               '<report_doctor>' || c.完成人 || '</report_doctor>' || 
               '<result_positive>' || c.结果阳性 || '</result_positive>' || 
           '</report_inf>'  Into v_return 
 
        From 病人信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 病人医嘱报告 d, 病人挂号记录 e 
        Where a.病人id=b.病人id And b.id=c.医嘱id And b.id=d.医嘱id And b.挂号单=e.no(+) And b.相关ID Is Null 
              And d.病历id Is Null And a.病人id=b.病人id And b.id=医嘱ID_In And d.检查报告id=消息标记_In; 
    End If;
    
    Return v_return; 
  End Get_ZLHIS_PACS_001; 
 
  --ZLHIS_PACS_002(患者状态同步) 
  Function Get_ZLHIS_PACS_002 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<identity_card>' || a.身份证号 || '</identity_card>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(d.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.当前科室id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<study_state>' || 
               '<study_cur_state>' || nvl(c.执行过程,0) || '</study_cur_state>' || 
               '<study_cur_time>' || sysdate || '</study_cur_time>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<study_item_id>' || b.诊疗项目id || '</study_item_id>' || 
               '<study_item_title>' || b.医嘱内容 || '</study_item_title>' || 
               '<study_oper_person>' || 当前用户_In || '</study_oper_person>' || 
               '<study_execute_id>' || nvl(b.执行科室id,0) || '</study_execute_id>' || 
           '</study_state>' Into v_return 
 
        From 病人信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 病人挂号记录 d 
        Where a.病人id=b.病人id And b.id=c.医嘱id And b.挂号单=d.no(+) And b.相关ID Is Null 
              And a.病人id=b.病人id And b.id=医嘱ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_PACS_002; 
 
 
  --ZLHIS_PACS_003(检查状态回退) 
  Function Get_ZLHIS_PACS_003 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<identity_card>' || a.身份证号 || '</identity_card>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(d.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.当前科室id, 0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<study_state>' || 
               '<study_cur_state>' || nvl(c.执行过程,0) || '</study_cur_state>' || 
               '<study_cur_time>' || Sysdate || '</study_cur_time>' || 
               '<study_order_id>' || b.id || '</study_order_id>' || 
               '<study_item_id>' || b.诊疗项目id || '</study_item_id>' || 
               '<study_item_title>' || b.医嘱内容 || '</study_item_title>' || 
               '<study_oper_person>' || 当前用户_In || '</study_oper_person>' || 
               '<study_execute_id>' || nvl(b.执行科室id,0) || '</study_execute_id>' || 
           '</study_state>' Into v_return 
 
        From 病人信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 病人挂号记录 d 
        Where a.病人id=b.病人id And b.id=c.医嘱id And b.挂号单=d.no(+) And b.相关ID Is Null 
              And a.病人id=b.病人id And b.id=医嘱ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_PACS_003; 
 
 
  --ZLHIS_PACS_004(检查报告撤销) 
  Function Get_ZLHIS_PACS_004 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
    If 消息标记_In Is Null Then
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(e.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.当前科室id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<advice_info>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<cur_state>' || nvl(c.执行过程,0) || '</cur_state>' || 
               '<check_item_id>' || b.诊疗项目id || '</check_item_id>' || 
               '<check_item_title>' || b.医嘱内容 || '</check_item_title>' || 
               '<create_doctor>' || b.开嘱医生 || '</create_doctor>' || 
               '<create_dept_id>' || nvl(b.开嘱科室id,0) || '</create_dept_id>' || 
               '<study_execute_id>' || nvl(b.执行科室id,0) || '</study_execute_id>' || 
           '</advice_info>' || 
           '<report_info>' || 
               '<report_id>' || d.病历id || '</report_id>' || 
           '</report_info>'  Into v_return 
 
        From 病人信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 病人医嘱报告 d, 病人挂号记录 e 
        Where a.病人id=b.病人id And b.id=c.医嘱id And b.id=d.医嘱id And b.挂号单=e.no(+) And b.相关ID Is Null 
              And d.检查报告id Is Null And a.病人id=b.病人id And b.id=医嘱ID_In; 
    Else
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(e.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.当前科室id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<advice_info>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<cur_state>' || nvl(c.执行过程,0) || '</cur_state>' || 
               '<check_item_id>' || b.诊疗项目id || '</check_item_id>' || 
               '<check_item_title>' || b.医嘱内容 || '</check_item_title>' || 
               '<create_doctor>' || b.开嘱医生 || '</create_doctor>' || 
               '<create_dept_id>' || nvl(b.开嘱科室id,0) || '</create_dept_id>' || 
               '<study_execute_id>' || nvl(b.执行科室id,0) || '</study_execute_id>' || 
           '</advice_info>' || 
           '<report_info>' || 
               '<report_id>' || d.检查报告id || '</report_id>' || 
           '</report_info>'  Into v_return 
 
        From 病人信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 病人医嘱报告 d, 病人挂号记录 e 
        Where a.病人id=b.病人id And b.id=c.医嘱id And b.id=d.医嘱id And b.挂号单=e.no(+) And b.相关ID Is Null 
              And d.病历id Is Null And a.病人id=b.病人id And b.id=医嘱ID_In And d.检查报告id=消息标记_In; 
    End If;
    
    Return v_return; 
  End Get_ZLHIS_PACS_004; 
  
  --ZLHIS_PACS_005(检查危急值通知) 
  Function Get_ZLHIS_PACS_005 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        With t As (Select id, 姓名 From 人员表) 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(i.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_area_id>' || nvl(a.当前病区id,0) || '</clinic_area_id>' || 
               '<clinic_dept_id>' || nvl(b.开嘱科室id,0) || '</clinic_dept_id>' || 
               '<in_doctor_id>' || nvl(e.id,0) || '</in_doctor_id>' || 
               '<director_doctor_id>' || nvl(g.id,0) || '</director_doctor_id>' || 
               '<treat_doctor_id>' || nvl(h.id,0) || '</treat_doctor_id>' || 
               '<duty_nurse_id>' || nvl(f.id,0) || '</duty_nurse_id>' || 
           '</patient_clinic>' || 
           '<check_order>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<check_item_id>' || b.诊疗项目id || '</check_item_id>' || 
               '<check_item_title>' || b.医嘱内容 || '</check_item_title>' || 
               '<study_execute_id>' || nvl(b.执行科室id,0) || '</study_execute_id>' || 
           '</check_order>'  Into v_return 
 
        From 病人信息 a, 病人医嘱记录 b, 病案主页 c, 病人挂号记录 i, 
             (Select a.病人ID,a.主页ID,主任医师,主治医师 From 病人变动记录 a,病人医嘱记录 b 
              Where a.病人id=b.病人id And a.开始原因 Is Not Null And a.终止时间 Is Null And b.id=0) d, 
              t e, t f, t g, t h 
        Where b.病人id = a.病人id And b.挂号单=i.no(+) 
              And b.病人id=c.病人id(+) And b.主页id=c.主页id(+) 
              And c.住院医师=e.姓名(+) And c.责任护士=f.姓名(+) 
              And c.病人id =d.病人id(+) And c.主页id=d.主页id(+) 
              And d.主任医师=g.姓名(+) And d.主治医师=h.姓名(+) 
              And b.相关id Is Null And a.病人id=b.病人id And b.id=医嘱ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_PACS_005;
  
  --ZLHIS_PACS_006(检查预约通知) 
  Function Get_ZLHIS_PACS_006 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.病人id || '</patient_id>' || 
               '<patient_name>' || v_姓名 || '</patient_name>' || 
               '<in_number>' || nvl(a.住院号,0) || '</in_number>' || 
               '<out_number>' || nvl(a.门诊号,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.病人来源 || '</patient_source>' || 
               '<clinic_id>' || Case b.病人来源 When 1 Then nvl(i.id,0) When 2 Then nvl(b.主页id, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_area_id>' || nvl(a.当前病区id,0) || '</clinic_area_id>' || 
               '<clinic_dept_id>' || nvl(b.开嘱科室id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<check_order>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<check_item_id>' || b.诊疗项目id || '</check_item_id>' || 
               '<check_item_title>' || b.医嘱内容 || '</check_item_title>' || 
               '<study_execute_id>' || nvl(b.执行科室id,0) || '</study_execute_id>' ||
               '<schedult_id>' || k.预约id || '</schedult_id>' ||
               '<machine_id>' || nvl(k.检查设备id,0) || '</machine_id>' ||
               '<machine_name>' || nvl(k.检查设备名称,'') || '</machine_name>' ||
               '<schedule_date>' || To_Char(k.预约日期, 'YYYY-MM-DD HH24:MI:SS') || '</schedule_date>' ||
               '<schedule_begin_time>' || To_Char(k.预约开始时间, 'YYYY-MM-DD HH24:MI:SS') || '</schedule_begin_time>' ||
               '<schedule_end_time>' || To_Char(k.预约结束时间, 'YYYY-MM-DD HH24:MI:SS')  || '</schedule_end_time>' ||
               '<schedule_sec_begin>' || To_Char(k.预约开始时间段, 'YYYY-MM-DD HH24:MI:SS')  || '</schedule_sec_begin>' ||
               '<schedule_sec_end>' || To_Char(k.预约结束时间段, 'YYYY-MM-DD HH24:MI:SS')  || '</schedule_sec_end>' ||
               '<schedule_call_no>' || nvl(k.序号,0) || '</schedule_call_no>' || 
           '</check_order>'  Into v_return 
 
        From 病人信息 a, 病人医嘱记录 b, 病案主页 c, 病人挂号记录 i, Ris检查预约 k 
        Where b.病人id = a.病人id And b.挂号单=i.no(+) 
              And b.病人id=c.病人id(+) And b.主页id=c.主页id(+) 
              And b.id =k.医嘱id 
              And b.相关id Is Null And a.病人id=b.病人id And b.id=医嘱ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_PACS_006; 
 
Begin 
  v_Context := ''; 
  
  --首先判断是否是婴儿，若是则v_姓名 提取婴儿姓名，否则v_姓名 为病人医嘱记录的姓名。
  Select Max(婴儿), Max(主页id) Into n_婴儿序号, n_主页id From 病人医嘱记录 Where ID = 医嘱id_In;
  If n_婴儿序号 > 0 And n_主页id > 0 Then
    Select Nvl(b.婴儿姓名, a.姓名 || '之子' || Trim(To_Char(b.序号, '9')))
    Into v_姓名
    From 病人医嘱记录 A, 病人新生儿记录 B
    Where a.病人id = b.病人id And b.主页id = n_主页id And b.序号 = n_婴儿序号 And a.Id = 医嘱id_In;
  Else
    Select 姓名 Into v_姓名 From 病人医嘱记录
    Where Id = 医嘱id_In;
  End If;
 
  Case 消息类型_In 
    When 'ZLHIS_CIS_005' Then 
      --ZLHIS_CIS_005(医技执行安排完成) 
      v_Context := Get_ZLHIS_CIS_005; 
 
    When 'ZLHIS_CIS_015' Then 
        --ZLHIS_CIS_015(医技拒绝执行) 
        v_Context := Get_ZLHIS_CIS_015; 
        
    When 'ZLHIS_CIS_017' Then 
        --ZLHIS_CIS_017(患者检查申请) 
        v_Context := Get_ZLHIS_CIS_017; 
 
    When 'ZLHIS_CIS_024' Then 
        --ZLHIS_PACS_024(患者医嘱撤销) 
        v_Context := Get_ZLHIS_CIS_024; 
 
    When 'ZLHIS_PACS_001' Then 
        --ZLHIS_PACS_001(检查报告完成) 
        v_Context := Get_ZLHIS_PACS_001; 
 
    When 'ZLHIS_PACS_002' Then 
        --ZLHIS_PACS_002(检查状态同步) 
        v_Context := Get_ZLHIS_PACS_002; 
 
    When 'ZLHIS_PACS_003' Then 
        --ZLHIS_PACS_003(检查状态回退) 
        v_Context := Get_ZLHIS_PACS_003; 
 
    When 'ZLHIS_PACS_004' Then 
        --ZLHIS_PACS_004(检查报告撤销) 
        v_Context := Get_ZLHIS_PACS_004; 
 
    When 'ZLHIS_PACS_005' Then 
        --ZLHIS_PACS_005(检查危急值通知) 
        v_Context := Get_ZLHIS_PACS_005; 
    
    When 'ZLHIS_PACS_006' Then 
        --ZLHIS_PACS_006(检查预约通知) 
        v_Context := Get_ZLHIS_PACS_006; 
    Else 
      Return ''; 
  End Case; 
 
  Return v_Context; 
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_影像消息_XML内容获取;
/

--117999:李南春,2017-12-12,序号锁号状态检查
Create Or Replace Procedure Zl_三方机构挂号_Insert
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
      Select Nvl(是否分时段, 0), 限号数, 已挂数, 其中已接收, 已约数, 是否序号控制, 限约数, 项目id, 科室id, 医生id, 医生姓名, 替诊医生id, 替诊医生姓名, 替诊开始时间, 替诊终止时间
      Into n_启用分时段, n_限号数, n_已挂数, n_其中已接收, n_已约数, n_序号控制, n_限约数, n_项目id, n_科室id, n_医生id, v_医生姓名, n_替诊医生id, v_替诊医生姓名,
           d_替诊开始时间, d_替诊终止时间
      From 临床出诊记录
      Where ID = 记录id_In And Nvl(是否锁定, 0) = 0;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '该号别没有在' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '中进行安排。';
      Raise Err_Item;
    End If;
  
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
        Select To_Number(Substr(Zl_Fun_Get临床出诊预约状态(记录id_In, 发生时间_In, 号序_In, 预约方式_In, NULL, 0, v_操作员姓名, v_机器名), 1, 1))
        Into n_Exists
        From Dual;
        If n_Exists <> 0 Then
          v_Err_Msg := '传入的预约方式' || 预约方式_In || '不可用,不能继续。';
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
      Select Nvl(控制方式, 0)
      Into n_合作单位限数量模式
      From 临床出诊挂号控制记录
      Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1 And Rownum < 2;
    
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

--117999:李南春,2017-12-13,序号锁号状态检查
CREATE OR REPLACE Function Zl_Fun_Get临床出诊预约状态
(
  记录id_In   In 临床出诊记录.Id%Type,
  预约时间_In In 病人挂号记录.预约时间%Type,
  序号_In     临床出诊序号控制.序号%Type := Null,
  预约方式_In 预约方式.名称%Type := Null,
  合作单位_In 挂号合作单位.名称%Type := Null,
  收费预约_In Number := 0,
  操作员姓名_In 挂号序号状态.操作员姓名%Type := Null,
  机器名_In   挂号序号状态.机器名%Type := Null
) Return Varchar2 As
  --功能：判断出诊记录在预约时间是否可预约
  --入参：
  --返回：
  --     格式：预约状态|提示信息，如："1|预约时间不在当前上班时段时间范围内。"
  --     预约状态：
  --         0-可预约
  --         ======================================================
  --         1-不可预约，预约时间不在当前上班时段时间范围内
  --         2-不可预约，当前上班时段禁止预约
  --         3-不可预约，当前上班时段在预约时间时已停诊
  --         4-不可预约，当前上班时段剩余可预约数为零
  --         ======================================================
  --         5-不可预约，当前预约时间在法定节假日时间范围内，不上班
  --         6-不可预约，当前预约时间在法定节假日时间范围内，禁止预约
  --         7-不可预约，当前预约时间在法定节假日不允许预约的时间范围内
  --         8-不可预约，当前预约时间在法定节假日不允许挂号的时间范围内
  --         9-不可预约，当前预约时间在法定节假日时间范围内，已停诊
  --         ======================================================
  --         10-不可预约，当前预约方式禁止预约
  --         11-不可预约，当前预约方式可预约数不足
  --         ======================================================
  --         12-不可预约，当前合作单位禁止预约
  --         13-不可预约，当前合作单位可预约数不足
  --         ======================================================
  --         14-不可预约，当前序号禁止预约
  --         15-不可预约，当前序号已经被使用
  --         16-不可预约，当前序号不可用
  --
  n_号源id         临床出诊记录.号源id%Type;
  n_是否分时段     临床出诊记录.是否分时段%Type;
  n_预约控制       临床出诊记录.预约控制%Type;
  d_停诊开始时间   临床出诊记录.停诊开始时间%Type;
  d_停诊终止时间   临床出诊记录.停诊终止时间%Type;
  v_停诊原因       临床出诊记录.停诊原因%Type;
  n_限约数         临床出诊记录.限约数%Type;
  n_已约数         临床出诊记录.已约数%Type;
  n_独占           临床出诊记录.是否独占%Type;
  n_控制方式       临床出诊挂号控制记录.控制方式%Type;
  n_数量           临床出诊挂号控制记录.数量%Type;
  n_数量限制       临床出诊挂号控制记录.数量%Type;
  n_序号控制       临床出诊记录.是否序号控制%Type;
  v_预约方式       临床出诊挂号控制记录.名称%Type;
  n_类型           临床出诊挂号控制记录.类型%Type;
  n_预约方式限约数 临床出诊记录.限约数%Type;
  n_预约方式已约数 临床出诊记录.已约数%Type;
  n_挂号状态       临床出诊序号控制.挂号状态%Type;
  n_是否预约       临床出诊序号控制.是否预约%Type;

  n_假日控制状态 临床出诊号源.假日控制状态%Type;

  v_允许预约 法定假日表.允许预约日期%Type;
  v_允许挂号 法定假日表.允许挂号日期%Type;
  n_Count    Number(2);
  n_已使用   Number(5);
  v_锁号机器名 挂号序号状态.机器名%Type;
  v_锁号操作员 挂号序号状态.操作员姓名%Type;
Begin
  Begin
    Select a.号源id, a.是否分时段, a.预约控制, a.停诊开始时间, a.停诊终止时间, a.停诊原因, Nvl(限约数, 限号数), 已约数, 是否独占, 是否序号控制
    Into n_号源id, n_是否分时段, n_预约控制, d_停诊开始时间, d_停诊终止时间, v_停诊原因, n_限约数, n_已约数, n_独占, n_序号控制
    From 临床出诊记录 A
    Where a.Id = 记录id_In And 预约时间_In Between 开始时间 And 终止时间;
  Exception
    When Others Then
      Return '1|预约时间不在当前上班时段时间范围内。';
  End;

  --预约方式检查
  If 预约方式_In Is Not Null Then
    Begin
      Select 控制方式
      Into n_控制方式
      From 临床出诊挂号控制记录
      Where 类型 = 2 And 性质 = 1 And 记录id = 记录id_In And 名称 = 预约方式_In And Rownum < 2;
    Exception
      When Others Then
        Begin
          Select 控制方式
          Into n_控制方式
          From 临床出诊挂号控制记录
          Where 类型 = 2 And 性质 = 1 And 记录id = 记录id_In And Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
    End;
    If n_控制方式 = 0 Then
      Return '10|当前预约方式禁止预约。';
    End If;
    If n_控制方式 = 1 Or n_控制方式 = 2 Then
      Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
      If n_独占 = 0 Then
        Begin
          Select 数量
          Into n_数量
          From 临床出诊挂号控制记录
          Where 类型 = 2 And 性质 = 1 And 名称 = 预约方式_In And 记录id = 记录id_In;
        Exception
          When Others Then
            n_数量 := Null;
        End;
        If n_数量 Is Not Null Then
          If n_控制方式 = 1 Then
            n_预约方式限约数 := Round(n_预约方式限约数 * n_数量 / 100);
          Else
            n_预约方式限约数 := n_数量;
          End If;
          Select Count(1)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 记录状态 = 1 And 预约方式 = 预约方式_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '11|当前预约方式可预约数不足。';
          End If;
        End If;
      Else
        --限数量独占
        Begin
          Select 数量
          Into n_数量
          From 临床出诊挂号控制记录
          Where 类型 = 2 And 性质 = 1 And 名称 = 预约方式_In And 记录id = 记录id_In;
        Exception
          When Others Then
            n_数量 := Null;
        End;
        If n_数量 Is Not Null Then
          If n_控制方式 = 1 Then
            n_预约方式限约数 := Round(n_预约方式限约数 * n_数量 / 100);
          Else
            n_预约方式限约数 := n_数量;
          End If;
          Select Count(1)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 记录状态 = 1 And 预约方式 = 预约方式_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '11|当前预约方式可预约数不足。';
          End If;
        Else
          If 收费预约_In = 0 Then
            For r_限制 In (Select 数量, 名称, 类型 From 临床出诊挂号控制记录 Where 性质 = 1 And 记录id = 记录id_In) Loop
              If r_限制.类型 = 1 Then
                Select Count(1)
                Into n_已使用
                From 病人挂号记录
                Where 出诊记录id = 记录id_In And 合作单位 = r_限制.名称 And 记录状态 = 1;
              Else
                Select Count(1)
                Into n_已使用
                From 病人挂号记录
                Where 出诊记录id = 记录id_In And 预约方式 = r_限制.名称 And 记录状态 = 1;
              End If;
              If n_控制方式 = 1 Then
                n_数量限制 := Nvl(n_数量限制, 0) + Round(r_限制.数量 * n_预约方式限约数 / 100) - Nvl(n_已使用, 0);
              Else
                n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
              End If;
            End Loop;
            Select Count(1) Into n_已使用 From 病人挂号记录 Where 出诊记录id = 记录id_In And 记录状态 = 1;
            If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
              Null;
            Else
              Return '11|当前预约方式可预约数不足。';
            End If;
          Else
            For r_限制 In (Select 数量, 名称, 类型
                         From 临床出诊挂号控制记录
                         Where 性质 = 1 And 类型 = 2 And 记录id = 记录id_In) Loop
              Select Count(1)
              Into n_已使用
              From 病人挂号记录
              Where 出诊记录id = 记录id_In And 预约方式 = r_限制.名称 And 记录状态 = 1;
              If n_控制方式 = 1 Then
                n_数量限制 := Nvl(n_数量限制, 0) + Round(r_限制.数量 * n_预约方式限约数 / 100) - Nvl(n_已使用, 0);
              Else
                n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
              End If;
            End Loop;
            Select Count(1) Into n_已使用 From 病人挂号记录 Where 出诊记录id = 记录id_In And 记录状态 = 1;
            If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
              Null;
            Else
              Return '11|当前预约方式可预约数不足。';
            End If;
          End If;
        End If;
      End If;
    End If;
    If n_控制方式 = 3 Then
      If n_序号控制 = 1 Then
        If 收费预约_In = 0 Then
          Begin
            Select 数量, 名称, 类型
            Into n_预约方式限约数, v_预约方式, n_类型
            From 临床出诊挂号控制记录
            Where 性质 = 1 And 记录id = 记录id_In And 序号 = 序号_In;
          Exception
            When Others Then
              n_预约方式限约数 := Null;
          End;
          If n_预约方式限约数 Is Not Null Then
            If v_预约方式 <> 预约方式_In Or n_类型 = 1 Then
              Return '11|当前预约方式可预约数不足。';
            End If;
            Select Nvl(Max(1), 0)
            Into n_预约方式已约数
            From 病人挂号记录
            Where 出诊记录id = 记录id_In And 号序 = 序号_In;
            If n_预约方式已约数 >= n_预约方式限约数 Then
              Return '11|当前预约方式可预约数不足。';
            End If;
          End If;
        Else
          Begin
            Select 数量, 名称, 类型
            Into n_预约方式限约数, v_预约方式, n_类型
            From 临床出诊挂号控制记录
            Where 性质 = 1 And 类型 = 2 And 记录id = 记录id_In And 序号 = 序号_In;
          Exception
            When Others Then
              n_预约方式限约数 := Null;
          End;
          If n_预约方式限约数 Is Not Null Then
            If v_预约方式 <> 预约方式_In Then
              Return '11|当前预约方式可预约数不足。';
            End If;
            Select Nvl(Max(1), 0)
            Into n_预约方式已约数
            From 病人挂号记录
            Where 出诊记录id = 记录id_In And 号序 = 序号_In;
            If n_预约方式已约数 >= n_预约方式限约数 Then
              Return '11|当前预约方式可预约数不足。';
            End If;
          End If;
        End If;
      Else
        If 收费预约_In = 0 Then
          For r_限制 In (Select 数量, 名称, 类型
                       From 临床出诊挂号控制记录
                       Where 性质 = 1 And 记录id = 记录id_In And 序号 = 序号_In) Loop
            If r_限制.名称 <> 预约方式_In Or r_限制.类型 = 1 Then
              If r_限制.类型 = 1 Then
                Select Count(1)
                Into n_已使用
                From 临床出诊序号控制 A, 病人挂号记录 B
                Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                      b.合作单位 = r_限制.名称 And b.记录状态 = 1;
              Else
                Select Count(1)
                Into n_已使用
                From 临床出诊序号控制 A, 病人挂号记录 B
                Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                      b.预约方式 = r_限制.名称 And b.记录状态 = 1;
              End If;
              n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
            Else
              Select Count(1)
              Into n_预约方式已约数
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.预约方式 = 预约方式_In And b.记录状态 = 1;
              If n_预约方式已约数 >= n_预约方式限约数 Then
                Return '11|当前预约方式可预约数不足。';
              End If;
            End If;
          End Loop;
          Select Count(1)
          Into n_已使用
          From 临床出诊序号控制 A
          Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And 序号 = 序号_In;
          Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
          If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
            Null;
          Else
            Return '11|当前预约方式可预约数不足。';
          End If;
        Else
          For r_限制 In (Select 数量, 名称, 类型
                       From 临床出诊挂号控制记录
                       Where 性质 = 1 And 类型 = 2 And 记录id = 记录id_In And 序号 = 序号_In) Loop
            If r_限制.名称 <> 预约方式_In Then
              Select Count(1)
              Into n_已使用
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.预约方式 = r_限制.名称 And b.记录状态 = 1;
              n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
            Else
              Select Count(1)
              Into n_预约方式已约数
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.预约方式 = 预约方式_In And b.记录状态 = 1;
              If n_预约方式已约数 >= n_预约方式限约数 Then
                Return '11|当前预约方式可预约数不足。';
              End If;
            End If;
          End Loop;
          Select Count(1)
          Into n_已使用
          From 临床出诊序号控制 A
          Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And 序号 = 序号_In;
          Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
          If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
            Null;
          Else
            Return '11|当前预约方式可预约数不足。';
          End If;
        End If;
      End If;
    End If;
  End If;

  --合作单位检查
  If 合作单位_In Is Not Null Then
    Begin
      Select 控制方式
      Into n_控制方式
      From 临床出诊挂号控制记录
      Where 类型 = 1 And 性质 = 1 And 记录id = 记录id_In And 名称 = 合作单位_In And Rownum < 2;
    Exception
      When Others Then
        Begin
          Select 控制方式
          Into n_控制方式
          From 临床出诊挂号控制记录
          Where 类型 = 1 And 性质 = 1 And 记录id = 记录id_In And Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
    End;
    If n_控制方式 = 0 Then
      Return '12|当前合作单位禁止预约。';
    End If;
    If n_控制方式 = 1 Or n_控制方式 = 2 Then
      Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
      If n_独占 = 0 Then
        Begin
          Select 数量
          Into n_数量
          From 临床出诊挂号控制记录
          Where 类型 = 1 And 性质 = 1 And 名称 = 合作单位_In And 记录id = 记录id_In;
        Exception
          When Others Then
            n_数量 := Null;
        End;
        If n_数量 Is Not Null Then
          If n_控制方式 = 1 Then
            n_预约方式限约数 := Round(n_预约方式限约数 * n_数量 / 100);
          Else
            n_预约方式限约数 := n_数量;
          End If;
          Select Count(1)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 记录状态 = 1 And 合作单位 = 合作单位_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '13|当前合作单位可预约数不足。';
          End If;
        End If;
      Else
        --限数量独占
        Begin
          Select 数量
          Into n_数量
          From 临床出诊挂号控制记录
          Where 类型 = 1 And 性质 = 1 And 名称 = 合作单位_In And 记录id = 记录id_In;
        Exception
          When Others Then
            n_数量 := Null;
        End;
        If n_数量 Is Not Null Then
          If n_控制方式 = 1 Then
            n_预约方式限约数 := Round(n_预约方式限约数 * n_数量 / 100);
          Else
            n_预约方式限约数 := n_数量;
          End If;
          Select Count(1)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 记录状态 = 1 And 合作单位 = 合作单位_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '13|当前合作单位可预约数不足。';
          End If;
        Else
          For r_限制 In (Select 数量, 名称, 类型 From 临床出诊挂号控制记录 Where 性质 = 1 And 记录id = 记录id_In) Loop
            If r_限制.类型 = 1 Then
              Select Count(1)
              Into n_已使用
              From 病人挂号记录
              Where 出诊记录id = 记录id_In And 合作单位 = r_限制.名称 And 记录状态 = 1;
            Else
              Select Count(1)
              Into n_已使用
              From 病人挂号记录
              Where 出诊记录id = 记录id_In And 预约方式 = r_限制.名称 And 记录状态 = 1;
            End If;
            If n_控制方式 = 1 Then
              n_数量限制 := Nvl(n_数量限制, 0) + Round(r_限制.数量 * n_预约方式限约数 / 100) - Nvl(n_已使用, 0);
            Else
              n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
            End If;
          End Loop;
          Select Count(1) Into n_已使用 From 病人挂号记录 Where 出诊记录id = 记录id_In And 记录状态 = 1;
          If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
            Null;
          Else
            Return '13|当前合作单位可预约数不足。';
          End If;
        End If;
      End If;
    End If;
    If n_控制方式 = 3 Then
      If n_序号控制 = 1 Then
        Begin
          Select 数量, 名称, 类型
          Into n_预约方式限约数, v_预约方式, n_类型
          From 临床出诊挂号控制记录
          Where 性质 = 1 And 记录id = 记录id_In And 序号 = 序号_In;
        Exception
          When Others Then
            n_预约方式限约数 := Null;
        End;
        If n_预约方式限约数 Is Not Null Then
          If v_预约方式 <> 合作单位_In Or n_类型 = 1 Then
            Return '13|当前合作单位可预约数不足。';
          End If;
          Select Nvl(Max(1), 0)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 号序 = 序号_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '13|当前合作单位可预约数不足。';
          End If;
        End If;
      Else
        For r_限制 In (Select 数量, 名称, 类型
                     From 临床出诊挂号控制记录
                     Where 性质 = 1 And 记录id = 记录id_In And 序号 = 序号_In) Loop
          If r_限制.名称 <> 合作单位_In Or r_限制.类型 = 1 Then
            If r_限制.类型 = 1 Then
              Select Count(1)
              Into n_已使用
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.合作单位 = r_限制.名称 And b.记录状态 = 1;
            Else
              Select Count(1)
              Into n_已使用
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.预约方式 = r_限制.名称 And b.记录状态 = 1;
            End If;
            n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
          Else
            Select Count(1)
            Into n_预约方式已约数
            From 临床出诊序号控制 A, 病人挂号记录 B
            Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And b.合作单位 = 合作单位_In And
                  b.记录状态 = 1;
            If n_预约方式已约数 >= n_预约方式限约数 Then
              Return '13|当前合作单位可预约数不足。';
            End If;
          End If;
        End Loop;
        Select Count(1)
        Into n_已使用
        From 临床出诊序号控制 A
        Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And 序号 = 序号_In;
        Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
        If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
          Null;
        Else
          Return '13|当前合作单位可预约数不足。';
        End If;
      End If;
    End If;
  End If;

  --0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
  If Nvl(n_预约控制, 0) = 1 Then
    Return '2|当前上班时段禁止预约。';
  End If;

  If d_停诊开始时间 Is Not Null And Not (Nvl(n_序号控制, 0) = 1 And Nvl(n_是否分时段, 0) = 1) Then
    If 预约时间_In >= d_停诊开始时间 And 预约时间_In <= d_停诊终止时间 Then
      Return '3|当前上班时段在预约时间时已停诊，不能预约！';
    End If;
  End If;

  If Nvl(n_限约数, 0) > 0 Then
    If Nvl(n_限约数, 0) - Nvl(n_已约数, 0) <= 0 Then
      Return '4|当前上班时段剩余可预约数为零，不能继续预约！';
    End If;
  End If;

  If Nvl(n_是否分时段, 0) = 0 Then
    --不分时段
    Begin
      Select Nvl(b.假日控制状态, 0) Into n_假日控制状态 From 临床出诊号源 B Where b.Id = n_号源id;
    Exception
      When Others Then
        n_假日控制状态 := 0;
    End;

    --1.查找包含预约时间的节假日
    Begin
      Select a.允许预约日期, a.允许挂号日期
      Into v_允许预约, v_允许挂号
      From 法定假日表 A
      Where a.性质 = 0 And 预约时间_In Between a.开始日期 And a.终止日期 + 1 - 1 / 24 / 60 / 60 And Rownum < 2;
    Exception
      When Others Then
        Return '0|正常预约。';
    End;

    --假日控制状态：0-不上班;1-上班且开放预约;2-上班但不开放预约;3-受节假日设置控制
    If Nvl(n_假日控制状态, 0) = 0 Then
      --不上班的肯定是不能预约的
      Return '5|当前预约时间在法定节假日时间范围内，不上班。';
    Elsif Nvl(n_假日控制状态, 0) = 1 Then
      Return '0|正常预约。';
    Elsif Nvl(n_假日控制状态, 0) = 2 Then
      --在节假日时间范围内，则不能预约
      Return '6|当前预约时间在法定节假日时间范围内，禁止预约。';
    Elsif Nvl(n_假日控制状态, 0) = 3 Then
      --没有"允许挂号"就一定没有"允许预约"
      If v_允许挂号 Is Not Null Then
        --2.检查是否有包含预约时间的"允许挂号"
        Select Max(1)
        Into n_Count
        From Table(f_Str2list(v_允许挂号, ';'))
        Where 预约时间_In Between To_Date(Column_Value, 'yyyy-mm-dd') And
              To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And Rownum < 2;

        If Nvl(n_Count, 0) <> 0 Then
          --3.检查是否有包含预约时间的"允许预约"
          Select Max(1)
          Into n_Count
          From Table(f_Str2list(v_允许预约, ';'))
          Where 预约时间_In Between To_Date(Column_Value, 'yyyy-mm-dd') And
                To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And Rownum < 2;

          If Nvl(n_Count, 0) = 0 Then
            --不在"允许预约"时间范围内，则不能预约
            Return '7|当前预约时间在法定节假日不允许预约的时间范围内，不能预约。';
          Else
            Return '0|正常预约。';
          End If;
        Else
          Return '8|当前预约时间在法定节假日不允许挂号的时间范围内，不能预约。';
        End If;
      Else
        --没有设置"允许挂号"/"允许预约"表示停诊，肯定不能预约
        Return '9|当前预约时间在法定节假日时间范围内，已停诊，不能预约。';
      End If;
    End If;
  Else
    --分时段
    If Nvl(序号_In, 0) <> 0 Then
      Begin
        Select Nvl(是否预约, 0), Nvl(挂号状态, 0), 操作员姓名, 工作站名称
        Into n_是否预约, n_挂号状态, v_锁号操作员, v_锁号机器名
        From 临床出诊序号控制
        Where 记录id = 记录id_In And 序号 = 序号_In;
      Exception
        When Others Then
          Return '16|当前选择的序号不可用。';
      End;
      If n_是否预约 = 0 Then
        Return '14|当前选择的序号禁止预约。';
      End If;
      If n_挂号状态 <> 0 Then
        If n_挂号状态 = 5 And (Nvl(操作员姓名_In, '-') <> Nvl(v_锁号操作员, '_') Or Nvl(机器名_In, '-') <> Nvl(v_锁号机器名, '_')) Then
           Return '15|当前选择的序号已经被'|| Nvl(v_锁号机器名,'') ||'锁定。';
        Elsif n_挂号状态 <> 5 Then
           Return '15|当前选择的序号已经被使用。';
        End if;
      End If;
    End If;
    Return '0|正常预约。';
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Get临床出诊预约状态;
/

--115264:冉俊明,2017-12-11,住院病人按门诊收费，直接收费，病人无门诊号时，门诊标志和标识号填写错误
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
  n_门诊标志 门诊费用记录.门诊标志%Type;
  v_付款方式 医疗付款方式.名称%Type;

  --临时变量 
  n_Count      Number;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  n_新病人模式 Number;
  v_出库no     药品收发记录.No%Type;
  v_Date       Date;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;
  n_组id 财务缴款分组.Id%Type;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  Select Count(ID)
  Into n_Count
  From 门诊费用记录
  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And 操作员姓名 Is Null;
  If n_Count = 0 Then
    Select Max(操作员姓名) Into v_操作员姓名 From 门诊费用记录 Where 记录性质 = 1 And NO = No_In;
    If v_操作员姓名 Is Not Null Then
      If v_操作员姓名 = 操作员姓名_In Then
        v_Err_Msg := '不能读取划价单内容,该单据已经被收费！';
        Raise Err_Special;
      Else
        v_Err_Msg := '不能读取划价单内容,该单据已经被收费！';
        Raise Err_Item;
      End If;
    Else
      v_Err_Msg := '不能读取划价单内容,该单据已经被删除！';
      Raise Err_Item;
    End If;
  End If;
  v_Date := 登记时间_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    --根据门诊标志获取门诊号/住院号
    Select Max(门诊标志), Max(标识号) Into n_门诊标志, v_标识号 From 门诊费用记录 Where 记录性质 = 1 And NO = No_In;
    If v_标识号 Is Null Then
      Select Decode(n_门诊标志, 2, 住院号, 门诊号) Into v_标识号 From 病人信息 Where 病人id = 病人id_In;
    End If;
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
      Set 记录状态 = 1, 病人id = Decode(病人id_In, 0, Null, 病人id_In), 标识号 = Nvl(标识号, v_标识号), 付款方式 = 付款方式_In, 姓名 = 姓名_In,
          年龄 = 年龄_In, 性别 = 性别_In,
          --可能保持医嘱发送的内容 
          病人科室id = Nvl(病人科室id_In, 病人科室id), 开单部门id = Nvl(开单部门id_In, 开单部门id), 开单人 = Nvl(开单人_In, 开单人), 结帐金额 = 实收金额,
          结帐id = 结帐id_In, 发生时间 = 发生时间_In, 登记时间 = v_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 是否急诊 = 是否急诊_In,
          缴款组id = n_组id, 费用状态 = 1, 执行状态 = Decode(Nvl(执行状态, 0), -1, Null, Nvl(执行状态, 0))
      Where ID = t_费用id(I) And 记录状态 = 0;
  
    If Sql%RowCount <> t_费用id.Count Then
      Select Count(1)
      Into n_Count
      From 门诊费用记录
      Where 记录状态 = 1 And ID In (Select Column_Value From Table(t_费用id));
      If n_Count <> t_费用id.Count Then
        v_Err_Msg := '由于并发操作,该单据已经删除！';
        Raise Err_Item;
      Else
        Select Max(操作员姓名)
        Into v_操作员姓名
        From 门诊费用记录
        Where 记录状态 = 1 And ID In (Select Column_Value From Table(t_费用id));
        If v_操作员姓名 = 操作员姓名_In Then
          v_Err_Msg := '由于并发操作,该单据已经收费！';
          Raise Err_Special;
        Else
          v_Err_Msg := '由于并发操作,该单据已经收费！';
          Raise Err_Item;
        End If;
      End If;
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
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人划价收费_Insert;
/

--118128:刘涛,2017-12-11,不分批批次等于0处理
Create Or Replace Procedure Zl_材料其他入库_Insert
(
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  入出类别id_In In 药品收发记录.入出类别id%Type,
  材料id_In     In 药品收发记录.药品id%Type,
  实际数量_In   In 药品收发记录.实际数量%Type,
  成本价_In     In 药品收发记录.成本价%Type,
  成本金额_In   In 药品收发记录.成本金额%Type,
  零售价_In     In 药品收发记录.零售价%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  差价_In       In 药品收发记录.差价%Type,
  零售差价_In   In 药品收发记录.差价%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  产地_In       In 药品收发记录.产地%Type := Null,
  批号_In       In 药品收发记录.批号%Type := Null,
  生产日期_In   In 药品收发记录.生产日期%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
  商品条码_In   In 药品收发记录.商品条码%Type := Null,
  批准文号_In   In 药品收发记录.批准文号%Type := Null
) Is
  n_Id       药品收发记录.Id%Type; --收发ID
  n_入出系数 药品收发记录.入出系数%Type;
  n_批次     药品收发记录.批次%Type := Null; --批次
  n_库房分批 Integer; --是否分批核算    1:分批;0：不分批
  n_在用分批 Integer; --是否分批核算    1:分批;0：不分批
Begin
  If Not 批准文号_In Is Null And Not 产地_In Is Null Then
    Update 药品生产商对照 Set 批准文号 = 批准文号_In Where 药品id = 材料id_In And 厂家名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 药品生产商对照 (药品id, 厂家名称, 批准文号) Values (材料id_In, 产地_In, 批准文号_In);
    End If;
  End If;

  n_入出系数 := 1;

  Select 药品收发记录_Id.Nextval Into n_Id From Dual;

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
	Else
      n_批次 := 0;
    End If;
  Else
    n_批次 := n_Id;
  End If;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌日期, 灭菌效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价,
     摘要, 填制人, 填制日期, 生产日期, 用法, 商品条码, 批准文号)
  Values
    (n_Id, 1, 17, No_In, 序号_In, 库房id_In, 入出类别id_In, n_入出系数, 材料id_In, n_批次, 产地_In, 批号_In, 效期_In, 灭菌日期_In, 灭菌效期_In,
     实际数量_In, 实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, 生产日期_In, 零售差价_In, 商品条码_In,
     批准文号_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料其他入库_Insert;
/

--118128:刘涛,2017-12-11,不分批批次等于0处理
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
  验收结论_In   In 药品收发记录.验收结论%Type := Null
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
     扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 发药方式, 配药人, 配药日期, 注册证号, 用法, 商品条码, 内部条码, 费用id, 批准文号, 验收结论)
  Values
    (v_Lngid, 1, 15, No_In, 序号_In, 库房id_In, 供药单位id_In, n_入出类别id, n_入出系数, 材料id_In, Decode(退货_In, -1, 批次_In, n_批次), 产地_In,
     批号_In, 生产日期_In, 效期_In, 灭菌日期_In, 灭菌效期_In, 退货_In * 实际数量_In, 退货_In * 实际数量_In, 成本价_In, 退货_In * 成本金额_In, 扣率_In, 零售价_In,
     退货_In * 零售金额_In, 退货_In * 差价_In, 摘要_In, 填制人_In, 填制日期_In, Decode(退货_In, -1, 1, 0), 核查人_In, 核查日期_In, 注册证号_In, 零售差价_In,
     商品条码_In, 内部条码_In, 费用id_In, 批准文号_In, 验收结论_In);

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
  
    If v_可用数量 - 实际数量_In < 0 Then
      v_Err_Msg := '[ZLSOFT]第' || 序号_In || '行的可用数量不够,请检查[ZLSOFT]';
      Raise Err_Item;
    End If;
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

--114951:陈刘,2017-12-08,体温单批量录入增加骑线体温自动标记

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
      End Case;
    End Loop;
  Else
    n_开始时点 := 4;
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
--114951:陈刘,2017-12-08,体温单批量录入增加骑线体温自动标记
CREATE OR REPLACE Procedure Zl_体温单数据_Update
(
  文件id_In   In 病人护理文件.Id%Type, --病人护理文件ID
  发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
  记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
  项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
  记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
  体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
  复试合格_In In Number := 0,
  未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
  他人记录_In In Number := 1,
  数据来源_In In 病人护理明细.数据来源%Type := 0,
  来源id_In   In 病人护理明细.来源id%Type := Null, --始终为原始记录的来源ID
  共用_In     In 病人护理明细.共用%Type := 0,
  项目首次_In In Number := 0, --汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
  开始时间_In In 病人护理数据.发生时间%Type := Null, --本记录有效跨度的开始时间
  结束时间_In In 病人护理数据.发生时间%Type := Null, --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除


  操作员_In   In 病人护理数据.保存人%Type := Null,
  检查科室_In In Number := 1,
  显示_In     In Number := 0,
  骑线_In     In Number := 0
) Is
  n_项目序号 病人护理明细.项目序号%Type;
  n_记录标记 病人护理明细.记录标记%Type; --记录内容的特殊标志
  v_保存人   病人护理数据.保存人%Type;
  v_记录人   病人护理明细.记录人%Type;
  d_结束时间 病人护理数据.发生时间%Type;
  d_发生时间 病人护理数据.发生时间%Type;
  d_开始时间 病人护理数据.发生时间%Type;
  n_记录id   病人护理明细.记录id%Type;
  v_科室id   病人护理文件.科室id%Type;
  n_心率应用 护理记录项目.应用方式%Type;
  n_脉搏     护理记录项目.项目序号%Type := 2;
  n_体温     护理记录项目.项目序号%Type := 1;
  n_心率     护理记录项目.项目序号%Type := -1;
  n_项目性质 护理记录项目.项目性质%Type := 1;
  n_开始版本 病人护理明细.开始版本%Type;
  n_疼痛强度 护理记录项目.项目序号%Type;
  n_Newid    病人护理明细.Id%Type;

  n_病人id       病人护理文件.病人id%Type;
  n_主页id       病人护理文件.主页id%Type;
  n_婴儿         病人护理文件.婴儿%Type;
  d_婴儿出院时间 病人医嘱记录.开始执行时间%Type;
  d_文件开始时间 病人护理文件.开始时间%Type;
  n_Preblue      Number;
  n_i            Number;
  n_Sqlrowcount  Number;
  n_Count        Number(1);
  v_记录内容     病人护理明细.记录内容%Type;
  v_Data         病人护理明细.记录内容%Type;
  --主过程
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  d_发生时间 := 发生时间_In;

  If d_发生时间 Is Null Then
    v_Error := '数据发生时间不能为空！';
    Raise Err_Custom;
  End If;

  If 开始时间_In Is Null Then
    d_开始时间 := d_发生时间;
  Else
    d_开始时间 := 开始时间_In;
  End If;

  If 结束时间_In Is Null Then
    d_结束时间 := d_开始时间;
  Else
    d_结束时间 := 结束时间_In;
  End If;

  --提取记录ID
  n_记录id := 0;
  If 操作员_In Is Null Then
    v_保存人 := Zl_Username;
  Else
    v_保存人 := 操作员_In;
  End If;
  ----------------------------------------------------------------------------------------------------------------------
  Begin
    Select ID Into n_记录id From 病人护理数据 Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
  Exception
    When Others Then
      n_记录id := 0;
  End;

  --检查数据的发生时间是否对应科室
  ---------------------------------------------------------------------------------------------------------------------
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
            (发生时间_In >= a.开始时间 And (发生时间_In < = Nvl(a.终止时间, Sysdate) Or a.终止时间 Is Null)) And Rownum < 2;
    Exception
      When Others Then
        v_科室id := 0;
    End;
    If v_科室id = 0 And 检查科室_In = 1 Then
      v_Error := '数据发生时间 ' || To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI:SS') || ' 不在病人有效变动时间范围内，不能操作！';
      Raise Err_Custom;
    End If;
  Else
    If 发生时间_In < d_文件开始时间 Or 发生时间_In > d_婴儿出院时间 Then
      v_Error := '数据发生时间 ' || To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI:SS') || ' 不在病人有效变动时间范围内，不能操作！';
      Raise Err_Custom;
    End If;
  End If;
  --检查是不是本人的记录
  ---------------------------------------------------------------------------------------------------------------------
  If 他人记录_In = 0 And n_记录id > 0 Then
    v_记录人 := '';
    Begin
      Select 记录人
      Into v_记录人
      From 病人护理明细
      Where 记录id = n_记录id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And 终止版本 Is Null And Rownum < 2
      Order By Nvl(记录标记, 0);
    Exception
      When Others Then
        v_记录人 := '';
    End;
    If v_记录人 Is Not Null And v_记录人 <> v_保存人 Then
      v_Error := '在' || To_Char(d_开始时间, 'yyyy-mm-dd hh24:mi:ss') || '至' || To_Char(d_结束时间, 'yyyy-mm-dd hh24:mi:ss') ||
                 '段内记录人不是当前人，你无权修改！';
      Raise Err_Custom;
    End If;
  End If;

  --提取疼痛强度曲线项目的项目序号
  Begin
    Select 项目序号 Into n_疼痛强度 From 体温记录项目 Where 记录名 = '疼痛强度';
  Exception
    When Others Then
      n_疼痛强度 := -999;
  End;
  --检查脉搏心率是否共用
  If 项目序号_In = n_脉搏 Then
    n_项目序号 := n_心率;
  Else
    n_项目序号 := 项目序号_In;
  End If;
  Begin
    Select 应用方式, 项目性质 Into n_心率应用, n_项目性质 From 护理记录项目 Where 项目序号 = n_项目序号;
  Exception
    When Others Then
      n_心率应用 := 0;
  End;

  ----清除某段时间内的护理数据信息
  --项目首次_In 汇总项目根据汇总时间段保存一天数据时先清除在保存 项目首次_In：=1
  --记录内容_In Is Null And 未记说明_In Is Null 则认为删除数据
  ---------------------------------------------------------------------------------------------------------------------
  If (项目首次_In = 1) Or (记录内容_In Is Null And 未记说明_In Is Null) Then
    For r_List In (Select l.Id, Count(*) As 记录数, Min(l.发生时间) 发生时间
                   From 病人护理文件 A, 病人护理数据 L, 病人护理明细 D
                   Where a.Id = l.文件id And l.Id = d.记录id And a.Id = 文件id_In And d.终止版本 Is Null And l.发生时间 >= d_开始时间 And
                         l.发生时间 <= d_结束时间
                   Group By l.Id) Loop
      n_Sqlrowcount := 0;
      If 记录类型_In = 2 Or 记录类型_In = 6 Then
        Delete 病人护理明细
        Where 记录id = r_List.Id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And 终止版本 Is Null;
        n_Sqlrowcount := Sql%RowCount;
      Else
        If 体温部位_In Is Not Null Then
          --此处主要针对活动项目
          Delete 病人护理明细
          Where 记录id = r_List.Id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And Nvl(体温部位, '无') = Nvl(体温部位_In, '无') And
                终止版本 Is Null;
        Else
          Delete 病人护理明细
          Where 记录id = r_List.Id And 记录类型 = 记录类型_In And 项目序号 = 项目序号_In And 终止版本 Is Null;
        End If;

        n_Sqlrowcount := Sql%RowCount;
        --如果脉搏和心率共用删除脉搏是同时删除心率数据
        If 项目序号_In = n_脉搏 And n_心率应用 = 2 Then
          Delete 病人护理明细
          Where 记录id = r_List.Id And 记录类型 = 记录类型_In And 项目序号 = n_心率 And 终止版本 Is Null;
          n_Sqlrowcount := n_Sqlrowcount + Sql%RowCount;
        End If;
        --如果为收缩压/舒张压删除收缩压时同时删除舒张压数据
        If 项目序号_In = 4 Then
          Delete 病人护理明细
          Where 记录id = r_List.Id And 记录类型 = 记录类型_In And 项目序号 = 5 And 终止版本 Is Null;
          n_Sqlrowcount := n_Sqlrowcount + Sql%RowCount;
        End If;
      End If;
      If n_Sqlrowcount >= r_List.记录数 Then
        Delete 病人护理数据 Where ID = r_List.Id;
      End If;
      --更新打印
      Update 体温单打印
      Set 打印人 = Null, 打印时间 = Null
      Where 文件id = 文件id_In And
            开始时间 = (Select Max(开始时间) From 体温单打印 Where 文件id = 文件id_In And 开始时间 <= r_List.发生时间);
    End Loop;
  End If;

  If 记录内容_In Is Null And 未记说明_In Is Null Then
    Return;
  End If;

  --分解项目记录内容
  n_Preblue := 0;
  If (记录类型_In = 1 Or 记录类型_In = 7) And Instr(',' || n_疼痛强度 || ',1,2,4,', ',' || 项目序号_In || ',', 1) > 0 Then
    n_Preblue := Nvl(Instr(Nvl(记录内容_In, ''), '/', 1), 0);
    If n_Preblue > 1 Then
      n_Preblue := 1;
    End If;
  End If;

  If 项目序号_In = 4 And n_Preblue = 0 Then
    v_Error := '血压数据格式错误! 格式:收缩压/舒张压。';
    Raise Err_Custom;
  End If;

  --确认开始版本号
  ---------------------------------------------------------------------------------------------------------------------
  n_开始版本 := 1;

  --改写病人护理数据：如果已经存在与病人、科室和发生时间相同的记录则修改，否则增加新的记录
  ---------------------------------------------------------------------------------------------------------------------
  --汇总项目是删除后在增加，可能开始提取的记录ID已经不存在。
  Begin
    Select ID Into n_记录id From 病人护理数据 Where ID = n_记录id;
  Exception
    When Others Then
      n_记录id := 0;
  End;

  If n_记录id = 0 Then
    Select 病人护理数据_Id.Nextval Into n_记录id From Dual;
    Insert Into 病人护理数据
      (ID, 文件id, 显示, 发生时间, 保存人, 保存时间, 最后版本)
    Values
      (n_记录id, 文件id_In, 0, d_发生时间, v_保存人, Sysdate, n_开始版本);
  End If;

  --检查删除物理降温数据或脉搏短轴数据
  If (项目序号_In = n_体温 Or 项目序号_In = n_疼痛强度 Or (项目序号_In = n_脉搏 And n_心率应用 = 2)) And n_Preblue = 0 Then
    Delete From 病人护理明细
    Where 记录id = n_记录id And 项目序号 = Decode(项目序号_In, n_脉搏, n_心率, 项目序号_In) And Decode(项目序号_In, n_脉搏, 1, Nvl(记录标记, 0)) = 1 And
          记录类型 = 记录类型_In And 终止版本 Is Null;
  End If;

  --改写病人护理明细：如果已经存在与病人、科室和发生时间相同的记录则修改，否则增加新的记录
  -----------------------------------------------------------------------------------------------------------------------
  v_Data     := 记录内容_In;
  n_项目序号 := 项目序号_In;
  For n_i In 0 .. n_Preblue Loop
    If n_i = 0 Then
      If 项目序号_In = n_心率 Then
        n_记录标记 := 1;
      Else
        n_记录标记 := 0;
      End If;
    Else
      --收缩压/舒张压
      If 项目序号_In = 4 Then
        n_记录标记 := 0;
        n_项目序号 := 5;
      Else
        n_记录标记 := 1;
        If 项目序号_In = n_脉搏 Then
          n_项目序号 := n_心率;
        End If;
      End If;

    End If;
    If n_Preblue > 0 Then
      v_记录内容 := Substr(v_Data, 1, Instr(v_Data, '/', 1) - 1);
      If v_记录内容 Is Null Then
        v_记录内容 := v_Data;
      End If;
    Else
      v_记录内容 := v_Data;
    End If;

    --检查是否需要标记骑线
    Select Count(b.记录名)
    Into n_Count
    From 护理记录项目 A, 体温记录项目 B
    Where a.项目序号 = b.项目序号 And a.项目序号 = 项目序号_In And b.记录法 = 1 And a.分组名 = '1)体温曲线项目';

    --为了兼容以前同步过来的心率数据记录标记为0
    if 项目序号_In=0  then
    If n_i = 0 Then
      Update 病人护理明细
      Set 记录内容 = v_记录内容, 体温部位 = 体温部位_In, 复试合格 = 复试合格_In,
          未记说明 = Decode(n_项目序号, n_体温, Decode(v_记录内容, '不升', Null, 未记说明_In), 未记说明_In), 记录人 = v_保存人, 记录时间 = Sysdate
      Where 记录id = n_记录id And 项目序号 = n_项目序号 And 记录类型 = 记录类型_In And
            Decode(项目序号_In, n_体温, Nvl(记录标记, 0), n_疼痛强度, Nvl(记录标记, 0), Nvl(n_记录标记, 0)) = Nvl(n_记录标记, 0) And
            Decode(n_项目性质, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And 终止版本 Is Null;
    Else
      Update 病人护理明细
      Set 记录内容 = v_记录内容, 记录人 = v_保存人, 记录时间 = Sysdate
      Where 记录id = n_记录id And 项目序号 = n_项目序号 And 记录类型 = 记录类型_In And
            Decode(项目序号_In, n_体温, Nvl(记录标记, 0), n_疼痛强度, Nvl(记录标记, 0), Nvl(n_记录标记, 0)) = Nvl(n_记录标记, 0) And 终止版本 Is Null;
    End If;
    else
      If n_i = 0 Then
      Update 病人护理明细
      Set 记录内容 = v_记录内容, 体温部位 = 体温部位_In, 复试合格 = 复试合格_In, 记录类型 = 记录类型_In,
          未记说明 = Decode(n_项目序号, n_体温, Decode(v_记录内容, '不升', Null, 未记说明_In), 未记说明_In), 记录人 = v_保存人, 记录时间 = Sysdate
      Where 记录id = n_记录id And 项目序号 = n_项目序号  And
            Decode(项目序号_In, n_体温, Nvl(记录标记, 0), n_疼痛强度, Nvl(记录标记, 0), Nvl(n_记录标记, 0)) = Nvl(n_记录标记, 0) And
            Decode(n_项目性质, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And 终止版本 Is Null;
      Else
        Update 病人护理明细
        Set 记录内容 = v_记录内容, 记录人 = v_保存人,记录类型 = 记录类型_In, 记录时间 = Sysdate
        Where 记录id = n_记录id And 项目序号 = n_项目序号  And
              Decode(项目序号_In, n_体温, Nvl(记录标记, 0), n_疼痛强度, Nvl(记录标记, 0), Nvl(n_记录标记, 0)) = Nvl(n_记录标记, 0) And 终止版本 Is Null;
      End If;
    end if;
    If Sql%RowCount = 0 Then
      --插入本次登记的病人护理内容
      If Mod(记录类型_In, 10) = 1 Or 记录类型_In = 7 Then
        Select 病人护理明细_Id.Nextval Into n_Newid From Dual;
        Insert Into 病人护理明细
          (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 记录人, 体温部位, 复试合格, 开始版本, 终止版本, 记录组号, 未记说明,
           记录时间, 数据来源, 显示, 来源id, 共用)
          Select n_Newid, n_记录id, 记录类型_In, 分组名, 项目id, 项目序号, 项目名称, 项目类型, v_记录内容, 项目单位, n_记录标记, v_保存人, 体温部位_In, 复试合格_In,
                 n_开始版本, Null, Null, Decode(n_项目序号, n_体温, Decode(v_记录内容, '不升', Null, 未记说明_In), 未记说明_In), Sysdate,
                 数据来源_In, 0, 来源id_In, 共用_In
          From 护理记录项目
          Where 项目序号 = n_项目序号;
        If 显示_In = 1 Then
          Zl_体温单数据_设置显示(n_Newid, 1);
        End If;
        If n_Count > 0 And 骑线_In = 1 Then
          Zl_体温单骑线设置_Update(文件id_In, 发生时间_In, 项目序号_In, n_记录id);
        End If;
      Else
        Insert Into 病人护理明细
          (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 记录人, 体温部位, 复试合格, 开始版本, 终止版本, 记录组号, 未记说明,
           记录时间, 数据来源, 显示, 来源id, 共用)
        Values
          (病人护理明细_Id.Nextval, n_记录id, 记录类型_In, Null, Null, 0,
           Decode(记录类型_In, 2, '上标说明', 6, '下标说明', 3, '入出转', 4, v_记录内容), Decode(记录类型_In, 3, 0, 1),
           Decode(记录类型_In, 4, '1', 记录内容_In), '', n_记录标记, v_保存人, 体温部位_In, 复试合格_In, n_开始版本, Null, Null, 未记说明_In, Sysdate,
           数据来源_In, 0, 来源id_In, 共用_In);
      End If;
    Else

      If n_Count > 0 And 骑线_In = 1 Then
        Zl_体温单骑线设置_Update(文件id_In, 发生时间_In, 项目序号_In, n_记录id, 1);
      End If;
    End If;
    If n_Preblue > 0 Then
      v_Data := Substr(v_Data, Instr(v_Data, '/', 1) + 1);
    End If;
  End Loop;
  --更新打印
  Update 体温单打印
  Set 打印人 = Null, 打印时间 = Null
  Where 文件id = 文件id_In And
        开始时间 = (Select Max(开始时间) From 体温单打印 Where 文件id = 文件id_In And 开始时间 <= d_发生时间);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_体温单数据_Update;
/

--117925:刘涛,2017-12-08,排除导致死锁处理
Create Or Replace Procedure Zl_材料领用_Delete(
                                           --删除药品收发记录及恢复相应的表：药品库存
                                           No_In In 药品收发记录.No%Type) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(100);
  v_下库存   Zlparameters.参数值%Type;
  n_实价卫材 收费项目目录.是否变价%Type;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_下库存 From Dual;

  If v_下库存 = 1 Then
    --通过循环，恢复原来的可用数量
    For v_单据 In (Select ID, 填写数量, 库房id, 零售价, 批次, 批号, 药品id, 供药单位id, 成本价, 效期, 灭菌效期, 产地, 生产日期, 批准文号
                 From 药品收发记录
                 Where NO = No_In And 单据 = 20 And 入出系数 = -1
                 Order By 药品id, 批次) Loop
      Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = v_单据.药品id;
    
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + v_单据.填写数量
      Where 库房id = v_单据.库房id And 药品id = v_单据.药品id And Nvl(批次, 0) = Nvl(v_单据.批次, 0) And 性质 = 1;
    
      If Sql%NotFound Then
      
        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
        Values
          (v_单据.库房id, v_单据.药品id, v_单据.批次, 1, v_单据.填写数量, v_单据.效期, v_单据.灭菌效期, v_单据.供药单位id, v_单据.成本价, v_单据.批号, v_单据.生产日期,
           v_单据.产地, v_单据.批准文号, Decode(n_实价卫材, 1, Decode(Nvl(v_单据.批次, 0), 0, Null, v_单据.零售价), Null));
      End If;
    
      Delete From 药品库存
      Where 库房id = v_单据.库房id And 药品id = v_单据.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
            Nvl(实际差价, 0) = 0;
      Delete From 材料领用信息 Where 收发id = v_单据.Id;
    End Loop;
  End If;

  Delete From 药品收发记录 Where NO = No_In And 单据 = 20 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料领用_Delete;
/

--117925:刘涛,2017-12-08,排序导致死锁处理
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
  v_条码前缀 := Nvl(Zl_Getsysparameter(159), '');

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
  
    --重新计算库存表中的平均成本价
    Update 药品库存
    Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
    Where 药品id = c_收发.药品id And Nvl(批次, 0) = Nvl(c_收发.批次, 0) And 库房id = c_收发.库房id And Nvl(实际数量, 0) <> 0 And 性质 = 1;
    If Sql%NotFound Then
      Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = c_收发.药品id;
      Update 药品库存
      Set 平均成本价 = n_平均成本价
      Where 药品id = c_收发.药品id And 库房id = c_收发.库房id And Nvl(批次, 0) = Nvl(c_收发.批次, 0) And 性质 = 1;
    End If;
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料其他入库_Verify;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_材料申领_Delete(No_In In 药品收发记录.NO%Type) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(500);
  v_下库存   zlParameters.参数值%Type;
  v_明确批次 zlParameters.参数值%Type;
  d_发送日期 药品收发记录.配药日期%Type;
  n_实价卫材 收费项目目录.是否变价%Type;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_下库存 From Dual;

  --只有在明确批次的情况下才能下可用库存
  Select Nvl(zl_GetSysParameter(83), '0') Into v_明确批次 From Dual;

  --检查是否已经发料或审核
  Select 配药日期 Into d_发送日期 From 药品收发记录 Where 单据 = 19 And NO = No_In And Rownum < 2;

  If d_发送日期 Is Not Null Then
    v_Err_Msg := '[ZLSOFT]该申领单已经被他人发料,不能进行删除![ZLSOFT]';
    Raise Err_Item;
  End If;

  If To_Number(v_下库存, '9999') = 1 And To_Number(v_明确批次, '9999') = 1 Then
    --需要还原可用库存
    --通过循环，恢复原来的可用数量
    For v_单据 In (Select 实际数量, 库房id, 批次, 药品id, 零售价, 批号, 效期, 产地, 供药单位id, 成本价, 生产日期,
                          批准文号, 灭菌效期
                   From 药品收发记录
                   Where NO = No_In And 单据 = 19 And 入出系数 = -1
                   Order By 药品id, 批次) Loop
      Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = v_单据.药品id;

      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + v_单据.实际数量
      Where 库房id = v_单据.库房id And 药品id = v_单据.药品id And Nvl(批次, 0) = Nvl(v_单据.批次, 0) And 性质 = 1;

      If Sql%NotFound Then

        Insert Into 药品库存
          (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期,
           上次产地, 批准文号, 零售价)
        Values
          (v_单据.库房id, v_单据.药品id, v_单据.批次, 1, v_单据.实际数量, v_单据.效期, v_单据.灭菌效期,
           v_单据.供药单位id, v_单据.成本价, v_单据.批号, v_单据.生产日期, v_单据.产地, v_单据.批准文号,
           Decode(n_实价卫材, 1, Decode(Nvl(v_单据.批次, 0), 0, Null, v_单据.零售价), Null));
      End If;

      Delete From 药品库存
      Where 库房id = v_单据.库房id And 药品id = v_单据.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
            Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
    End Loop;
  End If;

  Delete 药品收发记录 Where NO = No_In And 单据 = 19 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception

  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料申领_Delete;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_材料盘点_Strike
(
  No_In     In 药品收发记录.No%Type,
  审核人_In In 药品收发记录.审核人%Type
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  n_Batch_Count Integer; --原不分批现在分批的材料的数量
  n_Count       Integer; --原分批现不分批

  n_批次       药品收发记录.批次%Type;
  n_成本价     药品收发记录.成本价%Type;
  n_材料id     药品收发记录.药品id%Type;
  n_实价卫材   收费项目目录.是否变价%Type;
  n_平均成本价 药品库存.平均成本价%Type;

Begin
  Update 药品收发记录 Set 记录状态 = 3 Where NO = No_In And 单据 = 22 And 记录状态 = 1;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
    Raise Err_Item;
  End If;

  --主要针对原不分批现在分批的材料，不能对其审核
  Select Count(*), Max(a.药品id)
  Into n_Batch_Count, n_材料id
  From 药品收发记录 A, 材料特性 B
  Where a.药品id = b.材料id And a.No = No_In And a.单据 = 22 And a.记录状态 = 3 And Nvl(a.批次, 0) = 0 And
        ((Nvl(b.库房分批, 0) = 1 And
        a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '发料部门') Or (工作性质 Like '制剂室'))) Or Nvl(b.在用分批, 0) = 1);

  If n_Batch_Count > 0 Then
    Begin
      Select 编码 || '-' || 名称 Into v_Err_Msg From 收费项目目录 Where ID = n_材料id;
    Exception
      When Others Then
        Null;
    End;
    v_Err_Msg := '[ZLSOFT]该单据中材料为:' || v_Err_Msg || Chr(10) || Chr(13) || '的材料,原来不分批,而现在分批，因此不能审核！[ZLSOFT]';
    Raise Err_Item;
  End If;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 填写数量, 扣率, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 审核人, 审核日期, 频次, 供药单位id, 生产日期, 批准文号, 单量)
    Select 药品收发记录_Id.Nextval, 2, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, a.药品id,
           Decode(Nvl(a.批次, 0), 0, Null, (Decode(Nvl(b.库房分批, 0), 0, Null, a.批次))), a.产地, 批号, a.效期, a.灭菌效期, 填写数量, a.扣率,
           -实际数量, a.成本价, 成本金额, 零售价, -零售金额, -差价, 摘要, 审核人_In, Sysdate, 审核人_In, Sysdate, a.频次, a.供药单位id, a.生产日期, a.批准文号,
           a.单量
    From 药品收发记录 A, 材料特性 B
    Where NO = No_In And a.药品id = b.材料id And 单据 = 22 And 记录状态 = 3;

  For c_单据 In (Select ID, 实际数量, 零售价, 零售金额, 差价, 库房id, 药品id 材料id, 批次, 批号, 效期, 灭菌效期, 产地, 入出类别id, 入出系数, 供药单位id, 生产日期, 批准文号,
                      单量
               From 药品收发记录
               Where NO = No_In And 单据 = 22 And 记录状态 = 2
               Order By 药品id, 批次) Loop
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = c_单据.材料id;
  
    --原分批现不分批的材料,在C冲消时，要处理他
    Begin
      Select Count(*)
      Into n_Count
      From 药品收发记录 A, 材料特性 B
      Where b.材料id + 0 = c_单据.材料id And a.No = No_In And a.药品id = b.材料id And a.单据 = 22 And a.库房id + 0 = c_单据.库房id And
            a.记录状态 = 3 And Nvl(a.批次, 0) > 0 And
            (Nvl(b.库房分批, 0) = 0 Or
            (Nvl(b.在用分批, 0) = 0 And
            a.库房id In (Select 部门id From 部门性质说明 Where (工作性质 Like '发料部门') Or (工作性质 Like '制剂室'))));
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      n_批次 := 0;
    Else
      n_批次 := Nvl(c_单据.批次, 0);
    End If;
  
    --更改药品库存表的相应数据
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + Nvl(c_单据.实际数量, 0) * c_单据.入出系数, 实际数量 = Nvl(实际数量, 0) + Nvl(c_单据.实际数量, 0) * c_单据.入出系数,
        实际金额 = Nvl(实际金额, 0) + Nvl(c_单据.零售金额, 0) * c_单据.入出系数, 实际差价 = Nvl(实际差价, 0) + Nvl(c_单据.差价, 0) * c_单据.入出系数,
        零售价 = Decode(n_实价卫材, 1, Decode(Nvl(n_批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_单据.零售价, 零售价)), Null)
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(批次, 0) = n_批次 And 性质 = 1;
  
    If Sql%NotFound Then
      If Nvl(c_单据.实际数量, 0) <> 0 Then
        n_成本价 := Round((Nvl(c_单据.零售金额, 0) - Nvl(c_单据.差价, 0)) / c_单据.实际数量, 7);
      Else
        n_成本价 := 0;
      End If;
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价, 平均成本价)
      
      Values
        (c_单据.库房id, c_单据.材料id, n_批次, 1, c_单据.实际数量 * c_单据.入出系数, c_单据.实际数量 * c_单据.入出系数, c_单据.零售金额 * c_单据.入出系数,
         c_单据.差价 * c_单据.入出系数, c_单据.效期, c_单据.灭菌效期, c_单据.供药单位id, n_成本价, c_单据.批号, c_单据.生产日期, c_单据.产地, c_单据.批准文号,
         Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, c_单据.零售价), Null), n_成本价);
    End If;
  
    Delete From 药品库存
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
    Zl_材料收发记录_调价修正(c_单据.Id);
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料盘点_Strike;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
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
  
    --更改药品库存表的相应数据
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + Decode(c_单据.入出系数, 1, Nvl(c_单据.实际数量, 0), 0),
        实际数量 = Nvl(实际数量, 0) + Nvl(c_单据.实际数量, 0) * c_单据.入出系数, 实际金额 = Nvl(实际金额, 0) + Nvl(c_单据.零售金额, 0) * c_单据.入出系数,
        实际差价 = Nvl(实际差价, 0) + Nvl(c_单据.差价, 0) * c_单据.入出系数,
        零售价 = Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_单据.零售价, 零售价)), Null)
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      If Nvl(c_单据.实际数量, 0) <> 0 Then
        n_成本价 := Round((Nvl(c_单据.零售金额, 0) - Nvl(c_单据.差价, 0)) / c_单据.实际数量, 7);
      Else
        n_成本价 := 0;
      End If;
    
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价,平均成本价)
      
      Values
        (c_单据.库房id, c_单据.材料id, c_单据.批次, 1, Decode(c_单据.入出系数, 1, Nvl(c_单据.实际数量, 0), 0), c_单据.实际数量 * c_单据.入出系数,
         c_单据.零售金额 * c_单据.入出系数, c_单据.差价 * c_单据.入出系数, c_单据.效期, c_单据.灭菌效期, c_单据.供药单位id, n_成本价, c_单据.批号, c_单据.生产日期, c_单据.产地,
         c_单据.批准文号, Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, c_单据.零售价), Null),n_成本价);
    End If;
  
    Delete From 药品库存
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  
    --重新计算平均成本价
    Update 药品库存
    Set 平均成本价 = Decode(Nvl(批次, 0), 0, decode((实际金额 - 实际差价) / 实际数量,0,上次采购价,(实际金额 - 实际差价) / 实际数量), 上次采购价)
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And 性质 = 1 And Nvl(实际数量, 0) <> 0;
    If Sql%NotFound Then
      Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = c_单据.材料id;
      Update 药品库存
      Set 平均成本价 = n_平均成本价
      Where 药品id = c_单据.材料id And 库房id = c_单据.库房id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) and 性质=1;
    End If;
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料盘点_Verify;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_材料盘点_Delete(

                                               --删除药品收发记录及恢复相应的表：药品库存
                                               No_In In 药品收发记录.NO%Type) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(100);
  n_成本价   药品收发记录.成本价%Type;
  n_实价卫材 收费项目目录.是否变价%Type;

Begin
  --通过循环，恢复出库类别原来的可用数量，
  --实际数量保存的是数量差
  For c_单据 In (Select 实际数量, 库房id, 批次, 药品id 材料id, 零售价, 供药单位id, 零售金额, 差价, 效期, 灭菌效期, 产地,
                        批号, 生产日期, 批准文号
                 From 药品收发记录
                 Where NO = No_In And 单据 = 22 And 入出系数 = -1
                 Order By 药品id, 批次) Loop
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = c_单据.材料id;

    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + c_单据.实际数量,
        零售价 = Decode(n_实价卫材, 1,
                         Decode(Nvl(c_单据.批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_单据.零售价, 零售价)), Null)
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And 性质 = 1;

    If Sql%NotFound Then
      If Nvl(c_单据.实际数量, 0) <> 0 Then
        n_成本价 := Round((Nvl(c_单据.零售金额, 0) - Nvl(c_单据.差价, 0)) / c_单据.实际数量, 7);
      Else
        n_成本价 := 0;
      End If;
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 上次供应商id, 上次采购价, 零售价)
      Values
        (c_单据.库房id, c_单据.材料id, c_单据.批次, 1, c_单据.实际数量, c_单据.供药单位id, n_成本价,
         Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, c_单据.零售价), Null));
    End If;

    Delete From 药品库存
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
          Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
  End Loop;

  Delete From 药品收发记录 Where NO = No_In And 单据 = 22 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料盘点_Delete;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_材料移库_Back(No_In In 药品收发记录.NO%Type) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;

  d_发送日期 药品收发记录.配药日期%Type;
  v_备料     药品收发记录.配药人%Type;
  v_审核     药品收发记录.审核人%Type;
  v_下库存   zlParameters.参数值%Type;
  n_实价卫材 收费项目目录.是否变价%Type;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_下库存 From Dual;

  Select 配药人, 配药日期, 审核人
  Into v_备料, d_发送日期, v_审核
  From 药品收发记录
  Where 单据 = 19 And NO = No_In And Rownum < 2;

  If v_审核 Is Not Null Then
    v_Err_Msg := '[ZLSOFT]该单据已被库房接收，不再允许回退！[ZLSOFT]';
    Raise Err_Item;
  End If;

  If v_备料 Is Null Then
    Return;
  End If;

  If d_发送日期 Is Null Then
    --仅更新配药人为空即可
    Update 药品收发记录 Set 配药人 = Null, 外观 = Null Where 单据 = 19 And NO = No_In;
  Else

    --需要恢复出库库房的可用数量
    Update 药品收发记录 Set 配药日期 = Null Where 单据 = 19 And NO = No_In;
    --如果是在增加单据时已经下了库存的,则本次回退不再恢愎可用可存了.
    If To_Number(v_下库存, '9999') <> 1 Then

      For v_单据 In (Select 实际数量, 库房id, 零售价, 批次, 药品id, 批号, 效期, 产地, 供药单位id, 成本价, 生产日期,
                            灭菌效期, 批准文号
                     From 药品收发记录
                     Where NO = No_In And 单据 = 19 And 入出系数 = -1
                     Order By 药品id, 批次) Loop
        Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = v_单据.药品id;

        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + v_单据.实际数量
        Where 库房id = v_单据.库房id And 药品id = v_单据.药品id And Nvl(批次, 0) = Nvl(v_单据.批次, 0) And 性质 = 1;

        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期,
             上次产地, 批准文号, 零售价)
          Values
            (v_单据.库房id, v_单据.药品id, Nvl(v_单据.批次, 0), 1, v_单据.实际数量, v_单据.效期, v_单据.灭菌效期,
             v_单据.供药单位id, v_单据.成本价, v_单据.批号, v_单据.生产日期, v_单据.产地, v_单据.批准文号,
             Decode(n_实价卫材, 1, Decode(Nvl(v_单据.批次, 0), 0, Null, v_单据.零售价), Null));

        End If;

        Delete From 药品库存
        Where 库房id = v_单据.库房id And 药品id = v_单据.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
              Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
      End Loop;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料移库_Back;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_材料移库_Prepare
(
  No_In     In 药品收发记录.NO%Type,
  操作员_In Varchar2 := Null
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  v_下库存   zlParameters.参数值%Type;
  n_实价卫材 收费项目目录.是否变价%Type;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_下库存 From Dual;

  If 操作员_In Is Not Null Then
    Update 药品收发记录
    Set 配药人 = 操作员_In, 外观 = To_Char(Sysdate, 'yyyy-MM-dd hh24:mi:ss')
    Where 单据 = 19 And NO = No_In;

  Else

    Update 药品收发记录 Set 配药日期 = Sysdate Where 单据 = 19 And NO = No_In;

    If To_Number(v_下库存, '9999') <> 1 Then
      For v_单据 In (Select 实际数量, 库房id, 零售价, 批次, 药品id, 批号, 效期, 产地, 供药单位id, 成本价, 灭菌效期,
                            生产日期, 批准文号
                     From 药品收发记录
                     Where NO = No_In And 单据 = 19 And 入出系数 = -1
                     Order By 药品id, 批次) Loop

        Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = v_单据.药品id;

        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - v_单据.实际数量
        Where 库房id = v_单据.库房id And 药品id = v_单据.药品id And Nvl(批次, 0) = Nvl(v_单据.批次, 0) And 性质 = 1;

        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期,
             上次产地, 批准文号, 零售价)
          Values
            (v_单据.库房id, v_单据.药品id, Nvl(v_单据.批次, 0), 1, -1 * v_单据.实际数量, v_单据.效期, v_单据.灭菌效期,
             v_单据.供药单位id, v_单据.成本价, v_单据.批号, v_单据.生产日期, v_单据.产地, v_单据.批准文号,
             Decode(n_实价卫材, 1, Decode(Nvl(v_单据.批次, 0), 0, Null, v_单据.零售价), Null));

        End If;

        Delete From 药品库存
        Where 库房id = v_单据.库房id And 药品id = v_单据.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
              Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;

      End Loop;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料移库_Prepare;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
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
  n_库存数     药品库存.实际数量%Type;
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
    Order By a.药品id,a.批次,a.序号;

  Cursor c_冲销申请记录 Is
    Select a.Id, a.序号, a.库房id, a.对方部门id, a.入出类别id, a.入出系数, a.药品id, Nvl(a.批次, 0) As 批次, a.产地, a.批号, a.效期, a.配药人,
           a.配药日期 As 发送日期, a.摘要, a.供药单位id, a.批准文号, a.生产日期, a.成本价, a.实际数量, a.零售金额, a.差价, a.零售价, Nvl(b.是否变价, 0) As 时价,
           a.扣率, a.单量, a.频次, a.商品条码, a.内部条码, a.灭菌日期, a.灭菌效期
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.No = No_In And a.单据 = 19 And (a.序号 >= 序号_In And a.序号 <= 序号_In + 1) And
          (a.记录状态 = 原记录状态_In And Mod(a.记录状态, 3) = 2) And a.审核日期 Is Null
    Order By a.药品id,a.批次;
Begin
  --获取金额小数位数
  Select Nvl(精度, 2) Into n_小数 From 药品卫材精度 Where 性质 = 0 And 类别 = 2 And 内容 = 4 And 单位 = 5;
  Select Nvl(Zl_Getsysparameter(95), '0') Into v_下可用库存 From Dual;

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
  
    --取库存数
    Begin
      Select Nvl(实际数量, 0)
      Into n_库存数
      From 药品库存
      Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = n_批次 And 性质 = 1;
    Exception
      When Others Then
        n_库存数 := 0;
    End;
  
    If Nvl(n_剩余数量, 0) = 0 Then
      v_Err_Msg := '[ZLSOFT]该单据中第' || Ceil(序号_In / 2) || '行的材料已经被冲销完成,不能再冲！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --如果库存数大于剩余数量,取剩余数量;否则取库存数
    If n_库存数 < n_剩余数量 Then
      n_剩余成本金额 := n_库存数 / n_剩余数量 * n_剩余成本金额;
      n_剩余零售金额 := n_库存数 / n_剩余数量 * n_剩余零售金额;
      n_剩余数量     := n_库存数;
    End If;
  
    --冲销数量大于剩余数量，不允许
    If n_剩余数量 < 冲销数量_In Then
      v_Err_Msg := '[ZLSOFT]该单据中第' || Ceil(序号_In / 2) || '行的卫生材料冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]';
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
    End Loop;
  
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
  
    --取库存数
    Begin
      Select Nvl(实际数量, 0)
      Into n_库存数
      From 药品库存
      Where 库房id = n_库房id And 药品id = 材料id_In And Nvl(批次, 0) = n_批次 And 性质 = 1;
    Exception
      When Others Then
        n_库存数 := 0;
    End;
  
    If Nvl(n_剩余数量, 0) = 0 Then
      v_Err_Msg := '[ZLSOFT]该单据中第' || Ceil(序号_In / 2) || '行的材料已经被冲销完成,不能再冲！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --如果库存数大于剩余数量,取剩余数量;否则取库存数
    If n_库存数 < n_剩余数量 Then
      n_剩余成本金额 := n_库存数 / n_剩余数量 * n_剩余成本金额;
      n_剩余零售金额 := n_库存数 / n_剩余数量 * n_剩余零售金额;
      n_剩余数量     := n_库存数;
    End If;
  
    --冲销数量大于剩余数量，不允许
    If n_剩余数量 < 冲销数量_In Then
      v_Err_Msg := '[ZLSOFT]该单据中第' || Ceil(序号_In / 2) || '行的卫生材料冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]';
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

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_材料移库_Delete
(
  No_In       In 药品收发记录.No%Type,
  记录状态_In In 药品收发记录.记录状态%Type := 1
) Is
  v_发送 药品收发记录.配药日期%Type;

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  v_下库存   Zlparameters.参数值%Type;
  n_实价卫材 收费项目目录.是否变价%Type;

  Cursor c_药品收发记录 Is
    Select 实际数量, 零售价, 库房id, Nvl(批次, 0) As 批次, 药品id, 批号, 效期, 产地, 供药单位id, 成本价, 灭菌效期, 生产日期, 批准文号 
    From 药品收发记录
    Where NO = No_In And 单据 = 19 And 入出系数 = -1
    Order By 药品id,批次;

  Cursor c_申请冲销记录 Is
    Select (-1 * 实际数量) 实际数量, 库房id, Nvl(批次, 0) As 批次, 药品id, 批号, 效期, 产地, 供药单位id, 批准文号, 成本价, 生产日期 
    From 药品收发记录
    Where NO = No_In And 单据 = 19 And 入出系数 = 1 And 记录状态 = 记录状态_In
    Order By 药品id,批次;
Begin
  Select Nvl(Zl_Getsysparameter(95), '0') Into v_下库存 From Dual;

  If 记录状态_In = 1 Then
    --检查是否已发送，已发送的单据需要还原可用数量 
    Select 配药日期 Into v_发送 From 药品收发记录 Where 单据 = 19 And NO = No_In And Rownum < 2;
  
    If v_发送 Is Not Null Or To_Number(v_下库存, '9999') = 1 Then
      --通过循环，恢复原来的可用数量 
      For c_单据 In c_药品收发记录 Loop
        Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = c_单据.药品id;
      
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + c_单据.实际数量
        Where 库房id = c_单据.库房id And 药品id = c_单据.药品id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And 性质 = 1;
      
        If Sql%NotFound Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
          Values
            (c_单据.库房id, c_单据.药品id, c_单据.批次, 1, c_单据.实际数量, c_单据.效期, c_单据.灭菌效期, c_单据.供药单位id, c_单据.成本价, c_单据.批号, c_单据.生产日期,
             c_单据.产地, c_单据.批准文号, Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, c_单据.零售价), Null));
        End If;
      
        Delete From 药品库存
        Where 库房id = c_单据.库房id And 药品id = c_单据.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
              Nvl(实际差价, 0) = 0;
      End Loop;
    End If;
  Else
    --处理移库申请冲销单据 
  
    --如果参数值为1也要恢复原来的可用数量 
    If v_下库存 = '1' Then
      --通过循环，恢复原来的可用数量 
      For v_申请冲销记录 In c_申请冲销记录 Loop
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + v_申请冲销记录.实际数量
        Where 库房id = v_申请冲销记录.库房id And 药品id = v_申请冲销记录.药品id And Nvl(批次, 0) = v_申请冲销记录.批次 And 性质 = 1;
      
        If Sql%NotFound Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 上次批号, 效期, 上次产地, 上次供应商id, 批准文号, 上次采购价, 上次生产日期)
          Values
            (v_申请冲销记录.库房id, v_申请冲销记录.药品id, v_申请冲销记录.批次, 1, v_申请冲销记录.实际数量, v_申请冲销记录.批号, v_申请冲销记录.效期, v_申请冲销记录.产地,
             v_申请冲销记录.供药单位id, v_申请冲销记录.批准文号, v_申请冲销记录.成本价, v_申请冲销记录.生产日期);
        End If;
      
        Delete From 药品库存
        Where 库房id = v_申请冲销记录.库房id And 药品id = v_申请冲销记录.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
              Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
      End Loop;
    End If;
  End If;

  Delete --把入和出两种类别的移库单都删除 
  From 药品收发记录
  Where NO = No_In And 单据 = 19 And 记录状态 = 记录状态_In And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料移库_Delete;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
CREATE OR REPLACE Procedure Zl_材料库存差价调整_Verify
(
  No_In     In 药品收发记录.NO%Type := Null,
  审核人_In In 药品收发记录.审核人%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  Cursor c_材料单据信息 Is
    Select 库房id, 药品id 材料id, 批次, 入出类别id, 差价, 产地, 效期, 灭菌效期, 成本价,单量 as 新成本价, 供药单位id, 生产日期, 批准文号,
           批号
    From 药品收发记录
    Where NO = No_In And 单据 = 18 And 记录状态 = 1
    Order By 药品id, 批次;
Begin
  Update 药品收发记录
  Set 审核人 = 审核人_In, 审核日期 = Sysdate
  Where NO = No_In And 单据 = 18 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;

  For c_单据 In c_材料单据信息 Loop
    --更改药品库存表的相应数据

    Update 药品库存
    Set 实际差价 = Nvl(实际差价, 0) + Nvl(c_单据.差价, 0),上次采购价=c_单据.新成本价,平均成本价=c_单据.新成本价
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And 性质 = 1;

    If Sql%NotFound Then

      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期,
         上次产地, 批准文号,平均成本价)
      Values
        (c_单据.库房id, c_单据.材料id, c_单据.批次, 1, c_单据.差价, c_单据.效期, c_单据.灭菌效期, c_单据.供药单位id,
         c_单据.新成本价, c_单据.批号, c_单据.生产日期, c_单据.产地, c_单据.批准文号,c_单据.新成本价);
    End If;

    Delete From 药品库存
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
          Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料库存差价调整_Verify;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_材料其他出库_Delete(
                                                   --删除药品收发记录及恢复相应的表：药品库存
                                                   No_In In 药品收发记录.NO%Type) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_实价卫材 收费项目目录.是否变价%Type;

Begin
  --通过循环，恢复原来的可用数量
  For c_单据 In (Select 填写数量, 库房id, 零售价, 批次, 效期, 灭菌效期, 药品id 材料id, 成本价, 供药单位id, 生产日期, 批号,
                        批准文号, 产地
                 From 药品收发记录
                 Where NO = No_In And 单据 = 21
                 Order By 药品id, 批次) Loop
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = c_单据.材料id;

    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + c_单据.填写数量,
        零售价 = Decode(n_实价卫材, 1,
                         Decode(Nvl(c_单据.批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_单据.零售价, 零售价)), Null)
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And 性质 = 1;

    If Sql%NotFound Then

      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期,
         上次产地, 批准文号, 零售价)
      Values
        (c_单据.库房id, c_单据.材料id, c_单据.批次, 1, c_单据.填写数量, c_单据.效期, c_单据.灭菌效期, c_单据.供药单位id,
         c_单据.成本价, c_单据.批号, c_单据.生产日期, c_单据.产地, c_单据.批准文号,
         Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, c_单据.零售价), Null));
    End If;

    Delete From 药品库存
    Where 库房id = c_单据.库房id And 药品id = c_单据.材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
          Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
  End Loop;

  Delete From 药品收发记录 Where NO = No_In And 单据 = 21 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料其他出库_Delete;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_材料自制原料出库_Insert
(
  No_In         In 药品收发记录.NO%Type,
  对方部门id_In In 药品收发记录.对方部门id%Type
) As
  v_Err_Msg Varchar2(100);
  Err_Item Exception;

  n_数量       药品收发记录.实际数量%Type;
  n_成本价     药品收发记录.成本价%Type;
  n_成本金额   药品收发记录.成本金额%Type;
  n_差价       药品收发记录.差价%Type;
  n_售价       药品收发记录.零售价%Type;
  n_零售金额   药品收发记录.零售金额%Type;
  n_库存金额   药品库存.实际金额%Type;
  n_库存差价   药品库存.实际差价%Type;
  n_可用数量   药品库存.可用数量%Type;
  n_实际数量   药品库存.实际数量%Type;
  n_出的类别id 药品收发记录.入出类别id%Type; --入出类别ID
  n_Max_序号   药品收发记录.序号%Type;

  v_上次产地   药品库存.上次产地%Type;
  v_负成本计算 zlParameters.参数值%Type;
Begin
  Select B.ID
  Into n_出的类别id
  From 药品单据性质 A, 药品入出类别 B
  Where A.类别id = B.ID And A.单据 = 31 And B.系数 = -1 And Rownum < 2;

  Select zl_GetSysParameter(120) Into v_负成本计算 From Dual;

  Select Max(序号) Into n_Max_序号 From 药品收发记录 Where NO = No_In And 单据 = 16 And 入出系数 = 1;

  For v_自制 In (Select * From 药品收发记录 Where NO = No_In And 单据 = 16 And 入出系数 = 1 Order By 药品id, 批次) Loop

    For v_组成 In (Select A.*, B.是否变价, C.指导差价率, C.成本价
                   From 自制材料构成 A, 收费项目目录 B, 材料特性 C
                   Where A.原料材料id = B.ID And A.自制材料id = v_自制.药品id And A.原料材料id = C.材料id
				   Order By A.原料材料id) Loop

      Begin
        Select 可用数量, 实际数量, 实际差价, 实际金额, 上次产地
        Into n_可用数量, n_实际数量, n_库存差价, n_库存金额, v_上次产地
        From 药品库存
        Where 药品id = v_组成.原料材料id And 性质 = 1 And 库房id = 对方部门id_In;
      Exception
        When Others Then
          n_可用数量 := 0;
          n_实际数量 := 0;
          n_库存差价 := 0;
          n_库存金额 := 0;
      End;
      If Nvl(v_组成.是否变价, 0) = 1 Then
        --实价
        If Nvl(n_实际数量, 0) > 0 Then
          n_售价 := Nvl(n_库存金额, 0) / n_实际数量;
        Else
          --无库数:需提示
          v_Err_Msg := '[ZLSOFT]该单据中存在一笔以上原料的实际数量不足[ZLSOFT]';
          Raise Err_Item;
        End If;
      Else
        --定价,以现价为准
        Begin
          Select Nvl(现价, 0)
          Into n_售价
          From 收费价目
          Where 收费细目id = v_组成.原料材料id And
                ((Sysdate Between 执行日期 And 终止日期) Or (Sysdate >= 执行日期 And 终止日期 Is Null));

        Exception
          When Others Then
            v_Err_Msg := 'Err';
        End;
        If v_Err_Msg = 'Err' Then
          v_Err_Msg := '[ZLSOFT]该单据中存在一笔以上原料还未进行定价！[ZLSOFT]';
          Raise Err_Item;
        End If;
      End If;
      n_数量 := Nvl(v_自制.实际数量, 0) * v_组成.分子 / v_组成.分母;

      If n_数量 = 0 Then
        v_Err_Msg := '[ZLSOFT]该单据中存在一笔以上原料的数量为零了！[ZLSOFT]';
        Raise Err_Item;
      End If;
      n_零售金额 := n_数量 * n_售价;

      --算成本价
      If Nvl(n_库存金额, 0) <= 0 Then
        If v_负成本计算 = '1' And Nvl(v_组成.成本价, 0) > 0 Then
          n_成本价 := v_组成.成本价;
          n_差价   := n_零售金额 - n_数量 * n_成本价;
        Else
          n_差价   := n_零售金额 * v_组成.指导差价率 / 100;
          n_成本价 := (n_零售金额 - n_差价) / n_数量;
        End If;
      Else
        n_差价   := n_零售金额 * (n_库存差价 / n_库存金额);
        n_成本价 := (n_零售金额 - n_差价) / n_数量;
      End If;
      n_成本价 := Nvl(n_成本价, 0);

      n_成本金额 := n_成本价 * n_数量;
      n_Max_序号 := n_Max_序号 + 1;

      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 产地, 填写数量, 实际数量,
         成本价, 成本金额, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 费用id, 扣率)
      Values
        (药品收发记录_Id.Nextval, 1, 16, No_In, n_Max_序号, v_自制.对方部门id, v_自制.库房id, n_出的类别id, -1,
         v_组成.原料材料id, v_上次产地, n_数量, n_数量, n_成本价, n_成本金额, n_售价, n_零售金额, n_差价, v_自制.摘要,
         v_自制.填制人, v_自制.填制日期, v_自制.药品id, v_自制.序号);

      --IF n_可用数量<0 then
      --    v_Err_Msg:='[ZLSOFT]该单据中存在一笔以上原料的可用数量不足[ZLSOFT]';
      --    RAISE Err_Item;
      --End IF ;

      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) - n_数量
      Where 库房id = v_自制.对方部门id And 药品id = v_组成.原料材料id And 性质 = 1;

      If Sql%NotFound Then
        Insert Into 药品库存
          (库房id, 药品id, 性质, 可用数量)
        Values
          (v_自制.对方部门id, v_组成.原料材料id, 1, -n_数量);
      End If;

      Delete From 药品库存
      Where 库房id = v_自制.对方部门id And 药品id = v_组成.原料材料id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And
            Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;

    End Loop;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料自制原料出库_Insert;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
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

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_自制材料入库_Strike
(
  No_In     In 药品收发记录.No%Type,
  审核人_In In 药品收发记录.审核人%Type
) Is
  Err_Item Exception;
  v_Err_Msg    Varchar2(500);
  n_实价卫材   收费项目目录.是否变价%Type;
  n_平均成本价 药品库存.平均成本价%Type;

Begin
  Update 药品收发记录 Set 记录状态 = 3 Where NO = No_In And 单据 = 16 And 记录状态 = 1;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
    Raise Err_Item;
  End If;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 审核人, 审核日期, 费用id, 扣率)
    Select 药品收发记录_Id.Nextval, 2, 16, No_In, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 效期, -填写数量, -实际数量, 成本价,
           -成本金额, 零售价, -零售金额, -差价, 摘要, 审核人_In, Sysdate, 审核人_In, Sysdate, 费用id, 扣率
    From 药品收发记录
    Where NO = No_In And 单据 = 16 And 记录状态 = 3;

  For c_单据 In (Select ID, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 零售价, 填写数量, 实际数量, 成本价, 零售金额, 差价
               From 药品收发记录 A
               Where NO = No_In And 单据 = 16 And 记录状态 = 2
			   Order By 药品id,批次) Loop
  
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = c_单据.药品id;
  
    --更改材料库存表的相应数据
    --自制材料与原料材料的处理通过入出系数来实现
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + Nvl(c_单据.填写数量, 0) * c_单据.入出系数, 实际数量 = Nvl(实际数量, 0) + Nvl(c_单据.填写数量, 0) * c_单据.入出系数,
        实际金额 = Nvl(实际金额, 0) + Nvl(c_单据.零售金额, 0) * c_单据.入出系数, 实际差价 = Nvl(实际差价, 0) + Nvl(c_单据.差价, 0) * c_单据.入出系数,
        上次采购价 = Nvl(c_单据.成本价, 上次采购价), 上次批号 = Nvl(c_单据.批号, 上次批号), 上次产地 = Nvl(c_单据.产地, 上次产地), 效期 = Nvl(c_单据.效期, 效期),
        零售价 = Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, c_单据.零售价, 零售价)), Null)
    Where 库房id = c_单据.库房id And 药品id = c_单据.药品id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次采购价, 上次批号, 上次产地, 效期, 零售价, 平均成本价)
      Values
        (c_单据.库房id, c_单据.药品id, c_单据.批次, 1, c_单据.填写数量 * c_单据.入出系数, c_单据.填写数量 * c_单据.入出系数, c_单据.零售金额 * c_单据.入出系数,
         c_单据.差价 * c_单据.入出系数, c_单据.成本价, c_单据.批号, c_单据.产地, c_单据.效期,
         Decode(n_实价卫材, 1, Decode(Nvl(c_单据.批次, 0), 0, Null, c_单据.零售价), Null), c_单据.成本价);
    End If;
  
    Delete From 药品库存
    Where 库房id = c_单据.库房id And 药品id = c_单据.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  
    --重新计算库存表中的平均成本价
    Update 药品库存
    Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
    Where 性质 = 1 And 库房id = c_单据.库房id And 药品id = c_单据.药品id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And Nvl(实际数量, 0) <> 0;
    If Sql%NotFound Then
      Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = c_单据.药品id;
      Update 药品库存
      Set 平均成本价 = n_平均成本价
      Where 性质 = 1 And 库房id = c_单据.库房id And 药品id = c_单据.药品id And Nvl(批次, 0) = Nvl(c_单据.批次, 0) And
            Nvl(平均成本价, 0) <> c_单据.成本价;
    End If;
    Zl_材料收发记录_调价修正(c_单据.Id);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_自制材料入库_Strike;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_自制材料入库_Verify
(
  No_In     In 药品收发记录.No%Type := Null,
  审核人_In In 药品收发记录.审核人%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  v_负成本计算 Zlparameters.参数值%Type;

  n_实际库存金额 药品库存.实际金额%Type;
  n_实际库存差价 药品库存.实际差价%Type;
  n_实际库存数量 药品库存.实际数量%Type;
  n_出库差价     药品库存.实际差价%Type;
  n_实价卫材     收费项目目录.是否变价%Type;

  n_成本价     药品收发记录.成本价%Type;
  n_成本金额   药品收发记录.成本金额%Type;
  n_差价率     Number(18, 8);
  n_小数       Number(2);
  n_平均成本价 药品库存.平均成本价%Type;

Begin
  Select zl_GetSysParameter(120) Into v_负成本计算 From Dual;

  Update 药品收发记录
  Set 审核人 = Nvl(审核人_In, 审核人), 审核日期 = Sysdate
  Where NO = No_In And 单据 = 16 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_小数 From Dual;

  Update 药品收发记录 Set 成本金额 = 0 Where NO = No_In And 单据 = 16 And 记录状态 = 1 And 入出系数 = 1;

  For v_原料 In (Select ID, 实际数量, 零售价, 零售金额, 差价, 库房id, 药品id, 批次, 成本价, 批号, 效期, 产地, 灭菌效期, 批准文号, 生产日期, 入出类别id, 入出系数, 对方部门id,
                      费用id As 自制材料id, Trunc(扣率) As 序号
               From 药品收发记录
               Where NO = No_In And 单据 = 16 And 记录状态 = 1 And 入出系数 = -1
               Order By 药品id, 批次) Loop
  
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = v_原料.药品id;
  
    Begin
      Select Nvl(实际金额, 0), Nvl(实际差价, 0), Nvl(实际数量, 0)
      Into n_实际库存金额, n_实际库存差价, n_实际库存数量
      From 药品库存
      Where 药品id = v_原料.药品id And Nvl(批次, 0) = Nvl(v_原料.批次, 0) And 库房id = v_原料.库房id And 性质 = 1 And Rownum = 1;
    Exception
      When Others Then
        n_实际库存金额 := 0;
        n_实际库存数量 := 0;
    End;
  
    If n_实际库存金额 <= 0 Then
      If (n_实际库存金额 - n_实际库存差价) <= 0 Or n_实际库存数量 <= 0 Then
      
        Begin
          Select 指导差价率 / 100 Into n_差价率 From 材料特性 Where 材料id = v_原料.药品id;
        Exception
          When Others Then
            n_差价率 := 0;
        End;
        If v_负成本计算 = '1' Then
          Begin
            Select Nvl(成本价, 0) Into n_成本价 From 材料特性 Where 材料id = v_原料.药品id;
          Exception
            When Others Then
              n_成本价 := 0;
          End;
          If n_成本价 = 0 Then
            n_出库差价 := Round(v_原料.零售金额 * n_差价率, 4);
          Else
            n_出库差价 := Round(v_原料.零售金额 - v_原料.实际数量 * n_成本价, 4);
          End If;
        Else
          n_出库差价 := Round(v_原料.零售金额 * n_差价率, n_小数);
        End If;
      Else
        --主要处理零售价为零的情况，从而造成无成本的问题
        n_成本价   := ((n_实际库存金额 - n_实际库存差价) / n_实际库存数量);
        n_出库差价 := Round(v_原料.零售金额 - n_成本价 * v_原料.实际数量, n_小数);
      End If;
    Else
      n_差价率   := n_实际库存差价 / n_实际库存金额;
      n_出库差价 := Round(v_原料.零售金额 * n_差价率, n_小数);
    End If;
  
    If Nvl(v_原料.实际数量, 0) = 0 Then
      n_成本价 := (v_原料.零售金额 - n_出库差价);
    Else
      n_成本价 := (v_原料.零售金额 - n_出库差价) / v_原料.实际数量;
    End If;
    n_成本价   := Nvl(n_成本价, 0);
    n_成本金额 := Round(n_成本价 * v_原料.实际数量, n_小数);
  
    Update 药品收发记录 Set 成本价 = n_成本价, 成本金额 = n_成本金额, 差价 = n_出库差价 Where ID = v_原料.Id;
  
    Update 药品库存
    Set 实际数量 = Nvl(实际数量, 0) - Nvl(v_原料.实际数量, 0), 实际金额 = Nvl(实际金额, 0) - Nvl(v_原料.零售金额, 0), 实际差价 = Nvl(实际差价, 0) - n_出库差价
    Where 库房id = v_原料.库房id And 药品id = v_原料.药品id And Nvl(批次, 0) = Nvl(v_原料.批次, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次批号, 上次生产日期, 效期, 上次产地, 灭菌效期, 批准文号, 上次采购价, 零售价,平均成本价)
      Values
        (v_原料.库房id, v_原料.药品id, Decode(v_原料.批次, Null, Null, 0, Null, v_原料.批次), 1,
         Decode(v_原料.入出系数, 1, Nvl(v_原料.实际数量, 0), 0), v_原料.实际数量 * v_原料.入出系数, v_原料.零售金额 * v_原料.入出系数, n_出库差价 * v_原料.入出系数,
         v_原料.批号, v_原料.生产日期, v_原料.效期, v_原料.产地, v_原料.灭菌效期, v_原料.批准文号, n_成本价,
         Decode(n_实价卫材, 1, Decode(Nvl(v_原料.批次, 0), 0, Null, v_原料.零售价), Null),n_成本价);
    End If;
    Delete From 药品库存
    Where 库房id = v_原料.库房id And 药品id = v_原料.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  
    Update 药品收发记录
    Set 成本金额 = Nvl(成本金额, 0) + n_成本金额
    Where NO = No_In And 序号 = v_原料.序号 And 药品id = Nvl(v_原料.自制材料id, 0) And 单据 = 16 And 记录状态 = 1 And 入出系数 = 1;
  End Loop;

  For v_自制材料 In (Select ID, 成本金额, 零售价, 实际数量, 零售金额, 差价, 库房id, 药品id, 批次, 成本价, 批号, 效期, 产地, 灭菌效期, 批准文号, 生产日期, 入出类别id, 入出系数,
                        对方部门id, 费用id As 自制材料id, Trunc(扣率) As 序号
                 From 药品收发记录
                 Where NO = No_In And 单据 = 16 And 记录状态 = 1 And 入出系数 = 1
                 Order By 药品id, 批次) Loop
    Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = v_自制材料.药品id;
  
    n_成本金额 := Nvl(v_自制材料.成本金额, 0);
    If Nvl(v_自制材料.实际数量, 0) <> 0 Then
      n_成本价 := n_成本金额 / Nvl(v_自制材料.实际数量, 0);
    Else
      n_成本价 := n_成本金额;
    End If;
    n_出库差价 := v_自制材料.零售金额 - n_成本金额;
  
    Update 药品收发记录 Set 成本价 = n_成本价, 成本金额 = n_成本金额, 差价 = n_出库差价 Where ID = v_自制材料.Id;
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + Decode(v_自制材料.入出系数, 1, Nvl(v_自制材料.实际数量, 0), 0), 实际数量 = Nvl(实际数量, 0) + Nvl(v_自制材料.实际数量, 0),
        实际金额 = Nvl(实际金额, 0) + Nvl(v_自制材料.零售金额, 0), 实际差价 = Nvl(实际差价, 0) + n_出库差价,
        零售价 = Decode(n_实价卫材, 1, Decode(Nvl(v_自制材料.批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, v_自制材料.零售价, 零售价)), Null)
    Where 库房id = v_自制材料.库房id And 药品id = v_自制材料.药品id And Nvl(批次, 0) = Nvl(v_自制材料.批次, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次批号, 上次生产日期, 效期, 上次产地, 灭菌效期, 批准文号, 上次采购价, 零售价,平均成本价)
      Values
        (v_自制材料.库房id, v_自制材料.药品id, Decode(Nvl(v_自制材料.批次, 0), 0, Null, v_自制材料.批次), 1,
         Decode(v_自制材料.入出系数, 1, Nvl(v_自制材料.实际数量, 0), 0), v_自制材料.实际数量 * v_自制材料.入出系数, v_自制材料.零售金额 * v_自制材料.入出系数,
         n_出库差价 * v_自制材料.入出系数, v_自制材料.批号, v_自制材料.生产日期, v_自制材料.效期, v_自制材料.产地, v_自制材料.灭菌效期, v_自制材料.批准文号, n_成本价,
         Decode(n_实价卫材, 1, Decode(Nvl(v_自制材料.批次, 0), 0, Null, v_自制材料.零售价), Null),n_成本价);
    End If;
    Delete From 药品库存
    Where 库房id = v_自制材料.库房id And 药品id = v_自制材料.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  
    --更新该材料的成本价
    Update 材料特性 Set 成本价 = n_成本价 Where 材料id = v_自制材料.药品id;
  
    --重新计算库存表中的平均成本价
    Update 药品库存
    Set 平均成本价 = Decode(Nvl(批次, 0), 0, decode((实际金额 - 实际差价) / 实际数量,0,上次采购价,(实际金额 - 实际差价) / 实际数量), 上次采购价)
    Where 药品id = v_自制材料.药品id And Nvl(批次, 0) = Nvl(v_自制材料.批次, 0) And 库房id = v_自制材料.库房id And 性质 = 1 And Nvl(实际数量, 0) <> 0;
    If Sql%NotFound Then
      Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = v_自制材料.药品id;
      Update 药品库存
      Set 平均成本价 = n_平均成本价
      Where 药品id = v_自制材料.药品id And 库房id = v_自制材料.库房id And Nvl(批次, 0) = Nvl(v_自制材料.批次, 0) and 性质=1;
    End If;
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_自制材料入库_Verify;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
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
                      Nvl(a.费用id, 0) As 费用id, 序号
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
  
    --重新计算库存表中的平均成本价
    Update 药品库存
    Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
    Where 药品id = v_收发.药品id And Nvl(批次, 0) = Nvl(v_收发.批次, 0) And 库房id = v_收发.库房id And 性质 = 1 And Nvl(实际数量, 0) <> 0;
    If Sql%NotFound Then
      Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = v_收发.药品id;
      Update 药品库存
      Set 平均成本价 = n_平均成本价
      Where 药品id = v_收发.药品id And 库房id = v_收发.库房id And Nvl(批次, 0) = Nvl(v_收发.批次, 0) And 性质 = 1;
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
                 Order By 药品id, 批次,序号) Loop
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
                   Order By 药品id, 批次,序号) Loop
    
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

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_材料外购_Delete(
                                           --删除药品收发记录及相应的表：应付记录
                                           No_In In 药品收发记录.No%Type) Is
  Merritem Exception;
  Merrmsg Varchar2(100);
Begin

  --恢复可用数量
  For v_收发 In (Select 实际数量, 库房id, 批次, 药品id, 成本价, 批号, 生产日期, 灭菌效期, 效期, 产地, 供药单位id, 批准文号
               From 药品收发记录
               Where NO = No_In And Nvl(发药方式, 0) = 1 And 单据 = 15
               Order By 药品id,批次,序号) Loop
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + (-1 * v_收发.实际数量)
    Where 库房id = v_收发.库房id And 药品id = v_收发.药品id And Nvl(批次, 0) = Nvl(v_收发.批次, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 灭菌效期, 上次生产日期, 批准文号)
      Values
        (v_收发.库房id, v_收发.药品id, v_收发.批次, 1, -1 * v_收发.实际数量, v_收发.供药单位id, v_收发.成本价, v_收发.批号, v_收发.产地, v_收发.效期, v_收发.灭菌效期,
         v_收发.生产日期, v_收发.批准文号);
    End If;
  
    Delete From 药品库存
    Where 库房id = v_收发.库房id And 药品id = v_收发.药品id And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
          Nvl(实际差价, 0) = 0;
  
  End Loop;

  Delete 应付记录 Where 系统标识 = 5 And 收发id In (Select ID From 药品收发记录 Where NO = No_In And 单据 = 15);
  --对应应付记录的删除通过级联删除
  Delete --删除本身
  From 药品收发记录
  Where NO = No_In And 单据 = 15 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Merrmsg := '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]';
    Raise Merritem;
  End If;
Exception
  When Merritem Then
    Raise_Application_Error(-20101, Merrmsg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料外购_Delete;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_药品盘点_Delete(
                                           --删除药品收发记录及恢复相应的表：药品库存
                                           No_In In 药品收发记录.No%Type) Is
  Err_Isverified Exception;

  Cursor c_药品收发记录 Is
    Select ID From 药品收发记录 Where NO = No_In And 单据 = 12 Order By 药品id,批次;
Begin
  --通过循环，恢复出库类别原来的可用数量，
  --实际数量保存的是数量差
  For v_药品收发记录 In c_药品收发记录 Loop
    Zl_药品库存_Update(v_药品收发记录.Id, 1);
  End Loop;

  Delete From 药品收发记录 Where NO = No_In And 单据 = 12 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品盘点_Delete;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_药品盘点_Verify
(
  No_In     In 药品收发记录.No%Type := Null,
  审核人_In In 药品收发记录.审核人%Type := Null
) Is
  Err_Isverified Exception;

  Cursor c_药品收发记录 Is
    Select ID, 实际数量, 零售金额, 差价, 库房id, 药品id, 批次, 批号, 效期, 产地, 入出类别id, 入出系数, 批准文号, 供药单位id, 生产日期, 单量
    From 药品收发记录
    Where NO = No_In And 单据 = 12 And 记录状态 = 1
    Order By 药品id,批次;
Begin
  Update 药品收发记录
  Set 审核人 = 审核人_In, 审核日期 = Sysdate
  Where NO = No_In And 单据 = 12 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  For v_药品收发记录 In c_药品收发记录 Loop
    --更改药品库存表的相应数据
    Zl_药品库存_Update(v_药品收发记录.Id, 0);
  End Loop;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品盘点_Verify;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_药品盘点_Strike
(
  No_In     In 药品收发记录.No%Type,
  审核人_In In 药品收发记录.审核人%Type
) Is
  Err_Isstriked Exception;

  Cursor c_药品收发记录 Is
    Select a.Id, a.实际数量, a.零售金额, a.差价, a.零售价, Nvl(b.是否变价, 0) As 是否变价, a.库房id, a.药品id, a.批次, a.批号, a.效期, a.产地, a.入出类别id,
           a.入出系数, a.单量, a.批准文号, a.供药单位id, a.生产日期
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And NO = No_In And 单据 = 12 And 记录状态 = 2
    Order By 药品id,批次;
Begin
  Update 药品收发记录 Set 记录状态 = 3 Where NO = No_In And 单据 = 12 And 记录状态 = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 扣率, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要, 填制人,
     填制日期, 审核人, 审核日期, 频次, 单量, 批准文号, 供药单位id, 生产日期, 库房货位)
    Select 药品收发记录_Id.Nextval, 2, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, a.药品id,
           Decode(Nvl(a.批次, 0), 0, Null, (Decode(Nvl(b.药库分批, 0), 0, Null, a.批次))), a.产地, 批号, 效期, 填写数量, a.扣率, -实际数量,
           a.成本价, 成本金额, 零售价, -零售金额, -差价, 摘要, 审核人_In, Sysdate, 审核人_In, Sysdate, 频次, 单量, a.批准文号, a.供药单位id, a.生产日期, a.库房货位
    From (Select * From 药品收发记录 Where NO = No_In And 单据 = 12 And 记录状态 = 3 Order By 药品id) A, 药品规格 B
    Where a.药品id = b.药品id;

  For v_药品收发记录 In c_药品收发记录 Loop
    --处理库存
    Zl_药品库存_Update(v_药品收发记录.Id, 0);
  
    --处理调价后冲销
    Zl_药品收发记录_调价修正(v_药品收发记录.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品盘点_Strike;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_药品移库_Verify
(
  序号_In         In 药品收发记录.序号%Type,
  库房id_In       In 药品收发记录.库房id%Type,
  对方部门id_In   In 药品收发记录.对方部门id%Type,
  药品id_In       In 药品收发记录.药品id%Type,
  产地_In         In 药品收发记录.产地%Type,
  出批次_In       In 药品收发记录.批次%Type,
  实际数量_In     In 药品收发记录.实际数量%Type,
  成本价_In       In 药品收发记录.成本价%Type,
  成本金额_In     In 药品收发记录.成本金额%Type,
  零售金额_In     In 药品收发记录.零售金额%Type,
  差价_In         In 药品收发记录.差价%Type,
  No_In           In 药品收发记录.No%Type,
  审核人_In       In 药品收发记录.审核人%Type,
  批号_In         In 药品收发记录.批号%Type := Null,
  效期_In         In 药品收发记录.效期%Type := Null,
  审核日期_In     In 药品收发记录.审核日期%Type := Null,
  上次供应商id_In In 药品收发记录.供药单位id%Type := Null,
  批准文号_In     In 药品收发记录.批准文号%Type := Null,
  零售价_In       In 药品收发记录.零售价%Type := Null
) Is
  Err_Isverified Exception;
  Err_Isnonumber Exception;
  Err_Isbatch Exception;
  Err_Isprice Exception;
  v_Druginf  Varchar2(50); --原不分批现在分批的药品信息 
  v_实际数量 药品库存.实际数量%Type;
  v_编码     收费项目目录.编码%Type;
  Intdigit   Number;
  v_上次扣率 药品库存.上次扣率%Type;
  Cursor c_药品收发记录 Is
    Select ID
    From 药品收发记录
    Where NO = No_In And 单据 = 6 And 药品id = 药品id_In And 记录状态 = 1 And 序号 In (序号_In, 序号_In + 1) And 审核日期 Is Not Null
	Order By 药品id,批次;
Begin
  --获取金额小数位数 
  Select Nvl(精度, 2) Into Intdigit From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;

  --主要针对原不分批现在分批的药品，不能对其审核 
  --仅检查入类别 
  Begin
    Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
    Into v_Druginf
    From 药品收发记录 A, 药品规格 B, 收费项目目录 I, 收费项目别名 N
    Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 6 And
          a.记录状态 = 1 And Nvl(a.批次, 0) = 0 And a.药品id + 0 = 药品id_In And a.序号 = 序号_In + 1 And
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

  Begin
    Select Nvl(实际数量, 0), Nvl(上次扣率, 100)
    Into v_实际数量, v_上次扣率
    From 药品库存
    Where 药品id = 药品id_In And Nvl(批次, 0) = 出批次_In And 库房id = 库房id_In And 性质 = 1 And Rownum = 1;
  Exception
    When Others Then
      v_实际数量 := 0;
      v_上次扣率 := 100;
  End;

  If 出批次_In > 0 Then
    If v_实际数量 < 实际数量_In Then
      Raise Err_Isnonumber;
    End If;
  End If;

  Update 药品收发记录
  Set 审核人 = Nvl(审核人_In, 审核人), 审核日期 = 审核日期_In, 实际数量 = 实际数量_In, 成本价 = 成本价_In, 成本金额 = 成本金额_In, 零售价 = 零售价_In, 零售金额 = 零售金额_In,
      差价 = 差价_In, 扣率 = v_上次扣率
  Where NO = No_In And 单据 = 6 And 药品id = 药品id_In And 记录状态 = 1 And 序号 In (序号_In, 序号_In + 1) And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  --更新入出库库存
  For v_药品收发记录 In c_药品收发记录 Loop
    Zl_药品库存_Update(v_药品收发记录.Id);
  End Loop;

Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]');
  When Err_Isnonumber Then
    Select 编码 Into v_编码 From 收费项目目录 Where ID = 药品id_In;
    Raise_Application_Error(-20101,
                            '[ZLSOFT]编码为' || v_编码 || ',批号为' || 批号_In || '的药库分批药品' || Chr(10) || Chr(13) ||
                             '可用库存数量不够！[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品[' || v_Druginf || ']，不能审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品移库_Verify;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_药品移库_Delete
(
  No_In       In 药品收发记录.No%Type,
  记录状态_In In 药品收发记录.记录状态%Type := 1
) Is
  Err_Isverified Exception;

  Cursor c_药品收发记录 Is
    Select ID From 药品收发记录 Where NO = No_In And 单据 = 6 Order By 药品id,批次;

  Cursor c_申请冲销记录 Is
    Select ID From 药品收发记录 Where NO = No_In And 单据 = 6 And 记录状态 = 记录状态_In Order By 药品id,批次;
Begin

  If 记录状态_In = 1 Then
    --处理未审核移库单据
    --通过循环，恢复原来的可用数量
    For v_药品收发记录 In c_药品收发记录 Loop
      Zl_药品库存_Update(v_药品收发记录.Id, 1);
    End Loop;
  Else
    --处理移库申请冲销单据
    --通过循环，恢复原来的可用数量
    For v_申请冲销记录 In c_申请冲销记录 Loop
      Zl_药品库存_Update(v_申请冲销记录.Id, 1);
    End Loop;
  End If;

  --删除未审核单据
  Delete From 药品收发记录 Where NO = No_In And 单据 = 6 And 记录状态 = 记录状态_In And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品移库_Delete;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_药品移库_Strike
(
  行次_In       In Integer,
  原记录状态_In In 药品收发记录.记录状态%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  药品id_In     In 药品收发记录.药品id%Type,
  冲销数量_In   In 药品收发记录.实际数量%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  冲销方式_In   In Integer := 0 --0－正常冲销方式；1－产生冲销申请单据；2－审核已产生的冲销申请单据
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf      Varchar2(50); --原不分批现在分批的药品信息
  v_库房id       药品收发记录.库房id%Type;
  v_批次         药品收发记录.批次%Type;
  v_成本价       药品收发记录.成本价%Type;
  v_成本金额     药品收发记录.成本金额%Type;
  v_零售价       药品收发记录.零售价%Type;
  v_零售金额     药品收发记录.零售金额%Type;
  v_差价         药品收发记录.差价%Type;
  v_剩余数量     药品收发记录.实际数量%Type;
  v_剩余成本金额 药品收发记录.成本金额%Type;
  v_剩余零售金额 药品收发记录.零售金额%Type;
  v_收发id       药品收发记录.Id%Type;
  v_批准文号     药品收发记录.批准文号%Type;

  v_药库分批 Integer;
  v_药房分批 Integer;
  Intdigit   Number;

  Cursor c_药品收发记录 Is
    Select a.Id, a.序号, a.库房id, a.对方部门id, a.入出类别id, a.入出系数, a.药品id, a.批次, a.产地, a.批号, a.效期, a.配药人, a.配药日期, a.摘要, a.供药单位id,
           a.批准文号, a.生产日期, a.成本价, a.零售价, Nvl(b.是否变价, 0) As 时价, a.扣率, a.单量, a.频次
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.No = No_In And a.单据 = 6 And (a.序号 >= 序号_In And a.序号 <= 序号_In + 1) And
          (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id,a.批次;

  Cursor c_冲销申请记录 Is
    Select a.Id, a.序号, a.库房id, a.对方部门id, a.入出类别id, a.入出系数, a.药品id, a.批次, a.产地, a.批号, a.效期, a.配药人, a.配药日期, a.摘要, a.供药单位id,
           a.批准文号, a.生产日期, a.成本价, a.实际数量, a.零售金额, a.差价, a.零售价, Nvl(b.是否变价, 0) As 时价, a.扣率, a.单量, a.频次
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.No = No_In And a.单据 = 6 And (a.序号 >= 序号_In And a.序号 <= 序号_In + 1) And
          (a.记录状态 = 原记录状态_In And Mod(a.记录状态, 3) = 2) And a.审核日期 Is Null
    Order By a.药品id,a.批次;
Begin
  --获取金额小数位数
  Select Nvl(精度, 2) Into Intdigit From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;

  If 冲销方式_In = 1 Then
    --产生冲销申请单据，不填写审核人、审核日期，不更新库存记录
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where NO = No_In And 单据 = 6 And 记录状态 = 原记录状态_In;
      If Sql%RowCount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --主要针对原不分批现在分批的药品，不能对其冲销
    Begin
      Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
      Into v_Druginf
      From 药品收发记录 A, 药品规格 B, 收费项目目录 I, 收费项目别名 N
      Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 6 And
            a.药品id + 0 = 药品id_In And Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And a.序号 = 序号_In And
            ((Nvl(b.药库分批, 0) = 1 And
            a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or
            Nvl(b.药房分批, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(a.实际数量) As 剩余数量, Sum(a.成本金额) As 剩余成本金额, Sum(a.零售金额) As 剩余零售金额, a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0),
           b.药库分批, b.药房分批, a.批准文号
    Into v_剩余数量, v_剩余成本金额, v_剩余零售金额, v_成本价, v_零售价, v_库房id, v_批次, v_药库分批, v_药房分批, v_批准文号
    From 药品收发记录 A, 药品规格 B
    Where a.No = No_In And a.药品id = b.药品id And a.单据 = 6 And a.药品id = 药品id_In And a.序号 = 序号_In
    Group By a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0), b.药库分批, b.药房分批, a.批准文号;
  
    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    --因为对于对方库房，相当于出库，而对于当前库房，相当于入库，所以当前库房不予检查，仅检查退库那条记录
    Select Nvl(a.批次, 0)
    Into v_批次
    From 药品收发记录 A
    Where a.No = No_In And a.单据 = 6 And a.药品id = 药品id_In And a.序号 = 序号_In + 1 And Mod(a.记录状态, 3) = 0;
  
    --冲销数量大于剩余数量，不允许
    If v_剩余数量 < 冲销数量_In Then
      Raise Err_Isnonum;
    End If;
  
    v_成本金额 := Round(冲销数量_In / v_剩余数量 * v_剩余成本金额, Intdigit);
    v_零售金额 := Round(冲销数量_In / v_剩余数量 * v_剩余零售金额, Intdigit);
    v_差价     := v_零售金额 - v_成本金额;
  
    For v_药品收发记录 In c_药品收发记录 Loop
    
      Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价,
         摘要, 填制人, 填制日期, 审核人, 审核日期, 配药人, 配药日期, 供药单位id, 批准文号, 生产日期, 扣率, 单量, 频次)
      Values
        (v_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 6, No_In, v_药品收发记录.序号, v_药品收发记录.库房id, v_药品收发记录.对方部门id,
         v_药品收发记录.入出类别id, v_药品收发记录.入出系数, 药品id_In, v_药品收发记录.批次, v_药品收发记录.产地, v_药品收发记录.批号, v_药品收发记录.效期, -冲销数量_In, -冲销数量_In,
         v_成本价, -v_成本金额, v_零售价, -v_零售金额, -v_差价, 摘要_In, 填制人_In, 填制日期_In, Null, Null, v_药品收发记录.配药人, v_药品收发记录.配药日期,
         v_药品收发记录.供药单位id, v_药品收发记录.批准文号, v_药品收发记录.生产日期, v_药品收发记录.扣率, v_药品收发记录.单量, v_药品收发记录.频次);
    
      --原入库库房再勾选填单下库存时要下库存
      Zl_药品库存_Update(v_收发id, 0, 1);
    End Loop;
  
  Elsif 冲销方式_In = 2 Then
    --审核已产生的冲销申请单据，填写审核人、审核日期，更新库存记录
    For v_药品收发记录 In c_冲销申请记录 Loop
      --填写审核人、审核日期
      Update 药品收发记录
      Set 审核人 = 填制人_In, 审核日期 = 填制日期_In
      Where NO = No_In And 单据 = 6 And ID = v_药品收发记录.Id;
    
      --更改药品库存表的相应数据，注意这时传入的数量等是负数
      --参数为1表示申请冲销时下可用数量，仅对原移入库房，下了可用数量就不用再更新可用数量了
      Zl_药品库存_Update(v_药品收发记录.Id, 0, 1);
    
      --处理调价后冲销
      Zl_药品收发记录_调价修正(v_药品收发记录.Id);
    End Loop;
  Else
    --正常冲销方式，产生冲销记录，填写审核人、审核日期，更新库存记录
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where NO = No_In And 单据 = 6 And 记录状态 = 原记录状态_In;
      If Sql%RowCount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --主要针对原不分批现在分批的药品，不能对其冲销
    Begin
      Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
      Into v_Druginf
      From 药品收发记录 A, 药品规格 B, 收费项目目录 I, 收费项目别名 N
      Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 6 And
            a.药品id + 0 = 药品id_In And Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And a.序号 = 序号_In And
            ((Nvl(b.药库分批, 0) = 1 And
            a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or
            Nvl(b.药房分批, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(a.实际数量) As 剩余数量, Sum(a.成本金额) As 剩余成本金额, Sum(a.零售金额) As 剩余零售金额, a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0),
           b.药库分批, b.药房分批, a.批准文号
    Into v_剩余数量, v_剩余成本金额, v_剩余零售金额, v_成本价, v_零售价, v_库房id, v_批次, v_药库分批, v_药房分批, v_批准文号
    From 药品收发记录 A, 药品规格 B
    Where a.No = No_In And a.药品id = b.药品id And a.单据 = 6 And a.药品id = 药品id_In And a.序号 = 序号_In
    Group By a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0), b.药库分批, b.药房分批, a.批准文号;
  
    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    --因为对于对方库房，相当于出库，而对于当前库房，相当于入库，所以当前库房不予检查，仅检查退库那条记录
    Select Nvl(a.批次, 0)
    Into v_批次
    From 药品收发记录 A
    Where a.No = No_In And a.单据 = 6 And a.药品id = 药品id_In And a.序号 = 序号_In + 1 And Mod(a.记录状态, 3) = 0;
  
    --冲销数量大于剩余数量，不允许
    If v_剩余数量 < 冲销数量_In Then
      Raise Err_Isnonum;
    End If;
  
    v_成本金额 := Round(冲销数量_In / v_剩余数量 * v_剩余成本金额, Intdigit);
    v_零售金额 := Round(冲销数量_In / v_剩余数量 * v_剩余零售金额, Intdigit);
    v_差价     := v_零售金额 - v_成本金额;
  
    For v_药品收发记录 In c_药品收发记录 Loop
      Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价,
         摘要, 填制人, 填制日期, 审核人, 审核日期, 配药人, 配药日期, 供药单位id, 批准文号, 生产日期, 扣率, 单量, 频次)
      Values
        (v_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 6, No_In, v_药品收发记录.序号, v_药品收发记录.库房id, v_药品收发记录.对方部门id,
         v_药品收发记录.入出类别id, v_药品收发记录.入出系数, 药品id_In, v_药品收发记录.批次, v_药品收发记录.产地, v_药品收发记录.批号, v_药品收发记录.效期, -冲销数量_In, -冲销数量_In,
         v_成本价, -v_成本金额, v_零售价, -v_零售金额, -v_差价, 摘要_In, 填制人_In, 填制日期_In, 填制人_In, 填制日期_In, v_药品收发记录.配药人, v_药品收发记录.配药日期,
         v_药品收发记录.供药单位id, v_药品收发记录.批准文号, v_药品收发记录.生产日期, v_药品收发记录.扣率, v_药品收发记录.单量, v_药品收发记录.频次);
    
      --更改药品库存表的相应数据
      Zl_药品库存_Update(v_收发id, 0, 0);
    
      --处理调价后冲销
      Zl_药品收发记录_调价修正(v_收发id);
    End Loop;
  End If;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品[' || v_Druginf || ']，不能冲销！[ZLSOFT]');
  When Err_Isnonum Then
    Raise_Application_Error(-20103, '[ZLSOFT]该单据中第' || Ceil(序号_In / 2) || '行的药品冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品移库_Strike;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_药品其他出库_Delete(No_In In 药品收发记录.No%Type) Is
  Err_Isverified Exception;

  Cursor c_药品收发记录 Is
    Select ID From 药品收发记录 Where NO = No_In And 单据 = 11 Order By 药品id,批次;
Begin

  --根据系统参数，如果填单时下了可用数量，则要恢复原来的可用数量
  For v_药品收发记录 In c_药品收发记录 Loop
    Zl_药品库存_Update(v_药品收发记录.Id, 1);
  End Loop;

  Delete From 药品收发记录 Where NO = No_In And 单据 = 11 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品其他出库_Delete;
/

--117925:刘涛,2017-12-07,排序导致死锁处理
Create Or Replace Procedure Zl_药品领用_Delete(
                                           --删除药品收发记录及恢复相应的表：药品库存
                                           No_In In 药品收发记录.No%Type) Is
  Err_Isverified Exception;

  Cursor c_药品收发记录 Is
    Select ID, 填写数量, 库房id, 批次, 药品id, 批号, 效期, 产地, 批准文号, 对方部门id, 发药方式
    From 药品收发记录
    Where NO = No_In And 单据 = 7
    Order By 药品id,批次;
  v_按月留存领用 Varchar2(4000);
Begin

  Select Zl_Getsysparameter('按月留存领用', 1305) Into v_按月留存领用 From Dual;
  --通过循环，恢复原来的可用数量
  For v_药品收发记录 In c_药品收发记录 Loop
    Zl_药品库存_Update(v_药品收发记录.Id, 1);
  
    If v_药品收发记录.发药方式 = 1 Then
      Update 药品留存
      Set 可用数量 = Nvl(可用数量, 0) + v_药品收发记录.填写数量
      Where 期间 = To_Char(Sysdate, Decode(v_按月留存领用, '1', 'yyyymm', 'yyyy')) And 科室id = v_药品收发记录.对方部门id And
            库房id = v_药品收发记录.库房id And 药品id = v_药品收发记录.药品id;
    End If;
  End Loop;

  Delete From 药品收发记录 Where NO = No_In And 单据 = 7 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品领用_Delete;
/

--116996:胡俊勇,2017-12-06,静配药取消销帐申请
Create Or Replace Procedure Zl_病人医嘱记录_收回
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
    Select a.病人id, a.主页id, d.姓名, Nvl(x.剂量系数, 1) As 剂量系数, Nvl(x.住院包装, 1) As 住院包装, x.最大效期, Nvl(b.付数, 1) * b.实际数量 As 数量,
           b.Id As 收发id, b.单据, b.药品id, b.对方部门id, b.库房id, b.费用id, Nvl(x.药房分批, 0) As 分批, b.批次, b.批号, b.效期, a.记录状态, a.No,
           a.序号, a.收费细目id, a.执行状态 As 执行标志
    From 住院费用记录 A, 药品收发记录 B, 病人医嘱发送 C, 病人信息 D, 药品规格 X
    Where c.医嘱id = 医嘱id_In And a.No = c.No And a.记录性质 = c.记录性质 And a.记录状态 In (0, 1, 3) And a.医嘱序号 + 0 = 医嘱id_In And
          a.No = b.No And a.Id = b.费用id + 0 And b.单据 In (9, 10, 25, 26) And (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And
          a.病人id = d.病人id And b.药品id = x.药品id(+)
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
                    For r_Otherdrug In (Select a.病人id, a.主页id, d.姓名, Nvl(x.剂量系数, 1) As 剂量系数, Nvl(x.住院包装, 1) As 住院包装,
                                               x.最大效期, Nvl(b.付数, 1) * b.实际数量 As 数量, b.Id As 收发id, b.单据, b.药品id, b.对方部门id,
                                               b.库房id, b.费用id, Nvl(x.药房分批, 0) As 分批, b.批次, b.批号, b.效期, a.记录状态, a.No, a.序号,
                                               a.收费细目id
                                        From 住院费用记录 A, 药品收发记录 B, 病人信息 D, 药品规格 X
                                        Where a.Id = r_Other.费用id And a.记录状态 In (0, 1, 3) And a.No = b.No And
                                              a.Id = b.费用id + 0 And b.单据 In (9, 10, 25, 26) And
                                              (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And a.病人id = d.病人id And
                                              b.药品id = x.药品id(+)
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

--116996:胡俊勇,2017-12-06,静配药取消销帐申请
--116388:殷瑞,2017-11-16,修正取消申请销帐后恢复操作类型的赋值错误
Create Or Replace Procedure Zl_病人费用销帐_Delete
(
  Ids_In    In Varchar2,
  配药id_In In 输液配药记录.Id%Type := Null
) As
  n_Id  病人费用销帐.费用id%Type;
  v_Ids Varchar2(4000);

  n_医嘱id   住院费用记录.Id%Type;
  v_No       住院费用记录.No%Type;
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;
  n_操作类型 输液配药记录.操作状态%Type;
  n_销帐时间 输液配药记录.操作时间%Type;
Begin
  If 配药id_In Is Not Null Then
    Select 操作时间
    Into n_销帐时间
    From (Select 操作人员, 操作时间, 操作类型
           From 输液配药状态
           Where 配药id = 配药id_In And 操作类型 = 9
           Order By 操作时间 Desc)
    Where Rownum = 1;
  End If;
  
  v_Ids := Ids_In || ',';
  While v_Ids Is Not Null Loop
    n_Id  := To_Number(Substr(v_Ids, 1, Instr(v_Ids, ',') - 1));
    v_Ids := Substr(v_Ids, Instr(v_Ids, ',') + 1);
  
    If n_销帐时间 Is Null Then
      Delete 病人费用销帐 Where 费用id = n_Id And 状态 = 0;
      Select a.No, a.医嘱序号 Into v_No, n_医嘱id From 住院费用记录 A Where a.Id = n_Id;
      If Not n_医嘱id Is Null Then
        --暂未提供按配药批次取消的功能，所有已申请的批次一起取消
        For R In (Select d.Id
                  From 病人医嘱记录 A, 病人医嘱发送 B, 输液配药记录 D
                  Where a.Id = n_医嘱id And a.Id = b.医嘱id And b.No = v_No And a.相关id = d.医嘱id And b.发送号 = d.发送号 And
                        b.记录性质 = 2) Loop
          Select 操作人员, 操作时间, 操作类型
          Into v_操作人员, d_操作时间, n_操作类型
          From (Select 操作人员, 操作时间, 操作类型
                 From 输液配药状态
                 Where 配药id = r.Id And 操作类型 <> 9
                 Order By 操作时间 Desc, 操作类型 Desc)
          Where Rownum = 1;        
          Update 输液配药记录 Set 操作人员 = v_操作人员, 操作时间 = d_操作时间, 操作状态 = n_操作类型 Where ID = r.Id;
        End Loop;
      End If;
    Else
      Delete 病人费用销帐 Where 费用id = n_Id And 状态 = 0 And 申请时间 = n_销帐时间;    
      Select 操作人员, 操作时间, 操作类型
      Into v_操作人员, d_操作时间, n_操作类型
      From (Select 操作人员, 操作时间, 操作类型
             From 输液配药状态
             Where 配药id = 配药id_In And 操作类型 <> 9
             Order By 操作时间 Desc, 操作类型 Desc)
      Where Rownum = 1;    
      Update 输液配药记录 Set 操作人员 = v_操作人员, 操作时间 = d_操作时间, 操作状态 = n_操作类型 Where ID = 配药id_In;    
    End If;  
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人费用销帐_Delete;
/
--117925:刘涛,2017-12-06,排序导致死锁处理
Create Or Replace Procedure Zl_药品外购_Delete(No_In In 药品收发记录.No%Type) Is
  Err_Isverified Exception;

  Cursor c_药品收发记录 Is
    Select ID From 药品收发记录 Where NO = No_In And 单据 = 1 Order By 药品id,批次,序号;
Begin

  --通过循环，恢复原来的可用数量
  For v_药品收发记录 In c_药品收发记录 Loop
    --调用库存更新过程
    Zl_药品库存_Update(v_药品收发记录.Id, 1);
  End Loop;

  Delete 应付记录 Where 系统标识 = 1 And 收发id In (Select ID From 药品收发记录 Where NO = No_In And 单据 = 1);

  --对应应付记录的删除通过级联删除
  Delete --删除本身
  From 药品收发记录
  Where NO = No_In And 单据 = 1 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品外购_Delete;
/

--117925:刘涛,2017-12-06,排序导致死锁处理
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
    Order By a.药品id,a.批次;
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
    If Oldno_In Is Null Then
      Zl_药品库存_Update(v_药品收发记录.Id, 0);
    Else
      Zl_药品库存_Update(v_药品收发记录.Id, 0, 0, 0, 1);
    End If;
  
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
      Select distinct 成本价, 零售价
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

--117925:刘涛,2017-12-06,排序导致死锁处理
Create Or Replace Procedure Zl_药品其他入库_Delete(No_In In 药品收发记录.No%Type) Is
  Err_Isverified Exception;

  Cursor c_药品收发记录 Is
    Select ID From 药品收发记录 Where NO = No_In And 单据 = 4 Order By 药品id,批次,序号;
Begin
  --通过循环，恢复原来的可用数量
  For v_药品收发记录 In c_药品收发记录 Loop
    Zl_药品库存_Update(v_药品收发记录.Id, 1);
  End Loop;

  --删除药品收发记录
  Delete --删除本身
  From 药品收发记录
  Where NO = No_In And 单据 = 4 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品其他入库_Delete;
/

--117925:刘涛,2017-12-06,排序导致死锁处理
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
    Order By a.药品id,a.批次;
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
    Set 成本价 = v_药品收发记录.成本价, 上次售价 = Decode(v_药品收发记录.时价, 1, v_药品收发记录.零售价, Null),
        上次供应商id = Decode(v_药品收发记录.供药单位id, Null, 上次供应商id, v_药品收发记录.供药单位id), 上次批号 = v_药品收发记录.批号, 上次生产日期 = v_药品收发记录.生产日期,
        上次产地 = v_药品收发记录.产地, 上次批准文号 = v_药品收发记录.批准文号
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

--117925:刘涛,2017-12-06,排序导致死锁处理
Create Or Replace Procedure Zl_协定入库_Delete(No_In In 药品收发记录.No%Type) Is
  Err_Isverified Exception;

  Cursor c_药品收发记录 Is
    Select ID From 药品收发记录 Where NO = No_In And 单据 = 3 Order By 药品id,批次;
Begin
  --通过循环，恢复所有构成协定药原来的可用数量
  For v_药品收发记录 In c_药品收发记录 Loop
    Zl_药品库存_Update(v_药品收发记录.Id, 1);
  End Loop;

  Delete --删除本身及相应的构成协定药
  From 药品收发记录
  Where NO = No_In And 单据 = 3 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_协定入库_Delete;
/

--117925:刘涛,2017-12-06,排序导致死锁处理
Create Or Replace Procedure Zl_协定入库_Verify
(
  No_In     In 药品收发记录.No%Type := Null,
  审核人_In In 药品收发记录.审核人%Type := Null
) Is
  Err_Isverified Exception;
  v_出库差价 药品库存.实际差价%Type;
  v_成本价   药品收发记录.成本价%Type;
  v_成本金额 药品收发记录.成本金额%Type;

  Cursor c_药品收发记录 Is
    Select ID, 填写数量, 零售价, 零售金额, 差价, 库房id, 药品id, 批次, 成本价, 批号, 产地, 入出类别id, 入出系数, 对方部门id, 供药单位id, 生产日期, 批准文号, 效期
    From 药品收发记录
    Where NO = No_In And 单据 = 3 And 记录状态 = 1
	Order By 药品id,批次;
Begin
  Update 药品收发记录
  Set 审核人 = Nvl(审核人_In, 审核人), 审核日期 = Sysdate
  Where NO = No_In And 单据 = 3 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  For v_药品收发记录 In c_药品收发记录 Loop
    --更改药品库存表的相应数据
    If v_药品收发记录.入出系数 = -1 Then
      v_成本价   := Zl_Fun_Getoutcost(v_药品收发记录.药品id, Nvl(v_药品收发记录.批次, 0), v_药品收发记录.库房id);
      v_成本金额 := v_成本价 * v_药品收发记录.填写数量;
      v_出库差价 := v_药品收发记录.零售金额 - v_成本金额;
    Else
      Begin
        Select Sum(成本价)
        Into v_成本价
        From (Select Decode(Nvl(c.实际金额, 0), 0,
                              Decode(Nvl(c.上次采购价, 0), 0, Decode(Nvl(b.成本价, 0), 0, (d.现价 - d.现价 * (b.指导差价率 / 100)), b.成本价),
                                      c.上次采购价), (d.现价 - d.现价 * (c.实际差价 / c.实际金额))) * (a.分子 / a.分母) As 成本价               
               From 协定药品对照 A,
                    (Select b.药品id, b.成本价, b.指导差价率
                      From 收费项目目录 A, 药品规格 B
                      Where a.Id = b.药品id And Nvl(是否变价, 0) = 0) B,
                    (Select 库房id, 药品id, 实际金额, 实际差价, 上次采购价
                      From 药品库存
                      Where 性质 = 1 And 库房id = v_药品收发记录.对方部门id) C,
                    (Select 收费细目id, 现价
                      From 收费价目
                      Where ((Sysdate Between 执行日期 And 终止日期) Or (Sysdate >= 执行日期 And 终止日期 Is Null))) D
               Where a.协定药品id = b.药品id And b.药品id = d.收费细目id And b.药品id = c.药品id(+) And a.药品id = v_药品收发记录.药品id
               Union All
               Select Decode(Nvl(c.实际金额, 0), 0,
                              Decode(Nvl(c.上次采购价, 0), 0, Decode(Nvl(b.成本价, 0), 0, (c.现价 - c.现价 * (b.指导差价率 / 100)), b.成本价),
                                      c.上次采购价), (c.现价 - c.现价 * (c.实际差价 / c.实际金额))) * (a.分子 / a.分母) As 成本价               
               From 协定药品对照 A,
                    (Select b.药品id, b.成本价, b.指导差价率
                      From 收费项目目录 A, 药品规格 B
                      Where a.Id = b.药品id And Nvl(是否变价, 0) = 1) B,
                    (Select 库房id, 药品id, 实际金额, 实际差价, 实际金额 / 实际数量 As 现价, 上次采购价
                      From 药品库存
                      Where 性质 = 1 And 库房id = v_药品收发记录.对方部门id And 实际数量 > 0) C
               Where a.协定药品id = b.药品id And b.药品id = c.药品id And a.药品id = v_药品收发记录.药品id);
      Exception
        When Others Then
          v_成本价 := 0;
      End;
    
      v_成本金额 := v_成本价 * v_药品收发记录.填写数量;
      v_出库差价 := v_药品收发记录.零售金额 - v_成本金额;
    End If;
  
    Update 药品收发记录 Set 成本价 = v_成本价, 成本金额 = v_成本金额, 差价 = v_出库差价 Where ID = v_药品收发记录.Id;
  
    --更新库存
    Zl_药品库存_Update(v_药品收发记录.Id, 0);
  
  End Loop;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_协定入库_Verify;
/

--117925:刘涛,2017-12-06,排序导致死锁处理
Create Or Replace Procedure Zl_协定入库_Strike
(
  No_In     In 药品收发记录.No%Type,
  审核人_In In 药品收发记录.审核人%Type
) Is
  Err_Isstriked Exception;

  Cursor c_药品收发记录 Is
    Select ID, 库房id, 入出类别id, 入出系数, 药品id, 填写数量, 批次, 实际数量, 成本价, 零售金额, 差价, 产地, 批号, 效期, 供药单位id, 生产日期, 批准文号
    From 药品收发记录 A
    Where NO = No_In And 单据 = 3 And 记录状态 = 2
	Order By 药品id,批次;
Begin
  Update 药品收发记录 Set 记录状态 = 3 Where NO = No_In And 单据 = 3 And 记录状态 = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 审核人, 审核日期, 费用id, 扣率, 供药单位id, 生产日期, 批准文号)
    Select 药品收发记录_Id.Nextval, 2, 单据, No_In, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 效期, -填写数量, -实际数量, 成本价,
           -成本金额, 零售价, -零售金额, -差价, 摘要, 审核人_In, Sysdate, 审核人_In, Sysdate, 费用id, 扣率, 供药单位id, 生产日期, 批准文号
    From 药品收发记录
    Where NO = No_In And 单据 = 3 And 记录状态 = 3;

  For v_药品收发记录 In c_药品收发记录 Loop
    --更改药品库存表的相应数据
    Zl_药品库存_Update(v_药品收发记录.Id, 0);
  
    --处理调价后冲销
    Zl_药品收发记录_调价修正(v_药品收发记录.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_协定入库_Strike;
/

--117925:刘涛,2017-12-06,排序导致死锁处理
Create Or Replace Procedure Zl_自制入库_Delete(
                                           
                                           --删除药品收发记录及相应的表
                                           No_In In 药品收发记录.No%Type) Is
  Err_Isverified Exception;

  Cursor c_药品收发记录 Is
    Select ID From 药品收发记录 Where NO = No_In And 单据 = 2 Order By 药品id,批次;
Begin

  --通过循环，恢复所有构成原料药原来的可用数量
  For v_药品收发记录 In c_药品收发记录 Loop
    --更新库存
    Zl_药品库存_Update(v_药品收发记录.Id, 1);
  End Loop;

  Delete --删除本身及相应的构成原料药
  From 药品收发记录
  Where NO = No_In And 单据 = 2 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_自制入库_Delete;
/

--117925:刘涛,2017-12-06,排序导致死锁处理
Create Or Replace Procedure Zl_自制入库_Verify
(
  No_In     In 药品收发记录.No%Type := Null,
  审核人_In In 药品收发记录.审核人%Type := Null
) Is
  Err_Isverified Exception;
  v_差价率     Number;
  v_出库差价   药品库存.实际差价%Type;
  v_成本价     药品收发记录.成本价%Type;
  v_成本金额   药品收发记录.成本金额%Type;
  v_成本价方式 Zlparameters.参数值%Type;
  Intdigit     Number;

  Cursor c_药品收发记录 Is
    Select ID, 实际数量, Nvl(零售价, 0) As 零售价, 零售金额, 差价, 库房id, 药品id, 批次, 成本价, 批号, 效期, 产地, 入出类别id, 入出系数, 对方部门id, 供药单位id, 生产日期,
           批准文号
    From 药品收发记录
    Where NO = No_In And 单据 = 2 And 记录状态 = 1
    Order By 药品id,批次;
Begin
  Select Zl_To_Number(Nvl(Zl_Getsysparameter(9), '2')) Into Intdigit From Dual;
  Select Nvl(参数值, 0)
  Into v_成本价方式
  From Zlparameters
  Where 参数名 = '药品自制入库成本价计算方式' And 模块 = 1301;

  Update 药品收发记录
  Set 审核人 = Nvl(审核人_In, 审核人), 审核日期 = Sysdate
  Where NO = No_In And 单据 = 2 And 记录状态 = 1 And 审核人 Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  For v_药品收发记录 In c_药品收发记录 Loop
    --更改药品库存表的相应数据
    If v_药品收发记录.入出系数 = -1 Then
      v_成本价   := Zl_Fun_Getoutcost(v_药品收发记录.药品id, Nvl(v_药品收发记录.批次, 0), v_药品收发记录.库房id);
      v_成本金额 := Round(v_成本价 * v_药品收发记录.实际数量, Intdigit);
      v_出库差价 := Round(v_药品收发记录.零售金额 - v_成本金额, Intdigit);
    Else
      If v_成本价方式 = 0 Then
        Begin
          Select Sum(成本价)
          Into v_成本价
          From (Select Decode(Nvl(c.实际金额, 0), 0,
                                Decode(Nvl(c.上次采购价, 0), 0, Decode(Nvl(b.成本价, 0), 0, (d.现价 - d.现价 * (b.指导差价率 / 100)), b.成本价),
                                        c.上次采购价), (d.现价 - d.现价 * (c.实际差价 / c.实际金额))) * (a.分子 / a.分母) * (e.剂量系数 / b.构剂量系数) As 成本价
                 From 自制药品构成 A,
                      (Select b.药品id, b.成本价, b.指导差价率, b.剂量系数 As 构剂量系数
                        From 收费项目目录 A, 药品规格 B
                        Where a.Id = b.药品id And Nvl(是否变价, 0) = 0) B,
                      (Select 库房id, 药品id, 实际金额, 实际差价, 上次采购价
                        From 药品库存
                        Where 性质 = 1 And 库房id = v_药品收发记录.对方部门id) C,
                      (Select 收费细目id, 现价
                        From 收费价目
                        Where ((Sysdate Between 执行日期 And 终止日期) Or (Sysdate >= 执行日期 And 终止日期 Is Null))) D, 药品规格 E
                 Where a.原料药品id = b.药品id And b.药品id = d.收费细目id And b.药品id = c.药品id(+) And e.药品id = v_药品收发记录.药品id And
                       a.自制药品id = v_药品收发记录.药品id
                 Union All
                 Select Decode(Nvl(c.实际金额, 0), 0,
                                Decode(Nvl(c.上次采购价, 0), 0, Decode(Nvl(b.成本价, 0), 0, (c.现价 - c.现价 * (b.指导差价率 / 100)), b.成本价),
                                        c.上次采购价), (c.现价 - c.现价 * (c.实际差价 / c.实际金额))) * (a.分子 / a.分母) * (e.剂量系数 / b.构剂量系数) As 成本价
                 From 自制药品构成 A,
                      (Select b.药品id, b.成本价, b.指导差价率, b.剂量系数 As 构剂量系数
                        From 收费项目目录 A, 药品规格 B
                        Where a.Id = b.药品id And Nvl(是否变价, 0) = 1) B,
                      (Select 库房id, 药品id, 实际金额, 实际差价, 实际金额 / 实际数量 As 现价, 上次采购价
                        From 药品库存
                        Where 性质 = 1 And 库房id = v_药品收发记录.对方部门id And 实际数量 > 0) C, 药品规格 E
                 Where a.原料药品id = b.药品id And b.药品id = c.药品id And e.药品id = v_药品收发记录.药品id And a.自制药品id = v_药品收发记录.药品id);
        Exception
          When Others Then
            v_成本价 := 0;
        End;
        v_成本金额 := Round(v_成本价 * v_药品收发记录.实际数量, Intdigit);
        v_出库差价 := Round(v_药品收发记录.零售金额 - v_成本金额, Intdigit);
      Else
        v_成本价   := v_药品收发记录.成本价;
        v_成本金额 := Round(v_药品收发记录.成本价 * v_药品收发记录.实际数量, Intdigit);
        v_出库差价 := Round(v_药品收发记录.零售金额 - v_成本金额, Intdigit);
      End If;
    End If;
  
    Update 药品收发记录 Set 成本价 = v_成本价, 成本金额 = v_成本金额, 差价 = v_出库差价 Where ID = v_药品收发记录.Id;
  
    --调用库存更新过程
    Zl_药品库存_Update(v_药品收发记录.Id, 0);
  
    If v_药品收发记录.入出系数 = 1 Then
      --只有入业务才处理
      --更新该药品的成本价
      Update 药品规格 Set 成本价 = v_成本价 Where 药品id = v_药品收发记录.药品id;
    End If;
  
  End Loop;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_自制入库_Verify;
/

--117925:刘涛,2017-12-06,排序导致死锁处理
Create Or Replace Procedure Zl_自制入库_Strike
(
  No_In     In 药品收发记录.No%Type,
  审核人_In In 药品收发记录.审核人%Type
) Is
  Err_Isstriked Exception;

  v_入出类别id 药品收发记录.入出类别id%Type;

  Cursor c_药品收发记录 Is
    Select ID, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 零售金额, 差价, 供药单位id, 生产日期, 批准文号
    From 药品收发记录 A
    Where NO = No_In And 单据 = 2 And 记录状态 = 2
    Order By 药品id,批次;
Begin
  Update 药品收发记录 Set 记录状态 = 3 Where NO = No_In And 单据 = 2 And 记录状态 = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 审核人, 审核日期, 费用id, 扣率, 供药单位id, 生产日期, 批准文号)
    Select 药品收发记录_Id.Nextval, 2, 2, No_In, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 效期, -填写数量, -实际数量, 成本价,
           -成本金额, 零售价, -零售金额, -差价, 摘要, 审核人_In, Sysdate, 审核人_In, Sysdate, 费用id, 扣率, 供药单位id, 生产日期, 批准文号
    From 药品收发记录
    Where NO = No_In And 单据 = 2 And 记录状态 = 3;

  For v_药品收发记录 In c_药品收发记录 Loop
    --更改药品库存表的相应数据
    Zl_药品库存_Update(v_药品收发记录.Id, 0);
  
    --处理调价后冲销
    Zl_药品收发记录_调价修正(v_药品收发记录.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_自制入库_Strike;
/

--115026:胡俊勇,2017-12-04,病人危急值
CREATE OR REPLACE Procedure Zl_病人信息_基本信息调整_医嘱
(
  病人id_In 病人信息变动.病人id%Type,
  就诊id_In Number,
  姓名_In   病人信息.姓名%Type,
  性别_In   病人信息.性别%Type,
  年龄_In   病人信息.年龄%Type,
  场合_In   Number, --1-门诊;2-住院
  说明_Out  Out 病人信息变动.说明%Type
) As
  ------------------------------------------------------------------------------------------
  --功能:更新医嘱相关业务数据的病人基本信息
  --入参:病人id_In:病人ID
  --     就诊id_In:门诊病人为挂号ID;住院病人为主页ID;为0说明是外来或非挂号就诊的病人(就诊id_In为空时,将批量更改该病人的所有业务数据)
  --     姓名_In:需要更改的病人姓名
  --     性别_In:需要更改的病人性别
  --     年龄_In:需要更改的病人年龄
  --     场合_In:1-门诊;2-住院
  --出参:说明_Out:病人信息调整后的说明信息，用于提示操作员进行相关操作
  ------------------------------------------------------------------------------------------
  Err_Custom Exception;
  V_Error Varchar2(2000);
  N_Count Number(3);
  V_No    病人挂号记录.No%Type;
  V_Tmp   Varchar2(100);
Begin
  --外来人员，不处理
  If Nvl(就诊id_In, 0) = 0 Then
    Return;
  End If;
  --门诊取挂号单
  If Nvl(场合_In, 0) = 1 Then
    Select NO Into V_No From 病人挂号记录 Where ID = 就诊id_In;
    If V_No Is Null Then
      V_Error := '问找到该病人的挂号记录,不能更新病人基本信息.';
      Raise Err_Custom;
    End If;
    --门诊医嘱签名,则不允许修改病人基本信息
    Select Nvl(Count(1), 0)
    Into N_Count
    From 病人医嘱记录
    Where 病人id = 病人id_In And 挂号单 = V_No And 新开签名id Is Not Null And Rownum < 2;
    If N_Count <> 0 Then
      V_Error := '病人医嘱已经签名,不能更新病人基本信息.';
      Raise Err_Custom;
    End If;

    --更新病人本次就诊的医嘱中的病人基本信息
    Update 病人医嘱记录
    Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
    Where 病人id = 病人id_In And 挂号单 = V_No;
    Return;
  End If;
  --住院病人
  If Nvl(场合_In, 0) = 2 Then
    --住院医嘱签名,则不允许修改病人基本信息
    Select Nvl(Count(1), 0)
    Into N_Count
    From 病人医嘱记录
    Where 病人id = 病人id_In And 主页id = 就诊id_In And 新开签名id Is Not Null And Rownum < 2;

    If N_Count <> 0 Then
      V_Error := '该病人医嘱已经签名,不能更新病人基本信息.';
      Raise Err_Custom;
    End If;
    --住院首页存在签名的，则不允许修改病人基本信息
    Select Nvl(Count(1), 0)
    Into N_Count
    From 病案主页从表
    Where 病人id = 病人id_In And 主页id = 就诊id_In And 信息名 In ('住院医师签名', '主治医师签名', '主任医师签名', '科主任签名') And Rownum < 2;
    If N_Count <> 0 Then
      V_Error := '该病人住院首页已经签名,不能更新病人基本信息.';
      Raise Err_Custom;
    End If;
    --病案处于锁定状态，则不允许修改病人基本信息
    Select Decode(病案状态, 1, '等待审查中', 3, '正在审查中', 5, '已经审查归档', 10, '接收待审中', Null)
    Into V_Tmp
    From 病案主页
    Where 病人id = 病人id_In And 主页id = 就诊id_In;

    If Not V_Tmp Is Null Then
      V_Error := '该病人的病案' || V_Tmp || ',不能更新病人基本信息.';
      Raise Err_Custom;
    End If;
    --病案处于编目状态，则不允许修改病人基本信息
    Select Nvl(Count(1), 0)
    Into N_Count
    From 病案主页
    Where 病人id = 病人id_In And 主页id = 就诊id_In And 编目日期 Is Not Null;
    If N_Count <> 0 Then
      V_Error := '该病人的病案已经编目,不能更新病人基本信息.';
      Raise Err_Custom;
    End If;

    --已经打印了医嘱清单的提示重新打印
    Select Nvl(Count(1), 0)
    Into N_Count
    From 病人医嘱打印
    Where 病人id = 病人id_In And 主页id = 就诊id_In And Rownum < 2;

    If N_Count <> 0 Then
      If Not 说明_Out Is Null Then
        说明_Out := 说明_Out || Chr(13);
      End If;
      说明_Out := 说明_Out || '医嘱清单:已经打印需重新打印.';
    End If;

    --已经打印了首页的提示重新打印
    Select Nvl(Count(1), 0)
    Into N_Count
    From 电子病历打印
    Where 病人id = 病人id_In And 主页id = 就诊id_In And 文件id Is Null And 种类 = 9 And Rownum < 2;
    If N_Count <> 0 Then
      If Not 说明_Out Is Null Then
        说明_Out := 说明_Out || Chr(13);
      End If;
      说明_Out := 说明_Out || '病人首页:已经打印需重新打印.';
    End If;

    Update 病人医嘱记录
    Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
    Where 病人id = 病人id_In And 主页id = 就诊id_In;

    Update 输液配药记录
    Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
    Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = 病人id_In And 主页id = 就诊id_In);

	  ---更新病人危急值记录
    Update 病人危急值记录
    Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
    Where 病人id = 病人id_In And 主页id = 就诊id_In;

    Return;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人信息_基本信息调整_医嘱;
/

--115026:胡俊勇,2017-12-04,病人危急值
Create Or Replace Procedure Zl_病人危急值医嘱_Update
(
  功能_In     In Number,
  危急值id_In In 病人危急值医嘱.危急值id%Type,
  医嘱id_In   In 病人危急值医嘱.医嘱id%Type
) Is
  --功能：危急值医嘱设置关系
  --参数：功能_In-1新增对应关系，2-删除对应关系，3-医嘱作废时删除关系
  n_病人id 病人医嘱记录.病人id%Type;
  n_主页id 病人医嘱记录.主页id%Type;
  v_挂号单 病人医嘱记录.挂号单%Type;

  n_Cnt   Number;
  v_Error Varchar2(2000);
  Err_Custom Exception;
Begin
  If 功能_In = 1 Then
  
    --只能关联同一次就诊的医嘱
    Select a.病人id, a.主页id, a.挂号单
    Into n_病人id, n_主页id, v_挂号单
    From 病人医嘱记录 A, 病人危急值记录 B
    Where a.Id = b.医嘱id And b.Id = 危急值id_In;      
    If v_挂号单 Is Null Then
      Select Count(1)
      Into n_Cnt
      From 病人医嘱记录 A
      Where a.Id = 医嘱id_In And a.病人id = n_病人id And a.主页id = n_主页id;
    Else
      Select Count(1) Into n_Cnt From 病人医嘱记录 A Where a.Id = 医嘱id_In And a.挂号单 = v_挂号单;
    End If;
    If n_Cnt = 0 Then
      v_Error := '只能关联本次就诊的医嘱。';
      Raise Err_Custom;
    End If;
  
    Insert Into 病人危急值医嘱 (危急值id, 医嘱id) Values (危急值id_In, 医嘱id_In);
  Elsif 功能_In = 2 Then
    Delete 病人危急值医嘱 A Where a.危急值id = 危急值id_In And a.医嘱id = 医嘱id_In;
  Elsif 功能_In = 3 Then
    Delete 病人危急值医嘱 A Where a.医嘱id = 医嘱id_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人危急值医嘱_Update;
/

--115026:胡俊勇,2017-12-04,病人危急值
CREATE OR REPLACE Procedure Zl_病人危急值记录_处理
(
  Id_In         In 病人危急值记录.Id%Type,
  处理情况_In   In 病人危急值记录.处理情况%Type,
  确认时间_In   In 病人危急值记录.确认时间%Type,
  确认人_In     In 病人危急值记录.确认人%Type,
  确认科室id_In In 病人危急值记录.确认科室id%Type,
  是否危急值_In In 病人危急值记录.是否危急值%Type
) Is
  --病人危急值处理，处理后状态更新为2表示医生已处理
Begin
  Update 病人危急值记录
  Set 处理情况 = 处理情况_In, 确认时间 = 确认时间_In, 确认人 = 确认人_In, 确认科室id = 确认科室id_In, 是否危急值 = 是否危急值_In, 状态 = 2
  Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人危急值记录_处理;
/

--115026:胡俊勇,2017-12-04,病人危急值
CREATE OR REPLACE Procedure Zl_病人危急值记录_Insert
(
  Id_In         In 病人危急值记录.Id%Type,
  数据来源_In   In 病人危急值记录.数据来源%Type,
  病人id_In     In 病人危急值记录.病人id%Type,
  主页id_In     In 病人危急值记录.主页id%Type,
  挂号单_In     In 病人危急值记录.挂号单%Type,
  婴儿_In       In 病人危急值记录.婴儿%Type,
  姓名_In       In 病人危急值记录.姓名%Type,
  性别_In       In 病人危急值记录.性别%Type,
  年龄_In       In 病人危急值记录.年龄%Type,
  医嘱id_In     In 病人危急值记录.医嘱id%Type,
  标本id_In     In 病人危急值记录.标本id%Type,
  危急值描述_In In 病人危急值记录.危急值描述%Type,
  报告时间_In   In 病人危急值记录.报告时间%Type,
  报告科室id_In In 病人危急值记录.报告科室id%Type,
  报告人_In     In 病人危急值记录.报告人%Type
) Is
--功能：危急值登记，新增时 状态 缺省为1
Begin
  Insert Into 病人危急值记录
    (ID, 数据来源, 病人id, 主页id, 挂号单, 婴儿, 姓名, 性别, 年龄, 医嘱id, 标本id, 危急值描述, 报告时间, 报告科室id, 报告人, 状态)
  Values
    (Id_In, 数据来源_In, 病人id_In, 主页id_In, 挂号单_In, 婴儿_In, 姓名_In, 性别_In, 年龄_In, 医嘱id_In, 标本id_In, 危急值描述_In, 报告时间_In,
     报告科室id_In, 报告人_In, 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人危急值记录_Insert;
/

--115026:胡俊勇,2017-12-04,病人危急值
Create Or Replace Procedure Zl_病人危急值记录_Update
(
  Id_In         In 病人危急值记录.Id%Type,
  数据来源_In   In 病人危急值记录.数据来源%Type,
  病人id_In     In 病人危急值记录.病人id%Type,
  主页id_In     In 病人危急值记录.主页id%Type,
  挂号单_In     In 病人危急值记录.挂号单%Type,
  婴儿_In       In 病人危急值记录.婴儿%Type,
  姓名_In       In 病人危急值记录.姓名%Type,
  性别_In       In 病人危急值记录.性别%Type,
  年龄_In       In 病人危急值记录.年龄%Type,
  医嘱id_In     In 病人危急值记录.医嘱id%Type,
  标本id_In     In 病人危急值记录.标本id%Type,
  危急值描述_In In 病人危急值记录.危急值描述%Type,
  报告时间_In   In 病人危急值记录.报告时间%Type,
  报告科室id_In In 病人危急值记录.报告科室id%Type,
  报告人_In     In 病人危急值记录.报告人%Type
) Is
  --功能：修改危急值记录
  n_状态  病人危急值记录.状态%Type;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select 状态 Into n_状态 From 病人危急值记录 Where ID = Id_In;
  If Nvl(n_状态, 0) <> 1 Then
    v_Error := '当前危急值记录已被医生处理确认，不能修改。';
    Raise Err_Custom;
  End If;
  Update 病人危急值记录
  Set 数据来源 = 数据来源_In, 病人id = 病人id_In, 主页id = 主页id_In, 挂号单 = 挂号单_In, 婴儿 = 婴儿_In, 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In,
      医嘱id = 医嘱id_In, 标本id = 标本id_In, 危急值描述 = 危急值描述_In, 报告时间 = 报告时间_In, 报告科室id = 报告科室id_In, 报告人 = 报告人_In
  Where ID = Id_In;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人危急值记录_Update;
/

--115026:胡俊勇,2017-12-04,病人危急值
Create Or Replace Procedure Zl_病人危急值记录_Delete(Id_In In 病人危急值记录.Id%Type) Is
  n_状态  病人危急值记录.状态%Type;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select 状态 Into n_状态 From 病人危急值记录 Where ID = Id_In;
  If Nvl(n_状态, 0) <> 1 Then
    v_Error := '当前危急值记录已被医生处理确认，不能删除。';
    Raise Err_Custom;
  End If;
  Delete 病人危急值记录 Where ID = Id_In;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人危急值记录_Delete;
/

--117527:冉俊明,2017-12-01,预交款管理使用消费卡支付后退余额，金额未退回消费卡
Create Or Replace Procedure Zl_病人卡结算记录_退款
(
  原预交id_In 病人预交记录.Id%Type,
  预交id_In   病人预交记录.Id%Type,
  退款金额_In 病人预交记录.冲预交%Type
) Is
  --功能：退回消费卡
  --说明：仅预交款管理调用
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_存在卡片 Number(3);
  n_序号     消费卡目录.序号%Type;
  n_最大序号 消费卡目录.序号%Type;
  d_停用日期 消费卡目录.停用日期%Type;
  d_回收时间 消费卡目录.回收时间%Type;
  v_卡号     消费卡目录.卡号%Type;

  n_Id       病人卡结算记录.Id%Type;
  n_结算金额 病人卡结算记录.结算金额%Type;
  n_本次金额 病人卡结算记录.结算金额%Type;
Begin
  n_结算金额 := Nvl(退款金额_In, 0);
  For c_冲销 In (Select Distinct Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, a.消费卡id,
                               Decode(Max(a.记录状态), 1, 2, Max(a.记录状态) + 2) As 记录状态, Sum(a.结算金额) As 结算金额
               From 病人卡结算记录 A, 病人卡结算记录 B, 病人卡结算对照 C
               Where a.接口编号 = b.接口编号 And a.消费卡id = b.消费卡id And a.序号 = b.序号 And b.Id = c.卡结算id And c.预交id = 原预交id_In
               Group By a.接口编号, a.消费卡id, a.序号
               Having Nvl(Sum(a.结算金额), 0) > 0) Loop
  
    If c_冲销.结算金额 < n_结算金额 Then
      n_本次金额 := c_冲销.结算金额;
      n_结算金额 := n_结算金额 - c_冲销.结算金额;
    Else
      n_本次金额 := n_结算金额;
      n_结算金额 := 0;
    End If;
  
    --检查当前卡号是否已经使用 
    Begin
      Select 卡号, 1, 停用日期, (Select Max(序号) From 消费卡目录 B Where a.卡号 = b.卡号 And a.接口编号 = b.接口编号), 序号, 回收时间
      Into v_卡号, n_存在卡片, d_停用日期, n_最大序号, n_序号, d_回收时间
      From 消费卡目录 A
      Where ID = c_冲销.消费卡id;
    Exception
      When Others Then
        n_存在卡片 := 0;
    End;
  
    --取消停用 
    If n_存在卡片 = 0 Then
      v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡可能被他人删除，不能再退费到该卡！';
      Raise Err_Item;
    End If;
    If Nvl(n_序号, 0) < Nvl(n_最大序号, 0) Then
      v_Err_Msg := '不能再退费到历史发放卡(卡号为"' || v_卡号 || '")！';
      Raise Err_Item;
    End If;
    If Nvl(d_停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
      v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡已经被他人停用，不能再进行退费！';
      Raise Err_Item;
    End If;
    If Nvl(d_回收时间, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
      v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡已经回收，不能再进行退费！';
      Raise Err_Item;
    End If;
  
    Update 消费卡目录 Set 余额 = Nvl(余额, 0) + n_本次金额 Where ID = c_冲销.消费卡id;
  
    Select 病人卡结算记录_Id.Nextval Into n_Id From Dual;
    Insert Into 病人卡结算记录
      (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
      Select n_Id, 接口编号, 消费卡id, 序号, c_冲销.记录状态, 结算方式, -1 * n_本次金额, 卡号, 交易流水号, 交易时间, 备注, 1
      From 病人卡结算记录
      Where ID = c_冲销.Id;
  
    Insert Into 病人卡结算对照 (预交id, 卡结算id) Values (预交id_In, n_Id);
  
    Update 病人卡结算记录 Set 记录状态 = 3 Where ID = c_冲销.Id;
  
    If n_结算金额 = 0 Then
      Exit;
    End If;
  End Loop;

  If n_结算金额 > 0 Then
    v_Err_Msg := '消费卡剩余可退金额(' || LTrim(To_Char(退款金额_In - n_结算金额, '9999999990.00')) || ')不足本次退款金额(' ||
                 LTrim(To_Char(退款金额_In, '9999999990.00')) || ')，不能退费！';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人卡结算记录_退款;
/

--117279:刘硕,2017-12-01,停用人员时锁定账户
Create Or Replace Procedure Zl_人员表_启用
(
Id_In In 人员表.Id%Type
) Is
  v_User 上机人员表.用户名%Type;
Begin
  Update 人员表 Set 撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'), 撤档原因 = '' Where ID = Id_In;
  Select Max(用户名) Into v_User From 上机人员表 Where 人员id = Id_In;
  If Not v_User Is Null Then
    Begin
      --启用数据库用户
      Execute Immediate 'Alter User ' || v_User || '  Account UnLock';
    Exception
      When Others Then
        Null;
        --1、由于用户可能不存在，上级人员表为我们记录的用户，并不是数据库实际存在用户 
      --2、系统所有者没有权限，以前系统所有者没有ALter User权限
      --因此采取错误屏蔽
    End;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_人员表_启用;
/

--117279:刘硕,2017-12-01,停用人员时锁定账户
Create Or Replace Procedure Zl_人员表_Delete(Id_In In 人员表.Id%Type) Is
  v_User 上机人员表.用户名%Type;
Begin
  Select Max(用户名) Into v_User From 上机人员表 Where 人员id = Id_In;
  Delete From 人员表 Where ID = Id_In;
  If Not v_User Is Null Then
    Begin
      --停用数据库用户
      Execute Immediate 'Alter User ' || v_User || '  Account Lock';
    Exception
      When Others Then
        Null;
        --1、由于用户可能不存在，上级人员表为我们记录的用户，并不是数据库实际存在用户 
      --2、系统所有者没有权限，以前系统所有者没有ALter User权限
      --因此采取错误屏蔽
    End;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_人员表_Delete;
/

--117279:刘硕,2017-12-01,停用人员时锁定账户
Create Or Replace Procedure Zl_人员表_停用
(
  Id_In       In 人员表.Id%Type,
  撤档原因_In 人员表.撤档原因%Type := Null
) Is
  v_User 上机人员表.用户名%Type;
Begin
  Update 人员表 Set 撤档时间 = Sysdate, 撤档原因 = 撤档原因_In Where ID = Id_In;
  Select Max(用户名) Into v_User From 上机人员表 Where 人员id = Id_In;
  If Not v_User Is Null Then
    Begin
      --停用数据库用户
      Execute Immediate 'Alter User ' || v_User || '  Account Lock';
    Exception
      When Others Then
        Null;
        --1、由于用户可能不存在，上级人员表为我们记录的用户，并不是数据库实际存在用户 
      --2、系统所有者没有权限，以前系统所有者没有ALter User权限
      --因此采取错误屏蔽
    End;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_人员表_停用;
/

--107618:胡俊勇,2017-11-29,分诊诊室
Create Or Replace Procedure Zl_病人接诊_Cancel
(
  病人id_In 病人信息.病人id%Type,
  No_In     病人挂号记录.No%Type
) As
  v_门诊号   病人信息.门诊号%Type;
  v_挂号id   病人挂号记录.Id%Type;
  v_分诊方式 挂号安排.分诊方式%Type;
  n_挂号模式 Number(3);

  n_转诊       Number(1);
  n_申请科室id 病人转诊记录.申请科室id%Type;
  v_申请医生   病人转诊记录.申请医生%Type;
Begin
  n_挂号模式 := To_Number(Nvl(Substr(zl_GetSysParameter(256), 1, 1), 0));

  Select 门诊号 Into v_门诊号 From 病人信息 Where 病人id = 病人id_In;
  Select ID Into v_挂号id From 病人挂号记录 Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;

  --确定原挂号号别的医生,用于还原
  Begin
    If Nvl(n_挂号模式, 0) = 0 Then
      Select Nvl(a.分诊方式, 0) Into v_分诊方式 From 挂号安排 A, 病人挂号记录 B Where a.号码 = b.号别 And b.No = No_In;
    Else
      Select Nvl(a.分诊方式, 0)
      Into v_分诊方式
      From 临床出诊记录 A, 病人挂号记录 B
      Where a.Id = b.出诊记录id And b.No = No_In;
    End If;
  Exception
    When Others Then
      Null;
  End;

  --判断病人是否是转诊方式(强制续诊/转诊),如果是该回恢到以前的 科室和医生 然后撤消转诊变动记录
  For R In (Select a.挂号id, a.申请科室id, a.申请医生, a.接收科室id, a.接收医生, a.接收时间
            From 病人转诊记录 A
            Where a.No = No_In
            Order By a.接收时间 Desc) Loop
    n_申请科室id := r.申请科室id;
    v_申请医生   := r.申请医生;
    n_转诊       := 1;
    Delete 病人转诊记录 Where 挂号id = r.挂号id And 接收时间 = r.接收时间;
    Exit;
  End Loop;

  --就诊状态
  Update 病人信息 Set 就诊时间 = Null, 就诊状态 = 1 Where 病人id = 病人id_In;

  Update 门诊费用记录
  Set 执行状态 = 0, 执行时间 = Null, 发药窗口 = Decode(v_分诊方式, 0, Null, 发药窗口), 结论 = Null
  Where NO = No_In And 记录性质 = 4 And 记录状态 In (1, 3);

  If n_转诊 = 1 Then
    Update 病人挂号记录
    Set 执行部门id = n_申请科室id, 执行人 = v_申请医生, 执行状态 = 0, 执行时间 = Null, 诊室 = Decode(v_分诊方式, 0, Null, 诊室), 摘要 = Null
    Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
  Else
    Update 病人挂号记录
    Set 执行状态 = 0, 执行时间 = Null, 诊室 = Decode(v_分诊方式, 0, Null, 诊室), 摘要 = Null
    Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
  End If;
  
  --删除过敏，诊断信息
  Zl_病人过敏记录_Delete(病人id_In, v_挂号id);
  Zl_病人诊断记录_Delete(病人id_In, v_挂号id, Null, Null, '1,11');
  Update 排队叫号队列 Set 排队状态 = 0 Where 业务类型 = 0 And 业务id = v_挂号id;

  Delete From 病人医嘱记录 Where 病人id = 病人id_In And 挂号单 = No_In And 医嘱状态 = 1;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人接诊_Cancel;
/

--97553:胡俊勇,2017-11-29,过敏皮试结果登记时间
Create Or Replace Procedure Zl_病人医嘱记录_皮试
(
  --功能：填写医嘱皮试结果
  --说明：同时处理病人过敏记录
  --参数：标注_In=比如阳性："(+)",阴性："(-)",免试："免试"，等
  --      结果_IN=0-阴性,1-阳性，NULL=免试
  Id_In         病人医嘱记录.Id%Type,
  标注_In       病人医嘱记录.皮试结果%Type,
  结果_In       病人过敏记录.结果%Type,
  操作员姓名_In Varchar2 := Null,
  皮试时间_In   病人过敏记录.过敏时间%Type := Null,
  过敏反应_In   病人过敏药物.过敏反应%Type := Null
) Is
  --跟该过敏试验相关的所有药品信息项目
  Cursor c_Data Is
    Select Distinct c.病人id, Decode(c.挂号单, Null, c.主页id, d.Id) As 主页id, a.项目id, b.名称
    From 诊疗用法用量 A, 诊疗项目目录 B, 病人医嘱记录 C, 病人挂号记录 D
    Where Nvl(a.性质, 0) = 0 And a.用法id = c.诊疗项目id And a.项目id = b.Id And b.类别 In ('5', '6') And c.挂号单 = d.No(+) And
          c.Id = Id_In;

  v_挂号单   病人医嘱记录.挂号单%Type;
  v_状态     病人医嘱记录.医嘱状态%Type;
  v_医嘱内容 病人医嘱记录.医嘱内容%Type;

  v_Date     Date;
  d_操作时间 Date;
  v_Temp     Varchar2(255);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 病人医嘱状态.操作人员%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --检查医嘱状态是否正确:并发操作
  Select 挂号单, 医嘱状态, 医嘱内容 Into v_挂号单, v_状态, v_医嘱内容 From 病人医嘱记录 Where ID = Id_In;
  If v_状态 = 4 Then
    v_Error := '医嘱"' || v_医嘱内容 || '"已经作废，不能登记过敏试验结果。';
    Raise Err_Custom;
  End If;
  If v_挂号单 Is Not Null And v_状态 = 1 And Not 结果_In Is Null Then
    v_Error := '医嘱"' || v_医嘱内容 || '"尚未发送，不能登记过敏试验结果。';
    Raise Err_Custom;
  End If;

  --当前操作人员
  If 操作员姓名_In Is Null Then
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  Else
    v_人员姓名 := 操作员姓名_In;
    Select 编号 Into v_人员编号 From 人员表 Where 姓名 = v_人员姓名;
  End If;

  If 皮试时间_In Is Null Then
    Select Sysdate Into v_Date From Dual;
    d_操作时间 := v_Date;
  Else
    v_Date := 皮试时间_In;
    Select Sysdate Into d_操作时间 From Dual;
  End If;

  --处理医嘱记录:清除免试一样记录
  Update 病人医嘱记录 Set 皮试结果 = 标注_In, 标本部位 = To_Char(v_Date, 'YYYY-MM-DD HH24:MI:SS') Where ID = Id_In;
  Insert Into 病人医嘱状态 (医嘱id, 操作类型, 操作人员, 操作时间) Values (Id_In, 10, v_人员姓名, d_操作时间);

  --登记病人过敏记录(即使以前有同类药的过敏结果登记)
  If Not 结果_In Is Null Then
    For r_Data In c_Data Loop
      Insert Into 病人过敏记录
        (ID, 病人id, 主页id, 记录来源, 药物id, 药物名, 结果, 过敏时间, 记录时间, 记录人, 过敏反应)
      Values
        (病人过敏记录_Id.Nextval, r_Data.病人id, r_Data.主页id, 2, r_Data.项目id, r_Data.名称, 结果_In, v_Date, d_操作时间, v_人员姓名, 过敏反应_In);
      If 结果_In = 1 Then
        Update 病人过敏药物
        Set 过敏反应 = 过敏反应_In, 过敏药物id = r_Data.项目id
        Where 病人id = r_Data.病人id And 过敏药物 = r_Data.名称;
        If Sql%RowCount = 0 Then
          Insert Into 病人过敏药物
            (病人id, 过敏药物id, 过敏药物, 过敏反应)
          Values
            (r_Data.病人id, r_Data.项目id, r_Data.名称, 过敏反应_In);
        End If;
      Else
        --如果没有过敏的记录就删除该药品的过敏记录
        Delete From 病人过敏药物 A
        Where a.病人id = r_Data.病人id And a.过敏药物 = r_Data.名称 And a.过敏药物id = r_Data.项目id And Not Exists
         (Select 1
               From 病人过敏记录 B
               Where b.病人id = a.病人id And b.药物id = a.过敏药物id And b.药物名 = a.过敏药物 And 结果 = 1);
      End If;
    End Loop;
    --标记皮试结果时将医嘱自动设为执行完成
    For X In (Select 执行状态, 发送号, 执行部门id From 病人医嘱发送 Where 医嘱id = Id_In) Loop
      If x.执行状态 <> 1 Then
        Zl_病人医嘱执行_Finish(Id_In, x.发送号, Null, 0, v_人员编号, v_人员姓名, x.执行部门id);
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_皮试;
/

--116388:殷瑞,2017-11-27,修正取消发药后不能正确恢复操作状态的错误
Create Or Replace Procedure Zl_输液配药记录_取消配药(配药id_In In Varchar2 --ID串:ID1,ID2....
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
  n_Row Number(10);
  n_Out Number(1);

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_Out      := Nvl(zl_GetSysParameter('出院病人不收配置费', 1345), 0);

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
    
      If n_操作状态 != 4 Then
        v_Error := '该数据当前不是配药状态，不能进行取消配药！';
        Raise Err_Custom;
      End If;
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
  
    --向[输液配药状态]表中记录“取消配药”的操作
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
    Values
      (v_Tansid, 2, v_操作人员, Sysdate, '取消配药');
  
    Select 是否打包 Into n_打包 From 输液配药记录 Where ID = v_Tansid;
    If n_打包 <> 1 Then
      For r_Item In (Select a.No, b.序号
                     From 输液配药附费 A, 住院费用记录 B
                     Where a.病人id = b.病人id And a.No = b.No And b.记录状态 = 1 And a.配药id = v_Tansid) Loop
        If r_Item.No Is Not Null Then
          Zl_住院记帐记录_Delete(r_Item.No, r_Item.序号, v_Usercode, Zl_Username);
        End If;
      End Loop;
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

--116388:殷瑞,2017-11-27,修正取消摆药后不能正确恢复操作状态的错误
Create Or Replace Procedure Zl_输液配药记录_取消摆药(配药id_In In Varchar2 --ID串:配药ID1,配药ID2....
                                           ) Is
  v_Id       Varchar2(20);
  v_发药id   Varchar2(20);
  v_退药id   Varchar2(20);
  v_收发id   Varchar2(20);
  v_配药id   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_Date     Date;
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;
  n_操作状态 输液配药记录.操作状态%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;

  Cursor c_配药内容 Is
    Select /*+ rule*/
    Distinct c.记录id, a.Id As 退药id, c.收发id
    From 药品收发记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(配药id_In) As Zltools.t_Numlist)) D
    Where c.记录id = d.Column_Value And b.Id = c.收发id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And
          a.药品id + 0 = b.药品id And a.序号 = b.序号 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0);

  v_配药内容 c_配药内容%RowType;

  Cursor c_退药记录 Is
    Select a.Id, a.批号, a.效期, a.产地, b.数量 As 退药数
    From 药品收发记录 A, 输液配药内容 B
    Where a.Id = v_退药id And a.审核人 Is Not Null And b.收发id = v_收发id And b.记录id = v_配药id;

  v_退药记录 c_退药记录%RowType;

Begin
  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --分解单据ID串 
    v_Id  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp := Replace(',' || v_Tmp, ',' || v_Id || ',');
  
    --检查当前输液单的状态是否为待摆药状态 
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Id;
    
      If n_操作状态 != 2 Then
        v_Error := '该数据已被操作，不能进行取消摆药操作！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    Select 操作人员, 操作时间
    Into v_操作人员, d_操作时间
    From 输液配药状态
    Where 配药id = v_Id And 操作类型 = 1 And Rownum = 1;
  
    Update 输液配药记录 Set 操作状态 = 1, 操作人员 = v_操作人员, 操作时间 = d_操作时间 Where ID = v_Id;
  
    --向[输液配药状态]表中记录“取消摆药”的操作
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
    Values
      (v_Id, 1, v_操作人员, Sysdate, '取消摆药');
  
  End Loop;

  Select Sysdate Into v_Date From Dual;

  For v_配药内容 In c_配药内容 Loop
    v_退药id := v_配药内容.退药id;
    v_收发id := v_配药内容.收发id;
    v_配药id := v_配药内容.记录id;
    For v_退药记录 In c_退药记录 Loop
      --处理退药 
      Zl_药品收发记录_部门退药(v_退药记录.Id,
                     Zl_Username,
                     v_Date,
                     v_退药记录.批号,
                     v_退药记录.效期,
                     v_退药记录.产地,
                     v_退药记录.退药数,
                     Null,
                     Zl_Username);
    
      Select Max(a.Id)
      Into v_发药id
      From 药品收发记录 A, 药品收发记录 B
      Where b.Id = v_退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
            a.序号 = b.序号 And Mod(a.记录状态, 3) = 1 And a.审核日期 Is Null;
    
      --替换输液配药内容中的收发ID 
      Update 输液配药内容 Set 收发id = v_发药id Where 记录id = v_配药id And 收发id = v_配药内容.收发id;
    End Loop;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_取消摆药;
/

--116388:殷瑞,2017-11-27,修正取消发药后不能正确恢复操作状态的错误
Create Or Replace Procedure Zl_输液配药记录_取消发送(配药id_In In Varchar2 --ID串:ID1,ID2....
                                           ) Is
  v_Tansid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;
  n_打包     Number(2);

  v_Error    Varchar2(255);
  n_操作状态 输液配药记录.操作状态%Type;
  v_Usercode Varchar2(100);
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
  
    --检查当前输液单的状态是否为已发送状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;
    
      If n_操作状态 != 5 Then
        v_Error := '该数据已被操作，不能进行取消发送操作！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    Select 操作人员, 操作时间
    Into v_操作人员, d_操作时间
    From (Select 操作人员, 操作时间 From 输液配药状态 Where 配药id = v_Tansid And 操作类型 = 4 Order By 操作时间 Desc)
    Where Rownum = 1;
  
    Update 输液配药记录 Set 操作状态 = 4, 操作人员 = v_操作人员, 操作时间 = d_操作时间 Where ID = v_Tansid;
  
    --向[输液配药状态]表中记录“取消发送”的操作
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
    Values
      (v_Tansid, 4, v_操作人员, Sysdate, '取消发送');
  
    Select 是否打包 Into n_打包 From 输液配药记录 Where ID = v_Tansid;
    If n_打包 <> 0 Then
      For r_Item In (Select a.No, b.序号
                     From 输液配药附费 A, 住院费用记录 B
                     Where a.病人id = b.病人id And a.No = b.No And b.记录状态 = 1 And a.配药id = v_Tansid) Loop
        If r_Item.No Is Not Null Then
          Zl_住院记帐记录_Delete(r_Item.No, r_Item.序号, v_Usercode, Zl_Username);
        End If;
      End Loop;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_取消发送;
/

--92026:蒋廷中,2017-11-22,完善自由录入医嘱可以执行
CREATE OR REPLACE Procedure Zl_病人医嘱执行_Insert
( 
  医嘱id_In       病人医嘱执行.医嘱id%Type, 
  发送号_In       病人医嘱执行.发送号%Type, 
  要求时间_In     病人医嘱执行.要求时间%Type, 
  本次数次_In     病人医嘱执行.本次数次%Type, 
  执行摘要_In     病人医嘱执行.执行摘要%Type, 
  执行人_In       病人医嘱执行.执行人%Type, 
  执行时间_In     病人医嘱执行.执行时间%Type, 
  单独执行_In     Number := 0, 
  自动完成_In     Number := 0, 
  执行结果_In     病人医嘱执行.执行结果%Type := 1, 
  未执行原因_In   病人医嘱执行.说明%Type := Null, 
  操作员编号_In   人员表.编号%Type := Null, 
  操作员姓名_In   人员表.姓名%Type := Null, 
  执行部门id_In   门诊费用记录.执行部门id%Type := 0, 
  配液检查_In     Number := 0, 
  检验项目记帐_In Number := 0, 
  输液通道_In     病人医嘱执行.输液通道%Type := Null 
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。 
  --      执行结果_In=1- 完成   =0  -未执行 
  --      如果是台式机调用 操作员编号_In 操作员姓名_In 这两个参数必须传入 
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门 
  --配液检查_In=移动工作站调用时，是否检查配液信息。 
  --检验项目记帐_In=如果是检验项目时，需要记帐但不完成医嘱发送状态 
) Is 
  --除了要执行的主记录,还包含了附加手术,检查部位的记录 
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同 
  v_组id     病人医嘱记录.Id%Type; 
  v_诊疗类别 病人医嘱记录.诊疗类别%Type; 
  v_自动完成 Number; 
  v_病人来源 病人医嘱记录.病人来源%Type; 
  v_费用性质 病人医嘱发送.记录性质%Type; 
  v_操作类型 诊疗项目目录.操作类型%Type; 
  v_病区id   病案主页.当前病区id%Type; 
  v_配液病区 Varchar2(200); 
  v_Count    Number; 
  v_Temp     Varchar2(255); 
  v_人员编号 人员表.编号%Type; 
  v_人员姓名 人员表.姓名%Type; 
  n_期效     病人医嘱记录.医嘱期效%Type; 
  n_诊疗项目id 病人医嘱记录.诊疗项目id%Type;
  v_叮嘱执行   Varchar2(5);
 
  n_执行次数 Number; 
  n_剩余次数 Number; 
  n_执行状态 Number; 
  d_终止时间 Date; 
  d_开始时间 Date; 
  n_发送数次 Number;
  n_登记数次 Number;
  n_单次数次 Number;
  d_要求时间 Date;
 
  v_Date  Date; 
  v_Error Varchar2(255); 
  Err_Custom Exception; 
Begin 
  --并发查检，防止产生多条执行记录 
  Begin 
    Select (a.发送数次 - c.登记次数) As 剩余数次, a.发送数次, Nvl(D.诊疗项目id, 0)
    Into v_Count, n_发送数次, n_诊疗项目id
    From 病人医嘱发送 A, 
         (Select 医嘱id_In As 医嘱id, 发送号_In As 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数 
           From 病人医嘱执行 B 
           Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In) C, 病人医嘱记录 D
    Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.医嘱id = d.Id And a.发送号 = 发送号_In;
  Exception 
    When Others Then 
      v_Count := 本次数次_In; 
  End; 
  v_叮嘱执行 := zl_GetSysParameter(288);
  If 本次数次_In > v_Count And (Not (n_诊疗项目id = 0 And v_叮嘱执行 = 1)) Then
    v_Error := '由于并发操作可能已经被他人登记，请刷新后再试。'; 
    Raise Err_Custom; 
  End If; 
  --当前操作人员 
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then 
    v_人员编号 := 操作员编号_In; 
    v_人员姓名 := 操作员姓名_In; 
  Else 
    Begin 
      Select 姓名, 编号 Into v_人员姓名, v_人员编号 From 人员表 Where 姓名 = 执行人_In; 
    Exception 
      When Others Then 
        v_Temp     := Zl_Identity; 
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1); 
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1); 
        v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1); 
        v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1); 
    End; 
  End If; 
  --对医嘱终止时间进行检查 
  Select a.执行终止时间, a.开始执行时间, a.医嘱期效 
  Into d_终止时间, d_开始时间, n_期效 
  From 病人医嘱记录 A 
  Where a.Id = 医嘱id_In; 
  If Not d_终止时间 Is Null And n_期效 = 0 Then 
    If 要求时间_In > d_终止时间 Then 
      v_Error := '要求时间超过了医嘱终止时间，请确认医嘱是否提前停止！'; 
      Raise Err_Custom; 
    End If; 
  End If; 
  If Not d_开始时间 Is Null Then 
    If 执行时间_In < d_开始时间 Then 
      v_Error := '执行时间必须大于医嘱的开始执行时间''' || To_Char(d_开始时间, 'yyyy-mm-dd HH24:mi:ss') || '''！'; 
      Raise Err_Custom; 
    End If; 
  End If; 
  Select Sysdate Into v_Date From Dual; 
  Select a.病人来源, 执行科室id, Nvl(a.相关id, a.Id), Nvl(a.诊疗类别, '*'), Nvl(b.操作类型, '0') 操作类型
  Into v_病人来源, v_病区id, v_组id, v_诊疗类别, v_操作类型
  From 病人医嘱记录 A, 诊疗项目目录 B
  Where a.Id = 医嘱id_In And a.诊疗项目id = b.Id(+);

  If v_病人来源 = 2 Then 
    Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2)) 
    Into v_费用性质 
    From 病人医嘱发送 
    Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In; 
  Else 
    v_费用性质 := 1; 
  End If; 
 
  --移动系统配液检查 
  If 配液检查_In = 1 Then 
    --检查当前病人所属病区是否进行配液登记管理 
    Select Nvl(Zl_Getsysparameter(184), '') Into v_配液病区 From Dual; 
 
    If v_配液病区 Is Not Null And 执行结果_In <> 0 Then 
      If Instr(',' || v_配液病区 || ',', ',' || v_病区id || ',') > 0 Then 
        v_病区id   := 0; 
        v_配液病区 := 'Select 1 From 病区配液记录 where 医嘱ID=:YZID AND 发送号=:FSH AND 要求时间=:YQSJ'; 
        Begin 
          Execute Immediate v_配液病区 
            Into v_病区id 
            Using 医嘱id_In, 发送号_In, 要求时间_In; 
        Exception 
          When Others Then 
            Null; 
        End; 
        If v_病区id = 0 Then 
          v_Error := '当前医嘱还未进行配液，不允许进行执行登记！'; 
          Raise Err_Custom; 
        End If; 
      End If; 
    End If; 
    --检查当前医嘱是否已配液 
  End If; 
 
  --病人医嘱执行 
  Select Count(1) 
  Into v_Count 
  From 病人医嘱执行 
  Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 执行时间 = 执行时间_In; 
  If v_Count > 0 Then 
    v_Error := '您指定的执行时间，已经执行过本条医嘱，请更改一个执行时间。'; 
    Raise Err_Custom; 
  End If; 
  Insert Into 病人医嘱执行 
    (医嘱id, 发送号, 要求时间, 本次数次, 执行摘要, 执行人, 执行时间, 登记时间, 登记人, 执行结果, 说明, 输液通道) 
  Values 
    (医嘱id_In, 发送号_In, 要求时间_In, 本次数次_In, 执行摘要_In, 执行人_In, 执行时间_In, v_Date, v_人员姓名, 执行结果_In, 未执行原因_In, 输液通道_In); 
 
  --费用记录的执行状态进行更新 
  Select Decode(a.执行状态, 1, a.发送数次, c.登记次数), Decode(a.执行状态, 1, 0, a.发送数次 - c.登记次数) ,c.登记次数
  Into n_执行次数, n_剩余次数 ,n_登记数次
  From 病人医嘱发送 A, 
       (Select 医嘱id_In 医嘱id, 发送号_In 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数 
         From 病人医嘱执行 B 
         Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In And Nvl(b.执行结果, 1) <> 0) C 
  Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In; 
  --如果全部执行则状态为1，未执行状态为0，部分执行状态为2 
  Select Decode(n_剩余次数, 0, 1, Decode(n_执行次数, 0, 0, 2)) Into n_执行状态 From Dual; 
 
  --填写了执行状态后就标记为正在执行 
  If Nvl(单独执行_In, 0) = 1 Then 
    Update 病人医嘱发送 
    Set 执行状态 = Decode(n_执行次数, 0, 0, 3) 
    Where 执行状态 In (0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In; 
  Else 
    Update 病人医嘱发送 
    Set 执行状态 = Decode(n_执行次数, 0, 0, 3) 
    Where 执行状态 In (0, 3) And 发送号 + 0 = 发送号_In And 
          医嘱id In (Select ID 
                   From 病人医嘱记录 
                   Where ID = v_组id And Nvl(诊疗类别, '*') = v_诊疗类别
                   Union All
                   Select ID
                   From 病人医嘱记录
                   Where 相关id = v_组id And Nvl(诊疗类别, '*') = v_诊疗类别);
  End If; 
 
  --更新对应的费用执行状态为已执行(无正在执行) 
  --不应该处理药品和跟踪在用的卫材 
  If 执行结果_In = 1 Then 
    If v_费用性质 = 2 Then 
      If Nvl(单独执行_In, 0) = 1 Then 
        Update 住院费用记录 A 
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, Null, 执行时间_In) 
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists 
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And 
              (医嘱序号, NO, 记录性质) In 
              (Select 医嘱id, NO, 记录性质 
               From 病人医嘱发送 
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In); 
      Else 
        Update 住院费用记录 A 
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, Null, 执行时间_In) 
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists 
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And 
              (医嘱序号, NO, 记录性质) In 
              (Select 医嘱id, NO, 记录性质 
               From 病人医嘱发送 
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And 
                     医嘱id In (Select ID 
                              From 病人医嘱记录 
                              Where ID = v_组id And 诊疗类别 = v_诊疗类别 
                              Union All 
                              Select ID From 病人医嘱记录 Where 相关id = v_组id And 诊疗类别 = v_诊疗类别)); 
      End If; 
    Else 
      If Nvl(单独执行_In, 0) = 1 Then 
        --对于门诊单据n_执行状态可能为0（登记执行情况，选择执行结果为未执行），因此需判断 
        Update 门诊费用记录 A 
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, Null, 执行时间_In) 
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists 
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And 
              (医嘱序号, NO, 记录性质) In 
              (Select 医嘱id, NO, 记录性质 
               From 病人医嘱发送 
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In); 
      Else 
        Update 门诊费用记录 A 
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, Null, 执行时间_In) 
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists 
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And 
              (医嘱序号, NO, 记录性质) In 
              (Select 医嘱id, NO, 记录性质 
               From 病人医嘱发送 
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And 
                     医嘱id In (Select ID 
                              From 病人医嘱记录 
                              Where ID = v_组id And 诊疗类别 = v_诊疗类别 
                              Union All 
                              Select ID From 病人医嘱记录 Where 相关id = v_组id And 诊疗类别 = v_诊疗类别)); 
      End If; 
    End If; 
    --检验自动完成采集 
    If v_诊疗类别 = 'E' And v_操作类型 = '6' Then 
      Update 病人医嘱发送 A 
      Set a.采样人 = 执行人_In, a.采样时间 = 执行时间_In 
      Where 医嘱id In (Select ID 
                     From 病人医嘱记录 
                     Where ID = v_组id 
                     Union All 
                     Select ID From 病人医嘱记录 Where 相关id = v_组id) And 发送号 = 发送号_In; 
    End If; 
 
    --执行数次达到之后自动完成执行(主要用于PDA自动执行)，如果启用了移动临床，则护士站和PDA一致。 
    v_自动完成 := 自动完成_In; 
    If Nvl(v_自动完成, 0) = 0 And v_病人来源 = 2 And Instr('C,D', v_诊疗类别) = 0 Then 
      Begin 
        Execute Immediate 'Select Count(1) From ZLMBSYSTEMS' 
          Into v_Count; 
      Exception 
        When Others Then 
          Null; 
      End; 
      If v_Count > 0 Then 
        v_自动完成 := 1; 
      End If; 
    End If; 
 
    If Nvl(v_自动完成, 0) = 1 Or 检验项目记帐_In = 1 Then 
      Begin 
        Select Decode(Sign(Nvl(Sum(b.本次数次), 0) - a.发送数次), 1, 1, 0, 1, 0) 
        Into v_自动完成 
        From 病人医嘱发送 A, 病人医嘱执行 B 
        Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And a.执行状态 In (0, 3) And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In 
        Group By a.发送数次; 
      Exception 
        When Others Then 
          Null; 
      End; 
 
      If Nvl(v_自动完成, 0) = 1 Or 检验项目记帐_In = 1 Then 
        Zl_病人医嘱执行_Finish(医嘱id_In, 发送号_In, Null, 单独执行_In, v_人员编号, v_人员姓名, 执行部门id_In, 检验项目记帐_In); 
      End If; 
    End If; 
    --更新医嘱执行计价.执行状态
    If n_发送数次 > 0 Then
      Select Count(distinct 要求时间) Into v_Count From 医嘱执行计价 Where 医嘱ID = 医嘱ID_IN And 发送号 = 发送号_IN;
      If v_Count > 0 Then
        n_单次数次 := n_发送数次 / v_Count;
        --已执行数量+本次数次 总共能够执行多少个时间点,取最大整数
        v_Count := ceil((n_登记数次) / n_单次数次);
        --获取执行截至要求时间 
        Select 要求时间 Into d_要求时间
        From (Select 要求时间, Rownum As 次数
               From (Select Distinct 要求时间 From 医嘱执行计价 Where 医嘱id = 医嘱ID_IN And 发送号 = 发送号_IN Order By 要求时间))
        Where 次数 = v_Count;
        
        If Not d_要求时间 Is Null Then
          --先检查是否已经退费
          Select Max(NVL(执行状态,0)) Into v_Count From 医嘱执行计价 Where 医嘱id = 医嘱ID_IN And 发送号 = 发送号_IN And 要求时间 <= d_要求时间;
          If v_Count = 2 Then
            v_Error := '您指定的执行时间段的医嘱费用已经被退费，不允许再执行。'; 
            Raise Err_Custom; 
          End If;
          --更新截至要求时间之前(含)的记录执行状态；
          Update 医嘱执行计价 Set 执行状态 = 1 Where 医嘱id = 医嘱ID_IN And 发送号 = 发送号_IN And 要求时间 <= d_要求时间 And NVL(执行状态,0) <> 2;
        End If;
      End If;
    End If;
  End If; 
Exception 
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]'); 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End Zl_病人医嘱执行_Insert;
/

--114434:余伟节,2017-11-17,杭州逸曜合理用药新增处方序号
CREATE OR REPLACE Procedure Zl_病人医嘱记录_Insert
(
  Id_In           病人医嘱记录.Id%Type,
  相关id_In       病人医嘱记录.相关id%Type,
  序号_In         病人医嘱记录.序号%Type,
  病人来源_In     病人医嘱记录.病人来源%Type,
  病人id_In       病人医嘱记录.病人id%Type,
  主页id_In       病人医嘱记录.主页id%Type,
  婴儿_In         病人医嘱记录.婴儿%Type,
  医嘱状态_In     病人医嘱记录.医嘱状态%Type,
  医嘱期效_In     病人医嘱记录.医嘱期效%Type,
  诊疗类别_In     病人医嘱记录.诊疗类别%Type,
  诊疗项目id_In   病人医嘱记录.诊疗项目id%Type,
  收费细目id_In   病人医嘱记录.收费细目id%Type,
  天数_In         病人医嘱记录.天数%Type,
  单次用量_In     病人医嘱记录.单次用量%Type,
  总给予量_In     病人医嘱记录.总给予量%Type,
  医嘱内容_In     病人医嘱记录.医嘱内容%Type,
  医生嘱托_In     病人医嘱记录.医生嘱托%Type,
  标本部位_In     病人医嘱记录.标本部位%Type,
  执行频次_In     病人医嘱记录.执行频次%Type,
  频率次数_In     病人医嘱记录.频率次数%Type,
  频率间隔_In     病人医嘱记录.频率间隔%Type,
  间隔单位_In     病人医嘱记录.间隔单位%Type,
  执行时间方案_In 病人医嘱记录.执行时间方案%Type,
  计价特性_In     病人医嘱记录.计价特性%Type,
  执行科室id_In   病人医嘱记录.执行科室id%Type,
  执行性质_In     病人医嘱记录.执行性质%Type,
  紧急标志_In     病人医嘱记录.紧急标志%Type,
  开始执行时间_In 病人医嘱记录.开始执行时间%Type,
  执行终止时间_In 病人医嘱记录.执行终止时间%Type,
  病人科室id_In   病人医嘱记录.病人科室id%Type,
  开嘱科室id_In   病人医嘱记录.开嘱科室id%Type,
  开嘱医生_In     病人医嘱记录.开嘱医生%Type,
  开嘱时间_In     病人医嘱记录.开嘱时间%Type,
  挂号单_In       病人医嘱记录.挂号单%Type := Null,
  前提id_In       病人医嘱记录.前提id%Type := Null,
  检查方法_In     病人医嘱记录.检查方法%Type := Null,
  执行标记_In     病人医嘱记录.执行标记%Type := Null,
  可否分零_In     病人医嘱记录.可否分零%Type := Null,
  摘要_In         病人医嘱记录.摘要%Type := Null,
  操作员姓名_In   病人医嘱状态.操作人员%Type := Null,
  零费记帐_In     病人医嘱记录.零费记帐%Type := Null,
  用药目的_In     病人医嘱记录.用药目的%Type := Null,
  用药理由_In     病人医嘱记录.用药理由%Type := Null,
  审核状态_In     病人医嘱记录.审核状态%Type := Null,
  申请序号_In     病人医嘱记录.申请序号%Type := Null,
  超量说明_In     病人医嘱记录.超量说明%Type := Null,
  首次用量_In     病人医嘱记录.首次用量%Type := Null,
  配方id_In       病人医嘱记录.配方id%Type := Null,
  手术情况_In     病人医嘱记录.手术情况%Type := Null,
  组合项目id_In   病人医嘱记录.组合项目id%Type := Null,
  皮试结果_In     病人医嘱记录.皮试结果%Type := Null,
  处方序号_In       病人医嘱记录.处方序号%Type := Null
  --功能：医生或护士新开,补录医嘱时新产生的医嘱记录。可用于门诊或住院。
) Is
  v_Temp     Varchar2(255);
  v_人员姓名 病人医嘱状态.操作人员%Type;

  v_姓名 病人信息.姓名%Type;
  v_性别 病人信息.性别%Type;
  v_年龄 病人信息.年龄%Type;

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

  --病人医嘱记录
  Insert Into 病人医嘱记录
    (ID, 相关id, 序号, 病人来源, 病人id, 主页id, 姓名, 性别, 年龄, 婴儿, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id, 收费细目id, 天数, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 标本部位,
     检查方法, 执行标记, 执行频次, 频率次数, 频率间隔, 间隔单位, 执行时间方案, 计价特性, 执行科室id, 执行性质, 紧急标志, 可否分零, 开始执行时间, 执行终止时间, 病人科室id, 开嘱科室id, 开嘱医生,
     开嘱时间, 挂号单, 前提id, 摘要, 零费记帐, 手术时间, 用药目的, 用药理由, 审核状态, 申请序号, 超量说明, 首次用量, 配方id, 手术情况, 组合项目id, 皮试结果,处方序号)
  Values
    (Id_In, 相关id_In, 序号_In, 病人来源_In, 病人id_In, 主页id_In, v_姓名, v_性别, v_年龄, 婴儿_In, 医嘱状态_In, 医嘱期效_In, 诊疗类别_In, 诊疗项目id_In,
     收费细目id_In, 天数_In, 单次用量_In, 总给予量_In, 医嘱内容_In, 医生嘱托_In, 标本部位_In, 检查方法_In, 执行标记_In, 执行频次_In, 频率次数_In, 频率间隔_In, 间隔单位_In,
     执行时间方案_In, 计价特性_In, 执行科室id_In, 执行性质_In, 紧急标志_In, 可否分零_In, 开始执行时间_In, 执行终止时间_In, 病人科室id_In, 开嘱科室id_In, 开嘱医生_In,
     开嘱时间_In, 挂号单_In, 前提id_In, 摘要_In, 零费记帐_In,
     Decode(诊疗类别_In, 'F', To_Date(标本部位_In, 'yyyy-mm-dd hh24:mi:ss'), 'K', To_Date(标本部位_In, 'yyyy-mm-dd hh24:mi:ss'),
             Null), 用药目的_In, 用药理由_In, 审核状态_In, 申请序号_In, 超量说明_In, 首次用量_In, 配方id_In, 手术情况_In, 组合项目id_In, 皮试结果_In,处方序号_In);

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

--111635:梁唐彬,2017-11-16,自定义申请单
Create Or Replace Procedure Zl_诊疗单据目录_Edit
(
  操作_In  In Number, --1:增加;2-修改;3-删除
  Id_In    In 病历文件列表.ID%Type,
  编号_In  In 病历文件列表.编号%Type := Null,
  名称_In  In 病历文件列表.名称%Type := Null,
  说明_In  In 病历文件列表.说明%Type := Null,
  保留_In  In 病历文件列表.保留%Type := Null,
  通用_In  In 病历文件列表.通用%Type := Null,
  附项_In  In Varchar2 := Null, --按chr(10)行分隔，chr(9)字段分隔
  密码1_In In Varchar2 := Null,
  密码2_In In Varchar2 := Null,
  子类_In  In 病历文件列表.子类%Type := Null,
  格式_In  In 病历文件列表.格式%Type := Null
) Is
  v_编号 病历文件列表.编号%Type; --原编号
  n_系统 zlTools.zlReports.系统%Type;
  v_密码 zlTools.zlReports.密码%Type := 'Wait';

  --获得当前系统号
  Function f_Cur_Sys Return Number Is
    n_Sys_No zlTools.zlSystems.编号%Type;
  Begin
    Select Min(编号)
    Into n_Sys_No
    From zlTools.zlSystems
    Where 所有者 In
          (Select Owner From All_Objects Where Object_Name = Upper('Zl_诊疗单据目录_Edit') And Object_Type = 'PROCEDURE');
    Return n_Sys_No;
  End f_Cur_Sys;

  --填写申请附项
  Procedure p_Append_Items Is
    v_All    Varchar2(2000);
    v_Row    Varchar2(1000);
    v_Val    Varchar2(1000);
    v_项目   病历单据附项.项目%Type;
    n_必填   病历单据附项.必填%Type := 0;
    n_排列   病历单据附项.排列%Type := 0;
    n_要素id 病历单据附项.要素id%Type := Null;
    v_内容   病历单据附项.内容%Type;
    n_只读   病历单据附项.只读%Type := 0;
  Begin
    Delete 病历单据附项 Where 文件id = Id_In;
    
    If 附项_In Is Null Then
      Return;
    End If;
    
    v_All := 附项_In || Chr(10);
    Loop
      v_Row  := Substr(v_All, 1, Instr(v_All, Chr(10)) - 1);
      v_项目 := Substr(v_Row, 1, Instr(v_Row, Chr(9)) - 1);

      v_Val := Substr(v_Row, Instr(v_Row, Chr(9), 1, 1) + 1, Instr(v_Row, Chr(9), 1, 2) - Instr(v_Row, Chr(9), 1, 1) - 1);
      If v_Val Is Null Then
        n_必填 := 0;
      Else
        n_必填 := To_Number(v_Val);
      End If;
      
      v_Val := Substr(v_Row, Instr(v_Row, Chr(9), 1, 2) + 1, Instr(v_Row, Chr(9), 1, 3) - Instr(v_Row, Chr(9), 1, 2) - 1);
      If v_Val Is Null Then
        n_只读 := 0;
      Else
        n_只读 := To_Number(v_Val);
      End If;
            
      v_Val := Substr(v_Row, Instr(v_Row, Chr(9), 1, 3) + 1, Instr(v_Row, Chr(9), 1, 4) - Instr(v_Row, Chr(9), 1, 3) - 1);
      If v_Val Is Null Then
        n_要素id := Null;
      Else
        n_要素id := To_Number(v_Val);
      End If;
      
      v_Val  := Substr(v_Row, Instr(v_Row, Chr(9), 1, 4) + 1);
      v_内容 := v_Val;

      n_排列 := n_排列 + 1;
      Insert Into 病历单据附项
        (文件id, 项目, 必填, 排列, 要素id, 内容,只读)
      Values
        (Id_In, v_项目, n_必填, n_排列, n_要素id, v_内容,n_只读);
      v_All := Substr(v_All, Instr(v_All, Chr(10)) + 1);
      Exit When v_All Is Null;
    End Loop;
    
    delete 病历附项模板 a where a.病历文件Id=Id_In and not exists(select 1 from 病历单据附项 where 项目=a.单据附项);
  End p_Append_Items;

  --按诊疗单据报表模板添加本单据对应报表
  Procedure p_Add_Report(Form_In Number) Is
    --参数：form_In=1,申请; form_In=2,报告
    n_Mdl_Id zlTools.zlReports.ID%Type;
    n_Rpt_Id zlTools.zlReports.ID%Type;
    n_Dat_Id zlTools.zlRPTDatas.ID%Type;
    e_Mod_Lost Exception;
  Begin
    Begin
      Select ID Into n_Mdl_Id From zlReports Where 系统 = n_系统 And Upper(编号) = 'ZLEMRBILLMOLD1-' || Form_In;
    Exception
      When Others Then
        n_Mdl_Id := 0;
    End;
    If n_Mdl_Id = 0 Then
      Raise e_Mod_Lost;
    End If;
    -- 11698 产生的报表，缺密码，不能直接设计
    If Form_In = 1 Then
      v_密码 := 密码1_In;
    Elsif Form_In = 2 Then
      v_密码 := 密码2_In;
    End If;
    If v_密码 Is Null Then
      v_密码 := 'Wait...';
    End If;
    Select zlTools.Zlreports_Id.Nextval Into n_Rpt_Id From Dual;
    Insert Into zlTools.zlReports
      (ID, 编号, 名称, 说明, 密码, 进纸, 打印机, 票据, 系统, 程序id, 功能, 修改时间, 发布时间)
      Select n_Rpt_Id, 'ZLCISBILL00' || 编号_In || '-' || Form_In, 名称_In, 说明_In, v_密码, 进纸, 打印机, 票据, 系统, Null, Null,
             Sysdate, Null
      From zlTools.zlReports
      Where ID = n_Mdl_Id;
    For r_Rptdatas In (Select ID From zlTools.zlRPTDatas Where 报表id = n_Mdl_Id) Loop
      Select zlTools.Zlrptdatas_Id.Nextval Into n_Dat_Id From Dual;
      Insert Into zlTools.zlRPTDatas
        (ID, 报表id, 名称, 字段, 对象, 类型)
        Select n_Dat_Id, n_Rpt_Id, 名称, 字段, 对象, 类型 From zlTools.zlRPTDatas Where ID = r_Rptdatas.ID;
      Insert Into zlTools.zlRPTPars
        (源id, 组名, 序号, 名称, 类型, 缺省值, 格式, 值列表, 分类sql, 明细sql, 分类字段, 明细字段, 对象)
        Select n_Dat_Id, 组名, 序号, 名称, 类型, 缺省值, 格式, 值列表, 分类sql, 明细sql, 分类字段, 明细字段, 对象
        From zlTools.zlRPTPars
        Where 源id = r_Rptdatas.ID;
      Insert Into zlTools.zlRPTSQLs
        (源id, 行号, 内容)
        Select n_Dat_Id, 行号, 内容 From zlTools.zlRPTSQLs Where 源id = r_Rptdatas.ID;
    End Loop;
    Insert Into zlTools.zlRPTFMTs
      (报表id, 序号, 说明, W, H, 纸张, 纸向, 动态纸张, 图样)
      Select n_Rpt_Id, 序号, 说明, W, H, 纸张, 纸向, 动态纸张, 图样 From zlTools.zlRPTFMTs Where 报表id = n_Mdl_Id;
    For r_Rptitems In (Select ID From zlTools.zlRPTItems Where 报表id = n_Mdl_Id Order By ID) Loop
      Insert Into zlTools.zlRPTItems
        (ID, 报表id, 格式号, 名称, 类型, 上级id, 序号, 参照, 性质, 内容, 表头, X, Y, W, H, 行高, 对齐, 自调, 字体, 字号, 粗体, 斜体, 下线, 前景, 背景, 边框, 排序,
         格式, 汇总, 分栏, 网格, 系统)
        Select zlTools.Zlrptitems_Id.Nextval, n_Rpt_Id, 格式号, 名称, 类型, zlTools.Zlrptitems_Id.Nextval - (ID - 上级id), 序号, 参照,
               性质, 内容, 表头, X, Y, W, H, 行高, 对齐, 自调, 字体, 字号, 粗体, 斜体, 下线, 前景, 背景, 边框, 排序, 格式, 汇总, 分栏, 网格, 系统
        From zlTools.zlRPTItems
        Where ID = r_Rptitems.ID;
    End Loop;
    Update zlTools.zlRPTItems Set 内容 = 名称_In Where 报表id = n_Rpt_Id And 名称 = 'ZLBILLCAPTION';
  Exception
    When e_Mod_Lost Then
      Raise_Application_Error(-20101, '[ZLSOFT]诊疗单据模扳丢失，请联系系统管理员！[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Report;

  --主过程
Begin
  n_系统 := f_Cur_Sys;
  If 操作_In = 1 Then
    Insert Into 病历页面格式 (种类, 编号, 名称) Values (7, 编号_In, 名称_In);
    Insert Into 病历文件列表
      (ID, 种类, 编号, 名称, 说明, 保留, 通用, 页面,子类, 格式)
    Values
      (Id_In, 7, 编号_In, 名称_In, 说明_In, 保留_In, 通用_In, 编号_In,子类_In, 格式_In);
    p_Append_Items;
    p_Add_Report(1);
    --If 通用_In = 2 Then
    -- 11547 新增诊疗单据时，只要“执行后有报告”有效，新增的诊疗单据需要有对应的自定义报表
    p_Add_Report(2);
    --End If;

  Elsif 操作_In = 2 Then
    Select 编号 Into v_编号 From 病历文件列表 Where ID = Id_In;
    Update 病历页面格式 Set 编号 = 编号_In, 名称 = 名称_In Where 种类 = 7 And 编号 = v_编号;
    Update 病历文件列表
    Set 编号 = 编号_In, 名称 = 名称_In, 说明 = 说明_In, 通用 = 通用_In, 子类=子类_In, 格式 = 格式_In
    Where 种类 = 7 And ID = Id_In;
    p_Append_Items;

    Update zlTools.zlReports
    Set 编号 = 'ZLCISBILL00' || 编号_In || '-1', 名称 = 名称_In, 说明 = 说明_In, 密码 = 密码1_In
    Where 系统 = n_系统 And 编号 = 'ZLCISBILL00' || v_编号 || '-1';
    If Sql%RowCount = 0 Then
      p_Add_Report(1);
    End If;
    --If 通用_In <> 2 Then
    -- 11323 改变诊疗单据的格式,不删除自定义报表(2007-08-15 陈东)
    -- Delete zlTools.zlReports Where 系统 = n_系统 And 编号 = 'ZLCISBILL00' || v_编号 || '-2';
    --  Null;
    --Else
    Update zlTools.zlReports
    Set 编号 = 'ZLCISBILL00' || 编号_In || '-2', 名称 = 名称_In, 说明 = 说明_In, 密码 = 密码2_In
    Where 系统 = n_系统 And 编号 = 'ZLCISBILL00' || v_编号 || '-2';
    If Sql%RowCount = 0 Then
      p_Add_Report(2);
    End If;
    --End If;

  Elsif 操作_In = 3 Then
    Select 编号 Into v_编号 From 病历文件列表 Where ID = Id_In;
    Delete 病历文件列表 Where ID = Id_In;
    Delete 病历页面格式 Where 种类 = 7 And 编号 = v_编号;

    Delete zlTools.zlReports Where 系统 = n_系统 And 编号 Like 'ZLCISBILL00' || v_编号 || '-_';
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗单据目录_Edit;
/

--111635:梁唐彬,2017-11-16,自定义申请单
Create Or Replace Function Zl_Lob_Read
(
  Tab_In     In Number,
  Key_In     In Varchar2,
  Pos_In     In Number,
  Moved_In   In Number := 0,
  Lobtype_In In Number := 0
  --参数说明：
  --Tab_In：包含LOB的数据表
  --        0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形;
  --        5-电子病历格式;6-电子病历图形;7-病历页面格式(图形);8-电子病历附件;9-体温重叠标记;
  --        10-临床路径文件;11-临床路径图标;12-病历页面格式(页眉文件);13-病历页面格式(页脚文件);
  --        14-人员证书记录;19-部门扩展信息;20-人员扩展信息;22-医嘱报告内容;23-供应商照片;24-自定义申请单文件;25-医嘱申请单文件
  --Key_In：数据记录的关键字
  --Pos_In：从0开始不断读取，直到返回为空
  --Moved_In: 0正常记录,1读取转储后备表记录
  --LobType_IN:0-BLOb,1-CLOB
) Return Varchar2 Is
  l_Blob   Blob;
  l_Clob   Clob;
  v_Buffer Varchar2(32767);
  n_Amount Number := 2000;
  n_Offset Number := 1;
  t_Key    t_Strlist;
Begin
  If Tab_In = 0 Then
    Select 图形 Into l_Blob From 病历标记图形 Where 编码 = Key_In;
  Elsif Tab_In = 1 Then
    Select 内容 Into l_Blob From 病历文件格式 Where 文件id = To_Number(Key_In);
  Elsif Tab_In = 2 Then
    Select 图形 Into l_Blob From 病历文件图形 Where 对象id = To_Number(Key_In);
  Elsif Tab_In = 3 Then
    Select 内容 Into l_Blob From 病历范文格式 Where 文件id = To_Number(Key_In);
  Elsif Tab_In = 4 Then
    Select 图形 Into l_Blob From 病历范文图形 Where 对象id = To_Number(Key_In);
  Elsif Tab_In = 5 Then
    If Moved_In = 0 Then
      Select 内容 Into l_Blob From 电子病历格式 Where 文件id = To_Number(Key_In);
    Else
      Select 内容 Into l_Blob From H电子病历格式 Where 文件id = To_Number(Key_In);
    End If;
  Elsif Tab_In = 6 Then
    If Moved_In = 0 Then
      Select 图形 Into l_Blob From 电子病历图形 Where 对象id = To_Number(Key_In);
    Else
      Select 图形 Into l_Blob From H电子病历图形 Where 对象id = To_Number(Key_In);
    End If;
  Elsif Tab_In = 7 Then
    Select 图形
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
  Elsif Tab_In = 8 Then
    If Moved_In = 0 Then
      Select 内容
      Into l_Blob
      From 电子病历附件
      Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1);
    Else
      Select 内容
      Into l_Blob
      From H电子病历附件
      Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
  Elsif Tab_In = 9 Then
    Select 标记图形 Into l_Blob From 体温重叠标记 Where 序号 = To_Number(Key_In);
  Elsif Tab_In = 10 Then
    Select 内容
    Into l_Blob
    From 临床路径文件
    Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 文件名 = Substr(Key_In, Instr(Key_In, ',') + 1);
  Elsif Tab_In = 11 Then
    Select 图标 Into l_Blob From 临床路径图标 Where ID = To_Number(Key_In);
  Elsif Tab_In = 12 Then
    Select 页眉文件
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
  Elsif Tab_In = 13 Then
    Select 页脚文件
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
  Elsif Tab_In = 14 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    Select 签章信息 Into l_Clob From 人员证书记录 Where 人员id = To_Number(t_Key(1)) And Certsn = t_Key(2);
  Elsif Tab_In = 19 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    Select 图片 Into l_Blob From 部门扩展信息 Where 部门id = To_Number(t_Key(1)) And 项目 = t_Key(2);
  Elsif Tab_In = 20 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    Select 图片 Into l_Blob From 人员扩展信息 Where 人员id = To_Number(t_Key(1)) And 项目 = t_Key(2);
  Elsif Tab_In = 22 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    Select 内容 Into l_Blob From 医嘱报告内容 Where ID = To_Number(Key_In);
  Elsif Tab_In = 23 Then
    If To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=0 Then
       Select 许可证号照片 Into l_Blob From 供应商照片 Where 供应商ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
    Elsif  To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=1 Then
       Select 执照号照片 Into l_Blob From 供应商照片 Where 供应商ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
    Elsif To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=2 Then
       Select 授权号照片 Into l_Blob From 供应商照片 Where 供应商ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
    End If;
  Elsif Tab_In = 24 Then
    Select 内容
    Into l_Clob
    From 自定义申请单文件
    Where 文件id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
          类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) ;
  Elsif Tab_In = 25 Then
    Select 内容
    Into l_Clob
    From 医嘱申请单文件
    Where 医嘱id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
          类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) ;
  End If;

  n_Offset := n_Offset + Pos_In * n_Amount;
  If Lobtype_In = 1 Then
    If l_Clob Is Null Then
      v_Buffer := Null;
    Else
      Dbms_Lob.Read(l_Clob, n_Amount, n_Offset, v_Buffer);
    End If;
  Else
    If l_Blob Is Null Then
      v_Buffer := Null;
    Else
      Dbms_Lob.Read(l_Blob, n_Amount, n_Offset, v_Buffer);
    End If;
  End If;
  Return v_Buffer;
Exception
  When No_Data_Found Then
    Return Null;
End Zl_Lob_Read;
/

--111635:梁唐彬,2017-11-16,自定义申请单
Create Or Replace Procedure Zl_Lob_Append
(
  Tab_In     In Number,
  Key_In     In Varchar2,
  Txt_In     In Varchar2, --16进制的文件片段或文字片段
  Cls_In     In Number := 0, --是否清除原来的内容，第一片段传递时为1，以后为0
  Lobtype_In In Number := 0 --0-BLOB;1-CLOB
  --参数说明：
  --Tab_In：包含LOB的数据表
  --        0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形;
  --        5-电子病历格式;6-电子病历图形;7-病历页面格式；8-电子病历附件;9-体温重叠标记
  --        10-临床路径文件,11-临床路径图标;14-人员证书记录;
  --        19-部门扩展信息;20-人员扩展信息;22-医嘱报告内容;23-供应商照片;24-自定义申请单文件;25-医嘱申请单文件
  --Key_In：数据记录的关键字
  --Txt_In：16进制的文件片段或文字片段
  --Cls_In：是否清除原来的内容，第一片段传递时为1，以后为0
  --Lobtype_In:--0-BLOB;1-CLOB
) Is
  l_Blob Blob;
  l_Clob Clob;
  t_Key  t_Strlist;

Begin
  If Tab_In = 0 Then
    If Cls_In = 1 Then
      Update 病历标记图形 Set 图形 = Empty_Blob() Where 编码 = Key_In;
    End If;
    Select 图形 Into l_Blob From 病历标记图形 Where 编码 = Key_In For Update;
  Elsif Tab_In = 1 Then
    If Cls_In = 1 Then
      Update 病历文件格式 Set 内容 = Empty_Blob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历文件格式 (文件id, 内容) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 内容 Into l_Blob From 病历文件格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 2 Then
    If Cls_In = 1 Then
      Update 病历文件图形 Set 图形 = Empty_Blob() Where 对象id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历文件图形 (对象id, 图形) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 图形 Into l_Blob From 病历文件图形 Where 对象id = To_Number(Key_In) For Update;
  Elsif Tab_In = 3 Then
    If Cls_In = 1 Then
      Update 病历范文格式 Set 内容 = Empty_Blob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历范文格式 (文件id, 内容) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 内容 Into l_Blob From 病历范文格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 4 Then
    If Cls_In = 1 Then
      Update 病历范文图形 Set 图形 = Empty_Blob() Where 对象id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历范文图形 (对象id, 图形) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 图形 Into l_Blob From 病历范文图形 Where 对象id = To_Number(Key_In) For Update;
  Elsif Tab_In = 5 Then
    If Cls_In = 1 Then
      Update 电子病历格式 Set 内容 = Empty_Blob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 电子病历格式 (文件id, 内容) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 内容 Into l_Blob From 电子病历格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 6 Then
    If Cls_In = 1 Then
      Update 电子病历图形 Set 图形 = Empty_Blob() Where 对象id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 电子病历图形 (对象id, 图形) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 图形 Into l_Blob From 电子病历图形 Where 对象id = To_Number(Key_In) For Update;
  Elsif Tab_In = 7 Then
    If Cls_In = 1 Then
      Update 病历页面格式
      Set 图形 = Empty_Blob()
      Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
    End If;
    Select 图形
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 8 Then
    If Cls_In = 1 Then
      Update 电子病历附件
      Set 内容 = Empty_Blob()
      Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select 内容
    Into l_Blob
    From 电子病历附件
    Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 9 Then
    If Cls_In = 1 Then
      Update 体温重叠标记 Set 标记图形 = Empty_Blob() Where 序号 = To_Number(Key_In);
    End If;
    Select 标记图形 Into l_Blob From 体温重叠标记 Where 序号 = To_Number(Key_In) For Update;
  Elsif Tab_In = 10 Then
    If Cls_In = 1 Then
      Update 临床路径文件
      Set 内容 = Empty_Blob()
      Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            文件名 = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select 内容
    Into l_Blob
    From 临床路径文件
    Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 文件名 = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 11 Then
    If Cls_In = 1 Then
      Update 临床路径图标 Set 图标 = Empty_Blob() Where ID = To_Number(Key_In);
    End If;
    Select 图标 Into l_Blob From 临床路径图标 Where ID = To_Number(Key_In) For Update;
  Elsif Tab_In = 12 Then
    If Cls_In = 1 Then
      Update 病历页面格式
      Set 页眉文件 = Empty_Blob()
      Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
    End If;
    Select 页眉文件
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 13 Then
    If Cls_In = 1 Then
      Update 病历页面格式
      Set 页脚文件 = Empty_Blob()
      Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
    End If;
    Select 页脚文件
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 14 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update 人员证书记录 Set 签章信息 = Empty_Clob() Where 人员id = To_Number(t_Key(1)) And Certsn = t_Key(2);
    End If;
    Select 签章信息 Into l_Clob From 人员证书记录 Where 人员id = To_Number(t_Key(1)) And Certsn = t_Key(2) For Update;
  Elsif Tab_In = 19 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update 部门扩展信息 Set 图片 = Empty_Blob() Where 部门id = To_Number(t_Key(1)) And 项目 = t_Key(2);
    End If;
    Select 图片 Into l_Blob From 部门扩展信息 Where 部门id = To_Number(t_Key(1)) And 项目 = t_Key(2) For Update;
    Update 部门表 Set 最后修改时间 = Sysdate Where ID = To_Number(t_Key(1));
  Elsif Tab_In = 20 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update 人员扩展信息 Set 图片 = Empty_Blob() Where 人员id = To_Number(t_Key(1)) And 项目 = t_Key(2);
    End If;
    Select 图片 Into l_Blob From 人员扩展信息 Where 人员id = To_Number(t_Key(1)) And 项目 = t_Key(2) For Update;
    Update 人员表 Set 最后修改时间 = Sysdate Where ID = To_Number(t_Key(1));
  Elsif Tab_In = 22 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update 医嘱报告内容 Set 内容 = Empty_Blob() Where ID = To_Number(t_Key(1));
    End If;
    Select 内容 Into l_Blob From 医嘱报告内容 Where ID = To_Number(t_Key(1)) For Update;
  Elsif Tab_In = 23 Then
    If To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=0 Then
      If Cls_In = 1 Then
        Update 供应商照片 Set 许可证号照片 = Empty_Blob() Where 供应商ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into 供应商照片 (供应商ID, 许可证号照片,执照号照片,授权号照片) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select 许可证号照片 Into l_Blob From 供应商照片 Where 供应商ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    Elsif  To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=1 Then
      If Cls_In = 1 Then
        Update 供应商照片 Set 执照号照片 = Empty_Blob() Where 供应商ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into 供应商照片 (供应商ID, 许可证号照片,执照号照片,授权号照片) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select 执照号照片 Into l_Blob From 供应商照片 Where 供应商ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    Elsif To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=2 Then
     If Cls_In = 1 Then
        Update 供应商照片 Set 授权号照片 = Empty_Blob() Where 供应商ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into 供应商照片 (供应商ID, 许可证号照片,执照号照片,授权号照片) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select 授权号照片 Into l_Blob From 供应商照片 Where 供应商ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    End If;
  Elsif Tab_In = 24 Then
    If Cls_In = 1 Then
      Update 自定义申请单文件
      Set 内容 = Empty_Clob()
      Where 文件id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) ;
    End If;
    Select 内容
    Into l_Clob
    From 自定义申请单文件
    Where 文件id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) 
    For Update;
  ElsIf Tab_In = 25 Then
    If Cls_In = 1 Then
      Update 医嘱申请单文件
      Set 内容 = Empty_Clob()
      Where 医嘱id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 
            类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))  ;
    End If;
    Select 内容
    Into l_Clob
    From 医嘱申请单文件
    Where 医嘱id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) 
    For Update;
  End If;

  If Lobtype_In = 1 Then
    Dbms_Lob.Writeappend(l_Clob, Length(Txt_In), Txt_In);
  Else
    Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Lob_Append;
/

--115026:胡俊勇,2017-12-04,病人危急值
--111635:梁唐彬,2017-11-16,自定义申请单
CREATE OR REPLACE Procedure Zl_Retu_Clinic
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
              From Table(f_Str2list('电子病历附件,电子病历格式,电子病历内容,疾病申报记录,疾病报告反馈,疾病申报反馈'))) Loop
      v_Table := r.Column_Value;
      If v_Table = '电子病历附件' Then
        v_Field := '病历id';
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
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
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
      Elsif v_Table = '病人医嘱报告' Then
        v_Fields := Getfields('医嘱报告内容');
        v_Sql    := 'Insert Into 医嘱报告内容(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H医嘱报告内容 Where ID In (Select 报告id From H病人医嘱报告 Where 医嘱id = :1 And 报告id Is Not Null)';
        Execute Immediate v_Sql
          Using n_Rec_Id;

        Delete H医嘱报告内容
        Where ID In (Select 报告id From H病人医嘱报告 Where 医嘱id = n_Rec_Id And 报告id Is Not Null);

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

--111635:梁唐彬,2017-07-14,XML自定义申请单
Create Or Replace Procedure Zl_自定义申请单文件_Edit
(
  模式_In   Number, --1-新增/修改;2-删除
  文件id_In 自定义申请单文件.文件id%Type,
  类别_In   自定义申请单文件.类别%Type,
  文件名_In 自定义申请单文件.文件名%Type := Null
) As
  v_Temp     Varchar(500);
  v_人员姓名 Varchar(100);
Begin
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  If 模式_In = 1 Then
    Update 自定义申请单文件
    Set 文件名 = 文件名_In, 创建人 = v_人员姓名, 创建时间 = Sysdate
    Where 文件id = 文件id_In And 类别 = 类别_In;
    If Sql%RowCount = 0 Then
      Insert Into 自定义申请单文件
        (文件id, 文件名, 类别, 创建人, 创建时间)
      Values
        (文件id_In, 文件名_In, 类别_In, v_人员姓名, Sysdate);
    End If;
  Elsif 模式_In = 2 Then
    Delete From 自定义申请单文件 Where 文件id = 文件id_In And 类别 = 类别_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_自定义申请单文件_Edit;
/

--111635:梁唐彬,2017-07-14,XML自定义申请单
CREATE OR REPLACE Procedure Zl_医嘱申请单文件_Edit
(
  文件id_In 医嘱申请单文件.文件id%Type,
  文件名_IN 医嘱申请单文件.文件名%Type,
  类别_In   医嘱申请单文件.类别%Type,
  医嘱ID_In   医嘱申请单文件.医嘱ID%Type 
) As

Begin
  Insert Into 医嘱申请单文件
      (文件id,文件名, 医嘱ID, 类别)
  Values
      (文件id_In,文件名_IN, 医嘱ID_In, 类别_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_医嘱申请单文件_Edit;
/

--115519:李小东,2017-11-16,拒收后重采标本允许医嘱回退和作废
Create Or Replace Procedure Zl_检验预置条码_采集完成
(
  医嘱内容_In Varchar2, --内容包括多个医嘱ID使用","分隔 
  人员编号_In 人员表.编号%Type := Null,
  人员姓名_In 人员表.姓名%Type := Null, --Null=取消，不为空时完成采集
  操作_In     Number := 0, --0=完成采集，1=取消采集
  医嘱类别_In Number := 0 --0=检验医嘱,1=输血医嘱 
) Is
  n_自动发料 Number;
  --查找当前标本的相关申请 
  Cursor c_Samplequest(v_医嘱id In Varchar2) Is
    Select /*+ rule */
    Distinct ID As 医嘱id, 病人来源
    From 病人医嘱记录 A, 病人医嘱发送 B
    Where a.Id = b.医嘱id And b.接收人 Is Null And Sign(Nvl(a.相关id, 0)) = 医嘱类别_In And
          a.Id In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist)));

  --未审核的费用行(不包含药品) 
  Cursor c_Verify(v_医嘱id In Varchar2) Is
    Select /*+ rule */
    Distinct 记录性质, NO, 序号, 记录状态, 门诊标志
    From 住院费用记录
    Where 收费类别 Not In ('5', '6', '7') And
          医嘱序号 + 0 In (Select ID
                       From 病人医嘱记录
                       Where ID In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist))) And
                             Sign(Nvl(相关id, 0)) = 医嘱类别_In) And 记帐费用 = 1 And 记录状态 = 0 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist)))
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID
                                        From 病人医嘱记录
                                        Where ID In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist))) And
                                              Sign(Nvl(相关id, 0)) = 医嘱类别_In) And 接收人 Is Null)
    Union All
    Select /*+ rule */
    Distinct 记录性质, NO, 序号, 记录状态, 门诊标志
    From 门诊费用记录
    Where 收费类别 Not In ('5', '6', '7') And
          医嘱序号 + 0 In (Select ID
                       From 病人医嘱记录
                       Where ID In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist))) And
                             Sign(Nvl(相关id, 0)) = 医嘱类别_In) And 记帐费用 = 1 And 记录状态 = 0 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist)))
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID
                                        From 病人医嘱记录
                                        Where ID In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist))) And
                                              Sign(Nvl(相关id, 0)) = 医嘱类别_In) And 接收人 Is Null)
    Order By 记录性质, NO, 序号;

  v_检验标本记录 Number(18);
  v_执行状态     Number(1);
  v_接收人       Varchar2(50);
  v_Error        Varchar2(100);
  v_No           病人医嘱发送.No%Type;
  v_性质         病人医嘱发送.记录性质%Type;
  v_序号         Varchar2(1000);

  v_收发ids Varchar2(4000);
  n_库房id  Number;
  n_发料号  Number;

  Err_Custom Exception;
  n_Par Number;
Begin
  Select zl_GetSysParameter('自动发料退料', 1211) Into n_自动发料 From Dual;
  If 人员姓名_In Is Not Null And 操作_In = 0 Then
    --检查标本是否被核收或接收 
    Begin
      Select /*+ rule */
       Nvl(c.Id, 0), b.执行状态, b.接收人
      Into v_检验标本记录, v_执行状态, v_接收人
      From 病人医嘱记录 A, 病人医嘱发送 B, 检验标本记录 C
      Where a.Id = b.医嘱id And a.相关id = c.医嘱id(+) And
            a.Id In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist)));
    Exception
      When Others Then
        v_检验标本记录 := 0;
    End;
  
    If v_检验标本记录 <> 0 Then
      v_Error := '标本已被检验科核收不能完成采集!';
      Raise Err_Custom;
    End If;
  
    If v_执行状态 <> 2 And v_接收人 Is Not Null Then
      v_Error := '标本已被检验科签收不能完成采集!';
      Raise Err_Custom;
    End If;
  
    --检查医嘱是否收费
    n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
    If n_Par = 1 Then
      For r_Verify In c_Verify(医嘱内容_In) Loop
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
    End If;
  
    Update /*+ rule */ 检验拒收记录
    Set 重采人 = 人员姓名_In, 重采时间 = Sysdate
    Where 医嘱id In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist)));
  
    --更新采集信息(检验和采集） 
    Update /*+ rule */ 病人医嘱发送
    Set 采样人 = 人员姓名_In, 采样时间 = Sysdate, 执行状态 = Decode(执行状态, 2, 0, 执行状态),
        重采标本 = Decode(Nvl(重采标本, 0), 0, Decode(执行状态, 2, 1, 0), 重采标本), 执行说明 = Null
    Where 医嘱id In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist)));
  
    --更新医嘱和费用记录 
    For r_Samplequest In c_Samplequest(医嘱内容_In) Loop
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
               Where 医嘱id In (Select ID
                              From 病人医嘱记录 A, 病人医嘱发送 B
                              Where a.Id = b.医嘱id And r_Samplequest.医嘱id In (a.Id) And Sign(Nvl(a.相关id, 0)) = 医嘱类别_In And
                                    b.执行状态 In (0, 2) And b.接收人 Is Null));
      Else
        --发料
        If n_自动发料 = 1 Then
          For c_Stuff In (Select a.记录性质, a.记录状态, b.Id, b.库房id
                          From 门诊费用记录 A, 药品收发记录 B
                          Where a.Id = b.费用id And a.收费类别 = '4' And b.审核人 Is Null And
                                (a.医嘱序号, a.记录性质, a.No) In
                                (Select 医嘱id, 记录性质, NO
                                 From 病人医嘱附费
                                 Where 医嘱id = r_Samplequest.医嘱id
                                 Union All
                                 Select 医嘱id, 记录性质, NO
                                 From 病人医嘱发送
                                 Where 医嘱id In
                                       (Select ID
                                        From 病人医嘱记录 A, 病人医嘱发送 B
                                        Where a.Id = b.医嘱id And r_Samplequest.医嘱id In (a.Id) And
                                              Sign(Nvl(a.相关id, 0)) = 医嘱类别_In And b.执行状态 In (0, 2) And b.接收人 Is Null))) Loop
            If Mod(Nvl(c_Stuff.记录性质, 0), 10) = 1 And Nvl(c_Stuff.记录状态, 0) = 1 Then
              If n_发料号 Is Null Then
                n_发料号 := Nextno(20);
              End If;
            
              If c_Stuff.库房id <> Nvl(n_库房id, 0) Then
                If Nvl(n_库房id, 0) <> 0 And v_收发ids Is Not Null Then
                  v_收发ids := Substr(v_收发ids, 2);
                  Zl_药品收发记录_批量发料(v_收发ids, n_库房id, 人员姓名_In, Sysdate, 1, 人员姓名_In, n_发料号, 人员姓名_In);
                End If;
              
                n_库房id  := c_Stuff.库房id;
                v_收发ids := Null;
              End If;
            
              v_收发ids := v_收发ids || '|' || c_Stuff.Id || ',0';
            End If;
          End Loop;
          If Nvl(n_库房id, 0) <> 0 And v_收发ids Is Not Null Then
            v_收发ids := Substr(v_收发ids, 2);
            Zl_药品收发记录_批量发料(v_收发ids, n_库房id, 人员姓名_In, Sysdate, 1, 人员姓名_In, n_发料号, 人员姓名_In);
          End If;
        End If;
      
        --2.费用执行处理 
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
               Where 医嘱id In (Select ID
                              From 病人医嘱记录 A, 病人医嘱发送 B
                              Where a.Id = b.医嘱id And r_Samplequest.医嘱id In (a.Id) And Sign(Nvl(a.相关id, 0)) = 医嘱类别_In And
                                    b.执行状态 In (0, 2) And b.接收人 Is Null));
      End If;
    End Loop;
  
    --更新执行状态(只更新采集） 
    Update /*+ rule */ 病人医嘱发送
    Set 执行状态 = 1, 完成人 = 人员姓名_In, 完成时间 = Sysdate
    Where 医嘱id In (Select ID
                   From 病人医嘱记录
                   Where ID In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist))) And
                         Sign(Nvl(相关id, 0)) = 医嘱类别_In);
    --3.自动审核记帐 
    For r_Verify In c_Verify(医嘱内容_In) Loop
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
    End If;
  
  Else
    --检查标本是否被核收或接收 
    Begin
      Select /*+ rule */
       Nvl(c.Id, 0), b.执行状态, b.接收人
      Into v_检验标本记录, v_执行状态, v_接收人
      From 病人医嘱记录 A, 病人医嘱发送 B, 检验标本记录 C
      Where a.Id = b.医嘱id And a.相关id = c.医嘱id(+) And
            a.Id In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist)));
    Exception
      When Others Then
        v_检验标本记录 := 0;
    End;
  
    If v_检验标本记录 <> 0 Then
      v_Error := '标本已被检验科核收不能取消完成采集!';
      Raise Err_Custom;
    End If;
  
    If v_执行状态 <> 2 And v_接收人 Is Not Null Then
      v_Error := '标本已被检验科签收不能取消完成采集!';
      Raise Err_Custom;
    End If;
  
    Update /*+ rule */ 病人医嘱发送
    Set 采样人 = Null, 采样时间 = Null, 执行状态 = 0, 执行说明 = Null, 完成人 = Null, 完成时间 = Null
    Where 医嘱id In (Select ID
                   From 病人医嘱记录
                   Where ID In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist))));
  
    For r_Samplequest In c_Samplequest(医嘱内容_In) Loop
    
      If r_Samplequest.病人来源 = 2 Then
        --2.费用执行处理 
        Update 住院费用记录
        Set 执行状态 = 0, 执行时间 = Null, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID
                              From 病人医嘱记录
                              Where r_Samplequest.医嘱id In (ID) And Sign(Nvl(相关id, 0)) = 医嘱类别_In) And 执行状态 In (0, 2) And
                     接收人 Is Null);
      Else
        --退料
        If n_自动发料 = 1 Then
          For c_Stuff In (Select b.Id, b.实际数量
                          From 门诊费用记录 A, 药品收发记录 B
                          Where a.Id = b.费用id And a.收费类别 = '4' And (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And
                                b.审核人 Is Not Null And (a.医嘱序号, a.记录性质, a.No) In
                                (Select 医嘱id, 记录性质, NO
                                                       From 病人医嘱附费
                                                       Where 医嘱id = r_Samplequest.医嘱id
                                                       Union All
                                                       Select 医嘱id, 记录性质, NO
                                                       From 病人医嘱发送
                                                       Where 医嘱id In (Select ID
                                                                      From 病人医嘱记录
                                                                      Where r_Samplequest.医嘱id In (ID) And
                                                                            Sign(Nvl(相关id, 0)) = 医嘱类别_In) And
                                                             执行状态 In (0, 2) And 接收人 Is Null)) Loop
          
            Zl_材料收发记录_部门退料(c_Stuff.Id, 人员姓名_In, Sysdate, Null, Null, Null, c_Stuff.实际数量);
          End Loop;
        End If;
        --退费
        Update 门诊费用记录
        Set 执行状态 = 0, 执行时间 = Null, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID
                              From 病人医嘱记录
                              Where r_Samplequest.医嘱id In (ID) And Sign(Nvl(相关id, 0)) = 医嘱类别_In) And 执行状态 In (0, 2) And
                     接收人 Is Null);
      End If;
    End Loop;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验预置条码_采集完成;
/

--116673:刘兴洪,2017-11-16,将体检费用纳入合约单位结帐
Create Or Replace Procedure Zl_结帐费用记录_Unit
(
  Patientids_In Varchar2,
  结帐id_In     门诊费用记录.结帐id%Type,
  零费用结帐_In Number --初次结帐时是否排开记帐再销帐的费用,后续结帐时不能排开
) As
  Cursor c_Fee(v_病人id 门诊费用记录.病人id%Type) Is
    Select A.ID, A.NO, A.序号, A.记录性质, A.记录状态, A.执行状态, Nvl(A.实收金额, 0) As 未结金额
    From 门诊费用记录 A
    Where A.病人id = v_病人id And A.结帐id Is Null And A.记录状态 <> 0 And A.记帐费用 = 1 And A.门诊标志 In (1, 4) And
          Not Exists
     (Select 1
           From 门诊费用记录 B
           Where B.NO = A.NO And B.记录性质 = A.记录性质 And B.序号 = A.序号
           Group By B.NO, B.记录性质, B.序号
           Having Nvl(Sum(B.实收金额), 0) = Decode(零费用结帐_In, 1, 1 + Nvl(Sum(B.实收金额), 0), 0))
    Union All
    Select 0 As ID, A.NO, A.序号, Mod(A.记录性质, 10) As 记录性质, A.记录状态, A.执行状态,
           Nvl(Sum(A.实收金额), 0) - Nvl(Sum(A.结帐金额), 0) As 未结金额
    From (Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志,
                  记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数,
                  发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人,
                  开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号,
                  操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要,
                  是否急诊
           From 门诊费用记录
           Union All
           Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志,
                  记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数,
                  发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人,
                  开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号,
                  操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要,
                  是否急诊
           From H门诊费用记录) A
    Where A.病人id = v_病人id And A.结帐id Is Not Null And A.记帐费用 = 1 And A.门诊标志 in( 1,4) And
          Nvl(A.实收金额, 0) <> Nvl(A.结帐金额, 0)
    Group By A.NO, A.序号, Mod(A.记录性质, 10), A.记录状态, A.执行状态
    Having Nvl(Sum(A.实收金额), 0) - Nvl(Sum(A.结帐金额), 0) <> 0;

  v_Patientids  Varchar2(4000);
  v_Patientid   Varchar2(4000);
  v_Banlanceids Varchar2(4000);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin

  v_Patientids := Patientids_In || ',';
  While v_Patientids Is Not Null Loop
    v_Patientid   := Substr(v_Patientids, 1, Instr(v_Patientids, ',') - 1);
    v_Patientids  := Substr(v_Patientids, Instr(v_Patientids, ',') + 1);
    v_Banlanceids := '';

    For r_Fee In c_Fee(v_Patientid) Loop
      If r_Fee.ID = 0 Then
        Zl_结帐费用记录_Insert(r_Fee.ID, r_Fee.NO, r_Fee.记录性质, r_Fee.记录状态, r_Fee.执行状态, r_Fee.序号,
                               r_Fee.未结金额, 结帐id_In);
      Else
        v_Banlanceids := v_Banlanceids || ',' || r_Fee.ID;
        If Length(v_Banlanceids) > 3980 Then
          Zl_结帐费用记录_Batch(Substr(v_Banlanceids, 2), v_Patientid, 结帐id_In);
          v_Banlanceids := '';
        End If;
      End If;
    End Loop;

    If Not v_Banlanceids Is Null Then
      Zl_结帐费用记录_Batch(Substr(v_Banlanceids, 2), v_Patientid, 结帐id_In);
    End If;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_结帐费用记录_Unit;
/

--115026:胡俊勇,2017-12-04,病人危急值
--111635:梁唐彬,2017-11-16,自定义申请单
--116697:张永康,2017-11-18,预交款退款和就诊卡记帐零费用的历史数据转出特殊处理
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
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where 待转出 Is Null And 结帐id Is Null And
        (病人id, 主页id) In (Select 病人id, 主页id
                         From 病案主页 C
                         Where 出院日期 < d_End And 待转出 Is Null And 数据转出 Is Null And Not Exists
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
  Set 待转出 = n_批次
  Where ID In (Select c.病历id
               From 病人医嘱记录 B, 病人医嘱报告 C
               Where c.医嘱id = b.Id And Nvl(b.主页id, 0) = 0 And b.挂号单 Is Null And b.相关id Is Null And b.待转出 Is Null And
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

--117082:秦龙,2017-11-21,编码压缩后的处理
--115695:秦龙,2017-11-09,修改扩充时的编码字符长度
Create Or Replace Procedure Zl_诊疗分类目录_Update
(
  Id_In      诊疗分类目录.Id%Type,
  上级id_In  诊疗分类目录.上级id%Type,
  编码_In    诊疗分类目录.编码%Type,
  名称_In    诊疗分类目录.名称%Type,
  简码_In    诊疗分类目录.简码%Type,
  v_Brethren Number
  --是否对同级编码进行长度处理,0-否,1-是
) As
  v_Oldcode Varchar2(20); --原来的编码
  v_Parent  Varchar2(20); --上级编码
  v_Extend  Number(20); --扩充长度(为负表示压缩)
  v_Kind    Number(1); --当前项目的类型
  Err_Notfind Exception;

Begin
  Select RTrim(编码), 类型 Into v_Oldcode, v_Kind From 诊疗分类目录 Where ID = Id_In;
  If v_Oldcode Is Null Then
    Raise Err_Notfind;
  End If;
  --修改项目本身
  Update 诊疗分类目录
  Set 上级id = Decode(上级id_In, 0, Null, 上级id_In), 编码 = 编码_In, 名称 = 名称_In, 简码 = 简码_In
  Where ID = Id_In;
  --修改本系各级下属编码
  Update 诊疗分类目录
  Set 编码 = 编码_In || Substr(编码, Length(v_Oldcode) + 1)
  Where 编码 <> 编码_In And 编码 Like v_Oldcode || '_%' And 类型 = v_Kind;
  --调整同级编码的长度
  If v_Brethren = 1 Then
    If Nvl(上级id_In, 0) <> 0 Then
      Select 编码 Into v_Parent From 诊疗分类目录 Where ID = 上级id_In;
    Else
      v_Parent := Null;
    End If;
    Begin
      Select Length(RTrim(编码_In)) - Length(RTrim(编码))
      Into v_Extend
      From 诊疗分类目录
      Where (上级id = 上级id_In Or 上级id Is Null And Nvl(上级id_In, 0) = 0) And ID <> Id_In And 类型 = v_Kind And Rownum = 1;
    Exception
      When Others Then
        v_Extend := 0;
    End;
    If v_Extend > 0 Then
      --扩充处理
      If v_Parent Is Null Then
        Update 诊疗分类目录
        Set 编码 = LPad('0', v_Extend, '0') || 编码
        Where 类型 = v_Kind And ID Not In (Select ID From 诊疗分类目录 Start With ID = Id_In Connect By Prior ID = 上级id);
      Else
        Update 诊疗分类目录
        Set 编码 = v_Parent || LPad('0', v_Extend, '0') || Substr(编码, Length(v_Parent) + 1)
        Where 类型 = v_Kind And 编码 Like v_Parent || '_%' And
              ID Not In (Select ID From 诊疗分类目录 Start With ID = Id_In Connect By Prior ID = 上级id);
      End If;
    End If;
    If v_Extend < 0 Then
      --压缩处理
      If v_Parent Is Null Then
        Update 诊疗分类目录
        Set 编码 = Substr(编码, 1 + Abs(v_Extend))
        Where ID Not In (Select ID
                         From 诊疗分类目录
                         Where 类型 = v_Kind
                         Start With 上级id = Id_In
                         Connect By Prior ID = 上级id
                         Union All
                         Select ID From 诊疗分类目录 Where 类型 = v_Kind And ID = Id_In) And 类型 = v_Kind;
      Else
        Update 诊疗分类目录
        Set 编码 = v_Parent || Substr(编码, Length(v_Parent) + 1 + Abs(v_Extend))
        Where 编码 Like v_Parent || '_%' And
              ID Not In (Select ID
                         From 诊疗分类目录
                         Where 类型 = v_Kind
                         Start With 上级id = Id_In
                         Connect By Prior ID = 上级id
                         Union All
                         Select ID From 诊疗分类目录 Where 类型 = v_Kind And ID = Id_In) And 类型 = v_Kind;
      End If;
    End If;
  End If;

Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该项目不存在，可能已被其他用户删除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗分类目录_Update;
/

--115695:秦龙,2017-11-09,修改上级编码、扩充长度的长度
Create Or Replace Procedure Zl_诊疗分类目录_Insert
(
  Id_In      诊疗分类目录.Id%Type,
  上级id_In  诊疗分类目录.上级id%Type,
  编码_In    诊疗分类目录.编码%Type,
  名称_In    诊疗分类目录.名称%Type,
  简码_In    诊疗分类目录.简码%Type,
  类型_In    诊疗分类目录.类型%Type,
  v_Brethren Number
  --是否对同级编码进行长度处理,0-否,1-是
) As
  v_Parent Varchar2(20); --上级编码
  v_Extend Number(20); --扩充长度(为负表示压缩)
Begin
  --调整同级编码的长度
  If v_Brethren = 1 Then
    If Nvl(上级id_In, 0) <> 0 Then
      Select 编码 Into v_Parent From 诊疗分类目录 Where ID = 上级id_In;
    Else
      v_Parent := Null;
    End If;
    Begin
      Select Length(RTrim(编码_In)) - Length(RTrim(编码))
      Into v_Extend
      From 诊疗分类目录
      Where (上级id = 上级id_In Or 上级id Is Null And Nvl(上级id_In, 0) = 0) And ID <> Id_In And 类型 = 类型_In And Rownum = 1;
    Exception
      When Others Then
        v_Extend := 0;
    End;
    If v_Extend > 0 Then
      --扩充处理
      If v_Parent Is Null Then
        Update 诊疗分类目录 Set 编码 = LPad('0', v_Extend, '0') || 编码 Where ID <> Id_In And 类型 = 类型_In;
      Else
        Update 诊疗分类目录
        Set 编码 = v_Parent || LPad('0', v_Extend, '0') || Substr(编码, Length(v_Parent) + 1)
        Where 编码 Like v_Parent || '_%' And 类型 = 类型_In;
      End If;
    End If;
    If v_Extend < 0 Then
      --压缩处理
      If v_Parent Is Null Then
        Update 诊疗分类目录 Set 编码 = Substr(编码, 1 + Abs(v_Extend)) Where ID <> Id_In And 类型 = 类型_In;
      Else
        Update 诊疗分类目录
        Set 编码 = v_Parent || Substr(编码, Length(v_Parent) + 1 + Abs(v_Extend))
        Where 编码 Like v_Parent || '_%' And 类型 = 类型_In;
      End If;
    End If;
  End If;
  --添加本记录
  Insert Into 诊疗分类目录
    (ID, 上级id, 编码, 名称, 类型, 简码)
  Values
    (Id_In, Decode(上级id_In, 0, Null, 上级id_In), 编码_In, 名称_In, 类型_In, 简码_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗分类目录_Insert;
/

--116034:梁唐彬,2017-11-03,路径生成病历无内容问题
CREATE OR REPLACE Function Zl_Lob_ReadForPath
(
  Tab_In     In Number,
  Key_In     In Varchar2,
  Pos_In     In Number,
  Moved_In   In Number := 0,
  Lobtype_In In Number := 0
  --参数说明：
  --Tab_In：包含LOB的数据表
  --        5-电子病历格式
  --Key_In：数据记录的关键字
  --Pos_In：从0开始不断读取，直到返回为空
  --Moved_In: 0正常记录,1读取转储后备表记录
  --LobType_IN:0-BLOb,1-CLOB
) Return Varchar2 Is
  l_Blob   Blob;
  l_Clob   Clob;
  v_Buffer Varchar2(32767);
  n_Amount Number := 2000;
  n_Offset Number := 1;
Begin
  If Tab_In = 5 Then
    If Moved_In = 0 Then
      Select 内容 Into l_Blob From 电子病历格式 Where 文件id = To_Number(Key_In);
    Else
      Select 内容 Into l_Blob From H电子病历格式 Where 文件id = To_Number(Key_In);
    End If;
  End If;

  n_Offset := n_Offset + Pos_In * n_Amount;
  If Lobtype_In = 1 Then
    If l_Clob Is Null Then
      v_Buffer := Null;
    Else
      Dbms_Lob.Read(l_Clob, n_Amount, n_Offset, v_Buffer);
    End If;
  Else
    If l_Blob Is Null Then
      v_Buffer := Null;
    Else
      Dbms_Lob.Read(l_Blob, n_Amount, n_Offset, v_Buffer);
    End If;
  End If;
  Return v_Buffer;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Lob_ReadForPath;
/

--114647:冉俊明,2017-11-02,诊记帐单在审核时，当发药窗口发生变化后，没有更新药品收发记录的发药窗口
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
  
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(r_Bill.实收金额, 0)
    Where 病人id = r_Bill.病人id And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (r_Bill.病人id, 1, 1, r_Bill.实收金额, 0);
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

--92026:蒋廷中,2017-11-01,自由录入医嘱支持执行
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

--112953:梁唐彬,2017-09-11,药品说明书知识库
CREATE OR REPLACE Function Zl_Drugexplain_Readlob
(
  Key_In In Varchar2,
  Col_In In Varchar2,
  Pos_In     In Number
) Return Varchar2 Is
  l_Clob   Clob;
  v_Buffer Varchar2(32767);
  n_Amount Number := 2000;
  n_Offset Number := 1;
Begin

  If Col_In = '化学名称' Then
    Select t.化学名称 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '性状' Then
    Select t.性状 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '药理毒理' Then
    Select t.药理毒理 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '药代动力学' Then
    Select t.药代动力学 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '适应症' Then
    Select t.适应症 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '用法用量' Then
    Select t.用法用量 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '不良反应' Then
    Select t.不良反应 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '禁忌症' Then
    Select t.禁忌症 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '注意事项' Then
    Select t.注意事项 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '孕妇用药' Then
    Select t.孕妇用药 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '儿童用药' Then
    Select t.儿童用药 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '相互作用' Then
    Select t.相互作用 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '药物过量' Then
    Select t.药物过量 Into l_Clob From 药品说明书 T Where ID = Key_In;
  Elsif Col_In = '贮藏条件' Then
    Select t.贮藏条件 Into l_Clob From 药品说明书 T Where ID = Key_In;
  End If;

  n_Offset := n_Offset + Pos_In * n_Amount;
  If l_Clob Is Null Then
    v_Buffer := Null;
  Else
    Dbms_Lob.Read(l_Clob, n_Amount, n_Offset, v_Buffer);
  End If;
  Return v_Buffer;
Exception
  When No_Data_Found Then
    Return Null;
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugexplain_Readlob;
/

--109619:李业庆,2017-10-20,库存价格特殊情况处理
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
           a.内部条码, b.是否变价, a.单量, a.频次, a.摘要, Nvl(a.费用id, 0) As 费用id
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
      Select Nvl(实际数量, 0), 平均成本价
      Into n_库存数量, n_库存平均价
      From 药品库存
      Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
    Exception
      When Others Then
        n_库存数量 := 0;
    End;
  
    If n_库存平均价 Is Null Or n_库存平均价 < 0 Then
      Select 成本价 Into n_库存平均价 From 药品规格 Where 药品id = v_Detail.药品id;
    
      If n_库存平均价 Is Null Or n_库存平均价 < 0 Then
        n_库存平均价 := 0;
      End If;
    End If;
  
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
        
          Insert Into 药品入库信息
            (药品id, 库房id, 批次, 入库日期)
            Select v_Detail.药品id, v_Detail.库房id, v_Detail.批次, v_Detail.审核日期
            From Dual
            Where Not Exists (Select 1
                   From 药品入库信息
                   Where 药品id = v_Detail.药品id And 库房id = v_Detail.库房id And 批次 = v_Detail.批次);
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
          
            Insert Into 药品入库信息
              (药品id, 库房id, 批次, 入库日期)
              Select v_Detail.药品id, v_Detail.库房id, v_Detail.批次, v_Detail.审核日期
              From Dual
              Where Not Exists (Select 1
                     From 药品入库信息
                     Where 药品id = v_Detail.药品id And 库房id = v_Detail.库房id And 批次 = v_Detail.批次);
          End If;
        
          --调整药品批号对照表中的价格
          If v_Detail.摘要 = '成本价调价' Then
            Update 药品批号对照 Set 成本价 = n_成本价 Where 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
          End If;
        End If;
      Elsif v_Detail.单据 = 13 Then
        --单据=13 的售价修正记录 同步更新的金额和差价，所以不需要重算平均成本价
        If v_Detail.费用id = 0 Then
          Update 药品库存
          Set 零售价 = Decode(n_时价分批, 1, n_零售价, Null), 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_差价
          Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
        Else
          --调价修正时不更新零售价
          Update 药品库存
          Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_差价
          Where 性质 = 1 And 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
        End If;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 零售价)
          Values
            (v_Detail.库房id, v_Detail.药品id, v_Detail.批次, 1, 0, 0, n_零售金额, n_零售金额, Decode(n_时价分批, 1, n_零售价, Null));
        
          Insert Into 药品入库信息
            (药品id, 库房id, 批次, 入库日期)
            Select v_Detail.药品id, v_Detail.库房id, v_Detail.批次, v_Detail.审核日期
            From Dual
            Where Not Exists (Select 1
                   From 药品入库信息
                   Where 药品id = v_Detail.药品id And 库房id = v_Detail.库房id And 批次 = v_Detail.批次);
        End If;
      
        --调整药品批号对照表中的价格
        If v_Detail.摘要 = '药品调价' Then
          Update 药品批号对照 Set 售价 = n_零售价 Where 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次;
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
        
          --104843：34版本入库时不更新零售价（如果是无库存的是通过插入新库存记录产生了价格，如果已存在了就不更新）
          /*          --外购入库和其他入库审核时
          If (v_Detail.单据 = 1 And v_Detail.记录状态 = 1 And 财务审核_In = 0) Or (v_Detail.单据 = 4 And v_Detail.记录状态 = 1) Then
            Update 药品库存
            Set 零售价 = Decode(n_时价分批, 1, n_零售价, Null)
            Where 库房id = v_Detail.库房id And 药品id = v_Detail.药品id And Nvl(批次, 0) = v_Detail.批次 And 性质 = 1;
          End If;*/
        
          --不分批入库需要重算成本价
          --外购退货、财务审核和所有冲销业务不更新平均成本价，保持当前价格
          If (v_Detail.单据 = 1 And v_Detail.发药方式 = 1) Or Mod(v_Detail.记录状态, 3) = 2 Or (v_Detail.单据 = 1 And 财务审核_In = 1) Then
            Null;
          Else
            --按总金额/总数量方式计算平均成本价而不用（金额-差价）/数量是为了数据的准确性
            n_总数量 := (n_库存数量 + n_实际数量);
            If n_总数量 <> 0 And v_Detail.批次 = 0 Then
              --104843：不分批的才重算，分批的不处理
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
              实际差价 = Nvl(实际差价, 0) + n_差价, 平均成本价 = Decode(平均成本价, Null, n_成本价, 平均成本价),
              上次采购价 = Decode(上次采购价, Null, n_成本价, 上次采购价)
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
        
          Insert Into 药品入库信息
            (药品id, 库房id, 批次, 入库日期)
            Select v_Detail.药品id, v_Detail.库房id, v_Detail.批次, v_Detail.审核日期
            From Dual
            Where Not Exists (Select 1
                   From 药品入库信息
                   Where 药品id = v_Detail.药品id And 库房id = v_Detail.库房id And 批次 = v_Detail.批次);
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

--118817:殷瑞,2017-12-21,处理错误的查询异常
--118364:殷瑞,2017-12-19,修正退发药品“配药人”和“配药日期”可能为空的情况
--82526:李业庆,2017-12-05,批量发药产生汇总发药号
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
  未取药_In        In 药品收发记录.是否未取药%Type := Null,
  汇总发药号_In    In 药品收发记录.汇总发药号%Type := Null
) Is
  --住院病人
  Cursor c_Modifybillin Is
    Select a.Id, a.药品id, a.入出类别id, a.入出系数, a.费用id, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) 数量, Nvl(a.零售金额, 0) 金额, Nvl(a.批次, 0) 批次,
           a.供药单位id, a.成本价, a.批号, a.产地, a.效期, a.生产日期, a.批准文号, b.病人id, b.序号, Nvl(c.处方类型, Nvl(a.注册证号, 0)) 处方类型,
           Nvl(a.零售价, 0) As 零售价, a.记录状态
    From 药品收发记录 A, 住院费用记录 B, 未发药品记录 C
    Where a.单据 = c.单据 And a.No = c.No And Nvl(a.库房id, 0) = Nvl(c.库房id, 0) And a.No = No_In And a.单据 = Bill_In And
          (a.库房id + 0 = Partid_In Or a.库房id Is Null) And Nvl(a.摘要, '小宝') <> '拒发' And a.费用id = b.Id And
          Nvl(b.执行状态, 0) <> 1 And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null;

  --门诊病人
  Cursor c_Modifybillout Is
    Select a.Id, a.药品id, a.入出类别id, a.入出系数, a.费用id, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) 数量, Nvl(a.零售金额, 0) 金额, Nvl(a.批次, 0) 批次,
           a.供药单位id, a.成本价, a.批号, a.产地, a.效期, a.生产日期, a.批准文号, b.病人id, b.序号, Nvl(c.处方类型, Nvl(a.注册证号, 0)) 处方类型,
           Nvl(a.零售价, 0) As 零售价, b.No, b.记录性质, a.记录状态
    From 药品收发记录 A, 门诊费用记录 B, 未发药品记录 C
    Where a.单据 = c.单据 And a.No = c.No And Nvl(a.库房id, 0) = Nvl(c.库房id, 0) And a.No = No_In And a.单据 = Bill_In And
          (a.库房id + 0 = Partid_In Or a.库房id Is Null) And Nvl(a.摘要, '小宝') <> '拒发' And a.费用id = b.Id And
          Nvl(b.执行状态, 0) <> 1 And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null
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
  v_配药人          药品收发记录.配药人%Type;
  v_配药日期        药品收发记录.配药日期%Type;
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
  Set 配药人 = Decode(配药人_In, Null, 配药人, 配药人_In), 配药日期 = Decode(配药人_In, Null, 配药日期, Date操作时间), 汇总发药号 = 汇总发药号_In
  Where NO = No_In And 单据 = Bill_In And (库房id + 0 = Partid_In Or 库房id Is Null) And Mod(记录状态, 3) = 1 And 审核人 Is Not Null;

  --修正退发药品“配药人”和“配药日期”可能为空的情况
  Begin
    If 配药人_In Is Null Then
      Select 配药人, 配药日期
      Into v_配药人, v_配药日期
      From 药品收发记录
      Where NO = No_In And 单据 = Bill_In And (库房id + 0 = Partid_In Or 库房id Is Null) And Mod(记录状态, 3) = 1 And
            配药人 Is Not Null And Rownum = 1
      Order By 记录状态 Desc;
    
      Update 药品收发记录
      Set 配药人 = v_配药人, 配药日期 = v_配药日期
      Where NO = No_In And 单据 = Bill_In And (库房id + 0 = Partid_In Or 库房id Is Null) And Mod(记录状态, 3) = 1 And 审核人 Is Null;
    End If;
  Exception
    When Others Then
      v_配药人   := Null;
      v_配药日期 := Null;
  End;

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
      If v_Modifybillout.记录状态 = 1 Then
        --原始发药记录，取最新价格
        n_平均成本价 := Round(Zl_Fun_Getoutcost(v_Modifybillout.药品id, v_Modifybillout.批次, Partid_In), 5);
      Else
        --退药再发记录，取原始单据价格
        Select a.成本价
        Into n_平均成本价
        From 药品收发记录 A, 药品收发记录 B
        Where b.Id = v_Modifybillout.Id And a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And a.药品id + 0 = b.药品id And
              a.序号 = b.序号 And Nvl(a.批次, 0) = Nvl(b.批次, 0) And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0);
      End If;
    
      Dbl成本金额 := Round(n_平均成本价 * Nvl(v_Modifybillout.数量, 0), Intdigit_In);
      --零售金额
      Dbl实际金额 := Nvl(v_Modifybillout.金额, 0);
      --差价
      Dbl实际差价 := Round(Dbl实际金额 - Dbl成本金额, Intdigit_In);
    
      --更新药品收发记录的零售金额、成本金额、差价、审核人等信息
      Update 药品收发记录
      Set 成本价 = n_平均成本价, 成本金额 = Dbl成本金额, 差价 = Dbl实际差价, 库房id = Partid_In, 配药人 = Decode(配药人_In, Null, 配药人, 配药人_In),
          核查人 = 核查人_In, 核查日期 = v_核查日期, 配药日期 = Decode(配药人_In, Null, 配药日期, Date操作时间),
          填制人 = Decode(校验人_In, Null, 填制人, 校验人_In), 审核人 = Decode(People_In, Null, Zl_Username, People_In),
          审核日期 = Date操作时间, 发药方式 = 发药方式_In, 注册证号 = v_Modifybillout.处方类型, 是否未取药 = 未取药_In, 汇总发药号 = 汇总发药号_In
      Where ID = v_Modifybillout.Id;
    
      If Bln收费与发药分离 = 1 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - Nvl(v_Modifybillout.数量, 0), 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybillout.数量, 0),
            实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybillout.金额, 0), 实际差价 = Nvl(实际差价, 0) - Dbl实际差价,
            平均成本价 = Decode(平均成本价, Null, n_平均成本价, 平均成本价), 上次采购价 = Decode(上次采购价, Null, n_平均成本价, 上次采购价)
        Where 库房id + 0 = Partid_In And 药品id = v_Modifybillout.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybillout.批次;
      Else
        Update 药品库存
        Set 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybillout.数量, 0), 实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybillout.金额, 0),
            实际差价 = Nvl(实际差价, 0) - Dbl实际差价, 平均成本价 = Decode(平均成本价, Null, n_平均成本价, 平均成本价),
            上次采购价 = Decode(上次采购价, Null, n_平均成本价, 上次采购价)
        Where 库房id + 0 = Partid_In And 药品id = v_Modifybillout.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybillout.批次;
      End If;
    
      If Sql%RowCount = 0 Then
        If Bln收费与发药分离 = 1 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号, 平均成本价)
          Values
            (Partid_In, v_Modifybillout.药品id, v_Modifybillout.批次, 1, 0 - Nvl(v_Modifybillout.数量, 0),
             0 - Nvl(v_Modifybillout.数量, 0), 0 - Nvl(v_Modifybillout.金额, 0), 0 - Dbl实际差价, v_Modifybillout.供药单位id,
             n_平均成本价, v_Modifybillout.批号, v_Modifybillout.产地, v_Modifybillout.效期, v_Modifybillout.生产日期,
             v_Modifybillout.批准文号, n_平均成本价);
        Else
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号, 平均成本价)
          Values
            (Partid_In, v_Modifybillout.药品id, v_Modifybillout.批次, 1, 0 - Nvl(v_Modifybillout.数量, 0),
             0 - Nvl(v_Modifybillout.金额, 0), 0 - Dbl实际差价, v_Modifybillout.供药单位id, n_平均成本价, v_Modifybillout.批号,
             v_Modifybillout.产地, v_Modifybillout.效期, v_Modifybillout.生产日期, v_Modifybillout.批准文号, n_平均成本价);
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
      If v_Modifybillin.记录状态 = 1 Then
        --原始发药记录，取最新价格
        n_平均成本价 := Round(Zl_Fun_Getoutcost(v_Modifybillin.药品id, v_Modifybillin.批次, Partid_In), 5);
      Else
        --退药再发记录，取原始单据价格
        Select a.成本价
        Into n_平均成本价
        From 药品收发记录 A, 药品收发记录 B
        Where b.Id = v_Modifybillin.Id And a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And a.药品id + 0 = b.药品id And
              a.序号 = b.序号 And Nvl(a.批次, 0) = Nvl(b.批次, 0) And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0);
      End If;
    
      Dbl成本金额 := Round(n_平均成本价 * Nvl(v_Modifybillin.数量, 0), Intdigit_In);
      --零售金额
      Dbl实际金额 := Nvl(v_Modifybillin.金额, 0);
      --差价
      Dbl实际差价 := Round(Dbl实际金额 - Dbl成本金额, Intdigit_In);
    
      --更新药品收发记录的零售金额、成本金额、差价、审核人等信息
      Update 药品收发记录
      Set 成本价 = n_平均成本价, 成本金额 = Dbl成本金额, 差价 = Dbl实际差价, 库房id = Partid_In, 配药人 = Decode(配药人_In, Null, 配药人, 配药人_In),
          核查人 = 核查人_In, 核查日期 = v_核查日期, 配药日期 = Decode(配药人_In, Null, 配药日期, Date操作时间),
          填制人 = Decode(校验人_In, Null, 填制人, 校验人_In), 审核人 = Decode(People_In, Null, Zl_Username, People_In),
          审核日期 = Date操作时间, 发药方式 = 发药方式_In, 注册证号 = v_Modifybillin.处方类型, 汇总发药号 = 汇总发药号_In
      Where ID = v_Modifybillin.Id;
    
      If Bln收费与发药分离 = 1 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) - Nvl(v_Modifybillin.数量, 0), 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybillin.数量, 0),
            实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybillin.金额, 0), 实际差价 = Nvl(实际差价, 0) - Dbl实际差价,
            平均成本价 = Decode(平均成本价, Null, n_平均成本价, 平均成本价), 上次采购价 = Decode(上次采购价, Null, n_平均成本价, 上次采购价)
        Where 库房id + 0 = Partid_In And 药品id = v_Modifybillin.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybillin.批次;
      Else
        Update 药品库存
        Set 实际数量 = Nvl(实际数量, 0) - Nvl(v_Modifybillin.数量, 0), 实际金额 = Nvl(实际金额, 0) - Nvl(v_Modifybillin.金额, 0),
            实际差价 = Nvl(实际差价, 0) - Dbl实际差价, 平均成本价 = Decode(平均成本价, Null, n_平均成本价, 平均成本价),
            上次采购价 = Decode(上次采购价, Null, n_平均成本价, 上次采购价)
        Where 库房id + 0 = Partid_In And 药品id = v_Modifybillin.药品id And 性质 = 1 And Nvl(批次, 0) = v_Modifybillin.批次;
      End If;
    
      If Sql%RowCount = 0 Then
        If Bln收费与发药分离 = 1 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号, 平均成本价)
          Values
            (Partid_In, v_Modifybillin.药品id, v_Modifybillin.批次, 1, 0 - Nvl(v_Modifybillin.数量, 0),
             0 - Nvl(v_Modifybillin.数量, 0), 0 - Nvl(v_Modifybillin.金额, 0), 0 - Dbl实际差价, v_Modifybillin.供药单位id, n_平均成本价,
             v_Modifybillin.批号, v_Modifybillin.产地, v_Modifybillin.效期, v_Modifybillin.生产日期, v_Modifybillin.批准文号, n_平均成本价);
        Else
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次产地, 效期, 上次生产日期, 批准文号, 平均成本价)
          Values
            (Partid_In, v_Modifybillin.药品id, v_Modifybillin.批次, 1, 0 - Nvl(v_Modifybillin.数量, 0),
             0 - Nvl(v_Modifybillin.金额, 0), 0 - Dbl实际差价, v_Modifybillin.供药单位id, n_平均成本价, v_Modifybillin.批号,
             v_Modifybillin.产地, v_Modifybillin.效期, v_Modifybillin.生产日期, v_Modifybillin.批准文号, n_平均成本价);
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

--104221:余伟节,2017-12-07,身份证号检查
Create Or Replace Function Zl_Fun_Checkidcard
(
  Idcard_In   In varchar2,
  Calcdate_In In Date := Null
) Return varchar2 Is
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
  v_校验位    varchar2(50);
  v_Pattern   varchar2(500);
  v_Err_Msg   varchar2(2000);
  v_性别      varchar2(100);
  v_年龄      varchar2(100);
  d_Curr_Time Date;
  d_出生日期  Date;

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
        d_出生日期 := To_Date('19' || Substr(Idcard_In, 7, 6), 'yyyy-mm-dd');
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
        d_出生日期 := To_Date(Substr(Idcard_In, 7, 8), 'yyyy-mm-dd');
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


--118323:李业庆,2017-12-12,修改删除划价记账时高值卫材库存处理遗漏
Create Or Replace Procedure Zl_门诊划价记录_Delete
(
  No_In       门诊费用记录.No%Type,
  序号_In     Varchar2 := Null,
  自动清除_In Number := 0
) As
  --功能：删除一张门诊划价单据
  --入参：
  --       序号_In：主要用于门诊医生站作废单个药品
  --      自动清除_in：是否自动清除划价单 zl_门诊划价记录_clear 在调用
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
    If Nvl(自动清除_In, 0) = 1 Then
      --自动清除划价单调用时不报错，直接退出
      Return;
    Else
      v_Err_Msg := '要删除的费用记录不存在，可能已经删除或已经收费。';
      Raise Err_Item;
    End If;
  End If;
  --是否已经执行
  If Nvl(n_已执行_Count, 0) > 0 Then
    v_Err_Msg := '要删除的费用记录中包含已执行的内容！';
    Raise Err_Item;
  End If;

  --医嘱费用：检查正在执行的医嘱(注意已执行的情况在下面检查,因为不传 序号_IN 这种情况费用界面已限制)
  --自动清除划价单调用时，由于只会传入药品卫材的对应序号，所以不用检查医嘱；
  --如果检查医嘱，可能同一个医嘱中既有药品，也有其它项目，而其它项目正在执行或已执行时该药品划价记录将删除不掉
  If Nvl(自动清除_In, 0) = 0 Then
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
    If Nvl(自动清除_In, 0) = 1 Then
      --自动清除划价单调用时不报错，直接退出
      Return;
    Else
      v_Err_Msg := '要删除的费用记录不存在，可能已经删除或已经收费。';
      Raise Err_Item;
    End If;
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


--118323:李业庆,2017-12-12,修改删除划价记账时高值卫材库存处理遗漏
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

---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--113763:杨周一,2017-12-07,DBA工具名称修改
Insert Into Zltools.Zlfilesupgrade
  (序号, 加入日期, 安装路径, 文件类型, 文件名, 版本号, 修改日期, 所属系统, 业务部件, Md5, 文件说明, 自动注册, 强制覆盖, 附加安装路径)
  Select 序号, To_Date('2017-07-05 17:22:54', 'yyyy-mm-dd hh24:mi:ss'), '[APPSOFT]', 0, 'ZLDBATOOLS.EXE', Null, Null, Null,
         Null, Null, '部件功能:DBA管理工具单独执行文件', 0, 0, Null
  From Dual A, (Select Nvl(Max(To_Number(序号)), 0) + 1 序号 From zlFilesUpgrade) B
  Where Not Exists (Select 1 From Zltools.Zlfilesupgrade Where Upper(文件名) = 'ZLDBATOOLS.EXE');
--00000:刘硕,2017-12-27,文件清单调整
Update Zltools.Zlfilesupgrade
Set 安装路径 = '[APPSOFT]'
Where Upper(文件名) = 'ZLRISDUMPTOOL.EXE' And Not Exists
 (Select 1 From Zltools.Zlfilesupgrade Where Upper(文件名) = 'ZLRISDUMPTOOL.EXE' And Upper(安装路径) = '[APPSOFT]');
Delete Zltools.Zlfilesupgrade Where Upper(文件名) = 'ZLRISDUMPTOOL.EXE' And Upper(安装路径) <> '[APPSOFT]';
Delete Zltools.Zlfilesupgrade Where Upper(文件名) In ('ZL9PEISDEVANALYSE', 'ZL9PEISINSTRUMENT');

--系统版本号
Update zlSystems Set 版本号='10.34.140' Where 编号=&n_System;
--部件版本号
Commit;

