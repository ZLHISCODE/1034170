--本脚本支持从ZLHIS+ v10.11.0 升级到 v10.12.0
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--6859
Alter Table 药品库房货位 Modify 编码 Varchar2(5);

--PACS
Alter Table 影像检查记录 Add 联系电话 Varchar2(20);
Alter Table 影像检查记录 Drop Constraint 影像检查记录_UQ_检查号 Cascade;
Alter Table H影像检查记录 Drop Constraint H影像检查记录_UQ_检查号 Cascade;
Create Index 影像检查记录_IX_检查号 on 影像检查记录 (检查号, 影像类别) PCTFREE 10 TABLESPACE zl9CisRec;
Create Index 影像临时记录_IX_检查号 on 影像临时记录 (检查号, 影像类别) PCTFREE 10 TABLESPACE zl9CisRec;

--6950
Alter Table 药品收发记录 Add 批准文号 VARCHAR2(40);
Alter Table H药品收发记录 Add 批准文号 VARCHAR2(40);
Alter Table 药品库存 Add 批准文号 VARCHAR2(40);
--7135
ALTER TABLE 病人过敏药物 DROP CONSTRAINT 病人过敏药物_UQ_过敏药物ID;
--7303
Create Or Replace View 诊断情况 As 
Select 病人id, 主页Id,疾病id,诊断描述 As 描述信息,诊断类型, 
   出院情况,诊断次序,编码序号, 是否未治, 是否疑诊
From 病人诊断记录 Where 记录来源=2;

--7163
Create Table 病人担保记录(
    病人ID      NUMBER(18),
    担保人      VARCHAR2(20),
    担保额      NUMBER(16,5),
	担保性质    NUMBER(1),
    操作员编号  VARCHAR2(6),
    操作员姓名  VARCHAR2(20),
    发生时间    Date
    )
    TABLESPACE zl9Patient
    PCTFREE 10 PCTUSED 60 STORAGE (NEXT 8K PCTINCREASE 0 MAXEXTENTS UNLIMITED);

ALTER TABLE 病人担保记录 ADD CONSTRAINT 病人担保记录_PK PRIMARY KEY (病人ID,发生时间) USING INDEX PCTFREE 5 TABLESPACE zl9Patient;
ALTER TABLE 病人担保记录 ADD CONSTRAINT 病人担保记录_FK_病人ID FOREIGN KEY (病人ID) REFERENCES 病人信息(病人ID) ON DELETE CASCADE;

--7149
Alter Table 药品采购计划 Modify 期间 Varchar(8);

--朱玉宝 2005-12-16 调整医保部分数据结构
ALTER TABLE 保险结算记录 ADD (就诊流水号 VARCHAR2(30),结算时间 DATE ,工作站 VARCHAR2(50),版本号 VARCHAR2(15));
ALTER TABLE 保险结算记录 ADD (医疗类别 VARCHAR2(3));
ALTER TABLE 保险结算记录 ADD (病种ID NUMBER(18));
ALTER TABLE 保险结算记录 ADD (病种名称 VARCHAR2(100));
ALTER TABLE 保险结算记录 ADD (并发症 VARCHAR2(200));
ALTER TABLE 保险结算记录 MODIFY 备注 VARCHAR2(500);

--因建了病人ID的外键，需同时修改病人身份合并过程
CREATE TABLE 就诊登记记录(
	险类 NUMBER(18),
	病人ID NUMBER(18),
	主页ID NUMBER(18),
	就诊时间 DATE ,
	状态 NUMBER(2),		--1-就诊中;0-未就诊
	医疗类别 VARCHAR2(3),
	帐户余额 NUMBER(16,5),
	病种ID NUMBER(18),
	病种名称 VARCHAR2(100),
	并发症 VARCHAR2(200),
	IC卡信息 VARCHAR2(200),
	HIS流水号 VARCHAR2(30),
	YB流水号 VARCHAR2(30),
	记录ID NUMBER(18),	--结帐ID，门诊可用此字段来关联，住院不必
	备注 VARCHAR2(200));
ALTER TABLE 就诊登记记录 ADD CONSTRAINT 就诊登记记录_PK PRIMARY KEY (险类,病人ID,就诊时间);
ALTER TABLE 就诊登记记录 ADD CONSTRAINT 就诊登记记录_FK_险类 FOREIGN KEY (险类) REFERENCES 保险类别(序号);
ALTER TABLE 就诊登记记录 ADD CONSTRAINT 就诊登记记录_FK_病人ID FOREIGN KEY (病人ID) REFERENCES 病人信息(病人ID);


--陈福容：体检
Alter Table 体检诊断建议 Add 是否疾病 Number(1);
Alter Table 体检诊断建议 Drop Column 描述;
Alter Table 体检诊断建议 Drop Column 疾病id;

Alter Table 体检人员结论 Drop Column 疾病id;

Update 体检诊断建议 Set 是否疾病=1 Where 末级=1;
--}陈福容

--增加卫材加成方案
CREATE TABLE 材料加成方案(
	序号		number(18), 
	最低价		number(16,5), 
	最高价		number(16,5), 
	加成率		number(16,5), 
	说明		varchar2(50))
    TableSpace zl9BaseItem
    PCTFREE 5 PCTUSED 90 STORAGE (NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED);
ALTER TABLE 材料加成方案 ADD CONSTRAINT 材料加成方案_PK PRIMARY KEY (序号) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;

--7157(ZT)
Alter Table 病人医嘱计价 Add 执行科室ID Number(18);
Alter Table 病人医嘱计价 
    Add CONSTRAINT 病人医嘱计价_FK_执行科室ID
    Foreign Key (执行科室ID) 
    References 部门表(ID);

--7108
Alter Table 诊疗项目目录 Modify 编码 VARCHAR2(20);
Alter Table 收费项目目录 Modify 编码 VARCHAR2(20);

--7058,7053
alter table 病人变动记录 add 病情 VARCHAR2(20);
alter table 病人变动记录 add 主治医师 VARCHAR2(20);
alter table 病人变动记录 add 主任医师 VARCHAR2(20);

--7038
Alter Table 药品特性 Add 品种医嘱 Number(1);

alter table 费别明细 add 计算方法 number(1) default 0;

--医嘱内容规则定义:用于类似公式编辑的方式来定义下达医嘱时,医嘱内容的生成规则
Create Table 医嘱内容定义(
    诊疗类别 VARCHAR2(1),
    医嘱内容 VARCHAR2(500))
    TABLESPACE zl9BaseItem
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);
ALTER TABLE 医嘱内容定义 ADD CONSTRAINT 医嘱内容定义_PK PRIMARY KEY (诊疗类别) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;

--电子签名结构调整
CREATE SEQUENCE 人员证书记录_ID START WITH 1;
CREATE TABLE 人员证书记录(
	ID NUMBER(18),
	人员ID NUMBER(18),
	CertDN VARCHAR2(300),
	CertSN VARCHAR2(100),
	SignCert VARCHAR2(2000),
	EncCert VARCHAR2(2000),
	注册时间 DATE)
    TABLESPACE zl9BaseItem
    PCTFREE 5 PCTUSED 90 STORAGE (NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED);
ALTER TABLE 人员证书记录 ADD CONSTRAINT 人员证书记录_PK PRIMARY KEY(ID) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;
Alter Table 人员证书记录 Add CONSTRAINT 人员证书记录_FK_人员ID Foreign Key (人员ID) References 人员表(ID);

CREATE SEQUENCE 医嘱签名记录_ID START WITH 1;
CREATE TABLE 医嘱签名记录(
	ID NUMBER(18),
	签名规则 NUMBER(2),--从1开始顺序编号,用于防止因产品升级，医嘱源文本生成规则发生变化
	签名信息 VARCHAR2(2000),
	证书ID	NUMBER(18),
	签名时间 DATE,
    签名人 VARCHAR2(20))
    TABLESPACE zl9CisRec
    PCTFREE 15 PCTUSED 70 
	STORAGE (NEXT 1K PCTINCREASE 0 MAXEXTENTS UNLIMITED);
Alter Table 医嘱签名记录 Add CONSTRAINT 医嘱签名记录_PK Primary Key (ID) USING INDEX PCTFREE 5 TABLESPACE zl9CisRec;
Alter Table 医嘱签名记录 Add CONSTRAINT 医嘱签名记录_FK_证书ID Foreign Key (证书ID) References 人员证书记录(ID);
CREATE INDEX 医嘱签名记录_IX_证书ID ON 医嘱签名记录(证书ID) PCTFREE 10 TABLESPACE zl9CisRec
/

--医嘱部份结构调整
Alter Table 病人医嘱记录 Add(姓名 VARCHAR2(20),性别 VARCHAR2(4),年龄 VARCHAR2(10));
Alter Table 病人医嘱状态 Add 签名ID Number(18);
Alter Table 病人医嘱状态 Add CONSTRAINT 病人医嘱状态_FK_签名ID Foreign Key (签名ID) References 医嘱签名记录(ID);
CREATE INDEX 病人医嘱状态_IX_签名ID ON 病人医嘱状态(签名ID) PCTFREE 10 TABLESPACE zl9CisRec
/

--历史后备数据结构
Create Table H医嘱签名记录 Tablespace zl9History As Select * From 医嘱签名记录 where 1=0;
Alter Table H医嘱签名记录 Add Constraint H医嘱签名记录_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9History;
Alter Table H医嘱签名记录 Add CONSTRAINT H医嘱签名记录_FK_证书ID Foreign Key (证书ID) References 人员证书记录(ID);
CREATE INDEX H医嘱签名记录_IX_证书ID ON H医嘱签名记录(证书ID) PCTFREE 10 TABLESPACE zl9History
/

Alter Table H病人医嘱记录 Add(姓名 VARCHAR2(20),性别 VARCHAR2(4),年龄 VARCHAR2(10));
Alter Table H病人医嘱状态 Add 签名ID Number(18);
Alter Table H病人医嘱状态 Add CONSTRAINT H病人医嘱状态_FK_签名ID Foreign Key (签名ID) References H医嘱签名记录(ID);
CREATE INDEX H病人医嘱状态_IX_签名ID ON H病人医嘱状态(签名ID) PCTFREE 10 TABLESPACE zl9History
/

-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--7427
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1342,'基本',user,'号码控制表','SELECT');

--修正临床性质数据把编码为1位或3位的临床性质修正为2位或4位在最前面加0
alter table 临床部门 drop CONSTRAINT 临床部门_FK_工作性质;
update 临床性质 set 编码='0'||编码 where length(编码)=1 or length(编码)=3;
update 临床部门 set 工作性质='0'||工作性质 where length(工作性质)=1 or length(工作性质)=3;
ALTER TABLE 临床部门 ADD CONSTRAINT 临床部门_FK_工作性质 FOREIGN KEY (工作性质) REFERENCES 临床性质(编码) ON DELETE CASCADE;
Insert into 临床性质(编码,名称,简码,序号) Values ('61','重症监护室(综合)','ZZJHS',163);
Insert into 临床性质(编码,名称,简码,序号) Values ('79','其他','ZZJHS',164);
Insert into 临床性质(编码,名称,简码,序号) Values ('99','管理科室','ZZJHS',165);
Insert into 临床性质(编码,名称,简码,序号) Values ('9901','感染(管理)科','ZZJHS',166);

--LIS基础
UPDATE ZLPrograms SET 标题='病人历史记录查询' WHERE 序号=1210;

--增加卫材加成方案
insert into 材料加成方案 (序号, 最低价, 最高价, 加成率, 说明) values (1, null, 500, 10, null);
insert into 材料加成方案 (序号, 最低价, 最高价, 加成率, 说明) values (2, 500, 2000, 8, null);
insert into 材料加成方案 (序号, 最低价, 最高价, 加成率, 说明) values (3, 2000, 5000, 5, null);
insert into 材料加成方案 (序号, 最低价, 最高价, 加成率, 说明) values (4, 5000, null, 2, null);

--固定增加中药方剂
Insert Into 药品剂型(编码,名称,简码)
Select * From (
	Select zl_Incstr(Max(编码)),'方剂','FJ' From 药品剂型)
Where Not Exists(Select 名称 From 药品剂型 Where 名称='方剂');

--7058
Update 病情 Set 名称='重' ,简码='Z'  Where 名称='急';

--7058 由于数据量可能很大,缺省只修正在院病人数据,请根据情况决定是否修正出院病人数据
Update 病案主页 Set 入院病况='重' Where 入院病况='急' And 出院日期 IS Null;
Update 病案主页 Set 当前病况='重' Where 当前病况='急' And 出院日期 IS Null;

--陈福容
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检项目报告_EMPTY','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'zl_GetTextRows','EXECUTE');

--7058  初始在院病人在病人变动记录中的病情
Create Or Replace Procedure zl_病人变动记录_修正
AS 
Begin
    For r_InPatient IN
        ( Select 病人ID,主页ID,当前病况 From 病案主页 Where 出院日期 IS Null)
    Loop
        Update 病人变动记录 Set 病情=r_InPatient.当前病况 Where 病人ID=r_InPatient.病人ID And 主页ID=r_InPatient.主页ID;
    End Loop;
EXCEPTION
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
End zl_病人变动记录_修正;
/
Execute zl_病人变动记录_修正;
DROP Procedure zl_病人变动记录_修正;


----外购入库、其他入库、移库过程中部分冲销造成售价金额、成本金额、差价计算错误修正（见7126问题登记）
--该问题只在固定条件下发生，会影响外购入库、其他入库、移库冲销的明细单据和库存表数据。
--由于只影响了明细的金额没有影响数量，故只提供冲销单据的明细金额修正脚本。
--如果要修正库存金额：时价药品可通过调价修正，定价药品可通过盘点修正。
--执行该脚本可能会出现的问题：由于小数精度问题，可能脚本会修正正常数据。用户可自己决定是否执行该修正过程。

CREATE OR REPLACE PROCEDURE ZL_修正冲销明细记录
IS
  	INTDIGIT NUMBER;
    INT冲销记录 Number;
BEGIN
    --获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';
    
    --修正冲销明细
    Update 药品收发记录 Set 成本金额=round(实际数量*成本价,INTDIGIT),零售金额=round(实际数量*零售价,INTDIGIT),
    差价=round(实际数量*零售价,INTDIGIT)-round(实际数量*成本价,INTDIGIT)
    Where 单据 In(1,4,6) And 零售金额<>round(零售价*实际数量,INTDIGIT) And Mod(记录状态,3)=2;
    
    --如果存在数量全部冲销而金额无法完全冲销的情况，则把差额更新到最后一条冲销记录上去
    For V_调整差额 In(Select 单据,No,药品id,实际数量,成本金额,零售金额 From 
        (Select 单据,No,药品id,Sum(实际数量) 实际数量,Sum(成本金额) 成本金额,Sum(零售金额) 零售金额  From 药品收发记录 Where 单据 In(1,4,6) 
        Group By 单据,No ,药品id Having Sum(实际数量)=0) Where 成本金额<>0 Or 零售金额<>0) Loop 
        
        --取最大冲销记录
        Select Max(记录状态) Into INT冲销记录 From 药品收发记录 Where 单据=V_调整差额.单据 And No=V_调整差额.No And 
        药品id=V_调整差额.药品id And Mod(记录状态,3)=2;
        
        --调整差额
        Update 药品收发记录 Set 成本金额=成本金额-V_调整差额.成本金额,零售金额=零售金额-V_调整差额.零售金额,
        差价=差价-( V_调整差额.零售金额-V_调整差额.成本金额)
        Where 单据=V_调整差额.单据 And No=V_调整差额.No And 药品id=V_调整差额.药品id And 记录状态=INT冲销记录;
        
    End Loop;
    
    Commit;
EXCEPTION
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_修正冲销明细记录;
/
--根据情况决定是否执行
--Execute ZL_修正冲销明细记录;
--Drop Procedure  ZL_修正冲销明细记录;
----部分冲销造成售价金额、成本金额、差价计算错误修正（见7126问题登记）

--7049
Insert Into 部门性质分类(编码,名称,简码,服务病人,说明)
Values('1','中医科','ZYK',0,'特指采用中医学给病人诊断并决定治疗方案的临床部门。在本系统中具备中医科特性科室的病人，允许填写中医的诊断，录入中医病案内容。');


--数据调整
Update 系统参数表 Set 参数名='门诊药嘱先作废后退药',参数说明='表示门诊药品医嘱需要先作废，然后再去退药，退费' Where 参数号=68;
Delete From 系统参数表 Where 参数号=25 And 参数名<>'电子签名认证中心';
Delete From 系统参数表 Where 参数号=26 And 参数名<>'电子签名使用场合';
Insert Into 系统参数表(参数号,参数名,参数值,缺省值,参数说明) 
	Select 25,'电子签名认证中心','0','0','电子签名认证中心的编号,从1开始顺序编号,0表示不使用电子签名;如1为辽宁CA认证中心' From Dual 
	Where Not Exists(Select 1 From 系统参数表 Where 参数号=25);
Insert Into 系统参数表(参数号,参数名,参数值,缺省值,参数说明) 
	Select 26,'电子签名使用场合','000','000','对不同场合是否使用电子签名进行控制,数字位数分别为:门诊,住院,医技,0-不控制,1-控制' From Dual
	Where Not Exists(Select 1 From 系统参数表 Where 参数号=26);

--刘兴宏加入
Insert Into 系统参数表(参数号,参数名,参数值,缺省值,参数说明)
Select 120,'卫材负数出库计算方式','0','0','对于卫材负数出库,确定成本价的计算方式：0-按指导差价率计算,1-按最后进价算(取卫材目录中的成本价）'  From dual;

Insert Into 系统参数表(参数号,参数名,参数值,缺省值,参数说明)
Select 121,'卫材分段加成率','0','0','对于时价卫材在入库时,按分段加成率计算.'  From dual;

Insert Into 系统参数表(参数号,参数名,参数值,缺省值,参数说明)
Select 123,'不严格控制指导价格','0','0','不严格控制卫材指导批价和指导售价' From dual;



Insert Into 医嘱内容定义(诊疗类别,医嘱内容) 
Select '5','[输入名]+iif([产地]<>"","("+[产地]+")","")+iif([规格]<>""," "+[规格],"")' From Dual Union All
Select '6','[输入名]+iif([产地]<>"","("+[产地]+")","")+iif([规格]<>""," "+[规格],"")' From Dual Union All
Select '8','"中药"+[付数]+"付,"+[中文频率]+","+[煎法]+","+[用法]+":"+[配方组成]' From Dual Union All
Select 'C','[检验项目]+iif([检验标本]<>"","("+[检验标本]+")","")' From Dual Union All
Select 'D','[检查项目]+iif([检查部位]<>"","("+[检查部位]+")","")' From Dual Union All
Select 'E','[诊疗项目]' From Dual Union All
Select 'F','Format([开始时间],"MM月dd日HH:mm")+iif([麻醉方法]<>""," 在 "+[麻醉方法]+" 下行 "," 行 ")+[主要手术]+iif([附加手术]<>""," 及 "+[附加手术],"")' From Dual Union All
Select 'H','[诊疗项目]' From Dual Union All
Select 'I','[诊疗项目]' From Dual Union All
Select 'K','[诊疗项目]' From Dual Union All
Select 'L','[诊疗项目]' From Dual Union All
Select 'M','[诊疗项目]' From Dual Union All
Select 'Z','[诊疗项目]' From Dual;

----本段因为执行较慢,可不执行,不影响新开医嘱的签名
--Update 病人医嘱记录 A Set 姓名=(Select 姓名 From 病人信息 B Where B.病人ID=A.病人ID);
--Update 病人医嘱记录 A Set 年龄=(Select 年龄 From 病人信息 B Where B.病人ID=A.病人ID);
--Update 病人医嘱记录 A Set 性别=(Select 性别 From 病人信息 B Where B.病人ID=A.病人ID);
----本段因为执行较慢,可不执行,不影响新开医嘱的签名
--Update H病人医嘱记录 A Set 姓名=(Select 姓名 From 病人信息 B Where B.病人ID=A.病人ID);
--Update H病人医嘱记录 A Set 年龄=(Select 年龄 From 病人信息 B Where B.病人ID=A.病人ID);
--Update H病人医嘱记录 A Set 性别=(Select 性别 From 病人信息 B Where B.病人ID=A.病人ID);

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--7404
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1342,'基本',user,'药品储备限额','SELECT');
Delete From zlProgPrivs Where 系统=100 And 序号=1342 And 功能='基本' And 对象='H药品收发记录';
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1342,'基本',user,'H药品收发记录','SELECT');

--6888
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1023,'所有库房','允许对所有库房药品进行储存库房设置');

--病人信息管理工作
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'基本',user,'病人担保记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'zl_病人信息_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'职业','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'医疗付款方式','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'学历','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'性别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'收入项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'收费细目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'收费特定项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'收费价目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'身份','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'社会关系','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'民族','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'结算方式应用','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'结算方式','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'婚姻状况','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'合约单位','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'号码控制表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'地区','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'国籍','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1101,'担保信息',user,'费别','SELECT');

--临床医护：增加医生站观片
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1203,'观片处理','进入观片工作站处理相关的影像。');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1203,'基本','影像检查图象',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1203,'基本','影像检查序列',USER,'SELECT');

Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1204,'观片处理','进入观片工作站处理相关的影像。');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1204,'基本','影像检查图象',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1204,'基本','影像检查序列',USER,'SELECT');

--6986
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1309,'医生查询','允许查看开单医生。');
--6884
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1205,'超期收回调整','超期发送收回医嘱时调整收回量的操作权限。');

--6896,7074:护士确认停止,暂停权限独立
Delete From zlProgPrivs Where 系统=100 And 序号=1205 And 功能='医嘱停止' And Upper(对象)='ZL_病人医嘱记录_确认停止';
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1205,'医嘱确认停止','对医生已停止的医嘱进行确认停止的操作权限。');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1205,'医嘱确认停止',user,'ZL_病人医嘱记录_确认停止','EXECUTE');

Delete From zlprogPrivs Where 系统=100 And 序号=1205 And 功能='医嘱停止' And Upper(对象) IN('ZL_病人医嘱记录_暂停','ZL_病人医嘱记录_启用');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1205,'医嘱暂停','已下达的医嘱进行暂停,启用的操作权限。');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1205,'医嘱暂停',user,'ZL_病人医嘱记录_暂停','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1205,'医嘱暂停',user,'ZL_病人医嘱记录_启用','EXECUTE');

--7080:皮试结果权限控制
Delete From zlProgPrivs Where 系统=100 And 序号=1203 And 功能='医嘱下达' And Upper(对象)='ZL_病人医嘱记录_皮试';
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1203,'皮试医嘱结果','对皮试的医嘱结果编辑的操作权限。');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1203,'皮试医嘱结果',user,'ZL_病人医嘱记录_皮试','EXECUTE');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1206,'参数设置','设置参数的操作权限。');

--朱玉宝 2005-12-16
--权限修正部分
--门诊/住院身份验证：就诊登记记录、zl_就诊登记记录_UPDATE、zl_就诊登记记录_DELETE
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1121,'基本',user,'就诊登记记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1121,'基本',user,'zl_就诊登记记录_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1121,'基本',user,'zl_就诊登记记录_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1131,'基本',user,'就诊登记记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1131,'基本',user,'zl_就诊登记记录_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1131,'基本',user,'zl_就诊登记记录_DELETE','EXECUTE');
--门诊结算：zl_就诊登记记录_结束
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1121,'基本',user,'zl_就诊登记记录_结束','EXECUTE');
--入院登记/出院登记及撤销：就诊登记记录、zl_就诊登记记录_UPDATE、zl_就诊登记记录_DELETE、zl_就诊登记记录_更新状态
--此处权限不含身份验证处权限
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1131,'基本',user,'zl_就诊登记记录_更新状态','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1132,'基本',user,'就诊登记记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1132,'基本',user,'zl_就诊登记记录_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1132,'基本',user,'zl_就诊登记记录_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1132,'基本',user,'zl_就诊登记记录_更新状态','EXECUTE');
--住院虚拟结算/病人费用查询：就诊登记记录
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1137,'基本',user,'就诊登记记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1139,'基本',user,'就诊登记记录','SELECT');


--电子签名权限调整
----医嘱内容定义权限
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1203,'基本',user,'医嘱内容定义','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1204,'基本',user,'医嘱内容定义','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1205,'基本',user,'医嘱内容定义','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1206,'基本',user,'医嘱内容定义','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1207,'基本',user,'医嘱内容定义','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1804,'基本',user,'医嘱内容定义','SELECT');
----基础参数设置
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1011,'基本',user,'诊疗项目类别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1011,'基本',user,'医嘱内容定义','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1011,'基本',user,'zl_医嘱内容定义_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1011,'基本',user,'zl_医嘱内容定义_Insert','EXECUTE');
----人员管理
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1002,'数字证书注册','对人员的数字证书进行注册的权限(发放或更换注册)。');
--医嘱签名权限填写错误修正
Delete From zlProgPrivs Where 系统=100 And 序号=1002 And 功能='基本' And 对象 IN('系统参数','系统参数表');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1002,'基本',user,'系统参数表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1002,'基本',user,'人员证书记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1002,'数字证书注册',user,'zl_人员证书记录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1002,'数字证书注册',user,'zl_人员证书记录_Delete','EXECUTE');
----门诊挂号管理
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1111,'下医嘱后退号','仅有已作废医嘱的病人允许退号。');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1111,'挂号',user,'ZL_就诊卡记录_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1111,'退号',user,'ZL_就诊卡记录_DELETE','EXECUTE');
----病人入出管理
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1132,'基本',user,'病案主页从表','SELECT');
----医技科室记帐
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1135,'诊疗负数记帐','诊疗单据负数记帐的操作权限。');
----门诊医生工作站
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1203,'基本',user,'人员证书记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1203,'基本',user,'医嘱签名记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1203,'基本',user,'H医嘱签名记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1203,'医嘱下达',user,'医嘱签名记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1203,'医嘱下达',user,'zl_医嘱签名记录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1203,'医嘱下达',user,'zl_医嘱签名记录_Delete','EXECUTE');
----住院医生工作站
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1204,'基本',user,'人员证书记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1204,'基本',user,'医嘱签名记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1204,'基本',user,'H医嘱签名记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1204,'医嘱下达',user,'医嘱签名记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1204,'医嘱下达',user,'zl_医嘱签名记录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1204,'医嘱下达',user,'zl_医嘱签名记录_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1204,'医嘱下达',user,'zl_病人医嘱记录_回退','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1204,'医嘱下达',user,'zl_病人医嘱记录_批量回退','EXECUTE');
--住院护士工作站:7031
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1205,'医嘱下达',user,'ZL_病人医嘱记录_校对','EXECUTE');
----医技工作站
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1206,'基本',user,'人员证书记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1206,'基本',user,'医嘱签名记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1206,'基本',user,'H医嘱签名记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1206,'补充下达医嘱',user,'医嘱签名记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1206,'补充下达医嘱',user,'zl_医嘱签名记录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1206,'补充下达医嘱',user,'zl_医嘱签名记录_Delete','EXECUTE');

----其它权限修正
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1205,'医嘱下达',user,'ZL_病人医嘱记录_更新审查','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1206,'补充下达医嘱',user,'ZL_病人医嘱记录_更新审查','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1804,'术后医嘱',user,'ZL_病人医嘱记录_更新审查','EXECUTE');

--导诊咨询系统
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1536,'基本',user,'保险模拟结算','SELECT');

--卫材目录
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1711,'分段加成率','按单价分段设置加成率。');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1711,'基本',user,'材料加成方案','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1711,'分段加成率',user,'ZL_材料加成方案_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1711,'分段加成率',user,'ZL_材料加成方案_INSERT','EXECUTE');
--卫材外购入库
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1712,'基本',user,'材料加成方案','SELECT');
--卫材其他入库
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1714,'基本',user,'材料加成方案','SELECT');


-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--PACS
CREATE OR REPLACE Procedure ZL_影像检查_BEGIN(
  执行间_IN   病人医嘱发送.执行间%Type,
  检查号_IN   影像检查记录.检查号%Type,
	医嘱ID_IN		影像检查记录.医嘱ID%Type,
	发送号_IN		影像检查记录.发送号%Type,
  影像类别_IN 影像检查记录.影像类别%Type,
  姓名_IN     影像检查记录.姓名%Type,
  英文名_IN   影像检查记录.英文名%Type,
  性别_IN     影像检查记录.性别%Type,
  年龄_IN     影像检查记录.年龄%Type,
  出生日期_IN 影像检查记录.出生日期%Type,
  身高_IN     影像检查记录.身高%Type,
  体重_IN     影像检查记录.体重%Type,
  病理检查_IN 影像检查记录.病理检查%Type,
  发放胶片_IN 影像检查记录.发放胶片%Type,
  检查设备_IN 影像检查记录.检查设备%Type,
  修改_IN     Number:=0,
  电话_IN     影像检查记录.联系电话%Type:=Null
) IS
  Cursor c_Advice IS
    Select A.医嘱ID
    From 病人医嘱发送 A,病人医嘱记录 B
    Where (B.ID=医嘱ID_IN Or (B.相关ID=医嘱ID_IN And B.诊疗类别 IN('F','G','D')))
      And A.医嘱ID=B.ID And A.发送号+0=发送号_IN;
  iRecCount Number;
Begin
	Select Count(*) Into iRecCount From 影像检查记录 Where 检查号=检查号_IN And 影像类别=影像类别_IN;

	Update 影像检查记录
		Set 影像类别=影像类别_IN,
   			检查号=检查号_IN,
				姓名=姓名_IN,
				英文名=英文名_IN,
				性别=性别_IN,
				年龄=年龄_IN,
				出生日期=出生日期_IN,
				身高=身高_IN,
				体重=体重_IN,
				病理检查=病理检查_IN,
				发放胶片=发放胶片_IN,
				检查设备=检查设备_IN,
				联系电话=电话_IN
	Where 医嘱ID=医嘱ID_IN And 发送号=发送号_IN;
	If SQl%RowCount=0 Then
    Insert Into 影像检查记录(医嘱ID,发送号,影像类别,检查号,姓名,英文名,性别,年龄,出生日期,
      身高,体重,病理检查,发放胶片,检查设备,联系电话)
    Values(医嘱ID_IN,发送号_IN,影像类别_IN,检查号_IN,姓名_IN,英文名_IN,性别_IN,年龄_IN,出生日期_IN,
      身高_IN,体重_IN,病理检查_IN,发放胶片_IN,检查设备_IN,电话_IN);
  End if;
  
  If 修改_IN=0 Then
    For r_Advice In c_Advice Loop
      Update 病人医嘱发送 Set 首次时间=Sysdate,末次时间=Sysdate,执行状态=3,执行过程=2,执行间=执行间_IN Where 医嘱ID=r_Advice.医嘱ID And 发送号=发送号_IN;
    End Loop;
    If iRecCount=0 Then
      Update 影像检查类别 Set 最大号码=检查号_IN Where 编码=影像类别_IN;
    End If;
  Else
    For r_Advice In c_Advice Loop
      Update 病人医嘱发送 Set 执行间=执行间_IN Where 医嘱ID=r_Advice.医嘱ID And 发送号=发送号_IN;
    End Loop;
  End If;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_影像检查_BEGIN;
/
--LIS
CREATE OR REPLACE Procedure ZL_检验标本记录_报告审核(
	ID_IN			检验标本记录.ID%Type
) IS

	--未审核的费用行(不包含药品)
	Cursor c_Verify(v_医嘱id In Number) is	Select Distinct 记录性质,NO,序号 From 病人费用记录
						Where 收费类别 Not IN('5','6','7') And 医嘱序号+0=v_医嘱id
							And 记帐费用=1 And 记录状态=0 And 价格父号 IS NULL
							And (记录性质,NO) IN(
								Select 记录性质,NO From 病人医嘱附费 Where 医嘱ID=v_医嘱id
								Union ALL
								Select 记录性质,NO From 病人医嘱发送 Where 医嘱ID In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID,相关id)))
						Order By 记录性质,NO,序号;

	--提取病人的相关信息
	CURSOR c_Advice2(v_医嘱id In Number) IS	SELECT * FROM 病人医嘱记录 WHERE ID=v_医嘱id;

	r_Advice2 c_Advice2%RowType;

	--查找文件的组成元素
	CURSOR c_File(v_File number) IS	SELECT 类型,编码,文本转储,标题文本,标题显示,标题字体,标题位置,内容字体,内容位置,嵌入方式
					FROM 病历文件组成 A,病历元素目录 B
					where A.病历元素id=B.ID
					      AND A.病历文件id=v_File
					order by A.排列序号;

	--检验项目结果
	CURSOR c_Result(v_医嘱id In Number) IS	Select	B.中文名,B.ID,B.类型,B.单位,A.检验结果,DECODE(A.结果标志,1,'1-正常',2,'2-偏低',3,'3-偏高',4,'4-阳性',5,'5-阴性','') AS 结果标志,A.结果参考,C.ID As 标本id
						From	检验普通结果 A,
							诊治所见项目 B,
							检验标本记录 C,
							检验项目分布 D
						Where	A.检验项目id=B.ID
							AND A.检验标本id=C.ID
							AND A.记录类型=C.报告结果
							AND C.样本状态=2
							AND D.标本id=A.检验标本id
							AND D.项目id=A.检验项目id
							AND D.医嘱id=v_医嘱id;

	--查找当前标本的相关申请
	CURSOR c_SampleQuest(v_微生物 In Number) IS	Select Distinct 医嘱id From (
							Select 医嘱id From 检验项目分布 Where 0=v_微生物 And 标本id=ID_IN
							Union
							Select 医嘱id From 检验标本记录 Where 1=v_微生物 And ID=ID_IN);

	Cursor c_Stuff(v_NO Varchar2,v_主页id Number) is
		Select NO,单据,库房ID From 未发药品记录
		Where NO=v_NO And 单据=25 And 库房ID IS Not Null
			And Not Exists(Select 参数值 From 系统参数表 Where 参数号=Decode(v_主页id,NULL,92,63) And 参数值='1')
			And Exists(
				Select A.序号 From 病人费用记录 A,材料特性 B
				Where A.记录性质=2 And A.记录状态=1 And A.NO=v_NO
					And A.收费细目ID=B.材料ID And B.跟踪在用=1
				)
		Order BY 库房ID;

	v_执行			Number(1);
	v_NO			病人医嘱发送.NO%Type;
	v_性质			病人医嘱发送.记录性质%Type;
	v_序号			Varchar2(1000);

	v_Temp			Varchar2(255);
	v_人员部门ID		部门人员.部门ID%Type;
	 v_人员编号		人员表.编号%Type;
	v_人员姓名		人员表.姓名%Type;
	v_Count			Number(18);


	v_病历内容id		number(18);
	v_病历id		number(18);
	v_报告id		number(18);
	v_文件ID		number(18);
	v_Index			number(18);
	v_病历种类		病历文件目录.种类%Type;
	v_病历名称 		病历文件目录.名称%Type;

	v_FLAG			Number(1);
	v_MaxIndex		Number(18);
	v_微生物标本		Number(1):=0;
	v_主页id	Number(18);
Begin

	--0.操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
	v_Temp:=zl_Identity;
	v_人员部门ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
	v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
	v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
	v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
	v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

	--执行后自动审核对应的记帐划价单(不包含药品)
	v_执行:=0;
	Begin
		Select To_Number(Nvl(参数值,'0')) Into v_执行 From 系统参数表 Where 参数号=81;
	Exception
		When Others Then v_执行:=0;
	End;

	v_微生物标本:=0;
	Begin
		Select 1 Into v_微生物标本 From 检验标本记录 Where 微生物标本=1 And ID=ID_IN;
	Exception
		When Others Then v_微生物标本:=0;
	End;

	--1.置本标本的状态及审核人和时间
	UPDATE 检验标本记录 SET 审核人=v_人员姓名,审核时间=SYSDATE,样本状态=2 WHERE ID=ID_IN;

	--2.检查当前标本相关的申请的相关标本是否完成审核
	For r_SampleQuest In c_SampleQuest(v_微生物标本) Loop

		v_Count:=0;

		If v_微生物标本=0 Then
			Begin
				Select NVL(COUNT(1),0) INTO v_Count  From 检验标本记录 Where 样本状态<2 And ID In (Select 标本id From 检验项目分布 Where 医嘱id=r_SampleQuest.医嘱id);
			Exception
				When Others Then v_Count:=0;
			End;
		End If;

		--r_SampleQuest.医嘱id申请已经完成,处理后续环节
		If v_Count=0 Then

			--1.置申请单的执行状态
			Update 病人医嘱发送 Set 执行状态=1 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_SampleQuest.医嘱id In (ID,相关id));

			--2.费用执行处理
			Update 病人费用记录 Set 执行状态=1,执行时间=Sysdate,执行人=v_人员姓名
			Where 收费类别 Not IN('5','6','7')
				And (医嘱序号,记录性质,NO) IN (
					Select 医嘱ID,记录性质,NO From 病人医嘱附费 Where 医嘱ID=r_SampleQuest.医嘱id
					Union ALL
					Select 医嘱ID,记录性质,NO From 病人医嘱发送 Where 医嘱ID In (Select ID From 病人医嘱记录 Where r_SampleQuest.医嘱id In (ID,相关ID)));

			--3.自动审核记帐
			If v_执行=1 Then
				For r_Verify IN c_Verify(r_SampleQuest.医嘱id) Loop
					IF 	r_Verify.NO||','||r_Verify.记录性质<>v_NO||','||v_性质 Then
						If v_序号 IS Not NULL Then
							If v_性质=1 Then
								zl_门诊记帐记录_Verify(v_NO,v_人员编号,v_人员姓名,Substr(v_序号,2));
							Elsif v_性质=2 Then
								zl_住院记帐记录_Verify(v_NO,v_人员编号,v_人员姓名,Substr(v_序号,2));
							End IF;
						End IF;
						v_序号:=NULL;
					End IF;
					v_NO:=r_Verify.NO;
					v_性质:=r_Verify.记录性质;
					v_序号:=v_序号||','||r_Verify.序号;
				End Loop;
				If v_序号 IS Not NULL Then
					If v_性质=1 Then
						zl_门诊记帐记录_Verify(v_NO,v_人员编号,v_人员姓名,Substr(v_序号,2));
					Elsif v_性质=2 Then
						zl_住院记帐记录_Verify(v_NO,v_人员编号,v_人员姓名,Substr(v_序号,2));
					End IF;
				End IF;
			End IF;

			--审核试剂消耗单
			v_No:=NextNo(14);

			Update 检验试剂记录 Set No=v_No Where 医嘱id=r_SampleQuest.医嘱id;

			If v_No Is Not Null Then

				ZL_检验试剂记录_BILL(r_SampleQuest.医嘱id,v_No);

				v_主页id:=Null;
				Select 主页id Into v_主页id From 病人医嘱记录 A WHERE ID=r_SampleQuest.医嘱id;

				If v_主页id Is Null Then
					zl_门诊记帐记录_Verify(v_No,v_人员编号,v_人员姓名);
				Else
					zl_住院记帐记录_Verify(v_No,v_人员编号,v_人员姓名);
				End If;

				--如果记帐没有自动发料,则自动发料,否则不处理
				For r_Stuff In c_Stuff(v_No,v_主页id) Loop
					zl_材料收发记录_处方发料(r_Stuff.库房ID,25,v_No,v_人员姓名,v_人员姓名,v_人员姓名,1,Sysdate);
				End Loop;

			End If;

			--4.自动填写报告,包括普通项目,微生物
			If v_微生物标本>=0 Then
				v_报告id:=0;
				v_病历id:=0;
				begin
					Select Distinct nvl(报告id,0) into v_报告id from 病人医嘱发送  Where 医嘱id in (SELECT ID FROM 病人医嘱记录 WHERE r_SampleQuest.医嘱id In (ID,相关id));
				exception
					when others then v_报告id:=0;
				end;

				If Nvl(v_报告id,0)=0 then

					--产生病人病历记录
					Open c_Advice2(r_SampleQuest.医嘱id);
					Fetch c_Advice2 Into r_Advice2;

					--检查要填写的报告格式中是否含有检验专用纸,如果没有则返回
					v_文件ID:=0;
					begin
						Select U.ID,U.种类,U.名称
						Into v_文件ID,v_病历种类,v_病历名称
						From 病历文件组成 X,病历元素目录 Y,病历文件目录 U
						Where X.病历文件id in (select A.病历文件id
									from 诊疗单据应用 A,病人医嘱记录 B
									where A.诊疗项目id=B.诊疗项目id
										and B.相关ID=r_SampleQuest.医嘱id
										and A.应用场合=r_Advice2.病人来源)
							AND X.病历元素id=Y.ID
							AND U.ID=X.病历文件id
							AND Y.类型=4 and Y.编码='000009';
					exception
						when others then v_文件ID:=0;
					end;

					--有LIS专用纸,要填写
					If nvl(v_文件ID,0)>0 then

						--新产生报告id
						v_病历id:=0;
						Select 病人病历记录_ID.Nextval Into v_报告id From Dual;

						ZL_病人病历_INSERT(v_报告id,r_Advice2.病人id,r_Advice2.主页id,r_Advice2.挂号单,r_Advice2.婴儿,r_Advice2.病人科室id,v_病历种类,v_文件ID,v_病历名称,v_人员姓名,r_SampleQuest.医嘱id);

						--按病历组成依次产生病历组成元素记录
						v_Index:=0;
						FOR r_File In c_File(v_文件ID) LOOP
							v_Index:=v_Index+1;

							Select 病人病历内容_ID.Nextval Into v_病历内容id From Dual;

							if r_File.类型=4 and r_File.编码='000009' then
								v_病历id:=v_病历内容id;
							end if;

							ZL_病人病历内容_INSERT(v_病历内容id,NULL,v_报告id,v_Index,r_File.类型,r_File.编码,r_File.文本转储,r_File.标题文本,r_File.标题显示,r_File.标题字体,r_File.标题位置,0,r_File.内容字体,r_File.内容位置,0,r_File.嵌入方式);
						END LOOP;

					End if;
					Close c_Advice2;
				Else

					--检查要已填写的报告格式中是否含有检验专用纸,并找出检验专用纸在报告中的位置,如果没有则返回
					v_病历id:=0;
					begin
						select nvl(id,0) into v_病历id from 病人病历内容 where 元素类型=4 and 元素编码='000009' and 病历记录id=v_报告id;
					exception
						when others then v_病历id:=0;
					End;
				End If;

				--有Lis专用纸病历,则填写检验结果到此专用纸中
				If v_病历id>0 And v_报告id>0 Then

					v_Index:=0;
          Delete from 病人病历所见单 where 所见项id+0 In
            (Select	D.项目id From	检验标本记录 C,检验项目分布 D
              Where	D.标本id=C.ID AND C.样本状态=2 AND D.医嘱id=r_SampleQuest.医嘱id)
            And 病历id In (Select ID From 病人病历内容 Where 元素类型=4 and 元素编码='000009' and 病历记录id=v_报告id);
					FOR r_Result In c_Result(r_SampleQuest.医嘱id) LOOP

						--Delete from 病人病历所见单 where 所见项id=r_Result.ID and 病历id In (Select ID From 病人病历内容 Where 元素类型=4 and 元素编码='000009' and 病历记录id=v_报告id);

						v_MaxIndex:=1;
						Begin
							Select Nvl(Max(A.控件号),0)+1 Into v_MaxIndex From 病人病历所见单 A,病人病历内容 B Where A.病历id=B.ID AND B.元素类型=4 and B.元素编码='000009' and B.病历记录id=v_报告id;
						Exception
							When Others Then v_MaxIndex:=1;
						End;

						Insert Into 病人病历所见单(病历ID,控件号,控件类,标题,所见项ID,数值类型,计量单位,所见内容)
						Select ID,v_MaxIndex,2,r_Result.中文名,r_Result.ID,r_Result.类型,r_Result.单位,r_Result.检验结果||''''||r_Result.结果标志||''''||r_Result.结果参考
						From 病人病历内容 Where 元素类型=4 and 元素编码='000009' and 病历记录id=v_报告id;

					END LOOP;
				End If;
				--修改病人医嘱发送记录的报告id列
				IF v_报告id>0 THEN
					Update 病人医嘱发送 SET 报告id=v_报告id WHERE 医嘱id in (SELECT ID FROM 病人医嘱记录 WHERE r_SampleQuest.医嘱id In (ID,相关id));
				END IF;
			End If;
		End If;
	End Loop;
Exception
	When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_检验标本记录_报告审核;
/

-------------------------------------------------------------------------
--功能说明： 自动计算一个病人的费用
--入口参数：
--       PatiID  number    病人身份ID
--       PageID  number    病案主页ID，两个参数共同确定需要计算的病人
--       ReCalcBDate  Date 重算开始时间
--调用关系： 外部应用程序调用本过程
-------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE Zl1_autocptpati (
    PatiID IN NUMBER,
    PageID IN NUMBER,
    ReCalcBDate IN 病人变动记录.上次计算时间%Type:=Null
)
AS
    Modilast NUMBER (1);--是否修正上期自动计费参数
    Period VARCHAR2 (6);--需要计算的最小期间
BEGIN
    SELECT 期间
      INTO Period
      FROM 期间表
     WHERE TRUNC (SYSDATE) BETWEEN TRUNC (开始日期) AND TRUNC (终止日期);
    SELECT 参数值
      INTO Modilast
      FROM 系统参数表
     WHERE 参数号 = 7;

    IF Modilast = 1 THEN
        Period :=
          TO_CHAR (ADD_MONTHS (TO_DATE (Period || '05', 'yyyymmdd'), -1),
              'yyyymm');
    END IF;
    
    IF ReCalcBDate IS NOT NULL THEN 
        Update 病人变动记录 Set 上次计算时间=Null Where 病人ID=PatiID And 主页ID=PageID And 上次计算时间>=ReCalcBDate;
        COMMIT;
    END IF; 

    Zl1_autocptone (PatiID, PageID, Period);
    COMMIT;
END;
/
-------------------------------------------------------------------------
--功能说明： 自动计算一个病区所有病人费用
--入口参数： WardID    number  病区ID,ReCalcBDate  Date 重算开始时间
--调用关系： 外部应用程序调用本过程
-------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE Zl1_autocptward (
    WardID IN NUMBER,
    ReCalcBDate IN 病人变动记录.上次计算时间%Type:=Null
)
AS
    Modilast NUMBER (1);--是否修正上期自动计费参数
    Period VARCHAR2 (6);--需要计算的最小期间

    CURSOR Patitab
    IS
        SELECT Distinct 病人ID, 主页ID
          FROM 病人自动费用
         WHERE 病区ID = WardID
            AND TRUNC (终止日期) >=
                                (SELECT MIN (开始日期)
                                    FROM 期间表
                                  WHERE 期间 >= Period);
BEGIN
    SELECT 期间
      INTO Period
      FROM 期间表
     WHERE TRUNC (SYSDATE) - 1 BETWEEN TRUNC (开始日期) AND TRUNC (终止日期);
    SELECT 参数值
      INTO Modilast
      FROM 系统参数表
     WHERE 参数号 = 7;

    IF Modilast = 1 THEN
        Period :=
          TO_CHAR (ADD_MONTHS (TO_DATE (Period || '05', 'yyyymmdd'), -1),
              'yyyymm');
    END IF;
    
    IF ReCalcBDate IS NOT NULL THEN 
        Update 病人变动记录 Set 上次计算时间=Null 
        Where (病人ID,主页ID) IN (Select 病人ID,主页ID From 病案主页 Where 当前病区ID=WardID And 出院日期 Is Null)  
                    And 上次计算时间>=ReCalcBDate;
        COMMIT;
    END IF; 

    FOR Patifld IN Patitab LOOP
        IF      Patifld.病人ID IS NOT NULL
            AND Patifld.主页ID IS NOT NULL THEN
            Zl1_autocptone (Patifld.病人ID, Patifld.主页ID, Period);
            COMMIT;
        END IF;
    END LOOP;
END;
/

--朱玉宝 2005-12-16
--用于门诊就诊、入院
CREATE OR REPLACE PROCEDURE zl_就诊登记记录_UPDATE(
	险类_IN				就诊登记记录.险类%TYPE,
	病人ID_IN			就诊登记记录.病人ID%TYPE,
	主页ID_IN			就诊登记记录.主页ID%TYPE,
	就诊时间_IN			就诊登记记录.就诊时间%TYPE,
	状态_IN				就诊登记记录.状态%TYPE:=0,
	医疗类别_IN			就诊登记记录.医疗类别%TYPE:=NULL,
	帐户余额_IN			就诊登记记录.帐户余额%TYPE:=0,
	病种ID_IN			就诊登记记录.病种ID%TYPE:=NULL,
	病种名称_IN			就诊登记记录.病种名称%TYPE:=NULL,
	并发症_IN			就诊登记记录.并发症%TYPE:=NULL,
	IC卡信息_IN			就诊登记记录.IC卡信息%TYPE:=NULL,
	HIS流水号_IN		就诊登记记录.HIS流水号%TYPE:=NULL,
	YB流水号_IN			就诊登记记录.YB流水号%TYPE:=NULL,
	备注_IN				就诊登记记录.备注%TYPE:=NULL
)
AS 
BEGIN 
	UPDATE 就诊登记记录
	SET 状态=状态_IN,
		医疗类别=医疗类别_IN,
		帐户余额=帐户余额_IN,
		病种ID=病种ID_IN,
		病种名称=病种名称_IN,
		并发症=并发症_IN,
		IC卡信息=IC卡信息_IN,
		HIS流水号=HIS流水号_IN,
		YB流水号=YB流水号_IN,
		备注=备注_IN
	WHERE 险类=险类_IN AND 病人ID=病人ID_IN 
	AND Nvl(主页ID,0)=Nvl(主页ID_IN,0) AND 就诊时间=就诊时间_IN;

	IF SQL%ROWCOUNT =0 THEN 
		INSERT INTO 就诊登记记录
		(险类,病人ID,主页ID,就诊时间,状态,医疗类别,帐户余额,
		病种ID,病种名称,并发症,IC卡信息,HIS流水号,YB流水号,备注)
		VALUES 
		(险类_IN,病人ID_IN,Nvl(主页ID_IN,0),就诊时间_IN,状态_IN,医疗类别_IN,帐户余额_IN,
		病种ID_IN,病种名称_IN,并发症_IN,IC卡信息_IN,HIS流水号_IN,YB流水号_IN,备注_IN);
	END IF ;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_就诊登记记录_UPDATE;
/


--仅用于门诊取消就诊登记、撤销入院，无费出院时将状态更新为零
CREATE OR REPLACE PROCEDURE zl_就诊登记记录_DELETE(
	险类_IN				就诊登记记录.险类%TYPE,
	病人ID_IN			就诊登记记录.病人ID%TYPE,
	主页ID_IN			就诊登记记录.主页ID%TYPE,
	就诊时间_IN			就诊登记记录.就诊时间%TYPE
)
AS 
BEGIN 
	DELETE  就诊登记记录
	WHERE 险类=险类_IN AND 病人ID=病人ID_IN 
	AND NVL(主页ID,0)=NVL(主页ID_IN,0) AND 就诊时间=就诊时间_IN;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_就诊登记记录_DELETE;
/


--此过程仅用于入出院过程中同步更新就诊登记记录的状态
CREATE OR REPLACE PROCEDURE zl_就诊登记记录_更新状态(
	险类_IN				就诊登记记录.险类%TYPE,
	病人ID_IN			就诊登记记录.病人ID%TYPE,
	主页ID_IN			就诊登记记录.主页ID%TYPE,
	状态_IN 			就诊登记记录.状态%TYPE:=0
)
AS 
BEGIN 
	UPDATE 就诊登记记录
	SET 状态=状态_IN
	WHERE 险类=险类_IN AND 病人ID=病人ID_IN AND Nvl(主页ID,0)=Nvl(主页ID_IN,0) ;

EXCEPTION 
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_就诊登记记录_更新状态;
/



--此过程仅用于门诊
CREATE OR REPLACE PROCEDURE zl_就诊登记记录_结束(
	险类_IN				就诊登记记录.险类%TYPE,
	病人ID_IN			就诊登记记录.病人ID%TYPE,
	主页ID_IN			就诊登记记录.主页ID%TYPE,
	就诊时间_IN			就诊登记记录.就诊时间%TYPE,
	记录ID_IN			就诊登记记录.记录ID%TYPE
)
AS 
BEGIN 
	UPDATE 就诊登记记录
	SET 记录ID=记录ID_IN,
		状态=0
	WHERE 险类=险类_IN AND 病人ID=病人ID_IN AND Nvl(主页ID,0)=Nvl(主页ID_IN,0) AND 就诊时间=就诊时间_IN;

EXCEPTION 
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_就诊登记记录_结束;
/



Create OR Replace Procedure zl_保险结算记录_Insert(
    性质_IN             保险结算记录.性质%Type,
    记录ID_IN           保险结算记录.记录ID%Type,
    险类_IN             保险结算记录.险类%Type,
    病人ID_IN           保险结算记录.病人ID%Type,
    年度_IN             保险结算记录.年度%Type,
    帐户累计增加_IN     保险结算记录.帐户累计增加%Type,
    帐户累计支出_IN     保险结算记录.帐户累计支出%Type,
    累计进入统筹_IN     保险结算记录.累计进入统筹%Type,
    累计统筹报销_IN     保险结算记录.累计统筹报销%Type,
    住院次数_IN         保险结算记录.住院次数%Type,
    起付线_IN           保险结算记录.起付线%Type,
    封顶线_IN           保险结算记录.封顶线%Type,
    实际起付线_IN       保险结算记录.实际起付线%Type,
    发生费用金额_IN     保险结算记录.发生费用金额%Type,
    全自付金额_IN       保险结算记录.全自付金额%Type,
    首先自付金额_IN     保险结算记录.首先自付金额%Type,
    进入统筹金额_IN     保险结算记录.进入统筹金额%Type,
    统筹报销金额_IN     保险结算记录.统筹报销金额%Type,
    大病自付金额_IN     保险结算记录.大病自付金额%Type,
    超限自付金额_IN     保险结算记录.超限自付金额%Type,
    个人帐户支付_IN     保险结算记录.个人帐户支付%Type,
    支付顺序号_IN       保险结算记录.支付顺序号%Type,
    主页ID_IN           保险结算记录.主页ID%Type := null,
    中途结帐_IN         保险结算记录.中途结帐%Type := null,
    备注_IN             保险结算记录.备注%Type := null,
	校正_IN             保险结算记录.校正%TYPE:=0,
	工作站_IN			保险结算记录.工作站%TYPE:=NULL,
	版本号_IN			保险结算记录.版本号%TYPE:=NULL,
	医疗类别_IN			保险结算记录.医疗类别%TYPE:=NULL,
	就诊流水号_IN		保险结算记录.就诊流水号%TYPE:=NULL,
	病种ID_IN			保险结算记录.病种ID%TYPE:=NULL,
	病种名称_IN			保险结算记录.病种名称%TYPE:=NULL,
	并发症_IN			保险结算记录.并发症%TYPE:=NULL,
	结算时间_IN			保险结算记录.结算时间%TYPE:=SYSDATE
)
AS
BEGIN
    Update 保险结算记录
        Set 性质=性质_IN,
            记录ID=记录ID_IN,
            险类=险类_IN,
            病人ID=病人ID_IN,
            年度=年度_IN,
            帐户累计增加=帐户累计增加_IN,
            帐户累计支出=帐户累计支出_IN,
            累计进入统筹=累计进入统筹_IN,
            累计统筹报销=累计统筹报销_IN,
            住院次数=住院次数_IN,
            起付线=起付线_IN,
            封顶线=封顶线_IN,
            实际起付线=实际起付线_IN,
            发生费用金额=发生费用金额_IN,
            全自付金额=全自付金额_IN,
            首先自付金额=首先自付金额_IN,
            进入统筹金额=进入统筹金额_IN,
            统筹报销金额=统筹报销金额_IN,
            大病自付金额=大病自付金额_IN,
            超限自付金额=超限自付金额_IN,
            个人帐户支付=个人帐户支付_IN,
            支付顺序号=支付顺序号_IN,
            主页ID=nvl(主页ID_IN,主页ID),
            中途结帐=nvl(中途结帐_IN,中途结帐),
            备注=nvl(备注_IN,备注),
			校正=校正_IN,
			版本号=版本号_IN,
			医疗类别=医疗类别_IN,
			病种ID=病种ID_IN,
			病种名称=病种名称_IN,
			并发症=并发症_IN,
			结算时间=结算时间_IN ,
			工作站=工作站_IN,
			就诊流水号=就诊流水号_IN
    Where 记录ID=记录ID_IN And 性质=性质_IN;

    IF SQl%RowCount=0 Then
        Insert Into 保险结算记录(
            性质,记录id,险类,病人id,年度,帐户累计增加,帐户累计支出,累计进入统筹,累计统筹报销,住院次数,
            起付线,封顶线,实际起付线,发生费用金额,全自付金额,首先自付金额,进入统筹金额,统筹报销金额,
            大病自付金额,超限自付金额,个人帐户支付,支付顺序号,主页ID,中途结帐,备注,校正,
			工作站,版本号,医疗类别,就诊流水号,病种ID,病种名称,并发症,结算时间)
        Values(
            性质_IN,记录id_IN,险类_IN,病人id_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,
            累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
            进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN,校正_IN,
			工作站_IN,版本号_IN,医疗类别_IN,就诊流水号_IN,病种ID_IN,病种名称_IN,并发症_IN,结算时间_IN);
    End if;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_保险结算记录_Insert;
/


-------------------------------------------------------
--模块：保险项目管理
--功能：将对照信息保存到医保对照明细中，为了保持与以前兼容，将缺省类别下的对照信息同步保存到保险支付项目中
----------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_保险支付项目_Modify(
    收费细目ID_IN	IN 保险支付项目.收费细目ID%TYPE,
    险类_IN         IN 保险支付项目.险类%TYPE,
    大类ID_IN		IN 保险支付项目.大类ID%TYPE,
    项目编码_IN		IN 保险支付项目.项目编码%TYPE,
    项目名称_IN		IN 保险支付项目.项目名称%TYPE,
    附注_IN         IN 保险支付项目.附注%TYPE,
    是否医保_IN     IN 保险支付项目.是否医保%TYPE,
	类别_IN			IN NUMBER := 0
)
IS 
BEGIN 
	--首先更新医保对照明细数据，没有则插入
	IF 项目编码_IN IS NOT NULL THEN 
		DELETE 医保对照明细
		WHERE 收费细目ID=收费细目ID_IN AND 险类=险类_IN AND 类别=类别_IN AND 项目编码=项目编码_IN;

		INSERT INTO 医保对照明细(险类,类别,收费细目ID,项目编码)
		VALUES (险类_IN,类别_IN,收费细目ID_IN,项目编码_IN);
	END IF ;

	--如果类别_IN=0则表示是缺省类别，为了保持与以前的模式兼容，直接更新保险支付项目中的数据
	IF 类别_IN =0 THEN 
		--进行修改
		UPDATE 保险支付项目
		SET 大类ID=大类ID_IN,项目编码=项目编码_IN,项目名称=项目名称_IN,附注=附注_IN,是否医保=是否医保_IN
		WHERE 收费细目ID=收费细目ID_IN AND 险类=险类_IN;
		
		IF SQL%NOTFOUND THEN 
			--不存在，改为新增
			INSERT INTO 保险支付项目(收费细目ID,险类,大类ID,项目编码,项目名称,附注,是否医保)
			VALUES (收费细目ID_IN,险类_IN,大类ID_IN,项目编码_IN,项目名称_IN,附注_IN,是否医保_IN);
		END IF;
	END IF ;
EXCEPTION 
    WHEN OTHERS THEN 
        zl_ErrorCenter (SQLCODE, SQLERRM); 
END ZL_保险支付项目_Modify;
/



Create or Replace View 病人自动费用 as
Select p.病人id, p.主页id, i.姓名, i.性别, i.年龄, i.住院号, a.费别, p.科室id, p.病区id, p.床号, p.附加床位, p.收费细目id,
       p.收入项目id, 1 As 标志, p.现价 As 标准单价, p.开始日期, p.终止日期, p.终止日期 - p.开始日期 As 天数, p.数量, p.经治医师,
       p.责任护士, p.操作员编号, p.操作员姓名
From 病人信息 i, 病案主页 a,
     (Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士,
              b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Sysdate), Sysdate), p.执行日期,
                                     Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量
       From 自动计价项目 a,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 床位等级id, 1 As 数量, 责任护士, 经治医师, 终止时间,
                     操作员编号, 操作员姓名, 上次计算时间
              From 病人变动记录
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 责任护士,
                     经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间
              From 病人变动记录 b, 收费从属项目 i
              Where b.床位等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) b, 收费价目 p
       Where a.病区id = b.病区id And Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And
             p.现价 <> 0 And a.计算标志 = 1 And b.床位等级id = p.收费细目id And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))
       Union All
       Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士,
              b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Sysdate), Sysdate), p.执行日期,
                                     Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量
       From 自动计价项目 a,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 护理等级id, 1 As 数量, 责任护士, 经治医师, 终止时间,
                     操作员编号, 操作员姓名, 上次计算时间
              From 病人变动记录
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 护理等级id, i.从项数次 As 数量, 责任护士,
                     经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间
              From 病人变动记录 b, 收费从属项目 i
              Where b.护理等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) b, 收费价目 p
       Where a.病区id = b.病区id And b.附加床位 <> 1 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.计算标志 = 2 And b.护理等级id = p.收费细目id And Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))
       Union All
       Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士,
              b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Sysdate), Sysdate), p.执行日期,
                                     Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, a.数量
       From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期
              From 自动计价项目
              Union All
              Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期
              From 自动计价项目 a, 收费从属项目 i
              Where a.收费细目id = i.主项id And i.固有从属 > 0) a, 病人变动记录 b, 收费价目 p
       Where a.病区id = b.病区id And b.附加床位 <> 1 And b.开始原因 <> 10 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.收费细目id = p.收费细目id And (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志 = 7) And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))) p
Where i.病人id = p.病人id And a.病人id = p.病人id And a.主页id = p.主页id;

-------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_费别明细_UPDATE (
    费别_IN IN 费别明细.费别%TYPE,
    收入项目ID_IN IN 费别明细.收入项目ID%TYPE,
    比率_IN IN VARCHAR2,
    计算方法_IN In Number:=0
)
IS
-----------------------------------------------------------
--比率_IN参数的填写方式如下：  "1:0:34:100;2:35:101:50;3:102:100000:100;
-----------------------------------------------------------
    V_段号 费别明细.段号%TYPE;
    V_应收段首值 费别明细.应收段首值%TYPE;
    V_应收段尾值 费别明细.应收段尾值%TYPE;
    V_实收比率 费别明细.实收比率%TYPE;
    Intpos PLS_INTEGER;
    Str比率 VARCHAR2 (2000);
BEGIN
    --先删除已有的
    Delete  FROM 费别明细 WHERE 费别 = 费别_IN AND 收入项目ID = 收入项目ID_IN;

    --插入新的分段
    Str比率 := 比率_IN;

    WHILE Str比率 IS NOT NULL LOOP
        Intpos := INSTR (Str比率, ':');

        IF Intpos = 0 THEN
            Str比率 := '';
        ELSE
            --得到段号
            V_段号 := TO_NUMBER (SUBSTR (Str比率, 1, Intpos - 1));
            Str比率 := SUBSTR (Str比率, Intpos + 1);
            --得到应收段首值
            Intpos := INSTR (Str比率, ':');
            V_应收段首值 := TO_NUMBER (SUBSTR (Str比率, 1, Intpos - 1));
            Str比率 := SUBSTR (Str比率, Intpos + 1);
            --得到应收段尾值
            Intpos := INSTR (Str比率, ':');
            V_应收段尾值 := TO_NUMBER (SUBSTR (Str比率, 1, Intpos - 1));
            Str比率 := SUBSTR (Str比率, Intpos + 1);
            --得到实收比率
            Intpos := INSTR (Str比率, ';');
            V_实收比率 := TO_NUMBER (SUBSTR (Str比率, 1, Intpos - 1));
            Str比率 := SUBSTR (Str比率, Intpos + 1);

            Insert INTO 费别明细(费别,收入项目ID,计算方法,段号,应收段首值,应收段尾值,实收比率)
                   VALUES (费别_IN,收入项目ID_IN,计算方法_IN,V_段号,V_应收段首值,V_应收段尾值,V_实收比率);
        END IF;
    END LOOP;
END zl_费别明细_UPDATE;
/

CREATE OR REPLACE PROCEDURE zl_病人信息_Insert (
    病人ID_IN		病人信息.病人ID%TYPE,
    门诊号_IN       病人信息.门诊号%TYPE,
    费别_IN         病人信息.费别%TYPE,
    医疗付款_IN     病人信息.医疗付款方式%TYPE,
    姓名_IN         病人信息.姓名%TYPE,
    性别_IN         病人信息.性别%TYPE,
    年龄_IN         病人信息.年龄%TYPE,
    出生日期_IN     病人信息.出生日期%TYPE,
    出生地点_IN     病人信息.出生地点%TYPE,
    身份证号_IN     病人信息.身份证号%TYPE,
    身份_IN         病人信息.身份%TYPE,
    职业_IN         病人信息.职业%TYPE,
    民族_IN         病人信息.民族%TYPE,
    国籍_IN         病人信息.国籍%TYPE,
    学历_IN         病人信息.学历%TYPE,
    婚姻_IN         病人信息.婚姻状况%TYPE,
    家庭地址_IN     病人信息.家庭地址%TYPE,
    家庭电话_IN     病人信息.家庭电话%TYPE,
    户口邮编_IN     病人信息.户口邮编%TYPE,
    联系人姓名_IN   病人信息.联系人姓名%TYPE,
    联系人关系_IN   病人信息.联系人关系%TYPE,
    联系人地址_IN   病人信息.联系人地址%TYPE,
    联系人电话_IN   病人信息.联系人电话%TYPE,
    合同单位ID_IN   病人信息.合同单位ID%TYPE,
    工作单位_IN     病人信息.工作单位%TYPE,
    单位电话_IN     病人信息.单位电话%TYPE,
    单位邮编_IN     病人信息.单位邮编%TYPE,
    单位开户行_IN   病人信息.单位开户行%TYPE,
    单位帐号_IN     病人信息.单位帐号%TYPE,
    担保人_IN       病人信息.担保人%TYPE,
    担保额_IN       病人信息.担保额%TYPE,
    险类_IN         病人信息.险类%TYPE,
    登记时间_IN     病人信息.登记时间%TYPE,
	区域_IN			病人信息.区域%Type:=NULL,
	担保性质_IN		病人信息.担保性质%Type:=NULL,
    操作员编号_IN   病人担保记录.操作员编号%Type:=NULL,
    操作员姓名_IN   病人担保记录.操作员姓名%Type:=NULL
)
AS
BEGIN
    Insert INTO 病人信息 (
        病人ID,门诊号,费别,医疗付款方式,姓名,性别,年龄,出生日期,出生地点,
        身份证号,身份,职业,民族,国籍,区域,学历,婚姻状况,家庭地址,家庭电话,
        户口邮编,联系人姓名,联系人关系,联系人地址,联系人电话,合同单位ID,
        工作单位,单位电话,单位邮编,单位开户行,单位帐号,担保人,担保额,担保性质,险类,登记时间)
    VALUES (
        病人ID_IN,门诊号_IN,费别_IN,医疗付款_IN,姓名_IN,性别_IN,年龄_IN,出生日期_IN,
        出生地点_IN,身份证号_IN,身份_IN,职业_IN,民族_IN,国籍_IN,区域_IN,学历_IN,婚姻_IN,
        家庭地址_IN,家庭电话_IN,户口邮编_IN,联系人姓名_IN,联系人关系_IN,联系人地址_IN,
        联系人电话_IN,DECODE (合同单位ID_IN, 0, NULL, 合同单位ID_IN),工作单位_IN,
        单位电话_IN,单位邮编_IN,单位开户行_IN,单位帐号_IN,担保人_IN,
		DECODE (担保额_IN,0,NULL,担保额_IN),担保性质_IN,险类_IN,登记时间_IN);
    
    IF 门诊号_IN is Not NULL Then
        Insert Into 门诊病案记录(
            病人ID,病案号,建立日期,病案类别,存储状态,存放位置)
        Values(
            病人ID_IN,门诊号_IN,登记时间_IN,'一般','正常',NULL);
    End IF;
    
    If 担保人_IN Is Not Null Then 
      Insert Into 病人担保记录(病人Id,担保人,担保额,担保性质,操作员编号,操作员姓名,发生时间)
         Values(病人ID_IN,担保人_IN,担保额_IN,担保性质_IN,操作员编号_IN,操作员姓名_IN,登记时间_IN);
    End If;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病人信息_Insert;
/

CREATE OR REPLACE PROCEDURE zl_病人信息_UPDATE (
    病人ID_IN			病人信息.病人ID%TYPE,
    门诊号_IN           病人信息.门诊号%TYPE,
    住院号_IN           病人信息.住院号%TYPE,
    费别_IN             病人信息.费别%TYPE,
    医疗付款_IN         病人信息.医疗付款方式%TYPE,
    姓名_IN             病人信息.姓名%TYPE,
    性别_IN             病人信息.性别%TYPE,
    年龄_IN             病人信息.年龄%TYPE,
    出生日期_IN         病人信息.出生日期%TYPE,
    出生地点_IN         病人信息.出生地点%TYPE,
    身份证号_IN         病人信息.身份证号%TYPE,
    身份_IN             病人信息.身份%TYPE,
    职业_IN             病人信息.职业%TYPE,
    民族_IN             病人信息.民族%TYPE,
    国籍_IN             病人信息.国籍%TYPE,
    学历_IN             病人信息.学历%TYPE,
    婚姻_IN             病人信息.婚姻状况%TYPE,
    家庭地址_IN         病人信息.家庭地址%TYPE,
    家庭电话_IN         病人信息.家庭电话%TYPE,
    户口邮编_IN         病人信息.户口邮编%TYPE,
    联系人姓名_IN       病人信息.联系人姓名%TYPE,
    联系人关系_IN       病人信息.联系人关系%TYPE,
    联系人地址_IN       病人信息.联系人地址%TYPE,
    联系人电话_IN       病人信息.联系人电话%TYPE,
    合同单位ID_IN       病人信息.合同单位ID%TYPE,
    工作单位_IN         病人信息.工作单位%TYPE,
    单位电话_IN         病人信息.单位电话%TYPE,
    单位邮编_IN         病人信息.单位邮编%TYPE,
    单位开户行_IN       病人信息.单位开户行%TYPE,
    单位帐号_IN         病人信息.单位帐号%TYPE,
    担保人_IN           病人信息.担保人%TYPE,
    担保额_IN           病人信息.担保额%TYPE,
    险类_IN             病人信息.险类%TYPE,
    住院费别_IN         Number:=0,--是否修改的是病人的住院费别
	医保号_IN			保险帐户.医保号%Type:=NULL,
	区域_IN				病人信息.区域%Type:=NULL,
	担保性质_IN			病人信息.担保性质%Type:=NULL,
    操作员编号_IN       病人担保记录.操作员编号%Type:=NULL,
    操作员姓名_IN       病人担保记录.操作员姓名%Type:=NULL
)
AS
	v_主页ID	        病案主页.主页ID%Type;
    v_担保人            病人信息.担保人%Type;
    v_担保额            病人信息.担保额%Type;
    v_担保性质          病人信息.担保性质%Type;
Begin    
    Select Nvl(担保人,'无人能及'),Nvl(担保额,0),Nvl(担保性质,0) 
           Into v_担保人,v_担保额,v_担保性质 
    From 病人信息 Where 病人Id=病人ID_IN;
    If 担保人_IN<>v_担保人 Or (担保人_IN Is Null And v_担保人<>'无人能及') Or 担保额_IN<>v_担保额 Or 担保性质_IN<>v_担保性质 Then 
       Insert Into 病人担保记录(病人Id,担保人,担保额,担保性质,操作员编号,操作员姓名,发生时间)
       Values(病人ID_IN,担保人_IN,担保额_IN,担保性质_IN,操作员编号_IN,操作员姓名_IN,Sysdate);
    End If;

    UPDATE 病人信息
        SET 门诊号=门诊号_IN,住院号=住院号_IN,医疗付款方式=医疗付款_IN,费别=Decode(Nvl(住院费别_IN,0),0,费别_IN,费别),
            姓名=姓名_IN,性别=性别_IN,年龄=年龄_IN,出生日期=出生日期_IN,出生地点=出生地点_IN,
            身份证号=身份证号_IN,身份=身份_IN,职业=职业_IN,
            民族=民族_IN,国籍=国籍_IN,区域=区域_IN,学历=学历_IN,
            婚姻状况=婚姻_IN,家庭地址=家庭地址_IN,家庭电话=家庭电话_IN,
            户口邮编=户口邮编_IN,联系人姓名=联系人姓名_IN,联系人关系=联系人关系_IN,
            联系人地址=联系人地址_IN,联系人电话=联系人电话_IN,
            合同单位ID=DECODE (合同单位ID_IN, 0, NULL, 合同单位ID_IN),
            工作单位=工作单位_IN,单位电话=单位电话_IN,单位邮编=单位邮编_IN,
            单位开户行=单位开户行_IN,单位帐号=单位帐号_IN,
            担保人=担保人_IN,担保额=DECODE (担保额_IN,0,NULL,担保额_IN),
            担保性质=担保性质_IN,险类=险类_IN
    WHERE 病人ID=病人ID_IN;
    
    IF 门诊号_IN is Not NULL Then
        Update 门诊病案记录 Set 病案号=门诊号_IN Where 病人ID=病人ID_IN;
        IF SQL%RowCount=0 Then
            Insert Into 门诊病案记录(
                病人ID,病案号,建立日期,病案类别,存储状态,存放位置)
            Values(
                病人ID_IN,门诊号_IN,Sysdate,'一般','正常',NULL);
        End IF;
	Else
		Delete From 门诊病案记录 Where 病人ID=病人ID_IN;
    End IF;

    IF 住院号_IN is Not NULL Then
        Update 住院病案记录 Set 病案号=住院号_IN Where 病人ID=病人ID_IN;
        IF SQL%RowCount=0 Then
            Insert Into 住院病案记录(
                病人ID,病案号,建立日期,病案类别,存储状态,存放位置)
            Values(
                病人ID_IN,住院号_IN,Sysdate,'一般','在院',NULL);
        End IF;
	Else
		Delete From 住院病案记录 Where 病人ID=病人ID_IN;
    End IF;
    
	Begin
		Select Max(主页ID) Into v_主页ID From 病案主页 Where 病人ID=病人ID_IN;
	Exception
		When Others Then NULL;
	End;
	If v_主页ID IS Not NULL Then
		Update 病案主页 
			Set 费别=Decode(Nvl(住院费别_IN,0),1,费别_IN,费别),
				医疗付款方式=医疗付款_IN,
				区域=Decode(区域_IN,NULL,区域,区域_IN)
		Where 病人ID=病人ID_IN And 主页ID=v_主页ID;
		
		If 医保号_IN IS Not NULL Then
			Update 病案主页从表 Set 信息值=医保号_IN Where 病人ID=病人ID_IN And 主页ID=v_主页ID And 信息名='医保号';
			If SQL%RowCount=0 Then
				Insert Into 病案主页从表(
					病人ID,主页ID,信息名,信息值)
				Values(
					病人ID_IN,v_主页ID,'医保号',医保号_IN);
			End IF;
		Else
			Delete From 病案主页从表 Where 病人ID=病人ID_IN And 主页ID=v_主页ID And 信息名='医保号';
		End IF;
	End IF;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病人信息_UPDATE;
/

CREATE OR REPLACE Procedure zl_病人结帐记录_Delete(
    No_IN            病人结帐记录.No%Type, 
    操作员编号_IN    病人结帐记录.操作员编号%Type, 
    操作员姓名_IN    病人结帐记录.操作员姓名%Type, 
    误差金额_IN      病人预交记录.冲预交%Type :=0,        --医保不支持退回时,转现金产生的误差
    误差NO_IN        病人费用记录.No%Type:=Null,          --当结帐时无误差,并且作废时有误差时才有值
    结帐作废结算_IN      Varchar2:=Null                   --结算方式|结算金额|结算号码||......   
) 
AS 
    --该游标用于预交记录相关信息
    Cursor c_Deposit(v_ID 病人预交记录.结帐ID%Type) is 
        Select * From 病人预交记录 Where 结帐ID = v_ID; 
    r_DepositRow c_Deposit%RowType; 
 
    --该游标用于处理费用相关汇总表 
    Cursor c_Money (v_ID 病人预交记录.结帐ID%Type) is 
        Select * From 病人费用记录 Where 结帐ID = v_ID; 
    r_MoneyRow c_Money%RowType; 
    
    --该游标包含误差项目的相关信息
    Cursor c_ErrItem is 
        Select 
            A.类别 AS 收费类别,A.ID AS 收费细目ID,A.计算单位,C.ID AS 收入项目ID,C.收据费目 
        From 收费细目 A,收费价目 B,收入项目 C,收费特定项目 D
        Where D.特定项目='误差项' And D.收费细目ID=A.ID
            And A.ID=B.收费细目ID And B.收入项目ID=C.ID
            And ((Sysdate Between B.执行日期 And B.终止日期) 
                Or (Sysdate>=B.执行日期 And B.终止日期 is NULL));
    r_ErrItem c_ErrItem%RowType;

    --该游标包含病人的相关信息
    Cursor c_Pati(v_病人Id 病人信息.病人ID%Type) is
        Select A.姓名,A.性别,A.年龄,A.住院号,A.门诊号,B.主页ID,B.出院病床,
            B.当前病区ID,B.出院科室ID,Nvl(B.费别,A.费别) AS 费别
        From 病人信息 A,病案主页 B
        Where A.病人ID=v_病人Id And A.病人ID=B.病人ID(+) And Nvl(A.住院次数,0)=B.主页ID(+);
    r_Pati c_Pati%RowType;

    --过程变量
    v_结算内容		Varchar2(500);
    v_当前结算		Varchar2(50);
    v_结算方式		病人预交记录.结算方式%Type;
    v_结算金额		病人预交记录.冲预交%Type;
    v_结算号码		病人预交记录.结算号码%Type;
    
    v_Temp          Varchar2(255);
    v_病人Id        病人信息.病人ID%Type;
    v_人员部门ID    部门人员.部门ID%Type;
 
    v_原ID		      病人结帐记录.ID%Type; 
    v_结帐ID        病人结帐记录.ID%Type; 
    v_打印ID        票据打印内容.ID%Type;    
    v_实际票号      病人预交记录.实际票号%Type; 

    v_误差NO        病人费用记录.NO%Type;     
    v_Date		      Date;  
    Err_Custom      Exception; 
    v_Error         Varchar2 (255); 
Begin 
    Begin 
        Select ID,病人Id,实际票号 Into v_原ID,v_病人Id,v_实际票号 From 病人结帐记录 Where 记录状态 = 1 And No = No_IN; 
        --最后一次打印的内容
        Select Max(ID) Into v_打印ID From 票据打印内容 Where 数据性质=3 And NO=NO_IN;
    Exception 
        When Others Then 
        Begin 
            v_Error := '没有发现要作废的结帐单据,可能已经作废！'; 
            Raise Err_Custom; 
        End; 
    End; 
    Open c_Pati(v_病人Id);
    Fetch c_Pati Into r_Pati;    --体检系统调用此过程,团体结帐时没有病人信息
 
    Select Sysdate Into v_Date From Dual; 
    Select 病人结帐记录_ID.Nextval Into v_结帐ID From Dual;   
 
    --病人结帐记录 
    Insert Into 病人结帐记录( 
        ID,No,实际票号,记录状态,病人ID,操作员编号,操作员姓名,开始日期,结束日期,收费时间) 
    Select 
        v_结帐ID,No,实际票号,2,病人ID,操作员编号_IN,操作员姓名_IN,开始日期,结束日期,v_Date 
    From 病人结帐记录 Where ID = v_原ID; 
 
    Update 病人结帐记录 Set 记录状态=3 Where ID=v_原ID; 
 
    --作废收回票据(可能以前没有使用票据,无法收回)
    IF v_打印ID is NOT Null Then 
        Insert Into 票据使用明细(
            ID,票种,号码,性质,原因,领用ID,打印ID,使用时间,使用人) 
        Select 
            票据使用明细_ID.Nextval,票种,号码,2,2,领用ID,打印ID,v_Date,操作员姓名_IN
        From 票据使用明细
        Where 打印ID=v_打印ID And 票种=3 And 性质=1;
    End IF; 
 
    --病人预交记录(冲预交及缴款)     
    IF 结帐作废结算_IN Is Null Then 
        --非医保结帐作废 
        Insert Into 病人预交记录( 
            ID,No,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,结算方式,结算号码,摘要, 
            缴款单位,单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,冲预交,结帐ID) 
        Select 
            病人预交记录_ID.Nextval,No,实际票号,to_Number('1'||Substr(记录性质,Length(记录性质),1)), 
            记录状态,病人ID,主页ID,科室ID,Null,结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号, 
            收款时间,操作员姓名,操作员编号,-1*冲预交,v_结帐ID 
        From 病人预交记录 
        Where 结帐ID=v_原ID; 
    Else          
        --1.先处理冲预交部分        
        Insert Into 病人预交记录( 
            ID,No,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,结算方式,结算号码,摘要, 
            缴款单位,单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,冲预交,结帐ID) 
        Select 
            病人预交记录_ID.Nextval,No,实际票号,to_Number('1'||Substr(记录性质,Length(记录性质),1)), 
            记录状态,病人ID,主页ID,科室ID,Null,结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号, 
            收款时间,操作员姓名,操作员编号,-1*冲预交,v_结帐ID 
        From 病人预交记录 
        Where 结帐ID=v_原ID And 记录性质 In(1,11); 
 
        --2.再处理结帐结算,包括医保和非医保         
        v_结算内容:=结帐作废结算_IN||' ||';--以空格分开以|结尾,没有结算号码的
        While v_结算内容 IS Not NULL Loop
            v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
            v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
            v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
            v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
            v_结算号码:=LTrim(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
                        
            Insert Into 病人预交记录( 
                ID,No,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,结算方式,结算号码,摘要, 
                缴款单位,单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,冲预交,结帐ID) 
            Values(
                病人预交记录_ID.Nextval,No_IN,v_实际票号,12,1,v_病人ID,r_Pati.主页ID,r_Pati.出院科室ID,Null,v_结算方式,v_结算号码,'结帐作废医保结算退费',
                Null,Null,Null,v_Date,操作员姓名_IN,操作员编号_IN,-1*v_结算金额,v_结帐ID );
                            
            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
    End IF; 
 
    --病人费用记录 
    --读取原结帐时产生的误差费用,并销帐,然后将销帐记录处理在本次结帐作废中    
    Begin
        Select No Into v_误差No	From 病人费用记录 Where 结帐ID=v_原ID And Nvl(附加标志,0)=9 And 记录性质=2 And 记录状态=1;
    Exception
        When Others Then NULL;
    End;
    If v_误差NO IS Not NULL Then
    		--a.结帐时有误差       作废时三种情况:
          --1.原样退(普通收费或医保全部退允回退)  :误差原样退(前面加负号,结帐ID为新的),更新旧结帐ID的误差的记录状态,结帐此误差
          --2.医保只允许部分退,但没有误差,        :不做任何处理
          --3.医保只允许部分退,但有误差,          :以新误差退(前面加负号,结帐ID为新的),更新旧结帐ID的误差的记录状态,结帐此误差             
        If 结帐作废结算_IN Is Null Or 误差金额_IN<>0 Then
            Insert Into 病人费用记录(
                ID,NO,实际票号,记录性质,记录状态,序号,从属父号,价格父号,多病人单,病人ID,主页ID,医嘱序号,
                门诊标志,姓名,性别,年龄,标识号,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,
                发药窗口,数次,加班标志,附加标志,婴儿费,记帐费用,收入项目ID,收据费目,标准单价,应收金额,实收金额,
                划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行状态,操作员姓名,操作员编号,结帐金额,结帐ID,是否上传)
            Select
                病人费用记录_ID.Nextval,No,NULL,记录性质,2,1,从属父号,价格父号,多病人单,病人ID,主页ID,医嘱序号,
                门诊标志,姓名,性别,年龄,标识号,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,
                发药窗口,-1*数次,加班标志,附加标志,婴儿费,记帐费用,收入项目ID,收据费目,Decode(结帐作废结算_IN,Null,应收金额,误差金额_IN),
                Decode(结帐作废结算_IN,Null,-1*应收金额,-1*误差金额_IN),Decode(结帐作废结算_IN,Null,-1*实收金额,-1*误差金额_IN),
                划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行状态,操作员姓名,操作员编号, 0,v_结帐ID,1
            From 病人费用记录
            Where 结帐ID=v_原ID And Nvl(附加标志,0)=9 And 记录性质=2 And 记录状态=1;
            
            --处理旧的误差记录
            Update 病人费用记录
                Set 记录状态=3,执行状态=0
            Where 结帐ID=v_原ID And Nvl(附加标志,0)=9 And 记录性质=2 And 记录状态=1;        
            
            --对新产生的作废误差记录进行结帐
            Update 病人费用记录 Set 结帐金额=实收金额,结帐ID=v_结帐ID,是否上传=1
                Where NO=v_误差NO And Nvl(附加标志,0)=9 And 记录性质=2 And 记录状态=2;  
        End If;      
    Else           
        --b.作废时新产生的误差
        IF 误差金额_IN<>0 Then
            v_Temp:=zl_Identity;
            v_人员部门ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));       
            
            Open c_ErrItem;
            Fetch c_ErrItem Into r_ErrItem;
            If c_ErrItem%RowCount=0 Then
                Close c_ErrItem;
                v_Error:='不能正确读取处理费用误差的项目信息，请先检查该项目是否正确设置。';
                Raise Err_Custom;
            End IF;
            
            Insert Into 病人费用记录(
                ID,NO,实际票号,记录性质,记录状态,序号,从属父号,价格父号,多病人单,病人ID,主页ID,医嘱序号,
                门诊标志,姓名,性别,年龄,标识号,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,
                发药窗口,数次,加班标志,附加标志,婴儿费,记帐费用,收入项目ID,收据费目,标准单价,应收金额,实收金额,
                划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行状态,操作员姓名,操作员编号,
                结帐金额,结帐ID,是否上传)
            Values(
                病人费用记录_ID.Nextval,误差NO_IN,NULL,2,1,1,NULL,NULL,0,v_病人Id,r_Pati.主页ID,NULL,
                Decode(r_Pati.主页ID,NULL,1,2),r_Pati.姓名,r_Pati.性别,r_Pati.年龄,Decode(r_Pati.主页ID,NULL,r_Pati.门诊号,r_Pati.住院号),
                r_Pati.出院病床,Nvl(r_Pati.当前病区ID,v_人员部门ID),Nvl(r_Pati.出院科室ID,v_人员部门ID),r_Pati.费别,
                r_ErrItem.收费类别,r_ErrItem.收费细目ID,r_ErrItem.计算单位,1,NULL,1,NULL,9,0,1,r_ErrItem.收入项目ID,
                r_ErrItem.收据费目,-1*误差金额_IN,-1*误差金额_IN,-1*误差金额_IN,操作员姓名_IN,v_人员部门ID,操作员姓名_IN,v_Date,
                v_Date,v_人员部门ID,0,操作员姓名_IN,操作员编号_IN,-1*误差金额_IN,v_结帐Id,1);   
  
            --结帐此误差       
            Update 病人费用记录 Set 结帐金额=实收金额,结帐ID=v_结帐ID,是否上传=1
            Where NO=误差NO_IN And Nvl(附加标志,0)=9 And 记录性质=2 And 记录状态=1;  
            
            v_误差NO:=误差NO_IN;  --用来在后面排开误差的汇总处理
            Close c_ErrItem;
        End If;  
    End IF;
    
    --作废结帐对应的费用记录:不包含原始结帐产生的误差项目
    Insert Into 病人费用记录( 
        ID,No,实际票号,记录性质,记录状态,序号,从属父号,价格父号,多病人单,记帐单ID,病人ID, 
        主页ID,医嘱序号,门诊标志,姓名,性别,年龄,标识号,床号,病人病区ID,病人科室ID,费别, 
        收费类别,收费细目ID,计算单位,付数,发药窗口,数次,加班标志,附加标志,婴儿费,记帐费用, 
        收入项目ID,收据费目,标准单价,应收金额,实收金额,划价人,开单部门ID,开单人,发生时间, 
        登记时间,执行部门ID,执行状态,执行人,执行时间,操作员姓名,操作员编号,结帐金额,结帐ID,
        保险项目否,保险大类ID,统筹金额,是否急诊,保险编码,摘要) 
    Select 
        病人费用记录_ID.Nextval,No,实际票号,to_Number('1'||Substr(记录性质,Length(记录性质),1)), 
        记录状态,序号,从属父号,价格父号,多病人单,记帐单ID,病人ID,主页ID,医嘱序号,门诊标志,姓名, 
        性别,年龄,标识号,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,发药窗口, 
        数次,加班标志,附加标志,婴儿费,记帐费用,收入项目ID,收据费目,标准单价,Null,Null,划价人, 
        开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行状态,执行人,执行时间,操作员姓名,操作员编号, 
        -1 * 结帐金额, v_结帐ID,
		保险项目否,保险大类ID,统筹金额,是否急诊,保险编码,摘要
    From 病人费用记录 
    Where 结帐ID=v_原ID And Nvl(附加标志,0)<>9;
 
    --相关汇总表处理
    For r_DepositRow in c_Deposit(v_结帐ID) Loop 
        IF r_DepositRow.记录性质 in (1,11) Then 
            --病人余额(预交)     
            Update 病人余额 
                Set 预交余额=Nvl(预交余额,0)-r_DepositRow.冲预交 --注:新的结帐ID产生的是负数金额
             Where 病人ID=r_DepositRow.病人ID And 性质=1; 
 
            IF SQL%RowCount = 0 Then 
                Insert Into 病人余额( 
                    病人ID,性质,预交余额,费用余额) 
                Values( 
                    r_DepositRow.病人ID,1,-1*r_DepositRow.冲预交,0); 
            End IF; 
        Else 
           --人员缴款余额,医保不支持作废的结算方式在新的预交结算中已被处理为了退现金,
            --此处用加,表示收回退给病人的现金(结帐时,退款是负,作废时是正)
            Update 人员缴款余额 
                Set 余额=Nvl(余额,0)+r_DepositRow.冲预交 
             Where 收款员=操作员姓名_IN And 结算方式 = r_DepositRow.结算方式 And 性质 = 1; 

            IF SQL%RowCount = 0 Then 
                Insert Into 人员缴款余额( 
                    收款员,结算方式,性质,余额) 
                Values( 
                    操作员姓名_IN,r_DepositRow.结算方式,1,r_DepositRow.冲预交); 
            End IF; 
            Delete From 人员缴款余额 Where 收款员=操作员姓名_IN And 结算方式=r_DepositRow.结算方式 And 性质=1 And Nvl(余额,0)=0;  
        End IF; 
    End Loop; 
 
    For r_MoneyRow in c_Money(v_结帐ID) Loop 
        --病人余额 ,误差项已结帐,所以不需要更新这两个汇总表
        If Nvl(v_误差NO,'sc')<>Nvl(r_MoneyRow.No,'sc') Then
            Update 病人余额 
                Set 费用余额 = Nvl(费用余额,0)-r_MoneyRow.结帐金额  --注:新的结帐ID产生的是负数金额
             Where 病人ID=r_MoneyRow.病人ID And 性质=1; 
     
            IF SQL%RowCount = 0 Then 
                Insert Into 病人余额( 
                    病人ID,性质,预交余额,费用余额) 
                Values( 
                    r_MoneyRow.病人ID,1,0,-1*r_MoneyRow.结帐金额); 
            End IF; 
     
            --病人未结费用 
            Update 病人未结费用 
                Set 金额 = Nvl(金额,0)-r_MoneyRow.结帐金额 
             Where 病人ID=r_MoneyRow.病人ID 
                And Nvl(主页ID,0)=Nvl(r_MoneyRow.主页ID,0) 
                And Nvl(病人病区ID,0)=Nvl(r_MoneyRow.病人病区ID,0) 
                And Nvl(病人科室ID,0)=Nvl(r_MoneyRow.病人科室ID,0) 
                And Nvl(开单部门ID,0)=Nvl(r_MoneyRow.开单部门ID,0) 
                And Nvl(执行部门ID,0)=Nvl(r_MoneyRow.执行部门ID,0) 
                And 收入项目ID+0=r_MoneyRow.收入项目ID 
                And 来源途径+0=r_MoneyRow.门诊标志; 
     
            IF SQL%RowCount = 0 Then 
                Insert Into 病人未结费用( 
                    病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额) 
                Values( 
                    r_MoneyRow.病人ID,r_MoneyRow.主页ID,r_MoneyRow.病人病区ID,r_MoneyRow.病人科室ID,r_MoneyRow.开单部门ID, 
                    r_MoneyRow.执行部门ID,r_MoneyRow.收入项目ID,r_MoneyRow.门诊标志,-1*r_MoneyRow.结帐金额); 
            End IF; 
        End If;
 
        --病人费用汇总 
        Update 病人费用汇总 
            Set 结帐金额 = Nvl(结帐金额, 0) + r_MoneyRow.结帐金额 
         Where 日期 = Trunc(v_Date) 
            And Nvl(病人病区ID,0) = Nvl(r_MoneyRow.病人病区ID,0) 
            And Nvl(病人科室ID,0) = Nvl(r_MoneyRow.病人科室ID,0) 
            And Nvl(开单部门ID,0) = Nvl(r_MoneyRow.开单部门ID,0) 
            And Nvl(执行部门ID,0) = Nvl(r_MoneyRow.执行部门ID,0) 
            And 收入项目ID+0 = r_MoneyRow.收入项目ID 
            And 来源途径 = r_MoneyRow.门诊标志 
            And 记帐费用 = r_MoneyRow.记帐费用; 
 
        IF SQL%RowCount = 0 Then 
            Insert Into 病人费用汇总( 
                日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额) 
            Values( 
                Trunc (v_Date),r_MoneyRow.病人病区ID,r_MoneyRow.病人科室ID,r_MoneyRow.开单部门ID,r_MoneyRow.执行部门ID, 
                r_MoneyRow.收入项目ID,r_MoneyRow.门诊标志,r_MoneyRow.记帐费用,0,0,r_MoneyRow.结帐金额); 
        End IF; 
    End Loop; 
   
    Close c_Pati;
    
Exception 
    When Err_Custom Then Raise_application_errOr (-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]'); 
    When Others Then zl_ErrOrCenter (SQLCODE, SQLERRM); 
End zl_病人结帐记录_Delete;
/

CREATE OR REPLACE PROCEDURE zl_结帐费用记录_Insert(
    ID_IN           病人费用记录.ID%TYPE,
    NO_IN           病人费用记录.No%TYPE,
    记录性质_IN		病人费用记录.记录性质%TYPE,
    记录状态_IN     病人费用记录.记录状态%TYPE,
    执行状态_IN     病人费用记录.执行状态%TYPE,
    序号_IN         病人费用记录.序号%TYPE,
    结帐金额_IN     病人费用记录.结帐金额%TYPE,
    结帐ID_IN       病人费用记录.结帐ID%TYPE
)
AS
    v_NextID        病人费用记录.ID%TYPE;
    v_病人ID        病人费用记录.病人ID%TYPE;
    v_主页ID        病人费用记录.主页ID%TYPE;
    v_病人病区ID    病人费用记录.病人病区ID%TYPE;
    v_病人科室ID    病人费用记录.病人科室ID%TYPE;
    v_开单部门ID    病人费用记录.开单部门ID%TYPE;
    v_执行部门ID    病人费用记录.执行部门ID%TYPE;
    v_收入项目ID    病人费用记录.收入项目ID%TYPE;
    v_门诊标志      病人费用记录.门诊标志%TYPE;
    v_记帐费用      病人费用记录.记帐费用%TYPE;
    
    v_结帐金额      病人费用记录.结帐金额%Type;
    v_实收金额      病人费用记录.实收金额%Type;

    Err_Custom      Exception;
    v_Error         Varchar2(255);
BEGIN
    IF ID_IN <> 0 THEN
        --第一次结帐
        UPDATE 病人费用记录
            SET 结帐金额=结帐金额_IN,
                 结帐ID=结帐ID_IN
         WHERE ID=ID_IN And 结帐ID IS NULL;

        IF SQL%RowCount=0 Then
            v_Error:='发现已经被其他人结帐的费用,当前结帐操作不能继续。';
            Raise Err_Custom;
        End IF;

        v_NextID:=ID_IN;
    ELSE
        --结以前的余帐
        SELECT 病人费用记录_ID.Nextval INTO v_NextID FROM Dual;

        Insert INTO 病人费用记录(
            ID,No,实际票号,记录性质,记录状态,序号,从属父号,价格父号,多病人单,记帐单ID,
            病人ID,主页ID,医嘱序号,门诊标志,姓名,性别,年龄,标识号,床号,
            病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,发药窗口,数次,
            加班标志,附加标志,婴儿费,记帐费用,收入项目ID,收据费目,标准单价,应收金额,实收金额,
            划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行状态,执行人,执行时间,
            操作员姓名,操作员编号,结帐金额,结帐ID,
			保险项目否,保险大类ID,统筹金额,保险编码,是否急诊,摘要)
        SELECT v_NextID,NO,实际票号,TO_NUMBER('1'||记录性质_IN),记录状态,
             序号,从属父号,价格父号,多病人单,记帐单ID,病人ID,主页ID,
             医嘱序号,门诊标志,姓名,性别,年龄,标识号,床号,
             病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,
             付数,发药窗口,数次,加班标志,附加标志,婴儿费,记帐费用,
             收入项目ID,收据费目,标准单价,NULL,NULL,划价人,
             开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行状态,
             执行人,执行时间,操作员姓名,操作员编号,结帐金额_IN,结帐ID_IN,
			 保险项目否,保险大类ID,统筹金额,保险编码,是否急诊,摘要
         FROM 病人费用记录
         WHERE No=No_IN AND 序号=序号_IN AND 记录状态=记录状态_IN And Nvl(执行状态,0)=Nvl(执行状态_IN,0)
            AND SUBSTR(记录性质,LENGTH(记录性质),1)=记录性质_IN AND ROWNUM=1;

        --检查多次结帐后结帐金额是否高于原金额
        Select Nvl(Sum(实收金额),0),Nvl(Sum(结帐金额),0) Into v_实收金额,v_结帐金额
        From 病人费用记录
        WHERE NO=NO_IN AND 序号=序号_IN AND 记录状态=记录状态_IN
            AND SUBSTR(记录性质,LENGTH(记录性质),1)=记录性质_IN And Nvl(执行状态,0)=执行状态_IN;
        If v_结帐金额>v_实收金额 Then
            v_Error:='发现已经被其他人结帐的费用,当前结帐操作不能继续。';
            Raise Err_Custom;
        End IF;
    END IF;

    Select 
        病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,门诊标志,记帐费用
    Into 
        v_病人ID,v_主页ID,v_病人病区ID,v_病人科室ID,v_开单部门ID,v_执行部门ID,v_收入项目ID,v_门诊标志,v_记帐费用
    From 病人费用记录 Where ID=v_NextID;

    --病人余额
    UPDATE 病人余额 SET 费用余额=NVL(费用余额,0)-结帐金额_IN WHERE 病人ID=v_病人ID AND 性质=1;
    IF SQL%ROWCOUNT=0 THEN
        Insert INTO 病人余额(
            病人ID,性质,预交余额,费用余额) 
        VALUES(
            v_病人ID,1,0,-1 * 结帐金额_IN);
    END IF;
    DELETE FROM 病人余额 WHERE NVL(预交余额,0)=0 AND NVL(费用余额,0)=0 AND 病人ID=v_病人ID;

    --病人未结费用
    UPDATE 病人未结费用
        SET 金额=NVL(金额,0)-结帐金额_IN
    WHERE 病人ID=v_病人ID
        AND NVL(主页ID,0)=NVL(v_主页ID,0)
        AND NVL(病人病区ID,0)=NVL(v_病人病区ID,0)
        AND NVL(病人科室ID,0)=NVL(v_病人科室ID,0)
        AND NVL(开单部门ID,0)=NVL(v_开单部门ID,0)
        AND NVL(执行部门ID,0)=NVL(v_执行部门ID,0)
        AND 收入项目ID+0=v_收入项目ID
        AND 来源途径+0=v_门诊标志;
    IF SQL%ROWCOUNT=0 THEN
        Insert INTO 病人未结费用(
            病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
        VALUES(
            v_病人ID,v_主页ID,v_病人病区ID,v_病人科室ID,v_开单部门ID,v_执行部门ID,
            v_收入项目ID,v_门诊标志,-1 * 结帐金额_IN);
    END IF;
    DELETE FROM 病人未结费用 WHERE 病人ID=v_病人ID And Nvl(金额,0)=0;

    --病人费用汇总
    UPDATE 病人费用汇总
        SET 结帐金额=NVL(结帐金额,0) + 结帐金额_IN
    WHERE 日期=TRUNC(SYSDATE)
        AND NVL(病人病区ID,0)=NVL(v_病人病区ID,0)
        AND NVL(病人科室ID,0)=NVL(v_病人科室ID,0)
        AND NVL(开单部门ID,0)=NVL(v_开单部门ID,0)
        AND NVL(执行部门ID,0)=NVL(v_执行部门ID,0)
        AND 收入项目ID+0=v_收入项目ID
        AND 来源途径=v_门诊标志
        AND 记帐费用=v_记帐费用;
    IF SQL%ROWCOUNT=0 THEN
        Insert INTO 病人费用汇总(
            日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
        VALUES(
            TRUNC(SYSDATE),v_病人病区ID,v_病人科室ID,v_开单部门ID,v_执行部门ID,v_收入项目ID,
            v_门诊标志,v_记帐费用,0,0,结帐金额_IN);
    END IF;
EXCEPTION
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_结帐费用记录_Insert;
/

CREATE OR REPLACE PROCEDURE zl_病人结帐记录_Insert(
--功能：插入一条病人结帐记录,同时处理费用结算误差
    ID_IN           病人结帐记录.ID%TYPE,
    单据号_IN       病人结帐记录.No%TYPE,
    病人ID_IN       病人结帐记录.病人ID%TYPE,
    误差NO_IN		病人费用记录.NO%TYPE,
    误差金额_IN     病人费用记录.结帐金额%TYPE,
    收费时间_IN     病人结帐记录.收费时间%TYPE,
    开始日期_IN     病人结帐记录.开始日期%TYPE,
    结束日期_IN     病人结帐记录.结束日期%TYPE
) AS
    --该游标包含误差项目的相关信息
    Cursor c_ErrItem is 
        Select 
            A.类别 AS 收费类别,A.ID AS 收费细目ID,A.计算单位,C.ID AS 收入项目ID,C.收据费目 
        From 收费细目 A,收费价目 B,收入项目 C,收费特定项目 D
        Where D.特定项目='误差项' And D.收费细目ID=A.ID
            And A.ID=B.收费细目ID And B.收入项目ID=C.ID
            And ((Sysdate Between B.执行日期 And B.终止日期) 
                Or (Sysdate>=B.执行日期 And B.终止日期 is NULL));
    r_ErrItem c_ErrItem%RowType;

    --该游标包含病人的相关信息
    Cursor c_Pati is
        Select A.姓名,A.性别,A.年龄,A.住院号,A.门诊号,B.主页ID,B.出院病床,
            B.当前病区ID,B.出院科室ID,Nvl(B.费别,A.费别) AS 费别
        From 病人信息 A,病案主页 B
        Where A.病人ID=病人ID_IN And A.病人ID=B.病人ID(+) And Nvl(A.住院次数,0)=B.主页ID(+);
    r_Pati c_Pati%RowType;

    v_Temp          Varchar2(255);
    v_人员部门ID    部门人员.部门ID%Type;
    v_人员编号		人员表.编号%Type;
    v_人员姓名      人员表.姓名%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;
BEGIN
    --操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp:=zl_Identity;
    v_人员部门ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);
 
    --病人结帐记录
    Insert INTO 病人结帐记录(
        ID,No,实际票号,记录状态,病人ID,开始日期,结束日期,收费时间,操作员编号,操作员姓名)
    VALUES(
        ID_IN,单据号_IN,NULL,1,病人ID_IN,开始日期_IN,结束日期_IN,收费时间_IN,v_人员编号,v_人员姓名);

    --处理结帐时结算产生的误差金额
    If Nvl(误差金额_IN,0)<>0 Then
        Open c_ErrItem;
        Fetch c_ErrItem Into r_ErrItem;
        If c_ErrItem%RowCount=0 Then
            Close c_ErrItem;
            v_Error:='不能正确读取处理费用误差的项目信息，请先检查该项目是否正确设置。';
            Raise Err_Custom;
        End IF;
        
        Open c_Pati;
        Fetch c_Pati Into r_Pati;        
        If c_Pati%RowCount=0 Then
            Close c_Pati;
            v_Error:='不能正确读取结帐病人信息。';
            Raise Err_Custom;
        End IF;
        
        --病人费用记录(记帐同时结帐):附加标志=9
        Insert Into 病人费用记录(
            ID,NO,实际票号,记录性质,记录状态,序号,从属父号,价格父号,多病人单,病人ID,主页ID,医嘱序号,
            门诊标志,姓名,性别,年龄,标识号,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,
            发药窗口,数次,加班标志,附加标志,婴儿费,记帐费用,收入项目ID,收据费目,标准单价,应收金额,实收金额,
            划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行状态,操作员姓名,操作员编号,
            结帐金额,结帐ID,是否上传)
        Values(
            病人费用记录_ID.Nextval,误差NO_IN,NULL,2,1,1,NULL,NULL,0,病人ID_IN,r_Pati.主页ID,NULL,
            Decode(r_Pati.主页ID,NULL,1,2),r_Pati.姓名,r_Pati.性别,r_Pati.年龄,Decode(r_Pati.主页ID,NULL,r_Pati.门诊号,r_Pati.住院号),
            r_Pati.出院病床,Nvl(r_Pati.当前病区ID,v_人员部门ID),Nvl(r_Pati.出院科室ID,v_人员部门ID),r_Pati.费别,
            r_ErrItem.收费类别,r_ErrItem.收费细目ID,r_ErrItem.计算单位,1,NULL,1,NULL,9,0,1,r_ErrItem.收入项目ID,
            r_ErrItem.收据费目,误差金额_IN,误差金额_IN,误差金额_IN,v_人员姓名,v_人员部门ID,v_人员姓名,收费时间_IN,
            收费时间_IN,v_人员部门ID,0,v_人员姓名,v_人员编号,误差金额_IN,ID_IN,1);

        --病人余额,病人未结费用(余额=余额+实收-结帐,所以可以不处理)

        --病人费用汇总
        Update 病人费用汇总
            Set 应收金额=Nvl(应收金额,0)+误差金额_IN,
                实收金额=Nvl(实收金额,0)+误差金额_IN,
                结帐金额=Nvl(结帐金额,0)+误差金额_IN
        Where 日期=Trunc(收费时间_IN)
            And Nvl(病人病区ID,0)=Nvl(r_Pati.当前病区ID,v_人员部门ID)
            And Nvl(病人科室ID,0)=Nvl(r_Pati.出院科室ID,v_人员部门ID)
            And Nvl(开单部门ID,0)=v_人员部门ID
            And Nvl(执行部门ID,0)=v_人员部门ID
            And 收入项目ID+0=r_ErrItem.收入项目ID
            And 来源途径=Decode(r_Pati.主页ID,NULL,1,2)
            And 记帐费用=1;
        IF SQL%RowCount=0 Then
            Insert Into 病人费用汇总(
                日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
            Values(
                Trunc(收费时间_IN),Nvl(r_Pati.当前病区ID,v_人员部门ID),Nvl(r_Pati.出院科室ID,v_人员部门ID),v_人员部门ID,
                v_人员部门ID,r_ErrItem.收入项目ID,Decode(r_Pati.主页ID,NULL,1,2),1,误差金额_IN,误差金额_IN,误差金额_IN);
        End IF;
            
        Close c_Pati;
        Close c_ErrItem;
    End IF;
EXCEPTION
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病人结帐记录_Insert;
/

Create Or Replace Procedure zl_住院收费结算_Update(
	   结帐Id_IN			病人费用记录.结帐ID%Type,
	   结帐结算_IN		    Varchar2,                   --结帐结算_IN-非医保时:结算方式|结算金额|结算号码||.....医保时:结算方式|结算金额|保险类别,保险密码,保险帐号||.....
	   冲预交_IN		    Varchar2,                   --冲预交_IN= ID|单据号|金额|记录状态||.....
	   误差金额_IN		    病人费用记录.实收金额%Type,
       误差NO_IN		    病人费用记录.NO%Type
 )
 As 
 --功能:处理结帐时和医保正式结算后,相关结算信息的调整
 --     因为虚拟结帐后,生成的医保结算金额总额及分摊可能会与正式结算时有差异,所以提供了校对功能,
 --		操作员在结算校对时,可以调整非医保结算方式的各种结算金额及方式,重新生成结算串,并且可能产生误差金额.
 
--病人信息 
     Cursor c_Pati(v_病人ID 病人信息.病人ID%Type) is 
     Select A.姓名,A.性别,A.年龄,A.住院号,A.门诊号,B.主页Id,B.出院病床,
                B.当前病区ID,B.出院科室Id,Nvl(B.费别,A.费别) AS 费别
            From 病人信息 A,病案主页 B
            Where A.病人ID=v_病人Id And A.病人ID=B.病人ID(+) And Nvl(A.住院次数,0)=B.主页ID(+);
     r_pati c_Pati%RowType;
 
 --过程变量
    v_结算内容		Varchar2(4000);
    v_当前结算		Varchar2(100);
    v_结算方式		病人预交记录.结算方式%Type;
    v_结算金额		病人预交记录.冲预交%Type;
    v_结算号码		Varchar2(100);          --保险结算记录时,存入:保险类别,保险密码,保险帐号
	
    v_收费类别		病人费用记录.收费类别%Type;
    v_收费细目ID	病人费用记录.收费细目ID%Type;
    v_计算单位		病人费用记录.计算单位%Type;
    v_收入项目ID	病人费用记录.收入项目ID%Type;
    v_收据费目		病人费用记录.收据费目%Type;
	
    v_No			病人预交记录.No%Type;
    v_病人Id		病人预交记录.病人Id%Type;   
    v_收款时间		病人预交记录.收款时间%Type;
    v_操作员编号	病人预交记录.操作员编号%Type;
    v_操作员姓名	病人预交记录.操作员姓名%Type;
    v_人员部门ID	部门人员.部门ID%Type;
    v_Temp			Varchar2(500);    
	
    v_预交金额		病人预交记录.冲预交%Type;	
    v_预交ID		病人预交记录.Id%Type;
    v_记录状态		病人预交记录.记录状态%Type;
	 
    v_保险类别		病人预交记录.缴款单位%Type;
    v_保险帐号		病人预交记录.单位开户行%Type;
    v_保险密码		病人预交记录.单位帐号%Type;
   
    v_Error			VARCHAR2(255);
    Err_Custom		EXCEPTION;
 Begin
 
 --1.取预交记录中的需要的相关信息
    Select No,病人Id,收费时间,操作员编号,操作员姓名 
		   Into  v_No,v_病人Id,v_收款时间,v_操作员编号,v_操作员姓名
    From 病人结帐记录 Where ID=结帐ID_IN;  
    
    Open c_Pati(v_病人Id);
    Fetch c_Pati Into r_Pati;
    
    --误差相关信息
    Begin
        Select A.类别,A.ID,A.计算单位,C.ID,C.收据费目 
        Into v_收费类别,v_收费细目ID,v_计算单位,v_收入项目ID,v_收据费目
        From 收费项目目录 A,收费价目 B,收入项目 C,收费特定项目 D
        Where D.特定项目='误差项' And D.收费细目ID=A.Id  And A.ID=B.收费细目ID And B.收入项目ID=C.ID
            And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-01-01','YYYY-MM-DD')) ;
    Exception
        When Others Then
        Begin
            v_Error:='不能正确读取收费误差项的信息，请先检查该项目是否设置正确。';
            Raise Err_Custom;
        End;
    End;	
	
 --2.删除旧的记录,回退汇总数据
    --回退人员缴款余额,病人余额,
    For C_DEL In (SELECT * FROM 病人预交记录 WHERE 结帐ID=结帐ID_IN And 记录性质=2) Loop
	      Update 人员缴款余额 Set 余额=Nvl(余额,0)-Nvl(C_DEL.冲预交,0) Where 结算方式=C_DEL.结算方式;	   
      	If SQL%RowCount=0 Then
               Insert Into 人员缴款余额(收款员,结算方式,性质,余额) Values(C_DEL.操作员姓名,C_DEL.结算方式,1,-1*C_DEL.冲预交);
    		End If;
    End Loop;
	
    If v_病人Id>0 Then
    	Begin
        	Select Sum(冲预交) Into V_预交金额 From 病人预交记录 Where 结帐Id=结帐id_IN And 记录性质 In (1,11);
        Exception
        	When Others Then NULL;
    	End;	
    	If v_预交金额<>0 Then
        	Update 病人余额 Set 预交余额=Nvl(预交余额,0)+V_预交金额 Where 病人ID=v_病人Id And 性质=1;
        	IF SQL%RowCount=0 Then
            	Insert Into 病人余额(病人ID,预交余额,性质) Values(v_病人Id,V_预交金额,1);
            End IF;
    	End If;
    End If;
    
	--回退病人费用汇总.         病人未结费用(因为新误差将立即结帐,所以不处理)  
	--只可能产生误差金额的变化. 旧误差只可能存在一行,仅为了变量处理方便而用游标
    For C_Error In (
        Select TRUNC(登记时间) as 日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,门诊标志,应收金额,实收金额,结帐金额
        From 病人费用记录
        Where 记录性质=2 And 记录状态=1 And 结帐Id=结帐Id_IN And 附加标志=9
    ) Loop
        Update 病人费用汇总
            Set 应收金额=Nvl(应收金额,0)-C_Error.应收金额,实收金额=Nvl(实收金额,0)-C_Error.实收金额,结帐金额=Nvl(结帐金额,0)-C_Error.结帐金额
        Where 日期=C_Error.日期
            And Nvl(病人病区ID,0)=Nvl(C_Error.病人病区ID,0) And Nvl(病人科室ID,0)=Nvl(C_Error.病人科室ID,0)
            And Nvl(开单部门ID,0)=Nvl(C_Error.开单部门ID,0) And Nvl(执行部门ID,0)=Nvl(C_Error.执行部门ID,0)
            And 收入项目ID+0=C_Error.收入项目Id And 来源途径=C_Error.门诊标志 And 记帐费用=1; 
        If SQL%RowCount=0 Then
            Insert Into 病人费用汇总(
                日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
            Values(
                C_Error.日期,C_Error.病人病区ID,C_Error.病人科室ID,C_Error.开单部门ID,C_Error.执行部门ID,
                C_Error.收入项目ID,C_Error.门诊标志,1,-1*C_Error.应收金额,-1*C_Error.实收金额,-1*C_Error.结帐金额);
        End If;
    End Loop; 
 
    --删除结帐缴款,保险结算记录		     
    Delete 病人预交记录 Where 结帐ID=结帐ID_IN And 记录性质=2; 
    --第一次冲预交的,清空冲减额
    Update 病人预交记录 Set 冲预交=Null,结帐Id=Null	Where 结帐Id=结帐ID_IN And 记录性质=1;
    --删除冲余款
    Delete 病人预交记录 Where 结帐Id=结帐ID_IN And 记录性质=11;
    --删除误差记录
    Delete 病人费用记录 Where 结帐Id=结帐Id_IN And 附加标志=9;	
 
 --3.产生病人费用记录的误差记录
    If 误差金额_IN <>0 Then		    
        --操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
        v_Temp:=zl_Identity;
        v_人员部门ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));             
        Insert Into 病人费用记录(
            ID,NO,实际票号,记录性质,记录状态,序号,从属父号,价格父号,多病人单,病人ID,主页ID,医嘱序号,
            门诊标志,姓名,性别,年龄,标识号,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,
            发药窗口,数次,加班标志,附加标志,婴儿费,记帐费用,收入项目ID,收据费目,标准单价,应收金额,实收金额,
            划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行状态,操作员姓名,操作员编号,
            结帐金额,结帐ID,是否上传)
        Values(
            病人费用记录_ID.Nextval,误差NO_IN,NULL,2,1,1,NULL,NULL,0,v_病人Id,r_Pati.主页ID,NULL,
            Decode(r_Pati.主页ID,NULL,1,2),r_Pati.姓名,r_Pati.性别,r_Pati.年龄,Decode(r_Pati.主页ID,NULL,r_Pati.门诊号,r_Pati.住院号),
            r_Pati.出院病床,Nvl(r_Pati.当前病区ID,v_人员部门ID),Nvl(r_Pati.出院科室ID,v_人员部门ID),r_Pati.费别,
            v_收费类别,v_收费细目ID,v_计算单位,1,NULL,1,NULL,9,0,1,v_收入项目ID,
            v_收据费目,误差金额_IN,误差金额_IN,误差金额_IN,v_操作员姓名,v_人员部门ID,v_操作员姓名,v_收款时间,
            v_收款时间,v_人员部门ID,0,v_操作员姓名,v_操作员编号,误差金额_IN,结帐ID_IN,1);
    End If;
  
 --4.重新生成病人预交记录相关数据	
    --4.1.补款结算,保险结算
    If 结帐结算_IN IS Not NULL Then
		v_结算内容:=结帐结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
      			v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
      			v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
      			v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
      			v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
      			v_结算号码:=LTrim(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
            If Instr(v_结算号码,',')>0 Then   --医保结算:保险类别,保险密码,保险帐号
                v_结算号码:=v_结算号码||',';
                  v_保险类别:=Substr(v_结算号码,1,Instr(v_结算号码,',')-1);
                v_结算号码:=Substr(v_结算号码,Instr(v_结算号码,',')+1);
                  v_保险密码:=Substr(v_结算号码,1,Instr(v_结算号码,',')-1);
                v_结算号码:=Substr(v_结算号码,Instr(v_结算号码,',')+1);
                  v_保险帐号:=Substr(v_结算号码,1,Instr(v_结算号码,',')-1);
                v_结算号码:=Null;
            Else
                v_保险类别:=Null;
                v_保险密码:=Null;
                v_保险帐号:=Null;
            End If;
			
  			If Nvl(v_结算金额,0)<>0 Then
  				Insert Into 病人预交记录(
                    ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号,
                    收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
                Values(
                    病人预交记录_ID.NextVal,v_No,Null,2,1,v_病人Id,r_Pati.主页Id,r_Pati.出院科室Id,
                    Null,v_结算方式,v_结算号码,'结帐缴款',v_保险类别,v_保险密码,v_保险帐号,
                    v_收款时间,v_操作员编号,v_操作员姓名,v_结算金额,结帐ID_IN);
  			End IF;
            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
     End IF;

    --4.2.预交结算
    IF 冲预交_IN Is Not Null Then
        v_结算内容:=冲预交_IN||'||';
        V_预交金额:=0;              --前面回退预交余额时用过此变量
        While v_结算内容 Is Not Null Loop
            v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
  		v_预交ID:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));      --是记录冲预交的ID
            v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
            v_结算号码:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);               --是记录冲预交的NO号
            v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
  			v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
            v_记录状态:=To_Number(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
                         
            IF v_预交Id <> 0 Then              --第一次冲预交
                UPDATE 病人预交记录 SET 冲预交 = v_结算金额, 结帐ID = 结帐ID_IN WHERE ID = v_预交Id;
            Else                            --冲上次剩余额
                Insert INTO 病人预交记录(ID,No,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,结算方式,结算号码,摘要, 
                                        缴款单位,单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,冲预交,结帐Id)
                    SELECT 病人预交记录_ID.Nextval, No, 实际票号, 11, v_记录状态, 病人ID,主页ID, 科室ID, NULL, 结算方式, 结算号码, 摘要, 
                            缴款单位,单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,v_结算金额, 结帐ID_IN
                    FROM 病人预交记录
                    WHERE No = v_结算号码 AND 记录性质 In(1,11) AND 记录状态 = v_记录状态 AND ROWNUM = 1;
            END IF;
            v_预交金额:=v_预交金额+v_结算金额;
            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
        
        --更新病人余额
        Update 病人余额 Set 预交余额=Nvl(预交余额,0)-v_预交金额 Where 病人ID=v_病人Id And 性质=1;
        IF SQL%RowCount=0 Then
            Insert Into 病人余额(病人ID,预交余额,性质) Values(v_病人Id,-1*v_预交金额,1);
        End IF;
        Delete From 病人余额 Where 病人ID=v_病人Id And 性质=1 And Nvl(费用余额,0)=0 And Nvl(预交余额,0)=0;
    End IF;
	
    --5.相关汇总表的处理	
    --汇总"人员缴款余额"
	--缴款结算,保险结算
    IF 结帐结算_IN IS Not NULL Then
        v_结算内容:=结帐结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
      			v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
      			v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
      			v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
      			v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
      			
      			If Nvl(v_结算金额,0)<>0 Then
      				Update 人员缴款余额	Set 余额=Nvl(余额,0)+Nvl(v_结算金额,0)
      				    Where 收款员=v_操作员姓名 And 性质=1 And 结算方式=v_结算方式;
      				If SQL%RowCount=0 Then
      					Insert Into 人员缴款余额(收款员,结算方式,性质,余额)
      					Values(v_操作员姓名,v_结算方式,1,Nvl(v_结算金额,0));
      				End If;
      			End IF;
            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
    End IF;
    Delete From 人员缴款余额 Where 性质=1 And 收款员=v_操作员姓名 And Nvl(余额,0)=0;

    --病人费用汇总,只需重汇误差行,因为其它项不会变,未结费用不变(新产生的误差项已结帐),只有一行误差记录,仅为使用变量方便而用游标
    For r_MoneyRow In (
        Select TRUNC(登记时间) as 日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,门诊标志,应收金额,实收金额,结帐金额
        From 病人费用记录
        Where 记录性质=2 And 记录状态=1 And 结帐Id=结帐Id_IN And 附加标志=9
	) Loop
        Update 病人费用汇总
            Set 应收金额=Nvl(应收金额,0)+r_MoneyRow.应收金额,实收金额=Nvl(实收金额,0)+r_MoneyRow.实收金额,结帐金额=Nvl(结帐金额,0)+r_MoneyRow.结帐金额
        Where 日期=r_MoneyRow.日期
            And Nvl(病人病区ID,0)=Nvl(r_MoneyRow.病人病区ID,0) And Nvl(病人科室ID,0)=Nvl(r_MoneyRow.病人科室ID,0)
            And Nvl(开单部门ID,0)=Nvl(r_MoneyRow.开单部门ID,0) And Nvl(执行部门ID,0)=Nvl(r_MoneyRow.执行部门ID,0)
            And 收入项目ID+0=r_MoneyRow.收入项目Id  And 来源途径=r_MoneyRow.门诊标志 And 记帐费用=1;

        If SQL%RowCount=0 Then
            Insert Into 病人费用汇总(
                日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
            Values(
                r_MoneyRow.日期,r_MoneyRow.病人病区ID,r_MoneyRow.病人科室ID,r_MoneyRow.开单部门ID,r_MoneyRow.执行部门ID,
                r_MoneyRow.收入项目ID,r_MoneyRow.门诊标志,1,r_MoneyRow.应收金额,r_MoneyRow.实收金额,r_MoneyRow.结帐金额);
        End If;
    End Loop; 
 
 	--6.医保相关表的处理
    Delete 医保核对表 Where 结帐Id=结帐Id_IN;
    
    Close c_Pati;
 
EXCEPTION
    WHEN Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS Then zl_ErrOrCenter(SQLCODE,SQLERRM);
End zl_住院收费结算_Update;
/



-------------------------------------------------------
--模块：就诊卡记录.SQL
Create Or Replace Procedure zl_就诊卡记录_Insert(
--参数：发卡类型=0-发卡,1-补卡,2-换卡(相当于重打)
--      换卡时,单据号_IN传入的是原发/补卡的单据号。
--      补卡/换卡后,再换卡时是以最后一次卡号为准。
    发卡类型_IN            Number, 
    单据号_IN            病人费用记录.No%Type,
    病人ID_IN            病人费用记录.病人id%Type, 
    主页ID_IN            病人费用记录.主页id%Type,
    标识号_IN            病人费用记录.标识号%Type, 
    费别_IN                病人费用记录.费别%Type,
    原卡号_IN            病人信息.就诊卡号%Type, 
    卡号_IN                病人信息.就诊卡号%Type,
    密码_IN                病人信息.卡验证码%Type, 
    姓名_IN                病人费用记录.姓名%Type,
    性别_IN                病人费用记录.性别%Type, 
    年龄_IN                病人费用记录.年龄%Type,
    病人病区ID_IN        病人费用记录.病人病区id%Type,
    病人科室ID_IN        病人费用记录.病人科室id%Type,
    收费细目ID_IN        病人费用记录.收费细目id%Type,
    收费类别_IN            病人费用记录.收费类别%Type,
    计算单位_IN            病人费用记录.计算单位%Type,
    收入项目ID_IN        病人费用记录.收入项目id%Type,
    收据费目_IN            病人费用记录.收据费目%Type,
    金额_IN                病人费用记录.实收金额%Type,
    执行部门ID_IN        病人费用记录.执行部门id%Type,
    开单部门ID_IN        病人费用记录.开单部门id%Type,
    操作员编号_IN        病人费用记录.操作员编号%Type,
    操作员姓名_IN        病人费用记录.操作员姓名%Type,
    加班标志_IN            病人费用记录.加班标志%Type,
    发卡时间_IN            病人费用记录.登记时间%Type,
    结算方式_IN            病人预交记录.结算方式%Type,
    领用ID_IN            票据使用明细.领用id%Type
) As
    Cursor c_PreCard Is
        Select Id As 费用ID From 病人费用记录 
        Where 记录性质 = 5 And 实际票号=原卡号_IN And 病人id = 病人ID_IN;
    r_CardRow c_PreCard%Rowtype;
    

    v_费用id    病人费用记录.Id%Type;
    v_结帐id    病人费用记录.结帐id%Type;
    v_收回ID    票据打印内容.ID%Type;
    v_打印ID    票据打印内容.ID%Type;

    Err_NoPreCard Exception;
Begin
    If Not 结算方式_IN Is Null Then
        Select 病人结帐记录_Id.Nextval Into v_结帐id From Dual;
    End If;

    If 发卡类型_IN <> 2 Then
        --就诊卡费用记录
        Select 病人费用记录_Id.Nextval Into v_费用id From Dual;

        Insert Into 病人费用记录
            (Id, 记录性质, 记录状态, No,实际票号,序号,病人id, 主页id, 病人病区id, 病人科室id, 标识号, 姓名, 性别, 年龄, 费别,
             记帐费用, 门诊标志, 加班标志, 开单部门id,开单人, 操作员编号, 操作员姓名, 发生时间, 登记时间, 收费细目id,
             收费类别, 计算单位, 付数, 数次, 发药窗口,附加标志, 执行部门id, 收入项目id, 收据费目, 标准单价, 应收金额,
             实收金额, 结帐id, 结帐金额)
        Values
            (v_费用id,5,1,单据号_IN,卡号_IN,1,病人ID_IN, Decode(主页ID_IN, 0, Null, 主页ID_IN),
             Decode(病人病区ID_IN, 0, Null, 病人病区ID_IN), Decode(病人科室ID_IN, 0, Null, 病人科室ID_IN),
             Decode(标识号_IN, 0, Null, 标识号_IN), 姓名_IN, 性别_IN, 年龄_IN, 费别_IN, Decode(结算方式_IN, Null, 1, 0), 3,
             加班标志_IN, 开单部门ID_IN, 操作员姓名_IN, 操作员编号_IN, 操作员姓名_IN, 发卡时间_IN, 发卡时间_IN, 收费细目ID_IN,
             收费类别_IN, 计算单位_IN, 1, 1, 卡号_IN, 发卡类型_IN, 执行部门ID_IN, 收入项目ID_IN, 收据费目_IN, 金额_IN, 金额_IN,
             金额_IN, v_结帐id, Decode(结算方式_IN, Null, Null, 金额_IN));
    
        --如果是现收就诊卡费用，则将结算填入病人预交记录
        If Not 结算方式_IN Is Null Then
            Insert Into 病人预交记录
                (Id, No, 记录性质, 记录状态, 病人id, 主页id, 科室id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id,
                 摘要)
            Values
                (病人预交记录_Id.Nextval, 单据号_IN, 5, 1, 病人ID_IN, Decode(主页ID_IN, 0, Null, 主页ID_IN),
                 Decode(病人科室ID_IN, 0, Null, 病人科室ID_IN), 结算方式_IN, 发卡时间_IN, 操作员编号_IN, 操作员姓名_IN, 金额_IN,
                 v_结帐id, '就诊卡费用');
        End If;
    
        IF Not 领用ID_IN Is Null then
            --发卡使用票据
            Select 票据打印内容_ID.Nextval Into v_打印ID From Dual;
            Insert Into 票据打印内容(
                ID,数据性质,NO)
            Values(
                v_打印ID,5,单据号_IN);

            Insert Into 票据使用明细(
                ID,票种,号码,性质,原因,领用id,打印id,使用时间,使用人)
            Values(
                票据使用明细_ID.Nextval,5,卡号_IN,1,1,领用ID_IN,v_打印ID,发卡时间_IN,操作员姓名_IN);
        
            --该批领用状态变化
            Update 票据领用记录
                Set 当前号码=卡号_IN,
                    剩余数量=Decode(Sign(剩余数量-1),-1,0,剩余数量-1)
            Where Id=Nvl(领用ID_IN,0);
        End IF;
    
        --相关汇总表的处理
        If 结算方式_IN Is Null Then
            --汇总'病人余额'
            Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) + 金额_IN Where 性质 = 1 And 病人id = 病人ID_IN;
        
            If Sql%Rowcount = 0 Then
                Insert Into 病人余额 (病人id, 性质, 预交余额, 费用余额) Values (病人ID_IN, 1, 0, 金额_IN);
            End If;
        
            Delete From 病人余额 Where 病人id = 病人ID_IN And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
        
            --汇总'病人未结费用'
            Update 病人未结费用
            Set 金额 = Nvl(金额, 0) + 金额_IN
            Where 病人id = 病人ID_IN And Nvl(主页id, 0) = Nvl(主页ID_IN, 0) And Nvl(病人病区id, 0) = Nvl(病人病区ID_IN, 0) And
                        Nvl(病人科室id, 0) = Nvl(病人科室ID_IN, 0) And Nvl(开单部门id, 0) = Nvl(开单部门ID_IN, 0) And
                        Nvl(执行部门id, 0) = Nvl(执行部门ID_IN, 0) And 收入项目id+0 = 收入项目ID_IN And 来源途径 = 3;
        
            If Sql%Rowcount = 0 Then
                Insert Into 病人未结费用
                    (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
                Values
                    (病人ID_IN, Decode(主页ID_IN, 0, Null, 主页ID_IN), Decode(病人病区ID_IN, 0, Null, 病人病区ID_IN),
                     Decode(病人科室ID_IN, 0, Null, 病人科室ID_IN), 开单部门ID_IN, 执行部门ID_IN, 收入项目ID_IN, 3, 金额_IN);
            End If;
        Else
            --汇总"人员缴款余额"
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) + 金额_IN
            Where 收款员 = 操作员姓名_IN And 性质 = 1 And 结算方式 = 结算方式_IN;
        
            If Sql%Rowcount = 0 Then
                Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_IN, 结算方式_IN, 1, 金额_IN);
            End If;
        
            Delete From 人员缴款余额
            Where 性质 = 1 And 收款员 = 操作员姓名_IN And 结算方式 = 结算方式_IN And Nvl(余额, 0) = 0;
        End If;
    
        --汇总'病人费用汇总'
        Update 病人费用汇总
        Set 应收金额 = Nvl(应收金额, 0) + 金额_IN, 实收金额 = Nvl(实收金额, 0) + 金额_IN,
                结帐金额 = Nvl(结帐金额, 0) + Decode(结算方式_IN, Null, 0, 金额_IN)
        Where 日期 = Trunc(发卡时间_IN) And Nvl(病人病区id, 0) = Nvl(病人病区ID_IN,0) 
                    And Nvl(病人科室id, 0) = Nvl(病人科室ID_IN,0) 
                    And 开单部门id = 开单部门ID_IN And 执行部门id = 执行部门ID_IN 
                    And 收入项目id+0 = 收入项目ID_IN And 来源途径 = 3 
                    And 记帐费用 = Decode(结算方式_IN, Null, 1, 0);
    
        If Sql%Rowcount = 0 Then
            Insert Into 病人费用汇总
                (日期, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 记帐费用, 应收金额, 实收金额,
                 结帐金额)
            Values
                (Trunc(发卡时间_IN), Decode(病人病区ID_IN, 0, Null, 病人病区ID_IN),
                 Decode(病人科室ID_IN, 0, Null, 病人科室ID_IN), 开单部门ID_IN, 执行部门ID_IN, 收入项目ID_IN, 3,
                 Decode(结算方式_IN, Null, 1, 0), 金额_IN, 金额_IN, Decode(结算方式_IN, Null, 0, 金额_IN));
        End If;
    Else
        --处理换卡方式
        --首先查找需要换卡的原就诊卡费用记录
        Open c_PreCard;
        Fetch c_PreCard Into r_CardRow;
    
        If c_PreCard%Rowcount = 0 Then
            Close c_PreCard;
            Raise Err_NoPreCard;
        Else
            --仅当有原费用记录时才处理
            --重打收回票据
            Begin
                Select Max(ID) Into v_收回ID From 票据打印内容 Where 数据性质=5 And NO=单据号_IN;
            Exception
                When Others Then NULL;
            End;
            If v_收回ID Is Not Null Then
                Insert Into 票据使用明细(
                    ID,票种,号码,性质,原因,领用id,打印id,使用时间,使用人)
                Select 
                    票据使用明细_ID.Nextval,票种,号码,2,4,领用ID,打印ID,发卡时间_IN,操作员姓名_IN
                From 票据使用明细
                Where 打印ID=v_收回ID And 票种=5 And 性质=1;
            End If;
            
            --重打发出票据
            Select 票据打印内容_ID.Nextval Into v_打印ID From Dual;

            Insert Into 票据打印内容(
                ID,数据性质,NO)
            Values(
                v_打印ID,5,单据号_IN);

            Insert Into 票据使用明细(
                ID,票种,号码,性质,原因,领用id,打印id,使用时间,使用人)
            Values(
                票据使用明细_ID.Nextval,5,卡号_IN,1,Decode(v_收回ID,NULL,1,3),领用ID_IN,v_打印ID,发卡时间_IN,操作员姓名_IN);
        
            --领用状态变化
            Update 票据领用记录
                Set 当前号码=卡号_IN, 
                    剩余数量=Decode(Sign(剩余数量-1),-1,0,剩余数量-1)
            Where Id=Nvl(领用ID_IN,0);

            --更改原发卡记录状态
            Update 病人费用记录 
                Set 实际票号=卡号_IN,
                    发药窗口=卡号_IN,
                    附加标志=2
            Where Id=r_CardRow.费用id;
        
            Close c_PreCard;
        End If;
    End If;

    --病人就诊卡信息变化
    Update 病人信息 Set 就诊卡号=卡号_IN, 卡验证码=密码_IN Where 病人id=病人ID_IN;
Exception
    When Err_NoPreCard Then Raise_Application_Error(-20101, '[ZLSOFT]没有发现原就诊卡发放记录,换卡操作失败！[ZLSOFT]');
    When Others Then Zl_Errorcenter(Sqlcode, Sqlerrm);
End zl_就诊卡记录_Insert;
/

Create Or Replace Procedure zl_就诊卡记录_Delete(
    单据号_In        病人费用记录.No%Type,
    操作员编号_In    病人费用记录.操作员编号%Type,
    操作员姓名_In    病人费用记录.操作员姓名%Type
) As
    Cursor c_Cardinfo Is
        Select a.Id As 费用id, Nvl(a.记帐费用, 0) As 记帐, a.结帐id, a.实际票号, a.病人id, Nvl(a.主页id, 0) As 主页id,
             Nvl(a.病人病区id, 0) As 病人病区id, Nvl(a.病人科室id, 0) As 病人科室id, Nvl(a.开单部门id, 0) As 开单部门id,
             Nvl(a.执行部门id, 0) As 执行部门id, a.收入项目id, a.实收金额, b.结算方式, b.冲预交
        From 病人费用记录 a, 病人预交记录 b
        Where a.记录性质 = 5 And a.记录状态 = 1 And a.No = 单据号_In And a.结帐id = b.结帐id(+);
    r_Cardrow c_Cardinfo%Rowtype;
    
    v_费用id    病人费用记录.Id%Type;
    v_结帐id    病人费用记录.结帐id%Type;
    v_打印ID    票据打印内容.ID%Type;

    v_Date Date;
    Err_Custom Exception;
Begin
    Open c_Cardinfo;
    Fetch c_Cardinfo Into r_Cardrow;

    --首先判断要退卡的记录是否存在
    If c_Cardinfo%Rowcount = 0 Then
        Close c_Cardinfo;
        Raise Err_Custom;
    Else
        Select Sysdate Into v_Date From Dual;
        Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
    
        If r_Cardrow.记帐 = 0 Then
            Select 病人结帐记录_Id.Nextval Into v_结帐id From Dual;
        End If;
    
        --退除就诊卡费用记录
        Insert Into 病人费用记录
            (Id, No,实际票号,记录性质, 记录状态, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别,
             收费类别, 收费细目id, 计算单位, 付数, 数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用,
             标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id,
             结帐金额)
            Select v_费用id, No,实际票号,记录性质, 2, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别,
                 收费类别, 收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用,
                 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号_In, 操作员姓名_In, 发生时间,
                 v_Date, v_结帐id, Decode(v_结帐id, Null, Null, -结帐金额)
            From 病人费用记录
            Where Id = r_Cardrow.费用id;
    
        Update 病人费用记录 Set 记录状态 = 3 Where Id = r_Cardrow.费用id;
    
        --预交款里现收的结算金额
        If r_Cardrow.记帐 = 0 Then
            Insert Into 病人预交记录(
                Id, No, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交,结帐id)
            Select 病人预交记录_Id.Nextval, No, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 
                结算方式, v_Date,操作员编号_In, 操作员姓名_In, -冲预交, v_结帐id
            From 病人预交记录
            Where 记录性质 = 5 And 记录状态 = 1 And 结帐id = r_Cardrow.结帐id;
        
            Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 5 And 记录状态 = 1 And 结帐id = r_Cardrow.结帐id;
        End If;
    
        --退卡收回票据
        Begin
            Select Max(ID) Into v_打印ID From 票据打印内容 Where 数据性质=5 And NO=单据号_IN;
        Exception
            When Others Then NULL;
        End;
        If v_打印ID Is Not Null Then
            Insert Into 票据使用明细(
                ID,票种,号码,性质,原因,领用id,打印id,使用时间,使用人)
            Select
                票据使用明细_ID.Nextval,票种,号码,2,2,领用id,打印ID,v_Date,操作员姓名_In
            From 票据使用明细
            Where 打印ID=v_打印ID And 票种=5 And 性质=1;
        End If;
    
        --更新病人信息
        Update 病人信息 Set 就诊卡号=Null, 卡验证码=Null Where 就诊卡号=r_Cardrow.实际票号;
    
        --相关汇总表的处理
        If r_Cardrow.记帐 = 1 Then
            --汇总'病人余额'
            Update 病人余额
                Set 费用余额 = Nvl(费用余额, 0) + (-1 * r_Cardrow.实收金额)
            Where 性质 = 1 And 病人id = r_Cardrow.病人id;
        
            If Sql%Rowcount = 0 Then
                Insert Into 病人余额
                    (病人id, 性质, 预交余额, 费用余额)
                Values
                    (r_Cardrow.病人id, 1, 0, -1 * r_Cardrow.实收金额);
            End If;
        
            --汇总'病人未结费用'
            Update 病人未结费用
                Set 金额 = Nvl(金额, 0) + (-1 * r_Cardrow.实收金额)
            Where 病人id = r_Cardrow.病人id And Nvl(主页id, 0) = r_Cardrow.主页id And
                Nvl(病人病区id, 0) = r_Cardrow.病人病区id And Nvl(病人科室id, 0) = r_Cardrow.病人科室id And
                Nvl(开单部门id, 0) = r_Cardrow.开单部门id And Nvl(执行部门id, 0) = r_Cardrow.执行部门id And
                收入项目id+0 = r_Cardrow.收入项目id And 来源途径 = 3;
        
            If Sql%Rowcount = 0 Then
                Insert Into 病人未结费用
                    (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
                Values
                    (r_Cardrow.病人id, Decode(r_Cardrow.主页id, 0, Null, r_Cardrow.主页id),
                     Decode(r_Cardrow.病人病区id, 0, Null, r_Cardrow.病人病区id),
                     Decode(r_Cardrow.病人科室id, 0, Null, r_Cardrow.病人科室id),
                     Decode(r_Cardrow.开单部门id, 0, Null, r_Cardrow.开单部门id),
                     Decode(r_Cardrow.执行部门id, 0, Null, r_Cardrow.执行部门id), r_Cardrow.收入项目id, 3, -1 * r_Cardrow.实收金额);
            End If;
        Elsif r_Cardrow.结算方式 Is Not Null Then
            --汇总"人员缴款余额"
            Update 人员缴款余额
                Set 余额 = Nvl(余额, 0) + (-1 * r_Cardrow.冲预交)
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Cardrow.结算方式;
        
            If Sql%Rowcount = 0 Then
                Insert Into 人员缴款余额
                    (收款员, 结算方式, 性质, 余额)
                Values
                    (操作员姓名_In, r_Cardrow.结算方式, 1, -1 * r_Cardrow.冲预交);
            End If;
        
            Delete From 人员缴款余额
            Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = r_Cardrow.结算方式 And Nvl(余额, 0) = 0;
        End If;
    
        --汇总'病人费用汇总'
        Update 病人费用汇总
            Set 应收金额 = Nvl(应收金额, 0) + (-1 * r_Cardrow.实收金额), 实收金额 = Nvl(实收金额, 0) + (-1 * r_Cardrow.实收金额),
                结帐金额 = Nvl(结帐金额, 0) + Decode(r_Cardrow.记帐, 0, -1 * r_Cardrow.实收金额, Null)
        Where 日期 = Trunc(v_Date) And Nvl(病人病区id, 0) = r_Cardrow.病人病区id And
                    Nvl(病人科室id, 0) = r_Cardrow.病人科室id And Nvl(开单部门id, 0) = r_Cardrow.开单部门id And
                    Nvl(执行部门id, 0) = r_Cardrow.执行部门id And 收入项目id+0 = r_Cardrow.收入项目id And 来源途径 = 3 And
                    记帐费用 = r_Cardrow.记帐;
    
        If Sql%Rowcount = 0 Then
            Insert Into 病人费用汇总
                (日期, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 记帐费用, 应收金额, 实收金额,
                 结帐金额)
            Values
                (Trunc(v_Date), Decode(r_Cardrow.病人病区id, 0, Null, r_Cardrow.病人病区id),
                 Decode(r_Cardrow.病人科室id, 0, Null, r_Cardrow.病人科室id),
                 Decode(r_Cardrow.开单部门id, 0, Null, r_Cardrow.开单部门id),
                 Decode(r_Cardrow.执行部门id, 0, Null, r_Cardrow.执行部门id), r_Cardrow.收入项目id, 3, r_Cardrow.记帐,
                 -1 * r_Cardrow.实收金额, -1 * r_Cardrow.实收金额, Decode(r_Cardrow.记帐, 0, -1 * r_Cardrow.实收金额, Null));
        End If;
    
        Close c_Cardinfo;
    End If;
Exception
    When Err_Custom Then Raise_Application_Error(-20999, '没有发现要退卡的记录,该记录可能已经退除！');
    When Others Then Zl_ErrorCenter(Sqlcode, Sqlerrm);
End zl_就诊卡记录_Delete;
/

-------------------------------------------------------
--模块：门诊记帐记录.SQL
Create Or Replace Procedure zl_门诊记帐记录_Delete(
    NO_IN            病人费用记录.NO%Type,
    序号_IN          Varchar2,
    操作员编号_IN    病人费用记录.操作员编号%Type,
    操作员姓名_IN    病人费用记录.操作员姓名%Type
)
AS
	--功能：冲销一张门诊记帐单据中指定序号行
	--序号：格式如"1,3,5,7,8",为空表示冲销所有可冲销行
    --该光标用于销帐指定费用行

    --该游标为要退费单据的所有原始记录
    Cursor c_Bill is
        Select * From 病人费用记录
        Where NO=NO_IN And 记录性质=2 And 记录状态 IN(0,1,3) And 门诊标志=1
        Order by 收费细目ID,序号;

    --该游标用于处理药品库存可用数量
    --不要管费用的执行状态,因为先于此步处理
    Cursor c_Stock is
        Select * From 药品收发记录
        Where NO=NO_IN And 单据 IN(9,25) And Mod(记录状态,3)=1 And 审核人 IS NULL
            And 费用ID IN(
                Select ID From 病人费用记录 
                Where NO=NO_IN And 记录性质=2 And 记录状态 IN(0,1,3) 
                    And 收费类别 IN('4','5','6','7') And 门诊标志=1
                    And (INSTR(','||序号_IN||',',','||序号||',')>0 Or 序号_IN Is Null)
                )
        Order BY 药品ID;
    
    --该游标用于处理未发药品记录
    Cursor c_Spare is
        Select * From 未发药品记录 Where NO=NO_IN And 单据 IN(9,25);

    --该游标用于处理费用记录序号
    Cursor c_Serial is
        Select 序号,价格父号 From 病人费用记录 Where NO=NO_IN And 记录性质=2 And 记录状态 IN(0,1,3) Order BY 序号;

	v_医嘱ID		病人医嘱记录.ID%Type;
	v_划价			Number;
	v_父号			病人费用记录.价格父号%Type;

    --部分退费计算变量
    v_剩余数量		Number;
    v_剩余应收		Number;
    v_剩余实收		Number;
    v_剩余统筹		Number;

    v_准退数量		Number;
    v_退费次数		Number;

    v_应收金额		Number;
    v_实收金额		Number;
    v_统筹金额		Number;

    v_Dec			Number;
	
    v_Count			Number;
    v_CurDate		Date;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
    --是否已经全部完全执行(只是整张单据的检查)
    Select Nvl(Count(*),0) Into v_Count 
    From 病人费用记录 
    Where NO=NO_IN And 记录性质=2 And 记录状态 IN(0,1,3) And Nvl(执行状态,0)<>1;
    IF v_Count = 0 Then
        v_Error := '该单据中的项目已经全部完全执行！';
        Raise Err_Custom;
    End IF;

    --未完全执行的项目是否有剩余数量(只是整张单据的检查)
    Select Nvl(Count(*),0) Into v_Count
    From (
        Select 序号,Sum(数量) as 剩余数量
        From (
            Select 记录状态,Nvl(价格父号,序号) as 序号,
                Avg(Nvl(付数,1)*数次) as 数量 
            From 病人费用记录
            Where NO=NO_IN And 记录性质=2 And 门诊标志=1
                And Nvl(价格父号,序号) IN (
                        Select Nvl(价格父号,序号) 
                        From 病人费用记录 
                        Where NO=NO_IN And 记录性质=2 And 门诊标志=1
                            And 记录状态 IN(0,1,3) And Nvl(执行状态,0)<>1)
            Group by 记录状态,Nvl(价格父号,序号)
            )
        Group by 序号 Having Sum(数量)<>0);
    IF v_Count = 0 Then
        v_Error := '该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！';
        Raise Err_Custom;
    End IF;
    
    ---------------------------------------------------------------------------------
    --公用变量
    Select Sysdate Into v_CurDate From Dual;

    --金额小数位数
    Begin
        Select To_Number(参数值) Into v_Dec From 系统参数表 Where 参数号=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --循环处理每行费用(收入项目行)
	For r_Bill IN c_Bill Loop
		IF INSTR(','||序号_IN||',',','||Nvl(r_Bill.价格父号,r_Bill.序号)||',') >0 Or 序号_IN Is Null Then
			Select Decode(记录状态,0,1,0) Into v_划价 From 病人费用记录 Where ID=r_Bill.ID;
			If v_划价=0 Then
				IF Nvl(r_Bill.执行状态,0)<>1 Then
					--求剩余数量,剩余应收,剩余实收
					Select 
						Sum(Nvl(付数,1)*数次),Sum(应收金额),Sum(实收金额),Sum(统筹金额)
						Into v_剩余数量,v_剩余应收,v_剩余实收,v_剩余统筹
					From 病人费用记录 
					Where NO=NO_IN And 记录性质=2 And 序号=r_Bill.序号;

					IF v_剩余数量=0 Then
						IF 序号_IN IS Not NULL Then 
							v_Error := '单据中第'||Nvl(r_Bill.价格父号,r_Bill.序号)||'行费用已经全部销帐！';
							Raise Err_Custom;
						End IF;
						--情况：未限定行号,原始单据中的该笔已经全部销帐(执行状态=0的一种可能)
					Else
						--准销数量(非药品项目为剩余数量,原始数量)
						IF Instr(',4,5,6,7,',r_Bill.收费类别)=0 Then
							v_准退数量:=v_剩余数量;
						Else
							Select Sum(Nvl(付数,1)*实际数量) Into v_准退数量
							From 药品收发记录
							Where NO=NO_IN And 单据 IN(9,25) And MOD(记录状态,3)=1 
								And 审核人 is NULL And 费用ID=r_Bill.ID;

							--不跟踪在用的卫生材料
							If r_Bill.收费类别='4' And Nvl(v_准退数量,0)=0 Then
								v_准退数量:=v_剩余数量;
							End IF;
						End if;

						--处理病人费用记录
						
						--该笔项目第几次销帐
						Select Nvl(Max(Abs(执行状态)),0)+1 Into v_退费次数
						From 病人费用记录 
						Where NO=NO_IN And 记录性质=2 And 记录状态=2 And 序号=r_Bill.序号;
						
						--金额=剩余金额*(准退数/剩余数)
						v_应收金额:=Round(v_剩余应收*(v_准退数量/v_剩余数量),v_Dec);
						v_实收金额:=Round(v_剩余实收*(v_准退数量/v_剩余数量),v_Dec);
						v_统筹金额:=Round(v_剩余统筹*(v_准退数量/v_剩余数量),v_Dec);

						--插入退费记录
						Insert Into 病人费用记录(
							ID,NO,记录性质,记录状态,序号,从属父号,价格父号,病人ID,医嘱序号,门诊标志,多病人单,婴儿费,姓名,
							性别,年龄,标识号,床号,费别,病人病区ID,病人科室ID,收费类别,收费细目ID,计算单位,付数,发药窗口,
							数次,加班标志,附加标志,收入项目ID,收据费目,记帐费用,标准单价,应收金额,实收金额,开单部门ID,
							开单人,执行部门ID,划价人,执行人,执行状态,执行时间,操作员编号,操作员姓名,发生时间,登记时间,
							保险项目否,保险大类ID,统筹金额,记帐单ID,摘要)
						Select 病人费用记录_ID.Nextval,NO,记录性质,2,序号,从属父号,价格父号,病人ID,医嘱序号,门诊标志,多病人单,
							婴儿费,姓名,性别,年龄,标识号,床号,费别,病人病区ID,病人科室ID,收费类别,收费细目ID,计算单位,
							Decode(Sign(v_准退数量-Nvl(付数,1)*数次),0,付数,1),发药窗口,
							Decode(Sign(v_准退数量-Nvl(付数,1)*数次),0,-1*数次,-1*v_准退数量),加班标志,附加标志,
							收入项目ID,收据费目,记帐费用,标准单价,-1*v_应收金额,-1*v_实收金额,开单部门ID,开单人,执行部门ID,
							划价人,执行人,-1*v_退费次数,执行时间,操作员编号_IN,操作员姓名_IN,发生时间,v_CurDate,
							保险项目否,保险大类ID,-1*v_统筹金额,记帐单ID,摘要
						From 病人费用记录 Where ID=r_Bill.ID;

						--记录病人医嘱附费对应的医嘱ID(不是主费用)
						If v_医嘱ID IS Null And r_Bill.医嘱序号 IS Not Null Then
							v_医嘱ID:=r_Bill.医嘱序号;
						End IF;

						--病人余额
						Update 病人余额
							Set 费用余额=Nvl(费用余额,0) - v_实收金额
						 Where 病人ID=r_Bill.病人ID And 性质=1;
						IF SQL%RowCount=0 Then
							Insert Into 病人余额(
								病人ID,性质,费用余额,预交余额)
							Values(
								r_Bill.病人ID,1,-1*v_实收金额,0);
						End IF;
						
						--病人未结费用
						Update 病人未结费用
							Set 金额=Nvl(金额,0) - v_实收金额
						 Where 病人ID=r_Bill.病人ID
							And Nvl(主页ID,0)=Nvl(r_Bill.主页ID,0)
							And Nvl(病人病区ID,0)=Nvl(r_Bill.病人病区ID,0)
							And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
							And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
							And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
							And 收入项目ID+0=r_Bill.收入项目ID And 来源途径+0=1;
						IF SQL%RowCount=0 Then
							Insert Into 病人未结费用(
								病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
							Values(
								r_Bill.病人ID,r_Bill.主页ID,r_Bill.病人病区ID,r_Bill.病人科室ID,
								r_Bill.开单部门ID,r_Bill.执行部门ID,r_Bill.收入项目ID,1,-1*v_实收金额);
						End IF;

						--处理病人费用汇总
						Update 病人费用汇总
							Set 应收金额=Nvl(应收金额,0) - v_应收金额,
								实收金额=Nvl(实收金额,0) - v_实收金额
						 Where 日期=Trunc(v_CurDate)
							And Nvl(病人病区ID,0)=Nvl(r_Bill.病人病区ID,0)
							And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
							And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
							And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
							And 收入项目ID+0=r_Bill.收入项目ID
							And 来源途径=1 And 记帐费用=1;
						IF SQL%RowCount=0 Then
							Insert Into 病人费用汇总(
								日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
							Values(
								Trunc(v_CurDate),r_Bill.病人病区ID,r_Bill.病人科室ID,r_Bill.开单部门ID,r_Bill.执行部门ID,
								r_Bill.收入项目ID,1,1,-1 * v_应收金额,-1 * v_实收金额,0);
						End IF;
						
						--标记原费用记录
						--执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1
						Update 病人费用记录 
							Set 记录状态=3,
								执行状态=Decode(Sign(v_准退数量-v_剩余数量),0,0,1) 
						Where ID=r_Bill.ID;
					End IF;
				Else
					IF 序号_IN Is Not Null Then
						v_Error := '单据中第'||Nvl(r_Bill.价格父号,r_Bill.序号)||'行费用已经完全执行,不能销帐！';
						Raise Err_Custom;
					End IF;
					--情况:没限定行号,原始单据中包括已经完全执行的
				End IF;
			End IF;
		End IF;
	End Loop;
    
    ---------------------------------------------------------------------------------
    --药品相关内容
    For r_Stock in c_Stock Loop
        --处理药品库存
        If r_Stock.库房ID IS Not Null then
            Update 药品库存
                Set 可用数量=Nvl(可用数量,0)+Nvl(r_Stock.付数,1)*Nvl(r_Stock.实际数量,0)
             Where 库房ID=r_Stock.库房ID And 药品ID=r_Stock.药品ID
                And Nvl(批次,0)=Nvl(r_Stock.批次,0) And 性质=1;
            IF SQL%RowCount=0 Then
                Insert Into 药品库存(
                    库房ID,药品ID,性质,批次,效期,可用数量,上次批号,上次产地,灭菌效期)
                Values(
                    r_Stock.库房ID,r_Stock.药品ID,1,r_Stock.批次,r_Stock.效期,
                    Nvl(r_Stock.付数,1)*Nvl(r_Stock.实际数量,0),r_Stock.批号,r_Stock.产地,r_Stock.灭菌效期);
            End IF;
        End IF;

        --删除药品收发记录
        Delete From 药品收发记录 Where ID=r_Stock.ID;
    End Loop;

    --未发药品记录
    For r_Spare IN c_Spare Loop
        Select Nvl(Count(*),0) Into v_Count
        From 药品收发记录 
        Where NO=NO_IN And 单据=r_Spare.单据 And Mod(记录状态,3)=1 
            And 审核人 is NULL And Nvl(库房ID,0)=Nvl(r_Spare.库房ID,0);
        If v_Count=0 Then
            Delete From 未发药品记录 Where 单据=r_Spare.单据 And NO=NO_IN And Nvl(库房ID,0)=Nvl(r_Spare.库房ID,0);
        End IF;
    End Loop;
	
	---------------------------------------------------------------------------------
	--如果是划价,直接删除费用记录(药品处理后)
	v_Count:=0;
	For r_Bill IN c_Bill Loop
		IF INSTR(','||序号_IN||',',','||Nvl(r_Bill.价格父号,r_Bill.序号)||',') >0 Or 序号_IN Is Null Then
			Select Decode(记录状态,0,1,0) Into v_划价 From 病人费用记录 Where ID=r_Bill.ID;
			If v_划价=1 Then
				IF Nvl(r_Bill.执行状态,0)<>1 Then
					Delete From 病人费用记录 Where ID=r_Bill.ID;
					v_Count:=v_Count+1;--记录是否有删除行

					--记录病人医嘱附费对应的医嘱ID(不是主费用)
					If v_医嘱ID IS Null And r_Bill.医嘱序号 IS Not Null Then
						v_医嘱ID:=r_Bill.医嘱序号;
					End IF;
				Else
					IF 序号_IN Is Not Null Then
						v_Error := '单据中第'||Nvl(r_Bill.价格父号,r_Bill.序号)||'行费用已经完全执行,不能销帐！';
						Raise Err_Custom;
					End IF;
					--情况:没限定行号,原始单据中包括已经完全执行的
				End IF;
			End IF;
		End IF;
	End Loop;

	--删除之后再统一调整序号
	If v_Count>0 Then
		v_Count:=1;
		For r_Serial In c_Serial Loop
			If r_Serial.价格父号 IS NULL Then 
				v_父号:=v_Count;
			End IF;

			Update 病人费用记录 
				Set 序号=v_Count,
					价格父号=Decode(价格父号,NULL,NULL,v_父号)
			Where NO=NO_IN And 记录性质=2 And 序号=r_Serial.序号;
			
			Update 病人费用记录
				Set 从属父号=v_Count
			Where NO=NO_IN And 记录性质=2 And 从属父号=r_Serial.序号;

			v_Count:=v_Count+1;
		End Loop;
	End IF;

	--整张单据全部冲完时，删除病人医嘱附费
	If 序号_IN IS NULL And v_医嘱ID IS Not NULL Then
		Select Nvl(Count(*),0) Into v_Count
		From (
			Select 序号,Sum(数量) as 剩余数量
			From (
				Select 记录状态,Nvl(价格父号,序号) as 序号,
					Avg(Nvl(付数,1)*数次) as 数量 
				From 病人费用记录
				Where NO=NO_IN And 记录性质=2 And 医嘱序号+0=v_医嘱ID
				Group by 记录状态,Nvl(价格父号,序号)
				)
			Group by 序号 Having Nvl(Sum(数量),0)<>0);
		IF v_Count = 0 Then
			Delete From 病人医嘱附费 Where 医嘱ID=v_医嘱ID And 记录性质=2 And NO=NO_IN;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_门诊记帐记录_Delete;
/

CREATE OR REPLACE PROCEDURE zl_门诊记帐记录_Verify (
    NO_IN			病人费用记录.NO%TYPE,
    操作员编号_IN   病人费用记录.操作员编号%TYPE,
    操作员姓名_IN   病人费用记录.操作员姓名%TYPE,
	序号_IN			Varchar2:=NULL,
	审核时间_IN		病人费用记录.登记时间%Type:=NULL
) AS
--功能：审核一张门诊记帐划价单
--参数：
--		序号_IN：格式如"1,3,5,7,8",为空表示审核所有未审核的行
--		审核时间_IN：用于部份需要统一控制或返回时间的地方
	--只读取指定序号的,未审核的部份进行处理
	Cursor c_Bill is
		Select * From 病人费用记录 
		Where 记录性质=2 And 记录状态=0 And NO=NO_IN
			And (Instr(','||序号_IN||',',','||Nvl(价格父号,序号)||',')>0 Or 序号_IN Is Null)
		Order BY 序号;

	--审核中包含跟踪在用的未发卫料时，根据参数设置是否自动发料  25-记帐单处方发料
	Cursor c_Stuff is
		Select NO,单据,库房ID From 未发药品记录
		Where NO=NO_IN And 单据=25 And 库房ID IS Not Null
			And Exists(Select 参数值 From 系统参数表 Where 参数号=92 And 参数值='1')
			And Exists(
				Select A.序号 From 病人费用记录 A,材料特性 B
				Where A.记录性质=2 And A.记录状态=1 And A.NO=NO_IN
					And (Instr(','||序号_IN||',',','||Nvl(A.价格父号,A.序号)||',')>0 Or 序号_IN Is Null)					
					And A.收费细目ID=B.材料ID And B.跟踪在用=1
				)
		Order BY 库房ID;

	v_Date	Date;
BEGIN
	If 审核时间_IN IS Null Then
		Select Sysdate Into v_Date From Dual;
	Else
		v_Date:=审核时间_IN;
	End IF;

	For r_Bill IN c_Bill Loop
		Update 病人费用记录
			Set 记录状态=1,
				操作员编号=操作员编号_IN,
				操作员姓名=操作员姓名_IN,
				登记时间=v_Date --已产生的药品记录的时间不变
		Where ID=r_Bill.ID;

		--药品收发记录.填制日期
		Update 药品收发记录
			Set 填制日期=Decode(Sign(Nvl(审核日期,v_Date)-v_Date),-1,填制日期,v_Date)  
		Where NO=NO_IN AND 单据 IN(9,25)  AND 费用ID=r_Bill.ID;

		--病人余额
		Update 病人余额
			Set 费用余额=Nvl(费用余额,0)+Nvl(r_Bill.实收金额,0)
		Where 病人ID=r_Bill.病人ID And 性质=1;

		IF SQL%RowCount=0 Then
			Insert Into 病人余额(
				病人ID,性质,费用余额,预交余额)
			Values(
				r_Bill.病人ID,1,r_Bill.实收金额,0);
		End IF;

		--病人未结费用
		Update 病人未结费用
			Set 金额=Nvl(金额,0)+Nvl(r_Bill.实收金额,0)
		 Where 病人ID=r_Bill.病人ID
			And Nvl(主页ID,0)=Nvl(r_Bill.主页ID,0)
			And Nvl(病人病区ID,0)=Nvl(r_Bill.病人病区ID,0)
			And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
			And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
			And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
			And 收入项目ID+0=r_Bill.收入项目ID
			And 来源途径+0=r_Bill.门诊标志;

		IF SQL%RowCount=0 Then
			Insert Into 病人未结费用(
				病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
			Values(
				r_Bill.病人ID,r_Bill.主页ID,r_Bill.病人病区ID,r_Bill.病人科室ID,r_Bill.开单部门ID,r_Bill.执行部门ID,r_Bill.收入项目ID,r_Bill.门诊标志,Nvl(r_Bill.实收金额,0));
		End IF;

		--病人费用汇总
		Update 病人费用汇总
			Set 应收金额=Nvl(应收金额,0)+Nvl(r_Bill.应收金额,0),
				实收金额=Nvl(实收金额,0)+Nvl(r_Bill.实收金额,0)
		 Where 日期=Trunc(v_Date)
			And Nvl(病人病区ID,0)=Nvl(r_Bill.病人病区ID,0)
			And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
			And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
			And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
			And 收入项目ID+0=r_Bill.收入项目ID
			And 来源途径=r_Bill.门诊标志 And 记帐费用=1;

		IF SQL%RowCount=0 Then
			Insert Into 病人费用汇总(
				日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,
				来源途径,记帐费用,应收金额,实收金额,结帐金额)
			Values(
				Trunc(v_Date),r_Bill.病人病区ID,r_Bill.病人科室ID,r_Bill.开单部门ID,
				r_Bill.执行部门ID,r_Bill.收入项目ID,r_Bill.门诊标志,1,r_Bill.应收金额,r_Bill.实收金额,0);
		End IF;
	End Loop;

	--库房中的药品已全部审核则标为已收费
	Update 未发药品记录 Set 已收费=1,填制日期=v_Date
	Where NO=NO_IN And 单据=9 And Nvl(已收费,0)=0 
		And Nvl(库房ID,0) Not IN(
			Select Distinct Nvl(执行部门ID,0) From 病人费用记录 
				Where 记录性质=2 And NO=NO_IN And 收费类别 IN('5','6','7') And 记录状态=0);

	Update 未发药品记录 Set 已收费=1,填制日期=v_Date
	Where NO=NO_IN And 单据=25 And Nvl(已收费,0)=0 
		And Nvl(库房ID,0) Not IN(
			Select Distinct Nvl(执行部门ID,0) From 病人费用记录 
				Where 记录性质=2 And NO=NO_IN And 收费类别='4' And 记录状态=0);

	--处理卫料自动发料
	For r_Stuff In c_Stuff Loop
		zl_材料收发记录_处方发料(r_Stuff.库房ID,r_Stuff.单据,r_Stuff.NO,操作员姓名_IN,操作员姓名_IN,操作员姓名_IN,1,Sysdate);
	End Loop;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_门诊记帐记录_Verify;
/

Create Or Replace Procedure zl_门诊记帐记录_Insert(
    NO_IN				病人费用记录.NO%Type,
    序号_IN				病人费用记录.序号%Type,
    病人ID_IN			病人费用记录.病人ID%Type,
    标识号_IN			病人费用记录.标识号%Type,
    姓名_IN				病人费用记录.姓名%Type,
    性别_IN				病人费用记录.性别%Type,
    年龄_IN				病人费用记录.年龄%Type,
    费别_IN				病人费用记录.费别%Type,
    加班标志_IN			病人费用记录.加班标志%Type,
    婴儿费_IN			病人费用记录.婴儿费%Type,
	病人病区ID_IN		病人费用记录.病人病区ID%Type,
	病人科室ID_IN		病人费用记录.病人科室ID%Type,
    开单部门ID_IN		病人费用记录.开单部门ID%Type,
    开单人_IN			病人费用记录.开单人%Type,
    从属父号_IN			病人费用记录.从属父号%Type,
    收费细目ID_IN		病人费用记录.收费细目ID%Type,
    收费类别_IN			病人费用记录.收费类别%Type,
    计算单位_IN			病人费用记录.计算单位%Type,
    付数_IN				病人费用记录.付数%Type,
    数次_IN				病人费用记录.数次%Type,
    附加标志_IN			病人费用记录.附加标志%Type,
    执行部门ID_IN		病人费用记录.执行部门ID%Type,
    价格父号_IN			病人费用记录.价格父号%Type,
    收入项目ID_IN		病人费用记录.收入项目ID%Type,
    收据费目_IN			病人费用记录.收据费目%Type,
    标准单价_IN			病人费用记录.标准单价%Type,
    应收金额_IN			病人费用记录.应收金额%Type,
    实收金额_IN			病人费用记录.实收金额%Type,
    发生时间_IN			病人费用记录.发生时间%Type,
    登记时间_IN			病人费用记录.登记时间%Type,
    药品摘要_IN			药品收发记录.摘要%Type,
    划价_IN				Number,
    操作员编号_IN		病人费用记录.操作员编号%Type,
    操作员姓名_IN		病人费用记录.操作员姓名%Type,
    类别ID_IN			药品单据性质.类别ID%Type:=Null,
    记帐单ID_IN			病人费用记录.记帐单ID%Type:=Null,
    费用摘要_IN			病人费用记录.摘要%Type:=Null,
    医嘱序号_IN			病人费用记录.医嘱序号%TYPE:=NULL,
    频次_IN				药品收发记录.频次%Type:=NULL,
    单量_IN				药品收发记录.单量%Type:=NULL,
    用法_IN				药品收发记录.用法%Type:=NULL,--用法[|煎法]
    期效_IN				药品收发记录.扣率%Type:=NULL,
    计价特性_IN			药品收发记录.扣率%Type:=NULL
)
AS
    --功能：新收一张门诊记帐单据
    --参数：
    --   药品摘要_IN:修改保存新单据时用。目前仅用于存放于药品收发记录的摘要中。
    --         原单据(记录状态=2)记录修改产生的新单据号。
    --         新单据(记录状态=1)记录所修改的原单据号。
    v_费用ID 病人费用记录.ID%Type;
    v_优先级 未发药品记录.优先级%Type;

    --药房分批、时价药品--
    ------------------------------------------------------------
    --该游标用于分批药品数量分解
    Cursor c_Stock is
        Select * From 药品库存 
        Where 药品ID=收费细目ID_IN And 库房ID=执行部门ID_IN
            And 性质=1 And(Nvl(批次,0)=0 Or 效期 is Null Or 效期>Trunc(Sysdate))
            And Nvl(可用数量,0)<>0
        Order By Nvl(批次,0);
    r_Stock c_Stock%RowType;
    
    --属性
    v_分批			药品规格.药房分批%Type;
    v_时价			收费项目目录.是否变价%Type;
    v_名称			收费项目目录.名称%Type;
    --临时变量
    v_总数量		Number;
    v_当前数量		Number;
    v_总金额		Number;
    v_当前单价		Number;
    --药品收发记录
    v_批次			药品收发记录.批次%Type;
    v_产地			药品收发记录.产地%Type;
    v_批号			药品收发记录.批号%Type;
    v_效期			药品收发记录.效期%Type;
    v_序号			药品收发记录.序号%Type;
    v_扣率			药品收发记录.扣率%Type;
	v_灭菌效期		药品收发记录.灭菌效期%Type;
	v_灭菌日期		药品收发记录.灭菌日期%Type;
    ------------------------------------------------------------
	v_用法			药品收发记录.用法%Type;
	v_煎法			药品收发记录.外观%Type;

    v_Dec			Number;
	v_Count			Number;
	v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
	--药品用法煎法分解
	IF 用法_IN IS Not NULL Then
		IF Instr(用法_IN,'|')>0 Then
			v_用法:=Substr(用法_IN,1,Instr(用法_IN,'|')-1);
			v_煎法:=Substr(用法_IN,Instr(用法_IN,'|')+1);
		Else
			v_用法:=用法_IN;
		End IF;
	End IF;

    --金额小数位数
    Begin
        Select To_Number(参数值) Into v_Dec From 系统参数表 Where 参数号=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --病人费用记录
    Select 病人费用记录_ID.Nextval Into v_费用ID From Dual;

    Insert Into 病人费用记录(
        ID,记录性质,NO,记录状态,序号,从属父号,价格父号,门诊标志,病人ID,标识号,
        姓名,性别,年龄,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,
        付数,数次,加班标志,附加标志,收入项目ID,收据费目,标准单价,应收金额,实收金额,
        记帐费用,划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行状态,
        操作员编号,操作员姓名,婴儿费,记帐单ID,摘要,医嘱序号)
    Values(
        v_费用ID,2,NO_IN,Decode(划价_IN,1,0,1),序号_IN,Decode(从属父号_IN,0,Null,从属父号_IN),
        Decode(价格父号_IN,0,Null,价格父号_IN),1,病人ID_IN,
        Decode(标识号_IN,0,Null,标识号_IN),姓名_IN,性别_IN,年龄_IN,病人病区ID_IN,
        病人科室ID_IN,费别_IN,收费类别_IN,收费细目ID_IN,计算单位_IN,付数_IN,数次_IN,
        加班标志_IN,附加标志_IN,收入项目ID_IN,收据费目_IN,标准单价_IN,应收金额_IN,
        实收金额_IN,1,操作员姓名_IN,开单部门ID_IN,开单人_IN,发生时间_IN,登记时间_IN,
        执行部门ID_IN,0,Decode(划价_IN,1,Null,操作员编号_IN),
        Decode(划价_IN,1,Null,操作员姓名_IN),婴儿费_IN,记帐单ID_IN,费用摘要_IN,医嘱序号_IN);

    --相关汇总表的处理
	If Nvl(划价_IN,0)=0 Then
		--病人余额
		Update 病人余额
			Set 费用余额=Nvl(费用余额,0)+实收金额_IN
		Where 病人ID=病人ID_IN And 性质=1;

		IF SQL%RowCount=0 Then
			Insert Into 病人余额(
				病人ID,性质,费用余额,预交余额)
			Values(
				病人ID_IN,1,实收金额_IN,0);
		End IF;

		--病人未结费用
		Update 病人未结费用
			Set 金额=Nvl(金额,0)+实收金额_IN
		 Where 病人ID=病人ID_IN
			And Nvl(主页ID,0)=0
			And Nvl(病人病区ID,0)=Nvl(病人病区ID_IN,0)
			And Nvl(病人科室ID,0)=Nvl(病人科室ID_IN,0)
			And Nvl(开单部门ID,0)=Nvl(开单部门ID_IN,0)
			And Nvl(执行部门ID,0)=Nvl(执行部门ID_IN,0)
			And 收入项目ID+0=收入项目ID_IN
			And 来源途径+0=1;

		IF SQL%RowCount=0 Then
			Insert Into 病人未结费用(
				病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
			Values(
				病人ID_IN,Null,病人病区ID_IN,病人科室ID_IN,开单部门ID_IN,执行部门ID_IN,收入项目ID_IN,1,实收金额_IN);
		End IF;

		--病人费用汇总
		Update 病人费用汇总
			Set 应收金额=Nvl(应收金额,0)+应收金额_IN,
				实收金额=Nvl(实收金额,0)+实收金额_IN
		 Where 日期=Trunc(登记时间_IN)
			And Nvl(病人病区ID,0)=Nvl(病人病区ID_IN,0)
			And Nvl(病人科室ID,0)=Nvl(病人科室ID_IN,0)
			And Nvl(开单部门ID,0)=Nvl(开单部门ID_IN,0)
			And Nvl(执行部门ID,0)=Nvl(执行部门ID_IN,0)
			And 收入项目ID+0=收入项目ID_IN
			And 来源途径=1 And 记帐费用=1;

		IF SQL%RowCount=0 Then
			Insert Into 病人费用汇总(
				日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,
				来源途径,记帐费用,应收金额,实收金额,结帐金额)
			Values(
				Trunc(登记时间_IN),病人病区ID_IN,病人科室ID_IN,开单部门ID_IN,
				执行部门ID_IN,收入项目ID_IN,1,1,应收金额_IN,实收金额_IN,0);
		End IF;
	End IF;

    --药品和卫生材料部分
	v_Count:=0;
	If 收费类别_IN='4' Then--跟踪在用的卫材才处理
		Select 跟踪在用 Into v_Count From 材料特性 Where 材料ID=收费细目ID_IN;
	End IF;
    IF 收费类别_IN in('5','6','7') Or (收费类别_IN='4' And Nvl(v_Count,0)=1) Then
		If 收费类别_IN='4' Then
			Select Nvl(A.在用分批,0),Nvl(B.是否变价,0),B.名称 
				Into v_分批,v_时价,v_名称
			From 材料特性 A,收费项目目录 B
			Where A.材料ID=B.ID And B.ID=收费细目ID_IN;
		Else
			Select Nvl(A.药房分批,0),Nvl(B.是否变价,0),B.名称 
				Into v_分批,v_时价,v_名称
			From 药品规格 A,收费项目目录 B
			Where A.药品ID=B.ID And B.ID=收费细目ID_IN;
		End IF;

        v_总数量:=付数_IN*数次_IN;
        v_总金额:=0;
        Open c_Stock;

        While v_总数量<>0 Loop
            Fetch c_Stock Into r_Stock;
            IF c_Stock%NotFound Then
                --第一次就没有库存,分批或时价都不允许。
                --分批药品数量分解不完,也就是库存不足。
                IF v_分批=1 Or v_时价=1 Then
                    Close c_Stock;
					If 医嘱序号_IN IS NULL Then
						If 收费类别_IN='4' Then
							v_Error:='第 '||序号_IN||' 行的分批或时价卫生材料"'||v_名称||'"没有足够的库存！';
						Else
							v_Error:='第 '||序号_IN||' 行的分批或时价药品"'||v_名称||'"没有足够的库存！';
						End IF;
					Else
						If 收费类别_IN='4' Then
							v_Error:='在处理病人"'||姓名_IN||'"时发现分批或时价卫生材料"'||v_名称||'"没有足够的库存！';
						Else
							v_Error:='在处理病人"'||姓名_IN||'"时发现分批或时价药品"'||v_名称||'"没有足够的库存！';
						End IF;
					End IF;
                    Raise Err_Custom;
                End IF;
            ElsIF(v_分批=1 And Nvl(r_Stock.批次,0)=0) Or(v_分批=0 And Nvl(r_Stock.批次,0)<>0) Then 
                Close c_Stock;
                If 医嘱序号_IN IS NULL Then
					If 收费类别_IN='4' Then
						v_Error:='第 '||序号_IN||' 行卫生材料"'||v_名称||'"的分批属性与库存记录不相符,请检查材料数据的正确性！';
					Else
	                    v_Error:='第 '||序号_IN||' 行药品"'||v_名称||'"的分批属性与库存记录不相符,请检查药品数据的正确性！';
					End IF;
                Else
					If 收费类别_IN='4' Then
						v_Error:='在处理病人"'||姓名_IN||'"时发现卫生材料"'||v_名称||'"的分批属性与库存记录不相符,请检查材料数据的正确性！';
					Else
	                    v_Error:='在处理病人"'||姓名_IN||'"时发现药品"'||v_名称||'"的分批属性与库存记录不相符,请检查药品数据的正确性！';
					End IF;
                End IF;
                Raise Err_Custom;
            End IF;

            --确定本次分解数量
            IF v_分批=1 Or v_时价=1 Then
                --对于不分批的时价只可能分解一次,分解不完上面判断了.它分解是为了计算单价.
                --每次分解取小者,库存不够分解不完在上面判断.
                IF v_总数量<=Nvl(r_Stock.可用数量,0) Then
                    v_当前数量:=v_总数量;
                Else
                    v_当前数量:=Nvl(r_Stock.可用数量,0);
                End if;
                IF v_时价=1 Then 
                    If r_Stock.实际数量=0 Then
                        v_当前单价:=0;
                    Else
                        v_当前单价:=Round(Nvl(r_Stock.实际金额/r_Stock.实际数量,0),5);
                    End IF;
                ElsIf v_分批=1 Then
                    v_当前单价:=标准单价_IN;
                End IF;
            Else
                --普通药品
                --不管够不够,程序中已根据参数判断
                v_当前数量:=v_总数量;
                v_当前单价:=标准单价_IN;
            End IF;

            --药品库存(普通情况可能没有记录)
            IF c_Stock%Found Then
                Update 药品库存
                    Set 可用数量=Nvl(可用数量,0)-v_当前数量
                Where 库房ID=执行部门ID_IN And 药品ID=收费细目ID_IN
                    And Nvl(批次,0)=Nvl(r_Stock.批次,0) And 性质=1;
            ElsIf 执行部门ID_IN IS Not NULL Then
                --只有不分批非时价药品可能库存不足出库
                Update 药品库存
                    Set 可用数量=Nvl(可用数量,0)-v_当前数量
                Where 库房ID=执行部门ID_IN And 药品ID=收费细目ID_IN
                    And Nvl(批次,0)=0 And 性质=1;
                IF SQL%RowCount=0 Then
                    Insert Into 药品库存(
                        库房ID,药品ID,性质,可用数量)
                    Values(
                        执行部门ID_IN,收费细目ID_IN,1,-1*v_当前数量);
                End IF;
            End IF;

            --药品收发记录
			v_批次:=Null;v_批号:=Null;
			v_效期:=Null;v_产地:=Null;
			v_灭菌效期:=Null;v_灭菌日期:=Null;
            IF c_Stock%Found Then
                v_批次:=r_Stock.批次;
                v_批号:=r_Stock.上次批号;
                v_效期:=r_Stock.效期;
                v_产地:=r_Stock.上次产地;

				--卫材灭菌效期:一次性材料且有效期
				IF 收费类别_IN='4' Then
					v_Count:=0;
					Begin
						Select 灭菌效期 Into v_Count From 材料特性 Where Nvl(一次性材料,0)=1 And 材料ID=收费细目ID_IN;
					Exception
						When Others Then Null;
					End;
					IF Nvl(v_Count,0)>0 Then
						v_灭菌效期:=r_Stock.灭菌效期;	
						v_灭菌日期:=v_灭菌效期-v_Count*30;
					End IF;
				End IF;
            End IF;

            Select Nvl(Max(序号),0)+1 Into v_序号 From 药品收发记录 
				Where 单据=Decode(收费类别_IN,'4',25,9) And 记录状态=1 And NO=NO_IN;

            --修改的原单据号存放在摘要中
			v_扣率:=NULL;
            If 期效_IN IS Not NULL Or 计价特性_IN IS Not NULL THEN 
                v_扣率:=Nvl(期效_IN,0)||Nvl(计价特性_IN,0);
            End IF;
            Insert Into 药品收发记录(
                ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
                药品ID,批次,产地,批号,效期,付数,填写数量,实际数量,零售价,零售金额,
                摘要,填制人,填制日期,费用ID,频次,单量,用法,外观,扣率,灭菌效期,灭菌日期)
            Values(
                药品收发记录_ID.Nextval,1,Decode(收费类别_IN,'4',25,9),NO_IN,v_序号,执行部门ID_IN,开单部门ID_IN,
                类别ID_IN,-1,收费细目ID_IN,v_批次,v_产地,v_批号,v_效期,Decode(v_分批,1,1,付数_IN),
                Decode(v_分批,1,v_当前数量,v_当前数量/付数_IN),Decode(v_分批,1,v_当前数量,v_当前数量/付数_IN),
                v_当前单价,Round(v_当前单价*v_当前数量,v_Dec),药品摘要_IN,操作员姓名_IN,登记时间_IN,
                v_费用ID,频次_IN,单量_IN,v_用法,v_煎法,v_扣率,v_灭菌效期,v_灭菌日期);

            --未发药品记录
            Update 未发药品记录
                Set 病人ID=病人ID_IN,姓名=姓名_IN
             Where 单据=Decode(收费类别_IN,'4',25,9) And NO=NO_IN 
				And Nvl(库房ID,0)=Nvl(执行部门ID_IN,0);

            IF SQL%RowCount=0 Then
                --取身份优先级
                Begin
                    Select B.优先级 Into v_优先级 From 病人信息 A,身份 B
                     Where A.身份=B.名称(+) And A.病人ID=病人ID_IN;
                Exception
                    When Others Then Null;
                End;

                Insert Into 未发药品记录(
                    单据,NO,病人ID,姓名,优先级,对方部门ID,库房ID,填制日期,已收费,打印状态)
                Values(
                    Decode(收费类别_IN,'4',25,9),NO_IN,病人ID_IN,姓名_IN,v_优先级,
					开单部门ID_IN,执行部门ID_IN,登记时间_IN,Decode(划价_IN,1,0,1),0);
            End IF;

            v_总数量:=v_总数量-v_当前数量;
            v_总金额:=v_总金额+Round(v_当前数量*v_当前单价,v_Dec);
        End Loop;
        
        --可能分批时价药品分解的批次变了
        IF v_时价=1 Then
            IF Round(v_总金额/(付数_IN*数次_IN),5)<>标准单价_IN Then 
                Close c_Stock;    
                If 医嘱序号_IN IS NULL Then
					If 收费类别_IN='4' Then
						v_Error:='第 '||序号_IN||' 行的时价卫生材料"'||v_名称||'"当前计算单价不一致,请重新输入数量计算！';
					Else
	                    v_Error:='第 '||序号_IN||' 行的时价药品"'||v_名称||'"当前计算单价不一致,请重新输入数量计算！';
					End IF;
                Else
					If 收费类别_IN='4' Then
	                    v_Error:='在处理病人"'||姓名_IN||'"时发现时价卫生材料"'||v_名称||'"当前计算的单价发生变化。'||CHR(13)||CHR(10)||'请检查该病人是否同时使用了两笔相同的"'||v_名称||'"！';
					Else
						v_Error:='在处理病人"'||姓名_IN||'"时发现时价药品"'||v_名称||'"当前计算的单价发生变化。'||CHR(13)||CHR(10)||'请检查该病人是否同时使用了两笔相同的"'||v_名称||'"！';
					End IF;
                End IF;
                Raise Err_Custom;
            End IF;
        End IF;
        
        Close c_Stock;
    End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_门诊记帐记录_Insert;
/

-------------------------------------------------------
--模块：门诊收费记录.SQL
Create Or Replace Procedure zl_门诊收费记录_Insert(
    NO_IN				病人费用记录.NO%Type,
    序号_IN             病人费用记录.序号%Type,
    病人ID_IN           病人费用记录.病人ID%Type,
    主页ID_IN           病人费用记录.主页ID%Type,
    标识号_IN           病人费用记录.标识号%Type,
    床号_IN             病人费用记录.床号%Type,
    姓名_IN             病人费用记录.姓名%Type,
    性别_IN             病人费用记录.性别%Type,
    年龄_IN             病人费用记录.年龄%Type,
    费别_IN             病人费用记录.费别%Type,
    加班标志_IN         病人费用记录.加班标志%Type,
    病人病区ID_IN       病人费用记录.病人病区ID%Type,
    病人科室ID_IN       病人费用记录.病人科室ID%Type,
    开单部门ID_IN       病人费用记录.开单部门ID%Type,
    开单人_IN           病人费用记录.开单人%Type,
    从属父号_IN         病人费用记录.从属父号%Type,
    收费细目ID_IN       病人费用记录.收费细目ID%Type,
    收费类别_IN         病人费用记录.收费类别%Type,
    计算单位_IN         病人费用记录.计算单位%Type,
    保险项目否_IN       病人费用记录.保险项目否%Type,
    保险大类ID_IN       病人费用记录.保险大类ID%Type,
    发药窗口_IN         病人费用记录.发药窗口%Type,
    付数_IN             病人费用记录.付数%Type,
    数次_IN             病人费用记录.数次%Type,
    附加标志_IN         病人费用记录.附加标志%Type,
    执行部门ID_IN       病人费用记录.执行部门ID%Type,
    价格父号_IN         病人费用记录.价格父号%Type,
    收入项目ID_IN       病人费用记录.收入项目ID%Type,
    收据费目_IN         病人费用记录.收据费目%Type,
    标准单价_IN         病人费用记录.标准单价%Type,
    应收金额_IN         病人费用记录.应收金额%Type,
    实收金额_IN         病人费用记录.实收金额%Type,
    统筹金额_IN         病人费用记录.统筹金额%Type,
    发生时间_IN         病人费用记录.发生时间%Type,
    登记时间_IN         病人费用记录.登记时间%Type,
    原NO_IN             病人费用记录.NO%Type,
    结帐ID_IN           病人费用记录.结帐ID%Type,
	收费结算_IN			Varchar2,
    冲预交额_IN         病人预交记录.冲预交%Type,
    保险结算_IN         Varchar2,
    操作员编号_IN       病人费用记录.操作员编号%Type,
    操作员姓名_IN       病人费用记录.操作员姓名%Type,
    类别ID_IN           药品单据性质.类别ID%Type:=Null,
    摘要_IN             病人费用记录.摘要%Type:=Null,
    是否急诊_IN         病人费用记录.是否急诊%Type:=0,
    用法_IN                 药品收发记录.用法%Type:=NULL--用法[|煎法]
)
AS
    --功能：新收一张门诊收费单据
    --参数：
    --  主页ID_IN:住院病人收费时用。
    --  原NO_IN:修改保存新单据时用。目前用于存放于药品收发记录的摘要中。
    --         原单据(记录状态=2)记录修改产生的新单据号。
    --         新单据(记录状态=1)记录所修改的原单据号。
	--	收费结算_IN:格式="结算方式|结算金额|结算号码||.....",注意无结算号码要用空格填充
	--	保险结算_IN:格式="结算方式|结算金额||....."

    --该游标用于收费冲预交的可用预交列表(该SQL参考住院结帐)
    --以ID排序，优先冲上次未冲完的。
	--不包含结算方式为代收款项的预交款。
    Cursor c_Deposit(v_病人ID 病人信息.病人ID%Type) is
    Select * From(
        Select A.ID,A.记录状态,A.NO,Nvl(A.金额,0) as 金额
        From 病人预交记录 A,(
                Select NO,Sum(Nvl(A.金额,0)) as 金额 
                From 病人预交记录 A
				Where A.结帐ID is Null And Nvl(A.金额,0)<>0 
					And A.病人ID=v_病人ID
				Group by NO Having Sum(Nvl(A.金额,0))<>0
                ) B
        Where A.结帐ID is Null And Nvl(A.金额,0)<>0 
			And A.结算方式 Not IN(Select 名称 From 结算方式 Where 性质=5)
			And A.NO=B.NO And A.病人ID=v_病人ID
        Union All
        Select 0 as ID,记录状态,NO,Sum(Nvl(金额,0)-Nvl(冲预交,0)) as 金额
        From 病人预交记录
        Where 记录性质 IN(1,11) And 结帐ID is Not NULL 
			And Nvl(金额,0)<>Nvl(冲预交,0) And 病人ID=v_病人ID
        Having Sum(Nvl(金额,0)-Nvl(冲预交,0))<>0
        Group by 记录状态,NO)
    Order by ID,NO;

    v_费用ID		病人费用记录.ID%Type;
    v_优先级		未发药品记录.优先级%Type;
    v_预交金额		病人预交记录.冲预交%Type;

    --药房分批、时价药品--
    ------------------------------------------------------------
    --该游标用于分批药品数量分解
    Cursor c_Stock is
        Select * From 药品库存 
        Where 药品ID=收费细目ID_IN And 库房ID=执行部门ID_IN
            And 性质=1 And(Nvl(批次,0)=0 Or 效期 is Null Or 效期>Trunc(Sysdate))
            And Nvl(可用数量,0)<>0
        Order By Nvl(批次,0);
    r_Stock c_Stock%RowType;
    
    --属性
    v_分批			药品规格.药房分批%Type;
    v_时价			收费项目目录.是否变价%Type;
    v_名称			收费项目目录.名称%Type;
    --临时变量
    v_总数量		Number;
    v_当前数量		Number;
    v_总金额		Number;
    v_当前单价		Number;
    --药品收发记录
    v_批次			药品收发记录.批次%Type;
    v_产地			药品收发记录.产地%Type;
    v_批号			药品收发记录.批号%Type;
    v_效期			药品收发记录.效期%Type;
    v_序号			药品收发记录.序号%Type;
    v_灭菌效期		药品收发记录.灭菌效期%Type;
    v_灭菌日期		药品收发记录.灭菌日期%Type;
    v_煎法			药品收发记录.外观%Type;
    ------------------------------------------------------------
	--结算方式串
	v_结算内容	Varchar2(500);
	v_当前结算	Varchar2(50);
	v_结算方式	病人预交记录.结算方式%Type;
	v_结算金额	病人预交记录.冲预交%Type;
	v_结算号码	病人预交记录.结算号码%Type;

    v_Dec			Number;

    --临时变量
    v_Count			Number;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
   	--药品用法煎法分解
	IF 用法_IN IS Not NULL Then
		IF Instr(用法_IN,'|')>0 Then			
			v_煎法:=Substr(用法_IN,Instr(用法_IN,'|')+1);		
		End IF;
	End IF;

    --金额小数位数
    Begin
        Select To_Number(参数值) Into v_Dec From 系统参数表 Where 参数号=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --病人费用记录
    Select 病人费用记录_ID.Nextval Into v_费用ID From Dual;
    Insert Into 病人费用记录(
        ID,记录性质,NO,记录状态,序号,从属父号,价格父号,门诊标志,病人ID,主页ID,标识号,床号,姓名,性别,
        年龄,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,保险项目否,保险大类ID,付数,数次,发药窗口,加班标志,附加标志,
        收入项目ID,收据费目,标准单价,应收金额,实收金额,统筹金额,记帐费用,划价人,开单部门ID,开单人,发生时间,登记时间,
        执行部门ID,执行状态,结帐ID,结帐金额,操作员编号,操作员姓名,摘要,是否急诊)
    Values(
        v_费用ID,1,NO_IN,1,序号_IN,Decode(从属父号_IN,0,Null,从属父号_IN),Decode(价格父号_IN,0,Null,价格父号_IN),Decode(主页ID_IN,NULL,1,2),
        Decode(病人ID_IN,0,Null,病人ID_IN),主页ID_IN,Decode(标识号_IN,0,Null,标识号_IN),床号_IN,姓名_IN,性别_IN,年龄_IN,病人病区ID_IN,
        病人科室ID_IN,费别_IN,收费类别_IN,收费细目ID_IN,计算单位_IN,保险项目否_IN,保险大类ID_IN,付数_IN,数次_IN,发药窗口_IN,加班标志_IN,
        附加标志_IN,收入项目ID_IN,收据费目_IN,标准单价_IN,应收金额_IN,实收金额_IN,统筹金额_IN,0,操作员姓名_IN,开单部门ID_IN,开单人_IN,
        发生时间_IN,登记时间_IN,执行部门ID_IN,0,结帐ID_IN,实收金额_IN,操作员编号_IN,操作员姓名_IN,摘要_IN,是否急诊_IN);

    IF 序号_IN=1 Then
        --病人预交记录(第一行时处理)
		--正常结算
		IF 收费结算_IN is Not Null Then
			--各个收费结算	
			v_结算内容:=收费结算_IN||'||';
			While v_结算内容 IS Not NULL Loop
				v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
				v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
				v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
				v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
				v_结算号码:=LTrim(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
				
				If Nvl(v_结算金额,0)<>0 Then
					Insert Into 病人预交记录(
						ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,结算号码,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
					Values(
						病人预交记录_ID.Nextval,3,NO_IN,1,Decode(病人ID_IN,0,Null,病人ID_IN),主页ID_IN,'收费结算',
						v_结算方式,v_结算号码,登记时间_IN,操作员编号_IN,操作员姓名_IN,v_结算金额,结帐ID_IN);
				End IF;

				v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
			End Loop;
		End IF;

       --各个保险结算    
	IF 保险结算_IN IS NOT NULL Then 
		v_结算内容:=保险结算_IN||'||';
		While v_结算内容 IS Not NULL Loop
			v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);

			v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
			v_结算金额:=To_Number(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
					
			If Nvl(v_结算金额,0)<>0 Then
				Insert Into 病人预交记录(
					ID,记录性质,NO,记录状态,病人ID,摘要,结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
				Values(
					病人预交记录_ID.Nextval,3,NO_IN,1,Decode(病人ID_IN,0,Null,病人ID_IN),'保险结算',
					v_结算方式,登记时间_IN,操作员编号_IN,操作员姓名_IN,v_结算金额,结帐ID_IN);
			End IF;
			v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
		End Loop;        
	END if;

        --预交结算
        IF Nvl(冲预交额_IN,0)<>0 Then
            IF Nvl(病人ID_IN,0)=0 Then
                v_Error:='不能确定病人病人ID,收费使用预交款结算失败！';
                Raise Err_Custom;
            End if;

            v_预交金额:=冲预交额_IN;
            For r_Deposit IN c_Deposit(病人ID_IN) Loop
                IF r_Deposit.ID<>0 Then
                    --第一次冲预交
                    Update 病人预交记录 
                        Set 冲预交=Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额),
                            结帐ID=结帐ID_IN
                    Where ID=r_Deposit.ID;
                Else
                    --冲上次剩余额
                    INSERT Into 病人预交记录(
                        ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,
                        结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号,收款时间,
                        操作员姓名,操作员编号,冲预交,结帐ID)
                    Select 病人预交记录_ID.Nextval,NO,实际票号,11,记录状态,病人ID,
                         主页ID,科室ID,NULL,结算方式,结算号码,摘要,缴款单位,
                         单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,
                         Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额),结帐ID_IN
                    From 病人预交记录
                    Where NO=r_Deposit.NO And 记录状态=r_Deposit.记录状态
                        And 记录性质 IN(1,11) And RowNum=1;
                End IF;
                --检查是否已经处理完
                IF r_Deposit.金额<v_预交金额 Then
                    v_预交金额:=v_预交金额-r_Deposit.金额;
                Else
                    v_预交金额:=0;
                End IF;
                IF v_预交金额=0 Then 
                    Exit;
                End IF;
            End Loop;
            --检查金额是否足够
            IF v_预交金额>0 Then
                v_Error:='病人的当前预交余额不足金额 '||Ltrim(To_Char(冲预交额_IN,'9999999990.00'))||' ！';
                Raise Err_Custom;
            End IF;

            --更新病人预交余额
            Update 病人余额 Set 预交余额=Nvl(预交余额,0)-冲预交额_IN Where 病人ID=病人ID_IN And 性质=1;
            IF SQL%RowCount=0 Then
                Insert Into 病人余额(病人ID,预交余额,性质) Values(病人ID_IN,-冲预交额_IN,1);
            End IF;
            Delete From 病人余额 Where 病人ID=病人ID_IN And 性质=1 And Nvl(费用余额,0)=0 And Nvl(预交余额,0)=0;
        End IF;
    End IF;

    --相关汇总表的处理
    --汇总"人员缴款余额"(注意要处理个人帐户的结算)
    IF 序号_IN=1 Then
		--各个收费结算	
		IF 收费结算_IN IS Not NULL Then 
			v_结算内容:=收费结算_IN||'||';
			While v_结算内容 IS Not NULL Loop
				v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
				v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
				v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
				v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
				
				If Nvl(v_结算金额,0)<>0 Then
					Update 人员缴款余额
						Set 余额=Nvl(余额,0)+Nvl(v_结算金额,0)
					 Where 收款员=操作员姓名_IN And 性质=1
						And 结算方式=v_结算方式;

					IF SQL%RowCount=0 Then
						Insert Into 人员缴款余额(
							收款员,结算方式,性质,余额)
						Values(
							操作员姓名_IN,v_结算方式,1,Nvl(v_结算金额,0));
					End IF;
				End IF;

				v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
			End Loop;
		End IF;

        --各个保险结算    
        IF 保险结算_IN IS Not NULL Then 
            v_结算内容:=保险结算_IN||'||';
            While v_结算内容 IS Not NULL Loop
                v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);

                v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
                v_结算金额:=To_Number(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
				
				If Nvl(v_结算金额,0)<>0 Then
					Update 人员缴款余额
						Set 余额=Nvl(余额,0)+Nvl(v_结算金额,0)
					 Where 收款员=操作员姓名_IN And 性质=1
						And 结算方式=v_结算方式;

					IF SQL%RowCount=0 Then
						Insert Into 人员缴款余额(
							收款员,结算方式,性质,余额)
						Values(
							操作员姓名_IN,v_结算方式,1,Nvl(v_结算金额,0));
					End IF;
				End IF;                

                v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
            End Loop;
        End IF;
        Delete From 人员缴款余额 Where 性质=1 And 收款员=操作员姓名_IN And Nvl(余额,0)=0;
    End IF;

    --病人费用汇总
    Update 病人费用汇总
        Set 应收金额=Nvl(应收金额,0)+应收金额_IN,
            实收金额=Nvl(实收金额,0)+实收金额_IN,
            结帐金额=Nvl(结帐金额,0)+实收金额_IN
    Where 日期=Trunc(登记时间_IN)
        And Nvl(病人病区ID,0)=Nvl(病人病区ID_IN,0)
        And Nvl(病人科室ID,0)=Nvl(病人科室ID_IN,0)
        And Nvl(开单部门ID,0)=Nvl(开单部门ID_IN,0)
        And Nvl(执行部门ID,0)=Nvl(执行部门ID_IN,0)
        And 收入项目ID+0=收入项目ID_IN
        And 来源途径=Decode(主页ID_IN,NULL,1,2) And 记帐费用=0;

    IF SQL%RowCount=0 Then
        Insert Into 病人费用汇总(
            日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
        Values(
            Trunc(登记时间_IN),病人病区ID_IN,病人科室ID_IN,开单部门ID_IN,执行部门ID_IN,收入项目ID_IN,
            Decode(主页ID_IN,NULL,1,2),0,应收金额_IN,实收金额_IN,实收金额_IN);
    End IF;

    --药品和卫生材料部分
	v_Count:=0;
	If 收费类别_IN='4' Then--跟踪在用的卫材才处理
		Select 跟踪在用 Into v_Count From 材料特性 Where 材料ID=收费细目ID_IN;
	End IF;
    IF 收费类别_IN in('5','6','7') Or (收费类别_IN='4' And Nvl(v_Count,0)=1) Then
		If 收费类别_IN='4' Then
			Select Nvl(A.在用分批,0),Nvl(B.是否变价,0),B.名称 
				Into v_分批,v_时价,v_名称
			From 材料特性 A,收费项目目录 B
			Where A.材料ID=B.ID And B.ID=收费细目ID_IN;
		Else
			Select Nvl(A.药房分批,0),Nvl(B.是否变价,0),B.名称 
				Into v_分批,v_时价,v_名称
			From 药品规格 A,收费项目目录 B
			Where A.药品ID=B.ID And B.ID=收费细目ID_IN;
		End IF;

        v_总数量:=付数_IN*数次_IN;
        v_总金额:=0;
        Open c_Stock;

        While v_总数量<>0 Loop
            Fetch c_Stock Into r_Stock;
            IF c_Stock%NotFound Then
                --第一次就没有库存,分批或时价都不允许。
                --分批药品数量分解不完,也就是库存不足。
                IF v_分批=1 Or v_时价=1 Then
                    Close c_Stock;
					If 收费类别_IN='4' Then
						v_Error:='第 '||序号_IN||' 行的分批或时价卫生材料"'||v_名称||'"没有可用的库存！';
					Else
	                    v_Error:='第 '||序号_IN||' 行的分批或时价药品"'||v_名称||'"没有可用的药品库存！';
					End IF;
                    Raise Err_Custom;
                End IF;
            ElsIF(v_分批=1 And Nvl(r_Stock.批次,0)=0) Or(v_分批=0 And Nvl(r_Stock.批次,0)<>0) Then 
                Close c_Stock;
				If 收费类别_IN='4' Then
					v_Error:='第 '||序号_IN||' 行卫生材料"'||v_名称||'"的在用分批属性与库存记录不相符,请检查材料数据的正确性！';
				Else
	                v_Error:='第 '||序号_IN||' 行药品"'||v_名称||'"的分批属性与库存记录不相符,请检查药品数据的正确性！';
				End IF;
                Raise Err_Custom;
            End IF;

            --确定本次分解数量
            IF v_分批=1 Or v_时价=1 Then
                --对于不分批的时价只可能分解一次,分解不完上面判断了.它分解是为了计算单价.
                --每次分解取小者,库存不够分解不完在上面判断.
                IF v_总数量<=Nvl(r_Stock.可用数量,0) Then
                    v_当前数量:=v_总数量;
                Else
                    v_当前数量:=Nvl(r_Stock.可用数量,0);
                End if;
                IF v_时价=1 Then 
                    If r_Stock.实际数量=0 Then
                        v_当前单价:=0;
                    Else
                        v_当前单价:=Round(Nvl(r_Stock.实际金额/r_Stock.实际数量,0),5);
                    End IF;
                ElsIf v_分批=1 Then
                    v_当前单价:=标准单价_IN;
                End IF;
            Else
                --普通药品
                --不管够不够,程序中已根据参数判断
                v_当前数量:=v_总数量;
                v_当前单价:=标准单价_IN;
            End IF;

            --药品库存(普通情况可能没有记录)
            IF c_Stock%Found Then
                Update 药品库存
                    Set 可用数量=Nvl(可用数量,0)-v_当前数量
                Where 库房ID=执行部门ID_IN And 药品ID=收费细目ID_IN
                    And Nvl(批次,0)=Nvl(r_Stock.批次,0) And 性质=1;
            Elsif 执行部门ID_IN IS Not NULL Then
                --只有不分批非时价药品可能库存不足出库
                Update 药品库存
                    Set 可用数量=Nvl(可用数量,0)-v_当前数量
                Where 库房ID=执行部门ID_IN And 药品ID=收费细目ID_IN
                    And Nvl(批次,0)=0 And 性质=1;
                IF SQL%RowCount=0 Then
                    Insert Into 药品库存(
                        库房ID,药品ID,性质,可用数量)
                    Values(
                        执行部门ID_IN,收费细目ID_IN,1,-1*v_当前数量);
                End IF;
            End IF;

            --药品收发记录
			v_批次:=Null;v_批号:=Null;
			v_效期:=Null;v_产地:=Null;
			v_灭菌效期:=Null;v_灭菌日期:=Null;
            IF c_Stock%Found Then
                v_批次:=r_Stock.批次;
                v_批号:=r_Stock.上次批号;
                v_效期:=r_Stock.效期;
                v_产地:=r_Stock.上次产地;
				
				--卫材灭菌效期:一次性材料且有效期
				IF 收费类别_IN='4' Then
					v_Count:=0;
					Begin
						Select 灭菌效期 Into v_Count From 材料特性 Where Nvl(一次性材料,0)=1 And 材料ID=收费细目ID_IN;
					Exception
						When Others Then Null;
					End;
					IF Nvl(v_Count,0)>0 Then
						v_灭菌效期:=r_Stock.灭菌效期;	
						v_灭菌日期:=v_灭菌效期-v_Count*30;
					End IF;
				End IF;
            End IF;

            Select Nvl(Max(序号),0)+1 Into v_序号 From 药品收发记录 
				Where 单据=Decode(收费类别_IN,'4',24,8) And 记录状态=1 And NO=NO_IN;

            --修改的原单据号存放在摘要中
			--注意卫材单据与药品单据不同
            Insert Into 药品收发记录(
                ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
                药品ID,批次,产地,批号,效期,付数,填写数量,实际数量,零售价,零售金额,
                摘要,填制人,填制日期,费用ID,发药窗口,灭菌效期,灭菌日期,外观)
            Values(
                药品收发记录_ID.Nextval,1,Decode(收费类别_IN,'4',24,8),NO_IN,v_序号,执行部门ID_IN,开单部门ID_IN,
                类别ID_IN,-1,收费细目ID_IN,v_批次,v_产地,v_批号,v_效期,Decode(v_分批,1,1,付数_IN),
                Decode(v_分批,1,v_当前数量,v_当前数量/付数_IN),Decode(v_分批,1,v_当前数量,v_当前数量/付数_IN),
                v_当前单价,Round(v_当前单价*v_当前数量,v_Dec),原NO_IN,操作员姓名_IN,登记时间_IN,v_费用ID,
                发药窗口_IN,v_灭菌效期,v_灭菌日期,v_煎法);

            --未发药品记录:可能同一个库房,但一个为药品,一个为卫材,插入两条记录。
            Update 未发药品记录
                Set 病人ID=Decode(病人ID_IN,0,Null,病人ID_IN),
                    主页ID=主页ID_IN,姓名=姓名_IN,
					发药窗口=Nvl(发药窗口_IN,发药窗口)--可能药品和材料用同一个库房,但材料无发药窗口
             Where 单据=Decode(收费类别_IN,'4',24,8) 
				And NO=NO_IN And Nvl(库房ID,0)=Nvl(执行部门ID_IN,0);

            IF SQL%RowCount=0 Then
                --取身份优先级
                IF Nvl(病人ID_IN,0)<>0 And v_优先级 is Null Then
                    Begin
                        Select B.优先级 Into v_优先级 From 病人信息 A,身份 B
                         Where A.身份=B.名称(+) And A.病人ID=病人ID_IN;
                    Exception
                        When Others Then Null;
                    End;
                End IF;

                Insert Into 未发药品记录(
                    单据,NO,病人ID,主页ID,姓名,优先级,对方部门ID,库房ID,发药窗口,填制日期,已收费,打印状态)
                Values(
                    Decode(收费类别_IN,'4',24,8),NO_IN,Decode(病人ID_IN,0,Null,病人ID_IN),主页ID_IN,姓名_IN,
					v_优先级,开单部门ID_IN,执行部门ID_IN,发药窗口_IN,登记时间_IN,1,0);
            End IF;

            v_总数量:=v_总数量-v_当前数量;
            v_总金额:=v_总金额+Round(v_当前数量*v_当前单价,v_Dec);
        End Loop;
        
        --可能分批时价药品分解的批次变了
        IF v_时价=1 Then
            IF Round(v_总金额/(付数_IN*数次_IN),5)<>标准单价_IN Then 
                Close c_Stock;  
				If 收费类别_IN='4' Then
					v_Error:='第 '||序号_IN||' 行的时价卫生材料"'||v_名称||'"当前计算单价不一致,请重新输入数量计算！';
				Else
	                v_Error:='第 '||序号_IN||' 行的时价药品"'||v_名称||'"当前计算单价不一致,请重新输入数量计算！';
				End IF;
                Raise Err_Custom;
            End IF;
        End IF;

        Close c_Stock;
    End IF;

	--更新部份病人信息
	If 序号_IN=1 And 病人ID_IN IS Not NULL Then
		Update 病人信息 
			Set 性别=Nvl(性别_IN,性别),
				年龄=Nvl(年龄_IN,年龄)				
		Where 病人ID=病人ID_IN;
		UPDATE 病人信息 SET 费别=Nvl(费别_IN,费别) WHERE 病人ID=病人ID_IN AND NOT Exists (SELECT 'X' FROM 费别 WHERE 名称=费别_IN AND 属性=2);
		If 床号_IN IS Not NULL  And 主页ID_IN IS NULL  Then
			Update 病人信息 Set 医疗付款方式=(Select 名称 From 医疗付款方式 Where 编码=床号_IN) Where 病人ID=病人ID_IN;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_门诊收费记录_Insert;
/

Create Or Replace Procedure zl_划价收费记录_Insert(
    NO_IN				病人费用记录.NO%Type,
    病人ID_IN			病人费用记录.病人ID%Type,
    主页ID_IN			病人费用记录.主页ID%Type,
    床号_IN				病人费用记录.床号%Type,
    姓名_IN				病人费用记录.姓名%Type,
    性别_IN				病人费用记录.性别%Type,
    年龄_IN				病人费用记录.年龄%Type,
    病人病区ID_IN		病人费用记录.病人病区ID%Type,
    病人科室ID_IN		病人费用记录.病人科室ID%Type,
    开单部门ID_IN		病人费用记录.开单部门ID%Type,
    开单人_IN			病人费用记录.开单人%Type,
    收费结算_IN			Varchar2,
    冲预交额_IN			病人预交记录.冲预交%Type,
    保险结算_IN			Varchar2,
    结帐ID_IN			病人费用记录.结帐ID%Type,
    发生时间_IN			病人费用记录.发生时间%Type,
    操作员编号_IN		病人费用记录.操作员编号%Type,
    操作员姓名_IN		病人费用记录.操作员姓名%Type,
    发药窗口_IN			病人费用记录.发药窗口%Type:=Null,
    是否急诊_IN			病人费用记录.是否急诊%Type:=0
) AS
     --功能：用于收费时收取划价单费用    
     --参数：
     --        收费结算_IN:格式="结算方式|结算金额|结算号码||.....",注意无结算号码要用空格填充
     --        保险结算_IN:格式="结算方式|结算金额||....."
     --说明：
     --        1.收取划价费用时,才计算费用相关汇总,在划价时不处理;但药品相关汇总(姓名除外)划价时已经计算。
     --        2.收取划价费用时,目前界面及过程中未处理加收工本费,由划价时直接处理。
    --该游标为划价原单据内容
    Cursor c_Price is
        Select * From 病人费用记录
        Where NO=NO_IN And 记录性质=1 And 记录状态=0 And 操作员姓名 is Null
        Order by 序号;
    r_PriceRow c_Price%RowType;

    --该游标用于收费冲预交的可用预交列表(该SQL参考住院结帐)
    --以ID排序，优先冲上次未冲完的。
    Cursor c_Deposit(v_病人ID 病人信息.病人ID%Type) is
        Select * From(
            Select A.ID,A.记录状态,A.NO,Nvl(A.金额,0) as 金额
            From 病人预交记录 A,(
                    Select NO,Sum(Nvl(A.金额,0)) as 金额 
                    From 病人预交记录 A
                Where A.结帐ID Is Null And Nvl(A.金额,0)<>0 And A.病人ID=v_病人ID
                  Group by NO Having Sum(Nvl(A.金额,0))<>0
                    ) B
            Where A.结帐ID Is Null And Nvl(A.金额,0)<>0 And A.NO=B.NO And A.病人ID=v_病人ID
            Union All
            Select 0 as ID,记录状态,NO,Sum(Nvl(金额,0)-Nvl(冲预交,0)) as 金额
            From 病人预交记录
            Where 记录性质 IN(1,11) And 结帐ID is Not NULL And Nvl(金额,0)<>Nvl(冲预交,0) And 病人ID=v_病人ID
            Having Sum(Nvl(金额,0)-Nvl(冲预交,0))<>0
            Group by 记录状态,NO)
        Order by ID,NO;

    --该游标用于病人费用汇总处理
    Cursor c_Money is
        Select TRUNC(登记时间) as 日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,门诊标志,
             Sum(应收金额) as 应收金额,Sum(实收金额) as 实收金额,Sum(结帐金额) as 结帐金额
        From 病人费用记录
        Where 记录性质=1 And 记录状态=1 And NO=NO_IN
        Group by TRUNC(登记时间),病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,门诊标志;
    r_MoneyRow c_Money%RowType;

    --预交与结算相关变量
    v_预交金额		病人预交记录.冲预交%Type;
    v_结算内容		Varchar2(500);
    v_当前结算		Varchar2(50);
    v_结算方式		病人预交记录.结算方式%Type;
    v_结算金额		病人预交记录.冲预交%Type;
    v_结算号码		病人预交记录.结算号码%Type;

    v_标识号		病人费用记录.标识号%Type;

    --临时变量
    v_Count			NUMBER;
    v_Date			DATE;
    v_Error			VARCHAR2(255);
    Err_Custom		EXCEPTION;
BEGIN
    Select Count(ID) Into v_Count From 病人费用记录 Where 记录性质=1 And 记录状态=0 And NO=NO_IN And 操作员姓名 is Null;
    If v_Count=0 Then
        v_Error:='不能读取划价单内容,该单据可能已经删除或已经收费！';
        Raise Err_Custom;
    End If;
   
    Select Sysdate Into v_Date From Dual;
	
  	If Nvl(病人ID_IN,0)<>0 Then
  		Select 
  			Decode(当前科室ID,NULL,门诊号,住院号) Into v_标识号
  		From 病人信息
  		Where 病人ID=病人ID_IN;
  	End IF;

    --循环处理病人费用记录
    For r_PriceRow In c_Price Loop
        --执行状态相关字段不处理,在划价时处理;因为可能未收费发药,这种已执行的划价单是允许收费操作的。
        --为保证与预交结算记录的时间相同,重新填写登记时间,但药品部分不变动。
        Update 病人费用记录
            Set 记录状态=1,
                病人ID=Decode(病人ID_IN,0,Null,病人ID_IN),
                主页ID=主页ID_IN,
				标识号=v_标识号,
                床号=床号_IN,
                姓名=姓名_IN,
                年龄=年龄_IN,
                性别=性别_IN,
                病人病区ID=Nvl(病人病区ID_IN,病人病区ID),--可能保持医嘱发送的内容
                病人科室ID=Nvl(病人科室ID_IN,病人科室ID),
                开单部门ID=Nvl(开单部门ID_IN,开单部门ID),
                开单人=Nvl(开单人_IN,开单人),
                结帐金额=实收金额,
                结帐ID=结帐ID_IN,
                发生时间=发生时间_IN,
                登记时间=v_Date,
                操作员编号=操作员编号_IN,
                操作员姓名=操作员姓名_IN,
                发药窗口=Nvl(发药窗口_IN,发药窗口),
                是否急诊=是否急诊_IN
        Where ID=r_Pricerow.ID;
    End Loop;

    --预交款相关结算
    --收费结算
    If 收费结算_IN IS Not NULL Then
        v_结算内容:=收费结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
			v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
			v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
			v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
			v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
			v_结算号码:=LTrim(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
			
			If Nvl(v_结算金额,0)<>0 Then
				Insert Into 病人预交记录(
					ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,结算号码,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
				Values(
					病人预交记录_ID.Nextval,3,NO_IN,1,Decode(病人ID_IN,0,Null,病人ID_IN),主页ID_IN,
					'收费结算',v_结算方式,v_结算号码,v_Date,操作员编号_IN,操作员姓名_IN,v_结算金额,结帐ID_IN);
			End IF;

            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
     End IF;

     IF 保险结算_IN IS NOT NULL Then 
        --各个保险结算    
        v_结算内容:=保险结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
            v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);

            v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
            v_结算金额:=To_Number(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
			
			If Nvl(v_结算金额,0)<>0 Then
				Insert Into 病人预交记录(
					ID,记录性质,NO,记录状态,病人ID,摘要,结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
				Values(
					病人预交记录_ID.Nextval,3,NO_IN,1,Decode(病人ID_IN,0,Null,病人ID_IN),'保险结算',
					v_结算方式,v_Date,操作员编号_IN,操作员姓名_IN,v_结算金额,结帐ID_IN);
			End IF;

            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
    End IF;

    --预交结算
    IF Nvl(冲预交额_IN,0)<>0 THEN
        IF Nvl(病人ID_IN,0)=0 Then
            v_Error:='不能确定病人病人ID,收费使用预交款结算失败！';
            Raise Err_Custom;
        End if;

        v_预交金额:=冲预交额_IN;
        For r_Deposit IN c_Deposit(病人ID_IN) Loop
            IF r_Deposit.ID<>0 Then
                --第一次冲预交
                Update 病人预交记录 
                    Set 冲预交=Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额),
                        结帐ID=结帐ID_IN
                Where ID=r_Deposit.ID;
            Else
                --冲上次剩余额
                INSERT Into 病人预交记录(
                    ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,
                    结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号,收款时间,
                    操作员姓名,操作员编号,冲预交,结帐ID)
                Select 病人预交记录_ID.Nextval,NO,实际票号,11,记录状态,病人ID,
                     主页ID,科室ID,NULL,结算方式,结算号码,摘要,缴款单位,
                     单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,
                     Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额),结帐ID_IN
                From 病人预交记录
                Where NO=r_Deposit.NO And 记录状态=r_Deposit.记录状态
                    AND 记录性质 IN(1,11) And RowNum=1;
            End IF;
            --检查是否已经处理完
            IF r_Deposit.金额<v_预交金额 Then
                v_预交金额:=v_预交金额-r_Deposit.金额;
            Else
                v_预交金额:=0;
            End IF;
            IF v_预交金额=0 Then 
                Exit;
            End IF;
        End Loop;
        --检查金额是否足够
        IF v_预交金额>0 Then
            v_Error:='病人的当前预交余额不足金额 '||Ltrim(To_Char(冲预交额_IN,'9999999990.00'))||' ！';
            Raise Err_Custom;
        End IF;

        --更新病人预交余额
        Update 病人余额 Set 预交余额=Nvl(预交余额,0)-冲预交额_IN Where 病人ID=病人ID_IN And 性质=1;
        IF SQL%RowCount=0 Then
            Insert Into 病人余额(病人ID,预交余额,性质) Values(病人ID_IN,-冲预交额_IN,1);
        End IF;
        Delete From 病人余额 Where 病人ID=病人ID_IN And 性质=1 And Nvl(费用余额,0)=0 And Nvl(预交余额,0)=0;
    End IF;

    --相关汇总表的处理

    --汇总"人员缴款余额"
	--收费结算
    IF 收费结算_IN IS Not NULL Then
        v_结算内容:=收费结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
			v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
			v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
			v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
			v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
			
			If Nvl(v_结算金额,0)<>0 Then
				Update 人员缴款余额
					Set 余额=Nvl(余额,0)+Nvl(v_结算金额,0)
				 Where 收款员=操作员姓名_IN
					And 性质=1 And 结算方式=v_结算方式;
				If SQL%RowCount=0 Then
					Insert Into 人员缴款余额(
						收款员,结算方式,性质,余额)
					Values(
						操作员姓名_IN,v_结算方式,1,Nvl(v_结算金额,0));
				End If;
			End IF;

            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
    End IF;

    --各个保险结算
    IF 保险结算_IN IS Not NULL Then
        v_结算内容:=保险结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
            v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);

            v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
            v_结算金额:=To_Number(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
			
			If Nvl(v_结算金额,0)<>0 Then
				Update 人员缴款余额
					Set 余额=Nvl(余额,0)+Nvl(v_结算金额,0)
				 Where 收款员=操作员姓名_IN
					And 性质=1 And 结算方式=v_结算方式;
				If SQL%RowCount=0 Then
					Insert Into 人员缴款余额(
						收款员,结算方式,性质,余额)
					Values(
						操作员姓名_IN,v_结算方式,1,Nvl(v_结算金额,0));
				End If;
			End IF;

            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
    End IF;
    Delete From 人员缴款余额 Where 性质=1 And 收款员=操作员姓名_IN And Nvl(余额,0)=0;

    --病人费用汇总
    For r_MoneyRow In c_Money Loop
        Update 病人费用汇总
            Set 应收金额=Nvl(应收金额,0)+r_MoneyRow.应收金额,
                实收金额=Nvl(实收金额,0)+r_MoneyRow.实收金额,
                结帐金额=Nvl(结帐金额,0)+r_MoneyRow.结帐金额
         Where 日期=r_MoneyRow.日期
            And Nvl(病人病区ID,0)=Nvl(r_MoneyRow.病人病区ID,0)
            And Nvl(病人科室ID,0)=Nvl(r_MoneyRow.病人科室ID,0)
            And Nvl(开单部门ID,0)=Nvl(r_MoneyRow.开单部门ID,0)
            And Nvl(执行部门ID,0)=Nvl(r_MoneyRow.执行部门ID,0)
            And 收入项目ID+0=r_MoneyRow.收入项目ID
            And 来源途径=r_MoneyRow.门诊标志 And 记帐费用=0;

        If SQL%RowCount=0 Then
            Insert Into 病人费用汇总(
                日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
            Values(
                r_MoneyRow.日期,r_MoneyRow.病人病区ID,r_MoneyRow.病人科室ID,r_MoneyRow.开单部门ID,r_MoneyRow.执行部门ID,
                r_MoneyRow.收入项目ID,r_MoneyRow.门诊标志,0,r_MoneyRow.应收金额,r_MoneyRow.实收金额,r_MoneyRow.结帐金额);
        End If;
    End Loop;

    --药品部分非费用信息的修改
    --药品未发记录(如果已发药则修改不到),分离发药时无库房ID
	--可能存在材料和药品库房相同，但材料无发药窗口
    Update 未发药品记录
        Set 病人ID=Decode(病人ID_IN,0,Null,病人ID_IN),
            主页ID=主页ID_IN,姓名=姓名_IN,
            对方部门ID=开单部门ID_IN,已收费=1,填制日期=v_Date
     Where 单据=24 And NO=NO_IN And Nvl(库房ID,0) IN(
		Select Distinct Nvl(执行部门ID,0) From 病人费用记录 
			Where 记录性质=1 And 记录状态=1 And NO=NO_IN And 收费类别='4');
	
	Update 未发药品记录
        Set 病人ID=Decode(病人ID_IN,0,Null,病人ID_IN),
            主页ID=主页ID_IN,姓名=姓名_IN,
            对方部门ID=开单部门ID_IN,已收费=1,
            发药窗口=Nvl(发药窗口_IN,发药窗口),填制日期=v_Date
     Where 单据=8 And NO=NO_IN And Nvl(库房ID,0) IN(
		Select Distinct Nvl(执行部门ID,0) From 病人费用记录 
			Where 记录性质=1 And 记录状态=1 And NO=NO_IN And 收费类别 IN('5','6','7'));

    --药品收发记录(可能已经发药或取消发药,所有记录更改)
    Update 药品收发记录
        Set 对方部门ID=开单部门ID_IN,填制日期=Decode(Sign(Nvl(审核日期,v_Date)-v_Date),-1,填制日期,v_Date)
     Where 单据=24 And NO=NO_IN And 费用ID+0 IN(
		Select ID From 病人费用记录 
			Where 记录性质=1 And 记录状态=1 And NO=NO_IN And 收费类别='4');

	Update 药品收发记录
        Set 对方部门ID=开单部门ID_IN,发药窗口=Nvl(发药窗口_IN,发药窗口),填制日期=Decode(Sign(Nvl(审核日期,v_Date)-v_Date),-1,填制日期,v_Date)
     Where 单据=8 And NO=NO_IN And 费用ID+0 IN(
		Select ID From 病人费用记录 
			Where 记录性质=1 And 记录状态=1 And NO=NO_IN And 收费类别 IN('5','6','7'));

	--更新部份病人信息
	If 病人ID_IN IS Not NULL Then
		Update 病人信息 
			Set 性别=Nvl(性别_IN,性别),
				年龄=Nvl(年龄_IN,年龄)
		Where 病人ID=病人ID_IN;
		If 床号_IN IS Not NULL  And 主页ID_IN IS NULL Then
			Update 病人信息 Set 医疗付款方式=(Select 名称 From 医疗付款方式 Where 编码=床号_IN) Where 病人ID=病人ID_IN;
		End IF;
	End IF;
EXCEPTION
    WHEN Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS Then zl_ErrOrCenter(SQLCODE,SQLERRM);
End zl_划价收费记录_Insert;
/

Create Or Replace Procedure zl_门诊收费误差_Insert(
--功能：填写门诊收费时所产生的误差费用,以保证与结算金额一致。
--      用于门诊收费,门诊收取划价单费用,门诊退费。
--参数：误差_IN=Sum(结算金额)-Sum(结帐金额)
--      退费_IN=表明是否门诊退费时调用(仅单张单据部份退费时才调用)
    NO_IN			病人费用记录.NO%Type,
    误差_IN         病人费用记录.实收金额%Type,
    退费_IN         Number:=0
) AS
    v_收费类别      病人费用记录.收费类别%Type;
    v_收费细目ID    病人费用记录.收费细目ID%Type;
    v_计算单位      病人费用记录.计算单位%Type;
    v_收入项目ID    病人费用记录.收入项目ID%Type;
    v_收据费目      病人费用记录.收据费目%Type;
    
    v_病人病区ID    病人费用记录.病人病区ID%Type;
    v_病人科室ID    病人费用记录.病人科室ID%Type;
    v_开单部门ID    病人费用记录.开单部门ID%Type;
    v_执行部门ID    病人费用记录.执行部门ID%Type;
    v_登记时间      病人费用记录.登记时间%Type;
    v_病人来源		病人费用记录.门诊标志%Type;

    v_费用ID		病人费用记录.ID%Type;
    v_序号			病人费用记录.序号%Type;
    v_结帐ID		病人费用记录.结帐ID%Type;
    v_执行状态		病人费用记录.执行状态%Type;

    v_Temp			Varchar2(255);
    v_人员编号		病人费用记录.操作员编号%Type;
    v_人员姓名		病人费用记录.操作员姓名%Type;

    v_Sign			Number;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
    If Nvl(误差_IN,0)=0 THEN 
        Return;
    End IF;

    --操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp:=zl_Identity;
    v_执行部门ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --误差项目内容
    Begin
        Select A.类别,A.ID,A.计算单位,C.ID,C.收据费目 
            Into v_收费类别,v_收费细目ID,v_计算单位,v_收入项目ID,v_收据费目
        From 收费项目目录 A,收费价目 B,收入项目 C,收费特定项目 D
        Where D.特定项目='误差项' And D.收费细目ID=A.ID
            And A.ID=B.收费细目ID And B.收入项目ID=C.ID
            And ((Sysdate Between B.执行日期 And B.终止日期) 
                Or (Sysdate>=B.执行日期 And B.终止日期 is NULL));
    Exception
        When Others Then
        Begin
            v_Error:='不能正确读取处理费用误差的项目信息，请先检查该项目是否正确设置。';
            Raise Err_Custom;
        End;
    End;
    
    If Nvl(退费_IN,0)<>0 Then
        --退费处理误差时,在收费的误差记录上处理；如果无,则直接新增退费记录
        Begin
            Select 序号 Into v_序号 From 病人费用记录 Where NO=NO_IN And 记录性质=1 And 记录状态 IN(1,3) And 附加标志=9;
        Exception
            When Others Then NULL;
        End;
    End IF;
    If v_序号 IS NULL Then
        Select Max(序号)+1 Into v_序号 From 病人费用记录 Where NO=NO_IN And 记录性质=1;    
    End IF;

    v_Sign:=1;v_执行状态:=0;
    --该笔项目第几次退费(退费时)
    If Nvl(退费_IN,0)<>0 Then
        v_Sign:=-1;
        Select -1*(Nvl(Max(Abs(执行状态)),0)+1) Into v_执行状态
        From 病人费用记录 
        Where NO=NO_IN And 记录性质=1 And 记录状态=2 And 序号=v_序号;
    End IF;

    --取最近收费或退费的结帐ID(主要是为了确定最后一次退费)
    Select Max(结帐ID) Into v_结帐ID From 病人费用记录 Where NO=NO_IN And 记录性质=1;
    Select 病人费用记录_ID.Nextval Into v_费用ID From Dual;

    --病人费用记录:附加标志=9
    Insert Into 病人费用记录(
        ID,记录性质,NO,实际票号,记录状态,序号,从属父号,价格父号,门诊标志,病人ID,标识号,床号,姓名,性别,
        年龄,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,发药窗口,付数,数次,加班标志,附加标志,
        收入项目ID,收据费目,标准单价,应收金额,实收金额,记帐费用,划价人,开单部门ID,开单人,发生时间,登记时间,
        执行部门ID,执行状态,结帐ID,结帐金额,操作员编号,操作员姓名,是否上传)
    Select
        v_费用ID,记录性质,NO,实际票号,记录状态,v_序号,NULL,NULL,门诊标志,病人ID,标识号,床号,姓名,性别,年龄,
        病人病区ID,病人科室ID,费别,v_收费类别,v_收费细目ID,v_计算单位,发药窗口,1,v_Sign*1,加班标志,9,
        v_收入项目ID,v_收据费目,误差_IN,v_Sign*误差_IN,v_Sign*误差_IN,记帐费用,划价人,开单部门ID,开单人,发生时间,登记时间,
        v_执行部门ID,v_执行状态,结帐ID,v_Sign*误差_IN,v_人员编号,v_人员姓名,1
    From 病人费用记录
    Where NO=NO_IN And 记录性质=1 And 记录状态=Decode(Nvl(退费_IN,0),0,1,2) And 结帐ID=v_结帐ID And Rownum=1;

    If Nvl(退费_IN,0)<>0 Then
        Update 病人费用记录 Set 记录状态=3 Where 记录性质=1 And NO=NO_IN And 记录状态=1 And 序号=v_序号;
    End IF;

    --病人费用汇总
    Select 病人病区ID,病人科室ID,开单部门ID,登记时间,门诊标志
        Into v_病人病区ID,v_病人科室ID,v_开单部门ID,v_登记时间,v_病人来源
    From 病人费用记录 Where ID=v_费用ID;
    Update 病人费用汇总
        Set 应收金额=Nvl(应收金额,0)+v_Sign*误差_IN,
            实收金额=Nvl(实收金额,0)+v_Sign*误差_IN,
            结帐金额=Nvl(结帐金额,0)+v_Sign*误差_IN
    Where 日期=Trunc(v_登记时间)
        And Nvl(病人病区ID,0)=Nvl(v_病人病区ID,0)
        And Nvl(病人科室ID,0)=Nvl(v_病人科室ID,0)
        And Nvl(开单部门ID,0)=Nvl(v_开单部门ID,0)
        And Nvl(执行部门ID,0)=Nvl(v_执行部门ID,0)
        And 收入项目ID+0=v_收入项目ID
        And 来源途径=v_病人来源 And 记帐费用=0;
    IF SQL%RowCount=0 Then
        Insert Into 病人费用汇总(
            日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
        Values(
            Trunc(v_登记时间),v_病人病区ID,v_病人科室ID,v_开单部门ID,v_执行部门ID,v_收入项目ID,v_病人来源,0,v_Sign*误差_IN,v_Sign*误差_IN,v_Sign*误差_IN);
    End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_门诊收费误差_Insert;
/

Create Or Replace Procedure zl_门诊收费结算_Update(
	   结帐Id_IN			病人费用记录.结帐ID%Type,
	   收费结算_IN			Varchar2,
	   冲预交额_IN			病人预交记录.冲预交%Type,
	   保险结算_IN			Varchar2,
	   误差_IN				病人费用记录.实收金额%Type
 )
 As 
 --功能:处理收费时和医保正式结算后,相关结算信息的调整
 --     因为预结算后,生成的医保结算金额总额及分摊可能会与正式结算时有差异,所以提供了校对功能,
 --		操作员在结算校对时,可以调整非医保结算方式的各种结算金额及方式,重新生成结算串,并且可能产生误差金额.
  
  --该游标用于收费冲预交的可用预交列表(该SQL参考住院结帐)
    --以ID排序，优先冲上次未冲完的。
    Cursor c_Deposit(v_病人ID 病人信息.病人ID%Type) is
      Select * From(
          Select A.ID,A.记录状态,A.NO,Nvl(A.金额,0) as 金额
          From 病人预交记录 A,(
                  Select NO,Sum(Nvl(A.金额,0)) as 金额 
                  From 病人预交记录 A
              Where A.结帐ID Is Null And Nvl(A.金额,0)<>0 And A.病人ID=v_病人ID
                Group by NO Having Sum(Nvl(A.金额,0))<>0
                  ) B
          Where A.结帐ID Is Null And Nvl(A.金额,0)<>0 And A.NO=B.NO And A.病人ID=v_病人ID
          Union All
          Select 0 as ID,记录状态,NO,Sum(Nvl(金额,0)-Nvl(冲预交,0)) as 金额
          From 病人预交记录
          Where 记录性质 IN(1,11) And 结帐ID is Not NULL And Nvl(金额,0)<>Nvl(冲预交,0) And 病人ID=v_病人ID
          Having Sum(Nvl(金额,0)-Nvl(冲预交,0))<>0
          Group by 记录状态,NO)
      Order by ID,NO;
 
 --过程变量
    v_结算内容		Varchar2(500);
    v_当前结算		Varchar2(50);
    v_结算方式		病人预交记录.结算方式%Type;
    v_结算金额		病人预交记录.冲预交%Type;
    v_结算号码		病人预交记录.结算号码%Type;
	
    v_费用ID		病人费用记录.ID%Type;
    v_序号			病人费用记录.序号%Type;
	
    v_收费类别		病人费用记录.收费类别%Type;
    v_收费细目ID	病人费用记录.收费细目ID%Type;
    v_计算单位		病人费用记录.计算单位%Type;
    v_收入项目ID	病人费用记录.收入项目ID%Type;
    v_收据费目		病人费用记录.收据费目%Type;
    v_执行部门ID	病人费用记录.执行部门ID%Type;
    v_Temp			Varchar2(500);
	
    v_No			病人预交记录.No%Type;
    v_病人Id		病人预交记录.病人Id%Type;
    v_主页Id		病人预交记录.主页Id%Type;
    v_收款时间		病人预交记录.收款时间%Type;
    v_操作员编号	病人预交记录.操作员编号%Type;
    v_操作员姓名	病人预交记录.操作员姓名%Type;
	
    v_预交金额		病人预交记录.冲预交%Type;	
	 
    v_Error			VARCHAR2(255);
    Err_Custom		EXCEPTION;
 Begin
 
 --1.取预交记录需要的相关信息
    Select No,病人Id,主页Id,登记时间,操作员编号,操作员姓名 
		   Into  v_No,v_病人Id,v_主页Id,v_收款时间,v_操作员编号,v_操作员姓名
    From 病人费用记录 Where 结帐ID=结帐ID_IN And Rownum=1 And 记录性质=1;		
    --误差相关信息
    Begin
        Select A.类别,A.ID,A.计算单位,C.ID,C.收据费目 
        Into v_收费类别,v_收费细目ID,v_计算单位,v_收入项目ID,v_收据费目
        From 收费项目目录 A,收费价目 B,收入项目 C,收费特定项目 D
        Where D.特定项目='误差项' And D.收费细目ID=A.Id  And A.ID=B.收费细目ID And B.收入项目ID=C.ID
            And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-01-01','YYYY-MM-DD')) ;
    Exception
        When Others Then
        Begin
            v_Error:='不能正确读取收费误差项的信息，请先检查该项目是否设置正确。';
            Raise Err_Custom;
        End;
    End;	
	
 --2.删除旧的记录,回退汇总数据
    --回退人员缴款余额,病人余额,
    For C_DEL In (SELECT * FROM 病人预交记录 WHERE 结帐ID=结帐ID_IN And 记录性质=3) Loop
	    Update 人员缴款余额 Set 余额=Nvl(余额,0)-Nvl(C_DEL.冲预交,0) Where 结算方式=C_DEL.结算方式;	   
      	If SQL%RowCount=0 Then
           Insert Into 人员缴款余额(收款员,结算方式,性质,余额) Values(C_DEL.操作员姓名,C_DEL.结算方式,1,-1*C_DEL.冲预交);
		End If;
    End Loop;
	
    If v_病人Id>0 Then
    	Begin
        	Select Sum(冲预交) Into V_预交金额 From 病人预交记录 Where 结帐Id=结帐id_IN And 记录性质 In (1,11);
        Exception
        	When Others Then NULL;
    	End;	
    	If v_预交金额<>0 Then
        	Update 病人余额 Set 预交余额=Nvl(预交余额,0)+V_预交金额 Where 病人ID=v_病人Id And 性质=1;
        	IF SQL%RowCount=0 Then
            	Insert Into 病人余额(病人ID,预交余额,性质) Values(v_病人Id,V_预交金额,1);
            End IF;
    	End If;
    End If;
	
	--回退病人费用汇总
	--只可能产生误差金额的变化,其它的不会变,并且只可能存在一行,仅为了变量处理方便而用游标
    For C_Error In (
        Select TRUNC(登记时间) as 日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,门诊标志,应收金额,实收金额,结帐金额
        From 病人费用记录
        Where 记录性质=1 And 记录状态=1 And 结帐Id=结帐Id_IN And 附加标志=9
    ) Loop
        Update 病人费用汇总
            Set 应收金额=Nvl(应收金额,0)-C_Error.应收金额,实收金额=Nvl(实收金额,0)-C_Error.实收金额,结帐金额=Nvl(结帐金额,0)-C_Error.结帐金额
        Where 日期=C_Error.日期
            And Nvl(病人病区ID,0)=Nvl(C_Error.病人病区ID,0) And Nvl(病人科室ID,0)=Nvl(C_Error.病人科室ID,0)
            And Nvl(开单部门ID,0)=Nvl(C_Error.开单部门ID,0) And Nvl(执行部门ID,0)=Nvl(C_Error.执行部门ID,0)
            And 收入项目ID+0=C_Error.收入项目Id And 来源途径=C_Error.门诊标志 And 记帐费用=0; 
        If SQL%RowCount=0 Then
            Insert Into 病人费用汇总(
                日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
            Values(
                C_Error.日期,C_Error.病人病区ID,C_Error.病人科室ID,C_Error.开单部门ID,C_Error.执行部门ID,
                C_Error.收入项目ID,C_Error.门诊标志,0,-1*C_Error.应收金额,-1*C_Error.实收金额,-1*C_Error.结帐金额);
        End If;
    End Loop; 
 
    --删除收费结算,保险结算记录		     
    Delete 病人预交记录 Where 结帐ID=结帐ID_IN And 记录性质=3; 
    --第一次冲预交的,清空冲减额
    Update 病人预交记录 Set 冲预交=Null,结帐Id=Null	Where 结帐Id=结帐ID_IN And 记录性质=1;
    --删除冲余款
    Delete 病人预交记录 Where 结帐Id=结帐ID_IN And 记录性质=11;
    --删除误差记录
    Delete 病人费用记录 Where 结帐Id=结帐Id_IN And 附加标志=9;	
 
 --3.产生病人费用记录的误差记录
    If 误差_IN <>0 Then 
        Select 病人费用记录_ID.Nextval Into v_费用ID From Dual;
        Select Max(序号)+1 Into v_序号 From 病人费用记录 Where 结帐ID=结帐ID_IN And 记录性质=1;
	 v_Temp:=zl_Identity;
	 v_执行部门ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
			
        Insert Into 病人费用记录(
            ID,记录性质,NO,实际票号,记录状态,序号,从属父号,价格父号,门诊标志,病人ID,标识号,床号,姓名,性别,
            年龄,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,发药窗口,付数,数次,加班标志,附加标志,
            收入项目ID,收据费目,标准单价,应收金额,实收金额,记帐费用,划价人,开单部门ID,开单人,发生时间,登记时间,
            执行部门ID,执行状态,结帐ID,结帐金额,操作员编号,操作员姓名,是否上传)
        Select
            v_费用ID,记录性质,NO,实际票号,记录状态,v_序号,NULL,NULL,门诊标志,病人ID,标识号,床号,姓名,性别,年龄,
            病人病区ID,病人科室ID,费别,v_收费类别,v_收费细目ID,v_计算单位,发药窗口,1,1,加班标志,9,
            v_收入项目ID,v_收据费目,误差_IN,误差_IN,误差_IN,记帐费用,划价人,开单部门ID,开单人,发生时间,登记时间,
            v_执行部门ID,执行状态,结帐ID,误差_IN,操作员编号,操作员姓名,1
        From 病人费用记录
        Where 记录性质=1 And 记录状态=1 And 结帐ID=结帐ID_IN And Rownum=1;
    End If;
  
 --4.重新生成病人预交记录相关数据	
    --4.1.收费结算
    If 收费结算_IN IS Not NULL Then
		v_结算内容:=收费结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
			v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
			v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
			v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
			v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
			v_结算号码:=LTrim(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
			
			If Nvl(v_结算金额,0)<>0 Then
				Insert Into 病人预交记录(
					ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,结算号码,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
				Values(
					病人预交记录_ID.Nextval,3,v_No,1,v_病人Id,v_主页Id,'收费结算',
					v_结算方式,v_结算号码,v_收款时间,v_操作员编号,v_操作员姓名,v_结算金额,结帐ID_IN);
			End IF;
            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
     End IF;
	 
    --4.2.保险结算
    If 保险结算_IN IS Not NULL Then  
		v_结算内容:=保险结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
            v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
            v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
            v_结算金额:=To_Number(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
			
			If Nvl(v_结算金额,0)<>0 Then
				Insert Into 病人预交记录(
					ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
				Values(
					病人预交记录_ID.Nextval,3,v_No,1,v_病人Id,v_主页Id,'保险结算',
					v_结算方式,v_收款时间,v_操作员编号,v_操作员姓名,v_结算金额,结帐ID_IN);
			End IF;
            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
    End IF;

    --4.3.预交结算
    IF Nvl(冲预交额_IN,0)<>0 THEN
        IF Nvl(v_病人Id,0)=0 Then
            v_Error:='不能确定病人的病人ID,收费不能使用预交款结算,结算操作失败！';
            Raise Err_Custom;
        End if;
		
        v_预交金额:=冲预交额_IN;
        For r_Deposit IN c_Deposit(v_病人Id) Loop
            IF r_Deposit.ID<>0 Then
                --第一次冲预交
                Update 病人预交记录 
                    Set 冲预交=Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额),结帐ID=结帐ID_IN
                Where ID=r_Deposit.ID;
            Else
                --冲上次剩余额
                INSERT Into 病人预交记录(
                    ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,
                    结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号,收款时间,
                    操作员姓名,操作员编号,冲预交,结帐ID)
                Select 病人预交记录_ID.Nextval,NO,实际票号,11,记录状态,病人ID,
                     主页ID,科室ID,NULL,结算方式,结算号码,摘要,缴款单位,
                     单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,
                     Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额),结帐ID_IN
                From 病人预交记录
                Where NO=r_Deposit.NO And 记录状态=r_Deposit.记录状态 AND 记录性质 IN(1,11) And RowNum=1;
            End IF;
            --检查是否已经处理完
            IF r_Deposit.金额<v_预交金额 Then
                v_预交金额:=v_预交金额-r_Deposit.金额;
            Else
                v_预交金额:=0;
            End IF;
            IF v_预交金额=0 Then 
                Exit;
            End IF;
        End Loop;
        --检查金额是否足够
        IF v_预交金额>0 Then
            v_Error:='病人的当前预交余额不足金额 '||Ltrim(To_Char(冲预交额_IN,'9999999990.00'))||' ！';
            Raise Err_Custom;
        End IF;

        --更新病人预交余额
        Update 病人余额 Set 预交余额=Nvl(预交余额,0)-冲预交额_IN Where 病人ID=v_病人Id And 性质=1;
        IF SQL%RowCount=0 Then
            Insert Into 病人余额(病人ID,预交余额,性质) Values(v_病人Id,-1*冲预交额_IN,1);
        End IF;
        Delete From 病人余额 Where 病人ID=v_病人Id And 性质=1 And Nvl(费用余额,0)=0 And Nvl(预交余额,0)=0;
    End IF;
	
    --5.相关汇总表的处理	
    --汇总"人员缴款余额"
	--收费结算
    IF 收费结算_IN IS Not NULL Then
        v_结算内容:=收费结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
			v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
			v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
			v_当前结算:=Substr(v_当前结算,Instr(v_当前结算,'|')+1);
			v_结算金额:=To_Number(Substr(v_当前结算,1,Instr(v_当前结算,'|')-1));
			
			If Nvl(v_结算金额,0)<>0 Then
				Update 人员缴款余额
					Set 余额=Nvl(余额,0)+Nvl(v_结算金额,0)
				 Where 收款员=v_操作员姓名
					And 性质=1 And 结算方式=v_结算方式;
				If SQL%RowCount=0 Then
					Insert Into 人员缴款余额(
						收款员,结算方式,性质,余额)
					Values(
						v_操作员姓名,v_结算方式,1,Nvl(v_结算金额,0));
				End If;
			End IF;
            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
    End IF;

    --各个保险结算
    IF 保险结算_IN IS Not NULL Then
        v_结算内容:=保险结算_IN||'||';
        While v_结算内容 IS Not NULL Loop
            v_当前结算:=Substr(v_结算内容,1,Instr(v_结算内容,'||')-1);
            v_结算方式:=Substr(v_当前结算,1,Instr(v_当前结算,'|')-1);
            v_结算金额:=To_Number(Substr(v_当前结算,Instr(v_当前结算,'|')+1));
			
			If Nvl(v_结算金额,0)<>0 Then
				Update 人员缴款余额
					Set 余额=Nvl(余额,0)+Nvl(v_结算金额,0)
				 Where 收款员=v_操作员姓名
					And 性质=1 And 结算方式=v_结算方式;
				If SQL%RowCount=0 Then
					Insert Into 人员缴款余额(
						收款员,结算方式,性质,余额)
					Values(
						v_操作员姓名,v_结算方式,1,Nvl(v_结算金额,0));
				End If;
			End IF;
            v_结算内容:=Substr(v_结算内容,Instr(v_结算内容,'||')+2);
        End Loop;
    End IF;
    Delete From 人员缴款余额 Where 性质=1 And 收款员=v_操作员姓名 And Nvl(余额,0)=0;

    --病人费用汇总,只需重汇误差行,因为其它项不会变
    For r_MoneyRow In (
        Select TRUNC(登记时间) as 日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,门诊标志,应收金额,实收金额,结帐金额
        From 病人费用记录
        Where 记录性质=1 And 记录状态=1 And 结帐Id=结帐Id_IN And 附加标志=9
	) Loop
        Update 病人费用汇总
            Set 应收金额=Nvl(应收金额,0)+r_MoneyRow.应收金额,实收金额=Nvl(实收金额,0)+r_MoneyRow.实收金额,结帐金额=Nvl(结帐金额,0)+r_MoneyRow.结帐金额
        Where 日期=r_MoneyRow.日期
            And Nvl(病人病区ID,0)=Nvl(r_MoneyRow.病人病区ID,0) And Nvl(病人科室ID,0)=Nvl(r_MoneyRow.病人科室ID,0)
            And Nvl(开单部门ID,0)=Nvl(r_MoneyRow.开单部门ID,0) And Nvl(执行部门ID,0)=Nvl(r_MoneyRow.执行部门ID,0)
            And 收入项目ID+0=r_MoneyRow.收入项目Id  And 来源途径=r_MoneyRow.门诊标志 And 记帐费用=0;

        If SQL%RowCount=0 Then
            Insert Into 病人费用汇总(
                日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
            Values(
                r_MoneyRow.日期,r_MoneyRow.病人病区ID,r_MoneyRow.病人科室ID,r_MoneyRow.开单部门ID,r_MoneyRow.执行部门ID,
                r_MoneyRow.收入项目ID,r_MoneyRow.门诊标志,0,r_MoneyRow.应收金额,r_MoneyRow.实收金额,r_MoneyRow.结帐金额);
        End If;
    End Loop; 
 
 	--6.医保相关表的处理
    Delete 医保核对表 Where 结帐Id=结帐Id_IN;
 
EXCEPTION
    WHEN Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS Then zl_ErrOrCenter(SQLCODE,SQLERRM);
End zl_门诊收费结算_Update;
/

CREATE OR REPLACE Procedure zl_门诊收费记录_Delete(
    NO_IN				病人费用记录.NO%Type,
    操作员编号_IN		病人费用记录.操作员编号%Type,
    操作员姓名_IN		病人费用记录.操作员姓名%Type,
    医保结算方式_IN		Varchar2:=Null,
    序号_IN				Varchar2:=NULL,
    结算方式_IN			病人预交记录.结算方式%Type:=NULL,
    票据号_IN			票据使用明细.号码%Type:=NULL,
    领用ID_IN			票据使用明细.领用ID%Type:=NULL,
    误差_IN				病人费用记录.实收金额%Type:=0,
	退费时间_IN			病人费用记录.登记时间%Type:=NULL,
	票据处理_IN			Number:=0
)
AS
--功能：删除一张门诊收费单据
--参数：
--        序号_IN:要退费的项目序号,格式为"1,3,5,6...",缺省NULL表示退"未退的"所有行。
--        结算方式_IN:当为部分退费时,退费金额的结算方式。
--		  误差_IN=部份退费或医保全退但某种结算退现金时才有,用于这两种情况产生的误差。
--		  医保结算方式_IN=个人帐户,退休补助
--          医保退费时,不支持结算作废的结算方式,如果为空表示非医保退费或医保退费全部结算允许作废。
--        票据号_IN,领用ID_IN:当为部分退费时,需要重打收据的(起始)票据号和领用批次ID。
--        票据处理_IN=0-用于单张单据退费,按正常方式处理。
--                    1-用于多张单据退费,且是全部退,只收回票据(注意只能收回一次)。
--                    2-用于多张单据退费,且是部份退,这里不处理票据,全部退费后单独处理。

    --该游标为要退费单据的所有原始记录
      --医保全退但某种结算退现金从而产生了新的误差时,排开此处的误差处理,执行完本过程后,界面程序中单独处理新误差    
    Cursor c_Bill is
        Select * From 病人费用记录
        Where NO=NO_IN And 记录性质=1 And 记录状态 IN(1,3) And 
              NVL(附加标志,0)<>Decode(医保结算方式_IN,Null,999,Decode(sign(误差_IN),0,999,9))
        Order by 收费细目ID,序号;

    --该游标用于处理药品库存可用数量
    --不要管费用的执行状态,因为先于此步处理
    Cursor c_Stock is
        Select * From 药品收发记录
        Where NO=NO_IN And 单据 IN(8,24) --@@@
			And Mod(记录状态,3)=1 And 审核人 IS NULL
            And 费用ID IN(
                Select ID From 病人费用记录
                Where NO=NO_IN And 记录性质=1 And 记录状态 IN(1,3)
                    And 收费类别 IN('4','5','6','7')--@@@
                    And (INSTR(','||序号_IN||',',','||序号||',')>0 Or 序号_IN Is Null)
                )
        Order BY 药品ID;

    --该游标用于处理未发药品记录
    Cursor c_Spare is
        Select * From 未发药品记录 Where NO=NO_IN And 单据 IN(8,24);--@@@

    --该光标用于处理人员缴款余额中退的不同结算方式的金额
    Cursor c_Money(v结帐ID 病人预交记录.结帐ID%Type) is
        Select 结算方式,冲预交
        From 病人预交记录
        Where 记录性质=3 And 记录状态=2 And 结帐ID=v结帐ID
            And 结算方式 is Not Null And Nvl(冲预交,0)<>0;

	v_医嘱ID		病人医嘱记录.ID%Type;

    v_病人ID		病人信息.病人ID%Type;
    v_结帐ID		病人费用记录.结帐ID%Type;
    v_打印ID		票据打印内容.ID%Type;
    v_已退金额		病人预交记录.冲预交%Type;
    v_预交金额		病人预交记录.冲预交%Type;

    --部分退费计算变量
    v_剩余数量		Number;
    v_剩余应收		Number;
    v_剩余实收		Number;
    v_剩余统筹		Number;

    v_准退数量		Number;
    v_退费次数		Number;

    v_应收金额		Number;
    v_实收金额		Number;
    v_统筹金额		Number;
    v_总金额		Number;

	v_首次全退		Number;--是否第一次退费且是全退,用于序号_IN=NULL时判断是否也退误差
    v_正常退费		Number;--是否第一次退费且全部退费,在每行退费过程中判断得到。
    v_全部退完		Number;--本次退费是否将剩余部分全部退完了,退费完成后读SQL得到。

    v_退费结算		结算方式.名称%Type;

    v_结算内容      Varchar2(500);

    v_Dec			Number;

	v_Date			Date;
    v_Count			Number;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
    --是否已经全部完全执行(只是该单据整张单据的检查)
    Select Nvl(Count(*),0) Into v_Count From 病人费用记录
    Where NO=NO_IN And 记录性质=1 And 记录状态 IN(1,3) And Nvl(执行状态,0)<>1;
    IF v_Count = 0 Then
        v_Error := '该单据中的项目已经全部完全执行！';
        Raise Err_Custom;
    End IF;

    --未完全执行的项目是否有剩余数量(只是整张单据的检查)
    --执行状态在原始记录上判断
    Select Nvl(Count(*),0) Into v_Count
    From (
        Select 序号,Sum(数量) as 剩余数量
        From (
            Select 记录状态,Nvl(价格父号,序号) as 序号,
                Avg(Nvl(付数,1)*数次) as 数量
            From 病人费用记录
            Where NO=NO_IN And 记录性质=1 And Nvl(附加标志,0)<>9
                And Nvl(价格父号,序号) IN (
                        Select Nvl(价格父号,序号)
                        From 病人费用记录
                        Where NO=NO_IN And 记录性质=1
                            And 记录状态 IN(1,3) And Nvl(执行状态,0)<>1)
            Group by 记录状态,Nvl(价格父号,序号)
            )
        Group by 序号 Having Sum(数量)<>0);
    IF v_Count = 0 Then
        v_Error := '该单据中未完全执行部分项目剩余数量为零,没有可以退费的费用！';
        Raise Err_Custom;
    End IF;

    ---------------------------------------------------------------------------------
    --公用变量
	If 退费时间_IN IS Not NULL Then
		v_Date:=退费时间_IN;
	Else
		Select Sysdate Into v_Date From Dual;
	End IF;
    Select 病人结帐记录_ID.Nextval Into v_结帐ID From Dual;

    --金额小数位数
    Begin
        Select To_Number(参数值) Into v_Dec From 系统参数表 Where 参数号=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --获取结算方式名称
    v_退费结算:=结算方式_IN;
    If v_退费结算 IS NULL Then
        Begin
            Select 名称 Into v_退费结算 From 结算方式 Where 性质=1;
        Exception
            When Others Then v_退费结算:='现金';
        End;
    End IF;

	--是否首次全退:第一次退且所有剩余数不为零为行都满足准退数=剩余数
	Select Decode(Count(*),0,1,0) Into v_首次全退 From 病人费用记录 Where 记录性质=1 And 记录状态=2 And NO=NO_IN;
	If v_首次全退=1 Then
		Select
			Decode(Count(A.序号),0,1,0) Into v_首次全退
		From (
			Select --每行剩余数量和准退数量
				A.序号,A.剩余数量,Decode(A.执行状态,1,0,
					Decode(Instr(',4,5,6,7,',A.收费类别),0,A.剩余数量,Nvl(B.准退数量,A.剩余数量))) As 准退数量--@@@
			From (
				Select --有剩余数量的每行剩余数量及原始费用ID和执行状态
					Sum(A.ID) As ID,Sum(A.执行状态) As 执行状态,A.序号,A.收费类别,Sum(数量) As 剩余数量
				From (
					Select
						Decode(A.记录状态,2,0,A.ID) As ID,
						Decode(A.记录状态,2,0,Nvl(A.执行状态,0)) As 执行状态,
						A.序号,A.收费类别,Nvl(A.付数,1)*A.数次 As 数量
					From 病人费用记录 a
					Where A.价格父号 Is Null And Nvl(A.附加标志,0)<>9
						And A.记录性质=1 And A.NO=NO_IN
					) A
				Group By A.序号,A.收费类别 Having Nvl(Sum(数量),0)<>0
				) A,(
				Select --药品准退数量
					费用ID,Sum(Nvl(付数,1)*实际数量) As 准退数量
				From 药品收发记录
				Where NO=NO_IN And Mod(记录状态,3)=1
					And 审核人 Is Null And 单据 IN(8,24)--@@@
				Group By 费用ID
				) B
			Where A.ID=B.费用ID(+)) A
		Where Nvl(A.准退数量,0)<>A.剩余数量;
	End IF;

    --循环处理每行费用(收入项目行)
    v_总金额:=0;
    v_正常退费:=1;
    For r_Bill IN c_Bill Loop
        IF INSTR(','||序号_IN||',',','||Nvl(r_Bill.价格父号,r_Bill.序号)||',') >0 Or 序号_IN Is Null Then
			If 序号_IN IS NULL And Nvl(r_Bill.附加标志,0)=9 And v_首次全退=0 Then
				--不是第一次退费的全退时,不管误差项费用
				v_正常退费:=0;--应该是部份退费
            ElsIF Nvl(r_Bill.执行状态,0)<>1 Then
                --求剩余数量,剩余应收,剩余实收
                Select
                    Sum(Nvl(付数,1)*数次),Sum(应收金额),Sum(实收金额),Sum(统筹金额)
                    Into v_剩余数量,v_剩余应收,v_剩余实收,v_剩余统筹
                From 病人费用记录
                Where NO=NO_IN And 记录性质=1 And 序号=r_Bill.序号;

                IF v_剩余数量=0 Then
                    IF 序号_IN IS Not NULL Then
                        v_Error := '单据中第'||Nvl(r_Bill.价格父号,r_Bill.序号)||'行费用已经全部退费！';
                        Raise Err_Custom;
                    End IF;
                    --情况：未限定行号,原始单据中的该笔已经全部退费(执行状态=0的一种可能)
                    v_正常退费:=0;
                Else
                    --准退数量(非药品项目为剩余数量,原始数量)
                    IF Instr(',4,5,6,7,',r_Bill.收费类别)=0 Then--@@@
                        v_准退数量:=v_剩余数量;
                    Else
                        Select Sum(Nvl(付数,1)*实际数量) Into v_准退数量
                        From 药品收发记录
                        Where NO=NO_IN And 单据 IN(8,24) And Mod(记录状态,3)=1 --@@@
                            And 审核人 is NULL And 费用ID=r_Bill.ID;

						--不跟踪在用的卫生材料@@@
						--有剩余数量无准退数量的有两种情况：
							--1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
							--2.收发记录中已全部发放,即已全部执行,已排除这种情况
						If r_Bill.收费类别='4' And Nvl(v_准退数量,0)=0 Then
							v_准退数量:=v_剩余数量;
						End IF;
                    End if;
                    --是否部分退费
                    If r_Bill.执行状态=2 Or v_准退数量<>Nvl(r_Bill.付数,1)*r_Bill.数次 Then
                        v_正常退费:=0;
                    End IF;

                    --处理病人费用记录

                    --该笔项目第几次退费
                    Select Nvl(Max(Abs(执行状态)),0)+1 Into v_退费次数
                    From 病人费用记录
                    Where NO=NO_IN And 记录性质=1 And 记录状态=2 And 序号=r_Bill.序号;

                    --金额=剩余金额*(准退数/剩余数)
                    v_应收金额:=Round(v_剩余应收*(v_准退数量/v_剩余数量),v_Dec);
                    v_实收金额:=Round(v_剩余实收*(v_准退数量/v_剩余数量),v_Dec);
                    v_统筹金额:=Round(v_剩余统筹*(v_准退数量/v_剩余数量),v_Dec);
                    v_总金额:=v_总金额+v_实收金额;

                    --插入退费记录
                    Insert Into 病人费用记录(
                        ID,NO,实际票号,记录性质,记录状态,序号,从属父号,价格父号,病人ID,主页ID,医嘱序号,门诊标志,姓名,
                        性别,年龄,标识号,床号,费别,病人病区ID,病人科室ID,收费类别,收费细目ID,计算单位,付数,发药窗口,
                        数次,加班标志,附加标志,收入项目ID,收据费目,记帐费用,标准单价,应收金额,实收金额,开单部门ID,
                        开单人,执行部门ID,划价人,执行人,执行状态,执行时间,操作员编号,操作员姓名,发生时间,登记时间,
                        结帐ID,结帐金额,保险项目否,保险大类ID,统筹金额,摘要,是否上传)
                    Select 病人费用记录_ID.Nextval,NO,实际票号,记录性质,2,序号,从属父号,价格父号,病人ID,主页ID,
                        医嘱序号,门诊标志,姓名,性别,年龄,标识号,床号,费别,病人病区ID,病人科室ID,收费类别,
                        收费细目ID,计算单位,Decode(Sign(v_准退数量-Nvl(付数,1)*数次),0,付数,1),发药窗口,
                        Decode(Sign(v_准退数量-Nvl(付数,1)*数次),0,-1*数次,-1*v_准退数量),加班标志,附加标志,
                        收入项目ID,收据费目,记帐费用,标准单价,-1*v_应收金额,-1*v_实收金额,开单部门ID,开单人,执行部门ID,
                        划价人,执行人,-1*v_退费次数,执行时间,操作员编号_IN,操作员姓名_IN,发生时间,v_Date,v_结帐ID,
                        -1*v_实收金额,保险项目否,保险大类ID,-1*v_统筹金额,摘要,Decode(Nvl(附加标志,0),9,1,0)
                    From 病人费用记录 Where ID=r_Bill.ID;

                    --处理病人费用汇总
                    Update 病人费用汇总
                        Set 应收金额=Nvl(应收金额,0) - v_应收金额,
                            实收金额=Nvl(实收金额,0) - v_实收金额,
                            结帐金额=Nvl(结帐金额,0) - v_实收金额
                     Where 日期=Trunc(v_Date)
                        And Nvl(病人病区ID,0)=Nvl(r_Bill.病人病区ID,0)
                        And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
                        And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
                        And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
                        And 收入项目ID+0=r_Bill.收入项目ID
                        And 来源途径=r_Bill.门诊标志 And 记帐费用=0;
                    IF SQL%RowCount=0 Then
                        Insert Into 病人费用汇总(
                            日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
                        Values(
                            Trunc(v_Date),r_Bill.病人病区ID,r_Bill.病人科室ID,r_Bill.开单部门ID,r_Bill.执行部门ID,
                            r_Bill.收入项目ID,r_Bill.门诊标志,0,-1 * v_应收金额,-1 * v_实收金额,-1 * v_实收金额);
                    End IF;

                    --标记原费用记录
                    --执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1
                    Update 病人费用记录
                        Set 记录状态=3,
                            执行状态=Decode(Sign(v_准退数量-v_剩余数量),0,0,1)
                    Where ID=r_Bill.ID;
                End IF;
            Else
                IF 序号_IN Is Not Null Then
                    v_Error := '单据中第'||Nvl(r_Bill.价格父号,r_Bill.序号)||'行费用已经完全执行,不能退费！';
                    Raise Err_Custom;
                End IF;
                --情况:没限定行号,原始单据中包括已经完全执行的
                v_正常退费:=0;
            End IF;
        Else
            v_正常退费:=0;--未指定该笔,属于部分退费
        End IF;
    End Loop;

    ---------------------------------------------------------------------------------
    --处理病人预交记录

    --原单据的结帐ID
    Select 结帐ID Into v_Count From 病人费用记录 Where NO=NO_IN And 记录性质=1 And 记录状态 IN(1,3) And Rownum=1;

    IF v_正常退费=1 THEN --单据第一次退费且全部退完

		--冲预交部分记录
        Insert Into 病人预交记录(
            ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,结算方式,结算号码,摘要,
            缴款单位,单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,冲预交,结帐ID)
        Select
            病人预交记录_ID.Nextval,NO,实际票号,11,记录状态,病人ID,主页ID,科室ID,Null,
            结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号,收款时间,操作员姓名,
            操作员编号,-1*冲预交,v_结帐ID
        From 病人预交记录
        Where 记录性质 IN(1,11) And 结帐ID=v_Count;

        --处理病人预交余额
        Begin
            Select 病人ID,Sum(Nvl(冲预交,0)) Into v_病人ID,v_预交金额 From 病人预交记录
            Where 记录性质 IN(1,11) And 结帐ID=v_Count
            Group by 病人ID;
        Exception
            When Others Then NULL;
        End;
        IF Nvl(v_病人ID,0)<>0 And Nvl(v_预交金额,0)<>0 Then
            Update 病人余额 Set 预交余额=Nvl(预交余额,0)+v_预交金额 Where 病人ID=v_病人ID And 性质=1;
            IF SQL%RowCount=0 Then
                Insert Into 病人余额(病人ID,预交余额,性质) Values(v_病人ID,v_预交金额,1);
            End IF;
            Delete From 病人余额 Where 病人ID=v_病人ID And 性质=1 And Nvl(预交余额,0)=0 And Nvl(费用余额,0)=0;
        End IF;

        --非医保全退,和医保所有结算方式都允许回退,原样退回(冲预交在前面已处理)
        IF 医保结算方式_IN Is Null Then
            Insert Into 病人预交记录(
                ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,结算号码,收款时间,缴款单位,单位开户行,单位帐号,操作员编号,操作员姓名,冲预交,结帐ID)
            Select
                病人预交记录_ID.Nextval,记录性质,NO,2,病人ID,主页ID,摘要,结算方式,结算号码,v_Date,缴款单位,单位开户行,单位帐号,
                操作员编号_IN,操作员姓名_IN,-1*冲预交,v_结帐ID
            From 病人预交记录
            Where 记录性质=3 And 记录状态=1 And 结帐ID=v_Count;

        --医保按允许作废的结算方式退,不允许的,退到指定的结算方式上
        Else
            --a.原样退回
            v_结算内容:=','||医保结算方式_IN ||','||v_退费结算||',' ;           
            Insert Into 病人预交记录(
                ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,结算号码,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
            Select
                病人预交记录_ID.Nextval,记录性质,NO,2,病人ID,主页ID,摘要,结算方式,结算号码,v_Date,操作员编号_IN,操作员姓名_IN,-1*冲预交,v_结帐ID
            From 病人预交记录
            Where 记录性质=3 And 记录状态=1 And 结帐ID=v_Count And instr(v_结算内容,','||结算方式||',')=0;
               
            --b.余下的就是医保不允许作废的结算方式,加上到指定的结算方式上,加上误差(因为界面程序会在这之后退误差)          
            Begin
                Select -1*Nvl(Sum(冲预交),0) Into v_已退金额 From 病人预交记录 Where 结帐ID=v_结帐Id;
            Exception
                When Others Then  v_已退金额:=0;
            End;
            IF (v_总金额-v_已退金额)<>0 Then             --此时的总金额还没有包含误差,因为界面程序中在调用本过程后才产生误差费用记录
                Insert Into 病人预交记录(
                    ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
                Select
                    病人预交记录_ID.Nextval,3,NO,2,病人ID,主页ID,'门诊医保结算退费',v_退费结算,v_Date,操作员编号_IN,
                    操作员姓名_IN,-1*(v_总金额-v_已退金额+Nvl(误差_IN,0)),v_结帐ID
                From 病人预交记录
                Where 记录性质=3 And 记录状态=1 And 结帐ID=v_Count And Rownum=1;
            End IF;
        End IF;      
    Else
        -------------------------------------------------
        --部分退费直接退为指定结算方式
        Insert Into 病人预交记录(
            ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,收款时间,缴款单位,单位开户行,单位帐号,操作员编号,操作员姓名,冲预交,结帐ID)
        Select
            病人预交记录_ID.Nextval,3,NO_IN,2,病人ID,主页ID,'部分退费结算',v_退费结算,v_Date,NULL,NULL,NULL,
            操作员编号_IN,操作员姓名_IN,-1*(v_总金额+Nvl(误差_IN,0)),v_结帐ID
        From 病人预交记录
        Where 记录性质=3 And 记录状态 IN(1,3) And 结帐ID=v_Count And Rownum=1;
    End IF;

    --更新原记录
    Update 病人预交记录 Set 记录状态=3 Where 记录性质=3 And 记录状态 IN(1,3) And 结帐ID=v_Count;

    ---------------------------------------------------------------------------------
    --人员缴款余额(注意是预交记录处理后才处理，包括个人帐户等的结算金额,不含退冲预交款)
    For r_MoneyRow in c_Money(v_结帐ID) Loop
        Update 人员缴款余额
            Set 余额=Nvl(余额,0)+r_MoneyRow.冲预交
         Where 收款员=操作员姓名_IN And 性质=1 And 结算方式=r_MoneyRow.结算方式;
        IF SQL%RowCount=0 Then
            Insert Into 人员缴款余额(
                收款员,结算方式,性质,余额)
            Values(
                操作员姓名_IN,r_MoneyRow.结算方式,1,r_MoneyRow.冲预交);
        End IF;
        Delete From 人员缴款余额 Where 收款员=操作员姓名_IN And 性质=1 And 结算方式=r_MoneyRow.结算方式 And Nvl(余额,0)=0;
    End Loop;

    ---------------------------------------------------------------------------------
    --收费票据处理
    --获取单据最后一次的打印ID(可能是多张单据收费打印)
    --可能以前没有打印内容,这时如果是全部退完,则不用收回；但未全部退完时需要重打。
    Begin
        Select Max(ID) Into v_打印ID From 票据打印内容 Where 数据性质=1 And NO=NO_IN;
    Exception
        When Others Then NULL;
    End;

	If Nvl(票据处理_IN,0)=0 Then
		--判断是全部退完了,以决定是否不重打票据。
		--不管执行状态=1的情况(属本次不可退部分),整张单据看
		--多单据收费时,必须所有单据都退完了才不重打
		Select Nvl(Count(*),0) Into v_Count
		From (
			Select NO,序号,Sum(数量) as 剩余数量
			From (
				Select NO,记录状态,Nvl(价格父号,序号) as 序号,
					Avg(Nvl(付数,1)*数次) as 数量
				From 病人费用记录
				Where 记录性质=1 And Nvl(附加标志,0)<>9
					And NO IN(
						Select NO From 票据打印内容 Where ID=v_打印ID And 数据性质=1
						Union ALL
						Select NO_IN From Dual)
				Group by NO,记录状态,Nvl(价格父号,序号)
				)
			Group by NO,序号 Having Sum(数量)<>0);
		If v_Count=0 Then
			v_全部退完:=1;
		Else
			v_全部退完:=0;
		End IF;

		IF v_全部退完=1 Then
			--全部退完收回多张票据(可能以前没有打印,无收回)
			If v_打印ID IS Not NULL Then
				Insert Into 票据使用明细(
					ID,票种,号码,性质,原因,领用ID,打印ID,使用时间,使用人)
				Select
					票据使用明细_ID.Nextval,票种,号码,2,2,领用ID,打印ID,v_Date,操作员姓名_IN
				From 票据使用明细
				Where 打印ID=v_打印ID And 票种=1 And 性质=1;
			End IF;
		Else
			--是部分退费,重打票据
			If 票据号_IN is Not NULL Then
				--现在可以打,但以前如果没打过程内就无收回
				zl_门诊收费记录_RePrint(NO_IN,票据号_IN,领用ID_IN,操作员姓名_IN,v_Date,1);
			ElsIf v_打印ID IS Not NULL Then
				--现在不打印,但要收回以前的
				Insert Into 票据使用明细(
					ID,票种,号码,性质,原因,领用ID,打印ID,使用时间,使用人)
				Select
					票据使用明细_ID.Nextval,票种,号码,2,2,领用ID,打印ID,v_Date,操作员姓名_IN
				From 票据使用明细
				Where 打印ID=v_打印ID And 票种=1 And 性质=1;
			End IF;
		End If;
	ElsIf Nvl(票据处理_IN,0)=1 Then
		--多张单据全退，仅回收票据(多张单据只能收回一次)
		If v_打印ID IS Not NULL Then
			Select Count(*) Into v_Count From 票据使用明细 Where 票种=1 And 性质=2 And 打印ID=v_打印ID;
			If v_Count=0 Then
				Insert Into 票据使用明细(
					ID,票种,号码,性质,原因,领用ID,打印ID,使用时间,使用人)
				Select
					票据使用明细_ID.Nextval,票种,号码,2,2,领用ID,打印ID,v_Date,操作员姓名_IN
				From 票据使用明细
				Where 打印ID=v_打印ID And 票种=1 And 性质=1;
			End IF;
		End IF;
	ElsIf Nvl(票据处理_IN,0)=2 Then
		NULL;--多张单据部份退，这时不处理票据
	End IF;

    ---------------------------------------------------------------------------------
    --药品相关内容
    For r_Stock in c_Stock Loop
        --处理药品库存
        If r_Stock.库房ID IS Not NULL Then
            Update 药品库存
                Set 可用数量=Nvl(可用数量,0)+Nvl(r_Stock.付数,1)*Nvl(r_Stock.实际数量,0)
             Where 库房ID=r_Stock.库房ID And 药品ID=r_Stock.药品ID
                And Nvl(批次,0)=Nvl(r_Stock.批次,0) And 性质=1;
            IF SQL%RowCount=0 Then
                Insert Into 药品库存(
                    库房ID,药品ID,性质,批次,效期,可用数量,上次批号,上次产地,灭菌效期)--@@@
                Values(
                    r_Stock.库房ID,r_Stock.药品ID,1,r_Stock.批次,r_Stock.效期,
                    Nvl(r_Stock.付数,1)*Nvl(r_Stock.实际数量,0),r_Stock.批号,r_Stock.产地,r_Stock.灭菌效期);
            End IF;
        End IF;

        --删除药品收发记录
        Delete From 药品收发记录 Where ID=r_Stock.ID;
    End Loop;

    --未发药品记录
    For r_Spare IN c_Spare Loop
        Select Nvl(Count(*),0) Into v_Count
        From 药品收发记录
        Where NO=NO_IN And 单据=r_Spare.单据 --@@@
			And Mod(记录状态,3)=1 And 审核人 is NULL
			And Nvl(库房ID,0)=Nvl(r_Spare.库房ID,0);

        If v_Count=0 Then
            Delete From 未发药品记录 Where 单据=r_Spare.单据 --@@@
				And NO=NO_IN And Nvl(库房ID,0)=Nvl(r_Spare.库房ID,0);
        End IF;
    End Loop;

	--整张单据全部冲完时，删除病人医嘱附费
	If 序号_IN IS NULL Then
		Begin
			Select 医嘱序号 Into v_医嘱ID From 病人费用记录 Where 记录性质=1 And 记录状态=3 And NO=NO_IN And Rownum=1;
		Exception
			When Others Then NULL;
		End;
		If v_医嘱ID IS Not NULL Then
			Select Nvl(Count(*),0) Into v_Count
			From (
				Select 序号,Sum(数量) as 剩余数量
				From (
					Select 记录状态,Nvl(价格父号,序号) as 序号,
						Avg(Nvl(付数,1)*数次) as 数量
					From 病人费用记录
					Where 记录性质=1 And Nvl(附加标志,0)<>9
						And 医嘱序号+0=v_医嘱ID And NO=NO_IN
					Group by 记录状态,Nvl(价格父号,序号)
					)
				Group by 序号 Having Sum(数量)<>0);
			IF v_Count = 0 Then
				Delete From 病人医嘱附费 Where 医嘱ID=v_医嘱ID And 记录性质=1 And NO=NO_IN;
			End IF;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_门诊收费记录_Delete;
/

-------------------------------------------------------
--模块：住院记帐记录.SQL
Create Or Replace Procedure zl_住院记帐记录_Insert(
    NO_IN			病人费用记录.NO%Type,
    序号_IN         病人费用记录.序号%Type,
    病人ID_IN       病人费用记录.病人ID%Type,
    主页ID_IN       病人费用记录.主页ID%Type,
    标识号_IN       病人费用记录.标识号%Type,
    姓名_IN         病人费用记录.姓名%Type,
    性别_IN         病人费用记录.性别%Type,
    年龄_IN         病人费用记录.年龄%Type,
    床号_IN         病人费用记录.床号%Type,
    费别_IN         病人费用记录.费别%Type,
    病区ID_IN       病人费用记录.病人病区ID%Type,
    科室ID_IN       病人费用记录.病人科室ID%Type,
    加班标志_IN     病人费用记录.加班标志%Type,
    婴儿费_IN       病人费用记录.婴儿费%Type,
    开单部门ID_IN   病人费用记录.开单部门ID%Type,
    开单人_IN       病人费用记录.开单人%Type,
    从属父号_IN     病人费用记录.从属父号%Type,
    收费细目ID_IN   病人费用记录.收费细目ID%Type,
    收费类别_IN     病人费用记录.收费类别%Type,
    计算单位_IN     病人费用记录.计算单位%Type,
    保险项目否_IN   病人费用记录.保险项目否%Type,
    保险大类ID_IN   病人费用记录.保险大类ID%Type,
    保险编码_IN     病人费用记录.保险编码%Type,
    付数_IN         病人费用记录.付数%Type,
    数次_IN         病人费用记录.数次%Type,
    附加标志_IN     病人费用记录.附加标志%Type,
    执行部门ID_IN   病人费用记录.执行部门ID%Type,
    价格父号_IN     病人费用记录.价格父号%Type,
    收入项目ID_IN   病人费用记录.收入项目ID%Type,
    收据费目_IN     病人费用记录.收据费目%Type,
    标准单价_IN     病人费用记录.标准单价%Type,
    应收金额_IN     病人费用记录.应收金额%Type,
    实收金额_IN     病人费用记录.实收金额%Type,
    统筹金额_IN     病人费用记录.统筹金额%Type,
    发生时间_IN     病人费用记录.发生时间%Type,
    登记时间_IN     病人费用记录.登记时间%Type,
    药品摘要_IN     药品收发记录.摘要%Type,
    划价_IN         Number,
    操作员编号_IN   病人费用记录.操作员编号%Type,
    操作员姓名_IN   病人费用记录.操作员姓名%Type,
    多病人单_IN     Number := 0,
    类别ID_IN       药品单据性质.类别ID%Type:=Null,
    记帐单ID_IN     病人费用记录.记帐单ID%Type:=Null,
    费用摘要_IN     病人费用记录.摘要%Type:=Null,
    是否急诊_IN     病人费用记录.是否急诊%Type:=0,
    医嘱序号_IN     病人费用记录.医嘱序号%TYPE:=NULL,
    频次_IN         药品收发记录.频次%Type:=NULL,
    单量_IN         药品收发记录.单量%Type:=NULL,
    用法_IN         药品收发记录.用法%Type:=NULL,--用法[|煎法]
    期效_IN         药品收发记录.扣率%Type:=NULL,
    计价特性_IN     药品收发记录.扣率%Type:=NULL,
    简单记帐_IN     Number:=0
)
AS
    --功能：新收一张住院记帐单据
    --参数：
    --   药品摘要_IN:存放医嘱中的附加说明或修改保存新单据时用。目前仅用于存放于药品收发记录的摘要中。
    --         原单据(记录状态=2)记录修改产生的新单据号。
    --         新单据(记录状态=1)记录所修改的原单据号。
    --   划价-是否属于住院划价。
    v_费用ID 病人费用记录.ID%Type;
    v_优先级 未发药品记录.优先级%Type;

    --药房分批、时价药品--
    ------------------------------------------------------------
    --该游标用于分批药品数量分解
    Cursor c_Stock is
        Select * From 药品库存 
        Where 药品ID=收费细目ID_IN And 库房ID=执行部门ID_IN
            And 性质=1 And(Nvl(批次,0)=0 Or 效期 is Null Or 效期>Trunc(Sysdate))
            And Nvl(可用数量,0)<>0
        Order By Nvl(批次,0);
    r_Stock c_Stock%RowType;
    
    --属性
    v_分批			药品规格.药房分批%Type;
    v_时价			收费项目目录.是否变价%Type;
    v_名称			收费项目目录.名称%Type;
    --临时变量
    v_总数量		Number;
    v_当前数量		Number;
    v_总金额		Number;
    v_当前单价		Number;
    --药品收发记录
    v_批次			药品收发记录.批次%Type;
    v_产地			药品收发记录.产地%Type;
    v_批号			药品收发记录.批号%Type;
    v_效期			药品收发记录.效期%Type;
    v_序号			药品收发记录.序号%Type;
    v_扣率			药品收发记录.扣率%Type;
	v_灭菌效期		药品收发记录.灭菌效期%Type;
	v_灭菌日期		药品收发记录.灭菌日期%Type;
    ------------------------------------------------------------
	v_用法			药品收发记录.用法%Type;
	v_煎法			药品收发记录.外观%Type;

	v_Dec			Number;
	v_Count			Number;
	v_Error			Varchar2(255);
    Err_custom		Exception;
Begin
	--药品用法煎法分解
	IF 用法_IN IS Not NULL Then
		IF Instr(用法_IN,'|')>0 Then
			v_用法:=Substr(用法_IN,1,Instr(用法_IN,'|')-1);
			v_煎法:=Substr(用法_IN,Instr(用法_IN,'|')+1);
		Else
			v_用法:=用法_IN;
		End IF;
	End IF;

    --金额小数位数
    Begin
        Select To_Number(参数值) Into v_Dec From 系统参数表 Where 参数号=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --病人费用记录
    Select 病人费用记录_ID.Nextval Into v_费用ID From Dual;

    Insert Into 病人费用记录(
        ID,记录性质,NO,记录状态,序号,从属父号,价格父号,多病人单,门诊标志,病人ID,主页ID,
        标识号,姓名,性别,年龄,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,
        保险项目否,保险大类ID,保险编码,发药窗口,付数,数次,加班标志,附加标志,婴儿费,收入项目ID,收据费目,
        标准单价,应收金额,实收金额,统筹金额,记帐费用,开单部门ID,开单人,发生时间,登记时间,
        执行部门ID,执行状态,划价人,操作员编号,操作员姓名,记帐单ID,摘要,是否急诊,医嘱序号)
    Values(
        v_费用ID,2,NO_IN,Decode(划价_IN,1,0,1),序号_IN,Decode(从属父号_IN,0,Null,从属父号_IN),
        Decode(价格父号_IN,0,Null,价格父号_IN),多病人单_IN,2,病人ID_IN,主页ID_IN,
        Decode(标识号_IN,0,Null,标识号_IN),姓名_IN,性别_IN,年龄_IN,
        Decode(床号_IN,0,Null,床号_IN),Decode(病区ID_IN,0,Null,病区ID_IN),
        Decode(科室ID_IN,0,Null,科室ID_IN),费别_IN,收费类别_IN,收费细目ID_IN,
        计算单位_IN,保险项目否_IN,保险大类ID_IN,保险编码_IN,Decode(Nvl(简单记帐_IN,0),0,NULL,收费类别_IN),
        付数_IN,数次_IN,加班标志_IN,附加标志_IN,婴儿费_IN,收入项目ID_IN,收据费目_IN,标准单价_IN,应收金额_IN,
        实收金额_IN,统筹金额_IN,1,开单部门ID_IN,开单人_IN,发生时间_IN,登记时间_IN,
        执行部门ID_IN,0,操作员姓名_IN,Decode(划价_IN,1,Null,操作员编号_IN),
        Decode(划价_IN,1,Null,操作员姓名_IN),记帐单ID_IN,费用摘要_IN,是否急诊_IN,医嘱序号_IN);

    --相关汇总表的处理
	If Nvl(划价_IN,0)=0 Then
		--病人余额
		Update 病人余额
			Set 费用余额=Nvl(费用余额,0)+实收金额_IN
		 Where 病人ID=病人ID_IN And 性质=1;
		IF SQL%RowCount=0 Then
			Insert Into 病人余额(
				病人ID,性质,费用余额,预交余额)
			Values(
				病人ID_IN,1,实收金额_IN,0);
		End IF;

		--病人未结费用
		Update 病人未结费用
			Set 金额=Nvl(金额,0)+实收金额_IN
		 Where 病人ID=病人ID_IN
			And Nvl(主页ID,0)=Nvl(主页ID_IN,0)
			And Nvl(病人病区ID,0)=Nvl(病区ID_IN,0)
			And Nvl(病人科室ID,0)=Nvl(科室ID_IN,0)
			And Nvl(开单部门ID,0)=Nvl(开单部门ID_IN,0)
			And Nvl(执行部门ID,0)=Nvl(执行部门ID_IN,0)
			And 收入项目ID+0=收入项目ID_IN
			And 来源途径+0=2;
		IF SQL%RowCount=0 Then
			Insert Into 病人未结费用(
				病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
			Values(
				病人ID_IN,主页ID_IN,病区ID_IN,科室ID_IN,开单部门ID_IN,执行部门ID_IN,收入项目ID_IN,2,实收金额_IN);
		End IF;

		--病人费用汇总
		Update 病人费用汇总
			Set 应收金额=Nvl(应收金额,0)+应收金额_IN,
				 实收金额=Nvl(实收金额,0)+实收金额_IN
		 Where 日期=Trunc(登记时间_IN)
			And Nvl(病人病区ID,0)=Nvl(病区ID_IN,0)
			And Nvl(病人科室ID,0)=Nvl(科室ID_IN,0)
			And Nvl(开单部门ID,0)=Nvl(开单部门ID_IN,0)
			And Nvl(执行部门ID,0)=Nvl(执行部门ID_IN,0)
			And 收入项目ID+0=收入项目ID_IN
			And 来源途径=2 And 记帐费用=1;
		IF SQL%RowCount=0 Then
			Insert Into 病人费用汇总(
				日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
			Values(
				Trunc(登记时间_IN),病区ID_IN,科室ID_IN,开单部门ID_IN,执行部门ID_IN,收入项目ID_IN,2,1,应收金额_IN,实收金额_IN,0);
		End IF;
	End IF;

    --药品和卫生材料部分
	v_Count:=0;--@@@
	If 收费类别_IN='4' Then--跟踪在用的卫材才处理
		Select 跟踪在用 Into v_Count From 材料特性 Where 材料ID=收费细目ID_IN;
	End IF;
    IF 收费类别_IN in('5','6','7') Or (收费类别_IN='4' And Nvl(v_Count,0)=1) Then
		If 收费类别_IN='4' Then
			Select Nvl(A.在用分批,0),Nvl(B.是否变价,0),B.名称 
				Into v_分批,v_时价,v_名称
			From 材料特性 A,收费项目目录 B
			Where A.材料ID=B.ID And B.ID=收费细目ID_IN;
		Else
			Select Nvl(A.药房分批,0),Nvl(B.是否变价,0),B.名称 
				Into v_分批,v_时价,v_名称
			From 药品规格 A,收费项目目录 B
			Where A.药品ID=B.ID And B.ID=收费细目ID_IN;
		End IF;

        v_总数量:=付数_IN*数次_IN;
        v_总金额:=0;
        Open c_Stock;

        While v_总数量<>0 Loop
            Fetch c_Stock Into r_Stock;
            IF c_Stock%NotFound Then
                --第一次就没有库存,分批或时价都不允许。
                --分批药品数量分解不完,也就是库存不足。
                IF v_分批=1 Or v_时价=1 Then
                    Close c_Stock;
                    If 医嘱序号_IN IS NULL Then
						If 收费类别_IN='4' Then
							v_Error:='第 '||序号_IN||' 行的分批或时价卫生材料"'||v_名称||'"没有足够的材料库存！';
						Else
	                        v_Error:='第 '||序号_IN||' 行的分批或时价药品"'||v_名称||'"没有足够的药品库存！';
						End IF;
                    Else
						If 收费类别_IN='4' Then
							v_Error:='在处理病人"'||姓名_IN||'"时发现分批或时价卫生材料"'||v_名称||'"没有足够的材料库存！';
						Else
	                        v_Error:='在处理病人"'||姓名_IN||'"时发现分批或时价药品"'||v_名称||'"没有足够的药品库存！';
						End IF;
                    End IF;
                    Raise Err_Custom;
                End IF;
            ElsIF(v_分批=1 And Nvl(r_Stock.批次,0)=0) Or(v_分批=0 And Nvl(r_Stock.批次,0)<>0) Then 
                Close c_Stock;
                If 医嘱序号_IN IS NULL Then
					If 收费类别_IN='4' Then
						v_Error:='第 '||序号_IN||' 行卫生材料"'||v_名称||'"的分批属性与库存记录不相符,请检查材料数据的正确性！';
					Else
	                    v_Error:='第 '||序号_IN||' 行药品"'||v_名称||'"的分批属性与库存记录不相符,请检查药品数据的正确性！';
					End IF;
                Else
					If 收费类别_IN='4' Then
						v_Error:='在处理病人"'||姓名_IN||'"时发现卫生材料"'||v_名称||'"的分批属性与库存记录不相符,请检查材料数据的正确性！';
					Else
	                    v_Error:='在处理病人"'||姓名_IN||'"时发现药品"'||v_名称||'"的分批属性与库存记录不相符,请检查药品数据的正确性！';
					End IF;
                End IF;
                Raise Err_Custom;
            End IF;

            --确定本次分解数量
            IF v_分批=1 Or v_时价=1 Then
                --对于不分批的时价只可能分解一次,分解不完上面判断了.它分解是为了计算单价.
                --每次分解取小者,库存不够分解不完在上面判断.
                IF v_总数量<=Nvl(r_Stock.可用数量,0) Then
                    v_当前数量:=v_总数量;
                Else
                    v_当前数量:=Nvl(r_Stock.可用数量,0);
                End if;
                IF v_时价=1 Then 
                    If r_Stock.实际数量=0 Then
                        v_当前单价:=0;
                    Else
                        v_当前单价:=Round(Nvl(r_Stock.实际金额/r_Stock.实际数量,0),5);
                    End IF;
                ElsIf v_分批=1 Then
                    v_当前单价:=标准单价_IN;
                End IF;
            Else
                --普通药品
                --不管够不够,程序中已根据参数判断
                v_当前数量:=v_总数量;
                v_当前单价:=标准单价_IN;
            End IF;

            --药品库存(普通情况可能没有记录)
            IF c_Stock%Found Then
                Update 药品库存
                    Set 可用数量=Nvl(可用数量,0)-v_当前数量
                Where 库房ID=执行部门ID_IN And 药品ID=收费细目ID_IN
                    And Nvl(批次,0)=Nvl(r_Stock.批次,0) And 性质=1;
            ElsIf 执行部门ID_IN IS Not NULL Then
                --只有不分批非时价药品可能库存不足出库
                Update 药品库存
                    Set 可用数量=Nvl(可用数量,0)-v_当前数量
                Where 库房ID=执行部门ID_IN And 药品ID=收费细目ID_IN
                    And Nvl(批次,0)=0 And 性质=1;
                IF SQL%RowCount=0 Then
                    Insert Into 药品库存(
                        库房ID,药品ID,性质,可用数量)
                    Values(
                        执行部门ID_IN,收费细目ID_IN,1,-1*v_当前数量);
                End IF;
            End IF;

            --药品收发记录
			v_批次:=Null;v_批号:=Null;
			v_效期:=Null;v_产地:=Null;
			v_灭菌效期:=Null;v_灭菌日期:=Null;
            IF c_Stock%Found Then
                v_批次:=r_Stock.批次;
                v_批号:=r_Stock.上次批号;
                v_效期:=r_Stock.效期;
                v_产地:=r_Stock.上次产地;

				--卫材灭菌效期:一次性材料且有效期
				IF 收费类别_IN='4' Then
					v_Count:=0;
					Begin
						Select 灭菌效期 Into v_Count From 材料特性 Where Nvl(一次性材料,0)=1 And 材料ID=收费细目ID_IN;
					Exception
						When Others Then Null;
					End;
					IF Nvl(v_Count,0)>0 Then
						v_灭菌效期:=r_Stock.灭菌效期;	
						v_灭菌日期:=v_灭菌效期-v_Count*30;
					End IF;
				End IF;
            End IF;

            Select Nvl(Max(序号),0)+1 Into v_序号 From 药品收发记录 
				Where NO=NO_IN And 记录状态=1 And 单据=Decode(多病人单_IN,1,10,9)+Decode(收费类别_IN,'4',16,0);
					

            --修改的原单据号存放在摘要中
			v_扣率:=Null;
            If 期效_IN IS Not NULL Or 计价特性_IN IS Not NULL THEN 
                v_扣率:=Nvl(期效_IN,0)||Nvl(计价特性_IN,0);
            End IF;
            Insert Into 药品收发记录(
                ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
                药品ID,批次,产地,批号,效期,付数,填写数量,实际数量,零售价,零售金额,
                摘要,填制人,填制日期,费用ID,频次,单量,用法,外观,扣率,灭菌效期,灭菌日期)
            Values(
                药品收发记录_ID.Nextval,1,Decode(多病人单_IN,1,10,9)+Decode(收费类别_IN,'4',16,0),
				NO_IN,v_序号,执行部门ID_IN,开单部门ID_IN,类别ID_IN,-1,收费细目ID_IN,v_批次,v_产地,v_批号,v_效期,
				Decode(v_分批,1,1,付数_IN),Decode(v_分批,1,v_当前数量,v_当前数量/付数_IN),Decode(v_分批,1,v_当前数量,v_当前数量/付数_IN),
                v_当前单价,Round(v_当前单价*v_当前数量,v_Dec),药品摘要_IN,操作员姓名_IN,登记时间_IN,v_费用ID,
                频次_IN,单量_IN,v_用法,v_煎法,v_扣率,v_灭菌效期,v_灭菌日期);

            --未发药品记录
            Update 未发药品记录
                Set 病人ID=病人ID_IN,
                    主页ID=主页ID_IN,
                    姓名=姓名_IN
            Where 单据=Decode(多病人单_IN,1,10,9)+Decode(收费类别_IN,'4',16,0)
				And NO=NO_IN And Nvl(库房ID,0)=Nvl(执行部门ID_IN,0);

            IF SQL%RowCount=0 Then
                --取身份优先级
                Begin
                    Select B.优先级 Into v_优先级 From 病人信息 A,身份 B
                     Where A.身份=B.名称(+) And A.病人ID=病人ID_IN;
                Exception
                    When Others Then Null;
                End;

                Insert Into 未发药品记录(
                    单据,NO,病人ID,主页ID,姓名,优先级,对方部门ID,库房ID,填制日期,已收费,打印状态)
                Values(
                    Decode(多病人单_IN,1,10,9)+Decode(收费类别_IN,'4',16,0),NO_IN,病人ID_IN,
					主页ID_IN,姓名_IN,v_优先级,开单部门ID_IN,执行部门ID_IN,登记时间_IN,Decode(划价_IN,1,0,1),0);
            End IF;

            v_总数量:=v_总数量-v_当前数量;
            v_总金额:=v_总金额+Round(v_当前数量*v_当前单价,v_Dec);
        End Loop;
        
        --可能时价药品的库存金额和数量变化了
        IF v_时价=1 Then
            IF Round(v_总金额/(付数_IN*数次_IN),5)<>标准单价_IN Then 
                Close c_Stock;    
                If 医嘱序号_IN IS NULL Then
					If 收费类别_IN='4' Then
						v_Error:='第 '||序号_IN||' 行的时价卫生材料"'||v_名称||'"当前计算单价不一致,请重新输入数量计算！';
					Else
						v_Error:='第 '||序号_IN||' 行的时价药品"'||v_名称||'"当前计算单价不一致,请重新输入数量计算！';
					End IF;
                Else
                    --医嘱摆药时是按病人分次计算并提交数据库,因此不同病人使用相同实价药品没有问题。
                    --但同一病人同时使用两笔以上相同实价药品则会有问题。
					IF 收费类别_IN='4' Then
						v_Error:='在处理病人"'||姓名_IN||'"时发现时价卫生材料"'||v_名称||'"当前计算的单价发生变化。'||CHR(13)||CHR(10)||'请检查该病人是否同时使用了两笔相同的"'||v_名称||'"！';
					Else
	                    v_Error:='在处理病人"'||姓名_IN||'"时发现时价药品"'||v_名称||'"当前计算的单价发生变化。'||CHR(13)||CHR(10)||'请检查该病人是否同时使用了两笔相同的"'||v_名称||'"！';
					End IF;
                End IF;
                Raise Err_Custom;
            End IF;
        End IF;

        Close c_Stock;
    End IF;
Exception
    When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_住院记帐记录_Insert;
/

CREATE OR REPLACE PROCEDURE zl_住院记帐记录_Verify (
    NO_IN           病人费用记录.NO%TYPE,
    操作员编号_IN   病人费用记录.操作员编号%TYPE,
    操作员姓名_IN	病人费用记录.操作员姓名%TYPE,
	序号_IN			Varchar2:=NULL,
	病人ID_IN		病人费用记录.病人ID%Type:=NULL,
	审核时间_IN		病人费用记录.登记时间%Type:=NULL
) AS
--功能：审核一张住院记帐划价单
--参数：
--		序号_IN：格式如"1,3,5,7,8",为空表示审核所有未审核的行
--		病人ID_IN：只审核指定病人,用于按病人审核记帐表。
--		审核时间_IN：用于部份需要统一控制或返回时间的地方
	--只读取指定序号的,未审核的部份进行处理
	Cursor c_Bill is
		Select * From 病人费用记录 
		Where 记录性质=2 And 记录状态=0 And NO=NO_IN
			And (Instr(','||序号_IN||',',','||Nvl(价格父号,序号)||',')>0 Or 序号_IN Is Null)
			And (病人ID+0=病人ID_IN Or 病人ID_IN IS NULL)
		Order BY 序号;
	
	--审核中包含跟踪在用的未发卫料时，根据参数设置是否自动发料
	Cursor c_Stuff is
		Select NO,单据,库房ID From 未发药品记录
		Where NO=NO_IN And 单据 IN(25,26) And 库房ID IS Not Null
			And Exists(Select 参数值 From 系统参数表 Where 参数号=63 And 参数值='1')
			And Exists(
				Select A.序号 From 病人费用记录 A,材料特性 B
				Where A.记录性质=2 And A.记录状态=1 And A.NO=NO_IN
					And (Instr(','||序号_IN||',',','||Nvl(A.价格父号,A.序号)||',')>0 Or 序号_IN Is Null)
					And (A.病人ID+0=病人ID_IN Or 病人ID_IN IS NULL)
					And A.收费细目ID=B.材料ID And B.跟踪在用=1
				)
		Order BY 库房ID;
	
	v_Date	Date;
BEGIN
	If 审核时间_IN IS Null Then
		Select Sysdate Into v_Date From Dual;
	Else
		v_Date:=审核时间_IN;
	End IF;

	For r_Bill IN c_Bill Loop
		Update 病人费用记录
			Set 记录状态=1,
				操作员编号=操作员编号_IN,
				操作员姓名=操作员姓名_IN,
				登记时间=v_Date --已产生的药品记录的时间不变
		Where ID=r_Bill.ID;

		--药品收发记录.填制日期
		Update 药品收发记录
			Set 填制日期=Decode(Sign(Nvl(审核日期,v_Date)-v_Date),-1,填制日期,v_Date)  
		Where NO=NO_IN AND 单据 IN(9,10,25,26)  AND 费用ID=r_Bill.ID;

		--病人余额
		Update 病人余额
			Set 费用余额=Nvl(费用余额,0)+Nvl(r_Bill.实收金额,0)
		Where 病人ID=r_Bill.病人ID And 性质=1;

		IF SQL%RowCount=0 Then
			Insert Into 病人余额(
				病人ID,性质,费用余额,预交余额)
			Values(
				r_Bill.病人ID,1,r_Bill.实收金额,0);
		End IF;

		--病人未结费用
		Update 病人未结费用
			Set 金额=Nvl(金额,0)+Nvl(r_Bill.实收金额,0)
		 Where 病人ID=r_Bill.病人ID
			And Nvl(主页ID,0)=Nvl(r_Bill.主页ID,0)
			And Nvl(病人病区ID,0)=Nvl(r_Bill.病人病区ID,0)
			And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
			And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
			And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
			And 收入项目ID+0=r_Bill.收入项目ID
			And 来源途径+0=r_Bill.门诊标志;

		IF SQL%RowCount=0 Then
			Insert Into 病人未结费用(
				病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
			Values(
				r_Bill.病人ID,r_Bill.主页ID,r_Bill.病人病区ID,r_Bill.病人科室ID,r_Bill.开单部门ID,r_Bill.执行部门ID,r_Bill.收入项目ID,r_Bill.门诊标志,Nvl(r_Bill.实收金额,0));
		End IF;

		--病人费用汇总
		Update 病人费用汇总
			Set 应收金额=Nvl(应收金额,0)+Nvl(r_Bill.应收金额,0),
				实收金额=Nvl(实收金额,0)+Nvl(r_Bill.实收金额,0)
		 Where 日期=Trunc(v_Date)
			And Nvl(病人病区ID,0)=Nvl(r_Bill.病人病区ID,0)
			And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
			And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
			And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
			And 收入项目ID+0=r_Bill.收入项目ID
			And 来源途径=r_Bill.门诊标志 And 记帐费用=1;

		IF SQL%RowCount=0 Then
			Insert Into 病人费用汇总(
				日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,
				来源途径,记帐费用,应收金额,实收金额,结帐金额)
			Values(
				Trunc(v_Date),r_Bill.病人病区ID,r_Bill.病人科室ID,r_Bill.开单部门ID,
				r_Bill.执行部门ID,r_Bill.收入项目ID,r_Bill.门诊标志,1,r_Bill.应收金额,r_Bill.实收金额,0);
		End IF;
	End Loop;

	--库房中的药品已全部审核则标为已收费
	Update 未发药品记录 Set 已收费=1, 填制日期=v_Date
	Where NO=NO_IN And 单据 IN(9,10) And Nvl(已收费,0)=0 
		And Nvl(库房ID,0) Not IN(
			Select Distinct Nvl(执行部门ID,0) From 病人费用记录 
				Where 记录性质=2 And NO=NO_IN And 收费类别 IN('5','6','7') And 记录状态=0);

	Update 未发药品记录 Set 已收费=1, 填制日期=v_Date
	Where NO=NO_IN And 单据 IN(25,26) And Nvl(已收费,0)=0 
		And Nvl(库房ID,0) Not IN(
			Select Distinct Nvl(执行部门ID,0) From 病人费用记录 
				Where 记录性质=2 And NO=NO_IN And 收费类别='4' And 记录状态=0);

	--处理卫料自动发料
	For r_Stuff In c_Stuff Loop
		zl_材料收发记录_处方发料(r_Stuff.库房ID,r_Stuff.单据,r_Stuff.NO,操作员姓名_IN,操作员姓名_IN,操作员姓名_IN,1,Sysdate);
	End Loop;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_住院记帐记录_Verify;
/



Create OR Replace Procedure zl_病人信息_Merge(
	A病人ID_IN		病人信息.病人ID%Type,--要合并的病人信息
	B病人ID_IN		病人信息.病人ID%Type --要保留的病人信息
--涉及表：
--病人信息,病案主页,病案主页从表,病人变动记录,特殊病人
--门诊病案记录,住院病案记录,检查病案记录,床位状况记录
--保险帐户,保险模拟结算,保险结算记录,帐户年度信息
--病人余额,病人未结费用,病人费用记录,病人预交记录,病人结帐记录,未发药品记录
--病人挂号记录,病人过敏药物,病人过敏记录,病人诊断记录,病人诊断记录
--病人医嘱记录,病人病历记录,病人手术记录,病人手麻记录
--后备表：
--H病人结帐记录,H病人预交记录,H病人费用记录
--H病人医嘱记录,H病人诊断记录,H病人过敏记录
--H病人病历记录,H病人手麻记录,H病人手术记录
) AS
	--被合并的病人
	Cursor c_InfoA IS
		Select A.*,B.主页ID,B.入院日期,B.出院日期
		From 病人信息 A,病案主页 B
		Where A.病人ID=B.病人ID(+) And A.病人ID=A病人ID_IN
		Order by 主页ID;
	r_InfoA c_InfoA%RowType;

	--要保留的病人 
	Cursor c_InfoB IS
		Select A.*,B.主页ID,B.入院日期,B.出院日期
		From 病人信息 A,病案主页 B
		Where A.病人ID=B.病人ID(+) And A.病人ID=B病人ID_IN
		Order by 主页ID;
	r_InfoB c_InfoB%RowType;

	--合并后的信息
	Cursor c_Info(v_病人ID 病人信息.病人ID%Type) is
		Select * From 病案主页
		Where 主页ID=(Select Max(主页ID) From 病案主页 Where 病人ID=v_病人ID)
			And 病人ID=v_病人ID;
	r_Info c_Info%RowType;

	--合并两个住院病人
	Cursor c_MergePati IS
		Select A.姓名,A.门诊号,A.住院号,B.* 
		From 病人信息 A,病案主页 B
		Where A.病人ID=B.病人ID And A.病人ID IN(A病人ID_IN,B病人ID_IN)
		Order by B.入院日期 Desc,NVL(B.出院日期,SYSDATE) Desc;
	
	v_保留ID	病人信息.病人ID%Type;
	v_合并ID	病人信息.病人ID%Type;
	v_门诊号	病人信息.门诊号%Type;
	v_住院号	病人信息.住院号%Type;

	--病人未结费用(门诊部份)
	Cursor c_Owe(v_病人ID 病人信息.病人ID%Type) IS
		Select 
			病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,Sum(金额) AS 金额
		From 病人未结费用 
		Where 主页ID IS NULL And 病人ID=v_病人ID
		Group BY 病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径;

	--病人余额
	Cursor c_Spare(v_病人ID 病人信息.病人ID%Type) IS
		Select 性质,预交余额,费用余额
		From 病人余额 Where 病人ID=v_病人ID;
	
	--检查病案记录(ZLHIS+)
	Cursor c_Check(v_病人ID 病人信息.病人ID%Type) is
		Select * From 检查病案记录 Where 病人ID=v_病人ID Order BY 检查类别;

	--保险帐户
	Cursor c_Insure(v_病人ID 病人信息.病人ID%Type) IS
		Select * From 保险帐户 Where 病人ID=v_病人ID Order BY 险类;
	
	--要保留的保险帐户
	Cursor c_KeepInsure(v_病人ID 病人信息.病人ID%Type,v_险类 保险帐户.险类%Type) IS
		Select * From 保险帐户 Where 病人ID=v_病人ID And 险类=v_险类;
	r_KeepInsure c_KeepInsure%RowType;
	
	Cursor c_Year(v_病人ID 病人信息.病人ID%Type,v_险类 保险帐户.险类%Type) is
		Select * From 帐户年度信息 Where 病人ID=v_病人ID And 险类=v_险类;

	v_卡号		保险帐户.卡号%Type;
	v_医保号	保险帐户.医保号%Type;
	
	v_Count		NUMBER;
	v_Error		VARCHAR2(255);
	Err_Custom	EXCEPTION;

	--字符串变换函数
	Function strSwitch(v_InStr In Varchar2,v_Mask In Number)
		Return Varchar2 Is
		v_OutStr	Varchar2(1000);
	Begin
		For v_Bit In 1 .. Length(v_InStr) Loop
			v_OutStr:=v_OutStr||Chr(Ascii(Substr(v_InStr,v_Bit,1))+v_Mask);
		End Loop;
		Return(v_OutStr);
	End strSwitch;
BEGIN
	--程序中已检查：
	--1.选择了同一个病人
	--2.两个住院病人先入院的却在院(包括两个都在院)。
	--3.两个住院病人的住院期间存在交叉的情况

	OPEN c_InfoA;
	FETCH c_InfoA Into r_InfoA;
	IF c_InfoA%Rowcount=0 THEN
		Close c_InfoA;
		v_Error:='没有发现被合并的病人信息！';
		RAISE Err_Custom;
	END IF;

	OPEN c_InfoB;
	FETCH c_InfoB Into r_InfoB;
	IF c_InfoB%Rowcount=0 THEN
		Close c_InfoB;
		v_Error:='没有发现要保留的病人信息！';
		RAISE Err_Custom;
	END IF;

	--以先住院或先登记的病人ID作为实际上要保留的病人ID
	Select 病人ID Into v_保留ID
	From (
		Select A.病人ID
		From 病人信息 A,病案主页 B
		Where A.病人ID=B.病人ID(+) 
			And A.病人ID IN(A病人ID_IN,B病人ID_IN)
		Order by Nvl(B.入院日期,To_Date('3000-01-01','YYYY-MM-DD')),NVL(B.出院日期,To_Date('3000-01-01','YYYY-MM-DD')),A.登记时间,A.病人ID --住院病人优先
		)
	Where Rownum=1;

	--另外一个就是实际最后要删除的病人ID
	If v_保留ID=A病人ID_IN Then
		v_合并ID:=B病人ID_IN;
	Else
		v_合并ID:=A病人ID_IN;
	End IF;
	v_门诊号:=Nvl(r_InfoB.门诊号,r_InfoA.门诊号);
	v_住院号:=Nvl(r_InfoB.住院号,r_InfoA.住院号);
	
	IF r_InfoA.主页ID IS Not NULL And r_InfoB.主页ID IS Not NULL THEN
		--求两个病人总共的住院次数
		Select Count(*) Into v_Count From 病案主页 Where 病人ID IN(A病人ID_IN,B病人ID_IN);
		
		FOR r_Merge IN c_MergePati LOOP
			--处理病案主页部份(涉及病人ID,主页ID字段的表)
			If Not (r_Merge.病人ID=v_保留ID And r_Merge.主页ID=v_Count) Then
				--该病案主页要删除时,不能是已编目了的。
				If r_Merge.编目日期 IS Not NULL Then
					Close c_InfoA; Close c_InfoB;
					If r_Merge.住院号 IS NULL Then
						v_Error:='病人'||r_Merge.姓名||'(病人ID='||r_Merge.病人ID||')存在已编目的病案,不允许合并该病人。';
					Else
						v_Error:='病人'||r_Merge.姓名||'(病人ID='||r_Merge.病人ID||',住院号='||r_Merge.住院号||')存在已编目的病案,不允许合并该病人。';
					End IF;
					Raise Err_Custom;
				End IF;

				Insert Into 病案主页(
					病人ID,主页ID,病人性质,医疗付款方式,费别,入院病区ID,入院科室ID,入院日期,入院病况,
					入院方式,二级院转入,住院目的,入院病床,是否陪伴,当前病况,当前病区ID,护理等级ID,
					出院科室ID,出院病床,出院日期,住院天数,出院方式,是否确诊,确诊日期,新发肿瘤,血型,
					抢救次数,成功次数,随诊标志,随诊期限,尸检标志,门诊医师,责任护士,住院医师,
					编目员编号,编目员姓名,编目日期,状态,费用和,年龄,婚姻状况,职业,国籍,学历,单位电话,
					单位邮编,单位地址,区域,家庭地址,家庭电话,户口邮编,联系人姓名,联系人关系,联系人地址,
					联系人电话,中医治疗类别,登记人,登记时间,险类,是否上传,备注,数据转出)
				Values(
					v_保留ID,v_Count,r_Merge.病人性质,r_Merge.医疗付款方式,r_Merge.费别,r_Merge.入院病区ID,
					r_Merge.入院科室ID,r_Merge.入院日期,r_Merge.入院病况,r_Merge.入院方式,r_Merge.二级院转入,
					r_Merge.住院目的,r_Merge.入院病床,r_Merge.是否陪伴,r_Merge.当前病况,r_Merge.当前病区ID,
					r_Merge.护理等级ID,r_Merge.出院科室ID,r_Merge.出院病床,r_Merge.出院日期,r_Merge.住院天数,
					r_Merge.出院方式,r_Merge.是否确诊,r_Merge.确诊日期,r_Merge.新发肿瘤,r_Merge.血型,
					r_Merge.抢救次数,r_Merge.成功次数,r_Merge.随诊标志,r_Merge.随诊期限,r_Merge.尸检标志,
					r_Merge.门诊医师,r_Merge.责任护士,r_Merge.住院医师,r_Merge.编目员编号,r_Merge.编目员姓名,
					r_Merge.编目日期,r_Merge.状态,r_Merge.费用和,r_Merge.年龄,r_Merge.婚姻状况,r_Merge.职业,
					r_Merge.国籍,r_Merge.学历,r_Merge.单位电话,r_Merge.单位邮编,r_Merge.单位地址,r_Merge.区域,
					r_Merge.家庭地址,r_Merge.家庭电话,r_Merge.户口邮编,r_Merge.联系人姓名,r_Merge.联系人关系,
					r_Merge.联系人地址,r_Merge.联系人电话,r_Merge.中医治疗类别,r_Merge.登记人,r_Merge.登记时间,
					r_Merge.险类,r_Merge.是否上传,r_Merge.备注,r_Merge.数据转出);
				
				--更新病人相关表的病人指向
				---------------------------------------------------------------
				--病人变动记录
				Update 病人变动记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;

				--病案主页从表
				Update 病案主页从表
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				
				--病人费用记录
				Update 病人费用记录
					Set 病人ID=v_保留ID,主页ID=v_Count,
						标识号=Nvl(Decode(门诊标志,1,v_门诊号,v_住院号),标识号)
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				Update H病人费用记录
					Set 病人ID=v_保留ID,主页ID=v_Count,
						标识号=Nvl(Decode(门诊标志,1,v_门诊号,v_住院号),标识号)
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;

				--病人预交记录
				Update 病人预交记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				Update H病人预交记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;

				--病人未结费用
				Update 病人未结费用
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				
				--未发药品记录
				Update 未发药品记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				
				--病人诊断记录
				Update 病人诊断记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				
				--保险结算记录(比较特殊,主页ID无外键,病人外键对应保险帐户,这里仅改主页ID,病人ID后面改)
				Update 保险结算记录
					Set 主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;

				--保险模拟结算
				Update 保险模拟结算
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;

				--病人医嘱记录(ZLHIS+)
				Update 病人医嘱记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				Update H病人医嘱记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				
				--病人过敏记录(ZLHIS+)
				Update 病人过敏记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				Update H病人过敏记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;

				--病人诊断记录(ZLHIS+)
				Update 病人诊断记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				Update H病人诊断记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;

				--病人病历记录(ZLHIS+)
				Update 病人病历记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				Update H病人病历记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;

				--病人手麻记录(ZLHIS+)
				Update 病人手麻记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				Update H病人手麻记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				
				--病人手术记录(ZLHIS+)
				Update 病人手术记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
				Update H病人手术记录
					Set 病人ID=v_保留ID,主页ID=v_Count
				Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;

				--删除已调整后的病案主页
				Delete From 病案主页 Where 病人ID=r_Merge.病人ID And 主页ID=r_Merge.主页ID;
			End IF;
			v_Count:=v_Count-1;
		End Loop;
	End IF;

	--不涉及主页ID部份的更改(非住院病人或住院病人住院前发生的)
	---------------------------------------------------------------
	--病人费用记录
	Update 病人费用记录
		Set 病人ID=v_保留ID,
			标识号=Nvl(Decode(门诊标志,2,v_住院号,v_门诊号),标识号)
	Where 病人ID=v_合并ID And 主页ID IS NULL;
	Update H病人费用记录
		Set 病人ID=v_保留ID,
			标识号=Nvl(Decode(门诊标志,2,v_住院号,v_门诊号),标识号)
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--病人预交记录
	Update 病人预交记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;
	Update H病人预交记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--未发药品记录
	Update 未发药品记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--病人诊断记录
	Update 病人诊断记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--病人医嘱记录(ZLHIS+)
	Update 病人医嘱记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;
	Update H病人医嘱记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--病人过敏记录(ZLHIS+)
	Update 病人过敏记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;
	Update H病人过敏记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--病人诊断记录(ZLHIS+)
	Update 病人诊断记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;
	Update H病人诊断记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--病人病历记录(ZLHIS+)
	Update 病人病历记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;
	Update H病人病历记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--病人手麻记录(ZLHIS+)
	Update 病人手麻记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;
	Update H病人手麻记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--病人手术记录(ZLHIS+)
	Update 病人手术记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;
	Update H病人手术记录
		Set 病人ID=v_保留ID
	Where 病人ID=v_合并ID And 主页ID IS NULL;

	--病人挂号记录(ZLHIS+)
	Update 病人挂号记录 
		Set 病人ID=v_保留ID,
			门诊号=Nvl(v_门诊号,门诊号)
	Where 病人ID=v_合并ID;

	--病人结帐记录
	Update 病人结帐记录 Set 病人ID=v_保留ID Where 病人ID=v_合并ID;
	Update H病人结帐记录 Set 病人ID=v_保留ID Where 病人ID=v_合并ID;
	
	--床位状况记录
	Update 床位状况记录 Set 病人ID=v_保留ID Where 病人ID=v_合并ID;

	--特殊病人
	Select Count(*) Into v_Count From 特殊病人 Where 病人ID=v_保留ID;
	If v_Count=0 Then
		Update 特殊病人 Set 病人ID=v_保留ID Where 病人ID=v_合并ID;	
	Else
		Delete From 特殊病人 Where 病人ID=v_合并ID;			
	End IF;
	
	--病人未结费用
	For r_Owe In c_Owe(v_合并ID) Loop
		Update 病人未结费用
			Set 金额=Nvl(金额,0)+Nvl(r_Owe.金额,0)
		Where 主页ID IS NULL And 病人ID=v_保留ID
			And Nvl(病人病区ID,0)=Nvl(r_Owe.病人病区ID,0)
			And Nvl(病人科室ID,0)=Nvl(r_Owe.病人科室ID,0)
			And Nvl(开单部门ID,0)=Nvl(r_Owe.开单部门ID,0)
			And Nvl(执行部门ID,0)=Nvl(r_Owe.执行部门ID,0)
			And Nvl(收入项目ID,0)=Nvl(r_Owe.收入项目ID,0)
			And Nvl(来源途径,0)=Nvl(r_Owe.来源途径,0);
		If SQl%RowCount=0 Then
			Insert Into 病人未结费用(
				病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
			Values(
				v_保留ID,NULL,r_Owe.病人病区ID,r_Owe.病人科室ID,r_Owe.开单部门ID,r_Owe.执行部门ID,r_Owe.收入项目ID,r_Owe.来源途径,r_Owe.金额);
		End IF;
	End Loop;
	Delete From 病人未结费用 Where 病人ID=v_合并ID;
	Delete From 病人未结费用 Where 病人ID=v_保留ID And Nvl(金额,0)=0;

	--病人余额
	For r_Spare In c_Spare(v_合并ID) Loop
		Update 病人余额
			Set 预交余额=Nvl(预交余额,0)+Nvl(r_Spare.预交余额,0),
				费用余额=Nvl(费用余额,0)+Nvl(r_Spare.费用余额,0)
		Where Nvl(性质,0)=Nvl(r_Spare.性质,0) And 病人ID=v_保留ID;
		If SQL%RowCount=0 Then
			Insert Into 病人余额(
				病人ID,性质,预交余额,费用余额)
			Values(
				v_保留ID,r_Spare.性质,r_Spare.预交余额,r_Spare.费用余额);
		End If;
	End Loop;
	Delete From 病人余额 Where 病人ID=v_合并ID;
	Delete From 病人余额 Where 病人ID=v_保留ID And Nvl(预交余额,0)=0 And Nvl(费用余额,0)=0 And 性质=1;

	--病人过敏药物
	Insert Into 病人过敏药物(
		病人ID,过敏药物ID,过敏药物)
	Select 
		v_保留ID,过敏药物ID,过敏药物
	From 病人过敏药物
	Where 病人ID=v_合并ID 
		And 过敏药物ID Not IN(Select 过敏药物ID From 病人过敏药物 Where 病人ID=v_保留ID);

	Delete From 病人过敏药物 Where 病人ID=v_合并ID;

	--门诊病案记录
	Select Count(*) Into v_Count From 门诊病案记录 Where 病人ID=v_保留ID;
	If v_Count=0 Then
		Select Count(*) Into v_Count From 门诊病案记录 Where 病人ID=v_合并ID;
		If v_Count>0 Then
			Update 门诊病案记录 Set 病人ID=v_保留ID,病案号=v_门诊号 Where 病人ID=v_合并ID;
		End IF;
	Else
		Delete From 门诊病案记录 Where 病人ID=v_合并ID;
		Update 门诊病案记录 Set 病案号=v_门诊号 Where 病人ID=v_保留ID;
	End IF;

	--住院病案记录
	Select Count(*) Into v_Count From 住院病案记录 Where 病人ID=v_保留ID;
	If v_Count=0 Then
		Select Count(*) Into v_Count From 住院病案记录 Where 病人ID=v_合并ID;
		If v_Count>0 Then
			Update 住院病案记录 Set 病人ID=v_保留ID,病案号=v_住院号 Where 病人ID=v_合并ID;
		End IF;
	Else
		Delete From 住院病案记录 Where 病人ID=v_合并ID;
		Update 住院病案记录 Set 病案号=v_住院号 Where 病人ID=v_保留ID;
	End IF;
	
	--检查病案记录(ZLHIS+)
	For r_Check In c_Check(v_合并ID) Loop
		Select Count(*) Into v_Count From 检查病案记录 Where 病人ID=v_保留ID And 检查类别=r_Check.检查类别;
		If v_Count=0 Then
			Update 检查病案记录
				Set 病人ID=v_保留ID
			Where 病人ID=v_合并ID And 检查类别=r_Check.检查类别;
		Else
			Delete From 检查病案记录 Where 病人ID=v_合并ID And 检查类别=r_Check.检查类别;
		End IF;
	End Loop;

	--保险帐户
	For r_Insure In c_Insure(v_合并ID) Loop
		Select Count(*) Into v_Count From 保险帐户 Where 病人ID=v_保留ID And 险类=r_Insure.险类;
		If v_Count>0 Then
			--两个病人具有相同险类的帐户
			--转移帐户年度信息
			For r_Year In c_Year(v_合并ID,r_Insure.险类) Loop
				Update 帐户年度信息
					Set 帐户增加累计=Nvl(帐户增加累计,0)+Nvl(r_Year.帐户增加累计,0),
						帐户支出累计=Nvl(帐户支出累计,0)+Nvl(r_Year.帐户支出累计,0),
						进入统筹累计=Nvl(进入统筹累计,0)+Nvl(r_Year.进入统筹累计,0),
						统筹报销累计=Nvl(统筹报销累计,0)+Nvl(r_Year.统筹报销累计,0),
						住院次数累计=Nvl(住院次数累计,0)+Nvl(r_Year.住院次数累计,0),
						大额统筹累计=Nvl(大额统筹累计,0)+Nvl(r_Year.大额统筹累计,0),
						起付线累计=Nvl(起付线累计,0)+Nvl(r_Year.起付线累计,0),
						本次起付线=Nvl(本次起付线,r_Year.本次起付线),
						基本统筹限额=Nvl(基本统筹限额,r_Year.基本统筹限额),
						大额统筹限额=Nvl(大额统筹限额,r_Year.大额统筹限额),
						封销信息=Nvl(封销信息,r_Year.封销信息)
				Where 病人ID=v_保留ID And 险类=r_Insure.险类 And 年度=r_Year.年度;
				If SQL%RowCount=0 Then
					Insert Into 帐户年度信息(
						病人ID,险类,年度,帐户增加累计,帐户支出累计,进入统筹累计,统筹报销累计,
						住院次数累计,本次起付线,基本统筹限额,大额统筹限额,起付线累计,大额统筹累计,封销信息)
					Values(
						v_保留ID,r_Insure.险类,r_Year.年度,r_Year.帐户增加累计,r_Year.帐户支出累计,r_Year.进入统筹累计,r_Year.统筹报销累计,
						r_Year.住院次数累计,r_Year.本次起付线,r_Year.基本统筹限额,r_Year.大额统筹限额,r_Year.起付线累计,r_Year.大额统筹累计,r_Year.封销信息);
				End IF;
			End Loop;

			--转移保险结算记录
			Update 保险结算记录 
				Set 病人ID=v_保留ID
			Where 病人ID=v_合并ID And 险类=r_Insure.险类;
			
			--读取用户指定要保留病人的帐户信息
			If v_合并ID=B病人ID_IN Then
				Open c_KeepInsure(B病人ID_IN,r_Insure.险类);
				Fetch c_KeepInsure Into r_KeepInsure;
			End IF;

			Delete From 保险帐户 Where 病人ID=v_合并ID And 险类=r_Insure.险类;	
			
			--保留用户指定要保留病人的帐户信息
			If v_合并ID=B病人ID_IN Then
				If c_KeepInsure%RowCount>0 Then
					Update 保险帐户
						Set 中心=r_KeepInsure.中心,
							卡号=r_KeepInsure.卡号,
							医保号=r_KeepInsure.医保号,
							密码=r_KeepInsure.密码,
							人员身份=r_KeepInsure.人员身份,
							单位编码=r_KeepInsure.单位编码,
							顺序号=r_KeepInsure.顺序号,
							退休证号=r_KeepInsure.退休证号,
							帐户余额=r_KeepInsure.帐户余额,
							当前状态=r_KeepInsure.当前状态,
							病种ID=r_KeepInsure.病种ID,
							在职=r_KeepInsure.在职,
							年龄段=r_KeepInsure.年龄段,
							灰度级=r_KeepInsure.灰度级,
							就诊时间=r_KeepInsure.就诊时间
					Where 险类=r_Insure.险类 And 病人ID=v_保留ID;
				End IF;
				Close c_KeepInsure;
			End IF;
		Else
			--两个病人具有不同险类的帐户(或保留病人没有)

			--两个病人分别属于不同险类时不允许合并
			Select Count(*) Into v_Count From 保险帐户 Where 病人ID=v_保留ID;
			IF v_Count>0 Then
				Close c_InfoA; Close c_InfoB;
				v_Error:='两个病人分别属于不同的保险类别，不允许合并。';
				Raise Err_Custom;
			End IF;
			
			--为避免插入重复,处理医保号和卡号
			v_卡号:=strSwitch(r_Insure.卡号,-31);
			v_医保号:=strSwitch(r_Insure.医保号,-31);
			Insert Into 保险帐户(
				病人ID,险类,中心,卡号,医保号,密码,人员身份,
				单位编码,顺序号,退休证号,帐户余额,当前状态,
				病种ID,在职,年龄段,灰度级,就诊时间)
			Values(
				v_保留ID,r_Insure.险类,r_Insure.中心,v_卡号,v_医保号,
				r_Insure.密码,r_Insure.人员身份,r_Insure.单位编码,r_Insure.顺序号,r_Insure.退休证号,
				r_Insure.帐户余额,r_Insure.当前状态,r_Insure.病种ID,r_Insure.在职,r_Insure.年龄段,
				r_Insure.灰度级,r_Insure.就诊时间);

			--转移帐户年度信息
			Update 帐户年度信息 
				Set 病人ID=v_保留ID
			Where 病人ID=v_合并ID And 险类=r_Insure.险类;

			--转移保险结算记录
			Update 保险结算记录 
				Set 病人ID=v_保留ID
			Where 病人ID=v_合并ID And 险类=r_Insure.险类;

			Delete From 保险帐户 Where 病人ID=v_合并ID And 险类=r_Insure.险类;
			
			--还原医保号和卡号
			v_卡号:=strSwitch(v_卡号,31);
			v_医保号:=strSwitch(v_医保号,31);
			Update 保险帐户
				Set 卡号=v_卡号,医保号=v_医保号
			Where 病人ID=v_保留ID And 险类=r_Insure.险类;
		End IF;
	End Loop;

	--删除实际不保留的病人信息
	Delete From 病人信息 Where 病人ID=v_合并ID;

	--根据界面选择保留病人信息
	Update 病人信息
		Set 姓名=Nvl(r_InfoB.姓名,r_InfoA.姓名),
			性别=Nvl(r_InfoB.性别,r_InfoA.性别),
			年龄=Nvl(r_InfoB.年龄,r_InfoA.年龄),
			门诊号=Nvl(r_InfoB.门诊号,r_InfoA.门诊号),
			住院号=Nvl(r_InfoB.住院号,r_InfoA.住院号),
			就诊卡号=Nvl(r_InfoB.就诊卡号,r_InfoA.就诊卡号),
			卡验证码=Decode(r_InfoB.就诊卡号,NULL,r_InfoA.卡验证码,r_InfoB.卡验证码),
			费别=Nvl(r_InfoB.费别,r_InfoA.费别),
			医疗付款方式=Nvl(r_InfoB.医疗付款方式,r_InfoA.医疗付款方式),
			出生日期=Nvl(r_InfoB.出生日期,r_InfoA.出生日期),
			出生地点=Nvl(r_InfoB.出生地点,r_InfoA.出生地点),
			身份证号=Nvl(r_InfoB.身份证号,r_InfoA.身份证号),
			身份=Nvl(r_InfoB.身份,r_InfoA.身份),
			职业=Nvl(r_InfoB.职业,r_InfoA.职业),
			民族=Nvl(r_InfoB.民族,r_InfoA.民族),
			国籍=Nvl(r_InfoB.国籍,r_InfoA.国籍),
			学历=Nvl(r_InfoB.学历,r_InfoA.学历),
			婚姻状况=Nvl(r_InfoB.婚姻状况,r_InfoA.婚姻状况),
			家庭地址=Nvl(r_InfoB.家庭地址,r_InfoA.家庭地址),
			家庭电话=Nvl(r_InfoB.家庭电话,r_InfoA.家庭电话),
			户口邮编=Nvl(r_InfoB.户口邮编,r_InfoA.户口邮编),
			联系人姓名=Nvl(r_InfoB.联系人姓名,r_InfoA.联系人姓名),
			联系人关系=Nvl(r_InfoB.联系人关系,r_InfoA.联系人关系),
			联系人地址=Nvl(r_InfoB.联系人地址,r_InfoA.联系人地址),
			联系人电话=Nvl(r_InfoB.联系人电话,r_InfoA.联系人电话),
			合同单位id=Nvl(r_InfoB.合同单位id,r_InfoA.合同单位id),
			工作单位=Nvl(r_InfoB.工作单位,r_InfoA.工作单位),
			单位电话=Nvl(r_InfoB.单位电话,r_InfoA.单位电话),
			单位邮编=Nvl(r_InfoB.单位邮编,r_InfoA.单位邮编),
			单位开户行=Nvl(r_InfoB.单位开户行,r_InfoA.单位开户行),
			单位帐号=Nvl(r_InfoB.单位帐号,r_InfoA.单位帐号),
			担保人=Nvl(r_InfoB.担保人,r_InfoA.担保人),
			担保额=Decode(r_InfoB.担保人,NULL,r_InfoA.担保额,r_InfoB.担保额),
			担保性质=Decode(r_InfoB.担保人,NULL,r_InfoA.担保性质,r_InfoB.担保性质),
			就诊时间=Nvl(r_InfoB.就诊时间,r_InfoA.就诊时间),
			就诊状态=Nvl(r_InfoB.就诊状态,r_InfoA.就诊状态),
			就诊诊室=Nvl(r_InfoB.就诊诊室,r_InfoA.就诊诊室),
			险类=Nvl(r_InfoB.险类,r_InfoA.险类),
			登记时间=Nvl(r_InfoB.登记时间,r_InfoA.登记时间),
			住院次数=NULL,当前床号=NULL,
			当前科室ID=NULL,当前病区ID=NULL,
			入院时间=NULL,出院时间=NULL
	Where 病人ID=v_保留ID;

	OPEN c_Info(v_保留ID);
	FETCH c_Info Into r_Info;
	IF c_Info%Rowcount>0 THEN
		Update 病人信息
			Set 住院次数=r_Info.主页ID,
				当前床号=Decode(r_Info.出院日期,NULL,r_Info.出院病床,NULL),
				当前病区ID=Decode(r_Info.出院日期,NULL,r_Info.当前病区ID,NULL),
				当前科室ID=Decode(r_Info.出院日期,NULL,r_Info.出院科室ID,NULL),
				入院时间=r_Info.入院日期,出院时间=r_Info.出院日期
		Where 病人ID=v_保留ID;
	End IF;
	Close c_Info;

	Close c_InfoA;
	Close c_InfoB;
EXCEPTION
	WHEN Err_Custom THEN Raise_application_error(-20101,'[ZLSOFT]' || v_Error || '[ZLSOFT]');
	WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_病人信息_Merge;
/


CREATE OR REPLACE Procedure ZL_病人变动记录_转住院(
--功能：将住院留观病人转为住院病人
    病人ID_IN    病案主页.病人ID%Type,
    主页ID_IN    病案主页.主页ID%Type,
    住院号_IN    病人信息.住院号%Type
) IS
    Cursor c_Info IS
        Select * From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;
    r_Info    c_Info%RowType;

    v_Count        Number;
    v_Date        Date;
    v_Temp        Varchar2(255);
    v_人员编号    病人费用记录.操作员编号%Type;
    v_人员姓名    病人费用记录.操作员姓名%Type;

    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --并发操作检查
    Select Nvl(状态,0) Into v_Count From 病案主页 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 病人性质=2;
    if v_Count=1 Then
        v_Error:='病人当前尚未入科,不能转为住院病人。请先将病人入科后再试。';
        Raise Err_Custom;
    ElsIf v_Count=2 Then
        v_Error:='病人当前正在转科,不能转为住院病人。请先将病人转科或取消转科后再试。';
        Raise Err_Custom;
    End IF;

    Select Sysdate Into v_Date From Dual;
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

    Open c_Info;--必须先打开

    --取消上次变动
    Update 病人变动记录
        Set 终止时间=v_Date,终止原因=9,终止人员=v_人员姓名
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;

    --产生新变动
    Fetch c_Info Into r_Info;
    if c_Info%RowCount=0 Then
        Close c_Info;
        v_Error:='未发现该病人当前有效的变动记录！';
        Raise Err_Custom;
    End IF;

    --产生变动记录
    While c_Info%Found Loop
        Insert Into 病人变动记录(
            病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,
            护理等级ID,床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
        Values(
            病人ID_IN,主页ID_IN,v_Date,9,r_Info.附加床位,r_Info.病区ID,
            r_Info.科室ID,r_Info.护理等级ID,r_Info.床位等级ID,r_Info.床号,
            r_Info.责任护士,r_Info.经治医师,r_Info.主治医师,r_Info.主任医师,r_Info.病情,v_人员编号,v_人员姓名);
        Fetch c_Info Into r_Info;
    End Loop;

    Close c_Info;

    Update 病案主页 Set 病人性质=0 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;    
    Update 病人信息 Set 住院号=住院号_IN Where 病人ID=病人ID_IN;

    --并发操作检查
    Select Count(*) Into v_Count 
    From 病人变动记录 
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
        AND NVL(附加床位,0)=0 And 开始时间 IS NOT Null
        AND 终止时间 is Null;

    if v_Count > 1 Then
        v_Error:='发现病人存在非法的变动记录,当前操作不能继续！'||Chr(13)||Chr(10)||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人变动记录_转住院;
/

Create Or Replace Procedure zl_病人变动记录_Nurse(
    病人ID_IN        病案主页.病人ID%Type,
    主页ID_IN        病案主页.主页ID%Type,
    护理ID_IN        病人变动记录.护理等级ID%Type,
    生效时间_IN        病人变动记录.开始时间%Type,
    操作员编号_IN    病人变动记录.操作员编号%Type,
    操作员姓名_IN    病人变动记录.操作员姓名%Type
)
AS
-----------------------------------------------------------
--说明：更改病人护理等级
-----------------------------------------------------------
    Cursor c_OldInfo IS
        Select * From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;

    r_OldInfo    c_OldInfo%RowType;
    v_Count        Number;
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    Open c_OldInfo;--必须先打开

    --取消上次变动
    Update 病人变动记录
        Set 终止时间=生效时间_IN,终止原因=6,终止人员=操作员姓名_IN
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;

    --产生新变动
    Fetch c_OldInfo Into r_OldInfo;

    if c_OldInfo%RowCount=0 Then
        Close c_OldInfo;
        v_Error:='未发现该病人当前有效的变动记录！';
        Raise Err_Custom;
    End IF;

    While c_OldInfo%Found Loop
        Insert Into 病人变动记录(
            病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,
            护理等级ID,床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
        Values(
            病人ID_IN,主页ID_IN,生效时间_IN,6,r_OldInfo.附加床位,r_OldInfo.病区ID,
            r_OldInfo.科室ID,护理ID_IN,r_OldInfo.床位等级ID,r_OldInfo.床号,
            r_OldInfo.责任护士,r_OldInfo.经治医师,r_OldInfo.主治医师,r_OldInfo.主任医师,r_OldInfo.病情,操作员编号_IN,操作员姓名_IN);
        Fetch c_OldInfo Into r_OldInfo;
    End Loop;

    Close c_OldInfo;

    Update 病案主页 Set 护理等级ID=护理ID_IN
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 出院日期 is Null;

    --并发操作检查
    Select Count(*) Into v_Count
    From 病人变动记录
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
        AND NVL(附加床位,0)=0 And 开始时间 IS NOT Null
        AND 终止时间 is Null;

    if v_Count > 1 Then
        v_Error:='发现病人存在非法的变动记录,当前操作不能继续！'||Chr(13)||Chr(10)||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_病人变动记录_Nurse;
/

Create Or Replace Procedure zl_病人变动记录_BedLevel(
    病人ID_IN        病案主页.病人ID%Type,
    主页ID_IN        病案主页.主页ID%Type,
    床号_IN            病人变动记录.床号%Type,
    等级ID_IN        病人变动记录.床位等级ID%Type,
    生效时间_IN        病人变动记录.开始时间%Type,
    操作员编号_IN    病人变动记录.操作员编号%Type,
    操作员姓名_IN    病人变动记录.操作员姓名%Type
)
AS
-----------------------------------------------------------
--说明：更改床位等级
-----------------------------------------------------------
    Cursor c_OldInfo IS
        Select * From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;

    r_OldInfo    c_OldInfo%RowType;
    v_Count        Number;
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    Open c_OldInfo;--必须先打开

    --取消上次变动
    Update 病人变动记录 
        Set 终止时间=生效时间_IN,终止原因=5,终止人员=操作员姓名_IN
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;

    --产生新变动
    Fetch c_OldInfo Into r_OldInfo;

    if c_OldInfo%RowCount=0 Then
        Close c_OldInfo;
        v_Error:='未发现该病人当前有效的变动记录！';
        Raise Err_Custom;
    End IF;

    While c_OldInfo%Found Loop
        if r_OldInfo.床号=床号_IN Then
            Update 床位状况记录 Set 等级ID=等级ID_IN
            Where 病区ID=r_OldInfo.病区ID And 床号=床号_IN;

            Insert Into 病人变动记录(
                病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,
                护理等级ID,床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
            Values(
                病人ID_IN,主页ID_IN,生效时间_IN,5,r_OldInfo.附加床位,r_OldInfo.病区ID,
                r_OldInfo.科室ID,r_OldInfo.护理等级ID,等级ID_IN,床号_IN,r_OldInfo.责任护士,
                r_OldInfo.经治医师,r_OldInfo.主治医师,r_OldInfo.主任医师,r_OldInfo.病情,操作员编号_IN,操作员姓名_IN);
        ELSE
            Insert Into 病人变动记录(
                病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,护理等级ID,
                床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
            Values(
                病人ID_IN,主页ID_IN,生效时间_IN,5,r_OldInfo.附加床位,r_OldInfo.病区ID,
                r_OldInfo.科室ID,r_OldInfo.护理等级ID,r_OldInfo.床位等级ID,r_OldInfo.床号,
                r_OldInfo.责任护士,r_OldInfo.经治医师,r_OldInfo.主治医师,r_OldInfo.主任医师,r_OldInfo.病情,操作员编号_IN,操作员姓名_IN);
        End IF;

        Fetch c_OldInfo Into r_OldInfo;
    End Loop;

    Close c_OldInfo;
    --并发操作检查
    Select Count(*) Into v_Count From 病人变动记录
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
        AND NVL(附加床位,0)=0 And 开始时间 IS NOT Null
        AND 终止时间 is Null;

    if v_Count > 1 Then
        v_Error:='发现病人存在非法的变动记录,当前操作不能继续！'||Chr(13)||Chr(10) ||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_病人变动记录_BedLevel;
/

Create Or Replace Procedure zl_病人变动记录_Move(
    病人ID_IN        病案主页.病人ID%Type,
    主页ID_IN        病案主页.主页ID%Type,
    换床时间_IN        病人变动记录.开始时间%Type,
    床号_IN            Varchar2,
    操作员编号_IN    病人变动记录.操作员编号%Type,
    操作员姓名_IN    病人变动记录.操作员姓名%Type,
    病区ID_IN    病人变动记录.病区ID%Type:=Null
)
AS
-----------------------------------------------------------
--说明：病人换床
--参数：
--       床号=Null:家庭病床;"床号1,床号2,....床号n"
-----------------------------------------------------------
    Cursor c_BedInfo IS
        Select 病区ID,床号 From 床位状况记录 Where 病人ID=病人ID_IN;

    Cursor c_OldInfo IS
        Select * From 病人变动记录 Where 病人ID=病人ID_IN
            AND 主页ID=主页ID_IN And 终止时间 is Null And NVL(附加床位,0)=0;

    r_OldInfo    c_OldInfo%RowType;
    v_床号串    Varchar2(255);
    v_床号        病人变动记录.床号%Type;
    v_等级ID    床位状况记录.等级ID%Type;
    v_病区科室独立 Number(1);
    v_Tmp        Number;
    v_Count        Number;
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --并发操作检查
    Select Count(*) Into v_Count From 病案主页
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And NVL(状态,0)=0;

    if v_Count=0 Then
        v_Error:='病人当前不处于正常住院状态,可能尚未入科,操作不能继续！'||Chr(13) ||Chr(10) ||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
        Raise Err_Custom;
    End IF;

    v_病区科室独立:=To_Number(Nvl(ZL_GetSysParameter(99),0));    --病人换床时可以换病区

    --退除病人原床位
    For r_bedrow IN c_BedInfo Loop
        Update 床位状况记录 
            Set 状态='空床',
                病人ID=Null,
                科室ID=Decode(共用,1,NULL,科室ID)
        Where 病区ID=r_bedrow.病区ID And 床号=r_bedrow.床号;
    End Loop;    

    Open c_OldInfo;
    Fetch c_OldInfo Into r_OldInfo;

    --取消上批变动记录
    Update 病人变动记录
        Set 终止时间=换床时间_IN,终止原因=4,终止人员=操作员姓名_IN
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;

    --新增病人床位
    if 床号_IN is Null Then
        --家庭病床
        Insert Into 病人变动记录(
            病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,
            护理等级ID,床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
        Values(
            病人ID_IN,主页ID_IN,换床时间_IN,4,0,r_OldInfo.病区ID,r_OldInfo.科室ID,
            r_OldInfo.护理等级ID,Null,Null,r_OldInfo.责任护士,r_OldInfo.经治医师,r_OldInfo.主治医师,r_OldInfo.主任医师,r_OldInfo.病情,操作员编号_IN,操作员姓名_IN);
    ELSE
        --入住一张或多张病床病床
        v_Count:=0;
        v_床号串:=床号_IN||',';

        While v_床号串 IS NOT Null Loop
            v_床号:=to_Number(Substr(v_床号串,1,Instr(v_床号串,',')-1));
            
            Select 等级ID Into v_等级ID From 床位状况记录 Where 病区ID=Decode(v_病区科室独立,1,病区ID_IN,r_OldInfo.病区ID) And 床号=v_床号;

            Insert Into 病人变动记录(
                病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,护理等级ID,
                床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
            Values(
                病人ID_IN,主页ID_IN,换床时间_IN,4,Decode(v_Count,0,0,1),
                Decode(v_病区科室独立,1,病区ID_IN,r_OldInfo.病区ID),r_OldInfo.科室ID,r_OldInfo.护理等级ID,v_等级ID,
                v_床号,r_OldInfo.责任护士,r_OldInfo.经治医师,r_OldInfo.主治医师,r_OldInfo.主任医师,r_OldInfo.病情,操作员编号_IN,操作员姓名_IN);

            Select Count(*) Into v_Tmp From 床位状况记录 Where 病区ID=Decode(v_病区科室独立,1,病区ID_IN,r_OldInfo.病区ID) And 床号=v_床号 And 状态='空床';
            if v_Tmp=0 Then
                v_Error:='操作失败,床位 '||v_床号||' 不是空床！';
                Raise Err_Custom;
            End IF;

            Update 床位状况记录 
                Set 状态='占用',
                    病人ID=病人ID_IN ,
                    科室ID=Decode(共用,1,r_OldInfo.科室ID,科室ID)
            Where 病区ID=Decode(v_病区科室独立,1,病区ID_IN,r_OldInfo.病区ID) And 床号=v_床号;

            v_床号串:=Substr(v_床号串,Instr(v_床号串,',')+1);
            v_Count:=v_Count+1;
        End Loop;
    End IF;

    Close c_OldInfo;
    --病人信息、病案主页(记录第一张床位)
    v_床号串:=床号_IN||',';
    v_床号:=to_Number(Substr(v_床号串,1,Instr(v_床号串,',')-1));

    IF v_病区科室独立=1 THEN 
        Update 病案主页 Set 出院病床=v_床号,当前病区ID=病区ID_IN Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
        Update 病人信息 Set 当前床号=v_床号,当前病区ID=病区ID_IN Where 病人ID=病人ID_IN;
    ELSE 
        Update 病案主页 Set 出院病床=v_床号 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
        Update 病人信息 Set 当前床号=v_床号 Where 病人ID=病人ID_IN;
    END IF;

    --并发操作检查
    Select Count(*) Into v_Count From 病人变动记录
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
        AND NVL(附加床位,0)=0 And 开始时间 IS NOT Null
        AND 终止时间 is Null;

    if v_Count > 1 Then
        v_Error:='发现病人存在非法的变动记录,当前操作不能继续！'||Chr(13)||Chr(10) ||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_病人变动记录_Move;
/

Create Or Replace Procedure zl_病人变动记录_InDept(
    病人ID_IN        病案主页.病人ID%Type,
    主页ID_IN        病案主页.主页ID%Type,
    床号_IN            Varchar2,
    病区ID_IN        病案主页.当前病区ID%Type,
    科室ID_IN        病案主页.出院科室ID%Type,
    护理等级ID_IN    病案主页.护理等级ID%Type,
    当前病况_IN        病案主页.当前病况%Type,
    责任护士_IN        病案主页.责任护士%Type,
    门诊医师_IN        病案主页.门诊医师%Type,
    住院医师_IN        病案主页.住院医师%Type,
    是否陪伴_IN        病案主页.是否陪伴%Type,
    入科时间_IN        病人变动记录.开始时间%Type,
    操作员编号_IN    人员表.编号%Type,
    操作员姓名_IN    人员表.姓名%Type,
    入院_IN            Number,
    主治医师_IN        病案主页.住院医师%Type:=Null,
    主任医师_IN        病案主页.住院医师%Type:=Null
)
AS
-----------------------------------------------------------
--说明：完成病人入院或转科入科处理。
--参数：
--       入院_IN:病人是入院还是转科入科。
--       床号_IN:为空表示家庭病床,否则为"床号1,床号2,...床号n",多个床号时,表示包房。
-----------------------------------------------------------
    Cursor c_BedInfo IS
        Select 病区ID,床号 From 床位状况记录 Where 病人ID=病人ID_IN;

    v_床号        Varchar2(255);
    v_当前床号		床位状况记录.床号%Type;
    v_等级ID		床位状况记录.等级ID%Type;
    v_病区ID		病案主页.当前病区ID%Type;
    v_终止人员		病人变动记录.终止人员%Type;
    v_Count        Number;
    Err_Custom    Exception;
    v_Error        Varchar2(255);
Begin
    --包房时,病床只取一个填写出院病床。
    v_床号:=床号_IN||',';
    v_床号:=Substr(v_床号,1,Instr(v_床号,',')-1);

    --都要更新病人信息
    Update 病人信息
        Set 当前病区ID=病区ID_IN,
            当前科室ID=科室ID_IN,
            当前床号=to_Number(v_床号)
    Where 病人ID=病人ID_IN;

    if 入院_IN=1 Then
        --入院入科
        Select Count(*) Into v_Count From 病案主页 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 状态=1;

        if v_Count=0 Then
            v_Error:='病人当前不处于入院状态,可能已经撤入院，操作不能继续！'||Chr(13) ||Chr(10) ||'这可能是由于网络并发操作引起的，请刷新病人状态后再试！';
            Raise Err_Custom;
        End IF;

        Select 入院病区ID Into v_病区ID From 病案主页 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;

        if v_病区ID <> 病区ID_IN Then
            v_Error:='当前入住病区与病人登记病区不一致,病人状态已经更改,操作不能继续！'||Chr(13) ||Chr(10) ||'这可能是由于网络并发操作引起的，请刷新病人状态后再试！';
            Raise Err_Custom;
        End IF;

        --病案主页
        --同时更改了入院登记时的科室,病区,病况
        Update 病案主页
            Set 入院科室ID=科室ID_IN,
                入院病区ID=病区ID_IN,
                入院病况=当前病况_IN,
                状态=0,
                入院病床=to_Number(v_床号),
                出院病床=to_Number(v_床号),
                当前病区ID=病区ID_IN,
                出院科室ID=科室ID_IN,
                护理等级ID=Decode(护理等级ID_IN,0,Null,护理等级ID_IN),
                当前病况=当前病况_IN,
                责任护士=责任护士_IN,
                门诊医师=门诊医师_IN,
                住院医师=住院医师_IN,
                是否陪伴=是否陪伴_IN
         Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
         Insert Into 病案主页从表 (病人ID,主页ID,信息名,信息值) 
                Values (病人ID_IN,主页ID_IN,'主治医师',主治医师_IN);
         Insert Into 病案主页从表 (病人ID,主页ID,信息名,信息值) 
                Values (病人ID_IN,主页ID_IN,'主任医师',主任医师_IN);

        --记录上一步的终止操作人员
        Update 病人变动记录 
            Set 终止时间=入科时间_IN,终止原因=2,终止人员=操作员姓名_IN
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 开始时间 IS Not NULL And 终止时间 is Null;

        if 床号_IN is Null Then
            --仅家庭病床
            Insert Into 病人变动记录(
                病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,
                护理等级ID,床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,
                操作员姓名)
            Values(
                病人ID_IN,主页ID_IN,入科时间_IN,2,0,病区ID_IN,科室ID_IN,
                Decode(护理等级ID_IN,0,Null,护理等级ID_IN),Null,Null,
                责任护士_IN,住院医师_IN,主治医师_IN,主任医师_IN,当前病况_IN,操作员编号_IN,操作员姓名_IN);
        ELSE
            --多张床位
            v_Count:=0;
            v_床号:=床号_IN||',';

            While v_床号 IS NOT Null Loop
                v_当前床号:=to_Number(Substr(v_床号,1,Instr(v_床号,',')-1));
                Select 等级ID Into v_等级ID From 床位状况记录 Where 病区ID=病区ID_IN And 床号=v_当前床号;

                Insert Into 病人变动记录(
                    病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,护理等级ID,
                    床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
                Values(
                    病人ID_IN,主页ID_IN,入科时间_IN,2,Decode(v_Count,0,0,1),
                    病区ID_IN,科室ID_IN,Decode(护理等级ID_IN,0,Null,护理等级ID_IN),
                    v_等级ID,v_当前床号,责任护士_IN,住院医师_IN,主治医师_IN,主任医师_IN,当前病况_IN,操作员编号_IN,操作员姓名_IN);

                Select Count(*) Into v_Count From 床位状况记录 Where 病区ID=病区ID_IN And 床号=v_当前床号 And 状态='空床';

                if v_Count=0 Then
                    v_Error:='操作失败,床位 '||v_当前床号||' 不是空床！';
                    Raise Err_Custom;
                End IF;

                Update 床位状况记录 Set 状态='占用',病人ID=病人ID_IN,科室ID=Decode(共用,1,科室ID_IN,科室ID) Where 病区ID=病区ID_IN And 床号=v_当前床号;

                v_床号:=Substr(v_床号,Instr(v_床号,',')+1);
                v_Count:=v_Count+1;
            End Loop;
        End IF;
    ELSE
        --转科入科
        Select Count(*) Into v_Count From 病案主页 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 状态=2;

        if v_Count=0 Then
            v_Error:='病人当前不处于转科状态,可能已经撤转科，操作不能继续！'||Chr(13) ||Chr(10) ||'这可能是由于网络并发操作引起的，请刷新病人状态后再试！';
            Raise Err_Custom;
        End IF;

        Select 病区ID,操作员姓名 Into v_病区ID,v_终止人员 From 病人变动记录       --病区与科室独立时,临时记录里没有填,为Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 开始时间 is Null And 终止时间 is Null;

        if v_病区ID <> 病区ID_IN Then                         --病区与科室独立时,由于是Null判断,所以不会提示
            v_Error:='当前入住病区与病人登记病区不一致,病人状态已经更改,操作不能继续！' ||Chr(13) ||Chr(10) ||'这可能是由于网络并发操作引起的，请刷新病人状态后再试！';
            Raise Err_Custom;
        End IF;

        --病案主页
        Update 病案主页
            Set 状态=0,
                出院病床=to_Number(v_床号),
                当前病区ID=病区ID_IN,
                出院科室ID=科室ID_IN,
                护理等级ID=Decode(护理等级ID_IN,0,Null,护理等级ID_IN),
                当前病况=当前病况_IN,
                责任护士=责任护士_IN,
                门诊医师=门诊医师_IN,
                住院医师=住院医师_IN,
                是否陪伴=是否陪伴_IN
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;

         Update 病案主页从表 Set 信息值=主治医师_IN Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主治医师';
         IF SQL%RowCount=0 Then
             Insert Into 病案主页从表 (病人ID,主页ID,信息名,信息值) Values (病人ID_IN,主页ID_IN,'主治医师',主治医师_IN);
         End IF;
         Update 病案主页从表 Set 信息值=主任医师_IN Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主任医师';
         IF SQL%RowCount=0 Then
             Insert Into 病案主页从表 (病人ID,主页ID,信息名,信息值) Values (病人ID_IN,主页ID_IN,'主任医师',主任医师_IN);
         End IF;


        --退除病人当前床位
        For r_bedrow IN c_BedInfo Loop
            Update 床位状况记录 Set 状态='空床',病人ID=Null,科室ID=Decode(共用,1,NULL,科室ID) Where 病区ID=r_bedrow.病区ID And 床号=r_bedrow.床号;
        End Loop;

        --记录上一步的终止操作人员
        Update 病人变动记录 
            Set 终止时间=入科时间_IN,终止原因=3,终止人员=v_终止人员
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 开始时间 IS NOT Null And 终止时间 is Null;

        Delete From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 开始原因=3 And 开始时间 is Null And 终止时间 is Null;

        --新的床位记录
        if 床号_IN is Null Then
            --仅家庭病床
            Insert Into 病人变动记录(
                病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,
                护理等级ID,床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
            Values(
                病人ID_IN,主页ID_IN,入科时间_IN,3,0,病区ID_IN,科室ID_IN,
                Decode(护理等级ID_IN,0,Null,护理等级ID_IN),Null,
                Null,责任护士_IN,住院医师_IN,主治医师_IN,主任医师_IN,当前病况_IN,操作员编号_IN,操作员姓名_IN);
        ELSE
            v_Count:=0;
            v_床号:=床号_IN||',';

            While v_床号 IS NOT Null Loop
                v_当前床号:=to_Number(Substr(v_床号,1,Instr(v_床号,',')-1));
                Select 等级ID Into v_等级ID From 床位状况记录 Where 病区ID=病区ID_IN And 床号=v_当前床号;

                Insert Into 病人变动记录(
                    病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,
                    护理等级ID,床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
                Values(
                    病人ID_IN,主页ID_IN,入科时间_IN,3,Decode(v_Count,0,0,1),
                    病区ID_IN,科室ID_IN,Decode(护理等级ID_IN,0,Null,护理等级ID_IN),
                    v_等级ID,v_当前床号,责任护士_IN,住院医师_IN,主治医师_IN,主任医师_IN,当前病况_IN,操作员编号_IN,操作员姓名_IN);

                Select Count(*) Into v_Count From 床位状况记录 Where 病区ID=病区ID_IN And 床号=v_当前床号 And 状态='空床';

                if v_Count=0 Then
                    v_Error:='操作失败,床位 '||v_当前床号||' 不是空床！';
                    Raise Err_Custom;
                End IF;

                Update 床位状况记录 Set 状态='占用',病人ID=病人ID_IN,科室ID=Decode(共用,1,科室ID_IN,科室ID) Where 病区ID=病区ID_IN And 床号=v_当前床号;

                v_床号:=Substr(v_床号,Instr(v_床号,',')+1);
                v_Count:=v_Count+1;
            End Loop;
        End IF;
    End IF;

    --并发操作检查
    Select Count(*) Into v_Count From 病人变动记录
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
        AND NVL(附加床位,0)=0 And 开始时间 IS NOT Null
        AND 终止时间 is Null;

    if v_Count > 1 Then
        v_Error:='发现病人存在非法的变动记录,当前操作不能继续！'||Chr(13)||Chr(10) ||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_病人变动记录_InDept;
/

Create Or Replace Procedure zl_病人变动记录_Undo(
    病人ID_IN        病案主页.病人ID%Type,
    主页ID_IN        病案主页.主页ID%Type,
    操作员编号_IN    病人变动记录.操作员编号%Type,
    操作员姓名_IN    病人变动记录.操作员姓名%Type,
    数据_IN            Varchar2:=NULL--附加数据,不一定用得到
)
AS
 -----------------------------------------------------------
 --说明：1.撤消病人最近一次的变动
 --        2.前提：当病人包床时,对其中一张床位作变动,则所有床位相应产生变动
 -----------------------------------------------------------
    --要撤消的变动记录(如果包床,可能多条)
    Cursor c_CurLog IS
        Select * From 病人变动记录 
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND(终止时间 is Null Or 终止原因=1)
        Order by 终止时间 DESC,开始时间 DESC;
    r_CurLogRow c_CurLog%RowType;

    --撤消后要恢复的变动记录(如果包床,可能多条)
    Cursor c_PreLog(
        v_终止时间 病人变动记录.终止时间%Type,
        v_终止原因 病人变动记录.终止原因%Type) IS
        Select * From 病人变动记录
        Where 病人ID=病人ID_IN
            AND 主页ID=主页ID_IN
            AND 终止时间=v_终止时间
            AND 终止原因=v_终止原因
        Order by 终止时间 DESC,开始时间 DESC;
    r_PreLogRow        c_PreLog%RowType;

    v_开始时间        病人变动记录.开始时间%Type;
    v_开始原因        病人变动记录.开始原因%Type;       
    v_终止人员		病人变动记录.终止人员%Type; 

    v_病区科室独立 Number(1);
    v_Count            Number;
    Err_Custom        Exception;
    v_Error            Varchar2(255);
Begin
    Open c_CurLog;
    Fetch c_CurLog Into r_CurLogRow;
    If c_CurLog%RowCount=0 Then
        v_Error:='[ZLSOFT]病人当前没有可以撤消的操作！[ZLSOFT]';
        Close c_CurLog;
        Raise Err_Custom;
    End IF;

    if r_CurLogRow.终止时间 is Null And r_CurLogRow.开始时间 is Null And r_CurLogRow.开始原因=3 Then
        --撤消转科(标志)
        Delete From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 开始时间 is Null And 终止时间 is Null And 开始原因=3;

        Update 病案主页 Set 状态=0 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;

        Close c_CurLog;
    Elsif r_CurLogRow.终止时间 IS NOT Null And r_CurLogRow.终止原因=1 Then
        --撤消出院
        Close c_CurLog;

        --处理床位
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.床号 IS NOT Null Then
                Select Count(*) Into v_Count From 床位状况记录 Where 病区ID=r_CurLogRow.病区ID AND 床号=r_CurLogRow.床号 And 状态='空床';
                if v_Count=0 Then
                    v_Error:='[ZLSOFT]该病人出院前所入住的病床 '||r_CurLogRow.床号 ||' 当前非空床或已经撤消！[ZLSOFT]';
                    Raise Err_Custom;
                End IF;

                --重新占用床位(床位信息如果有变动,则强行恢复)
                Update 床位状况记录
                    Set 状态='占用',
                        病人ID=病人ID_IN,
                        等级ID=r_CurLogRow.床位等级ID,
                        科室ID=r_CurLogRow.科室ID--强行恢复以前的科室,共用床也不用处理了。
                Where 病区ID=r_CurLogRow.病区ID And 床号=r_CurLogRow.床号;
            End IF;

            if NVL(r_CurLogRow.附加床位,0)=0 Then
                Update 病人信息
                    Set 出院时间=Null,
                        当前病区ID=r_CurLogRow.病区ID,
                        当前科室ID=r_CurLogRow.科室ID,
                        当前床号=r_CurLogRow.床号
                Where 病人ID=病人ID_IN;
            End IF;
        End Loop;

        --恢复入院
        Update 病人变动记录
            Set 终止时间=Null,终止原因=Null,终止人员=null,
                上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 终止时间 IS NOT Null And 终止原因=1;

        Select 开始原因 Into v_开始原因 From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            And 终止时间 is Null And Nvl(附加床位,0)=0;

        Update 病案主页
            Set 状态=Decode(v_开始原因,10,3,状态),
                出院日期=Null,出院方式=Null,
                随诊标志=Null,随诊期限=Null
         Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;

        --删除出院诊断
        Delete From 病人诊断记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 诊断类型=3 AND 记录来源=2;
    Elsif r_CurLogRow.开始原因=1 Then
        --撤消入科(入院同时入科)
        v_开始时间:=r_CurLogRow.开始时间;
        Close c_CurLog;

        --退除当前床位
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.床号 IS NOT Null Then
                Update 床位状况记录 
                    Set 状态='空床',
                    病人ID=Null,
                    科室ID=Decode(共用,1,NULL,科室ID)
                Where 病区ID=r_CurLogRow.病区ID
                    AND 床号=r_CurLogRow.床号;
            End IF;
        End Loop;

        --相关信息还原
        Update 病案主页
            Set 入院病床=Null,出院病床=Null,状态=1
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;

        Update 病人信息 Set 当前床号=Null Where 病人ID=病人ID_IN;

        --恢复变动(入院同时入科不会有包床)
        --因为是同一条记录中的撤消,所以不处理人员
        Update 病人变动记录
            Set 床位等级ID=Null,床号=Null,
                责任护士=Null,经治医师=Null,
                上次计算时间=Null
        Where 病人ID=病人ID_IN AND 主页ID=主页ID_IN
            AND 开始原因=1 AND 终止时间 is Null;
    Elsif r_CurLogRow.开始原因=2 Then
        --撤消入院入科
        v_开始时间:=r_CurLogRow.开始时间;
        Close c_CurLog;

        --退除当前床位
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.床号 IS NOT Null Then
                Update 床位状况记录 
                    Set 状态='空床',
                        病人ID=Null,
                        科室ID=Decode(共用,1,NULL,科室ID)
                Where 病区ID=r_CurLogRow.病区ID 
                    And 床号=r_CurLogRow.床号;
            End IF;
        End Loop;

        --相关信息还原
        Open c_PreLog(v_开始时间,2);
        Fetch c_PreLog Into r_PreLogRow;
        Update 病案主页 Set 入院病床=Null,出院病床=Null,状态=1,
                        当前病况=r_PreLogRow.病情,入院病况=r_PreLogRow.病情,护理等级ID=r_PreLogRow.护理等级ID 
                    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
        Close c_PreLog;
        
        Update 病人信息 Set 当前床号=Null Where 病人ID=病人ID_IN;
        Delete 病案主页从表 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And (信息名='主治医师' Or 信息名='主任医师');

        --恢复变动
        Delete From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 开始原因=2 And 终止时间 is Null;

        Update 病人变动记录
            Set 终止时间=Null,终止原因=Null,终止人员=null,
                上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 终止原因=2 And 终止时间=v_开始时间;
    Elsif r_CurLogRow.开始原因=3 Then
        --撤消转科入科
        v_开始时间:=r_CurLogRow.开始时间;
        Close c_CurLog;

        --退除当前床位
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.床号 IS NOT Null Then
                Update 床位状况记录 
                    Set 状态='空床',
                        病人ID=Null,
                        科室ID=Decode(共用,1,NULL,科室ID)
                Where 病区ID=r_CurLogRow.病区ID 
                    And 床号=r_CurLogRow.床号;
            End IF;
        End Loop;

        --检查及还原原床位
        For r_PreLogRow IN c_PreLog(v_开始时间,3) Loop
            if r_PreLogRow.床号 IS NOT Null Then
                Select Count(*) Into v_Count From 床位状况记录 Where 病区ID=r_PreLogRow.病区ID AND 床号=r_PreLogRow.床号 And 状态='空床';
                if v_Count=0 Then
                    v_Error:='[ZLSOFT]病人转入该科室前的床位 '||r_PreLogRow.床号 ||' 当前非空床或已经撤消！[ZLSOFT]';
                    Raise Err_Custom;
                End IF;

                Update 床位状况记录
                    Set 状态='占用',
                        病人ID=病人ID_IN,
                        科室ID=r_PreLogRow.科室ID--强行恢复以前的科室,共用床也不用处理了。
                Where 病区ID=r_PreLogRow.病区ID 
                    And 床号=r_PreLogRow.床号;
            End IF;

            --相关信息还原
            if NVL(r_PreLogRow.附加床位,0)=0 Then
                Update 病案主页
                    Set 状态=2,
                        当前病区ID=r_PreLogRow.病区ID,
                        出院科室ID=r_PreLogRow.科室ID,
                        出院病床=r_PreLogRow.床号,
                        护理等级ID=r_PreLogRow.护理等级ID,
                        责任护士=r_PreLogRow.责任护士,
                        住院医师=r_PreLogRow.经治医师,
                        当前病况=r_CurLogRow.病情
                Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;

                Update 病案主页从表
                    SET 信息值=r_PreLogRow.主治医师
                Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主治医师';
                Update 病案主页从表
                    SET 信息值=r_PreLogRow.主任医师
                Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主任医师';

                Update 病人信息
                    Set 当前病区ID=r_PreLogRow.病区ID,
                        当前科室ID=r_PreLogRow.科室ID,
                        当前床号=r_PreLogRow.床号
                Where 病人ID=病人ID_IN;
            End IF;
        End Loop;

        --恢复变动(恢复到临时转科标记状态)
        Delete From 病人变动记录
        Where 附加床位=1 And 病人ID=病人ID_IN
            AND 主页ID=主页ID_IN And 开始原因=3 And 终止时间 is Null;
        	
		   Select 终止人员 Into v_终止人员 From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN 
			  And 终止原因=3 And 终止时间=v_开始时间 And Nvl(附加床位,0)=0;

        v_病区科室独立:=To_Number(Nvl(ZL_GetSysParameter(99),0)); 
        
        --临时记录的操作员信息记录的是终止人员,因为没有记录终止人员编号,就不恢复
        IF v_病区科室独立=1 THEN 
             Update 病人变动记录
                Set 开始时间=Null,护理等级ID=Null,
                    床位等级ID=Null,床号=Null,
                    责任护士=Null,经治医师=Null,
                    操作员编号=Null,操作员姓名=v_终止人员,
                    上次计算时间=Null,病区ID=Null
            Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
                AND 开始原因=3 And 终止时间 is Null;
        ELSE 
            Update 病人变动记录
                Set 开始时间=Null,护理等级ID=Null,
                    床位等级ID=Null,床号=Null,
                    责任护士=Null,经治医师=Null,
                    操作员编号=Null,操作员姓名=v_终止人员,
                    上次计算时间=Null
            Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
                AND 开始原因=3 And 终止时间 is Null;
        END if;

        Update 病人变动记录
            Set 终止时间=Null,终止原因=Null,终止人员=Null,
                上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 终止原因=3 And 终止时间=v_开始时间;
        
    Elsif r_CurLogRow.开始原因=4 Then
        --撤消换床
        v_开始时间:=r_CurLogRow.开始时间;
        Close c_CurLog;

        --退除当前床位
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.床号 IS NOT Null Then
                Update 床位状况记录 
                    Set 状态='空床',
                        病人ID=Null,
                        科室ID=Decode(共用,1,NULL,科室ID)
                Where 病区ID=r_CurLogRow.病区ID 
                    And 床号=r_CurLogRow.床号;
            End IF;
        End Loop;

        --检查及还原原床位
        For r_PreLogRow IN c_PreLog(v_开始时间,4) Loop
            if r_PreLogRow.床号 IS NOT Null Then
                Select Count(*) Into v_Count From 床位状况记录 Where 病区ID=r_PreLogRow.病区ID And 床号=r_PreLogRow.床号 And 状态='空床';
                if v_Count=0 Then
                    v_Error:='[ZLSOFT]病人最近一次换床前所入住的床位 '||r_PreLogRow.床号 ||' 当前非空床或已经撤消！[ZLSOFT]';
                    Raise Err_Custom;
                End IF;

                Update 床位状况记录 
                    Set 状态='占用',
                        病人ID=病人ID_IN,
                        科室ID=Decode(共用,1,r_PreLogRow.科室ID,科室ID)
                Where 病区ID=r_PreLogRow.病区ID 
                    And 床号=r_PreLogRow.床号;
            End IF;

            --病人信息、病案主页,仅当病区与科室独立时,换床才可以换病区,此处为简化判断,统一还原病区
            if NVL(r_PreLogRow.附加床位,0)=0 Then
                Update 病案主页 Set 出院病床=r_PreLogRow.床号,当前病区ID=r_PreLogRow.病区ID      
                Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;

                Update 病人信息 Set 当前床号=r_PreLogRow.床号,当前病区ID=r_PreLogRow.病区ID Where 病人ID=病人ID_IN;
            End IF;
        End Loop;

        --恢复变动
        Delete From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 开始原因=4 And 终止时间 is Null;

        Update 病人变动记录
            Set 终止时间=Null,终止原因=Null,终止人员=Null,上次计算时间=Null
         Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 终止原因=4 And 终止时间=v_开始时间;
    Elsif r_CurLogRow.开始原因=5 Then
        --撤消床位等级变动
        --还原原床位的等级
        For r_PreLogRow IN c_PreLog(r_CurLogRow.开始时间,5) Loop
            if r_PreLogRow.床号 IS NOT Null Then
                Update 床位状况记录 Set 等级ID=r_PreLogRow.床位等级ID Where 病区ID=r_PreLogRow.病区ID And 床号=r_PreLogRow.床号;
            End IF;
        End Loop;
        --恢复变动
        Delete From 病人变动记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 开始原因=5 And 终止时间 is Null;

        Update 病人变动记录 
            Set 终止时间=Null,终止原因=Null,终止人员=Null,上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 终止原因=5 And 终止时间=r_CurLogRow.开始时间;

        Close c_CurLog;
    Elsif r_CurLogRow.开始原因=6 Then
        --撤消护理等级变动
        Open c_PreLog(r_CurLogRow.开始时间,6);
        Fetch c_PreLog Into r_PreLogRow;
        --恢复原护理等级
        Update 病案主页
            Set 护理等级ID=r_PreLogRow.护理等级ID
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
        --恢复变动
        Delete From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 开始原因=6 And 终止时间 is Null;

        Update 病人变动记录
            Set 终止时间=Null,终止原因=Null,终止人员=Null,上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 终止原因=6 And 终止时间=r_CurLogRow.开始时间;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.开始原因=7 Then
        --撤消经治医师改变
        Open c_PreLog(r_CurLogRow.开始时间,7);
        Fetch c_PreLog Into r_PreLogRow;
        --恢复原医师
        Update 病案主页 Set 住院医师=r_PreLogRow.经治医师 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
        --恢复变动
        Delete From 病人变动记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 开始原因=7 And 终止时间 is Null;

        Update 病人变动记录 
            Set 终止时间=Null,终止原因=Null,终止人员=Null,上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN 
            And 终止原因=7 And 终止时间=r_CurLogRow.开始时间;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.开始原因=8 Then
        --撤消责任护士改变
        Open c_PreLog(r_CurLogRow.开始时间,8);
        Fetch c_PreLog Into r_PreLogRow;

        --恢复原责任护士
        Update 病案主页 Set 责任护士=r_PreLogRow.责任护士
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
        --恢复变动
        Delete From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 开始原因=8 And 终止时间 is Null;

        Update 病人变动记录 
            Set 终止时间=Null,终止原因=Null,终止人员=Null,上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 终止原因=8 And 终止时间=r_CurLogRow.开始时间;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.开始原因=9 Then
        --撤消转为住院病人

        --恢复原责任护士
        Update 病案主页 Set 病人性质=2 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
        --恢复变动
        Delete From 病人变动记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 开始原因=9 And 终止时间 is Null;

        Update 病人变动记录 
            Set 终止时间=Null,终止原因=Null,终止人员=Null,上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
            AND 终止原因=9 And 终止时间=r_CurLogRow.开始时间;
        
        If 数据_IN IS Not NULL Then 
            Update 病人信息 Set 住院号=NULL Where 病人ID=病人ID_IN;
        END if;

        Close c_CurLog;
    Elsif r_CurLogRow.开始原因=10 Then
        --撤消预出院

        --恢复住院状态
        Update 病案主页 Set 状态=0 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
        --恢复变动
        Delete From 病人变动记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 开始原因=10 And 终止时间 is Null;

        Update 病人变动记录 
            Set 终止时间=Null,终止原因=Null,上次计算时间=Null,终止人员=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN 
            And 终止原因=10 And 终止时间=r_CurLogRow.开始时间;

        Close c_CurLog;
    Elsif r_CurLogRow.开始原因=11 Then
        --撤消主治医师改变
        Open c_PreLog(r_CurLogRow.开始时间,11);
        Fetch c_PreLog Into r_PreLogRow;
        --恢复原主治医师
        Update 病案主页从表 Set 信息值=r_PreLogRow.主治医师 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主治医师';
        --恢复变动
        Delete From 病人变动记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 开始原因=11 And 终止时间 is Null;

        Update 病人变动记录 
            Set 终止时间=Null,终止原因=Null,终止人员=Null,上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN 
            And 终止原因=11 And 终止时间=r_CurLogRow.开始时间;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.开始原因=12 Then
        --撤消主任医师改变
        Open c_PreLog(r_CurLogRow.开始时间,12);
        Fetch c_PreLog Into r_PreLogRow;
        --恢复原主任医师
        Update 病案主页从表 Set 信息值=r_PreLogRow.主任医师 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主任医师';
        --恢复变动
        Delete From 病人变动记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 开始原因=12 And 终止时间 is Null;

        Update 病人变动记录 
            Set 终止时间=Null,终止原因=Null,终止人员=Null,上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN 
            And 终止原因=12 And 终止时间=r_CurLogRow.开始时间;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.开始原因=13 Then
        --撤消病情改变
        Open c_PreLog(r_CurLogRow.开始时间,13);
        Fetch c_PreLog Into r_PreLogRow;
        --恢复原病情
        Update 病案主页 Set 当前病况=r_PreLogRow.病情 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN ;
        --恢复变动
        Delete From 病人变动记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 开始原因=13 And 终止时间 is Null;

        Update 病人变动记录 
            Set 终止时间=Null,终止原因=Null,终止人员=Null,上次计算时间=Null
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN 
            And 终止原因=13 And 终止时间=r_CurLogRow.开始时间;

        Close c_PreLog;
        Close c_CurLog;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,v_Error);
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_病人变动记录_Undo;
/

Create Or Replace Procedure ZL_病人变动记录_PreOut(
    病人ID_IN		病案主页.病人ID%Type,
    主页ID_IN		病案主页.主页ID%Type,
	发生时间_IN		病人变动记录.开始时间%Type	
) AS
-----------------------------------------------------------
--功能：将病人标为预出院状态，并产生一条变动
-----------------------------------------------------------
    Cursor c_OldInfo IS
        Select * From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;
    r_OldInfo    c_OldInfo%RowType;
    
    v_Temp        Varchar2(255);
    v_人员编号    病人费用记录.操作员编号%Type;
    v_人员姓名    病人费用记录.操作员姓名%Type;

    v_Count        Number;
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --并发操作检查
    Select Nvl(状态,0) Into v_Count From 病案主页 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
    If v_Count<>0 Then
        v_Error:='该病人当前正在转科或尚未入科，不能执行预出院。';
        Raise Err_Custom;
    End IF;
    
    --操作员信息
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);
    
    Open c_OldInfo;--必须在处理之前先打开

    --取消上次变动
    Update 病人变动记录
        Set 终止时间=发生时间_IN,终止原因=10,终止人员=v_人员姓名
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;

    --产生新变动
    Fetch c_OldInfo Into r_OldInfo;
    If c_OldInfo%RowCount=0 Then
        Close c_OldInfo;
        v_Error:='未发现该病人当前有效的变动记录！';
        Raise Err_Custom;
    End IF;

    While c_OldInfo%Found Loop
        Insert Into 病人变动记录(
            病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,
            护理等级ID,床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
        Values(
            病人ID_IN,主页ID_IN,发生时间_IN,10,r_OldInfo.附加床位,r_OldInfo.病区ID,
            r_OldInfo.科室ID,r_OldInfo.护理等级ID,r_OldInfo.床位等级ID,r_OldInfo.床号,
            r_OldInfo.责任护士,r_OldInfo.经治医师,r_OldInfo.主治医师,r_OldInfo.主任医师,r_OldInfo.病情,v_人员编号,v_人员姓名);
        Fetch c_OldInfo Into r_OldInfo;
    End Loop;

    Close c_OldInfo;

    Update 病案主页 Set 状态=3 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;

    --并发操作检查
    Select Count(*) Into v_Count
    From 病人变动记录
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
        AND NVL(附加床位,0)=0 And 开始时间 IS NOT Null
        AND 终止时间 is Null;
    If v_Count > 1 Then
        v_Error:='发现病人存在非法的变动记录,当前操作不能继续！'||Chr(13)||Chr(10)||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人变动记录_PreOut;
/


Create Or Replace Procedure zl_入院病案主页_Insert(
	登记模式_IN			Number,
    病人性质_IN			病案主页.病人性质%Type,
    病人ID_IN           病人信息.病人ID%Type,
    住院号_IN           病人信息.住院号%Type,
	医保号_IN			保险帐户.医保号%Type,
    姓名_IN             病人信息.姓名%Type,
    性别_IN             病人信息.性别%Type,
    年龄_IN             病人信息.年龄%Type,
    费别_IN             病人信息.费别%Type,
    出生日期_IN         病人信息.出生日期%Type,
    国籍_IN             病人信息.国籍%Type,
    民族_IN             病人信息.民族%Type,
    学历_IN             病人信息.学历%Type,
    婚姻状况_IN         病人信息.婚姻状况%Type,
    职业_IN             病人信息.职业%Type,
    身份_IN             病人信息.身份%Type,
    身份证号_IN         病人信息.身份证号%Type,
    出生地点_IN         病人信息.出生地点%Type,
    家庭地址_IN         病人信息.家庭地址%Type,
    户口邮编_IN         病人信息.户口邮编%Type,
    家庭电话_IN         病人信息.家庭电话%Type,
    联系人姓名_IN       病人信息.联系人姓名%Type,
    联系人关系_IN       病人信息.联系人关系%Type,
    联系人地址_IN       病人信息.联系人地址%Type,
    联系人电话_IN       病人信息.联系人电话%Type,
    工作单位_IN         病人信息.工作单位%Type,
    合同单位ID_IN       病人信息.合同单位ID%Type,
    单位电话_IN         病人信息.单位电话%Type,
    单位邮编_IN         病人信息.单位邮编%Type,
    单位开户行_IN       病人信息.单位开户行%Type,
    单位帐号_IN         病人信息.单位帐号%Type,
    担保人_IN           病人信息.担保人%Type,
    担保额_IN           病人信息.担保额%Type,
	担保性质_IN			病人信息.担保性质%Type,
    入院科室ID_IN       病案主页.入院科室ID%Type,
    护理等级ID_IN       病案主页.护理等级ID%Type,
    入院病况_IN         病案主页.入院病况%Type,
    入院方式_IN         病案主页.入院方式%Type,
    住院目的_IN         病案主页.住院目的%Type,
    二级院转入_IN       病案主页.二级院转入%Type,
    门诊医师_IN         病案主页.门诊医师%Type,
    区域_IN             病案主页.区域%Type,
    入院时间_IN         病案主页.入院日期%Type,
    是否陪伴_IN         病案主页.是否陪伴%Type,
    床号_IN             病案主页.入院病床%Type,
    付款方式_IN         病案主页.医疗付款方式%Type,
    疾病ID_IN           病人诊断记录.疾病ID%Type,
    门诊诊断_IN         病人诊断记录.诊断描述%Type,
    中医疾病ID_IN       病人诊断记录.疾病ID%Type,
    中医诊断_IN			病人诊断记录.诊断描述%Type,
    险类_IN             病案主页.险类%Type,
    操作员编号_IN       病案主页.编目员编号%Type,
    操作员姓名_IN       病案主页.编目员姓名%Type,
    新病人_IN           Number:=1,
	备注_IN				病案主页.备注%Type:=Null,
    入院病区ID_IN       病案主页.入院病区ID%Type:=Null
) AS
-----------------------------------------------------------
--功能：对入院病人新增一张病案主页，同时可能处理入科。
--参数：
--      登记模式_IN=0-正常登记,1-预约登记,2-接收预约(新病人_IN=0)
--      病人性质_IN=对应"病案主页.病人性质"
--      床号_IN=Null:不同时入科;0:分配家庭病床,填为空;数字:分配具体床位。
--      新病人_IN=如果是已有档案的病人入院,则该参数为0；缺省为新病人
--      入院病区ID_IN=只有当使用[病区管理病床]模式(参数号99)时,并且入院同时入科分床时,才有值
-----------------------------------------------------------
    v_主页ID	                病案主页.主页ID%Type;
    v_病区ID	                病案主页.入院病区ID%Type;
    v_等级ID                  床位状况记录.等级ID%Type;
   	v_病区科室独立     Number(1);

    v_Count     Number;
    v_Date      Date;
    v_Error     Varchar2(255);
    Err_Custom  Exception;
Begin
    Select Sysdate Into v_Date From Dual;

    --病人基本信息
    IF 新病人_IN=1 Then
        Insert Into 病人信息(
            病人ID,住院号,姓名,性别,年龄,费别,医疗付款方式,出生日期,国籍,民族,区域,学历,
            婚姻状况,职业,身份,身份证号,出生地点,家庭地址,户口邮编,家庭电话,联系人姓名,
            联系人关系,联系人地址,联系人电话,工作单位,合同单位ID,单位电话,单位邮编,
            单位开户行,单位帐号,担保人,担保额,担保性质,险类,登记时间)
        Values(
            病人ID_IN,住院号_IN,姓名_IN,性别_IN,年龄_IN,费别_IN,付款方式_IN,出生日期_IN,
            国籍_IN,民族_IN,区域_IN,学历_IN,婚姻状况_IN,职业_IN,身份_IN,身份证号_IN,出生地点_IN,
            家庭地址_IN,户口邮编_IN,家庭电话_IN,联系人姓名_IN,联系人关系_IN,联系人地址_IN,
            联系人电话_IN,工作单位_IN,Decode(合同单位ID_IN,0,Null,合同单位ID_IN),单位电话_IN,
            单位邮编_IN,单位开户行_IN,单位帐号_IN,担保人_IN,Decode(担保额_IN,0,Null,担保额_IN),
            担保性质_IN,险类_IN,入院时间_IN);
    Else
        --老病人的门诊费别不变,除非是门诊留观病人
        Update 病人信息
            Set 住院号=住院号_IN,姓名=姓名_IN,
                性别=性别_IN,年龄=年龄_IN,
				费别=Decode(病人性质_IN,1,费别_IN,费别),
                医疗付款方式=付款方式_IN,
                出生日期=出生日期_IN,国籍=国籍_IN,
                民族=民族_IN,区域=区域_IN,学历=学历_IN,
                婚姻状况=婚姻状况_IN,职业=职业_IN,
                身份=身份_IN,身份证号=身份证号_IN,
                出生地点=出生地点_IN,家庭地址=家庭地址_IN,
                户口邮编=户口邮编_IN,家庭电话=家庭电话_IN,
                联系人姓名=联系人姓名_IN,联系人关系=联系人关系_IN,
                联系人地址=联系人地址_IN,联系人电话=联系人电话_IN,
                工作单位=工作单位_IN,合同单位ID=Decode(合同单位ID_IN,0,Null,合同单位ID_IN),
                单位电话=单位电话_IN,单位邮编=单位邮编_IN,
                单位开户行=单位开户行_IN,单位帐号=单位帐号_IN,
                担保人=担保人_IN,担保额=Decode(担保额_IN,0,Null,担保额_IN),
                担保性质=担保性质_IN,险类=险类_IN
        Where 病人ID=病人ID_IN;
    End if;

    --住院病案记录:预约时不产生,接收时才产生,但可能先预约住院号
	If 登记模式_IN<>1 Then
		If 住院号_IN IS Not NULL Then
			Update 住院病案记录
				Set 病案号=住院号_IN,病案类别='一般',存储状态='在院'
			Where 病人ID=病人ID_IN;
			If SQL%RowCount=0 Then
				Insert Into 住院病案记录(
					病人ID,病案号,病案类别,存储状态,建立日期)
				Values(
					病人ID_IN,住院号_IN,'一般','在院',入院时间_IN);
			End IF;
		Else
			Delete From 住院病案记录 Where 病人ID=病人ID_IN;
		End IF;
	End IF;

    --病案信息
    Begin
        If 登记模式_IN=1 Then
            v_主页ID:=0;--预约登记记录的主页ID=0
        Else
            Select Nvl(Max(主页ID),0)+1 Into v_主页ID From 病案主页 Where 病人ID=病人ID_IN And Nvl(主页ID,0)<>0;
        End IF;
        
         --病区科室独立模式,病人入院同时入科配床,需要在界面选择病区,如果没有分床,则传入空
        v_病区科室独立:=To_Number(Nvl(ZL_GetSysParameter(99),0)); 
        IF v_病区科室独立=1 Then 
            v_病区ID :=入院病区ID_IN;
        Else        
           Select DISTINCT 病区ID Into v_病区ID From 床位状况记录 Where 科室ID=入院科室ID_IN;        
        End If;
    Exception
        When OTHERS Then Null;
    End;
	
	If 登记模式_IN<>1 Then
		Update 病人信息
			Set 住院次数=v_主页ID,当前病区ID=v_病区ID,
				当前科室ID=入院科室ID_IN,当前床号=Decode(床号_IN,Null,Null,0,Null,床号_IN),
				入院时间=入院时间_IN,出院时间=Null
		Where 病人ID=病人ID_IN;
	End IF;

    --状态：0-正常在院,1-等待入科,2-等待转科
	IF 登记模式_IN=2 Then
		--接收预约
		Update 病案主页
			Set 主页ID=v_主页ID,病人性质=病人性质_IN,--主页ID变更,病人性质可能变更
				费别=费别_IN,入院病区ID=v_病区ID,
				入院科室ID=入院科室ID_IN,入院日期=入院时间_IN,
				入院病况=入院病况_IN,入院方式=入院方式_IN,
				二级院转入=二级院转入_IN,住院目的=住院目的_IN,
				入院病床=Decode(床号_IN,Null,Null,0,Null,床号_IN),
				是否陪伴=是否陪伴_IN,
				当前病况=入院病况_IN,当前病区ID=v_病区ID,
				护理等级ID=Decode(护理等级ID_IN,0,Null,护理等级ID_IN),
				出院科室ID=入院科室ID_IN,
				出院病床=Decode(床号_IN,Null,Null,0,Null,床号_IN),
				门诊医师=门诊医师_IN,
				编目员编号=操作员编号_IN,编目员姓名=操作员姓名_IN,
				年龄=年龄_IN,婚姻状况=婚姻状况_IN,
				职业=职业_IN,国籍=国籍_IN,
				学历=学历_IN,单位电话=单位电话_IN,
				单位邮编=单位邮编_IN,单位地址=工作单位_IN,
				区域=区域_IN,家庭地址=家庭地址_IN,
				家庭电话=家庭电话_IN,户口邮编=户口邮编_IN,
				联系人姓名=联系人姓名_IN,联系人关系=联系人关系_IN,
				联系人地址=联系人地址_IN,联系人电话=联系人电话_IN,
				医疗付款方式=付款方式_IN,备注=备注_IN,
				险类=险类_IN,状态=Decode(床号_IN,Null,1,0),
				登记人=操作员姓名_IN,登记时间=v_Date
		Where 病人ID=病人ID_IN And Nvl(主页ID,0)=0;
	Else
		--入院登记或预约登记
		Insert Into 病案主页(
			病人性质,病人ID,主页ID,费别,入院病区ID,入院科室ID,入院日期,入院病况,入院方式,二级院转入,住院目的,
			入院病床,是否陪伴,当前病况,当前病区ID,护理等级ID,出院科室ID,出院病床,门诊医师,编目员编号,
			编目员姓名,状态,年龄,婚姻状况,职业,国籍,学历,单位电话,单位邮编,单位地址,区域,家庭地址,
			家庭电话,户口邮编,联系人姓名,联系人关系,联系人地址,联系人电话,医疗付款方式,险类,备注,登记人,登记时间)
		Values(
			病人性质_IN,病人ID_IN,v_主页ID,费别_IN,v_病区ID,入院科室ID_IN,入院时间_IN,入院病况_IN,入院方式_IN,
			二级院转入_IN,住院目的_IN,Decode(床号_IN,Null,Null,0,Null,床号_IN),是否陪伴_IN,
			入院病况_IN,v_病区ID,Decode(护理等级ID_IN,0,Null,护理等级ID_IN),入院科室ID_IN,
			Decode(床号_IN,Null,Null,0,Null,床号_IN),门诊医师_IN,操作员编号_IN,操作员姓名_IN,
			Decode(床号_IN,Null,1,0),年龄_IN,婚姻状况_IN,职业_IN,国籍_IN,学历_IN,单位电话_IN,
			单位邮编_IN,工作单位_IN,区域_IN,家庭地址_IN,家庭电话_IN,户口邮编_IN,联系人姓名_IN,
			联系人关系_IN,联系人地址_IN,联系人电话_IN,付款方式_IN,险类_IN,备注_IN,操作员姓名_IN,v_Date);
	End If;
	
	--医保号
	If 登记模式_IN<>1 Then
		If 医保号_IN IS Not Null Then
			Insert Into 病案主页从表(
				病人ID,主页ID,信息名,信息值)
			Values(
				病人ID_IN,v_主页ID,'医保号',医保号_IN);
		End IF;

		--病人变动记录
		--同时入科且非家庭病床时有等级
		if Nvl(床号_IN,0) <> 0 Then
			Select 等级ID Into v_等级ID From 床位状况记录 Where 病区ID=v_病区ID And 床号=床号_IN;
		End IF;

		--如果同时入科,则入院和入科填写到一条入院变动
		Insert Into 病人变动记录(
			病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,护理等级ID,床位等级ID,床号,病情,操作员编号,操作员姓名)
		Values(
			病人ID_IN,v_主页ID,入院时间_IN,1,0,v_病区ID,入院科室ID_IN,Decode(护理等级ID_IN,0,Null,护理等级ID_IN),
			v_等级ID,Decode(床号_IN,0,Null,床号_IN),入院病况_IN,操作员编号_IN,操作员姓名_IN);

		--同时入科且非家庭病床时床位被占用
		If Nvl(床号_IN,0) <> 0 Then
			Select Count(*) Into v_Count From 床位状况记录 Where 病区ID=v_病区ID And 床号=床号_IN And 状态='空床';

			if v_Count=0 Then
				v_Error:='操作失败,床位 '||床号_IN||' 不是空床！';
				Raise Err_Custom;
			End IF;

			Update 床位状况记录 Set 状态='占用',病人ID=病人ID_IN,科室ID=Decode(共用,1,入院科室ID_IN,科室ID) Where 病区ID=v_病区ID And 床号=床号_IN;
		End IF;

		--病人诊断记录
		If 门诊诊断_IN IS Not Null Or 疾病ID_IN IS Not NULL Then
			Insert Into 病人诊断记录(
				ID,病人ID,主页ID,记录来源,诊断类型,诊断次序,疾病ID,诊断描述,记录日期,记录人) 
			Values(
				病人诊断记录_ID.Nextval,病人ID_IN,v_主页ID,2,1,1,疾病ID_IN,门诊诊断_IN,sysdate,操作员姓名_IN);
		End IF;        
		If 中医诊断_IN IS Not Null Or 中医疾病ID_IN IS Not NULL Then
			Insert Into 病人诊断记录(
				ID,病人ID,主页ID,记录来源,诊断类型,诊断次序,疾病ID,诊断描述,记录日期,记录人) 
			Values(
				病人诊断记录_ID.Nextval,病人ID_IN,v_主页ID,2,11,1,中医疾病ID_IN,中医诊断_IN,sysdate,操作员姓名_IN);
		End IF;        

		--并发操作检查
		Select Count(*) Into v_Count From 病案主页 Where 病人ID=病人ID_IN And 出院日期 is Null;
		If v_Count>1 Then
			v_Error:='发现病人存在非法的病案记录,当前操作不能继续！'||Chr(13)||Chr(10) ||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
			Raise Err_Custom;
		End IF;

		Select Count(*) Into v_Count
			From 病人变动记录
			Where 病人ID=病人ID_IN And 主页ID=v_主页ID And Nvl(附加床位,0)=0
				And 开始时间 IS Not Null And 终止时间 is Null;
		If v_Count>1 Then
			v_Error:='发现病人存在非法的变动记录,当前操作不能继续！'||Chr(13)||Chr(10) ||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
			Raise Err_Custom;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_入院病案主页_Insert;
/


Create Or Replace Procedure zl_入院病案主页_Update(
	登记模式_IN		Number,
	病人ID_IN		病人信息.病人ID%Type,
    住院号_IN       病人信息.住院号%Type,
	医保号_IN		保险帐户.医保号%Type,
    姓名_IN         病人信息.姓名%Type,
    性别_IN         病人信息.性别%Type,
    年龄_IN         病人信息.年龄%Type,
    费别_IN         病人信息.费别%Type,
    出生日期_IN     病人信息.出生日期%Type,
    国籍_IN         病人信息.国籍%Type,
    民族_IN         病人信息.民族%Type,
    学历_IN         病人信息.学历%Type,
    婚姻状况_IN     病人信息.婚姻状况%Type,
    职业_IN         病人信息.职业%Type,
    身份_IN         病人信息.身份%Type,
    身份证号_IN     病人信息.身份证号%Type,
    出生地点_IN     病人信息.出生地点%Type,
    家庭地址_IN     病人信息.家庭地址%Type,
    户口邮编_IN     病人信息.户口邮编%Type,
    家庭电话_IN     病人信息.家庭电话%Type,
    联系人姓名_IN   病人信息.联系人姓名%Type,
    联系人关系_IN   病人信息.联系人关系%Type,
    联系人地址_IN   病人信息.联系人地址%Type,
    联系人电话_IN   病人信息.联系人电话%Type,
    工作单位_IN     病人信息.工作单位%Type,
    合同单位ID_IN   病人信息.合同单位ID%Type,
    单位电话_IN     病人信息.单位电话%Type,
    单位邮编_IN     病人信息.单位邮编%Type,
    单位开户行_IN   病人信息.单位开户行%Type,
    单位帐号_IN     病人信息.单位帐号%Type,
    担保人_IN       病人信息.担保人%Type,
    担保额_IN       病人信息.担保额%Type,
	担保性质_IN		病人信息.担保性质%Type,
    主页ID_IN       病案主页.主页ID%Type,
    入院科室ID_IN   病案主页.入院科室ID%Type,
    护理等级ID_IN   病案主页.护理等级ID%Type,
    入院病况_IN     病案主页.入院病况%Type,
    入院方式_IN     病案主页.入院方式%Type,
    住院目的_IN     病案主页.住院目的%Type,
    二级院转入_IN   病案主页.二级院转入%Type,
    门诊医师_IN     病案主页.门诊医师%Type,
    区域_IN         病案主页.区域%Type,
    入院时间_IN     病案主页.入院日期%Type,
    付款方式_IN     病案主页.医疗付款方式%Type,
    疾病ID_IN       病人诊断记录.疾病ID%Type,
    门诊诊断_IN     病人诊断记录.诊断描述%Type,
    中医疾病ID_IN   病人诊断记录.疾病ID%Type,
    中医诊断_IN     病人诊断记录.诊断描述%Type,
    操作员编号_IN   病案主页.编目员编号%Type,
    操作员姓名_IN   病案主页.编目员姓名%Type,
	备注_IN			病案主页.备注%Type:=Null,
    病区ID_IN       病案主页.入院病区Id%Type:=Null
) AS
-----------------------------------------------------------
--说明：本函数仅用于入院未入科登记病人信息的修改
--      登记模式_IN=0-正常登记,1-预约登记
--      病区ID_IN=只有当病区管理病床模式下,入院时入科时,才会有值
-----------------------------------------------------------    
	v_病区ID				    病案主页.入院病区ID%Type;
	v_等级ID				    床位状况记录.等级ID%Type;
	v_病人性质			    病案主页.病人性质%Type;
	v_病区科室独立     Number(1);
    
    v_Count			Number;
    v_Date			Date;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --判断病人是否未入院
    Select Count(*) Into v_Count From 病案主页 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 状态=1;
    if v_Count=0 Then
        v_Error:='病人当前不处于等待入科状态，操作不能继续！'||Chr(13)||Chr(10)||'可能该病人已经被其它操作员取消登记或分配床位。';
        Raise Err_Custom;
    End IF;
	
    Select Sysdate Into v_Date From Dual;
	Select 病人性质 Into v_病人性质 From 病案主页 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
	
    --病区科室独立模式,病人入院同时入科配床,需要在界面选择病区,如果没有分床,则传入空
    v_病区科室独立:=To_Number(Nvl(ZL_GetSysParameter(99),0));    
    If v_病区科室独立=1 Then
        v_病区ID:=病区ID_IN;
    Else        
        Begin
	        Select Distinct 病区ID Into v_病区ID From 床位状况记录 Where 科室ID=入院科室ID_IN;
        Exception
             When OTHERS Then Null;
        End;
    End If;

    --病人基本信息
    --非第一次入院时,门诊费别保持不变,除非是门诊留观病人
	Update 病人信息
        Set 住院号=住院号_IN,姓名=姓名_IN,性别=性别_IN,
            年龄=年龄_IN,医疗付款方式=付款方式_IN,
            费别=Decode(v_病人性质,1,费别_IN,费别),
            出生日期=出生日期_IN,国籍=国籍_IN,民族=民族_IN,
            区域=区域_IN,学历=学历_IN,婚姻状况=婚姻状况_IN,职业=职业_IN,
            身份=身份_IN,身份证号=身份证号_IN,出生地点=出生地点_IN,
            家庭地址=家庭地址_IN,户口邮编=户口邮编_IN,家庭电话=家庭电话_IN,
            联系人姓名=联系人姓名_IN,联系人关系=联系人关系_IN,
            联系人地址=联系人地址_IN,联系人电话=联系人电话_IN,
            工作单位=工作单位_IN,合同单位ID=Decode(合同单位ID_IN,0,Null,合同单位ID_IN),
            单位电话=单位电话_IN,单位邮编=单位邮编_IN,单位开户行=单位开户行_IN,
            单位帐号=单位帐号_IN,担保人=担保人_IN,
			担保额=Decode(担保额_IN,0,Null,担保额_IN),担保性质=担保性质_IN
	Where 病人ID=病人ID_IN;

	If 登记模式_IN=0 Then
		--住院病案记录
		If 住院号_IN IS Not NULL Then
			Update 住院病案记录 Set 病案号=住院号_IN Where 病人ID=病人ID_IN;
		Else
			Delete From 住院病案记录 Where 病人ID=病人ID_IN;
		End IF;

		--病案信息
		Update 病人信息
			Set 当前病区ID=v_病区ID,当前科室ID=入院科室ID_IN,
				入院时间=入院时间_IN,出院时间=Null
		Where 病人ID=病人ID_IN;
	End If;

    --修改病案主页
    Update 病案主页
        Set 费别=费别_IN,入院病区ID=v_病区ID,
            入院科室ID=入院科室ID_IN,入院日期=入院时间_IN,
            入院病况=入院病况_IN,入院方式=入院方式_IN,
            二级院转入=二级院转入_IN,住院目的=住院目的_IN,
            当前病况=入院病况_IN,当前病区ID=v_病区ID,
            护理等级ID=Decode(护理等级ID_IN,0,Null,护理等级ID_IN),
            出院科室ID=入院科室ID_IN,门诊医师=门诊医师_IN,
            编目员编号=操作员编号_IN,编目员姓名=操作员姓名_IN,
            年龄=年龄_IN,婚姻状况=婚姻状况_IN,
            职业=职业_IN,国籍=国籍_IN,
            学历=学历_IN,单位电话=单位电话_IN,
            单位邮编=单位邮编_IN,单位地址=工作单位_IN,
            区域=区域_IN,家庭地址=家庭地址_IN,
            家庭电话=家庭电话_IN,户口邮编=户口邮编_IN,
            联系人姓名=联系人姓名_IN,联系人关系=联系人关系_IN,
            联系人地址=联系人地址_IN,联系人电话=联系人电话_IN,
            医疗付款方式=付款方式_IN,备注=备注_IN,
            登记人=操作员姓名_IN,登记时间=v_Date
    Where 病人ID=病人ID_IN And Nvl(主页ID,0)=Nvl(主页ID_IN,0);
	
	If 登记模式_IN=0 Then
		--医保号
		If 医保号_IN IS Not NULL Then
			Update 病案主页从表 Set 信息值=医保号_IN Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='医保号';
			If SQL%RowCount=0 Then
				Insert Into 病案主页从表(
					病人ID,主页ID,信息名,信息值)
				Values(
					病人ID_IN,主页ID_IN,'医保号',医保号_IN);
			End IF;
		Else
			Delete From 病案主页从表 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='医保号';
		End IF;

		--修改病人变动记录(肯定为入院变动;单独入科的不准修改,入院同时入科的病人界面程序禁止修改)
		Update 病人变动记录
			Set 开始时间=入院时间_IN,病区ID=v_病区ID,科室ID=入院科室ID_IN,
				护理等级ID=Decode(护理等级ID_IN,0,Null,护理等级ID_IN),病情=入院病况_IN,
				操作员编号=操作员编号_IN,操作员姓名=操作员姓名_IN,上次计算时间=Null
		Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
			And 终止时间 is Null And 开始原因=1;
		
		--处理门诊诊断
		If 门诊诊断_IN is Null AND 疾病ID_IN IS NULL Then
			Delete From 病人诊断记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 诊断类型=1 And 记录来源=2;
		Else
			Update 病人诊断记录 Set 疾病ID=疾病ID_IN,诊断描述=门诊诊断_IN,记录日期=sysdate,记录人=操作员姓名_IN
            Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 诊断类型=1  And 记录来源=2;
			IF SQL%RowCount=0 Then
				Insert Into 病人诊断记录(
					ID,病人ID,主页ID,记录来源,诊断类型,诊断次序,疾病ID,诊断描述,记录日期,记录人) 
				Values(
					病人诊断记录_ID.Nextval ,病人ID_IN,主页ID_IN,2,1,1,疾病ID_IN,门诊诊断_IN,sysdate,操作员姓名_IN);
			End IF;
		End IF;

		--处理中医诊断
		If 中医诊断_IN is Null AND 中医疾病ID_IN IS NULL Then
			Delete From 病人诊断记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 诊断类型=11  And 记录来源=2;
		Else
			Update 病人诊断记录 Set 疾病ID=中医疾病ID_IN,诊断描述=中医诊断_IN ,记录日期=sysdate,记录人=操作员姓名_IN
            Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 诊断类型=11  And 记录来源=2;
			IF SQL%RowCount=0 Then
				Insert Into 病人诊断记录(
					ID,病人ID,主页ID,记录来源,诊断类型,诊断次序,疾病ID,诊断描述,记录日期,记录人) 
				Values(
					病人诊断记录_ID.Nextval,病人ID_IN,主页ID_IN,2,11,1,中医疾病ID_IN,中医诊断_IN,sysdate,操作员姓名_IN);
			End IF;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_入院病案主页_UpDate;
/

Create Or Replace Procedure zl_住院病案主页_Update(
    病人ID_IN		病案主页.病人ID%Type,
    主页ID_IN       病案主页.主页ID%Type,
    年龄_IN         病案主页.年龄%Type,
    费别_IN         病案主页.费别%Type,
    婚姻状况_IN     病案主页.婚姻状况%Type,
    学历_IN         病案主页.学历%Type,
    职业_IN         病案主页.职业%Type,
    当前病况_IN     病案主页.当前病况%Type,
    单位地址_IN     病案主页.单位地址%Type,
    合同单位ID_IN   病人信息.合同单位ID%Type,
    单位电话_IN     病案主页.单位电话%Type,
    单位邮编_IN     病案主页.单位邮编%Type,
    家庭地址_IN     病案主页.家庭地址%Type,
    家庭电话_IN     病案主页.家庭电话%Type,
    户口邮编_IN     病案主页.户口邮编%Type,
    联系人姓名_IN   病案主页.联系人姓名%Type,
    联系人关系_IN   病案主页.联系人关系%Type,
    联系人电话_IN   病案主页.联系人电话%Type,
    联系人地址_IN   病案主页.联系人地址%Type,
    责任护士_IN     病案主页.责任护士%Type,
    门诊医师_IN     病案主页.门诊医师%Type,
    住院医师_IN     病案主页.住院医师%Type,
    疾病ID_IN       病人诊断记录.疾病ID%Type,
    入院诊断_IN     病人诊断记录.诊断描述%Type,
    中医疾病ID_IN   病人诊断记录.疾病ID%Type,
    中医诊断_IN     病人诊断记录.诊断描述%Type,
    操作员编号_IN   病案主页.编目员编号%Type,
    操作员姓名_IN   病案主页.编目员姓名%Type,
    主治医师_IN       病案主页.住院医师%Type,
    主任医师_IN       病案主页.住院医师%Type
)
AS
    Cursor c_OldInfo IS
        Select * From 病人变动记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;
    r_OldInfo c_OldInfo%RowType;

    v_责任护士		病案主页.责任护士%Type;
    v_住院医师		病案主页.住院医师%Type;
    v_主治医师		病案主页.住院医师%Type;
    v_主任医师		病案主页.住院医师%Type;
	v_病人性质		病案主页.病人性质%Type;
    v_当前病况      病案主页.当前病况%Type;
    v_原因			病人变动记录.开始原因%Type;
    v_Count         Number;
    v_CurDate		Date;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --取更改前的内容(用NoneData和新的比较)
    Select 
		病人性质,Nvl(责任护士,'NoneData'),Nvl(住院医师,'NoneData'),Nvl(当前病况,'NoneData') Into v_病人性质,v_责任护士,v_住院医师,v_当前病况
    From 病案主页 
	Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
    Begin 
        Select Nvl(信息值,'NoneData') Into v_主治医师 From 病案主页从表 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主治医师';
    Exception
        When Others Then v_主治医师:='NoneData';
    End;
    Begin
        Select Nvl(信息值,'NoneData') Into v_主任医师 From 病案主页从表 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主任医师';
    Exception
        When Others Then v_主任医师:='NoneData';
    End;

    Update 病案主页
    Set 年龄=年龄_IN,费别=费别_IN,
        婚姻状况=婚姻状况_IN,学历=学历_IN,
        当前病况=当前病况_IN,职业=职业_IN,
        单位地址=单位地址_IN,单位电话=单位电话_IN,
        单位邮编=单位邮编_IN,家庭地址=家庭地址_IN,
        家庭电话=家庭电话_IN,户口邮编=户口邮编_IN,
        联系人姓名=联系人姓名_IN,联系人关系=联系人关系_IN,
        联系人地址=联系人地址_IN,联系人电话=联系人电话_IN,
        责任护士=责任护士_IN,门诊医师=门诊医师_IN,
        住院医师=住院医师_IN,编目员编号=操作员编号_IN,
        编目员姓名=操作员姓名_IN
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN;
    
    --仅修改住院费别,除非是门诊留观病人
    Update 病人信息
    Set 年龄=年龄_IN,费别=Decode(v_病人性质,1,费别_IN,费别),
        婚姻状况=婚姻状况_IN,学历=学历_IN,
        职业=职业_IN,工作单位=单位地址_IN,
        合同单位ID=Decode(合同单位ID_IN,0,合同单位ID,合同单位ID_IN),
        单位电话=单位电话_IN,单位邮编=单位邮编_IN,
        家庭地址=家庭地址_IN,家庭电话=家庭电话_IN,
        户口邮编=户口邮编_IN,联系人姓名=联系人姓名_IN,
        联系人关系=联系人关系_IN,联系人地址=联系人地址_IN,
        联系人电话=联系人电话_IN
    Where 病人ID=病人ID_IN;

    --产生变动记录
    if v_住院医师 <> 住院医师_IN Or v_责任护士 <> 责任护士_IN Or v_主治医师 <> 主治医师_IN Or v_主任医师 <> 主任医师_IN Or v_当前病况<>当前病况_IN Then
        Select Sysdate Into v_CurDate From Dual;
        Open c_OldInfo;
        Fetch c_OldInfo Into r_OldInfo;
        if c_OldInfo%RowCount=0 Then
            Close c_OldInfo;
            v_Error:='未发现该病人当前有效的变动记录！';
            Raise Err_Custom;
        End IF;

        if v_住院医师 <> 住院医师_IN Then
            v_原因:=7;    
            Update 病人变动记录 
                Set 终止时间=v_CurDate,终止原因=v_原因,
                    操作员编号=操作员编号_IN,操作员姓名=操作员姓名_IN
            Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;    
            While c_OldInfo%Found Loop                  --注意:有附加床位时有多条记录                                                                    
                Insert Into 病人变动记录(
                    病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,护理等级ID,
                    床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
                Values(
                    病人ID_IN,主页ID_IN,v_CurDate,v_原因,r_OldInfo.附加床位,r_OldInfo.病区ID,
                    r_OldInfo.科室ID,r_OldInfo.护理等级ID,r_OldInfo.床位等级ID,r_OldInfo.床号,
                    r_OldInfo.责任护士,住院医师_IN,r_OldInfo.主治医师,r_OldInfo.主任医师,r_OldInfo.病情,操作员编号_IN,操作员姓名_IN); 
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo;    --重新打开,以便取最新信息
            Fetch c_OldInfo Into r_OldInfo; 
        End IF;
    
       if v_责任护士 <> 责任护士_IN Then
            v_原因:=8;    
            Update 病人变动记录 
                Set 终止时间=v_CurDate,终止原因=v_原因,
                    操作员编号=操作员编号_IN,操作员姓名=操作员姓名_IN
            Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;             
            While c_OldInfo%Found Loop            
                Insert Into 病人变动记录(
                    病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,护理等级ID,
                    床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
                Values(
                    病人ID_IN,主页ID_IN,v_CurDate,v_原因,r_OldInfo.附加床位,r_OldInfo.病区ID,
                    r_OldInfo.科室ID,r_OldInfo.护理等级ID,r_OldInfo.床位等级ID,r_OldInfo.床号,
                    责任护士_IN,r_OldInfo.经治医师,r_OldInfo.主治医师,r_OldInfo.主任医师,r_OldInfo.病情,操作员编号_IN,操作员姓名_IN);                
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo;
            Fetch c_OldInfo Into r_OldInfo; 
        End IF;

        if v_主治医师 <> 主治医师_IN Then
            Update 病案主页从表 Set 信息值=主治医师_IN Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主治医师';
            IF SQL%RowCount=0 Then
                Insert Into 病案主页从表(病人ID,主页ID,信息名,信息值) Values (病人ID_IN,主页ID_IN,'主治医师',主治医师_IN);
            End IF;

            v_原因:=11;    
            Update 病人变动记录 
                Set 终止时间=v_CurDate,终止原因=v_原因,
                    操作员编号=操作员编号_IN,操作员姓名=操作员姓名_IN
            Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;    
            While c_OldInfo%Found Loop                                                                     
                Insert Into 病人变动记录(
                    病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,护理等级ID,
                    床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
                Values(
                    病人ID_IN,主页ID_IN,v_CurDate,v_原因,r_OldInfo.附加床位,r_OldInfo.病区ID,
                    r_OldInfo.科室ID,r_OldInfo.护理等级ID,r_OldInfo.床位等级ID,r_OldInfo.床号,
                    r_OldInfo.责任护士,r_OldInfo.经治医师,主治医师_IN,r_OldInfo.主任医师,r_OldInfo.病情,操作员编号_IN,操作员姓名_IN);                
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo; 
            Fetch c_OldInfo Into r_OldInfo; 
        End IF;

        if v_主任医师 <> 主任医师_IN Then
            Update 病案主页从表 Set 信息值=主任医师_IN Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 信息名='主任医师';
            IF SQL%RowCount=0 Then
                Insert Into 病案主页从表(病人ID,主页ID,信息名,信息值) Values (病人ID_IN,主页ID_IN,'主任医师',主任医师_IN);
            End IF;  

            v_原因:=12;    
            Update 病人变动记录 
                Set 终止时间=v_CurDate,终止原因=v_原因,
                    操作员编号=操作员编号_IN,操作员姓名=操作员姓名_IN
            Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;    
            While c_OldInfo%Found Loop
                Insert Into 病人变动记录(
                    病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,护理等级ID,
                    床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
                Values(
                    病人ID_IN,主页ID_IN,v_CurDate,v_原因,r_OldInfo.附加床位,r_OldInfo.病区ID,
                    r_OldInfo.科室ID,r_OldInfo.护理等级ID,r_OldInfo.床位等级ID,r_OldInfo.床号,
                    r_OldInfo.责任护士,r_OldInfo.经治医师,r_OldInfo.主治医师,主任医师_IN,r_OldInfo.病情,操作员编号_IN,操作员姓名_IN);                
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo;
            Fetch c_OldInfo Into r_OldInfo; 
        End IF;

        if v_当前病况<>当前病况_IN Then
            v_原因:=13;    
            Update 病人变动记录 
                Set 终止时间=v_CurDate,终止原因=v_原因,
                    操作员编号=操作员编号_IN,操作员姓名=操作员姓名_IN
            Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 终止时间 is Null;    
            While c_OldInfo%Found Loop                                                          
                Insert Into 病人变动记录(
                    病人ID,主页ID,开始时间,开始原因,附加床位,病区ID,科室ID,护理等级ID,
                    床位等级ID,床号,责任护士,经治医师,主治医师,主任医师,病情,操作员编号,操作员姓名)
                Values(
                    病人ID_IN,主页ID_IN,v_CurDate,v_原因,r_OldInfo.附加床位,r_OldInfo.病区ID,
                    r_OldInfo.科室ID,r_OldInfo.护理等级ID,r_OldInfo.床位等级ID,r_OldInfo.床号,
                    r_OldInfo.责任护士,r_OldInfo.经治医师,r_OldInfo.主治医师,r_OldInfo.主任医师,当前病况_IN,操作员编号_IN,操作员姓名_IN);                
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo;
            Fetch c_OldInfo Into r_OldInfo; 
        End IF; 
        Close c_OldInfo;
    End IF;

	--处理入院诊断
    IF 入院诊断_IN is Null AND 疾病ID_IN IS NULL Then
        Delete From 病人诊断记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 诊断类型=2 AND 记录来源=2;
    Else
        Update 病人诊断记录 Set 疾病ID=疾病ID_IN,诊断描述=入院诊断_IN,记录日期=sysdate,记录人=操作员姓名_IN
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 诊断类型=2 AND 记录来源=2;
        IF SQL%RowCount=0 Then
            Insert Into 病人诊断记录(
                ID,病人ID,主页ID,记录来源,诊断类型,诊断次序,疾病ID,诊断描述,记录日期,记录人) 
            Values(
                病人诊断记录_ID.Nextval,病人ID_IN,主页ID_IN,2,2,1,疾病ID_IN,入院诊断_IN,sysdate,操作员姓名_IN);
        End IF;
    End IF;

	--处理中医入院诊断
    IF 中医诊断_IN is Null AND 中医疾病ID_IN IS NULL Then
        Delete From 病人诊断记录 Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 诊断类型=12 AND 记录来源=2;
    Else
        Update 病人诊断记录 Set 疾病ID=中医疾病ID_IN,诊断描述=中医诊断_IN,记录日期=sysdate,记录人=操作员姓名_IN
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And 诊断类型=12 AND 记录来源=2;
        IF SQL%RowCount=0 Then
            Insert Into 病人诊断记录(
                ID,病人ID,主页ID,记录来源,诊断类型,诊断次序,疾病ID,诊断描述,记录日期,记录人) 
            Values(
                病人诊断记录_ID.Nextval,病人ID_IN,主页ID_IN,2,12,1,中医疾病ID_IN,中医诊断_IN,sysdate,操作员姓名_IN);
        End IF;
    End IF;

    Select Count(*) Into v_Count
        From 病人变动记录
        Where 病人ID=病人ID_IN And 主页ID=主页ID_IN And Nvl(附加床位,0)=0
            And 开始时间 IS Not Null And 终止时间 is Null;
    If v_Count>1 Then
        v_Error:='发现病人存在非法的变动记录,当前操作不能继续！'||Chr(13)||Chr(10) ||'这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
        Raise Err_Custom;
    End IF;

Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_住院病案主页_UpDate;
/

CREATE OR REPLACE PROCEDURE ZL_病人结算记录_UPDATE(
	结帐ID_IN			病人预交记录.结帐ID%TYPE,
	保险结算_IN			VARCHAR2,--"结算方式|结算金额||....."
	结帐_IN				NUMBER:=0,
	缺省结算方式_IN VARCHAR2:=NULL
) AS
	--该游标为要删除的由费用记录产生的结算记录
	CURSOR C_DEL IS
		SELECT A.*,B.性质 FROM 病人预交记录 A,结算方式 B 
		WHERE A.结算方式=B.名称 AND A.结帐ID=结帐ID_IN;

	--相关信息
	V_NO			病人预交记录.NO%TYPE;
	V_病人ID		病人费用记录.病人ID%TYPE;
	V_主页ID		病人费用记录.主页ID%TYPE;
	V_登记时间		病人费用记录.登记时间%TYPE;
	V_操作员编号	病人费用记录.操作员编号%TYPE;
	V_操作员姓名	病人费用记录.操作员姓名%TYPE;
	
	--本次结算变量
	V_金额合计	    病人预交记录.冲预交%TYPE;
	V_冲预交额	    病人预交记录.冲预交%TYPE;
    
	--保险结算
	V_保险结算	VARCHAR2(255);
	V_当前结算	VARCHAR2(50);
	V_结算方式	病人预交记录.结算方式%TYPE;
	V_结算金额	病人预交记录.冲预交%TYPE;

	v_记录性质	病人预交记录.记录性质%Type;
	v_缺省 病人预交记录.结算方式%TYPE;
        
    --分币处理及误差变量
    v_险类          保险帐户.险类%TYPE;    
    v_CentMode      Number;
    v_现金金额	    病人预交记录.冲预交%TYPE;
    v_CashCented    病人预交记录.冲预交%TYPE;
    v_误差金额      病人预交记录.冲预交%TYPE;
    v_费用ID		病人费用记录.ID%Type;
    v_序号			病人费用记录.序号%Type;
    v_收费类别		病人费用记录.收费类别%Type;
    v_收费细目ID	病人费用记录.收费细目ID%Type;    
    v_收入项目ID	病人费用记录.收入项目ID%Type;
    v_收据费目		病人费用记录.收据费目%Type;
    
	--临时变量    
	ERR_CUSTOM	EXCEPTION;
	V_ERROR		VARCHAR2(255);
BEGIN
	--如果缺省结算方式为空，则取现金结算方式
	IF 缺省结算方式_IN IS NULL THEN 
		BEGIN 
			SELECT 名称 INTO v_缺省 FROM 结算方式 WHERE 性质=1 AND ROWNUM<2;
		EXCEPTION 
			WHEN OTHERS THEN v_缺省:='现金';
		END ;
	ELSE 
		v_缺省:=缺省结算方式_IN;
	END IF ;

	--取得本次结算的相关信息
	IF NVL(结帐_IN,0)=1 THEN
		SELECT NO,病人ID,收费时间,操作员编号,操作员姓名
			INTO V_NO,V_病人ID,V_登记时间,V_操作员编号,V_操作员姓名
		FROM 病人结帐记录 WHERE ID=结帐ID_IN;
	ELSE
		SELECT NO,病人ID,登记时间,操作员编号,操作员姓名
			INTO V_NO,V_病人ID,V_登记时间,V_操作员编号,V_操作员姓名
		FROM 病人费用记录 WHERE 结帐ID=结帐ID_IN AND ROWNUM=1;

		Begin --20071027 陈东
			Select 记录性质 Into v_记录性质
			From 病人预交记录 Where 结帐ID=结帐ID_IN And Rownum=1;
		Exception --20071027 陈东
			WHEN OTHERS Then v_记录性质:=-1; --20071027 陈东
		End; --20071027 陈东

	END IF;
	IF NVL(V_病人ID,0)<>0 THEN
		SELECT 住院次数 INTO V_主页ID FROM 病人信息 WHERE 病人ID=V_病人ID;
	END IF;
	
	--回退缴款,预交不动,因为没有改冲预交的
	V_金额合计:=0;V_冲预交额:=0;
	FOR R_DEL IN C_DEL LOOP
		IF r_Del.记录性质 Not IN(1,11) THEN 
			UPDATE 人员缴款余额
				SET 余额=NVL(余额,0)-R_DEL.冲预交
			 WHERE 收款员=V_操作员姓名 AND 性质=1
				AND 结算方式=R_DEL.结算方式;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO 人员缴款余额(
					收款员,结算方式,性质,余额)
				VALUES(
					V_操作员姓名,R_DEL.结算方式,1,-1*R_DEL.冲预交);
			END IF;
			
			V_金额合计:=V_金额合计+R_DEL.冲预交;

			DELETE FROM 病人预交记录 WHERE ID=R_DEL.ID;
		END IF ;
	END LOOP;
	
	--------------------------------------------------------------------------------------------------------------
	--------------------------------------------------------------------------------------------------------------
	--产生医保支付结算
	IF 保险结算_IN IS NOT NULL THEN 
		--各个保险结算	
		V_保险结算:=保险结算_IN||'||';
		WHILE V_保险结算 IS NOT NULL LOOP
			V_当前结算:=SUBSTR(V_保险结算,1,INSTR(V_保险结算,'||')-1);

			V_结算方式:=SUBSTR(V_当前结算,1,INSTR(V_当前结算,'|')-1);
			V_结算金额:=TO_NUMBER(SUBSTR(V_当前结算,INSTR(V_当前结算,'|')+1));

			INSERT INTO 病人预交记录(
				ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
			VALUES(
				病人预交记录_ID.NEXTVAL,DECODE(结帐_IN,1,2,v_记录性质),V_NO,1,V_病人ID,V_主页ID,'保险部份',
				V_结算方式,V_登记时间,V_操作员编号,V_操作员姓名,V_结算金额,结帐ID_IN);
			
			V_金额合计:=V_金额合计-V_结算金额;

			V_保险结算:=SUBSTR(V_保险结算,INSTR(V_保险结算,'||')+2);
		END LOOP;
	END IF;

	--剩余部份全部用缺省结算方式结算，(小于零也不进行额外处理)
	IF V_金额合计<>0 Then             
		UPDATE 病人预交记录
			SET 冲预交=冲预交+V_金额合计
		WHERE 结帐ID=结帐ID_IN AND 结算方式=v_缺省 AND 记录性质=DECODE(结帐_IN,1,2,v_记录性质);
		IF SQL%ROWCOUNT=0 THEN 
			INSERT INTO 病人预交记录(
				ID,记录性质,NO,记录状态,病人ID,主页ID,摘要,结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
			VALUES(
				病人预交记录_ID.NEXTVAL,DECODE(结帐_IN,1,2,v_记录性质),V_NO,1,V_病人ID,V_主页ID,'现金部份',v_缺省,
				V_登记时间,V_操作员编号,V_操作员姓名,V_金额合计,结帐ID_IN);
		END IF ;
        
        --挂号结算,分币处理(由于挂号界面没有预结算,所以在此过程中根据分币处理规则来修正)
        If v_记录性质=4 Then
           Begin
               Select 险类 Into v_险类 From 保险帐户 Where 病人Id=V_病人Id;                              
           Exception
               When Others Then v_险类:=0;
           End;
           If v_险类=413 Then --上海医保,挂号支持分币处理
               Begin
                   Select A.冲预交 Into V_现金金额 From 病人预交记录 A,结算方式 B
                   Where A.结算方式=B.名称 And B.性质=1 And A.结帐Id=结帐ID_IN;
               Exception
                   When Others Then V_现金金额:=0;
               End;           
               If FLOOR(V_现金金额*10)<>V_现金金额*10 Then    
                   --v_CentMode:=Nvl(zl_GetSysParameter(14),0);              
                   --If v_CentMode=1 Then                                                                      
                   --   v_CashCented:=round(V_现金金额,1);                      --1.四舍五入法,eg:0.51=0.50;0.56=0.60  
                   --Elsif  v_CentMode=2 Then
                   --   v_CashCented:=CEIL(V_现金金额*10)/10;                   --2.补整收法,eg:0.51=0.60,0.56=0.60
                   --Elsif  v_CentMode=3 Then
                      v_CashCented:=FLOOR(V_现金金额*10)/10;                  --3.舍分收法,eg:0.51=0.50,0.56=0.50
                   --Else
                   --   v_CashCented:=V_现金金额;
                   --End If;        
                   v_误差金额:=v_CashCented-V_现金金额;
                   If v_误差金额<>0 Then                 
                      --1.更新预交记录(一定存在记录)
                      Update 病人预交记录 Set 冲预交=v_CashCented
                      Where 结算方式=(Select 名称 From 结算方式 Where 性质=1 And Rownum=1) And 结帐Id=结帐ID_IN;
                      
                      --2.生成误差费用记录(注:计算单位记录的是号别,所以不取误差项的)
                      Begin
                          Select A.类别,A.ID,C.ID,C.收据费目 
                          Into v_收费类别,v_收费细目ID,v_收入项目ID,v_收据费目
                          From 收费项目目录 A,收费价目 B,收入项目 C,收费特定项目 D
                          Where D.特定项目='误差项' And D.收费细目ID=A.Id  And A.ID=B.收费细目ID And B.收入项目ID=C.ID
                              And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-01-01','YYYY-MM-DD')) ;
                      Exception
                          When Others Then                        
                          v_Error:='不能正确读取收费误差项的信息，请先检查该项目是否设置正确。';
                          Raise Err_Custom;
                      End; 
                      Select 病人费用记录_ID.Nextval Into v_费用ID From Dual;
                      Select Max(序号)+1 Into v_序号 From 病人费用记录 Where 结帐ID=结帐ID_IN;
                                           
                      Insert Into 病人费用记录(
                          ID,记录性质,NO,实际票号,记录状态,序号,从属父号,价格父号,门诊标志,病人ID,标识号,床号,姓名,性别,
                          年龄,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,发药窗口,付数,数次,加班标志,附加标志,
                          收入项目ID,收据费目,标准单价,应收金额,实收金额,记帐费用,划价人,开单部门ID,开单人,发生时间,登记时间,
                          执行部门ID,执行人,执行状态,结帐ID,结帐金额,操作员编号,操作员姓名,是否上传)
                      Select
                          v_费用ID,记录性质,NO,实际票号,记录状态,v_序号,NULL,NULL,门诊标志,病人ID,标识号,床号,姓名,性别,年龄,
                          病人病区ID,病人科室ID,费别,v_收费类别,v_收费细目ID,计算单位,发药窗口,1,1,加班标志,9,
                          v_收入项目ID,v_收据费目,v_误差金额,v_误差金额,v_误差金额,记帐费用,划价人,开单部门ID,开单人,发生时间,登记时间,
                          执行部门ID,执行人,执行状态,结帐ID_IN,v_误差金额,操作员编号,操作员姓名,1
                      From 病人费用记录
                      Where 结帐ID=结帐ID_IN And Rownum=1; 
                                         
                      --3.更新"病人费用汇总"  
                      --只可能产生误差金额的变化.仅为了变量处理方便而用游标
                      For C_Error In (
                          Select TRUNC(登记时间) as 日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,门诊标志,应收金额,实收金额,结帐金额
                          From 病人费用记录
                          Where 结帐Id=结帐Id_IN And 附加标志=9
                      ) Loop
                          Update 病人费用汇总
                              Set 应收金额=Nvl(应收金额,0)+C_Error.应收金额,实收金额=Nvl(实收金额,0)+C_Error.实收金额,结帐金额=Nvl(结帐金额,0)+C_Error.结帐金额
                          Where 日期=C_Error.日期
                              And Nvl(病人病区ID,0)=Nvl(C_Error.病人病区ID,0) And Nvl(病人科室ID,0)=Nvl(C_Error.病人科室ID,0)
                              And Nvl(开单部门ID,0)=Nvl(C_Error.开单部门ID,0) And Nvl(执行部门ID,0)=Nvl(C_Error.执行部门ID,0)
                              And 收入项目ID+0=C_Error.收入项目Id And 来源途径=C_Error.门诊标志 And 记帐费用=0; 
                          If SQL%RowCount=0 Then
                              Insert Into 病人费用汇总(
                                  日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
                              Values(
                                  C_Error.日期,C_Error.病人病区ID,C_Error.病人科室ID,C_Error.开单部门ID,C_Error.执行部门ID,
                                  C_Error.收入项目ID,C_Error.门诊标志,0,C_Error.应收金额,C_Error.实收金额,C_Error.结帐金额);
                          End If;
                      End Loop;                   
                   End If;
               End If;
           End If;
        End If;        
	END IF;
	
	--最后再处理"人员缴款余额"(没有动冲预交那部分,所以"病人余额"的预交余额不用更新)
	FOR R_DEL IN C_DEL LOOP
		IF r_Del.记录性质 Not IN(1,11) THEN 
			UPDATE 人员缴款余额
				SET 余额=NVL(余额,0)+R_DEL.冲预交
			 WHERE 收款员=V_操作员姓名 AND 性质=1
				AND 结算方式=R_DEL.结算方式;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO 人员缴款余额(
					收款员,结算方式,性质,余额)
				VALUES(
					V_操作员姓名,R_DEL.结算方式,1,R_DEL.冲预交);
			END IF;
		END IF ;
	END LOOP;
	DELETE FROM 人员缴款余额 WHERE 性质=1 AND 收款员=V_操作员姓名 AND NVL(余额,0)=0;
EXCEPTION
	WHEN ERR_CUSTOM THEN RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]'||V_ERROR||'[ZLSOFT]');
	WHEN OTHERS THEN ZL_ERRORCENTER(SQLCODE,SQLERRM);
END ZL_病人结算记录_UPDATE;
/

Create Or Replace Procedure zl_住院记帐记录_Delete(
    NO_IN            病人费用记录.NO%Type,
    序号_IN            Varchar2,
    操作员编号_IN    病人费用记录.操作员编号%Type,
    操作员姓名_IN    病人费用记录.操作员姓名%Type,
    记录性质_IN        病人费用记录.记录性质%Type:=2
)
AS
     --功能：冲销一张住院记帐单据中指定序号行
     --序号：格式如"1,3,5,7,8",为空表示冲销所有可冲销行
     --记录性质:    2-人工记帐单,3-自动记帐单
    --该光标用于销帐指定费用行

    --该游标为要退费单据的所有原始记录
    Cursor c_Bill is
        Select * From 病人费用记录
        Where NO=NO_IN And 记录性质=记录性质_IN And 记录状态 IN(0,1,3) And 门诊标志=2
        Order by 收费细目ID,序号;

    --该游标用于处理药品库存可用数量
    --不要管费用的执行状态,因为先于此步处理
    Cursor c_Stock is
        Select * From 药品收发记录
        Where NO=NO_IN And 单据 IN(9,10,25,26) And Mod(记录状态,3)=1 And 审核人 IS NULL
            And 费用ID IN(
                Select ID From 病人费用记录 
                Where NO=NO_IN And 记录性质=记录性质_IN And 记录状态 IN(0,1,3) 
                    And 收费类别 IN('4','5','6','7') And 门诊标志=2
                    And (INSTR(','||序号_IN||',',','||序号||',')>0 Or 序号_IN Is Null)
                )
        Order BY 药品ID;
	r_Stock c_Stock%RowType;
    
    --该游标用于处理未发药品记录
    Cursor c_Spare is
        Select * From 未发药品记录 Where NO=NO_IN And 单据 IN(9,10,25,26);

    --该游标用于处理费用记录序号
    Cursor c_Serial is
        Select 序号,价格父号 From 病人费用记录 Where NO=NO_IN And 记录性质=记录性质_IN And 记录状态 IN(0,1,3) Order BY 序号;

	v_医嘱ID		病人医嘱记录.ID%Type;
	v_划价			Number;
	v_父号			病人费用记录.价格父号%Type;
	
    --部分退费计算变量
    v_剩余数量		Number;
    v_剩余应收		Number;
    v_剩余实收		Number;
    v_剩余统筹		Number;

    v_准退数量		Number;
    v_退费次数		Number;

    v_应收金额		Number;
    v_实收金额		Number;
    v_统筹金额		Number;

    v_Dec			Number;    
    v_Count			Number;
    v_CurDate		Date;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
    --是否已经全部完全执行(只是整张单据的检查)
    Select Nvl(Count(*),0) Into v_Count 
    From 病人费用记录 
    Where NO=NO_IN And 记录性质=记录性质_IN And 记录状态 IN(0,1,3) And Nvl(执行状态,0)<>1 And 门诊标志=2;
    IF v_Count = 0 Then
        v_Error := '该单据中的项目已经全部完全执行！';
        Raise Err_Custom;
    End IF;

    --未完全执行的项目是否有剩余数量(只是整张单据的检查)
    Select Nvl(Count(*),0) Into v_Count
    From (
        Select 序号,Sum(数量) as 剩余数量
        From (
            Select 记录状态,Nvl(价格父号,序号) as 序号,
                Avg(Nvl(付数,1)*数次) as 数量 
            From 病人费用记录
            Where NO=NO_IN And 记录性质=记录性质_IN And 门诊标志=2
                And Nvl(价格父号,序号) IN (
                        Select Nvl(价格父号,序号) 
                        From 病人费用记录 
                        Where NO=NO_IN And 记录性质=记录性质_IN 
                            And 记录状态 IN(0,1,3) And Nvl(执行状态,0)<>1)
            Group by 记录状态,Nvl(价格父号,序号)
            )
        Group by 序号 Having Sum(数量)<>0);
    IF v_Count = 0 Then
        v_Error := '该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！';
        Raise Err_Custom;
    End IF;
    
    ---------------------------------------------------------------------------------
	--先打开药品对应数据集,以确保当前条件下有数据,为了处理并发判断
	--不能在游标条件中取消"审核人 is Null"条件，因为多次退药可能部份又已发
	Open c_Stock;

    --公用变量
    Select Sysdate Into v_CurDate From Dual;
    --金额小数位数
    Begin
        Select To_Number(参数值) Into v_Dec From 系统参数表 Where 参数号=9;
    Exception
        When Others Then v_Dec:=2;
    End;
    
    --循环处理每行费用(收入项目行)
	For r_Bill IN c_Bill Loop
		IF INSTR(','||序号_IN||',',','||Nvl(r_Bill.价格父号,r_Bill.序号)||',') >0 Or 序号_IN Is Null Then
			Select Decode(记录状态,0,1,0) Into v_划价 From 病人费用记录 Where ID=r_Bill.ID;
			If v_划价=0 Then
				IF Nvl(r_Bill.执行状态,0)<>1 Then
					--求剩余数量,剩余应收,剩余实收
					Select 
						Sum(Nvl(付数,1)*数次),Sum(应收金额),Sum(实收金额),Sum(统筹金额)
						Into v_剩余数量,v_剩余应收,v_剩余实收,v_剩余统筹
					From 病人费用记录 
					Where NO=NO_IN And 记录性质=记录性质_IN And 序号=r_Bill.序号;

					IF v_剩余数量=0 Then
						IF 序号_IN IS Not NULL Then 
							v_Error := '单据中第'||Nvl(r_Bill.价格父号,r_Bill.序号)||'行费用已经全部销帐！';
							Raise Err_Custom;
						End IF;
						--情况：未限定行号,原始单据中的该笔已经全部销帐(执行状态=0的一种可能)
					Else
						--准销数量(非药品项目为剩余数量,原始数量)
						IF Instr(',4,5,6,7,',r_Bill.收费类别)=0 Then
							v_准退数量:=v_剩余数量;
						Else
							Select Sum(Nvl(付数,1)*实际数量) Into v_准退数量
							From 药品收发记录
							Where NO=NO_IN And 单据 IN(9,10,25,26) And MOD(记录状态,3)=1 
								And 审核人 is NULL And 费用ID=r_Bill.ID;

							--不跟踪在用的卫生材料
							If r_Bill.收费类别='4' And Nvl(v_准退数量,0)=0 Then
								v_准退数量:=v_剩余数量;
							End IF;
						End if;

						--处理病人费用记录
						
						--该笔项目第几次销帐
						Select Nvl(Max(Abs(执行状态)),0)+1 Into v_退费次数
						From 病人费用记录 
						Where NO=NO_IN And 记录性质=记录性质_IN And 记录状态=2 And 序号=r_Bill.序号 And 门诊标志=2;
						
						--金额=剩余金额*(准退数/剩余数)
						v_应收金额:=Round(v_剩余应收*(v_准退数量/v_剩余数量),v_Dec);
						v_实收金额:=Round(v_剩余实收*(v_准退数量/v_剩余数量),v_Dec);
						v_统筹金额:=Round(v_剩余统筹*(v_准退数量/v_剩余数量),v_Dec);

						--插入退费记录
						Insert Into 病人费用记录(
							ID,NO,记录性质,记录状态,序号,从属父号,价格父号,主页ID,病人ID,医嘱序号,门诊标志,多病人单,婴儿费,姓名,
							性别,年龄,标识号,床号,费别,病人病区ID,病人科室ID,收费类别,收费细目ID,计算单位,付数,发药窗口,
							数次,加班标志,附加标志,收入项目ID,收据费目,记帐费用,标准单价,应收金额,实收金额,开单部门ID,
							开单人,执行部门ID,划价人,执行人,执行状态,执行时间,操作员编号,操作员姓名,发生时间,登记时间,
							保险项目否,保险大类ID,统筹金额,保险编码,记帐单ID,摘要)
						Select 病人费用记录_ID.Nextval,NO,记录性质,2,序号,从属父号,价格父号,主页ID,病人ID,医嘱序号,门诊标志,多病人单,
							婴儿费,姓名,性别,年龄,标识号,床号,费别,病人病区ID,病人科室ID,收费类别,收费细目ID,计算单位,
							Decode(Sign(v_准退数量-Nvl(付数,1)*数次),0,付数,1),发药窗口,
							Decode(Sign(v_准退数量-Nvl(付数,1)*数次),0,-1*数次,-1*v_准退数量),加班标志,附加标志,
							收入项目ID,收据费目,记帐费用,标准单价,-1*v_应收金额,-1*v_实收金额,开单部门ID,开单人,执行部门ID,
							划价人,执行人,-1*v_退费次数,执行时间,操作员编号_IN,操作员姓名_IN,发生时间,v_CurDate,
							保险项目否,保险大类ID,-1*v_统筹金额,保险编码,记帐单ID,摘要
						From 病人费用记录 Where ID=r_Bill.ID;
						
						--记录病人医嘱附费对应的医嘱ID(不是主费用)
						If v_医嘱ID IS Null And r_Bill.医嘱序号 IS Not Null Then
							v_医嘱ID:=r_Bill.医嘱序号;
						End IF;

						--病人余额
						Update 病人余额
							Set 费用余额=Nvl(费用余额,0) - v_实收金额
						 Where 病人ID=r_Bill.病人ID And 性质=1;
						IF SQL%RowCount=0 Then
							Insert Into 病人余额(
								病人ID,性质,费用余额,预交余额)
							Values(
								r_Bill.病人ID,1,-1*v_实收金额,0);
						End IF;
						
						--病人未结费用
						Update 病人未结费用
							Set 金额=Nvl(金额,0) - v_实收金额
						 Where 病人ID=r_Bill.病人ID
							And Nvl(主页ID,0)=Nvl(r_Bill.主页ID,0)
							And Nvl(病人病区ID,0)=Nvl(r_Bill.病人病区ID,0)
							And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
							And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
							And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
							And 收入项目ID+0=r_Bill.收入项目ID And 来源途径+0=2;
						IF SQL%RowCount=0 Then
							Insert Into 病人未结费用(
								病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
							Values(
								r_Bill.病人ID,r_Bill.主页ID,r_Bill.病人病区ID,r_Bill.病人科室ID,
								r_Bill.开单部门ID,r_Bill.执行部门ID,r_Bill.收入项目ID,2,-1*v_实收金额);
						End IF;

						--处理病人费用汇总
						Update 病人费用汇总
							Set 应收金额=Nvl(应收金额,0) - v_应收金额,
								实收金额=Nvl(实收金额,0) - v_实收金额
						 Where 日期=Trunc(v_CurDate)
							And Nvl(病人病区ID,0)=Nvl(r_Bill.病人病区ID,0)
							And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
							And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
							And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
							And 收入项目ID+0=r_Bill.收入项目ID
							And 来源途径=2 And 记帐费用=1;
						IF SQL%RowCount=0 Then
							Insert Into 病人费用汇总(
								日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
							Values(
								Trunc(v_CurDate),r_Bill.病人病区ID,r_Bill.病人科室ID,r_Bill.开单部门ID,r_Bill.执行部门ID,
								r_Bill.收入项目ID,2,1,-1 * v_应收金额,-1 * v_实收金额,0);
						End IF;
						
						--标记原费用记录
						--执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1
						Update 病人费用记录 
							Set 记录状态=3,
								执行状态=Decode(Sign(v_准退数量-v_剩余数量),0,0,1) 
						Where ID=r_Bill.ID;
					End IF;
				Else
					IF 序号_IN Is Not Null Then
						v_Error := '单据中第'||Nvl(r_Bill.价格父号,r_Bill.序号)||'行费用已经完全执行,不能销帐！';
						Raise Err_Custom;
					End IF;
					--情况:没限定行号,原始单据中包括已经完全执行的
				End IF;
			End IF;
		End IF;
	End Loop;
	
    ---------------------------------------------------------------------------------
    --药品相关内容
	Fetch c_Stock Into r_Stock;
    While c_Stock%Found Loop
        --处理药品库存
        If r_Stock.库房ID IS Not NULL Then
            Update 药品库存
                Set 可用数量=Nvl(可用数量,0)+Nvl(r_Stock.付数,1)*Nvl(r_Stock.实际数量,0)
             Where 库房ID=r_Stock.库房ID And 药品ID=r_Stock.药品ID
                And Nvl(批次,0)=Nvl(r_Stock.批次,0) And 性质=1;
            IF SQL%RowCount=0 Then
                Insert Into 药品库存(
                    库房ID,药品ID,性质,批次,效期,可用数量,上次批号,上次产地,灭菌效期)
                Values(
                    r_Stock.库房ID,r_Stock.药品ID,1,r_Stock.批次,r_Stock.效期,
                    Nvl(r_Stock.付数,1)*Nvl(r_Stock.实际数量,0),r_Stock.批号,r_Stock.产地,r_Stock.灭菌效期);
            End IF;
        End IF;

        --删除药品收发记录(加上并发操作检查:审核人 Is Null)
        Delete From 药品收发记录 Where ID=r_Stock.ID And 审核人 Is Null;
		IF SQL%RowCount=0 Then
			If r_Stock.单据 IN(9,10) Then
				v_Error:='要销帐的费用中存在已发药的药品，或已被其他人销帐；这可能是并发操作引起的。';
			Else
				v_Error:='要销帐的费用中存在已发料的卫材，或已被其他人销帐；这可能是并发操作引起的。';
			End IF;
			Raise Err_Custom;
		End IF;

		Fetch c_Stock Into r_Stock;
    End Loop;
	Close c_Stock;

    --未发药品记录
    For r_Spare IN c_Spare Loop
        Select Nvl(Count(*),0) Into v_Count
        From 药品收发记录 
        Where NO=NO_IN And 单据=r_Spare.单据 And Mod(记录状态,3)=1 
            And 审核人 is NULL And Nvl(库房ID,0)=Nvl(r_Spare.库房ID,0);
        If v_Count=0 Then
            Delete From 未发药品记录 Where 单据=r_Spare.单据 And NO=NO_IN And Nvl(库房ID,0)=Nvl(r_Spare.库房ID,0);
        End IF;
    End Loop;

	---------------------------------------------------------------------------------
	--如果是划价,直接删除费用记录(药品处理后)
	v_Count:=0;
	For r_Bill IN c_Bill Loop
		IF INSTR(','||序号_IN||',',','||Nvl(r_Bill.价格父号,r_Bill.序号)||',') >0 Or 序号_IN Is Null Then
			Select Decode(记录状态,0,1,0) Into v_划价 From 病人费用记录 Where ID=r_Bill.ID;
			If v_划价=1 Then
				IF Nvl(r_Bill.执行状态,0)<>1 Then
					Delete From 病人费用记录 Where ID=r_Bill.ID;
					v_Count:=v_Count+1;--记录是否有删除行

					--记录病人医嘱附费对应的医嘱ID(不是主费用)
					If v_医嘱ID IS Null And r_Bill.医嘱序号 IS Not Null Then
						v_医嘱ID:=r_Bill.医嘱序号;
					End IF;
				Else
					IF 序号_IN Is Not Null Then
						v_Error := '单据中第'||Nvl(r_Bill.价格父号,r_Bill.序号)||'行费用已经完全执行,不能销帐！';
						Raise Err_Custom;
					End IF;
					--情况:没限定行号,原始单据中包括已经完全执行的
				End IF;
			End IF;
		End IF;
	End Loop;

	--删除之后再统一调整序号
	If v_Count>0 Then
		v_Count:=1;
		For r_Serial In c_Serial Loop
			If r_Serial.价格父号 IS NULL Then 
				v_父号:=v_Count;
			End IF;

			Update 病人费用记录 
				Set 序号=v_Count,
					价格父号=Decode(价格父号,NULL,NULL,v_父号)
			Where NO=NO_IN And 记录性质=记录性质_IN And 序号=r_Serial.序号;
			
			Update 病人费用记录
				Set 从属父号=v_Count
			Where NO=NO_IN And 记录性质=记录性质_IN And 从属父号=r_Serial.序号;

			v_Count:=v_Count+1;
		End Loop;
	End IF;

	--整张单据全部冲完时，删除病人医嘱附费
	If 序号_IN IS NULL And v_医嘱ID IS Not NULL Then
		Select Nvl(Count(*),0) Into v_Count
		From (
			Select 序号,Sum(数量) as 剩余数量
			From (
				Select 记录状态,Nvl(价格父号,序号) as 序号,
					Avg(Nvl(付数,1)*数次) as 数量 
				From 病人费用记录
				Where NO=NO_IN And 记录性质=2 And 医嘱序号+0=v_医嘱ID
				Group by 记录状态,Nvl(价格父号,序号)
				)
			Group by 序号 Having Nvl(Sum(数量),0)<>0);
		IF v_Count = 0 Then
			Delete From 病人医嘱附费 Where 医嘱ID=v_医嘱ID And 记录性质=2 And NO=NO_IN;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_住院记帐记录_Delete;
/

CREATE OR REPLACE Procedure zl_住院一次费用_Delete(
    病人ID_IN        病人费用记录.病人ID%Type,
    主页ID_IN        病人费用记录.主页ID%Type
)
AS
    --功能：删除住院病人计算的一次性费用。
    Cursor c_Money is
        Select 病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,
            Nvl(Sum(应收金额),0) AS 应收金额,Nvl(Sum(实收金额),0) AS 实收金额
        From 病人费用记录
        Where 记录性质=3 And 记录状态=1 And 附加标志=8 And 病人ID=病人ID_IN And 主页ID=主页ID_IN
        Group BY 病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID;

    v_人员编号        病人费用记录.操作员编号%Type;
    v_人员姓名        病人费用记录.操作员姓名%Type;
    
    v_Date        Date;
    v_Temp        Varchar2(255);
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --取操作员信息(部门ID,部门名称;人员ID,人员编号,人员姓名)
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

    Select Sysdate Into v_Date From Dual;
    
    --产生作废记录
    Insert Into 病人费用记录(
        ID,记录性质,NO,记录状态,序号,价格父号,病人ID,主页ID,门诊标志,记帐费用,姓名,性别,年龄,标识号,
        床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,数次,附加标志,收入项目ID,
        收据费目,标准单价,应收金额,实收金额,划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,
        操作员编号,操作员姓名)
    Select
        病人费用记录_ID.Nextval,记录性质,NO,2,序号,价格父号,病人ID,主页ID,门诊标志,记帐费用,姓名,
        性别,年龄,标识号,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,-1*数次,
        附加标志,收入项目ID,收据费目,标准单价,-1*应收金额,-1*实收金额,划价人,开单部门ID,开单人,发生时间,
        登记时间,执行部门ID,v_人员编号,v_人员姓名
    From 病人费用记录
    Where 记录性质=3 And 记录状态=1 And 附加标志=8 And 病人ID=病人ID_IN And 主页ID=主页ID_IN;
    
    --处理汇总表
    For r_Money IN c_Money Loop
        --病人余额
        Update 病人余额
            Set 费用余额=Nvl(费用余额,0)-r_Money.实收金额
         Where 病人ID=病人ID_IN And 性质=1;
        IF SQL%RowCount=0 Then
            Insert Into 病人余额(
                病人ID,性质,费用余额,预交余额)
            Values(
                病人ID_IN,1,-1*r_Money.实收金额,0);
        End IF;

        --病人未结费用
        Update 病人未结费用
            Set 金额=Nvl(金额,0)-r_Money.实收金额
         Where 病人ID=病人ID_IN
            And Nvl(主页ID,0)=Nvl(主页ID_IN,0)
            And Nvl(病人病区ID,0)=Nvl(r_Money.病人病区ID,0)
            And Nvl(病人科室ID,0)=Nvl(r_Money.病人科室ID,0)
            And Nvl(开单部门ID,0)=r_Money.开单部门ID
            And Nvl(执行部门ID,0)=r_Money.执行部门ID
            And 收入项目ID+0=r_Money.收入项目ID
            And 来源途径=2;

        IF SQL%RowCount=0 Then
            Insert Into 病人未结费用(
                病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
            Values(
                病人ID_IN,主页ID_IN,r_Money.病人病区ID,r_Money.病人科室ID,r_Money.开单部门ID,r_Money.执行部门ID,r_Money.收入项目ID,2,-1*r_Money.实收金额);
        End IF;

        --病人费用汇总
        Update 病人费用汇总
            Set 应收金额=Nvl(应收金额,0)-r_Money.应收金额,
                 实收金额=Nvl(实收金额,0)-r_Money.实收金额
         Where 日期=Trunc(v_Date)
            And Nvl(病人病区ID,0)=Nvl(r_Money.病人病区ID,0)
            And Nvl(病人科室ID,0)=Nvl(r_Money.病人科室ID,0)
            And Nvl(开单部门ID,0)=r_Money.开单部门ID
            And Nvl(执行部门ID,0)=r_Money.执行部门ID
            And 收入项目ID+0=r_Money.收入项目ID
            And 来源途径=2 And 记帐费用=1;
        IF SQL%RowCount=0 Then
            Insert Into 病人费用汇总(
                日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
            Values(
                Trunc(v_Date),r_Money.病人病区ID,r_Money.病人科室ID,r_Money.开单部门ID,r_Money.执行部门ID,r_Money.收入项目ID,2,1,-1*r_Money.应收金额,-1*r_Money.实收金额,0);
        End IF;
    End Loop;

    --更改原始记录
    Update 病人费用记录 Set 记录状态=3 Where 记录性质=3 And 记录状态=1 And 附加标志=8 And 病人ID=病人ID_IN And 主页ID=主页ID_IN;
Exception
    When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_住院一次费用_Delete;
/

CREATE OR REPLACE Procedure zl_住院一次费用_Insert(
    病人ID_IN        病人费用记录.病人ID%Type,
    主页ID_IN        病人费用记录.主页ID%Type
)
AS
    Cursor c_Money is
        Select 
            E.姓名,E.性别,E.年龄,E.住院号,D.出院病床,D.入院病区ID,D.入院科室ID,D.费别,
            A.类别,C.收费细目ID,A.计算单位,B.收入项目ID,F.收据费目,B.现价,
            D.入院日期,Nvl(A.执行科室,0) AS 执行科室,Nvl(A.屏蔽费别,0) AS 屏蔽费别
        From 收费细目 A,收费价目 B,自动计价项目 C,病案主页 D,病人信息 E,收入项目 F
        Where A.ID=B.收费细目ID And A.ID=C.收费细目ID 
            And C.病区ID=D.入院病区ID And C.计算标志=8 And D.入院日期>=Nvl(C.启用日期,To_Date('3000-01-01','YYYY-MM-DD'))
            And D.主页ID=主页ID_IN And D.病人ID=病人ID_IN 
            And E.病人ID=D.病人ID And B.收入项目ID=F.ID
            And ((D.入院日期 Between B.执行日期 and B.终止日期) or (D.入院日期>=B.执行日期 And B.终止日期 is NULL))
        Order BY A.ID,B.收入项目ID;

    --功能：对住院病人计算一次性费用。
    v_BillNO        病人费用记录.NO%Type;
    v_执行部门ID	病人费用记录.执行部门ID%Type;
    v_实收金额      病人费用记录.实收金额%Type;
    v_价格父号      病人费用记录.价格父号%Type;
    v_项目ID        收费项目目录.ID%Type;

    v_人员编号      病人费用记录.操作员编号%Type;
    v_人员姓名      病人费用记录.操作员姓名%Type;
    v_人员部门ID    部门表.ID%Type;

    v_Dec			Number;        
    v_Date			Date;
    v_Count			Number;
    v_Temp			Varchar2(255);
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --检查是否是计算一次的项目
    Select Count(*) Into v_Count
    From 病案主页 A,自动计价项目 B
    Where A.病人ID=病人ID_IN And A.主页ID=主页ID_IN
        And A.入院病区ID=B.病区ID And B.计算标志=8 And A.入院日期>=Nvl(B.启用日期,To_Date('3000-01-01','YYYY-MM-DD'));
    If v_Count=0 Then 
        Return;
    End IF;

    --检查该病人本次住院是否已经计算过
    Select 
        Count(*) Into v_Count 
    From 病人费用记录 
    Where 病人ID=病人ID_IN And 主页ID=主页ID_IN
        And 记录性质=3 And 记录状态=1 And 附加标志=8;
    If v_Count>0 Then 
        Return;
    End IF;
    
    --取操作员信息(部门ID,部门名称;人员ID,人员编号,人员姓名)
    v_Temp:=zl_Identity;
    v_人员部门ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --取单据号
    v_BillNO:=zl1_BillAutoNO;
    Update 号码控制表 Set 最大号码=v_BillNO Where 项目名称='自动记帐号';
    
    --取时间
    Select Sysdate Into v_Date From Dual;

    --金额小数位数
    Begin
        Select To_Number(参数值) Into v_Dec From 系统参数表 Where 参数号=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --产生费用信息
    v_项目ID:=NULL;
    v_Count:=1;--序号
    For r_Money In c_Money Loop
        If Nvl(v_项目ID,0)<>r_Money.收费细目ID Then
            --求执行部门
            IF r_Money.执行科室=2 Then
                --入住病区    
                v_执行部门ID:=r_Money.入院病区ID;
            ElsIF r_Money.执行科室=1 Then
                --指定科室
                Begin
                    Select 执行部门ID Into v_执行部门ID From 收费执行部门 Where 收费细目ID=r_Money.收费细目ID And Rownum<2;
                Exception
                    When Others Then v_执行部门ID:=v_人员部门ID;
                End;
            Else
                --未指定或操作员科室
                v_执行部门ID:=v_人员部门ID;
            End IF;
            --该项目各后续收入项目的价格父号
            v_价格父号:=v_Count;

        End IF;

        --求实收金额
        IF r_Money.屏蔽费别=1 Then
            v_实收金额:=Round(r_Money.现价,v_Dec);
        Else
            Begin
                Select 
                    Round(Round(r_Money.现价,5)*实收比率/100,v_Dec) Into v_实收金额
                From 费别明细 
                Where 收入项目ID=r_Money.收入项目ID And 费别=r_Money.费别 
                    And Round(r_Money.现价,5) Between 应收段首值 and 应收段尾值;
            Exception
                When Others Then v_实收金额:=Round(r_Money.现价,v_Dec);
            End;
        End IF;
        
        --插入费用记录(附加标志=8,记录性质=3)
        Insert Into 病人费用记录(
            ID,记录性质,NO,记录状态,序号,价格父号,病人ID,主页ID,门诊标志,记帐费用,姓名,性别,年龄,标识号,
            床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,数次,附加标志,收入项目ID,
            收据费目,标准单价,应收金额,实收金额,划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,
            操作员编号,操作员姓名)
        Values(
            病人费用记录_ID.Nextval,3,v_BillNo,1,v_Count,Decode(Sign(Nvl(v_项目ID,0)-r_Money.收费细目ID),0,v_价格父号,NULL),
            病人ID_IN,主页ID_IN,2,1,r_Money.姓名,r_Money.性别,r_Money.年龄,r_Money.住院号,r_Money.出院病床,
            r_Money.入院病区ID,r_Money.入院科室ID,r_Money.费别,r_Money.类别,r_Money.收费细目ID,r_Money.计算单位,
            1,1,8,r_Money.收入项目ID,r_Money.收据费目,Round(r_Money.现价,5),Round(r_Money.现价,v_Dec),v_实收金额,
            v_人员姓名,v_人员部门ID,v_人员姓名,r_Money.入院日期,v_Date,v_执行部门ID,v_人员编号,v_人员姓名);

        v_Count:=v_Count+1;
        v_项目ID:=r_Money.收费细目ID;--记录上次处理行的项目

        --相关汇总表的处理
        --病人余额
        Update 病人余额
            Set 费用余额=Nvl(费用余额,0)+v_实收金额
         Where 病人ID=病人ID_IN And 性质=1;
        IF SQL%RowCount=0 Then
            Insert Into 病人余额(
                病人ID,性质,费用余额,预交余额)
            Values(
                病人ID_IN,1,v_实收金额,0);
        End IF;

        --病人未结费用
        Update 病人未结费用
            Set 金额=Nvl(金额,0)+v_实收金额
         Where 病人ID=病人ID_IN
            And Nvl(主页ID,0)=Nvl(主页ID_IN,0)
            And Nvl(病人病区ID,0)=Nvl(r_Money.入院病区ID,0)
            And Nvl(病人科室ID,0)=Nvl(r_Money.入院科室ID,0)
            And Nvl(开单部门ID,0)=v_人员部门ID
            And Nvl(执行部门ID,0)=v_执行部门ID
            And 收入项目ID+0=r_Money.收入项目ID
            And 来源途径=2;

        IF SQL%RowCount=0 Then
            Insert Into 病人未结费用(
                病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
            Values(
                病人ID_IN,主页ID_IN,r_Money.入院病区ID,r_Money.入院科室ID,v_人员部门ID,v_执行部门ID,r_Money.收入项目ID,2,v_实收金额);
        End IF;

        --病人费用汇总
        Update 病人费用汇总
            Set 应收金额=Nvl(应收金额,0)+Round(r_Money.现价,v_Dec),
                 实收金额=Nvl(实收金额,0)+v_实收金额
         Where 日期=Trunc(v_Date)
            And Nvl(病人病区ID,0)=Nvl(r_Money.入院病区ID,0)
            And Nvl(病人科室ID,0)=Nvl(r_Money.入院科室ID,0)
            And Nvl(开单部门ID,0)=v_人员部门ID
            And Nvl(执行部门ID,0)=v_执行部门ID
            And 收入项目ID+0=r_Money.收入项目ID
            And 来源途径=2 And 记帐费用=1;
        IF SQL%RowCount=0 Then
            Insert Into 病人费用汇总(
                日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
            Values(
                Trunc(v_Date),r_Money.入院病区ID,r_Money.入院科室ID,v_人员部门ID,v_执行部门ID,r_Money.收入项目ID,2,1,Round(r_Money.现价,v_Dec),v_实收金额,0);
        End IF;
    End Loop;
Exception
    When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_住院一次费用_Insert;
/

-------------------------------------------------------
--模块：病人挂号记录.SQL
Create Or Replace Procedure zl_病人挂号记录_Delete(
    单据号_IN                   病人费用记录.NO%Type,
    操作员编号_IN           病人费用记录.操作员编号%Type,
    操作员姓名_IN           病人费用记录.操作员姓名%Type,
    删除门诊号_IN           Number:=0,
    医保退费结算_IN		Varchar2:=Null                     --医保不允许的退费结算方式,空表示全部允许
) AS
    --该游标用于判断是否单独收病历费,及挂号汇总表处理
    Cursor c_RegistInfo(v_状态 病人费用记录.记录状态%Type) IS
        Select A.发生时间,A.登记时间,B.项目ID,B.科室ID,B.医生姓名,B.医生ID
        From 病人费用记录 A,挂号安排 B
        Where A.记录性质=4 And A.记录状态=v_状态
            And A.NO=单据号_IN
            And Nvl(A.计算单位,'号别')=B.号码
            And ROWNUM=1;
    r_RegistRow c_RegistInfo%ROWTYPE;

    --该游标用于判断记录是否存在,及费用汇总表处理
    Cursor c_MoneyInfo(v_状态 病人费用记录.记录状态%Type) IS
        Select 病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,
                Nvl(Sum(应收金额),0) AS 应收,
                Nvl(Sum(实收金额),0) AS 实收,
                Nvl(Sum(结帐金额),0) AS 结帐
        From 病人费用记录
        Where 记录性质=4 And 记录状态=v_状态 And NO=单据号_IN
        Group BY 病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID;
    r_MoneyRow c_MoneyInfo%ROWTYPE;
	
    --该光标用于处理人员缴款余额中退的不同结算方式的金额
    Cursor c_OperMoney is
		Select DISTINCT B.结算方式,B.冲预交 
		From 病人费用记录 A,病人预交记录 B
		Where A.结帐ID=B.结帐ID And A.NO=单据号_IN
			And A.记录性质=4 And A.记录状态=3
			And B.记录性质=4 And B.记录状态=3 And Nvl(B.冲预交,0)<>0;

	v_打印ID    票据打印内容.ID%Type;
	v_结帐ID    病人费用记录.结帐ID%Type;
	v_费用ID    病人费用记录.ID%Type;--用于作费收回票据
	v_病人ID    病人信息.病人ID%Type;
	v_预交金额	病人预交记录.冲预交%Type;
  V_退费金额  病人预交记录.冲预交%Type;
	
	v_预约挂号	Number;

	v_Count		    Number;
	v_Date		    Date;
	v_Error		    Varchar(255);
	Err_Custom	Exception;
Begin
	--首先判断要退号/取消预约的记录是否存在
    Open c_MoneyInfo(1);
    Fetch c_MoneyInfo Into r_MoneyRow;
	IF c_MoneyInfo%RowCount=0 Then
		Close c_MoneyInfo;
		Open c_MoneyInfo(0);
		Fetch c_MoneyInfo Into r_MoneyRow;
		IF c_MoneyInfo%RowCount=0 Then
			v_Error:='要处理的单据不存在。';
			Raise Err_Custom;
		End IF;
		v_预约挂号:=1;
	End IF;        
	Close c_MoneyInfo;

	If Nvl(v_预约挂号,0)=1 Then
		--减少已约数
		Open c_RegistInfo(0);
		Fetch c_RegistInfo Into r_RegistRow;
		Update 病人挂号汇总
			Set 已约数=Nvl(已约数,0) - 1
		Where 日期=Trunc(r_RegistRow.发生时间)
			And 科室ID=r_RegistRow.科室ID
			And 项目ID=r_RegistRow.项目ID
			And Nvl(医生姓名,'医生')=Nvl(r_RegistRow.医生姓名,'医生')
			And Nvl(医生ID,0)=Nvl(r_RegistRow.医生ID,0);

		IF SQL%RowCount=0 Then
			Insert Into 病人挂号汇总(
				日期,科室ID,项目ID,医生姓名,医生ID,已约数)
			Values(
				Trunc(r_RegistRow.发生时间),r_RegistRow.科室ID,r_RegistRow.项目ID,r_RegistRow.医生姓名,
				Decode(r_RegistRow.医生ID,0,Null,r_RegistRow.医生ID),-1);
		End If;
		Close c_RegistInfo;

		--删除病人费用记录
		Delete From 病人费用记录 Where NO=单据号_IN And 记录性质=4 And 记录状态=0;
	Else
		Select Sysdate,病人结帐记录_ID.Nextval Into v_Date,v_结帐ID From Dual;
		
		--病人就诊状态
		Select 病人ID Into v_病人ID From 病人费用记录 Where 记录性质=4 And 记录状态=1 And NO=单据号_IN And 序号=1;
		IF v_病人ID IS Not NULL Then
			Update 病人信息 Set 就诊状态=0,就诊诊室=NULL Where 病人ID=v_病人ID;      		
			--删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
			If 删除门诊号_IN=1 Then
				Delete 门诊病案记录 Where 病人Id=v_病人ID;
				Update 病人信息 Set 门诊号=Null Where 病人Id=v_病人ID;
				--费用记录包括挂号及病案、就诊卡费用,以及病人交费后退费或销帐的费用,挂号记录在最后处理
				Update 病人费用记录 Set 标识号=Null Where 门诊标志=1 And 病人Id=v_病人ID;	
			End If;
		End IF;

		--如果挂时收了就诊卡费,退费时清除就诊卡号
		v_病人ID:=Null;
		Begin
			Select 病人ID Into v_病人ID From 病人费用记录 Where 记录性质=4 And 记录状态=1 And NO=单据号_IN And 附加标志=2 And Rownum<2;			
		Exception
			When Others Then NULL;
		End;
		IF v_病人ID IS Not NULL Then
			Update 病人信息 Set 就诊卡号=NULL,卡验证码=NULL Where 病人ID=v_病人ID;
		End IF;

		--病人费用记录
		Insert Into 病人费用记录(
			ID,NO,实际票号,记录性质,记录状态,序号,价格父号,从属父号,病人ID,主页ID,病人病区ID,
			病人科室ID,门诊标志,标识号,姓名,性别,年龄,费别,收费类别,收费细目ID,计算单位,
			付数,数次,加班标志,附加标志,发药窗口,收入项目ID,收据费目,记帐费用,标准单价,
			应收金额,实收金额,开单部门ID,开单人,执行部门ID,执行人,操作员编号,操作员姓名,发生时间,
			登记时间,结帐ID,结帐金额)
		Select 病人费用记录_ID.Nextval,NO,实际票号,记录性质,2,序号,价格父号,从属父号,病人ID,主页ID,
			病人病区ID,病人科室ID,门诊标志,标识号,姓名,性别,年龄,费别,收费类别,收费细目ID,
			计算单位,付数,-数次,加班标志,附加标志,发药窗口,收入项目ID,收据费目,记帐费用,
			标准单价,-应收金额,-实收金额,开单部门ID,开单人,执行部门ID,执行人,操作员编号_IN,
			操作员姓名_IN,发生时间,v_Date,v_结帐ID,-结帐金额
		From 病人费用记录
		Where 记录性质=4 And 记录状态=1 And NO=单据号_IN;

		Update 病人费用记录 Set 记录状态=3 Where 记录性质=4 And 记录状态=1 And NO=单据号_IN;

		Select 结帐ID Into v_Count From 病人费用记录 Where 记录性质=4 And 记录状态=3 And NO=单据号_IN And Rownum=1;
		--病人挂号结算:现金和个人帐户部份
		IF 医保退费结算_IN IS NOT NULL THEN         
            --a.允许的结算方式原样退
            Insert Into 病人预交记录(
              ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,摘要,
              结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
            Select 病人预交记录_ID.Nextval,NO,实际票号,记录性质,2,病人ID,
              主页ID,科室ID,摘要,结算方式,v_Date,操作员编号_IN,
              操作员姓名_IN,-冲预交,v_结帐ID
            From 病人预交记录
            Where 记录性质=4 And 记录状态=1 And 结帐ID=v_Count And 结算方式<>医保退费结算_IN;
            
            --b.不允许的退现金
            Begin
              Select 冲预交 Into V_退费金额 From 病人预交记录
                   Where 记录性质=4 And 记录状态=1 And 结帐ID=v_Count And 结算方式=医保退费结算_IN;
            Exception
                   When Others Then V_退费金额:=0;
            End;
            If V_退费金额<>0 Then
               Update 病人预交记录 Set 冲预交=冲预交-V_退费金额 Where 记录性质=4 And 记录状态=2 And 结帐ID=v_结帐ID 
                      And 结算方式=(Select 名称 From 结算方式 Where 性质=1);
                      
               If Sql%Rowcount=0 Then
                  Insert Into 病人预交记录(
                    ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,摘要,
                    结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
                  Select 病人预交记录_ID.Nextval,A.NO,A.实际票号,A.记录性质,2,A.病人ID,
                    A.主页ID,A.科室ID,A.摘要,B.名称,v_Date,操作员编号_IN,
                    操作员姓名_IN,-1*V_退费金额,v_结帐ID
                  From 病人预交记录 A,结算方式 B
                  Where B.性质=1 And A.记录性质=4 And A.记录状态=1 And A.结帐ID=v_Count And Rownum=1;
               End If;        
            End If;        
        Else
          Insert Into 病人预交记录(
                ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,摘要,
                结算方式,收款时间,操作员编号,操作员姓名,冲预交,结帐ID)
            Select 病人预交记录_ID.Nextval,NO,实际票号,记录性质,2,病人ID,
                主页ID,科室ID,摘要,结算方式,v_Date,操作员编号_IN,
                操作员姓名_IN,-冲预交,v_结帐ID
            From 病人预交记录
            Where 记录性质=4 And 记录状态=1 And 结帐ID=v_Count;
        End If;

		Update 病人预交记录 Set 记录状态=3 Where 记录性质=4 And 记录状态=1 And 结帐ID=v_Count;

		--病人挂号结算:冲预交款部份
		Insert Into 病人预交记录(
		    ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额,结算方式,结算号码,摘要,
		    缴款单位,单位开户行,单位帐号,收款时间,操作员姓名,操作员编号,冲预交,结帐ID)
		Select
		    病人预交记录_ID.Nextval,NO,实际票号,11,记录状态,病人ID,主页ID,科室ID,Null,
		    结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号,收款时间,操作员姓名,
		    操作员编号,-1*冲预交,v_结帐ID
		From 病人预交记录
		Where 记录性质 IN(1,11) And 结帐ID=v_Count;

		--处理病人预交余额
		Begin
		    Select 病人ID,Sum(Nvl(冲预交,0)) Into v_病人ID,v_预交金额 From 病人预交记录 
		    Where 记录性质 IN(1,11) And 结帐ID=v_Count 
		    Group by 病人ID;
		Exception
		    When Others Then NULL;
		End;
		IF Nvl(v_病人ID,0)<>0 And Nvl(v_预交金额,0)<>0 Then
		    Update 病人余额 Set 预交余额=Nvl(预交余额,0)+v_预交金额 Where 病人ID=v_病人ID And 性质=1;
		    IF SQL%RowCount=0 Then
			     Insert Into 病人余额(病人ID,预交余额,性质) Values(v_病人ID,v_预交金额,1);
		    End IF;
		    Delete From 病人余额 Where 病人ID=v_病人ID And 性质=1 And Nvl(预交余额,0)=0 And Nvl(费用余额,0)=0;
		End IF;

		--退卡收回票据(可能上次挂号使用票据,不能收回)
		Begin
			--从最后一次的打印内容中取
			Select Max(ID) Into v_打印ID From 票据打印内容 Where 数据性质=4 And NO=单据号_IN;
		Exception
			When Others Then NULL;
		End;
		IF v_打印ID IS Not NULL Then
			Insert Into 票据使用明细(
				ID,票种,号码,性质,原因,领用ID,打印ID,使用时间,使用人)
			Select
				票据使用明细_ID.Nextval,票种,号码,2,2,领用ID,打印ID,v_Date,操作员姓名_IN
			From 票据使用明细
			Where 打印ID=v_打印ID And 性质=1;
		End If;

		--相关汇总表的处理

		--病人挂号汇总
		Open c_RegistInfo(3);
		Fetch c_RegistInfo Into r_RegistRow;
		IF c_RegistInfo%RowCount=0 Then
			--只收病历费时无号别,不处理
			Close c_RegistInfo;
		Else
			Update 病人挂号汇总
				Set 已挂数=Nvl(已挂数,0) - 1
			Where 日期=Trunc(r_RegistRow.登记时间)
				And 科室ID=r_RegistRow.科室ID
				And 项目ID=r_RegistRow.项目ID
				And Nvl(医生姓名,'医生')=Nvl(r_RegistRow.医生姓名,'医生')
				And Nvl(医生ID,0)=Nvl(r_RegistRow.医生ID,0);

			IF SQL%RowCount=0 Then
				Insert Into 病人挂号汇总(
					日期,科室ID,项目ID,医生姓名,医生ID,已挂数)
				Values(
					Trunc(r_RegistRow.登记时间),r_RegistRow.科室ID,r_RegistRow.项目ID,r_RegistRow.医生姓名,
					Decode(r_RegistRow.医生ID,0,Null,r_RegistRow.医生ID),-1);
			End If;
			Close c_RegistInfo;
		End If;

		--病人费用汇总:一张挂号单多个收入(包括病历)
		For r_MoneyRow IN c_MoneyInfo(3) Loop
			Update 病人费用汇总
				Set 应收金额=Nvl(应收金额,0) +(-1 * r_MoneyRow.应收),
					实收金额=Nvl(实收金额,0) +(-1 * r_MoneyRow.实收),
					结帐金额=Nvl(结帐金额,0) +(-1 * r_MoneyRow.结帐)
			Where 日期=Trunc(v_Date)
				And Nvl(病人病区ID,0)=Nvl(r_MoneyRow.病人病区ID,0)
				And Nvl(病人科室ID,0)=Nvl(r_MoneyRow.病人科室ID,0)
				And Nvl(开单部门ID,0)=Nvl(r_MoneyRow.开单部门ID,0)
				And Nvl(执行部门ID,0)=Nvl(r_MoneyRow.执行部门ID,0)
				And 收入项目ID+0=r_MoneyRow.收入项目ID
				And 来源途径=1
				And 记帐费用=0;

			IF SQL%RowCount=0 Then
				Insert Into 病人费用汇总(
					日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,
					来源途径,记帐费用,应收金额,实收金额,结帐金额)
				Values(
					Trunc(v_Date),
					Decode(r_MoneyRow.病人病区ID,0,Null,r_MoneyRow.病人病区ID),
					Decode(r_MoneyRow.病人科室ID,0,Null,r_MoneyRow.病人科室ID),
					Decode(r_MoneyRow.开单部门ID,0,Null,r_MoneyRow.开单部门ID),
					Decode(r_MoneyRow.执行部门ID,0,Null,r_MoneyRow.执行部门ID),
					r_MoneyRow.收入项目ID,1,0,
					-1 * r_MoneyRow.应收,-1 * r_MoneyRow.实收,-1 * r_MoneyRow.结帐);
			End If;
		End Loop;

		--人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
		For r_OperMoney in c_OperMoney Loop
			Update 人员缴款余额
				Set 余额=Nvl(余额,0) +(-1*r_OperMoney.冲预交)
			 Where 收款员=操作员姓名_IN And 结算方式=r_OperMoney.结算方式 And 性质=1;
			IF SQL%RowCount=0 Then
				Insert Into 人员缴款余额(
					收款员,结算方式,性质,余额)
				Values(
					操作员姓名_IN,r_OperMoney.结算方式,1,-1*r_OperMoney.冲预交);
			End IF;
		End Loop;
		Delete From 人员缴款余额 Where 收款员=操作员姓名_IN And 性质=1 And Nvl(余额,0)=0;

		--病人挂号记录
		Delete From 病人挂号记录 Where NO=单据号_IN;
	End IF;
Exception
    When Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]'); 
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_病人挂号记录_Delete;
/

--功能：处理门诊病人挂号和自助挂号。
CREATE OR REPLACE Procedure zl_病人挂号记录_Insert(
    病人ID_IN			病人费用记录.病人ID%Type, 
    门诊号_IN			病人费用记录.标识号%Type, 
    姓名_IN				病人费用记录.姓名%Type, 
    性别_IN				病人费用记录.性别%Type, 
    年龄_IN				病人费用记录.年龄%Type, 
	床号_IN				病人费用记录.床号%Type,--用于存放病人的医疗付款方式编号
    费别_IN				病人费用记录.费别%Type, 
    单据号_IN			病人费用记录.NO%Type, 
    票据号_IN			病人费用记录.实际票号%Type, 
    序号_IN				病人费用记录.序号%Type, 
    价格父号_IN			病人费用记录.价格父号%Type, 
	从属父号_IN			病人费用记录.从属父号%Type,
    收费类别_IN			病人费用记录.收费类别%Type, 
    收费细目ID_IN		病人费用记录.收费细目ID%Type, 
	数次_IN				病人费用记录.数次%Type,
	标准单价_IN			病人费用记录.标准单价%Type,
    收入项目ID_IN		病人费用记录.收入项目ID%Type, 
    收据费目_IN			病人费用记录.收据费目%Type, 
    结算方式_IN			病人预交记录.结算方式%Type,--现金的结算名称
    应收金额_IN			病人费用记录.应收金额%Type, 
    实收金额_IN			病人费用记录.实收金额%Type, 
    病人科室ID_IN		病人费用记录.病人科室ID%Type, 
    开单部门ID_IN		病人费用记录.开单部门ID%Type, 
    执行部门ID_IN		病人费用记录.执行部门ID%Type, 
    操作员编号_IN		病人费用记录.操作员编号%Type, 
    操作员姓名_IN		病人费用记录.操作员姓名%Type, 
    发生时间_IN			病人费用记录.发生时间%Type, 
    登记时间_IN			病人费用记录.登记时间%Type, 
    医生姓名_IN			挂号安排.医生姓名%Type,
    医生ID_IN			挂号安排.医生ID%Type, 
    病历费_IN			Number,--该条记录是否病历工本费
    急诊_IN				Number,
    号别_IN				挂号安排.号码%Type, 
    诊室_IN				病人费用记录.发药窗口%Type, 
    结帐ID_IN			病人费用记录.结帐ID%Type, 
    领用ID_IN			票据使用明细.领用ID%Type,
    预交支付_IN			病人预交记录.冲预交%Type,--刷卡挂号时使用的预交金额,序号为1传入.
    现金支付_IN			病人预交记录.冲预交%Type,--挂号时现金支付部份金额,序号为1传入.
    个帐支付_IN			病人预交记录.冲预交%Type,--挂号时个人帐户支付金额,,序号为1传入.
    保险大类ID_IN		病人费用记录.保险大类ID%Type,
    保险项目否_IN		病人费用记录.保险项目否%Type,
    统筹金额_IN			病人费用记录.统筹金额%Type,
	摘要_IN				病人费用记录.摘要%Type,--预约挂号摘要信息
	预约挂号_IN			Number:=0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
	收费票据_IN			Number:=0, --挂号是否使用收费票据
    保险编码_IN         病人费用记录.保险编码%Type
) AS 
	--该游标用于收费冲预交的可用预交列表
    --以ID排序，优先冲上次未冲完的。 
    Cursor c_Deposit(v_病人ID 病人信息.病人ID%Type) is 
        Select * From( 
            Select A.ID,A.记录状态,A.NO,Nvl(A.金额,0) as 金额 
            From 病人预交记录 A,( 
                Select NO,Sum(Nvl(A.金额,0)) as 金额 
                From 病人预交记录 A 
                Where A.结帐ID is Null And Nvl(A.金额,0)<>0 
					And A.病人ID=v_病人ID 
				Group by NO Having Sum(Nvl(A.金额,0))<>0 
				) B
        Where A.结帐ID is Null And Nvl(A.金额,0)<>0 
			And A.结算方式 Not IN(Select 名称 From 结算方式 Where 性质=5)
			And A.NO=B.NO And A.病人ID=v_病人ID 
        Union All 
        Select 0 as ID,记录状态,NO,Sum(Nvl(金额,0)-Nvl(冲预交,0)) as 金额 
        From 病人预交记录 
        Where 记录性质 IN(1,11) And 结帐ID is Not NULL
			And Nvl(金额,0)<>Nvl(冲预交,0) And 病人ID=v_病人ID 
        Having Sum(Nvl(金额,0)-Nvl(冲预交,0))<>0 
        Group by 记录状态,NO) 
        Order by ID,NO; 
    --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录 
    --       同时汇总相关的汇总表(病人挂号汇总、病人费用汇总) 
    --       第一行费用处理票据使用情况(领用ID_IN>0) 
    
    v_打印ID    票据打印内容.ID%Type;
    v_费用ID    病人费用记录.ID%Type;
    v_预交金额    病人预交记录.金额%Type;

    v_现金        结算方式.名称%Type;
    v_个人帐户    结算方式.名称%Type;

    v_Count        Number;  
    v_Error        Varchar2(255); 
    Err_Custom    Exception; 
Begin 
	If Nvl(预约挂号_IN,0)=0 Then
		--因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
		Select Count(*) Into v_Count From 病人费用记录 Where 记录性质=4 And 记录状态 IN(1,3) And 序号=序号_IN And NO=单据号_IN;
		If v_Count<>0 THEN 
			v_Error:='挂号单据号重复,不能保存！'||CHR(13)||'如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
			Raise Err_Custom;
		End IF;

		--获取结算方式名称
		Begin
			Select 名称 Into v_现金 From 结算方式 Where 性质=1;
		Exception
			When Others Then v_现金:='现金';
		End;
		Begin
			Select 名称 Into v_个人帐户 From 结算方式 Where 性质=3;
		Exception
			When Others Then v_个人帐户:='个人帐户';
		End;
	End IF;

    --产生病人挂号费用(可能单独是或包括病历费用) 
    Select 病人费用记录_ID.Nextval Into v_费用ID From Dual; --应该通过程序得到
    Insert Into 病人费用记录( 
        ID,记录性质,记录状态,序号,价格父号,从属父号,NO,实际票号,门诊标志,加班标志,附加标志,发药窗口,病人ID,
		标识号, 床号,姓名,性别,年龄,费别,病人病区ID,病人科室ID,收费类别,计算单位,收费细目ID,收入项目ID, 
        收据费目,付数,数次,标准单价,应收金额,实收金额,结帐金额,结帐ID,记帐费用,开单部门ID,开单人,
		执行部门ID,执行人,操作员编号,操作员姓名,发生时间,登记时间,保险大类ID,保险项目否,保险编码,统筹金额,摘要) 
    Values( 
        v_费用ID,4,Decode(预约挂号_IN,1,0,1),序号_IN,Decode(价格父号_IN,0,Null,价格父号_IN),从属父号_IN,
		单据号_IN,票据号_IN,1,急诊_IN,病历费_IN,诊室_IN,Decode(病人ID_IN,0,Null,病人ID_IN), 
        Decode(门诊号_IN,0,Null,门诊号_IN),床号_IN,姓名_IN,Decode(姓名_IN,Null,Null,性别_IN), 
        Decode(姓名_IN,Null,Null,年龄_IN),费别_IN,病人科室ID_IN,病人科室ID_IN,收费类别_IN, 
        号别_IN,收费细目ID_IN,收入项目ID_IN,收据费目_IN,1,数次_IN,标准单价_IN,应收金额_IN,实收金额_IN, 
        Decode(预约挂号_IN,1,NULL,实收金额_IN),Decode(预约挂号_IN,1,NULL,结帐ID_IN),0,开单部门ID_IN,操作员姓名_IN,
		执行部门ID_IN,医生姓名_IN,操作员编号_IN, 操作员姓名_IN,发生时间_IN,登记时间_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,统筹金额_IN,摘要_IN); 
 
    --汇总结算到病人预交记录
	If Nvl(预约挂号_IN,0)=0 Then
		IF Nvl(现金支付_IN,0)<> 0 And 序号_IN=1 THEN     
			Insert Into 病人预交记录( 
				ID,记录性质,记录状态,NO,病人ID,结算方式,冲预交, 
				收款时间,操作员编号,操作员姓名,结帐ID,摘要) 
			Values( 
				病人预交记录_ID.Nextval,4,1,单据号_IN,Decode(病人ID_IN,0,Null,病人ID_IN), 
				Nvl(结算方式_IN,v_现金),现金支付_IN,登记时间_IN,操作员编号_IN,操作员姓名_IN,结帐ID_IN,'挂号收费'); 
		END IF;
		
		--对于医保挂号
		IF Nvl(个帐支付_IN,0)<> 0 And 序号_IN=1 THEN
			Insert Into 病人预交记录( 
				ID,记录性质,记录状态,NO,病人ID,结算方式,冲预交, 
				收款时间,操作员编号,操作员姓名,结帐ID,摘要) 
			Values( 
				病人预交记录_ID.Nextval,4,1,单据号_IN,Decode(病人ID_IN,0,Null,病人ID_IN), 
				v_个人帐户,个帐支付_IN,登记时间_IN,操作员编号_IN,操作员姓名_IN,结帐ID_IN,'医保挂号');
		END IF;
	  
		--对于就诊卡通过预交金挂号 
		IF Nvl(预交支付_IN,0)<> 0 And 序号_IN=1 THEN
			v_预交金额:=预交支付_IN; 
			For r_Deposit IN c_Deposit(病人ID_IN) Loop 
				IF r_Deposit.ID <> 0 Then 
					--第一次冲预交 
					Update 病人预交记录 
						Set 冲预交=Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额), 
							结帐ID=结帐ID_IN 
					Where ID=r_Deposit.ID; 
				Else 
					--冲上次剩余额 
					INSERT Into 病人预交记录( 
						ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额, 
						结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号,收款时间, 
						操作员姓名,操作员编号,冲预交,结帐ID) 
					Select 病人预交记录_ID.Nextval,NO,实际票号,11,记录状态,病人ID, 
						主页ID,科室ID,NULL,结算方式,结算号码,摘要,缴款单位, 
						单位开户行,单位帐号,收款时间,操作员姓名,操作员编号, 
						Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额),结帐ID_IN 
					From 病人预交记录 
					Where NO=r_Deposit.NO And 记录状态=r_Deposit.记录状态 
						And 记录性质 IN(1,11) And RowNum=1; 
				End IF; 

				--检查是否已经处理完 
				IF r_Deposit.金额<v_预交金额 Then 
					v_预交金额:=v_预交金额-r_Deposit.金额; 
				Else 
					v_预交金额:=0; 
				End IF; 

				IF v_预交金额=0 Then 
					Exit; 
				End IF; 
			End Loop; 
		
			--更新病人预交余额 
			Update 病人余额 Set 预交余额=Nvl(预交余额,0)-预交支付_IN Where 病人ID=病人ID_IN And 性质=1; 
			Delete From 病人余额 Where 病人ID=病人ID_IN And 性质=1 And Nvl(费用余额,0)=0 And Nvl(预交余额,0)=0; 
		End IF; 
		
		--相关汇总表的处理 
		--人员缴款余额 
		IF 序号_IN=1 THEN 
			IF Nvl(现金支付_IN,0)<> 0 THEN
				Update 人员缴款余额 Set 余额=Nvl(余额,0)+现金支付_IN 
				Where 性质=1 And 收款员=操作员姓名_IN And 结算方式=Nvl(结算方式_IN,v_现金); 

				IF SQL%RowCount=0 Then 
					Insert Into 人员缴款余额( 
						收款员,结算方式,性质,余额) 
					Values( 
						操作员姓名_IN,Nvl(结算方式_IN,v_现金),1,现金支付_IN); 
				End If; 
			END IF;

			IF Nvl(个帐支付_IN,0)<> 0 THEN
				Update 人员缴款余额 Set 余额=Nvl(余额,0)+个帐支付_IN 
				Where 性质=1 And 收款员=操作员姓名_IN And 结算方式=v_个人帐户; 

				IF SQL%RowCount=0 Then 
					Insert Into 人员缴款余额( 
						收款员,结算方式,性质,余额) 
					Values( 
						操作员姓名_IN,v_个人帐户,1,个帐支付_IN); 
				End If; 
			END IF;
			Delete From 人员缴款余额 Where 收款员=操作员姓名_IN And 性质=1 And Nvl(余额,0)=0; 
		END if;
	End IF;

    --病人挂号汇总(只处理一次,且单独收取病历费不处理) 
    IF 号别_IN is Not Null And 序号_IN=1 Then 
		If Nvl(预约挂号_IN,0)=0 Then
			Update 病人挂号汇总 
				Set 已挂数=Nvl(已挂数,0)+1 
			Where 日期=Trunc(登记时间_IN) 
				And Nvl(科室ID,0)=执行部门ID_IN 
				And Nvl(项目ID,0)=收费细目ID_IN 
				And Nvl(医生姓名,'医生')=Nvl(医生姓名_IN,'医生') 
				And Nvl(医生ID,0)=Nvl(医生ID_IN,0); 

			IF SQL%RowCount=0 Then 
				Insert Into 病人挂号汇总( 
					日期,科室ID,项目ID,医生姓名,医生ID,已挂数) 
				Values( 
					Trunc(登记时间_IN),执行部门ID_IN,收费细目ID_IN,医生姓名_IN,Decode(医生ID_IN,0,Null,医生ID_IN),1); 
			End If; 
		Else
			Update 病人挂号汇总 
				Set 已约数=Nvl(已约数,0)+1 
			Where 日期=Trunc(发生时间_IN) 
				And Nvl(科室ID,0)=执行部门ID_IN 
				And Nvl(项目ID,0)=收费细目ID_IN 
				And Nvl(医生姓名,'医生')=Nvl(医生姓名_IN,'医生') 
				And Nvl(医生ID,0)=Nvl(医生ID_IN,0); 

			IF SQL%RowCount=0 Then 
				Insert Into 病人挂号汇总( 
					日期,科室ID,项目ID,医生姓名,医生ID,已约数) 
				Values( 
					Trunc(发生时间_IN),执行部门ID_IN,收费细目ID_IN,医生姓名_IN,Decode(医生ID_IN,0,Null,医生ID_IN),1); 
			End If; 
		End IF;
    End If; 
    
    --病人费用汇总 
	If Nvl(预约挂号_IN,0)=0 Then
		Update 病人费用汇总 
			Set 应收金额=Nvl(应收金额,0)+应收金额_IN, 
				实收金额=Nvl(实收金额,0)+实收金额_IN, 
				结帐金额=Nvl(结帐金额,0)+实收金额_IN 
			Where 日期=Trunc(登记时间_IN) 
				And Nvl(病人病区ID,0)=病人科室ID_IN 
				And Nvl(病人科室ID,0)=病人科室ID_IN 
				And Nvl(开单部门ID,0)=开单部门ID_IN 
				And Nvl(执行部门ID,0)=执行部门ID_IN 
				And 收入项目ID+0=收入项目ID_IN 
				And 来源途径=1 And 记帐费用=0; 
		IF SQL%RowCount=0 Then 
			Insert Into 病人费用汇总( 
				日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID, 
				收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额) 
			Values( 
				Trunc(登记时间_IN),病人科室ID_IN,病人科室ID_IN,开单部门ID_IN, 
				执行部门ID_IN,收入项目ID_IN,1,0,应收金额_IN,实收金额_IN,实收金额_IN); 
		End If; 
		
		--处理票据使用情况
		IF 序号_IN=1 And 票据号_IN is Not Null Then 
			Select 票据打印内容_ID.Nextval Into v_打印ID From Dual;

			--发出票据 
			Insert Into 票据打印内容(
				ID,数据性质,NO)
			Values(
				v_打印ID,4,单据号_IN);

			Insert Into 票据使用明细( 
				ID,票种,号码,性质,原因,领用ID,打印ID,使用时间,使用人) 
			Values( 
				票据使用明细_ID.Nextval,Decode(收费票据_IN,1,1,4),票据号_IN,1,1,领用ID_IN,v_打印ID,登记时间_IN,操作员姓名_IN); 

			--状态改动 
			Update 票据领用记录 
				Set 当前号码=票据号_IN,剩余数量=Decode(SIGN(剩余数量-1),-1,0,剩余数量-1) 
			Where ID=Nvl(领用ID_IN,0); 
		End If; 
	 
		--病人本次就诊(以发生时间为准) 
		IF Nvl(病人ID_IN,0)<>0 And 序号_IN=1 Then 
			Update 病人信息 
				Set 就诊时间=登记时间_IN, 
					就诊状态=1, 
					就诊诊室=诊室_IN 
			Where 病人ID=病人ID_IN; 
		End If; 
	End IF;

	--病人挂号记录
	IF 号别_IN is Not Null And 序号_IN=1 And Nvl(预约挂号_IN,0)=0 Then 
		Insert Into 病人挂号记录(
			ID,NO,病人ID,门诊号,姓名,性别,年龄,号别,急诊,诊室,附加标志,
			执行部门ID,执行人,执行状态,执行时间,登记时间,操作员编号,操作员姓名,摘要)
		Values(
			病人挂号记录_ID.Nextval,单据号_IN,病人ID_IN,门诊号_IN,姓名_IN,性别_IN,
			年龄_IN,号别_IN,急诊_IN,诊室_IN,NULL,执行部门ID_IN,医生姓名_IN,0,NULL,
			登记时间_IN,操作员编号_IN,操作员姓名_IN,摘要_IN);
	End IF;
Exception 
    When Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]'); 
    When Others Then zl_ErrOrCenter(SQLCODE,SQLERRM); 
End zl_病人挂号记录_Insert; 
/

CREATE OR REPLACE Procedure ZL_预约挂号接收_INSERT(
    NO_IN				病人费用记录.NO%Type, 
    票据号_IN			病人费用记录.实际票号%Type, 
    领用ID_IN			票据使用明细.领用ID%Type,
    结帐ID_IN			病人费用记录.结帐ID%Type, 
    诊室_IN				病人费用记录.发药窗口%Type, 
	病人ID_IN			病人费用记录.病人ID%Type, 
    门诊号_IN			病人费用记录.标识号%Type, 
    姓名_IN				病人费用记录.姓名%Type, 
    性别_IN				病人费用记录.性别%Type, 
    年龄_IN				病人费用记录.年龄%Type, 
	床号_IN				病人费用记录.床号%Type,--用于存放病人的医疗付款方式编号
    费别_IN				病人费用记录.费别%Type, 
    结算方式_IN			病人预交记录.结算方式%Type,--现金的结算名称
    现金支付_IN			病人预交记录.冲预交%Type,--挂号时现金支付部份金额
	预交支付_IN			病人预交记录.冲预交%Type,--挂号时使用的预交金额
    个帐支付_IN			病人预交记录.冲预交%Type,--挂号时个人帐户支付金额
    发生时间_IN			病人费用记录.发生时间%Type 
) AS 
	--该游标用于收费冲预交的可用预交列表
    --以ID排序，优先冲上次未冲完的。 
    Cursor c_Deposit(v_病人ID 病人信息.病人ID%Type) is 
        Select * From( 
            Select A.ID,A.记录状态,A.NO,Nvl(A.金额,0) as 金额 
            From 病人预交记录 A,( 
                Select NO,Sum(Nvl(A.金额,0)) as 金额 
                From 病人预交记录 A 
                Where A.结帐ID is Null And Nvl(A.金额,0)<>0 
					And A.病人ID=v_病人ID 
				Group by NO Having Sum(Nvl(A.金额,0))<>0 
				) B
        Where A.结帐ID is Null And Nvl(A.金额,0)<>0 
			And A.结算方式 Not IN(Select 名称 From 结算方式 Where 性质=5)
			And A.NO=B.NO And A.病人ID=v_病人ID 
        Union All 
        Select 0 as ID,记录状态,NO,Sum(Nvl(金额,0)-Nvl(冲预交,0)) as 金额 
        From 病人预交记录 
        Where 记录性质 IN(1,11) And 结帐ID is Not NULL
			And Nvl(金额,0)<>Nvl(冲预交,0) And 病人ID=v_病人ID 
        Having Sum(Nvl(金额,0)-Nvl(冲预交,0))<>0 
        Group by 记录状态,NO) 
        Order by ID,NO; 

	--处理病人费用汇总
	Cursor c_Money is
		Select 病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,
			Sum(应收金额) AS 应收金额,Sum(实收金额) AS 实收金额,Sum(结帐金额) AS 结帐金额
		From 病人费用记录
		Where 记录性质=4 And 记录状态=1 And NO=NO_IN
		Group By 病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID;

	--号别信息
	Cursor c_Regist is
		Select B.科室ID,B.项目ID,B.医生ID,B.医生姓名
		From 病人费用记录 A,挂号安排 B
		Where A.记录性质=4 And A.记录状态=1 And A.NO=NO_IN
			And A.序号=1 And A.计算单位=B.号码;
    r_Regist c_Regist%RowType;	

    v_现金			结算方式.名称%Type;
    v_个人帐户		结算方式.名称%Type;
    v_打印ID		票据打印内容.ID%Type;
	v_预交金额		病人预交记录.金额%Type;
	
    v_Temp			Varchar2(255);
    v_人员编号		病人费用记录.操作员编号%Type;
    v_人员姓名		病人费用记录.操作员姓名%Type;

	v_Date			Date;
    v_Error			Varchar2(255); 
    Err_Custom		Exception; 
Begin 
	--获取结算方式名称
	Begin
		Select 名称 Into v_现金 From 结算方式 Where 性质=1;
	Exception
		When Others Then v_现金:='现金';
	End;
	Begin
		Select 名称 Into v_个人帐户 From 结算方式 Where 性质=3;
	Exception
		When Others Then v_个人帐户:='个人帐户';
	End;
	Select Sysdate Into v_Date From Dual;
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

	--更新病人费用记录
	Update 病人费用记录
		Set 记录状态=1,
			实际票号=票据号_IN,
			结帐ID=结帐ID_IN,
			结帐金额=实收金额,
			发药窗口=诊室_IN,
			病人ID=病人ID_IN,
			标识号=门诊号_IN,
			姓名=姓名_IN,
			年龄=年龄_IN,
			性别=性别_IN,
			床号=床号_IN,
			费别=费别_IN,
			发生时间=发生时间_IN,
			登记时间=v_Date,
			操作员编号=v_人员编号,
			操作员姓名=v_人员姓名
	Where 记录性质=4 And 记录状态=0 And NO=NO_IN;
	 
	--病人挂号记录
	Insert Into 病人挂号记录(
		ID,NO,病人ID,门诊号,姓名,性别,年龄,号别,急诊,诊室,附加标志,
		执行部门ID,执行人,执行状态,执行时间,登记时间,操作员编号,操作员姓名,摘要)
	Select
		病人挂号记录_ID.Nextval,NO_IN,病人ID_IN,门诊号_IN,姓名_IN,性别_IN,
		年龄_IN,计算单位,加班标志,诊室_IN,NULL,执行部门ID,执行人,0,NULL,
		v_Date,v_人员编号,v_人员姓名,Nvl(摘要,结论)
	From 病人费用记录
	Where 记录性质=4 And 记录状态=1 And 序号=1 And NO=NO_IN;

    --汇总结算到病人预交记录
	IF Nvl(现金支付_IN,0)<> 0  THEN     
		Insert Into 病人预交记录( 
			ID,记录性质,记录状态,NO,病人ID,结算方式,冲预交, 
			收款时间,操作员编号,操作员姓名,结帐ID,摘要) 
		Values( 
			病人预交记录_ID.Nextval,4,1,NO_IN,病人ID_IN, 
			Nvl(结算方式_IN,v_现金),现金支付_IN,v_Date,v_人员编号,v_人员姓名,结帐ID_IN,'挂号收费'); 
	END IF;
	
	--对于就诊卡通过预交金挂号 
	IF Nvl(预交支付_IN,0)<> 0 THEN
		v_预交金额:=预交支付_IN; 
		For r_Deposit IN c_Deposit(病人ID_IN) Loop 
			IF r_Deposit.ID <> 0 Then 
				--第一次冲预交 
				Update 病人预交记录 
					Set 冲预交=Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额), 
						结帐ID=结帐ID_IN 
				Where ID=r_Deposit.ID; 
			Else 
				--冲上次剩余额 
				INSERT Into 病人预交记录( 
					ID,NO,实际票号,记录性质,记录状态,病人ID,主页ID,科室ID,金额, 
					结算方式,结算号码,摘要,缴款单位,单位开户行,单位帐号,收款时间, 
					操作员姓名,操作员编号,冲预交,结帐ID) 
				Select 病人预交记录_ID.Nextval,NO,实际票号,11,记录状态,病人ID, 
					主页ID,科室ID,NULL,结算方式,结算号码,摘要,缴款单位, 
					单位开户行,单位帐号,收款时间,操作员姓名,操作员编号, 
					Decode(Sign(r_Deposit.金额-v_预交金额),-1,r_Deposit.金额,v_预交金额),结帐ID_IN 
				From 病人预交记录 
				Where NO=r_Deposit.NO And 记录状态=r_Deposit.记录状态 
					And 记录性质 IN(1,11) And RowNum=1; 
			End IF; 

			--检查是否已经处理完 
			IF r_Deposit.金额<v_预交金额 Then 
				v_预交金额:=v_预交金额-r_Deposit.金额; 
			Else 
				v_预交金额:=0; 
			End IF; 

			IF v_预交金额=0 Then 
				Exit; 
			End IF; 
		End Loop; 
	
		--更新病人预交余额 
		Update 病人余额 Set 预交余额=Nvl(预交余额,0)-预交支付_IN Where 病人ID=病人ID_IN And 性质=1; 
		Delete From 病人余额 Where 病人ID=病人ID_IN And 性质=1 And Nvl(费用余额,0)=0 And Nvl(预交余额,0)=0; 
	End IF; 

	--对于医保挂号
	IF Nvl(个帐支付_IN,0)<> 0 THEN
		Insert Into 病人预交记录( 
			ID,记录性质,记录状态,NO,病人ID,结算方式,冲预交, 
			收款时间,操作员编号,操作员姓名,结帐ID,摘要) 
		Values( 
			病人预交记录_ID.Nextval,4,1,NO_IN,病人ID_IN, v_个人帐户,个帐支付_IN,
			v_Date,v_人员编号,v_人员姓名,结帐ID_IN,'医保挂号');
	END IF;
  
	--相关汇总表的处理 
	--人员缴款余额 
	IF Nvl(现金支付_IN,0)<> 0 THEN
		Update 人员缴款余额 
			Set 余额=Nvl(余额,0)+现金支付_IN 
		Where 性质=1 And 收款员=v_人员姓名 And 结算方式=Nvl(结算方式_IN,v_现金); 

		IF SQL%RowCount=0 Then 
			Insert Into 人员缴款余额( 
				收款员,结算方式,性质,余额) 
			Values( 
				v_人员姓名,Nvl(结算方式_IN,v_现金),1,现金支付_IN); 
		End If; 
	END IF;

	IF Nvl(个帐支付_IN,0)<> 0 THEN
		Update 人员缴款余额 Set 余额=Nvl(余额,0)+个帐支付_IN 
		Where 性质=1 And 收款员=v_人员姓名 And 结算方式=v_个人帐户; 

		IF SQL%RowCount=0 Then 
			Insert Into 人员缴款余额( 
				收款员,结算方式,性质,余额) 
			Values( 
				v_人员姓名,v_个人帐户,1,个帐支付_IN); 
		End If; 
	END IF;
	Delete From 人员缴款余额 Where 收款员=v_人员姓名 And 性质=1 And Nvl(余额,0)=0; 

    --病人挂号汇总(只处理一次,且单独收取病历费不处理) 
	Open c_Regist;
	Fetch c_Regist Into r_Regist;
	Update 病人挂号汇总 
		Set 已挂数=Nvl(已挂数,0)+1 
	Where 日期=Trunc(v_Date) 
		And Nvl(科室ID,0)=Nvl(r_Regist.科室ID,0)
		And Nvl(项目ID,0)=Nvl(r_Regist.项目ID,0)
		And Nvl(医生姓名,'医生')=Nvl(r_Regist.医生姓名,'医生') 
		And Nvl(医生ID,0)=Nvl(r_Regist.医生ID,0); 
	IF SQL%RowCount=0 Then 
		Insert Into 病人挂号汇总( 
			日期,科室ID,项目ID,医生姓名,医生ID,已挂数) 
		Values(
			Trunc(v_Date),r_Regist.科室ID,r_Regist.项目ID,r_Regist.医生姓名,r_Regist.医生ID,1); 
	End If;
	Close c_Regist;
    
    --病人费用汇总 
	For r_Money In c_Money Loop
		Update 病人费用汇总 
			Set 应收金额=Nvl(应收金额,0)+Nvl(r_Money.应收金额,0), 
				实收金额=Nvl(实收金额,0)+Nvl(r_Money.实收金额,0), 
				结帐金额=Nvl(结帐金额,0)+Nvl(r_Money.结帐金额,0) 
			Where 日期=Trunc(v_Date) 
				And Nvl(病人病区ID,0)=Nvl(r_Money.病人病区ID,0)
				And Nvl(病人科室ID,0)=Nvl(r_Money.病人科室ID,0)
				And Nvl(开单部门ID,0)=Nvl(r_Money.开单部门ID,0)
				And Nvl(执行部门ID,0)=Nvl(r_Money.执行部门ID,0)
				And 收入项目ID+0=r_Money.收入项目ID
				And 来源途径=1 And 记帐费用=0;
		IF SQL%RowCount=0 Then 
			Insert Into 病人费用汇总( 
				日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID, 
				收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额) 
			Values( 
				Trunc(v_Date),r_Money.病人病区ID,r_Money.病人科室ID,r_Money.开单部门ID, 
				r_Money.执行部门ID,r_Money.收入项目ID,1,0,r_Money.应收金额,r_Money.实收金额,r_Money.结帐金额); 
		End If; 
	End Loop;

	--处理票据使用情况
	IF 票据号_IN is Not Null Then 
		Select 票据打印内容_ID.Nextval Into v_打印ID From Dual;

		--发出票据 
		Insert Into 票据打印内容(
			ID,数据性质,NO)
		Values(
			v_打印ID,4,NO_IN);

		Insert Into 票据使用明细( 
			ID,票种,号码,性质,原因,领用ID,打印ID,使用时间,使用人) 
		Values( 
			票据使用明细_ID.Nextval,4,票据号_IN,1,1,领用ID_IN,v_打印ID,v_Date,v_人员姓名); 

		--状态改动 
		Update 票据领用记录 
			Set 当前号码=票据号_IN,剩余数量=Decode(SIGN(剩余数量-1),-1,0,剩余数量-1) 
		Where ID=Nvl(领用ID_IN,0); 
	End If; 
 
	--病人本次就诊(以发生时间为准) 
	IF Nvl(病人ID_IN,0)<>0 Then 
		Update 病人信息 
			Set 就诊时间=v_Date, 
				就诊状态=1, 
				就诊诊室=诊室_IN 
		Where 病人ID=病人ID_IN; 
	End If; 
Exception 
    When Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]'); 
    When Others Then zl_ErrOrCenter(SQLCODE,SQLERRM); 
End ZL_预约挂号接收_INSERT; 
/

Create Or Replace Procedure ZL_病人挂号记录_换号(
--功能：完成病人换号功能，在挂号项目ID相同的情况下。
    NO_IN			病人挂号记录.NO%Type,
    号别_IN			病人挂号记录.号别%Type,
    诊室_IN			病人挂号记录.诊室%Type,
    科室ID_IN		病人挂号记录.执行部门ID%Type,
    原医生_IN		病人挂号记录.执行人%Type,
    原医生ID_IN		病人挂号汇总.医生ID%Type,
    新医生_IN		病人挂号记录.执行人%Type,
    新医生ID_IN		病人挂号汇总.医生ID%Type
) AS
    Cursor c_Bill  IS
	    Select * From 病人费用记录 Where 记录性质=4 And 记录状态=1 And NO=NO_IN Order BY 序号;

	v_病人ID	病人费用记录.ID%Type;
    v_Error		Varchar2(255);
    Err_Custom	Exception;
Begin
	v_病人ID:=0;
    Begin
      Select 病人ID Into v_病人ID From 病人挂号记录 Where NO=NO_IN;
    Exception
        When OTHERS Then Null;
    End;
	If v_病人ID=0 Then
        v_Error:='没有找到病人的挂号信息。';
        Raise Err_Custom;
    ElsIf v_病人ID IS Null Then 
        v_Error:='没有找到病人信息。';
        Raise Err_Custom;
    End If;

    ---先更新病人信息的就诊诊室和状态
    Update 病人信息 Set 就诊诊室=诊室_IN,就诊状态=1 Where 病人ID=v_病人ID And 就诊状态 IN(1,2);
     
    For r_Bill IN c_Bill  Loop 
		--恢复以前的挂号汇总
		Update 病人挂号汇总 
			Set 已挂数=Nvl(已挂数,0)-1 
		Where 日期=Trunc(r_Bill.登记时间)
			And Nvl(科室ID,0)=Nvl(r_Bill.执行部门ID,0)
			And Nvl(项目ID,0)=Nvl(r_Bill.收费细目ID,0)
			And Nvl(医生姓名,'医生')=Nvl(原医生_IN,'医生') 
			And Nvl(医生ID,0)=Nvl(原医生ID_IN,0); 
		If SQL%RowCount=0 Then 
			Insert Into 病人挂号汇总(
				日期,科室ID,项目ID,医生姓名,医生ID,已挂数) 
			Values(
				Trunc(r_Bill.登记时间),r_Bill.执行部门ID,r_Bill.收费细目ID,原医生_IN,Decode(原医生ID_IN,0,Null,原医生ID_IN),-1); 
		End If;     
			
		--恢复以前的费用汇总
		Update 病人费用汇总 
			Set 应收金额=Nvl(应收金额,0)-Nvl(r_Bill.应收金额,0), 
				实收金额=Nvl(实收金额,0)-Nvl(r_Bill.实收金额,0), 
				结帐金额=Nvl(结帐金额,0)-Nvl(r_Bill.结帐金额,0) 
		Where 日期=Trunc(r_Bill.登记时间)
			And Nvl(病人病区ID,0)=Nvl(r_Bill.病人科室ID,0)
			And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
			And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
			And Nvl(执行部门ID,0)=Nvl(r_Bill.执行部门ID,0)
			And 收入项目ID+0=r_Bill.收入项目ID
			And 来源途径=1 And 记帐费用=0; 
		If SQL%RowCount=0 Then 
			Insert Into 病人费用汇总( 
				日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID, 
				收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额) 
			Values( 
				Trunc(r_Bill.登记时间),r_Bill.病人科室ID,r_Bill.病人科室ID,r_Bill.开单部门ID, 
				r_Bill.执行部门ID,r_Bill.收入项目ID,1,0,-1*r_Bill.应收金额,-1*r_Bill.实收金额,-1*r_Bill.结帐金额); 
		End If; 

		----然后再更新挂号汇总
		Update 病人挂号汇总 
			Set 已挂数=Nvl(已挂数,0)+1 
		Where 日期=Trunc(r_Bill.登记时间)
			And Nvl(科室ID,0)=科室ID_IN
			And Nvl(项目ID,0)=Nvl(r_Bill.收费细目ID,0)
			And Nvl(医生姓名,'医生')=Nvl(新医生_IN,'医生') 
			And Nvl(医生ID,0)=Nvl(新医生ID_IN,0); 
		If SQL%RowCount=0 Then 
			Insert Into 病人挂号汇总(
				日期,科室ID,项目ID,医生姓名,医生ID,已挂数) 
			Values(
				Trunc(r_Bill.登记时间),科室ID_IN,r_Bill.收费细目ID,新医生_IN,Decode(新医生ID_IN,0,Null,新医生ID_IN),1); 
		End If;     
			
		-----然后再更新挂号费用汇总
		Update 病人费用汇总 
			Set 应收金额=Nvl(应收金额,0)+Nvl(r_Bill.应收金额,0), 
				实收金额=Nvl(实收金额,0)+Nvl(r_Bill.实收金额,0), 
				结帐金额=Nvl(结帐金额,0)+Nvl(r_Bill.结帐金额,0) 
		Where 日期=Trunc(r_Bill.登记时间)
			And Nvl(病人病区ID,0)=Nvl(r_Bill.病人科室ID,0)
			And Nvl(病人科室ID,0)=Nvl(r_Bill.病人科室ID,0)
			And Nvl(开单部门ID,0)=Nvl(r_Bill.开单部门ID,0)
			And Nvl(执行部门ID,0)=科室ID_IN
			And 收入项目ID+0=r_Bill.收入项目ID
			And 来源途径=1 And 记帐费用=0; 
		If SQL%RowCount=0 Then 
			Insert Into 病人费用汇总( 
				日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID, 
				收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额) 
			Values( 
				Trunc(r_Bill.登记时间),r_Bill.病人科室ID,r_Bill.病人科室ID,r_Bill.开单部门ID, 
				科室ID_IN,r_Bill.收入项目ID,1,0,r_Bill.应收金额,r_Bill.实收金额,r_Bill.结帐金额); 
		End If; 
		  
		---更新挂号记录
		Update 病人费用记录
			Set 执行部门ID=科室ID_IN,
				病人科室ID=科室ID_IN,
				病人病区ID=科室ID_IN,
				计算单位=号别_IN,
				发药窗口=诊室_IN,
				执行人=新医生_IN,
				执行状态=0,执行时间=Null
		Where ID=r_Bill.ID;

		--更新病人挂号记录
		If r_Bill.序号=1 Then
			Update 病人挂号记录
				Set 执行部门ID=科室ID_IN,
					号别=号别_IN,
					诊室=诊室_IN,
					执行人=新医生_IN,
					执行状态=0,
					执行时间=Null
			Where NO=r_Bill.NO;
		End If;
    End Loop;
Exception
    When Err_Custom Then RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When OTHERS Then ZL_ERRORCENTER (SQLCODE, SQLERRM);
End ZL_病人挂号记录_换号;
/

CREATE OR REPLACE Procedure ZL_病人医嘱记录_收回(
--功能：将指定医嘱超期发送部分收回。如果上次发送没有产生费用，则仅收回医嘱的上次执行时间。
--参数：NO_IN=用于收回产生冲销单据的新单据号(供费用及药品使用),当前处理的只是新NO的一部份。
--            因为药品可能分批,所以序号在处理时取。
--      收回量_IN=对药品为按住院单位的收回量,对其它医嘱为收回总量或次数。
--      医嘱ID_IN=每条要收回的医嘱记录的ID(明细存储的ID),对成药或配方,不一定包含给药途径或用法煎法(可能为叮嘱而未读取)
--      上次时间_IN=医嘱超期发送部分收回后应该还原的上次执行时间(严格按频率计算得来),为空时表示被全部收回了。
    NO_IN				病人费用记录.NO%Type,
    收回量_IN			病人医嘱发送.发送数次%Type,
    医嘱ID_IN			病人医嘱记录.ID%Type,
    诊疗类别_IN			病人医嘱记录.诊疗类别%Type,
    上次时间_IN			病人医嘱记录.上次执行时间%Type
) IS
    --包含指定成药长嘱发送时产生的相关费用及药品记录信息(因多次发送或分批而有多条记录)
    --药品医嘱填写了"病人医嘱发送"记录,对应的给药途径不一定填写了的(可能为叮嘱),且NO不同。
    --因为要收回的次数可能包含了多次发送的内容,所以要将多次发送的收发记录都取出来
    Cursor c_Drug is
        Select A.病人ID,A.主页ID,D.姓名,
            X.住院包装,X.最大效期,Nvl(B.付数,1)*B.实际数量 AS 数量,
            B.ID AS 收发ID,B.单据,B.药品ID,B.对方部门ID,B.库房ID,B.费用ID,
            Nvl(X.药房分批,0) AS 分批,B.批次,B.批号,B.效期
        From 病人费用记录 A,药品收发记录 B,病人医嘱发送 C,病人信息 D,药品规格 X
        Where C.医嘱ID=医嘱ID_IN And A.NO=C.NO And A.记录性质=C.记录性质 And A.记录状态 IN(0,1,3)
            And A.医嘱序号+0=医嘱ID_IN And A.NO=B.NO And A.ID=B.费用ID+0
            And B.单据 IN(9,10) And (B.记录状态=1 Or Mod(B.记录状态,3)=0)
            And A.病人ID=D.病人ID And B.药品ID=X.药品ID
        Order BY B.NO Desc,B.ID Desc;
    
    --包含非药长嘱(含给药途径)发送时所产生的费用(因多个收入而有多条记录)
    --对非药医嘱,直接收回指定量,不管多次发送(如果多次发送价格不同,则收回的价格是以最后次的；不然就要根据多个收入依次减收回量)。
    --非药长嘱应该都填写了发送记录(除开了叮嘱及护理等级)
    Cursor c_Other is
        Select A.ID AS 费用ID,Nvl(A.付数,1)*A.数次 AS 数量
        From 病人费用记录 A,病人医嘱发送 B
        Where A.NO=B.NO And A.记录性质=B.记录性质 And A.记录状态 IN(0,1,3)
            And A.医嘱序号+0=医嘱ID_IN And B.医嘱ID=医嘱ID_IN
            And B.发送号=(Select Max(发送号) From 病人医嘱发送 Where 医嘱ID=医嘱ID_IN)
        Order BY A.收费细目ID,A.序号;

    --该游标用于处理费用相关汇总表
    Cursor c_Money(v_Start 病人费用记录.序号%Type,v_End 病人费用记录.序号%Type) IS
        Select 病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,
            Sum(Nvl(应收金额,0)) AS 应收金额,Sum(Nvl(实收金额,0)) AS 实收金额
        From 病人费用记录 
        Where 记录性质=2 And 记录状态=1 And NO=NO_IN And 序号 Between v_Start And v_End
        Group BY 病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID;

    v_Date          Date;
    v_费用序号      病人费用记录.序号%Type;
    v_收发序号      药品收发记录.序号%Type;
    v_费用ID        病人费用记录.ID%Type;
    v_批次          药品收发记录.批次%Type;
    v_效期          药品收发记录.效期%Type;
    v_批号          药品收发记录.批号%Type;
    v_优先级        身份.优先级%Type;
    v_实收金额      病人费用记录.实收金额%Type;
    
    v_开始序号      病人费用记录.序号%Type;
    v_结束序号      病人费用记录.序号%Type;

    v_剩余量        药品收发记录.实际数量%Type;
    v_当前量        药品收发记录.实际数量%Type;
    
	v_组ID			病人医嘱记录.ID%Type;

    v_Temp          Varchar2(255);
    v_人员编号      病人费用记录.操作员编号%Type;
    v_人员姓名      病人费用记录.操作员姓名%Type;

	v_药品划价		Number;
	v_其他划价		Number;

    v_Dec			Number;
    v_Error         Varchar2(255);
    Err_Custom      Exception;
Begin
    --金额小数位数
    Begin
        Select To_Number(参数值) Into v_Dec From 系统参数表 Where 参数号=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --取操作员信息(部门ID,部门名称;人员ID,人员编号,人员姓名)
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

	--生成划价单系统参数
	Begin
		Select To_Number(Nvl(参数值,'0')) Into v_药品划价 From 系统参数表 Where 参数号=79;
	Exception
		When Others Then v_药品划价:=0;
	End;
	Begin
		Select To_Number(Nvl(参数值,'0')) Into v_其他划价 From 系统参数表 Where 参数号=80;
	Exception
		When Others Then v_其他划价:=0;
	End;

    Select Sysdate Into v_Date From Dual;
    v_开始序号:=NULL;v_结束序号:=NULL;

    If Nvl(收回量_IN,0)<>0 Then
        If 诊疗类别_IN IN('5','6') Then
            --中，西成药
            -----------------------------------------------------------------------------------------------------
            v_剩余量:=NULL;
            Select Nvl(Max(序号),0)+1 Into v_收发序号 From 药品收发记录 Where 单据 IN(9,10) And 记录状态=1 And NO=NO_IN;
            Select Nvl(Max(序号),0)+1 Into v_费用序号 From 病人费用记录 Where 记录性质=2 And 记录状态 IN(0,1) And NO=NO_IN;
            For r_Drug In c_Drug Loop
                --初始化要收回的总数量(零售数量)
                IF v_剩余量 IS NULL Then
                    v_剩余量:=Round(收回量_IN*r_Drug.住院包装,5);
                End IF;
                
                If v_剩余量>=r_Drug.数量 Then
                    v_当前量:=r_Drug.数量;
                Else
                    v_当前量:=v_剩余量;
                End IF;
                v_剩余量:=v_剩余量-v_当前量;

                Select 病人费用记录_ID.Nextval Into v_费用ID From Dual;
                
                --确定批次
                If Nvl(r_Drug.批次,0)<>0 And r_Drug.分批=0 Then
                    --原分批,现不分批
                    v_批次:=NULL;
                    v_批号:=r_Drug.批号;
                    v_效期:=r_Drug.效期;
                ElsIf Nvl(r_Drug.批次,0)=0 And r_Drug.分批=1 Then
                    --原不分批,现分批
                    Select 药品收发记录_ID.Nextval Into v_批次 From Dual;
                    Select To_Char(Sysdate,'YYYYMMDD') Into v_批号 From Dual;
                    If r_Drug.最大效期 is Not Null Then
                        v_效期:=Trunc(Sysdate+r_Drug.最大效期*30);
                    Else
                        v_效期:=NULL;
                    End IF;
                Else
                    v_批次:=r_Drug.批次;
                    v_批号:=r_Drug.批号;
                    v_效期:=r_Drug.效期;
                End IF;

                Insert Into 药品收发记录(
                    ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
                    药品ID,批次,产地,批号,效期,付数,填写数量,实际数量,零售价,零售金额,
                    摘要,填制人,填制日期,费用ID,单量,频次,用法)
                Select
                    药品收发记录_ID.Nextval,1,单据,NO_IN,v_收发序号,库房ID,对方部门ID,
                    入出类别ID,-1,药品ID,v_批次,产地,v_批号,v_效期,1,-1*v_当前量,-1*v_当前量,
                    零售价,Round(-1*v_当前量*零售价,v_Dec),'超期发送收回',v_人员姓名,v_Date,v_费用ID,
                    单量,频次,用法
                From 药品收发记录 Where ID=r_Drug.收发ID;

                --药品库存
                Update 药品库存
                    Set 可用数量=Nvl(可用数量,0)-(-1*v_当前量)
                Where 库房ID=r_Drug.库房ID And 药品ID=r_Drug.药品ID
                    And Nvl(批次,0)=Nvl(v_批次,0) And 性质=1;
                IF SQL%RowCount=0 Then
                    Insert Into 药品库存(
                        库房ID,药品ID,性质,可用数量,批次,效期)
                    Values(
                        r_Drug.库房ID,r_Drug.药品ID,1,v_当前量,v_批次,v_效期);
                End IF;

                --未发药品记录
                Update 未发药品记录
                    Set 病人ID=r_Drug.病人ID,
                        主页ID=r_Drug.主页ID,
                        姓名=r_Drug.姓名
                 Where 单据=r_Drug.单据 And NO=NO_IN And 库房ID+0=r_Drug.库房ID;

                IF SQL%RowCount=0 Then
                    --取身份优先级
                    Begin
                        Select B.优先级 Into v_优先级 From 病人信息 A,身份 B
                         Where A.身份=B.名称(+) And A.病人ID=r_Drug.病人ID;
                    Exception
                        When Others Then Null;
                    End;

                    Insert Into 未发药品记录(
                        单据,NO,病人ID,主页ID,姓名,优先级,对方部门ID,库房ID,填制日期,已收费,打印状态)
                    Values(
                        r_Drug.单据,NO_IN,r_Drug.病人ID,r_Drug.主页ID,r_Drug.姓名,v_优先级,r_Drug.对方部门ID,r_Drug.库房ID,v_Date,Decode(v_药品划价,1,0,1),0);
                End IF;
                
                v_收发序号:=v_收发序号+1;

                --病人费用记录
                -------------------------------------------------------------------------------------
                --记录序号范围以处理汇总表
                IF v_开始序号 IS NULL Then
                    v_开始序号:=v_费用序号;    
                End IF;
                v_结束序号:=v_费用序号;

                Insert Into 病人费用记录(
                    ID,记录性质,NO,记录状态,序号,从属父号,价格父号,多病人单,门诊标志,病人ID,主页ID,
                    标识号,姓名,性别,年龄,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,
                    保险项目否,保险大类ID,付数,数次,加班标志,附加标志,婴儿费,收入项目ID,收据费目,标准单价,
                    应收金额,实收金额,统筹金额,记帐费用,开单部门ID,开单人,发生时间,登记时间,执行部门ID,
                    执行状态,医嘱序号,划价人,操作员编号,操作员姓名)
                Select 
                    v_费用ID,2,NO_IN,Decode(v_药品划价,1,0,1),v_费用序号,NULL,NULL,多病人单,2,病人ID,主页ID,标识号,姓名,性别,年龄,
                    床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,保险项目否,保险大类ID,
                    1,-1*v_当前量,加班标志,附加标志,婴儿费,收入项目ID,收据费目,标准单价,
                    Round(-1*v_当前量*标准单价,v_Dec),Round(-1*v_当前量*标准单价,v_Dec),NULL,1,开单部门ID,开单人,
                    v_Date,v_Date,执行部门ID,0,医嘱序号,v_人员姓名,Decode(v_药品划价,1,NULL,v_人员编号),Decode(v_药品划价,1,NULL,v_人员姓名)
                From 病人费用记录 Where ID=r_Drug.费用ID;

                Begin
                    Select Round(B.应收金额*A.实收比率/100,v_Dec) Into v_实收金额
                    From 费别明细 A,病人费用记录 B
                    Where B.ID=v_费用ID And A.收入项目ID=B.收入项目ID And A.费别=B.费别 
                        And Abs(B.应收金额) Between A.应收段首值 And A.应收段尾值;

                    Update 病人费用记录 A Set 实收金额=v_实收金额 Where ID=v_费用ID;
                Exception
                    When Others Then NULL;
                End;
                
                v_费用序号:=v_费用序号+1;

                If v_剩余量<=0 Then 
                    Exit;
                End IF;
            End Loop;

            If v_剩余量<>0 Then
                --没有收回所有数量,收发记录本身有问题(如记录不全或数量为负)
                NULL;
            End IF;
        Else
            --其它非药医嘱(包括给药途径)
            -----------------------------------------------------------------------------------------------------
            Select Nvl(Max(序号),0)+1 Into v_费用序号 From 病人费用记录 Where 记录性质=2 And 记录状态 IN(0,1) And NO=NO_IN;
            For r_Other In c_Other Loop
                --记录序号范围以处理汇总表
                IF v_开始序号 IS NULL Then
                    v_开始序号:=v_费用序号;    
                End IF;
                v_结束序号:=v_费用序号;
                
                --病人费用记录:按理如果收回量大于了上次发送量,则不正确
                Select 病人费用记录_ID.Nextval Into v_费用ID From Dual;
                Insert Into 病人费用记录(
                    ID,记录性质,NO,记录状态,序号,从属父号,价格父号,多病人单,门诊标志,病人ID,主页ID,
                    标识号,姓名,性别,年龄,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,
                    保险项目否,保险大类ID,付数,数次,加班标志,附加标志,婴儿费,收入项目ID,收据费目,标准单价,
                    应收金额,实收金额,统筹金额,记帐费用,开单部门ID,开单人,发生时间,登记时间,执行部门ID,
                    执行状态,医嘱序号,划价人,操作员编号,操作员姓名)
                Select 
                    v_费用ID,2,NO_IN,Decode(v_其他划价,1,0,1),v_费用序号,NULL,Decode(价格父号,NULL,NULL,v_费用序号+价格父号-序号),多病人单,2,
                    病人ID,主页ID,标识号,姓名,性别,年龄,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,
                    保险项目否,保险大类ID,1,-1*收回量_IN,加班标志,附加标志,婴儿费,收入项目ID,收据费目,标准单价,
                    Round(-1*收回量_IN*标准单价,v_Dec),Round(-1*收回量_IN*标准单价,v_Dec),NULL,1,开单部门ID,开单人,
                    v_Date,v_Date,执行部门ID,0,医嘱序号,v_人员姓名,Decode(v_其他划价,1,NULL,v_人员编号),Decode(v_其他划价,1,NULL,v_人员姓名)
                From 病人费用记录 Where ID=r_Other.费用ID;
                
                Begin
                    Select Round(B.应收金额*A.实收比率/100,v_Dec) Into v_实收金额
                    From 费别明细 A,病人费用记录 B
                    Where B.ID=v_费用ID And A.收入项目ID=B.收入项目ID And A.费别=B.费别 
                        And Abs(B.应收金额) Between A.应收段首值 And A.应收段尾值;

                    Update 病人费用记录 A Set 实收金额=v_实收金额 Where ID=v_费用ID;
                Exception
                    When Others Then NULL;
                End;

                v_费用序号:=v_费用序号+1;
            End Loop;
        End IF;
    End IF;

    --处理费用汇总表
    -----------------------------------------------------------------------------------------------------
    If v_开始序号 IS Not NULL And v_结束序号 IS Not NULL Then
        --最后统一处理费用相关汇总表
        For r_Money IN c_Money(v_开始序号,v_结束序号) Loop
            --病人余额
            Update 病人余额
                Set 费用余额=Nvl(费用余额,0)+r_Money.实收金额
            Where 病人ID=r_Money.病人ID And 性质=1;

            IF SQL%RowCount=0 Then
                Insert Into 病人余额(
                    病人ID,性质,费用余额,预交余额)
                Values(
                    r_Money.病人ID,1,r_Money.实收金额,0);
            End IF;

            --病人未结费用
            Update 病人未结费用
                Set 金额=Nvl(金额,0)+r_Money.实收金额
            Where 病人ID=r_Money.病人ID
                And 主页ID=r_Money.主页ID
                And Nvl(病人病区ID,0)=Nvl(r_Money.病人病区ID,0)
                And Nvl(病人科室ID,0)=Nvl(r_Money.病人科室ID,0)
                And 开单部门ID+0=r_Money.开单部门ID
                And 执行部门ID+0=r_Money.执行部门ID
                And 收入项目ID+0=r_Money.收入项目ID
                And 来源途径+0=2;

            IF SQL%RowCount=0 Then
                Insert Into 病人未结费用(
                    病人ID,主页ID,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,金额)
                Values(
                    r_Money.病人ID,r_Money.主页ID,r_Money.病人病区ID,r_Money.病人科室ID,r_Money.开单部门ID,
                    r_Money.执行部门ID,r_Money.收入项目ID,2,r_Money.实收金额);
            End IF;

            --病人费用汇总
            Update 病人费用汇总
                Set 应收金额=Nvl(应收金额,0)+r_Money.应收金额,
                    实收金额=Nvl(实收金额,0)+r_Money.实收金额
            Where 日期=Trunc(v_Date)
                And Nvl(病人病区ID,0)=Nvl(r_Money.病人病区ID,0)
                And Nvl(病人科室ID,0)=Nvl(r_Money.病人科室ID,0)
                And 开单部门ID+0=r_Money.开单部门ID
                And 执行部门ID+0=r_Money.执行部门ID
                And 收入项目ID+0=r_Money.收入项目ID
                And 来源途径=2
                And 记帐费用=1;

            IF SQL%RowCount=0 Then
                Insert Into 病人费用汇总(
                    日期,病人病区ID,病人科室ID,开单部门ID,执行部门ID,收入项目ID,来源途径,记帐费用,应收金额,实收金额,结帐金额)
                Values(
                    Trunc(v_Date),r_Money.病人病区ID,r_Money.病人科室ID,r_Money.开单部门ID,r_Money.执行部门ID,
                    r_Money.收入项目ID,2,1,r_Money.应收金额,r_Money.实收金额,0);
            End IF;
        End Loop;                            
    End IF;

    --处理医嘱的上次执行时间:给药途径等可能因为未发送而没调用收回过程。
    -----------------------------------------------------------------------------------------------------
	Select Nvl(相关ID,ID) Into v_组ID From 病人医嘱记录 Where ID=医嘱ID_IN;
    Update 病人医嘱记录 Set 上次执行时间=上次时间_IN Where ID=v_组ID Or 相关ID=v_组ID;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱记录_收回;
/

CREATE OR REPLACE Procedure ZL_病人诊断记录_Insert(
--功能：插入病人诊断记录
    病人ID_IN		病人诊断记录.病人ID%Type,
    主页ID_IN		病人诊断记录.主页ID%Type,
    记录来源_IN		病人诊断记录.记录来源%Type,
    病历ID_IN		病人诊断记录.病历ID%Type,
    诊断类型_IN		病人诊断记录.诊断类型%Type,
    疾病ID_IN		病人诊断记录.疾病ID%Type,
    诊断ID_IN		病人诊断记录.诊断ID%Type,
    证候ID_IN		病人诊断记录.证候ID%Type,
    诊断描述_IN		病人诊断记录.诊断描述%Type,
    出院情况_IN		病人诊断记录.出院情况%Type,
    是否未治_IN		病人诊断记录.是否未治%Type,
    是否疑诊_IN		病人诊断记录.是否疑诊%Type,
    记录日期_IN		病人诊断记录.记录日期%Type,
    医嘱ID_IN		病人诊断记录.医嘱ID%Type:=NULL,
	诊断次序_IN		病人诊断记录.诊断次序%Type:=1
) IS
    v_Temp            Varchar2(255);
    v_人员编号        人员表.编号%Type;
    v_人员姓名        人员表.姓名%Type;
Begin
    --当前操作人员
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);
    
    Insert Into 病人诊断记录(
        ID,病人ID,主页ID,记录来源,病历ID,诊断类型,诊断次序,疾病ID,诊断ID,证候ID,诊断描述,出院情况,是否未治,是否疑诊,记录日期,记录人,医嘱id)
    Values(
        病人诊断记录_ID.Nextval,病人ID_IN,主页ID_IN,记录来源_IN,病历ID_IN,诊断类型_IN,诊断次序_IN,疾病ID_IN,
        诊断ID_IN,证候ID_IN,诊断描述_IN,出院情况_IN,是否未治_IN,是否疑诊_IN,记录日期_IN,v_人员姓名,医嘱id_IN);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人诊断记录_Insert;
/

CREATE OR REPLACE Procedure zl_医嘱内容定义_Delete
IS
Begin
	Delete From 医嘱内容定义;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_医嘱内容定义_Delete;
/

CREATE OR REPLACE Procedure zl_医嘱内容定义_Insert(
	诊疗类别_IN		医嘱内容定义.诊疗类别%Type,
	医嘱内容_IN		医嘱内容定义.医嘱内容%Type
) IS
Begin
	Update 医嘱内容定义 Set 医嘱内容=医嘱内容_IN Where 诊疗类别=诊疗类别_IN;
	If SQL%RowCount=0 Then
		Insert Into 医嘱内容定义(
			诊疗类别,医嘱内容)
		Values(
			诊疗类别_IN,医嘱内容_IN);
	End IF;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_医嘱内容定义_Insert;
/

CREATE OR REPLACE Procedure zl_人员证书记录_Insert(
	人员ID_IN		人员证书记录.人员ID%Type,
	CertDN_IN		人员证书记录.CertDN%Type,
	CertSN_IN		人员证书记录.CertSN%Type,
	SignCert_IN		人员证书记录.SignCert%Type,
	EncCert_IN		人员证书记录.EncCert%Type
) IS
	v_姓名			人员表.姓名%Type;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
	Begin
		Select A.姓名 Into v_姓名 From 人员表 A,人员证书记录 B Where A.ID=B.人员ID And B.CertSN=CertSN_IN;
	Exception
		When Others Then Null;
	End;
	If v_姓名 Is Not Null Then
		v_Error:='该数字证书已经注册给"'||v_姓名||'"，不能重复注册。';
		Raise Err_Custom;
	End IF;

	Insert Into 人员证书记录(
		ID,人员ID,CertDN,CertSN,SignCert,EncCert,注册时间)
	Values(
		人员证书记录_ID.Nextval,人员ID_IN,CertDN_IN,CertSN_IN,SignCert_IN,EncCert_IN,Sysdate);
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_人员证书记录_Insert;
/

CREATE OR REPLACE Procedure zl_人员证书记录_Delete(
	证书ID_IN		人员证书记录.ID%Type
) IS
Begin
	Delete From 人员证书记录 Where ID=证书ID_IN;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_人员证书记录_Delete;
/

CREATE OR REPLACE Procedure zl_医嘱签名记录_Insert(
	签名ID_IN		医嘱签名记录.ID%Type,
	签名类型_IN		病人医嘱状态.操作类型%Type,--对应为1-新开,4-作废,8-停止
	签名规则_IN		医嘱签名记录.签名规则%Type,
	签名信息_IN		医嘱签名记录.签名信息%Type,
	证书ID_IN		医嘱签名记录.证书ID%Type,
	医嘱IDs_IN		Varchar2 --本次签名的医嘱ID序列,格式为'1,2,3,...'
) IS
	v_医嘱IDs		Varchar2(4000);
	v_当前ID		病人医嘱记录.ID%Type;

    v_Temp			Varchar2(255);
    v_人员姓名		病人费用记录.操作员姓名%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --当前操作人员
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);
	
	Insert Into 医嘱签名记录(
		ID,签名规则,签名信息,证书ID,签名时间,签名人)
	Values(
		签名ID_IN,签名规则_IN,签名信息_IN,证书ID_IN,Sysdate,v_人员姓名);
	
	--关联签名对应的医嘱
    v_医嘱IDs:=医嘱IDs_IN||',';
	While v_医嘱IDs Is Not Null Loop
		v_当前ID:=to_Number(Substr(v_医嘱IDs,1,Instr(v_医嘱IDs,',')-1));
		
		--因为这几个操作都不可重复，因此操作时间的判断也可以不要。
		Update 病人医嘱状态 
			Set 签名ID=签名ID_IN 
		Where 医嘱ID=v_当前ID And 操作类型=签名类型_IN And 签名ID Is Null
			And 操作时间=(Select Max(操作时间) From 病人医嘱状态 Where 医嘱ID=v_当前ID And 操作类型=签名类型_IN);
		If SQL%RowCount=0 Then
			v_Error:='没有找到要签名的医嘱，不能有效完成电子签名。';
			Raise Err_Custom;
		End IF;

		v_医嘱IDs:=Substr(v_医嘱IDs,Instr(v_医嘱IDs,',')+1);
	End Loop;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_医嘱签名记录_Insert;
/

CREATE OR REPLACE Procedure zl_医嘱签名记录_Delete(
	签名ID_IN		医嘱签名记录.ID%Type
) IS
	v_Count			Number;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
	--取消新开医嘱签名的并发检查
	Select 
		Count(A.ID) Into v_Count
	From 病人医嘱记录 A,病人医嘱状态 B 
	Where A.医嘱状态 Not IN(1,2) And A.ID=B.医嘱ID 
		And B.操作类型=1 And B.签名ID=签名ID_IN;
	If Nvl(v_Count,0)>0 Then
		v_Error:='相关医嘱已经校对或发送，不能取消电子签名。';
		Raise Err_Custom;
	End IF;
	
	Update 病人医嘱状态 Set 签名ID=NULL Where 签名ID=签名ID_IN;
	Delete From 医嘱签名记录 Where ID=签名ID_IN;
	If SQL%RowCount=0 Then
		v_Error:='没有找到医嘱的签名记录，无法取消电子签名。';
		Raise Err_Custom;
	End IF;
Exception
	When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_医嘱签名记录_Delete;
/

CREATE OR REPLACE Procedure ZL_病人医嘱记录_Insert(
--功能：医生或护士新开,补录医嘱时新产生的医嘱记录。可用于门诊或住院。
    ID_IN				病人医嘱记录.ID%TYPE,
    相关ID_IN           病人医嘱记录.相关ID%TYPE,
    序号_IN				病人医嘱记录.序号%TYPE,
    病人来源_IN         病人医嘱记录.病人来源%TYPE,
    病人ID_IN           病人医嘱记录.病人ID%TYPE,
    主页ID_IN           病人医嘱记录.主页ID%TYPE,
    婴儿_IN             病人医嘱记录.婴儿%TYPE,
    医嘱状态_IN         病人医嘱记录.医嘱状态%TYPE,
    医嘱期效_IN         病人医嘱记录.医嘱期效%TYPE,
    诊疗类别_IN         病人医嘱记录.诊疗类别%TYPE,
    诊疗项目ID_IN       病人医嘱记录.诊疗项目ID%TYPE,
    收费细目ID_IN       病人医嘱记录.收费细目ID%TYPE,
	天数_IN				病人医嘱记录.天数%TYPE,
    单次用量_IN         病人医嘱记录.单次用量%TYPE,
    总给予量_IN         病人医嘱记录.总给予量%TYPE,
    医嘱内容_IN         病人医嘱记录.医嘱内容%TYPE,
    医生嘱托_IN         病人医嘱记录.医生嘱托%TYPE,
    标本部位_IN         病人医嘱记录.标本部位%TYPE,
    执行频次_IN         病人医嘱记录.执行频次%TYPE,
    频率次数_IN         病人医嘱记录.频率次数%TYPE,
    频率间隔_IN         病人医嘱记录.频率间隔%TYPE,
    间隔单位_IN         病人医嘱记录.间隔单位%TYPE,
    执行时间方案_IN		病人医嘱记录.执行时间方案%TYPE,
    计价特性_IN         病人医嘱记录.计价特性%TYPE,
    执行科室ID_IN       病人医嘱记录.执行科室ID%TYPE,
    执行性质_IN         病人医嘱记录.执行性质%TYPE,
    紧急标志_IN         病人医嘱记录.紧急标志%TYPE,
    开始执行时间_IN     病人医嘱记录.开始执行时间%TYPE,
    执行终止时间_IN     病人医嘱记录.执行终止时间%TYPE,
    病人科室ID_IN       病人医嘱记录.病人科室ID%TYPE,
    开嘱科室ID_IN       病人医嘱记录.开嘱科室ID%TYPE,
    开嘱医生_IN         病人医嘱记录.开嘱医生%TYPE,
    开嘱时间_IN         病人医嘱记录.开嘱时间%TYPE,
    挂号单_IN           病人医嘱记录.挂号单%TYPE:=NULL,
	前提ID_IN			病人医嘱记录.前提ID%Type:=NULL
) IS
    v_Temp			Varchar2(255);
    v_人员编号		病人费用记录.操作员编号%Type;
    v_人员姓名		病人费用记录.操作员姓名%Type;
	
	v_姓名			病人医嘱记录.姓名%Type;
	v_性别			病人医嘱记录.性别%Type;
	v_年龄			病人医嘱记录.年龄%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --当前操作人员
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);
	
	Select 姓名,性别,年龄 Into v_姓名,v_性别,v_年龄 From 病人信息 Where 病人ID=病人ID_IN;

    --病人医嘱记录
    Insert Into 病人医嘱记录(
        ID,相关ID,序号,病人来源,病人ID,主页ID,姓名,性别,年龄,婴儿,医嘱状态,医嘱期效,诊疗类别,诊疗项目ID,收费细目ID,
        天数,单次用量,总给予量,医嘱内容,医生嘱托,标本部位,执行频次,频率次数,频率间隔,间隔单位,执行时间方案,
        计价特性,执行科室ID,执行性质,紧急标志,开始执行时间,执行终止时间,病人科室ID,开嘱科室ID,开嘱医生,
		开嘱时间,挂号单,前提ID)
    VALUES(
        ID_IN,相关ID_IN,序号_IN,病人来源_IN,病人ID_IN,主页ID_IN,v_姓名,v_性别,v_年龄,婴儿_IN,医嘱状态_IN,医嘱期效_IN,
        诊疗类别_IN,诊疗项目ID_IN,收费细目ID_IN,天数_IN,单次用量_IN,总给予量_IN,医嘱内容_IN,医生嘱托_IN,
        标本部位_IN,执行频次_IN,频率次数_IN,频率间隔_IN,间隔单位_IN,执行时间方案_IN,计价特性_IN,
        执行科室ID_IN,执行性质_IN,紧急标志_IN,开始执行时间_IN,执行终止时间_IN,
        病人科室ID_IN,开嘱科室ID_IN,开嘱医生_IN,开嘱时间_IN,挂号单_IN,前提ID_IN);

    --病人医嘱状态
	Delete From 病人医嘱状态 Where 医嘱ID=ID_IN And 操作类型=1;
	If SQL%RowCount<>0 Then
		v_Error:='相同ID的新开医嘱已经存在。';
		Raise Err_Custom;
	End IF;
	--因为可能同时：新开->自动校对->互斥自动停止,因此分别-2,-1秒
    Insert Into 病人医嘱状态(
        医嘱ID,操作类型,操作人员,操作时间)
    Values(
        ID_IN,1,v_人员姓名,Sysdate-2/60/60/24);
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱记录_Insert;
/

CREATE OR REPLACE Procedure ZL_病人医嘱记录_Update(
--功能：被医生或护士修改了部分内容的医嘱记录。可用于门诊或住院。
--说明：Update时之所以涉及诊疗项目ID,计价特性变化,是因为给药途径,用法的变化
--      Update时之所以涉及期效变化,是因为自由录入医嘱可任意改变期效
    ID_IN               病人医嘱记录.ID%TYPE,
    相关ID_IN           病人医嘱记录.相关ID%TYPE,
    序号_IN             病人医嘱记录.序号%TYPE,
    医嘱状态_IN         病人医嘱记录.医嘱状态%TYPE,
	医嘱期效_IN			病人医嘱记录.医嘱期效%TYPE,
    诊疗项目ID_IN       病人医嘱记录.诊疗项目ID%TYPE,
	天数_IN				病人医嘱记录.天数%TYPE,
    单次用量_IN         病人医嘱记录.单次用量%TYPE,
    总给予量_IN         病人医嘱记录.总给予量%TYPE,
    医嘱内容_IN         病人医嘱记录.医嘱内容%TYPE,
    医生嘱托_IN         病人医嘱记录.医生嘱托%TYPE,
    标本部位_IN         病人医嘱记录.标本部位%TYPE,
    执行频次_IN         病人医嘱记录.执行频次%TYPE,
    频率次数_IN         病人医嘱记录.频率次数%TYPE,
    频率间隔_IN         病人医嘱记录.频率间隔%TYPE,
    间隔单位_IN         病人医嘱记录.间隔单位%TYPE,
    执行时间方案_IN     病人医嘱记录.执行时间方案%TYPE,
    计价特性_IN         病人医嘱记录.计价特性%TYPE,
    执行科室ID_IN       病人医嘱记录.执行科室ID%TYPE,
    执行性质_IN         病人医嘱记录.执行性质%TYPE,
    紧急标志_IN         病人医嘱记录.紧急标志%TYPE,
    开始执行时间_IN     病人医嘱记录.开始执行时间%TYPE,
    执行终止时间_IN     病人医嘱记录.执行终止时间%TYPE,
    病人科室ID_IN       病人医嘱记录.病人科室ID%TYPE,
    开嘱科室ID_IN       病人医嘱记录.开嘱科室ID%TYPE,
    开嘱医生_IN         病人医嘱记录.开嘱医生%TYPE,
    开嘱时间_IN         病人医嘱记录.开嘱时间%TYPE
) IS
    v_Count			Number;

    v_Temp			Varchar2(255);
    v_人员编号		病人费用记录.操作员编号%Type;
    v_人员姓名		病人费用记录.操作员姓名%Type;

    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --检查该医嘱状态:并发操作
    Begin
        Select 医嘱状态 Into v_Count From 病人医嘱记录 Where ID=ID_IN;
    Exception
        When Others Then
        Begin
            v_Error:='医嘱"'||医嘱内容_IN||'"已经不存在,可能已被其他人删除。';
            Raise Err_Custom;
        End;
    End;
    If v_Count Not IN(1,2) Then
        v_Error:='医嘱"'||医嘱内容_IN||'"已经校对或发送,不能再修改。';
        Raise Err_Custom;
    End IF;

	Select Count(*) Into v_Count From 病人医嘱状态 Where 医嘱ID=ID_IN And 操作类型=1 And 签名ID Is Not Null;
	If Nvl(v_Count,0)>0 Then
        v_Error:='医嘱"'||医嘱内容_IN||'"已经电子签名,不能再修改。';
        Raise Err_Custom;
	End IF;

    --当前操作人员    
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --病人医嘱记录
    Update 病人医嘱记录 
        Set 相关ID=相关ID_IN,--比如一并给药，重新设置检查部位等引起的相关ID变化
            序号=序号_IN,
            医嘱状态=医嘱状态_IN,--!因为只能修改未校对医嘱，所以应该为新开，校对疑问的医嘱修改后为新开
			医嘱期效=医嘱期效_IN,
            诊疗项目ID=诊疗项目ID_IN,
			天数=天数_IN,
            单次用量=单次用量_IN,
            总给予量=总给予量_IN,
            医嘱内容=医嘱内容_IN,
            医生嘱托=医生嘱托_IN,
            标本部位=标本部位_IN,
            执行频次=执行频次_IN,
            频率次数=频率次数_IN,
            频率间隔=频率间隔_IN,
            间隔单位=间隔单位_IN,
            执行时间方案=执行时间方案_IN,
            计价特性=计价特性_IN,
            执行科室ID=执行科室ID_IN,
            执行性质=执行性质_IN,--药品根据外购药,出院带药的调整时会发生变化
            紧急标志=紧急标志_IN,
            开始执行时间=开始执行时间_IN,
            执行终止时间=执行终止时间_IN,--!长嘱的终止时间可以修改,临嘱应该为空
            病人科室ID=病人科室ID_IN,--修改时更新为病人的当前科室
            开嘱科室ID=开嘱科室ID_IN,--修改后会根据当前科室变化
            开嘱医生=开嘱医生_IN,--护士开医嘱时可以更改
            开嘱时间=开嘱时间_IN--补录的可以修改
    Where ID=ID_IN;
    
    --病人医嘱状态:更新医生新开这条
    Update 病人医嘱状态
        Set 操作人员=v_人员姓名,
            操作时间=Sysdate
    Where 医嘱ID=ID_IN And 操作类型=1;--新开这条始终有,校对疑问保留作为历史记录
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱记录_Update;
/

CREATE OR REPLACE Procedure ZL_病人医嘱记录_Delete(
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
    医嘱ID_IN		病人医嘱记录.ID%TYPE,
    删相关_IN       Number:=0
) IS
	Cursor c_Case is
		Select 申请ID From 病人医嘱记录 Where 申请ID IS Not NULL And (ID=医嘱ID_IN Or 相关ID=医嘱ID_IN); 
	r_Case	c_Case%RowType;

    v_状态			病人医嘱记录.医嘱状态%Type;
    v_相关ID		病人医嘱记录.相关ID%Type;
    v_病人ID		病人医嘱记录.病人ID%Type;
    v_挂号单		病人医嘱记录.挂号单%Type;
    v_主页ID		病人医嘱记录.主页ID%Type;
    v_婴儿			病人医嘱记录.婴儿%Type;
    v_序号			病人医嘱记录.序号%Type;
    v_内容			病人医嘱记录.医嘱内容%Type;
    v_Count			Number(5);

    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --检查医嘱状态:并发操作
    Begin
        Select 病人ID,挂号单,主页ID,婴儿,医嘱状态,相关ID,医嘱内容
            Into v_病人ID,v_挂号单,v_主页ID,v_婴儿,v_状态,v_相关ID,v_内容
        From 病人医嘱记录 Where ID=医嘱ID_IN;
    Exception
        When Others Then
        Begin
            v_Error:='未发现要删除的医嘱记录，可能已被其他人删除。';
            Raise Err_Custom;
        End;
    End;
    If v_挂号单 IS NULL Then
        IF Not v_状态 IN(1,2) Then
            v_Error:='医嘱"'||v_内容||'"已经过校对，不能再删除。';
            Raise Err_Custom;
        End IF;
    Else
        IF v_状态<>1 Then
            v_Error:='医嘱"'||v_内容||'"已经被发送或作废，不能删除。';
            Raise Err_Custom;
        End IF;
    End IF;

	Select Count(*) Into v_Count From 病人医嘱状态 Where 医嘱ID=医嘱ID_IN And 操作类型=1 And 签名ID Is Not Null;
	If Nvl(v_Count,0)>0 Then
        v_Error:='医嘱"'||v_内容||'"已经电子签名,不能删除。';
        Raise Err_Custom;
	End IF;

	IF Nvl(删相关_IN,0)=0 then
		Begin
			Select 申请ID Into v_Count From 病人医嘱记录 Where ID=医嘱ID_IN;
		Exception
			When Others Then v_Count:=NULL;
		End;

		--删除医嘱
		Delete From 病人医嘱记录 Where ID=医嘱ID_IN;
		
		--删除对应的申请单
		IF v_Count IS Not NULL Then
			Delete From 病人病历记录 Where ID=v_Count;
		End IF;
    Else
		If v_相关ID IS NULL Then
            --检查组合,手术及附加,中药配方,检验组合,以及独立医嘱
            Select Max(序号),Count(*) Into v_序号,v_Count From 病人医嘱记录 Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;
			
			Open c_Case;--必须先打开

			--删除医嘱
			Delete From 病人医嘱记录 Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;

			--删除对应的申请单
			Fetch c_Case Into r_Case;
			While c_Case%Found Loop
				Delete From 病人病历记录 Where ID=r_Case.申请ID;
				Fetch c_Case Into r_Case;
			End Loop;
			Close c_Case;
        Else
            --成药一并给药的情况(无申请)
            --先判断是否一并给药
            Select Count(*) Into v_Count From 病人医嘱记录 Where 相关ID=v_相关ID;
            
            If v_Count=1 Then
                --单独给药:同时删除其给药途径
                Select Max(序号),Count(*) Into v_序号,v_Count From 病人医嘱记录 Where ID=医嘱ID_IN Or ID=v_相关ID;
                Delete From 病人医嘱记录 Where ID=医嘱ID_IN Or ID=v_相关ID;
            Else
                --一并给药:只删除当前药品
                v_Count:=1;
                Select 序号 Into v_序号 From 病人医嘱记录 Where ID=医嘱ID_IN;
                Delete From 病人医嘱记录 Where ID=医嘱ID_IN;
            End If;
        End if;

        --调整序号
        Update 病人医嘱记录 
            Set 序号=序号-v_Count
        Where 病人ID=v_病人ID 
            And Nvl(主页ID,0)=Nvl(v_主页ID,0) 
            And Nvl(挂号单,'空')=Nvl(v_挂号单,'空')
            And Nvl(婴儿,0)=Nvl(v_婴儿,0) 
            And 序号>v_序号;
    End if;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱记录_Delete;
/

CREATE OR REPLACE Procedure ZL_病人医嘱记录_校对(
--功能：校对指定的医嘱
--参数：医嘱ID_IN=Nvl(相关ID,ID)
--      状态_IN=校对通过3或校对疑问2
--      自动校对_IN=保存之后调用自动校对,自动填写计价内容
--说明：一组医嘱只能调用一次,过程同时完成处理一组医嘱的校对
    医嘱ID_IN		病人医嘱记录.ID%TYPE,
    状态_IN			病人医嘱记录.医嘱状态%TYPE,
    校对时间_IN		病人医嘱状态.操作时间%TYPE,
	自动校对_IN		Number:=Null
) IS
    --用于医嘱检查
    v_状态			病人医嘱记录.医嘱状态%Type;
	v_期效			病人医嘱记录.医嘱期效%Type;
    v_病人ID        病人医嘱记录.病人ID%Type;
    v_主页ID        病人医嘱记录.主页ID%Type;
    v_婴儿			病人医嘱记录.婴儿%Type;
    v_医嘱内容		病人医嘱记录.医嘱内容%Type;
	v_开嘱时间		病人医嘱记录.开嘱时间%Type;
	v_开始时间		病人医嘱记录.开始执行时间%Type;
	v_开嘱医生		病人医嘱记录.开嘱医生%Type;
	v_前提ID		病人医嘱记录.前提ID%Type;

    --用于变更护理等级
    v_诊疗类别		病人医嘱记录.诊疗类别%TYPE;
    v_诊疗项目ID    病人医嘱记录.诊疗项目ID%TYPE;
    v_操作类型      诊疗项目目录.操作类型%TYPE;
    v_护理等级ID    病案主页.护理等级ID%TYPE;

    --与该项目同一自动停止互斥组的项目:组中应该都是长嘱(包括当前医嘱),程序应已检查。
    --注意应加婴儿条件,同时也应停止除当前医嘱外的其它相同诊疗项目的医嘱。
    Cursor c_Exclude IS
        Select Distinct B.ID AS 医嘱ID,B.开始执行时间,B.执行终止时间,B.上次执行时间,B.开嘱医生,
            B.执行时间方案,B.频率间隔,B.频率次数,B.间隔单位
        From 诊疗互斥项目 A,病人医嘱记录 B
        Where A.类型=3 And A.项目ID=B.诊疗项目ID And B.ID<>医嘱ID_IN
            And Nvl(B.医嘱期效,0)=0 And B.医嘱状态 IN(3,5,6,7)
            And B.病人ID=v_病人ID And Nvl(B.主页ID,0)=Nvl(v_主页ID,0) And Nvl(B.婴儿,0)=Nvl(v_婴儿,0)
            And A.组编号 IN(Select Distinct 组编号 From 诊疗互斥项目 Where 类型=3 And 项目ID=v_诊疗项目ID)
            Order by B.ID;
    v_终止时间 病人医嘱记录.执行终止时间%TYPE;

    Cursor c_Nurse IS
        Select A.ID AS 医嘱ID,A.开始执行时间,A.执行终止时间,A.上次执行时间,A.开嘱医生
        From 病人医嘱记录 A,诊疗项目目录 B
        Where A.诊疗项目ID=B.ID And A.诊疗类别='H' And B.操作类型='1'
            And A.病人ID=v_病人ID And Nvl(A.主页ID,0)=Nvl(v_主页ID,0) And Nvl(A.婴儿,0)=Nvl(v_婴儿,0)
            And Nvl(A.医嘱期效,0)=0 And A.医嘱状态 IN(3,5,6,7) And A.ID<>医嘱ID_IN;

    --包含病人(婴儿)的所有未停长嘱(含配方长嘱)
    Cursor c_NeedStop(
        v_病人ID    病人医嘱记录.病人ID%Type,
        v_主页ID    病人医嘱记录.主页ID%Type,
        v_婴儿		病人医嘱记录.婴儿%Type,
        v_StopTime	Date) is
        Select ID From 病人医嘱记录 
        Where 病人ID=v_病人ID And 主页ID=v_主页ID And Nvl(婴儿,0)=Nvl(v_婴儿,0) 
            And Nvl(医嘱期效,0)=0 And 医嘱状态 Not IN(1,2,4,8,9)
            And 开始执行时间<v_StopTime
        Order BY 序号;
	--包含病人(婴儿)的已停但未确认的长嘱,终止执行时间在指定时间之后
    Cursor c_HaveStop(
        v_病人ID    病人医嘱记录.病人ID%Type,
        v_主页ID    病人医嘱记录.主页ID%Type,
        v_婴儿		病人医嘱记录.婴儿%Type,
        v_StopTime	Date) is
        Select ID From 病人医嘱记录 
        Where 病人ID=v_病人ID And 主页ID=v_主页ID And Nvl(婴儿,0)=Nvl(v_婴儿,0) 
            And Nvl(医嘱期效,0)=0 And 医嘱状态=8 And 执行终止时间>v_StopTime
			And 开始执行时间<v_StopTime
        Order BY 序号;
	
	--取一组医嘱的计价内容
	Cursor c_Price Is
		Select A.ID,B.收费项目ID,B.收费数量,B.从属项目,
			Sum(Decode(Nvl(C.是否变价,0),1,D.原价,Null)) as 单价
		From 病人医嘱记录 A,诊疗收费关系 B,收费项目目录 C,收费价目 D
		Where A.诊疗项目ID=B.诊疗项目ID And B.收费项目ID=C.ID
			And C.ID=D.收费细目ID And A.诊疗类别 Not IN('5','6','7') 
			And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5)
			And C.服务对象 IN(1,3) And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)
			And Sysdate Between D.执行日期 And Nvl(D.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))
			And Nvl(B.收费数量,0)<>0 And Not(Nvl(C.是否变价,0)=1 And Nvl(D.原价,0)=0)
			And (A.ID=医嘱ID_IN Or A.相关ID=医嘱ID_IN)
		Group by A.ID,B.收费项目ID,B.收费数量,B.从属项目;

    --其它临时变量
	v_参数值		系统参数表.参数值%Type;
	v_Count			Number;
    v_Date			Date;
    v_Temp			Varchar2(255);
    v_人员编号		病人费用记录.操作员编号%Type;
    v_人员姓名		病人费用记录.操作员姓名%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;

	Function GetAdviceText(v_医嘱ID 病人医嘱记录.ID%Type)
		Return Varchar2 Is
		v_Text	病人医嘱记录.医嘱内容%Type;
		v_类别	病人医嘱记录.诊疗类别%Type;
		v_配方	Number;
	Begin
		Select 诊疗类别,医嘱内容 Into v_类别,v_Text From 病人医嘱记录 Where ID=v_医嘱ID;
		If v_类别='E' Then
			--西药，中成药的医嘱内容
			Begin
				Select 诊疗类别,Decode(诊疗类别,'7',v_Text,医嘱内容) 
					Into v_类别,v_Text 
				From 病人医嘱记录 Where 相关ID=v_医嘱ID And 诊疗类别 IN('5','6','7') And Rownum=1;
			Exception
				When Others Then Null;
			End;
			If v_类别='7' Then
				v_配方:=1;
			End IF;
		End IF;
		If Length(v_Text)>30 Then 
			v_Text:=Substr(v_Text,1,30)||'...';
		End IF;
		If Length(v_Text)>20 Then 
			v_Text:='"'||v_Text||'"'||CHR(13)||CHR(10);
		Else
			v_Text:='"'||v_Text||'"';
		End IF;
		If v_配方=1 Then
			v_Text:='中药配方'||v_Text;
		End IF;
		Return(v_Text);
	End;
Begin
    --检查医嘱状态是否正确:并发操作
	Begin
		Select A.医嘱期效,A.医嘱状态,A.开嘱时间,A.开嘱医生,A.开始执行时间,A.病人ID,A.主页ID,A.婴儿,A.医嘱内容,A.诊疗类别,A.诊疗项目ID,A.前提ID,Nvl(B.操作类型,'0')
			Into v_期效,v_状态,v_开嘱时间,v_开嘱医生,v_开始时间,v_病人ID,v_主页ID,v_婴儿,v_医嘱内容,v_诊疗类别,v_诊疗项目ID,v_前提ID,v_操作类型
		From 病人医嘱记录 A,诊疗项目目录 B
		Where A.诊疗项目ID=B.ID(+) And A.ID=医嘱ID_IN;
	Exception
		When Others Then
		Begin
			v_Error:='医嘱已被删除，不能进行校对。'||CHR(13)||CHR(10)||'这可能是并发操作引起的，请重新读取校对数据。';
			Raise Err_Custom;
		End;
	End;
	IF v_状态<>1 Then
		v_Error:='医嘱"'||GetAdviceText(医嘱ID_IN)||'"不是新开的医嘱，不能通过校对。'||CHR(13)||CHR(10)||'这可能是并发操作引起的，请重新读取校对数据。';
		Raise Err_Custom;
	End IF;
	--再次检查校对时间的有效性:并发操作
	If To_Char(v_开嘱时间,'YYYY-MM-DD HH24:MI') <= To_Char(v_开始时间,'YYYY-MM-DD HH24:MI') Then
		If To_Char(校对时间_IN,'YYYY-MM-DD HH24:MI') < To_Char(v_开嘱时间,'YYYY-MM-DD HH24:MI') Then
			v_Error:='医嘱"'||GetAdviceText(医嘱ID_IN)||'"的校对时间不能小于开嘱时间 '||To_Char(v_开嘱时间,'YYYY-MM-DD HH24:MI')||'。'||CHR(13)||CHR(10)||'这可能是并发操作引起的，请重新读取校对数据。';
			Raise Err_Custom;
		End If;
	Else
		If To_Char(校对时间_IN,'YYYY-MM-DD HH24:MI') < To_Char(v_开始时间,'YYYY-MM-DD HH24:MI') Then
			v_Error:='医嘱"'||GetAdviceText(医嘱ID_IN)||'"的校对时间不能小于开始执行时间 '||To_Char(v_开始时间,'YYYY-MM-DD HH24:MI')||'。'||CHR(13)||CHR(10)||'这可能是并发操作引起的，请重新读取校对数据。';
			Raise Err_Custom;
		End If;
	End If;
	
	
	--如果要求签名，检查校对时是否有签名(并发取消签名)
	If 状态_IN=3 Then
		Begin
			Select 参数值 Into v_参数值 From 系统参数表 Where 参数号=25;
		Exception
			When Others Then Null;
		End;
		If Nvl(v_参数值,'0')<>'0' Then
			v_参数值:=Null;
			Begin
				Select 参数值 Into v_参数值 From 系统参数表 Where 参数号=26;
			Exception
				When Others Then Null;
			End;
			If Nvl(Substr(v_参数值,2,1),'0')='1' And v_前提ID Is Null 
				Or Nvl(Substr(v_参数值,3,1),'0')='1' And v_前提ID Is Not Null Then
				Select Count(*) Into v_Count From 病人医嘱状态 Where 操作类型=1 And 签名ID Is Not Null And 医嘱ID=医嘱ID_IN;
				If Nvl(v_Count,0)=0 Then
					v_Error:='医嘱"'||GetAdviceText(医嘱ID_IN)||'"还没有电子签名，不能通过校对。';
					Raise Err_Custom;
				End If;
			End If;
		End IF;
	End IF;

    --当前操作人员    
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);
    
    --因为可能同时：新开->自动校对->互斥自动停止,因此分别-2,-1秒
    Select Sysdate-1/60/60/24 Into v_Date From Dual;

	Update 病人医嘱记录 
        Set 医嘱状态=状态_IN,
            校对护士=v_人员姓名,
            校对时间=校对时间_IN
    Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;

    Insert Into 病人医嘱状态(
        医嘱ID,操作类型,操作人员,操作时间)
    Select 
        ID,状态_IN,v_人员姓名,v_Date
    From 病人医嘱记录
    Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;
    
    --校对通过时的其它处理
    If 状态_IN=3 Then
		--自动校对时，自动填写缺省的计价内容
		If Nvl(自动校对_IN,0)=1 Then
			--1.变价的计价项目,如果最低限价不为0,则缺省为最低限价,否则不加入;可再手工计价.
			--2.对于非药嘱药品和在用卫材未定执行科室,发送时会取缺省的,可再手工设置。
			For r_Price In c_Price Loop
				Insert Into 病人医嘱计价(
					医嘱ID,收费细目ID,数量,单价,从项,执行科室ID)
				Values(
					r_Price.ID,r_Price.收费项目ID,r_Price.收费数量,r_Price.单价,r_Price.从属项目,NULL);
			End Loop;
		End IF;
		
		--自由录入的临嘱医嘱标记为停止
		If Nvl(v_期效,0)=1 And v_诊疗项目ID Is Null Then
            Update 病人医嘱记录 
                Set 医嘱状态=8,
                    停嘱时间=校对时间_IN,
                    停嘱医生=v_开嘱医生
            Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;
            
            Insert Into 病人医嘱状态(
                医嘱ID,操作类型,操作人员,操作时间) 
            Select 
                ID,8,v_人员姓名,校对时间_IN 
            From 病人医嘱记录 
            Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;
		End IF;

        --将同一自动停止互斥组中的病人其它医嘱停止(如果尚未停止)
        For r_Exclude In c_Exclude Loop
			Select Decode(Sign(r_Exclude.开始执行时间-v_开始时间),1,r_Exclude.开始执行时间,v_开始时间) Into v_终止时间 From Dual;
            ZL_病人医嘱记录_停止(r_Exclude.医嘱ID,v_终止时间,v_开嘱医生);
        End Loop;

        --对一些特殊医嘱的处理
        If v_诊疗类别='H' And v_操作类型='1' And Nvl(v_婴儿,0)=0 Then
            --校对护理等级时,同步更改病人护理等级
			
			--病人当前应处于正常住院状态
			v_Temp:=Null;
			Begin
				Select Decode(状态,1,'等待入科',2,'正在转科',3,'已预出院',Null) Into v_Temp From 病案主页 Where 病人ID=v_病人ID And 主页ID=v_主页ID;
			Exception
				When Others Then Null;
			End;
			If v_Temp IS Not Null Then
				v_Error:='病人当前处于'||v_Temp||'状态,医嘱"'||v_医嘱内容||'"不能通过校对。';
				Raise Err_Custom;
			End If;

            Begin
                --未设置时,不处理,有多个时,只取一个。
                Select 收费项目ID Into v_护理等级ID From 诊疗收费关系 Where 诊疗项目ID=v_诊疗项目ID And Rownum=1;
            Exception
                When Others Then NULL;
            End;
            IF v_护理等级ID IS Not NULL Then
                zl_病人变动记录_Nurse(v_病人ID,v_主页ID,v_护理等级ID,v_Date,v_人员编号,v_人员姓名);
            End IF;
            
            --并停止其它护理等级医嘱(护理等级应该都为"持续性"长嘱,且只有一个未停)
            For r_Nurse In c_Nurse Loop
				Select Decode(Sign(r_Nurse.开始执行时间-v_开始时间),1,r_Nurse.开始执行时间,v_开始时间) Into v_终止时间 From Dual;
                ZL_病人医嘱记录_停止(r_Nurse.医嘱ID,v_终止时间,v_开嘱医生);
            End Loop;
		ElsIf v_诊疗类别='Z' And v_操作类型='4' Then
			--术后医嘱校对时停止前面的长嘱,在术后开始时终止
			For r_NeedStop IN c_NeedStop(v_病人ID,v_主页ID,v_婴儿,v_开始时间) Loop
				Update 病人医嘱记录
					Set 医嘱状态=8,
						执行终止时间=Decode(Sign(开始执行时间-v_开始时间),1,开始执行时间,v_开始时间),
						停嘱时间=校对时间_IN,
						停嘱医生=v_开嘱医生
				Where ID=r_NeedStop.ID;

				Insert Into 病人医嘱状态(
					医嘱ID,操作类型,操作人员,操作时间) 
				Select 
					ID,8,v_人员姓名,校对时间_IN
				From 病人医嘱记录 
				Where ID=r_NeedStop.ID;
			End Loop;
			--已停止未确认的长嘱,终止时间在术后开始后的,调前其终止时间(同时多个术后或特殊医嘱的情况)
			For r_HaveStop IN c_HaveStop(v_病人ID,v_主页ID,v_婴儿,v_开始时间) Loop
				Update 病人医嘱记录
					Set 执行终止时间=Decode(Sign(开始执行时间-v_开始时间),1,开始执行时间,v_开始时间),
						停嘱时间=校对时间_IN,
						停嘱医生=v_开嘱医生
				Where ID=r_HaveStop.ID;
				
				Update 病人医嘱状态
					Set 操作时间=校对时间_IN,
						操作人员=v_人员姓名
				Where 医嘱ID=r_HaveStop.ID And 操作类型=8;
			End Loop;
        End IF;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱记录_校对;
/

CREATE OR REPLACE Procedure ZL_病人医嘱发送_Insert(
--功能：填写病人医嘱发送记录
    医嘱ID_IN		病人医嘱发送.医嘱ID%Type,
    发送号_IN       病人医嘱发送.发送号%Type,
    记录性质_IN     病人医嘱发送.记录性质%Type,
    NO_IN           病人医嘱发送.NO%Type,
    记录序号_IN     病人医嘱发送.记录序号%Type,
    发送数次_IN     病人医嘱发送.发送数次%Type,
    首次时间_IN     病人医嘱发送.首次时间%Type,
    末次时间_IN     病人医嘱发送.末次时间%Type,
    发送时间_IN     病人医嘱发送.发送时间%Type,
    执行状态_IN     病人医嘱发送.执行状态%Type,
    执行部门ID_IN   病人医嘱发送.执行部门ID%Type,
    计费状态_IN     病人医嘱发送.计费状态%Type,
    First_IN        Number:=0
--参数：First_IN=表示是否一组医嘱的第一医嘱行,以便处理医嘱相关内容(如成药,配方的第一行,因为给药途径,配方煎法,用法可能为叮嘱不发送)
--      发送数次_IN,首次时间_IN,末次时间_IN:对"持续性"长嘱,不填写发送数次,可填写首末次时间(用于回退)。
) IS
    --包含病人及医嘱(一组医嘱中第一行)相关信息的游标
    Cursor c_Advice is
        Select 
            Nvl(A.相关ID,A.ID) AS 组ID,A.序号,A.病人ID,A.主页ID,A.婴儿,B.姓名,B.当前科室ID,C.操作类型,
            A.诊疗类别,A.医嘱期效,A.医嘱状态,A.医嘱内容,A.开嘱医生,A.开嘱时间,A.开始执行时间,A.上次执行时间,A.执行终止时间,
            A.执行时间方案,A.频率次数,A.频率间隔,A.间隔单位
        From 病人医嘱记录 A,病人信息 B,诊疗项目目录 C
        Where A.病人ID=B.病人ID And A.诊疗项目ID=C.ID And A.ID=医嘱ID_IN
        Group BY Nvl(A.相关ID,A.ID),A.序号,A.病人ID,A.主页ID,A.婴儿,B.姓名,B.当前科室ID,C.操作类型,A.诊疗类别,A.医嘱期效,
            A.医嘱状态,A.医嘱内容,A.开嘱医生,A.开嘱时间,A.开始执行时间,A.上次执行时间,A.执行终止时间,
            A.执行时间方案,A.频率次数,A.频率间隔,A.间隔单位;
    r_Advice c_Advice%RowType;

    --包含病人(婴儿)的所有未停长嘱(含配方长嘱)
    Cursor c_NeedStop(
        v_病人ID    病人医嘱记录.病人ID%Type,
        v_主页ID    病人医嘱记录.主页ID%Type,
        v_婴儿      病人医嘱记录.婴儿%Type,
        v_StopTime  Date) is
        Select ID From 病人医嘱记录 
        Where 病人ID=v_病人ID And 主页ID=v_主页ID And Nvl(婴儿,0)=Nvl(v_婴儿,0) 
            And Nvl(医嘱期效,0)=0 And 医嘱状态 Not IN(1,2,4,8,9)
			And 开始执行时间<v_StopTime
        Order BY 序号;
	--包含病人(婴儿)的已停但未确认的长嘱,终止执行时间在指定时间之后
    Cursor c_HaveStop(
        v_病人ID    病人医嘱记录.病人ID%Type,
        v_主页ID    病人医嘱记录.主页ID%Type,
        v_婴儿		病人医嘱记录.婴儿%Type,
        v_StopTime	Date) is
        Select ID From 病人医嘱记录 
        Where 病人ID=v_病人ID And 主页ID=v_主页ID And Nvl(婴儿,0)=Nvl(v_婴儿,0) 
            And Nvl(医嘱期效,0)=0 And 医嘱状态=8 And 执行终止时间>v_StopTime
			And 开始执行时间<v_StopTime
        Order BY 序号;

    --其它临时变量
    v_持续性        Number(1);--是否持续性长嘱
    v_AutoStop		Number(1);
    v_Temp			Varchar2(255);
    v_人员编号		病人费用记录.操作员编号%Type;
    v_人员姓名		病人费用记录.操作员姓名%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --当前操作人员    
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --是一组医嘱的第一行时处理医嘱内容
    If Nvl(First_IN,0)=1 Then
        Open c_Advice;
        Fetch c_Advice Into r_Advice;
        
        --并发操作检查
        ---------------------------------------------------------------------------------------
        IF Nvl(r_Advice.医嘱状态,0)=4 Then
            --检查要发送的医嘱是否被作废
            v_Error:='"'||r_Advice.姓名||'"的医嘱"'||r_Advice.医嘱内容||'"已经被其他人作废。'
                ||CHR(13)||CHR(10)||'该病人的医嘱发送失败。请重新读取发送清单再试。';
            Raise Err_Custom;
        End IF;

        If Nvl(r_Advice.医嘱期效,0)=0 And r_Advice.诊疗类别<>'7' Then
            --长嘱：含成药长嘱,非药"可选频率"长嘱,非药"持续性"长嘱
            
            --检查长嘱是否已被发送
            If r_Advice.上次执行时间 IS Not NULL Then
                If r_Advice.上次执行时间>=首次时间_IN Then
                    v_Error:='"'||r_Advice.姓名||'"的医嘱"'||r_Advice.医嘱内容||'"已经被其他人发送。'
                        ||CHR(13)||CHR(10)||'该病人的医嘱发送失败。请重新读取发送清单再试。';
                    Raise Err_Custom;
                End IF;
            End IF;

            --检查长嘱发送前是否已被自动停止(如术后)
            If r_Advice.执行终止时间 Is Not NULL Then
                If 首次时间_IN>r_Advice.执行终止时间 Then
                    v_Error:='"'||r_Advice.姓名||'"的医嘱"'||r_Advice.医嘱内容||'"已经被停止。'
                        ||CHR(13)||CHR(10)||'该病人的医嘱发送失败。请重新读取发送清单再试。';
                    Raise Err_Custom;
                End IF;
            End IF;
        ElsIF Nvl(r_Advice.医嘱状态,0) IN(8,9) Then
            --临嘱：含配方长嘱

            --检查是否已被发送(或因其它原因自动停止)
            v_Error:='"'||r_Advice.姓名||'"的医嘱"'||r_Advice.医嘱内容||'"已经被其他人发送。'
                ||CHR(13)||CHR(10)||'该病人的医嘱发送失败。请重新读取发送清单再试。';
            Raise Err_Custom;
        End IF;
        
        --发送后的医嘱处理
        ---------------------------------------------------------------------------------------
        If Nvl(r_Advice.医嘱期效,0)=0 And r_Advice.诊疗类别<>'7' Then
            --长期医嘱:更新上次执行时间
            Update 病人医嘱记录 
                Set 上次执行时间=末次时间_IN 
            Where ID=r_Advice.组ID Or 相关ID=r_Advice.组ID;
            
            --判断是否持续性长嘱
            v_持续性:=0;
            If r_Advice.执行时间方案 IS NULL 
                And (Nvl(r_Advice.频率次数,0)=0 Or Nvl(r_Advice.频率间隔,0)=0 Or r_Advice.间隔单位 IS NULL) Then
                v_持续性:=1;
            End IF;

            --预定了终止时间且未停止的自动停止
            IF r_Advice.执行终止时间 IS Not NULL And Nvl(r_Advice.医嘱状态,0) Not IN(8,9) Then
                v_AutoStop:=0;
                If v_持续性=1 Then
                    --非药"持续性"长嘱
                    If Trunc(末次时间_IN)=Trunc(r_Advice.执行终止时间-1) Then
                        v_AutoStop:=1; --终止这天不执行
                    End IF;
                ElsIf zl_AdviceNextTime(医嘱ID_IN)>r_Advice.执行终止时间 Then
                    --成药长嘱或非药"可选频率"长嘱
                    v_AutoStop:=1; --如果是等于,还可以执行一次
                End IF;

                If v_AutoStop=1 Then
                    Update 病人医嘱记录 
                        Set 医嘱状态=8,
                            停嘱时间=末次时间_IN,
                            停嘱医生=r_Advice.开嘱医生
                    Where ID=r_Advice.组ID Or 相关ID=r_Advice.组ID;
                    
                    Insert Into 病人医嘱状态(
                        医嘱ID,操作类型,操作人员,操作时间) 
                    Select 
                        ID,8,v_人员姓名,发送时间_IN 
                    From 病人医嘱记录 
                    Where ID=r_Advice.组ID Or 相关ID=r_Advice.组ID;
                End IF;
            End IF;
        Else
            --临嘱(含配方长嘱):停止。(为防万一,配方长嘱也更新执行终止时间)
            Update 病人医嘱记录 
                Set 医嘱状态=8,
                    执行终止时间=末次时间_IN,--为一次性临嘱时没有
                    上次执行时间=末次时间_IN,--为一次性临嘱时没有
                    停嘱时间=发送时间_IN,
                    停嘱医生=r_Advice.开嘱医生
            Where ID=r_Advice.组ID Or 相关ID=r_Advice.组ID;
            
            Insert Into 病人医嘱状态(
                医嘱ID,操作类型,操作人员,操作时间) 
            Select 
                ID,8,v_人员姓名,发送时间_IN 
            From 病人医嘱记录 
            Where ID=r_Advice.组ID Or 相关ID=r_Advice.组ID;
        End IF;

        --特殊医嘱的处理
        ---------------------------------------------------------------------------------------
        If r_Advice.诊疗类别='Z' And Nvl(r_Advice.操作类型,'0')<>'0' Then
            --(1-留观;2-住院;)3-转科;4-术后(不发送);5-出院;6-转院,7-会诊

            --几种特殊医嘱要自动停止病人该医嘱之前(按时间算)所有未停的长嘱
            If r_Advice.操作类型 IN('3','5','6') Then
                For r_NeedStop IN c_NeedStop(r_Advice.病人ID,r_Advice.主页ID,r_Advice.婴儿,r_Advice.开始执行时间) Loop
                    Update 病人医嘱记录
                        Set 医嘱状态=8,
							执行终止时间=Decode(Sign(开始执行时间-r_Advice.开始执行时间),1,开始执行时间,r_Advice.开始执行时间),
                            停嘱时间=发送时间_IN,
                            停嘱医生=r_Advice.开嘱医生
                    Where ID=r_NeedStop.ID;

                    Insert Into 病人医嘱状态(
                        医嘱ID,操作类型,操作人员,操作时间) 
                    Select 
                        ID,8,v_人员姓名,发送时间_IN 
                    From 病人医嘱记录 
                    Where ID=r_NeedStop.ID;
                End Loop;
				--已停止未确认的长嘱,终止时间在术后开始后的,调前其终止时间(同时多个术后或特殊医嘱的情况)
				For r_HaveStop IN c_HaveStop(r_Advice.病人ID,r_Advice.主页ID,r_Advice.婴儿,r_Advice.开始执行时间) Loop
					Update 病人医嘱记录
						Set 执行终止时间=Decode(Sign(开始执行时间-r_Advice.开始执行时间),1,开始执行时间,r_Advice.开始执行时间),
							停嘱时间=发送时间_IN,
							停嘱医生=r_Advice.开嘱医生
					Where ID=r_HaveStop.ID;
					
					Update 病人医嘱状态
						Set 操作时间=发送时间_IN,
							操作人员=v_人员姓名
					Where 医嘱ID=r_HaveStop.ID And 操作类型=8;
				End Loop;
            End IF;

            --具体的特殊处理
			If Nvl(r_Advice.婴儿,0)=0 Then
				If r_Advice.操作类型='3' And 执行部门ID_IN IS Not NULL 
					And r_Advice.当前科室ID IS Not NULL And Nvl(r_Advice.当前科室ID,0)<>Nvl(执行部门ID_IN,0) Then
					--转科医嘱,将病人登记转科到"执行科室ID"(在院病人且当前科室与转入科室不同才处理)
					zl_病人变动记录_Change(r_Advice.病人ID,r_Advice.主页ID,执行部门ID_IN,v_人员编号,v_人员姓名);
				ElsIf r_Advice.操作类型='5' Then
					--出院医嘱,将病人标记为预出院
					ZL_病人变动记录_PreOut(r_Advice.病人ID,r_Advice.主页ID,r_Advice.开始执行时间);
				ElsIf r_Advice.操作类型='6' Then
					--转院医嘱,将病人标记为预出院
					ZL_病人变动记录_PreOut(r_Advice.病人ID,r_Advice.主页ID,r_Advice.开始执行时间);
				End IF;
			End IF;
        End IF;

        Close c_Advice;
    End IF;
    
    --填写发送记录
    ---------------------------------------------------------------------------------------
    Insert Into 病人医嘱发送(
        医嘱ID,发送号,记录性质,NO,记录序号,发送数次,发送人,发送时间,执行状态,执行部门ID,计费状态,首次时间,末次时间)
    Values(
        医嘱ID_IN,发送号_IN,记录性质_IN,NO_IN,记录序号_IN,发送数次_IN,
        v_人员姓名,发送时间_IN,执行状态_IN,执行部门ID_IN,计费状态_IN,
        首次时间_IN,末次时间_IN);

	--自动填为已执行时，需要同步处理费用执行状态及审核划价状态
	If 执行状态_IN=1 Then
		ZL_病人医嘱执行_Finish(医嘱ID_IN,发送号_IN);
	End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱发送_Insert;
/

--数据转移过程调整(签名数据随对应的医嘱数据同步转移)
Create Or Replace Procedure Zl1_Datamoveout1(d_Demoded In Number) As
   --------------------------------------------
   --参数:d_Demoded,转出数据必须是多少天以前的数据
   --------------------------------------------
   d_Current Date;
   v_Version Varchar2(128);

   --------------------------------------------
   --转移指定ID的病人预交记录子过程；
   --------------------------------------------
   Procedure Zl_Move_Prepay(n_Settle_Id 病人预交记录.结帐id%Type) As
   Begin
      For r_Rec In (Select * From 病人预交记录 Where 结帐id = n_Settle_Id) Loop
         Update 人员缴款余额
         Set 余额 = Nvl(余额, 0) - Nvl(Decode(r_Rec.记录性质, 1, r_Rec.金额, 11, r_Rec.金额, r_Rec.冲预交), 0)
         Where 收款员 = r_Rec.操作员姓名 And 结算方式 = r_Rec.结算方式 And 性质 = 0;
         If Sql%Rowcount = 0 Then
            Insert Into 人员缴款余额
               (收款员, 结算方式, 性质, 余额)
            Values
               (r_Rec.操作员姓名, r_Rec.结算方式, 0,
                -1 * Nvl(Decode(r_Rec.记录性质, 1, r_Rec.金额, 11, r_Rec.金额, r_Rec.冲预交), 0));
         End If;
      End Loop;
      Insert Into H病人预交记录
         (Id, 记录性质, No, 实际票号, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 金额, 结算方式,
          结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id)
         Select Id, 记录性质, No, 实际票号, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 金额, 结算方式,
                结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id
         From 病人预交记录
         Where 结帐id = n_Settle_Id;
      Delete 病人预交记录 Where 结帐id = n_Settle_Id;
   End Zl_Move_Prepay;

   --------------------------------------------
   --转移指定ID的病人费用记录子过程；
   --------------------------------------------
   Procedure Zl_Move_Fee(n_Settle_Id 病人费用记录.结帐id%Type) As
   Begin
      Insert Into H病人费用记录
         (Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志,
          记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口,
          数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人,
          发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id,
          保险项目否, 保险编码, 统筹金额, 是否上传, 摘要, 是否急诊)
         Select Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号,
                门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位,
                付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人,
                开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名,
                结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 统筹金额, 是否上传, 摘要, 是否急诊
         From 病人费用记录
         Where 结帐id = n_Settle_Id;
      Delete 病人费用记录 Where 结帐id = n_Settle_Id;
   End Zl_Move_Fee;

   --------------------------------------------
   --转移指定ID的药品收发记录子过程；
   --------------------------------------------
   Procedure Zl_Move_Medilist(n_Rec_Id 药品收发记录.Id%Type) As
      r_Rec 药品收发记录%Rowtype;
   Begin
      Select * Into r_Rec From 药品收发记录 Where Id = n_Rec_Id;
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) - r_Rec.入出系数 * Nvl(r_Rec.实际数量 * r_Rec.付数, 0),
          实际数量 = Nvl(实际数量, 0) - r_Rec.入出系数 * Nvl(r_Rec.实际数量 * r_Rec.付数, 0),
          实际金额 = Nvl(实际金额, 0) - r_Rec.入出系数 * Nvl(r_Rec.零售金额, 0),
          实际差价 = Nvl(实际差价, 0) - r_Rec.入出系数 * Nvl(r_Rec.差价, 0), 上次供应商id = Nvl(上次供应商id, r_Rec.供药单位id),
          上次采购价 = Nvl(上次采购价, r_Rec.成本价), 上次批号 = Nvl(上次批号, r_Rec.批号), 上次产地 = Nvl(上次产地, r_Rec.产地),上次生产日期=nvl(上次生产日期,r_rec.生产日期)
      Where 库房id = r_Rec.库房id And 药品id = r_Rec.药品id And Nvl(批次, 0) = Nvl(r_Rec.批次, 0) And 性质 = 0;
      If Sql%Notfound Then
         Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 效期,
             上次产地,上次生产日期)
         Values
            (r_Rec.库房id, r_Rec.药品id, r_Rec.批次, 0, -r_Rec.入出系数 * Nvl(r_Rec.实际数量 * r_Rec.付数, 0),
             -r_Rec.入出系数 * Nvl(r_Rec.实际数量 * r_Rec.付数, 0), -r_Rec.入出系数 * Nvl(r_Rec.零售金额, 0),
             -r_Rec.入出系数 * Nvl(r_Rec.差价, 0), r_Rec.供药单位id, r_Rec.成本价, r_Rec.批号, r_Rec.效期, r_Rec.产地, r_Rec.生产日期);
      End If;
      Insert Into H药品收发记录
         (Id, 记录状态, 单据, No, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次,生产日期, 产地, 批号, 效期,
          付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 审核人,
          审核日期, 价格id, 费用id, 单量, 频次, 用法, 发药方式, 发药窗口, 配药日期, 外观, 产品合格证, 灭菌日期, 灭菌效期,领用人)
         Select Id, 记录状态, 单据, No, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次,生产日期, 产地, 批号, 效期,
                付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 审核人,
                审核日期, 价格id, 费用id, 单量, 频次, 用法, 发药方式, 发药窗口, 配药日期, 外观, 产品合格证, 灭菌日期, 灭菌效期,领用人
         From 药品收发记录
         Where Id = r_Rec.Id;
      Delete 药品收发记录 Where Id = r_Rec.Id;
   End Zl_Move_Medilist;

   --------------------------------------------
   --转移指定ID的病人病历记录子过程；
   --------------------------------------------
   Procedure Zl_Move_Cpr(n_Rec_Id 病人病历记录.Id%Type) As
   Begin
      Insert Into H病人病历记录
         (Id, 病人id, 主页id, 挂号单, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 书写人, 书写日期, 审阅人, 审阅日期, 归档人,
          归档日期, 作废人, 作废日期, 医嘱id)
         Select Id, 病人id, 主页id, 挂号单, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 书写人, 书写日期, 审阅人, 审阅日期, 归档人,
                归档日期, 作废人, 作废日期, 医嘱id
         From 病人病历记录
         Where Id = n_Rec_Id;
   
      Insert Into H病历打印记录
         (Id, 病历记录id, 起始页号, 结束页号, 起始位置, 结束位置, 打印时间, 打印人)
         Select Id, 病历记录id, 起始页号, 结束页号, 起始位置, 结束位置, 打印时间, 打印人
         From 病历打印记录
         Where 病历记录id = n_Rec_Id;
   
      Insert Into H病人病历修订记录
         (Id, 病历记录id, 书写人, 书写日期, 版本序号)
         Select Id, 病历记录id, 书写人, 书写日期, 版本序号 From 病人病历修订记录 Where 病历记录id = n_Rec_Id;
   
      Insert Into H病人病历内容
         (Id, 病历示范id, 病历记录id, 排列序号, 元素类型, 元素编码, 填写时机, 标题文本, 文本转储, 标题显示, 标题字体, 标题颜色,
          标题位置, 内容字体, 内容颜色, 内容位置, 嵌入方式, 病历修订id)
         Select Id, 病历示范id, 病历记录id, 排列序号, 元素类型, 元素编码, 填写时机, 标题文本, 文本转储, 标题显示, 标题字体,
                标题颜色, 标题位置, 内容字体, 内容颜色, 内容位置, 嵌入方式, 病历修订id
         From 病人病历内容
         Where 病历记录id = n_Rec_Id;
   
      Insert Into H病人病历文本段
         (病历id, 行号, 内容)
         Select 病历id, 行号, 内容 From 病人病历文本段 Where 病历id In (Select Id From 病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into H病人病历所见单
         (病历id, 控件号, 控件类, 标题, 固定行, 固定列, 列, 行, 宽, 高, 对齐, 合并号, 不可写, 可屏蔽, 所见项id, 数值类型,
          所见内容, 计量单位)
         Select 病历id, 控件号, 控件类, 标题, 固定行, 固定列, 列, 行, 宽, 高, 对齐, 合并号, 不可写, 可屏蔽, 所见项id, 数值类型,
                所见内容, 计量单位
         From 病人病历所见单
         Where 病历id In (Select Id From 病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into H病人病历标记图
         (病历id, 类型, 内容, 字体, 点集, X1, Y1, X2, Y2, 填充色, 填充方式, 线条色, 线型, 线宽)
         Select 病历id, 类型, 内容, 字体, 点集, X1, Y1, X2, Y2, 填充色, 填充方式, 线条色, 线型, 线宽
         From 病人病历标记图
         Where 病历id In (Select Id From 病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into H病人病历外部图
         (病历id, 序号, 图象类型, 图象路径, 图象文件)
         Select 病历id, 序号, 图象类型, 图象路径, 图象文件
         From 病人病历外部图
         Where 病历id In (Select Id From 病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into H病人过敏记录
         (Id, 病人id, 主页id, 记录来源, 病历id, 药物id, 药物名, 记录时间, 记录人, 结果)
         Select Id, 病人id, 主页id, 记录来源, 病历id, 药物id, 药物名, 记录时间, 记录人, 结果
         From 病人过敏记录
         Where 病历id In (Select Id From 病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into H病人诊断记录
         (Id, 病人id, 主页id, 记录来源, 病历id, 诊断类型, 疾病id, 诊断id, 证候id, 诊断描述, 出院情况, 是否未治, 是否疑诊,
          记录日期, 记录人, 取消时间, 取消人, 医嘱id, 诊断次序, 编码序号)
         Select Id, 病人id, 主页id, 记录来源, 病历id, 诊断类型, 疾病id, 诊断id, 证候id, 诊断描述, 出院情况, 是否未治, 是否疑诊,
                记录日期, 记录人, 取消时间, 取消人, 医嘱id, 诊断次序, 编码序号
         From 病人诊断记录
         Where 病历id In (Select Id From 病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into H病人家系图
         (病历id, 序号, 父亲, 母亲, 姓名, 称谓, 性别, 状态, 说明, 怀孕, 养育关系, 婚姻关系)
         Select 病历id, 序号, 父亲, 母亲, 姓名, 称谓, 性别, 状态, 说明, 怀孕, 养育关系, 婚姻关系
         From 病人家系图
         Where 病历id In (Select Id From 病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into H病人手麻记录
         (Id, 病人id, 主页id, 记录来源, 病历id, 手术日期, 手术开始时间, 手术结束时间, 拟行手术, 手术操作id, 诊疗项目id, 已行手术,
          主刀医师, 第一助手, 第二助手, 手术护士, 麻醉开始时间, 麻醉结束时间, 麻醉方式, 麻醉类型, 麻醉质量, 输液总量, 麻醉医师,
          输氧开始时间, 输氧结束时间, 切口, 愈合, 记录日期, 记录人, 取消时间, 取消人)
         Select Id, 病人id, 主页id, 记录来源, 病历id, 手术日期, 手术开始时间, 手术结束时间, 拟行手术, 手术操作id, 诊疗项目id,
                已行手术, 主刀医师, 第一助手, 第二助手, 手术护士, 麻醉开始时间, 麻醉结束时间, 麻醉方式, 麻醉类型, 麻醉质量,
                输液总量, 麻醉医师, 输氧开始时间, 输氧结束时间, 切口, 愈合, 记录日期, 记录人, 取消时间, 取消人
         From 病人手麻记录
         Where 病历id In (Select Id From 病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into H病人麻醉用药
         (记录id, 类型, 序号, 药物id, 药物名, 总量, 方式)
         Select 记录id, 类型, 序号, 药物id, 药物名, 总量, 方式
         From 病人麻醉用药
         Where 记录id In (Select Id From 病人手麻记录 Where 病历id In (Select Id From 病人病历内容 Where 病历记录id = n_Rec_Id));
      Delete 病人病历记录 Where Id = n_Rec_Id;
   End Zl_Move_Cpr;

   --------------------------------------------
   --转移指定ID的病人医嘱记录子过程；
   --------------------------------------------
   Procedure Zl_Move_Order(n_Rec_Id 病人医嘱记录.Id%Type) As
   Begin
      Insert Into H病人医嘱记录
         (Id, 相关id, 病人来源, 病人id, 主页id, 姓名, 性别, 年龄, 挂号单, 婴儿, 病人科室id, 序号, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id,
          标本部位, 收费细目id, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 执行科室id, 皮试结果, 执行频次, 频率次数, 频率间隔,
          间隔单位, 执行时间方案, 计价特性, 执行性质, 紧急标志, 开始执行时间, 执行终止时间, 上次执行时间, 开嘱科室id, 开嘱医生,
          开嘱时间, 校对护士, 校对时间, 停嘱医生, 停嘱时间, 确认停嘱时间, 申请id, 前提id, 是否上传, 天数,审查结果)
         Select Id, 相关id, 病人来源, 病人id, 主页id, 姓名, 性别, 年龄, 挂号单, 婴儿, 病人科室id, 序号, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id,
                标本部位, 收费细目id, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 执行科室id, 皮试结果, 执行频次, 频率次数, 频率间隔,
                间隔单位, 执行时间方案, 计价特性, 执行性质, 紧急标志, 开始执行时间, 执行终止时间, 上次执行时间, 开嘱科室id,
                开嘱医生, 开嘱时间, 校对护士, 校对时间, 停嘱医生, 停嘱时间, 确认停嘱时间, 申请id, 前提id, 是否上传, 天数,审查结果
         From 病人医嘱记录
         Where Id = n_Rec_Id;
   
      Insert Into H病人医嘱计价
         (医嘱id, 收费细目id, 数量, 单价, 从项)
         Select 医嘱id, 收费细目id, 数量, 单价, 从项 From 病人医嘱计价 Where 医嘱id = n_Rec_Id;
      Delete 病人医嘱计价 Where 医嘱id = n_Rec_Id;
      
	  --转出医嘱签名数据(因为多条医嘱共用一个签名，又是逐条转出，需要先禁用"FK_签名ID"外键)
      Insert Into H医嘱签名记录(
		ID, 签名规则, 签名信息, 证书ID, 签名时间, 签名人)
      Select 
	    ID, 签名规则, 签名信息, 证书ID, 签名时间, 签名人 
	  From 医嘱签名记录 
	  Where ID IN (Select 签名ID From 病人医嘱状态 Where 医嘱id = n_Rec_Id);
	  IF SQL%RowCount<>0 Then--可能转其它医嘱时已经删了
          Delete 医嘱签名记录 Where ID IN (Select 签名ID From 病人医嘱状态 Where 医嘱id = n_Rec_Id);
      End IF;
   
      Insert Into H病人医嘱状态
         (医嘱id, 操作类型, 操作人员, 操作时间, 签名ID)
         Select 医嘱id, 操作类型, 操作人员, 操作时间, 签名ID From 病人医嘱状态 Where 医嘱id = n_Rec_Id;
      Delete 病人医嘱状态 Where 医嘱id = n_Rec_Id;
   
      --医嘱发送部分   
      Insert Into H病人医嘱发送
         (医嘱id, 发送号, 记录性质, No, 记录序号, 发送数次, 发送人, 发送时间, 首次时间, 末次时间, 执行状态, 执行部门id, 计费状态,
          报告id, 执行间, 执行过程, 采样人, 采样时间, 样本条码)
         Select 医嘱id, 发送号, 记录性质, No, 记录序号, 发送数次, 发送人, 发送时间, 首次时间, 末次时间, 执行状态, 执行部门id,
                计费状态, 报告id, 执行间, 执行过程, 采样人, 采样时间, 样本条码
         From 病人医嘱发送
         Where 医嘱id = n_Rec_Id;
   
      Insert Into H病人医嘱附费
         (医嘱id, 发送号, 记录性质, No)
         Select 医嘱id, 发送号, 记录性质, No From 病人医嘱附费 Where 医嘱id = n_Rec_Id;
      Delete 病人医嘱附费 Where 医嘱id = n_Rec_Id;
   
      Insert Into H病人医嘱执行
         (医嘱id, 发送号, 要求时间, 本次数次, 执行摘要, 执行人, 执行时间, 登记人, 登记时间)
         Select 医嘱id, 发送号, 要求时间, 本次数次, 执行摘要, 执行人, 执行时间, 登记人, 登记时间
         From 病人医嘱执行
         Where 医嘱id = n_Rec_Id;
      Delete 病人医嘱执行 Where 医嘱id = n_Rec_Id;
   
      Insert Into H影像检查记录
         (医嘱id, 发送号, 影像类别, 检查号, 姓名, 英文名, 性别, 年龄, 出生日期, 身高, 体重, 病理检查, 发放胶片, 检查uid, 位置一,
          位置二, 位置三, 检查设备, 报告图象, 接收日期)
         Select 医嘱id, 发送号, 影像类别, 检查号, 姓名, 英文名, 性别, 年龄, 出生日期, 身高, 体重, 病理检查, 发放胶片, 检查uid,
                位置一, 位置二, 位置三, 检查设备, 报告图象, 接收日期
         From 影像检查记录
         Where 医嘱id = n_Rec_Id;
      For r_Ris In (Select 检查uid
                    From 影像检查记录 r, 病人医嘱发送 s
                    Where r.医嘱id = s.医嘱id And r.发送号 = s.发送号 And s.医嘱id = n_Rec_Id) Loop
         Insert Into H影像检查序列
            (序列uid, 检查uid, 序列号, 序列描述, 采集时间)
            Select 序列uid, 检查uid, 序列号, 序列描述, 采集时间 From 影像检查序列 Where 检查uid = r_Ris.检查uid;
         For r_Seq In (Select 序列uid From 影像检查序列 Where 检查uid = r_Ris.检查uid) Loop
            Insert Into H影像检查图象
               (图像uid, 序列uid, 图像号, 图像描述)
               Select 图像uid, 序列uid, 图像号, 图像描述 From 影像检查图象 Where 序列uid = r_Seq.序列uid;
            Delete 影像检查图象 Where 序列uid = r_Seq.序列uid;
         End Loop;
         Delete 影像检查序列 Where 检查uid = r_Ris.检查uid;
      End Loop;
      Delete 影像检查记录 Where 医嘱id = n_Rec_Id;
   
      Delete 病人医嘱发送 Where 医嘱id = n_Rec_Id;
   
      --手术记录部分
      Insert Into H病人手术记录
         (Id, 医嘱id, 病历id, 病人id, 主页id, 记录来源, 手术日期, 手术开始时间, 手术结束时间, 手术规模, 麻醉开始时间,
          麻醉结束时间, 麻醉方式, 麻醉类型, 麻醉质量, 输液总量, 输氧开始时间, 输氧结束时间, 切口, 愈合, 无菌手术, 手术间, 手术间id,
          手术室id, 记录日期, 记录人, 取消时间, 取消人)
         Select Id, 医嘱id, 病历id, 病人id, 主页id, 记录来源, 手术日期, 手术开始时间, 手术结束时间, 手术规模, 麻醉开始时间,
                麻醉结束时间, 麻醉方式, 麻醉类型, 麻醉质量, 输液总量, 输氧开始时间, 输氧结束时间, 切口, 愈合, 无菌手术, 手术间,
                手术间id, 手术室id, 记录日期, 记录人, 取消时间, 取消人
         From 病人手术记录
         Where 医嘱id = n_Rec_Id;
      For r_Ops In (Select Id From 病人手术记录 Where 医嘱id = n_Rec_Id) Loop
         Insert Into H病人手术情况
            (记录id, 性质, 缺省, 手术名称, 手术操作id, 诊疗项目id)
            Select 记录id, 性质, 缺省, 手术名称, 手术操作id, 诊疗项目id From 病人手术情况 Where 记录id = r_Ops.Id;
         Delete 病人手术情况 Where 记录id = r_Ops.Id;
      
         Insert Into H病人手术用药
            (记录id, 类型, 序号, 药品id, 药名id, 药品名称, 准备总量, 使用总量, 给药方式, 执行科室id)
            Select 记录id, 类型, 序号, 药品id, 药名id, 药品名称, 准备总量, 使用总量, 给药方式, 执行科室id
            From 病人手术用药
            Where 记录id = r_Ops.Id;
         Delete 病人手术用药 Where 记录id = r_Ops.Id;
      
         Insert Into H病人手术人员
            (记录id, 科室id, 岗位, 人员id, 编码, 姓名)
            Select 记录id, 科室id, 岗位, 人员id, 编码, 姓名 From 病人手术人员 Where 记录id = r_Ops.Id;
         Delete 病人手术人员 Where 记录id = r_Ops.Id;
      
         Insert Into H病人手术材料
            (记录id, 序号, 材料id, 准备数量, 实用数量, 清点结果, 执行科室id, 附加说明)
            Select 记录id, 序号, 材料id, 准备数量, 实用数量, 清点结果, 执行科室id, 附加说明
            From 病人手术材料
            Where 记录id = r_Ops.Id;
         Delete 病人手术材料 Where 记录id = r_Ops.Id;
      
         Insert Into H病人手术计价
            (记录id, 序号, 数量, 单价, 收费细目id, 执行科室id)
            Select 记录id, 序号, 数量, 单价, 收费细目id, 执行科室id From 病人手术计价 Where 记录id = r_Ops.Id;
         Delete 病人手术计价 Where 记录id = r_Ops.Id;
      
      End Loop;
      Delete 病人手术记录 Where 医嘱id = n_Rec_Id;
   
      Insert Into H病人手术状态
         (医嘱id, 序号, 上次序号, 处理性质, 记录状态, 处理人, 处理时间, 附加说明, 当前状态, 单据号)
         Select 医嘱id, 序号, 上次序号, 处理性质, 记录状态, 处理人, 处理时间, 附加说明, 当前状态, 单据号
         From 病人手术状态
         Where 医嘱id = n_Rec_Id;
      Delete 病人手术状态 Where 医嘱id = n_Rec_Id;
   
      --检验部分
      Insert Into H检验标本记录
         (Id, 医嘱id, 标本序号, 采样时间, 采样人, 标本类型, 核收人, 核收时间, 样本状态, 检验人, 检验时间, 审核人, 审核时间,
          合并报告号, 打印次数, 申请类型, 仪器id, 样本条码, 报告结果, 备注, 未通过审核原因, 申请时间, 标本形态, 是否质控品,
          执行科室id)
         Select Id, 医嘱id, 标本序号, 采样时间, 采样人, 标本类型, 核收人, 核收时间, 样本状态, 检验人, 检验时间, 审核人, 审核时间,
                合并报告号, 打印次数, 申请类型, 仪器id, 样本条码, 报告结果, 备注, 未通过审核原因, 申请时间, 标本形态, 是否质控品,
                执行科室id
         From 检验标本记录
         Where 医嘱id = n_Rec_Id;
      For r_Retu In (Select Id From 检验普通结果 Where 检验标本id In (Select Id From 检验标本记录 Where 医嘱id = n_Rec_Id)) Loop
         Insert Into H检验普通结果
            (Id, 检验标本id, 检验项目id, 检验结果, 结果标志, 结果参考, 修改者, 修改时间, 记录类型, 原始结果, 原始记录时间, 记录者,
             是否检验, 修改原因, 细菌id, 仪器id, 培养描述)
            Select Id, 检验标本id, 检验项目id, 检验结果, 结果标志, 结果参考, 修改者, 修改时间, 记录类型, 原始结果, 原始记录时间,
                   记录者, 是否检验, 修改原因, 细菌id, 仪器id, 培养描述
            From 检验普通结果
            Where Id = r_Retu.Id;
         Insert Into H检验药敏结果
            (细菌结果id, 抗生素id, 修改者, 修改时间, 结果, 结果类型, 记录类型, 仪器id)
            Select 细菌结果id, 抗生素id, 修改者, 修改时间, 结果, 结果类型, 记录类型, 仪器id
            From 检验药敏结果
            Where 细菌结果id = r_Retu.Id;

		 Delete 检验药敏结果 Where 细菌结果id = r_Retu.Id;
         Delete 检验普通结果 Where Id = r_Retu.Id;
      End Loop;
      Delete 检验标本记录 Where 医嘱id = n_Rec_Id;
   
      Delete 病人医嘱记录 Where Id = n_Rec_Id;
   End Zl_Move_Order;

   --------------------------------------------
   --以下为主程序体
   --------------------------------------------
Begin
   --防止数据不一致，先禁用一些约束
   Begin
      Execute Immediate 'Alter Table 病人医嘱记录 Modify Constraint 病人医嘱记录_FK_申请ID Disable';
      Execute Immediate 'Alter Table 病人医嘱发送 Modify Constraint 病人医嘱发送_FK_报告ID Disable';
	  Execute Immediate 'Alter Table 病人医嘱状态 Modify Constraint 病人医嘱状态_FK_签名ID Disable';
   Exception
      When Others Then Null;
   End;
   
   Select Trunc(Sysdate) Into d_Current From Dual;

   --------------------------------------------
   --1、指定时间前可转出的没有冲预交收费和对应发药记录转移
   For r_Settle In (Select l.结帐id
                    From 病人预交记录 l,
                         (Select 结帐id From 病人预交记录 Where 收款时间 < d_Current - d_Demoded And 记录性质 In (3, 4, 5)) c
                    Where l.结帐id = c.结帐id
                    Group By l.结帐id
                    Having Sum(Decode(记录性质, 1, 1, 11, 1, 0)) = 0
					Union
					Select 结帐id From 病人费用记录 
					Where 登记时间 < d_Current - d_Demoded And 记录性质=4 
						And Nvl(应收金额,0)=0 And Nvl(实收金额,0)=0
                    Minus
                    Select Distinct d.结帐id
                    From 药品收发记录 l,
                         (Select d.No, d.Id, d.记录性质, d.结帐id
                           From 病人费用记录 d
                           Where d.登记时间 < d_Current - d_Demoded And d.记录性质 = 1 And d.收费类别 In ('4', '5', '6', '7')) d
                    Where l.No = d.No And l.费用id = d.Id And Nvl(发药方式, 0) <> -1 And
                          (l.审核日期 >= d_Current - d_Demoded Or l.审核日期 Is Null) And l.单据 In (8, 24)) Loop
   
      Zl_Move_Prepay(r_Settle.结帐id);
      For r_Rxlist In (Select m.Id
                       From 药品收发记录 m,
                            (Select Id, No, 序号, 记录性质
                              From 病人费用记录
                              Where 结帐id = r_Settle.结帐id And 收费类别 In ('4', '5', '6', '7') And 记录性质 In (1, 2)) e
                       Where m.No = e.No And m.费用id = e.Id And
                             (e.记录性质 = 1 And m.单据 In (8, 24) Or e.记录性质 <> 1 And m.单据 In (9, 10, 25, 26))) Loop
         Zl_Move_Medilist(r_Rxlist.Id);
      End Loop;
      Zl_Move_Fee(r_Settle.结帐id);
   
      Commit;
   End Loop;

   --------------------------------------------
   --2、指定时间前可转出的结帐记录、病人预交记录、病人记帐费用和对应记帐发药记录转移
   For r_Settle In (Select 结帐id
                    From 病人费用记录
                    Where 登记时间 < d_Current - d_Demoded And 记录性质 In (1, 4, 5) And Nvl(记帐费用,0) <> 1
                    Union
                    Select Id As 结帐id
                    From 病人结帐记录 l
                    Where l.收费时间 < d_Current - d_Demoded
                    Minus
                    Select Distinct d.结帐id
                    From 病人预交记录 d,
                         (Select d.No
                           From 病人预交记录 d,
                                (Select 结帐id
                                  From 病人费用记录
                                  Where 登记时间 < d_Current - d_Demoded And 记录性质 In (1, 4, 5) And Nvl(记帐费用,0) <> 1
                                  Union
                                  Select Id As 结帐id From 病人结帐记录 Where 收费时间 < d_Current - d_Demoded) l
                           Where d.结帐id = l.结帐id And d.记录性质 In (1, 11)
                           Group By d.No
                           Having d.No Is Not Null And Sum(d.金额) - Sum(d.冲预交) <> 0) n
                    Where d.No = n.No And d.记录性质 In (1, 11)
                    Minus
                    Select Distinct d.结帐id
                    From 病人费用记录 d,
                         (Select d.No, d.序号, Decode(d.记录性质, 12, 2, 13, 3, 15, 5, d.记录性质) As 记录性质
                           From 病人费用记录 d, 病人结帐记录 l
                           Where d.结帐id = l.Id And l.收费时间 < d_Current - d_Demoded And d.记录性质 In (2, 12, 3, 13, 5, 15) And
                                 d.记帐费用 = 1
                           Group By d.No, d.序号, Decode(d.记录性质, 12, 2, 13, 3, 15, 5, d.记录性质)
                           Having d.No Is Not Null And d.序号 Is Not Null And Nvl(Sum(d.实收金额),0) - Nvl(Sum(d.结帐金额),0) <> 0) n
                    Where d.No = n.No And d.序号 = n.序号 And Decode(d.记录性质, 12, 2, 13, 3, 15, 5, d.记录性质) = n.记录性质
                    Minus
                    Select Distinct d.结帐id
                    From 药品收发记录 l,
                         (Select d.No, d.Id, d.记录性质, d.结帐id
                           From 病人费用记录 d
                           Where d.登记时间 < d_Current - d_Demoded And d.记录性质 In (1, 2) And d.收费类别 In ('4', '5', '6', '7')) d
                    Where l.No = d.No And l.费用id = d.Id And Nvl(发药方式, 0) <> -1 And
                          (l.审核日期 >= d_Current - d_Demoded Or l.审核日期 Is Null) And
                          (d.记录性质 = 1 And l.单据 In (8, 24) Or d.记录性质 <> 1 And l.单据 In (9, 10, 25, 26))) Loop
   
      Insert Into H病人结帐记录
         (Id, No, 实际票号, 记录状态, 病人id, 操作员编号, 操作员姓名, 收费时间, 开始日期, 结束日期)
         Select Id, No, 实际票号, 记录状态, 病人id, 操作员编号, 操作员姓名, 收费时间, 开始日期, 结束日期
         From 病人结帐记录
         Where Id = r_Settle.结帐id;
   
      Zl_Move_Prepay(r_Settle.结帐id);
      For r_Rxlist In (Select m.Id
                       From 药品收发记录 m,
                            (Select Id, No, 序号, 记录性质
                              From 病人费用记录
                              Where 结帐id = r_Settle.结帐id And 收费类别 In ('4', '5', '6', '7') And 记录性质 In (1, 2)) e
                       Where m.No = e.No And m.费用id = e.Id And
                             (e.记录性质 = 1 And m.单据 In (8, 24) Or e.记录性质 <> 1 And m.单据 In (9, 10, 25, 26))) Loop
         Zl_Move_Medilist(r_Rxlist.Id);
      End Loop;
      Zl_Move_Fee(r_Settle.结帐id);
   
      Delete 病人结帐记录 Where Id = r_Settle.结帐id;
   
      Commit;
   End Loop;

   --------------------------------------------
   --3、指定时间前门诊就诊病人医嘱病历数据(前提是病人本次就诊费用已经结清转出)
   For r_Regist In (Select r.Id, r.病人id, r.No
                    From 病人挂号记录 r, (Select No From 病人费用记录 Where 登记时间 < d_Current - d_Demoded And 记录性质 = 4) d,
                         (Select r.Id As 挂号id
                           From 病人医嘱记录 a, 病人挂号记录 r
                           Where a.挂号单 = r.No And r.登记时间 < d_Current - d_Demoded
                           Group By r.Id
                           Having Max(a.停嘱时间) >= d_Current - d_Demoded) a,
                         (Select a.挂号id
                           From 病人费用记录 e,
                                (Select a.Id, r.Id As 挂号id
                                  From 病人医嘱记录 a, 病人挂号记录 r
                                  Where a.挂号单 = r.No And r.登记时间 < d_Current - d_Demoded) a
                           Where e.医嘱序号 = a.Id) e
                    Where r.No = d.No(+) And r.Id = a.挂号id(+) And r.Id = e.挂号id(+) 
						And r.执行状态<>2 And r.登记时间 < d_Current - d_Demoded
                    Group By r.Id, r.病人id, r.No
                    Having Count(d.No) = 0 And Count(a.挂号id) = 0 And Count(e.挂号id) = 0) Loop
   
      Insert Into H病人挂号记录
         (Id, No, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间,
          操作员编号, 操作员姓名, 摘要)
         Select Id, No, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间,
                登记时间, 操作员编号, 操作员姓名, 摘要
         From 病人挂号记录
         Where Id = r_Regist.Id;
   
      For r_Cpr In (Select Id From 病人病历记录 Where 病人id = r_Regist.病人id And 挂号单 = r_Regist.No) Loop
         Zl_Move_Cpr(r_Cpr.Id);
      End Loop;
   
      For r_Order In (Select Id From 病人医嘱记录 Where 病人id = r_Regist.病人id And 挂号单 = r_Regist.No) Loop
         Zl_Move_Order(r_Order.Id);
      End Loop;
   
      Delete 病人挂号记录 Where Id = r_Regist.Id;
   
      Commit;
   End Loop;

   --------------------------------------------
   --4、指定时间前出院的住院就诊病人医嘱病历数据(前提是病人本次就诊费用已经结清转出)
   For r_Page In (Select 病人id, 主页id
                  From 病案主页 p
                  Where 出院日期 < d_Current - d_Demoded And Nvl(数据转出, 0) <> 1 And Not Exists
                   (Select 1 From 病人费用记录 Where 病人id = p.病人id And 主页id = p.主页id)) Loop
   
      For r_Cpr In (Select Id From 病人病历记录 Where 病人id = r_Page.病人id And 主页id = r_Page.主页id) Loop
         Zl_Move_Cpr(r_Cpr.Id);
      End Loop;
   
      For r_Order In (Select Id From 病人医嘱记录 Where 病人id = r_Page.病人id And 主页id = r_Page.主页id) Loop
         Zl_Move_Order(r_Order.Id);
      End Loop;
   
      Update 病案主页 Set 数据转出 = 1 Where 病人id = r_Page.病人id And 主页id = r_Page.主页id;
   
      Commit;
   End Loop;

   --------------------------------------------
   --5、指定时间前人员缴款记录的转移
   For r_Hand In (Select * From 人员缴款记录 Where 登记时间 < d_Current - d_Demoded) Loop
      Insert Into H人员缴款记录
         (Id, 单据id, 收款员, 结算方式, 结算号, 金额, 摘要, 截止时间, 登记时间, 登记人)
         Select Id, 单据id, 收款员, 结算方式, 结算号, 金额, 摘要, 截止时间, 登记时间, 登记人 
         From 人员缴款记录
         Where Id = r_Hand.Id;
   
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + r_Hand.金额
      Where 收款员 = r_Hand.收款员 And 结算方式 = r_Hand.结算方式 And 性质 = 0;
      If Sql%Rowcount = 0 Then
         Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (r_Hand.收款员, r_Hand.结算方式, 0, r_Hand.金额);
      End If;
   
      If r_Hand.Id = r_Hand.单据id Then
         Insert Into H人员缴款对照
            (单据id, 性质, 记录id)
            Select 单据id, 性质, 记录id From 人员缴款对照 Where 单据id = r_Hand.单据id;
         Delete From 人员缴款对照 Where 单据id = r_Hand.单据id;
      End If;
   
      Delete 人员缴款记录 Where Id = r_Hand.Id;
      Commit;
   End Loop;

   --------------------------------------------
   --6、指定时间前用完的票据数据的转移
   For r_Bill In (Select d.领用id As Id
                  From 票据使用明细 d, (Select Id From 票据领用记录 Where 登记时间 < d_Current - d_Demoded And 剩余数量 = 0) l
                  Where d.领用id = l.Id
                  Group By d.领用id
                  Having Max(d.使用时间) < d_Current - d_Demoded) Loop
      Insert Into H票据领用记录
         (Id, 票种, 领用人, 前缀文本, 开始号码, 终止号码, 使用方式, 登记时间, 登记人, 当前号码, 剩余数量)
         Select Id, 票种, 领用人, 前缀文本, 开始号码, 终止号码, 使用方式, 登记时间, 登记人, 当前号码, 剩余数量
         From 票据领用记录
         Where Id = r_Bill.Id;
   
      For r_Used In (Select Id, 打印id From 票据使用明细 Where 领用id = r_Bill.Id) Loop
         Insert Into H票据使用明细
            (Id, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
            Select Id, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人 From 票据使用明细 Where Id = r_Used.Id;
      
         Insert Into H票据打印内容
            (Id, 数据性质, No)
            Select Id, 数据性质, No From 票据打印内容 Where Id = r_Used.打印id;
         Delete 票据打印内容 Where Id = r_Used.打印id;
      
         Delete 票据使用明细 Where Id = r_Used.Id;
      End Loop;
   
      Delete 票据领用记录 Where Id = r_Bill.Id;
      Commit;
   End Loop;

   --   --------------------------------------------
   --   --7、指定时间前已审核的药品入出记录
   --   For r_Flow In (Select Id From 药品收发记录 Where 审核日期 < d_Current - d_Demoded And 单据 Not In (8, 9, 10, 24, 25, 26)) Loop
   --      Zl_Move_Medilist(r_Flow.Id);
   --      Commit;
   --   End Loop;
   --
   --   --------------------------------------------
   --   --8、指定时间前已经正确勾对结清的应付款和付款记录转移
   --   For r_Clear In (Select s.付款序号
   --                   From 付款记录 p, 应付记录 m, (Select Distinct 付款序号 From 付款记录 Where 审核日期 < d_Current - d_Demoded) s
   --                   Where p.付款序号 = s.付款序号 And m.付款序号 = s.付款序号
   --                   Group By s.付款序号
   --                   Having Max(p.审核日期) < d_Current - d_Demoded And Max(m.审核日期) < d_Current - d_Demoded) Loop
   --
   --      Insert Into H应付记录
   --         (Id, 记录性质, 记录状态, No, 收发id, 单位id, 品名, 规格, 产地, 批号, 计量单位, 入库单据号, 单据金额, 数量, 采购价,
   --          采购金额, 发票号, 发票日期, 发票金额, 制定日期, 计划金额, 计划人, 计划日期, 填制人, 填制日期, 审核人, 审核日期, 摘要,
   --          付款序号, 计划序号, 系统标识)
   --         Select Id, 记录性质, 记录状态, No, 收发id, 单位id, 品名, 规格, 产地, 批号, 计量单位, 入库单据号, 单据金额, 数量, 采购价,
   --                采购金额, 发票号, 发票日期, 发票金额, 制定日期, 计划金额, 计划人, 计划日期, 填制人, 填制日期, 审核人, 审核日期,
   --                摘要, 付款序号, 计划序号, 系统标识
   --         From 应付记录
   --         Where 付款序号 = r_Clear.付款序号;
   --      Delete 应付记录 Where 付款序号 = r_Clear.付款序号;
   --
   --      Insert Into H付款记录
   --         (Id, 记录状态, No, 序号, 预付款, 单位id, 金额, 结算方式, 结算号码, 摘要, 填制人, 填制日期, 审核人, 审核日期, 付款序号)
   --         Select Id, 记录状态, No, 序号, 预付款, 单位id, 金额, 结算方式, 结算号码, 摘要, 填制人, 填制日期, 审核人, 审核日期,
   --                付款序号
   --         From 付款记录
   --         Where 付款序号 = r_Clear.付款序号;
   --      Delete 付款记录 Where 付款序号 = r_Clear.付款序号;
   --
   --      Commit;
   --   End Loop;

   --------------------------------------------
   --9、在线数据表索引重建:Oracle 8.1.6/8.1.7版本需要手工执行重建索引
   --启用约束
   Begin
      Execute Immediate 'Alter Table 病人医嘱记录 Modify Constraint 病人医嘱记录_FK_申请ID Enable';
      Execute Immediate 'Alter Table 病人医嘱发送 Modify Constraint 病人医嘱发送_FK_报告ID Enable';
	  Execute Immediate 'Alter Table 病人医嘱状态 Modify Constraint 病人医嘱状态_FK_签名ID Enable';
   Exception
      When Others Then Null;
   End;

   Select Version Into v_Version From Product_Component_Version Where Upper(Substr(PRODUCT,1,6))=Upper('Oracle');
   If v_Version>='9' Then
	   For r_Sql In (Select 'alter index ' || Index_Name || ' rebuild online nologging' As Sqltext
					 From User_Indexes
					 Where Table_Name In (Select Substr(Table_Name, 2) From User_Tables Where Table_Name Like 'H%')) Loop
		  Execute Immediate r_Sql.Sqltext;
	   End Loop;
   End If;

   --------------------------------------------
   Update Zldatamove
   Set 上次日期 = Greatest(d_Current - d_Demoded, Nvl(上次日期, d_Current - d_Demoded))
   Where 系统 In (Select 编号 From Zlsystems Where Upper(所有者) = Zl_Owner And 编号 Like '1%') And 组号 = 1;
   Commit;
End Zl1_Datamoveout1;
/

--------------------------------------------
--数据返回过程2：抽选返回病人某次门诊住院医疗数据
--------------------------------------------
CREATE OR REPLACE Procedure Zl_Retu_Clinic(n_Patiid In Number, v_Times In Varchar2, n_Flag In Number) As
   --------------------------------------------
   --参数:n_Patiid,病人id
   --     v_Times,挂号单号或住院主页id
   --     n_Flag,门诊或住院标志:0-门诊,1-住院
   --------------------------------------------
   --------------------------------------------
   --返回指定ID的病人病历记录子过程；
   --------------------------------------------
   Procedure Zl_Retu_Cpr(n_Rec_Id H病人病历记录.Id%Type) As
   Begin
      Insert Into 病人病历记录
         (Id, 病人id, 主页id, 挂号单, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 书写人, 书写日期, 审阅人, 审阅日期, 归档人,
          归档日期, 作废人, 作废日期, 医嘱id)
         Select Id, 病人id, 主页id, 挂号单, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 书写人, 书写日期, 审阅人, 审阅日期, 归档人,
                归档日期, 作废人, 作废日期, 医嘱id
         From H病人病历记录
         Where Id = n_Rec_Id;
   
      Insert Into 病历打印记录
         (Id, 病历记录id, 起始页号, 结束页号, 起始位置, 结束位置, 打印时间, 打印人)
         Select Id, 病历记录id, 起始页号, 结束页号, 起始位置, 结束位置, 打印时间, 打印人
         From H病历打印记录
         Where 病历记录id = n_Rec_Id;
   
      Insert Into 病人病历修订记录
         (Id, 病历记录id, 书写人, 书写日期, 版本序号)
         Select Id, 病历记录id, 书写人, 书写日期, 版本序号 
		 From H病人病历修订记录 Where 病历记录id = n_Rec_Id;
   
      Insert Into 病人病历内容
         (Id, 病历示范id, 病历记录id, 排列序号, 元素类型, 元素编码, 填写时机, 标题文本, 文本转储, 标题显示, 标题字体, 标题颜色,
          标题位置, 内容字体, 内容颜色, 内容位置, 嵌入方式, 病历修订id)
         Select Id, 病历示范id, 病历记录id, 排列序号, 元素类型, 元素编码, 填写时机, 标题文本, 文本转储, 标题显示, 标题字体,
                标题颜色, 标题位置, 内容字体, 内容颜色, 内容位置, 嵌入方式, 病历修订id
         From H病人病历内容
         Where 病历记录id = n_Rec_Id;
   
      Insert Into 病人病历文本段
         (病历id, 行号, 内容)
         Select 病历id, 行号, 内容
         From H病人病历文本段
         Where 病历id In (Select Id From H病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into 病人病历所见单
         (病历id, 控件号, 控件类, 标题, 固定行, 固定列, 列, 行, 宽, 高, 对齐, 合并号, 不可写, 可屏蔽, 所见项id, 数值类型,
          所见内容, 计量单位)
         Select 病历id, 控件号, 控件类, 标题, 固定行, 固定列, 列, 行, 宽, 高, 对齐, 合并号, 不可写, 可屏蔽, 所见项id, 数值类型,
                所见内容, 计量单位
         From H病人病历所见单
         Where 病历id In (Select Id From H病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into 病人病历标记图
         (病历id, 类型, 内容, 字体, 点集, X1, Y1, X2, Y2, 填充色, 填充方式, 线条色, 线型, 线宽)
         Select 病历id, 类型, 内容, 字体, 点集, X1, Y1, X2, Y2, 填充色, 填充方式, 线条色, 线型, 线宽
         From H病人病历标记图
         Where 病历id In (Select Id From H病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into 病人病历外部图
         (病历id, 序号, 图象类型, 图象路径, 图象文件)
         Select 病历id, 序号, 图象类型, 图象路径, 图象文件
         From H病人病历外部图
         Where 病历id In (Select Id From H病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into 病人过敏记录
         (Id, 病人id, 主页id, 记录来源, 病历id, 药物id, 药物名, 记录时间, 记录人, 结果)
         Select Id, 病人id, 主页id, 记录来源, 病历id, 药物id, 药物名, 记录时间, 记录人, 结果
         From H病人过敏记录
         Where 病历id In (Select Id From H病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into 病人诊断记录
         (Id, 病人id, 主页id, 记录来源, 病历id, 诊断类型, 疾病id, 诊断id, 证候id, 诊断描述, 出院情况, 是否未治, 是否疑诊,
          记录日期, 记录人, 取消时间, 取消人, 医嘱id, 诊断次序, 编码序号)
         Select Id, 病人id, 主页id, 记录来源, 病历id, 诊断类型, 疾病id, 诊断id, 证候id, 诊断描述, 出院情况, 是否未治, 是否疑诊,
                记录日期, 记录人, 取消时间, 取消人, 医嘱id, 诊断次序, 编码序号
         From H病人诊断记录
         Where 病历id In (Select Id From H病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into 病人家系图
         (病历id, 序号, 父亲, 母亲, 姓名, 称谓, 性别, 状态, 说明, 怀孕, 养育关系, 婚姻关系)
         Select 病历id, 序号, 父亲, 母亲, 姓名, 称谓, 性别, 状态, 说明, 怀孕, 养育关系, 婚姻关系
         From H病人家系图
         Where 病历id In (Select Id From H病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into 病人手麻记录
         (Id, 病人id, 主页id, 记录来源, 病历id, 手术日期, 手术开始时间, 手术结束时间, 拟行手术, 手术操作id, 诊疗项目id, 已行手术,
          主刀医师, 第一助手, 第二助手, 手术护士, 麻醉开始时间, 麻醉结束时间, 麻醉方式, 麻醉类型, 麻醉质量, 输液总量, 麻醉医师,
          输氧开始时间, 输氧结束时间, 切口, 愈合, 记录日期, 记录人, 取消时间, 取消人)
         Select Id, 病人id, 主页id, 记录来源, 病历id, 手术日期, 手术开始时间, 手术结束时间, 拟行手术, 手术操作id, 诊疗项目id,
                已行手术, 主刀医师, 第一助手, 第二助手, 手术护士, 麻醉开始时间, 麻醉结束时间, 麻醉方式, 麻醉类型, 麻醉质量,
                输液总量, 麻醉医师, 输氧开始时间, 输氧结束时间, 切口, 愈合, 记录日期, 记录人, 取消时间, 取消人
         From H病人手麻记录
         Where 病历id In (Select Id From H病人病历内容 Where 病历记录id = n_Rec_Id);
   
      Insert Into 病人麻醉用药
         (记录id, 类型, 序号, 药物id, 药物名, 总量, 方式)
         Select 记录id, 类型, 序号, 药物id, 药物名, 总量, 方式
         From H病人麻醉用药
         Where 记录id In
               (Select Id From H病人手麻记录 Where 病历id In (Select Id From H病人病历内容 Where 病历记录id = n_Rec_Id));
   
      Delete H病人病历记录 Where Id = n_Rec_Id;
   End Zl_Retu_Cpr;

   --------------------------------------------
   --返回指定ID的病人医嘱记录子过程；
   --------------------------------------------
   Procedure Zl_Retu_Order(n_Rec_Id H病人医嘱记录.Id%Type) As
   Begin
      Insert Into 病人医嘱记录
         (Id, 相关id, 病人来源, 病人id, 主页id, 姓名, 性别, 年龄, 挂号单, 婴儿, 病人科室id, 序号, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id,
          标本部位, 收费细目id, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 执行科室id, 皮试结果, 执行频次, 频率次数, 频率间隔,
          间隔单位, 执行时间方案, 计价特性, 执行性质, 紧急标志, 开始执行时间, 执行终止时间, 上次执行时间, 开嘱科室id, 开嘱医生,
          开嘱时间, 校对护士, 校对时间, 停嘱医生, 停嘱时间, 确认停嘱时间, 申请id, 前提id, 是否上传, 天数,审查结果)
         Select Id, 相关id, 病人来源, 病人id, 主页id, 姓名, 性别, 年龄, 挂号单, 婴儿, 病人科室id, 序号, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id,
                标本部位, 收费细目id, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 执行科室id, 皮试结果, 执行频次, 频率次数, 频率间隔,
                间隔单位, 执行时间方案, 计价特性, 执行性质, 紧急标志, 开始执行时间, 执行终止时间, 上次执行时间, 开嘱科室id,
                开嘱医生, 开嘱时间, 校对护士, 校对时间, 停嘱医生, 停嘱时间, 确认停嘱时间, 申请id, 前提id, 是否上传, 天数,审查结果
         From H病人医嘱记录
         Where Id = n_Rec_Id;
   
      Insert Into 病人医嘱计价
         (医嘱id, 收费细目id, 数量, 单价, 从项)
         Select 医嘱id, 收费细目id, 数量, 单价, 从项 From H病人医嘱计价 Where 医嘱id = n_Rec_Id;
      Delete H病人医嘱计价 Where 医嘱id = n_Rec_Id;
   
      --返回医嘱签名数据
      Insert Into 医嘱签名记录(
		ID, 签名规则, 签名信息, 证书ID, 签名时间, 签名人)
      Select 
	    ID, 签名规则, 签名信息, 证书ID, 签名时间, 签名人 
	  From H医嘱签名记录 
	  Where ID IN (Select 签名ID From H病人医嘱状态 Where 医嘱id = n_Rec_Id);
	  IF SQL%RowCount<>0 Then--可能转其它医嘱时已经删了
          Delete H医嘱签名记录 Where ID IN (Select 签名ID From H病人医嘱状态 Where 医嘱id = n_Rec_Id);
      End IF;
	  
      Insert Into 病人医嘱状态
         (医嘱id, 操作类型, 操作人员, 操作时间, 签名ID)
         Select 医嘱id, 操作类型, 操作人员, 操作时间, 签名ID From H病人医嘱状态 Where 医嘱id = n_Rec_Id;
      Delete H病人医嘱状态 Where 医嘱id = n_Rec_Id;
   
      --医嘱发送部分
      Insert Into 病人医嘱发送
         (医嘱id, 发送号, 记录性质, No, 记录序号, 发送数次, 发送人, 发送时间, 首次时间, 末次时间, 执行状态, 执行部门id, 计费状态,
          报告id, 执行间, 执行过程, 采样人, 采样时间, 样本条码)
         Select 医嘱id, 发送号, 记录性质, No, 记录序号, 发送数次, 发送人, 发送时间, 首次时间, 末次时间, 执行状态, 执行部门id,
                计费状态, 报告id, 执行间, 执行过程, 采样人, 采样时间, 样本条码
         From H病人医嘱发送
         Where 医嘱id = n_Rec_Id;
   
      Insert Into 病人医嘱附费
         (医嘱id, 发送号, 记录性质, No)
         Select 医嘱id, 发送号, 记录性质, No From H病人医嘱附费 Where 医嘱id = n_Rec_Id;
      Delete H病人医嘱附费 Where 医嘱id = n_Rec_Id;
   
      Insert Into 病人医嘱执行
         (医嘱id, 发送号, 要求时间, 本次数次, 执行摘要, 执行人, 执行时间, 登记人, 登记时间)
         Select 医嘱id, 发送号, 要求时间, 本次数次, 执行摘要, 执行人, 执行时间, 登记人, 登记时间
         From H病人医嘱执行
         Where 医嘱id = n_Rec_Id;
      Delete H病人医嘱执行 Where 医嘱id = n_Rec_Id;
   
      Insert Into 影像检查记录
         (医嘱id, 发送号, 影像类别, 检查号, 姓名, 英文名, 性别, 年龄, 出生日期, 身高, 体重, 病理检查, 发放胶片, 检查uid, 位置一,
          位置二, 位置三, 检查设备, 报告图象, 接收日期)
         Select 医嘱id, 发送号, 影像类别, 检查号, 姓名, 英文名, 性别, 年龄, 出生日期, 身高, 体重, 病理检查, 发放胶片, 检查uid,
                位置一, 位置二, 位置三, 检查设备, 报告图象, 接收日期
         From H影像检查记录
         Where 医嘱id = n_Rec_Id;
      For r_Ris In (Select 检查uid
                    From H影像检查记录 r, H病人医嘱发送 s
                    Where r.医嘱id = s.医嘱id And r.发送号 = s.发送号 And s.医嘱id = n_Rec_Id) Loop
         Insert Into 影像检查序列
            (序列uid, 检查uid, 序列号, 序列描述, 采集时间)
            Select 序列uid, 检查uid, 序列号, 序列描述, 采集时间 From H影像检查序列 Where 检查uid = r_Ris.检查uid;
         For r_Seq In (Select 序列uid From H影像检查序列 Where 检查uid = r_Ris.检查uid) Loop
            Insert Into 影像检查图象
               (图像uid, 序列uid, 图像号, 图像描述)
               Select 图像uid, 序列uid, 图像号, 图像描述 From H影像检查图象 Where 序列uid = r_Seq.序列uid;
            Delete H影像检查图象 Where 序列uid = r_Seq.序列uid;
         End Loop;
         Delete H影像检查序列 Where 检查uid = r_Ris.检查uid;
      End Loop;
      Delete H影像检查记录 Where 医嘱id = n_Rec_Id;
   
      Delete H病人医嘱发送 Where 医嘱id = n_Rec_Id;
   
      --手术记录部分
      Insert Into 病人手术记录
         (Id, 医嘱id, 病历id, 病人id, 主页id, 记录来源, 手术日期, 手术开始时间, 手术结束时间, 手术规模, 麻醉开始时间,
          麻醉结束时间, 麻醉方式, 麻醉类型, 麻醉质量, 输液总量, 输氧开始时间, 输氧结束时间, 切口, 愈合, 无菌手术, 手术间, 手术间id,
          手术室id, 记录日期, 记录人, 取消时间, 取消人)
         Select Id, 医嘱id, 病历id, 病人id, 主页id, 记录来源, 手术日期, 手术开始时间, 手术结束时间, 手术规模, 麻醉开始时间,
                麻醉结束时间, 麻醉方式, 麻醉类型, 麻醉质量, 输液总量, 输氧开始时间, 输氧结束时间, 切口, 愈合, 无菌手术, 手术间,
                手术间id, 手术室id, 记录日期, 记录人, 取消时间, 取消人
         From H病人手术记录
         Where 医嘱id = n_Rec_Id;
      For r_Ops In (Select Id From H病人手术记录 Where 医嘱id = n_Rec_Id) Loop
      
         Insert Into 病人手术情况
            (记录id, 性质, 缺省, 手术名称, 手术操作id, 诊疗项目id)
            Select 记录id, 性质, 缺省, 手术名称, 手术操作id, 诊疗项目id From H病人手术情况 Where 记录id = r_Ops.Id;
         Delete H病人手术情况 Where 记录id = r_Ops.Id;
      
         Insert Into 病人手术用药
            (记录id, 类型, 序号, 药品id, 药名id, 药品名称, 准备总量, 使用总量, 给药方式, 执行科室id)
            Select 记录id, 类型, 序号, 药品id, 药名id, 药品名称, 准备总量, 使用总量, 给药方式, 执行科室id
            From H病人手术用药
            Where 记录id = r_Ops.Id;
         Delete H病人手术用药 Where 记录id = r_Ops.Id;
      
         Insert Into 病人手术人员
            (记录id, 科室id, 岗位, 人员id, 编码, 姓名)
            Select 记录id, 科室id, 岗位, 人员id, 编码, 姓名 From H病人手术人员 Where 记录id = r_Ops.Id;
         Delete H病人手术人员 Where 记录id = r_Ops.Id;
      
         Insert Into 病人手术材料
            (记录id, 序号, 材料id, 准备数量, 实用数量, 清点结果, 执行科室id, 附加说明)
            Select 记录id, 序号, 材料id, 准备数量, 实用数量, 清点结果, 执行科室id, 附加说明
            From H病人手术材料
            Where 记录id = r_Ops.Id;
         Delete H病人手术材料 Where 记录id = r_Ops.Id;
      
         Insert Into 病人手术计价
            (记录id, 序号, 数量, 单价, 收费细目id, 执行科室id)
            Select 记录id, 序号, 数量, 单价, 收费细目id, 执行科室id From H病人手术计价 Where 记录id = r_Ops.Id;
         Delete H病人手术计价 Where 记录id = r_Ops.Id;
      End Loop;
      Delete H病人手术记录 Where 医嘱id = n_Rec_Id;
   
      Insert Into 病人手术状态
         (医嘱id, 序号, 上次序号, 处理性质, 记录状态, 处理人, 处理时间, 附加说明, 当前状态, 单据号)
         Select 医嘱id, 序号, 上次序号, 处理性质, 记录状态, 处理人, 处理时间, 附加说明, 当前状态, 单据号
         From H病人手术状态
         Where 医嘱id = n_Rec_Id;
      Delete H病人手术状态 Where 医嘱id = n_Rec_Id;
   
      --检验部分
      Insert Into 检验标本记录
         (Id, 医嘱id, 标本序号, 采样时间, 采样人, 标本类型, 核收人, 核收时间, 样本状态, 检验人, 检验时间, 审核人, 审核时间,
          合并报告号, 打印次数, 申请类型, 仪器id, 样本条码, 报告结果, 备注, 未通过审核原因, 申请时间, 标本形态, 是否质控品,
          执行科室id)
         Select Id, 医嘱id, 标本序号, 采样时间, 采样人, 标本类型, 核收人, 核收时间, 样本状态, 检验人, 检验时间, 审核人, 审核时间,
                合并报告号, 打印次数, 申请类型, 仪器id, 样本条码, 报告结果, 备注, 未通过审核原因, 申请时间, 标本形态, 是否质控品,
                执行科室id
         From H检验标本记录
         Where 医嘱id = n_Rec_Id;
      For r_Retu In (Select Id From H检验普通结果 Where 检验标本id In (Select Id From H检验标本记录 Where 医嘱id = n_Rec_Id)) Loop
         Insert Into 检验普通结果
            (Id, 检验标本id, 检验项目id, 检验结果, 结果标志, 结果参考, 修改者, 修改时间, 记录类型, 原始结果, 原始记录时间, 记录者,
             是否检验, 修改原因, 细菌id, 仪器id, 培养描述)
            Select Id, 检验标本id, 检验项目id, 检验结果, 结果标志, 结果参考, 修改者, 修改时间, 记录类型, 原始结果, 原始记录时间,
                   记录者, 是否检验, 修改原因, 细菌id, 仪器id, 培养描述
            From H检验普通结果
            Where Id = r_Retu.Id;
         Insert Into 检验药敏结果
            (细菌结果id, 抗生素id, 修改者, 修改时间, 结果, 结果类型, 记录类型, 仪器id)
            Select 细菌结果id, 抗生素id, 修改者, 修改时间, 结果, 结果类型, 记录类型, 仪器id
            From H检验药敏结果
            Where 细菌结果id = r_Retu.Id;
         Delete H检验药敏结果 Where 细菌结果id = r_Retu.Id;
         Delete H检验普通结果 Where Id = r_Retu.Id;
      End Loop;
      Delete H检验标本记录 Where 医嘱id = n_Rec_Id;
   
      Delete H病人医嘱记录 Where Id = n_Rec_Id;
   End Zl_Retu_Order;

   --------------------------------------------
   --以下为主程序体
   --------------------------------------------
Begin
   --防止数据不一致，先禁用一些约束
   Begin
	  Execute Immediate 'Alter Table H病人医嘱状态 Modify Constraint H病人医嘱状态_FK_签名ID Disable';
   Exception
      When Others Then Null;
   End;

   If n_Flag = 0 Then
      Insert Into 病人挂号记录
         (Id, No, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间,
          操作员编号, 操作员姓名, 摘要)
         Select Id, No, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间,
                登记时间, 操作员编号, 操作员姓名, 摘要
         From H病人挂号记录
         Where No = v_Times;
      For r_Cpr In (Select Id From H病人病历记录 Where 病人id = n_Patiid And 挂号单 = v_Times) Loop
         Zl_Retu_Cpr(r_Cpr.Id);
      End Loop;
      For r_Order In (Select Id From H病人医嘱记录 Where 病人id = n_Patiid And 挂号单 = v_Times) Loop
         Zl_Retu_Order(r_Order.Id);
      End Loop;
      Delete H病人挂号记录 Where No = v_Times;
   Else
      For r_Cpr In (Select Id From H病人病历记录 Where 病人id = n_Patiid And 主页id = To_Number(v_Times)) Loop
         Zl_Retu_Cpr(r_Cpr.Id);
      End Loop;
      For r_Order In (Select Id From H病人医嘱记录 Where 病人id = n_Patiid And 主页id = To_Number(v_Times)) Loop
         Zl_Retu_Order(r_Order.Id);
      End Loop;
      Update 病案主页 Set 数据转出 = 0 Where 病人id = n_Patiid And 主页id = To_Number(v_Times);
   End If;

   --启用约束
   Begin
	  Execute Immediate 'Alter Table H病人医嘱状态 Modify Constraint H病人医嘱状态_FK_签名ID Enable';
   Exception
      When Others Then Null;
   End;

   Commit;
End Zl_Retu_Clinic;
/

CREATE OR REPLACE Procedure ZL_病人医嘱记录_作废(
--功能：作废指定的医嘱(未发送的长嘱或临嘱)
--说明：一并给药的只能调用一次(界面显示有多行)
--参数：ID_IN=相关ID为NULL的医嘱的ID(给药途径,中药用法,检查项目,主要手术,及独立医嘱)
    ID_IN    病人医嘱记录.ID%TYPE
) IS
    --包含医嘱相关信息
    Cursor c_Advice is
        Select A.病人ID,A.挂号单,A.主页ID,A.婴儿,A.医嘱状态,
			A.上次执行时间,A.医嘱内容,A.诊疗类别,B.操作类型,A.执行科室ID
        From 病人医嘱记录 A,诊疗项目目录 B
        Where A.诊疗项目ID=B.ID And A.ID=ID_IN;
    r_Advice c_Advice%RowType;
	
	--门诊医嘱作废：
    --根据医嘱及发送NO求出本次回退要销帐或删除(门诊划价单)的费用记录
    --一组医嘱并不是都填写了发送记录,且可能NO不同(药品有,用法煎法不一定有)
    --不管发送记录的计费状态(可能无需计费),有费用记录自然关联出来
    --费用只求价格父号为空的,以便取序号销帐(门诊记帐单)
    --只管记录状态为1的记录,如果已经销帐或部份销帐的记录,不再处理
    --ZYL:如果是药品，即使已经收费和发药,仍然允许作废
    Cursor c_RollMoney(v_发送号 病人医嘱发送.发送号%Type) is
        Select A.记录性质,A.记录状态,A.NO,A.序号,
			A.执行状态 as 费用执行,C.执行状态 as 医嘱执行
        From 病人费用记录 A,病人医嘱记录 B,病人医嘱发送 C,诊疗项目目录 I
        Where C.医嘱ID=B.ID And C.发送号=v_发送号
            And (B.ID=ID_IN Or B.相关ID=ID_IN)
            And A.医嘱序号=B.ID And A.记录状态 IN(0,1)
            And A.NO=C.NO And A.记录性质=C.记录性质
            And B.诊疗项目ID=I.ID And A.价格父号 IS NULL
            And (
				A.收费类别 Not In ('5','6','7','E')
                Or A.收费类别='E' And I.操作类型 Not In ('2','3','4')
                Or A.收费类别 In ('5','6','7') And Nvl(A.执行状态,0)=0 And Not(A.记录性质=1 And A.记录状态<>0)
				Or Exists(Select 参数值 From 系统参数表 Where 参数号=68 And Nvl(参数值,0)=0)
				)
        Order BY A.收费细目ID;

	--用于删除报告记录
	Cursor c_Case is
		Select 报告ID From 病人医嘱发送 
		Where 报告ID IS Not NULL And 医嘱ID IN(
			Select ID From 病人医嘱记录 Where ID=ID_IN OR 相关ID=ID_IN);
	r_Case c_Case%RowType;

    v_发送号        病人医嘱发送.发送号%Type;
    v_费用NO        病人费用记录.NO%Type;
    v_记录性质      病人费用记录.记录性质%Type;
    v_费用序号      Varchar2(255);

    v_Date          Date;
	v_Count			Number;
    v_Temp          Varchar2(255);
    v_人员编号      病人费用记录.操作员编号%Type;
    v_人员姓名      病人费用记录.操作员姓名%Type;

    v_Error			Varchar2(255);
    Err_Custom      Exception;
Begin
    --检查医嘱状态是否正确:并发操作
    Open c_Advice;
    Fetch c_Advice Into r_Advice;

    If r_Advice.挂号单 IS NULL Then
        IF r_Advice.医嘱状态 IN(4,8,9) Then
            v_Error:='医嘱"'||r_Advice.医嘱内容||'"已经被作废或停止，不能再作废。';
            Raise Err_Custom;
        Elsif r_Advice.上次执行时间 IS Not NULL Then
            v_Error:='医嘱"'||r_Advice.医嘱内容||'"已经发送，不能被作废。';
            Raise Err_Custom;
        End IF;
    Else
        IF r_Advice.医嘱状态<>8 Then
            v_Error:='医嘱"'||r_Advice.医嘱内容||'"尚未发送或已经作废。';
            Raise Err_Custom;
        End IF;

		--门诊医嘱只可能发送一次
		Select Count(*) Into v_Count 
		From 病人医嘱发送 
		Where 执行状态 IN(1,3) And 医嘱ID IN(
			Select ID From 病人医嘱记录 Where ID=ID_IN Or 相关ID=ID_IN);
		If v_Count>0 Then
			v_Error:='该医嘱已经执行或正在执行，不能作废。';
			Raise Err_Custom;
		End IF;
    End IF;

    --当前操作人员
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);
    Select Sysdate Into v_Date From Dual;

    Update 病人医嘱记录
        Set 医嘱状态=4
    Where ID=ID_IN Or 相关ID=ID_IN;

    Insert Into 病人医嘱状态(
        医嘱ID,操作类型,操作人员,操作时间)
    Select
        ID,4,v_人员姓名,v_Date
    From 病人医嘱记录
    Where ID=ID_IN Or 相关ID=ID_IN;

    --其它处理
    ---------------------------------------------------------------------------------------
    --门诊/住院医嘱作废时把对应的申请单作废
	Update 病人病历记录
		Set 作废人=v_人员姓名,作废日期=v_Date
	Where 作废人 IS NULL And ID IN(
		Select 申请ID From 病人医嘱记录 Where ID=ID_IN Or 相关ID=ID_IN);

    If r_Advice.挂号单 IS Not NULL Then
        --门诊医嘱(临嘱)作废时还需要回退相关内容:只有一次发送
        --回退划价或记帐费用
        Begin
            --医嘱ID为传入值的这条医嘱不一定发送了的,甚至无发送。
            Select Distinct 发送号 Into v_发送号 From 病人医嘱发送
            Where 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=ID_IN Or 相关ID=ID_IN);
        Exception
            When Others Then v_发送号:=NULL;
        End;
        If v_发送号 IS Not NULL Then
            --将该组医嘱的费用删除或销帐(按一组医嘱可能有不同NO处理)
            --门诊记帐：如果原始费用已被销帐(或部分销帐),调用过程中有判断
            --门诊划价：如果已收费，则不允许删除
            v_费用NO:=NULL;v_费用序号:=NULL;
            For r_RollMoney In c_RollMoney(v_发送号) Loop
                If Nvl(v_费用NO,'空')<>r_RollMoney.NO Then
                    If v_费用序号 IS Not NULL And v_费用NO IS Not NULL Then
                        v_费用序号:=Substr(v_费用序号,2);
                        IF v_记录性质=1 Then
                            zl_门诊划价记录_Delete(v_费用NO,v_费用序号);
                        Elsif v_记录性质=2 Then
                            zl_门诊记帐记录_Delete(v_费用NO,v_费用序号,v_人员编号,v_人员姓名);
                        End If;
                    End IF;
                    v_费用序号:=NULL;
                End IF;
                v_记录性质:=r_RollMoney.记录性质;
                v_费用NO:=r_RollMoney.NO;
                v_费用序号:=v_费用序号||','||r_RollMoney.序号;

                If Nvl(r_RollMoney.医嘱执行,0) IN(1,3) Then --1-完全执行;3-正在执行
                    v_Error:='医嘱"'||r_Advice.医嘱内容||'"已经执行或正在执行，不能作废。';
                    Raise Err_Custom;
                End IF;
                If Nvl(r_RollMoney.费用执行,0) IN(1,2) Then --1-完全执行;2-部份执行
                    v_Error:='医嘱费用单据"'||r_RollMoney.NO||'"中的内容已经全部或部分执行，不能作废。';
                    Raise Err_Custom;
                End IF;
                If r_RollMoney.记录性质=1 And r_RollMoney.记录状态<>0 Then
                    v_Error:='医嘱费用单据"'||r_RollMoney.NO||'"已经收费，不能作废。';
                    Raise Err_Custom;
                End IF;
            End Loop;
            If v_费用序号 IS Not NULL And v_费用NO IS Not NULL Then
                v_费用序号:=Substr(v_费用序号,2);
                IF v_记录性质=1 Then
                    zl_门诊划价记录_Delete(v_费用NO,v_费用序号);
                Elsif v_记录性质=2 Then
                    zl_门诊记帐记录_Delete(v_费用NO,v_费用序号,v_人员编号,v_人员姓名);
                End If;
            End IF;
			
			Open c_Case;--必须先打开

            --回退医嘱发送记录
            Delete From 病人医嘱发送 Where 医嘱ID IN(
				Select ID From 病人医嘱记录 Where ID=ID_IN Or 相关ID=ID_IN);

            --删除对应的报告单
			Fetch c_Case Into r_Case;
			While c_Case%Found Loop
	            Delete From 病人病历记录 Where ID=r_Case.报告ID;
				Fetch c_Case Into r_Case;
			End Loop;
			Close c_Case;

            --回退特殊医嘱的处理
            If r_Advice.诊疗类别='Z' And Nvl(r_Advice.操作类型,'0')<>'0' And Nvl(r_Advice.婴儿,0)=0 Then
                If r_Advice.操作类型='1' And r_Advice.执行科室ID IS Not NULL Then
                    --留观医嘱
					Select Count(*) Into v_Count From 病案主页 Where 病人ID=r_Advice.病人ID And Nvl(主页ID,0)=0 And 入院科室ID=r_Advice.执行科室ID And 病人性质 IN(1,2);
					If v_Count=1 Then
						zl_入院病案主页_Delete(r_Advice.病人ID,0);
					End IF;
                ElsIf r_Advice.操作类型='2' And r_Advice.执行科室ID IS Not NULL Then
                    --住院医嘱
					Select Count(*) Into v_Count From 病案主页 Where 病人ID=r_Advice.病人ID And Nvl(主页ID,0)=0 And 入院科室ID=r_Advice.执行科室ID And Nvl(病人性质,0)=0;
					If v_Count=1 Then
						zl_入院病案主页_Delete(r_Advice.病人ID,0);
					End IF;
                End IF;
            End IF;
        End IF;
    End IF;
	
	--删除过敏登记记录
	If r_Advice.诊疗类别='E' And r_Advice.操作类型='1'  Then

		--Update 病人医嘱记录 Set 皮试结果=Null Where ID=ID_IN; --保留最后的皮试结果

		For r_Test IN(Select 操作时间 From 病人医嘱状态 Where 医嘱ID=ID_IN And 操作类型=10) Loop
			Delete From 病人过敏记录 
			Where 病人ID=r_Advice.病人ID And 记录来源=2
				And Nvl(主页ID,0)=Nvl(r_Advice.主页ID,0)
				And 记录时间=r_Test.操作时间;
		End Loop;
	End IF;

    Close c_Advice;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱记录_作废;
/

-------------------------------------------------------------------------
--刘兴宏:对冲销记录进行发票信息修改，但不修改发票金额.

-- 材料外购入库发票信息的修改
CREATE OR REPLACE PROCEDURE zl_材料外购发票信息_UPDATE (
    NO_IN		IN 药品收发记录.NO%TYPE := NULL,
    记录状态_IN		IN 药品收发记录.记录状态%type:=NULL,
    序号_IN		IN 药品收发记录.序号%TYPE:=NULL,
    发票号_IN		IN 应付记录.发票号%TYPE := NULL,
    发票日期_IN		IN 应付记录.发票日期%TYPE := NULL,
    发票金额_IN		IN 应付记录.发票金额%TYPE := NULL,
    供药单位_IN		in 应付记录.单位ID%TYPE:=0
)
IS
    mErrMsg		varchar2(255);
    mErrItem		exception;

    V_NO		应付记录.NO%TYPE;
    V_应付ID		应付记录.ID%TYPE;
    V_收发ID		应付记录.收发ID%TYPE;
    V_付款序号		应付记录.付款序号%TYPE;
    V_发票金额		应付记录.发票金额%TYPE;--旧发票金额
    V_供药单位ID	应付记录.单位ID%TYPE;
BEGIN
    --取是否付款及总额
    BEGIN 
        Select max(付款序号),sum(nvl(发票金额,0)) INTO v_付款序号,v_发票金额 
        FROM 应付记录 
        WHERE 收发id=(Select ID From 药品收发记录 Where NO=NO_IN And 序号=序号_IN And 单据=15) 
		AND 系统标识=5 And 记录性质=-1;
    EXCEPTION 
        WHEN OTHERS THEN 
        v_发票金额:=0;
    END ;


    v_付款序号:=nvl(v_付款序号,0);
    
    IF v_付款序号<>0 then
       mErrMsg:='[ZLSOFT]该单据已经被付了款，不能再修改发票信息[ZLSOFT]';
       RAISE mErrItem;
    END IF ;

    if 发票金额_IN>v_发票金额 And v_发票金额<>0 then 
        mErrMsg:='[ZLSOFT]发票金额不能小于计划付款金额[ZLSOFT]';
        raise mErrItem;
    end if ;

    
    --判断是否冲销后的记录
    IF 记录状态_IN<>1 THEN 
	IF nvl(发票号_IN,' ') <>' ' AND 发票金额_IN=0  THEN 
	       mErrMsg:='[ZLSOFT]不能对发票金额为零的发票信息进行修改。[ZLSOFT]';
	       RAISE mErrItem;
	END IF ;

	IF nvl(发票号_IN,' ') =' ' AND 发票金额_IN<>0  THEN 
	       mErrMsg:='[ZLSOFT]不能对冲销或被冲销记录的发票号改为空,不能保存！[ZLSOFT]';
	       RAISE mErrItem;
	END IF ;

	--更新相关的发票信息,只更改发票号，发票日期
	FOR V_收发 IN (Select ID From 药品收发记录 WHERE 单据=15 AND NO=NO_IN AND 序号=序号_IN )
	LOOP 
	    UPDATE 应付记录
	    SET 发票号 = 发票号_IN,
		发票日期 = 发票日期_IN
	    WHERE 收发ID = V_收发.ID And 系统标识=5 AND 记录性质=0;
	END LOOP ;
	RETURN ;
	
    END IF ;

    SELECT A.ID,nvl(B.发票金额,0),A.供药单位ID    INTO V_收发ID,V_发票金额,V_供药单位ID
    FROM 药品收发记录 A,(Select * From 应付记录 Where 系统标识=5  AND 记录性质<>-1 And 付款序号 Is NULL) B
    WHERE A.ID = B.收发ID(+)
        AND A.NO = NO_IN
        AND A.单据 = 15
        AND A.记录状态 = 1
        AND A.序号 = 序号_IN; 
    
    UPDATE 应付记录
    SET 发票号 = 发票号_IN,
        发票日期 = 发票日期_IN,
        发票金额 = 发票金额_IN,
        单位ID=供药单位_IN
    WHERE 收发ID = V_收发ID And 系统标识=5 And 记录状态=1 And 记录性质=0;


    if sql%rowcount=0 then 
        IF 发票号_IN IS NOT NULL THEN
            --如果是第一笔明细,则产生应付记录的NO
            BEGIN 
                SELECT NO INTO V_NO FROM 应付记录 
                WHERE 系统标识=5 AND 记录性质=0 AND 记录状态=1 
                    AND 入库单据号=NO_IN AND ROWNUM<2;
            EXCEPTION
                WHEN OTHERS THEN V_NO:=NEXTNO(69);
            END ;

            SELECT 应付记录_ID.NEXTVAL INTO V_应付ID FROM DUAL;
            
            INSERT INTO 应付记录
            (ID,记录性质,记录状态,项目id,序号,单位ID,NO,系统标识,收发ID,入库单据号,单据金额,发票号,发票日期,发票金额,品名,
            规格,产地,批号,计量单位,数量,采购价,采购金额,填制人,填制日期,审核人,审核日期,摘要)
            select V_应付ID,0,1,A.药品id,A.序号,供药单位_IN,V_NO,5,V_收发ID,A.NO,A.零售金额,发票号_IN,发票日期_IN,发票金额_IN,B.名称,
            B.规格,B.产地,A.批号,B.计算单位,A.实际数量,A.成本价,A.成本金额,A.填制人,A.填制日期,A.审核人,A.审核日期,A.摘要
            from 药品收发记录 A,收费项目目录 B
            Where A.单据=15 And A.NO=NO_in And A.序号=序号_IN And A.药品ID=B.ID;
        END IF;
    END IF;

    UPDATE 应付余额 SET 金额 = NVL (金额,0) - V_发票金额
    WHERE 单位ID = V_供药单位ID AND 性质 = 1;
    IF SQL%NOTFOUND THEN
        INSERT INTO 应付余额(单位ID,性质,金额) VALUES (V_供药单位ID,1,-V_发票金额);
    END IF; 
    UPDATE 应付余额 SET 金额=NVL(金额,0)+发票金额_IN
    WHERE 单位ID=供药单位_IN AND 性质=1;

    IF SQL%NOTFOUND THEN
        INSERT INTO 应付余额(单位ID,性质,金额) VALUES (供药单位_IN,1,发票金额_IN);
    END IF; 

    --更新药品收发记录中的供药单位
    UPDATE 药品收发记录 SET 供药单位ID=供药单位_IN WHERE NO=NO_IN AND 单据=15 And 序号=序号_IN;

    --更新药品库存里的上次供应商
    UPDATE 药品库存 SET 上次供应商ID=供药单位_IN WHERE 性质=1 and  (库房ID,药品ID) IN (SELECT 库房ID,药品ID FROM 药品收发记录 WHERE NO=NO_IN AND 单据=15);
EXCEPTION
    when mErrItem then Raise_application_error (-20101,mErrMsg); 
    WHEN NO_data_found THEN
        Raise_application_error (-20101,'[ZLSOFT]该单据已经被他人冲销或已经付过款！[ZLSOFT]'); 
    WHEN OTHERS THEN
    zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料外购发票信息_UPDATE;
/



CREATE OR REPLACE PROCEDURE zl_材料自制原料出库_Insert (
		NO_IN		IN 药品收发记录.NO%TYPE,
		对方部门ID_IN	IN 药品收发记录.对方部门ID%TYPE)
AS
		mErrMsg         varchar2(100);
		mErrItem        EXCEPTION ;

		v_数量		药品收发记录.实际数量%type;
		v_成本价	药品收发记录.成本价%type;
		v_成本金额	药品收发记录.成本金额%type;
		v_差价		药品收发记录.差价%type;

		v_售价		药品收发记录.零售价%type;
		v_零售金额	药品收发记录.零售金额%type;

		v_库存金额	药品库存.实际金额%type;
		v_库存差价	药品库存.实际差价%type;
		v_可用数量	药品库存.可用数量%type;
		v_实际数量	药品库存.实际数量%type;
		v_上次产地	药品库存.上次产地%type;
		v_负成本计算	系统参数表.参数值%type;
		V_出的类别ID	药品收发记录.入出类别ID%TYPE;--入出类别ID
		V_maxserial	药品收发记录.序号%TYPE;
BEGIN
	SELECT B.ID INTO V_出的类别ID    
	FROM 药品单据性质 A,药品入出类别 B
	WHERE A.类别ID = B.ID AND A.单据 = 31 AND B.系数 = -1
		AND ROWNUM < 2;
	
	SELECT 参数值 INTO v_负成本计算 FROM 系统参数表 WHERE 参数号=120;

	SELECT MAX (序号) INTO V_maxserial
	FROM 药品收发记录
	WHERE NO = NO_IN AND 单据 = 16 AND 入出系数 = 1;


	FOR v_自制 IN (SELECT * FROM 药品收发记录 WHERE NO = NO_IN AND 单据 = 16 AND 入出系数 = 1)
	LOOP 
		FOR v_组成 IN (	Select a.*,b.是否变价,c.指导差价率,c.成本价
				From 自制材料构成 a,收费项目目录 b,材料特性 c
				WHERE  a.原料材料id=b.id  AND  a.自制材料ID=v_自制.药品id AND a.原料材料id=c.材料id
			)
		LOOP 
			BEGIN 
				SELECT 可用数量,实际数量,实际差价,实际金额,上次产地 
					INTO v_可用数量,v_实际数量,v_库存差价,v_库存金额,v_上次产地
				FROM 药品库存
				WHERE 药品id=v_组成.原料材料id AND 性质 = 1 AND 库房ID=对方部门ID_IN;
			EXCEPTION 
				WHEN OTHERS THEN 
				     v_可用数量:=0;
				     v_实际数量:=0;
				     v_库存差价:=0;
				     v_库存金额:=0;
			END ;
			IF nvl(v_组成.是否变价,0)=1 THEN 
				--实价
				IF nvl(v_实际数量,0)> 0 THEN 
					v_售价:=nvl(v_库存金额,0)/v_实际数量;
				ELSE	
					--无库数:需提示
					mErrMsg:='[ZLSOFT]该单据中存在一笔以上原料的实际数量不足[ZLSOFT]';
					RAISE mErrItem;
				END IF ;
			ELSE 
				--定价,以现价为准
				BEGIN 
					SELECT nvl(现价,0) INTO v_售价  
					FROM 收费价目 
					WHERE  收费细目ID=v_组成.原料材料id AND ((SYSDATE BETWEEN 执行日期 AND 终止日期) OR (SYSDATE >= 执行日期 AND 终止日期 IS NULL));

				EXCEPTION 
					WHEN OTHERS THEN mErrMsg:='Err';
				END ;
				IF mErrMsg='Err' THEN 
					mErrMsg:='[ZLSOFT]该单据中存在一笔以上原料还未进行定价！[ZLSOFT]';
					RAISE mErrItem;
				END IF ;
			END IF ;
			v_数量:=nvl(v_自制.实际数量,0)*v_组成.分子/v_组成.分母;
			
			If v_数量=0 Then 
				mErrMsg:='[ZLSOFT]该单据中存在一笔以上原料的数量为零了！[ZLSOFT]';
				RAISE mErrItem;
			End If ;
			v_零售金额:=v_数量*v_售价;

			--算成本价
			IF nvl(v_库存金额,0)<=0 THEN 
				IF v_负成本计算='1' AND nvl(v_组成.成本价,0)>0 THEN 
					v_成本价:=v_组成.成本价;
					v_差价:=v_零售金额-v_数量*v_成本价;
				ELSE 
					v_差价:=v_零售金额*v_组成.指导差价率/100;
					v_成本价:=(v_零售金额-v_差价)/v_数量 ;
				END if;
			ELSE 
				v_差价 := v_零售金额 * (v_库存差价 / v_库存金额);
				v_成本价 := (v_零售金额 - v_差价) / v_数量;
			END IF ;
			v_成本价:=nvl(v_成本价,0);
			v_成本金额:=v_成本价*v_数量;
			V_maxserial:=V_maxserial+1;

			Insert INTO 药品收发记录
			    (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,产地,填写数量,实际数量,成本价,成本金额,
			    零售价,零售金额,差价,摘要,填制人,填制日期,费用ID,扣率)
			VALUES (
			    药品收发记录_ID.Nextval,1,16,NO_IN,V_maxserial,v_自制.对方部门ID,v_自制.库房ID,
			    V_出的类别ID,-1,v_组成.原料材料id,v_上次产地,v_数量,v_数量,
			    v_成本价,v_成本金额,v_售价,v_零售金额,v_差价,v_自制.摘要,
			    v_自制.填制人,v_自制.填制日期,v_自制.药品ID,v_自制.序号);

			--IF v_可用数量<0 then
			--    mErrMsg:='[ZLSOFT]该单据中存在一笔以上原料的可用数量不足[ZLSOFT]';
			--    RAISE mErrItem;
			--END IF ;

			UPDATE 药品库存
			SET 可用数量 = NVL (可用数量,0) - v_数量
			WHERE 库房ID = v_自制.对方部门ID AND 药品ID = v_组成.原料材料id AND 性质 = 1;

			IF SQL%NOTFOUND THEN
			    Insert INTO 药品库存 (库房ID,药品ID,性质,可用数量)
			    VALUES (v_自制.对方部门ID,v_组成.原料材料id,1,-v_数量);
			END IF;

			DELETE
			FROM 药品库存
			WHERE 库房ID=v_自制.对方部门ID And 药品ID=v_组成.原料材料id
				And nvl(可用数量,0)=0 And nvl(实际数量,0)=0
				And nvl(实际金额,0)=0 And nvl(实际差价,0)=0;

		END LOOP ;
	END LOOP ;
EXCEPTION
    WHEN mErrItem  THEN
        Raise_application_error (-20101,mErrMsg    );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料自制原料出库_Insert;
/



-----------------------------------------------------------
-- 材料自制入库的审核处理
--说明：首先对药品收发记录表中的审核人和审核时间进行处理，
--接着对药品库存和药品收发汇总表中的相应数量和金额进行处理
--特别说明：对药品库存的处理是分开的，对自制药品的处理与其他入库一样，
--对原料的处理，只对实际数量，实际金额，实际差价进行减少的处理，不对可用数量进行处理，因为以前保存时已处理
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_自制材料入库_verify (
    NO_IN        IN 药品收发记录.NO%TYPE := NULL,
    审核人_IN    IN 药品收发记录.审核人%TYPE := NULL
)
IS
	mErrItem		EXCEPTION;
	mErrMsg			varchar2(100);

	V_实际库存金额		药品库存.实际金额%TYPE;
	V_实际库存差价		药品库存.实际差价%TYPE;
	V_出库差价		药品库存.实际差价%TYPE;
	V_差价率		number(18,8);
	V_成本价		药品收发记录.成本价%TYPE;
	V_成本金额		药品收发记录.成本金额%TYPE;
	v_负成本计算		系统参数表.参数值%type;
	V_小数			number(2);

BEGIN

    	SELECT 参数值 INTO v_负成本计算 FROM 系统参数表 WHERE 参数号=120;


	UPDATE 药品收发记录
	SET 审核人 = NVL (审核人_IN,审核人),审核日期 = SYSDATE
	WHERE NO = NO_IN AND 单据 = 16    AND 记录状态 = 1     AND 审核人 IS NULL;

	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';    
		RAISE mErrItem;
	END IF;

	BEGIN 
	    SELECT  to_number(Nvl(参数值,缺省值),'99999')  INTO v_小数 
	    From 系统参数表 where 参数名='费用金额保留位数';
	EXCEPTION 
	WHEN OTHERS THEN 
		v_小数:=2;		
	END;


	FOR V_药品收发记录 IN (
			SELECT ID,实际数量,零售金额,差价,库房ID,药品ID,批次,成本价,批号,效期,产地,入出类别ID,入出系数,对方部门ID
			FROM 药品收发记录
			WHERE NO = NO_IN AND 单据 = 16 AND 记录状态 = 1    
			ORDER BY 药品ID    ) 
	LOOP
		--更改药品库存表的相应数据
		IF V_药品收发记录.入出系数 = -1 THEN
		    BEGIN
			SELECT nvl(实际金额,0),nvl(实际差价,0) INTO V_实际库存金额,V_实际库存差价
			FROM 药品库存
			WHERE 药品ID = V_药品收发记录.药品ID
			    AND NVL (批次,0) = NVL (V_药品收发记录.批次,0)
			    AND 库房ID = V_药品收发记录.库房ID
			    AND 性质 = 1
			    AND ROWNUM = 1;
		    EXCEPTION
			WHEN OTHERS THEN
			    V_实际库存金额 := 0;
		    END;

		    IF V_实际库存金额 <= 0 THEN
			BEGIN
			    SELECT 指导差价率 / 100 INTO V_差价率
			    FROM 材料特性
			    WHERE 材料ID = V_药品收发记录.药品ID;
			EXCEPTION
			    WHEN OTHERS THEN
			    V_差价率 := 0;
			END;
			IF v_负成本计算 ='1' THEN 
				BEGIN 
					SELECT nvl(成本价,0) INTO v_成本价 FROM 材料特性 WHERE 材料id=v_药品收发记录.药品id;
				EXCEPTION 
					WHEN OTHERS THEN v_成本价:=0;
				END ;
				IF v_成本价=0 THEN 
					V_出库差价 := round(V_药品收发记录.零售金额 * V_差价率,4);
				ELSE 
					V_出库差价 :=round(v_药品收发记录.零售金额-v_药品收发记录.实际数量*v_成本价,4);
				END IF ;
			ELSE 
				V_出库差价 :=round( V_药品收发记录.零售金额 * V_差价率,V_小数);
			END IF ;
		    ELSE
			V_差价率 := V_实际库存差价 / V_实际库存金额;
			V_出库差价 :=round( V_药品收发记录.零售金额 * V_差价率,V_小数);
		    END IF;


		    IF NVL (V_药品收发记录.实际数量,0) = 0 THEN
			V_成本价 := (V_药品收发记录.零售金额 - V_出库差价);
		    ELSE
			V_成本价 :=(V_药品收发记录.零售金额 - V_出库差价) / V_药品收发记录.实际数量;
		    END IF;
  
		    v_成本价:=nvl(v_成本价,0);
		    V_成本金额 := round(V_成本价 * V_药品收发记录.实际数量,V_小数);
		ELSE
		    BEGIN
			SELECT SUM(成本价) INTO V_成本价
			FROM (
			    SELECT DECODE(SIGN(NVL(C.实际金额,0)),1,(D.现价-D.现价*(C.实际差价/C.实际金额)),(D.现价-D.现价*(B.指导差价率/100)))
				*(A.分子/A.分母) AS 成本价
			    FROM 自制材料构成 A,
				(SELECT j.材料id AS 药品id,j.指导差价率 FROM 材料特性 j,收费项目目录 q  WHERE j.材料id=q.id and  nvl(q.是否变价,0)=0) B,
				(SELECT 库房ID,药品ID,实际金额,实际差价 FROM 药品库存 WHERE 性质 = 1 AND 库房ID = V_药品收发记录.对方部门ID ) C,
				(SELECT 收费细目ID,现价 FROM 收费价目 WHERE ((SYSDATE BETWEEN 执行日期 AND 终止日期) OR (SYSDATE >= 执行日期 AND 终止日期 IS NULL))
				) D
			    WHERE A.原料材料ID = B.药品ID AND B.药品ID = D.收费细目ID AND B.药品ID = C.药品ID (+)
				AND A.自制材料ID = V_药品收发记录.药品ID
			    UNION all
			    SELECT DECODE(SIGN(NVL(C.实际金额,0)),1,(C.现价-C.现价*(C.实际差价/C.实际金额)),(C.现价-C.现价*(B.指导差价率/100)))*(A.分子/A.分母) AS 成本价
			    FROM 自制材料构成 A,
				(SELECT j.材料id AS 药品id,j.指导差价率 FROM 材料特性 j,收费项目目录 q  WHERE j.材料id=q.id and  nvl(q.是否变价,0)=1) B,
				(SELECT 库房ID,药品ID,实际金额,实际差价,实际金额/实际数量 AS 现价 FROM 药品库存 WHERE 性质 = 1 AND 库房ID = V_药品收发记录.对方部门ID AND 实际数量>0 ) C 
			    WHERE A.原料材料ID = B.药品ID AND B.药品ID = C.药品ID AND A.自制材料ID = V_药品收发记录.药品ID);
		    EXCEPTION
			WHEN OTHERS THEN
			    V_成本价 := 0;
		    END;

		    V_成本金额 := V_成本价 * V_药品收发记录.实际数量;
		    V_出库差价 := V_药品收发记录.零售金额 - V_成本金额;

			--更新该材料的成本价
			UPDATE 材料特性
			SET 成本价=V_成本价 
			WHERE 材料ID=V_药品收发记录.药品ID;
		    
		END IF;

		UPDATE 药品收发记录
		SET 成本价 = V_成本价,成本金额 = V_成本金额,差价 = V_出库差价
		WHERE ID = V_药品收发记录.ID;

		UPDATE 药品库存
		SET 可用数量=NVL(可用数量,0)+DECODE(V_药品收发记录.入出系数,1,NVL(V_药品收发记录.实际数量,0),0),
		    实际数量=NVL(实际数量,0)+NVL(V_药品收发记录.实际数量,0)*V_药品收发记录.入出系数,
		    实际金额=NVL(实际金额,0)+NVL(V_药品收发记录.零售金额,0)*V_药品收发记录.入出系数,
		    实际差价=NVL(实际差价,0)+V_出库差价*V_药品收发记录.入出系数
		WHERE 库房ID = V_药品收发记录.库房ID
		    AND 药品ID = V_药品收发记录.药品ID
		    AND NVL (批次,0) = 0
		    AND 性质 = 1;

		IF SQL%NOTFOUND THEN
		Insert INTO 药品库存
		    (库房ID,药品ID,性质,可用数量,实际数量,实际金额,实际差价)
		VALUES (V_药品收发记录.库房ID,V_药品收发记录.药品ID,1,
		    DECODE (V_药品收发记录.入出系数,1,NVL (V_药品收发记录.实际数量,0),0),
		    V_药品收发记录.实际数量 * V_药品收发记录.入出系数,
		    V_药品收发记录.零售金额 * V_药品收发记录.入出系数,
		    V_出库差价 * V_药品收发记录.入出系数);
		END IF;

		DELETE
		FROM 药品库存
		WHERE 库房ID = V_药品收发记录.库房ID
		    AND 药品ID = V_药品收发记录.药品ID
		    AND NVL (可用数量,0) = 0
		    AND NVL (实际数量,0) = 0
		    AND NVL (实际金额,0) = 0
		    AND NVL (实际差价,0) = 0;

		--更改药品收发汇总表的相应数据
		UPDATE 药品收发汇总
		SET 数量 =NVL(数量,0)+NVL(V_药品收发记录.实际数量,0)*V_药品收发记录.入出系数,
		    金额 =NVL(金额,0)+NVL (V_药品收发记录.零售金额,0) * V_药品收发记录.入出系数,
		    差价 =NVL (差价,0) + V_出库差价 * V_药品收发记录.入出系数
		WHERE 日期 = TRUNC (SYSDATE)
		    AND 库房ID = V_药品收发记录.库房ID
		    AND 药品ID = V_药品收发记录.药品ID
		    AND 类别ID = V_药品收发记录.入出类别ID
		    AND 单据 = 16;

		IF SQL%NOTFOUND THEN
		    Insert INTO 药品收发汇总
			(日期,库房ID,药品ID,类别ID,单据,数量,金额,差价)
		    VALUES (
			TRUNC (SYSDATE),
			V_药品收发记录.库房ID,
			V_药品收发记录.药品ID,
			V_药品收发记录.入出类别ID,
			16,
			V_药品收发记录.实际数量 * V_药品收发记录.入出系数,
			V_药品收发记录.零售金额 * V_药品收发记录.入出系数,
			V_出库差价 * V_药品收发记录.入出系数
			);
		END IF;
    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_自制材料入库_verify;
/





-----------------------------------------------------------
--材料领用的审核处理
--说明：首先对药品收发记录表中的审核人和审核时间及实际数量进行处理，
--接着对药品库存和材料收发汇总表中的相应数量和金额进行处理
--特别说明：对药品库存和收发汇总表的处理是分开的，
--
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_材料领用_verify (
	序号_IN		IN 药品收发记录.序号%TYPE,
	NO_IN		IN 药品收发记录.NO%TYPE,
	库房ID_IN	IN 药品收发记录.库房ID%TYPE,
	对方部门ID_IN   IN 药品收发记录.对方部门ID%TYPE,
	材料ID_IN	IN 药品收发记录.药品ID%TYPE,
	产地_IN		IN 药品收发记录.产地%TYPE,
	批次_IN		IN 药品收发记录.批次%TYPE,
	填写数量_IN	IN 药品收发记录.填写数量%TYPE,
	实际数量_IN	IN 药品收发记录.实际数量%TYPE,
	成本价_IN	IN 药品收发记录.成本价%TYPE,
	成本金额_IN	IN 药品收发记录.成本金额%TYPE,
	零售金额_IN	IN 药品收发记录.零售金额%TYPE,
	差价_IN		IN 药品收发记录.差价%TYPE,
	入出类别ID_IN   IN 药品收发记录.入出类别ID%TYPE,
	审核人_IN	IN 药品收发记录.审核人%TYPE,
	审核日期_IN	IN 药品收发记录.审核日期%TYPE,
	批号_IN		IN 药品收发记录.批号%TYPE := NULL,
	效期_IN		IN 药品收发记录.效期%TYPE := NULL
)
IS
	mErrMsg		varchar2(100);
	mErrItem		EXCEPTION;
	V_可用数量		药品库存.可用数量%TYPE;
	V_编码		收费项目目录.编码%TYPE;
	V_实际库存金额      药品库存.实际金额%TYPE;
	V_实际库存差价      药品库存.实际差价%TYPE;
	V_差价率            number(18,8);
	V_出库差价		药品库存.实际差价%TYPE;
	V_成本价            药品收发记录.成本价%TYPE;
	V_成本金额		药品收发记录.成本金额%TYPE;
	V_小数		number(2);
	v_负成本计算		系统参数表.参数值%type;
BEGIN
    	SELECT 参数值 INTO v_负成本计算 FROM 系统参数表 WHERE 参数号=120;

	BEGIN 
	    SELECT  to_number(Nvl(参数值,缺省值),'99999')  INTO v_小数 
	    From 系统参数表 where 参数名='费用金额保留位数';
	EXCEPTION 
		WHEN OTHERS THEN 
			v_小数:=2;		
	END;

	--由于领用处理允许在审核时改变实际数量，
	--所以首先对实际数量和其他相应的字段进行更新。

	BEGIN
		SELECT nvl(实际金额,0),nvl(实际差价,0),nvl(可用数量,0)    INTO V_实际库存金额,V_实际库存差价,V_可用数量
		FROM 药品库存
		WHERE 药品ID = 材料ID_IN     AND NVL(批次,0) = 批次_IN AND 库房ID = 库房ID_IN AND 性质 = 1 AND ROWNUM = 1;

	EXCEPTION
		WHEN OTHERS THEN
		    V_实际库存金额 := 0;
		    V_可用数量 := 0;
	END;

	IF V_实际库存金额 <= 0 THEN
		BEGIN
		    SELECT 指导差价率 / 100,nvl(成本价,0)    INTO V_差价率,v_成本价
		    FROM 材料特性
		    WHERE 材料ID = 材料ID_IN;
		EXCEPTION
		    WHEN OTHERS THEN
			V_差价率 := 0;
		END;
		IF v_负成本计算 ='1' THEN 
			IF v_成本价=0 THEN 
				V_差价率 := V_实际库存差价 / V_实际库存金额;
				V_出库差价 :=round( 零售金额_IN * V_差价率,v_小数);
			ELSE 
				V_出库差价 :=round( 零售金额_IN-实际数量_IN * V_成本价,v_小数);
			END IF ;
		ELSE 
			V_出库差价 :=round( 零售金额_IN * V_差价率,v_小数);
		END IF ;
	ELSE
		V_差价率 := V_实际库存差价 / V_实际库存金额;
		V_出库差价 :=round( 零售金额_IN * V_差价率,v_小数);
	END IF;

	IF 实际数量_IN=0 THEN 
		V_成本价 :=成本价_IN; 
	ELSE 
		V_成本价 := (零售金额_IN - V_出库差价) / 实际数量_IN; 
	END IF; 

	V_成本金额 := round(V_成本价 * 实际数量_IN,v_小数);

	UPDATE 药品收发记录
	SET 审核人 = NVL (审核人_IN,审核人),
		审核日期 = 审核日期_IN,
		实际数量 = 实际数量_IN,
		成本价 = V_成本价,
		成本金额 = V_成本金额,
		零售金额 = 零售金额_IN,
		差价 = V_出库差价
	WHERE NO = NO_IN
		AND 单据 =20
		AND 药品ID = 材料ID_IN
		AND 序号 = 序号_IN
		AND 记录状态 = 1
		AND 审核人 IS NULL;

	--更改药品库存的相应数据
	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
		RAISE mErrItem;
	END IF;

	IF 批次_IN > 0 AND (V_可用数量 + 填写数量_IN - 实际数量_IN) < 0 THEN
		SELECT 编码 INTO V_编码 FROM 收费项目目录  WHERE ID = 材料ID_IN;
		mErrMsg:='[ZLSOFT]编码为' || V_编码 || ',批号为' || 批号_IN ||'的分批核算材料' || CHR (10) ||CHR (13) ||'可用库存数量不够！[ZLSOFT]';
		RAISE mErrItem;
	END IF;

	UPDATE 药品库存
	SET 可用数量 = NVL (可用数量,0) + 填写数量_IN - 实际数量_IN,
		实际数量 = NVL (实际数量,0) - 实际数量_IN,
		实际金额 = NVL (实际金额,0) - 零售金额_IN,
		实际差价 = NVL (实际差价,0) - V_出库差价
	WHERE 库房ID = 库房ID_IN
		AND 药品ID = 材料ID_IN
		AND NVL (批次,0) = NVL (批次_IN,0)
		AND 性质 = 1;

	IF SQL%NOTFOUND THEN
		Insert INTO 药品库存(库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价)
		VALUES (库房ID_IN,材料ID_IN,批次_IN,1,-实际数量_IN,-实际数量_IN,-零售金额_IN,-V_出库差价);
	END IF;

	DELETE
	FROM 药品库存
	WHERE 库房ID = 库房ID_IN
		AND 药品ID = 材料ID_IN
		AND NVL (可用数量,0) = 0
		AND NVL (实际数量,0) = 0
		AND NVL (实际金额,0) = 0
		AND NVL (实际差价,0) = 0;

	--更材料收发汇总表的相应数据
	UPDATE 药品收发汇总
	SET 数量 = NVL (数量,0) - 实际数量_IN,
		金额 = NVL (金额,0) - 零售金额_IN,
		差价 = NVL (差价,0) - V_出库差价
	WHERE 日期 = TRUNC (SYSDATE)
		AND 库房ID = 库房ID_IN
		AND 药品ID = 材料ID_IN
		AND 类别ID = 入出类别ID_IN
		AND 单据 = 20;

	IF SQL%NOTFOUND THEN
		Insert INTO 药品收发汇总
		    (日期,库房ID,药品ID,类别ID,单据,数量,金额,差价)
		VALUES (TRUNC (SYSDATE),库房ID_IN,材料ID_IN,入出类别ID_IN,20,-实际数量_IN,-零售金额_IN,-V_出库差价    );
	END IF;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料领用_verify;
/




-----------------------------------------------------------
--材料其他出库的审核处理
--说明：首先对药品收发记录表中的审核人和审核时间及实际数量进行处理，
--接着对药品库存和药品收发汇总表中的相应数量和金额进行处理
--特别说明：对药品库存和收发汇总表的处理是分开的，
--
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_材料其他出库_verify (
	序号_IN		IN 药品收发记录.序号%TYPE,
	NO_IN		IN 药品收发记录.NO%TYPE,
	库房ID_IN	IN 药品收发记录.库房ID%TYPE,
	材料ID_IN	IN 药品收发记录.药品ID%TYPE,
	批次_IN         IN 药品收发记录.批次%TYPE,
	实际数量_IN     IN 药品收发记录.实际数量%TYPE,
	成本价_IN	IN 药品收发记录.成本价%TYPE,
	成本金额_IN     IN 药品收发记录.成本金额%TYPE,
	零售金额_IN     IN 药品收发记录.零售金额%TYPE,
	差价_IN         IN 药品收发记录.差价%TYPE,
	入出类别ID_IN   IN 药品收发记录.入出类别ID%TYPE,
	审核人_IN	IN 药品收发记录.审核人%TYPE,
	审核日期_IN	IN 药品收发记录.审核日期%TYPE
)
IS
	mErrMsg            varchar2(100);
	mErrItem        EXCEPTION;

	V_可用数量        药品库存.可用数量%TYPE;
	V_编码            收费项目目录.编码%TYPE;
	V_实际库存金额        药品库存.实际金额%TYPE;
	V_实际库存差价        药品库存.实际差价%TYPE;
	V_差价率            number(18,8);
	V_出库差价        药品库存.实际差价%TYPE;
	V_成本价            药品收发记录.成本价%TYPE;
	V_成本金额        药品收发记录.成本金额%TYPE;
	V_小数		number(2);
	v_负成本计算		系统参数表.参数值%type;

BEGIN
	BEGIN 
		SELECT  to_number(Nvl(参数值,缺省值),'99999')  INTO v_小数 
		From 系统参数表 where 参数名='费用金额保留位数';
	EXCEPTION 
		WHEN OTHERS THEN 
			v_小数:=2;		
	END;
    	SELECT 参数值 INTO v_负成本计算 FROM 系统参数表 WHERE 参数号=120;



	--由于领用处理允许在审核时改变实际数量，
	--所以首先对实际数量和其他相应的字段进行更新。
	BEGIN
		SELECT nvl(实际金额,0),nvl(实际差价,0)    INTO V_实际库存金额,V_实际库存差价
		FROM 药品库存
		WHERE 药品ID = 材料ID_IN
		    AND NVL (批次,0) = 批次_IN
		    AND 库房ID = 库房ID_IN
		    AND 性质 = 1
		    AND ROWNUM = 1;
	EXCEPTION
		WHEN OTHERS THEN
		    V_实际库存金额 := 0;
	END;
	IF V_实际库存金额 <= 0 THEN
		BEGIN
		    SELECT 指导差价率 / 100,nvl(成本价,0)    INTO V_差价率,v_成本价
		    FROM 材料特性
		    WHERE 材料ID = 材料ID_IN;
		EXCEPTION
		    WHEN OTHERS THEN
			V_差价率 := 0;
		END;
		IF v_负成本计算 ='1' THEN 
			IF v_成本价=0 THEN 
				V_差价率 := V_实际库存差价 / V_实际库存金额;
				V_出库差价 :=round( 零售金额_IN * V_差价率,v_小数);
			ELSE 
				V_出库差价 :=round( 零售金额_IN-实际数量_IN * V_成本价,v_小数);
			END IF ;
		ELSE 
			V_出库差价 :=round( 零售金额_IN * V_差价率,v_小数);
		END IF ;
	ELSE
		V_差价率 := V_实际库存差价 / V_实际库存金额;
		V_出库差价 :=round( 零售金额_IN * V_差价率,v_小数);
	END IF;
	IF 实际数量_IN<=0 THEN 
		V_成本价 := 成本价_IN;
	ELSE 
		V_成本价 := (零售金额_IN - V_出库差价) / 实际数量_IN;
	END IF ;
	V_成本金额 :=round( V_成本价 * 实际数量_IN,v_小数);

	UPDATE 药品收发记录
	SET 审核人 = NVL (审核人_IN,审核人),
		审核日期 = 审核日期_IN,
		成本价 = V_成本价,
		成本金额 = V_成本金额,
		差价 = V_出库差价
	WHERE NO = NO_IN
		AND 单据 = 21
		AND 药品ID = 材料ID_IN
		AND 序号 = 序号_IN
		AND 记录状态 = 1
		AND 审核人 IS NULL;

	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
		RAISE mErrItem;
	END IF;

	--更改药品库存的相应数据
	UPDATE 药品库存
	SET 实际数量 = NVL (实际数量,0) - 实际数量_IN,
		实际金额 = NVL (实际金额,0) - 零售金额_IN,
		实际差价 = NVL (实际差价,0) - V_出库差价
	WHERE 库房ID = 库房ID_IN
		AND 药品ID = 材料ID_IN
		AND NVL (批次,0) = NVL (批次_IN,0)
		AND 性质 = 1;

	IF SQL%NOTFOUND THEN
		Insert INTO 药品库存
		    (库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价)
		VALUES (库房ID_IN,材料ID_IN,批次_IN,1,-实际数量_IN,-实际数量_IN,-零售金额_IN,-V_出库差价);
	END IF;

	DELETE
	FROM 药品库存
	WHERE 库房ID = 库房ID_IN
		AND 药品ID = 材料ID_IN
		AND NVL (可用数量,0) = 0
		AND NVL (实际数量,0) = 0
		AND NVL (实际金额,0) = 0
		AND NVL (实际差价,0) = 0;

	--更药品收发汇总表的相应数据
	UPDATE 药品收发汇总
	SET 数量 = NVL (数量,0) - 实际数量_IN,
		金额 = NVL (金额,0) - 零售金额_IN,
		差价 = NVL (差价,0) - V_出库差价
	WHERE 日期 = TRUNC (SYSDATE)
		AND 库房ID = 库房ID_IN
		AND 药品ID = 材料ID_IN
		AND 类别ID = 入出类别ID_IN
		AND 单据 = 21;

	IF SQL%NOTFOUND THEN
		Insert INTO 药品收发汇总
		    (日期,库房ID,药品id ,类别ID,单据,数量,金额,差价)
		VALUES (TRUNC (SYSDATE),库房ID_IN,材料ID_IN,入出类别ID_IN,21,-实际数量_IN,-零售金额_IN,-V_出库差价);
	END IF;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料其他出库_verify;
/




-----------------------------------------------------------
--材料移库的审核处理
--说明：首先对药品收发记录表中的审核人和审核时间及实际数量进行处理，
--接着对药品库存和药品收发汇总表中的相应数量和金额进行处理
--特别说明：对药品库存和收发汇总表的处理是分开的，
------------------------------------------------------------

    
CREATE OR REPLACE PROCEDURE ZL_材料移库_VERIFY (
    序号_IN		IN 药品收发记录.序号%TYPE,
    库房ID_IN		IN 药品收发记录.库房ID%TYPE,
    对方部门ID_IN	IN 药品收发记录.对方部门ID%TYPE,
    材料ID_IN		IN 药品收发记录.药品ID%TYPE,
    产地_IN		IN 药品收发记录.产地%TYPE,
    出批次_IN		IN 药品收发记录.批次%TYPE,
    填写数量_IN		IN 药品收发记录.填写数量%TYPE,
    实际数量_IN		IN 药品收发记录.实际数量%TYPE,
    成本价_IN		IN 药品收发记录.成本价%TYPE,
    成本金额_IN		IN 药品收发记录.成本金额%TYPE,
    零售金额_IN		IN 药品收发记录.零售金额%TYPE,
    差价_IN		IN 药品收发记录.差价%TYPE,
    出类别ID_IN		IN 药品收发记录.入出类别ID%TYPE,
    入类别ID_IN		IN 药品收发记录.入出类别ID%TYPE,
    NO_IN		IN 药品收发记录.NO%TYPE,
    审核人_IN		IN 药品收发记录.审核人%TYPE,
    批号_IN		IN 药品收发记录.批号%TYPE := NULL,
    效期_IN		IN 药品收发记录.效期%TYPE := NULL,
    灭菌效期_IN        IN 药品收发记录.灭菌效期%type:=NULL,
    审核日期_IN		IN 药品收发记录.审核日期%TYPE := NULL,
    移库单_IN		IN NUMBER:=1)
IS
	mErrMsg		varchar2(500);
	mErrItem	EXCEPTION;

	V_入批次	药品收发记录.批次%TYPE := NULL;
	V_实际库存金额	药品库存.实际金额%TYPE;
	V_实际库存差价	药品库存.实际差价%TYPE;
	V_差价率	NUMBER(18,8);
	V_出库差价	药品库存.实际差价%TYPE;
	V_成本价	药品收发记录.成本价%TYPE;
	V_成本金额	药品收发记录.成本金额%TYPE;
	V_实际数量	药品库存.实际数量%TYPE;
	V_编码		收费项目目录.编码%TYPE;
	v_小数		NUMBER ;
	v_负成本计算		系统参数表.参数值%type;

BEGIN
	--获取金额小数位数
	SELECT to_number(Nvl(参数值,缺省值),'99999') INTO v_小数 FROM 系统参数表 WHERE 参数名='费用金额保留位数';
    	SELECT 参数值 INTO v_负成本计算 FROM 系统参数表 WHERE 参数号=120;

	--由于移库处理允许在审核时改变实际数量，
        --所以首先对实际数量和其他相应的字段进行更新。
	BEGIN
		SELECT NVL(实际金额,0), NVL(实际差价,0), NVL(实际数量,0) INTO V_实际库存金额, V_实际库存差价, V_实际数量
		FROM 药品库存
		WHERE 药品ID = 材料ID_IN
			AND NVL (批次, 0) = 出批次_IN
			AND 库房ID = 库房ID_IN
			AND 性质 = 1
			AND ROWNUM = 1;
	EXCEPTION
		WHEN OTHERS THEN
		    V_实际库存金额 := 0;
		    V_实际数量 := 0;
	END;

	IF V_实际库存金额 <= 0 THEN
		BEGIN
		    SELECT 指导差价率 / 100,nvl(成本价,0)    INTO V_差价率,v_成本价
		    FROM 材料特性
		    WHERE 材料ID = 材料ID_IN;
		EXCEPTION
		    WHEN OTHERS THEN
			V_差价率 := 0;
		END;
		IF v_负成本计算 ='1' THEN 
			IF v_成本价=0 THEN 
				V_差价率 := V_实际库存差价 / V_实际库存金额;
				V_出库差价 :=round( 零售金额_IN * V_差价率,v_小数);
			ELSE 
				V_出库差价 :=round( 零售金额_IN-实际数量_IN * V_成本价,v_小数);
			END IF ;
		ELSE 
			V_出库差价 :=round( 零售金额_IN * V_差价率,v_小数);
		END IF ;
	ELSE
		V_差价率 := V_实际库存差价 / V_实际库存金额;
		V_出库差价 :=round( 零售金额_IN * V_差价率,v_小数);
	END IF;


	IF 实际数量_IN=0 THEN
		V_成本价 :=成本价_IN;
	ELSE
		V_成本价 := (零售金额_IN - V_出库差价) / 实际数量_IN; 
	END IF; 
	
	V_成本金额 := ROUND(V_成本价 * 实际数量_IN,v_小数);

	UPDATE 药品收发记录
	SET 审核人 = NVL (审核人_IN, 审核人),
	     审核日期 = 审核日期_IN,
	     实际数量 = 实际数量_IN,
	     成本价 = V_成本价,
	     成本金额 = V_成本金额,
	     零售金额 = 零售金额_IN,
	     差价 = V_出库差价
	WHERE NO = NO_IN
	AND 单据 = 19
	AND 药品ID = 材料ID_IN
	AND 记录状态 = 1
	AND 序号 IN (序号_IN, 序号_IN + 1)
	AND 审核人 IS NULL;

	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
		RAISE mErrItem;
	END IF;

	IF 出批次_IN > 0 THEN
		IF V_实际数量 < 实际数量_IN THEN
			SELECT 编码 INTO V_编码 FROM 收费项目目录 WHERE ID = 材料ID_IN;
			mErrMsg:= '[ZLSOFT]编码为' || V_编码 || ',批号为' || 批号_IN ||'的库房分批材料' ||
				CHR(10) || CHR(13) || '可用库存数量不够！[ZLSOFT]';
			RAISE mErrItem;
		END IF;
	END IF;

	--取入类别的批次
	SELECT 批次 INTO V_入批次 FROM 药品收发记录 WHERE NO = NO_IN AND 单据 = 19 AND 记录状态 = 1 AND 序号 = 序号_IN+1;
        
	--更改入类别的材料库存的相应数据

	UPDATE 药品库存
	SET 可用数量 = NVL (可用数量, 0) + 实际数量_IN,
		实际数量 = NVL (实际数量, 0) + 实际数量_IN,
		实际金额 = NVL (实际金额, 0) + 零售金额_IN,
		实际差价 = NVL (实际差价, 0) + V_出库差价,
		上次采购价 = V_成本价,
		上次批号 = NVL (批号_IN, 上次批号),
		上次产地 = NVL (产地_IN, 上次产地),
		灭菌效期=nvl(灭菌效期_IN,灭菌效期)
	WHERE 库房ID = 对方部门ID_IN AND 药品ID = 材料ID_IN AND NVL (批次, 0) = NVL (V_入批次, 0) AND 性质 = 1;

	IF SQL%NOTFOUND THEN
		INSERT INTO 药品库存
			(库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次采购价,上次批号,上次产地,效期,灭菌效期)
		VALUES 
			(对方部门ID_IN,材料ID_IN,V_入批次,1,实际数量_IN,实际数量_IN,零售金额_IN,V_出库差价,V_成本价,批号_IN,产地_IN,效期_IN,灭菌效期_IN);
	END IF;

	--更改出类别的材料库存的相应数据

	UPDATE 药品库存
	SET	实际数量 = NVL (实际数量, 0) - 实际数量_IN,
		实际金额 = NVL (实际金额, 0) - 零售金额_IN,
		实际差价 = NVL (实际差价, 0) - V_出库差价
	WHERE 库房ID = 库房ID_IN AND 药品ID = 材料ID_IN
		AND NVL (批次, 0) = NVL (出批次_IN, 0)
		AND 性质 = 1;

	IF SQL%NOTFOUND THEN
		INSERT INTO 药品库存(库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次批号,效期,灭菌效期)
		VALUES (库房ID_IN,材料ID_IN,出批次_IN,1,0,-实际数量_IN,-零售金额_IN,-V_出库差价,批号_IN,效期_IN,灭菌效期_IN);
	END IF;

	DELETE
	FROM 药品库存
	WHERE 库房ID = 库房ID_IN
		AND 药品ID = 材料ID_IN
		AND NVL (可用数量, 0) = 0
		AND NVL (实际数量, 0) = 0
		AND NVL (实际金额, 0) = 0
		AND NVL (实际差价, 0) = 0;

	--更改入类别的药品收发汇总表的相应数据
	UPDATE 药品收发汇总
	SET	数量 = NVL (数量, 0) + 实际数量_IN,
		金额 = NVL (金额, 0) + 零售金额_IN,
		差价 = NVL (差价, 0) + V_出库差价
	WHERE 日期 = TRUNC (SYSDATE)
		AND 库房ID = 对方部门ID_IN
		AND 药品ID = 材料ID_IN
		AND 类别ID = 入类别ID_IN
		AND 单据 = 19;

	IF SQL%NOTFOUND THEN
		INSERT INTO 药品收发汇总
			(日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
		VALUES(
			TRUNC(SYSDATE),对方部门ID_IN,材料ID_IN,入类别ID_IN,19,实际数量_IN,零售金额_IN,V_出库差价);
	END IF;

	--更改出类别的药品收发汇总表的相应数据
	UPDATE 药品收发汇总
	SET 数量 = NVL (数量, 0) - 实际数量_IN,
		金额 = NVL (金额, 0) - 零售金额_IN,
		差价 = NVL (差价, 0) - V_出库差价
	WHERE 日期 = TRUNC (SYSDATE)
		AND 库房ID = 库房ID_IN
		AND 药品ID = 材料ID_IN
		AND 类别ID = 出类别ID_IN
		AND 单据 = 19;

	IF SQL%NOTFOUND THEN
		INSERT INTO 药品收发汇总
				(日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
			VALUES (
				TRUNC(SYSDATE),库房ID_IN,材料ID_IN,出类别ID_IN,19,-实际数量_IN,-零售金额_IN,-V_出库差价);
	END IF;
EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101, mErrMsg);
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_材料移库_VERIFY;
/

CREATE OR REPLACE PROCEDURE zl_成药品种_INSERT(
    类别_IN IN 诊疗项目目录.类别%TYPE := NULL,
    分类ID_IN IN 诊疗项目目录.分类ID%TYPE := NULL,
    ID_IN IN 诊疗项目目录.ID%TYPE,
    编码_IN IN 诊疗项目目录.编码%TYPE := NULL,
    名称_IN IN 诊疗项目目录.名称%TYPE := NULL,
    拼音_IN IN 诊疗项目别名.简码%TYPE := NULL,
    五笔_IN IN 诊疗项目别名.简码%TYPE := NULL,
    英文_IN IN 诊疗项目别名.名称%TYPE := NULL,
    单位_IN IN 诊疗项目目录.计算单位%TYPE := NULL,
    药品剂型_IN IN 药品特性.药品剂型%TYPE := NULL,
    毒理分类_IN IN 药品特性.毒理分类%TYPE := NULL,
    价值分类_IN IN 药品特性.价值分类%TYPE := NULL,
    货源情况_IN IN 药品特性.货源情况%TYPE := NULL,
    用药梯次_IN IN 药品特性.用药梯次%TYPE := NULL,
    药品类型_IN IN 药品特性.药品类型%TYPE := NULL,
    处方职务_IN IN 药品特性.处方职务%TYPE := '00',
    处方限量_IN IN 药品特性.处方限量%TYPE := NULL,
    急救药否_IN IN 药品特性.急救药否%TYPE := 0,
    是否新药_IN IN 药品特性.是否新药%TYPE := 0,
    是否原料_IN IN 药品特性.是否原料%TYPE := 0,
    是否皮试_IN IN 药品特性.是否皮试%TYPE := 0,
    参考目录Id_IN In 诊疗项目目录.参考目录Id%Type:=Null,
    品种医嘱_IN In 药品特性.品种医嘱%TYPE := 0,
    其他别名_IN IN Varchar2      --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织

) IS
    v_Records VARCHAR2(4000);   --临时记录别名数据的字符串
    v_CurrRec VARCHAR2(1000);   --包含在别名记录中的一条别名
    v_Fields  VARCHAR2(1000);   --临时记录一条别名的字符串
    v_名称 诊疗项目目录.名称%TYPE;
    v_拼音 诊疗项目别名.简码%TYPE;
    v_五笔 诊疗项目别名.简码%TYPE;
	v_诊疗项目ID number(18);
BEGIN
    INSERT INTO 诊疗项目目录(类别,分类ID,ID,编码,名称,计算单位,
        计算方式,执行频率,适用性别,单独应用,组合项目,执行安排,计价性质,服务对象,建档时间,撤档时间,参考目录Id)
    VALUES (类别_IN,分类ID_IN,ID_IN,编码_IN,名称_IN,单位_IN,
        1,0,0,1,0,0,0,3,sysdate,to_date('3000-01-01','YYYY-MM-DD'),参考目录Id_IN);

    INSERT INTO 药品特性(药名ID,药品剂型,毒理分类,价值分类,货源情况,用药梯次,
        药品类型,处方职务,处方限量,急救药否,是否新药,是否原料,是否皮试,品种医嘱)
    VALUES (ID_IN,药品剂型_IN,毒理分类_IN,价值分类_IN,货源情况_IN,用药梯次_IN,
        药品类型_IN,处方职务_IN,处方限量_IN,急救药否_IN,是否新药_IN,是否原料_IN,是否皮试_IN,品种医嘱_IN);

    IF 拼音_IN IS NOT NULL THEN
        INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,名称_IN,1,拼音_IN,1);
    END IF;
    IF 五笔_IN IS NOT NULL THEN
        INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,名称_IN,1,五笔_IN,2);
    END IF;
    IF 英文_IN IS NOT NULL THEN
        INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,英文_IN,2,null,0);
    END IF;

    IF 其他别名_IN IS NULL THEN
        v_Records := null;
    ELSE
        v_Records := 其他别名_IN||'|';
    END IF;
    WHILE v_Records IS NOT NULL LOOP
        v_CurrRec:=substr(v_Records,1,instr(v_Records,'|')-1);
        v_Fields:=v_CurrRec;
        v_名称:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_拼音:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_五笔:=v_Fields;
        IF V_拼音 IS NOT NULL THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,v_名称,9,v_拼音,1);
        END IF;
        IF v_五笔 IS NOT NULL THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,v_名称,9,v_五笔,2);
        END IF;
        v_Records:=replace('|'||v_Records,'|'||v_CurrRec||'|');
    END LOOP;
    --添加缺省的对应输出单据
    INSERT INTO 诊疗单据应用(病历文件id,应用场合,诊疗项目id)
    SELECT A.病历文件id,1,ID_IN
    FROM 诊疗单据应用 A,诊疗项目目录 I
    Where A.诊疗项目id=I.Id And I.类别=类别_IN And 应用场合=1 And Rownum<2;
    INSERT INTO 诊疗单据应用(病历文件id,应用场合,诊疗项目id)
    SELECT A.病历文件id,2,ID_IN
    FROM 诊疗单据应用 A,诊疗项目目录 I
    Where A.诊疗项目id=I.Id And I.类别=类别_IN And 应用场合=2 And Rownum<2;
	
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_成药品种_INSERT;
/

CREATE OR REPLACE PROCEDURE zl_成药品种_UPDATE(
    分类ID_IN IN 诊疗项目目录.分类ID%TYPE := NULL,
    ID_IN IN 诊疗项目目录.ID%TYPE,
    编码_IN IN 诊疗项目目录.编码%TYPE := NULL,
    名称_IN IN 诊疗项目目录.名称%TYPE := NULL,
    拼音_IN IN 诊疗项目别名.简码%TYPE := NULL,
    五笔_IN IN 诊疗项目别名.简码%TYPE := NULL,
    英文_IN IN 诊疗项目别名.名称%TYPE := NULL,
    单位_IN IN 诊疗项目目录.计算单位%TYPE := NULL,
    药品剂型_IN IN 药品特性.药品剂型%TYPE := NULL,
    毒理分类_IN IN 药品特性.毒理分类%TYPE := NULL,
    价值分类_IN IN 药品特性.价值分类%TYPE := NULL,
    货源情况_IN IN 药品特性.货源情况%TYPE := NULL,
    用药梯次_IN IN 药品特性.用药梯次%TYPE := NULL,
    药品类型_IN IN 药品特性.药品类型%TYPE := NULL,
    处方职务_IN IN 药品特性.处方职务%TYPE := '00',
    处方限量_IN IN 药品特性.处方限量%TYPE := NULL,
    急救药否_IN IN 药品特性.急救药否%TYPE := 0,
    是否新药_IN IN 药品特性.是否新药%TYPE := 0,
    是否原料_IN IN 药品特性.是否原料%TYPE := 0,
    是否皮试_IN IN 药品特性.是否皮试%TYPE := 0,
    参考目录Id_IN In 诊疗项目目录.参考目录Id%Type:=Null,
    品种医嘱_IN In 药品特性.品种医嘱%TYPE := 0,
    其他别名_IN IN Varchar2      --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织
) IS
    v_Records VARCHAR2(4000);   --临时记录别名数据的字符串
    v_CurrRec VARCHAR2(1000);   --包含在别名记录中的一条别名
    v_Fields  VARCHAR2(1000);   --临时记录一条别名的字符串
    v_名称 诊疗项目目录.名称%TYPE;
    v_拼音 诊疗项目别名.简码%TYPE;
    v_五笔 诊疗项目别名.简码%TYPE;
    Err_NotFind  EXCEPTION;
BEGIN
    UPDATE 诊疗项目目录
    SET 分类ID=分类ID_IN,编码=编码_IN,名称=名称_IN,计算单位=单位_IN,参考目录Id=参考目录Id_IN
    WHERE ID=ID_IN;
    IF SQL%ROWCOUNT=0 THEN
        RAISE Err_NotFind;
    END IF;

    UPDATE 药品特性
    SET 药品剂型=药品剂型_IN,毒理分类=毒理分类_IN,价值分类=价值分类_IN,货源情况=货源情况_IN,用药梯次=用药梯次_IN,
        药品类型=药品类型_IN,处方职务=处方职务_IN,处方限量=处方限量_IN,
        急救药否=急救药否_IN,是否新药=是否新药_IN,是否原料=是否原料_IN,是否皮试=是否皮试_IN,品种医嘱=品种医嘱_IN
    WHERE 药名ID=ID_IN;

    update 收费项目目录
    set 名称=名称_IN
    where ID in (select 药品id from 药品规格 where 药名ID=ID_IN);

    IF 拼音_IN IS NULL THEN
        DELETE FROM 诊疗项目别名 WHERE 诊疗项目ID=ID_IN AND 性质=1 AND 码类=1;
        DELETE FROM 收费项目别名 
        WHERE 收费细目id in (select 药品id from 药品规格 where 药名id=ID_IN) AND 性质=1 AND 码类=1;
    ELSE
        UPDATE 诊疗项目别名 SET 名称=名称_IN, 简码=拼音_IN WHERE 诊疗项目ID=ID_IN AND 性质=1 AND 码类=1;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,名称_IN,1,拼音_IN,1);
        END IF;
        for r_Spec in (
            select 药品id from 药品规格 where 药名id=ID_IN)
        loop
            update 收费项目别名 SET 名称=名称_IN, 简码=拼音_IN WHERE 收费细目ID=r_Spec.药品id AND 性质=1 AND 码类=1;
            if sql%rowcount=0 then
               insert into 收费项目别名(收费细目ID,名称,性质,简码,码类) values(r_Spec.药品id,名称_IN,1,拼音_IN,1);
            end if;
        end loop;
    END IF;
    IF 五笔_IN IS NULL THEN
        DELETE FROM 诊疗项目别名 WHERE 诊疗项目ID=ID_IN AND 性质=1 AND 码类=2;
        DELETE FROM 收费项目别名 
        WHERE 收费细目id in (select 药品id from 药品规格 where 药名id=ID_IN) AND 性质=1 AND 码类=2;
    ELSE
        UPDATE 诊疗项目别名 SET 名称=名称_IN, 简码=五笔_IN WHERE 诊疗项目ID=ID_IN AND 性质=1 AND 码类=2;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,名称_IN,1,五笔_IN,2);
        END IF;
        for r_Spec in (
            select 药品id from 药品规格 where 药名id=ID_IN)
        loop
            update 收费项目别名 SET 名称=名称_IN, 简码=五笔_IN WHERE 收费细目ID=r_Spec.药品id AND 性质=1 AND 码类=2;
            if sql%rowcount=0 then
               insert into 收费项目别名(收费细目ID,名称,性质,简码,码类) values(r_Spec.药品id,名称_IN,1,五笔_IN,2);
            end if;
        end loop;
    END IF;
    IF 英文_IN IS NULL THEN
        DELETE FROM 诊疗项目别名 WHERE 诊疗项目ID=ID_IN AND 性质=2;
        DELETE FROM 收费项目别名 
        WHERE 收费细目id in (select 药品id from 药品规格 where 药名id=ID_IN) AND 性质=2;
    ELSE
        UPDATE 诊疗项目别名 SET 名称=英文_IN WHERE 诊疗项目ID=ID_IN AND 性质=2;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,英文_IN,2,null,0);
        END IF;
        for r_Spec in (
            select 药品id from 药品规格 where 药名id=ID_IN)
        loop
            update 收费项目别名 SET 名称=英文_IN WHERE 收费细目ID=r_Spec.药品id AND 性质=2;
            if sql%rowcount=0 then
               insert into 收费项目别名(收费细目ID,名称,性质,简码,码类) values(r_Spec.药品id,英文_IN,2,null,0);
            end if;
        end loop;
    END IF;

    DELETE FROM 诊疗项目别名 WHERE 诊疗项目ID=ID_IN AND 性质=9;
    DELETE FROM 收费项目别名 
    WHERE 收费细目id in (select 药品id from 药品规格 where 药名id=ID_IN) AND 性质=9;
    IF 其他别名_IN IS NULL THEN
        v_Records := null;
    ELSE
        v_Records := 其他别名_IN||'|';
    END IF;
    WHILE v_Records IS NOT NULL LOOP
        v_CurrRec:=substr(v_Records,1,instr(v_Records,'|')-1);
        v_Fields:=v_CurrRec;
        v_名称:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_拼音:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_五笔:=v_Fields;
        IF V_拼音 IS NOT NULL THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,v_名称,9,v_拼音,1);
            insert into 收费项目别名(收费细目ID,名称,性质,简码,码类) 
            select 药品id,v_名称,9,v_拼音,1 from 药品规格 where 药名id=ID_IN;
        END IF;
        IF v_五笔 IS NOT NULL THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,v_名称,9,v_五笔,2);
            insert into 收费项目别名(收费细目ID,名称,性质,简码,码类) 
            select 药品id,v_名称,9,v_五笔,2 from 药品规格 where 药名id=ID_IN;
        END IF;
        v_Records:=replace('|'||v_Records,'|'||v_CurrRec||'|');
    END LOOP;

EXCEPTION
    WHEN Err_NotFind THEN
        Raise_application_error (-20101, '[ZLSOFT]该品种不存在，可能已被其他用户删除！[ZLSOFT]');
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_成药品种_UPDATE;
/


-----------------------------------------------------------
-- 材料外购入库的审核处理
--说明：首先对材料收发记录表中的审核人和审核时间进行处理，
--接着对材料库存和材料收发汇总表中的相应数量和金额进行处理
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE ZL_材料外购_VERIFY (
    NO_IN        IN 药品收发记录.NO%TYPE := NULL,
    审核人_IN    IN 药品收发记录.审核人%TYPE := NULL
)
IS
    mErrItem        EXCEPTION;
    mErrMsg        varchar2(100);

    V_BATCHCOUNT    INTEGER;        --原不分批现在分批的材料的数量
    V_单位ID        药品收发记录.供药单位ID%TYPE;

    V_发票金额    应付记录.发票金额%TYPE;
    V_库存金额    NUMBER(16,5);
    V_库存差价    NUMBER(16,5);
    V_库存数量    NUMBER(16,5);
    V_成本价        NUMBER(16,5);

    CURSOR C_药品收发记录    IS
    SELECT ID,实际数量,零售金额,差价,库房ID,药品ID,批次,供药单位ID,成本价,批号,效期,灭菌效期,生产日期,产地,入出类别ID
    FROM 药品收发记录
    WHERE NO = NO_IN AND 单据 = 15    AND 记录状态 = 1
    ORDER BY 药品ID;
BEGIN

    UPDATE 药品收发记录
    SET 审核人 = NVL (审核人_IN,审核人),审核日期 = SYSDATE
    WHERE NO = NO_IN AND 单据 = 15 AND 记录状态 = 1 AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        mErrMsg:='[ZLSOFT]该单据已经被他人审核或删除，不能进行审核！[ZLSOFT]';
        RAISE mErrItem;
    END IF;

    --主要针对原不分批现在分批的材料，不能对其审核
    SELECT COUNT(*) INTO V_BATCHCOUNT 
    FROM    药品收发记录 A,材料特性 B
    WHERE    A.药品ID=B.材料ID AND A.NO=NO_IN     AND A.单据=15 AND A.记录状态=1    AND NVL(A.批次,0)=0
        AND ((NVL(B.库房分批,0)=1 AND 
        A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 
                 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室'))) OR NVL(B.在用分批,0)=1);

    IF V_BATCHCOUNT>0 THEN
        mErrMsg:='[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能审核！[ZLSOFT]';
        RAISE mErrItem;
    END IF;

    --原分批现不分批的材料,在审核时，要处理他
    UPDATE 药品收发记录 SET 批次=0
    WHERE    ID IN (    SELECT ID FROM 药品收发记录 A,材料特性 B
            WHERE A.药品ID=B.材料ID    AND A.NO=NO_IN    AND A.单据 = 15    AND A.记录状态 = 1
                AND NVL(A.批次,0)>0 AND (NVL(B.库房分批,0)=0 OR    (NVL(B.在用分批,0)=0 AND
                A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室'))))
            );

    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        --更改材料库存表的相应数据

        UPDATE 药品库存
        SET    可用数量 = NVL (可用数量,0) + NVL (V_药品收发记录.实际数量,0),
            实际数量 = NVL (实际数量,0) + NVL (V_药品收发记录.实际数量,0),
            实际金额 = NVL (实际金额,0) + NVL (V_药品收发记录.零售金额,0),
            实际差价 = NVL (实际差价,0) + NVL (V_药品收发记录.差价,0),
            上次供应商ID = NVL (V_药品收发记录.供药单位ID,上次供应商ID),
            上次采购价 = NVL (V_药品收发记录.成本价,上次采购价),
            上次批号 = NVL (V_药品收发记录.批号,上次批号),
            上次产地 = NVL (V_药品收发记录.产地,上次产地),
            灭菌效期=NVL (V_药品收发记录.灭菌效期,灭菌效期),
            上次生产日期 = NVL (V_药品收发记录.生产日期,上次生产日期),
            效期 = NVL (V_药品收发记录.效期,效期)
        WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次,0) = NVL (V_药品收发记录.批次,0)
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            INSERT INTO 药品库存
                (库房ID,药品ID,批次,   性质,可用数量,实际数量,实际金额,实际差价,上次供应商ID,上次采购价,上次批号,上次生产日期,上次产地,灭菌效期,效期)
            VALUES (
                V_药品收发记录.库房ID,
                V_药品收发记录.药品ID,
                V_药品收发记录.批次,
                1,
                V_药品收发记录.实际数量,
                V_药品收发记录.实际数量,
                V_药品收发记录.零售金额,
                V_药品收发记录.差价,
                V_药品收发记录.供药单位ID,
                V_药品收发记录.成本价,
                V_药品收发记录.批号,
                V_药品收发记录.生产日期,
                V_药品收发记录.产地,
                V_药品收发记录.灭菌效期,
                V_药品收发记录.效期);
        END IF;

        --清除数量金额为零的记录
        DELETE
        FROM 药品库存
        WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL(可用数量,0) = 0
            AND NVL(实际数量,0) = 0
            AND NVL(实际金额,0) = 0
            AND NVL(实际差价,0) = 0;

        --更改材料收发汇总表的相应数据

        UPDATE 药品收发汇总 SET 
            数量 = NVL (数量,0) + NVL (V_药品收发记录.实际数量,0),
            金额 = NVL (金额,0) + NVL (V_药品收发记录.零售金额,0),
            差价 = NVL (差价,0) + NVL (V_药品收发记录.差价,0)
        WHERE 日期 = TRUNC (SYSDATE)
            AND 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND 类别ID = V_药品收发记录.入出类别ID
            AND 单据 = 15;

        IF SQL%NOTFOUND THEN
            INSERT INTO 药品收发汇总
                (日期,库房ID,药品ID,类别ID,单据,数量,金额,差价)
            VALUES (
                  TRUNC (SYSDATE),
                  V_药品收发记录.库房ID,
                  V_药品收发记录.药品ID,
                  V_药品收发记录.入出类别ID,
                  15,
                  V_药品收发记录.实际数量,
                  V_药品收发记录.零售金额,
                  V_药品收发记录.差价
                  );
        END IF;

        --更新该材料的成本价
        BEGIN 
            SELECT SUM(NVL(实际金额,0)),SUM(NVL(实际差价,0)),SUM(NVL(实际数量,0))
                INTO V_库存金额,V_库存差价,V_库存数量
            FROM 药品库存
            WHERE 性质=1 and 药品ID=V_药品收发记录.药品ID;
        EXCEPTION 
            WHEN OTHERS THEN V_库存数量:=0;
        END ;

	--更新该药品的成本价
	UPDATE 材料特性
	SET 成本价=V_药品收发记录.成本价 
	WHERE 材料ID=V_药品收发记录.药品ID;

    END LOOP;


    --对应付余额表进行处理
    --此处用一个块，主要是解决没有对应发票号的记录
    BEGIN
        UPDATE 应付记录
        SET 审核人=审核人_IN,审核日期=SYSDATE
        WHERE 入库单据号=NO_IN AND 系统标识=5 and 记录性质=0 And 记录状态=1;

        SELECT B.单位ID,SUM (发票金额)    INTO V_单位ID,V_发票金额
        FROM 药品收发记录 A,应付记录 B
        WHERE A.ID = B.收发ID
            AND A.NO = NO_IN
            AND A.单据 = 15 AND B.系统标识=5
        GROUP BY B.单位ID;

        IF NVL (V_单位ID,0) <> 0 THEN
            UPDATE 应付余额    SET 
                金额 = NVL (金额,0) + NVL (V_发票金额,0)
            WHERE 单位ID = V_单位ID    AND 性质 = 1;

            IF SQL%NOTFOUND THEN
                INSERT INTO 应付余额 (单位ID,性质,金额)
                VALUES (V_单位ID,1,V_发票金额);
            END IF;
        END IF;
    EXCEPTION
        WHEN NO_DATA_FOUND THEN
            NULL;
    END;

EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101,mErrMsg);
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE,SQLERRM);
END ZL_材料外购_VERIFY;
/
-----------------------------------------------------------
-- 材料外购入库的冲销处理
--说明：首先改原单据的记录状态为3或+3;
--再生成一张单据号相同，记录状态为2或+3，数量和金额为负的冲销单据;
--同时增加相应的应付记录;
--最后更改药品库存表和药品收发汇总表的相应数量和金额。
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_材料外购_STRIKE (
    行次_IN        IN INTEGER,
    原记录状态_IN    IN 药品收发记录.记录状态%TYPE,
    NO_IN        IN 药品收发记录.NO%TYPE,
    序号_IN        IN 药品收发记录.序号%TYPE,
    材料id_IN    IN 药品收发记录.药品ID%TYPE,
    冲销数量_IN    IN 药品收发记录.实际数量%TYPE,
    填制人_IN    IN 药品收发记录.填制人%TYPE ,
    填制日期_IN    IN 药品收发记录.填制日期%TYPE ,
    发票号_IN    IN 应付记录.发票号%TYPE := NULL,
    发票日期_IN    IN 应付记录.发票日期%TYPE := NULL,
    发票金额_IN    IN 应付记录.发票金额%TYPE := NULL,
    全部冲销_IN    IN 药品收发记录.实际数量%TYPE := 0    --用于财务审核
) 
IS
    mErrItem    EXCEPTION ;
    mErrMsg        varchar2(100);

    V_BATCHCOUNT    INTEGER;    --原不分批现在分批的材料的数量 
    v_应付ID        应付记录.ID%TYPE;
    V_库房ID        药品收发记录.库房ID%TYPE; 
    V_供药单位ID    药品收发记录.供药单位ID%TYPE; 
    V_入出类别ID    药品收发记录.入出类别ID%TYPE ;
    V_产地        药品收发记录.产地%TYPE ; 
    V_批次        药品收发记录.批次%TYPE ; 
    V_批号        药品收发记录.批号%TYPE ; 
    V_效期        药品收发记录.效期%TYPE ; 
    V_成本价        药品收发记录.成本价%TYPE ; 
    V_成本金额    药品收发记录.成本金额%TYPE ; 
    V_扣率        药品收发记录.扣率%TYPE ; 
    V_零售价        药品收发记录.零售价%TYPE ; 
    V_零售金额    药品收发记录.零售金额%TYPE ; 
    V_差价        药品收发记录.差价%TYPE ; 
    V_摘要        药品收发记录.摘要%TYPE ; 
    V_剩余数量    药品收发记录.实际数量%TYPE; 
    V_剩余成本金额 药品收发记录.成本金额%Type;
    V_剩余零售金额 药品收发记录.零售金额%Type;

    V_入出系数    药品收发记录.入出系数%TYPE; 
    V_冲销数量    药品收发记录.实际数量%TYPE;
    V_灭菌效期    药品收发记录.灭菌效期%TYPE; 
    v_灭菌日期    药品收发记录.灭菌日期%TYPE; 
    v_生产日期    药品收发记录.生产日期%TYPE; 

    V_记录数 NUMBER; 
    V_收发ID        药品收发记录.ID%TYPE; 

    --对冲销数量进行检查
    V_库存数        药品库存.实际数量%TYPE;
    V_库房分批    INTEGER;
    V_在用分批    INTEGER;
    V_分批属性    INTEGER;
    v_库房        INTEGER;
    v_记录状态	  药品收发记录.记录状态%type;
    V_分批        NUMBER;
    V_小数		number(2);
    V_发票金额	  NUMBER(16,5);

BEGIN
    BEGIN 
	    SELECT  to_number(Nvl(参数值,缺省值),'99999')  INTO v_小数 
	    From 系统参数表 where 参数名='费用金额保留位数';
    EXCEPTION 
	WHEN OTHERS THEN 
		v_小数:=2;		
    END;

    V_冲销数量:=冲销数量_IN;
    IF 行次_IN =1 THEN
        UPDATE 药品收发记录 
        SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3) 
        WHERE NO = NO_IN 
            AND 单据 = 15 
            AND 记录状态 =原记录状态_IN ; 

        IF SQL%ROWCOUNT = 0 THEN 
            mErrMsg:='[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
            RAISE mErrItem; 
        END IF; 
    END IF;
    
    --主要针对原不分批现在分批的材料，不能对其审核 
    SELECT COUNT(*) INTO V_BATCHCOUNT 
    FROM     药品收发记录 A,材料特性 b
    WHERE A.药品ID=B.材料id    AND A.NO=NO_IN     AND A.单据=15 AND MOD(A.记录状态,3)=0
        AND NVL(A.批次,0)=0 AND A.药品ID+0=材料id_IN
        AND ((NVL(B.在用分批,0)=1 AND 
        A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')))    or nvl(b.在用分批,0)=1); 
    
    IF V_BATCHCOUNT>0 THEN 
        mErrMsg:='[ZLSOFT]该单据中第' || 序号_IN || '行的材料原来不分批，现在分批的材料，不能冲销！[ZLSOFT]';
        RAISE mErrItem; 
    END IF;
        
    SELECT SUM(A.实际数量) AS 剩余数量,SUM(A.成本金额) AS 剩余成本金额,SUM(A.零售金额) AS 剩余零售金额,A.库房ID,A.供药单位ID,A.入出类别ID,A.入出系数,NVL(A.批次,0),A.产地,A.批号,A.效期,A.灭菌效期,A.灭菌日期,A.生产日期,A.成本价,A.扣率,A.零售价,A.摘要,B.库房分批,B.在用分批
        INTO  V_剩余数量,v_剩余成本金额,V_剩余零售金额,V_库房ID,V_供药单位ID,V_入出类别ID,V_入出系数,V_批次,V_产地,V_批号,V_效期,v_灭菌效期,v_灭菌日期,v_生产日期,V_成本价,V_扣率,V_零售价,V_摘要,V_库房分批,V_在用分批
    FROM 药品收发记录 A,材料特性 B
    WHERE A.NO=NO_IN And A.药品ID=B.材料ID AND A.单据=15 AND A.药品ID+0=材料id_IN AND A.序号=序号_IN
    GROUP BY A.库房ID,A.供药单位ID,A.入出类别ID,A.入出系数,NVL(A.批次,0),A.产地,A.批号,A.效期,A.灭菌效期,A.灭菌日期,A.生产日期,A.成本价,A.扣率,A.零售价,A.摘要,B.库房分批,B.在用分批;

    --判断该部门是库房还是发料部门
    BEGIN
        SELECT DISTINCT 0 INTO v_库房
        FROM 部门性质说明
        WHERE (工作性质 ='发料部门' OR 工作性质 = '制剂室')
        AND 部门ID = V_库房ID;
    EXCEPTION 
        WHEN OTHERS THEN v_库房:=1;
    END ;
    
    --根据部门性质,判断分批特性
    IF v_库房=0 THEN 
        v_分批属性:=V_在用分批;
    ELSE
        V_分批属性:=V_库房分批;
    END IF ;

    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    V_分批:=0;
    IF V_分批属性=1 AND V_批次<>0 THEN 
        V_分批:=V_批次;
    END IF ;
    
    --取库存数
    BEGIN
        SELECT Nvl(实际数量,0) INTO V_库存数 
        FROM 药品库存 
        WHERE 库房ID=V_库房ID AND 药品ID=材料id_IN AND Nvl(批次,0)=V_分批 And 性质=1;
    EXCEPTION 
    WHEN OTHERS THEN V_库存数:=0;
    END ;

    IF nvl(V_剩余数量,0)=0 THEN 
            mErrMsg:='[ZLSOFT]该单据中第' || 序号_IN || '行的材料已经被冲销完成,不能再冲！[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

    --如果库存数大于剩余数量,取剩余数量;否则取库存数
    IF V_库存数<V_剩余数量 THEN 
        if 全部冲销_IN=1 then 
            --不允许
            mErrMsg:='[ZLSOFT]该单据中第' || 序号_IN || '行的材料冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]';
            RAISE mErrItem; 
        else
            v_剩余成本金额:=V_库存数/V_剩余数量*v_剩余成本金额;
            V_剩余零售金额:=V_库存数/V_剩余数量*V_剩余零售金额;
            V_剩余数量:=V_库存数;
 
        end if ;
    END IF ;
    
    IF 全部冲销_IN=1  THEN 
        V_冲销数量:=V_剩余数量;
    END IF;

    --冲销数量大于剩余数量，不允许
    IF V_剩余数量<V_冲销数量  THEN
        mErrMsg:='[ZLSOFT]该单据中第' || 序号_IN || '行的材料冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]';
        RAISE mErrItem; 
    END IF;

    V_成本金额:= ROUND(V_冲销数量/v_剩余数量*v_剩余成本金额,v_小数);
    V_零售金额:= ROUND(V_冲销数量/v_剩余数量*V_剩余零售金额,v_小数);
    V_差价:=round(V_零售金额-V_成本金额,v_小数);

    SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;
         
    INSERT INTO 药品收发记录 
    ( ID,记录状态,单据,NO,序号,库房ID,供药单位ID,入出类别ID,入出系数,药品ID,批次,产地,批号,生产日期,效期,灭菌效期,灭菌日期,
    填写数量,实际数量,成本价,成本金额,扣率,零售价,零售金额,差价,摘要,填制人,填制日期,审核人,审核日期 
    ) 
    VALUES (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),15,NO_IN,序号_IN,V_库房ID,V_供药单位ID,
    V_入出类别ID,1,材料id_IN,V_批次,V_产地,V_批号,v_生产日期,V_效期,v_灭菌效期,v_灭菌日期,-V_冲销数量,-V_冲销数量,V_成本价,-V_成本金额,
    V_扣率,V_零售价,-V_零售金额,-V_差价,V_摘要,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN); 


    --对于冲销的单据也应该对应付余额表进行处理 
    --只对填了发票号的记录进行处理 
    V_发票金额:=NVL (发票金额_IN,0);

    IF NVL (发票号_IN,' ') <> ' ' AND NVL (V_发票金额,0)<>0 THEN 
	--对于财务审核的，要将剩余的发票金额全部冲销
	IF 全部冲销_IN=1 THEN
		SELECT SUM(B.发票金额) INTO v_发票金额
		FROM
			(SELECT ID
			FROM 药品收发记录
			WHERE 单据=15 AND NO=NO_IN AND 序号=序号_IN) A,应付记录 B
		WHERE A.ID=B.收发ID AND B.系统标识=5 And B.记录性质<>-1;
	END IF;

	UPDATE 应付余额    SET 金额 = NVL (金额,0) - NVL (v_发票金额,0)
	WHERE 单位ID = V_供药单位ID  AND 性质 = 1; 

        IF SQL%NOTFOUND THEN 
            INSERT INTO 应付余额 (单位ID,性质,金额) VALUES (V_供药单位ID,1,-NVL (v_发票金额,0)); 
        END IF; 
    END IF; 
    
    UPDATE 药品库存 
    SET 可用数量 = NVL (可用数量,0) -V_冲销数量,
         实际数量 = NVL (实际数量,0) - V_冲销数量,
         实际金额 = NVL (实际金额,0) - V_零售金额,
         实际差价 = NVL (实际差价,0) -V_差价,
         上次供应商ID = V_供药单位ID,
         上次采购价 = V_成本价,
         上次批号 = V_批号,
         上次产地 = V_产地,
         灭菌效期=v_灭菌效期,
	       上次生产日期=v_生产日期,
         效期 = V_效期 
    WHERE 库房ID = V_库房ID 
        AND 药品ID = 材料id_IN 
        AND NVL (批次,0) = NVL(V_分批,0) 
        AND 性质 = 1; 
 
    IF SQL%NOTFOUND THEN 
        INSERT INTO 药品库存 
            (库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次供应商ID,上次采购价,上次批号,上次产地,效期,灭菌效期,上次生产日期) 
        VALUES 
            (V_库房ID,材料id_IN,V_分批,1,-V_冲销数量,-V_冲销数量,-V_零售金额,-V_差价,V_供药单位ID,V_成本价,V_批号,V_产地,V_效期,v_灭菌效期,v_生产日期) ;
    END IF; 
 
    --清除数量金额为零的记录 
    DELETE     FROM 药品库存 
    WHERE 库房ID = V_库房ID 
        AND 药品ID = 材料id_IN 
        AND NVL (可用数量,0) = 0 
        AND NVL (实际数量,0) = 0 
        AND NVL (实际金额,0) = 0 
        AND NVL (实际差价,0) = 0; 
 
    --更改药品收发汇总表的相应数据 
    UPDATE 药品收发汇总 
    SET 数量 =NVL(数量,0) - V_冲销数量,
        金额 =NVL (金额,0) -V_零售金额,
        差价 =NVL (差价,0) -V_差价 
    WHERE 日期 = TRUNC (SYSDATE) 
        AND 库房ID = V_库房ID 
        AND 药品ID = 材料id_IN 
        AND 类别ID = V_入出类别ID 
        AND 单据 = 15; 

    IF SQL%NOTFOUND THEN 
        INSERT INTO 药品收发汇总 
            (日期,库房ID,药品ID,类别ID,单据,数量,金额,差价) 
        VALUES ( 
            TRUNC (SYSDATE),V_库房ID,材料id_IN,V_入出类别ID,15,-V_冲销数量,-V_零售金额,-V_差价); 
    END IF; 

    --产生应付记录的冲销记录(先判断应付记录中是否已存在该记录对应的冲销记录,是则更新;否则新增)
    SELECT 应付记录_ID.NEXTVAL INTO V_应付ID FROM DUAL;

	begin 
		select max(记录状态)+3 into v_记录状态
		from 应付记录 
		where (系统标识,记录性质,no,项目id,序号) in (	select 系统标识,记录性质,NO,项目id,序号
								from 应付记录
								where 收发ID=(SELECT ID FROM 药品收发记录 WHERE NO=NO_IN AND 单据=15 AND 序号=序号_IN And Mod(记录状态,3)=0) AND 系统标识=5 and  记录性质=0)
		       and 记录状态<>1 and mod(记录状态,3)<>0;

	exception 
		when others then 
			v_记录状态:=2;
	end ;
	if v_记录状态 is null then 
		v_记录状态:=2;
	end if ;
	if mod(v_记录状态,3)<>2 then 
		v_记录状态:=v_记录状态+1;
	end if ;
	if mod(v_记录状态,3)<>2 then 
		v_记录状态:=v_记录状态+1;
	end if ;
	
    INSERT INTO 应付记录
    (ID,记录性质,记录状态,项目ID,序号,单位ID,NO,系统标识,收发ID,入库单据号,单据金额,发票号,发票日期,发票金额,品名,
		规格,产地,批号,计量单位,数量,采购价,采购金额,填制人,填制日期,审核人,审核日期,摘要)
    SELECT V_应付ID,记录性质,v_记录状态,材料ID_In,序号_IN,单位ID,NO,5,V_收发ID,入库单据号,-1* V_零售金额,发票号,发票日期,-v_发票金额,品名,
    规格,产地,批号,计量单位,-1*V_冲销数量,采购价,-1* 采购价*V_冲销数量,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN,摘要
    FROM 应付记录
    WHERE 收发ID=(SELECT ID FROM 药品收发记录 WHERE NO=NO_IN AND 单据=15 AND 序号=序号_IN And Mod(记录状态,3)=0) AND 系统标识=5 AND 记录性质=0;

    update 应付记录    set 记录状态=3
    WHERE 收发ID=(SELECT ID FROM 药品收发记录 WHERE NO=NO_IN AND 单据=15 AND 序号=序号_IN And Mod(记录状态,3)=0) AND 系统标识=5 AND 记录性质=0;
    
EXCEPTION 
    WHEN mErrItem THEN 
        RAISE_APPLICATION_ERROR ( -20101,mErrMsg); 
    WHEN OTHERS THEN 
        ZL_ERRORCENTER (SQLCODE,SQLERRM); 
END ZL_材料外购_STRIKE; 
/

-----------------------------------------------------------
--材料其他入库的冲销处理
--说明：首先改原单据的记录状态为3;
--再生成一张单据号相同，记录状态为2，数量和金额为负的冲销单据;
--最后更改药品库存表和药品收发汇总表的相应数量和金额。
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_材料其他入库_strike (
    行次_IN        IN INTEGER,
    原记录状态_IN    IN 药品收发记录.记录状态%TYPE,
    NO_IN        IN 药品收发记录.NO%TYPE,
    序号_IN        IN 药品收发记录.序号%TYPE,
    材料ID_IN    IN 药品收发记录.药品ID%TYPE,
    冲销数量_IN    IN 药品收发记录.实际数量%TYPE,
    填制人_IN    IN 药品收发记录.填制人%TYPE,
    填制日期_IN    IN 药品收发记录.填制日期%TYPE
)
IS
    mErrMsg            varchar2(100);
    mErrItem        EXCEPTION;

    v_BatchCount        INTEGER;    --原不分批现在分批的材料的数量

    V_库房ID            药品收发记录.库房ID%TYPE; 
    V_入出类别ID        药品收发记录.入出类别ID%TYPE ;
    V_产地            药品收发记录.产地%TYPE ; 
    V_批次            药品收发记录.批次%TYPE ; 
    V_批号            药品收发记录.批号%TYPE ; 
    V_生产日期        药品收发记录.生产日期%TYPE ; 
    V_效期            药品收发记录.效期%TYPE ; 
    V_成本价            药品收发记录.成本价%TYPE ; 
    V_成本金额        药品收发记录.成本金额%TYPE ; 
    V_扣率            药品收发记录.扣率%TYPE ; 
    V_零售价            药品收发记录.零售价%TYPE ; 
    V_零售金额        药品收发记录.零售金额%TYPE ; 
    V_差价            药品收发记录.差价%TYPE ; 
    V_摘要            药品收发记录.摘要%TYPE ; 
    V_入出系数        药品收发记录.入出系数%TYPE; 
    V_灭菌效期        药品收发记录.灭菌效期%type;
    v_灭菌日期        药品收发记录.灭菌日期%type;
    V_记录数            NUMBER; 
    V_收发ID            药品收发记录.ID%TYPE; 
    V_剩余数量		药品收发记录.实际数量%TYPE; 
    V_剩余成本金额	药品收发记录.成本金额%Type;
    V_剩余零售金额	药品收发记录.零售金额%Type;

    --对冲销数量进行检查
    V_库存数            药品库存.实际数量%TYPE;
    V_库房分批        INTEGER;
    V_在用分批        INTEGER;
    V_分批属性        INTEGER;
    v_库房            INTEGER;
    V_分批            NUMBER;
    V_小数		number(2);
BEGIN
    BEGIN 
	    SELECT  to_number(Nvl(参数值,缺省值),'99999')  INTO v_小数 
	    From 系统参数表 where 参数名='费用金额保留位数';
    EXCEPTION 
	WHEN OTHERS THEN 
		v_小数:=2;		
    END;

    IF 行次_IN =1 THEN
        UPDATE 药品收发记录 
        SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3) 
        WHERE NO = NO_IN AND 单据 = 17 AND 记录状态 =原记录状态_IN ; 

        IF SQL%ROWCOUNT = 0 THEN 
            mErrMsg:='[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
            RAISE mErrItem; 
        END IF; 
    END IF;
    
    --主要针对原不分批现在分批的材料，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount 
    FROM 药品收发记录 a,材料特性 b
    WHERE a.药品id=b.材料id    AND a.no=NO_IN     AND a.单据=17 AND MOD(a.记录状态,3)=0 AND a.药品ID+0=材料ID_IN
        AND nvl(a.批次,0)=0
        AND ((nvl(b.库房分批,0)=1 AND a.库房id not in (select 部门id from  部门性质说明 where (工作性质 LIKE '发料部门') or (工作性质 LIKE '制剂室')))
        or nvl(b.在用分批,0)=1);

    IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能冲销！[ZLSOFT]';
        RAISE mErrItem; 
    END IF;  

    SELECT SUM(A.实际数量) AS 剩余数量,SUM(A.成本金额) AS 剩余成本金额,SUM(A.零售金额) AS 剩余零售金额,A.库房ID,A.入出类别ID,A.入出系数,Nvl(A.批次,0),A.产地,A.批号,A.生产日期,A.效期,A.灭菌效期,A.灭菌日期,A.成本价,A.扣率,A.零售价,
        A.摘要,B.库房分批,B.在用分批    INTO  V_剩余数量,v_剩余成本金额,V_剩余零售金额,V_库房ID,V_入出类别ID,V_入出系数,V_批次,V_产地,V_批号,V_生产日期,V_效期,V_灭菌效期,v_灭菌日期,
        V_成本价,V_扣率,V_零售价,V_摘要,V_库房分批,V_在用分批
    FROM 药品收发记录 A,材料特性 B
    WHERE A.NO=NO_IN AND A.单据=17 AND A.药品ID=B.材料ID AND A.药品ID+0=材料ID_IN AND A.序号=序号_IN
    GROUP BY A.库房ID,A.入出类别ID,A.入出系数,NVL(A.批次,0),A.产地,A.批号,A.生产日期,A.效期,A.灭菌效期,a.灭菌日期,A.成本价,A.扣率,A.零售价,A.摘要,B.库房分批,B.在用分批;

    --判断该部门是库房还是发料部门
    BEGIN
        SELECT DISTINCT 0 INTO v_库房 
        FROM 部门性质说明
        WHERE ((工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')) AND 部门ID = V_库房ID;
    EXCEPTION 
        WHEN OTHERS THEN v_库房:=1;
    END ;
    
    --根据部门性质,判断分批特性
    IF v_库房=0 THEN 
        v_分批属性:=V_在用分批;
    ELSE
        V_分批属性:=V_库房分批;
    END IF ;

    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    V_分批:=0;
    IF V_分批属性=1 AND V_批次<>0 THEN 
        V_分批:=V_批次;
    END IF ;
    
    --取库存数
    BEGIN
        SELECT Nvl(实际数量,0) INTO V_库存数 FROM 药品库存 
        WHERE 库房ID=V_库房ID AND 药品ID=材料ID_IN AND Nvl(批次,0)=V_分批 And 性质=1;
    EXCEPTION 
        WHEN OTHERS THEN V_库存数:=0;
    END ;

    IF nvl(V_剩余数量,0)=0 THEN 
            mErrMsg:='[ZLSOFT]该单据中第' || 序号_IN || '行的材料已经被冲销完成,不能再冲！[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

  
    --如果库存数大于剩余数量,取剩余数量;否则取库存数
    IF V_库存数<V_剩余数量 THEN 
	v_剩余成本金额:=V_库存数/V_剩余数量*v_剩余成本金额;
	V_剩余零售金额:=V_库存数/V_剩余数量*V_剩余零售金额;
	V_剩余数量:=V_库存数;
    END IF ;

    --冲销数量大于剩余数量，不允许
    IF V_剩余数量<冲销数量_IN THEN
        mErrMsg:='[ZLSOFT]该单据中第' || 序号_IN || '行的材料冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]';
        RAISE mErrItem; 
    END IF;

    V_成本金额:= ROUND(冲销数量_IN/v_剩余数量*v_剩余成本金额,v_小数);
    V_零售金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余零售金额,v_小数);
    V_差价:=round(V_零售金额-V_成本金额,v_小数);


    SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;

    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,药品ID,批次,产地,批号,生产日期,
        效期,灭菌日期,灭菌效期,填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,审核人,审核日期)
    VALUES 
        (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),17,NO_IN,序号_IN,V_库房ID,V_入出类别ID,
        V_入出系数,材料ID_IN,V_批次,V_产地,V_批号,V_生产日期,V_效期,v_灭菌日期,v_灭菌效期,-冲销数量_IN,-冲销数量_IN,V_成本价,-V_成本金额,
        V_零售价,-V_零售金额,-V_差价,V_摘要,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN);

    --更改药品库存表的相应数据

    UPDATE 药品库存
    SET 可用数量 = NVL (可用数量,0) - NVL (冲销数量_IN,0),
        实际数量 = NVL (实际数量,0) - NVL (冲销数量_IN,0),
        实际金额 = NVL (实际金额,0) - NVL (V_零售金额,0),
        实际差价 = NVL (实际差价,0) - NVL (V_差价,0),
        上次采购价 = NVL (V_成本价,上次采购价),
        上次批号 = NVL (V_批号,上次批号),
        上次产地 = NVL (V_产地,上次产地),
	上次生产日期=nvl(V_生产日期,上次生产日期),
        灭菌效期=nvl(v_灭菌效期,灭菌效期),
        效期 = NVL (V_效期,效期)
    WHERE 库房ID = V_库房ID
        AND 药品ID = 材料ID_IN
        AND NVL (批次,0) = NVL (V_分批,0)
        AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品库存
            (库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,
            实际差价,上次采购价,上次批号,上次生产日期,上次产地,效期,灭菌效期)
        VALUES (
            V_库房ID,材料ID_IN,V_分批,1,-冲销数量_IN,-冲销数量_IN,-V_零售金额,
            -V_差价,V_成本价,V_批号,V_生产日期,V_产地,V_效期,v_灭菌效期);
    END IF;

    DELETE
    FROM 药品库存
    WHERE 库房ID = V_库房ID
        AND 药品ID = 材料ID_IN
        AND nvl(可用数量,0) = 0
        AND nvl(实际数量,0) = 0
        AND nvl(实际金额,0) = 0
        AND nvl(实际差价,0) = 0;

    --更改药品收发汇总表的相应数据

    UPDATE 药品收发汇总
    SET 数量 = NVL (数量,0) - NVL (冲销数量_IN,0),
        金额 = NVL (金额,0) - NVL (V_零售金额,0),
        差价 = NVL (差价,0) - NVL (V_差价,0)
    WHERE 日期 = TRUNC (填制日期_IN)
        AND 库房ID = V_库房ID
        AND 药品ID = 材料ID_IN
        AND 类别ID = V_入出类别ID
        AND 单据 = 17;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品收发汇总
            (日期,库房ID,药品ID,类别ID,单据,数量,金额,差价)
        VALUES (
            TRUNC (填制日期_IN),V_库房ID,材料ID_IN,
            V_入出类别ID,17,-冲销数量_IN,-V_零售金额,-V_差价);
    END IF;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN 
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料其他入库_strike;
/


-----------------------------------------------------------
--材料库存差价调整的冲销处理
--说明：首先改原单据的记录状态为3;
--再生成一张单据号相同，记录状态为2，数量和金额为负的冲销单据;
--最后更改药品库存表和药品收发汇总表的相应数量和金额。
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_材料库存差价调整_strike (
    NO_IN        IN 药品收发记录.NO%TYPE,
    审核人_IN    IN 药品收发记录.审核人%TYPE
)
IS
    mErrMsg            varchar2(100);
    mErrItem        EXCEPTION;
    v_BatchCount        integer;    --原不分批现在分批的材料的数量
    V_COUNT            INTEGER;    --原分批现不分批
    V_批次            药品收发记录.批次%TYPE;

    CURSOR C_药品收发记录    IS
    SELECT 入出类别ID,库房ID,药品id 材料ID,批次,差价,批号,效期,灭菌效期,灭菌日期
    FROM 药品收发记录 A
    WHERE NO = NO_IN
        AND 单据 = 18
        AND 记录状态 = 2
    ORDER BY 药品ID;
BEGIN
    UPDATE 药品收发记录
    SET 记录状态 = 3
    WHERE NO = NO_IN AND 单据 = 18    AND 记录状态 = 1;

    IF SQL%ROWCOUNT = 0 THEN
        mErrMsg:='[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
        RAISE mErrItem; 
    END IF;
    
    --主要针对原不分批现在分批的材料，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount 
    FROM  药品收发记录 a,材料特性 b
    WHERE a.药品id=b.材料id
        AND a.no=NO_IN  AND a.单据=18 AND a.记录状态=3 AND nvl(a.批次,0)=0
        AND ((NVL(B.库房分批,0)=1 AND  
        A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')))
        OR NVL(B.在用分批,0)=1);
    IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能冲销！[ZLSOFT]';
        RAISE mErrItem; 
    END IF;  

    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,药品ID,批次,
        产地,批号,效期,灭菌效期,灭菌日期,零售金额,差价,摘要,填制人,填制日期,审核人,审核日期)
        SELECT 药品收发记录_ID.Nextval,2,单据,NO_IN,序号,库房ID,
            入出类别ID,入出系数,a.药品ID,
            DECODE(NVL(a.批次,0),0,NULL,(DECODE(NVL(b.库房分批,0),0,NULL,a.批次))),
            a.产地,a.批号,a.效期,a.灭菌效期,a.灭菌日期,a.零售金额,-a.差价,a.摘要,
            审核人_IN,SYSDATE,审核人_IN,SYSDATE
        FROM 药品收发记录 a,材料特性 b
        WHERE NO = NO_IN
            AND a.药品id=b.材料id
            AND 单据 = 18
            AND 记录状态 = 3;

    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        --原分批现不分批的材料,在C冲消时，要处理他
        BEGIN 
            SELECT COUNT(*) INTO V_COUNT
            FROM 药品收发记录 A,材料特性 B
            WHERE a.药品ID =b.材料id AND a.药品ID+0=V_药品收发记录.材料ID
                AND A.NO=NO_IN AND A.单据 = 18  and a.库房id+0=V_药品收发记录.库房id
                AND A.记录状态 = 3 AND NVL(A.批次,0)>0
                AND (NVL(B.库房分批,0)=0 OR (NVL(B.在用分批,0)=0 AND
                A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室'))));
        EXCEPTION 
            WHEN OTHERS THEN
                V_COUNT:=0;
        END;
        IF V_COUNT>0 THEN
            V_批次:=0;
        ELSE
            V_批次:=NVL (V_药品收发记录.批次,0);
        END IF;

        --更改药品库存表的相应数据

        UPDATE 药品库存
        SET 实际差价 = NVL (实际差价,0) + NVL (V_药品收发记录.差价,0)
        WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.材料ID
            AND NVL (批次,0) = V_批次
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
                (库房ID,药品ID,批次,性质,实际差价,上次批号,效期,灭菌效期)
            VALUES (
            V_药品收发记录.库房ID,
            V_药品收发记录.材料ID,
            V_批次,
            18,
            V_药品收发记录.差价,
            V_药品收发记录.批号,
            V_药品收发记录.效期,
            v_药品收发记录.灭菌效期
            );
        END IF;

        DELETE
        FROM 药品库存
        WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.材料ID
            AND nvl(可用数量,0) = 0 AND nvl(实际数量,0) = 0 AND nvl(实际金额,0) = 0 AND nvl(实际差价,0) = 0;

        --更改药品收发汇总表的相应数据

        UPDATE 药品收发汇总
        SET 差价 = NVL (差价,0) + NVL (V_药品收发记录.差价,0)
        WHERE 日期 = TRUNC (SYSDATE)
            AND 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.材料ID
            AND 类别ID = V_药品收发记录.入出类别ID
            AND 单据 = 18;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品收发汇总
                (日期,库房ID,药品ID,类别ID,单据,差价)
            VALUES (
                TRUNC (SYSDATE),
                V_药品收发记录.库房ID,
                V_药品收发记录.材料ID,
                V_药品收发记录.入出类别ID,
                18,
                V_药品收发记录.差价);
        END IF;
    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN 
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料库存差价调整_strike;
/

CREATE OR REPLACE PROCEDURE zl_材料移库_Insert (
	NO_IN		IN 药品收发记录.NO%TYPE,
	序号_IN		IN 药品收发记录.序号%TYPE,
	库房ID_IN	IN 药品收发记录.库房ID%TYPE,
	对方部门ID_IN   IN 药品收发记录.对方部门ID%TYPE,
	材料ID_IN	IN 药品收发记录.药品ID%TYPE,
	批次_IN		IN 药品收发记录.批次%TYPE,
	填写数量_IN	IN 药品收发记录.填写数量%TYPE,
	实际数量_IN	IN 药品收发记录.实际数量%TYPE,
	成本价_IN	IN 药品收发记录.成本价%TYPE,
	成本金额_IN	IN 药品收发记录.成本金额%TYPE,
	零售价_IN	IN 药品收发记录.零售价%TYPE,
	零售金额_IN	IN 药品收发记录.零售金额%TYPE,
	差价_IN		IN 药品收发记录.差价%TYPE,
	填制人_IN	IN 药品收发记录.填制人%TYPE,
	产地_IN		IN 药品收发记录.产地%TYPE := NULL,
	批号_IN		IN 药品收发记录.批号%TYPE := NULL,
	效期_IN		IN 药品收发记录.效期%TYPE := NULL,
	灭菌效期_IN	IN 药品收发记录.灭菌效期%TYPE := NULL,
	摘要_IN		IN 药品收发记录.摘要%TYPE := NULL,
	填制日期_IN	IN 药品收发记录.填制日期%TYPE := NULL
	)
IS
	mErrItem	EXCEPTION;
	mErrMsg		varchar2(100);

	v_下库存	系统参数表.参数值%type;
	V_编码		收费项目目录.编码%TYPE;
	V_可用数量	药品库存.可用数量%TYPE;
	V_lngID		药品收发记录.ID%TYPE;--收发ID
	V_入的类别ID    药品收发记录.入出类别ID%TYPE;--入出类别ID
	V_出的类别ID    药品收发记录.入出类别ID%TYPE;--入出类别ID
	V_批次		药品收发记录.批次%TYPE := NULL;--主要针对入库中实行分批核算的材料
	V_是否分批	INTEGER;--判断入库是否分批核算   1:分批；0：不分批
	V_库房分批	INTEGER;--判断入库是否分批核算   1:分批；0：不分批
	V_在用分批	INTEGER;--判断入库是否分批核算   1:分批；0：不分批
	intRecords	NUMBER ;

BEGIN
	BEGIN
		SELECT nvl(参数值,'0') INTO v_下库存 FROM 系统参数表 WHERE 参数号=95;
	EXCEPTION 
		WHEN OTHERS THEN v_下库存:='-99';
	END;

	IF v_下库存='-99' THEN 
		mErrMsg:='[ZLSOFT]在系统参数中无"卫材申领下可用库存"参数,请与系统管员联系![ZLSOFT]';
		RAISE mErrItem;
	END IF ;

	IF 批次_IN > 0 THEN
		BEGIN
		    SELECT 可用数量    INTO V_可用数量
		    FROM 药品库存
		    WHERE 药品ID = 材料ID_IN
			AND NVL (批次,0) = 批次_IN
			AND 库房ID = 库房ID_IN
			AND 性质 = 1
			AND ROWNUM = 1;
		EXCEPTION
		    WHEN OTHERS THEN
			V_可用数量 := 0;
		END;

		IF V_可用数量 - 实际数量_IN < 0 THEN
		    SELECT 编码 INTO V_编码    FROM 收费项目目录 WHERE ID = 材料ID_IN;
		    mErrMsg:='[ZLSOFT]编码为' || V_编码 || ',批号为' || 批号_IN ||'的分批核算材料' || CHR (10) || CHR (13) || '可用库存数量不够！[ZLSOFT]';
		    RAISE mErrItem;
		END IF;
	END IF;

	--首先找出入和出的类别ID
	SELECT B.ID INTO V_入的类别ID 
	FROM 药品单据性质 A,药品入出类别 B 
	WHERE A.类别ID = B.ID AND A.单据 = 34 AND B.系数 = 1 AND ROWNUM < 2;

	SELECT B.ID INTO V_出的类别ID
	FROM 药品单据性质 A,药品入出类别 B
	WHERE A.类别ID = B.ID AND A.单据 = 34 AND B.系数 = -1 AND ROWNUM < 2;

	--插入类别为出的那一笔
	Insert INTO 药品收发记录
		(ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,灭菌效期,
		填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期)
	VALUES (药品收发记录_ID.Nextval,1,19,NO_IN,序号_IN,库房ID_IN,对方部门ID_IN,
		V_出的类别ID,-1,材料ID_IN,批次_IN,产地_IN,批号_IN,效期_IN,灭菌效期_IN,
		填写数量_IN,实际数量_IN,成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,
		差价_IN,摘要_IN,填制人_IN,填制日期_IN);


	IF to_number(v_下库存,'9999')=1 THEN 
		

		UPDATE 药品库存
		SET 可用数量 = NVL(可用数量, 0) - 实际数量_IN
		WHERE 库房ID = 库房ID_IN AND 药品ID = 材料ID_IN 
			AND NVL(批次, 0) = NVL(批次_IN, 0) AND 性质 = 1;

		IF SQL%ROWCOUNT = 0 THEN
			INSERT INTO 药品库存(库房ID, 药品ID, 批次, 性质, 可用数量)
			VALUES(库房ID_IN, 材料ID_IN, NVL(批次_IN, 0), 1, -实际数量_IN);
		END IF;
			
		--同时更新库存数
		DELETE
		FROM 药品库存
		WHERE 库房ID = 库房ID_IN
			AND 药品ID = 材料ID_IN
			AND nvl(可用数量,0) = 0
			AND nvl(实际数量,0) = 0
			AND nvl(实际金额,0) = 0
			AND nvl(实际差价,0) = 0;
	END IF ;

	--下面是判断入库的材料是否是分批核算材料
	SELECT NVL (库房分批,0),nvl(在用分批,0) INTO V_库房分批,v_在用分批
	FROM 材料特性
	WHERE 材料ID = 材料ID_IN;

	V_是否分批 := 0;

	IF v_在用分批=0 then
		IF V_库房分批 = 1 THEN
		    BEGIN
			SELECT DISTINCT 0 INTO V_是否分批
			FROM 部门性质说明
			WHERE ( (工作性质 LIKE '发料部门')    OR (工作性质 LIKE '制剂室')) AND 部门ID = 对方部门ID_IN;
		    EXCEPTION
			WHEN OTHERS THEN
			V_是否分批 := 1;
		    END;
		END IF;
	ELSE 
		V_是否分批 := 1;
	END if;

	SELECT 药品收发记录_ID.Nextval INTO V_lngID FROM Dual;

	IF V_是否分批 = 1 AND NVL (批次_IN,0) = 0 THEN--入库分批且出库不分批
		V_批次 := V_lngID;
	ELSIF    V_是否分批 = 0 THEN--入库不分批
		V_批次 := 0;
	ELSIF NVL (批次_IN,0) <> 0 THEN--入库分批且出库也分批
		V_批次 := 批次_IN;
	END IF;

	--插入类别为入的那一笔
	Insert INTO 药品收发记录
		(ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,灭菌效期,
		填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期)
	VALUES (V_lngID,1,19,NO_IN,序号_IN + 1,对方部门ID_IN,库房ID_IN,V_入的类别ID,
		1,材料ID_IN,V_批次,产地_IN,批号_IN,效期_IN,灭菌效期_IN,填写数量_IN,实际数量_IN,
		成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,差价_IN,摘要_IN,
		填制人_IN,填制日期_IN
		);

	--检查是否存在相同材料相同批次的数据，如果存在不允许保存
	SELECT COUNT(*) INTO intRecords 
	FROM 药品收发记录
	WHERE 单据=19 AND NO=NO_IN AND 入出系数=-1 AND 药品ID+0=材料ID_IN AND Nvl(批次,0)=NVL(批次_IN,0);

	IF intRecords>1 THEN
		SELECT 编码 INTO V_编码 FROM 收费项目目录 WHERE ID = 材料ID_IN;
		mErrMsg:='[ZLSOFT]编码为'||V_编码||'的药品，存在多条重复的记录，请合并为一条记录！[ZLSOFT]';
		RAISE mErrItem;
	END IF ;

EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg  );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料移库_Insert;
/


-----------------------------------------------------------
-- 材料移库的冲销处理
--说明：首先改原单据的记录状态为3;
--再生成一张单据号相同，记录状态为2，数量和金额为负的冲销单据;
-- 对库存表的处理要分开处理，对入类别进行加，出类别药进行减。
--最后更改药品收发汇总表的相应数量和金额。
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE ZL_材料移库_STRIKE (
	行次_IN			IN INTEGER,
	原记录状态_IN		IN 药品收发记录.记录状态%TYPE,
	NO_IN			IN 药品收发记录.NO%TYPE,
	序号_IN			IN 药品收发记录.序号%TYPE,
	材料id_IN		IN 药品收发记录.药品ID%TYPE,
	冲销数量_IN		IN 药品收发记录.实际数量%TYPE,
	填制人_IN		IN 药品收发记录.填制人%TYPE,
	填制日期_IN		IN 药品收发记录.填制日期%TYPE
)
IS
	mErrMsg		varchar2(500);
	mErrItem	EXCEPTION;

	V_BATCHCOUNT INTEGER;    --原不分批现在分批的药品的数量
	V_序号		药品收发记录.序号%TYPE;
	V_库房ID	药品收发记录.库房ID%TYPE;
	V_对方部门ID	药品收发记录.对方部门ID%TYPE;
	V_入出类别ID	药品收发记录.入出类别ID%TYPE ;
	V_产地		药品收发记录.产地%TYPE ;
	V_批次		药品收发记录.批次%TYPE ;
	V_批号		药品收发记录.批号%TYPE ;
	V_效期		药品收发记录.效期%TYPE ;
	V_成本价	药品收发记录.成本价%TYPE ;
	V_成本金额	药品收发记录.成本金额%TYPE ;
	V_扣率		药品收发记录.扣率%TYPE ;
	V_零售价	药品收发记录.零售价%TYPE ;
	V_零售金额	药品收发记录.零售金额%TYPE ;
	V_差价		药品收发记录.差价%TYPE ;
	V_摘要		药品收发记录.摘要%TYPE ;
	V_剩余数量	药品收发记录.实际数量%TYPE; 
	V_剩余成本金额	药品收发记录.成本金额%Type;
	V_剩余零售金额	药品收发记录.零售金额%Type;
	V_入出系数	药品收发记录.入出系数%TYPE;
	V_灭菌日期	药品收发记录.灭菌日期%TYPE;
	V_灭菌效期	药品收发记录.灭菌效期%TYPE;
	V_记录数	NUMBER;
	V_收发ID	药品收发记录.ID%TYPE;
	V_备药人	药品收发记录.配药人%TYPE;
	V_发送日期	药品收发记录.配药日期%TYPE;
 
	--对冲销数量进行检查
	V_库存数	药品库存.实际数量%TYPE;
	V_库房分批	INTEGER;
	V_在用分批	INTEGER;
	V_分批属性	INTEGER;
	V_卫材库		INTEGER;
	V_分批		NUMBER;
	INTDIGIT	NUMBER;
     
BEGIN
         --获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';
 

    IF 行次_IN =1 THEN
	UPDATE 药品收发记录
	SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3)
	WHERE NO = NO_IN AND 单据 = 19 AND 记录状态 =原记录状态_IN ;

	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
		RAISE mErrItem;
	END IF;
    END IF;
 

    SELECT COUNT(*) INTO V_BATCHCOUNT FROM
            药品收发记录 A,材料特性 B
    WHERE A.药品ID=B.材料ID
        AND A.NO=NO_IN
        AND A.单据=19
        AND A.药品ID+0=材料id_IN
        AND MOD(A.记录状态,3)=0
        AND NVL(A.批次,0)=0
        AND ((NVL(B.库房分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%发料部门') OR (工作性质 LIKE '制剂室')))
            OR NVL(B.在用分批,0)=1);
 

    IF V_BATCHCOUNT>0 THEN
	mErrMsg:='[ZLSOFT]该单据中包含有一条原来不分批，现在分批的卫生材料，不能冲销！[ZLSOFT]';
        RAISE mErrItem;
    END IF;
 
    SELECT SUM(A.实际数量) AS 剩余数量,SUM(A.成本金额) AS 剩余成本金额,SUM(A.零售金额) AS 剩余零售金额,A.成本价,A.零售价,A.对方部门ID,NVL(A.批次,0),B.库房分批,B.在用分批
    INTO  V_剩余数量,v_剩余成本金额,V_剩余零售金额,V_成本价,V_零售价,V_库房ID,V_批次,V_库房分批,V_在用分批
    FROM 药品收发记录 A,材料特性 B
    WHERE A.NO=NO_IN AND A.药品ID=B.材料ID AND A.单据=19 AND A.药品ID+0=材料id_IN AND A.序号=序号_IN
    GROUP BY A.成本价,A.零售价,A.对方部门ID,NVL(A.批次,0),B.库房分批,B.在用分批;
 
    --判断该部门是库房还是发料部门
    BEGIN
        SELECT DISTINCT 0   INTO V_卫材库
        FROM 部门性质说明
        WHERE ((工作性质 LIKE '%卫材库')
		OR (工作性质 LIKE '制剂室'))
		AND 部门ID = V_库房ID;
    EXCEPTION
        WHEN OTHERS THEN V_卫材库:=1;
    END ;
 
    --根据部门性质,判断分批特性
    IF V_卫材库=0 THEN
        V_分批属性:=V_在用分批;
    ELSE
        V_分批属性:=V_库房分批;
    END IF ;
 
    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    --因为对于对方库房，相当于出库，而对于当前库房，相当于入库，所以当前库房不予检查，仅检查退库那条记录
    SELECT NVL(A.批次,0) INTO V_批次
    FROM 药品收发记录 A
    WHERE A.NO=NO_IN AND A.单据=19 AND A.药品ID+0=材料id_IN AND A.序号=序号_IN+1 AND MOD(A.记录状态,3)=0;
 
    --取库存数
    BEGIN
        SELECT NVL(实际数量,0) INTO V_库存数 FROM 药品库存
        WHERE 库房ID=V_库房ID AND 药品ID=材料id_IN AND NVL(批次,0)=V_批次 AND 性质=1;
    EXCEPTION
        WHEN OTHERS THEN V_库存数:=0;
    END ;
    
    IF nvl(V_剩余数量,0)=0 THEN 
            mErrMsg:='[ZLSOFT]该单据中第' || ceil(序号_IN/2) || '行的材料已经被冲销完成,不能再冲！[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

    --如果库存数大于剩余数量,取剩余数量;否则取库存数
    IF V_库存数<V_剩余数量 THEN
            v_剩余成本金额:=V_库存数/V_剩余数量*v_剩余成本金额;
            V_剩余零售金额:=V_库存数/V_剩余数量*V_剩余零售金额;
            V_剩余数量:=V_库存数;
    END IF ;
 
    --冲销数量大于剩余数量，不允许
    IF V_剩余数量<冲销数量_IN THEN
	mErrMsg:='[ZLSOFT]该单据中第' || ceil(序号_IN/2) || '行的卫生材料冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]';
        RAISE mErrItem;
    END IF;
 
    V_成本金额:= ROUND(冲销数量_IN/v_剩余数量*v_剩余成本金额,INTDIGIT);
    V_零售金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余零售金额,INTDIGIT);
    V_差价:=round(V_零售金额-V_成本金额,INTDIGIT);





    FOR v_单据 IN (SELECT 序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,灭菌日期,灭菌效期,配药人,配药日期,摘要
			   FROM 药品收发记录
			   WHERE NO = NO_IN AND 单据 = 19 AND (序号>=序号_IN AND 序号<=序号_IN+1) AND (记录状态=1 OR MOD(记录状态,3)=0)
		           ORDER BY 药品ID)
    LOOP
        V_序号:=v_单据.序号;
        V_库房ID:=v_单据.库房ID;
        V_对方部门ID:=v_单据.对方部门ID;
        V_入出类别ID:=v_单据.入出类别ID;
        V_入出系数:=v_单据.入出系数;
        V_批次:=v_单据.批次;
        V_产地:=v_单据.产地;
        V_批号:=v_单据.批号;
        V_效期:=v_单据.效期;
        v_摘要:=v_单据.摘要;
        V_备药人:=v_单据.配药人;
        V_发送日期:=v_单据.配药日期;
	v_灭菌日期:=v_单据.灭菌日期;
	v_灭菌效期:=v_单据.灭菌效期;

        SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;

        INSERT INTO 药品收发记录
        (	ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
		药品ID,批次,产地,批号,效期,灭菌日期,灭菌效期,填写数量,实际数量,成本价,
		成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,审核人,审核日期,配药人,配药日期)
        VALUES
        (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),19,NO_IN,V_序号,V_库房ID,V_对方部门ID,V_入出类别ID,V_入出系数,
        材料id_IN,V_批次,V_产地,V_批号,V_效期,v_灭菌日期,v_灭菌效期,-冲销数量_IN,-冲销数量_IN,V_成本价,-V_成本金额,V_零售价,-V_零售金额,
        -V_差价,V_摘要,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN,V_备药人,V_发送日期);
 
        --原分批现不分批的材料,在冲消时，要处理他
        BEGIN
            SELECT COUNT(*) INTO V_记录数
            FROM 药品收发记录 A, 材料特性 B
            WHERE B.材料ID=A.药品ID
            AND A.药品ID=材料id_IN
            AND A.NO=NO_IN
            AND A.单据 = 19
            AND A.库房ID=V_库房ID
            AND MOD(A.记录状态,3)=0
            AND NVL(A.批次,0)>0
            AND (NVL(B.库房分批,0)=0 OR
                (NVL(B.在用分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%发料部门') OR (工作性质 LIKE '制剂室'))))
            ;
        EXCEPTION
            WHEN OTHERS THEN V_记录数:=0;
        END;
        IF V_记录数>0 THEN
            V_批次:=0;
        ELSE
            V_批次:=NVL (V_批次, 0);
        END IF;
 
        --更改药品库存表的相应数据
        UPDATE 药品库存
            SET 可用数量=NVL(可用数量,0)-NVL(冲销数量_IN,0)*V_入出系数,
                实际数量=NVL(实际数量,0)-NVL(冲销数量_IN,0)*V_入出系数,
                实际金额=NVL(实际金额,0)-NVL(V_零售金额,0)*V_入出系数,
                实际差价=NVL(实际差价,0)-NVL(V_差价,0)*V_入出系数,
                上次采购价=NVL(V_成本价,上次采购价),
                上次批号=NVL(V_批号,上次批号),
                上次产地=NVL(V_产地,上次产地),
                效期=NVL(V_效期,效期)
          WHERE 库房ID = V_库房ID
            AND 药品ID = 材料id_IN
            AND NVL (批次, 0) = V_批次
            AND 性质 = 1;
 
        IF SQL%NOTFOUND THEN
            INSERT INTO 药品库存
            (库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额, 实际差价,上次采购价,上次批号,上次产地,效期,灭菌效期)
            VALUES
            (V_库房ID,材料id_IN,V_批次,1,-冲销数量_IN*V_入出系数,-冲销数量_IN*V_入出系数,
            -V_零售金额*V_入出系数,-V_差价*V_入出系数,V_成本价,V_批号,V_产地,V_效期,v_灭菌效期);
        END IF;
 
        DELETE
          FROM 药品库存
         WHERE 库房ID = V_库房ID
           AND 药品ID = 材料id_IN
           AND NVL(可用数量,0)=0
           AND NVL(实际数量,0)=0
           AND NVL(实际金额,0)=0
           AND NVL(实际差价,0)=0;
 
        --更改药品收发汇总表的相应数据
        UPDATE 药品收发汇总
         SET 数量 =    NVL (数量,0)  - NVL (冲销数量_IN,0)*V_入出系数,
             金额 = NVL (金额, 0) - NVL (V_零售金额, 0)*V_入出系数,
             差价 = NVL (差价, 0) - NVL (V_差价, 0)*V_入出系数
        WHERE 日期 = TRUNC (填制日期_IN)
         AND 库房ID = V_库房ID
         AND 药品ID = 材料id_IN
         AND 类别ID = V_入出类别ID
         AND 单据 = 19;
        IF SQL%NOTFOUND THEN
            INSERT INTO 药品收发汇总
            (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
            VALUES
            (TRUNC (填制日期_IN),V_库房ID,材料id_IN,V_入出类别ID,
            19,-冲销数量_IN*V_入出系数,-V_零售金额*V_入出系数,-V_差价*V_入出系数);
        END IF;
    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101, mErrMsg);
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_材料移库_STRIKE;
/

-----------------------------------------------------------
-- 材料领用的冲销处理
--说明：首先改原单据的记录状态为3;
--再生成一张单据号相同，记录状态为2，数量和金额为负的冲销单据;
--最后更改药品收发汇总表的相应数量和金额。
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_材料领用_STRIKE (
    行次_IN        IN INTEGER,
    原记录状态_IN    IN 药品收发记录.记录状态%TYPE,
    NO_IN        IN 药品收发记录.NO%TYPE,
    序号_IN        IN 药品收发记录.序号%TYPE,
    材料ID_IN    IN 药品收发记录.药品ID%TYPE,
    冲销数量_IN    IN 药品收发记录.实际数量%TYPE,
    填制人_IN    IN 药品收发记录.填制人%TYPE,
    填制日期_IN    IN 药品收发记录.填制日期%TYPE
)
IS
	mErrMsg         varchar2(100);
	mErrItem        EXCEPTION;
	v_BatchCount    INTEGER;    --原不分批现在分批的材料的数量

	V_库房ID        药品收发记录.库房ID%TYPE; 
	V_对方部门ID    药品收发记录.对方部门ID%TYPE;
	V_入出类别ID    药品收发记录.入出类别ID%TYPE ;
	V_产地          药品收发记录.产地%TYPE ; 
	V_批次          药品收发记录.批次%TYPE ; 
	V_批号          药品收发记录.批号%TYPE ; 
	V_效期          药品收发记录.效期%TYPE ; 
	V_成本价        药品收发记录.成本价%TYPE ; 
	V_成本金额      药品收发记录.成本金额%TYPE ; 
	V_扣率          药品收发记录.扣率%TYPE ; 
	V_零售价        药品收发记录.零售价%TYPE ; 
	V_零售金额      药品收发记录.零售金额%TYPE ; 
	V_差价          药品收发记录.差价%TYPE ; 
	V_摘要          药品收发记录.摘要%TYPE ; 
	V_剩余数量	药品收发记录.实际数量%TYPE; 
	V_剩余成本金额	药品收发记录.成本金额%Type;
	V_剩余零售金额	药品收发记录.零售金额%Type;
	V_入出系数      药品收发记录.入出系数%TYPE; 
	V_收发ID        药品收发记录.ID%TYPE; 
	v_灭菌日期        药品收发记录.灭菌日期%TYPE; 
	v_灭菌效期        药品收发记录.灭菌效期%TYPE; 
	V_记录数            NUMBER; 
	V_小数		number(2);
BEGIN
    BEGIN 
	    SELECT  to_number(Nvl(参数值,缺省值),'99999')  INTO v_小数 
	    From 系统参数表 where 参数名='费用金额保留位数';
    EXCEPTION 
	WHEN OTHERS THEN 
		v_小数:=2;		
    END;

    IF 行次_IN =1 THEN
        UPDATE 药品收发记录 
        SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3) 
        WHERE NO = NO_IN AND 单据 = 20 AND 记录状态 =原记录状态_IN ; 

        IF SQL%ROWCOUNT = 0 THEN 
            mErrMsg:='[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
            RAISE mErrItem;
        END IF; 
    END IF;

    --主要针对原不分批现在分批的材料，不能对其审核
    SELECT COUNT(*) INTO V_BATCHCOUNT 
    FROM  药品收发记录 A,材料特性 B
    WHERE A.药品ID=B.材料ID
        AND A.NO=NO_IN     AND A.单据=20 AND Mod(A.记录状态,3)=0 AND NVL(A.批次,0)=0
        AND ((NVL(B.库房分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')))
        OR NVL(B.在用分批,0)=1);
        
    IF V_BATCHCOUNT>0 THEN
        mErrMsg:='[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能冲销！[ZLSOFT]';
        RAISE mErrItem;
    END IF;  

    SELECT SUM(实际数量) AS 剩余数量,SUM(成本金额) AS 剩余成本金额,SUM(零售金额) AS 剩余零售金额,库房ID,对方部门ID,入出类别ID,入出系数,批次,产地,批号,效期,灭菌日期,灭菌效期,成本价,扣率,零售价,摘要
        INTO  V_剩余数量,v_剩余成本金额,V_剩余零售金额,V_库房ID,V_对方部门ID,V_入出类别ID,V_入出系数,V_批次,V_产地,V_批号,V_效期,v_灭菌日期,
        v_灭菌效期,V_成本价,V_扣率,V_零售价,V_摘要
    FROM 药品收发记录 
    WHERE NO=NO_IN 
        AND 单据=20
        AND 药品ID+0=材料ID_IN 
        AND 序号=序号_IN
    GROUP BY 库房ID,对方部门ID,入出类别ID,入出系数,批次,产地,批号,效期,灭菌日期,灭菌效期,成本价,扣率,零售价,摘要;

    IF nvl(V_剩余数量,0)=0 THEN 
            mErrMsg:='[ZLSOFT]该单据中第' || 序号_IN || '行的材料已经被冲销完成,不能再冲！[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

    --冲销数量大于剩余数量，不允许
    IF V_剩余数量<冲销数量_IN THEN
        mErrMsg:='[ZLSOFT]剩余数据不能于被冲销数量,不能冲销！[ZLSOFT]';
        RAISE mErrItem;
    END IF;

    V_成本金额:= ROUND(冲销数量_IN/v_剩余数量*v_剩余成本金额,v_小数);
    V_零售金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余零售金额,v_小数);
    V_差价:=round(V_零售金额-V_成本金额,v_小数);

    SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;

    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
        药品ID,批次,产地,批号,效期,灭菌日期,灭菌效期,填写数量,实际数量,成本价,
        成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,审核人,审核日期)
    VALUES 
        (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),20,NO_IN,序号_IN,V_库房ID,V_对方部门ID,V_入出类别ID,V_入出系数,
        材料ID_IN,V_批次,V_产地,V_批号,V_效期,V_灭菌日期,V_灭菌效期,-冲销数量_IN,-冲销数量_IN,V_成本价,-V_成本金额,V_零售价,-V_零售金额,
        -V_差价,V_摘要,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN);

    --原分批现不分批的材料,在C冲消时，要处理他
    BEGIN 
        SELECT COUNT(*) INTO V_记录数
        FROM 药品收发记录 A,材料特性 B
        WHERE A.药品id=b.材料id AND  B.材料ID+0=材料ID_IN     AND A.NO=NO_IN    AND A.单据 = 20     AND Mod(A.记录状态,3)=0    AND NVL(A.批次,0)>0
        AND (NVL(B.库房分批,0)=0 OR 
            (NVL(B.在用分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室'))));
    EXCEPTION 
        WHEN OTHERS THEN V_记录数:=0;
    END;

    IF V_记录数>0 THEN
        V_批次:=0;
    ELSE
        V_批次:=NVL (V_批次,0);
    END IF;

    --更改药品库存表的相应数据
    UPDATE 药品库存
    SET 可用数量 = NVL (可用数量,0) + NVL (冲销数量_IN,0),
        实际数量 = NVL (实际数量,0) + NVL (冲销数量_IN,0),
        实际金额 = NVL (实际金额,0) + NVL (V_零售金额,0),
        实际差价 = NVL (实际差价,0) + NVL (V_差价,0)
    WHERE 库房ID = V_库房ID
        AND 药品ID = 材料ID_IN
        AND NVL (批次,0) = V_批次
        AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO 药品库存
            (库房ID,药品ID,批次,性质,可用数量,实际数量,
            实际金额,实际差价,上次批号,效期,灭菌效期)
        VALUES 
            (V_库房ID,材料ID_IN,V_批次,1,冲销数量_IN,冲销数量_IN,
            V_零售金额,V_差价,V_批号,V_效期,V_灭菌效期);
    END IF;

    DELETE 药品库存
    WHERE 库房ID = V_库房ID
        AND 药品ID = 材料ID_IN
        AND NVL (可用数量,0) = 0
        AND NVL (实际数量,0) = 0
        AND NVL (实际金额,0) = 0
        AND NVL (实际差价,0) = 0;

    --更改药品收发汇总表的相应数据
    UPDATE 药品收发汇总
    SET 数量 = NVL (数量,0)  +NVL(冲销数量_IN,0),
        金额 = NVL (金额,0) +NVL(V_零售金额,0),
        差价 = NVL (差价,0) +NVL(V_差价,0)
    WHERE 日期 = TRUNC (填制日期_IN)
        AND 库房ID = V_库房ID
        AND 药品ID = 材料ID_IN
        AND 类别ID = V_入出类别ID
        AND 单据 = 20;

    IF SQL%NOTFOUND THEN
        INSERT INTO 药品收发汇总
            (日期,库房ID,药品ID,类别ID,单据,数量,金额,差价)
        VALUES 
            (TRUNC (填制日期_IN),V_库房ID,材料ID_IN,V_入出类别ID,20,冲销数量_IN,V_零售金额,V_差价);
    END IF;
EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101,mErrMsg);
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE,SQLERRM);
END ZL_材料领用_STRIKE;
/



-----------------------------------------------------------
-- 材料其他入库的审核处理
--说明：首先对药品收发记录表中的审核人和审核时间进行处理，
--接着对药品库存和药品收发汇总表中的相应数量和金额进行处理
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_材料其他入库_verify (
    NO_IN        IN 药品收发记录.NO%TYPE := NULL,
    审核人_IN    IN 药品收发记录.审核人%TYPE := NULL
)
IS
    mErrMsg            varchar2(100);
    mErrItem        EXCEPTION;

    v_BatchCount integer;    --原不分批现在分批的材料的数量
 

    CURSOR C_药品收发记录    IS
    SELECT ID,实际数量,零售金额,差价,库房ID,药品ID,批次,成本价,批号,效期,灭菌效期,灭菌日期,产地,入出类别ID,生产日期
    FROM 药品收发记录
    WHERE NO = NO_IN AND 单据 = 17    AND 记录状态 = 1
    ORDER BY 药品ID;
BEGIN
    UPDATE 药品收发记录
    SET 审核人 = 审核人_IN,审核日期 = SYSDATE
    WHERE NO = NO_IN AND 单据 = 17 AND 记录状态 = 1 AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        mErrMsg:='[ZLSOFT]该单据已经被他人审核！[ZLSOFT]';
        RAISE mErrItem;
    END IF;
    
    --主要针对原不分批现在分批的材料，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount 
    FROM   药品收发记录 a,材料特性 b
    WHERE a.药品id=b.材料id AND a.no=NO_IN  AND a.单据=17 AND a.记录状态=1 AND nvl(a.批次,0)=0
        AND ((nvl(b.库房分批,0)=1 AND a.库房id not in (select 部门id from  部门性质说明 where (工作性质 LIKE '发料部门') or (工作性质 LIKE '制剂室')))
            or nvl(b.在用分批,0)=1);
        
    IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能审核！[ZLSOFT';
        RAISE mErrItem;
    END IF;  
    
    --原分批现不分批的材料,在审核时，要处理他
    UPDATE 药品收发记录 SET 批次=0
    WHERE  id=(    SELECT id FROM 药品收发记录 a,材料特性 b 
            WHERE b.材料id=a.药品ID AND a.no=no_in AND a.单据 = 17
                AND a.记录状态 = 1 AND nvl(a.批次,0)>0 AND (nvl(b.库房分批,0)=0 or 
                (nvl(b.在用分批,0)=0 and a.库房id in (select 部门id from  部门性质说明 where (工作性质 LIKE '发料部门') or 
                (工作性质 LIKE '制剂室'))))
            );

    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        --更改药品库存表的相应数据

        UPDATE 药品库存
        SET 可用数量 = NVL (可用数量,0) + NVL (V_药品收发记录.实际数量,0),
            实际数量 = NVL (实际数量,0) + NVL (V_药品收发记录.实际数量,0),
            实际金额 = NVL (实际金额,0) + NVL (V_药品收发记录.零售金额,0),
            实际差价 = NVL (实际差价,0) + NVL (V_药品收发记录.差价,0),
            上次采购价 = NVL (V_药品收发记录.成本价,上次采购价),
            上次批号 = NVL (V_药品收发记录.批号,上次批号),
            上次生产日期 = NVL (V_药品收发记录.生产日期,上次生产日期),
            上次产地 = NVL (V_药品收发记录.产地,上次产地),
            效期 = NVL (V_药品收发记录.效期,效期),
            灭菌效期 = NVL (V_药品收发记录.灭菌效期,灭菌效期)
        WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次,0) = NVL (V_药品收发记录.批次,0)
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存(库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次采购价,上次批号,上次生产日期,上次产地,效期,灭菌效期)
            VALUES (
                V_药品收发记录.库房ID,
                V_药品收发记录.药品ID,
                V_药品收发记录.批次,
                1,
                V_药品收发记录.实际数量,
                V_药品收发记录.实际数量,
                V_药品收发记录.零售金额,
                V_药品收发记录.差价,
                V_药品收发记录.成本价,
                V_药品收发记录.批号,
                V_药品收发记录.生产日期,
                V_药品收发记录.产地,
                V_药品收发记录.效期,
                V_药品收发记录.灭菌效期
                );
        END IF;

        DELETE
        FROM 药品库存
        WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND nvl(可用数量,0) = 0 AND nvl(实际数量,0) = 0 AND nvl(实际金额,0) = 0 AND nvl(实际差价,0) = 0;


        --更改药品收发汇总表的相应数据
        UPDATE 药品收发汇总
        SET 数量 = NVL (数量,0) + NVL (V_药品收发记录.实际数量,0),
            金额 = NVL (金额,0) + NVL (V_药品收发记录.零售金额,0),
            差价 = NVL (差价,0) + NVL (V_药品收发记录.差价,0)
        WHERE 日期 = TRUNC (SYSDATE)
            AND 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND 类别ID = V_药品收发记录.入出类别ID
            AND 单据 = 17;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品收发汇总
                (日期,库房ID,药品ID,类别ID,单据,数量,金额,差价)
            VALUES (
                TRUNC (SYSDATE),
                V_药品收发记录.库房ID,
                V_药品收发记录.药品ID,
                V_药品收发记录.入出类别ID,
                17,
                V_药品收发记录.实际数量,
                V_药品收发记录.零售金额,
                V_药品收发记录.差价
                );
        END IF;
	--更新该材料的成本价
	UPDATE 材料特性
	SET 成本价=V_药品收发记录.成本价 
	WHERE 材料ID=V_药品收发记录.药品ID;

    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料其他入库_verify;
/

-----------------------------------------------------------
-- 材料其他出库的冲销处理
--说明：首先改原单据的记录状态为3;
--再生成一张单据号相同，记录状态为2，数量和金额为负的冲销单据;
--最后更改材料收发汇总表的相应数量和金额。
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_材料其他出库_strike (
    行次_IN            IN INTEGER,
    原记录状态_IN        IN 药品收发记录.记录状态%TYPE,
    NO_IN            IN 药品收发记录.NO%TYPE,
    序号_IN            IN 药品收发记录.序号%TYPE,
    材料ID_IN        IN 药品收发记录.药品ID%TYPE,
    冲销数量_IN        IN 药品收发记录.实际数量%TYPE,
    填制人_IN        IN 药品收发记录.填制人%TYPE,
    填制日期_IN        IN 药品收发记录.填制日期%TYPE
)
IS
    mErrMsg            varchar2(100);
    mErrItem        EXCEPTION;
    v_BatchCount        INTEGER;    --原不分批现在分批的材料的数量

    V_库房ID            药品收发记录.库房ID%TYPE; 
    V_入出类别ID        药品收发记录.入出类别ID%TYPE ;
    V_产地            药品收发记录.产地%TYPE ; 
    V_批次            药品收发记录.批次%TYPE ; 
    V_批号            药品收发记录.批号%TYPE ; 
    V_效期            药品收发记录.效期%TYPE ; 
    V_成本价            药品收发记录.成本价%TYPE ; 
    V_成本金额        药品收发记录.成本金额%TYPE ; 
    V_扣率            药品收发记录.扣率%TYPE ; 
    V_零售价            药品收发记录.零售价%TYPE ; 
    V_零售金额        药品收发记录.零售金额%TYPE ; 
    V_差价            药品收发记录.差价%TYPE ; 
    V_摘要            药品收发记录.摘要%TYPE ; 
    V_剩余数量	药品收发记录.实际数量%TYPE; 
    V_剩余成本金额	药品收发记录.成本金额%Type;
    V_剩余零售金额	药品收发记录.零售金额%Type;

    V_入出系数        药品收发记录.入出系数%TYPE; 

    v_灭菌效期        药品收发记录.灭菌效期%TYPE; 
    V_记录数            NUMBER; 
    V_收发ID            药品收发记录.ID%TYPE; 
    V_小数		number(2);
BEGIN
    BEGIN 
	    SELECT  to_number(Nvl(参数值,缺省值),'99999')  INTO v_小数 
	    From 系统参数表 where 参数名='费用金额保留位数';
    EXCEPTION 
	WHEN OTHERS THEN 
		v_小数:=2;		
    END;
    
    IF 行次_IN =1 THEN
        UPDATE 药品收发记录 
        SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3) 
        WHERE NO = NO_IN AND 单据 = 21 AND 记录状态 =原记录状态_IN ; 
        IF SQL%ROWCOUNT = 0 THEN 
            mErrMsg:='[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
            RAISE mErrItem;
        END IF; 
    END IF;
    
    --主要针对原不分批现在分批的材料，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount 
    FROM 药品收发记录 a,材料特性 b
    WHERE a.药品id=b.材料id    AND a.no=NO_IN     AND a.单据=21    AND A.药品ID+0=材料ID_IN    AND MOD(a.记录状态,3)=0    AND nvl(a.批次,0)=0
        AND ((NVL(B.库房分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')))
            OR NVL(B.在用分批,0)=1);
    
    IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能冲销！[ZLSOFT]';
        RAISE mErrItem;
    END IF;  

    SELECT SUM(实际数量) AS 剩余数量,SUM(成本金额) AS 剩余成本金额,SUM(零售金额) AS 剩余零售金额,库房ID,入出类别ID,入出系数,批次,产地,批号,效期,灭菌效期,成本价,扣率,零售价,摘要
        INTO  V_剩余数量,v_剩余成本金额,V_剩余零售金额,V_库房ID,V_入出类别ID,V_入出系数,V_批次,V_产地,V_批号,V_效期,v_灭菌效期,V_成本价,V_扣率,V_零售价,V_摘要
    FROM 药品收发记录 
    WHERE NO=NO_IN AND 单据=21 AND 药品ID+0=材料ID_IN AND 序号=序号_IN
    GROUP BY 库房ID,入出类别ID,入出系数,批次,产地,批号,效期,灭菌效期,成本价,扣率,零售价,摘要;

    IF nvl(V_剩余数量,0)=0 THEN 
            mErrMsg:='[ZLSOFT]该单据中包含一条的材料已经被冲销完成,不能再冲！[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

    --冲销数量大于剩余数量，不允许
    IF V_剩余数量<冲销数量_IN THEN
        mErrMsg:='[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能冲销！[ZLSOFT]';
        RAISE mErrItem;
    END IF;

    V_成本金额:= ROUND(冲销数量_IN/v_剩余数量*v_剩余成本金额,v_小数);
    V_零售金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余零售金额,v_小数);
    V_差价:=round(V_零售金额-V_成本金额,v_小数);

    SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;

    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,
        药品ID,批次,产地,批号,效期,灭菌效期,填写数量,实际数量,成本价,
        成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,审核人,审核日期)
    VALUES 
        (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),21,NO_IN,序号_IN,V_库房ID,V_入出类别ID,V_入出系数,
        材料ID_IN,V_批次,V_产地,V_批号,V_效期,v_灭菌效期,-冲销数量_IN,-冲销数量_IN,V_成本价,-V_成本金额,V_零售价,-V_零售金额,
        -V_差价,V_摘要,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN);

    --原分批现不分批的材料,在冲消时，要处理他
    BEGIN 
        SELECT COUNT(*) INTO V_记录数
        FROM 药品收发记录 A,材料特性 B
        WHERE A.药品ID=B.材料ID     AND B.材料ID+0=材料ID_IN    AND A.NO=NO_IN    AND A.单据 = 21    
            AND MOD(A.记录状态,3)=0    AND NVL(A.批次,0)>0
            AND (NVL(B.库房分批,0)=0 OR 
            (NVL(B.在用分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室'))));
    EXCEPTION 
        WHEN OTHERS THEN
            V_记录数:=0;
    END;
    IF V_记录数>0 THEN
        V_批次:=0;
    ELSE
        V_批次:=NVL (V_批次,0);
    END IF;
    --更改药品库存表的相应数据
    UPDATE 药品库存
    SET 可用数量 = NVL (可用数量,0) + NVL (冲销数量_IN,0),
        实际数量 = NVL (实际数量,0) + NVL (冲销数量_IN,0),
        实际金额 = NVL (实际金额,0) + NVL (V_零售金额,0),
        实际差价 = NVL (实际差价,0) + NVL (V_差价,0)
    WHERE 库房ID = V_库房ID
        AND 药品ID = 材料ID_IN
        AND NVL (批次,0) = v_批次
        AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品库存
            (库房ID,药品ID,灭菌效期,批次,性质,可用数量,实际数量,实际金额,实际差价,上次批号,效期)
        VALUES 
            (V_库房ID,材料ID_IN,v_灭菌效期,V_批次,1,冲销数量_IN,冲销数量_IN,V_零售金额,V_差价,V_批号,V_效期);
    END IF;

    DELETE
    FROM 药品库存
    WHERE 库房ID = V_库房ID
        AND 药品ID = 材料ID_IN
        AND NVL (可用数量,0) = 0
        AND NVL (实际数量,0) = 0
        AND NVL (实际金额,0) = 0
        AND NVL (实际差价,0) = 0;

    --更改药品收发汇总表的相应数据

    UPDATE 药品收发汇总
    SET 数量 = NVL (数量,0) + NVL (冲销数量_IN,0),
         金额 = NVL (金额,0) + NVL (V_零售金额,0),
         差价 = NVL (差价,0) + NVL (V_差价,0)
    WHERE 日期 = TRUNC (填制日期_IN)
        AND 库房ID = V_库房ID
        AND 药品ID = 材料ID_IN
        AND 类别ID = V_入出类别ID
        AND 单据 = 21;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品收发汇总
            (日期,库房ID,药品ID,类别ID,单据,数量,金额,差价)
        VALUES 
            (TRUNC (填制日期_IN),V_库房ID,材料ID_IN,V_入出类别ID,21,冲销数量_IN,V_零售金额,V_差价);
    END IF;
EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101,mErrMsg);
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE,SQLERRM);
END zl_材料其他出库_strike;
/

-----------------------------------------------------------
-- 材料盘点的冲销处理
--说明：首先改原单据的记录状态为3;
--再生成一张单据号相同，记录状态为2，数量和金额为负的冲销单据;
--最后更改药品收发汇总表的相应数量和金额。
--输出：无
--输入：NO_IN,审核人_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_材料盘点_strike (
    NO_IN        IN 药品收发记录.NO%TYPE,
    审核人_IN    IN 药品收发记录.审核人%TYPE)
IS
    mErrMsg        varchar2(100);
    mErrItem    EXCEPTION;
    v_BatchCount    integer;    --原不分批现在分批的材料的数量
    V_COUNT        INTEGER;    --原分批现不分批
    V_批次        药品收发记录.批次%TYPE;

    CURSOR C_药品收发记录  IS
    SELECT ID,实际数量,零售金额,差价,库房ID,药品ID 材料id,批次,批号,效期,产地,入出类别ID,入出系数
    FROM 药品收发记录
    WHERE NO = NO_IN AND 单据 = 22 AND 记录状态 = 2
    ORDER BY 药品ID;
BEGIN
    UPDATE 药品收发记录
    SET 记录状态 = 3
    WHERE NO = NO_IN AND 单据 = 22     AND 记录状态 = 1;

    IF SQL%ROWCOUNT = 0 THEN
        mErrMsg:='[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]';
        RAISE mErrItem;
    END IF;
   
    --主要针对原不分批现在分批的材料，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount 
    FROM  药品收发记录 a,材料特性 b
    WHERE a.药品id=b.材料id AND a.no=NO_IN  AND a.单据=22 AND a.记录状态=3 AND nvl(a.批次,0)=0
        AND ((NVL(B.库房分批,0)=1 
        AND  A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室'))) OR 
                    NVL(B.在用分批,0)=1);
        IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]该单据中包含有一条原来不分批，现在分批的材料，不能冲销！[ZLSOFT]';
        RAISE mErrItem;
    END IF;  

    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,
        药品ID,批次,产地,批号,效期,灭菌效期,填写数量,扣率,实际数量,
        成本价,成本金额,零售价,零售金额,差价,摘要,
        填制人,填制日期,审核人,审核日期,频次)
        SELECT 药品收发记录_ID.Nextval,2,单据,NO,序号,库房ID,入出类别ID,入出系数,a.药品ID,
            DECODE(NVL(a.批次,0),0,NULL,(DECODE(NVL(b.库房分批,0),0,NULL,a.批次))),
            a.产地,批号,a.效期,a.灭菌效期,填写数量,a.扣率,
            -实际数量,a.成本价,成本金额,零售价,-零售金额,-差价,摘要,
            审核人_IN,SYSDATE,审核人_IN,SYSDATE,频次
        FROM 药品收发记录 a,材料特性 b
        WHERE NO = NO_IN AND a.药品id=b.材料id AND 单据 = 22 AND 记录状态 = 3;

    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        --原分批现不分批的材料,在C冲消时，要处理他
        BEGIN 
            SELECT COUNT(*) INTO V_COUNT
            FROM 药品收发记录 A,材料特性 B
            WHERE B.材料ID+0=V_药品收发记录.材料ID
                AND A.NO=NO_IN AND a.药品id=b.材料id
                AND A.单据 = 22
                and a.库房id+0=V_药品收发记录.库房id
                AND A.记录状态 = 3 
                AND NVL(A.批次,0)>0
                AND (NVL(B.库房分批,0)=0 OR 
                (NVL(B.在用分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室'))));
        EXCEPTION 
            WHEN OTHERS THEN
            V_COUNT:=0;
        END;
        IF V_COUNT>0 THEN
            V_批次:=0;
        ELSE
            V_批次:=NVL (V_药品收发记录.批次,0);
        END IF;

        --更改药品库存表的相应数据
        UPDATE 药品库存
        SET 可用数量=NVL(可用数量,0)+NVL(V_药品收发记录.实际数量,0)*V_药品收发记录.入出系数,
            实际数量=NVL(实际数量,0)+NVL(V_药品收发记录.实际数量,0)*V_药品收发记录.入出系数,
            实际金额=NVL(实际金额,0)+NVL(V_药品收发记录.零售金额,0)*V_药品收发记录.入出系数,
            实际差价=NVL(实际差价,0)+NVL(V_药品收发记录.差价,0)*V_药品收发记录.入出系数
        WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.材料ID
            AND NVL (批次,0) = V_批次
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
                (库房ID,药品ID,批次,性质,可用数量,实际数量,
                实际金额,实际差价,上次批号,上次产地,效期)
            VALUES (
                V_药品收发记录.库房ID,
                V_药品收发记录.材料ID,
                V_批次,
                1,
                V_药品收发记录.实际数量*V_药品收发记录.入出系数,
                V_药品收发记录.实际数量*V_药品收发记录.入出系数,
                V_药品收发记录.零售金额*V_药品收发记录.入出系数,
                V_药品收发记录.差价*V_药品收发记录.入出系数,
                V_药品收发记录.批号,
                V_药品收发记录.产地,
                V_药品收发记录.效期);
        END IF;


        DELETE 
        FROM 药品库存
        WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.材料ID
            AND nvl(可用数量,0)=0 
            And nvl(实际数量,0)=0 
            And nvl(实际金额,0)=0 
            And nvl(实际差价,0)=0;

        --更改药品收发汇总表的相应数据
        UPDATE 药品收发汇总
        SET 数量=NVL(数量,0)+NVL(V_药品收发记录.实际数量,0)*V_药品收发记录.入出系数,
            金额=NVL(金额,0)+NVL(V_药品收发记录.零售金额,0)*V_药品收发记录.入出系数,
            差价=NVL(差价,0)+NVL(V_药品收发记录.差价,0)*V_药品收发记录.入出系数
        WHERE 日期 = TRUNC (SYSDATE)
            AND 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.材料ID
            AND 类别ID = V_药品收发记录.入出类别ID
            AND 单据 = 22;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品收发汇总
                (日期,库房ID,药品ID,类别ID,单据,数量,金额,差价)
            VALUES (TRUNC (SYSDATE),V_药品收发记录.库房ID,V_药品收发记录.材料ID,V_药品收发记录.入出类别ID,22,
            V_药品收发记录.实际数量*V_药品收发记录.入出系数,V_药品收发记录.零售金额*V_药品收发记录.入出系数,
            V_药品收发记录.差价*V_药品收发记录.入出系数);
        END IF;
    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料盘点_strike;
/
-------------------------------------------------
--不管材料的批次或时价属性，直接原样保存
-------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_材料申领_Insert (
	NO_IN		IN 药品收发记录.NO%TYPE,
	序号_IN		IN 药品收发记录.序号%TYPE,
	库房ID_IN	IN 药品收发记录.库房ID%TYPE,
	对方部门ID_IN   IN 药品收发记录.对方部门ID%TYPE,
	材料ID_IN	IN 药品收发记录.药品ID%TYPE,
	批次_IN		IN 药品收发记录.批次%TYPE,
	填写数量_IN	IN 药品收发记录.填写数量%TYPE,
	实际数量_IN	IN 药品收发记录.实际数量%TYPE,
	成本价_IN	IN 药品收发记录.成本价%TYPE,
	成本金额_IN	IN 药品收发记录.成本金额%TYPE,
	零售价_IN	IN 药品收发记录.零售价%TYPE,
	零售金额_IN	IN 药品收发记录.零售金额%TYPE,
	差价_IN		IN 药品收发记录.差价%TYPE,
	填制人_IN	IN 药品收发记录.填制人%TYPE,
	产地_IN		IN 药品收发记录.产地%TYPE := NULL,
	批号_IN		IN 药品收发记录.批号%TYPE := NULL,
	效期_IN		IN 药品收发记录.效期%TYPE := NULL,
	灭菌效期_IN	IN 药品收发记录.灭菌效期%TYPE := NULL,
	摘要_IN		IN 药品收发记录.摘要%TYPE := NULL,
	填制日期_IN	IN 药品收发记录.填制日期%TYPE := NULL
)
IS
	mErrItem	EXCEPTION ;
	mErrMsg		varchar2(100);

	V_lngID		药品收发记录.ID%TYPE;--收发ID
	V_入的类别ID    药品收发记录.入出类别ID%TYPE;--入出类别ID
	V_出的类别ID    药品收发记录.入出类别ID%TYPE;--入出类别ID
	v_下库存	系统参数表.参数值%type;
	v_明确批次	系统参数表.参数值%type;

	V_编码		收费项目目录.编码%TYPE;
	V_可用数量	药品库存.可用数量%TYPE;
	V_批次		药品收发记录.批次%TYPE := NULL;--主要针对入库中实行分批核算的材料
	V_是否分批	INTEGER;--判断入库是否分批核算   1:分批；0：不分批
	V_库房分批	INTEGER;--判断入库是否分批核算   1:分批；0：不分批
	V_在用分批	INTEGER;--判断入库是否分批核算   1:分批；0：不分批
	intRecords	NUMBER ;
BEGIN
	BEGIN
		SELECT nvl(参数值,'0') INTO v_下库存 FROM 系统参数表 WHERE 参数号=95;
	EXCEPTION 
		WHEN OTHERS THEN v_下库存:='-99';
	END;

	IF v_下库存='-99' THEN 
		mErrMsg:='[ZLSOFT]在系统参数中无"卫材填单下可用库存"参数,请与系统管员联系![ZLSOFT]';
		RAISE mErrItem;
	END IF ;

	--只有在明确批次的情况下才能下可用库存
	BEGIN
		SELECT nvl(参数值,'0') INTO v_明确批次 FROM 系统参数表 WHERE 参数号=83;
	EXCEPTION 
		WHEN OTHERS THEN v_明确批次:='-99';
	END;

	
	IF v_明确批次='-99' THEN 
		mErrMsg:='[ZLSOFT]在系统参数中无"按批次申领卫生材料"参数,请与系统管员联系![ZLSOFT]';
		RAISE mErrItem;
	END IF ;

	--首先找出入和出的类别ID
	SELECT B.ID INTO V_入的类别ID
	FROM 药品单据性质 A,药品入出类别 B
	WHERE A.类别ID = B.ID AND A.单据 = 34 AND B.系数 = 1 AND ROWNUM < 2;

	SELECT B.ID INTO V_出的类别ID
	FROM 药品单据性质 A,药品入出类别 B
	WHERE A.类别ID = B.ID AND A.单据 = 34 AND B.系数 = -1 AND ROWNUM < 2;


	SELECT 药品收发记录_ID.Nextval INTO V_lngID FROM Dual;

	--插入类别为出的那一笔
	Insert INTO 药品收发记录
		(ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,灭菌效期,
		填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,发药方式)
	VALUES (药品收发记录_ID.Nextval,1,19,NO_IN,序号_IN,库房ID_IN,对方部门ID_IN,
		V_出的类别ID,-1,材料ID_IN,批次_IN,产地_IN,批号_IN,效期_IN,灭菌效期_IN,
		填写数量_IN,实际数量_IN,成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,
		差价_IN,摘要_IN,填制人_IN,填制日期_IN,1);

	
	IF to_number(v_下库存,'99999') =1 AND to_number(v_明确批次,'99999')=1 THEN 

		--需要下可用库存

		--判断是否有可用库存
		IF 批次_IN > 0 THEN
			BEGIN
			    SELECT 可用数量    INTO V_可用数量
			    FROM 药品库存
			    WHERE 药品ID = 材料ID_IN
				AND NVL (批次,0) = 批次_IN
				AND 库房ID = 库房ID_IN
				AND 性质 = 1
				AND ROWNUM = 1;
			EXCEPTION
			    WHEN OTHERS THEN
				V_可用数量 := 0;
			END;

			IF V_可用数量 - 实际数量_IN < 0 THEN
			    SELECT 编码 INTO V_编码    FROM 收费项目目录 WHERE ID = 材料ID_IN;
			    mErrMsg:='[ZLSOFT]编码为' || V_编码 || ',批号为' || 批号_IN ||'的分批核算材料' || CHR (10) || CHR (13) || '可用库存数量不够！[ZLSOFT]';
			    RAISE mErrItem;
			END IF;
		END IF;

		UPDATE 药品库存
		SET 可用数量 = NVL(可用数量, 0) - 实际数量_IN
		WHERE 库房ID = 库房ID_IN AND 药品ID = 材料ID_IN 
			AND NVL(批次, 0) = NVL(批次_IN, 0) AND 性质 = 1;

		IF SQL%ROWCOUNT = 0 THEN
			INSERT INTO 药品库存(库房ID, 药品ID, 批次, 性质, 可用数量)
			VALUES(库房ID_IN, 材料ID_IN, NVL(批次_IN, 0), 1, -实际数量_IN);
		END IF;

		--同时更新库存数
		DELETE
		FROM 药品库存
		WHERE 库房ID = 库房ID_IN
			AND 药品ID = 材料ID_IN
			AND nvl(可用数量,0) = 0
			AND nvl(实际数量,0) = 0
			AND nvl(实际金额,0) = 0
			AND nvl(实际差价,0) = 0;

		--下面是判断入库的材料是否是分批核算材料
		SELECT NVL (库房分批,0),nvl(在用分批,0) INTO V_库房分批,v_在用分批
		FROM 材料特性
		WHERE 材料ID = 材料ID_IN;

		V_是否分批 := 0;
		IF v_在用分批=0 then
			IF V_库房分批 = 1 THEN
			    BEGIN
				SELECT DISTINCT 0 INTO V_是否分批
				FROM 部门性质说明
				WHERE ( (工作性质 LIKE '发料部门')    OR (工作性质 LIKE '制剂室')) AND 部门ID = 对方部门ID_IN;
			    EXCEPTION
				WHEN OTHERS THEN V_是否分批 := 1;
			    END;
			END IF;
		ELSE 
			V_是否分批 := 1;
		END if;


		IF V_是否分批 = 1 AND NVL (批次_IN,0) = 0 THEN--入库分批且出库不分批
			V_批次 := V_lngID;
		ELSIF    V_是否分批 = 0 THEN--入库不分批
			V_批次 := 0;
		ELSIF NVL (批次_IN,0) <> 0 THEN--入库分批且出库也分批
			V_批次 := 批次_IN;
		END IF;


		--插入类别为入的那一笔
		Insert INTO 药品收发记录
		(ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,灭菌效期,
		填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,发药方式)
		VALUES (V_lngID,1,19,NO_IN,序号_IN + 1,对方部门ID_IN,库房ID_IN,V_入的类别ID,
		1,材料ID_IN,V_批次,产地_IN,批号_IN,效期_IN,灭菌效期_IN,填写数量_IN,实际数量_IN,
		成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,差价_IN,摘要_IN,
		填制人_IN,填制日期_IN,1);

		--检查是否存在相同材料相同批次的数据，如果存在不允许保存
		SELECT COUNT(*) INTO intRecords 
		FROM 药品收发记录
		WHERE 单据=19 AND NO=NO_IN AND 入出系数=-1 AND 药品ID+0=材料ID_IN AND Nvl(批次,0)=NVL(批次_IN,0);

		IF intRecords>1 THEN
			SELECT 编码 INTO V_编码 FROM 收费项目目录 WHERE ID = 材料ID_IN;
			mErrMsg:='[ZLSOFT]编码为'||V_编码||'的药品，存在多条重复的记录，请合并为一条记录！[ZLSOFT]';
			RAISE mErrItem;
		END IF;
	ELSE 
		--插入类别为入的那一笔
		Insert INTO 药品收发记录
		(ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,灭菌效期,
		填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,发药方式)
		VALUES (V_lngID,1,19,NO_IN,序号_IN + 1,对方部门ID_IN,库房ID_IN,V_入的类别ID,
		1,材料ID_IN,批次_IN,产地_IN,批号_IN,效期_IN,灭菌效期_IN,填写数量_IN,实际数量_IN,
		成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,差价_IN,摘要_IN,
		填制人_IN,填制日期_IN,1);

	END IF ;
        
EXCEPTION
	WHEN mErrItem THEN 
		raise_application_error(-20101,mErrMsg);
	WHEN OTHERS THEN
	        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_材料申领_Insert;
/

CREATE OR REPLACE PROCEDURE zl_草药药品_INSERT(
    分类ID_IN IN 诊疗项目目录.分类ID%TYPE := NULL,
    ID_IN IN 诊疗项目目录.ID%TYPE,
    编码_IN IN 诊疗项目目录.编码%TYPE := NULL,
    标识码_IN IN 药品规格.标识码%TYPE := NULL,
    名称_IN IN 诊疗项目目录.名称%TYPE := NULL,
    拼音_IN IN 诊疗项目别名.简码%TYPE := NULL,
    五笔_IN IN 诊疗项目别名.简码%TYPE := NULL,
    产地_IN IN 收费项目目录.产地%TYPE := NULL,
    单位_IN IN 诊疗项目目录.计算单位%TYPE := NULL,
    规格_IN IN 收费项目目录.规格%Type := NULL,           --新增的
    售价单位_IN IN 收费项目目录.计算单位%TYPE := NULL,   --新增的
  	剂量系数_IN IN 药品规格.剂量系数%TYPE := NULL,       --新增的
  	门诊单位_IN IN 药品规格.门诊单位%TYPE := NULL,       --新增的
  	门诊包装_IN IN 药品规格.门诊包装%TYPE := NULL,       --新增的
  	住院单位_IN IN 药品规格.住院单位%TYPE := NULL,       --新增的
  	住院包装_IN IN 药品规格.住院包装%TYPE := NULL,       --新增的
    药库单位_IN IN 药品规格.药库单位%TYPE := NULL,       
    药库包装_IN IN 药品规格.药库包装%TYPE := NULL,
	  申领单位_IN IN 药品规格.申领单位%TYPE := 1,
	  申领阀值_IN IN 药品规格.申领阀值%TYPE := NULL,
    毒理分类_IN IN 药品特性.毒理分类%TYPE := NULL,
    价值分类_IN IN 药品特性.价值分类%TYPE := NULL,
    货源情况_IN IN 药品特性.货源情况%TYPE := NULL,
    用药梯次_IN IN 药品特性.用药梯次%TYPE := NULL,
    药品类型_IN IN 药品特性.药品类型%TYPE := NULL,
    处方职务_IN IN 药品特性.处方职务%TYPE := '00',
    处方限量_IN IN 药品特性.处方限量%TYPE := NULL,
    单独应用_IN IN 诊疗项目目录.单独应用%TYPE := NULL,
    是否原料_IN IN 药品特性.是否原料%TYPE := 0,
    是否变价_IN IN 收费项目目录.是否变价%TYPE := NULL,
    指导批发价_IN IN 药品规格.指导批发价%TYPE := NULL,
    扣率_IN IN 药品规格.扣率%TYPE := 95,
    指导零售价_IN IN 药品规格.指导零售价%TYPE := NULL,
    指导差价率_IN IN 药品规格.指导差价率%TYPE := NULL,
	  管理费比例_IN IN 药品规格.管理费比例%TYPE := NULL,
    药价级别_IN IN 药品规格.药价级别%TYPE := NULL,
    费用类型_IN IN 收费项目目录.费用类型%TYPE := NULL,
    服务对象_IN IN 诊疗项目目录.服务对象%TYPE := NULL,
    屏蔽费别_IN IN 收费项目目录.屏蔽费别%TYPE := 0,
    药库分批_IN IN 药品规格.药库分批%TYPE := NULL,
    药房分批_IN IN 药品规格.药房分批%TYPE := NULL,
    参考目录Id_IN In 诊疗项目目录.参考目录Id%Type:=Null,
    其他别名_IN IN VARCHAR2 :=NULL,      --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织
	  成本价_IN in 药品规格.成本价%TYPE := 0,
    当前售价_IN IN 收费价目.现价%TYPE := 0,
    收入ID_IN IN 收费价目.收入项目ID%TYPE := NULL,
    合同单位id_IN IN 药品规格.合同单位id%TYPE := Null,
    说明_IN In 收费项目目录.说明%Type:=Null,
    可否分零_IN In 药品规格.可否分零%Type := NULL
) IS
    v_药品ID NUMBER;             --收费项目ID，根据序列提取产生
    v_Records VARCHAR2(4000);    --临时记录别名数据的字符串
    v_CurrRec VARCHAR2(1000);    --包含在别名记录中的一条别名
    v_Fields  VARCHAR2(1000);    --临时记录一条别名的字符串
    v_名称 诊疗项目目录.名称%TYPE;
    v_拼音 诊疗项目别名.简码%TYPE;
    v_五笔 诊疗项目别名.简码%TYPE;
    v_诊疗项目ID number;
	
	--查找库房id
	Cursor c_StorageID Is 
		Select DISTINCT 部门id 
		From 部门性质说明
		Where 工作性质 Like '中药%' Or 工作性质 = '制剂室';
	r_StorageID c_StorageID%RowType;
BEGIN
    INSERT INTO 诊疗项目目录(类别,分类ID,ID,编码,名称,计算单位,
        计算方式,执行频率,适用性别,服务对象,单独应用,组合项目,执行安排,计价性质,建档时间,撤档时间,参考目录Id)
    VALUES ('7',分类ID_IN,ID_IN,编码_IN,名称_IN,单位_IN,
        1,0,0,服务对象_IN,单独应用_IN,0,0,0,sysdate,to_date('3000-01-01','YYYY-MM-DD'),参考目录Id_IN);
    INSERT INTO 药品特性(药名ID,药品剂型,毒理分类,价值分类,货源情况,用药梯次,
        药品类型,处方职务,处方限量,急救药否,是否新药,是否原料,是否皮试)
    VALUES (ID_IN,'方剂',毒理分类_IN,价值分类_IN,货源情况_IN,用药梯次_IN,
        药品类型_IN,处方职务_IN,处方限量_IN,0,0,是否原料_IN,0);

    select 收费项目目录_ID.NEXTVAL into v_药品ID from dual;
    insert into 收费项目目录(类别,ID,编码,名称,规格,产地,计算单位,
           费用类型,服务对象,屏蔽费别,是否变价,建档时间,撤档时间,说明)
    values ('7',v_药品ID,编码_IN,名称_IN,规格_IN,产地_IN,售价单位_IN,
           费用类型_IN,服务对象_IN,屏蔽费别_IN,是否变价_IN,sysdate,to_date('3000-01-01','YYYY-MM-DD'),说明_IN);
    Insert INTO 药品规格(药名ID,药品ID,标识码,药品来源,剂量系数,
           门诊单位,门诊包装,住院单位,住院包装,药库单位,药库包装,申领单位,申领阀值,
           指导批发价,扣率,指导零售价,指导差价率,管理费比例,药价级别,成本价,
           可否分零,药库分批,药房分批,最大效期,合同单位id)
    VALUES (ID_IN,v_药品ID,标识码_IN,'国产',剂量系数_IN,
           门诊单位_IN,门诊包装_IN,住院单位_IN,住院包装_IN,
           药库单位_IN,药库包装_IN,申领单位_IN,申领阀值_IN,
           指导批发价_IN,扣率_IN,指导零售价_IN,指导差价率_IN,管理费比例_IN,药价级别_IN,成本价_IN,
           可否分零_IN,药库分批_IN,药房分批_IN,0,合同单位id_IN);

    IF 拼音_IN IS NOT NULL THEN
        INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,名称_IN,1,拼音_IN,1);
        INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (v_药品ID,名称_IN,1,拼音_IN,1);
    END IF;
    IF 五笔_IN IS NOT NULL THEN
        INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,名称_IN,1,五笔_IN,2);
        INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (v_药品ID,名称_IN,1,五笔_IN,2);
    END IF;
    IF 其他别名_IN IS NULL THEN
        v_Records := null;
    ELSE
        v_Records := 其他别名_IN||'|';
    END IF;
    WHILE v_Records IS NOT NULL LOOP
        v_CurrRec:=substr(v_Records,1,instr(v_Records,'|')-1);
        v_Fields:=v_CurrRec;
        v_名称:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_拼音:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_五笔:=v_Fields;
        IF V_拼音 IS NOT NULL THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,v_名称,9,v_拼音,1);
            INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (v_药品ID,v_名称,9,v_拼音,1);
        END IF;
        IF v_五笔 IS NOT NULL THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,v_名称,9,v_五笔,2);
            INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (v_药品ID,v_名称,9,v_五笔,2);
        END IF;
        v_Records:=replace('|'||v_Records,'|'||v_CurrRec||'|');
    END LOOP;
    --定价信息
    if 收入ID_IN is not null then
       insert into 收费价目(ID,原价ID,收费细目ID,原价,现价,收入项目ID,变动原因,调价说明,调价人,执行日期,终止日期)
       values (收费价目_ID.Nextval,null,v_药品ID,0,当前售价_IN,收入ID_IN,1,'新增定价',user,sysdate,to_date('3000-01-01','YYYY-MM-DD'));
    end if;

    --添加缺省的对应输出单据
    INSERT INTO 诊疗单据应用(病历文件id,应用场合,诊疗项目id)
    SELECT A.病历文件id,1,ID_IN
    FROM 诊疗单据应用 A,诊疗项目目录 I
    Where A.诊疗项目id=I.Id And I.类别='7' And 应用场合=1 And Rownum<2;
    INSERT INTO 诊疗单据应用(病历文件id,应用场合,诊疗项目id)
    SELECT A.病历文件id,2,ID_IN
    FROM 诊疗单据应用 A,诊疗项目目录 I
    Where A.诊疗项目id=I.Id And I.类别='7' And 应用场合=2 And Rownum<2;

	--插入盘点属性
	For r_StorageID In c_StorageID Loop
		Insert Into 药品储备限额(库房ID, 药品ID, 上限, 下限, 盘点属性, 库房货位) Values(r_StorageID.部门id, v_药品ID, 0, 0, '1111', null);
	End Loop;
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_草药药品_INSERT;
/

CREATE OR REPLACE PROCEDURE zl_草药药品_Update(
    分类ID_IN IN 诊疗项目目录.分类ID%TYPE := NULL,
    ID_IN IN 诊疗项目目录.ID%TYPE,
    编码_IN IN 诊疗项目目录.编码%TYPE := NULL,
    标识码_IN IN 药品规格.标识码%TYPE := NULL,
    名称_IN IN 诊疗项目目录.名称%TYPE := NULL,
    拼音_IN IN 诊疗项目别名.简码%TYPE := NULL,
    五笔_IN IN 诊疗项目别名.简码%TYPE := NULL,
    产地_IN IN 收费项目目录.产地%TYPE := NULL,
    单位_IN IN 诊疗项目目录.计算单位%TYPE := NULL,
    规格_IN IN 收费项目目录.规格%Type := NULL,           --新增的
    售价单位_IN IN 收费项目目录.计算单位%TYPE := NULL,   --新增的
  	剂量系数_IN IN 药品规格.剂量系数%TYPE := NULL,       --新增的
  	门诊单位_IN IN 药品规格.门诊单位%TYPE := NULL,       --新增的
  	门诊包装_IN IN 药品规格.门诊包装%TYPE := NULL,       --新增的
  	住院单位_IN IN 药品规格.住院单位%TYPE := NULL,       --新增的
  	住院包装_IN IN 药品规格.住院包装%TYPE := NULL,       --新增的
    药库单位_IN IN 药品规格.药库单位%TYPE := NULL,
    药库包装_IN IN 药品规格.药库包装%TYPE := NULL,
	  申领单位_IN IN 药品规格.申领单位%TYPE := 1,
	  申领阀值_IN IN 药品规格.申领阀值%TYPE := NULL,
    毒理分类_IN IN 药品特性.毒理分类%TYPE := NULL,
    价值分类_IN IN 药品特性.价值分类%TYPE := NULL,
    货源情况_IN IN 药品特性.货源情况%TYPE := NULL,
    用药梯次_IN IN 药品特性.用药梯次%TYPE := NULL,
    药品类型_IN IN 药品特性.药品类型%TYPE := NULL,
    处方职务_IN IN 药品特性.处方职务%TYPE := '00',
    处方限量_IN IN 药品特性.处方限量%TYPE := NULL,
    单独应用_IN IN 诊疗项目目录.单独应用%TYPE := NULL,
    是否原料_IN IN 药品特性.是否原料%TYPE := 0,
    是否变价_IN IN 收费项目目录.是否变价%TYPE := NULL,
    指导批发价_IN IN 药品规格.指导批发价%TYPE := NULL,
    扣率_IN IN 药品规格.扣率%TYPE := 95,
    指导零售价_IN IN 药品规格.指导零售价%TYPE := NULL,
    指导差价率_IN IN 药品规格.指导差价率%TYPE := NULL,
	  管理费比例_IN IN 药品规格.管理费比例%TYPE := NULL,
    药价级别_IN IN 药品规格.药价级别%TYPE := NULL,
    费用类型_IN IN 收费项目目录.费用类型%TYPE := NULL,
    服务对象_IN IN 诊疗项目目录.服务对象%TYPE := NULL,
    屏蔽费别_IN IN 收费项目目录.屏蔽费别%TYPE := 0,
    药库分批_IN IN 药品规格.药库分批%TYPE := NULL,
    药房分批_IN IN 药品规格.药房分批%TYPE := NULL,
    参考目录Id_IN In 诊疗项目目录.参考目录Id%Type:=Null,
    其他别名_IN IN VARCHAR2,      --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织
	  成本价_IN in 药品规格.成本价%TYPE := 0,
    当前售价_IN IN 收费价目.现价%TYPE := 0,
    收入ID_IN IN 收费价目.收入项目ID%TYPE := NULL,
    合同单位id_IN IN 药品规格.合同单位id%TYPE := Null,
    说明_IN In 收费项目目录.说明%Type:=Null,
    可否分零_IN In 药品规格.可否分零%Type := NULL
) IS
    v_药品ID NUMBER;             --收费项目ID，根据序列提取产生
    v_Records VARCHAR2(4000);    --临时记录别名数据的字符串
    v_CurrRec VARCHAR2(1000);    --包含在别名记录中的一条别名
    v_Fields  VARCHAR2(1000);    --临时记录一条别名的字符串
    v_名称 诊疗项目目录.名称%TYPE;
    v_拼音 诊疗项目别名.简码%TYPE;
    v_五笔 诊疗项目别名.简码%TYPE;
    v_发生 Number(2);
    Err_NotFind  EXCEPTION;
BEGIN
    UPDATE 诊疗项目目录
    SET 分类ID=分类ID_IN,编码=编码_IN,名称=名称_IN,计算单位=单位_IN,服务对象=服务对象_IN,单独应用=单独应用_IN,参考目录Id=参考目录Id_IN
    WHERE ID=ID_IN;
    IF SQL%ROWCOUNT=0 THEN
        RAISE Err_NotFind;
    END IF;
    UPDATE 药品特性
    SET 毒理分类=毒理分类_IN,价值分类=价值分类_IN,货源情况=货源情况_IN,用药梯次=用药梯次_IN,
        药品类型=药品类型_IN,处方职务=处方职务_IN,处方限量=处方限量_IN,
        是否原料=是否原料_IN
    WHERE 药名ID=ID_IN;

    select 药品id into v_药品ID from 药品规格 where 药名id=ID_IN and rownum<2;
    update 收费项目目录
    set 编码=编码_IN,名称=名称_IN,产地=产地_IN,计算单位=售价单位_IN,规格=规格_IN,
        费用类型=费用类型_IN,服务对象=服务对象_IN,屏蔽费别=屏蔽费别_IN,是否变价=是否变价_IN,说明=说明_IN
    where ID=v_药品ID;

    update 药品规格
    set 标识码=标识码_IN,剂量系数=剂量系数_IN,
        门诊单位=门诊单位_IN,门诊包装=门诊包装_IN,
        住院单位=住院单位_IN,住院包装=住院包装_IN,
        药库单位=药库单位_IN,药库包装=药库包装_IN,申领单位=申领单位_IN,申领阀值=申领阀值_IN,
        指导批发价=指导批发价_IN,扣率=扣率_IN,指导零售价=指导零售价_IN,指导差价率=指导差价率_IN,
		管理费比例=管理费比例_IN,药价级别=药价级别_IN,药库分批=药库分批_IN,药房分批=药房分批_IN,合同单位id=合同单位id_IN,
    可否分零=可否分零_IN
    where 药品ID=v_药品ID;

    IF 拼音_IN IS NULL THEN
        DELETE FROM 诊疗项目别名 WHERE 诊疗项目ID=ID_IN AND 性质=1 AND 码类=1;
        DELETE FROM 收费项目别名 WHERE 收费细目id=v_药品ID AND 性质=1 AND 码类=1;
    ELSE
        UPDATE 诊疗项目别名 SET 名称=名称_IN, 简码=拼音_IN WHERE 诊疗项目ID=ID_IN AND 性质=1 AND 码类=1;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,名称_IN,1,拼音_IN,1);
        END IF;
        update 收费项目别名 SET 名称=名称_IN, 简码=拼音_IN WHERE 收费细目ID=v_药品ID AND 性质=1 AND 码类=1;
        if sql%rowcount=0 then
           insert into 收费项目别名(收费细目ID,名称,性质,简码,码类) values(v_药品ID,名称_IN,1,拼音_IN,1);
        end if;
    END IF;
    IF 五笔_IN IS NULL THEN
        DELETE FROM 诊疗项目别名 WHERE 诊疗项目ID=ID_IN AND 性质=1 AND 码类=2;
        DELETE FROM 收费项目别名 WHERE 收费细目id=v_药品ID AND 性质=1 AND 码类=2;
    ELSE
        UPDATE 诊疗项目别名 SET 名称=名称_IN, 简码=五笔_IN WHERE 诊疗项目ID=ID_IN AND 性质=1 AND 码类=2;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,名称_IN,1,五笔_IN,2);
        END IF;
        update 收费项目别名 SET 名称=名称_IN, 简码=五笔_IN WHERE 收费细目ID=v_药品ID AND 性质=1 AND 码类=2;
        if sql%rowcount=0 then
           insert into 收费项目别名(收费细目ID,名称,性质,简码,码类) values(v_药品ID,名称_IN,1,五笔_IN,2);
        end if;
    END IF;

    DELETE FROM 诊疗项目别名 WHERE 诊疗项目ID=ID_IN AND 性质=9;
    DELETE FROM 收费项目别名 WHERE 收费细目id=v_药品id AND 性质=9;
    IF 其他别名_IN IS NULL THEN
        v_Records := null;
    ELSE
        v_Records := 其他别名_IN||'|';
    END IF;
    WHILE v_Records IS NOT NULL LOOP
        v_CurrRec:=substr(v_Records,1,instr(v_Records,'|')-1);
        v_Fields:=v_CurrRec;
        v_名称:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_拼音:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_五笔:=v_Fields;
        IF V_拼音 IS NOT NULL THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,v_名称,9,v_拼音,1);
            INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (v_药品ID,v_名称,9,v_拼音,1);
        END IF;
        IF v_五笔 IS NOT NULL THEN
            INSERT INTO 诊疗项目别名(诊疗项目ID,名称,性质,简码,码类) VALUES (ID_IN,v_名称,9,v_五笔,2);
            INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (v_药品ID,v_名称,9,v_五笔,2);
        END IF;
        v_Records:=replace('|'||v_Records,'|'||v_CurrRec||'|');
    END LOOP;
    --定价信息：如果已经有发生，则不允许直接更改这些信息
    Select nvl(Count(*),0) Into v_发生 From 药品收发记录 Where 药品id=v_药品ID And rownum<2;
    If v_发生=0 Then
        update 收费项目目录 set 是否变价=是否变价_IN where ID=v_药品ID;
        update 药品规格 set 成本价=成本价_IN where 药品ID=v_药品ID;
        if 收入ID_IN is not null Then
           Update 收费价目
           Set 现价=当前售价_IN,收入项目ID=收入ID_IN,变动原因=1,调价说明='修改定价',调价人=User
           Where 收费细目ID=v_药品ID
                 And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因=1;
           If Sql%Rowcount=0 Then
              insert into 收费价目(ID,原价ID,收费细目ID,原价,现价,收入项目ID,变动原因,调价说明,调价人,执行日期,终止日期)
              values (收费价目_ID.Nextval,null,v_药品ID,0,当前售价_IN,收入ID_IN,1,'新增定价',user,sysdate,to_date('3000-01-01','YYYY-MM-DD'));
           End If;
        end if;
    End If;
EXCEPTION
    WHEN Err_NotFind THEN
        Raise_application_error (-20101, '[ZLSOFT]该药品不存在，可能已被其他用户删除！[ZLSOFT]');
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_草药药品_Update;
/

--by 陈福容 {
--计算一段内容占用行数
CREATE OR REPLACE Function zl_GetTextRows(文本_IN Varchar2,个数_IN Number) Return Number
As
	v_Rows		Number(18);
	v_Len		Number(18);
	v_Pos		Number(18);
	v_Tmp		Varchar2(4000);
	v_TmpLine	Varchar2(4000);
Begin

	v_Rows:=0;

	v_Tmp:=文本_IN;
	v_Pos:=Instrb(v_Tmp,chr(10));
	While v_Pos>0	loop
		v_TmpLine:=SubStrb(v_Tmp,1,v_Pos-1);

		v_Len:=Lengthb(v_TmpLine);

		v_Rows:=v_Rows+Ceil(v_Len/个数_IN);

		v_Tmp:=SubStrb(v_Tmp,v_Pos+1);
		v_Pos:=Instrb(v_Tmp,chr(10));

	END loop;
	v_Len:=Lengthb(v_Tmp);
	v_Rows:=v_Rows+Ceil(v_Len/个数_IN);

	If v_Rows<1 Then
		v_Rows:=1;
	End If;

	Return(v_Rows);
End;
/
----------------------------------------------------------------------------
---  INSERT   for   体检诊断建议
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检诊断建议_INSERT(
	序号_IN IN 体检诊断建议.序号%TYPE,
	编码_IN IN 体检诊断建议.编码%TYPE,
	名称_IN IN 体检诊断建议.名称%TYPE,
	简码_IN IN 体检诊断建议.简码%TYPE,
	是否疾病_IN IN 体检诊断建议.是否疾病%TYPE,
	诊断建议_IN IN 体检诊断建议.诊断建议%TYPE,
	上级序号_IN IN 体检诊断建议.上级序号%TYPE:=NULL,
	末级_IN IN 体检类型.末级%TYPE:=1,
	同级调整_IN  NUMBER:=0
)
IS
	v_Extend number(18);
	v_Parent varchar2(30);
BEGIN		
	IF 末级_IN=0 THEN
		IF 同级调整_IN=1 THEN
			    --调整同级编码的长度
			IF NVL(上级序号_IN,0)<>0 THEN
			    SELECT 编码 INTO v_Parent FROM 体检诊断建议 WHERE 序号=上级序号_IN;
			ELSE
			    v_Parent:=NULL;
			END IF;

			BEGIN
			    SELECT length(rtrim(编码_IN))-length(rtrim(编码)) INTO v_Extend
			    FROM 体检诊断建议
			    WHERE 末级=0 AND (上级序号=上级序号_IN OR 上级序号 IS NULL AND NVL(上级序号_IN,0)=0) AND Rownum=1;
			EXCEPTION
			    WHEN OTHERS THEN v_Extend:=0;
			END;

			IF v_Extend>0 THEN
			    --扩充处理
			    IF v_Parent IS null THEN
				UPDATE 体检诊断建议 SET 编码=lpad('0',v_Extend,'0')||编码 WHERE 序号<>序号_IN AND 末级=0;
			    ELSE
				UPDATE 体检诊断建议 SET 编码=v_Parent||lpad('0',v_Extend,'0')||substr(编码,length(v_Parent)+1) WHERE 编码 LIKE v_Parent||'_%' AND 末级=0;
			    END IF;
			END IF;

			IF v_Extend<0 THEN
			    --压缩处理
			    IF v_Parent IS null THEN
				UPDATE 体检诊断建议 SET 编码=substr(编码,1+abs(v_Extend)) WHERE 序号<>序号_IN AND 末级=0;
			    ELSE
				UPDATE 体检诊断建议 SET 编码=v_Parent||substr(编码,length(v_Parent)+1+abs(v_Extend)) WHERE 编码 LIKE v_Parent||'_%' AND 末级=0;
			    END IF;
			END IF;

		END IF;
	END IF;
	Insert Into 体检诊断建议(序号,上级序号,末级,编码,名称,简码,是否疾病,诊断建议) VALUES(序号_IN,DECODE(上级序号_IN,0,NULL,上级序号_IN),末级_IN,编码_IN,名称_IN,简码_IN,是否疾病_IN,诊断建议_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检诊断建议_INSERT;
/
----------------------------------------------------------------------------
---  UPDATE   for   体检诊断建议
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检诊断建议_UPDATE(
	序号_IN IN 体检诊断建议.序号%TYPE,
	编码_IN IN 体检诊断建议.编码%TYPE,
	名称_IN IN 体检诊断建议.名称%TYPE,
	简码_IN IN 体检诊断建议.简码%TYPE,
	是否疾病_IN IN 体检诊断建议.是否疾病%TYPE,
	诊断建议_IN IN 体检诊断建议.诊断建议%TYPE,
	上级序号_IN IN 体检诊断建议.上级序号%TYPE:=NULL,
	同级调整_IN  NUMBER:=0
)
IS
	v_OldCode  VARCHAR2(30);  --原来的编码
	v_Parent  VARCHAR2(30);  --上级编码
	v_Extend  NUMBER(18);    --扩充长度(为负表示压缩)
	Err_NotFind  EXCEPTION;
BEGIN
	
	SELECT rtrim(编码) INTO v_OldCode FROM 体检诊断建议 WHERE 序号=序号_IN;
	IF v_OldCode is null THEN
		RAISE Err_NotFind;
	END IF;

	--修改项目本身
	Update 体检诊断建议
		Set 编码=编码_IN,
		    名称=名称_IN,
		    简码=简码_IN,
		    是否疾病=是否疾病_IN,
		    诊断建议=诊断建议_IN,
		    上级序号=DECODE(上级序号_IN,0,NULL,上级序号_IN)
	WHERE 序号=序号_IN;    

	--修改本系各级下属编码

	UPDATE 体检诊断建议 SET 编码=编码_IN||substr(编码,length(v_OldCode)+1) WHERE 编码<>编码_IN And 编码 LIKE v_OldCode||'_%' And 末级=0;

	--调整同级编码的长度
	IF 同级调整_IN=1 THEN
		IF NVL(上级序号_IN,0)<>0 THEN
		    SELECT 编码 INTO v_Parent FROM 体检诊断建议 WHERE 序号=上级序号_IN;
		ELSE
		    v_Parent:=NULL;
		END IF;

		BEGIN
		    SELECT length(rtrim(编码_IN))-length(rtrim(编码)) INTO v_Extend FROM 体检诊断建议 WHERE 末级=0 AND (上级序号=上级序号_IN OR 上级序号 IS NULL AND nvl(上级序号_IN,0)=0) AND 序号<>序号_IN AND Rownum=1;
		EXCEPTION
		    WHEN OTHERS THEN v_Extend:=0;
		END;

		IF v_Extend>0 THEN
		    --扩充处理
		    IF v_Parent IS null THEN
			UPDATE 体检诊断建议 SET 编码=lpad('0',v_Extend,'0')||编码  WHERE 末级=0 and 序号 not in (select 序号 from 体检诊断建议 WHERE 末级=0 start with 序号=序号_IN connect by prior 序号=上级序号);
		    ELSE
			UPDATE 体检诊断建议	SET 编码=v_Parent||lpad('0',v_Extend,'0')||substr(编码,length(v_Parent)+1) WHERE 末级=0 AND 编码 LIKE v_Parent||'_%' and 序号 not in (select 序号 from 体检诊断建议 where 末级=0 start with 序号=序号_IN connect by prior 序号=上级序号);
		    END IF;
		END IF;

		IF v_Extend<0 THEN
		    --压缩处理
		    IF v_Parent IS null THEN
			UPDATE 体检诊断建议 SET 编码=substr(编码,1+abs(v_Extend)) WHERE 序号<>序号_IN AND 末级=0;
		    ELSE
			UPDATE 体检诊断建议 SET 编码=v_Parent||substr(编码,length(v_Parent)+1+abs(v_Extend)) WHERE 编码 LIKE v_Parent||'_%' AND 序号<>序号_IN AND 末级=0;
		    END IF;
		END IF;
	END IF;
EXCEPTION
	WHEN Err_NotFind THEN Raise_application_error (-20101, '[ZLSOFT]该项目不存在，可能已被其他用户删除！[ZLSOFT]');
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检诊断建议_UPDATE;
/

CREATE OR REPLACE PROCEDURE ZL_体检人员结论_INSERT(
	病人id_IN IN 体检人员结论.病人id%TYPE,
	主页id_IN IN 体检人员结论.主页id%TYPE,
	病历id_IN IN 体检人员结论.病历id%TYPE,
	记录性质_IN IN 体检人员结论.记录性质%TYPE,
	记录序号_IN IN 体检人员结论.记录序号%TYPE,
	结论描述_IN IN 体检人员结论.结论描述%TYPE,
	参考建议_IN IN 体检人员结论.参考建议%TYPE,
	结论id_IN IN 体检人员结论.结论id%TYPE,
	是否疾病_IN IN 体检人员结论.是否疾病%TYPE:=0,
	诊断建议_IN IN 体检人员结论.诊断建议%TYPE:=Null

)
IS
BEGIN
	INSERT INTO 体检人员结论(病人id,主页id,病历id,记录性质,记录序号,结论描述,参考建议,结论id,是否疾病,诊断建议) 
	VALUES (病人id_IN,DECODE(主页id_IN,0,NULL,主页id_IN),病历id_IN,记录性质_IN,记录序号_IN,结论描述_IN,参考建议_IN,DECODE(结论id_IN,0,NULL,结论id_IN),是否疾病_IN,诊断建议_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检人员结论_INSERT;
/

--产生空的病人病历，书写人及书写日期为空
CREATE OR REPLACE PROCEDURE ZL_体检项目报告_EMPTY(
	病人id_IN	IN	体检人员档案.病人id%TYPE,
	清单ID_IN	IN	NUMBER
)
IS
	--医嘱
	CURSOR c_Advice(v_医嘱id In Number) IS	SELECT * FROM 病人医嘱记录 WHERE ID=v_医嘱id;
	r_Advice c_Advice%RowType;
	
	--病历文件
	CURSOR c_File(v_File number) IS	SELECT * FROM 病历文件目录 A where A.ID=v_File;
	r_File c_File%RowType;

	--组成元素
	CURSOR c_Element(v_File number) IS	
		SELECT 类型,编码,B.ID,文本转储,标题文本,标题显示,标题字体,标题位置,内容字体,内容位置,嵌入方式,B.部件
		FROM 病历文件组成 A,病历元素目录 B
		Where A.病历元素id=B.ID
		      And A.病历文件id=v_File
		Order By A.排列序号;
	
	--所见单元素
	CURSOR c_ElementPaper(v_Element number) IS
		Select * From 病历所见单 Where 元素ID=v_Element Order By 控件号;
	
	--检验数据
	CURSOR c_ElementVerfy(v_医嘱id Number) IS
		SELECT DISTINCT C.报告项目ID,
                               G.中文名,
                               zlGetReference(C.报告项目ID,A.标本部位,DECODE(E.性别,'男',1,'女',2,0),E.出生日期) AS 结果参考,
                               D.结果类型,
                               B.计算单位,
                               D.计算公式,C.排列序号
                        FROM 病人医嘱记录 A,
                             诊疗项目目录 B,
                             检验报告项目 C,
                             检验项目 D,
                             病人信息 E,
                             诊治所见项目 G
                        Where A.相关ID =v_医嘱id
                              AND E.病人ID=A.病人id
                              AND A.诊疗项目ID=B.ID
                              AND C.诊疗项目ID=B.ID
                              AND D.诊治项目ID=C.报告项目ID
                              AND G.ID=C.报告项目ID Order By C.排列序号;


	v_Index			Number(18);
	v_VerIndex		Number(18);
	v_报告id		Number(18);
	v_病历内容id		Number(18);
	v_病历文件id		Number(18);
	v_医嘱id		NUMBER(18);
BEGIN
	v_病历文件id:=0;
	v_医嘱id:=0;
	Begin 
		Select B.病历文件id,C.医嘱id Into v_病历文件id,v_医嘱id From 体检项目清单 A,诊疗单据应用 B,体检项目医嘱 C Where C.清单id=A.ID AND C.病人id=病人id_IN AND A.ID=清单ID_IN AND A.诊疗项目id=B.诊疗项目id AND B.应用场合=4;
	Exception
		When Others Then v_病历文件id:=0;				
	End;
	
	If v_病历文件id>0 Then 
		Open c_Advice(v_医嘱id);
		Fetch c_Advice Into r_Advice;

		Open c_File(v_病历文件id);
		Fetch c_File Into r_File;

		Select 病人病历记录_ID.Nextval Into v_报告id From Dual;
		
		ZL_病人病历_INSERT(v_报告id,r_Advice.病人id,r_Advice.主页id,r_Advice.挂号单,r_Advice.婴儿,r_Advice.病人科室id,r_File.种类,r_File.ID,r_File.名称,NULL,v_医嘱id);
		UPDATE 病人病历记录 SET 书写人=NULL,书写日期=NULL WHERE ID=v_报告id;
	
		v_Index:=0;
		FOR r_Element In c_Element(v_病历文件id) LOOP
			v_Index:=v_Index+1;

			Select 病人病历内容_ID.Nextval Into v_病历内容id From Dual;
			ZL_病人病历内容_INSERT(v_病历内容id,NULL,v_报告id,v_Index,r_Element.类型,r_Element.编码,r_Element.文本转储,r_Element.标题文本,r_Element.标题显示,r_Element.标题字体,r_Element.标题位置,0,r_Element.内容字体,r_Element.内容位置,0,r_Element.嵌入方式);
			
			--0-文本段；1-附加表；2-所见单；3-标记图；4-专用纸
			If r_Element.类型=4 Then
				--专用纸
				If UPPER(r_Element.部件)='ZL9CISCORE.USRVERIFYREPORT' then				
					--检验专用纸
					v_VerIndex:=0;
					FOR r_ElementVerfy In c_ElementVerfy(v_医嘱id) LOOP
						
						v_VerIndex:=v_VerIndex+1;

						ZL_病人病历所见单_SAVE(v_病历内容id,v_VerIndex,2,r_ElementVerfy.中文名,NULL,NULL,NULL,NULL,
									NULL,NULL,NULL,r_ElementVerfy.报告项目ID,r_ElementVerfy.结果类型,r_ElementVerfy.计算单位,''||''''''||r_ElementVerfy.结果参考);
					END LOOP;
				End If;

				If UPPER(r_Element.部件)='ZL9CISCORE.USRMEDICALGROUP' then				
					--体检小结专用纸
					NULL;
				End If;

				If UPPER(r_Element.部件)='ZL9CISCORE.USRMEDICALSUM' then				
					--体检总结专用纸
					NULL;
				End If;
			End If;

			If r_Element.类型=2 Then
				--所见单
				FOR r_ElementPaper In c_ElementPaper(r_Element.ID) LOOP
					
					ZL_病人病历所见单_SAVE(v_病历内容id,r_ElementPaper.控件号,r_ElementPaper.控件类,r_ElementPaper.标题,r_ElementPaper.行,r_ElementPaper.列,r_ElementPaper.宽,r_ElementPaper.高,
								r_ElementPaper.对齐,r_ElementPaper.不可写,r_ElementPaper.可屏蔽,r_ElementPaper.所见项ID,r_ElementPaper.数值类型,r_ElementPaper.计量单位,NULL);

				END LOOP;
			End If;

		END LOOP;

		Close c_File;
		Close c_Advice;

		Update 病人医嘱发送 Set 报告id=v_报告id Where 医嘱id In (Select ID From 病人医嘱记录 Where ID=v_医嘱id Union All Select 相关ID As ID From 病人医嘱记录 Where 相关ID=v_医嘱id);
	End If;


EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检项目报告_EMPTY;
/


CREATE OR REPLACE Procedure ZL_体检登记记录_单项填写(
	病历id_IN		病人病历所见单.病历id%TYPE,
	所见项id_IN		病人病历所见单.所见项id%TYPE,
	所见内容_IN		病人病历所见单.所见内容%TYPE
) IS
	v_Temp			Varchar2(255);
	v_人员编号		人员表.编号%Type;
	v_人员姓名		人员表.姓名%Type;
Begin
	
	--当前操作人员
	v_Temp:=zl_Identity;
	v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
	v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
	v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
	v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

	UPDATE 病人病历所见单 SET 所见内容=所见内容_IN WHERE 病历id=病历id_IN AND 所见项id=所见项id_IN;	
	UPDATE 病人病历记录 SET 书写人=v_人员姓名,书写日期=SYSDATE WHERE 书写人 IS NULL AND ID=(SELECT A.病历记录id FROM 病人病历内容 A,病人病历所见单 B WHERE A.ID=B.病历id AND A.ID=病历id_IN AND ROWNUM<2);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检登记记录_单项填写;
/
--}陈福容

--病历部分 By:赵彤宇
CREATE OR REPLACE PROCEDURE ZL_病人病历修订_INSERT(
  RecID 病人病历修订记录.病历记录ID%Type,
  Writer 病人病历修订记录.书写人%Type
)
AS
  nVersion number;
  vLastWriter varchar2(50);
  dLastDate date;
begin
  Select Max(版本序号)+1 Into nVersion From 病人病历修订记录 Where 病历记录ID=Recid;
  If nVersion Is Null Then
    nVersion:=1;
  End If;
  Select 审阅人,审阅日期 Into vLastWriter,dLastDate From 病人病历记录 Where ID=RecID;

  Insert Into 病人病历修订记录(ID,病历记录ID,书写人,书写日期,版本序号) Values(
    病人病历修订记录_ID.Nextval,RecID,vLastWriter,dLastDate,nVersion);
  Update 病人病历记录 Set 书写人=Decode(书写人,Null,Writer,书写人),书写日期=Decode(书写日期,Null,SYSDATE,书写日期),审阅人=Writer,审阅日期=Sysdate Where ID=RecID;

  Update 病人病历内容 Set 病历记录ID=null,病历修订ID=病人病历修订记录_ID.CurrVal
    Where 病历记录ID=RecID;
  EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
end ZL_病人病历修订_INSERT;
/



----------------------------------------------------------------------------
---  DELETE   for   材料加成方案
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_材料加成方案_DELETE
IS
BEGIN
	Delete From 材料加成方案;
EXCEPTION
	WHEN OTHERS THEN
		zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_材料加成方案_DELETE;
/


----------------------------------------------------------------------------
---  INSERT   for   材料加成方案
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_材料加成方案_INSERT(
	序号_IN		IN 材料加成方案.序号%TYPE,
	最低价_IN	IN 材料加成方案.最低价%TYPE,
	最高价_IN	IN 材料加成方案.最高价%TYPE,
	加成率_IN	IN 材料加成方案.加成率%TYPE,
	说明_IN		IN 材料加成方案.说明%TYPE
)
IS
BEGIN
	Insert Into 材料加成方案
		(序号,最低价,最高价,加成率,说明)
		VALUES
		(序号_IN,最低价_IN,最高价_IN,加成率_IN,说明_IN);
EXCEPTION
	WHEN OTHERS THEN
		zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_材料加成方案_INSERT;
/


CREATE OR REPLACE Procedure ZL_病人医嘱计价_Insert(
--功能：指定医嘱的计价
    医嘱ID_IN		病人医嘱计价.医嘱ID%TYPE,
    收费细目ID_IN	病人医嘱计价.收费细目ID%TYPE,
    数量_IN			病人医嘱计价.数量%TYPE,
    单价_IN			病人医嘱计价.单价%TYPE,
	从项_IN			病人医嘱计价.从项%TYPE:=Null,
	执行科室ID_IN	病人医嘱计价.执行科室ID%Type:=Null
) IS
Begin
    Insert Into 病人医嘱计价(
        医嘱ID,收费细目ID,数量,单价,从项,执行科室ID)
    Values(
        医嘱ID_IN,收费细目ID_IN,数量_IN,单价_IN,从项_IN,执行科室ID_IN);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱计价_Insert;
/


CREATE OR REPLACE PROCEDURE zl_卫生材料_Update (
    材料ID_IN        IN 材料特性.材料ID%TYPE,
    分类ID_IN        IN 诊疗分类目录.id%type,
    编码_IN            IN 收费项目目录.编码%TYPE,
    品名_IN            IN 收费项目别名.名称%TYPE := NULL,
    规格_IN            IN 收费项目目录.规格%TYPE,
    产地_IN            IN 收费项目目录.产地%TYPE := NULL,
    拼音_IN            IN 收费项目别名.简码%TYPE := NULL,
    五笔_IN            IN 收费项目别名.简码%TYPE := NULL,
    标识主码_IN		IN 收费项目目录.标识主码%TYPE := NULL,
    标识子码_IN		IN 收费项目目录.标识子码%TYPE := NULL,
    材料来源_IN        IN 材料特性.材料来源%TYPE := NULL,
    货源情况_IN        IN 材料特性.货源情况%TYPE := NULL,
    散装单位_IN        IN 收费项目目录.计算单位%TYPE := NULL,
    包装单位_IN        IN 材料特性.包装单位%TYPE := NULL,
    换算系数_IN        IN 材料特性.换算系数%TYPE := NULL,
    是否变价_IN        IN 收费项目目录.是否变价%TYPE := NULL,
    指导批发价_IN        IN 材料特性.指导批发价%TYPE := NULL,
    扣率_IN            IN 材料特性.扣率%TYPE := 95,
    指导零售价_IN        IN 材料特性.指导零售价%TYPE := NULL,
    指导差价率_IN        IN 材料特性.指导差价率%TYPE := NULL,
    费用类型_IN        IN 收费项目目录.费用类型%TYPE := NULL,
    服务对象_IN        IN 收费项目目录.服务对象%TYPE := NULL,
    屏蔽费别_IN        IN 收费项目目录.屏蔽费别%TYPE := 0,
    库房分批_IN        IN 材料特性.库房分批%TYPE := NULL,
    在用分批_IN        IN 材料特性.在用分批%TYPE := NULL,
    最大效期_IN        IN 材料特性.最大效期%TYPE := NULL,
    灭菌效期_IN        IN 材料特性.灭菌效期%TYPE := NULL,
    无菌性材料_IN        IN 材料特性.无菌性材料%TYPE := NULL,
    一次性材料_IN        IN 材料特性.一次性材料%TYPE := NULL,
    原材料_IN        IN 材料特性.原材料%TYPE := NULL,
    差价让利比_IN        IN 材料特性.差价让利比%TYPE := 0,
    成本价_IN        IN 材料特性.成本价%TYPE := 0,
    跟踪在用_IN        IN 材料特性.跟踪在用%TYPE := NULL,
    当前售价_IN        IN 收费价目.现价%TYPE := 0,
    收入ID_IN        IN 收费价目.收入项目ID%TYPE := NULL,
    批准文号_IN  IN 材料特性.批准文号%TYPE := NULL,
    注册商标_IN  IN 材料特性.注册商标%TYPE := NULL
) IS
    m诊疗ID        诊疗项目目录.ID%type;
    mErrMsg        varchar2(200);
    mErrItem    EXCEPTION;
    m发生        integer;
    m跟踪在用    INTEGER;
    mCount         INTEGER ;
BEGIN
    mErrMsg:='无';
    --修改诊疗项目
    BEGIN 
        SELECT 诊疗ID,跟踪在用  INTO  m诊疗ID,m跟踪在用 FROM 材料特性 WHERE 材料id=材料id_IN;
    EXCEPTION 
        WHEN OTHERS THEN 
        mErrMsg:='[ZLSOFT]不存在诊疗项目,可能被其他用户删除了,请检查![ZLSOFT]';
    END;
    IF mErrMsg<>'无' THEN 
        RAISE mErrItem;
    END IF ;
    
    --如果更新前的材料为跟踪在用,如果改为了不跟踪则需判断库存
    IF m跟踪在用=1 AND 跟踪在用_IN<>1 THEN 
        BEGIN 
            SELECT count(*) INTO mCount  FROM 药品库存 
            WHERE 药品id=材料id_In AND ( nvl(可用数量,0)<>0 or nvl(实际数量,0)<>0 or 
                nvl(实际金额,0)<>0 or nvl(实际差价,0)<>0);
            IF mcount <>0 THEN 
                mErrMsg:='[ZLSOFT]该卫生材料存在库存,不能取消跟踪在用属性,请检查![ZLSOFT]';
            END IF ;
        EXCEPTION 
            WHEN OTHERS THEN 
            null;
        END ;

    END IF ;
    IF mErrMsg<>'无' THEN 
        RAISE mErrItem;
    END IF ;

    UPDATE 诊疗项目目录 
    SET    分类id = 分类id_IN,
        编码 = 编码_IN,
        名称 = substr(品名_IN||' '||规格_IN,1,60),
        计算单位 = 散装单位_IN,
        服务对象 = 服务对象_IN
    WHERE id=m诊疗id;

    
    IF 拼音_IN IS NULL THEN 
        DELETE 诊疗项目别名 WHERE 诊疗项目id = m诊疗ID AND 码类=1 AND 名称= 品名_IN  AND 性质=1; 
    ELSE 
        UPDATE 诊疗项目别名
        SET 名称 = 品名_IN,
            简码 = 拼音_IN
        where 诊疗项目id = m诊疗ID AND 性质=1 AND 码类=1;
    END IF ;

    IF 五笔_IN IS NULL THEN 
        DELETE 诊疗项目别名 WHERE 诊疗项目id = m诊疗ID AND 码类=2 AND 名称= 品名_IN  AND 性质=1; 
    ELSE 
        UPDATE 诊疗项目别名
        SET 名称 = 品名_IN,
            简码 = 五笔_IN
        where 诊疗项目id = m诊疗ID AND 性质=1 AND 码类=2;
    END IF ;

    --规格信息
    update 收费项目目录
        set 编码=编码_IN,名称=品名_IN,规格=规格_IN,标识主码=标识主码_IN,标识子码=标识子码_IN,产地=产地_IN,是否变价=是否变价_IN,计算单位=散装单位_IN,
        费用类型=费用类型_IN,服务对象=服务对象_IN,屏蔽费别=屏蔽费别_IN
    where ID=材料ID_IN;

    IF SQL%ROWCOUNT=0 THEN
        mErrMsg:='[ZLSOFT]该卫生材料可能被其他用户删除了,请检查![ZLSOFT]';
        RAISE mErrItem;
    END IF;

    --材料特性
    UPDATE  材料特性
        SET 最大效期=最大效期_IN,
            灭菌效期=灭菌效期_IN,
            无菌性材料=无菌性材料_IN,
            一次性材料=一次性材料_IN,
            原材料=原材料_IN,
            货源情况=货源情况_IN,
            包装单位=包装单位_IN,
            换算系数=换算系数_IN,
            指导批发价=指导批发价_IN,
            指导零售价=指导零售价_IN,
            指导差价率=指导差价率_IN,
            扣率=扣率_IN,
            库房分批=库房分批_IN,
            在用分批=在用分批_IN,
            材料来源=材料来源_IN,
            差价让利比=差价让利比_IN,
            成本价=成本价_IN,
            跟踪在用=跟踪在用_IN,
	    批准文号=批准文号_IN,
	    注册商标=注册商标_In
    WHERE 材料id=材料id_IN;

    IF 拼音_IN IS NULL THEN 
        DELETE 收费项目别名 WHERE 收费细目id = 材料ID_IN AND 码类=1 AND 名称= 品名_IN  AND 性质=1; 
    ELSE 
        UPDATE 收费项目别名
        SET 名称 = 品名_IN,
            简码 = 拼音_IN
        where 收费细目id = 材料ID_IN AND 性质=1 AND 码类=1;
            if sql%rowcount=0 then
               INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (材料ID_IN,品名_IN,1,拼音_IN,1);
            end if;
    END IF ;

    IF 五笔_IN IS NULL THEN 
        DELETE 收费项目别名 WHERE 收费细目id = 材料ID_IN AND 码类=2 AND 名称= 品名_IN  AND 性质=1; 
    ELSE 
        UPDATE 收费项目别名
        SET 名称 = 品名_IN,
            简码 = 五笔_IN
        WHERE 收费细目id = 材料ID_IN AND 性质=1 AND 码类=2;
        if sql%rowcount=0 then
            INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (材料ID_IN,品名_IN,1,拼音_IN,2);
        end if;
    END IF ;

    IF 品名_IN is null then
        delete 收费项目别名 where 收费细目id=材料ID_IN and 性质=3;
    else
        if 拼音_IN IS NULL THEN
            delete 收费项目别名 where 收费细目id=材料ID_IN and 性质=1 and 码类=1;
        else
            update 收费项目别名 set 名称=品名_IN,简码=拼音_IN where 收费细目id=材料ID_IN and 性质=1 and 码类=1;
            if sql%rowcount=0 then
               INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (材料ID_IN,品名_IN,1,拼音_IN,1);
            end if;
        end if;
        if 五笔_IN IS NULL THEN
           delete 收费项目别名 where 收费细目id=材料ID_IN and 性质=1 and 码类=2;
        else
            update 收费项目别名 set 名称=品名_IN,简码=五笔_IN where 收费细目id=材料ID_IN and 性质=1 and 码类=2;
            if sql%rowcount=0 then
               INSERT INTO 收费项目别名(收费细目id,名称,性质,简码,码类) VALUES (材料ID_IN,品名_IN,1,五笔_IN,2);
            end if;
        end if;
    END IF;

    --定价信息：如果已经有发生，则不允许直接更改这些信息
    Select nvl(Count(*),0) Into m发生 From 药品收发记录 Where 药品id=材料ID_IN And rownum<2;

    If m发生=0 Then 
        Update 收费项目目录 set 是否变价=是否变价_IN where ID=材料ID_IN;
        Update 材料特性 Set 成本价=成本价_IN Where 材料ID=材料ID_IN;

        if 收入ID_IN is not null Then
           Update 收费价目
           Set 现价=当前售价_IN,收入项目ID=收入ID_IN,变动原因=1,调价说明='修改定价',调价人=User
           Where 收费细目ID=材料ID_IN
             --And (终止日期 Is Null Or 终止日期=to_date('3000-01-01','YYYY-MM-DD'));
              And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因=1;
 
           If Sql%Rowcount=0 Then
              insert into 收费价目(ID,原价ID,收费细目ID,原价,现价,收入项目ID,变动原因,调价说明,调价人,执行日期,终止日期)
              values (收费价目_ID.Nextval,null,材料ID_IN,0,当前售价_IN,收入ID_IN,1,'新增定价',user,sysdate,to_date('3000-01-01','YYYY-MM-DD'));
           End If;
        end if;
    End If;

    --材料生产商比较增加
    if 产地_IN is not null then
        update 材料生产商 set 名称=产地_IN where 名称=产地_IN;
        if sql%rowcount=0 then
              Insert INTO 材料生产商(编码,名称,简码)
              select nvl(max(to_number(编码)),0)+1,产地_IN,ZLSpellCode(产地_IN) from 材料生产商;
        end if;
    end if;

EXCEPTION
    WHEN mErrItem THEN 
        raise_application_error(-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_卫生材料_Update;
/


CREATE OR REPLACE PROCEDURE zl_药品外购_Insert (
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    供药单位ID_IN IN 药品收发记录.供药单位ID%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    产地_IN IN 药品收发记录.产地%TYPE := NULL,
    批号_IN IN 药品收发记录.批号%TYPE := NULL,
    效期_IN IN 药品收发记录.效期%TYPE := NULL,
    实际数量_IN IN 药品收发记录.实际数量%TYPE := NULL,
    成本价_IN IN 药品收发记录.成本价%TYPE := NULL,
    成本金额_IN IN 药品收发记录.成本金额%TYPE := NULL,
    扣率_IN IN 药品收发记录.扣率%TYPE := NULL,
    零售价_IN IN 药品收发记录.零售价%TYPE := NULL,
    零售金额_IN IN 药品收发记录.零售金额%TYPE := NULL,
    差价_IN IN 药品收发记录.差价%TYPE := NULL,
    摘要_IN IN 药品收发记录.摘要%TYPE := NULL,	
    填制人_IN IN 药品收发记录.填制人%TYPE := NULL,
    发票号_IN IN 应付记录.发票号%TYPE := NULL,
    发票日期_IN IN 应付记录.发票日期%TYPE := NULL,
    发票金额_IN IN 应付记录.发票金额%TYPE := NULL,
    填制日期_IN IN 药品收发记录.填制日期%TYPE := NULL,
	外观_IN IN 药品收发记录.外观%TYPE:=NULL,
	产品合格证_IN IN 药品收发记录.产品合格证%TYPE:=NULL,
	核查人_IN IN 药品收发记录.配药人%TYPE := NULL,
	核查日期_IN IN 药品收发记录.配药日期%TYPE :=NULL,
	批次_IN IN 药品收发记录.批次%TYPE:=0,
	退货_IN IN NUMBER:=1,
  生产日期_IN In 药品收发记录.生产日期%TYPE := Null,
  批准文号_IN In 药品收发记录.批准文号%TYPE := Null
)
IS
    v_NO 应付记录.NO%TYPE;		 --应付记录的NO
	v_商品名 收费项目目录.名称%TYPE;--通用名称
	v_规格 收费项目目录.规格%TYPE;
	v_产地 收费项目目录.规格%TYPE;
	v_单位 收费项目目录.计算单位%TYPE;
    V_lngID 药品收发记录.ID%TYPE;--收发ID
	V_应付ID 应付记录.ID%TYPE;	 --应付记录的ID
    V_入出类别ID 药品收发记录.入出类别ID%TYPE;--入出类别ID
    V_入出系数 药品收发记录.入出系数%TYPE;--入出系数
    V_批次 药品收发记录.批次%TYPE := NULL;--批次
    v_药库分批 INTEGER;--是否药库分批    1:分批；0：不分批
    v_药房分批 integer;--是否药房分批       1:分批；0：不分批
	v_指导批价 药品规格.指导批发价%TYPE;

	v_库存数量 药品库存.实际数量%TYPE;
	v_退货数量 药品库存.实际数量%TYPE;
	err_MSG VARCHAR2(255);
	ERR_NOENOUGH EXCEPTION ;
BEGIN
	--取该药品的商品名
	v_产地:='';
	SELECT 名称,规格,计算单位 INTO V_商品名,v_规格,v_单位 FROM 收费项目目录 WHERE ID=药品ID_IN;
	IF V_规格 IS NOT NULL THEN
		IF INSTR(V_规格,'|')<>0 THEN
			V_产地:=SUBSTR(V_规格,INSTR(V_规格,'|'));
			V_规格:=SUBSTR(V_规格,INSTR(V_规格,'|')-1);
		END IF ;
	END IF ;

    SELECT 药品收发记录_ID.Nextval
      INTO V_lngID
      FROM Dual;
    SELECT NVL (药库分批, 0),NVL (药房分批, 0),NVL(指导批发价,0) 
      INTO v_药库分批,v_药房分批,v_指导批价
      FROM 药品规格
     WHERE 药品ID = 药品ID_IN;

    IF v_药房分批=0 then
        IF v_药库分批 = 1 THEN
            BEGIN
                SELECT DISTINCT 0 INTO v_药库分批
                  FROM 部门性质说明
                 WHERE ((工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室'))
                    AND 部门ID = 库房ID_IN;
            EXCEPTION
                WHEN OTHERS THEN
                    v_药库分批 := 1;
            END;

            IF v_药库分批 = 1 THEN
                V_批次 := V_lngID;
            END IF;
        END IF;
    else
        V_批次 := V_lngID;
    END if;

    SELECT B.ID, B.系数
      INTO V_入出类别ID, V_入出系数
      FROM 药品单据性质 A, 药品入出类别 B
     WHERE A.类别ID = B.ID
        AND A.单据 = 1
        AND ROWNUM < 2;

    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,供药单位ID,入出类别ID,入出系数,药品ID,
        批次,产地,批号,效期,填写数量,实际数量,成本价,成本金额,扣率,零售价,零售金额,
        差价,摘要,填制人,填制日期,配药人,配药日期,发药方式,单量,外观,产品合格证,生产日期,批准文号)
	VALUES (V_lngID,1,1,NO_IN,序号_IN,库房ID_IN,供药单位ID_IN,V_入出类别ID,
		V_入出系数,药品ID_IN,Decode(退货_IN,-1,批次_IN,V_批次),产地_IN,批号_IN,效期_IN,退货_IN*实际数量_IN,
		退货_IN*实际数量_IN,成本价_IN,退货_IN*成本金额_IN,扣率_IN,零售价_IN,退货_IN*零售金额_IN,
		退货_IN*差价_IN,摘要_IN,填制人_IN,填制日期_IN,核查人_IN,核查日期_IN,DECODE(退货_IN,-1,1,0),v_指导批价,外观_IN,产品合格证_IN,生产日期_IN,批准文号_IN);

    IF 发票号_IN IS NOT NULL Then
      --如果是第一笔明细,则产生应付记录的NO
    	BEGIN
    		SELECT NO INTO V_NO FROM 应付记录
    		WHERE 系统标识=1 AND 记录性质=0 AND 记录状态=1
    			AND 入库单据号=NO_IN AND ROWNUM<2;
    	EXCEPTION
    		WHEN OTHERS THEN V_NO:=NEXTNO(67);
    	END ; 
    	SELECT 应付记录_ID.NEXTVAL INTO V_应付ID FROM DUAL;
        INSERT INTO 应付记录
		(ID,记录性质,记录状态,单位ID,NO,系统标识,收发ID,入库单据号,单据金额,发票号,发票日期,发票金额,品名,
		规格,产地,批号,计量单位,数量,采购价,采购金额,填制人,填制日期,审核人,审核日期,摘要,项目ID,序号)
        VALUES (V_应付ID,0,1,供药单位ID_IN,V_NO,1,V_LNGID,NO_IN,退货_IN*零售金额_IN,发票号_IN,发票日期_IN,退货_IN*发票金额_IN,V_商品名,
		V_规格,V_产地,批号_IN,V_单位,退货_IN*实际数量_IN,成本价_IN,退货_IN*成本金额_IN,填制人_IN,填制日期_IN,NULL,NULL,摘要_IN,药品ID_IN,序号_IN);
    END IF;
EXCEPTION
    WHEN ERR_NOENOUGH THEN
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]'||err_MSG||'[ZLSOFT]');
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品外购_Insert;
/

CREATE OR REPLACE PROCEDURE ZL_药品外购_VERIFY (
    NO_IN IN 药品收发记录.NO%TYPE := NULL,
    审核人_IN IN 药品收发记录.审核人%TYPE := NULL
)
IS
    ERR_ISVERIFIED EXCEPTION;
    ERR_ISBATCH EXCEPTION;
    V_BATCHCOUNT INTEGER;    --原不分批现在分批的药品的数量
    V_供药单位ID 药品收发记录.供药单位ID%TYPE;
    V_发票金额 应付记录.发票金额%TYPE;
	V_库存金额 药品库存.实际金额%TYPE;
	V_库存差价 药品库存.实际差价%TYPE;
	V_库存数量 药品库存.实际数量%TYPE;
	V_成本价   药品库存.上次采购价%TYPE;

    CURSOR C_药品收发记录
    IS
        SELECT ID, 实际数量, 零售金额, 差价, 库房ID, 药品ID, 批次, 供药单位ID,
                 成本价, 批号, 效期, 产地, 入出类别ID,生产日期,批准文号
          FROM 药品收发记录
         WHERE NO = NO_IN
            AND 单据 = 1
            AND 记录状态 = 1
         ORDER BY 药品ID;
BEGIN
    UPDATE 药品收发记录
        SET 审核人 = NVL (审核人_IN, 审核人),
             审核日期 = SYSDATE
     WHERE NO = NO_IN
        AND 单据 = 1
        AND 记录状态 = 1
        AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE ERR_ISVERIFIED;
    END IF;

    --主要针对原不分批现在分批的药品，不能对其审核
    SELECT COUNT(*) INTO V_BATCHCOUNT FROM
            药品收发记录 A,药品规格 B
    WHERE A.药品ID=B.药品ID
        AND A.NO=NO_IN
        AND A.单据=1
        AND A.记录状态=1
        AND NVL(A.批次,0)=0
        AND ((NVL(B.药库分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')))
            OR NVL(B.药房分批,0)=1);
    IF V_BATCHCOUNT>0 THEN
        RAISE ERR_ISBATCH;
    END IF;

    --原分批现不分批的药品,在审核时，要处理他
    UPDATE 药品收发记录 SET 批次=0
    WHERE
    ID in 
    (SELECT ID
        FROM 药品收发记录 A, 药品规格 B
        WHERE B.药品ID=A.药品ID
        AND A.NO=NO_IN
        AND A.单据 = 1
            AND A.记录状态 = 1
            AND NVL(A.批次,0)>0
            AND (NVL(B.药库分批,0)=0 OR
                (NVL(B.药房分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室'))))
            );

    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        --更改药品库存表的相应数据

        UPDATE 药品库存
            SET 可用数量 = NVL (可用数量, 0) + NVL (V_药品收发记录.实际数量, 0),
                 实际数量 = NVL (实际数量, 0) + NVL (V_药品收发记录.实际数量, 0),
                 实际金额 = NVL (实际金额, 0) + NVL (V_药品收发记录.零售金额, 0),
                 实际差价 = NVL (实际差价, 0) + NVL (V_药品收发记录.差价, 0),
                 上次供应商ID = NVL (V_药品收发记录.供药单位ID, 上次供应商ID),
                 上次采购价 = NVL (V_药品收发记录.成本价, 上次采购价),
                 上次批号 = NVL (V_药品收发记录.批号, 上次批号),
                 上次产地 = NVL (V_药品收发记录.产地, 上次产地),
                 效期 = NVL (V_药品收发记录.效期, 效期),
                 上次生产日期 = NVL (V_药品收发记录.生产日期, 上次生产日期),
                 批准文号 = NVL (V_药品收发记录.批准文号, 批准文号)
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次, 0) = NVL (V_药品收发记录.批次, 0)
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            INSERT INTO 药品库存
                            (
                                库房ID,
                                药品ID,
                                批次,
                                性质,
                                可用数量,
                                实际数量,
                                实际金额,
                                实际差价,
                                上次供应商ID,
                                上次采购价,
                                上次批号,
                                上次产地,
                                效期,
                                上次生产日期,
                                批准文号
                            )
                  VALUES (
                      V_药品收发记录.库房ID,
                      V_药品收发记录.药品ID,
                      V_药品收发记录.批次,
                      1,
                      V_药品收发记录.实际数量,
                      V_药品收发记录.实际数量,
                      V_药品收发记录.零售金额,
                      V_药品收发记录.差价,
                      V_药品收发记录.供药单位ID,
                      V_药品收发记录.成本价,
                      V_药品收发记录.批号,
                      V_药品收发记录.产地,
                      V_药品收发记录.效期,
                      V_药品收发记录.生产日期,
                      V_药品收发记录.批准文号
                  );
        END IF;

        --清除数量金额为零的记录
        DELETE
          FROM 药品库存
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL(可用数量,0) = 0
            AND NVL(实际数量,0) = 0
            AND NVL(实际金额,0) = 0
            AND NVL(实际差价,0) = 0;

        --更改药品收发汇总表的相应数据

        UPDATE 药品收发汇总
            SET 数量 = NVL (数量, 0) + NVL (V_药品收发记录.实际数量, 0),
                 金额 = NVL (金额, 0) + NVL (V_药品收发记录.零售金额, 0),
                 差价 = NVL (差价, 0) + NVL (V_药品收发记录.差价, 0)
         WHERE 日期 = TRUNC (SYSDATE)
            AND 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND 类别ID = V_药品收发记录.入出类别ID
            AND 单据 = 1;

        IF SQL%NOTFOUND THEN
            INSERT INTO 药品收发汇总
                            (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
                  VALUES (
                      TRUNC (SYSDATE),
                      V_药品收发记录.库房ID,
                      V_药品收发记录.药品ID,
                      V_药品收发记录.入出类别ID,
                      1,
                      V_药品收发记录.实际数量,
                      V_药品收发记录.零售金额,
                      V_药品收发记录.差价
                  );
        END IF;

		--更新该药品的成本价
		UPDATE 药品规格
		SET 成本价=V_药品收发记录.成本价 
		WHERE 药品ID=V_药品收发记录.药品ID;
    END LOOP;

    --对应付余额表进行处理
      --此处用一个块，主要是解决没有对应发票号的记录
    BEGIN
		UPDATE 应付记录
		SET 审核人=审核人_IN,审核日期=SYSDATE
		WHERE 入库单据号=NO_IN AND 系统标识=1 and 记录性质=0 And 记录状态=1;

        SELECT B.单位ID, SUM (发票金额)
          INTO V_供药单位ID, V_发票金额
          FROM 药品收发记录 A, 应付记录 B
         WHERE A.ID = B.收发ID
            AND A.NO = NO_IN
            AND A.单据 = 1 AND B.系统标识=1
         GROUP BY B.单位ID;

        IF NVL (V_供药单位ID, 0) <> 0 THEN
            UPDATE 应付余额
                SET 金额 = NVL (金额, 0) + NVL (V_发票金额, 0)
             WHERE 单位ID = V_供药单位ID
                AND 性质 = 1;

            IF SQL%NOTFOUND THEN
                INSERT INTO 应付余额
                                (单位ID, 性质, 金额)
                      VALUES (V_供药单位ID, 1, V_发票金额);
            END IF;
        END IF;
    EXCEPTION
        WHEN NO_DATA_FOUND THEN
            NULL;
    END;
EXCEPTION
    WHEN ERR_ISVERIFIED THEN
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]'
        );
    WHEN ERR_ISBATCH THEN
        RAISE_APPLICATION_ERROR (
            -20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品，不能审核！[ZLSOFT]'
        );
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_药品外购_VERIFY;
/


CREATE OR REPLACE PROCEDURE ZL_药品外购_STRIKE (
	行次_IN IN INTEGER,
    原记录状态_IN IN 药品收发记录.记录状态%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    冲销数量_IN Number,
    填制人_IN IN 药品收发记录.填制人%TYPE ,
    填制日期_IN IN 药品收发记录.填制日期%TYPE ,
    发票号_IN IN 应付记录.发票号%TYPE := NULL,
    发票日期_IN IN 应付记录.发票日期%TYPE := NULL,
    发票金额_IN IN 应付记录.发票金额%TYPE := NULL,
    全部冲销_IN IN 药品收发记录.实际数量%TYPE := 0,    --用于财务审核
    财务审核_IN In Number :=0  --财务审核标志
)
IS
    ERR_ISSTRIKED EXCEPTION;
    ERR_ISOUTSTOCK EXCEPTION;
    ERR_ISBATCH EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    V_BATCHCOUNT INTEGER;    --原不分批现在分批的药品的数量

	  v_应付ID 应付记录.ID%TYPE;
    V_库房ID 药品收发记录.库房ID%TYPE;
    V_供药单位ID 药品收发记录.供药单位ID%TYPE;
    V_入出类别ID 药品收发记录.入出类别ID%TYPE ;
    V_产地 药品收发记录.产地%TYPE ;
    V_批次 药品收发记录.批次%TYPE ;
    V_批号 药品收发记录.批号%TYPE ;
    V_效期 药品收发记录.效期%TYPE ;
    V_成本价 药品收发记录.成本价%TYPE ;
    V_成本金额 药品收发记录.成本金额%TYPE ;
    V_扣率 药品收发记录.扣率%TYPE ;
    V_零售价 药品收发记录.零售价%TYPE ;
    V_零售金额 药品收发记录.零售金额%TYPE ;
    V_差价 药品收发记录.差价%TYPE ;
    V_摘要 药品收发记录.摘要%TYPE ;
    V_剩余数量 Number;
    V_剩余成本金额 药品收发记录.成本金额%Type;
    V_剩余零售金额 药品收发记录.零售金额%Type;
    V_入出系数 药品收发记录.入出系数%TYPE;
    V_冲销数量 Number;
	  v_单量 药品收发记录.单量%TYPE;
    V_生产日期 药品收发记录.生产日期%TYPE;
    V_批准文号 药品收发记录.批准文号%TYPE;

    v_核查人 药品收发记录.配药人%TYPE;
    v_核查日期 药品收发记录.配药日期%TYPE;
    V_记录数 NUMBER;
	  V_发药方式 NUMBER ;
    V_收发ID 药品收发记录.ID%TYPE;

    --对冲销数量进行检查
    V_库存数 药品库存.实际数量%TYPE;
    V_药库分批 INTEGER;
    V_药房分批 INTEGER;
    V_分批属性 INTEGER;
    V_药库 INTEGER;
    V_分批 NUMBER;

    V_库存金额 NUMBER(16,5);
	  V_库存差价 NUMBER(16,5);
	  V_库存数量 NUMBER(16,5);
  	INTDIGIT NUMBER;
	  V_发票金额 NUMBER ;

BEGIN
	--获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';

	--取核查人
	SELECT min(配药人) 配药人,min(配药日期) 配药日期,sum(实际数量) 实际数量
	INTO v_核查人,v_核查日期,V_冲销数量
	FROM 药品收发记录
	WHERE NO=NO_IN AND 单据=1 AND 序号=序号_IN Group By 配药人,配药日期 ;

  If 财务审核_IN=0 Then
    V_冲销数量:=冲销数量_IN;
  End If;

    IF 行次_IN =1 THEN
        UPDATE 药品收发记录
            SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3)
         WHERE NO = NO_IN
            AND 单据 = 1
            AND 记录状态 =原记录状态_IN ;

        IF SQL%ROWCOUNT = 0 THEN
            RAISE ERR_ISSTRIKED;
        END IF;
    END IF;

    --主要针对原不分批现在分批的药品，不能对其审核
    SELECT COUNT(*) INTO V_BATCHCOUNT FROM
            药品收发记录 A,药品规格 B
    WHERE A.药品ID=B.药品ID
        AND A.NO=NO_IN
        AND A.单据=1
        AND MOD(A.记录状态,3)=0
        AND NVL(A.批次,0)=0
        AND A.药品ID+0=药品ID_IN
        AND ((NVL(B.药库分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')))
            or nvl(b.药房分批,0)=1);
    IF V_BATCHCOUNT>0 THEN
        RAISE ERR_ISBATCH;
    END IF;

    SELECT SUM(A.实际数量) AS 剩余数量,SUM(A.成本金额) AS 剩余成本金额,SUM(A.零售金额) AS 剩余零售金额,A.库房ID,A.供药单位ID,A.入出类别ID,A.入出系数,NVL(A.批次,0),A.产地,A.批号,A.效期,A.成本价,A.扣率,A.零售价,A.摘要,B.药库分批,B.药房分批,A.发药方式,A.单量,A.生产日期,A.批准文号
    INTO  V_剩余数量,v_剩余成本金额,V_剩余零售金额,V_库房ID,V_供药单位ID,V_入出类别ID,V_入出系数,V_批次,V_产地,V_批号,V_效期,V_成本价,V_扣率,V_零售价,V_摘要,V_药库分批,V_药房分批,V_发药方式,V_单量,V_生产日期,V_批准文号
    FROM 药品收发记录 A,药品规格 B
    WHERE A.NO=NO_IN And A.药品ID=B.药品ID AND A.单据=1 AND A.药品ID=药品ID_IN AND A.序号=序号_IN
    GROUP BY A.库房ID,A.供药单位ID,A.入出类别ID,A.入出系数,NVL(A.批次,0),A.产地,A.批号,A.效期,A.成本价,A.扣率,A.零售价,A.摘要,B.药库分批,B.药房分批,A.发药方式,A.单量,A.生产日期,A.批准文号;

    --判断该部门是药库还是药房
    BEGIN
        SELECT DISTINCT 0
        INTO v_药库
        FROM 部门性质说明
        WHERE (   (工作性质 LIKE '%药房')
              OR (工作性质 LIKE '制剂室'))
        AND 部门ID = V_库房ID;
    EXCEPTION
        WHEN OTHERS THEN V_药库:=1;
    END ;

    --根据部门性质,判断分批特性
    IF V_药库=0 THEN
        v_分批属性:=V_药房分批;
    ELSE
        V_分批属性:=V_药库分批;
    END IF ;

    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    V_分批:=0;
    IF V_分批属性=1 AND V_批次<>0 THEN
        V_分批:=V_批次;
    END IF ;

    --取库存数
    BEGIN
        SELECT Nvl(实际数量,0) INTO V_库存数 FROM 药品库存
        WHERE 库房ID=V_库房ID AND 药品ID=药品ID_IN AND Nvl(批次,0)=V_分批 And 性质=1;
    EXCEPTION
        WHEN OTHERS THEN V_库存数:=0;
    END ;
   
    --如果库存数大于剩余数量,取剩余数量;否则取库存数
    IF V_库存数<V_剩余数量 And 财务审核_IN=0 THEN
        if 全部冲销_IN=1  then
            --不允许
            raise ERR_ISNONUM;
        Else
            v_剩余成本金额:=V_库存数/V_剩余数量*v_剩余成本金额;
            V_剩余零售金额:=V_库存数/V_剩余数量*V_剩余零售金额;
            V_剩余数量:=V_库存数;
        end if ;
    END IF ;

    IF 全部冲销_IN=1 And 财务审核_IN=0 THEN
        V_冲销数量:=V_剩余数量;
    END IF;

    --冲销数量大于剩余数量，不允许(财务审核除外)
    IF ABS(V_剩余数量)<ABS(V_冲销数量) And 财务审核_IN=0 THEN
        RAISE ERR_ISNONUM;
    END IF;

    If 财务审核_IN=0 Then
       V_成本金额:= ROUND(V_冲销数量/v_剩余数量*v_剩余成本金额,INTDIGIT);
       V_零售金额:= ROUND(V_冲销数量/v_剩余数量*V_剩余零售金额,INTDIGIT);
    Else
       V_成本金额:= ROUND(v_剩余成本金额,INTDIGIT);
       V_零售金额:= ROUND(V_剩余零售金额,INTDIGIT);
    End If;
    V_差价:=V_零售金额-V_成本金额;

    SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;

    INSERT INTO 药品收发记录
        ( ID,记录状态,单据,NO,序号,库房ID,供药单位ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,
        填写数量,实际数量,成本价,成本金额,扣率,零售价,零售金额,差价,摘要,填制人,填制日期,配药人,配药日期,审核人,审核日期,发药方式,单量,生产日期,批准文号
        )
    VALUES (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),1,NO_IN,序号_IN,V_库房ID,V_供药单位ID,
        V_入出类别ID,1,药品ID_IN,V_批次,V_产地,V_批号,V_效期,-V_冲销数量, -V_冲销数量,V_成本价,-V_成本金额,
        V_扣率,V_零售价,-V_零售金额,-V_差价,V_摘要,填制人_IN, 填制日期_IN,v_核查人,v_核查日期, 填制人_IN, 填制日期_IN,V_发药方式,v_单量,V_生产日期,V_批准文号);


    --对于冲销的单据也应该对应付余额表进行处理
    --只对填了发票号的记录进行处理
	v_发票金额:=Nvl(发票金额_IN,0);
    IF NVL (发票号_IN,' ') <> ' ' AND NVL (发票金额_IN, 0)<>0 THEN
		--对于财务审核的，要将剩余的发票金额全部冲销
		IF 全部冲销_IN=1 THEN
			SELECT SUM(B.发票金额) INTO v_发票金额
			FROM
				(SELECT ID
				FROM 药品收发记录
				WHERE 单据=1 AND NO=NO_IN AND 序号=序号_IN) A,应付记录 B
			WHERE A.ID=B.收发ID AND B.系统标识=1 And B.记录性质<>-1;
		END IF;

		UPDATE 应付余额
            SET 金额 = NVL (金额, 0) - NVL (v_发票金额, 0)
         WHERE 单位ID = V_供药单位ID
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            INSERT INTO 应付余额
            (单位ID, 性质, 金额)
            VALUES (V_供药单位ID, 1, -NVL (v_发票金额, 0));
        END IF;
    END IF;

    UPDATE 药品库存
    SET 可用数量 = NVL (可用数量, 0) -V_冲销数量,
             实际数量 = NVL (实际数量, 0) - V_冲销数量,
             实际金额 = NVL (实际金额, 0) - V_零售金额,
             实际差价 = NVL (实际差价, 0) -V_差价,
             上次供应商ID = V_供药单位ID,
             上次采购价 = V_成本价,
             上次批号 = V_批号,
             上次产地 = V_产地,
             效期 = V_效期,
             上次生产日期=V_生产日期,
             批准文号=V_批准文号
     WHERE 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND NVL (批次, 0) = NVL(V_分批,0)
        AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO 药品库存
        (库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次供应商ID,上次采购价,上次批号,上次产地,效期,上次生产日期,批准文号)
        VALUES
        (V_库房ID,药品ID_IN,V_分批,1,-V_冲销数量,-V_冲销数量,-V_零售金额,-V_差价,V_供药单位ID,V_成本价,V_批号,V_产地,V_效期,V_生产日期,V_批准文号) ;

    END IF;

    --清除数量金额为零的记录
    DELETE
    FROM 药品库存
    WHERE 库房ID = V_库房ID
    AND 药品ID = 药品ID_IN
    AND NVL (可用数量, 0) = 0
    AND NVL (实际数量, 0) = 0
    AND NVL (实际金额, 0) = 0
    AND NVL (实际差价, 0) = 0;

    --更改药品收发汇总表的相应数据

    UPDATE 药品收发汇总
        SET 数量 =NVL(数量,0) - V_冲销数量,
            金额 =NVL (金额, 0) -V_零售金额,
            差价 =NVL (差价, 0) -V_差价
     WHERE 日期 = TRUNC (SYSDATE)
        AND 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND 类别ID = V_入出类别ID
        AND 单据 = 1;

    IF SQL%NOTFOUND THEN
		INSERT INTO 药品收发汇总
        (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
        VALUES (
            TRUNC (SYSDATE),V_库房ID,药品ID_IN,V_入出类别ID,1,-V_冲销数量,-V_零售金额,-V_差价);
    END IF;

	--产生应付记录的冲销记录(先判断应付记录中是否已存在该记录对应的冲销记录,是则更新;否则新增)
	SELECT 应付记录_ID.NEXTVAL INTO V_应付ID FROM DUAL;
	INSERT INTO 应付记录
	(ID,记录性质,记录状态,单位ID,NO,系统标识,收发ID,入库单据号,单据金额,发票号,发票日期,发票金额,品名,
	规格,产地,批号,计量单位,数量,采购价,采购金额,填制人,填制日期,审核人,审核日期,摘要,项目ID,序号)
	SELECT V_应付ID,记录性质,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),单位ID,NO,1,V_收发ID,入库单据号,-V_零售金额,发票号,发票日期,-v_发票金额,品名,
	规格,产地,批号,计量单位,-V_冲销数量,采购价,-采购价*V_冲销数量,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN,摘要,项目ID,序号
	FROM 应付记录
	WHERE 收发ID=(SELECT ID FROM 药品收发记录 WHERE NO=NO_IN AND 单据=1 AND 序号=序号_IN And Mod(记录状态,3)=0) AND 系统标识=1 AND 记录性质=0;

	update 应付记录
	set 记录状态=3
	WHERE 收发ID=(SELECT ID FROM 药品收发记录 WHERE NO=NO_IN AND 单据=1 AND 序号=序号_IN And Mod(记录状态,3)=0) AND 系统标识=1 AND 记录性质=0;

EXCEPTION
    WHEN ERR_ISSTRIKED THEN
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]'
        );
    WHEN ERR_ISOUTSTOCK THEN
        RAISE_APPLICATION_ERROR (
            -20102, '[ZLSOFT]该单据中第' || 序号_IN || '行的药品已出库，不能冲销！[ZLSOFT]'
        );
    WHEN ERR_ISBATCH THEN
        RAISE_APPLICATION_ERROR (
            -20102, '[ZLSOFT]该单据中第' || 序号_IN || '行的药品原来不分批，现在分批的药品，不能冲销！[ZLSOFT]'
        );
    WHEN ERR_ISNONUM THEN
        RAISE_APPLICATION_ERROR (
            -20102, '[ZLSOFT]该单据中第' || 序号_IN || '行的药品冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]'
        );
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_药品外购_STRIKE;
/

CREATE OR REPLACE PROCEDURE zl_药品其他入库_Insert (
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    入出类别ID_IN IN 药品收发记录.入出类别ID%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    实际数量_IN IN 药品收发记录.实际数量%TYPE,
    成本价_IN IN 药品收发记录.成本价%TYPE,
    成本金额_IN IN 药品收发记录.成本金额%TYPE,
    零售价_IN IN 药品收发记录.零售价%TYPE,
    零售金额_IN IN 药品收发记录.零售金额%TYPE,
    差价_IN IN 药品收发记录.差价%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    填制日期_IN IN 药品收发记录.填制日期%TYPE,
    摘要_IN IN 药品收发记录.摘要%TYPE := NULL,
    产地_IN IN 药品收发记录.产地%TYPE := NULL,
    批号_IN IN 药品收发记录.批号%TYPE := NULL,
    效期_IN IN 药品收发记录.效期%TYPE := NULL,
    生产日期_IN In 药品收发记录.生产日期%TYPE := Null,
    批准文号_IN In 药品收发记录.批准文号%TYPE := Null
)
Is      
    V_lngID 药品收发记录.ID%TYPE;--收发ID
    V_入出系数 药品收发记录.入出系数%TYPE;
    V_批次 药品收发记录.批次%TYPE := NULL;--批次
    v_药库分批 INTEGER;--是否药库分批    1:分批;0：不分批
    v_药房分批 INTEGER;--是否药库分批    1:分批;0：不分批
BEGIN
    V_入出系数 := 1;
    SELECT 药品收发记录_ID.Nextval
      INTO V_lngID
      FROM Dual;
    SELECT NVL (药库分批, 0),nvl(药房分批,0)
      INTO v_药库分批,v_药房分批
      FROM 药品规格
     WHERE 药品ID = 药品ID_IN;

    IF v_药房分批=0 then
        IF v_药库分批 = 1 THEN
            BEGIN
                SELECT DISTINCT 0
                  INTO v_药库分批
                  FROM 部门性质说明
                 WHERE (   (工作性质 LIKE '%药房')
                          OR (工作性质 LIKE '制剂室'))
                    AND 部门ID = 库房ID_IN;
            EXCEPTION
                WHEN OTHERS THEN
                    v_药库分批 := 1;
            END;

            IF v_药库分批 = 1 THEN
                V_批次 := V_lngID;
            END IF;
        END IF;
    else
        V_批次 := V_lngID;
    END if;

    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,药品ID,
        批次,产地,批号,效期,填写数量,实际数量,成本价,成本金额,零售价,
        零售金额,差价,摘要,填制人,填制日期,生产日期,批准文号)
    VALUES (V_lngID,1,4,NO_IN,序号_IN,库房ID_IN,入出类别ID_IN,
        V_入出系数,药品ID_IN,V_批次,产地_IN,批号_IN,效期_IN,
        实际数量_IN,实际数量_IN,成本价_IN,成本金额_IN,零售价_IN,
        零售金额_IN,差价_IN,摘要_IN,填制人_IN,填制日期_IN,生产日期_IN,批准文号_IN
        );

Exception
   WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品其他入库_Insert;
/


CREATE OR REPLACE PROCEDURE zl_药品其他入库_verify (
    NO_IN IN 药品收发记录.NO%TYPE := NULL,
    审核人_IN IN 药品收发记录.审核人%TYPE := NULL
)
IS
    Err_isverified EXCEPTION;
    Err_isBatch exception;
    v_BatchCount integer;    --原不分批现在分批的药品的数量

	V_库存金额 药品库存.实际金额%TYPE;
	V_库存差价 药品库存.实际差价%TYPE;
	V_库存数量 药品库存.实际数量%TYPE;
	V_成本价   药品库存.上次采购价%TYPE;

    CURSOR C_药品收发记录
    IS
        SELECT ID, 实际数量, 零售金额, 差价, 库房ID, 药品ID, 批次, 成本价, 批号,
                 效期, 产地, 入出类别ID,生产日期,批准文号
          FROM 药品收发记录
         WHERE NO = NO_IN
            AND 单据 = 4
            AND 记录状态 = 1
         ORDER BY 药品ID;
BEGIN
    UPDATE 药品收发记录
        SET 审核人 = 审核人_IN,
             审核日期 = SYSDATE
     WHERE NO = NO_IN
        AND 单据 = 4
        AND 记录状态 = 1
        AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;
    
    --主要针对原不分批现在分批的药品，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount FROM        
            药品收发记录 a,药品规格 b
    WHERE a.药品id=b.药品id
        AND a.no=NO_IN 
        AND a.单据=4
        AND a.记录状态=1
        AND nvl(a.批次,0)=0
        AND ((nvl(b.药库分批,0)=1 AND a.库房id not in (select 部门id from  部门性质说明 where (工作性质 LIKE '%药房') or (工作性质 LIKE '制剂室')))
            or nvl(b.药房分批,0)=1);
        
    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  
    
    --原分批现不分批的药品,在审核时，要处理他
    UPDATE 药品收发记录 SET 批次=0
    WHERE 
    id=
    (SELECT id
        FROM 药品收发记录 a, 药品规格 b
        WHERE b.药品id=a.药品ID
        AND a.no=no_in
        AND a.单据 = 4 
            AND a.记录状态 = 1 
            AND nvl(a.批次,0)>0
            AND (nvl(b.药库分批,0)=0 or 
                (nvl(b.药房分批,0)=0 and a.库房id in (select 部门id from  部门性质说明 where (工作性质 LIKE '%药房') or (工作性质 LIKE '制剂室'))))
            );

    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        --更改药品库存表的相应数据

        UPDATE 药品库存
            SET 可用数量 = NVL (可用数量, 0) + NVL (V_药品收发记录.实际数量, 0),
                 实际数量 = NVL (实际数量, 0) + NVL (V_药品收发记录.实际数量, 0),
                 实际金额 = NVL (实际金额, 0) + NVL (V_药品收发记录.零售金额, 0),
                 实际差价 = NVL (实际差价, 0) + NVL (V_药品收发记录.差价, 0),
                 上次采购价 = NVL (V_药品收发记录.成本价, 上次采购价),
                 上次批号 = NVL (V_药品收发记录.批号, 上次批号),
                 上次产地 = NVL (V_药品收发记录.产地, 上次产地),
                 效期 = NVL (V_药品收发记录.效期, 效期),
                 上次生产日期 = NVL (V_药品收发记录.生产日期, 上次生产日期),
                 批准文号= NVL (V_药品收发记录.批准文号, 批准文号)
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次, 0) = NVL (V_药品收发记录.批次, 0)
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
                            (
                                库房ID,
                                药品ID,
                                批次,
                                性质,
                                可用数量,
                                实际数量,
                                实际金额,
                                实际差价,
                                上次采购价,
                                上次批号,
                                上次产地,
                                效期,
                                上次生产日期,
                                批准文号
                            )
                  VALUES (
                      V_药品收发记录.库房ID,
                      V_药品收发记录.药品ID,
                      V_药品收发记录.批次,
                      1,
                      V_药品收发记录.实际数量,
                      V_药品收发记录.实际数量,
                      V_药品收发记录.零售金额,
                      V_药品收发记录.差价,
                      V_药品收发记录.成本价,
                      V_药品收发记录.批号,
                      V_药品收发记录.产地,
                      V_药品收发记录.效期,
                      V_药品收发记录.生产日期,
                      V_药品收发记录.批准文号
                  );
        END IF;

        DELETE
          FROM 药品库存
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND nvl(可用数量,0) = 0 AND nvl(实际数量,0) = 0 AND nvl(实际金额,0) = 0 AND nvl(实际差价,0) = 0;

        --更改药品收发汇总表的相应数据

        UPDATE 药品收发汇总
            SET 数量 = NVL (数量, 0) + NVL (V_药品收发记录.实际数量, 0),
                 金额 = NVL (金额, 0) + NVL (V_药品收发记录.零售金额, 0),
                 差价 = NVL (差价, 0) + NVL (V_药品收发记录.差价, 0)
         WHERE 日期 = TRUNC (SYSDATE)
            AND 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND 类别ID = V_药品收发记录.入出类别ID
            AND 单据 = 4;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品收发汇总
                            (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
                  VALUES (
                      TRUNC (SYSDATE),
                      V_药品收发记录.库房ID,
                      V_药品收发记录.药品ID,
                      V_药品收发记录.入出类别ID,
                      4,
                      V_药品收发记录.实际数量,
                      V_药品收发记录.零售金额,
                      V_药品收发记录.差价
                  );
        END IF;

		UPDATE 药品规格
		SET 成本价=V_药品收发记录.成本价 
		WHERE 药品ID=V_药品收发记录.药品ID;
    END LOOP;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]'
        );
    when Err_isBatch then
        Raise_application_error ( 
            -20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品，不能审核！[ZLSOFT]' 
        );  
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品其他入库_verify;
/



CREATE OR REPLACE PROCEDURE zl_药品其他入库_strike (
    行次_IN IN INTEGER,
    原记录状态_IN IN 药品收发记录.记录状态%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    冲销数量_IN IN 药品收发记录.实际数量%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    填制日期_IN IN 药品收发记录.填制日期%TYPE
)
IS
    Err_isstriked EXCEPTION;
    Err_isoutstock EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    Err_isBatch EXCEPTION;
    v_BatchCount INTEGER;    --原不分批现在分批的药品的数量

    V_库房ID 药品收发记录.库房ID%TYPE; 
    V_入出类别ID 药品收发记录.入出类别ID%TYPE ;
    V_产地 药品收发记录.产地%TYPE ; 
    V_批次 药品收发记录.批次%TYPE ; 
    V_批号 药品收发记录.批号%TYPE ; 
    V_效期 药品收发记录.效期%TYPE ; 
    V_成本价 药品收发记录.成本价%TYPE ; 
    V_成本金额 药品收发记录.成本金额%TYPE ; 
    V_扣率 药品收发记录.扣率%TYPE ; 
    V_零售价 药品收发记录.零售价%TYPE ; 
    V_零售金额 药品收发记录.零售金额%TYPE ; 
    V_差价 药品收发记录.差价%TYPE ; 
    V_摘要 药品收发记录.摘要%TYPE ; 
    V_剩余数量 药品收发记录.实际数量%TYPE;
    V_剩余成本金额 药品收发记录.成本金额%Type;
    V_剩余零售金额 药品收发记录.零售金额%Type; 
    V_入出系数 药品收发记录.入出系数%TYPE; 
    V_生产日期 药品收发记录.生产日期%TYPE;
    V_批准文号 药品收发记录.批准文号%TYPE;
    
	  V_库存金额 NUMBER(16,5);
	  V_库存差价 NUMBER(16,5);
	  V_库存数量 NUMBER(16,5);

    V_记录数 NUMBER; 
    V_收发ID 药品收发记录.ID%TYPE; 

    --对冲销数量进行检查
    V_库存数 药品库存.实际数量%TYPE;
    V_药库分批 INTEGER;
    V_药房分批 INTEGER;
    V_分批属性 INTEGER;
    V_药库 INTEGER;
    V_分批 NUMBER;
	INTDIGIT NUMBER;
BEGIN
	--获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';

    IF 行次_IN =1 THEN
        UPDATE 药品收发记录 
        SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3) 
        WHERE NO = NO_IN AND 单据 = 4 AND 记录状态 =原记录状态_IN ; 

        IF SQL%ROWCOUNT = 0 THEN 
            RAISE ERR_ISSTRIKED; 
        END IF; 
    END IF;
    
    --主要针对原不分批现在分批的药品，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount 
    FROM 药品收发记录 a,药品规格 b
    WHERE a.药品id=b.药品id
        AND a.no=NO_IN 
        AND a.单据=4
        AND MOD(a.记录状态,3)=0
        AND a.药品ID+0=药品ID_IN
        AND nvl(a.批次,0)=0
        AND ((nvl(b.药库分批,0)=1 AND a.库房id not in (select 部门id from  部门性质说明 where (工作性质 LIKE '%药房') or (工作性质 LIKE '制剂室')))
            or nvl(b.药房分批,0)=1);

    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  
    
    SELECT SUM(A.实际数量) AS 剩余数量,SUM(A.成本金额) AS 剩余成本金额,SUM(A.零售金额) AS 剩余零售金额,A.库房ID,A.入出类别ID,A.入出系数,Nvl(A.批次,0),A.产地,A.批号,A.效期,A.成本价,A.扣率,A.零售价,A.摘要,B.药库分批,B.药房分批,A.生产日期,A.批准文号
    INTO  V_剩余数量,v_剩余成本金额,V_剩余零售金额,V_库房ID,V_入出类别ID,V_入出系数,V_批次,V_产地,V_批号,V_效期,V_成本价,V_扣率,V_零售价,V_摘要,V_药库分批,V_药房分批,V_生产日期,V_批准文号
    FROM 药品收发记录 A,药品规格 B
    WHERE A.NO=NO_IN AND A.单据=4 AND A.药品ID=B.药品ID AND A.药品ID=药品ID_IN AND A.序号=序号_IN
    GROUP BY A.库房ID,A.入出类别ID,A.入出系数,NVL(A.批次,0),A.产地,A.批号,A.效期,A.成本价,A.扣率,A.零售价,A.摘要,B.药库分批,B.药房分批,A.生产日期,A.批准文号;

    --判断该部门是药库还是药房
    BEGIN
        SELECT DISTINCT 0
        INTO v_药库
        FROM 部门性质说明
        WHERE (   (工作性质 LIKE '%药房')
              OR (工作性质 LIKE '制剂室'))
        AND 部门ID = V_库房ID;
    EXCEPTION 
        WHEN OTHERS THEN V_药库:=1;
    END ;
    
    --根据部门性质,判断分批特性
    IF V_药库=0 THEN 
        v_分批属性:=V_药房分批;
    ELSE
        V_分批属性:=V_药库分批;
    END IF ;

    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    V_分批:=0;
    IF V_分批属性=1 AND V_批次<>0 THEN 
        V_分批:=V_批次;
    END IF ;
    
    --取库存数
    BEGIN
        SELECT Nvl(实际数量,0) INTO V_库存数 FROM 药品库存 
        WHERE 库房ID=V_库房ID AND 药品ID=药品ID_IN AND Nvl(批次,0)=V_分批 And 性质=1;
    EXCEPTION 
        WHEN OTHERS THEN V_库存数:=0;
    END ;

    --如果库存数大于剩余数量,取剩余数量;否则取库存数
    IF V_库存数<V_剩余数量 Then
       v_剩余成本金额:=V_库存数/V_剩余数量*v_剩余成本金额;
       V_剩余零售金额:=V_库存数/V_剩余数量*V_剩余零售金额;
       V_剩余数量:=V_库存数; 
    END IF ;

    --冲销数量大于剩余数量，不允许
    IF abs(V_剩余数量)<abs(冲销数量_IN) THEN
        RAISE ERR_ISNONUM; 
    END IF;

    V_成本金额:= ROUND(冲销数量_IN/v_剩余数量*v_剩余成本金额,INTDIGIT);
    V_零售金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余零售金额,INTDIGIT);
    V_差价:=V_零售金额-V_成本金额;

    SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;

    Insert INTO 药品收发记录
    (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,药品ID,批次,产地,批号,
    效期,填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,审核人,审核日期,生产日期,批准文号)
    VALUES 
    (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2), 4, NO_IN, 序号_IN, V_库房ID, V_入出类别ID,
    V_入出系数, 药品ID_IN,V_批次,V_产地, V_批号, V_效期, -冲销数量_IN, -冲销数量_IN,V_成本价, -V_成本金额, 
    V_零售价, -V_零售金额, -V_差价, V_摘要, 填制人_IN,填制日期_IN,填制人_IN,填制日期_IN,V_生产日期,V_批准文号);

    --更改药品库存表的相应数据

    UPDATE 药品库存
        SET 可用数量 = NVL (可用数量, 0) - NVL (冲销数量_IN, 0),
             实际数量 = NVL (实际数量, 0) - NVL (冲销数量_IN, 0),
             实际金额 = NVL (实际金额, 0) - NVL (V_零售金额, 0),
             实际差价 = NVL (实际差价, 0) - NVL (V_差价, 0),
             上次采购价 = NVL (V_成本价, 上次采购价),
             上次批号 = NVL (V_批号, 上次批号),
             上次产地 = NVL (V_产地, 上次产地),
             效期 = NVL (V_效期, 效期),
             上次生产日期=NVL(V_生产日期,上次生产日期),
             批准文号=NVL(V_批准文号,批准文号)
     WHERE 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND NVL (批次,0) = NVL (V_分批, 0)
        AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品库存
        (库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,
        实际差价,上次采购价,上次批号,上次产地,效期,上次生产日期,批准文号)
        VALUES (
        V_库房ID,药品ID_IN,V_分批,1,-冲销数量_IN,-冲销数量_IN,-V_零售金额,
        -V_差价,V_成本价,V_批号,V_产地,V_效期,V_生产日期,V_批准文号);
    END IF;

    DELETE
      FROM 药品库存
     WHERE 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND nvl(可用数量,0) = 0
        AND nvl(实际数量,0) = 0
        AND nvl(实际金额,0) = 0
        AND nvl(实际差价,0) = 0;

    --更改药品收发汇总表的相应数据

    UPDATE 药品收发汇总
        SET 数量 = NVL (数量, 0) - NVL (冲销数量_IN, 0),
             金额 = NVL (金额, 0) - NVL (V_零售金额, 0),
             差价 = NVL (差价, 0) - NVL (V_差价, 0)
     WHERE 日期 = TRUNC (填制日期_IN)
        AND 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND 类别ID = V_入出类别ID
        AND 单据 = 4;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品收发汇总
        (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
        VALUES (
        TRUNC (填制日期_IN),V_库房ID,药品ID_IN,
        V_入出类别ID,4,-冲销数量_IN,-V_零售金额,-V_差价);
    END IF;

EXCEPTION
    WHEN Err_isstriked THEN
        Raise_application_error (-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
    WHEN Err_isoutstock THEN
        Raise_application_error (-20102, '[ZLSOFT]该单据中有一笔分批药品已出库，不能冲销！[ZLSOFT]');
    when Err_isBatch then
        Raise_application_error (-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品，不能冲销！[ZLSOFT]' );  
    WHEN ERR_ISNONUM THEN 
        RAISE_APPLICATION_ERROR (-20102, '[ZLSOFT]该单据中第' || 序号_IN || '行的药品冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]' ); 
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品其他入库_strike;
/

CREATE OR REPLACE PROCEDURE zl_药品其他出库_Insert (
    入出类别ID_IN IN 药品收发记录.入出类别ID%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    批次_IN IN 药品收发记录.批次%TYPE,
    填写数量_IN IN 药品收发记录.填写数量%TYPE,
    成本价_IN IN 药品收发记录.成本价%TYPE,
    成本金额_IN IN 药品收发记录.成本金额%TYPE,
    零售价_IN IN 药品收发记录.零售价%TYPE,
    零售金额_IN IN 药品收发记录.零售金额%TYPE,
    差价_IN IN 药品收发记录.差价%TYPE,
	外调价_IN IN 药品收发记录.单量%TYPE,
	外调单位_IN IN 药品收发记录.发药窗口%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    填制日期_IN IN 药品收发记录.填制日期%TYPE,
    产地_IN IN 药品收发记录.产地%TYPE := NULL,
    批号_IN IN 药品收发记录.批号%TYPE := NULL,
    效期_IN IN 药品收发记录.效期%TYPE := NULL,
    摘要_IN IN 药品收发记录.摘要%TYPE := NULL,
    批准文号_IN In 药品收发记录.批准文号%TYPE := Null
)
Is
    ERR_MutilROW EXCEPTION ;
    Err_isNOnumber EXCEPTION;
    intRecords NUMBER ;
    V_编码 收费项目目录.编码%TYPE;
    V_可用数量 药品库存.可用数量%TYPE;
    V_入出系数 药品收发记录.入出系数%TYPE;--收发ID
BEGIN
    V_入出系数 := -1;

    IF 批次_IN > 0 THEN
        BEGIN
            SELECT 可用数量
              INTO V_可用数量
              FROM 药品库存
             WHERE 药品ID = 药品ID_IN
                AND NVL (批次, 0) = 批次_IN
                AND 库房ID = 库房ID_IN
                AND 性质 = 1
                AND ROWNUM = 1;
        EXCEPTION
            WHEN OTHERS THEN
                V_可用数量 := 0;
        END;

        IF V_可用数量 - 填写数量_IN < 0 THEN
            RAISE Err_isNOnumber;
        END IF;
    END IF;

    Insert INTO 药品收发记录
    (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,
    填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,单量,发药窗口,批准文号)
	VALUES (
	药品收发记录_ID.Nextval,1,11,NO_IN,序号_IN,库房ID_IN,入出类别ID_IN,V_入出系数,药品ID_IN,批次_IN,
	产地_IN,批号_IN,效期_IN,填写数量_IN,填写数量_IN,成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,差价_IN,
	摘要_IN,填制人_IN,填制日期_IN,外调价_IN,外调单位_IN,批准文号_IN);
  
  --检查是否存在相同药品相同批次的数据，如果存在不允许保存
	SELECT COUNT(*) INTO intRecords
	FROM 药品收发记录
	WHERE 单据=11 AND NO=NO_IN AND 入出系数=-1 AND 药品ID+0=药品ID_IN AND Nvl(批次,0)=NVL(批次_IN,0);
	IF intRecords>1 THEN
		RAISE ERR_MutilROW;
	END IF ;
  
    --同时更新库存数
	UPDATE 药品库存
	SET 可用数量 = NVL (可用数量, 0) - 填写数量_IN
	WHERE 库房ID = 库房ID_IN
	AND 药品ID = 药品ID_IN
	AND NVL (批次, 0) = NVL (批次_IN, 0)
	AND 性质 = 1;

    --不插入批次是因为批次药品不够，不准出库
    IF SQL%NOTFOUND THEN
        Insert INTO 药品库存
		(库房ID, 药品ID, 性质, 可用数量,上次批号,效期,上次产地,批准文号)
		VALUES 
		(库房ID_IN, 药品ID_IN, 1, -填写数量_IN,批号_IN,效期_IN,产地_IN,批准文号_IN);
    END IF;

    DELETE
      FROM 药品库存
     WHERE 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND NVL (可用数量, 0) = 0
        AND NVL (实际数量, 0) = 0
        AND NVL (实际金额, 0) = 0
        AND NVL (实际差价, 0) = 0;
Exception
    WHEN ERR_MutilROW THEN
		SELECT 编码 INTO V_编码 FROM 收费项目目录 WHERE ID = 药品ID_IN;
		RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]编码为'||V_编码||'的药品，存在多条重复的记录，请合并为一条记录！[ZLSOFT]');
    
    WHEN Err_isNOnumber THEN
        SELECT 编码
          INTO V_编码
          FROM 收费项目目录
         WHERE ID = 药品ID_IN;
        Raise_application_error (
            -20101, '[ZLSOFT]编码为' || V_编码 || ',批号为' || 批号_IN ||
                          '的药库分批药品' ||
                          CHR (10) ||
                          CHR (13) ||
                          '可用库存数量不够！[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品其他出库_Insert;
/


CREATE OR REPLACE PROCEDURE zl_药品其他出库_DELETE (
    
    --删除药品收发记录及恢复相应的表：药品库存
    NO_IN IN 药品收发记录.NO%TYPE
)
IS
    Err_isverified EXCEPTION;

    CURSOR C_药品收发记录
    IS
        SELECT 填写数量, 库房ID, 批次, 药品ID,批号,效期,产地,批准文号
          FROM 药品收发记录
         WHERE NO = NO_IN
            AND 单据 = 11
         ORDER BY 药品ID;
BEGIN
    --通过循环，恢复原来的可用数量
    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        UPDATE 药品库存
            SET 可用数量 = NVL (可用数量, 0) + V_药品收发记录.填写数量
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次, 0) = NVL (V_药品收发记录.批次, 0)
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
                            (库房ID, 药品ID, 批次, 性质, 可用数量,上次批号,效期,上次产地,批准文号)
                  VALUES (
                      V_药品收发记录.库房ID,V_药品收发记录.药品ID,V_药品收发记录.批次,1,
                      V_药品收发记录.填写数量,V_药品收发记录.批号,V_药品收发记录.效期,V_药品收发记录.产地,V_药品收发记录.批准文号
                  );
        END IF;

        DELETE
          FROM 药品库存
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (可用数量, 0) = 0
            AND NVL (实际数量, 0) = 0
            AND NVL (实际金额, 0) = 0
            AND NVL (实际差价, 0) = 0;
    END LOOP;

    DELETE
      FROM 药品收发记录
     WHERE NO = NO_IN
        AND 单据 = 11
        AND 记录状态 = 1
        AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品其他出库_DELETE;
/


CREATE OR REPLACE PROCEDURE zl_药品其他出库_verify (
    序号_IN IN 药品收发记录.序号%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    批次_IN IN 药品收发记录.批次%TYPE,
    实际数量_IN IN 药品收发记录.实际数量%TYPE,
    成本价_IN IN 药品收发记录.成本价%TYPE,
    成本金额_IN IN 药品收发记录.成本金额%TYPE,
    零售金额_IN IN 药品收发记录.零售金额%TYPE,
    差价_IN IN 药品收发记录.差价%TYPE,
    入出类别ID_IN IN 药品收发记录.入出类别ID%TYPE,
    审核人_IN IN 药品收发记录.审核人%TYPE,
    审核日期_IN IN 药品收发记录.审核日期%Type
)
IS
    Err_isverified EXCEPTION;
    V_实际库存金额 药品库存.实际金额%TYPE;
    V_实际库存差价 药品库存.实际差价%TYPE;
    V_差价率 number(18,8);
    V_出库差价 药品库存.实际差价%TYPE;
    V_成本价 药品收发记录.成本价%TYPE;
    V_成本金额 药品收发记录.成本金额%TYPE;
	v_批号 药品收发记录.批号%TYPE;
	v_效期 药品收发记录.效期%TYPE;
	v_产地 药品收发记录.产地%TYPE;
  v_批准文号 药品收发记录.批准文号%TYPE;
	INTDIGIT NUMBER ;
BEGIN
	--获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';

    --由于领用处理允许在审核时改变实际数量，
      --所以首先对实际数量和其他相应的字段进行更新。
    BEGIN
        SELECT nvl(实际金额,0), nvl(实际差价,0)
          INTO V_实际库存金额, V_实际库存差价
          FROM 药品库存
         WHERE 药品ID = 药品ID_IN
            AND NVL (批次, 0) = 批次_IN
            AND 库房ID = 库房ID_IN
            AND 性质 = 1
            AND ROWNUM = 1;
    EXCEPTION
        WHEN OTHERS THEN
            V_实际库存金额 := 0;
    END;

    IF V_实际库存金额 <= 0 THEN
        BEGIN
            SELECT 指导差价率 / 100
              INTO V_差价率
              FROM 药品规格
             WHERE 药品ID = 药品ID_IN;
        EXCEPTION
            WHEN OTHERS THEN
                V_差价率 := 0;
        END;
    ELSE
        V_差价率 := V_实际库存差价 / V_实际库存金额;
    END IF;

    V_出库差价 := round(零售金额_IN * V_差价率,INTDIGIT);
    IF 实际数量_IN<=0 THEN 
        V_成本价 := 成本价_IN;
    ELSE 
        V_成本价 := (零售金额_IN - V_出库差价) / 实际数量_IN;
    END IF ;
    V_成本金额 := round(V_成本价 * 实际数量_IN,INTDIGIT);
	
	--提取药品其他出库单指定明细的批号,效期与产地信息
	SELECT 批号,效期,产地,批准文号 INTO v_批号,v_效期,v_产地,v_批准文号
	FROM 药品收发记录
	WHERE 单据=11 AND NO=NO_IN AND 序号=序号_IN;

    UPDATE 药品收发记录
        SET 审核人 = NVL (审核人_IN, 审核人),
             审核日期 = 审核日期_IN,
             成本价 = V_成本价,
             成本金额 = V_成本金额,
             差价 = V_出库差价
     WHERE NO = NO_IN
        AND 单据 = 11
        AND 药品ID = 药品ID_IN
        AND 序号 = 序号_IN
        AND 记录状态 = 1
        AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;

    --更改药品库存的相应数据

    UPDATE 药品库存
        SET 实际数量 = NVL (实际数量, 0) - 实际数量_IN,
             实际金额 = NVL (实际金额, 0) - 零售金额_IN,
             实际差价 = NVL (实际差价, 0) - V_出库差价
     WHERE 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND NVL (批次, 0) = NVL (批次_IN, 0)
        AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品库存
                        (
                            库房ID,
                            药品ID,
                            批次,
                            性质,
                            可用数量,
                            实际数量,
                            实际金额,
                            实际差价,
							上次批号,效期,上次产地,批准文号
                        )
              VALUES (
                  库房ID_IN,
                  药品ID_IN,
                  批次_IN,
                  1,
                  -实际数量_IN,
                  -实际数量_IN,
                  -零售金额_IN,
                  -V_出库差价,
				  v_批号,v_效期,v_产地,v_批准文号
              );
    END IF;

    DELETE
      FROM 药品库存
     WHERE 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND NVL (可用数量, 0) = 0
        AND NVL (实际数量, 0) = 0
        AND NVL (实际金额, 0) = 0
        AND NVL (实际差价, 0) = 0;

    --更药品收发汇总表的相应数据
    UPDATE 药品收发汇总
        SET 数量 = NVL (数量, 0) - 实际数量_IN,
             金额 = NVL (金额, 0) - 零售金额_IN,
             差价 = NVL (差价, 0) - V_出库差价
     WHERE 日期 = TRUNC (SYSDATE)
        AND 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND 类别ID = 入出类别ID_IN
        AND 单据 = 11;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品收发汇总
                        (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
              VALUES (
                  TRUNC (SYSDATE),
                  库房ID_IN,
                  药品ID_IN,
                  入出类别ID_IN,
                  11,
                  -实际数量_IN,
                  -零售金额_IN,
                  -V_出库差价
              );
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品其他出库_verify;
/


CREATE OR REPLACE PROCEDURE zl_药品其他出库_strike (
    行次_IN IN INTEGER,
    原记录状态_IN IN 药品收发记录.记录状态%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    冲销数量_IN IN 药品收发记录.实际数量%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    填制日期_IN IN 药品收发记录.填制日期%TYPE
)
IS
    Err_isstriked EXCEPTION;
    Err_isoutstock EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    Err_isBatch EXCEPTION;
    v_BatchCount INTEGER;    --原不分批现在分批的药品的数量

    V_库房ID 药品收发记录.库房ID%TYPE; 
    V_入出类别ID 药品收发记录.入出类别ID%TYPE ;
    V_产地 药品收发记录.产地%TYPE ; 
    V_批次 药品收发记录.批次%TYPE ; 
    V_批号 药品收发记录.批号%TYPE ; 
    V_效期 药品收发记录.效期%TYPE ; 
    V_成本价 药品收发记录.成本价%TYPE ; 
    V_成本金额 药品收发记录.成本金额%TYPE ; 
    V_扣率 药品收发记录.扣率%TYPE ; 
    V_零售价 药品收发记录.零售价%TYPE ; 
    V_零售金额 药品收发记录.零售金额%TYPE ; 
    V_差价 药品收发记录.差价%TYPE ; 
    V_摘要 药品收发记录.摘要%TYPE ; 
    V_剩余数量 药品收发记录.实际数量%TYPE;
    V_剩余成本金额 药品收发记录.成本金额%Type;
    V_剩余零售金额 药品收发记录.零售金额%Type; 
    V_入出系数 药品收发记录.入出系数%TYPE; 
	V_外调价 药品收发记录.单量%TYPE;
	V_外调单位 药品收发记录.发药窗口%TYPE;
    v_批准文号 药品收发记录.批准文号%TYPE;

    V_记录数 NUMBER; 
    V_收发ID 药品收发记录.ID%TYPE; 
	INTDIGIT NUMBER;
BEGIN
	--获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';

    IF 行次_IN =1 THEN
        UPDATE 药品收发记录 
        SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3) 
        WHERE NO = NO_IN AND 单据 = 11 AND 记录状态 =原记录状态_IN ; 
        IF SQL%ROWCOUNT = 0 THEN 
            RAISE ERR_ISSTRIKED; 
        END IF; 
    END IF;
    
    --主要针对原不分批现在分批的药品，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount 
    FROM 药品收发记录 a,药品规格 b
    WHERE a.药品id=b.药品id
        AND a.no=NO_IN 
        AND a.单据=11
        AND A.药品ID+0=药品ID_IN
        AND MOD(a.记录状态,3)=0
        AND nvl(a.批次,0)=0
        AND ((NVL(B.药库分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')))
            OR NVL(B.药房分批,0)=1);
    
    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  
    
    SELECT SUM(实际数量) AS 剩余数量,SUM(成本金额) AS 剩余成本金额,SUM(零售金额) AS 剩余零售金额,库房ID,入出类别ID,入出系数,批次,产地,批号,效期,成本价,扣率,零售价,摘要,单量,发药窗口,批准文号
    INTO  V_剩余数量,V_剩余成本金额,V_剩余零售金额,V_库房ID,V_入出类别ID,V_入出系数,V_批次,V_产地,V_批号,V_效期,V_成本价,V_扣率,V_零售价,V_摘要,V_外调价,V_外调单位,v_批准文号
    FROM 药品收发记录 
    WHERE NO=NO_IN 
    AND 单据=11 
    AND 药品ID=药品ID_IN 
    AND 序号=序号_IN
    GROUP BY 库房ID,入出类别ID,入出系数,批次,产地,批号,效期,成本价,扣率,零售价,摘要,单量,发药窗口,批准文号;

    --冲销数量大于剩余数量，不允许
    IF abs(V_剩余数量)<abs(冲销数量_IN) THEN
        RAISE ERR_ISNONUM; 
    END IF;

    V_成本金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余成本金额,INTDIGIT);
    V_零售金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余零售金额,INTDIGIT);
    V_差价:=V_零售金额-V_成本金额;

    SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;
    Insert INTO 药品收发记录
    (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,
    药品ID,批次,产地,批号,效期,填写数量,实际数量,成本价,
    成本金额,零售价,零售金额,差价,单量,发药窗口,
	摘要,填制人,填制日期,审核人,审核日期,批准文号)
    VALUES 
    (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),11,NO_IN,序号_IN,V_库房ID,V_入出类别ID,V_入出系数,
    药品ID_IN,V_批次,V_产地,V_批号,V_效期,-冲销数量_IN,-冲销数量_IN,V_成本价,-V_成本金额,V_零售价,-V_零售金额,
    -V_差价,V_外调价,V_外调单位,V_摘要,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN,v_批准文号);

    --原分批现不分批的药品,在冲消时，要处理他
    BEGIN 
        SELECT COUNT(*) INTO V_记录数
        FROM 药品收发记录 A, 药品规格 B
        WHERE A.药品ID+0=B.药品ID 
        AND B.药品ID=药品ID_IN
        AND A.NO=NO_IN
        AND A.单据 = 11 
        AND MOD(A.记录状态,3)=0
        AND NVL(A.批次,0)>0
        AND (NVL(B.药库分批,0)=0 OR 
            (NVL(B.药房分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室'))));
    EXCEPTION 
        WHEN OTHERS THEN
            V_记录数:=0;
    END;
    IF V_记录数>0 THEN
        V_批次:=0;
    ELSE
        V_批次:=NVL (V_批次, 0);
    END IF;
    --更改药品库存表的相应数据
    UPDATE 药品库存
        SET 可用数量 = NVL (可用数量, 0) + NVL (冲销数量_IN, 0),
             实际数量 = NVL (实际数量, 0) + NVL (冲销数量_IN, 0),
             实际金额 = NVL (实际金额, 0) + NVL (V_零售金额, 0),
             实际差价 = NVL (实际差价, 0) + NVL (V_差价, 0)
     WHERE 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND NVL (批次, 0) = v_批次
        AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品库存
        (库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次批号,效期,上次产地,批准文号)
        VALUES 
        (V_库房ID,药品ID_IN,V_批次,1,冲销数量_IN,冲销数量_IN,V_零售金额,V_差价,V_批号,V_效期,v_产地,v_批准文号);
    END IF;

    DELETE
      FROM 药品库存
     WHERE 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND NVL (可用数量, 0) = 0
        AND NVL (实际数量, 0) = 0
        AND NVL (实际金额, 0) = 0
        AND NVL (实际差价, 0) = 0;

    --更改药品收发汇总表的相应数据

    UPDATE 药品收发汇总
        SET 数量 = NVL (数量, 0) + NVL (冲销数量_IN, 0),
             金额 = NVL (金额, 0) + NVL (V_零售金额, 0),
             差价 = NVL (差价, 0) + NVL (V_差价, 0)
     WHERE 日期 = TRUNC (填制日期_IN)
        AND 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND 类别ID = V_入出类别ID
        AND 单据 = 11;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品收发汇总
        (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
        VALUES 
        (TRUNC (填制日期_IN),V_库房ID,药品ID_IN,V_入出类别ID,11,冲销数量_IN,V_零售金额,V_差价);
    END IF;
EXCEPTION
    WHEN Err_isstriked THEN
        Raise_application_error (-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
    WHEN Err_isBatch THEN
        Raise_application_error (-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品，不能冲销！[ZLSOFT]'); 
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品其他出库_strike;
/

CREATE OR REPLACE PROCEDURE zl_药品盘点_Insert (
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    批次_IN IN 药品收发记录.批次%TYPE,
    入出类别ID_IN IN 药品收发记录.入出类别ID%TYPE,
    --每次都使用，所以用外面传入比较好
    入出系数_IN IN 药品收发记录.入出系数%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    帐面数量_IN IN 药品收发记录.填写数量%TYPE,
    实盘数量_IN IN 药品收发记录.扣率%TYPE,
    数量差_IN IN 药品收发记录.实际数量%TYPE,
    售价_IN IN 药品收发记录.零售价%TYPE,
    金额差_IN IN 药品收发记录.零售金额%TYPE,
    差价差_IN IN 药品收发记录.差价%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    填制日期_IN IN 药品收发记录.填制日期%TYPE,
    摘要_IN IN 药品收发记录.摘要%TYPE := NULL,
    产地_IN IN 药品收发记录.产地%TYPE := NULL,
    批号_IN IN 药品收发记录.批号%TYPE := NULL,
    效期_IN IN 药品收发记录.效期%TYPE := NULL,
    盘点时间_IN IN 药品收发记录.频次%TYPE := NULL,
    库存金额_IN IN 药品收发记录.成本价%TYPE := NULL,
    库存差价_IN IN 药品收发记录.成本金额%TYPE := NULL,
    成本价_IN In 药品收发记录.单量%TYPE := NULL,
    批准文号_IN In 药品收发记录.批准文号%TYPE := Null
)
IS
    V_lngID NUMBER(18);
BEGIN
    --如果批次_IN为-1,则表示新产生一个批次药品
    SELECT 药品收发记录_ID.Nextval
    INTO V_lngID
    FROM Dual;

    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,药品ID,批次,
        产地,批号,效期,填写数量,扣率,实际数量,零售价,零售金额,差价,
        摘要,填制人,填制日期,频次,成本价,成本金额,单量,批准文号)
    VALUES (
        V_lngID,1,12,NO_IN,序号_IN,库房ID_IN,入出类别ID_IN,入出系数_IN,药品ID_IN,
        DECODE(批次_IN,-1,V_lngID,批次_IN),产地_IN,批号_IN,效期_IN,帐面数量_IN,
        实盘数量_IN,数量差_IN,售价_IN,金额差_IN,差价差_IN,摘要_IN,填制人_IN,
        填制日期_IN,盘点时间_IN,库存金额_IN,库存差价_IN,成本价_IN,批准文号_IN);

    IF 入出系数_IN = -1 THEN
        --同时更新库存数
        UPDATE 药品库存
        SET 可用数量 = NVL (可用数量, 0) - 数量差_IN
        WHERE 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND NVL (批次, 0) = NVL (批次_IN, 0)
        AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
            (库房ID, 药品ID, 性质, 批次, 可用数量,上次批号,效期,上次产地,批准文号)
            VALUES 
            (库房ID_IN, 药品ID_IN, 1, 批次_IN, -数量差_IN,批号_IN,效期_IN,产地_IN,批准文号_IN);
        END IF;

        DELETE
        FROM 药品库存
        WHERE 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND NVL (可用数量, 0) = 0
        AND NVL (实际数量, 0) = 0
        AND NVL (实际金额, 0) = 0
        AND NVL (实际差价, 0) = 0;
    END IF;
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品盘点_Insert;
/




CREATE OR REPLACE PROCEDURE zl_药品盘点_DELETE (
    
    --删除药品收发记录及恢复相应的表：药品库存
    NO_IN IN 药品收发记录.NO%TYPE
)
IS
    Err_isverified EXCEPTION;

    CURSOR C_药品收发记录
    IS
        SELECT 实际数量, 库房ID, 批次, 药品ID,批号,效期,产地,批准文号
          FROM 药品收发记录
         WHERE NO = NO_IN
            AND 单据 = 12
            AND 入出系数 = -1
         ORDER BY 药品ID;
BEGIN
    --通过循环，恢复出库类别原来的可用数量，
     --实际数量保存的是数量差
    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        UPDATE 药品库存
            SET 可用数量 = NVL (可用数量, 0) + V_药品收发记录.实际数量
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次, 0) = NVL (V_药品收发记录.批次, 0)
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
                            (库房ID, 药品ID, 批次, 性质, 可用数量,上次批号,效期,上次产地,批准文号)
                  VALUES (
                      V_药品收发记录.库房ID,V_药品收发记录.药品ID,V_药品收发记录.批次,1,
                      V_药品收发记录.实际数量,V_药品收发记录.批号,V_药品收发记录.效期,V_药品收发记录.产地,V_药品收发记录.批准文号
                  );
        END IF;

        DELETE
          FROM 药品库存
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (可用数量, 0) = 0
            AND NVL (实际数量, 0) = 0
            AND NVL (实际金额, 0) = 0
            AND NVL (实际差价, 0) = 0;
    END LOOP;

    DELETE
      FROM 药品收发记录
     WHERE NO = NO_IN
        AND 单据 = 12
        AND 记录状态 = 1
        AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品盘点_DELETE;
/


CREATE OR REPLACE PROCEDURE zl_药品盘点_verify (
    NO_IN IN 药品收发记录.NO%TYPE := NULL,
    审核人_IN IN 药品收发记录.审核人%TYPE := NULL
)
IS
    Err_isverified EXCEPTION;
    Err_isBatch exception;
    v_BatchCount integer;    --原不分批现在分批的药品的数量

    CURSOR C_药品收发记录
    IS
        SELECT ID, 实际数量, 零售金额, 差价, 库房ID, 药品ID, 批次, 批号, 效期,
                 产地, 入出类别ID, 入出系数,批准文号
          FROM 药品收发记录
         WHERE NO = NO_IN
            AND 单据 = 12
            AND 记录状态 = 1
         ORDER BY 药品ID;
BEGIN
    UPDATE 药品收发记录
        SET 审核人 = 审核人_IN,
             审核日期 = SYSDATE
     WHERE NO = NO_IN
        AND 单据 = 12
        AND 记录状态 = 1
        AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;

    --主要针对原不分批现在分批的药品，不能对其审核
    SELECT count(*) INTO v_BatchCount FROM        
            药品收发记录 a,药品规格 b
    WHERE a.药品id=b.药品id
        AND a.no=NO_IN 
        AND a.单据=12
        AND a.记录状态=1
        AND nvl(a.批次,0)=0
        AND nvl(b.药库分批,0)=1
        AND a.库房id not in
        (select 部门id from  部门性质说明 where (工作性质 LIKE '%药房') or (工作性质 LIKE '制剂室'));
    IF v_batchcount>0 THEN
        raise Err_isBatch;
    END IF;  
    
    --原分批现不分批的药品,在审核时，要处理他
    UPDATE 药品收发记录 SET 批次=0
    WHERE 
    id=
    (SELECT id
        FROM 药品收发记录 a, 药品规格 b
        WHERE b.药品id=a.药品ID
        AND a.no=no_in
        AND a.单据 = 12 
            AND a.记录状态 = 1 
            AND nvl(a.批次,0)>0
            AND nvl(b.药库分批,0)=0
            );
    
    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        --更改药品库存表的相应数据
        UPDATE 药品库存
            SET 可用数量 =
                     NVL (可用数量, 0) +
                         DECODE (
                             V_药品收发记录.入出系数, 1,
                             NVL (V_药品收发记录.实际数量, 0), 0
                         ),
                 实际数量 =
                     NVL (实际数量, 0) +
                         NVL (V_药品收发记录.实际数量, 0) * V_药品收发记录.入出系数,
                 实际金额 =
                     NVL (实际金额, 0) +
                         NVL (V_药品收发记录.零售金额, 0) * V_药品收发记录.入出系数,
                 实际差价 =
                     NVL (实际差价, 0) +
                         NVL (V_药品收发记录.差价, 0) * V_药品收发记录.入出系数
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次, 0) = NVL (V_药品收发记录.批次, 0)
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
                            (
                                库房ID,
                                药品ID,
                                批次,
                                性质,
                                可用数量,
                                实际数量,
                                实际金额,
                                实际差价,
                                上次批号,
                                上次产地,
                                效期,
                                批准文号
                            )
                  VALUES (
                      V_药品收发记录.库房ID,
                      V_药品收发记录.药品ID,
                      V_药品收发记录.批次,
                      1,
                      DECODE (
                          V_药品收发记录.入出系数, 1, NVL (
                                                                    V_药品收发记录.实际数量, 0
                                                                ), 0
                      ),
                      V_药品收发记录.实际数量 * V_药品收发记录.入出系数,
                      V_药品收发记录.零售金额 * V_药品收发记录.入出系数,
                      V_药品收发记录.差价 * V_药品收发记录.入出系数,
                      V_药品收发记录.批号,
                      V_药品收发记录.产地,
                      V_药品收发记录.效期,
                      V_药品收发记录.批准文号
                  );
        END IF;

        DELETE
          FROM 药品库存
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (可用数量, 0) = 0
            AND NVL (实际数量, 0) = 0
            AND NVL (实际金额, 0) = 0
            AND NVL (实际差价, 0) = 0;

        --更改药品收发汇总表的相应数据

        UPDATE 药品收发汇总
            SET 数量 =
                     NVL (数量, 0) +
                         NVL (V_药品收发记录.实际数量, 0) * V_药品收发记录.入出系数,
                 金额 =
                     NVL (金额, 0) +
                         NVL (V_药品收发记录.零售金额, 0) * V_药品收发记录.入出系数,
                 差价 =
                     NVL (差价, 0) +
                         NVL (V_药品收发记录.差价, 0) * V_药品收发记录.入出系数
         WHERE 日期 = TRUNC (SYSDATE)
            AND 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND 类别ID = V_药品收发记录.入出类别ID
            AND 单据 = 12;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品收发汇总
                (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
                  VALUES (
                      TRUNC (SYSDATE),
                      V_药品收发记录.库房ID,
                      V_药品收发记录.药品ID,
                      V_药品收发记录.入出类别ID,
                      12,
                      V_药品收发记录.实际数量 * V_药品收发记录.入出系数,
                      V_药品收发记录.零售金额 * V_药品收发记录.入出系数,
                      V_药品收发记录.差价 * V_药品收发记录.入出系数
                  );
        END IF;
    END LOOP;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]'
        );
    WHEN Err_isBatch THEN
        Raise_application_error ( 
            -20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品，不能审核！[ZLSOFT]' 
        ); 
        
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品盘点_verify;
/


CREATE OR REPLACE PROCEDURE zl_药品盘点_strike (
   NO_IN IN 药品收发记录.NO%TYPE,
   审核人_IN IN 药品收发记录.审核人%TYPE
)
IS
   Err_isstriked EXCEPTION;
   Err_isBatch exception;
   v_BatchCount integer;    --原不分批现在分批的药品的数量
   V_COUNT INTEGER;    --原分批现不分批
   V_批次 药品收发记录.批次%TYPE;

   CURSOR C_药品收发记录
   IS
      SELECT ID, 实际数量, 零售金额, 差价, 库房ID, 药品ID, 批次, 批号,效期, 产地,
             入出类别ID, 入出系数,单量,批准文号
        FROM 药品收发记录
       WHERE NO = NO_IN
         AND 单据 = 12
         AND 记录状态 = 2
       ORDER BY 药品ID;
BEGIN
   UPDATE 药品收发记录
      SET 记录状态 = 3
    WHERE NO = NO_IN
      AND 单据 = 12
      AND 记录状态 = 1;

   IF SQL%ROWCOUNT = 0 THEN
      RAISE Err_isstriked;
   END IF;
   
   --主要针对原不分批现在分批的药品，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount FROM        
            药品收发记录 a,药品规格 b
    WHERE a.药品id=b.药品id
        AND a.no=NO_IN 
        AND a.单据=12
        AND a.记录状态=3
        AND nvl(a.批次,0)=0
        AND ((NVL(B.药库分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')))
            OR NVL(B.药房分批,0)=1);
        
    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  

   Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,
         药品ID,批次,产地,批号,效期,填写数量,扣率,实际数量,
         成本价,成本金额,零售价,零售金额,差价,摘要,
         填制人,填制日期,审核人,审核日期,频次,单量,批准文号
         )
      SELECT 药品收发记录_ID.Nextval, 2, 单据, NO, 序号, 库房ID, 入出类别ID,
             入出系数, a.药品ID, 
             DECODE(NVL(a.批次,0),0,NULL,(DECODE(NVL(b.药库分批,0),0,NULL,a.批次))), 
             a.产地, 批号, 效期, 填写数量, a.扣率,
             -实际数量, a.成本价, 成本金额, 零售价, -零售金额, -差价, 摘要,
             审核人_IN, SYSDATE, 审核人_IN, SYSDATE, 频次,单量,a.批准文号
        FROM 药品收发记录 a,药品规格 b
       WHERE NO = NO_IN
         AND a.药品id=b.药品id
         AND 单据 = 12
         AND 记录状态 = 3;

   FOR V_药品收发记录 IN C_药品收发记录 LOOP
    --原分批现不分批的药品,在C冲消时，要处理他
    BEGIN 
        SELECT COUNT(*) INTO V_COUNT
        FROM 药品收发记录 A, 药品规格 B
        WHERE B.药品ID=V_药品收发记录.药品ID
        AND A.NO=NO_IN
        AND A.单据 = 12
        and a.库房id+0=V_药品收发记录.库房id
        AND A.记录状态 = 3 
        AND NVL(A.批次,0)>0
        AND (NVL(B.药库分批,0)=0 OR 
            (NVL(B.药房分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室'))))
        ;
    EXCEPTION 
        WHEN OTHERS THEN
            V_COUNT:=0;
    END;
    IF V_COUNT>0 THEN
        V_批次:=0;
    ELSE
        V_批次:=NVL (V_药品收发记录.批次, 0);
    END IF;
      --更改药品库存表的相应数据
      
        UPDATE 药品库存
            SET 可用数量=NVL(可用数量,0)+NVL(V_药品收发记录.实际数量,0)*V_药品收发记录.入出系数,
                实际数量=NVL(实际数量,0)+NVL(V_药品收发记录.实际数量,0)*V_药品收发记录.入出系数,
                实际金额=NVL(实际金额,0)+NVL(V_药品收发记录.零售金额,0)*V_药品收发记录.入出系数,
                实际差价=NVL(实际差价,0)+NVL(V_药品收发记录.差价,0)*V_药品收发记录.入出系数
          WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次, 0) = V_批次
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
                (库房ID,药品ID,批次,性质,可用数量,实际数量,
                 实际金额,实际差价,上次批号,上次产地,效期,上次采购价,批准文号
                 )
                VALUES (
                    V_药品收发记录.库房ID,
                    V_药品收发记录.药品ID,
                    V_批次,
                    1,
                    V_药品收发记录.实际数量*V_药品收发记录.入出系数,
                    V_药品收发记录.实际数量*V_药品收发记录.入出系数,
                    V_药品收发记录.零售金额*V_药品收发记录.入出系数,
                    V_药品收发记录.差价*V_药品收发记录.入出系数,
                    V_药品收发记录.批号,
                    V_药品收发记录.产地,
                    V_药品收发记录.效期,
                    V_药品收发记录.单量,
                    V_药品收发记录.批准文号
                 );
        END IF;
      

        DELETE 
          FROM 药品库存
         WHERE 库房ID = V_药品收发记录.库房ID
           AND 药品ID = V_药品收发记录.药品ID
           AND nvl(可用数量,0)=0 
           And nvl(实际数量,0)=0 
           And nvl(实际金额,0)=0 
           And nvl(实际差价,0)=0;

      --更改药品收发汇总表的相应数据

       UPDATE 药品收发汇总
          SET 数量=NVL(数量,0)+NVL(V_药品收发记录.实际数量,0)*V_药品收发记录.入出系数,
              金额=NVL(金额,0)+NVL(V_药品收发记录.零售金额,0)*V_药品收发记录.入出系数,
              差价=NVL(差价,0)+NVL(V_药品收发记录.差价,0)*V_药品收发记录.入出系数
        WHERE 日期 = TRUNC (SYSDATE)
          AND 库房ID = V_药品收发记录.库房ID
          AND 药品ID = V_药品收发记录.药品ID
          AND 类别ID = V_药品收发记录.入出类别ID
          AND 单据 = 12;

      IF SQL%NOTFOUND THEN
         Insert INTO 药品收发汇总
            (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
              VALUES (
                 TRUNC (SYSDATE),
                 V_药品收发记录.库房ID,
                 V_药品收发记录.药品ID,
                 V_药品收发记录.入出类别ID,
                 12,
                 V_药品收发记录.实际数量*V_药品收发记录.入出系数,
                 V_药品收发记录.零售金额*V_药品收发记录.入出系数,
                 V_药品收发记录.差价*V_药品收发记录.入出系数
              );
      END IF;
   END LOOP;
EXCEPTION
   WHEN Err_isstriked THEN
      Raise_application_error (
         -20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]'
      );
   WHEN Err_isBatch THEN
      Raise_application_error ( 
            -20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品，不能冲销！[ZLSOFT]' 
        ); 
   WHEN OTHERS THEN
      zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品盘点_strike;
/

CREATE OR REPLACE PROCEDURE zl_药品盘点记录单_Insert (
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    批次_IN IN 药品收发记录.批次%TYPE,
    入出类别ID_IN IN 药品收发记录.入出类别ID%TYPE,
    --每次都使用，所以用外面传入比较好
    入出系数_IN IN 药品收发记录.入出系数%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    帐面数量_IN IN 药品收发记录.填写数量%TYPE,
    实盘数量_IN IN 药品收发记录.扣率%TYPE,
    数量差_IN IN 药品收发记录.实际数量%TYPE,
    售价_IN IN 药品收发记录.零售价%TYPE,
    金额差_IN IN 药品收发记录.零售金额%TYPE,
    差价差_IN IN 药品收发记录.差价%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    填制日期_IN IN 药品收发记录.填制日期%TYPE,
    摘要_IN IN 药品收发记录.摘要%TYPE := NULL,
    产地_IN IN 药品收发记录.产地%TYPE := NULL,
    批号_IN IN 药品收发记录.批号%TYPE := NULL,
    效期_IN IN 药品收发记录.效期%TYPE := NULL,
    盘点时间_IN IN 药品收发记录.频次%TYPE := NULL,
    库存金额_IN IN 药品收发记录.成本价%TYPE := NULL,
    库存差价_IN IN 药品收发记录.成本金额%TYPE := NULL,
    批准文号_IN In 药品收发记录.批准文号%TYPE := Null
)
IS
BEGIN
    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,
        填写数量,扣率,实际数量,零售价,零售金额,差价,摘要,填制人,填制日期,频次,成本价,成本金额,批准文号)
    VALUES (
        药品收发记录_ID.Nextval,1,14,NO_IN,序号_IN,库房ID_IN,入出类别ID_IN,入出系数_IN,药品ID_IN,
        批次_IN,产地_IN,批号_IN,效期_IN,帐面数量_IN,实盘数量_IN,数量差_IN,售价_IN,金额差_IN,
        差价差_IN,摘要_IN,填制人_IN,填制日期_IN,盘点时间_IN,库存金额_IN,库存差价_IN,批准文号_IN);
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品盘点记录单_Insert;
/

CREATE OR REPLACE PROCEDURE zl_药品申领_Insert (
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    对方部门ID_IN IN 药品收发记录.对方部门ID%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    批次_IN IN 药品收发记录.批次%TYPE,
    填写数量_IN IN 药品收发记录.填写数量%TYPE,
    实际数量_IN in 药品收发记录.实际数量%TYPE,
    成本价_IN IN 药品收发记录.成本价%TYPE,
    成本金额_IN IN 药品收发记录.成本金额%TYPE,
    零售价_IN IN 药品收发记录.零售价%TYPE,
    零售金额_IN IN 药品收发记录.零售金额%TYPE,
    差价_IN IN 药品收发记录.差价%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    产地_IN IN 药品收发记录.产地%TYPE := NULL,
    批号_IN IN 药品收发记录.批号%TYPE := NULL,
    效期_IN IN 药品收发记录.效期%TYPE := NULL,
    摘要_IN IN 药品收发记录.摘要%TYPE := NULL,
    填制日期_IN IN 药品收发记录.填制日期%TYPE := NULL,
    上次供应商ID_IN In 药品收发记录.供药单位ID%TYPE := Null,
    批准文号_IN In 药品收发记录.批准文号%TYPE := Null
)
IS
    V_lngID 药品收发记录.ID%TYPE;--收发ID
    V_入的类别ID 药品收发记录.入出类别ID%TYPE;--入出类别ID
    V_出的类别ID 药品收发记录.入出类别ID%TYPE;--入出类别ID
    V_编码 收费项目目录.编码%TYPE;
    V_下可用库存 系统参数表.参数值%Type;

	ERR_MutilROW EXCEPTION ;
	intRecords NUMBER ;
BEGIN
    --首先找出入和出的类别ID
    SELECT B.ID INTO V_入的类别ID
    FROM 药品单据性质 A, 药品入出类别 B
    WHERE A.类别ID = B.ID AND A.单据 = 6 AND B.系数 = 1 AND ROWNUM < 2;
    SELECT B.ID INTO V_出的类别ID
    FROM 药品单据性质 A, 药品入出类别 B
    WHERE A.类别ID = B.ID AND A.单据 = 6 AND B.系数 = -1 AND ROWNUM < 2;

	SELECT 药品收发记录_ID.Nextval INTO V_lngID FROM Dual;

    --插入类别为出的那一笔
    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,
        填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,发药方式,供药单位ID,批准文号)
    VALUES (药品收发记录_ID.Nextval,1,6,NO_IN,序号_IN,库房ID_IN,对方部门ID_IN,
		V_出的类别ID,-1,药品ID_IN,批次_IN,产地_IN,批号_IN,效期_IN,
		填写数量_IN,实际数量_IN,成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,
		差价_IN,摘要_IN,填制人_IN,填制日期_IN,1,上次供应商ID_IN,批准文号_IN);

    --插入类别为入的那一笔
    Insert INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,
        填写数量,实际数量,成本价,成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,发药方式,供药单位ID,批准文号)
          VALUES (V_lngID,1,6,NO_IN,序号_IN + 1,对方部门ID_IN,库房ID_IN,V_入的类别ID,
              1,药品ID_IN,批次_IN,产地_IN,批号_IN,效期_IN,填写数量_IN,实际数量_IN,
              成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,差价_IN,摘要_IN,
              填制人_IN,填制日期_IN,1,上次供应商ID_IN,批准文号_IN);

	--检查是否存在相同药品相同批次的数据，如果存在不允许保存
	SELECT COUNT(*) INTO intRecords 
	FROM 药品收发记录
	WHERE 单据=6 AND NO=NO_IN AND 入出系数=-1 AND 药品ID+0=药品ID_IN AND Nvl(批次,0)=NVL(批次_IN,0);
	IF intRecords>1 THEN 
		RAISE ERR_MutilROW;
	END IF ;
  
	--根据参数决定是否下发药库房的可用库存
	Select 参数值 Into v_下可用库存 From 系统参数表 Where 参数号 = 96;
  
  --参数为1表示在填单时下可用数量
	If v_下可用库存 = '1' Then
      UPDATE 药品库存
			SET 可用数量=NVL(可用数量,0)-实际数量_IN
			WHERE 库房ID=库房ID_IN AND 药品ID=药品ID_IN AND NVL(批次,0)=批次_IN AND 性质=1;
      
			IF SQL%ROWCOUNT=0 THEN 
				INSERT INTO 药品库存(库房ID,药品ID,批次,性质,可用数量,上次批号,效期,上次产地)
				VALUES (库房ID_IN,药品ID_IN,批次_IN,1,-1*实际数量_IN,
					批号_IN,效期_IN,产地_IN);
			END IF ;

			DELETE
			FROM 药品库存
			WHERE 库房ID = 库房ID_IN
			AND 药品ID = 药品ID_IN
			AND NVL(可用数量,0)=0
			AND NVL(实际数量,0)=0
			AND NVL(实际金额,0)=0
			AND NVL(实际差价,0)=0;
  End If;
  
EXCEPTION
	WHEN ERR_MutilROW THEN 
		SELECT 编码 INTO V_编码 FROM 收费项目目录 WHERE ID = 药品ID_IN;
		RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]编码为'||V_编码||'的药品，存在多条重复的记录，请合并为一条记录！[ZLSOFT]');
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品申领_Insert;
/


CREATE OR REPLACE PROCEDURE zl_药品申领_DELETE (
    NO_IN IN 药品收发记录.NO%TYPE
)
Is
    V_下可用库存 系统参数表.参数值%Type;
    
    CURSOR C_药品收发记录
    IS
		SELECT 实际数量,库房ID,批次,药品ID,批号,效期,产地,供药单位ID,批准文号
		FROM 药品收发记录
		WHERE NO = NO_IN
		AND 单据 = 6
		AND 入出系数 = -1
		ORDER BY 药品ID;
BEGIN
    --根据参数决定是否恢复发药库房的可用库存
    Select 参数值 Into v_下可用库存 From 系统参数表 Where 参数号 = 96;
    
    --参数为1表示需要恢复可用数量
    IF  v_下可用库存='1' THEN
      FOR V_药品收发记录 IN C_药品收发记录 LOOP
  			UPDATE 药品库存
  			SET 可用数量 = NVL (可用数量, 0) + V_药品收发记录.实际数量
  			WHERE 库房ID = V_药品收发记录.库房ID
  			AND 药品ID = V_药品收发记录.药品ID
  			AND NVL (批次, 0) = NVL (V_药品收发记录.批次, 0)
  			AND 性质 = 1;
  
  			IF SQL%NOTFOUND THEN
  				INSERT INTO 药品库存
  					(库房ID, 药品ID, 批次, 性质, 可用数量,上次批号,效期,上次产地,上次供应商ID,批准文号)
  				VALUES (
  					V_药品收发记录.库房ID,V_药品收发记录.药品ID,V_药品收发记录.批次,1,
  					V_药品收发记录.实际数量,V_药品收发记录.批号,V_药品收发记录.效期,V_药品收发记录.产地,V_药品收发记录.供药单位ID,V_药品收发记录.批准文号
  					);
  			END IF;
  
  			DELETE
  			FROM 药品库存
  			WHERE 库房ID = V_药品收发记录.库房ID
  			AND 药品ID = V_药品收发记录.药品ID
  			AND NVL(可用数量,0) = 0
  			AND NVL(实际数量,0) = 0
  			AND NVL(实际金额,0) = 0
  			AND NVL(实际差价,0) = 0;
  		END LOOP;
	  END IF ;
    
    DELETE 药品收发记录
    WHERE NO = NO_IN AND 单据 = 6 AND 记录状态 = 1 AND 审核人 IS NULL;
    
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品申领_DELETE;
/

CREATE OR REPLACE PROCEDURE zl_药品领用_Insert (
    入出类别ID_IN IN 药品收发记录.入出类别ID%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    对方部门ID_IN IN 药品收发记录.对方部门ID%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    批次_IN IN 药品收发记录.批次%TYPE,
    填写数量_IN IN 药品收发记录.填写数量%TYPE,
    成本价_IN IN 药品收发记录.成本价%TYPE,
    成本金额_IN IN 药品收发记录.成本金额%TYPE,
    零售价_IN IN 药品收发记录.零售价%TYPE,
    零售金额_IN IN 药品收发记录.零售金额%TYPE,
    差价_IN IN 药品收发记录.差价%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    填制日期_IN IN 药品收发记录.填制日期%TYPE,
    产地_IN IN 药品收发记录.产地%TYPE := NULL,
    批号_IN IN 药品收发记录.批号%TYPE := NULL,
    效期_IN IN 药品收发记录.效期%TYPE := NULL,
    摘要_IN IN 药品收发记录.摘要%TYPE := Null,
    领用人_IN IN 药品收发记录.领用人%TYPE := Null,
    批准文号_IN In 药品收发记录.批准文号%TYPE := Null
)
IS
    V_入出系数 药品收发记录.入出系数%TYPE;--收发ID
    Err_isNOnumber EXCEPTION;
    V_编码 收费项目目录.编码%TYPE;
    V_可用数量 药品库存.可用数量%TYPE;
BEGIN
    V_入出系数 := -1;

    IF 批次_IN > 0 THEN
        BEGIN
            SELECT 可用数量
              INTO V_可用数量
              FROM 药品库存
             WHERE 药品ID = 药品ID_IN
                AND NVL (批次, 0) = 批次_IN
                AND 库房ID = 库房ID_IN
                AND 性质 = 1
                AND ROWNUM = 1;
        EXCEPTION
            WHEN OTHERS THEN
                V_可用数量 := 0;
        END;

        IF V_可用数量 - 填写数量_IN < 0 THEN
            RAISE Err_isNOnumber;
        END IF;
    END IF;

    --插入类别为出的那一笔
    Insert INTO 药品收发记录
                    (
                        ID,
                        记录状态,
                        单据,
                        NO,
                        序号,
                        库房ID,
                        对方部门ID,
                        入出类别ID,
                        入出系数,
                        药品ID,
                        批次,
                        产地,
                        批号,
                        效期,
                        填写数量,
                        实际数量,
                        成本价,
                        成本金额,
                        零售价,
                        零售金额,
                        差价,
                        摘要,
                        填制人,
                        填制日期,
                        领用人,
                        批准文号
                    )
          VALUES (
              药品收发记录_ID.Nextval,
              1,
              7,
              NO_IN,
              序号_IN,
              库房ID_IN,
              对方部门ID_IN,
              入出类别ID_IN,
              V_入出系数,
              药品ID_IN,
              批次_IN,
              产地_IN,
              批号_IN,
              效期_IN,
              填写数量_IN,
              填写数量_IN,
              成本价_IN,
              成本金额_IN,
              零售价_IN,
              零售金额_IN,
              差价_IN,
              摘要_IN,
              填制人_IN,
              填制日期_IN,
              领用人_IN,
              批准文号_IN
          );

    --同时更新库存数
    UPDATE 药品库存
        SET 可用数量 = NVL (可用数量, 0) - 填写数量_IN
     WHERE 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND NVL (批次, 0) = NVL (批次_IN, 0)
        AND 性质 = 1;

    --不插入批次是因为批次药品不够，不准出库
    IF SQL%NOTFOUND THEN
        Insert INTO 药品库存
                        (库房ID, 药品ID, 性质, 可用数量,上次批号,效期,上次产地,批准文号)
              VALUES (库房ID_IN, 药品ID_IN, 1, -填写数量_IN,批号_IN,效期_IN,产地_IN,批准文号_IN);
    END IF;

    DELETE
      FROM 药品库存
     WHERE 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND nvl(可用数量,0) = 0
        AND nvl(实际数量,0) = 0
        AND nvl(实际金额,0) = 0
        AND nvl(实际差价,0) = 0;
EXCEPTION
    WHEN Err_isNOnumber THEN
        SELECT 编码
          INTO V_编码
          FROM 收费项目目录
         WHERE ID = 药品ID_IN;
        Raise_application_error (
            -20101, '[ZLSOFT]编码为' || V_编码 || ',批号为' || 批号_IN ||
                          '的分批核算药品' ||
                          CHR (10) ||
                          CHR (13) ||
                          '可用库存数量不够！[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品领用_Insert;
/


CREATE OR REPLACE PROCEDURE zl_药品领用_DELETE (
    
    --删除药品收发记录及恢复相应的表：药品库存
    NO_IN IN 药品收发记录.NO%TYPE
)
IS
    Err_isverified EXCEPTION;

    CURSOR C_药品收发记录
    IS
        SELECT 填写数量, 库房ID, 批次, 药品ID,批号,效期,产地,批准文号
          FROM 药品收发记录
         WHERE NO = NO_IN
            AND 单据 = 7
            AND 入出系数 = -1
         ORDER BY 药品ID;
BEGIN
    --通过循环，恢复原来的可用数量
    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        UPDATE 药品库存
            SET 可用数量 = NVL (可用数量, 0) + V_药品收发记录.填写数量
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次, 0) = NVL (V_药品收发记录.批次, 0)
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
                            (库房ID, 药品ID, 批次, 性质, 可用数量,上次批号,效期,上次产地,批准文号)
                  VALUES (
                      V_药品收发记录.库房ID,V_药品收发记录.药品ID,V_药品收发记录.批次,1,
                      V_药品收发记录.填写数量,V_药品收发记录.批号,V_药品收发记录.效期,V_药品收发记录.产地,V_药品收发记录.批准文号
                  );
        END IF;

        DELETE
          FROM 药品库存
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (可用数量, 0) = 0
            AND NVL (实际数量, 0) = 0
            AND NVL (实际金额, 0) = 0
            AND NVL (实际差价, 0) = 0;
    END LOOP;

    DELETE
      FROM 药品收发记录
     WHERE NO = NO_IN
        AND 单据 = 7
        AND 记录状态 = 1
        AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品领用_DELETE;
/


CREATE OR REPLACE PROCEDURE zl_药品领用_verify (
    序号_IN IN 药品收发记录.序号%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    对方部门ID_IN IN 药品收发记录.对方部门ID%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    产地_IN IN 药品收发记录.产地%TYPE,
    批次_IN IN 药品收发记录.批次%TYPE,
    填写数量_IN IN 药品收发记录.填写数量%TYPE,
    实际数量_IN IN 药品收发记录.实际数量%TYPE,
    成本价_IN IN 药品收发记录.成本价%TYPE,
    成本金额_IN IN 药品收发记录.成本金额%TYPE,
    零售金额_IN IN 药品收发记录.零售金额%TYPE,
    差价_IN IN 药品收发记录.差价%TYPE,
    入出类别ID_IN IN 药品收发记录.入出类别ID%TYPE,
    审核人_IN IN 药品收发记录.审核人%TYPE,
    审核日期_IN IN 药品收发记录.审核日期%TYPE,
    批号_IN IN 药品收发记录.批号%TYPE := NULL,
    效期_IN IN 药品收发记录.效期%TYPE := NULL,
    批准文号_IN In 药品收发记录.批准文号%TYPE := Null
)
IS
    Err_isverified EXCEPTION;
    Err_isNOnumber EXCEPTION;
    V_可用数量 药品库存.可用数量%TYPE;
    V_编码 收费项目目录.编码%TYPE;
    V_实际库存金额 药品库存.实际金额%TYPE;
    V_实际库存差价 药品库存.实际差价%TYPE;
    V_差价率 number(18,8);
    V_出库差价 药品库存.实际差价%TYPE;
    V_成本价 药品收发记录.成本价%TYPE;
    V_成本金额 药品收发记录.成本金额%TYPE;
	INTDIGIT NUMBER ;
BEGIN
	--获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';

    --由于领用处理允许在审核时改变实际数量，
      --所以首先对实际数量和其他相应的字段进行更新。
    BEGIN
        SELECT nvl(实际金额,0), nvl(实际差价,0), nvl(可用数量,0)
          INTO V_实际库存金额, V_实际库存差价, V_可用数量
          FROM 药品库存
         WHERE 药品ID = 药品ID_IN
            AND NVL (批次, 0) = 批次_IN
            AND 库房ID = 库房ID_IN
            AND 性质 = 1
            AND ROWNUM = 1;
    EXCEPTION
        WHEN OTHERS THEN
            V_实际库存金额 := 0;
            V_可用数量 := 0;
    END;

    IF V_实际库存金额 <= 0 THEN
        BEGIN
            SELECT 指导差价率 / 100
              INTO V_差价率
              FROM 药品规格
             WHERE 药品ID = 药品ID_IN;
        EXCEPTION
            WHEN OTHERS THEN
                V_差价率 := 0;
        END;
    ELSE
        V_差价率 := V_实际库存差价 / V_实际库存金额;
    END IF;

    V_出库差价 := round(零售金额_IN * V_差价率,INTDIGIT);
    IF 实际数量_IN=0 THEN 
        V_成本价 :=成本价_IN; 
    ELSE 
        V_成本价 := (零售金额_IN - V_出库差价) / 实际数量_IN; 
    END IF; 
    V_成本金额 := round(V_成本价 * 实际数量_IN,INTDIGIT);

    UPDATE 药品收发记录
        SET 审核人 = NVL (审核人_IN, 审核人),
             审核日期 = 审核日期_IN,
             实际数量 = 实际数量_IN,
             成本价 = V_成本价,
             成本金额 = V_成本金额,
             零售金额 = 零售金额_IN,
             差价 = V_出库差价
     WHERE NO = NO_IN
        AND 单据 = 7
        AND 药品ID = 药品ID_IN
        AND 序号 = 序号_IN
        AND 记录状态 = 1
        AND 审核人 IS NULL;

    --更改药品库存的相应数据
    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;

    IF 批次_IN > 0 THEN
        IF V_可用数量 + 填写数量_IN - 实际数量_IN < 0 THEN
            RAISE Err_isNOnumber;
        END IF;
    END IF;

    UPDATE 药品库存
        SET 可用数量 = NVL (可用数量, 0) + 填写数量_IN - 实际数量_IN,
             实际数量 = NVL (实际数量, 0) - 实际数量_IN,
             实际金额 = NVL (实际金额, 0) - 零售金额_IN,
             实际差价 = NVL (实际差价, 0) - V_出库差价
     WHERE 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND NVL (批次, 0) = NVL (批次_IN, 0)
        AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品库存
                        (
                            库房ID,
                            药品ID,
                            批次,
                            性质,
                            可用数量,
                            实际数量,
                            实际金额,
                            实际差价,
							上次批号,效期,上次产地,批准文号
                        )
              VALUES (
                  库房ID_IN,
                  药品ID_IN,
                  批次_IN,
                  1,
                  -实际数量_IN,
                  -实际数量_IN,
                  -零售金额_IN,
                  -V_出库差价,
				  批号_IN,效期_IN,产地_IN,批准文号_IN
              );
    END IF;

    DELETE
      FROM 药品库存
     WHERE 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND NVL (可用数量, 0) = 0
        AND NVL (实际数量, 0) = 0
        AND NVL (实际金额, 0) = 0
        AND NVL (实际差价, 0) = 0;

    --更药品收发汇总表的相应数据
    UPDATE 药品收发汇总
        SET 数量 = NVL (数量, 0) - 实际数量_IN,
             金额 = NVL (金额, 0) - 零售金额_IN,
             差价 = NVL (差价, 0) - V_出库差价
     WHERE 日期 = TRUNC (SYSDATE)
        AND 库房ID = 库房ID_IN
        AND 药品ID = 药品ID_IN
        AND 类别ID = 入出类别ID_IN
        AND 单据 = 7;

    IF SQL%NOTFOUND THEN
        Insert INTO 药品收发汇总
                        (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
              VALUES (
                  TRUNC (SYSDATE),
                  库房ID_IN,
                  药品ID_IN,
                  入出类别ID_IN,
                  7,
                  -实际数量_IN,
                  -零售金额_IN,
                  -V_出库差价
              );
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]'
        );
    WHEN Err_isNOnumber THEN
        SELECT 编码
          INTO V_编码
          FROM 收费项目目录
         WHERE ID = 药品ID_IN;
        Raise_application_error (
            -20101, '[ZLSOFT]编码为' || V_编码 || ',批号为' || 批号_IN ||
                          '的分批核算药品' ||
                          CHR (10) ||
                          CHR (13) ||
                          '可用库存数量不够！[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品领用_verify;
/

CREATE OR REPLACE PROCEDURE ZL_药品领用_STRIKE (
    行次_IN IN INTEGER,
    原记录状态_IN IN 药品收发记录.记录状态%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    冲销数量_IN IN 药品收发记录.实际数量%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    填制日期_IN IN 药品收发记录.填制日期%TYPE
)
IS
    Err_isstriked EXCEPTION;
    Err_isoutstock EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    Err_isBatch EXCEPTION;
    v_BatchCount INTEGER;    --原不分批现在分批的药品的数量

    V_库房ID 药品收发记录.库房ID%TYPE; 
    V_对方部门ID 药品收发记录.对方部门ID%TYPE;
    V_入出类别ID 药品收发记录.入出类别ID%TYPE ;
    V_产地 药品收发记录.产地%TYPE ; 
    V_批次 药品收发记录.批次%TYPE ; 
    V_批号 药品收发记录.批号%TYPE ; 
    V_效期 药品收发记录.效期%TYPE ; 
    V_成本价 药品收发记录.成本价%TYPE ; 
    V_成本金额 药品收发记录.成本金额%TYPE ; 
    V_扣率 药品收发记录.扣率%TYPE ; 
    V_零售价 药品收发记录.零售价%TYPE ; 
    V_零售金额 药品收发记录.零售金额%TYPE ; 
    V_差价 药品收发记录.差价%TYPE ; 
    V_摘要 药品收发记录.摘要%TYPE ; 
    V_剩余数量 药品收发记录.实际数量%TYPE; 
    V_剩余成本金额 药品收发记录.成本金额%Type;
    V_剩余零售金额 药品收发记录.零售金额%Type;
    V_入出系数 药品收发记录.入出系数%TYPE; 

    V_记录数 NUMBER; 
    V_收发ID 药品收发记录.ID%TYPE; 
    V_领用人 药品收发记录.领用人%Type;
    V_批准文号 药品收发记录.批准文号%TYPE;
	INTDIGIT NUMBER;
BEGIN
	--获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';

	IF 行次_IN =1 THEN
        UPDATE 药品收发记录 
        SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3) 
        WHERE NO = NO_IN AND 单据 = 7 AND 记录状态 =原记录状态_IN ; 
        IF SQL%ROWCOUNT = 0 THEN 
            RAISE ERR_ISSTRIKED; 
        END IF; 
    END IF;
    
    --主要针对原不分批现在分批的药品，不能对其审核
    SELECT COUNT(*) INTO V_BATCHCOUNT FROM        
            药品收发记录 A,药品规格 B
    WHERE A.药品ID=B.药品ID
        AND A.NO=NO_IN 
        AND A.单据=7
        AND Mod(A.记录状态,3)=0
        AND NVL(A.批次,0)=0
        AND ((NVL(B.药库分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')))
            OR NVL(B.药房分批,0)=1);
        
    IF V_BATCHCOUNT>0 THEN
        RAISE ERR_ISBATCH;
    END IF;  
    
    SELECT SUM(实际数量) AS 剩余数量,SUM(成本金额) AS 剩余成本金额,SUM(零售金额) AS 剩余零售金额,库房ID,对方部门ID,入出类别ID,入出系数,批次,产地,批号,效期,成本价,扣率,零售价,摘要,领用人,批准文号
    INTO  V_剩余数量,V_剩余成本金额,V_剩余零售金额,V_库房ID,V_对方部门ID,V_入出类别ID,V_入出系数,V_批次,V_产地,V_批号,V_效期,V_成本价,V_扣率,V_零售价,V_摘要,V_领用人,V_批准文号
    FROM 药品收发记录 
    WHERE NO=NO_IN 
    AND 单据=7
    AND 药品ID=药品ID_IN 
    AND 序号=序号_IN
    GROUP BY 库房ID,对方部门ID,入出类别ID,入出系数,批次,产地,批号,效期,成本价,扣率,零售价,摘要,领用人,批准文号;

    --冲销数量大于剩余数量，不允许
    IF V_剩余数量<冲销数量_IN THEN
        RAISE ERR_ISNONUM; 
    END IF;

    V_成本金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余成本金额,INTDIGIT);
    V_零售金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余零售金额,INTDIGIT);
    V_差价:=V_零售金额-V_成本金额;

    SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;
    Insert INTO 药品收发记录
    (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
    药品ID,批次,产地,批号,效期,填写数量,实际数量,成本价,
    成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,审核人,审核日期,领用人,批准文号)
    VALUES 
    (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),7,NO_IN,序号_IN,V_库房ID,V_对方部门ID,V_入出类别ID,V_入出系数,
    药品ID_IN,V_批次,V_产地,V_批号,V_效期,-冲销数量_IN,-冲销数量_IN,V_成本价,-V_成本金额,V_零售价,-V_零售金额,
    -V_差价,V_摘要,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN,V_领用人,V_批准文号);

    --原分批现不分批的药品,在C冲消时，要处理他
    BEGIN 
        SELECT COUNT(*) INTO V_记录数
        FROM 药品收发记录 A, 药品规格 B
        WHERE B.药品ID=药品ID_IN
        AND A.NO=NO_IN
        AND A.单据 = 7 
        AND Mod(A.记录状态,3)=0
        AND NVL(A.批次,0)>0
        AND (NVL(B.药库分批,0)=0 OR 
            (NVL(B.药房分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室'))))
        ;
    EXCEPTION 
        WHEN OTHERS THEN V_记录数:=0;
    END;
    IF V_记录数>0 THEN
        V_批次:=0;
    ELSE
        V_批次:=NVL (V_批次, 0);
    END IF;

    --更改药品库存表的相应数据
    UPDATE 药品库存
       SET 可用数量 = NVL (可用数量, 0) + NVL (冲销数量_IN, 0),
           实际数量 = NVL (实际数量, 0) + NVL (冲销数量_IN, 0),
           实际金额 = NVL (实际金额, 0) + NVL (V_零售金额, 0),
           实际差价 = NVL (实际差价, 0) + NVL (V_差价, 0)
     WHERE 库房ID = V_库房ID
       AND 药品ID = 药品ID_IN
       AND NVL (批次, 0) = V_批次
       AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO 药品库存
        (库房ID,药品ID,批次,性质,可用数量,实际数量,
            实际金额,实际差价,上次批号,效期,上次产地,批准文号)
        VALUES 
        (V_库房ID,药品ID_IN,V_批次,1,冲销数量_IN,冲销数量_IN,
        V_零售金额,V_差价,V_批号,V_效期,v_产地,V_批准文号);
    END IF;

    DELETE 药品库存
     WHERE 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND NVL (可用数量, 0) = 0
        AND NVL (实际数量, 0) = 0
        AND NVL (实际金额, 0) = 0
        AND NVL (实际差价, 0) = 0;

    --更改药品收发汇总表的相应数据

    UPDATE 药品收发汇总
        SET 数量 = NVL (数量, 0)  +NVL(冲销数量_IN, 0),
             金额 = NVL (金额, 0) +NVL(V_零售金额, 0),
             差价 = NVL (差价, 0) +NVL(V_差价, 0)
     WHERE 日期 = TRUNC (填制日期_IN)
        AND 库房ID = V_库房ID
        AND 药品ID = 药品ID_IN
        AND 类别ID = V_入出类别ID
        AND 单据 = 7;

    IF SQL%NOTFOUND THEN
        INSERT INTO 药品收发汇总
        (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
        VALUES 
        (TRUNC (填制日期_IN),V_库房ID,药品ID_IN,V_入出类别ID,7,冲销数量_IN,V_零售金额,V_差价);
    END IF;
EXCEPTION
    WHEN ERR_ISSTRIKED THEN
        RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
    WHEN ERR_ISBATCH THEN
        RAISE_APPLICATION_ERROR (-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品，不能冲销！[ZLSOFT]');  
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_药品领用_STRIKE;
/

CREATE OR REPLACE PROCEDURE ZL_药品移库_INSERT(
	NO_IN         IN 药品收发记录.NO%TYPE,
	序号_IN       IN 药品收发记录.序号%TYPE,
	库房ID_IN     IN 药品收发记录.库房ID%TYPE,
	对方部门ID_IN IN 药品收发记录.对方部门ID%TYPE,
	药品ID_IN     IN 药品收发记录.药品ID%TYPE,
	批次_IN       IN 药品收发记录.批次%TYPE,
	填写数量_IN   IN 药品收发记录.填写数量%TYPE,
	实际数量_IN   IN 药品收发记录.实际数量%TYPE,
	成本价_IN     IN 药品收发记录.成本价%TYPE,
	成本金额_IN   IN 药品收发记录.成本金额%TYPE,
	零售价_IN     IN 药品收发记录.零售价%TYPE,
	零售金额_IN   IN 药品收发记录.零售金额%TYPE,
	差价_IN       IN 药品收发记录.差价%TYPE,
	填制人_IN     IN 药品收发记录.填制人%TYPE,
	产地_IN       IN 药品收发记录.产地%TYPE := NULL,
	批号_IN       IN 药品收发记录.批号%TYPE := NULL,
	效期_IN       IN 药品收发记录.效期%TYPE := NULL,
	摘要_IN       IN 药品收发记录.摘要%TYPE := NULL,
	填制日期_IN   IN 药品收发记录.填制日期%TYPE := NULL,
  上次供应商ID_IN In 药品收发记录.供药单位ID%TYPE := Null,
  批准文号_IN In 药品收发记录.批准文号%TYPE := Null
	)
IS
	ERR_MutilROW EXCEPTION ;
	ERR_ISNONUMBER EXCEPTION;
	V_编码       收费项目目录.编码%TYPE;
  V_LNGID      药品收发记录.ID%TYPE; --收发ID
	V_入的类别ID 药品收发记录.入出类别ID%TYPE; --入出类别ID
	V_出的类别ID 药品收发记录.入出类别ID%TYPE; --入出类别ID
	V_批次       药品收发记录.批次%TYPE := NULL; --主要针对入库中实行药库分批的药品
	V_是否分批   INTEGER; --判断入库是否药库分批   1:分批；0：不分批
	V_药库分批   INTEGER; --判断入库是否药库分批   1:分批；0：不分批
	V_药房分批   INTEGER; --判断入库是否药库分批   1:分批；0：不分批
	intRecords NUMBER ;
  V_下可用库存 系统参数表.参数值%Type;
BEGIN
	SELECT B.ID
	INTO V_入的类别ID
	FROM 药品单据性质 A, 药品入出类别 B
	WHERE A.类别ID = B.ID AND A.单据 = 6 AND B.系数 = 1 AND ROWNUM < 2;
	SELECT B.ID
	INTO V_出的类别ID
	FROM 药品单据性质 A, 药品入出类别 B
	WHERE A.类别ID = B.ID AND A.单据 = 6 AND B.系数 = -1 AND ROWNUM < 2;

  INSERT INTO 药品收发记录
      (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
       药品ID,批次,产地,批号,效期,填写数量,实际数量,
    成本价,成本金额,零售价,零售金额,差价,
       摘要,填制人,填制日期,供药单位ID,批准文号)
  VALUES
      (药品收发记录_ID.NEXTVAL,1,6,NO_IN,序号_IN,库房ID_IN,对方部门ID_IN,V_出的类别ID,-1,
       药品ID_IN,批次_IN,产地_IN,批号_IN,效期_IN,填写数量_IN,实际数量_IN,
 成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,差价_IN,
       摘要_IN,填制人_IN,填制日期_IN,上次供应商ID_IN,批准文号_IN);

  SELECT NVL(药库分批, 0), NVL(药房分批, 0)
	INTO V_药库分批, V_药房分批
	FROM 药品规格
	WHERE 药品ID = 药品ID_IN;

  V_是否分批 := 0;
  IF V_药房分批 = 0 THEN
      IF V_药库分批 = 1 THEN
	BEGIN
		SELECT DISTINCT 0
		INTO V_是否分批
		FROM 部门性质说明
		WHERE ((工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室'))
		AND 部门ID = 对方部门ID_IN;
	EXCEPTION
		WHEN OTHERS THEN V_是否分批 := 1;
	END;
      END IF;
  ELSE
      V_是否分批 := 1;
  END IF;

  SELECT 药品收发记录_ID.NEXTVAL INTO V_LNGID FROM DUAL;

  IF V_是否分批 = 1 AND NVL(批次_IN, 0) = 0 THEN
      --入库分批且出库不分批
      V_批次 := V_LNGID;
  ELSIF V_是否分批 = 0 THEN
      --入库不分批
      V_批次 := 0;
  ELSIF NVL(批次_IN, 0) <> 0 THEN
      --入库分批且出库也分批
      V_批次 := 批次_IN;
  END IF;

  INSERT INTO 药品收发记录
      (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
       药品ID,批次,产地,批号,效期,填写数量,实际数量,
   成本价,成本金额,零售价,零售金额,差价,
       摘要,填制人,填制日期,供药单位ID,批准文号)
  VALUES
      (V_LNGID,1,6,NO_IN,序号_IN + 1,对方部门ID_IN,库房ID_IN,V_入的类别ID,1,
       药品ID_IN,V_批次,产地_IN,批号_IN,效期_IN,填写数量_IN,实际数量_IN,
       成本价_IN,成本金额_IN,零售价_IN,零售金额_IN,差价_IN,
       摘要_IN,填制人_IN,填制日期_IN,上次供应商ID_IN,批准文号_IN);

	--检查是否存在相同药品相同批次的数据，如果存在不允许保存
	SELECT COUNT(*) INTO intRecords
	FROM 药品收发记录
	WHERE 单据=6 AND NO=NO_IN AND 入出系数=-1 AND 药品ID+0=药品ID_IN AND Nvl(批次,0)=NVL(批次_IN,0);
	IF intRecords>1 THEN
		RAISE ERR_MutilROW;
	END IF ;
  
  --根据参数决定是否下发药库房的可用库存
	Select 参数值 Into v_下可用库存 From 系统参数表 Where 参数号 = 96;
  
  --参数为1表示在填单时下可用数量
	If v_下可用库存 = '1' Then
      UPDATE 药品库存
			SET 可用数量=NVL(可用数量,0)-实际数量_IN
			WHERE 库房ID=库房ID_IN AND 药品ID=药品ID_IN AND NVL(批次,0)=批次_IN AND 性质=1;
      
			IF SQL%ROWCOUNT=0 THEN 
				INSERT INTO 药品库存(库房ID,药品ID,批次,性质,可用数量,上次批号,效期,上次产地)
				VALUES (库房ID_IN,药品ID_IN,批次_IN,1,-1*实际数量_IN,
					批号_IN,效期_IN,产地_IN);
			END IF ;

			DELETE
			FROM 药品库存
			WHERE 库房ID = 库房ID_IN
			AND 药品ID = 药品ID_IN
			AND NVL(可用数量,0)=0
			AND NVL(实际数量,0)=0
			AND NVL(实际金额,0)=0
			AND NVL(实际差价,0)=0;
  End If;
EXCEPTION
    WHEN ERR_ISNONUMBER THEN
        SELECT 编码 INTO V_编码 FROM 收费项目目录 WHERE ID = 药品ID_IN;
        RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]编码为'||V_编码||',批号为'||
			批号_IN||'的药库分批药品'||CHR(10) ||CHR(13)||'可用库存数量不够！[ZLSOFT]');
	WHEN ERR_MutilROW THEN
		SELECT 编码 INTO V_编码 FROM 收费项目目录 WHERE ID = 药品ID_IN;
		RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]编码为'||V_编码||'的药品，存在多条重复的记录，请合并为一条记录！[ZLSOFT]');
    WHEN OTHERS THEN
        ZL_ERRORCENTER(SQLCODE, SQLERRM);
END ZL_药品移库_INSERT;
/


CREATE OR REPLACE PROCEDURE ZL_药品移库_DELETE (
    --删除药品收发记录及相应的表：药品库存
    NO_IN IN 药品收发记录.NO%TYPE
)
IS
	V_发送 药品收发记录.配药日期%TYPE;
	ERR_ISVERIFIED EXCEPTION;
  V_下可用库存 系统参数表.参数值%Type;

    CURSOR C_药品收发记录
    IS
		SELECT 实际数量, 库房ID, 批次, 药品ID,批号,效期,产地,供药单位id,批准文号
		FROM 药品收发记录
		WHERE NO = NO_IN
		AND 单据 = 6
		AND 入出系数 = -1
		ORDER BY 药品ID;
BEGIN
	--检查是否已发送，已发送的单据需要还原可用数量
	SELECT 配药日期 INTO V_发送
	FROM 药品收发记录
	WHERE 单据=6 AND NO=NO_IN AND ROWNUM<2;

	Select 参数值 Into v_下可用库存 From 系统参数表 Where 参数号 = 96;  
  
  --如果参数值为1也要恢复原来的可用数量
	IF V_发送 IS NOT NULL Or v_下可用库存='1' THEN
		--通过循环，恢复原来的可用数量
		FOR V_药品收发记录 IN C_药品收发记录 LOOP
			UPDATE 药品库存
			SET 可用数量 = NVL (可用数量, 0) + V_药品收发记录.实际数量
			WHERE 库房ID = V_药品收发记录.库房ID
			AND 药品ID = V_药品收发记录.药品ID
			AND NVL (批次, 0) = NVL (V_药品收发记录.批次, 0)
			AND 性质 = 1;

			IF SQL%NOTFOUND THEN
				INSERT INTO 药品库存
					(库房ID, 药品ID, 批次, 性质, 可用数量,上次批号,效期,上次产地,上次供应商Id,批准文号)
				VALUES (
					V_药品收发记录.库房ID,V_药品收发记录.药品ID,V_药品收发记录.批次,1,
					V_药品收发记录.实际数量,V_药品收发记录.批号,V_药品收发记录.效期,V_药品收发记录.产地,V_药品收发记录.供药单位id,V_药品收发记录.批准文号
					);
			END IF;

			DELETE
			FROM 药品库存
			WHERE 库房ID = V_药品收发记录.库房ID
			AND 药品ID = V_药品收发记录.药品ID
			AND NVL(可用数量,0) = 0
			AND NVL(实际数量,0) = 0
			AND NVL(实际金额,0) = 0
			AND NVL(实际差价,0) = 0;
		END LOOP;
	END IF ;

    DELETE--把入和出两种类别的移库单都删除
	FROM 药品收发记录
	WHERE NO = NO_IN
	AND 单据 = 6
	AND 记录状态 = 1
	AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE ERR_ISVERIFIED;
    END IF;
EXCEPTION
    WHEN ERR_ISVERIFIED THEN
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]该单据已经被他人删除或已被人审核！[ZLSOFT]');
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_药品移库_DELETE;
/


CREATE OR REPLACE PROCEDURE ZL_药品移库_VERIFY (
    序号_IN IN 药品收发记录.序号%TYPE,
    库房ID_IN IN 药品收发记录.库房ID%TYPE,
    对方部门ID_IN IN 药品收发记录.对方部门ID%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    产地_IN IN 药品收发记录.产地%TYPE,
    出批次_IN IN 药品收发记录.批次%TYPE,
    填写数量_IN IN 药品收发记录.填写数量%TYPE,
    实际数量_IN IN 药品收发记录.实际数量%TYPE,
    成本价_IN IN 药品收发记录.成本价%TYPE,
    成本金额_IN IN 药品收发记录.成本金额%TYPE,
    零售金额_IN IN 药品收发记录.零售金额%TYPE,
    差价_IN IN 药品收发记录.差价%TYPE,
    出类别ID_IN IN 药品收发记录.入出类别ID%TYPE,
    入类别ID_IN IN 药品收发记录.入出类别ID%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    审核人_IN IN 药品收发记录.审核人%TYPE,
    批号_IN IN 药品收发记录.批号%TYPE := NULL,
    效期_IN IN 药品收发记录.效期%TYPE := NULL,
    审核日期_IN IN 药品收发记录.审核日期%TYPE := NULL,
    移库单_IN IN NUMBER:=1,
    上次供应商ID_IN In 药品收发记录.供药单位ID%TYPE := Null,
    批准文号_IN In 药品收发记录.批准文号%TYPE := Null
)
IS
    ERR_ISVERIFIED EXCEPTION;
    ERR_ISNONUMBER EXCEPTION;
    V_入批次 药品收发记录.批次%TYPE := NULL;
    V_实际库存金额 药品库存.实际金额%TYPE;
    V_实际库存差价 药品库存.实际差价%TYPE;
    V_差价率 NUMBER(18,8);
    V_出库差价 药品库存.实际差价%TYPE;
    V_成本价 药品收发记录.成本价%TYPE;
    V_成本金额 药品收发记录.成本金额%TYPE;
    V_实际数量 药品库存.实际数量%TYPE;
    V_编码 收费项目目录.编码%TYPE;
	INTDIGIT NUMBER ;
BEGIN
	--获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';

	--由于移库处理允许在审核时改变实际数量，
      --所以首先对实际数量和其他相应的字段进行更新。
    BEGIN
		SELECT NVL(实际金额,0), NVL(实际差价,0), NVL(实际数量,0)
		INTO V_实际库存金额, V_实际库存差价, V_实际数量
		FROM 药品库存
		WHERE 药品ID = 药品ID_IN
		AND NVL (批次, 0) = 出批次_IN
		AND 库房ID = 库房ID_IN
		AND 性质 = 1
		AND ROWNUM = 1;
    EXCEPTION
        WHEN OTHERS THEN
            V_实际库存金额 := 0;
            V_实际数量 := 0;
    END;

    IF V_实际库存金额 <= 0 THEN
        BEGIN
			SELECT 指导差价率 / 100
			INTO V_差价率
			FROM 药品规格
			WHERE 药品ID = 药品ID_IN;
        EXCEPTION
            WHEN OTHERS THEN
                V_差价率 := 0;
        END;
    ELSE
        V_差价率 := V_实际库存差价 / V_实际库存金额;
    END IF;

    V_出库差价 := ROUND(零售金额_IN * V_差价率,INTDIGIT);
    IF 实际数量_IN=0 THEN
        V_成本价 :=成本价_IN;
    ELSE
        V_成本价 := (零售金额_IN - V_出库差价) / 实际数量_IN; 
    END IF; 
    V_成本金额 := ROUND(V_成本价 * 实际数量_IN,INTDIGIT);

    UPDATE 药品收发记录
        SET 审核人 = NVL (审核人_IN, 审核人),
             审核日期 = 审核日期_IN,
             实际数量 = 实际数量_IN,
             成本价 = V_成本价,
             成本金额 = V_成本金额,
             零售金额 = 零售金额_IN,
             差价 = V_出库差价
     WHERE NO = NO_IN
        AND 单据 = 6
        AND 药品ID = 药品ID_IN
        AND 记录状态 = 1
        AND 序号 IN (序号_IN, 序号_IN + 1)
        AND 审核人 IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE ERR_ISVERIFIED;
    END IF;

    IF 出批次_IN > 0 THEN
        IF V_实际数量 < 实际数量_IN THEN
            RAISE ERR_ISNONUMBER;
        END IF;
    END IF;

    --取入类别的批次
	SELECT 批次
	INTO V_入批次
	FROM 药品收发记录
	WHERE NO = NO_IN
	AND 单据 = 6
	AND 记录状态 = 1
	AND 序号 = 序号_IN+1;
        
    --更改入类别的药品库存的相应数据

	UPDATE 药品库存
	SET 可用数量 = NVL (可用数量, 0) + 实际数量_IN,
		实际数量 = NVL (实际数量, 0) + 实际数量_IN,
		实际金额 = NVL (实际金额, 0) + 零售金额_IN,
		实际差价 = NVL (实际差价, 0) + V_出库差价,
		上次采购价 = V_成本价,
		上次批号 = NVL (批号_IN, 上次批号),
		上次产地 = NVL (产地_IN, 上次产地),
		效期 = NVL (效期_IN, 效期),
    上次供应商ID=上次供应商ID_IN,
    批准文号=批准文号_IN
	WHERE 库房ID = 对方部门ID_IN
	AND 药品ID = 药品ID_IN
	AND NVL (批次, 0) = NVL (V_入批次, 0)
	AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO 药品库存
			(库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次采购价,上次批号,上次产地,效期,上次供应商ID,批准文号)
        VALUES (
			对方部门ID_IN,药品ID_IN,V_入批次,1,实际数量_IN,实际数量_IN,零售金额_IN,
            V_出库差价,V_成本价,批号_IN,产地_IN,效期_IN,上次供应商ID_IN,批准文号_IN);
    END IF;

    --更改出类别的药品库存的相应数据

    UPDATE 药品库存
	SET 
		实际数量 = NVL (实际数量, 0) - 实际数量_IN,
		实际金额 = NVL (实际金额, 0) - 零售金额_IN,
		实际差价 = NVL (实际差价, 0) - V_出库差价
	WHERE 库房ID = 库房ID_IN
	AND 药品ID = 药品ID_IN
	AND NVL (批次, 0) = NVL (出批次_IN, 0)
	AND 性质 = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO 药品库存(库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次批号,效期,上次供应商ID,批准文号)
        VALUES (库房ID_IN,药品ID_IN,出批次_IN,1,0,-实际数量_IN,-零售金额_IN,-V_出库差价,批号_IN,效期_IN,上次供应商ID_IN,批准文号_IN);
    END IF;

	DELETE
	FROM 药品库存
	WHERE 库房ID = 库房ID_IN
	AND 药品ID = 药品ID_IN
	AND NVL (可用数量, 0) = 0
	AND NVL (实际数量, 0) = 0
	AND NVL (实际金额, 0) = 0
	AND NVL (实际差价, 0) = 0;

    --更改入类别的药品收发汇总表的相应数据
	UPDATE 药品收发汇总
	SET 数量 = NVL (数量, 0) + 实际数量_IN,
		金额 = NVL (金额, 0) + 零售金额_IN,
		差价 = NVL (差价, 0) + V_出库差价
	WHERE 日期 = TRUNC (SYSDATE)
	AND 库房ID = 对方部门ID_IN
	AND 药品ID = 药品ID_IN
	AND 类别ID = 入类别ID_IN
	AND 单据 = 6;

    IF SQL%NOTFOUND THEN
		INSERT INTO 药品收发汇总
			(日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
		VALUES(
			TRUNC(SYSDATE),对方部门ID_IN,药品ID_IN,入类别ID_IN,6,实际数量_IN,零售金额_IN,V_出库差价);
    END IF;

    --更改出类别的药品收发汇总表的相应数据
    UPDATE 药品收发汇总
	SET 数量 = NVL (数量, 0) - 实际数量_IN,
		金额 = NVL (金额, 0) - 零售金额_IN,
		差价 = NVL (差价, 0) - V_出库差价
	WHERE 日期 = TRUNC (SYSDATE)
	AND 库房ID = 库房ID_IN
	AND 药品ID = 药品ID_IN
	AND 类别ID = 出类别ID_IN
	AND 单据 = 6;

    IF SQL%NOTFOUND THEN
        INSERT INTO 药品收发汇总
			(日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
		VALUES (
			TRUNC(SYSDATE),库房ID_IN,药品ID_IN,出类别ID_IN,6,-实际数量_IN,-零售金额_IN,-V_出库差价);
    END IF;
EXCEPTION
    WHEN ERR_ISVERIFIED THEN
        RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]该单据已经被他人审核！[ZLSOFT]');
    WHEN ERR_ISNONUMBER THEN
		SELECT 编码
		INTO V_编码
		FROM 收费项目目录
		WHERE ID = 药品ID_IN;
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]编码为' || V_编码 || ',批号为' || 批号_IN ||
                          '的药库分批药品' || CHR(10) || CHR(13) || '可用库存数量不够！[ZLSOFT]');
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_药品移库_VERIFY;
/


CREATE OR REPLACE PROCEDURE ZL_药品移库_STRIKE (
    行次_IN IN INTEGER,
    原记录状态_IN IN 药品收发记录.记录状态%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    序号_IN IN 药品收发记录.序号%TYPE,
    药品ID_IN IN 药品收发记录.药品ID%TYPE,
    冲销数量_IN IN 药品收发记录.实际数量%TYPE,
    填制人_IN IN 药品收发记录.填制人%TYPE,
    填制日期_IN IN 药品收发记录.填制日期%TYPE
)
IS
    ERR_ISSTRIKED EXCEPTION;
    ERR_ISOUTSTOCK EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    ERR_ISBATCH EXCEPTION;
    V_BATCHCOUNT INTEGER;    --原不分批现在分批的药品的数量
    V_序号 药品收发记录.序号%TYPE;
    V_库房ID 药品收发记录.库房ID%TYPE;
    V_对方部门ID 药品收发记录.对方部门ID%TYPE;
    V_入出类别ID 药品收发记录.入出类别ID%TYPE ;
    V_产地 药品收发记录.产地%TYPE ;
    V_批次 药品收发记录.批次%TYPE ;
    V_批号 药品收发记录.批号%TYPE ;
    V_效期 药品收发记录.效期%TYPE ;
    V_成本价 药品收发记录.成本价%TYPE ;
    V_成本金额 药品收发记录.成本金额%TYPE ;
    V_扣率 药品收发记录.扣率%TYPE ;
    V_零售价 药品收发记录.零售价%TYPE ;
    V_零售金额 药品收发记录.零售金额%TYPE ;
    V_差价 药品收发记录.差价%TYPE ;
    V_摘要 药品收发记录.摘要%TYPE ;
    V_剩余数量 药品收发记录.实际数量%TYPE;
    V_剩余成本金额 药品收发记录.成本金额%Type;
    V_剩余零售金额 药品收发记录.零售金额%Type;
    V_入出系数 药品收发记录.入出系数%TYPE;
    V_记录数 NUMBER;
    V_收发ID 药品收发记录.ID%TYPE;
    V_备药人 药品收发记录.配药人%TYPE;
    V_发送日期 药品收发记录.配药日期%TYPE;
    V_上次供应商ID 药品收发记录.供药单位id%Type;
    V_批准文号 药品收发记录.批准文号%TYPE;
 
    --对冲销数量进行检查
    V_库存数 药品库存.实际数量%TYPE;
    V_药库分批 INTEGER;
    V_药房分批 INTEGER;
    V_分批属性 INTEGER;
    V_药库 INTEGER;
    V_分批 NUMBER;
	INTDIGIT NUMBER;
 
    CURSOR C_药品收发记录
    IS
    SELECT 序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,批次,产地,批号,效期,配药人,配药日期,摘要,供药单位ID,批准文号
    FROM 药品收发记录
    WHERE NO = NO_IN AND 单据 = 6 AND (序号>=序号_IN AND 序号<=序号_IN+1) AND (记录状态=1 OR MOD(记录状态,3)=0)
    ORDER BY 药品ID;
BEGIN
         --获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';
 
    IF 行次_IN =1 THEN
        UPDATE 药品收发记录
        SET 记录状态 = DECODE(原记录状态_IN,1,3,原记录状态_IN+3)
        WHERE NO = NO_IN AND 单据 = 6 AND 记录状态 =原记录状态_IN ;
        IF SQL%ROWCOUNT = 0 THEN
            RAISE ERR_ISSTRIKED;
        END IF;
    END IF;
 
    SELECT COUNT(*) INTO V_BATCHCOUNT FROM
            药品收发记录 A,药品规格 B
    WHERE A.药品ID=B.药品ID
        AND A.NO=NO_IN
        AND A.单据=6
        AND A.药品ID+0=药品ID_IN
        AND MOD(A.记录状态,3)=0
        AND NVL(A.批次,0)=0
        AND ((NVL(B.药库分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')))
            OR NVL(B.药房分批,0)=1);
 
    IF V_BATCHCOUNT>0 THEN
        RAISE ERR_ISBATCH;
    END IF;
    
    SELECT SUM(A.实际数量) AS 剩余数量,SUM(A.成本金额) AS 剩余成本金额,SUM(A.零售金额) AS 剩余零售金额,A.成本价,A.零售价,A.对方部门ID,NVL(A.批次,0),B.药库分批,B.药房分批,A.批准文号
    INTO  V_剩余数量,V_剩余成本金额,V_剩余零售金额,V_成本价,V_零售价,V_库房ID,V_批次,V_药库分批,V_药房分批,V_批准文号
    FROM 药品收发记录 A,药品规格 B
    WHERE A.NO=NO_IN AND A.药品ID=B.药品ID AND A.单据=6 AND A.药品ID=药品ID_IN AND A.序号=序号_IN
    GROUP BY A.成本价,A.零售价,A.对方部门ID,NVL(A.批次,0),B.药库分批,B.药房分批,A.批准文号;
 
    --判断该部门是药库还是药房
    BEGIN
        SELECT DISTINCT 0
        INTO V_药库
        FROM 部门性质说明
        WHERE (   (工作性质 LIKE '%药房')
              OR (工作性质 LIKE '制剂室'))
        AND 部门ID = V_库房ID;
    EXCEPTION
        WHEN OTHERS THEN V_药库:=1;
    END ;
 
    --根据部门性质,判断分批特性
    IF V_药库=0 THEN
        V_分批属性:=V_药房分批;
    ELSE
        V_分批属性:=V_药库分批;
    END IF ;
 
    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    --因为对于对方库房，相当于出库，而对于当前库房，相当于入库，所以当前库房不予检查，仅检查退库那条记录
    SELECT NVL(A.批次,0) INTO V_批次
    FROM 药品收发记录 A
    WHERE A.NO=NO_IN AND A.单据=6 AND A.药品ID=药品ID_IN AND A.序号=序号_IN+1 AND MOD(A.记录状态,3)=0;
 
    --取库存数
    BEGIN
        SELECT NVL(实际数量,0) INTO V_库存数 FROM 药品库存
        WHERE 库房ID=V_库房ID AND 药品ID=药品ID_IN AND NVL(批次,0)=V_批次 AND 性质=1;
    EXCEPTION
        WHEN OTHERS THEN V_库存数:=0;
    END ;
 
    --如果库存数大于剩余数量,取剩余数量;否则取库存数
    IF V_库存数<V_剩余数量 Then
       v_剩余成本金额:=V_库存数/V_剩余数量*v_剩余成本金额;
       V_剩余零售金额:=V_库存数/V_剩余数量*V_剩余零售金额;
       V_剩余数量:=V_库存数;
    END IF ;
 
    --冲销数量大于剩余数量，不允许
    IF V_剩余数量<冲销数量_IN THEN
        RAISE ERR_ISNONUM;
    END IF;
 
    V_成本金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余成本金额,INTDIGIT);
    V_零售金额:= ROUND(冲销数量_IN/v_剩余数量*V_剩余零售金额,INTDIGIT);
    V_差价:=V_零售金额-V_成本金额;
 
    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        V_序号:=V_药品收发记录.序号;
        V_库房ID:=V_药品收发记录.库房ID;
        V_对方部门ID:=V_药品收发记录.对方部门ID;
        V_入出类别ID:=V_药品收发记录.入出类别ID;
        V_入出系数:=V_药品收发记录.入出系数;
        V_批次:=V_药品收发记录.批次;
        V_产地:=V_药品收发记录.产地;
        V_批号:=V_药品收发记录.批号;
        V_效期:=V_药品收发记录.效期;
         v_摘要:=v_药品收发记录.摘要;
         V_备药人:=V_药品收发记录.配药人;
         V_发送日期:=V_药品收发记录.配药日期;
         V_上次供应商ID:=V_药品收发记录.供药单位ID;
         V_批准文号:=V_药品收发记录.批准文号;
 
        SELECT 药品收发记录_ID.NEXTVAL INTO V_收发ID  FROM DUAL;
        INSERT INTO 药品收发记录
        (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,
        药品ID,批次,产地,批号,效期,填写数量,实际数量,成本价,
        成本金额,零售价,零售金额,差价,摘要,填制人,填制日期,审核人,审核日期,配药人,配药日期,供药单位ID,批准文号)
        VALUES
        (V_收发ID,DECODE(原记录状态_IN,1,2,原记录状态_IN+2),6,NO_IN,V_序号,V_库房ID,V_对方部门ID,V_入出类别ID,V_入出系数,
        药品ID_IN,V_批次,V_产地,V_批号,V_效期,-冲销数量_IN,-冲销数量_IN,V_成本价,-V_成本金额,V_零售价,-V_零售金额,
        -V_差价,V_摘要,填制人_IN,填制日期_IN,填制人_IN,填制日期_IN,V_备药人,V_发送日期,V_上次供应商ID,V_批准文号);
 
        --原分批现不分批的药品,在C冲消时，要处理他
        BEGIN
            SELECT COUNT(*) INTO V_记录数
            FROM 药品收发记录 A, 药品规格 B
            WHERE B.药品ID=A.药品ID
            AND A.药品ID+0=药品ID_IN
            AND A.NO=NO_IN
            AND A.单据 = 6
            AND A.库房ID+0=V_库房ID
            AND MOD(A.记录状态,3)=0
            AND NVL(A.批次,0)>0
            AND (NVL(B.药库分批,0)=0 OR
                (NVL(B.药房分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室'))))
            ;
        EXCEPTION
            WHEN OTHERS THEN V_记录数:=0;
        END;
        IF V_记录数>0 THEN
            V_批次:=0;
        ELSE
            V_批次:=NVL (V_批次, 0);
        END IF;
 
        --更改药品库存表的相应数据
        UPDATE 药品库存
            SET 可用数量=NVL(可用数量,0)-NVL(冲销数量_IN,0)*V_入出系数,
                实际数量=NVL(实际数量,0)-NVL(冲销数量_IN,0)*V_入出系数,
                实际金额=NVL(实际金额,0)-NVL(V_零售金额,0)*V_入出系数,
                实际差价=NVL(实际差价,0)-NVL(V_差价,0)*V_入出系数,
                上次采购价=NVL(V_成本价,上次采购价),
                上次批号=NVL(V_批号,上次批号),
                上次产地=NVL(V_产地,上次产地),
                效期=NVL(V_效期,效期)
          WHERE 库房ID = V_库房ID
            AND 药品ID = 药品ID_IN
            AND NVL (批次, 0) = V_批次
            AND 性质 = 1;
 
        IF SQL%NOTFOUND THEN
            INSERT INTO 药品库存
            (库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,
            实际差价,上次采购价,上次批号,上次产地,效期,上次供应商id,批准文号)
            VALUES
            (V_库房ID,药品ID_IN,V_批次,1,-冲销数量_IN*V_入出系数,-冲销数量_IN*V_入出系数,
            -V_零售金额*V_入出系数,-V_差价*V_入出系数,V_成本价,V_批号,V_产地,V_效期,V_上次供应商ID,V_批准文号);
        END IF;
 
        DELETE
          FROM 药品库存
         WHERE 库房ID = V_库房ID
           AND 药品ID = 药品ID_IN
           AND NVL(可用数量,0)=0
           AND NVL(实际数量,0)=0
           AND NVL(实际金额,0)=0
           AND NVL(实际差价,0)=0;
 
        --更改药品收发汇总表的相应数据
        UPDATE 药品收发汇总
         SET 数量 =    NVL (数量,0)  - NVL (冲销数量_IN,0)*V_入出系数,
             金额 = NVL (金额, 0) - NVL (V_零售金额, 0)*V_入出系数,
             差价 = NVL (差价, 0) - NVL (V_差价, 0)*V_入出系数
        WHERE 日期 = TRUNC (填制日期_IN)
         AND 库房ID = V_库房ID
         AND 药品ID = 药品ID_IN
         AND 类别ID = V_入出类别ID
         AND 单据 = 6;
        IF SQL%NOTFOUND THEN
            INSERT INTO 药品收发汇总
            (日期, 库房ID, 药品ID, 类别ID, 单据, 数量, 金额, 差价)
            VALUES
            (TRUNC (填制日期_IN),V_库房ID,药品ID_IN,V_入出类别ID,
            6,-冲销数量_IN*V_入出系数,-V_零售金额*V_入出系数,-V_差价*V_入出系数);
        END IF;
    END LOOP;
EXCEPTION
    WHEN ERR_ISSTRIKED THEN
        RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
    WHEN ERR_ISBATCH THEN
        RAISE_APPLICATION_ERROR (-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品，不能冲销！[ZLSOFT]');
    WHEN ERR_ISNONUM THEN
        RAISE_APPLICATION_ERROR (-20103, '[ZLSOFT]该单据中第' || ceil(序号_IN/2) || '行的药品冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]' );
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_药品移库_STRIKE;
/

CREATE OR REPLACE PROCEDURE ZL_药品移库_PREPARE(
    NO_IN IN 药品收发记录.NO%TYPE,
    操作员_IN VARCHAR2:=NULL
)
IS
  str发送 VARCHAR2(20);
  V_下可用库存 系统参数表.参数值%Type;

    CURSOR C_药品收发记录
    IS
		SELECT 实际数量,库房ID,批次,药品ID,批号,效期,产地
		FROM 药品收发记录
		WHERE NO = NO_IN
		AND 单据 = 6
		AND 入出系数 = -1
		ORDER BY 药品ID;
BEGIN
    IF 操作员_IN IS NOT NULL THEN
		UPDATE 药品收发记录
		SET 配药人=操作员_IN,
			外观=to_char(SYSDATE,'yyyy-MM-dd hh24:mi:ss')
		WHERE 单据=6 AND NO=NO_IN;
	ELSE
		SELECT to_char(SYSDATE,'yyyy-MM-dd hh24:mi:ss') INTO str发送 FROM dual;

		UPDATE 药品收发记录
		SET 配药日期=to_date(str发送,'yyyy-MM-dd hh24:mi:ss')
		WHERE 单据=6 AND NO=NO_IN;
    
    --根据参数决定是否下发药库房的可用库存
    Select 参数值 Into v_下可用库存 From 系统参数表 Where 参数号 = 96;
    
    --参数为0表示在发送时才下可用数量
    If v_下可用库存='0' Then
  		FOR v_药品收发记录 IN c_药品收发记录 LOOP
  			UPDATE 药品库存
  			SET 可用数量=NVL(可用数量,0)-v_药品收发记录.实际数量
  			WHERE 库房ID=v_药品收发记录.库房ID AND 药品ID=v_药品收发记录.药品ID AND NVL(批次,0)=NVL(v_药品收发记录.批次,0) AND 性质=1;
  			IF SQL%ROWCOUNT=0 THEN
  				INSERT INTO 药品库存(库房ID,药品ID,批次,性质,可用数量,上次批号,效期,上次产地)
  				VALUES (v_药品收发记录.库房ID,v_药品收发记录.药品ID,NVL(v_药品收发记录.批次,0),1,-1*v_药品收发记录.实际数量,
  					v_药品收发记录.批号,v_药品收发记录.效期,v_药品收发记录.产地);
  			END IF ;
  
  			DELETE
  			FROM 药品库存
  			WHERE 库房ID = v_药品收发记录.库房ID
  			AND 药品ID = v_药品收发记录.药品ID
  			AND NVL(可用数量,0)=0
  			AND NVL(实际数量,0)=0
  			AND NVL(实际金额,0)=0
  			AND NVL(实际差价,0)=0;
  		END LOOP ;
    End If;
	END IF ;
EXCEPTION
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_药品移库_PREPARE;
/


CREATE OR REPLACE PROCEDURE ZL_药品移库_BACK(
    NO_IN IN 药品收发记录.NO%TYPE
)
IS
	str发送 药品收发记录.配药日期%TYPE;
	str备药 药品收发记录.配药人%TYPE;
	str审核 药品收发记录.审核人%TYPE;

	Err_Note VARCHAR2(255);
	Err_Custom EXCEPTION ;
  
  V_下可用库存 系统参数表.参数值%Type;
  
    CURSOR C_药品收发记录
    IS
		SELECT 实际数量,库房ID,批次,药品ID,批号,效期,产地,供药单位ID,批准文号
		FROM 药品收发记录
		WHERE NO = NO_IN
		AND 单据 = 6
		AND 入出系数 = -1
		ORDER BY 药品ID;
BEGIN
    SELECT 配药人,配药日期,审核人 INTO str备药,str发送,str审核
	FROM 药品收发记录
	WHERE 单据=6 AND NO=NO_IN AND ROWNUM<2;

	IF str审核 IS NOT NULL THEN
		Err_Note:='该单据已被接收库房接收，不再允许回退！';
		RAISE Err_Custom;
	END IF ;
	IF str备药 IS NULL THEN
		RETURN ;
	END IF ;
	IF str发送 IS NULL THEN
		--仅更新配药人为空即可
		UPDATE 药品收发记录
		SET 配药人=NULL,
			外观=NULL
		WHERE 单据=6 AND NO=NO_IN;
	ELSE
		--需要恢复出库库房的可用数量
		UPDATE 药品收发记录
		SET 配药日期=NULL
		WHERE 单据=6 AND NO=NO_IN;

    --根据参数决定是否恢复发药库房的可用库存
  	Select 参数值 Into v_下可用库存 From 系统参数表 Where 参数号 = 96;
    
    --参数为0表示回退时要恢复可用数量
  	If v_下可用库存 = '0' Then  
      FOR v_药品收发记录 IN c_药品收发记录 LOOP
  			UPDATE 药品库存
  			SET 可用数量=NVL(可用数量,0)+v_药品收发记录.实际数量
  			WHERE 库房ID=v_药品收发记录.库房ID AND 药品ID=v_药品收发记录.药品ID AND NVL(批次,0)=NVL(v_药品收发记录.批次,0) AND 性质=1;
  			IF SQL%ROWCOUNT=0 THEN
  				INSERT INTO 药品库存(库房ID,药品ID,批次,性质,可用数量,上次批号,效期,上次产地,上次供应商ID,批准文号)
  				VALUES (v_药品收发记录.库房ID,v_药品收发记录.药品ID,NVL(v_药品收发记录.批次,0),1,v_药品收发记录.实际数量,
  					v_药品收发记录.批号,v_药品收发记录.效期,v_药品收发记录.产地,v_药品收发记录.供药单位ID,v_药品收发记录.批准文号);
  			END IF ;
  
  			DELETE
  			FROM 药品库存
  			WHERE 库房ID = v_药品收发记录.库房ID
  			AND 药品ID = v_药品收发记录.药品ID
  			AND NVL(可用数量,0)=0
  			AND NVL(实际数量,0)=0
  			AND NVL(实际金额,0)=0
  			AND NVL(实际差价,0)=0;
  		END LOOP ;
  	END IF ;
  End If;
EXCEPTION
    WHEN Err_Custom THEN
        RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]'||Err_Note||'[ZLSOFT]');
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_药品移库_BACK;
/

CREATE OR REPLACE PROCEDURE zl_药品库存差价调整_strike (
    NO_IN IN 药品收发记录.NO%TYPE,
    审核人_IN IN 药品收发记录.审核人%TYPE
)
IS
    Err_isstriked EXCEPTION;
    Err_isBatch exception;
    v_BatchCount integer;    --原不分批现在分批的药品的数量
    V_COUNT INTEGER;    --原分批现不分批
    V_批次 药品收发记录.批次%TYPE;

    CURSOR C_药品收发记录
    IS
        SELECT 入出类别ID, 库房ID, 药品ID, 批次, 差价,批号,效期,产地
          FROM 药品收发记录 A
         WHERE NO = NO_IN
            AND 单据 = 5
            AND 记录状态 = 2
         ORDER BY 药品ID;
BEGIN
    UPDATE 药品收发记录
        SET 记录状态 = 3
     WHERE NO = NO_IN
        AND 单据 = 5
        AND 记录状态 = 1;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isstriked;
    END IF;
    
    --主要针对原不分批现在分批的药品，不能对其审核
    SELECT COUNT(*) INTO v_BatchCount FROM        
            药品收发记录 a,药品规格 b
    WHERE a.药品id=b.药品id
        AND a.no=NO_IN 
        AND a.单据=5
        AND a.记录状态=3
        AND nvl(a.批次,0)=0
        AND ((NVL(B.药库分批,0)=1 AND A.库房ID NOT IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')))
            OR NVL(B.药房分批,0)=1);
    
    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  

    Insert INTO 药品收发记录
                    (
                        ID,
                        记录状态,
                        单据,
                        NO,
                        序号,
                        库房ID,
                        入出类别ID,
                        入出系数,
                        药品ID,
                        批次,
                        产地,
                        批号,
                        效期,
                        零售金额,
                        差价,
                        摘要,
                        填制人,
                        填制日期,
                        审核人,
                        审核日期
                    )
        SELECT 药品收发记录_ID.Nextval, 2, 单据, NO_IN, 序号, 库房ID,
                 入出类别ID,
                 入出系数, a.药品ID, 
                 DECODE(NVL(a.批次,0),0,NULL,(DECODE(NVL(b.药库分批,0),0,NULL,a.批次))), 
                 a.产地, 批号, 效期, 零售金额, -差价, 摘要,
                 审核人_IN, SYSDATE, 审核人_IN, SYSDATE
          FROM 药品收发记录 a,药品规格 b
         WHERE NO = NO_IN
            AND a.药品id=b.药品id
            AND 单据 = 5
            AND 记录状态 = 3;

    FOR V_药品收发记录 IN C_药品收发记录 LOOP
        --原分批现不分批的药品,在C冲消时，要处理他
        BEGIN 
            SELECT COUNT(*) INTO V_COUNT
            FROM 药品收发记录 A, 药品规格 B
            WHERE B.药品ID=V_药品收发记录.药品ID
            AND A.NO=NO_IN
            AND A.单据 = 5 
            and a.库房id+0=V_药品收发记录.库房id
            AND A.记录状态 = 3 
            AND NVL(A.批次,0)>0
            AND (NVL(B.药库分批,0)=0 OR 
                (NVL(B.药房分批,0)=0 AND A.库房ID IN (SELECT 部门ID FROM  部门性质说明 WHERE (工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室'))))
            ;
        EXCEPTION 
            WHEN OTHERS THEN
                V_COUNT:=0;
        END;
        IF V_COUNT>0 THEN
            V_批次:=0;
        ELSE
            V_批次:=NVL (V_药品收发记录.批次, 0);
        END IF;

        --更改药品库存表的相应数据

        UPDATE 药品库存
            SET 实际差价 = NVL (实际差价, 0) + NVL (V_药品收发记录.差价, 0)
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND NVL (批次, 0) = V_批次
            AND 性质 = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品库存
                (库房ID, 药品ID, 批次, 性质, 实际差价,上次批号,效期,上次产地)
                  VALUES (
                      V_药品收发记录.库房ID,
                      V_药品收发记录.药品ID,
                      V_批次,
                      1,
                      V_药品收发记录.差价,
                      V_药品收发记录.批号,
                      V_药品收发记录.效期,
					  v_药品收发记录.产地
                  );
        END IF;

        DELETE
          FROM 药品库存
         WHERE 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND nvl(可用数量,0) = 0 AND nvl(实际数量,0) = 0 AND nvl(实际金额,0) = 0 AND nvl(实际差价,0) = 0;

        --更改药品收发汇总表的相应数据

        UPDATE 药品收发汇总
            SET 差价 = NVL (差价, 0) + NVL (V_药品收发记录.差价, 0)
         WHERE 日期 = TRUNC (SYSDATE)
            AND 库房ID = V_药品收发记录.库房ID
            AND 药品ID = V_药品收发记录.药品ID
            AND 类别ID = V_药品收发记录.入出类别ID
            AND 单据 = 5;

        IF SQL%NOTFOUND THEN
            Insert INTO 药品收发汇总
                            (日期, 库房ID, 药品ID, 类别ID, 单据, 差价)
                  VALUES (
                      TRUNC (SYSDATE),
                      V_药品收发记录.库房ID,
                      V_药品收发记录.药品ID,
                      V_药品收发记录.入出类别ID,
                      5,
                      V_药品收发记录.差价
                  );
        END IF;
    END LOOP;
EXCEPTION
    WHEN Err_isstriked THEN
        Raise_application_error (
            -20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]'
        );
    WHEN Err_isBatch THEN
        Raise_application_error ( 
            -20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品，不能冲销！[ZLSOFT]' 
        );  
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品库存差价调整_strike;
/

CREATE OR REPLACE PROCEDURE zl_药品外购发票信息_UPDATE (
    NO_IN IN 药品收发记录.NO%TYPE := NULL,
    序号_IN IN 药品收发记录.序号%TYPE,
    发票号_IN IN 应付记录.发票号%TYPE := NULL,
    发票日期_IN IN 应付记录.发票日期%TYPE := NULL,
    发票金额_IN IN 应付记录.发票金额%TYPE := NULL,
    供药单位_IN in 应付记录.单位ID%TYPE:=0,
    操作标志_IN Number                        --1、未冲销单据修改发票信息; 2、部分冲销单据修改发票信息
)
IS
	ErrInfor varchar2(255);
	ErrItem exception;

	V_NO 应付记录.NO%TYPE;
	V_应付ID 应付记录.ID%TYPE;
    V_收发ID 应付记录.收发ID%TYPE;
	V_付款序号 应付记录.付款序号%TYPE;
    V_发票金额 应付记录.发票金额%TYPE;--旧发票金额
    V_供药单位ID 应付记录.单位ID%TYPE;
BEGIN
	If 操作标志_IN=1 Then    --未冲销单据
    --取是否付款及总额
    Begin
  		Select max(付款序号),sum(nvl(发票金额,0)) INTO v_付款序号,v_发票金额
  		FROM 应付记录
  		WHERE 收发id=(Select ID From 药品收发记录 Where NO=NO_IN And 序号=序号_IN And 单据=1) AND 系统标识=1 And 记录性质=-1;
  	EXCEPTION
  		WHEN OTHERS THEN
  		v_发票金额:=0;
  		NULL;
  	END ;
  	v_付款序号:=nvl(v_付款序号,0);
  	IF v_付款序号<>0 then
  	   ErrInfor:='[ZLSOFT]该单据已经被付了款，不能再修改发票信息[ZLSOFT]';
  	   RAISE ErrItem;
  	END IF ;
  	if 发票金额_IN>v_发票金额 And v_发票金额<>0 then
  		ErrInfor:='[ZLSOFT]发票金额不能大于计划付款金额[ZLSOFT]';
  		raise ErrItem;
  	end if ;

  	SELECT A.ID, nvl(B.发票金额,0), A.供药单位ID
        INTO V_收发ID, V_发票金额, V_供药单位ID
        FROM 药品收发记录 A, (Select * From 应付记录 Where 系统标识=1 And 记录性质=0 AND 记录状态=1 And 付款序号 Is NULL) B
       WHERE A.ID = B.收发ID(+)
          AND A.NO = NO_IN
          AND A.单据 = 1
          AND A.记录状态 = 1
          AND A.序号 = 序号_IN;

  	UPDATE 应付记录
  	SET 发票号 = 发票号_IN,
  		发票日期 = 发票日期_IN,
  		发票金额 = 发票金额_IN,
  		单位ID=供药单位_IN
  	WHERE 收发ID = V_收发ID And 系统标识=1 And 记录状态=1 And 记录性质=0;

  	if sql%rowcount=0 then
  		IF 发票号_IN IS NOT NULL THEN
  			--如果是第一笔明细,则产生应付记录的NO
  			BEGIN
  				SELECT NO INTO V_NO FROM 应付记录
  				WHERE 系统标识=1 AND 记录性质=0 AND 记录状态=1
  					AND 入库单据号=NO_IN AND ROWNUM<2;
  			EXCEPTION
  				WHEN OTHERS THEN V_NO:=NEXTNO(67);
  			END ;

  			SELECT 应付记录_ID.NEXTVAL INTO V_应付ID FROM DUAL;
  			INSERT INTO 应付记录
  			(ID,记录性质,记录状态,单位ID,NO,系统标识,收发ID,入库单据号,单据金额,发票号,发票日期,发票金额,品名,
  			规格,产地,批号,计量单位,数量,采购价,采购金额,填制人,填制日期,审核人,审核日期,摘要,项目ID,序号)
  			select V_应付ID,0,1,供药单位_IN,V_NO,1,V_收发ID,A.NO,A.零售金额,发票号_IN,发票日期_IN,发票金额_IN,B.名称,
  			B.规格,B.产地,A.批号,B.计算单位,A.实际数量,A.成本价,A.成本金额,A.填制人,A.填制日期,A.审核人,A.审核日期,A.摘要,A.药品ID,A.序号
  			from 药品收发记录 A,收费项目目录 B
  			Where A.单据=1 And A.NO=NO_in And A.序号=序号_IN And A.药品ID=B.ID;
  		END IF;
  	END IF;

      UPDATE 应付余额
          SET 金额 = NVL (金额, 0) - V_发票金额
      WHERE 单位ID = V_供药单位ID AND 性质 = 1;
      IF SQL%NOTFOUND THEN
          INSERT INTO 应付余额(单位ID, 性质, 金额)
          VALUES (V_供药单位ID, 1,-V_发票金额);
      END IF;
  	UPDATE 应付余额
          SET 金额=NVL(金额,0)+发票金额_IN
      WHERE 单位ID=供药单位_IN AND 性质=1;
      IF SQL%NOTFOUND THEN
          INSERT INTO 应付余额(单位ID, 性质, 金额)
          VALUES (供药单位_IN, 1,发票金额_IN);
      END IF;

      --更新药品收发记录中的供药单位
      UPDATE 药品收发记录 SET 供药单位ID=供药单位_IN WHERE NO=NO_IN AND 单据=1 And 序号=序号_IN;

      --更新药品库存里的上次供应商
      UPDATE 药品库存 SET 上次供应商ID=供药单位_IN WHERE (库房ID,药品ID) IN (SELECT 库房ID,药品ID FROM 药品收发记录 WHERE NO=NO_IN AND 单据=1) AND 性质=1;
   Else      --部分冲销单据，只更新发票号
      Update 应付记录 Set 发票号=发票号_IN Where 入库单据号=NO_IN And 序号=序号_IN;
   End If;
EXCEPTION
	when ErrItem then Raise_application_error (-20101, ErrInfor);
    WHEN NO_data_found THEN
        Raise_application_error (-20101, '[ZLSOFT]该单据已经被他人冲销或已经付过款！[ZLSOFT]');
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品外购发票信息_UPDATE;
/

CREATE OR REPLACE PROCEDURE zl_BillCopy (
    单据_IN IN 药品收发记录.单据%TYPE,
    NO_IN IN 药品收发记录.NO%TYPE,
    NewNO_IN IN 药品收发记录.NO%TYPE
)
IS
	V_NO 应付记录.NO%TYPE;
	V_应付ID 应付记录.ID%TYPE;
BEGIN
    --复制产生新单据(可能存在对已部分冲销的单据进行财务审核,需要更正数量,而采购价、采购金额及差价由其它过程更新)
    INSERT INTO 药品收发记录
    (ID,记录状态,单据,NO,序号,库房ID,供药单位ID,入出类别ID,入出系数,药品ID,
    批次,产地,批号,效期,填写数量,实际数量,成本价,成本金额,扣率,零售价,零售金额,
    差价,摘要,填制人,填制日期,配药人,配药日期,单量,外观,产品合格证,频次,生产日期,批准文号)
    SELECT 药品收发记录_ID.NEXTVAL ID,1 记录状态,A.单据,A.NO,A.序号,A.库房ID,A.供药单位ID,A.入出类别ID,A.入出系数,A.药品ID,
    A.批次,A.产地,A.批号,A.效期,B.实际数量,B.实际数量,A.成本价,B.成本金额,A.扣率,A.零售价,B.零售金额,
    B.差价,A.摘要,A.填制人,A.填制日期,配药人,配药日期,单量,外观,产品合格证,to_char(Sysdate,'YYYY-MM-DD HH24:MI:SS'),A.生产日期,A.批准文号
    FROM 
        (SELECT 单据,NEWNO_IN NO,序号,库房ID,供药单位ID,入出类别ID,入出系数,药品ID,
        批次,产地,批号,效期,成本价,成本金额,扣率,零售价,零售金额,
        差价,摘要,填制人,填制日期,配药人,配药日期,单量,外观,产品合格证,生产日期,批准文号
        FROM 药品收发记录
        WHERE 单据=单据_IN AND NO=NO_IN AND (记录状态=1 OR MOD(记录状态,3)=0)) A,
        (SELECT 序号,SUM(实际数量) 实际数量,Sum(零售金额) 零售金额,Sum(差价) 差价,Sum(成本金额) 成本金额
        FROM 药品收发记录
        WHERE 单据=单据_IN AND NO=NO_IN 
        GROUP BY 序号) B
    WHERE A.序号=B.序号;

    IF 单据_IN=1 THEN 
		V_NO:=NEXTNO(67);
		
		INSERT INTO 应付记录
			(ID,记录性质,记录状态,单位ID,NO,系统标识,收发ID,入库单据号,单据金额,发票号,发票日期,发票金额,品名,
			规格,产地,批号,计量单位,数量,采购价,采购金额,填制人,填制日期,审核人,审核日期,摘要,项目ID,序号)
		SELECT 应付记录_ID.NEXTVAL,0,1,C.单位ID,V_NO,1,A.收发ID,NewNO_IN,B.零售金额,C.发票号,C.发票日期,C.发票金额,B.名称,
			B.规格,B.产地,B.批号,B.售价单位, B.实际数量,B.成本价,B.成本金额,B.填制人,B.填制日期,B.审核人,B.审核日期,B.摘要,B.药品ID,B.序号
   FROM (SELECT Id 收发ID,序号 FROM 药品收发记录 WHERE 单据=单据_IN AND NO=NEWNO_IN) A,
			(SELECT A.No,A.药品ID,A.序号,sum(A.零售金额) 零售金额,A.批号,sum(A.实际数量) 实际数量,A.成本价,sum(A.成本金额) 成本金额,
              A.填制人,min(A.填制日期) 填制日期,A.审核人,min(A.审核日期) 审核日期,A.摘要 ,C.规格,C.产地,C.名称,C.计算单位 AS 售价单位
			FROM 药品收发记录 A,收费项目目录 C 
			WHERE A.单据=单据_IN AND A.NO=NO_IN AND A.药品ID=C.Id
      Group By A.No,A.药品ID,A.序号,A.批号,A.成本价,A.填制人,A.审核人,A.摘要 ,C.规格,C.产地,C.名称,C.计算单位) B,
			(Select 入库单据号,单位ID,发票号,发票日期,SUM(发票金额) 发票金额 ,序号
			From 应付记录 
			Where 系统标识=1 And 记录性质=0
			GROUP BY 入库单据号,单位ID,发票号,发票日期,序号) C
		WHERE A.序号=B.序号 AND B.No=C.入库单据号 And b.序号=c.序号;
    
 END IF ;
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
        NULL;
END zl_BillCopy;
/

CREATE OR REPLACE PROCEDURE zl_药品收发记录_部门退药 (
    BillID_IN IN 药品收发记录.ID%TYPE,
    People_IN IN 药品收发记录.审核人%TYPE,
    Date_IN IN 药品收发记录.审核日期%TYPE,
    批号_IN IN 药品库存.上次批号%TYPE:=NULL,
    效期_IN IN 药品库存.效期%TYPE:=NULL,
    产地_IN IN 药品库存.上次产地%TYPE:=NULL,
    退药数量_IN IN 药品收发记录.实际数量%TYPE:=NULL,
    退药库房_IN IN 药品收发记录.库房ID%TYPE:=Null,
    退药人_IN In 药品收发记录.领用人%TYPE:=Null
)
IS
    --只读变量
    int记录状态 药品收发记录.记录状态%TYPE;
    int执行状态 病人费用记录.执行状态%TYPE;
    bln部分退药 NUMBER;
    lng入出类别ID NUMBER (18);
    strNO 药品收发记录.NO%TYPE;
    int单据 药品收发记录.单据%TYPE;
    lng库房ID 药品收发记录.库房ID%TYPE;
    lng药品ID 药品收发记录.药品ID%TYPE;
    Dbl实际数量 药品收发记录.实际数量%TYPE;
    Dbl实际金额 药品收发记录.零售金额%TYPE;
    Dbl实际成本 药品收发记录.成本金额%TYPE;
    Dbl实际差价 药品收发记录.差价%TYPE;
    lng费用ID 药品收发记录.费用ID%TYPE;
    BillNO NUMBER (8);        --调价单号
    dbl原价 药品收发记录.零售价%TYPE;
    dbl现价 药品收发记录.零售价%TYPE;

    --20020731 Modified by zyb
    --处理退药时，分批核算性质改变后的处理
    lng新批次 药品收发记录.批次%TYPE;
    lng分批 药品规格.药房分批%TYPE;
    lng批次 药品收发记录.批次%TYPE;			--原批次
	str批号 药品收发记录.批号%TYPE;			--原批号
	date效期 药品收发记录.效期%TYPE;		--原效期

    intDigit number(1);
    Err_custom Exception;
    v_Error Varchar2(255);
BEGIN
    --获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';

	IF 退药数量_IN IS NOT NULL THEN
        IF 退药数量_IN=0 THEN
            RETURN;
        END IF ;
    END IF ;
    --获取该收发记录的单据、药品ID、库房ID
    SELECT 单据,NO,库房ID,药品ID,费用ID,入出类别ID,记录状态,Nvl(批次,0),批号,效期
    INTO int单据,strNO,lng库房ID,lng药品ID,lng费用ID,lng入出类别ID,int记录状态,lng批次,str批号,date效期
    FROM 药品收发记录
    WHERE ID = BillID_IN;
    --获取该笔记录剩余未退数量、金额及差价
    --尽量避免金额及差价未出完的现象
    SELECT SUM(NVL(实际数量,0)*NVL(付数,1)),SUM(NVL(零售金额,0)),SUM(NVL(成本金额,0)),SUM(NVL(差价,0))
    INTO Dbl实际数量,Dbl实际金额,Dbl实际成本,Dbl实际差价
    FROM 药品收发记录
    WHERE 审核人 IS NOT NULL AND NO=strNO AND 单据=int单据
    AND 序号=(SELECT 序号 FROM 药品收发记录 WHERE ID=BillID_IN);

    --如果允许退药数为零，表示已退药
    IF Dbl实际数量=0 THEN
        v_Error:='该单据已被其他操作员退药，请刷新后再试！';
        RAISE Err_custom;
    END IF ;
    IF NVL(退药数量_IN,0)>Dbl实际数量 THEN
        v_Error:='该单据已被其他操作员部分退药，请刷新后再试！';
        RAISE Err_custom;
    END IF ;

    --获取该药品当前是否分批的信息
    SELECT Nvl(药房分批,0)  INTO lng分批
    FROM 药品规格
    WHERE 药品ID=lng药品ID;
    --如果是部分退药，则重新计算零售金额及差价
    bln部分退药:=0;
    IF NOT (退药数量_IN IS NULL OR NVL(退药数量_IN,0)=Dbl实际数量) THEN
        bln部分退药:=1;
    END IF ;
    IF bln部分退药=1 THEN
        Dbl实际金额:=ROUND(Dbl实际金额*退药数量_IN/Dbl实际数量,INTDIGIT);
        Dbl实际成本:=ROUND(Dbl实际成本*退药数量_IN/Dbl实际数量,INTDIGIT);
        Dbl实际差价:=ROUND(Dbl实际差价*退药数量_IN/Dbl实际数量,INTDIGIT);
        Dbl实际数量:=退药数量_IN;
    END IF ;

    --lng分批:0-不分批;1-分批;2-原分批，现不分批，按不分批处理;3-原不分批，现分批，产生新批次
    IF lng分批=0 AND lng批次<>0 THEN
        --原分批，现不分批，按不分批处理
        lng分批:=2 ;
    ELSIF lng分批<>0 AND lng批次=0 THEN
        --原不分批,现分批,产生新的批次，并在新产生的发药记录中使用
        lng分批:=3;
    ELSE
        IF lng批次=0 THEN
            lng分批:=0;
        ELSE
            lng分批:=1;
        END IF ;
    END IF ;

    --记录状态的含义有所变化
    --冲销的记录状态        :iif(int记录状态=1,0,1)+1
    --被冲销的记录状态        :iif(int记录状态=1,0,1)+2
    --等待发药的记录状态    :iif(int记录状态=1,0,1)+3

    --产生冲销记录
    Insert INTO 药品收发记录
    (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,
    批次,产地,批号,效期,付数,填写数量,实际数量,成本价,成本金额,扣率,零售价,
    零售金额,差价,摘要,填制人,填制日期,配药人,审核人,审核日期,费用ID,单量,频次,用法,发药窗口,外观,领用人)
    SELECT 药品收发记录_ID.Nextval, int记录状态+DECODE(int记录状态,1,0,1)+1, int单据, strNO, 序号,
        库房ID,对方部门ID, 入出类别ID, 入出系数, 药品ID, 批次, 产地, 批号, 效期,
        1, -Dbl实际数量, -Dbl实际数量, 成本价, -Dbl实际成本, 扣率, 零售价,
        -Dbl实际金额, -Dbl实际差价, 摘要, People_IN, Date_IN, 配药人, People_IN,
        Date_IN,费用ID,单量,频次,用法,发药窗口,退药库房_IN,退药人_IN
    FROM 药品收发记录
    WHERE ID = BillID_IN;

    --如果是部分冲销，则付数填为1，实际数量为付数与实际数量的积
    --产生正常记录以供继续发药
    SELECT 药品收发记录_ID.Nextval INTO lng新批次 FROM dual;
    Insert INTO 药品收发记录
    (ID,记录状态,单据,NO,序号,库房ID,对方部门ID,入出类别ID,入出系数,药品ID,
    批次,产地,批号,效期,付数,填写数量,实际数量,成本价,成本金额,扣率,零售价,
    零售金额,差价,摘要,填制人,填制日期,配药人,审核人,审核日期,费用ID,单量,频次,用法,发药窗口)
    SELECT lng新批次, int记录状态+DECODE(int记录状态,1,0,1)+3, int单据, strNO, 序号,
         库房ID, 对方部门ID, 入出类别ID, 入出系数, 药品ID, DECODE(lng分批,1,批次,3,lng新批次,NULL), DECODE(lng分批,3,产地_IN,1,产地,产地),
         DECODE(lng分批,3,批号_IN,1,批号,NULL),DECODE(lng分批,3,效期_IN,1,效期,NULL),1,
         Dbl实际数量,Dbl实际数量,成本价,Dbl实际成本,扣率,零售价,Dbl实际金额,
         Dbl实际差价,摘要,填制人,填制日期,NULL,NULL,NULL,费用ID,单量,频次,用法,发药窗口
    FROM 药品收发记录
    WHERE ID = BillID_IN;

    --更新病人费用记录的执行状态(0-未执行;1-完全执行;2-部分执行)
    SELECT DECODE(SUM(NVL(付数,1)*实际数量),NULL,0,0,0,2) INTO int执行状态
    FROM 药品收发记录
    WHERE 单据=int单据 AND No=strNO AND 费用id=lng费用ID AND 审核人 IS NOT NULL;
    UPDATE 病人费用记录
    SET 执行状态 = int执行状态
    WHERE ID =lng费用ID;

    --插入未发药品记录
    BEGIN
        Insert INTO 未发药品记录
        (单据,NO,病人ID,主页ID,姓名,优先级,对方部门ID,
        库房ID,发药窗口,填制日期,已收费,配药人,打印状态,未发数)
        SELECT A.单据, A.NO, A.病人ID, A.主页ID, A.姓名,NVL (B.优先级, 0) 优先级,
            A.对方部门ID, A.库房ID, A.发药窗口, A.填制日期, A.已收费,NULL, 1, 1
        FROM (
            SELECT B.单据, B.NO, A.病人ID, A.主页ID, A.姓名,
            DECODE (A.记录状态,0,0,1) 已收费,
            B.对方部门ID, B.库房ID, B.发药窗口, B.填制日期,C.身份
            FROM 病人费用记录 A, 药品收发记录 B,病人信息 C
            WHERE B.ID = BillID_IN
            AND A.ID = B.费用ID+0 And A.病人ID=C.病人ID(+)) A,身份 B
            Where B.名称(+)=A.身份;
    EXCEPTION
        WHEN OTHERS THEN NULL;
    END;

    --修改原记录为被冲销记录
    UPDATE 药品收发记录
    SET 记录状态 = int记录状态+DECODE(int记录状态,1,0,1)+2
    WHERE ID = BillID_IN;

    --修改药品库存(反冲库存)
    IF lng分批<>3 THEN
        UPDATE 药品库存
        SET 实际数量 = NVL (实际数量, 0) + Dbl实际数量,
            实际金额 = NVL (实际金额, 0) + Dbl实际金额,
            实际差价 = NVL (实际差价, 0) + Dbl实际差价
        WHERE 库房ID+0 = lng库房ID AND 药品ID = lng药品ID AND 性质 = 1 AND NVL(批次,0)=lng批次;

        IF SQL%ROWCOUNT = 0 THEN
            INSERT INTO 药品库存
            (库房ID,药品ID,批次,性质,实际数量,实际金额,实际差价,上次批号,效期)
            VALUES
            (lng库房ID,lng药品ID,DECODE(lng分批,2,NULL,lng批次),1,Dbl实际数量,Dbl实际金额,Dbl实际差价,DECODE(lng分批,1,str批号,NULL),DECODE(lng分批,1,date效期,NULL));
        END IF;
    ELSE
        INSERT INTO 药品库存
        (库房ID,药品ID,批次,效期,性质,实际数量,实际金额,实际差价,上次批号,上次产地)
        VALUES
        (lng库房ID,lng药品ID,lng新批次,效期_IN,1,dbl实际数量,Dbl实际金额,dbl实际差价,批号_IN,产地_IN);
    END IF ;

    DELETE 药品库存
    WHERE 库房ID+0 = lng库房ID AND 药品ID = lng药品ID AND 性质=1
    AND NVL(可用数量,0) = 0 AND NVL(实际数量,0) = 0 AND NVL(实际金额,0) = 0 AND NVL(实际差价,0) = 0;

    --更新药品收发汇总
    UPDATE 药品收发汇总
    SET 数量 = NVL (数量, 0) + Dbl实际数量 ,
        金额 = NVL (金额, 0) + Dbl实际金额 ,
        差价 = NVL (差价, 0) + Dbl实际差价
    WHERE 库房ID+0 = lng库房ID AND 药品ID+0 = lng药品ID AND 类别ID+0 = lng入出类别ID
    AND 日期 = TRUNC (Date_IN) AND 单据 = int单据;

    IF SQL%ROWCOUNT = 0 THEN
        Insert INTO 药品收发汇总
        (日期, 库房ID, 药品ID, 单据, 类别ID, 数量, 金额, 差价)
        VALUES
        (TRUNC (Date_IN),lng库房ID,lng药品ID,int单据,lng入出类别ID,Dbl实际数量 ,Dbl实际金额 ,Dbl实际差价 );
    END IF;

    DELETE 药品收发汇总
    WHERE 库房ID+0 = lng库房ID AND 药品ID+0 = lng药品ID AND 类别ID+0 = lng入出类别ID
    AND 日期 = TRUNC (Date_IN) AND 单据 = int单据
    AND Nvl(数量,0)=0 AND Nvl(金额,0)=0 And Nvl(差价,0)=0;
    
    --处理调价退药
    select nvl(零售价,0) 原价,现价 Into dbl原价,dbl现价
    from 药品收发记录 a,收费价目 b
    where a.药品id=b.收费细目id  And (SYSDATE BETWEEN b.执行日期 AND b.终止日期 Or  SYSDATE >= b.执行日期 AND b.终止日期 IS Null) 
    And a.id=BillID_IN;
        
    If dbl原价<>dbl现价 Then
       SELECT 药品收发记录_ID.Nextval INTO BillNO FROM Dual; 
       
       SELECT 类别ID INTO lng入出类别ID FROM 药品单据性质 WHERE 单据 = 13;
       
       Insert INTO 药品收发记录 ( ID,记录状态,单据,NO,序号,入出类别ID,药品ID,批次,批号,效期,产地,付数,填写数量,实际数量,成本价,成本金额, 
                                   零售价,扣率,零售金额,差价,摘要,填制人,填制日期,库房ID,入出系数,价格ID,审核人,审核日期,费用ID) 
       Select  药品收发记录_ID.Nextval,记录状态,13,BillNO,序号,lng入出类别ID,药品ID,批次,批号,效期,产地,付数,Dbl实际数量,0,成本价,0, 
                                   dbl现价-dbl原价,扣率,(dbl现价-dbl原价)*Dbl实际数量,(dbl现价-dbl原价)*Dbl实际数量,'调价退药',People_IN,Sysdate,库房ID,1,价格ID,People_IN,Sysdate,BillID_IN
       From 药品收发记录 Where id=BillID_IN;
    End If;
       
EXCEPTION
    When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_药品收发记录_部门退药;
/

CREATE OR REPLACE PROCEDURE zl_药品储备限额_Update (
    库房ID_IN In 药品储备限额.库房ID%Type,
    药品ID_IN IN 药品储备限额.药品ID%Type,
    上限_IN In 药品储备限额.上限%TYPE := 0,
    下限_IN In 药品储备限额.下限%TYPE := 0,
    盘点_IN In 药品储备限额.盘点属性%TYPE := '0000',
    货位_IN In 药品储备限额.库房货位%TYPE := NULL 
) IS
    v_类别 收费项目目录.类别%TYPE;
BEGIN
    IF 上限_IN<>0 or 下限_IN<>0 or 盘点_IN<>'0000' and 盘点_IN is not null or 货位_IN is not null THEN
       update 药品储备限额
       set 上限=上限_IN,下限=下限_IN,盘点属性=nvl(盘点_IN,'0000'),库房货位=货位_IN
       where 库房ID=库房ID_IN and 药品ID=药品ID_IN;
       if sql%rowcount=0 then
          INSERT INTO 药品储备限额 (库房ID,药品ID,上限,下限,盘点属性,库房货位)
          values(库房ID_IN,药品ID_IN,上限_IN,下限_IN,nvl(盘点_IN,'0000'),货位_IN);
       end if;
    Else
       Delete 药品储备限额
       where 库房ID=库房ID_IN and 药品ID=药品ID_IN;
    END IF;
    --删除已经修改性质的非库房限额
    select 类别 into v_类别 from 收费项目目录 where ID=药品ID_IN;
    if v_类别='5' then
       delete 药品储备限额
       where 药品ID=药品ID_IN
             and 库房id not in (
                 select distinct 部门id from 部门性质说明 where 工作性质 like '西药%' or 工作性质='制剂室');
    end if;
    if v_类别='6' then
       delete 药品储备限额
       where 药品ID=药品ID_IN
             and 库房id not in (
                 select distinct 部门id from 部门性质说明 where 工作性质 like '成药%' or 工作性质='制剂室');
    end if;
    if v_类别='7' then
       delete 药品储备限额
       where 药品ID=药品ID_IN
             and 库房id not in (
                 select distinct 部门id from 部门性质说明 where 工作性质 like '中药%' or 工作性质='制剂室');
    end if;

      --增加新库房货位  
	If 货位_In Is Not Null Then
		Update 药品库房货位 Set 名称 = 货位_In Where 名称 = 货位_In;
		If Sql%Rowcount = 0 Then
			Insert Into 药品库房货位
				(编码, 名称, 简码)
				Select trim(to_char(Nvl(Max(To_Number(编码)), 0) + 1,'00000')), 货位_In, Zlspellcode(货位_In) From 药品库房货位;
		End If;
	End If;

EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品储备限额_Update;
/

CREATE OR REPLACE Procedure ZL_病人医嘱执行_Cancel(
	医嘱ID_IN		病人医嘱执行.医嘱ID%Type,
	发送号_IN		病人医嘱执行.发送号%Type,
	取消皮试_IN		Number:=Null
--参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
) IS
	Cursor c_Advice is
		Select A.ID,A.相关ID,A.病人ID,A.主页ID,A.诊疗类别,B.操作类型 
			From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID And A.ID=医嘱ID_IN;
	r_Advice c_Advice%RowType;

    v_Temp            Varchar2(255);
    v_人员编号        病人费用记录.操作员编号%Type;
    v_人员姓名        病人费用记录.操作员姓名%Type;
	
	v_Date		Date;
	v_Count		Number;
Begin
	--对检验组合,执行情况只填写在了第一个检验项目中
	Select Count(*) Into v_Count From 病人医嘱执行 Where 医嘱ID=医嘱ID_IN And 发送号+0=发送号_IN;

	Open c_Advice;
	Fetch c_Advice Into r_Advice;

	If r_Advice.诊疗类别='C' And r_Advice.相关ID IS Not NULL Then
		--包含一并采集的所有检验项目
		Update 病人医嘱发送
			Set 执行状态=Decode(v_Count,0,0,3) 
		Where 发送号+0=发送号_IN And 医嘱ID IN(
			Select ID From 病人医嘱记录 Where 相关ID=r_Advice.相关ID);
	Else
		--包含附加手术,检验部位,以及其它独立医嘱;麻醉和中药煎法是单独安排
		Update 病人医嘱发送
			Set 执行状态=Decode(v_Count,0,0,3)
		Where 发送号+0=发送号_IN And 医嘱ID IN(
			Select ID From 病人医嘱记录 Where ID=医嘱ID_IN
			Union ALL
			Select ID From 病人医嘱记录 Where 相关ID=医嘱ID_IN And 诊疗类别 IN('F','D'));
	End If;

	--主费用可能需要限制医嘱序号
	Update 病人费用记录 
		Set 执行状态=0,执行时间=NULL,执行人=NULL
	Where 收费类别 Not IN('5','6','7') And 医嘱序号+0=医嘱ID_IN
		And (记录性质,NO) IN(
			Select 记录性质,NO From 病人医嘱附费 Where 医嘱ID=医嘱ID_IN And 发送号+0=发送号_IN
			Union ALL
			Select 记录性质,NO From 病人医嘱发送 Where 医嘱ID=医嘱ID_IN And 发送号+0=发送号_IN);

	--删除过敏登记记录(当前人员登记的)
	If r_Advice.诊疗类别='E' And r_Advice.操作类型='1' Then
		v_Temp:=zl_Identity;
		v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
		v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
		v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
		v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);
		
		Begin
			Select Max(操作时间) Into v_Date From 病人医嘱状态 
			Where 医嘱ID=医嘱ID_IN And 操作类型=10 
				And (操作人员=v_人员姓名 Or Nvl(取消皮试_IN,0)=1);
		Exception
			When Others Then Null;
		End;
		If v_Date IS Not Null Then
			Delete From 病人过敏记录 
			Where 病人ID=r_Advice.病人ID And 记录来源=2
				And Nvl(主页ID,0)=Nvl(r_Advice.主页ID,0)
				And 记录时间=v_Date And (记录人=v_人员姓名 Or Nvl(取消皮试_IN,0)=1);
			If SQL%RowCount>0 Then
				Update 病人医嘱记录 Set 皮试结果=Null Where ID=医嘱ID_IN;
			End IF;
		End If;
	End IF;

	Close c_Advice;			
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱执行_Cancel;
/

CREATE OR REPLACE Procedure ZL_病人医嘱记录_回退(
--功能：回退住院医嘱的状态操作或发送操作
    医嘱ID_IN       病人医嘱记录.ID%Type,
    FLAG_IN			Number:=0,
	医嘱内容_IN		病人医嘱记录.医嘱内容%Type:=Null,
	操作类型_IN		病人医嘱状态.操作类型%Type:=Null
--参数：医嘱ID_IN=相关ID为空的医嘱的ID(给药途径,中药用法,主要手术,检查项目,及独立医嘱),相当于医嘱组ID
--      FLAG_IN=附加数据。回退停止：0=清除执行终止时间,1=保留现有的执行终止时间。
--      医嘱内容_IN=该过程被批量回退调用时才用，用于错误提示。
--      操作类型_IN=该过程被批量回退调用时才用，用于核对回退数据。0-回退发送,n=回退具体医嘱操作
) IS
    --包含指定医嘱的操作记录,第一条为要回退的内容(状态操作优先)
	--临嘱不回退发送后的自动停止,在回退发送时自动回退停止操作
    Cursor c_RollAdvice is
        Select Distinct B.操作人员,B.操作时间,0 AS 发送号,B.操作类型,
            0 AS 执行状态,Sysdate+NULL AS 首次时间,Sysdate+NULL AS 末次时间,
			A.上次执行时间,A.医嘱期效,A.诊疗类别 AS 类别,Null AS 类型,
			A.病人ID,A.主页ID,A.婴儿,A.皮试结果
        From 病人医嘱记录 A,病人医嘱状态 B
        Where A.ID=B.医嘱ID And (A.ID=医嘱ID_IN Or A.相关ID=医嘱ID_IN)
            And (Nvl(A.医嘱期效,0)=0 And B.操作类型 Not IN(1,2,3) 
				Or Nvl(A.医嘱期效,0)=1 And B.操作类型 Not IN(1,2,3,8))
        Union ALL
        Select Distinct B.发送人 AS 操作人员,B.发送时间 AS 操作时间,
			B.发送号,-NULL as 操作类型,B.执行状态,B.首次时间,B.末次时间,
			A.上次执行时间,A.医嘱期效,C.类别,C.操作类型 AS 类型,
			A.病人ID,A.主页ID,A.婴儿,A.皮试结果
        From 病人医嘱记录 A,病人医嘱发送 B,诊疗项目目录 C
        Where A.ID=B.医嘱ID And A.诊疗项目ID=C.ID
			And (A.ID=医嘱ID_IN Or A.相关ID=医嘱ID_IN)
        Order by 操作时间 Desc,发送号;
    r_RollAdvice c_RollAdvice%RowType;
    
    --根据医嘱及发送NO求出本次回退要销帐的费用记录
    --一组医嘱并不是都填写了发送记录,且可能NO不同(药品有,用法煎法不一定有)
    --不管发送记录的计费状态(可能无需计费),有费用记录自然关联出来
    --费用只求价格父号为空的,以便取序号销帐
    --只管记录状态为1的费用,对于已销帐或部份销帐的记录,不再处理
    Cursor c_RollMoney(v_发送号 病人医嘱发送.发送号%Type) is
        Select A.NO,A.序号,A.执行状态
        From 病人费用记录 A,病人医嘱记录 B,病人医嘱发送 C
        Where C.医嘱ID=B.ID And C.发送号=v_发送号
            And (B.ID=医嘱ID_IN Or B.相关ID=医嘱ID_IN)
            And A.医嘱序号=B.ID And A.记录状态 IN(0,1)
            And A.NO=C.NO And A.记录性质=C.记录性质
            And A.价格父号 IS NULL
        Order BY A.收费细目ID;
	
	--用于删除报告记录
	Cursor c_Case(v_发送号 病人医嘱发送.发送号%Type) is
		Select 报告ID From 病人医嘱发送 
		Where 报告ID IS Not NULL And 发送号=v_发送号 
			And 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=医嘱ID_IN OR 相关ID=医嘱ID_IN);
	r_Case c_Case%RowType;

	--用于处理特殊医嘱的回退
	Cursor c_PatiLog(
		v_病人ID 病人变动记录.病人ID%Type,
		v_主页ID 病人变动记录.主页ID%Type) is
		Select * From 病人变动记录 
		Where 病人ID=v_病人ID And 主页ID=v_主页ID And 终止时间 IS NULL
		Order by 开始时间 Desc;
	r_PatiLog c_PatiLog%RowType;

    v_医嘱状态      病人医嘱记录.医嘱状态%Type;
    v_费用NO        病人费用记录.NO%Type;
    v_费用序号      Varchar2(255);
    v_末次时间      病人医嘱发送.末次时间%Type;

    v_Count         Number(5);
    v_Temp          Varchar2(255);
    v_人员编号      病人费用记录.操作员编号%Type;
    v_人员姓名      病人费用记录.操作员姓名%Type;

    v_Error         Varchar2(255);
    Err_Custom      Exception;
Begin
	Open c_RollAdvice;
    Fetch c_RollAdvice Into r_RollAdvice;
    If c_RollAdvice%RowCount=0 Then
        v_Error:=Nvl(医嘱内容_IN,'该医嘱')||'当前没有可以回退的内容。';
        Raise Err_Custom;
    End IF;
	--批量回退调用时判断
	If 医嘱内容_IN Is Not Null Then
		If Nvl(r_RollAdvice.操作类型,0)<>Nvl(操作类型_IN,0) Then
			v_Error:=Nvl(医嘱内容_IN,'该医嘱')||'不能与当前医嘱一起回退，可能该医嘱已经执行了其他操作。';
			Raise Err_Custom;
		End IF;
	End IF;

    If r_RollAdvice.发送号=0 Then
        --回退医嘱状态操作(以时间关键字)
        --4-作废；5-重整；6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果
        ------------------------------------------------------------------

        --最多只能退回到校对状态
        If r_RollAdvice.操作类型=3 Then
            v_Error:=Nvl(医嘱内容_IN,'该医嘱')||'当前处于通过校对状态，不能再回退。';
            Raise Err_Custom;
        End IF;
        
        --回退作废时恢复申请单的作废
        If r_RollAdvice.操作类型=4 Then
			Update 病人病历记录
				Set 作废人=NULL,作废日期=NULL
			Where 作废人=r_RollAdvice.操作人员 And 作废日期=r_RollAdvice.操作时间
				And ID IN(Select 申请ID From 病人医嘱记录 Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN);
        End IF;

        --删除(该组医嘱)最近的状态操作记录
        Delete From 病人医嘱状态 
        Where 医嘱ID IN (Select ID From 病人医嘱记录 Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN) 
            And 操作时间=r_RollAdvice.操作时间;

        --取删除后应恢复的医嘱状态
        Select 操作类型
            Into v_医嘱状态 
        From 病人医嘱状态 
        Where 操作时间=(Select Max(操作时间) From 病人医嘱状态 Where 医嘱ID=医嘱ID_IN)
            And 医嘱ID=医嘱ID_IN;
        
        --恢复(该组医嘱)回退后的状态
        Update 病人医嘱记录 
            Set 医嘱状态=v_医嘱状态 
        Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;
        
        --其它额外的处理
        If r_RollAdvice.操作类型=8 Then
            --被超期发送收回过的医嘱不允许再撤消停止(撤消停止,再发送,再回退就有问题)
            --可能超期发送收回时被全部收回(无上次执行时间)
            Select Nvl(Count(*),0) Into v_Count
            From 病人医嘱记录 A,病人医嘱发送 B
            Where B.医嘱ID=A.ID And (A.ID=医嘱ID_IN Or A.相关ID=医嘱ID_IN)
                And B.发送号=(Select Max(发送号) From 病人医嘱发送 Where 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN))
                And A.执行终止时间 IS Not NULL
                And ((A.上次执行时间<B.末次时间) 
                    Or (A.上次执行时间 IS NULL And B.末次时间 IS Not NULL));
            If v_Count>0 Then
                v_Error:=Nvl(医嘱内容_IN,'该医嘱')||'已被超期发送收回，不能再撤消停止操作。';
                Raise Err_Custom;
            End if;

            --回退医嘱停止时,清空停嘱医生和时间
            Update 病人医嘱记录 
                Set 执行终止时间=Decode(FLAG_IN,1,执行终止时间,NULL),
                    停嘱医生=NULL,停嘱时间=NULL
            Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;
		ElsIf r_RollAdvice.操作类型=9 Then
            --回退医嘱停止时,清空停嘱医生和时间
            Update 病人医嘱记录 
				Set 确认停嘱时间=NULL
			Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;
        ElsIf r_RollAdvice.操作类型=10 Then
            --回退标注皮试结果,同时删除过敏登记(+)或(-),根据记录时间
			Delete From 病人过敏记录 
			Where 病人ID=r_RollAdvice.病人ID
				And Nvl(主页ID,0)=Nvl(r_RollAdvice.主页ID,0)
				And 记录时间=r_RollAdvice.操作时间;

            Update 病人医嘱记录 
                Set 皮试结果=NULL
            Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;
        End IF;
    Else
        --回退医嘱发送(以发送号关键字)
        ------------------------------------------------------------------
        --被超期收回的长期药品医嘱不允许回退(再退费用就多退了)
		If Nvl(r_RollAdvice.医嘱期效,0)=0 Then
			If r_RollAdvice.上次执行时间 IS Not NULL And r_RollAdvice.末次时间 IS Not NULL Then
				If r_RollAdvice.上次执行时间<r_RollAdvice.末次时间 Then
					v_Error:=Nvl(医嘱内容_IN,'该医嘱')||'最近超期发送的内容已被收回，不能再回退。';
					Raise Err_Custom;
				End IF;
			ElsIF r_RollAdvice.上次执行时间 IS NULL And r_RollAdvice.末次时间 IS Not NULL Then
				--长嘱可能被全部超期收回
				v_Error:=Nvl(医嘱内容_IN,'该医嘱')||'未被发送，或发送的内容已被全部超期收回，不能再回退。';
				Raise Err_Custom;
			End IF;
		End IF;

        If Nvl(r_RollAdvice.执行状态,0) IN(1,3) Then --1-完全执行;3-正在执行
            v_Error:=Nvl(医嘱内容_IN,'该医嘱')||'最近发送的内容已经执行或正在执行，不能回退。';            
            Raise Err_Custom;
        End IF;

        --当前操作人员    
        v_Temp:=zl_Identity;
        v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
        v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
        v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
        v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

        --将该组医嘱的费用销帐(按一组医嘱可能有不同NO处理)
        --如果原始费用已被销帐(或部分销帐),调用过程中有判断
        v_费用NO:=NULL;v_费用序号:=NULL;
        For r_RollMoney In c_RollMoney(r_RollAdvice.发送号) Loop
            If Nvl(v_费用NO,'空')<>r_RollMoney.NO Then
                If v_费用序号 IS Not NULL And v_费用NO IS Not NULL Then
                    v_费用序号:=Substr(v_费用序号,2);
                    zl_住院记帐记录_Delete(v_费用NO,v_费用序号,v_人员编号,v_人员姓名);
                End IF;
                v_费用序号:=NULL;
            End IF;
            v_费用NO:=r_RollMoney.NO;
            v_费用序号:=v_费用序号||','||r_RollMoney.序号;

            If Nvl(r_RollMoney.执行状态,0)<>0 Then
                v_Error:=Nvl(医嘱内容_IN,'该医嘱')||'发送的费用单据"'||r_RollMoney.NO||'"中的内容已被部分或完全执行，不能回退。';
                Raise Err_Custom;
            End IF;
        End Loop;
        If v_费用序号 IS Not NULL And v_费用NO IS Not NULL Then
            v_费用序号:=Substr(v_费用序号,2);
            zl_住院记帐记录_Delete(v_费用NO,v_费用序号,v_人员编号,v_人员姓名);
        End IF;

		Open c_Case(r_RollAdvice.发送号);--必须先打开

        --删除发送记录(该组医嘱的)
        Delete From 病人医嘱发送 
        Where 发送号=r_RollAdvice.发送号
            And 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=医嘱ID_IN OR 相关ID=医嘱ID_IN);
        
        --删除对应的报告单
		Fetch c_Case Into r_Case;
		While c_Case%Found Loop
			Delete From 病人病历记录 Where ID=r_Case.报告ID;
			Fetch c_Case Into r_Case;				
		End Loop;
		Close c_Case;

        --标记(该组医嘱)上次执行时间(以上次发送的末次执行时间)

		--所有长嘱(包括持续性长嘱)发送时都填写了末次时间
		--临嘱可能没有，且只可能发送了一次。
        v_末次时间:=NULL;
        Begin
			--一组医嘱的发送首末时间相同,一并给药是取最小的
            --取相关ID为NULL的医嘱的发送记录的时间
			--但给药途径或中药用法可能未填写发送记录
            Select 末次时间 Into v_末次时间
            From 病人医嘱发送
            Where 医嘱ID IN(
					Select ID From 病人医嘱记录 Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN)
                And 发送号=(
					Select Max(发送号) From 病人医嘱发送 Where 医嘱ID IN(
						Select ID From 病人医嘱记录 Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN)
						)
				And Rownum=1;
        Exception
            When Others Then NULL;
        End;
		Update 病人医嘱记录 
			Set 上次执行时间=v_末次时间
		Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;

		--回退临嘱发送时，同时自动回退停止
		If Nvl(r_RollAdvice.医嘱期效,0)=1 Then
			--删除(该组医嘱)最近的状态操作记录
			Delete From 病人医嘱状态 
			Where 医嘱ID IN (Select ID From 病人医嘱记录 Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN) 
				And 操作时间=r_RollAdvice.操作时间;

			--取删除后应恢复的医嘱状态
			Select 操作类型 Into v_医嘱状态 From 病人医嘱状态 
			Where 操作时间=(Select Max(操作时间) From 病人医嘱状态 Where 医嘱ID=医嘱ID_IN)
				And 医嘱ID=医嘱ID_IN;
			
			--恢复(该组医嘱)回退后的状态
			Update 病人医嘱记录 
				Set 医嘱状态=v_医嘱状态,
					执行终止时间=NULL,
					停嘱医生=NULL,
					停嘱时间=NULL
			Where ID=医嘱ID_IN Or 相关ID=医嘱ID_IN;
		End IF;

		--住院特殊医嘱发送后的回退(3-转科;5-出院;6-转院)
		If r_RollAdvice.类别='Z' And Instr(',3,5,6,',Nvl(r_RollAdvice.类型,'0'))>0 And Nvl(r_RollAdvice.婴儿,0)=0 Then
			Open c_PatiLog(r_RollAdvice.病人ID,r_RollAdvice.主页ID);
			Fetch c_PatiLog Into r_PatiLog;
			If c_PatiLog%Found Then
				If r_RollAdvice.类型='3' And r_PatiLog.开始原因=3 And r_PatiLog.开始时间 Is Null Then
					--取消病人转科状态
					zl_病人变动记录_Undo(r_RollAdvice.病人ID,r_RollAdvice.主页ID,v_人员编号,v_人员姓名);
				ElsIF r_RollAdvice.类型='5' And r_PatiLog.开始原因=10 Then
					--取消病人预出院状态
					zl_病人变动记录_Undo(r_RollAdvice.病人ID,r_RollAdvice.主页ID,v_人员编号,v_人员姓名);
				ElsIF r_RollAdvice.类型='6' And r_PatiLog.开始原因=10 Then
					--取消病人预出院状态
					zl_病人变动记录_Undo(r_RollAdvice.病人ID,r_RollAdvice.主页ID,v_人员编号,v_人员姓名);
				End IF;
			End If;
			Close c_PatiLog;
		End IF;
    End IF;

    Close c_RollAdvice;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱记录_回退;
/

CREATE OR REPLACE Procedure ZL_门诊医嘱发送_Insert(
--功能：填写病人医嘱发送记录
    医嘱ID_IN		病人医嘱发送.医嘱ID%Type,
    发送号_IN       病人医嘱发送.发送号%Type,
    记录性质_IN     病人医嘱发送.记录性质%Type,
    NO_IN           病人医嘱发送.NO%Type,
    记录序号_IN     病人医嘱发送.记录序号%Type,
    发送数次_IN     病人医嘱发送.发送数次%Type,
    首次时间_IN     病人医嘱发送.首次时间%Type,
    末次时间_IN     病人医嘱发送.末次时间%Type,
    发送时间_IN     病人医嘱发送.发送时间%Type,
    执行状态_IN     病人医嘱发送.执行状态%Type,
    执行部门ID_IN   病人医嘱发送.执行部门ID%Type,
    计费状态_IN     病人医嘱发送.计费状态%Type,
    First_IN        Number:=0
--参数：First_IN=表示是否一组医嘱的第一医嘱行,以便处理医嘱相关内容(如成药,配方的第一行,因为给药途径,配方煎法,用法可能为叮嘱不发送)
) IS
    --包含病人及医嘱(一组医嘱中第一行)相关信息的游标
    Cursor c_Advice is
        Select
            Nvl(A.相关ID,A.ID) AS 组ID,A.序号,A.病人ID,A.挂号单,A.婴儿,B.姓名,C.操作类型,
            A.诊疗类别,A.医嘱状态,A.医嘱内容,A.开嘱医生,A.开始执行时间,A.执行时间方案,A.频率次数,A.频率间隔,A.间隔单位
        From 病人医嘱记录 A,病人信息 B,诊疗项目目录 C
        Where A.病人ID=B.病人ID And A.诊疗项目ID=C.ID And A.ID=医嘱ID_IN
        Group BY Nvl(A.相关ID,A.ID),A.序号,A.病人ID,A.挂号单,A.婴儿,B.姓名,C.操作类型,
            A.诊疗类别,A.医嘱状态,A.医嘱内容,A.开嘱医生,A.开始执行时间,A.执行时间方案,A.频率次数,A.频率间隔,A.间隔单位;
    r_Advice c_Advice%RowType;

	Cursor c_Pati(v_病人ID 病人信息.病人ID%Type) is
		Select * From 病人信息 Where 病人ID=v_病人ID;
	r_Pati c_pati%RowType;

    --其它临时变量
    v_Temp			Varchar2(255);
	v_Count			Number;
	v_病人性质		病案主页.病人性质%Type;
    v_人员编号      病人费用记录.操作员编号%Type;
    v_人员姓名      病人费用记录.操作员姓名%Type;

	v_Error         Varchar2(255);
    Err_Custom      Exception;
Begin
    --当前操作人员
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --是一组医嘱的第一行时处理医嘱内容
    If Nvl(First_IN,0)=1 Then
        Open c_Advice;
        Fetch c_Advice Into r_Advice;

        --并发操作检查
        ---------------------------------------------------------------------------------------
        IF Nvl(r_Advice.医嘱状态,0)<>1 Then
            v_Error:='"'||r_Advice.姓名||'"的医嘱"'||r_Advice.医嘱内容||'"已经被其他人发送。'
                ||CHR(13)||CHR(10)||'该病人的医嘱发送失败。请重新读取发送清单再试。';
            Raise Err_Custom;
        End IF;

        --发送后的医嘱处理:临嘱发送后自动停止
        ---------------------------------------------------------------------------------------
        Update 病人医嘱记录
            Set 医嘱状态=8,
                执行终止时间=末次时间_IN,--可能没有
                停嘱时间=发送时间_IN,--要作为发送时间显示
                停嘱医生=v_人员姓名--要作为发送人显示,不同于住院,门诊医嘱无护士操作
        Where ID=r_Advice.组ID Or 相关ID=r_Advice.组ID;

        Insert Into 病人医嘱状态(
            医嘱ID,操作类型,操作人员,操作时间)
        Select
            ID,8,v_人员姓名,发送时间_IN
        From 病人医嘱记录
        Where ID=r_Advice.组ID Or 相关ID=r_Advice.组ID;

        --特殊医嘱的处理
        ---------------------------------------------------------------------------------------
        If r_Advice.诊疗类别='Z' And Nvl(r_Advice.操作类型,'0')<>'0' And Nvl(r_Advice.婴儿,0)=0 Then
            --1-留观;2-住院;
			If Instr(',1,2,',r_Advice.操作类型)>0 And 执行部门ID_IN IS Not NULL Then
				--满足产生新的预约登记的条件：1.当前无预约,2.当前不在院,3-无要求预约时间内的住院记录
				Select Count(*) Into v_Count From 病案主页 Where 病人ID=r_Advice.病人ID And Nvl(主页ID,0)=0;
				If v_Count=0 Then
					Select Count(*) Into v_Count From 病案主页 Where 病人ID=r_Advice.病人ID And 出院日期 IS NULL;
				End IF;
				If v_Count=0 Then
					Select Count(*) Into v_Count From 病案主页 Where 病人ID=r_Advice.病人ID 
						And (入院日期>=r_Advice.开始执行时间 Or 出院日期>=r_Advice.开始执行时间);
				End IF;
				If v_Count=0 Then
					If r_Advice.操作类型='1' Then
						--留观医嘱,将病人在"开始时间"留观到临床执行科室
						Begin
							v_病人性质:=2;
							Select Decode(服务对象,1,1,2) Into v_病人性质 From 部门性质说明 Where 工作性质='临床' And 部门ID=执行部门ID_IN;
						Exception
							When Others Then Null;
						End;
					ElsIf r_Advice.操作类型='2' Then
						--住院医嘱,将病人在"开始时间"登记到临床执行科室
						v_病人性质:=0;
					End IF;
					
					Open c_Pati(r_Advice.病人ID);
					Fetch c_Pati Into r_Pati;

					zl_入院病案主页_Insert(1,v_病人性质,r_Pati.病人ID,r_Pati.住院号,NULL,r_Pati.姓名,r_Pati.性别,r_Pati.年龄,
						r_Pati.费别,r_Pati.出生日期,r_Pati.国籍,r_Pati.民族,r_Pati.学历,r_Pati.婚姻状况,r_Pati.职业,r_Pati.身份,r_Pati.身份证号,
						r_Pati.出生地点,r_Pati.家庭地址,r_Pati.户口邮编,r_Pati.家庭电话,r_Pati.联系人姓名,r_Pati.联系人关系,r_Pati.联系人地址,
						r_Pati.联系人电话,r_Pati.工作单位,r_Pati.合同单位ID,r_Pati.单位电话,r_Pati.单位邮编,r_Pati.单位开户行,r_Pati.单位帐号,
						r_Pati.担保人,r_Pati.担保额,r_Pati.担保性质,执行部门ID_IN,NULL,NULL,NULL,NULL,NULL,r_Advice.开嘱医生,NULL,r_Advice.开始执行时间,
						NULL,NULL,r_Pati.医疗付款方式,NULL,NULL,NULL,NULL,r_Pati.险类,v_人员编号,v_人员姓名,0,NULL);

					Close c_Pati;
				End If;
			End If;
        End IF;

        Close c_Advice;
    End IF;

    --填写发送记录
    ---------------------------------------------------------------------------------------
    Insert Into 病人医嘱发送(
        医嘱ID,发送号,记录性质,NO,记录序号,发送数次,发送人,发送时间,执行状态,执行部门ID,计费状态,首次时间,末次时间)
    Values(
        医嘱ID_IN,发送号_IN,记录性质_IN,NO_IN,记录序号_IN,发送数次_IN,
        v_人员姓名,发送时间_IN,执行状态_IN,执行部门ID_IN,计费状态_IN,
        首次时间_IN,末次时间_IN);
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_门诊医嘱发送_Insert;
/

CREATE OR REPLACE Procedure ZL_病人医嘱记录_更新序号(
--功能：用于新开或修改医嘱后，其它医嘱的序号受影响发生变化
    医嘱ID_IN		病人医嘱记录.ID%TYPE:=Null,--传入这两个参数时，表示更新指定医嘱序号
    序号_IN			病人医嘱记录.序号%TYPE:=Null,
	病人ID_IN		病人医嘱记录.病人ID%Type:=Null,--传入这两个参数时，表示对所有医嘱序号进行整理
	就诊ID_IN		Varchar2:=Null--主页ID或挂号单,以字符类型传入
) IS
	v_主页ID		病人医嘱记录.主页ID%Type;
	v_挂号单		病人医嘱记录.挂号单%Type;
	v_婴儿			病人医嘱记录.婴儿%Type;

	Cursor c_Advice Is
		Select A.ID,A.婴儿
		From 病人医嘱记录 A,病人医嘱状态 B,病人医嘱记录 C
		Where A.ID=B.医嘱ID And B.操作类型=1
			And A.相关ID=C.ID(+) And A.病人ID=病人ID_IN 
			And (A.主页ID=v_主页ID Or A.挂号单=v_挂号单)
		Order by Nvl(A.婴儿,0),Nvl(C.序号,A.序号),Nvl(A.相关ID,A.ID),A.序号,B.操作时间;
	
	v_Count			Number;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
	If 医嘱ID_IN is Not Null Then
		Update 病人医嘱记录 Set 序号=序号_IN Where ID=医嘱ID_IN;
	Else
		v_主页ID:=Null;v_挂号单:=Null;
		Begin
			Select To_Number(就诊ID_IN) Into v_主页ID From Dual;
		Exception
			When Others Then Null;
		End;
		If v_主页ID Is Null Then
			v_挂号单:=就诊ID_IN;
		End IF;
		
		--重新整理序号
		v_Count:=1;
		For r_Advice In c_Advice Loop
			If Nvl(v_婴儿,0)<>Nvl(r_Advice.婴儿,0) Then
				v_Count:=1;
			End IF;
			Update 病人医嘱记录 Set 序号=v_Count Where ID=r_Advice.ID;
			v_婴儿:=r_Advice.婴儿;
			v_Count:=v_Count+1;
		End Loop;
	End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱记录_更新序号;
/

CREATE OR REPLACE PROCEDURE zl_药品差价重整_UPDATE (
    期间_IN IN 期间表.期间%TYPE
)
IS
        CURSOR C_期间表 
        IS 
              SELECT 开始日期,终止日期 
              FROM 期间表 
              WHERE 期间>=期间_IN and sysdate>=开始日期;

        CURSOR C_平均差价率 (
              V_开始日期 DATE,
              V_终止日期 DATE)
        IS 
              select D.药库,S.库房ID,S.药品ID,S.批次,decode(sign(S.金额),1,S.差价/S.金额,M.指导差价率/100) as 差价率
              from (select O.库房ID,O.药品ID,O.批次,
                           nvl(E.当前金额,0)-nvl(J.发生金额,0)-O.出库金额 as 金额,
                           nvl(E.当前差价,0)-nvl(J.发生差价,0)-O.出库差价 as 差价
                    from (select 库房ID,药品ID,nvl(批次,0) as 批次,
                                sum(入出系数*零售金额) as 出库金额,
                                sum(入出系数*差价) as 出库差价
                          from 药品收发记录 L
                          where 审核日期 between trunc(V_开始日期) and trunc(V_终止日期)+1-1/24/60/60
                                and (单据=6
                                     and exists 
                                         (select 1 
                                          from 部门性质说明 C 
                                          where C.部门ID=L.库房ID
                                                and C.工作性质 in('西药库','中药库','成药库'))
                                     and NOt exists 
                                         (select 1 
                                          from 部门性质说明 C 
                                          where C.部门ID=L.对方部门ID
                                                and C.工作性质 in('西药库','中药库','成药库','制剂室'))
                                    or 单据 between 7 and 11)
                          group by 库房ID,药品ID,nvl(批次,0)) O,
                         (select 库房ID,药品ID,nvl(批次,0) as 批次,
                                 sum(入出系数*零售金额) as 发生金额,
                                 sum(入出系数*差价) as 发生差价
                          from 药品收发记录
                          where 审核日期>=trunc(V_终止日期)+1
                          group by 库房ID,药品ID,nvl(批次,0)) J,
                         (select 库房ID,药品ID,nvl(批次,0) as 批次,
                                 sum(实际金额) as 当前金额,sum(实际差价) as 当前差价
                          from 药品库存
                          where 性质=1
                          group by 库房ID,药品ID,nvl(批次,0)) E
                    where O.库房ID=J.库房ID(+)
                          and O.药品ID=J.药品ID(+) 
                          and O.批次=J.批次(+)
                          and O.库房ID=E.库房ID(+)
                          and O.药品ID=E.药品ID(+) 
                          and O.批次=E.批次(+)) S,
                   药品规格 M,
                   (select 部门ID,min(decode(工作性质,'西药库',1,'中药库',1,'成药库',1,2)) as 药库
                    from 部门性质说明 
                    where 工作性质 in('西药库','中药库','成药库','西药房','中药房','成药房')
                    group by 部门ID) D
              where S.药品ID=M.药品ID and S.库房ID=D.部门ID 
              order by D.药库,S.库房ID,S.药品ID,S.批次;

        CURSOR C_药品出库记录 (
              V_开始日期 DATE,
              V_终止日期 DATE,
              V_药库 INTEGER,
              V_库房ID INTEGER,
              V_药品ID INTEGER,
              V_批次 INTEGER)
        IS 
            select ID,单据,NO,审核日期,入出类别ID,入出系数,
                   成本价,实际数量*付数 as 实际数量,零售金额,差价,产地,批号,效期,对方部门ID
            from 药品收发记录 L
            where 审核日期 between trunc(V_开始日期) and trunc(V_终止日期)+1-1/24/60/60
                  and 库房ID=V_库房ID
                  and 药品ID=V_药品ID
                  and NVL(批次,0)=NVL(V_批次,0)
                  and (V_药库=1 
                       and 单据=6
                       and NOt exists 
                           (select 1 
                            from 部门性质说明 C 
                            where C.部门ID=L.对方部门ID
                                  and C.工作性质 in('西药库','中药库','成药库','制剂室'))
                      or 单据 between 7 and 11);

       v_原差价 药品库存.实际差价%Type;
       v_现差价 药品库存.实际差价%Type;
       v_成本价 药品库存.上次采购价%Type;
       v_对方类别ID INTEGER;
       INTDIGIT NUMBER;
Begin
       --获取金额小数位数
       SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';

        FOR v_Period IN C_期间表 LOOP
            FOR v_AvgTax IN C_平均差价率 (v_Period.开始日期,v_Period.终止日期) LOOP
                FOR v_OutRec IN C_药品出库记录 (v_Period.开始日期,v_Period.终止日期,v_AvgTax.药库,v_AvgTax.库房ID,v_AvgTax.药品ID,v_AvgTax.批次) LOOP
                    v_原差价:=v_OutRec.差价;
                    v_现差价:=round(nvl(v_OutRec.零售金额,0)* v_AvgTax.差价率,INTDIGIT);
                    IF nvl(v_OutRec.实际数量,0)=0 THEN
                        v_成本价:=v_OutRec.成本价;
                    ELSE
                        v_成本价:=round((NVL(v_OutRec.零售金额,0)-v_现差价)/v_OutRec.实际数量,7);
                    END IF;

                    UPDATE 药品收发记录
                    SET    差价=v_现差价,
                           成本金额=NVL(v_OutRec.零售金额,0)-v_现差价,
                           成本价=v_成本价
                    WHERE ID = v_OutRec.ID;

                    UPDATE 药品库存
                    SET 实际差价 = nvl(实际差价,0)+(v_现差价-v_原差价)*v_OutRec.入出系数
                    WHERE 库房ID=v_AvgTax.库房ID
                          and 药品ID=v_AvgTax.药品ID
                          and Nvl(批次,0)=NVL(v_AvgTax.批次,0)
                          and 性质=1;
                    IF SQL%NOTFOUND THEN
                          Insert into 药品库存(库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次供应商ID,上次采购价,上次批号,上次产地,效期)
                          values (v_AvgTax.库房ID,v_AvgTax.药品ID,v_AvgTax.批次,1,0,0,0,(v_现差价-v_原差价)*v_OutRec.入出系数,null,v_成本价,v_OutRec.批号,v_OutRec.产地,v_OutRec.效期);
                    END IF;

                    UPDATE 药品收发汇总
                    SET 差价 = nvl(差价,0)+(v_现差价-v_原差价)*v_OutRec.入出系数
                    WHERE 日期 = TRUNC(v_OutRec.审核日期)
                          and 库房ID=v_AvgTax.库房ID
                          and 药品ID=v_AvgTax.药品ID
                          AND 类别ID = v_OutRec.入出类别ID
                          AND 单据 = v_OutRec.单据;

                    IF SQL%NOTFOUND THEN
                          Insert into 药品收发汇总(日期,库房ID,药品ID,类别ID,数量,金额,差价,单据)
                          values (TRUNC(v_OutRec.审核日期),v_AvgTax.库房ID,v_AvgTax.药品ID,v_OutRec.入出类别ID,0,0,(v_现差价-v_原差价)*v_OutRec.入出系数,v_OutRec.单据);
                    END IF;

                    IF v_OutRec.单据=6 THEN
                        UPDATE 药品收发记录
                        SET    差价=v_现差价,
                               成本金额=NVL(v_OutRec.零售金额,0)-v_现差价,
                               成本价=v_成本价
                        WHERE NO=v_OutRec.NO
                              AND 单据 = 6
                              and 药品ID+0=v_AvgTax.药品ID
                              and NVL(批次,0)=NVL(v_AvgTax.批次,0)
                              and 库房ID+0=v_OutRec.对方部门ID
                              and 对方部门ID+0=v_AvgTax.库房ID
                              and 入出系数=-1*v_OutRec.入出系数;
                        IF SQL%NOTFOUND THEN
                            NULL;
                        ELSE
                            SELECT 入出类别ID
                            INTO   v_对方类别ID
                            FROM   药品收发记录
                            WHERE NO=v_OutRec.NO
                                  AND 单据 = 6
                                  and 药品ID+0=v_AvgTax.药品ID
                                  and NVL(批次,0)=NVL(v_AvgTax.批次,0)
                                  and 库房ID+0=v_OutRec.对方部门ID
                                  and 对方部门ID+0=v_AvgTax.库房ID
                                  and 入出系数=-1*v_OutRec.入出系数
                                  and ROWNUM<2;

                            UPDATE 药品库存
                            SET 实际差价 = nvl(实际差价,0)+(v_现差价-v_原差价)*v_OutRec.入出系数*-1
                            WHERE 库房ID=v_OutRec.对方部门ID
                                  and 药品ID=v_AvgTax.药品ID
                                  and NVL(批次,0)=NVL(v_AvgTax.批次,0)
                                  and 性质=1;
                            IF SQL%NOTFOUND THEN
                                  Insert into 药品库存(库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价,上次供应商ID,上次采购价,上次批号,上次产地,效期)
                                  values (v_OutRec.对方部门ID,v_AvgTax.药品ID,v_AvgTax.批次,1,0,0,0,(v_现差价-v_原差价)*v_OutRec.入出系数*-1,null,v_成本价,v_OutRec.批号,v_OutRec.产地,v_OutRec.效期);
                            END IF;

                            UPDATE 药品收发汇总
                            SET 差价 = nvl(差价,0)+(v_现差价-v_原差价)*v_OutRec.入出系数*-1
                            WHERE 日期 = TRUNC(v_OutRec.审核日期)
                                  and 库房ID=v_OutRec.对方部门ID
                                  and 药品ID=v_AvgTax.药品ID
                                  AND 类别ID = v_对方类别ID
                                  AND 单据 = v_OutRec.单据;

                            IF SQL%NOTFOUND THEN
                                  Insert into 药品收发汇总(日期,库房ID,药品ID,类别ID,数量,金额,差价,单据)
                                  values (TRUNC(v_OutRec.审核日期),v_OutRec.对方部门ID,v_AvgTax.药品ID,v_对方类别ID,0,0,(v_现差价-v_原差价)*v_OutRec.入出系数*-1,v_OutRec.单据);
                            END IF;
                        END IF;
                    END IF;
                END LOOP;
                DELETE FROM 药品库存
                WHERE 库房ID=v_AvgTax.库房ID
                      and 药品ID=v_AvgTax.药品ID
                      AND nvl(可用数量,0) = 0
                      AND nvl(实际数量,0) = 0
                      AND nvl(实际金额,0) = 0
                      AND nvl(实际差价,0) = 0;
                COMMIT;
            END LOOP;
        END LOOP;
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_药品差价重整_UPDATE;
/

CREATE OR REPLACE PROCEDURE ZL_药品收发记录_部门发药 (
    PARTID_IN IN 药品收发记录.库房ID%TYPE,
    BILLID_IN IN 药品收发记录.ID%TYPE,
    PEOPLE_IN IN 药品收发记录.审核人%TYPE,
    DATE_IN IN 药品收发记录.审核日期%TYPE,
    批次_IN IN 药品收发记录.批次%TYPE:=NULL,
    发药方式_IN IN 药品收发记录.发药方式%TYPE:=3,
    领药人_IN In 药品收发记录.领用人%TYPE:=Null
)
IS
    --只读变量
    LNG入出类别ID NUMBER (18);
    INT入出系数 NUMBER;
    INT执行状态 NUMBER;
    INT单据 药品收发记录.单据%TYPE;
    STRNO 药品收发记录.NO%TYPE;
    LNG库房ID 药品收发记录.库房ID%TYPE;
    LNG药品ID 药品收发记录.药品ID%TYPE;
    LNG费用ID 药品收发记录.费用ID%TYPE;
    DBL库存金额 药品库存.实际金额%TYPE;
    DBL库存差价 药品库存.实际差价%TYPE;
    DBL差价率 药品规格.指导差价率%TYPE;
    INT未发数 未发药品记录.未发数%TYPE;
    --可写变量
    DBL实际数量 药品收发记录.实际数量%TYPE;
    DBL实际金额 药品收发记录.零售金额%TYPE;
    DBL成本金额 药品收发记录.成本金额%TYPE;
    DBL实际差价 药品收发记录.差价%TYPE;
    --2002-07-31朱玉宝
    --LNGLAST批次 发药前确定的批次(已减可用数量)
    INT记录状态 病人费用记录.执行状态%TYPE;
    STR药名 VARCHAR2(200);
    DBL可用数量 药品收发记录.填写数量%TYPE;
    LNGLAST批次 药品收发记录.批次%TYPE;
    LNGCUR批次 药品收发记录.批次%TYPE;
    STR批号 药品收发记录.批号%TYPE;
    STR效期 药品收发记录.效期%TYPE;
    BLN收费与发药分离 NUMBER(1);
    V_ERROR VARCHAR2(255);
    intDigit number(1);
    ERR_CUSTOM EXCEPTION ;

	--自动审核费用
	intAutoVerify NUMBER(1);
	str操作员编号 人员表.编号%TYPE;
	str操作员姓名 人员表.姓名%TYPE;
	int序号 病人费用记录.序号%TYPE;
	lng病人ID 病人费用记录.病人ID%TYPE;
BEGIN
	--取操作员编号与姓名
	SELECT 编号,姓名 INTO str操作员编号,str操作员姓名
	FROM 人员表 A,上机人员表 B
	WHERE A.ID=B.人员ID AND B.用户名=USER;

	--获取金额小数位数
    SELECT NVL(参数值,缺省值) INTO INTDIGIT FROM 系统参数表 WHERE 参数名='费用金额保留位数';
	--判断划价单发药后是否自动审核为记帐单
	SELECT NVL(参数值,缺省值) INTO intAutoVerify FROM 系统参数表 WHERE 参数名='执行后自动审核划价单';

    --获取该收发记录的单据、药品ID、库房ID,零售金额及实际数量、入出类别ID
    SELECT A.单据,A.NO,A.药品ID,A.库房ID,A.费用ID,NVL(A.零售金额,0),NVL(A.实际数量, 0)*NVL(A.付数,1),
        A.入出类别ID,A.入出系数,NVL(A.批次,0),'['||C.编码||']'||C.名称
    INTO INT单据, STRNO, LNG药品ID, LNG库房ID,LNG费用ID,DBL实际金额, DBL实际数量,
        LNG入出类别ID,INT入出系数,LNGLAST批次,STR药名
    FROM 药品收发记录 A,收费项目目录 C
    WHERE A.ID = BILLID_IN AND A.药品ID=C.ID;
    IF NVL(批次_IN,0)=0 THEN
        LNGCUR批次:=LNGLAST批次;
    ELSE
        LNGCUR批次:=NVL(批次_IN,0);
    END IF ;

    --检查是否已经填写库房
    BLN收费与发药分离:=0;
    IF LNG库房ID IS NULL THEN
        BLN收费与发药分离:=1;
    END IF ;
    LNG库房ID:=PARTID_IN;

    --取该批药品的批号
    BEGIN
        SELECT 上次批号,效期,NVL(可用数量,0)
        INTO STR批号,STR效期,DBL可用数量
        FROM 药品库存
        WHERE 库房ID=LNG库房ID AND 药品ID=LNG药品ID AND 性质=1 AND NVL(批次,0)=LNGCUR批次;
    EXCEPTION
        WHEN OTHERS THEN
            SELECT '','',0 INTO STR批号,STR效期,DBL可用数量 FROM DUAL ;
    END ;

    --可用数量不足则退出
    IF LNGCUR批次<>NVL(LNGLAST批次,0) THEN
        IF DBL可用数量<DBL实际数量 AND LNGCUR批次<>0 THEN
            V_ERROR:=STR药名||'的可用数量不足，操作中止！';
            RAISE ERR_CUSTOM;
        END IF ;
    END IF ;

    --计算该药品收发记录的成本价、成本金额、零售金额及差价(先计算零售金额，再计算差价，最后计算成本金额及成本价)
    --获取该药品的差价率(因为只有售价及库存数量充足,才允许更改批次,所以重新计算处使用当前确定的批次做为计算条件)
    BEGIN
        SELECT NVL(实际金额, 0) 实际金额, NVL(实际差价, 0) 实际差价 INTO DBL库存金额,DBL库存差价
        FROM 药品库存
        WHERE 库房ID+0 = LNG库房ID AND 药品ID = LNG药品ID
        AND 性质=1 AND NVL(批次,0)=LNGCUR批次 ;

        IF DBL库存金额<=0 OR DBL库存差价<0 THEN
            SELECT NVL (指导差价率, 15) / 100 指导差价率 INTO DBL差价率
            FROM 药品规格 WHERE 药品ID = LNG药品ID;
        ELSE
            DBL差价率 := DBL库存差价/DBL库存金额 ;
        END IF ;
    EXCEPTION
        WHEN OTHERS THEN
            SELECT NVL (指导差价率, 15) / 100 指导差价率 INTO DBL差价率
            FROM 药品规格 WHERE 药品ID = LNG药品ID;
    END ;
    --差价
    DBL实际差价 := round(DBL实际金额 * DBL差价率,intDigit);
    --成本金额
    DBL成本金额 := round(DBL实际金额 - DBL实际差价,intDigit);

    --更新药品收发记录的零售金额、成本金额及差价
    UPDATE 药品收发记录
    SET 库房ID=lng库房ID,
        成本价 = round(DBL成本金额 / DECODE(DBL实际数量,NULL,1,0,1,DBL实际数量),5),
        成本金额 = DBL成本金额,
        差价 = DBL实际差价,
        批次=LNGCUR批次,
        批号=STR批号,
        效期=STR效期,
        审核人=PEOPLE_IN,
        审核日期=DATE_IN,
        发药方式=发药方式_IN,
        领用人=领药人_IN
    WHERE ID = BILLID_IN;
	--并发操作检查
	IF SQL%RowCount=0 Then
		v_Error:='要发药的药品记录"'||STR药名||'"不存在，操作中止！';
		Raise Err_Custom;
	End IF;

    --更改所有已发药处方的配药人为发药人
    UPDATE 药品收发记录
    SET 配药人 = PEOPLE_IN
    WHERE NO = STRNO AND 单据 = INT单据 AND (库房ID+0=LNG库房ID OR 库房ID IS NULL) AND 审核人 IS NOT NULL AND MOD (记录状态, 3) = 1;

    --更新原批次库存的可用数量
    --更新发药批次库存的可用及实际数量
    IF LNGLAST批次<>LNGCUR批次 THEN
        UPDATE 药品库存
        SET 可用数量=NVL(可用数量,0)+DBL实际数量
        WHERE 库房ID+0 = LNG库房ID AND 药品ID = LNG药品ID AND 性质 = 1 AND NVL(批次,0)=LNGLAST批次;

        UPDATE 药品库存
        SET 可用数量=NVL(可用数量,0)-DBL实际数量
        WHERE 库房ID+0 = LNG库房ID AND 药品ID = LNG药品ID AND 性质 = 1 AND NVL(批次,0)=LNGCUR批次;
    END IF ;

    IF BLN收费与发药分离=1 THEN
        UPDATE 药品库存
        SET 可用数量 = NVL (可用数量, 0) - DBL实际数量,
            实际数量 = NVL (实际数量, 0) - DBL实际数量,
            实际金额 = NVL (实际金额, 0) - DBL实际金额,
            实际差价 = NVL (实际差价, 0) - DBL实际差价
        WHERE 库房ID+0 = LNG库房ID AND 药品ID = LNG药品ID AND 性质 = 1 AND NVL(批次,0)=LNGCUR批次;
    ELSE
        UPDATE 药品库存
        SET 实际数量 = NVL (实际数量, 0) - DBL实际数量,
            实际金额 = NVL (实际金额, 0) - DBL实际金额,
            实际差价 = NVL (实际差价, 0) - DBL实际差价
        WHERE 库房ID+0 = LNG库房ID AND 药品ID = LNG药品ID AND 性质 = 1 AND NVL(批次,0)=LNGCUR批次;
    END IF ;

    IF SQL%ROWCOUNT = 0 THEN
        IF BLN收费与发药分离=1 THEN
            INSERT INTO 药品库存
            (库房ID,药品ID,批次,性质,可用数量,实际数量,实际金额,实际差价)
            VALUES
            (LNG库房ID,LNG药品ID,LNGCUR批次,1,0 - DBL实际数量,0 - DBL实际数量,0 - DBL实际金额,0 - DBL实际差价);
        ELSE
            INSERT INTO 药品库存
            (库房ID,药品ID,批次,性质,实际数量,实际金额,实际差价)
            VALUES
            (LNG库房ID,LNG药品ID,LNGCUR批次,1,0 - DBL实际数量,0 - DBL实际金额,0 - DBL实际差价);
        END IF ;
    END IF;

    DELETE 药品库存
    WHERE 库房ID+0 = LNG库房ID AND 药品ID = LNG药品ID AND 性质=1
    AND NVL(可用数量,0) = 0 AND NVL(实际数量,0) = 0 AND NVL(实际金额,0) = 0 AND NVL(实际差价,0) = 0;

    --更新药品收发汇总
    UPDATE 药品收发汇总
    SET 数量 = NVL (数量, 0) + DBL实际数量 * INT入出系数,
        金额 = NVL (金额, 0) + DBL实际金额 * INT入出系数,
        差价 = NVL (差价, 0) + DBL实际差价 * INT入出系数
    WHERE 库房ID+0 = LNG库房ID AND 药品ID+0 = LNG药品ID AND 类别ID+0 = LNG入出类别ID
    AND 日期 = TRUNC (DATE_IN) AND 单据 = INT单据;

    IF SQL%ROWCOUNT = 0 THEN
        INSERT INTO 药品收发汇总
        (日期,库房ID,药品ID,单据,类别ID,数量,金额,差价)
        VALUES
        (TRUNC (DATE_IN),LNG库房ID,LNG药品ID,INT单据,LNG入出类别ID,
        DBL实际数量 * INT入出系数,DBL实际金额 * INT入出系数,DBL实际差价 * INT入出系数);
    END IF;

    --更新病人费用记录的执行状态(已执行)
    SELECT DECODE(SUM(NVL(付数,1)*实际数量),NULL,1,0,1,2) INTO INT执行状态
    FROM 药品收发记录
    WHERE 单据=INT单据 AND NO=STRNO AND 费用ID=LNG费用ID AND 审核人 IS NULL AND 记录状态<>1 AND MOD(记录状态,3)<>0;
    UPDATE 病人费用记录
    SET 执行状态 = INT执行状态
    WHERE ID = LNG费用ID;

    --更新未发药品记录(如果未发数为零则删除)
    SELECT COUNT(*) INTO INT未发数
    FROM 药品收发记录
    WHERE 单据=INT单据 AND (库房ID+0=LNG库房ID OR 库房ID IS NULL) AND NO=STRNO AND 审核人 IS NULL AND NVL(LTRIM(RTRIM(摘要)),'小宝')<>'拒发';

    IF INT未发数 = 0 THEN
        DELETE 未发药品记录 WHERE NO = STRNO AND 单据 = INT单据 AND (库房ID+0=LNG库房ID OR 库房ID IS NULL);
    END IF;

	--费用审核（重复审核也没有关系）
	IF intAutoVerify=1 THEN
		SELECT 序号,病人ID,NO INTO int序号,lng病人ID,strNO
		FROM 病人费用记录
		WHERE id = (SELECT 费用ID FROM 药品收发记录 WHERE ID=BILLID_IN);

		zl_住院记帐记录_Verify(strNO,str操作员编号,str操作员姓名,int序号,lng病人ID,DATE_IN);
	END IF ;
EXCEPTION
    WHEN ERR_CUSTOM THEN RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]'||V_ERROR||'[ZLSOFT]');
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_药品收发记录_部门发药;
/

-------------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.12.0' Where 编号=100;
--部件版本号
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9BaseItem')	And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9CISBase')		And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9Patient')		And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9InPatient')   And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9RegEvent')	And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9OutExse')		And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9InExse')		And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9CustAcc')		And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9CashBill')	And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9DrugStore')   And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9MediStore')   And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9Stuff')		And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9Due')			And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9Analysis')	And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9Ops')			And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9Medical')		And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9CISWork')		And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9PacsWork')	And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9LisWork')		And 系统=100;
Update zlComponent Set 主版本=10,次版本=12,附版本=0 Where Upper(部件)=Upper('zl9ImgCapture')	And 系统=100;
--医保部件
Update zlComponent Set 主版本=9,次版本=23,附版本=50 Where Upper(部件)=Upper('zl9Insure')		And 系统=100;

Commit;