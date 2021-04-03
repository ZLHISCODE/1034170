--分类目录
--1.公共基础,2.医保基础,3.病人病案基础,4.费用基础,5.药品卫材基础
--6.临床基础,7.临床路径基础,8.病历基础,9.护理基础,10.检验基础
--11.检查基础,12.医保业务,13.病人病案业务,14.费用业务,15.药品卫材业务
--16.临床医嘱,17.临床路径,18.病历业务,19.护理业务,20.检验业务,21.检查业务

----------------------------------------------------------------------------
--[[1.公共基础]]
----------------------------------------------------------------------------
Create Index 人员表_IX_签名 On 人员表(签名) Tablespace zl9Indexhis;
Create Index 人员性质说明_IX_人员性质 On 人员性质说明(人员性质) Tablespace zl9Indexhis;
Create Index 人员证书记录_IX_人员ID On 人员证书记录(人员ID) Tablespace zl9Indexhis;
Create Index 部门性质说明_IX_工作性质 On 部门性质说明(工作性质) Tablespace zl9Indexhis;
Create Index 部门人员_IX_人员ID On 部门人员(人员ID) Tablespace zl9Indexhis;
Create Index 临床部门_IX_部门ID On 临床部门(部门ID) Tablespace zl9Indexhis;
Create Index 病区科室对应_IX_科室ID On 病区科室对应(科室ID) Tablespace zl9Indexhis;

Create Index 排队叫号队列_IX_科室ID On 排队叫号队列(科室id) Tablespace zl9Indexhis;
Create Index 排队叫号队列_IX_病人ID On 排队叫号队列(病人ID) Tablespace zl9Indexhis;
create index 排队叫号队列_IX_业务ID on 排队叫号队列(业务ID,业务类型) tablespace zl9indexhis;
Create index 排队语音呼叫_IX_队列ID on 排队语音呼叫(队列ID,站点) Tablespace zl9indexhis;

----------------------------------------------------------------------------
--[[2.医保基础]]
----------------------------------------------------------------------------
Create Index 保险结算记录_IX_病人ID On 保险结算记录(病人ID) Tablespace zl9Indexhis;
Create Index 保险结算记录_IX_结算时间 On 保险结算记录(结算时间) Tablespace zl9Indexhis;
Create Index 保险支付项目_IX_大类ID On 保险支付项目(大类ID,险类) Tablespace zl9Indexhis;
Create Index 保险支付项目_IX_项目编码 On 保险支付项目(项目编码,险类) Tablespace zl9Indexhis;
Create Index 审批项目模板_IX_项目ID On 审批项目模板(项目ID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[3.病人病案基础]]
----------------------------------------------------------------------------
Create Index 疾病编码分类_IX_上级ID On 疾病编码分类(上级ID) Tablespace zl9Indexhis;
Create Index 疾病编码目录_IX_分类ID On 疾病编码目录(分类ID) Tablespace zl9Indexhis;
Create Index 疾病编码目录_IX_名称 On 疾病编码目录(名称) Tablespace zl9Indexhis;
Create Index 疾病编码目录_IX_简码 On 疾病编码目录(简码) Tablespace zl9Indexhis;
Create Index 疾病编码目录_IX_五笔码 On 疾病编码目录(五笔码) Tablespace zl9Indexhis;
Create Index 疾病编码科室_IX_科室ID On 疾病编码科室(科室ID) Tablespace zl9Indexhis;
Create Index 疾病编码科室_IX_人员ID On 疾病编码科室(人员ID) Tablespace zl9Indexhis;
Create Index 疾病诊断科室_IX_科室ID On 疾病诊断科室(科室ID) Tablespace zl9Indexhis;
Create Index 疾病诊断科室_IX_人员ID On 疾病诊断科室(人员ID) Tablespace zl9Indexhis;
Create Index 疾病诊断分类_IX_上级ID On 疾病诊断分类(上级ID) Tablespace zl9Indexcis;
Create Index 疾病诊断别名_IX_诊断ID On 疾病诊断别名(诊断id) Tablespace zl9Indexcis;
Create Index 疾病诊断别名_IX_名称 On 疾病诊断别名(名称) Tablespace zl9Indexcis;
Create Index 疾病诊断别名_IX_简码 On 疾病诊断别名(简码) Tablespace zl9Indexcis;
Create Index 疾病诊疗措施_IX_诊疗项目ID On 疾病诊疗措施(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 疾病诊断规则_IX_项目ID On 疾病诊断规则(项目ID) Tablespace zl9Indexcis;
Create Index 疾病诊断对照_IX_诊断ID On 疾病诊断对照(诊断ID) Tablespace zl9Indexcis;
Create Index 疾病诊断对照_IX_手术ID On 疾病诊断对照(手术ID) Tablespace zl9Indexcis;

Create Index 咨询表格内容_IX_表号 On 咨询表格内容(表号) Tablespace zl9Indexhis;
Create Index 咨询广告序列_IX_图片序号 On 咨询广告序列(图片序号) Tablespace zl9Indexhis;
Create Index 咨询页面目录_IX_宣传标语 On 咨询页面目录(宣传标语) Tablespace zl9Indexhis;
Create Index 咨询页面目录_IX_页面背景 On 咨询页面目录(页面背景) Tablespace zl9Indexhis;
Create Index 咨询页面目录_IX_上级序号 On 咨询页面目录(上级序号) Tablespace zl9Indexhis;
Create Index 咨询页面排列_IX_页面 On 咨询页面排列(页面) Tablespace zl9Indexhis;
Create Index 咨询页面排列_IX_父序号 On 咨询页面排列(父序号) Tablespace zl9Indexhis;
Create Index 咨询页面排列_IX_页面图标 On 咨询页面排列(页面图标) Tablespace zl9Indexhis;
Create Index 咨询段落目录_IX_页面序号 On 咨询段落目录(页面序号) Tablespace zl9Indexhis;
Create Index 咨询段落目录_IX_标题图标 On 咨询段落目录(标题图标) Tablespace zl9Indexhis;
Create Index 咨询段落目录_IX_插表序号 On 咨询段落目录(插表序号) Tablespace zl9Indexhis;
Create Index 咨询段落目录_IX_插图序号 On 咨询段落目录(插图序号) Tablespace zl9Indexhis;
Create Index 咨询段落链接_IX_链接 On 咨询段落链接(页面序号,段落序号) Tablespace zl9Indexhis;
Create Index 咨询段落链接_IX_链接页面 On 咨询段落链接(链接页面) Tablespace zl9Indexhis;
Create Index 咨询专家清单_IX_人员id On 咨询专家清单(人员id) Tablespace zl9Indexhis;
Create Index 咨询专家清单_IX_科室id On 咨询专家清单(科室id) Tablespace zl9Indexhis;


----------------------------------------------------------------------------
--[[4.费用基础]]
----------------------------------------------------------------------------
Create Index 费别明细_IX_收费细目id On 费别明细(费别, 收费细目id) Tablespace zl9Indexhis;
Create Index 收费分类目录_IX_上级ID On 收费分类目录(上级ID) Tablespace zl9Indexhis;
Create Index 收费项目目录_IX_分类ID On 收费项目目录(分类ID) Tablespace zl9Indexhis;
Create Index 收费项目别名_IX_名称 On 收费项目别名(名称) Tablespace zl9Indexhis;
Create Index 收费项目别名_IX_简码 On 收费项目别名(简码) Tablespace zl9Indexhis;
Create Index 收费执行科室_IX_开单科室ID On 收费执行科室(开单科室ID) Tablespace zl9Indexhis;
Create Index 收费执行科室_IX_执行科室ID On 收费执行科室(执行科室ID) Tablespace zl9Indexhis;
Create Index 收费价目_IX_收费细目id On 收费价目(收费细目id) Tablespace zl9Indexhis;
Create Index 成套项目分类_IX_简码 On 成套项目分类(简码) Tablespace zl9Indexhis;
Create Index 成套收费项目_IX_拼音 On 成套收费项目(拼音) Tablespace zl9Indexhis;
Create Index 成套收费项目_IX_五笔 On 成套收费项目(五笔) Tablespace zl9Indexhis;
Create Index 成套收费项目_IX_分类ID On 成套收费项目(分类ID) Tablespace zl9Indexhis;

Create Index 挂号安排_IX_执行计划ID On 挂号安排(执行计划ID) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_安排时间 On 挂号安排计划(安排时间) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_审核时间 On 挂号安排计划(审核时间) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_生效时间 On 挂号安排计划(生效时间) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_失效时间 On 挂号安排计划(失效时间) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_实际生效 On 挂号安排计划(实际生效) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_安排ID On 挂号安排计划(安排ID) Tablespace zl9Indexhis;
Create Index 挂号安排停用状态_IX_开始时间 On 挂号安排停用状态(开始停止时间) Tablespace zl9Indexhis;
Create Index 挂号安排停用状态_IX_结束时间 On 挂号安排停用状态(结束停止时间) Tablespace zl9Indexhis;
Create Index 常用退费原因_IX_简码 On 常用退费原因(简码) Tablespace zl9Indexhis;

Create Index 常用发卡原因_IX_简码 On 常用发卡原因(简码) Tablespace zl9Indexhis;
Create Index 消费卡目录_IX_发卡序号 On 消费卡目录(发卡序号) Tablespace zl9Indexhis;
Create Index 消费卡目录_IX_有效期 On 消费卡目录(有效期) Tablespace zl9Indexhis;
Create Index 消费卡目录_IX_发卡时间 On 消费卡目录(发卡时间) Tablespace zl9Indexhis;
Create Index 消费卡目录_IX_回收时间 On 消费卡目录(回收时间) Tablespace zl9Indexhis;
Create Index 消费卡目录_IX_当前状态 On 消费卡目录(当前状态) Tablespace zl9Indexhis;
Create Index 消费卡目录_IX_停用日期 On 消费卡目录(停用日期) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[5.药品卫材基础]]
----------------------------------------------------------------------------
Create Index 药品规格_IX_药名ID On 药品规格(药名ID) Tablespace zl9Indexhis;
Create Index 药品规格_IX_标识码 On 药品规格(标识码) Tablespace zl9Indexhis;
Create Index 材料特性_IX_诊疗ID On 材料特性(诊疗ID) Tablespace zl9Indexhis;
Create Index 材料领用信息_IX_主页ID On 材料领用信息(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 材料领用用途_IX_简码 On 材料领用用途(简码) Tablespace zl9Indexhis;
Create Index 供应商_IX_上级ID On 供应商(上级ID) Tablespace zl9Indexhis;
Create Index 供应商_IX_简码 On 供应商(简码) Tablespace zl9Indexhis;
Create Index 收费价目_IX_调价汇总号 On 收费价目(调价汇总号) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[6.临床基础]]
----------------------------------------------------------------------------
Create Index 输血检验对照_IX_检验项目id On 输血检验对照(检验项目id) Tablespace zl9Indexhis;
Create Index 抗菌药物抽样记录_IX_抽样时间 On 抗菌药物抽样记录(抽样时间) Tablespace zl9Indexcis;
Create Index 抗菌药物抽样明细_IX_病人ID On 抗菌药物抽样明细(病人ID,主页ID) Tablespace zl9Indexcis;
Create Index 抗菌药物抽样明细_IX_临床症状 On 抗菌药物抽样明细(临床症状) Tablespace zl9Indexcis;
Create Index 抗菌药物抽样明细_IX_感染诊断 On 抗菌药物抽样明细(感染诊断) Tablespace zl9Indexcis;
Create Index 抗菌药物抽样手术_IX_手术ID On 抗菌药物抽样手术(手术ID) Tablespace zl9Indexcis;
Create Index 诊疗分类目录_IX_上级ID On 诊疗分类目录(上级ID) Tablespace zl9Indexhis;
Create Index 诊疗项目目录_IX_分类ID On 诊疗项目目录(分类ID) Tablespace zl9Indexhis;
Create Index 诊疗项目别名_IX_名称 On 诊疗项目别名(名称) Tablespace zl9Indexhis;
Create Index 诊疗项目别名_IX_简码 On 诊疗项目别名(简码) Tablespace zl9Indexhis;
Create Index 诊疗执行科室_IX_开单科室ID On 诊疗执行科室(开单科室ID) Tablespace zl9Indexcis;
Create Index 诊疗执行科室_IX_执行科室ID On 诊疗执行科室(执行科室ID) Tablespace zl9Indexcis;
Create Index 诊疗项目组合_IX_诊疗项目ID On 诊疗项目组合(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 诊疗项目组合_IX_配方ID On 诊疗项目组合(配方ID) Tablespace zl9Indexcis;
Create Index 诊疗收费关系_IX_收费项目ID On 诊疗收费关系(收费项目id) Tablespace zl9Indexcis;

Create Index 人员抗菌药物权限_Ix_人员id On 人员抗菌药物权限(人员id) Tablespace Zl9Indexhis;
Create Index 人员手术权限_IX_诊疗项目ID On 人员手术权限(诊疗项目ID) Tablespace zl9Indexhis;


----------------------------------------------------------------------------
--[[7.临床路径基础]]
----------------------------------------------------------------------------
Create Index 临床路径项目_IX_版本号 On 临床路径项目(路径ID,版本号) Tablespace zl9Indexcis;
Create Index 临床路径项目_IX_阶段ID On 临床路径项目(阶段ID) Tablespace zl9Indexcis;
Create Index 临床路径项目_IX_图标ID On 临床路径项目(图标ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_相关ID On 路径医嘱内容(相关ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_诊疗项目ID On 路径医嘱内容(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_收费细目ID On 路径医嘱内容(收费细目ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_执行科室ID On 路径医嘱内容(执行科室ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_配方ID On 路径医嘱内容(配方ID) Tablespace zl9Indexcis;
Create Index 临床路径分支_IX_前一阶段ID On 临床路径分支(前一阶段ID) Tablespace zl9Indexhis;
Create Index 临床路径阶段_IX_分支ID On 临床路径阶段(分支ID) Tablespace zl9Indexhis;
Create Index 临床路径阶段_IX_父ID On 临床路径阶段(父ID) Tablespace zl9Indexcis;
Create Index 临床路径分类_IX_分支ID On 临床路径分类(分支ID) Tablespace zl9Indexhis;
Create Index 临床路径项目_IX_分支ID On 临床路径项目(分支ID) Tablespace zl9Indexhis;
Create Index 临床路径评估_IX_分支ID On 临床路径评估(分支ID) Tablespace zl9Indexhis;
Create Index 临床路径评估_IX_阶段ID On 临床路径评估(阶段ID) Tablespace zl9Indexcis;
Create Index 路径评估条件_IX_评估ID On 路径评估条件(评估ID) Tablespace zl9Indexcis;
Create Index 路径评估条件_IX_项目ID On 路径评估条件(项目ID) Tablespace zl9Indexcis;
Create Index 临床路径医嘱_IX_医嘱内容ID On 临床路径医嘱(医嘱内容ID) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[8.病历基础]]
----------------------------------------------------------------------------
Create Index 诊治所见项目_IX_分类ID On 诊治所见项目(分类ID) Tablespace zl9Indexcis;
Create Index 病历提纲词句_IX_词句分类ID On 病历提纲词句(词句分类ID) Tablespace zl9Indexcis;
Create Index 病历替代关系_IX_替代ID On 病历替代关系(替代ID) Tablespace zl9Indexcis;
Create Index 病历应用科室_IX_科室ID On 病历应用科室(科室ID) Tablespace zl9Indexcis;
Create Index 疾病报告前提_IX_疾病ID On 疾病报告前提(疾病ID) Tablespace zl9Indexcis;
Create Index 疾病报告前提_IX_诊断ID On 疾病报告前提(诊断ID) Tablespace zl9Indexcis;
Create Index 病历单据应用_IX_病历文件ID On 病历单据应用(病历文件ID) Tablespace zl9Indexcis;
Create Index 病历附项模板_IX_病历文件Id On 病历附项模板(病历文件Id,单据附项) Tablespace zl9Indexhis;
Create Index 病历文件结构_IX_父ID On 病历文件结构(父ID) Tablespace zl9Indexcis;
Create Index 病历文件结构_IX_预制提纲ID On 病历文件结构(预制提纲ID) Tablespace zl9Indexcis;
Create Index 病历文件结构_IX_诊治要素ID On 病历文件结构(诊治要素ID) Tablespace zl9Indexcis;
Create Index 病历词句示范_IX_科室id On 病历词句示范(科室id) Tablespace zl9Indexcis;
Create Index 病历词句示范_IX_人员id On 病历词句示范(人员id) Tablespace zl9Indexcis;
Create Index 病历词句示范_IX_编号 On 病历词句示范(编号) Tablespace zl9Indexcis;
Create Index 病历词句示范_IX_名称 On 病历词句示范(名称) Tablespace zl9Indexcis;
Create Index 病历词句组成_IX_内容文本 On 病历词句组成(内容文本) Tablespace zl9Indexcis;
Create Index 病历范文目录_IX_科室id On 病历范文目录(科室id) Tablespace zl9Indexcis;
Create Index 病历范文目录_IX_人员id On 病历范文目录(人员id) Tablespace zl9Indexcis;
Create Index 病历范文内容_IX_父ID On 病历范文内容(父ID) Tablespace zl9Indexcis;
Create Index 病历范文内容_IX_预制提纲ID On 病历范文内容(预制提纲ID) Tablespace zl9Indexcis;
Create Index 病历范文内容_IX_诊治要素ID On 病历范文内容(诊治要素ID) Tablespace zl9Indexcis;

Create Index 病案审查分类_IX_上级id On 病案审查分类(上级id) Tablespace zl9Indexcis;
Create Index 病案审查分类_IX_方案id On 病案审查分类(方案id) Tablespace zl9Indexcis;
Create Index 病案审查目录_IX_分类id On 病案审查目录(分类id) Tablespace zl9Indexcis;
----------------------------------------------------------------------------
--[[9.护理基础]]
----------------------------------------------------------------------------
Create Index 体温重叠标记_IX_上级序号 On 体温重叠标记(上级序号) Tablespace zl9Indexcis;
Create Index 护理适用科室_IX_科室ID On 护理适用科室(科室ID) Tablespace zl9Indexcis;
----------------------------------------------------------------------------
--[[10.检验基础]]
----------------------------------------------------------------------------
Create Index 检验细菌_IX_简码 On 检验细菌(简码) Tablespace zl9Indexcis;
Create Index 检验试剂关系_IX_材料id On 检验试剂关系(材料id) Tablespace zl9Indexcis;
Create Index 检验备注文字_IX_分类 On 检验备注文字(分类) Tablespace zl9Indexcis;
Create Index 检验评语文字_IX_分类 On 检验评语文字(分类) Tablespace zl9Indexcis;
Create Index 检验报告项目_IX_细菌ID On 检验报告项目(细菌id) Tablespace zl9Indexcis;
Create Index 检验报告项目_IX_报告项目ID On 检验报告项目(报告项目ID) Tablespace zl9Indexcis;

Create Index 检验模板目录_IX_诊疗项目ID On 检验模板目录(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 检验模板内容_IX_模板ID On 检验模板内容(模板ID) Tablespace zl9Indexcis;
Create Index 检验模板内容_IX_项目ID On 检验模板内容(项目ID) Tablespace zl9Indexcis;
Create Index 检验模板内容_IX_细菌ID On 检验模板内容(细菌ID) Tablespace zl9Indexcis;
Create Index 检验模板药敏_IX_抗生素ID On 检验模板药敏(抗生素ID) Tablespace zl9Indexcis;
Create Index 检验合并规则_IX_主项目ID On 检验合并规则(主项目ID) Tablespace zl9Indexcis;
Create Index 检验合并规则_IX_合并项目ID On 检验合并规则(合并项目ID) Tablespace zl9Indexcis;

Create Index 检验仪器_IX_使用小组ID On 检验仪器(使用小组ID) Tablespace zl9Indexcis;
Create Index 检验仪器抗生素_IX_抗生素ID On 检验仪器项目(抗生素id) Tablespace zl9Indexcis;
Create Index 检验仪器抗生素_IX_项目ID On 检验仪器项目(项目id) Tablespace zl9Indexcis;
Create Index 检验仪器状态_IX_项目ID On 检验仪器状态(项目ID) Tablespace zl9Indexcis;
Create Index 检验仪器规则_IX_上级ID On 检验仪器规则(上级ID) Tablespace zl9Indexcis;
Create Index 检验仪器规则_IX_仪器ID On 检验仪器规则(仪器ID) Tablespace zl9Indexcis;
Create Index 检验仪器规则_IX_规则ID On 检验仪器规则(规则ID) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[11.检查基础]]
----------------------------------------------------------------------------
Create Index 病理号码记录_IX_年 On 病理号码记录(年) Pctfree 5 Tablespace zl9Indexcis;
Create Index 影像查询方案_IX_所属科室 On 影像查询方案(所属科室) Tablespace zl9Indexhis;
Create Index 影像查询配置_IX_方案ID On 影像查询配置(方案ID) Tablespace zl9Indexhis;
Create Index 快捷功能信息_IX_模块号 On 快捷功能信息(模块号,项目) Tablespace zl9Indexhis;
create index 医技执行房间_IX_分组ID on 医技执行房间(分组ID) Tablespace zl9Indexhis;
create index 影像分组关联_IX_分组ID on 影像分组关联(分组ID) Tablespace zl9Indexhis;
create index 影像执行分组_IX_科室ID on 影像执行分组(科室ID) Tablespace zl9Indexhis;
----------------------------------------------------------------------------
--[[12.医保业务]]
----------------------------------------------------------------------------
Create Index 医保病人档案_IX_就诊时间 On 医保病人档案(就诊时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 医保病人关联表_IX_病人ID On 医保病人关联表(病人ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人审批项目_IX_项目ID On 病人审批项目(项目ID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[13.病人病案业务]]
----------------------------------------------------------------------------
Create Index 床位增减记录_IX_病区ID On 床位增减记录(病区ID) Tablespace zl9Indexhis;
Create Index 床位状况记录_IX_科室ID On 床位状况记录(科室ID) Tablespace zl9Indexhis;
Create Index 床位状况记录_IX_病人ID On 床位状况记录(病人ID) Tablespace zl9Indexhis;

Create Index 病人信息_IX_姓名 On 病人信息(姓名) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_登记时间 On 病人信息(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_身份证号 On 病人信息(身份证号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_IC卡号 On 病人信息(IC卡号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_医保号 On 病人信息(医保号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_合同单位id On 病人信息(合同单位id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_在院 On 病人信息(在院) Pctfree 5 Tablespace zl9Indexhis;
Create Index 在院病人_IX_病人ID On 在院病人(病人ID) Tablespace zl9Indexhis;
Create Index 病人合并记录_IX_病人ID On 病人合并记录(病人id) Tablespace zl9Indexhis;
Create Index 特殊病人_IX_病人ID On 特殊病人(病人ID) Tablespace zl9Indexhis;
Create Index 特殊病人_IX_加入时间 On 特殊病人(加入时间) Tablespace zl9Indexhis;
Create Index 病人担保记录_IX_主页ID On 病人担保记录(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 病人变动记录_IX_病人ID On 病人变动记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人变动记录_IX_科室id On 病人变动记录(科室id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人变动记录_IX_医疗小组ID On 病人变动记录(医疗小组ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人变动记录_IX_病区ID On 病人变动记录(病区ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人变动记录_IX_开始时间 On 病人变动记录(开始时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人变动记录_IX_终止时间 On 病人变动记录(终止时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_入院日期 On 病案主页(入院日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_出院日期 On 病案主页(出院日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_出院科室ID On 病案主页(出院科室ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_当前病区ID On 病案主页(当前病区ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_医疗小组ID On 病案主页(医疗小组ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_住院号 On 病案主页(住院号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_病案号 On 病案主页(病案号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_婴儿科室ID On 病案主页(婴儿科室ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_婴儿病区ID On 病案主页(婴儿病区ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_待转出 On 病案主页(待转出) Tablespace zl9Indexhis;

Create Index 病人地址信息_IX_省 On 病人地址信息(省) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人地址信息_IX_市 On 病人地址信息(市) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人地址信息_IX_县 On 病人地址信息(县) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院病案记录_IX_病人ID On 住院病案记录(病案号) PCTFREE 5 Tablespace zl9Indexhis;

Create Index 病人过敏记录_IX_病人ID On 病人过敏记录(病人ID) Tablespace zl9Indexcis;
Create Index 病人过敏记录_IX_待转出 On 病人过敏记录(待转出) Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_病人ID On 病人诊断记录(病人ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_医嘱id On 病人诊断记录(医嘱id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人诊断记录_IX_病历ID On 病人诊断记录(病历ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_病例ID On 病人诊断记录(病例ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_待转出 On 病人诊断记录(待转出) Tablespace zl9Indexcis;
Create Index 病人诊断医嘱_IX_医嘱ID On 病人诊断医嘱(医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断医嘱_IX_待转出 On 病人诊断医嘱(待转出) Tablespace zl9Indexcis;
Create Index 病人手麻记录_IX_主页ID On 病人手麻记录(病人ID,主页ID ) Tablespace zl9Indexcis;
Create Index 病人手麻记录_IX_待转出 On 病人手麻记录(待转出) Tablespace zl9Indexcis;
Create Index 病人抗生素记录_IX_药名id On 病人抗生素记录(药名id) Tablespace zl9Indexcis;

Create Index 病案化疗记录_IX_开始日期 On 病案化疗记录(开始日期) PCTFREE 5 Tablespace zl9Indexcis;
Create Index 病案化疗记录_IX_结束日期 On 病案化疗记录(结束日期) PCTFREE 5 Tablespace zl9Indexcis;
Create Index 病案放疗记录_IX_开始日期 On 病案放疗记录(开始日期) PCTFREE 5 Tablespace zl9Indexcis;
Create Index 病案放疗记录_IX_结束日期 On 病案放疗记录(结束日期) PCTFREE 5 Tablespace zl9Indexcis;
Create Index 病案精神治疗_IX_登记时间 On 病案精神治疗(药物名称) PCTFREE 5 Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[14.费用业务]]
----------------------------------------------------------------------------
Create Index 凭条打印记录_IX_待转出 On 凭条打印记录(待转出) Tablespace zl9Indexhis;
Create Index 三方结算交易_IX_待转出 On 三方结算交易(待转出) Tablespace zl9Indexhis;
Create Index 消费卡充值记录_IX_充值时间 On 消费卡充值记录(充值时间) Tablespace zl9Indexhis;
Create Index 病人卡结算记录_IX_交易时间 On 病人卡结算记录(交易时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人卡结算记录_IX_待转出 On 病人卡结算记录(待转出) Tablespace zl9Indexhis;
Create Index 病人卡结算对照_IX_卡结算ID On 病人卡结算对照(卡结算ID,预交ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人卡结算对照_IX_待转出 On 病人卡结算对照(待转出) Tablespace zl9Indexhis;
Create Index 病人医疗卡变动_IX_卡类别ID On 病人医疗卡变动(卡类别ID) Tablespace zl9Indexhis;
Create Index 病人医疗卡变动_IX_变动ID On 病人医疗卡变动(变动ID) Tablespace zl9Indexhis;
Create Index 病人医疗卡信息_IX_挂失时间 On 病人医疗卡信息(挂失时间) Tablespace zl9Indexhis;
Create Index 病人医疗卡信息_IX_发卡日期 On 病人医疗卡信息(发卡日期) Tablespace zl9Indexhis;

Create Index 病人挂号汇总_IX_号码 On 病人挂号汇总(号码) Tablespace zl9Indexhis;
Create Index 病人挂号汇总_IX_项目ID On 病人挂号汇总(项目ID) Tablespace zl9Indexhis;
Create Index 病人挂号汇总_IX_待转出 On 病人挂号汇总(待转出) Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_病人ID On 病人挂号记录(病人ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_接收时间 On 病人挂号记录(接收时间) Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_登记时间 On 病人挂号记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_预约时间 On 病人挂号记录(预约时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_发生时间 On 病人挂号记录(发生时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_执行时间 On 病人挂号记录(执行时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_执行状态 On 病人挂号记录(执行状态) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_待转出 On 病人挂号记录(待转出) Tablespace zl9Indexhis;
Create Index 挂号序号状态_IX_日期 On 挂号序号状态(日期) Tablespace zl9Indexhis;
Create Index 挂号序号状态_IX_登记时间 On 挂号序号状态(登记时间) Tablespace zl9Indexhis;
Create Index 挂号序号状态_IX_号码 On 挂号序号状态(号码) Tablespace zl9Indexhis;
Create Index 病人转诊记录_IX_待转出 On 病人转诊记录(待转出) Tablespace zl9Indexhis;

Create Index 人员收缴记录_IX_收款员 On 人员收缴记录(收款员) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_缴款组ID On 人员收缴记录(缴款组ID) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_小组收款ID On 人员收缴记录(小组收款ID) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_小组轧账ID On 人员收缴记录(小组轧账ID) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_财务收款ID On 人员收缴记录(财务收款ID) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_作废时间 On 人员收缴记录(作废时间) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_登记时间 On 人员收缴记录(登记时间) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_待转出 On 人员收缴记录(待转出) Tablespace zl9Indexhis;
Create Index 人员收缴明细_IX_待转出 On 人员收缴明细(待转出) Tablespace zl9Indexhis;
Create Index 人员收缴票据_IX_待转出 On 人员收缴票据(待转出) Tablespace zl9Indexhis;
Create Index 人员收缴对照_IX_记录ID On 人员收缴对照(记录ID, 性质) Tablespace zl9Indexhis;
Create Index 人员收缴对照_IX_待转出 On 人员收缴对照(待转出) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_收缴ID On 人员暂存记录(收缴ID) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_收回时间 On 人员暂存记录(收回时间) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_登记时间 On 人员暂存记录(登记时间) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_领用时间 On 人员暂存记录(领用时间) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_收款员 On 人员暂存记录(收款员) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_待转出 On 人员暂存记录(待转出) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_借款人 On 人员借款记录(借款人) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_申请时间 On 人员借款记录(申请时间) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_借出人 On 人员借款记录(借出人) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_借出时间 On 人员借款记录(借出时间) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_待转出 On 人员借款记录(待转出) Tablespace zl9Indexhis;
Create Index 病人催款记录_IX_病人ID On 病人催款记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人催款记录_IX_打印日期 On 病人催款记录(打印日期) Pctfree 5 Tablespace zl9Indexhis;

Create Index 病人结帐记录_IX_收费时间 On 病人结帐记录(收费时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人结帐记录_IX_病人id On 病人结帐记录(病人id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人结帐记录_IX_待转出 On 病人结帐记录(待转出) Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_收费细目id On 住院费用记录(收费细目id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_收入项目id On 住院费用记录(收入项目id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_医嘱序号 On 住院费用记录(医嘱序号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_结帐ID On 住院费用记录(结帐ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_登记时间 On 住院费用记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_发生时间 On 住院费用记录(发生时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_病人id On 住院费用记录(病人id,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_保险大类ID On 住院费用记录(保险大类ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_待转出 On 住院费用记录(待转出) Tablespace zl9Indexhis;

Create Index 门诊费用记录_IX_收费细目id On 门诊费用记录(收费细目id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_收入项目id On 门诊费用记录(收入项目id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_医嘱序号 On 门诊费用记录(医嘱序号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_结帐ID On 门诊费用记录(结帐ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_登记时间 On 门诊费用记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_发生时间 On 门诊费用记录(发生时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_病人id On 门诊费用记录(病人id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_保险大类ID On 门诊费用记录(保险大类ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_待转出 On 门诊费用记录(待转出) Tablespace zl9Indexhis;

Create Index 病人费用销帐_IX_申请时间 On 病人费用销帐(申请时间) Tablespace zl9Indexhis;
Create Index 费用审核记录_IX_病人ID On 费用审核记录(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 费用审核记录_IX_审核日期 On 费用审核记录(审核日期) Tablespace zl9Indexhis;
Create Index 病人费用销帐_IX_审核时间 On 病人费用销帐(审核时间) Tablespace zl9Indexhis;
Create Index 病人费用销帐_IX_状态 On 病人费用销帐(状态) Tablespace zl9Indexhis;
Create Index 病人退费申请_IX_申请时间 On 病人退费申请(申请时间) Tablespace zl9Indexhis;
Create Index 病人退费申请_IX_审核时间 On 病人退费申请(审核时间) Tablespace zl9Indexhis;
Create Index 病人费用汇总_IX_收入项目id On 病人费用汇总(收入项目id) Tablespace zl9Indexhis;
Create Index 病人结帐汇总_IX_结帐ID On 病人结帐汇总(结帐ID) Tablespace zl9Indexhis;
Create Index 病人结帐汇总_IX_收入项目id On 病人结帐汇总(收入项目id) Tablespace zl9Indexhis;
Create Index 医生收入汇总_IX_执行人 On 医生收入汇总(日期,执行人) Tablespace zl9Indexhis;
Create Index 医生收入汇总_IX_收入项目id On 医生收入汇总(收入项目id) Tablespace zl9Indexhis;
Create Index 病人未结费用_IX_病人id On 病人未结费用(病人id,主页ID) Tablespace zl9Indexhis;
Create Index 病人未结费用_IX_收入项目id On 病人未结费用(收入项目id) Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_主页ID On 病人预交记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_结帐id On 病人预交记录(结帐id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_收款时间 On 病人预交记录(收款时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_结算序号 On 病人预交记录(结算序号) Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_待转出 On 病人预交记录(待转出) Tablespace zl9Indexhis;

Create Index 票据入库记录_IX_登记人 On 票据入库记录(登记人) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据入库记录_IX_登记时间 On 票据入库记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据入库记录_IX_有无票据 On 票据入库记录(有无票据) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据报损记录_IX_报损人 On 票据报损记录(报损人) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据报损记录_IX_报损时间 On 票据报损记录(报损时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据报损记录_IX_入库ID On 票据报损记录(入库ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据领用记录_IX_领用人 On 票据领用记录(领用人) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据领用记录_IX_批次 On 票据领用记录(批次,票种) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据领用记录_IX_登记时间 On 票据领用记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据领用记录_IX_待转出 On 票据领用记录(待转出) Tablespace zl9Indexhis;
Create Index 票据打印内容_IX_NO On 票据打印内容(NO) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据使用明细_IX_领用ID On 票据使用明细(领用ID,票种,性质) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据使用明细_IX_使用时间 On 票据使用明细(使用时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据使用明细_IX_打印ID On 票据使用明细(打印ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据使用明细_IX_待转出 On 票据使用明细(待转出) Tablespace zl9Indexhis;
Create Index 票据打印明细_IX_使用ID On 票据打印明细(使用ID,NO) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据打印明细_IX_关联票号序号 On 票据打印明细(关联票号序号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据打印明细_IX_待转出 On 票据打印明细(待转出) Tablespace zl9Indexhis;
Create Index 票据打印内容_IX_待转出 On 票据打印内容(待转出) Tablespace zl9Indexhis;
Create Index 缴款成员组成_IX_成员ID On 缴款成员组成(成员ID) Pctfree 5 Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[15.药品卫材业务]]
----------------------------------------------------------------------------
Create Index 药品财务审核_IX_审核日期 On 药品财务审核(审核日期) Tablespace zl9Indexcis;
Create Index 药品采购计划_IX_库房id On 药品采购计划(库房id) Tablespace zl9Indexhis;
Create Index 药品采购计划_IX_NO On 药品采购计划(no) Tablespace zl9Indexhis;
Create Index 药品采购计划_IX_编制日期 On 药品采购计划(编制日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品采购计划_IX_审核日期 On 药品采购计划(审核日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品采购计划_IX_复核日期 On 药品采购计划(复核日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品计划内容_IX_药品id On 药品计划内容(药品id) Tablespace zl9Indexhis;
Create Index 药品退药计划_IX_药品id On 药品退药计划(药品id) Tablespace zl9Indexhis;
Create Index 药品退药计划_IX_供药单位id On 药品退药计划(供药单位id) Tablespace zl9Indexhis;
Create Index 药品退药计划_IX_填制日期 On 药品退药计划(填制日期) Tablespace zl9Indexhis;
Create Index 药品退药计划_IX_审核日期 On 药品退药计划(审核日期) Tablespace zl9Indexhis;
Create Index 材料采购计划_IX_库房id On 材料采购计划(库房id) Tablespace zl9Indexhis;
Create Index 材料采购计划_IX_NO On 材料采购计划(no) Tablespace zl9Indexhis;
Create Index 材料计划内容_IX_材料id On 材料计划内容(材料id) Tablespace zl9Indexhis;
Create Index 药品留存计划_IX_状态 On 药品留存计划(部门ID,状态) Tablespace zl9Indexhis;
Create Index 药品留存计划_IX_留存ID On 药品留存计划(留存ID) Tablespace zl9Indexhis;
Create Index 药品留存计划_IX_待转出 On 药品留存计划(待转出) Tablespace zl9Indexhis;

Create Index 药品库存_IX_药品id On 药品库存(药品id) Tablespace zl9Indexhis;
Create Index 药品库存_IX_商品条码 On 药品库存(商品条码) Tablespace zl9Indexhis;
Create Index 药品库存_IX_内部条码 On 药品库存(内部条码) Tablespace zl9Indexhis;
Create Index 药品结存_IX_药品id On 药品结存(药品id) Tablespace zl9Indexhis;
Create Index 药品留存_IX_药品id On 药品留存(药品id) Tablespace zl9Indexhis;
Create Index 药品收发汇总_IX_药品id On 药品收发汇总(药品id) Tablespace zl9Indexhis;
Create Index 药品收发汇总_IX_类别id On 药品收发汇总(类别id) Tablespace zl9Indexhis;
Create Index 未发药品记录_IX_填制日期 On 未发药品记录(填制日期) Tablespace zl9Indexhis;
Create Index 未发药品记录_IX_对方部门ID On 未发药品记录(对方部门ID) Tablespace zl9Indexhis;
Create Index 未发药品记录_IX_主页ID On 未发药品记录(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_费用id On 药品收发记录(费用id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_药品id On 药品收发记录(药品id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_入出类别id On 药品收发记录(入出类别id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_供药单位id On 药品收发记录(供药单位id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_填制日期 On 药品收发记录(填制日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_审核日期 On 药品收发记录(审核日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_价格ID On 药品收发记录(价格ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_汇总发药号 On 药品收发记录(汇总发药号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_商品条码 On 药品收发记录(商品条码) Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_内部条码 On 药品收发记录(内部条码) Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_待转出 On 药品收发记录(待转出) Tablespace zl9Indexhis;
Create Index 药品收发记录_计划id On 药品收发记录(计划id) Tablespace zl9Indexhis;
Create Index 收发记录补充信息_IX_收发ID On 收发记录补充信息(收发id) Tablespace zl9Indexhis;
Create Index 收发记录补充信息_IX_待转出 On 收发记录补充信息(待转出) Tablespace zl9Indexhis;

Create Index 药品签名记录_IX_证书ID On 药品签名记录(证书ID) Tablespace zl9Indexhis;
Create Index 药品签名记录_IX_待转出 On 药品签名记录(待转出) Tablespace zl9Indexhis;
Create Index 药品签名明细_IX_收发ID On 药品签名明细(收发ID) Tablespace zl9Indexhis;
Create Index 药品签名明细_IX_待转出 On 药品签名明细(待转出) Tablespace zl9Indexhis;

Create Index 成本价调价信息_IX_执行日期 On 成本价调价信息(执行日期) Tablespace zl9Indexhis;
Create Index 成本价调价信息_IX_库房id On 成本价调价信息(库房id) Tablespace zl9Indexhis;
Create Index 成本价调价信息_IX_药品ID On 成本价调价信息(药品ID) Tablespace zl9Indexhis;
Create Index 成本价调价信息_IX_收发id On 成本价调价信息(收发id) Tablespace zl9Indexhis;
Create Index 成本价调价信息_IX_供药单位ID On 成本价调价信息(供药单位ID) Tablespace zl9Indexhis;

Create Index 药品质量记录_IX_库房id On 药品质量记录(库房id) Tablespace zl9Indexhis;
Create Index 药品质量记录_IX_药品id On 药品质量记录(药品id) Tablespace zl9Indexhis;
Create Index 药品质量记录_IX_供药单位id On 药品质量记录(供药单位id) Tablespace zl9Indexhis;
Create Index 药品质量记录_IX_登记时间 On 药品质量记录(登记时间) Tablespace zl9Indexhis;
Create Index 药品质量记录_IX_处理时间 On 药品质量记录(处理时间) Tablespace zl9Indexhis;

Create Index 药品结存记录_IX_填制日期 On 药品结存记录(填制日期) Tablespace zl9Indexhis;
Create Index 药品结存记录_IX_结存日期 On 药品结存记录(审核日期) Tablespace zl9Indexhis;
Create Index 药品结存明细_IX_药品id On 药品结存明细(药品id) Tablespace zl9Indexhis;
Create Index 药品结存误差_IX_药品id On 药品结存误差(药品id) Tablespace zl9Indexhis;
Create Index 药品结存误差_IX_结存id On 药品结存误差(结存id) Tablespace zl9Indexhis;
Create Index 暂存药品记录_IX_病人ID On 暂存药品记录(病人ID) Tablespace zl9Indexhis;
Create Index 暂存药品记录_IX_登记时间 On 暂存药品记录(登记时间) Tablespace zl9Indexhis;
Create Index 暂存药品记录_IX_医嘱ID On 暂存药品记录(医嘱ID, 发送号) Tablespace zl9Indexhis;

Create Index 输液配药记录_IX_执行时间 On 输液配药记录(执行时间) Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_操作时间 On 输液配药记录(操作时间,操作状态) Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_摆药单号 On 输液配药记录(摆药单号) Pctfree 20 Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_瓶签号 On 输液配药记录(瓶签号) Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_待转出 On 输液配药记录(待转出) Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_打印时间 On 输液配药记录(打印时间) Tablespace zl9Indexcis;

Create Index 输液配药状态_IX_操作时间 On 输液配药状态(操作时间,操作类型) Tablespace zl9Indexhis;
Create Index 输液配药状态_IX_待转出 On 输液配药状态(待转出) Tablespace zl9Indexhis;
Create Index 输液配药内容_IX_收发ID On 输液配药内容(收发ID) Tablespace zl9Indexhis;
Create Index 输液配药内容_IX_待转出 On 输液配药内容(待转出) Tablespace zl9Indexhis;
Create Index 输液配药附费_IX_待转出 On 输液配药附费(待转出) Tablespace zl9Indexhis;

Create Index 应付记录_IX_收发ID On 应付记录(收发ID) Tablespace zl9Indexhis;
Create Index 应付记录_IX_单位ID On 应付记录(单位ID) Tablespace zl9Indexhis;
Create Index 应付记录_IX_付款序号 On 应付记录(付款序号) Tablespace zl9Indexhis;
Create Index 应付记录_IX_审核日期 On 应付记录(审核日期) Tablespace zl9Indexhis;
Create Index 应付记录_IX_发票号 On 应付记录(发票号) Tablespace zl9Indexhis;
Create Index 应付记录_IX_随货单号 On 应付记录(随货单号) Tablespace zl9Indexhis;
Create Index 付款记录_IX_单位id On 付款记录(单位id) Tablespace zl9Indexhis;
Create Index 付款记录_IX_填制日期 On 付款记录(填制日期) Tablespace zl9Indexhis;
Create Index 付款记录_IX_预审日期 On 付款记录(预审日期) Tablespace zl9Indexhis;
Create Index 付款记录_IX_审核日期 On 付款记录(审核日期) Tablespace zl9Indexhis;
Create Index 付款记录_IX_付款序号 On 付款记录(付款序号) Tablespace zl9Indexhis;

Create Index 调价汇总记录_IX_执行日期 On 调价汇总记录(执行日期) Tablespace zl9Indexhis;
Create Index 调价汇总记录_IX_填制日期 On 调价汇总记录(填制日期) Tablespace zl9Indexhis;
Create Index 成本价调价信息_IX_调价汇总号 On 成本价调价信息(调价汇总号) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[16.临床医嘱]]
----------------------------------------------------------------------------
Create Index 输血检验结果_IX_检验项目ID On 输血检验结果(检验项目ID) Tablespace zl9Indexcis;
Create Index 输血检验结果_IX_待转出 On 输血检验结果(待转出) Tablespace zl9Indexcis;
Create Index 排队记录_IX_病人ID On 排队记录(病人ID) Tablespace zl9Indexcis;
Create Index 排队记录_IX_日期 On 排队记录(日期) Tablespace zl9Indexcis;
Create Index 排队记录_IX_呼叫标志 On 排队记录(呼叫标志) Tablespace zl9Indexcis;
Create Index 座位状况记录_IX_病人ID On 座位状况记录(病人ID) Tablespace zl9Indexcis;
Create Index 座位状况记录_IX_收费细目id On 座位状况记录(收费细目id) Tablespace zl9Indexcis;
Create Index 呼叫器日志_IX_科室ID On 呼叫器日志(科室ID) Tablespace ZL9INDEXCIS;
Create Index 门诊输液操作日志_IX_科室ID On 门诊输液操作日志(科室ID) Tablespace ZL9INDEXCIS;
Create Index 门诊输液操作日志_IX_时间 On 门诊输液操作日志(时间) Tablespace ZL9INDEXCIS;
Create Index 门诊输液操作日志_IX_操作员 On 门诊输液操作日志(操作员) Tablespace ZL9INDEXCIS;
Create Index 门诊输液操作日志_IX_挂号单 On 门诊输液操作日志(挂号单) Tablespace ZL9INDEXCIS;

Create Index 病人医嘱记录_IX_相关ID On 病人医嘱记录(相关ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_主页ID On 病人医嘱记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_诊疗项目ID On 病人医嘱记录(诊疗项目id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_收费细目ID On 病人医嘱记录(收费细目id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_挂号单 On 病人医嘱记录(挂号单) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_开嘱时间 On 病人医嘱记录(开嘱时间,医嘱状态) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_开始执行时间 On 病人医嘱记录(开始执行时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_手术时间 On 病人医嘱记录(手术时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_审核状态 On 病人医嘱记录(审核状态) Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_申请序号 On 病人医嘱记录(申请序号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_配方ID On 病人医嘱记录(配方ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_待转出 On 病人医嘱记录(待转出) Tablespace zl9Indexcis;

Create Index 病人医嘱状态_IX_操作时间 On 病人医嘱状态(操作时间,操作类型) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱状态_IX_签名ID On 病人医嘱状态(签名ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱状态_IX_待转出 On 病人医嘱状态(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱计价_IX_收费细目ID On 病人医嘱计价(收费细目ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱计价_IX_待转出 On 病人医嘱计价(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_发送号 On 病人医嘱发送(发送号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_发送时间 On 病人医嘱发送(发送时间,执行状态) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_首次时间 On 病人医嘱发送(首次时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_样本条码 On 病人医嘱发送(样本条码) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_接收批次 On 病人医嘱发送(接收批次) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_待转出 On 病人医嘱发送(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱报告_IX_病历ID On 病人医嘱报告(病历ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱报告_IX_待转出 On 病人医嘱报告(待转出) Tablespace zl9Indexcis;

Create Index 病人医嘱附费_IX_NO	On 病人医嘱附费(NO,记录性质) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱附费_IX_待转出 On 病人医嘱附费(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱附件_IX_待转出 On 病人医嘱附件(待转出) Tablespace zl9Indexcis;
Create Index 医嘱签名记录_IX_证书ID On 医嘱签名记录(证书ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 医嘱签名记录_IX_待转出 On 医嘱签名记录(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱打印_IX_主页ID On 病人医嘱打印(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱打印_IX_打印时间 On 病人医嘱打印(打印时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱打印_IX_待转出 On 病人医嘱打印(待转出) Tablespace zl9Indexcis;
Create Index 医嘱执行时间_Ix_要求时间 On 医嘱执行时间(要求时间) Pctfree 5 Tablespace Zl9indexcis;
Create Index 医嘱执行时间_IX_待转出 On 医嘱执行时间(待转出) Tablespace zl9Indexcis;
Create Index 医嘱执行计价_IX_收费细目id On 医嘱执行计价(收费细目id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 医嘱执行计价_IX_待转出 On 医嘱执行计价(待转出) Tablespace zl9Indexcis;
Create Index 医嘱执行打印_IX_待转出 On 医嘱执行打印(待转出) Tablespace zl9Indexcis;
Create Index 执行打印记录_IX_流水号 On 执行打印记录(流水号) Tablespace zl9Indexcis;
Create Index 执行打印记录_IX_待转出 On 执行打印记录(待转出) Tablespace zl9Indexcis;

Create Index 病人医嘱执行_IX_流水号 On 病人医嘱执行(流水号) Tablespace zl9Indexcis;
Create Index 病人医嘱执行_IX_待转出 On 病人医嘱执行(待转出) Tablespace zl9Indexcis;
Create Index 诊疗单据打印_IX_待转出 On 诊疗单据打印(待转出) Tablespace zl9Indexcis;
Create Index 输血申请记录_IX_待转出 On 输血申请记录(待转出) Tablespace zl9Indexcis;
Create Index 报告查阅记录_IX_病历ID On 报告查阅记录(病历ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 报告查阅记录_IX_待转出 On 报告查阅记录(待转出) Tablespace zl9Indexcis;

Create Index 业务消息清单_IX_病人ID On 业务消息清单(病人ID,就诊ID) Tablespace zl9Indexcis;
Create Index 业务消息清单_IX_登记时间 On 业务消息清单(登记时间) Tablespace zl9Indexcis;
Create Index 业务消息状态_IX_阅读时间 On 业务消息状态(阅读时间) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[17.临床路径]]
----------------------------------------------------------------------------
Create Index 路径医嘱变动_IX_诊疗项目ID On 路径医嘱变动(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 路径医嘱变动_IX_收费细目ID On 路径医嘱变动(收费细目Id)   Tablespace zl9Indexcis;
Create Index 路径医嘱变动_IX_配方ID On 路径医嘱变动(配方ID)  Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_病人ID On 病人临床路径(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_科室ID On 病人临床路径(科室ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_路径ID On 病人临床路径(路径ID,版本号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_导入时间 On 病人临床路径(导入时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_疾病ID On 病人临床路径(疾病ID) Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_诊断ID On 病人临床路径(诊断ID) Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_未导入原因 On 病人临床路径(未导入原因) Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_待转出 On 病人临床路径(待转出) Tablespace zl9Indexcis;

Create Index 病人路径变异_IX_变异原因 On 病人路径变异(变异原因) Tablespace zl9Indexcis;
Create Index 病人路径变异_IX_待转出 On 病人路径变异(待转出) Tablespace zl9Indexcis;
Create Index 病人路径医嘱_IX_病人医嘱ID On 病人路径医嘱(病人医嘱ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人路径医嘱_IX_待转出 On 病人路径医嘱(待转出) Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_日期 On 病人路径执行(日期) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_路径记录ID On 病人路径执行(路径记录ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_阶段ID On 病人路径执行(阶段ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_项目ID On 病人路径执行(项目ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_图标ID On 病人路径执行(图标ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_登记时间 On 病人路径执行(登记时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_变异原因 On 病人路径执行(变异原因) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_待转出 On 病人路径执行(待转出) Tablespace zl9Indexcis;
Create Index 病人路径评估_IX_日期 On 病人路径评估(日期) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径评估_IX_登记时间 On 病人路径评估(登记时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径评估_IX_变异原因 On 病人路径评估(变异原因) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径评估_IX_阶段ID On 病人路径评估(阶段ID) Tablespace zl9Indexcis;
Create Index 病人路径评估_IX_待转出 On 病人路径评估(待转出) Tablespace zl9Indexcis;
Create Index 病人合并路径_IX_主页ID On 病人合并路径(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_版本号 On 病人合并路径(路径ID,版本号) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_疾病ID On 病人合并路径(疾病ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_当前阶段ID On 病人合并路径(当前阶段ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_前一阶段ID On 病人合并路径(前一阶段ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_首要路径阶段ID On 病人合并路径(首要路径阶段ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_首要路径记录ID On 病人合并路径(首要路径记录ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_待转出 On 病人合并路径(待转出) Tablespace zl9Indexcis;
Create Index 病人合并路径评估_IX_待转出 On 病人合并路径评估(待转出) Tablespace zl9Indexcis;

Create Index 病人路径执行_IX_合并路径阶段ID On 病人路径执行(合并路径阶段ID) Tablespace zl9Indexhis;
Create Index 病人路径执行_IX_合并路径记录ID On 病人路径执行(合并路径记录ID) Tablespace zl9Indexhis;
Create Index 病人路径指标_IX_合并路径阶段ID On 病人路径指标(合并路径记录ID) Tablespace zl9Indexhis;
Create Index 病人路径指标_IX_日期 On 病人路径指标(日期) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径指标_IX_阶段ID On 病人路径指标(阶段ID) Tablespace zl9Indexcis;
Create Index 病人路径指标_IX_待转出 On 病人路径指标(待转出) Tablespace zl9Indexcis;
Create Index 病人出径记录_IX_路径记录ID On 病人出径记录(路径记录ID) Tablespace zl9Indexhis;
Create Index 病人路径取消_IX_病人ID On 病人路径取消(病人ID,主页ID) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[18.病历业务]]
----------------------------------------------------------------------------
Create Index 电子病历记录_IX_病人ID On 电子病历记录(病人ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_文件ID On 电子病历记录(文件ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_完成时间 On 电子病历记录(完成时间,病历种类,科室id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_创建时间 On 电子病历记录(创建时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_路径执行ID On 电子病历记录(路径执行ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_待转出 On 电子病历记录(待转出) Tablespace zl9Indexcis;

Create Index 电子病历内容_IX_父ID On 电子病历内容(父ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 电子病历内容_IX_预制提纲ID On 电子病历内容(预制提纲ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 电子病历内容_IX_诊治要素ID On 电子病历内容(诊治要素ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 电子病历内容_IX_待转出 On 电子病历内容(待转出) Tablespace zl9Indexcis;
Create Index 病历变动原因_IX_病历文件id On 病历变动原因(病历文件id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动原因_IX_原因要件id On 病历变动原因(原因要件id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动原因_IX_原因要素 On 病历变动原因(原因要素) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动结果_IX_变动原因id On 病历变动结果(变动原因id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动结果_IX_结果要件id On 病历变动结果(结果要件id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动结果_IX_结果要素 On 病历变动结果(结果要素) Pctfree 5 Tablespace zl9Indexhis;
Create Index 电子病历打印_IX_病人ID On 电子病历打印(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 电子病历时机_IX_病人ID On 电子病历时机(病人ID,主页ID) Pctfree 20 Tablespace zl9Indexcis;
Create Index 电子病历时机_IX_文件ID On 电子病历时机(文件ID) Pctfree 20 Tablespace zl9Indexcis;
Create Index 疾病申报记录_IX_文档ID On 疾病申报记录(文档ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 疾病申报记录_IX_待转出 On 疾病申报记录(待转出) Tablespace zl9Indexcis;

Create Index 电子病历附件_IX_待转出 On 电子病历附件(待转出) Tablespace zl9Indexcis;
Create Index 电子病历格式_IX_待转出 On 电子病历格式(待转出) Tablespace zl9Indexcis;
Create Index 电子病历图形_IX_待转出 On 电子病历图形(待转出) Tablespace zl9Indexcis;

--临时表,不要指定表空间,Pctfree等参数
Create Index 病历时限监测_IX_病人id On 病历时限监测(病人ID,主页ID,病人来源);
Create Index 病历内容监测_IX_病人id On 病历内容监测(病人ID,主页ID,病人来源);

--病案审查归档
Create Index 病案提交记录_IX_主页ID On 病案提交记录(病人ID,主页ID) Tablespace zl9Indexcis;
Create Index 病案提交记录_IX_提交时间 On 病案提交记录(提交时间) Tablespace zl9Indexcis;
Create Index 病案打印记录_IX_主页ID On 病案打印记录(病人id,主页id) Tablespace zl9Indexcis;
Create Index 病案打印记录_IX_打印时间 On 病案打印记录(打印时间) Tablespace zl9Indexcis;
Create Index 病案审阅书签_IX_提交id On 病案审阅书签(提交id) Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_主页ID On 病案反馈记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_提交id On 病案反馈记录(提交id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_相关id On 病案反馈记录(相关id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_反馈时间 On 病案反馈记录(反馈时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_处理时间 On 病案反馈记录(处理时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_医嘱id On 病案反馈记录(医嘱id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_科室id On 病案反馈记录(科室id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_主页ID On 病案反馈历史(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_医嘱id On 病案反馈历史(医嘱id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_科室id On 病案反馈历史(科室id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_提交id On 病案反馈历史(提交id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_相关id On 病案反馈历史(相关id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_反馈时间 On 病案反馈历史(反馈时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_处理时间 On 病案反馈历史(处理时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案封存记录_IX_主页ID On 病案封存记录(病人ID,主页ID) Tablespace zl9Indexcis;
Create Index 病案借阅内容_IX_主页ID On 病案借阅内容(病人ID,主页ID) Tablespace zl9Indexcis;
Create Index 病案借阅内容_IX_借阅id On 病案借阅内容(借阅id) Tablespace zl9Indexcis;
Create Index 病案借阅内容_IX_病人id On 病案借阅内容(病人id) Tablespace zl9Indexcis;
Create Index 病案借阅人员_IX_借阅id On 病案借阅人员(借阅id) Tablespace zl9Indexcis;
Create Index 病案评分标准_IX_方案ID On 病案评分标准(方案ID) Tablespace zl9Indexcis;
Create Index 病案评分标准_IX_上级ID On 病案评分标准(上级ID) Tablespace zl9Indexcis;
Create Index 病案评分结果_IX_方案ID On 病案评分结果(方案ID) Tablespace zl9Indexcis;
Create Index 病案评分明细_IX_结果ID On 病案评分明细(主表ID) Tablespace zl9Indexcis;
Create Index 病案评分明细_IX_评分标准ID On 病案评分明细(评分标准ID) Tablespace zl9Indexcis;
Create Index 病案借阅记录_IX_登记时间 On 病案借阅记录(登记时间) Tablespace zl9Indexcis;


----------------------------------------------------------------------------
--[[19.护理业务]]
----------------------------------------------------------------------------
Create Index 病人护理文件_IX_主页ID On 病人护理文件(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理文件_IX_待转出 On 病人护理文件(待转出) Tablespace zl9Indexcis;
Create Index 病人护理记录_IX_待转出 On 病人护理记录(待转出) Tablespace zl9Indexcis;
Create Index 病人护理记录_IX_主页ID On 病人护理记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理记录_IX_发生时间 On 病人护理记录(发生时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理内容_IX_待转出 On 病人护理内容(待转出) Tablespace zl9Indexcis;
Create Index 病人护理内容_IX_记录id On 病人护理内容(记录id) Pctfree 5 Tablespace zl9Indexcis;

Create Index 病人护理数据_IX_文件ID On 病人护理数据(文件ID,发生时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理数据_IX_待转出 On 病人护理数据(待转出) Tablespace zl9Indexcis;
Create Index 病人护理明细_IX_记录ID On 病人护理明细(记录ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理明细_IX_来源ID On 病人护理明细(来源ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理明细_IX_待转出 On 病人护理明细(待转出) Tablespace zl9Indexcis;

Create Index 病人护理打印_IX_文件ID On 病人护理打印(文件ID,发生时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理打印_IX_待转出 On 病人护理打印(待转出) Tablespace zl9Indexcis;
Create Index 病区标记记录_IX_主页ID On 病区标记记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理活动项目_IX_待转出 On 病人护理活动项目(待转出) Tablespace zl9Indexcis;
Create Index 产程要素内容_IX_待转出 On 产程要素内容(待转出) Tablespace zl9Indexcis;
Create Index 病人护理要素内容_IX_待转出 On 病人护理要素内容(待转出) Tablespace zl9Indexcis;

----------------------------------------------------------------------------

--[[20.检验业务]]

----------------------------------------------------------------------------
Create Index 检验流水线标本_IX_待转出 On 检验流水线标本(待转出) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验流水线指标_IX_待转出 On 检验流水线指标(待转出) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验流水线指标_IX_项目ID On 检验流水线指标(项目ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验流水线标本_IX_标本ID On 检验流水线标本(标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验流水线指标_IX_标本ID On 检验流水线指标(标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_医嘱ID On 检验标本记录(医嘱id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_检验时间 On 检验标本记录(检验时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_申请时间 On 检验标本记录(申请时间) Pctfree 5 Tablespace ZL9INDEXCIS;
Create Index 检验标本记录_IX_审核时间 On 检验标本记录(审核时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_年龄数字 On 检验标本记录(年龄数字) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_挂号单 On 检验标本记录(挂号单) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_合并ID On 检验标本记录(合并ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_主页ID On 检验标本记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_标识号 On 检验标本记录(标识号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_待转出 On 检验标本记录(待转出) Tablespace zl9Indexcis;

Create Index 检验普通结果_IX_细菌ID On 检验普通结果(细菌ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_仪器ID On 检验普通结果(仪器ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_检验标本ID On 检验普通结果(检验标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_药敏组ID On 检验普通结果(药敏组ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_待转出 On 检验普通结果(待转出) Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_标本id On 检验项目分布(标本id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_项目id On 检验项目分布(项目id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_医嘱id On 检验项目分布(医嘱id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_细菌ID On 检验项目分布(细菌id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_待转出 On 检验项目分布(待转出) Tablespace zl9Indexcis;

Create Index 检验质控记录_IX_仪器ID On 检验质控记录(仪器ID) Tablespace zl9Indexcis;
Create Index 检验质控记录_IX_质控品ID On 检验质控记录(质控品ID) Tablespace zl9Indexcis;
Create Index 检验质控记录_IX_待转出 On 检验质控记录(待转出) Tablespace zl9Indexcis;
Create Index 检验图像结果_IX_标本id On 检验图像结果(标本id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验酶标记录_IX_测试时间 On 检验酶标记录(测试时间) Tablespace zl9Indexhis;
Create Index 检验操作记录_IX_标本id On 检验操作记录(标本id) Tablespace zl9Indexcis;
Create Index 检验操作记录_IX_待转出 On 检验操作记录(待转出) Tablespace zl9Indexhis;
Create Index 检验分析记录_IX_标本ID On 检验分析记录(标本ID) Tablespace zl9Indexhis;
Create Index 检验分析记录_IX_用途 On 检验分析记录(用途) Tablespace zl9Indexhis;
Create Index 检验分析记录_IX_待转出 On 检验分析记录(待转出) Tablespace zl9Indexcis;
Create Index 检验拒收记录_IX_医嘱ID On 检验拒收记录(医嘱ID) Tablespace zl9Indexcis;
Create Index 检验拒收记录_IX_待转出 On 检验拒收记录(待转出) Tablespace zl9Indexcis;

Create Index 检验申请项目_IX_待转出 On 检验申请项目(待转出) Tablespace zl9Indexcis;
Create Index 检验试剂记录_IX_待转出 On 检验试剂记录(待转出) Tablespace zl9Indexcis;
Create Index 检验质控报告_IX_待转出 On 检验质控报告(待转出) Tablespace zl9Indexcis;
Create Index 检验药敏结果_IX_待转出 On 检验药敏结果(待转出) Tablespace zl9Indexcis;
Create Index 检验签名记录_IX_待转出 On 检验签名记录(待转出) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[21.检查业务]]
----------------------------------------------------------------------------
Create Index 影像报告驳回_IX_医嘱ID On 影像报告驳回(医嘱ID,病历ID) Tablespace ZL9INDEXCIS;
Create Index 影像报告驳回_IX_待转出 On 影像报告驳回(待转出) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_检查号 On 影像检查记录(检查号, 影像类别) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_位置一 On 影像检查记录(位置一) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_位置二 On 影像检查记录(位置二) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_位置三 On 影像检查记录(位置三) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_接收日期 On 影像检查记录(接收日期) Pctfree 5 Tablespace zl9Indexcis;
Create Index 影像检查记录_Ix_执行科室id On 影像检查记录(执行科室id) Pctfree 5 Tablespace Zl9Indexcis;
Create Index 影像检查记录_IX_关联ID On 影像检查记录(关联ID) Pctfree 5 Tablespace Zl9Indexcis;
Create Index 影像检查记录_IX_待转出 On 影像检查记录(待转出) Tablespace zl9Indexcis;

Create Index 影像临时记录_IX_检查号 On 影像临时记录(检查号, 影像类别) Tablespace zl9Indexcis;
Create Index 影像临时记录_IX_位置一 On 影像临时记录(位置一) Tablespace zl9Indexcis;
Create Index 影像临时记录_IX_位置二 On 影像临时记录(位置二) Tablespace zl9Indexcis;
Create Index 影像临时记录_IX_位置三 On 影像临时记录(位置三) Tablespace zl9Indexcis;
Create Index 影像临时记录_IX_接收日期 On 影像临时记录(接收日期) Tablespace zl9Indexcis;
Create Index 胶片打印记录_IX_相关ID On 胶片打印记录(相关ID) Tablespace zl9Indexhis;
Create Index 胶片打印记录_IX_打印时间 On 胶片打印记录(打印时间) Tablespace zl9Indexhis;
Create Index 影像收藏类别_IX_上级ID On 影像收藏类别(上级ID) Tablespace zl9Indexcis;
Create Index 影像申请单图像_IX_医嘱ID On 影像申请单图像(医嘱ID) Tablespace zl9Indexcis;
Create Index 影像申请单图像_IX_待转出 On 影像申请单图像(待转出) Tablespace zl9Indexcis;
Create Index 影像收藏内容_IX_医嘱ID On 影像收藏内容(医嘱ID) Tablespace zl9Indexcis;
Create Index 影像收藏内容_IX_待转出 On 影像收藏内容(待转出) Tablespace zl9Indexcis;

Create Index 影像检查图象_IX_待转出 On 影像检查图象(待转出) Tablespace zl9Indexcis;
Create Index 影像检查序列_IX_待转出 On 影像检查序列(待转出) Tablespace zl9Indexcis;
Create Index 影像危急值记录_IX_待转出 On 影像危急值记录(待转出) Tablespace zl9Indexcis;

Create Index 病理检查信息_IX_医嘱ID On 病理检查信息(医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理检查信息_IX_报到时间 On 病理检查信息(报到时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理质量信息_IX_病理医嘱ID On 病理质量信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理标本信息_IX_医嘱ID On 病理标本信息(医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理标本信息_IX_送检ID On 病理标本信息(送检ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理送检信息_IX_医嘱ID On 病理送检信息(医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理申请信息_IX_病理医嘱ID On 病理申请信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理申请信息_IX_申请时间 On 病理申请信息(申请时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理取材信息_IX_病理医嘱ID On 病理取材信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理取材信息_IX_申请ID On 病理取材信息(申请ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理取材信息_IX_标本ID On 病理取材信息(标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理取材信息_IX_取材时间 On 病理取材信息(取材时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理脱钙信息_IX_标本ID On 病理脱钙信息(标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理制片信息_IX_材块ID On 病理制片信息(材块ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理制片信息_IX_申请ID On 病理制片信息(申请ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理制片信息_IX_病理医嘱ID On 病理制片信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理制片信息_IX_制片时间 On 病理制片信息(制片时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理过程报告_IX_病理医嘱ID On 病理过程报告(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理特检信息_IX_申请ID On 病理特检信息(申请ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理特检信息_IX_抗体ID On 病理特检信息(抗体ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理特检信息_IX_材块ID On 病理特检信息(材块ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理特检信息_IX_完成时间 On 病理特检信息(完成时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理报告延迟_IX_病理医嘱ID On 病理报告延迟(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理会诊信息_IX_病理医嘱ID On 病理会诊信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理抗体反馈_IX_抗体ID On 病理抗体反馈(抗体ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理套餐关联_IX_抗体ID On 病理套餐关联(抗体ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理档案信息_IX_分类ID On 病理档案信息(分类ID) Tablespace zl9Indexcis;
Create Index 病理档案信息_IX_创建日期 On 病理档案信息(创建日期) Tablespace zl9Indexcis;
Create Index 病理归档信息_IX_材块ID On 病理归档信息(材块ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index 病理归档信息_IX_制片ID On 病理归档信息(制片ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index 病理归档信息_IX_特检ID On 病理归档信息(特检ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index 病理归档信息_IX_档案ID On 病理归档信息(档案ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index 病理借阅信息_IX_借阅时间 On 病理借阅信息(借阅时间) TableSpace zl9Indexcis;
Create Index 病理借阅信息_IX_证件号码 On 病理借阅信息(证件号码,证件类型) TableSpace zl9Indexcis;
Create Index 病理遗失信息_IX_借阅ID On 病理遗失信息(借阅ID) TableSpace zl9Indexcis;
Create Index 病理遗失信息_IX_归档ID On 病理遗失信息(归档ID) TableSpace zl9Indexcis;
Create Index 病理遗失信息_IX_遗失日期 On 病理遗失信息(遗失日期) TableSpace zl9Indexcis;
Create Index 病理归还信息_IX_借阅ID On 病理归还信息(借阅ID) TableSpace zl9Indexcis;
Create Index 病理借阅关联_IX_借阅ID On 病理借阅关联(借阅ID) TableSpace zl9Indexcis;
Create Index 病理玻片信息_IX_来源ID On 病理玻片信息(来源ID,材块Id,病理医嘱ID) Tablespace zl9Indexcis;
