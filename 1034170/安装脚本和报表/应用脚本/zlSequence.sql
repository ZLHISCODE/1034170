--分类目录
--1.公共基础,2.医保基础,3.病人病案基础,4.费用基础,5.药品卫材基础
--6.临床基础,7.临床路径基础,8.病历基础,9.护理基础,10.检验基础
--11.检查基础,12.医保业务,13.病人病案业务,14.费用业务,15.药品卫材业务
--16.临床医嘱,17.临床路径,18.病历业务,19.护理业务,20.检验业务,21.检查业务

----------------------------------------------------------------------------
--[[1.公共基础]]
----------------------------------------------------------------------------
Create Sequence 部门表_ID Start With 1;
Create Sequence 人员表_ID Start With 1;
Create Sequence 人员证书记录_ID Start With 1;
Create Sequence 排队叫号队列_ID Start WITH 1;
create sequence 排队叫号队列_排队序号 Start with 1;
Create Sequence 排队语音呼叫_ID Start WITH 1 Cache 100;
Create Sequence 合约单位_ID Start With 1;

----------------------------------------------------------------------------
--[[2.医保基础]]
----------------------------------------------------------------------------
Create Sequence 保险支付大类_ID Start With 1;
Create Sequence 保险病种_ID Start With 1;

----------------------------------------------------------------------------
--[[3.病人病案基础]]
----------------------------------------------------------------------------
Create Sequence 疾病编码分类_ID Start With 1;
Create Sequence 疾病编码目录_ID Start With 1;

----------------------------------------------------------------------------
--[[4.费用基础]]
----------------------------------------------------------------------------
Create Sequence 医疗卡类别_ID Start With 1;
Create Sequence 挂号安排_ID Start With 1;
Create Sequence 挂号安排计划_ID Start WITH 1;
Create Sequence 收入项目_ID Start With 1;
Create Sequence 收费分类目录_ID Start With 1;
Create Sequence 收费项目目录_ID Start With 1;
Create Sequence 收费价目_ID Start With 1;
Create Sequence 收费记帐单_ID Start With 1;
Create Sequence 成套项目分类_ID Start With 1;
Create Sequence 成套收费项目_ID Start With 1;
Create Sequence 收费项目组成_ID Start With 1;
Create Sequence 消费卡目录_ID Start With 1;

----------------------------------------------------------------------------
--[[5.药品卫材基础]]
----------------------------------------------------------------------------
Create Sequence 药品入出类别_ID Start With 1;
Create Sequence 供应商_ID Start With 1;

----------------------------------------------------------------------------
--[[6.临床基础]]
----------------------------------------------------------------------------
Create Sequence 抗菌药物抽样记录_ID Start With 1;
Create Sequence 临床医疗小组_ID Start With 1;
Create Sequence 诊疗分类目录_ID Start With 1;
Create Sequence 诊疗项目目录_ID Start With 1;
Create Sequence 诊疗项目部位_ID Start With 1;
Create Sequence 疾病诊断分类_ID Start With 1;
Create Sequence 疾病诊断目录_ID Start With 1;
Create Sequence 诊疗参考分类_ID Start With 1;
Create Sequence 诊疗参考目录_ID Start With 1;

----------------------------------------------------------------------------
--[[7.临床路径基础]]
----------------------------------------------------------------------------
Create Sequence 临床路径图标_ID Start With 1;
Create Sequence 临床路径目录_ID Start With 1;
Create Sequence 临床路径阶段_ID Start With 1;
Create Sequence 临床路径项目_ID Start With 1;
Create Sequence 路径医嘱内容_ID Start With 1;
Create Sequence 临床路径评估_ID Start With 1;
Create Sequence 路径评估指标_ID Start With 1;
Create Sequence 路径报表目录_ID Start With 1;
Create Sequence 路径报表文件_ID Start With 1;
Create Sequence 临床路径分支_ID Start With 1;
Create Sequence 标准路径目录_ID Start With 1;

----------------------------------------------------------------------------
--[[8.病历基础]]
----------------------------------------------------------------------------
Create Sequence 诊治所见分类_ID Start With 1;
Create Sequence 诊治所见项目_ID Start With 1;
Create Sequence 病历附项模板_ID start with 1;
Create Sequence 病历文件列表_ID Start With 1;
Create Sequence 病历文件结构_ID Start With 1;
Create Sequence 病历词句分类_ID Start With 1;
Create Sequence 病历词句示范_ID Start With 1;
Create Sequence 病历范文目录_ID Start With 1;
Create Sequence 病历范文内容_ID Start With 1;
Create Sequence 病历范文包_ID Start With 1;

Create Sequence 病案评分方案_ID start with 1;
Create Sequence 病案评分标准_ID start with 1;
Create Sequence 病案评分结果_ID start with 1;
Create Sequence 病案评分明细_ID start with 1;
Create Sequence 病案审查方案_ID Start With 1;
Create Sequence 病案审查分类_ID Start With 1;
Create Sequence 病案审查目录_ID Start With 1;

----------------------------------------------------------------------------
--[[9.护理基础]]
----------------------------------------------------------------------------
CREATE Sequence 病区公告栏样式_ID START WITH 1;

----------------------------------------------------------------------------
--[[10.检验基础]]
----------------------------------------------------------------------------
Create Sequence 检验抗生素组_ID Start With 1;
Create Sequence 检验细菌类型_ID Start With 1;
Create Sequence 检验仪器_ID Start With 1;
Create Sequence 检验用抗生素_ID Start With 1;
Create Sequence 检验质控品_ID Start With 1;
Create Sequence 检验细菌_ID Start With 1;
Create Sequence 检验报告_ID Start With 1;
Create Sequence 检验质控规则_ID Start With 1;
Create Sequence 检验模板目录_ID Start With 1;
Create Sequence 检验模板内容_ID Start With 1;
Create Sequence 检验仪器规则_ID Start With 1;
Create Sequence 检验报告项目_ID start with 1;
Create Sequence 检验试剂关系_ID start with 1;
Create Sequence 检验项目参考_ID start with 1;
Create Sequence 检验审核规则_ID Start With 1;
Create Sequence 检验酶标模板_ID Start With 1;

----------------------------------------------------------------------------
--[[11.检查基础]]
----------------------------------------------------------------------------
create sequence 影像执行分组_ID start with 1; 
create sequence 影像分组关联_ID start with 1;  
Create Sequence 快捷功能关联_ID Start With 1;
Create Sequence 影像查询方案_ID Start With 1;
Create Sequence 影像查询配置_ID Start With 1;
Create Sequence 快捷功能信息_ID Start With 1;
Create Sequence 影像滤镜模板_ID Start With 1;
Create Sequence 影像图像信息表_ID Start With 1;
Create Sequence 影像收藏类别_ID Start With 1;
Create Sequence 影像颜色清单_序号 Start With 1;
Create Sequence 影像MWL部位对码_ID Start With 1;
Create Sequence 影像DICOM服务对_服务ID Start With 1;
Create Sequence 影像流程参数_ID Start With 1;
Create Sequence 影像DICOM服务参数_服务参数ID Start With 1;
Create Sequence 影像图像消隐表_ID Start With 1;
Create Sequence 影像屏幕布局_ID Start With 1;
Create Sequence 影像打印机设置_ID Start With 1;
Create Sequence 影像预设窗宽窗位_ID Start With 1;
Create Sequence 影像接入设备_ID Start With 1;

Create Sequence 病理号码规则_ID Start With 1;
Create Sequence 病理号码记录_ID Start With 1;
Create Sequence 病理检查标本_ID Start With 1;
Create Sequence 病理套餐信息_套餐ID Start With 1;
Create Sequence 病理套餐关联_ID Start With 1;
Create Sequence 病理档案分类_ID Start With 1;

----------------------------------------------------------------------------
--[[12.医保业务]]
----------------------------------------------------------------------------


----------------------------------------------------------------------------
--[[13.病人病案业务]]
----------------------------------------------------------------------------
Create Sequence 病人信息_ID Start With 1;
Create Sequence 病人变动记录_Id Start With 1;
Create Sequence 病人过敏记录_ID Start With 1;
Create Sequence 病人诊断记录_ID Start With 1;
Create Sequence 病人手麻记录_ID Start With 1;

----------------------------------------------------------------------------
--[[14.费用业务]]
----------------------------------------------------------------------------
Create Sequence 病人结帐记录_ID Start With 1;
Create Sequence 病人缴款记录_ID Start With 1;
Create Sequence 病人催款记录_ID Start With 1;
Create Sequence 病人费用记录_ID Start With 1 Cache 100;
Create Sequence 病人挂号记录_ID Start With 1;
Create Sequence 病人预交记录_ID Start With 1;
Create Sequence 病人备注信息_ID Start With 1;
Create Sequence 票据入库记录_ID Start With 1;
Create Sequence 票据报损记录_ID Start With 1;
Create Sequence 票据领用记录_ID Start With 1;
Create Sequence 票据使用明细_ID Start With 1;
Create Sequence 票据打印内容_ID Start With 1;
Create Sequence 财务缴款分组_ID Start With 1;
Create Sequence 人员收缴记录_ID Start With 1;
Create Sequence 人员暂存记录_ID Start With 1;
Create Sequence 人员借款记录_ID Start WITH 1;
Create Sequence 病人医疗卡变动_ID Start WITH 1;
Create Sequence 消费卡充值记录_ID Start With 1;
Create Sequence 病人卡结算记录_ID Start With 1;

----------------------------------------------------------------------------
--[[15.药品卫材业务]]
----------------------------------------------------------------------------
Create Sequence 药品采购计划_ID Start With 1;
Create Sequence 药品退药计划_ID Start With 1;
Create Sequence 材料采购计划_ID Start With 1;
Create Sequence 药品签名记录_ID Start With 1;
Create Sequence 成本价调价信息_ID Start With 1;
Create Sequence 药品收发记录_ID Start With 1 Cache 100;
Create Sequence 药品质量记录_ID Start With 1;
Create Sequence 药品结存记录_ID Start With 1;
Create Sequence 药品结存误差_ID Start With 1;
Create Sequence 输液配药记录_ID Start With 1;
Create Sequence 应付记录_ID Start With 1;
Create Sequence 付款记录_ID Start With 1;

----------------------------------------------------------------------------
--[[16.临床医嘱]]
----------------------------------------------------------------------------
Create Sequence 病人医嘱记录_ID Start With 1 Cache 100;
Create Sequence 病人医嘱记录_申请序号 Start With 1 Cache 100;
Create Sequence 病人医嘱发送_接收批次 Start With 1;
Create Sequence 病人医嘱执行_流水号 Start With 1;
Create Sequence 医嘱签名记录_ID Start With 1;
Create Sequence 门诊输液操作日志_ID Start With 1;
Create Sequence 呼叫器日志_ID Start With 1;
Create Sequence 业务消息清单_ID Start With 1;

----------------------------------------------------------------------------
--[[17.临床路径]]
----------------------------------------------------------------------------
Create Sequence 病人临床路径_ID Start With 1;
Create Sequence 病人路径执行_ID Start With 1;
Create Sequence 病人合并路径_ID Start With 1;

----------------------------------------------------------------------------
--[[18.病历业务]]
----------------------------------------------------------------------------
Create Sequence 电子病历记录_ID Start With 1;
Create Sequence 电子病历内容_ID Start With 1 Cache 100;
Create Sequence 病历变动原因_ID Start With 1;
Create Sequence 病历变动结果_ID Start With 1;
Create Sequence 电子病历时机_ID Start With 1 Cache 100;
Create Sequence 电子病历打印_ID Start With 1;

Create Sequence 病案接收记录_Id Start With 1;
Create Sequence 病案提交记录_ID start with 1;
Create Sequence 病案反馈记录_ID start with 1;
Create Sequence 病案借阅记录_ID start with 1;
Create Sequence 病案封存记录_ID start with 1;

----------------------------------------------------------------------------
--[[19.护理业务]]
----------------------------------------------------------------------------
Create Sequence 病人护理记录_ID Start With 1;
Create Sequence 病人护理内容_ID Start With 1;
Create Sequence 病人护理文件_ID Start WITH 1;
Create Sequence 病人护理数据_ID Start WITH 1 Cache 100;
Create Sequence 病人护理明细_ID Start WITH 1 Cache 100;

----------------------------------------------------------------------------
--[[20.检验业务]]
----------------------------------------------------------------------------
Create Sequence 检验流水线标本_ID Start With 1;
Create Sequence 检验流水线指标_ID Start With 1;
Create Sequence 检验标本记录_ID Start With 1 Cache 100;
Create Sequence 检验普通结果_ID Start With 1 Cache 100;
Create Sequence 检验项目分布_ID start with 1 Cache 100;
Create Sequence 检验图像结果_ID start with 1;
Create Sequence 检验酶标记录_ID Start With 1;
Create Sequence 检验操作记录_ID start with 1;
Create Sequence 检验拒收记录_ID Start With 1;
Create Sequence 检验分析记录_ID Start With 1;

----------------------------------------------------------------------------
--[[21.检查业务]]
----------------------------------------------------------------------------
Create Sequence 影像危急值记录_id start with 1;
Create Sequence 影像报告驳回_ID Start With 1;
Create Sequence 影像归档作业_ID Start With 1;
Create Sequence 影像检查UID序号_ID Start With 1;
Create Sequence 胶片打印记录_ID Start With 1;
Create Sequence 影像收藏内容_ID Start With 1;
Create Sequence 影像申请单图像_ID Start With 1;

Create Sequence 病理档案信息_ID Start With 1;
Create Sequence 病理检查信息_病理医嘱ID Start With 1;
Create Sequence 病理质量信息_ID Start With 1;
Create Sequence 病理标本信息_标本ID Start With 1;
Create Sequence 病理送检信息_ID Start With 1;
Create Sequence 病理申请信息_申请ID Start With 1;
Create Sequence 病理取材信息_材块ID Start With 1;
Create Sequence 病理脱钙信息_ID Start With 1;
Create Sequence 病理制片信息_ID Start With 1;
Create Sequence 病理过程报告_ID Start With 1;
Create Sequence 病理抗体信息_抗体ID Start With 1;
Create Sequence 病理特检信息_ID Start With 1;
Create Sequence 病理报告延迟_ID Start With 1;
Create Sequence 病理会诊信息_ID Start With 1;
Create Sequence 病理抗体反馈_ID Start With 1;
Create Sequence 病理归档信息_ID Start With 1;
Create Sequence 病理借阅信息_ID Start With 1;
Create Sequence 病理遗失信息_ID Start With 1;
Create Sequence 病理归还信息_ID Start With 1;
Create Sequence 病理玻片信息_ID Start With 1;



