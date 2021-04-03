Attribute VB_Name = "mdl广元旺苍"
Option Explicit

Public Enum 业务类型_广元旺苍
    获得社保机构_旺苍 = 0
    获得参保人员资料_旺苍
    获得人员资料_医保号_旺苍
    获得人员资料_读卡_旺苍
    
    获取帐户余额_旺苍
    检查拔号连接_旺苍
    建立拔号连接_旺苍
    断开拔号连接_旺苍
    个人帐户消费_旺苍
    个人帐户消费_金额_旺苍
    初始化_旺苍
    消费冲正_旺苍
    下载交易记录_旺苍
    提取基础资料_旺苍
    门诊预处理_旺苍
    修改密码_旺苍
    
    获得社保机构_住院_旺苍
    入院登记_旺苍
    取消入院登记_旺苍
    获取处方记录号_旺苍
    增加处方单据_旺苍
    单条处方传输_旺苍
    增加处方明细_旺苍
    出院结算_旺苍
    取消出院结算_旺苍
    根据住院号获取记录号_旺苍
    打印结算报表_旺苍
    住院病人跨月重提_旺苍
    
    申报项目_资阳
    提取项目_资阳
End Enum
Private Type InitbaseInfor
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    医院编码 As String                      '初始医院编码
    机构编码 As String                      '默认的社保机构编码
    明细时实上传 As Boolean
    数据不等不可结算  As Boolean
    适用地区 As String
    
End Type
Public InitInfor_广元旺苍 As InitbaseInfor

Private Type 病人身份
    医保卡号    As String
    医保证号    As String
    身份证号码  As String
    记录号      As String
    姓名        As String
    性别        As String
    出生日期    As String
    年龄        As Integer
    单位名称    As String
    机构编码    As String
    
    帐户余额    As String
    费用总额    As Double
    密码        As String
    社保中心    As Long
    病人ID      As Long
    直输金额    As Boolean
    
    个人ID As String
    
    参加工作年月 As String
    退休年月 As String
    职务级别 As String
    职称级别 As String
    人员分类 As String
    异地居住标志 As String
    单位ID As String
    年月 As String
    住院性质 As String
    基本医疗标志 As String
    补充医疗标志 As String
    公务员标志  As String
    基本待遇状态 As String
    补充待遇状态 As String
    公务员待遇状态  As String
    年内住院次数    As String
    年内已报销金额 As String
    缴费年限    As String
    提取时间    As String
    住院记录号  As String
    
    str住院号 As String
    str住院信息  As String
    个人帐户支付 As Double '门诊虚结时保存
End Type

Private Type 结算数据
    卡号 As String
    姓名    As String
    消费前帐户余额 As Double
    个人帐户支付金额 As Double
    自费金额 As Double
    消费后帐户余额 As Double
    交易时间  As String
    前端单据号  As String
    中心单据号  As String
    处方号  As String
    操作员姓名  As String
    前端名称  As String
    人员分类 As String
    
    结帐ID As Long
    结算标志 As Byte    '0-门诊,1-住院
    基本报销金额 As Double
    补充报销金额 As Double
    公务员报销金额 As Double

End Type
Private g结算数据 As 结算数据
Public g病人身份_广元旺苍 As 病人身份
Public gcnOracle_广元旺苍 As ADODB.Connection     '中间库连接

Private gbln检查连接 As Boolean
Private gbln已经初始 As Boolean             '已经被初始化了.

'1.获得社保机构_旺苍编号和名称列表
Private Declare Function GetSBJGLB Lib "CDGK_GRZH.dll" Alias "GETSBJGLB" () As String
'===============================================================================================================
'原型:FUNCTION GETSBJGLB:PCHAR
'功能: 获得社保机构_旺苍编号和名称列表
'入口参数: 无
'出口参数: 无
'返回: A社保机构编号+列间隔符+A社保机构名称+行间隔符+B社保机构编号+列间隔符+B社保机构名称+……
'===============================================================================================================

'2．获得参保人员的基本资料
'   A.门诊
Private Declare Function GETKZL Lib "CDGK_GRZH.dll" () As String
'===============================================================================================================
'原型:FUNCTION GETKZL:PCHAR
'功能: 获得参保人员的基本资料
'入口参数:
'出口参数: 无
'返回: OK(或错误信息)@$医保卡号||医保证号||个人记录号||姓名||身份证号码||单位名称||性别||出生日期（格式：YYYY-MM-DD）
'===============================================================================================================

'   B.住院(获得参保人员的基本资料(输入医保证号))
Private Declare Function GETRYJBZL Lib "CDGK_HOS.dll" (ByVal str医保证号 As String, ByVal str机构编码 As String) As String
'===============================================================================================================
'原型:FUNCTION GETRYJBZL(AYBZH, ABXJGBH:PCHAR):PCHAR;
'功能:通过输入医保证号从社保机构医保数据库提取医保病人的基本资料。
'入口参数:AYBZH   PCHAR   参保人员的医保证号
'         ABXJGBH PCHAR   参保人员所在的保险机构编号
'出口参数: 无
'返回: OK(或错误信息)@$个人ID||社保编号||姓名||性别||出生日期（格式：YYYY-MM-DD）||参加工作年月||退休年月||职务级别||职称级别||人员分类||异地居住标志||单位ID||单位名称||年龄||年月||医保证号||住院性质||基本医疗标志
'===============================================================================================================

'   C.住院(获得参保人员的基本资料(直接读医保卡))
Private Declare Function GETRYJBZL_BYYBK Lib "CDGK_HOS.dll" (ByVal str机构编码 As String) As String
'===============================================================================================================
'原型:FUNCTION GETRYJBZL_BYYBK(ABXJGBH:PCHAR):PCHAR;
'功能:通过读取医保卡数据从社保机构医保数据库提取医保病人的基本资料。
'入口参数:ABXJGBH PCHAR   参保人员所在的保险机构编号
'出口参数: 无
'返回: OK(或错误信息)@$个人ID||社保编号||姓名||性别||出生日期（格式：YYYY-MM-DD）||参加工作年月||退休年月||职务级别||职称级别||人员分类||异地居住标志||单位ID||单位名称||年龄||年月||医保证号||住院性质||基本医疗标志
'===============================================================================================================


'3.个人帐户余额查询
Private Declare Function GETZHYE Lib "CDGK_GRZH.dll" (ByVal str机构编码 As String, ByVal strPassWord As String) As String
'===============================================================================================================
'原型:FUNCTION GETZHYE(YBJGBH,CPASSWORD:PCHAR):PCHAR
'功能: 获得持卡人员个人帐户余额
'入口参数:YBJGBH  PCHAR   保险机构编号
'         CPASSWORD   PCHAR   持卡人卡密码
'出口参数: 无
'返回:  OK(或错误信息)@$个人帐户余额
'===============================================================================================================

'4.检测拔号连接是否连接成功
Private Declare Function CheckCon Lib "CDGK_GRZH.dll" () As String
'===============================================================================================================
'原型:FUNCTION CHECKCON:PCHAR;
'功能:检测拔号连接是否连接成功
'入口参数:
'返回:OK或错误信息
'===============================================================================================================

'5.建立拔号连接
Private Declare Function RasDial Lib "CDGK_GRZH.dll" (ByVal str机构代码 As String) As String
'===============================================================================================================
'原型:FUNCTION RASDIAL(SBXJGBH:PCHAR):PCHAR
'功能:拔号至选择的社保局，与其建立连接
'入口参数:SBXJGBH PCHAR   保险机构编号
'返回:  成功    川大金键HIS拔号器状态栏显示"连接"
'       失败 错误信息
'===============================================================================================================

'6.断开与社保局的连接
Private Declare Function DisDial Lib "CDGK_GRZH.dll" () As String
'===============================================================================================================
'原型:FUNCTION DISDIAL:PCHAR
'功能:拔号至选择的社保局，与其建立连接
'入口参数:
'返回:
'===============================================================================================================

'7.个人帐户销费
Private Declare Function GRZHXF_CF Lib "CDGK_GRZH.dll" (ByVal str机构编号 As String, ByVal str处方号 As String, _
            ByVal str明细数据 As String, ByVal strPassWord As String, ByVal str操作员 As String) As String
'===============================================================================================================
'原型:Function GRZHXF_CF()(YBJGBH,CFH:PCHAR;CFMXDATA:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'功能:进行个人帐户消费
'入口参数:YBJGBH  PCHAR   保险机构编号
'        CFH PCHAR   处方号
'        CFMXDATA    PCHAR   处方明细数据    格式说明：处方1(医保药品编号+列间隔符+单价+列间隔符数量+)+行间隔符+        ……        处方N(医保药品编号+列间隔符+单价+列间隔符+数量
'        CPASSWORD   PCHAR   持卡人卡密码
'        CCZYXM  PCHAR   操作员姓名
'返回:个人帐户消费信息(OK@$个人帐户消费信息)
'   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
'===============================================================================================================


'8.个人帐户消费（直接输入消费金额）

Private Declare Function GRZHXF_JE Lib "CDGK_GRZH.dll" (ByVal str机构编号 As String, ByVal str金额 As String, _
             ByVal strPassWord As String, ByVal str操作员 As String) As String
'===============================================================================================================
'原型:FUNCTION GRZHXF_JE(YBJGBH,XFJE:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'功能:进行个人帐户消费
'入口参数:YBJGBH  PCHAR   保险机构编号
'    XFJE    PCHAR   消费金额(保证为小数，并且保留二位小数)
'    CPASSWORD   PCHAR   持卡人卡密码
'    CCZYXM  PCHAR   操作员姓名
'返回:个人帐户消费信息(OK@$个人帐户消费信息)
'   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
'===============================================================================================================

'个人预帐户消费
Private Declare Function GRZHXF_CFPRE Lib "CDGK_GRZH.dll" (ByVal str机构编号 As String, ByVal str明细数据 As String, _
             ByVal strPassWord As String) As String
'===============================================================================================================
'原型:FUNCTION GRZHXF_CFPRE(YBJGBH,CFMXDATA,CPASSWORD:PCHAR):PCHAR
'功能:个人预帐户消费
'入口参数:YBJGBH  PCHAR   保险机构编号
'    CFMXDATA    PCHAR   处方明细数据    格式说明：处方1(医保药品编号+列间隔符+单价+列间隔符数量+)+行间隔符+
'    ……
'    处方N(医保药品编号+列间隔符+单价+列间隔符+数量
'    CPASSWORD   PCHAR   持卡人卡密码
'返回:OK@$个人帐户支付金额@$自付金额
'===============================================================================================================

'9.消费冲正

Private Declare Function XFCZ Lib "CDGK_GRZH.dll" (ByVal str机构编号 As String, ByVal str中心单据号 As String, _
             ByVal strPassWord As String, ByVal str操作员 As String) As String
'===============================================================================================================
'原型:FUNCTION XFCZ(YBJGBH ，CZXDJH:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'功能:对已经消费的记录进行冲正。
'入口参数:YBJGBH  PCHAR   保险机构编号
'        cZXDJH  PCHAR   中心单据号(消费时返回)
'        CPASSWORD   PCHAR   持卡人卡密码
'        CCZYXM  PCHAR   操作员姓名
'返回:个人帐户消费信息(OK@$个人帐户消费信息)
'   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
'===============================================================================================================




'10.持卡人员进行卡密码修改

Private Declare Function CHANGPASSWORD Lib "CDGK_GRZH.dll" (ByVal str机构编号 As String, ByVal str旧密码 As String, _
             ByVal str新密码 As String) As String
'===============================================================================================================
'原型:FUNCTION CHANGPASSWORD(YBJGBH ,COLDPASS,CNEWPASS:PCHAR):PCHAR
'功能:持卡人员进行卡密码修改
'入口参数:YBJGBH  PCHAR   保险机构编号
'    COLDPASS    PCHAR   旧密码
'    CNEWPAS PCHAR   新密码
'返回:(OK或错误信息)
'===============================================================================================================



'11.前端初始化
Private Declare Function QDINIT Lib "CDGK_GRZH.dll" (ByVal str机构编号 As String) As String
'===============================================================================================================
'原型:FUNCTION QDINIT(AYBJGBH:STRING):PCHAR;
'功能:对前端进行初始化操作，以便前端个人帐户消费流水号与中心保持一致。
'入口参数:AYBJGBH PCHAR   医保机构编号
'返回:(OK或错误信息)
'===============================================================================================================


'12.下载交易记录
Private Declare Function DOWNJYJL Lib "CDGK_GRZH.dll" (ByVal str机构编号 As String) As String
'===============================================================================================================
'原型:FUNCTION DOWNJYJL(AYBJGBH:PCHAR):PCHAR
'功能:当本地医保数据库破坏后，从中心下载本定点前端所有还未审核结算的消费记录。
'入口参数:AYBJGBH PCHAR   医保机构编号
'返回:(OK或错误信息)
'===============================================================================================================


'*****************************************************************************************************************************************
'**住院部分
'*****************************************************************************************************************************************
'1.获得社保机构_旺苍编号和名称列表
Private Declare Function GetSBJGLB1 Lib "CDGK_HOS.dll" Alias "GETSBJGLB" () As String
'===============================================================================================================
'原型:FUNCTION GETSBJGLB:PCHAR
'功能: 获得社保机构_旺苍编号和名称列表
'入口参数: 无
'出口参数: 无
'返回: A社保机构编号+列间隔符+A社保机构名称+行间隔符+B社保机构编号+列间隔符+B社保机构名称+……
'===============================================================================================================



'1.入院登记
Private Declare Function RYDJ Lib "CDGK_HOS.dll" (ByVal str住院号 As String, ByVal str个人资料 As String, ByVal str机构编号 As String, ByVal str操作员姓名 As String) As String
'===============================================================================================================
'原型:FUNCTION RYDJ(AZYH,;ARYZL,ABXJGBH,ACZYXM:PCHAR):PCHAR;
'功能: 在社保机构医保数据库和医院本地医保数据库中对住院的医保病人进行登记。
'入口参数:AZYH    PCHAR   住院号
'    ARYZL   PCHAR   参保人员的入院资料
'    ABXJGBH PCHAR   参保人员所在的社保机构编号
'    ACZYXM  PCHAR   操作员姓名
'出口参数: 无
'说明:
'   参保人员的入院资料:个人ID||社保编号||姓名||性别||出生日期（格式：YYYY-MM-DD）||参加工作年月||退休年月||职务级别||职称级别||人员分类||异地居住标志||单位ID||单位名称||年龄||年月||医保证号||住院性质||基本医疗标志||补充医疗标志||公务员标志||基本医疗待遇状态||补充医疗待遇状态||公务员待遇状态||年内住院次数||年内已报销金额||缴费年限||提取时间||住院记录号||(以上参数实际为提取到的参保人员基础资料)||入院日期（格式：YYYY-MM-DD）||入院诊断||入院指征||病区||床号||科室
'返回:返回"OK或错识信息"
'===============================================================================================================

'2.取消住院
Private Declare Function ZYQX Lib "CDGK_HOS.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION ZYQX(AZYH:PCHAR):PCHAR;
'功能: 在没有正式结算前，在社保机构医保数据库和医院本地医保数据库中删除医保病人住院记录。
'入口参数: AZYH    PCHAR   住院号
'出口参数: 无
'返回:返回标志
'===============================================================================================================


'3.获得一个新的处方记录号，以保证处方的唯一性。
Private Declare Function GETNEWCFID Lib "CDGK_HOS.dll" () As String
'===============================================================================================================
'原型:FUNCTION GETNEWCFID:PCHAR
'功能: 获得一个新的处方记录号，以保证处方的唯一性。
'入口参数:
'出口参数: 无
'返回:返回标志:OK(或错误信息)@$处方记录号
'===============================================================================================================




'4.增加一个处方单据
Private Declare Function AddCFJL Lib "CDGK_HOS.dll" _
    Alias "ADDCFJL" (ByVal str住院号 As String, ByVal str处方号 As String, ByVal str序号 As String, ByVal str处方日期 As _
    String, ByVal str医生 As String, ByVal str科室 As String, ByVal str药品 As String, ByVal str数量 As _
    String, ByVal str单价 As String) As String
'===============================================================================================================
'原型:function ADDCFJL(AZYH,ACFID,ACFMXID,ACFRQ,AYS,AKS,AYPBH,ASL,ADJ:PCHAR):PCHAR
'功能:增加一个处方单据,必须保证ACFID，ACFMXID唯一
'入口参数:
'      AZYH    PCHAR   住院号
'    ACFID   PCHAR   处方单号(在整个数据库中保证唯一)
'    ACFMXID PCHAR   明细序号(在一个处方中保证唯一)
'    ACFRQ   PCHAR   处方日期（YYYY-MM-DD）
'    AYS PCHAR   医生
'    AKS PCHAR   科室
'    AYPBH   PCHAR   药品编号(社保药品编号)
'    ASL PCHAR   数量(可以为负数)
'    ADJ PCHAR   单价

'出口参数: 无
'返回:''OK'@$自费比例@$自费金额
'===============================================================================================================

'5.单条处方传输
Private Declare Function CFCS Lib "CDGK_HOS.dll" (ByVal str住院号 As String, ByVal str处方记录号 As String) As String
'===============================================================================================================
'原型:FUNCTION CFCS(AZYH:PCHAR;ACFID:PCHAR):PCHAR
'功能:将社保病人每天的处方情况向社保局中心数据库传输（同一个处方可以多次重复传输，后一次传输的数据将覆盖前一次传输的数据）。
'入口参数:
'       AZYH    PCHAR   住院号
'       ACFID   PCHAR   处方记录号（通过调用ADDCFJL返回的值）
'出口参数: 无
'返回:'OK或错误信息
'===============================================================================================================



'6.增加处方明细
Private Declare Function AddCFMX Lib "CDGK_HOS.dll" (ByVal str处方记录号 As String, ByVal str医保编码 As String, ByVal str数量 As String, ByVal str单价 As String) As String
'===============================================================================================================
'原型:FUNCTION ADDCFMX(ACFID,AYPBH,ASL,ADJ:PCHAR):PCHAR;
'功能:增加一个处方明细。
'入口参数:
'    ACFID   PCHAR   处方记录号
'    AYPBH   PCHAR   药品编号(社保药品编号)
'    ASL PCHAR   数量(可以为负数)
'    ADJ PCHAR   单价
'出口参数: 无
'返回:OK@$处方明细记录号@$自费比例@$自费金额
'===============================================================================================================



'7.出院结算
Private Declare Function CYJS Lib "CDGK_HOS.dll" (ByVal str住院号 As String, ByVal str操作员姓名 As String, ByVal lng预结标志 As Long, ByVal str治疗效果 As String, ByVal str出院诊断 As String, ByVal str出院日期 As String) As String
'===============================================================================================================
'原型:FNCTION CYJS(AZYH:PCHAR; ISPREV:INTEGER;ZLXG,CYZD,CYRQ:PCHAR):PCHAR
'功能:住院参保病人出院或住院中预结算,当为预结算时, ZLXG,CYZD,CYRQ三个参数不须要传输
'入口参数:
'    AZYH    PCHAR   住院号
'    ISPREV  PCHAR   预结算标志（'0'－表示预结算）
'    ZLXG    PCHAR   治疗效果
'    CYZD    PCHAR   出院诊断1||出院诊断2||出院诊断3||出院诊断4
'    CYRQ    PCHAR   出院日期（YYYY-MM-DD）
'出口参数: 无
'返回:OK@$住院费用结算结果@$报销分段明细
'   说明:
'       住院费用结算结果:基本医疗待遇状态||起付金额||基本封顶金额||基本报销比例||年内已报销金额||基本报销金额||补充报销金额||公务员报销金额||补充医疗待遇状态||公务员待遇状态||补充报销比例||公务员报销比例||本次住院费用||甲类费用||甲类药品费||甲类诊疗费||甲类服务费||乙类费用||乙类药品费||乙类诊疗费||乙类手术费||乙类自付||丙类费用||丙类药品费||丙类诊疗费||丙类服务费||报销合计||个人支付
'       报销分段明细(多条):险种||名称||段起始金额||段终止金额||本段基数||本段报销比例||本段报销金额||本段自付金额@$.......
'===============================================================================================================


'8.取消出院结算
Private Declare Function CYJSQX Lib "CDGK_HOS.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION CYJSQX(AZYH:PCHAR):PCHAR
'功能:取消参保病人出院结算
'入口参数:
'    AZYH    PCHAR   住院号
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================

'9.根据住院号得到住院记录号

Private Declare Function GETZYIDBYZYBH Lib "CDGK_HOS.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION GETZYIDBYZYBH(AZYH:PCHAR):PCHAR
'功能:根据住院号得到住院记录号。
'入口参数:
'    AZYH    PCHAR   住院号
'出口参数: 无
'返回:'OK@$住院记录号
'===============================================================================================================

'10.打印出院结算报表函数
Private Declare Function JSReport Lib "CDGK_HOS.dll" (ByVal str开始住院号 As String, ByVal str结束住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION JSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL
'功能: 打印结算报表，此报表比较耗用资源，最好由HIS系统来打印,格式由我公司提供。
'入口参数:
'    ASTARTZYH   PCHAR   打印开始住院号
'    AENDZYH PCHAR      打印结束住院号
'   注意:
'    1 ?二个住院号之间所有的住院记录必须为同一个社保局?
'    2、当只打印一个住院号的报表时，二个参数值一样。
'出口参数: 无
'返回:无须注意返回值
'===============================================================================================================


'11.住院病人跨月重提人员基本资料
Private Declare Function CWJSREPORT Lib "CDGK_HOS.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION CWJSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL;
'功能:住院病人跨月重提人员基本资料
'入口参数:
'   AZYH    PCHAR   住院号
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================

'12 提取基础资料_旺苍

Private Declare Function GETJCXX Lib "CDGK_HOS.dll" (ByVal str机构编码 As String, ByVal str下载标志 As String) As String
'===============================================================================================================
'原型:FUNCTION GETJCXX(SBXJGBH:PCHAR;DOWNALL:INTEGER):PCHAR
'功能:向指定的社保机构提取基础资料。
'入口参数:
'    SBXJGBH PCHAR   保险机构编号
'    DOWNALL PCHAR   当值为0时表示下载本地医保数据库中没有的基础资料，为其他时表示全部重新下载
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================

'20.住院病人跨月重提人员基本资料
Private Declare Function GETNEWRYZL Lib "CDGK_HOS.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION GETNEWRYZL(AZYH:PCHAR):PCHAR;STDCALL;
'功能:住院病人跨月重提人员基本资料。
'入口参数:
'    AZYH    PCHAR   住院号
'返回:OK@$错误信息
'===============================================================================================================

'21.增加一条待启用的药品诊疗项目
Private Declare Function ADDSINYP Lib "CDGK_HOS.dll" (ByVal str保险机构编号 As String, ByVal str明细数据 As String, ByVal str操作员姓名 As String, ByVal str上传日期 As String) As String
'===============================================================================================================
'原型:FUNCTION ADDSINYP(SBXJGBH,YPMXDATA,CZYXM,SCDATE:PCHAR):PCHAR;STDCALL;
'功能:增加一条待启用的药品诊疗项目。
'入口参数:
'    SBXJGBH    PCHAR   保险机构编号
'    YPMXDATA   PCHAR   处方明细数据
'    CZYXM      PCHAR   操作员姓名
'    SCDATE     PCHAR   上传日期
'返回:OK@$错误信息
'===============================================================================================================

'22.提取一条可启用的药品诊疗项目
Private Declare Function DOWNSINYP Lib "CDGK_HOS.dll" (ByVal str保险机构编号 As String, ByVal str药品编号 As String) As String
'===============================================================================================================
'原型:FUNCTION GETNEWRYZL(AZYH:PCHAR):PCHAR;STDCALL;
'功能:提取一条可启用的药品诊疗项目。
'入口参数:
'    SBXJGBH    PCHAR   保险机构编号
'    CYPBH      PCHAR   药品编号
'返回:'OK'+行间隔符+药品编号+列间隔符+药品启用标志+行间隔符+药品费用项目+行间隔符+费用类别@$错误信息
'===============================================================================================================


Public Function 医保初始化_广元旺苍() As Boolean
    Dim strReg As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    If mblnInit Then
        医保初始化_广元旺苍 = True
        Exit Function
    End If
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_广元旺苍.模拟数据 = True
    Else
        InitInfor_广元旺苍.模拟数据 = False
    End If
    
   Call GetRegInFor(g公共模块, "医保", "社保机构代码", strReg)
   
   InitInfor_广元旺苍.机构编码 = strReg
   g病人身份_广元旺苍.机构编码 = strReg
   
   If strReg = "" Then
        MsgBox "你未设置默认的社保机构编码，请检查参数设置!"
        Exit Function
   End If
   
    '取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & TYPE_广元旺苍
    Call OpenRecordset(rsTemp, "读取医院编码")
    InitInfor_广元旺苍.医院编码 = Nvl(rsTemp!医院编码)
    
    
    '中间库连接
    gstrSQL = "select 参数名,参数值 from 保险参数 where  险类=" & TYPE_广元旺苍
    Call OpenRecordset(rsTemp, "渝北医保")
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "操作员直下个人帐户"
                 g病人身份_广元旺苍.直输金额 = Nvl(rsTemp("参数值"), 0) = 1
            Case "明细时实上传"
                InitInfor_广元旺苍.明细时实上传 = IIf(Nvl(rsTemp!参数值, 1) = 1, 1, 0)
            Case "比较结算数据"
                 InitInfor_广元旺苍.数据不等不可结算 = IIf(Nvl(rsTemp!参数值, 1) = 1, 1, 0)
            Case "适用地区"
                 InitInfor_广元旺苍.适用地区 = Nvl(rsTemp!参数值, 0)
        End Select
        rsTemp.MoveNext
    Loop

    
    Set gcnOracle_广元旺苍 = New ADODB.Connection
    If OraDataOpen(gcnOracle_广元旺苍, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    
   '建立拔号连接
   If gbln已经初始 = False And gbln检查连接 Then
       If 建立拨号连接() = False Then Exit Function
   End If
   
   If gbln检查连接 Then
        '检查拔号连接
        If 业务请求_广元旺苍(检查拔号连接_旺苍, "", strOutput) = False Then
             Exit Function
        End If
    End If
    gbln已经初始 = True
    mblnInit = True
    医保初始化_广元旺苍 = True
End Function

Public Function 医保终止_广元旺苍() As Boolean
    Dim strOutput As String
    mblnInit = False
    If gcnOracle_广元旺苍.State = 1 Then
        gcnOracle_广元旺苍.Close
    End If
    If gbln检查连接 Then
         '建立拔号连接
        Call 业务请求_广元旺苍(断开拔号连接_旺苍, "", strOutput)
    End If
    Err = 0
    On Error Resume Next
    医保终止_广元旺苍 = True
End Function

Public Function 身份标识_广元旺苍(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '返回：空或信息串
    Err = 0
    On Error GoTo ErrHand:
    'If bytType = 1 Or bytType = 3 or  Then Exit Function
    
    身份标识_广元旺苍 = frmIdentify广元旺苍.GetPatient(bytType, lng病人ID)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_广元旺苍 = ""
End Function


Public Function 个人余额_广元旺苍(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID='" & lng病人ID & "' and 险类=" & TYPE_广元旺苍
    Call OpenRecordset(rsTemp, "读取个人帐户余额")
    
    If rsTemp.EOF Then
        个人余额_广元旺苍 = 0
    Else
        个人余额_广元旺苍 = rsTemp("帐户余额")
    End If
End Function
Public Function 门诊虚拟结算_广元旺苍(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim str明细 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    Dim strArr As Variant
    g病人身份_广元旺苍.费用总额 = 0
    str明细 = ""
    With rs明细
        Do While Not .EOF
            gstrSQL = "Select * From 医保支付项目 where 险类=" & TYPE_广元旺苍 & " and 中心=" & g病人身份_广元旺苍.社保中心 & " and 收费细目id=" & Nvl(!收费细目ID, 0)
            Call OpenRecordset(rsTemp, "确定医保支付项目")
            
            If Val(Nvl(rs明细("实收金额"), 0)) <> 0 Then
                If rsTemp.EOF Then
                    str明细 = str明细 & "@$" & ""
                Else
                    str明细 = str明细 & "@$" & Nvl(rsTemp!项目编码)
                End If
                '曾明春(2006-05-16):资阳机车厂单价和数量支持4位小数
                If InitInfor_广元旺苍.适用地区 = 0 Then
                    str明细 = str明细 & "||" & Nvl(!单价)
                    str明细 = str明细 & "||" & Nvl(!数量)
                Else
                    '曾明春(2005-07-06):如果单价精度超过2位小数，则数量传1，单价传实收金额。
                    '曾明春(2005-12-12):修改为使用列间隔符"||"连接
                    If Round(Nvl(!单价) * 100) = Nvl(!单价) * 100 Then
                       str明细 = str明细 & "||" & Nvl(!单价)
                       str明细 = str明细 & "||" & Nvl(!数量)
                    Else
                       str明细 = str明细 & "||" & Nvl(!实收金额)
                       str明细 = str明细 & "||" & "1"
                    End If
                End If
            End If
'            BJGBH  PCHAR   保险机构编号
'            '    CFMXDATA    PCHAR   处方明细数据    格式说明：处方1(医保药品编号+列间隔符+单价+列间隔符数量+)+行间隔符+
'            '    ……
'            '    处方N(医保药品编号+列间隔符+单价+列间隔符+数量
'            '    CPASSWORD   PCHAR   持卡人卡密码

            g病人身份_广元旺苍.费用总额 = g病人身份_广元旺苍.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    If str明细 <> "" Then
        str明细 = Mid(str明细, 3)
    End If
    g病人身份_广元旺苍.个人帐户支付 = 0
    If g病人身份_广元旺苍.直输金额 Then
        If g病人身份_广元旺苍.费用总额 > g病人身份_广元旺苍.帐户余额 Then
            str结算方式 = str结算方式 & "个人帐户;" & g病人身份_广元旺苍.帐户余额 & ";1"
        Else
            str结算方式 = str结算方式 & "个人帐户;" & g病人身份_广元旺苍.费用总额 & ";1"
        End If
    Else
         strInput = g病人身份_广元旺苍.机构编码
         strInput = strInput & vbTab & str明细
         strInput = strInput & vbTab & g病人身份_广元旺苍.密码
         If 业务请求_广元旺苍(门诊预处理_旺苍, strInput, strOutput) = False Then
            Exit Function
         End If
         strArr = Split(strOutput, "||")
         '个人帐户支付金额||自付金额
         
        str结算方式 = str结算方式 & "个人帐户;" & Val(strArr(0)) & ";0"
        g病人身份_广元旺苍.个人帐户支付 = Val(strArr(0))
    End If
    门诊虚拟结算_广元旺苍 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function 建立拨号连接() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Static str机构编号 As String
    Dim strInput As String, strOutput As String
    建立拨号连接 = False
    
    Err = 0: On Error GoTo ErrHand:
    If str机构编号 <> g病人身份_广元旺苍.机构编码 Then
        '检查网络是否正常连接
        If str机构编号 = "" Then
            '表求第一次远行,需断开
            If 业务请求_广元旺苍(建立拔号连接_旺苍, g病人身份_广元旺苍.机构编码, strOutput) = False Then
                Exit Function
            End If
        Else
            '表示至少有两次以上的操作,则需断开连接
            Call 业务请求_广元旺苍(断开拔号连接_旺苍, "", strOutput)
            If 业务请求_广元旺苍(建立拔号连接_旺苍, g病人身份_广元旺苍.机构编码, strOutput) = False Then Exit Function
        End If
        If 业务请求_广元旺苍(检查拔号连接_旺苍, "", strOutput) = False Then Exit Function
    Else
        If 业务请求_广元旺苍(检查拔号连接_旺苍, "", strOutput) = False Then
            '需重新建立拨号连接
            If 业务请求_广元旺苍(建立拔号连接_旺苍, g病人身份_广元旺苍.机构编码, strOutput) = False Then
                Exit Function
            End If
        End If
    End If
    str机构编号 = g病人身份_广元旺苍.机构编码
    建立拨号连接 = True
    Exit Function
ErrHand:
        If ErrCenter = 1 Then Resume
End Function
Public Function 门诊结算_广元旺苍(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    Dim strInput As String, strOutput As String
    Dim lng病人ID  As Long
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strArr As Variant
'    If 建立拨号连接() = False Then Exit Function
'
    On Error GoTo errHandle
    
    Call DebugTool("进入门诊结算")
    
    gstrSQL = "" & _
        "   Select a.*,a.*,a.付数*a.数次 as 数量,a.实收金额/(nvl(a.付数,1)*nvl(a.数次,1)) as 单价 " & _
        "   From 病人费用记录 a " & _
        "   Where nvl(实收金额,0)<>0 and  结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
        
    Call OpenRecordset(rs明细, "获取明细记录")
    
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If

    lng病人ID = rs明细("病人ID")
    
    If g病人身份_广元旺苍.病人ID <> lng病人ID Then
        MsgBox "该病人还没有经过身份验证，不能进行医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    If 业务请求_广元旺苍(初始化_旺苍, g病人身份_广元旺苍.机构编码, strOutput) = False Then Exit Function
    
    g结算数据.结帐ID = lng结帐ID
    g结算数据.结算标志 = 0
    '写入明细
    If 门诊明细写入(rs明细, False) = False Then Exit Function
    
    If g病人身份_广元旺苍.直输金额 = False Then
'        '显示其结处方式
        Call 结算方式更正
        DebugTool "结算已经显示完成"
    End If
    DebugTool "开始保存数据"
    
    
    Dim dbl个人帐户 As Double
    dbl个人帐户 = cur个人帐户
    If dbl个人帐户 <> g结算数据.个人帐户支付金额 Then
        If g病人身份_广元旺苍.直输金额 Then
            '更新个人帐户支付
            '入:YBJGBH  PCHAR   保险机构编号
            '    XFJE    PCHAR   消费金额(保证为小数，并且保留二位小数)
            '    CPASSWORD   PCHAR   持卡人卡密码
            '    CCZYXM  PCHAR   操作员姓名
            '返回:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
            strInput = g病人身份_广元旺苍.机构编码
            strInput = strInput & vbTab & Format(dbl个人帐户, "###0.00;-###0.00;0.00;0.00")
            strInput = strInput & vbTab & g病人身份_广元旺苍.密码
            strInput = strInput & vbTab & gstrUserName
            If 业务请求_广元旺苍(个人帐户消费_金额_旺苍, strInput, strOutput) = False Then Exit Function
            If strOutput = "" Then Exit Function
            strArr = Split(strOutput, "||")
            
            With g结算数据
                .卡号 = strArr(0)
                .姓名 = strArr(1)
                .消费前帐户余额 = Val(strArr(2))
                .个人帐户支付金额 = Val(strArr(3))
                .自费金额 = Val(strArr(4))
                .消费后帐户余额 = Val(strArr(5))
                .交易时间 = strArr(6)
                .前端单据号 = strArr(7)
                .中心单据号 = strArr(8)
                .处方号 = strArr(9)
                .操作员姓名 = strArr(10)
                .前端名称 = strArr(11)
                If InitInfor_广元旺苍.适用地区 = 0 Then
                   .人员分类 = strArr(12)
                End If
            End With
        End If
    End If
       
    '填写结算表
    Call DebugTool("填写结算记录")
    

    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(消费前帐户余额),累计统筹报销_IN(消费后帐户余额),住院次数_IN(无),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(自费金额),
    '   进入统筹金额_IN(无),统筹报销金额_IN(无),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(中心单据号),主页ID_IN(无),中途结帐_IN,备注_IN(前端单据号|处方号|操作员姓名|前端名称)
    '其中中途结帐ID保存人员分类
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    
    If InitInfor_广元旺苍.适用地区 = 0 Then
        gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_广元旺苍 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
                 "NULL,NULL," & g结算数据.消费前帐户余额 & "," & g结算数据.消费后帐户余额 & ",null,0,0,0," & _
                g病人身份_广元旺苍.费用总额 & ",0," & g结算数据.自费金额 & "," & _
              "0,0,0,0," & g结算数据.个人帐户支付金额 & ",'" & _
                g结算数据.中心单据号 & " ',NULL,NULL,'" & g结算数据.人员分类 & "|" & g结算数据.前端单据号 & "|" & g结算数据.处方号 & "|" & g结算数据.操作员姓名 & "|" & g结算数据.前端单据号 & "')"
    Else
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_广元旺苍 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL," & g结算数据.消费前帐户余额 & "," & g结算数据.消费后帐户余额 & ",null,0,0,0," & _
            g病人身份_广元旺苍.费用总额 & ",0," & g结算数据.自费金额 & "," & _
          "0,0,0,0," & g结算数据.个人帐户支付金额 & ",'" & _
            g结算数据.中心单据号 & " ',NULL,NULL,'" & g结算数据.前端单据号 & "|" & g结算数据.处方号 & "|" & g结算数据.操作员姓名 & "|" & g结算数据.前端单据号 & "')"
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '---------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------
    门诊结算_广元旺苍 = True
    Exit Function

'Err反结算:
'
''入口参数:YBJGBH  PCHAR   保险机构编号
''        cZXDJH  PCHAR   中心单据号(消费时返回)
''        CPASSWORD   PCHAR   持卡人卡密码
''        CCZYXM  PCHAR   操作员姓名
'    strInput = g病人身份_广元旺苍.机构编码
'    strInput = strInput & vbTab & g结算数据.中心单据号
'    strInput = strInput & vbTab & g病人身份_广元旺苍.密码
'    strInput = strInput & vbTab & gstrUserName
'
'    If 业务请求_广元旺苍(消费冲正_旺苍, strInput, strOutPut) = False Then Exit Function
''返回:个人帐户消费信息(OK@$个人帐户消费信息)
''   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
'    If strOutPut = "" Then Exit Function
'     strArr = Split(strOutPut, "||")
'
'    With g结算数据
'        .卡号 = strArr(0)
'        .姓名 = strArr(1)
'        .消费前帐户余额 = Val(strArr(2))
'        .个人帐户支付金额 = Val(strArr(3))
'        .自费金额 = Val(strArr(4))
'        .消费后帐户余额 = Val(strArr(5))
'        .交易时间 = strArr(6)
'        .前端单据号 = strArr(7)
'        .中心单据号 = strArr(8)
'        .处方号 = strArr(9)
'        .操作员姓名 = strArr(10)
'        .前端名称 = strArr(11)
'    End With

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function 门诊结算冲销_广元旺苍(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    
    Dim intMouse As Integer
    Dim lng冲销ID  As Long
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    Dim strArr As Variant
    Dim lng病人id1 As Long
    门诊结算冲销_广元旺苍 = False
    
    '身份验证
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If 身份标识_广元旺苍(2, lng病人id1) = "" Then
        If lng病人id1 = 0 Then
            ShowMsgbox "你不是当前持卡人!"
            Screen.MousePointer = intMouse
            Exit Function
        End If
    End If
    Screen.MousePointer = intMouse
    If 业务请求_广元旺苍(初始化_旺苍, g病人身份_广元旺苍.机构编码, strOutput) = False Then Exit Function
    
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "重庆医保")
    lng冲销ID = rsTemp("结帐ID")
    
    
    
    gstrSQL = "Select * From 病人费用记录 " & _
        " Where 结帐ID=" & lng冲销ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
        
    Call OpenRecordset(rs明细, "获取冲销记录")
    g病人身份_广元旺苍.费用总额 = 0
    With rs明细
        Do While Not .EOF
                '写上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            g病人身份_广元旺苍.费用总额 = g病人身份_广元旺苍.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    
    '冲正:
    gstrSQL = "Select 支付顺序号,substr(备注,1,instr(备注,'|')-1) as 人员分类 from 保险结算记录 where 性质=1 and 记录id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取中心单据号"
    If rsTemp.EOF Then
        ShowMsgbox "不存在结算记录,不能冲销!"
        Exit Function
    End If
    
    '入口参数:YBJGBH  PCHAR   保险机构编号
    '        cZXDJH  PCHAR   中心单据号(消费时返回)
    '        CPASSWORD   PCHAR   持卡人卡密码
    '        CCZYXM  PCHAR   操作员姓名
    strInput = g病人身份_广元旺苍.机构编码
    strInput = strInput & vbTab & Nvl(rsTemp!支付顺序号)
    strInput = strInput & vbTab & g病人身份_广元旺苍.密码
    strInput = strInput & vbTab & gstrUserName
    
    If 业务请求_广元旺苍(消费冲正_旺苍, strInput, strOutput) = False Then Exit Function
    '返回:个人帐户消费信息(OK@$个人帐户消费信息)
    '   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
    If strOutput = "" Then Exit Function
     strArr = Split(strOutput, "||")
    
    With g结算数据
        .卡号 = strArr(0)
        .姓名 = strArr(1)
        .消费前帐户余额 = Val(strArr(2))
        .个人帐户支付金额 = Val(strArr(3))
        .自费金额 = Val(strArr(4))
        .消费后帐户余额 = Val(strArr(5))
        .交易时间 = strArr(6)
        .前端单据号 = strArr(7)
        .中心单据号 = strArr(8)
        .处方号 = strArr(9)
        .操作员姓名 = strArr(10)
        .前端名称 = strArr(11)
    End With
    门诊结算冲销_广元旺苍 = True
        
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(消费前帐户余额),累计统筹报销_IN(消费后帐户余额),住院次数_IN(无),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(自费金额),
    '   进入统筹金额_IN(无),统筹报销金额_IN(无),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(中心单据号),主页ID_IN(无),中途结帐_IN,备注_IN(前端单据号|处方号|操作员姓名|前端名称)
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    
    If InitInfor_广元旺苍.适用地区 = 0 Then
        gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_广元旺苍 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
                 "NULL,NULL," & -1 * g结算数据.消费前帐户余额 & "," & -1 * g结算数据.消费后帐户余额 & ",null,0,0,0," & _
               -1 * g病人身份_广元旺苍.费用总额 & ",0," & -1 * g结算数据.自费金额 & "," & _
              "0,0,0,0," & -1 * g结算数据.个人帐户支付金额 & ",'" & _
                g结算数据.中心单据号 & " ',NULL,NULL,'" & rsTemp!人员分类 & "|" & g结算数据.前端单据号 & "|" & g结算数据.处方号 & "|" & g结算数据.操作员姓名 & "|" & g结算数据.前端单据号 & "')"
    Else
        gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_广元旺苍 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL," & -1 * g结算数据.消费前帐户余额 & "," & -1 * g结算数据.消费后帐户余额 & ",null,0,0,0," & _
           -1 * g病人身份_广元旺苍.费用总额 & ",0," & -1 * g结算数据.自费金额 & "," & _
          "0,0,0,0," & -1 * g结算数据.个人帐户支付金额 & ",'" & _
            g结算数据.中心单据号 & " ',NULL,NULL,'" & g结算数据.前端单据号 & "|" & g结算数据.处方号 & "|" & g结算数据.操作员姓名 & "|" & g结算数据.前端单据号 & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '---------------------------------------------------------------------------------------------
    门诊结算冲销_广元旺苍 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Get个人资料(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    '    个人ID||社保编号||姓名||性别||出生日期（格式：YYYY-MM-DD）||参加工作年月||退休年月||职务级别||
    '   职称级别||人员分类||异地居住标志||单位ID||单位名称||年龄||年月||医保证号||住院性质||基本医疗标志||
    '   补充医疗标志||公务员标志||基本医疗待遇状态||补充医疗待遇状态||公务员待遇状态||年内住院次数||
    '   年内已报销金额||缴费年限||提取时间||住院记录号||入院日期（格式：YYYY-MM-DD）||
    '   入院诊断||入院指征||病区||床号||科室
    
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String
    
    gstrSQL = "" & _
        "   Select  to_char(a.入院日期,'yyyy-mm-dd') as 入院日期,a.入院病况,b.名称 as 科室,c.名称 as 病区,d.入院诊断,a.入院病床" & _
        "   From 病案主页 a,部门表 b,部门表 c, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.名称,'')) AS 入院诊断, max(DECODE(a.诊断次序,2,b.编码||'-'||b.名称,'')) AS 入院诊断1,max(DECODE(a.诊断次序,3,b.编码||'-'||b.名称,'')) AS 入院诊断2, max(DECODE(a.诊断次序,4,b.编码||'-'||b.名称,'')) AS 入院诊断3 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1  and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by 病人id,主页id) D" & _
        " Where  " & _
        "        A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & " and A.出院科室ID=b.id(+) and a.入院病区ID =c.id(+) " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) "
        
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取主页信息"
    
    With g病人身份_广元旺苍
        strInput = .str住院信息
        strInput = strInput & "||" & Nvl(rsTemp!入院日期, "")
        strInput = strInput & "||" & Nvl(rsTemp!入院诊断)
        strInput = strInput & "||" & Nvl(rsTemp!入院病况)     '入院指征,目前没有传
        strInput = strInput & "||" & Nvl(rsTemp!病区)
        strInput = strInput & "||" & Nvl(rsTemp!入院病床)
        strInput = strInput & "||" & Nvl(rsTemp!科室)
    End With
    Get个人资料 = strInput
End Function

Public Function 入院登记_广元旺苍(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
  '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, strInput As String
    Dim strArr
    Err = 0: On Error GoTo ErrHand:
    
    '获取住院号
    DebugTool "进入入院登记接口"
    
    
   If InitInfor_广元旺苍.机构编码 <> g病人身份_广元旺苍.机构编码 Then
        '建立拔号连接
        If gbln已经初始 = False And gbln检查连接 Then
             If 业务请求_广元旺苍(建立拔号连接_旺苍, g病人身份_广元旺苍.机构编码, strOutput) = False Then
                  Exit Function
             End If
        End If
        
        If gbln检查连接 Then
             '检查拔号连接
             If 业务请求_广元旺苍(检查拔号连接_旺苍, "", strOutput) = False Then
                  Exit Function
             End If
         End If
    End If
    
'    gstrSQL = "Select 医保住院号_ID.nextval  as 住院号  From dual "
'    OpenRecordset_广元旺苍 rsTemp, "获取住院号"
'
    '    AZYH    PCHAR   住院号
    '    ARYZL   PCHAR   参保人员的入院资料
    '    ABXJGBH PCHAR   参保人员所在的社保机构编号
    '    ACZYXM  PCHAR   操作员姓名

    'strInput = Rpad(lng主页ID, 4, "0") & lng病人ID
    '宋献平  2006年12月19日  处理住院次数为1和10时重复问题
    '考虑到在院病人，所以只处理住院次数大于9的病人
     strInput = Rpad(lng主页ID, 4, "0") & lng病人ID
     '宋献平  2006年12月19日  处理住院次数为1和10时住院号上传重复问题
     '由于当前编号规则已经使用，不能全新编规则，故住院号小于10时仍按原规则处理，以保证现有病人的数据正常
     '医保中心住院号的规则为不重复的数字，和年无关
     If lng主页ID > 9 Then strInput = "9" & Lpad(lng主页ID, 3, "0") & lng病人ID

    strInput = strInput & vbTab & Get个人资料(lng病人ID, lng主页ID)
    strInput = strInput & vbTab & g病人身份_广元旺苍.机构编码
    strInput = strInput & vbTab & gstrUserName
    

    If 业务请求_广元旺苍(入院登记_旺苍, strInput, strOutput) = False Then
        Exit Function
    End If
    
    '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_广元旺苍 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_广元旺苍 = False
End Function
Private Function Get交易代码(ByVal intType As 业务类型_广元旺苍, Optional bln读名称 As Boolean = False) As String
    '代码暂没用
    Select Case intType
        Case 获得社保机构_旺苍
            Get交易代码 = IIf(bln读名称, "获得社保机构", "01")
        Case 获得社保机构_住院_旺苍
            Get交易代码 = IIf(bln读名称, "获得社保机构_住院_旺苍", "27")
        
        Case 获得参保人员资料_旺苍
            Get交易代码 = IIf(bln读名称, "获得参保人员资料", "02")
        Case 获取帐户余额_旺苍
                Get交易代码 = IIf(bln读名称, "获取帐户余额", "03")
        Case 检查拔号连接_旺苍
            Get交易代码 = IIf(bln读名称, "检查拔号连接", "04")
        Case 建立拔号连接_旺苍
            Get交易代码 = IIf(bln读名称, "建立拔号连接", "05")
        Case 断开拔号连接_旺苍
            Get交易代码 = IIf(bln读名称, "断开拔号连接", "06")
        Case 个人帐户消费_旺苍
            Get交易代码 = IIf(bln读名称, "个人帐户消费", "07")
        Case 个人帐户消费_金额_旺苍
            Get交易代码 = IIf(bln读名称, "个人帐户消费_金额", "08")
        Case 消费冲正_旺苍
            Get交易代码 = IIf(bln读名称, "消费冲正", "09")
        Case 修改密码_旺苍
            Get交易代码 = IIf(bln读名称, "修改密码", "10")
        Case 初始化_旺苍
            Get交易代码 = IIf(bln读名称, "初始化", "11")
        Case 下载交易记录_旺苍
            Get交易代码 = IIf(bln读名称, "下载交易记录", "12")
        Case 获得人员资料_医保号_旺苍
            Get交易代码 = IIf(bln读名称, "获得人员资料_医保号_旺苍", "13")
        Case 获得人员资料_读卡_旺苍
            Get交易代码 = IIf(bln读名称, "获得人员资料_读卡_旺苍", "14")
        Case 入院登记_旺苍
            Get交易代码 = IIf(bln读名称, "入院登记_旺苍", "15")
        Case 取消入院登记_旺苍
            Get交易代码 = IIf(bln读名称, "取消入院登记_旺苍", "16")
        Case 获取处方记录号_旺苍
            Get交易代码 = IIf(bln读名称, "获取处方记录号_旺苍", "17")
        Case 增加处方单据_旺苍
            Get交易代码 = IIf(bln读名称, "增加处方单据_旺苍", "18")
        Case 单条处方传输_旺苍
            Get交易代码 = IIf(bln读名称, "单条处方传输_旺苍", "19")
        Case 增加处方明细_旺苍
            Get交易代码 = IIf(bln读名称, "增加处方明细_旺苍", "20")
        Case 出院结算_旺苍
            Get交易代码 = IIf(bln读名称, "出院结算_旺苍", "21")
        Case 取消出院结算_旺苍
            Get交易代码 = IIf(bln读名称, "取消出院结算_旺苍", "22")
        Case 根据住院号获取记录号_旺苍
            Get交易代码 = IIf(bln读名称, "根据住院号获取记录号_旺苍", "23")
        Case 打印结算报表_旺苍
            Get交易代码 = IIf(bln读名称, "打印结算报表_旺苍", "24")
        Case 住院病人跨月重提_旺苍
            Get交易代码 = IIf(bln读名称, "住院病人跨月重提_旺苍", "25")
        Case 提取基础资料_旺苍
            Get交易代码 = IIf(bln读名称, "提取基础资料_旺苍", "26")
        Case 门诊预处理_旺苍
            Get交易代码 = IIf(bln读名称, "门诊预处理_旺苍", "28")
        Case Else
            Get交易代码 = IIf(bln读名称, "错误的交易代码", "-1")
    End Select
End Function
Public Function 业务请求_广元旺苍(ByVal intType As 业务类型_广元旺苍, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInput As String, lngReturn As Long, strOutput As String, strReturn As String
    Dim strInValue(0 To 20) As String
    
    Dim str交易代码 As String
    Dim i As Integer
    Dim strArr
    
    str交易代码 = Get交易代码(intType, True)
    strInput = strInputString
    DebugTool "进入业务请求函数(业务类型代码为:" & intType & " 业务名称：" & str交易代码 & ")" & vbCrLf & "        输入参数为:" & strInputString
    
    业务请求_广元旺苍 = False
    If InitInfor_广元旺苍.模拟数据 Then
        '读取模拟数据
        Read模拟数据 intType, strInput, strOutPutstring
         业务请求_广元旺苍 = True
        Exit Function
    End If
    strArr = Split(strInputString, vbTab)
    For i = 0 To UBound(strArr)
        strInValue(i) = strArr(i)
    Next
    
    Err = 0
    On Error GoTo ErrHand:
    
    Select Case intType
        Case 获得社保机构_旺苍
            strOutput = GetSBJGLB()
            
            If strOutput = "" Then
                MsgBox "获取社保机构时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Replace(strOutput, "OK@$", "")
        Case 获得社保机构_住院_旺苍
            strOutput = GetSBJGLB1()
            
            If strOutput = "" Then
                MsgBox "获得社保机构_住院_旺苍时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Replace(strOutput, "OK@$", "")
        
        Case 获得参保人员资料_旺苍
            strOutput = GETKZL()
            If strOutput = "" Then
                MsgBox "获得参保人员资料_旺苍时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case 获得人员资料_医保号_旺苍
            strOutput = GETRYJBZL(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "获得参保人员资料_旺苍时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        
        Case 获得人员资料_读卡_旺苍
            strOutput = GETRYJBZL_BYYBK(strInValue(0))
            If strOutput = "" Then
                MsgBox "获得参保人员资料_旺苍时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        
        Case 获取帐户余额_旺苍
            strOutput = GETZHYE(strInValue(0), strInValue(1))
            ''OK'+行间隔符+个人帐户余额
            If strOutput = "" Then
                MsgBox "获取帐户余额_时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        
        Case 检查拔号连接_旺苍
            strOutput = CheckCon()
            If strOutput = "" Then
                MsgBox "检查拔号连接时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 建立拔号连接_旺苍
            strOutput = RasDial(strInValue(0))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 断开拔号连接_旺苍
            strOutput = DisDial()
            strOutput = Split(strOutput, Chr(0))(0)
            strOutput = ""
        Case 个人帐户消费_旺苍
            strOutput = GRZHXF_CF(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
            strOutput = strArr(1)
            If InitInfor_广元旺苍.适用地区 = 0 Then
               For i = 2 To UBound(strArr)
                   strOutput = strOutput & "@$" & strArr(i)
               Next
            End If
        Case 门诊预处理_旺苍
            strOutput = GRZHXF_CFPRE(strInValue(0), strInValue(1), strInValue(2))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            'OK@$个人帐户支付金额||自付金额
            strOutput = strArr(1)
        
        Case 个人帐户消费_金额_旺苍
            strOutput = GRZHXF_JE(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
            strOutput = strArr(1)
        Case 消费冲正_旺苍
            strOutput = XFCZ(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
            strOutput = strArr(1)
        Case 修改密码_旺苍
            strOutput = CHANGPASSWORD(strInValue(0), strInValue(1), strInValue(2))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If Left(strArr(0), 2) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
           
        Case 初始化_旺苍
        
                '杨芬20050201:取消该函数
                '            strOutPut = QDINIT(strInValue(0))
                '            strOutPut = Split(strOutPut, Chr(0))(0)
                '            strArr = Split(strOutPut, "@$")
                '            If Left(strArr(0), 2) <> "OK" Then
                '                MsgBox strArr(0), vbInformation, gstrSysName
                '                Exit Function
                '            End If
            strOutput = ""
        Case 下载交易记录_旺苍
        
            strOutput = DOWNJYJL(strInValue(0))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If Left(strArr(0), 2) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            
            strOutput = ""
        Case 入院登记_旺苍
            strOutput = RYDJ(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            If strOutput = "" Then
                MsgBox "入院登记时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 取消入院登记_旺苍
            strOutput = ZYQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "取消入院登记时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 获取处方记录号_旺苍
            strOutput = GETNEWCFID()
            If strOutput = "" Then
                MsgBox "获取处方记录号时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""

        Case 增加处方单据_旺苍
            '入参:
            '    AZYH    PCHAR   住院号
            '    ACFID   PCHAR   处方单号(在整个数据库中保证唯一)
            '    ACFMXID PCHAR   明细序号(在一个处方中保证唯一)
            '    ACFRQ   PCHAR   处方日期（YYYY-MM-DD）
            '    AYS PCHAR   医生
            '    AKS PCHAR   科室
            '    AYPBH   PCHAR   药品编号(社保药品编号)
            '    ASL PCHAR   数量(可以为负数)
            '    ADJ PCHAR   单价

            strOutput = AddCFJL(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4), strInValue(5), strInValue(6), strInValue(7), strInValue(8))
            If strOutput = "" Then
                MsgBox "增加处方单据时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            
            strOutput = ""
            For i = 1 To UBound(strArr)
                strOutput = "||" & strArr(i)
            Next
            If strOutput <> "" Then
                strOutput = Mid(strOutput, 3)
            End If
        Case 单条处方传输_旺苍
            strOutput = CFCS(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "单条处方传输时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 增加处方明细_旺苍
            strOutput = AddCFMX(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            If strOutput = "" Then
                MsgBox "增加处方明细时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
            For i = 1 To UBound(strArr)
                strOutput = "||" & strArr(i)
            Next
            If strOutput <> "" Then
                strOutput = Mid(strOutput, 3)
            End If
        Case 出院结算_旺苍
            strOutput = CYJS(strInValue(0), strInValue(1), Val(strInValue(2)), strInValue(3), strInValue(4), strInValue(5))
            If strOutput = "" Then
                MsgBox "出院结算时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Mid(strOutput, 5)
        
        Case 取消出院结算_旺苍
            strOutput = CYJSQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "取消出院结算时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 根据住院号获取记录号_旺苍
            
            strOutput = GETZYIDBYZYBH(strInValue(0))
            If strOutput = "" Then
                MsgBox "根据住院号获取记录号时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case 打印结算报表_旺苍
            strOutput = JSReport(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "打印结算报表时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 住院病人跨月重提_旺苍
            strOutput = GETNEWRYZL(strInValue(0))
            If strOutput = "" Then
                MsgBox "跨月重提人员基本资料时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""

        Case 提取基础资料_旺苍
             strOutput = GETJCXX(strInValue(1), strInValue(2))
            If strOutput = "" Then
                MsgBox "提取基础资料时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 申报项目_资阳
            strOutput = ADDSINYP(strInValue(1), strInValue(2), strInValue(3), strInValue(4))
            If strOutput = "" Then
                MsgBox "提取基础资料时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 提取项目_资阳
             strOutput = DOWNSINYP(strInValue(1), strInValue(2))
            If strOutput = "" Then
                MsgBox "提取基础资料时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
    End Select
    strOutPutstring = strOutput
    业务请求_广元旺苍 = True
    DebugTool "    输出参数为:" & strOutPutstring
     Exit Function
    
ErrHand:
    DebugTool "    输出参数为:" & strOutPutstring
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_广元旺苍(lng病人ID As Long, lng主页ID As Long) As Boolean
  '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
            
    '刘兴宏:20040923增加的
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
     Err = 0
    On Error GoTo ErrHand
    
    DebugTool "进入出院登撤消接口"
    
    入院登记撤销_广元旺苍 = False
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "存在未结费用，不能撤消入院登记"
        Exit Function
    End If
    'strInput = Rpad(lng主页ID, 4, "0") & lng病人ID
    strInput = Rpad(lng主页ID, 4, "0") & lng病人ID
     If lng主页ID > 9 Then strInput = "9" & Lpad(lng主页ID, 3, "0") & lng病人ID '宋献平  2006年12月19日  处理住院次数为1和10时重复问题
    
    If 业务请求_广元旺苍(取消入院登记_旺苍, strInput, strOutput) = False Then Exit Function

    '更新医保帐户
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_广元旺苍 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    DebugTool "取消成功"
    入院登记撤销_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_广元旺苍(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    
    Err = 0:    On Error GoTo ErrHand:
    出院登记_广元旺苍 = False
    
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "当前病人不存在未结费用，请在入院撤消即可"
        Exit Function
    End If
 
    '改变当前状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_广元旺苍 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function 出院登记撤销_广元旺苍(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
  '出院登记撤消
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    Dim strArr As Variant
    
    出院登记撤销_广元旺苍 = False
    
    Err = 0: On Error GoTo ErrHand:
     
     If Not 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "该病人已经出院结算了,不能再取消出院!"
        Exit Function
     End If
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_广元旺苍 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    出院登记撤销_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_广元旺苍(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
  '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
 '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    
    Dim rsTemp As New ADODB.Recordset, strInput As String, strOutput As String
    
    Dim lng主页ID As Long
    Dim dbl费用总额 As Double
    Dim strArr As Variant, strTmpArr As Variant
    Dim str结算方式  As String, str住院号 As String
    Dim obj结算 As 结算数据
    住院结算_广元旺苍 = False
        
 
    Err = 0: On Error GoTo ErrHand:
    Call DebugTool("进入住院结算")
    
    
    If g病人身份_广元旺苍.病人ID <> lng病人ID Then
        MsgBox "该病人没有完成医保的预结算操作，不能进行结算。", vbInformation, gstrSysName
        Exit Function
    End If
        
    gstrSQL = "Select 当前状态 From 保险帐户  where 病人ID=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断当前的住院状态!"
    
    If Nvl(rsTemp!当前状态, 0) = 1 Then
        ShowMsgbox "当前病人还处于在院状态,请出院后再结算!"
        Exit Function
    End If
    
    
    With g结算数据
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & lng病人ID
        Call OpenRecordset(rsTemp, "虚拟结算")
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        lng主页ID = rsTemp("主页ID")
    End With
    
    str住院号 = Rpad(lng主页ID, 4, "0") & lng病人ID
   If lng主页ID > 9 Then str住院号 = "9" & Lpad(lng主页ID, 3, "0") & lng病人ID
    
    
    gstrSQL = " " & _
          " Select sum(nvl(结帐金额,0)) as 实收金额 " & _
          " From 病人费用记录 " & _
          " Where 记录状态<>0 and 结帐ID=" & lng结帐ID & " and  Nvl(附加标志,0)<>9"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取总费用"
    
    dbl费用总额 = Nvl(rsTemp!实收金额, 0)
    If dbl费用总额 <> g病人身份_广元旺苍.费用总额 Then
        ShowMsgbox "虚拟结算数据的费用总额与本次结算的费用总额不等，可能又有处方记帐单上传了!"
        Exit Function
    End If
    
    gstrSQL = "" & _
        "   Select C.住院号,C.当前病区id,A.入院病床 ,c.住院号,to_char(A.确诊日期,'yyyyMMdd') as 确诊日期,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyyMMdd') 入院经办时间," & _
        "           to_char(A.入院日期,'yyyyMMdd') 入院日期, A.出院方式,to_char(a.出院日期,'yyyy-mm-dd') as 出院日期 ,a.出院病床,H.名称 as 出院科室," & _
        "           g.治疗情况,G.出院诊断1,G.出院诊断2,G.出院诊断3,G.出院诊断4" & _
        " From 病案主页 A,部门表 B,病人信息 C,部门表 H, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,出院情况,'')) as 治疗情况,max(DECODE(a.诊断次序,1,b.编码||'-'||b.名称,'')) AS 出院诊断1,max(DECODE(a.诊断次序,2,b.编码||'-'||b.名称,'')) AS 出院诊断2,max(DECODE(a.诊断次序,3,b.编码||'-'||b.名称,'')) AS 出院诊断3,max(DECODE(a.诊断次序,4,b.编码||'-'||b.名称,'')) AS 出院诊断4 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 = 3 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by 病人id,主页id)   G" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID and A.出院科室ID=H.id(+) " & _
        "       and A.主页id=G.主页id(+) and a.病人id=G.病人id(+) " & _
        ""
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取出院信息"
    
    '再次结算
    strInput = str住院号
    strInput = strInput & vbTab & gstrUserName
    strInput = strInput & vbTab & "1"
    strInput = strInput & vbTab & Get治渝情况_旺苍(lng病人ID, lng主页ID)
    strInput = strInput & vbTab & Nvl(rsTemp!出院诊断1) & "||" & Nvl(rsTemp!出院诊断2) & "||" & Nvl(rsTemp!出院诊断3) & "||" & Nvl(rsTemp!出院诊断4)
    strInput = strInput & vbTab & Nvl(rsTemp!出院日期)
    If 业务请求_广元旺苍(出院结算_旺苍, strInput, strOutput) = False Then Exit Function
    
    strArr = Split(strOutput, "@$")
    strTmpArr = Split(strArr(0), "||")
    With obj结算
        .基本报销金额 = Val(strTmpArr(5))
        .补充报销金额 = Val(strTmpArr(6))
        .公务员报销金额 = Val(strTmpArr(7))
        If InitInfor_广元旺苍.适用地区 = 0 Then
           .个人帐户支付金额 = Val(strTmpArr(28))
           .人员分类 = Trim(strTmpArr(29))
        End If
    End With
    
    gcnOracle_广元旺苍.BeginTrans

    If InsertInto医保结算记录(strArr, lng结帐ID) = False Then Exit Function
    
    
    '填写结算表
    Call DebugTool("填写结算记录")
    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(),累计统筹报销_IN(年内已报销金额),住院次数_IN(主页ID),起付线(起付金额),封顶线_IN(基本封顶金额),实际起付线_IN(基本报销比例),
    '   发生费用金额_IN(费用总额),全自付金额_IN(补充报销比例),首先自付金额_IN(公务员报销比例),
    '   进入统筹金额_IN(基本报销金额),统筹报销金额_IN(补充报销金额),大病自付金额_IN(公务员报销金额),超限自付金额_IN(无),个人帐户支付_IN(),"
    '   支付顺序号_IN(住院号),主页ID_IN(主页ID),中途结帐_IN,备注_IN(中心)
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
   
   '基本医疗待遇状态||起付金额||基本封顶金额||基本报销比例||年内已报销金额||基本报销金额||补充报销金额||公务员报销金额||
   '补充医疗待遇状态||公务员待遇状态||补充报销比例||公务员报销比例||本次住院费用||甲类费用||甲类药品费||甲类诊疗费||甲类服务费||乙类费用||乙类药品费||乙类诊疗费||乙类手术费||乙类自付||丙类费用||丙类药品费||丙类诊疗费||丙类服务费||报销合计||个人支付
    If InitInfor_广元旺苍.适用地区 = 0 Then
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_广元旺苍 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
                 "NULL,NULL,NULL," & Val(strTmpArr(4)) & "," & lng主页ID & "," & Val(strTmpArr(1)) & "," & Val(strTmpArr(2)) & "," & Val(strTmpArr(3)) & "," & _
                dbl费用总额 & "," & Val(strTmpArr(6)) & "," & Val(strTmpArr(7)) & "," & _
                obj结算.基本报销金额 & "," & obj结算.补充报销金额 & "," & obj结算.公务员报销金额 & ",0," & obj结算.个人帐户支付金额 & ",'" & _
                str住院号 & "'," & lng主页ID & ",NULL,'" & obj结算.人员分类 & "|" & g病人身份_广元旺苍.社保中心 & "')"
    Else
            gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_广元旺苍 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
                 "NULL,NULL,NULL," & Val(strTmpArr(4)) & "," & lng主页ID & "," & Val(strTmpArr(1)) & "," & Val(strTmpArr(2)) & "," & Val(strTmpArr(3)) & "," & _
                dbl费用总额 & "," & Val(strTmpArr(6)) & "," & Val(strTmpArr(7)) & "," & _
                obj结算.基本报销金额 & "," & obj结算.补充报销金额 & "," & obj结算.公务员报销金额 & ",0,0,'" & _
                str住院号 & "'," & lng主页ID & ",NULL,'" & g病人身份_广元旺苍.社保中心 & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '---------------------------------------------------------------------------------------------
     gcnOracle_广元旺苍.CommitTrans

      
    住院结算_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function 住院结算冲销_广元旺苍(lng结帐ID As Long) As Boolean
     '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim rs结算记录 As New ADODB.Recordset
    
    Dim strInput As String, strOutput  As String
    Dim lng冲销ID As Long, str住院号 As String
    Dim strArr
    Dim lng病人ID As Long, intMouse As Integer
    Err = 0: On Error GoTo ErrHand:
    
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "结算冲销")
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=" & TYPE_广元旺苍 & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "结算冲销")
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    g病人身份_广元旺苍.病人ID = Nvl(rsTemp!病人ID, 0)
    
    
    gstrSQL = "select * from 医保结算记录 where 性质=2  and 结帐ID=" & lng结帐ID
    Call OpenRecordset_广元旺苍(rs结算记录, "结算冲销")
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    lng病人ID = g病人身份_广元旺苍.病人ID
    
    Screen.MousePointer = 1
    If 身份标识_广元旺苍(88, g病人身份_广元旺苍.病人ID) = "" Then
        Screen.MousePointer = intMouse
        住院结算冲销_广元旺苍 = False
        Exit Function
    End If
    Screen.MousePointer = intMouse
    If lng病人ID <> g病人身份_广元旺苍.病人ID Then
        ShowMsgbox "不是当前要冲销结算的病人!"
        Exit Function
    End If
    
    
    '判断病人的住院结算数据是否允许作废。判断标准是检查病人有新的住院记录，如果有，就不能冲销
    If Can住院结算冲销(rsTemp("病人ID"), rsTemp("主页ID")) = False Then Exit Function
    
    str住院号 = Nvl(rsTemp("支付顺序号"))
    
    strInput = str住院号
    If 业务请求_广元旺苍(取消出院结算_旺苍, strInput, strOutput) = False Then
        Exit Function
    End If
    
    '保存数据
    '    性质_IN     IN 医保结算记录.性质%TYPE,
    '    结帐ID_IN   IN 医保结算记录.结帐ID%TYPE,
    '    冲销ID_IN   IN 医保结算记录.结帐ID%TYPE)
    
    gcnOracle_广元旺苍.BeginTrans
    gstrSQL = "ZL_医保结算记录_冲销("
    gstrSQL = gstrSQL & "2"
    gstrSQL = gstrSQL & "," & lng结帐ID
    gstrSQL = gstrSQL & "," & lng冲销ID & ")"
    ExecuteProcedure_广元旺苍 "保存结算记录"
    
 
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(),累计统筹报销_IN(年内已报销金额),住院次数_IN(主页ID),起付线(起付金额),封顶线_IN(基本封顶金额),实际起付线_IN(基本报销比例),
    '   发生费用金额_IN(费用总额),全自付金额_IN(补充报销比例),首先自付金额_IN(公务员报销比例),
    '   进入统筹金额_IN(基本报销金额),统筹报销金额_IN(补充报销金额),大病自付金额_IN(公务员报销金额),超限自付金额_IN(无),个人帐户支付_IN(),"
    '   支付顺序号_IN(住院号),主页ID_IN(主页ID),中途结帐_IN,备注_IN(中心)
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
   
    '---------------------------------------------------------------------------------------------
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_广元旺苍 & "," & rsTemp("病人ID") & "," & Year(zlDatabase.Currentdate) & "," & _
        "NULL,NULL,NULL," & Nvl(rsTemp("累计统筹报销"), 0) * -1 & "," & Nvl(rsTemp!主页ID, 0) & "," & Nvl(rsTemp("起付线"), 0) * -1 & "," & Nvl(rsTemp("封顶线"), 0) * -1 & "," & Nvl(rsTemp("实际起付线"), 0) * -1 & "," & _
        Nvl(rsTemp("发生费用金额"), 0) * -1 & "," & Nvl(rsTemp("全自付金额"), 0) * -1 & "," & Nvl(rsTemp("首先自付金额"), 0) * -1 & "," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & "," & Nvl(rsTemp!大病自付金额, 0) * -1 & ",0," & Nvl(rsTemp!个人帐户支付, 0) * -1 & ",'" & _
        str住院号 & "'," & Nvl(rsTemp!主页ID, 0) & ",NULL,'" & Nvl(rsTemp!备注) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存医保结算记录")
    gcnOracle_广元旺苍.CommitTrans
    
    住院结算冲销_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function 处方登记_广元旺苍(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------
   '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    Dim str处方记录号 As String, str摘要 As String
    Dim strArr
    
    
    处方登记_广元旺苍 = False
    
    
   '读出该张单据的费用明细
  gstrSQL = "Select A.ID,a.标识号 住院号,a.序号,a.NO,A.记录性质,A.记录状态,A.病人ID,A.主页ID,to_char(A.发生时间,'yyyy-mm-dd') as 登记时间,Round(A.实收金额,4) 实收金额 " & _
              "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,M.中心,Q.名称 as 开单部门,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位 " & _
              "  From 病人费用记录 A,收费细目 B,保险帐户 M,部门表 Q,病人信息 J" & _
              "  where A.NO='" & str单据号 & "' and A.记录性质=" & lng记录性质 & " and A.记录状态 = " & lng记录状态 & " And Nvl(A.是否上传,0)=0 " & _
              "        and A.收费细目ID=B.ID and A.病人ID=J.病人ID  and A.主页ID=J.住院次数 And M.险类=" & TYPE_广元旺苍 & _
              "        and a.病人id=m.病人id and a.开单部门id=q.id(+)" & _
              "  Order by A.病人ID,A.记录性质,a.记录状态,A.NO,A.序号,A.发生时间"
        
    Call OpenRecordset(rs明细, "处方明细上传")
    If InitInfor_广元旺苍.明细时实上传 = False Then
        处方登记_广元旺苍 = True
        Exit Function
    End If
    
    Err = 0
    On Error GoTo ErrHand:
    gcnOracle.BeginTrans
    
    Dim lng处方号 As Long
    
    lng病人ID = 0
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            gstrSQL = "Select * From 医保支付项目 where 险类=" & TYPE_广元旺苍 & " and 中心=" & Nvl(!中心, 0) & " and 收费细目id=" & Nvl(!收费细目ID, 0)
            Call OpenRecordset(rsTemp, "确定医保支付项目")
            If rsTemp.EOF Then
                ShowMsgbox "注意：" & vbCrLf & "   收费细目为:[" & Nvl(!编码) & "]" & Nvl(!名称) & " 还未进行医保对码!"
            End If
            
            '增加处方明细
            '    AZYH    PCHAR   住院号
            '    ACFID   PCHAR   处方单号(在整个数据库中保证唯一)
            '    ACFMXID PCHAR   明细序号(在一个处方中保证唯一)
            '    ACFRQ   PCHAR   处方日期（YYYY-MM-DD）
            '    AYS PCHAR   医生
            '    AKS PCHAR   科室
            '    AYPBH   PCHAR   药品编号(社保药品编号)
            '    ASL PCHAR   数量(可以为负数)
            '    ADJ PCHAR   单价
            If lng病人ID <> Nvl(!病人ID, 0) Then
                lng病人ID = Nvl(!病人ID, 0)
                lng处方号 = !ID
            End If
            lng主页ID = Nvl(!主页ID, 0)
            
            strInput = Rpad(lng主页ID, 4, "0") & lng病人ID
             If lng主页ID > 9 Then strInput = "9" & Lpad(lng主页ID, 3, "0") & lng病人ID '宋献平  2006年12月19日  处理住院次数为1和10时重复问题
            strInput = strInput & vbTab & lng处方号 'Val(Mid(Nvl(!登记时间), 3, 4)) & Val(Substr(Nvl(!登记时间, "05"), 3, 2)) & Mid(Nvl(!no), 2) & Lpad(!记录性质, 3, "0") & Lpad(!记录状态, 3, "0") '如果是多病人单，可以考虑将病人id加入其中
            strInput = strInput & vbTab & Nvl(!序号)
            strInput = strInput & vbTab & Nvl(!登记时间)
            strInput = strInput & vbTab & Nvl(!医生)
            strInput = strInput & vbTab & Nvl(!开单部门)
            
            If rsTemp.EOF Then
                strInput = strInput & vbTab & ""
            Else
                strInput = strInput & vbTab & Nvl(rsTemp!项目编码)
            End If
            
            '曾明春:2005-07-06 如果单价精度超过2位小数，则数量传1，单价传实收金额。
            If Round(Nvl(!价格) * 100) = Nvl(!价格) * 100 Then
               strInput = strInput & vbTab & Nvl(!数量)
               strInput = strInput & vbTab & Nvl(!价格)
            Else
               strInput = strInput & vbTab & "1"
               strInput = strInput & vbTab & Nvl(!实收金额)
            End If
            
            If rsTemp.EOF Then
                '需对码才能上传
            Else
                If 业务请求_广元旺苍(增加处方单据_旺苍, strInput, strOutput) = False Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
                strOutput = Replace(strOutput, "@$", "||")
                '处方明细记录号@$自费比例@$自费金额
                '摘要保存值:处方号||自费比例||自费金额||住院号
                str摘要 = lng处方号 & "||" & strOutput & "||" & Rpad(lng主页ID, 4, "0") & lng病人ID
          
                If lng主页ID > 9 Then str摘要 = lng处方号 & "||" & strOutput & "||" & "9" & Lpad(lng主页ID, 3, "0") & lng病人ID
                '更改上传标志
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str摘要 & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            End If
            .MoveNext
        Loop
    End With
    gcnOracle.CommitTrans
    处方登记_广元旺苍 = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
End Function
Private Function Get治渝情况_旺苍(lng病人ID As Long, lng主页ID As Long) As String
    '功能:获取治渝情况标识

    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.出院情况" & _
             " From 诊断情况 A,疾病编码目录 B " & _
             " Where A.病人ID=" & lng病人ID & " And A.疾病ID=B.ID(+) And A.主页ID=" & lng主页ID & _
             "       And A.诊断类型 in (2,3)" & _
             " Order by A.诊断类型 Desc"
    
    rsInNote.CursorLocation = adUseClient
    Call OpenRecordset(rsInNote, "医保接口", strTmp)
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = Nvl(rsInNote!出院情况)
    End If
    If strTmp = "" Then
        strTmp = "治愈"
        
    End If
   ' strTmp = Decode(strTmp, "治愈", "1", "好转", "2", "未愈", "3", "死亡", "4", "其他", "9", "1")
    Get治渝情况_旺苍 = strTmp
   Call WriteDebugInfor_大连("Get治渝情况_吉林", lng病人ID)
End Function


Private Function Read模拟数据(ByVal int业务类型 As 业务类型_广元旺苍, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--功  能:通过该功能读取模拟数据,以便测试
    '--入参数:
    '--出参数:
    '--返  回:字串
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim strName As String
    
    If int业务类型 = 读取卡内数据 Then
        strFile = App.Path & "\解析卡.txt"
    Else
        strFile = App.Path & "\模拟提交串.txt"
    End If
    
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    strName = Get交易代码(int业务类型, True)
    
    Dim blnStart As Boolean
    Dim strArr
    
    Err = 0
    On Error GoTo ErrHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If int业务类型 = 读取卡内数据 Then
                    strArr = Split(strText, vbTab & "|")
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab & "|"
                            End If
                            strArr = Split(strText, vbTab & "|")
                            
                            If Val(strArr(0)) = 1 Then
                                str = strArr(1)
                                Exit Do
                            End If
                        Else
                             If "<" & strName & ">" = strText Then
                                 blnStart = True
                             End If
                        End If
                        If "</" & strName & ">" = strText Then
                            Exit Do
                        End If
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
'    If InStr(1, strOutPutstring, "@$") <> 0 Then
'        strOutPutstring = Split(strOutPutstring, "@$")(1)
'    End If
    Exit Function
ErrHand:
    DebugTool Err.Description
    Exit Function
End Function
Private Sub OpenRecordset_广元旺苍(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_广元旺苍, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub
Private Function 补传住院明细记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '补传相关明细记录
    Dim cnTemp As New ADODB.Connection
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strInput  As String, strOutput As String
    Dim strArr
    Dim str住院号 As String, str处方记录号 As String
    Dim strNO  As String
    Dim strSQL As String, strTmp As String
    Dim str摘要 As String
    Dim str重传 As Boolean
    Dim str传 As Boolean
    Err = 0
    On Error GoTo ErrHand:
      
    
    Call DebugTool("打开新连接")
    cnTemp.ConnectionString = gcnOracle.ConnectionString
    cnTemp.Open
    Call DebugTool("打开连接成功，开始检查明细数据的合法性。")
    
      
      
    str重传 = False
    补传住院明细记录 = False
    
    gstrSQL = "Select A.ID,a.标识号 住院号,A.摘要,a.序号,a.NO,A.记录性质,A.记录状态,A.病人ID,A.主页ID,to_char(A.发生时间,'yyyy-mm-dd') as 登记时间,Round(A.实收金额,4) 实收金额 " & _
                "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
                "         ,M.中心,Q.名称 as 开单部门,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,Nvl(A.是否上传,0) as 上传标志 " & _
                "  From 病人费用记录 A,收费细目 B,保险帐户 M,部门表 Q,病人信息 J" & _
                "  where Nvl(附加标志,0)<>9 and nvl(a.实收金额,0)<>0 " & _
                "        and A.收费细目ID=B.ID and A.病人ID=J.病人ID   and A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & " And M.险类=" & TYPE_广元旺苍 & _
                "        and a.病人id=m.病人id and a.开单部门id=q.id(+)" & _
                "  Order by A.病人ID,A.记录性质,a.记录状态,A.NO,A.序号,A.发生时间"
                
    Call OpenRecordset(rs明细, "虚拟结算", strSQL)
    If MsgBox("是否要全部重新上传病人费用明细?", vbInformation + vbYesNo, "中联软件") = vbYes Then str重传 = True
   With rs明细
        strNO = ""
        Do While Not .EOF
            strTmp = Nvl(!记录性质, 0) & "_" & Nvl(!记录状态, 0) & "_" & Nvl(!NO, 0)
            If strNO <> strTmp Then
                strNO = strTmp
                str处方记录号 = Split(Nvl(!摘要) & "||||", "||")(0)
                If str处方记录号 = "" Then
                   str处方记录号 = Nvl(!ID)
                End If
            Else
                If str处方记录号 = "" Then
                   str处方记录号 = Nvl(!ID)
                End If
            End If
            If !上传标志 = 0 Then str传 = True Else str传 = False
            If str重传 = True Then str传 = True
            If str传 = True Then
                gstrSQL = "Select * From 医保支付项目 where 险类=" & TYPE_广元旺苍 & " and 中心=" & Nvl(!中心, 0) & " and 收费细目id=" & Nvl(!收费细目ID, 0)
                Call OpenRecordset(rsTemp, "确定医保支付项目")
                If rsTemp.EOF Then
                    ShowMsgbox "注意：" & vbCrLf & "   收费细目为:[" & Nvl(!编码) & "]" & Nvl(!名称) & " 还未进行医保对码,请立即对码!"
                    Exit Function
                End If
 
               
                '增加处方明细
                '    AZYH    PCHAR   住院号
                '    ACFID   PCHAR   处方单号(在整个数据库中保证唯一)
                '    ACFMXID PCHAR   明细序号(在整个数据库中保证唯一)
                '    ACFRQ   PCHAR   处方日期（YYYY-MM-DD）
                '    AYS PCHAR   医生
                '    AKS PCHAR   科室
                '    AYPBH   PCHAR   药品编号(社保药品编号)
                '    ASL PCHAR   数量(可以为负数)
                '    ADJ PCHAR   单价
    
                
                strInput = Rpad(lng主页ID, 4, "0") & lng病人ID
                 If lng主页ID > 9 Then strInput = "9" & Lpad(lng主页ID, 3, "0") & lng病人ID '宋献平  2006年12月19日  处理住院次数为1和10时重复问题
                strInput = strInput & vbTab & str处方记录号 ' Val(Mid(Nvl(!登记时间), 3, 4)) & Val(Substr(Nvl(!登记时间, "05"), 3, 2)) & Mid(Nvl(!no), 2) & Lpad(!记录性质, 3, "0") & Lpad(!记录状态, 3, "0") '如果是多病人单，可以考虑将病人id加入其中
                strInput = strInput & vbTab & Nvl(!ID)
                strInput = strInput & vbTab & Nvl(!登记时间)
                strInput = strInput & vbTab & Nvl(!医生)
                strInput = strInput & vbTab & Nvl(!开单部门)
                If rsTemp.EOF Then
                    strInput = strInput & vbTab & ""
                Else
                    strInput = strInput & vbTab & Nvl(rsTemp!项目编码)
                End If
                '曾明春(2006-05-16):资阳机车厂医保单价和金额支持位小数
                If InitInfor_广元旺苍.适用地区 = 0 Then
                    strInput = strInput & vbTab & Nvl(!数量)
                    strInput = strInput & vbTab & Nvl(!价格)
                Else
                    '曾明春:2005-07-06 如果单价精度超过2位小数，则数量传1，单价传实收金额。
                    If Round(Nvl(!价格) * 100) = Nvl(!价格) * 100 Then
                       strInput = strInput & vbTab & Nvl(!数量)
                       strInput = strInput & vbTab & Nvl(!价格)
                    Else
                       strInput = strInput & vbTab & "1"
                       strInput = strInput & vbTab & Nvl(!实收金额)
                    End If
                End If
                
                
                If 业务请求_广元旺苍(增加处方单据_旺苍, strInput, strOutput) = False Then
                    Exit Function
                End If
                strOutput = Replace(strOutput, "@$", "||")
                '自费比例@$自费金额
                '摘要保存值:处方号||自费比例||自费金额||住院号
                str摘要 = str处方记录号 & "||" & strOutput & "||" & Rpad(lng主页ID, 4, "0") & lng病人ID
                If lng主页ID > 9 Then strInput = str处方记录号 & "||" & strOutput & "||" & "9" & Lpad(lng主页ID, 3, "0") & lng病人ID
                '更改上传标志
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str摘要 & "')"
                cnTemp.Execute gstrSQL, , adCmdStoredProc
            End If
            .MoveNext
        Loop
    End With
    补传住院明细记录 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_广元旺苍(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    'rsExse:字符集
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng主页ID As Long
    Dim strInput As String, strOutput   As String
    Dim strArr As Variant
    Dim str住院号 As String, str结算方式 As String
    Dim lng病人id1 As Long
    Dim intMouse As Integer
    
    Err = 0: On Error GoTo ErrHand:
    g病人身份_广元旺苍.病人ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
    intMouse = Screen.MousePointer
    
    gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "虚拟结算")
    
    If IsNull(rsTemp("主页ID")) = True Then
        MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    lng主页ID = rsTemp("主页ID")
    
'    If bln结帐处 Then
'        Screen.MousePointer = 1
'        If 身份标识_广元旺苍(4, lng病人id1) = "" Then
'            Screen.MousePointer = intMouse
'            住院虚拟结算_广元旺苍 = ""
'            Exit Function
'        End If
'        Screen.MousePointer = intMouse
'        If lng病人id <> lng病人id1 Then
'            ShowMsgbox "不是当前要结算的病人!"
'            Exit Function
'        End If
'    End If
    
    gstrSQL = "Select b.住院号,a.中心 From 保险帐户 a,病人信息 b  where a.病人id=b.病人id and a.病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取住院号"
    If rsTemp.EOF Then
        ShowMsgbox "该病人不是医保病人!"
        Exit Function
    End If
    
    str住院号 = Rpad(lng主页ID, 4, "0") & lng病人ID
   If lng主页ID > 9 Then str住院号 = "9" & Lpad(lng主页ID, 3, "0") & lng病人ID

    g病人身份_广元旺苍.社保中心 = Nvl(rsTemp!中心, 0)
    
    
    Screen.MousePointer = vbHourglass
   
    With rsExse
        g病人身份_广元旺苍.费用总额 = 0
        Do While Not .EOF
            g病人身份_广元旺苍.费用总额 = g病人身份_广元旺苍.费用总额 + Nvl(!金额, 0)
            .MoveNext
        Loop
    End With
     
    If 补传住院明细记录(lng病人ID, lng主页ID) = False Then Exit Function
    
    'AZYH    PCHAR   住院号
    'ISPREV  PCHAR   预结算标志（'0'－表示预结算）
    'ZLXG    PCHAR   治疗效果
    'CYZD    PCHAR   出院诊断1+列间隔符+ 出院诊断2+列间隔符+ 出院诊断3+列间隔符+出院诊断4
    'CYRQ    PCHAR   出院日期（YYYY-MM-DD）

    strInput = str住院号
    strInput = strInput & vbTab & gstrUserName
    strInput = strInput & vbTab & "0"
    strInput = strInput & vbTab & ""
    strInput = strInput & vbTab & ""
    strInput = strInput & vbTab & ""
    
    If 业务请求_广元旺苍(出院结算_旺苍, strInput, strOutput) = False Then Exit Function
    '返回:OK@$住院费用结算结果@$报销分段明细
    '   说明:
    '       住院费用结算结果:基本医疗待遇状态||起付金额||基本封顶金额||基本报销比例||年内已报销金额||基本报销金额||补充报销金额||公务员报销金额||补充医疗待遇状态||公务员待遇状态||补充报销比例||公务员报销比例||本次住院费用||甲类费用||甲类药品费||甲类诊疗费||甲类服务费||乙类费用||乙类药品费||乙类诊疗费||乙类手术费||乙类自付||丙类费用||丙类药品费||丙类诊疗费||丙类服务费||报销合计||个人支付
    '       报销分段明细(多条):险种||名称||段起始金额||段终止金额||本段基数||本段报销比例||本段报销金额||本段自付金额@$.......
    '   资阳三院多返回2个值[个人帐户支付]、[人员分类]:
    '       住院费用结算结果:基本医疗待遇状态||起付金额||基本封顶金额||基本报销比例||年内已报销金额||基本报销金额||补充报销金额||公务员报销金额||补充医疗待遇状态||公务员待遇状态||补充报销比例||公务员报销比例||本次住院费用||甲类费用||甲类药品费||甲类诊疗费||甲类服务费||乙类费用||乙类药品费||乙类诊疗费||乙类手术费||乙类自付||丙类费用||丙类药品费||丙类诊疗费||丙类服务费||报销合计||个人支付||个人帐户支付||人员分类
    '       报销分段明细(多条):险种||名称||段起始金额||段终止金额||本段基数||本段报销比例||本段报销金额||本段自付金额@$.......
        

    strArr = Split(strOutput, "||")
    With g结算数据
        .基本报销金额 = Val(strArr(5))
        .补充报销金额 = Val(strArr(6))
        .公务员报销金额 = Val(strArr(7))
        If InitInfor_广元旺苍.适用地区 = 0 Then
           .个人帐户支付金额 = Val(strArr(28))
           .人员分类 = Trim(strArr(28))
        End If
    End With
     If Format(strArr(12), "####0.00;-####0.00;0;0") <> Format(g病人身份_广元旺苍.费用总额, "####0.00;-####0.00;0;0") Then
        ShowMsgbox "结算数据不等!" & vbCrLf & "医保中心费用总额:" & Format(strArr(12), "####0.00;-####0.00;0;0") & vbCrLf & " 医院端为:" & Format(g病人身份_广元旺苍.费用总额, "####0.00;-####0.00;0;0")
        If InitInfor_广元旺苍.数据不等不可结算 Then
            Exit Function
        End If
    End If
    
    str结算方式 = "基本医疗报销;" & g结算数据.基本报销金额 & ";0"
    str结算方式 = str结算方式 & "|补充报销;" & g结算数据.补充报销金额 & ";0"
    str结算方式 = str结算方式 & "|公务员报销;" & g结算数据.公务员报销金额 & ";0"
    If InitInfor_广元旺苍.适用地区 = 0 Then
       str结算方式 = str结算方式 & "|个人帐户支付;" & g结算数据.个人帐户支付金额 & ";0"
    End If
    住院虚拟结算_广元旺苍 = str结算方式
    g病人身份_广元旺苍.病人ID = lng病人ID   '表示该病人已经进行了虚拟结算
    
    '打印结算单:
'    If InitInfor_广元旺苍.打印结算单 Then
        '调打印接口
        '    ASTARTZYH   PCHAR   打印开始住院号
        '    AENDZYH PCHAR   打印结束住院号
            
'        StrInput = str住院号 & "||"
'        StrInput = StrInput & str住院号
'        Call 业务请求_广元旺苍(打印结算报表_旺苍, StrInput, strOutput)
'    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 医保设置_广元旺苍(ByVal lng险类 As Long, ByVal lng医保中心 As Integer) As Boolean
    医保设置_广元旺苍 = frmSet广元旺苍.参数设置
End Function
Public Sub ExecuteProcedure_广元旺苍(ByVal strCaption As String)
'功能：执行SQL语句
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_广元旺苍.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Private Function 门诊明细写入(ByVal rs明细 As ADODB.Recordset, Optional ByVal bln虚拟 As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传明细记录
    '--入参数:
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    Dim str明细 As String
    Dim strTmpArr As Variant, strArr As Variant
    
    门诊明细写入 = False
    g病人身份_广元旺苍.费用总额 = 0
    
    Err = 0:    On Error GoTo ErrHand:
    '然后插入处方明细
    With rs明细
        If .RecordCount = 0 Then
            ShowMsgbox "不存在相关的明细费用记录!"
            Exit Function
        End If
        'YBJGBH  PCHAR   保险机构编号
        'CFH PCHAR   处方号
        'CFMXDATA    PCHAR   处方明细数据    格式说明：处方1(医保药品编号+列间隔符+单价+列间隔符数量+)+行间隔符+
        'CPASSWORD   PCHAR   持卡人卡密码
        'CCZYXM  PCHAR   操作员姓名
        strInput = g病人身份_广元旺苍.机构编码
        strInput = strInput & vbTab & Nvl(!NO)
        
        Do While Not rs明细.EOF
            gstrSQL = "Select * From 医保支付项目 where 险类=" & TYPE_广元旺苍 & " and 中心=" & g病人身份_广元旺苍.社保中心 & " and 收费细目id=" & Nvl(!收费细目ID, 0)
            Call OpenRecordset(rsTemp, "确定医保支付项目")
            If Val(Nvl(rs明细("实收金额"), 0)) <> 0 Then
                If rsTemp.EOF Then
                    str明细 = str明细 & "@$" & ""
                Else
                    str明细 = str明细 & "@$" & Nvl(rsTemp!项目编码)
                End If
                '曾明春(2006-05-16):资阳机车厂单价和数量支持4位小数
                If InitInfor_广元旺苍.适用地区 = 0 Then
                    str明细 = str明细 & "||" & Nvl(!单价)
                    str明细 = str明细 & "||" & Nvl(!数量)
                Else
                    '曾明春(2005-07-06):如果单价精度超过2位小数，则数量传1，单价传实收金额。
                    '曾明春(2005-12-12):修改为使用列间隔符"||"连接
                    If Round(Nvl(!单价) * 100) = Nvl(!单价) * 100 Then
                       str明细 = str明细 & "||" & Nvl(!单价)
                       str明细 = str明细 & "||" & Nvl(!数量)
                    Else
                       str明细 = str明细 & "||" & Nvl(!实收金额)
                       str明细 = str明细 & "||" & "1"
                    End If
                End If
                '曾明春(2006-05-16):增加传入明细ID
                If InitInfor_广元旺苍.适用地区 = 0 Then
                    str明细 = str明细 & "||" & Nvl(rs明细!ID)
                End If
            End If
            g病人身份_广元旺苍.费用总额 = g病人身份_广元旺苍.费用总额 + Nvl(!实收金额, 0)
            
            '写上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            
            rs明细.MoveNext
        Loop
    End With
    If g病人身份_广元旺苍.直输金额 = False Then
        str明细 = Mid(str明细, 3)
        strInput = strInput & vbTab & str明细
        strInput = strInput & vbTab & g病人身份_广元旺苍.密码
        strInput = strInput & vbTab & gstrUserName
        
        If 业务请求_广元旺苍(个人帐户消费_旺苍, strInput, strOutput) = False Then Exit Function
        If strOutput = "" Then Exit Function
        
        If InitInfor_广元旺苍.适用地区 = 0 Then
            strTmpArr = Split(strOutput, "@$")
            strArr = Split(strTmpArr(0), "||")
            If 更新处方明细信息(strTmpArr) = False Then Exit Function
        Else
            strArr = Split(strOutput, "||")
        End If
        With g结算数据
            .卡号 = strArr(0)
            .姓名 = strArr(1)
            .消费前帐户余额 = Val(strArr(2))
            .个人帐户支付金额 = Val(strArr(3))
            .自费金额 = Val(strArr(4))
            .消费后帐户余额 = Val(strArr(5))
            .交易时间 = strArr(6)
            .前端单据号 = strArr(7)
            .中心单据号 = strArr(8)
            .处方号 = strArr(9)
            .操作员姓名 = strArr(10)
            .前端名称 = strArr(11)
            If InitInfor_广元旺苍.适用地区 = 0 Then
               .人员分类 = strArr(12)
            End If
        End With
    End If
    门诊明细写入 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 结算方式更正() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:更正及显示结算结果
    '--入参数:
    '--出参数:str结算方式
    '--返  回:成功返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    Dim dbl费用总额 As Double
        
    '费用总额=病人自费金额+基本统筹支付金额+大病统筹金额      此解释是由刘兴宏根据以面公式转换而来的
    
    '病人自费金额 = 总费用额 - 基本统筹支付金额 - 大病 / 高额统筹支付金额
    '自费金额＝现金支付额＋帐户支付额 (即:可选择由现金或用帐户支付)
    '大病统筹与高额统筹意义相同
    '统筹支付金额等于医保内费用根据不同的起付标准和报销比例由医保中心算
    '此说明依据北京科瑞奇技术开发股份有限公司蒋红彬负责的解释
    结算方式更正 = False
    
    Err = 0:    On Error GoTo ErrHand:
    DebugTool "进入(" & "Get结算方式" & ")"
    
    '卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
    dbl费用总额 = g结算数据.个人帐户支付金额 + g结算数据.自费金额
    str结算方式 = "||个人帐户|" & g结算数据.个人帐户支付金额
    
    If Format(g病人身份_广元旺苍.费用总额, "###0.00;-###0.00;0;0") <> Format(dbl费用总额, "###0.00;-###0.00;0;0") Then
        Dim blnYes As Boolean
        '费用总额与医保中心返回总额不致,不能进行结算
        ShowMsgbox "本次结算总额(" & g病人身份_广元旺苍.费用总额 & ") 与" & vbCrLf & _
                    "   中心返回的总额(" & dbl费用总额 & ")，结算数据发生错误，请立即与医保中心联系!"
        Exit Function
    End If
    If g病人身份_广元旺苍.个人帐户支付 <> g结算数据.个人帐户支付金额 Then
        ShowMsgbox "本次虚拟结算个人帐户支付(" & g病人身份_广元旺苍.个人帐户支付 & ") 与" & vbCrLf & _
                    "   结算的个人帐户支付(" & g结算数据.个人帐户支付金额 & ")不等，结算数据发生错误，请立即与医保中心联系!"
        Exit Function
    End If
    结算方式更正 = True
'
'    Exit Function
'   '如果存在,则保存冲预交记录中
'    If str结算方式 <> "" Then
'        str结算方式 = Mid(str结算方式, 3)
'        g病人身份_成都内江.结算方式 = str结算方式
'
'        If g结算数据.结算标志 = 0 Then
'            gstrSQL = "zl_病人结算记录_Update(" & g结算数据.结帐ID & ",'" & str结算方式 & "', 0)"
'            Call zldatabase.ExecuteProcedure(gstrsql,"更新预交记录")
'        Else
'                gstrSQL = "zl_病人结算记录_Update(" & g结算数据.结帐ID & ",'" & str结算方式 & "',1)"
'                Call zldatabase.ExecuteProcedure(gstrsql,"更新预交记录")
'        End If
'    End If
'
'    DebugTool "开始显示结算方式"
'    '显示结算信息
'    If frm结算信息.ShowME(g结算数据.结帐ID, False, "", IIf(g结算数据.结算标志 = 0, 0, 1)) = False Then
'
'        DebugTool "结算方式显示失败"
'        结算方式更正 = False
'        Exit Function
'    End If
    DebugTool "结算方式显示成功"
    结算方式更正 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function 获取个人帐户支付() As Double
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取个人帐户值(从预交记录中获取)
    '--入参数:
    '--出参数:
    '--返  回:成功,返回本次个人帐户支付,否则返回0
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 金额 From 病人预交记录 where 结帐ID=" & g结算数据.结帐ID & " and  结算方式='个人帐户'"
    
    OpenRecordset rsTemp, "获取个人帐户支付"
    If Not rsTemp.EOF Then
        获取个人帐户支付 = Nvl(rsTemp!金额, 0)
    End If
    
End Function
Private Function InsertInto医保结算记录(ByVal strArr As Variant, ByVal lng结帐ID As Long) As Boolean
    '功能:往中间库插入医保结算记录
    '参数:strarr以split(stroutput,"||")产生的数组
    '返回:strArr(0)-住院费用结算结果,strArr(1-n)报销分段明细
      '   说明:
      '       住院费用结算结果:基本医疗待遇状态||起付金额||基本封顶金额||基本报销比例||年内已报销金额||基本报销金额||补充报销金额||公务员报销金额||补充医疗待遇状态||公务员待遇状态||补充报销比例||公务员报销比例||本次住院费用||甲类费用||甲类药品费||甲类诊疗费||甲类服务费||乙类费用||乙类药品费||乙类诊疗费||乙类手术费||乙类自付||丙类费用||丙类药品费||丙类诊疗费||丙类服务费||报销合计||个人支付
      '       报销分段明细(多条):险种||名称||段起始金额||段终止金额||本段基数||本段报销比例||本段报销金额||本段自付金额@$.......
    Dim tmpArr As Variant
    Dim i As Long
    Err = 0
    On Error GoTo ErrHand:
    InsertInto医保结算记录 = False
    
    '保存住院结算数据
    '曾明春(2006-01-04):资阳地区(适用地区为0)个人支付保存的是个人帐户支付金额
    tmpArr = Split(strArr(0), "||")
    
    DebugTool "进入InsertInto医保结算记录"
       
    '    性质        number(2),
    gstrSQL = "ZL_医保结算记录_INSERT(2"
    '    结帐ID      number(18),
    gstrSQL = gstrSQL & "," & lng结帐ID
    '    基本医疗待遇状态_IN IN 医保结算记录.基本医疗待遇状态%TYPE,
    gstrSQL = gstrSQL & ",'" & tmpArr(0) & "'"
    '    起付金额_IN IN 医保结算记录.起付金额%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(1)) & ""
    '    基本封顶金额_IN IN 医保结算记录.基本封顶金额%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(2)) & ""
    '    基本报销比例_IN IN 医保结算记录.基本报销比例%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(3)) & ""
    '    年内已报销金额_IN   IN 医保结算记录.年内已报销金额%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(4)) & ""
    '    基本报销金额_IN IN 医保结算记录.基本报销金额%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(5)) & ""
    '    补充报销金额_IN IN 医保结算记录.补充报销金额%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(6)) & ""
    '    公务员报销金额_IN   IN 医保结算记录.公务员报销金额%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(7)) & ""
    '    补充医疗待遇状态_IN IN 医保结算记录.补充医疗待遇状态%TYPE,
    gstrSQL = gstrSQL & ",'" & Trim(tmpArr(8)) & "'"
    '    公务员待遇状态_IN   IN 医保结算记录.公务员待遇状态%TYPE,
    gstrSQL = gstrSQL & ",'" & Trim(tmpArr(9)) & "'"
    '    补充报销比例_IN IN 医保结算记录.公务员报销比例%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(10)) & ""
    '    公务员报销比例_IN   IN 医保结算记录.公务员报销比例%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(11)) & ""
    '    本次住院费用_IN IN 医保结算记录.本次住院费用%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(12)) & ""
    '    甲类费用_IN IN 医保结算记录.甲类费用%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(13)) & ""
    '    甲类药品费_IN   IN 医保结算记录.甲类药品费%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(14)) & ""
    '    甲类诊疗费_IN   IN 医保结算记录.甲类诊疗费%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(15)) & ""
    '    甲类服务费_IN   IN 医保结算记录.甲类服务费%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(16)) & ""
    '    乙类费用_IN IN 医保结算记录.乙类费用%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(17)) & ""
    '    乙类药品费_IN   IN 医保结算记录.乙类药品费%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(18)) & ""
    '    乙类诊疗费_IN   IN 医保结算记录.乙类诊疗费%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(19)) & ""
    '    乙类手术费_IN   IN 医保结算记录.乙类手术费%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(20)) & ""
    '    乙类自付_IN IN 医保结算记录.乙类自付%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(21)) & ""
    '    丙类费用_IN IN 医保结算记录.丙类费用%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(22)) & ""
    '    丙类药品费_IN   IN 医保结算记录.丙类药品费%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(23)) & ""
    '    丙类诊疗费_IN   IN 医保结算记录.丙类诊疗费%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(24)) & ""
    '    丙类服务费_IN   IN 医保结算记录.丙类服务费%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(25)) & ""
    '    报销合计_IN IN 医保结算记录.报销合计%TYPE,
    gstrSQL = gstrSQL & "," & Val(tmpArr(26)) & ""
    '    个人支付_IN IN 医保结算记录.个人支付%TYPE
    If InitInfor_广元旺苍.适用地区 = 0 Then
       gstrSQL = gstrSQL & "," & Val(tmpArr(28)) & ")"
    Else
       gstrSQL = gstrSQL & "," & Val(tmpArr(27)) & ")"
    End If
    
    ExecuteProcedure_广元旺苍 "保存结算记录到中间库"
        
    '保存明细数据
    '曾明春(2005-07-26):病人未到起付线，返回的基本报销为空,必须进行判断
    For i = 1 To UBound(strArr)
        '保存明细数据
         '名称||段起始金额||段终止金额||本段基数||本段报销比例||本段报销金额||本段自付金额||险种
        '过程参数:
        '性质_IN     IN 医保结算分段明细.性质%TYPE,
        '结帐ID_IN   IN 医保结算分段明细.结帐ID%TYPE,
        '险种_IN     IN 医保结算分段明细.险种%TYPE,
        '名称_IN     IN 医保结算分段明细.名称%TYPE,
        '段起始金额_IN   IN 医保结算分段明细.段起始金额%TYPE,
        '段终止金额_IN   IN 医保结算分段明细.段终止金额%TYPE,
        '本段基数_IN IN 医保结算分段明细.本段基数%TYPE,
        '本段报销比例_IN IN 医保结算分段明细.本段报销比例%TYPE,
        '本段报销金额_IN IN 医保结算分段明细.本段报销金额%TYPE,
        '本段自付金额_IN IN 医保结算分段明细.本段自付金额%TYPE)

        tmpArr = Split(strArr(i), "||")
        If UBound(tmpArr) >= 7 Then
            gstrSQL = "ZL_医保结算分段明细_INSERT("
            gstrSQL = gstrSQL & "2"
            gstrSQL = gstrSQL & "," & lng结帐ID & ""
            gstrSQL = gstrSQL & ",'" & IIf(tmpArr(7) = "", "反回空值了", tmpArr(7)) & "'"
            '曾明春(2006-4-7):不知道名称这个参数为什么有时候会返回空,增加对空值的处理;同时可能存在多段补充的情况
            gstrSQL = gstrSQL & ",'" & IIf(tmpArr(0) = "", "报销" & i, tmpArr(7) & i) & "'"
            gstrSQL = gstrSQL & "," & Val(tmpArr(1)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(2)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(3)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(4)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(5)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(6)) & ")"
      
            ExecuteProcedure_广元旺苍 "保存结算分段信息到中间库"
        End If
    Next
    InsertInto医保结算记录 = True
    DebugTool "保存医保结算记录成功"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 撤消医保入院_川大金键(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intInsure As Integer) As Boolean
'功能：更新病人的出院疾病。如果是肿瘤，则结算时起付线会减半
    Dim strInput As String
    Dim strOutput As String
    Dim blnYes  As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    
    On Error GoTo errHandle
    '曾明春(2006-2-17):可能存在冲销费用,不能直接在入院登记中进行入院撤消
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        If MsgBox("该病人已经出院或无费用发生，可以通过入院登记进行入院撤消!但有可能存在冲销费用,点否只取消医保登记,点是退出!", vbYesNo) = vbYes Then
           Exit Function
        End If
    End If
    
    
    gstrSQL = "Select * From 病人费用记录 where nvl(是否上传,0)=1 and rownum<=1 and 病人id=" & lng病人ID & " and 主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断是否存在上传记录"
        
    If Not rsTemp.EOF Then
        ShowMsgbox "已经有上传到中心明细费用，是否真的要取消医保入院?", True, blnYes
        If blnYes = False Then
            Exit Function
        End If
    End If
    
    
    strInput = Rpad(lng主页ID, 4, "0") & lng病人ID
     If lng主页ID > 9 Then strInput = "9" & Lpad(lng主页ID, 3, "0") & lng病人ID '宋献平  2006年12月19日  处理住院次数为1和10时重复问题
    
    If 业务请求_广元旺苍(取消入院登记_旺苍, strInput, strOutput) = False Then Exit Function

    '更新医保帐户
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & intInsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    
    gstrSQL = "ZL_病案主页_撤消医保入院(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消医保入院")
    
    '处理上传标志
    gstrSQL = "update 病人费用记录 set 是否上传=0 where 结帐金额 is null and 病人ID= " & lng病人ID & " and 主页ID= " & lng主页ID
    gcnOracle.Execute gstrSQL
    
    DebugTool "取消成功"
    
    撤消医保入院_川大金键 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function 更新处方明细信息(ByVal strArr As Variant) As Boolean
    '功能:更新明细摘要字段,保存分割信息
    '参数:strarr以split(stroutput,"||")产生的数组
    '返回:strArr(0)个人帐户消费信息,strArr(1-n)报销分段明细
      '   说明:
      '        明细ID,
    Dim tmpArr As Variant
    Dim i As Long
    Err = 0
    On Error GoTo ErrHand:
    更新处方明细信息 = False
    
    '更新明细数据
    For i = 1 To UBound(strArr)
        tmpArr = Split(strArr(i), "||")
        If UBound(tmpArr) >= 3 Then
            gstrSQL = "ZL_病人费用记录_更新医保("
            gstrSQL = gstrSQL & tmpArr(0)
            gstrSQL = gstrSQL & ",Null,Null,Null,Null,Null"
            gstrSQL = gstrSQL & ",'" & IIf(IsNull(tmpArr(1)), "0", tmpArr(1)) & "|" & IIf(IsNull(tmpArr(2)), "0", tmpArr(2)) & "|" & IIf(IsNull(tmpArr(3)), "0", tmpArr(3)) & "')"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新结算分段信息到摘要")
        End If
    Next
    更新处方明细信息 = True
    DebugTool "保存医保结算记录成功"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
