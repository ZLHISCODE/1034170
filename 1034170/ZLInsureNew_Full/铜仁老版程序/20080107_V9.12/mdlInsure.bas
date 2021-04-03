Attribute VB_Name = "mdlInsure"
Option Explicit

Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As String
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_BROWSEFORCOMPUTER = &H1000  'Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000   'Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000 'Browsing for Everything
Private Const CSIDL_NETWORK As Long = &H12

Private Const MAX_PATH = 260
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2
Private Const LVM_SETCOLUMNWIDTH = &H101E

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'输入法控制API----------------------------------------------------------------------------------------------
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Public Const KLF_REORDER = &H8

'下列语句用于检测是否合法调用
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

'对文本串进行加密或解密的函数
Public Declare Function EncryptStr Lib "FTP_Trans.dll" (ByVal SourceStr As String, ByVal Key As String, ByVal IsEncrypt As Boolean) As String

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    部门 As String
    站点 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public glngSys As Long                      '系统编号参数
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSQL As String                    '用着作为所有临时SQL语句

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrDbUser As String                 '当前数据库用户
Public gstrUserName As String               '当前用户姓名
Public gstr单位名称 As String
Public gbln特殊门诊 As Boolean              '西铝专用,用于返回是否为特殊门诊
Public gstr特殊病种 As String               '特殊病种类

Public gstrMatchMethod As String    '匹配方式:0表示双向匹配

Public gintInsure As Integer
Public gstr医院编码 As String * 10               '医院编号
Public gstr医保机构编码 As String

Public Type T结算数据
    病人ID       As Long
    年度         As Long
    住院次数     As Long
    帐户累计增加   As Currency
    帐户累计支出   As Currency
    累计进入统筹   As Currency
    累计统筹报销   As Currency
    起付线         As Currency
    封顶线         As Currency
    实际起付线     As Currency
    发生费用金额   As Currency
    全自费金额   As Currency
    首先自付金额   As Currency
    进入统筹金额   As Currency
    统筹报销金额   As Currency
    超限自付金额   As Currency
    个人帐户支付   As Currency
    支付顺序号     As String
    主页ID         As Long
    中途结帐       As Long
    住院床日       As Long
End Type
Public g结算数据 As T结算数据           '保存预结算之后计算的结果，可用以填写保险结算记录
Public gcol结算计算 As New Collection   '保存预结算之后计算的结果，可用以填写保险结算计算
                                        '每个成员为一个数组，依次为档次、进入统筹金额、统筹报销金额、比例

Public Enum 医保Enum
    TYPE_重庆市 = 10
    TYPE_重庆松藻 = 11                  'Modified by ZYB ##2002-10-28
    TYPE_重庆壁山 = 12
    TYPE_重庆中梁山 = 13
    TYPE_重庆银海版 = 14
    TYPE_成都市 = 20
    TYPE_成都莲合 = 21
    TYPE_成都郊县 = 22
    TYPE_成都南充 = 23
    type_米易 = 24
    TYPE_四川眉山 = 25
    TYPE_乐山 = 26
    TYPE_云南省 = 30
    TYPE_昆明市 = 31
    TYPE_云南建水 = 32
    TYPE_自贡市 = 40
    TYPE_沈阳市 = 43
    TYPE_贵阳市 = 50
    TYPE_咸阳市 = 61
    TYPE_福建巨龙 = 70
    'Modified by 朱玉宝 20031218 地区：福州
    TYPE_福建省 = 71
    TYPE_福州市 = 72
    TYPE_南平市 = 73
    TYPE_泸州市 = 80
    TYPE_铜仁 = 81
    TYPE_大连市 = 82
    TYPE_大连开发区 = 83
    TYPE_涪陵 = 84
    TYPE_凯里 = 85
    TYPE_重大校园卡 = 86    '刘兴宏:200403
    TYPE_北京 = 90
    TYPE_广元 = 87
End Enum

Public Enum 余额Enum
    balan门诊 = 10
    balan入院 = 20
    balan预交 = 30
    balan结算 = 40
End Enum

Public Enum 身份验证Enum
    id门诊收费 = 0
    id入院登记 = 1
    id帐户管理 = 2
    id挂号 = 3
    id结帐 = 4
    id门诊确认 = 5
End Enum

Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤消出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付

    support门诊连续收费 = 22        '门诊在身份验证后，可进行多次收费操作
    support门诊收费完成后验证 = 23  '在门诊收费完成，是否再次调用身份验证
    
    support医嘱上传 = 24            '医嘱产生费用时是否实时传输
    support分币处理 = 25            '医保病人是否处理分币
    support中途结算仅处理已上传部分 = 26 '提供对已上传部分数据的结算功能
    support允许冲销已结帐的记帐单据 = 27 '是否允许冲销记帐单据，如果该单据已经结帐
End Enum

Public Function GetErrInfo(strCode As String) As String
'功能：根据错误代码返回错误信息
'参数：bytType=保险类别,strCode=错误代码
    Dim rsTmpErr As New ADODB.Recordset
    
    Select Case gintInsure
        Case TYPE_云南省, TYPE_昆明市, TYPE_云南建水
            Select Case strCode
                Case "0000":      GetErrInfo = "正常"
                Case "0001":      GetErrInfo = "无法读取配置文件，请关闭本程序后检查配置文件！"
                Case "0002":      GetErrInfo = "与应用程序服务器连接失败(无法找到应用程序服务器),请确认Socket Server是否正常启动!"
                Case "0003":      GetErrInfo = "应用程序服务器出错、无法完成交易!"
                Case "0004":      GetErrInfo = "无法得到系统配置信息!"
                Case "0005":      GetErrInfo = "找不到参保人所在中心的程序服务器，网络连接有问题!"
                Case "0009":      GetErrInfo = "检索不到该卡号对应的分中心编号"
                Case "1":         GetErrInfo = "终端设备不支持此功能"
                Case "10":        GetErrInfo = "验证密码,输入的个人密码错误"
                Case "1001":      GetErrInfo = "顺序号长度非法"
                Case "1002":      GetErrInfo = "收费项目大类编码非法"
                Case "1003":      GetErrInfo = "收费项目编码非法"
                Case "1004":      GetErrInfo = "数量或价格不能为空"
                Case "1005":      GetErrInfo = "数量或价格不能小于等于0"
                Case "1006":      GetErrInfo = "其它数据项非法"
                Case "11":        GetErrInfo = "支付交易,余额不足"
                Case "1101":      GetErrInfo = "顺序号错误"
                Case "1102":      GetErrInfo = "病人已结算不能再传递费用明细"
                Case "1103":      GetErrInfo = "没有检索到需要修改的费用明细资料!可能是输入参数不正确!"
                Case "1104":      GetErrInfo = "修改费用明细资料出错!"
                Case "1105":      GetErrInfo = "该病人住院费用已进入大病!乙类项目将视同甲类处理"
                Case "12":        GetErrInfo = "支付交易,用户卡初始化失败"
                Case "13":        GetErrInfo = "支付交易,SAM卡初始化失败"
                Case "14":        GetErrInfo = "支付交易,用户卡验证MAC1失败"
                Case "15":        GetErrInfo = "支付交易,SAM卡验证MAC2失败"
                Case "16":        GetErrInfo = "查余额交易,读取余额失败"
                Case "17":        GetErrInfo = "更新动态信息,用户卡更新错误"
                Case "18":        GetErrInfo = "未知卡类别"
                Case "19":        GetErrInfo = "更新动态信息,PSAM卡读取错误"
                Case "2":         GetErrInfo = "交易初始化,检测不到终端设备+卡设备类型"
                Case "20":        GetErrInfo = "无此数据项"
                Case "2001":      GetErrInfo = "经办人或科室名称非法,不能再进行结算!"
                Case "21":        GetErrInfo = "支付交易,TAC校验错误"
                Case "2101":      GetErrInfo = "病人已办理出院结算,不能再进行结算"
                Case "2102":      GetErrInfo = "病人审批未通过，在院期间的费用为全自费"
                Case "2103":      GetErrInfo = "费用结算时检测到存储过程的输入参数“顺序号”位数不正确!"
                Case "22":        GetErrInfo = "圈存交易,MAC1校验错误"
                Case "2201":      GetErrInfo = "不能提取病人的支付类别，无法进行费用结算<bnzxx>！"
                Case "2202":      GetErrInfo = "不能提取特殊病人的支付类别，无法进行费用结算<by21bzxx>！"
                Case "2203":      GetErrInfo = "费用结算时向By10cyjsb中写数据出错!"
                Case "2204":      GetErrInfo = "费用回退失败!"
                Case "2205":      GetErrInfo = "预结算失败!"
                Case "2206":      GetErrInfo = "公务员存储过程执行出错!"
                Case "2207":      GetErrInfo = "请首先进行费用的与结算之后再进行结算汇总!"
                Case "2208":      GetErrInfo = "没有有效的预结算纪录，无法回退!"
                Case "2209":      GetErrInfo = "住院数据已经超过月份限制，系统只能清本月的费用!"
                Case "2210":      GetErrInfo = "未查询到就诊登记或回退日期超过允许回退的最后期限！"
                Case "2211":      GetErrInfo = "当前回退结算记录不是最后一次结算，必须把当前结算记录之后的所有结算记录回退之后才能进行回退业务操作!"
                Case "23":        GetErrInfo = "圈存交易,TAC校验错误"
                Case "24":        GetErrInfo = "圈存交易,用户卡初始化失败"
                Case "25":        GetErrInfo = "圈存交易,用户卡验证MAC2失败"
                Case "29":        GetErrInfo = "更改密码失败"
                Case "3":         GetErrInfo = "交易初始化,检测不到PSAM卡"
                Case "30":        GetErrInfo = "无此交易代码"
                Case "3001":      GetErrInfo = "输入医院编码非法，请检查配置文件的设置！"
                Case "3002":      GetErrInfo = "卡类型不符或未插卡，请（重）插卡！"
                Case "3003":      GetErrInfo = "无法读取卡中信息，请重试！"
                Case "31":        GetErrInfo = "发送交易请求失败,请检查通讯端口"
                Case "3100":      GetErrInfo = "病人已办理住院登记，不能进行该项业务！"
                Case "3101":      GetErrInfo = "病人所使用的卡非法，不能凭卡享受医保待遇！"
                Case "3102":      GetErrInfo = "病人在其它医院未进行结算，无法进行该业务！"
                Case "3103":      GetErrInfo = "未建立此卡的基本资料，请重新输入!"
                Case "3104":      GetErrInfo = "不能获取病人的出生日期，不能进入院登记！"
                Case "3105":      GetErrInfo = "该病历号/住院号已经被占用，请输入别的住院号/病历号！"
                Case "3106":      GetErrInfo = "不能检索到“医院等级”数据，请核实医院编码！"
                Case "3107":      GetErrInfo = "密码检验不正确,请重新输入!"
                Case "3108":      GetErrInfo = "无法检索到特殊群体的比例参数!"
                Case "3109":      GetErrInfo = "执行特殊人群的存储过程、包出现错误!"
                Case "3110":      GetErrInfo = "医保中心编码与IC卡实际纪录的医保中心编码不一致!"
                Case "3111":      GetErrInfo = "公务员待遇审核时出错!"
                Case "3128":      GetErrInfo = "当前病人的就诊类别不是急诊，不能执行急诊转住院操作！"
                Case "32":        GetErrInfo = "接收响应数据超时或交易被取消,请检查通讯端口"
                Case "33":        GetErrInfo = "校验响应数据(ETX)错误"
                Case "34":        GetErrInfo = "校验响应数据(LRC)错误"
                Case "35":        GetErrInfo = "校验响应数据(STX)错误"
                Case "36":        GetErrInfo = "校验响应数据传输密钥错误"
                Case "37":        GetErrInfo = "接收响应数据失败,请检查通讯端口"
                Case "38":        GetErrInfo = "未知错误,操作被取消"
                Case "4":         GetErrInfo = "交易初始化,PSAM卡读取错误"
                Case "40":        GetErrInfo = "通讯错误"
                Case "4001":      GetErrInfo = "文件名非法"
                Case "4002":      GetErrInfo = "写文件过程出错"
                Case "41":        GetErrInfo = "圈存交易验证失败，请将卡交至医保中心处理"
                Case "4101":      GetErrInfo = "无费用明细信息"
                Case "42":        GetErrInfo = "磁卡封锁"
                Case "5":         GetErrInfo = "交易初始化,检测不到用户卡"
                Case "5001":      GetErrInfo = "支付原因长度非法,支付失败"
                Case "5002":      GetErrInfo = "支付金额应该大于0,支付失败"
                Case "5003":      GetErrInfo = "支付额大于卡上余额,支付失败"
                Case "5004":      GetErrInfo = "写卡失败"
                Case "5101":      GetErrInfo = "病人已出院或顺序号错误"
                Case "5102":      GetErrInfo = "本卡与登记时所使用的卡不符"
                Case "5103":      GetErrInfo = "无法得到有效的个人账户支付和现金支付。"
                Case "5104":      GetErrInfo = "病员尚未进行费用传输/或者检索不到该病员的支付数据。"
                Case "6":         GetErrInfo = "交易初始化,非本系统卡"
                Case "6101":      GetErrInfo = "无法享受待遇"
                Case "6102":      GetErrInfo = "卡支付额大于费用总额，无法支付"
                Case "7":         GetErrInfo = "用户卡读取错误"
                Case "7101":      GetErrInfo = "与中心数据库连接失败,请确认网络畅通以及NET8服务配置正确!"
                Case "7102":      GetErrInfo = "与前置机数据库连接失败,请确认Socket Server是否正常启动!"
                Case "7103":      GetErrInfo = "省医保中心数据下载失败"
                Case "7104":      GetErrInfo = "操作已经被取消！"
                Case "7106":      GetErrInfo = "市医保中心数据下载失败"
                Case "7107":      GetErrInfo = "省市医保中心数据下载失败!"
                Case "8":         GetErrInfo = "验证卡号不符"
                Case "8001":      GetErrInfo = "获取个人基本信息出错"
                Case "8002":      GetErrInfo = "还有有效的审批记录未结算，不能新审批"
                Case "8003":      GetErrInfo = "已经办理入院登记"
                Case "8004":      GetErrInfo = "该人员无法享受医疗待遇"
                Case "8005":      GetErrInfo = "审核医疗待遇享受资格时，系统出错"
                Case "8006":      GetErrInfo = "请先到医保中心进行资格审批"
                Case "8007":      GetErrInfo = "生成序号时出错"
                Case "8008":      GetErrInfo = "卡处于封锁状态"
                Case "8009":      GetErrInfo = "费用结算时，系统出错"
                Case "8010":      GetErrInfo = "没有有效的审批记录"
                Case "8011":      GetErrInfo = "有未结算的审批记录，请核实"
                Case "8012":      GetErrInfo = "全自费部分加挂钩自费部分大于费用总额，请核实"
                Case "8013":      GetErrInfo = "该审批记录审批未通过，作为不享受待遇结算，全部自费"
                Case "8014":      GetErrInfo = "该人员因个人封锁无法享受医疗待遇"
                Case "8015":      GetErrInfo = "该人员非医疗保险照顾人员"
                Case "8016":      GetErrInfo = "该人员为医疗保险照顾人员"
                Case "8017":      GetErrInfo = "当前费用已结算"
                Case "8080":      GetErrInfo = "医保中心尚未在贵医院开通此项医保业务,请与医保中心联系"
                Case "9":         GetErrInfo = "验证密码,用户卡个人密码被锁"
                Case "9001":      GetErrInfo = "应用服务器执行存储过程/程序包发生错误！"
                Case "9002":      GetErrInfo = "不能连接到本地数据库(hisint/hisintkm),无法进行交易处理!"
                Case "9003":      GetErrInfo = "向本地数据库中提交数据修改出错,无法对数据进行提交或者回滚!"
                Case "9004":      GetErrInfo = "该病人的基本资料还没有登记或者已经提交成功,无法回滚数据!"
                Case "9005":      GetErrInfo = "数据库中未检索到该病人的未提交的数据资料,无法提交数据!"
                Case "9006":      GetErrInfo = "外部应用程序传入的事物控制号的位数不为18位!"
                Case "9201":      GetErrInfo = "查询分段费用明细纪录数出错!"
                Case "9202":      GetErrInfo = "查询住院病人待遇变更记录数出错!"
                Case "9203":      GetErrInfo = "无效的查询类别!"
                Case "9204":      GetErrInfo = "待遇变更数据下载失败,无法查阅最新的变更信息!"
                Case "9205":      GetErrInfo = "待遇变更信息查询出错!"
                Case "9301":      GetErrInfo = "无法定位病人的医保机构，无法就诊!"
                Case "9996":      GetErrInfo = "省中心数据传输失败!"
                Case "9997":      GetErrInfo = "市中心数据传输失败!"
                Case "9998":      GetErrInfo = "省/市中心数据传输失败!"
                Case "9999":      GetErrInfo = "应用服务器发生异常错误"
                Case Else
                    GetErrInfo = "医保支持部分出现错误"
            End Select
            GetErrInfo = GetErrInfo & "[错误代码—" & strCode & "]"
        Case TYPE_成都市
            gstrSQL = "select errtext from errcode where code='" & strCode & "'"
            rsTmpErr.CursorLocation = adUseClient
            rsTmpErr.Open gstrSQL, gcnSybase, adOpenKeyset
            If Not rsTmpErr.EOF Then
                GetErrInfo = IIf(IsNull(rsTmpErr!errtext), "未知原因的错误。", rsTmpErr!errtext)
            Else
                GetErrInfo = "未知原因的错误"
            End If
        Case TYPE_大连开发区, TYPE_大连市
               Select Case Val(strCode)
                    Case 0: GetErrInfo = "正常"
                    Case -2:      GetErrInfo = "分配内存错，系统坏，重新启动系统可能解决"
                Case -1001, -1003, -1004, -1005 - 1006 - 1007:
                        GetErrInfo = "系统发生网络错误，请确认网络是否正常连接!"
                Case -1002:
                        GetErrInfo = "与中心连接错误,发生的原因可能是:" & vbCrLf & _
                                     "    （1）网络不通" & vbCrLf & _
                                     "    （2）服务程序运行失败" & vbCrLf & _
                                     "    （3）客户端配置错误" & vbCrLf & _
                                     "    （4）客户端配置医院编码错误" & vbCrLf & _
                                     "解决办法为:确认网络是否正常、确认服务是否正常工作"
                Case -5555
                    GetErrInfo = "读卡器读卡错误，IC卡非法或读卡器类型不匹配!"
                '周海全调试 2003-12-17
                Case -5556
                    GetErrInfo = "保号不一致！"
                Case -6001, -6002, -6003, -6004, -6005, -6007, -6008
                    GetErrInfo = "系统进行数据解析时错误,可能系统文件package.dat" & vbCrLf & _
                                 "文件遭到破坏或传递的参数值不对!"
                Case -6009
                    GetErrInfo = "数据中的医院编号和注册的医院编号不一致!"
                Case -6006, -7001
                    GetErrInfo = "系统进行合法性认证错误，可能设置系统参数错误" & vbCrLf & _
                                 "（注：如果非法使用，医保中心将冻结相关医院的结算能力）!"
                Case 1001
                    GetErrInfo = "不存在该保号!"
                Case 1002
                    GetErrInfo = "卡号错（验卡）!"
                Case 1003
                    GetErrInfo = "止付后进行支付，请验卡!"
                Case 1004
                    GetErrInfo = "无医院字典，医院编号错!"
                Case 1005
                    GetErrInfo = "医院已被冻结!"
                Case 1007
                    GetErrInfo = "数据时间大于当前系统时间，应用错误！"
                Case 1008
                    GetErrInfo = "治疗序号重复，IC卡上的数据错误，请验卡！"
                '周海全调试 2003-12-17
                Case 1009: GetErrInfo = "校验数据错误（特治费为负数）！"
                Case 1011
                    GetErrInfo = "结算信息中参保人的基本信息错误，请验卡！"
                Case 1016
                    GetErrInfo = "中心服务暂停，正在进行更新，请三分钟后再试！"
                Case 1020: GetErrInfo = "非法请求，业务类型错误，应用错误！"
                Case 1022: GetErrInfo = "不允许门诊！"
                Case 1023: GetErrInfo = "不允许住院！"
                Case 1024: GetErrInfo = "不允许重病！"
                Case 1025: GetErrInfo = "中心服务错误，请及时和中心服务管理员联系！"
                Case 1026: GetErrInfo = "结算序列号错误，医院系统被非法拷贝或遭到破坏！"
                Case 1027: GetErrInfo = "个人账户卡库不一致，请验卡！"
                Case 1028: GetErrInfo = "非法工作时间，中心服务暂停！"
                Case 1030: GetErrInfo = "医院门诊结算算法错误！"
                Case 1031: GetErrInfo = "医院门诊冲账结算算法错误！"
                Case 1032: GetErrInfo = "结算类别错误！"
                Case 1033: GetErrInfo = "医院住院结算算法错误！"
                Case 1034: GetErrInfo = "医院住院冲账结算算法错误！"
                Case 1035: GetErrInfo = "统筹累计错误！"
                Case 1036: GetErrInfo = "写卡序号错！"
                Case 1037: GetErrInfo = "卡号错误！"
                Case 1038: GetErrInfo = "账户余额错误！"
                Case 1039: GetErrInfo = "住院没有住院天数！"
                Case 1041: GetErrInfo = "转诊没有转诊单！"
                Case 1042: GetErrInfo = "门诊大病码错误！"
                Case 1043: GetErrInfo = "医院的统内比例为零，不允许住院！"
                Case 1044: GetErrInfo = "账户余额与中心不符！"
                Case 1045: GetErrInfo = "账户与差额不同！"
                Case 1046: GetErrInfo = "医院门诊大病结算算法错误！"
                Case 1047: GetErrInfo = "医院门诊大病冲账结算算法错误！"
                Case 1048: GetErrInfo = "就诊种类错误！"
                Case 1049: GetErrInfo = "医院门诊冲账结算数据错误！"
                Case 1050: GetErrInfo = "冲账时不允许产生差额！"
                Case 1052: GetErrInfo = "住院病人已登记在院！"
                Case 1053: GetErrInfo = "该病人未登记不在院！"
                Case 1054: GetErrInfo = "该大病受限制，您已经超额！"
                Case 1058: GetErrInfo = "大病登记时大病编码不能为空！"
                Case 1059: GetErrInfo = "就医限定医院错误！"
                Case 1062: GetErrInfo = "转诊单00000E的患者年龄小于70！"
                Case 1063: GetErrInfo = "转诊单00000E的患者身份证号错误！"
                Case 1064: GetErrInfo = "出院的登记日期和入院日期错误！"
                Case 1301: GetErrInfo = "大病编码已存在！"
                Case 1302: GetErrInfo = "无此大病编码！"
                Case 1303: GetErrInfo = "有限制大病编码已存在！"
                Case 1304: GetErrInfo = "无此有限制大病编码 , 没用对应的慢病帐户！"
                Case 1305: GetErrInfo = "此转诊单号已存在！"
                Case 1306: GetErrInfo = "无此转诊单号！"
                Case 1307: GetErrInfo = "此限定医院已存在！"
                Case 1308: GetErrInfo = "无此限定医院！"
                Case 7001, 7002, 7003, 7004, 7005
                    GetErrInfo = "读卡器连接错误；系统使用的读卡动态连接库错误；IC卡错误！"
                Case 7006: GetErrInfo = "系统错误，调入动态连接库错误！"
                Case 7007: GetErrInfo = "写卡时校验卡产生错误 (可能卡被换了)！"
                Case -8001: GetErrInfo = "完整性校验错误！"
                Case -8002: GetErrInfo = "医院编号错误，系统配置错误！"
                Case -8003: GetErrInfo = "系统版本错误，需要更新客户的程序！"
                Case -8004: GetErrInfo = "系统日期错误，需要更改客户端日期！"
                Case 1401: GetErrInfo = "不允许医保结算！"
                Case 1402: GetErrInfo = "子门诊不存在！"
                Case 1403: GetErrInfo = "家床数已达最大！"
                Case 1404: GetErrInfo = "参保类型1等于2 (医疗保险不可用)！"
                Case 1405: GetErrInfo = "不允许进行生育保险！"
                Case 1406: GetErrInfo = "不允许进行工伤保险！"
                Case 1407: GetErrInfo = "卡为在院状态不允许继续操作！"
                Case 1408: GetErrInfo = "无效卡！"
                Case 1409: GetErrInfo = "冲帐结算统筹累计出现负值！"
                Case 1410: GetErrInfo = "非法医院 , 医院不存在！"
                Case 1411: GetErrInfo = "此特殊转诊单已被禁用！"
                Case 1412: GetErrInfo = "非法日期格式 应是'YYYYMMDD'"
                Case 1413: GetErrInfo = "非法日期时间格式 应是'YYYYMMDDhhmmss'！"
                Case 1414: GetErrInfo = "住院登记的时候写卡序号不相等！"
                Case 1415: GetErrInfo = "这个医院不可以生育住院！"
                Case 1416: GetErrInfo = "这个医院不可以工伤住院！"
                Case 1417: GetErrInfo = "慢病冲账结算错误！"
                Case 1418: GetErrInfo = "工伤冲账结算错误！"
                Case 1419: GetErrInfo = "生育冲账结算错误！"
                '周海全调试 2003-12-17
                '加入如下错误代码
                Case 1424: GetErrInfo = "费用总额出现负数！"
                Case 1425: GetErrInfo = "结算方式与入院登记方式不一致！"
                Case 1427: GetErrInfo = "入院日期大于出院日期或者出院日期大于结算日期！"
                Case Else
                    GetErrInfo = "医保支持部分出现错误！"
                End Select
                '周海全调试 2003-12-17
                '同时加上错误编号，以方便查错。
                GetErrInfo = "错误编号：" & strCode & vbCr & "错误描述：" & GetErrInfo
        Case TYPE_重大校园卡
                Select Case Val(strCode)
                Case 0: GetErrInfo = " 成功"
                Case -1: GetErrInfo = "打开串口失败"
                Case -2: GetErrInfo = "读写器连接失败"
                Case -3: GetErrInfo = "参数错误"
                Case -4: GetErrInfo = "超时错误"
                Case -5: GetErrInfo = "无卡"
                Case -6: GetErrInfo = "用户卡错误"
                Case -7: GetErrInfo = "读卡错误"
                Case -8: GetErrInfo = "写卡错误"
                Case -9: GetErrInfo = "充值失败"
                Case -10: GetErrInfo = "减值失败"
                Case -11: GetErrInfo = "创建Licence错误"
                Case -12: GetErrInfo = "Licence错误"
                Case -13: GetErrInfo = "系统卡错误"
                Case -14: GetErrInfo = "金额不足"
                Case -15: GetErrInfo = "扇区未启动"
                Case -16: GetErrInfo = "网络通讯错误"
                Case -17: GetErrInfo = "配置文件错误"
                Case -18: GetErrInfo = "应答错误"
                Case -19: GetErrInfo = "黑名单中的卡"
                Case -20: GetErrInfo = "卡已到期"
                Case -21: GetErrInfo = "数据库操作失败"
                Case -22: GetErrInfo = "联机交易失败"
                Case -23: GetErrInfo = "密码错误"
                Case -24: GetErrInfo = "读磁卡错误"
                Case -25: GetErrInfo = "超出日消费限额"
                Case -100: GetErrInfo = "无法识别的卡"
                Case Else
                    GetErrInfo = "校园卡支持部分出现错误！"
                End Select
                GetErrInfo = "错误编号：" & strCode & vbCr & "错误描述：" & GetErrInfo
               Case Else
    End Select
End Function

Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional blnMessage As Boolean = True) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
    If Err <> 0 Then
        If blnMessage = True Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            Else
                MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
            End If
        End If
        
        Err.Clear
        OraDataOpen = False
        Exit Function
    End If
    OraDataOpen = True
End Function

Public Sub GetUserInfo()
 '功能：获取登陆用户信息
    Dim rsUser As New ADODB.Recordset
    Dim strSql As String
    
    Set rsUser = New ADODB.Recordset
    rsUser.CursorLocation = adUseClient
    'rsUser.Open "Select A.ID,A.部门ID,A.编号,A.简码,A.姓名,B.用户名,C.名称 as 部门 from 人员表 A,上机人员表 B,部门表 C Where A.部门ID=C.ID And  B.人员ID=A.ID AND Upper(B.用户名)=Upper(User)", gcnOracle, adOpenKeyset
    
    strSql = "select P.*,D.编码 as 部门编码,D.名称 as 部门名称,M.部门ID,u.用户名 " & _
                " from 上机人员表 U,人员表 P,部门表 D,部门人员 M " & _
                " Where U.人员id = P.id And P.ID=M.人员ID and  M.缺省=1 and M.部门id = D.id and U.用户名=user"
    rsUser.Open strSql, gcnOracle, adOpenKeyset
    
    If rsUser.RecordCount <> 0 Then
        UserInfo.ID = rsUser!ID
        UserInfo.编号 = rsUser!编号
        UserInfo.部门ID = IIf(IsNull(rsUser!部门ID), 0, rsUser!部门ID)
        UserInfo.简码 = IIf(IsNull(rsUser!简码), "", rsUser!简码)
        UserInfo.姓名 = IIf(IsNull(rsUser!姓名), "", rsUser!姓名)
        UserInfo.部门 = rsUser!部门名称
        UserInfo.用户名 = rsUser!用户名
        UserInfo.站点 = rsUser!用户名
        
        '为了不改其它程序，重复增加了一个变量
        gstrUserName = UserInfo.姓名
    End If
End Sub

Public Function DateStr() As String
    Dim rsTmp As New ADODB.Recordset

    rsTmp.Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    DateStr = Format(rsTmp.Fields(0).Value, "yyyy-MM-dd HH:mm:ss")
End Function

Public Function TrimStr(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Function NextNo(intBillId As Integer) As Variant
'------------------------------------------------------------------------------------
'功能：根据特定规则产生新的号码,规则如下：
'   一、编号原则：
'   1   病人ID         数字    永远递增编号 自动补缺号
'   二、年度位确定原则:
'       以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码
'参数：
'   intBillId:由“号码控制表”指定的单据标识
'返回：
'------------------------------------------------------------------------------------
    Dim rsCtrl As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim vntNo As Variant        '获取的号码的中间变量
    Dim blnNext As Boolean      '获得的号码是否为递增号码，否则为补缺号码
    Dim intYear, strYear As String      '年度标志位
    
    Dim blnByDate As Boolean, curDate As Date
RESTART:
    Err = 0
    On Error GoTo errHand
    
    If intBillId = 1 Then
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSql = "select * from 号码控制表 where 项目序号=" & intBillId
                Call SQLTest(App.ProductName, "mdlInPatient", strSql) 'SQLTest
                .Open strSql, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            strSql = "select nvl(max(病人ID),0)+1 from 病人信息 where 病人ID>=" & vntNo & ""
            
            With rsTmp
                If .State = adStateOpen Then .Close
                Call SQLTest(App.ProductName, "mdlInsure", strSql) 'SQLTest
                .Open strSql, gcnOracle
                Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            On Error Resume Next
            .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            If Err <> 0 Then
                .CancelUpdate
                GoTo RESTART
            End If
            NextNo = vntNo
        End With
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    NextNo = Null
End Function

Public Function Get入院诊断(lng病人ID As Long, lng主页ID As Long, _
Optional ByVal bln允许空 As Boolean = True, Optional ByVal bln疾病编码 As Boolean = False) As String
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.描述信息 as 入院诊断,B.编码 疾病编码 " & _
             " From 诊断情况 A,疾病编码目录 B " & _
             " Where A.病人ID=" & lng病人ID & " And A.疾病ID=B.ID(+) And A.主页ID=" & lng主页ID & " And A.诊断类型=2"
    rsInNote.CursorLocation = adUseClient
    Call OpenRecordset(rsInNote, "医保接口", strTmp)
    
    If Not rsInNote.EOF Then
        Get入院诊断 = IIf(IsNull(rsInNote!入院诊断), "", rsInNote!入院诊断)
    End If
    If Not bln允许空 Then
        Get入院诊断 = Trim(Get入院诊断)
        If Get入院诊断 = "" Then Get入院诊断 = "无"
    End If
    If bln疾病编码 Then
        If Not rsInNote.EOF Then
            Get入院诊断 = Get入院诊断 & "|" & NVL(rsInNote!疾病编码)
        Else
            Get入院诊断 = Get入院诊断 & "|"
        End If
    End If
End Function

Public Function BuildPatiInfo(ByVal bytType As Byte, ByVal strInfo As String, ByVal lng病人ID As Long) As Long
'功能：建立病人帐户信息
'参数：bytType=0-门诊,1-住院
'      strInfo='0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
'      8中心;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(1,2,3);15退休证号;16年龄段;17灰度级
'      18帐户增加累计;19帐户支出累计;20进入统筹累计;21统筹报销累计;22住院次数累计;23就诊类别
'      24本次起付线;25起付线累计;26基本统筹限额
'返回：病人ID
    Const MAX_BOUND = 26 '要求传入的信息段数
    
    Dim rsPati As ADODB.Recordset, str单位编码 As String, lng年龄 As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String, curDate As Date
    Dim lng中心 As Long, array信息 As Variant
    Dim lngTemp As Long
    
    gcnOracle.BeginTrans
    On Error GoTo ErrHandle
    
    If Len(Trim(strInfo)) <> 0 Then
        curDate = zlDatabase.Currentdate
        
        '200308z012:保证传入的信息串够用
        If UBound(Split(strInfo, ";")) < MAX_BOUND Then
            strInfo = strInfo & String(MAX_BOUND - UBound(Split(strInfo, ";")), ";")
        End If
        array信息 = Split(strInfo, ";")
        
        '从第7项内容中取出单位编码
        If array信息(7) Like "*(*" Then
            str单位编码 = Split(array信息(7), "(")(UBound(Split(array信息(7), "(")))
            str单位编码 = Mid(str单位编码, 1, Len(str单位编码) - 1)
        End If
        '取年龄
        If IsDate(array信息(5)) Then
            lng年龄 = Int(curDate - CDate(array信息(5))) / 365
        End If
        
        lng中心 = Val(array信息(8))
        
        If lng病人ID > 0 Then
            '该病人已经存在
            gstrSQL = "Select nvl(病人ID,0) 病人ID from 保险帐户 where 医保号='" & CStr(array信息(1)) & "' and 中心=" & lng中心 & " and 险类=" & gintInsure
            Call OpenRecordset(rsTemp, "建立帐户")
            If rsTemp.EOF = False Then
                If rsTemp("病人ID") <> lng病人ID Then
                    If gintInsure = TYPE_成都市 Then
                        If MsgBox("已经存在相同医保号的另外一位病人，您需要将这两位病人合并吗？", vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then
                            gcnOracle.RollbackTrans
                            Exit Function
                        End If
                        '对这两个病人进行合并
                        lngTemp = MergePatient(lng病人ID, rsTemp!病人ID)
                        If lngTemp = 0 Then
                            gcnOracle.RollbackTrans
                            Exit Function
                        End If
                        lng病人ID = lngTemp
                    Else
                        MsgBox "已经存在相同医保号的另外一位病人，请您在病人管理中将这两位病人合并", vbInformation, gstrSysName
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                End If
            End If
        End If
        
        '帐户唯一：险类,中心,医保号
        strSql = "Select A.*,B.医保号 From 病人信息 A," & _
            " (Select * From 保险帐户" & _
            " Where 险类=" & gintInsure & _
            " And 医保号='" & CStr(array信息(1)) & "'" & _
            " And 中心=" & lng中心 & ") B" & _
            " Where " & IIf(lng病人ID = 0, "A.病人ID=B.病人ID", "A.病人ID=B.病人ID(+) and A.病人ID=" & lng病人ID) '可能病人ID已经确定
        Set rsPati = New ADODB.Recordset
        rsPati.CursorLocation = adUseClient
        Call OpenRecordset(rsPati, "医保接口", strSql)
        
        If rsPati.EOF Then
            '无保险帐户则认为没有病人信息
            If lng病人ID = 0 Then lng病人ID = NextNo(1)
            strSql = "zl_病人信息_Insert(" & lng病人ID & ",NULL,NULL,'社会基本医疗保险'," & _
                "'" & array信息(3) & "','" & array信息(4) & "'," & IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
                "To_Date('" & Format(array信息(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "NULL,'" & array信息(6) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,'" & array信息(7) & "',NULL,NULL,NULL," & _
                "NULL,NULL,NULL," & gintInsure & "," & _
                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call SQLTest(App.ProductName, "医保接口", strSql)
            gcnOracle.Execute strSql, , adCmdStoredProc
            Call SQLTest
        Else
            '有病人信息和保险帐户信息
            If rsPati("姓名") <> array信息(3) Then
                If MsgBox("病人原有登记的姓名是 " & rsPati("姓名") & " ，与刷卡得到的姓名 " & array信息(3) & " 不符，" & vbCrLf & _
                          "继续会更新病人原有的登记信息，是否确定？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            End If
            If lng病人ID = 0 Then lng病人ID = rsPati!病人ID
            strSql = "zl_病人信息_Update(" & _
                lng病人ID & "," & IIf(IsNull(rsPati!门诊号), "NULL", rsPati!门诊号) & "," & _
                IIf(IsNull(rsPati!住院号), "NULL", rsPati!住院号) & ",'" & IIf(IsNull(rsPati!费别), "", rsPati!费别) & "'," & _
                "'" & IIf(IsNull(rsPati!医疗付款方式), "", rsPati!医疗付款方式) & "'," & _
                "'" & array信息(3) & "','" & array信息(4) & "'," & IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
                "To_Date('" & Format(array信息(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "'" & IIf(IsNull(rsPati!出生地点), "", rsPati!出生地点) & "','" & array信息(6) & "'," & _
                "'" & IIf(IsNull(rsPati!身份), "", rsPati!身份) & "','" & IIf(IsNull(rsPati!职业), "", rsPati!职业) & "'," & _
                "'" & IIf(IsNull(rsPati!民族), "", rsPati!民族) & "','" & IIf(IsNull(rsPati!国籍), "", rsPati!国籍) & "'," & _
                "'" & IIf(IsNull(rsPati!学历), "", rsPati!学历) & "','" & IIf(IsNull(rsPati!婚姻状况), "", rsPati!婚姻状况) & "'," & _
                "'" & IIf(IsNull(rsPati!家庭地址), "", rsPati!家庭地址) & "','" & IIf(IsNull(rsPati!家庭电话), "", rsPati!家庭电话) & "'," & _
                "'" & IIf(IsNull(rsPati!户口邮编), "", rsPati!户口邮编) & "','" & IIf(IsNull(rsPati!联系人姓名), "", rsPati!联系人姓名) & "'," & _
                "'" & IIf(IsNull(rsPati!联系人关系), "", rsPati!联系人关系) & "','" & IIf(IsNull(rsPati!联系人地址), "", rsPati!联系人地址) & "'," & _
                "'" & IIf(IsNull(rsPati!联系人电话), "", rsPati!联系人电话) & "'," & IIf(IsNull(rsPati!合同单位ID), "NULL", rsPati!合同单位ID) & "," & _
                "'" & array信息(7) & "','" & IIf(IsNull(rsPati!单位电话), "", rsPati!单位电话) & "'," & _
                "'" & IIf(IsNull(rsPati!单位邮编), "", rsPati!单位邮编) & "','" & IIf(IsNull(rsPati!单位开户行), "", rsPati!单位开户行) & "'," & _
                "'" & IIf(IsNull(rsPati!单位帐号), "", rsPati!单位帐号) & "','" & IIf(IsNull(rsPati!担保人), "", rsPati!担保人) & "'," & _
                "" & IIf(IsNull(rsPati!担保额), "NULL", rsPati!担保额) & "," & gintInsure & ")"
            Call SQLTest(App.ProductName, "医保接口", strSql)
            gcnOracle.Execute strSql, , adCmdStoredProc
            Call SQLTest
        End If
        
        '插入或更新保险帐户信息(自动)
        strSql = "zl_保险帐户_insert(" & lng病人ID & "," & gintInsure & "," & _
            lng中心 & "," & _
            "'" & IIf(array信息(0) = "-1", array信息(1), array信息(0)) & "'," & _
            "'" & array信息(1) & "'," & _
            "'" & array信息(2) & "'," & _
            "'" & array信息(9) & "'," & _
            "'" & array信息(15) & "'," & _
            "'" & array信息(10) & "'," & _
            "'" & str单位编码 & "'," & _
            Val(array信息(11)) & "," & _
            Val(array信息(12)) & "," & _
            IIf(Val(array信息(13)) = 0, "NULL", Val(array信息(13))) & "," & _
            IIf(Val(array信息(14)) = 0, 1, Val(array信息(14))) & "," & _
            IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
            "'" & array信息(17) & "'," & _
            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call SQLTest(App.ProductName, "医保接口", strSql)
        gcnOracle.Execute strSql, , adCmdStoredProc
        Call SQLTest
        
        '插入或更新帐户年度信息(自动)
        '200308z012:成都:保存"24本次起付线=zyjs,25起付线累计=tcbxbl,26基本统筹限额=zyxe"
        strSql = "zl_帐户年度信息_Insert(" & lng病人ID & "," & gintInsure & "," & Year(curDate) & "," & _
            Val(array信息(18)) & "," & Val(array信息(19)) & "," & _
            Val(array信息(20)) & "," & Val(array信息(21)) & "," & _
            Val(array信息(22)) & "," & Val(array信息(24)) & "," & Val(array信息(25)) & "," & Val(array信息(26)) & ")"
        Call SQLTest(App.ProductName, "医保接口", strSql)
        gcnOracle.Execute strSql, , adCmdStoredProc
        Call SQLTest
    End If
    gcnOracle.CommitTrans
    BuildPatiInfo = lng病人ID
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
End Function

Public Function GetTextFromCombo(cmbTemp As ComboBox, ByVal blnAfter As Boolean, Optional strSplit As String = ".") As String
'参数：cmbTemp  准备获取数据的ComboBox控件
'      blnAfter 表示在.之前或之后取值
    Dim lngPos As Long
    
    lngPos = InStr(cmbTemp.Text, strSplit)
    If lngPos = 0 Then
        '直接返回整个字符串
        GetTextFromCombo = "'" & cmbTemp.Text & "'"
    Else
        If blnAfter = False Then
            '圆点之前
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, 1, lngPos - 1) & "'"
        Else
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, lngPos + 1) & "'"
        End If
    End If
End Function

Public Sub SetComboByText(cmbTemp As ComboBox, ByVal strText As String, ByVal blnAfter As Boolean, Optional strSplit As String = ".")
'参数：cmbTemp  准备设置的ComboBox控件
'      blnAfter 表示在.之前或之后取值
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cmbTemp.ListCount - 1
        strTemp = cmbTemp.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            '直接返回整个字符串
            If strText = strTemp Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                '圆点之前
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '已经找到
        cmbTemp.ListIndex = lngCount
    Else
        cmbTemp.ListIndex = -1
        If blnAfter = True Then
            cmbTemp.AddItem strText
        End If
    End If
End Sub

Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'功能：按数据库规则得到字符串的子集，也就是汉字按两个字符算，而字母仍是一个
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    '去掉可能出现的半个字符
    MidUni = Replace(MidUni, Chr(0), "")
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
'功能：将文本按Varchar2的长度计算方法进行截断
    Dim strText As String
    
    strText = IIf(IsNull(varText), "", varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    '去掉可能出现的半个字符
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Public Function GetComputer(frmParant As Form, Optional ByVal strCaption As String = "选择计算机") As String
'功能：返回计算机名
   Dim BI As BrowseInfo
   Dim pidl As Long
   Dim sPath As String
   Dim pos As Integer
   
  'obtain the pidl to the special folder 'network'
   If SHGetSpecialFolderLocation(frmParant.hwnd, CSIDL_NETWORK, pidl) = 0 Then
     'fill in the required members, limiting the
     'Browse to the network by specifying the
     'returned pidl as pidlRoot
      With BI
         .hwndOwner = frmParant.hwnd
         .pIDLRoot = pidl
         .pszDisplayName = Space$(MAX_PATH)
         .lpszTitle = lstrcat(strCaption, "")
         .ulFlags = BIF_BROWSEFORCOMPUTER
      End With
         
     'show the browse dialog. We don't need
     'a pidl, so it can be used in the If..then directly.
      If SHBrowseForFolder(BI) <> 0 Then
               
         'a server was selected. Although a valid pidl
         'is returned, SHGetPathFromIDList only return
         'paths to valid file system objects, of which
         'a networked machine is not. However, the
         'BROWSEINFO displayname member does contain
         'the selected item, which we return
          GetComputer = TrimStr(BI.pszDisplayName)
            
      End If  'If SHBrowseForFolder
      
      Call CoTaskMemFree(pidl)
               
   End If  'If SHGetSpecialFolderLocation
   
End Function

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSql As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSql = "", gstrSQL, strSql))
    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Sub ExecuteProcedure(ByVal strCaption As String)
'功能：执行SQL语句
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Public Sub CenterTableCaption(mshTemp As Object)
'功能：设置表格的列头居中对齐
    With mshTemp
        .Col = 0
        .Row = .FixedRows - 1
        .ColSel = .Cols - 1
        .RowSel = .Row
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = .FixedRows: .Col = .FixedCols
    End With
End Sub

Public Function Get住院次数(lng病人ID As Long) As Integer
'功能：获取指定病人本年度住院次数
'说明：跨年住院的情况两年都各算一次住院。
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Count(*) as 次数 From 病案主页 Where Nvl(出院日期,Sysdate)=To_Date(To_Char(Sysdate,'YYYY')||'-01-01','YYYY-MM-DD') And 病人ID=" & lng病人ID
    rsTmp.CursorLocation = adUseClient
    Call OpenRecordset(rsTmp, "医保接口")
    
    If Not rsTmp.EOF Then Get住院次数 = IIf(IsNull(rsTmp!次数), 0, rsTmp!次数)
End Function

Public Function Get帐户信息(ByVal lng病人ID As Long, ByVal str年度 As String, int住院次数累计 As Integer, _
    cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, _
    cur统筹报销累计 As Currency, Optional cur本次起付线 As Currency, Optional cur起付线累计 As Currency, _
    Optional cur基本统筹限额 As Currency) As Boolean
'功能：得到帐户年度信息
'200308z012:新增几个返回参数
    Dim rsTemp As New ADODB.Recordset
    
    cur帐户增加累计 = 0
    cur帐户支出累计 = 0
    cur进入统筹累计 = 0
    cur统筹报销累计 = 0
    int住院次数累计 = 0
    cur本次起付线 = 0
    cur起付线累计 = 0
    cur基本统筹限额 = 0
    
    '帐户年度信息
    gstrSQL = "Select * From 帐户年度信息 Where 病人ID=" & lng病人ID & " And 险类=" & gintInsure & " And 年度=" & str年度
    Call OpenRecordset(rsTemp, "医保接口")
    
    If rsTemp.EOF = False Then
        cur帐户增加累计 = IIf(IsNull(rsTemp("帐户增加累计")), 0, rsTemp("帐户增加累计"))
        cur帐户支出累计 = IIf(IsNull(rsTemp("帐户支出累计")), 0, rsTemp("帐户支出累计"))
        cur进入统筹累计 = IIf(IsNull(rsTemp("进入统筹累计")), 0, rsTemp("进入统筹累计"))
        cur统筹报销累计 = IIf(IsNull(rsTemp("统筹报销累计")), 0, rsTemp("统筹报销累计"))
        int住院次数累计 = IIf(IsNull(rsTemp("住院次数累计")), 0, rsTemp("住院次数累计"))
        cur本次起付线 = IIf(IsNull(rsTemp("本次起付线")), 0, rsTemp("本次起付线"))
        cur起付线累计 = IIf(IsNull(rsTemp("起付线累计")), 0, rsTemp("起付线累计"))
        cur基本统筹限额 = IIf(IsNull(rsTemp("基本统筹限额")), 0, rsTemp("基本统筹限额"))
    End If

End Function

Public Function 门诊虚拟结算(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim cls医保 As New clsInsure
    Dim dbl全自费 As Currency, dbl首先自付 As Currency, dbl进入统筹 As Currency
    Dim dbl个人帐户 As Double
    Dim lng病人ID As Long
    Dim rs特准项目 As New ADODB.Recordset
    
    If rs明细.RecordCount > 0 Then
        rs明细.MoveFirst
        lng病人ID = rs明细("病人ID")
    End If
    
    gstrSQL = "select A.收费细目ID from 保险特准项目 A,保险帐户 B " & _
            "where A.病种ID=B.病种ID and B.病人ID=" & lng病人ID & " and 险类=" & gintInsure
    Call OpenRecordset(rs特准项目, "虚拟结算")
    
    Do Until rs明细.EOF
        rs特准项目.Filter = "收费细目ID = " & rs明细("收费细目ID")
        
        If rs明细("是否医保") = 1 Or rs特准项目.EOF = False Then
            '如果是特准项目，强行进入统筹
            dbl进入统筹 = dbl进入统筹 + rs明细("统筹金额")
            dbl首先自付 = dbl首先自付 + rs明细("实收金额") - rs明细("统筹金额")
        Else
            dbl全自费 = dbl全自费 + rs明细("实收金额")
        End If
            
        rs明细.MoveNext
    Loop
    If cls医保.GetCapability(support收费帐户全自费) = True Then
        dbl个人帐户 = dbl个人帐户 + dbl全自费
    End If
    
    If Is全额统筹(lng病人ID) = True Then
        '首先自付也是由医保基金支付
        str结算方式 = "个人帐户;" & dbl个人帐户 & ";0|医保基金;" & dbl进入统筹 + dbl首先自付 & ";0"
    Else
        If cls医保.GetCapability(support收费帐户首先自付) = True Then
            dbl个人帐户 = dbl个人帐户 + dbl首先自付
        End If
        
        str结算方式 = "个人帐户;" & dbl个人帐户 & ";0|医保基金;" & dbl进入统筹 & ";0"
    End If
    
    门诊虚拟结算 = True
End Function

Public Function Is全额统筹(ByVal 病人ID As Long) As Boolean
'功能：判断是否全额统筹病人(注意：传的病人ID可能非医保病人的)
    Dim rsTemp As New ADODB.Recordset
    
    If gintInsure = TYPE_自贡市 Then
        '对于自贡医保：只要病人是离休人员，那就是全额统筹
        gstrSQL = "select 在职 from 保险帐户 where 病人ID=" & 病人ID & " and 险类=" & TYPE_自贡市
        Call OpenRecordset(rsTemp, "医保接口")
        If rsTemp.EOF = False Then
            Is全额统筹 = IIf(rsTemp("在职") = 3, True, False)
        End If
    Else
        gstrSQL = _
            "Select Nvl(B.全额统筹,0) as 全额统筹" & _
            " From 保险帐户 A,保险年龄段 B" & _
            " Where A.险类 = B.险类 And Nvl(A.中心, 0) = Nvl(B.中心, 0)" & _
            " And Nvl(A.在职,0)=Nvl(B.在职,0)" & _
            " And B.下限<=Nvl(A.年龄段,0) And (A.年龄段<=B.上限 Or B.上限=0)" & _
            " And A.病人ID=" & 病人ID & " And A.险类=" & gintInsure
        Set rsTemp = New ADODB.Recordset
        rsTemp.CursorLocation = adUseClient
        Call OpenRecordset(rsTemp, "医保接口")
        
        If Not rsTemp.EOF Then Is全额统筹 = (rsTemp!全额统筹 = 1)
    End If
End Function

Public Function AddDate(ByVal strOrin As String, Optional ByVal bln时 As Boolean = False) As String
'功能：为不全的日期信息补充完整
    Dim strTemp As String
    Dim intPos As Integer
    
    strTemp = Trim(strOrin)
    
    If strTemp = "" Then
        AddDate = ""
        Exit Function
    End If
    
    intPos = InStr(strTemp, "-")
    If intPos = 0 Then
        intPos = InStr(strTemp, ".")
        If intPos <> 0 Then
            '使用 . 隔
            strTemp = Replace(strTemp, ".", "-")
        End If
    End If
    
    If intPos = 0 Then
        '没有"-",手工加上
        intPos = Len(strTemp)
        If intPos <= 8 Then
            If intPos = 8 Then
                strTemp = Mid(strTemp, 1, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7, 2)
            ElseIf intPos > 4 Then
                strTemp = Left(strTemp, intPos - 4) & "-" & Mid(Right(strTemp, 4), 1, 2) & "-" & Right(strTemp, 2)
            ElseIf intPos > 2 Then
                strTemp = Format(Date, "yyyy") & "-" & Left(strTemp, intPos - 2) & "-" & Right(strTemp, 2)
            Else
                strTemp = Format(Date, "yyyy") & "-" & Format(Date, "MM") & "-" & strTemp
            End If
        End If
    Else
        If bln时 = False Then
            If IsDate(strTemp) Then
                strTemp = Format(CDate(strTemp), "yyyy-MM-dd")
            End If
        Else
            '处理小时
            If InStr(strTemp, " ") > 0 Then
                '输入了小时
                If IsDate(strTemp & ":00") Then
                    strTemp = Format(CDate(strTemp & ":00"), "yyyy-MM-dd HH:ss")
                End If
            Else
                If IsDate(strTemp) Then
                    strTemp = Format(CDate(strTemp), "yyyy-MM-dd HH:ss")
                End If
            End If
        End If
    End If
    
    AddDate = strTemp
End Function

Public Function Insert虚拟结算数据(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str结算方式 As String) As Boolean
'功能：将虚拟结算的数据保存起来
'参数：结算方式  "报销方式;金额;是否允许修改|...."
    Dim cnTemp As New ADODB.Connection
    Dim strDate As String
    Dim lngCount As Long, arr结算方式 As Variant, arr金额 As Variant
    
    cnTemp.Open gcnOracle.ConnectionString '为了防止一个连接串多次进放事务
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    cnTemp.BeginTrans
    On Error GoTo ErrHandle
    
    gstrSQL = "zl_保险模拟结算_Clear(" & lng病人ID & "," & lng主页ID & ")"
    cnTemp.Execute gstrSQL, , adCmdStoredProc
    
    arr结算方式 = Split(str结算方式, "|")
    For lngCount = 0 To UBound(arr结算方式)
        If arr结算方式(lngCount) <> "" Then
            arr金额 = Split(arr结算方式(lngCount), ";")
            If UBound(arr金额) > 1 Then
                If Val(arr金额(1)) <> 0 Then
                    gstrSQL = "zl_保险模拟结算_Insert(" & lng病人ID & "," & IIf(lng主页ID = 0, "null", lng主页ID) & _
                        ",'" & arr金额(0) & "'," & Val(arr金额(1)) & "," & strDate & ")"
                    cnTemp.Execute gstrSQL, , adCmdStoredProc
                End If
            End If
        End If
    Next
    
    cnTemp.CommitTrans
    Insert虚拟结算数据 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    cnTemp.RollbackTrans
End Function

Public Function Clear虚拟结算数据(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：在结帐之后，将虚拟结算的数据清除
    
    gstrSQL = "zl_保险模拟结算_Clear(" & lng病人ID & "," & lng主页ID & ")"
    Call ExecuteProcedure("虚拟结算")
    
    Clear虚拟结算数据 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get出生日期(ByVal str身份证 As String, ByVal lng年龄 As Long) As String
'功能：根据身份证号码或年龄得到出生日期
    Dim strDate As String
    If Len(str身份证) = 15 Then
        '老式的身份证号
        strDate = AddDate(Mid(str身份证, 7, 6))
        strDate = "19" & strDate
    ElseIf Len(str身份证) = 18 Then
        '新式的身份证号
        strDate = AddDate(Mid(str身份证, 7, 8))
    Else
        '没有身份证号
        strDate = Format(DateAdd("yyyy", lng年龄 * -1, Date), "yyyy-MM-dd")
    End If
    
    If IsDate(strDate) = True Then
        Get出生日期 = Format(CDate(strDate), "yyyy-MM-dd")
    End If
End Function

Public Function GetOracleFormat(ByVal dat日期 As Date)
    GetOracleFormat = "To_Date('" & Format(dat日期, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Function NVL(ByVal varValue As Variant, Optional varDefalut As Variant = "") As Variant
'功能：模仿Oracle的函数
    NVL = IIf(IsNull(varValue) = True, varDefalut, varValue)
End Function

Public Sub RemoveSelect(lvw As ListView)
'功能：删除当前选中项
    Dim lngIndex  As Long
    
    With lvw
        If .SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = .SelectedItem.Index
        .ListItems.Remove lngIndex
        
        If .ListItems.Count > 0 Then
            '如果仍有列表，则进行下一个选择
            lngIndex = IIf(.ListItems.Count > lngIndex, lngIndex, .ListItems.Count)
            .ListItems(lngIndex).Selected = True
            .ListItems(lngIndex).EnsureVisible
        End If
    End With

End Sub

Public Function Can住院结算冲销(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：判断病人的住院结算数据是否允许作废。判断标准是检查病人有新的住院记录，如果有，就不能交冲销
'参数：lng病人ID     病人ID
'      lng主页ID     该结帐记录所在的住院次数
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle

    gstrSQL = "SELECT COUNT(*) as 住院次数 FROM 病案主页 WHERE 病人ID=" & lng病人ID & " AND 主页ID>" & lng主页ID
    Call OpenRecordset(rsTemp, "结帐作废")
    If rsTemp("住院次数") > 0 Then
        MsgBox "该病人已经有新的住院记录，不能作废以前住院的结帐数据。", vbInformation, gstrSysName
        Exit Function
    End If

    Can住院结算冲销 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 医保病人已经出院(ByVal lng病人ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Nvl(当前状态,0) 状态 From 保险帐户 Where 病人ID=" & lng病人ID
    Call OpenRecordset(rsTmp, "判断医保病人是否出院")
    
    医保病人已经出院 = (rsTmp!状态 = 0)
End Function

Public Function 存在未结费用(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim rs费用 As New ADODB.Recordset
    '检查该次住院是否还有费用未结算
    gstrSQL = "Select nvl(费用余额,0) as 金额  from 病人余额 where 病人ID=" & lng病人ID & " and 性质=1"
    Call OpenRecordset(rs费用, "是否存在未结费用")
    If rs费用.EOF = True Then
        存在未结费用 = False
    Else
        存在未结费用 = (rs费用("金额") <> 0)
    End If
End Function

Public Function 获取入出院诊断(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
Optional ByVal bln入院诊断 As Boolean = True, Optional ByVal bln允许空 As Boolean = True, _
Optional ByVal bln疾病编码 As Boolean = False) As String
    
    '1-门诊诊断;2-入院诊断;3-出院诊断
    Dim rs诊断 As New ADODB.Recordset
    If bln疾病编码 = False Then
        gstrSQL = " Select A.描述信息" & _
                  " From 诊断情况 A" & _
                  " Where A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & _
                  " And A.诊断类型=" & IIf(bln入院诊断, "1", "3") & " And 诊断次序=1"
    Else
        gstrSQL = " Select A.描述信息,B.编码 疾病编码" & _
                  " From 诊断情况 A,疾病编码目录 B" & _
                  " Where A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & _
                  " And A.疾病ID=B.ID(+) And A.诊断类型=" & IIf(bln入院诊断, "1", "3")
    End If
    Call OpenRecordset(rs诊断, "获取入出院诊断")
    
    获取入出院诊断 = ""
    If Not rs诊断.EOF Then
        获取入出院诊断 = IIf(IsNull(rs诊断!描述信息), "", rs诊断!描述信息)
    End If
    
    获取入出院诊断 = Trim(获取入出院诊断)
    If Not bln允许空 And 获取入出院诊断 = "" Then
        获取入出院诊断 = "无"
    End If
    If bln疾病编码 Then
        If Not rs诊断.EOF Then
            获取入出院诊断 = 获取入出院诊断 & "|" & NVL(rs诊断!疾病编码, " ")
        Else
            获取入出院诊断 = 获取入出院诊断 & "| "
        End If
    End If
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '获取注册表后，马上清零
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "公共全局", "公共", 0)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", 0)
    blnValid = (intAtom <> 0)
    
    '如果存在，则对串进行解析
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '如果为空，则表示非法
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '判断时间间隔是否大于1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '如果相等，则通过
                    Else
                        '不等，表示存在进位，则分应该为零
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function 存在中心(ByVal int险类 As Integer) As Boolean
    Dim rs中心 As New ADODB.Recordset
    
    存在中心 = False
    gstrSQL = "Select Nvl(具有中心,0) 中心 From 保险类别 Where 序号=" & int险类
    Call OpenRecordset(rs中心, "是否有中心")
    If Not rs中心.EOF Then
        存在中心 = (rs中心!中心 = 1)
    End If
End Function

Private Function GetPatiInfo(lngID As Long) As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select * From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID(+) And A.病人ID=" & lngID & " Order by 主页ID"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnOracle, adOpenKeyset
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
        
Private Function MergePatient(ByVal lngOld As Long, ByVal lngInsure As Long) As Long
    Dim i As Integer, j As Integer
    Dim lngNew As Long
    Dim curDate As Date
    Dim strSql As String
    Dim rsPatiS As New ADODB.Recordset
    Dim rsPatiO As New ADODB.Recordset
    Set rsPatiS = GetPatiInfo(lngOld)
    Set rsPatiO = GetPatiInfo(lngInsure)
        
    'AB都住过院
    If Not IsNull(rsPatiS!主页ID) And Not IsNull(rsPatiO!主页ID) Then
        '1.先住院的在院,不允许(先后住院可以为：出院-出院,出院-在院；不允许：在院-出院,在院-在院)
        '因为除病人合并外,程序不额外处理自动出院或撤消出院
        rsPatiS.MoveLast
        rsPatiO.MoveLast
        If rsPatiS!入院时间 <= rsPatiO!入院时间 Then
            If IsNull(rsPatiS!出院时间) Then
                MsgBox "病人:" & rsPatiS!住院号 & " 最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If IsNull(rsPatiO!出院时间) Then
                MsgBox "病人:" & rsPatiO!住院号 & " 最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '2.时间交叉提示是否继续
        curDate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
                If Not (rsPatiO!入院日期 >= IIf(IsNull(rsPatiS!出院日期), curDate, rsPatiS!出院日期) Or _
                    IIf(IsNull(rsPatiO!出院日期), curDate, rsPatiO!出院日期) <= rsPatiS!入院日期) Then
                    If MsgBox("发现病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]第 " & rsPatiS!主页ID & " 次住院的期间" & Format(rsPatiS!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiS!出院日期), curDate, rsPatiS!出院日期), "yyyy-MM-dd") & vbCrLf & _
                        "与病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]的第 " & rsPatiO!主页ID & " 次住院的期间" & Format(rsPatiO!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiO!出院日期), curDate, rsPatiO!出院日期), "yyyy-MM-dd") & _
                        vbCrLf & "互相交叉，应该不是同一个病人，确实要合并吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
                rsPatiO.MoveNext
            Next
            rsPatiS.MoveNext
        Next
        
        lngNew = NextNo(1)
    End If
    
    strSql = "zl_病人信息_MERGE(" & lngOld & "," & lngInsure & IIf(lngNew <> 0, "," & lngNew, "") & ")"
    Screen.MousePointer = 11
    DoEvents
    
    gcnOracle.Execute strSql, , adCmdStoredProc
    Screen.MousePointer = 0
    
    If lngNew <> 0 Then
        If glngSys Like "8??" Then
            MsgBox "客户合并成功,合并后的客户ID为""" & lngNew & """！", vbInformation, gstrSysName
        Else
            MsgBox "病人合并成功,合并后的病人ID为""" & lngNew & """！", vbInformation, gstrSysName
        End If
        MergePatient = lngNew
    Else
        If glngSys Like "8??" Then
            MsgBox "客户合并成功！", vbInformation, gstrSysName
        Else
            MsgBox "病人合并成功！", vbInformation, gstrSysName
        End If
        MergePatient = lngInsure
    End If
End Function

Public Sub DebugTool(ByVal strInfo As String)
    Dim intDebug As Integer
    '判断是否是调试状态，是则显示提示框
    intDebug = GetSetting("ZLSOFT", "医保", "调试", 0)
    If intDebug = 0 Then Exit Sub
    MsgBox strInfo
End Sub

Public Function SystemImes() As Variant
'功能：将系统中文输入法名称返回到一个字符串数组中
'返回：如果不存在中文输入法,则返回空串
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function OpenIme(Optional strIme As String) As Boolean
'功能:按名称打开中文输入法,不指定名称时关闭中文输入法。支持部分名称。
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If strIme = "不自动开启" Then OpenIme = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function
