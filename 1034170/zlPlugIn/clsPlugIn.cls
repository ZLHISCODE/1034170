VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------------
'功能：外挂功能接口部件，用于一些工作站集成外挂调用功能，并传入相应的业务参数
'编制人：周韬
'编制时间：2010-07-22
'说明：
'  2010-07-23:集成到了门诊/住院医嘱虚拟模块中
'  2010-11-25:增加了对自动调用功能的支持,以及对医技站前提ID的传入
'---------------------------------------------------------------------------

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      病历调用 lngModual=1070，无int场合参数
    Set gcnOracle = cnOracle
    
    Call HandlePlugIn(1, lngSys, lngModual, cnOracle, int场合)
    
End Sub

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1, Optional ByVal strReserve As String) As String
'功能：获取外挂功能列表；该接口一般在初始化之后调用一次，调用时机和频度可能因不同的主程序而不同
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号，病历调用时 lngModual=1070，无int场合参数
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=保留参数,用于扩展使用
'      医技站中的调用时,传入为当前执行的前提医嘱ID
'返回：外挂功能列表，格式为"功能名1,功能名2,功能名3,..."
'      如果某个功能需要在病人切换(或刷新)时自动调用，则在功能名前面加上"Auto:"前缀，这种功能只允许有一个。

'特殊说明：当插件在人员管理模块调用时只传入lngSys和lngModual，并且其中的lngModual＝0
'          自动调用功能在人员管理模块调用时不会自动调用

    Dim strTmp As String
    
    Call HandlePlugIn(2, lngSys, lngModual, , int场合, strReserve, strTmp)
    
    GetFuncNames = "功能名1,功能名2,功能名3,功能名4,Auto:功能名5" & IIf(strTmp = "", "", "," & strTmp)
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, _
    ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String, Optional ByVal int场合 As Integer = -1)
'功能：执行某个外挂功能
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号，病历调用时 lngModual＝1070，无int场合参数
'      strFuncName=要执行的功能名，该功能名之前由GetFuncNames函数返回
'      lngPatiID=当前病人ID
'      varRecId=数字或者字符串；对门诊病人，为当前挂号单号或者挂号ID；对住院病人，为当前住院主页ID
'      varKeyId=数字或者字符串；当前的关键业务数据唯一标识ID，如医嘱ID
'           病历调用时 varKeyId=电子病历记录.id，在新增编辑未保存前其值=0,保存后或修改、审订时执行有值；
'           临床路径 varKeyId 病人路径id
'      strReserve=保留参数,用于扩展使用 当是病历调用时 strReserve＝ 病历文件列表.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

'特殊说明：当插件在人员管理模块调用时只传入lngSys、lngModual、strFuncName、lngPatiID、varRecId、varKeyId，
'          lngPatiID就是人员id，如果lngPatiID＝0表示界面上未选中人员；其中的varRecId和varKeyId等值都是传的0

    '注意：在采用分离插件后，功能的执行时，应该先对功能明进行判断。
    '先检查功能是否在本插件（zlPlugIn）中，然后再去执行扩展插件
    
    If InStr("功能名1,功能名2,功能名3,功能名4,功能名5", strFuncName) > 0 Then
        MsgBox strFuncName & "执行成功！"
    Else
        Call HandlePlugIn(3, lngSys, lngModual, , int场合, strReserve, strFuncName, lngPatiID, varRecId, varKeyId)
    End If
End Sub

Public Function QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal str卡号 As String, ByRef rsInfo As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回病人信息集。病人信息集为动态记录集，在接口中预先写好构建及返回字段模板；如果返回字段中"病人ID"<>0，则表示接口返回前已新建档案，挂号程序直接提取病人ID对应的档案信息，而不直接读取其他字段；
    '入参:lngModule-调用的模块号
    '出参:rsInfo-返回的病人信息集(直接更新该表,字段为:病人ID,姓名,性别,年龄,出生日期,出生地点,身份证号,其他证件,身份,职业,家庭地址,家庭电话,家庭邮编,工作单位,单位邮编,医保号,医疗付款方式,费别,国籍,民族,婚姻状况,区域)
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:33:19
    '说明:准格尔旗区域卫生系统接口(读取病人信息)
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '例:
    '    With rsInfo
    '        .AddNew
    '            !病人ID = Null
    '            !姓名 = "区域卫生" & Int(Rnd * 10)
    '            !性别 = "女"
    '            !年龄 = "44"
    '            !出生日期 = Now - 3000
    '            !出生地点 = "重庆市北碚区万寿桥村1组"
    '            !身份证号 = "80223197405197890"
    '            !其他证件 = Null
    '            !身份 = "工人"
    '            !职业 = "IT"
    '            !家庭地址 = "渝北区工业园区"
    '            !家庭电话 = "67136410"
    '            !家庭邮编 = "5123"
    '            !工作单位 = "中联信息产业"
    '            !单位邮编 = "423"
    '            !医保号 = "42143Yb"
    '            !医疗付款方式 = "生育保险"
    '            !费别 = "普通"
    '            !国籍 = "中国"
    '            !民族 = "汉族"
    '            !婚姻状况 = "已婚"
    '            !区域 = "渝北区"
    '        .Update
    '    End With
    QueryPatiInfo = True
End Function
 Public Function CommitPatiInfo(ByVal str卡号 As String, ByRef rsInfo As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:提交发卡信息
    '入参:str卡号-卡号
    '        rsInfo-返回的病人信息集(直接更新该表,字段为:病人ID,姓名,性别,年龄,出生日期,出生地点,身份证号,其他证件,身份,职业,家庭地址,家庭电话,家庭邮编,工作单位,单位邮编,医保号,医疗付款方式,费别,国籍,民族,婚姻状况,区域)
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:33:19
    '说明:准格尔旗区域卫生系统接口(提交发卡信息)
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CommitPatiInfo = True
End Function
Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口终止
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    Call HandlePlugIn(4, lngSys, lngModual, , int场合)
End Sub

Public Function ClinicReceive(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'功能：病人接诊操作后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）

    ClinicReceive = True
End Function

Public Function ClinicFinish(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'功能：病人完成接诊后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）

    ClinicFinish = True
End Function

Public Function DiagnosisEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lng原诊断id As Long, Optional ByVal int场合 As Integer = -1) As Boolean
'功能：门诊病案中输入主要诊断后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）
'      lngDiagID=诊断ID或疾病ID(根据系统参数决定),strDiagDesc=诊断描述
'      lng原诊断id=0 表示新添加的诊断， 不为0表示修改诊断，lng原诊断id的值就是修改前的 诊断ID或疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
      
    DiagnosisEnter = True
End Function

Public Function DiagnosisOtherEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lng原诊断id As Long, Optional ByVal int场合 As Integer = -1) As Boolean
'功能：门诊病案中输入次要诊断后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）
'      lngDiagID=诊断ID或疾病ID(根据系统参数决定),strDiagDesc=诊断描述
'      lng原诊断id=0 表示新添加的诊断， 不为0表示修改诊断，lng原诊断id的值就是修改前的 诊断ID或疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    
    DiagnosisOtherEnter = True
End Function

Public Function DiagnosisDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional ByVal int场合 As Integer = -1) As Boolean
'功能：门诊病案中删除诊断后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）
'      lngDiagID=诊断ID或疾病ID(根据系统参数决定),strDiagDesc=诊断描述
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    
    DiagnosisDeleted = True
End Function

Public Function AdviceEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional ByVal int场合 As Integer = -1) As Boolean
'功能：每输入一条医嘱项目时调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      rsAdvice=当前录入的一条或一组医嘱记录（中药配方、成套方案、批量复制医嘱在输入时不调用）
'      字段为表“病人医嘱记录”的大部分字段，包括：ID, 相关id, 序号, 病人来源（1=门诊,2=住院), 病人id, 挂号单/主页ID,
'                姓名, 性别, 年龄, 婴儿, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id, 收费细目id, 天数, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 标本部位,
'                检查方法 , 执行标记, 执行频次, 频率次数, 频率间隔, 间隔单位, 执行时间方案, 计价特性, 执行科室id, 执行性质, 紧急标志, 可否分零,
'                开始执行时间,执行终止时间(住院调用才有，没填时为cdate(0)), 病人科室id, 开嘱科室id, 开嘱医生, 开嘱时间, 挂号单, 前提id, 摘要
'                ,除此之外，门诊调用时，还包含医嘱对应的诊断的疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceEnter = True
End Function

Public Function AdviceSave(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int场合 As Integer = -1) As Boolean
'功能：保存医嘱时(事务开始前)调用,返回False时不保存医嘱
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      rsAdvice=当前病人本次新开的所有医嘱记录，字段同AdviceEnter接口的该参数,除此之外，门诊调用时，还包含医嘱对应的诊断的疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceSave = True
End Function


Public Function AdviceSend(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngSendID As Long) As Boolean
'功能：发送医嘱完成后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达(门诊没有“发送”模块),1254=住院医嘱发送)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngSendID=发送号(病人医嘱发送.id)

    AdviceSend = True
End Function

Public Function AdviceRowChange(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int场合 As Integer = -1) As Boolean
'功能：医嘱编辑时切换医嘱行时调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      rsAdvice=当前病人本次新开的所有医嘱记录，字段同AdviceEnter接口的该参数,除此之外，门诊调用时，还包含医嘱对应的诊断的疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceRowChange = True
End Function

Public Function AdviceDeletBefor(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int场合 As Integer = -1) As Boolean
'功能：医嘱删除前调用（返回False则取消删除医嘱）
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceDeletBefor = True
End Function

Public Function AdviceDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int场合 As Integer = -1) As Boolean
'功能：医嘱删除后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceDeleted = True
End Function


Public Function AdviceRevoked(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int场合 As Integer = -1) As Boolean
'功能：医嘱作废后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceRevoked = True
End Function

Public Function AddEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'功能：新增电子病历时候调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(门诊病历 1250,住院病历 1251,护理病历 1255)
'      lngPatientID=病人ID,lngClinicID=就诊ID/
'      EMRFileID=病历文件列表.id
'      arrInput是数组：扩展参数用于扩展使用，通过下标访一直接问时可能会报两种错，9－下标越界/13－类型不匹配：均表示参数没有传入
'           所以访问参数前应先判断：If TypeName(arrInput(i)) <> "Error" Then 满足这个条件后才使用这个参数，否则就认为这个参数没有传入
'           arrInput 目前未传入任何参数
    AddEMRBefore = True
End Function

Public Function SignEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'功能：电子病历签名前调用，事务未开启时
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号，lngModual固定传递1070
'      lngPatientID=病人ID,lngClinicID=就诊ID/主页ID
'      EMRFileID=电子病历记录.id
'      arrInput是数组：扩展参数用于扩展使用，通过下标访一直接问时可能会报两种错，9－下标越界/13－类型不匹配：均表示参数没有传入
'           所以访问参数前应先判断：If TypeName(arrInput(i)) <> "Error" Then 满足这个条件后才使用这个参数，否则就认为这个参数没有传入
'           arrInput 目前未传入任何参数
'返回：true时，允许继续签名，否则不弹出签名窗口，病历内部不做提示。
    SignEMRBefore = True
End Function

Public Sub SignEMRAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant)
'功能：在签名窗口关闭且文档数据保存后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号，lngModual固定传递1070
'      lngPatientID=病人ID,lngClinicID=就诊ID/主页ID
'      EMRFileID=电子病历记录.id  在新增编辑未保存前其值=0，保存后签名、修改时签名、审签有值；
'      arrInput是数组：扩展参数用于扩展使用，通过下标访一直接问时可能会报两种错，9－下标越界/13－类型不匹配：均表示参数没有传入
'           所以访问参数前应先判断：If TypeName(arrInput(i)) <> "Error" Then 满足这个条件后才使用这个参数，否则就认为这个参数没有传入
'           arrInput(0)－String：签名人的姓名

End Sub

Public Sub LogInAfter()
'功能：登陆航导台之后

End Sub

Public Sub LogOutBefore()
'功能：关闭航导台或者注销当前
    
End Sub

Public Function PathImportBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径ID As Long, Optional ByVal lng版本号ID As Long, Optional ByVal int场合 As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng疾病ID As Long, Optional ByVal lng诊断ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'功能：路径导入前
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lng路径ID=临床路径目录.id
'      lng版本号ID=导入的版本号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      lngDiagnosisType 导入诊断类型 '诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
'      lng疾病ID=导入疾病（疾病编码目录.ID）、lng诊断ID=导入诊断(疾病诊断目录.ID)
'      strReserve=扩展参数

    PathImportBefore = True

End Function

Public Function PathImportAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径ID As Long, Optional ByVal lng版本号ID As Long, Optional ByVal int场合 As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng疾病ID As Long, Optional ByVal lng诊断ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'功能：路径导入前
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lng路径ID=临床路径目录.id
'      lng版本号ID=导入的版本号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=扩展参数

    PathImportAfter = True

End Function

Public Function PathEvaluateBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径记录ID As Long, Optional ByVal lng阶段ID As Long, Optional ByVal int场合 As Integer = -1, Optional ByVal strReserve As String) As Boolean
'功能：评估路径评估后
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号1256=临床路径应用
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lng路径记录ID=病人临床路径.id
'      lng阶段ID=当前评估的阶段ID
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=扩展参数

   PathEvaluateBefore = True

End Function

Public Function PathEvaluateAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径记录ID As Long, Optional ByVal lng阶段ID As Long, Optional ByVal int场合 As Integer = -1, Optional ByVal strReserve As String) As Boolean
'功能：评估路径评估后
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号1256=临床路径应用
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lng路径记录ID=病人临床路径.id
'      lng阶段ID=当前评估的阶段ID
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=扩展参数

   PathEvaluateAfter = True

End Function

Public Function DiagnosisEnterIn(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lng原诊断id As Long, Optional ByVal int场合 As Integer = -1, _
    Optional ByVal int诊断类型 As Integer, Optional ByVal int诊断次序 As Integer, Optional ByVal strReserve As String) As Boolean
'功能：住院病案中输入诊断后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lngDiagID=诊断ID或疾病ID(根据系统参数决定),strDiagDesc=诊断描述
'      lng原诊断id=0 表示新添加的诊断， 不为0表示修改诊断，lng原诊断id的值就是修改前的诊断ID或疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      int诊断类型=1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
'      int诊断次序=诊断类型下的顺序号
'      strReserve=扩展参数
      
    DiagnosisEnterIn = True

End Function

Public Function GetFormHwnd(Optional ByVal strReserve As String) As Long
'功能：病人信息管理 模块中调用 获取要嵌入的窗体句柄，
'参数：
'      strReserve=扩展参数
    GetFormHwnd = 0
End Function

Public Function PatiInfoSaveBefore(ByVal lng病人ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：病人信息管理 模块中调用 病人信息管理中保存病人信息前，用于控制一些检查类的业务，返回False终止保存操作。
'参数：lng病人ID   病人ID
'      strReserve=扩展参数
    PatiInfoSaveBefore = True
End Function

Public Function PatiInfoSaveAfter(ByVal lng病人ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：病人信息管理 模块中调用 病人信息管理中保存病人信息后，事务提交后调用，插件中渠道自行处理数据保存
'参数：lng病人ID   病人ID
'      strReserve=扩展参数
    PatiInfoSaveAfter = True
End Function
