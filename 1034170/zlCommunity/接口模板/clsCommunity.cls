VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommunity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Initialize(cnMain As ADODB.Connection) As Boolean
'功能：初始化社区接口部件
'参数：cnMain=数据库连接对象
'返回：如果初始化成功，返回True
    Dim strSQL As String
    
    On Error GoTo errH
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    '其他必要的代码(如果需要)
    '###
        
    Initialize = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Terminate() As Boolean
'功能：终止社区接口，在HIS调用主程序退出时调用
'返回：执行成功/失败
    
    Terminate = True
End Function

Public Function Setup(ByVal int社区 As Integer) As Boolean
'功能：对指定的社区接口参数进行设置
    
    '在参数设置窗体中根据需要添加控件
    '窗体中保存时调用过程"zl_社区参数_Update"
    Setup = frmSetup.ShowMe
End Function

Public Function Identify(ByVal lngSys As Long, ByVal lngModual As Long, _
    ByVal int社区 As Integer, str社区号 As String, colInfo As Collection, _
    Optional lng病人ID As Long, Optional ByVal lng就诊ID As Long) As Boolean
'功能：社区身份验证，验证失败返回False
'参数：
'     lngSys,lngModual(In)：调用该接口的系统及模块号，如100-标准版,1111-挂号。
'     int社区(In)：该病人所属的社区序号(如1-成都市青羊区)
'     str社区号(Out)：该病人在该社区中的唯一编号，如ABC123
'     colInfo(Out)：包含病人各个信息项目内容的集合，Key为"_字段名"，如"_姓名"；失败时返回空集(非Nothing)
'     lng病人ID(In/Out)：可选参数，如果已确定病人或者补充身份验证时传入。可用为Out参数。
'     lng就诊ID(In)：可选参数，补充身份验证时使用；对应"挂号ID"或"主页ID"，可通过调用模块场合确定。
'返回：如果验证成功则返回True及相关Out参数
    
    '在身份验证窗体中根据需要添加控件及增加窗体函数的参数
    'HIS中根据返回的病人相关信息，在业务确认时保存(涉及过程"Zl_病人社区信息_Insert")
    Set colInfo = New Collection
    
    If frmIdentify.ShowMe Then
        '其他返回参数
        str社区号 = "ABC123"
        
        '病人基本信息(目前挂号支持)
        colInfo.Add "张爱玲", "_姓名"
        colInfo.Add "女", "_性别"
        colInfo.Add "28岁", "_年龄" '注意要包含单位(岁月天),不含单位时缺省为岁
        colInfo.Add "1980-01-01 08:08", "_出生日期" '注意日期格式,可以包含出生时间
        colInfo.Add "汉族", "_民族"
        colInfo.Add "未婚", "_婚姻状况"
        colInfo.Add "", "_职业"
        colInfo.Add "110123198001011234", "_身份证号"
        colInfo.Add "成都市青羊区", "_家庭地址"
        colInfo.Add "028-88661234", "_家庭电话"
        colInfo.Add "400000", "_家庭地址邮编" '户口邮编
        colInfo.Add "成都中联信息有限公司", "_工作单位"
        colInfo.Add "028-88664321", "_单位电话"
        colInfo.Add "400000", "_单位邮编"
                
        '门诊医生站补充验证支持
        colInfo.Add "中国", "_国籍"
        colInfo.Add "成都市青羊区", "_出生地点"
        colInfo.Add "400000", "_户口邮编"
        colInfo.Add "张三", "_联系人姓名"
        colInfo.Add "朋友", "_联系人关系"
        colInfo.Add "028-88660000", "_联系人电话"
        colInfo.Add "成都市青羊区", "_联系人地址"
                
        '注：补充登记时，特殊信息可以在这里直接写入数据库(独立事务)
        
        Identify = True
    End If
End Function

Public Function GetCommunityFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int社区 As Integer) As String
'功能：返回特定社区在特定模块中的可调用功能，用于调用模块的社区菜单显示
'参数：都是传入参数
'      lngSys,lngModual=系统号，模块号
'      int社区=特定的社区；当医院可对多个社区的病人就诊时，用于区分不同社区的功能
'返回：功能描述串="功能号1,功能名1,功能说明1;功能号2,功能名2,功能说明2;..."
'      功能号为>0的整数，可以不连续；功能说明可为空，但需加上间隔符；功能名和功能说明中不能包含半角的间隔符
'      如果某个功能需要在病人接诊时自动调用(也可手工调用),则在前面加上"Auto:"前缀，如"Auto:1,档案查看,查看社区档案;...",这种功能只允许有一个
'      返回空表示不提供功能

    If lngSys \ 100 = 1 And lngModual = 1260 Then
        '门诊医生站
        GetCommunityFunc = "Auto:1,档案查阅,;2,档案修改,;3,档案下载,;4,档案上传,;5,其它管理,"
    Else
        '###
    End If
End Function

Public Function CommunityFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int功能号 As Integer, ByVal int社区 As Integer, ByVal str社区号 As String, ByVal lng病人ID As Long, ByVal lng就诊ID As Long) As Boolean
'功能：社区功能调用，调用平台接口中提供的社区病人业务功能。
'参数：都是传入参数
'      int功能号：由"GetCommunityFunc"返回的预定的一个编号序列，接口中根据这个编号再调用平台接口的相关功能。
'      int社区、str社区号、lng病人ID、lng就诊ID：当前病人的就诊标识，其中就诊ID对应"挂号ID"或"主页ID"，可通过调用模块场合确定
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    
    Select Case int功能号
    Case 1
        If lngSys \ 100 = 1 And lngModual = 1260 Then
            '标准版,门诊调用,根据需要调整SQL
            strSQL = "Select A.姓名,A.性别,A.年龄,B.执行状态 From 病人信息 A,病人挂号记录 B Where A.病人ID=B.病人ID And A.病人ID=[1] And B.ID=[2]"
        ElseIf lngSys \ 100 = 1 And lngModual = 1261 Then
            '标准版,住院调用,根据需要调整SQL
            strSQL = "Select A.姓名,A.性别,A.年龄,B.出院日期 From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID And B.病人ID=[1] And B.主页ID=[2]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CommunityFunc", lng病人ID, lng就诊ID)
        
        '调用对方接口功能
        '###
        MsgBox "测试病人=" & rsTmp!姓名, vbInformation, gstrSysName
        
        CommunityFunc = True
    Case Else
        CommunityFunc = True
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function UpdateInfo(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int社区 As Integer, ByVal str社区号 As String, ByVal lng病人ID As Long, ByVal lng就诊ID As Long) As Boolean
'功能：社区档案同步，HIS中的社区病人信息发生变化时，调用平台接口对中心档案库的数据进行同步。
'参数：都是传入参数
'说明：该接口在HIS对应功能的同一个数据库事务中调用，如果调用失败，则HIS中回退事务。
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    
    If lngSys \ 100 = 1 And lngModual = 1260 Then
        '标准版,门诊调用,根据需要调整SQL
        strSQL = "Select A.姓名,A.性别,A.年龄 From 病人信息 A,病人挂号记录 B Where A.病人ID=B.病人ID And A.病人ID=[1] And B.ID=[2]"
    ElseIf lngSys \ 100 = 1 And lngModual = 1261 Then
        '标准版,住院调用,根据需要调整SQL
        strSQL = "Select A.姓名,A.性别,A.年龄 From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID And B.病人ID=[1] And B.主页ID=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CommunityFunc", lng病人ID, lng就诊ID)
    
    '调用对方接口功能
    '###
    
    UpdateInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ClinicSubmit(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int社区 As Integer, ByVal str社区号 As String, ByVal lng病人ID As Long, ByVal lng就诊ID As Long) As Boolean
'功能：就诊信息提交，社区病人完成本次就诊时(比如医生站完成就诊)，调用接口将本次就诊过程中的业务数据信息上传到中心档案库。
'参数：都是传入参数
'说明：该接口在HIS对应功能的同一个数据库事务中调用，如果调用失败，则HIS中回退事务。
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    
    If lngSys \ 100 = 1 And lngModual = 1260 Then
        '标准版,门诊调用,根据需要调整SQL
        strSQL = "Select A.姓名,A.性别,A.年龄 From 病人信息 A,病人挂号记录 B Where A.病人ID=B.病人ID And A.病人ID=[1] And B.ID=[2]"
    ElseIf lngSys \ 100 = 1 And lngModual = 1261 Then
        '标准版,住院调用,根据需要调整SQL
        strSQL = "Select A.姓名,A.性别,A.年龄 From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID And B.病人ID=[1] And B.主页ID=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CommunityFunc", lng病人ID, lng就诊ID)
    
    '调用对方接口功能
    '###
    
    ClinicSubmit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
