Attribute VB_Name = "mdlDefine"
Option Explicit

Public gobjComLib As Object
Public gstrMessage As String                    '消息
Public gobjConn As ADODB.Connection             'HIS的DB连接对象
Public gfrmOwner As Form                        '主窗体对象
Public glngSys As Long                          '主调程序系统号
Public glngModule As Long                       '主调程序模块号
Public gstrDBUser As String                     'HIS的DB用户名
Public gstrRegHospital As String                '注册医院名称
Public gcolConn As Collection                   'clsConnect对象集合
Public gcolDevice As Collection                 'clsDevice对象集合
Public gobjSOAP As Object                       'MSSOAP对象
Public gstrPrivs As String
'Public grsUpload As ADODB.Recordset
Public grsParam As ADODB.Recordset               '参数数据集

Public glngUserId As Long
Public glngDeptId As Long
Public gstrUserCode As String
Public gstrUserName As String
Public gstrUserAbbr As String
Public gstrDeptCode As String
Public gstrDeptName As String


Public gstrSQL As String

Public Const GINT_INTERFACE_MODULENO = 1348
Public Const GSTR_INTERFACE_NAME = "药房自动化接口"
Public Const GSTR_SEPARAT = "|"
Public Const GSTR_SEPARAT_CHILD = ";"
Public Const GSTR_DEVICE_KEY = "D_"

'自动化系统连接类型
Public Enum enuLinkType
    DB
    WEBServices
    Directory
End Enum

'嵌入菜单号
Public Enum enuMenuNo
    药品信息 = 1
    药品库存
    设备开关
    上传设置
End Enum

Private Type Type_Params
    '设备对应的参数
    int服务对象() As Integer                  '1-门诊；2-住院
    int配药对应业务() As Integer              '1-门诊收费；2-处方发药配药功能；3-处方发药发药功能
    bln启用发送通知() As Boolean              '1-启用
    str单据类型() As String                   '按位分别表示长嘱、临嘱、记账单；1表示选择，0表示未选择
    str药品剂型() As String                   'Null表示所有药品剂型；如果需要指定某些剂型，格式：“粉型,片剂,…
    
    lngDeviceID() As Long                     '设备ID
    lngStockID() As Long                      '设备对应的药房ID
    blnStart() As Boolean                     '设备是否启用
End Type
Public gDeviceParams As Type_Params

Public Function GetDeviceID()

End Function

Public Function GetJudge_IsNeedUpload(ByVal lngModule As Boolean, ByVal bytType As Byte) As Boolean
'功能：判断当前业务环节是否需要上传数据
'参数：
'   lngModule：模块号
'   bytType：
'       0: 药品基本信息上传
'       1: 门诊处方上传 (配药)
'       2: 门诊发药通知 (发药)
'       3: 住院药品医嘱上传 (配、发药)
'       4: 药品库存上传

'    If grsParam Is Nothing Then
'        Call GetParam(0)
'    End If
    
    Select Case lngModule
        Case 1121   '门诊收费
            If bytType = 1 Then
                grsParam.Filter = "参数名='配药对应业务' And 参数值=1 "
                GetJudge_IsNeedUpload = Not grsParam.EOF
            End If
            Exit Function
        Case 1341   '处方发药
            grsParam.Filter = "参数名='服务对象' And 参数值=1 "
            If grsParam.EOF Then
                Exit Function
            End If
            
            Select Case bytType
                Case 1
                    grsParam.Filter = "参数名='配药对应业务' And 参数值=1 "
                Case 2
                    grsParam.Filter = "参数名='发送对应业务' And 参数值=1 "
                Case 0, 4
                    GetJudge_IsNeedUpload = True
                    Exit Function
                Case Else
                    Exit Function
            End Select
            
            GetJudge_IsNeedUpload = Not grsParam.EOF
            Exit Function
        Case 1342   '部门发药
            Select Case bytType
                Case 3
                    grsParam.Filter = "参数名='服务对象' And 参数值=2 "
                Case 0, 4
                    GetJudge_IsNeedUpload = True
                    Exit Function
                Case Else
                    Exit Function
            End Select
            GetJudge_IsNeedUpload = Not grsParam.EOF
            Exit Function
    End Select

End Function

Public Sub GetDeviceParam()
'功能：获取设备对应的参数值，并存放到公共变量中
'参数：
'   lngDevicdID：设备ID
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    gstrSQL = "Select * From 药房注册设备 Order by 设备ID "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDeviceParam")
    
    Do While Not rsData.EOF
        ReDim Preserve gDeviceParams.lngDeviceID(UBound(gDeviceParams.lngDeviceID) + 1)
        ReDim Preserve gDeviceParams.lngStockID(UBound(gDeviceParams.lngStockID) + 1)
        ReDim Preserve gDeviceParams.blnStart(UBound(gDeviceParams.blnStart) + 1)
        
        ReDim Preserve gDeviceParams.int服务对象(UBound(gDeviceParams.int服务对象) + 1)
        ReDim Preserve gDeviceParams.int配药对应业务(UBound(gDeviceParams.int配药对应业务) + 1)
        ReDim Preserve gDeviceParams.bln启用发送通知(UBound(gDeviceParams.bln启用发送通知) + 1)
        ReDim Preserve gDeviceParams.str单据类型(UBound(gDeviceParams.str单据类型) + 1)
        ReDim Preserve gDeviceParams.str药品剂型(UBound(gDeviceParams.str药品剂型) + 1)
        
        gDeviceParams.lngDeviceID(UBound(gDeviceParams.lngDeviceID)) = Val(rsData!设备ID)
        gDeviceParams.lngStockID(UBound(gDeviceParams.lngStockID)) = Val(rsData!部门ID)
        gDeviceParams.blnStart(UBound(gDeviceParams.blnStart)) = (Val(NVL(rsData!启用, 0)) = 1)
        
        rsData.MoveNext
    Loop
    
    gstrSQL = "Select a.设备id, b.参数号, b.参数名, a.参数值, b.缺省值 From 药房设备参数 A, 自动发药参数 B Where a.参数id = b.Id Order by a.设备id "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDeviceParam")
    Do While Not rsData.EOF
        rsData.Filter = "参数名='服务对象'"
        If Not rsData.EOF Then gDeviceParams.int服务对象(rsData.AbsolutePosition - 1) = Val(NVL(rsData!参数值, rsData!缺省值))
        
        rsData.Filter = "参数名='配药对应业务'"
        If Not rsData.EOF Then gDeviceParams.int配药对应业务(rsData.AbsolutePosition - 1) = Val(NVL(rsData!参数值, rsData!缺省值))
        
        rsData.Filter = "参数名='发送对应业务'"
        If Not rsData.EOF Then gDeviceParams.bln启用发送通知(rsData.AbsolutePosition - 1) = (Val(NVL(rsData!参数值, rsData!缺省值)) = 1)
        
        rsData.Filter = "参数名='单据类型'"
        If Not rsData.EOF Then gDeviceParams.str单据类型(rsData.AbsolutePosition - 1) = Val(NVL(rsData!参数值, rsData!缺省值))
        
        rsData.Filter = "参数名='药品剂型'"
        If Not rsData.EOF Then gDeviceParams.str药品剂型(rsData.AbsolutePosition - 1) = Val(NVL(rsData!参数值, rsData!缺省值))
        
        rsData.MoveNext
    Loop
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetUserInfo()
    Dim strSQL As String
    Dim rsUser As New ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select R.*,D.编码 as 部门编码,D.名称 as 部门名称,P.编号,P.姓名,P.简码, USER 用户名 " & _
            " From 上机人员表 U,人员表 P,部门表 D,部门人员 R" & _
            " Where U.人员ID = P.ID And R.部门ID = D.ID And P.ID=R.人员ID and U.用户名=USER and R.缺省=1 " & _
            "       and (p.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or p.撤档时间 Is Null)"
    Set rsUser = gobjComLib.zldatabase.OpenSQLRecord(strSQL, "获取用户信息")
    With rsUser
        If Not .EOF Then
            gstrDBUser = !用户名
            glngUserId = !人员ID '当前用户id
            gstrUserCode = !编号 '当前用户编码
            gstrUserName = IIf(IsNull(!姓名), "", !姓名) '当前用户姓名
            gstrUserAbbr = IIf(IsNull(!简码), "", !简码) '当前用户简码
            glngDeptId = !部门ID '当前用户部门id
            gstrDeptCode = !部门编码 '当前用户
            gstrDeptName = !部门名称 '当前用户
        Else
            gstrDBUser = ""
            glngUserId = 0 '当前用户id
            gstrUserCode = "" '当前用户编码
            gstrUserName = "" '当前用户姓名
            gstrUserAbbr = "" '当前用户简码
            glngDeptId = 0 '当前用户部门id
            gstrDeptCode = "" '当前用户
            gstrDeptName = "" '当前用户
        End If
    End With
    Exit Function

errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Function

Public Function FindDeviceID(ByVal fldDeptID As Field, ByVal fldDrugType As Field, ByVal fldBill As Field, ByVal fldServiceObject As Field) As Long
'功能：获取注册设备ID
'参数：
'  fldDeptID：药房ID
'  fldDrugType：药品剂型
'  fldBill：单据类型
'  fldServiceObject：服务对象
'返回：设备ID

    Dim rsDevice As ADODB.Recordset
    Dim strTmp As String
    Dim strDrugType As String
    Dim strBill As String
    Dim strServiceObject As String
    Dim lngDeptID As Long, lngDeviceID As Long

    On Error GoTo errHandle

    '药房ID
    lngDeptID = fldDeptID
    
    '服务对象
    strTmp = Trim(gobjComLib.zlCommFun.NVL(fldServiceObject))
    strServiceObject = IIf(strTmp = "", "0", IIf(strTmp = "门诊", "1", "2"))
    
    '药品剂型
    strTmp = Trim(gobjComLib.zlCommFun.NVL(fldDrugType))
    strDrugType = "%|" & IIf(strTmp = "", "????", strTmp) & "|%"
    
    '单据类型
    strTmp = Trim(gobjComLib.zlCommFun.NVL(fldBill))
    strBill = IIf(strTmp = "", "0", IIf(strTmp = "长嘱", "1", IIf(strTmp = "临嘱", "2", "3")))
    strBill = "%;" & strBill & ";%"
    
    gstrSQL = "Select Id " & _
              "From (Select a.Id, a.编码, a.名称, a.型号, a.启用, Max(b.名称) 连接名, " & _
              "        Max(Decode(d.参数号, 1, d.参数值, Null)) 服务对象, " & _
              "        Max(Decode(d.参数号, 4, d.参数值, Null)) 单据类型, " & _
              "        Max(Decode(d.参数号, 5, d.参数值, Null)) 药品剂型, " & _
              "        Max(Decode(d.参数号, 2, d.参数值, Null)) 配药业务, " & _
              "        Max(Decode(d.参数号, 3, d.参数值, Null)) 发药业务  " & _
              "      From 药房注册设备 A, 药房设备连接 B," & _
              "        (Select b.设备id, b.参数值, a.参数号 From Zlparameters A, 药房设备参数 B Where a.Id = b.参数id) D " & _
              "         Where a.连接id = b.Id And a.Id = d.设备id(+) And a.部门id = [1] " & _
              "      Group By a.Id, a.编码, a.名称, a.型号, a.启用) A " & _
              "Where '|' || 药品剂型 || '|' Like [2] and 服务对象 = [3] "
    If strServiceObject = "2" Then
        '服务于门诊，忽略单据类型；只有服务于住院才判断单据类型
        gstrSQL = gstrSQL & " and 单据类型 like [4] "
    End If
    On Error GoTo errSQL
    Set rsDevice = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药房注册设备ID", lngDeptID, strDrugType, strServiceObject, strBill)
    On Error GoTo errHandle
    
    If rsDevice.EOF = False Then
        FindDeviceID = rsDevice!ID
    End If
    rsDevice.Close
    Exit Function
    
errHandle:
    gstrMessage = Err.Description
    Exit Function

errSQL:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Function

Public Function FindDevice(ByVal lngID As Long) As clsDevice
'功能：找到设备对象，如果没有找到，就实例一个
'参数：
'   lngID：设备ID
'返回：clsDevice对象

    Dim strKey As String
    Dim i As Integer
    
    If lngID = 0 Then Exit Function
    
    If gcolDevice Is Nothing Then
        strKey = CreateDevice(lngID)
        If strKey <> "" Then Set FindDevice = gcolDevice(strKey)
    Else
        '找设备对象
        If gcolDevice(GSTR_DEVICE_KEY & lngID) Is Nothing Then
            strKey = CreateDevice(lngID)
            If strKey <> "" Then Set FindDevice = gcolDevice(strKey)
        Else
            FindDevice = gcolDevice(GSTR_DEVICE_KEY & lngID)
        End If
    End If
    
    Exit Function
    
errHandle:
    Set FindDevice = Nothing
    gstrMessage = "未找到条例条件的注册设备。"
End Function

Public Function CreateDevice(ByVal lngID As Long) As String
'功能：实例设备对象
'参数：
'   lngDeptID：设备ID
'返回：设备对象Key
    Dim rsTmp As ADODB.Recordset
    Dim strKey As String
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select a.Id, a.编码, a.名称, a.型号, a.部门id, a.启用, Max(d.参数号) 参数号, Max(b.名称) 连接名, " & _
              "    Max(Decode(d.参数号, 1, d.参数值, Null)) 服务对象," & _
              "    Max(Decode(d.参数号, 4, d.参数值, Null)) 单据类型," & _
              "    Max(Decode(d.参数号, 5, d.参数值, Null)) 药品剂型," & _
              "    Max(Decode(d.参数号, 2, d.参数值, Null)) 配药业务," & _
              "    Max(Decode(d.参数号, 3, d.参数值, Null)) 发药业务 " & _
              "From 药房注册设备 A, 药房设备连接 B, " & _
              "    (Select b.设备id, b.参数值, a.参数号 From Zlparameters A, 药房设备参数 B Where a.Id = b.参数id) D " & _
              "Where a.连接id = b.Id And a.Id = d.设备id(+) And a.Id = [1] " & _
              "Group By a.Id, a.编码, a.名称, a.型号, a.部门id, a.启用 "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药房注册设备", lngID)
    If Not rsTmp.EOF Then
        strKey = GSTR_DEVICE_KEY & rsTmp!ID
        gcolDevice.Add New clsDevice, strKey
        With gcolDevice(strKey)
            .ID = rsTmp!ID
            .DeptID = rsTmp!DeptID
            .Link = gcolConn(rsTmp!连接名)
            .ServiceObject = gobjComLib.zlCommFun.NVL(rsTmp!服务对象, 0)
            .Bill = gobjComLib.zlCommFun.NVL(rsTmp!单据类型)
            .Enabled = gobjComLib.zlCommFun.NVL(rsTmp!启用, 0) = 1
            .DrugType = gobjComLib.zlCommFun.NVL(rsTmp!药品剂型)
            .DispenseFunc = Val(gobjComLib.zlCommFun.NVL(rsTmp!配药业务))
            .DispensingFunc = Val(gobjComLib.zlCommFun.NVL(rsTmp!发药业务))
        End With
        CreateDevice = strKey
    End If
    rsTmp.Close
    Exit Function
    
errHandle:
    gstrMessage = "尚未注册设备信息，实例设备对象失败。"
End Function

Public Function TestURL(ByVal strURL As String) As Boolean
'功能：测试URL是否连接
'参数：
'  strURL：URL地址
'返回：True连接；False未连接
    Dim objSOAP As Object

    On Error Resume Next
    Set objSOAP = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        gstrMessage = Err.Description
        Err.Clear
        On Error GoTo errSOAP
        Set objSOAP = CreateObject("MSSOAP.SoapClient")
    End If
    
    '测试
    objSOAP.MSSoapInit strURL
    If objSOAP.FaultCode <> "" Then
        gstrMessage = objSOAP.FaultString
        Set objSOAP = Nothing
    Else
        TestURL = True
        Set objSOAP = Nothing
    End If
    Exit Function
    
errSOAP:
    gstrMessage = Err.Description
End Function

Public Sub CreateWebServices(ByVal strURL As String, ByRef objWS As Object)
'功能：创建WebServices对象
'参数：
'  strURL：
'  objWS：实参对象

    On Error Resume Next
    Set objWS = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        gstrMessage = Err.Description
        Err.Clear
        On Error GoTo errSOAP
        Set objWS = CreateObject("MSSOAP.SoapClient")
    End If
    
    objWS.MSSoapInit strURL
    If objWS.FaultCode <> "" Then
        gstrMessage = objWS.FaultString
        Set objWS = Nothing
    End If
    Exit Sub
    
errSOAP:
    gstrMessage = Err.Description
    Set objWS = Nothing
End Sub

Public Function GetConnectStrEle(ByVal strConnect As String, ByVal bytType As Byte, ByVal strName As String) As String
'功能：提取连接内容的元素值
'参数：
'  strConnect：连接内容
'  bytType：连接类型
'  strName：要提取的元素名
'返回：元素值

    Dim arrEle As Variant
    Dim i As Integer

    Select Case bytType
        Case enuLinkType.WEBServices
            
            arrEle = Split(strConnect, GSTR_SEPARAT_CHILD)
            For i = LBound(arrEle) To UBound(arrEle)
                If UCase(strName) = Split(UCase(arrEle(i)), "=")(0) Then
                    GetConnectStrEle = Mid(arrEle(i), InStr(arrEle(i), "=") + 1)
                    Exit For
                End If
            Next
            Set arrEle = Nothing
    End Select
End Function

Public Function SetMenuItem(ByVal intFunc As Integer) As Boolean
'功能：设置功能菜单项
'参数：
'  intFunc：功能号
    
    Dim objMenuItem As Object
    
    On Error GoTo errHandle
    Load gfrmOwner.mnuDrugPackerItems(gfrmOwner.mnuDrugPackerItems.UBound + 1)
    Set objMenuItem = gfrmOwner.mnuDrugPackerItems(gfrmOwner.mnuDrugPackerItems.UBound)
    Select Case intFunc
        Case enuMenuNo.药品信息
            objMenuItem.Caption = "药品信息上传(&D)"
        Case enuMenuNo.药品库存
            objMenuItem.Caption = "药品库存上传(&R)"
        Case enuMenuNo.设备开关
            objMenuItem.Caption = "设备开关(&S)"
        Case enuMenuNo.上传设置
            objMenuItem.Caption = "上传参数设置(&U)"
    End Select
    objMenuItem.Tag = intFunc
    SetMenuItem = True
    
errHandle:
    Set objMenuItem = Nothing
    If Err.Number <> 0 Then gstrMessage = "自动化接口嵌入式菜单创建失败！"
End Function

Public Sub ShowMenuItem()
'功能：显示嵌入菜单栏
    With gfrmOwner
        .mnuDrugPacker.Visible = True
        .mnuDrugPackerItems(0).Visible = False
    End With
End Sub

'Public Function GetDeviceParam(ByVal lngDeviceID As Long, ByVal lngParamNO As Long) As String
''功能：获取指定设备、指定参数号的参数值
''参数：
''  lngDeviceID：设备ID
''  lngParamNO：参数号
''返回：设备参数值
'
'    Dim rsTmp As ADODB.Recordset
'
'    On Error GoTo errHandle
'    gstrSQL = "Select b.参数值 From Zlparameters A, 药房设备参数 B Where a.Id = b.参数id And b.设备id = [1] And a.参数号 = [2] "
'    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取设备的服务对象", lngDeviceID, lngParamNO)
'    If rsTmp.EOF = False Then
'        rstmp!参数值
'    End If
'    Exit Function
'
'errHandle:
'    If gobjComLib.ErrCenter = 1 Then Resume
'    gstrMessage = Err.Description
'End Function
 
Public Function GetHisRecord_DrugInf(ByVal lngDrugID As Long) As ADODB.Recordset
'功能：获取HIS端基本数据，药品基本信息
'参数：
'   lngDrugID：药品ID，如果为0表示所有药品

    gstrSQL = "Select Decode(a.类别, '5', '西药', '6', '成药', '草药') As 材质, e.分类id, f.名称 As 分类名称, g.药名id As 品种id, e.名称 As 品种名称," & vbNewLine & _
        " g.药品id As 规格id, h.药品剂型 As 剂型, e.编码, a.名称 As 通用名, b.简码 As 拼音简码, c.名称 As 商品名, d.名称 As 英文名, a.规格, e.计算单位 As 剂量单位," & vbNewLine & _
        " g.剂量系数, a.计算单位, g.门诊单位, g.门诊包装, g.住院单位, g.住院包装, g.药库单位, g.药库包装, j.编码 As 生产商编号, a.产地 As 生产商名称, i.现价 As 售价" & vbNewLine & _
        " From 收费项目目录 A, 收费项目别名 B, 收费项目别名 C, 收费项目别名 D, 诊疗项目目录 E, 诊疗分类目录 F, 药品规格 G, 药品特性 H, 收费价目 I, 药品生产商 J" & vbNewLine & _
        " Where a.Id = b.收费细目id(+) And b.性质(+) = 1 And b.码类(+) = 1 And a.Id = c.收费细目id(+) And c.性质(+) = 3 And c.码类(+) = 1 And" & vbNewLine & _
        " a.Id = d.收费细目id(+) And d.性质(+) = 2 And a.Id = g.药品id And g.药名id = e.Id And e.分类id = f.Id And g.药名id = h.药名id And" & vbNewLine & _
        " a.Id = i.收费细目id And a.产地 = j.名称(+) And a.类别 In ('5', '6', '7') And Sysdate Between i.执行日期 And" & vbNewLine & _
        " Nvl(i.终止日期, Sysdate) And a.撤档时间 = Nvl(a.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD'))"
    If lngDrugID <> 0 Then
        gstrSQL = gstrSQL & " And A.id=[1] "
    End If
    gstrSQL = gstrSQL & " Order By Decode(a.类别, '5', '西药', '6', '成药', '草药'), a.Id"
    
    Set GetHisRecord_DrugInf = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_DrugInf", lngDrugID)

End Function

Public Function GetHisRecord_ReceipDetail(ByVal strKey As String) As ADODB.Recordset
'功能：获取HIS端基本数据，处方药品明细信息
'参数：
'   strKey：单据;库房ID;NO[|单据;库房ID;NO][|...]
    Dim rsData As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim int单据 As Integer
    Dim lng库房ID As Long
    Dim strNO As String
    Dim i As Integer
    Dim n As Integer
    Dim arrKey As Variant
    
    '分解为数组
    arrKey = Split(strKey, "|")
    For i = 0 To UBound(arrKey)
        If arrKey(i) = "" Or InStr(1, arrKey(i), ";") = 0 Then Exit For
        
        '将格式字符串分解并分别执行SQL
        int单据 = Split(arrKey(i), ";")(0)
        lng库房ID = Split(arrKey(i), ";")(1)
        strNO = Split(arrKey(i), ";")(2)
        
        gstrSQL = "Select Distinct a.单据, a.No, a.填制日期 As 处方时间, a.库房id As 发药药房id, i.名称 As 发药药房, a.序号," & vbNewLine & _
            " Decode(b.类别, '5', '西药', '6', '成药', '草药') As 材质, g.分类id, k.名称 As 分类名称, g.Id As 品种id, g.名称 As 品种名称, j.药品剂型," & vbNewLine & _
            " a.药品id, b.编码 As 药品编码, b.名称 As 药品名称, c.名称 As 药品商品名, h.名称 As 药品英文名, b.规格 As 药品规格, g.计算单位 As 剂量单位, d.剂量系数," & vbNewLine & _
            " b.计算单位, d.门诊单位, d.门诊包装, a.批次, a.产地 As 生产商, a.批号, a.单量, Nvl(a.付数, 1) * a.实际数量 / d.门诊包装 As 数量," & vbNewLine & _
            " a.成本价 * d.门诊包装 As 成本价, a.零售价 * d.门诊包装 As 售价, e.应收金额, e.实收金额, a.用法 As 药品用法, a.频次" & vbNewLine & _
            " From 药品收发记录 A, 收费项目目录 B, 收费项目别名 C, 药品规格 D, 门诊费用记录 E, 诊疗项目目录 G, 收费项目别名 H, 部门表 I, 药品特性 J, 诊疗分类目录 K" & vbNewLine & _
            " Where a.药品id = b.Id And a.药品id = c.收费细目id(+) And c.性质(+) = 3 And c.码类(+) = 1 And a.药品id = h.收费细目id(+) And h.性质(+) = 2 And" & vbNewLine & _
            " a.药品id = d.药品id And a.费用id = e.Id And d.药名id = g.Id And a.库房id = i.Id And d.药名id = j.药名id And g.分类id = k.Id And" & vbNewLine & _
            " a.单据 = [1] And a.库房id = [2] And a.No = [3] " & vbNewLine & _
            " Order By a.单据, a.No, a.序号"
        
        If i = 0 Then
            Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipDetail", int单据, lng库房ID, strNO)
        Else
            Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipDetail", int单据, lng库房ID, strNO)
            
            '将数据结果添加到初始数据集中
            Do While Not rsTmp.EOF
                rsData.AddNew
                
                '注意：如果SQL列数增加或减少，相应调整n的结束值，目前SQL为34列
                For n = 0 To 33
                    rsData.Fields(n).Value = rsTmp.Fields(n).Value
                Next
                
                rsData.Update
                
                rsTmp.MoveNext
            Loop
        End If
    Next
    
    Set GetHisRecord_ReceipDetail = rsData
End Function

Public Function GetHisRecord_ReceipList(ByVal strKey As String) As ADODB.Recordset
'功能：获取HIS端基本数据，处方概要信息
'参数：
'   strKey：单据;库房ID;NO[|单据;库房ID;NO][|...]
    Dim rsData As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim int单据 As Integer
    Dim lng库房ID As Long
    Dim strNO As String
    Dim i As Integer
    Dim n As Integer
    Dim arrKey As Variant
    
    '分解为数组
    arrKey = Split(strKey, "|")
    For i = 0 To UBound(arrKey)
        If arrKey(i) = "" Or InStr(1, arrKey(i), ";") = 0 Then Exit For
        
        '将格式字符串分解并分别执行SQL
        int单据 = Split(arrKey(i), ";")(0)
        lng库房ID = Split(arrKey(i), ";")(1)
        strNO = Split(arrKey(i), ";")(2)
        
        gstrSQL = "Select a.单据, a.No, Decode(a.处方类型, 1, '儿科', 2, '急诊', 3, '精一', '4', '精二', '5', '麻醉', '普通') As 处方类型, a.病人id, a.主页id, a.姓名," & vbNewLine & _
            " c.性别, c.年龄, c.出生日期, c.身份, c.就诊卡号, c.门诊号, c.住院号, c.医保号, c.身份证号, c.Ic卡号, c.民族, c.国籍, c.区域, c.医疗付款方式 As 医保类型," & vbNewLine & _
            " Sum(d.应收金额) As 处方金额, Sum(d.实收金额) As 实收金额, a.填制日期 As 处方时间, d.开单部门id As 开单科室id, f.名称 As 开单科室, d.开单人 As 开单医生," & vbNewLine & _
            " a.库房id As 发药药房id, g.名称 As 发药药房, Decode(a.优先级, 1, '1', '2') As 优先级, h.编码 As 发药窗口编号, a.发药窗口" & vbNewLine & _
            " From 未发药品记录 A, 病人信息 C, 门诊费用记录 D, 药品收发记录 E, 部门表 F, 部门表 G, 发药窗口 H" & vbNewLine & _
            " Where a.单据 = e.单据 And a.No = e.No And a.库房id = e.库房id And a.病人id = c.病人id And e.费用id = d.Id And d.开单部门id = f.Id And" & vbNewLine & _
            " a.库房id = g.Id And a.发药窗口 = h.名称(+) And a.单据 = [1] And a.库房id = [2]  And a.No = [3] " & vbNewLine & _
            " Group By a.单据, a.No, Decode(a.处方类型, 1, '儿科', 2, '急诊', 3, '精一', '4', '精二', '5', '麻醉', '普通'), a.病人id, a.主页id, a.姓名, c.性别," & vbNewLine & _
            " c.年龄, c.出生日期, c.身份, c.就诊卡号, c.门诊号, c.住院号, c.医保号, c.身份证号, c.Ic卡号, c.民族, c.国籍, c.区域, c.医疗付款方式, a.填制日期, d.开单部门id," & vbNewLine & _
            " f.名称, d.开单人, a.库房id, g.名称, Decode(a.优先级, 1, '1', '2'), h.编码, a.发药窗口" & vbNewLine & _
            " Order By a.单据, a.库房id, Decode(a.优先级, 1, '1', '2'), a.No, a.填制日期"
        
        If i = 0 Then
            Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipList", int单据, lng库房ID, strNO)
        Else
            Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipList", int单据, lng库房ID, strNO)
            
            '将数据结果添加到初始数据集中
            Do While Not rsTmp.EOF
                rsData.AddNew
                
                '注意：如果SQL列数增加或减少，相应调整n的结束值，目前SQL为31列
                For n = 0 To 30
                    rsData.Fields(n).Value = rsTmp.Fields(n).Value
                Next
                
                rsData.Update
                
                rsTmp.MoveNext
            Loop
        End If
    Next
    
    Set GetHisRecord_ReceipList = rsData
End Function

Public Function GetHisRecord_ReceipInf(ByVal strKey As String) As ADODB.Recordset
'功能：获取HIS端基本数据，处方信息含药品明细
'参数：
'   strKey：单据;库房ID;NO[|单据;库房ID;NO][|...]
    Dim rsData As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim int单据 As Integer
    Dim lng库房ID As Long
    Dim strNO As String
    Dim i As Integer
    Dim n As Integer
    Dim arrKey As Variant
    
    '分解为数组
    arrKey = Split(strKey, "|")
    For i = 0 To UBound(arrKey)
        If arrKey(i) = "" Or InStr(1, arrKey(i), ";") = 0 Then Exit For
        
        '将格式字符串分解并分别执行SQL
        int单据 = Split(arrKey(i), ";")(0)
        lng库房ID = Split(arrKey(i), ";")(1)
        strNO = Split(arrKey(i), ";")(2)
        
        gstrSQL = "Select a.单据, a.No, Decode(a.处方类型, 1, '儿科', 2, '急诊', 3, '精一', '4', '精二', '5', '麻醉', '普通') As 处方类型, a.病人id, a.主页id, a.姓名," & vbNewLine & _
            " c.性别, c.年龄, c.出生日期, c.身份, c.就诊卡号, c.门诊号, c.住院号, c.医保号, c.身份证号, c.Ic卡号, c.民族, c.国籍, c.区域, c.医疗付款方式 As 医保类型," & vbNewLine & _
            " a.填制日期 As 处方时间, d.开单部门id As 开单科室id, f.名称 As 开单科室, d.开单人 As 开单医生, a.库房id As 发药药房id, g.名称 As 发药药房," & vbNewLine & _
            " Decode(a.优先级, 1, '1', '2') As 优先级, h.编码 As 发药窗口编号, a.发药窗口, e.序号, Decode(i.类别, '5', '西药', '6', '成药', '草药') As 材质," & vbNewLine & _
            " l.分类id, o.名称 As 分类名称, l.Id As 品种id, l.名称 As 品种名称, n.药品剂型, e.药品id, i.编码 As 药品编码, i.名称 As 药品名称, j.名称 As 药品商品名," & vbNewLine & _
            " m.名称 As 药品英文名, i.规格 As 药品规格, l.计算单位 As 剂量单位, k.剂量系数, i.计算单位, k.门诊单位, k.门诊包装, e.批次, e.产地 As 生产商, e.批号, e.单量," & vbNewLine & _
            " Nvl(e.付数, 1) * e.实际数量 / k.门诊包装 As 数量, e.成本价 * k.门诊包装 As 成本价, e.零售价 * k.门诊包装 As 售价, d.应收金额, d.实收金额, e.用法 As 药品用法," & vbNewLine & _
            " e.频次" & vbNewLine & _
            " From 未发药品记录 A, 病人信息 C, 门诊费用记录 D, 药品收发记录 E, 部门表 F, 部门表 G, 发药窗口 H, 收费项目目录 I, 收费项目别名 J, 药品规格 K, 诊疗项目目录 L, 收费项目别名 M, 药品特性 N," & vbNewLine & _
            " 诊疗分类目录 O" & vbNewLine & _
            " Where a.单据 = e.单据 And a.No = e.No And a.库房id = e.库房id And a.病人id = c.病人id And e.费用id = d.Id And d.开单部门id = f.Id And" & vbNewLine & _
            " a.库房id = g.Id And a.发药窗口 = h.名称(+) And e.药品id = i.Id And e.药品id = j.收费细目id(+) And j.性质(+) = 3 And j.码类(+) = 1 And" & vbNewLine & _
            " e.药品id = m.收费细目id(+) And m.性质(+) = 2 And e.药品id = k.药品id And k.药名id = l.Id And k.药名id = n.药名id And l.分类id = o.Id And" & vbNewLine & _
            " a.单据 = [1] And a.库房id = [2] And a.No = [3] " & vbNewLine & _
            " Order By a.单据, a.库房id, a.No, e.序号"

        If i = 0 Then
            Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipInf", int单据, lng库房ID, strNO)
        Else
            Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipInf", int单据, lng库房ID, strNO)
            
            '将数据结果添加到初始数据集中
            Do While Not rsTmp.EOF
                rsData.AddNew
                
                '注意：如果SQL列数增加或减少，相应调整n的结束值，目前SQL为58列
                For n = 0 To 57
                    rsData.Fields(n).Value = rsTmp.Fields(n).Value
                Next
                
                rsData.Update
                
                rsTmp.MoveNext
            Loop
        End If
    Next
    
    Set GetHisRecord_ReceipInf = rsData
End Function

Public Function GetHisRecord_AdviceInf(ByVal strKey As String) As ADODB.Recordset
'功能：获取HIS端基本数据，医嘱信息含药品明细
'参数：
'   strKey：药品ID串，格式为"药品ID,药品ID..."

    gstrSQL = "Select /*+ rule*/ a.病人id, a.标识号 As 住院号, a.床号, a.姓名, a.性别, a.年龄, q.出生日期, q.身份, q.就诊卡号, q.医保号, q.身份证号, q.Ic卡号, q.民族, q.国籍, q.区域," & vbNewLine & _
        " a.开单部门id As 开单部门id, r.编码 As 开单部门编码, r.名称 As 开单部门名称, a.病人科室id, s.编码 As 病人科室编码, s.名称 As 病人科室名称, a.病人病区id," & vbNewLine & _
        " f.编码 As 病人病区编码, f.名称 As 病人病区名称, b.对方部门id As 领药部门id, t.编码 As 领药部门编码, t.名称 As 领药部门名称," & vbNewLine & _
        " Decode(d.医嘱期效, 1, '长期', '临时') As 医嘱类型, a.开单人 As 开单医生, c.发送时间 As 医嘱发送时间, c.首次时间, c.末次时间, d.执行频次, d.频率次数, d.频率间隔," & vbNewLine & _
        " d.间隔单位, d.执行时间方案, d.医生嘱托, b.用法 As 药品用法, Decode(g.类别, '5', '西药', '6', '成药', '草药') As 材质, h.分类id, m.名称 As 分类名称," & vbNewLine & _
        " i.药名id As 品种id, h.名称 As 品种名称, l.药品剂型, b.药品id, g.编码 As 药品编码, g.名称 As 药品名称, n.名称 As 药品商品名, o.名称 As 药品英文名, g.规格," & vbNewLine & _
        " b.产地 As 生产商, b.批号, b.批次, i.剂量系数, h.计算单位 As 剂量单位, g.计算单位, i.住院单位, i.住院包装, b.单量," & vbNewLine & _
        " Nvl(b.付数, 1) * b.实际数量 / i.住院包装 As 数量, b.成本价 * i.住院包装 As 成本价, b.零售价 * i.住院包装 As 售价, b.零售金额 As 金额, b.Id As 收发id," & vbNewLine & _
        " b.库房id As 发药药房id, u.编码 As 发药药房编号, u.名称 As 发药药房, b.填制日期, b.审核日期" & vbNewLine & _
        " From 住院费用记录 A, 药品收发记录 B, 病人医嘱发送 C, 病人医嘱记录 D, 部门表 F, 收费项目目录 G, 诊疗项目目录 H, 药品规格 I, 药品特性 L, 诊疗分类目录 M, 收费项目别名 N, 收费项目别名 O," & vbNewLine & _
        " 病人信息 Q, 部门表 R, 部门表 S, 部门表 T, 部门表 U , Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V " & vbNewLine & _
        " Where a.Id = b.费用id And a.医嘱序号 = c.医嘱id And c.医嘱id = d.Id And b.No = c.No And a.病人病区id = f.Id And b.药品id = g.Id And" & vbNewLine & _
        " h.Id = i.药名id And b.药品id = i.药品id And i.药名id = l.药名id And h.分类id = m.Id And g.Id = n.收费细目id(+) And n.性质(+) = 3 And" & vbNewLine & _
        " n.码类(+) = 1 And g.Id = o.收费细目id(+) And o.性质(+) = 2 And a.病人id = q.病人id And a.开单部门id = r.Id And a.病人科室id = s.Id And" & vbNewLine & _
        " b.对方部门id = t.Id And b.库房id = u.Id And b.Id = v.Column_Value "
    Set GetHisRecord_AdviceInf = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_AdviceInf", strKey)

End Function

Public Function GetHisRecord_DrugStock(ByVal lngStockID As Long) As ADODB.Recordset
'功能：获取HIS端基本数据，药品库存信息
'参数：
'   lngStockID：库房ID

    gstrSQL = "Select Decode(a.类别, '5', '西药', '6', '成药', '草药') As 材质, e.分类id, f.名称 As 分类名称, g.药名id As 品种id, e.名称 As 品种名称," & vbNewLine & _
        " g.药品id As 规格id, h.药品剂型 As 剂型, e.编码, a.名称 As 通用名, b.简码 As 拼音简码, c.名称 As 商品名, d.名称 As 英文名, a.规格," & vbNewLine & _
        " Decode(a.是否变价, 1, '时价', '定价') As 价格属性, e.计算单位 As 剂量单位, g.剂量系数, a.计算单位, g.门诊单位, g.门诊包装, g.住院单位, g.住院包装, g.药库单位," & vbNewLine & _
        " g.药库包装, i.现价 As 售价, k.批次, k.效期, k.可用数量, k.实际数量, k.实际金额 As 实际金额, k.实际差价 As 实际差价, l.名称 As 供应商, k.上次采购价 As 采购价," & vbNewLine & _
        " k.上次批号 As 批号, k.上次生产日期 As 生产日期, k.上次产地 As 产地, k.批准文号, k.平均成本价" & vbNewLine & _
        " From 收费项目目录 A, 收费项目别名 B, 收费项目别名 C, 收费项目别名 D, 诊疗项目目录 E, 诊疗分类目录 F, 药品规格 G, 药品特性 H, 收费价目 I, 药品库存 K, 供应商 L" & vbNewLine & _
        " Where a.Id = b.收费细目id(+) And b.性质(+) = 1 And b.码类(+) = 1 And a.Id = c.收费细目id(+) And c.性质(+) = 3 And c.码类(+) = 1 And" & vbNewLine & _
        " a.Id = d.收费细目id(+) And d.性质(+) = 2 And a.Id = g.药品id And g.药名id = e.Id And e.分类id = f.Id And g.药名id = h.药名id And" & vbNewLine & _
        " a.Id = i.收费细目id And a.类别 In ('5', '6', '7') And Sysdate Between i.执行日期 And Nvl(i.终止日期, Sysdate) And" & vbNewLine & _
        " a.撤档时间 = Nvl(a.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And g.药品id = k.药品id And k.性质 = 1 And" & vbNewLine & _
        " k.上次供应商id = l.Id(+) And 库房id = [1] " & vbNewLine & _
        " Order By Decode(a.类别, '5', '西药', '6', '成药', '草药'), a.Id, k.批次"
    Set GetHisRecord_DrugStock = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_DrugStock", lngStockID)

End Function
