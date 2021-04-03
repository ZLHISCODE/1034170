VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugPacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    On Error Resume Next
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    If Err <> 0 Or gobjComLib Is Nothing Then
        If gblnShowMsg Then
            MsgBox "引用zl9ComLib部件失败！", vbInformation, GSTR_MESSAGE
        End If
        Exit Sub
    End If
End Sub

Private Sub Class_Terminate()
    Set gobjComLib = Nothing
    If gcnOutside Is Nothing Then Exit Sub
    If gcnOutside.State = adStateOpen Then gcnOutside.Close
End Sub

Public Function DBConnect() As Boolean
'连接中间数据库
    Dim strServer As String, strDBName As String, strUser As String, strPassword As String
    Dim blnConnectFinish As Boolean
    '查询注册表有无连接服务器的信息
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    If Trim(strUser) = "" Then
        '无：默认信息
        DBConnect = MSSQLServerOpen(MSTR_SERVER, MSTR_DBNAME, MSTR_USER, MSTR_PASSWORD)
    Else
        '有：注册表信息
        strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
        strDBName = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
        strPassword = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
        strPassword = StringEnDeCodecn(strPassword, 68)     '解密
        DBConnect = MSSQLServerOpen(strServer, strDBName, strUser, strPassword)
    End If
End Function

Public Function SetDBConnect() As Boolean
'设置数据库服务器连接参数
    frmOutsideLinkSet.Show vbModal
    '返回值
    SetDBConnect = frmOutsideLinkSet.gblnSetupFinish
End Function

Public Function TranDrugPacker(ByVal strData As String) As String
'功能： 传送药品自动分包数据
'参数： 分包数据字符串
'格式： 病区编码;库房组号;分包设备编号|收发ID1;病例号;...|收发ID2;病例号;...|收发ID3;病例号;...
'规则： 收发ID,病例号,病人ID,姓名,病区编码,病区名称,药师姓名,床号,服用方法,服药时间说明,
'       药品编码,药品名称,规格,剂量,剂量单位,服用数量,服用时间,分包设备编号,医嘱类型
'返回值：未成功传送的收发ID字符串
    Dim arrPrimary As Variant, arrSecondly As Variant, arrSecondlyVals As Variant
    Dim strInsert As String, strTmp As String, strID As String, strPageNO As String
    Dim i As Integer, j As Integer, intPageNO As Integer
    Dim rsInsert As New ADODB.Recordset
    Dim blnRollback As Boolean, blnInsert As Boolean, blnInserted As Boolean
    
    If gcnOutside Is Nothing Or gcnOutside.State = adStateClosed Then
        MsgBox "你未连接数据库，请先执行DBConnect()函数！", vbCritical, GSTR_MESSAGE
        TranDrugPacker = "NOT"
        Exit Function
    End If
    
    strTmp = Trim(strData)
    If strTmp = "" Then Exit Function
    
    arrPrimary = Split(Mid(strTmp, 1, InStr(1, strTmp, "|") - 1), ";")
    strTmp = Mid(strTmp, InStr(1, strTmp, "|") + 1)
    arrSecondly = Split(strTmp, "|")
    '取PageNO号
    strTmp = "select convert(char(6),getdate(),12) + right('000000'+cast(isnull(max(substring(page_no,7,len(page_no))),0)+1 as varchar(4)),4) max_no " _
           & "from dbo.atf_ypxx where convert(char(6),getdate(),12)=left(page_no,6)"
    rsInsert.Open strTmp, gcnOutside
    strPageNO = rsInsert!max_no
    rsInsert.Close
    
    '先传送表数据(从)
'    intPageNO = 1   '计数
'    intAbate = 0    '回滚数
    strInsert = "insert into dbo.atf_ypxx " _
              & "(DETAIL_SN,inpatient_no,p_id,name,ward_sn,ward_name,doctor,bed_no,comment,comm2,drug_code,drugname" _
              & ",specification,dosage,dos_unit,amount,occ_time,atf_no,pri_flag,Mz_flag,dept_name,page_no) " & Chr(13)
    strTmp = ""
    For i = LBound(arrSecondly) To UBound(arrSecondly)
        '得到元素
        arrSecondlyVals = Split(arrSecondly(i), ";")
        '组织字符串
        strTmp = strTmp & "select "
        For j = LBound(arrSecondlyVals) To UBound(arrSecondlyVals)
            Select Case j
            Case 0
                strTmp = strTmp & "'" & arrSecondlyVals(j) & "'"
            Case 1 To 12, 14, 20, 16 To 18
                strTmp = strTmp & ",'" & arrSecondlyVals(j) & "'"
            Case 13, 15, 19
                strTmp = strTmp & "," & arrSecondlyVals(j)
            End Select
        Next
        strTmp = strTmp & ",'" & strPageNO & "'"
        strTmp = strTmp & " union all " & Chr(13)
        '判断下条记录是否为同一收发ID
        strID = arrSecondlyVals(0)
        If i = UBound(arrSecondly) Then
            blnInsert = True
        Else
            If Mid(arrSecondly(i + 1), 1, InStr(1, arrSecondly(i + 1), ";") - 1) = strID Then
                blnInsert = False
            Else
                blnInsert = True
            End If
        End If
        '是否执行Insert语句
        If blnInsert = True Then
            blnRollback = False
            strTmp = Left(strTmp, Len(strTmp) - 11)
            
            gcnOutside.BeginTrans
            On Error GoTo errRollback
            rsInsert.Open strInsert & strTmp, gcnOutside
            On Error GoTo 0
            If blnRollback = False Then
                gcnOutside.CommitTrans
                blnInserted = True
            Else
'                intPageNO = intPageNO - intAbate - 1
                '记录未提交的收发ID
                TranDrugPacker = TranDrugPacker & strID & ";"
            End If
            If rsInsert.State = adStateOpen Then rsInsert.Close
            strTmp = ""
'            intAbate = 0
        Else
            strTmp = strTmp & Chr(13)
            '记录多少条相同的
'            intAbate = intAbate + 1
        End If
'        intPageNO = intPageNO + 1
    Next
    If rsInsert.State = adStateOpen Then rsInsert.Close
    '先传送表数据(主)
    If blnInserted Then
        blnRollback = False
        strTmp = "insert into dbo.atf_yp_page_no (ward_sn,group_no,atf_no,submit_time,page_no) " & Chr(13)
        strTmp = strTmp & "select "
        For i = LBound(arrPrimary) To UBound(arrPrimary)
            Select Case i
            Case 0 To 2
                strTmp = strTmp & "'" & arrPrimary(i) & "',"
'            Case 3
'                strTmp = strTmp & "getdate(),"
'            Case 4
'                strTmp = strTmp & "'" & strPageNO & "'"
            End Select
        Next
        strTmp = strTmp & "getdate(),'" & strPageNO & "'"
        'strTmp = Left(strTmp, Len(strTmp) - 1)
        '提交数据
        gcnOutside.BeginTrans
        On Error GoTo errRollback
        rsInsert.Open strTmp, gcnOutside
        On Error GoTo 0
        If blnRollback = False Then
            gcnOutside.CommitTrans
        Else
            '如果主表数据失败，同事删除从表对应数据
            strTmp = "delete dbo.atf_ypxx where page_no='" & strPageNO & "'"
            On Error Resume Next
            If rsInsert.State = adStateOpen Then rsInsert.Close
            rsInsert.Open strTmp, gcnOutside
            If rsInsert.State = adStateOpen Then rsInsert.Close
            '返回所有收发ID字符串
            strID = "": TranDrugPacker = ""
            For i = LBound(arrSecondly) To UBound(arrSecondly)
                If Left(arrSecondly(i), InStr(1, arrSecondly(i), ";") - 1) <> strID Then
                    strID = Left(arrSecondly(i), InStr(1, arrSecondly(i), ";") - 1)
                    TranDrugPacker = TranDrugPacker & strID & ";"
                End If
            Next
        End If
    End If
    'If gcnOutside.State = adStateOpen Then gcnOutside.Close
    If Trim(TranDrugPacker) <> "" Then
        '返回收发ID字符串
        TranDrugPacker = Left(TranDrugPacker, Len(TranDrugPacker) - 1)
    End If
    
    Exit Function

errRollback:
    gcnOutside.RollbackTrans
    blnRollback = True
    Resume Next
End Function

Public Function DrugPackerSet(ByVal cnZLHIS As ADODB.Connection, ByVal blnOutsideConnected As Boolean) As String
    Dim frmTmp As New frmDrugPackerSet
    frmTmp.ConnectHIS = cnZLHIS
    frmTmp.OutsideConnected = blnOutsideConnected
    frmTmp.Show vbModal
    '返回值
    DrugPackerSet = frmTmp.ResultString
    Set frmTmp = Nothing
End Function

Public Function TranDrugSingle(ByVal cnHIS As ADODB.Connection, ByVal lngID As Long) As Boolean
    Dim cmdInsert As New ADODB.Command, rsTmp As New ADODB.Recordset
    Dim strTmp As String, strInsert As String
    Dim lngExec As Long

    If cnHIS Is Nothing Or cnHIS.State = adStateClosed Then
        'MsgBox "ZLHIS数据库未连接！", vbCritical, GSTR_MESSAGE
        Exit Function
    End If
    If gcnOutside Is Nothing Or gcnOutside.State = adStateClosed Then
        'MsgBox "你未连接数据库，请先执行DBConnect()函数！", vbCritical, GSTR_MESSAGE
        Exit Function
    End If

    strTmp = "Select A.*, B.编码 厂商编码 " _
           & "From (Select A.药品id, B.编码, B.名称, B.规格, D.药品剂型 药品类型, A.剂量系数 剂量, E.计算单位 剂量单位," _
           & "        A.住院包装 包装数量, A.住院单位 包装单位, Nvl(A.上次产地, B.产地) 生产厂商," _
           & "        zlTools.zlSpellCode(B.名称) 拼音码" _
           & "      From 药品规格 A, 收费项目目录 B, 药品特性 D, 诊疗项目目录 E" _
           & "      Where A.药品id = B.ID And A.药名id = D.药名id And A.药名id = E.ID And B.类别 In ('5', '6', '7') And" _
           & "        Nvl(B.撤档时间, To_Date('3000-1-1', 'yyyy-mm-dd')) = To_Date('3000-1-1', 'yyyy-mm-dd') And" _
           & "        A.药品ID=" & lngID & ") A, 药品生产商 B " _
           & "Where A.生产厂商 = B.名称(+) order by cast(a.编码 as int) "
    rsTmp.Open strTmp, cnHIS
    If rsTmp.RecordCount <> 1 Then
        rsTmp.Close
        Exit Function
    End If
    gcnOutside.BeginTrans
    On Error GoTo errHand
    strInsert = "delete dbo.atf_his_druginfo where drug_code='" & rsTmp!编码 & "' and drugname='" & rsTmp!名称 & "' " & Chr(13) _
              & "insert into dbo.atf_his_druginfo (drug_code,drugname,specification,drug_type," _
              & "dosage,dos_unit,pack_amount,pack_name,manufactory,py_code,manu_no) " & Chr(13)
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        strTmp = "select '" & rsTmp!编码 & "'," _
               & "'" & rsTmp!名称 & "'," _
               & "'" & rsTmp!规格 & "'," _
               & "'" & rsTmp!药品类型 & "'," _
               & CDbl(rsTmp!剂量) & "," _
               & "'" & rsTmp!剂量单位 & "'," _
               & CDbl(rsTmp!包装数量) & "," _
               & "'" & rsTmp!包装单位 & "'," _
               & "'" & IIf(IsNull(rsTmp!生产厂商), "", rsTmp!生产厂商) & "'," _
               & "'" & rsTmp!拼音码 & "'," _
               & "'" & IIf(IsNull(rsTmp!厂商编码), "", rsTmp!厂商编码) & "' union all "
        strInsert = strInsert & strTmp & Chr(13)
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    strInsert = Left(strInsert, Len(strInsert) - 11)
    '更新数据
    With cmdInsert
        .ActiveConnection = gcnOutside
        .CommandText = strInsert
        .Execute 'lngExec, , adCmdText
    End With
    'If lngExec <= 0 Then GoTo errHand

    gcnOutside.CommitTrans
    TranDrugSingle = True
    Exit Function

errHand:
    gcnOutside.RollbackTrans
End Function

Public Function DYEY_MZ_IniSoap(Optional ByVal blnShowMsg As Boolean = True, Optional ByRef strMsg As String, Optional ByVal strUser As String, Optional ByVal cnOracle As ADODB.Connection) As Boolean
'功能：初始化接口
'参数：
'  blnShowMsg：显示消息提示对话框，默认显示
'  strMsg：返回消息
'返回值：true-成功，false-失败

    Dim strUrl As String '获取路径
    
    If gobjComLib Is Nothing Then
        Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    End If
    If cnOracle Is Nothing Then
        gobjComLib.InitCommon cnOracle
    End If
    
    strUser = gobjComLib.GetUnitName
    If strUser <> gstrUnit_DYEY And strUser <> gstrUnit_YZSZYY And strUser <> gstrUnit_JLSZXYY Then Exit Function
    On Error Resume Next
    Set gobjSOAP = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo errSoap
        Set gobjSOAP = CreateObject("MSSOAP.SoapClient")
    End If
    
    gblnShowMsg = blnShowMsg
    
    GoTo RRow
    
errSoap:
    If blnShowMsg Then
        MsgBox "系统初始化失败！" & vbCrLf & vbCrLf & "客户端未安装SOAP！" & vbCrLf & vbCrLf & "错误信息如下：" & vbCrLf & vbCrLf & Err.Description, vbCritical, GSTR_MESSAGE
    Else
        strMsg = "客户端未安装SOAP！"
    End If
    Exit Function
    
RRow:
    On Error GoTo errHandle
    strUrl = GetSetting("ZLSOFT", "公共模块\WebService路径", "WebUrl")
    
    If strUrl = "" Then
        If blnShowMsg Then
            MsgBox "未设置WebServic服务地址，请在程序中进行设置！", vbInformation + vbOKOnly, GSTR_MESSAGE
        Else
            strMsg = "未设置WebServic服务地址！"
        End If
        DYEY_MZ_IniSoap = False
    Else
        Call gobjSOAP.MSSoapInit(strUrl & "?WSDL")
        DYEY_MZ_IniSoap = True
    End If
    Exit Function
    
errHandle:
    If blnShowMsg Then
        MsgBox "初始化失败，请检测服务器地址。", vbCritical, GSTR_MESSAGE
    Else
        strMsg = "初始化失败，请检测服务器地址。"
    End If
    DYEY_MZ_IniSoap = False
End Function

Public Function DYEY_MZ_SetServer() As Boolean
'设置WebService服务地址
'返回值：true-成功，false-失败
    Dim strUrl As String
        
    On Error GoTo errHandle
    Call frmDYEY_MZ_SetServer.ShowMe(strUrl)
            
    If strUrl = "" Then
        MsgBox "未设置WebServic服务地址，请在程序中进行设置！", vbInformation + vbOKOnly, GSTR_MESSAGE
    Else
        If strUrl <> GetSetting("ZLSOFT", "公共模块\WebService路径", "WebUrl") Then
            Call gobjSOAP.MSSoapInit(strUrl & "?WSDL")
            Call SaveSetting("zlsoft", "公共模块\WebService路径", "WebUrl", strUrl)
        End If
        DYEY_MZ_SetServer = True
    End If
    Exit Function
errHandle:
    MsgBox "初始化失败，请检测服务器地址。", vbCritical, GSTR_MESSAGE
End Function

Public Function DYEY_MZ_TransDrug(ByVal intOprId As Integer, ByVal strUserCode As String, ByVal strUserName As String, ByRef strReturn As String) As Boolean
'HIS端向接口传递药品基础信息
'返回值：true-成功，false-失败
    Dim arrXML As Variant
    
    arrXML = GetXML_Drug
    
    If UBound(arrXML) = -1 Then
        DYEY_MZ_TransDrug = True
        Exit Function
    End If
    
    If DYEY_MZ_TransData(gType.IntDrug, intOprId, strUserCode, strUserName, arrXML, strReturn) Then
        DYEY_MZ_TransDrug = True
    End If
    
End Function

Public Function DYEY_MZ_TransStock(ByVal intOprId As Integer, ByVal strUserCode As String, ByVal strUserName As String, ByVal LngStockID As Long, ByRef strReturn As String) As Boolean
'HIS端向接口传递药品库存信息
'返回值：true-成功，false-失败
    Dim arrXML As Variant
    
    arrXML = GetXML_Stock(LngStockID)
    
    If UBound(arrXML) = -1 Then
        DYEY_MZ_TransStock = True
        Exit Function
    End If
    
    If DYEY_MZ_TransData(gType.IntStore, intOprId, strUserCode, strUserName, arrXML, strReturn) Then
        DYEY_MZ_TransStock = True
    End If
End Function

Public Function DYEY_MZ_TransRecipeDetail(ByVal intOprId As Integer, ByVal strUserCode As String, ByVal strUserName As String, ByVal LngStockID As Long, ByVal strNO As String, ByRef strReturn As String) As Boolean
'HIS端向接口传递处方明细信息
'返回值：true-成功，false-失败
    Dim arrXML As Variant

    arrXML = GetXML_RecipeDetail(LngStockID, strNO)
    
    If UBound(arrXML) = -1 Then
        DYEY_MZ_TransRecipeDetail = True
        Exit Function
    End If
    
    If DYEY_MZ_TransData(gType.IntDetail, intOprId, strUserCode, strUserName, arrXML, strReturn, strNO, LngStockID) Then
        DYEY_MZ_TransRecipeDetail = True
    End If
End Function

Public Function DYEY_MZ_TransRecipeList(ByVal intOprId As Integer, ByVal strUserCode As String, ByVal strUserName As String, ByVal LngStockID As Long, ByVal strNO As String, ByRef strReturn As String, Optional ByVal intSendType As Integer) As Boolean
'HIS端向接口传递处方主单
'intSendType:0-开始发药,1-结束发药
'返回值：true-成功，false-失败
    Dim arrXML As Variant
    
    arrXML = GetXML_RecipeList(LngStockID, strNO)
    
    If UBound(arrXML) = -1 Then
        DYEY_MZ_TransRecipeList = True
        Exit Function
    End If
    
    If DYEY_MZ_TransData(IIf(intSendType = 0, gType.IntStartList, gType.IntEndList), intOprId, strUserCode, strUserName, arrXML, strReturn) Then
        DYEY_MZ_TransRecipeList = True
    End If
End Function


Public Function DYEY_MZ_GetSendType() As Integer
'向HIS端返回发药流程方式
    DYEY_MZ_GetSendType = GINT_SEND_TYPE
End Function

Public Function DYEY_MZ_GetStartSendType() As Integer
'向HIS端返回开始发药操作方式
    DYEY_MZ_GetStartSendType = GINT_STARTSEND_TYPE
End Function



