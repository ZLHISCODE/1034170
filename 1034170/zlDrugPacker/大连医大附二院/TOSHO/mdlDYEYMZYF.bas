Attribute VB_Name = "mdlDYEYMZYF"
Option Explicit

Public gobjSOAP As Object  '接口对象
Public gstrIP As String    '本机ip
Public gblnShowMsg As Boolean   '是否弹出对话框提示（自助收费需要）
Public Const gstrUnit_DYEY = "大连医科大学附属第二医院"
Public Const gstrUnit_YZSZYY = "扬州市中医院"
Public Const gstrUnit_JLSZXYY = "吉林市中心医院"

Public Const GINT_SEND_TYPE = 1           '0-仅开始发药流程，1-有开始发药，结束发药流程
Public Const GINT_STARTSEND_TYPE = 1      '0-按钮方式开始发药，1-刷卡方式开始发药

Private Type IPINFO
    dwAddr As Long   ' IP address
    dwIndex As Long ' interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(5) As IPINFO  'array of IP address entries
End Type
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Enum gType
    IntDrug = 101 '上传药品基础数据
    IntStore = 102 '上传药品库存数据
    IntDetail = 201 '上传处方明细
    IntStartList = 202  '上传主处方单，开始发药
    IntEndList = 203    '上传主处方单，结束发药
End Enum

Private mStrSql As String

Public Function DYEY_MZ_TransData(ByVal intType As Integer, ByVal intOprId As Integer, ByVal strUserCode As String, ByVal strUserName As String, ByVal arrXML As Variant, ByRef strReturn As String, Optional ByVal strNO As String, Optional ByVal LngStockID As Long) As Boolean
'1.向WebService传递数据
'2.供接口函数调用
    Dim i As Integer
    Dim intRetval As Integer
    Dim strRetmsg As String
    Dim blnShow As Boolean
    Dim lngDrugStockID As Long
    
    On Error GoTo errHandle
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "进入上传！", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "进入上传！"
        End If
    End If
    If gstrIP = "" Then
        gstrIP = GetLocalIP
    End If
    
    For i = 0 To UBound(arrXML)
        If gobjSOAP.TransConsisData(intOprId, intType, CStr(arrXML(i)), gstrIP, strUserCode, strUserName, intRetval, strRetmsg) <> 1 Then
            If gblnShowMsg Then
                MsgBox strRetmsg, vbInformation + vbOKOnly, GSTR_MESSAGE
            Else
                strReturn = strRetmsg
            End If
            If blnShow Then frmDYEY_MZ_TransDrug.UnloadMe
            Exit Function
        End If
        
        If intType = gType.IntDrug Or intType = gType.IntStore Then
            If i = 0 Then
                frmDYEY_MZ_TransDrug.Show
                blnShow = True
            End If
            
            Call frmDYEY_MZ_TransDrug.ChangePrg(i + 1, UBound(arrXML) + 1, intType)
        ElseIf intType = gType.IntDetail Then
            lngDrugStockID = GetStockID(arrXML(i))
            If lngDrugStockID = 0 Then lngDrugStockID = 176
            'If Not SetSendWin(LngStockID, strNO, intRetval) Then
            'If Not SetSendWin(176, strNO, intRetval) Then   '暂时库房id为176''''''''''''''''''''''''''''''''''''''''''''
            If Not SetSendWin(lngDrugStockID, strNO, intRetval) Then
                If gblnShowMsg Then
                    MsgBox "调整处方的发药窗口失败！", vbCritical, GSTR_MESSAGE
                Else
                    strReturn = "调整处方的发药窗口失败！"
                End If
            End If
        End If
    Next
    
    DYEY_MZ_TransData = True
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "上传完成！", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "上传完成！"
        End If
    End If
    Exit Function
errHandle:
    If gblnShowMsg Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    End If
End Function

Public Function GetXML_Drug() As Variant
'将药品基础信息组织成指定的XML格式
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    Dim strErrMsg As String
    
    On Error GoTo errHandle
'    MsgBox "获取数据"
    strErrMsg = "获取数据"
    mStrSql = "Select Distinct a.id 药品编号, a.名称 药品名称, e.名称 药品商品名, a.规格 药品规格, a.规格 药品包装规格, b.门诊单位 药品单位," & vbNewLine & _
              "    round(b.药库包装/b.门诊包装, 2) 包装比,b.门诊可否分零,a.产地 药品厂家, c.现价 * b.门诊包装 药品价格, d.药品剂型, " & vbNewLine & _
              "    b.门诊包装, a.建档时间 最后更新时间, f.简码 药品拼音, d.毒理分类 " & vbNewLine & _
              "From 收费项目目录 a, 药品规格 b, 收费价目 c, 药品特性 d, 收费项目别名 e, 收费项目别名 f " & vbNewLine & _
              "Where a.Id = b.药品id And a.Id = c.收费细目id And b.药名id = d.药名id And a.Id = e.收费细目id(+) And a.Id = f.收费细目id(+) And " & vbNewLine & _
              "    (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And Sysdate Between c.执行日期 And " & vbNewLine & _
              "    Nvl(c.终止日期, Sysdate) And e.性质(+) = 3 And f.性质(+) = 1 And f.码类(+) = 1"
    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_Drug")
    strErrMsg = "数据获取完毕"
    strXML = ""
    arrXML = Array()
    
    strErrMsg = "XML开始"
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_BASIC_DRUGSVW"
                strDrug = strDrug & vbCrLf & "DRUG_CODE = """ & SpecialChar(!药品编号) & """"
                strDrug = strDrug & vbCrLf & "DRUG_NAME = """ & SpecialChar(!药品名称) & """"
                strDrug = strDrug & vbCrLf & "TRADE_NAME = """ & SpecialChar(!药品商品名) & """"
                strDrug = strDrug & vbCrLf & "DRUG_SPEC = """ & SpecialChar(!药品规格) & """"
                strDrug = strDrug & vbCrLf & "DRUG_PACKAGE = """ & NVL(!门诊包装) & """"  ' & SpecialChar(!药品包装规格) & """"
                strDrug = strDrug & vbCrLf & "DRUG_UNIT = """ & SpecialChar(!药品单位) & """"
                strDrug = strDrug & vbCrLf & "FIRM_ID = """ & SpecialChar(!药品厂家) & """"
                strDrug = strDrug & vbCrLf & "DRUG_PRICE = """ & NVL(!药品价格) & """"
                strDrug = strDrug & vbCrLf & "DRUG_FORM = """ & SpecialChar(!药品剂型) & """"
                strDrug = strDrug & vbCrLf & "DRUG_SORT = """ & SpecialChar(!毒理分类) & """"
                strDrug = strDrug & vbCrLf & "BARCODE = """""
                strDrug = strDrug & vbCrLf & "LAST_DATE = """ & Format(!最后更新时间, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PINYIN = """ & SpecialChar(!药品拼音) & """"
                strDrug = strDrug & vbCrLf & "DRUG_CONVERTATION = """ & NVL(!包装比) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_BASIC_DRUGSVW>"
                
                If Len(strXML & strDrug) > 3900 Then
                    '将以前的添加到数组
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    strErrMsg = "装入数据1"
                    '重新拼凑新的XML
                    strXML = strTitle & vbCrLf & strDrug
                Else
                    strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                End If
                
                rsTemp.MoveNext
                If .EOF And strXML <> "" Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    strErrMsg = "装入数据2"
                End If
            Loop
        End If
    End With
    
    strErrMsg = "获取数据"
    GetXML_Drug = arrXML
    strErrMsg = "返回数据"
    Exit Function

errHandle:
    Debug.Print strErrMsg
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetXML_RecipeDetail(ByVal LngStockID As Long, ByVal strNO As String) As Variant
'将处方明细组织成指定的XML格式
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    Dim strSQL As String
    Dim i As Integer
    Dim rsDetails As Recordset
    Dim strDetail As String
    
    '暂时库房为176''''''''''''''''''''''''''''''''''''''''''''''''
'    LngStockID = 176
    
    On Error GoTo errHandle
    '获取处方单信息
    strSQL = "Select a.填制日期 处方时间, a.单据, a.No 处方编号, a.库房id 发药药局, c.病人id 就诊卡号, a.姓名 患者姓名, Decode(a.优先级, 1, '01', '00') 患者类型, " & vbNewLine & _
             "    c.出生日期 患者出生日期, c.性别 患者性别, c.身份 患者身份, c.医疗付款方式 医保类型, Sum(d.应收金额) 费用, Sum(d.实收金额) 实付费用," & vbNewLine & _
             "    f.id 开单科室, d.开单人 开方医生, d.开单人 录方人, Decode(a.优先级, 1, '1', '2') 配药优先级 " & vbNewLine & _
             "From 未发药品记录 a, 病人信息 c, 门诊费用记录 d, 药品收发记录 e, 部门表 f " & vbNewLine & _
             "Where a.单据 = e.单据 And a.No = e.No And a.库房id = e.库房id And a.病人id = c.病人id And e.费用id = d.Id And " & vbNewLine & _
             "    d.开单部门id = f.Id " & IIf(LngStockID = 0, "", " And a.库房id=[1] ")

    If InStr(1, strNO, "|") < 1 Then
        strSQL = strSQL & " And a.单据=[2] And a.NO=[3] "
    Else
        strSQL = strSQL & " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strSQL = strSQL & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strSQL = strSQL & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strSQL = strSQL & ") "
    End If
            
    strSQL = strSQL & _
             "Group By a.填制日期, a.单据, a.No, a.库房id, c.病人id, a.姓名, Decode(a.优先级, 1, '01', '00'), c.出生日期, c.性别, " & vbNewLine & _
             "    c.身份, c.医疗付款方式, f.id, d.开单人,d.开单人, Decode(a.优先级, 1, '1', '2') "
             

    mStrSql = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "门诊费用记录", "住院费用记录")
    mStrSql = "select * from (" & mStrSql & ") Order By 发药药局, 就诊卡号 "
    
    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_RecipeDetail", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    
    
    '获取处方明细信息
    strSQL = "Select Distinct a.填制日期, a.单据, a.No, a.序号, b.id 药品编码, b.名称 药品名称, c.名称 药品商品名, b.规格 药品规格, b.规格 药品包装规格, " & vbNewLine & _
             "    d.门诊单位 药品单位, a.产地 药品厂家, a.零售价 * d.门诊包装 药品价格, a.实际数量 / d.门诊包装 数量, e.应收金额 费用,e.病人id," & vbNewLine & _
             "    e.实收金额 实付费用, a.单量 药品剂量, a.库房id, a.用法, f.执行频次, g.计算单位 剂量单位 " & vbNewLine & _
             "From 药品收发记录 a, 收费项目目录 b, 收费项目别名 c, 药品规格 d, 门诊费用记录 e, 病人医嘱记录 f, 诊疗项目目录 g " & vbNewLine & _
             "Where a.药品id = b.Id And a.药品id = c.收费细目id(+) And a.药品id = d.药品id And a.费用id = e.Id and d.药名id=g.id " & vbNewLine & _
             "    And e.医嘱序号 = f.Id(+) And c.性质(+) = 3 " & IIf(LngStockID = 0, "", " And a.库房id=[1] ")

    If InStr(1, strNO, "|") < 1 Then
        strSQL = strSQL & " And a.单据=[2] And a.NO=[3] "
    Else
        strSQL = strSQL & " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strSQL = strSQL & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strSQL = strSQL & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strSQL = strSQL & ") "
    End If
    
    mStrSql = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "门诊费用记录", "住院费用记录")
    Set rsDetails = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_RecipeDetail", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    strXML = ""
    arrXML = Array()
    
    '库房ID为0的情况单独函数处理
    If LngStockID = 0 Then
        If GetXML_RecipeDetailEx(rsTemp, rsDetails, arrXML) Then
            GetXML_RecipeDetail = arrXML
        End If
        Exit Function
    End If
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW"
                strDrug = strDrug & vbCrLf & "PRESC_DATE = """ & Format(!处方时间, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PRESC_NO = """ & SpecialChar(!处方编号) & """"
                strDrug = strDrug & vbCrLf & "DISPENSARY = """ & NVL(!发药药局) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_ID = """ & NVL(!就诊卡号) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_NAME = """ & SpecialChar(!患者姓名) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_TYPE = """ & NVL(!患者类型) & """"
                strDrug = strDrug & vbCrLf & "DATE_OF_BIRTH = """ & Format(NVL(!患者出生日期), "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "SEX = """ & SpecialChar(!患者性别) & """"
                strDrug = strDrug & vbCrLf & "PRESC_IDENTITY = """ & SpecialChar(!患者身份) & """"
                strDrug = strDrug & vbCrLf & "CHARGE_TYPE = """ & SpecialChar(!医保类型) & """"
                strDrug = strDrug & vbCrLf & "PRESC_ATTR = """""
                strDrug = strDrug & vbCrLf & "PRESC_INFO = """""
                strDrug = strDrug & vbCrLf & "RCPT_INFO = " & GetRCPT_INFO(NVL(!处方编号))
                strDrug = strDrug & vbCrLf & "RCPT_REMARK = """""
                strDrug = strDrug & vbCrLf & "REPETITION = ""1"""
                strDrug = strDrug & vbCrLf & "COSTS = """ & NVL(!费用) & """"
                strDrug = strDrug & vbCrLf & "PAYMENTS = """ & NVL(!实付费用) & """"
                strDrug = strDrug & vbCrLf & "ORDERED_BY = """ & NVL(!开单科室) & """"
                strDrug = strDrug & vbCrLf & "PRESCRIBED_BY = """ & SpecialChar(!开方医生) & """"
                strDrug = strDrug & vbCrLf & "ENTERED_BY = """ & SpecialChar(!录方人) & """"
                strDrug = strDrug & vbCrLf & "DISPENSE_PRI = """ & NVL(!配药优先级) & """"
                strDrug = strDrug & vbCrLf & ">"
                
                '过滤明细记录，确保与单据对应
                rsDetails.Filter = "no='" & !处方编号 & "' and 单据=" & NVL(!单据) & " and 库房id=" & NVL(!发药药局)
                rsDetails.Sort = "序号"
'                rsDetails.Filter = "no='" & !处方编号 & "' and 填制日期='" & CDate(!处方时间) & "'"
                
                strDetail = ""
                Do While Not rsDetails.EOF
                    strDetail = strDetail & vbCrLf & "<CONSIS_PRESC_DTLVW"
                    strDetail = strDetail & vbCrLf & "PRESC_DATE = """ & Format(rsDetails!填制日期, "yyyy-MM-DDThh:mm:ss") & """"
                    strDetail = strDetail & vbCrLf & "PRESC_NO = """ & NVL(rsDetails!no) & """"
                    strDetail = strDetail & vbCrLf & "ITEM_NO = """ & NVL(rsDetails!序号) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_CODE = """ & SpecialChar(rsDetails!药品编码) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_NAME = """ & SpecialChar(rsDetails!药品名称) & """"
                    strDetail = strDetail & vbCrLf & "TRADE_NAME = """ & SpecialChar(rsDetails!药品商品名) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_SPEC= """ & SpecialChar(rsDetails!药品规格) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_PACKAGE = """ & SpecialChar(rsDetails!药品包装规格) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_UNIT = """ & SpecialChar(rsDetails!药品单位) & """"
                    strDetail = strDetail & vbCrLf & "FIRM_ID = """ & SpecialChar(rsDetails!药品厂家) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_PRICE = """ & NVL(rsDetails!药品价格) & """"
                    strDetail = strDetail & vbCrLf & "QUANTITY = """ & NVL(rsDetails!数量) & """"
                    strDetail = strDetail & vbCrLf & "COSTS = """ & NVL(rsDetails!费用) & """"
                    strDetail = strDetail & vbCrLf & "PAYMENTS = """ & NVL(rsDetails!实付费用) & """"
                    strDetail = strDetail & vbCrLf & "DOSAGE = """ & NVL(rsDetails!药品剂量) & """"
                    strDetail = strDetail & vbCrLf & "DOSAGE_UNITS = """ & SpecialChar(rsDetails!剂量单位) & """"
                    strDetail = strDetail & vbCrLf & "ADMINISTRATION = """ & SpecialChar(rsDetails!用法) & """"
                    strDetail = strDetail & vbCrLf & "FREQUENCY = """ & SpecialChar(rsDetails!执行频次) & """"
                    strDetail = strDetail & vbCrLf & ">"
                    strDetail = strDetail & vbCrLf & "</CONSIS_PRESC_DTLVW>"
                    rsDetails.MoveNext
                Loop
                strDrug = strDrug & strDetail
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeDetail = arrXML
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function GetXML_RecipeDetailEx(ByVal rsBill As ADODB.Recordset, ByVal rsDetail As ADODB.Recordset, ByRef varXML As Variant) As Boolean
'功能：处理库房ID为0的情况，分处理库房ID与病人ID生成XML字符串
'参数：
'  rsBill：单据数据集；
'  rsDetail：明细数据集；
'  varXML：生成的XML字符串数组（实参）。
'返回：True成功   False失败
    Const STR_ROOT_BEGIN = "<ROOT>"
    Const STR_ROOT_END = "</ROOT>"
    Const STR_BILL = "CONSIS_PRESC_MSTVW"
    Const STR_DETAIL = "CONSIS_PRESC_DTLVW"
    Dim strXML As String, strBill As String, strDetail As String
    Dim lng库房ID As Long, lng病人ID As Long
    Dim varReturn As Variant
    
    On Error GoTo errHandle
    varReturn = Array()
    With rsBill
        If .RecordCount <= 0 Then Exit Function
        .MoveFirst
        lng库房ID = NVL(!发药药局, 0)
        lng病人ID = NVL(!就诊卡号, 0)
        Do
            If .EOF Then Exit Do
            '单据
            strBill = "<" & STR_BILL & " "
            strBill = strBill & vbCrLf & "PRESC_DATE = """ & Format(!处方时间, "yyyy-MM-DDThh:mm:ss") & """"
            strBill = strBill & vbCrLf & "PRESC_NO = """ & SpecialChar(!处方编号) & """"
            strBill = strBill & vbCrLf & "DISPENSARY = """ & NVL(!发药药局) & """"
            strBill = strBill & vbCrLf & "PATIENT_ID = """ & NVL(!就诊卡号) & """"
            strBill = strBill & vbCrLf & "PATIENT_NAME = """ & SpecialChar(!患者姓名) & """"
            strBill = strBill & vbCrLf & "PATIENT_TYPE = """ & NVL(!患者类型) & """"
            strBill = strBill & vbCrLf & "DATE_OF_BIRTH = """ & Format(NVL(!患者出生日期), "yyyy-MM-DDThh:mm:ss") & """"
            strBill = strBill & vbCrLf & "SEX = """ & SpecialChar(!患者性别) & """"
            strBill = strBill & vbCrLf & "PRESC_IDENTITY = """ & SpecialChar(!患者身份) & """"
            strBill = strBill & vbCrLf & "CHARGE_TYPE = """ & SpecialChar(!医保类型) & """"
            strBill = strBill & vbCrLf & "PRESC_ATTR = """""
            strBill = strBill & vbCrLf & "PRESC_INFO = """""
            strBill = strBill & vbCrLf & "RCPT_INFO = " & GetRCPT_INFO(NVL(!处方编号))
            strBill = strBill & vbCrLf & "RCPT_REMARK = """""
            strBill = strBill & vbCrLf & "REPETITION = ""1"""
            strBill = strBill & vbCrLf & "COSTS = """ & NVL(!费用) & """"
            strBill = strBill & vbCrLf & "PAYMENTS = """ & NVL(!实付费用) & """"
            strBill = strBill & vbCrLf & "ORDERED_BY = """ & NVL(!开单科室) & """"
            strBill = strBill & vbCrLf & "PRESCRIBED_BY = """ & SpecialChar(!开方医生) & """"
            strBill = strBill & vbCrLf & "ENTERED_BY = """ & SpecialChar(!录方人) & """"
            strBill = strBill & vbCrLf & "DISPENSE_PRI = """ & NVL(!配药优先级) & """"
            strBill = strBill & vbCrLf & ">"
            
            '过滤明细记录，确保与单据对应
            strDetail = ""
            rsDetail.Filter = "no='" & !处方编号 & "' and 单据=" & NVL(!单据) & " and 库房id=" & NVL(!发药药局) & " and 病人id=" & NVL(!就诊卡号)
            rsDetail.Sort = "序号"
            Do
                If rsDetail.EOF Then Exit Do
                '明细
                strDetail = strDetail & vbCrLf & "<" & STR_DETAIL & " "
                strDetail = strDetail & vbCrLf & "PRESC_DATE = """ & Format(rsDetail!填制日期, "yyyy-MM-DDThh:mm:ss") & """"
                strDetail = strDetail & vbCrLf & "PRESC_NO = """ & NVL(rsDetail!no) & """"
                strDetail = strDetail & vbCrLf & "ITEM_NO = """ & rsDetail!序号 & """"
                strDetail = strDetail & vbCrLf & "DRUG_CODE = """ & SpecialChar(rsDetail!药品编码) & """"
                strDetail = strDetail & vbCrLf & "DRUG_NAME = """ & SpecialChar(rsDetail!药品名称) & """"
                strDetail = strDetail & vbCrLf & "TRADE_NAME = """ & SpecialChar(rsDetail!药品商品名) & """"
                strDetail = strDetail & vbCrLf & "DRUG_SPEC= """ & SpecialChar(rsDetail!药品规格) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PACKAGE = """ & SpecialChar(rsDetail!药品包装规格) & """"
                strDetail = strDetail & vbCrLf & "DRUG_UNIT = """ & SpecialChar(rsDetail!药品单位) & """"
                strDetail = strDetail & vbCrLf & "FIRM_ID = """ & SpecialChar(rsDetail!药品厂家) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PRICE = """ & NVL(rsDetail!药品价格) & """"
                strDetail = strDetail & vbCrLf & "QUANTITY = """ & NVL(rsDetail!数量) & """"
                strDetail = strDetail & vbCrLf & "COSTS = """ & NVL(rsDetail!费用) & """"
                strDetail = strDetail & vbCrLf & "PAYMENTS = """ & NVL(rsDetail!实付费用) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE = """ & NVL(rsDetail!药品剂量) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE_UNITS = """ & SpecialChar(rsDetail!剂量单位) & """"
                strDetail = strDetail & vbCrLf & "ADMINISTRATION = """ & SpecialChar(rsDetail!用法) & """"
                strDetail = strDetail & vbCrLf & "FREQUENCY = """ & SpecialChar(rsDetail!执行频次) & """"
                strDetail = strDetail & vbCrLf & ">"
                strDetail = strDetail & vbCrLf & "</" & STR_DETAIL & ">"
                rsDetail.MoveNext
            Loop While Not rsDetail.EOF
            
            strBill = strBill & strDetail
            strBill = strBill & "</" & STR_BILL & ">"
            
            '拆分不同库房ID和病人ID的单据明细
            If lng库房ID = NVL(!发药药局, 0) And lng病人ID = NVL(!就诊卡号, 0) Then
                strXML = strXML & strBill & vbCrLf
            Else
                strXML = STR_ROOT_BEGIN & vbCrLf & strXML & STR_ROOT_END
                ReDim Preserve varReturn(UBound(varReturn) + 1)
                varReturn(UBound(varReturn)) = strXML
                strXML = strBill & vbCrLf
            End If
            
            lng库房ID = NVL(!发药药局, 0)
            lng病人ID = NVL(!就诊卡号, 0)
            
            .MoveNext
        Loop While Not .EOF
        
        strXML = STR_ROOT_BEGIN & vbCrLf & strXML & STR_ROOT_END
        ReDim Preserve varReturn(UBound(varReturn) + 1)
        varReturn(UBound(varReturn)) = strXML
        varXML = varReturn
        GetXML_RecipeDetailEx = True
    
    End With
    
    Exit Function
    
errHandle:
    Set varXML = Nothing
End Function

Public Function GetXML_RecipeList(ByVal LngStockID As Long, ByVal strNO As String) As Variant
'将处方单组织成指定的XML格式
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant, arrTmp As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    mStrSql = "Select 填制日期,No From 药品收发记录 Where 库房id=[1]"
    
    If InStr(1, strNO, "|") < 1 Then
        mStrSql = mStrSql & " And 单据=[2] And NO=[3]"
    Else
        mStrSql = mStrSql & " And ("
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(arrTmp)
            If i = UBound(arrTmp) Then
                mStrSql = mStrSql & "(单据=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "')"
            Else
                mStrSql = mStrSql & "(单据=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "') or "
            End If
        Next
        mStrSql = mStrSql & ")"
    End If
    mStrSql = mStrSql & " and (记录状态=1 or mod(记录状态,3)=1) "
    
    If InStr(1, strNO, "|") < 1 Then
        Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_RecipeList", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_RecipeList", LngStockID)
    End If
    
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW"
                strDrug = strDrug & vbCrLf & "PRESC_DATE = """ & Format(!填制日期, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PRESC_NO = """ & NVL(!no) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeList = arrXML
    Exit Function
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetXML_Stock(ByVal LngStockID As Long) As Variant
'将药品库存信息组织成指定的XML格式
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    
    On Error GoTo errHandle
    mStrSql = "Select a.id 药品编号,c.库房id 发药药局,sum(c.实际数量/e.门诊包装) 药品数量,d.库房货位 药品货位 " & vbNewLine & _
              "From 收费项目目录 a, 药品库存 c, 药品储备限额 d,药品规格 e " & vbNewLine & _
              "Where a.Id = c.药品id And e.药品id=c.药品id And d.库房id(+) = c.库房id And d.药品id(+) = c.药品id And c.库房id=[1] " & vbNewLine & _
              "Group By a.id, c.库房id, d.库房货位 " & vbNewLine & _
              "Having Sum(c.实际数量/e.门诊包装)<>0 "

    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "GetXML_Stock", LngStockID)
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PHC_STORAGEVW"
                strDrug = strDrug & vbCrLf & "DRUG_CODE = """ & SpecialChar(!药品编号) & """"
                strDrug = strDrug & vbCrLf & "DISPENSARY = """ & NVL(!发药药局) & """"
                strDrug = strDrug & vbCrLf & "DRUG_QUANTITY = """ & NVL(!药品数量) & """"
                strDrug = strDrug & vbCrLf & "LOCATIONINFO = """ & SpecialChar(!药品货位) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_PHC_STORAGEVW>"

'该业务功能可以不用4K限制
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                
'                If Len(strXML & strDrug) > 3900 Then
'                    '将以前的添加到数组
'                    strXML = strXML & vbCrLf & "</ROOT>"
'                    ReDim Preserve arrXML(UBound(arrXML) + 1)
'                    arrXML(UBound(arrXML)) = strXML
'
'                    '重新拼凑新的XML
'                    strXML = strTitle & vbCrLf & strDrug
'                Else
'                    strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
'                End If
                
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
        End If
    End With
    
    GetXML_Stock = arrXML
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function SetSendWin(ByVal LngStockID As Long, ByVal strNO As String, ByVal intOpr As Integer) As Boolean
'设置HIS中指定处方的发药窗口
    Dim i As Integer
    Dim arrTmp As Variant
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    mStrSql = "Select 名称 From 发药窗口 Where 药房id=[1] And 编码=[2]"
    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(mStrSql, "SetSendWin", LngStockID, CStr(intOpr))
    
    If Not rsTemp.EOF Then
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(Split(strNO, "|"))
            mStrSql = "Zl_未发药品记录_分配发药窗口("
'            mStrSql = mStrSql & "'" & Split(Split(strNO, "|"), ",")(1) & "',"
'            mStrSql = mStrSql & Split(Split(strNO, "|"), ",")(0) & ","
            mStrSql = mStrSql & "'" & Split(arrTmp(i), ",")(1) & "',"
            mStrSql = mStrSql & Split(arrTmp(i), ",")(0) & ","
            mStrSql = mStrSql & LngStockID & ","
            mStrSql = mStrSql & "'" & rsTemp!名称 & "')"
            Call gobjComLib.zldatabase.ExecuteProcedure(mStrSql, "SetSendWin")
        Next
        SetSendWin = True
    Else
        If gblnShowMsg Then
            MsgBox "没有找到编码为【" & intOpr & "】的窗口，请检查！", vbCritical, GSTR_MESSAGE
        End If
    End If
    
    Exit Function
errHandle:
    If gblnShowMsg Then
        If gobjComLib.ErrCenter() = 1 Then Resume
        Call gobjComLib.SaveErrLog
    End If
End Function


Public Function GetLocalIP() As String
'取本机IP
    Dim Ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    
    
    On Error GoTo EndRow
        GetIpAddrTable ByVal 0&, Ret, True
    
    
        If Ret <= 0 Then Exit Function
        ReDim bBytes(0 To Ret - 1) As Byte
        ReDim TempList(0 To Ret - 1) As String
        
        'retrieve the data
        GetIpAddrTable bBytes(0), Ret, False
          
        'Get the first 4 bytes to get the entry's.. ip installed
        CopyMemory Listing.dEntrys, bBytes(0), 4
        
        For Tel = 0 To Listing.dEntrys - 1
            'Copy whole structure to Listing..
            CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
            TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
        Next Tel
        'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        GetLocalIP = TempIP 'Return The TempIP
    Exit Function
EndRow:
    GetLocalIP = ""
End Function

Private Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Private Function GetRCPT_INFO(ByVal strNO As String) As String
'功能：获取诊断信息
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select MAX(DECODE(Id,1,诊断描述,''))||';'||MAX(DECODE(Id,2,诊断描述,'')) as 诊断 " & vbNewLine & _
             "From ( " & vbNewLine & _
             "      Select Rownum As Id,诊断描述 " & vbNewLine & _
             "      From (Select 诊断描述||decode(是否疑诊,1,'?','') 诊断描述 " & vbNewLine & _
             "            From 病人诊断记录 " & vbNewLine & _
             "            Where 病人id=(Select distinct 病人id " & vbNewLine & _
             "                          From ( Select a.病人id From 门诊费用记录 a Left Join 病人医嘱记录 b On a.医嘱序号=b.Id " & vbNewLine & _
             "                                 Where a.No=[1] And 记录性质=1 ) ) " & vbNewLine & _
             "              And 主页id=(Select distinct Case When 主页id Is Null Then (Select Id From 病人挂号记录 Where No=c.挂号单) Else 主页Id End As 主页id " & vbNewLine & _
             "                          From ( Select null 主页id, b.挂号单 From 门诊费用记录 a Left Join 病人医嘱记录 b On a.医嘱序号=b.Id " & vbNewLine & _
             "                                 Where a.No=[1] And 记录性质=1 ) c ) " & vbNewLine & _
             "union all " & vbNewLine & _
             "Select a.摘要 As 诊断描述 From 病人挂号记录 a " & vbNewLine & _
             "Where No= (Select distinct Case When b.挂号单 Is Null Then ' ' Else b.挂号单 End As No " & vbNewLine & _
             "           From 门诊费用记录 a Left Join 病人医嘱记录 b On a.医嘱序号 = b.Id " & vbNewLine & _
             "           Where a.No = [1] And 记录性质 = 1 ) ) ) "
    On Error GoTo errHandle
    Set rsTemp = gobjComLib.zldatabase.OpenSQLRecord(strSQL, "获取诊断信息", strNO)
    If Not rsTemp.EOF Then
        GetRCPT_INFO = IIf(Trim(NVL(rsTemp!诊断)) = ";", """""", """" & Trim(NVL(rsTemp!诊断)) & """")
    Else
        GetRCPT_INFO = """"""
    End If
    rsTemp.Close
    Exit Function
    
errHandle:
    GetRCPT_INFO = """"""
End Function

Private Function SpecialChar(ByVal strVal As Variant) As String
'功能：特殊字符转换
'说明：
' < 转 &lt;
' > 转 &gt;
' & 转 &amp;
' ' 转 &apos;
' " 转 &quot;
    Dim strReturn As String
    
    If IsNull(strVal) Then
        strVal = ""
        GoTo errHandle
    End If
    If strVal = "" Then
        GoTo errHandle
    End If
    On Error GoTo errHandle
    strReturn = strVal
    strReturn = Replace(strReturn, "<", "&lt;")
    strReturn = Replace(strReturn, ">", "&gt;")
    strReturn = Replace(strReturn, "&", "&amp;")
    strReturn = Replace(strReturn, "'", "&apos;")
    strReturn = Replace(strReturn, """", "&quot;")
    SpecialChar = strReturn
    Exit Function
    
errHandle:
    SpecialChar = strVal
End Function

Private Function GetStockID(ByVal strText As String) As Long
'功能：获取XML文本中的药房ID
    Const STR_KEY = "DISPENSARY = "
    Dim LngStockID As Long
    Dim intStart As Integer
    
    If strText = "" Then Exit Function
    
    intStart = InStr(strText, STR_KEY)
    If intStart > 0 Then
        LngStockID = Val(Mid(strText, intStart + Len(STR_KEY) + 1))
    End If
    GetStockID = LngStockID
    
End Function

