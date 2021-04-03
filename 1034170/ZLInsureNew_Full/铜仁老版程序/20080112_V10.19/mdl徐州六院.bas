Attribute VB_Name = "mdl徐州六院"
Option Explicit

Public Function 门诊虚拟结算_徐州六院(rs明细 As ADODB.Recordset, str结算方式 As String, Optional ByRef strAdvance As String = "") As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim rs算法 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim cls医保 As New clsInsure
    Dim rs大类汇总 As New ADODB.Recordset
    Dim dbl全自费 As Currency, dbl首先自付 As Currency, dbl进入统筹 As Currency, dblTemp As Double
    Dim dbl最大金额 As Double, cur诊疗报销 As Currency, cur药品报销 As Currency
    Dim dbl个人帐户 As Double, cur报销额 As Currency
    Dim lng病人ID As Long, cur总额 As Currency
    Dim rs特准项目 As New ADODB.Recordset
    Dim dblTemp1 As Double, datCurr As Date
    
    datCurr = zlDatabase.Currentdate
    If rs明细.RecordCount > 0 Then
        rs明细.MoveFirst
        lng病人ID = rs明细("病人ID")
    End If
    cur总额 = 0: cur药品报销 = 0: cur诊疗报销 = 0
    While Not rs明细.EOF
        gstrSQL = "select a.名称,b.名称,b.统筹比额,b.算法,a.类别 from 收费细目 a,保险支付大类 b,保险支付项目 c where " & _
            "b.id=c.大类id and a.id=c.收费细目id and c.险类=" & TYPE_徐州六院 & " and a.id=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        If rsTemp.EOF Then
            dbl全自费 = dbl全自费 + rs明细!实收金额
            cur报销额 = 0
        ElseIf rsTemp!算法 = 1 Then
            dbl全自费 = dbl全自费 + rs明细!实收金额 * (1 - rsTemp!统筹比额 / 100)
            dbl进入统筹 = dbl进入统筹 + rs明细!实收金额 * rsTemp!统筹比额 / 100
            cur报销额 = rs明细!实收金额 * rsTemp!统筹比额 / 100
        ElseIf rsTemp!算法 = 2 Then
            dbl进入统筹 = dbl进入统筹 + IIf(rs明细!实收金额 < rsTemp!统筹比额, rs明细!实收金额, rsTemp!统筹比额)
            dbl全自费 = dbl全自费 + IIf(rs明细!实收金额 < rsTemp!统筹比额, 0, rs明细!实收金额 - rsTemp!统筹比额)
            cur报销额 = IIf(rs明细!实收金额 < rsTemp!统筹比额, rs明细!实收金额, rsTemp!统筹比额)
        End If
        If rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7" Then
            cur药品报销 = cur药品报销 + cur报销额
        Else
            cur诊疗报销 = cur诊疗报销 + cur报销额
        End If
        
        cur总额 = cur总额 + Nvl(rs明细!实收金额, 0)
        rs明细.MoveNext
    Wend
    g结算数据.发生费用金额 = cur总额
'    dblTemp = dbl进入统筹
    
    '每天报销金额不高于80
'    gstrSQL = "Select nvl(sum(a.统筹报销金额),0) From 保险结算记录 a,病人费用记录 b Where a.记录ID=b.结帐id " & _
'        "and to_char(b.发生时间,'yyyy-mm-dd')='" & _
'        Format(datCurr, "yyyy-mm-dd") & "' And a.性质=1 And b.病人ID=" & lng病人id & " And a.险类=" & TYPE_徐州六院
'    Call OpenRecordset(rsTemp, gstrSysName)
'    If dblTemp + rsTemp(0) > 80 Then dblTemp = 80 - rsTemp(0)
    '每张单据药品报销不超80元
    
    '20051220 陈东 徐州丰县加
    Dim cur限额 As Currency, bln超额禁止 As Boolean
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=" & TYPE_徐州六院 & " And 参数名='门诊限额'"
    Call OpenRecordset(rsTemp, "保险参数")
    If rsTemp.EOF Then
        cur限额 = 80
    Else
        If Val(rsTemp!参数值) > 0 Then
            cur限额 = Val(rsTemp!参数值)
        Else
            cur限额 = 80
        End If
    End If
    
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=" & TYPE_徐州六院 & " And 参数名='超额禁止'"
    Call OpenRecordset(rsTemp, "保险参数")
    If rsTemp.EOF Then
        bln超额禁止 = False
    Else
        If Val(rsTemp!参数值) = 1 Then
            bln超额禁止 = True
        Else
            bln超额禁止 = False
        End If
    End If
    If cur药品报销 > cur限额 Then
        If bln超额禁止 = True Then
            MsgBox "已超过处方限额" & Format(cur限额, "0.00") & "，不能收费！", vbInformation, gstrSysName
            门诊虚拟结算_徐州六院 = False
            Exit Function
        Else
            dblTemp = cur诊疗报销 + cur限额
            dbl全自费 = dbl全自费 + cur药品报销 - cur限额
        End If
    Else
        dblTemp = cur药品报销 + cur诊疗报销
    End If
    
    g结算数据.进入统筹金额 = dbl进入统筹
    g结算数据.全自费金额 = dbl全自费
    g结算数据.首先自付金额 = 0
    g结算数据.统筹报销金额 = dblTemp
    str结算方式 = "离休医疗;" & dblTemp & ";0"
   
    门诊虚拟结算_徐州六院 = True
End Function

Public Function 住院虚拟结算_徐州六院(rs明细 As ADODB.Recordset) As String
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim rs算法 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim cls医保 As New clsInsure
    Dim rs大类汇总 As New ADODB.Recordset
    Dim dbl全自费 As Currency, dbl首先自付 As Currency, dbl进入统筹 As Currency, dblTemp As Double
    Dim dbl最大金额 As Double, lng中心 As Long, lng在职 As Long, lng年龄 As Long
    Dim dbl个人帐户 As Double, lng年龄段  As Long, bln全额统筹 As Boolean, bln无起付线 As Boolean, bln无封顶线 As Boolean
    Dim lng病人ID As Long, cur全额 As Currency, dbl统筹 As Currency
    Dim rs特准项目 As New ADODB.Recordset, strTemp As String, str类型 As String
    Dim dblTemp1 As Double, datCurr As Date

    datCurr = zlDatabase.Currentdate
    If rs明细.RecordCount > 0 Then
        rs明细.MoveFirst
        lng病人ID = rs明细("病人ID")
    End If
    gstrSQL = "Select max(主页id) from 病案主页 Where 病人id=" & lng病人ID
    Call OpenRecordset(rsTemp, gstrSysName)
    g结算数据.主页ID = Nvl(rsTemp(0), 1)
    g结算数据.病人ID = lng病人ID
    g结算数据.年度 = Format(datCurr, "yyyy")
    
    gstrSQL = "select 入院日期,nvl(出院日期,to_date('3000-01-01','yyyy-MM-dd')) as 出院日期 " & _
              "from 病案主页 where 病人ID=" & g结算数据.病人ID & " and 主页ID=" & g结算数据.主页ID
    Call OpenRecordset(rsTemp, "虚拟结算")
    If rsTemp("出院日期") = CDate("3000-01-01") Then
        g结算数据.中途结帐 = 1
    Else
        '表示该病人已经出院
        g结算数据.中途结帐 = 0
    End If
    
    With g结算数据
        gstrSQL = "select A.中心,A.人员身份,A.在职,A.年龄段," & _
                  "      B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计" & _
                  " from 保险帐户 A,帐户年度信息 B" & _
                  " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) " & _
                  "     and B.年度(+)=" & .年度 & " and A.病人ID=" & .病人ID & " and A.险类=" & TYPE_徐州六院
        Call OpenRecordset(rsTemp, "虚拟结算")
        
        lng中心 = IIf(IsNull(rsTemp("中心")), 0, rsTemp("中心"))
        lng在职 = IIf(IsNull(rsTemp("在职")), 1, rsTemp("在职"))
        lng年龄 = IIf(IsNull(rsTemp("年龄段")), 0, rsTemp("年龄段"))
        .住院次数 = IIf(IsNull(rsTemp("住院次数累计")), 0, rsTemp("住院次数累计"))
        .帐户累计增加 = IIf(IsNull(rsTemp("帐户增加累计")), 0, rsTemp("帐户增加累计"))
        .帐户累计支出 = IIf(IsNull(rsTemp("帐户支出累计")), 0, rsTemp("帐户支出累计"))
        .累计进入统筹 = IIf(IsNull(rsTemp("进入统筹累计")), 0, rsTemp("进入统筹累计"))
        .累计统筹报销 = IIf(IsNull(rsTemp("统筹报销累计")), 0, rsTemp("统筹报销累计"))
    
        
        gstrSQL = "select 年龄段,nvl(全额统筹,0) as 全额统筹 ,nvl(无起付线,0) as 无起付线 ,nvl(无封顶线,0) as 无封顶线 " & _
                " from 保险年龄段" & _
                " where 险类=" & TYPE_徐州六院 & " and nvl(中心,0)=" & lng中心 & _
                "       and 在职=" & lng在职 & " and 下限<=" & lng年龄 & " and (" & lng年龄 & "<=上限 or 上限=0)"
        Call OpenRecordset(rsTemp, "虚拟结算")
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在“保险类别管理”中设置年龄段与费用档。", vbInformation, gstrSysName
            Exit Function
        End If
        lng年龄段 = rsTemp("年龄段")
        bln全额统筹 = (rsTemp("全额统筹") = 1)
        bln无起付线 = (rsTemp("无起付线") = 1)
        bln无封顶线 = (rsTemp("无封顶线") = 1)
    End With
    
    While Not rs明细.EOF
        gstrSQL = "select a.名称,b.名称,b.统筹比额,b.算法 from 收费细目 a,保险支付大类 b,保险支付项目 c where " & _
            "b.id=c.大类id and a.id=c.收费细目id and c.险类=" & TYPE_徐州六院 & " and a.id=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        If rsTemp.EOF Then
            dbl全自费 = dbl全自费 + rs明细!金额
        ElseIf rsTemp!算法 = 1 Then
            dbl全自费 = dbl全自费 + rs明细!金额 * (1 - rsTemp!统筹比额 / 100)
            dbl进入统筹 = dbl进入统筹 + rs明细!金额 * rsTemp!统筹比额 / 100
        ElseIf rsTemp!算法 = 2 Then
'            dbl进入统筹 = dbl进入统筹 + IIf(rs明细!金额 < rsTemp!统筹比额, rs明细!金额, rsTemp!统筹比额)
'            dbl全自费 = dbl全自费 + IIf(rs明细!金额 < rsTemp!统筹比额, 0, rs明细!金额 - rsTemp!统筹比额)

            'Beging 20051228 陈东 原冲销记录和负数记录此公式计算有误
            If rs明细!金额 >= 0 Then
                dbl进入统筹 = dbl进入统筹 + IIf(rs明细!金额 < rsTemp!统筹比额, rs明细!金额, rsTemp!统筹比额)
                dbl全自费 = dbl全自费 + IIf(rs明细!金额 < rsTemp!统筹比额, 0, rs明细!金额 - rsTemp!统筹比额)
            Else
                dbl进入统筹 = dbl进入统筹 + IIf(Abs(rs明细!金额) < rsTemp!统筹比额, rs明细!金额, -rsTemp!统筹比额)
                dbl全自费 = dbl全自费 + IIf(Abs(rs明细!金额) < rsTemp!统筹比额, 0, rs明细!金额 + rsTemp!统筹比额)
            End If
            'End    20051228 陈东 原冲销记录和负数记录此公式计算有误
        End If
        cur全额 = cur全额 + Nvl(rs明细!金额, 0)
        rs明细.MoveNext
    Wend
    dblTemp = dbl进入统筹
    
    g结算数据.发生费用金额 = cur全额
    g结算数据.进入统筹金额 = dbl进入统筹
    g结算数据.全自费金额 = dbl全自费
    g结算数据.首先自付金额 = 0
    g结算数据.统筹报销金额 = dbl进入统筹
    
    gstrSQL = "Select * From 病人费用记录 Where 门诊标志=2 And 记录状态<>0 And nvl(附加标志,0)<>9 and nvl(实收金额,0)<>0 and 病人id=" & lng病人ID & " And 主页id=" & g结算数据.主页ID & " order by 主页ID,序号"
    Call OpenRecordset(rsTemp, gstrSysName)
    While Not rsTemp.EOF
        gstrSQL = "select a.名称,b.名称,b.统筹比额,b.算法,c.大类ID from 收费细目 a,保险支付大类 b,保险支付项目 c where " & _
            "b.id=c.大类id and a.id=c.收费细目id and c.险类=" & TYPE_徐州六院 & " and a.id=" & rsTemp!收费细目ID
        Call OpenRecordset(rs算法, gstrSysName)
        If rs算法.EOF Then
            dbl统筹 = 0
        ElseIf rs算法!算法 = 1 Then
            dbl统筹 = rsTemp!实收金额 * rs算法!统筹比额 / 100
        ElseIf rs算法!算法 = 2 Then
'            dbl统筹 = IIf(rsTemp!实收金额 < rs算法!统筹比额, rsTemp!实收金额, rs算法!统筹比额)
            'Beging 20051228 陈东 原冲销记录和负数记录此公式计算有误
            If rsTemp!实收金额 >= 0 Then
                dbl统筹 = IIf(rsTemp!实收金额 < rs算法!统筹比额, rsTemp!实收金额, rs算法!统筹比额)
            Else
                dbl统筹 = IIf(Abs(rsTemp!实收金额) < rs算法!统筹比额, rsTemp!实收金额, -rs算法!统筹比额)
            End If
            'End    20051228 陈东 原冲销记录和负数记录此公式计算有误
        End If
        If Not rs算法.EOF Then
            strTemp = rs算法!大类id
            str类型 = rs算法(1)
        Else
            str类型 = "自费"
            strTemp = "NULL"
        End If
        gcnOracle.Execute "Delete From 离休明细 Where 记录ID=" & rsTemp!ID
        gcnOracle.Execute "insert into 离休明细 values (" & dbl统筹 & "," & strTemp & ",'" & str类型 & "'," & rsTemp!ID & ")"
        rsTemp.MoveNext
    Wend
    
    '循环更新所有项目的费用类型
    Call UpdateClass(g结算数据.病人ID, g结算数据.主页ID)
    
    住院虚拟结算_徐州六院 = "离休医疗;" & dblTemp & ";0"
End Function

Public Function 医保项目_徐州六院(病人ID As Long, 收费细目ID As Long, 金额 As Currency, _
    ByVal bln门诊 As Boolean, Optional ByVal intInsure As Integer) As String
    '提取医保大类做为费用类型返回给主调程序
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select B.名称  " & _
             " From 保险支付项目 A,保险支付大类 B " & _
             " Where A.大类ID=B.Id And A.险类=B.险类 And A.险类=" & intInsure & " And A.收费细目ID=" & 收费细目ID
    Call OpenRecordset(rsTemp, "提取医保大类做为费用类型返回给主调程序")
    
    If rsTemp.RecordCount <> 0 Then
        医保项目_徐州六院 = Nvl(rsTemp!名称)
    End If
End Function

Private Sub UpdateClass(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    Dim str费用类型 As String
    Dim rsTemp As New ADODB.Recordset
    '循环更新所有项目的费用类型
    gstrSQL = "Select ID,病人ID,收费细目ID,费用类型 From 病人费用记录" & _
        " Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID & _
        " And Nvl(是否上传,0)=1 And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0 And Nvl(实收金额,0)<>0 And 费用类型 is null"
    Call OpenRecordset(rsTemp, "循环更新所有项目的费用类型")
    
    With rsTemp
        Do While Not .EOF
            str费用类型 = 医保项目_徐州六院(!病人ID, !收费细目ID, 0, False, TYPE_徐州六院)
            gstrSQL = "ZL_病人记帐记录_上传(" & !ID & ",NULL,NULL,'" & str费用类型 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新费用类型")
            .MoveNext
        Loop
    End With
End Sub
