Attribute VB_Name = "mdlPivaNormal"
Option Explicit


Public Function PIVA_GetAdvice(ByVal bln核查 As Boolean, ByVal lngCenterID As Long, ByVal str病区id As String, ByVal dateExeStart As Date, ByVal dateExeEnd As Date) As ADODB.Recordset
    '取病区已发送但还未发药的医嘱记录，要排除已分解为输液单据的记录
    'lngCenterID：输液配置中心ID
    'dateExeStart、dateExeEnd：医嘱的开始执行时间范围
    '注意：返回的是医嘱和医嘱内容（药品）记录，在取医嘱明细时需要按相关ID排序和合并
    On Error GoTo errHandle
    gstrSQL = "Select /*+ Rule*/ Distinct A.ID As 医嘱id, Nvl(A.相关id, A.ID) As 相关id, M.发送号, E.库房id, D.姓名, D.性别, D.年龄, D.标识号 As 住院号, D.床号, D.病人病区id, " & _
        " D.病人科室id, A.开始执行时间,  M.发送时间, M.发送人, B.名称 As 病人病区, C.名称 As 病人科室, E.ID As 收发id, E.单据, E.NO, F.编码 As 药品编码, F.名称 As 通用名, " & _
        " H.名称 As 商品名, I.名称 As 英文名, F.规格, E.产地, E.批号, E.单量, J.计算单位 As 剂量单位, E.频次, " & _
        " (E.实际数量 / G.住院包装) As 数量, G.住院单位 As 单位 , E.批次, Decode(A.医嘱期效, 0, '长期', '临时') As 医嘱类型, A.执行频次," & _
        " K.名称 As 开嘱科室, A.开嘱医生, A.医生嘱托, A.开嘱时间, A.校对护士, A.校对时间, Nvl(A.审查结果,-1) 审查结果, E.用法, E.药品id " & _
        " From 病人医嘱记录 A, 病人医嘱发送 M, 部门表 B, 部门表 C, 住院费用记录 D, 药品收发记录 E, 收费项目目录 F, 药品规格 G, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 部门表 K "
        
    If str病区id <> "" Then
        gstrSQL = gstrSQL & ",Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) L "
    End If
    
    gstrSQL = gstrSQL & " Where D.病人病区id = B.ID And A.ID = M.医嘱id And M.NO = D.NO And D.病人科室id = C.ID And A.开嘱科室id = K.ID And A.ID = D.医嘱序号 And D.ID = E.费用id And E.药品id = F.ID And F.ID = G.药品id And " & _
        " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID And " & _
        " E.审核日期 Is Null And E.实际数量 > 0 And A.单次用量 > 0 And Not Exists (Select 1 From 输液配药内容 Where 收发id = E.ID And Rownum = 1) " & _
        " And E.库房id = [1] And M.发送时间 Between [3] And [4]  " & _
        " And Exists (Select 1 From 诊疗项目目录 N, 病人医嘱记录 O " & _
        " Where N.类别 = 'E' And N.操作类型 = '2' And N.执行分类 = 1 And O.诊疗项目id = N.ID And Nvl(A.相关id, A.ID) = O.ID) "
        
    If str病区id <> "" Then
        gstrSQL = gstrSQL & " And D.病人病区id + 0 =L.Column_Value "
    End If
    
    If bln核查 = True Then
        gstrSQL = Replace(gstrSQL, "Not Exists", "Exists")
    End If
    
    Set PIVA_GetAdvice = zlDatabase.OpenSQLRecord(gstrSQL, "读取医嘱记录", lngCenterID, str病区id, dateExeStart, dateExeEnd)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PIVA_GetExcStatus(ByVal str配药ids As String, ByVal intStatus As Integer) As ADODB.Recordset
    '检查不符合当前状态的输液单
    'str配药ids：输液单ID串
    'intStatus：当前应该的业务状态
    Dim i As Integer
    Dim arrExecute As Variant
    
    On Error GoTo errHandle
    arrExecute = GetArrayByStr(str配药ids, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = " Select ID, 瓶签号, 操作状态,是否打包 " & _
            " From 输液配药记录 Where (操作状态 <> [2] " & IIf(intStatus = 2, " or 是否打包<>0", "") & ") And ID In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) "
        Set PIVA_GetExcStatus = zlDatabase.OpenSQLRecord(gstrSQL, "PIVA_GetStatus", CStr(arrExecute(i)), intStatus)
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function PIVA_GetAdviceCount(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal dateExeEnd As Date) As ADODB.Recordset
    '取病区已发送但还未发药的医嘱记录数，要排除已分解为输液单据的记录
    'lngCenterID：输液配置中心ID
    'dateExeStart、dateExeEnd：医嘱的执行时间范围（发送时间）
    Dim strTmp As String
    
    On Error GoTo errHandle
    gstrSQL = "Select /*+ rule*/ 病区id, 病区, Count(病区id) As 数量, 0 核查标志 " & _
        " From (Select  Distinct B.病人病区id As 病区id, '[' || D.编码 || ']' || D.名称 As 病区, Nvl(A.相关id, A.ID) As 相关id, E.发送号 " & _
        " From 病人医嘱记录 A, 病人医嘱发送 E, 住院费用记录 B, 药品收发记录 C, 部门表 D " & _
        " Where A.Id = B.医嘱序号 And A.ID = E.医嘱id And E.NO = B.NO And B.Id = C.费用id And B.病人病区id = D.Id And C.审核日期 Is Null And " & _
        " C.实际数量 > 0 And A.单次用量 > 0 And C.库房id = [1] And Not Exists " & _
        " (Select 1 From 输液配药内容 Where 收发id = C.ID And Rownum = 1) And E.发送时间 Between [2] And [3] " & _
        " And Exists (Select 1 From 诊疗项目目录 F, 病人医嘱记录 G " & _
        " Where F.类别 = 'E' And F.操作类型 = '2' And F.执行分类 = 1 And G.诊疗项目id = F.ID And Nvl(A.相关id, A.ID) = G.ID)" & _
        " Order By '[' || D.编码 || ']' || D.名称) " & _
        " Group By 病区id, 病区"
    
    strTmp = Replace(gstrSQL, "0 核查标志", "1 核查标志")
    strTmp = Replace(strTmp, "Not Exists", "Exists")
    
    '合并未核查和已核查的医嘱
    gstrSQL = gstrSQL & " Union All " & strTmp
    
    Set PIVA_GetAdviceCount = zlDatabase.OpenSQLRecord(gstrSQL, "读取病区医嘱记录数", lngCenterID, dateExeStart, dateExeEnd)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Piva_GetMedi(ByVal lngCenterID As Long, ByVal str病区id As String, ByVal intStemp As Integer, ByVal dateExeStart As Date, _
        ByVal dateExeEnd As Date, ByVal int显示所有 As Integer) As Recordset
    
    On Error GoTo errHandle
    If int显示所有 = 0 Then
        gstrSQL = "Select Distinct a.id,a.相关Id,A.病人id,A.主页id,A.开嘱医生,A.审查结果,A.药师审核原因,H.病人病区ID,H.病人科室ID, b.名称 科室名称, f.当前床号 床号,P.配药类型,decode(A.医嘱期效,0,'长期',1,'临时') 医嘱期效,M.名称 给药途径, " & _
            " g.标识号 As 住院号, a.姓名, a.性别, a.年龄,c.id 药品id,  c.名称 药品名称, c.规格, a.单次用量,I.计算单位,I.id 药名id,a.执行频次,nvl(a.药师审核标志,0) 审核标志,a.执行时间方案,A.皮试结果,A.开嘱时间,nvl(T.是否皮试,0) 是否皮试 " & vbNewLine & _
            "From 病人医嘱记录 A, 部门表 B, 收费项目目录 C, 输液配药内容 D, 药品收发记录 E, 病人信息 F, 住院费用记录 G,输液配药记录 H,诊疗项目目录 I,药品规格 J, 药品特性 T,输液药品属性 P,病人医嘱记录 L,诊疗项目目录 M " & vbNewLine & _
             ",Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) K " & vbNewLine & _
            "Where a.病人id = f.病人id  And e.费用id = g.Id And a.病人科室id = b.Id And E.单据=9 and a.Id = g.医嘱序号 And g.收费细目id = c.Id and J.药品id=c.id and J.药品id=P.药品id and J.药名id=I.id And A.相关id=L.id and L.诊疗项目id=M.id And H.id=D.记录id And J.药名id=T.药名id " & vbNewLine & _
            "      And e.Id = d.收发id And (nvl(a.药师审核标志,0)=[1] " & IIf(intStemp = 0, " or nvl(a.药师审核标志,0)=3", "") & ") " & IIf(intStemp = 0, " And H.操作状态 = 1 ", "") & "  And H.病人病区id + 0 =K.Column_Value  and H.执行时间 between [3] and [4] And h.部门id=[5] " & vbNewLine & _
            " order by b.名称,A.病人id,a.相关Id"
            
        Set Piva_GetMedi = zlDatabase.OpenSQLRecord(gstrSQL, "Piva_GetMedi", intStemp, str病区id, dateExeStart, dateExeEnd, lngCenterID)
    Else
        If intStemp = 0 Then
            gstrSQL = "Select Distinct a.id,a.相关Id,A.病人id,A.主页id,A.开嘱医生,A.审查结果,A.药师审核原因,G.病人病区ID,G.病人科室ID, b.名称 科室名称, f.当前床号 床号,P.配药类型,decode(A.医嘱期效,0,'长期',1,'临时') 医嘱期效,M.名称 给药途径, " & _
                " g.标识号 As 住院号, a.姓名, a.性别, a.年龄,c.id 药品id,  c.名称 药品名称, c.规格, a.单次用量,I.计算单位,I.id 药名id,a.执行频次,nvl(a.药师审核标志,0) 审核标志,a.执行时间方案,A.皮试结果,A.开嘱时间,nvl(T.是否皮试,0) 是否皮试 " & vbNewLine & _
                "From 病人医嘱记录 A, 部门表 B, 收费项目目录 C,  药品收发记录 E, 病人信息 F, 住院费用记录 G,诊疗项目目录 I,药品规格 J, 药品特性 T,输液药品属性 P,病人医嘱记录 L,诊疗项目目录 M " & vbNewLine & _
                ",Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) K " & vbNewLine & _
                "Where a.病人id = f.病人id  And e.费用id = g.Id And a.病人科室id = b.Id And E.单据=9 and a.Id = g.医嘱序号 And g.收费细目id = c.Id and J.药品id=c.id and J.药品id=P.药品id and J.药名id=I.id And A.相关id=L.id and L.诊疗项目id=M.id And J.药名id=T.药名id " & vbNewLine & _
                "      And G.病人病区id  =K.Column_Value  and E.填制日期 between [1] and [2] And E.库房id=[3] " & vbNewLine & _
                " order by b.名称,A.病人id,a.相关Id"
            Set Piva_GetMedi = zlDatabase.OpenSQLRecord(gstrSQL, "Piva_GetMedi", dateExeStart, dateExeEnd, lngCenterID, str病区id)
        Else
            gstrSQL = "Select Distinct a.id,a.相关Id,A.病人id,A.主页id,A.开嘱医生,A.审查结果,A.药师审核原因,G.病人病区ID,G.病人科室ID, b.名称 科室名称, f.当前床号 床号,P.配药类型,decode(A.医嘱期效,0,'长期',1,'临时') 医嘱期效,M.名称 给药途径, " & _
                " g.标识号 As 住院号, a.姓名, a.性别, a.年龄,c.id 药品id,  c.名称 药品名称, c.规格, a.单次用量,I.计算单位,I.id 药名id,a.执行频次,nvl(a.药师审核标志,0) 审核标志,a.执行时间方案,A.皮试结果,A.开嘱时间,nvl(T.是否皮试,0) 是否皮试 " & vbNewLine & _
                "From 病人医嘱记录 A, 部门表 B, 收费项目目录 C,  药品收发记录 E, 病人信息 F, 住院费用记录 G,诊疗项目目录 I,药品规格 J, 药品特性 T,输液药品属性 P,病人医嘱记录 L,诊疗项目目录 M " & vbNewLine & _
                ",Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) K " & vbNewLine & _
                "Where a.病人id = f.病人id  And e.费用id = g.Id And a.病人科室id = b.Id And E.单据=9 and a.Id = g.医嘱序号 And g.收费细目id = c.Id and J.药品id=c.id and J.药品id=P.药品id and J.药名id=I.id And A.相关id=L.id and L.诊疗项目id=M.id And J.药名id=T.药名id " & vbNewLine & _
                "      And G.病人病区id  =K.Column_Value  and E.填制日期 between [1] and [2] And E.库房id=[3] And (nvl(a.药师审核标志,0)=[5] " & IIf(intStemp = 0, " or nvl(a.药师审核标志,0)=3", "") & ") " & vbNewLine & _
                " order by b.名称,A.病人id,a.相关Id"
            Set Piva_GetMedi = zlDatabase.OpenSQLRecord(gstrSQL, "Piva_GetMedi", dateExeStart, dateExeEnd, lngCenterID, str病区id, intStemp)
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function Piva_GetTrans(ByVal lngCenterID As Long, ByVal lng病区id As Long, ByVal dateExeStart As Date, _
        ByVal dateExeEnd As Date, ByVal strStep As String, ByVal intPack As Integer, ByVal intSend As Integer, ByVal bln审核 As Boolean, ByVal bln启用审方 As Boolean) As ADODB.Recordset
        
    '取输液配药记录
    'lngCenterID：输液配置中心ID
    'str病区ID：病区ID串
    'dateExeStart、dateExeEnd：输液配药单据的执行时间范围
    'strStep(操作类型)：01-摆药印签(1)，02-配药核查(2)，03-发送核查(4)，04-销帐审核(9)，10-审核已通过医嘱(10)，11-审核未通过医嘱(10)，12-已发送查看(5), 13-已签收查看(6)，14-拒绝签收查(7)，15-已作废查看
    '操作类型：1、填制，2、摆药，3、校对，4、配药，5、发送，6、签收，7、拒绝签收  8，确认拒收，9，销帐申请，10，销帐审核
    'intPack传入：0-所有；1-仅配药；2-仅打包
    '是否打包：0-不打包(配液),1-病区打包,2-配置中心打包
    On Error GoTo errHandle
    
    If strStep = "15" Then
        '已摆药状态
        '1.销帐审核通过
        gstrSQL = "Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,P.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次, A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' NO, F.编码 As 药品编码, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, e.产地, e.批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, e.频次, '销帐审核通过' As 作废类型, " & _
            " 0 As 发药数量, (e.入出系数*e.实际数量 / G.住院包装) As 数量,e.入出系数*e.实际数量 As 实际数量, G.住院单位 As 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, e.用法, e.药品id,0 as 费用序号,0 As 费用id,null As 险类, A.摆药单号,L.发送时间 As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,A.发送号,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1 " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G,输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 住院费用记录 D, 药品收发记录 E, 病人医嘱发送 L ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 P,输液配药内容 Z "

        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID And A.病人科室id = C.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And g.药品id = e.药品id And T.药名id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=P.名称(+) And " & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID " & _
            " And m.Id = d.医嘱序号 And d.Id = e.费用id And a.医嘱id = l.医嘱id(+) And a.发送号 = l.发送号(+) And a.id = z.记录id And z.收发id = e.id " & _
            " And a.操作状态=10 And A.部门id = [1] And A.执行时间 Between [3] And [4] And Exists (Select 1 From 输液配药内容 D, 药品收发记录 E Where d.收发id = e.Id And d.记录id = a.Id)"
            
        If lng病区id <> 0 Then
            gstrSQL = gstrSQL & " And A.病人病区id + 0 =[2] "
        End If

        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
        
        '2.销账审核拒绝
        gstrSQL = gstrSQL & " Union All " & _
            "Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,P.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次, A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' NO, F.编码 As 药品编码, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, e.产地, e.批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, e.频次, '销帐审核拒绝' As 作废类型, " & _
            " 0 As 发药数量, (e.入出系数*e.实际数量 / G.住院包装) As 数量,e.入出系数*e.实际数量 As 实际数量, G.住院单位 As 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, e.用法, e.药品id,0 as 费用序号,0 As 费用id,null As 险类, A.摆药单号,L.发送时间 As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,A.发送号,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1 " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G,输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 住院费用记录 D, 药品收发记录 E, 病人医嘱发送 L ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 P,输液配药内容 Z "

        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID And A.病人科室id = C.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And g.药品id = e.药品id And T.药名id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=P.名称(+) And " & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID " & _
            " And m.Id = d.医嘱序号 And d.Id = e.费用id And a.医嘱id = l.医嘱id(+) And a.发送号 = l.发送号(+)  and e.实际数量>0 And a.id = z.记录id And z.收发id = e.id " & _
            " And a.操作状态=11 And A.部门id = [1] And A.执行时间 Between [3] And [4] And Exists (Select 1 From 输液配药内容 D, 药品收发记录 E Where d.收发id = e.Id And d.记录id = a.Id)"
            
        If lng病区id <> 0 Then
            gstrSQL = gstrSQL & " And A.病人病区id + 0 =[2] "
        End If

        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
        
        '未摆药状态
        '按规格
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,P.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' As NO, F.编码 As 药品编码, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, '' As 产地, '' As 批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, '' As 频次, '未摆药销帐' As 作废类型, " & _
            " 0 As 发药数量, (M.单次用量/ G.剂量系数 / G.住院包装) As 数量,M.单次用量/ G.剂量系数 As 实际数量, G.住院单位 As 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, '' As 用法, M.收费细目id As 药品id,0 as 费用序号,0 As 费用id,null As 险类, " & _
            " A.摆药单号,Null As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,A.发送号,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1 " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G,输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 P "
        
        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID  And A.病人科室id = C.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And M.收费细目id = F.ID And T.药名id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=P.名称(+) And " & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID And a.操作状态=10  " & _
            " And A.部门id = [1] And A.执行时间 Between [3] And [4] And Not Exists (Select 1 From 输液配药内容 D, 药品收发记录 E Where d.收发id = e.Id And d.记录id = a.Id) "
            
        If lng病区id <> 0 Then
            gstrSQL = gstrSQL & " And A.病人病区id + 0 =[2] "
        End If
        
        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
                
        '按品种
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,P.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' As NO, J.编码 As 药品编码, " & _
            " J.名称 As 通用名, '' As 商品名, I.名称 As 英文名, '' as 规格, '' As 产地, '' As 批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, '' As 频次, '未摆药销帐' As 作废类型, " & _
            " 0 As 发药数量, 0 As 数量,0 As 实际数量, '' 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, '' As 用法, Decode(Nvl(m.收费细目id, 0), 0, j.Id, m.收费细目id) As 药品id,0 as 费用序号,0 As 费用id,null As 险类, " & _
            " A.摆药单号,Null As 医嘱发送时间,0 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,A.发送号,nvl(T.是否皮试,0) 是否皮试,'' As 配药类型1 " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C,诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 P "

        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID  And A.病人科室id = C.ID and M.收费细目id is null and M.诊疗项目id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=P.名称(+) And a.操作状态=10  " & _
            " And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And J.id = I.诊疗项目id(+) And I.性质(+) = 2 And j.Id = t.药名id " & _
            " And A.部门id = [1] And A.执行时间 Between [3] And [4] And Not Exists (Select 1 From 输液配药内容 D, 药品收发记录 E Where d.收发id = e.Id And d.记录id = a.Id) "

        If lng病区id <> 0 Then
            gstrSQL = gstrSQL & " And A.病人病区id + 0 =[2] "
        End If

        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
    ElseIf strStep = "16" Then
        '医嘱回退(按规格)
        gstrSQL = " Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,P.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' As NO, F.编码 As 药品编码, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, '' As 产地, '' As 批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, '' As 频次, '医嘱回退' As 作废类型, " & _
            " 0 As 发药数量, (M.单次用量/ G.剂量系数 / G.住院包装) As 数量,M.单次用量/ G.剂量系数 As 实际数量, G.住院单位 As 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, '' As 用法, M.收费细目id As 药品id,0 as 费用序号,0 As 费用id,null As 险类, " & _
            " A.摆药单号,Null As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,A.发送号,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1 " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G,输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 P "
        
        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID  And A.病人科室id = C.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And M.收费细目id = F.ID And T.药名id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=P.名称(+) And " & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID And a.操作状态=12  " & _
            " And A.部门id = [1] And A.执行时间 Between [3] And [4] "
            
        If lng病区id <> 0 Then
            gstrSQL = gstrSQL & " And A.病人病区id + 0 =[2] "
        End If
        
        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
                
        '合并医嘱回退(按品种发送)
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,P.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' As NO, J.编码 As 药品编码, " & _
            " J.名称 As 通用名, '' As 商品名, I.名称 As 英文名, '' as 规格, '' As 产地, '' As 批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, '' As 频次, '医嘱回退' As 作废类型, " & _
            " 0 As 发药数量, 0 As 数量,0 As 实际数量, '' 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, '' As 用法, Decode(Nvl(m.收费细目id, 0), 0, j.Id, m.收费细目id) As 药品id,0 as 费用序号,0 As 费用id,null As 险类, " & _
            " A.摆药单号,Null As 医嘱发送时间,0 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,A.发送号,nvl(T.是否皮试,0) 是否皮试,'' As 配药类型1 " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C,诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 P "

        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID  And A.病人科室id = C.ID and M.收费细目id is null and M.诊疗项目id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=P.名称(+) And a.操作状态=12  " & _
            " And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And J.id = I.诊疗项目id(+) And I.性质(+) = 2 And j.Id = t.药名id " & _
            " And A.部门id = [1] And A.执行时间 Between [3] And [4] "

        If lng病区id <> 0 Then
            gstrSQL = gstrSQL & " And A.病人病区id + 0 =[2] "
        End If

        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
    Else
        '其他
        gstrSQL = "Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次, S.颜色,A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,P.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,A.是否调整批次,A.是否锁定,A.手工调整批次," & IIf(strStep = "13", "W.操作说明 拒收原因,", "'' 拒收原因,") & _
            "  A.操作人员,A.操作时间,Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, D.收发id, E.单据, E.NO, F.编码 As 药品编码, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, E.产地, E.批号, E.单量, J.计算单位 As 剂量单位,J.id 药名id, E.频次, '' As 作废类型, " & _
            " Case Nvl(E.审核人, '未审核') When '未审核' Then E.实际数量 * Nvl(E.付数, 1) / G.住院包装 Else 0 End As 发药数量,M.病人id,M.主页id,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,A.发送号, " & _
            " (D.数量 / G.住院包装)  As 数量,D.数量 As 实际数量, G.住院单位 As 单位,Nvl(E.批次,0) As 批次, Nvl(L.实际数量, 0)/ G.住院包装 As 库存数量, Nvl(M.审查结果,-1) 审查结果, E.用法, E.药品id, n.序号 As 费用序号,E.费用id, o.险类, A.摆药单号,r.发送时间 As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1 " & _
            " From  输液配药记录 A, 部门表 B, 部门表 C, 输液配药内容 D, 药品收发记录 E, 收费项目目录 F, 药品规格 G,输液药品属性 X,  收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 住院费用记录 N, 病案主页 O ,配药工作批次 S,药品特性 T,床位状况记录 Q,床位编制分类 P "
        
        If strStep = "13" Then gstrSQL = gstrSQL & ",输液配药状态 W "
        
        If strStep = "01" And bln启用审方 Then
            gstrSQL = gstrSQL & ",处方审查记录 Q,处方审查明细 K "
        End If
        
        gstrSQL = gstrSQL & ",(Select 库房id, 药品id, Nvl(批次, 0) As 批次, Nvl(实际数量, 0) As 实际数量 " & _
            " From 药品库存 Where 性质 = 1 And 库房id = [1]) L, 药品收发记录 P, 病人医嘱发送 R "
        
        gstrSQL = gstrSQL & " Where A.病人病区id = B.ID And A.病人科室id = C.ID And A.ID = D.记录id And D.收发id = E.ID And E.药品id = F.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And E.费用id = N.ID And N.医嘱序号 = M.ID And " & IIf(strStep = "13", "W.配药id=A.id And A.操作状态=W.操作类型 And A.操作时间=W.操作时间 And ", "") & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID And T.药名id=J.ID And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And E.库房id = L.库房id(+) And E.药品id = L.药品id(+) And A.床号=Q.床号(+) And  A.病人病区id=Q.病区id(+) And A.病人科室id=Q.科室id(+) and Q.床位编制=P.名称(+) And Nvl(E.批次, 0) = L.批次(+) " & _
            " And n.病人id = o.病人id(+) And n.主页id = o.主页id(+) And A.部门id = [1] And a.医嘱id = r.医嘱id And a.发送号 = r.发送号 And " & IIf(strStep = "04", "A.执行时间", "A.执行时间") & " Between [3] And [4] " & _
            " And e.单据 = p.单据 And e.No = p.No And e.库房id + 0 = p.库房id And e.药品id + 0 = p.药品id And e.序号 = p.序号 And (p.记录状态 = 1 Or Mod(p.记录状态, 3) = 0) "
            
        If lng病区id <> 0 Then
            gstrSQL = gstrSQL & " And A.病人病区id + 0 =[2] "
        End If
        
        If strStep = "01" Then
            '待摆药
            If bln启用审方 Then gstrSQL = gstrSQL & " And Q.id=K.审方ID and K.医嘱id=M.id and Q.审查结果=1 and K.最后提交=1 "
            
            gstrSQL = gstrSQL & " And (" & IIf(bln审核, "M.药师审核标志=1 And", "") & " A.操作状态=1) "
        ElseIf strStep = "02" Then
            '待配药
            gstrSQL = gstrSQL & " And A.操作状态=2 "
        ElseIf strStep = "03" Then
            '待发送
            gstrSQL = gstrSQL & " And A.操作状态=4 "
        ElseIf strStep = "11" Then
            '已销账审核
            gstrSQL = gstrSQL & " And A.操作状态=10 "
        ElseIf strStep = "12" Then
            '已发送
            gstrSQL = gstrSQL & " And A.操作状态=5 "
        ElseIf strStep = "13" Then
            '已签收
            gstrSQL = gstrSQL & " And A.操作状态=6 "
        ElseIf strStep = "14" Then
            '已拒绝签收
            gstrSQL = gstrSQL & " And A.操作状态=7 "
        ElseIf strStep = "04" Then
            gstrSQL = gstrSQL & " And A.操作状态=9 "
        End If
        
        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
    End If
    
    Set Piva_GetTrans = zlDatabase.OpenSQLRecord(gstrSQL, "读取输液配药记录", lngCenterID, lng病区id, dateExeStart, dateExeEnd)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function PIVA_GetTransCount(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal dateExeEnd As Date, ByVal bln审核 As Boolean, ByVal bln启用审方 As Boolean, Optional intCheck As Integer) As ADODB.Recordset
    '取病区输液单数目
    'lngCenterID：输液配置中心ID
    'dateExeStart、dateExeEnd：输液配药单据的执行时间范围
    On Error GoTo errHandle
    
    gstrSQL = "select 类型, 病区id, 病区,  数量,药师审核标志,名称,编码 from (with W as (Select Distinct a.操作状态,c.药师审核标志, a.病人病区id As 病区id, '[' || b.编码 || ']' || b.名称 As 病区,b.名称,b.编码, c.相关id As 医嘱id,A.id " & vbNewLine & _
        "       From 输液配药记录 A, 部门表 B, 病人医嘱记录 C" & IIf(bln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & vbNewLine & _
        "       Where a.病人病区id = b.Id And a.医嘱id = c.相关id And c.执行性质 <> 5 And a.部门id = [1] And" & IIf(bln启用审方, " c.id=k.医嘱id and Q.id=K.审方id and K.最后提交=1 and Q.审查结果=1 and", "") & vbNewLine & _
        "             a.执行时间 Between [2] And [3]" & vbNewLine & _
        "             And Exists" & vbNewLine & _
        "        (Select 1 From 输液配药内容 D Where d.记录id = a.Id))," & vbNewLine & _
        "       R as (Select Distinct a.操作状态, a.病人病区id As 病区id, '[' || b.编码 || ']' || b.名称 As 病区,b.名称,b.编码," & vbNewLine & _
        "                                 a.Id" & vbNewLine & _
        "                 From 输液配药记录 A, 部门表 B" & vbNewLine & _
        "                 Where a.病人病区id = b.Id  And a.部门id = [1] And" & vbNewLine & _
        "                       a.执行时间 Between [2]  and" & vbNewLine & _
        "                       [3] And Exists" & vbNewLine & _
        "                  (Select 1 From 输液配药内容 D Where d.记录id = a.Id))"


    If bln审核 = True Then
        '审核医嘱
        If intCheck = 0 Then
            gstrSQL = gstrSQL & " Select 类型, 病区id, 病区, Count(医嘱id) As 数量,0 药师审核标志,名称,编码 " & vbNewLine & _
            "From ( select Distinct '00' 类型,病区id,病区,医嘱id,名称,编码 from  W where (Nvl(药师审核标志, 0) = 0 or Nvl(药师审核标志, 0)=3)  and 操作状态=1)" & vbNewLine & _
            "Group By 类型, 病区id, 病区,名称,编码" & vbNewLine & _
            "union all"
        Else
            gstrSQL = gstrSQL & "select 类型,病区id,病区,count(医嘱id) as 数量,Nvl(药师审核标志,0) 药师审核标志,名称,编码 from (" & _
                " Select distinct '00' As 类型, D.病人病区id As 病区id, '[' || B.编码 || ']' || B.名称 As 病区, c.相关id As 医嘱id,c.药师审核标志,B.名称,b.编码 " & _
                " From 药品收发记录 A, 部门表 B,病人医嘱记录 C,住院费用记录 D " & IIf(bln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & vbNewLine & _
                " Where D.病人病区id = B.ID And D.医嘱序号=C.id And A.费用id=D.id And C.执行性质<>5  And A.库房id = [1] and A.单据=9 And A.填制日期 Between [2] And [3]) " & IIf(bln启用审方, " c.id=k.医嘱id and Q.id=K.审方id and K.最后提交=1 and Q.审查结果=1 and", "") & vbNewLine & _
                " Group By 类型,病区id,病区,Nvl(药师审核标志,0),名称,编码 " & vbNewLine & _
                "union all"
        End If
        
'        gstrSQL = gstrSQL & " Select 类型, 病区id, 病区, Count(医嘱id) As 数量,1 药师审核标志" & vbNewLine & _
'            "From ( select Distinct '00' 类型,病区id,病区,医嘱id from  W where Nvl(药师审核标志, 0) = 0 and 操作状态=1)" & vbNewLine & _
'            "Group By 类型, 病区id, 病区" & vbNewLine & _
'            "union all"
            
         '摆药
        gstrSQL = gstrSQL & " Select 类型, 病区id, 病区, Count(id) As 数量,1 药师审核标志,名称,编码 " & vbNewLine & _
            "From ( select '01' 类型,病区id,病区,医嘱id,id,名称,编码 from  W where Nvl(药师审核标志, 0) =1 and 操作状态=1)" & vbNewLine & _
            "Group By 类型, 病区id, 病区,名称,编码"
    Else
        '摆药
        gstrSQL = gstrSQL & " Select 类型, 病区id, 病区, Count(id) As 数量,1 药师审核标志,名称,编码 " & vbNewLine & _
            "From ( select '01' 类型,病区id,病区,医嘱id,id,名称,编码 from  W where 操作状态=1)" & vbNewLine & _
            "Group By 类型, 病区id, 病区,名称,编码"
    End If
    '配药
    gstrSQL = gstrSQL & " Union All " & _
        "Select 类型, 病区id, 病区, Count(id) As 数量,1 药师审核标志,名称,编码 " & vbNewLine & _
        "From ( select '02' 类型,病区id,病区,id,名称,编码 from  R where 操作状态=2)" & vbNewLine & _
        "Group By 类型, 病区id, 病区,名称,编码"

    '发送
    gstrSQL = gstrSQL & " Union All " & _
        "Select 类型, 病区id, 病区, Count(id) As 数量,1 药师审核标志,名称,编码 " & vbNewLine & _
        "From ( select '03' 类型,病区id,病区 ,id,名称,编码 from  R where 操作状态=4)" & vbNewLine & _
        "Group By 类型, 病区id, 病区,名称,编码"

    '销账审核
    gstrSQL = gstrSQL & " Union All " & _
        "Select 类型, 病区id, 病区, Count(id) As 数量,1 药师审核标志 ,名称,编码" & vbNewLine & _
        "From ( select '04' 类型,病区id,病区,id,名称,编码 from  R where 操作状态=9)" & vbNewLine & _
        "Group By 类型, 病区id, 病区,名称,编码"

        
    If bln审核 = True Then
        If intCheck = 0 Then
            '已审核通过医嘱查看
            gstrSQL = gstrSQL & " Union All " & _
                "Select 类型, 病区id, 病区, Count(医嘱id) As 数量,1 药师审核标志,名称,编码 " & vbNewLine & _
                "From ( select Distinct  '10' 类型,病区id,病区,医嘱id,名称,编码 from  W where  Nvl(药师审核标志, 0) =1)" & vbNewLine & _
                "Group By 类型, 病区id, 病区,名称,编码"
    
            '未审核通过医嘱查看
            gstrSQL = gstrSQL & " Union All " & _
                "Select 类型, 病区id, 病区, Count(医嘱id) As 数量,2 药师审核标志,名称,编码 " & vbNewLine & _
                "From ( select Distinct  '11' 类型,病区id,病区,医嘱id,名称,编码 from  W where  Nvl(药师审核标志, 0) =2)" & vbNewLine & _
                "Group By 类型, 病区id, 病区,名称,编码"
        Else
            gstrSQL = gstrSQL & " Union All " & _
                "select 类型,病区id,病区,count(医嘱id) as 数量,药师审核标志,名称,编码 from (" & _
                " Select distinct '10' As 类型, D.病人病区id As 病区id, '[' || B.编码 || ']' || B.名称 As 病区, c.相关id As 医嘱id,c.药师审核标志,B.名称,B.编码 " & _
                " From 药品收发记录 A, 部门表 B,病人医嘱记录 C,住院费用记录 D " & _
                " Where D.病人病区id = B.ID And D.医嘱序号=C.id And A.费用id=D.id and A.单据=9  And C.执行性质<>5 and c.药师审核标志=1 And A.库房id = [1] And A.填制日期 Between [2] And [3]) " & _
                " Group By 类型,病区id,病区,药师审核标志,名称,编码 "
                
            gstrSQL = gstrSQL & " Union All " & _
                "select 类型,病区id,病区,count(医嘱id) as 数量,药师审核标志,名称,编码 from (" & _
                " Select distinct '11' As 类型, D.病人病区id As 病区id, '[' || B.编码 || ']' || B.名称 As 病区, c.相关id As 医嘱id,c.药师审核标志,B.名称,B.编码 " & _
                " From 药品收发记录 A, 部门表 B,病人医嘱记录 C,住院费用记录 D " & _
                " Where D.病人病区id = B.ID And D.医嘱序号=C.id And A.费用id=D.id and A.单据=9 And C.执行性质<>5 and c.药师审核标志=2 And A.库房id = [1] And A.填制日期 Between [2] And [3]) " & _
                " Group By 类型,病区id,病区,药师审核标志,名称,编码 "
        End If

    End If
    '已发送查看
    gstrSQL = gstrSQL & " Union All " & _
        "Select 类型, 病区id, 病区, Count(id) As 数量,1 药师审核标志,名称,编码 " & vbNewLine & _
        "From ( select '12' 类型,病区id,病区,id,名称,编码 from  R where 操作状态=5)" & vbNewLine & _
        "Group By 类型, 病区id, 病区,名称,编码"

    '已签收查看
    gstrSQL = gstrSQL & " Union All " & _
        "Select 类型, 病区id, 病区, Count(id) As 数量,1 药师审核标志,名称,编码 " & vbNewLine & _
        "From ( select '13' 类型,病区id,病区,id,名称,编码 from  R where 操作状态=6)" & vbNewLine & _
        "Group By 类型, 病区id, 病区,名称,编码"

    '拒绝签收查看
    gstrSQL = gstrSQL & " Union All " & _
        "Select 类型, 病区id, 病区, Count(id) As 数量,1 药师审核标志,名称,编码 " & vbNewLine & _
        "From ( select '14' 类型,病区id,病区,id,名称,编码 from  R where 操作状态=7)" & vbNewLine & _
        "Group By 类型, 病区id, 病区,名称,编码"

    '已作废审核查看
    gstrSQL = gstrSQL & " Union All " & _
        "Select '15' As 类型, 病区id, 病区, Sum(数量) As 数量,1 药师审核标志,名称,编码 " & vbNewLine & _
        "From (Select a.病人病区id As 病区id, '[' || b.编码 || ']' || b.名称 As 病区,名称,b.编码, Count(a.Id) As 数量" & vbNewLine & _
        "       From (Select ID, 病人病区id" & vbNewLine & _
        "              From 输液配药记录 A" & vbNewLine & _
        "              Where a.部门id = [1] And a.执行时间 Between [2] And [3] And Nvl(a.操作状态, 0) In (10,11)) A, 部门表 B" & vbNewLine & _
        "       Where a.病人病区id = b.Id" & vbNewLine & _
        "       Group By a.病人病区id, '[' || b.编码 || ']' || b.名称,b.名称,编码)" & vbNewLine & _
        "   Group By 病区id, 病区,名称,编码"
    '医嘱回退查看
    gstrSQL = gstrSQL & " Union All " & _
        " Select '16' As 类型, A.病人病区id As 病区id, '[' || B.编码 || ']' || B.名称 As 病区, Count(A.ID) As 数量,1 药师审核标志,b.名称,b.编码 " & _
        " From 输液配药记录 A, 部门表 B " & _
        " Where A.病人病区id = B.ID And A.操作状态=12 And A.部门id = [1] And A.执行时间 Between [2] And [3] " & _
        " Group By A.病人病区id, '[' || B.编码 || ']' || B.名称,名称,编码 "
        
    gstrSQL = gstrSQL & " Order By 类型, 名称 )"
        
    Set PIVA_GetTransCount = zlDatabase.OpenSQLRecord(gstrSQL, "取病区输液单数目", lngCenterID, dateExeStart, dateExeEnd)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Sub PIVA_AnalysisTrans(ByVal lngCenterID As Long, ByVal dateStart As String, ByVal dateEnd As String)
    'PIVA后台工作：分解发药单，产生输液单
    'lngCenterID：输液配置中心ID
    'dateStart、dateEnd：发药单据的填制时间范围
    On Error GoTo ErrHand
    gstrSQL = "Zl_输液配药记录_Insert("
    '配置中心ID
    gstrSQL = gstrSQL & lngCenterID
    '开始时间
    gstrSQL = gstrSQL & "," & dateStart
    '结束时间
    gstrSQL = gstrSQL & "," & dateEnd
    gstrSQL = gstrSQL & ")"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "产生输液配药记录")
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Function DeptSendWork_Get科室名称() As Recordset
'获取病人科室名称，取工作性质为临床或护理的部门
    On Error GoTo ErrHand
    
    gstrSQL = "Select distinct a.Id, a.编码, a.名称,zlSpellCode(a.名称) 简码,zlWBCode(a.名称) 五笔简码, a.撤档时间" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B" & vbNewLine & _
            "Where a.Id = b.部门id And (b.工作性质 = '临床' Or b.工作性质 = '护理') And" & vbNewLine & _
            "      (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000/1/1', 'yyyy/mm/dd'))"
    
    
    Set DeptSendWork_Get科室名称 = zlDatabase.OpenSQLRecord(gstrSQL, "获取科室信息")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DeptSendWork_Get配药类型() As Recordset
'获取药品的配药类型
    On Error GoTo ErrHand
    gstrSQL = "select 编码,名称 from 输液配药类型"
    
    Set DeptSendWork_Get配药类型 = zlDatabase.OpenSQLRecord(gstrSQL, "获取配药类型")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get频次() As Recordset
'获取药品的配药类型
    On Error GoTo ErrHand
    gstrSQL = "select 编码,名称,英文名称 from 诊疗频率项目 where 编码 not like '-%'"
    
    Set DeptSendWork_Get频次 = zlDatabase.OpenSQLRecord(gstrSQL, "获取频次")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get收费项目() As Recordset
'获取收费项目
    On Error GoTo ErrHand
    gstrSQL = "select id,编码,名称,计算单位,说明 from 收费项目目录 where 类别='Z' and nvl(是否变价,0)=0"
    
    Set DeptSendWork_Get收费项目 = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费项目")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_给药途径() As Recordset
'获取给药途径,目前只针对“静脉营养”类
    On Error GoTo ErrHand
    gstrSQL = "select ID, 名称 from 诊疗项目目录 where 类别 = 'E' and 操作类型 = '2' and 执行分类 = '1' and 执行标记 = 2"
    
    Set DeptSendWork_给药途径 = zlDatabase.OpenSQLRecord(gstrSQL, "获取给药途径")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PIVA_已摆药输液单(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal lng病人ID As Long) As Recordset
'获取该病人当天的已经摆药还未配药的输液单
    On Error GoTo errHandle

    gstrSQL = "Select Distinct A.ID As 配药ID, A.部门id, A.序号, A.配药批次, S.颜色,A.姓名, A.性别, A.年龄, A.住院号, A.床号,M.药师审核时间, A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.执行频次,A.是否调整批次,A.是否锁定,A.手工调整批次," & _
            " A.操作人员, A.操作时间,Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, D.收发id, E.单据, E.NO, F.编码 As 药品编码, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, E.产地, E.批号, E.单量, J.计算单位 As 剂量单位,J.id 药名id, E.频次," & _
            " Case Nvl(E.审核人, '未审核') When '未审核' Then E.实际数量 * Nvl(E.付数, 1) / G.住院包装 Else 0 End As 发药数量,M.病人id,M.主页id, " & _
            " (D.数量 / G.住院包装)  As 数量,D.数量 As 实际数量, G.住院单位 As 单位,Nvl(E.批次,0) As 批次, Nvl(L.实际数量, 0)/ G.住院包装 As 库存数量, Nvl(M.审查结果,-1) 审查结果, E.用法, E.药品id, n.序号 As 费用序号,E.费用id, o.险类, A.摆药单号,r.发送时间 As 医嘱发送时间,nvl(X.配药类型,'') 配药类型 " & _
            " From  输液配药记录 A, 部门表 B, 部门表 C, 输液配药内容 D, 药品收发记录 E, 收费项目目录 F, 药品规格 G,输液药品属性 X,  收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 住院费用记录 N, 病案主页 O ,配药工作批次 S "


        gstrSQL = gstrSQL & ",(Select 库房id, 药品id, Nvl(批次, 0) As 批次, Nvl(实际数量, 0) As 实际数量 " & _
            " From 药品库存 Where 性质 = 1 And 库房id = [1]) L, 药品收发记录 P, 病人医嘱发送 R "

        gstrSQL = gstrSQL & " Where A.病人病区id = B.ID And A.病人科室id = C.ID And A.ID = D.记录id And D.收发id = E.ID And E.药品id = F.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And E.费用id = N.ID And N.医嘱序号 = M.ID And " & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID And A.配药批次=S.批次(+) And E.库房id = L.库房id(+) And E.药品id = L.药品id(+) And Nvl(E.批次, 0) = L.批次(+) " & _
            " And n.病人id = o.病人id(+) And n.主页id = o.主页id(+) And A.部门id = [1] And a.医嘱id = r.医嘱id And a.发送号 = r.发送号 And A.执行时间 between [2] and [3] " & _
            " And e.单据 = p.单据 And e.No = p.No And e.库房id + 0 = p.库房id And e.药品id + 0 = p.药品id And e.序号 = p.序号 And (p.记录状态 = 1 Or Mod(p.记录状态, 3) = 0) "



        gstrSQL = gstrSQL & " And A.操作状态=2 and M.病人id=[4] "
        

        Set PIVA_已摆药输液单 = zlDatabase.OpenSQLRecord(gstrSQL, "读取输液配药记录", lngCenterID, CDate(Format(dateExeStart, "yyyy-mm-dd 00:00:00")), CDate(Format(dateExeStart, "yyyy-mm-dd 23:59:59")), lng病人ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function






