Attribute VB_Name = "mdlAdvice"
Option Explicit

'刷新数据时传入的病人状态
Public Enum TYPE_PATI_State
    ps在院 = 0
    ps预出 = 1
    ps出院 = 2
    ps待诊 = 3          '医生站:待会诊病人(在院)
    ps已诊 = 4          '医生站:已会诊病人
    ps最近转出 = 5      '医护站:最近转科或转病区的病人(在院)
    ps待转入 = 6        '医护站:入科待入住或转病区待入往病人
End Enum

Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Public Function StringMask(ByVal strText As String, ByVal strMask As String) As Boolean
'功能：检查字符串是否只包含指定的字符
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strMask, Mid(strText, i, 1)) = 0 Then Exit Function
    Next
    StringMask = True
End Function

Public Function BillExpend(ByVal strNO As String) As Boolean
'功能：判断挂号单是否已经超过有效挂号天数。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    '如果启用了允许处理超过挂号有效天数的病人参数，则表示不做检查
    If Val(zlDatabase.GetPara(210, glngSys)) = 1 Then Exit Function
    '按时点算
    strSQL = "Select  Sysdate-发生时间 as 间隔,急诊 From 病人挂号记录 Where NO=[1] And 记录性质=1 And 记录状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)
    If Not rsTmp.EOF Then
        BillExpend = Val(rsTmp!间隔) > IIf(Val("" & rsTmp!急诊) = 1, IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数), IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数))
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckOutAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：检查病人是否已下达了出院医嘱
    Dim strSQL As String, rsTmp As Recordset
    
    strSQL = "Select 1 from 病人医嘱记录 A,诊疗项目目录 B Where a.诊疗项目ID=b.ID And a.病人ID=[1] And a.主页ID=[2] And b.类别='Z' And b.操作类型='5' And a.医嘱状态<>4  and nvl(A.婴儿,0)=0"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng病人ID, lng主页ID)
    CheckOutAdvice = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExeTimeValid(ByVal strTime As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, Optional ByVal bln首日 As Boolean) As Boolean
'功能：检查指定的执行时间是否合法
    Dim arrTime() As String, strTmp As String, i As Integer
    Dim strPreTime As String, intPreDay As Long, intCurDay As Long
    
    If strTime = "" Then
        If str间隔单位 = "分钟" Then ExeTimeValid = True
        Exit Function
    End If
    
    If str间隔单位 = "周" Then
        '1/8:00-3/15:00-5/9:00；1/8:00-3/15-5/9:00
        If Not StringMask(strTime, "0123456789:-/") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If bln首日 Then
            If Not Between(UBound(arrTime) + 1, 1, int频率次数) Then Exit Function
        Else
            If UBound(arrTime) + 1 <> int频率次数 Then Exit Function
        End If
        
        For i = 0 To UBound(arrTime)
            If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
            '星期部份
            strTmp = Split(arrTime(i), "/")(0)
            If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
            intCurDay = Val(strTmp)
            If intCurDay < 1 Or intCurDay > 7 Then Exit Function
            If intPreDay <> 0 Then
                If intCurDay < intPreDay Then Exit Function
            End If
            
            '绝对时间部分
            strTmp = Split(arrTime(i), "/")(1)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
            If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            
            strPreTime = Format(strTmp, "HH:mm")
            intPreDay = intCurDay
        Next
    ElseIf str间隔单位 = "天" Then
        If int频率间隔 = 1 Then
            '8:00-12:00-14:00；8:00-12-14:00
            If Not StringMask(strTime, "0123456789:-") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If bln首日 Then
                If Not Between(UBound(arrTime) + 1, 1, int频率次数) Then Exit Function
            Else
                If UBound(arrTime) + 1 <> int频率次数 Then Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                strTmp = arrTime(i)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
                If strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                strPreTime = Format(strTmp, "HH:mm")
            Next
        Else
            '1/8:00-1/15:00-2/9:00；1/8:00-1/15-2/9:00
            If Not StringMask(strTime, "0123456789:-/") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If bln首日 Then
                If Not Between(UBound(arrTime) + 1, 1, int频率次数) Then Exit Function
            Else
                If UBound(arrTime) + 1 <> int频率次数 Then Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
                '相对天数部份
                strTmp = Split(arrTime(i), "/")(0)
                If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
                intCurDay = Val(strTmp)
                If intCurDay < 1 Or intCurDay > int频率间隔 Then Exit Function
                If intPreDay <> 0 Then
                    If intCurDay < intPreDay Then Exit Function
                End If
                
                '绝对时间部分
                strTmp = Split(arrTime(i), "/")(1)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
                If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                
                strPreTime = Format(strTmp, "HH:mm")
                intPreDay = intCurDay
            Next
        End If
    ElseIf str间隔单位 = "小时" Then
        '1:30-2-3:30
        If Not StringMask(strTime, "0123456789:-") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> int频率次数 Then Exit Function
        
        For i = 0 To UBound(arrTime)
            strTmp = arrTime(i)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) < 1 Or Val(Split(strTmp, ":")(0)) > int频率间隔 Or Split(strTmp, ":")(0) = "" Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
            If strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            strPreTime = Format(strTmp, "HH:mm")
        Next
    End If
    
    ExeTimeValid = True
End Function

Public Function GetWeekBase(ByVal datDate As Date) As Date
'功能：获取指定时间所在星期的星期一的时间
    'Oracle:Select Sysdate-(To_Number(To_Char(Sysdate,'D'))-1)+1 From Dual
    GetWeekBase = Format(datDate - (Weekday(datDate, vbMonday) - 1), "yyyy-MM-dd 00:00:00")
End Function

Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个时间是否在暂停的时间段中
'参数：strPause="暂停时间,开始时间;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '可能尚未启用或暂停的时候被停止
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function


Public Function GetMaxBedLen(Optional lng部门ID As Long, Optional bln科室 As Boolean) As Integer
'功能：获取指定部门的床位号的最大长度
'参数：lng部门ID=病区ID或科室ID,为0表示所有病区或科室
'      bln占用=是否只管被占用的床
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not bln科室 Or lng部门ID = 0 Then
        strSQL = "Select Max(LengthB(床号)) as 长度 From 床位状况记录 Where 状态='占用' And 病区ID" & IIf(lng部门ID = 0, " is Not NULL", "= [1] ")
    Else
        strSQL = "Select Max(LengthB(床号)) as 长度 From 床位状况记录 Where 状态='占用' And 科室ID" & IIf(lng部门ID = 0, " is Not NULL", "= [1] ")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng部门ID)
    
    If Not rsTmp.EOF Then GetMaxBedLen = IIf(IsNull(rsTmp!长度), 0, rsTmp!长度)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DateIsPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个日期是否在暂停的时间段中
'参数：strPause="暂停时间,开始时间;...."
'说明：不按时点判断,对暂停日期按算始不算止规则判断
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Format(Split(arrPause(i), ",")(0), "yyyy-MM-dd")
        strEnd = Format(Split(arrPause(i), ",")(1), "yyyy-MM-dd")
        If strEnd = "" Then strEnd = "3000-01-01" '可能尚未启用或暂停的时候被停止
        If strEnd > strBegin Then
            If Between(Format(vDate, "yyyy-MM-dd"), strBegin, _
                Format(DateAdd("d", -1, CDate(strEnd)), "yyyy-MM-dd")) Then
                DateIsPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function TimeisLastPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个时间是否在最后一次暂停的时间内,且最后一次暂停没有启用
'说明：因为这种情况下,如果长嘱没有终止时间,某些计算会死循环
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    
    For i = UBound(arrPause) To 0 Step -1
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then
            strEnd = "3000-01-01 00:00:00"
            If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
                TimeisLastPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function Calc次数分解时间(lng次数 As Long, ByVal dat开始时间 As Date, dat终止时间 As Date, strPause As String, _
    ByVal str执行时间 As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, _
    Optional ByVal dat首日日期 As Date) As String
'功能：按次数计算各次的分解执行时间,要求<=终止时间及不在暂停时间段内
'参数：dat开始时间=医嘱的开始执行时间
'      dat终止时间=医嘱的执行终止时间,没有时传入"3000-01-01"
'      strPause=医嘱的暂停时间段
'      dat首日日期=用于首日时间计算参照
'返回：1."时间1,时间2,...."(yyyy-MM-dd HH:mm:ss)
'      2.lng次数=实际能够分解的次数
'说明：1.因为终止时间的限制,因此分解出来的时间个数可能小于要分解的次数
'      2.本函数是假定在执行时间及频率性质完全正确的情况下计算。
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrFirst As Variant, arrNormal As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    If InStr(str执行时间, ",") > 0 Then
        arrNormal = Split(Split(str执行时间, ",")(1), "-")
        arrFirst = Split(Split(str执行时间, ",")(0), "-")
    Else
        arrNormal = Split(str执行时间, "-")
        arrFirst = Array()
    End If
    
    vCurTime = dat开始时间
    
    If str间隔单位 = "周" Then
        vCurTime = GetWeekBase(dat开始时间)
        
        Do While lng次数 > 0
            blnFirst = (GetWeekBase(vCurTime) = GetWeekBase(dat首日日期)) And dat首日日期 <> Empty And UBound(arrFirst) <> -1
            arrTime = IIf(blnFirst, arrFirst, arrNormal)

            '1/8:00-3/15:00-5/9:00
            For i = 1 To int频率次数
                If i - 1 <= UBound(arrTime) Then '首周可能次数不足
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime > dat终止时间 Then
                        Exit Do
                    ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                        Exit Do
                    ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        lng次数 = lng次数 - 1
                        If lng次数 = 0 Then Exit Do
                    End If
                End If
            Next
            vCurTime = vCurTime + 7
        Loop
    ElseIf str间隔单位 = "天" Then
        Do While lng次数 > 0
            blnFirst = (Int(vCurTime) = Int(dat首日日期)) And dat首日日期 <> Empty And UBound(arrFirst) <> -1
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
        
            If int频率间隔 = 1 Then
                '8:00-12:00-14:00；8-12-14
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        
                        If vTmpTime > dat终止时间 Then
                            Exit Do
                        ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                            Exit Do
                        ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            lng次数 = lng次数 - 1
                            If lng次数 = 0 Then Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime > dat终止时间 Then
                            Exit Do
                        ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                            Exit Do
                        ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            lng次数 = lng次数 - 1
                            If lng次数 = 0 Then Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = vCurTime + int频率间隔
        Loop
    ElseIf str间隔单位 = "小时" Then
        '10:00-20:00-40:00；10-20-40；02:30
        arrTime = arrNormal
        Do While lng次数 > 0
            For i = 1 To int频率次数
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime > dat终止时间 Then
                    Exit Do
                ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                    Exit Do
                ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    lng次数 = lng次数 - 1
                    If lng次数 = 0 Then Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + int频率间隔 / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str间隔单位 = "分钟" Then
        '无执行时间
        Do While lng次数 > 0
            vTmpTime = vCurTime
            
            If vTmpTime > dat终止时间 Then
                Exit Do
            ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                Exit Do
            ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                lng次数 = lng次数 - 1
                If lng次数 = 0 Then Exit Do
            End If

            vCurTime = Format(vCurTime + int频率间隔 / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If

    lng次数 = UBound(Split(Mid(strDetailTime, 2), ",")) + 1
    Calc次数分解时间 = Mid(strDetailTime, 2)
End Function

Public Function Calc段内分解时间(ByVal datBegin As Date, ByVal datEnd As Date, ByVal strPause As String, _
    ByVal str执行时间 As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, _
    Optional ByVal dat首日日期 As Date) As String
'功能：按时间段计算各次的分解执行时间及次数
'参数：datBegin-datEnd=要计算的时间段,其中datBegin应为每个周期的开始基准时间
'      strPause=暂停的时间段
'      dat首日日期=用于首日时间计算参照
'返回："时间1,时间2,...."(yyyy-MM-dd HH:mm:ss),时间个数即为次数
'说明：1.时间段内要排除暂停的时间段,次数可能因此而减少
'      2.本函数是假定在执行时间及频率性质完全正确的情况下计算。
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrNormal As Variant, arrFirst As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    If InStr(str执行时间, ",") > 0 Then
        arrNormal = Split(Split(str执行时间, ",")(1), "-")
        arrFirst = Split(Split(str执行时间, ",")(0), "-")
    Else
        arrNormal = Split(str执行时间, "-")
        arrFirst = Array()
    End If
        
    vCurTime = datBegin
    
    If str间隔单位 = "周" Then
        vCurTime = GetWeekBase(datBegin)
        If dat首日日期 <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (vCurTime = GetWeekBase(dat首日日期))
        Else
            blnFirst = False
        End If

        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False
                        
            '1/8:00-3/15:00-5/9:00
            For i = 1 To int频率次数
                If i - 1 <= UBound(arrTime) Then '首周可能次数不足
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                End If
            Next
            vCurTime = Format(vCurTime + 7, "yyyy-MM-dd") '！！！
        Loop
    ElseIf str间隔单位 = "天" Then
        If dat首日日期 <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (Int(vCurTime) = Int(dat首日日期))
        Else
            blnFirst = False
        End If
        
        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False
            
            If int频率间隔 = 1 Then
                '8:00-12:00-14:00；8-12-14
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = Format(vCurTime + int频率间隔, "yyyy-MM-dd") '！！！
        Loop
    ElseIf str间隔单位 = "小时" Then
        '10:00-20:00-40:00；10-20-40；02:30
        arrTime = arrNormal
        Do While vCurTime <= datEnd
            For i = 1 To int频率次数
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + int频率间隔 / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str间隔单位 = "分钟" Then
        '无执行时间
        Do While vCurTime <= datEnd
            vTmpTime = vCurTime
            
            If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                If Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                End If
            ElseIf vTmpTime > datEnd Then
                Exit Do
            End If

            vCurTime = Format(vCurTime + int频率间隔 / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If
    
    Calc段内分解时间 = Mid(strDetailTime, 2)
End Function

Public Function Calc缺省药品总量(ByVal dbl单量 As Double, ByVal int疗程 As Integer, _
    ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, Optional ByVal str执行时间 As String, _
    Optional ByVal dbl剂量系数 As Double, Optional ByVal dbl包装系数 As Double, Optional ByVal int分零 As Integer, Optional ByVal dbl首次用量 As Double) As Double
'功能：按疗程及分零特性计算药品临嘱的缺省总量(或配方缺省付数)
'参数：dbl单量=按剂量单位的一次用量
'      int疗程=一个疗程的天数
'      int分零=0-可分零,1-不分零,2-一次性(即时失效),-N-N天内分零使用有效
'      dbl包装系数=门诊包装或住院包装
'返回：按住院单位计算的药品总量
'说明：
'     1.药品分零特性是针对门诊或住院包装而言的。
'     2.dbl剂量系数,dbl包装系数,int分零=中药不传递,只计算付数
    Dim dbl天次 As Double, dbl总量 As Double
    Dim dbl剩余 As Double, dblOne As Double
    Dim intStep As Integer, dblEnd As Double
    Dim arrTime() As String, strBegin As String
    Dim strTime As String, i As Integer, j As Integer
    
    '疗程不足一个频率周期时就不管疗程
    If str间隔单位 = "周" Then
        If int疗程 < 7 Then int疗程 = 1
    ElseIf str间隔单位 = "天" Then
        If int疗程 < int频率间隔 Then int疗程 = 1
    ElseIf str间隔单位 = "小时" Then
        If int疗程 < int频率间隔 / 24 Then int疗程 = 1
    ElseIf str间隔单位 = "分钟" Then
        If int疗程 < int频率间隔 / (24 * 60) Then int疗程 = 1
    End If
    
    '一个频率周期的次数(按天)
    If str间隔单位 = "周" Then
        dbl天次 = int频率次数 / 7
    ElseIf str间隔单位 = "天" Then
        dbl天次 = int频率次数 / int频率间隔
    ElseIf str间隔单位 = "小时" Then
        dbl天次 = (int频率次数 / int频率间隔) * 24
    ElseIf str间隔单位 = "分钟" Then
        dbl天次 = (int频率次数 / int频率间隔) * (24 * 60)
    End If
    
    If dbl剂量系数 = 0 And dbl包装系数 = 0 Then
        '中药总量(付数) = 单付*疗程*(频率次数/频率间隔)
        dbl总量 = IntEx(int疗程 * dbl天次)
    Else
        '药品临嘱总量 = 门诊/住院包装(单量*疗程*(频率次数/频率间隔))
        If int分零 = 0 Then
            '可分零
            dbl总量 = dbl单量 * int疗程 * dbl天次 / dbl剂量系数 / dbl包装系数
        ElseIf int分零 = 1 Then
            '不分零
            dbl总量 = IntEx(dbl单量 * int疗程 * dbl天次 / dbl剂量系数 / dbl包装系数)
        ElseIf int分零 = 2 Then
            '一次性(即时失效)
            dbl总量 = IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * IntEx(int疗程 * dbl天次)
        ElseIf int分零 < 0 Then
            'ABS(int分零)天内分零使用有效(但不分零计算)
            If str执行时间 <> "" Then
                '一次门诊/住院包装的剂量
                dblOne = IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * (dbl剂量系数 * dbl包装系数)
                '缺省执行的次数和时间分解
                strTime = Calc次数分解时间(IntEx(int疗程 * dbl天次), Date, CDate("3000-01-01"), "", str执行时间, int频率次数, int频率间隔, str间隔单位)
                If strTime <> "" Then
                    arrTime = Split(strTime, ",")
                    dbl剩余 = dblOne: dbl总量 = 1
                    strBegin = arrTime(0)
                    
                    '计算总量
                    For i = 0 To UBound(arrTime)
                        If dbl剩余 < dbl单量 Or CDate(arrTime(i)) - CDate(strBegin) >= Abs(int分零) Then
                            If CDate(arrTime(i)) - CDate(strBegin) >= Abs(int分零) Then
                                dbl剩余 = dblOne
                            Else
                                dbl剩余 = dbl剩余 + dblOne
                            End If
                            dbl总量 = dbl总量 + 1
                            strBegin = arrTime(i)
                        End If
                        dbl剩余 = dbl剩余 - dbl单量
                    Next
                End If
            End If
        End If
    End If
    If dbl总量 > 0 And dbl首次用量 > 0 Then
        dbl总量 = dbl总量 + (dbl首次用量 - dbl单量) / dbl剂量系数 / dbl包装系数
    End If
    Calc缺省药品总量 = dbl总量
End Function

Public Function Calc缺省药品天数(ByVal dbl总量 As Double, ByVal dbl单量 As Double, _
    ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, _
    Optional ByVal dbl剂量系数 As Double, Optional ByVal dbl包装系数 As Double, Optional ByVal int分零 As Integer) As Long
'功能：根据总量，单量，及药品特性计算用药天数
'参数：dbl总量=用户输入的总量
'      dbl单量=按剂量单位的一次用量
'      int分零=0-可分零,1-不分零,2-一次性(即时失效),-N-N天内分零使用有效
'      dbl包装系数=门诊包装或住院包装
'返回：用药天数(中药无天数输入)
    Dim dbl天次 As Double
    Dim lng天数 As Long
    
    '一个频率周期的次数(按天)
    If str间隔单位 = "周" Then
        dbl天次 = int频率次数 / 7
    ElseIf str间隔单位 = "天" Then
        dbl天次 = int频率次数 / int频率间隔
    ElseIf str间隔单位 = "小时" Then
        dbl天次 = (int频率次数 / int频率间隔) * 24
    ElseIf str间隔单位 = "分钟" Then
        dbl天次 = (int频率次数 / int频率间隔) * (24 * 60)
    End If
    
    If int分零 = 0 Then
        '可分零
        'dbl总量 = dbl单量 * int疗程 * dbl天次 / dbl剂量系数 / dbl包装系数
        lng天数 = Format(dbl总量 * dbl包装系数 * dbl剂量系数 / dbl单量 / dbl天次, "0")
    ElseIf int分零 = 1 Then
        '不分零
        'dbl总量 = IntEx(dbl单量 * int疗程 * dbl天次 / dbl剂量系数 / dbl包装系数)
        lng天数 = Format(dbl总量 * dbl包装系数 * dbl剂量系数 / dbl单量 / dbl天次, "0")
    ElseIf int分零 = 2 Then
        '一次性(即时失效)
        'dbl总量 = IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * IntEx(int疗程 * dbl天次)
        lng天数 = Format(dbl总量 / IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) / dbl天次, "0")
    ElseIf int分零 < 0 Then
        'ABS(int分零)天内分零使用有效(但不分零计算)
        lng天数 = Format(dbl总量 * dbl包装系数 * dbl剂量系数 / dbl单量 / dbl天次, "0")
    End If

    Calc缺省药品天数 = lng天数
End Function

Public Function Calc发送药品总量(ByVal dat开始执行时间 As Date, lng次数 As Long, str分解时间 As String, _
    ByVal dbl单量 As Double, ByVal dbl剂量系数 As Double, ByVal dbl包装系数 As Double, _
    ByVal int分零 As Integer, ByVal dat终止时间 As Date, ByVal strPause As String, ByVal str执行时间 As String, _
    ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, _
    Optional ByVal blnLimit As Boolean, Optional ByVal dbl首次用量 As Double, Optional ByVal dat上次执行时间 As Date) As Double
'功能：按发送次数及分零特性计算成药总量
'参数：dat开始执行时间=医嘱的开始执行时间,用于计算下一执行周期开始基准时间
'      lng次数=本次计划要发送的次数
'      dbl单量=按剂量单位的一次用量
'      int分零=0-可分零,1-不分零,2-一次性(即时失效),-N-N天内分零使用有效(按24小时计算)
'      dbl包装系数=门诊包装或住院包装
'      blnLimit=是否按时间限制计算给药途径，不管剩余部份
'下列参数用于不分零药品计算(包括-N型)：
'      str分解时间=本次发送计划执行的分解时间,与次数对应
'      strPause=医嘱的暂停时间段
'      dat终止时间=医嘱的执行终止时间,没有时传入"3000-01-01"
'返回：1.按门诊/住院单位计算的药品总量
'      2.lng次数=不分零药品(包括-N型分零药品)计算后的实际执行次数(增加)
'      3.str分解时间=不分零药品(包括-N型分零药品)计算后的分解时间(增加)
'说明：药品分零特性是针对门诊或住院包装而言的。
    Dim dbl总量 As Double, dbl剩余 As Double
    Dim arrTime() As String, dblOne As Double
    Dim strBegin As String, datBase As Date
    Dim strTmp As String, i As Long
    Dim blnIsFirst As Boolean
    
    '注：一些地方加Val是因为运算结果的Double在某些地方判断时，内部精度有问题，导致比如0.9<>0.9
    If int分零 = 0 Then
        '可分零
        dbl总量 = Val(dbl单量 * lng次数 / dbl剂量系数 / dbl包装系数)
        '如果上次执行时间为NULL，说明包含首次
        If dat上次执行时间 = CDate(0) And dbl首次用量 > 0 Then
            dbl总量 = Val(dbl总量 + (dbl首次用量 - dbl单量) / dbl剂量系数 / dbl包装系数)
        End If
    ElseIf int分零 = 1 Then
        '不分零
        dbl总量 = Val(dbl单量 * lng次数 / dbl剂量系数 / dbl包装系数)
        '如果上次执行时间为NULL，说明包含首次
        If dat上次执行时间 = CDate(0) And dbl首次用量 > 0 Then
            dbl总量 = Val(dbl总量 + (dbl首次用量 - dbl单量) / dbl剂量系数 / dbl包装系数)
        End If
        dbl总量 = Val(IntEx(dbl总量))
        '按不分零计算时,多余的尽可能使用,从而使发送次数增加
        If Not blnLimit Then
            dbl剩余 = Val(dbl总量 * dbl包装系数 * dbl剂量系数 - dbl单量 * lng次数)
            If dbl剩余 >= dbl单量 And dbl单量 <> 0 Then
                '剩余理论可以执行的次数
                i = Int(Val(dbl剩余 / dbl单量))
                '剩余实际可以执行的次数及时间分解(受终止时间限制)
                arrTime = Split(str分解时间, ",")
                datBase = Calc本周期开始时间(dat开始执行时间, CDate(arrTime(UBound(arrTime))), int频率间隔, str间隔单位)
                
                '在往后扩展时间时,最后一个周期内已执行的时间不再计算,按暂停处理
                strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
                If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                
                strTmp = Calc次数分解时间(i, datBase, dat终止时间, strPause, str执行时间, int频率次数, int频率间隔, str间隔单位, dat开始执行时间)
                If strTmp <> "" Then
                    lng次数 = lng次数 + i
                    str分解时间 = str分解时间 & "," & strTmp
                End If
            End If
        End If
    ElseIf int分零 = 2 Then
        '一次性(即时失效)
        dbl总量 = Val(IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * lng次数)
        '如果上次执行时间为NULL，说明包含首次
        If dat上次执行时间 = CDate(0) And dbl首次用量 > 0 Then
            dbl总量 = Val(dbl总量 + IntEx(dbl首次用量 / dbl剂量系数 / dbl包装系数) - IntEx(dbl单量 / dbl剂量系数 / dbl包装系数))
        End If
    ElseIf int分零 < 0 Then
        'ABS(int分零)天内分零使用有效(但不分零计算)
        arrTime = Split(str分解时间, ",")
        strBegin = arrTime(0)
        
        '一次门诊/住院包装的剂量(剂量单位)
        dblOne = Val(IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * (dbl剂量系数 * dbl包装系数))
        '一次门诊/住院包装的剂量(包装单位)
        dbl总量 = Val(IntEx(dbl单量 / dbl剂量系数 / dbl包装系数))
        '如果上次执行时间为NULL，说明包含首次
        If dat上次执行时间 = CDate(0) And dbl首次用量 > 0 Then
            dbl总量 = Val(IntEx(dbl首次用量 / dbl剂量系数 / dbl包装系数))
            dblOne = Val(IntEx(dbl首次用量 / dbl剂量系数 / dbl包装系数) * (dbl剂量系数 * dbl包装系数))
            blnIsFirst = True
        End If
         '计算总量
        dbl剩余 = dblOne
        For i = 0 To UBound(arrTime)
            '第一次循环肯定够,所以不进入条件
            If dbl剩余 < IIf(blnIsFirst, dbl首次用量, dbl单量) Or CDate(arrTime(i)) - CDate(strBegin) >= Abs(int分零) Then
                If CDate(arrTime(i)) - CDate(strBegin) >= Abs(int分零) Then
                    dbl剩余 = dblOne
                    dbl总量 = dbl总量 + IntEx(dbl单量 / dbl剂量系数 / dbl包装系数)
                Else
                    If dbl剩余 + dbl剂量系数 * dbl包装系数 >= IIf(blnIsFirst, dbl首次用量, dbl单量) Then
                        '只需剩余加一个包装单位即够
                        dbl剩余 = dbl剩余 + dbl剂量系数 * dbl包装系数
                        dbl总量 = dbl总量 + 1
                    Else
                        '需要剩余加一次包装单位才够
                        dbl剩余 = dbl剩余 + dblOne
                        dbl总量 = dbl总量 + IntEx(IIf(blnIsFirst, dbl首次用量, dbl单量) / dbl剂量系数 / dbl包装系数)
                    End If
                End If
                strBegin = arrTime(i)
            End If
            dbl剩余 = dbl剩余 - IIf(blnIsFirst, dbl首次用量, dbl单量)
            If blnIsFirst Then
                blnIsFirst = False
                dblOne = Val(IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * (dbl剂量系数 * dbl包装系数))
            End If
        Next
        
        '剩余部分继续在有效期内按不分零计算,从而使发送次数增加
        If Not blnLimit Then
            If dbl剩余 >= dbl单量 And dbl单量 <> 0 Then
                '剩余理论可以执行的次数
                i = Int(Val(dbl剩余 / dbl单量))
                '剩余实际可以执行的次数及时间分解(受终止时间限制)
                datBase = Calc本周期开始时间(dat开始执行时间, CDate(arrTime(UBound(arrTime))), int频率间隔, str间隔单位)
                
                '在往后扩展时间时,最后一个周期内已执行的时间不再计算,按暂停处理
                strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
                If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                
                strTmp = Calc次数分解时间(i, datBase, dat终止时间, strPause, str执行时间, int频率次数, int频率间隔, str间隔单位, dat开始执行时间)
                If strTmp <> "" Then
                    arrTime = Split(strTmp, ",")
                    For i = 0 To UBound(arrTime)
                        If dbl剩余 < dbl单量 Or CDate(arrTime(i)) - CDate(strBegin) >= Abs(int分零) Then
                            Exit For
                        End If
                        lng次数 = lng次数 + 1
                        str分解时间 = str分解时间 & "," & arrTime(i)
                        dbl剩余 = dbl剩余 - dbl单量
                    Next
                End If
            End If
        End If
    End If
    
    Calc发送药品总量 = dbl总量
End Function

Public Function Calc本周期开始时间(ByVal dat开始执行时间 As Date, ByVal dat某次执行时间 As Date, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String) As Date
'功能：根据长嘱的某次执行时间，得到它在该周期内的开始基准时间
    Dim datBegin As Date, datCurr As Date
    
    datCurr = dat开始执行时间
    datBegin = datCurr
    If str间隔单位 = "周" Then datCurr = GetWeekBase(datCurr)
    
    If str间隔单位 = "" Then Exit Function
    Do While datCurr <= dat某次执行时间
        datBegin = datCurr
        If str间隔单位 = "周" Then
            datCurr = datCurr + 7
        ElseIf str间隔单位 = "天" Then
            datCurr = datCurr + int频率间隔
        ElseIf str间隔单位 = "小时" Then
            datCurr = DateAdd("h", int频率间隔, datCurr)
        ElseIf str间隔单位 = "分钟" Then
            datCurr = DateAdd("n", int频率间隔, datCurr)
        End If
    Loop
    Calc本周期开始时间 = datBegin
End Function

Public Function Trim分解时间(ByVal lng次数 As Long, ByVal str分解时间 As String) As String
'功能：将医嘱执行的分解时间按次数进行截断
    Dim arrTime() As String, strTmp As String, i As Long
    
    arrTime = Split(str分解时间, ",")
    For i = 0 To lng次数 - 1
        strTmp = strTmp & "," & arrTime(i)
    Next
    Trim分解时间 = Mid(strTmp, 2)
End Function

Public Function Calc持续性长嘱次数(ByVal datBegin As Date, ByVal datEnd As Date, _
    ByVal str上次执行时间 As String, ByVal str执行终止时间 As String, _
    ByVal strPause As String, Optional str首次时间 As String, _
    Optional str末次时间 As String, Optional str分解时间 As String) As Long
'功能：对持续性非药长嘱计算它本次应该发送的次数,及首末时间
'参数：str上次执行时间=不一定等于本次发送的开始时间
'      str执行终止时间=
'返回：本次该医嘱发送的次数
'      str首次时间,str末次时间=返回yyyy-MM-dd HH:mm:ss
'说明：持续性长嘱按日期每天发送一次处理,处理规则与床位费类似(暂停时间算始不算止)
    Dim curDate As Date, lng次数 As Long, blnSend As Boolean
    
    str首次时间 = "": str末次时间 = "": str分解时间 = ""
    curDate = CDate(Format(datBegin, "yyyy-MM-dd"))
    Do While curDate <= CDate(Format(datEnd, "yyyy-MM-dd"))
        If Not DateIsPause(curDate, strPause) Then
            blnSend = True
            If str上次执行时间 <> "" Then
                If Format(curDate, "yyyy-MM-dd") <= Format(str上次执行时间, "yyyy-MM-dd") Then
                    blnSend = False '应大于上次执行时间才执行
                End If
            End If
            If str执行终止时间 <> "" Then
                If Format(curDate, "yyyy-MM-dd") > Format(str执行终止时间, "yyyy-MM-dd") Then
                    blnSend = False '应小于等于执行终止时间才执行
                End If
            End If
            If blnSend Then
                lng次数 = lng次数 + 1
                If lng次数 = 1 Then
                    str首次时间 = Format(curDate, "yyyy-MM-dd 00:00:00") '定为零点执行
                    If str首次时间 < Format(datBegin, "yyyy-MM-dd HH:mm:ss") Then
                        str首次时间 = Format(datBegin, "yyyy-MM-dd HH:mm:ss")
                    End If
                    str末次时间 = str首次时间
                    str分解时间 = str首次时间
                Else
                    str末次时间 = Format(curDate, "yyyy-MM-dd 00:00:00")
                    str分解时间 = str分解时间 & "," & str末次时间
                End If
            End If
        End If
        curDate = curDate + 1
    Loop
    
    Calc持续性长嘱次数 = lng次数
End Function

 Public Function Calc总量单量天数(ByVal dbl总量 As Double, ByVal dbl单量 As Double, ByVal dbl剂量系数 As Double, ByVal dbl包装系数 As Double, _
    ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String) As Double
'功能：根据指定的总量、单量、频率计算药品可以使用的天数
    Dim dbl天次 As Double
    Dim dbl总单量 As Double
    
    '一个频率周期的次数(按天)
    If str间隔单位 = "周" Then
        dbl天次 = int频率次数 / 7
    ElseIf str间隔单位 = "天" Then
        dbl天次 = int频率次数 / int频率间隔
    ElseIf str间隔单位 = "小时" Then
        dbl天次 = (int频率次数 / int频率间隔) * 24
    ElseIf str间隔单位 = "分钟" Then
        dbl天次 = (int频率次数 / int频率间隔) * (24 * 60)
    End If
    
    dbl总单量 = dbl总量 * dbl包装系数 * dbl剂量系数
    
    Calc总量单量天数 = dbl总单量 / dbl单量 / dbl天次
End Function

Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str姓名 As String, ByVal cur剩余款额 As Currency, _
    ByVal cur当日金额 As Currency, ByVal cur记帐金额 As Currency, ByVal cur担保金额 As Currency, _
    ByVal str收费类别 As String, ByVal str类别名称 As String, str已报类别 As String, _
    intWarn As Integer, Optional ByVal bln划价 As Boolean, _
    Optional blnNotCheck类别 As Boolean = False) As Integer
'功能:对病人记帐进行报警提示
'参数:rsWarn=包含报警参数设置的记录集(该病人病区,并区分好了医保)
'     str收费类别=当前要检查的类别,用于分类报警
'     str类别名称=类别名称,用于提示
'     bln划价=生成划价费用时的报警，类似具有欠费强制记帐权限时的处理
'     intWarn=是否显示询问性的提示,-1=要显示,0=缺省为否,1-缺省为是
'     blnNotCheck类别:不对类别进行检查(主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
'返回:str已报类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
'     intWarn=本次询问性提示中的选择结果,0=为否,1-为是
'     0;没有报警,继续
'     1:报警提示后用户选择继续
'     2:报警提示后用户选择中断
'     3:报警提示必须中断
'     4:强制记帐报警,继续
    Dim bln已报警 As Boolean, byt标志 As Byte
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim arrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str担保 As String, i As Long
    
    BillingWarn = 0
    
    '报警参数检查:NULL是没有设置,0是设置了的
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str收费类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志1 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str收费类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志2 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str收费类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志3 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    If byt标志 = 0 Then Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    '示例："-" 或 ",ABC,567,DEF"
    '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
    bln已报警 = InStr(str已报类别, str收费类别) > 0 Or str已报类别 Like "-*"
    
    If bln已报警 Then '当intWarn = -1时,也可强行再报警
        If byt标志 = 2 Then
            If str已报类别 Like "-*" Then
                byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
            Else
                arrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(arrTmp)
                    If InStr(arrTmp(i), str收费类别) > 0 Then
                        byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                        'Exit For '取消说明见住院记帐模块
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str类别名称 <> "" Then str类别名称 = """" & str类别名称 & """费用"
    str担保 = IIf(cur担保金额 = 0, "", "(含担保额:" & Format(cur担保金额, "0.00") & ")")
    cur剩余款额 = cur剩余款额 + cur担保金额 - cur记帐金额
    cur当日金额 = cur当日金额 + cur记帐金额
        
    '---------------------------------------------------------------------
    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur剩余款额 < 0 Then
                        byt方式 = 2
                        If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽,允许该病人记帐吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    ElseIf cur剩余款额 < rsWarn!报警值 Then
                        byt方式 = 1
                        If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur剩余款额 < 0 Then
                            byt方式 = 2
                            If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽,允许该病人记帐吗？", frmParent)
                                    If vMsg = vbNo Or vMsg = vbCancel Then
                                        If vMsg = vbCancel Then intWarn = 0
                                        BillingWarn = 2
                                    ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                        If vMsg = vbIgnore Then intWarn = 1
                                        BillingWarn = 4
                                    End If
                                Else
                                    If intWarn = 0 Then
                                        BillingWarn = 2
                                    ElseIf intWarn = 1 Then
                                        BillingWarn = 4
                                    End If
                                End If
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    
    '对于继续类的操作,返回已报警类别
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志3
            End If
        End If
    End If
End Function

Public Sub GetPatiLastChange(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByRef lng病区ID As Long, ByRef lng科室ID As Long, Optional ByVal int场合 As Integer = -1, Optional ByRef strTurnDate As String)
'功能：获取病人最近的转科或转病区信息
'参数：int场合 -1-参数不传；0-医生站，1-护士站
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    If int场合 = -1 Or int场合 = 1 Then
        strSQL = " And (终止原因 = 3 Or 终止原因 = 15)"
    ElseIf int场合 = 0 Or int场合 = 2 Then
        strSQL = " And (终止原因 = 3 )"
    End If
    
    strSQL = "Select 病区id, 科室id,终止时间" & vbNewLine & _
        "From (Select 病区id, 科室id,终止时间" & vbNewLine & _
        "       From 病人变动记录" & vbNewLine & _
        "       Where 病人id = [1] And 主页id = [2]  And 终止时间 Is Not Null" & strSQL & vbNewLine & _
        "       Order By 终止时间 Desc)" & vbNewLine & _
        "Where Rownum = 1"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiLastChange", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        lng病区ID = Val("" & rsTmp!病区ID)
        lng科室ID = Val("" & rsTmp!科室ID)
        strTurnDate = Format(rsTmp!终止时间, "yyyy-MM-dd HH:mm:ss")
    Else
        lng病区ID = 0
        lng科室ID = 0
        strTurnDate = ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetPatiDayMoney(lng病人ID As Long) As Currency
'功能：获取指定病人当天发生的费用总额
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!金额, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiCanBilling(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String, Optional ByVal lngModual As Long) As Boolean
'功能：检查指定病人是否具有相关权限
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    
    If InStr(strPrivs, "出院未结强制记帐") > 0 And InStr(strPrivs, "出院结清强制记帐") > 0 Then
        Exit Function
    End If
    On Error GoTo errH
    strSQL = "Select NVL(B.姓名,A.姓名) 姓名,B.出院日期,B.状态,X.费用余额" & _
        " From 病人信息 A,病案主页 B,病人余额 X" & _
        " Where A.病人ID=B.病人ID And A.病人ID=X.病人ID(+) And X.类型(+) = 2" & _
        " And A.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!出院日期) And Nvl(rsTmp!状态, 0) <> 3 Then Exit Function
        If InStr(strPrivs, "出院未结强制记帐") = 0 Then
            If Nvl(rsTmp!费用余额, 0) <> 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用未结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If InStr(strPrivs, "出院结清强制记帐") = 0 Then
            If Nvl(rsTmp!费用余额, 0) = 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用已结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If lngModual = p医嘱附费管理 Or lngModual = p住院医嘱发送 Or lngModual = p住院医嘱下达 Then
            '68081不允许出院病人处理医嘱费用
            strMsg = """" & rsTmp!姓名 & """已经出院(或预出院)，不能对该病人的医嘱进行发送、超期收回、执行、回退。"
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckEPRReport(ByVal lng医嘱ID As Long, Optional lng报告ID As Long, Optional blnBySign As Boolean, Optional ByVal int执行状态 As Integer = -999) As Integer
'功能：检查对应项目的报告填写情况
'参数：lng医嘱ID=可见行的医嘱ID
'      lng报告ID=可以传入，主要用于返回报告病历ID
'      int执行状态=用于检验完成时，传入综合的执行状态
'参数：blnBySign=报告是否完成通过签名级别判断(用于医技工作站)
'返回：0-报告还没有填写
'      1-报告已填写完成(已签名,包括修订后签名,或已执行完成)
'      2-报告未填写完成(未签名,或修订后未签名,且未执行完成)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str检查报告ID As String
    
    On Error GoTo errH
    
    '检查报告是否已书写
    If lng报告ID = 0 Then
        strSQL = "Select 病历ID,检查报告ID || ''  as 检查报告ID From 病人医嘱报告 Where 医嘱ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng医嘱ID)
        If Not rsTmp.EOF Then lng报告ID = Val(rsTmp!病历id & ""): str检查报告ID = rsTmp!检查报告ID & ""
    End If
    If lng报告ID = 0 And str检查报告ID = "" Then
        CheckEPRReport = 0: Exit Function
    End If
    
    If Not blnBySign Then
        '检查报告执行过程(5-审核;6-报告完成)和状态(1-完成)
        '检验报告是关联到采集方式上面的，但采集方式可能为叮嘱未产生发送记录
        strSQL = _
            " Select 2 as 排序,医嘱ID,执行过程,执行状态,发送时间 From 病人医嘱发送 Where 医嘱ID=[1]" & _
            " Union ALL" & _
            " Select 排序,医嘱ID,执行过程,Decode([2],-999,执行状态,[2]) as 执行状态,发送时间" & _
            " From (" & _
                " Select 1 as 排序,B.医嘱ID,B.执行过程,B.执行状态,B.发送时间 From 病人医嘱记录 A,病人医嘱发送 B" & _
                " Where A.ID=B.医嘱ID And A.相关ID=(" & _
                    " Select A.ID From 病人医嘱记录 A,诊疗项目目录 B Where A.ID=[1] And A.诊疗项目ID=B.ID And A.诊疗类别='E' And B.操作类型='6')" & _
                " Order by A.序号" & _
            " ) Where Rownum=1" & _
            " Order by 排序,发送时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng医嘱ID, int执行状态)
        If Nvl(rsTmp!执行过程, 0) >= 5 Or Nvl(rsTmp!执行状态, 0) = 1 Then
            CheckEPRReport = 1
        Else
            CheckEPRReport = 2
        End If
    Else
        '通过签名版本判断报告完成的方式
        strSQL = "Select B.文件ID,Max(B.开始版) as 签名版本 From 电子病历内容 B Where B.文件ID=[1] And B.对象类型=8 Group by B.文件ID"
        strSQL = "Select B.完成时间,B.最后版本,C.签名版本 From 电子病历记录 B,(" & strSQL & ") C Where B.ID=[1] And B.ID=C.文件ID(+)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng报告ID)
            
        '(签名后不能直接修改，除非修订；因此签名后最后版本应与签名版本一致)
        If IsNull(rsTmp!完成时间) Or Nvl(rsTmp!最后版本, 0) <> Nvl(rsTmp!签名版本, 0) Then
            '如果医嘱本身已经执行,即使没有签名或不符也视同完成
            strSQL = _
                " Select 2 as 排序,医嘱ID,执行状态,发送时间 From 病人医嘱发送 Where 医嘱ID=[1]" & _
                " Union ALL" & _
                " Select 排序,医嘱ID,Decode([2],-999,执行状态,[2]) as 执行状态,发送时间" & _
                " From (" & _
                    " Select 1 as 排序,B.医嘱ID,B.执行状态,B.发送时间 From 病人医嘱记录 A,病人医嘱发送 B" & _
                    " Where A.ID=B.医嘱ID And A.相关ID=(" & _
                        " Select A.ID From 病人医嘱记录 A,诊疗项目目录 B Where A.ID=[1] And A.诊疗项目ID=B.ID And A.诊疗类别='E' And B.操作类型='6')" & _
                    " Order by A.序号" & _
                " ) Where Rownum=1" & _
                " Order by 排序,发送时间 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng医嘱ID, int执行状态)
            If Nvl(rsTmp!执行状态, 0) = 1 Then
                CheckEPRReport = 1
            Else
                CheckEPRReport = 2
            End If
        Else
            CheckEPRReport = 1
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub GetTestLabel(ByVal strScript As String, ByVal strSelect As String, strLabel As String, intResult As Integer)
'功能：获取皮试标注和结果
'参数：strScript=皮试结果描述串，如"阳性(+),大阳性(++);阴性(-)"
'      strSelect=所选择的皮试结果中文名，如"阳性"
'返回：strLabel = 皮试结果标注，如"(+)"
'      intResult=皮试结果：0-阴性，1-阳性
    Dim arr阳性 As Variant, arr阴性 As Variant
    Dim i As Integer
    
    strLabel = "": intResult = 0
    
    arr阳性 = Split(Split(strScript, ";")(0), ",")
    arr阴性 = Split(Split(strScript, ";")(1), ",")
    
    For i = 0 To UBound(arr阳性)
        If arr阳性(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr阳性(i), Len(strSelect) + 1)
            intResult = 1: Exit Sub
        End If
    Next
    For i = 0 To UBound(arr阴性)
        If arr阴性(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr阴性(i), Len(strSelect) + 1)
            intResult = 0: Exit Sub
        End If
    Next
End Sub


Public Function ItemHaveCash(ByVal int病人来源 As Integer, ByVal bln单独执行 As Boolean, ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, _
    ByVal lng发送号 As Long, ByVal str类别 As String, ByVal str单据号 As String, ByVal int记录性质 As Integer, ByVal int门诊记帐 As Integer, ByVal int方式 As Integer, _
    Optional ByVal blnMove As Boolean, Optional ByVal dat发送时间 As Date, Optional ByRef str医嘱IDs As String, Optional ByRef strNos As String, Optional ByRef blnIsAbnormal As Boolean) As Boolean
'功能：判断当前的执行医嘱是否已收费或记帐划价单是否已审核
'参数：int病人来源=1-门诊,2-住院
'      str类别=诊疗类别，用于从一组医嘱中区分分开执行的内容
'      int方式=0-检查是否存在未收费记录
'              1-检查是否存在已收费记录
'      int门诊记帐=1=住院发送到门诊记帐
'      返回：str医嘱IDs=该医嘱及相关的医嘱ID,NOs=医嘱发送的单据号和补的附费中的单据号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTab As String
    
    If int病人来源 = 2 And int记录性质 = 2 And int门诊记帐 = 0 Then
        strTab = "住院费用记录"
    Else
        strTab = "门诊费用记录"
    End If
    ItemHaveCash = True
    str医嘱IDs = ""
    strNos = ""
    
    '对应的费用中是否存在未收费[或已作废]的内容
    '和清单只显示已收费不同：
    '1.检查了医嘱附费(不加记录性质的条件，因为可能补收费单或记帐单)
    '2.记帐划价也显示为未收(清单需要先显出来执行后审核)
    '3.按NO对应到相关医嘱的费用检查(清单是按显示的医嘱ID)
    strSQL = _
        " Select A.记录状态,Nvl(B.相关ID,B.ID) as 医嘱ID,B.诊疗类别,A.执行状态,nvl(a.结帐id,0) as 结帐id,A.NO" & IIf(strTab = "住院费用记录", ",0 as 费用状态", ",NVL(A.费用状态,0) as 费用状态") & _
        " From " & strTab & " A,病人医嘱记录 B" & _
        " Where A.NO=[4] And A.记录状态 IN(0,1,3) And A.医嘱序号+0=B.ID And MOD(A.记录性质,10)=[5]" & IIf(bln单独执行, " And B.ID=[2]", "") & _
        " Union ALL " & _
        " Select B.记录状态,Nvl(C.相关ID,C.ID) as 医嘱ID,C.诊疗类别,B.执行状态,nvl(b.结帐id,0) as 结帐id,A.NO" & IIf(strTab = "住院费用记录", ",0 as 费用状态", ",NVL(b.费用状态,0) as 费用状态") & _
        " From 病人医嘱记录 C," & strTab & " B,病人医嘱附费 A" & _
        " Where A.NO=B.NO And A.记录性质=MOD(B.记录性质,10) And A.医嘱ID=B.医嘱序号+0" & IIf(bln单独执行, " And A.医嘱ID=[2]", _
            " And A.医嘱ID IN (Select ID From 病人医嘱记录 Where (ID=[1] Or 相关ID=[1]) And 诊疗类别=[6])") & _
        " And A.发送号=[3] And B.记录状态 IN(0,1,3) And A.医嘱ID=C.ID And A.记录性质=[5]"
    If blnMove Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
        strSQL = Replace(strSQL, strTab, "H" & strTab)
    ElseIf zlDatabase.DateMoved(dat发送时间) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, strTab, "H" & strTab)
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ItemHaveCash", IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID), lng医嘱ID, lng发送号, str单据号, int记录性质, str类别)
    If Not rsTmp.EOF Then
        If int方式 = 0 Then
            rsTmp.Filter = "医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 费用状态=1 and 结帐id<>0"
            If Not rsTmp.EOF Then
                blnIsAbnormal = True
                ItemHaveCash = False
            Else
                rsTmp.Filter = "(医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 记录状态=0)" & _
                     " or (医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 记录状态=1 and 结帐id=0)"
                If Not rsTmp.EOF Then ItemHaveCash = False
            End If
            
            While Not rsTmp.EOF
                If InStr("," & str医嘱IDs & ",", "," & rsTmp!医嘱ID & ",") = 0 Then
                    str医嘱IDs = str医嘱IDs & "," & rsTmp!医嘱ID
                End If
                If InStr("," & strNos & ",", "," & rsTmp!NO & ",") = 0 Then
                    strNos = strNos & "," & rsTmp!NO
                End If
                rsTmp.MoveNext
            Wend
            strNos = Mid(strNos, 2)
            str医嘱IDs = Mid(str医嘱IDs, 2)
        ElseIf int方式 = 1 Then
            rsTmp.Filter = "医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 记录状态<>1 And 费用状态<>1"
            If Not rsTmp.EOF Then ItemHaveCash = False
        End If
    ElseIf int方式 = 1 Then
        ItemHaveCash = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceMoney(ByVal str组ID As String, ByVal str医嘱ID As String, ByVal str发送号 As String, _
    str类别 As String, str类别名 As String, ByVal bln单独执行 As Boolean, ByVal byt来源 As Byte) As Currency
'功能：根据指定的医嘱ID串，获取医嘱对应未审核的记帐费用合计
'参数：str组ID,str医嘱ID,str发送号="ID1,ID2,..."
'      bln单独执行=检验项目单独执行，这时只有一个医嘱ID
'      byt来源，1:门诊，2-住院
'返回：str类别,str类别名=用于报警提示
'说明：当系统参数为执行后审核费用时才返回。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, curMoney As Currency
    Dim strTab As String
    
    str类别 = "": str类别名 = ""
    
    On Error GoTo errH
    
    If zlDatabase.GetPara(81, glngSys) <> "1" Then Exit Function
    strTab = IIf(byt来源 = 1, "门诊费用记录", "住院费用记录")
    
    If bln单独执行 Then
        strSQL = _
            " Select B.编码,B.名称,Sum(A.实收金额) as 金额" & _
            " From " & strTab & " A,收费项目类别 B" & _
            " Where A.医嘱序号 + 0 = [2] And (A.记录性质, A.NO) In" & _
            "      (Select 记录性质, NO From 病人医嘱附费 Where 医嘱id = [2] And 发送号 + 0 = [3]" & _
            "       Union All" & _
            "       Select 记录性质, NO From 病人医嘱发送 Where 医嘱id = [2] And 发送号 + 0 = [3])" & _
            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别=B.编码" & _
            " Group by B.编码,B.名称"
    Else
        strSQL = _
            " Select /*+ RULE */ B.编码,B.名称,Sum(A.实收金额) as 金额" & _
            " From " & strTab & " A,收费项目类别 B" & _
            " Where A.医嘱序号 + 0 In" & _
            "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
            "       Union All" & _
            "       Select ID From 病人医嘱记录" & _
            "       Where 相关id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "  And (A.记录性质, A.NO) In" & _
            "      (Select 记录性质, NO From 病人医嘱附费" & _
            "       Where 医嘱id In" & _
                "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
                "       Union All" & _
                "       Select ID From 病人医嘱记录" & _
                "       Where 相关id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "         And 发送号 + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)))" & _
            "       Union All" & _
            "       Select 记录性质, NO From 病人医嘱发送" & _
            "       Where 医嘱id In (Select Column_Value From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))" & _
            "         And 发送号 + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))))" & _
            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别=B.编码" & _
            " Group by B.编码,B.名称"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str组ID, str医嘱ID, str发送号, glngSys)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!金额, 0)
        str类别 = str类别 & rsTmp!编码
        str类别名 = str类别名 & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    
    str类别名 = Mid(str类别名, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceStuffMoney(ByVal str组ID As String, ByVal str医嘱ID As String, _
    ByVal str发送号 As String, ByVal bln单独执行 As Boolean, ByVal int病人来源 As Integer, ByVal int记录性质 As Integer, ByVal int门诊记帐 As Integer) As Currency
'功能：根据指定的医嘱ID串，获取医嘱对应未审核的跟踪在用卫材记帐费用合计
'参数：str组ID,str医嘱ID,str发送号="ID1,ID2,..."
'      bln单独执行=检验项目单独执行，这时只有一个医嘱ID
'      int病人来源，1:门诊，2-住院
'      int门诊记帐=1=住院发送到门诊记帐
'返回：str类别,str类别名=用于报警提示
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    On Error GoTo errH
    If int病人来源 = 2 And int记录性质 = 2 And int门诊记帐 = 0 Then
        strTab = "住院费用记录"
    Else
        strTab = "门诊费用记录"
    End If
    
    If bln单独执行 Then
        strSQL = _
            " Select Sum(A.实收金额) as 金额" & _
            " From " & strTab & " A,材料特性 B" & _
            " Where A.医嘱序号 + 0 = [2] And (A.记录性质, A.NO) In" & _
            "      (Select 记录性质, NO From 病人医嘱附费 Where 医嘱id = [2] And 发送号 + 0 = [3]" & _
            "       Union All" & _
            "       Select 记录性质, NO From 病人医嘱发送 Where 医嘱id = [2] And 发送号 + 0 = [3])" & _
            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别='4' And A.收费细目ID=B.材料ID And B.跟踪在用=1"
    Else
        strSQL = _
            " Select /*+ RULE */ Sum(A.实收金额) as 金额" & _
            " From " & strTab & " A,材料特性 B" & _
            " Where A.医嘱序号 + 0 In" & _
            "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
            "       Union All" & _
            "       Select ID From 病人医嘱记录" & _
            "       Where 相关id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "  And (A.记录性质, A.NO) In" & _
            "      (Select 记录性质, NO From 病人医嘱附费" & _
            "       Where 医嘱id In" & _
                "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
                "       Union All" & _
                "       Select ID From 病人医嘱记录" & _
                "       Where 相关id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "         And 发送号 + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)))" & _
            "       Union All" & _
            "       Select 记录性质, NO From 病人医嘱发送" & _
            "       Where 医嘱id In (Select Column_Value From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))" & _
            "         And 发送号 + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))))" & _
            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别='4' And A.收费细目ID=B.材料ID And B.跟踪在用=1"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceStuffMoney", str组ID, str医嘱ID, str发送号, glngSys)
    If Not rsTmp.EOF Then GetAdviceStuffMoney = Nvl(rsTmp!金额, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FinishBillingWarn(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng病人ID As Long, _
    ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal cur金额 As Currency, ByVal str类别 As String, ByVal str类别名 As String) As Boolean
'功能：当执行完成有自动审核的费用时，对病人费用进行记帐报警。
'参数：str类别="CDE..."，报警金额涉及到的收费类别
'      str类别名="检查,检验,..."，对应的类别名用于提示
    Dim rsPati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strSQL As String, intR As Integer, i As Long
    Dim cur当日 As Currency
    
    On Error GoTo errH
    
    If lng主页ID <> 0 Then
        '住院病人报警
        strSQL = _
            " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1] And 类型 = 2" & _
            " Union ALL" & _
            " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
        strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
        
        strSQL = "Select NVL(B.姓名,A.姓名) 姓名,zl_PatiWarnScheme(A.病人ID,B.主页ID) as 适用病人,C.剩余款," & _
            " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额" & _
            " From 病人信息 A,病案主页 B,(" & strSQL & ") C" & _
            " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID(+)" & _
            " And A.病人ID=[1] And B.主页ID=[2]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病人ID, lng主页ID)
    Else
        '其他按门诊报警
        strSQL = "Select 病人ID,预交余额,费用余额 From 病人余额 Where 性质=1 And 病人ID=[1] And 类型 = 1"
        strSQL = "Select A.姓名,zl_PatiWarnScheme(A.病人ID) as 适用病人,A.担保额," & _
            " Nvl(B.预交余额,0)-Nvl(B.费用余额,0) as 剩余款" & _
            " From 病人信息 A,(" & strSQL & ") B" & _
            " Where A.病人ID=B.病人ID(+)" & _
            " And A.病人ID=[1]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病人ID)
    End If
    
    intWarn = -1 '记帐报警时缺省要提示
    '执行报警:门诊病人病区ID=0
    strSQL = "Select Nvl(报警方法,1) as 报警方法,报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线 Where Nvl(病区ID,0)=[1] And 适用病人=[2]"
    Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病区ID, CStr(Nvl(rsPati!适用病人)))
    If Not rsWarn.EOF Then
        If rsWarn!报警方法 = 2 Then cur当日 = GetPatiDayMoney(lng病人ID)
        str类别名 = Mid(str类别名, 2)
        For i = 1 To Len(str类别)
            intR = BillingWarn(frmParent, strPrivs, rsWarn, Nvl(rsPati!姓名), Nvl(rsPati!剩余款, 0), cur当日, cur金额, Nvl(rsPati!担保额, 0), Mid(str类别, i, 1), Split(str类别名, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ItemCanCancel(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal lng组ID As Long, str诊疗类别 As String, _
    ByVal bln单独执行 As Boolean, ByVal blnMove As Boolean, ByVal byt来源 As Byte) As Boolean
'功能：判断指定项目是否可以取消执行
'参数：byt来源=1:门诊，2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If gbytBillOpt = 0 Then ItemCanCancel = True: Exit Function
    
    On Error GoTo errH
    
    If bln单独执行 Then
        strSQL = _
            " Select Distinct NO From 病人医嘱发送 Where 记录性质=2 And 医嘱ID=[1] And 发送号=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From 病人医嘱附费 Where 记录性质=2 And 医嘱ID=[1] And 发送号=[2]"
    Else
        strSQL = _
            " Select Distinct NO From 病人医嘱发送 Where 记录性质=2 And 医嘱ID=[1] And 发送号=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From 病人医嘱附费 Where 记录性质=2 And 发送号=[2]" & _
            " And 医嘱ID IN(Select ID From 病人医嘱记录 Where (ID=[3] Or 相关ID=[3]) And 诊疗类别=[4])"
    End If
    If blnMove Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ItemCanCancel", lng医嘱ID, lng发送号, lng组ID, str诊疗类别)
    
    Do While Not rsTmp.EOF
        '处理中排开了结帐金额为0的，即零耗费用登记
        If HaveBilling(rsTmp!NO, True, "", IIf(bln单独执行, lng医嘱ID, 0), byt来源) <> 0 Then
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("该项目包含已经结帐的费用,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
                    MsgBox "该项目包含已经结帐的费用,操作不能继续。", vbExclamation, gstrSysName
                    Exit Function
            End Select
        End If
        rsTmp.MoveNext
    Loop
    ItemCanCancel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get对应科室IDs(ByVal lng病区ID As Long) As Recordset
'功能：获取病区对应的科室
    Dim strSQL As String, i As Long

    strSQL = " Select B.科室ID From 病区科室对应 B" & _
            " Where B.病区ID=[1]"
    On Error GoTo errH
    Set Get对应科室IDs = zlDatabase.OpenSQLRecord(strSQL, "Get对应科室IDs", lng病区ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser科室IDs(Optional ByVal bln病区 As Boolean) As String
'功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
'参数：是否取所属病区下的科室
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
    
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    '没有强制限制临床,可能医技科室用
    If blnNew Then
        strSQL = "Select 1 as 类别,部门ID From 部门人员 Where 人员ID=[1] Union" & _
                " Select Distinct 2 as 类别,B.科室ID From 部门人员 A,病区科室对应 B" & _
                " Where A.部门ID=B.病区ID And A.人员ID=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", UserInfo.ID)
    End If
    If bln病区 = False Then
        rsTmp.Filter = "类别 = 1"
    Else
        rsTmp.Filter = ""
    End If
    
    For i = 1 To rsTmp.RecordCount
        If InStr("," & GetUser科室IDs & ",", "," & rsTmp!部门ID & ",") = 0 Then
            GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
        End If
        rsTmp.MoveNext
    Next
    GetUser科室IDs = Mid(GetUser科室IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser病区IDs() As String
'功能：获取操作员所属的病区(直接属于病区或所在科室所属的病区),可能有多个
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
        
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    If blnNew Then
        strSQL = _
            "Select Distinct 病区ID From (" & _
            " Select A.部门ID as 病区ID" & _
            " From 部门性质说明 A,部门人员 B" & _
            " Where A.部门ID=B.部门ID And B.人员ID=[1]" & _
            " And A.服务对象 in(1,2,3) And A.工作性质='护理'" & _
            " Union" & _
            " Select A.病区ID From 病区科室对应 A,部门人员 B" & _
            " Where A.科室ID=B.部门ID And B.人员ID=[1])"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
    ElseIf rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
    End If
    For i = 1 To rsTmp.RecordCount
        GetUser病区IDs = GetUser病区IDs & "," & rsTmp!病区ID
        rsTmp.MoveNext
    Next
    
    GetUser病区IDs = Mid(GetUser病区IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Have部门性质(ByVal lng科室ID As Long, ByVal str性质 As String, Optional ByRef blnOutDept As Boolean) As Boolean
'功能：检查指定科室是否具有指定工作性质
'说明：因为部门性质一般不变动，又大量使用，利用缓存读取
'返回：blnOutDept=是否为仅服务于门诊的部门
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
    
    blnOutDept = False
     
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    On Error GoTo errH
    If blnNew Then
        strSQL = "Select 部门ID,工作性质,服务对象 From 部门性质说明"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "Have部门性质")
    End If
    rsTmp.Filter = "部门ID=" & lng科室ID & " And 工作性质='" & str性质 & "'"
    Have部门性质 = Not rsTmp.EOF
    If rsTmp.RecordCount > 0 Then
        rsTmp.Filter = "部门ID=" & lng科室ID & " And 工作性质='" & str性质 & "' And 服务对象<>1"
        blnOutDept = rsTmp.RecordCount = 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveBilling(ByVal strNO As String, ByVal blnALL As Boolean, _
     ByVal strTime As String, ByVal lng医嘱ID As Long, ByVal byt来源 As Byte) As Integer
'功能：判断一张记帐单/表是否已经结帐
'参数：strNO=记帐单据号,不分门诊及住院
'      blnALL=是否对整张单据内容进行判断,否则只对未销帐部分进行判断(销帐时)
'      byt来源=1:门诊，2-住院
'返回：0-未结帐,1=已全部结帐,2-已部分结帐
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    Dim strTab As String
    
    On Error GoTo errH
    strTab = IIf(byt来源 = 1, "门诊费用记录", "住院费用记录")
        
    '求未作废的费用行
    strSQL = _
        " Select 序号 From (" & _
        " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号," & _
        " Avg(Nvl(付数, 1) * 数次) As 数量" & _
        " From " & strTab & "" & _
        " Where NO=[1] And 记录性质=2" & _
        " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    
    '求每行的结帐情况
    strSQL = _
        "Select Nvl(价格父号,序号) as 序号,Sum(Nvl(结帐金额,0)) as 结帐金额" & _
        " From " & strTab & "" & _
        " Where NO=[1] And 记录性质 IN(2,12)" & _
        IIf(Not blnALL, " And Nvl(价格父号,序号) IN(" & strSQL & ")", "") & _
        IIf(strTime <> "", " And 登记时间=[2]", "") & _
        IIf(lng医嘱ID <> 0, " And 医嘱序号+0=[3]", "") & _
        " Group by Nvl(价格父号,序号)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "HaveBilling", strNO, CDate(IIf(strTime = "", "1990-01-01", strTime)), lng医嘱ID)
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '单据行数
        rsTmp.Filter = "结帐金额<>0"
        If rsTmp.EOF Then
            HaveBilling = 0 '无结帐行
        ElseIf rsTmp.RecordCount = lngTmp Then
            HaveBilling = 1 '全部行已结帐
        ElseIf rsTmp.RecordCount > 0 Then
            HaveBilling = 2 '部分行已结帐
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckKSSPrivilege(ByVal int场合 As Integer) As Boolean
'功能：检查系统是否存在抗菌药物授权的人员，并且设置当前操作员的用药级别到UserInfo对象
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    UserInfo.用药级别 = 0
    
    On Error GoTo errH
    strSQL = "Select 级别 From 人员抗菌药物权限 Where 记录状态=1 and 人员ID = [1] And 场合=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID, int场合)
    If rsTmp.RecordCount > 0 Then
        UserInfo.用药级别 = Val("" & rsTmp!级别)
        CheckKSSPrivilege = True
    Else
        strSQL = "Select 1 From 人员抗菌药物权限 Where 记录状态=1 and Rownum<2 And 场合=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int场合)
        CheckKSSPrivilege = rsTmp.RecordCount > 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function CheckLISShowVer(lng相关ID As Long) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能说明:          通过相关ID检查老版LIS中是否有记录，如果有记录表示使用老版打开，否则使用新版打开。
    '参数:              lng相关ID       医嘱相关ID
    '返回:              True 老版中找到记录  False 没有找到记录
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    CheckLISShowVer = False
    If lng相关ID = 0 Then
        Exit Function
    End If
    strSQL = "select id from 检验项目分布 where 医嘱id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng相关ID)
    If rsTmp.RecordCount > 0 Then
        CheckLISShowVer = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SetPublicFontSize(ByRef frmMe As Object, ByVal bytSize As Byte, Optional ByVal strOther As String)
'功能：设置窗体及所有控件的字体大小
'参数：frmMe=需要设置字体的窗体对象
'      bytSize:设置为9号字体,0:设置为9号字体,1,设置为12号字体
'      strOther:不进行字体设置的控件父容器的集合,格式为：容器名字1,容器名字2,容器名字3,....
'说明：1.如果涉及到VsFlexGrid等表格控件，需要根据所在的环境重新调整列宽和行高
'      2.如果存在未列出的其他控件或自定义控件,需要用特定方法指定字体大小及相关处理的，需另外单独设置

    Dim objCtrol As Control, objrptCol As ReportColumn
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    Dim strContainer As String
    
    lngFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    frmMe.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In frmMe.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "ReportControl", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox", "IDKind", "PatiIdentify"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '对于CommandBars用户自定义控件读取objCtrol.Container会出错
            strContainer = ""
            On Error Resume Next
            strContainer = objCtrol.Container.Name
            err.Clear: On Error GoTo 0
            If InStr(1, strOther, "," & strContainer & ",") > 0 Then
                 blnDo = False
            End If
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Height = frmMe.TextHeight("字") + 20
                        'Label宽度需要自行调整
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.Count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("字体" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("2012-01-01    ")
                        objCtrol.Height = frmMe.TextHeight("字") + IIf(bytSize = 0, 100, 120)
                Case "TextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("字")
                Case "MaskEdBox"
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = frmMe.TextWidth(objCtrol.Mask)
                        objCtrol.Height = frmMe.TextHeight("字")
                Case "ReportControl"
                        lngOldSize = objCtrol.PaintManager.TextFont.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        Set CtlFont = objCtrol.PaintManager.TextFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.TextFont = CtlFont
                        For Each objrptCol In objCtrol.Columns
                            objrptCol.Width = objrptCol.Width * dblRate
                        Next
                        objCtrol.Redraw
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        lngOldSize = objCtrol.FontSize
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                Case "IDKind"
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "PatiIdentify"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        objCtrol.IDKindFont.Size = lngFontSize
                        objCtrol.objIDKind.Refrash
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
            End Select
        End If
    Next
End Function

Public Sub SetCtrlPosOnLine(ByVal blnvertical As Boolean, ByVal intAligType As Integer, ParamArray arrControls() As Variant)
'功能:对同一行的控件进行位置设置
'参数：
'blnvertical  true ,垂直方向设置控件位置，false,水平方向设置控件位置
'blnvertical=false :intAligType=-1,顶端对齐，0-中间对齐，1-底端对齐,blnvertical=true,intAligType=-1,左对齐，0-水平中心对齐，1-右对齐
'   arrControls格式为控件1,间距1,控件2,间距2,控件3,...
    Dim i As Long
    Dim lngPos As Long '第一个控件的某一位置
    Dim dblRate As Double
    If UBound(arrControls) = -1 Then Exit Sub
    If blnvertical Then
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Left
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Left + 0.5 * arrControls(0).Width
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Left + arrControls(0).Width
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Top = arrControls(i - 2).Top + arrControls(i - 2).Height + arrControls(i - 1)
                arrControls(i).Left = lngPos - arrControls(i).Width * dblRate
            End If
        Next
    Else
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Top
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Top + 0.5 * arrControls(0).Height
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Top + arrControls(0).Height
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Left = arrControls(i - 2).Left + arrControls(i - 2).Width + arrControls(i - 1)
                arrControls(i).Top = lngPos - arrControls(i).Height * dblRate
            End If
        Next
    End If
End Sub

Public Function InitAdviceDefine() As Recordset
'功能：读取医嘱内容定义记录集
'参数：blnNew-是否创建objVBA和objScript对象
'说明：
    Dim strSQL As String
    Dim rsDefine As Recordset
    

    On Error GoTo errH
    strSQL = "Select 诊疗类别,医嘱内容 From 医嘱内容定义 Order by 诊疗类别"
    Set rsDefine = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsDefine, strSQL, "InitAdviceDefine")
    Set InitAdviceDefine = rsDefine
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckSign(ByVal int签名场合 As Long, ByVal lng开嘱科室ID As Long, Optional ByVal lng医技科室ID As Long, Optional ByVal lng病人科室ID As Long, _
    Optional ByVal int病人范围 As Integer = 2, Optional ByVal blnCheckCert As Boolean = True, Optional ByRef objESign As Object) As Boolean
'功能：判断一个部门或是一组部门中是否存在启用了电子签名控制的
'参数：int病人范围=1-门诊,2-住院(缺省)
'     int签名场合:0-门诊医嘱和病历；1-住院医生医嘱和病历；2-住院护士医嘱；3-医技医嘱和报告；4-护理记录和护理病历；5-药品发药；6-LIS;7-PACS;
'     lng开嘱科室ID=如果lng开嘱科室ID=0，则需要根据传入的医技科室，病人科室ID求对应的默认开嘱科室
'                   护士站校对和确认停止时，传入的病区ID，可判断病区是否启用了电子签名
'                   传入-1（抗菌药物审核时，如果判断是否分科室启用）
'     blnCheckCert=true 检查证书是否停用，=false表示不检查
    Dim strSQL As String, intTmp As Integer
    Dim rsTmp As Recordset
    
    '如果场合都未启用，则返回false
    If int签名场合 = 0 Or int签名场合 = 1 Then
        intTmp = int签名场合 + 1
    ElseIf int签名场合 > 1 And int签名场合 <= 7 Then
        intTmp = int签名场合
    End If
    If Mid(gstrESign, intTmp, 1) <> "1" Then Exit Function
    If lng开嘱科室ID = 0 And (lng病人科室ID <> 0 Or lng医技科室ID <> 0) Then
        '取开嘱科室
        lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, lng医技科室ID, lng病人科室ID, int病人范围)
        If lng开嘱科室ID = 0 Then Exit Function
    End If
    grsSign.Filter = "部门ID=" & lng开嘱科室ID & " and 场合=" & int签名场合
    If grsSign.RecordCount = 0 Then
        strSQL = "Select Zl_Fun_Getsignpar([1],[2]) as 是否启用 From dual"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlAdvice", int签名场合, lng开嘱科室ID)
        If rsTmp.RecordCount > 0 Then
            CheckSign = Val(rsTmp!是否启用 & "") = 1
            grsSign.AddNew
            grsSign!部门ID = lng开嘱科室ID
            grsSign!场合 = int签名场合
            grsSign!是否启用 = Val(rsTmp!是否启用 & "")
        End If
    Else
        grsSign.MoveFirst
        CheckSign = Val(grsSign!是否启用 & "") = 1
    End If
    If CheckSign = True And blnCheckCert Then
        If objESign Is Nothing Then
            On Error Resume Next
            Set objESign = CreateObject("zl9ESign.clsESign")
            err.Clear: On Error GoTo 0
            If Not objESign Is Nothing Then
                Call objESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        '检查证书是否停用
        If objESign.CertificateStoped(UserInfo.姓名) Then CheckSign = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get开嘱科室ID(ByVal lng医生ID As Long, ByVal lng医技科室ID As Long, ByVal lng病人科室ID As Long, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal lng执行科室ID As Long, Optional ByVal lng会诊科室ID As Long) As Long
'功能：由医生确定开嘱科室
'参数：int范围=1-门诊,2-住院(缺省)
'说明：在医生所属科室范围内,优先顺序如下：
'      1、医技科室(医技开嘱)
'      2、会诊科室
'      3、病人科室
'      4、服务于门诊/住院病人的某些特殊医嘱的执行科室
'      5、服务于门诊/住院病人的科室且为默认科室
'      6、服务于门诊/住院病人的科室
'      7、默认科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr科室ID(1 To 7) As Long
    
    '开单部门必须是临床或医技
    strSQL = "Select Distinct A.部门ID,Nvl(A.缺省,0) as 缺省" & _
        " From 部门人员 A,部门性质说明 B,部门表 C" & _
        " Where A.部门ID=C.ID And A.部门ID=B.部门ID" & _
        " And B.服务对象 IN([2],3) And A.人员ID=[1]" & _
        " And B.工作性质 IN('临床','检查','检验','手术','治疗','营养')" & _
        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医生ID, int范围)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!部门ID = lng医技科室ID Then
            arr科室ID(1) = rsTmp!部门ID
        ElseIf rsTmp!部门ID = lng会诊科室ID Then
            arr科室ID(2) = rsTmp!部门ID
        ElseIf rsTmp!部门ID = lng病人科室ID Then
            arr科室ID(3) = rsTmp!部门ID
        ElseIf rsTmp!部门ID = lng执行科室ID Then
            arr科室ID(4) = rsTmp!部门ID
        ElseIf rsTmp!缺省 = 1 Then
            arr科室ID(5) = rsTmp!部门ID
        ElseIf arr科室ID(5) = 0 Then
            arr科室ID(6) = rsTmp!部门ID
        End If
        rsTmp.MoveNext
    Next
    arr科室ID(7) = UserInfo.部门ID
    
    For i = LBound(arr科室ID) To UBound(arr科室ID)
        If arr科室ID(i) <> 0 Then
            Get开嘱科室ID = arr科室ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function CreatePlugInOK(ByVal lngMod As Long, Optional ByVal int场合 As Integer) As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod, int场合)
        Call zlPlugInErrH(err, "Initialize")
        err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Function Get输液配置中心() As String
'功能：获取输液配置中心的科室IDs
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strReturn As String
    
    On Error GoTo errH

    strSQL = "Select 部门id From 部门性质说明 Where 工作性质 = '配制中心' Order by 部门id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get输液配置中心")
    
    For i = 1 To rsTmp.RecordCount
        strReturn = strReturn & "," & rsTmp!部门ID
        rsTmp.MoveNext
    Next
    Get输液配置中心 = Mid(strReturn, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HavePath(ByVal lng部门ID As Long) As Boolean
'功能：检查指定科室或病区是否有可用的临床路径
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "Select a.Id" & vbNewLine & _
            "From 临床路径目录 A, 临床路径版本 B, 临床路径科室 C," & vbNewLine & _
            "     (Select 科室id From 病区科室对应 Where 病区id = [1]" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select ID From 部门表 Where ID = [1]) D" & vbNewLine & _
            "Where a.Id = b.路径id And a.最新版本 = b.版本号 And a.Id = c.路径id(+) And (c.科室id = d.科室id or c.科室id is null) And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng部门ID)
    HavePath = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病种ID(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, _
                Optional ByRef bln中医 As Boolean = False, Optional ByVal bytType As Byte) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSQL As String
'bytType =0 取首要诊断,=1首要诊断找不到时取次要诊断

    bln中医 = Have部门性质(lng科室ID, "中医科")
    If bytType = 0 Then
        If bln中医 Then
            strSQL = "Select 疾病id, 诊断id, 诊断描述,诊断类型 " & vbNewLine & _
                "From 病人诊断记录" & vbNewLine & _
                "Where 记录来源 In (1, 2, 3) And 诊断类型 In (1, 2, 11, 12) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 = 1 And" & vbNewLine & _
                "      Nvl(是否疑诊, 0) = 0 And NVL(编码序号,1) = 1 And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
                "Order By Decode(诊断类型, 12, 1, 2, 2, 11, 3, 1, 4), Decode(记录来源, 1, 4, 记录来源) Desc"
        Else
            strSQL = "Select 疾病id, 诊断id, 诊断描述,诊断类型 " & vbNewLine & _
            "From 病人诊断记录" & vbNewLine & _
            "Where 记录来源 In (1, 2, 3) And 诊断类型 In (1, 2, 11, 12) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 = 1 And Nvl(是否疑诊,0) = 0 And NVL(编码序号,1) = 1 And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
            "Order By Sign(诊断类型-10),诊断类型 Desc, Decode(记录来源, 1, 4, 记录来源) Desc"
        End If
    ElseIf bytType = 2 Then
        strSQL = "Select a.疾病id, a.诊断id, a.诊断描述, a.诊断类型, a.记录来源" & vbNewLine & _
            "From 病人诊断记录 A" & vbNewLine & _
            "Where a.记录来源 In (1, 2, 3) And a.取消时间 Is Null And a.病人id = [1] And a.主页id = [2] And" & vbNewLine & _
            "      (a.诊断类型 In (1, 2, 3, 11, 12,13) And a.诊断次序 <> 1 Or" & vbNewLine & _
            "      a.诊断类型 = 10) And Nvl(a.是否疑诊, 0) = 0 And NVL(编码序号,1) = 1 And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
            "Order By Sign(a.诊断类型 - 10), a.诊断类型 Desc, Decode(a.记录来源, 1, 4, a.记录来源) Desc"
    Else
        If bln中医 Then
            strSQL = "Select Distinct a.Id, k.疾病id, k.诊断id, k.诊断描述, K.诊断类型, K.记录来源,k.排序 " & vbNewLine & _
                "From 临床路径目录 A, 临床路径病种 B, 临床路径版本 C," & vbNewLine & _
                "     (Select Rownum As 排序, 疾病id, 诊断id, 诊断描述, 诊断类型, 记录来源 " & vbNewLine & _
                "       From 病人诊断记录" & vbNewLine & _
                "       Where 记录来源 In (1, 2, 3) And 诊断类型 In (1, 2, 11, 12) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 <> 1 And" & vbNewLine & _
                "             Nvl(是否疑诊, 0) = 0 And NVL(编码序号,1) = 1 And Not (Nvl(疾病id, 0) = 0 And Nvl(诊断id, 0) = 0)" & vbNewLine & _
                "       Order By Decode(诊断类型, 12, 1, 2, 2, 11, 3, 1, 4), Decode(记录来源, 1, 4, 记录来源) Desc, 诊断次序) K" & vbNewLine & _
                "Where a.Id = b.路径id And a.Id = b.路径id And a.Id = c.路径id And a.最新版本 = c.版本号 And a.性质 = 0 And b.性质 = 0 And" & vbNewLine & _
                "      (b.疾病id = k.疾病id Or b.诊断id = k.诊断id) And" & vbNewLine & _
                "      (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 临床路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & vbNewLine & _
                "Order By k.排序"
        Else
            strSQL = "Select Distinct a.Id, k.疾病id, k.诊断id, k.诊断描述,K.诊断类型, K.记录来源,k.排序 " & vbNewLine & _
            "From 临床路径目录 A, 临床路径病种 B, 临床路径版本 C," & vbNewLine & _
            "     (Select Rownum As 排序, 疾病id, 诊断id, 诊断描述, 诊断类型, 记录来源 " & vbNewLine & _
            "       From 病人诊断记录" & vbNewLine & _
            "       Where 记录来源 In (1, 2, 3) And 诊断类型 In (1, 2, 11, 12) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 <> 1 And" & vbNewLine & _
            "             Nvl(是否疑诊, 0) = 0 And NVL(编码序号,1) = 1 And Not (Nvl(疾病id, 0) = 0 And Nvl(诊断id, 0) = 0)" & vbNewLine & _
            "       Order By Sign(诊断类型 - 10), 诊断类型 Desc, Decode(记录来源, 1, 4, 记录来源) Desc, 诊断次序) K" & vbNewLine & _
            "Where a.Id = b.路径id And a.Id = b.路径id And a.Id = c.路径id And a.最新版本 = c.版本号 And a.性质 = 0 And b.性质 = 0 And" & vbNewLine & _
            "      (b.疾病id = k.疾病id Or b.诊断id = k.诊断id) And" & vbNewLine & _
            "      (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 临床路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & vbNewLine & _
            "Order By k.排序"
        End If
    End If
    '记录来源:1-病历；2-入院登记；3-首页整理;4-病案
    '诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
    '有多个诊断的情况下，根据诊断次序，只取第一个主要诊断
    '病历里面的诊断优先，主要是为了支持修正诊断。
    '剔除自由录入诊断
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取病种", lng病人ID, lng主页ID, lng科室ID)
    Set Get病种ID = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPathTable(lng疾病ID As Long, lng诊断ID As Long, lng科室ID As Long) As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select a.Id, a.分类, a.编码, a.名称, a.说明, Nvl(a.适用病情,'通用') 适用病情, a.适用性别, a.适用年龄, a.最新版本, c.标准住院日,Nvl(a.病例分型,'无') as 病例分型,Nvl(a.确诊天数,0) as 确诊天数" & vbNewLine & _
            "From 临床路径目录 A, 临床路径病种 B,临床路径版本 C" & vbNewLine & _
            "Where a.Id = b.路径id And (b.疾病id = [1] Or b.诊断id = [2]) And a.最新版本 is not null And a.id = b.路径ID And a.最新版本 = c.版本号" & vbNewLine & _
            "And a.Id = c.路径id And b.性质=0 And (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 临床路径科室 D Where a.Id = D.路径id And d.科室id = [3]))"
    On Error GoTo errH
    Set GetPathTable = zlDatabase.OpenSQLRecord(strSQL, "读取路径目录", lng疾病ID, lng诊断ID, lng科室ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CreateObjectPacs(objPublicPACS As Object) As Boolean
    If objPublicPACS Is Nothing Then
        On Error Resume Next
        Set objPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        err.Clear: On Error GoTo 0
        If Not objPublicPACS Is Nothing Then
            Call objPublicPACS.InitInterface(gcnOracle, UserInfo.用户名)
        End If
        If objPublicPACS Is Nothing Then
            MsgBox "PACS公共部件未创建成功！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreateObjectPacs = True
End Function

Public Function HaveRIS() As Boolean
'功能：判断 RIS接口部件 是否存在，并启用
    gbln启用影像信息系统预约 = False
    If Not gbln启用影像信息系统接口 Then Exit Function
    If gobjRis Is Nothing Then
        On Error Resume Next
        Set gobjRis = CreateObject("zl9XWInterface.clsHISInner")
        err.Clear: On Error GoTo 0
    End If
    If gobjRis Is Nothing Then
        MsgBox "RIS接口创建失败，不能继续当前操作。可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Function
    End If
    gbln启用影像信息系统预约 = gobjRis.HISSchedulingjudge = 0
    HaveRIS = True
End Function


Public Function InitObjBlood(Optional ByVal blnMsg As Boolean = False) As Boolean
'判断如果血库部件为空就初始化
    If gobjPublicBlood Is Nothing Then
        On Error Resume Next
        Set gobjPublicBlood = CreateObject("zlPublicBlood.clsPublicBlood")
        If Not gobjPublicBlood Is Nothing Then
            If gobjPublicBlood.zlInitCommon(gcnOracle, gstrDBUser) = False Then
                Set gobjPublicBlood = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    If gobjPublicBlood Is Nothing Then
        If blnMsg = True Then
            MsgBox "血库公共部件zlPublicBlood创建失败，不能继续当前操作。可能是部件未安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    InitObjBlood = Not gobjPublicBlood Is Nothing
End Function
