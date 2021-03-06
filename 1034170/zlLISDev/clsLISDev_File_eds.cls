VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_eds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:细菌分析仪
'厂商:北京金山
'作者:陈东
'用户:晋煤集团总医院

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'标准函数：从文件获取检验结果
'   strFile：包含路径的文件名
'   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
'   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
'   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
'   返回的检验结果数组(各解析程序必须按以下标准组织结果)
'   元素之间以|分隔
'   第0个元素：检验时间
'   第1个元素：样本序号
'   第2个元素：检验人
'   第3个元素：标本
'   第4个元素：是否质控品
'   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
'       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    
    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strBh As String

 
  
    ReDim aResults(0)
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    If Dir(strFile) = "" Then
        ResultFromFile = aResults
        Exit Function
    End If
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then
            ResultFromFile = aResults
            Exit Function
        End If
        strSQL = "select * From " & _
                "(Select Val(样品号) as 标本号,检验结果,Format(Mid(检验日期,1,2) & '-' & Mid(检验日期,3,2) & '-' & Mid(检验日期,5,2) ,'yyyy-MM-dd') as 日期" & _
                " from eds)" & vbNewLine & _
                " Where 标本号<>0 And 日期>='" & Format(dtStart, "yyyy-MM-dd") & "' And 日期<='" & Format(dtEnd, "yyyy-MM-dd") & "' " & _
                IIf(lngSampleNo = -1, "", " And val(标本号)=" & lngSampleNo)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            ResultFromFile = aResults
            Exit Function
        End If
 
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("标本号") Then '新标本
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = Format(rsTmp("日期"), "yyyy-mm-dd")  '检验时间
                
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsTmp("标本号"))   '样本序号
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "  '检验人
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|血液"  '标本
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|001|" & rsTmp!检验结果
                strBh = rsTmp("标本号")
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function


