VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_IMMU1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: IMMULITE 1000 全自动发光免疫分析仪
'厂商: DPC
'作者: 陈东
'用户: 汉中3201医院
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
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
    Dim strID As String
    Dim rsErs As ADODB.Recordset
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd=www.dpc.com"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = "SELECT Rpt_no as 标本号,Format(Rpt_date,'yyyy-MM-dd') as 日期,Rpt_ID  " & _
                 " From TBL_Rpt As A " & _
                 "Where Format(A.RPT_Date,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'" & _
                 " And Format(A.RPT_Date,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                 IIf(lngSampleNO = -1, "", " And val(Rpt_no)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        strID = "***"
        Do While Not rsTmp.EOF
        
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                
            strID = "" & rsTmp("Rpt_ID")
            
            strSQL = "Select ITEM_CODE,ITEM_RESULT From TBL_Rpt_Item As B Where B.RPT_ID='" & strID & "'"
            Set rsErs = cnAccess.Execute(strSQL)
            Do Until rsErs.EOF
                If aResults(UBound(aResults)) = "" Then
                    aResults(UBound(aResults)) = Trim("" & rsTmp("日期"))
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsTmp("标本号"))   '样本序号
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '检验人
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血清"                      '标本
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                              '质控
                End If
                If Trim("" & rsErs!ITEM_RESULT) <> "" And Trim("" & rsErs!ITEM_CODE) <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Trim("" & rsErs!ITEM_CODE) & _
                         "|" & Trim("" & rsErs!ITEM_RESULT)  '结果
                End If
                rsErs.MoveNext
            Loop
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function




