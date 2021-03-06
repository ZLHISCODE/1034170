VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_RT_2100C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:雷杜酶标仪
'型号:RT-2100C
'厂商:Rayto(雷杜)
'作者:李建飞
'用户:
'日期:2009-08-21

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
                        ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'标准函数：从文件获取检验结果
    'strFile：包含路径的文件名
    'lngSampleNo：标本号。为-1表示选取所有时间范围内的标本
    'dtStart：开始时间，如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
    'dtEnd：结束时间。只有选取多个标本(lngSample=-1)时该参数有效。如果不指定该参数，则时间范围为>=dtStart
    '返回的检验结果数组（各解析程序必须按照以下标准组织结果）
    '元素之间以 | 分隔
    '第 0 个元素：检验时间
    '第 1 个元素：样本序号
    '第 2 个元素：检验人
    '第 3 个元素：标本
    '第 4 个元素：是否质控品
    '从第 5 个元素开始为检验结果，每两个元素表示一个检验项目
    '如：第 5i 个元素为检验项目，第 5i+1 个元素表示检验结果
 
    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim cnErs As New ADODB.Connection, strFileErs As String, rsErs As ADODB.Recordset, strImgErs As String
    Dim str_检验时间 As String, str_检验项目 As String, str_样本序号 As String, str_检验人 As String, str_标本 As String
    
        
    strFilePath = objFile.GetParentFolderName(strFile)  '取路径
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFileErs = strFile
    
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
            .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};"
            .Open
            If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
    
    ReDim aResults(0)
    strImgErs = ""
    
    If strFileErs <> "" Then
        strSql = "select * from Table_Pa_Item " & _
                 "as a where format(Item_test_date,'yyyy-mm-dd')>= '" & Format(dtStart, "yyyy-mm-dd") & "' " & _
                 "and format(Item_test_date,'yyyy-mm-dd')<= '" & Format(dtEnd, "YYYY-MM-DD") & "'" & _
                 IIf(lngSampleNo = -1, "", " and VAL(Sample_id)=" & lngSampleNo)
        Set rsErs = cnErs.Execute(strSql)
        
        While Not rsErs.EOF
            '根据文件结构读取检验信息
            If Val("" & rsErs!Sample_id) <> 0 Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                '第 0 个元素：检验时间
                '第 1 个元素：样本序号
                '第 2 个元素：检验人
                '第 3 个元素：标本
                '第 4 个元素：是否质控品
                '第 5 个元素开始为检验结果，每两个元素表示一个检验项目
                '如：第 5i 个元素为检验项目，第 5i+1 个元素表示检验结果
                str_检验时间 = Format(Trim(rsErs!Item_test_date), "yyyy-mm-dd")                    '检验时间
                str_样本序号 = Trim(rsErs!Sample_id)
                str_检验人 = " "
                str_标本 = "血清|0"
                
                '检验项目
                str_检验项目 = rsErs!Item_ename & "|" & rsErs!Item_result_det & "^" & rsErs!Item_result_od & _
                                "^" & rsErs!Item_cutoff & "^" & rsErs!Item_result
                aResults(UBound(aResults)) = str_检验时间 & "|" & str_样本序号 & "|" & str_检验人 & "|" & str_标本 & "|" & str_检验项目
            End If

            rsErs.MoveNext
        Wend
    End If
    ResultFromFile = aResults

End Function

