VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_URISYS_2400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:URISYS-2400 (尿液分析仪)
'厂商:罗氏
'作者:陈东
'用户:晋煤集团医院

Private mstrCode As String

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
        
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号^是否急诊^条码号  可以只有样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
   
    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim str_标本号 As String, str_项目 As String, str_结果 As String, str_日期 As String
    Dim str_条码 As String, str_SendLine As String
    
    
    strResult = ""
    str_in = strSource
    '-------------------------------------------------------------------------------------------------------------
    If InStr(str_in, Chr(&H5)) > 0 Then                                         '收到开始符 05 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(str_in, Chr(&H3)) > 0 Or InStr(str_in, Chr(&H17)) > 0 Then         '收到行结束符 03,17 ,应答 06
        strCmd = Chr(&H6)
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    If mstrCode Like "*1H|*" & Chr(&H4) & "*" Then                         '收到一个完整的记录,开始解析
        WriteLog "解析 URISYS－2400", mstrCode, ""
        Do While InStr(mstrCode, Chr(&H4)) > 0
            lng_S = InStr(mstrCode, "1H|")
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
            
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            
            '处理 校验码
            Do While InStr(str_Code, Chr(&H17)) > 0
                lng_S = InStr(str_Code, Chr(&H17))
                str_Code = Mid(str_Code, 1, lng_S - 1) & Mid(str_Code, lng_S + 7)
            Loop
            Do While InStr(str_Code, Chr(&H3)) > 0
                lng_S = InStr(str_Code, Chr(&H3))
                str_Code = Mid(str_Code, 1, lng_S - 1) & Mid(str_Code, lng_S + 5)
            Loop
            
            Do While str_Code <> ""
                lng_S = 1
                lng_E = InStr(str_Code, Chr(&HD))
                
                If lng_S > 0 And lng_E > 0 Then
                    str_Line = Mid(str_Code, lng_S, lng_E - lng_S + 1)
                    str_Code = Mid(str_Code, lng_E + 1)
                    
                    If str_Line Like "O|*|*" Then                            '标本号,项目 行
                        str_标本号 = Val(Split(str_Line, "|")(3))
                        str_日期 = Split(str_Line, "|")(14)
                        If str_日期 <> "" Then
                            str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2)
                        End If
                    ElseIf InStr(str_Line, "1H|") > 0 Then                       '项目 结果 行
                        str_标本号 = "": str_项目 = ""
                    ElseIf str_Line Like "R|*|*" Then
                        
                        str_结果 = Trim(Split(str_Line, "|")(3))
                        If str_结果 <> "" Then
                            str_项目 = str_项目 & "|" & Replace(Split(str_Line, "|")(2), "^^^", "") & "|" & str_结果
                        End If
                    End If
                Else
                    str_Code = ""
                End If
            Loop
            If Val(str_标本号) <> 0 And IsDate(str_日期) And str_项目 <> "" Then
                strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |尿液|0" & str_项目
            End If
        Loop
    End If
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "解析 URISYS－2400 End", strResult, ""
    End If
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
        SendSample = ""
End Function






