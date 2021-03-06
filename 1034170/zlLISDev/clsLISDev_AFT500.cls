VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AFT500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AFT-500 电解质分析仪
'厂商:梅州康立
'作者:陈东
'用户:甘肃宋巨强

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
    Dim strEndChar As String
    Dim lng_Loop As Long, var_Code As Variant
    strResult = ""
    str_in = strSource
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    strEndChar = "tCa"
    If InStr(mstrCode, "PID:") > 0 Then strEndChar = "iCa"
    
    If mstrCode Like "*####-##-## ##:##*" & strEndChar & "*" & Chr(&HA) & "*" Then                          '收到一个完整的记录,开始解析
        WriteLog "解析 AFT500", mstrCode, ""
        Do While InStr(mstrCode, strEndChar) > 0
            lng_S = InStr(mstrCode, "-")
            If lng_S <= 4 Then Exit Do
            lng_S = lng_S - 4
            lng_E = InStr(lng_S, mstrCode, strEndChar)
            lng_E = InStr(lng_E, mstrCode, Chr(&HA))
            
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            var_Code = Split(str_Code, Chr(&HA))
            
            For lng_Loop = LBound(var_Code) To UBound(var_Code)

                str_Line = var_Code(lng_Loop)
                
                If str_Line Like "*####-##-## ##:##*" Then                           '标本号,项目 行
                    str_日期 = str_Line
                    str_标本号 = "": str_项目 = ""
                ElseIf InStr(str_Line, "样品编号:") > 0 Then                       '项目 结果 行
                    str_标本号 = Replace(str_Line, "样品编号:", "")
                ElseIf InStr(str_Line, "PID:") > 0 Then                       '项目 结果 行
                    str_标本号 = Replace(str_Line, "PID:", "")
                ElseIf InStr(str_Line, "=") > 0 Then
                    str_结果 = Trim(Replace(Split(str_Line, "=")(1), "mmol/L", ""))
                    If str_结果 <> "" And Trim(Split(str_Line, "=")(0)) <> "" Then
                        str_项目 = str_项目 & "|" & Trim(Split(str_Line, "=")(0)) & "|" & str_结果
                    End If
                End If

            Next
            If Val(str_标本号) <> 0 And IsDate(str_日期) And str_项目 <> "" Then
                strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_项目
            End If
        Loop
    End If
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "解析 AFT500", strResult, ""
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










