VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HT150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:HT-150
'厂商:桂林华通
'作者:陈东
'用户:商丘市第一人民医院

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
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    '请在这个过程中写解析程序，上面的注释行请不要删除。
    Dim str_时间 As String, str_样本号 As String, str_项目 As String, str_结果 As String
    Dim lon_str As Long, lon_end As Long, str_Tmp As String
    Dim str_Data As String
    Dim varline As Variant
    Dim i As Integer, str_Line As String
    
    lon_str = InStr(strSource, Chr(&H2))
    lon_end = InStr(strSource, Chr(&H3))
    strResult = ""
    If lon_str > 0 And lon_end > 0 Then
        str_Data = Mid(strSource, lon_str, lon_end)
        strReserved = Mid(strSource, lon_end + 1)
        str_项目 = ""
        str_时间 = ""
        varline = Split(str_Data, Chr(&HD) & Chr(&HA))
        For i = LBound(varline) To UBound(varline)
            str_Line = varline(i)
            If UCase(str_Line) Like "*NO*" Then
                str_Line = UCase(str_Line)
                str_样本号 = Val(Mid(str_Line, InStr(str_Line, "NO") + 2))

            ElseIf UCase(str_Line) Like "*DATE*" Then
                str_时间 = Trim(Mid(str_Line, InStr(UCase(str_Line), "DATE") + 4))
                If IsDate(str_时间) Then
                    str_时间 = Format(CDate(str_时间), "yyyy-MM-dd")
                End If
            ElseIf Not (UCase(str_Line) Like "*TIME*") Then
                    str_Tmp = Trim(Mid(str_Line, 5))
                    str_Tmp = Replace(str_Tmp, "mmol/L", "")
                    str_Tmp = Replace(str_Tmp, "Cell/uL", "")
                    str_Tmp = Replace(str_Tmp, "umol/L", "")
                    str_Tmp = Replace(str_Tmp, "mg/L", "")
                    str_Tmp = Replace(str_Tmp, "Leu/uL", "")
                    str_Tmp = Replace(str_Tmp, "g/L", "")
                    If str_Tmp <> "" And Trim(Mid(str_Line, 2, 4)) <> "" Then
                        str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 2, 4)) & "|" & Trim(str_Tmp)
                    End If
            End If
            
        Next
        If Val(str_样本号) > 0 And IsDate(str_时间) And str_项目 <> "" Then
            strResult = str_时间 & "|" & str_样本号 & "|" & " |" & "尿液" & "|0" & str_项目
        End If
    Else
        strReserved = strSource
    End If

End Sub


Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function






