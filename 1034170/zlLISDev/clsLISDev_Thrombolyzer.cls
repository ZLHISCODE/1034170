VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Thrombolyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血凝仪
'型号:Compact-X
'厂商:Thrombolyzer
'作者:曾超

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
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim iStrPoint As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim varIn As Variant '存放输入数据
    Dim StrTime As String
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strLine As String
    Dim str标本号 As String, str日期 As String, str检验人标本质控 As String, str指标 As String
    Dim aLine() As String
    Dim strDate As String
    Dim intloop As Integer
    Dim aItem() As String
    Dim strNO As String
    Dim strNow As String
    
    
    On Error GoTo errHandle
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    '没有开始和结束字符时退出
    If InStr(strSource, Chr(&H2)) = 0 Or InStr(strSource, Chr(&H3)) = 0 Then
        strReserved = strSource
        strCmd = Chr(&H6)
        Exit Sub
    End If
    
    strNow = Mid(strSource, InStr(strSource, Chr(&H2)), InStr(strSource, Chr(&H3)))
    strReserved = Mid(strSource, InStr(strSource, Chr(&H3)) + 1)
    
    '先取标本号和日期信息
    aLine = Split(strNow, vbCrLf)
    
    For intloop = 0 To UBound(aLine)
        aItem = Split(aLine(intloop), "|")
        If UBound(aItem) > 3 Then
            strData = strData & "||" & Now & "|" & Val(aItem(1)) & "| |血清|0|" & Trim(aItem(2)) & "1" & "|" & Trim(aItem(5))
            strData = strData & "||" & Now & "|" & Val(aItem(1)) & "| |血清|0|" & Trim(aItem(2)) & "2" & "|" & Trim(aItem(6))
        End If
    Next
    
    
    strResult = Mid(strData, 3)
    strCmd = Chr(&H6)
    Exit Sub
errHandle:
    WriteLog "Thrombolyzer", Err.Description, ""

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










