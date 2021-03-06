VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LG_PABER_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血小板聚集凝血因子分析仪
'型号:LG-PABER-I
'厂商:
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
    Dim intLoop As Integer
    Dim aItem() As String
    Dim strNO As String
    Dim strNow As String
    Dim strItem As String
    Dim intCount As Integer
    Const StrAllitem As String = ",ADP,PT,TT,APTT,FIB,"
    
    On Error GoTo errHandle
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    '没有开始和结束字符时退出
    If InStr(strSource, "@c") = 0 Or InStr(strSource, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD)) = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    strNow = Mid(strSource, InStr(strSource, "@c") + 5, InStr(strSource, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD)) + 5)
    
    If InStr(strNow, "Meter") > 0 Then
        strNow = Mid(strNow, InStr(strNow, "Meter") + 6)
    End If
    
    aLine = Split(strNow, vbCr)
    
    strDate = Replace(Mid(aLine(0), 1, 10), ".", "-") & " " & Time
    
    For intLoop = 3 To UBound(aLine)
        strLine = aLine(intLoop)
        strLine = Replace(strLine, Chr(30), " ")
        strLine = Replace(strLine, Chr(0), " ")
        strLine = Replace(strLine, Chr(9), " ")
        
        strLine = DelSapce(strLine)
        aItem = Split(strLine, " ")
        
        If InStr(strLine, "") > 0 Or InStr(strLine, Chr(&H14) & Chr(&H1E)) > 0 Then
            strNO = aItem(1)
        End If
       
        If UBound(aItem) > 1 Then
        
            If InStr(strLine, " ADP ") Then
                strItem = "ADP"
            End If
            
            If InStr(strLine, " PT ") Then
                strItem = "PT"
                strData = strData & "||" & strDate & "|" & strNO & "| |血清|0|PT|" & aItem(UBound(aItem) - 2)
            End If
            
            If InStr(strLine, " TT ") Then
                strData = strData & "||" & strDate & "|" & strNO & "| |血清|0|TT|" & aItem(UBound(aItem) - 1)
                strItem = ""
            End If
            
            If InStr(strLine, " APTT ") Then
                strData = strData & "||" & strDate & "|" & strNO & "| |血清|0|APTT|" & aItem(UBound(aItem) - 1)
                strItem = ""
            End If
            
            If InStr(strLine, " FIB ") Then
                strItem = ""
                strData = strData & "||" & strDate & "|" & strNO & "| |血清|0|FIB|" & aItem(UBound(aItem) - 1)
            End If
            
            If strItem = "ADP" Then
                If aItem(1) = "180" Then
                    strData = strData & "||" & strDate & "|" & strNO & "| |血清|0|ADP|" & aItem(UBound(aItem) - 1)
                End If
            End If
            
            If strItem = "PT" Then
                If aItem(2) = "INR" Then
                    strData = strData & "||" & strDate & "|" & strNO & "| |血清|0|INR|" & aItem(UBound(aItem) - 1)
                End If
            End If
        End If
        
    Next
    
    strResult = Mid(strData, 3)

    Exit Sub
    

errHandle:
    WriteLog "lg_Paber_i,解码错误", Err.Description, ""

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












