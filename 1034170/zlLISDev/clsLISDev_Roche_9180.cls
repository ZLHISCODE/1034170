VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Roche_9180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Roche 9180
'厂商:罗氏
'作者:
'用户:塘厦医院
Private Const LABITEMS As String = "Na,K,Cl"
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
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    Dim iPos As Long, iEndPos As Long, i As Integer, aItems() As String
    Dim strDate As String, lngSampleNO As Long, strItemName As String, dblItemValue As Double
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1: blnReadItem = False
    strDate = "": lngSampleNO = -1
    Do While iStrPoint <= Len(strSource)
        iStrPoint = InStr(iStrPoint, strSource, "")
        '判断一组结果是否开始
        If iStrPoint = 0 Then Exit Do
        '记录一组结果的入口
        iCurrGrpEntry = iStrPoint
        
        iEndPos = InStr(iStrPoint, strSource, "")
        If iEndPos = 0 Then Exit Do
        
        '检验时间
        '时间格式必须为：yyyy-mm-dd hh:mm:ss
        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
        
        '样本序号，必须为数字
        iStrPoint = InStr(iStrPoint, strSource, "Sample No.") + 10
        strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 5))
        
        '检验人
        strData = strData & "| " '本仪器不支持
        '标本
        strData = strData & "|血液" '本仪器不支持
        '质控品
        strData = strData & "|0"
        
        aItems = Split(LABITEMS, ",")
        For i = 0 To UBound(aItems)
            iPos = InStr(iStrPoint, strSource, Mid(aItems(i) & "  ", 1, 2) & "=")
            If iPos > 0 Then
                strData = strData & "|" & aItems(i)
                strData = strData & "|" & Val(Mid(strSource, iPos + 3, 7))
            End If
        Next
    
        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
        
        '该组结果已解析
        iCurrGrpEntry = 0
        iStrPoint = iEndPos + 1
    Loop
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    '向设备发送‘数据收到’命令
    strCmd = Chr(&H6&)
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

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    '指针定位到下一个换行之后
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '读取子串，同时指针移动到读取的字符串之后
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '读取一行字符，并将字符指针移到下一行开始位置
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(13) + Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 2
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '读取一个检验项目名称，并将字符指针移到项目名称之后
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

