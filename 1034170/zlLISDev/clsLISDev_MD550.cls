VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MD550"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:MD-550
'厂商:三和医疗设备有限公司
'作者:陈东
'用户:曲阳医院

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = Chr(6)
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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, lngLength As Long, strRecord As String
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "*[" & Chr(&H23&) & "]" & Chr(&H24&) & "*" & Chr(&H24&) & Chr(&H26&) & "*" Or _
       strSource Like "*" & Chr(&H27&) & Chr(&H28&) & "*" & Chr(&H28&) & Chr(&H2A&) & "*" Then
       
        strSource = Replace(strSource, Chr(&H23&) & Chr(&H24&), Chr(&H27&) & Chr(&H28&))
        strSource = Replace(strSource, Chr(&H24&) & Chr(&H26&), Chr(&H28&) & Chr(&H2A&))
        
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            lngLength = InStr(strReserved, Chr(&H2A&)) - InStr(strReserved, Chr(&H27&))
            strLine = Mid(strReserved, InStr(strReserved, Chr(&H27&)), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H2A&)) + 1)
            
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            '解析
            strItem = ""
            strLine = Replace(strLine, Chr(&H27&) & Chr(&H28&), "")
            strLine = Replace(strLine, Chr(&H28&) & Chr(&H2A&), "")
            Do While Len(strLine) >= 64
                strRecord = Mid(strLine, 1, 64)
                strLine = Mid(strLine, 65)
                str标本号 = Val(Mid(strRecord, 1, 3)) & IIf(Mid(strRecord, 5, 1) = "1", "^1", "")
                str检验人标本质控 = " |" & Switch(Mid(strRecord, 4, 1) = "0", "血液", Mid(strRecord, 4, 1) = "1", "尿液", Mid(strRecord, 4, 1) = "2", "其他") & "|0"
                str日期 = Mid(strRecord, 7, 4) & "-" & Mid(strRecord, 11, 2) & "-" & Mid(strRecord, 13, 2)
                
                strItem = IIf(InStr(Mid(strRecord, 22, 6), ".") > 0, "|01|" & Trim(Mid(strRecord, 22, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 28, 6), ".") > 0, "|02|" & Trim(Mid(strRecord, 28, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 34, 6), ".") > 0, "|03|" & Trim(Mid(strRecord, 34, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 40, 6), ".") > 0, "|04|" & Trim(Mid(strRecord, 40, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 46, 6), ".") > 0, "|05|" & Trim(Mid(strRecord, 46, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 52, 6), ".") > 0, "|06|" & Trim(Mid(strRecord, 52, 6)), "") & _
                          IIf(InStr(Mid(strRecord, 58, 6), ".") > 0, "|07|" & Trim(Mid(strRecord, 58, 6)), "")
                If strItem <> "" And Val(str标本号) <> 0 Then
                    strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
                End If
            Loop
            If Not strReserved Like "*" & Chr(&H27&) & Chr(&H28&) & "*" & Chr(&H28&) & Chr(&H2A&) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        '收到的信息不全,返回,等下次再解析
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    'strCmd = Chr(6)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '           元素之间以|分隔
    '           第0个元素：检验时间
    '           第1个元素：样本序号
    '           第2个元素：检验人
    '           第3个元素：标本
    '           第4个元素：是否质控品
    '           第5～9元素：系统保留
    '           从第10个元素开始为需要的检验项目。
    '   iSendStep  : 步骤
    '   blnSuccess : 是否传输失败
    '   strResponse: 仪器的应答信息
    '   blnUndo    : 是否取消.
    '
    '   返回值：发到仪器的字串或指令。
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function

