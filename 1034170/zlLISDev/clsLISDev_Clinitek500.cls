VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Clinitek500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Clinitek 500 尿液分析仪
'厂商:德国拜尔(Bayer)
'作者:陈东
'用户:包头中心医院
Dim mstrCode As String

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
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant
    Dim str_Line As String, int_Loop As Integer
    Dim str_日期 As String, str_标本号 As String, str_项目 As String
    mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(mstrCode, Chr(&H3))
    
    
    If lng_S > 0 And lng_E > 0 Then
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        var_Code = Split(str_Code, vbCrLf)
        For int_Loop = 0 To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If Left(str_Line, 1) = "#" Then
                str_日期 = "20" & Right(str_Line, 8)
                str_标本号 = Trim(Mid(str_Line, InStr(str_Line, "-") + 1, Len(str_Line) - InStr(str_Line, "-") - 8))
            ElseIf Len(str_Line) = 20 Then
                str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, 3)) & "|" & Trim(Mid(str_Line, 4))
            End If
        Next
        
        If Val(str_标本号) <> 0 And IsDate(str_日期) And str_项目 <> "" Then
            str_项目 = Replace(str_项目, "umol/L", "")
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |尿液|0" & str_项目
        End If
        
        If strResult <> "" Then strResult = Mid(strResult, 3)
    End If
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


