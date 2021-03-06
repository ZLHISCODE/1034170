VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Immu1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:IMMULITE 1000 全自动发光免疫分析仪
'厂商: DPC
'作者: 陈东
'用户: 汉中3201医院
Private mstrCode As String      '待解析的原始数据
Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    'GetStartCmd = Chr(6)
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
    Dim lng_S As Long, lng_E As Long, str_Code As String, str_Line As String, var_Code As Variant
    Dim int_Loop As Integer, str_结果 As String, str_标本号 As String, str_项目 As String, str_日期 As String
    
    '-- 处理 应答
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    'If InStr(strSource, Chr(&H4)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)

    If Len(strSource) <> 0 Then mstrCode = mstrCode & strSource
    
    '--解析
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If Left(str_Line, 1) = Chr(&H2) And Mid(str_Line, 3, 2) = "O|" Then
                If str_标本号 <> Split(str_Line, "|")(2) Then
                
                    If str_标本号 <> "" And str_结果 <> "" And str_日期 <> "" Then
                        strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_结果
                        str_结果 = "": str_日期 = ""
                    End If
                    str_标本号 = Split(str_Line, "|")(2)
                    
                End If
                str_项目 = Trim(Split(str_Line, "|")(4))
            ElseIf Left(str_Line, 1) = Chr(&H2) And Mid(str_Line, 3, 2) = "R|" Then
                
                If str_标本号 <> "" And str_项目 = Trim(Split(str_Line, "|")(2)) Then
                    str_日期 = Trim(Split(str_Line, "|")(11))
                    
                    If IsNumeric(str_日期) And Len(str_日期) > 13 Then
                        str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2)
                    End If
                    str_结果 = str_结果 & "|" & Replace(str_项目, "^", "") & "|" & Trim(Split(str_Line, "|")(3))
                End If
            End If
        Next
        
        If str_标本号 <> "" And str_结果 <> "" And str_日期 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_结果
            str_日期 = "": str_结果 = ""
        End If
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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








