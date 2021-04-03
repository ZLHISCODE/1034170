VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ctrlComm 
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   Picture         =   "ctrlComm.ctx":0000
   ScaleHeight     =   2175
   ScaleWidth      =   2190
   ToolboxBitmap   =   "ctrlComm.ctx":0842
   Begin VB.Timer timInData 
      Interval        =   3000
      Left            =   30
      Top             =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   885
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   570
      Top             =   210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   30
      Top             =   570
   End
End
Attribute VB_Name = "ctrlComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objLISComm  As clsLISComm
Attribute objLISComm.VB_VarHelpID = -1
Private strBuffer As String '数据缓冲区
'strSampleInfo：发送的标本信息
'iSendStep：发送步骤。从1开始递增，0代表不执行发送
Private strSampleInfo As String, iSendStep As Integer, dtSendTime As Date, mblnUndo As Boolean, miType As Integer
Private mstrConnType As String '连接方式
Private mstrLOG As String '保存日志信息
Private mLastrBuffer As String '保存上次解析失败时传入的数据
Public Event DataReceived()

Public Event DevOnComm(ByVal comPort As String, ByVal lngEvent As Long, ByVal strR As String)  ' 显示日志事件
Public Event DevSenComm(ByVal comPort As String, ByVal strR As String, ByVal intErr As Integer)
Public Event DevDecode(ByVal commport As String, ByVal str结果 As String)

Public Event ItemUnknown(ByVal commport As String, ByVal strItems As String) '返回未知项
Public Event ReturnCompute(ByVal strReturn As String)  '返回自动计算结果
Private mInterVal As Double  '自动应答间隔
Private mbln关闭端口并打开 As Boolean

Public Property Get CommSetting() As String
    '串口设置参数
    CommSetting = objLISComm.CommSetting
End Property

Public Property Get DevProgName() As String
    '串口设置参数
    DevProgName = objLISComm.DecodeProgName
End Property

Public Property Get AutoAnswer() As Boolean
    AutoAnswer = mInterVal > 0
End Property

Public Sub OpenPort(Optional blnShowError As Boolean = True)
'打开串口
    Dim lngBit As Long, varTemp As Variant
    Dim aCommSetting() As String
    Dim lngInterval As Long, dStart As Date
    Dim lngHost As Long
    Dim blnOpenOk As Boolean
    
    On Error GoTo OpenError
    mbln关闭端口并打开 = False
    blnOpenOk = False
    If mstrConnType = "TCPIP" Then
        '--- TCPIP方式
    
        If Not Winsock1.State = sckOpen Then
            Winsock1.Close
            lngHost = Val(Split(objLISComm.CommSetting, "|")(0))
            Winsock1.Tag = objLISComm.InputMode   '存接收模式
            If lngHost = 1 Then
                aCommSetting = Split(Split(objLISComm.CommSetting, "|")(1), ":")
                Winsock1.Protocol = sckTCPProtocol
                Winsock1.Bind Val(aCommSetting(1)), aCommSetting(0)
                Winsock1.Listen
                blnOpenOk = True
                mbln关闭端口并打开 = True
            Else
                aCommSetting = Split(Split(objLISComm.CommSetting, "|")(1), ":")
                Winsock1.Protocol = sckTCPProtocol  '设置通讯协议
                Winsock1.RemoteHost = aCommSetting(0)    '远端IP
                Winsock1.RemotePort = Val(aCommSetting(1))   '端口
                Winsock1.Connect  '连接

                blnOpenOk = True
            End If
            
        End If
    
    Else
        '-------Comm方式
        If MSComm1.PortOpen = False Then
            aCommSetting = Split(objLISComm.CommSetting, "|")
        
            MSComm1.Settings = aCommSetting(0)
            MSComm1.InputMode = Val(aCommSetting(2))
            MSComm1.RThreshold = 1
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            MSComm1.Handshaking = Val(aCommSetting(1))
            MSComm1.RTSEnable = True
            '计算串口的接收缓冲区大小
            lngBit = Val(Split(MSComm1.Settings, ",")(0))
            If lngBit = 0 Then lngBit = 9600
            
            On Error Resume Next
            MSComm1.InBufferSize = CLng(lngBit / 8) + 1 '读取的间隔为1秒
            MSComm1.InBufferSize = lngBit * 10   '两倍缓冲
            
            
'            lngInterval = CLng(1000 / (lngBit / 8)): If lngInterval < 200 Then lngInterval = 200
'
'            If lngBit <= 4800 And lngInterval <= 200 Then
'                lngInterval = 600
'            End If
'            Timer1.Interval = lngInterval
            
            '设置时间间隔
            If mInterVal > 0 Then
                If mInterVal < 0.1 And mInterVal > 3600 Then mInterVal = 0.5
                Timer1.Interval = mInterVal * 1000
                Timer1.Enabled = True
            Else
                Timer1.Enabled = False
            End If
            
            On Error GoTo OpenError
            
            MSComm1.PortOpen = True
            blnOpenOk = True


        End If
    End If
    
    If blnOpenOk Then
'            '向设备发送开始传送命令
        varTemp = objLISComm.GetDeviceStartCmd
        If Len(varTemp) > 0 Then Call SendCmd(varTemp)
    End If
    
    Exit Sub
OpenError:
    If Err.Number = 8005 Or Err.Description Like "*端口*" Then
        If blnShowError Then MsgBox Err.Description, vbInformation, gstrSysName
    ElseIf blnShowError Then
        If gobjComLib.ErrCenter() = 1 Then Resume
    End If
    Call WriteLog("CtrlComm.OpenPort", LOG_错误日志, Err.Number, Err.Description)
End Sub

Public Sub ClosePort()

    '关闭连接
    Dim varTemp As Variant
    
    On Error Resume Next
    Timer1.Enabled = False
    If mstrConnType = "TCPIP" Then
        If Winsock1.State = sckOpen Then
            
            '向设备发送停止传送命令
            varTemp = objLISComm.GetDeviceEndCmd
            If Len(varTemp) > 0 Then Call Winsock1.SendData(varTemp)
            '关闭连接
            mbln关闭端口并打开 = False
            Winsock1.Close

        End If
    Else
        If MSComm1.PortOpen = True Then
            
            '向设备发送停止传送命令
            varTemp = objLISComm.GetDeviceEndCmd
            If Len(varTemp) > 0 Then MSComm1.Output = varTemp
            '关闭连接
            MSComm1.PortOpen = False
        End If
    End If
    SaveData
End Sub

Public Function SendSample(ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As Boolean
    '向仪器发送标本信息
    On Error Resume Next
    Dim strSendData As String
    
    If iSendStep > 0 Then SendSample = False '当前正在发送，不能发送新的
    
    If Len(strBuffer) > 0 Then '当前正在接收数据，不能发送
        SendSample = False
    Else
        strSampleInfo = objLISComm.GetSampleInfo(lngDeviceID, strSampleDate, strSampleNO, "", strAdviceIDs, iType)

        mblnUndo = blnUndo: miType = iType
        iSendStep = 0 '开始发送
        strSendData = objLISComm.SendSample(iSendStep, strSampleInfo, SendSample, "", blnUndo, iType)
        Call SendCmd(strSendData)

        If Not SendSample Then
            iSendStep = 0 '如果传输失败，则取消发送
        Else
            dtSendTime = Now
        End If
    End If
End Function

Public Property Get PortOpened() As Boolean
    If mstrConnType = "TCPIP" Then
        PortOpened = (Winsock1.State = sckConnected)
    Else
        PortOpened = MSComm1.PortOpen
    End If
End Property

Public Property Get DeviceID() As Long
    DeviceID = objLISComm.DeviceID
End Property

Public Sub InitContrl(ByVal intIndex As Integer)
'初始化控件
'   intindex: 索引
'
    
    If objLISComm Is Nothing Then Set objLISComm = New clsLISComm
    mInterVal = 0
    If g仪器(intIndex).ID > 0 Then
    
        If g仪器(intIndex).类型 = 0 Then
            mstrConnType = "COM"
            MSComm1.commport = g仪器(intIndex).COM口
        Else
            mstrConnType = "TCPIP"
        End If
        
        mInterVal = Val(g仪器(intIndex).自动应答)
        
        objLISComm.InitClsLisComm intIndex
        Call OpenPort
    End If
End Sub

Private Sub MSComm1_OnComm()
    Dim lngMaxSize As Long
    Dim strInstr As String
    Dim byt_Bit() As Byte '-接收二进制数据
    Dim i As Integer
    Dim strRt As String '返回的字串
    Dim blnTimeEnable  As Boolean
    '--------暂停计时器
    blnTimeEnable = False
    If Timer1.Enabled Then
        blnTimeEnable = True
        Timer1.Enabled = False
    End If
    
    timInData.Enabled = False
    
    strRt = ""
    Select Case MSComm1.CommEvent
        Case comEventRxOver '接收缓冲区已满
            Call WriteLog("CtrlComm.MSCOMM1_OnComm", LOG_错误日志, vbObjectError + 1, "接收缓冲区满，当前值为：" & MSComm1.InBufferSize)
            'objLISComm.WriteErrorLog 2, vbObjectError + 1, "接收缓冲区满，当前值为：" & MSComm1.InBufferSize
        Case comEventTxFull '传输缓冲区已满
'            objLISComm.WriteErrorLog 2, vbObjectError + 1, "传输缓冲区满，当前值为：" & MSComm1.OutBufferSize
            Call WriteLog("CtrlComm.MSCOMM1_OnComm", LOG_错误日志, vbObjectError + 1, "传输缓冲区满，当前值为：" & MSComm1.InBufferSize)
        Case comEvReceive
            
            
            If MSComm1.InputMode = comInputModeText Then
                strInstr = MSComm1.Input
                strBuffer = strBuffer & strInstr
                strRt = strInstr
            Else
                byt_Bit = MSComm1.Input
                For i = 0 To UBound(byt_Bit)
                    strBuffer = strBuffer & "," & IIf(Len(Hex(byt_Bit(i))) = 1, "0" & Hex(byt_Bit(i)), Hex(byt_Bit(i)))
                    strRt = strRt & "," & IIf(Len(Hex(byt_Bit(i))) = 1, "0" & Hex(byt_Bit(i)), Hex(byt_Bit(i)))
                Next
                
            End If
            RaiseEvent DevOnComm(MSComm1.commport, MSComm1.CommEvent, strRt)
            If MSComm1.InBufferCount <= 0 Then
                Call SaveData
            End If
            
        Case 3 To 5
            ' clear-to-send ，data-set ready ，carrier detect   线变化
            
        Case 1001 To 1011
            WriteLog "MScomm_onComm", LOG_错误日志, MSComm1.CommEvent, "COM通讯错误"
    End Select
    
    If blnTimeEnable Then Timer1.Enabled = True  '恢复定时器
    timInData.Enabled = True
    ' 触发显示日志事件
    
End Sub



Private Sub objLISComm_AutoCompute(ByVal strReturn As String)
    '传自动计算结果
    RaiseEvent ReturnCompute(strReturn)
End Sub

Private Sub objLISComm_Decode(ByVal strReturn As String)
    '传解析结果
    Call Return_Decode(strReturn)
End Sub

Private Sub objLISComm_DecodeErr(ByVal strErr As String)
    '传错误提示
    Call SendCmd(strErr, 1)
End Sub

Private Sub objLISComm_ItemUnknown(ByVal strItems As String)
    '返回ItemUnknown项
    If strItems = "" Then Exit Sub
    If mstrConnType = "TCPIP" Then
        If Winsock1.RemoteHostIP = "" Then
            RaiseEvent ItemUnknown(Winsock1.LocalIP, strItems)
        Else
            RaiseEvent ItemUnknown(Winsock1.RemoteHostIP, strItems)
        End If
    Else
        RaiseEvent ItemUnknown(MSComm1.commport, strItems)
    End If
End Sub

Private Sub Timer1_Timer()
    '定时调用自动应答指令，发往仪器。
    Dim strCmd As String
    If mInterVal > 0 Then
        Timer1.Enabled = False
        If MSComm1.InBufferCount <= 0 Then
            strCmd = objLISComm.GetDeviceAnswerCmd
            If strCmd <> "" Then Call SendCmd(strCmd)
        End If
        Timer1.Enabled = True
    End If
End Sub

Private Function SaveData() As Boolean
    '保存数据到本地硬盘
    Dim strResult As String, strReserved As String, strCmd As String
    Dim lngDataID As Long
    Dim blnGetSample As Boolean
    Dim aSampleInfo() As String, aSamples() As String, i As Long
    Dim strResponse As String, blnSuccess As Boolean
    Dim strSampleNO As String, aTmp() As String, strBarcode As String
    Dim strSendData As String
    Dim blnClearData As Boolean
    
    On Error GoTo ErrHandle
    

    
    SaveData = False
    
    strSendData = "" '初始化临时变量
    
    If Len(strBuffer) = 0 Then
        '如果发送超时3秒，则取消发送
        If iSendStep = 0 Or DateAdd("s", 3, dtSendTime) > Now Then Exit Function
    End If
    
    strResponse = strBuffer
    '保存原始的接收数据
    blnClearData = gblnClearData
    If blnClearData Then lngDataID = objLISComm.SaveToLocal(strBuffer)
    
    '如果当前在发送期间
    If iSendStep > 0 Then
        
        strSendData = objLISComm.SendSample(iSendStep, strSampleInfo, blnSuccess, strResponse, mblnUndo, miType)
        
        Call SendCmd(strSendData)
                 
        If Not blnSuccess Then
            iSendStep = 0 '如果传输失败，则取消发送
        Else
            strBuffer = ""
        End If
        Exit Function
    End If
'    If lngDataID = 0 Then Exit Function
    '上次传入数据和本次传入数据相同，则不解析。
    strCmd = ""
    
    If strBuffer <> mLastrBuffer Or mLastrBuffer <> "" Then
        '还在接收数据，退出
        If mstrConnType = "TCPIP" Then
            If Winsock1.BytesReceived > 0 Then
                WriteLog "SaveData", LOG_错误日志, 0, "缓冲区还有数据，退出"
                Exit Function
            End If
        Else
            If MSComm1.InBufferCount > 0 Then
                WriteLog "SaveData", LOG_错误日志, 0, "缓冲区还有数据，退出"
                Exit Function
            End If
        End If
        '解析数据串并保存
    
        If Not objLISComm.Analyse(lngDataID, strResult, strReserved, strCmd, strBuffer, blnGetSample) Then Exit Function
    Else
        Exit Function
    End If
    '有结果，则清空mLastrBuffer，无结果，则保留上次传入数据在mLastrBuffer
    If strResult = "" And strReserved = strBuffer Then
        mLastrBuffer = strBuffer
    Else
        mLastrBuffer = ""
    End If
    '为解析的数据继续保留
    strBuffer = strReserved
    If Not blnGetSample Then
        
        '发送已解析完成消息（）
        If Len(strCmd) > 0 Then
            Call SendCmd(strCmd)
        End If
        If Len(strResult) > 0 Then SaveData = True

    Else
        '发送已解析完成消息（）
        If Len(strCmd) > 0 Then
            Call SendCmd(strCmd)
        End If
        
        If Len(strResult) > 0 Then '向仪器发送标本信息
            aSamples = Split(strResult, "||")
            strSampleInfo = "": miType = 0
            For i = 0 To UBound(aSamples)
                aSampleInfo = Split(aSamples(i), "|")
                If UBound(aSampleInfo) > 0 Then
                    aTmp = Split(aSampleInfo(1), "^")
                    If UBound(aTmp) = 0 Then
                        strSampleNO = Val(aTmp(0)): miType = 0: strBarcode = ""
                    Else
                        strSampleNO = Val(aTmp(0)): miType = Val(aTmp(1)): strBarcode = ""
                        If UBound(aTmp) > 1 Then
                            strBarcode = Trim(aTmp(2))
                        End If
                    End If
                    strSampleInfo = strSampleInfo & "||" & objLISComm.GetSampleInfo(objLISComm.DeviceID, _
                        Format(aSampleInfo(0), "yyyy-mm-dd"), strSampleNO, strBarcode, , miType)
                End If
            Next
            If Len(strSampleInfo) > 0 Then strSampleInfo = Mid(strSampleInfo, 3)
            
            mblnUndo = False
            iSendStep = 0 '开始发送
            strSendData = objLISComm.SendSample(iSendStep, strSampleInfo, blnSuccess, strResponse, mblnUndo, miType)
            Call SendCmd(strSendData)
            If Not blnSuccess Then
                iSendStep = 0 '如果传输失败，则取消发送
            Else
                dtSendTime = Now
            End If
        End If
    End If
    Exit Function
ErrHandle:
    Call WriteLog("CtrlComm.SaveData", LOG_错误日志, Err.Number, Err.Description)
    Call SendCmd(Err.Number & " " & Err.Description, 1)
End Function

Private Sub timInData_Timer()
    '每两秒触发保存数据的过程
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim strLine As String, strPath As String
    On Error GoTo errH
 
    strPath = App.Path & "\Apply\Decode\"
    If Not objLISComm Is Nothing Then
        
        strFileName = Dir(strPath & objLISComm.DeviceID & "_*_*.txt")
        
        Do While strFileName <> ""
            strFileName = strPath & strFileName
            If objFileSystem.FileExists(strFileName) Then
                Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
                strLine = ""
                Do
                    If objStream.AtEndOfStream Then Exit Do
                    strLine = strLine & objStream.ReadLine
                Loop
                objStream.Close
                Set objStream = Nothing
                '收到新的数据，暂时不保存数据，退出
                If mstrConnType = "TCPIP" Then
                    If Winsock1.BytesReceived > 0 Then
                        Exit Sub
                    End If
                Else
                    If MSComm1.InBufferCount > 0 Then
                        Exit Sub
                    End If
                End If
                If strLine <> "" Then
                    strLine = Replace(strLine, "CHR(10) CHR(13)", vbCrLf)
                    If objLISComm.InDataBase(strLine) = True Then RaiseEvent DataReceived
                End If
            End If
            If objFileSystem.FileExists(strFileName) Then Kill strFileName
            strFileName = Dir(strPath & objLISComm.DeviceID & "_*_*.txt")
        Loop
    End If
 
    Exit Sub
errH:
    'Resume
    WriteLog "timInData", LOG_错误日志, Err.Number, Err.Description
End Sub

Private Sub UserControl_Initialize()
    strBuffer = ""
    iSendStep = 0 '开始不执行发送
End Sub

Private Sub UserControl_Terminate()
    Set objLISComm = Nothing
    ClosePort
End Sub

Private Sub Winsock1_Close()
    '关闭成功，写日志
    'Call WriteLog("Winsock1_Close", LOG_通讯日志, 0, "关闭")
    If mbln关闭端口并打开 Then
        '主机模式，关闭后，重新初始化，以供下次连接
        Call OpenPort
    End If
End Sub

Private Sub Winsock1_Connect()
    '连接成功,写日志
    Call WriteLog("Winsock1_Connect", LOG_通讯日志, 0, "连接")
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    '只支持一个连接请求
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Winsock1.Accept requestID
    Call WriteLog("Winsock1_ConnectionRequest", LOG_通讯日志, 0, "接受" & requestID & "连接")
'                    varTemp = objLISComm.GetDeviceStartCmd
'                    If Len(varTemp) > 0 Then Call Winsock1.SendData(varTemp)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    '接到数据，保存到strBuffer中
    Dim strData As String
    Dim byt_Bit() As Byte '-接收二进制数据
    Dim i As Long
    
    Dim blnTimeEnable  As Boolean
    '--------暂停计时器
    blnTimeEnable = False
    If Timer1.Enabled Then
        blnTimeEnable = True
        Timer1.Enabled = False
    End If
    timInData.Enabled = False
    If Val(Winsock1.Tag) = 0 Then
        Winsock1.GetData strData
        strBuffer = strBuffer & strData
    Else
        Winsock1.GetData byt_Bit, vbByte
        For i = 0 To UBound(byt_Bit)
            strBuffer = strBuffer & "," & IIf(Len(Hex(byt_Bit(i))) = 1, "0" & Hex(byt_Bit(i)), Hex(byt_Bit(i)))
            strData = strData & "," & IIf(Len(Hex(byt_Bit(i))) = 1, "0" & Hex(byt_Bit(i)), Hex(byt_Bit(i)))
        Next
    End If
    
    If Winsock1.RemoteHostIP = "" Then
        RaiseEvent DevOnComm(Winsock1.LocalIP, -1, strData)
    Else
        RaiseEvent DevOnComm(Winsock1.RemoteHostIP, -1, strData)
    End If
    '保存数据，触发事件
    If Winsock1.BytesReceived <= 0 Then
        Call SaveData
    End If
    timInData.Enabled = True
    If blnTimeEnable Then Timer1.Enabled = True '恢复定时器
    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '返回错误,同mscomm的onComm事件功能相同
'    Call objLISComm.WriteErrorLog(2, Number, Description)
    Call WriteLog("CtrlComm.Winsock1_Error", LOG_错误日志, Number, Description)
End Sub

Private Sub Winsock1_SendComplete()
    '发送成功,写日志
End Sub

Private Sub SendCmd(ByVal strSendCmd As String, Optional intErr As Integer = 0)
    '发送消息
    'interr= 0时才发送，为1时不发送到仪器
    Dim bitByte() As Byte
    Dim lngBits As Long, lngloop As Long
    Dim strCode As String
    Dim ReturnBin As Boolean
    Dim blnErr As Boolean
    On Error GoTo errH
    If strSendCmd = "" Then Exit Sub
    
    
    '根据接收模式确定发送模式
    If mstrConnType = "TCPIP" Then
        ReturnBin = Val(Winsock1.Tag) = 1
    Else
        ReturnBin = MSComm1.InputMode = comInputModeBinary
    End If
    
    If ReturnBin Then
        '二进制数据，转为字符数组
        strCode = strSendCmd
        lngBits = Len(strCode) / 3
        If lngBits > 0 Then
            ReDim bitByte(lngBits - 1)
            For lngloop = 1 To lngBits
                bitByte(lngloop - 1) = Val("&H" & Mid(Left(strCode, 3), 2))
                strCode = Mid(strCode, 4)
            Next
        Else
            blnErr = True
            WriteLog "sendcmd", LOG_错误日志, 1, "不是二进制格式的数据！" & vbNewLine & strSendCmd
        End If
    End If
    
    If mstrConnType = "TCPIP" Then
        If intErr = 0 Then
            If ReturnBin Then
                Call Winsock1.SendData(bitByte)    '传字符数组
            Else
                Call Winsock1.SendData(strSendCmd) '传文本
            End If
        End If
        If Winsock1.RemoteHostIP = "" Then
            RaiseEvent DevSenComm(Winsock1.LocalIP, strSendCmd, intErr)
        Else
            RaiseEvent DevSenComm(Winsock1.RemoteHostIP, strSendCmd, intErr)
        End If
    Else
        If intErr = 0 Then
            If ReturnBin Then
                If blnErr = False Then
                    MSComm1.Output = bitByte
                End If
            Else
                MSComm1.Output = strSendCmd
            End If
        End If
        RaiseEvent DevSenComm(MSComm1.commport, strSendCmd, intErr)
    End If
    Exit Sub
errH:
    WriteLog "sendcmd", LOG_错误日志, Err.Number, Err.Description & vbNewLine & strSendCmd
End Sub

Private Sub Return_Decode(ByVal strDecode As String)
    '返回解码结果
    If strDecode = "" Then Exit Sub
    If mstrConnType = "TCPIP" Then
        If Winsock1.RemoteHostIP = "" Then
            RaiseEvent DevDecode(Winsock1.LocalIP, strDecode)
        Else
            RaiseEvent DevDecode(Winsock1.RemoteHostIP, strDecode)
        End If
    Else
        RaiseEvent DevDecode(MSComm1.commport, strDecode)
    End If
End Sub
