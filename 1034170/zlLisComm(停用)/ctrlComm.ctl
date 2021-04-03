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
Private strBuffer As String '���ݻ�����
'strSampleInfo�����͵ı걾��Ϣ
'iSendStep�����Ͳ��衣��1��ʼ������0����ִ�з���
Private strSampleInfo As String, iSendStep As Integer, dtSendTime As Date, mblnUndo As Boolean, miType As Integer
Private mstrConnType As String '���ӷ�ʽ
Private mstrLOG As String '������־��Ϣ
Private mLastrBuffer As String '�����ϴν���ʧ��ʱ���������
Public Event DataReceived()

Public Event DevOnComm(ByVal comPort As String, ByVal lngEvent As Long, ByVal strR As String)  ' ��ʾ��־�¼�
Public Event DevSenComm(ByVal comPort As String, ByVal strR As String, ByVal intErr As Integer)
Public Event DevDecode(ByVal commport As String, ByVal str��� As String)

Public Event ItemUnknown(ByVal commport As String, ByVal strItems As String) '����δ֪��
Public Event ReturnCompute(ByVal strReturn As String)  '�����Զ�������
Private mInterVal As Double  '�Զ�Ӧ����
Private mbln�رն˿ڲ��� As Boolean

Public Property Get CommSetting() As String
    '�������ò���
    CommSetting = objLISComm.CommSetting
End Property

Public Property Get DevProgName() As String
    '�������ò���
    DevProgName = objLISComm.DecodeProgName
End Property

Public Property Get AutoAnswer() As Boolean
    AutoAnswer = mInterVal > 0
End Property

Public Sub OpenPort(Optional blnShowError As Boolean = True)
'�򿪴���
    Dim lngBit As Long, varTemp As Variant
    Dim aCommSetting() As String
    Dim lngInterval As Long, dStart As Date
    Dim lngHost As Long
    Dim blnOpenOk As Boolean
    
    On Error GoTo OpenError
    mbln�رն˿ڲ��� = False
    blnOpenOk = False
    If mstrConnType = "TCPIP" Then
        '--- TCPIP��ʽ
    
        If Not Winsock1.State = sckOpen Then
            Winsock1.Close
            lngHost = Val(Split(objLISComm.CommSetting, "|")(0))
            Winsock1.Tag = objLISComm.InputMode   '�����ģʽ
            If lngHost = 1 Then
                aCommSetting = Split(Split(objLISComm.CommSetting, "|")(1), ":")
                Winsock1.Protocol = sckTCPProtocol
                Winsock1.Bind Val(aCommSetting(1)), aCommSetting(0)
                Winsock1.Listen
                blnOpenOk = True
                mbln�رն˿ڲ��� = True
            Else
                aCommSetting = Split(Split(objLISComm.CommSetting, "|")(1), ":")
                Winsock1.Protocol = sckTCPProtocol  '����ͨѶЭ��
                Winsock1.RemoteHost = aCommSetting(0)    'Զ��IP
                Winsock1.RemotePort = Val(aCommSetting(1))   '�˿�
                Winsock1.Connect  '����

                blnOpenOk = True
            End If
            
        End If
    
    Else
        '-------Comm��ʽ
        If MSComm1.PortOpen = False Then
            aCommSetting = Split(objLISComm.CommSetting, "|")
        
            MSComm1.Settings = aCommSetting(0)
            MSComm1.InputMode = Val(aCommSetting(2))
            MSComm1.RThreshold = 1
            MSComm1.InBufferCount = 0
            MSComm1.InputLen = 0
            MSComm1.Handshaking = Val(aCommSetting(1))
            MSComm1.RTSEnable = True
            '���㴮�ڵĽ��ջ�������С
            lngBit = Val(Split(MSComm1.Settings, ",")(0))
            If lngBit = 0 Then lngBit = 9600
            
            On Error Resume Next
            MSComm1.InBufferSize = CLng(lngBit / 8) + 1 '��ȡ�ļ��Ϊ1��
            MSComm1.InBufferSize = lngBit * 10   '��������
            
            
'            lngInterval = CLng(1000 / (lngBit / 8)): If lngInterval < 200 Then lngInterval = 200
'
'            If lngBit <= 4800 And lngInterval <= 200 Then
'                lngInterval = 600
'            End If
'            Timer1.Interval = lngInterval
            
            '����ʱ����
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
'            '���豸���Ϳ�ʼ��������
        varTemp = objLISComm.GetDeviceStartCmd
        If Len(varTemp) > 0 Then Call SendCmd(varTemp)
    End If
    
    Exit Sub
OpenError:
    If Err.Number = 8005 Or Err.Description Like "*�˿�*" Then
        If blnShowError Then MsgBox Err.Description, vbInformation, gstrSysName
    ElseIf blnShowError Then
        If gobjComLib.ErrCenter() = 1 Then Resume
    End If
    Call WriteLog("CtrlComm.OpenPort", LOG_������־, Err.Number, Err.Description)
End Sub

Public Sub ClosePort()

    '�ر�����
    Dim varTemp As Variant
    
    On Error Resume Next
    Timer1.Enabled = False
    If mstrConnType = "TCPIP" Then
        If Winsock1.State = sckOpen Then
            
            '���豸����ֹͣ��������
            varTemp = objLISComm.GetDeviceEndCmd
            If Len(varTemp) > 0 Then Call Winsock1.SendData(varTemp)
            '�ر�����
            mbln�رն˿ڲ��� = False
            Winsock1.Close

        End If
    Else
        If MSComm1.PortOpen = True Then
            
            '���豸����ֹͣ��������
            varTemp = objLISComm.GetDeviceEndCmd
            If Len(varTemp) > 0 Then MSComm1.Output = varTemp
            '�ر�����
            MSComm1.PortOpen = False
        End If
    End If
    SaveData
End Sub

Public Function SendSample(ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As Boolean
    '���������ͱ걾��Ϣ
    On Error Resume Next
    Dim strSendData As String
    
    If iSendStep > 0 Then SendSample = False '��ǰ���ڷ��ͣ����ܷ����µ�
    
    If Len(strBuffer) > 0 Then '��ǰ���ڽ������ݣ����ܷ���
        SendSample = False
    Else
        strSampleInfo = objLISComm.GetSampleInfo(lngDeviceID, strSampleDate, strSampleNO, "", strAdviceIDs, iType)

        mblnUndo = blnUndo: miType = iType
        iSendStep = 0 '��ʼ����
        strSendData = objLISComm.SendSample(iSendStep, strSampleInfo, SendSample, "", blnUndo, iType)
        Call SendCmd(strSendData)

        If Not SendSample Then
            iSendStep = 0 '�������ʧ�ܣ���ȡ������
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
'��ʼ���ؼ�
'   intindex: ����
'
    
    If objLISComm Is Nothing Then Set objLISComm = New clsLISComm
    mInterVal = 0
    If g����(intIndex).ID > 0 Then
    
        If g����(intIndex).���� = 0 Then
            mstrConnType = "COM"
            MSComm1.commport = g����(intIndex).COM��
        Else
            mstrConnType = "TCPIP"
        End If
        
        mInterVal = Val(g����(intIndex).�Զ�Ӧ��)
        
        objLISComm.InitClsLisComm intIndex
        Call OpenPort
    End If
End Sub

Private Sub MSComm1_OnComm()
    Dim lngMaxSize As Long
    Dim strInstr As String
    Dim byt_Bit() As Byte '-���ն���������
    Dim i As Integer
    Dim strRt As String '���ص��ִ�
    Dim blnTimeEnable  As Boolean
    '--------��ͣ��ʱ��
    blnTimeEnable = False
    If Timer1.Enabled Then
        blnTimeEnable = True
        Timer1.Enabled = False
    End If
    
    timInData.Enabled = False
    
    strRt = ""
    Select Case MSComm1.CommEvent
        Case comEventRxOver '���ջ���������
            Call WriteLog("CtrlComm.MSCOMM1_OnComm", LOG_������־, vbObjectError + 1, "���ջ�����������ǰֵΪ��" & MSComm1.InBufferSize)
            'objLISComm.WriteErrorLog 2, vbObjectError + 1, "���ջ�����������ǰֵΪ��" & MSComm1.InBufferSize
        Case comEventTxFull '���仺��������
'            objLISComm.WriteErrorLog 2, vbObjectError + 1, "���仺����������ǰֵΪ��" & MSComm1.OutBufferSize
            Call WriteLog("CtrlComm.MSCOMM1_OnComm", LOG_������־, vbObjectError + 1, "���仺����������ǰֵΪ��" & MSComm1.InBufferSize)
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
            ' clear-to-send ��data-set ready ��carrier detect   �߱仯
            
        Case 1001 To 1011
            WriteLog "MScomm_onComm", LOG_������־, MSComm1.CommEvent, "COMͨѶ����"
    End Select
    
    If blnTimeEnable Then Timer1.Enabled = True  '�ָ���ʱ��
    timInData.Enabled = True
    ' ������ʾ��־�¼�
    
End Sub



Private Sub objLISComm_AutoCompute(ByVal strReturn As String)
    '���Զ�������
    RaiseEvent ReturnCompute(strReturn)
End Sub

Private Sub objLISComm_Decode(ByVal strReturn As String)
    '���������
    Call Return_Decode(strReturn)
End Sub

Private Sub objLISComm_DecodeErr(ByVal strErr As String)
    '��������ʾ
    Call SendCmd(strErr, 1)
End Sub

Private Sub objLISComm_ItemUnknown(ByVal strItems As String)
    '����ItemUnknown��
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
    '��ʱ�����Զ�Ӧ��ָ�����������
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
    '�������ݵ�����Ӳ��
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
    
    strSendData = "" '��ʼ����ʱ����
    
    If Len(strBuffer) = 0 Then
        '������ͳ�ʱ3�룬��ȡ������
        If iSendStep = 0 Or DateAdd("s", 3, dtSendTime) > Now Then Exit Function
    End If
    
    strResponse = strBuffer
    '����ԭʼ�Ľ�������
    blnClearData = gblnClearData
    If blnClearData Then lngDataID = objLISComm.SaveToLocal(strBuffer)
    
    '�����ǰ�ڷ����ڼ�
    If iSendStep > 0 Then
        
        strSendData = objLISComm.SendSample(iSendStep, strSampleInfo, blnSuccess, strResponse, mblnUndo, miType)
        
        Call SendCmd(strSendData)
                 
        If Not blnSuccess Then
            iSendStep = 0 '�������ʧ�ܣ���ȡ������
        Else
            strBuffer = ""
        End If
        Exit Function
    End If
'    If lngDataID = 0 Then Exit Function
    '�ϴδ������ݺͱ��δ���������ͬ���򲻽�����
    strCmd = ""
    
    If strBuffer <> mLastrBuffer Or mLastrBuffer <> "" Then
        '���ڽ������ݣ��˳�
        If mstrConnType = "TCPIP" Then
            If Winsock1.BytesReceived > 0 Then
                WriteLog "SaveData", LOG_������־, 0, "�������������ݣ��˳�"
                Exit Function
            End If
        Else
            If MSComm1.InBufferCount > 0 Then
                WriteLog "SaveData", LOG_������־, 0, "�������������ݣ��˳�"
                Exit Function
            End If
        End If
        '�������ݴ�������
    
        If Not objLISComm.Analyse(lngDataID, strResult, strReserved, strCmd, strBuffer, blnGetSample) Then Exit Function
    Else
        Exit Function
    End If
    '�н���������mLastrBuffer���޽���������ϴδ���������mLastrBuffer
    If strResult = "" And strReserved = strBuffer Then
        mLastrBuffer = strBuffer
    Else
        mLastrBuffer = ""
    End If
    'Ϊ���������ݼ�������
    strBuffer = strReserved
    If Not blnGetSample Then
        
        '�����ѽ��������Ϣ����
        If Len(strCmd) > 0 Then
            Call SendCmd(strCmd)
        End If
        If Len(strResult) > 0 Then SaveData = True

    Else
        '�����ѽ��������Ϣ����
        If Len(strCmd) > 0 Then
            Call SendCmd(strCmd)
        End If
        
        If Len(strResult) > 0 Then '���������ͱ걾��Ϣ
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
            iSendStep = 0 '��ʼ����
            strSendData = objLISComm.SendSample(iSendStep, strSampleInfo, blnSuccess, strResponse, mblnUndo, miType)
            Call SendCmd(strSendData)
            If Not blnSuccess Then
                iSendStep = 0 '�������ʧ�ܣ���ȡ������
            Else
                dtSendTime = Now
            End If
        End If
    End If
    Exit Function
ErrHandle:
    Call WriteLog("CtrlComm.SaveData", LOG_������־, Err.Number, Err.Description)
    Call SendCmd(Err.Number & " " & Err.Description, 1)
End Function

Private Sub timInData_Timer()
    'ÿ���봥���������ݵĹ���
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
                '�յ��µ����ݣ���ʱ���������ݣ��˳�
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
    WriteLog "timInData", LOG_������־, Err.Number, Err.Description
End Sub

Private Sub UserControl_Initialize()
    strBuffer = ""
    iSendStep = 0 '��ʼ��ִ�з���
End Sub

Private Sub UserControl_Terminate()
    Set objLISComm = Nothing
    ClosePort
End Sub

Private Sub Winsock1_Close()
    '�رճɹ���д��־
    'Call WriteLog("Winsock1_Close", LOG_ͨѶ��־, 0, "�ر�")
    If mbln�رն˿ڲ��� Then
        '����ģʽ���رպ����³�ʼ�����Թ��´�����
        Call OpenPort
    End If
End Sub

Private Sub Winsock1_Connect()
    '���ӳɹ�,д��־
    Call WriteLog("Winsock1_Connect", LOG_ͨѶ��־, 0, "����")
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    'ֻ֧��һ����������
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Winsock1.Accept requestID
    Call WriteLog("Winsock1_ConnectionRequest", LOG_ͨѶ��־, 0, "����" & requestID & "����")
'                    varTemp = objLISComm.GetDeviceStartCmd
'                    If Len(varTemp) > 0 Then Call Winsock1.SendData(varTemp)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    '�ӵ����ݣ����浽strBuffer��
    Dim strData As String
    Dim byt_Bit() As Byte '-���ն���������
    Dim i As Long
    
    Dim blnTimeEnable  As Boolean
    '--------��ͣ��ʱ��
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
    '�������ݣ������¼�
    If Winsock1.BytesReceived <= 0 Then
        Call SaveData
    End If
    timInData.Enabled = True
    If blnTimeEnable Then Timer1.Enabled = True '�ָ���ʱ��
    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '���ش���,ͬmscomm��onComm�¼�������ͬ
'    Call objLISComm.WriteErrorLog(2, Number, Description)
    Call WriteLog("CtrlComm.Winsock1_Error", LOG_������־, Number, Description)
End Sub

Private Sub Winsock1_SendComplete()
    '���ͳɹ�,д��־
End Sub

Private Sub SendCmd(ByVal strSendCmd As String, Optional intErr As Integer = 0)
    '������Ϣ
    'interr= 0ʱ�ŷ��ͣ�Ϊ1ʱ�����͵�����
    Dim bitByte() As Byte
    Dim lngBits As Long, lngloop As Long
    Dim strCode As String
    Dim ReturnBin As Boolean
    Dim blnErr As Boolean
    On Error GoTo errH
    If strSendCmd = "" Then Exit Sub
    
    
    '���ݽ���ģʽȷ������ģʽ
    If mstrConnType = "TCPIP" Then
        ReturnBin = Val(Winsock1.Tag) = 1
    Else
        ReturnBin = MSComm1.InputMode = comInputModeBinary
    End If
    
    If ReturnBin Then
        '���������ݣ�תΪ�ַ�����
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
            WriteLog "sendcmd", LOG_������־, 1, "���Ƕ����Ƹ�ʽ�����ݣ�" & vbNewLine & strSendCmd
        End If
    End If
    
    If mstrConnType = "TCPIP" Then
        If intErr = 0 Then
            If ReturnBin Then
                Call Winsock1.SendData(bitByte)    '���ַ�����
            Else
                Call Winsock1.SendData(strSendCmd) '���ı�
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
    WriteLog "sendcmd", LOG_������־, Err.Number, Err.Description & vbNewLine & strSendCmd
End Sub

Private Sub Return_Decode(ByVal strDecode As String)
    '���ؽ�����
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
