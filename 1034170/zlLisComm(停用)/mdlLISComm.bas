Attribute VB_Name = "mdlLISComm"
Option Explicit

'Public gcnOracle As ADODB.Connection    '�������ݿ�����
Public gstrSQL As String

'Public gstrSysName As String                'ϵͳ����

Public lngExeDeptID As Long 'ִ�п���
Public ParentWnd As Object
Public blnDataReceived As Boolean
'------������ͼ�괦��
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_ACTIVATE = &H6
Public Const WM_KEYDOWN = &H100
Public Const WM_PAINT = &HF

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'Public Const GWL_EXSTYLE = (-20)
'Public Const WinStyle = &H40000
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4

'ø���ǲ���
Public glngMBDeviceID As Long, gstrMBChannel As String, glngMBNo As Long, gstrMBPosition As String

Private mItem() As Variant

Public Const LOG_������־ = 0
Public Const LOG_ͨѶ��־ = 1
Public Const LOG_δ֪�� = 2

Public pLast������־ As String '�ϴδ�����Ϣ,���ڱ�������ظ�����־
Public pLastͨѶ��־ As String
Public mMakeNoRule As String    '�걾�������ʱ�����

Public gblnFromDB As Boolean ' �Ƿ��Ǵ����ݿ��ȡ����.

Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip

Public Sub SavePortsSetting()
'���ܣ��������Ӽ��������Ĵ�������
    Dim i As Integer
    Dim strSet As String
    Dim aPorts As Variant
    On Error GoTo errH
    
    strSet = ""
    If gblnFromDB Then
        '���ԭ��������
        Call gobjDatabase.SetPara("������������", "", glngSys, 1208)
        For i = LBound(g����) To UBound(g����)
            '����id , ����, COM��, ������, ����λ, У��λ, ֹͣλ, ����, TCPIP�˿�, IP��ַ, �ַ�ģʽ, ���Ϊ������ID, ����,�Զ�Ӧ��,�ɷ��Ѻ˱걾
            If g����(i).ID > 0 Then
                strSet = strSet & ";" & g����(i).ID & "," & g����(i).���� & "," & g����(i).COM�� & "," & g����(i).������ & _
                   "," & g����(i).����λ & "," & g����(i).У��λ & "," & g����(i).ֹͣλ & "," & g����(i).���� & _
                   "," & g����(i).IP�˿� & "," & g����(i).IP & "," & g����(i).�ַ�ģʽ & "," & g����(i).SaveAsID & "," & g����(i).���� & _
                   "," & g����(i).�Զ�Ӧ�� & "," & g����(i).�ɷ��Ѻ˱걾
            End If
        Next
        If strSet <> "" Then
            Call gobjDatabase.SetPara("������������", strSet, glngSys, 1208)
        End If
    Else
        'DeleteSetting "ZLSOFT", "����ģ��", "ZlLISSrv"
        Err = 0: On Error Resume Next
        aPorts = GetAllSettings("ZLSOFT", "����ģ��\ZlLISSrv")
        On Error GoTo errH
        If IsEmpty(aPorts) Then
            ReDim aPorts(8, 0)
            For i = 0 To 7
                aPorts(i, 0) = "COM" & i + 1
            Next
        End If
        Err = 0: On Error Resume Next
        For i = LBound(aPorts) To UBound(aPorts)
            DeleteSetting "ZLSOFT", "����ģ��\ZLLISSrv", aPorts(i, 0)
            DeleteSetting "ZLSOFT", "����ģ��\ZLLISSrv\" & aPorts(i, 0)
        Next
        On Error GoTo errH
        For i = LBound(g����) To UBound(g����)
            If g����(i).���� = 1 Then
                'TCP
                If g����(i).ID > 0 Then
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv", "IP" & g����(i).ID, "")
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Device", g����(i).ID)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Enabled", g����(i).����)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Host", g����(i).����)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "InMode", g����(i).�ַ�ģʽ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "IP", g����(i).IP)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Port", g����(i).IP�˿�)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "SaveAs", g����(i).SaveAsID)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Auto", g����(i).�Զ�Ӧ��)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "blnSend", g����(i).�ɷ��Ѻ˱걾)
                End If
            Else
                If g����(i).COM�� > 0 And g����(i).ID > 0 Then
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv", "COM" & g����(i).COM��, "")
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "Device", g����(i).ID)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "Speed", g����(i).������)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "DataBit", g����(i).����λ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "Parity", g����(i).У��λ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "StopBit", g����(i).ֹͣλ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "HandShaking", g����(i).����)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "InputMode", g����(i).�ַ�ģʽ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "SaveAs", g����(i).SaveAsID)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "Auto", g����(i).�Զ�Ӧ��)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "blnSend", g����(i).�ɷ��Ѻ˱걾)
                End If
            End If
        Next
    End If
    Exit Sub
errH:
    MsgBox Err.Description

End Sub

Public Function GetConnectDevs() As Variant
'���ܣ���ȡϵͳ���ӵļ�������
    Dim aSettings() As Variant
    Dim aPorts As Variant, i As Integer, PortIndex As Integer
    Dim lngDeviceID As Long, rsTmp As New adodb.Recordset, rsTmp1 As New adodb.Recordset
    Dim strConnType As String  '��������
    Dim strIP As String, strPort As String 'ip �� Port
    Dim varIPSet As Variant 'IP������
    Dim lngSaveAsID As Long '���Ϊ������ID
    Dim strSaveAsName As String
    
    aSettings = Array()
    
    Err = 0: On Error Resume Next
    aPorts = GetAllSettings("ZLSOFT", "����ģ��\ZlLISSrv")
    On Error GoTo errH
    If IsEmpty(aPorts) Then
        ReDim aPorts(8, 0)
        For i = 0 To 7
            aPorts(i, 0) = "COM" & i + 1
        Next
    End If
   
    If Not IsEmpty(aPorts) Then
        
        ReDim g����(UBound(aPorts))
        
        For i = LBound(g����) To UBound(g����)
            g����(i).ID = 0
            g����(i).IP = "127.0.0.1"
            g����(i).IP�˿� = 6666
            g����(i).SaveAsID = 0
            g����(i).������ = 9600
            g����(i).���� = 1
            g����(i).COM�� = 0
            g����(i).����λ = 8
            g����(i).ֹͣλ = 1
            g����(i).���� = 0
            g����(i).У��λ = "N"
            g����(i).�ַ�ģʽ = 0
            g����(i).���� = 0
            g����(i).�Զ�Ӧ�� = "0"
            g����(i).�ɷ��Ѻ˱걾 = 1
        Next
        
        For i = LBound(aPorts) To UBound(aPorts)
            
            strIP = "": strPort = ""
            lngSaveAsID = 0
            strSaveAsName = ""
            
            lngSaveAsID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "SaveAs", 0))
            If lngSaveAsID > 0 Then
                Set rsTmp1 = gobjDatabase.OpenSQLRecord("Select ���� From �������� where ID=[1]", "ȡ������������", lngSaveAsID)
                Do Until rsTmp1.EOF
                    strSaveAsName = "" & rsTmp1!����
                    rsTmp1.MoveNext
                Loop
            End If
            
            strConnType = aPorts(i, 0)

            If strConnType Like "IP*" Then
                'TCPIP����
                g����(i).���� = 1
                lngDeviceID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                
                If lngDeviceID > 0 Then

                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    gstrSQL = "Select * From �������� Where ID=" & lngDeviceID
                    OpenRecordset rsTmp, App.ProductName
                    If Not rsTmp.EOF Then

                        If Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Enabled", 0)) = 1 Then
                            '������IP��ʽ,���IP�Ͷ˿��Ƿ�Ϸ�
                            strIP = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1")
                            strPort = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Port", 6666)
                            g����(i).IP = strIP
                            g����(i).IP�˿� = Val(strPort)
                            g����(i).���� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Host", 0))
                            
                            g����(i).�Զ�Ӧ�� = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Auto", "0")
                            g����(i).�ɷ��Ѻ˱걾 = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                            If Not ValidateIP(strIP) And Not ValidatePort(strPort) Then

                                If UBound(aSettings) = -1 Then
                                    ReDim aSettings(2, 0) As Variant
                                Else
                                    ReDim Preserve aSettings(2, UBound(aSettings, 2) + 1) As Variant
                                End If

                                aSettings(0, UBound(aSettings, 2)) = strIP & ":" & strPort
                                aSettings(1, UBound(aSettings, 2)) = "IP " & strIP & " " & rsTmp("����") & IIf(strSaveAsName = "", "", " -> " & strSaveAsName)
                                aSettings(2, UBound(aSettings, 2)) = lngDeviceID
                            End If

                        End If
                    End If
                End If
            ElseIf strConnType Like "COM*" Then
                'COM����
                PortIndex = Val(Mid(aPorts(i, 0), 4)) - 1
                lngDeviceID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                g����(i).���� = 0
                g����(i).COM�� = Val(PortIndex + 1)
                If lngDeviceID > 0 Then
                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    gstrSQL = "Select * From �������� Where ID=" & lngDeviceID
                    OpenRecordset rsTmp, App.ProductName
                    If Not rsTmp.EOF Then
                        If UBound(aSettings) = -1 Then
                            ReDim aSettings(2, 0) As Variant
                        Else
                            ReDim Preserve aSettings(2, UBound(aSettings, 2) + 1) As Variant
                        End If
                        aSettings(0, UBound(aSettings, 2)) = PortIndex
                        aSettings(1, UBound(aSettings, 2)) = "COM" & PortIndex + 1 & " " & rsTmp("����") & IIf(strSaveAsName = "", "", " -> " & strSaveAsName)
                        aSettings(2, UBound(aSettings, 2)) = lngDeviceID
                    End If
                
                    With g����(i)
                        .ID = lngDeviceID
                        .������ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Speed", "9600"))
                        .����λ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "DataBit", "8"))
                        .ֹͣλ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "StopBit", "1"))
                        .У��λ = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Parity", "n")
                        .���� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & aPorts(i, 0), "HandShaking", "0"))
                        .�ַ�ģʽ = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "InputMode", "0")
                        .SaveAsID = lngSaveAsID
                        .�Զ�Ӧ�� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Auto", "0"))
                        .�ɷ��Ѻ˱걾 = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                    End With
                End If
            End If
        Next
    End If
    
    If UBound(aSettings) > -1 Then GetConnectDevs = aSettings
    Exit Function
errH:
    MsgBox Err.Description

End Function

Public Function GetDevices() As adodb.Recordset
'���ܣ���ȡ���м�������
    On Error GoTo DBError
    Set GetDevices = Nothing
    gstrSQL = "Select * From ��������"
    Set GetDevices = gobjDatabase.OpenSQLRecord(gstrSQL, "�������ݽ���")
    Exit Function
DBError:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetComboxIndex(objCbo As ComboBox, ByVal SeekValue As Long) As Long
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If objCbo.ItemData(i) = SeekValue Then Exit For
    Next
    If i > objCbo.ListCount - 1 Then i = 0
    GetComboxIndex = i
End Function

Public Sub OpenRecordset(rsTemp As adodb.Recordset, ByVal strFormCaption As String, Optional cnOracle As adodb.Connection = Nothing)
'���ܣ��򿪼�¼��ͬʱ����SQL���
    If rsTemp.State = adStateOpen Then rsTemp.Close
    If cnOracle Is Nothing Then Set cnOracle = gcnOracle
    
    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, gstrSQL)
    rsTemp.Open gstrSQL, cnOracle, adOpenStatic, adLockReadOnly
    Call gobjComLib.SQLTest
End Sub

Public Sub ExecuteProcedure(ByVal strFormCaption As String, Optional cnOracle As adodb.Connection = Nothing)
'���ܣ�ִ�й���ʽ��SQL���
    If cnOracle Is Nothing Then Set cnOracle = gcnOracle
    
    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, gstrSQL)
    cnOracle.Execute gstrSQL, , adCmdStoredProc
    Call gobjComLib.SQLTest
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub WriteLog(ByVal ModuleName As String, ByVal ErrorType As Integer, ByVal ErrorNum As Long, ByVal ErrorDesc As String)
    'Module:ģ���������
    'ErrorType:��־����
    'errorNum:����Ż���־���
    'errorDesc:������Ϣ����־��Ϣ
    Dim strSQL As String
    
    Call WriteTxtLog(ErrorType, ModuleName, IIf(ErrorNum = 0, "", " ") & ErrorDesc)
    
End Sub

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '���ܣ���������������һ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "�������ݽ���", strTip) & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub ModifyIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "", Optional ByVal blnMessage As Boolean = True)
    
    '���ܣ���������������һ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = IIf(blnMessage, WM_MOUSEMOVE, 0)
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "�������ݽ���", strTip) & Chr$(0)

    Shell_NotifyIcon NIM_MODIFY, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '���ܣ�����������ɾ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼�����������
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Public Sub ResultFromFile(ByVal strFile As String, ByVal lngDeviceID As Long, ByVal strSampleNO As String, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"))
'���ļ���ȡ������
'   strFile������·�����ļ���
'   lngDeviceID�������豸ID
'   strSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��strSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
    Dim rsTmp As New adodb.Recordset
    Dim strDevice As String
    Dim objDevice As Object, strInput As String
    Dim aRecord() As String, aItem() As String, aItemInfo() As Variant
    Dim strDate As String, strSampleID As String '2007-08-16 ��Ϊ�ַ���
    Dim strName As String, strSample As String, strSex As String, strBirth As String
    Dim iResultFlag As Integer, strResultRef As String, aResultRef() As String
    Dim i As Long, j As Long
    Dim strSQL As String, rsContent As adodb.Recordset
    Dim rsRef As New adodb.Recordset
    Dim lngID As Long
    Dim blnAuditing As Boolean '�Ƿ����
    Dim lngItemID As Long '��ĿID
    Dim strItemRecords As String
    Dim aNos() As String, iType As Integer '�걾������
    Dim blnBeginTrans As Boolean, strδ֪�� As String
    Dim intMicrobe As Integer   '΢���� =1 ��ʾ΢����
    Dim strStartDate As String
    Dim strEndDate As String
    Dim strBarcode As String
    Dim blnQryWithSampleNO As Boolean

    Dim aTmp() As String '�ָ�ͼ������
    
    If Len(Trim(strFile)) = 0 Then Exit Sub
    
    gstrSQL = "Select ͨѶ������,nvl(΢����,0) as ΢���� From �������� Where ID=" & lngDeviceID
    OpenRecordset rsTmp, App.ProductName, gcnOracle
    If Not rsTmp.EOF Then strDevice = rsTmp(0): intMicrobe = Nvl(rsTmp(1), 0)
    
    If intMicrobe = 0 Then
        gstrSQL = "Select ͨ������,��ĿID,Nvl(С��λ��,2) As С��λ�� From ����������Ŀ Where ����ID=[1]"
    Else
        gstrSQL = "Select ͨ������,������ID As ��ĿID, 2 as С��λ��  From ����ϸ������ Where ����id = [1] "
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, App.ProductName, lngDeviceID)
    
    If rsTmp.EOF Then
        ReDim mItem(1, 0) As Variant
        mItem(1, 0) = -1
    Else
        mItem = rsTmp.GetRows
    End If
    
    On Error Resume Next
    Set objDevice = CreateObject(strDevice)
    If objDevice Is Nothing Then Call WriteLog("ResultFromFile", LOG_������־, Err.Number, "��������:" & strDevice & "����ʧ��!" & vbNewLine & Err.Description)
    On Error GoTo DBError
    
    blnBeginTrans = False
    gcnOracle.BeginTrans
    blnBeginTrans = True
    Call WriteLog(strDevice & ".ResultFromFile", LOG_ͨѶ��־, 0, "strFile:" & strFile & vbNewLine & "strSampleNO:" & strSampleNO & vbNewLine & "dtStart:" & CStr(dtStart) & vbNewLine & "dtEnd:" & CStr(dtEnd))
    aRecord = objDevice.ResultFromFile(strFile, strSampleNO, dtStart, dtEnd)
    'aRecord�����صļ���������(������������밴���±�׼��֯���)
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    
    '�з��ؽ��
    If UBound(aRecord) > -1 Then
        
        
        For i = 0 To UBound(aRecord)
            Call WriteLog("mdlLISComm.ResultFromFile", LOG_ͨѶ��־, 0, "��¼" & i & ":" & aRecord(i))
            blnAuditing = False
            
            If InStr(aRecord(i), "|") > 0 Then
                aTmp = Split(aRecord(i), vbCrLf)
                aItem = Split(aTmp(0), "|")
                If UBound(aItem) > 4 Then
                    '��Ч�ı�����
                    aNos = Split(aItem(1), "^") '�걾�Ÿ�ʽ���걾��^�걾���0�����棬1�����
                    If UBound(aNos) = 0 Then
                        'û�б걾����򰴳���걾����
                        strDate = Trim(aItem(0)): strSampleID = Val(aNos(0)): iType = 0: strBarcode = ""
                    Else
                        strDate = Trim(aItem(0)): strSampleID = Val(aNos(0)): iType = Val(aNos(1)): strBarcode = ""
                        If UBound(aNos) > 1 Then
                            strBarcode = Trim(aNos(2))
                        End If
                    End If
                    '��������걾���ɹ��򣨰�ʱ�䣩
                    strStartDate = GetDateTime(mMakeNoRule, 1, strDate)
                    strEndDate = GetDateTime(mMakeNoRule, 2, strDate)
                    
                    strName = Trim(aItem(2)): strSample = Trim(aItem(3))
                    '�ж��Ƿ������걾
                      
                    '-------------------------------------------------------------------------------
                    If Len(Trim(strBarcode)) = 0 Then
                        '���걾�Ų�
                        blnQryWithSampleNO = True
                    Else
                        '�������ѯ
                        gstrSQL = "Select a.*,Decode(A.�Ա�,Null,0,'��',1,'Ů',2,0) As �Ա�A,to_char(c.��������,'yyyy-mm-dd') As ��������A From ����걾��¼ a,����ҽ����¼ b,������Ϣ c " & _
                            " Where a.ҽ��id=b.id(+) And b.����id=c.����id(+)" & _
                            " And a.����ʱ�� Between [1] And [2]" & _
                            " And a.����ID=[3] And a.��������=[6]"
                        Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "��ѯ�걾��¼", CDate(strStartDate), _
                            CDate(strEndDate), lngDeviceID, strSampleID, iType, strBarcode)
                        If Not rsTmp.EOF Then
                            blnQryWithSampleNO = False
                        Else
                            '�����Ƿ����б걾
                            gstrSQL = "Select a.*,Decode(A.�Ա�,Null,0,'��',1,'Ů',2,0) As �Ա�A,to_char(c.��������,'yyyy-mm-dd') As ��������A From ����걾��¼ a,����ҽ����¼ b,������Ϣ c " & _
                            " Where a.ҽ��id=b.id(+) And b.����id=c.����id(+)" & _
                            " And a.����ʱ�� Between [1] And [2]" & _
                            " And a.����ID=[3] And a.�걾���=[4] And Nvl(a.�걾���,0)=[5]"
                            Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "��ѯ�걾��¼", CDate(Format(strDate, "yyyy-MM-dd") & " 00:00:00"), _
                                CDate(Format(strDate, "yyyy-MM-dd") & " 23:59:59"), lngDeviceID, strSampleID, iType, strBarcode)
                            If rsTmp.EOF = True Then
                                '�����������ɱ걾
                                Call CreateSample(lngDeviceID, strBarcode, strSampleID, CDate(strDate), iType)
                                blnQryWithSampleNO = True
                            Else
                                If Val(Nvl(rsTmp("ҽ��id"))) = 0 Then
                                    '�걾Ϊ����ʱҲ����
                                    Call CreateSample(lngDeviceID, strBarcode, strSampleID, CDate(strDate), iType)
                                    blnQryWithSampleNO = True
                                End If
                            End If
                        End If
                    End If
                    
                    If blnQryWithSampleNO Then
                        gstrSQL = "Select a.*,Decode(A.�Ա�,Null,0,'��',1,'Ů',2,0) As �Ա�A,to_char(c.��������,'yyyy-mm-dd') As �������� From ����걾��¼ a,����ҽ����¼ b,������Ϣ c " & _
                            " Where a.ҽ��id=b.id(+) And b.����id=c.����id(+)" & _
                            " And a.����ʱ�� Between [1] And [2]" & _
                            " And a.����ID=[3] And a.�걾���=[4] And Nvl(a.�걾���,0)=[5] and a.�걾��� = [6] "
                        Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "��ѯ�걾��¼", CDate(strStartDate), _
                            CDate(strEndDate), lngDeviceID, strSampleID, iType, strSampleID)
                    End If
                    '-------------------------------------------------------------------------------
                    If rsTmp.EOF Then
                        '�����걾������ʱ�걾��¼
                        strSex = -1
                        strBirth = ""
                        lngID = gobjDatabase.GetNextId("����걾��¼")
                        gstrSQL = "ZL_����걾��¼_INSERT(" & lngID & ",NULL,'" & _
                            strSampleID & "',NULL,NULL," & lngDeviceID & ",NULL," & _
                            "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                            "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strSample & "'," & _
                            "Null,To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strName & "','" & aItem(4) & "'," & lngExeDeptID & "," & iType & "," & intMicrobe & ")"
                        ExecuteProcedure "���������ʱ��¼", gcnOracle
                    Else
                        strSex = Nvl(rsTmp("�Ա�A"), 0)
                        strBirth = Nvl(rsTmp("��������"))
                        If intMicrobe = 0 Then
                            strSample = Nvl(rsTmp("�걾����"))
                        End If
                        lngID = rsTmp("ID")
                        blnAuditing = Not IsNull(rsTmp("�����"))
                    End If
                        
                    If Not blnAuditing Then
                        '���������Ŀ
                        strItemRecords = ""
                        strδ֪�� = ""
                        For j = 5 To UBound(aItem) Step 2
                            '����ͨ�����޸���Ӧ��Ŀ�����δ�ҵ�����ֱ�����ӣ�����ͨ�����Ҳ�����Ŀ���ݲ�����
                            '����ͨ��������Ŀ
                            lngItemID = GetItemID(aItem(j))
                            If lngItemID > 0 Then
                                strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1)
                            Else
    
                                If strδ֪�� = "" Then strδ֪�� = "�걾��    ��Ŀ��ʶ    ��Ŀֵ" & vbNewLine
                                strδ֪�� = strδ֪�� & strSampleID & vbTab & aItem(j) & vbTab & aItem(j + 1) & vbNewLine
    '                            gcnAccess.Execute strSql
                            End If
                        Next
                        If strδ֪�� <> "" Then Call WriteLog("mdlLISComm.ResultFromFile", LOG_δ֪��, 0, strδ֪��)
                        If Len(strItemRecords) > 0 Then
                            strItemRecords = Mid(strItemRecords, 2)
                            
                            gstrSQL = "ZL_������ͨ���_BATCHUPDATE(" & lngID & "," & _
                                lngDeviceID & ",'" & strSample & "'," & strSex & "," & _
                                IIf(strBirth = "", "Null", "To_Date('" & strBirth & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                strItemRecords & "'," & intMicrobe & ")"
                            ExecuteProcedure "����������", gcnOracle
                        End If
                    End If
                    
                    If UBound(aTmp) > 0 Then
                        If Trim(aTmp(1)) <> "" Then
                            '����ͼ������
                            Call WriteLog("SaveImg", LOG_ͨѶ��־, 0, "��ʼʱ��:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
                            Call SaveImg(lngID, aTmp(1))
                            Call WriteLog("SaveImg", LOG_ͨѶ��־, 0, "����ʱ��:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
                        End If
                    End If
                    
                End If
            End If
        Next
    End If
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
'    If gcnAccess.State <> adStateClosed Then gcnAccess.CommitTrans
    Exit Sub
DBError:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    Call WriteLog("mdlLISComm.ResultFromFile", LOG_������־, Err.Number, Err.Description & vbCrLf & gstrSQL)
End Sub

Private Function GetItemID(ByVal strChannel As String) As Long
    Dim i As Integer
    For i = 0 To UBound(mItem, 2)
        If UCase(strChannel) = UCase(mItem(0, i)) Then Exit For
    Next
    If i > UBound(mItem, 2) Then
        GetItemID = -1
    Else
        GetItemID = CLng(mItem(1, i))
    End If
End Function

Public Function ValidateIP(ByVal strIP As String, Optional strErrInfo As String) As Boolean
    '���IP��ַ����ȷ�ԡ�
    
    Dim varIP As Variant
    Dim IPError As Integer
    Dim IPd As Integer
    Dim i As Integer
    
    varIP = Split(strIP, ".")
    If UBound(varIP) <> 3 Then
        IPError = 0
    Else
        For i = 0 To 3
            If Not IsNumeric(varIP(i)) Then
                IPError = 1
                Exit For
            Else
                IPd = CInt(varIP(i))
                If IPd < 0 Or IPd > 255 Then
                    IPError = 2
                    Exit For
                Else
                    IPError = -1
                End If
            End If
        Next i
    End If
    
    ValidateIP = True
    Select Case IPError
        Case -1
            If strIP <> "0.0.0.0" Then
                ValidateIP = False
                strErrInfo = ""
            Else
                strErrInfo = "IP������Ϊ0.0.0.0��"
            End If
        Case 0
            strErrInfo = "IP��ʽ���ԣ�ӦΪXXX.XXX.XXX.XXX������XXXΪ0-255�����֡�"
        Case 1
            strErrInfo = "IP��ַֻ��Ϊ0-255�����֡�"
        Case 2
            strErrInfo = "IP��ַ�ķ�Χֻ��Ϊ0-255֮�䡣"
    End Select
End Function

Public Function ValidatePort(ByVal strPort As String, Optional strErrInfo As String) As Boolean
    '���˿ںŵ���ȷ�ԡ�
    ValidatePort = True
    If Not IsNumeric(Trim(strPort)) Then
        strErrInfo = "�˿ں�ֻ��Ϊ1-65535�����֡�"
    Else
        If Val(Trim(strPort)) > 0 And Val(Trim(strPort)) <= 65535 Then
            ValidatePort = False
            strErrInfo = ""
        Else
            strErrInfo = "�˿ںŵķ�Χֻ����1-65535֮�䡣"
        End If
    End If
End Function

Private Sub WriteTxtLog(ByVal lng���� As String, ByVal str��Ŀ As String, ByVal str���� As String)
    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim blnClearData As Boolean
    
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    'If Val(GetSetting("ZLSOFT", "zlLisLog", "Test", 0)) = 0 Then Exit Sub
    
    blnClearData = gblnClearData
    
    '������־(����ʱ��,��������,�����,������Ϣ
    If str��Ŀ <> "" Or str���� <> "" Then
        
        If lng���� = LOG_������־ Then
            '������־
            strFileName = App.Path & "\zlLis������־_" & Format(date, "yyyyMMdd") & ".LOG"
            If pLast������־ = str��Ŀ & "|" & str���� Then
                Exit Sub
            Else
                pLast������־ = str��Ŀ & "|" & str����
            End If
        ElseIf lng���� = LOG_ͨѶ��־ Then
            'ͨѶ��־
            
            If blnClearData Then Exit Sub '���������־ѡ���д��־
            strFileName = App.Path & "\zlLisͨѶ��־_" & Format(date, "yyyyMMdd") & ".LOG"
            If pLastͨѶ��־ = str��Ŀ & "|" & str���� Then
                Exit Sub
            Else
                pLastͨѶ��־ = str��Ŀ & "|" & str����
            End If
        ElseIf lng���� = LOG_δ֪�� Then
            'δ֪��
            If blnClearData Then Exit Sub '���������־ѡ���д��־
            strFileName = App.Path & "\zlLisδ֪��Ŀ_" & Format(date, "yyyyMMdd") & ".LOG"
        End If
        
        If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
        Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
        
        
        strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
        objStream.WriteLine ("ʱ��:" & strDate & " �汾:" & App.major & "." & App.minor & "." & App.Revision)
        
        objStream.WriteLine (str��Ŀ)
        objStream.WriteLine (str����)
        
        'objStream.WriteLine (String(50, "-"))
        objStream.Close
        Set objStream = Nothing
    End If
End Sub

Public Sub SaveImg(ByVal lngID As Long, ByVal strImg As String)
    '����ͼ�����ݵ����ݿ���
    
    Dim aGraphItem() As String
    Dim strImageVal As String
    Dim strImageType As String
    Dim strImageData As String
    Dim intLoop As Integer
    Dim IntCount As Integer
    Dim blnDeleImg As Boolean '������Ƿ�ɾ��ԭ����ͼƬ
    Dim strPicPath As String, strSQL() As String
    Dim intLayOut As Integer 'ͼƬ����ʾ��ʽ
    Dim strBMPFile As String
    
    On Error GoTo ErrHandle
    aGraphItem = Split(strImg, "^")
        
    For intLoop = 0 To UBound(aGraphItem)
        strImageVal = Replace(aGraphItem(intLoop), vbCrLf, "")
        strImageType = Mid(strImageVal, 1, InStr(strImageVal, ";") - 1)
        strImageData = Mid(strImageVal, InStr(strImageVal, ";") + 1)
        
        If Mid(strImageData, 1, InStr(strImageData, ";") - 1) >= 100 And Mid(strImageData, 1, InStr(strImageData, ";") - 1) <= 227 Then
            '��֯ͼƬ����
            intLayOut = Mid(strImageData, 1, InStr(strImageData, ";") - 1)
            strPicPath = Mid(strImageData, InStr(strImageData, ";") + 1)
            
            If InStr(strPicPath, ";") > 0 Then
                strPicPath = Mid(strPicPath, InStr(strPicPath, ";") + 1)
                If Left(strPicPath, 2) = "1;" Then
                    blnDeleImg = True
                End If
            End If
            
            If Dir(strPicPath) <> "" Then
                If UCase(Right(strPicPath, 4)) = ".BMP" And intLayOut >= 100 And intLayOut <= 107 Then
                    strBMPFile = strPicPath
                ElseIf (UCase(Right(strPicPath, 4)) = ".JPG" Or UCase(Right(strPicPath, 4)) = ".GIF") And intLayOut >= 110 And intLayOut <= 127 Then
                    strBMPFile = strPicPath
                ElseIf intLayOut >= 200 And intLayOut <= 227 Then
                    strPicPath = UCase$(strPicPath)
                    strBMPFile = zlFileZip(strPicPath)
                Else
                    frmLISSrv.picTmp.Picture = LoadPicture(strPicPath)
                    If Dir(App.Path & "\zlLisIn.bmp") <> "" Then Kill App.Path & "\zlLisIn.bmp"
                    SavePicture frmLISSrv.picTmp.Picture, App.Path & "\zlLisIn.bmp"
                    strBMPFile = App.Path & "\zlLisIn.bmp"
                End If
                
                
                If zlLisBlobSql(lngID, strImageType, strBMPFile, intLayOut, strSQL) Then
                    WriteLog "ִ�� SaveImg", LOG_ͨѶ��־, 0, "��ʼʱ��"
                    For IntCount = LBound(strSQL) To UBound(strSQL)
                        If strSQL(IntCount) <> "" Then
                            gstrSQL = strSQL(IntCount)
                            ExecuteProcedure Replace(strSQL(IntCount), "Call", ""), gcnOracle
                        End If
                    Next
                    WriteLog "ִ�� SaveImg", LOG_ͨѶ��־, 0, "����ʱ��"
                End If
                If blnDeleImg Then
                    Kill strPicPath
                End If
                If intLayOut >= 200 And intLayOut <= 227 Then
                    Kill strBMPFile
                End If
            End If
        Else
            'ͼ������
            If Len(strImageData) > 2000 Then
                '�������2000��������

                For IntCount = 1 To CInt(Len(strImageData) / 1000) + 1
                    If Len(strImageData) > 0 Then
                        
                        gstrSQL = "Zl_����ͼ����_Update(" & lngID & ",'" & strImageType & "','" & _
                                                Mid(strImageData, IntCount * 1000 - 999, 1000) & "'," & _
                                                "1," & IntCount & ")"
                        ExecuteProcedure "����ͼ�񱣴�", gcnOracle
                    End If
                Next

            Else
                gstrSQL = "Zl_����ͼ����_Update(" & lngID & ",'" & strImageType & "','" & strImageData & "',0,1)"
                ExecuteProcedure "����ͼ�񱣴�", gcnOracle
            End If
        End If
    Next

    Exit Sub
ErrHandle:
    Call WriteLog("SaveImg", LOG_������־, Err.Number, Err.Description)

End Sub


Private Function zlLisBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, ByVal layOut As Integer, ByRef arySql() As String) As Boolean
    '���ɱ���ͼƬ��SQL
    'Action ����ID
    'KeyWord ����
    'strFile ͼƬ�ļ�
    'arySql ���ɵ�SQL����ڴ�������
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Dim lngLBound As Long, lngUBound As Long    '�����������С����±�
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo 0
    
    lngFileNum = FreeFile
    WriteLog "����BlobSQL", LOG_ͨѶ��־, 0, "��ʼʱ��"
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    conChunkSize = 512
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        If strText <> "" Then
            If lngCount = 0 Then strText = layOut & ";" & strText
            arySql(lngUBound + lngCount + 1) = "Zl_����ͼ����_Update(" & Action & ",'" & KeyWord & "','" & strText & "',1," & IIf(lngCount = 0, 1, 0) & ")"
        End If
    Next
    Close lngFileNum
    WriteLog "����BlobSQL", LOG_ͨѶ��־, 0, "����ʱ��"
    zlLisBlobSql = True
    Exit Function

errHand:
    Close lngFileNum
    zlLisBlobSql = False
End Function
Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1, Optional ByVal BeginDate As String) As String
    '-----------------------------------------------------------------------------------------
    '����:��ȡ����ʱ��
    '����:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    Dim dateNow As Date
    
    If BeginDate = "" Then
        dateNow = gobjDatabase.Currentdate
    Else
        dateNow = BeginDate
    End If
    
    Select Case strMode
    Case "��  ʱ"      '��ʱ
        GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����,bytFlag=1,���ܿ�ʼʱ��,=2,���ܽ���ʱ��
        intDay = Weekday(CDate(Format(dateNow, "YYYY-MM-DD")))
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 2, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 8 - intDay, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(dateNow, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"      '������
        Select Case Format(dateNow, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "������"      '������
        If Val(Format(dateNow, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "��  ��"   'ȫ��
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "���ظ�"
        If bytFlag = 1 Then
            GetDateTime = "2000-01-01 00:00:00"
        Else
            GetDateTime = "3000-12-31 23:59:59"
        End If
    End Select
    
End Function

Public Function CreateSample(ByVal lngDeviceID As Long, ByVal strBarcode As String, _
    ByRef strSampleNO As String, ByVal dtSampleDate As Date, ByVal intType As Integer) As Boolean
    'inttype=0
    Dim strSQL As String, rsTmp As adodb.Recordset, rs As New adodb.Recordset
    Dim lngKey As Long, strItemRecords As String
    Dim lngDeptID As Long '��ǰ��������
    Dim rsItem As New adodb.Recordset
    Dim strItem As String                           '������Ŀ
    Dim str���� As String, str�Ա� As String, str���� As String
    On Error GoTo DBErr
    
    CreateSample = False
    
    '������������
    strSQL = "Select ʹ��С��id From �������� Where ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��������걾", lngDeviceID)
    lngDeptID = lngExeDeptID
    If Not rsTmp.EOF Then
        lngDeptID = Nvl(rsTmp("ʹ��С��id"), lngExeDeptID)
    End If
    
    If Val(strSampleNO) <= 0 Then
        strSampleNO = Val(CalcNextCode(lngDeviceID, 0, intType))
    End If

    '���ҷ����������Ŀָ��
'    strSql = "Select A.���ID AS ID," & _
        "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����,A.�Ա�,A.����,F.No," & _
        "I.������ĿID As ��ĿID,Decode(I.�������,3,Nvl(I.Ĭ��ֵ,'-'),2,I.Ĭ��ֵ,'') As ���,'' As ��־," & _
        "Trim(REPLACE(REPLACE(' '||zlGetReference(I.������ĿID,A.�걾��λ,DECODE(A.�Ա�,'��',1,'Ů',2,0),C.��������,Y.����ID,A.����),' .','0.'),'��.','��0.')) AS ����ο�," & _
        "NVL(A.������־,0) AS ����,F.����ʱ��,F.������ " & _
        "FROM ����ҽ����¼ A," & _
        "������Ϣ C,����ҽ������ F,���鱨����Ŀ G,������Ŀ I,����������Ŀ Y " & _
        "WHERE A.������� = 'C' " & _
        "AND A.����ID=C.����ID " & _
        "AND A.���id IS NOT NULL " & _
        "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
        "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null AND G.������Ŀid=Y.��Ŀid(+) " & _
        "AND G.������ĿID=I.������ĿID " & _
        "AND (Y.����ID+0=[1] Or (Y.����ID Is Null And F.ִ�в���ID=[3])) " & _
        "And F.��������=[2] "
'        "AND F.ִ��״̬=0 "
    
    strSQL = "Select ID, ����, �Ա�, ����, NO, ��Ŀid, ���, ��־, ����ο�, ����, ����ʱ��, ������, Rownum As �������, ������Ŀid," & vbNewLine & _
            "       ����,�걾��λ,��������ID,����ҽ��,��ʶ��,��ǰ����,���˿��� " & vbNewLine & _
            "From (Select A.���id As ID, C.���� || Decode(A.Ӥ��, 0, '', Null, '', '(Ӥ��)') As ����, A.�Ա�, A.����, F.NO," & vbNewLine & _
            "              I.������Ŀid As ��Ŀid, Decode(I.�������, 3, Nvl(I.Ĭ��ֵ, '-'), 2, I.Ĭ��ֵ, '') As ���, '' As ��־," & vbNewLine & _
            "              Trim(Replace(Replace(' ' || Zlgetreference(I.������Ŀid, A.�걾��λ, Decode(A.�Ա�, '��', 1, 'Ů', 2, 0)," & vbNewLine & _
            "                                                          C.��������, Y.����id, A.����), ' .', '0.'), '��.', '��0.')) As ����ο�," & vbNewLine & _
            "              Nvl(A.������־, 0) As ����, F.����ʱ��, F.������, G.�������, A.������Ŀid, M.����, " & vbNewLine & _
            "              a.�걾��λ,��������ID,����ҽ��,decode(a.������Դ,2,c.סԺ��,c.�����) as ��ʶ��,c.��ǰ����,l.���� as ���˿��� " & vbNewLine & _
            "       From ����ҽ����¼ A, ������Ϣ C, ����ҽ������ F, ���鱨����Ŀ G, ������Ŀ I, ����������Ŀ Y, ������ĿĿ¼ M ,���ű� L " & vbNewLine & _
            "       Where A.������� = 'C' And A.����id = C.����id And A.���id Is Not Null And A.ҽ��״̬ = 8 And A.ID = F.ҽ��id And" & vbNewLine & _
            "             A.������Ŀid = G.������Ŀid And G.ϸ��id Is Null And G.������Ŀid = Y.��Ŀid(+) And" & vbNewLine & _
            "             G.������Ŀid = I.������Ŀid And A.������Ŀid = M.ID(+) And a.���˿���ID = l.ID" & vbNewLine & _
            "             and (Y.����id + 0 = [1] Or (Y.����id Is Null And F.ִ�в���id = [3])) And nvl(F.ִ��״̬,0) = 0  And F.�������� = [2]" & vbNewLine & _
            "       Order By M.����, G.�������)"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��������걾", lngDeviceID, strBarcode, lngDeptID)
    If rsTmp.EOF Then Exit Function
    
    gstrSQL = "Select B.����id, B.��ҳid, B.���, B.Ӥ������, B.Ӥ���Ա�" & vbNewLine & _
                    "From ����ҽ����¼ A, ������������¼ B" & vbNewLine & _
                    "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.Ӥ�� = B.��� And A.���id = [1] And Rownum = 1"
    Set rs = gobjDatabase.OpenSQLRecord(gstrSQL, "CreateSample", CLng(rsTmp("ID")))
    If rs.EOF = False Then
        str���� = Nvl(rs("Ӥ������"))
        str�Ա� = Nvl(rs("Ӥ���Ա�"))
        str���� = "Ӥ��"
    Else
        str���� = Nvl(rsTmp("����"))
        str�Ա� = Nvl(rsTmp("�Ա�"))
        str���� = Nvl(rsTmp("����"))
    End If
    
    '����������Ŀ
    gstrSQL = "select distinct ҽ������ from ����ҽ����¼ a , ����ҽ������ b, ���鱨����Ŀ c , ����������Ŀ d " & vbNewLine & _
              "  where a.id = b.ҽ��ID and a.���id is not null and a.������ĿID = c.������ĿID and " & vbNewLine & _
              "  c.������ĿID = d.��ĿID(+) and  (d.����id + 0 = [1] Or (d.����id Is Null And b.ִ�в���id = [3])) and b.�������� = [2] "
    Set rsItem = gobjDatabase.OpenSQLRecord(gstrSQL, "��������걾_1", lngDeviceID, strBarcode, lngDeptID)
    Do Until rsItem.EOF
        strItem = strItem & " " & Nvl(rsItem("ҽ������"))
        rsItem.MoveNext
    Loop
    strItem = Trim(strItem) & "(" & Nvl(rsTmp("�걾��λ")) & ")"
        
    '�����걾��¼
    lngKey = gobjDatabase.GetNextId("����걾��¼")
    gstrSQL = "ZL_����걾��¼_�걾����(" & lngKey & "," & _
        rsTmp("ID") & ",'" & _
        strSampleNO & "'," & _
        IIf(IsNull(rsTmp("����ʱ��")), "Null", "TO_DATE('" & Format(rsTmp("����ʱ��"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
        IIf(IsNull(rsTmp("������")), "Null", "'" & rsTmp("������") & "'") & "," & _
        lngDeviceID & "," & _
        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
        "1,'" & _
        gobjDatabase.GetUserInfo.Fields("����").Value & "'," & _
        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0,0,0," & _
        intType & ",NULL,'" & _
        str���� & "','" & str�Ա� & "','" & str���� & "','" & Nvl(rsTmp("No")) & "','" & _
        Nvl(rsTmp("�걾��λ")) & "'," & Nvl(rsTmp("��������ID")) & ",'" & Nvl(rsTmp("����ҽ��")) & "'," & _
        Nvl(rsTmp("��ʶ��")) & ",'" & Nvl(rsTmp("��ǰ����")) & "','" & Nvl(rsTmp("���˿���")) & "','" & _
        strItem & "')"
    ExecuteProcedure "��������걾", gcnOracle
                                                                
    '��дָ��
    strItemRecords = ""
    Do While Not rsTmp.EOF
        strItemRecords = strItemRecords & "|" & rsTmp("ID") & "^" & rsTmp("��ĿID") & "^" & _
            Nvl(rsTmp("���")) & "^" & Nvl(rsTmp("��־"), 0) & "^" & Nvl(rsTmp("����ο�")) & "^" & _
            Nvl(rsTmp("������ĿID")) & "^" & Nvl(rsTmp("�������"))
            
        rsTmp.MoveNext
    Loop
    
    If Len(strItemRecords) > 0 Then
        strItemRecords = Mid(strItemRecords, 2)
            
        gstrSQL = "Zl_������ͨ���_Write(" & lngKey & "," & _
            lngDeviceID & ",'" & strItemRecords & "',0,0)"
        ExecuteProcedure "��������걾", gcnOracle
    End If
    Exit Function
DBErr:
    Call WriteLog("clsLISComm.CreateSample", LOG_������־, Err.Number, Err.Description)
End Function

Private Function CalcNextCode(ByVal lngKey As Long, ByVal intRow As Integer, ByVal iType As Integer) As String
    '--------------------------------------------------------------------------------------------------------
    '����:����ָ�������ڵ����ڵ���һ��ȱʡ�걾��
    '����:lngKey                ��������ID
    '     iType                 �걾���0=��ͨ��1=����
    '����:ȱʡ�걾����
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New adodb.Recordset
    Dim strToday As String
    Dim strTmp As String
    Dim lng���� As Long
    Dim strLabNo As String, strLabQCNo As String '����걾���ʿر걾
    Dim mstrSQL As String, mlngLoop As Long
    Dim mlngDefaultItemID As Long
    
    'ʱ��,����,�걾��
    On Error GoTo errHand
    mlngDefaultItemID = 0
    strToday = Format(gobjDatabase.Currentdate, "YYYY-MM-DD")
    
    On Error GoTo point1
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(�걾���)),0) AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾id(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                        IIf(mlngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id= [1] ") & " And ҽ��ID Is Not Null" & _
                    IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1")
    Set rs = gobjDatabase.OpenSQLRecord(mstrSQL, "����", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                           CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabNo = gobjCommFun.Nvl(rs("������"))
    
    On Error GoTo errHand
    GoTo point2
    
point1:
    On Error GoTo errHand
    
    mstrSQL = "SELECT NVL(MAX(�걾���),'') AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾id(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id= [1] ") & " And ҽ��ID Is Not Null" & _
                    IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1")
    Set rs = gobjDatabase.OpenSQLRecord(mstrSQL, "����", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabNo = gobjCommFun.Nvl(rs("������"))
    
point2:
    On Error GoTo point3
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(�걾���)),0) AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾ID(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id= [1] ") & _
                    IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1")
    Set rs = gobjDatabase.OpenSQLRecord(mstrSQL, "����", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabQCNo = gobjCommFun.Nvl(rs("������"))
    
    On Error GoTo errHand
    GoTo point4
    
point3:
    On Error GoTo errHand
    
    mstrSQL = "SELECT NVL(MAX(�걾���),'') AS ������ FROM ����걾��¼ a,����������Ŀ b" & _
                " WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾ID(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id=[1] ") & _
                    IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1")
    Set rs = gobjDatabase.OpenSQLRecord(mstrSQL, "����", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabQCNo = gobjCommFun.Nvl(rs("������"))
    
point4:
    If strLabNo >= strLabQCNo Then
        CalcNextCode = strLabNo
    Else
        CalcNextCode = strLabQCNo
    End If
'    If Val(strLabQCNo) > Val(strLabNo) + 100 Then CalcNextCode = strLabNo

'    For mlngLoop = 1 To vsf2.Rows - 1
'        If mlngLoop <> intRow Then
'            If Val(vsf2.RowData(mlngLoop)) = lngKey Then
'                If Val(CalcNextCode) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
'                    CalcNextCode = Val(vsf2.TextMatrix(mlngLoop, 2))
'                End If
'            End If
'        End If
'    Next
'
    If Val(CalcNextCode) <= 0 Then
        CalcNextCode = "1"
        Exit Function
    End If
'
    CalcNextCode = Val(CalcNextCode) + 1
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function
'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String, Optional ByVal strUnZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If gobjFSO.FileExists(strUnZipFile) Then gobjFSO.DeleteFile strUnZipFile
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strUnZipFile) <> "" Then
        zlFileUnzip = strUnZipFile
    Else
        zlFileUnzip = ""
    End If
End Function
'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLLIS" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function
