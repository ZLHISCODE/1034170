Attribute VB_Name = "mdlPublic"
Option Explicit

'���ò����� {+}1405{+}ZLHIS[+]ZLHIS[+]HIS[+]0{+}false{+}false{+}0{+}0{+}false

Public gstrLogPath As String        '��־�ļ�
Public gstrImages As String         '��Ϣ���� strImages
Public glngOrderID As Long          '��Ϣ���� lngOrderID
Public gstrDBConnection As String   '��Ϣ���� strDBConnection
Public gblnMoved As Boolean         '��Ϣ���� blnMoved
Public gbAdd As Boolean             '��Ϣ���� bAdd
Public gintImageInterval As Integer '��Ϣ���� intImageInterval
Public glngSys As Long              '��Ϣ���� lngSys
Public gblnReconnectDB As Boolean   '��Ϣ���� blnReconnectDB
Public gstrZLHIS�����ַ��� As String '��Ϣ���� strDBServer
Public gstr�û��� As String          '��Ϣ���� strDBUser
Public gstr���� As String            '��Ϣ���� strDBPassword
Public gbln�Ƿ�ת������ As Boolean '��Ϣ���� blnTransPassword
Public gfrmViewImage As frmViewImage    '��Ϣѭ����������
Public gobjPacsCore As Object       '��Ƭ����
Public glngPreWndProc As Long       'ԭ������Ϣ�������
Public glngLog As Long              '�Ƿ��¼��־��0---����δ��ֵ��1---��¼��־��2---����¼��־

Public Const HIS_CAPTION = "����Ӱ���Ƭ����"
Public Const MSG_SPLIT = "{+}"

Private mobjRegister As Object                  '10.35.10֮���ע�����
Public glngModule As Long                       'ģ���
Public gblnBefore3510 As Boolean                '����10.35.10ǰ��汾��True=10.35.10֮ǰ�汾,��ʹ��zlRegister����ʼ��comlibʱ��ҪSetDbUser��RegCheck
Public gzlComLib As Object                      '�������ݿ⴦��ģ��zlComLib
Public gcnOracle As ADODB.Connection            '�������ݿ�����

Public Const gstrSysName As String = "Ӱ���Ƭ"

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Enum LogType
    ltError = 0
    ltDebug = 1
End Enum

Public Function errHandle(errSubName As String, errTitle As String, Optional errDesc As String = "") As Long
'------------------------------------------------
'���ܣ�������
'������ logSubName  --  ��������ĺ�����
'       logTitle   -- ��������
'       logDesc   --  ��������
'���أ�1-�������Resume��0-�����˳�
'------------------------------------------------
    
    errHandle = 0
    
    '��¼������־
    Call WriteCommLog("zlSoftViewImage,����--" & errSubName, errTitle & "���������= " & err.Number, errDesc & "����������=" & err.Description, ltError)
    
    '��ʾ����
    MsgBox errTitle & errDesc, vbOKOnly, "��Ƭ�ӿ�zlSoftViewImage���ִ���"
    
    '�������
    err.Clear
    
End Function

Public Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String, ByVal ltLogType As LogType)
'------------------------------------------------
'���ܣ���¼ͨѶ��־
'������ logSubName  --  ������־�ĺ�����
'       logTitle   -- ��־����
'       logDesc   --  ��־����
'       ltLogType --  ��־����
'���أ���
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    
    On Error GoTo err
    
    If glngLog = 0 Then
        glngLog = Val(GetSetting("ZLSOFT", "����ģ��\zl9PacsCore\zlSoftViewImage\", "Log", 2))
    End If
    
    'Log=1���ż�¼��־
    If glngLog <> 1 And ltLogType <> ltError Then Exit Sub
    
    strFileName = gstrLogPath & "\Interface" & Format(Date, "YYYY-MM-DD") & ".log"
    
    strLog = Now() & " ���⣺ " & logTitle & vbCrLf & "   ������ " & logSubName & vbCrLf & "   ��־���ݣ�" & logDesc & vbCrLf
    
    '������־���ӱ�ǣ�����鿴����
    If ltLogType = ltError Then
        strLog = "�������������" & strLog
    End If
    
    Open strFileName For Append As #1
    Print #1, strLog
    Close #1
    
    Exit Sub
err:
    Close #1
End Sub

Public Function GetLogDir() As String
'------------------------------------------------
'���ܣ���ȡ��־Ŀ¼�����Ŀ¼�����ڣ��򴴽�Ŀ¼
'��������
'���أ���־����Ŀ¼
'------------------------------------------------
    Dim strLogPath As String
    Dim strBackupPath As String
    
    On Error GoTo err
    
    strLogPath = Mid(App.Path, 1, InStr(5, App.Path, "\"))
    strLogPath = strLogPath & "Log\��־����\100_PACS��Ƭ��־"
    
    
    Call MkLocalDir(strLogPath + "\")
    
    GetLogDir = strLogPath
   
    Exit Function
err:
    GetLogDir = App.Path & "\100_PACS��Ƭ��־"
    Call MkLocalDir(GetLogDir + "\")
End Function

Public Function ProcessMessage(strMsg As String) As Long
'------------------------------------------------
'���ܣ�������յ�����Ϣ
'������strMsg -- ����exeʱ����Ĳ�����
'���أ���
'------------------------------------------------
    
    Dim lngPartType As Long
    Dim strDBUser As String
    Dim lngPatientID As Long
    Dim lngClinicID As Long
    Dim lngDeptID As Long
    Dim lngOrderID As Long
    
    On Error GoTo err
    ProcessMessage = 1
    
    '����Ĳ������壬���������ӷ��������ַ���{+}��
    '������ʽ��strImages{+}lngOrderID{+}strDBConnection{+}blnMoved{+}bAdd{+}intImageInterval{+}lngSys{+}blnReconnectDB
    '�������ͣ� strImages --- ͼ���,�����ǡ�����UID1|1-3;5-27;33-100+����UID2|ȫ����,ȫ����ʾ��ȫ��ͼ��
    '           lngOrderID --- ҽ��ID
    '           strDBConnection --- ���ݿ����Ӵ���������������[+]�û���[+]����[+]�����Ƿ�ת���������ӷ��������ַ���[+]��
    '                          �������롱���û���¼����ʱ���������Ƿ�ת����=1���������롱�����ݿ��¼����ʱ���������Ƿ�ת����=0
    '           blnMoved --- �����Ƿ�ת��
    '           bAdd --- ��ѡ������Ĭ��ֵFalse����ͼ�������ӽ���Ƭվ�������滻ԭ��Ƭվ����ͼ��TrueΪ���ӣ�FasleΪ�滻
    '           intImageInterval --- ��ѡ������Ĭ��ֵ0����ͼ��ļ����ֻ�Դ�ȫ������,��������ͼ������>100ʱ��Ч
    '           lngSys --- ��ѡ������Ĭ��,100��ϵͳ���
    '           blnReconnectDB --- ��ѡ������Ĭ��ֵFalse���Ƿ������������ݿ⡣��һ�δ򿪹�Ƭʱ�Զ��������ݿ⣬֮���ٴ򿪹�Ƭ��
    '                           ��blnReconnectDB���������Ƿ������������ݿ⡣
    '                           =True��ʹ��strDBConnection���������������ݿ⣻=False�����������������ݿ⣬ʹ�ù�Ƭ�������ڵ����ݿ�����
    '
    
    '�ȴ���̶�����
    If UBound(Split(strMsg, MSG_SPLIT)) >= 3 Then
        gstrImages = Split(strMsg, MSG_SPLIT)(0)
        glngOrderID = Val(Split(strMsg, MSG_SPLIT)(1))
        gstrDBConnection = Split(strMsg, MSG_SPLIT)(2)
        gblnMoved = (UCase(Split(strMsg, MSG_SPLIT)(3)) = "TRUE")
    Else
        Call WriteCommLog("����--zlSoftShowHisForms.ProcessMessage", "��������", "������������������������4��������Ϊ��" & strMsg, ltError)
        Exit Function
    End If
    
    '�ٴ����ѡ����
    If UBound(Split(strMsg, MSG_SPLIT)) >= 4 Then
        gbAdd = (UCase(Split(strMsg, MSG_SPLIT)(4)) = "TRUE")
    Else
        gbAdd = False
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) >= 5 Then
        gintImageInterval = Val(Split(strMsg, MSG_SPLIT)(5))
    Else
        gintImageInterval = 0
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) >= 6 Then
        glngSys = Val(Split(strMsg, MSG_SPLIT)(6))
    Else
        glngSys = 100
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) = 7 Then
        gblnReconnectDB = (UCase(Split(strMsg, MSG_SPLIT)(7)) = "TRUE")
    Else
        gblnReconnectDB = False
    End If
    
    If CreatePacsCore = False Then
        Exit Function
    End If
    
    Call WriteCommLog("zlSoftShowHisForms.ProcessMessage", "���ù�Ƭ", "��Ƭ�Ĳ����ǣ�gstrImages=" & gstrImages & ",glngOrderID=" & glngOrderID _
        & ",gstrDBConnection=" & gstrDBConnection & ",gblnMoved=" & gblnMoved & ",gbAdd=" & gbAdd & ",gintImageInterval=" & gintImageInterval _
        & ",glngSys=" & glngSys & ",gblnReconnectDB=" & gblnReconnectDB, ltDebug)
    
    Call gobjPacsCore.CallOpenViewer(gstrImages, glngOrderID, Nothing, gcnOracle, gblnMoved, gbAdd, gintImageInterval, glngSys)
    
    ProcessMessage = 0
    Exit Function
    
err:
    Call WriteCommLog("����--zlSoftShowHisForms.ProcessMessage", "������յ�����Ϣ�����ִ����յ�����Ϣ�ǣ�" & strMsg & "���������= " & err.Number, "����������=" & err.Description, ltError)
End Function

'******************************************************************************************************************
'���ܣ�����PACS��Ƭ����
'��������
'���أ������ɹ�,����true,���򷵻�False
'˵����
'******************************************************************************************************************
Private Function CreatePacsCore() As Boolean

    err = 0: On Error Resume Next
    If Not gobjPacsCore Is Nothing Then CreatePacsCore = True: Exit Function
    
    Set gobjPacsCore = CreateObject("zl9PacsCore.clsViewer")
    
    If err <> 0 Then
        MsgBox "δ�ҵ� zl9PacsCore �����������ǳ���汾��֧�֣������վ���Ƿ����˴˲���!", vbInformation + vbOKOnly, "��ʾ��Ϣ"
        Exit Function
    End If
    
    CreatePacsCore = True
    
End Function

Public Function CloseAllForms() As Boolean

    On Error GoTo err
    
    '�ر���Ϣѭ��������
    If Not gfrmViewImage Is Nothing Then
        Unload gfrmViewImage
        Set gfrmViewImage = Nothing
    End If
    
    CloseAllForms = True
    
    Exit Function
err:
    Call WriteCommLog("����--zlSoftViewImage.CloseAllForms", "�˳����򣬹ر����д��ڣ����ִ��󣬴������= " & err.Number, "����������=" & err.Description, ltError)
    Resume Next
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function InitInterface(ByVal strDBUser As String) As Boolean
'------------------------------------------------
'���ܣ���ʼ���ӿڣ�����ComLib���������ݿ�
'��������
'���أ�True-�ɹ���False-ʧ��
'------------------------------------------------
    
    On Error GoTo err
    InitInterface = False
    
    '��ʼ��ϵͳ��Ϊ100��ģ���Ϊ1289
    glngSys = 100
    glngModule = 1289
        
On Error Resume Next
    If mobjRegister Is Nothing Then
        Set mobjRegister = GetObject("", "zlRegister.clsRegister")
        If mobjRegister Is Nothing Then gblnBefore3510 = True '35.10֮ǰ�İ汾
    End If
    
    err.Clear
On Error GoTo err
    If gzlComLib Is Nothing Then
        If gblnBefore3510 Then
            '10.35.10֮ǰ�İ汾
            Set gzlComLib = CreateObject("zl9ComLib.clsComLib")
        Else
            '10.35.10֮��İ汾
            Set gzlComLib = GetObject("", "zl9ComLib.clsComLib")
        End If
    End If
    
    '����Ǵ�RIS������DLL�����ݿ�����gzlComLib.CurrentConn�ǿյģ���Ҫ��ע����ȡ�û������룬�����������ݿ�
    If gzlComLib.CurrentConn Is Nothing Then
        '��ע����ȡ�û������룬�������ݿ�
        
        '���gcnOracle�����ڣ�Ҫ�½�һ��
        If gcnOracle Is Nothing Or gblnReconnectDB = True Then
            Set gcnOracle = New ADODB.Connection
            Call ConnectDB(strDBUser)
        End If

        '��ʼ����������
        gzlComLib.InitCommon gcnOracle
        
        If gblnBefore3510 = True Then
            '10.35.10֮ǰ�İ汾
            If gzlComLib.RegCheck = False Then
                
                Exit Function
            End If
        End If
    Else
        '����Ǵ�HIS����̨������DLL���򴴽�zl9ComLib֮�󣬻��Զ�������gzlComLib.CurrentConn
        '������ʱû�д� CodeMan��ȡ�� gcnOracle��������Ҫ��zl9ComLibȡ��gcnOracle����
        
        If gcnOracle Is Nothing Then Set gcnOracle = gzlComLib.CurrentConn
    End If
    
    InitInterface = True
    
  
    Exit Function
err:
    If errHandle("zlSoftShowHisForms.InitInterface", "��ʼ���ӿڳ���", err.Description) = 1 Then Resume
End Function

Public Function ConnectDB(ByVal strDBUser As String) As Boolean
'------------------------------------------------
'���ܣ��������ݿ⣬��ע����ж�ȡ���ܺ�����ݿ�������Ϣ���û��������룬������
'������
'���أ�True-�ɹ���False-ʧ��
'------------------------------------------------
    Dim strDBPassword As String
    Dim strDBServer As String
    Dim blnTransPassword As Boolean
    
    ConnectDB = False
    
    On Error GoTo err
    
    If gcnOracle.State <> adStateOpen Then
        strDBServer = gstrZLHIS�����ַ���
        strDBUser = gstr�û���
        strDBPassword = gstr����
        blnTransPassword = gbln�Ƿ�ת������
                
        '�������ݿ�
        If OraDataOpen(strDBServer, strDBUser, strDBPassword, blnTransPassword) = False Then
           
            Exit Function
        End If
    End If
    
    ConnectDB = True
    Exit Function
err:
    If errHandle("zlSoftViewImage.ConnectDB", "�������ݿ⺯�����ִ���", err.Description) = 1 Then Resume
End Function

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, ByVal blnTransPassword As Boolean) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '   blnTransPassword �� �Ƿ���Ҫת������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strError As String
    
    On Error GoTo ErrHand
    

    If gblnBefore3510 = True Then
        '�����10.35.10֮ǰ�İ汾��ֱ�����û����������¼���ݿ�
        OraDataOpen = OpenOracle(gcnOracle, strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strUserPwd, IIf(blnTransPassword = True, TranPasswd(strUserPwd), strUserPwd)))
    Else
        '�����10.35.10֮��İ汾��ʹ��zlRegister��ȡ���ݿ�����
        Set gcnOracle = mobjRegister.GetConnection(strServerName, strUserName, strUserPwd, blnTransPassword, , strError, True)
        If gcnOracle.State = adStateOpen Then
            OraDataOpen = True
        Else
            OraDataOpen = False
        End If
    End If
    
    If OraDataOpen = True Then
        strUserName = UCase(strUserName) '����ΪʲôҪǿ�ƴ�д���ǲ���comlib��Ҫ��
        If gblnBefore3510 = True Then
            '10.35.10֮ǰ�İ汾
            gzlComLib.SetDbUser strUserName
        End If
    End If
    
    Exit Function
    
ErrHand:
    
    If errHandle("zlSoftViewImage.OraDataOpen", "�������ݿ����", err.Description) = 1 Then Resume
    OraDataOpen = False
End Function

Private Function OpenOracle(ByRef cnOrcle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ����Oracle���ݿ�
    '������
    '   cnOrcle �����ݿ�����
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    err = 0
    DoEvents
    With cnOrcle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '���������Ϣ
            strError = err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OpenOracle = False
            Exit Function
        End If
    End With
    
    OpenOracle = True
    err = 0
    
    Exit Function
    
End Function

Private Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    
    strNew = ""
    
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    
    TranPasswd = strNew
End Function

