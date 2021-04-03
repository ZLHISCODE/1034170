Attribute VB_Name = "mdlHisCrust"
Option Explicit
'ȫ�ֱ�������
Public gobjMe               As clsHisCrsutCom
Public gstrAppPath          As String                   '�Զ���������·��
Public gstrSetupPath        As String                   '����İ�װ·��
Public garrKillProcess      As Variant                  '������ɱ���Ľ�������
Public gstrPreTempPath      As String                   'ϵͳĿ¼System32Ŀ¼
Public gstrSystemPath       As String                   'ϵͳĿ¼System32Ŀ¼
Public gstrTempPath         As String                   '��ʱ���Ŀ¼
Public grsFileUpgrade       As adodb.Recordset          '�����ļ��嵥
Public gcnOracle            As adodb.Connection
Public gstrComputerName     As String                   '��������
Public gstrComputerIp       As String                   '������IP��ַ

Public gobjFSO              As New FileSystemObject     '�ļ���������
Public gobjTrace            As New clsTrace             '��־���ٶ���
Public gcllSetPath          As New Collection           '���а�װ·��
Public gclsRegCom           As New clsRegCom            '����ע�����
Public grsErrRec            As adodb.Recordset          '�����¼
Public gclsConnect          As clsConnect               '�ļ��ռ�������
Public gobj7zZip            As New cls7zZip             '7zѹ����

Public glngNoteLength       As Long                     '˵���ֶγ���
Public glngFileBatch        As Long                     '�����ļ�����
Private mblnWriteRunErrLog  As Boolean                  '�Ƿ���д���д�����־�������ݿ��������
Public gblnReCheckComs      As Boolean                  '�Ƿ����¼�鰲װ����
Public gintWaite            As Integer                  '�ȴ�������ʱ�䡣0-����������<>0�ȴ�N���Ӻ�ʼ������һ��Ϊ1
Public gblnIs64Bits         As Boolean                  '�Ƿ���64λϵͳ
Public gblnHaveVersion      As Boolean                  '�Ƿ�����ļ��汾���ֶ�
Public gblnSameFTP          As Boolean                  '�Ƿ�ʹ�ü���FTP����
'�����н�������
Public gstrCommand          As String                   '�Զ�������������
Public gstrConnectString    As String                   '�����ַ���
Public gotCurType           As OperateType              '����ִ�еĲ�������
Public gstrHisInput         As String                   'ZLHIS������û�������,��ʽΪUSER=ZLHIS PASS=HIS SERVER=TXYY(�������������)
Public gstrHisCommand       As String                   'ԭʼ��ZLHIS����
Public gstrAppEXE           As String                   '���ñ���ǳ�����ļ�
Public gintCallTimes        As Integer                  '���ô���
Public gblnAutoLogin        As Boolean                  '�Ƿ��Զ���¼
Public gstrTerminal         As String                   '��ǰ������
Public gstrAudsid           As String                   '��ǰaudsid
Public gblnSilence          As Boolean

Public gobjRegistry         As New clsRegistry

Public Function InitCommon() As Boolean
    On Error GoTo errH
    gblnAutoLogin = True
    gblnIs64Bits = Is64bit
    gstrSetupPath = GetSetupPath
    'Call gobjTrace.OpenTace("ZLHISCRUST", gstrSetupPath)
    gobjTrace.WriteSection "�ͻ����Զ�����"
    gobjTrace.WriteSection "������ʼ��", SL_LevelTwo
    gstrCommand = GetCommand()
    If gstrCommand = "" Then GoTo ReCall
    gstrTerminal = InitTerminal(gstrAudsid)
    If Not GetBaseInfo Then GoTo ReCall
    '�������
    If Not CheckJobs Then
        GoTo ReCall
    ElseIf gclsConnect Is Nothing Then                   'û�������Զ��˳�����¼ZLHIS
        GoTo AutoLogin
    End If

    Call EnablePrivilege(GetCurrentProcess(), SE_DEBUG_NAME)
    If Not SetOperateProcess(gotCurType, OS_InProcessing, SumErrMsg) Then GoTo ReCall
    '��װ·���޸�
    If Not CheckAndAdjustFolder() Then
        Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:����ʧ�� ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        Call SetOperateProcess(gotCurType, OS_Failure, SumErrMsg)          '��ʶ��������
        GoTo ReCall
    End If
    If gotCurType <> OT_CheckFile Then
        'ʣ��ռ���
        If Not CheckFreeSpace Then
            Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:����ʧ�� ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
            Call SetOperateProcess(gotCurType, OS_Failure, SumErrMsg)          '��ʶ��������
            GoTo ReCall
        End If
    End If
    InitCommon = True
    Exit Function
errH:
    MsgBox Err.Description, vbInformation, App.Title
    Err.Clear
ReCall:
    Call CallHISEXE
    Call gobjMe.ExitApp
AutoLogin:
    Call CallHISEXE(True)
    Call gobjMe.ExitApp
End Function

Private Function GetSetupPath() As String
'���ܣ���ȡ����İ�װ·��
    If IsDesinMode Then
        GetSetupPath = "C:\APPSOFT"
    Else
        '������ǰ����Apply���������ڿ��ܱ�ɱ���������������ٴδ����ʧ��
        '�������ZLuptmp����Ŀ¼����Ŀ¼Ϊ����·�����+ʱ�䣬��ֹ��ɱ��
        '��2016-12-12 12:12Ŀ¼ΪAPPSost\ZLUpTmp\1612121212
        '��ǰZLHISCrust.EXE����APPLY,�·�ʽ������APPSOFT\ZLUPTMP,��ѹͬʱҲ���ڴ˴�APPSOFT\ZLUPTMP
        If InStr(UCase(gstrAppPath), "\ZLUPTMP") > 0 Then
            GetSetupPath = gobjFSO.GetParentFolderName(gobjFSO.GetParentFolderName(gstrAppPath))
        ElseIf InStr(UCase(gstrAppPath), "\APPLY") > 0 Then
            GetSetupPath = gobjFSO.GetParentFolderName(gstrAppPath)
        Else
            GetSetupPath = gstrAppPath
        End If
    End If
End Function

Private Function GetCommand() As String
    Dim strCommand      As String, strServer        As String
    Dim objText         As TextStream
    Dim strErrInfo      As String
    
    On Error GoTo errH
    gobjTrace.WriteSection "��ȡ����", SL_LevelThree
    strCommand = gstrCommand
    gobjTrace.WriteInfo "GetCommand", "ԭʼ����������", Cipher(strCommand)
    'ZLRunAS.exe����û��������,ͨ�������ļ�����ԭʼ������
    If strCommand = "" Then
        If gobjFSO.FileExists(gstrSetupPath & "\ZLRUNAS.ini") Then
            'ZLRunAS.exe����û��������
            Set objText = gobjFSO.OpenTextFile(gstrSetupPath & "\ZLRUNAS.ini", ForReading)
            If Not objText.AtEndOfLine Then
                strCommand = objText.ReadLine
            End If
            objText.Close
            Set objText = Nothing
            Call gobjFSO.DeleteFile(gstrSetupPath & "\ZLRUNAS.ini", True)
            gobjTrace.WriteInfo "GetCommand", "ZLRUNAS����������", strCommand
            strCommand = DeCipher(strCommand)
        End If
    End If
    'ͨ�������ļ����ɼ��ܴ�
    If strCommand = "" Then
        If gobjFSO.FileExists(gstrSetupPath & "\ZLHISCRUST.ini") Then
            Set objText = gobjFSO.OpenTextFile(gstrSetupPath & "\ZLHISCRUST.ini", ForReading)
            If Not objText.AtEndOfLine Then
                strCommand = Trim(objText.ReadLine)
            End If
            objText.Close
            Set objText = Nothing
            Call gobjFSO.DeleteFile(gstrSetupPath & "\ZLHISCRUST.ini", True)
            If strCommand Like "ZLUPDATE:*" Then
            Else
                strCommand = "ZLUPDATE:" & Cipher(strCommand)
            End If
            gobjTrace.WriteInfo "GetCommand", "��������������", strCommand
        End If
    End If
    'û��������
    If strCommand = "" Then
        If IsDesinMode Then
'            strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=TESTBASE"";Persist Security Info=True;User ID=ZLHIS;Password=HIS;Data Provider=MSDASQL||0||OfficialUpgrade||||USER=ZLHIS PASS=aqa||CMDCHECK:1,171,174,191,193,214"
            strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=TESTBASE_LF"";Persist Security Info=True;User ID=ZLHIS;Password=HIS;Data Provider=MSDASQL||0||OfficialUpgrade||||USER=ZLHIS PASS=aqa||"
            gobjTrace.WriteInfo "GetCommand", "Դ������������", strCommand
        End If
    End If
    If strCommand = "" Then
        strServer = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="SERVER", Default:="")
        If MsgBox("�Ƿ���Ҫǿ��������", vbInformation + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
            Exit Function
        End If
        If strServer = "" Then strServer = InputBox("���������ӵķ�����", "��ʾ")
        If strServer = "" Then Exit Function
        'ʹ��ZLTOOLS/ZLTOOLS��¼
        strCommand = "ZLUPDATE:" & Cipher("USER=ZLTOOLS PASS=ZLTOOLS SERVER=" & strServer & " MODE=0")
        gobjTrace.WriteInfo "GetCommand", "ǿ������(1)������", strCommand
        Set gcnOracle = GetConnByCommand(strCommand)
        'ʹ��ZLTOOLS/ZLSOFT��¼
        If gcnOracle Is Nothing Then
            strCommand = "ZLUPDATE:" & Cipher("USER=ZLTOOLS PASS=ZLSOFT SERVER=" & strServer & " MODE=0")
            gobjTrace.WriteInfo "GetCommand", "ǿ������(2)������", strCommand
            Set gcnOracle = GetConnByCommand(strCommand)
        End If
        '�û����������¼
        If gcnOracle Is Nothing Then
            strCommand = InputBox("������ZLTOOLS������", "��ʾ")
            If strCommand = "" Then Exit Function
            strCommand = "ZLUPDATE:" & Cipher("USER=ZLTOOLS PASS=" & strCommand & " SERVER=" & strServer & " MODE=0")
            gobjTrace.WriteInfo "GetCommand", "ǿ������(3)������", strCommand
            Set gcnOracle = GetConnByCommand(strCommand, True)
        End If
    Else
        gobjTrace.WriteInfo "GetCommand", "��������������", Cipher(strCommand)
        Set gcnOracle = GetConnByCommand(strCommand, True)
    End If
    If gcnOracle Is Nothing Then Exit Function
    GetCommand = strCommand
    Exit Function
errH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "GetCommand", "��ȡ������ʧ��", strErrInfo
    MsgBox "��ȡ��������Ϣ����������������ϵ����Ա����Ϣ��" & vbNewLine & strErrInfo, vbInformation, App.Title
    Err.Clear
End Function

Private Function GetConnByCommand(ByVal strCommand As String, Optional ByVal blnMsg As Boolean) As adodb.Connection
'���ܣ�ͨ�������л�ȡ����
'       strCommand=������
'       blnMsg=�Ƿ���ʾ
'���أ�����������
    Dim strUser     As String, strPwd       As String, strServer    As String, intMode      As Integer
    Dim strTmp      As String, arrCommand   As Variant, i           As Integer
    Dim cnTmp       As adodb.Connection
    Dim strCur      As String, lngWait      As Long
    
    On Error GoTo errH
    gstrHisInput = "": gstrHisCommand = "": gstrAppEXE = "": gintCallTimes = 0: gstrConnectString = "": gintWaite = 0
    If strCommand Like "ZLUPDATEEX:*" Then
        gobjTrace.WriteInfo "GetConnByCommand", "��������", "���ηǳ�������"
        strCommand = DeCipher(Mid(strCommand, Len("ZLUPDATEEX:*")))
        gintCallTimes = 1
    End If
    
    'ʹ��ZLHIS���ù����˻�����
    If strCommand Like "ZLUPDATE:*" Then
        arrCommand = Split(DeCipher(Mid(strCommand, Len("ZLUPDATE:*"))), " ")
        For i = LBound(arrCommand) To UBound(arrCommand)
            If arrCommand(i) Like "USER=*" Then
                strUser = Mid(arrCommand(i), Len("USER=*"))
            ElseIf arrCommand(i) Like "PASS=*" Then
                strPwd = Mid(arrCommand(i), Len("PASS=*"))
            ElseIf arrCommand(i) Like "SERVER=*" Then
                strServer = Mid(arrCommand(i), Len("SERVER=*"))
            ElseIf arrCommand(i) Like "MODE=*" Then
                intMode = Val(Mid(arrCommand(i), Len("MODE=*")))
            End If
        Next
        gblnAutoLogin = False
        If strServer = "" Then strServer = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="SERVER", Default:="")
        If strUser = "" Or strPwd = "" Or strServer = "" Then
            gobjTrace.WriteInfo "GetConnByCommand", "����ʧ��", "��������Ϣ������������ϵ����!"
            If blnMsg Then
                MsgBox "��������Ϣ������������ϵ����Ա��", vbInformation, App.Title
            End If
            Exit Function
        End If
        gstrConnectString = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & strServer & """;Persist Security Info=True;User ID=" & strUser & ";Password=" & strPwd & ";Data Provider=MSDASQL"
        '��������
        gotCurType = Decode(intMode, 0, OT_Repair, 1, OT_OfficialUpgrade, 2, OT_PreUpgrade, 3, OT_CheckFile, OT_OfficialUpgrade)
    Else
        If strCommand Like "ZLUPDATENEW:*" Then
            gobjTrace.WriteInfo "GetConnByCommand", "��������", "���γ�������"
            strCommand = DeCipher(Mid(strCommand, Len("ZLUPDATENEW:*")))
            gintCallTimes = 1
        End If
        arrCommand = Split(strCommand, "||")
        'У�鷽ʽ����������������������н�����׼ȷ��
        If arrCommand(UBound(arrCommand)) Like "CMDCHECK:" Then
            strTmp = Mid(arrCommand(UBound(arrCommand)), 10)
            arrCommand = Split(strTmp, ",")
            strTmp = Mid(strCommand, 1, Len(strCommand) - Len(arrCommand(UBound(arrCommand))) - 2)
            For i = UBound(arrCommand) To LBound(arrCommand) Step -1
                If i = 5 Then
                    strCur = Mid(strTmp, Val(arrCommand(i)) + 2)
                    If strCur Like "W:*" Then '������ǰ�Ϸ�ʽ�Ĳ��Դ�����For+Sleep����ʵ�ֵȴ����÷������ڳ���������⣬�������ǰ׺W:
                        gintWaite = Val(Mid(strCur, 3))
                    End If
                ElseIf i = 4 Then               'HIS������û���������
                    gstrHisInput = Mid(strTmp, Val(arrCommand(i)) + 2)
                ElseIf i = 3 Then
                    gstrHisCommand = Mid(strTmp, Val(arrCommand(i)) + 2)
                ElseIf i = 2 Then
                    gstrAppEXE = Mid(strTmp, Val(arrCommand(i)) + 2)
                ElseIf i = 1 Then
                    If gintCallTimes = 0 Then gintCallTimes = Val(Mid(strTmp, Val(arrCommand(i)) + 2))
                ElseIf i = 0 Then
                    gstrConnectString = strTmp
                End If
                strTmp = Mid(strTmp, 1, Val(arrCommand(i)) - 1)
            Next
        Else
            gstrConnectString = arrCommand(0)
            If gintCallTimes = 0 Then gintCallTimes = Val(arrCommand(1))
            gstrAppEXE = arrCommand(2)          '"PreUpgrade","OfficialUpgrade","zlActMain.exe"
            If UBound(arrCommand) >= 3 Then
                gstrHisCommand = arrCommand(3)
                If UBound(arrCommand) >= 4 Then
                    gstrHisInput = arrCommand(4)
                    If UBound(arrCommand) >= 5 Then
                        If arrCommand(5) Like "W:*" Then '������ǰ�Ϸ�ʽ�Ĳ��Դ�������For+Sleep����ʵ�ֵȴ����÷������ڳ���������⣬�������ǰ׺W:
                            gintWaite = Val(Mid(arrCommand(5), 3))
                        End If
                    End If
                End If
            End If
        End If
        gotCurType = Decode(gstrAppEXE, "Repair", OT_Repair, "PreUpgrade", OT_PreUpgrade, "OfficialUpgrade", OT_OfficialUpgrade, "CheckUpgrade", OT_CheckFile, OT_OfficialUpgrade)
    End If
    gblnSilence = gotCurType = OT_CheckFile Or gotCurType = OT_PreUpgrade
    If gintWaite > 0 And gintCallTimes = 0 Then 'ֻ�е�һ�ε��òų�˯
        lngWait = gintWaite * 60000
        Call Sleep(lngWait)
    End If
    Err.Clear: On Error Resume Next
    Set cnTmp = New adodb.Connection
    cnTmp.CursorLocation = adUseClient
    cnTmp.ConnectionString = gstrConnectString
    cnTmp.Open
    If Err.Number <> 0 Then
        gobjTrace.WriteInfo "GetConnByCommand", "����ʧ��", Err.Description
        If blnMsg And Not gblnSilence Then
            MsgBox "�����ݿ�����ʧ�ܣ�����ϵ����Ա����Ϣ��" & vbNewLine & Err.Description, vbInformation, App.Title
        End If
        Err.Clear
        Exit Function
    End If
    gobjTrace.WriteInfo "GetConnByCommand", "����", Decode(gotCurType, OT_Repair, "�޸�", OT_PreUpgrade, "Ԥ����", OT_OfficialUpgrade, "��ʽ����", OT_CheckFile, "�ռ�"), _
                "�����ó���", gstrAppEXE, "�����ó�������", Cipher(gstrHisCommand), "������������", Cipher(gstrHisInput), "���ҵ��ô���", gintCallTimes
    Set GetConnByCommand = cnTmp
    Exit Function
errH:
    gobjTrace.WriteInfo "GetConnByCommand", "������ȡ����ʧ��", Err.Description
    If Not gblnSilence Then MsgBox "������ȡ����ʧ�ܣ�����ϵ����Ա����Ϣ��" & vbNewLine & Err.Description, vbInformation, App.Title
    Err.Clear
End Function

Public Sub CallHISEXE(Optional blnAutoLogin As Boolean)
    '����HIS
    Dim mError              As String
    Dim strFile             As String
    Dim strCommand          As String
    Dim strUserName         As String, strPass      As String, strServer As String
    Dim lngRet              As Long
    
    '�����ZLBH�ں����������ٻص�
    If UCase(gstrAppEXE) = "ZLACTMAIN.EXE" Or gotCurType = OT_PreUpgrade Or gotCurType = OT_CheckFile Then Exit Sub
    'ȷ���ļ��Ƿ����
    '1�����ٴ��� "ZLHIS90.exe"
    '2��Ԥ����Ҳ���Զ����õ���̨����
    If gstrAppEXE <> "" Then
        strFile = gstrSetupPath & "\" & gstrAppEXE
        If Not gobjFSO.FileExists(strFile) Then
            If UCase(gstrAppEXE) <> "ZLHIS+.EXE" Then
                strFile = gstrSetupPath & "\ZLHIS+.exe"
            End If
        End If
    Else
        strFile = gstrSetupPath & "\ZLHIS+.exe"
    End If
    gobjTrace.WriteInfo "CallHISEXE", "��������·��", strFile
    On Error Resume Next
    If blnAutoLogin And gblnAutoLogin Then
        '�����˻��������Զ���¼
        If gstrHisCommand = "" And gstrHisInput = "" And Not (gstrCommand Like "ZLUPDATE:*" Or gstrCommand Like "ZLUPDATEEX:*") Then
            If GetConnectionInfo(gstrConnectString, strServer, strUserName, strPass) Then
                strCommand = strUserName & "/" & strPass & "@" & strServer
            End If
        ElseIf gstrHisCommand <> "" Then
            strCommand = gstrHisCommand
            If gstrHisCommand Like "USER=*" Then
                strCommand = gstrHisCommand & " ZLHISCRUSTCALL=1"
            End If
        Else
            strCommand = gstrHisInput & IIf(gstrHisInput <> "", " ZLHISCRUSTCALL=1", "")
        End If
        gobjTrace.WriteInfo "CallHISEXE", "������", Cipher(strCommand)
        strCommand = strFile & " " & strCommand
    Else
        strCommand = strFile
    End If
    
    lngRet = Shell(strCommand, vbNormalFocus)
    Call Sleep(100)
End Sub

Public Function GetConnectionInfo(ByVal strConect As String, ByRef strServerName As String, ByRef strUserName As String, ByRef strPassword As String) As Boolean
'���ܣ� ����MSODBC���Ӷ����е�ORACLE���е� ���������û���������
'���أ� �ɹ�ʧ�ܣ�����True��ʧ�ܣ�����False

    Dim i As Integer
    Dim strTemp As String
    If strConect = "" Then Exit Function
            
    strServerName = ""
    strUserName = ""
    strPassword = ""
    strConect = Replace(strConect, """", "")
    
    If InStr(strConect, "ODBC") > 0 Then
        'Provider=MSDataShape.1;Extended Properties="Driver={Microsoft ODBC for Oracle};Server=DYYY";Persist Security Info=True;User ID=zlhis;Password=his;Data Provider=MSDASQL"
        'Provider=MSDataShape.1;Persist Security Info=False;User ID=ZLHIS;Data Provider=MSDASQL;
        '��ȡ strServerName(SecurityΪfalseʱ���޷����)
        i = InStrRev(strConect, "Server=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 6)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            End If
        End If
    Else
        'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
        'Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=ZLHIS;Data Source="DYYY"
        i = InStrRev(strConect, "Data Source=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 11)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            Else    'SecurityΪfalseʱ��û��;��
                strServerName = strTemp
            End If
        End If
    End If
    
    '��ȡ strUserName
    i = InStrRev(strConect, "User ID=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 7)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strUserName = Left(strTemp, i - 1)
        End If
    End If
    
    '��ȡ strPassword
    i = InStrRev(strConect, "Password=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 8)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strPassword = Left(strTemp, i - 1)
        End If
    End If
    GetConnectionInfo = strPassword <> "" And strUserName <> "" And strServerName <> ""
End Function

Private Function GetBaseInfo() As Boolean
    Dim strErrInfo      As String
    
    On Error GoTo errH
    gstrComputerName = ComputerName
    gstrComputerIp = IP
    gstrSystemPath = gobjFSO.GetSpecialFolder(SystemFolder)
    If gblnIs64Bits Then '64ϵͳ��32λ����Ӧ�÷���C:\windows\SysWOW64
        gstrSystemPath = gobjFSO.GetParentFolderName(gstrSystemPath) & "\SysWOW64"
    End If
    gblnReCheckComs = False
    gstrTempPath = gstrSetupPath & "\ZLUPTMP"
    If Not gobjFSO.FolderExists(gstrTempPath) Then
        Call gobjFSO.CreateFolder(gstrTempPath)
    End If
    gstrPreTempPath = gstrTempPath & "\ZLPRETMP"
    If Not gobjFSO.FolderExists(gstrPreTempPath) Then
        Call gobjFSO.CreateFolder(gstrPreTempPath)
    End If
    '��ʱĿ¼���붯̬Ŀ¼
    gstrTempPath = gstrTempPath & "\" & Format(Now, "YYMMDDHHmmss")
    If Not gobjFSO.FolderExists(gstrTempPath) Then
        Call gobjFSO.CreateFolder(gstrTempPath)
    End If
    mblnWriteRunErrLog = IsWriteRunErrLog()
    glngNoteLength = GetNoteLength
    gblnHaveVersion = IsHaveVersion()
    gblnSameFTP = IsSampleFTP()
    Set grsErrRec = CopyNewRec(Nothing, True, , Array("����", adInteger, 3, 0, "����", adVarChar, 100, Empty, "��Ϣ", adVarChar, 1000, Empty))
    Call RecordErrMsg(MT_MsgHeader, "��Ϣͷ", "����:" & Decode(gotCurType, OT_OfficialUpgrade, "����", OT_PreUpgrade, "Ԥ��", OT_Repair, "�޸�", OT_CheckFile, "�ռ�") & " ��ʼ:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
    gobjTrace.WriteInfo "GetBaseInfo", "����վ", gstrComputerName, "IP", gstrComputerIp, "System·��", gstrSystemPath, "��ʱĿ¼", gstrTempPath, "��¼������־", mblnWriteRunErrLog, "˵����Ϣ����", glngNoteLength
    GetBaseInfo = True
    Exit Function
errH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "GetBaseInfo", "������Ϣ��ȡ�������ش���", strErrInfo
    MsgBox "��ȡ������Ϣ������������ϵ����Ա����Ϣ��" & vbNewLine & strErrInfo, vbInformation, App.Title
    Err.Clear
    Resume
End Function

Public Sub RecordErrMsg(ByVal mtInput As MsgType, ByVal strErrObject As String, ByVal strErrInfo As String)
    Dim strSql As String
    grsErrRec.AddNew Array("����", "����", "��Ϣ"), Array(mtInput, strErrObject, strErrInfo)
    If mtInput > MT_MsgHeader And mtInput < MT_MsgFoot Then
        On Error Resume Next
        '��д������־
        strSql = "Zl_Zlclientupdatelog_Insert(" & SQLAdjust(strErrObject & ":" & strErrInfo) & "," & SQLAdjust(gstrTerminal) & ")"
        Call ExecuteProcedure(strSql, "RecordErrMsg")
        If Err.Number <> 0 Then Err.Clear
        
        '��д������־
        If mblnWriteRunErrLog Then
            '����=4 �ͻ�����������
            '�������=0
            strSql = "Zl_Zlerrorlog_Insert(" & SQLAdjust(gstrTerminal) & ",4,0," & SQLAdjust(strErrObject & ":" & strErrInfo) & "," & SQLAdjust(gstrAudsid) & " )"
            Call ExecuteProcedure(strSql, "RecordErrMsg")
            If Err.Number <> 0 Then Err.Clear
        End If
    ElseIf mtInput = MT_MsgHeader Or mtInput = MT_MsgFoot Then
        On Error Resume Next
        
        '��д������־
        strSql = "Zl_Zlclientupdatelog_Insert(" & SQLAdjust(strErrObject & ":" & strErrInfo) & "," & SQLAdjust(gstrTerminal) & ")"
        Call ExecuteProcedure(strSql, "RecordErrMsg")
        If Err.Number <> 0 Then Err.Clear
    End If
End Sub

Public Function ReportInfo(ByVal strInfo As String)
    Dim strSql      As String
    '��д������־
    On Error Resume Next
    strSql = "Zl_Zlclientupdatelog_Insert(" & SQLAdjust(strInfo) & "," & SQLAdjust(gstrTerminal) & ",1)"
    Call ExecuteProcedure(strSql, "ReportInfo")
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function SumErrMsg() As String
'���ܣ��ϲ�������Ϣ��������Ϣ����
    Dim strMsg As String, strPreObject As String
    
    grsErrRec.Filter = "����=" & MT_MsgHeader
    If Not grsErrRec.EOF Then strMsg = grsErrRec!��Ϣ
    grsErrRec.Filter = "����=" & MT_InitEnv
    Do While Not grsErrRec.EOF
        strMsg = strMsg & vbNewLine & grsErrRec!���� & ":" & grsErrRec!��Ϣ
        grsErrRec.MoveNext
    Loop
    grsErrRec.Filter = "����=" & MT_SvrConn
    Do While Not grsErrRec.EOF
        strMsg = strMsg & vbNewLine & grsErrRec!���� & ":" & grsErrRec!��Ϣ
        grsErrRec.MoveNext
    Loop
    
    grsErrRec.Filter = "����>" & MT_SvrConn & " And  ����<" & MT_ExeBat
    grsErrRec.Sort = "����,����"
    Do While Not grsErrRec.EOF
        If strPreObject <> grsErrRec!���� Then
            strPreObject = grsErrRec!����
            strMsg = strMsg & vbNewLine & grsErrRec!���� & ":"
        End If
        strMsg = strMsg & vbNewLine & "  " & grsErrRec!��Ϣ
        grsErrRec.MoveNext
    Loop
    grsErrRec.Filter = "����=" & MT_ExeBat
    Do While Not grsErrRec.EOF
        strMsg = strMsg & vbNewLine & grsErrRec!���� & ":" & grsErrRec!��Ϣ
        grsErrRec.MoveNext
    Loop
    grsErrRec.Filter = "����=" & MT_MsgFoot
    If Not grsErrRec.EOF Then strMsg = strMsg & vbNewLine & grsErrRec!��Ϣ
    SumErrMsg = strMsg
End Function

Private Function CheckFreeSpace() As Boolean
'���ܣ������̵�ʣ��ռ�
    '�����̿ռ䣬������1.5G,����ʾ����Ԥ����
    If gotCurType = OT_PreUpgrade Then
        If gobjFSO.Drives(Left(gstrSetupPath, 2)).FreeSpace / 1024 / 1024 < 500 Then
            gobjTrace.WriteInfo "���̿ռ���", "��Ϣ", "���пռ����500MB,�����޷�����Ԥ����"
            Call RecordErrMsg(MT_InitEnv, "���̿ռ���", "���пռ����500MB,�����޷�����Ԥ����")
        End If
    '��ʽ�������޸�������Ҫ��200M�ռ�
    Else
        If gobjFSO.Drives(Left(gstrSetupPath, 2)).FreeSpace / 1024 / 1024 < 200 Then
            gobjTrace.WriteInfo "���̿ռ���", "��Ϣ", "���пռ����200MB,�����޷�����" & Decode(gotCurType, OT_OfficialUpgrade, "����", OT_Repair, "�޸�", OT_CheckFile, "�ռ�")
            Call RecordErrMsg(MT_InitEnv, "���̿ռ���", "���пռ����200MB,�����޷�����" & Decode(gotCurType, OT_OfficialUpgrade, "����", OT_Repair, "�޸�", OT_CheckFile, "�ռ�"))
        End If
    End If
    CheckFreeSpace = True
End Function

Public Function GetActualPath(ByVal strSetupPath As String, ByVal ftFileType As FileType, ByVal strFile As String) As String
'���ܣ���ȡ�ļ���ʵ��·��
    Dim strKey As String, strPath As String
    
    If strSetupPath = "" Then
        Select Case ftFileType
            Case FT_Public
                strKey = "K_[PUBLIC]"
            Case FT_Apply
                strKey = "K_[APPSOFT]\APPLY"
            Case FT_Other, FT_AdditionFile
                strKey = "K_[APPSOFT]"
            Case FT_System
                strKey = "K_[SYSTEM]"
            Case FT_Help
                strKey = "K_[HELP]"
        End Select
    Else
        strKey = "K_" & strSetupPath
    End If
    strPath = gcllSetPath(strKey)
    GetActualPath = strPath & "\" & strFile
End Function

Public Function IsFileUpgade(ByVal strLoaclFile As String, ByVal strSvrVersion As String, ByVal strSvrModiTime As String, ByVal strSvrMD5 As String, Optional ByVal blnCheckReleated As Boolean)
'���ܣ��Ƿ��������
    Dim strlocVersion As String, strLocModiTime As String, strLocMd5 As String
    
    If Not gobjFSO.FileExists(strLoaclFile) Then
        '���ز����ڣ����жϷ��������Ƿ���ڣ�������������������������
        IsFileUpgade = strSvrMD5 <> ""
        Exit Function
    End If
    '�������ļ����ܴ��ڣ�������
    If strSvrMD5 = "" Then Exit Function
    '�޸����ںͰ汾�����������ж�MD5
    If strSvrVersion = "" Or strSvrModiTime = "" Then
        strLocMd5 = FileMD5(strLoaclFile)
        IsFileUpgade = strLocMd5 <> strSvrMD5
    Else
        strlocVersion = GetCommpentVersion(strLoaclFile)
        If Len(strlocVersion) <> Len(strSvrVersion) And UCase(gobjFSO.GetFileName(strLoaclFile)) Like "ZL*" Then
            strLocMd5 = FileMD5(strLoaclFile)
            IsFileUpgade = strLocMd5 <> strSvrMD5
            Exit Function
        End If
        strLocModiTime = gobjFSO.GetFile(strLoaclFile).DateLastModified
        IsFileUpgade = strlocVersion <> strSvrVersion Or Format(strSvrModiTime, "YYYY-MM-DD hh:mm:ss") <> Format(strLocModiTime, "YYYY-MM-DD hh:mm:ss")
    End If
End Function

Public Function GetHisUpdateCommand(Optional ByVal blnOld As Boolean) As String
'���ܣ���ȡ�Զ�������������
    Dim strCheck As String, strCommand As String
    Dim strUserName         As String, strPass      As String, strServer As String
    
    If blnOld Then
        GetHisUpdateCommand = gstrConnectString & "||1||" & gstrAppEXE & "||" & gstrHisCommand & "||" & gstrHisInput
    ElseIf gstrCommand Like "ZLUPDATE:*" Then
        GetHisUpdateCommand = "ZLUPDATEEX:" & Cipher(gstrCommand)
    ElseIf gstrCommand Like "ZLUPDATEEX:*" Or gstrCommand Like "ZLUPDATENEW:*" Then
        GetHisUpdateCommand = gstrCommand
    Else
        GetHisUpdateCommand = "ZLUPDATENEW:" & Cipher(gstrCommand)
    End If
End Function

Public Sub ClearFolder(ByVal strFolder As String, Optional ByVal blnOk As Boolean)
'���ܣ�����ִ���ļ���
    Dim objFolder       As Folder, objFile          As File, objTmpFolder           As Folder
    Dim cllFolders      As New Collection, cllFiles       As New Collection
    Dim strTmpFile      As String, strTmpFloder As String
    Dim blnAdd          As Boolean
    Dim i               As Long
    On Error Resume Next
    If InStr(UCase(gstrAppPath), "\ZLUPTMP") > 0 Or InStr(UCase(gstrAppPath), "\APPLY") > 0 Then
        FileNormal gstrSetupPath & "\ZLHisCrust.EXE"
        Call gobjFSO.CopyFile(gstrAppPath & "\ZLHisCrust.EXE", gstrSetupPath & "\ZLHisCrust.EXE", True)
        FileNormal gstrAppPath & "\ZLHisCrust.EXE"
        Call gobjFSO.DeleteFile(gstrAppPath & "\ZLHisCrust.EXE", True)
    End If
    If Err.Number <> 0 Then Err.Clear
    For Each objFolder In gobjFSO.GetFolder(strFolder).SubFolders
        'Ԥ��������ɾ��Ԥ��������Ŀ¼
        blnAdd = False
        If UCase(objFolder.Name) = "ZLPRETMP" Then
            If blnOk And (gotCurType = OT_OfficialUpgrade Or gotCurType = OT_Repair) Then
                blnAdd = True
            End If
        Else
            blnAdd = True
        End If
        If blnAdd Then
            cllFolders.Add objFolder.Path
            For Each objFile In objFolder.Files
                cllFiles.Add objFile.Path
            Next
        End If
    Next
    For i = 1 To cllFiles.Count
        Call gobjFSO.DeleteFile(cllFiles(i), True)
        If Err.Number <> 0 Then Err.Clear
    Next
    For i = 1 To cllFolders.Count
        Call gobjFSO.DeleteFolder(cllFolders(i), True)
        If Err.Number <> 0 Then Err.Clear
    Next
End Sub

Public Function FileNormal(ByVal strSource As String) As Boolean
'���ܣ����ļ����Լ�����Ŀ¼���Ƶ���һ��Ŀ¼
    On Error Resume Next
    If gobjFSO.FileExists(strSource) Then
        If FileSystem.GetAttr(strSource) <> vbNormal Then
            FileSystem.SetAttr strSource, vbNormal
        End If
    End If
    
    FileNormal = Err.Number = 0
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function GetAdditionSetup(ByVal strFileName As String, ByVal strMD5 As String, ByVal strAdditionSetup As String) As String
'���ܣ���ȡ���Ӱ�װ·���������ļ�����·�����е����������ļ�·�����ܰ������Ӱ�װ·���е�·��
    Dim arrAdd      As Variant, i           As Integer, j       As Integer
    Dim arrTmp      As Variant, strLast     As String
    Dim arrAllPath  As Variant, strAllPath  As String, strTmp   As String
    Dim strAllFile  As String, strLocMd5    As String
    Dim strPath     As String
    
    If strAdditionSetup = "" Or strMD5 = "" Then Exit Function
    arrAdd = Split(UCase(strAdditionSetup), "|")
    For i = LBound(arrAdd) To UBound(arrAdd)
        arrTmp = Split(arrAdd(i), "\")
        strPath = ""
        If UBound(arrTmp) <> -1 Then
            If arrTmp(0) = "[APPSOFT]" Then
                strPath = gstrSetupPath
            ElseIf arrTmp(0) = "[PUBLIC]" Then
                If Not gobjFSO.FolderExists(gstrSetupPath & "\PUBLIC") Then
                    gobjFSO.CreateFolder (gstrSetupPath & "\PUBLIC")
                End If
                strPath = gstrSetupPath & "\PUBLIC"
            ElseIf arrTmp(0) = "[APPLY]" Then
                strPath = gstrSetupPath & "\APPLY"
            ElseIf arrTmp(0) = "[OS:]" Then 'ϵͳ��
                strPath = Left(gstrSystemPath, 2)
            ElseIf arrTmp(0) = "[X:]" Then '��ǰ��װ��
                strPath = Left(gstrSetupPath, 2)
            End If
            If strPath <> "" Then
                strLast = Mid(arrAdd(i), Len(arrTmp(0) & "\") + 1)
                If strLast = "" Then
                    strTmp = strPath
                Else
                    strTmp = GetSubFloderByMach(strPath, strLast)
                End If
                If strTmp <> "" Then strAllPath = strAllPath & "|" & strTmp
            End If
        End If
    Next
    If strAllPath <> "" Then
        strAllPath = Mid(strAllPath, 2)
        arrAllPath = Split(strAllPath, "|")
        For i = LBound(arrAllPath) To UBound(arrAllPath)
            If gobjFSO.FileExists(arrAllPath(i) & "\" & strFileName) Then
                strLocMd5 = FileMD5(arrAllPath(i) & "\" & strFileName)
                If strMD5 <> strLocMd5 Then
                    strAllFile = strAllFile & "|" & arrAllPath(i) & "\" & strFileName
                    gobjTrace.WriteInfo "���Ӱ�װ���", "�ļ�", arrAllPath(i) & "\" & strFileName, "��Ϣ", "��·���ļ��ͷ������ļ�MD5����ͬ����Ҫ���Ӱ�װ"
                Else
                    gobjTrace.WriteInfo "���Ӱ�װ���", "�ļ�", arrAllPath(i) & "\" & strFileName, "��Ϣ", "��·���ļ��ͷ������ļ�MD5��ͬ������Ҫ���Ӱ�װ"
                End If
            Else
                strAllFile = strAllFile & "|" & arrAllPath(i) & "\" & strFileName
                gobjTrace.WriteInfo "���Ӱ�װ���", "�ļ�", arrAllPath(i) & "\" & strFileName, "��Ϣ", "��·�����ڵ����ļ������ڣ������Ҫ���ز����Ӱ�װ"
            End If
        Next
        If strAllFile <> "" Then strAllFile = Mid(strAllFile, 2)
    End If
    GetAdditionSetup = strAllFile
End Function

Private Function GetSubFloderByMach(ByVal strParentFloder As String, strMach As String) As String
'���ܣ���ȡƥ������ļ���
'strParentFloder:�����ļ���
'strMach:ƥ��·����
    Dim arrTmp      As Variant, strLast As String
    Dim objFolder   As Folder, blnLike  As Boolean, strLike As String
    Dim strTmp      As String, strReturn As String
    
    arrTmp = Split(strMach, "\")
    strLast = Mid(strMach, Len(arrTmp(0) & "\") + 1)
    If InStr(arrTmp(0), "[*]") > 0 Then
        strLike = Replace(arrTmp(0), "[*]", "*")
        For Each objFolder In gobjFSO.GetFolder(strParentFloder).SubFolders
            If UCase(objFolder.Name) Like strLike Then
                If strLast = "" Then
                    strTmp = strParentFloder & "\" & objFolder.Name
                Else
                    strTmp = GetSubFloderByMach(strParentFloder & "\" & objFolder.Name, strLast)
                End If
                If strTmp <> "" Then
                    strReturn = strReturn & "|" & strTmp
                End If
            End If
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        GetSubFloderByMach = strReturn
    Else
        If gobjFSO.FolderExists(strParentFloder & "\" & arrTmp(0)) Then
            If strLast = "" Then
                GetSubFloderByMach = strParentFloder & "\" & arrTmp(0)
            Else
                GetSubFloderByMach = GetSubFloderByMach(strParentFloder & "\" & arrTmp(0), strLast)
            End If
        End If
    End If
End Function

Public Function GetWrongFiles(ByVal strFileName As String, ByVal strSetupFile As String) As String
'���ܣ���ȡ�����ļ�·��
    Dim varItem         As Variant, strFileTmp              As String
    Dim strWrongFile    As String
    
    For Each varItem In gcllSetPath
        strFileTmp = varItem & "\" & strFileName
        If UCase(strFileTmp) <> UCase(strSetupFile) Then
            If gobjFSO.FileExists(strFileTmp) Then
                If strWrongFile <> "" Then '����[System]·����[help]·����ͬ����
                    If strWrongFile = "|" & strFileTmp Then
                    ElseIf InStr(strWrongFile, strFileTmp) = 0 Then
                        strWrongFile = strWrongFile & "|" & strFileTmp
                    End If
                Else
                    strWrongFile = strWrongFile & "|" & strFileTmp
                End If
            End If
        End If
    Next
    If strWrongFile <> "" Then strWrongFile = Mid(strWrongFile, 2)
    GetWrongFiles = strWrongFile
End Function

Private Function InitTerminal(ByRef strAudsid As String) As String
    Dim rsTmp As adodb.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select Userenv('SessionID') Audsid ,Userenv('Terminal')  Terminal From dual"
    Set rsTmp = OpenSQLRecord(strSql, "InitTerminal")
    
    If Not rsTmp.EOF Then
        strAudsid = rsTmp!Audsid
        InitTerminal = rsTmp!Terminal
    End If
    
    Exit Function
errH:
    MsgBox Err.Description
End Function

