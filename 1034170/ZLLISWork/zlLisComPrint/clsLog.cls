VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'********************************************************
'* ������־��
'*
'********************************************************
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'---- ��־��Ϣ�����̶�
Private Enum LOG_Severity
    LOG_OFF = 1         '�����
    LOG_FATAL = 2       '������
    LOG_ERR = 3         '������Ϣ
    LOG_WARNING = 4     '������Ϣ
    LOG_INFO = 6        '֪ͨ����Ϣ
    LOG_DEBUG = 7       '���Լ�����Ϣ�����
End Enum

Private mblnInit As Boolean '�Ƿ��ѳɹ���ʼ��

Private Type logConfig
    'ͨ��������
    LogAppender As String        '��־�����ʽ��FILE,Ŀǰֻ֧���ļ���ʽ��������
    LogLevel As Integer          'OFF>FATAL>ERROR>WARN>INFO>DEBUG
    LogFormat    As String       '��־��ʽ
    LogIni      As String       '�����ļ����·��
    '%level:��־��Ϣ����
    '%date{}  :��־��Ϣ����ʱ��
    '%prg  :��־��Ϣ���ڵأ����������������
    '%pid  : ������־��Pig
    '%tag  : ��������־���     strtag
    '%Const: ��������־������Ϣ strConst
    '%n:   : ��־��Ϣ����
    
    '����%level %date{yyyy-MM-dd HH:mm:ss} %prg=>%tag  [%pid]  %Const  %n
    '���: ERR    2010-03-11 12:03:44 ͨѶ����=>��ʼ���ӿ�    [5] �಻֧���Զ���! <CR>
    
    '��־�ļ���������
    LogFile  As String           '��־�ļ�������·��
    LogMaxFileSize  As Integer   '��־�ļ���С��byte��
    LogMaxBackupIndex As Integer '��־�ļ���������
    LogBackupIndex  As Integer   '��ǰ��־�ļ�˳��
    LogProgName     As String    '�����ߣ�һ��Ϊ������
    'DataBase������
    'Sock������
End Type

Private mConfig As logConfig
Private mFSO As New FileSystemObject

Public Sub LogDebug(ByVal strTag As String, ByVal lngPid As Long, ByVal strConst As String, ByRef strErr As String)
    'д�������Ϣ
    'lngPid   :����Ż��к�
    'strTag   :���
    'strConst :����
    If mConfig.LogLevel < LOG_DEBUG Then Exit Sub

    If InStr(mConfig.LogAppender, "FILE") > 0 Then
        'д�ļ�
        Call WritLogFile(LOG_DEBUG, lngPid, strTag, strConst, strErr)
    End If
End Sub

Public Sub LogInfo(ByVal strTag As String, ByVal lngPid As Long, ByVal strConst As String, ByRef strErr As String)
    'д����ʾ��Ϣ
    
    'strTag   :���
    'lngPid   :����Ż��к�
    'strConst :����
    If mConfig.LogLevel < LOG_INFO Then Exit Sub
    If InStr(mConfig.LogAppender, "FILE") > 0 Then
        'д�ļ�
        Call WritLogFile(LOG_INFO, lngPid, strTag, strConst, strErr)
    End If
End Sub

Public Sub LogWarn(ByVal strTag As String, ByVal lngPid As Long, ByVal strConst As String, ByRef strErr As String)
    'д�뾯����Ϣ
    
    'strTag   :���
    'lngPid   :����Ż��к�
    'strConst :����
    If mConfig.LogLevel < LOG_WARNING Then Exit Sub
    If InStr(mConfig.LogAppender, "FILE") > 0 Then
        'д�ļ�
        Call WritLogFile(LOG_WARNING, lngPid, strTag, strConst, strErr)
    End If
End Sub

Public Sub LogError(ByVal strTag As String, ByVal lngPid As Long, ByVal strConst As String, ByRef strErr As String)
    'д�������Ϣ
    'strTag   :���
    'lngPid   :����Ż��к�
    'strConst :����
    strErr = ""
    If mConfig.LogLevel < LOG_ERR Then Exit Sub
    If InStr(mConfig.LogAppender, "FILE") > 0 Then
        'д�ļ�
        Call WritLogFile(LOG_ERR, lngPid, strTag, strConst, strErr)
    End If
End Sub

Public Function LogInit(ByVal strProgName As String, ByRef strErr As String, Optional ByVal strConFile As String, Optional blnReLoad As Boolean = False, Optional strLogFileName As String) As Boolean
        '��ʼ��Log����
        'strProgName:���ó��������硰ͨ�ų���
        'strErr     :���ش�����ʾ
        'strConFile :�����ļ������Ϊ�գ���ȡ��������Ŀ¼��log\logConfig.ini�ļ�
        'blnReLoad  :�Ƿ����¶�ȡ���á�
        
        On Error GoTo hErr
100     strErr = ""
    
102     If mblnInit And blnReLoad = False Then
104         LogInit = True
            Exit Function
        End If
106     If strProgName = "" Then
108         strErr = "���ó���������Ϊ��!"
            Exit Function
        End If
    
110     If strConFile = "" Then
112         strConFile = App.Path & "\log\logConfig.ini"
        End If
114     mConfig.LogIni = strConFile
116     If Not mFSO.FolderExists(mFSO.GetParentFolderName(strConFile)) Then
118         Call mFSO.CreateFolder(mFSO.GetParentFolderName(strConFile))
        End If
120     If Not mFSO.FileExists(strConFile) Then
122         mFSO.CreateTextFile strConFile
        End If
    
124     mConfig.LogProgName = strProgName
126     mConfig.LogAppender = "FILE"    '��ֻ֧���ļ�����
    
128     If Dir(strConFile) <> "" Then
            '��ȡ�����ļ��е�����
130         mConfig.LogFile = IniRead(strProgName, "LogFile", "")
            
132         If mConfig.LogFile = "" Then
                
134             If InStr(strLogFileName, ":") < 0 And Right(strLogFileName, 8) <> "_log.txt" Then
                    '����ȫ·������Ĭ������
                    mConfig.LogFile = mFSO.GetParentFolderName(strConFile) & "\" & strProgName & "_log.txt"
                Else
138                 mConfig.LogFile = strLogFileName
                End If
140         ElseIf mConfig.LogFile <> strLogFileName And strLogFileName <> "" Then
142             If InStr(strLogFileName, ":") < 0 And Right(strLogFileName, 8) <> "_log.txt" Then
144                 mConfig.LogFile = mFSO.GetParentFolderName(strConFile) & "\" & strProgName & "_log.txt"
                Else
146                 mConfig.LogFile = strLogFileName
                End If
            End If
148         IniWrite strProgName, "LogFile", mConfig.LogFile
        
150         mConfig.LogFormat = IniRead(strProgName, "LogFormat", "")
152         If mConfig.LogFormat = "" Then mConfig.LogFormat = "%level" & vbTab & "%date{yyyy-MM-dd HH:mm:ss}" & vbTab & "[%pid]" & "%Const" & vbTab & "%prg" & vbTab & "%tag%n"
154         IniWrite strProgName, "LogFormat", mConfig.LogFormat
        
156         mConfig.LogLevel = Val(IniRead(strProgName, "LogLevel", LOG_INFO))
158         If mConfig.LogLevel < LOG_OFF Or mConfig.LogLevel > LOG_DEBUG Then mConfig.LogLevel = LOG_INFO
160         IniWrite strProgName, "LogLevel", mConfig.LogLevel
        
162         mConfig.LogMaxBackupIndex = Val(IniRead(strProgName, "LogMaxBackupIndex", 3))
164         If mConfig.LogMaxBackupIndex <= 0 Or mConfig.LogMaxBackupIndex > 512 Then mConfig.LogMaxBackupIndex = 3
166         IniWrite strProgName, "LogMaxBackupIndex", mConfig.LogMaxBackupIndex
        
168         mConfig.LogMaxFileSize = Val(IniRead(strProgName, "LogMaxFileSize", 512))
170         If mConfig.LogMaxFileSize <= 10 Or mConfig.LogMaxFileSize > 3072 Then mConfig.LogMaxFileSize = 512
172         IniWrite strProgName, "LogMaxFileSize", mConfig.LogMaxFileSize
        
        Else
            '�������ļ�����Ĭ��ֵ��ʼ������д�������ļ���
174         If InStr(strLogFileName, ":") < 0 And Right(strLogFileName, 8) <> "_log.txt" Then
                mConfig.LogFile = mFSO.GetParentFolderName(strConFile) & "\" & strProgName & "_log.txt"
            Else
178             mConfig.LogFile = strLogFileName
            End If
180         mConfig.LogFormat = "%level" & vbTab & "%date{yyyy-MM-dd HH:mm:ss}" & vbTab & "[%pid]" & "%Const" & vbTab & "%prg" & vbTab & "%tag%n"
182         mConfig.LogLevel = LOG_INFO
184         mConfig.LogMaxBackupIndex = 3
186         mConfig.LogMaxFileSize = 1024
        
188         IniWrite strProgName, "LogFile", mConfig.LogFile
190         IniWrite strProgName, "LogLevel", mConfig.LogLevel
192         IniWrite strProgName, "LogFormat", mConfig.LogFormat
194         IniWrite strProgName, "LogMaxBackupIndex", mConfig.LogMaxBackupIndex
196         IniWrite strProgName, "LogMaxFileSize", mConfig.LogMaxFileSize
        End If

198     mblnInit = True
200     LogInit = True
        Exit Function
hErr:
202     LogInit = False
204     strErr = CStr(Erl()) & "��,��ʼ����־����ʧ�ܣ�" & Err.Description
End Function

Public Function LogSetConfig(ByVal strItem As String, ByVal strValue As String, ByRef strErr As String) As Boolean
        '����init��������
        '���ⲿ����
        'strItem����LogFile����LogLevel��LogFormat��LogMaxBackupIndex��LogMaxFileSize
        Dim strProgName As String
        On Error GoTo hErr
100     strProgName = mConfig.LogProgName
102     strErr = ""
104     Select Case strItem
            Case "LogFile"
106             If strValue = "" Then strValue = strProgName
108             mConfig.LogFile = App.Path & "\log\" & strValue & "_log.txt"
110             IniWrite strProgName, "LogFile", mConfig.LogFile

112         Case "LogFormat"
    
114             If strValue = "" Then strValue = "%level" & vbTab & "%date{yyyy-MM-dd HH:mm:ss}" & vbTab & "[%pid]" & "%Const" & vbTab & "%prg" & vbTab & "%tag%n"
116             mConfig.LogFormat = strValue
118             IniWrite strProgName, "LogFormat", mConfig.LogFormat
    
120         Case "LogLevel"
122             If Val(strValue) < LOG_OFF Or Val(strValue) > LOG_DEBUG Then strValue = LOG_INFO
124             mConfig.LogLevel = Val(strValue)
126             IniWrite strProgName, "LogLevel", mConfig.LogLevel
    
128         Case "LogMaxBackupIndex"
130             If Val(strValue) <= 0 Or Val(strValue) > 512 Then strValue = 3
132             mConfig.LogMaxBackupIndex = Val(strValue)
134             IniWrite strProgName, "LogMaxBackupIndex", mConfig.LogMaxBackupIndex
    
136         Case "LogMaxFileSize"
138             If Val(strValue) <= 10 Or Val(strValue) > 3072 Then strValue = 1024
140             mConfig.LogMaxFileSize = Val(strValue)
142             IniWrite strProgName, "LogMaxFileSize", mConfig.LogMaxFileSize
144         Case Else
146             strErr = "�����������"
        End Select
148     LogSetConfig = strErr = ""
        Exit Function
hErr:
150     LogSetConfig = False
152     strErr = CStr(Erl()) & "��,������Ŀ" & strItem & "ʧ�ܣ�" & Err.Description
End Function

Private Sub WritLogFile(ByVal lngLogType As Long, ByVal lngPid As Long, ByVal strTag As String, ByVal strConst As String, ByRef strErr As String)
        Dim strPath As String, objStream As TextStream
        Dim i As Integer, lngCount As Long, strWritInfo As String
        
        Dim lngFileSize As String
        On Error GoTo hErr
    
100     If Not mblnInit Then Exit Sub
102     If strTag = "" Or strConst = "" Then
104         strErr = "��Tag��Const���ݣ�д����־ʧ�ܣ�"
            Exit Sub
        End If
106     lngFileSize = mConfig.LogMaxFileSize
108     lngFileSize = lngFileSize * 1024
109     strPath = mFSO.GetParentFolderName(mConfig.LogFile)
110     If Not mFSO.FolderExists(strPath) Then mFSO.CreateFolder strPath
111     If GetFreeSpace(strPath) <= lngFileSize Then
112         strErr = "���̿ռ䲻�㣬д����־ʧ�ܣ�"
            Exit Sub
        End If
        '------ ��־��Ϣ����
        strWritInfo = FormatLogInfo(lngLogType, strTag, lngPid, strConst)
    
        '------ LOG�ļ�����
158     mConfig.LogBackupIndex = Val(IniRead(mConfig.LogProgName, "LogBackupIndex"))
    
160
162

164     If Not mFSO.FileExists(mConfig.LogFile) Then
            '--- û����־�ļ����򴴽�
166         Call mFSO.CreateTextFile(mConfig.LogFile)
168     ElseIf Val(mFSO.GetFile(mConfig.LogFile).Size) >= lngFileSize Then
            '����־�ļ����������˹涨��С���򽫵�ǰ�ļ�������Ȼ�󴴽�һ���յ���־�ļ���
170         lngCount = 0
172         For i = 1 To mConfig.LogMaxBackupIndex
174             If mFSO.FileExists(mConfig.LogFile & "bak" & i) Then
                    lngCount = i
                Else
                    Exit For
                End If
            Next
        
176         If lngCount >= mConfig.LogMaxBackupIndex Then
                '�б����ļ����ļ��Ŵ��ڵ�������ļ���,ȡ�����ļ��е����+1������������ԭΪ��С��
178             If mConfig.LogBackupIndex + 1 <= mConfig.LogMaxBackupIndex Then
180                 mConfig.LogBackupIndex = mConfig.LogBackupIndex + 1
                Else
182                 mConfig.LogBackupIndex = 1
                End If
            Else
                '�ļ���С������ļ�
184             mConfig.LogBackupIndex = lngCount + 1
            End If
        
186         If mFSO.FileExists(mConfig.LogFile & "bak" & mConfig.LogBackupIndex) Then Call mFSO.DeleteFile(mConfig.LogFile & "bak" & mConfig.LogBackupIndex, True)
188         Call mFSO.CopyFile(mConfig.LogFile, mConfig.LogFile & "bak" & mConfig.LogBackupIndex, True)
190         Call mFSO.DeleteFile(mConfig.LogFile, True)
192         Call mFSO.CreateTextFile(mConfig.LogFile)
            '������־�ļ����
194         IniWrite mConfig.LogProgName, "LogBackupIndex", mConfig.LogBackupIndex
        End If
196     Set objStream = mFSO.OpenTextFile(mConfig.LogFile, ForAppending)
198     If strWritInfo <> "" Then objStream.Write strWritInfo
200     objStream.Close
202     Set objStream = Nothing
204     strErr = ""
        Exit Sub
hErr:
206     strErr = CStr(Erl()) & "��,����" & lngLogType & "��־ʧ�ܣ�" & Err.Description
End Sub

Public Function FormatLogInfo(lngLogType As Long, ByVal strTag As String, ByVal lngPid As Long, ByVal strConst As String) As String
    '�������ȶ���ĸ�ʽ����ʽ����־��Ϣ����
    Dim strWritInfo As String, strType As String
    Dim lng_S As Long, lng_E As Long, strDateFormat As String
    
        If lngLogType = LOG_DEBUG Then
            strType = "����"
        ElseIf lngLogType = LOG_INFO Then
            strType = "��ʾ"
        ElseIf lngLogType = LOG_WARNING Then
            strType = "����"
        Else
            strType = "����"
        End If
114     strWritInfo = mConfig.LogFormat
    
116     strWritInfo = Replace(strWritInfo, "%level", strType)
118     strWritInfo = Replace(strWritInfo, "%prg", mConfig.LogProgName)
120     strWritInfo = Replace(strWritInfo, "%tag", strTag)
122     strWritInfo = Replace(strWritInfo, "%pid", lngPid)
124     strWritInfo = Replace(strWritInfo, "%Const", strConst)
126     strWritInfo = Replace(strWritInfo, "%n", vbCrLf)
    
128     strDateFormat = ""
130     lng_S = InStr(strWritInfo, "%date{")
132     If lng_S > 0 Then
134         lng_E = InStr(lng_S, strWritInfo, "}")
        Else
136         lng_E = 0
        End If
138     Do While lng_S > 0 And lng_E > 0
140         strDateFormat = Mid(strWritInfo, lng_S, lng_E - lng_S + 1)
142         strDateFormat = Replace(strDateFormat, "%date{", "")
144         strDateFormat = Replace(strDateFormat, "}", "")
146         strDateFormat = Format(Now, strDateFormat)
        
148         strWritInfo = Mid(strWritInfo, 1, lng_S - 1) & strDateFormat & Mid(strWritInfo, lng_E + 1)
150         lng_S = InStr(strWritInfo, "%date{")
152         If lng_S > 0 Then
154             lng_E = InStr(lng_S, strWritInfo, "}")
            Else
156             lng_E = 0
            End If
        Loop
        FormatLogInfo = strWritInfo
        
End Function
Private Function IniRead(ByVal SessionName As String, ByVal KeyString As String, Optional ByVal strDefault As String) As String
    Dim Filename As String
    Dim Retval As Long
    Dim value As String
    Filename = mConfig.LogIni
    
    Filename = IIf(Mid(Filename, 2, 2) = ":\", Filename, Replace(App.Path & "\" & Filename, "\\", "\"))
    value = Space(128)
    Retval = GetPrivateProfileString(SessionName, KeyString, strDefault, value, Len(value), Filename)
    IniRead = Left(Trim(value), Len(Trim(value)) - 1)
End Function

Private Function IniWrite(ByVal SessionName As String, ByVal KeyName As String, ByVal KeyValue As String) As Boolean

    Dim Retval As Long, Filename As String
    Filename = mConfig.LogIni
    Filename = IIf(Mid(Filename, 2, 2) = ":\", Filename, Replace(App.Path & "\" & Filename, "\\", "\"))
    IniWrite = WritePrivateProfileString(SessionName, KeyName, KeyValue, Filename)
End Function

Private Function GetFreeSpace(ByVal strPath As String) As Double
    Dim strDriv As String, Drv As Drive
    Dim strDir As String
    
    If mFSO.FolderExists(strPath) Then
        strDriv = mFSO.GetDriveName(mFSO.GetAbsolutePathName(strPath))
        Set Drv = mFSO.GetDrive(strDriv)
        If Drv.IsReady Then
            GetFreeSpace = Drv.FreeSpace
        End If
        Set Drv = Nothing
    End If
End Function



