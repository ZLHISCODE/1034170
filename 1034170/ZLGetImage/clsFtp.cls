VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFtp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''FTP��API����'''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '''''''''''''''''''''
                                ''''FTP���Ӳ���'''''''
                                '''''''''''''''''''''
'��һ�������������͵�Internet����
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'hInternetSession--����InternetOpen������Internet�Ի����ص�ֵ
'sServerName--Ҫ���ӵķ����������ƻ�IP
'nServerPort--�����ӵ�Internet�˿�
'sUsername--��¼���û��ʺ�
'sPassword--��¼�Ŀ���
'lService--Ҫ���ӵķ��������ͣ�����������FTP�����������ӵ�����Ϊ����INTERNET_SERVICE_FTP��

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'����Internet����ĳ���
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_GOPHER = 2
Private Const INTERNET_SERVICE_HTTP = 3
Private Const INTERNET_FLAG_PASSIVE = &H8000000     '���ñ�������
Private Const MAX_FILENAME = 260

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_FILENAME
        cAlternate As String * 14
End Type

                                ''''''''''''''''''''''''
                                '''''''FTPĿ¼����''''''
                                '''''''''''''''''''''''
'��ftp�������ϴ���Ŀ¼
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'lpszDirectory--����Ҫ����Ŀ¼���ַ�����������һ�����·�������·��
 '���ݴ�internet���ӵĺ���internetopen�������صľ�����û���������
        '����ftp������������
        
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, ByRef lpdwCurrentDirectory As Long) As Boolean




Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean


                                ''''''''''''''''''''''''
                                '''''''FTP�ļ�����''''''
                                '''''''''''''''''''''''
    
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean

Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
'FtpFindFirstFile
'InternetFindNextFile

   
Private Const FILE_ATTRIBUTE_NORMAL = &H80
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''����'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mstrHost As String
Private mstrUser As String
Private mstrPwd As String
Private mstrFtpRoot As String

Public hOpen As Long
Public hConnection As Long

'��¼��־״̬�ı���
Private mblnLogEnable As Boolean     '�Ƿ�������־
Private mstrLogPath As String        '��־·��
Private mlngLogLevel As Long         '��־����1��Ϊ��ͨ��־��2��Ϊ��ϸ��־

Private mstrIPAddress As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''����'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get IPAddress() As String
    IPAddress = mstrIPAddress
End Property

Public Property Get User() As String
    User = mstrUser
End Property


Public Property Get PassWord() As String
    PassWord = mstrPwd
End Property

'mblnLogEnable
Public Property Let blnLogEnable(ByVal vData As Boolean)
    mblnLogEnable = vData
    '���������־��������־��·��Ϊ�գ���������־·��ΪĬ��·��
    On Error Resume Next
    
    If mblnLogEnable = True Then
        If mstrLogPath = "" Then
            mstrLogPath = App.Path & "\GetImgLog"


        End If
        
         '�����־·�������ڣ��򴴽�
        If Dir(mstrLogPath, vbDirectory) = "" Then
            'Ĭ��·�������ڣ��������Ŀ¼
            If Dir(mstrLogPath, vbDirectory) = "" Then
                Call MkLocalDir(mstrLogPath)
            End If
        End If
    End If
End Property

Public Property Get blnLogEnable() As Boolean
    blnLogEnable = mblnLogEnable
End Property

'mstrLogPath
Public Property Let strLogPath(ByVal vData As String)
    mstrLogPath = vData
    
    If mstrLogPath = "" Then
        mstrLogPath = App.Path & "\GetImgLog"
    End If

    
     '�����־·�������ڣ��򴴽�
    If Dir(mstrLogPath, vbDirectory) = "" Then
        'Ĭ��·�������ڣ��������Ŀ¼
        If Dir(mstrLogPath, vbDirectory) = "" Then
            Call MkLocalDir(mstrLogPath)
        End If
    End If
    
End Property


Public Property Get strLogPath() As String
    strLogPath = mstrLogPath
End Property

'mlngLogLevel
Public Property Let lngLogLevel(ByVal vData As Long)
    mlngLogLevel = vData
End Property

Public Property Get lngLogLevel() As Long
    lngLogLevel = mlngLogLevel
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ReadFtpTransferState() As Boolean
On Error GoTo errHandle
    ReadFtpTransferState = IIf(Val(GetSetting("ZLSOFT", "����ģ��\Ftp", "���ñ�������", 0)) = 1, True, False)
    Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "���ñ�������", IIf(ReadFtpTransferState, 1, 0))
Exit Function
errHandle:
    ReadFtpTransferState = 0
End Function

Public Function FuncFtpConnect(strHost As String, strUser As String, strPassWord As String) As Long
    Dim strDecryptionPassW As String
    Dim blnPasv As Boolean
    Dim strRootPath As String
    
    blnPasv = ReadFtpTransferState
    
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    
    '��������Ѿ����ܣ�������ܼ�������
    If Len(strPassWord) >= 3 Then
        If Mid(strPassWord, 1, 1) & Mid(strPassWord, 3, 1) & Mid(strPassWord, Len(strPassWord), 1) = "�����" Then
            strDecryptionPassW = Mid(strPassWord, 2)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, Len(strDecryptionPassW) - 1)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, 1) & Mid(strDecryptionPassW, 3)
            strDecryptionPassW = GetDecryptionPassW(strDecryptionPassW)
            
            strPassWord = strDecryptionPassW
        End If
    End If
    
    hConnection = InternetConnect(hOpen, strHost, 0, strUser, strPassWord, INTERNET_SERVICE_FTP, IIf(blnPasv = True, INTERNET_FLAG_PASSIVE, 0), IIf(blnPasv = True, 0, 255))
    If hConnection = 0 Then
        '���Ĵ���ģʽ��������
        blnPasv = Not blnPasv
        hConnection = InternetConnect(hOpen, strHost, 0, strUser, strPassWord, INTERNET_SERVICE_FTP, IIf(blnPasv = True, INTERNET_FLAG_PASSIVE, 0), IIf(blnPasv = True, 0, 255))
        
        If hConnection <> 0 Then
        '���洫����������
            If blnPasv Then
                Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "���ñ�������", 1)
            Else
                Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "���ñ�������", 0)
            End If
        End If
    End If
    
    mstrIPAddress = strHost
    mstrHost = strHost
    mstrUser = strUser
    mstrPwd = strPassWord
    
    '��ȡftp��Ŀ¼
    mstrFtpRoot = GetFtpRootPath
    
    FuncFtpConnect = hConnection
End Function

Private Function GetFtpCWD(Optional ByVal strDefault As String = "") As String
'��ȡftp��ǰ����Ŀ¼
On Error GoTo errHandle
    Dim strCwd As String
    Dim lngsize As Long
    Dim blnResult As Boolean
    
    GetFtpCWD = strDefault
    
    If hConnection = 0 Then Exit Function
    
    lngsize = 255
    strCwd = String(lngsize, 0)
    
    blnResult = FtpGetCurrentDirectory(hConnection, strCwd, lngsize)
    If blnResult = False Then Exit Function
    
    strCwd = Mid$(strCwd, 1, lngsize)
    
    GetFtpCWD = strCwd
Exit Function
errHandle:
    GetFtpCWD = strDefault
End Function

Private Function GetFtpRootPath(Optional blnRestore As Boolean = False) As String
'��ȡ��Ŀ¼·��
    If blnRestore Then
        Call FuncFtpDisConnect
        Call FuncFtpConnect(mstrHost, mstrUser, mstrPwd)
    End If
    
    GetFtpRootPath = GetFtpCWD("/")
End Function


Public Function ResotreFtpConnect() As Long
    ResotreFtpConnect = FuncFtpConnect(mstrHost, mstrUser, mstrPwd)
End Function

Public Function FuncFtpDisConnect()
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    
    mstrIPAddress = ""
    mstrUser = ""
    mstrPwd = ""
End Function

Private Function FuncSetVirtualPath(ByVal strVirtualPath As String, Optional ByVal blnIsReStart As Boolean = True) As Boolean
'���õ�ǰ����Ŀ¼
    Dim blnResult As Boolean
    Dim strFormatVirtualPath As String
    Dim strCurWorkDirectory As String
    Dim strFullPath As String
    
    FuncSetVirtualPath = False
    If hConnection = 0 Then Exit Function
    
    strFormatVirtualPath = FormatVirtualPath(strVirtualPath)
    
    If blnIsReStart Then
        strCurWorkDirectory = GetFtpCWD(mstrFtpRoot)
        
        strFullPath = mstrFtpRoot & IIf(strFormatVirtualPath <> "" And mstrFtpRoot <> "/", "/", "") & strFormatVirtualPath
        
        If UCase(strCurWorkDirectory) = UCase(strFullPath) Then Exit Function
        
        
        blnResult = True
        While (UCase(strCurWorkDirectory) <> UCase(mstrFtpRoot)) And blnResult And strCurWorkDirectory <> "/"
            blnResult = IIf(FtpSetCurrentDirectory(hConnection, "../"), True, False)
            strCurWorkDirectory = GetFtpCWD(mstrFtpRoot)
        Wend

    End If
    
    blnResult = IIf(FtpSetCurrentDirectory(hConnection, "./" & strFormatVirtualPath), True, False)
    
    If Not blnResult Then
        blnResult = IIf(FtpSetCurrentDirectory(hConnection, "/" & strFormatVirtualPath), True, False)
    End If
    
    If Not (blnResult) Then
        '�������Ŀ¼�����豸����Ͽ��������½�������Ŀ¼����
        FuncFtpDisConnect
        
        FuncFtpConnect mstrHost, mstrUser, mstrPwd
        
        If hConnection = 0 Then Exit Function
        
        blnResult = IIf(FtpSetCurrentDirectory(hConnection, "./" & strFormatVirtualPath), True, False)
        If Not blnResult Then
            blnResult = IIf(FtpSetCurrentDirectory(hConnection, "/" & strFormatVirtualPath), True, False)
        End If
        
        Debug.Print "Set Ftp Path Result:" & blnResult & " Path Way: // To ./" & strFormatVirtualPath
    End If
End Function


Public Function FuncFtpMkDir(strVirtualPath As String, strDir As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ�����FTPĿ¼
    '''������strIPAddress-----FTP����������IP��ַ
    '''������strMainPath----FTP�����е�����Ŀ¼��
    '''������strUser---FTP��½���û���
    '''������strPsw---FTP��½������
    '''������strDir---��Ҫ������Ŀ¼��ʹ�á�/��������Ŀ¼�ָ���
    '''���أ��ɹ�����0������ʧ�ܷ���1������Ŀ¼ʧ�ܷ���2
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim i As Integer
    Dim aNestDirs() As String               '�洢������Ҫ������Ŀ¼
    Dim bResult As Boolean                  '�������
        
    FuncFtpMkDir = 1
    
    '��FTP����
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
        '����FTP������Ŀ¼
        FuncSetVirtualPath strVirtualPath
        
        '��ȡȫ����Ҫ������Ŀ¼��Ϣ
        aNestDirs = Split(Replace(strDir, "//", "/"), "/")
        
        '����ȫ��Ŀ¼
        For i = 0 To UBound(aNestDirs)
            If aNestDirs(i) <> vbNullString Then
                bResult = FtpCreateDirectory(hConnection, aNestDirs(i))
                FuncSetVirtualPath aNestDirs(i), False
            End If
        Next
    End If
    
    FuncFtpMkDir = IIf(bResult, 0, 2)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FuncFtpDelDir(strVirtualPath As String, strDir As String) As Long
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ�ɾ��ָ����Ŀ¼
    '''������strIPAddress---FTP�����������ƻ�IP��ַ
    '''������strVirtualPath---FTP������Ŀ¼����
    '''������strUser---FTP��½���û���
    '''������strPsw---FTP��½������
    '''������strDir---��Ҫɾ����Ŀ¼��ʹ�á�/��������Ŀ¼�ָ�����ֻɾ�����һ��Ŀ¼
    '''���أ��ɹ�����0������ʧ�ܷ���1��ɾ��ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim bResult As Boolean      '���淵��ֵ
    
    FuncFtpDelDir = 1
    '��FTP����
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        FuncSetVirtualPath strVirtualPath
           
        bResult = FtpRemoveDirectory(hConnection, Replace(strDir, "//", "/"))
    End If
    
    FuncFtpDelDir = IIf(bResult, 0, 2)
End Function

Public Function FuncUploadFile(strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String) As Long
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ���FTP��������ָ��Ŀ¼�ϴ�һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strLocalFileName---�����ļ�����·�����硰C:\SSS.TXT��
    '''������strRemoteFileName---�ϴ�����ļ���
    '''���أ��ɹ�����0������ʧ�ܷ���1���ϴ��ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim bResult As Boolean
    
    FuncUploadFile = 1
    '��FTP����
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        FuncSetVirtualPath strVirtualPath
         
        bResult = FtpPutFile(hConnection, strLocalFileName, Replace(strRemoteFileName, "//", "/"), 2, 0)
    End If
    
    FuncUploadFile = IIf(bResult, 0, 2)
End Function

Public Function FuncChangeDir(strVirtualPath As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ�FTP��������ָ��Ŀ¼����һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''���أ��ɹ�����0��ת��Ŀ¼ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncChangeDir = 1
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        bResult = FuncSetVirtualPath(strVirtualPath)
    End If
    
    FuncChangeDir = IIf(bResult, 0, 2)
End Function

Public Function FuncDownloadOneFile(strLocalFileName As String, strRemoteFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ�FTP��������ָ��Ŀ¼����һ���ļ�
    '''������strLocalFileName---�����ļ�����·�����硰C:\SSS.TXT��
    '''������strRemoteFileName---��Ҫ���ص��ļ���
    '''���أ��ɹ�����0������ʧ�ܷ���1������ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncDownloadOneFile = 1
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
        bResult = FtpGetFile(hConnection, strRemoteFileName, strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
    End If
    
    FuncDownloadOneFile = IIf(bResult, 0, 2)
End Function


Public Function DownloadFileValid(strFtpPath As String, ByVal strLocalFile As String, ByVal strFileName As String) As Boolean
'����ϴ����ļ��ͱ����ļ���С�Ƿ�һ�£���һ���򷵻�true
    Dim blnResult As Boolean, blnReUpload As Boolean
    Dim lngFtpFileSzie As Long, lngDestFileSize As Long
    Dim strMessage As String
    Dim objFileSystem As New FileSystemObject
    
On Error GoTo errHandle
    DownloadFileValid = False
    
    '�ϴ���Ա�һ�´�С���ж��Ƿ������ϴ�
    lngDestFileSize = objFileSystem.GetFile(strLocalFile).Size
    lngFtpFileSzie = FuncFtpGetFileSize(strFtpPath, strFileName)

    If lngFtpFileSzie <> lngDestFileSize Then
        'ǿ���ض�
        Call FuncDownloadFile(strFtpPath, strLocalFile, strFileName, True)
        
        lngDestFileSize = objFileSystem.GetFile(strLocalFile).Size
    End If

    If lngFtpFileSzie <> lngDestFileSize Then
        strMessage = "���غ�ı����ļ���С[" & lngDestFileSize & "]��FTPԭ�ļ���С[" & lngFtpFileSzie & "]��һ��" & vbCrLf & _
                     "�����ļ���" & strLocalFile & vbCrLf & _
                     "FTPԭ�ļ���" & strFtpPath & strFileName & vbCrLf & _
                     "�Ƿ���Ҫ�������أ�"
        
        If MsgBox(strMessage, vbQuestion + vbYesNo, "��ʾ") = vbYes Then
            DownloadFileValid = True
        End If
    End If
Exit Function
errHandle:
    If MsgBox("ͼ���ļ�[����:" & strLocalFile & "  FTP:" & strFtpPath & "/" & strFileName & "]һ���Լ�����,����ԭ��:" & err.Description & "��" & vbCrLf & "�Ƿ����ԣ�", vbQuestion + vbYesNo, "��ʾ") = vbYes Then
        DownloadFileValid = True
    End If
End Function


Public Function FuncDownloadFile(strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String, Optional ByVal blnIsForceRead As Boolean = False) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ�FTP��������ָ��Ŀ¼����һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strLocalFileName---�����ļ�����·�����硰C:\SSS.TXT��
    '''������strRemoteFileName---��Ҫ���ص��ļ���
    '''���أ��ɹ�����0������ʧ�ܷ���1������ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncDownloadFile = 1
    '��FTP����
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        FuncSetVirtualPath strVirtualPath
           
        '����Ftp�����ȡ�ļ�
        bResult = FtpGetFile(hConnection, Replace(strRemoteFileName, "//", "/"), strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, IIf(blnIsForceRead, &H80000000, 2), 0)
    
        If bResult = False Then
            '������ʱʧ�ܽ�����������
'            Debug.Print "zlRichEpr.cFTP(" & Now & "): DownLoad File Time Out."
            bResult = FtpGetFile(hConnection, Replace(strRemoteFileName, "//", "/"), strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, IIf(blnIsForceRead, &H80000000, 2), 0)
            
            If bResult Then
                MsgBox "ͼ�� " & strRemoteFileName & " ���س�ʱ���������سɹ���", vbInformation, "��ʾ"
            Else
                MsgBox "ͼ�� " & strRemoteFileName & " ���س�ʱ����������ʧ�ܡ�", vbInformation, "��ʾ"
            End If
        End If
        
        If bResult = False Then
            '���ݴ���
            '��ǰ��������δ������Ŀ¼ʱ��ͼ����ܱ��洢��ftp�ĸ�Ŀ¼��
            FuncSetVirtualPath ""
            bResult = FtpGetFile(hConnection, Replace(strRemoteFileName, "//", "/"), strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, IIf(blnIsForceRead, &H80000000, 2), 0)
        End If
    End If
    
    FuncDownloadFile = IIf(bResult, 0, 2)
End Function

Public Function FuncDownLoadFiles(strVirtualPath As String, strLocalPath As String, strRemoteFileNames() As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ���FTP��������ָ��Ŀ¼�����ؼ����ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strLocalPath---����·�����硰C:\��
    '''������strRemoteFileNames()---��Ҫ���ص��ļ�������
    '''���أ��ɹ�����0������ʧ�ܷ���1�������ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDownLoadFiles = 1
    
    '��FTP����
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        FuncSetVirtualPath strVirtualPath
        
        For i = 1 To UBound(strRemoteFileNames)
            bResult = FtpGetFile(hConnection, Replace(strRemoteFileNames(i), "//", "/"), strLocalPath & "\" & strRemoteFileNames(i), False, FILE_ATTRIBUTE_NORMAL, 2, 0)
            If bResult = False Then
                FuncDownLoadFiles = 2
                Exit Function
            End If
        Next
    End If
    
    FuncDownLoadFiles = 0
End Function


Public Function FuncUpLoadFiles(strVirtualPath As String, strLocalPath As String, strLocalFileNames() As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''����: ������Ŀ¼�м����ļ�ͬʱ�ϴ���FTP��������ָ��Ŀ¼��
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strLocalPath---����·�����硰C:\��
    '''������strLocalFileNames()---��Ҫ�ϴ����ļ�������
    '''���أ��ɹ�����0������ʧ�ܷ���1���ϴ��ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncUpLoadFiles = 1
    
    '��FTP����
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        FuncSetVirtualPath strVirtualPath
 
        For i = 1 To UBound(strLocalFileNames)
            bResult = FtpPutFile(hConnection, strLocalPath & "\" & strLocalFileNames(i), Replace(strLocalFileNames(i), "//", "/"), 2, 0)
            If bResult = False Then
                FuncUpLoadFiles = 2
                Exit Function
            End If
        Next
    End If

    FuncUpLoadFiles = 0
End Function


Public Function FuncDelFile(strVirtualPath As String, strFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''����: ɾ��FTP��������ָ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strFileName()---��Ҫɾ�����ļ���
    '''���أ��ɹ�����0������ʧ�ܷ���1��ɾ���ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncDelFile = 1
    '��FTP����
    
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        FuncSetVirtualPath strVirtualPath
 
        bResult = FtpDeleteFile(hConnection, Replace(strFileName, "//", "/"))
    End If
    
    FuncDelFile = IIf(bResult, 0, 2)
End Function


Public Function FuncReNameFile(ByVal strSourceFile As String, ByVal strNewFile As String) As Long
    Dim bResult As Boolean
    Dim strCurFtpRoot As String
    Dim strCurSourceFile As String
    Dim strCurNewFile As String
    
    FuncReNameFile = 1
    
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
'         '����FTP������Ŀ¼
        Call FuncSetVirtualPath("")
        
        strCurFtpRoot = IIf(mstrFtpRoot = "" Or mstrFtpRoot = "/", "", mstrFtpRoot)
        strCurSourceFile = Replace(strSourceFile, "//", "/")
        strCurSourceFile = Replace(strCurFtpRoot & "/" & strCurSourceFile, "//", "/")
        
        strCurNewFile = Replace(strNewFile, "//", "/")
        strCurNewFile = Replace(strCurFtpRoot & "/" & strCurNewFile, "//", "/")
        
        bResult = FtpRenameFile(hConnection, strCurSourceFile, strCurNewFile)
    End If
    
    FuncReNameFile = IIf(bResult, 0, 2)
End Function


Public Function FuncDelFiles(strVirtualPath As String, strFileNames() As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''����: ɾ��FTP������ͬһĿ¼�µ�һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strFileNames()---��Ҫɾ�����ļ�������
    '''���أ��ɹ�����0������ʧ�ܷ���1��ɾ���ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDelFiles = 1
    '��FTP����
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        FuncSetVirtualPath strVirtualPath
        
        For i = 1 To UBound(strFileNames)
            bResult = FtpDeleteFile(hConnection, Replace(strFileNames(i), "//", "/"))
            If bResult = False Then
                FuncDelFiles = 2
                Exit Function
            End If
        Next
    End If
    
    FuncDelFiles = 0
End Function

Public Function FuncDirFiles(ByVal strVirtualPath As String) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''����: �г�FTP������ͬһĿ¼�µ�һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''���أ��ļ������ļ���֮����|�ָ�
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim pData As WIN32_FIND_DATA
    Dim lngHINet As Long ', hOpen As Long, hConnection As Long
    Dim strTemp As String
    Dim blnRC As Boolean
    Dim strFormatVirtualPath As String
    
    '��ʽ������Ŀ¼��������Ŀ¼ǰʹ����"/\VirtualPath\"����ʽʱ�������ܴ�ftp�ϻ�ȡ�ļ���Ϣ
    Call FuncSetVirtualPath(strVirtualPath)
    
    pData.cFileName = String(MAX_FILENAME, 0)
    FuncDirFiles = ""
    lngHINet = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
    
    If lngHINet > 0 Then
        FuncDirFiles = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        Do While True
            pData.cFileName = String(MAX_FILENAME, 0)
            blnRC = InternetFindNextFile(lngHINet, pData)
            If Not blnRC Then
                Exit Do
            Else
                FuncDirFiles = FuncDirFiles & "|" & Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            End If
            DoEvents
        Loop
        InternetCloseHandle lngHINet
        'ɾ����.���͡�..��Ŀ¼
        If Left(Trim(FuncDirFiles), 1) = "." Then
            FuncDirFiles = Mid(Trim(FuncDirFiles), 6)
        End If
    End If
End Function

Public Function FuncFtpFileExists(ByVal strVirtualPath As String, ByVal strFile As String) As Boolean
'�ж�ftp�ļ��Ƿ����
    Dim pData As WIN32_FIND_DATA
    Dim lngHINet As Long ', hOpen As Long, hConnection As Long
    Dim strTemp As String
    Dim blnRC As Boolean
    
    FuncFtpFileExists = False
    
    Call FuncSetVirtualPath(strVirtualPath)
    
    pData.cFileName = String(MAX_FILENAME, 0)
    lngHINet = FtpFindFirstFile(hConnection, Replace(strFile, "//", "/"), pData, 0, 0)
    
    If lngHINet > 0 Then
        strTemp = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        
        If strTemp = strFile Then FuncFtpFileExists = True
                
        InternetCloseHandle lngHINet
    End If
End Function

Public Function FuncFtpGetFileSize(ByVal strVirtualPath As String, ByVal strFile As String) As Long
'��ȡFTP�ϵ��ļ��Ĵ�С
    Dim pData As WIN32_FIND_DATA
    Dim lngHINet As Long '
    Dim strTemp As String
    Dim strFtpFile As String
    
    FuncFtpGetFileSize = 0
    
    Call FuncSetVirtualPath(strVirtualPath)

    pData.cFileName = String(MAX_FILENAME, 0)
    lngHINet = FtpFindFirstFile(hConnection, Replace(strFile, "//", "/"), pData, 0, 0)
    
    If lngHINet > 0 Then
        strTemp = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        
        If strTemp = strFile Then FuncFtpGetFileSize = pData.nFileSizeLow
                
        InternetCloseHandle lngHINet
    End If
End Function


Private Function FormatVirtualPath(ByVal strVirtualPath As String) As String
'�罫��//20170222/1.23.5674.104/����ʽ�����·����ʽΪ��20170222/1.23.5674.104��
    Dim i As Long
    Dim strPaths() As String
    Dim strTmp As String
    
    FormatVirtualPath = ""
    If Trim$(strVirtualPath) = "" Then Exit Function
    
    strTmp = Trim$(strVirtualPath) & "/"
    
    strTmp = Replace(strTmp, "/\", "/")
    strTmp = Replace(strTmp, "\", "/")
    strTmp = Replace(strTmp, "//", "/")
    strTmp = Replace(strTmp, "//", "/")
    
    strPaths = Split(strTmp, "/")
    strTmp = ""
    For i = 0 To UBound(strPaths)
        If Trim(strPaths(i)) <> "" Then
            If strTmp <> "" Then strTmp = strTmp & "/"
            strTmp = strTmp & strPaths(i)
        End If
    Next i
    
    FormatVirtualPath = strTmp
End Function


Private Function GetRandom(ByVal lngBase As Long) As String
    Dim lngNum As Long
    
    Randomize 99
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function

'��ȡ��������
Public Function getEncryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Long
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim strRandom As String
    Dim strBase As String
        
    i = 0
    
    lngPassWLength = Len(strPassW)
    
    strBase = GetRandom(30)
    strRandom = GetRandom(30)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
     
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassW, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strBase) Xor Asc(strRandom)
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop
    
    getEncryptionPassW = strBase & Join(strTemp, "") & strRandom '���ܺ���ִ�
End Function

'��ȡ��������
Public Function GetDecryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Integer
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim lngBase As Long
    Dim strRandom As String
    Dim strPassSouce As String

    i = 0
    
    strPassSouce = Mid(strPassW, 2, Len(strPassW) - 2)
    lngPassWLength = Len(strPassSouce)
    lngBase = Asc(Mid(strPassW, 1, 1))
    
    strRandom = Right(strPassW, 1)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
    
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassSouce, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strRandom) Xor lngBase
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop

    GetDecryptionPassW = Join(strTemp, "") '���ܺ���ִ�
End Function

Public Function funcDownLoadAllFiles(strVirtualPath As String, strLocalPath As String, Optional blnOverWrite As Boolean = False) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''����: ��ָ��Ŀ¼�����������ļ�
'''������   strVirtualPath---FTP����Ŀ¼����Ҫ���ص��ļ����ڵ�Ŀ¼
'''         strLocalPath -- ����Ŀ¼������֮���ŵ�Ŀ¼
'''         blnOverWrite -- �Ƿ񸲸�ԭ���ļ���True -- ���ǣ�False--������
'''���أ��ɹ�����0��ʧ�ܷ��ش������
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim pData As WIN32_FIND_DATA
    Dim lngFindFirstResult As Long
    Dim lngFindNextResult As Long
    Dim blnDownResult As Boolean
    Dim strFileName As String
    
    On Error GoTo err
    
    funcDownLoadAllFiles = 1
    
    '���FTP�����Ƿ���Ч,��Ч���˳�
    If hConnection = 0 Then
        '��FTP������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        
        '��¼��־
        Call WriteCommLog("funcDownLoadAllFiles", "FTP������Ч", "FTP����ʧ�ܣ��˳�����", 0)
        
        Exit Function
    End If
    
    '�л���FTP����Ŀ¼
    Call FuncSetVirtualPath(strVirtualPath)

    
    '�Ȳ���FTPĿ¼�µ��ļ�
    pData.cFileName = String(MAX_FILENAME, 0)
    lngFindFirstResult = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)


    '��¼��־
    Call WriteCommLog("funcDownLoadAllFiles", "���ҵ�һ���ļ�", "����FTPĿ¼�µ��ļ������ҽ���ǣ�" & IIf(lngFindFirstResult = 0, "û�ҵ���", "�ҵ�����׼�����ء�"), 1)
    
    If lngFindFirstResult > 0 Then
        '���ص�һ���ļ�
        strFileName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        If strFileName <> "." And strFileName <> ".." Then
            '���жϱ����ļ��Ƿ���ڣ��Լ��Ƿ�ǿ�и���
            If blnOverWrite = False Then
                If Dir(strLocalPath & "\" & strFileName) = "" Then
                    '�����ļ������ڣ���������ļ�
                    blnDownResult = FtpGetFile(hConnection, strFileName, strLocalPath & "\" & strFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
                    '��¼��־
                    Call WriteCommLog("funcDownLoadAllFiles", "���ص�һ���ļ�", "�����ص�һ���ļ���" & strFileName & " ,���ؽ����" & IIf(blnDownResult, "�ɹ���", "ʧ�ܡ�"), 1)
                Else
                    '��¼��־
                    Call WriteCommLog("funcDownLoadAllFiles", "���ص�һ���ļ�", "���ļ���" & strFileName & " �Ѿ��ڱ��ش��ڣ�δ���ء�", 1)
                End If
            Else
                '��������ļ�
                blnDownResult = FtpGetFile(hConnection, strFileName, strLocalPath & "\" & strFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
                '��¼��־
                Call WriteCommLog("funcDownLoadAllFiles", "���ص�һ���ļ�", "�����ص�һ���ļ���" & strFileName & " ,���ؽ����" & IIf(blnDownResult, "�ɹ���", "ʧ�ܡ�"), 1)
            End If
        End If
         
        Do While True
             pData.cFileName = String(MAX_FILENAME, 0)
             lngFindNextResult = InternetFindNextFile(lngFindFirstResult, pData)
             
             '��¼��־
             Call WriteCommLog("funcDownLoadAllFiles", "����ʣ����ļ�", "ʣ���ļ����������" & IIf(lngFindNextResult = 0, "û�ҵ���", "�ҵ���"), 2)
             
             If lngFindNextResult = 0 Then
                Exit Do
             Else
                strFileName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
                If strFileName <> "." And strFileName <> ".." Then
                    '���жϱ����ļ��Ƿ���ڣ��Լ��Ƿ�ǿ�и���
                    If blnOverWrite = False Then
                        If Dir(strLocalPath & "\" & strFileName) = "" Then
                            '�����ļ������ڣ���������ļ�
                            blnDownResult = FtpGetFile(hConnection, strFileName, strLocalPath & "\" & strFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
                            '��¼��־
                            Call WriteCommLog("funcDownLoadAllFiles", "����ʣ����ļ�", "����ʣ���ļ���" & strFileName & " ���������" & IIf(blnDownResult, "�ɹ���", "ʧ�ܡ�"), 2)
                        Else
                            '��¼��־
                            Call WriteCommLog("funcDownLoadAllFiles", "����ʣ����ļ�", "���ļ���" & strFileName & " �Ѿ��ڱ��ش��ڣ�δ���ء�", 2)
                        End If
                    Else
                        '��������ļ�
                        blnDownResult = FtpGetFile(hConnection, strFileName, strLocalPath & "\" & strFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
                        '��¼��־
                        Call WriteCommLog("funcDownLoadAllFiles", "����ʣ����ļ�", "����ʣ���ļ���" & strFileName & " ���������" & IIf(blnDownResult, "�ɹ���", "ʧ�ܡ�"), 2)
                    End If
                 End If
             End If
             
             DoEvents
         Loop
         
         InternetCloseHandle lngFindFirstResult
    End If

    funcDownLoadAllFiles = 0
    '��¼��־
    Call WriteCommLog("funcDownLoadAllFiles", "�������", "������ء�", 1)
    Exit Function
err:
    '��¼������־
    Call WriteCommLog("funcDownLoadAllFiles", "����Ŀ¼�е�ͼ��ʧ��", "������룺" & err.Number & " ������Ϣ��" & err.Description & " ����֮ǰ������ص�ͼ�� ��" & strFileName, 0)
    
End Function


Private Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String, lngLogLevel As Long)
'------------------------------------------------
'���ܣ���¼ͨѶ��־
'������ logSubName  --  ������־�ĺ�����
'       logTitle   -- ��־����
'       logDesc   --  ��־����
'       lngLogLevel -- ��־����ͨ����־����ȷ����ǰ��־�Ƿ���Ҫ��¼
'���أ���
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    Dim intHour As Integer
    
    On Error GoTo err



    
    If pblnLogEnable = True Then        '�����˼�¼��־���ż�¼��ǰ����־
        '�ж���־����ȷ��������־�Ƿ���Ҫ��¼
        If plngLogLevel >= lngLogLevel Then
            'ͨ����ǰʱ�䣬������־�ļ�����ÿ����Сʱ����һ����־�ļ�
            intHour = Hour(Time)
            intHour = intHour / 2
            intHour = intHour * 2
            strFileName = pstrLogPath & "\" & Date & "-" & intHour & ".log"
            
            '������־����
            strLog = Now() & " ��־���� " & lngLogLevel & " ���⣺ " & logTitle & vbCrLf & "      ������ " & logSubName & vbCrLf & "     ��־���ݣ�" & logDesc & vbCrLf
            
            '����־�ļ�����¼��־
            Open strFileName For Append As #1
            Print #1, strLog
            Close #1
        End If
    End If
    Exit Sub
err:
    Close #1
End Sub


Private Sub Class_Initialize()
    '���ó�ʼֵ
    mblnLogEnable = False
    mstrLogPath = ""
    mlngLogLevel = 1
End Sub