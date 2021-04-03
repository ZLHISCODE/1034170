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
Attribute VB_Creatable = False
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

Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean


                                ''''''''''''''''''''''''
                                '''''''FTP�ļ�����''''''
                                '''''''''''''''''''''''
    
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean

Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

   
Private Const FILE_ATTRIBUTE_NORMAL = &H80
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''����'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public hOpen As Long
Public hConnection As Long

'��¼��־״̬�ı���
Private mblnLogEnable As Boolean     '�Ƿ�������־
Private mstrLogPath As String        '��־·��
Private mlngLogLevel As Long         '��־����1��Ϊ��ͨ��־��2��Ϊ��ϸ��־

Private mstrIPAddress As String
Private mstrUser As String
Private mstrPswd As String



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
    PassWord = mstrPswd
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
Public Function FuncFtpConnect(strHost As String, strUser As String, strPassWord As String) As Long
    Dim i As Integer
    Dim strDecryptionPassW As String
    
    '��������Ѿ����ܣ�������ܼ�������
    If Len(strPassWord) >= 3 Then
        If Mid(strPassWord, 1, 1) & Mid(strPassWord, 3, 1) & Mid(strPassWord, Len(strPassWord), 1) = "�����" Then
            strDecryptionPassW = Mid(strPassWord, 2)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, Len(strDecryptionPassW) - 1)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, 1) & Mid(strDecryptionPassW, 3)
            strDecryptionPassW = getDecryptionPassW(strDecryptionPassW)
            
            strPassWord = strDecryptionPassW
        End If
    End If
    
    mstrIPAddress = strHost
    mstrUser = strUser
    mstrPswd = strPassWord
    
    hOpen = InternetOpen("WinINet1", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strHost, 0, strUser, strPassWord, INTERNET_SERVICE_FTP, 0, 255)
    FuncFtpConnect = hConnection
End Function

Public Function FuncFtpDisConnect()
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    mstrIPAddress = ""
    mstrUser = ""
    mstrPswd = ""
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
    
'    Dim hOpen As Long
'    Dim hConnection As Long
    Dim i As Integer
    Dim aNestDirs() As String               '�洢������Ҫ������Ŀ¼
    Dim bResult As Boolean                  '�������
        
    FuncFtpMkDir = 1
    
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncFtpMkDir = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
        '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        '��ȡȫ����Ҫ������Ŀ¼��Ϣ
        aNestDirs = Split(strDir, "/")
        '����ȫ��Ŀ¼
        For i = 0 To UBound(aNestDirs)
            If aNestDirs(i) <> vbNullString Then
                bResult = FtpCreateDirectory(hConnection, aNestDirs(i))
                FtpSetCurrentDirectory hConnection, aNestDirs(i)
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
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncFtpDelDir = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        bResult = FtpRemoveDirectory(hConnection, strDir)
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
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncUploadFile = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        bResult = FtpPutFile(hConnection, strLocalFileName, strRemoteFileName, 2, 0)
    End If
    FuncUploadFile = IIf(bResult, 0, 2)
End Function

Public Function FuncDownloadFile(strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''���ܣ�FTP��������ָ��Ŀ¼����һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strLocalFileName---�����ļ�����·�����硰C:\SSS.TXT��
    '''������strRemoteFileName---��Ҫ���ص��ļ���
    '''���أ��ɹ�����0������ʧ�ܷ���1������ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bResult As Boolean
    
    FuncDownloadFile = 1
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDownloadFile = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        bResult = FtpGetFile(hConnection, strRemoteFileName, strLocalFileName, False, FILE_ATTRIBUTE_NORMAL, 2, 0)
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
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDownLoadFiles = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        For i = 1 To UBound(strRemoteFileNames)
            bResult = FtpGetFile(hConnection, strRemoteFileNames(i), strLocalPath & "\" & strRemoteFileNames(i), False, FILE_ATTRIBUTE_NORMAL, 2, 0)
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
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncUpLoadFiles = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        For i = 1 To UBound(strLocalFileNames)
            bResult = FtpPutFile(hConnection, strLocalPath & "\" & strLocalFileNames(i), strLocalFileNames(i), 2, 0)
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
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDelFile = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        bResult = FtpDeleteFile(hConnection, strFileName)
    End If
    FuncDelFile = IIf(bResult, 0, 2)
End Function


Public Function FuncDelFiles(strVirtualPath As String, strFileNames() As String) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''����: ɾ��FTP������ͬһĿ¼�µ�һ���ļ�
    '''������strVirtualPath---FTP����Ŀ¼
    '''������strFileNames()---��Ҫɾ�����ļ�������
    '''���أ��ɹ�����0������ʧ�ܷ���1��ɾ���ļ�ʧ�ܷ���2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    Dim bResult As Boolean
    
    FuncDelFiles = 1
    If hConnection = 0 Then
        '��ftp������������ʧ�ܣ���ʾ���Ӵ�����Ϣ
        InternetCloseHandle hConnection
        InternetCloseHandle hOpen
        FuncDelFiles = 1
        Exit Function
    Else               '��FTP�����������ӳɹ�
         '����FTP������Ŀ¼
        If strVirtualPath <> vbNullString Then
           FtpSetCurrentDirectory hConnection, strVirtualPath
        End If
        For i = 1 To UBound(strFileNames)
            bResult = FtpDeleteFile(hConnection, strFileNames(i))
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
    Dim lngHINet As Long
    Dim strTemp As String
    Dim blnRC As Boolean
    
    pData.cFileName = String(MAX_FILENAME, 0)
    FuncDirFiles = ""
    lngHINet = FtpFindFirstFile(hConnection, strVirtualPath & "/*.*", pData, 0, 0)
    
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
    If strVirtualPath <> vbNullString Then
        FtpSetCurrentDirectory hConnection, strVirtualPath
    End If
    
    '�Ȳ���FTPĿ¼�µ��ļ�
    pData.cFileName = String(MAX_FILENAME, 0)
    lngFindFirstResult = FtpFindFirstFile(hConnection, strVirtualPath & "\*.*", pData, 0, 0)
    
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

Private Sub Class_Initialize()
    
    '���ó�ʼֵ
    mblnLogEnable = False
    mstrLogPath = ""
    mlngLogLevel = 1
    
End Sub

Private Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As Object
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    Set objFile = CreateObject("Scripting.FileSystemObject")
    
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

    If mblnLogEnable = True Then        '�����˼�¼��־���ż�¼��ǰ����־
        '�ж���־����ȷ��������־�Ƿ���Ҫ��¼
        If mlngLogLevel >= lngLogLevel Then
            'ͨ����ǰʱ�䣬������־�ļ�����ÿ����Сʱ����һ����־�ļ�
            intHour = Hour(Time)
            intHour = intHour / 2
            intHour = intHour * 2
            strFileName = mstrLogPath & "\" & Date & "-" & intHour & ".log"

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
Public Function getDecryptionPassW(ByVal strPassW As String) As String
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

    getDecryptionPassW = Join(strTemp, "") '���ܺ���ִ�
End Function