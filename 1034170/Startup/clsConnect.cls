VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************
'                һ��API�볣��
'**************************************************************************
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'���ܣ�������Internet�ĻỰ
'˵����
'    sAgent--Ҫ����Internet�Ի���Ӧ�ó�����
'    lAccessType--�����������ʵ�����
'��ע�����lAccessType����ΪINTERNET_OPEN_TYPE_PRECONFIG������ʱ��Ҫ����
'    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings
'    ע���·���µ�ע�����ֵProxyEnable��ProxyServer�� ProxyOverride
'    sProxyName--ָ�����������������֣�������������ΪINTERNET_OPEN_TYPE_PROXY����Ч
'    sProxyBypass--ָ�����������������ֻ��ַ�������ô���ʱlpszProxyNameָ���Ľ�ʧЧ
'��������ֵ�������������ʧ�ܣ�lngINet Ϊ0��
'    lFlags-
Private Const INTERNET_FLAG_ASYNC           As Long = &H10000000    'ʹ���첽����
Private Const INTERNET_FLAG_FROM_CACHE      As Long = &H1000000
Private Const INTERNET_FLAG_OFFLINE         As Long = &H1000000     'ֻͨ�����û���������ز���
'   lAccessType
Private Const INTERNET_OPEN_TYPE_PRECONFIG  As Long = &H0           ' use registry configuration
Private Const INTERNET_OPEN_TYPE_DIRECT     As Long = &H1           'direct to net

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'���ܣ�����Internet���ӣ���FTP�Ự
'˵����
'    hInternetSession--����InternetOpen���ص�Internet�Ự���
'    sServerName--Ҫ���ӵķ����������ƻ�IP
'    nServerPort--Ҫ���ӵ�Internet�˿�
'    sUsername--��¼���û��ʺ�
'    sPassword--��¼�Ŀ���
'    lService--Ҫ���ӵķ��������ͣ�����������FTP�����������ӵ�����Ϊ����INTERNET_SERVICE_FTP��
'    lFlags--�������x8000000�����ӽ�ʹ�ñ���FTP���壬����0ʹ�÷Ǳ�������
'    lContext--��ʹ�ûص�����ʱʹ�øò�������ʹ�ûص����񴫵�0
'��������ֵ�������������ʧ�ܣ�lngINetConn Ϊ0
'    lService
Private Const INTERNET_SERVICE_FTP          As Long = &H1
'    lFlags
Private Const INTERNET_FLAG_KEEP_CONNECTION  As Long = &H400000    ' use keep-alive semantics
Private Const INTERNET_FLAG_PASSIVE         As Long = &H8000000   ' used for FTP connections

Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
'���ܣ���FTP������������һ���ļ�
'˵����
'    hFtpSession--����InternetConnect���ص�Internet���Ӿ��
'    lpszRemoteFile--��Ҫ��õ�FTP�������ϵ��ļ���
'    lpszNewFile--Ҫ�����ڱ��ػ����е��ļ���
'    fFailIfExists--0���滻�����ļ�����1 ����������ļ��Ѿ����������ʧ�ܣ���
'    dwFlagsAndAttributes--����ָ�������ļ����ļ����ԣ�����0����
'    dwFlags--�ļ��Ĵ��䷽ʽ���ܰ�������ֵ��
'        FTP_TRANSFER_TYPE_ASCII                   1           ��ASCII �����ļ���A�ഫ�䷽����
'        FTP_TRANSFER_TYPE_BINARY                 2           �ö����ƴ����ļ���B�ഫ�䷽����
'    dwContext--Ҫȡ�ص��ļ�����������ʶ��
'��������ֵ�������������ʧ�ܣ�blnRC ΪFALSE
Private Const FTP_TRANSFER_TYPE_BINARY      As Long = &H2
Private Const INTERNET_FLAG_RELOAD          As Long = &H80000000  'retrieve the original item
Private Const INTERNET_FLAG_NO_CACHE_WRITE  As Long = &H4000000
Private Const INTERNET_FLAG_DONT_CACHE      As Long = INTERNET_FLAG_NO_CACHE_WRITE
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
' ����ָ����·��
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime                           As Long
    dwHighDateTime                          As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes                        As Long
    ftCreationTime                          As FILETIME
    ftLastAccessTime                        As FILETIME
    ftLastWriteTime                         As FILETIME
    nFileSizeHigh                           As Long
    nFileSizeLow                            As Long
    dwReserved0                             As Long
    dwReserved1                             As Long
    cFileName                               As String * MAX_PATH
    cAlternate                              As String * 14
End Type

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'���ܣ��ر�Internet����
'˵����
'hInet--Ҫ�رյĻỰ��InternetOpen�������ӣ�InternetConnect�����
'��������ֵ��
'��������
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE_STRING, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lplngEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal lngEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal lngEnum As Long) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

Private Type NETRESOURCE
    dwScope                                 As Long
    dwType                                  As Long
    dwDisplayType                           As Long
    dwUsage                                 As Long
    pLocalName                              As Long
    pRemoteName                             As Long
    pComment                                As Long
    pProvider                               As Long
End Type
'string�汾��NETRESOURCE�����㴫��
Private Type NETRESOURCE_STRING
    dwScope                                 As Long
    dwType                                  As Long
    dwDisplayType                           As Long
    dwUsage                                 As Long
    pLocalName                              As String
    pRemoteName                             As String
    pComment                                As String
    pProvider                               As String
End Type

Private Type NETRESOURCE_REAL
    dwScope                                 As Long
    dwType                                  As Long
    dwDisplayType                           As Long
    dwUsage                                 As Long
    sLocalName                              As String
    sRemoteName                             As String
    sComment                                As String
    sProvider                               As String
End Type
Private Const RESOURCE_CONNECTED            As Long = &H1
Private Const RESOURCE_ENUM_ALL             As Long = &HFFFF
Private Const RESOURCE_GLOBALNET            As Long = &H2
Private Const RESOURCETYPE_DISK             As Long = &H1
Private Const RESOURCEDISPLAYTYPE_SHARE     As Long = &H3
Private Const RESOURCEUSAGE_CONNECTABLE     As Long = &H1
Private Const CONNECT_UPDATE_PROFILE        As Long = &H1
Private Const NO_ERROR                      As Long = 0
'**************************************************************************
'                ����ģ�鼶����
'**************************************************************************
'==========================================================================
'                1��ģ�����
'==========================================================================
Private mlngINet                            As Long                 'FTP��ʽ����������
Private mlngINetConn                        As Long                 'FTP����
'Private mobjAPI As APILoad 'API��̬������
Private mcllInsPath                         As Collection
Private mintConnTimes                       As Integer              '�������Ӵ�����
Private mstrComputerName                    As String
Private mobjFSO                             As New FileSystemObject
'==========================================================================
'                2��ö�١��ṹ��
'==========================================================================
Public Enum ServerConnectType
    SCT_NONE = -1 '������
    SCT_Share = 0 '�����ļ�����
    SCT_FTP = 1 'FTP����
End Enum

'***************************************************************************
'                ����������ĸ������Ե�ģ�鼶�������Լ����Զ���
'***************************************************************************
'===========================================================================
'                1���������Եı���
'===========================================================================
'ConnectType��ǰ��������(R/W)
Private msctCur                             As ServerConnectType
'ServerPath��ǰ������λ��(R/W)
Private mstrCurServerPath                   As String
'ConnUser���ӵ�ǰ���������û���(R/W)
Private mstrCurConnUser                     As String
'ConnPassword���ӵ�ǰ������������(R/W)
Private mstrCurConnPWD                      As String
'ServerPort���ӵ�ǰ�������Ķ˿�(R/W)������FTP��ʽ
Private mstrCurServerPort                   As String
'CollectType�ռ����ļ�����(R/W)�������ռ��ļ�
Private mstrCollectType                     As String
'ISOpen�����Ƿ��(R)
Private mblnOpen                            As Boolean
'===========================================================================
'                2�����Զ���
'===========================================================================
'ConnectType��ǰ��������(R/W)
Public Property Let ConnectType(ByVal sctCur As ServerConnectType)
    msctCur = sctCur
End Property

Public Property Get ConnectType() As ServerConnectType
    ConnectType = msctCur
End Property

'ServerPath��ǰ������λ��(R/W)
Public Property Let ServerPath(ByVal strCurServerPath As String)
    mstrCurServerPath = strCurServerPath
End Property

Public Property Get ServerPath() As String
    ServerPath = mstrCurServerPath
End Property

'ConnUser���ӵ�ǰ���������û���(R/W)
Public Property Let ConnUser(ByVal strCurConnUser As String)
    mstrCurConnUser = strCurConnUser
End Property

Public Property Get ConnUser() As String
    ConnUser = mstrCurConnUser
End Property
'ConnPassword���ӵ�ǰ������������(R/W)
Public Property Let ConnPassword(ByVal strCurConnPWD As String)
    mstrCurConnPWD = strCurConnPWD
End Property

Public Property Get ConnPassword() As String
    ConnPassword = mstrCurConnPWD
End Property

'ServerPort���ӵ�ǰ�������Ķ˿�(R/W)������FTP��ʽ
Public Property Let ServerPort(ByVal strCurServerPort As String)
    mstrCurServerPort = strCurServerPort
End Property

Public Property Get ServerPort() As String
    ServerPort = mstrCurServerPort
End Property
'CollectType�ռ����ļ�����(R/W)�������ռ��ļ�
Public Property Let CollectType(ByVal strCollectType As String)
    mstrCollectType = strCollectType
End Property

Public Property Get CollectType() As String
    CollectType = mstrCollectType
End Property
'ISOpen�����Ƿ��(R)
Public Property Get IsOpen() As String
    IsOpen = mblnOpen
End Property

'***************************************************************************
'                �ġ����ⲿ�����ķ���
'***************************************************************************
Public Function GetFileConnect(ByVal strComputerName As String) As Boolean
'���ܣ���ȡ�������ļ�����
    Dim objConn As New clsConnect
    Dim sctConnType As ServerConnectType
    Dim strServerID As String, strServer As String, strUser As String, strPWD As String, strPort As String, strCollectType As String
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim blnDefalut As Boolean, blnConnOK As Boolean
    
    On Error Resume Next
    strSql = "Select �����ļ������� From ZLTools.zlClients Where ����վ=[1]"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ�������������", strComputerName)
    If Err.Number = 0 Then
        If Not rsTmp.EOF Then strServerID = rsTmp!�����ļ������� & ""
    Else
        Err.Clear
        Exit Function
    End If
    On Error GoTo errH
    If strServerID <> "" Then
        strSql = "Select ���,����, λ��, �û���, ����, �˿�,Nvl(�Ƿ�ȱʡ,0) �Ƿ�ȱʡ , ���� From Zltools.Zlupgradeserver Where ��� = [1]"
        Set rsTmp = OpenSQLRecord(strSql, "��ȡ����������", Val(strServerID))
        If Not rsTmp.EOF Then
            strServerID = rsTmp!��� & ""
            sctConnType = IIf(rsTmp!���� = 0, SCT_Share, SCT_FTP)
            strServer = rsTmp!λ��
            strUser = rsTmp!�û���
            strPWD = DeCipher(rsTmp!���� & "")
            strPort = rsTmp!�˿� & ""
            blnDefalut = rsTmp!�Ƿ�ȱʡ = 1
        Else
            strServerID = ""
        End If
    End If
    If strServerID <> "" Then
        blnConnOK = Me.ToConnect(sctConnType, strServer, strUser, strPWD, strPort, strCollectType)
    End If
    '���Ӳ��ɹ��������������Զ�����Ĭ�Ϸ�����
    If Not blnConnOK And Not blnDefalut Then
        strSql = "Select ���,����, λ��, �û���, ����, �˿�, ���� From Zltools.Zlupgradeserver Where Nvl(�Ƿ�ȱʡ,0) = 1"
        Set rsTmp = OpenSQLRecord(strSql, "��ȡĬ������������")
        If Not rsTmp.EOF Then
            strServerID = rsTmp!��� & ""
            sctConnType = IIf(rsTmp!���� = 0, SCT_Share, SCT_FTP)
            strServer = rsTmp!λ��
            strUser = rsTmp!�û���
            strPWD = DeCipher(rsTmp!���� & "")
            strPort = rsTmp!�˿� & ""
            blnConnOK = Me.ToConnect(sctConnType, strServer, strUser, strPWD, strPort)
        End If
    End If

    If blnConnOK Then
        GetFileConnect = True
    Else
        GetFileConnect = GetFileConnectOld
    End If
    Exit Function
errH:
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "����ļ����������ó�����" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Private Function GetFileConnectOld() As Boolean
'���ܣ���ȡ�ļ����������ӣ��Ϸ�ʽ
'������blnUpgrade=True-Ԥ���������������� ��false-�ļ��ռ�������
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim sctConnType As ServerConnectType, strServerID As String
    Dim objConn As New clsConnect
    Dim arrParas() As Variant, arrValues(4) As String
    Dim strSQLPars As String, i As Integer
    Dim blnReadOk As Boolean, blnConnOK As Boolean, blnGo As Boolean
    
    On Error GoTo errH
    '��ȡ��������
    sctConnType = SCT_Share
    strSql = "Select ��Ŀ,���� From ZLTools.zlregInfo where ��Ŀ=[1]"
    Set rsTmp = OpenSQLRecord(strSql, "��������", "��������")
    If Not rsTmp.EOF Then
        If NVL(rsTmp!����, 0) = 1 Then sctConnType = SCT_FTP
    End If
    '��ȡ������ID
    strSql = "Select ����������,FTP������ From ZLTools.zlClients Where ����վ=[1]"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ�������������", mstrComputerName)
    If Not rsTmp.EOF Then strServerID = IIf(sctConnType = SCT_FTP, rsTmp!FTP������ & "", rsTmp!���������� & "")
    '��ȡ��������Ϣ
    If sctConnType = SCT_FTP Then
        arrParas = Array("FTP������", "FTP�û�", "FTP����", "FTP�˿�", "")
    Else
        arrParas = Array("������Ŀ¼", "�����û�", "��������", "", "")
    End If
ReGetParas:
    '�Ȼ�ȡSQL����
    strSQLPars = ""
    For i = LBound(arrParas) To UBound(arrParas)
        If arrParas(i) <> "" Then
            strSQLPars = strSQLPars & ",'" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
        End If
    Next
    strSQLPars = Mid(strSQLPars, 2)
    strSql = "Select ��Ŀ,���� From ZLTools.zlregInfo where ��Ŀ in(" & strSQLPars & ")"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ������")
    If Not rsTmp.EOF Then
        For i = LBound(arrParas) To UBound(arrParas)
            If arrParas(i) <> "" Then
                rsTmp.Filter = "��Ŀ='" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
                If Not rsTmp.EOF Then arrValues(i) = rsTmp!���� & ""
            End If
        Next
    End If
    
    blnReadOk = True
    '���������û�������Ϊ�գ����ܽ����ռ�������
    If arrValues(0) = "" Or arrValues(1) = "" Or arrValues(2) = "" Then
        blnReadOk = False
    'FTP��ʽ��Ҫһ���˿�
    ElseIf sctConnType = SCT_FTP And arrValues(3) = "" Then
        blnReadOk = False
    End If
    If blnReadOk Then
        blnConnOK = Me.ToConnect(sctConnType, arrValues(0), arrValues(1), arrValues(2), arrValues(3), arrValues(4))
    End If
    If (Not blnConnOK Or Not blnReadOk) Then
        If strServerID <> "" And strServerID <> "0" Then
            strServerID = "0"
            GoTo ReGetParas '���»�ȡ���ӷ������Ĳ���
        ElseIf (strServerID = "0" Or strServerID = "") And Not blnGo Then
            blnGo = True '��ֹѭ��
            strServerID = IIf(strServerID = "0", "", "0")
            GoTo ReGetParas '���»�ȡ���ӷ������Ĳ���
        End If
    End If
    GetFileConnectOld = blnConnOK
    Exit Function
errH:
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "��ȡ��ǰ��ʽ��ȱʡ������������" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Public Function ToConnect(ByVal sctConnectType As ServerConnectType, ByVal strPath As String, strUser As String, strUserPwd As String, Optional ByVal strPort As String, Optional ByVal strCollectType As String) As Boolean
'���ܣ�����һ������
'������intConnectType=�������ͣ�0-���繲�����ӣ�1-Ftp����
'      strPath=��Դλ��
'      strUser\strUserPwd=������Դ���û���������
'      strPort=������Դ�Ķ˿�
'      strCollectType=�ռ��ļ�����
'���أ��Ƿ����ӳɹ�
    Dim nrCur As NETRESOURCE_STRING
    Dim lngRet As Long
    
    mintConnTimes = 0
    mblnOpen = False
    If sctConnectType = SCT_Share Then '���������Ƿ�����������
        If Not mobjFSO.FolderExists(strPath) Then
            With nrCur
                .dwScope = RESOURCE_GLOBALNET
                .dwType = RESOURCETYPE_DISK
                .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
                .dwUsage = RESOURCEUSAGE_CONNECTABLE
                .pLocalName = "" 'ӳ���������
                .pRemoteName = strPath  '������·��
            End With
            On Error GoTo errH
            lngRet = WNetAddConnection2(nrCur, strUserPwd, strUser, CONNECT_UPDATE_PROFILE)
            If lngRet <> NO_ERROR Then
                '������������һ���û�ʹ��һ�������û����������������Դ�Ķ������ӡ��ж���˷�����������Դ���������ӣ�Ȼ������һ�Ρ�
                Exit Function
            End If
        End If
    Else '����Ƿ�����������FTP������
        mlngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If mlngINet <= 0 Then
            Exit Function
        End If
        '����������������
        mlngINetConn = InternetConnect(mlngINet, strPath, strPort, strUser, strUserPwd, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_PASSIVE, 0)
        If mlngINetConn = 0 Then
            Call InternetCloseHandle(mlngINet)
            Exit Function
        End If
    End If
    Me.ConnectType = sctConnectType
    Me.ServerPath = strPath
    Me.ConnUser = strUser
    Me.ConnPassword = strUserPwd
    Me.ServerPort = strPort
    Me.CollectType = strCollectType
    mblnOpen = True
    ToConnect = True
    Exit Function
errH:
    If Not ReConnect And mintConnTimes <= 3 Then
        ToConnect = ReConnect
    Else
        ToConnect = False
    End If
End Function

Public Function ReConnect() As Boolean
'���ܣ�����һ������
'������intConnectType=�������ͣ�0-���繲�����ӣ�1-Ftp����
'      strPath=��Դλ��
'      strUser\strUserPwd=������Դ���û���������
'      strPort=������Դ�Ķ˿�
'      strCollectType=�ռ��ļ�����
'���أ��Ƿ����ӳɹ�
    Dim nrCur As NETRESOURCE_STRING
    Dim lngRet As Long
    '�л�����3�Σ������л���ֱ�ӷ���True
    If mintConnTimes > 3 Then
        ReConnect = True: Exit Function
    End If
    mblnOpen = False
    mintConnTimes = mintConnTimes + 1
    Call CloseConnect
    If Me.ConnectType = SCT_Share Then '���������Ƿ�����������
        If Not mobjFSO.FolderExists(Me.ServerPath) Then
            With nrCur
                .dwScope = RESOURCE_GLOBALNET
                .dwType = RESOURCETYPE_DISK
                .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
                .dwUsage = RESOURCEUSAGE_CONNECTABLE
                .pLocalName = "" 'ӳ���������
                .pRemoteName = Me.ServerPath  '������·��
            End With
            On Error GoTo errH
            lngRet = WNetAddConnection2(nrCur, Me.ConnUser, Me.ConnPassword, CONNECT_UPDATE_PROFILE)
            If lngRet <> NO_ERROR Then
                '������������һ���û�ʹ��һ�������û����������������Դ�Ķ������ӡ��ж���˷�����������Դ���������ӣ�Ȼ������һ�Ρ�
                Exit Function
            End If
        End If
    Else '����Ƿ�����������FTP������
        mlngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        If mlngINet <= 0 Then
            Exit Function
        End If
        '�Զ��л�������FTP���壬ͬʱ����������������
        mlngINetConn = InternetConnect(mlngINet, Me.ServerPath, Me.ServerPort, Me.ConnUser, Me.ConnPassword, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or IIf(mintConnTimes Mod 2 = 1, INTERNET_FLAG_PASSIVE, 0), 0)
        If mlngINetConn = 0 Then
            Call InternetCloseHandle(mlngINet)
            Exit Function
        End If
    End If
    mblnOpen = True
    ReConnect = True
    Exit Function
errH:
    ReConnect = False
End Function

Public Function CloseConnect() As Boolean
'���ܣ��ر�����
'      strPath=�ر��������ӵ�·��
    Dim lngReturn As Long
    mblnOpen = False
    On Error Resume Next
    If Me.ConnectType = SCT_Share Then
        lngReturn = WNetCancelConnection(Me.ServerPath, CONNECT_UPDATE_PROFILE, True)
        If lngReturn <> NO_ERROR Then
        Else
            CloseConnect = True
        End If
    ElseIf Me.ConnectType = SCT_FTP Then
        If mlngINetConn <> 0 Then
           InternetCloseHandle mlngINetConn
        End If
    
        If mlngINet <> 0 Then
           InternetCloseHandle mlngINet
        End If
        CloseConnect = True
    End If
End Function

Public Function DownloadFile(ByVal strFileName As String, ByVal strDownPath As String) As Boolean
'���ܣ���������
'������strFileName=���ļ����ļ���,\
'      strDownPath=���ص�ַ
'���أ������Ƿ�ɹ���
    Dim blnDowload As Boolean
    Dim lngReturn As Long, strErr As String
    '�ȼ��������ļ��Ƿ���ڣ����������˳�
    If Not IsServerFileExists(strFileName) Then Exit Function
    '���з������ļ�����
    On Error Resume Next
    '�����������ļ�
    If mobjFSO.FileExists(strDownPath) Then
        If FileSystem.GetAttr(strDownPath) <> vbNormal Then
             Call FileSystem.SetAttr(strDownPath, vbNormal)
        End If
        Call mobjFSO.DeleteFile(strDownPath)
    End If
    
    If Err.Number <> 0 Then Err.Clear
    If Me.ConnectType = SCT_FTP Then
        DoEvents
        lngReturn = FtpGetFile(mlngINetConn, strFileName, strDownPath, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
        blnDowload = lngReturn <> 0 '��������������ֱ��ͨ�����ز����ͣ��ò����͵�ΪTrueʱ��not ��ΪTrue,ԭ��δ֪
        If Not blnDowload Then
            strErr = GetLastDllErr
            If ReConnect() Then
                DoEvents
                lngReturn = FtpGetFile(mlngINetConn, strFileName, strDownPath, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
                blnDowload = lngReturn <> 0 '��������������ֱ��ͨ�����ز����ͣ��ò����͵�ΪTrueʱ��not ��ΪTrue,ԭ��δ֪
                If Not blnDowload Then strErr = GetLastDllErr
            End If
        End If
    Else
        DoEvents
        Call mobjFSO.CopyFile(Me.ServerPath & "\" & strFileName, strDownPath, True)
        If Err.Number <> 0 Then
            strErr = Err.Description
            Err.Clear
            blnDowload = False
        Else
            blnDowload = True
        End If
    End If
    DownloadFile = blnDowload
End Function

Public Function IsServerFileExists(ByVal strFileName As String) As Boolean
'���ܣ��жϷ������ļ��Ƿ����
    Dim udtFindData As WIN32_FIND_DATA, lnghFile    As Long
    Dim strFile    As String

    If Me.ConnectType = SCT_Share Then
        IsServerFileExists = mobjFSO.FileExists(Me.ServerPath & "\" & strFileName)
    Else
        lnghFile = FtpFindFirstFile(mlngINetConn, strFileName, udtFindData, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0)
        strFile = Left(udtFindData.cFileName, InStr(udtFindData.cFileName, vbNullChar) - 1)
        IsServerFileExists = strFile <> ""
        If lnghFile <> 0 Then
            Call InternetCloseHandle(lnghFile)
        End If
    End If
End Function

Private Function CloseNetConn(Optional strIp As String = "", Optional strComputerName As String = "") As Boolean
'���ܣ���ֹ����������Դ����
'���أ���ֹ�ɹ�,����true,���򷵻�False
'˵����ֻҪ��һ������û��ɹ�,��Ҳ����false,���򷵻�true
    Dim lngReturn           As Long, lngEnum        As Long, lngCount           As Long
    Dim lngLength           As Long, i              As Long, lngBufferSize      As Long
    Dim uNetApi(0 To 256)   As NETRESOURCE
    Dim uNet()              As NETRESOURCE_REAL
    Dim blnReturn           As Boolean
    
    On Error Resume Next
    lngReturn = WNetOpenEnum(RESOURCE_CONNECTED, 0, RESOURCEUSAGE_CONNECTABLE, ByVal 0&, lngEnum)
    If lngReturn = NO_ERROR Then
        lngCount = RESOURCE_ENUM_ALL
        lngBufferSize = UBound(uNetApi) * Len(uNetApi(0)) / 2
        lngReturn = WNetEnumResource(lngEnum, lngCount, uNetApi(0), lngBufferSize)
        If lngCount > 0 Then
            ReDim Preserve uNet(0 To lngCount - 1) As NETRESOURCE_REAL
            For i = 0 To lngCount - 1
                'Each   Resource   will   appear   here   as   uNet(i)
                uNet(i).dwScope = uNetApi(i).dwScope
                uNet(i).dwType = uNetApi(i).dwType
                uNet(i).dwDisplayType = uNetApi(i).dwDisplayType
                uNet(i).dwUsage = uNetApi(i).dwUsage
                If uNetApi(i).pLocalName <> 0 Then
                    lngLength = lstrlen(uNetApi(i).pLocalName)
                    uNet(i).sLocalName = Space$(lngLength)
                    CopyMemory ByVal uNet(i).sLocalName, ByVal uNetApi(i).pLocalName, lngLength
                End If
                If uNetApi(i).pRemoteName <> 0 Then
                    lngLength = lstrlen(uNetApi(i).pRemoteName)
                    uNet(i).sRemoteName = Space$(lngLength)
                    CopyMemory ByVal uNet(i).sRemoteName, ByVal uNetApi(i).pRemoteName, lngLength
                End If
                If uNetApi(i).pComment <> 0 Then
                    lngLength = lstrlen(uNetApi(i).pComment)
                    uNet(i).sComment = Space$(lngLength)
                    CopyMemory ByVal uNet(i).sComment, ByVal uNetApi(i).pComment, lngLength
                End If
                If uNetApi(i).pProvider <> 0 Then
                    lngLength = lstrlen(uNetApi(i).pProvider)
                    uNet(i).sProvider = Space$(lngLength)
                    CopyMemory ByVal uNet(i).sProvider, ByVal uNetApi(i).pProvider, lngLength
                End If
            Next
        End If
    End If
    
    If lngEnum > 0 Then
        lngReturn = WNetCloseEnum(lngEnum)
    End If
    If lngCount = 0 Then
        CloseNetConn = True
        Exit Function
    End If
    
    '��������
    blnReturn = True
    For i = 0 To UBound(uNet)
        lngReturn = WNetCancelConnection(IIf(uNet(i).sLocalName = "", uNet(i).sRemoteName, uNet(i).sLocalName), CONNECT_UPDATE_PROFILE, True)
        If lngReturn <> NO_ERROR Then
            blnReturn = False
        End If
    Next
    CloseNetConn = blnReturn
End Function