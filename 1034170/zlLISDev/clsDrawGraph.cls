VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrawGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DrawImg(ByVal strType As String, ByVal strData As String, ByVal strFileName As String, _
                        Optional ByVal intSaveType As Integer) As Boolean
    '�ⲿ���ýӿ�һ
    '���ܣ����ݴ����ͼ�����ݣ���ͼ
    DrawImg = DrawImgAndSaveFile(strType, strData, strFileName, intSaveType)
End Function

Public Function GetSampleImages(ByVal lngSampleID As Long, ByVal strPath As String, ByVal blnClearOldFile As Boolean, _
                                ByRef strErr As String, Optional intSaveImgType As Integer = 0, Optional intVer As Integer = 0) As String
        '��ȡһ���걾������ͼƬ��
        '
        'lngsampleid : �걾ID
        'strPath     : ͼƬ����·��,���������Զ�����
        'blnClearOldFile :�Ƿ���ջ����ڱ��ص�ͼ���ļ�,True��ÿ�ζ������ݿ���ļ����浽����;False-��һ�ε���ʱ�����ݿ��ͼ�β���ͼƬ��֮��ֱ��ʹ��
        'strErr      : ��������ֵΪ�մ�ʱ�����ص���ʾ��Ϣ
        'intImgType  : ���ص�ͼƬ�ļ���ʽ��0��cht(Ĭ��),1-jgp,2-png
        'intVer      : ���°�LIS�����ϰ�LIS�ڵ��ñ��������� 0-�ϰ�LIS��Ĭ�ϣ��ӡ�����ͼ��������ȡͼ�����ݣ���1-�°�LIS���ӡ����鱨��ͼ����ȡͼ�����ݣ�
    
        Dim rsTmp As New ADODB.Recordset
        Dim DrawIndex As Integer, strFileType As String
        Dim StrTime As Date
        Dim strSavePath As String
        Dim strReturn As String
        Dim strFileName As String
        Dim objStream As textStream
        
        
        Dim blnFtp As Boolean       'FTP�Ƿ����
        Static strFtpPara As String       '����FTP����
        Dim strFtpUser As String, strFtpPass As String, strFtpIP As String, strFtpDir As String
        Dim strFtpPath As String
        
        Dim intLayOut As Integer, strImageType As String
        Dim strImageData As String, strLocalFile As String, strFileNames() As String
        Dim strFileAll() As String, strFtpPathName As String, intLoop As Integer
        Dim blnDown As Boolean
         
        ReDim Preserve strFileNames(0)
        ReDim Preserve strFileAll(0)
        
100     strReturn = ""
102     If Not gblnInit Then
104         strErr = "���ñ�����ǰ�����ȵ��ó�ʼ��������"
            Exit Function
        End If
    
106     If intSaveImgType = 1 Then
108         strFileType = ".jpg"
110     ElseIf intSaveImgType = 2 Then
112         strFileType = ".png"
        Else
114         strFileType = ".cht"
        End If
116     strSavePath = strPath
118     If Trim(strSavePath) = "" Then
120         strErr = "��ָ������ͼƬ��·����"
            Exit Function
122     ElseIf Not gobjFSO.FolderExists(strSavePath) Then
124         strErr = strSavePath & "�����ڣ�"
            Exit Function
        End If
126     If Right(strSavePath, 1) = "\" Then strSavePath = Left(strSavePath, Len(strSavePath) - 1)
        '---------------------------------------------------
        On Error GoTo errH
128     StrTime = Now
130     If intVer = 0 Then
132         gstrSql = "select id ,�걾ID,ͼ������,ͼ��λ�� from ����ͼ���� where �걾id = [1] order by ID"
        Else
134         gstrSql = "select id ,�걾ID,ͼ������,ͼ��λ�� from ���鱨��ͼ�� where �걾id = [1] order by ID"
        End If
    
136     Set rsTmp = OpenSQLRecord(gstrSql, "GetSampleImage", lngSampleID)
138     If rsTmp.EOF Then strErr = "��ͼ�����ݣ�"
        
        Call FunFtpSet(blnFtp, intVer, strFtpPara, strFtpUser, strFtpPass, strFtpIP, strFtpDir)
    
140     Do Until rsTmp.EOF
'            If strFileNames(UBound(strFileNames)) <> "" Then
            ReDim Preserve strFileNames(UBound(strFileNames) + 1)
            ReDim Preserve strFileAll(UBound(strFileAll) + 1)
'            End If
            If intVer = 0 Then
                strFileName = rsTmp("�걾ID") & strFileType
            Else
                strFileName = rsTmp("�걾ID") & "_" & Trim$("" & rsTmp("ͼ������")) & strFileType
            End If
142         If blnClearOldFile = True Then
                'ǿ����������ڱ��ص�ͼƬ
144             If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then gobjFSO.DeleteFile strSavePath & "\" & strFileName, True
                WriteLog "ǿ��������ػ����ͼƬ", strSavePath, strFileName
            End If
        
146         If Not gobjFSO.FileExists(strSavePath & "\" & strFileName) Then
                If blnFtp = False Then
148                 If LoadImageDataTwo(strSavePath, rsTmp("ID"), intSaveImgType, intVer, strFileName) Then
150                     If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then strReturn = strReturn & "," & strFileName
                    End If
                Else
                    strFtpPath = Trim("" & rsTmp!ͼ��λ��)
'                    If InStr(strFtpPath, ":") > 0 Then
                        strFtpPathName = Trim(Split(strFtpPath, ";")(1))
'                    End If
                    strFileNames(UBound(strFileNames)) = Split(strFtpPathName, "/")(UBound(Split(strFtpPathName, "/")))
                    strFileAll(UBound(strFileAll)) = Trim("" & rsTmp!ͼ������) & ";" & strFtpPath
                End If
152             DrawIndex = DrawIndex + 1
            Else
                strReturn = strReturn & "," & strFileName
            End If
154         rsTmp.MoveNext
        Loop
        
        If blnFtp Then
            strLocalFile = strPath
            blnDown = True
            For intLoop = 1 To UBound(strFileNames)
                If Dir(strPath & "\" & strFileNames(intLoop)) <> "" Then
                    strFileName = Replace(strFileNames(intLoop), ".bmp", ".cht")
                    strFileName = Replace(strFileName, ".jpg", ".cht")
                    strFileName = Replace(strFileName, ".png", ".cht")
                    strFileName = Replace(strFileName, ".BMP", ".cht")
                    strFileName = Replace(strFileName, ".JPG", ".cht")
                    strFileName = Replace(strFileName, ".PNG", ".cht")
                    strReturn = strReturn & "," & strFileName
                    blnDown = False
                End If
            Next
            If blnDown = True Then
                If DownFiles(strFtpUser, strFtpPass, strFtpIP, strFtpPathName, strLocalFile, strFileNames) = "" Then
                    For intLoop = 1 To UBound(strFileNames)
166                     strFtpPath = ""
168                     strFtpPath = strFileAll(intLoop)
                        'ͼ�����FTP�У���FTP��ȡ����
                        'ͼ��λ�õ����ݸ�ʽΪ��ͼ���ʽ;FTP�ļ�·��
                        strImageType = Trim(Split(strFtpPath, ";")(0))
262                     intLayOut = Val(Split(strFtpPath, ";")(1))
264                     strFtpPath = Trim(Split(strFtpPath, ";")(2))
266                     strImageData = ""
268                     If intLayOut >= 100 And intLayOut <= 227 Then
                            ' ͼƬ�ļ���ֱ�����ص�����
270                         strLocalFile = strPath & "\" & Split(strFtpPath, "/")(UBound(Split(strFtpPath, "/")))
278                         strImageData = intLayOut & ";" & strLocalFile
                        Else
                            ' ͼ�����ݣ���Ҫ�����ص��ı��ļ��ж�ȡ����
280                         strLocalFile = strPath & "\" & strFileNames(intLoop)
288                         Set objStream = gobjFSO.OpenTextFile(strLocalFile, ForReading)
290                             Do Until objStream.AtEndOfLine
292                                 strImageData = strImageData & objStream.ReadLine
                                Loop
294                             objStream.Close
296                             Set objStream = Nothing
298                             strImageData = Replace(Replace(Trim(strImageData), vbCr, ""), vbLf, "")
300                             strImageData = intLayOut & ";" & strImageData
302                         If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
                        End If
                    
304                     If Len(strImageData) <> 0 Then
                            strFileName = Replace(strFileNames(intLoop), ".bmp", ".cht")
                            strFileName = Replace(strFileName, ".jpg", ".cht")
                            strFileName = Replace(strFileName, ".png", ".cht")
                            strFileName = Replace(strFileName, ".BMP", ".cht")
                            strFileName = Replace(strFileName, ".JPG", ".cht")
                            strFileName = Replace(strFileName, ".PNG", ".cht")
                            
                            '��ͼ������ͼ���ļ�
306                         Call DrawImgAndSaveFile(strImageType, strImageData, strPath & "\" & strFileName, intSaveImgType)
                        End If
                        If gobjFSO.FileExists(strPath & "\" & strFileName) Then strReturn = strReturn & "," & strFileName
'                        intLoop = 0
'                        Do Until intLoop > 100
'                            intLoop = intLoop + 1
'                            If gobjFSO.FileExists(strLocalFile) Then
'                                WriteLog "LoadImageData", "��" & intLoop & "��ɾ��FTP���ص�ԭʼ�ļ�" & strLocalFile, ""
'                                Call gobjFSO.DeleteFile(strLocalFile)
'                            Else
'                                If strLocalFile <> "" Then WriteLog "LoadImageData", "FTP���ص�ԭʼ�ļ�" & strLocalFile & "��ɾ��!", ""
'                                Exit Do
'                            End If
'                        Loop
                    Next
                End If
            End If
        End If
156     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
158     GetSampleImages = strReturn
        WriteLog "����ͼ���ļ���", "", strReturn
        'Debug.Print "ID=" & lngKeyID & ",��ʱ:" & DateDiff("s", strTime, Now)
        Exit Function
errH:
160     GetSampleImages = ""
162    strErr = Err.Description
164    WriteLog "GetSampleImages", CStr(Erl()) & "��", Err.Description
End Function

Public Function GetSampleImagesOld(ByVal lngSampleID As Long, ByVal strPath As String, ByVal blnClearOldFile As Boolean, _
                                ByRef strErr As String, Optional intSaveImgType As Integer = 0, Optional intVer As Integer = 0) As String
        '��ȡһ���걾������ͼƬ��
        '
        'lngsampleid : �걾ID
        'strPath     : ͼƬ����·��,���������Զ�����
        'blnClearOldFile :�Ƿ���ջ����ڱ��ص�ͼ���ļ�,True��ÿ�ζ������ݿ���ļ����浽����;False-��һ�ε���ʱ�����ݿ��ͼ�β���ͼƬ��֮��ֱ��ʹ��
        'strErr      : ��������ֵΪ�մ�ʱ�����ص���ʾ��Ϣ
        'intImgType  : ���ص�ͼƬ�ļ���ʽ��0��cht(Ĭ��),1-jgp,2-png
        'intVer      : ���°�LIS�����ϰ�LIS�ڵ��ñ��������� 0-�ϰ�LIS��Ĭ�ϣ��ӡ�����ͼ��������ȡͼ�����ݣ���1-�°�LIS���ӡ����鱨��ͼ����ȡͼ�����ݣ�
    
        Dim rsTmp As New ADODB.Recordset
        Dim DrawIndex As Integer, strFileType As String
        Dim StrTime As Date
        Dim strSavePath As String
        Dim strReturn As String
        Dim strFileName As String
    
100     strReturn = ""
102     If Not gblnInit Then
104         strErr = "���ñ�����ǰ�����ȵ��ó�ʼ��������"
            Exit Function
        End If
    
106     If intSaveImgType = 1 Then
108         strFileType = ".jpg"
110     ElseIf intSaveImgType = 2 Then
112         strFileType = ".png"
        Else
114         strFileType = ".cht"
        End If
116     strSavePath = strPath
118     If Trim(strSavePath) = "" Then
120         strErr = "��ָ������ͼƬ��·����"
            Exit Function
122     ElseIf Not gobjFSO.FolderExists(strSavePath) Then
124         strErr = strSavePath & "�����ڣ�"
            Exit Function
        End If
126     If Right(strSavePath, 1) = "\" Then strSavePath = Left(strSavePath, Len(strSavePath) - 1)
        '---------------------------------------------------
        On Error GoTo errH
128     StrTime = Now
130     If intVer = 0 Then
132         gstrSql = "select id ,�걾ID,ͼ������ from ����ͼ���� where �걾id = [1] order by ID"
        Else
134         gstrSql = "select id ,�걾ID,ͼ������ from ���鱨��ͼ�� where �걾id = [1] order by ID"
        End If
    
136     Set rsTmp = OpenSQLRecord(gstrSql, "GetSampleImage", lngSampleID)
138     If rsTmp.EOF Then strErr = "��ͼ�����ݣ�"
        
    
140     Do Until rsTmp.EOF
            If intVer = 0 Then
                strFileName = rsTmp("�걾ID") & strFileType
            Else
                strFileName = rsTmp("�걾ID") & "_" & Trim$("" & rsTmp("ͼ������")) & strFileType
            End If
142         If blnClearOldFile = True Then
                'ǿ����������ڱ��ص�ͼƬ
144             If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then gobjFSO.DeleteFile strSavePath & "\" & strFileName, True
                WriteLog "ǿ��������ػ����ͼƬ", strSavePath, strFileName
            End If
        
146         If Not gobjFSO.FileExists(strSavePath & "\" & strFileName) Then
        
148             If LoadImageData(strSavePath, rsTmp("ID"), intSaveImgType, intVer, strFileName) Then
150                 If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then strReturn = strReturn & "," & strFileName
                End If
            
152             DrawIndex = DrawIndex + 1
            Else
                strReturn = strReturn & "," & strFileName
            End If
154         rsTmp.MoveNext
        Loop
156     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
158     GetSampleImagesOld = strReturn
        WriteLog "����ͼ���ļ���", "", strReturn
        'Debug.Print "ID=" & lngKeyID & ",��ʱ:" & DateDiff("s", strTime, Now)
        Exit Function
errH:
160     GetSampleImagesOld = ""
162    strErr = Err.Description
164    WriteLog "GetSampleImages", CStr(Erl()) & "��", Err.Description
End Function

Public Function GetImage(ByVal lngImgID As Long, ByVal strPath As String, ByVal blnClearOldFile As Boolean, _
                                ByRef strErr As String, Optional intSaveImgType As Integer = 0, Optional intVer As Integer = 0) As String
        '��ȡָ����һ��ͼƬ��
        '
        'lngImgid    : ͼƬID
        'strPath     : ͼƬ����·��,���������Զ�����
        'blnClearOldFile :�Ƿ���ջ����ڱ��ص�ͼ���ļ�,True��ÿ�ζ������ݿ���ļ����浽����;False-��һ�ε���ʱ�����ݿ��ͼ�β���ͼƬ��֮��ֱ��ʹ��
        'strErr      : ��������ֵΪ�մ�ʱ�����ص���ʾ��Ϣ
        'intImgType  : ���ص�ͼƬ�ļ���ʽ��0��cht(Ĭ��),1-jgp,2-png
        'intVer      : ���°�LIS�����ϰ�LIS�ڵ��ñ��������� 0-�ϰ�LIS��Ĭ�ϣ��ӡ�����ͼ��������ȡͼ�����ݣ���1-�°�LIS���ӡ����鱨��ͼ����ȡͼ�����ݣ�
        Dim strSavePath As String, strFileType As String
        Dim strReturn As String
        Dim rsTmp As New ADODB.Recordset
        Dim strFileName As String
100     If Not gblnInit Then
102         strErr = "���ñ�����ǰ�����ȵ��ó�ʼ��������"
            Exit Function
        End If
104     If intSaveImgType = 1 Then
106         strFileType = ".jpg"
108     ElseIf intSaveImgType = 2 Then
110         strFileType = ".png"
        Else
112         strFileType = ".cht"
        End If
114     strSavePath = strPath
116     If Trim(strSavePath) = "" Then
118         strErr = "��ָ������ͼƬ��·����"
            Exit Function
120     ElseIf Not gobjFSO.FolderExists(strSavePath) Then
122         strErr = strSavePath & "�����ڣ�"
            Exit Function
        End If
124     If Right(strSavePath, 1) = "\" Then strSavePath = Left(strSavePath, Len(strSavePath) - 1)
        On Error GoTo errH
    
126     If intVer = 0 Then
128         gstrSql = "select id ,�걾ID,ͼ������ from ����ͼ���� where id = [1]"
        Else
130         gstrSql = "select id ,�걾ID,ͼ������ from ���鱨��ͼ�� where id = [1]"
        End If
    
132     Set rsTmp = OpenSQLRecord(gstrSql, "GetImage", lngImgID)
134     If rsTmp.EOF Then strErr = "��ͼ�����ݣ�"
        If intVer = 0 Then
            strFileName = rsTmp("ID") & strFileType
        Else
            strFileName = rsTmp("ID") & "_" & Trim$("" & rsTmp("ͼ������")) & strFileType
        End If
136     If blnClearOldFile = True Then
            'ǿ����������ڱ��ص�ͼƬ
138         If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then gobjFSO.DeleteFile strSavePath & "\" & strFileName, True
            WriteLog "ǿ��������ػ����ͼƬ", strSavePath, strFileName
        End If
    
140     If Not gobjFSO.FileExists(strSavePath & "\" & strFileName) Then
142         If LoadImageData(strSavePath, rsTmp("ID"), intSaveImgType, intVer, strFileName) Then
144             If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then strReturn = strReturn & "," & strFileName
            Else
                strErr = "����ͼ���ļ�ʧ�ܣ�"
            End If
        Else
            strReturn = strReturn & "," & strFileName
        End If
146     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
148     GetImage = strReturn
        WriteLog "����ͼ���ļ���", "", strReturn
        Exit Function
errH:
150     GetImage = ""
152     strErr = Err.Description
154     WriteLog "GetImage", CStr(Erl()) & "��", Err.Description
End Function


Public Function GetSampleImgInit(ByVal lngSys As Long, ByVal cnnOra As ADODB.Connection, ByRef strErr As String) As Boolean
    '��ʼ�����Ӽ���������
    '
    Dim arrData
    Dim i As Integer
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    
    
    GetSampleImgInit = gblnInit
    glngSys = lngSys
    If gblnInit = True Then Exit Function
    Set gcnOracle = cnnOra
    strSQL = "Select SYS_CONTEXT('USERENV','TERMINAL') as MName From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "��ʼ��")
    gstrComputerName = rsTmp("MName")
'    With gcnOracle
'        Set gComLib = CreateObject("zl9Comlib.clsComlib")
'        Call gComLib.initCommon(gcnOracle)
'    End With
    
    gblnInit = True
    GetSampleImgInit = True
    Exit Function
errH:
    strErr = Err.Description
    WriteLog "GetImgInit", CStr(Erl()) & "��", Err.Description
End Function

Public Function GetSampleImgExit(ByRef strErr As String) As Boolean
    '�ر�����
    On Error GoTo errH
    strErr = ""
    If gblnInit Then
        Set gComLib = Nothing
        Set gcnOracle = Nothing
        gblnInit = False
    End If
    GetSampleImgExit = True
    Exit Function
errH:
    strErr = Err.Description
    WriteLog "GetImgExit", CStr(Erl()) & "��", Err.Description
End Function


Public Function LocalImageSaveToCht(ByVal strTitle As String, ByVal strImageData As String, ByVal strSaveFileName As String) As Boolean
    '������ͼ���ļ�����Cht����ָ��Ŀ¼
    Dim strPath As String
    Dim strChtName As String
    strPath = App.Path & "\LisImage"
    If Not gobjFSO.FolderExists(strPath) Then gobjFSO.CreateFolder (strPath)
    strChtName = strPath & "\" & strSaveFileName
    If gobjFSO.FileExists(strChtName) Then gobjFSO.DeleteFile strChtName, True
    
    LocalImageSaveToCht = DrawImgAndSaveFile(strTitle, strImageData, strChtName, 0)
End Function


