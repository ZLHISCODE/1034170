VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LX7680"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�����LX7680
'����:��������
'ҽԺ��˫������ҽԺ
'��ʽ��SQLserver���ݿ���ȡ��ʽ
'��ȡ�ļ���Examdir_Data.mdf���ݿ��е�ITEMLIST��
'���ߣ��µ���

Private FileNames As String

Private Const MAX_PATH = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

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
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

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

Dim fso As New FileSystemObject
Dim fld As Folder

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Function StripNulls(OriginalStr As String) As String
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, _
      InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function

Public Function FindFilesAPI(ByVal Path As String, ByVal SearchStr As String)
   Dim Filename As String   ' Walking filename variable...
   Dim DirName As String    ' ���ļ�����
   Dim dirNames() As String ' Buffer for directory name entries
   Dim nDir As Integer   ' Number of directories in this path
   Dim i As Integer
   Dim hSearch As Long   ' ���Ҿ��
   Dim WFD As WIN32_FIND_DATA
   Dim Cont As Integer
   Dim FT As FILETIME
   Dim ST As SYSTEMTIME
   Dim DateCStr As String, DateMStr As String
      
   If Right(Path, 1) <> "\" Then Path = Path & "\"
   ' Search for subdirectories.
   nDir = 0
   ReDim dirNames(nDir)
   Cont = True
   hSearch = FindFirstFile(Path & "*", WFD)
   If hSearch <> INVALID_HANDLE_VALUE Then
      Do While Cont
         DirName = StripNulls(WFD.cFileName)
         If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
               dirNames(nDir) = DirName
               nDir = nDir + 1
               ReDim Preserve dirNames(nDir)
            End If
         End If
         Cont = FindNextFile(hSearch, WFD) ' Get next subdirectory.
      Loop
      Cont = FindClose(hSearch)
   End If
 
   ' Walk through this directory and sum file sizes.
   hSearch = FindFirstFile(Path & SearchStr, WFD)
   Cont = True
   If hSearch <> INVALID_HANDLE_VALUE Then
      While Cont
         Filename = StripNulls(WFD.cFileName)
            If (Filename <> ".") And (Filename <> "..") And _
              ((GetFileAttributes(Path & Filename) And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
            FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * _
             MAXDWORD) + WFD.nFileSizeLow
             
           '�ļ�������Ϣ
           FileTimeToLocalFileTime WFD.ftCreationTime, FT
           FileTimeToSystemTime FT, ST
           DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
              " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
           ' and Last Modified Date
           FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
           FileTimeToSystemTime FT, ST
           DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
              " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
             
           If FileNames <> "" Then FileNames = FileNames + "|"
           FileNames = FileNames & Filename
           'List1.AddItem Path & FileName & vbTab & _
           '   Format(DateCStr, "mm/dd/yyyy hh:nn:ss") _
           '   & vbTab & Format(DateMStr, "mm/dd/yyyy hh:nn:ss")
          End If
         Cont = FindNextFile(hSearch, WFD) '������һ���ļ�
      Wend
      Cont = FindClose(hSearch)
   End If
 
    '�������ļ���
    If nDir > 0 Then
      ' Recursively walk into them...
      For i = 0 To nDir - 1
        FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) _
         & "\", SearchStr)
      Next i
   End If
End Function

Private Function FindFile(ByVal sFol As String, sFile As String) As Currency
   Dim tFld As Folder, tFil As File, Filename As String
    
   On Error GoTo Catch
   Set fld = fso.GetFolder(sFol)
   Filename = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
                  vbHidden Or vbSystem Or vbReadOnly)
   While Len(Filename) <> 0
      FindFile = FindFile + FileLen(fso.BuildPath(fld.Path, _
      Filename))
      
      If FileNames <> "" Then FileNames = FileNames + "|"
      FileNames = FileNames & fso.BuildPath("", Filename)
      
      Filename = Dir() ' Get next file
      DoEvents
   Wend
   If fld.SubFolders.count > 0 Then
      For Each tFld In fld.SubFolders
         DoEvents
         FindFile = FindFile + FindFile(tFld.Path, sFile)
      Next
   End If
   Exit Function
Catch: Filename = ""
       Resume Next
End Function


Private Function FindFile2(ByVal sPath As String, sTyp As String) As String
   Dim aTmp As String, bTmp As String
   aTmp = Dir(sPath & sTyp)   ' ��Ѱ��һ�
   bTmp = ""
   Do While aTmp <> ""   '��ʼѭ����
     ' ������ǰ��Ŀ¼���ϲ�Ŀ¼��
     If aTmp <> "." And aTmp <> ".." Then
       If bTmp <> "" Then bTmp = bTmp & "|"
         bTmp = bTmp & aTmp   ' �������һ��Ŀ¼������������ʾ������
     End If
    aTmp = Dir   ' ������һ��Ŀ¼��
   Loop
   FindFile2 = bTmp
End Function

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'��׼���������ļ���ȡ������
'   strFile������·�����ļ���
'   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
'   ���صļ���������(������������밴���±�׼��֯���)
'   Ԫ��֮����|�ָ�
'   ��0��Ԫ�أ�����ʱ��
'   ��1��Ԫ�أ��������
'   ��2��Ԫ�أ�������
'   ��3��Ԫ�أ��걾
'   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
'   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
' Dim DBConn As New ADODB.Connection
'��SQL���ݿ�
  ' DBConn.Open "Provider   =   sqloledb;Data   Source= '�������� ';
   'Initial   Catalog= '���ݿ����� ';User   Id= '�û��� ';Password= '���� ' "
   '�粻�뽫�����ַ���д�ڳ�����Ļ���������ͨ����INI�ļ�������������������:
    '���鼼ʦ����վָ���ļ�ʱֻ��ָ�����ini�ļ�����
    '[LX2000]
    'Database Name = ExamDir
    'SERVER Name = (local)
    'USER Name = sa
    'PASSWORD=
    Dim aResults() As String
    Dim ConnSQL As New ADODB.Connection
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim lngResultCount As Long, strBh As String
    Dim str���� As String, strImg As String, strTmp As String
    Dim strTDM As String, strPicPath As String, strPicFix As String, strTmp2 As String, strTmp3 As String
    Dim intloop As Integer
    Dim aFiles() As String, strLog As String, strImgBmp As String, YearMonth As String, YearMonthDay As String, YearMonthSample As String
    
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream
    
'    strLog = "�Ѿ����ò���"
'    Call writelislog(strLog)
    
    strIniFile = App.Path & "\LX7860.ini"
    
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "[Options]"
       objStream.WriteLine "PicPath=C:\Program Files\LX3000\Images"
       objStream.WriteLine "PicFix=U100N"
       objStream.Close
       Set objStream = Nothing
       strPicPath = ""
    Else
       strPicPath = ReadIni("Options", "PicPath", strIniFile)
       strPicFix = ReadIni("Options", "PicFix", strIniFile)
       If (Len(strPicPath) > 0) Then
         If Mid(strPicPath, Len(strPicPath)) <> "\" Then strPicPath = strPicPath + "\"
       End If
    End If
    
    With ConnSQL
        .ConnectionString = "Provider=SQLOLEDB;Data Source=(local);user id=sa;password=;initial catalog=ExamDir"
        .Open
        If .State = adStateClosed Then
'          strLog = "����SQL SERVER���ݿ�ʧ�ܡ�"
'          Call writelislog(strLog)
          Exit Function
        End If
        
'        strLog = "����SQL SERVER���ݿ�ɹ���"
'        Call writelislog(strLog)
        
        strSQL = "SELECT EXAMINEDATE, SAMPLETYPE,SAMPLEID,ITEMCODE, ITEMVALUE,ITEMCHAR,QUERYID " & vbNewLine & _
                "From ITEMLIST " & vbNewLine & _
                "Where EXAMINEDATE >= '" & Format(dtStart, "YYYY-MM-DD") & "'" & _
                " and EXAMINEDATE <= '" & Format(dtEnd, "YYYY-MM-DD") & "'" & _
                 IIf(lngSampleNO = -1, "", " And SAMPLEID=" & lngSampleNO) & _
                " order by SAMPLEID,EXAMINEDATE,ITEMCODE"
        Set rsSQL = ConnSQL.Execute(strSQL)
        
        If rsSQL.EOF Then
            .Close
            ReDim Preserve aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If
        
        rsSQL.MoveFirst
        lngResultCount = -1
        strBh = "***"
        Do While Not rsSQL.EOF
            If strBh <> rsSQL("SAMPLEID") Then '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsSQL("EXAMINEDATE"), "yyyy-mm-dd hh:mm:ss")  '����ʱ��
                
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsSQL("SAMPLEID"))   '�������   Val("" & rsSQL("SAMPLEID"))
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      " "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      "��Һ"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                   ' Val(rsSQL("�ʿ�")) '�ʿ�
                    
                strBh = rsSQL("SAMPLEID")
                   
            End If
              '������ͨ�����滻��Ӣ��
                strTDM = rsSQL("ITEMCODE")
                strTDM = Replace(strTDM, "��ϸ������", "rbc")
                strTDM = Replace(strTDM, "��ϸ������", "TWBC")
                strTDM = Replace(strTDM, "��Ƥϸ������", "spxb")
                strTDM = Replace(strTDM, "�����Թ���", "blxgx")
                strTDM = Replace(strTDM, "�Ǿ�������ᾧ", "nsjj")
                strTDM = Replace(strTDM, "͸������", "tmgx")
                strTDM = Replace(strTDM, "ճҺ˿", "nys")
                
                strTDM = Replace(strTDM, "�γ�", "dc")
                strTDM = Replace(strTDM, "���", "zj")
                strTDM = Replace(strTDM, "ϸ��", "xj")
                strTDM = Replace(strTDM, "�ᾧ", "jj")
                strTDM = Replace(strTDM, "Unknown", "kong")
                strTDM = Replace(strTDM, "����ƽᾧ", "csgjj")
                strTDM = Replace(strTDM, "����", "jz")
                strTDM = Replace(strTDM, "��������", "klgx")
            If strTDM <> "" Or IsNull(strTDM) = False Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Trim(strTDM) & "|" & Trim(Replace(Replace(Replace(IIf(rsSQL("ITEMCHAR") = "", rsSQL("ITEMVALUE"), rsSQL("ITEMCHAR")), "+-", "������(+-)"), "1+", "����(1+)"), "2+", "����(2+)"))
            End If
            
            If Trim(rsSQL("ITEMCODE")) = "���" Then
              '��ȡ����ѡ��ͼƬ
              strImgBmp = ""
              strTmp = ""
              YearMonth = Format(rsSQL("EXAMINEDATE"), "yyyymm")
              YearMonthDay = Format(rsSQL("EXAMINEDATE"), "yyyymmdd")
              YearMonthSample = YearMonthDay & Trim(rsSQL("SAMPLEID"))
              strImgBmp = strPicPath & YearMonth & "\" & YearMonthDay & "\" & YearMonthSample & "\"
'              strLog = "����ͼƬ·��: " & strImgBmp
'              Call writelislog(strLog)

              strPicFix = Trim(rsSQL("QUERYID")) & "*p.jpg"

'              strLog = "����ͼƬ����: " & strPicFix
'              Call writelislog(strLog)

              'Call FindFilesAPI(strPicPath, strPicFix)
              'Call FindFile(strPicPath, strPicFix)
              FileNames = FindFile2(strImgBmp, strPicFix)
              aFiles = Split(FileNames, "|")
              For intloop = LBound(aFiles) To UBound(aFiles)
                strImg = aFiles(intloop)
                If (Trim(strImg) <> "") And (Dir(strImgBmp + strImg) <> "") Then
                   strTmp = strTmp & "^" & CStr(intloop + 1) & ";113;0;" & strImgBmp & strImg
'                  strTmp = strTmp & "^���������" & CStr(intloop + 1) & ";113;0;" & strImgBmp & strImg
'                  strLog = "�ҵ�ѡ��ͼƬ: " & strImgBmp & strImg
'                  Call writelislog(strLog)
                End If
              Next
              If strTmp <> "" Then aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & Mid(strTmp, 2)

'              strLog = "��ǰ����: " & aResults(lngResultCount)
'              Call writelislog(strLog)
            End If
            rsSQL.MoveNext
        Loop
    End With
    
    If rsSQL.EOF Then
        ConnSQL.Close
    End If
    ResultFromFile = aResults
End Function

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_LX2000_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLog)
    objStream.Close
    Set objStream = Nothing
End Function

