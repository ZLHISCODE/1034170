Attribute VB_Name = "mdlZip"
Option Explicit

Public Type ZIPnames
    s(0 To 99) As String
End Type

'ZPOPT is used to set options in the zip32.dll
Private Type ZPOPT
    fSuffix As Long
    fEncrypt As Long
    fSystem As Long
    fVolume As Long
    fExtra As Long
    fNoDirEntries As Long
    fExcludeDate As Long
    fIncludeDate As Long
    fVerbose As Long
    fQuiet As Long
    fCRLF_LF As Long
    fLF_CRLF As Long
    fJunkDir As Long
    fRecurse As Long
    fGrow As Long
    fForce As Long
    fMove As Long
    fDeleteEntries As Long
    fUpdate As Long
    fFreshen As Long
    fJunkSFX As Long
    fLatestTime As Long
    fComment As Long
    fOffsets As Long
    fPrivilege As Long
    fEncryption As Long
    fRepair As Long
    flevel As Byte
    date As String ' 8 bytes long
    szRootDir As String ' up to 256 bytes long
End Type

Private Type ZIPUSERFUNCTIONS
    DllPrnt As Long
    DLLPASSWORD As Long
    DLLCOMMENT As Long
    DLLSERVICE As Long
End Type

'Structure ZCL - not used by VB
'Private Type ZCL
'    argc As Long            'number of files
'    filename As String      'Name of the Zip file
'    fileArray As ZIPnames   'The array of filenames
'End Type

' Call back "string" (sic)
' Callback large "string" (sic)
Private Type CBChar
    ch(4096) As Byte
End Type

' Callback small "string" (sic)
Private Type CBCh
    ch(256) As Byte
End Type


' DCL structure
Private Type DCLIST
    ExtractOnlyNewer As Long
    SpaceToUnderscore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nUflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    fPrivilege As Long
    Zip As String
    ExtractDir As String
End Type

' Userfunctions structure
Private Type USERFUNCTION
    DllPrnt As Long
    DLLSND As Long
    DLLREPLACE As Long
    DLLPASSWORD As Long
    DLLMESSAGE As Long
    DLLSERVICE As Long
    TotalSizeComp As Long
    TotalSize As Long
    CompFactor As Long
    NumMembers As Long
    cchComment As Integer
End Type

' Unzip32.dll version structure
Private Type UZPVER
    structlen As Long
    flag As Long
    beta As String * 10
    date As String * 20
    zlib As String * 10
    unzip(1 To 4) As Byte
    zipinfo(1 To 4) As Byte
    os2dll As Long
    windll(1 To 4) As Byte
End Type

Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

'This assumes zip32.dll is in your \windows\system directory!
Private Declare Function ZpInit Lib "zip32.dll" (ByRef Zipfun As ZIPUSERFUNCTIONS) As Long ' Set Zip Callbacks
Private Declare Function ZpSetOptions Lib "zip32.dll" (ByRef Opts As ZPOPT) As Long ' Set Zip options
Private Declare Function ZpGetOptions Lib "zip32.dll" () As ZPOPT ' used to check encryption flag only
Private Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long ' Real zipping action
Private Declare Function windll_unzip Lib "unzip32.dll" _
    (ByVal ifnc As Long, ByRef ifnv As ZIPnames, _
     ByVal xfnc As Long, ByRef xfnv As ZIPnames, _
     dcll As DCLIST, Userf As USERFUNCTION) As Long

Private Declare Sub UzpVersion2 Lib "unzip32.dll" (uzpv As UZPVER)
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Function FnPtr(ByVal lp As Long) As Long
'���ܣ�ȡ�ú�����ָ��ֵ
    FnPtr = lp
End Function

' Callback for unzip32.dll
Sub ReceiveDllMessage(ByVal ucsize As Long, _
    ByVal csiz As Long, _
    ByVal cfactor As Integer, _
    ByVal mo As Integer, _
    ByVal dy As Integer, _
    ByVal yr As Integer, _
    ByVal hh As Integer, _
    ByVal mm As Integer, _
    ByVal c As Byte, ByRef fname As CBCh, _
    ByRef meth As CBCh, ByVal crc As Long, _
    ByVal fCrypt As Byte)

'���ս�ѹ�����з��ص���Ϣ
    Dim strTemp As String, lngCount As Long
    Dim strInfo As String * 80

    ' always put this in callback routines!
    On Error Resume Next
    strInfo = Space(80)
'    If vbzipnum = 0 Then
'        Mid$(strInfo, 1, 50) = "Filename:"
'        Mid$(strInfo, 53, 4) = "Size"
'        Mid$(strInfo, 62, 4) = "Date"
'        Mid$(strInfo, 71, 4) = "Time"
'        vbzipmes = strInfo + vbCrLf
'        strInfo = Space(80)
'    End If
    strTemp = ""
    For lngCount = 0 To 255
        If fname.ch(lngCount) = 0 Then
            lngCount = 99999
        Else
            strTemp = strTemp & Chr$(fname.ch(lngCount))
        End If
    Next lngCount
    Mid$(strInfo, 1, 50) = Mid$(strTemp, 1, 50)
    Mid$(strInfo, 51, 7) = Right$("        " + Str$(ucsize), 7)
    Mid$(strInfo, 60, 3) = Right$(Str$(dy), 2) + "/"
    Mid$(strInfo, 63, 3) = Right$("0" + Trim$(Str$(mo)), 2) + "/"
    Mid$(strInfo, 66, 2) = Right$("0" + Trim$(Str$(yr)), 2)
    Mid$(strInfo, 70, 3) = Right$(Str$(hh), 2) + ":"
    Mid$(strInfo, 73, 2) = Right$("0" + Trim$(Str$(mm)), 2)
    ' Mid$(strInfo, 75, 2) = Right$(" " + Str$(cfactor), 2)
    ' Mid$(strInfo, 78, 8) = Right$("        " + Str$(csiz), 8)
    ' strTemp = ""
    ' For lngCount = 0 To 255
    '     If meth.ch(lngCount) = 0 Then lngCount = 99999 Else strTemp = strTemp + Chr(meth.ch(lngCount))
    ' Next lngCount
    '��ѹ���ļ�����
'    vbzipmes = vbzipmes + strInfo + vbCrLf
'    vbzipnum = vbzipnum + 1
End Sub

' Callback for unzip32.dll
Function DllPrnt(ByRef fname As CBChar, ByVal lngLength As Long) As Long
    Dim strTemp As String, lngCount As Long

    ' always put this in callback routines!
    On Error Resume Next
    strTemp = ""
    For lngCount = 0 To lngLength
        If fname.ch(lngCount) = 0 Then
            lngCount = 99999
        Else
            strTemp = strTemp + Chr(fname.ch(lngCount))
        End If
    Next lngCount
    DllPrnt = 0
End Function

' Callback for unzip32.dll
Function DllPass(ByRef s1 As Byte, x As Long, _
    ByRef s2 As Byte, _
    ByRef s3 As Byte) As Long

    ' always put this in callback routines!
    On Error Resume Next
    ' not supported - always return 1
    DllPass = 1
End Function

Function DllRep(ByRef fname As CBChar) As Long
'���ܣ��ļ�����ʱ�����֡��Ƿ��滻�ļ�������Ϣ
'      ��unzip32.dll����

    Dim strTemp As String, lngCount As Long
    
    On Error Resume Next
    
    DllRep = 100 ' 100=do not overwrite - keep asking user
    '����ļ���
    strTemp = ""
    For lngCount = 0 To 255
        If fname.ch(lngCount) = 0 Then
            lngCount = 99999
        Else
            strTemp = strTemp + Chr(fname.ch(lngCount))
        End If
    Next lngCount
    
    lngCount = MsgBox("�ļ���" + strTemp + "���Ѿ����ڣ��Ƿ��滻��", vbQuestion Or vbYesNoCancel, gstrSysName)
    
    If lngCount = vbNo Then Exit Function
    If lngCount = vbCancel Then
        DllRep = 104 ' 104=overwrite none
        Exit Function
    End If
    DllRep = 102 ' 102=overwrite 103=overwrite all
End Function

Function szTrim(szString As String) As String
'���ܣ�ȥ��\0�Ժ���ַ���ASCIIZ to String
    
    Dim pos As Integer, ln As Integer

    pos = InStr(szString, Chr$(0))
    ln = Len(szString)
    Select Case pos
        Case Is > 1
            szTrim = Trim(Left(szString, pos - 1))
        Case 1
            szTrim = ""
        Case Else
            szTrim = Trim(szString)
    End Select
End Function

' Callback for zip32.dll
Function DllComm(ByRef s1 As CBChar) As CBChar
    
    ' always put this in callback routines!
    On Error Resume Next
    ' not supported always return \0
    s1.ch(0) = vbNullString
    DllComm = s1
End Function

' Main subroutine
Function VBUnzip(fname As String, extdir As String, _
    prom As Integer, over As Integer, _
    mess As Integer, dirs As Integer, numfiles As Long, numxfiles As Long, _
    vbzipnam As ZIPnames, vbxnames As ZIPnames) As Boolean
'���ܣ���ѹ����
'����˵��
'    zipfile    ҪUnzip���ļ�
'    unzipdir   ���ý�ѹ���ļ���Ŀ¼
'    prom       1 = ���ڸ��ǽ�����ʾ
'    over       1 = ���Ǹ���
'    mess       1 = ֻ�г��ļ�����  0 = ��ѹ
'    dirs       1 = ����ZIP�ļ��е�·��
'    vbzipnam  ��ѡ�Ľ�ѹ���ļ�
'    vbxnames  Ҫ���ų��Ľ�ѹ�ļ�
    
    Dim lngCount As Long ' , s1 As String * 20, s2 As String * 256
    
    Dim MYUSER As USERFUNCTION
    Dim MYDCL As DCLIST
    Dim MYVER As UZPVER

    ' Set options
    With MYDCL
        .ExtractOnlyNewer = 0      ' 1=extract only newer
        .SpaceToUnderscore = 0     ' 1=convert space to underscore
        .PromptToOverwrite = prom  ' 1=prompt to overwrite required
        .fQuiet = 0                ' 2=no messages 1=less 0=all
        .ncflag = 0                ' 1=write to stdout
        .ntflag = 0                ' 1=test zip
        .nvflag = mess             ' 0=extract 1=list contents
        .nUflag = 0                ' 1=extract only newer
        .nzflag = 0                ' 1=display zip file comment
        .ndflag = dirs             ' 1=honour directories
        .noflag = over              ' 1=overwrite files
        .naflag = 0                ' 1=convert CR to CRLF
        .nZIflag = 0               ' 1=Zip Info Verbose
        .C_flag = 0                ' 1=Case insensitivity, 0=Case Sensitivity
        .fPrivilege = 0            ' 1=ACL 2=priv
        .Zip = fname               ' ZIP name
        .ExtractDir = extdir       ' Extraction directory, NULL if extracting
    End With                              ' to current directory
    
    '�����ڲ������ĵ�ַ
    With MYUSER
        .DllPrnt = FnPtr(AddressOf DllPrnt)
        .DLLSND = 0& ' not supported
        .DLLREPLACE = FnPtr(AddressOf DllRep)
        .DLLPASSWORD = FnPtr(AddressOf DllPass)
        .DLLMESSAGE = FnPtr(AddressOf ReceiveDllMessage)
        .DLLSERVICE = 0& ' not coded yet :)
    End With
    ' Set Version space
    ' Do not change
    With MYVER
        .structlen = Len(MYVER)
        .beta = Space(9) & vbNullChar
        .date = Space(19) & vbNullChar
        .zlib = Space(9) & vbNullChar
    End With
    
    ' Get version
    Call UzpVersion2(MYVER)
    
    ' Go for it!
    lngCount = windll_unzip(numfiles, vbzipnam, _
        numxfiles, vbxnames, MYDCL, MYUSER)
        
    If lngCount = 0 Then
        VBUnzip = True
    Else
        VBUnzip = False
        MsgBox "�����ļ� " & fname & " ��ѹʧ�ܡ�", vbInformation, gstrSysName
    End If
End Function

'Main Subroutine
Function VBZip(argc As Integer, zipname As String, _
        mynames As ZIPnames, junk As Integer, _
        recurse As Integer, updat As Integer, _
        freshen As Integer, basename As String) As Boolean
        
'���ܣ�ѹ���ļ�
'������argc         �ļ�����
'      zipname      ZIP�ļ���
'      mynames      Ҫѹ�����ļ��б�
'      junk         1 �׿�Ŀ¼��
'      recurse      ZIP�ļ���
'      updat        ZIP�ļ���
    Dim hmem As Long, lngCount As Integer
    Dim retcode As Long
    Dim MYOPT As ZPOPT
    Dim MYUSER As ZIPUSERFUNCTIONS
    
    On Error Resume Next ' nothing will go wrong :-)
    
    '�����ڲ������ĵ�ַ
    With MYUSER
        .DllPrnt = FnPtr(AddressOf DllPrnt)
        .DLLPASSWORD = FnPtr(AddressOf DllPass)
        .DLLCOMMENT = FnPtr(AddressOf DllComm)
        .DLLSERVICE = 0& ' not coded yet :-)
    End With
    retcode = ZpInit(MYUSER)
    
    '����ѹ��ѡ��
    With MYOPT
        .fSuffix = 0        ' include suffixes (not yet implemented)
        .fEncrypt = 0       ' 1 if encryption wanted
        .fSystem = 0        ' 1 to include system/hidden files
        .fVolume = 0        ' 1 if storing volume label
        .fExtra = 0         ' 1 if including extra attributes
        .fNoDirEntries = 0  ' 1 if ignoring directory entries
        .fExcludeDate = 0   ' 1 if excluding files earlier than a specified date
        .fIncludeDate = 0   ' 1 if including files earlier than a specified date
        .fVerbose = 0       ' 1 if full messages wanted
        .fQuiet = 0         ' 1 if minimum messages wanted
        .fCRLF_LF = 0       ' 1 if translate CR/LF to LF
        .fLF_CRLF = 0       ' 1 if translate LF to CR/LF
        .fJunkDir = junk    ' 1 if junking directory names
        .fRecurse = recurse ' 1 if recursing into subdirectories
        .fGrow = 0          ' 1 if allow appending to zip file
        .fForce = 0         ' 1 if making entries using DOS names
        .fMove = 0          ' 1 if deleting files added or updated
        .fDeleteEntries = 0 ' 1 if files passed have to be deleted
        .fUpdate = updat    ' 1 if updating zip file--overwrite only if newer
        .fFreshen = freshen ' 1 if freshening zip file--overwrite only
        .fJunkSFX = 0       ' 1 if junking sfx prefix
        .fLatestTime = 0    ' 1 if setting zip file time to time of latest file in archive
        .fComment = 0       ' 1 if putting comment in zip file
        .fOffsets = 0       ' 1 if updating archive offsets for sfx Files
        .fPrivilege = 0     ' 1 if not saving privelages
        .fEncryption = 0    'Read only property!
        .fRepair = 0        ' 1=> fix archive, 2=> try harder to fix
        .flevel = 0         ' compression level - should be 0!!!
        .date = vbNullString ' "12/31/79"? US Date?
        .szRootDir = basename
    End With
    ' Set options
    retcode = ZpSetOptions(MYOPT)
    
    ' ZCL not needed in VB
    ' MYZCL.argc = 2
    ' MYZCL.filename = "c:\wiz\new.zip"
    ' MYZCL.fileArray = MYNAMES
    
    ' Go for it!
    retcode = ZpArchive(argc, zipname, mynames)
    
    If retcode = 0 Then
        VBZip = True
    Else
        VBZip = False
        MsgBox "���ϴ��ļ� " & zipname & " ѹ��ʧ�ܡ�", vbInformation, gstrSysName
    End If
End Function


Public Function OpenDir(frmOwner As Form, Optional strTitle As String) As String
'���ܣ����Ŀ¼��
'������frmOwner    �����ߴ���
'      strTitle    ѡ�񴰿ڱ���
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = strTitle
   With tBrowseInfo
      .hwndOwner = frmOwner.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      OpenDir = sBuffer
   End If
End Function



