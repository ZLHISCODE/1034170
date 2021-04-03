VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mblnIsLoadReportImage As Boolean '�Ǹ��ݱ���ͼ���ֶμ��صı���ͼ
Private mslQueryLevel As TQueryLevel      'ͼ����ʾ����
Private mtyFileLoadType As FileLoadType

'�ļ����ط�ʽ
Property Get ImgLoadType() As FileLoadType
    ImgLoadType = mtyFileLoadType
End Property

Property Let ImgLoadType(value As FileLoadType)
    mtyFileLoadType = value
End Property


'ͼ����ʾ����
Property Get QueryLevel() As TQueryLevel
    QueryLevel = mslQueryLevel
End Property

Property Let QueryLevel(value As TQueryLevel)
    mslQueryLevel = value
End Property

Public Function DownloadImages(ByRef arrImages() As String, strQueryValue As String, ByVal lngStartIndex As Long, ByVal lngEndIndex As Long, _
                blnOnlyLoadReportImage As Boolean, blnMoved As Boolean, _
                Optional blnQueryTmpRecord As Boolean) As Boolean
'lngAdviceId����ѯ����ֵ
'blnOnlyLoadReportImage��ΪTrueʱ���� ����ͼ�� �ֶ��еı���ͼ,��֮�������б���ͼ
'lngStartIndex����һ��ͼ������
'lngEndIndex�����һ��ͼ������
'blnMoved��ͼ���Ƿ�ת��
'slQueryLevel��ͼ�����ؼ���
'blnQueryTmpRecord���Ƿ���ʱͼ��
    Dim rsData As Recordset

    If strQueryValue = "0" Then Exit Function
    If QueryLevel = slAdvice And Not IsNumeric(Val(strQueryValue)) Then Exit Function

    If blnOnlyLoadReportImage Then
        '���� Ӱ�����¼.����ͼ�� �ֶ��е�ֵ���أ����Ϊ�գ� ���������б���ͼ��
        Set rsData = GetImageRptData(strQueryValue, lngStartIndex, lngEndIndex, blnMoved)
        
        mblnIsLoadReportImage = rsData.RecordCount > 0
        
        If rsData.RecordCount <= 0 Then
            Set rsData = GetImageViewData(mslQueryLevel, strQueryValue, lngStartIndex, lngEndIndex, blnQueryTmpRecord, blnMoved, blnOnlyLoadReportImage)
        End If
    Else
        Set rsData = GetImageViewData(mslQueryLevel, strQueryValue, lngStartIndex, lngEndIndex, blnQueryTmpRecord, blnMoved, blnOnlyLoadReportImage)
    End If
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If mtyFileLoadType = Normal Then
        DownloadImages = LoadViewImageToFaceWithNormal(rsData, arrImages)     'ʹ��ԭʼģʽ����
    Else
        DownloadImages = LoadViewImageToFaceWithService(rsData, arrImages)   'ʹ��ZLPacsServerCenter����,��̨����
    End If

End Function

Private Function GetImageRptData(ByVal lngOrderID As Long, ByVal lngStartIndex As Long, ByVal lngEndIndex As Long, blnMoved As Boolean) As ADODB.Recordset
'���ݱ���ͼ�� �ֶλ�ȡ���ͼ��
    Dim strSql As String

    strSql = "Select rownum As ˳���,a.ҽ��id,a.����,a.�Ա�,a.����, rownum As ͼ���,Replace(Trim(D.Column_Value),'.jpg','') as ͼ��UID, A.���UID, " & _
            "'' As ����UID, 0 as ��̬ͼ,'' as ��������,'' as �ɼ�ʱ��, '' as ¼�Ƴ���, '' as ����ͼ," & _
            "B.FTP�û��� As User1,B.FTP���� As Pwd1,B.IP��ַ As Host1,'/'||B.FtpĿ¼||'/' As Root1, " & _
            "B.����Ŀ¼ as ����Ŀ¼1,B.����Ŀ¼�û��� as ����Ŀ¼�û���1,B.����Ŀ¼���� as ����Ŀ¼����1, " & _
            "Decode(A.��������,Null,'',to_Char(A.��������,'YYYYMMDD')||'/') ||A.���UID||'/'||Replace(Trim(D.Column_Value),'.jpg','') As URL,B.�豸�� as �豸��1, B.�豸�� as �豸��1, " & _
            "C.FTP�û��� As User2,C.FTP���� As Pwd2,C.IP��ַ As Host2,'/'||C.FtpĿ¼||'/' As Root2, " & _
            "C.����Ŀ¼ as ����Ŀ¼2,C.����Ŀ¼�û��� as ����Ŀ¼�û���2,C.����Ŀ¼���� as ����Ŀ¼����2,C.�豸�� as �豸��2, C.�豸�� as �豸��2 " & _
            "From Ӱ�����¼ A, Ӱ���豸Ŀ¼ B, Ӱ���豸Ŀ¼ C, Table(Cast(f_Str2list(A.����ͼ��,';') As zlTools.t_Strlist)) D " & _
            "Where A.λ��һ = B.�豸��(+) And A.λ�ö� = C.�豸��(+) And A.ҽ��id = [1]"

    If blnMoved = True Then strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")


    strSql = "select * from (" & strSql & " order by ����UID, ͼ���) where ˳���>=" & lngStartIndex & " and ˳���<=" & lngEndIndex

    Set GetImageRptData = zlDatabase.OpenSQLRecord(strSql, "��ȡ����ͼ��", lngOrderID)
End Function

Private Function GetImageViewData(ByVal slQueryLevel As TQueryLevel, ByVal strSearchValue As String, _
    ByVal lngStartIndex As Long, ByVal lngEndIndex As Long, ByVal blnTmpRecord As Boolean, blnMoved As Boolean, blnOnlyLoadReportImage As Boolean) As ADODB.Recordset
'��ȡԤ��ͼ������

    Dim strSql As String
    

    strSql = "Select [2] ҽ��id,c.����,c.�Ա�,c.����, A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
        "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,D.����Ŀ¼ as ����Ŀ¼1,D.����Ŀ¼�û��� as ����Ŀ¼�û���1,D.����Ŀ¼���� as ����Ŀ¼����1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/') " & _
        "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, D.�豸�� As �豸��1," & _
        "E.FTP�û��� As User2,E.FTP���� As Pwd2,'' as ����ͼ," & _
        "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,E.����Ŀ¼ as ����Ŀ¼2,E.����Ŀ¼�û��� as ����Ŀ¼�û���2,E.����Ŀ¼���� as ����Ŀ¼����2," & _
        "E.�豸�� as �豸��2, E.�豸�� As �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+)" & IIf(blnOnlyLoadReportImage, " And nvl(A.��̬ͼ,0) = 0 ", "")

    
    If blnTmpRecord Then
        strSql = Replace(strSql, "Ӱ����", "Ӱ����ʱ")
    Else
        If blnMoved Then
            strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
            strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
        End If
    End If

    Select Case slQueryLevel
        Case slAdvice
            strSql = "Select rownum As ˳���,T.* From (" & strSql & " and C.ҽ��ID=[1] order by A.�ɼ�ʱ�� " & IIf(gblnShowImgDesc, " desc ", "") & " ) T"
        Case slStudy
            strSql = "Select rownum As ˳���,T.* From (" & strSql & " and C.���UID=[1] order by A.�ɼ�ʱ�� " & IIf(gblnShowImgDesc, " desc ", "") & " ) T"
        Case slSeries
            strSql = "Select rownum As ˳���,T.* From (" & strSql & " and B.����UID=[1] order by A.�ɼ�ʱ�� " & IIf(gblnShowImgDesc, " desc ", "") & " ) T"
        Case slImage
            strSql = "Select rownum As ˳���,T.* From (" & strSql & " and A.ͼ��UID=[1] order by A.�ɼ�ʱ�� " & IIf(gblnShowImgDesc, " desc ", "") & " ) T"
    End Select

    strSql = "select * from (" & strSql & ") where ˳���>=" & lngStartIndex & " and ˳���<=" & lngEndIndex

    Set GetImageViewData = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ����Ϣ", strSearchValue, IIf(blnTmpRecord, "-1", strSearchValue))
End Function


Private Function LoadViewImageToFaceWithNormal(rsCurImageData As ADODB.Recordset, ByRef arrImages() As String) As Boolean
'����Ԥ��ͼ�񵽽���
    Dim strTmpFile As String
    Dim strCachePath As String

    Dim curImage As DicomImage

    Dim objFile As New Scripting.FileSystemObject

    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp

    Dim dcmTag As clsImageTagInf
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim objImgInfo As Object
    Dim strItem As String

BUGEX "LoadViewImageToFaceWithNormal 1"

    blnIsAddImage = False
    LoadViewImageToFaceWithNormal = False

BUGEX "LoadViewImageToFaceWithNormal 2"

    '��������ͼ�񻺴�Ŀ¼
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsCurImageData("URL")))
    
    ReDim arrImages(0)
BUGEX "LoadViewImageToFaceWithNormal 3"
    Do While Not rsCurImageData.EOF
        'ѭ������ͼ��DicomViewer��
        strImgInstanceUid = Trim(Nvl(rsCurImageData!ͼ��UID))

        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then

            blnIsAddImage = True

            '��������Ƶ����ʾ�ļ������Ϊ����Ƶ�ļ�ʱ���ù��̽����ӷ�������ֱ�����������ļ�
            If Nvl(rsCurImageData!��̬ͼ, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf Nvl(rsCurImageData!��̬ͼ, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & Nvl(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If

            strTmpFile = Replace(Trim(strTmpFile), "/", "\")

            If Dir(strTmpFile) = vbNullString Then
                '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                '����FTP����
                If Nvl(rsCurImageData("�豸��1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(Nvl(rsCurImageData("Host1")), Nvl(rsCurImageData("User1")), Nvl(rsCurImageData("Pwd1"))) = 0 Then
                        If Nvl(rsCurImageData("�豸��2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))) = 0 Then
                                MsgBox "FTP�����������ӣ������������á�", vbOKOnly, CON_STR_HINT_TITLE
                                Exit Function
                            End If
                        Else
                            MsgBox "FTP�����������ӣ������������á�", vbOKOnly, CON_STR_HINT_TITLE
                            Exit Function
                        End If
                    End If
                End If

                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")) <> 0 Then
                    '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                    If Nvl(rsCurImageData("�豸��2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", ""))
                    End If
                End If
            End If
            
            
            
            strItem = strTmpFile & "|" & Nvl(rsCurImageData!��̬ͼ, imgTag) & "|" & Nvl(rsCurImageData!����ͼ) & "|" & _
                Nvl(rsCurImageData("��������"), "") & "|" & Nvl(rsCurImageData("�ɼ�ʱ��")) & "|" & strCachePath & Nvl(rsCurImageData("URL")) & ".avi" & "|" & _
                Val(Nvl(rsCurImageData("¼�Ƴ���"), "0")) & "|" & Nvl(rsCurImageData("ͼ��UID")) & "|" & Nvl(rsCurImageData("����UID")) & "|" & Nvl(rsCurImageData("���UID"))
            ReDim Preserve arrImages(UBound(arrImages) + 1)
            
            arrImages(UBound(arrImages)) = strItem
        End If
        rsCurImageData.MoveNext
    Loop
    
    LoadViewImageToFaceWithNormal = IIf(Trim(strCurInstanceUids) <> "" And blnIsAddImage = True, True, False)
    
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect

BUGEX "LoadViewImageToFaceWithNormal End"
End Function


Private Function LoadViewImageToFaceWithService(rsCurImageData As ADODB.Recordset, ByRef arrImages() As String) As Boolean
'ͨ��ZLPacsServerCenter�������Ԥ��ͼ�񵽽���
'����Ԥ��ͼ�񵽽���
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim dcmTag As clsImageTagInf
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim fileMsg As TransferFileMsg
    Dim blnIsSendOk As Boolean
    Dim strItem As String
    
    blnIsAddImage = False
    
    LoadViewImageToFaceWithService = False
        
    '��������ͼ�񻺴�Ŀ¼
    MkLocalDir GetResourceDir
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsCurImageData("URL")))
    
    ReDim arrImages(0)
    Do While Not rsCurImageData.EOF
        'ѭ������ͼ��DicomViewer��
        strImgInstanceUid = Trim(Nvl(rsCurImageData!ͼ��UID))
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then
            blnIsAddImage = True
            
            '��������Ƶ����ʾ�ļ������Ϊ����Ƶ�ļ�ʱ���ù��̽����ӷ�������ֱ�����������ļ�
            If Nvl(rsCurImageData!��̬ͼ, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf Nvl(rsCurImageData!��̬ͼ, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & Nvl(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If
            
            strTmpFile = Trim(strTmpFile)
            
            blnIsSendOk = True
            
            If Dir(strTmpFile) = vbNullString Then
                '���ػ���ͼ�񲻴��ڣ����ļ����ݷ�����������ʹ�÷����̨����
                With fileMsg
                    fileMsg.strAdviceId = Val(Nvl(rsCurImageData("ҽ��ID")))
                    fileMsg.strName = Nvl(rsCurImageData("����"))
                    fileMsg.strSex = Nvl(rsCurImageData("�Ա�"))
                    fileMsg.strAge = Nvl(rsCurImageData("����"))
                    
                    fileMsg.ftpInfo.strDeviceId = Nvl(rsCurImageData("�豸��1"))
                    fileMsg.ftpInfo.strFtpDir = Nvl(rsCurImageData("Root1"))
                    fileMsg.ftpInfo.strFTPIP = Nvl(rsCurImageData("Host1"))
                    fileMsg.ftpInfo.strFTPPwd = Nvl(rsCurImageData("Pwd1"))
                    fileMsg.ftpInfo.strFTPUser = Nvl(rsCurImageData("User1"))
                    fileMsg.ftpInfo.strSDDir = Nvl(rsCurImageData("����Ŀ¼1"))
                    fileMsg.ftpInfo.strSDPswd = Nvl(rsCurImageData("����Ŀ¼����1"))
                    fileMsg.ftpInfo.strSDUser = Nvl(rsCurImageData("����Ŀ¼�û���1"))
                    
                    fileMsg.bakFtpInfo.strDeviceId = Nvl(rsCurImageData("�豸��2"))
                    fileMsg.bakFtpInfo.strFtpDir = Nvl(rsCurImageData("Root2"))
                    fileMsg.bakFtpInfo.strFTPIP = Nvl(rsCurImageData("Host2"))
                    fileMsg.bakFtpInfo.strFTPPwd = Nvl(rsCurImageData("Pwd2"))
                    fileMsg.bakFtpInfo.strFTPUser = Nvl(rsCurImageData("User2"))
                    fileMsg.bakFtpInfo.strSDDir = Nvl(rsCurImageData("����Ŀ¼2"))
                    fileMsg.bakFtpInfo.strSDPswd = Nvl(rsCurImageData("����Ŀ¼����2"))
                    fileMsg.bakFtpInfo.strSDUser = Nvl(rsCurImageData("����Ŀ¼�û���2"))
                    
                    fileMsg.strLocalDir = strTmpFile
                    fileMsg.strFileName = Nvl(rsCurImageData("ͼ��UID")) & IIf(mblnIsLoadReportImage, ".jpg", "")
                    fileMsg.strSubDir = Nvl(rsCurImageData("URL"))
                    fileMsg.strMediaType = Nvl(rsCurImageData!��̬ͼ, imgTag)
                End With
                
                If Not SendDataToservice("����ͼ", LoadCommand.COMMAND_RPTIMG_DOWNLOAD, "ͼ������", fileMsg) Then
                    blnIsSendOk = False
                End If
            End If
            
            strItem = strTmpFile & "|" & Nvl(rsCurImageData!��̬ͼ, imgTag) & "|" & Nvl(rsCurImageData!����ͼ) & "|" & _
                Nvl(rsCurImageData("��������"), "") & "|" & Nvl(rsCurImageData("�ɼ�ʱ��")) & "|" & strCachePath & Nvl(rsCurImageData("URL")) & ".avi" & "|" & _
                Val(Nvl(rsCurImageData("¼�Ƴ���"), "0")) & "|" & Nvl(rsCurImageData("ͼ��UID")) & "|" & Nvl(rsCurImageData("����UID")) & "|" & Nvl(rsCurImageData("���UID"))
            ReDim Preserve arrImages(UBound(arrImages) + 1)
            arrImages(UBound(arrImages)) = strItem
        End If
        
        rsCurImageData.MoveNext
    Loop
    
    LoadViewImageToFaceWithService = IIf(Trim(strCurInstanceUids) <> "" And blnIsAddImage = True, True, False)
End Function

Private Function SendDataToservice(ByVal strDataTag As String, ByVal intCommandIdentify As Integer, ByVal strDataFrom As String, fileMsg As TransferFileMsg)
    Dim objServiceHelper As New clsServiceHelper
    
    SendDataToservice = objServiceHelper.SendDataToservice(strDataTag, intCommandIdentify, strDataFrom, fileMsg)
    
    Set objServiceHelper = Nothing
End Function

Public Function GetRpRecordCount(ByVal slQueryLevel As TQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean, blnMoved As Boolean) As Long
'��ȡͼ������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    strSql = "Select Count(B.Column_Value) ����ֵ From Ӱ�����¼ A, Table(Cast(f_Str2list(Replace(A.����ͼ��,';',',')) As zlTools.t_Strlist)) B Where ҽ��ID = [1]"
    '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
    If blnTmpRecord Then
        strSql = Replace(strSql, "Ӱ����", "Ӱ����ʱ")
    Else
        If blnMoved Then
            strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        End If
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ������", strSearchValue)
    If rsData.RecordCount > 0 Then lngRecordCount = Nvl(rsData!����ֵ)
    
    If lngRecordCount <= 0 Then
        Select Case slQueryLevel
            Case slAdvice
                strSql = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and nvl(a.��̬ͼ,0)=0 and c.ҽ��ID=[1]"
            Case slStudy
                strSql = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b where a.����UID=b.����UID and nvl(a.��̬ͼ,0)=0 and b.���UID=[1]"
            Case slSeries
                strSql = "select count(1)  as ����ֵ from Ӱ����ͼ��  where nvl(��̬ͼ,0)=0 and ����UID=[1]"
            Case slImage
                strSql = "select count(1)  as ����ֵ from Ӱ����ͼ��  where nvl(��̬ͼ,0)=0 and ͼ��UID=[1]"
        End Select
        
        '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
        If blnTmpRecord Then
            strSql = Replace(strSql, "Ӱ����", "Ӱ����ʱ")
        Else
            If blnMoved Then
                strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
                strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
            End If
        End If
    
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ������", strSearchValue)
        
        If rsData.RecordCount > 0 Then
            lngRecordCount = Nvl(rsData!����ֵ)
        Else
            lngRecordCount = 0
        End If
    End If
    
    GetRpRecordCount = lngRecordCount
End Function

Public Function GetRecordCount(ByVal slQueryLevel As TQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean, ByVal blnMoved As Boolean) As Long
'��ȡͼ������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    Select Case slQueryLevel
        Case slAdvice
            strSql = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=[1]"
        Case slStudy
            strSql = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b where a.����UID=b.����UID and b.���UID=[1]"
        Case slSeries
            strSql = "select count(1)  as ����ֵ from Ӱ����ͼ��  where  ����UID=[1]"
        Case slImage
            strSql = "select count(1)  as ����ֵ from Ӱ����ͼ��  where  ͼ��UID=[1]"
    End Select
    
    '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
    If blnTmpRecord Then
        strSql = Replace(strSql, "Ӱ����", "Ӱ����ʱ")
    Else
        If blnMoved Then
            strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        End If
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ������", strSearchValue)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = Nvl(rsData!����ֵ)
    Else
        lngRecordCount = 0
    End If
    
    GetRecordCount = lngRecordCount
End Function

Public Function GetImageIdex(ByVal slQueryLevel As TQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean, ByVal blnMoved As Boolean, strUid As String) As Long
'����ͼ��UID��ȡ����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    Select Case slQueryLevel
        Case slAdvice
            strSql = "select a.ͼ��UID,a.����UID,a.ͼ��� from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=[1]"
        Case slStudy
            strSql = "select a.ͼ��UID,a.����UID,a.ͼ��� from Ӱ����ͼ�� a, Ӱ�������� b where a.����UID=b.����UID and b.���UID=[1]"
        Case slSeries
            strSql = "select a.ͼ��UID,a.����UID,a.ͼ��� from Ӱ����ͼ��  where  ����UID=[1]"
        Case slImage
            strSql = "select a.ͼ��UID,a.����UID,a.ͼ��� from Ӱ����ͼ��  where  ͼ��UID=[1]"
    End Select
    strSql = "select * from(select rownum as ˳���,T.* from ( " & strSql & " order by a.�ɼ�ʱ�� " & IIf(gblnShowImgDesc, " desc ", "") & " ) T ) where ͼ��UID = [2]"
    '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
    If blnTmpRecord Then
        strSql = Replace(strSql, "Ӱ����", "Ӱ����ʱ")
    Else
        If blnMoved Then
            strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        End If
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ������", strSearchValue, strUid)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = Nvl(rsData!˳���)
    Else
        lngRecordCount = 0
    End If
    
    GetImageIdex = lngRecordCount
End Function
