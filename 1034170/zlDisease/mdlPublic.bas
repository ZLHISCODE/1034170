Attribute VB_Name = "mdlPublic"
Option Explicit

Public gstrUnitName As String       '��ǰ�û���λ����
Public gfrmMain As Object           '����̨����
Public gcnOracle As ADODB.Connection  '���ݿ�����
Public gstrSysName As String                'ϵͳ���ƣ����磺�������
Public gstrProductName As String            '��Ʒ��ƣ����磺����
Public glngModul As Long                    'ģ����
Public glngSys As Long                      'ϵͳ��ţ����磺100
Public gstrDBUser As String
Public gstrPrivs As String                     '�û��ڸ�ģ�������Ȩ��
Public gblnShowInTaskBar As Boolean         '�Ƿ���ʾ��������������
Public UserInfo As TYPE_USER_INFO            '�û���Ϣ
Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Public gMainPrivs As String
Public gstrNodeNo As String          '��ǰվ���ţ����δ��������վ�㣬��Ϊ"-"
Private mclsZip As New cZip
Private mclsUnzip As New cUnzip
Public gclsMipModule As zl9ComLib.clsMipModule
Public gstrLike As String  '��Ŀƥ�䷽��,%���
Public gbytCode As Byte '�������뷽ʽ
Public gstrDBOwer As String
Public gobjComlib As zl9ComLib.clsComLib
Public glngPreHWnd As Long '����֧�������ֹ���
Public glngOpenedID As Long 'ҽ��վ����ʱ�򿪵ķ�����ID
Public gObjRichEPR As zlRichEPR.cRichEPR

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'�ı䴰��λ�á�Zorder���ߴ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Public Const WM_MOUSEWHEEL = &H20A

Public Type TYPE_USER_INFO
    ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    ����ID As Long
    ������ As String
    ������ As String
    ��ҩ���� As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Type POINTAPI
        X As Long
        Y As Long
End Type

'Public Enum zlEnumDClick
'    cprEmDClickApplyTo = 1         '˫�����ÿ���
'    cprEmDClickRequest = 2         '˫��ʱ��Ҫ��
'End Enum

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    Dim strTmp As String
    gstrLike = IIf(gobjComlib.zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    gbytCode = Val(gobjComlib.zlDatabase.GetPara("���뷽ʽ"))
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.����ID = NVL(rsTmp!����ID, 0)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.���� = Get��Ա����
            GetUserInfo = True
        End If
    End If
    Exit Function
errH:
   If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
'���ܣ���ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    If str���� <> "" Then
        strSQL = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str����)
    Else
        strSQL = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## ���ܣ�  �����ݴ�һ��XtremeReportControl�ؼ����Ƶ�VSFlexGrid���Ա���д�ӡ
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '��ȫ����ǿ��չ��,�������ݱ��
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow

    Dim lngCol As Long, lngRow As Long

    On Error GoTo errHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        '�����и���
        .Row = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = rptCol.Width * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol

        '�����и���
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function
errHand:
    zlReportToVSFlexGrid = False
End Function
'
Public Function DynamicCreate(ByVal strClass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strClass)
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'���ܣ��ж�ָ������֮ǰ���Ƿ�����Ѿ�ִ��������ת��
'������vDate=ʱ����ʱ��εĿ�ʼʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "Select �ϴ����� From zlDataMove Where ϵͳ=[1] And ���=1 And �ϴ����� is Not Null"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '�ϴ�����û��ʱ��,"<"�ж���ת��������һ��
        If vDate < rsTmp!�ϴ����� Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetDbOwner(ByVal lngSys As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    GetDbOwner = ""
    On Error GoTo errHand
    strSQL = "Select ������ From Zlsystems Where ��� = [1]"
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetDbOwner", lngSys)
    If rsTemp.RecordCount <> 0 Then GetDbOwner = "" & rsTemp!������
    rsTemp.Close
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'################################################################################################################
'## ���ܣ�  ��ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  ������ݵ��ļ�����ʧ���򷵻��㳤��""
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    Dim strSQL As String

    Err = 0: On Error GoTo errHand

    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum

    strSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as Ƭ�� From Dual"
    lngCount = 0
    Do
        Set rsLob = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value

        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next

        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile
    If ErrCenter = 1 Then
        Resume
    End If
    zlBlobRead = ""
End Function


'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String

    On Error GoTo errHand

    If Not gobjFSO.FileExists(strZipFile) Then zlFileUnzip = "": Exit Function

    strZipPath = gobjFSO.GetSpecialFolder(2) 'ȡ��ʱĿ¼
    strZipPathTmp = strZipPath & "\" & Format(Now, "yyMMdd") & CStr(100 * Timer)
    If Not gobjFSO.FolderExists(strZipPathTmp) Then Call gobjFSO.CreateFolder(strZipPathTmp)

    strZipFileTmp = strZipPathTmp & "\TMP.RTF"
    If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp

    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FileExists(strZipFileTmp) Then

        strZipFileName = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer) & ".RTF"
        If gobjFSO.FileExists(strZipFileName) Then gobjFSO.DeleteFile strZipFileName

        Call gobjFSO.CopyFile(strZipFileTmp, strZipFileName)
        If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp, True
        On Error Resume Next
        If gobjFSO.FolderExists(strZipPathTmp) Then gobjFSO.DeleteFolder strZipPathTmp, True

        zlFileUnzip = strZipFileName
    Else
        zlFileUnzip = ""
    End If
    Exit Function
errHand:
    Call SaveErrLog
End Function


'################################################################################################################
'## ���ܣ�  �滻����Ҫ�صĴ���
'##
'## ������  ElementName     :�滻��Ŀ������
'##         sPatientID      :����ID
'##         sPageID         :��ҳID��Һ�id
'##         iPatientType    :0=���1=סԺ
'##         lngҽ��ID       :ҽ��ID
'##
'## ���أ�  �����滻���
'################################################################################################################
Public Function GetReplaceEleValue(ByVal ElementName As String, _
    ByVal sPatientID As String, _
    ByVal sPageID As String, _
    ByVal iPatientType As PatiFromEnum, _
    ByVal lngҽ��id As Long, Optional lngBabyNum As Long) As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

    strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5],[6]) From Dual"
    Err = 0: On Error GoTo DBError
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ�滻��", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lngҽ��id, lngBabyNum)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value))
    End If
    Exit Function
DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function


'################################################################################################################
'## ���ܣ�  ���������ı�����ָ���ؼ�������Ķ�λ��Ϣ
'##
'## ������  edtThis         :   IN  ���༭�ؼ�
'##         strKeyType      :   IN  �������ؼ������ơ�ȡֵΪ��"O"��"P"��"T"��"E"��"U"
'##         lngKey           :   IN  �����������ҵĹؼ���ID�š�
'##         lngKSS��lngKSE  :   OUT ���ֱ��ʾ��ʼ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         lngKES��lngKEE  :   OUT ���ֱ��ʾ��ֹ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         blnNeeded:      :   OUT ���Ƿ��Ǳ�������
'##
'## ���أ�  ����ҵ��ùؼ��־���λ�ã��򷵻�True�����򷵻�False
'################################################################################################################
Public Function FindKey(ByRef edtThis As Object, _
        ByRef strKeyType As String, _
        ByRef lngKey As Long, _
        ByRef lngKSS As Long, _
        ByRef lngKSE As Long, _
        ByRef lngKES As Long, _
        ByRef lngKEE As Long, _
        ByRef blnNeeded As Boolean) As Boolean

    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��

    sTMP = strKeyType & "S(" & Format(lngKey, "00000000")
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = 1
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���

            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = strKeyType & "E(" & Format(lngKey, "00000000")
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindKey = True
            End If
        End If
    End With
End Function


Public Function GetControlRect(ByVal lngHwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip/Pixel)
'���أ�blnTwip=True-����Twip��λ��False-�������ص�λ
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function

Public Function RPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'���ܣ���ͬOracle��RPAD����
'����:��ָ���������ƿո�
 '������
 '       strText:����ַ���
 '       intCount:����ĳ���
 '       StrPAD:�����ַ�
 '       blnAutoSub:�ַ����������Զ���ȡ
'����:�����ִ�
   
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    
    lngFill = LenB(StrConv(StrPAD, vbFromUnicode))
    lngTmp = LenB(StrConv(strText, vbFromUnicode))
    If lngTmp <= intCount - lngFill Then
        RPAD = strText & String((intCount - lngTmp) \ lngFill, StrPAD)
    ElseIf lngTmp > intCount And blnAutoSub Then
        RPAD = SubB(strText, 1, intCount)
    Else
        RPAD = strText
    End If
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
 '���:strInfor-ԭ��
 '         lngStart-ֱʼλ��
'         lngLen-����
'����:�Ӵ�
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo errH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
errH:
    Err.Clear
    SubB = ""
End Function


Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun���ڸú���
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function


Public Sub FormLock(Optional ByVal lngHwnd As Long)
'���ܣ�������������ˢ�¡����߽������
'������lngHwnd=0-�������,<>0Ҫ���������Hwnd
    LockWindowUpdate (lngHwnd)
End Sub


Public Function RowValue(ByVal strTable As String, Optional ByVal arrValues As Variant, Optional ByVal strGetFields As String, Optional ByVal strWhereField As String = "ID") As Variant
'���ܣ���ȡָ����ָ���ֶ���Ϣ
'������strTable=��ȡ���ݵı�
'          arrValues=����ֵ�����Դ����飬Ҳ���Դ�����ֵ��Ҳ���Բ�����������ȡȫ��
'          strGetField=��ȡ���ֶ�,����ֶ��Զ��ŷָͬSQL��д��ȡ�ֶ�һ��
'          strWhereField=���˵��ֶΣ����ֶ�Ϊ�򵥵���ֵ���ַ����ͻ��������ͣ����������޷�֧��
'���أ�
'ֻ������һ����������ض���һ��ֵ��δ����NULLֵ����
'      strGetField=�����ֶ�
'      arrValues=Ϊ����ֵ���򲻸���һ��Ԫ�ص�����
'������������ؼ�¼��

    Dim rsTmp As New ADODB.Recordset, blnReturnRec As Boolean
    Dim strSQL As String
    Dim strWhere As String
    Dim arrPars As Variant
    Dim i As Long, strPars As String
    
    On Error GoTo errH
    blnReturnRec = True
    If TypeName(arrValues) = "Variant()" Then
        arrPars = arrValues
        For i = LBound(arrValues) To UBound(arrValues)
            strPars = strPars & ",[" & i + 1 & "]"
        Next
        If strGetFields <> "" Then '�������Ԫ�ز�����һ��,�һ�ȡ����Ԫ�أ��򲻷��ؼ�¼��
            If UBound(arrValues) - LBound(arrValues) + 1 <= 1 And Not strGetFields Like "*,*" Then blnReturnRec = False
        End If
        If strPars <> "" Then
            strWhere = " Where " & strWhereField & " In (" & strPars & ")"
        End If
    ElseIf TypeName(arrValues) <> "Error" Then
        '����ֵʱ������ȡ�����ֶΣ��򲻷��ڼ�¼��
         If strGetFields <> "" And Not strGetFields Like "*,*" Then blnReturnRec = False
        arrPars = Array(arrValues)
        strWhere = " Where " & strWhereField & "=[1]"
    Else
        strWhere = ""
    End If
    
    If strGetFields = "" Then strGetFields = "*"
    strSQL = "Select " & strGetFields & " From " & strTable & strWhere
    If strWhere <> "" Then
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "RowValue", arrPars)
    Else
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "RowValue")
    End If
    If blnReturnRec Then
        Set RowValue = rsTmp
    Else
        If Not rsTmp.EOF Then
            RowValue = rsTmp.Fields(strGetFields).Value
        Else '��ȡ��ֵʱ��û�л�ȡ����ֵ���򷵻�Ĭ��ֵ
            If IsType(rsTmp.Fields(strGetFields).Type, adVarChar) Then
                RowValue = ""
            ElseIf IsType(rsTmp.Fields(strGetFields).Type, adInteger) Then
                RowValue = 0
            Else
                RowValue = Null
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function

Public Function GetDeptID(ByVal strDeptCode As String) As Long
'���ܣ����ݲ��ű����ȡ����ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
On Error GoTo errH
    strSQL = "Select a.Id, a.���� From ���ű� A Where A.���� = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "���Ų�ѯ", strDeptCode)
    If rsTmp.RecordCount > 0 Then
        GetDeptID = rsTmp!ID
    Else
        GetDeptID = 0
        MsgBox "û�в�ѯ������Ϊ��" & strDeptCode & "���Ĳ��ſ��ң�����ϵ����Ա���룡", vbInformation, gstrSysName
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'���ܣ�֧�ֹ��ֵĹ���
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '���¹�
            gobjComlib.ZLCommFun.PressKey vbKeyPageDown
        Case 7864320   '���Ϲ�
            gobjComlib.ZLCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
End Function

Public Function CheckOperateState(ByVal lngID As Long, ByRef intCode As Integer) As Boolean
'����: ��ѯ�Ƿ��ܹ�����÷�������ɾ�������޸ģ�
'����: lngID - ������ID ��intCode - ���ܲ�����ԭ�� ��1-δ���ҵ���2-���˵ķ�������3-ҽ���Ѿ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '��ȡ�����������Ϣ
    On Error GoTo errH
    strSQL = "Select a.Id, a.��¼״̬, a.�Ǽ��� From �������Լ�¼ A  Where a.Id = [1] "

    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ���Խ��������", lngID)
    
    If rsTmp.RecordCount > 0 Then
        If UserInfo.���� <> NVL(rsTmp!�Ǽ���) Then
            intCode = 2
            Exit Function
        ElseIf rsTmp!��¼״̬ > 1 Then
            intCode = 3
            Exit Function
        End If
    Else
        intCode = 1
        Exit Function
    End If
    CheckOperateState = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub PrintDiseaseRegist(ByVal intType As Integer, ByVal lngID As Long, ByRef frmParent As Object)
'����: ��ӡ���Խ��������
'������lngID : ������ID��intType:1-Ԥ����2-��ӡ
    Dim objReport As clsReport
    
    On Error GoTo errH
  
    If objReport Is Nothing Then Set objReport = New clsReport
    Call objReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1278_1", frmParent, "������ID=" & lngID, intType)
    If Not objReport Is Nothing Then Set objReport = Nothing
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckDisNum(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByRef rsDisease As ADODB.Recordset, Optional ByVal lngID As Long) As Boolean
'����: ���ò����ж���û����д���濨�ķ�����
'lngPatFrom: 2-סԺ, 1-����
    Dim strSQL As String
    On Error GoTo errH
    If lngID <> 0 Then
        strSQL = " and a.ID = " & lngID
    End If
    If lngPatFrom = 1 Then
        strSQL = "select rowNum as NO,a.ID,c.���� as ����, a.�Ǽ�ʱ��, a.��¼״̬, a.�������˵�� from  �������Լ�¼ A ,���˹Һż�¼ B ,���ű� C where A.�ļ�ID is NULL  and A.�Һŵ� = B.NO and A.����ID = B.����ID and A.��¼״̬ <> 3 and A.�Ǽǿ���ID = C.ID  and A.����ID = [1] and B.ID = [2]" & strSQL
    ElseIf lngPatFrom = 2 Then
        strSQL = "select rowNum as NO,a.ID ,c.���� as ����,a.�Ǽ�ʱ��, a.��¼״̬, a.�������˵�� from  �������Լ�¼ A ,���ű� C  where A.�ļ�ID is NULL  and A.��¼״̬ <> 3  and A.�Ǽǿ���ID = C.ID and A.����ID = [1] and A.��ҳID = [2] " & strSQL
    End If
    Set rsDisease = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ���Խ��������", lngPatiID, lngPageId)
    
    If rsDisease.RecordCount > 0 Then
        CheckDisNum = True
    Else
        CheckDisNum = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SaveReason(ByVal strReason As String, ByVal lngID As Long, ByVal lngState As Long) As Boolean
'����: �洢����д���濨��ԭ��
'������strReason-ԭ��lngID-������ID ��lngState-��������ǰ�ļ�¼״̬
    Dim strSQL As String
    Dim str����ʱ�� As String
    Dim str����ҽ�� As String
    Dim str������� As String, strTmp As String

    On Error GoTo errH
    str����ʱ�� = "to_date('" & Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    str����ҽ�� = "'" & UserInfo.���� & "'"
    str������� = "'" & strReason & "'"
    lngState = IIf(lngState = 1, 2, lngState)

    strSQL = "Zl_�������Լ���¼_update(1," & lngID & "," & "NULL" & "," & CStr(lngState) & "," & str����ҽ�� & "," & str����ʱ�� & "," & str������� & ")"
    Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, "���淴�����Ĵ������")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
