Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gstrSysName As String                'ϵͳ����
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String    '��Ʒ����

Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gobjXWHIS As Object     'RIS�ӿڲ���zl9XWInterface.clsHISInner
Public gblnXW As Boolean      'ϵͳ������������ҽѧӰ����Ϣϵͳרҵ��ӿڡ�
Public gobjLIS As Object     'Lis����
Public gobjPlugIn As Object    '�������

Public glngSys As Long
Public glngModule As Long
Public gMainPrivs As String
Public gstrDBUser As String
Public gstrNodeNo As String          '��ǰվ���ţ����δ��������վ�㣬��Ϊ"-"
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��

Public gstrLike As String  '��Ŀƥ�䷽��,%���
Public gblnMyStyle As Boolean 'ʹ�ø��Ի����
Public gstrIme As String '�Զ��Ŀ������뷨
Public gbytCode As Byte '�������ɷ�ʽ��0-ƴ��,1-���,2-����
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
    רҵ����ְ�� As String
    ��ҩ���� As Long
End Type
Public UserInfo As TYPE_USER_INFO

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    '����27554 by lesfeng 2010-01-19 lngTXTProc �޸�ΪglngTXTProc
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function MoveObj(lngHwnd As Long) As RECT
'���ܣ��ڶ����MouseDown�¼��е���,����������Hwnd����
'���أ������Ļ������ֵ
   
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional strName As String = "")
'���ܣ���PictureBoxģ���3Dƽ�水ť
'������intStyle:0=ƽ��,-1=����,1=͹��
    
    Dim picRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            picRect.Left = .ScaleLeft
            picRect.Top = .ScaleTop
            picRect.Right = .ScaleWidth
            picRect.Bottom = .ScaleHeight
            DrawEdge .hdc, picRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Sub AutoSizeCol(lvw As Object)
'���ܣ������Զ�ListView��ǰ�����Զ��������п��
'������blnByHead=�Ƿ���ͷ�ı�����,Col=ָ���л���������(1-N)
    Dim i As Integer, lngW As Long
    For i = 1 To lvw.ColumnHeaders.Count
        SendMessage lvw.hWnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If lvw.ColumnHeaders(i).Width < 200 Then lvw.ColumnHeaders(i).Width = 0
        If lvw.ColumnHeaders(i).Width < (gobjCommFun.ActualLen(lvw.ColumnHeaders(i).Text) + 2) * 90 And lvw.ColumnHeaders(i).Width <> 0 Then lvw.ColumnHeaders(i).Width = (gobjCommFun.ActualLen(lvw.ColumnHeaders(i).Text) + 2) * 90
    Next
End Sub

Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 0.5) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������lngHwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�-2=δ�Ӵ���,����=ƥ�������(����ƥ�������)
'˵�����뽫�ú�����KeyPress�¼��е��á�

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
        
    On Error Resume Next
        
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '������(ȱʡΪ0.5��)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 'ʹComboBox����ĵ���ƥ�书��ʧЧ
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '������Իس���������
    End If
End Function

Public Function GetCboIndex(cbo As ComboBox, strFind As String, Optional blnKeep As Boolean, Optional blnLike As Boolean) As Long
'���ܣ����ַ�����ComboBox�в�������
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '�Ⱦ�ȷ����
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), "-") > 0 Then
            If NeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '���ģ������
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Public Function FindCboIndex(cbo As ComboBox, lngID As Long) As Long
'���ܣ�����Ŀֵ����ComboBox����Ŀ����
    Dim i As Integer
    If lngID = 0 Then FindCboIndex = -1: Exit Function
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = lngID Then
            FindCboIndex = i
            Exit Function
        End If
    Next
    FindCboIndex = -1
End Function

Public Function SetWidth(cboHwnd As Long, NewWidthPixel As Long) As Boolean
'���ܣ����� Combo �����Ŀ��,��λΪ pixels
    Dim lRetVal As Long
    lRetVal = SendMessage(cboHwnd, &H160, NewWidthPixel, 0)
    If lRetVal <> -1 Then
        SetWidth = True
    Else
        SetWidth = False
    End If
End Function

Public Function GetWidth(cboHwnd As Long) As Long
'���ܣ� ȡ�� Combo �����Ŀ��,��λΪ pixels
    Dim lRetVal As Long
    lRetVal = SendMessage(cboHwnd, &H15F, 0, 0)
    If lRetVal <> -1 Then
        GetWidth = lRetVal
    Else
        GetWidth = 0
    End If
End Function

Public Sub SelAll(objTxt As Control)
'���ܣ����ı���ĵ��ı�ѡ��
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function SetCboDefault(cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = 1 Then
            cbo.ListIndex = i
            SetCboDefault = i: Exit Function
        End If
    Next
End Function

Public Function CheckLen(txt As TextBox, intLen As Integer) As Boolean
'���ܣ���鹤�������ʵ�����Ƿ���ָ�����Ƴ�����
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox Mid(txt.Name, 4) & "ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�", vbExclamation, gstrSysName
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(gobjDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function CaptionHeight() As Long
'����:����ϵͳ����������߶�(������Ϊ��λ)
    CaptionHeight = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(&HA)) > 0 Then
        NeedName = Mid(strList, InStr(strList, Chr(&HA)) + 1)
    Else
        NeedName = Mid(strList, InStr(strList, "-") + 1)
    End If
    '51299,������,2012-07-17
    If InStr(NeedName, Chr(&HD)) > 0 Then
        NeedName = Replace(NeedName, Chr(&HD), "")
    End If
End Function

Public Sub SetItemInfo(lvw As Object, pan As Object)
'���ܣ�����Listview��ǰѡ���У���ʾ��״̬����
    Dim i As Integer, strInfo As String
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If lvw.SelectedItem.Text <> "" Then
        strInfo = "/" & lvw.ColumnHeaders(1).Text & ":" & lvw.SelectedItem.Text
    End If
    
    For i = 2 To lvw.ColumnHeaders.Count
        If lvw.SelectedItem.SubItems(i - 1) <> "" Then
            strInfo = strInfo & "/" & lvw.ColumnHeaders(i).Text & ":" & lvw.SelectedItem.SubItems(i - 1)
        End If
    Next
    If strInfo <> "" Then pan.Text = Mid(strInfo, 2)
End Sub

Public Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function

Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If gobjCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function CheckFormInput(objForm As Object, Optional ByVal strToNumText As String = "") As Boolean
'����:strToNumText--��Ҫ���н�ǧ��λ��ʽ�Ľ��ת����������ʽ���ı��ؼ�����,�����ж��,����,�ŵȷָ�
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled And Not obj.Locked Then
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                    If InStr(1, "," & UCase(strToNumText) & ",", "," & UCase(obj.Name) & ",") > 0 Then
                        strText = StrToNum(strText)
                    End If
                Case "ComboBox"
                    If obj.Style = 0 Then strText = obj.Text
                End Select
                If InStr(strText, "'") > 0 _
                    Or InStr(strText, ",") > 0 _
                    Or InStr(strText, ";") > 0 _
                    Or InStr(strText, "|") > 0 _
                    Or InStr(strText, "~") > 0 _
                    Or InStr(strText, "^") > 0 Then
                    MsgBox "���������а����Ƿ��ַ���", vbInformation, gstrSysName
                    obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                    obj.SetFocus: Exit Function
                End If
            End If
        End If
    Next
    CheckFormInput = True
End Function

Public Function StrToNum(ByVal strNumber As String) As Double
    '����:���ַ���ת��������
    Dim strTmp As String
    strTmp = Replace(strNumber, ",", "")
    StrToNum = Val(strTmp)
End Function

Public Function GetIDDate(ID As String) As String
'���ܣ��������֤�ŷ��س�������,��ʽ"yyyy-MM-dd"
'������ID=���֤��,Ӧ��Ϊ15λ��18λ
    Dim strTmp As String
    
    If Len(ID) = 15 Then
        strTmp = Mid(ID, 7, 6)
        If Len(strTmp) = 6 And IsNumeric(strTmp) Then
            strTmp = "19" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2)
        End If
    ElseIf Len(ID) = 18 Then
        strTmp = Mid(ID, 7, 8)
        If Len(strTmp) = 8 And IsNumeric(strTmp) Then
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2)
        End If
    End If
    If IsDate(strTmp) Then GetIDDate = strTmp
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function


Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function
Public Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function
Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub HideFormCaption(ByVal lngHwnd As Long, Optional ByVal blnBorder As Boolean = True)
'���ܣ�����һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(lngHwnd, vRect)
    lngStyle = GetWindowLong(lngHwnd, GWL_STYLE)

    If blnBorder Then
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
    Else
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    End If
    SetWindowLong lngHwnd, GWL_STYLE, lngStyle
    SetWindowPos lngHwnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
'������blnForceNum=��ΪNullʱ���Ƿ�ǿ�Ʊ�ʾΪ������
    ZVal = IIf(Val(varValue) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varValue))
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.����ID = Nvl(rsTmp!����ID, 0)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.���� = Get��Ա����
            UserInfo.רҵ����ְ�� = Getרҵ����ְ��(UserInfo.ID)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.�û���
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Getרҵ����ְ��(ByVal lng��Աid As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��¼��Ա��רҵ����ְ��
    '����:����ָд��Ա��רҵ����ְ��
    '����:���˺�
    '����:2014-04-09 13:45:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
 
    strSQL = "Select רҵ����ְ�� From ��Ա�� Where ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Աרҵְ��", lng��Աid)
    
    Getרҵ����ְ�� = "" & rsTmp!רҵ����ְ��
  
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    '����:������Ա����,����ö��ŷ���
    '����:���˺�
    '����:2014-04-09 13:46:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    If str���� <> "" Then
        strSQL = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Ա����", str����)
    Else
        strSQL = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Ա����", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������ض���
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    Call gobjComlib.InitCommon(gcnOracle)
    Set gobjCommFun = gobjComlib.gobjCommFun
    Set gobjControl = gobjComlib.gobjControl
    Set gobjDatabase = gobjComlib.gobjDatabase
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
End Function

Public Sub InitLocPar()
    Dim strValue As String
    
    On Error Resume Next
    gstrLike = IIf(gobjDatabase.GetPara("����ƥ��") = 0, "%", "")
    strValue = gobjDatabase.GetPara("���뷨")
    gstrIme = IIf(strValue = "", "���Զ�����", strValue)
    gbytCode = Val(gobjDatabase.GetPara("���뷽ʽ"))
    gblnMyStyle = gobjDatabase.GetPara("ʹ�ø��Ի����") = "1"
    gblnXW = Val(gobjDatabase.GetPara(255, glngSys)) = 1
    
    If Err <> 0 Then Err.Clear
End Sub

Public Function Between(X, a, B) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < B Then
        Between = X >= a And X <= B
    Else
        Between = X >= B And X <= a
    End If
End Function

Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Public Function GetItemField(ByVal strTable As String, ByVal lngID As Long, Optional ByVal strField As String) As Variant
'���ܣ���ȡָ����ָ���ֶ���Ϣ
'˵����δ����NULLֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If strField = "" Then
        strSQL = "Select * From " & strTable & " Where ID=[1]"
    Else
        strSQL = "Select " & strField & " From " & strTable & " Where ID=[1]"
    End If
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngID)
    If Not rsTmp.EOF Then
        If strField = "" Then
            Set GetItemField = rsTmp
        Else
            GetItemField = rsTmp.Fields(strField).Value
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetColumnLength(strTable As String, strColumn As String) As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(Data_Precision, Data_Length) collen From All_Tab_Columns Where Table_Name = [1] And Column_Name = [2]"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, strTable, strColumn)
    GetColumnLength = Val("" & rsTmp!collen)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Long, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = gobjComlib.GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Sub LoadOldData(ByVal strOld As String, ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox)
'����:�����ݿ��б�������䰴�淶�ĸ�ʽ���ص�����,���淶��ԭ����ʾ
    Dim strTmp As String, lngIdx As Long
    
    If Trim(strOld) = "" Then Exit Sub
    
    lngIdx = -1
    strTmp = strOld
    If InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 0
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 1
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 2
        End If
    ElseIf IsNumeric(strOld) Then
        lngIdx = 0
    End If
    txt����.Text = strTmp
    If cbo���䵥λ.ListCount > 0 Then Call gobjControl.CboSetIndex(cbo���䵥λ.hWnd, lngIdx)
    If lngIdx = -1 Then
        cbo���䵥λ.Visible = False
    Else
        If cbo���䵥λ.Visible = False Then cbo���䵥λ.Visible = True
    End If
End Sub

Public Function RecalcBirth(ByVal strAge As String, ByRef strDateOfBirth As String, Optional ByVal strCalcDate As String, Optional ByRef strMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ��������ȡ���˳�������
    '���:strAge:��������,�磺23�ꡢ1��2��
    'strCalcDate-�����������
    '����:����Ĳ��������ʽ��ȷ����㷵�س�������,���򷵻ؿ�
    'strMsg-���ؾ�ʾ��Ϣ
    '��ȷ�����ʽ:X��[X��]��X��[X��]��X�졢XСʱ[X����]
    '    X��:X���ܴ���200,X��:X���ܴ���12,X��:X���ܴ���31,XСʱ:X���ܴ���24,X����:X���ܴ���59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBirthday As String, strSQL As String
    Dim strCurDate As String
    Dim intAge As Integer
    Dim rsTemp As New ADODB.Recordset
    
    '��鲡�˵������ʽ�Ƿ���ȷ
    strSQL = "Select Zl_Age_Check([1]) From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge)
    strMsg = Trim(Nvl(rsTemp.Fields(0).Value))
    If strMsg <> "" Then Exit Function
    
    '������������������
    strBirthday = ""
    If strCalcDate = "" Then
        strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    Else
        strCurDate = strCalcDate
    End If
    
    If strAge Like "*��" Or strAge Like "*��*��" Then
        intAge = Mid(strAge, 1, InStr(1, strAge, "��") - 1)
        strBirthday = Format(DateAdd("yyyy", -1 * intAge, CDate(strCurDate)), "YYYY-MM-DD HH:mm")
        If Right(strAge, 1) = "��" Then
            intAge = Mid(strAge, InStr(1, strAge, "��") + 1, Len(strAge) - InStr(1, strAge, "��") - 1)
            strBirthday = Format(DateAdd("m", -1 * intAge, CDate(strBirthday)), "YYYY-MM-DD HH:mm")
        End If
        strBirthday = Format(strBirthday, "YYYY-MM-DD")
    ElseIf strAge Like "*��" Or strAge Like "*��*��" Then
        intAge = Mid(strAge, 1, InStr(1, strAge, "��") - 1)
        strBirthday = Format(DateAdd("m", -1 * intAge, CDate(strCurDate)), "YYYY-MM-DD HH:mm")
        If Right(strAge, 1) = "��" Then
            intAge = Mid(strAge, InStr(1, strAge, "��") + 1, Len(strAge) - InStr(1, strAge, "��") - 1)
            strBirthday = Format(DateAdd("d", -1 * intAge, CDate(strBirthday)), "YYYY-MM-DD HH:mm")
        End If
        strBirthday = Format(strBirthday, "YYYY-MM-DD")
    ElseIf strAge Like "*��" Or strAge Like "*��*Сʱ" Then
        intAge = Mid(strAge, 1, InStr(1, strAge, "��") - 1)
        strBirthday = Format(DateAdd("d", -1 * intAge, CDate(strCurDate)), "YYYY-MM-DD HH:mm")
        If Right(strAge, 2) = "Сʱ" Then
            intAge = Mid(strAge, InStr(1, strAge, "��") + 1, Len(strAge) - InStr(1, strAge, "��") - 2)
            strBirthday = Format(DateAdd("h", -1 * intAge, CDate(strBirthday)), "YYYY-MM-DD HH:mm")
            strBirthday = Format(strBirthday, "YYYY-MM-DD HH:mm")
        End If
    ElseIf strAge Like "*Сʱ" Or strAge Like "*Сʱ*����" Then
        intAge = Mid(strAge, 1, InStr(1, strAge, "Сʱ") - 1)
        strBirthday = Format(DateAdd("h", -1 * intAge, CDate(strCurDate)), "YYYY-MM-DD HH:mm")
        If Right(strAge, 2) = "����" Then
            intAge = Mid(strAge, InStr(1, strAge, "Сʱ") + 2, Len(strAge) - InStr(1, strAge, "Сʱ") - 3)
            strBirthday = Format(DateAdd("n", -1 * intAge, CDate(strBirthday)), "YYYY-MM-DD HH:mm")
        End If
        strBirthday = Format(strBirthday, "YYYY-MM-DD HH:mm")
    End If
    
    strDateOfBirth = strBirthday
    RecalcBirth = True
End Function

Public Function CheckOldData(ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox) As Boolean
'���ܣ������������ֵ����Ч��
'���أ�
    If Not IsNumeric(txt����.Text) Then CheckOldData = True: Exit Function
    
    Select Case cbo���䵥λ.Text
        Case "��"
            If Val(txt����.Text) > 200 Then
                MsgBox "���䲻�ܴ���200��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txt����.Text) > 2400 Then
                MsgBox "���䲻�ܴ���2400��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txt����.Text) > 73000 Then
                MsgBox "���䲻�ܴ���73000��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function

Public Function GetOldAcademic(ByVal DateBir As Date, ByVal str���䵥λ As String) As Long
'���ܣ����ݵ�ǰ�ĳ������ں����䵥λ�����������ϵ�����ֵ
'���أ�����
    Dim DatCur As Date, lngOld As Long, strInterval As String
    If DateBir = CDate(0) Or InStr(" ������", str���䵥λ) < 2 Then Exit Function
    
    DatCur = gobjDatabase.Currentdate
    
    strInterval = Switch(str���䵥λ = "��", "yyyy", str���䵥λ = "��", "m", str���䵥λ = "��", "d")
    lngOld = DateDiff(strInterval, DateBir, DatCur)
    If DateAdd(strInterval, lngOld, DateBir) > DatCur Then
        lngOld = lngOld - 1
    End If
    GetOldAcademic = lngOld
End Function

Public Function ReCalcOld(ByVal DateBir As Date, ByRef cbo���䵥λ As ComboBox, Optional ByVal lng����ID As Long, Optional ByVal blnSetControl As Boolean = True, _
    Optional ByVal datCalc As Date) As String
'����:���ݳ����������¼��㲡�˵�����,�������䵥λ
'����:blnSetControl�Ƿ��������䵥λ�ؼ�
'     datCalc-ָ����������,δָ��ʱ��ϵͳʱ�����
'����:����,���䵥λ
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    If datCalc = CDate(0) Then
        strSQL = "Select Zl_Age_Calc([1],[2],Null) old From Dual"
    Else
        strSQL = "Select Zl_Age_Calc([1],[2],[3]) old From Dual"
    End If
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, DateBir, datCalc)
    If blnSetControl = False Then
        ReCalcOld = Trim(Nvl(rsTmp!old))
        Exit Function
    End If
    
    If Not IsNull(rsTmp!old) Then
        If rsTmp!old Like "*��" Or rsTmp!old Like "*��" Or rsTmp!old Like "*��" Then
            strTmp = Mid(rsTmp!old, 1, Len(rsTmp!old) - 1)
            If IsNumeric(strTmp) Then
                Call gobjControl.cboLocate(cbo���䵥λ, Mid(rsTmp!old, Len(rsTmp!old), 1))
            Else
                strTmp = rsTmp!old
                cbo���䵥λ.ListIndex = -1
            End If
        Else
            strTmp = rsTmp!old
            If IsNumeric(strTmp) Then
                cbo���䵥λ.ListIndex = 0
            Else
                cbo���䵥λ.ListIndex = -1
            End If
        End If
    End If
    If cbo���䵥λ.ListIndex = -1 Then
        cbo���䵥λ.Visible = False
    Else
        If cbo���䵥λ.Visible = False Then cbo���䵥λ.Visible = True
    End If
    
    ReCalcOld = strTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckTextLength(strName As String, txtObj As TextBox) As Boolean
'����:��鲢��ʾ�ı������볤���Ƿ���
    CheckTextLength = True
    If gobjCommFun.ActualLen(txtObj.Text) > txtObj.MaxLength Then
        MsgBox strName & "���������ֻ�������� " & txtObj.MaxLength & " ���ַ��� " & txtObj.MaxLength \ 2 & " �����֡�", vbInformation, gstrSysName
        If txtObj.Enabled And txtObj.Visible Then txtObj.SetFocus
        CheckTextLength = False
    End If
End Function

Public Function CreateXWHIS(Optional ByVal blnMsg As Boolean) As Boolean
'���ܣ��ж� RIS�ӿڲ���(zl9XWInterface.clsHISInner) �Ƿ���ڣ�������
'������blnMsg������ʧ��ʱ�Ƿ���ʾ

    If Not gblnXW Then Exit Function
    If Not gobjXWHIS Is Nothing Then CreateXWHIS = True: Exit Function
    
    On Error Resume Next
    Set gobjXWHIS = GetObject(, "zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    On Error Resume Next
    If gobjXWHIS Is Nothing Then Set gobjXWHIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    If gobjXWHIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
        Exit Function
    End If

    CreateXWHIS = True
End Function

Public Function InitObjLis(Optional ByVal blnMsg As Boolean) As Boolean
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = GetObject(, "zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
    
        On Error Resume Next
        If gobjLIS Is Nothing Then Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
        
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                If blnMsg Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
                Exit Function
            End If
        End If
    End If
    InitObjLis = True
End Function

