Attribute VB_Name = "mdlCashBill"
Option Explicit

Public gcnOracle As New ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrProductName As String            '��Ʒ����
Public gstrMatchMethod As String
Public gstrDbUser As String                 '��ǰ���ݿ��û�
Private mrsPayMode As ADODB.Recordset
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    �������� As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

'Ʊ�ݿ���
Public gobjBillPrint As Object '������Ʊ�ݴ�ӡ����
Public gblnBillPrint As Boolean '������Ʊ�ݴ�ӡ�����Ƿ����

Public gstrSQL As String
Public gstr��λ���� As String
Public glngSys  As Long
Public glngModul As Long

Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
    g��������ģ�� = 5
    g����˽��ģ�� = 6
End Enum
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48
'ϵͳ��������----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private mlng���ű���ƽ������ As Long
Public gstrLike  As String

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function MshGetColNum(msh As MSHFlexGrid, strColName As String) As Long
'����:������������MSHFlexGrid�ؼ��е������,û���ҵ�ʱ����-1
'����:strColName-����
    Dim i As Long
    
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strColName Then MshGetColNum = i: Exit Function
    Next
    MshGetColNum = -1
End Function

Public Function GetNextId(ByVal strTable As String) As Long
    '-------------------------------------------------------------
    '���ܣ���ȡָ�����ΨһID��
    '������strTable
    '      ���ڵı���
    '���أ���ǰ���ΨһID��
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo ErrHand
    With rsTemp
        .Open "SELECT " & strTable & "_id.NextVal FROM DUAL", gcnOracle
        GetNextId = .Fields(0).Value
        .Close
    End With
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    GetNextId = Null
    Err = 0
End Function

Public Function GetPersonnelDept(ByVal lngID As Long) As ADODB.Recordset
'���ܣ���ȡָ����Ա�����в���
    Dim strSQL As String
 
    strSQL = "Select B.����,B.ID From ������Ա A, ���ű� B Where A.����id = B.ID And A.��Աid = [1] Order by ȱʡ Desc"
    On Error GoTo errH
    Set GetPersonnelDept = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDbUser
    UserInfo.���� = gstrDbUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function TruncateDate(ByVal datFull As Date) As Date
'ȥ�������е�ʱ���֡���
    TruncateDate = CDate(Format(datFull, "yyyy-MM-dd"))
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
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
Public Function ReturnMovedExes(ByVal strNO As String, Optional ByVal bytType As Byte = 2, Optional ByVal strFormCaption As String) As Boolean
'����:�����û�ѡ���ѡ�����ݱ��е����ݵ���ǰ���ݱ���
'����:bytType��ʾ��������,ֵ::1-�շ�,2-����,3-�Զ�����,4-�Һ�,5-���￨,6-Ԥ��,7-���ʣ�
'����:�û�ѡ��ȡ������,���߳�ѡ����ת��ʧ��,�򷵻�False
    
    MsgBox "��ǰ�����ĵ���" & strNO & "�ں����ݱ���!" & vbCrLf _
        & "����ϵͳ����Ա��ϵ,ת�뵽�������ݱ��ٲ���!", vbInformation, gstrSysName
    ReturnMovedExes = False
    
'�����ǳ�ѡ�������ݵĹ��̣��ݴ棬���ڽ���͸������ʱ����
'    If MsgBox("��ǰ��������" & strNO & "�ں����ݱ���,ϵͳ��Ҫ�Ȱ���˵�����ص�����ת�뵽�������ݱ���ܼ���!" & vbCrLf & _
'                             "ȷ��Ҫ���д˲�����?", vbInformation + vbYesNo, gstrSysName) = vbNo Then
'        ReturnMovedExes = False     '�˾��ʡ
'        Exit Function
'    End If
'
'    If zlDatabase.ReturnMovedExes(strNO, bytType, strFormCaption) Then
'        ReturnMovedExes = True
'    Else
'        '��ϸ������֮ǰ��ִ�й��̳���ʱ����
'        MsgBox "��ϵͳ����,��õ�����ص�����δ��ת�뵽�������ݱ�." & vbCrLf & "����δ�ɹ�,����ϵͳ����Ա��ϵ!", vbInformation, gstrSysName
'        ReturnMovedExes = False
'    End If
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
Public Sub zlSetCrlEnbled(ByVal objCrl As Object, blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ؼ���Nabled����,���ΪFalse,ͬʱ��Ҫ������صı���ɫ
    '���:objCrl-ת���ָ���ؼ�
    '     blnEnabled-�������
    '����:
    '����:
    '����:���˺�
    '����:2009-09-08 14:44:25
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Select Case UCase(TypeName(objCrl))
    Case UCase("TextBox"), UCase("COMBOBOX")
        objCrl.Enabled = blnEnabled
        zlSetCtrolBackColor objCrl
    Case UCase("dtpicker"), UCase("frame"), UCase("CHECKBOX"), UCase("LABEL"), UCase("COMMANDBUTTON")
        objCrl.Enabled = blnEnabled
    Case Else
       ' objCrl.Enabled = blnEnabled
    End Select
End Sub
Public Sub zlSetCtrolBackColor(ByVal objCtl As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����ɫ����ɫ
    '���:objCtl-ת��Ŀؼ�
    '����:
    '����:
    '����:���˺�
    '����:2009-09-08 14:43:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If objCtl.Enabled = False Then
        objCtl.BackColor = &H8000000F
    Else
        objCtl.BackColor = vbWhite
    End If
End Sub
Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '����:�������ƶ��ؼ���:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If objCtl.Visible And objCtl.Enabled = True Then: objCtl.SetFocus
End Sub
Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    Err = 0: On Error GoTo ErrHand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDbUser & "\" & App.ProductName, frmMain.Name, "����"
    zlSaveDockPanceToReg = True
ErrHand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(zlDatabase.GetPara("������������", , , True)) = 1
    Err = 0: On Error GoTo ErrHand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDbUser & "\" & App.ProductName, frmMain.Name, "����"
    zlRestoreDockPanceToReg = True
ErrHand:
End Function
Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub
Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '����:���ָ����Ȩ���Ƿ����
    '����:strPrivs-Ȩ�޴�
    '     strMyPriv-����Ȩ��
    '����,����Ȩ��,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function
Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub
Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ�������߶�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT
    Err = 0: On Error GoTo ErrHand:
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
ErrHand:
End Function


Public Function zlPersonSelect(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboSel As ComboBox, ByVal rsPerson As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Աѡ��ѡ����
    '���:cboSel-ָ���Ĳ���ѡ�񲿼�
    '     rsPerson-ָ������Ա��Ϣ(ID,���,����,����)
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str����-��������(������,���в���Ա��)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String, strLike As String
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsPerson)
    
    strSearch = UCase(strSearch)
        
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���� <> "" Then
        str���� = zlCommFun.SpellCode(str����)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!���) = strSearch Then lngID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!���)) = Val(strSearch) Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!���) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!���) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!���)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngID = 0
    If lngID <> 0 And rsTemp.RecordCount = 1 Then lngID = Nvl(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngID <> 0 Then GoTo GoOver:
    If lngID < 0 Then lngID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboSel, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngID = Val(Nvl(rsReturn!ID))
    If lngID < 0 Then lngID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboSel, lngID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlPersonSelect = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboSel
End Function


Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲿����Ϣ�Ƿ���ر���
    '����:��ʾ����,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 13:11:01
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    If mlng���ű���ƽ������ = 0 Then
        strSQL = "Select Avg(length(����)) As ���� From ���ű�"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡ���ű����ƽ������")
        mlng���ű���ƽ������ = Val(Nvl(rsTemp!����))
    End If
    '���ڱ��볤�ȿ��ܹ���,�޷���ʾ���ŵ�����,����Զ���ʾ�Ͳ���ʾ����,������5ʱ,����ʾ.С��5ʱ,��ʾ
   zlIsShowDeptCode = mlng���ű���ƽ������ <= 5
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���в��� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���:cboDept-ָ���Ĳ��Ų���
    '     rsDept-ָ���Ĳ���
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str���в���-���в�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
 
      
    
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���в��� <> "" Then
        str���� = zlCommFun.SpellCode(str���в���)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str���в���) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!����) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!����)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function

Public Function zlGetFeeFields(Optional strTableName As String = "������ü�¼", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ�����ֵ
    '��Σ�strTableName:��:������ü�¼;סԺ���ü�¼;....
    '      blnReadDatabase-�����ݿ��ж�ȡ
    '���Σ�
    '���أ��ֶμ�
    '���ƣ����˺�
    '���ڣ�2010-03-10 10:41:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strFileds As String
    
    Err = 0: On Error GoTo ErrHand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "������ü�¼"
        zlGetFeeFields = "" & _
        "Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, " & _
        "����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, " & _
        "�Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, " & _
        "����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, " & _
        "���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
        Exit Function
    Case "סԺ���ü�¼"
        zlGetFeeFields = "" & _
         " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, " & _
         " �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, " & _
         " ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, " & _
         " ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, " & _
         " ����id , ���ʽ��, ���մ���ID, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
         Exit Function
    Case "���˽��ʼ�¼"
        zlGetFeeFields = "Id, No, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, �շ�ʱ��, ��ʼ����, ��������, ��ע"
        Exit Function
    Case "����Ԥ����¼"
        zlGetFeeFields = "" & _
        " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, " & _
        " ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�"
        Exit Function
    Case "��Ա��"
        zlGetFeeFields = "" & _
        "Id, ���, ����, ����, ���֤��, ��������, �Ա�, ����, ��������, �칫�ҵ绰, �����ʼ�, ִҵ���, ִҵ��Χ, " & _
        "����ְ��, רҵ����ְ��, Ƹ�μ���ְ��, ѧ��, ��ѧרҵ, ��ѧʱ��, ��ѧ����, ������ѵ, ���п���, ���˼��, ����ʱ��, " & _
        "����ʱ��, ����ԭ��, ����, վ��"
        Exit Function
    Case "Ʊ�����ü�¼"
        zlGetFeeFields = "ID,Ʊ��,ʹ�����,������,ǰ׺�ı�,��ʼ����,��ֹ����,ʹ�÷�ʽ,�Ǽ�ʱ��,ʹ��ʱ��," & _
        "�Ǽ���,��ǰ����,ʣ������,����,�˶���,�˶�ʱ��,�˶Խ��,�˶�ģʽ,��ע,ǩ����,ǩ��ʱ��"
        Exit Function
    Case "Ʊ��ʹ����ϸ"
        zlGetFeeFields = "ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,���մ���,ʹ��ʱ��,ʹ����,�˶���,�˶�ʱ��,�˶Խ��,��ע"
        Exit Function
    Case "��Ա�ɿ��¼"
        zlGetFeeFields = "ID,����ID,�տ�Ա,�տ��ID,���㷽ʽ,�����,���,ժҪ,��ֹʱ��,�Ǽ�ʱ��,�Ǽ���"
        Exit Function
    
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo ErrHand:
    strSQL = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID;"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ϣ", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & Nvl(!column_name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
ErrHand:
  zlGetFeeFields = "*"
End Function

Public Function zlGetFullFieldsTable(Optional strTableName As String = "������ü�¼", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡһ�����ݱ��е��ֶ�.������Select Id,....
    '��Σ�bytHistory-0-��������ʷ����,1-��������ʷ����,2-����������( select * from tablename Union select * from Htablename)
    '      strWhere-����
    '      blnSubTable-�Ƿ��ӱ�
    '      strAliasName-����
    '���Σ�
    '���أ�select ID ... From tableName Union ALL
    '���ƣ����˺�
    '���ڣ�2010-03-10 11:19:11
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSQL As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '��
        strSQL = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '����ʷ
        strSQL = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '���߶�����
        strSQL = " Select  " & strFields & " From " & Trim(strTableName) & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSQL = " (" & strSQL & ") " & strAliasName
    zlGetFullFieldsTable = strSQL
End Function



Public Function Select��Աѡ����(ByVal frmMain As Form, ByVal objCtl As Object, _
    ByVal strKey As String, Optional lng����ID As Long = 0, _
    Optional lng��ԱID As Long = 0, _
    Optional bln��������Ա��ʾ As Boolean = False, _
    Optional strSearchKey As String = "", _
    Optional str��Ա���� As String = "", _
    Optional str����ְ�� As String = "", _
    Optional strרҵ����ְ�� As String = "", _
    Optional strTittle As String = "��Աѡ����", _
    Optional strNote As String = "��ѡ����ص���Ա", _
    Optional strNotFindMsg As String = "δ�ҵ�ָ������Ա,����!", _
    Optional strShowField As String = "����", _
    Optional strShowSplit As String = "-") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ������Ա
    '���:frmMain-���õĸ�����
    '     objCtl-�ؼ�(Ŀǰֻ֧���ı���)
    '     strKey-����Ľ�ֵ
    '     lng����ID-�����Ϊ��,��������Ա,����, ��ָ�������µ���Ա
    '     str��Ա����: ��ҽ��,ҽ��1... ��ʽ
    '     str����ְ��strרҵ����ְ��: ��ְ��1,ְ��21... ��ʽ
    '����:lng��Աid-������ԱID
    '����:���ҳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/08/23
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, bytType As Byte, str��Ա����Table As String, strWhere As String
    Dim blnCancel As Boolean, sngX As Single, sngY As Single, lngH As Long, i As Long
    Dim vRect As RECT
    
    'zlDatabase.ShowSQLSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmMain=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    Err = 0: On Error GoTo ErrHand:
    bytType = 0: strWhere = ""
    If str��Ա���� <> "" Then
        str��Ա����Table = ",��Ա����˵�� Q1,(Select Column_Value From Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) Q2" & vbCrLf
        strWhere = strWhere & " And ( A.ID=Q1.��ԱID and Q1.��Ա���� = Q2.Column_Value ) " & vbCrLf
    End If
    If str����ְ�� <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)))  Where a.����ְ��=Column_Value) " & vbCrLf
    If strרҵ����ְ�� <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))  Where a.רҵ����ְ��=Column_Value) " & vbCrLf
    
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey, False)
        If lng����ID = 0 Then
            gstrSQL = "" & _
                "   Select /*+ rule */ distinct A.ID,A.���,A.����,A.����,A.����,A.�Ա�,A.����,A.��������,A.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                "   From ��Ա�� A " & str��Ա����Table & _
                "   Where (A.���� like [1] or A.��� like [1] or A.���� like Upper([1]) or A.���� like [1]) " & strWhere & zl_��ȡվ������(True, "A") & "" & _
                "       and (A.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                "   order by A.���"
        Else
            gstrSQL = "" & _
                "   Select /*+ rule */ distinct a.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                "   From ��Ա�� a,������Ա C " & str��Ա����Table & _
                "   Where a.id=c.��Աid and c.����Id=[2]   " & strWhere & zl_��ȡվ������(True, "a") & _
                "       and (a.���� like [1] or a.��� like [1] or a.���� like Upper([1]) or a.���� like [1]) " & _
                "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & _
                "   order by ���"
        End If
     Else
        If lng����ID = 0 Then
            If bln��������Ա��ʾ Then
                gstrSQL = "" & _
                "   Select /*+ rule */  id," & IIf(gstrNodeNo <> "-", "1 as ����ID,-1*NULL as �ϼ�ID", "Level as ����ID,�ϼ�id") & " ,����,����,0 ĩ��,'' as ����,'' as ����,''as �Ա�,''as ����, to_date(Null,'yyyy-mm-dd')  as ��������, '' as  �칫�ҵ绰 ,'' ִҵ���, '' ����ְ��,'' רҵ����ְ��" & _
                "   From ���ű� " & _
                "   where ����ʱ�� is null or ����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') " & zl_��ȡվ������() & _
                    IIf(gstrNodeNo <> "-", "", "   Start with �ϼ�id is null connect by prior id=�ϼ�id ") & _
                "   union all " & _
                "   Select  distinct a.ID,999999 AS ����ID,b.����id as �ϼ�ID,a.���,a.����,1 as ĩ��,����,����,�Ա�,����,��������,�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ�� " & _
                "   From ��Ա�� a,������Ա b  " & str��Ա����Table & _
                "   Where a.id=b.��Աid and b.ȱʡ=1  " & strWhere & zl_��ȡվ������(True, "a") & _
                "         And (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
                "   Order by ����ID,����"
                bytType = 2
            Else
                gstrSQL = "" & _
                    "   Select  /*+ rule */  distinct A.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                    "   From ��Ա�� A " & str��Ա����Table & _
                    "   Where (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & strWhere & zl_��ȡվ������(True, "a") & _
                    "   order by a.���"
            End If
        Else
            gstrSQL = "" & _
                "   Select /*+ rule */ distinct a.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                "   From ��Ա�� a,������Ա C " & str��Ա����Table & _
                "   Where a.id=c.��Աid and c.����Id=[2] " & _
                "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)  " & strWhere & zl_��ȡվ������(True, "a") & _
                "   order by a.���"
        End If
    End If
   
   
   '���궨λ
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, objCtl.MsfObj)
        lngH = objCtl.MsfObj.CellHeight
    Case Else
        vRect = GetControlRect(objCtl.hWnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = objCtl.Height
    End Select
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, bytType, strTittle, bytType = 2, strSearchKey, strNote, bytType = 2, False, Not (bytType = 2), sngX, sngY, lngH, blnCancel, False, False, strKey, lng����ID, str��Ա����, str����ְ��, strרҵ����ְ��)
    
    lng��ԱID = 0
    If blnCancel = True Then
        Call zl_CtlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        If strNotFindMsg <> "" Then ShowMsgbox strNotFindMsg
        Call zl_CtlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    Call zl_CtlSetFocus(objCtl, True)
    If bytType = 2 Then
        strShowField = "," & strShowField & ",M_��,"
        strShowField = Replace(strShowField, ",���,", ",����,")
        strShowField = Replace(strShowField, ",����,", ",����,")
        strShowField = Mid(strShowField, 2)
        strShowField = Replace(strShowField, ",M_��,", "")
    End If
    
    '������ص�ֵ
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .EditText = .TextMatrix(.Row, .Col)
            .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!ID)
        End With
    Case UCase("BILLEDIT")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .Text = .TextMatrix(.Row, .Col)
        End With
    Case UCase("ComboBox")
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!ID) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            ShowMsgbox "��ѡ��Ĳ����������б��в�����,����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        objCtl.Text = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
        objCtl.Tag = Val(rsTemp!ID)
        zlCommFun.PressKey vbKeyTab
    End Select
    lng��ԱID = Val(Nvl(rsTemp!ID))
    rsTemp.Close
    Select��Աѡ���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlCheckPrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ����Ȩ���Ƿ����
    '����:strPrivs-Ȩ�޴�
    '     strMyPriv-����Ȩ��
    '����,����Ȩ��,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-19 14:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlCheckPrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function

Public Function zl_��ȡվ������(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str���� As String = "") As String
    '����:��ȡվ����������:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str���� = "", "", str���� & ".") & "վ��"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_��ȡվ������ = strWhere
End Function
Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        x = objPoint.x * 15 'objBill.Left +
        y = objPoint.y * 15 + objBill.Height '+ objBill.Top
    Else
        x = objPoint.x * 15 + objBill.CellLeft
        y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Sub zl_CtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '����:�������ƶ��ؼ���:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
End Sub

Public Function zl_GetFieldValue(ByVal rsTemp As ADODB.Recordset, _
    Optional ByVal strShowFields As String = "����,����", _
    Optional ByVal strShowSplit As String = "-") As String
    '-----------------------------------------------------------------------------------------------------------
    '����:������ʾ�ֶε����ֵ
    '���:rsTemp-��¼��
    '     strShowFields-��ʾ���ֶ�
    '     strShowSplit-��ʾ�ķ����
    '����:
    '����:�ɹ�,������ص��ֶ�ֵ
    '����:���˺�
    '����:2009-03-06 11:59:19
    '-----------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, strValue As String, strLeft As String, strRight As String
    varData = Split(strShowFields, ",")
    
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    If rsTemp.RecordCount = 0 Then Exit Function
    
    Select Case strShowSplit
    Case "[", "[]", "]"
        strLeft = "[": strRight = "]"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "�ۣ�", "��", "��"
        strLeft = "��": strRight = "��"
    Case "[]", "[", "]"
        strLeft = "[": strRight = "]"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "{}", "{", "}"
        strLeft = "{": strRight = "}"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case Else
        strLeft = "": strRight = strShowSplit
    End Select
    
    strValue = ""
    With rsTemp
        For i = 0 To UBound(varData) - 1
            strValue = strValue & strLeft & Nvl(.Fields(varData(i))) & strRight
        Next
        strValue = strValue & Nvl(.Fields(varData(UBound(varData))))
    End With
    zl_GetFieldValue = strValue
End Function
Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '����:�жϿؼ��Ƿ��
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo ErrHand:
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
'*********************************************************************************************************************
Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub
Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '��Ҫ�пո������
        strTmp = Substr(strCode, 1, lngLen)
    End If
    'ȡ��������ַ�
    Rpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function zlIsOnlyNum(ByVal strAsk As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ���ַ����Ƿ�ȫ�������ֹ���
    '���:strAsk-��Ҫ�жϵ��ַ�
    '����:
    '����:���ȫ�����ֹ��ɣ�����true,���򷵻�False
    '����:���˺�
    '����:2010-11-17 11:19:15
    '˵��:
    '     isnumberic���ܼ����Щ:-099.22,22d2��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            strTemp = Mid(Trim(strAsk), i, 1)
            If InStr("0123456789", strTemp) = 0 Then Exit Function
        Next
        zlIsOnlyNum = True
    End If
End Function

Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '--�����:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '--������:
    '--��  ��:�Ӵ�
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    Err = 0
    On Error GoTo ErrHand:
    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
    Exit Function
ErrHand:
    Substr = ""
End Function
Public Function zlAddNum(ByVal strVal As String, Optional blnAdd As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݼ��ַ�
    '���:strVal=Ҫ��1���ַ���
    '����:blnAdd-true:����;false;�ݼ�
    '����:������ݼ�����ַ�
    '����:���˺�
    '����:2010-11-18 15:19:58
    '˵��:ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTmp As String, intUp As Integer, intAdd As Integer
    Dim strCur As String
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            intAdd = 1
        Else
            intAdd = 0
        End If
        strCur = Mid(strVal, i, 1)
        If IsNumeric(strCur) Then
            If blnAdd Then
                If CByte(Mid(strVal, i, 1)) + intAdd + intUp < 10 Then
                    strVal = Left(strVal, i - 1) & CByte(strCur) + intAdd + intUp & Mid(strVal, i + 1)
                    intUp = 0
                Else
                    strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                    intUp = 1
                End If
            Else
                If CByte(strCur) - intAdd - intUp < 0 Then
                    strVal = Left(strVal, i - 1) & "9" & Mid(strVal, i + 1)
                    intUp = 1
                Else
                    strVal = Left(strVal, i - 1) & CByte(strCur) - intAdd - intUp & Mid(strVal, i + 1)
                    intUp = 0
                End If
            End If
        Else
            If blnAdd Then
                If Asc(strCur) + intAdd + intUp <= Asc("Z") Then
                    strVal = Left(strVal, i - 1) & Chr(Asc(strCur) + intAdd + intUp) & Mid(strVal, i + 1)
                    intUp = 0
                Else
                    strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                    intUp = 1
                End If
            Else
                If Asc(strCur) - intAdd - intUp < Asc("A") Then
                    strVal = Left(strVal, i - 1) & "Z" & Mid(strVal, i + 1)
                    intUp = 1
                Else
                    strVal = Left(strVal, i - 1) & Chr(Asc(strCur) - intAdd - intUp) & Mid(strVal, i + 1)
                    intUp = 0
                End If
            End If
        End If
        If intUp = 0 Then Exit For
    Next
    zlAddNum = strVal
End Function

Public Function NumberSubtrac(str������ As String, str���� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ַ��������1
    '���:strOne ������: strTwo ����
    '����:������ַ���
    '����:����
    '����:2012-08-27 10:00:00
    '�����:43366
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int������ As Integer  '��������ǰλ���ϵ�ֵ������123ʮλ�ϵ�ֻΪ2
    Dim int���� As Integer  '������ǰλ���ϵ�ֵ
    Dim intLen������ As Integer '����������λ�� ������123����λ��Ϊ3λ
    Dim intLen���� As Integer '��������λ��
    Dim Index As Integer '��ǰ���㵽��λ��λ��
    Dim bln�Ƿ�ʮ As Boolean '�ж����������ʱ,�Ƿ񱻼���С�ڼ���,��������Ҫ����һλ�ϼ�ȥ1
    Dim bln�Ƿ��ʮ As Boolean '�ж����������ʱ,�Ƿ񳬹���10����������Ҫ����һλ�ϼ�1������129�� 9+1=10 ����Ҫ��2��+1
    Dim bln�Ƿ�Ϊ���� As Boolean
    Dim strCurr������ As String
    Dim strCurr���� As String
    Dim strʣ��λ��ֵ As String
    Dim str��ֵ As String
    Dim strTemp1 As String
    Dim strTemp2 As String
    Dim i As Integer
    
    'ȥ���������ͼ�����ߵ�0�ַ�
    While Left(str������, 1) = "0"
        str������ = Mid(str������, 2, Len(str������) - 1)
    Wend
    If str������ = "" Then str������ = "0"
    
    While Left(str����, 1) = "0"
        str���� = Mid(str����, 2, Len(str����) - 1)
    Wend
    If str���� = "" Then str���� = "0"
    
    bln�Ƿ��ʮ = True
    '��������1(��Ҫ�����Ʊ�ݺ���Ӧ�ð�����ʼ�ĺ���)
    While bln�Ƿ��ʮ
        Index = Index + 1
        If Index <= Len(str������) Then
            int������ = CInt(Left(Right(str������, Index), 1))
            If int������ + 1 = 10 Then
                bln�Ƿ��ʮ = True
                int������ = 0
            Else
                bln�Ƿ��ʮ = False
                int������ = int������ + 1
            End If
            
            strTemp2 = Left(str������, Len(str������) - Index)
            strTemp1 = Right(str������, Index)
            strTemp1 = int������ & Mid(strTemp1, 2, Len(strTemp1) - 1)
            str������ = strTemp2 & strTemp1
        Else
            str������ = "1" & str������
            bln�Ƿ��ʮ = False
        End If
    Wend
    
    Index = 0
    
    intLen������ = Len(str������)
    intLen���� = Len(str����)
    
    If intLen������ > intLen���� Then
        strCurr������ = str������
        strCurr���� = str����
        bln�Ƿ�Ϊ���� = False
    ElseIf intLen������ < intLen���� Then
        strCurr������ = str����
        strCurr���� = str������
        bln�Ƿ�Ϊ���� = True
    ElseIf str������ > str���� Then
        strCurr������ = str������
        strCurr���� = str����
        bln�Ƿ�Ϊ���� = False
    ElseIf str������ < str���� Then
        strCurr������ = str����
        strCurr���� = str������
        bln�Ƿ�Ϊ���� = True
    ElseIf str������ = str���� Then
        strCurr������ = str������
        strCurr���� = str����
        bln�Ƿ�Ϊ���� = False
    End If
    
    '��λ�����ֵ��ѭ��ȡÿλ�ϵ�ֵ
    For i = 1 To IIf(intLen������ <= intLen����, intLen������, intLen����)
          int������ = CInt(Left(Right(strCurr������, i), 1))
          int���� = CInt(Left(Right(strCurr����, i), 1))
          int������ = int������ - IIf(bln�Ƿ�ʮ, 1, 0)
          bln�Ƿ�ʮ = False
          If int������ < int���� Then
                int������ = int������ + 10
                bln�Ƿ�ʮ = True
          End If
          str��ֵ = (int������ - int����) & str��ֵ
          Index = i
    Next
    
    While bln�Ƿ�ʮ
        Index = Index + 1
        If bln�Ƿ�ʮ = True Then
            int������ = CInt(Left(Right(strCurr������, Index), 1))
            If int������ < 1 Then
                int������ = int������ + 9
                bln�Ƿ�ʮ = True
            Else
                int������ = int������ - 1
                bln�Ƿ�ʮ = False
            End If
        End If
        str��ֵ = int������ & str��ֵ
    Wend
    
    While Left(str��ֵ, 1) = "0" And intLen������ = intLen����
       str��ֵ = Mid(str��ֵ, 2, Len(str��ֵ) - 1)
    Wend
    If Index <= IIf(intLen������ >= intLen����, intLen������, intLen����) Then
        str��ֵ = Left(strCurr������, IIf(intLen������ >= intLen����, intLen������, intLen����) - Index) & str��ֵ
    End If
    
    While Left(str��ֵ, 1) = "0"
       str��ֵ = Mid(str��ֵ, 2, Len(str��ֵ) - 1)
    Wend
    
    If bln�Ƿ�Ϊ���� Then str��ֵ = "-" & str��ֵ
    
    NumberSubtrac = IIf(str��ֵ = "", 0, str��ֵ)
    
End Function

Public Function NumberSum(strOne As String, strTwo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ַ�����ͣ���ʱ������������и��������
    '���:strOne : strTwo
    '����:��ͺ���ַ���
    '����:���ϴ�
    '����:2014/9/2 17:28:17
    '�����:77390
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intOne As Integer  '�������ǰλ���ϵ�ֵ������123ʮλ�ϵ�ֻΪ2
    Dim intTwo As Integer
    Dim Index As Integer '��ǰ���㵽��λ��λ��
    Dim bln�Ƿ��ʮ As Boolean '�ж����������ʱ,�Ƿ񳬹���10����������Ҫ����һλ�ϼ�1������129�� 9+1=10 ����Ҫ��2��+1
    Dim str��ֵ As String
    Dim i As Integer
    Err = 0: On Error GoTo errHandle
    'ȥ���������ߵ�0�ַ�
    While Left(strOne, 1) = "0"
        strOne = Mid(strOne, 2, Len(strOne) - 1)
    Wend
    If strOne = "" Then strOne = "0"
    
    While Left(strTwo, 1) = "0"
        strTwo = Mid(strTwo, 2, Len(strTwo) - 1)
    Wend
    If strTwo = "" Then strTwo = "0"

    '��λ�����ֵ��ѭ��ȡÿλ�ϵ�ֵ
    For i = 1 To IIf(Len(strOne) <= Len(strTwo), Len(strTwo), Len(strOne))
        intOne = IIf(i > Len(intOne), 0, CInt(Left(Right(strOne, i), 1)))
        intTwo = IIf(i > Len(intTwo), 0, CInt(Left(Right(strTwo, i), 1)))
        intOne = intOne + IIf(bln�Ƿ��ʮ, 1, 0)
        bln�Ƿ��ʮ = False
        If intOne + intTwo >= 10 Then
            intOne = intOne - 10
            bln�Ƿ��ʮ = True
        End If
        str��ֵ = (intOne + intTwo) & str��ֵ
        Index = i
    Next
    '���λ��ͺ���Ҫ��һλ�����
    If bln�Ƿ��ʮ Then str��ֵ = "1" & str��ֵ
    NumberSum = str��ֵ
    Exit Function
errHandle:
    NumberSum = "0"
End Function

Public Function Get���㷽ʽ() As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㷽ʽ
    '����:���㷽ʽ��
    '����:���˺�
    '����:2013-09-04 17:22:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "" & _
    "   Select ����,����,����,nvl(Ӧ�տ�,0) as Ӧ�տ�,nvl(Ӧ����,0) as Ӧ����," & _
    "               nvl(ȱʡ��־,0) as ȱʡ��־  " & _
    "   From ���㷽ʽ"
    If mrsPayMode Is Nothing Then
        Set mrsPayMode = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���㷽ʽ")
    ElseIf mrsPayMode.State <> 1 Then
        Set mrsPayMode = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���㷽ʽ")
    End If
    Set Get���㷽ʽ = mrsPayMode
End Function
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNoBeginTrans:û������ʼ
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub



Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
    '���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
    '������intNum=��Ŀ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim dtCurDate As Date, strMaxNo As String
    Dim strYearStr As String
    Err = 0: On Error GoTo errH:
    
    strSQL = "Select ��Ź���,Sysdate as ����,������ From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������", intNum)
    If rsTmp.EOF Then GetFullNO = strNO: Exit Function
    Select Case Val(Nvl(rsTmp!��Ź���))
    Case 0, 1 '0-����˳����,1-����˳����
        If Len(strNO) >= 8 Then
            GetFullNO = Right(strNO, 8)
            Exit Function
        ElseIf Len(strNO) = 7 Then
            GetFullNO = PreFixNO & strNO
            Exit Function
        End If
        GetFullNO = strNO
        dtCurDate = Date
        If Not rsTmp.EOF Then
            intType = Val("" & rsTmp!��Ź���)
            dtCurDate = rsTmp!����
            strMaxNo = Nvl(rsTmp!������)
        End If
        strYearStr = PreFixNO
        If strMaxNo = "" Then strMaxNo = strYearStr & "000001"
        If intType = 1 Then
            '���ձ��
            strSQL = Format(CDate(Format(dtCurDate, "YYYY-MM-dd")) - CDate(Format(dtCurDate, "YYYY") & "-01-01") + 1, "000")
            GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
            Exit Function
        End If
        '������
        If Len(strNO) = 6 Then
            GetFullNO = Left(strMaxNo, 2) & strNO: Exit Function
        End If
        GetFullNO = Left(strMaxNo, 2) & zlLeftPad(Right(strNO, 6), 6, "0")
    Case 2  '2-�����ҷ��»���ձ����Ҫ��ȡ���Һ����,
    Case 3   '3-��������+˳���(��ȡ��λ,˳���ȡ4λ)
        If Len(strNO) <= 6 Then
            GetFullNO = Format(rsTmp!����, "YYMMDD") & zlLeftPad(strNO, 6, "0")
            Exit Function
        End If
        If Len(strNO) <= 8 Then
            GetFullNO = Format(rsTmp!����, "YYMM") & zlLeftPad(strNO, 8, "0")
            Exit Function
        End If
        If Len(strNO) <= 10 Then
            GetFullNO = Format(rsTmp!����, "YY") & zlLeftPad(strNO, 10, "0")
            Exit Function
        End If
        If Len(strNO) <= 12 Then
            GetFullNO = zlLeftPad(strNO, 12, "0")
            Exit Function
        End If
    Case 4    '4-��ִ�п��ҷ��ڼ���(��(�ڼ���е���)+ִ�п��ұ��+�·�(�ڼ���е���)+˳���)
    Case 5    '5-�����½��б��(yyyyMM000000)
    Case Else
    End Select
    GetFullNO = strNO
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function zlLeftPad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ���������ƿո�
    '����:�����ִ�
    '����:���˺�
    '����:2012-02-22 17:58:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = zlSubstr(strCode, 1, lngLen)
    End If
    zlLeftPad = Replace(strTmp, Chr(0), strChar)
End Function
Private Function zlSubstr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '���:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '����:�Ӵ�
    '����:���˺�
    '����:2012-02-22 18:00:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo ErrHand:
    zlSubstr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    zlSubstr = Replace(zlSubstr, Chr(0), " ")
    Exit Function
ErrHand:
    zlSubstr = ""
End Function
Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
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
    '���ܣ�����Ŀ���ݲ���ComboBox������ֵ
    '������lngID=ComboBox����Ŀֵ
    Dim i As Long
    If lngID = 0 Then FindCboIndex = -1: Exit Function
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = lngID Then
            FindCboIndex = i
            Exit Function
        End If
    Next
    FindCboIndex = -1
End Function
Public Function Get���ʽ�������(ByVal intRollingType As Integer, _
    ByRef strOut�������� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ʵĽ�������
    '���:intRollingType-�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�)
    '����:strOut��������-���ر��εĽ�������,����ö��ŷָ�,����:,2,...
    '     �������������Ԥ�������ѿ�,�򷵻ؿ�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-03-05 15:04:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strOut�������� = ""
    On Error GoTo errHandle
    '�������,Ԥ����,ֱ�ӷ���
    If intRollingType = 0 Or intRollingType = 2 Or intRollingType = 21 Or intRollingType = 22 _
        Or intRollingType = 6 Then Get���ʽ������� = True: Exit Function
 
    'Ԥ������NULL,2-����,3-�շ�,4-�Һ�,5-���￨,6-����ҽ����
    'intRollingType:1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�)
    If intRollingType = 1 Then  '�շ�
        strOut�������� = ",3,6,": Get���ʽ������� = True: Exit Function
    End If
    If intRollingType = 3 Then  '����
        strOut�������� = ",2,": Get���ʽ������� = True: Exit Function
    End If
    If intRollingType = 4 Then  '�Һ�
        strOut�������� = ",4,": Get���ʽ������� = True: Exit Function
    End If
    If intRollingType = 5 Then  '���￨
        strOut�������� = ",5,": Get���ʽ������� = True: Exit Function
    End If
    Get���ʽ������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


