Attribute VB_Name = "mdlAppTool"
Option Explicit
Public Const HKEY_CURRENT_USER = &H80000001
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SWP_SHOWWINDOW = &H40
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SMTO_ABORTIFHUNG = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWNOACTIVATE = 4

Public Type ChooseColorType
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As String
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColorType) As Long


Public Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SetActiveWindow Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "User32" (ByVal hWndLock As Long) As Long
Public Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'????????????????????????????????????????Layout,????????????????
Public Declare Function GetKeyboardLayoutList Lib "User32" (ByVal nBuff As Long, lpList As Long) As Long
'????????????????????
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'????????????????????????????
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

Public gcnOracle As New ADODB.Connection        '????????????????????????????????????????????
Public gclsAppTool As clsAppTool       '????APPTool????
Public gstrPrivs As String                   '????????????????????????????

Public gstrSysName As String                '????????
Public gstrVersion As String                '????????
Public gstrAviPath As String                'AVI??????????????

Public gstrDbUser As String                 '??????????????
Public glngUserId As Long                   '????????id
Public gstrUserCode As String               '????????????
Public gstrUserName As String               '????????????
Public gstrUserAbbr As String               '????????????

Public glngDeptId As Long                   '????????????id
Public gstrDeptCode As String               '????????????????
Public gstrDeptName As String               '????????????????

Public gstr???????? As String
Public gstrSQL As String
Public gstrMenuSys As String                '??????????????????????
Public glngSys As Long                      '????????

'??????????????????????????????
Public gfrmMain As Object                   '??????????????????????????????????????????
Public gblnMessageShow As Boolean           '??????????????????????????
Public gblnMessageGet  As Boolean           '????????????????????????????

Public Const glngLBound As Long = 99
Public Const glngUBound As Long = 240

Public Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Sub GetUserInfo()
'????:??????????????

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    rsTemp.CursorLocation = adUseClient
    On Error GoTo errHand
    
    With rsTemp
        strSQL = "select P.*,D.???? as ????????,D.???? as ????????,M.????ID" & _
                " from ?????????? U,?????? P,?????? D,???????? M " & _
                " Where U.????id = P.id And P.ID=M.????ID and  M.????=1 and M.????id = D.id and (P.???????? = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.???????? Is Null) And U.??????=user"
        .Open strSQL, gcnOracle, adOpenKeyset
                
        If .RecordCount <> 0 Then
            glngUserId = .Fields("ID").Value                '????????id
            gstrUserCode = .Fields("????").Value            '????????????
            gstrUserName = .Fields("????").Value            '????????????
            gstrUserAbbr = IIf(IsNull(.Fields("????").Value), "", .Fields("????").Value)          '????????????
            glngDeptId = .Fields("????id").Value            '????????????id
            gstrDeptCode = .Fields("????????").Value        '????????
            gstrDeptName = .Fields("????????").Value        '????????
        Else
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
        .Close
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Err = 0
End Sub

'Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strFormCaption As String)
''????????????????????????SQL????
'    If rsTemp.State = adStateOpen Then rsTemp.Close
'
'    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
'    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'    Call SQLTest
'End Sub

Private Function SystemImes() As Variant
'????????????????????????????????????????????????
'??????????????????????????,??????????
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, StrName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer

    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then '??1??????????????
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), StrName, Len(StrName))
            arrName(j) = Mid(StrName, 1, InStr(1, StrName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function ChooseIME(cmbIME As Object) As Boolean
    Dim varIME As Variant
    Dim i As Integer
    Dim strIme As String
    
    varIME = SystemImes
    If Not IsArray(varIME) Then
        MsgBox "??????????????????????????????????????????" & vbCrLf & _
               "????????????????????????????????", vbInformation, gstrSysName
        Exit Function
    End If
    cmbIME.Clear
    cmbIME.AddItem "??????????"
    strIme = zlDatabase.GetPara("??????")
    For i = LBound(varIME) To UBound(varIME)
        cmbIME.AddItem varIME(i)
        If strIme = varIME(i) Then cmbIME.ListIndex = i + 1
    Next
    If cmbIME.ListIndex < 0 Then cmbIME.ListIndex = 0
    ChooseIME = True
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'????????????????????????????????(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function NewClientRecord(ByVal strFilds As String) As ADODB.Recordset
    '??????????????????
    'strFilds:??????,????,????;??????,????,????...
    '    ????:??????,varchar2,30;????,varchar2,30
    
    Dim rs As ADODB.Recordset, i As Integer
    Dim varFilds As Variant
    Dim varFild As Variant
    Dim strTmp As String
    Set rs = New ADODB.Recordset
    
    varFilds = Split(strFilds, ";")
    With rs
        For i = LBound(varFilds) To UBound(varFilds)
            strTmp = varFilds(i)
            varFild = Split(strTmp, ",")
            
            If UCase(varFild(1)) = "VARCHAR2" Then
                .Fields.Append varFild(0), adVarWChar, CLng(varFild(2)), adFldIsNullable
            ElseIf UCase(varFild(1)) = "NUMBER" Then
                .Fields.Append varFild(0), adVarNumeric, CLng(varFild(2)), adFldIsNullable
            End If
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set NewClientRecord = rs
End Function

Public Function IsCheckConstraint(ByVal strOwner As String, ByVal strTableName As String, ByVal strColumnName As String, ByVal bytType As Byte) As Boolean
'????Check????????
'bytType
'  1: ?????? Check In (0,1) ????
'  2: ?????? Check Is Not Null ????
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo ErrH
    strTmp = "Select A.Search_Condition from All_Constraints A, All_Cons_Columns B " _
           & "Where A.Constraint_Name = B.Constraint_Name and A.owner=[1] and a.Table_Name=[2] and B.Column_Name=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "", strOwner, strTableName, strColumnName)
    If Not rsTmp.EOF And IsNull(rsTmp!search_condition) = False Then
        Select Case bytType
            Case 1: If InStr(rsTmp!search_condition, "(0,1)") Or InStr(rsTmp!search_condition, "(1,0)") Then IsCheckConstraint = True
            Case 2: If InStr(UCase(rsTmp!search_condition), "IS NOT NULL") Or InStr(UCase(rsTmp!search_condition), "IS NULL") And InStr(UCase(rsTmp!search_condition), "NOT") Then IsCheckConstraint = True
        End Select
    End If
    rsTmp.Close
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function IsPathProperty(strOwner As String, strTable As String) As String
'??????????????????????????????????????
'??????????????????;????????;????????
    Dim i As Integer
    Dim bln???? As Boolean, blnID As Boolean, bln???? As Boolean, bln???? As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    IsPathProperty = ";"
    On Error GoTo errHandle
    
    Set rsTmp = zlDatabase.OpenSQLRecord("select * from " & strOwner & "." & strTable & " where rownum=0", "")
    If rsTmp Is Nothing Then Exit Function
    
    For i = 0 To rsTmp.Fields.Count - 1
        If rsTmp.Fields(i).Name = "????" Then
            bln???? = True
        ElseIf rsTmp.Fields(i).Name = "ID" Then
            blnID = True
        ElseIf rsTmp.Fields(i).Name = "????" Then
            bln???? = True
        End If
    Next
    rsTmp.Close
    If ((blnID Or bln????) And bln????) = False Then Exit Function
    
    strTmp = "Select b.Column_Name, c.Column_Name r_column_name,c.TABLE_NAME r_table_name " _
           & "From All_Constraints A, All_Cons_Columns B, All_Cons_Columns C " _
           & "Where a.Constraint_Name = b.Constraint_Name And a.r_Constraint_Name = c.Constraint_Name And a.Constraint_Type = 'R' " _
           & "  And a.owner=[1] and a.table_name=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "??????????????????????", strOwner, strTable)
    Do While rsTmp.EOF = False
        If UCase(zlCommFun.Nvl(rsTmp!column_name)) = "????ID" And UCase(zlCommFun.Nvl(rsTmp!r_table_name)) = "RESOURCEINFO" Then
            '??????????BH??????????????
            IsPathProperty = ";;RESOURCEINFO"
        Else
            IsPathProperty = zlCommFun.Nvl(rsTmp!column_name) & ";" & zlCommFun.Nvl(rsTmp!r_column_name) & ";" & zlCommFun.Nvl(rsTmp!r_table_name)
            Exit Do
        End If
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'??????????Oracle??Decode????
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
