Attribute VB_Name = "mdlPublic"
Option Explicit 'Ҫ���������
'ϵͳ������ʱ����
Public gfrmMain As Object
Public glngMain As Long
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrSQL As String

Public glngSys As Long
Public glngModul As Long
Public gblnShowInTaskBar As Boolean

Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public gstrProductName As String

Public gstrUnitName As String '�û���λ����
Public gstrDBUser As String '��ǰ���ݿ��û���

Public gstrIme As String '�Զ��Ŀ������뷨
Public gblnOK As Boolean
Public Const LONG_MAX = 2147483647 'Long�����ֵ
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'�����������------------------------------------
Public Const conLineWide As Integer = 30 '������ռ���(��λΪ�)ռ�����߿��
Public Const conLineHigh As Integer = 30 '������ռ�߶�(��λΪ�)ռ�����߸߶�
Public gobjOutTo As Object

Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    pסԺ���� = 1133
    p���˽��� = 1137
    p���ò�ѯ = 1139
    pһ���嵥 = 1141
    p���ʲ��� = 1150
    pסԺҽ���´� = 1253
    pԤ���� = 1103
End Enum

'API��Ϣ-----------------------------------------
Public lngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
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
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type POINTAPI
     x As Long
     y As Long
End Type

Public Enum Em_Appearance
    Show_3D = 1     '3D��ʾ
    Show_Flat = 0   'ƽ��
End Enum

Public Enum Em_BorderStyle
    Show_Fixed_Single = 1
    Show_None = 0   '�ޱ߿���
End Enum

'�ؼ���λ
Public Type ty_ctlObject_Locale
    '�ؼ���λ��
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    '�����б����С�߶ȺͿ��
    minWidth As Single
    minHeight As Single
    
    '�½��б��ʵ��λ��
    DownLeft As Single
    DownTop As Single
    DownWidth As Single
    DownHeight As Single
 
    
    '��ģ���
    ScreenWidth As Single
    ScreenHeight As Single
    
End Type

Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Declare Sub GetKeyboardStateByString Lib "user32" Alias "GetKeyboardState" (ByVal pbKeyState As String)
Public Declare Sub SetKeyboardStateByString Lib "user32" Alias "SetKeyboardState" (ByVal lppbKeyState As String)
Public Const VK_CAPITAL = &H14
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_SHOWDROPDOWN = &H14F
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH = &H101E
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

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
'Windows���----------------------------------
Public Const ETO_OPAQUE = 2
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'���뷨����API----------------------------------------------------------------------------------------------
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Const KLF_REORDER = &H8
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 'ǳ����
Public Const BDR_RAISEDINNER = &H4 'ǳ͹��
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame������ʽ
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��Frame������ʽ
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
Public Enum gAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum
Public Enum EM_DrawStyle
    DW_Flat = 0  '= ƽ��
    Dw_SubKen = -1 '= ����
    Dw_Heave = 1  '= ͹��
    Dw_Deepen_Subken = -2 '= ���,
    Dw_Deepen_Heave = 2 ' = ��͹��
End Enum

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'��ͼAPI
Public Const PS_SOLID = 0
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional strName As String = "")
'���ܣ���PictureBoxģ���3Dƽ�水ť
'������intStyle:0=ƽ��,-1=����,1=͹��
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Function SystemImes() As Variant
'���ܣ���ϵͳ�������뷨���Ʒ��ص�һ���ַ���������
'���أ�����������������뷨,�򷵻ؿմ�
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Long, j As Long
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function OpenIme(Optional strIme As String) As Boolean
'����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If strIme = "���Զ�����" Then OpenIme = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function

Public Sub AutoSizeCol(lvw As Object)
'���ܣ������Զ�ListView��ǰ�����Զ��������п��
'������blnByHead=�Ƿ���ͷ�ı�����,Col=ָ���л���������(1-N)
    Dim i As Long, lngW As Long
    For i = 1 To lvw.ColumnHeaders.Count
        SendMessage lvw.hWnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If lvw.ColumnHeaders(i).Width < 200 Then lvw.ColumnHeaders(i).Width = 0
        If lvw.ColumnHeaders(i).Width < (zlCommFun.ActualLen(lvw.ColumnHeaders(i).Text) + 2) * 90 And lvw.ColumnHeaders(i).Width <> 0 Then lvw.ColumnHeaders(i).Width = (zlCommFun.ActualLen(lvw.ColumnHeaders(i).Text) + 2) * 90
    Next
End Sub

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

Public Function PreFixNO(Optional Curdate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If Curdate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(Curdate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
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

Public Function GetBillTotal(objBill As ExpenseBill) As Currency
'���ܣ���ȡ���ݷ�Ŀ�ϼƽ��
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    
    For Each objBillDetail In objBill.Details
        For Each objBillIncome In objBillDetail.InComes
            GetBillTotal = GetBillTotal + objBillIncome.ʵ�ս��
        Next
    Next
End Function

Public Function GetBillRowTotal(objBillInComes As BillInComes) As Currency
'���ܣ���ȡ���ݷ�Ŀ�ϼƽ��
    Dim objBillIncome As New BillInCome
    For Each objBillIncome In objBillInComes
        GetBillRowTotal = GetBillRowTotal + objBillIncome.ʵ�ս��
    Next
End Function


Public Function HaveExist(cbo As ComboBox, str As String, lng As Long) As Boolean
'���ܣ��ж�ָ����Ŀ���б����Ƿ��Ѿ�����
'˵������ͬ��ĿָText��ItemData����ͬ
    Dim i As Long
    For i = 0 To cbo.ListCount
        If cbo.List(i) = str And cbo.ItemData(i) = lng Then
            HaveExist = True: Exit For
        End If
    Next
End Function

Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Public Function GetFirstRow(curBill As ExpenseBill, Optional strClass As String) As Integer
'���ܣ���ȡ��ǰ�����е�һ��ΪҩƷ���շ��к�
'������strClass=ȡ��һ��ҩ����ҩ��,��ΪҩƷ
'���أ�0=û��ҩƷ�շ���
    Dim i As Long
    If curBill.Details.Count = 0 Then GetFirstRow = 0
    For i = 1 To curBill.Details.Count
        If strClass = "" Then
            If InStr(",5,6,7,", curBill.Details(i).�շ����) > 0 Then
                GetFirstRow = i: Exit Function
            End If
        Else
            If curBill.Details(i).�շ���� = strClass Then
                GetFirstRow = i: Exit Function
            End If
        End If
    Next
End Function

Public Function GetFirstClass(curBill As ExpenseBill) As String
'���ܣ���ȡ��ǰ�����е�һ��ΪҩƷ���շ��к�
'���أ�0=û��ҩƷ�շ���
    Dim i As Long
    If curBill.Details.Count = 0 Then GetFirstClass = ""
    For i = 1 To curBill.Details.Count
        If InStr(",5,6,7,", curBill.Details(i).�շ����) > 0 Then
            GetFirstClass = curBill.Details(i).�շ����: Exit Function
        End If
    Next
End Function

Public Function Custom_WndMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.x = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.x = 1600
        MinMax.ptMaxTrackSize.y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, msg, wp, lp)
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function strPad(ByVal strPre As String, ByVal intLen As Integer, ByVal strFill As String, ByVal bytAlign As Byte, Optional ByVal blnTrim As Boolean) As String
'���ܣ�����ַ���
'������
'     strPre=Ҫ�����ַ���
'     intLen=����ĳ���
'     strFill=Ҫ�����ַ�
'     bytAlign=1,2/��,�Ҷ��룬�����ʱ����ԭ�ַ����ұ����
'     blnTrim=���ַ�������ʱ���Ƿ�ǿ�а�ָ�����Ƚ�ȡ��
'���أ��������ַ���
'˵����һ�����ֵ��������ַ����ȴ���
    Dim i As Long
    
    If LenB(StrConv(strPre, vbFromUnicode)) >= intLen Then
        If blnTrim Then
            For i = 1 To Len(strPre)
                strPad = strPad & Mid(strPre, i, 1)
                If LenB(StrConv(strPad, vbFromUnicode)) >= intLen Then Exit For
            Next
        Else
            strPad = strPre
        End If
    Else
        If Len(strFill) > 1 Then strFill = Left(strFill, 1)
        If bytAlign = 1 Then
            strPad = strPre
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
        ElseIf bytAlign = 2 Then
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
            strPad = strPad & strPre
        End If
    End If
End Function

Public Sub PrintCell(ByVal Text As String, _
    ByVal x As Single, ByVal y As Single, _
    Optional ByVal Wide, _
    Optional ByVal High, _
    Optional Alignment As Byte = 0, _
    Optional ForeColor As Long = 0, _
    Optional GridColor As Long = 0, _
    Optional FillColor As Long = 0, _
    Optional LineStyle As String = "1111", _
    Optional FontName, Optional FontSize, _
    Optional FontBold, Optional FontItalic)
    '------------------------------------------------
    '���ܣ� ��ָ�������ӡһ�����ݵ�Ԫ,������ǰ�����ƶ�����Ԫ���Ͻ�λ��
    '������
    '   Text:    ������ַ���,���в������س����з�
    '   X:       ���Ͻ�X����
    '   Y:       ���Ͻ�Y����
    '   Wide:    ������
    '   High:    ����߶�
    '   Alignment:    ����ģʽ��0-�����(ȱʡ),1-�Ҷ���,2-����
    '   ForeColorǰ��ɫ,ȱʡΪ��ɫ
    '   GridColor����ɫ,ȱʡΪ��ɫ
    '   FillColor���ɫ,ȱʡΪ�豸����ɫ,����ϵͳ�����˺�ɫ��ɫ�룬���Խ�����������ɫ
    '   LineStyle:����ֱ�Ϊ�������µ��������
    '           0-���ߣ�1-9����Ӵ֣�1Ϊȱʡ
    '   FontName,FontSize,FontBold,FontItalic:��������
    '���أ�
    '------------------------------------------------
    Dim aryString() As String       '�س��ָ���ַ���
    Dim lngOldForeColor As Long     '����豸ȱʡǰ��ɫ
    Dim intRow As Integer, intAllRow As Integer
    Dim strRest As String, sngYMove As Single
    Dim oldFontName, oldFontSize, oldFontBold, oldFontItalic
    lngOldForeColor = gobjOutTo.ForeColor
    
    On Error Resume Next
    With gobjOutTo
        If Not IsMissing(FontName) Then
            oldFontName = gobjOutTo.FontName
            .FontName = FontName
        End If
        If Not IsMissing(FontSize) Then
            .FontSize = FontSize
            oldFontSize = gobjOutTo.FontSize
        End If
        If Not IsMissing(FontBold) Then
            .FontBold = FontBold
            oldFontBold = gobjOutTo.FontBold
        End If
        If Not IsMissing(FontItalic) Then
            .FontItalic = FontItalic
            oldFontItalic = gobjOutTo.FontItalic
        End If
    End With
    
    If IsMissing(Wide) Then Wide = gobjOutTo.TextWidth(Text) + 2 * conLineWide
    If IsMissing(High) Then High = gobjOutTo.TextHeight(Text) + 2 * conLineHigh
'    Wide = CLng(Wide)
'    High = CLng(High)
    If Wide * High = 0 Then Exit Sub
    
    If UCase(TypeName(LineStyle)) <> "STRING" Then LineStyle = CStr(LineStyle)
    If Len(LineStyle) < 4 Then
        LineStyle = Left(LineStyle & "1111", 4)
    End If
    
    '------------------------------------------
    '   ���ߴ�ӡ
    '------------------------------------------
    If Mid(LineStyle, 1, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 1, 1)
        gobjOutTo.Line (x, y)-(x + Wide, y), GridColor
    End If
    
    If Mid(LineStyle, 2, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 2, 1)
        gobjOutTo.Line (x, y)-(x, y + High), GridColor
    End If
    
    If Mid(LineStyle, 3, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 3, 1)
        gobjOutTo.Line (x + Wide, y)-(x + Wide, y + High), GridColor
    End If
    
    If Mid(LineStyle, 4, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 4, 1)
        gobjOutTo.Line (x, y + High)-(x + Wide, y + High), GridColor
    End If
    
    If Wide > conLineWide And High > conLineHigh Then
        '------------------------------------------
        '   ��ɫ���
        '------------------------------------------
'        If FillColor <> 0 Then
'            Printer.FillStyle = 1
'            gobjOutTo.Line (X + conLineWide / 2, Y + conLineHigh / 2)- _
'                (X + Wide - conLineWide / 2, Y + High - conLineHigh / 2), _
'                FillColor, BF
'        End If
        
        '------------------------------------------
        '   ���ִ�ӡ
        '------------------------------------------
        gobjOutTo.ForeColor = ForeColor
    
        If InStr(Text, vbCrLf) = 0 And InStr(Text, Chr(13)) = 0 Then
            If Wide - conLineWide < gobjOutTo.TextWidth("1") Then    'С��һ���ַ�
                intAllRow = 1
            Else
                If gobjOutTo.TextWidth(Text) Mod (Wide - conLineWide) = 0 Then
                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide)
                Else
                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide) + 1
                End If
            End If
            For intRow = intAllRow To 1 Step -1
                If High >= gobjOutTo.TextHeight(Text) * intRow Then
                    Exit For
                End If
            Next
            intAllRow = intRow
            sngYMove = (High - conLineHigh - gobjOutTo.TextHeight(Text) * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow - 1
                Do While gobjOutTo.TextWidth(Text) > Wide - conLineWide
                    If Len(Trim(Text)) <= 1 Then Exit Do
                    Text = Left(Text, Len(Text) - 1)
                Loop
                strRest = Mid(strRest, Len(Text) + 1)
                Select Case Alignment
                Case 2
                    gobjOutTo.CurrentX = x + (Wide - gobjOutTo.TextWidth(Text)) / 2
                Case 1
                    gobjOutTo.CurrentX = x - conLineWide / 2 + Wide - gobjOutTo.TextWidth(Text)
                Case Else
                    gobjOutTo.CurrentX = x + conLineWide / 2
                End Select
                gobjOutTo.CurrentY = y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(Text)
                gobjOutTo.Print Text
                Text = strRest
            Next
        Else
            If InStr(Text, vbCrLf) > 0 Then
                aryString = Split(Trim(Text), vbCrLf)
            Else
                aryString = Split(Trim(Text), Chr(13))
            End If

            intAllRow = UBound(aryString)
            sngYMove = (High - conLineHigh - gobjOutTo.TextHeight("ZYL") * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow
                strRest = aryString(intRow)
                Select Case Alignment
                Case 2
                    Dim blnLR As Boolean
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        blnLR = Not blnLR
                        strRest = IIf(blnLR, Left(strRest, Len(strRest) - 1), Right(strRest, Len(strRest) - 1))
                    Loop
                    gobjOutTo.CurrentX = x + (Wide - gobjOutTo.TextWidth(strRest)) / 2
                Case 1
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        strRest = Right(strRest, Len(strRest) - 1)
                    Loop
                    gobjOutTo.CurrentX = x - conLineWide / 2 + Wide - gobjOutTo.TextWidth(strRest)
                Case Else
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        strRest = Left(strRest, Len(strRest) - 1)
                    Loop
                    gobjOutTo.CurrentX = x + conLineWide / 2
                End Select
                
                gobjOutTo.CurrentY = y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(strRest)
                If gobjOutTo.CurrentY + gobjOutTo.TextHeight(strRest) > y + High Then Exit For
                If gobjOutTo.CurrentY >= y Then gobjOutTo.Print strRest
            
            Next
        End If
    End If
    gobjOutTo.CurrentX = x + Wide
    gobjOutTo.CurrentY = y
    gobjOutTo.DrawStyle = 0
    gobjOutTo.DrawWidth = 1
    gobjOutTo.ForeColor = lngOldForeColor

    If Not IsMissing(FontName) Then gobjOutTo.FontName = oldFontName
    If Not IsMissing(FontSize) Then gobjOutTo.FontSize = oldFontSize
    If Not IsMissing(FontBold) Then gobjOutTo.FontBold = oldFontBold
    If Not IsMissing(FontItalic) Then gobjOutTo.FontItalic = oldFontItalic
End Sub

Public Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
    Dim i As Long, strTmp As String, bytUp As Byte, bytAdd As Byte
    
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
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function CheckScope(varL As Double, varR As Double, varI As Double) As String
'���ܣ��ж��������Ƿ���ԭ�ۺ��ִ��޶��ķ�Χ��
'������varL=ԭ��,varR=�ּ�,varI=������
'���أ�������ڷ�Χ��,��Ϊ��ʾ��Ϣ,����Ϊ�մ�
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '�����ֵ������ͬ,���þ���ֵ�ж�
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "����ļ۸����ֵ���ڷ�Χ(" & FormatEx(Abs(varL), gintFeePrecision) & "-" & FormatEx(Abs(varR), gintFeePrecision) & ")��."
        End If
    Else
        '������Ų���ͬ,����ԭʼ��Χ�ж�
        If varI < varL Or varI > varR Then
            CheckScope = "����ļ۸�ֵ���ڷ�Χ(" & FormatEx(varL, gintFeePrecision) & "-" & FormatEx(varR, gintFeePrecision) & ")��."
        End If
    End If
End Function

Public Sub BandRectoGrid(objGrid As Object, rsData As ADODB.Recordset)
    Dim blnPre As Boolean, i As Long, j As Long
    
    objGrid.Clear: objGrid.Rows = 2: objGrid.Cols = 2
    objGrid.FixedRows = 1: objGrid.FixedCols = 0
    
    If rsData Is Nothing Then Exit Sub
    If rsData.State = adStateClosed Then Exit Sub
    
    blnPre = objGrid.Redraw
    objGrid.Redraw = False
    
    objGrid.Cols = rsData.Fields.Count
    objGrid.Rows = IIf(rsData.RecordCount = 0, 2, rsData.RecordCount + 1)
    objGrid.FixedRows = 1
    
    For j = 0 To rsData.Fields.Count - 1
        objGrid.TextMatrix(0, j) = rsData.Fields(j).Name
    Next
    
    If rsData.RecordCount = 0 Then objGrid.Redraw = blnPre: Exit Sub
    
    rsData.MoveFirst
    For i = 1 To rsData.RecordCount
        For j = 0 To rsData.Fields.Count - 1
            objGrid.TextMatrix(i, j) = "" & rsData.Fields(j).Value
        Next
        rsData.MoveNext
    Next
End Sub

Public Sub SetGridWidth(msh As Control, frmParent As Object)
'���ܣ��Զ���������п�,����С�ʺ�Ϊ׼
    Dim blnRedraw As Boolean
    Dim blnDo As Boolean, i As Long, j As Long, strText As String
    Dim lngStart As Long, lngEnd As Long, lngMaxLen As Long, lngCurLen As Long, lngMWRow As Long
        
    On Local Error Resume Next
    
    blnRedraw = msh.Redraw
    msh.Redraw = False
    lngStart = IIf(msh.FixedRows = 0, 0, msh.FixedRows - 1)
    lngEnd = msh.Rows - 1
    
    For i = 0 To msh.Cols - 1
        lngMaxLen = LenB(StrConv(msh.TextMatrix(0, i), vbFromUnicode))  '����Ϊ�������
        lngMWRow = 0
        For j = lngStart To lngEnd
            blnDo = True
            strText = msh.TextMatrix(j, i)
            
            If msh.MergeRow(j) Then
                If i > 0 Then If strText = msh.TextMatrix(j, i - 1) Then blnDo = False
                If blnDo Then
                    If i < msh.Cols - 1 Then If strText = msh.TextMatrix(j, i + 1) Then blnDo = False
                End If
            End If
            If blnDo Then
                lngCurLen = LenB(StrConv(strText, vbFromUnicode))
                If lngCurLen > lngMaxLen Then
                    lngMaxLen = lngCurLen
                    lngMWRow = j
                End If
            End If
        Next
        msh.ColWidth(i) = frmParent.TextWidth(msh.TextMatrix(lngMWRow, i)) + 100
        If msh.ColWidth(i) > 3090 Then msh.ColWidth(i) = 3000
    Next
    
    msh.Redraw = blnRedraw
End Sub

Public Function CheckInput(txt As TextBox, strName As String, Optional blnAllowNULL As Boolean) As Boolean
'���ܣ���鹤�������ʵ�����Ƿ���ָ�����Ƴ�����
    If txt.Text = "" And Not blnAllowNULL Then
        MsgBox strName & "������Ϊ�գ�", vbInformation, gstrSysName
        txt.SetFocus: Exit Function
    ElseIf LenB(StrConv(txt.Text, vbFromUnicode)) > txt.MaxLength Then
        MsgBox strName & "ֻ�������� " & txt.MaxLength & " ���ַ��� " & txt.MaxLength \ 2 & " �����֣�", vbInformation, gstrSysName
        txt.SetFocus: Exit Function
    End If
    CheckInput = True
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function CheckLen(txt As TextBox, intLen As Integer) As Boolean
'���ܣ���鹤�������ʵ�����Ƿ���ָ�����Ƴ�����
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox Mid(txt.Name, 4) & "ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�", vbInformation, gstrSysName
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function


Public Sub SaveRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
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
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
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
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, _
    Optional blnShowZero As Boolean = True) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = IIf(blnShowZero, 0, "")
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
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
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    
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

Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * Val(vNumber))
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

Public Function CentMoney(ByVal curMoney As Currency) As Currency
'���ܣ���ָ�����ֱҴ��������д���,���ش����Ľ��
'������curMoney=Ҫ���зֱҴ���Ľ��(ΪӦ�ɽ��,2λС��)
'      gBytMoney=
'         0.������
'         1.��ȡ�������뷨,eg:0.51=0.50;0.56=0.60
'         2.�����շ�,eg:0.51=0.60,0.56=0.60
'         3.����շ�,eg:0.51=0.50,0.56=0.50
'         4.�����������˫,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           �����������˫,����ҹ���ѧ����ίԱ����ʽ�䲼�ġ�������Լ����,������vb��Round����,�������������ְ�����λ����ʱ�����Ը����ֽ���������Լ
'           �����м����뷨:���������忼�ǣ�������ͽ�һ�������㿴��ż����ǰΪżӦ��ȥ����ǰΪ��Ҫ��һ
'         5.�������塢�������,�Խǽ��д�������Ҫ�ȶԷֱҽ�������,��0.29(��)���¶�����ǣ�0.80(��)���϶����ǣ�0.3-0.79����Ϊ0.5��
'         6.��������:eg:0.15=0.10:0.16=0.2:   ���˺� ����:34519  ����:2010-12-06 09:58:02
'91385,������5.�������塢������롱�����ȶԷֱҽ����������룬��0.24(��)���¶�����ǣ�0.75(��)���϶����ǣ�0.25-0.74������Ϊ0.5
'       �ֱ����������룬��ô0.00��0.24=0��0.25��0.5=0.50, 0.50��0.74=0.50��0.75��1.00=1������������ռ50%�ı���
    
    Dim intSign As Integer, curTmp As Currency

    If gBytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf gBytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '��ȡ��λ���,�ٴ���ֱ�,��:0.248 ��0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf gBytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf gBytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf gBytMoney = 4 Then
        CentMoney = Format(FormatEx(curMoney, 1), "0.00")
    ElseIf gBytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf gBytMoney = 6 Then
         '���˺� ����:34519 ��������:eg:0.15=0.10:0.16=0.2:    ����:2010-12-06 09:58:02
          CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function

Public Function Between(x, a, B) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < B Then
        Between = x >= a And x <= B
    Else
        Between = x >= B And x <= a
    End If
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Sub ExChangeLocate(objA As Object, objB As Object)
'���ܣ�����ҽ���Ϳ������ҵ�����λ��
    Dim X1 As Long, Y1 As Long, w1 As Long, t1 As Integer
    Dim X2 As Long, Y2 As Long, w2 As Long, t2 As Integer
    Dim obj1 As Object, obj2 As Object
    
    X1 = objA.Left
    Y1 = objA.Top
    w1 = objA.Width
    t1 = objA.TabIndex
    Set obj1 = objA.Container

    X2 = objB.Left
    Y2 = objB.Top
    w2 = objB.Width
    t2 = objB.TabIndex
    Set obj2 = objB.Container
    
    Set objB.Container = obj1
    If TypeName(objB) = "Label" Then
        objB.Left = X1 + w1 - objB.Width
    Else
        objB.Left = X1
        objB.Width = w1
    End If
    objB.Top = Y1
    objB.TabIndex = t1
    
    Set objA.Container = obj2
    If TypeName(objA) = "Label" Then
        objA.Left = X2 + w2 - objA.Width
    Else
        objA.Left = X2
        objA.Width = w2
    End If
    objA.Top = Y2
    objA.TabIndex = t2
End Sub

Public Function MoneyOverFlow(objBill As ExpenseBill) As Boolean
'���ܣ���鵥�ݺϼƽ���Ƿ����
'˵������Currency����922337203685477Ϊ׼
    Dim dblӦ�� As Double, dblʵ�� As Double
    Dim i As Integer, j As Integer
    
    'Ҫ��VALתΪDouble��������
    For i = 1 To objBill.Details.Count
        For j = 1 To objBill.Details(i).InComes.Count
            If Abs(dblӦ�� + Val(objBill.Details(i).InComes(j).Ӧ�ս��)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            If Abs(dblʵ�� + Val(objBill.Details(i).InComes(j).ʵ�ս��)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            dblӦ�� = dblӦ�� + Val(objBill.Details(i).InComes(j).Ӧ�ս��)
            dblʵ�� = dblʵ�� + Val(objBill.Details(i).InComes(j).ʵ�ս��)
        Next
    Next
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.x = lngX / Screen.TwipsPerPixelX: vPoint.y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.x = vPoint.x * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

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

Public Function ZVal(ByVal varValue As Variant) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    ZVal = IIf(Val(varValue) = 0, "NULL", varValue)
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
    Err = 0: On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
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
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional bln������� As Boolean = True, Optional bln���� As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ��Ľ��
    '���:strInput        ������ַ���
    '     intMax          ������λ��
    '     bln�������     �Ƿ���и������
    '     bln����         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
   
    Dim dblValue As Double
    If bln���� = True Then
        If strInput = "" Then
            ShowMsgbox str��Ŀ & "δ���룬����!"
            If hWnd <> 0 Then SetFocusHwnd hWnd
            Exit Function
        End If
    End If
    If strInput = "" Then zlDblIsValid = True: Exit Function
    
    If IsNumeric(strInput) = False Then
        MsgBox str��Ŀ & "������Ч�����ָ�ʽ��", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str��Ŀ & "��ֵ���󣬲��ܳ���" & 10 ^ intMax - 1 & "��", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    If bln������� = True And dblValue < 0 Then
        MsgBox str��Ŀ & "�������븺����", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str��Ŀ & "��ֵ��С������С��-" & 10 ^ intMax - 1 & "λ��", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    
    
    If bln���� = True And dblValue = 0 Then
        MsgBox str��Ŀ & "���������㡣", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    zlDblIsValid = True
End Function

Public Function Where����ʱ��(Optional strAlias As String) As String
    If strAlias = "" Then
        Where����ʱ�� = " (����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or ����ʱ�� is null) "
    Else
        Where����ʱ�� = " (" & strAlias & ".����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or " & strAlias & ".����ʱ�� is null) "
    End If
End Function

Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
'���ܣ��жϵ�ǰ��Ļ����Ƿ���ָ�����ڵ���ʾ������
    Dim vRect As RECT, vPos As POINTAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.x >= vRect.Left And vPos.x <= vRect.Right _
        And vPos.y >= vRect.Top And vPos.y <= vRect.Bottom Then
        MouseInRect = True
    End If
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


Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function
Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function

Public Sub zlRaisEffect(picBox As Object, Optional intStyle As EM_DrawStyle, _
    Optional strName As String = "", Optional TxtAlignment As gAlignment = 1)
    '���ܣ���PictureBoxģ���3Dƽ�水ť
    'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            If intStyle = 2 Then
                    DrawEdge .hDC, PicRect, EDGE_RAISED Or BF_SOFT, BF_RECT
            ElseIf intStyle = -2 Then
                    DrawEdge .hDC, PicRect, EDGE_SUNKEN Or BF_SOFT, BF_RECT
            Else
                DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
            End If
        End If
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If TxtAlignment = mCenterAgnmt Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            ElseIf TxtAlignment = mLeftAgnmt Then
                .CurrentX = .ScaleLeft
            Else
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) '-10
            End If
            picBox.Print strName
        End If
        .ScaleMode = lngTmp
        .Refresh
    End With
End Sub

Public Function GetModuleType() As Byte
    '99993:���ϴ�,2016/8/26,BH����ˢ������
    If gfrmMain Is Nothing And glngMain = 0 Then
        GetModuleType = 0
    Else
        GetModuleType = 1
    End If
End Function
