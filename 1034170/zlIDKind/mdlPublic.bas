Attribute VB_Name = "mdlPublic"
Option Explicit
Public gintSaveRegType As Integer
Public gstrSaveRegProceName As String
Public glngTXTProc As Long
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

Public Enum gAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24

Public Const SB_TOP = 6
Public Const WM_VSCROLL = &H115
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
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd&, ByVal nIndex&)

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetFocus Lib "user32" () As Long

Public Sub PressKey(bytKey As Byte)
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

'��ȡ��ǰ�ؼ����
Public Function GetHwnd() As Long
    Dim Hwnd As Long
    Dim PID As Long
    Dim TID As Long
    Dim hWndFocus As Long
            
    Hwnd = GetForegroundWindow
    If Hwnd <> 0 Then
        TID = GetWindowThreadProcessId(Hwnd, PID)
        AttachThreadInput App.ThreadID, TID, True
        GetHwnd = GetFocus
        AttachThreadInput App.ThreadID, TID, False
    End If
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal Hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, Hwnd, msg, wp, lp)
End Function

Public Function Nvl(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
    '-----------------------------------------------------------------------------------
    '����:ȡĳ�ֶε�ֵ
    '����:rsObj          �������ֶ�
    '     varValue       ��rsObjΪNULLֵʱ��ȡ��ֵ
    '����:�����Ϊ��ֵ,����ԭ����ֵ,���Ϊ��ֵ,�򷵻�ָ����varValueֵ
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        Nvl = varValue
    Else
        Nvl = rsObj
    End If
End Function
Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function
Public Sub zlRaisEffect(picBox As PictureBox, Optional intStyle As Integer, _
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
Public Sub TxtSelAll(objTxt As Object)
'���ܣ����༭��ĵ��ı�ȫ��ѡ��
'������objTxt=��Ҫȫѡ�ı༭�ؼ�,�ÿؼ�����SelStart,SelLength����
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    If TypeName(objTxt) = "TextBox" Then
        If objTxt.MultiLine Then
            SendMessage objTxt.Hwnd, WM_VSCROLL, SB_TOP, 0
        End If
    End If
End Sub

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
            SaveSetting "ZLSOFT", "����ģ��" & "\" & gstrSaveRegProceName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & gstrSaveRegProceName & "\" & strSection, strKey, strKeyValue
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
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & gstrSaveRegProceName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & gstrSaveRegProceName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Function InputIsCard(ByRef txtInput As Object, ByVal KeyAscii As Integer, _
    ByVal blnBrushPassShow As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ���ı����е�ǰ�����Ƿ���ˢ��(�Ƿ�ﵽ���ų��ȣ��ڵ��ó������ж�),������ϵͳ���������Ƿ�������ʾ
    '���:KeyAscii=��KeyPress�¼��е��õĲ���
    '       blnBrushPassShow-ˢ���Ƿ񿨺�������ʾ
    '����:��ˢ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-02 00:28:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
    
     'ˢ��ʱ����������ŵ��ɵ��÷�ȡ������
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then Exit Function
    
    '����ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    '�ж��Ƿ���ˢ��
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then  '�����������������ȫ���֣���Ϊ��ˢ��
        blnCard = True
    ElseIf KeyAscii > 32 Then
        sngNow = Timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '��һ̨�ʼǱ����ԣ�һ����0.014����
        End If
    End If
    'ˢ��ʱ�����Ƿ�������ʾ
    If blnCard Then
        txtInput.PasswordChar = IIf(blnBrushPassShow, "*", "")
    Else
        txtInput.PasswordChar = ""
    End If
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtInput.IMEMode = 0
    InputIsCard = blnCard
End Function

Public Sub zlInitCommLib()
   '��ʼ����������
    Err = 0: On Error Resume Next
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
        Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
        Set gobjControl = GetObject("", "zl9Comlib.clsControl")
        Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    End If
    Err = 0: On Error GoTo 0
 End Sub
