Attribute VB_Name = "mdlPublic"
Option Explicit 'Ҫ���������
Public gclsInsure As New clsInsure          'ҽ���ӿڶ���
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrPrivsStation As String '��ǰ�û���ҽ������վ��Ȩ��  ֻ��ͨ���ӿڵ���ʱ,�Ŵ���
Public gstrSysName As String                'ϵͳ����
Public gstrUnitName As String
Public glngSys As Long
Public glngModul As Long
Public gstrProductName As String

Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"
Public gbytDec As Byte '���ý���С����λ��
Public gbyt���������Ϣ As Byte '0-�����;1-���;2-��ʾ���
Public gblnOk As Boolean
Public gstrDBUser As String '��ǰ�û���
Public gfrmMain As Object
Public glngMain As Long
'�û���Ϣ------------------------
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

'ϵͳ����
Public Type TY_Reg_Para  '�Һ���ز���
    bytNODaysGeneral As Byte    '��ͨ�Һ���Ч����
    bytNoDayseMergency As Byte '����Һ���Ч����
End Type
Public Type TY_SysPara
    Sy_Reg  As TY_Reg_Para
End Type
Public gSysPara As TY_SysPara       'ϵͳ�������;�Ժ������չ(���˺�)


Public Type TY_VisitPlan_ModulePara '�ٴ����ﰲ��ģ�����
    byt������ӡ��ʽ As Byte
    str��Դά��վ�� As String 'δ����վ��Ŀ��Һ�Դ��ά��վ��
    byt����ȽϷ�ʽ  As Byte '��Դ���밴���ֱȽϷ�ʽ��������0-���ַ��Ƚϣ�1-����ֵ�Ƚ�
End Type
Public gVisitPlan_ModulePara As TY_VisitPlan_ModulePara

Public gstrLike As String   '����ƥ�䷽ʽ
Public glngInterval As Long '�ҺŰ��ű��Զ�ˢ�¼��,0��ʾ���Զ�ˢ��
Public gbytRegistMode As Byte '�Һ�ģʽ
Public gdatRegistTime As Date '�����ģʽ����ʱ��

Public gblnSharedInvoice As Boolean '�Һ�ʹ���շ�Ʊ��
Public gblnBill�Һ� As Boolean '�Ƿ��ϸ����Ʊ��

Public gbytFactLength As Byte '�Һ�Ʊ�ݺ��볤��
Public glng�Һ�ID As Long '�Һ�����ID
Public gbln������֤ As Boolean '����һ��ͨ���Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤

Public gstr�ſ�ID As String  '���￨����ID
'Public gblnBill�ſ� As Boolean '�Ƿ��ϸ����Ʊ��
'Public gbyt�ſ� As Byte '���￨�ų���
Public gstrCardPass As String 'ˢ��ʱҪ����������,'0000000000'��λ˳���ʾ��������,�ֱ�Ϊ:1.����Һ�,2.���ﻮ��,3.�����շ�,4.�������,5.��Ժ�Ǽ�,6.סԺ����,7.���˽���,8.����Ԥ����,9.���鼼ʦվ,10.Ӱ��ҽ��վ.'
Public gblnPrePayPriority As Boolean '����ʹ��Ԥ����

Public gintԤԼ���� As Integer '�Һ������ԤԼ����
Public gstr�ϰ�ʱ�� As String

Public gstr�Һſ���ID As String   '������վ����ҺŵĿ���ID
Public gstrIme As String '�Զ����������뷨

'��ѡ������Ŀ
Public gbln���� As Boolean '����
Public gbln�Ա� As Boolean  '�Ա�
Public gbln���� As Boolean  '����
Public gbln��ͥ��ַ As Boolean  '��ͥ��ַ
Public gbln���ʽ As Boolean  '���ʽ
Public gbln�ѱ� As Boolean '�ѱ�
Public gbln���㷽ʽ As Boolean '���㷽ʽ
Public gblnҽ�� As Boolean 'ҽ��
Public gbln�绰 As Boolean

'ȱʡֵ
Public gstr���ʽ As String 'ȱʡ���ʽ
Public gstr�ѱ� As String 'ȱʡ�ѱ�
Public gstr�Ա� As String 'ȱʡ�Ա�
Public gstr���㷽ʽ As String 'ȱʡ���㷽ʽ
'���˺� ����:????    ����:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '����С������
Public gstrFeePrecisionFmt As String '����С����ʽ:0.00000

'��������
Public gbln�ɿ���� As Boolean
Public gbln�Զ������ As Boolean
Public gblnAutoAddName As Boolean '����ʱ�Զ�������ʱ����
Public gblnNewCardNoPop As Boolean '����ʱ����������������
Public gbln���ѽ����� As Boolean
Public gbln�˷��ش� As Boolean '�˺Ų��˿�ʱ�Ƿ��ش�Ʊ
Public gint�ų� As Integer '�ű𳤶�
Public gblnLED As Boolean
Public gblnPrintFree As Boolean
Public gblnPrintCase As Boolean '��ӡ������ǩ
Public gbytInvoice As Byte   '��Ʊ��ӡ��ʽ
Public gByt��ӡ�������� As Byte '�������� ��ӡ��ʽ
Public gblnPrice As Boolean     '�������˹ҺŴ�Ϊ���۵�
Public gintNameDays As Integer  '������������N���ڵĲ���
Public gblnSeekName As Boolean
Public gByt�Һ�ƾ�� As Byte     '�Һ�ƾ����ӡ��ʽ
Public gBytԤԼ�Һŵ� As Byte  'ԤԼ�Һŵ���ӡ��ʽ
Public gByt�˺Żص� As Byte     '�˺Żص���ӡ��ʽ

Public glngOld As Long
Public glngMinW As Long, glngMaxW As Long
Public glngMinH As Long, glngMaxH As Long
Public gbln���֤Ψһ As Boolean
'WIN32����

'API����
Public Const CB_ADDSTRING = &H143
Public Const CB_SETITEMDATA = &H151
Public Const CB_FINDSTRING = &H14C
Public Const CB_SHOWDROPDOWN = &H14F


Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetComboData Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindComboStr Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24

Public Const GWL_STYLE = (-16)              'Set the window style
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000        '��߿�
Public Const WS_SYSMENU = &H80000           '�ڱ������Ƿ�߱�ϵͳ�˵�
Public Const WS_MINIMIZEBOX = &H20000       '�߱���С����ť
Public Const WS_MAXIMIZEBOX = &H10000       '�߱���󻯰�ť
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
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

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_RESETCONTENT = &H14B
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH = &H101E

'�ƶ��ؼ����ޱ߿���
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal Hwnd As Long) As Long
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MOVE = &HF010&
Public Const HTCAPTION = 2

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_F4 = vbKeyF4

Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Dim p As KBDLLHOOKSTRUCT
Public p1 As KBDLLHOOKSTRUCT
Public gblnBegin As Boolean
Public gblnLen As Boolean
Public gblnCard As Boolean
Public gsngStartTime As Single

'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long

'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long

'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'''''''''''''''''''''
'��ȡָ�����뷨����Layout,����Ϊ0ʱ��ʾ��ǰ���뷨��
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'��ȡ��ǰ���뷨����Layout��
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'�������뷨Layout���������뷨�л������뷨�л�˳�����ǰͷ(������������Ч),flags����=KLF_REORDER
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Public Const KLF_REORDER = &H8
'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Type Ty_CardProperty
       lng�����ID      As Long
       str������        As String
       str������        As String
       lng���ų���      As Long
       lng���㷽ʽ      As String
       bln���ƿ�        As Boolean
       bln�ϸ����      As Boolean
       lng����ID        As Long
       lng��������      As Long
       bln���          As Boolean
       blnˢ��          As Boolean
       int���볤��      As Integer
       int���볤������  As Integer
       int�������      As Integer
       bln���￨        As Boolean
       str��������      As String
       str��׼��Ŀ      As String
       blnȱʡ��־      As Boolean
       blnOneCard       As Boolean '  '�Ƿ�������һ��ͨ�ӿ�,��ģʽ�£�Ʊ���ϸ����Ʊ�ŷ�Χ��ķ�����󶨿����շ�
       rs����           As ADODB.Recordset
       dblӦ�ս��      As Double
       dblʵ�ս��      As Double
       bln�Ƿ��ƿ�      As Boolean
       bln�Ƿ񷢿�      As Boolean
       bln�Ƿ�д��      As Boolean
       lng��������      As Long '0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0 �����:57326
       bln�ظ�ʹ��      As Boolean
       byt��������      As Byte
       lng�շ�ϸĿID    As Long 'ҽԺ����������ѷ��ص��շ�ϸĿID,���뵱ǰ���ѵ��շ�ϸĿIDͬ��
End Type
Public gCurSendCard As Ty_CardProperty
Public gstrSQL  As String

Public Const TVM_SETBKCOLOR = 4381&
Public Const TVM_GETBKCOLOR = 4383&
Public Const TVS_HASLINES = 2&
'�ؼ�����λ�û�ȡת��
Public Const EM_EXGETSEL = (&H400 + 52)
Public Const EM_POSFROMCHAR = &HD6

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Public Enum mTextAlign
    taLeftAlign = 0
    taCenterAlign = 1
    taRightAlign = 2
End Enum

Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type


Public Function WndMessage(ByVal Hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
'���ܣ�ȥ��TextBox��Ĭ���Ҽ��˵�
    If msg <> WM_CONTEXTMENU Then
        WndMessage = CallWindowProc(glngTXTProc, Hwnd, msg, wp, lp)
    End If
End Function

Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������lngHwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�-2=δ�Ӵ���,����=ƥ�������(����ƥ�������)
'˵�����뽫�ú�����KeyPress�¼��е��á�

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
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

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim fEatKeystroke As Boolean
    Dim sngTime As Single
    Dim sngPreTime As Timer
    
    gblnCard = False
    
    sngTime = Timer
    If (nCode = HC_ACTION) Then
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            
            CopyMemory p, ByVal lParam, Len(p)
            gblnCard = (sngTime - gsngStartTime) < 0.4
            If gblnCard = False Then gblnLen = False
             
            gsngStartTime = sngTime
            fEatKeystroke = _
            ((p.vkCode = VK_TAB) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
            ((p.vkCode = VK_ESCAPE) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
            ((p.vkCode = VK_ESCAPE) And ((GetKeyState(VK_CONTROL) And &H8000) <> 0)) Or _
            ((p.vkCode = 91) Or (p.vkCode = 92) Or (p.vkCode = 93)) Or _
            ((p.vkCode = VK_F4) And (p.flags And LLKHF_ALTDOWN) <> 0) '�������д�������Alt+F4
            If p.vkCode = Asc(";") Then fEatKeystroke = True
        End If
        
        If p.vkCode = vbKeyBack Then
            LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
            Exit Function
        End If
    End If
    If (fEatKeystroke Or gblnLen) Then
        LowLevelKeyboardProc = -1
    Else
        LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
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
'���ܣ�����Ŀ���ݲ���ComboBox������ֵ
'������lngID=ComboBox����Ŀֵ
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

Public Function FindName(cbo As ComboBox) As String
'���ܣ�ȡ����ǰComboBox��ֵ(�����Ϊ�����-���ơ�)
'˵������ҪΪSQL���ʹ��
    If cbo.ListIndex = -1 Then
        FindName = "Null"
    Else
        FindName = "'" & Mid(cbo.Text, InStr(1, cbo.Text, "-") + 1) & "'"
    End If
End Function

Public Function FindText(txt As TextBox) As String
'���ܣ�����ǰTextBox��ֵת��Ϊ��׼SQL���
'˵������ҪΪSQL���ʹ��
    If Len(Trim(txt.Text)) = 0 Then
        FindText = "Null"
    Else
        FindText = "'" & txt.Text & "'"
    End If
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
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
End Sub

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

Public Function NeedName(strList As String, Optional ByVal blnLast As Boolean = False, _
Optional strSplit As String = "-") As String
    If Not blnLast Then
        NeedName = Mid(strList, InStr(strList, strSplit) + 1)
    Else
        NeedName = strList
        Do While (InStr(NeedName, strSplit)) > 0
            NeedName = Mid(NeedName, InStr(NeedName, strSplit) + 1)
        Loop
    End If
End Function
Public Function NeedCode(strList As String) As String
    If InStr(strList, "-") = 0 Then NeedCode = strList: Exit Function
    NeedCode = Mid(strList, 1, InStr(strList, "-") - 1)
End Function
Public Function Custom_WndMessage(ByVal Hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngMinW \ 15
        MinMax.ptMinTrackSize.Y = glngMinH \ 15
        MinMax.ptMaxTrackSize.X = glngMaxW \ 15
        MinMax.ptMaxTrackSize.Y = glngMaxH \ 15
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, Hwnd, msg, wp, lp)
End Function

Public Function InDesign() As Boolean
    'InDesign = False: Exit Function
    
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "")
'���ܣ���PictureBoxģ���3Dƽ�水ť
'������intStyle:0=ƽ��,-1=����,1=͹��
    
    Dim picRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If IntStyle <> 0 Then
            picRect.Left = .ScaleLeft
            picRect.Top = .ScaleTop
            picRect.Right = .ScaleWidth
            picRect.Bottom = .ScaleHeight
            DrawEdge .hdc, picRect, CLng(IIf(IntStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Function SetCboDefault(cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = 1 Then
            cbo.ListIndex = i
            SetCboDefault = i: Exit Function
        End If
    Next
    If cbo.ListCount > 0 And cbo.ListIndex = -1 Then cbo.ListIndex = 0
End Function

Public Sub AutoSizeCol(lvw As Object)
'���ܣ������Զ�ListView��ǰ�����Զ��������п��
'������blnByHead=�Ƿ���ͷ�ı�����,Col=ָ���л���������(1-N)
    Dim i As Integer, lngW As Long
    For i = 1 To lvw.ColumnHeaders.Count
        SendMessage lvw.Hwnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If lvw.ColumnHeaders(i).Width < 200 Then lvw.ColumnHeaders(i).Width = 0
        If lvw.ColumnHeaders(i).Width < (zlCommFun.ActualLen(lvw.ColumnHeaders(i).Text) + 2) * 90 And lvw.ColumnHeaders(i).Width <> 0 Then lvw.ColumnHeaders(i).Width = (zlCommFun.ActualLen(lvw.ColumnHeaders(i).Text) + 2) * 90
    Next
End Sub

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function SystemImes() As Variant
'���ܣ���ϵͳ�������뷨���Ʒ��ص�һ���ַ���������
'���أ�����������������뷨,�򷵻ؿմ�
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer
    
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

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZVal(ByVal varValue As Variant, Optional ByVal varDefault As Variant = 0) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    Dim varTmp As Variant
    varTmp = IIf(Val(varValue) = 0, varDefault, varValue)
    ZVal = IIf(Val(varTmp) = 0, "NULL", varTmp)
End Function
Public Function GetBaseDict() As ADODB.Recordset
'���ܣ����ֵ��ж�ȡ����
    Dim strSQL As String, strTmp As String, arrTmp As Variant, i As Integer
    strTmp = "����,����,����״��,ְҵ"
    arrTmp = Split(strTmp, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = arrTmp(i)
        If strSQL = "" Then
            strSQL = "Select '" & strTmp & "' ���,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strTmp
        Else
            strSQL = strSQL & " Union all Select '" & strTmp & "' ���,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strTmp
        End If
    Next
    strSQL = strSQL & " Order by ���,����"
    
    On Error GoTo errH
    Set GetBaseDict = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����,����,����״��,ְҵ")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function

Public Function GetModuleType() As Byte
    '99993:���ϴ�,2016/8/26,BH����ˢ������
    If gfrmMain Is Nothing And glngMain = 0 Then
        GetModuleType = 0
    Else
        GetModuleType = 1
    End If
End Function
