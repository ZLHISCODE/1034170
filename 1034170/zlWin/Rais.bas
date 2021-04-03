Attribute VB_Name = "Rais"
Option Explicit
Public Const GW_OWNER = 4
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
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
Public Const WM_CLOSE = &H10
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
Public Const SRCCOPY = &HCC0020
Public Const TOGGLE_HIDEWINDOW = &H80
Public Const TOGGLE_UNHIDEWINDOW = &H40
 Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
'���ô�����ʾ����
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal Hrgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'����PictureBox����ʾ״̬
Public Declare Function ClipCursor& Lib "user32" (lpRect As RECT)
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'��С������
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Public gConnect As New ADODB.Connection  '��������
Public gLngFormID As Long '�����������
Public CollMap As Collection
Public RsAllPic As New ADODB.Recordset
Public BlnExistBill As Boolean

'��PictureBoxģ���3Dƽ�水ť
'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "")
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .Cls
        .BorderStyle = 0
        
        If IntStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            
            Select Case IntStyle
                Case 1
                    DrawEdge .hdc, PicRect, CLng(BDR_RAISEDINNER), BF_RECT
                Case 2
                    DrawEdge .hdc, PicRect, CLng(EDGE_RAISED), BF_RECT
                Case -1
                    DrawEdge .hdc, PicRect, CLng(BDR_SUNKENOUTER), BF_RECT
                Case -2
                    DrawEdge .hdc, PicRect, CLng(EDGE_SUNKEN), BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub
'
Public Function ReadPicture(rsTable As ADODB.Recordset, strField As String) As String
    '-------------------------------------------------------------
    '���ܣ���ָ���ļ�¼��ͼ���ֶθ���Ϊͼ����ʱ�ļ�
    '������
    '       rsTable��ͼ�δ洢��¼��
    '       strField��ͼ���ֶ�
    '���أ�
    '-------------------------------------------------------------
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, FileNum, J
    Dim aryChunk() As Byte
    Dim strFile As String

    On Error GoTo ErrHand

    FileNum = FreeFile
    Do While True
        strFile = gstrAviPath & "\zlNewPicture" & CStr(rsTable!���) & ".pic"
        If Len(Dir(strFile)) <> 0 Then
                Kill strFile
        Else
                Exit Do
        End If
    Loop
    Open strFile For Binary As FileNum

    lngFileSize = rsTable.Fields(strField).ActualSize
    lngModSize = lngFileSize Mod conChunkSize
    intBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    rsTable.Move 0
    For J = 0 To intBolcks
        If J = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        ReDim aryChunk(lngCurSize - 1) As Byte
        aryChunk() = rsTable.Fields(strField).GetChunk(lngCurSize)
        Put FileNum, , aryChunk()
    Next
    Close FileNum
    ReadPicture = strFile
    Exit Function

ErrHand:
    Close FileNum
    Kill strFile
    ReadPicture = ""

End Function

Public Function InitPicToRead()
    '������:����
    '��������:2000-12-12
    'Ϊ����ٶ�,�Ȱ�ͼƬ�ӿ��ж���,����ͼƬ�ļ��б�

    Dim StrFilePath As String
    If BlnExistBill = False Then Exit Function
    With RsAllPic
        Do While Not .EOF
            StrFilePath = ReadPicture(RsAllPic, "ͼƬ")
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Function
