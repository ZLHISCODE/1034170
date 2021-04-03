VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaskIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'ʹ�÷�����1.�ڳ����м���һ���´���
'          2.���崰���ScaleMode=3
'          3.�ڴ����ж���������
'          4.�����������ϼ�ͼ��ǰ��ʼ���������������:
'              1).frmHwnd:�����Hwnd����(����)
'              2).Icon:IPictureDsp��Handle����(����)
'              3).Message:��ʾ�ı�(��ѡ)
'          5.���ö����AddIcon����
'          6.���޸Ķ����Icon��Message���Ըı�ͼ�����ʾ
'          7.����DelIcon����ɾ��ͼ��
'          8.Ϊ�˼��������¼�,�ڴ���ĵ�MouseMove�¼��е���MouseState����
'              ����:MouseMove�¼���X����
'ע�����1.�����б�����һ��������������Ϣ
'          2.����Ԥ�����ö����frmHwnd��Icon����

Option Explicit
'������Ϣ
Private Const NIM_ADD = &H0     '����ͼ�굽��������ʾ����
Private Const NIM_MODIFY = &H1  'ͼ�������Ѹı����Ϣ��
Private Const NIM_DELETE = &H2  'ɾ��ͼ�ꡣ
'�����ʶλ
Private Const NIF_MESSAGE = &H1  '��Ϣ�Ϸ���
Private Const NIF_ICON = &H2     'ͼ�����Ϸ���
Private Const NIF_TIP = &H4      '��ʾ�Ϸ���
Private Const WM_MOUSEMOVE = &H200 '����MouseMove��Ϣ������Ϣ�������͵������MouseMove�¼����������д�����
'���������Ϣ����
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private IconData As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long

Public Event MouseLeftDown()
Public Event MouseLeftUp()
Public Event MouseLeftDBClick()
Public Event MouseRightDown()
Public Event MouseRightUp()
Public Event MouseRightDBClick()
Private blnAdd As Boolean

Private Sub Class_Initialize()
    blnAdd = False
End Sub

Public Sub MouseState(XMessage As Single)
    Select Case CLng(XMessage / Screen.TwipsPerPixelX)
        Case WM_LBUTTONDBLCLK
            RaiseEvent MouseLeftDBClick
        Case WM_LBUTTONDOWN
            RaiseEvent MouseLeftDown
        Case WM_LBUTTONUP
            RaiseEvent MouseLeftUp
        Case WM_RBUTTONDBLCLK
            RaiseEvent MouseRightDBClick
        Case WM_RBUTTONDOWN
            RaiseEvent MouseRightDown
        Case WM_RBUTTONUP
            RaiseEvent MouseRightUp
    End Select
End Sub

'����:0=�ɹ�;1=δ���ô���Hwnd;2=δ����ͼ��;3=ʧ��
Public Function AddIcon() As Byte
    With IconData
        If .hwnd = 0 Then AddIcon = 1: Exit Function
        If .hIcon = 0 Then AddIcon = 2: Exit Function
        .cbSize = Len(IconData)
        .uCallbackMessage = WM_MOUSEMOVE
        .uID = &H1
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        If Shell_NotifyIcon(NIM_ADD, IconData) <> 0 Then
            AddIcon = 0: blnAdd = True
        Else
            AddIcon = 3
        End If
    End With
End Function

Private Function ModiIcon() As Boolean
    If Not blnAdd Then ModiIcon = False: Exit Function
    If Shell_NotifyIcon(NIM_MODIFY, IconData) <> 0 Then
        ModiIcon = True
    Else
        ModiIcon = False
    End If
End Function

Public Function DelIcon() As Boolean
    If Not blnAdd Then DelIcon = False: Exit Function
    If Shell_NotifyIcon(NIM_DELETE, IconData) <> 0 Then
        DelIcon = True
    Else
        DelIcon = False
    End If
End Function

'ͼ��(Handle)
Public Property Let Icon(lngHandle As Long)
    IconData.hIcon = lngHandle
    Call ModiIcon
End Property

Public Property Get Icon() As Long
    Icon = IconData.hIcon
End Property

'��ʾ
Public Property Let Message(strMessage As String)
    IconData.szTip = strMessage
    IconData.szTip = Trim(IconData.szTip) & Chr$(0)
    Call ModiIcon
End Property

Public Property Get Message() As String
    Message = IconData.szTip
End Property

'����Hwnd
Public Property Let frmHwnd(lngHwnd As Long)
    IconData.hwnd = lngHwnd
End Property

Public Property Get frmHwnd() As Long
    frmHwnd = IconData.hwnd
End Property
