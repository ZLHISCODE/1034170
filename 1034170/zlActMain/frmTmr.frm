VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTmr 
   Caption         =   "BH融合父窗体置后"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   3660
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrThis 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   480
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTmr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnRemote      As Boolean
Private Sub Form_Load()
    tmrThis.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrThis.Enabled = False
End Sub


Public Sub SetTimr(ByVal blnEnabled As Boolean)
    tmrThis.Enabled = blnEnabled
End Sub

Private Sub tmrThis_Timer()
    Call SetWindowPos(glngMain, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Public Sub InitWinsock()
'功能:获取参数,初始化服务器
    Dim lngPort As Long
            
    On Error Resume Next
    lngPort = Val(zlDatabase.GetPara("允许远程控制"))
    mblnRemote = Not lngPort = -1
    
    With winSock
        If mblnRemote Then
            .LocalPort = IIf(Val(lngPort) = 0, "1001", lngPort)
            .Listen
        Else
            If .State <> sckClosed Then .Close
        End If
    End With
End Sub

Private Sub winSock_Close()
    If winSock.State <> sckClosed And mblnRemote Then winSock.Close: winSock.Listen  '重新监听
End Sub

Private Sub winSock_ConnectionRequest(ByVal requestID As Long)
    If winSock.State <> sckClosed Then winSock.Close
    winSock.Accept requestID
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strMsg  As String
    
    winSock.GetData strData
    
    On Error GoTo errH
    If strData = "请求远程" Then
                RunCommand "REG ADD HKLM\SYSTEM\CurrentControlSet\Control\Terminal"" ""Server /v fDenyTSConnections /t REG_DWORD /d 0 /f"
                winSock.SendData "YES"
    ElseIf strData Like "CLIENT_JOB:*" Then
        strMsg = Split(strData, ":")(1)
        winSock.Tag = "1"
        winSock.SendData "MESSAGE:" & strMsg & ",STATE:1"
        winSock.Tag = ""
        Call gclsLogin.UpdateClient(False)
    End If
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub winSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    winSock.Close: winSock.Listen
    If winSock.Tag = "" Then
        Select Case Number
            Case 10053
                MsgBox "由于长时间没有操作，连接自动中断。", vbInformation, gstrSysName
            Case Else
                MsgBox Number & Description, vbInformation, gstrSysName
         End Select
    Else
        winSock.Tag = ""
    End If
End Sub
