VERSION 5.00
Begin VB.Form frmConn��ͨ 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   1
      Top             =   1455
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      Picture         =   "frmConn��ͨ.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   5325
   End
   Begin VB.Label LblNote 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���������Ľ������ݣ����Ժ�......"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1215
      TabIndex        =   2
      Top             =   495
      Width           =   2880
   End
End
Attribute VB_Name = "frmConn��ͨ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strReturnInfo As String, mlngRows As Long
Private strReturnData As String, bytData() As Byte, blnIsConnect As Boolean, blnFlag As Boolean, blnIsGet As Boolean

Public Function ReadCard(ByVal IntPort As Integer) As String
'    On Error GoTo errHand
'    Call ShowWindow(Me.hwnd, 9)
'    ReadCard = AppClient.ReadCard(IntPort)
'    Call ShowWindow(Me.hwnd, 0)
'    Exit Function
'
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function readPassword(ByVal IntPort As Integer) As String
'    On Error GoTo errHand
'    Call ShowWindow(Me.hwnd, 9)
'    readPassword = AppClient.readPassword(IntPort)
'    Call ShowWindow(Me.hwnd, 0)
'    Exit Function
'
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function ConnCenter(strServerIP As String, lngServerPort As Long, strCN As String, Optional HisUserID As Long = 0) As Boolean
'    Dim rsUser As New ADODB.Recordset
'    Dim strUser As String, strPass As String
'    Dim strDataSoure As String
'    On Error GoTo errHand
'
'    Call ShowWindow(Me.hwnd, 9)
'
'    If gcn��ͨ.State = 0 Then
'        strDataSoure = Mid(gcnOracle.ConnectionString, InStr(UCase(gcnOracle.ConnectionString), "SERVER=") + 7)
'        strDataSoure = Left(strDataSoure, InStr(strDataSoure, """;") - 1)
'
'        gcn��ͨ.ConnectionString = "Provider=MSDAORA.1;Password=his;User ID=ybuser;Data Source=" & strDataSoure & ";Persist Security Info=True"
'        gcn��ͨ.CursorLocation = adUseClient
'        gcn��ͨ.Open
'    End If
'
'    '��tab_czry����ŵ�¼ҽ��ʹ�õ��û���������
'    If HisUserID = 0 Then
'        strUser = "00"
'        strPass = "000000"
'    Else
'        gstrSQL = "Select * From tab_czry Where HISID=" & HisUserID
'        Set rsUser = gcn��ͨ.Execute(gstrSQL)
'        If rsUser.EOF Then
'            MsgBox "�û�û��ʹ��ҽ����Ȩ��", vbInformation, "Ȩ�޴���"
'            Call ShowWindow(Me.hwnd, 0)
'            Exit Function
'        End If
'        strUser = rsUser!oper
'        strPass = rsUser!Password
'    End If
'    ConnCenter = AppClient.Login(strServerIP, lngServerPort, strCN, strUser, strPass)
'
'    If ConnCenter = False Then
'        MsgBox AppClient.getMessages, vbInformation, gstrSysName
'    End If
'    Call ShowWindow(Me.hwnd, 0)
'    Exit Function
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function Query(lngRowNum As Long, lngRows As Long, Optional strMessage As String = "") As Boolean
'    Dim arrData() As String
'    On Error GoTo errHand
'
'    Call ShowWindow(Me.hwnd, 9)
'
'    arrData = AppClient.getResultSetARow(lngRowNum)
'    '��֯����ǰ�ĸ�ʽ����ʽ��;���зָ�����,���зָ�����
'    strReturnInfo = arrData(0)
'    Query = True
'    Call ShowWindow(Me.hwnd, 0)
'    Exit Function
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function Execute(str������ As String, lng���� As Long, str���� As String, str��ʾ��Ϣ As String) As Boolean
'    Dim intReturn As Integer
'    Dim intRow As Integer
'    On Error GoTo errHand
'
'    Call ShowWindow(Me.hwnd, 9)
'
'    intReturn = AppClient.executeTrade(str������, CInt(lng����), str����)
'    Execute = (intReturn = 0)
'
'    If Execute = False Then
'        MsgBox AppClient.getMessages, vbInformation, gstrSysName
'    Else
'        '��ȡ������
'        mlngRows = AppClient.GetRows
'        '��ȡ�ķ���ֵ����ʽ��chr(10)���зָ�����vbtab���зָ�����
'        strReturnInfo = AppClient.getResultSet(0, 0)
'    End If
'    Call ShowWindow(Me.hwnd, 0)
'    Exit Function
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function ConnClose() As Boolean
'    On Error GoTo errHand
'
'    Call ShowWindow(Me.hwnd, 9)
'    Call AppClient.logout
'
'    ConnClose = True
'    Unload Me
'    Exit Function
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Unload Me
End Function

Private Sub cmdCancel_Click()
'    If MsgBox("����δ�ɹ�ִ�У��Ƿ�ȡ����", vbQuestion + vbYesNo, "ȡ������") = vbYes Then
'        WriteInfo "ȡ��ҽ������"
'
'        On Error Resume Next
'        Call ShowWindow(Me.hwnd, 0)
'        Exit Sub
'    End If
End Sub

Private Sub TimeAvi_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub
'
'Private Function Byt2Long(bytInData() As Byte, lngStart As Long, lngLen As Long) As Long
'    Dim lngLoop As Long, strTemp As String
'    strTemp = ""
'    For lngLoop = lngStart To lngStart + lngLen - 1
'        strTemp = strTemp & Right("00" & hex(bytInData(lngLoop)), 2)
'    Next
'    Byt2Long = CLng("&H" & strTemp)
'End Function
'
'Private Function L2S(lngInData As Long, intStart As Integer) As String
'    Dim strTemp As String, bytTemp(3) As Byte
'    strTemp = Right("00000000" & hex(lngInData), 8)
'    bytData(intStart) = CLng("&H" & Mid(strTemp, 1, 2))
'    bytData(intStart + 1) = CLng("&H" & Mid(strTemp, 3, 2))
'    bytData(intStart + 2) = CLng("&H" & Mid(strTemp, 5, 2))
'    bytData(intStart + 3) = CLng("&H" & Mid(strTemp, 7, 2))
'End Function
'
'Private Sub wsckConn_Connect()
'    blnIsConnect = True
'End Sub
'
'Private Sub wsckConn_DataArrival(ByVal bytesTotal As Long)
'    Dim strFlag As String, lngCode As Long, lngRows As Long, lngType As Long, lngInfoLen As Long, _
'        strInfo As String, bytReturnData() As Byte, lngTheLen As Long, bytTemp As Byte
'    wsckConn.GetData strReturnData, vbString, 1
'    If strReturnData <> "R" And strReturnData <> "S" Then
'        MsgBox "�������󣺷�����Ϣ��ʽ����", vbInformation, "����"
'        WriteInfo "ҽ�����״���:������Ϣ��ʽ����[0x" & hex(asc(strReturnData)) & "]"
'        blnFlag = False
'        Exit Sub
'    End If
'    If strReturnData = "R" Then
'        wsckConn.GetData bytReturnData(), vbArray + vbByte, 13
'        lngCode = Byt2Long(bytReturnData, 0, 4)
'        lngRows = Byt2Long(bytReturnData, 4, 4)
'        lngType = bytReturnData(8)
'        lngInfoLen = Byt2Long(bytReturnData, 9, 4)
'
'        wsckConn.GetData strInfo, vbString, lngInfoLen
'
'        mlngRows = lngRows
'        If lngCode <> 0 And Len(strInfo) <> 0 Then
'            MsgBox "��������" & vbCrLf & "    " & strInfo & ";������:" & lngCode, vbInformation, "����"
'            WriteInfo "ҽ�����״���:" & strInfo
'            WriteInfo "������:" & lngCode
'            blnFlag = False
'        ElseIf lngCode <> 0 Then
'            MsgBox "�������󣬴�����Ϣδ����", vbInformation, "����"
'            WriteInfo "ҽ�����״���,δ���ش�����Ϣ"
'            blnFlag = False
'        Else
'            strReturnInfo = strInfo
'            WriteInfo "ҽ�����׽��:" & strInfo
'            WriteInfo "�����������" & lngRows
'            blnFlag = True
'        End If
'        On Error Resume Next
'        Call ShowWindow(Me.hwnd, 0)
'        unload me
'    Else
'        wsckConn.GetData bytReturnData(), vbArray + vbByte, 4
'        lngInfoLen = Byt2Long(bytReturnData, 0, 4)
'        wsckConn.GetData strInfo, vbString, lngInfoLen
'        strReturnInfo = strInfo
'        WriteInfo "ҽ�����׽��:" & strInfo
'        blnFlag = True
'    End If
'    blnIsGet = True
'End Sub
'
'Private Sub wsckConn_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    MsgBox "��������" & Description, vbInformation, "����"
'    WriteInfo "�������Ӵ���:" & Description
'    wsckConn.Close
'    blnIsConnect = False
'    blnFlag = False
''    SetPos Me.hwnd, True
'    On Error Resume Next
'    Call ShowWindow(Me.hwnd, 0)
'    unload me
'End Sub
Private Sub Form_Load()

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub
