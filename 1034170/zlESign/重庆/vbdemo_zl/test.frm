VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12900
   Icon            =   "test.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command15 
      Caption         =   "�ͻ���ǩ����֤"
      Height          =   495
      Left            =   7080
      TabIndex        =   23
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "�ͻ���ǩ��"
      Height          =   495
      Left            =   5760
      TabIndex        =   22
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ǩ��������"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command14 
      Caption         =   "��ʾ֤��"
      Height          =   495
      Left            =   4680
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "֤�鹫Կ"
      Height          =   495
      Left            =   3720
      TabIndex        =   19
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "֤�����к�"
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "֤������"
      Height          =   495
      Left            =   1560
      TabIndex        =   17
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton ֤��ʹ���� 
      Caption         =   "֤��ʹ����"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8040
      TabIndex        =   15
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Text            =   "100"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Text            =   "http://127.0.0.1:8090/ezca_signserver/services/ezcawebservice"
      Top             =   1680
      Width           =   9135
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Caption         =   "��ʼѹ������"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "������ǩ��"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ȡ�û�֤��"
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ǩ���÷���������֤"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ǩ����֤"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "test.frx":113A
      Top             =   5280
      Width           =   12255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ǩ������"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ѡ��֤��"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ؼ����"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "��ʱ"
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "д�����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "URL��ַ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private EZCAClient As Object

Public Cert_User As String
Public EZCAWebTools As New EZCA_SignTools
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Command1_Click()
Dim ls_ret, SignData As String, cert As String
SignData = "YIRiRXduIwDmQhSEUCJR63SJD0CkGYA/frusSFD9LFa6Qm3dSKwJF1v5ZSeieiEJon0SAOmwfSpr+v/UtpPaT4XI+YRCQhJhaJVgfd5JBj5sd3zAAi9TsKKIKjN2/059YJqc6z9jZs2YcFPEcLGZajDfXdbC8mG1eWeXcZ/TO38="
cert = "MIIDNTCCAp6gAwIBAgIQVHza1pdi+1GZ9HXncIkkSzANBgkqhkiG9w0BAQUFADBxMQswCQYDVQQGEwJDTjESMBAGA1UECBMJQ0hPTkdRSU5HMTYwNAYDVQQKEy1DaG9uZ3FpbmcgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IENlbnRlciBDTy5MVEQxFjAUBgNVBAMTDUNob25ncWluZyAgQ0EwHhcN" & _
"MTEwNDA2MDIzNzMxWhcNMTIwNDA1MDIzNzMxWjAgMQswCQYDVQQGEwJDTjERMA8GA1UEAxMIdGVzdDIwMTEwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBAObTbOvYsy3jSasKzOo2UhzwzkB91bDNxZPHkTvpklEu0ATZE9FD2K5eZ8EMwJd2qQzLDn2RkiTSdguoHCd8MUfaNXRalpawYwF" & _
"0P1pmTx1UQT52tgx3a/BDfhGSL54/P0lDZXxZZssjsqTRVtUweURutrmykURXBlGzrj3UPbPAgMBAAGjggEdMIIBGTBpBggrBgEFBQcBAQRdMFswMAYIKwYBBQUHMAKGJGh0dHA6Ly93d3cuZGZ6eGNhLmNvbS9DQS9jYWlzc3VlLmh0bTAnBggrBgEFBQcwAYYbaHR0cDovL3d3dy5kZnp4Y2Eu" & _
"Y29tOjIwNDQzMB8GA1UdIwQYMBaAFHrTXfuu1AQNTkXzgSsiItPP/FIWMAkGA1UdEwQCMAAwLwYDVR0fBCgwJjAkoCKgIIYeaHR0cDovL3d3dy5kZnp4Y2EuY29tL2NybDguY3JsMAsGA1UdDwQEAwIGwDAdBgNVHQ4EFgQUr8BpW18I7zRSKKHZGbF42BVWktIwIwYFKlYVAQEEGgwYRVpDQTU3" & _
"NUA1MDE0MDIwSkowMTIzNDU2MA0GCSqGSIb3DQEBBQUAA4GBABAlzIJXUNlFRz+adKhcrzdsy0tewwY5FHKqU5wl7f91v3VyQJakOwMgmNiECRRiBF198Z1/fx+ureBB5MA5EZINAiwkDS9EvV1lMiO+mESna0/0jkmQdWDGpwhfO7wrQFibK/0BYvTEfcqzqZZQ7ePsboBWhJaQvk2geYnUmJfQ"

ls_ret = EZCAWebTools.WF_VerifySigneData("dadfsf", SignData, cert)
'VD_VerifySigneData(inData, SignData, signMethodType, signdatatype, signCert, singType)
MsgBox ls_ret


End Sub


Private Sub Command10_Click()
    Dim ls_ret As String
    ls_ret = EZCAWebTools.WF_SignData("123", "ǩ������ԭ��", "EZCA@5014990setup")
    MsgBox ls_ret
End Sub

Private Sub Command11_Click()
Dim ls_ret As String

ls_ret = EZCAWebTools.WF_GetCertDN("EZCA@5014990setup")
MsgBox ls_ret
End Sub

Private Sub Command12_Click()
Dim ls_ret As String

ls_ret = EZCAWebTools.WF_GetCertSN("EZCA@5014990setup")
MsgBox ls_ret
End Sub

Private Sub Command13_Click()
Dim ls_ret As String

ls_ret = EZCAWebTools.WF_GetPublicCert("EZCA@5014990setup")
MsgBox ls_ret
End Sub

Private Sub Command14_Click()
EZCAWebTools.WF_ShowCert ("EZCA@5014990setup")
End Sub

Private Sub Command15_Click()
Dim ls_ret As String
Dim cert As String
cert = EZCAWebTools.WF_GetPublicCert("EZCA@5014990setup")
ls_ret = EZCAWebTools.WF_VerifySigneData(cert, EZCAWebTools.EncodeBase64String("ǩ������ԭ��"), "q2TY+PevVCc5fcflELSRgtZFVrXiIiTUQtcvkvidcwDRlrISC9MVSaLb/U/HshKzqggVshmAlfjOI8CJfQSzFTOZdKL97iMf7cba87cWzwoE+s+oXh9KPDz+SgFqTjnB9nrwGM8Osw31l49Z/E8Hgv+fhbWPGVG0NT02qURdyKA=")
MsgBox ls_ret
End Sub

Private Sub Command2_Click()
Dim ls_ret As String
ls_ret = EZCAClient.SOF_GetVersion()

'MsgBox "test", vbOKOnly, "", "", ls_ret, ""
MsgBox "�ؼ��汾:" & ls_ret

ls_ret = EZCAClient.SOF_CheckSupport()
If ls_ret = 0 Then
    MsgBox "�ؼ�֧��::" & "֧��"
Else
    MsgBox "�ؼ�֧��::" & "��֧��"
End If

End Sub

Private Sub Command3_Click()
Dim ls_ret As String
Dim users() As String
ls_ret = EZCAClient.SOF_GetUserList()
MsgBox "֤���û�::" & ls_ret

If ls_ret = "" Then
    MsgBox "�����գӣ£ˣţ٣�"
Else
    If InStr(ls_ret, "&&&") > 0 Then
    '����û���
        
        users = Split(ls_ret, "&&&")
        'ѡ���û�
        Cert_User = frmSelectUser.ShowMe(ls_ret)
        MsgBox Cert_User
    Else
        Cert_User = ls_ret
    End If
        
End If
End Sub

Private Sub Command4_Click()
Dim cert_id As String
Dim lstr() As String
Dim ls_ret As String
Cert_User = EZCAClient.SOF_GetUserList()
If Cert_User = "" Then
    MsgBox "�����KEY������"
    Exit Sub
End If

lstr = Split(Cert_User, "||")
cert_id = lstr(1)
'MsgBox cert_id

ls_ret = EZCAClient.SOF_SignData(cert_id, "ǩ��ԭ��")

MsgBox "P1��ʽ��ǩ�����ݣ�" & ls_ret
ls_ret = EZCAClient.SOF_SignDataByP7(cert_id, "ǩ��ԭ��")
MsgBox "P7��ʽ��ǩ�����ݣ�" & ls_ret


End Sub

Private Sub Command5_Click()
Dim li_ret As String
'ezcawebtools.Of_SetUrl (Text3.Text)
'ezcawebtools.Of_init

'sign_data'������bussinessId ҵ��Ψһ��ʶ��sdata ��ǩ��ԭ�ģ�certID �û���ʶ

li_ret = EZCAWebTools.WF_SignData("11223234234q32r", "�й�����Ҫ��˪", "EZCA@5014990setup")
If Len(li_ret) > 6 Then
    MsgBox "����ɹ���ǩ��ֵΪ��" & li_ret
Else
    MsgBox "����ʧ�ܣ������룺" & li_ret
End If

End Sub

Private Sub Command6_Click()
''BusinessId ҵ��ϵͳ��ΨһID��,SignCertSerialNoǩ��֤�����к�,signDataTypeǩ��ֵ���� ��0:P7,1:P1��
li_ret = EZCAWebTools.WF_VerifyServerSigneData("11223234234q32r", "10d755cfb9ba15023852aabd88887ee9", "1")
If li_ret <> "20001" And li_ret <> "20002" Then
    MsgBox "��֤�ɹ�������ֵ��" & li_ret
Else
    MsgBox "��֤ʧ�ܣ������룺" & li_ret
End If
End Sub

Private Sub Command7_Click()
EZCAWebTools.Of_SetUrl (Text3.Text)
Dim ls_ret As String

ls_ret = EZCAWebTools.WF_GetUserCert("EZCA@5014990setup", "02")
MsgBox ls_ret
End Sub

Private Sub Command8_Click()
Dim ls_ret As String
ls_ret = EZCAWebTools.WF_SignService("23232323232", "1")
MsgBox ls_ret
End Sub



Private Sub Command9_Click()
Dim li_ret As Integer
Dim li_count, i As Long
Dim ts, te As Date
EZCAWebTools.Of_SetUrl (Text3.Text)

Command9.Enabled = False
li_count = Int(Text4.Text)

ts = Now()
Text1.Text = "��ʼʱ�䣺" & str(ts)

Do While i < li_count
    li_ret = EZCAWebTools.Of_test("11223234234q32r232" & str(li_count), EZCAWebTools.EncodeBase64String("343434�й�����Ҫ��˪"), "0", "0", "232323")
    If li_ret <> 0 Then
        MsgBox "�������:" + str(li_ret)
    End If
    Text2.Text = str(i)
    i = i + 1
    Form1.Refresh
    
    'Sleep (Int(100000 * Rnd) + 90000)
Loop
te = Now()
Text5.Text = DateDiff("s", ts, te) & "��"
Text1.Text = Text1.Text & vbCrLf & "����ʱ�䣺" & str(te)


Command9.Enabled = True

End Sub

Private Sub Form_Load()
    Set EZCAClient = CreateObject("SANITATIONSYSTEMCLIENT.EZCASanitationSystemClient")
End Sub

Private Sub ֤��ʹ����_Click()
Dim ls_ret As String

ls_ret = EZCAWebTools.WF_GetCertOwner("EZCA@5014990setup")
MsgBox ls_ret
End Sub
