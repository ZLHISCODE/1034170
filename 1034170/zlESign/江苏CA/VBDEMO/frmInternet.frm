VERSION 5.00
Object = "{B3AA1750-BC62-4F7D-A8EA-C3940949399F}#1.0#0"; "gSeal.ocx"
Begin VB.Form frmInternet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEMO"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10410
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmd֤��ӵ���� 
      Caption         =   "֤��ӵ����"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdӡ��ǩ�� 
      Caption         =   "ӡ��ǩ��"
      Height          =   615
      Left            =   5760
      TabIndex        =   12
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdӡ�� 
      Caption         =   "��ȡӡ��"
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   2760
      Width           =   1695
   End
   Begin VB.PictureBox picǩ�� 
      Height          =   1455
      Left            =   2880
      ScaleHeight     =   1395
      ScaleWidth      =   2595
      TabIndex        =   10
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmd��ȡӡ�����ݲ����� 
      Caption         =   "��ȡӡ������"
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdӡ�±������� 
      Caption         =   "��ȡӡ�±�������"
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdǩ�� 
      Caption         =   "ǩ��"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmd��������֤ 
      Caption         =   "��������֤"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmd��֤���� 
      Caption         =   "������֤"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmd֤������ 
      Caption         =   "֤������"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0��ȡ֤��ID"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1��ȡkey��Ϣ"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmd��ʼ�� 
      Caption         =   "��ʼ��"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin GSEALLib.GSeal gSeal 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1296
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private JSCA_Sign   As CACltCoreLib.CltCore
Private mstrCertCode As String
Private mstrSignData As String     'ǩ������
Private mServer As MSSOAPLib30.SoapClient30
Private Const mCertID As String = "1@5013SF0500101198712265717"   '֤��ID JSCA_Sign.SOF_SelectCert(3)
Private Const mType֤��ӵ���� As String = "0x00000017"

'Private Sub Command1_Click()
'    Dim strURL As String
'    strURL = "wwww.baidu.com"
'    Text1.Text = netConn.OpenURL(strURL, icString)
'End Sub


Public Sub ICExecute(Optional ByVal strURL As String, Optional ByVal strOperation As String, _
    Optional ByVal strInputData As String, Optional ByVal strInputHdrs As String)
    
    inetConn.Execute strURL, strOperation, strInputData, strInputHdrs
End Sub

Private Sub cmdRequest_Click()
    Dim strURL As String
'    StrUrl = "www.baidu.com"
'    ICExecute StrUrl, "POST", "xx=223,222=23", ""
    strURL = "http://192.168.1.113:8080/test/myServlet"
    txtShow.Text = PostData(strURL, "name=xsd", responseText)
'    txtShow.Text = GetData(StrUrl, ResponseText)
End Sub



Private Sub cmd_Click()
    
End Sub

Private Sub cmd0_Click()
    Dim strMsg As String
    strMsg = JSCA_Sign.SOF_SelectCert(3)
    MsgBox strMsg
    Debug.Print strMsg
End Sub

Private Sub cmd1_Click()
    Dim strMsg As String
    Set JSCA_Sign = CreateObject("CACltCore.CltCore")
    strMsg = "����1(�û���):" & JSCA_Sign.SOF_GetUserInfo(mCertID, 1) & vbCrLf & _
             "����2��֤��ʵ��Ψһ��ʶ��:" & JSCA_Sign.SOF_GetUserInfo(mCertID, 2) & vbCrLf & _
             "����3�����ţ�:" & JSCA_Sign.SOF_GetUserInfo(mCertID, 3) & vbCrLf & _
             "����4���䷢��DN��:" & JSCA_Sign.SOF_GetUserInfo(mCertID, 4) & vbCrLf & _
             "����9��CA���ͣ�:" & JSCA_Sign.SOF_GetUserInfo(mCertID, 9) & vbCrLf & _
             "����22���û�֤��UniqueID��OID��:" & JSCA_Sign.SOF_GetUserInfo(mCertID, 22)
             
'����1(�û���):C=CN, S=����ʡ, L=�Ͼ���, O=����ʡ������, OU=CA����, E=ywj@126.com, CN=��ΰ��
'����2��֤��ʵ��Ψһ��ʶ��:1@5013SF0500101198712265717
'����3 (����): JSCA
'����4���䷢��DN��:C=CN, S=����ʡ, L=�Ͼ���, O=����ʡ��������֤����֤�����������ι�˾, OU=JSCA, CN=JSCA_CA
'����9 (CA����): JSCA_CA
'����22���û�֤��UniqueID��OID��:1.2.86.21.1.1
    Debug.Print strMsg
    MsgBox strMsg
End Sub

Private Sub cmd��ʼ��_Click()
    Set JSCA_Sign = CreateObject("CACltCore.CltCore")
End Sub

Private Sub inetConn_StateChanged(ByVal State As Integer)
    Dim varData As Variant
    Dim strData As String
    Dim blnDo As Boolean
    
    Select Case State
    
    Case icError
        MsgBox "�������:��" & inetConn.ResponseCode & "����������:" & inetConn.ResponseInfo
    Case icResponseCompleted
        'ȡ�õ�һ��
        blnDo = False
        varData = inetConn.GetChunk(1024, icString)
        
        DoEvents
        Do While Not blnDo
            strData = strData & varData
            DoEvents
            varData = inetConn.GetChunk(1024, icString)
            If Len(varData) = 0 Then
                blnDo = True
            End If
        Loop
        mstrReturn = strData
        inetConn.Cancel 'ȡ������
    End Select
    
End Sub

Private Sub cmd��������֤_Click()
    Dim intTimes As Integer
    Dim strCertId As String
    Dim strCertCode As String
    Dim strPassWord As String
    Dim strRet As String
    Dim strURL As String
    Dim strPara As String
    Dim strSource As String
    Dim strSignData As String
    Dim objSoapClient As New MSSOAPLib30.SoapClient30
'     CreateObject("MSSOAP.SoapClient30")
    Set JSCA_Sign = CreateObject("CACltCore.CltCore")
    strCertId = JSCA_Sign.SOF_SelectCert(3)
    strCertCode = JSCA_Sign.SOF_ExportUserCert(strCertId)
    strSource = "����"
    strCertCode = "MIIDxjCCAy+gAwIBAgIMMm7+Rlcg7856sc8pMA0GCSqGSIb3DQEBBQUAMIGOMQ0wCwYDVQQGHgQAQwBOMQ8wDQYDVQQIHgZsX4LPdwExDzANBgNVBAceBlNXT" & _
"qxeAjEvMC0GA1UECh4mbF+Cz3cBdTVbUFVGUqGLwU5mi6SLwU4tX8NnCZZQjSNO+1FsU/gxETAPBgNVBAseCABKAFMAQwBBMRcwFQYDVQQDHg4ASgBTAEMAQQBfAEMAQTAeFw0xNDA3MjkwMT" & _
"U3NTNaFw0xNTA3MjkwMTU3NTNaMIGVMQswCQYDVQQGDAJDTjESMBAGA1UECAwJ5rGf6IuP55yBMRIwEAYDVQQHDAnljZfkuqzluIIxGzAZBgNVBAoMEuaxn+iLj+ecgeWNq+eUn+WOhTERMA8" & _
"GA1UECwwIQ0HkuK3lv4MxGjAYBgkqhkiG9w0BCQEWC3l3akAxMjYuY29tMRIwEAYDVQQDDAnkvZnkvJ/oioIwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBALjjf3AQrsZtRaeVuGetAbeH" & _
"NNdhMMYDOZP7GDom5WS+fMBz2F0gBQBU2mur9jKuNADz03RKCbSjXiHu9eIgJHOnPWgkIoQJtWwhT4525r8GhsQ/J47sepB0YBrWvREY56eDGGH2DlBCirkJYvQOGkRvwHeNncpjQhiKdyZrR/kLAgMBAAGjggEe" & _
"MIIBGjAMBgNVHRMEBTADAQEAMAsGA1UdDwQEAwIAwDAkBgUqVhUBAQQbMUA1MDEzU0YwNTAwMTAxMTk4NzEyMjY1NzE3MB8GA1UdIwQYMBaAFFbAyBFUVTYGSn3tJlDoiL23o3oJMFAGCCsGAQUFBwEBBEQwQjBAB" & _
"ggrBgEFBQcwAoY0aHR0cDovL2NlcnQucHVibGlzaC5zZXJ2ZXI6ODg4MC9kb3dubG9hZC9KU0NBX0NBLmNlcjBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY2VydC5wdWJsaXNoLnNlcnZlcjo" & _
"4ODgwL2Rvd25sb2FkL0pTQ0FfQ0EuY3JsMB0GA1UdDgQWBBRFzWRJ9ZyN3gSXSWacjmzPNB6JlzANBgkqhkiG9w0BAQUFAAOBgQA5KSM2jdNwkvAk0bFRvh0oNQoyc/" & _
"umTsdTPFUsghptUmSPez/vZL7FYSz+DZn87kkgZSCbCbJEY8zkhBj+G8SEEHuANo3ArnldtPCMtWTLkxRw2L6RMq8vjwaWVIyRnoPX4nUFIJB5+LRdrR6u3trJUv3Y6lfAivAMTkGPNPegVA=="

    
    
    strSource = "10342   10343   ��¼��Դ    Ů  28��    0   1   2014-08-20 09:03:00 ��Ī���ֽ���(��������ŵ��) 0.25g*20��   ��Ī���ֽ���    2   2       ÿ�����    2   1   ��  10-16   4   0   ������  2014-08-20 09:03:00" & _
"10343       ��¼��Դ    Ů  28��    0   1   2014-08-20 09:03:00 �ڷ�                    ÿ�����    2   1   ��  10-16   0   0   ������  2014-08-20 09:03:00"
'strSignData = JSCA_Sign.SOF_SignData(strCertId, strSource)
strSignData = "Q0Aj/FNjfFq+eKEciufD7vnsfQpUadgqV+kSLnKEPKeT8syXcNEW4OS8iu1caLhJxKeK4t3VflbbE5ShJYM083nWi8xKqDLBjzyRKszu8xnuk7Yu7SIP9n3mw6rhf7IK+BMiOcFy8f+9E8W/5x89MRvbQsFV09Z1VuQbedNDfWo="
    strURL = "http://202.102.85.153:8080/HealthWebService.asmx?WSDL"
    Set objSoapClient = New SoapClient30
    Call objSoapClient.MSSoapInit(strURL)
    strRet = objSoapClient.VerifySignedData(strCertCode, strSource, strSignData)   '����0  ��֤�ɹ�
    
'    strPara = "Base64EncodeCert=" & strCertCode & "&InData=" & strSource & "&SignValue=" & strSignData
'    strRet = PostData(strURL, strPara, responseText)
    
    MsgBox "��������֤��" & strRet
    Debug.Print "��������֤��" & strRet
    
End Sub

Private Sub cmd��ȡӡ�����ݲ�����_Click()
    Dim strMsg As String
    Dim strFile As String
    
    strFile = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & mCertID & ".gif"
    strMsg = gSeal.JSCAGetSealPath(strFile)
    
    MsgBox strMsg
    Debug.Print strFile
End Sub

Private Sub cmdǩ��_Click()
    Dim strCertId As String
    Dim strSource As String
    
    Set JSCA_Sign = CreateObject("CACltCore.CltCore")
    strCertId = JSCA_Sign.SOF_SelectCert(3)
    strSource = "����"
    mstrSignData = JSCA_Sign.SOF_SignData(strCertId, strSource)
    MsgBox "ǩ������:" & mstrSignData
    Debug.Print mstrSignData
End Sub

Private Sub cmd��֤����_Click()
    Dim intTimes As Long
    Dim strCertId As String
    Dim strPassWord As String
    Dim strErr As String
    
    
    Set JSCA_Sign = CreateObject("CACltCore.CltCore")
    JSCA_Sign.IsShowError = 0
    strCertId = JSCA_Sign.SOF_SelectCert(3)
    strPassWord = "123456"

    intTimes = JSCA_Sign.SOF_Login(strCertId, strPassWord)   'У��ɹ����� 0,ʧ�ܷ���ʣ�����,-1 ����

    MsgBox "У�����뷵��ֵ��" & intTimes
    Debug.Print "У�����뷵��ֵ��" & intTimes
End Sub

Private Sub cmdӡ��_Click()
    Dim strMsg As String
    
    strMsg = gSeal.GetCert
    MsgBox "ӡ��Base64����:" & strMsg
    Debug.Print "ӡ��Base64����:" & strMsg
End Sub

Private Sub cmdӡ�±�������_Click()
    Dim strMsg As String
    Dim strBase64 As String
    Dim strFile As String  '�ļ�λ��
    
    strBase64 = gSeal.JSCAGetSeal
    strMsg = SaveBase64ToFile("gif", mCertID, strBase64)
    If strFile <> "" Then
        picǩ��.Picture = LoadPicture(strFile)
    End If
    MsgBox strMsg
    Debug.Print strMsg
End Sub

Private Sub cmdӡ��ǩ��_Click()
    Dim strMsg As String
    strMsg = gSeal.gtSignString("����")
    
    MsgBox "ӡ��ǩ��ֵ:" & strMsg
    Debug.Print "ӡ��ǩ��ֵ:" & strMsg
End Sub

Private Sub cmd֤������_Click()
    Dim strMsg As String
    
    Set JSCA_Sign = CreateObject("CACltCore.CltCore")
    strMsg = JSCA_Sign.SOF_ExportUserCert(mCertID)
    mstrCertCode = strMsg
    MsgBox strMsg
    Debug.Print "֤������BASE64��"; strMsg
End Sub

Private Sub Command1_Click()

    gSeal.JSCAGetSealPath ("")
End Sub

Private Sub cmd֤��ӵ����_Click()
    Dim strMsg As String
    Dim strCode As String
    
    Set JSCA_Sign = CreateObject("CACltCore.CltCore")
    strCode = JSCA_Sign.SOF_ExportUserCert(mCertID)
    strMsg = JSCA_Sign.SOF_GetCertInfo(strCode, 21)
    MsgBox strMsg
End Sub
