VERSION 5.00
Object = "{EDA0698C-EB1A-46F9-BAA9-3687D671FF68}#1.0#0"; "JITSEC~1.OCX"
Object = "{E2CBDEB6-97C0-476A-BF58-7292B4C1BF98}#1.0#0"; "IMGCON~1.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ǩ�¶�-������ݲ���������ǩ��"
   ClientHeight    =   5910
   ClientLeft      =   3990
   ClientTop       =   2670
   ClientWidth     =   7665
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7665
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   7455
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   7
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "Form1.frx":0000
         Top             =   3360
         Width           =   5415
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Text            =   "����"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Text            =   "Ů"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Text            =   "24"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   6
         Text            =   "��"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   5
         Text            =   "�����Ը�ð"
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "Form1.frx":0069
         Top             =   3000
         Width           =   2535
      End
      Begin JITSECURITYTOOLLib.JITSecurityTool dsnJit 
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   0
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   3
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "Form1.frx":0087
         Top             =   1440
         Width           =   5415
      End
      Begin IMGCONVERTLib.Imgconvert dsnImg 
         Height          =   375
         Left            =   4710
         TabIndex        =   28
         Top             =   375
         Visible         =   0   'False
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   0
      End
      Begin VB.Image pic 
         Height          =   945
         Left            =   4890
         Top             =   2310
         Width           =   2055
      End
      Begin VB.Label lbl 
         Caption         =   "�� ����"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "�� ��"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "�� �䣺"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "֢ ״��"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "�� ����"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "�� �ϣ�"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   15
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "�� ����"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   14
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "�� ע��"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   13
         Top             =   3360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ��"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   4730
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ǩ��"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   4730
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��֤"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   4730
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "������"
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   7455
      Begin VB.TextBox SignedDatatxt 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   27
         Text            =   "Form1.frx":00FA
         Top             =   600
         Width           =   6375
      End
      Begin VB.TextBox ErrCode 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox psw 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   23
         Text            =   "111111"
         Top             =   165
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "ǩ��ֵ��"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "������룺"
         Height          =   255
         Index           =   9
         Left            =   2160
         TabIndex        =   24
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "֤�����룺"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   22
         Top             =   210
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjFile As New FileSystemObject

Private mobjJit As Object
Private mobjImg As Object

Private mstrCert As String
Private mstrSign As String

Private nRet As Integer
Private ClientSignCert As String
Private plain As String

Private Sub Form_Load()
    Dim strFile As String
    
    'ʹ�ö�̬�ؼ�
    Set mobjJit = CreateObject("JITSECURITYTOOL.JITSecurityToolCtrl.1")
    Set mobjImg = CreateObject("IMGCONVERT.ImgconvertCtrl.1")
    
    strFile = GetSetting("ZLSOFT", "ע����Ϣ", "gstrAviPath", "")
    If strFile <> "" And mobjFile.FileExists(strFile & "\server.cer") And mobjFile.FileExists(strFile & "\server.pfx") Then
        'ʹ���ļ���ʽ֤��
        mstrCert = "FILE://" & Replace(strFile & "\server.cer", "\", "\\")
        mstrSign = "FILE://" & Replace(strFile & "\server.pfx", "\", "\\")
        psw.Text = "11111111" 'ȱʡ����
    Else
        'ʹ�����ܿ�֤��
        mstrCert = "USBCSP://.2CER"
        mstrSign = "USBCSP://.2CER"
        psw.Text = "111111" 'ȱʡ����
    End If
End Sub

Private Sub Command1_Click()
    nRet = dsnJit.initcontrol() 'ִ�нӿڲ���ǰ�����ʼ������
    ErrCode.Text = nRet
    If (nRet = 0 Or nRet = 50) Then
       MsgBox ("��ʼ�������ɹ�!")
    Else
       MsgBox ("��ʼ������ʧ��!")
       Exit Sub
    End If
   
    nRet = dsnJit.readcert(mstrCert, 2, psw.Text, mstrCert, 2, psw.Text)
    ErrCode.Text = nRet
    If (nRet = 0) Then
        MsgBox ("��ʼ���ͻ���֤��ɹ���")
    Else
        MsgBox ("��ʼ���ͻ���֤��ʧ�ܣ�")
        Exit Sub
    End If
   
    ClientSignCert = dsnJit.getsigncert()
End Sub

Private Sub Command2_Click()
    Dim strBMP As String, strGIF As String
    Dim i As Long
    
    plain = ""
    For i = 0 To 7
        plain = plain & txt(i).Text
    Next
    
    'LenB(StrConv(plain, vbFromUnicode))
    nRet = dsnJit.signdata(mstrSign, psw.Text, plain, Len(plain)) 'ʹ��USB֤�����ǩ��
    ErrCode.Text = nRet
    If (nRet = 0) Then
        MsgBox ("ǩ���ɹ���")
    Else
        MsgBox ("ǩ��ʧ�ܣ�")
        Exit Sub
    End If
    SignedDatatxt.Text = dsnJit.getconten()               '//������ĵ�Base64��
    
    '-------------------------------------------
    strBMP = mobjFile.GetSpecialFolder(TemporaryFolder).Path & "\" & mobjFile.GetTempName
    strGIF = mobjFile.GetSpecialFolder(TemporaryFolder).Path & "\" & mobjFile.GetTempName
    
    'picindex=KEY��ͼ�����кţ���KEYֻ֧�ִ������ͼ�£�0-ͼ��һ��1-ͼ�¶�
    'rgbflag=ͼ��ǰ��ɫ��0-��ɫ��1-��ɫ��2-��ɫ
    nRet = dsnJit.ShowSinglePic(psw.Text, 0, strBMP, 0) '�ļ�֤�鷽ʽ�����е���
    ErrCode.Text = nRet
    If (nRet = 0) Then
        MsgBox ("��ȡͼ�³ɹ���")
    Else
        MsgBox ("��ȡͼ��ʧ�ܣ�")
        GoTo DelFile
    End If
    
    nRet = dsnImg.Bmp2TransparentGif(strBMP, strGIF, 0)
    If (nRet = 0) Then
        MsgBox ("ͼ��ת���ɹ���")
    Else
        MsgBox ("ͼ��ת��ʧ�ܣ�")
        GoTo DelFile
    End If
    pic.Picture = LoadPicture(strGIF)
    
DelFile:
    If mobjFile.FileExists(strBMP) Then mobjFile.DeleteFile strBMP, True
    If mobjFile.FileExists(strGIF) Then mobjFile.DeleteFile strGIF, True
End Sub

Private Sub Command3_Click()
    Dim i As Long
    
    plain = ""
    For i = 0 To 7
       plain = plain & txt(i).Text
    Next
    
    'LenB(StrConv(plain, vbFromUnicode))
    nRet = dsnJit.verifySign(ClientSignCert, plain, Len(plain), SignedDatatxt.Text)
    ErrCode.Text = nRet
    If (nRet <> 0) Then
       MsgBox ("�ĵ����Ķ����������¸��£�")
       Exit Sub
    Else
        MsgBox "��֤�ɹ���"
    End If
End Sub
