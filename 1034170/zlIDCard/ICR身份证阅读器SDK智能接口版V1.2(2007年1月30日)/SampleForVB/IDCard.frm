VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ܽӿ����֤�Ķ���ʾ����"
   ClientHeight    =   8115
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IDCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10965
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdReadIINSNDN 
      Caption         =   "��оƬ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9240
      Picture         =   "IDCard.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4830
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "��Ƭ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   26
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton NewAddCmd 
      Caption         =   "����סַ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9240
      Picture         =   "IDCard.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3570
      Width           =   1095
   End
   Begin VB.CommandButton RdCmd 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9240
      Picture         =   "IDCard.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2310
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   19
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   1800
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   7620
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7057
            MinWidth        =   7057
            Key             =   "pg_status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "����RS-232C��"
            TextSave        =   "����RS-232C��"
            Key             =   "status"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13759
            MinWidth        =   13759
            Text            =   "��������һ�о���   ��Ȩ����  2005��12��"
            TextSave        =   "��������һ�о���   ��Ȩ����  2005��12��"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton EndCmd 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9240
      Picture         =   "IDCard.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6090
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   2048
      InputLen        =   1
      ParityReplace   =   0
      BaudRate        =   115200
      EOFEnable       =   -1  'True
      InputMode       =   1
   End
   Begin VB.Label IINSNDN 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   25
      Top             =   7080
      Width           =   5055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "оƬ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   1200
      TabIndex        =   24
      Top             =   7200
      Width           =   765
   End
   Begin VB.Label NewAdd 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   2400
      TabIndex        =   22
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����סַ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   1080
      TabIndex        =   21
      Top             =   6000
      Width           =   1020
   End
   Begin VB.Label ValidDate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label reg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   4800
      Width           =   5175
   End
   Begin VB.Label IDN 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label address 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   15
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label born 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label nation 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label namet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label sex 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ч����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ǩ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "������ݺ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "סַ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Photo 
      Height          =   1965
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   6015
      Left            =   600
      Top             =   960
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "���֤�Ķ���ʾ����"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Const ReadState = "����״̬"
Const DebugState = "����״̬"

Const OpenPortError = "�򿪴���ʧ��!"
Const TimeOutError = "ͨѶ��ʱ!"
Const RecError = "����ʧ��!"
Const XpError = "��Ƭ�������!"
Const FileExtError = "wlt�ļ���׺����!"
Const FileOpenError = "wlt�ļ��򿪴���!"
Const FileFormatError = "wlt�ļ���ʽ����!"
Const JmError = "���δ��Ȩ!"
Const CardError = "����֤����!"
Const UnknowError = "δ֪����!"

Const Swipe = "��ſ�..."
Const ReadOK = "�����ɹ�!�����һ�ſ�..."
Const ReadError = "����ʧ��!�����·ſ�..."
Const NewAddError = "������סַʧ��!"
Const IINSNDNError = "��оƬ��ʧ��!"
Const Reading = "���ڶ���..."

Const strPathName = "C:"

Dim bcc, TimeOutFlag As Byte
Dim state As Boolean
Dim OutByte() As Byte
Dim RecCount, i, j As Long
Dim ReadResult, PortNum As Integer
Dim ComPort, ReadMode, tmp As String
Dim nametmp, sextmp, nationtmp, borntmp, addresstmp, IDNtmp, regtmp, datetmp As String
Dim RecTmp(), RecByte() As String

Dim iFlag As Integer

Dim bFlagUSB As Boolean

'�����Ƿ���������
Private Sub Check1_Click()

    If Check1.Value = 0 Then
        RdCmd.Enabled = True
        NewAddCmd.Enabled = True
        If Not bFlagUSB Then cmdReadIINSNDN.Enabled = True
        Timer2.Enabled = False          '�ض�ʱ��2
    Else
        RdCmd.Enabled = False
        NewAddCmd.Enabled = False
        cmdReadIINSNDN.Enabled = False
        Timer2.Enabled = True           '����ʱ��2
    End If
    
End Sub

Private Sub cmdReadIINSNDN_Click()
    Dim rd(28) As Byte
    
    IINSNDN.Caption = ""
    MainForm.StatusBar1.Panels("pg_status").Text = Reading
    
    ans = AuthenticateExt()    '����֤
   If ans = 1 Then
        ans = Read_Content(4)          '������סַ
        
          Select Case ans
           Case 1                      '�����ɹ�
                Open "IINSNDN.bin" For Binary Access Read As #1
                    For i = 1 To 28
                        Get #1, i, rd(i - 1)
                    Next i
                    
                    '����
                    tmp = ""
                    For i = 0 To 27
                        tmp1 = Hex((rd(i)))
                        tmp1 = String(2 - Len(tmp1), "0") + tmp1
                        tmp = tmp + " " + tmp1
                    Next i
                Close #1
                
                IINSNDN.Caption = LTrim(tmp)
                MainForm.StatusBar1.Panels("pg_status").Text = ReadOK
            Case -5                     '���δ��Ȩ
               MainForm.StatusBar1.Panels("pg_status").Text = JmError
           Case Else                   '����ʧ��
                MainForm.StatusBar1.Panels("pg_status").Text = IINSNDNError
        End Select
    Else
        MainForm.StatusBar1.Panels("pg_status").Text = "�����·ſ���"
    End If
End Sub

'��ʼ
Private Sub Form_Load()

    If Check1.Value = 0 Then
        RdCmd.Enabled = True
        NewAddCmd.Enabled = True
        cmdReadIINSNDN.Enabled = True
    Else
        RdCmd.Enabled = False
        NewAddCmd.Enabled = False
        cmdReadIINSNDN.Enabled = False
    End If
    
    bFlagUSB = False
    
    PortNum = 2
    ans = InitCommExt '(PortNum)         '������
    If ans = 0 Then
        PortNum = 1001
        ans = InitComm(PortNum)         '��USB��
        If ans <> 1 Then
            ret = MsgBox("�򿪶˿�ʧ�ܣ�", , "����")
            End
        End If
    End If
    
    If ans >= 1001 Then
        MainForm.StatusBar1.Panels("status").Text = "����USB��"
         bFlagUSB = True
    End If
        
      
    Dim strSAMID As String '* 37
    
    strSAMID = GetSAMID()
    Dim s
    s = Split(strSAMID, "-", -1, 1)
    If UBound(s) > 3 Then MainForm.Caption = MainForm.Caption + "(" + "��Ȩ��: " + s(2) + "-" + s(3) + ") "
    
    Timer1.Interval = 2000          '2s
    Timer2.Interval = 300           '300ms
'    Timer2.Enabled = True           '����ʱ��2
    
    If Check1.Value = 0 Then
        RdCmd.Enabled = True
        NewAddCmd.Enabled = True
        cmdReadIINSNDN.Enabled = True
        Timer2.Enabled = False          '�ض�ʱ��2
    Else
        RdCmd.Enabled = False
        NewAddCmd.Enabled = False
        cmdReadIINSNDN.Enabled = False
        Timer2.Enabled = True           '����ʱ��2
    End If
    
    
    ReadResult = 0
    iFlag = 0
    state = True                    'ˢ��״̬
    
End Sub

'��ʱ��1�¼�
Private Sub Timer1_Timer()

    TimeOutFlag = 1
    
End Sub

'��ʱ��2�¼�(����֤/������)
Private Sub Timer2_Timer()
    
    '��ʾ״̬
    If state = True Then         '����״̬
        Select Case ReadResult
            Case 0
               MainForm.StatusBar1.Panels("pg_status").Text = Swipe
            Case 1
               MainForm.StatusBar1.Panels("pg_status").Text = ReadOK
            Case -1                     '��Ƭ�������
               Call Display(strPathName)
               Photo.Picture = LoadPicture()
               MainForm.StatusBar1.Panels("pg_status").Text = XpError
            Case -2               '�����
                MainForm.StatusBar1.Panels("pg_status").Text = FileExtError
            Case -3               '�����
                MainForm.StatusBar1.Panels("pg_status").Text = FileOpenError
            Case -4               '�����
                MainForm.StatusBar1.Panels("pg_status").Text = FileFormatError
            Case -5                     '���δ��Ȩ
               MainForm.StatusBar1.Panels("pg_status").Text = JmError
            Case Else                   '����ʧ��
               MainForm.StatusBar1.Panels("pg_status").Text = ReadError
        End Select
    End If
    
    ans = Authenticate()    '����֤
    
    '����֤�ɹ�
    If ans = 1 Then
        namet.Caption = ""
        sex.Caption = ""
        nation.Caption = ""
        born.Caption = ""
        address.Caption = ""
        IDN.Caption = ""
        reg.Caption = ""
        ValidDate.Caption = ""
        NewAdd.Caption = ""
        IINSNDN.Caption = ""
        Photo.Picture = LoadPicture()
        MainForm.StatusBar1.Panels("pg_status").Text = Reading
          
        If Check2.Value = 0 Then
'            ans = Read_Content(2)         'ֻ��������Ϣ,��������Ƭ����
            ans = Read_Content_Path(strPathName, 2)
        Else
            ans = Read_Content_Path(strPathName, 1)
'            ans = Read_Content(1)         '��������Ϣ
        End If
        
        Select Case ans
           Case 1                      '�����ɹ�
              ReadResult = 1
              Call Display(strPathName) 'App.Path)
           Case -1                     '��Ƭ�������
              Call Display(App.Path)
              Photo.Picture = LoadPicture()
              ReadResult = -1
           Case -2                     'wlt�ļ���׺����
              ReadResult = -2
           Case -3                     'wlt�ļ��򿪴���
              ReadResult = -3
           Case -4                     'wlt�ļ���ʽ����
              ReadResult = -4
           Case -5                     '���δ��Ȩ
              ReadResult = -5
        '   Case -12                    '·����̫��
        '      ReadResult = -12
           Case Else                   '����ʧ��
              ReadResult = 2
        End Select
    End If
      
End Sub

'������ť
Private Sub RdCmd_Click()

    namet.Caption = ""
    sex.Caption = ""
    nation.Caption = ""
    born.Caption = ""
    address.Caption = ""
    IDN.Caption = ""
    reg.Caption = ""
    ValidDate.Caption = ""
    NewAdd.Caption = ""
    IINSNDN.Caption = ""
    Photo.Picture = LoadPicture()
    MainForm.StatusBar1.Panels("pg_status").Text = Reading
    
    ans = AuthenticateExt()    '����֤
    
    If Check2.Value = 1 Then
        ans = Read_Content(4)        '��������Ϣ
    Else
        ans = Read_Content(2)       'ֻ��������Ϣ,��������Ƭ����
    End If
     
    Select Case ans
       Case 1                          '�����ɹ�
          Call Display(App.Path)
          MainForm.StatusBar1.Panels("pg_status").Text = ReadOK
    
        Case -1                     '��Ƭ�������
           Call Display(App.Path)
           Photo.Picture = LoadPicture()
           MainForm.StatusBar1.Panels("pg_status").Text = XpError
        Case -2                     'wlt�ļ���׺����
            MainForm.StatusBar1.Panels("pg_status").Text = FileExtError
        Case -3                     'wlt�ļ��򿪴���
            MainForm.StatusBar1.Panels("pg_status").Text = FileOpenError
        Case -4                     'wlt�ļ���ʽ����
            MainForm.StatusBar1.Panels("pg_status").Text = FileFormatError
        Case -5                     '���δ��Ȩ
           MainForm.StatusBar1.Panels("pg_status").Text = JmError
        Case Else                   '����ʧ��
           MainForm.StatusBar1.Panels("pg_status").Text = ReadError
    End Select

End Sub

'��ʾ��Ϣ
Private Sub Display(ByRef strFilePath As String)
    Dim tmp1 As Byte
    Dim tmp2 As Byte
    Dim rddata As String
    
    Open strFilePath & "\wz.txt" For Binary As #1
        Do While Not EOF(1)   ' ����ļ�β��
            Get #1, , tmp1
            Get #1, , tmp2
    
            rddata = rddata + ChrW(tmp2 * CLng(256) + tmp1)
        Loop
    Close #1
    
    '����
    nametmp = Mid(rddata, 1, 15)
    
    '�Ա�
    sextmp = Mid(rddata, 16, 1)
    
    '����
    nationtmp = Mid(rddata, 17, 2)
    
    '��������
    borntmp = Mid(rddata, 19, 8)
    
    'סַ
    addresstmp = Mid(rddata, 27, 35)
    
    '������ݺ���
    IDNtmp = Mid(rddata, 62, 18)
    
    'ǩ������
    regtmp = Mid(rddata, 80, 15)
    
    '��Ч����
    ValidDatetmp = Mid(rddata, 95, 16)
    
    '��ʾ������Ϣ
    namet.Caption = nametmp
    
    '�Ա�
    Select Case sextmp
        Case "0"
            sex.Caption = "δ֪"
        Case "1"
            sex.Caption = "��"
        Case "2"
            sex.Caption = "Ů"
        Case Else
            sex.Caption = "δ˵��"
    End Select

    '����
    Dim nationtmp1 As String
    ans = GetNation(nationtmp, nationtmp1)
    nation.Caption = nationtmp1
    
    born.Caption = Mid(borntmp, 1, 4) + "��" + Mid(borntmp, 5, 2) + "��" + Mid(borntmp, 7, 2) + "��"
    address.Caption = addresstmp
    IDN.Caption = IDNtmp
    reg.Caption = regtmp
    If Mid(ValidDatetmp, 9, 2) = "����" Then
        ValidDate.Caption = Mid(ValidDatetmp, 1, 4) + "." + Mid(ValidDatetmp, 5, 2) + "." + Mid(ValidDatetmp, 7, 2) + "-" + Mid(ValidDatetmp, 9, 2)
    Else
        ValidDate.Caption = Mid(ValidDatetmp, 1, 4) + "." + Mid(ValidDatetmp, 5, 2) + "." + Mid(ValidDatetmp, 7, 2) + "-" + Mid(ValidDatetmp, 9, 4) + "." + Mid(ValidDatetmp, 13, 2) + "." + Mid(ValidDatetmp, 15, 2)
    End If
    
    '��ʾ��Ƭ
    If Check2.Value = 1 Then Photo.Picture = LoadPicture(strFilePath & "\zp.bmp")

End Sub

'���������
Public Function GetNation(ByVal strNationcode As String, ByRef strNation As String)
    Dim strNationArray As Variant
    
    strNationArray = Array("��", "�ɹ�", "��", "��", "ά���", "��", "��", "׳", "����", "����", _
                        "��", "��", "��", "��", "����", "����", "������", "��", "��", "����", _
                        "��", "�", "��ɽ", "����", "ˮ", "����", "����", "����", "�¶�����", "��", _
                        "���Ӷ�", "����", "Ǽ", "����", "����", "ë��", "����", "����", "����", "����", _
                        "������", "ŭ", "���α��", "����˹", "���¿�", "�°�", "����", "ԣ��", "��", "������", _
                        "����", "���״�", "����", "�Ű�", "���", "��ŵ")
    
    If Trim(strNationcode) <> "" Then
        If ((CByte(Trim(strNationcode)) - 1) >= 0) And ((CByte(Trim(strNationcode)) - 1) <= 55) Then
            strNation = strNationArray(CByte(Trim(strNationcode)) - 1)
        Else
            strNation = "����"
        End If
    End If
    
End Function

'������סַ��ť
Private Sub NewAddCmd_Click()

    NewAdd.Caption = ""
    MainForm.StatusBar1.Panels("pg_status").Text = Reading
    
    ans = Authenticate()    '����֤
    ans = Read_Content(3)          '������סַ
    
    Select Case ans
       Case 1                      '�����ɹ�
            Dim tmp1 As Byte
            Dim tmp2 As Byte
            Dim addresstmp As String
            
            Open "newadd.txt" For Binary As #1
                Do While Not EOF(1)   ' ����ļ�β��
                    Get #1, , tmp1
                    Get #1, , tmp2
            
                    addresstmp = addresstmp + ChrW(tmp2 * CLng(256) + tmp1)
                Loop
            Close #1
            
            NewAdd.Caption = addresstmp
            MainForm.StatusBar1.Panels("pg_status").Text = ReadOK
        Case -5                     '���δ��Ȩ
           MainForm.StatusBar1.Panels("pg_status").Text = JmError
       Case Else                   '����ʧ��
            MainForm.StatusBar1.Panels("pg_status").Text = NewAddError
    End Select

End Sub

'�˳���ť
Private Sub EndCmd_Click()
   
   ret = CloseComm                  '�ش���
   End

End Sub

'�رմ���
Private Sub Form_Unload(Cancel As Integer)
   
   ret = CloseComm                  '�ش���
   End

End Sub

