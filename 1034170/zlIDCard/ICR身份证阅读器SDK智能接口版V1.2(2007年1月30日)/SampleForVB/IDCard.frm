VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "智能接口身份证阅读演示程序"
   ClientHeight    =   8115
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "宋体"
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdReadIINSNDN 
      Caption         =   "读芯片号"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "相片解码"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "最新住址"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "读 卡"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "连续读卡"
      BeginProperty Font 
         Name            =   "宋体"
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
            Text            =   "连接RS-232C口"
            TextSave        =   "连接RS-232C口"
            Key             =   "status"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13759
            MinWidth        =   13759
            Text            =   "公安部第一研究所   版权所有  2005年12月"
            TextSave        =   "公安部第一研究所   版权所有  2005年12月"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Caption         =   "退 出"
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "芯片号"
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "最新住址"
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "有效期限"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "签发机关"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "公民身份号码"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "住址"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "出生"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "民族"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "身份证阅读演示程序"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
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
 
Const ReadState = "读卡状态"
Const DebugState = "调试状态"

Const OpenPortError = "打开串口失败!"
Const TimeOutError = "通讯超时!"
Const RecError = "操作失败!"
Const XpError = "相片解码错误!"
Const FileExtError = "wlt文件后缀错误!"
Const FileOpenError = "wlt文件打开错误!"
Const FileFormatError = "wlt文件格式错误!"
Const JmError = "软件未授权!"
Const CardError = "卡认证错误!"
Const UnknowError = "未知错误!"

Const Swipe = "请放卡..."
Const ReadOK = "读卡成功!请放下一张卡..."
Const ReadError = "读卡失败!请重新放卡..."
Const NewAddError = "读最新住址失败!"
Const IINSNDNError = "读芯片号失败!"
Const Reading = "正在读卡..."

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

'设置是否连续读卡
Private Sub Check1_Click()

    If Check1.Value = 0 Then
        RdCmd.Enabled = True
        NewAddCmd.Enabled = True
        If Not bFlagUSB Then cmdReadIINSNDN.Enabled = True
        Timer2.Enabled = False          '关定时器2
    Else
        RdCmd.Enabled = False
        NewAddCmd.Enabled = False
        cmdReadIINSNDN.Enabled = False
        Timer2.Enabled = True           '开定时器2
    End If
    
End Sub

Private Sub cmdReadIINSNDN_Click()
    Dim rd(28) As Byte
    
    IINSNDN.Caption = ""
    MainForm.StatusBar1.Panels("pg_status").Text = Reading
    
    ans = AuthenticateExt()    '卡认证
   If ans = 1 Then
        ans = Read_Content(4)          '读最新住址
        
          Select Case ans
           Case 1                      '读卡成功
                Open "IINSNDN.bin" For Binary Access Read As #1
                    For i = 1 To 28
                        Get #1, i, rd(i - 1)
                    Next i
                    
                    '译码
                    tmp = ""
                    For i = 0 To 27
                        tmp1 = Hex((rd(i)))
                        tmp1 = String(2 - Len(tmp1), "0") + tmp1
                        tmp = tmp + " " + tmp1
                    Next i
                Close #1
                
                IINSNDN.Caption = LTrim(tmp)
                MainForm.StatusBar1.Panels("pg_status").Text = ReadOK
            Case -5                     '软件未授权
               MainForm.StatusBar1.Panels("pg_status").Text = JmError
           Case Else                   '读卡失败
                MainForm.StatusBar1.Panels("pg_status").Text = IINSNDNError
        End Select
    Else
        MainForm.StatusBar1.Panels("pg_status").Text = "请重新放卡！"
    End If
End Sub

'开始
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
    ans = InitCommExt '(PortNum)         '开串口
    If ans = 0 Then
        PortNum = 1001
        ans = InitComm(PortNum)         '开USB口
        If ans <> 1 Then
            ret = MsgBox("打开端口失败！", , "错误")
            End
        End If
    End If
    
    If ans >= 1001 Then
        MainForm.StatusBar1.Panels("status").Text = "连接USB口"
         bFlagUSB = True
    End If
        
      
    Dim strSAMID As String '* 37
    
    strSAMID = GetSAMID()
    Dim s
    s = Split(strSAMID, "-", -1, 1)
    If UBound(s) > 3 Then MainForm.Caption = MainForm.Caption + "(" + "授权号: " + s(2) + "-" + s(3) + ") "
    
    Timer1.Interval = 2000          '2s
    Timer2.Interval = 300           '300ms
'    Timer2.Enabled = True           '开定时器2
    
    If Check1.Value = 0 Then
        RdCmd.Enabled = True
        NewAddCmd.Enabled = True
        cmdReadIINSNDN.Enabled = True
        Timer2.Enabled = False          '关定时器2
    Else
        RdCmd.Enabled = False
        NewAddCmd.Enabled = False
        cmdReadIINSNDN.Enabled = False
        Timer2.Enabled = True           '开定时器2
    End If
    
    
    ReadResult = 0
    iFlag = 0
    state = True                    '刷卡状态
    
End Sub

'定时器1事件
Private Sub Timer1_Timer()

    TimeOutFlag = 1
    
End Sub

'定时器2事件(卡认证/读数据)
Private Sub Timer2_Timer()
    
    '显示状态
    If state = True Then         '读卡状态
        Select Case ReadResult
            Case 0
               MainForm.StatusBar1.Panels("pg_status").Text = Swipe
            Case 1
               MainForm.StatusBar1.Panels("pg_status").Text = ReadOK
            Case -1                     '相片解码错误
               Call Display(strPathName)
               Photo.Picture = LoadPicture()
               MainForm.StatusBar1.Panels("pg_status").Text = XpError
            Case -2               '解码错
                MainForm.StatusBar1.Panels("pg_status").Text = FileExtError
            Case -3               '解码错
                MainForm.StatusBar1.Panels("pg_status").Text = FileOpenError
            Case -4               '解码错
                MainForm.StatusBar1.Panels("pg_status").Text = FileFormatError
            Case -5                     '软件未授权
               MainForm.StatusBar1.Panels("pg_status").Text = JmError
            Case Else                   '读卡失败
               MainForm.StatusBar1.Panels("pg_status").Text = ReadError
        End Select
    End If
    
    ans = Authenticate()    '卡认证
    
    '卡认证成功
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
'            ans = Read_Content(2)         '只读文字信息,不进行相片解码
            ans = Read_Content_Path(strPathName, 2)
        Else
            ans = Read_Content_Path(strPathName, 1)
'            ans = Read_Content(1)         '读基本信息
        End If
        
        Select Case ans
           Case 1                      '读卡成功
              ReadResult = 1
              Call Display(strPathName) 'App.Path)
           Case -1                     '相片解码错误
              Call Display(App.Path)
              Photo.Picture = LoadPicture()
              ReadResult = -1
           Case -2                     'wlt文件后缀错误
              ReadResult = -2
           Case -3                     'wlt文件打开错误
              ReadResult = -3
           Case -4                     'wlt文件格式错误
              ReadResult = -4
           Case -5                     '软件未授权
              ReadResult = -5
        '   Case -12                    '路径名太长
        '      ReadResult = -12
           Case Else                   '读卡失败
              ReadResult = 2
        End Select
    End If
      
End Sub

'读卡按钮
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
    
    ans = AuthenticateExt()    '卡认证
    
    If Check2.Value = 1 Then
        ans = Read_Content(4)        '读基本信息
    Else
        ans = Read_Content(2)       '只读文字信息,不进行相片解码
    End If
     
    Select Case ans
       Case 1                          '读卡成功
          Call Display(App.Path)
          MainForm.StatusBar1.Panels("pg_status").Text = ReadOK
    
        Case -1                     '相片解码错误
           Call Display(App.Path)
           Photo.Picture = LoadPicture()
           MainForm.StatusBar1.Panels("pg_status").Text = XpError
        Case -2                     'wlt文件后缀错误
            MainForm.StatusBar1.Panels("pg_status").Text = FileExtError
        Case -3                     'wlt文件打开错误
            MainForm.StatusBar1.Panels("pg_status").Text = FileOpenError
        Case -4                     'wlt文件格式错误
            MainForm.StatusBar1.Panels("pg_status").Text = FileFormatError
        Case -5                     '软件未授权
           MainForm.StatusBar1.Panels("pg_status").Text = JmError
        Case Else                   '读卡失败
           MainForm.StatusBar1.Panels("pg_status").Text = ReadError
    End Select

End Sub

'显示信息
Private Sub Display(ByRef strFilePath As String)
    Dim tmp1 As Byte
    Dim tmp2 As Byte
    Dim rddata As String
    
    Open strFilePath & "\wz.txt" For Binary As #1
        Do While Not EOF(1)   ' 检查文件尾。
            Get #1, , tmp1
            Get #1, , tmp2
    
            rddata = rddata + ChrW(tmp2 * CLng(256) + tmp1)
        Loop
    Close #1
    
    '姓名
    nametmp = Mid(rddata, 1, 15)
    
    '性别
    sextmp = Mid(rddata, 16, 1)
    
    '民族
    nationtmp = Mid(rddata, 17, 2)
    
    '出生日期
    borntmp = Mid(rddata, 19, 8)
    
    '住址
    addresstmp = Mid(rddata, 27, 35)
    
    '公民身份号码
    IDNtmp = Mid(rddata, 62, 18)
    
    '签发机关
    regtmp = Mid(rddata, 80, 15)
    
    '有效期限
    ValidDatetmp = Mid(rddata, 95, 16)
    
    '显示文字信息
    namet.Caption = nametmp
    
    '性别
    Select Case sextmp
        Case "0"
            sex.Caption = "未知"
        Case "1"
            sex.Caption = "男"
        Case "2"
            sex.Caption = "女"
        Case Else
            sex.Caption = "未说明"
    End Select

    '民族
    Dim nationtmp1 As String
    ans = GetNation(nationtmp, nationtmp1)
    nation.Caption = nationtmp1
    
    born.Caption = Mid(borntmp, 1, 4) + "年" + Mid(borntmp, 5, 2) + "月" + Mid(borntmp, 7, 2) + "日"
    address.Caption = addresstmp
    IDN.Caption = IDNtmp
    reg.Caption = regtmp
    If Mid(ValidDatetmp, 9, 2) = "长期" Then
        ValidDate.Caption = Mid(ValidDatetmp, 1, 4) + "." + Mid(ValidDatetmp, 5, 2) + "." + Mid(ValidDatetmp, 7, 2) + "-" + Mid(ValidDatetmp, 9, 2)
    Else
        ValidDate.Caption = Mid(ValidDatetmp, 1, 4) + "." + Mid(ValidDatetmp, 5, 2) + "." + Mid(ValidDatetmp, 7, 2) + "-" + Mid(ValidDatetmp, 9, 4) + "." + Mid(ValidDatetmp, 13, 2) + "." + Mid(ValidDatetmp, 15, 2)
    End If
    
    '显示照片
    If Check2.Value = 1 Then Photo.Picture = LoadPicture(strFilePath & "\zp.bmp")

End Sub

'民族代码查表
Public Function GetNation(ByVal strNationcode As String, ByRef strNation As String)
    Dim strNationArray As Variant
    
    strNationArray = Array("汉", "蒙古", "回", "藏", "维吾尔", "苗", "彝", "壮", "布依", "朝鲜", _
                        "满", "侗", "瑶", "白", "土家", "哈尼", "哈萨克", "傣", "黎", "傈僳", _
                        "佤", "畲", "高山", "拉祜", "水", "东乡", "纳西", "景颇", "柯尔克孜", "土", _
                        "达斡尔", "仫佬", "羌", "布朗", "撒拉", "毛南", "仡佬", "锡伯", "阿昌", "普米", _
                        "塔吉克", "怒", "乌孜别克", "俄罗斯", "鄂温克", "德昂", "保安", "裕固", "京", "塔塔尔", _
                        "独龙", "鄂伦春", "赫哲", "门巴", "珞巴", "基诺")
    
    If Trim(strNationcode) <> "" Then
        If ((CByte(Trim(strNationcode)) - 1) >= 0) And ((CByte(Trim(strNationcode)) - 1) <= 55) Then
            strNation = strNationArray(CByte(Trim(strNationcode)) - 1)
        Else
            strNation = "其他"
        End If
    End If
    
End Function

'读最新住址按钮
Private Sub NewAddCmd_Click()

    NewAdd.Caption = ""
    MainForm.StatusBar1.Panels("pg_status").Text = Reading
    
    ans = Authenticate()    '卡认证
    ans = Read_Content(3)          '读最新住址
    
    Select Case ans
       Case 1                      '读卡成功
            Dim tmp1 As Byte
            Dim tmp2 As Byte
            Dim addresstmp As String
            
            Open "newadd.txt" For Binary As #1
                Do While Not EOF(1)   ' 检查文件尾。
                    Get #1, , tmp1
                    Get #1, , tmp2
            
                    addresstmp = addresstmp + ChrW(tmp2 * CLng(256) + tmp1)
                Loop
            Close #1
            
            NewAdd.Caption = addresstmp
            MainForm.StatusBar1.Panels("pg_status").Text = ReadOK
        Case -5                     '软件未授权
           MainForm.StatusBar1.Panels("pg_status").Text = JmError
       Case Else                   '读卡失败
            MainForm.StatusBar1.Panels("pg_status").Text = NewAddError
    End Select

End Sub

'退出按钮
Private Sub EndCmd_Click()
   
   ret = CloseComm                  '关串口
   End

End Sub

'关闭窗口
Private Sub Form_Unload(Cancel As Integer)
   
   ret = CloseComm                  '关串口
   End

End Sub

