VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PaneTwo 
   Appearance      =   0  'Flat
   BackColor       =   &H0080C0FF&
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   LockControls    =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   9825
   Begin MSComCtl2.MonthView MView 
      Height          =   2220
      Left            =   9120
      TabIndex        =   115
      Top             =   3840
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      Appearance      =   1
      StartOfWeek     =   116326402
      CurrentDate     =   41981
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1020
      TabIndex        =   29
      Tag             =   "137,236"
      Top             =   1236
      Width           =   4305
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType2 
      Height          =   270
      Index           =   1
      Left            =   2265
      TabIndex        =   67
      Tag             =   "215,394"
      Top             =   3840
      Width           =   675
      _ExtentX        =   25744
      _ExtentY        =   476
      Caption         =   "慢性"
   End
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6225
      TabIndex        =   30
      Tag             =   "455,236"
      Top             =   1236
      Width           =   1815
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType1 
      Height          =   270
      Index           =   4
      Left            =   6960
      TabIndex        =   65
      Tag             =   "541,371"
      Top             =   3483
      Width           =   2400
      _ExtentX        =   205740
      _ExtentY        =   476
      Caption         =   "阳性检测结果（献血员）"
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType1 
      Height          =   270
      Index           =   3
      Left            =   5655
      TabIndex        =   64
      Tag             =   "452,371"
      Top             =   3480
      Width           =   1455
      _ExtentX        =   204073
      _ExtentY        =   476
      Caption         =   "病原携带者、"
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType1 
      Height          =   270
      Index           =   2
      Left            =   4080
      TabIndex        =   63
      Tag             =   "339,371"
      Top             =   3483
      Width           =   1755
      _ExtentX        =   204602
      _ExtentY        =   476
      Caption         =   "实验室确诊病例、"
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType1 
      Height          =   270
      Index           =   1
      Left            =   2550
      TabIndex        =   62
      Tag             =   "238,371"
      Top             =   3483
      Width           =   1545
      _ExtentX        =   204232
      _ExtentY        =   476
      Caption         =   "临床诊断病例、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   17
      Left            =   8985
      TabIndex        =   60
      Tag             =   "677,350"
      Top             =   3090
      Width           =   810
      _ExtentX        =   202089
      _ExtentY        =   476
      Caption         =   "不详"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   16
      Left            =   7785
      TabIndex        =   59
      Tag             =   "594,350"
      Top             =   3090
      Width           =   1260
      _ExtentX        =   202883
      _ExtentY        =   476
      Caption         =   "其他（ ）、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   15
      Left            =   6510
      TabIndex        =   58
      Tag             =   "506,350"
      Top             =   3090
      Width           =   1350
      _ExtentX        =   203041
      _ExtentY        =   476
      Caption         =   "家务及待业、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   14
      Left            =   5415
      TabIndex        =   57
      Tag             =   "431,350"
      Top             =   3090
      Width           =   1200
      _ExtentX        =   202777
      _ExtentY        =   476
      Caption         =   "离退人员、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   13
      Left            =   4335
      TabIndex        =   56
      Tag             =   "356,350"
      Top             =   3090
      Width           =   1200
      _ExtentX        =   202777
      _ExtentY        =   476
      Caption         =   "干部职员、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   12
      Left            =   2985
      TabIndex        =   55
      Tag             =   "281,350"
      Top             =   3090
      Width           =   1455
      _ExtentX        =   203226
      _ExtentY        =   476
      Caption         =   "渔(船)民、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   11
      Left            =   2250
      TabIndex        =   54
      Tag             =   "230,350"
      Top             =   3090
      Width           =   810
      _ExtentX        =   202089
      _ExtentY        =   476
      Caption         =   "牧民、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   10
      Left            =   1515
      TabIndex        =   53
      Tag             =   "179,350"
      Top             =   3090
      Width           =   810
      _ExtentX        =   194892
      _ExtentY        =   476
      Caption         =   "农民、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   9
      Left            =   825
      TabIndex        =   52
      Tag             =   "128,350"
      Top             =   3090
      Width           =   810
      _ExtentX        =   194892
      _ExtentY        =   476
      Caption         =   "民工、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   7
      Left            =   8760
      TabIndex        =   50
      Tag             =   "654,326"
      Top             =   2640
      Width           =   1260
      _ExtentX        =   202883
      _ExtentY        =   476
      Caption         =   "医务 人员"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   6
      Left            =   7635
      TabIndex        =   49
      Tag             =   "579,326"
      Top             =   2640
      Width           =   1200
      _ExtentX        =   202777
      _ExtentY        =   476
      Caption         =   "商业服务、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   5
      Left            =   6360
      TabIndex        =   48
      Tag             =   "492,326"
      Top             =   2640
      Width           =   1350
      _ExtentX        =   203041
      _ExtentY        =   476
      Caption         =   "餐饮食品业、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   4
      Left            =   4905
      TabIndex        =   47
      Tag             =   "391,326"
      Top             =   2640
      Width           =   1530
      _ExtentX        =   203359
      _ExtentY        =   476
      Caption         =   "保育员及保姆、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   3
      Left            =   4170
      TabIndex        =   46
      Tag             =   "340,326"
      Top             =   2640
      Width           =   810
      _ExtentX        =   202089
      _ExtentY        =   476
      Caption         =   "教师、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   2
      Left            =   2355
      TabIndex        =   45
      Tag             =   "227,326"
      Top             =   2640
      Width           =   1890
      _ExtentX        =   203994
      _ExtentY        =   476
      Caption         =   "学生(大中小学)、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   1
      Left            =   1230
      TabIndex        =   44
      Tag             =   "152,326"
      Top             =   2644
      Width           =   1200
      _ExtentX        =   202777
      _ExtentY        =   476
      Caption         =   "散居儿童、"
   End
   Begin VB.TextBox txtDiagnose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   3900
      TabIndex        =   74
      Tag             =   "314,445"
      Top             =   4590
      Width           =   525
   End
   Begin VB.TextBox txtDeath 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   6075
      TabIndex        =   75
      Tag             =   "554,445"
      Top             =   4590
      Width           =   1095
   End
   Begin VB.TextBox txtDeath 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   7545
      TabIndex        =   76
      Tag             =   "620,445"
      Top             =   4590
      Width           =   615
   End
   Begin VB.TextBox txtDeath 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   8355
      TabIndex        =   77
      Tag             =   "660,445"
      Top             =   4590
      Width           =   525
   End
   Begin VB.TextBox txtDiagnose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   71
      Tag             =   "170,445"
      Top             =   4590
      Width           =   1095
   End
   Begin VB.TextBox txtDiagnose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2370
      TabIndex        =   72
      Tag             =   "235,445"
      Top             =   4590
      Width           =   615
   End
   Begin VB.TextBox txtDiagnose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   3165
      TabIndex        =   73
      Tag             =   "275,445"
      Top             =   4590
      Width           =   525
   End
   Begin VB.TextBox txtAttack 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   68
      Tag             =   "170,422"
      Top             =   4232
      Width           =   1095
   End
   Begin VB.TextBox txtAttack 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2370
      TabIndex        =   69
      Tag             =   "235,422"
      Top             =   4232
      Width           =   615
   End
   Begin VB.TextBox txtAttack 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   3165
      TabIndex        =   70
      Tag             =   "275,422"
      Top             =   4232
      Width           =   525
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1530
      TabIndex        =   37
      Tag             =   "179,281"
      Top             =   1940
      Width           =   690
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2400
      TabIndex        =   38
      Tag             =   "239,281"
      Top             =   1940
      Width           =   885
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   5025
      TabIndex        =   40
      Tag             =   "407,281"
      Top             =   1940
      Width           =   885
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   7035
      TabIndex        =   41
      Tag             =   "551,281"
      Top             =   1940
      Width           =   915
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   8145
      TabIndex        =   42
      Tag             =   "617,281"
      Top             =   1940
      Width           =   690
   End
   Begin zlDisReportCard.uCheckNorm ucAge 
      Height          =   270
      Index           =   0
      Left            =   8115
      TabIndex        =   26
      Tag             =   "599,210"
      Top             =   810
      Width           =   465
      _ExtentX        =   202327
      _ExtentY        =   476
      Caption         =   "岁"
   End
   Begin VB.TextBox txtParentName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3645
      TabIndex        =   1
      Tag             =   "317,164"
      Top             =   180
      Width           =   1455
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Tag             =   "137,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1381
      TabIndex        =   3
      Tag             =   "157,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   1682
      TabIndex        =   4
      Tag             =   "177,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   1983
      TabIndex        =   5
      Tag             =   "197,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   2284
      TabIndex        =   6
      Tag             =   "217,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   2585
      TabIndex        =   7
      Tag             =   "237,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   2886
      TabIndex        =   8
      Tag             =   "257,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   3187
      TabIndex        =   9
      Tag             =   "277,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   3488
      TabIndex        =   10
      Tag             =   "297,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   9
      Left            =   3789
      TabIndex        =   11
      Tag             =   "317,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   10
      Left            =   4090
      TabIndex        =   12
      Tag             =   "337,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   11
      Left            =   4391
      TabIndex        =   13
      Tag             =   "357,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   12
      Left            =   4692
      TabIndex        =   14
      Tag             =   "377,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   4993
      TabIndex        =   15
      Tag             =   "397,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   14
      Left            =   5294
      TabIndex        =   16
      Tag             =   "417,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   15
      Left            =   5595
      TabIndex        =   17
      Tag             =   "437,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   16
      Left            =   5896
      TabIndex        =   18
      Tag             =   "457,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   17
      Left            =   6210
      TabIndex        =   19
      Tag             =   "477,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtBirth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   22
      Tag             =   "170,212"
      Top             =   854
      Width           =   1095
   End
   Begin VB.TextBox txtBirth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2355
      TabIndex        =   23
      Tag             =   "235,212"
      Top             =   854
      Width           =   615
   End
   Begin VB.TextBox txtBirth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   3150
      TabIndex        =   24
      Tag             =   "275,212"
      Top             =   854
      Width           =   525
   End
   Begin VB.TextBox txtAge 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6585
      TabIndex        =   25
      Tag             =   "479,212"
      Top             =   854
      Width           =   525
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   735
      TabIndex        =   0
      Tag             =   "119,164"
      Top             =   180
      Width           =   1455
   End
   Begin zlDisReportCard.uCheckNorm ucSex 
      Height          =   270
      Index           =   0
      Left            =   7215
      TabIndex        =   20
      Tag             =   "539,183"
      Top             =   472
      Width           =   570
      _ExtentX        =   202512
      _ExtentY        =   476
      Caption         =   "男"
   End
   Begin zlDisReportCard.uCheckNorm ucSex 
      Height          =   270
      Index           =   1
      Left            =   7800
      TabIndex        =   21
      Tag             =   "583,183"
      Top             =   472
      Width           =   570
      _ExtentX        =   202512
      _ExtentY        =   476
      Caption         =   "女"
   End
   Begin zlDisReportCard.uCheckNorm ucAge 
      Height          =   270
      Index           =   1
      Left            =   8573
      TabIndex        =   27
      Tag             =   "625,210"
      Top             =   810
      Width           =   465
      _ExtentX        =   201057
      _ExtentY        =   476
      Caption         =   "月"
   End
   Begin zlDisReportCard.uCheckNorm ucAge 
      Height          =   270
      Index           =   2
      Left            =   9030
      TabIndex        =   28
      Tag             =   "651,210"
      Top             =   810
      Width           =   555
      _ExtentX        =   131789
      _ExtentY        =   476
      Caption         =   "天)"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   31
      Tag             =   "149,256"
      Top             =   1543
      Width           =   1005
      _ExtentX        =   203279
      _ExtentY        =   476
      Caption         =   "本县区"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   1
      Left            =   2070
      TabIndex        =   32
      Tag             =   "217,256"
      Top             =   1543
      Width           =   1425
      _ExtentX        =   204020
      _ExtentY        =   476
      Caption         =   "本市其他县区"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   2
      Left            =   3630
      TabIndex        =   33
      Tag             =   "321,256"
      Top             =   1543
      Width           =   1500
      _ExtentX        =   204153
      _ExtentY        =   476
      Caption         =   "本省其它地市"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   3
      Left            =   5190
      TabIndex        =   34
      Tag             =   "431,256"
      Top             =   1543
      Width           =   795
      _ExtentX        =   202909
      _ExtentY        =   476
      Caption         =   "外省"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   4
      Left            =   6150
      TabIndex        =   35
      Tag             =   "493,256"
      Top             =   1543
      Width           =   1005
      _ExtentX        =   203279
      _ExtentY        =   476
      Caption         =   "港澳台"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   5
      Left            =   7230
      TabIndex        =   36
      Tag             =   "561,256"
      Top             =   1543
      Width           =   795
      _ExtentX        =   202909
      _ExtentY        =   476
      Caption         =   "外籍"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   43
      Tag             =   "77,326"
      Top             =   2644
      Width           =   1200
      _ExtentX        =   203623
      _ExtentY        =   476
      Caption         =   "幼托儿童、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   8
      Left            =   105
      TabIndex        =   51
      Tag             =   "77,350"
      Top             =   3105
      Width           =   810
      _ExtentX        =   202935
      _ExtentY        =   476
      Caption         =   "工人、"
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType1 
      Height          =   270
      Index           =   0
      Left            =   1500
      TabIndex        =   61
      Tag             =   "163,371"
      Top             =   3483
      Width           =   1200
      _ExtentX        =   203623
      _ExtentY        =   476
      Caption         =   "疑似病例、"
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType2 
      Height          =   270
      Index           =   0
      Left            =   1500
      TabIndex        =   66
      Tag             =   "163,394"
      Top             =   3835
      Width           =   810
      _ExtentX        =   200395
      _ExtentY        =   476
      Caption         =   "急性、"
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   3510
      TabIndex        =   39
      Tag             =   "311,281"
      Top             =   1940
      Width           =   855
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(乙型肝炎、丙型肝炎、血吸虫病填写)"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2940
      TabIndex        =   117
      Tag             =   "255,397"
      Top             =   3885
      Width           =   3060
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8895
      Picture         =   "PaneTwo.ctx":0000
      Top             =   1380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(1)"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1110
      TabIndex        =   116
      Tag             =   "143,375"
      Top             =   3525
      Width           =   270
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "县(区)"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   16
      Left            =   4380
      TabIndex        =   94
      Tag             =   "360,281"
      Top             =   1935
      Width           =   540
   End
   Begin VB.Line Line1 
      Index           =   23
      Tag             =   "653,455,684"
      X1              =   8355
      X2              =   8900
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   22
      Tag             =   "611,455,645"
      X1              =   7530
      X2              =   8145
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   21
      Tag             =   "311,455,330"
      X1              =   3900
      X2              =   4515
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   20
      Tag             =   "269,455,300"
      X1              =   3165
      X2              =   3700
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   19
      Tag             =   "227,455,254"
      X1              =   2370
      X2              =   2985
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   18
      Tag             =   "269,432,300"
      X1              =   3165
      X2              =   3700
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line1 
      Index           =   17
      Tag             =   "227,432,254"
      X1              =   2370
      X2              =   2985
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line1 
      Index           =   16
      Tag             =   "527,455,600"
      X1              =   6075
      X2              =   7250
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   15
      Tag             =   "143,455,216"
      X1              =   1080
      X2              =   2175
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   0
      Tag             =   "143,432,216"
      X1              =   1080
      X2              =   2175
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line1 
      Index           =   7
      Tag             =   "137,245,394"
      X1              =   1020
      X2              =   5310
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "联系电话："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   5340
      TabIndex        =   114
      Tag             =   "396,236"
      Top             =   1236
      Width           =   900
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "工作单位："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   105
      TabIndex        =   113
      Tag             =   "78,236"
      Top             =   1236
      Width           =   900
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "病例分类*："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   21
      Left            =   105
      TabIndex        =   112
      Tag             =   "78,375"
      Top             =   3525
      Width           =   990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(2)"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   22
      Left            =   1110
      TabIndex        =   111
      Tag             =   "143,398"
      Top             =   3880
      Width           =   270
   End
   Begin VB.Line Line1 
      Index           =   11
      Tag             =   "311,292,351"
      X1              =   3465
      X2              =   4365
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label lblDiagnose 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "时"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   4470
      TabIndex        =   110
      Tag             =   "331,445 "
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "死亡日期 ："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   24
      Left            =   5175
      TabIndex        =   109
      Tag             =   "462,445"
      Top             =   4590
      Width           =   990
   End
   Begin VB.Label lblDeath 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   7275
      TabIndex        =   108
      Tag             =   "600,445"
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblDeath 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "月"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   8175
      TabIndex        =   107
      Tag             =   "642,445"
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblDeath 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "日"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   8925
      TabIndex        =   106
      Tag             =   "686,445"
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "诊断日期*："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   23
      Left            =   105
      TabIndex        =   105
      Tag             =   "78,445"
      Top             =   4590
      Width           =   990
   End
   Begin VB.Label lblDiagnose 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   2175
      TabIndex        =   104
      Tag             =   "216,445"
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblDiagnose 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "月"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2985
      TabIndex        =   103
      Tag             =   "258,445"
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblDiagnose 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "日"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   3690
      TabIndex        =   102
      Tag             =   "302,445"
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "发病日期*："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   25
      Left            =   105
      TabIndex        =   101
      Tag             =   "78,422"
      Top             =   4232
      Width           =   990
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   2175
      TabIndex        =   100
      Tag             =   "216,422"
      Top             =   4230
      Width           =   180
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "月"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2985
      TabIndex        =   99
      Tag             =   "258,422 "
      Top             =   4230
      Width           =   180
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "日 (病原携带者填初检日期或就诊时间)"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   3690
      TabIndex        =   98
      Tag             =   "302,422"
      Top             =   4230
      Width           =   3150
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "现住址(详填)*："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   105
      TabIndex        =   97
      Tag             =   "78,281"
      Top             =   1935
      Width           =   1350
   End
   Begin VB.Line Line1 
      Index           =   9
      Tag             =   "179,292,226"
      X1              =   1530
      X2              =   2200
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Line Line1 
      Index           =   10
      Tag             =   "239,292,300"
      X1              =   2400
      X2              =   3300
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "省"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   14
      Left            =   2220
      TabIndex        =   96
      Tag             =   "229,281"
      Top             =   1935
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "市"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   15
      Left            =   3285
      TabIndex        =   95
      Tag             =   "301,281"
      Top             =   1940
      Width           =   180
   End
   Begin VB.Line Line1 
      Index           =   12
      Tag             =   "407,292,470"
      X1              =   5025
      X2              =   5925
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Line Line1 
      Index           =   13
      Tag             =   "551,292,606"
      X1              =   7020
      X2              =   7920
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Line Line1 
      Index           =   14
      Tag             =   "617,292,670"
      X1              =   8145
      X2              =   8850
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "乡(镇、街道)"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   17
      Left            =   5925
      TabIndex        =   93
      Tag             =   "470,281"
      Top             =   1935
      Width           =   1080
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "村"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   18
      Left            =   7950
      TabIndex        =   92
      Tag             =   "606,281"
      Top             =   1940
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(门牌号)"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   19
      Left            =   8835
      TabIndex        =   91
      Tag             =   "672,281"
      Top             =   1935
      Width           =   720
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "患者职业*："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   105
      TabIndex        =   90
      Tag             =   "79,305"
      Top             =   2292
      Width           =   990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "病人属于*："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   12
      Left            =   105
      TabIndex        =   89
      Tag             =   "78,258"
      Top             =   1588
      Width           =   990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "姓名*："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   105
      TabIndex        =   78
      Tag             =   "78,164"
      Top             =   180
      Width           =   630
   End
   Begin VB.Line Line1 
      Index           =   1
      Tag             =   "119,175,210"
      X1              =   705
      X2              =   2265
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(患儿家长姓名："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   2325
      TabIndex        =   79
      Tag             =   "228,164"
      Top             =   180
      Width           =   1350
   End
   Begin VB.Line Line1 
      Index           =   2
      Tag             =   "317,175,390"
      X1              =   3645
      X2              =   5205
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   ")"
      Height          =   180
      Index           =   4
      Left            =   5235
      TabIndex        =   80
      Tag             =   "403,164"
      Top             =   180
      Width           =   90
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "身份证号:"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   105
      TabIndex        =   81
      Tag             =   "78,187"
      Top             =   525
      Width           =   810
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "性别*:"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   6585
      TabIndex        =   82
      Tag             =   "498,187"
      Top             =   510
      Width           =   540
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "出生日期*："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   105
      TabIndex        =   88
      Tag             =   "79,212"
      Top             =   884
      Width           =   990
   End
   Begin VB.Line Line1 
      Index           =   3
      Tag             =   "143,222,214"
      X1              =   1080
      X2              =   2280
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Label lblBirth 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   2175
      TabIndex        =   87
      Tag             =   "216,212"
      Top             =   855
      Width           =   180
   End
   Begin VB.Line Line1 
      Index           =   4
      Tag             =   "227,222,254"
      X1              =   2325
      X2              =   2985
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Label lblBirth 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "月"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2970
      TabIndex        =   86
      Tag             =   "258,212"
      Top             =   855
      Width           =   180
   End
   Begin VB.Line Line1 
      Index           =   5
      Tag             =   "269,222,300"
      X1              =   3135
      X2              =   3735
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Label lblBirth 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "日"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   3735
      TabIndex        =   85
      Tag             =   "301,212"
      Top             =   854
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(如出生日期不详，实足年龄："
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   4125
      TabIndex        =   84
      Tag             =   "318,212"
      Top             =   855
      Width           =   2430
   End
   Begin VB.Line Line1 
      Index           =   6
      Tag             =   "479,222,526"
      X1              =   6525
      X2              =   7125
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年龄单位:"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   7290
      TabIndex        =   83
      Tag             =   "540,212"
      Top             =   855
      Width           =   810
   End
   Begin VB.Line Line1 
      Index           =   8
      Tag             =   "455,245,598"
      X1              =   6225
      X2              =   8115
      Y1              =   1425
      Y2              =   1425
   End
End
Attribute VB_Name = "PaneTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDateType As Byte   '0表示出生日期，1表示发病日期，2表示诊断日期，3表示死亡日期
Private mblnFirst As Boolean
Private mbln身份证必填 As Boolean '身份证信息必填 参数：传染病报告身份证号码必填
Private mcolLoadData As Collection
Public Event ClickPositives(blnSelected As Boolean)  '选择了阳性检测结果时触发

Private Sub lblAttack_Click(Index As Integer)
    mDateType = 1
    Call ShowMView(lblAttack(Index).Left, lblAttack(Index).Top)
End Sub

Public Function HaveChanged() As Boolean
'判断控件显示信息是否发生变化
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    HaveChanged = False
    If mcolLoadData Is Nothing Then
        Set mcolLoadData = New Collection
    End If
    If mcolLoadData.Count <= 0 Then
        Exit Function
    End If
    For Each objCtl In UserControl.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                If objCtl.Text <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
            Case "uCheckNorm"
                If IIf(objCtl.Checked = True, 1, 0) <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
        End Select
        i = i + 1
    Next
End Function

Private Sub SaveLoadData()
'功能：保存控件显示信息
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    Set mcolLoadData = New Collection
    For Each objCtl In UserControl.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                Call mcolLoadData.Add(objCtl.Text, "K" & i)
            Case "uCheckNorm"
                Call mcolLoadData.Add(IIf(objCtl.Checked = True, 1, 0), "K" & i)
        End Select
        i = i + 1
    Next
End Sub

Public Sub SetCaption身份证(ByVal blnHave As Boolean)
    mbln身份证必填 = blnHave
    If mbln身份证必填 Then
        lblReport(5).Caption = "身份证号*:"
    Else
        lblReport(5).Caption = "身份证号:"
    End If
End Sub

Public Sub ClearMe()
    Dim objCtl As Control
    
    On Error GoTo errHand
    For Each objCtl In UserControl.Controls
        Call ClearInfo(objCtl)
    Next
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub PrintTwo()
    Dim objCtl As Control
    For Each objCtl In UserControl.Controls
        Call PrintInfo(objCtl)
    Next
End Sub
Public Sub LoadData(colData As Collection, bytType As Byte, ByVal strChkType As String)
    Dim strTmp As String
    Dim i As Integer
    Dim strInfo() As String
    Dim objCtl As Control
    Dim dteTmp As Date
    
    On Error GoTo errHand
    mblnFirst = True
    If bytType = 1 Then
        txtName.Text = CStr(colData("K3"))
        txtParentName.Text = CStr(colData("K4"))
        
        strTmp = CStr(colData("K5"))
        For i = 1 To 18
            txtIDCard(i - 1).Text = Mid(strTmp, i, 1)
        Next
        
        For Each objCtl In UserControl.Controls
            If TypeName(objCtl) = "uCheckNorm" Then
                strTmp = Trim(objCtl.Caption)
                strTmp = Replace(strTmp, "(", "")
                strTmp = Replace(strTmp, ")", "")
                strTmp = Replace(strTmp, "、", "")
                If objCtl.Name = "ucCheckJob" Then
                    If InStr(strChkType, "14," & strTmp) > 0 And Trim(strTmp) <> "" Then
                        objCtl.Checked = True
                    End If
                ElseIf InStr(strChkType, strTmp) > 0 And Trim(strTmp) <> "" Then
                    objCtl.Checked = True
                    If strTmp = "急性" Then
                        gbytAcute = 0
                    ElseIf strTmp = "慢性" Then
                        gbytAcute = 1
                    End If

                    If InStr("0疑似病例1临床诊断病例2实验室确诊病例3病原携带者4阳性检测结果(献血员)", strTmp) > 0 Then
                        strTmp = InStr("0疑似病例1临床诊断病例2实验室确诊病例3病原携带者4阳性检测结果(献血员)", strTmp)
                        gbytDiseaseType = val(Mid("0疑似病例1临床诊断病例2实验室确诊病例3病原携带者4阳性检测结果(献血员)", val(strTmp) - 1, val(strTmp)))
                    End If
                End If

            End If
        Next
        
        strTmp = CStr(colData("K7"))
        strInfo = Split(strTmp, "-")
        For i = 0 To UBound(strInfo)
            txtBirth(i).Text = IIf(val(strInfo(i)) = 0, "", strInfo(i))
        Next
        
        txtAge.Text = CStr(colData("K8"))
        
        txtAddress.Text = CStr(colData("K10"))
        txtNumber.Text = CStr(colData("K11"))

        strTmp = CStr(colData("K13"))
        strInfo = Split(strTmp, ";")
        For i = 0 To UBound(strInfo) - 1
            txtAddInfo(i).Text = strInfo(i)
        Next
        
        strInfo = Split(CStr(colData("K17")), "-")
        For i = 0 To UBound(strInfo)
            txtAttack(i) = IIf(val(strInfo(i)) = 0, "", strInfo(i))
        Next
        
        strInfo = Split(CStr(colData("K18")), " ")
        If UBound(strInfo) > 0 Then
            txtDiagnose(3) = IIf(val(strInfo(1)) = 0, "", strInfo(1))
        End If
        If UBound(strInfo) >= 0 Then
            strInfo = Split(strInfo(0), "-")
            For i = 0 To UBound(strInfo)
                txtDiagnose(i) = IIf(val(strInfo(i)) = 0, "", strInfo(i))
            Next
        End If
        strInfo = Split(CStr(colData("K19")), "-")
        For i = 0 To UBound(strInfo)
            txtDeath(i) = IIf(val(strInfo(i)) = 0, "", strInfo(i))
        Next
    Else
        txtName.Text = CStr(colData("K0"))
        txtParentName.Text = CStr(colData("KParent"))
        strTmp = CStr(colData("K1"))
        For i = 1 To 18
            txtIDCard(i - 1).Text = Mid(strTmp, i, 1)
        Next
        
        ucSex(IIf(CStr(colData("K2")) = "男", 0, 1)).Checked = True
        
        strTmp = Format(CStr(colData("K3")), "yyyy-mm-dd")
        If strTmp <> "年月日" Then
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtBirth(i).Text = strInfo(i)
            Next
        End If
        
        strTmp = Trim(CStr(colData("K4")))
        i = InStr("大岁月日天", Right(strTmp, 1))
        If i > 1 Then
            i = IIf(i > 4, 4, i)
            txtAge.Text = val(CStr(colData("K4")))
            ucAge(i - 2).Checked = True
        Else
            txtAge.Text = val(CStr(colData("K4")))
            ucAge(0).Checked = True
            If val(txtAge.Text) = 0 Then
                txtAge.Text = ""
                ucAge(0).Checked = False
            End If
        End If
        
        txtAddress.Text = CStr(colData("K5"))
        
        If CStr(colData("K6")) <> "" Then
            strTmp = CStr(colData("K6"))
        ElseIf CStr(colData("K7")) <> "" Then
            strTmp = CStr(colData("K7"))
        Else
            strTmp = CStr(colData("K8"))
        End If
        txtNumber.Text = strTmp
        
        lblReport(13).ToolTipText = CStr(colData("K13"))
        
        strTmp = Format(CStr(colData("K14")), "yyyy-mm-dd")
        If strTmp <> "年月日" Then
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtAttack(i).Text = strInfo(i)
            Next
        End If
        
        strTmp = Format(CStr(colData("K15")), "yyyy-mm-dd-hh")
        If strTmp <> "年月日" Then
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtDiagnose(i).Text = strInfo(i)
            Next
        End If
        
        strTmp = Format(CStr(colData("K17")), "yyyy-mm-dd")
        If strTmp <> "年月日" Then
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtDeath(i).Text = strInfo(i)
            Next
        End If
    End If
    mblnFirst = False
    Call SaveLoadData
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Public Function MakeSaveSql(arrSql() As Variant, colCls As Collection, strFileId As String) As Boolean
    Dim strObjNo As String
    Dim strContent As String
    Dim strReportInfo As String
    
    Dim i As Integer
    Dim strTmp As String
    Dim strTmp1 As String
    
    On Error GoTo errHand
    strObjNo = "3$4$5$6$7$8$9$10$11$12$13$14$15$16$17$18$19"
    
    '姓名、患者父母姓名

    strContent = txtName.Text & "$" & txtParentName.Text & "$"
    
    '身份证号码
    strTmp = ""
    For i = 0 To 17
        strTmp = Trim(strTmp) & Trim(txtIDCard(i).Text)
    Next

    strContent = strContent & strTmp & "$"
    
    '性别
    strTmp = IIf(ucSex(0).Checked = True, 1, IIf(ucSex(1).Checked = True, 2, 3))

    
    strTmp = IIf(ucSex(0).Checked = True, ucSex(0).Caption, IIf(ucSex(1).Checked = True, ucSex(1).Caption, ""))
    strContent = strContent & strTmp & "$"
    
    '出生日期
    strTmp = IIf(Trim(txtBirth(0).Text) = "", 0, Trim(txtBirth(0).Text)) & "-" & IIf(Trim(txtBirth(1).Text) = "", 0, Trim(txtBirth(1).Text)) & "-" & IIf(Trim(txtBirth(2).Text) = "", 0, Trim(txtBirth(2).Text))

    If Trim(strTmp) = "--" Then
        strTmp = ""
    End If
    strContent = strContent & strTmp & "$"
    '年龄

    strContent = strContent & Trim(txtAge.Text) & "$"
    
    strTmp = IIf(ucAge(0).Checked = True, 1, IIf(ucAge(1).Checked = True, 2, IIf(ucAge(2).Checked = True, 3, 4)))

    
    strTmp = IIf(ucAge(0).Checked = True, ucAge(0).Caption, IIf(ucAge(1).Checked = True, ucAge(1).Caption, IIf(ucAge(2).Checked = True, ucAge(2).Caption, "")))
    strContent = strContent & strTmp & "$"
    '工作单位、联系电话

    strContent = strContent & Trim(txtAddress.Text) & "$"
    

    strContent = strContent & Trim(txtNumber.Text) & "$"
    
    '病人属于
    For i = 0 To 5
        If ucFrom(i).Checked = True Then
            strTmp = i + 1
            strTmp1 = ucFrom(i).Caption
            Exit For
        End If
        strTmp = i + 2
        strTmp1 = ""
    Next

    strContent = strContent & strTmp1 & "$"
    
    '现居住
    strTmp = ""
    For i = 0 To 5
        strTmp = Trim(strTmp) & Trim(txtAddInfo(i).Text) & ";"
    Next

    strContent = strContent & strTmp & "$"
    
    '患者职业
    For i = 0 To 17
        If ucCheckJob(i).Checked = True Then
            strTmp = i + 1
            strTmp1 = ucCheckJob(i).Caption
            Exit For
        End If
        strTmp = i + 2
        strTmp1 = ""
    Next

    strContent = strContent & strTmp1 & "$"
    
    '病例分类1
    For i = 0 To 4
        If ucCaseType1(i).Checked = True Then
            strTmp = i + 1
            strTmp1 = ucCaseType1(i).Caption
            Exit For
        End If
        strTmp = i + 2
        strTmp1 = ""
    Next

    strContent = strContent & strTmp1 & "$"
    
    '病例分类2
    strTmp = IIf(ucCaseType2(0).Checked = True, 1, IIf(ucCaseType2(1).Checked = True, 2, 3))

    
    strTmp = IIf(ucCaseType2(0).Checked = True, ucCaseType2(0).Caption, IIf(ucCaseType2(1).Checked = True, ucCaseType2(1).Caption, ""))
    strContent = strContent & strTmp & "$"
    '发病日期
    strTmp = IIf(Trim(txtAttack(0).Text) = "", 0, Trim(txtAttack(0).Text)) & "-" & IIf(Trim(txtAttack(1).Text) = "", 0, Trim(txtAttack(1).Text)) & "-" & IIf(Trim(txtAttack(2).Text) = "", 0, Trim(txtAttack(2).Text))

    
    If Trim(strTmp) = "--" Then
        strTmp = ""
    End If
    strContent = strContent & strTmp & "$"
    
    '诊断日期
    strTmp = IIf(Trim(txtDiagnose(0).Text) = "", 0, Trim(txtDiagnose(0).Text)) & "-" & IIf(Trim(txtDiagnose(1).Text) = "", 0, Trim(txtDiagnose(1).Text)) & "-" & IIf(Trim(txtDiagnose(2).Text) = "", 0, Trim(txtDiagnose(2).Text)) & " " & IIf(Trim(txtDiagnose(3).Text) = "", 0, Trim(txtDiagnose(3).Text))

    strContent = strContent & strTmp & "$"
    
    '死亡日期
    strTmp = IIf(Trim(txtDeath(0).Text) = "", 0, Trim(txtDeath(0).Text)) & "-" & IIf(Trim(txtDeath(1).Text) = "", 0, Trim(txtDeath(1).Text)) & "-" & IIf(Trim(txtDeath(2).Text) = "", 0, Trim(txtDeath(2).Text))

    
    If Trim(strTmp) = "--" Then
        strTmp = ""
    End If
    strContent = strContent & strTmp & "$"
    
    strReportInfo = strObjNo & "|" & strContent
    MakeSaveSql = GetSaveSql(arrSql, colCls, strFileId, strReportInfo)
    Call SaveLoadData
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Function CheckValidity(ByRef strMsg As String) As Boolean
'功能：检查输入的合法性
    Dim strBirth As String      '生日，例如：1991-01-01
    Dim blnIsChild As Boolean   '判断是否为14岁以下的患者
    Dim intAge As Integer
    Dim blnDate As Boolean '判断日期是否输入完整
    Dim i As Integer
    Dim strTmp As String

    On Error GoTo errHand
    CheckValidity = False
    '检查姓名
    If Trim(txtName.Text) = "" Then
        strMsg = strMsg & "<姓名>为必选项，请检查！$"
    End If
    '检查身份证号
    If mbln身份证必填 Then
        For i = 0 To 17
            strTmp = strTmp & txtIDCard(i).Text
        Next
        If Trim(strTmp) = "" Then
            strMsg = strMsg & "<身份证号>为必选项，请检查！$"
        End If
    End If
    
    '检查性别，没有选择时必须选
    If ucSex(0).Checked = False And ucSex(1).Checked = False Then
        strMsg = strMsg & "<性别>为必选项，请检查！$"
    End If
    
    '检查出生日期 允许全空 不允许部份空或数值无效
    txtBirth(0).Text = Trim(txtBirth(0).Text): txtBirth(1).Text = Trim(txtBirth(1).Text): txtBirth(2).Text = Trim(txtBirth(2).Text)
    strBirth = txtBirth(0).Text & "-" & txtBirth(1).Text & "-" & txtBirth(2).Text
    
    If txtBirth(0).Text <> "" Or txtBirth(1).Text <> "" Or txtBirth(2).Text <> "" Then
        If Not IsDate(strBirth) Then
            strMsg = strMsg & "<出生日期>不完整或不是有效日期，请检查！$"
        Else
            If DateDiff("yyyy", strBirth, Now()) <= 14 And Trim(txtParentName.Text) = "" Then
                blnIsChild = True
            End If
        End If
    End If
    
    If Trim(txtBirth(0).Text) = "" And Trim(txtAge.Text) = "" Then
        strMsg = strMsg & "<出生日期>与<年龄>必需填写一项，请检查！$"
    End If
    
    '检查年龄，如果小于14，必须输入父母的名字,必须输入父母的联系电话
    If txtBirth(0).Text = "" Then
        intAge = val(txtAge.Text) * IIf(ucSex(0).Checked, 365, IIf(ucSex(1).Checked, 30, 1))
        If intAge <= (14 * 365) And Trim(txtParentName.Text) = "" Then
            blnIsChild = True
        End If
    End If
    If blnIsChild = True Then
        If Trim(txtNumber.Text) = "" Then
            strMsg = strMsg & "14岁以下患者要求填写<家长联系电话>，请检查！$"
        Else
            strMsg = strMsg & "14岁以下患者要求填写<家长姓名>，请检查！$"
        End If
    End If
    
        '检查病人属于
    For i = 0 To 5
        If ucFrom(i).Checked = True Then
            Exit For
        End If
        If i = 5 Then
            strMsg = strMsg & "<病人属于>为必选项，请检查！$"
        End If
    Next
    
    '检查地址
    For i = 0 To 5
        If Trim(txtAddInfo(i).Text) <> "" Then
            Exit For
        End If
        If i = 5 Then
            strMsg = strMsg & "<现住址>为必选项，请检查！$"
        End If
    Next
    
    '检查职业，必须选择一项
    For i = 0 To 17
        If ucCheckJob(i).Checked = True Then
            Exit For
        End If
        If i = 17 Then
            strMsg = strMsg & "<职业>为必选项，请检查！$"
        End If
    Next
    
    '检查病例分类
    For i = 0 To 4
        If ucCaseType1(i).Checked = True Then
            Exit For
        End If
        If i = 4 Then
            If ucCaseType2(0).Checked = False And ucCaseType2(1).Checked = False Then
                strMsg = strMsg & "<病例分类>为必选项，请检查！$"
            End If
        End If
    Next
    
    '发病日期必须填写
    txtAttack(0).Text = Trim(txtAttack(0).Text): txtAttack(1).Text = Trim(txtAttack(1).Text): txtAttack(2).Text = Trim(txtAttack(2).Text)
    If Not IsDate(txtAttack(0).Text & "-" & txtAttack(1).Text & "-" & txtAttack(2).Text) Then
        strMsg = strMsg & "<发病日期>不完整或不是有效日期，请检查！$"
    End If
    
    '诊断时间必须完整并精确到小时
    txtDiagnose(0).Text = Trim(txtDiagnose(0).Text): txtDiagnose(1).Text = Trim(txtDiagnose(1).Text):
    txtDiagnose(2).Text = Trim(txtDiagnose(2).Text): txtDiagnose(3).Text = Trim(txtDiagnose(3).Text)
    If (Not IsDate(txtDiagnose(0).Text & "-" & txtDiagnose(1).Text & "-" & txtDiagnose(2).Text)) Or (Decode(txtDiagnose(3).Text, "", "-1", txtDiagnose(3).Text) < 0) Or (Decode(txtDiagnose(3).Text, "", "-1", txtDiagnose(3).Text) > 23) Then
        strMsg = strMsg & "<诊断日期>不完整，或不是有效日期，或未精确到小时，请检查！$"
    End If
    
    '死亡日期 允许全空 不允许部份空或数值无效
    txtDeath(0).Text = Trim(txtDeath(0).Text): txtDeath(1).Text = Trim(txtDeath(1).Text): txtDeath(2).Text = Trim(txtDeath(2).Text)
    If txtDeath(0).Text <> "" Or txtDeath(1).Text <> "" Or txtDeath(2).Text <> "" Then
        If Not IsDate(txtDeath(0).Text & "-" & txtDeath(1).Text & "-" & txtDeath(2).Text) Then
            strMsg = strMsg & "<死亡日期>不完整或不是有效日期，请检查！$"
        End If
    End If
    
    CheckValidity = True
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Private Sub ShowMView(x As Long, y As Long)
'功能：显示日期选择控件
    
    MView.Left = x
    If mDateType = 0 Then
        MView.Top = y + 200
    ElseIf mDateType = 3 Then
        MView.Top = y - MView.Height
        MView.Left = txtDeath(0).Left
    Else
        MView.Top = y - MView.Height
    End If
    MView.Visible = True
    Call MView.SetFocus
End Sub

Private Sub lblAttack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lblAttack(Index).MouseIcon = Image1.Picture
    lblAttack(Index).MousePointer = vbCustom
End Sub

Private Sub lblBirth_Click(Index As Integer)
    mDateType = 0
    Call ShowMView(lblBirth(Index).Left, lblBirth(Index).Top)
End Sub

Private Sub lblBirth_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lblBirth(Index).MouseIcon = Image1.Picture
    lblBirth(Index).MousePointer = vbCustom
End Sub

Private Sub lblDeath_Click(Index As Integer)
    mDateType = 3
    Call ShowMView(lblDeath(Index).Left, lblDeath(Index).Top)
End Sub

Private Sub lblDeath_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lblDeath(Index).MouseIcon = Image1.Picture
    lblDeath(Index).MousePointer = vbCustom
End Sub

Private Sub lblDiagnose_Click(Index As Integer)
    If Index <> 3 Then
        mDateType = 2
        Call ShowMView(lblDiagnose(Index).Left, lblDiagnose(Index).Top)
    End If
End Sub

Private Sub lblDiagnose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index <> 3 Then
        Set lblDiagnose(Index).MouseIcon = Image1.Picture
        lblDiagnose(Index).MousePointer = vbCustom
    End If
End Sub

Private Sub MView_DateClick(ByVal DateClicked As Date)
    MView.Visible = False
    Select Case mDateType
        Case 0
            txtBirth(0).Text = MView.Year
            txtBirth(1).Text = MView.Month
            txtBirth(2).Text = MView.Day
        Case 1
            txtAttack(0).Text = MView.Year
            txtAttack(1).Text = MView.Month
            txtAttack(2).Text = MView.Day
        Case 2
            txtDiagnose(0).Text = MView.Year
            txtDiagnose(1).Text = MView.Month
            txtDiagnose(2).Text = MView.Day
        Case 3
            txtDeath(0).Text = MView.Year
            txtDeath(1).Text = MView.Month
            txtDeath(2).Text = MView.Day
    End Select
End Sub

Private Sub MView_LostFocus()
    MView.Visible = False
End Sub

Private Sub txtAge_Change()
    If mblnFirst = False Then
        ucAge(1).Checked = IIf(Trim(txtAge.Text) = "", False, True)
    End If
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Private Sub txtAttack_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Private Sub txtBirth_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
    txtAge.Text = ""
End Sub

Private Sub txtDeath_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Private Sub txtDiagnose_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Private Sub txtIDCard_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyRight Then
        SendKeys "{TAB}"
    Else
        txtIDCard(Index).SelStart = 0
        txtIDCard(Index).SelLength = Len(txtIDCard(Index).Text)
    End If
End Sub

Private Sub txtIDCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If CheckVal(KeyAscii) = True Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub ucCaseType1_Change(Index As Integer)
    gbytDiseaseType = 5
    If ucCaseType1(Index).Checked = True Then
        gbytDiseaseType = Index
    End If
    If Index = 4 Then
        RaiseEvent ClickPositives(ucCaseType1(4).Checked)
    End If
End Sub

Private Sub ucCaseType2_Change(Index As Integer)
    gbytAcute = 2
    If ucCaseType2(Index).Checked = True Then
        gbytAcute = Index
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackColor = vbWhite
End Sub

