VERSION 5.00
Begin VB.UserControl PaneThree 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   LockControls    =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   9825
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   17
      Left            =   7875
      TabIndex        =   36
      Tag             =   "588,608"
      Top             =   2130
      Width           =   1605
      _ExtentX        =   88344
      _ExtentY        =   476
      Caption         =   "新生儿破伤风、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   16
      Left            =   7110
      TabIndex        =   35
      Tag             =   "537,608"
      Top             =   2130
      Width           =   810
      _ExtentX        =   86942
      _ExtentY        =   476
      Caption         =   "白喉、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   15
      Left            =   6150
      TabIndex        =   34
      Tag             =   "474,608"
      Top             =   2130
      Width           =   1005
      _ExtentX        =   87286
      _ExtentY        =   476
      Caption         =   "百日咳、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   14
      Left            =   4275
      TabIndex        =   33
      Tag             =   "349,608"
      Top             =   2130
      Width           =   1905
      _ExtentX        =   88874
      _ExtentY        =   476
      Caption         =   "流行性脑脊髓膜炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucTyphia 
      Height          =   270
      Index           =   1
      Left            =   3210
      TabIndex        =   32
      Tag             =   "280,608"
      Top             =   2130
      Width           =   1185
      _ExtentX        =   50350
      _ExtentY        =   476
      Caption         =   "副伤寒)、"
   End
   Begin zlDisReportCard.uCheckNorm ucTyphia 
      Height          =   270
      Index           =   0
      Left            =   2445
      TabIndex        =   31
      Tag             =   "229,608"
      Top             =   2130
      Width           =   810
      _ExtentX        =   86942
      _ExtentY        =   476
      Caption         =   "伤寒、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   13
      Left            =   1935
      TabIndex        =   30
      Tag             =   "195,608"
      Top             =   2130
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   476
      Caption         =   "伤寒("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   11
      Left            =   4395
      TabIndex        =   22
      Tag             =   "366,584"
      Top             =   1743
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   476
      Caption         =   "痢疾("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   12
      Left            =   7215
      TabIndex        =   25
      Tag             =   "544,584"
      Top             =   1740
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   476
      Caption         =   "肺结核("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   9
      Left            =   8055
      TabIndex        =   18
      Tag             =   "617,561"
      Top             =   1365
      Width           =   1800
      _ExtentX        =   44662
      _ExtentY        =   476
      Caption         =   "流行性乙型脑炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   8
      Left            =   7110
      TabIndex        =   17
      Tag             =   "552,561"
      Top             =   1365
      Width           =   1005
      _ExtentX        =   43259
      _ExtentY        =   476
      Caption         =   "狂犬病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   7
      Left            =   5670
      TabIndex        =   16
      Tag             =   "452,561"
      Top             =   1365
      Width           =   1545
      _ExtentX        =   44212
      _ExtentY        =   476
      Caption         =   "流行性出血热、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   6
      Left            =   4935
      TabIndex        =   15
      Tag             =   "400,561"
      Top             =   1365
      Width           =   810
      _ExtentX        =   42915
      _ExtentY        =   476
      Caption         =   "麻疹、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   1
      Left            =   1950
      TabIndex        =   3
      Tag             =   "202,538"
      Top             =   977
      Width           =   990
      _ExtentX        =   43233
      _ExtentY        =   476
      Caption         =   "艾滋病("
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucAIDS 
      Height          =   270
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Tag             =   "304,538"
      Top             =   977
      Width           =   990
      _ExtentX        =   43233
      _ExtentY        =   476
      Caption         =   "AIDS)、"
   End
   Begin zlDisReportCard.uCheckNorm ucAIDS 
      Height          =   270
      Index           =   0
      Left            =   2925
      TabIndex        =   4
      Tag             =   "259,538"
      Top             =   977
      Width           =   690
      _ExtentX        =   42704
      _ExtentY        =   476
      Caption         =   "HIV、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   4
      Left            =   8895
      TabIndex        =   11
      Tag             =   "647,538"
      Top             =   977
      Width           =   990
      _ExtentX        =   43233
      _ExtentY        =   476
      Caption         =   "未分型)、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   3
      Left            =   8130
      TabIndex        =   10
      Tag             =   "595,538"
      Top             =   977
      Width           =   810
      _ExtentX        =   42915
      _ExtentY        =   476
      Caption         =   "戊型、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   2
      Left            =   7380
      TabIndex        =   9
      Tag             =   "544,538"
      Top             =   977
      Width           =   810
      _ExtentX        =   42915
      _ExtentY        =   476
      Caption         =   "丙型、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   1
      Left            =   6645
      TabIndex        =   8
      Tag             =   "493,538"
      Top             =   977
      Width           =   810
      _ExtentX        =   42915
      _ExtentY        =   476
      Caption         =   "乙型、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   0
      Left            =   5910
      TabIndex        =   7
      Tag             =   "442,538"
      Top             =   977
      Width           =   1200
      _ExtentX        =   43603
      _ExtentY        =   476
      Caption         =   "甲型、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   2
      Left            =   4590
      TabIndex        =   6
      Tag             =   "357,538"
      Top             =   977
      Width           =   1350
      _ExtentX        =   43868
      _ExtentY        =   476
      Caption         =   "病毒性肝炎("
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   4
      Left            =   1395
      TabIndex        =   13
      Tag             =   "164,561"
      Top             =   1365
      Width           =   2100
      _ExtentX        =   45191
      _ExtentY        =   476
      Caption         =   "人感染高致病性禽流感"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   5
      Left            =   3495
      TabIndex        =   14
      Tag             =   "300,561"
      Top             =   1365
      Width           =   1545
      _ExtentX        =   44212
      _ExtentY        =   476
      Caption         =   "甲型H1N1流感、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucAnthrax 
      Height          =   270
      Index           =   2
      Left            =   3435
      TabIndex        =   21
      Tag             =   "297,584"
      Top             =   1740
      Width           =   1080
      _ExtentX        =   74718
      _ExtentY        =   476
      Caption         =   "未分型)、"
   End
   Begin zlDisReportCard.uCheckNorm ucAnthrax 
      Height          =   270
      Index           =   0
      Left            =   1380
      TabIndex        =   19
      Tag             =   "159,584"
      Top             =   1743
      Width           =   983
      _ExtentX        =   82603
      _ExtentY        =   476
      Caption         =   "肺炭疽、"
   End
   Begin zlDisReportCard.uCheckNorm ucAnthrax 
      Height          =   270
      Index           =   1
      Left            =   2356
      TabIndex        =   20
      Tag             =   "221,584"
      Top             =   1743
      Width           =   1184
      _ExtentX        =   82947
      _ExtentY        =   476
      Caption         =   "皮肤炭疽、"
   End
   Begin zlDisReportCard.uCheckNorm ucDysentery 
      Height          =   270
      Index           =   1
      Left            =   5920
      TabIndex        =   24
      Tag             =   "464,584"
      Top             =   1743
      Width           =   1362
      _ExtentX        =   83264
      _ExtentY        =   476
      Caption         =   "阿米巴性)、"
   End
   Begin zlDisReportCard.uCheckNorm ucDysentery 
      Height          =   270
      Index           =   0
      Left            =   4973
      TabIndex        =   23
      Tag             =   "401,584"
      Top             =   1743
      Width           =   1005
      _ExtentX        =   87286
      _ExtentY        =   476
      Caption         =   "细菌性、"
   End
   Begin zlDisReportCard.uCheckNorm ucPTB 
      Height          =   270
      Index           =   0
      Left            =   7992
      TabIndex        =   26
      Tag             =   "591,584"
      Top             =   1743
      Width           =   810
      _ExtentX        =   86942
      _ExtentY        =   476
      Caption         =   "涂阳、"
   End
   Begin zlDisReportCard.uCheckNorm ucPTB 
      Height          =   270
      Index           =   1
      Left            =   8821
      TabIndex        =   27
      Tag             =   "641,584"
      Top             =   1743
      Width           =   1005
      _ExtentX        =   87286
      _ExtentY        =   476
      Caption         =   "仅培阳、"
   End
   Begin zlDisReportCard.uCheckNorm ucPTB 
      Height          =   270
      Index           =   3
      Left            =   888
      TabIndex        =   29
      Tag             =   "128,608"
      Top             =   2126
      Width           =   1172
      _ExtentX        =   82920
      _ExtentY        =   476
      Caption         =   "未痰检)、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   21
      Left            =   7410
      TabIndex        =   45
      Tag             =   "563,631"
      Top             =   2505
      Width           =   1365
      _ExtentX        =   83238
      _ExtentY        =   476
      Caption         =   "钩端螺旋体病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   4
      Left            =   6705
      TabIndex        =   44
      Tag             =   "513,631"
      Top             =   2505
      Width           =   840
      _ExtentX        =   86995
      _ExtentY        =   476
      Caption         =   "隐性)"
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   3
      Left            =   5970
      TabIndex        =   43
      Tag             =   "462,631"
      Top             =   2505
      Width           =   810
      _ExtentX        =   86942
      _ExtentY        =   476
      Caption         =   "胎传、"
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   2
      Left            =   5235
      TabIndex        =   42
      Tag             =   "411,631"
      Top             =   2505
      Width           =   810
      _ExtentX        =   86942
      _ExtentY        =   476
      Caption         =   "Ⅲ期、"
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   1
      Left            =   4530
      TabIndex        =   41
      Tag             =   "360,631"
      Top             =   2505
      Width           =   810
      _ExtentX        =   86942
      _ExtentY        =   476
      Caption         =   "Ⅱ期、"
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   0
      Left            =   3735
      TabIndex        =   40
      Tag             =   "309,631"
      Top             =   2505
      Width           =   810
      _ExtentX        =   86942
      _ExtentY        =   476
      Caption         =   "Ⅰ期、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   20
      Left            =   2310
      TabIndex        =   39
      Tag             =   "227,631"
      Top             =   2505
      Width           =   1500
      _ExtentX        =   68263
      _ExtentY        =   476
      Caption         =   "淋病、梅毒("
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   19
      Left            =   1035
      TabIndex        =   38
      Tag             =   "140,631"
      Top             =   2505
      Width           =   1545
      _ExtentX        =   88239
      _ExtentY        =   476
      Caption         =   "布鲁氏菌病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   18
      Left            =   105
      TabIndex        =   37
      Tag             =   "77,631"
      Top             =   2505
      Width           =   1200
      _ExtentX        =   81280
      _ExtentY        =   476
      Caption         =   "猩红热、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucMalaria 
      Height          =   270
      Index           =   0
      Left            =   915
      TabIndex        =   48
      Tag             =   "125,654"
      Top             =   2892
      Width           =   1005
      _ExtentX        =   87286
      _ExtentY        =   476
      Caption         =   "间日疟、"
   End
   Begin zlDisReportCard.uCheckNorm ucMalaria 
      Height          =   270
      Index           =   1
      Left            =   1965
      TabIndex        =   49
      Tag             =   "188,654"
      Top             =   2892
      Width           =   1005
      _ExtentX        =   87286
      _ExtentY        =   476
      Caption         =   "恶性疟、"
   End
   Begin zlDisReportCard.uCheckNorm ucMalaria 
      Height          =   270
      Index           =   2
      Left            =   2970
      TabIndex        =   50
      Tag             =   "247,654"
      Top             =   2892
      Width           =   1005
      _ExtentX        =   87286
      _ExtentY        =   476
      Caption         =   "未分型)"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   6
      Left            =   8790
      TabIndex        =   57
      Tag             =   "665,702"
      Top             =   3555
      Width           =   1035
      _ExtentX        =   81412
      _ExtentY        =   476
      Caption         =   "黑热病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   5
      Left            =   6420
      TabIndex        =   56
      Tag             =   "504,702"
      Top             =   3555
      Width           =   2475
      _ExtentX        =   83952
      _ExtentY        =   476
      Caption         =   "流行性和地方性斑疹伤寒、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   4
      Left            =   5475
      TabIndex        =   55
      Tag             =   "441,702"
      Top             =   3555
      Width           =   1035
      _ExtentX        =   81412
      _ExtentY        =   476
      Caption         =   "麻风病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   3
      Left            =   3615
      TabIndex        =   54
      Tag             =   "316,702"
      Top             =   3555
      Width           =   1980
      _ExtentX        =   83079
      _ExtentY        =   476
      Caption         =   "急性出血性结膜炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   2
      Left            =   2850
      TabIndex        =   53
      Tag             =   "265,702"
      Top             =   3555
      Width           =   960
      _ExtentX        =   81280
      _ExtentY        =   476
      Caption         =   "风疹、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   1
      Left            =   1380
      TabIndex        =   52
      Tag             =   "164,702"
      Top             =   3555
      Width           =   1560
      _ExtentX        =   82338
      _ExtentY        =   476
      Caption         =   "流行性腮腺炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   8
      Left            =   1200
      TabIndex        =   59
      Tag             =   "140,725"
      Top             =   3885
      Width           =   1095
      _ExtentX        =   87445
      _ExtentY        =   476
      Caption         =   "丝虫病，"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   9
      Left            =   2295
      TabIndex        =   60
      Tag             =   "203,725"
      Top             =   3885
      Width           =   5850
      _ExtentX        =   95832
      _ExtentY        =   476
      Caption         =   "除霍乱、细菌性和阿米巴性痢疾、伤寒和副伤寒以外的感染性腹泻病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   10
      Left            =   8160
      TabIndex        =   61
      Tag             =   "590,725"
      Top             =   3885
      Width           =   1095
      _ExtentX        =   87445
      _ExtentY        =   476
      Caption         =   "手足口病 "
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousA 
      Height          =   270
      Index           =   1
      Left            =   1050
      TabIndex        =   1
      Tag             =   "129,490"
      Top             =   301
      Width           =   825
      _ExtentX        =   84905
      _ExtentY        =   476
      Caption         =   "霍乱"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Tag             =   "77,538"
      Top             =   977
      Width           =   1905
      _ExtentX        =   82524
      _ExtentY        =   476
      Caption         =   "传染性非典型肺炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   3
      Left            =   105
      TabIndex        =   12
      Tag             =   "77,561"
      Top             =   1365
      Width           =   1395
      _ExtentX        =   81624
      _ExtentY        =   476
      Caption         =   "脊髓灰质炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   10
      Left            =   105
      TabIndex        =   65
      Tag             =   "77,584"
      Top             =   1743
      Width           =   1302
      _ExtentX        =   2302
      _ExtentY        =   476
      Caption         =   "登革热、炭疽("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucPTB 
      Height          =   270
      Index           =   2
      Left            =   105
      TabIndex        =   28
      Tag             =   "77,608"
      Top             =   2126
      Width           =   805
      _ExtentX        =   80592
      _ExtentY        =   476
      Caption         =   "菌阴、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   22
      Left            =   8760
      TabIndex        =   46
      Tag             =   "660,631"
      Top             =   2505
      Width           =   1005
      _ExtentX        =   84746
      _ExtentY        =   476
      Caption         =   "血吸虫病"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   23
      Left            =   105
      TabIndex        =   47
      Tag             =   "78,654"
      Top             =   2892
      Width           =   810
      _ExtentX        =   80592
      _ExtentY        =   476
      Caption         =   "疟疾("
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   51
      Tag             =   "77,702"
      Top             =   3555
      Width           =   1350
      _ExtentX        =   81545
      _ExtentY        =   476
      Caption         =   "流行性感冒、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   7
      Left            =   105
      TabIndex        =   58
      Tag             =   "77,725"
      Top             =   3885
      Width           =   1095
      _ExtentX        =   81095
      _ExtentY        =   476
      Caption         =   "包虫病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousA 
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Tag             =   "77,490"
      Top             =   270
      Width           =   825
      _ExtentX        =   80619
      _ExtentY        =   476
      Caption         =   "鼠疫、"
      CheckType       =   1
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   9982
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   9982
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "丙类传染病*："
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
      Index           =   29
      Left            =   105
      TabIndex        =   64
      Tag             =   "78,680"
      Top             =   3285
      Width           =   1170
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "乙类传染病*："
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
      Index           =   28
      Left            =   120
      TabIndex        =   63
      Tag             =   "79,517"
      Top             =   684
      Width           =   1170
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "甲类传染病*："
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
      Index           =   27
      Left            =   120
      TabIndex        =   62
      Tag             =   "79,468"
      Top             =   8
      Width           =   1170
   End
End
Attribute VB_Name = "PaneThree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolLoadData As Collection  '保存控件显示信息

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

Public Sub PrintThree()
    Dim objCtl As Control
    For Each objCtl In UserControl.Controls
        Call PrintInfo(objCtl)
    Next
End Sub

Public Sub LoadData(colData As Collection, bytType As Byte, ByVal strChkType As String)
    Dim strTmp As String
    Dim strName As String
    Dim i As Integer
    Dim strInfo() As String
    Dim objCtl As Control
    
    
    On Error GoTo errHand
    If bytType = 1 Then
        For Each objCtl In UserControl.Controls
            If TypeName(objCtl) = "uCheckNorm" Then
                strTmp = Trim(objCtl.Caption)
                strTmp = Replace(strTmp, "(", "")
                strTmp = Replace(strTmp, ")", "")
                strTmp = Replace(strTmp, "、", "")
                Select Case objCtl.Name
                    Case "ucHepatitis"
                        If InStr(strChkType, "23," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucAnthrax"
                        If InStr(strChkType, "24," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucMalaria"
                        If InStr(strChkType, "29," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case Else
                        If InStr(strChkType, strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                End Select
            End If
        Next
    Else
        strTmp = CStr(colData("K16"))
        
        ucInfectiousA(0).Checked = IIf(InStr(strTmp, "鼠疫") <> 0, True, False)
        ucInfectiousA(1).Checked = IIf(InStr(strTmp, "霍乱") <> 0, True, False)
        
        For i = 0 To 23
            strName = Trim(ucInfectiousB(i).Caption)
            strName = Mid(strName, 1, Len(strName) - 1)
            If InStr(strTmp, strName) <> 0 Then
                ucInfectiousB(i).Checked = True
            End If
        Next
        
        For i = 0 To 10
            strName = Trim(ucInfectiousC(i).Caption)
            strName = Mid(strName, 1, Len(strName) - 1)
            If InStr(strTmp, strName) <> 0 Then
                ucInfectiousC(i).Checked = True
            End If
        Next
        
        ucAIDS(0).Checked = IIf(InStr(strTmp, "HIV") <> 0, True, False)
        ucAIDS(1).Checked = IIf(InStr(strTmp, "AIDS") <> 0, True, False)
        
        For i = 0 To 4
            strName = Trim(ucHepatitis(i).Caption)
            strName = Mid(strName, 1, Len(strName) - IIf(i = 4, 2, 1))
            If InStr(strTmp, strName) <> 0 Then
                ucHepatitis(i).Checked = True
            End If
        Next
        
        ucAnthrax(0).Checked = IIf(InStr(strTmp, "肺炭疽") <> 0, True, False)
        ucAnthrax(1).Checked = IIf(InStr(strTmp, "皮肤炭疽") <> 0, True, False)
        ucAnthrax(2).Checked = IIf(InStr(strTmp, "未分型") <> 0, True, False)
        
        ucDysentery(0).Checked = IIf(InStr(strTmp, "细菌性") <> 0, True, False)
        ucDysentery(1).Checked = IIf(InStr(strTmp, "阿米巴性") <> 0, True, False)
        
        For i = 0 To 3
            strName = Trim(ucPTB(i).Caption)
            strName = Mid(strName, 1, Len(strName) - IIf(i = 3, 2, 1))
            If InStr(strTmp, strName) <> 0 Then
                ucPTB(i).Checked = True
            End If
        Next
        
        ucTyphia(0).Checked = IIf(InStr(strTmp, "伤寒") <> 0, True, False)
        ucTyphia(1).Checked = IIf(InStr(strTmp, "副伤寒") <> 0, True, False)
        
        For i = 0 To 4
            strName = Trim(ucSyphilis(i).Caption)
            strName = Mid(strName, 1, Len(strName) - IIf(i = 4, 2, 1))
            If InStr(strTmp, strName) <> 0 Then
                ucSyphilis(i).Checked = True
            End If
        Next
        
        ucMalaria(0).Checked = IIf(InStr(strTmp, "间日疟") <> 0, True, False)
        ucMalaria(1).Checked = IIf(InStr(strTmp, "恶性疟") <> 0, True, False)
        ucMalaria(2).Checked = IIf(InStr(strTmp, "未分型") <> 0, True, False)
    End If
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
    strObjNo = "20$21$22$23$24$25$26$27$28$29$30"
    
    '甲类传染病
    strTmp = IIf(ucInfectiousA(0).Checked = True, ucInfectiousA(0).Caption & ";", "")
    strTmp = Trim(strTmp) & Trim(IIf(ucInfectiousA(1).Checked = True, ucInfectiousA(1).Caption, ""))
    strContent = strContent & strTmp & "$"
    
    '乙类传染病
    strTmp = ""
    For i = 0 To ucInfectiousB.UBound
        If ucInfectiousB(i).Checked = True Then
            strTmp = strTmp & ";" & ucInfectiousB(i).Caption
        End If
    Next
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 2)
    End If
    strContent = strContent & strTmp & "$"
    
    '艾滋病
    strTmp = Decode(True, ucAIDS(0).Checked, ucAIDS(0).Caption, ucAIDS(1).Checked, ucAIDS(1).Caption, "")
    strContent = strContent & strTmp & "$"
    
    '病毒性肝炎
    strTmp = Decode(True, ucHepatitis(0).Checked, ucHepatitis(0).Caption, ucHepatitis(1).Checked, ucHepatitis(1).Caption, ucHepatitis(2).Checked, ucHepatitis(2).Caption, ucHepatitis(3).Checked, ucHepatitis(3).Caption, ucHepatitis(4).Checked, ucHepatitis(4).Caption, "")
    strContent = strContent & strTmp & "$"
    
    '炭疽
    strTmp = Decode(True, ucAnthrax(0).Checked, ucAnthrax(0).Caption, ucAnthrax(1).Checked, ucAnthrax(1).Caption, ucAnthrax(2).Checked, ucAnthrax(2).Caption, "")
    strContent = strContent & strTmp & "$"
    
    '痢疾
    strTmp = Decode(True, ucDysentery(0).Checked, ucDysentery(0).Caption, ucDysentery(1).Checked, ucDysentery(1).Caption, "")
    strContent = strContent & strTmp & "$"
    
    '肺结核
    strTmp = Decode(True, ucPTB(0).Checked, ucPTB(0).Caption, ucPTB(1).Checked, ucPTB(1).Caption, ucPTB(2).Checked, ucPTB(2).Caption, ucPTB(3).Checked, ucPTB(3).Caption, "")
    strContent = strContent & strTmp & "$"
    
    '伤寒
    strTmp = Decode(True, ucTyphia(0).Checked, ucTyphia(0).Caption, ucTyphia(1).Checked, ucTyphia(1).Caption, "")
    strContent = strContent & strTmp & "$"
    
    '淋病
    strTmp = Decode(True, ucSyphilis(0).Checked, ucSyphilis(0).Caption, ucSyphilis(1).Checked, ucSyphilis(1).Caption, ucSyphilis(2).Checked, ucSyphilis(2).Caption, ucSyphilis(3).Checked, ucSyphilis(3).Caption, ucSyphilis(4).Checked, ucSyphilis(4).Caption, "")
    strContent = strContent & strTmp & "$"
    
    '疟疾
    strTmp = Decode(True, ucMalaria(0).Checked, ucMalaria(0).Caption, ucMalaria(1).Checked, ucMalaria(1).Caption, ucMalaria(2).Checked, ucMalaria(2).Caption, "")
    strContent = strContent & strTmp & "$"
    
    '丙类传染病
    strTmp = ""
    For i = 0 To ucInfectiousC.UBound
        If ucInfectiousC(i).Checked = True Then
            strTmp = strTmp & ";" & ucInfectiousC(i).Caption
        End If
    Next
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 2)
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
'检查输入合法性

    On Error GoTo errHand
    CheckValidity = False
    '1.检查病例分类为"病原携带者"时，病种是否是"霍乱、脊髓灰质炎、艾滋病"
    If gbytDiseaseType = 3 Then
        If ucInfectiousA(1).Checked = False And ucInfectiousB(1).Checked = False And ucInfectiousB(3).Checked = False Then
            strMsg = strMsg & "需报告病原携带者的法定传染病病种包括<霍乱>、<脊髓灰质炎>、<艾滋病>，请检查！$"
        End If
    End If
    
    '2.检查病例分类为"阳性检测结果"时，病种必须是HIV
    If gbytDiseaseType = 4 Then
        If ucAIDS(0).Checked = False Then
            strMsg = strMsg & "病种是<HIV>时病例分类才能是<阳性检测结果>，请检查！$"
        End If
    End If
    
    '3."梅毒"、"淋病"的病例分类只能为"实验室诊断病例"和"疑似病例"
    If ucInfectiousB(20).Checked = True Then
        If gbytDiseaseType <> 0 And gbytDiseaseType <> 2 Then
            strMsg = strMsg & "<梅毒、淋病>的病例分类只能为<实验室诊断病例>和<疑似病例>！$"
        End If
    End If
    
    '4.乙肝、血吸虫病例须分急性或慢性填写
    If ucHepatitis(1).Checked = True Or ucInfectiousB(23).Checked = True Then
        If gbytAcute <> 0 And gbytAcute <> 1 Then
            strMsg = strMsg & "<乙肝>、<血吸虫病例>须分<急性>或<慢性>，请检查！$"
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

Public Sub SetAIDS(blnSelected As Boolean)
    ucAIDS(0).Checked = blnSelected
    ucInfectiousB(1).Checked = blnSelected
End Sub
Private Sub ucAIDS_Change(Index As Integer)
    ucInfectiousB(1).Checked = ucAIDS(Index).Checked
End Sub

Private Sub ucAnthrax_Change(Index As Integer)
    ucInfectiousB(10).Checked = ucAnthrax(Index).Checked
End Sub

Private Sub ucDysentery_Change(Index As Integer)
    ucInfectiousB(11).Checked = ucDysentery(Index).Checked
End Sub

Private Sub ucHepatitis_Change(Index As Integer)
    ucInfectiousB(2).Checked = ucHepatitis(Index).Checked
End Sub

Private Sub ucInfectiousB_Change(Index As Integer)
    Dim i As Integer
    
    On Error GoTo errHand
    Select Case Index
    '艾滋病
    Case 1
        If ucInfectiousB(1).Checked = True Then
            ucAIDS(0).Checked = True
        Else
            ucAIDS(0).Checked = False
            ucAIDS(1).Checked = False
        End If

    '病毒性肝炎
    Case 2
        If ucInfectiousB(2).Checked = True Then
            ucHepatitis(0).Checked = True
        Else
            For i = 0 To 4
                ucHepatitis(i).Checked = False
            Next
        End If
        
    '登革热、炭疽
    Case 10
        If ucInfectiousB(10).Checked = True Then
            ucAnthrax(0).Checked = True
        Else
            For i = 0 To 2
                ucAnthrax(i).Checked = False
            Next
        End If
        
    '痢疾
    Case 11
        If ucInfectiousB(11).Checked = True Then
            ucDysentery(0).Checked = True
        Else
            ucDysentery(0).Checked = False
            ucDysentery(1).Checked = False
        End If
        
    '肺结核
    Case 12
        If ucInfectiousB(12).Checked = True Then
            ucPTB(0).Checked = True
        Else
            For i = 0 To 3
                ucPTB(i).Checked = False
            Next
        End If
        
    '伤寒
    Case 13
        If ucInfectiousB(13).Checked = True Then
            ucTyphia(0).Checked = True
        Else
            ucTyphia(0).Checked = False
            ucTyphia(1).Checked = False
        End If
    
    '梅毒
    Case 20
        If ucInfectiousB(20).Checked = True Then
            ucSyphilis(0).Checked = True
        Else
            For i = 0 To 4
                ucSyphilis(i).Checked = False
            Next
        End If
        
    '疟疾
    Case 23
        If ucInfectiousB(23).Checked = True Then
            ucMalaria(0).Checked = True
        Else
            For i = 0 To 2
                ucMalaria(i).Checked = False
            Next
        End If
    End Select
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Sub ucMalaria_Change(Index As Integer)
    ucInfectiousB(23).Checked = ucMalaria(Index).Checked
End Sub

Private Sub ucPTB_Change(Index As Integer)
    ucInfectiousB(12).Checked = ucPTB(Index).Checked
End Sub

Private Sub ucSyphilis_Change(Index As Integer)
    ucInfectiousB(20).Checked = ucSyphilis(Index).Checked
End Sub

Private Sub ucTyphia_Change(Index As Integer)
    ucInfectiousB(13).Checked = ucTyphia(Index).Checked
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackColor = vbWhite
End Sub
