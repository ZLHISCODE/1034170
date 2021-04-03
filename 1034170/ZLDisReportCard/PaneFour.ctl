VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl PaneFour 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   LockControls    =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   9825
   Begin MSComCtl2.MonthView MView 
      Height          =   2220
      Left            =   7830
      TabIndex        =   54
      Top             =   4680
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12648447
      Appearance      =   1
      StartOfWeek     =   232914946
      CurrentDate     =   42010
   End
   Begin zlDisReportCard.uCheckNorm ucWay 
      Height          =   270
      Index           =   8
      Left            =   8055
      TabIndex        =   34
      Tag             =   "611,908"
      Top             =   2445
      Width           =   1050
      _ExtentX        =   56462
      _ExtentY        =   476
      Caption         =   "不详"
   End
   Begin zlDisReportCard.uCheckNorm ucWay 
      Height          =   270
      Index           =   7
      Left            =   7260
      TabIndex        =   33
      Tag             =   "559,908"
      Top             =   2445
      Width           =   1050
      _ExtentX        =   56462
      _ExtentY        =   476
      Caption         =   "其他、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay 
      Height          =   270
      Index           =   6
      Left            =   6510
      TabIndex        =   32
      Tag             =   "508,908"
      Top             =   2445
      Width           =   1050
      _ExtentX        =   56462
      _ExtentY        =   476
      Caption         =   "间接、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay 
      Height          =   270
      Index           =   5
      Left            =   5385
      TabIndex        =   31
      Tag             =   "432,908"
      Top             =   2445
      Width           =   1230
      _ExtentX        =   56780
      _ExtentY        =   476
      Caption         =   "职业暴露、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay2 
      Height          =   270
      Index           =   2
      Left            =   3870
      TabIndex        =   30
      Tag             =   "331,908"
      Top             =   2445
      Width           =   1635
      _ExtentX        =   57494
      _ExtentY        =   476
      Caption         =   "输血/血制品)、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay2 
      Height          =   270
      Index           =   1
      Left            =   2745
      TabIndex        =   29
      Tag             =   "255,908"
      Top             =   2460
      Width           =   1200
      _ExtentX        =   56727
      _ExtentY        =   476
      Caption         =   "注射毒品、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay2 
      Height          =   270
      Index           =   0
      Left            =   1950
      TabIndex        =   28
      Tag             =   "203,908"
      Top             =   2445
      Width           =   810
      _ExtentX        =   59849
      _ExtentY        =   476
      Caption         =   "采血、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay 
      Height          =   270
      Index           =   3
      Left            =   7515
      TabIndex        =   26
      Tag             =   "561,884"
      Top             =   2052
      Width           =   1815
      _ExtentX        =   57811
      _ExtentY        =   476
      Caption         =   "性接触+注射毒品、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay 
      Height          =   270
      Index           =   2
      Left            =   6330
      TabIndex        =   25
      Tag             =   "485,884"
      Top             =   2052
      Width           =   1215
      _ExtentX        =   56753
      _ExtentY        =   476
      Caption         =   "同性接触、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay 
      Height          =   270
      Index           =   1
      Left            =   5220
      TabIndex        =   24
      Tag             =   "409,884"
      Top             =   2052
      Width           =   1245
      _ExtentX        =   56806
      _ExtentY        =   476
      Caption         =   "母婴传播、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay1 
      Height          =   270
      Index           =   1
      Left            =   3825
      TabIndex        =   19
      Tag             =   "315,884"
      Top             =   2052
      Width           =   1550
      _ExtentX        =   57335
      _ExtentY        =   476
      Caption         =   "非婚性接触)、"
   End
   Begin zlDisReportCard.uCheckNorm ucWay1 
      Height          =   270
      Index           =   0
      Left            =   3030
      TabIndex        =   18
      Tag             =   "263,884"
      Top             =   2052
      Width           =   810
      _ExtentX        =   59849
      _ExtentY        =   476
      Caption         =   "配偶、"
   End
   Begin zlDisReportCard.uCheckNorm ucEducation 
      Height          =   270
      Index           =   6
      Left            =   5415
      TabIndex        =   16
      Tag             =   "428,859"
      Top             =   1725
      Width           =   1200
      _ExtentX        =   60537
      _ExtentY        =   476
      Caption         =   "硕士及以上 "
   End
   Begin zlDisReportCard.uCheckNorm ucEducation 
      Height          =   270
      Index           =   5
      Left            =   4770
      TabIndex        =   15
      Tag             =   "387,859"
      Top             =   1725
      Width           =   810
      _ExtentX        =   59849
      _ExtentY        =   476
      Caption         =   "大学"
   End
   Begin zlDisReportCard.uCheckNorm ucEducation 
      Height          =   270
      Index           =   4
      Left            =   4125
      TabIndex        =   14
      Tag             =   "346,859"
      Top             =   1725
      Width           =   810
      _ExtentX        =   59849
      _ExtentY        =   476
      Caption         =   "大专"
   End
   Begin zlDisReportCard.uCheckNorm ucEducation 
      Height          =   270
      Index           =   3
      Left            =   2955
      TabIndex        =   13
      Tag             =   "269,859"
      Top             =   1725
      Width           =   1185
      _ExtentX        =   56700
      _ExtentY        =   476
      Caption         =   "高中或中专"
   End
   Begin zlDisReportCard.uCheckNorm ucEducation 
      Height          =   270
      Index           =   2
      Left            =   2340
      TabIndex        =   12
      Tag             =   "228,859"
      Top             =   1719
      Width           =   810
      _ExtentX        =   56039
      _ExtentY        =   476
      Caption         =   "初中"
   End
   Begin zlDisReportCard.uCheckNorm ucEducation 
      Height          =   270
      Index           =   1
      Left            =   1710
      TabIndex        =   11
      Tag             =   "187,859"
      Top             =   1719
      Width           =   810
      _ExtentX        =   56039
      _ExtentY        =   476
      Caption         =   "小学"
   End
   Begin zlDisReportCard.uCheckNorm ucEducation 
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Tag             =   "146,859"
      Top             =   1719
      Width           =   810
      _ExtentX        =   56039
      _ExtentY        =   476
      Caption         =   "文盲"
   End
   Begin zlDisReportCard.uCheckNorm ucMarital 
      Height          =   270
      Index           =   3
      Left            =   3525
      TabIndex        =   9
      Tag             =   "305,836"
      Top             =   1386
      Width           =   810
      _ExtentX        =   56039
      _ExtentY        =   476
      Caption         =   "不详"
   End
   Begin zlDisReportCard.uCheckNorm ucMarital 
      Height          =   270
      Index           =   2
      Left            =   2340
      TabIndex        =   8
      Tag             =   "228,836"
      Top             =   1386
      Width           =   1230
      _ExtentX        =   56780
      _ExtentY        =   476
      Caption         =   "离异或丧偶"
   End
   Begin zlDisReportCard.uCheckNorm ucMarital 
      Height          =   270
      Index           =   1
      Left            =   1710
      TabIndex        =   7
      Tag             =   "187,836"
      Top             =   1386
      Width           =   810
      _ExtentX        =   56039
      _ExtentY        =   476
      Caption         =   "已婚"
   End
   Begin zlDisReportCard.uCheckNorm ucMarital 
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Tag             =   "146,836"
      Top             =   1386
      Width           =   810
      _ExtentX        =   56039
      _ExtentY        =   476
      Caption         =   "未婚"
   End
   Begin zlDisReportCard.uCheckNorm ucVD 
      Height          =   270
      Index           =   2
      Left            =   3255
      TabIndex        =   5
      Tag             =   "288,813"
      Top             =   1050
      Width           =   1800
      _ExtentX        =   61595
      _ExtentY        =   476
      Caption         =   "生殖道衣原体感染 "
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucVD 
      Height          =   270
      Index           =   1
      Left            =   2085
      TabIndex        =   4
      Tag             =   "211,813"
      Top             =   1050
      Width           =   1200
      _ExtentX        =   60537
      _ExtentY        =   476
      Caption         =   "生殖器疱疹"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucVD 
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Tag             =   "146,813"
      Top             =   1053
      Width           =   1050
      _ExtentX        =   60272
      _ExtentY        =   476
      Caption         =   "尖锐湿疣"
      CheckType       =   1
   End
   Begin VB.TextBox txtEnter 
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
      Left            =   7500
      TabIndex        =   53
      Tag             =   "577,1004"
      ToolTipText     =   "填卡时间在完成时由程序自动生成"
      Top             =   3805
      Width           =   450
   End
   Begin VB.TextBox txtEnter 
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
      Left            =   6240
      TabIndex        =   49
      Tag             =   "508,1004"
      ToolTipText     =   "填卡时间在完成时由程序自动生成"
      Top             =   3805
      Width           =   1095
   End
   Begin VB.TextBox txtEnter 
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
      Left            =   8130
      TabIndex        =   48
      Tag             =   "620,1004"
      ToolTipText     =   "填卡时间在完成时由程序自动生成"
      Top             =   3805
      Width           =   525
   End
   Begin VB.TextBox txtRemarks 
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
      Height          =   480
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   47
      Tag             =   "113,1027"
      Top             =   4185
      Width           =   9060
   End
   Begin VB.TextBox txtDoctor 
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
      Left            =   1035
      TabIndex        =   43
      Tag             =   "143,1004"
      ToolTipText     =   "填卡医生在完成时由程序自动生成"
      Top             =   3805
      Width           =   3330
   End
   Begin VB.TextBox txtDocNumber 
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
      Left            =   6150
      TabIndex        =   41
      Tag             =   "479,979"
      Top             =   3445
      Width           =   2505
   End
   Begin VB.TextBox txtUnit 
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
      Left            =   1050
      TabIndex        =   39
      Tag             =   "137,979"
      Top             =   3445
      Width           =   3255
   End
   Begin VB.TextBox txtReason 
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
      Left            =   6150
      TabIndex        =   37
      Tag             =   "479,957"
      Top             =   3100
      Width           =   2520
   End
   Begin VB.TextBox txtIName 
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
      Left            =   1005
      TabIndex        =   35
      Tag             =   "137,957"
      Top             =   3100
      Width           =   3300
   End
   Begin VB.TextBox txtImportant 
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
      Height          =   270
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   1
      Tag             =   "79,773"
      Top             =   270
      Width           =   9500
   End
   Begin zlDisReportCard.uCheckNorm ucWay 
      Height          =   270
      Index           =   0
      Left            =   2085
      TabIndex        =   17
      Tag             =   "204,884"
      Top             =   2052
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   476
      Caption         =   "异性传播("
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucWay 
      Height          =   270
      Index           =   4
      Left            =   1080
      TabIndex        =   27
      Tag             =   "144,908"
      Top             =   2445
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   476
      Caption         =   "血液传播("
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8625
      Picture         =   "PaneFour.ctx":0000
      Top             =   1035
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line5 
      Tag             =   "611,1014,644"
      X1              =   8130
      X2              =   8565
      Y1              =   3990
      Y2              =   3990
   End
   Begin VB.Line Line4 
      Tag             =   "569,1014,599"
      X1              =   7500
      X2              =   7935
      Y1              =   3990
      Y2              =   3990
   End
   Begin VB.Line Line1 
      Index           =   5
      Tag             =   "485,1014,555"
      X1              =   6240
      X2              =   7305
      Y1              =   3990
      Y2              =   3990
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
      Left            =   7335
      TabIndex        =   52
      Tag             =   "558,1004"
      Top             =   3810
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
      Left            =   7965
      TabIndex        =   51
      Tag             =   "600,1004"
      Top             =   3810
      Width           =   180
   End
   Begin VB.Label lblAttack 
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
      Left            =   8640
      TabIndex        =   50
      Tag             =   "644,1004"
      Top             =   3810
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "备注："
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
      Left            =   105
      TabIndex        =   46
      Tag             =   "78,1027"
      Top             =   4170
      Width           =   540
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "填卡日期*："
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
      Left            =   5265
      TabIndex        =   45
      Tag             =   "420,1004"
      ToolTipText     =   "填卡时间在完成时由程序自动生成"
      Top             =   3810
      Width           =   990
   End
   Begin VB.Line Line1 
      Index           =   4
      Tag             =   "143,1014,244"
      X1              =   1035
      X2              =   4350
      Y1              =   3990
      Y2              =   3990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "填卡医生*："
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
      Left            =   105
      TabIndex        =   44
      Tag             =   "78,1004"
      ToolTipText     =   "填卡医生在完成时由程序自动生成"
      Top             =   3810
      Width           =   990
   End
   Begin VB.Line Line1 
      Index           =   3
      Tag             =   "479,990,652"
      X1              =   6150
      X2              =   8685
      Y1              =   3630
      Y2              =   3630
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
      Index           =   8
      Left            =   5265
      TabIndex        =   42
      Tag             =   "420,979"
      Top             =   3450
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   2
      Tag             =   "137,990,358"
      X1              =   1020
      X2              =   4335
      Y1              =   3630
      Y2              =   3630
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "报告单位："
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
      TabIndex        =   40
      Tag             =   "79,979"
      Top             =   3450
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   1
      Tag             =   "479,967,652"
      X1              =   6150
      X2              =   8655
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "退卡原因："
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
      Left            =   5265
      TabIndex        =   38
      Tag             =   "420,957"
      Top             =   3105
      Width           =   900
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "订正病名："
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
      TabIndex        =   36
      Tag             =   "78,957"
      Top             =   3105
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   9
      Tag             =   "137,967,358"
      X1              =   1020
      X2              =   4335
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "监测性病*："
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
      TabIndex        =   23
      Tag             =   "79,818"
      Top             =   1098
      Width           =   990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "婚姻状况*："
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
      Left            =   105
      TabIndex        =   22
      Tag             =   "78,841 "
      Top             =   1431
      Width           =   990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "文化程度*："
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
      TabIndex        =   21
      Tag             =   "78,864"
      Top             =   1764
      Width           =   990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "感染途径（接触史）*："
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
      Left            =   105
      TabIndex        =   20
      Tag             =   "78,888"
      Top             =   2097
      Width           =   1890
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "性病报告附加栏（报告性病时须加填本栏项目）*"
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
      Left            =   105
      TabIndex        =   2
      Tag             =   "78,794"
      Top             =   765
      Width           =   3870
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "其他法定管理以及重点监测传染病*："
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
      Left            =   105
      TabIndex        =   0
      Tag             =   "79,750"
      Top             =   30
      Width           =   2970
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   9825
      Y1              =   4155
      Y2              =   4155
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9825
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   9825
      Y1              =   675
      Y2              =   675
   End
End
Attribute VB_Name = "PaneFour"
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

Public Sub PrintFour()
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
    
    On Error GoTo errHand
    If bytType = 1 Then
        txtImportant.Text = CStr(colData("K31"))
        
        For Each objCtl In UserControl.Controls
            If TypeName(objCtl) = "uCheckNorm" Then
                strTmp = Trim(objCtl.Caption)
                strTmp = Replace(strTmp, "(", "")
                strTmp = Replace(strTmp, ")", "")
                strTmp = Replace(strTmp, "、", "")
                Select Case objCtl.Name
                    Case "ucMarital"
                        If InStr(strChkType, "33," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucWay", "ucWay2"
                        If InStr(strChkType, "35," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucEducation"
                        If InStr(strChkType, "34," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case Else
                        If InStr(strChkType, strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                End Select
            End If
        Next

        txtIName.Text = CStr(colData("K38"))
        txtReason.Text = CStr(colData("K39"))
        txtUnit.Text = CStr(colData("K40"))
        txtDocNumber.Text = CStr(colData("K41"))
        txtDoctor.Text = CStr(colData("K42"))
        txtRemarks.Text = CStr(colData("K44"))
        
        strTmp = CStr(colData("K43"))
        strInfo = Split(strTmp, "-")
        For i = 0 To UBound(strInfo)
            txtEnter(i) = strInfo(i)
        Next
    Else
'       仅当性病报告时才填的内容不自动填充
'        For i = 0 To 3
'            If ucMarital(i).Caption = CStr(colData("K9")) Then
'                ucMarital(i).Checked = True
'                Exit For
'            End If
'            ucMarital(3).Checked = True
'        Next
'
'        For i = 0 To 6
'            If ucEducation(i).Caption = CStr(colData("K10")) Then
'                ucEducation(i).Checked = True
'                Exit For
'            End If
'            ucEducation(3).Checked = True
'        Next
        
        txtUnit.Text = CStr(colData("K11"))
        
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
    strObjNo = "31$32$33$34$35$36$37$38$39$40$41$42$43$44"
    
    '其它传染病
    strContent = Trim(txtImportant.Text) & "$"
    '监测性病
    strTmp = ""
    strTmp1 = ""
    For i = 1 To 3
        If ucVD(i - 1).Checked = True Then
            strTmp = strTmp & "1;"
            strTmp1 = strTmp1 & ucVD(i - 1).Caption & ";"
        Else
            strTmp = strTmp & "0;"
        End If
    Next
    strContent = strContent & strTmp1 & "$"
    '婚姻状况
    For i = 0 To 3
        If ucMarital(i).Checked = True Then
            strTmp = i + 1
            strTmp1 = ucMarital(i).Caption
            Exit For
        End If
        strTmp = i + 2
        strTmp1 = ""
    Next

    strContent = strContent & strTmp1 & "$"
    '学历
    For i = 0 To 6
        If ucEducation(i).Checked = True Then
            strTmp = i + 1
            strTmp1 = ucEducation(i).Caption
            Exit For
        End If
        strTmp = i + 2
        strTmp1 = ""
    Next

    strContent = strContent & strTmp1 & "$"
    '感染途径
    strTmp = ""
    strTmp1 = ""
    For i = 0 To 8
        If ucWay(i).Checked = True Then
            strTmp = strTmp & "1;"
            strTmp1 = strTmp1 & ucWay(i).Caption & ";"
        Else
            strTmp = strTmp & "0;"
        End If
    Next

    strContent = strContent & strTmp1 & "$"
    
    '异性传播
    strTmp = IIf(ucWay1(0).Checked = True, 1, IIf(ucWay1(1).Checked = True, 2, 3))

    
    strTmp = IIf(ucWay1(0).Checked = True, ucWay1(0).Caption, IIf(ucWay1(1).Checked = True, ucWay1(1).Caption, ""))
    strContent = strContent & strTmp & "$"
    
    '血液传播
    For i = 0 To 2
        If ucWay2(i).Checked = True Then
            strTmp = i + 1
            strTmp1 = ucWay2(i).Caption
            Exit For
        End If
        strTmp = i + 2
        strTmp1 = ""
    Next

    strContent = strContent & strTmp1 & "$"
    
    '订正病名、退卡原因、报告单位、联系电话、填卡医生

    strContent = strContent & txtIName.Text & "$" & txtReason.Text & "$" & txtUnit.Text & "$" & txtDocNumber.Text & "$" & txtDoctor.Text & "$"
    
    '填卡日期
    
    strTmp = txtEnter(0).Text & "-" & txtEnter(1).Text & "-" & txtEnter(2).Text
    If Trim(strTmp) = "--" Then
        strTmp = ""
    End If
    strContent = strContent & strTmp & "$"
    
    '备注
    strContent = strContent & txtRemarks.Text & "$"
    
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
'检查合法性
    Dim i As Integer, strTip As String, blnTip As Boolean
    
    On Error GoTo errHand
    CheckValidity = False
    
    '"尖锐湿疣"、"生殖器疱疹"的病例分类只能为"临床诊断病例"和"实验室诊断病例"
    If ucVD(0).Checked = True Or ucVD(1).Checked = True Then
        If gbytDiseaseType <> 1 And gbytDiseaseType <> 2 Then
            strMsg = strMsg & "<尖锐湿疣>、<生殖器疱疹>的病例分类只能为<临床诊断病例>和<实验室诊断病例>！$"
        End If
    End If
    
    For i = 0 To 2
        If ucVD(i).Checked Then blnTip = True: Exit For
        If i = 2 Then strTip = "性病报告时，<监测性病>为必选项，请检查！$"
    Next
    
    '婚姻状态
    For i = 0 To 3
        If ucMarital(i).Checked = True Then blnTip = True: Exit For
        If i = 3 Then strTip = strTip & "性病报告时，<婚姻状况>为必选项，请检查！$"
    Next
    
    '文化程度
    For i = 0 To 6
        If ucEducation(i).Checked = True Then blnTip = True: Exit For
        If i = 6 Then strTip = strTip & "性病报告时，<文化程度>为必选项，请检查！$"
    Next
    
    '感染途径（接触史)
    For i = 0 To 8
        If ucWay(i).Checked = True Then blnTip = True: Exit For
        If i = 8 Then strTip = strTip & "性病报告时，<感染途径(接触史)>为必选项，请检查！$"
    Next
    
    If blnTip Then '存在选中某项报告项
        strMsg = strMsg & strTip
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

Public Sub SetEnterInfo(ByVal strDoctor As String, ByVal strDate As String)
    Dim strDateInfo() As String
    Dim strCurTime() As String
    strDateInfo = Split(Format(strDate, "yyyy-mm-dd"), "-")
    strCurTime = Split(Format(zlDatabase.Currentdate, "yyyy-mm-dd"), "-")
    txtDoctor.Text = strDoctor
    If UBound(strDateInfo) < 2 Then
        txtEnter(0).Text = strCurTime(0)
        txtEnter(1).Text = strCurTime(1)
        txtEnter(2).Text = strCurTime(2)
    Else
        txtEnter(0).Text = strDateInfo(0)
        txtEnter(1).Text = strDateInfo(1)
        txtEnter(2).Text = strDateInfo(2)
    End If
End Sub

Public Sub ClearEnterInfo()
    txtDoctor.Text = ""
    txtEnter(0).Text = ""
    txtEnter(1).Text = ""
    txtEnter(2).Text = ""
End Sub

Private Sub lblAttack_Click(Index As Integer)
    MView.Top = txtEnter(0).Top - MView.Height - 10
    MView.Left = txtEnter(0).Left
    MView.Visible = True
    Call MView.SetFocus
End Sub

Private Sub lblAttack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lblAttack(Index).MouseIcon = Image1.Picture
    lblAttack(Index).MousePointer = vbCustom
End Sub

Private Sub MView_DateClick(ByVal DateClicked As Date)
    txtEnter(0).Text = MView.Year
    txtEnter(1).Text = MView.Month
    txtEnter(2).Text = MView.Day
    MView.Visible = False
End Sub

Private Sub MView_LostFocus()
    MView.Visible = False
End Sub

Private Sub ucWay_Change(Index As Integer)
    
    Select Case Index
        '异性传播
        Case 0
            If ucWay(0).Checked = True Then
                ucWay1(0).Checked = True
            Else
                ucWay1(0).Checked = False
                ucWay1(1).Checked = False
            End If
        '血液传播
        Case 4
            If ucWay(4).Checked = True Then
                ucWay2(0).Checked = True
            Else
                ucWay2(0).Checked = False
                ucWay2(1).Checked = False
                ucWay2(2).Checked = False
            End If
        Case Else
            ucWay1(0).Checked = False
            ucWay1(1).Checked = False
            ucWay2(0).Checked = False
            ucWay2(1).Checked = False
            ucWay2(2).Checked = False

    End Select
End Sub

Private Sub ucWay1_Change(Index As Integer)
    Call ChangeUcWay
    If ucWay1(Index).Checked = True Then
        ucWay2(0).Checked = False
        ucWay2(1).Checked = False
        ucWay2(2).Checked = False
    End If
    ucWay(0).Checked = ucWay1(Index).Checked
End Sub

Private Sub ucWay2_Change(Index As Integer)
    Call ChangeUcWay
    If ucWay2(Index).Checked = True Then
        ucWay1(0).Checked = False
        ucWay1(1).Checked = False
    End If
    ucWay(4).Checked = ucWay2(Index).Checked
End Sub

Private Sub ChangeUcWay()
    Dim i As Integer
    For i = 0 To 8
        ucWay(i).Checked = False
        
    Next
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackColor = vbWhite
End Sub
