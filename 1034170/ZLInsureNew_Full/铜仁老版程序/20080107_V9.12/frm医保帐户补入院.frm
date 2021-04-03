VERSION 5.00
Begin VB.Form frm医保帐户补入院 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人补办医保入院登记"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frm医保帐户补入院.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "登记(&X)"
      Height          =   350
      Left            =   6090
      TabIndex        =   8
      Top             =   6015
      Width           =   1100
   End
   Begin VB.Frame fra费用信息 
      Caption         =   "【费用信息】"
      ForeColor       =   &H00C00000&
      Height          =   705
      Left            =   75
      TabIndex        =   27
      Top             =   1815
      Width           =   8745
      Begin VB.TextBox txt费用余额 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txt预交余额 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txt担保额 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7380
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txt担保人 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未结费用"
         Height          =   180
         Left            =   2370
         TabIndex        =   30
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交余额"
         Height          =   180
         Left            =   375
         TabIndex        =   28
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   180
         Left            =   6765
         TabIndex        =   34
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   180
         Left            =   4695
         TabIndex        =   32
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame fra基本信息 
      Caption         =   "【基本信息】"
      ForeColor       =   &H00C00000&
      Height          =   3345
      Left            =   75
      TabIndex        =   58
      Top             =   2580
      Width           =   8745
      Begin VB.TextBox txt医疗付款 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt出生日期 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   570
         Width           =   1140
      End
      Begin VB.TextBox txt联系人关系 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1890
         Width           =   2000
      End
      Begin VB.TextBox txt身份 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt职业 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt婚姻状况 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt国籍 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt学历 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt民族 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt出生地点 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   900
         Width           =   3225
      End
      Begin VB.TextBox txt家庭地址 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1230
         Width           =   3150
      End
      Begin VB.TextBox txt户口邮编 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1170
      End
      Begin VB.TextBox txt联系人姓名 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1170
      End
      Begin VB.TextBox txt联系人地址 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1890
         Width           =   3225
      End
      Begin VB.TextBox txt联系人电话 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2220
         Width           =   2000
      End
      Begin VB.TextBox txt工作单位 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2220
         Width           =   3225
      End
      Begin VB.TextBox txt单位电话 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2550
         Width           =   2000
      End
      Begin VB.TextBox txt单位邮编 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1170
      End
      Begin VB.TextBox txt单位开户行 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txt单位帐号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3225
      End
      Begin VB.TextBox txt家庭电话 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2000
      End
      Begin VB.TextBox txt身份证号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   900
         Width           =   3150
      End
      Begin VB.Label lbl医疗付款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗付款"
         Height          =   180
         Left            =   345
         TabIndex        =   80
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Left            =   6570
         TabIndex        =   79
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Left            =   4470
         TabIndex        =   78
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   345
         TabIndex        =   77
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份"
         Height          =   180
         Left            =   4830
         TabIndex        =   76
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Left            =   2685
         TabIndex        =   75
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   180
         Left            =   4830
         TabIndex        =   74
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   180
         Left            =   2685
         TabIndex        =   73
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历"
         Height          =   180
         Left            =   6930
         TabIndex        =   72
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lvl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Left            =   345
         TabIndex        =   71
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址"
         Height          =   180
         Left            =   345
         TabIndex        =   70
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl家庭电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Left            =   345
         TabIndex        =   69
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl户口邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口邮编"
         Height          =   180
         Left            =   4470
         TabIndex        =   68
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl联系人姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人姓名"
         Height          =   180
         Left            =   4290
         TabIndex        =   67
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label lbl联系人关系 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人关系"
         Height          =   180
         Left            =   165
         TabIndex        =   66
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl联系人地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人地址"
         Height          =   180
         Left            =   4290
         TabIndex        =   65
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl联系人电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人电话"
         Height          =   180
         Left            =   165
         TabIndex        =   64
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label lbl工作单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         Height          =   180
         Left            =   4470
         TabIndex        =   63
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label lbl单位电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   180
         Left            =   345
         TabIndex        =   62
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl单位邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Left            =   4470
         TabIndex        =   61
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl单位开户行 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位开户行"
         Height          =   180
         Left            =   165
         TabIndex        =   60
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label lbl单位帐号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位帐号"
         Height          =   180
         Left            =   4470
         TabIndex        =   59
         Top             =   2940
         Width           =   720
      End
   End
   Begin VB.Frame fra在院信息 
      Caption         =   "【住院信息】"
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   8730
      Begin VB.TextBox txt诊断 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1260
         Width           =   7320
      End
      Begin VB.CommandButton cmdYB 
         Caption         =   "验证(&V)"
         Height          =   285
         Left            =   6330
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "热键:F12(医保病人验证)"
         Top             =   240
         Width           =   1020
      End
      Begin VB.TextBox txt护理 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   885
         Width           =   1065
      End
      Begin VB.TextBox txt床号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   885
         Width           =   1170
      End
      Begin VB.TextBox txt入院时间 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txt科室 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txt医保号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt费别 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt病人ID 
         Height          =   300
         Left            =   1125
         TabIndex        =   2
         Top             =   225
         Width           =   1110
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   555
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院诊断"
         Height          =   180
         Left            =   330
         TabIndex        =   82
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护理"
         Height          =   180
         Left            =   4800
         TabIndex        =   23
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   2715
         TabIndex        =   21
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院时间"
         Height          =   180
         Left            =   6540
         TabIndex        =   25
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   690
         TabIndex        =   19
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Height          =   180
         Left            =   4620
         TabIndex        =   5
         Top             =   285
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   2535
         TabIndex        =   3
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   690
         TabIndex        =   11
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2715
         TabIndex        =   13
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4800
         TabIndex        =   15
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   6900
         TabIndex        =   17
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl病人ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   510
         TabIndex        =   1
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   780
      TabIndex        =   10
      Top             =   6015
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7380
      TabIndex        =   9
      Top             =   6015
      Width           =   1100
   End
End
Attribute VB_Name = "frm医保帐户补入院"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long '要修改或查看的病人ID
Private mlng主页ID As Long '要修改或查看的主页ID
Private mstr医保号 As String

Private Function ReadCard() As Boolean
'功能：读取指定病人信息,并显示在界面上
    Dim rstmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
        
    gstrSQL = "Select * From 病人信息 Where 病人ID=" & mlng病人ID
    rstmp.CursorLocation = adUseClient
    
    Call OpenRecordset(rstmp, Me.Caption)
    
    If rstmp.EOF Then Exit Function
    If rstmp.RecordCount <> 1 Then Exit Function
    
    '住院信息
    txt病人ID.Locked = True
    txt病人ID.Text = mlng病人ID
    txt病人ID.Locked = False
    
    txt姓名.Text = rstmp!姓名
    txt住院号.Text = IIf(IsNull(rstmp!住院号), "", rstmp!住院号)
    
    '基本信息
    txt性别.Text = IIf(IsNull(rstmp!性别), "", rstmp!性别)
    txt年龄.Text = IIf(IsNull(rstmp!年龄), "", rstmp!年龄)
    txt费别.Text = IIf(IsNull(rstmp!费别), "", rstmp!费别)
    txt医疗付款.Text = IIf(IsNull(rstmp!医疗付款方式), "", rstmp!医疗付款方式)
    txt国籍.Text = IIf(IsNull(rstmp!国籍), "", rstmp!国籍)
    txt民族.Text = IIf(IsNull(rstmp!民族), "", rstmp!民族)
    txt学历.Text = IIf(IsNull(rstmp!学历), "", rstmp!学历)
    txt婚姻状况.Text = IIf(IsNull(rstmp!婚姻状况), "", rstmp!婚姻状况)
    txt职业.Text = IIf(IsNull(rstmp!职业), "", rstmp!职业)
    txt身份.Text = IIf(IsNull(rstmp!身份), "", rstmp!身份)
    txt出生日期.Text = Format(IIf(IsNull(rstmp!出生日期), "", rstmp!出生日期), "yyyy-MM-dd")
    txt身份证号.Text = IIf(IsNull(rstmp!身份证号), "", rstmp!身份证号)
    txt出生地点.Text = IIf(IsNull(rstmp!出生地点), "", rstmp!出生地点)
    txt家庭地址.Text = IIf(IsNull(rstmp!家庭地址), "", rstmp!家庭地址)
    txt家庭电话.Text = IIf(IsNull(rstmp!家庭电话), "", rstmp!家庭电话)
    txt户口邮编.Text = IIf(IsNull(rstmp!户口邮编), "", rstmp!户口邮编)
    txt联系人姓名.Text = IIf(IsNull(rstmp!联系人姓名), "", rstmp!联系人姓名)
    txt联系人关系.Text = IIf(IsNull(rstmp!联系人关系), "", rstmp!联系人关系)
    txt联系人地址.Text = IIf(IsNull(rstmp!联系人地址), "", rstmp!联系人地址)
    txt联系人电话.Text = IIf(IsNull(rstmp!联系人电话), "", rstmp!联系人电话)
    txt工作单位.Text = IIf(IsNull(rstmp!工作单位), "", rstmp!工作单位)
    txt单位电话.Text = IIf(IsNull(rstmp!单位电话), "", rstmp!单位电话)
    txt单位邮编.Text = IIf(IsNull(rstmp!单位邮编), "", rstmp!单位邮编)
    txt单位开户行.Text = IIf(IsNull(rstmp!单位开户行), "", rstmp!单位开户行)
    txt单位帐号.Text = IIf(IsNull(rstmp!单位帐号), "", rstmp!单位帐号)
        
    '费用信息
    txt担保人.Text = IIf(IsNull(rstmp!担保人), "", rstmp!担保人)
    txt担保额.Text = Format(IIf(IsNull(rstmp!担保额), "", rstmp!担保额), "0.00")
    
    gstrSQL = "Select * From 病人余额 Where 性质=1 And 病人ID=" & mlng病人ID
    Call OpenRecordset(rstmp, Me.Caption)
    
    If Not rstmp.EOF Then
        txt费用余额.Text = Format(IIf(IsNull(rstmp!费用余额), 0, rstmp!费用余额), "0.00")
        txt预交余额.Text = Format(IIf(IsNull(rstmp!预交余额), 0, rstmp!预交余额), "0.00")
    End If
    
    
    '病人医保信息
    txt医保号.Text = ""
    mstr医保号 = ""
    
    
    '病案主页信息
    gstrSQL = "Select A.入院日期,A.出院病床,b.名称 as 入院科室,C.名称 as 护理等级" & _
              " From 病案主页 A,部门表 B,护理等级 C" & _
              " Where A.病人ID=" & mlng病人ID & " And A.主页ID=" & mlng主页ID & _
              "       and A.入院科室ID=B.ID and A.护理等级ID=C.序号(+) "
    Call OpenRecordset(rstmp, Me.Caption)
    
    txt科室.Text = rstmp!入院科室
    txt护理.Text = IIf(IsNull(rstmp!护理等级), "无", rstmp!护理等级)
    txt床号.Text = IIf(IsNull(rstmp!出院病床), "", rstmp!出院病床)
    txt入院时间.Text = Format(rstmp!入院日期, "yyyy-MM-dd HH:mm")
    
    '入院诊断
    gstrSQL = "Select 描述信息" & _
              " From 诊断情况" & _
              " Where 病人ID=" & mlng病人ID & " And 主页ID=" & mlng主页ID & " and 诊断类型=1 "
    Call OpenRecordset(rstmp, Me.Caption)
    If rstmp.EOF = False Then
        txt诊断.Text = NVL(rstmp("描述信息"))
    End If
    
    Dim objInsure As New clsInsure
    If objInsure.GetCapability(support必须录入入出诊断) = True Then
        txt诊断.Locked = False
        txt诊断.BackColor = txt病人ID.BackColor
    Else
        txt诊断.Locked = True
        txt诊断.BackColor = txt住院号.BackColor
    End If
    
    ReadCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
'补办入院登记
    Dim clsInsure As New clsInsure
    
    If mlng病人ID = 0 Then
        MsgBox "请先确定等补办入院登记的病人。", vbInformation, gstrSysName
        txt病人ID.SetFocus
        Exit Sub
    End If
    If mstr医保号 = "" Then
        MsgBox "请先验证该病人是否可以进行医保入院。", vbInformation, gstrSysName
        cmdYB.SetFocus
        Exit Sub
    End If
    If txt诊断.Locked = False And txt诊断.Text = "" Then
        MsgBox "请填写入院诊断。", vbInformation, gstrSysName
        txt诊断.SetFocus
        Exit Sub
    End If
    If zlCommFun.StrIsValid(txt诊断.Text, txt诊断.MaxLength, txt诊断.hwnd, "入院诊断") = False Then
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    gstrSQL = "zl_病案主页_补办医保入院(" & mlng病人ID & "," & mlng主页ID & "," & gintInsure & ",'" & txt诊断 & "')"
    ExecuteProcedure Me.Caption
    
    If clsInsure.ComeInSwap(mlng病人ID, mlng主页ID, mstr医保号) = False Then
        '登记失败
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    
    gcnOracle.CommitTrans
    MsgBox "病人" & txt姓名.Text & "补办医保入院成功！" & IIf(gintInsure > 900, vbCrLf & "病人费用明细中医保数据已经按医保规则重算。", "") _
        , vbInformation, gstrSysName
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdYB_Click()
'验证医保病人身份
    Dim lng病人ID As Long
    Dim strYBPati As String
    Dim clsInsure As New clsInsure
    Dim arr信息 As Variant
    
    If mlng病人ID = 0 Then
        MsgBox "请先确定等补办入院登记的病人。", vbInformation, gstrSysName
        txt病人ID.SetFocus
        Exit Sub
    End If
    lng病人ID = mlng病人ID
    strYBPati = clsInsure.Identify(1, lng病人ID)
    If lng病人ID <> 0 Then mlng病人ID = lng病人ID
    
    arr信息 = Split(strYBPati, ";")
    '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID,...
    If UBound(arr信息) >= 8 Then
        txt医保号.Text = arr信息(1)
        mstr医保号 = txt医保号.Text
        
        txt姓名.Text = arr信息(3)
        txt性别.Text = arr信息(4)
        txt出生日期.Text = arr信息(5)
        txt身份证号.Text = arr信息(6)
        
        cmdOK.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdYB_Click
    End If
End Sub

Private Sub Form_Load()
    mlng病人ID = 0
    mlng主页ID = 0
End Sub

Private Sub txt病人ID_Change()
    If txt病人ID.Locked = False Then
        mlng病人ID = 0
        mlng主页ID = 0
    End If
End Sub

Private Sub txt病人ID_GotFocus()
    zlControl.TxtSelAll txt病人ID
End Sub

Private Sub txt病人ID_KeyPress(KeyAscii As Integer)
    Dim lng病人ID  As Long
    
    '转换成大写(汉字不可处理)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If InStr("0123456789", Chr(KeyAscii)) = 0 And (txt病人ID.Text = "" Or txt病人ID.SelLength = Len(txt病人ID.Text)) Then
        txt病人ID.MaxLength = 15
    End If
    
    If Len(Trim(Me.txt病人ID.Text)) = 0 And KeyAscii = 13 Then
        If frm医保病人选择.Get病人(lng病人ID) = True Then
            txt病人ID.Text = "A" & lng病人ID
        End If
    End If
    Me.Refresh
    
    '刷卡完毕或输入号码后回车
    If (KeyAscii = 13 And Trim(txt病人ID.Text) <> "") Then
        If Val(txt病人ID.Text) = mlng病人ID And mlng病人ID > 0 Then
            If mstr医保号 = "" Then
                cmdYB.SetFocus
            Else
                cmdOK.SetFocus
            End If
            Exit Sub
        End If
        
        If KeyAscii <> 13 Then
            txt病人ID.Text = txt病人ID.Text & Chr(KeyAscii)
            txt病人ID.SelStart = Len(txt病人ID.Text)
        End If
        KeyAscii = 0
        
        If Not GetPatient() Then
            MsgBox "没有发现该病人的住院信息,请重新输入！", vbInformation, gstrSysName
            txt病人ID.Text = ""
            txt病人ID.SetFocus
            Exit Sub
        Else
            Call ReadCard
            cmdYB.SetFocus
        End If
    End If

End Sub

Private Function GetPatient() As Boolean
'功能：读取病人信息
'返回:是否读取成功,成功时rsInfo中包含病人信息,失败时rsInfo=Close
    Dim rsInfo As New ADODB.Recordset
    Dim strCode As String
    
    strCode = Trim(txt病人ID.Text)
    On Error GoTo errH
    
    If (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID And Nvl(A.住院次数,0)=C.主页ID And A.病人ID=" & Val(Mid(strCode, 2)) & _
            "       and C.险类 is null and C.出院日期 is null"
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID And Nvl(A.住院次数,0)=C.主页ID And A.住院号=" & Val(Mid(strCode, 2)) & _
            "       and C.险类 is null and C.出院日期 is null"
    Else '当作姓名
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID And Nvl(A.住院次数,0)=C.主页ID And A.姓名='" & strCode & _
            "'       and C.险类 is null and C.出院日期 is null"
    End If
    
    rsInfo.CursorLocation = adUseClient
    Call OpenRecordset(rsInfo, Me.Caption)
    
    '读取失败
    If rsInfo.EOF Then
        Exit Function
    End If
        
    mlng病人ID = rsInfo("病人ID")
    mlng主页ID = rsInfo("主页ID")
    
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


