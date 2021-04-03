VERSION 5.00
Begin VB.Form frmIdentify贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   Icon            =   "frmIdentify贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt备注 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1380
      TabIndex        =   62
      Top             =   7080
      Width           =   7995
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "改密码(&M)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   64
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   63
      Top             =   510
      Width           =   1335
   End
   Begin VB.TextBox txt封锁信息 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1380
      TabIndex        =   60
      Top             =   6690
      Width           =   7995
   End
   Begin VB.Frame Frame2 
      Caption         =   "累计信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   180
      TabIndex        =   34
      Top             =   3720
      Width           =   9195
      Begin VB.TextBox txt普通门诊医疗补助结转可使用 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   58
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox txt普通门诊医疗补助起付线 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   56
         Top             =   1890
         Width           =   1965
      End
      Begin VB.TextBox txt普通门诊医疗补助起付标准 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   54
         Top             =   1500
         Width           =   1965
      End
      Begin VB.TextBox txt普通门诊医疗补助累计 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   52
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txt普通门诊医疗补助限额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   50
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox txt大额支付累计 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   48
         Top             =   330
         Width           =   1965
      End
      Begin VB.TextBox txt大额统筹限额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   46
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox txt统筹支付累计 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   44
         Top             =   1890
         Width           =   1965
      End
      Begin VB.TextBox txt基本统筹限额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   42
         Top             =   1500
         Width           =   1965
      End
      Begin VB.TextBox txt已支付起付线 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   40
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txt起付线 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   38
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox txt住院次数 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   36
         Top             =   330
         Width           =   1965
      End
      Begin VB.Label lbl普通门诊医疗补助结转可使用 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助结转可使用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4200
         TabIndex        =   57
         Top             =   2340
         Width           =   2730
      End
      Begin VB.Label lbl公务员门诊补助起付线 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助起付线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4620
         TabIndex        =   55
         Top             =   1950
         Width           =   2310
      End
      Begin VB.Label lbl公务员门诊补助起付标准 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助起付标准"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4410
         TabIndex        =   53
         Top             =   1560
         Width           =   2520
      End
      Begin VB.Label lbl公务员门诊补助累计 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助累计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4830
         TabIndex        =   51
         Top             =   1170
         Width           =   2100
      End
      Begin VB.Label lbl公务员门诊补助限额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊医疗补助限额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4830
         TabIndex        =   49
         Top             =   780
         Width           =   2100
      End
      Begin VB.Label lbl大额支付累计 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大额支付累计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5670
         TabIndex        =   47
         Top             =   390
         Width           =   1260
      End
      Begin VB.Label lbl大额统筹限额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大额统筹限额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   45
         Top             =   2340
         Width           =   1260
      End
      Begin VB.Label lbl统筹支付累计 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹支付累计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   43
         Top             =   1950
         Width           =   1260
      End
      Begin VB.Label lbl基本统筹限额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "基本统筹限额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   41
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label lbl已支付起付线 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "已支付起付线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   39
         Top             =   1170
         Width           =   1260
      End
      Begin VB.Label lbl起付线 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "起付线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1020
         TabIndex        =   37
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl住院次数 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   35
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   9195
      Begin VB.ComboBox cbo保险类别 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   1905
      End
      Begin VB.CheckBox chk生育标志 
         Caption         =   "生育标志"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txt缴费年度 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   33
         Top             =   3060
         Width           =   2595
      End
      Begin VB.TextBox txt帐户余额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   31
         Top             =   2670
         Width           =   2595
      End
      Begin VB.TextBox txt单位名称 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   29
         Top             =   2280
         Width           =   2595
      End
      Begin VB.TextBox txt单位编码 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   27
         Top             =   1890
         Width           =   1335
      End
      Begin VB.TextBox txt出生日期 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   25
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox txt身份证号 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   23
         Top             =   1110
         Width           =   2595
      End
      Begin VB.TextBox txt性别 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt姓名 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   19
         Top             =   330
         Width           =   2595
      End
      Begin VB.TextBox txt人员类别 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   15
         Top             =   2670
         Width           =   2595
      End
      Begin VB.TextBox txt医疗照顾人群 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   17
         Top             =   3060
         Width           =   2595
      End
      Begin VB.TextBox txt分中心编号 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   2280
         Width           =   2595
      End
      Begin VB.TextBox txt医保号 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1890
         Width           =   2595
      End
      Begin VB.TextBox txt密码 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1500
         Width           =   1245
      End
      Begin VB.ComboBox cbo支付类别 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   1905
      End
      Begin VB.TextBox txt卡号 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1110
         Width           =   2595
      End
      Begin VB.Label lbl保险类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "保险类别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   1
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lbl缴费年度 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "缴费年度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   32
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label lbl帐户余额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "帐户余额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   30
         Top             =   2730
         Width           =   840
      End
      Begin VB.Label lbl单位名称 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   28
         Top             =   2340
         Width           =   840
      End
      Begin VB.Label lbl单位编码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位编码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   26
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label lbl出生日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   24
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lbl身份证号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5460
         TabIndex        =   22
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label lbl性别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5880
         TabIndex        =   20
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5880
         TabIndex        =   18
         Top             =   390
         Width           =   420
      End
      Begin VB.Label lbl人员类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "人员类别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   14
         Top             =   2730
         Width           =   840
      End
      Begin VB.Label lbl医疗照顾人群 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医疗照顾人群"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   390
         TabIndex        =   16
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label lbl分中心编码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "分中心编码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   600
         TabIndex        =   12
         Top             =   2340
         Width           =   1050
      End
      Begin VB.Label lbl医保号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   10
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label lbl密码 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1230
         TabIndex        =   7
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label lbl支付类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "支付类别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   810
         TabIndex        =   3
         Top             =   780
         Width           =   840
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1230
         TabIndex        =   5
         Top             =   1170
         Width           =   420
      End
   End
   Begin VB.Label lbl备注 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   750
      TabIndex        =   61
      Top             =   7140
      Width           =   420
   End
   Begin VB.Label lbl封锁信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "封锁信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   330
      TabIndex        =   59
      Top             =   6750
      Width           =   840
   End
End
Attribute VB_Name = "frmIdentify贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mstr卡号 As String
Private mstr医保号 As String
Private mstr分中心编号 As String
Private mstr保险类别 As String
Private mstr密码 As String
Private mstr新密码 As String
Private mbln生育标志 As Boolean
Private mblnOK As Boolean
Private int门诊住院标志 As Integer   '门诊-0,住院-1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangePassword_Click()
    Dim strNewPass As String
    strNewPass = frm修改密码.ChangePassword("", Me.txt密码.Text, 40)
    If strNewPass <> "" Then mstr新密码 = strNewPass
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    If cmdOK.Enabled = False Then Exit Sub
    If Trim(txt卡号.Text) = "" Then
        MsgBox "未正确地刷卡,不能通过验证！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(txt医保号.Text) = "" Then
        MsgBox "未正确地刷卡,不能通过验证！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(mstr新密码) <> "" Then
        If 更改密码_贵阳市(txt卡号.Tag, txt密码.Text, mstr新密码) = False Then Exit Sub
        mstr密码 = mstr新密码
        mstr新密码 = ""
        txt密码.Text = mstr密码
    End If
    
    '2005.11.22,int门诊住院标志,住院强制选择保险类别
    If (int门诊住院标志 = 1 And cbo保险类别.ListIndex = 0) Then
       MsgBox "请选择保险类别！", vbInformation, gstrSysName
       cbo保险类别.SetFocus
       Exit Sub
    End If
    
    '有可能修改了密码，造成病人身份验证后返回的XML被破坏，再次调用读卡
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CARDDATA", txt卡号.Tag)            ' 磁卡数据
    Call InsertChild(mdomInput.documentElement, "PASSWORD", txt密码.Text)            ' 密码
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo支付类别.ItemData(Me.cbo支付类别.ListIndex))            ' 支付类别
 
    '2005.11.22,int门诊住院标志,医保返回
    If int门诊住院标志 = 0 Then
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo保险类别.ListIndex + 1)
    Else
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo保险类别.ListIndex)
    End If
    
    Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss"))           ' 开始时间
    '调用接口
    If CommServer("GETPSNINFO") = False Then Exit Sub
    
    'mstr卡号 = Trim(txt卡号.Text)
    '医保接口未返回卡号，以前的卡号字段改为保存磁卡数据，后面虚拟结算要用
    mstr卡号 = Me.txt卡号.Tag
    mstr医保号 = Trim(txt医保号.Text)
    mstr分中心编号 = Trim(txt分中心编号.Text)
    mstr密码 = Trim(txt密码.Text)
    mstr保险类别 = cbo保险类别.ListIndex + 1
    mbln生育标志 = (chk生育标志.Value = 1)
    
    '保存此病人的医保档案
'    医保号_IN IN 医保病人档案.医保号%TYPE,
'    住院次数_IN IN 医保病人档案.住院次数%TYPE,
'    起付线_IN IN 医保病人档案.起付线%TYPE,
'    已支付起付线_IN IN 医保病人档案.已支付起付线%TYPE,
'    基本统筹限额_IN IN 医保病人档案.基本统筹限额%TYPE,
'    统筹支付累计_IN IN 医保病人档案.统筹支付累计%TYPE,
'    大额统筹限额_IN IN 医保病人档案.大额统筹限额%TYPE,
'    大额支付累计_IN IN 医保病人档案.大额支付累计%TYPE,
'    公务员补助限额_IN IN 医保病人档案.公务员补助限额%TYPE,
'    公务员补助累计_IN IN 医保病人档案.公务员补助累计%TYPE,
'    公务员起付标准_IN IN 医保病人档案.公务员起付标准%TYPE,
'    公务员补助起付线_IN IN 医保病人档案.公务员补助起付线%TYPE,
'    参加75公务员补助_IN IN 医保病人档案.参加75公务员补助%TYPE)
    On Error GoTo errHand
    gstrSQL = "zl_医保病人档案_INSERT(" & _
        "'" & mstr医保号 & "'," & Val(txt住院次数.Text) & "," & Val(txt起付线.Text) & "," & Val(txt已支付起付线.Text) & "," & _
        "" & Val(txt基本统筹限额.Text) & "," & Val(txt统筹支付累计.Text) & "," & Val(txt大额统筹限额.Text) & "," & Val(txt大额支付累计.Text) & "," & _
        "" & Val(txt普通门诊医疗补助限额.Text) & "," & Val(txt普通门诊医疗补助累计.Text) & "," & Val(txt普通门诊医疗补助起付标准.Text) & "," & _
        "" & Val(txt普通门诊医疗补助起付线.Text) & ",'" & txt普通门诊医疗补助结转可使用.Text & "','" & txt备注.Text & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetIdentify(ByVal bytType As Byte, str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, _
    Optional ByRef bln生育标志 As Boolean = False) As Boolean
    mblnOK = False
    mstr密码 = ""
    mstr新密码 = ""
    mbytType = bytType
    
    frmIdentify贵阳.Show vbModal
    
    GetIdentify = mblnOK
    If mblnOK = True Then
        str卡号 = mstr卡号 & "^" & mstr保险类别
        str医保号 = mstr医保号
        str分中心编号 = mstr分中心编号
        str密码 = mstr密码
        bln生育标志 = mbln生育标志
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '2005.11.22,int门诊住院标志,加载cbo保险类别itemdata
    With cbo支付类别
        .Clear
        If mbytType = 0 Or mbytType = 3 Then
            int门诊住院标志 = 0
            .AddItem "普通门诊"
            .ItemData(.NewIndex) = 11
            .AddItem "特殊门诊"
            .ItemData(.NewIndex) = 18
            With cbo保险类别
                .Clear
                .AddItem "企业职工基本医疗保险"
                .AddItem "企业离休医疗保险"
                .AddItem "机关事业单位医疗保险"
                .AddItem "生育保险"
                .ListIndex = 0
            End With
            .ListIndex = 0
         Else
            int门诊住院标志 = 1
            .AddItem "普通住院"
            .ItemData(.NewIndex) = 31
            With cbo保险类别
                .Clear
                .AddItem ""
                .AddItem "企业职工基本医疗保险"
                .AddItem "企业离休医疗保险"
                .AddItem "机关事业单位医疗保险"
                .AddItem "生育保险"
                .ListIndex = 0
            End With
         End If
        .ListIndex = 0
    End With
End Sub

Private Sub txt密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim str封锁信息 As String
    
    If Trim(txt卡号.Text) = "" Then
        MsgBox "请刷卡！", vbInformation, gstrSysName
        txt卡号.SetFocus
        Exit Sub
    End If
    
    '2005.11.22,int门诊住院标志,住院强制选择保险类别
    If (int门诊住院标志 = 1 And cbo保险类别.ListIndex = 0) Then
       MsgBox "请选择保险类别！", vbInformation, gstrSysName
       cbo保险类别.SetFocus
       Exit Sub
    End If
    
    If InitXML = False Then Exit Sub
    
    '必须先修改密码
    If Trim(mstr新密码) <> "" Then
        If 更改密码_贵阳市(txt卡号.Text, mstr密码, mstr新密码) = False Then Exit Sub
        mstr密码 = mstr新密码
        mstr新密码 = ""
        txt密码.Text = mstr密码
    End If

    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CARDDATA", txt卡号.Text)            ' 磁卡数据
    Call InsertChild(mdomInput.documentElement, "PASSWORD", txt密码.Text)            ' 密码
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo支付类别.ItemData(Me.cbo支付类别.ListIndex))            ' 支付类别
        
    '2005.11.22,int门诊住院标志,医保返回
    If int门诊住院标志 = 0 Then
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo保险类别.ListIndex + 1)
    Else
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo保险类别.ListIndex)
    End If
    
    Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss"))           ' 开始时间
    '调用接口
    If CommServer("GETPSNINFO") = False Then Exit Sub
    
    '取得返回值
    '基本信息
    txt卡号.Tag = txt卡号.Text                    '保存卡内数据，以便更新密码时使用
    'txt卡号.Text = GetElemnetValue("CARDID")
    txt医保号.Text = GetElemnetValue("PERSONCODE")
    txt分中心编号.Text = GetElemnetValue("CENTERCODE")
    txt医疗照顾人群.Text = IIf(Val(GetElemnetValue("CAREPSNFLAG")) = 0, "否", "是")
    
    '2005.11.22,int门诊住院标志,住院必须自行选择保险类别,默认置空
    If int门诊住院标志 = 0 Then
        cbo保险类别.ListIndex = GetElemnetValue("INSURETYPE") - 1
    Else
'        cbo保险类别.ListIndex = 0
    End If
    
    txt人员类别.Text = GetElemnetValue("PERSONTYPE")
    txt人员类别.Text = Switch(txt人员类别.Text = "11", "在职", txt人员类别.Text = "21", "退休" _
                      , txt人员类别.Text = "32", "省属离休", txt人员类别.Text = "34", "市属离休", True, "其他")
    txt姓名.Text = GetElemnetValue("PERSONNAME")
    txt性别.Text = GetElemnetValue("SEX")
    txt性别.Text = Switch(txt性别.Text = "1", "男", txt性别.Text = "2", "女", txt性别.Text = "9", "其它", True, txt性别.Text)
    txt身份证号.Text = GetElemnetValue("PID")
    txt出生日期.Text = GetElemnetValue("BIRTHDAY")
    txt单位编码.Text = GetElemnetValue("DEPTCODE")
    txt单位名称.Text = GetElemnetValue("DEPTNAME")
    txt帐户余额.Text = GetElemnetValue("ACCTBALANCE")
    '累计信息
    txt住院次数.Text = GetElemnetValue("HOSPTIMES")
    txt起付线.Text = GetElemnetValue("STARTFEE")
    txt已支付起付线.Text = GetElemnetValue("STARTFEEPAID")
    txt基本统筹限额.Text = GetElemnetValue("FUND1LMT")
    txt统筹支付累计.Text = GetElemnetValue("FUND1PAID")
    txt大额统筹限额.Text = GetElemnetValue("FUND2LMT")
    txt大额支付累计.Text = GetElemnetValue("FUND2PAID")
    txt普通门诊医疗补助限额.Text = GetElemnetValue("FUND3LMT")
    txt普通门诊医疗补助累计.Text = GetElemnetValue("FUND3PAID")
    txt普通门诊医疗补助起付标准.Text = GetElemnetValue("STARTFEE2STD")
    txt普通门诊医疗补助起付线.Text = GetElemnetValue("STARTFEE2")
    txt普通门诊医疗补助结转可使用.Text = GetElemnetValue("FUND75BALANCE")
    txt备注.Text = GetElemnetValue("NOTE")
    txt封锁信息.Text = GetElemnetValue("LOCKINFO")

    cmdOK.Enabled = True
End Sub
