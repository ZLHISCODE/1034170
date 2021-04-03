VERSION 5.00
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#2.1#0"; "ZlPatiAddress.ocx"
Begin VB.Form frmEditPati 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人信息修改"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "frmEditPati.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   480
      TabIndex        =   64
      Top             =   8280
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7800
      TabIndex        =   63
      Top             =   8280
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   120
      TabIndex        =   65
      Top             =   0
      Width           =   8955
      Begin VB.CheckBox chk陪伴 
         Caption         =   "是否陪伴"
         Height          =   195
         Left            =   7605
         TabIndex        =   6
         Top             =   660
         Width           =   1020
      End
      Begin VB.TextBox txt护理 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txt入院时间 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   7215
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1530
      End
      Begin VB.TextBox txt等级 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   3300
      End
      Begin VB.TextBox txt床位 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox txt科室 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   3300
      End
      Begin VB.TextBox txt性别 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txt姓名 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   3255
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt住院号 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护理等级"
         Height          =   180
         Left            =   4965
         TabIndex        =   74
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院时间"
         Height          =   180
         Left            =   6450
         TabIndex        =   73
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位等级"
         Height          =   180
         Left            =   390
         TabIndex        =   72
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   5325
         TabIndex        =   71
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前科室"
         Height          =   180
         Left            =   390
         TabIndex        =   70
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   5325
         TabIndex        =   69
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   2850
         TabIndex        =   68
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   570
         TabIndex        =   60
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   6615
      Left            =   120
      TabIndex        =   59
      Top             =   1560
      Width           =   8955
      Begin VB.ComboBox cbo国籍 
         Height          =   300
         Left            =   7290
         TabIndex        =   111
         Top             =   240
         Width           =   1500
      End
      Begin VB.ComboBox cboIDNumber 
         Height          =   300
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   975
         Width           =   1695
      End
      Begin VB.ComboBox cbo学历 
         Height          =   300
         ItemData        =   "frmEditPati.frx":058A
         Left            =   4470
         List            =   "frmEditPati.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   630
         Width           =   1455
      End
      Begin VB.TextBox txt家庭地址邮编 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4230
         MaxLength       =   6
         TabIndex        =   30
         Top             =   2050
         Width           =   1695
      End
      Begin VB.TextBox txt联系人身份证号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1155
         MaxLength       =   18
         TabIndex        =   46
         Top             =   3960
         Width           =   4785
      End
      Begin VB.CommandButton cmd户口地址 
         Caption         =   "…"
         Height          =   240
         Left            =   5640
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "热键:F3"
         Top             =   2430
         Width           =   285
      End
      Begin VB.TextBox txt户口地址邮编 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7290
         MaxLength       =   6
         TabIndex        =   37
         Top             =   2400
         Width           =   1500
      End
      Begin VB.CommandButton cmd籍贯 
         Caption         =   "…"
         Height          =   240
         Left            =   8505
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "热键：F3"
         Top             =   2070
         Width           =   255
      End
      Begin VB.TextBox txt医疗小组 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1155
         MaxLength       =   6
         TabIndex        =   50
         Top             =   4725
         Width           =   1500
      End
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   0
         TabIndex        =   101
         Top             =   5475
         Width           =   8685
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   0
         TabIndex        =   100
         Top             =   5040
         Width           =   8685
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   0
         TabIndex        =   98
         Top             =   4665
         Width           =   8685
      End
      Begin VB.TextBox txt备注 
         Height          =   300
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   58
         Top             =   6225
         Width           =   7590
      End
      Begin VB.CommandButton cmd区域 
         Caption         =   "…"
         Height          =   240
         Left            =   8505
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "热键：F3"
         Top             =   1020
         Width           =   255
      End
      Begin VB.ComboBox cbo入院方式 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1320
         Width           =   1500
      End
      Begin VB.ComboBox cbo病人类型 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   3165
         Width           =   1500
      End
      Begin VB.ComboBox cbo入院属性 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1700
         Width           =   1500
      End
      Begin VB.ComboBox cbo职业 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   630
         Width           =   1500
      End
      Begin VB.ComboBox cbo联系人关系 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   3585
         Width           =   1500
      End
      Begin VB.ComboBox cbo住院医师 
         Height          =   300
         Left            =   4230
         Style           =   2  'Dropdown List
         TabIndex        =   51
         ToolTipText     =   "经治医师"
         Top             =   4725
         Width           =   1500
      End
      Begin VB.ComboBox cbo主治医师 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   4725
         Width           =   1500
      End
      Begin VB.ComboBox cbo主任医师 
         Height          =   300
         Left            =   7285
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   5145
         Width           =   1500
      End
      Begin VB.TextBox txt身份证号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1155
         TabIndex        =   16
         Top             =   975
         Width           =   2085
      End
      Begin VB.CommandButton cmd出生地点 
         Caption         =   "…"
         Height          =   240
         Left            =   5625
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "热键：F3"
         Top             =   1350
         Width           =   300
      End
      Begin VB.ComboBox cbo年龄单位 
         Height          =   300
         Left            =   6045
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   580
      End
      Begin VB.TextBox txt中医诊断 
         Height          =   300
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   57
         Top             =   5880
         Width           =   5445
      End
      Begin VB.CommandButton cmd联系人地址 
         Caption         =   "…"
         Height          =   240
         Left            =   8415
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "热键:F3"
         Top             =   4350
         Width           =   300
      End
      Begin VB.TextBox txt入院诊断 
         Height          =   300
         Left            =   1155
         MaxLength       =   200
         TabIndex        =   56
         Top             =   5535
         Width           =   5445
      End
      Begin VB.TextBox txt联系人地址 
         Height          =   300
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   47
         Top             =   4320
         Width           =   7590
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   5
         Left            =   1155
         TabIndex        =   49
         Tag             =   "联系人地址"
         Top             =   4320
         Visible         =   0   'False
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Items           =   5
         MaxLength       =   100
      End
      Begin VB.ComboBox cbo费别 
         Height          =   300
         Left            =   3150
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt年龄 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   900
      End
      Begin VB.ComboBox cbo婚姻状况 
         Height          =   300
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   630
         Width           =   1455
      End
      Begin VB.ComboBox cbo责任护士 
         Height          =   300
         Left            =   4230
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   5145
         Width           =   1500
      End
      Begin VB.TextBox txt联系人电话 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4230
         MaxLength       =   20
         TabIndex        =   44
         Top             =   3585
         Width           =   1695
      End
      Begin VB.TextBox txt联系人姓名 
         Height          =   300
         Left            =   1155
         MaxLength       =   64
         TabIndex        =   43
         Top             =   3585
         Width           =   1695
      End
      Begin VB.TextBox txt家庭电话 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1155
         MaxLength       =   20
         TabIndex        =   29
         Top             =   2050
         Width           =   1695
      End
      Begin VB.CommandButton cmd家庭地址 
         Caption         =   "…"
         Height          =   240
         Left            =   5640
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "热键:F3"
         Top             =   1730
         Width           =   285
      End
      Begin VB.TextBox txt单位电话 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4230
         MaxLength       =   20
         TabIndex        =   41
         Top             =   3165
         Width           =   1695
      End
      Begin VB.CommandButton cmd单位地址 
         Caption         =   "…"
         Height          =   240
         Left            =   8475
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "热键:F3"
         Top             =   2820
         Width           =   285
      End
      Begin VB.TextBox txt单位地址 
         Height          =   300
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   38
         Top             =   2790
         Width           =   7635
      End
      Begin VB.ComboBox cbo病况 
         Height          =   300
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cbo门诊医师 
         Height          =   300
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   5145
         Width           =   1500
      End
      Begin VB.TextBox txt出生地点 
         Height          =   300
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   21
         Top             =   1320
         Width           =   4785
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   1
         Left            =   1155
         TabIndex        =   23
         Tag             =   "出生地点"
         Top             =   1320
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Items           =   3
         MaxLength       =   100
      End
      Begin VB.TextBox txt单位邮编 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1155
         MaxLength       =   6
         TabIndex        =   40
         Top             =   3165
         Width           =   1695
      End
      Begin VB.CheckBox chk再入院 
         Caption         =   "再入院"
         Height          =   255
         Left            =   5070
         TabIndex        =   18
         ToolTipText     =   "再次入住相同诊疗科目编码的临床科室"
         Top             =   998
         Width           =   855
      End
      Begin VB.TextBox txt区域 
         Height          =   300
         Left            =   7290
         MaxLength       =   50
         TabIndex        =   19
         Top             =   990
         Width           =   1500
      End
      Begin VB.TextBox txt籍贯 
         Height          =   300
         Left            =   6570
         MaxLength       =   50
         TabIndex        =   31
         Top             =   2055
         Width           =   2220
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   2
         Left            =   6570
         TabIndex        =   33
         Tag             =   "籍贯"
         Top             =   2050
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Items           =   2
         MaxLength       =   100
      End
      Begin VB.TextBox txt户口地址 
         Height          =   300
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   34
         Top             =   2400
         Width           =   4785
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   4
         Left            =   1155
         TabIndex        =   36
         Tag             =   "户口地址"
         Top             =   2400
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Items           =   5
         MaxLength       =   100
      End
      Begin VB.TextBox txt家庭地址 
         Height          =   300
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   25
         Top             =   1700
         Width           =   4785
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   3
         Left            =   1155
         TabIndex        =   27
         Tag             =   "现住址"
         Top             =   1700
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Items           =   5
         MaxLength       =   100
      End
      Begin VB.Label Label32 
         Caption         =   "国籍"
         Height          =   180
         Left            =   6840
         TabIndex        =   112
         Top             =   300
         Width           =   375
      End
      Begin VB.Label lbl联系人身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人身份证"
         Height          =   180
         Left            =   45
         TabIndex        =   106
         Top             =   4005
         Width           =   1080
      End
      Begin VB.Label lbl户口地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口地址"
         Height          =   180
         Left            =   405
         TabIndex        =   104
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lbl户口地址邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口地址邮编"
         Height          =   180
         Left            =   6135
         TabIndex        =   103
         Top             =   2460
         Width           =   1080
      End
      Begin VB.Label lbl籍贯 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "籍贯"
         Height          =   180
         Left            =   6135
         TabIndex        =   102
         Top             =   2110
         Width           =   360
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗小组"
         Height          =   180
         Left            =   405
         TabIndex        =   99
         Top             =   4785
         Width           =   720
      End
      Begin VB.Label lbl备注 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   180
         Left            =   765
         TabIndex        =   97
         Top             =   6285
         Width           =   360
      End
      Begin VB.Label lbl入院方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院方式"
         Height          =   180
         Left            =   6495
         TabIndex        =   96
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lblPatiColor 
         Height          =   255
         Left            =   7815
         TabIndex        =   95
         Top             =   3180
         Width           =   225
      End
      Begin VB.Label lblPatiType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人类型"
         Height          =   180
         Left            =   6495
         TabIndex        =   94
         Top             =   3225
         Width           =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院属性"
         Height          =   180
         Left            =   6495
         TabIndex        =   93
         Top             =   1760
         Width           =   720
      End
      Begin VB.Label lbl区域 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "区域"
         Height          =   180
         Left            =   6855
         TabIndex        =   92
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Left            =   405
         TabIndex        =   91
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   405
         TabIndex        =   90
         Top             =   1035
         Width           =   720
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主任(副主任)医师"
         Height          =   180
         Left            =   5775
         TabIndex        =   89
         Top             =   5205
         Width           =   1440
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主治医师"
         Height          =   180
         Left            =   6495
         TabIndex        =   88
         Top             =   4785
         Width           =   720
      End
      Begin VB.Label lbl中医诊断 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中医诊断"
         Height          =   180
         Left            =   405
         TabIndex        =   61
         Top             =   5940
         Width           =   720
      End
      Begin VB.Label lbl入院诊断 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院诊断"
         Height          =   180
         Left            =   405
         TabIndex        =   66
         Top             =   5595
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人地址"
         Height          =   180
         Left            =   225
         TabIndex        =   85
         Top             =   4380
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   2730
         TabIndex        =   67
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4695
         TabIndex        =   84
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Left            =   405
         TabIndex        =   87
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院医师"
         Height          =   180
         Left            =   3450
         TabIndex        =   81
         Top             =   4785
         Width           =   720
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "责任护士"
         Height          =   180
         Left            =   3450
         TabIndex        =   79
         Top             =   5205
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电话"
         Height          =   180
         Left            =   3810
         TabIndex        =   86
         Top             =   3645
         Width           =   360
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "关系"
         Height          =   180
         Left            =   6855
         TabIndex        =   105
         Top             =   3645
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人"
         Height          =   180
         Left            =   585
         TabIndex        =   76
         Top             =   3645
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址邮编"
         Height          =   180
         Left            =   3090
         TabIndex        =   83
         Top             =   2115
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Left            =   405
         TabIndex        =   82
         Top             =   2110
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "现住址"
         Height          =   180
         Left            =   585
         TabIndex        =   107
         Top             =   1755
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电话"
         Height          =   180
         Left            =   3810
         TabIndex        =   108
         Top             =   3225
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         Height          =   180
         Left            =   405
         TabIndex        =   78
         Top             =   2850
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历"
         Height          =   180
         Left            =   4050
         TabIndex        =   77
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前病况"
         Height          =   180
         Left            =   405
         TabIndex        =   109
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Left            =   6855
         TabIndex        =   75
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊医师"
         Height          =   180
         Left            =   405
         TabIndex        =   80
         Top             =   5205
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Left            =   405
         TabIndex        =   110
         Top             =   3225
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6495
      TabIndex        =   62
      Top             =   8280
      Width           =   1100
   End
End
Attribute VB_Name = "frmEditPati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrPrivs As String
Public mlngUnit As Long, mstrUnit As String
Public mlng病人ID As Long, mlng主页ID As Long

Private mrsPati As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML
Private mstrPatiPlus    As String     '从表信息:信息名1:信息值1,信息名2:信息值2
Private mblnEMPI As Boolean       'T-返回EMPI平台病人,F-未返回EMPI平台病人
Private mstrBirthDay As String

Private Sub cbo病人类型_Click()
    If cbo病人类型.ListCount > 0 And cbo病人类型.ListIndex <> -1 Then
        lblPatiColor.BackColor = zlDatabase.GetPatiColor(NeedName(cbo病人类型.Text))
    End If
End Sub

Private Sub cbo病人类型_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo病人类型.hWnd, zlControl.CboMatchIndex(cbo病人类型.hWnd, KeyAscii))
End Sub

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo费别.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo费别.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo费别.ListIndex = lngIdx
End Sub

Private Sub cbo病况_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo病况.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo病况.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo病况.ListIndex = lngIdx
End Sub

Private Sub cbo婚姻状况_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo婚姻状况.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo婚姻状况.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo婚姻状况.ListIndex = lngIdx
End Sub

Private Sub cbo年龄单位_LostFocus()
    If Not CheckOldData(txt年龄, cbo年龄单位) Then Exit Sub
End Sub

Private Sub cbo主治医师_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo主治医师.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = MatchIndex(cbo主治医师.hWnd, KeyAscii)
        If lngIdx <> -2 Then cbo主治医师.ListIndex = lngIdx
    ElseIf cbo主治医师.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo住院医师_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo住院医师.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("医生", "主任医师,副主任医师,主治医师,医师,医士", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo住院医师.ListCount - 1
                If cbo住院医师.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo住院医师.ListIndex = i: Exit Sub
                End If
            Next
            cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo住院医师.ListCount - 1
            cbo住院医师.ListIndex = cbo住院医师.NewIndex
            cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!上级ID
        Else
            cbo住院医师.ListIndex = -1
        End If
    End If
End Sub
Private Sub cbo主治医师_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo主治医师.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("医生", "主任医师,副主任医师,主治医师", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo主治医师.ListCount - 1
                If cbo主治医师.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo主治医师.ListIndex = i: Exit Sub
                End If
            Next
            cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo主治医师.ListCount - 1
            cbo主治医师.ListIndex = cbo主治医师.NewIndex
            cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
        Else
            cbo主治医师.ListIndex = -1
        End If
    End If
End Sub
Private Sub cbo主任医师_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo主任医师.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("医生", "主任医师,副主任医师", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo主任医师.ListCount - 1
                If cbo主任医师.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo主任医师.ListIndex = i: Exit Sub
                End If
            Next
            cbo主任医师.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo主任医师.ListCount - 1
            cbo主任医师.ListIndex = cbo主任医师.NewIndex
            cbo主任医师.ItemData(cbo主任医师.NewIndex) = rsTmp!ID
        Else
            cbo主任医师.ListIndex = -1
        End If
    End If
End Sub

Private Sub cbo责任护士_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo责任护士.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("护士", "", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo责任护士.ListCount - 1
                If cbo责任护士.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo责任护士.ListIndex = i: Exit Sub
                End If
            Next
            cbo责任护士.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo责任护士.ListCount - 1
            cbo责任护士.ListIndex = cbo责任护士.NewIndex
            cbo责任护士.ItemData(cbo责任护士.NewIndex) = rsTmp!ID
        Else
            cbo责任护士.ListIndex = -1
        End If
    End If
End Sub

Private Sub cbo联系人关系_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo联系人关系.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo联系人关系.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo联系人关系.ListIndex = lngIdx
End Sub

Private Sub cbo门诊医师_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo门诊医师.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo门诊医师.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo门诊医师.ListIndex = lngIdx
End Sub

Private Sub cbo学历_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo学历.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo学历.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo学历.ListIndex = lngIdx
End Sub

Private Sub cbo责任护士_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo责任护士.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo责任护士.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo责任护士.ListIndex = lngIdx
End Sub

Private Sub cbo职业_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo职业.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo职业.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo职业.ListIndex = lngIdx
End Sub

Private Sub cbo住院医师_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo住院医师.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo住院医师.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo住院医师.ListIndex = lngIdx
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, "frmHosReg"
End Sub



Private Sub cmdOK_Click()
    Dim strSQL As String, strSQL_Recalc As String, blnTrans As Boolean
    Dim lng西医疾病ID As Long, lng中医疾病ID As Long
    Dim lng西医诊断ID As Long, lng中医诊断ID As Long, str年龄 As String
    Dim arrSQL() As String, i As Integer
    Dim lngTmp As Long
    Dim strBeginDate As String, strEndDate As String
    Dim str性别 As String, strAge As String, str出生日期 As String, strErrInfo As String
    Dim bln基本信息调整 As Boolean, blnMod As Boolean
    Dim strMsg As String
    Dim arrTmp As Variant
    
    If cbo费别.ListIndex = -1 Then
        MsgBox "请确定病人的费别！", vbInformation, gstrSysName
        cbo费别.SetFocus: Exit Sub
    End If
    
    '费别适用科室
    If Not Check费别适用科室(NeedName(cbo费别.Text), Val(txt科室.Tag)) Then
        MsgBox "当前费别对病人科室不适用,请重新选择费别!", vbInformation, gstrSysName
        cbo费别.SetFocus: Exit Sub
    End If
    
    If cbo国籍.ListIndex = -1 Then
        MsgBox "必须确定病人国籍！", vbInformation, gstrSysName
        cbo国籍.SetFocus: Exit Sub
    End If
    
    '入院诊断
    If Not CheckLen(txt入院诊断, txt入院诊断.MaxLength) Then Exit Sub
    If Not CheckLen(txt中医诊断, txt中医诊断.MaxLength) Then Exit Sub
    If Not IsNull(mrsPati!险类) Then
        If gclsInsure.GetCapability(support必须录入入出诊断, mlng病人ID, mrsPati!险类) Then
            If txt入院诊断.Text = "" Then
                MsgBox "请填写该病人的入院诊断！", vbInformation, gstrSysName
                txt入院诊断.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If gbln启用结构化地址 Then
        For i = PatiAddress.LBound To PatiAddress.UBound
            If Trim(PatiAddress(i).Value) <> "" And PatiAddress(i).CheckNullValue() <> "" Then
                MsgBox "病人的" & PatiAddress(i).Tag & "录入不完整,请重新录入。", vbInformation, gstrSysName
                If PatiAddress(i).Enabled And PatiAddress(i).Visible = True Then PatiAddress(i).SetFocus
                Exit Sub
            End If
        Next
    End If
    
    If Not CheckTextLength("年龄", txt年龄) Then Exit Sub
    If Not CheckOldData(txt年龄, cbo年龄单位) Then Exit Sub
    If Not CheckLen(txt出生地点, txt出生地点.MaxLength) Then Exit Sub
    If Not CheckLen(txt户口地址, txt户口地址.MaxLength) Then Exit Sub
    If Not CheckLen(txt家庭地址, txt家庭地址.MaxLength) Then Exit Sub
    If Not CheckLen(txt联系人姓名, txt联系人姓名.MaxLength) Then Exit Sub
    If Not CheckLen(txt联系人地址, txt联系人地址.MaxLength) Then Exit Sub
    If Not CheckLen(txt单位地址, txt单位地址.MaxLength) Then Exit Sub
    If Trim(zlCommFun.GetNeedName(cbo国籍.Text)) = "中国" Then
        If Not CheckLen(txt身份证号, 18) Then Exit Sub
    End If
    If Not CheckLen(txt联系人身份证号, 18) Then Exit Sub
    '问题27351 by lesfeng 2010-01-12
    If Not CheckLen(txt备注, txt备注.MaxLength) Then Exit Sub
        
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    
    '--46119,刘鹏飞,2012-08-16,检查输入的身份证性别是否和病人性别一致
    If Trim(zlCommFun.GetNeedName(cbo国籍.Text)) = "中国" Then
        lngTmp = LenB(StrConv(Trim(txt身份证号.Text), vbFromUnicode))
        If lngTmp > 0 Then
            If CreatePublicPatient() Then
                If gobjPublicPatient.CheckPatiIdcard(Trim(txt身份证号.Text), str出生日期, strAge, str性别, strErrInfo, CDate(txt入院时间.Text)) Then
                    '同一身份证只能对应一个建档病人非新建档病人时,检查身份证号
                    If gblnPatiByID And txt身份证号.Tag <> Trim(txt身份证号.Text) Then
                        If gobjPublicPatient.CheckPatiExistByID(Trim(txt身份证号.Text), mlng病人ID) Then
                            MsgBox "已存在身份证号为【" & Trim(txt身份证号.Text) & "】的建档病人，禁止修改身份证号！", vbInformation, gstrSysName
                            If CanFocus(txt身份证号) Then txt身份证号.SetFocus: Exit Sub
                        End If
                    End If
                    '有无基本信息调整权限
                    bln基本信息调整 = InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";基本信息调整;") > 0
                    If Format(mstrBirthDay, "HH:MM") <> "00:00" Then
                        str出生日期 = str出生日期 & " " & Format(mstrBirthDay, "HH:MM")
                    End If
                    '年龄
                    strMsg = ""
                    If str年龄 <> strAge Then
                        strMsg = "身份证号码中的年龄[" & strAge & "]" & "和病人年龄[" & str年龄 & "]不一致"
                        If str年龄 Like "*小时*分钟" Or str年龄 Like "*分钟" Or str年龄 Like "*天*小时" Or str年龄 Like "*小时" Then
                            strAge = str年龄
                        End If
                    ElseIf InStr(txt性别.Text, str性别) = 0 Then '性别
                        strMsg = "身份证号码中的性别[" & str性别 & "]和病人性别[" & txt性别.Text & "]不一致"
                    End If
                    If strMsg <> "" Then
                        If MsgBox(strMsg & ",是否继续？" & vbCrLf & IIf(bln基本信息调整, "选【是】,用身份证的信息替换病人的信息。", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            If CanFocus(txt身份证号) = True Then txt身份证号.SetFocus: Exit Sub
                        Else
                            blnMod = True
                        End If
                    End If
                Else
                    MsgBox strErrInfo, vbInformation + vbOKOnly, gstrSysName
                    If CanFocus(txt身份证号) = True Then txt身份证号.SetFocus: Exit Sub
                End If
            End If
        End If
    End If
    
    mstrPatiPlus = ""
    If zlCommFun.GetNeedName(cbo国籍.Text) = "中国" Then
        mstrPatiPlus = mstrPatiPlus & "," & "身份证号状态:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
        mstrPatiPlus = mstrPatiPlus & "," & "外籍身份证号:"
    Else
        If Trim(txt身份证号.Text) <> "" Then
            mstrPatiPlus = mstrPatiPlus & "," & "外籍身份证号:" & Trim(txt身份证号.Text)
            mstrPatiPlus = mstrPatiPlus & "," & "身份证号状态:"
            txt身份证号.Text = ""
        Else
            mstrPatiPlus = mstrPatiPlus & "," & "身份证号状态:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
            mstrPatiPlus = mstrPatiPlus & "," & "外籍身份证号:"
        End If
    End If
    If mstrPatiPlus <> "" Then mstrPatiPlus = Mid(mstrPatiPlus, 2)
    
    If InStr(1, txt入院诊断.Tag, ";") <= 0 Then
        lng西医疾病ID = Val(txt入院诊断.Tag)
    Else
        lng西医诊断ID = Val(txt入院诊断.Tag)
    End If
    If InStr(1, txt中医诊断.Tag, ";") <= 0 Then
        lng中医疾病ID = Val(txt中医诊断.Tag)
    Else
        lng中医诊断ID = Val(txt中医诊断.Tag)
    End If
    '问题27351 by lesfeng 2010-01-12
    '问题24463 by lesfeng 2010-03-22 增加陪伴
    '问题51167,刘鹏飞,2012-07-09,增加"联系人身份证号"
    strSQL = "zl_住院病案主页_Update(" & mlng病人ID & "," & mlng主页ID & ",'" & str年龄 & "'," & _
        "'" & NeedName(cbo费别.Text) & "','" & NeedName(cbo婚姻状况.Text) & "','" & NeedName(cbo学历.Text) & "'," & _
        "'" & NeedName(cbo职业.Text) & "','" & NeedName(cbo病况.Text) & "','" & txt单位地址.Text & "'," & _
        Val(txt单位地址.Tag) & ",'" & txt单位电话.Text & "','" & txt单位邮编.Text & "','" & txt家庭地址.Text & "'," & _
        "'" & txt家庭电话.Text & "','" & txt家庭地址邮编.Text & "','" & txt户口地址.Text & "','" & txt户口地址邮编.Text & "'," & _
        "'" & txt联系人姓名.Text & "','" & NeedName(cbo联系人关系.Text) & "','" & txt联系人电话.Text & "','" & txt联系人地址.Text & "'," & _
        "'" & NeedName(cbo责任护士.Text) & "','" & NeedName(cbo门诊医师.Text) & "'," & _
        "'" & NeedName(cbo住院医师.Text) & "'," & _
        ZVal(lng西医疾病ID) & "," & ZVal(lng西医诊断ID) & ",'" & Replace(txt入院诊断.Text, "'", "''") & "'," & _
        ZVal(lng中医疾病ID) & "," & ZVal(lng中医诊断ID) & ",'" & Replace(txt中医诊断.Text, "'", "''") & "'," & _
        "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & NeedName(cbo主治医师.Text) & "','" & NeedName(cbo主任医师.Text) & "'," & _
        chk再入院.Value & ",'" & Trim(txt身份证号.Text) & "','" & Trim(txt出生地点.Text) & "','" & NeedName(txt籍贯.Text) & "','" & NeedName(txt区域.Text) & "','" & _
        NeedName(cbo入院属性.Text) & "','" & NeedName(cbo病人类型.Text) & "','" & NeedName(cbo入院方式.Text) & "'," & _
        IIf(Trim(txt备注.Text) = "", "Null", "'" & Trim(txt备注.Text) & "'") & "," & chk陪伴.Value & ",'" & Trim(txt联系人身份证号.Text) & "','" & NeedName(cbo国籍.Text) & "')"
        
    ReDim Preserve arrSQL(0)
    arrSQL(UBound(arrSQL)) = strSQL
    
    If Val(cbo费别.Tag) <> cbo费别.ListIndex And InStr(";" & mstrPrivs & ";", ";重算费用;") > 0 And Nvl(mrsPati!险类, 0) = 0 Then
        If MsgBox("病人费别被改变，要将该病人的未结费用按新的费别重算吗?" & vbCrLf & vbCrLf & _
            "本操作将按病人当前费别对应的优惠比率对未结费用重新进行打折计算!", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            
            strSQL_Recalc = "Zl_病人未结费用_Recalc(" & mlng病人ID & "," & mlng主页ID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL_Recalc
        End If
    End If
    '从表信息更新
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",联系人附加信息,身份证号状态,外籍身份证号,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & mlng病人ID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
            End If
            
            If InStr(",身份证号状态,外籍身份证号,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "')"
            End If
        Next
    End If
    
    If gbln启用结构化地址 Then
        Call CreateStructAddressSQL(mlng病人ID, mlng主页ID, arrSQL, PatiAddress, 1)
    End If
    
    '如果只改了住院医师，则只产生住院医师变动，如果改了医疗小组，则只产生医疗小组变动。
    
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    strBeginDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    For i = LBound(arrSQL) To UBound(arrSQL)
        zlDatabase.ExecuteProcedure arrSQL(i), Me.Caption
    Next
    strEndDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    '调用医保病人信息修改接口
    If Not IsNull(mrsPati!险类) Then
        If Not gclsInsure.ModiPatiSwap(mlng病人ID, mlng主页ID, mrsPati!险类, "1") Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    End If
    '更新EMPI平台病人信息
    strMsg = ""
    If Not EMPI_AddORUpdatePati(mlng病人ID, mlng主页ID, strMsg) Then
        gcnOracle.RollbackTrans
        MsgBox strMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    gcnOracle.CommitTrans: blnTrans = False
    '新网118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng病人ID, mlng主页ID) <> 1 Then
            MsgBox "当前启用了影像信息系统接口，但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "当前启用了影像信息系统接口，但于由RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
    End If
    gblnOK = True
    On Error Resume Next
    
    '病人基本信息调整是否成功都不影响病人信息保存
    If bln基本信息调整 And blnMod Then
        strErrInfo = ""
        Call gobjPublicPatient.SavePatiBaseInfo(mlng病人ID, mlng主页ID, Trim(txt姓名.Text), str性别, strAge, str出生日期, Me.Caption, IIf(mlng主页ID <> 0, 2, 1), strErrInfo, True, True)
        '提示
        If strErrInfo <> "" Then
            MsgBox strErrInfo, vbInformation + vbOKOnly, Me.Caption
        End If
    End If
    
    '病情变化后触发消息
    If NeedName(cbo病况.Text) <> Nvl(mrsPati!当前病况) Then
        Call PatiInfoChange(13, strBeginDate, strEndDate)
    End If
    '住院医师变动后触发消息
    If NeedName(cbo住院医师.Text) <> Nvl(mrsPati!住院医师) Then
        Call PatiInfoChange(7, strBeginDate, strEndDate)
    End If
    '责任护士变动后触发消息
    If NeedName(cbo责任护士.Text) <> Nvl(mrsPati!责任护士) Then
        Call PatiInfoChange(8, strBeginDate, strEndDate)
    End If
    '主治医师变动后触发消息
    If NeedName(cbo主治医师.Text) <> Nvl(mrsPati!主治医师) Then
        Call PatiInfoChange(11, strBeginDate, strEndDate)
    End If
    '主任医师变动后触发消息
    If NeedName(cbo主任医师.Text) <> Nvl(mrsPati!主任医师) Then
        Call PatiInfoChange(12, strBeginDate, strEndDate)
    End If
    
    If Err <> 0 Then Err.Clear
    
    Unload Me
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd出生地点_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt出生地点, True)
    If Not rsTmp Is Nothing Then
        txt出生地点.Text = rsTmp!名称
        txt出生地点.SelStart = Len(txt出生地点.Text)
        txt出生地点.SetFocus
    End If
End Sub

Private Sub cmd户口地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt户口地址, True)
    If Not rsTmp Is Nothing Then
        txt户口地址.Text = rsTmp!名称
        txt户口地址.SelStart = Len(txt户口地址.Text)
        txt户口地址.SetFocus
    End If
End Sub

Private Sub cmd籍贯_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt籍贯, True)
    If Not rsTmp Is Nothing Then
        txt籍贯.Text = rsTmp!名称
        txt籍贯.SelStart = Len(txt籍贯.Text)
        txt籍贯.SetFocus
    Else
        SelAll txt籍贯
        txt籍贯.SetFocus
    End If
End Sub

Private Sub cmd区域_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt区域, True)
    If Not rsTmp Is Nothing Then
        txt区域.Text = rsTmp!名称
        txt区域.SelStart = Len(txt区域.Text)
        txt区域.SetFocus
    Else
        SelAll txt区域
        txt区域.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
        Case vbKeyReturn
            If InStr(UCase(",txt单位地址,txt户口地址,txt出生地点,txt家庭地址,txt联系人地址,txt入院诊断,txt中医诊断,txt籍贯,txt区域,PatiAddress,"), UCase("," & ActiveControl.Name & ",")) = 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case vbKeyF3
            If Me.ActiveControl.Name = txt单位地址.Name Then
                cmd单位地址_Click
            ElseIf Me.ActiveControl.Name = txt家庭地址.Name Then
                cmd家庭地址_Click
            ElseIf Me.ActiveControl.Name = txt出生地点.Name Then
                cmd出生地点_Click
            ElseIf Me.ActiveControl.Name = txt户口地址.Name Then
                cmd户口地址_Click
            ElseIf Me.ActiveControl.Name = txt联系人地址.Name Then
                cmd联系人地址_Click
            ElseIf Me.ActiveControl.Name = txt籍贯.Name Then
                cmd籍贯_Click
            ElseIf Me.ActiveControl.Name = txt区域.Name Then
                cmd区域_Click
            End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") And Not (Me.ActiveControl Is txt入院诊断 Or Me.ActiveControl Is txt中医诊断) Then KeyAscii = 0      '诊断内容中可能有'号
End Sub

Private Sub Form_Load()
    Dim strSQL As String, strTmp As String
    Dim rsDiagnosis As ADODB.Recordset, rsBeds As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    
    mblnEMPI = False
    gblnOK = False
    Call InitStructAddress
    '问题27351 by lesfeng 2010-01-12 ,A.备注
    '问题24463 by lesfeng 2010-03-22 增加陪伴
    On Error GoTo errH
    strSQL = "Select NVL(A.姓名,D.姓名) 姓名,NVL(A.性别,D.性别) 性别, NVL(A.年龄,D.年龄) 年龄,To_Char(A.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间,E.名称 as 当前科室,A.出院科室id as 当前科室ID,H.名称 as 当前病区,A.当前病区Id, A.医疗小组id, g.名称 as 医疗小组, " & vbNewLine & _
            "A.住院号,A.责任护士, A.门诊医师, A.住院医师, B.信息值 主治医师, C.信息值 主任医师, A.费别, A.婚姻状况, A.学历," & vbNewLine & _
            "       A.职业, A.当前病况, A.单位地址, A.单位邮编, A.单位电话, A.家庭地址, A.家庭电话, A.家庭地址邮编, A.户口地址, A.户口地址邮编, A.联系人地址," & vbNewLine & _
            "       A.联系人电话, A.联系人姓名, A.联系人关系,A.联系人身份证号, A.再入院, A.病人性质, A.险类, D.身份证号,D.国籍, D.籍贯, D.区域, D.出生地点," & vbNewLine & _
            "       D.出生日期, A.入院属性,D.合同单位id, F.名称 As 护理等级,Nvl(A.病人类型,Decode(A.险类,Null,'普通病人','医保病人')) 病人类型,A.入院方式,A.备注,A.是否陪伴" & vbNewLine & _
            "From 病案主页 A, 病案主页从表 B, 病案主页从表 C, 病人信息 D,部门表 E,部门表 H,收费项目目录 F, 临床医疗小组 G " & vbNewLine & _
            "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id(+) And A.主页id = B.主页id(+) And A.病人id = C.病人id(+) And" & vbNewLine & _
            "      A.主页id = C.主页id(+) And A.医疗小组id = G.id(+) And B.信息名(+) = '主治医师' And C.信息名(+) = '主任医师' And A.病人id = D.病人id And A.出院科室id = E.id And A.当前病区Id=H.id(+)" & vbNewLine & _
            " And A.护理等级id = F.ID(+)"
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        
    With mrsPati
        txt姓名.Text = !姓名
        txt住院号.Text = "" & !住院号
        txt性别.Text = "" & !性别
        txt入院时间.Text = !入院时间
        txt科室.Text = "" & !当前科室
        txt科室.Tag = Val("" & !当前科室id)
        '是否是中医科
        txt中医诊断.Enabled = (InStr(1, "," & GetDepCharacter(Val("" & !当前科室id)) & ",", ",中医科,") > 0)
        txt中医诊断.ToolTipText = "只有当病人所在科室的性质为中医科时才允许输入中医诊断!"
        
        txt护理.Text = "" & !护理等级
        txt单位地址.Tag = Val("" & !合同单位ID)
        txt医疗小组.Text = "" & !医疗小组
        txt医疗小组.Tag = "" & !医疗小组id
        mstrBirthDay = "" & !出生日期
    End With
    
    Set rsBeds = GetPatiBeds(mlng病人ID)
    With rsBeds
        If .RecordCount = 0 Then
            txt床位.Text = "家庭病床"
            txt等级.Text = "无"
        Else
            strTmp = ""
            Do While Not .EOF
                txt床位.Text = txt床位.Text & "," & !床号
                If InStr("," & strTmp & ",", "," & !床位等级 & ",") = 0 Then strTmp = strTmp & "," & !床位等级
                .MoveNext
            Loop
            txt床位.Text = Mid(txt床位.Text, 2)
            txt等级.Text = Mid(strTmp, 2)
        End If
    End With
    
        
    Call InitDicts
    
    Call LoadOldData("" & mrsPati!年龄, txt年龄, cbo年龄单位)   '不重算年龄,保持原有值以便修改
    
    txt身份证号.Text = "" & mrsPati!身份证号
    txt身份证号.Tag = "" & mrsPati!身份证号
    cboIDNumber.Enabled = txt身份证号.Text = ""
    txt区域.Text = Nvl(mrsPati!区域)
    
    cbo国籍.ListIndex = GetCboIndex(cbo国籍, IIf(IsNull(mrsPati!国籍), "", mrsPati!国籍))
    If cbo国籍.ListIndex = -1 Then Call SetCboDefault(cbo国籍)
    
    If InStr(mstrPrivs, "调整门诊医师") = 0 Then cbo门诊医师.Enabled = False
    cbo住院医师.ListIndex = GetCboIndex(cbo住院医师, IIf(IsNull(mrsPati!住院医师), "", mrsPati!住院医师))
    cbo主治医师.ListIndex = GetCboIndex(cbo主治医师, IIf(IsNull(mrsPati!主治医师), "", mrsPati!主治医师))
    
    cbo责任护士.ListIndex = GetCboIndex(cbo责任护士, IIf(IsNull(mrsPati!责任护士), "", mrsPati!责任护士))
    cbo门诊医师.ListIndex = GetCboIndex(cbo门诊医师, IIf(IsNull(mrsPati!门诊医师), "", mrsPati!门诊医师))
    cbo主任医师.ListIndex = GetCboIndex(cbo主任医师, IIf(IsNull(mrsPati!主任医师), "", mrsPati!主任医师))
            
    cbo费别.ListIndex = GetCboIndex(cbo费别, IIf(IsNull(mrsPati!费别), "", mrsPati!费别))
    cbo费别.Tag = cbo费别.ListIndex '记录原始费别，用于保存时判断是否进行重算费用
    cbo费别.Enabled = InStr(mstrPrivs, "调整病人费别") > 0
    
    cbo婚姻状况.ListIndex = GetCboIndex(cbo婚姻状况, IIf(IsNull(mrsPati!婚姻状况), "", mrsPati!婚姻状况))
    cbo学历.ListIndex = GetCboIndex(cbo学历, IIf(IsNull(mrsPati!学历), "", mrsPati!学历))
    cbo职业.ListIndex = GetCboIndex(cbo职业, IIf(IsNull(mrsPati!职业), "", mrsPati!职业))
    cbo病况.ListIndex = GetCboIndex(cbo病况, IIf(IsNull(mrsPati!当前病况), "", mrsPati!当前病况))
    cbo入院属性.ListIndex = GetCboIndex(cbo入院属性, IIf(IsNull(mrsPati!入院属性), "", mrsPati!入院属性))
    cbo入院方式.ListIndex = GetCboIndex(cbo入院方式, IIf(IsNull(mrsPati!入院方式), "", mrsPati!入院方式))
    
    txt单位地址.Text = IIf(IsNull(mrsPati!单位地址), "", mrsPati!单位地址)
    txt单位邮编.Text = IIf(IsNull(mrsPati!单位邮编), "", mrsPati!单位邮编)
    txt单位电话.Text = IIf(IsNull(mrsPati!单位电话), "", mrsPati!单位电话)
    cbo病人类型.ListIndex = GetCboIndex(cbo病人类型, mrsPati!病人类型)
    If InStr(mstrPrivs, "调整病人类型") = 0 Then cbo病人类型.Enabled = False
        
    txt家庭电话.Text = IIf(IsNull(mrsPati!家庭电话), "", mrsPati!家庭电话)
    txt家庭地址邮编.Text = IIf(IsNull(mrsPati!家庭地址邮编), "", mrsPati!家庭地址邮编)
    txt户口地址邮编.Text = IIf(IsNull(mrsPati!户口地址邮编), "", mrsPati!户口地址邮编)
    
    txt联系人电话.Text = IIf(IsNull(mrsPati!联系人电话), "", mrsPati!联系人电话)
    txt联系人姓名.Text = IIf(IsNull(mrsPati!联系人姓名), "", mrsPati!联系人姓名)
    txt联系人身份证号.Text = IIf(IsNull(mrsPati!联系人身份证号), "", mrsPati!联系人身份证号)
    cbo联系人关系.ListIndex = GetCboIndex(cbo联系人关系, IIf(IsNull(mrsPati!联系人关系), "", mrsPati!联系人关系))
    If gbln启用结构化地址 Then
        Call ReadStructAddress(mlng病人ID, mlng主页ID, PatiAddress)
        txt出生地点.Text = PatiAddress(E_IX_出生地点).Value
        txt籍贯.Text = PatiAddress(E_IX_籍贯).Value
        txt家庭地址.Text = PatiAddress(E_IX_现住址).Value
        txt户口地址.Text = PatiAddress(E_IX_户口地址).Value
        txt联系人地址.Text = PatiAddress(E_IX_联系人地址).Value
    Else
        txt出生地点.Text = "" & mrsPati!出生地点
        txt籍贯.Text = Nvl(mrsPati!籍贯)
        txt家庭地址.Text = IIf(IsNull(mrsPati!家庭地址), "", mrsPati!家庭地址)
        txt户口地址.Text = IIf(IsNull(mrsPati!户口地址), "", mrsPati!户口地址)
        txt联系人地址.Text = IIf(IsNull(mrsPati!联系人地址), "", mrsPati!联系人地址)
    End If
    '问题27351 by lesfeng 2010-01-12
    txt备注.Text = IIf(IsNull(mrsPati!备注), "", mrsPati!备注)
    '问题24463 by lesfeng 2010-03-22 增加陪伴
    chk陪伴.Value = IIf(IsNull(mrsPati!是否陪伴), 0, mrsPati!是否陪伴)
    
     '显示病人诊断记录
    Set rsDiagnosis = GetDiagnosticInfo(mlng病人ID, mlng主页ID, "1,11,2,12", "2")
    If Not rsDiagnosis Is Nothing Then
        'a.西医诊断
         rsDiagnosis.Filter = "诊断类型=2"        '先取以前保存的入院诊断
         If Not rsDiagnosis.EOF Then
             txt入院诊断.Text = Nvl(rsDiagnosis!诊断描述): txt入院诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl入院诊断.Tag = txt入院诊断.Text
         Else
             rsDiagnosis.Filter = "诊断类型=1"    '再取入院登记的门诊诊断
             If Not rsDiagnosis.EOF Then
                 txt入院诊断.Text = Nvl(rsDiagnosis!诊断描述): txt入院诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl入院诊断.Tag = txt入院诊断.Text
             End If
         End If
     
        'b.中医诊断
        If txt中医诊断.Enabled Then
            rsDiagnosis.Filter = "诊断类型=12"        '先取以前保存的入院诊断
            If Not rsDiagnosis.EOF Then
                txt中医诊断.Text = Nvl(rsDiagnosis!诊断描述): txt中医诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
            Else
                rsDiagnosis.Filter = "诊断类型=11"    '再取入院登记的门诊诊断
                If Not rsDiagnosis.EOF Then
                    txt中医诊断.Text = Nvl(rsDiagnosis!诊断描述): txt中医诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                End If
            End If
        End If
    End If
    chk再入院.Value = Val("" & mrsPati!再入院)
    
    
    '54045:刘鹏飞,2012-09-27,如果首页中对应的医师签名后，则不能修改
    strSQL = "Select 信息名,信息值 From 病案主页从表 Where 病人ID=[1] And 主页ID=[2] And 信息值 is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    rsTmp.Filter = "信息名='住院医师签名'"
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!信息值) Then
            cbo住院医师.Enabled = False
            cbo住院医师.BackColor = &HE0E0E0
        End If
    End If
    rsTmp.Filter = "信息名='主治医师签名'"
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!信息值) Then
            cbo主治医师.Enabled = False
            cbo主治医师.BackColor = &HE0E0E0
        End If
    End If
    rsTmp.Filter = "信息名='主任医师签名'"
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!信息值) Then
            cbo主任医师.Enabled = False
            cbo主任医师.BackColor = &HE0E0E0
        End If
    End If
    '病人从表
    Set rsTmp = Get病人信息从表(mlng病人ID, "身份证号状态")
    rsTmp.Filter = "信息名='身份证号状态'"
    If Not rsTmp.EOF Then
        Call cbo.Locate(cboIDNumber, zlCommFun.GetNeedName(rsTmp!信息值 & ""))
    End If
    If Trim(zlCommFun.GetNeedName(cbo国籍.Text)) <> "中国" And Trim(txt身份证号.Text) = "" Then
        If Trim(zlCommFun.GetNeedName(cboIDNumber.Text)) = "" Then
            Set rsTmp = Get病人信息从表(mlng病人ID, "外籍身份证号")
            rsTmp.Filter = "信息名='外籍身份证号'"
            If Not rsTmp.EOF Then
                txt身份证号.Text = "" & rsTmp!信息值
            End If
        End If
    End If
    '加载EMPI平台病人信息
    Call EMPI_LoadPati
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, P病人入出管理, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '卸载消息对象
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub

Private Sub PatiAddress_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True) '打开中文输入法
End Sub

Private Sub PatiAddress_LostFocus(Index As Integer)
'功能:
    Select Case Index
    
    Case E_IX_现住址
        txt家庭地址.Text = PatiAddress(Index).Value
    Case E_IX_出生地点
        txt出生地点.Text = PatiAddress(Index).Value
    Case E_IX_户口地址
        txt户口地址.Text = PatiAddress(Index).Value
    Case E_IX_籍贯
        txt籍贯.Text = PatiAddress(Index).Value
    Case E_IX_联系人地址
        txt联系人地址.Text = PatiAddress(Index).Value
    End Select
    Call zlCommFun.OpenIme '关闭中文输入法
End Sub

Private Sub PatiAddress_Validate(Index As Integer, Cancel As Boolean)
    Dim lngLen As Long
    
    lngLen = PatiAddress(Index).MaxLength
    If LenB(StrConv(PatiAddress(Index).Value, vbFromUnicode)) > lngLen Then
        MsgBox PatiAddress(Index).Tag & "只允许输入 " & lngLen & " 个字符或 " & lngLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

'问题27351 by lesfeng 2010-01-12  b
Private Sub txt备注_GotFocus()
    Call zlControl.TxtSelAll(txt备注)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt备注_KeyPress(KeyAscii As Integer)
    CheckInputLen txt备注, KeyAscii
End Sub

Private Sub txt备注_LostFocus()
    Call zlCommFun.OpenIme
End Sub
'问题27351 by lesfeng 2010-01-12 e
Private Sub txt出生地点_GotFocus()
    zlControl.TxtSelAll txt出生地点
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt出生地点_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt出生地点.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt出生地点)
            If Not rsTmp Is Nothing Then
                txt出生地点.Text = rsTmp!名称
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt出生地点, KeyAscii
    End If
End Sub

Private Sub txt出生地点_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt单位电话_GotFocus()
    SelAll txt单位电话
End Sub

Private Sub txt单位电话_KeyPress(KeyAscii As Integer)
    If InStr("01234567890()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt单位邮编_GotFocus()
    SelAll txt单位邮编
End Sub

Private Sub txt单位邮编_KeyPress(KeyAscii As Integer)
    If InStr("01234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt户口地址_GotFocus()
    SelAll txt户口地址
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt户口地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt户口地址.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt户口地址)
            If Not rsTmp Is Nothing Then
                txt户口地址.Text = rsTmp!名称
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt户口地址, KeyAscii
    End If
End Sub

Private Sub txt户口地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt户口地址邮编_GotFocus()
    SelAll txt户口地址邮编
End Sub

Private Sub txt户口地址邮编_KeyPress(KeyAscii As Integer)
    If InStr("01234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt籍贯_GotFocus()
    zlControl.TxtSelAll txt籍贯
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt籍贯_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt籍贯.Text <> "" Then
            Set rsTmp = GetArea(Me, txt籍贯)
            If Not rsTmp Is Nothing Then
                txt籍贯.Text = rsTmp!名称
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                SelAll txt籍贯
                txt籍贯.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt籍贯, KeyAscii
    End If
End Sub

Private Sub txt籍贯_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt家庭地址邮编_GotFocus()
    SelAll txt家庭地址邮编
End Sub

Private Sub txt家庭地址邮编_KeyPress(KeyAscii As Integer)
    If InStr("01234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt家庭地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt家庭电话_GotFocus()
    SelAll txt家庭电话
End Sub

Private Sub txt家庭电话_KeyPress(KeyAscii As Integer)
    If InStr("01234567890()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt联系人电话_GotFocus()
    SelAll txt联系人电话
End Sub

Private Sub txt联系人电话_KeyPress(KeyAscii As Integer)
    If InStr("01234567890()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


Private Sub txt联系人姓名_GotFocus()
    SelAll txt联系人姓名
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt联系人姓名_KeyPress(KeyAscii As Integer)
    CheckInputLen txt联系人姓名, KeyAscii
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt联系人姓名_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt年龄_GotFocus()
    Call zlCommFun.OpenIme
    SelAll txt年龄
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo年龄单位.Visible = False And IsNumeric(txt年龄.Text) Then
            Call txt年龄_Validate(False)
            Call cbo年龄单位.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt年龄.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt单位地址_Change()
    If txt单位地址.Text = "" Then txt单位地址.Tag = ""
End Sub

Private Sub txt单位地址_GotFocus()
    SelAll txt单位地址
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt单位地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt单位地址.Text <> "" Then
            Set rsTmp = GetOrgAddress(Me, txt单位地址)
            If Not rsTmp Is Nothing Then
                txt单位地址.Text = rsTmp!名称
                txt单位地址.Tag = rsTmp!ID
                txt单位电话.Text = Trim(rsTmp!电话 & "")
            Else
                txt单位地址.Tag = ""
            End If
        Else
            txt单位地址.Tag = ""
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt单位地址, KeyAscii
    End If
End Sub

Private Sub txt单位地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt家庭地址_GotFocus()
    SelAll txt家庭地址
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt家庭地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt家庭地址.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt家庭地址)
            If Not rsTmp Is Nothing Then
                txt家庭地址.Text = rsTmp!名称
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt家庭地址, KeyAscii
    End If
End Sub

Private Sub txt联系人地址_GotFocus()
    SelAll txt联系人地址
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt联系人地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt联系人地址.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt联系人地址)
            If Not rsTmp Is Nothing Then
                txt联系人地址.Text = rsTmp!名称
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt联系人地址, KeyAscii
    End If
End Sub

Private Sub txt联系人地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub cmd单位地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetOrgAddress(Me, txt单位地址, True)
    If Not rsTmp Is Nothing Then
        txt单位地址.Tag = rsTmp!ID
        txt单位地址.Text = rsTmp!名称
        txt单位地址.SelStart = Len(txt单位地址.Text)
        txt单位电话.Text = Trim(rsTmp!电话 & "")
        txt单位地址.SetFocus
    End If
End Sub

Private Sub cmd家庭地址_Click()
    Dim rsTmp As ADODB.Recordset
   Set rsTmp = GetAddress(Me, txt家庭地址, True)
    If Not rsTmp Is Nothing Then
        txt家庭地址.Text = rsTmp!名称
        txt家庭地址.SelStart = Len(txt家庭地址.Text)
        txt家庭地址.SetFocus
    End If
End Sub

Private Sub cmd联系人地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt联系人地址, True)
    If Not rsTmp Is Nothing Then
        txt联系人地址.Text = rsTmp!名称
        txt联系人地址.SelStart = Len(txt联系人地址.Text)
        txt联系人地址.SetFocus
    End If
End Sub

Private Sub InitDicts()
    Dim strSQL As String, i As Integer
    Dim strSQL医疗小组 As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    cbo年龄单位.Clear
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    cbo年龄单位.ListIndex = 0
    txt年龄.MaxLength = GetColumnLength("病人信息", "年龄")
    
    Call ReadDict("费别", cbo费别)
    Call ReadDict("病情", cbo病况)
    Call ReadDict("学历", cbo学历)
    Call ReadDict("婚姻状况", cbo婚姻状况)
    Call ReadDict("职业", cbo职业)
    Call ReadDict("社会关系", cbo联系人关系)
    Call ReadDict("入院属性", cbo入院属性)
    Call ReadDict("入院方式", cbo入院方式)
    Call ReadDict("病人类型", cbo病人类型, "病人类型")
   
    Call ReadDict("身份证未录原因", cboIDNumber)
    Call ReadDict("国籍", cbo国籍)
    
    mstrUnit = Get科室IDs(mlngUnit) & "," & mlngUnit

    '医疗小组
    strSQL = "Select Distinct A.ID, A.编号, A.简码, A.姓名" & vbNewLine & _
                        " From 人员表 A, 人员性质说明 B, 部门人员 C" & vbNewLine & _
                        " Where A.ID = B.人员id And A.ID = C.人员id And B.人员性质 = '医生' And" & vbNewLine & _
                        "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
                        "      (Instr(',' || [1] || ',', ',' || C.部门id || ',') > 0 Or A.姓名=[2]) And Instr(',' || [3] || ',', ',' || A.专业技术职务 || ',') > 0" & vbNewLine & _
                        "      And (A.站点=[4] Or A.站点 is Null)" & _
                        " Order By A.简码"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPati!住院医师), "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
    cbo住院医师.Clear
    Do Until rsTmp.EOF
        cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPati!主治医师), "主任医师,副主任医师,主治医师", gstrNodeNo)
    cbo主治医师.Clear
    Do Until rsTmp.EOF
        cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPati!主任医师), "主任医师,副主任医师", gstrNodeNo)
    Do Until rsTmp.EOF
        cbo主任医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo主任医师.ItemData(cbo主任医师.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop

    '门诊医师
    Set rsTmp = GetDoctorOrNurse(0)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo门诊医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            cbo门诊医师.ItemData(cbo门诊医师.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    
    'by lesfeng 2010-01-12 性能优化
    '住院护士
    strSQL = _
        "Select Distinct A.ID,A.编号,A.简码,A.姓名" & _
        " From 人员表 A,人员性质说明 B,部门人员 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And B.人员性质=[1] And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (Instr(','||[2]||',',','||C.部门ID||',')>0  Or A.姓名=[3])" & _
        " And (A.站点=[4] Or A.站点 is Null)" & _
        " Order by A.简码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "护士", mstrUnit, CStr("" & mrsPati!责任护士), gstrNodeNo)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo责任护士.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            cbo责任护士.ItemData(cbo责任护士.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    
    cbo住院医师.AddItem "其它..."
    cbo主治医师.AddItem "其它..."
    cbo主任医师.AddItem "其它..."
    cbo责任护士.AddItem "其它..."
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadDict(strDict As String, cbo As ComboBox, Optional strClass As String) As Boolean
'功能：初始化指定词典
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngMaxW As Long
    Dim strTemp As String

    On Error GoTo errH
    'by lesfeng 2010-01-12 性能优化
    If strDict = "费别" Then
        If Nvl(mrsPati!病人性质, 0) = 1 Then
            strTemp = "1,3" '门诊留观病人
        Else
            strTemp = "2,3"
        End If
'        strSql = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别 Where Nvl(服务对象,3) IN(" & strTemp & ") And  Sysdate Between NVL(有效开始,Sysdate-1) and NVL(有效结束,Sysdate+1) Order by 编码"
        strSQL = "Select A.编码,A.名称,A.简码,Nvl(A.缺省标志,0) as 缺省 From 费别 A,Table(Cast(f_Num2List([1]) As zlTools.t_Numlist)) B " & _
                 " Where (A.服务对象 = B.Column_Value or A.服务对象 is null) And  Sysdate Between NVL(A.有效开始,Sysdate-1) and NVL(A.有效结束,Sysdate+1) Order by A.编码"
    ElseIf strDict = "区域" Then
        strSQL = "Select 编码,名称,简码,0 as 缺省 From  区域 Order by 编码"
    ElseIf strDict = "病人类型" Then
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省,颜色 From 病人类型 Order by 编码"
    Else
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    End If
'    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption, strTemp)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTemp)
    
    cbo.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If strDict = "职业" Then
                cbo.AddItem rsTmp!编码 & "-" & Chr(&HA) & rsTmp!名称
            Else
                cbo.AddItem rsTmp!编码 & "-" & rsTmp!名称
            End If
            If rsTmp!缺省 = 1 Then
                cbo.ListIndex = cbo.NewIndex
                cbo.ItemData(cbo.NewIndex) = 1
            End If
            If TextWidth(cbo.List(cbo.NewIndex) & "字") > lngMaxW Then lngMaxW = TextWidth(cbo.List(cbo.NewIndex) & "字")
            rsTmp.MoveNext
        Next
    End If
    ReadDict = True
    If GetWidth(cbo.hWnd) * Screen.TwipsPerPixelX < lngMaxW Then SetWidth cbo.hWnd, lngMaxW / Screen.TwipsPerPixelX
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt年龄_Validate(Cancel As Boolean)
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False
    ElseIf cbo年龄单位.Visible = False Then
        cbo年龄单位.ListIndex = 0: cbo年龄单位.Visible = True
    End If
End Sub

Private Sub txt区域_GotFocus()
    zlControl.TxtSelAll txt区域
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt区域_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt区域.Text <> "" Then
            Set rsTmp = GetArea(Me, txt区域)
            If Not rsTmp Is Nothing Then
                txt区域.Text = rsTmp!名称
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                SelAll txt区域
                txt区域.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt区域, KeyAscii
    End If
End Sub

Private Sub txt区域_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt入院诊断_GotFocus()
    SelAll txt入院诊断
End Sub

Private Sub txt身份证号_GotFocus()
    zlControl.TxtSelAll txt身份证号
End Sub

Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt联系人身份证号_GotFocus()
    zlControl.TxtSelAll txt身份证号
End Sub

Private Sub txt联系人身份证号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt身份证号_LostFocus()
    If Trim(txt身份证号.Text) = "" Then
        cboIDNumber.Enabled = True
        cboIDNumber.SetFocus
    Else
        cboIDNumber.Enabled = False
        cboIDNumber.ListIndex = -1
    End If
End Sub

Private Sub txt中医诊断_GotFocus()
    SelAll txt中医诊断
End Sub

Private Sub txt入院诊断_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '问题25785 by lesfeng 2009-10-20 处理允许自由录入规则
            '************************************************
            If gint住院诊断输入 = 1 Then
                strInput = UCase(txt入院诊断.Text)
                strSex = NeedName(txt性别.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "名称 Like [2] or '('||编码||')'||名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gbytCode = 0, "简码", "五笔码") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gbytCode = 0, "简码", "五笔码 as 简码") & ",说明" & _
                        " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                        " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by 编码"
                '问题27613 by lesfeng 2010-01-21
                '自由录入时有多个匹配(汉字)不进行选择,数字及字母则进行选择
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "D", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                Else
                    vPoint = GetCoordPos(fraInfo.hWnd, txt入院诊断.Left, txt入院诊断.Top)
                    strInput = UCase(txt入院诊断.Text)
                    strSex = NeedName(txt性别.Text)
                    lngTxtHeight = txt入院诊断.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '数据库中只有一个匹配项目，则以该匹配的项目为准
                    txt入院诊断.Tag = rsTmp!ID
                    txt入院诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称 '
                    lbl入院诊断.Tag = txt入院诊断.Text '用于恢复显示
                Else
                    '多项或者无匹配项目时才以输入的为准
                    txt入院诊断.Tag = ""
                    lbl入院诊断.Tag = txt入院诊断.Text '用于恢复显示
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt入院诊断.Text = lbl入院诊断.Tag And txt入院诊断.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt入院诊断.Text = "" Then
            txt入院诊断.Tag = "": lbl入院诊断.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            vPoint = GetCoordPos(fraInfo.hWnd, txt入院诊断.Left, txt入院诊断.Top)
            strInput = UCase(txt入院诊断.Text)
            strSex = NeedName(txt性别.Text)
            lngTxtHeight = txt入院诊断.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
            If Not rsTmp Is Nothing Then
                txt入院诊断.Tag = rsTmp!ID
                txt入院诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                lbl入院诊断.Tag = txt入院诊断.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If lbl入院诊断.Tag <> "" Then txt入院诊断.Text = lbl入院诊断.Tag
                Call txt入院诊断_GotFocus
                txt入院诊断.SetFocus
            End If
        End If
    Else
        CheckInputLen txt入院诊断, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt中医诊断_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '问题25785 by lesfeng 2009-10-20 处理允许自由录入规则
            '************************************************
            If gint住院诊断输入 = 1 Then
                strInput = UCase(txt中医诊断.Text)
                strSex = NeedName(txt性别.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "名称 Like [2] or '('||编码||')'||名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gbytCode = 0, "简码", "五笔码") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gbytCode = 0, "简码", "五笔码 as 简码") & ",说明" & _
                        " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                        " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by 编码"
                '问题27613 by lesfeng 2010-01-21
                '自由录入时有多个匹配(汉字)不进行选择,数字及字母则进行选择
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "B", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                Else
                    vPoint = GetCoordPos(fraInfo.hWnd, txt中医诊断.Left, txt中医诊断.Top)
                    strInput = UCase(txt中医诊断.Text)
                    strSex = NeedName(txt性别.Text)
                    lngTxtHeight = txt中医诊断.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '数据库中只有一个匹配项目，则以该匹配的项目为准
                    txt中医诊断.Tag = rsTmp!ID
                    txt中医诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称 '
                    lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                Else
                    '多项或者无匹配项目时才以输入的为准
                    txt中医诊断.Tag = ""
                    lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt中医诊断.Text = lbl中医诊断.Tag And txt中医诊断.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt中医诊断.Text = "" Then
            txt中医诊断.Tag = "": lbl中医诊断.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            vPoint = GetCoordPos(fraInfo.hWnd, txt中医诊断.Left, txt中医诊断.Top)
            strInput = UCase(txt中医诊断.Text)
            strSex = NeedName(txt性别.Text)
            lngTxtHeight = txt中医诊断.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
            If Not rsTmp Is Nothing Then
                txt中医诊断.Tag = rsTmp!ID
                txt中医诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If lbl中医诊断.Tag <> "" Then txt中医诊断.Text = lbl中医诊断.Tag
                Call txt中医诊断_GotFocus
                txt中医诊断.SetFocus
            End If
        End If
    Else
        CheckInputLen txt中医诊断, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt入院诊断_Validate(Cancel As Boolean)
    If Val(txt入院诊断.Tag) > 0 And txt入院诊断.Text <> lbl入院诊断.Tag Then
        txt入院诊断.Text = lbl入院诊断.Tag
    ElseIf Val(txt入院诊断.Tag) = 0 And RequestCode Then
        txt入院诊断.Text = ""
    End If
End Sub

Private Sub txt中医诊断_Validate(Cancel As Boolean)
    If Val(txt中医诊断.Tag) > 0 And txt中医诊断.Text <> lbl中医诊断.Tag Then
        txt中医诊断.Text = lbl中医诊断.Tag
    ElseIf Val(txt中医诊断.Tag) = 0 And RequestCode Then
        txt中医诊断.Text = ""
    End If
End Sub

Private Function RequestCode() As Boolean
    RequestCode = gint住院诊断输入 = 2 Or (gint住院诊断输入 = 3 And Not IsNull(mrsPati!险类))
End Function

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    
    ShowMe = gblnOK
End Function

Private Function PatiInfoChange(ByVal intType As Integer, ByVal strBeginDate As String, ByVal strEndDate As String) As Boolean
'功能:病情、责任护士、住院医师、主任医师、主治医师变动后触发消息
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    Select Case intType
    Case 13 '病情变动
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '清除缓存中的XML
            '--进行消息组装
            '病人信息
            mclsXML.AppendNode "in_patient"
            'patient_id      病人id  1   N
            mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
            'page_id     主页id  1   N
            mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
            'patient_name        姓名    1   S
            mclsXML.appendData "patient_name", txt姓名.Text, xsString '姓名
            'patient_sex     性别    0..1    S
            mclsXML.appendData "patient_sex", txt性别.Text, xsString '性别
            'in_number       住院号  1   S
            mclsXML.appendData "in_number", txt住院号.Text, xsString  '住院号
            mclsXML.AppendNode "in_patient", True
            
            '当前情况
            'current_state       当前情况    1
            mclsXML.AppendNode "current_state"
            'current_area_id     当前病区id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!当前病区ID)), xsNumber
            'current_area_title      当前病区    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!当前病区), xsString
            'current_dept_id     当前科室id  1   N
            mclsXML.appendData "current_dept_id", Val(txt科室.Tag), xsNumber
            'current_dept_title      当前科室    1   S
            mclsXML.appendData "current_dept_title", txt科室.Text, xsString
            'current_situation       当前病况    1    S
            mclsXML.appendData "current_situation", Nvl(mrsPati!当前病况), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID 变动ID,开始时间 变动时间 From 病人变动记录 where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] And NVL(附加床位,0)=0 And 开始时间+0 between [4] And　[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID, intType, CDate(strBeginDate), CDate(strEndDate))
            '变更信息
            'change_state        变更信息    1
            mclsXML.AppendNode "change_state"
            'change_id       变更id  1   N
            mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
            'change_date     变更时间    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_situation        变更病况    0..1    S
            mclsXML.appendData "change_situation", NeedName(cbo病况.Text), xsString
            'change_operator         操作员      1   S
            mclsXML.appendData "change_operator", UserInfo.姓名, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_005", mclsXML.XmlText)
        End If
    
    Case 7 '住院医师
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '清除缓存中的XML
            '--进行消息组装
            '病人信息
            mclsXML.AppendNode "in_patient"
            'patient_id      病人id  1   N
            mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
            'page_id     主页id  1   N
            mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
            'patient_name        姓名    1   S
            mclsXML.appendData "patient_name", txt姓名.Text, xsString '姓名
            'patient_sex     性别    0..1    S
            mclsXML.appendData "patient_sex", txt性别.Text, xsString '性别
            'in_number       住院号  1   S
            mclsXML.appendData "in_number", txt住院号.Text, xsString  '住院号
            mclsXML.AppendNode "in_patient", True
            
            '当前情况
            'current_state       当前情况    1
            mclsXML.AppendNode "current_state"
            'current_area_id     当前病区id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!当前病区ID)), xsNumber
            'current_area_title      当前病区    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!当前病区), xsString
            'current_dept_id     当前科室id  1   N
            mclsXML.appendData "current_dept_id", Val(txt科室.Tag), xsNumber
            'current_dept_title      当前科室    1   S
            mclsXML.appendData "current_dept_title", txt科室.Text, xsString
            'curren_in_doctor        住院医师    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPati!住院医师), xsString
            'curren_director_doctor      主任医师    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPati!主任医师), xsString
            'curren_treat_doctor     主治医师    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPati!主治医师), xsString
            'curren_duty_nurse       责任护士    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPati!责任护士), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID 变动ID,开始时间 变动时间 From 病人变动记录 where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] And NVL(附加床位,0)=0 And 开始时间+0 between [4] And　[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID, intType, CDate(strBeginDate), CDate(strEndDate))
            '变更信息
            'change_state        变更信息    1
            mclsXML.AppendNode "change_state"
            'change_id       变更id  1   N
            mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
            'change_date     变更时间    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        住院医师    1   S
            mclsXML.appendData "change_in_doctor", NeedName(cbo住院医师.Text), xsString
            'change_director_doctor      主任医师    1   S
            mclsXML.appendData "change_director_doctor", Nvl(mrsPati!主任医师), xsString
            'change_treat_doctor     主治医师    1   S
            mclsXML.appendData "change_treat_doctor", Nvl(mrsPati!主治医师), xsString
            'change_duty_nurse       责任护士    1   S
            mclsXML.appendData "change_duty_nurse", Nvl(mrsPati!责任护士), xsString
            'change_operator         操作员      1   S
            mclsXML.appendData "change_operator", UserInfo.姓名, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_007", mclsXML.XmlText)
        End If
    Case 8 '责任护士
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '清除缓存中的XML
            '--进行消息组装
            '病人信息
            mclsXML.AppendNode "in_patient"
            'patient_id      病人id  1   N
            mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
            'page_id     主页id  1   N
            mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
            'patient_name        姓名    1   S
            mclsXML.appendData "patient_name", txt姓名.Text, xsString '姓名
            'patient_sex     性别    0..1    S
            mclsXML.appendData "patient_sex", txt性别.Text, xsString '性别
            'in_number       住院号  1   S
            mclsXML.appendData "in_number", txt住院号.Text, xsString  '住院号
            mclsXML.AppendNode "in_patient", True
            
            '当前情况
            'current_state       当前情况    1
            mclsXML.AppendNode "current_state"
            'current_area_id     当前病区id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!当前病区ID)), xsNumber
            'current_area_title      当前病区    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!当前病区), xsString
            'current_dept_id     当前科室id  1   N
            mclsXML.appendData "current_dept_id", Val(txt科室.Tag), xsNumber
            'current_dept_title      当前科室    1   S
            mclsXML.appendData "current_dept_title", txt科室.Text, xsString
            'curren_in_doctor        住院医师    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPati!住院医师), xsString
            'curren_director_doctor      主任医师    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPati!主任医师), xsString
            'curren_treat_doctor     主治医师    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPati!主治医师), xsString
            'curren_duty_nurse       责任护士    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPati!责任护士), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID 变动ID,开始时间 变动时间 From 病人变动记录 where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] And NVL(附加床位,0)=0 And 开始时间+0 between [4] And　[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID, intType, CDate(strBeginDate), CDate(strEndDate))
            '变更信息
            'change_state        变更信息    1
            mclsXML.AppendNode "change_state"
            'change_id       变更id  1   N
            mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
            'change_date     变更时间    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        住院医师    1   S
            mclsXML.appendData "change_in_doctor", NeedName(cbo住院医师.Text), xsString
            'change_director_doctor      主任医师    1   S
            mclsXML.appendData "change_director_doctor", Nvl(mrsPati!主任医师), xsString
            'change_treat_doctor     主治医师    1   S
            mclsXML.appendData "change_treat_doctor", Nvl(mrsPati!主治医师), xsString
            'change_duty_nurse       责任护士    1   S
            mclsXML.appendData "change_duty_nurse", NeedName(cbo责任护士.Text), xsString
            'change_operator         操作员      1   S
            mclsXML.appendData "change_operator", UserInfo.姓名, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_007", mclsXML.XmlText)
        End If
    Case 11 '主治医师
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '清除缓存中的XML
            '--进行消息组装
            '病人信息
            mclsXML.AppendNode "in_patient"
            'patient_id      病人id  1   N
            mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
            'page_id     主页id  1   N
            mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
            'patient_name        姓名    1   S
            mclsXML.appendData "patient_name", txt姓名.Text, xsString '姓名
            'patient_sex     性别    0..1    S
            mclsXML.appendData "patient_sex", txt性别.Text, xsString '性别
            'in_number       住院号  1   S
            mclsXML.appendData "in_number", txt住院号.Text, xsString  '住院号
            mclsXML.AppendNode "in_patient", True
            
            '当前情况
            'current_state       当前情况    1
            mclsXML.AppendNode "current_state"
            'current_area_id     当前病区id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!当前病区ID)), xsNumber
            'current_area_title      当前病区    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!当前病区), xsString
            'current_dept_id     当前科室id  1   N
            mclsXML.appendData "current_dept_id", Val(txt科室.Tag), xsNumber
            'current_dept_title      当前科室    1   S
            mclsXML.appendData "current_dept_title", txt科室.Text, xsString
            'curren_in_doctor        住院医师    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPati!住院医师), xsString
            'curren_director_doctor      主任医师    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPati!主任医师), xsString
            'curren_treat_doctor     主治医师    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPati!主治医师), xsString
            'curren_duty_nurse       责任护士    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPati!责任护士), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID 变动ID,开始时间 变动时间 From 病人变动记录 where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] And NVL(附加床位,0)=0 And 开始时间+0 between [4] And　[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID, intType, CDate(strBeginDate), CDate(strEndDate))
            '变更信息
            'change_state        变更信息    1
            mclsXML.AppendNode "change_state"
            'change_id       变更id  1   N
            mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
            'change_date     变更时间    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        住院医师    1   S
            mclsXML.appendData "change_in_doctor", NeedName(cbo住院医师.Text), xsString
            'change_director_doctor      主任医师    1   S
            mclsXML.appendData "change_director_doctor", Nvl(mrsPati!主任医师), xsString
            'change_treat_doctor     主治医师    1   S
            mclsXML.appendData "change_treat_doctor", NeedName(cbo主治医师.Text), xsString
            'change_duty_nurse       责任护士    1   S
            mclsXML.appendData "change_duty_nurse", NeedName(cbo责任护士.Text), xsString
            'change_operator         操作员      1   S
            mclsXML.appendData "change_operator", UserInfo.姓名, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_007", mclsXML.XmlText)
        End If
    Case 12 '主任医师
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '清除缓存中的XML
            '--进行消息组装
            '病人信息
            mclsXML.AppendNode "in_patient"
            'patient_id      病人id  1   N
            mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
            'page_id     主页id  1   N
            mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
            'patient_name        姓名    1   S
            mclsXML.appendData "patient_name", txt姓名.Text, xsString '姓名
            'patient_sex     性别    0..1    S
            mclsXML.appendData "patient_sex", txt性别.Text, xsString '性别
            'in_number       住院号  1   S
            mclsXML.appendData "in_number", txt住院号.Text, xsString  '住院号
            mclsXML.AppendNode "in_patient", True
            
            '当前情况
            'current_state       当前情况    1
            mclsXML.AppendNode "current_state"
            'current_area_id     当前病区id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!当前病区ID)), xsNumber
            'current_area_title      当前病区    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!当前病区), xsString
            'current_dept_id     当前科室id  1   N
            mclsXML.appendData "current_dept_id", Val(txt科室.Tag), xsNumber
            'current_dept_title      当前科室    1   S
            mclsXML.appendData "current_dept_title", txt科室.Text, xsString
            'curren_in_doctor        住院医师    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPati!住院医师), xsString
            'curren_director_doctor      主任医师    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPati!主任医师), xsString
            'curren_treat_doctor     主治医师    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPati!主治医师), xsString
            'curren_duty_nurse       责任护士    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPati!责任护士), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID 变动ID,开始时间 变动时间 From 病人变动记录 where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] And NVL(附加床位,0)=0 And 开始时间+0 between [4] And　[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID, intType, CDate(strBeginDate), CDate(strEndDate))
            '变更信息
            'change_state        变更信息    1
            mclsXML.AppendNode "change_state"
            'change_id       变更id  1   N
            mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
            'change_date     变更时间    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        住院医师    1   S
            mclsXML.appendData "change_in_doctor", NeedName(cbo住院医师.Text), xsString
            'change_director_doctor      主任医师    1   S
            mclsXML.appendData "change_director_doctor", NeedName(cbo主任医师.Text), xsString
            'change_treat_doctor     主治医师    1   S
            mclsXML.appendData "change_treat_doctor", NeedName(cbo主治医师.Text), xsString
            'change_duty_nurse       责任护士    1   S
            mclsXML.appendData "change_duty_nurse", NeedName(cbo责任护士.Text), xsString
            'change_operator         操作员      1   S
            mclsXML.appendData "change_operator", UserInfo.姓名, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_007", mclsXML.XmlText)
        End If
    End Select
End Function

Private Sub InitStructAddress()
'功能:根据是否启用结构化地址调整界面
    Dim i As Long
    
    If gbln启用结构化地址 Then
        For i = PatiAddress.LBound To PatiAddress.UBound
             PatiAddress(i).Visible = True
             PatiAddress(i).ShowTown = gbln显示乡镇
        Next
        txt家庭地址.Visible = False
        cmd家庭地址.Visible = False
        txt出生地点.Visible = False
        cmd出生地点.Visible = False
        txt户口地址.Visible = False
        cmd户口地址.Visible = False
        txt籍贯.Visible = False
        cmd籍贯.Visible = False
        txt联系人地址.Visible = False
        cmd联系人地址.Visible = False
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
             PatiAddress(i).Visible = False
        Next
        
        txt家庭地址.Visible = True
        cmd家庭地址.Visible = True
        txt出生地点.Visible = True
        cmd出生地点.Visible = True
        txt户口地址.Visible = True
        cmd户口地址.Visible = True
        txt籍贯.Visible = True
        cmd籍贯.Visible = True
        txt联系人地址.Visible = True
        cmd联系人地址.Visible = True
    End If
End Sub

Private Function CanFocus(ctlError As Control) As Boolean
    CanFocus = ctlError.Enabled And ctlError.Visible
End Function

Private Sub EMPI_LoadPati()
'功能:将EMPI返回来的病人信息更新到界面
    Dim rsPatiIn As ADODB.Recordset
    Dim rsPatiOut As ADODB.Recordset
    Dim str出生日期 As String
    Dim blnRet As Boolean
    
    If CreatePlugInOK(glngModul) Then
        '组织病人基本信息
        Set rsPatiIn = New ADODB.Recordset
        With rsPatiIn.Fields
            .Append "病人ID", adBigInt
            .Append "主页ID", adBigInt
            .Append "挂号ID", adBigInt
            '-------------------------------
            .Append "门诊号", adVarChar, 18
            .Append "住院号", adVarChar, 18
            .Append "医保号", adVarChar, 30
            .Append "身份证号", adVarChar, 18
            .Append "其他证件", adVarChar, 20
            .Append "姓名", adVarChar, 100
            .Append "性别", adVarChar, 4
            .Append "出生日期", adVarChar, 20 '日期格式：YYYY-MM-DD HH:MM:SS
            .Append "出生地点", adVarChar, 100
            .Append "国籍", adVarChar, 30
            .Append "民族", adVarChar, 20
            .Append "学历", adVarChar, 10
            .Append "职业", adVarChar, 80
            .Append "工作单位", adVarChar, 100
            .Append "邮箱", adVarChar, 30
            .Append "婚姻状况", adVarChar, 4
            .Append "家庭电话", adVarChar, 20
            .Append "联系人电话", adVarChar, 20
            .Append "单位电话", adVarChar, 20
            .Append "家庭地址", adVarChar, 100
            .Append "家庭地址邮编", adVarChar, 6
            .Append "户口地址", adVarChar, 100
            .Append "户口地址邮编", adVarChar, 6
            .Append "单位邮编", adVarChar, 6
            .Append "联系人地址", adVarChar, 100
            .Append "联系人关系", adVarChar, 30
            .Append "联系人姓名", adVarChar, 64
        End With
        rsPatiIn.CursorLocation = adUseClient
        rsPatiIn.LockType = adLockOptimistic
        rsPatiIn.CursorType = adOpenStatic
        rsPatiIn.Open

        With rsPatiIn
            .AddNew
            !病人ID = mlng病人ID
            !主页ID = mlng主页ID
            !住院号 = Trim(txt住院号.Text)
            '-要更新的字段--------------------------------------------
            !身份证号 = Trim(txt身份证号.Text)
            !姓名 = Trim(txt姓名.Text)
            !性别 = zlCommFun.GetNeedName(txt性别.Text)
            !出生地点 = Trim(txt出生地点.Text)
            !学历 = zlCommFun.GetNeedName(cbo学历.Text)
            !职业 = zlCommFun.GetNeedName(cbo职业.Text)
            !工作单位 = Trim(txt单位地址.Text)
            !婚姻状况 = zlCommFun.GetNeedName(cbo婚姻状况.Text)
            !家庭电话 = Trim(txt家庭电话.Text)
            !联系人电话 = Trim(txt联系人电话.Text)
            !单位电话 = Trim(txt单位电话.Text)
            !家庭地址 = Trim(txt家庭地址.Text)
            !家庭地址邮编 = Trim(txt家庭地址邮编.Text)
            !户口地址 = Trim(txt户口地址.Text)
            !户口地址邮编 = Trim(txt户口地址邮编.Text)
            !单位邮编 = Trim(txt单位邮编.Text)
            !联系人地址 = Trim(txt联系人地址.Text)
            !联系人关系 = zlCommFun.GetNeedName(cbo联系人关系.Text)
            !联系人姓名 = Trim(txt联系人姓名.Text)
            .Update
            '-------------------------------------------------------
        End With
        
        '调用查询接口
        On Error Resume Next
        blnRet = gobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModul, rsPatiIn, rsPatiOut)
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: On Error GoTo 0
        If Not blnRet Then Exit Sub
        If rsPatiOut Is Nothing Then Exit Sub
        If rsPatiOut.RecordCount = 0 Then Exit Sub
        '找到病人，将病人最新的信息更新到界面
        mblnEMPI = True
        With rsPatiOut
            Call cbo.Locate(cbo学历, !学历 & "")
            Call cbo.SeekIndex(cbo职业, !职业 & "")
            Call cbo.Locate(cbo婚姻状况, !婚姻状况 & "")
            Call cbo.Locate(cbo联系人关系, !联系人关系 & "")
            
            If gbln启用结构化地址 Then
                PatiAddress(E_IX_出生地点).Value = !出生地点 & ""
                PatiAddress(E_IX_现住址).Value = !家庭地址 & ""
                PatiAddress(E_IX_户口地址).Value = !户口地址 & ""
                PatiAddress(E_IX_联系人地址).Value = !联系人地址 & ""
            End If
            '姓名,性别,年龄,出生日期 要有病人基本信息修改权限才允许更新
            txt出生地点.Text = !出生地点 & ""
            txt家庭地址.Text = !家庭地址 & ""
            txt户口地址.Text = !户口地址 & ""
            txt联系人地址.Text = !联系人地址 & ""
            txt身份证号.Text = !身份证号 & ""
            txt身份证号.Tag = !身份证号 & ""
            txt姓名.Text = !姓名 & ""
            txt单位地址.Text = !工作单位 & ""
            txt家庭电话.Text = !家庭电话 & ""
            txt联系人电话.Text = !联系人电话 & ""
            txt单位电话.Text = !单位电话 & ""
            txt家庭地址邮编.Text = !家庭地址邮编 & ""
            txt户口地址邮编.Text = !户口地址邮编 & ""
            txt单位邮编.Text = !单位邮编 & ""
            txt联系人姓名.Text = !联系人姓名 & ""
        End With
    End If
End Sub

Private Function EMPI_AddORUpdatePati(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef strErr As String) As Boolean
'功能:增加或更新EMPI病人信息
    Dim lngRet  As Long
    Dim strPlugErr As String
    Dim strTmp As String
    
    lngRet = 1 '默认成功 兼容 老版zlPlug当不支持此接口错误号:438
    If CreatePlugInOK(glngModul) Then
        If Not mblnEMPI Then
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=成功;0-失败
            Call zlPlugInErrH(Err, "EMPI_AddPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "向EMPI平台新增病人信息失败！"
        Else
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=成功;0-失败
            Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "向EMPI平台更新病人信息失败！"
        End If
        If strPlugErr <> "" Then
            strErr = strTmp & vbCrLf & strPlugErr
             Exit Function
        ElseIf lngRet = 0 Then
            strErr = strTmp & vbCrLf & strErr
            Exit Function
        End If
    End If
    
    EMPI_AddORUpdatePati = True
End Function

