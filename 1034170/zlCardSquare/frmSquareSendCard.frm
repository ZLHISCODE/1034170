VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareSendCard 
   Caption         =   "消费卡发放"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13125
   Icon            =   "frmSquareSendCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   13125
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   61
      Top             =   10200
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSquareSendCard.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18071
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   13680
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareSendCard.frx":70E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareSendCard.frx":743A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCard 
      BorderStyle     =   0  'None
      Height          =   9240
      Left            =   2400
      ScaleHeight     =   9240
      ScaleWidth      =   11010
      TabIndex        =   62
      Top             =   600
      Width           =   11010
      Begin VB.PictureBox picCardInfor 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   8700
         Left            =   120
         ScaleHeight     =   8700
         ScaleWidth      =   10410
         TabIndex        =   63
         Top             =   240
         Width           =   10410
         Begin VB.Frame fraBaseInfor 
            Caption         =   "发卡基本信息"
            Height          =   4215
            Left            =   75
            TabIndex        =   33
            Top             =   165
            Width           =   7725
            Begin VB.CommandButton cmdSel 
               Caption         =   "…"
               Height          =   270
               Index           =   2
               Left            =   7335
               TabIndex        =   17
               TabStop         =   0   'False
               Tag             =   "领卡部门"
               Top             =   2010
               Width           =   285
            End
            Begin VB.CommandButton cmdSel 
               Caption         =   "…"
               Height          =   270
               Index           =   1
               Left            =   3105
               TabIndex        =   14
               TabStop         =   0   'False
               Tag             =   "领卡人"
               Top             =   2040
               Width           =   285
            End
            Begin VB.CommandButton cmdSel 
               Caption         =   "…"
               Height          =   270
               Index           =   0
               Left            =   7335
               TabIndex        =   11
               TabStop         =   0   'False
               Tag             =   "发卡原因"
               Top             =   1605
               Width           =   285
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   21
               Left            =   5340
               TabIndex        =   31
               Top             =   3630
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   22
               Left            =   1095
               TabIndex        =   29
               Top             =   3630
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   18
               Left            =   5340
               TabIndex        =   16
               Top             =   2010
               Width           =   2295
            End
            Begin VB.CheckBox chk是否充值 
               Caption         =   "是否充值卡"
               Height          =   450
               Left            =   5355
               TabIndex        =   4
               Top             =   690
               Width           =   1830
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   9
               Left            =   1095
               TabIndex        =   25
               Top             =   3225
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   8
               Left            =   5340
               TabIndex        =   23
               Top             =   2835
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   7
               Left            =   1095
               TabIndex        =   21
               Top             =   2835
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   6
               Left            =   1110
               TabIndex        =   19
               Top             =   2430
               Width           =   6525
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   4
               Left            =   1110
               TabIndex        =   13
               Top             =   2025
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   3
               Left            =   1125
               MaxLength       =   50
               TabIndex        =   10
               Top             =   1605
               Width           =   6525
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   5340
               PasswordChar    =   "*"
               TabIndex        =   8
               Top             =   1200
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   1110
               PasswordChar    =   "*"
               TabIndex        =   6
               Top             =   1200
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   0
               Left            =   1125
               TabIndex        =   3
               Top             =   795
               Width           =   2280
            End
            Begin VB.ComboBox cbo卡类型 
               Height          =   300
               Left            =   1125
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   345
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker dtp卡有效日期 
               Height          =   300
               Left            =   5340
               TabIndex        =   27
               Top             =   3240
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   529
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   111280131
               CurrentDate     =   40156.0854282407
            End
            Begin VB.TextBox txtEdit 
               Height          =   300
               Index           =   5
               Left            =   5340
               TabIndex        =   64
               Top             =   3240
               Width           =   2295
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "回收日期"
               Height          =   180
               Index           =   11
               Left            =   4530
               TabIndex        =   30
               Top             =   3690
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "回收人"
               Height          =   180
               Index           =   21
               Left            =   495
               TabIndex        =   28
               Top             =   3690
               Width           =   540
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "领卡部门(&M)"
               Height          =   180
               Index           =   20
               Left            =   4260
               TabIndex        =   15
               Top             =   2070
               Width           =   990
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "当前余额"
               Height          =   180
               Index           =   10
               Left            =   330
               TabIndex        =   24
               Top             =   3285
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "发卡日期"
               Height          =   180
               Index           =   9
               Left            =   4530
               TabIndex        =   22
               Top             =   2895
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "发卡人"
               Height          =   180
               Index           =   8
               Left            =   510
               TabIndex        =   20
               Top             =   2895
               Width           =   540
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "备注(&S)"
               Height          =   180
               Index           =   7
               Left            =   435
               TabIndex        =   18
               Top             =   2490
               Width           =   630
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "卡有效日期(&D)"
               Height          =   180
               Index           =   6
               Left            =   4080
               TabIndex        =   26
               Top             =   3285
               Width           =   1170
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "领卡人(&D)"
               Height          =   180
               Index           =   5
               Left            =   255
               TabIndex        =   12
               Top             =   2085
               Width           =   810
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "发卡原因(&Y)"
               Height          =   180
               Index           =   4
               Left            =   90
               TabIndex        =   9
               Top             =   1650
               Width           =   990
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "密码确认(&E)"
               Height          =   180
               Index           =   3
               Left            =   4260
               TabIndex        =   7
               Top             =   1260
               Width           =   990
            End
            Begin VB.Label lblEdit 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "密码(&W)"
               Height          =   180
               Index           =   2
               Left            =   435
               TabIndex        =   5
               Top             =   1245
               Width           =   630
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "卡号(&N)"
               Height          =   180
               Index           =   1
               Left            =   450
               TabIndex        =   2
               Top             =   840
               Width           =   630
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "卡类型(&T)"
               Height          =   180
               Index           =   0
               Left            =   270
               TabIndex        =   0
               Top             =   405
               Width           =   810
            End
         End
         Begin VB.Frame fra面值 
            Caption         =   "卡面值情况"
            Height          =   705
            Left            =   90
            TabIndex        =   53
            Top             =   4440
            Width           =   7710
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   11
               Left            =   5325
               TabIndex        =   38
               Top             =   270
               Width           =   2295
            End
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   10
               Left            =   1110
               TabIndex        =   36
               Top             =   270
               Width           =   2295
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "实际销售额(&J)"
               Height          =   180
               Index           =   13
               Left            =   4080
               TabIndex        =   37
               Top             =   330
               Width           =   1170
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "卡面额(&M)"
               Height          =   180
               Index           =   12
               Left            =   270
               TabIndex        =   35
               Top             =   330
               Width           =   810
            End
         End
         Begin VB.Frame fra充值情况 
            Caption         =   "充值信息情况"
            Height          =   1950
            Left            =   90
            TabIndex        =   66
            Tag             =   "1"
            Top             =   5235
            Width           =   10125
            Begin VB.TextBox txtEdit 
               Height          =   315
               Index           =   27
               Left            =   7185
               MaxLength       =   30
               TabIndex        =   46
               Tag             =   "1"
               Top             =   1470
               Width           =   2280
            End
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   14
               Left            =   7890
               TabIndex        =   72
               Top             =   330
               Width           =   2040
            End
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   12
               Left            =   4320
               TabIndex        =   71
               Top             =   330
               Width           =   1635
            End
            Begin VB.TextBox txtEdit 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   13
               Left            =   1230
               TabIndex        =   70
               Top             =   323
               Width           =   1515
            End
            Begin VB.ComboBox cboStyle 
               Height          =   300
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Tag             =   "1"
               Top             =   1110
               Width           =   1965
            End
            Begin VB.TextBox txtEdit 
               Height          =   315
               Index           =   25
               Left            =   4320
               MaxLength       =   50
               TabIndex        =   44
               Tag             =   "1"
               Top             =   1110
               Width           =   5640
            End
            Begin VB.TextBox txtEdit 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Index           =   26
               Left            =   1200
               MaxLength       =   20
               TabIndex        =   45
               Tag             =   "1"
               Top             =   1470
               Width           =   4905
            End
            Begin VB.TextBox txtEdit 
               Height          =   315
               Index           =   24
               Left            =   4305
               TabIndex        =   42
               Top             =   698
               Width           =   5640
            End
            Begin VB.TextBox txtEdit 
               Height          =   315
               Index           =   23
               Left            =   1230
               TabIndex        =   40
               Top             =   698
               Width           =   1695
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "结算号码"
               Height          =   180
               Left            =   6360
               TabIndex        =   77
               Top             =   1530
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "本次充值(&B)"
               Height          =   180
               Index           =   14
               Left            =   3150
               TabIndex        =   76
               Top             =   390
               Width           =   990
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "充值扣率(&K)"
               Height          =   180
               Index           =   15
               Left            =   135
               TabIndex        =   75
               Top             =   390
               Width           =   990
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "实际充值缴款(&I)"
               Height          =   180
               Index           =   16
               Left            =   6480
               TabIndex        =   74
               Top             =   390
               Width           =   1350
            End
            Begin VB.Label lblPer 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   2790
               TabIndex        =   73
               Top             =   375
               Width           =   120
            End
            Begin VB.Label lblzffs 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "支付方式"
               Height          =   240
               Left            =   405
               TabIndex        =   69
               Top             =   1140
               Width           =   720
            End
            Begin VB.Label lblkhh 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "开户行"
               Height          =   240
               Left            =   3720
               TabIndex        =   68
               Top             =   1140
               Width           =   600
            End
            Begin VB.Label lblzh 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "帐号"
               Height          =   240
               Left            =   720
               TabIndex        =   67
               Top             =   1500
               Width           =   480
            End
            Begin VB.Label lblEdit 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "充值说明(&Z)"
               Height          =   180
               Index           =   23
               Left            =   3255
               TabIndex        =   41
               Top             =   765
               Width           =   990
            End
            Begin VB.Label lblEdit 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "缴款人(&R)"
               Height          =   180
               Index           =   22
               Left            =   315
               TabIndex        =   39
               Top             =   765
               Width           =   810
            End
         End
         Begin VB.Frame fra缴款 
            Caption         =   "本次缴款情况"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   120
            TabIndex        =   65
            Top             =   7440
            Width           =   10095
            Begin VB.TextBox txtEdit 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   17
               Left            =   4320
               TabIndex        =   50
               Text            =   "12"
               ToolTipText     =   "本次缴款"
               Top             =   390
               Width           =   2010
            End
            Begin VB.TextBox txtEdit 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   16
               Left            =   825
               TabIndex        =   48
               Text            =   "12"
               ToolTipText     =   "实收合计"
               Top             =   405
               Width           =   2010
            End
            Begin VB.TextBox txtEdit 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   15
               Left            =   7770
               TabIndex        =   52
               Text            =   "12"
               ToolTipText     =   "本次缴款"
               Top             =   405
               Width           =   2010
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "缴款(&U)"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   19
               Left            =   3060
               TabIndex        =   49
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "合计"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   18
               Left            =   105
               TabIndex        =   47
               Top             =   480
               Width           =   660
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               Caption         =   "找补"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   17
               Left            =   6960
               TabIndex        =   51
               Top             =   480
               Width           =   660
            End
         End
         Begin VB.Frame fra限制类别 
            Caption         =   "限制类别(&X)"
            Height          =   4980
            Left            =   7905
            TabIndex        =   34
            Top             =   180
            Width           =   2310
            Begin MSComctlLib.ListView lvwType 
               Height          =   4665
               Left            =   75
               TabIndex        =   32
               Top             =   240
               Width           =   2085
               _ExtentX        =   3678
               _ExtentY        =   8229
               View            =   3
               Arrange         =   1
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FlatScrollBar   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Key             =   "类型"
                  Object.Tag             =   "类型"
                  Text            =   "类型"
                  Object.Width           =   2540
               EndProperty
            End
         End
      End
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   7755
      Left            =   480
      ScaleHeight     =   7755
      ScaleWidth      =   4350
      TabIndex        =   54
      Top             =   240
      Width           =   4350
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   20
         Left            =   1005
         TabIndex        =   60
         Top             =   525
         Width           =   2055
      End
      Begin VSFlex8Ctl.VSFlexGrid vsGrid 
         Height          =   5235
         Left            =   120
         TabIndex        =   55
         Top             =   930
         Width           =   4065
         _cx             =   7170
         _cy             =   9234
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   28
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSquareSendCard.frx":778E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picImg 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   56
            Top             =   60
            Width           =   210
            Begin VB.Image imgCol 
               Height          =   195
               Left            =   0
               Picture         =   "frmSquareSendCard.frx":7B98
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   19
         Left            =   1005
         TabIndex        =   58
         Top             =   165
         Width           =   2055
      End
      Begin VB.Label lbl开始卡号 
         AutoSize        =   -1  'True
         Caption         =   "开始单号"
         Height          =   180
         Left            =   195
         TabIndex        =   57
         Top             =   225
         Width           =   720
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "结束单号"
         Height          =   180
         Left            =   195
         TabIndex        =   59
         Top             =   555
         Width           =   720
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   180
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSquareSendCard.frx":80E6
      Left            =   210
      Top             =   345
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSquareSendCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String, mintSucces As Integer
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private mCardEditType As gCardEditType
Private mlng消费卡ID As Long, mblnFirst As Boolean, mlng接口编号 As Long
Private mblnNoClick As Boolean, mblnChange As Boolean
Private mfrmMain As Form
Private mrs类别 As ADODB.Recordset
Private mrs消费卡类型 As ADODB.Recordset
Private mblnUnLoad As Boolean
Private Type TyCardInfor
    str卡号 As String
    dbl卡余额 As Double
    dbl实际销售 As Double
    dbl卡面值 As Double
    dbl折扣率 As Double
    int可否充值 As Integer
    str效期 As String
    bln允许充值 As Boolean  '对修改时有效,如果对当前发卡的充值记录存只有一条,就可以重新更改相关的充值信息,
    bln已消费 As Boolean    '消费了的,就不能更改允值属性和面值
    lng充值次数 As Long
End Type
Private mCardInfor As TyCardInfor
Private Enum mPaneID
    Pane_Cards = 1     '批量打件
    Pane_CardInfor = 2  '卡信息
End Enum
Private mPanSearch As Pane
Private Enum mtxtIdx
    idx_txt卡号 = 0
    idx_txt密码 = 1
    idx_txt确认密码 = 2
    idx_txt发卡原因 = 3
    idx_txt领卡人 = 4
    idx_txt卡有效日期 = 5
    idx_txt备注 = 6
    idx_txt发卡人 = 7
    idx_txt发卡日期 = 8
    idx_txt当前余额 = 9
    idx_txt卡面额 = 10
    idx_txt实际销售额 = 11
    idx_txt本次充值 = 12
    idx_txt充值扣率 = 13
    idx_txt实际充值缴款 = 14
    idx_txt找补 = 15
    idx_txt实收合计 = 16
    idx_txt本次缴款 = 17
    idx_txt领卡部门 = 18
    idx_txt开始卡号 = 19
    idx_txt结束卡号 = 20
    idx_txt回收人 = 22
    idx_txt回收时间 = 21
    idx_txt缴款人 = 23
    idx_txt充值备注 = 24
    idx_txt开户行 = 25
    idx_txt帐号 = 26
    idx_txt结算号码 = 27
End Enum
Private Enum mcmdIdx
    idx_cmd发卡原因 = 0
    idx_cmd领卡人 = 1
    idx_cmd领卡部门 = 2
End Enum
Private mlngSel消费ID As Long
Private Enum mlblIdx
    idx_lbl找补 = 17
End Enum
Private Const mconMenu_Edit_Affirm = 225
Private Const FM_HEIGHT = 10000  '编辑窗口的窗口大小
Private Const FM_WIDTH = 10935  '编辑窗口的窗口大小
Private Const PIC_CARD_HEIGHT = 8700    '卡片项的高度
Private Const PIC_CARD_WIDTH = 7050     '卡片项的宽度
Private Type Ty_Para
    str卡号前缀 As String
    lng卡号长度 As Long
    bln卡号密文 As Boolean
    bln缴款单打印 As Boolean
End Type
Private mTy_MoudlePara As Ty_Para
Private mblnHaveOtherCard As Boolean '是否存在其他的卡片(发卡时)
Private WithEvents mobjBrushCard As clsBrushSequareCard
Attribute mobjBrushCard.VB_VarHelpID = -1
Private mdbl实收合计 As Double
Private mobjKeyboard As Object
Private mstrTitle As String '用于窗体个性化保存的窗体名

Public Function zlShowCard(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, ByVal CardEditType As gCardEditType, ByVal lng接口编号 As Long, Optional lng消费卡ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口,查看已发卡或增加发卡或修改发卡信息
    '返回:
    '编制:刘兴洪
    '日期:2009-12-09 13:40:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs: mintSucces = 0
    mlng消费卡ID = lng消费卡ID: mCardEditType = CardEditType
    mlng接口编号 = lng接口编号

    With gTy_TestBug
        If CardEditType = gEd_发卡 Then
            .BytType = 1
        Else
            .BytType = 2
        End If
    End With
    Me.Show 1, frmMain
    zlShowCard = mintSucces > 0
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub InitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块变量
    '编制:刘兴洪
    '日期:2009-12-10 17:31:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varData As Variant
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = zlGet消费卡接口()
    With rsTemp
        .Filter = 0
        .Find "编号=" & mlng接口编号, , adSearchForward, 1
        If Not rsTemp.EOF Then
            With mTy_MoudlePara
                .str卡号前缀 = Nvl(rsTemp!前缀文本)
                .lng卡号长度 = Val(Nvl(rsTemp!卡号长度))
                .bln卡号密文 = Val(Nvl(rsTemp!是否密文)) = 1
            End With
        End If
    End With
    txtEdit(mtxtIdx.idx_txt开始卡号).MaxLength = mTy_MoudlePara.lng卡号长度
    txtEdit(mtxtIdx.idx_txt结束卡号).MaxLength = mTy_MoudlePara.lng卡号长度
    txtEdit(mtxtIdx.idx_txt卡号).MaxLength = mTy_MoudlePara.lng卡号长度

    With mTy_MoudlePara
        .bln缴款单打印 = Val(zlDatabase.GetPara("缴款单打印", glngSys, mlngModule)) = 1
    End With
End Sub
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2009-12-10 10:19:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    With dkpMan
        .ImageList = imlPaneIcons
        Set mPanSearch = .CreatePane(mPaneID.Pane_Cards, 400, 400, DockLeftOf, Nothing)
        mPanSearch.Title = "批量发卡信息": mPanSearch.Options = PaneNoCloseable
        mPanSearch.Handle = picList.hWnd
        Set objPane = .CreatePane(mPaneID.Pane_CardInfor, 400, 400, DockRightOf, mPanSearch)
        objPane.Title = "卡信息"
        If mCardEditType = gEd_发卡 Or mCardEditType = gEd_回收 Then
            objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Else
            objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
            mPanSearch.Closed = True
        End If

        objPane.Handle = picCard.hWnd
        objPane.MaxTrackSize.Width = picCard.Width \ Screen.TwipsPerPixelX
        objPane.MinTrackSize.Width = picCard.Width \ Screen.TwipsPerPixelX
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With

'    If mCardEditType = gEd_回收 Then
'        zlRestoreDockPanceToReg Me, dkpMan, "区域-回收"
'    ElseIf mCardEditType = gEd_发卡 Then
'        zlRestoreDockPanceToReg Me, dkpMan, "区域-发卡"
'    End If

End Function

Private Function CheckDepented() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据关联检查
    '入参:
    '出参:
    '返回:数据关联合法,返回true, 否则返回False
    '编制:刘兴洪
    '日期:2009-12-09 14:28:26
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle

    Set mrs类别 = zlGet收费类别
    If mrs类别.RecordCount = 0 Then
        ShowMsgbox "注意:" & vbCrLf & "   没有相关的收费项目类别,请与系统管理员联系!"
        Exit Function
    End If

    With lvwType
        .ListItems.Clear
         Do While Not mrs类别.EOF
            .ListItems.Add , "K" & Nvl(mrs类别!名称), Nvl(mrs类别!编码) & "-" & Nvl(mrs类别!名称)
            mrs类别.MoveNext
         Loop
         mrs类别.MoveFirst
    End With
    gstrSQL = "Select rownum as ID, 编码,名称, 缺省面额, 缺省折扣, 缺省标志 From 消费卡类型"
    Set mrs消费卡类型 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If mrs消费卡类型.RecordCount = 0 Then
        ShowMsgbox "注意:" & vbCrLf & "   没有设置相关的消费卡类型,请在[字典管理]中设置!"
        Exit Function
    End If

    zlComboxLoadFromRecodeset Me.Caption, mrs消费卡类型, cbo卡类型, True
    '检查是否启用了相关的刷卡程序
    Set mobjBrushCard = New clsBrushSequareCard
    Call mobjBrushCard.zlInitInterFacel(mlng接口编号)

    CheckDepented = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-10 10:24:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup


    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With

    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    cbsThis.ActiveMenuBar.Visible = False

    '快键绑定
    With cbsThis.KeyBindings
        .Add FALT, Asc("O"), mconMenu_Edit_Affirm
        .Add FALT, Asc("X"), conMenu_Edit_CardModify
     End With

    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched

    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): mcbrControl.BeginGroup = True

        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_MoveCard, "移出卡片"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Apply_AllCard, "应用于所有"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Apply_AllColumn, "应用于此列"): mcbrControl.BeginGroup = True

        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_Affirm, "确定  "): mcbrControl.BeginGroup = True
        mcbrControl.Flags = xtpFlagRightAlign
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): mcbrControl.BeginGroup = True
        mcbrControl.Flags = xtpFlagRightAlign
    End With

    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ClearCtlData(Optional blnClearVsGridData As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除控件数据
    '编制:刘兴洪
    '日期:2009-12-09 15:24:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control, lngID As Long, i As Long
    For Each ctl In Me.Controls
        If UCase(TypeName(ctl)) = "TEXTBOX" Then
            ctl.Text = ""
        End If
    Next
    For i = 1 To Me.lvwType.ListItems.count
        lvwType.ListItems(i).Checked = False
    Next
    stbThis.Panels(2).Text = ""
    dtp卡有效日期.value = Null
    If blnClearVsGridData Then
        With vsGrid
            .Rows = 2
            .Clear 1
        End With
    End If
    Call SetDefaultValue
    Call Show发卡数量
End Sub
Private Sub SetDefaultValue(Optional bln面额 As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省值
    '编制:刘兴洪
    '日期:2009-12-09 16:30:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, blnDefault充值 As Boolean
    '设置缺省的值
    lngID = cbo卡类型.ItemData(cbo卡类型.ListIndex)
    mrs消费卡类型.Filter = 0
    If mrs消费卡类型.RecordCount <> 0 Then mrs消费卡类型.MoveFirst
    mrs消费卡类型.Find "ID=" & lngID, , adSearchForward, 0

    If Not mrs消费卡类型.EOF Then
'        If Val(txtEdit(idx_txt充值扣率).Tag) <> 0 Or Val(txtEdit(idx_txt充值扣率).Text) = 0 Then
            txtEdit(idx_txt充值扣率).Text = Format(Val(Nvl(mrs消费卡类型!缺省折扣, 100)), "0.00")
            txtEdit(idx_txt充值扣率).Tag = txtEdit(idx_txt充值扣率).Text
            txtEdit(idx_txt实际充值缴款).Text = Format(Val(Nvl(mrs消费卡类型!缺省折扣, 100)) * Val(txtEdit(idx_txt本次充值).Text) / 100, "0.00")
'        End If
        If bln面额 Then
            txtEdit(idx_txt卡面额).Text = Format(Val(Nvl(mrs消费卡类型!缺省面额)), "0.00")
            txtEdit(idx_txt实际销售额).Text = Format(Val(Nvl(mrs消费卡类型!缺省面额)) * (txtEdit(idx_txt充值扣率).Text / 100), "0.00")
        End If
        Call ModifyGridMoney
    Else
        Call Calc余额
        Call Calc实收合计
    End If
End Sub
Private Sub ModifyGridMoney()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重新更改网格中的充值扣率,本次允值及余额等
    '编制：刘兴洪
    '日期：2010-03-26 13:54:24
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnDefault充值  As Boolean
    blnDefault充值 = mCardInfor.bln允许充值 And chk是否充值.value = 1

    '先计算余额
    Call Calc余额
    If mCardEditType = gEd_发卡 Then
        With vsGrid
            If Split(.Cell(flexcpData, .Row, .ColIndex("卡号")) & ",", ",")(0) = txtEdit(mtxtIdx.idx_txt卡号).Text Then
                If blnDefault充值 Then
                    .TextMatrix(.Row, .ColIndex("充值扣率")) = Trim(txtEdit(mtxtIdx.idx_txt充值扣率).Text)
                    .TextMatrix(.Row, .ColIndex("本次充值")) = Trim(txtEdit(mtxtIdx.idx_txt本次充值).Text)
                    .TextMatrix(.Row, .ColIndex("实际充值缴款")) = Trim(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text)
                End If
                .TextMatrix(.Row, .ColIndex("卡面额")) = Trim(txtEdit(idx_txt卡面额).Text)
                .TextMatrix(.Row, .ColIndex("实际销售")) = Trim(txtEdit(idx_txt实际销售额).Text)
            End If
            .TextMatrix(.Row, .ColIndex("当前余额")) = Trim(txtEdit(mtxtIdx.idx_txt当前余额).Text)
        End With
    End If
    Call Calc实收合计
End Sub


Private Sub Calc实收合计()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算实收合计
    '编制:刘兴洪
    '日期:2009-12-09 16:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl实收合计 As Double, dbl余额 As Double, i As Long
    Dim dbl缴款 As Double
    dbl实收合计 = 0

    If mCardEditType = gEd_发卡 Then
        '计算总合计
        With vsGrid
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("卡号"))) <> "" Then
                    dbl缴款 = IIf(Val(.Cell(flexcpData, i, .ColIndex("是否充值"))) <> 1, 0, Val(.TextMatrix(i, .ColIndex("实际充值缴款")))) + Val(.TextMatrix(i, .ColIndex("实际销售")))
                    dbl缴款 = dbl缴款 * Val(.TextMatrix(i, .ColIndex("发卡数量")))
                    dbl实收合计 = dbl实收合计 + dbl缴款
                End If
            Next
        End With
    ElseIf mCardEditType = gEd_修改 Then
        dbl实收合计 = dbl实收合计 + IIf(chk是否充值.value = 0, 0, Val(txtEdit(idx_txt实际充值缴款).Text)) + Val(txtEdit(idx_txt实际销售额).Text)
    Else
        dbl实收合计 = dbl实收合计 + Val(txtEdit(idx_txt实际充值缴款).Text)
    End If
    mdbl实收合计 = dbl实收合计
    txtEdit(idx_txt实收合计).Text = Format(dbl实收合计, "0.00")
    Call SetLblCatpion
End Sub
Private Sub Calc余额()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算余额信息
    '编制:刘兴洪
    '日期:2009-12-09 16:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl实收合计 As Double, dbl余额 As Double, i As Long
    If Not (mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改 Or mCardEditType = gEd_充值) Then Exit Sub
    If mCardEditType = gEd_修改 And (mCardInfor.bln已消费) Then
        Exit Sub
    End If
    If mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改 Then
        dbl余额 = Val(txtEdit(idx_txt卡面额).Text)
    End If
    dbl余额 = dbl余额 + IIf(chk是否充值.value = 0, 0, Val(txtEdit(idx_txt本次充值).Text))
    '计算卡余额:
    txtEdit(idx_txt当前余额).Text = Format(IIf(mCardEditType = gEd_修改, 0, mCardInfor.dbl卡余额) + dbl余额, "0.00")
End Sub
Private Function zlFromCardNOGetDataToCtrl(ByVal strCardNo As String, Optional blnSetBase As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡号,读取相应的卡片信息给控件
    '入参:strCardNo-当前卡号
    '     blnSetBase-设置基本的信息(如:卡类型,限制类别,是否允值卡)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-14 16:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim rsTemp As ADODB.Recordset, i As Long, strTemp As String, blnFind As Boolean
    mlngSel消费ID = 0
    mCardInfor.bln允许充值 = False
    gstrSQL = "" & _
    "   Select a.Id,a.卡类型,a.卡号,a.序号,a.可否充值,a.有效期,a.发卡原因, a.密码," & _
    "          a.发卡人,a.领卡人,to_char(a.发卡时间,'yyyy-mm-dd hh24:mi:ss') as 发卡时间, " & _
    "          a.回收人,to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间 , " & _
    "          decode(a.当前状态,2,'回收',3,'退卡','回收') as 当前状态,a.备注, " & _
    "          to_char(a.卡面金额," & gOraFmtString.FM_金额 & ") as 卡面金额 ," & _
    "          to_char(a.销售金额," & gOraFmtString.FM_金额 & ") as 销售金额 ," & _
    "          to_char(a.充值折扣率," & gOraFmtString.FM_折扣率 & ") as 充值折扣率 ," & _
    "          to_char(a.余额," & gOraFmtString.FM_金额 & ") as 余额 ," & _
    "          a.停用人,to_char(a.停用日期,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & _
    "          a.领卡部门ID,decode(b.编码,NULL,'' ,b.编码||'-'||b.名称) AS 领卡部门,a.限制类别 " & _
    "   From 消费卡目录 A ,部门表 b" & _
    "   Where A.卡号 = [1] and A.接口编号=[2] And 序号 = (Select Max(序号) From 消费卡目录 B Where 卡号 = A.卡号 and 接口编号=A.接口编号) and a.领卡部门id=b.Id(+) " & _
    "   Order by a.序号"
    Err = 0: On Error GoTo Errhand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strCardNo, mlng接口编号)

    If rsTemp.EOF Then

        If blnSetBase = True Then
        Else
            ShowMsgbox "未找到相关的消费卡记录,可能已经被他人删除了该消费卡,请检查!"
        End If
        If mCardEditType = gEd_充值 Then
            Call ClearCtlData: mlng消费卡ID = 0
         End If
        Exit Function
    End If
    If Val(Nvl(rsTemp!可否充值)) <> 1 And mCardEditType = gEd_充值 Then
        ShowMsgbox "当前卡号为:" & Nvl(rsTemp!卡号) & " 不是充值卡,不能进行充值"
        Call ClearCtlData: mlng消费卡ID = 0
        Exit Function
    End If
    mblnNoClick = True
    mlng消费卡ID = Val(rsTemp!id)
    With cbo卡类型
        .ListIndex = -1
        strTemp = Nvl(rsTemp!卡类型): blnFind = False
        For i = 0 To .ListCount - 1
            If .List(i) & ";" Like "*." & strTemp & ";" Then
                blnFind = True
                .ListIndex = i: Exit For
            End If
        Next
        If blnFind = False Then
            .AddItem strTemp
            .ListIndex = .NewIndex
        End If
    End With

    With lvwType
        strTemp = Nvl(rsTemp!限制类别): blnFind = False
        For i = 1 To .ListItems.count
            If InStr(1, "," & strTemp & ",", "," & Mid(.ListItems(i).Key, 2) & ",") > 0 Then
              .ListItems(i).Checked = True
            Else
              .ListItems(i).Checked = False
            End If
        Next
    End With

    chk是否充值.value = IIf(Val(Nvl(rsTemp!可否充值)) = 1, 1, 0)
    mlngSel消费ID = Val(Nvl(rsTemp!id))
    If blnSetBase = True Then
        If mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改 Then
             Call SetDefaultValue
        End If
        mblnNoClick = False: zlFromCardNOGetDataToCtrl = True
        Exit Function
    End If
    txtEdit(mtxtIdx.idx_txt卡号).Text = Nvl(rsTemp!卡号)
    txtEdit(mtxtIdx.idx_txt密码).Text = Nvl(rsTemp!密码): txtEdit(mtxtIdx.idx_txt确认密码).Text = Nvl(rsTemp!密码)
    txtEdit(mtxtIdx.idx_txt发卡原因).Text = Nvl(rsTemp!发卡原因)
    txtEdit(mtxtIdx.idx_txt卡有效日期).Text = Format(rsTemp!有效期, "yyyy-MM-DD")
    If txtEdit(mtxtIdx.idx_txt卡有效日期).Text >= "3000-01-01" Then
        txtEdit(mtxtIdx.idx_txt卡有效日期).Text = ""
    End If
    If txtEdit(mtxtIdx.idx_txt卡有效日期).Text <> "" Then
        dtp卡有效日期.value = CDate(txtEdit(mtxtIdx.idx_txt卡有效日期).Text)
    Else
        dtp卡有效日期.value = Empty
    End If
    txtEdit(mtxtIdx.idx_txt备注).Text = Nvl(rsTemp!备注)
    txtEdit(mtxtIdx.idx_txt发卡人).Text = Nvl(rsTemp!发卡人)
    txtEdit(mtxtIdx.idx_txt发卡日期).Text = Nvl(rsTemp!发卡时间)
    txtEdit(mtxtIdx.idx_txt领卡人).Text = Nvl(rsTemp!领卡人)
    txtEdit(mtxtIdx.idx_txt当前余额).Text = Nvl(rsTemp!余额)
    txtEdit(mtxtIdx.idx_txt卡面额).Text = Nvl(rsTemp!卡面金额)
    txtEdit(mtxtIdx.idx_txt实际销售额).Text = Nvl(rsTemp!销售金额)
    txtEdit(mtxtIdx.idx_txt实际销售额).Tag = Nvl(rsTemp!销售金额)

    txtEdit(mtxtIdx.idx_txt本次充值).Text = ""
    txtEdit(mtxtIdx.idx_txt充值扣率).Text = Nvl(rsTemp!充值折扣率)
    txtEdit(mtxtIdx.idx_txt本次充值).Text = ""
    txtEdit(mtxtIdx.idx_txt实际充值缴款).Text = ""
    txtEdit(mtxtIdx.idx_txt找补).Text = ""
    txtEdit(mtxtIdx.idx_txt实收合计).Text = ""
    txtEdit(mtxtIdx.idx_txt本次缴款).Text = ""
    txtEdit(mtxtIdx.idx_txt领卡部门).Text = Nvl(rsTemp!领卡部门)
    txtEdit(mtxtIdx.idx_txt领卡部门).Tag = Nvl(rsTemp!领卡部门ID)

    If mCardEditType = gEd_回收 Or mCardEditType = gEd_回退 Then
        txtEdit(mtxtIdx.idx_txt回收人).Text = UserInfo.姓名
        txtEdit(mtxtIdx.idx_txt回收时间).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    End If
    With mCardInfor
        .str卡号 = Nvl(rsTemp!卡号)
        .int可否充值 = IIf(Val(Nvl(rsTemp!可否充值)) = 1, 1, 0)
        .str效期 = txtEdit(mtxtIdx.idx_txt卡有效日期).Text
        .dbl卡余额 = Val(Nvl(rsTemp!余额))
        .dbl实际销售 = Val(Nvl(rsTemp!销售金额))
        .dbl卡面值 = Val(Nvl(rsTemp!卡面金额))
        .dbl折扣率 = Val(Nvl(rsTemp!充值折扣率))
        .lng充值次数 = 0
    End With
    '77845:李南春,2014/9/15,避免扎帐时多扎，修改时不允许改动充值情况
    mCardInfor.bln允许充值 = False
    If chk是否充值.value = 1 Then
        If mCardEditType = gEd_充值 Then
            mCardInfor.bln允许充值 = InStr(1, mstrPrivs, ";充值;") > 0
            txtEdit(mtxtIdx.idx_txt充值扣率).Text = ""
            Call SetDefaultValue(False)
        End If
    End If

    mblnNoClick = False
    zlFromCardNOGetDataToCtrl = True
    Exit Function
Errhand:
mblnNoClick = False
    If ErrCenter = 1 Then Resume
End Function

Private Function LoadDatatoCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据到控件
    '编制:刘兴洪
    '日期:2009-12-09 15:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, strTemp As String, blnFind As Boolean

    Err = 0: On Error GoTo Errhand:
    If mCardEditType = gEd_发卡 Then
        Call ClearCtlData
        With mCardInfor
            .str卡号 = ""
            .int可否充值 = 0
            .str效期 = "3000-01-01"
            .dbl卡余额 = 0
            .dbl实际销售 = 0
            .dbl卡面值 = 0
            .dbl折扣率 = 0
        End With
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt卡号)
        mCardInfor.bln允许充值 = InStr(1, mstrPrivs, ";充值;") > 0
        txtEdit(mtxtIdx.idx_txt发卡人).Text = UserInfo.姓名: txtEdit(mtxtIdx.idx_txt发卡日期).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        dtp卡有效日期.Visible = True: txtEdit(mtxtIdx.idx_txt卡有效日期).Visible = False
        dtp卡有效日期.MinDate = CDate(txtEdit(mtxtIdx.idx_txt发卡日期).Text)
        dtp卡有效日期.value = DateAdd("m", 1, dtp卡有效日期.MinDate)
        dtp卡有效日期.value = Null
        Call Set可否充值
        LoadDatatoCard = True: Exit Function
    
    End If
    If mCardEditType = gEd_充值 And mlng消费卡ID = 0 Then
        '充值,又没传入值,就直接退出了
        Call ClearCtlData
        With mCardInfor
            .str卡号 = ""
            .int可否充值 = 0
            .str效期 = "3000-01-01"
            .dbl卡余额 = 0
            .dbl实际销售 = 0
            .dbl卡面值 = 0
            .dbl折扣率 = 0
        End With
        dtp卡有效日期.value = Null
        Call Set可否充值
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt卡号)
        LoadDatatoCard = True: Exit Function
    End If
    If mCardEditType = gEd_回收 And mlng消费卡ID = 0 Then
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt卡号)
        LoadDatatoCard = True: Exit Function
         Exit Function
    End If
    '查看或其他操作
    gstrSQL = "" & _
    "   Select a.Id,a.卡类型,a.卡号,a.序号,a.可否充值,a.有效期,a.发卡原因, a.密码," & _
    "          a.发卡人,a.领卡人,to_char(a.发卡时间,'yyyy-mm-dd hh24:mi:ss') as 发卡时间, " & _
    "          a.回收人,to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间 , " & _
    "          decode(a.当前状态,2,'回收',3,'退卡','回收') as 当前状态,a.备注, " & _
    "          to_char(a.卡面金额," & gOraFmtString.FM_金额 & ") as 卡面金额 ," & _
    "          to_char(a.销售金额," & gOraFmtString.FM_金额 & ") as 销售金额 ," & _
    "          to_char(a.充值折扣率," & gOraFmtString.FM_折扣率 & ") as 充值折扣率 ," & _
    "          to_char(a.余额," & gOraFmtString.FM_金额 & ") as 余额 ," & _
    "          a.停用人,to_char(a.停用日期,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & _
    "          a.领卡部门ID,decode(b.编码,null,'',b.编码||'-'||b.名称) AS 领卡部门,a.限制类别 " & _
    "   From 消费卡目录 A,部门表 B " & _
    "   Where   a.领卡部门id=b.Id(+) and A.Id =[1]   " & _
    "   Order by a.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng消费卡ID)

    If rsTemp.EOF Then
        ShowMsgbox "未找到相关的消费卡记录,可能已经被他人删除了该消费卡,请检查!"
        Exit Function
    End If

    If mCardEditType = gEd_充值 And IIf(Val(Nvl(rsTemp!可否充值)) = 1, 1, 0) <> 1 Then
        '如果是充值,且又不允许允值的,给提示，退出
        ShowMsgbox "注意:" & vbCrLf & "    该卡不允许允值，不能进行充值操作！"
        Exit Function
    End If

    mblnNoClick = True
    mlngSel消费ID = mlng消费卡ID
    chk是否充值.value = IIf(Val(Nvl(rsTemp!可否充值)) = 1, 1, 0)

    With cbo卡类型
        .ListIndex = -1
        strTemp = Nvl(rsTemp!卡类型): blnFind = False
        For i = 0 To .ListCount - 1
            If .List(i) & ";" Like "*." & strTemp & ";" Then
                blnFind = True
                .ListIndex = i: Exit For
            End If
        Next
        If blnFind = False Then
            .AddItem strTemp
            .ListIndex = .NewIndex
        End If
    End With

    With lvwType
        strTemp = Nvl(rsTemp!限制类别): blnFind = False
        For i = 1 To .ListItems.count
            If InStr(1, "," & strTemp & ",", "," & Mid(.ListItems(i).Key, 2) & ",") > 0 Then
                .ListItems(i).Checked = True
            Else
                .ListItems(i).Checked = False
            End If
        Next
    End With
    txtEdit(mtxtIdx.idx_txt卡号).Text = Nvl(rsTemp!卡号)
    txtEdit(mtxtIdx.idx_txt卡号).Tag = Nvl(rsTemp!id)
    txtEdit(mtxtIdx.idx_txt密码).Text = Nvl(rsTemp!密码): txtEdit(mtxtIdx.idx_txt确认密码).Text = Nvl(rsTemp!密码)
    txtEdit(mtxtIdx.idx_txt发卡原因).Text = Nvl(rsTemp!发卡原因)
    txtEdit(mtxtIdx.idx_txt卡有效日期).Text = Format(rsTemp!有效期, "yyyy-MM-DD")
    If txtEdit(mtxtIdx.idx_txt卡有效日期).Text >= "3000-01-01" Then
        txtEdit(mtxtIdx.idx_txt卡有效日期).Text = ""
    End If
    If txtEdit(mtxtIdx.idx_txt卡有效日期).Text <> "" Then
        dtp卡有效日期.value = CDate(txtEdit(mtxtIdx.idx_txt卡有效日期).Text)
    Else
        dtp卡有效日期.value = Null
    End If
    txtEdit(mtxtIdx.idx_txt备注).Text = Nvl(rsTemp!备注)
    txtEdit(mtxtIdx.idx_txt发卡人).Text = Nvl(rsTemp!发卡人)
    txtEdit(mtxtIdx.idx_txt发卡日期).Text = Nvl(rsTemp!发卡时间)
    txtEdit(mtxtIdx.idx_txt领卡人).Text = Nvl(rsTemp!领卡人)
    txtEdit(mtxtIdx.idx_txt当前余额).Text = Format(Val(Nvl(rsTemp!余额)), gVbFmtString.FM_金额)
    txtEdit(mtxtIdx.idx_txt卡面额).Text = Format(Val(Nvl(rsTemp!卡面金额)), gVbFmtString.FM_金额)
    txtEdit(mtxtIdx.idx_txt实际销售额).Text = Format(Val(Nvl(rsTemp!销售金额)), gVbFmtString.FM_金额)
    txtEdit(mtxtIdx.idx_txt实际销售额).Tag = Val(Nvl(rsTemp!销售金额))

    txtEdit(mtxtIdx.idx_txt本次充值).Text = ""
    txtEdit(mtxtIdx.idx_txt充值扣率).Text = Format(Val(Nvl(rsTemp!充值折扣率)), "0.00")
    txtEdit(mtxtIdx.idx_txt本次充值).Text = ""
    txtEdit(mtxtIdx.idx_txt实际充值缴款).Text = ""
    txtEdit(mtxtIdx.idx_txt找补).Text = ""
    txtEdit(mtxtIdx.idx_txt实收合计).Text = ""
    txtEdit(mtxtIdx.idx_txt本次缴款).Text = ""
    txtEdit(mtxtIdx.idx_txt领卡部门).Text = Nvl(rsTemp!领卡部门)
    txtEdit(mtxtIdx.idx_txt领卡部门).Tag = Nvl(rsTemp!领卡部门ID)

    If mCardEditType = gEd_回收 Or mCardEditType = gEd_退卡 Then
        txtEdit(mtxtIdx.idx_txt回收人).Text = UserInfo.姓名
        txtEdit(mtxtIdx.idx_txt回收时间).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        txtEdit(mtxtIdx.idx_txt回收人).Text = Nvl(rsTemp!回收人)
        txtEdit(mtxtIdx.idx_txt回收时间).Text = Format(rsTemp!回收时间, "yyyy-mm-dd HH:MM:SS")
        If txtEdit(mtxtIdx.idx_txt回收时间).Text >= "3000-01-01" Then txtEdit(mtxtIdx.idx_txt回收时间).Text = ""
    End If
    With mCardInfor
        .str卡号 = Nvl(rsTemp!卡号)
        .int可否充值 = IIf(Val(Nvl(rsTemp!可否充值)) = 1, 1, 0)
        .str效期 = txtEdit(mtxtIdx.idx_txt卡有效日期).Text
        .dbl卡余额 = Val(Nvl(rsTemp!余额))
        .dbl实际销售 = Val(Nvl(rsTemp!销售金额))
        .dbl卡面值 = Val(Nvl(rsTemp!卡面金额))
        .dbl折扣率 = Val(Nvl(rsTemp!充值折扣率))
        .bln允许充值 = InStr(1, mstrPrivs, ";充值;") > 0 And (mCardEditType = gEd_充值 Or mCardEditType = gEd_发卡)
        .lng充值次数 = 0
    End With
    '77845:李南春,2014/9/15,避免扎帐时多扎，修改时不允许改动充值情况
    If mCardEditType = gEd_回收 Then
        InsertIntoGrid txtEdit(mtxtIdx.idx_txt卡号).Text, False, True
    End If

    Call SetDefaultValue(False)
    Call Set可否充值

    Call SetLblCatpion

    mblnNoClick = False
    LoadDatatoCard = True
    Exit Function
Errhand:
    mblnNoClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub SetFormNOTResize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体不允许调整大小
    '编制:刘兴洪
    '日期:2009-12-10 15:15:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Me.Width = FM_WIDTH: Me.Height = FM_HEIGHT + IIf(mCardEditType = gEd_充值, 300, 0)
    Call zlSetWindowsBroldStyle(Me)  '将窗体设置成不可调
End Sub
Private Function InsertIntoGrid(ByVal strCardNo As String, Optional blnEndCard As Boolean = False, Optional blnNotCardNo As Boolean = False, Optional blnModifyCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:向网格插入指定卡的数据
    '入参:strCardNo-卡号
    '     blnEndCard=true:是在结束的文本框刷卡,false:在开始的文本框刷卡
    '     blnNotCardNo-是否不读取原来的卡片信息(对回收有效)
    '     blnModifyCard-是否修改改号
    '返回:插入成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2009-12-14 11:20:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, strCurCard As String, strCards As String, varData As Variant, i As Long, lng数量 As Long
    Dim strCardRange As String

    Err = 0: On Error GoTo Errhand:
    lng数量 = 1

    If mCardEditType = gEd_回收 Then
        '回收时,根据数据库的信息来确定
        If Not blnNotCardNo Then
            If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        End If
    Else
        Call zlFromCardNOGetDataToCtrl(strCardNo, True)
    End If
    
    '92796:李南春,2016/1/20,允许消费卡卡号长度小于限定长度
    '先检查输入项是否合法:
    If CheckInput = False Then Exit Function
    If zlCommFun.ActualLen(strCardNo) > txtEdit(mtxtIdx.idx_txt卡号).MaxLength Then
        ShowMsgbox "卡号长度不对正确,请检查!"
        Exit Function
    End If

    '先检查数据的合法性
    With vsGrid
        For lngRow = 1 To .Rows - 1
            '先看是否已经存在刷卡的信息
            strCurCard = Trim(.Cell(flexcpData, lngRow, .ColIndex("卡号")))
            If InStr(1, "," & strCurCard & ",", "," & strCardNo & ",") > 0 And lngRow <> .Row Then
                '卡号已经存在,不应该再插入
                ShowMsgbox "注意:" & vbCrLf & "    在第" & lngRow & "行中,已经存在此消费卡了(卡号:" & strCardNo & "),不能再继续!"
                Exit Function
            End If
            If .Row = lngRow Then

                If strCurCard = strCardNo Then
                    '同行的,肯定不能再插入了
                    InsertIntoGrid = True: Exit Function
                End If
                If InStr(1, strCurCard, ",") > 0 And InStr(1, "," & strCurCard & ",", "," & strCardNo & ",") > 0 Then
                    '卡号已经存在,不应该再插入
                    ShowMsgbox "注意:" & vbCrLf & "    在第" & lngRow & "行中,已经存在此消费卡了(卡号:" & strCardNo & "),不能再继续!"
                    Exit Function

                End If
                If blnEndCard Then
                    '检查结束单号是否正确
                    If Split(.TextMatrix(lngRow, .ColIndex("卡号")) & "～", "～")(1) = strCardNo Then
                         InsertIntoGrid = True: Exit Function
                    End If

                End If
            End If
        Next
        '如果是在一个范围内刷卡,也需要检查
        If blnEndCard Then
            strCurCard = Trim(.Cell(flexcpData, .Row, .ColIndex("卡号")))
            If InStr(1, strCurCard, ",") <= 0 Then  '本身就是一个范围,就当成新的刷卡记录处理
                If strCurCard < strCardNo Then
                  strCardRange = strCurCard & "～" & strCardNo

                  If zlCardNoRange(strCardRange, strCards) = False Then Exit Function
                  varData = Split(strCards, ","): lng数量 = UBound(varData) + 1
                  For i = 0 To UBound(varData)
                    '检查是否有重复的
                    For lngRow = 1 To .Rows - 1
                        '先看是否已经存在刷卡的信息
                        strCurCard = Trim(.Cell(flexcpData, lngRow, .ColIndex("卡号")))

                        If InStr(1, "," & strCurCard & ",", "," & strCardNo & ",") > 0 And lngRow <> .Row Then
                            '卡号已经存在,不应该再插入
                            ShowMsgbox "注意:" & vbCrLf & "    在第" & lngRow & "行中,已经存在此消费卡了(卡号:" & strCardNo & "),不能再继续!"
                            Exit Function
                        End If
                    Next
                  Next
                    '检查是否已经包含了存在的卡片信息,如果存在的卡片信息,需要检查:
                    '1. 卡类型不一致的,不能发.
                    '2. 是否充值不一致的,也不能发
                    '3.其他的些检查(比如:已经发了卡的,不能再次发卡)
                    If 检查卡号是否合法(strCards, True, True) = False Then Exit Function
                Else
                    '开始号大于结束号, 也只默认为新的号
                    blnEndCard = False
                End If
            Else
                blnEndCard = False
            End If
        End If

        If blnEndCard = False Then
            '检查是否已经包含了存在的卡片信息,如果存在的卡片信息,需要检查:
            '1. 卡类型不一致的,不能发.
            '2. 是否充值不一致的,也不能发
            '3.其他的些检查(比如:已经发了卡的,不能再次发卡)
            If 检查卡号是否合法(strCardNo, True, True) = False Then Exit Function
        End If
        mblnNoClick = True
        '可以增加数据了
        ' 根据界面中设置的值进行处理

        If blnEndCard Then
            '在原来的基础上加上范围
            .TextMatrix(.Row, .ColIndex("卡号")) = strCardRange & "   (共" & lng数量 & "张卡)"
            .Cell(flexcpData, .Row, .ColIndex("卡号")) = strCards
        Else
            If .TextMatrix(.Row, .ColIndex("卡号")) <> "" And blnModifyCard = False Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
            .TextMatrix(.Row, .ColIndex("卡号")) = strCardNo
            .Cell(flexcpData, .Row, .ColIndex("卡号")) = strCardNo
            If .RowIsVisible(.Row) = False Then .TopRow = .Row
        End If
        .TextMatrix(.Row, .ColIndex("ID")) = mlngSel消费ID
        .TextMatrix(.Row, .ColIndex("卡类型")) = Mid(cbo卡类型.Text, InStr(cbo卡类型.Text, ".") + 1)
        .TextMatrix(.Row, .ColIndex("是否充值")) = IIf(chk是否充值.value = 1, "√", ""): .Cell(flexcpData, .Row, .ColIndex("是否充值")) = IIf(chk是否充值.value = 1, 1, 0)
        .TextMatrix(.Row, .ColIndex("密码")) = "******************": .Cell(flexcpData, .Row, .ColIndex("密码")) = Trim(txtEdit(mtxtIdx.idx_txt密码).Text)
        .Cell(flexcpData, .Row, .ColIndex("ID")) = Trim(txtEdit(mtxtIdx.idx_txt确认密码).Text)
        .TextMatrix(.Row, .ColIndex("发卡原因")) = Trim(txtEdit(mtxtIdx.idx_txt发卡原因).Text)
        .TextMatrix(.Row, .ColIndex("领卡人")) = Trim(txtEdit(mtxtIdx.idx_txt领卡人).Text)
        .TextMatrix(.Row, .ColIndex("领卡部门")) = Trim(txtEdit(mtxtIdx.idx_txt领卡部门).Text): .Cell(flexcpData, .Row, .ColIndex("领卡部门")) = Val(txtEdit(mtxtIdx.idx_txt领卡部门).Tag)

        .TextMatrix(.Row, .ColIndex("备注")) = Trim(txtEdit(mtxtIdx.idx_txt备注).Text)
        .TextMatrix(.Row, .ColIndex("发卡人")) = Trim(txtEdit(mtxtIdx.idx_txt发卡人).Text)
        If Trim(.TextMatrix(.Row, .ColIndex("发卡人"))) = "" Then .TextMatrix(.Row, .ColIndex("发卡人")) = UserInfo.姓名
        .TextMatrix(.Row, .ColIndex("发卡日期")) = Trim(txtEdit(mtxtIdx.idx_txt发卡日期).Text)

        .TextMatrix(.Row, .ColIndex("当前余额")) = Trim(txtEdit(mtxtIdx.idx_txt当前余额).Text)
        .TextMatrix(.Row, .ColIndex("卡有效期")) = Format(dtp卡有效日期.value, "yyyy-mm-dd HH:MM")
        .TextMatrix(.Row, .ColIndex("卡面额")) = Trim(txtEdit(mtxtIdx.idx_txt卡面额).Text)
        .TextMatrix(.Row, .ColIndex("实际销售")) = Trim(txtEdit(mtxtIdx.idx_txt实际销售额).Text)

        If chk是否充值.value = 0 Then
            .TextMatrix(.Row, .ColIndex("充值扣率")) = ""
            .TextMatrix(.Row, .ColIndex("本次充值")) = ""
            .TextMatrix(.Row, .ColIndex("实际充值缴款")) = ""
        Else
            .TextMatrix(.Row, .ColIndex("充值扣率")) = Trim(txtEdit(mtxtIdx.idx_txt充值扣率).Text)
            .TextMatrix(.Row, .ColIndex("本次充值")) = Trim(txtEdit(mtxtIdx.idx_txt本次充值).Text)
            .TextMatrix(.Row, .ColIndex("实际充值缴款")) = Trim(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text)
        End If
        .TextMatrix(.Row, .ColIndex("限制类别")) = Get限制类别
        .TextMatrix(.Row, .ColIndex("充值说明")) = Trim(txtEdit(mtxtIdx.idx_txt充值备注).Text)
        .TextMatrix(.Row, .ColIndex("充值缴款人")) = Trim(txtEdit(mtxtIdx.idx_txt缴款人).Text)
        .TextMatrix(.Row, .ColIndex("发卡数量")) = lng数量
    End With
    Call Show发卡数量
    '检查是否存在其他单据信息
    Call CheckOtherCard
    Call Calc实收合计


    mblnNoClick = False
    InsertIntoGrid = True
    Exit Function
Errhand:
    mblnNoClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub Show发卡数量()
    Dim lngRow As Long, lngNum As Long
    If mCardEditType <> gEd_发卡 And mCardEditType <> gEd_回收 Then Exit Sub
    lngNum = 0
    With vsGrid
        For lngRow = 1 To .Rows - 1
            lngNum = lngNum + Val(.TextMatrix(lngRow, .ColIndex("发卡数量")))
        Next
    End With
    stbThis.Panels(2).Text = "本次共" & IIf(mCardEditType <> gEd_回收, "发", "回收") & lngNum & "张卡片"
End Sub
Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗口大小
    '编制:刘兴洪
    '日期:2010-01-04 16:40:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
  Select Case mCardEditType
    Case gEd_发卡
    Case gEd_修改
        Call SetFormNOTResize   '设置成不可调大小
        stbThis.Visible = False
    Case gEd_充值
         Call SetFormNOTResize   '设置成不可调大小
        stbThis.Visible = False
    Case gEd_回收
    Case gEd_取消回收
        Call SetFormNOTResize   '设置成不可调大小
        stbThis.Visible = False
    Case gEd_回退
        Call SetFormNOTResize   '设置成不可调大小
        stbThis.Visible = False
    Case gEd_退卡, gEd_取消退卡
        Call SetFormNOTResize   '设置成不可调大小
        stbThis.Visible = False
    Case gEd_查询
        Call SetFormNOTResize   '设置成不可调大小
        stbThis.Visible = False
    End Select
End Sub

Private Sub SetEditProperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置编辑属性
    '编制:刘兴洪
    '日期:2009-12-09 16:54:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    txtEdit(mtxtIdx.idx_txt发卡人).Enabled = False: txtEdit(mtxtIdx.idx_txt发卡日期).Enabled = False
    txtEdit(mtxtIdx.idx_txt实收合计).Locked = True: txtEdit(mtxtIdx.idx_txt找补).Locked = True
    txtEdit(mtxtIdx.idx_txt当前余额).Enabled = False
    txtEdit(mtxtIdx.idx_txt缴款人).Enabled = False
    txtEdit(mtxtIdx.idx_txt回收人).Enabled = False: txtEdit(mtxtIdx.idx_txt回收时间).Enabled = False
    cmdSel(mcmdIdx.idx_cmd发卡原因).Visible = False
    cmdSel(mcmdIdx.idx_cmd领卡人).Visible = False
    cmdSel(mcmdIdx.idx_cmd领卡部门).Visible = False

    Select Case mCardEditType
    Case gEd_发卡
        txtEdit(mtxtIdx.idx_txt卡面额).Enabled = zlCheckPrivs(mstrPrivs, "允许更改卡面额")
        txtEdit(mtxtIdx.idx_txt实际销售额).Enabled = txtEdit(mtxtIdx.idx_txt卡面额).Enabled
        cmdSel(mcmdIdx.idx_cmd发卡原因).Visible = True
        cmdSel(mcmdIdx.idx_cmd领卡人).Visible = True
        cmdSel(mcmdIdx.idx_cmd领卡部门).Visible = True
        If mCardInfor.bln允许充值 Then GoTo DoSetColor:
        fra缴款.Visible = txtEdit(mtxtIdx.idx_txt卡面额).Enabled
        If txtEdit(mtxtIdx.idx_txt卡面额).Enabled = False Then picCardInfor.Height = PIC_CARD_HEIGHT - fra缴款.Height
       ' txtEdit(mtxtIdx.idx_txt卡面额).Enabled = False: txtEdit(mtxtIdx.idx_txt实际销售额).Enabled = False
        txtEdit(mtxtIdx.idx_txt本次充值).Enabled = False: txtEdit(mtxtIdx.idx_txt实际充值缴款).Enabled = False
        txtEdit(mtxtIdx.idx_txt充值扣率).Enabled = False:

    Case gEd_修改
        txtEdit(mtxtIdx.idx_txt卡面额).Enabled = False ' zlCheckPrivs(mstrPrivs, "允许更改卡面额") And Not mCardInfor.bln已消费
        txtEdit(mtxtIdx.idx_txt实际销售额).Enabled = txtEdit(mtxtIdx.idx_txt卡面额).Enabled
        txtEdit(mtxtIdx.idx_txt领卡人).Enabled = True: txtEdit(mtxtIdx.idx_txt领卡部门).Enabled = True
        txtEdit(mtxtIdx.idx_txt发卡原因).Enabled = True:
        txtEdit(mtxtIdx.idx_txt卡号).Enabled = False
        cmdSel(mcmdIdx.idx_cmd发卡原因).Visible = True
        cmdSel(mcmdIdx.idx_cmd领卡人).Visible = True
        cmdSel(mcmdIdx.idx_cmd领卡部门).Visible = True

        chk是否充值.Enabled = (Not mCardInfor.bln已消费) And mCardInfor.lng充值次数 <= 1
        '77845:李南春,2014/9/15,避免扎帐时多扎，修改时不允许改动充值情况
         fra充值情况.Visible = False
         fra缴款.Visible = txtEdit(mtxtIdx.idx_txt卡面额).Enabled:

        If txtEdit(mtxtIdx.idx_txt卡面额).Enabled = False Then
            picCardInfor.Height = PIC_CARD_HEIGHT - fra缴款.Height - fra充值情况.Height - 300
            Me.Height = FM_HEIGHT - fra缴款.Height - fra充值情况.Height - 300
        Else
            fra缴款.Top = fra充值情况.Top
            picCardInfor.Height = PIC_CARD_HEIGHT - fra充值情况.Height - 300
            Me.Height = FM_HEIGHT - fra充值情况.Height - 300
        End If
        txtEdit(mtxtIdx.idx_txt本次充值).Enabled = False: txtEdit(mtxtIdx.idx_txt实际充值缴款).Enabled = False
        txtEdit(mtxtIdx.idx_txt充值扣率).Enabled = False

    Case gEd_充值
         For Each ctl In Controls
            Select Case UCase(TypeName(ctl))
            Case "TEXTBOX", "COMBOBOX"
                If Not ctl Is txtEdit(mtxtIdx.idx_txt卡号) Then
                    ctl.Enabled = False
                End If
            Case "CHECKBOX", "LISTVIEW", "LISTBOX"
                ctl.Enabled = False
            End Select
         Next
        txtEdit(mtxtIdx.idx_txt卡号).Enabled = True
        txtEdit(mtxtIdx.idx_txt实收合计).Enabled = True: txtEdit(mtxtIdx.idx_txt找补).Enabled = True
        txtEdit(mtxtIdx.idx_txt本次充值).Enabled = True: txtEdit(mtxtIdx.idx_txt实际充值缴款).Enabled = True
        txtEdit(mtxtIdx.idx_txt充值扣率).Enabled = True: dtp卡有效日期.Visible = False
        txtEdit(mtxtIdx.idx_txt缴款人).Enabled = True: txtEdit(mtxtIdx.idx_txt充值备注).Enabled = True
        Call Set可否充值

    Case gEd_回收
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        dtp卡有效日期.Visible = False
        fra缴款.Visible = False
        fra充值情况.Visible = False
        If Me.WindowState = 0 Then Me.Height = FM_HEIGHT - fra缴款.Height - fra充值情况.Height
        picCardInfor.Height = PIC_CARD_HEIGHT - fra缴款.Height - fra充值情况.Height
        
        txtEdit(mtxtIdx.idx_txt卡号).Enabled = True: txtEdit(mtxtIdx.idx_txt开始卡号).Visible = False: txtEdit(mtxtIdx.idx_txt结束卡号).Visible = False
        lbl开始卡号.Visible = False: lbl至.Visible = False
        Call picList_Resize
    Case gEd_取消回收
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        fra缴款.Visible = False
        dtp卡有效日期.Visible = False
        picCardInfor.Height = PIC_CARD_HEIGHT - fra缴款.Height - 300
        Me.Height = FM_HEIGHT - fra缴款.Height - 300

    Case gEd_回退
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        fra缴款.Visible = False
        picCardInfor.Height = PIC_CARD_HEIGHT - fra缴款.Height - 300
        Me.Height = FM_HEIGHT - fra缴款.Height - 300
    Case gEd_退卡, gEd_取消退卡
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        dtp卡有效日期.Visible = False
        fra缴款.Visible = False
        picCardInfor.Height = PIC_CARD_HEIGHT - fra缴款.Height - 300
        Me.Height = FM_HEIGHT - fra缴款.Height - 300
    Case gEd_查询
        For Each ctl In Controls
           Select Case UCase(TypeName(ctl))
           Case "TEXTBOX", "COMBOBOX"
              ctl.Enabled = False
           Case "CHECKBOX", "LISTVIEW", "LISTBOX"
               ctl.Enabled = False
           End Select
        Next
        dtp卡有效日期.Visible = False
        fra缴款.Visible = False
        fra充值情况.Visible = False

        picCardInfor.Height = PIC_CARD_HEIGHT - fra缴款.Height - fra充值情况.Height - 300
        Me.Height = FM_HEIGHT - fra缴款.Height - fra充值情况.Height - 300
    End Select
    Call picCard_Resize
DoSetColor:
    Call SetCtlBackColor
End Sub

Private Sub cboStyle_Click()
    '当是支票时,允许输入缴款单位
    Dim blnEnabled As Boolean
    If cboStyle.ListIndex = -1 Then Exit Sub
    Call Set支付方式Enabled
End Sub

Private Sub Set支付方式Enabled()
     '当是支票时,允许输入缴款单位
    Dim blnEnabled As Boolean
    If cboStyle.ListIndex = -1 Then Exit Sub

    blnEnabled = cboStyle.ItemData(cboStyle.ListIndex) = 2 And (cboStyle.Text Like "*票*" Or cboStyle.Text Like "*卡*")
    txtEdit(mtxtIdx.idx_txt结算号码).Enabled = blnEnabled
    txtEdit(mtxtIdx.idx_txt开户行).Enabled = blnEnabled
    txtEdit(mtxtIdx.idx_txt帐号).Enabled = blnEnabled
    If Not blnEnabled Then txtEdit(mtxtIdx.idx_txt结算号码).Text = "": txtEdit(mtxtIdx.idx_txt开户行).Text = "": txtEdit(mtxtIdx.idx_txt帐号).Text = ""
    Call SetCtlBackColor
End Sub

Private Sub cbo卡类型_Click()
    If mblnNoClick Then Exit Sub
    mblnChange = True
    '重新设置缺省值
    If Not (mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改) Then Exit Sub
    If mCardEditType = gEd_修改 And mCardInfor.bln允许充值 = False Then Exit Sub
    If mCardEditType = gEd_修改 And mCardInfor.bln已消费 = True Then Exit Sub
    Call SetDefaultValue(True)
    Calc实收合计
End Sub
Private Sub Set可否充值()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置可否充值
    '编制:刘兴洪
    '日期:2009-12-17 15:11:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    fra充值情况.Enabled = chk是否充值.value = 1 And IIf(mCardEditType = gEd_发卡, InStr(1, mstrPrivs, ";充值;") > 0, mCardInfor.bln允许充值)
    txtEdit(mtxtIdx.idx_txt本次充值).Enabled = fra充值情况.Enabled
    txtEdit(mtxtIdx.idx_txt充值备注).Enabled = fra充值情况.Enabled
    txtEdit(mtxtIdx.idx_txt充值扣率).Enabled = fra充值情况.Enabled
    txtEdit(mtxtIdx.idx_txt本次缴款).Enabled = fra充值情况.Enabled Or (Val(txtEdit(mtxtIdx.idx_txt实际销售额).Text) <> 0 And (mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改))
    txtEdit(mtxtIdx.idx_txt实际充值缴款).Enabled = fra充值情况.Enabled
    cboStyle.Enabled = fra充值情况.Enabled
    Call Set支付方式Enabled
End Sub
Private Sub SetCtlBackColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的可编制颜色
    '编制:刘兴洪
    '日期:2009-12-17 15:14:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    For Each ctl In Me.Controls
        If UCase(TypeName(ctl)) = "TEXTBOX" Or UCase(TypeName(ctl)) = "COMBOBOX" Then
            Call zl_SetCtlBackColor(ctl)
        End If
    Next
End Sub

Private Sub cbo卡类型_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab

End Sub

Private Sub cbo卡类型_Validate(Cancel As Boolean)
    If mCardEditType <> gEd_发卡 Then Exit Sub

    '修改的话,需要同步更新网格行的数据才行
    With vsGrid
        If .TextMatrix(.Row, .ColIndex("卡号")) <> txtEdit(mtxtIdx.idx_txt卡号).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("卡类型")) = Mid(cbo卡类型.Text, InStr(cbo卡类型.Text, ".") + 1)
    End With
 End Sub

Private Sub chk是否充值_Click()
    If mblnNoClick Then Exit Sub
    If mCardEditType <> gEd_发卡 And mCardEditType <> gEd_修改 Then Exit Sub
    Call Set可否充值
    If mCardEditType <> gEd_发卡 Then
        Call Calc余额:         Calc实收合计
        Call SetLblCatpion
        Exit Sub
    End If
    '修改的话,需要同步更新网格行的数据才行
    With vsGrid

        If Split(.Cell(flexcpData, .Row, .ColIndex("卡号")) & ",", ",")(0) <> txtEdit(mtxtIdx.idx_txt卡号).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("是否充值")) = IIf(chk是否充值.value = 1, "√", ""): .Cell(flexcpData, .Row, .ColIndex("是否充值")) = IIf(chk是否充值.value = 1, 1, 0)
        If chk是否充值.value = 0 Then
            .TextMatrix(.Row, .ColIndex("充值扣率")) = ""
            .TextMatrix(.Row, .ColIndex("本次充值")) = ""
            .TextMatrix(.Row, .ColIndex("实际充值缴款")) = ""
        Else
            .TextMatrix(.Row, .ColIndex("充值扣率")) = Trim(txtEdit(mtxtIdx.idx_txt充值扣率).Text)
            .TextMatrix(.Row, .ColIndex("本次充值")) = Trim(txtEdit(mtxtIdx.idx_txt本次充值).Text)
            .TextMatrix(.Row, .ColIndex("实际充值缴款")) = Trim(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text)
        End If
        .TextMatrix(.Row, .ColIndex("当前余额")) = Trim(txtEdit(mtxtIdx.idx_txt当前余额).Text)
    End With
    Call Calc余额:         Calc实收合计
    Call SetLblCatpion
End Sub

Private Sub chk是否充值_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub



Private Sub cmdSel_Click(Index As Integer)
    Dim lngID As Long, str编码 As String, str名称 As String
    Select Case cmdSel(Index).Tag
    Case "领卡人"
        '选择人员
        lngID = Val(txtEdit(mtxtIdx.idx_txt领卡部门).Tag)
        If Select人员选择器(Me, txtEdit(mtxtIdx.idx_txt领卡人), "", lngID, , True) = False Then
              Exit Sub
        End If
        If mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改 Then
            '领卡人就是缴款人
            txtEdit(mtxtIdx.idx_txt缴款人).Text = txtEdit(mtxtIdx.idx_txt领卡人).Text
            txtEdit(mtxtIdx.idx_txt缴款人).Tag = txtEdit(mtxtIdx.idx_txt领卡人).Tag
        End If
        '需要读取缺省部门:
        If zl_From人员获取缺省部门(Val(txtEdit(mtxtIdx.idx_txt领卡人).Tag), str编码, str名称, lngID) Then
            txtEdit(mtxtIdx.idx_txt领卡部门).Text = str编码 & "-" & str名称
            txtEdit(mtxtIdx.idx_txt领卡部门).Tag = lngID
        End If
    Case "领卡部门"
        '选择缺省部门
        lngID = Val(txtEdit(mtxtIdx.idx_txt领卡人).Tag)
        If Select部门选择器(Me, txtEdit(mtxtIdx.idx_txt领卡部门), "", "", IIf(lngID = 0, False, True), "", 0, "部门选择器", , , , , lngID) = False Then
            Exit Sub
        End If
    Case "发卡原因"
        If zl_SelectAndNotAddItem(Me, txtEdit(mtxtIdx.idx_txt发卡原因), "", "常用发卡原因", "常用发卡原因选择", True, True) = False Then
            Exit Sub
        End If
    Case Else
    End Select
End Sub

Private Sub dtp卡有效日期_Change()
    If mCardEditType <> gEd_发卡 Then Exit Sub

    '修改的话,需要同步更新网格行的数据才行
    With vsGrid
        If Split(.TextMatrix(.Row, .ColIndex("卡号")) & "～", "～")(0) <> txtEdit(mtxtIdx.idx_txt卡号).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("卡有效期")) = Format(dtp卡有效日期.value, "yyyy-mm-dd HH:MM")
    End With
End Sub

Private Sub dtp卡有效日期_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode <> vbKeyReturn Then Exit Sub
     zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtp卡有效日期_Validate(Cancel As Boolean)
    If mCardEditType <> gEd_发卡 Then Exit Sub

    '修改的话,需要同步更新网格行的数据才行
    '84792:李南春,2015/7/17,批量卡号判断不正确
    With vsGrid
        If Split(.TextMatrix(.Row, .ColIndex("卡号")) & "～", "～")(0) <> txtEdit(mtxtIdx.idx_txt卡号).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("卡有效期")) = Format(dtp卡有效日期.value, "yyyy-mm-dd HH:MM")
    End With
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnLoad Then Unload Me: Exit Sub
    If CheckDepented = False Then Unload Me: Exit Sub
    If LoadDatatoCard = False Then
        mlng消费卡ID = 0
        If mCardEditType <> gEd_充值 Then
            '可以再次冲值
            Unload Me: Exit Sub
        End If
    End If
    '窗体样式调整必须在窗体Load完成之后
    Call SetWindowsSize
    Call SetEditProperty
    If mCardEditType = gEd_充值 Then
         zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt本次充值)
    End If
    mblnChange = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mblnUnLoad = False
    Call CreateObjectKeyboard
    Call InitModulePara
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call InitPanel
    Call zlDefCommandBars '初始菜单及工具栏
    Call InitVsGrid
    
    Call Load支付方式
    Me.Caption = Switch(mCardEditType = gEd_查询, "消费卡信息查询", mCardEditType = gEd_退卡, "消费卡退卡", mCardEditType = gEd_修改, "消费卡信息修改", mCardEditType = gEd_充值, "消费卡允值管理", mCardEditType = gEd_发卡, "消费卡发放", mCardEditType = gEd_回收, "消费卡回收管理", mCardEditType = gEd_回退, "消费卡回退", mCardEditType = gEd_取消回收, "消费卡取消回收", True, "消费卡发放")
    mstrTitle = Me.Caption
    RaisEffect picCardInfor, -1
    '问题65902,刘尔旋:调整消费卡修改密码的方式
    If mCardEditType = gEd_修改 Then
        txtEdit(1).Enabled = False
        txtEdit(2).Enabled = False
        txtEdit(2).Visible = False
        lblEdit(3).Visible = False
    Else
        txtEdit(1).Enabled = True
        txtEdit(2).Enabled = True
        txtEdit(2).Visible = True
        lblEdit(3).Visible = True
    End If
    If mCardEditType = gEd_回收 Or mCardEditType = gEd_发卡 Then
        RestoreWinState Me, App.ProductName, mstrTitle
    End If
End Sub

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2009-12-10 11:39:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsGrid
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("卡号")) = "1|0"
        .ColData(.ColIndex("标志")) = "-1|1"
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("当前余额")) = "1|0"
        .Clear 1
        .Rows = 2
    End With
End Sub

Private Function zlMoveCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:移出当前卡片信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-18 09:56:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCurRow  As Long
    Err = 0: On Error GoTo Errhand:
    If mCardEditType <> gEd_发卡 And mCardEditType <> gEd_回收 Then
        Exit Function
    End If

    With vsGrid
        If .Rows < 2 Then Exit Function
        If .Rows <= 2 And .Row = 1 Then
            .Clear 1
            .Cell(flexcpData, 1, 0, .Rows - 1, .Cols - 1) = ""
            Call FromGridToCtlData
            Call SetDefaultValue
            Call CheckOtherCard
            zlMoveCard = True
            Exit Function
        End If
        lngCurRow = .Row
        .RemoveItem lngCurRow
        If lngCurRow < .Rows - 1 Then
            lngCurRow = lngCurRow + 1
        Else
            lngCurRow = .Rows - 1
        End If
        If lngCurRow < 1 Then lngCurRow = 1
        If lngCurRow > 1 Then .Row = lngCurRow
    End With
    Call Show发卡数量
    Call CheckOtherCard
    Call Calc实收合计
    zlMoveCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume

End Function

Private Function zlAppColumnData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:应用于列数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-18 09:56:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow  As Long, strTemp As String, strTempData As String, lngCurCol As Long
    Err = 0: On Error GoTo Errhand:
    If mCardEditType <> gEd_发卡 Then
        Exit Function
    End If

    With vsGrid
        If .Rows < 2 Then Exit Function
        If .Rows <= 2 And .Row = 1 Then
            Exit Function
        End If
        lngCurCol = .Col
        If .ColIndex("卡号") = lngCurCol Then Exit Function
        strTemp = .TextMatrix(.Row, lngCurCol)
        strTempData = .Cell(flexcpData, .Row, lngCurCol)
        For lngRow = 1 To .Rows - 1
            If lngRow <> .Row Then
                .TextMatrix(lngRow, lngCurCol) = strTemp
                .Cell(flexcpData, lngRow, lngCurCol) = strTempData
            End If
        Next
    End With
    zlAppColumnData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function zlAppAllCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:应用于列数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-18 09:56:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow  As Long, strTemp As String, strTempData As String, lngCol As Long
    Err = 0: On Error GoTo Errhand:
    If mCardEditType <> gEd_发卡 Then
        Exit Function
    End If

    With vsGrid
        If .Rows < 2 Then Exit Function
        If .Rows <= 2 And .Row = 1 Then
            Exit Function
        End If
        For lngRow = 1 To .Rows - 1
            If lngRow <> .Row Then
                For lngCol = 0 To .Cols - 1
                    Select Case lngCol
                    Case .ColIndex("卡号"), .ColIndex("标志"), .ColIndex("发卡数量")

                    Case Else
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(.Row, lngCol)
                        .Cell(flexcpData, lngRow, lngCol) = .Cell(flexcpData, .Row, lngCol)
                    End Select
                Next
            End If
        Next
    End With
    zlAppAllCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function



'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '------------------------------------
    Select Case Control.id
        Case conMenu_File_Exit: Unload Me
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case mconMenu_Edit_Affirm '确定
            Call SaveData
        Case conMenu_File_Print '打印
        Case conMenu_Edit_MoveCard   '移出卡片
             If zlMoveCard = False Then Exit Sub
         Case conMenu_Apply_AllCard     '应用于所有
            If zlAppAllCard = False Then Exit Sub
         Case conMenu_Apply_AllColumn     '应用于此列
            If zlAppColumnData = False Then Exit Sub
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub SaveData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '编制:刘兴洪
    '日期:2009-12-11 14:00:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng发卡序号 As Long, lngID As Long, rsTemp As ADODB.Recordset, txtTemp As TextBox
    If isValied = False Then Exit Sub
    
    '65048:刘尔旋,2013-10-30,保存时强制更新录入信息,避免遗留更改
    For Each txtTemp In txtEdit
        If Not (txtTemp.Index = 12 Or txtTemp.Index = 13 Or txtTemp.Index = 14 Or txtTemp.Index = 15 Or txtTemp.Index = 16 Or txtTemp.Index = 17) Then _
        Call txtEdit_Validate(txtTemp.Index, False)
    Next
    
    If mCardEditType = gEd_发卡 Then
        lng发卡序号 = zlDatabase.GetNextId("消费卡目录")
        If SavePayCard(lng发卡序号) = False Then Exit Sub
        If mTy_MoudlePara.bln缴款单打印 Then
            '打印缴款单
            '可能存在未缴款的情况
            If InStr(1, mstrPrivs, ";消费卡收费收据;") <> 0 Then
                'If Val(txtEdit(mtxtIdx.idx_txt实收合计)) <> 0 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "付款序号=" & lng发卡序号, "缴款=" & Val(txtEdit(mtxtIdx.idx_txt本次缴款).Text), "找补=" & Val(txtEdit(mtxtIdx.idx_txt找补).Tag), "充值ID=0", "ReportFormat=1", 2)
            End If
        End If
        Call ClearCtlData(True)
        Call zl_CtlSetFocus(Me.txtEdit(mtxtIdx.idx_txt卡号))
        mblnChange = False: mintSucces = mintSucces + 1
        Exit Sub
    End If

    If mCardEditType = gEd_修改 Then
        '修改处理
        If SaveModifyCard = False Then Exit Sub
        '打印缴款单
        '可能存在未缴款的情况
        If Val(txtEdit(mtxtIdx.idx_txt实收合计)) <> 0 And mCardInfor.bln允许充值 And mTy_MoudlePara.bln缴款单打印 Then
            '可能存在未缴款的情况
            If InStr(1, mstrPrivs, ";消费卡收费收据;") <> 0 Then
                gstrSQL = "Select 发卡序号 From 消费卡目录 where id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng消费卡ID)
                If rsTemp.EOF = False Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "付款序号=" & Val(Nvl(rsTemp!发卡序号)), "缴款=" & Val(txtEdit(mtxtIdx.idx_txt本次缴款).Text), "找补=" & Val(txtEdit(mtxtIdx.idx_txt找补).Text), "充值ID=0", "ReportFormat=1", 2)
                End If
            End If

         '   Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "消费卡ID=" & mlng消费卡ID, 2)
        End If
        mblnChange = False:: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If
    If mCardEditType = gEd_回收 Then
        '回收处理
        If SaveCallBack = False Then Exit Sub
        mblnChange = False:: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If

    If mCardEditType = gEd_取消回收 Then
        '回收处理
        If SaveCallBack(True) = False Then Exit Sub
        mblnChange = False:: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If

    If mCardEditType = gEd_退卡 Then
       '退卡处理操作
       If SaveBackCard(False) = False Then Exit Sub
        mblnChange = False: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If

    If mCardEditType = gEd_取消退卡 Then
       '退卡处理操作
       If SaveBackCard(True) = False Then Exit Sub
        mblnChange = False:: mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If
    If mCardEditType = gEd_充值 Then
        If mlng消费卡ID = 0 Then
            ShowMsgbox "注意:" & vbCrLf & "    不存指定的消费卡,不能充值!"
            Exit Sub
        End If

        If SaveInFull(lngID) = False Then Exit Sub
        '打印缴款单
        '可能存在未缴款的情况
        If Val(txtEdit(mtxtIdx.idx_txt实收合计)) <> 0 And mCardInfor.bln允许充值 And mTy_MoudlePara.bln缴款单打印 Then
            If InStr(1, mstrPrivs, ";消费卡收费收据;") <> 0 Then
                'If Val(txtEdit(mtxtIdx.idx_txt实收合计)) <> 0 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "充值ID=" & lngID, "缴款=" & Val(txtEdit(mtxtIdx.idx_txt本次缴款).Text), "找补=" & Val(txtEdit(mtxtIdx.idx_txt找补).Text), "付款序号=0", "ReportFormat=2", 2)
            End If
        End If
        mblnChange = False:: mintSucces = mintSucces + 1
        If IIf(Val(zlDatabase.GetPara("连续充值", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
            Call ClearCtlData(True): mlng消费卡ID = 0
            Call zl_CtlSetFocus(Me.txtEdit(mtxtIdx.idx_txt卡号))
            Call Set可否充值:             Calc余额: Calc实收合计
            Call SetEditProperty
        Else
            Unload Me: Exit Sub
        End If
    End If
End Sub

Private Function SaveInFull(ByRef lngID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存充值处理
    '出参:lngID-返回本次的充值的ID
    '返回:充值成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2009-12-14 10:51:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng序号 As Long
    Err = 0: On Error GoTo Errhand:
    
    '61905:刘尔旋,2013-10-29,支票充值填写信息超长检查
    If zlCommFun.ActualLen(Trim(txtEdit(mtxtIdx.idx_txt开户行).Text)) > 50 Then
        ShowMsgbox "注意:" & vbCrLf & "    开户行长度超长,请重新填写!"
        txtEdit(mtxtIdx.idx_txt开户行).SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtEdit(mtxtIdx.idx_txt帐号).Text)) > 20 Then
        ShowMsgbox "注意:" & vbCrLf & "    帐号长度超长,请重新填写!"
        txtEdit(mtxtIdx.idx_txt帐号).SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtEdit(mtxtIdx.idx_txt结算号码).Text)) > 30 Then
        ShowMsgbox "注意:" & vbCrLf & "    结算号码长度超长,请重新填写!"
        txtEdit(mtxtIdx.idx_txt结算号码).SetFocus
        Exit Function
    End If
    '78084:李南春,2014/9/18,充值金额判断
    If Val(txtEdit(mtxtIdx.idx_txt本次充值).Text) = 0 Then
        ShowMsgbox "注意:" & vbCrLf & "    本次充值的金额为零,请重新填写!"
        If txtEdit(mtxtIdx.idx_txt本次充值).Visible And txtEdit(mtxtIdx.idx_txt本次充值).Enabled Then txtEdit(mtxtIdx.idx_txt本次充值).SetFocus
        Exit Function
    End If
    
    lngID = zlDatabase.GetNextId("消费卡充值记录")
    lng序号 = GetMax充值序号(mlng消费卡ID)
    'Zl_消费卡充值记录_Insert
    gstrSQL = "Zl_消费卡充值记录_Insert("
    '  Id_In         In 消费卡充值记录.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  消费卡id_In   In 消费卡充值记录.消费卡id%Type,
    gstrSQL = gstrSQL & "" & mlng消费卡ID & ","
    '  序号_In       In 消费卡充值记录.序号%Type,
    gstrSQL = gstrSQL & "" & lng序号 & ","
    '  充值金额_In   In 消费卡充值记录.充值金额%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt本次充值).Text), 4) & ","
    '  充值折扣_In   In 消费卡充值记录.充值折扣%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt充值扣率).Text), 4) & ","
    '  缴款金额_In   In 消费卡充值记录.缴款金额%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text), 4) & ","
    '  充值时间_In   In 消费卡充值记录.充值时间%Type,
    gstrSQL = gstrSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  操作员姓名_In In 消费卡充值记录.操作员姓名%Type,
    gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
    '  缴款人_In     In 消费卡充值记录.缴款人%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt缴款人).Text) & "',"
    '  备注_In       In 消费卡充值记录.备注%Type
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt充值备注).Text) & "',"
     '  结算方式_IN  in 消费卡充值记录.结算方式%Type
    gstrSQL = gstrSQL & IIf(chk是否充值.value = 1, "'" & cboStyle.Text & "'", "NULL") & ","
    '  开户行_IN  in 消费卡充值记录.单位开户行%Type
    gstrSQL = gstrSQL & IIf(chk是否充值.value = 1, "'" & Trim(txtEdit(mtxtIdx.idx_txt开户行).Text) & "'", "NULL") & ","
    
   '  帐号_IN  in 消费卡充值记录.单位帐号%Type
    gstrSQL = gstrSQL & IIf(chk是否充值.value = 1, "'" & Trim(txtEdit(mtxtIdx.idx_txt帐号).Text) & "'", "NULL") & ","
    
   '  结算号码_IN  in 消费卡充值记录.单位结算号码%Type
    gstrSQL = gstrSQL & IIf(chk是否充值.value = 1, "'" & Trim(txtEdit(mtxtIdx.idx_txt结算号码).Text) & "'", "NULL") & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveInFull = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function GetMax充值序号(ByVal lng消费卡ID As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取最大的充值序号
    '编制:刘兴洪
    '日期:2009-12-14 11:06:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select nvl( Max(序号),0)+1 as 充值序号 From 消费卡充值记录 where 消费卡ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取最大的充值序号", lng消费卡ID)
    GetMax充值序号 = Val(Nvl(rsTemp!充值序号))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveCallBack(Optional blnCancelCallBack As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:回收处理
    '入参:blnCancelCallBack-取消回收
    '编制:刘兴洪
    '日期:2009-12-14 09:45:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As New Collection, lngRow As Long, strIDs As String, blnHaveData As Boolean
    
    Err = 0: On Error GoTo Errhand:
    If blnCancelCallBack Then
        'Zl_消费卡目录_Callback
        gstrSQL = "Zl_消费卡目录_Callback("
        '  Ids_In       IN varchar2,
        gstrSQL = gstrSQL & "" & mlng消费卡ID & ","
        '  回收人_In   In 消费卡目录.回收人%Type,
        gstrSQL = gstrSQL & IIf(blnCancelCallBack = True, "NULL", "'" & UserInfo.姓名 & "'") & ","
        '  回收时间_In In 消费卡目录.回收时间%Type
        gstrSQL = gstrSQL & "NULL)"
        AddArray cllPro, gstrSQL
    Else
        '可能存在批量回收操作,因此要传入ID的集合
        strIDs = ""
        blnHaveData = False
        With vsGrid
            For lngRow = 1 To .Rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("ID"))) <> 0 Then
                    If zlCommFun.ActualLen(strIDs) > 4000 Then
                        strIDs = Mid(2, strIDs)
                        'Zl_消费卡目录_Callback
                        gstrSQL = "Zl_消费卡目录_Callback("
                        '  Ids_In     IN varchar2,
                        gstrSQL = gstrSQL & "'" & strIDs & "',"
                        '  回收人_In   In 消费卡目录.回收人%Type,
                        gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
                        '  回收时间_In In 消费卡目录.回收时间%Type
                        gstrSQL = gstrSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
                        AddArray cllPro, gstrSQL
                        strIDs = ""
                    End If
                    strIDs = strIDs & "," & Val(.TextMatrix(lngRow, .ColIndex("ID")))
                    blnHaveData = True
                End If
            Next
            If strIDs <> "" Then
                strIDs = Mid(strIDs, 2)
                'Zl_消费卡目录_Callback
                gstrSQL = "Zl_消费卡目录_Callback("
                '  Ids_In       IN varchar2,
                gstrSQL = gstrSQL & "'" & strIDs & "',"
                '  回收人_In   In 消费卡目录.回收人%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
                '  回收时间_In In 消费卡目录.回收时间%Type
                gstrSQL = gstrSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))"
                AddArray cllPro, gstrSQL
            End If
            If blnHaveData = False Then
                ShowMsgbox "注意:" & vbCrLf & "    你没有刷要回收的消费卡,请检查!"
                zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt卡号), True
                Exit Function
            End If
        End With
    End If
    
Err = 0: On Error GoTo Errhand:
    ExecuteProcedureArrAy cllPro, Me.Caption
    SaveCallBack = True
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Function SaveBackCard(Optional blnCancelBackCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:回收处理
    '入参:blnCancelBackCard-取消退卡
    '编制:刘兴洪
    '日期:2009-12-14 09:45:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Err = 0: On Error GoTo Errhand:
    'Zl_消费卡目录_Backcard
    gstrSQL = "Zl_消费卡目录_Backcard("
    '  Id_In       In 消费卡目录.ID%Type,
    gstrSQL = gstrSQL & "" & mlng消费卡ID & ","
    '  回退人_In   In 消费卡目录.回收人%Type,
    gstrSQL = gstrSQL & IIf(blnCancelBackCard = True, "NULL", "'" & UserInfo.姓名 & "'") & ","
    '  回退时间_In In 消费卡目录.回收时间%Type
    gstrSQL = gstrSQL & IIf(blnCancelBackCard = True, "NULL", "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')") & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveBackCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function


Private Function Get限制类别() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取限制类别
    '编制:刘兴洪
    '日期:2009-12-11 15:36:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String, i As Long
    With lvwType
         For i = 1 To .ListItems.count
            If .ListItems.Item(i).Checked Then
                strType = strType & "," & Mid(.ListItems(i).Key, 2)
            End If
         Next
         If strType <> "" Then strType = Mid(strType, 2)
    End With
    Get限制类别 = strType
End Function

Private Function SaveModifyCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存卡片修改信息
    '返回:修改成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2009-12-11 14:28:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    
    'Zl_消费卡目录_Update
    gstrSQL = "Zl_消费卡目录_Update("
    '  Id_In         In 消费卡目录.ID%Type,
    gstrSQL = gstrSQL & "" & mlng消费卡ID & ","
    '  卡号_In       In 消费卡目录.卡号%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt卡号).Text) & "',"
    '  卡类型_In     In 消费卡目录.卡类型%Type,
    gstrSQL = gstrSQL & "" & zl_FromComboxGetData(cbo卡类型) & ","
    '  限制类别_In   In 消费卡目录.限制类别%Type,
    gstrSQL = gstrSQL & "'" & Get限制类别 & "',"
    '  可否充值_In   In 消费卡目录.可否充值%Type,
    gstrSQL = gstrSQL & "" & IIf(chk是否充值.value = 1, 1, 0) & ","
    '  有效期_In     In 消费卡目录.有效期%Type,
        If IsNull(dtp卡有效日期.value) Then
            gstrSQL = gstrSQL & "NULL,"
        Else
            gstrSQL = gstrSQL & "to_date('" & Format(dtp卡有效日期.value, "yyyy-mm-dd HH:MM") & "','yyyy-mm-dd hh24:mi'),"
        End If
    '  发卡原因_In   In 消费卡目录.发卡原因%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt发卡原因).Text) & "',"
    '  发卡人_In     In 消费卡目录.发卡人%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt发卡人).Text) & "',"
    '  领卡人_In     In 消费卡目录.领卡人%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt领卡人).Text) & "',"
    '  领卡部门id_In In 消费卡目录.领卡部门id%Type,
    gstrSQL = gstrSQL & "" & IIf(Val(txtEdit(mtxtIdx.idx_txt领卡部门).Tag) = 0, "NULL", Val(txtEdit(mtxtIdx.idx_txt领卡部门).Tag)) & ","
    '  发卡时间_In   In 消费卡目录.发卡时间%Type,
    gstrSQL = gstrSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  备注_In       In 消费卡目录.备注%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt备注).Text) & "',"
    '  结算方式_In   In 消费卡目录.结算方式%Type,
    gstrSQL = gstrSQL & "'" & Trim(cboStyle.Text) & "',"
    '  卡面金额_In   In 消费卡目录.卡面金额%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt卡面额).Text), 4) & ","
    '  销售金额_In   In 消费卡目录.销售金额%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt实际销售额).Text), 4) & ","
    '  充值折扣率_In In 消费卡目录.充值折扣率%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt充值扣率).Text), 4) & ","
    '  余额_In       In 消费卡目录.余额%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt当前余额).Text), 4) & ","
    '  充值金额_In   In 消费卡充值记录.充值金额%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt本次充值).Text), 4) & ","
    '  缴款金额_In   In 消费卡充值记录.缴款金额%Type,
    gstrSQL = gstrSQL & "" & Round(Val(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text), 4) & ","
    '  充值说明_In   In 消费卡充值记录.备注%Type
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mtxtIdx.idx_txt充值备注).Text) & "',"
    ' n_更新金额 Number:=0 --n_更新金额:1-需要更新面值及充值额及余额:0-只更新附加信息
    gstrSQL = gstrSQL & IIf(mCardInfor.bln允许充值, "1", "0") & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    SaveModifyCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case mconMenu_Edit_Affirm '确定
        Control.Enabled = Not (mCardEditType = gEd_查询)
        Control.Visible = Not (mCardEditType = gEd_查询)
    
    Case conMenu_File_Print '打印
    Case conMenu_Edit_MoveCard   '移出卡片
        If mCardEditType <> gEd_发卡 And mCardEditType <> gEd_回收 Then
            Control.Visible = False: Control.Enabled = False: Exit Sub
        End If
        Control.Enabled = vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("卡号")) <> "" Or vsGrid.Rows > 2
    Case conMenu_Apply_AllCard     '应用于所有
        If mCardEditType <> gEd_发卡 Then
            Control.Visible = False: Control.Enabled = False: Exit Sub
        End If
        With vsGrid
            Control.Caption = "应用于其他卡片"
            Control.ToolTipText = "将卡号为“" & .TextMatrix(.Row, .ColIndex("卡号")) & "”的信息 应用于其他卡信息"
            Control.Enabled = mblnHaveOtherCard '只有存在其他单据时，才会出现此列信息
        End With
    Case conMenu_Apply_AllColumn     '应用于此列
        If mCardEditType <> gEd_发卡 Then
            Control.Visible = False: Control.Enabled = False: Exit Sub
        End If
        With vsGrid
            Select Case .Col
            Case .ColIndex("卡号"), .ColIndex("发卡数量")
                 Control.Enabled = False: Exit Sub
            Case Else
                Control.Caption = "应用于〖" & Trim(.TextMatrix(0, .Col)) & "〗列"
                Control.ToolTipText = "将“" & .TextMatrix(.Row, .Col) & "” 应用于〖" & Trim(.TextMatrix(0, .Col)) & "〗列的其他卡片信息"
            End Select
            Control.Enabled = mblnHaveOtherCard '只有存在其他单据时，才会出现此列信息
        End With

    End Select
End Sub

Private Sub CheckOtherCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在其他卡片信息
    '编制:刘兴洪
    '日期:2009-12-18 09:47:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    mblnHaveOtherCard = False
    With vsGrid
        For i = 1 To .Rows - 1
            If i <> .Row And .TextMatrix(i, .ColIndex("卡号")) <> "" Then
                mblnHaveOtherCard = True
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
'    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case mPaneID.Pane_Cards     '搜索条件窗口
        Item.Handle = picList
    Case mPaneID.Pane_CardInfor    '详细卡信息
        Item.Handle = picCard.hWnd
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mCardEditType = gEd_回收 Or mCardEditType = gEd_发卡 Then
        SaveWinState Me, App.ProductName, mstrTitle
    End If
'    If mCardEditType = gEd_回收 Then
'        zlSaveDockPanceToReg Me, dkpMan, "区域-回收"
'    ElseIf mCardEditType = gEd_发卡 Then
'        zlSaveDockPanceToReg Me, dkpMan, "区域-发卡"
'    End If
    If mblnChange = False Then Exit Sub
    If mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改 Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub lvwType_Click()
    mblnChange = True
End Sub

'65048:刘尔旋,2013-10-30,发卡时限制类别未保存的问题
Private Sub lvwType_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If mCardEditType <> gEd_发卡 Then Exit Sub
    '84792:李南春,2015/7/17,批量卡号判断不正确
    With vsGrid
        If Split(.TextMatrix(.Row, .ColIndex("卡号")) & "～", "～")(0) <> txtEdit(mtxtIdx.idx_txt卡号).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("限制类别")) = Get限制类别
    End With
End Sub

Private Sub lvwType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改 Then
        If fra缴款.Visible = False Then
            If MsgBox("你是否真要进行相关的" & IIf(mCardEditType = gEd_修改, "修改", "发卡") & "操作吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call SaveData
            End If
            Exit Sub
        End If
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvwType_Validate(Cancel As Boolean)
    If mCardEditType <> gEd_发卡 Then Exit Sub
    '修改的话,需要同步更新网格行的数据才行
    '84792:李南春,2015/7/17,批量卡号判断不正确
    With vsGrid
        If Split(.TextMatrix(.Row, .ColIndex("卡号")) & "～", "～")(0) <> txtEdit(mtxtIdx.idx_txt卡号).Text Then Exit Sub
        .TextMatrix(.Row, .ColIndex("限制类别")) = Get限制类别
    End With
End Sub

Private Sub mobjBrushCard_zlBrushCarding(ByVal strCardNo As String)
    '刷卡操作
    If strCardNo = "" Then Exit Sub
    Call zlBrusCard(strCardNo, True)
        mobjBrushCard.zlSetAutoBrush False
End Sub

Private Sub picCard_Resize()
    Err = 0: On Error Resume Next
    Dim sngTop As Single, sngLeft As Long
    With picCard
        sngLeft = (.ScaleLeft + .ScaleWidth - picCardInfor.Width) \ 2
        sngLeft = IIf(sngLeft < 0, 0, sngLeft)
        sngTop = (.ScaleTop + .ScaleHeight - picCardInfor.Height) \ 2
        sngTop = IIf(sngTop < 0, 0, sngTop)
        picCardInfor.Move sngLeft, sngTop
    End With
End Sub

Private Sub picList_Resize()
    Dim sngWidth As Single
    Err = 0: On Error Resume Next
    With picList
        If .ScaleWidth - lbl至.Width - txtEdit(mtxtIdx.idx_txt结束卡号).Width - txtEdit(mtxtIdx.idx_txt开始卡号).Left - txtEdit(mtxtIdx.idx_txt开始卡号).Width - 120 > 0 Then
            txtEdit(mtxtIdx.idx_txt结束卡号).Top = txtEdit(mtxtIdx.idx_txt开始卡号).Top
            lbl至.Top = lbl开始卡号.Top
            lbl至.Left = txtEdit(mtxtIdx.idx_txt开始卡号).Left + txtEdit(mtxtIdx.idx_txt开始卡号).Width + 100
            txtEdit(mtxtIdx.idx_txt结束卡号).Left = lbl至.Left + lbl至.Width + 20
        Else
            txtEdit(mtxtIdx.idx_txt结束卡号).Left = txtEdit(mtxtIdx.idx_txt开始卡号).Left
            lbl至.Left = lbl开始卡号.Left

            txtEdit(mtxtIdx.idx_txt结束卡号).Top = txtEdit(mtxtIdx.idx_txt开始卡号).Top + txtEdit(mtxtIdx.idx_txt开始卡号).Height + 50
            lbl至.Top = txtEdit(mtxtIdx.idx_txt结束卡号).Top + (txtEdit(mtxtIdx.idx_txt结束卡号).Height - lbl至.Height) \ 2
        End If
    
'        If .Width < 4350 Then
'            lbl至.Caption = "结束单号": lbl至.FontBold = False
'            txtEdit(mtxtIdx.idx_txt结束卡号).Text
'            lbl至.Top = txtEdit(mtxtIdx.idx_txt开始卡号).Top + txtEdit(mtxtIdx.idx_txt开始卡号).Height + 50
'
'        Else
'            sngWidth = .ScaleWidth - (lbl卡号.Left + lbl卡号.Width) - (lbl至.Width * 3)
'            sngWidth = IIf(sngWidth < 3360, 3360, sngWidth)
'            sngWidth = sngWidth \ 2
'            txtEdit(mtxtIdx.idx_txt开始卡号).Width = sngWidth
'            lbl至.Left = txtEdit(mtxtIdx.idx_txt开始卡号).Left + txtEdit(mtxtIdx.idx_txt开始卡号).Width + lbl至.Width
'            txtEdit(mtxtIdx.idx_txt结束卡号).Left = lbl至.Left + (lbl至.Width * 2)
'            txtEdit(mtxtIdx.idx_txt结束卡号).Width = sngWidth
'        End If
        vsGrid.Top = IIf(txtEdit(mtxtIdx.idx_txt结束卡号).Visible = False, 0, txtEdit(mtxtIdx.idx_txt结束卡号).Top + txtEdit(mtxtIdx.idx_txt结束卡号).Height) + 50
        vsGrid.Left = .ScaleLeft
        vsGrid.Width = .ScaleWidth
        vsGrid.Height = .ScaleHeight - vsGrid.Top
    End With
End Sub
Private Function zlCardNoRange(ByVal strCardNoRange As String, ByRef strCardNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据传入的卡号范围，分解成相关的卡号
    '入参:strCardNoRange-卡号范围
    '出参:strCardNos-返回卡号数(用逗号分离)
    '返回:
    '编制:刘兴洪
    '日期:2009-12-10 16:30:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant
    Dim strCardStartNO As String, strCardEndNO As String, strCurNo As String
    Dim lngCount As Long


    varData = Split(strCardNoRange & "～", "～")
    strCardStartNO = varData(0): strCardEndNO = varData(1)
    If strCardEndNO = "" Then strCardNos = strCardStartNO: GoTo GoExit:

'    If strSartCardText = mTy_MoudlePara.str卡号前缀 Then
'        strCardStartNO = Mid(strCardStartNO, Len(mTy_MoudlePara.str卡号前缀) + 1)
'    End If
'    If strEndCardText = mTy_MoudlePara.str卡号前缀 Then
'        strCardEndNO = Mid(strCardEndNO, Len(mTy_MoudlePara.str卡号前缀) + 1)
'    End If
    If strCardStartNO > strCardEndNO Then
        Exit Function
    End If

    strCurNo = strCardStartNO
    strCardNos = strCardStartNO         'mTy_MoudlePara.str卡号前缀 &
    lngCount = 0
    Do While True
        If strCurNo >= strCardEndNO Then
            Exit Do
        End If
        strCurNo = zlCommFun.IncStr(strCurNo)
        strCardNos = strCardNos & "," & strCurNo  'mTy_MoudlePara.str卡号前缀 &
        lngCount = lngCount + 1
        If lngCount > 1000 Then Exit Do
Loop
    If lngCount > 1000 Then
        MsgBox "注意卡号范围不能大于1000，不能继续操作!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
GoExit:
    zlCardNoRange = True: Exit Function
End Function

Private Function zl消费卡InsertSQL(ByVal lng发卡序号 As Long, ByVal strCardNo As String, ByVal str发卡时间 As String, ByVal lngRow As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取插入SQL语句
    '入参:lng发卡序号-主要是标明一批发卡时的发卡序号,以便打印
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-11 09:21:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        ' Zl_消费卡目录_Insert
        gstrSQL = "Zl_消费卡目录_Insert("
        '   接口编号_IN      IN 消费卡目录.接口编号%Type,

        gstrSQL = gstrSQL & "'" & mlng接口编号 & "',"
        '  卡号_In       In Varchar2, --用,分离
        gstrSQL = gstrSQL & "'" & strCardNo & "',"
        '  卡类型_In     In 消费卡目录.卡类型%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("卡类型"))) & "',"
        '  密码_In       In 消费卡目录.密码%Type,
        gstrSQL = gstrSQL & "'" & zlCommFun.zlStringEncode(Trim(.Cell(flexcpData, lngRow, .ColIndex("密码")))) & "',"
        '  限制类别_In   In 消费卡目录.限制类别%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("限制类别"))) & "',"
        '  可否充值_In   In 消费卡目录.可否充值%Type,
        gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("是否充值"))) & ","
        '  有效期_In     In 消费卡目录.有效期%Type,
        If Trim(.TextMatrix(lngRow, .ColIndex("卡有效期"))) = "" Then
            gstrSQL = gstrSQL & "NULL,"
        Else
            gstrSQL = gstrSQL & "to_date('" & Trim(.TextMatrix(lngRow, .ColIndex("卡有效期"))) & "','yyyy-mm-dd hh24:mi'),"
        End If
        '  发卡原因_In   In 消费卡目录.发卡原因%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("发卡原因"))) & "',"
        '  发卡人_In     In 消费卡目录.发卡人%Type,
        gstrSQL = gstrSQL & "'" & IIf(Trim(.TextMatrix(lngRow, .ColIndex("发卡人"))) = "", UserInfo.姓名, Trim(.TextMatrix(lngRow, .ColIndex("发卡人")))) & "',"
        '  领卡人_In     In 消费卡目录.领卡人%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("领卡人"))) & "',"
        '  领卡部门id_In In 消费卡目录.领卡部门id%Type,
        gstrSQL = gstrSQL & "" & IIf(Val(.Cell(flexcpData, lngRow, .ColIndex("领卡部门"))) = 0, "NULL", Val(.Cell(flexcpData, lngRow, .ColIndex("领卡部门")))) & ","
        '  发卡时间_In   In 消费卡目录.发卡时间%Type,
        gstrSQL = gstrSQL & "to_date('" & str发卡时间 & "','yyyy-mm-dd hh24:mi:ss'),"
        '  备注_In       In 消费卡目录.备注%Type,
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("备注"))) & "',"
        '  卡面金额_In   In 消费卡目录.卡面金额%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("卡面额"))), 4) & ","
        '  销售金额_In   In 消费卡目录.销售金额%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("实际销售"))), 4) & ","
        '  充值折扣率_In In 消费卡目录.充值折扣率%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("充值扣率"))) * IIf(chk是否充值.value = 1, 1, 0), 4) & ","
        '  余额_In       In 消费卡目录.余额%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("当前余额"))), 4) & ","
        '    发卡序号_IN   IN 消费卡目录.发卡序号%type,
        gstrSQL = gstrSQL & "" & lng发卡序号 & ","
        '  充值金额_In   In 消费卡充值记录.充值金额%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("本次充值"))) * IIf(chk是否充值.value = 1, 1, 0), 4) & ","
        '  缴款金额_In   In 消费卡充值记录.缴款金额%Type,
        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(lngRow, .ColIndex("实际充值缴款"))) * IIf(chk是否充值.value = 1, 1, 0), 4) & ","
        '  充值说明_In   In 消费卡充值记录.备注%Type
        gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(lngRow, .ColIndex("充值说明"))) & "',"
        ' 结算方式In   In 消费卡充值记录.结算方式%Type
        gstrSQL = gstrSQL & IIf(chk是否充值.value = 1, "'" & Trim(.TextMatrix(lngRow, .ColIndex("结算方式"))) & "'", "NULL") & ","
         
        '  开户行_IN  in 消费卡充值记录.单位开户行%Type
        gstrSQL = gstrSQL & IIf(chk是否充值.value = 1, "'" & Trim(.TextMatrix(lngRow, .ColIndex("开户行"))) & "'", "NULL") & ","

        '  帐号_IN  in 消费卡充值记录.单位帐号%Type
        gstrSQL = gstrSQL & IIf(chk是否充值.value = 1, "'" & Trim(.TextMatrix(lngRow, .ColIndex("帐号"))) & "'", "NULL") & ","

        '  结算号码_IN  in 消费卡充值记录.单位结算号码%Type
        gstrSQL = gstrSQL & IIf(chk是否充值.value = 1, "'" & Trim(.TextMatrix(lngRow, .ColIndex("结算号码"))) & "'", "NULL") & ")"
    End With
    AddArray cllPro, gstrSQL
    zl消费卡InsertSQL = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 检查卡号是否合法(ByVal strCardNos As String, Optional bln检查卡类型 As Boolean = False, Optional bln检查是否允值 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的卡号是否合法
    '入参:strCardNos-卡号集(每个卡号用逗号分离)
    '     bln检查卡类型-需要检查卡类型(对发卡有效)
    '     bln检查是否允值-检查是否允值(对发卡有效)
    '编制:刘兴洪
    '日期:2009-12-11 11:08:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strErrCardNo As String
    Dim lngRow As Long, strTable As String, strValue(0 To 10) As String, varData As Variant
    Dim lngCurRow As Long, i As Long, strCardNosTemp As String, j As Long
    Err = 0: On Error GoTo Errhand:
    If mCardEditType <> gEd_发卡 And mCardEditType <> gEd_回收 Then
        gstrSQL = "" & _
        "   Select ID,卡类型,可否充值 ,卡号,序号,(Select Max(序号) From 消费卡目录 B where A.卡号=B.卡号 and A.接口编号=b.接口编号) as 最大序号, " & _
        "       to_char(回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间, to_char(停用日期 ,'yyyy-mm-dd hh24:mi:ss') as 停用日期 " & _
        "   From 消费卡目录 A " & _
        "   Where A.id = [2]"
    Else
        If zlCommFun.ActualLen(strCardNos) > 1990 Then
            varData = Split(strCardNos, ",")
            strCardNosTemp = ""
            j = 3
            For i = 0 To UBound(varData)
                If j - 3 > 10 Then
                    strCardNosTemp = strCardNosTemp & "," & varData(i)
                Else
                    If zlCommFun.ActualLen(strCardNosTemp) > 1990 Then
                        strValue(j - 3) = Mid(strCardNosTemp, 2)
                        strTable = strTable & " UNION ALL Select Column_Value From Table(Cast(f_Str2list([" & j & "]) As Zltools.t_Strlist)) "
                        strCardNosTemp = ""
                        j = j + 1
                    End If
                    strCardNosTemp = strCardNosTemp & "," & varData(i)
                End If
            Next
            If j - 3 > 10 And strCardNosTemp <> "" Then
                strCardNosTemp = Mid(strCardNosTemp, 2)
                strCardNosTemp = "'" & Replace(strCardNosTemp, ",", "','") & "'"
                strTable = strTable & " UNION ALL   Select 卡号 From 消费卡目录 Where 卡号 in(" & strCardNosTemp & ")"
            ElseIf strCardNosTemp <> "" Then
                strValue(j - 3) = Mid(strCardNosTemp, 2)
                strTable = strTable & " UNION ALL Select Column_Value From Table(Cast(f_Str2list([" & j & "]) As Zltools.t_Strlist)) "
            End If
            If strTable <> "" Then strTable = Mid(strTable, 11)
        Else
            strTable = "Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) "
        End If

        gstrSQL = "" & _
        "   Select  /*+ RULE */ ID,卡类型,可否充值, 卡号,序号,序号 as 最大序号,to_char(回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间, to_char(停用日期 ,'yyyy-mm-dd hh24:mi:ss') as 停用日期 " & _
        "   From 消费卡目录 A, (" & strTable & ") B " & _
        "   Where A.卡号 = B.Column_Value And 序号 = (Select Max(序号) From 消费卡目录 B Where 卡号 = A.卡号 and 接口编号=A.接口编号 ) and a.接口编号=[3]  "
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strCardNos, mlng消费卡ID, mlng接口编号, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))

    strErrCardNo = ""
    Do While Not rsTemp.EOF
        '检查卡号是否合法
        Select Case mCardEditType
        Case gEd_发卡
            If Nvl(rsTemp!回收时间, "3000-01-01") >= "3000-01-01" Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡正在使用，不能再发卡,请检查!"
                Exit Function
            End If
            If Nvl(rsTemp!停用日期, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡已被停止使用，不能再发卡,请检查!"
                Exit Function
            End If
            lngCurRow = -1
            For lngRow = 1 To vsGrid.Rows - 1
                If "," & vsGrid.Cell(flexcpData, lngRow, vsGrid.ColIndex("卡号")) & "," Like "*," & strCardNos & "," Then
                    lngCurRow = lngRow: Exit For
                End If
            Next
            If lngCurRow = -1 Then
                If Val(Nvl(rsTemp!可否充值)) <> IIf(chk是否充值.value = 1, 1, 0) Then
                    ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡原来为" & IIf(Val(Nvl(rsTemp!可否充值)) = 1, "充值卡", "非充值卡") & "，而现在为" & IIf(chk是否充值.value = 1, "充值卡", "非充值卡") & ",请检查!"
                    Exit Function
                End If
            Else
                If Val(Nvl(rsTemp!可否充值)) <> IIf((vsGrid.Cell(flexcpData, lngCurRow, vsGrid.ColIndex("是否充值"))) = 1, 1, 0) Then
                    ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡原来为" & IIf(Val(Nvl(rsTemp!可否充值)) = 1, "充值卡", "非充值卡") & "，而现在为" & IIf(chk是否充值.value = 1, "充值卡", "非充值卡") & ",请检查!"
                    Exit Function
                End If
            End If
            If Trim(Nvl(rsTemp!卡类型)) <> Mid(cbo卡类型.Text, InStr(cbo卡类型.Text, ".") + 1) Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡原来的卡类型为" & Trim(Nvl(rsTemp!卡类型)) & "，而现在为" & Mid(cbo卡类型.Text, InStr(cbo卡类型.Text, ".") + 1) & ",请检查!"
                Exit Function
            End If
        Case gEd_修改
            If Val(Nvl(rsTemp!序号)) < Val(Nvl(rsTemp!最大序号)) Then
                ShowMsgbox "注意:" & vbCrLf & "   不能修改历史卡号信息(卡号为:" & Nvl(rsTemp!卡号) & ") ,请检查!"
                Exit Function
            End If
        Case gEd_充值
            If Val(Nvl(rsTemp!序号)) < Val(Nvl(rsTemp!最大序号)) Then
                ShowMsgbox "注意:" & vbCrLf & "   不能对历史卡号进行充值(卡号为:" & Nvl(rsTemp!卡号) & "),请检查!"
                Exit Function
            End If
            If Nvl(rsTemp!回收时间, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡已被回收，不能再充值,请先发卡后,再进行充值!"
                Exit Function
            End If
            If Nvl(rsTemp!停用日期, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡已被停止使用，不能再充值,请检查!"
                Exit Function
            End If
            If Val(Nvl(rsTemp!可否充值)) <> 1 Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡不是允值卡，不能再充值,请检查!"
                Exit Function
            End If
        Case gEd_回收
            If Val(Nvl(rsTemp!序号)) < Val(Nvl(rsTemp!最大序号)) Then
                ShowMsgbox "注意:" & vbCrLf & "   不能对回收历史卡号(卡号为:" & Nvl(rsTemp!卡号) & ") ,请检查!"
                Exit Function
            End If
            If Nvl(rsTemp!回收时间, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡已被回收，不能再回收,请检查!"
                Exit Function
            End If
            If Nvl(rsTemp!停用日期, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡已被停止使用，不能回收,请检查!"
                Exit Function
            End If
        Case gEd_退卡
            If Val(Nvl(rsTemp!序号)) < Val(Nvl(rsTemp!最大序号)) Then
                ShowMsgbox "注意:" & vbCrLf & "   不能对回退历史卡号(卡号为:" & Nvl(rsTemp!卡号) & ") ,请检查!"
                Exit Function
            End If
            If Nvl(rsTemp!回收时间, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡已被回收，不能再回退,请检查!"
                Exit Function
            End If
            If Nvl(rsTemp!停用日期, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "注意:" & vbCrLf & "   卡号为:" & Nvl(rsTemp!卡号) & " 的消费卡已被停止使用，不能回退,请检查!"
                Exit Function
            End If
        End Select
        rsTemp.MoveNext
    Loop

    If rsTemp.RecordCount = 0 Then
        If mCardEditType = gEd_修改 Then
            ShowMsgbox "消费卡可能已经被他人删除，不能修改卡信息,请检查!"
            Exit Function
        End If
        If mCardEditType = gEd_充值 Then
            ShowMsgbox "消费卡可能已经被他人删除，不能充值,请检查!"
            Exit Function
        End If
        If mCardEditType = gEd_回收 Then
            ShowMsgbox "消费卡可能已经被他人删除，不能回收,请检查!"
            Exit Function
        End If
        If mCardEditType = gEd_退卡 Then
            ShowMsgbox "消费卡可能已经被他人删除，不能退卡,请检查!"
            Exit Function
        End If
    End If

    检查卡号是否合法 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckVsGridData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查网格数据中的输入是否合法
    '编制:刘兴洪
    '日期:2009-12-11 12:06:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, lngRow As Long, i As Long, strCurNos As String
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        For lngRow = 1 To .Rows - 1
            strTemp = Trim(.Cell(flexcpData, lngRow, .ColIndex("卡号")))
            If strTemp <> "" Then
                If zlCommFun.ActualLen(strTemp) > 4000 Then
                    Do While True
                        i = InStr(3900, strTemp, ",")
                        If i > 0 Then
                            strCurNos = Mid(strTemp, 1, i - 1)
                            strTemp = Mid(strTemp, i + 1)
                            If strCurNos <> "" Then
                                If 检查卡号是否合法(strCurNos) = False Then
                                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                                    Exit Function
                                End If
                            End If
                        Else
                            If strTemp <> "" Then
                                If 检查卡号是否合法(strTemp) = False Then
                                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                                    Exit Function
                                End If
                            End If
                            Exit Do
                        End If
                    Loop
                Else
                    If 检查卡号是否合法(strTemp) = False Then
                        .Row = lngRow: zlCtlSetFocus vsGrid, True
                        Exit Function
                    End If
                End If
            End If

            If mCardEditType = gEd_发卡 Then
                If zlCommFun.StrIsValid(.TextMatrix(lngRow, .ColIndex("发卡原因")), 50, 0, "发卡原因") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlCommFun.StrIsValid(.TextMatrix(lngRow, .ColIndex("备注")), 100, 0, "备注") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlCommFun.StrIsValid(.TextMatrix(lngRow, .ColIndex("充值说明")), 100, 0, "充值说明") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If

                If Trim(.Cell(flexcpData, lngRow, .ColIndex("ID"))) <> Trim(.Cell(flexcpData, lngRow, .ColIndex("密码"))) Then
                    ShowMsgbox "在第" & lngRow & "中的密码输入不正确,请检查密码和确认密码是否正确!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If

                If zlCommFun.StrIsValid(Trim(.Cell(flexcpData, lngRow, .ColIndex("密码"))), 20, 0, "密码") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
'                If zlCommFun.StrIsValid(Trim(.Cell(flexcpData, lngRow, .ColIndex("确认密码"))), 20, 0, "确认密码") = False Then
'                    .Row = lngRow: zlCtlSetFocus vsGrid, True
'                    Exit Function
'                End If
                If zlCommFun.StrIsValid(Trim(.Cell(flexcpData, lngRow, .ColIndex("领卡人"))), 20, 0, "领卡人") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                '金额检查
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("卡面额"))), 16, True, False, 0, "卡面额") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("实际销售"))), 16, True, False, 0, "实际销售") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("充值扣率"))), 3, True, False, 0, "充值扣率") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("本次充值"))), 16, True, False, 0, "本次充值") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If
                If zlDblIsValid(Trim(.TextMatrix(lngRow, .ColIndex("实际充值缴款"))), 16, True, False, 0, "实际充值缴款") = False Then
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If



                If Val(Trim(.TextMatrix(lngRow, .ColIndex("卡面额")))) < Val(Trim(.TextMatrix(lngRow, .ColIndex("实际销售")))) Then
                    ShowMsgbox "注意:" & vbCrLf & "卡面额不能小于实际销售额,请检查!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If

                If Val(.TextMatrix(lngRow, .ColIndex("本次充值"))) < Val(.TextMatrix(lngRow, .ColIndex("实际充值缴款"))) Then
                    ShowMsgbox "注意:" & vbCrLf & "本次充值不能小于实际充值缴款,请检查!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                    Exit Function
                End If

                 If Val(.TextMatrix(lngRow, .ColIndex("充值扣率"))) > 100 Then
                     ShowMsgbox "注意:" & vbCrLf & "充值扣率不能大于100%,请检查!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                     Exit Function
                 End If
                 If Val(.TextMatrix(lngRow, .ColIndex("充值扣率"))) < 0 Then
                     ShowMsgbox "注意:" & vbCrLf & "充值扣率不能小于0,请检查!"
                    .Row = lngRow: zlCtlSetFocus vsGrid, True
                     Exit Function
                 End If

            End If
        Next
    End With
    CheckVsGridData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Check缴款情况() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查缴款情况
    '编制:刘兴洪
    '日期:2009-12-11 13:56:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt实收合计).Text), 16, True, False, 0, "实收合计") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt实收合计), True
        Exit Function
    End If
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt本次缴款).Text), 16, True, False, 0, "本次缴款") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次缴款), True
        Exit Function
    End If
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt找补).Text), 16, True, False, 0, "找补") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt找补), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt找补).Tag) < 0 Then
        If Val(txtEdit(mtxtIdx.idx_txt本次缴款).Text) <> 0 Then
            ShowMsgbox "注意:" & vbCrLf & "    你输入的缴款金额不足(还应收:" & Trim(txtEdit(mtxtIdx.idx_txt找补).Text) & ")　 ,不能继续!"
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次缴款), True
            Exit Function
        End If
        If mCardEditType = gEd_发卡 Then
            If MsgBox("注意:" & vbCrLf & "   你还未收取领卡人的缴款,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次缴款), True
                Exit Function
            End If
        ElseIf mCardEditType = gEd_修改 Then
            If MsgBox("注意:" & vbCrLf & "   你还未收取领卡人的缴款,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次缴款), True
                Exit Function
            End If
        ElseIf mCardEditType = gEd_充值 Then
            If MsgBox("注意:" & vbCrLf & "   你还未收取相关的缴款,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次缴款), True
                Exit Function
            End If
        Else
            If MsgBox("注意:" & vbCrLf & "   你还未收取相关的缴款,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次缴款), True
                Exit Function
            End If
        End If
    End If
    Check缴款情况 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入项是否合法
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-14 15:30:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt发卡原因).Text), 50, 0, "发卡原因") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt发卡原因), True
        Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt备注).Text), 100, 0, "备注") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt备注), True
        Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt充值备注).Text), 100, 0, "充值说明") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt充值备注), True
        Exit Function
    End If
    '76213,李南春,2014-08-12,卡回收时不进行密码验证
    If mCardEditType <> gEd_修改 And mCardEditType <> gEd_回收 Then
        If Trim(txtEdit(mtxtIdx.idx_txt确认密码).Text) <> Trim(txtEdit(mtxtIdx.idx_txt密码).Text) Then
            ShowMsgbox "密码输入不正确,请检查密码和确认密码是否正确!"
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt密码), True
            Exit Function
        End If
        If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt密码).Text), 20, 0, "密码") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt密码), True
            Exit Function
        End If
        If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt确认密码).Text), 20, 0, "确认密码") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt确认密码), True
            Exit Function
        End If
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mtxtIdx.idx_txt领卡人).Text), 20, 0, "领卡人") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt领卡人), True
        Exit Function
    End If
    If Trim(txtEdit(mtxtIdx.idx_txt领卡部门).Text) <> "" And Val(txtEdit(mtxtIdx.idx_txt领卡部门).Tag) = 0 Then
        ShowMsgbox "注意:" & vbCrLf & "    你输入的领卡部门有误，请检查!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt领卡部门), True
        Exit Function
    End If
    '金额检查
   If CheckInput卡面额 = False Then Exit Function
   If CheckInput实际销售额 = False Then Exit Function
   If CheckInput实际充值缴款 = False Then Exit Function
   If CheckInput充值扣率 = False Then Exit Function
   If CheckInput本次充值 = False Then Exit Function
    CheckInput = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的合法性
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-11 10:50:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Err = 0: On Error GoTo Errhand:

    '检查卡号
    Select Case mCardEditType
    Case gEd_发卡
        If CheckVsGridData = False Then Exit Function
        If Check缴款情况 = False Then Exit Function
    Case gEd_回收
        If CheckVsGridData = False Then Exit Function
    Case gEd_修改
        If 检查卡号是否合法("") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡号), True
            Exit Function
        End If
        If CheckInput() = False Then Exit Function

    Case gEd_充值
        If mlng消费卡ID = 0 Then
            ShowMsgbox "未选择合法的消费卡,不能充值,请检查"
            Exit Function
        End If
        If 检查卡号是否合法("") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡号), True
            Exit Function
        End If
       If Check缴款情况 = False Then Exit Function
    Case gEd_回退

    Case gEd_退卡
        If 检查卡号是否合法("") = False Then
            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡号), True
            Exit Function
        End If
    End Select

    isValied = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function SavePayCard(ByVal lng发卡序号 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存发卡信息
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-10 16:14:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As New Collection, lngRow  As Long, lngID As Long, strCardNos As String, strTemp As String, strCurNos As String, i As Long
    Dim str发卡时间 As String, varData As Variant
    str发卡时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Set cllPro = New Collection
    With vsGrid
        For lngRow = 1 To .Rows - 1
            strTemp = Trim(.Cell(flexcpData, lngRow, .ColIndex("卡号")))
            If strTemp <> "" Then
                If zlCommFun.ActualLen(strTemp) > 4000 Then
                    Do While True
                        i = InStr(3900, strTemp, ",")
                        If i > 0 Then
                            strCurNos = Mid(strTemp, 1, i - 1)
                            strTemp = Mid(strTemp, i + 1)
                            If strCurNos <> "" Then
                                If zl消费卡InsertSQL(lng发卡序号, strCurNos, str发卡时间, lngRow, cllPro) = False Then Exit Function
                            End If
                        Else
                            If strTemp <> "" Then
                                If zl消费卡InsertSQL(lng发卡序号, strTemp, str发卡时间, lngRow, cllPro) = False Then Exit Function
                            End If
                            Exit Do
                        End If
                    Loop
                Else
                    If zl消费卡InsertSQL(lng发卡序号, strTemp, str发卡时间, lngRow, cllPro) = False Then Exit Function
                End If
            End If
        Next
    End With
    If cllPro.count = 0 Then Exit Function
    Err = 0: On Error GoTo Errhand:
    ExecuteProcedureArrAy cllPro, Me.Caption
    SavePayCard = True
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtEdit_Change(Index As Integer)
    If mblnNoClick Then Exit Sub
    mblnChange = True
    '--问题47855
    If Index = mtxtIdx.idx_txt密码 Or Index = mtxtIdx.idx_txt确认密码 Then
        vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("密码")) = "******************": vsGrid.Cell(flexcpData, vsGrid.Row, vsGrid.ColIndex("密码")) = Trim(txtEdit(mtxtIdx.idx_txt密码).Text)
        vsGrid.Cell(flexcpData, vsGrid.Row, vsGrid.ColIndex("ID")) = Trim(txtEdit(mtxtIdx.idx_txt确认密码).Text)  '存的确定密码
    End If
    txtEdit(Index).Tag = ""
    If Index = mtxtIdx.idx_txt开始卡号 Then
        txtEdit(mtxtIdx.idx_txt结束卡号) = ""
    ElseIf Index = mtxtIdx.idx_txt本次缴款 Or Index = mtxtIdx.idx_txt实收合计 Then
        Call SetLblCatpion
    ElseIf mtxtIdx.idx_txt实际销售额 = Index Then
        Call Set可否充值
    End If
End Sub

Public Sub SetLblCatpion()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置找补的标题
    '编制：刘兴洪
    '日期：2010-03-22 16:01:21
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim dbl找补 As Double
   dbl找补 = Val(txtEdit(mtxtIdx.idx_txt本次缴款).Text) - mdbl实收合计
   If dbl找补 >= 0 Then
         txtEdit(mtxtIdx.idx_txt找补).Text = Format(dbl找补, "0.00")
         txtEdit(mtxtIdx.idx_txt找补).ForeColor = &H80000008
         lblEdit(mlblIdx.idx_lbl找补).Caption = "找补"
         txtEdit(mtxtIdx.idx_txt找补).Tag = txtEdit(mtxtIdx.idx_txt找补).Text

   Else
         txtEdit(mtxtIdx.idx_txt找补).Text = Format(Abs(dbl找补), "0.00")
         txtEdit(mtxtIdx.idx_txt找补).Tag = Format(dbl找补, "0.00")
         txtEdit(mtxtIdx.idx_txt找补).ForeColor = vbRed
         lblEdit(mlblIdx.idx_lbl找补).Caption = "应收"
   End If

End Sub


Private Sub txtEdit_GotFocus(Index As Integer)
    Select Case Index
    Case mtxtIdx.idx_txt开始卡号, mtxtIdx.idx_txt结束卡号, mtxtIdx.idx_txt卡号

        If mtxtIdx.idx_txt结束卡号 = Index Then
            gTy_TestBug.strStartNo = Trim(txtEdit(mtxtIdx.idx_txt开始卡号))
        Else
            gTy_TestBug.strStartNo = ""
        End If
'        If Not mobjBrushCard Is Nothing Then Call mobjBrushCard.zlSetAutoBrush(Trim(txtEdit(Index).Text) = "")

    Case mtxtIdx.idx_txt备注, mtxtIdx.idx_txt充值备注, mtxtIdx.idx_txt发卡原因, mtxtIdx.idx_txt缴款人, mtxtIdx.idx_txt领卡部门, mtxtIdx.idx_txt领卡人
        zlCommFun.OpenIme True
    Case mtxtIdx.idx_txt密码
        Call OpenPassKeyboard(txtEdit(Index), False)
    Case mtxtIdx.idx_txt确认密码
        Call OpenPassKeyboard(txtEdit(Index), True)
    Case Else
        zlCommFun.OpenIme False
    End Select
    zlControl.TxtSelAll txtEdit(Index)
End Sub
Private Function zlBrusCard(ByVal strCardNo As String, Optional blnBrushCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡操作
    '编制:刘兴洪
    '日期:2009-12-16 10:33:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer, blnModifyCard As Boolean
    If Me.ActiveControl Is txtEdit(mtxtIdx.idx_txt开始卡号) Then
        intIndex = mtxtIdx.idx_txt开始卡号
    ElseIf Me.ActiveControl Is txtEdit(mtxtIdx.idx_txt结束卡号) Then
        intIndex = mtxtIdx.idx_txt结束卡号
    ElseIf Me.ActiveControl Is txtEdit(mtxtIdx.idx_txt卡号) Then
        intIndex = mtxtIdx.idx_txt卡号
    Else
        zlBrusCard = True
        Exit Function
    End If
    txtEdit(intIndex).Text = strCardNo
    txtEdit(intIndex).Tag = strCardNo

    Select Case intIndex
    Case mtxtIdx.idx_txt开始卡号, mtxtIdx.idx_txt结束卡号

        If mCardEditType = gEd_发卡 Then
           If InsertIntoGrid(strCardNo, mtxtIdx.idx_txt结束卡号 = intIndex, , False) = False Then
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex): Exit Function
           Else
                '从grid控件中往textbox控件中加数
                Call FromGridToCtlData
                '因为可能存在续续发卡的情况,因此就移动到下一控件上了
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex)
           End If

        ElseIf mCardEditType = gEd_回收 Then
           If InsertIntoGrid(strCardNo, False) = False Then
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex): Exit Function
           Else
                '因为可能存在续续回收的情况,因此就移动到下一控件上了
                '清除当前的卡号，主要是允许重新刷卡
                txtEdit(intIndex).Text = ""
                zl_CtlSetFocus txtEdit(intIndex)
           End If
        End If
    Case mtxtIdx.idx_txt卡号
        Select Case mCardEditType
        Case gEd_发卡
            '在编辑框输入时,才默认为修改原来的卡号
            ' 1.当前在卡号处手工输入
            ' 2.刷卡的自动增加
            blnModifyCard = (intIndex = mtxtIdx.idx_txt卡号) And (blnBrushCard = False)

           If InsertIntoGrid(strCardNo, False, , blnModifyCard) = False Then
                zlControl.TxtSelAll txtEdit(intIndex)
                 Exit Function
           Else
                '因为可能存在续续发卡的情况,因此就移动到下一控件上了
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex)
           End If
        Case gEd_充值
            If zlFromCardNOGetDataToCtrl(strCardNo) = False Then
                zl_CtlSetFocus txtEdit(intIndex), True
                Call SetEditProperty
                Exit Function
            End If
            Call SetEditProperty
            zl_CtlSetFocus txtEdit(intIndex)
            zlCommFun.PressKey vbKeyTab
        Case gEd_回收
           If InsertIntoGrid(strCardNo, mtxtIdx.idx_txt结束卡号 = intIndex) = False Then
                zlControl.TxtSelAll txtEdit(intIndex)
                zl_CtlSetFocus txtEdit(intIndex): Exit Function
           Else
                '因为可能存在续续回收的情况,因此就移动到下一控件上了
                '清除当前的卡号，主要是允许重新刷卡
                'txtEdit(intIndex).Text = ""
                zl_CtlSetFocus txtEdit(intIndex)
           End If
        End Select

    Case gEd_回退
        If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        zl_CtlSetFocus txtEdit(intIndex)
        zlCommFun.PressKey vbKeyTab
    Case gEd_取消回收
        If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        zl_CtlSetFocus txtEdit(intIndex)
        zlCommFun.PressKey vbKeyTab
    Case gEd_取消退卡
        If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        zl_CtlSetFocus txtEdit(intIndex)
        zlCommFun.PressKey vbKeyTab
    Case gEd_修改
        If zlFromCardNOGetDataToCtrl(strCardNo) = False Then Exit Function
        zl_CtlSetFocus txtEdit(intIndex)
        zlCommFun.PressKey vbKeyTab
    Case Else
    End Select

    zlBrusCard = True
End Function


Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim str编码 As String, str名称 As String, lngID As Long
    Dim strCardNo As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case mtxtIdx.idx_txt开始卡号, mtxtIdx.idx_txt结束卡号, mtxtIdx.idx_txt卡号
        'If txtEdit(Index) = "" Then Exit Sub
        If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab
''        If txtEdit(Index).Text = "" Then
''            '直接调读卡
''            If mobjBrushCard.zlReadCard(Me, strCardNo) = False Then
''                Exit Sub
''            End If
''            txtEdit(Index).Text = strCardNo
''            txtEdit(Index).Tag = strCardNo
''        End If
''        Call zlBrusCard(Trim(txtEdit(Index)), False)
    Case mtxtIdx.idx_txt领卡人
        If Trim(txtEdit(Index).Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txtEdit(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub

        '选择人员
        lngID = Val(txtEdit(mtxtIdx.idx_txt领卡部门).Tag)
        If Select人员选择器(Me, txtEdit(Index), Trim(txtEdit(Index).Text), lngID, , True, , , , , , , "") = False Then
            zlCommFun.PressKey vbKeyTab
        End If
        If mCardEditType = gEd_发卡 Then
            '领卡人就是缴款人
            txtEdit(mtxtIdx.idx_txt缴款人).Text = txtEdit(mtxtIdx.idx_txt领卡人).Text
            txtEdit(mtxtIdx.idx_txt缴款人).Tag = txtEdit(mtxtIdx.idx_txt领卡人).Tag
        End If

        '需要读取缺省部门:
        If zl_From人员获取缺省部门(Val(txtEdit(mtxtIdx.idx_txt领卡人).Tag), str编码, str名称, lngID) Then
            txtEdit(mtxtIdx.idx_txt领卡部门).Text = str编码 & "-" & str名称
            txtEdit(mtxtIdx.idx_txt领卡部门).Tag = lngID
        End If
        Exit Sub
    Case mtxtIdx.idx_txt缴款人
        '缴款人,不确定,不选择
        zlCommFun.PressKey vbKeyTab: Exit Sub
    Case mtxtIdx.idx_txt领卡部门
        '选择部门
        If Trim(txtEdit(Index).Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txtEdit(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        '选择缺省部门
        lngID = Val(txtEdit(mtxtIdx.idx_txt领卡人).Tag)
        If Select部门选择器(Me, txtEdit(Index), Trim(txtEdit(Index).Text), "", IIf(lngID = 0, False, True), "", 0, "部门选择器", , , , , lngID) = False Then
            Exit Sub
        End If
    Case mtxtIdx.idx_txt发卡原因
        '选择发卡原因
        If Trim(txtEdit(Index).Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txtEdit(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If zl_SelectAndNotAddItem(Me, txtEdit(Index), Trim(txtEdit(Index).Text), "常用发卡原因", "常用发卡原因选择", True, True, , , , True) = False Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    Case mtxtIdx.idx_txt确认密码
        If Trim(txtEdit(mtxtIdx.idx_txt密码).Text) <> "" Then
            If Trim(txtEdit(Index).Text) = "" Then
                ShowMsgbox "请输入确定密码,请检查!"
                zl_CtlSetFocus txtEdit(Index)
                zlControl.TxtSelAll txtEdit(Index)
                Exit Sub
            End If
            If Trim(txtEdit(Index).Text) <> Trim(txtEdit(mtxtIdx.idx_txt密码).Text) Then
                ShowMsgbox "输入的密码不一致,请检查!"
                zl_CtlSetFocus txtEdit(Index)
                zlControl.TxtSelAll txtEdit(Index)
                Exit Sub
            End If
        End If
        If Trim(txtEdit(Index).Text) <> "" And Trim(txtEdit(mtxtIdx.idx_txt密码).Text) = "" Then
            ShowMsgbox "密码未输入,请检查!"
            zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt密码)
            Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt卡面额
        If CheckInput卡面额 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab

    Case mtxtIdx.idx_txt实际销售额
        If CheckInput实际销售额 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt实际充值缴款
        If CheckInput实际充值缴款 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab

    Case mtxtIdx.idx_txt本次充值
        If CheckInput本次充值 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt充值扣率
        If CheckInput充值扣率 = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt本次缴款
        If mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改 Then
            If MsgBox("你是否真要进行相关的" & IIf(mCardEditType = gEd_修改, "修改", "发卡") & "操作吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call SaveData
            Else
                If txtEdit(mtxtIdx.idx_txt本次缴款).Enabled And txtEdit(mtxtIdx.idx_txt本次缴款).Visible Then
                    zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次缴款)
                End If
            End If
        End If
        If mCardEditType = gEd_充值 Then
            If MsgBox("你是否真要进行相关的充值操作吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call SaveData
            End If
        End If
    Case Else
        If Index = mtxtIdx.idx_txt充值备注 Then
            zlCommFun.PressKey vbKeyTab
        End If
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Function CheckInput充值扣率() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查充值扣率是否合法
    '编制:刘兴洪
    '日期:2009-12-17 16:03:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt充值扣率).Text), 3, True, False, 0, "充值扣率") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt充值扣率), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt充值扣率).Text) > 100 Then
        ShowMsgbox "注意:" & vbCrLf & "充值扣率不能大于100%,请检查!"
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt充值扣率)
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt充值扣率)
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt充值扣率).Text) < 0 Then
        ShowMsgbox "注意:" & vbCrLf & "充值扣率不能小于0,请检查!"
        zl_CtlSetFocus txtEdit(mtxtIdx.idx_txt充值扣率)
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt充值扣率)
        Exit Function
    End If
    CheckInput充值扣率 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInput实际销售额() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查实际销售额
    '返回:
    '编制:刘兴洪
    '日期:2009-12-17 16:11:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt实际销售额).Text), 16, True, False, 0, "实际销售") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt实际销售额), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt卡面额).Text) < Val(txtEdit(mtxtIdx.idx_txt实际销售额).Text) Then
        ShowMsgbox "注意:" & vbCrLf & "卡面额不能小于实际销售额,请检查!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt实际销售额), True
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt实际销售额)
        Exit Function
    End If
    CheckInput实际销售额 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInput本次充值() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查本次充值
    '返回:
    '编制:刘兴洪
    '日期:2009-12-17 16:11:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt本次充值).Text), 16, True, False, 0, "本次充值") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次充值), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt本次充值).Text) < Val(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text) Then
        ShowMsgbox "注意:" & vbCrLf & "本次充值不能小于实际充值缴款,请检查!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次充值), True
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt本次充值)
        Exit Function
    End If
    CheckInput本次充值 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function CheckInput实际充值缴款() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查本次充值
    '返回:
    '编制:刘兴洪
    '日期:2009-12-17 16:11:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text), 16, True, False, 0, "实际充值缴款") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt实际充值缴款), True
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_txt本次充值).Text) < Val(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text) Then
        ShowMsgbox "注意:" & vbCrLf & "本次充值不能小于实际充值缴款,请检查!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt实际充值缴款), True
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt实际充值缴款)
        Exit Function
    End If
    CheckInput实际充值缴款 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInput卡面额() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查卡面额
    '编制:刘兴洪
    '日期:2009-12-17 16:08:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:

    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt卡面额).Text), 16, True, False, 0, "卡面额") = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡面额), True
        Exit Function
    End If

    If Val(txtEdit(mtxtIdx.idx_txt卡面额).Text) < Val(txtEdit(mtxtIdx.idx_txt实际销售额).Text) Then
        ShowMsgbox "注意:" & vbCrLf & "卡面额不能小于实际销售额,请检查!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡面额), True
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_txt卡面额)
        Exit Function
    End If

    CheckInput卡面额 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function



Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case mtxtIdx.idx_txt本次充值, mtxtIdx.idx_txt本次缴款, mtxtIdx.idx_txt充值扣率, mtxtIdx.idx_txt当前余额, mtxtIdx.idx_txt卡面额, mtxtIdx.idx_txt实际充值缴款, mtxtIdx.idx_txt实际销售额, mtxtIdx.idx_txt实收合计, mtxtIdx.idx_txt找补
        Call zlControl.TxtCheckKeyPress(txtEdit(Index), KeyAscii, m金额式)
    Case mtxtIdx.idx_txt卡号, mtxtIdx.idx_txt开始卡号, mtxtIdx.idx_txt结束卡号
        Call zlControl.TxtCheckKeyPress(txtEdit(Index), KeyAscii, m文本式)
        If InStr(1, "'~～|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0

        Call BrushCard(txtEdit(Index), KeyAscii)

    Case Else
        Call zlControl.TxtCheckKeyPress(txtEdit(Index), KeyAscii, m文本式)
    End Select
End Sub
 Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡操作
    '编制:刘兴洪
    '日期:2010-02-09 14:07:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean

    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    blnCard = zlCommFun.InputIsCard(objEdit, KeyAscii, mTy_MoudlePara.bln卡号密文)
    If blnCard And Len(objEdit.Text) = objEdit.MaxLength - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(objEdit.Text) <> "" Then
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        Call zlBrusCard(Trim(objEdit), False)
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = idx_txt领卡部门 Then
        If txtEdit(idx_txt领卡部门).Tag = "" And txtEdit(idx_txt领卡部门).Text <> "" Then txtEdit(idx_txt领卡部门).Text = ""
    End If
    Select Case Index
    Case mtxtIdx.idx_txt密码
        Call ClosePassKeyboard(txtEdit(Index))
    Case mtxtIdx.idx_txt确认密码
        Call ClosePassKeyboard(txtEdit(Index))
    End Select
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If mCardEditType <> gEd_充值 And mCardEditType <> gEd_发卡 And mCardEditType <> gEd_修改 Then Exit Sub

    Select Case Index
    Case mtxtIdx.idx_txt卡面额
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
        Call Calc余额
        If mCardEditType = gEd_充值 Or mCardEditType = gEd_修改 Then Calc实收合计
    Case mtxtIdx.idx_txt实际销售额
        Call Calc余额
        Call Set可否充值
        If mCardEditType = gEd_充值 Or mCardEditType = gEd_修改 Then Calc实收合计
    Case mtxtIdx.idx_txt本次充值
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
        txtEdit(mtxtIdx.idx_txt实际充值缴款).Text = Format(Val(txtEdit(Index).Text) * (Round(Val(txtEdit(mtxtIdx.idx_txt充值扣率)) / 100, 6)), "0.00")
        Call Calc余额
        If mCardEditType = gEd_充值 Or mCardEditType = gEd_修改 Then Calc实收合计
    Case mtxtIdx.idx_txt充值扣率
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
        txtEdit(mtxtIdx.idx_txt实际充值缴款).Text = Format(Val(txtEdit(mtxtIdx.idx_txt本次充值).Text) * (Round(Val(txtEdit(mtxtIdx.idx_txt充值扣率)) / 100, 4)), "0.00")
        Call Calc余额
        If mCardEditType = gEd_充值 Or mCardEditType = gEd_修改 Then Calc实收合计
    Case mtxtIdx.idx_txt实际充值缴款
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
        If Val(txtEdit(mtxtIdx.idx_txt本次充值).Text) <> 0 Then
            txtEdit(mtxtIdx.idx_txt充值扣率).Text = Format((Round(Val(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text) / Val(txtEdit(mtxtIdx.idx_txt本次充值).Text), 6)) * 100, "0.00")
        Else
             txtEdit(mtxtIdx.idx_txt本次充值).Text = txtEdit(mtxtIdx.idx_txt实际充值缴款).Text
        End If
        Call Calc余额
        If mCardEditType = gEd_充值 Or mCardEditType = gEd_修改 Then Calc实收合计
    Case Else

    End Select

    If mCardEditType = gEd_发卡 Or mCardEditType = gEd_修改 Then
        '修改的话,需要同步更新网格行的数据才行
        With vsGrid
            If Split(.Cell(flexcpData, .Row, .ColIndex("卡号")) & ",", ",")(0) <> txtEdit(mtxtIdx.idx_txt卡号).Text Then Exit Sub
            '只有相同数据才可以更新
            Select Case Index
            Case idx_txt卡号
                '不管
            Case idx_txt密码
                '--问题47855
                '.TextMatrix(.Row, .ColIndex("密码")) = "******************": .Cell(flexcpData, .Row, .ColIndex("密码")) = Trim(txtEdit(mtxtIdx.idx_txt密码).Text)
                If Trim(txtEdit(mtxtIdx.idx_txt确认密码).Text) <> Trim(txtEdit(mtxtIdx.idx_txt密码).Text) And Trim(txtEdit(mtxtIdx.idx_txt确认密码).Text) <> "" Then
                     ShowMsgbox "输入密码与确定密码不一致,请检查"
                     txtEdit(mtxtIdx.idx_txt确认密码).SetFocus
                     'If Trim(txtEdit(mtxtIdx.idx_txt密码).Text) <> "" Then Cancel = True: Exit Sub
                End If
            Case idx_txt确认密码
                   If Trim(txtEdit(mtxtIdx.idx_txt确认密码).Text) <> Trim(txtEdit(mtxtIdx.idx_txt密码).Text) Then
                        ShowMsgbox "输入密码与确定密码不一致,请检查"
                        txtEdit(mtxtIdx.idx_txt密码).SetFocus
                        'If Trim(txtEdit(mtxtIdx.idx_txt确认密码).Text) <> "" Then Cancel = True: Exit Sub
                   End If
                   '--问题47855
                   ' .Cell(flexcpData, .Row, .ColIndex("ID")) = Trim(txtEdit(mtxtIdx.idx_txt确认密码).Text)  '存的确定密码
            Case idx_txt发卡原因
                .TextMatrix(.Row, .ColIndex("发卡原因")) = Trim(txtEdit(mtxtIdx.idx_txt发卡原因).Text)
            Case idx_txt领卡人
                .TextMatrix(.Row, .ColIndex("领卡人")) = Trim(txtEdit(mtxtIdx.idx_txt领卡人).Text): .Cell(flexcpData, .Row, .ColIndex("领卡人")) = Trim(txtEdit(mtxtIdx.idx_txt领卡人).Tag)
                .TextMatrix(.Row, .ColIndex("领卡部门")) = Trim(txtEdit(mtxtIdx.idx_txt领卡部门).Text): .Cell(flexcpData, .Row, .ColIndex("领卡部门")) = Val(txtEdit(mtxtIdx.idx_txt领卡部门).Tag)
                .TextMatrix(.Row, .ColIndex("充值缴款人")) = Trim(txtEdit(mtxtIdx.idx_txt领卡人).Text): .Cell(flexcpData, .Row, .ColIndex("充值缴款人")) = Trim(txtEdit(mtxtIdx.idx_txt领卡人).Tag)

            Case idx_txt卡有效日期
            Case idx_txt备注
                .TextMatrix(.Row, .ColIndex("备注")) = Trim(txtEdit(mtxtIdx.idx_txt备注).Text)
            Case idx_txt发卡人
            Case idx_txt发卡日期
            Case idx_txt当前余额
                txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
                .TextMatrix(.Row, .ColIndex("当前余额")) = Trim(txtEdit(mtxtIdx.idx_txt当前余额).Text)
            Case idx_txt卡面额
                txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
                .TextMatrix(.Row, .ColIndex("卡面额")) = Trim(txtEdit(mtxtIdx.idx_txt卡面额).Text)
                .TextMatrix(.Row, .ColIndex("当前余额")) = Trim(txtEdit(mtxtIdx.idx_txt当前余额).Text)
                Calc实收合计
            Case idx_txt实际销售额
                txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
                .TextMatrix(.Row, .ColIndex("实际销售")) = Trim(txtEdit(mtxtIdx.idx_txt实际销售额).Text)
                .TextMatrix(.Row, .ColIndex("当前余额")) = Trim(txtEdit(mtxtIdx.idx_txt当前余额).Text)
                Calc实收合计
            Case idx_txt本次充值, idx_txt充值扣率, idx_txt实际充值缴款
                txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
                If chk是否充值.value = 0 Then
                    .TextMatrix(.Row, .ColIndex("充值扣率")) = ""
                    .TextMatrix(.Row, .ColIndex("本次充值")) = ""
                    .TextMatrix(.Row, .ColIndex("实际充值缴款")) = ""
                    .TextMatrix(.Row, .ColIndex("结算方式")) = cboStyle.Text
                Else
                    .TextMatrix(.Row, .ColIndex("充值扣率")) = Trim(txtEdit(mtxtIdx.idx_txt充值扣率).Text)
                    .TextMatrix(.Row, .ColIndex("本次充值")) = Trim(txtEdit(mtxtIdx.idx_txt本次充值).Text)
                    .TextMatrix(.Row, .ColIndex("实际充值缴款")) = Trim(txtEdit(mtxtIdx.idx_txt实际充值缴款).Text)
                    .TextMatrix(.Row, .ColIndex("结算方式")) = cboStyle.Text
                End If
                .TextMatrix(.Row, .ColIndex("当前余额")) = Trim(txtEdit(mtxtIdx.idx_txt当前余额).Text)
                Calc实收合计
               Case idx_txt开户行, idx_txt帐号, idx_txt结算号码
                If chk是否充值.value = 0 Then
                    .TextMatrix(.Row, .ColIndex("开户行")) = ""
                    .TextMatrix(.Row, .ColIndex("帐号")) = ""
                    .TextMatrix(.Row, .ColIndex("结算号码")) = ""
                    .TextMatrix(.Row, .ColIndex("结算方式")) = cboStyle.Text
                Else
                    .TextMatrix(.Row, .ColIndex("开户行")) = Trim(txtEdit(mtxtIdx.idx_txt开户行).Text)
                    .TextMatrix(.Row, .ColIndex("帐号")) = Trim(txtEdit(mtxtIdx.idx_txt帐号).Text)
                    .TextMatrix(.Row, .ColIndex("结算号码")) = Trim(txtEdit(mtxtIdx.idx_txt结算号码).Text)
                    .TextMatrix(.Row, .ColIndex("结算方式")) = cboStyle.Text
                End If
                 
            
            Case idx_txt找补
            Case idx_txt实收合计
            Case idx_txt本次缴款
            Case idx_txt领卡部门
                .TextMatrix(.Row, .ColIndex("领卡部门")) = Trim(txtEdit(mtxtIdx.idx_txt领卡部门).Text): .Cell(flexcpData, .Row, .ColIndex("领卡部门")) = Val(txtEdit(mtxtIdx.idx_txt领卡部门).Tag)
            Case idx_txt开始卡号
                Call txtEdit_KeyPress(idx_txt开始卡号, 13)
            Case idx_txt结束卡号
            Case idx_txt回收人
            Case idx_txt回收时间
            Case idx_txt缴款人
                .TextMatrix(.Row, .ColIndex("充值缴款人")) = Trim(txtEdit(mtxtIdx.idx_txt缴款人).Text)
            Case idx_txt充值备注
                .TextMatrix(.Row, .ColIndex("充值说明")) = Trim(txtEdit(mtxtIdx.idx_txt充值备注).Text)
            End Select
       End With
    End If
End Sub
Private Function FromGridToCtlData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从指定行的数据向网格行插入数据
    '编制:刘兴洪
    '日期:2009-12-17 17:42:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFind As Boolean, lngRow As Long, strTemp As String, i As Long
    Dim strCards As String
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        lngRow = .Row
        If Trim(.TextMatrix(lngRow, .ColIndex("卡号"))) = "" Then
            Call ClearCtlData: Call SetDefaultValue(True): FromGridToCtlData = True: Exit Function
        End If
        
        '95344: 李南春,2016/4/22,卡号长度不够的情况下，正确获取批量的第一张卡号
        strCards = Split(Trim(.TextMatrix(lngRow, .ColIndex("卡号"))) & "(", "(")(0)
        txtEdit(mtxtIdx.idx_txt卡号).Text = Split(strCards & "～", "～")(0)
        mblnNoClick = True
        With cbo卡类型
            .ListIndex = -1
            strTemp = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("卡类型")): blnFind = False
            For i = 0 To .ListCount - 1
                If .List(i) & ";" Like "*." & strTemp & ";" Then
                    blnFind = True
                    .ListIndex = i: Exit For
                End If
            Next
            If blnFind = False And strTemp <> "" Then
                .AddItem strTemp
                .ListIndex = .NewIndex
            End If
        End With

        With lvwType
            strTemp = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("限制类别")): blnFind = False
            For i = 1 To .ListItems.count
                If InStr(1, "," & strTemp & ",", "," & Mid(.ListItems(i).Key, 2) & ",") > 0 Then
                    .ListItems(i).Checked = True
                Else
                    .ListItems(i).Checked = False
                End If
            Next
        End With

        chk是否充值.value = IIf(Val(.Cell(flexcpData, lngRow, .ColIndex("是否充值"))) = 1, 1, 0)
        txtEdit(mtxtIdx.idx_txt密码).Text = .Cell(flexcpData, lngRow, .ColIndex("密码"))
        txtEdit(mtxtIdx.idx_txt确认密码).Text = .Cell(flexcpData, lngRow, .ColIndex("ID"))
        txtEdit(mtxtIdx.idx_txt发卡原因).Text = .TextMatrix(lngRow, .ColIndex("发卡原因")): txtEdit(mtxtIdx.idx_txt发卡原因).Tag = txtEdit(mtxtIdx.idx_txt发卡原因).Text
        txtEdit(mtxtIdx.idx_txt领卡人).Text = .TextMatrix(lngRow, .ColIndex("领卡人")): txtEdit(mtxtIdx.idx_txt领卡人).Tag = .Cell(flexcpData, lngRow, .ColIndex("领卡人"))
        txtEdit(mtxtIdx.idx_txt领卡部门).Text = .TextMatrix(lngRow, .ColIndex("领卡部门")): txtEdit(mtxtIdx.idx_txt领卡部门).Tag = .Cell(flexcpData, lngRow, .ColIndex("领卡部门"))
        txtEdit(mtxtIdx.idx_txt备注).Text = .TextMatrix(lngRow, .ColIndex("备注"))

        txtEdit(mtxtIdx.idx_txt发卡人).Text = .TextMatrix(lngRow, .ColIndex("发卡人"))
        txtEdit(mtxtIdx.idx_txt发卡日期).Text = .TextMatrix(lngRow, .ColIndex("发卡日期"))
        txtEdit(mtxtIdx.idx_txt当前余额).Text = .TextMatrix(lngRow, .ColIndex("当前余额"))
        If .TextMatrix(lngRow, .ColIndex("卡有效期")) = "" Or IsDate(.TextMatrix(lngRow, .ColIndex("卡有效期"))) = False Then
            dtp卡有效日期.value = Null
        Else
            dtp卡有效日期.value = CDate(.TextMatrix(lngRow, .ColIndex("卡有效期")))
        End If
        txtEdit(mtxtIdx.idx_txt卡面额).Text = .TextMatrix(lngRow, .ColIndex("卡面额"))
        txtEdit(mtxtIdx.idx_txt实际销售额).Text = .TextMatrix(lngRow, .ColIndex("实际销售"))
        If chk是否充值.value = 0 Then
            txtEdit(mtxtIdx.idx_txt充值扣率).Text = ""
            txtEdit(mtxtIdx.idx_txt本次充值).Text = ""
            txtEdit(mtxtIdx.idx_txt实际充值缴款).Text = ""
            Call SetDefaultValue(False)
        Else
            txtEdit(mtxtIdx.idx_txt充值扣率).Text = .TextMatrix(lngRow, .ColIndex("充值扣率"))
            txtEdit(mtxtIdx.idx_txt本次充值).Text = .TextMatrix(lngRow, .ColIndex("本次充值"))
            txtEdit(mtxtIdx.idx_txt实际充值缴款).Text = .TextMatrix(lngRow, .ColIndex("实际充值缴款"))
        End If
        txtEdit(mtxtIdx.idx_txt充值备注).Text = .TextMatrix(lngRow, .ColIndex("充值说明"))
        txtEdit(mtxtIdx.idx_txt缴款人).Text = .TextMatrix(lngRow, .ColIndex("充值缴款人"))
    End With
    mblnNoClick = False
    FromGridToCtlData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldCol <> NewCol Then
        cbsThis.RecalcLayout
    End If
    If OldRow = NewRow Then
        With vsGrid
            If Trim(txtEdit(mtxtIdx.idx_txt卡号).Text) <> "" Then
                Exit Sub
            End If
        End With
    End If
    If mblnNoClick Then Exit Sub
    Call FromGridToCtlData
End Sub


Private Sub vsGrid_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsGrid.ColIndex("标志") Then Cancel = True
End Sub

Private Sub vsGrid_GotFocus()
    zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub

Private Sub vsGrid_LostFocus()
    zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub


Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function


Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:lgf
    '日期:2012-12-2 11:11:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, str性质 As String


    If str性质 = "" Then str性质 = ",1,2" '1-现金,2.支票
    str性质 = Mid(str性质, 2)
    strSQL = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质" & _
        " From 结算方式 B" & _
        " Where Nvl(B.性质,1) In(" & str性质 & ")" & _
        " Order by B.编码"

    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboStyle
        .Clear:
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!名称)
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
            rsTemp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If cboStyle.ListCount = 0 Then
        MsgBox "预交场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

