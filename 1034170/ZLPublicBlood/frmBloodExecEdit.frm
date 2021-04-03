VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBloodExecEdit 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输血执行登记"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6825
   Icon            =   "frmBloodExecEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraExe 
      Caption         =   "输注过程"
      Height          =   1725
      Index           =   3
      Left            =   60
      TabIndex        =   18
      Top             =   3030
      Width           =   6705
      Begin VB.PictureBox picLeakage 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1680
         ScaleHeight     =   330
         ScaleWidth      =   1080
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   585
         Width           =   1080
         Begin VB.OptionButton optLeakage 
            Caption         =   "有"
            Height          =   240
            Index           =   1
            Left            =   630
            TabIndex        =   26
            Top             =   45
            Width           =   525
         End
         Begin VB.OptionButton optLeakage 
            Caption         =   "无"
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   25
            Top             =   45
            Value           =   -1  'True
            Width           =   525
         End
      End
      Begin VB.ComboBox cbo执行人 
         Height          =   300
         Index           =   2
         Left            =   4845
         TabIndex        =   22
         Text            =   "cbo执行人"
         Top             =   210
         Width           =   1815
      End
      Begin VB.ComboBox cboReaction 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   960
         Width           =   1815
      End
      Begin VB.PictureBox picReaction 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   4845
         ScaleHeight     =   330
         ScaleWidth      =   1815
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   585
         Width           =   1815
         Begin VB.OptionButton optBegin 
            Caption         =   "无"
            Height          =   240
            Index           =   4
            Left            =   0
            TabIndex        =   29
            Top             =   45
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton optBegin 
            Caption         =   "有"
            Height          =   240
            Index           =   5
            Left            =   1335
            TabIndex        =   30
            Top             =   45
            Width           =   525
         End
      End
      Begin VB.CommandButton cmdDate 
         Enabled         =   0   'False
         Height          =   240
         Index           =   4
         Left            =   6360
         Picture         =   "frmBloodExecEdit.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   990
         Width           =   255
      End
      Begin VB.ComboBox cbo滴数 
         Height          =   300
         Index           =   1
         Left            =   1290
         TabIndex        =   20
         Top             =   210
         Width           =   930
      End
      Begin MSMask.MaskEdBox txt反应时间 
         Height          =   300
         Index           =   2
         Left            =   4845
         TabIndex        =   34
         Top             =   960
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm"
         Mask            =   "####-##-## ##:##"
         PromptChar      =   "_"
      End
      Begin zlPublicBlood.UCPatiVitalSigns UCPatiVS 
         Height          =   360
         Index           =   1
         Left            =   165
         TabIndex        =   35
         Top             =   1335
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   529
         TextBackColor   =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         ShowMode        =   0
         XDis            =   120
      End
      Begin VB.Label lbl单位 
         AutoSize        =   -1  'True
         Caption         =   "滴/分"
         Height          =   180
         Index           =   1
         Left            =   2265
         TabIndex        =   77
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblLeakage 
         AutoSize        =   -1  'True
         Caption         =   "输血部位有无渗漏"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   23
         Top             =   630
         Width           =   1440
      End
      Begin VB.Label lblPeople 
         AutoSize        =   -1  'True
         Caption         =   "执 行 人"
         Height          =   180
         Index           =   2
         Left            =   4050
         TabIndex        =   21
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblRe 
         AutoSize        =   -1  'True
         Caption         =   "不良反应"
         Height          =   180
         Index           =   2
         Left            =   165
         TabIndex        =   31
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblReactionTime 
         AutoSize        =   -1  'True
         Caption         =   "反应时间"
         Height          =   180
         Index           =   2
         Left            =   4050
         TabIndex        =   33
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblReaction 
         AutoSize        =   -1  'True
         Caption         =   "输血反应"
         Height          =   180
         Index           =   2
         Left            =   4050
         TabIndex        =   27
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "15分钟后滴速"
         Height          =   180
         Index           =   2
         Left            =   165
         TabIndex        =   19
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Frame fraExe 
      Caption         =   "输注核对"
      Height          =   1005
      Index           =   2
      Left            =   60
      TabIndex        =   67
      Top             =   120
      Width           =   6705
      Begin VB.Timer TimeFlash 
         Interval        =   250
         Left            =   3075
         Top             =   -15
      End
      Begin VB.TextBox txtCheck 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "2012-11-21 10:20"
         Top             =   630
         Width           =   1815
      End
      Begin VB.TextBox txtCheck 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   4845
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "管理员"
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox txtCheck 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "管理员"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Image imgMore 
         Height          =   225
         Left            =   2805
         Picture         =   "frmBloodExecEdit.frx":0680
         Top             =   660
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "核对验证"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   4050
         MouseIcon       =   "frmBloodExecEdit.frx":0A81
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   645
         Width           =   720
      End
      Begin VB.Line linB 
         Index           =   2
         X1              =   960
         X2              =   2775
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Label lblExeTime 
         AutoSize        =   -1  'True
         Caption         =   "核对时间"
         Height          =   180
         Index           =   2
         Left            =   165
         TabIndex        =   72
         Top             =   645
         Width           =   720
      End
      Begin VB.Line linB 
         Index           =   1
         X1              =   4845
         X2              =   6660
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line linB 
         Index           =   0
         X1              =   960
         X2              =   2775
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblCheck 
         AutoSize        =   -1  'True
         Caption         =   "复 查 人"
         Height          =   180
         Index           =   1
         Left            =   4050
         TabIndex        =   70
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblCheck 
         AutoSize        =   -1  'True
         Caption         =   "核 查 人"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   68
         Top             =   270
         Width           =   720
      End
   End
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   90
      TabIndex        =   66
      Top             =   7770
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   327942145
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VB.Frame fraLine 
      Height          =   90
      Left            =   -45
      TabIndex        =   65
      Top             =   7440
      Width           =   6900
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4500
      TabIndex        =   59
      Top             =   7710
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   5625
      TabIndex        =   60
      Top             =   7710
      Width           =   1100
   End
   Begin VB.TextBox txt执行摘要 
      Height          =   735
      Left            =   960
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   58
      Top             =   6645
      Width           =   5790
   End
   Begin VB.Frame fraExe 
      Caption         =   "输注结束"
      Height          =   1725
      Index           =   1
      Left            =   60
      TabIndex        =   37
      Top             =   4830
      Width           =   6705
      Begin VB.PictureBox picLeakage 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   1680
         ScaleHeight     =   330
         ScaleWidth      =   1080
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   585
         Width           =   1080
         Begin VB.OptionButton optLeakage 
            Caption         =   "无"
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   45
            Top             =   45
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton optLeakage 
            Caption         =   "有"
            Height          =   240
            Index           =   3
            Left            =   630
            TabIndex        =   46
            Top             =   45
            Width           =   525
         End
      End
      Begin VB.CommandButton cmdDate 
         Enabled         =   0   'False
         Height          =   240
         Index           =   3
         Left            =   6360
         Picture         =   "frmBloodExecEdit.frx":2403
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   990
         Width           =   255
      End
      Begin VB.CommandButton cmdDate 
         Height          =   240
         Index           =   1
         Left            =   2475
         Picture         =   "frmBloodExecEdit.frx":24F9
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   240
         Width           =   255
      End
      Begin MSMask.MaskEdBox txt执行时间 
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   39
         Top             =   210
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm"
         Mask            =   "####-##-## ##:##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbo执行人 
         Height          =   300
         Index           =   1
         Left            =   4845
         TabIndex        =   42
         Text            =   "cbo执行人"
         Top             =   210
         Width           =   1815
      End
      Begin VB.PictureBox picReaction 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   4845
         ScaleHeight     =   330
         ScaleWidth      =   1815
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   585
         Width           =   1815
         Begin VB.OptionButton optBegin 
            Caption         =   "有"
            Height          =   240
            Index           =   3
            Left            =   1335
            TabIndex        =   50
            Top             =   45
            Width           =   525
         End
         Begin VB.OptionButton optBegin 
            Caption         =   "无"
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   49
            Top             =   45
            Value           =   -1  'True
            Width           =   525
         End
      End
      Begin VB.ComboBox cboReaction 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   960
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txt反应时间 
         Height          =   300
         Index           =   1
         Left            =   4845
         TabIndex        =   54
         Top             =   960
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm"
         Mask            =   "####-##-## ##:##"
         PromptChar      =   "_"
      End
      Begin zlPublicBlood.UCPatiVitalSigns UCPatiVS 
         Height          =   360
         Index           =   2
         Left            =   165
         TabIndex        =   55
         Top             =   1335
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   529
         TextBackColor   =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         ShowMode        =   0
         XDis            =   120
      End
      Begin VB.Label lblLeakage 
         AutoSize        =   -1  'True
         Caption         =   "输血部位有无渗漏"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   43
         Top             =   630
         Width           =   1440
      End
      Begin VB.Label lblExeTime 
         AutoSize        =   -1  'True
         Caption         =   "结束时间"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   38
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblPeople 
         AutoSize        =   -1  'True
         Caption         =   "执 行 人"
         Height          =   180
         Index           =   1
         Left            =   4050
         TabIndex        =   41
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblReaction 
         AutoSize        =   -1  'True
         Caption         =   "输血反应"
         Height          =   180
         Index           =   1
         Left            =   4050
         TabIndex        =   47
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblReactionTime 
         AutoSize        =   -1  'True
         Caption         =   "反应时间"
         Height          =   180
         Index           =   1
         Left            =   4050
         TabIndex        =   53
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblRe 
         AutoSize        =   -1  'True
         Caption         =   "不良反应"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   51
         Top             =   1020
         Width           =   720
      End
   End
   Begin VB.Frame fraExe 
      Caption         =   "输注开始"
      Height          =   1725
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   1230
      Width           =   6705
      Begin VB.ComboBox cbo滴数 
         Height          =   300
         Index           =   0
         Left            =   1290
         TabIndex        =   7
         Top             =   585
         Width           =   930
      End
      Begin VB.CommandButton cmdDate 
         Height          =   240
         Index           =   0
         Left            =   2475
         Picture         =   "frmBloodExecEdit.frx":25EF
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   240
         Width           =   255
      End
      Begin MSMask.MaskEdBox txt执行时间 
         Height          =   300
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   210
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm"
         Mask            =   "####-##-## ##:##"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdDate 
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   6360
         Picture         =   "frmBloodExecEdit.frx":26E5
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   990
         Width           =   255
      End
      Begin MSMask.MaskEdBox txt反应时间 
         Height          =   300
         Index           =   0
         Left            =   4845
         TabIndex        =   15
         Top             =   960
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm"
         Mask            =   "####-##-## ##:##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboReaction 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   1815
      End
      Begin VB.PictureBox picReaction 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   4845
         ScaleHeight     =   330
         ScaleWidth      =   1815
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   585
         Width           =   1815
         Begin VB.OptionButton optBegin 
            Caption         =   "无"
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   45
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton optBegin 
            Caption         =   "有"
            Height          =   240
            Index           =   1
            Left            =   1335
            TabIndex        =   11
            Top             =   45
            Width           =   525
         End
      End
      Begin VB.ComboBox cbo执行人 
         Height          =   300
         Index           =   0
         Left            =   4845
         TabIndex        =   5
         Text            =   "cbo执行人"
         Top             =   210
         Width           =   1815
      End
      Begin zlPublicBlood.UCPatiVitalSigns UCPatiVS 
         Height          =   360
         Index           =   0
         Left            =   165
         TabIndex        =   16
         Top             =   1335
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   529
         TextBackColor   =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         ShowMode        =   0
         XDis            =   120
      End
      Begin VB.Label lbl单位 
         AutoSize        =   -1  'True
         Caption         =   "滴/分"
         Height          =   180
         Index           =   0
         Left            =   2265
         TabIndex        =   76
         Top             =   630
         Width           =   450
      End
      Begin VB.Label lblRe 
         AutoSize        =   -1  'True
         Caption         =   "不良反应"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   12
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblReactionTime 
         AutoSize        =   -1  'True
         Caption         =   "反应时间"
         Height          =   180
         Index           =   0
         Left            =   4050
         TabIndex        =   14
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblReaction 
         AutoSize        =   -1  'True
         Caption         =   "输血反应"
         Height          =   180
         Index           =   0
         Left            =   4050
         TabIndex        =   8
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "前15分钟滴速"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label lblPeople 
         AutoSize        =   -1  'True
         Caption         =   "执 行 人"
         Height          =   180
         Index           =   0
         Left            =   4050
         TabIndex        =   4
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblExeTime 
         AutoSize        =   -1  'True
         Caption         =   "开始时间"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.PictureBox picHide 
      Height          =   465
      Left            =   4320
      ScaleHeight     =   405
      ScaleWidth      =   765
      TabIndex        =   61
      Top             =   4035
      Visible         =   0   'False
      Width           =   825
      Begin VB.TextBox txt本次数次 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   195
         TabIndex        =   63
         Top             =   0
         Width           =   1005
      End
      Begin VB.TextBox txt发送数次 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   120
         TabIndex        =   62
         Top             =   0
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker dtp要求时间 
         Height          =   300
         Left            =   15
         TabIndex        =   64
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   327942147
         CurrentDate     =   38082
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "提醒：输注开始后请在4h内完成血液输注"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   45
      TabIndex        =   75
      Top             =   7800
      Width           =   3525
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "执行摘要"
      Height          =   180
      Left            =   105
      TabIndex        =   57
      Top             =   6900
      Width           =   720
   End
End
Attribute VB_Name = "frmBloodExecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnAcTive As Boolean
Private mstr缺省输血反应 As String
Private mlngModul As Long
Private mlng收发ID As Long
Private mlng医嘱ID As Long, mlng相关ID As Long
Private mlng发送号 As Long
Private mlng科室ID As Long
Private mlng执行科室ID As Long
Private mstrPrivs As String
Private mblnOk As Boolean
Private mstr接收时间 As String '血液的接收时间
Private mint血袋数 As Integer, mint已执行血袋数 As Integer
Private mintTimerCount As Integer
Private mblnReturn As Boolean  '执行人快速输入匹配控制
Private mblnOnlyRead As Boolean '是否是只读模式

Private Type ExeInfo
    收发ID As Long
    开始执行人 As String
    开始时间  As String
    前15分钟滴速 As Integer
    输注前输血反应  As String
    输注前反应时间 As String
    输注中执行人 As String
    后15分钟滴速 As Integer   '输注15分钟后的滴数
    输注中输血反应  As String
    输注中反应时间 As String
    结束执行人 As String
    结束时间  As String
    输注后输血反应  As String
    输注后反应时间 As String
    执行科室ID  As Long
    核查者 As String
    复查者 As String
    核对时间 As String
    登记人 As String
    登记时间  As String
    摘要  As String
    输血部位渗漏 As String
End Type
Private mExeInfo As ExeInfo
Private mExeInfoSave As ExeInfo

Public Function ShowEdit(ByVal frmParent As Object, ByVal lngModul As Enum_Inside_Program, ByVal lng医嘱ID As Long, _
    ByVal lng发送号 As Long, ByVal lng科室ID As Long, ByVal lng收发ID As Long, ByVal lng执行科室ID As Long, Optional ByVal strPrivs As String, Optional ByVal blnOnlyRead As Boolean) As Boolean
    mblnOk = True
    mlngModul = lngModul
    mlng医嘱ID = lng医嘱ID
    mlng发送号 = lng发送号
    mlng科室ID = lng科室ID
    mlng收发ID = lng收发ID
    mlng执行科室ID = lng执行科室ID
    mstrPrivs = strPrivs
    mblnOnlyRead = blnOnlyRead
    On Error Resume Next
    Me.Show 1, frmParent
    If Err <> 0 Then Err.Clear
    ShowEdit = mblnOk
End Function

Private Sub ClearExeInfo()
    With mExeInfo
        .收发ID = 0
        .开始执行人 = ""
        .开始时间 = ""
        .前15分钟滴速 = 0
        .输注前输血反应 = ""
        .输注前反应时间 = ""
        .输注中执行人 = ""
        .后15分钟滴速 = 0
        .输注中输血反应 = ""
        .输注中反应时间 = ""
        .结束执行人 = ""
        .结束时间 = ""
        .输注后输血反应 = ""
        .输注后反应时间 = ""
        .执行科室ID = 0
        .核查者 = ""
        .复查者 = ""
        .核对时间 = ""
        .登记人 = ""
        .登记时间 = ""
        .摘要 = ""
        .输血部位渗漏 = ""
    End With
End Sub

Private Sub cbo滴数_Click(Index As Integer)
    If cbo滴数(Index).ListIndex = 2 Or cbo滴数(Index).ListIndex = 3 Then
        lbl单位(Index).Visible = False
    Else
        lbl单位(Index).Visible = True
    End If
End Sub

Private Sub cbo滴数_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call cbo滴数_Click(Index)
End Sub

Private Sub cbo执行人_Click(Index As Integer)
    Call cbo执行人_KeyPress(Index, vbKeyReturn)
End Sub

Private Sub cbo执行人_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strPersons As String
    Dim blnOk As Boolean
    
    mblnReturn = False
    If cbo执行人(Index).ListIndex <> -1 Then cbo执行人(Index).Tag = cbo执行人(Index).ListIndex
    strPersons = cbo执行人(Index).Text
    If KeyAscii = 13 Then
        mblnReturn = True
        KeyAscii = 0
        If cbo执行人(Index).Text <> "" Then
            Set rsTmp = GetDataToPersons(cbo执行人(Index).Text)
            If Not rsTmp Is Nothing Then
                If rsTmp.State = adStateOpen Then
                    blnOk = Not rsTmp.EOF
                End If
            End If
            If blnOk Then
                Call FindCboIndex(cbo执行人(Index), rsTmp!id)
            Else
                cbo执行人(Index).ListIndex = Val(cbo执行人(Index).Tag)
            End If
            Call gobjControl.CboSetIndex(cbo执行人(Index).hWnd, Val(cbo执行人(Index).ListIndex))
        Else
            cbo执行人(Index).ListIndex = Val(cbo执行人(Index).Tag)
        End If
        If strPersons = cbo执行人(Index).Text Then
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo执行人_Validate(Index As Integer, Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call gobjControl.CboSetIndex(cbo执行人(Index).hWnd, Val(cbo执行人(Index).Tag))
    End If
End Sub

Private Sub CMDcancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdDate_Click(Index As Integer)
    Dim lngIndex  As Long
    Dim objControl As Object
    Dim lngLeft As Long, lngTop As Long
    
    If Index = 0 Then
        Set objControl = txt执行时间(0)
        lngLeft = fraExe(0).Left
        lngTop = fraExe(0).Top
    ElseIf Index = 1 Then
        Set objControl = txt执行时间(1)
        lngLeft = fraExe(0).Left
        lngTop = fraExe(0).Top
    ElseIf Index = 2 Then
        Set objControl = txt反应时间(0)
        lngLeft = fraExe(1).Left
        lngTop = fraExe(1).Top
    ElseIf Index = 3 Then
        Set objControl = txt反应时间(1)
        lngLeft = fraExe(1).Left
        lngTop = fraExe(1).Top
    ElseIf Index = 4 Then
        Set objControl = txt反应时间(2)
        lngLeft = fraExe(3).Left
        lngTop = fraExe(3).Top
    End If
    If IsDate(objControl.Text) Then
        dtpDate.Value = CDate(objControl.Text)
    Else
        dtpDate.Value = gobjDatabase.Currentdate
    End If
    dtpDate.Tag = Index
    If lngLeft + objControl.Left + objControl.Width - dtpDate.Width < 0 Then
        dtpDate.Left = lngLeft + objControl.Left
    Else
        dtpDate.Left = lngLeft + objControl.Left + objControl.Width - dtpDate.Width
    End If
    If lngTop + objControl.Top + objControl.Height + dtpDate.Height > Me.Height Then
        dtpDate.Top = lngTop + objControl.Top - dtpDate.Height
    Else
        dtpDate.Top = lngTop + objControl.Top + objControl.Height
    End If
    dtpDate.Visible = True
    dtpDate.ZOrder 0
    dtpDate.SetFocus
End Sub

Private Sub CMDok_Click()
    Dim dbl本次数次 As Double, dbl剩余次数 As Double
    Dim blnTrans As Boolean, strsql As String
    Dim i As Integer, intIndex As Integer
    Dim str输血部位有无渗漏 As String
    Dim str开始执行时间 As String
    Dim rsTmp As New Recordset
    Dim arrMsg As Variant
    
    If lblLink.Tag <> "已核对" Then
        MsgBox "请先进行输注前核对！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If IsDate(txt执行时间(0).Text) = False Then
        MsgBox "输注开始时间不是有效的日期格式！", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txt执行时间(0))
        Exit Sub
    End If
    
    If Trim(cbo执行人(0).Text) = "" Then
        MsgBox "请选择输注开始执行人。", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(cbo执行人(0))
        Exit Sub
    End If
    
    '117041:开始时间相同，则时间相比最后一次时间自动加一秒
    str开始执行时间 = txt执行时间(0).Text
    '检查本次执行时间是否大于上次执行时间
    If IsDate(txt执行时间(0).Tag) Then
        If CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm")) < CDate(Format(txt执行时间(0).Tag, "yyyy-MM-dd HH:mm")) Then
            MsgBox "本次执行时间不能小于上次执行时间 " & Format(txt执行时间(0).Tag, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txt执行时间(0))
            Exit Sub
        ElseIf CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm")) = CDate(Format(txt执行时间(0).Tag, "yyyy-MM-dd HH:mm")) Then
            str开始执行时间 = Format(DateAdd("s", 1, CDate(Format(txt执行时间(0).Tag, "yyyy-MM-dd HH:mm:ss"))), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
     If CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm")) < CDate(Format(txtCheck(2).Text, "yyyy-MM-dd HH:mm")) Then
        MsgBox "本次执行时间不能小于核对时间。", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txt执行时间(0))
        Exit Sub
    End If
    
    If cbo滴数(0).ListIndex <= 0 Then
        If LenB(StrConv(cbo滴数(0).Text, vbFromUnicode)) > 3 Or (Not IsNumeric(cbo滴数(0).Text) And cbo滴数(0).Text <> "") Then
            MsgBox "自由录入的前15分钟滴数只能是数字，且最多只允许录入3位数字！", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(cbo滴数(0))
            Exit Sub
        End If
    End If
    
    '输注过程中的校验
    If cbo滴数(1).ListIndex <= 0 Then
        If LenB(StrConv(cbo滴数(1).Text, vbFromUnicode)) > 3 Or (Not IsNumeric(cbo滴数(1).Text) And cbo滴数(1).Text <> "") Then
            MsgBox "自由录入的15分钟后滴数只能是数字，且最多只允许录入3位数字！", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(cbo滴数(1))
            Exit Sub
        End If
    End If
    
    If (cbo滴数(1).Text <> "" Or optLeakage(1).Value = True Or optBegin(5).Value = True) And Trim(cbo执行人(2).Text) = "" Then
        MsgBox "请选择输注过程执行人。", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(cbo执行人(2))
        Exit Sub
    End If
    
    '输注结束的校验
    If Not IsDate(txt执行时间(1).Text) And txt执行时间(1).Text <> "____-__-__ __:__" Then
        MsgBox "填写了输注结束时间时，必须是有效的日期格式！", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txt执行时间(1))
        Exit Sub
    End If
    
    If IsDate(txt执行时间(1).Text) And Trim(cbo执行人(1).Text) = "" Then
        MsgBox "请选择输注结束执行人。", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(cbo执行人(1))
        Exit Sub
    End If
    
    If optLeakage(3).Value = True And Not IsDate(txt执行时间(1).Text) Then
        MsgBox "输注结束后的输血部位有渗漏时，则必须填写输注结束时间！", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txt执行时间(1))
        Exit Sub
    End If
    '输注结束选择输血反应则必须填写结束时间
    If optBegin(3).Value = True And Not IsDate(txt执行时间(1).Text) Then
        MsgBox "输注结束后的存在输血反应时，则必须填写输注结束时间！", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txt执行时间(1))
        Exit Sub
    End If
    '输血反应填写统一校验
    For i = 1 To 5 Step 2
        If optBegin(i).Value = True Then  '录入输入反应
            intIndex = i \ 2
            If cboReaction(intIndex).ListIndex = -1 Then
                MsgBox "有输血反应时，必须录入不良反应情况！", vbInformation, gstrSysName
                Call gobjControl.ControlSetFocus(cboReaction(intIndex))
                Exit Sub
            End If
            If Not IsDate(txt反应时间(intIndex).Text) Then
                MsgBox "输血反应时间不是有效的日期格式！", vbInformation, gstrSysName
                Call gobjControl.ControlSetFocus(txt反应时间(intIndex))
                Exit Sub
            End If
            
            If CDate(Format(txt反应时间(intIndex).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm")) Then
                MsgBox "输血反应时间不能小于输注开始时间！", vbInformation, gstrSysName
                Call gobjControl.ControlSetFocus(txt反应时间(intIndex))
                Exit Sub
            End If
            '输注过程中的输血反应时间必须小于输注结束时间
            If intIndex = 2 And IsDate(txt执行时间(1).Text) Then
                If CDate(Format(txt反应时间(intIndex).Text, "YYYY-MM-DD HH:mm")) >= CDate(Format(txt执行时间(1).Text, "YYYY-MM-DD HH:mm")) Then
                    MsgBox "输注过程中的输血反应时间必须小于输注结束时间！", vbInformation, gstrSysName
                    Call gobjControl.ControlSetFocus(txt反应时间(intIndex))
                    Exit Sub
                End If
            End If
            '输注结束的输血反应时候必须大于等于输注结束时间
            If intIndex = 1 And IsDate(txt执行时间(1).Text) Then
                If CDate(Format(txt反应时间(intIndex).Text, "YYYY-MM-DD HH:mm")) < CDate(Format(txt执行时间(1).Text, "YYYY-MM-DD HH:mm")) Then
                    MsgBox "输注结束的输血反应时间不能小于输注结束时间！", vbInformation, gstrSysName
                    Call gobjControl.ControlSetFocus(txt反应时间(intIndex))
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    If optBegin(1).Value = True And optBegin(5).Value = True Then
        If CDate(Format(txt反应时间(2).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(txt反应时间(0).Text, "YYYY-MM-DD HH:mm")) Then
            MsgBox "输注过程中的输血反应时间必须大于输注开始的输血反应时间！", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txt反应时间(2))
            Exit Sub
        End If
    End If
    
    If optBegin(1).Value = True And optBegin(3).Value = True Then
        If CDate(Format(txt反应时间(1).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(txt反应时间(0).Text, "YYYY-MM-DD HH:mm")) Then
            MsgBox "输注结束的输血反应时间必须大于输注开始的输血反应时间！", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txt反应时间(1))
            Exit Sub
        End If
    End If
    
    If optBegin(3).Value = True And optBegin(5).Value = True Then
        If CDate(Format(txt反应时间(1).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(txt反应时间(2).Text, "YYYY-MM-DD HH:mm")) Then
            MsgBox "输注结束的输血反应时间必须大于输注过程中的输血反应时间！", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txt反应时间(1))
            Exit Sub
        End If
    End If
    
    If IsDate(txt执行时间(1).Text) Then
         If CDate(Format(txt执行时间(1).Text, "YYYY-MM-DD HH:mm")) <= CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm")) Then
            MsgBox "输注结束时间必须大于输注开始时间！", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txt执行时间(1))
            Exit Sub
         End If
    End If
    
    If gobjCommFun.ActualLen(txt执行摘要.Text) > txt执行摘要.MaxLength Then
        MsgBox "执行摘要内容过多，最多允许 " & txt执行摘要.MaxLength \ 2 & " 个汉字或 " & txt执行摘要.MaxLength & " 个字符。", vbInformation, gstrSysName
        Call gobjControl.ControlSetFocus(txt执行摘要)
        Exit Sub
    End If
    dbl本次数次 = Val(txt本次数次.Text)
    dbl剩余次数 = gobjComlib.FormatEx(Val(txt发送数次.Text) - Val(txt发送数次.Tag), 5)
    If mint血袋数 > mint已执行血袋数 Then
        dbl本次数次 = gobjComlib.FormatEx(dbl剩余次数 / (mint血袋数 - mint已执行血袋数), 5)
    Else
        dbl本次数次 = gobjComlib.FormatEx(dbl本次数次 / mint血袋数, 5)
    End If
    If Val(txt发送数次.Tag) + dbl本次数次 > Val(txt发送数次.Text) Then
        dbl本次数次 = gobjComlib.FormatEx(Val(txt发送数次.Text) - Val(txt发送数次.Tag), 5)
    End If
    
    '内容赋值
    With mExeInfoSave
        .收发ID = mlng收发ID
        .开始执行人 = Trim(gobjCommFun.GetNeedName(cbo执行人(0).Text))
        .开始时间 = Format(str开始执行时间, "yyyy-MM-dd HH:mm:ss")
        If cbo滴数(0).ListIndex > 0 Then
            .前15分钟滴速 = cbo滴数(0).ItemData(cbo滴数(0).ListIndex)
        Else
            .前15分钟滴速 = Val(cbo滴数(0).Text)
        End If
        .输注前输血反应 = IIf(optBegin(0).Value = True, "", cboReaction(0).Text)
        .输注前反应时间 = IIf(optBegin(0).Value = True, "", Format(txt反应时间(0).Text, "yyyy-MM-dd HH:mm:ss"))
        .输注中执行人 = Trim(gobjCommFun.GetNeedName(cbo执行人(2).Text))
        If .输注中执行人 = "" Then
            .后15分钟滴速 = 0
            .输注中输血反应 = ""
            .输注中反应时间 = ""
            str输血部位有无渗漏 = "0"
        Else
            If cbo滴数(1).ListIndex > 0 Then
                .后15分钟滴速 = cbo滴数(1).ItemData(cbo滴数(1).ListIndex)
            Else
                .后15分钟滴速 = Val(cbo滴数(1).Text)
            End If
            .输注中输血反应 = IIf(optBegin(4).Value = True, "", cboReaction(2).Text)
            .输注中反应时间 = IIf(optBegin(4).Value = True, "", Format(txt反应时间(2).Text, "yyyy-MM-dd HH:mm:ss"))
            str输血部位有无渗漏 = IIf(optLeakage(0).Value = True, 0, 1)
        End If
        If Not IsDate(txt执行时间(1).Text) Then
            .输注后输血反应 = ""
            .输注后反应时间 = ""
            .结束执行人 = ""
            .结束时间 = ""
        Else
            .输注后输血反应 = IIf(optBegin(2).Value = True, "", cboReaction(1).Text)
            .输注后反应时间 = IIf(optBegin(2).Value = True, "", Format(txt反应时间(1).Text, "yyyy-MM-dd HH:mm:ss"))
            .结束执行人 = Trim(gobjCommFun.GetNeedName(cbo执行人(1).Text))
            .结束时间 = Format(txt执行时间(1).Text, "yyyy-MM-dd HH:mm:ss")
            str输血部位有无渗漏 = str输血部位有无渗漏 & IIf(optLeakage(2).Value = True, 0, 1)
        End If
        If .输注中执行人 = "" And Not IsDate(txt执行时间(1).Text) Then
            .输血部位渗漏 = ""
        Else
            .输血部位渗漏 = str输血部位有无渗漏
        End If
        
        .核查者 = txtCheck(0).Text
        .复查者 = txtCheck(1).Text
        .核对时间 = txtCheck(2).Text
        .执行科室ID = mlng科室ID
        .登记人 = UserInfo.姓名
        .登记时间 = ""
        .摘要 = txt执行摘要.Text
    End With
    
    Call SetMessages(arrMsg)
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
    If mExeInfo.开始时间 = "" Then
'        If mlng执行科室ID <> mlng科室ID Then
'            strsql = "Zl_病人医嘱发送_科室变更(" & mlng医嘱ID & "," & mlng发送号 & "," & mlng科室ID & ")"
'            Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
'        End If
        
        strsql = "ZL_病人医嘱执行_Insert(" & mlng医嘱ID & "," & mlng发送号 & "," & _
            "To_Date('" & Format(dtp要求时间.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            dbl本次数次 & ",'" & txt执行摘要.Text & "','" & gobjCommFun.GetNeedName(cbo执行人(0).Text) & "'," & _
            "To_Date('" & Format(str开始执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            1 & "," & "0," & 1 & ",'','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
    Else
        strsql = "ZL_病人医嘱执行_Update(To_Date('" & mExeInfo.开始时间 & "','YYYY-MM-DD HH24:MI:SS')," & mlng医嘱ID & "," & mlng发送号 & "," & _
            "To_Date('" & Format(dtp要求时间.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            dbl本次数次 & ",'" & txt执行摘要.Text & "','" & gobjCommFun.GetNeedName(cbo执行人(0).Text) & "'," & _
            "To_Date('" & Format(str开始执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & "," & 1 & ",NULL," & 1 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
    End If
    
     '新增时的核对
    If lblLink.Enabled = True Then
        strsql = "Zl_血液执行记录_Check(" & mlng收发ID & ",'" & mExeInfoSave.核查者 & "','" & mExeInfoSave.复查者 & "',To_Date('" & mExeInfoSave.核对时间 & "','YYYY-MM-DD HH24:MI:SS'),'" & imgMore.Tag & "')"
        Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
    End If
    
    strsql = "zl_血液执行记录_Update(" & mlng收发ID & ",'" & mExeInfoSave.开始执行人 & "',To_Date('" & mExeInfoSave.开始时间 & "','YYYY-MM-DD HH24:MI:SS')," & _
        IIf(mExeInfoSave.前15分钟滴速 = 0, "NULL", mExeInfoSave.前15分钟滴速) & ",'" & mExeInfoSave.输注前输血反应 & "'," & IIf(mExeInfoSave.输注前反应时间 = "", "NULL", "To_Date('" & mExeInfoSave.输注前反应时间 & "','YYYY-MM-DD HH24:MI:SS')") & ",'" & mExeInfoSave.输注中执行人 & "'," & _
        IIf(mExeInfoSave.后15分钟滴速 = 0, "NULL", mExeInfoSave.后15分钟滴速) & ",'" & mExeInfoSave.输注中输血反应 & "'," & IIf(mExeInfoSave.输注中反应时间 = "", "NULL", "To_Date('" & mExeInfoSave.输注中反应时间 & "','YYYY-MM-DD HH24:MI:SS')") & ",'" & _
        mExeInfoSave.结束执行人 & "'," & IIf(mExeInfoSave.结束时间 = "", "NULL", "To_Date('" & mExeInfoSave.结束时间 & "','YYYY-MM-DD HH24:MI:SS')") & "," & _
        "'" & mExeInfoSave.输注后输血反应 & "'," & IIf(mExeInfoSave.输注后反应时间 = "", "NULL", "To_Date('" & mExeInfoSave.输注后反应时间 & "','YYYY-MM-DD HH24:MI:SS')") & "," & _
        mExeInfoSave.执行科室ID & ",'" & mExeInfoSave.输血部位渗漏 & "','" & mExeInfoSave.登记人 & "'," & IIf(mExeInfoSave.登记时间 = "", "NULL", "To_Date('" & mExeInfoSave.登记时间 & "','YYYY-MM-DD HH24:MI:SS')") & ",'" & mExeInfoSave.摘要 & "')"
    Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
    
    '保存生命体征数据
    For i = 0 To 2
        strsql = UCPatiVS(i).GetSaveSQL(mlng收发ID, i + 1)
        Call gobjDatabase.ExecuteProcedure(strsql, Me.Caption)
    Next
    For i = 0 To UBound(arrMsg)
        Call gobjDatabase.ExecuteProcedure(CStr(arrMsg(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    mblnOk = True
    Unload Me
    Exit Sub
errH:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SetMessages(ByRef arrSQL As Variant) As Boolean
    Dim rsTmp As New Recordset
    Dim strSQL As String
    Dim lng病人ID As Long, lng科室ID As Long, lng病区ID As Long, lng执行科室id As Long
    Dim lng就诊id As Long
    Dim int病人来源 As Integer
    Dim str提醒部门 As String
    arrSQL = Array()
    strSQL = "select a.主页id,a.挂号单,a.病人id,a.病人科室id,a.病人来源, b.执行部门id from 病人医嘱记录 a,血液配血记录 b,血液发送记录 c " & _
            "WHERE a.id = b.申请id AND b.id = c.配发id AND a.id = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng相关ID)
    If rsTmp.State = adStateClosed Then Exit Function
    If rsTmp.RecordCount = 0 Then Exit Function
    lng病人ID = Val(rsTmp!病人id)
    lng科室ID = Val(rsTmp!病人科室id)
    int病人来源 = Val(rsTmp!病人来源)
    lng执行科室id = Val(rsTmp!执行部门id)
    If int病人来源 = 2 Then
        lng就诊id = Val(rsTmp!主页id)
        strSQL = "select 当前病区id from 病案主页 where 病人id = [1] and 主页id = [2]  "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng就诊id)
        lng病区ID = Val(rsTmp!当前病区id)
    Else
        lng病区ID = Val(rsTmp!病人科室id)
        strSQL = "select id 挂号id from 病人挂号记录 where no = [1] and 病人id = [2] "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTmp!挂号单 & "", Val(rsTmp!病人id))
        lng就诊id = Val(rsTmp!挂号ID)
    End If
    strSQL = "select ID,类型编码,业务标识 from 业务消息清单 where 病人ID = [1] and 就诊id = [2] and 是否已阅 = 0 "
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng就诊id)
    '确定要提醒的部门
    str提醒部门 = IIf(Val(lng科室ID) = 0, "", lng科室ID)
    If lng病区ID <> lng科室ID Then
        If str提醒部门 = "" Then
            str提醒部门 = IIf(lng病区ID = 0, "", lng病区ID)
        Else
            str提醒部门 = str提醒部门 & IIf(lng病区ID = 0, "", "," & lng病区ID)
        End If
    End If
    '查询是否存在本医嘱本血袋输血反应消息
    rsTmp.Filter = "类型编码 = 'ZLHIS_BLOOD_006' And 业务标识 = '" & mlng相关ID & ":" & mlng收发ID & "'"
    If (optBegin(1).Value Or optBegin(3).Value Or optBegin(5).Value) Then
        If rsTmp.RecordCount = 0 Then
            strSQL = "Zl_业务消息清单_Insert(" & lng病人ID & "," & lng就诊id & ","  '病人id 就诊id
            strSQL = strSQL & Val(lng科室ID) & ","     '就诊科室id
            strSQL = strSQL & Val(lng病区ID) & ","      '就诊病区id
            strSQL = strSQL & int病人来源 & ","                                      '病人来源
            strSQL = strSQL & "'出现输血反应，请及时填写输血反应单。','"             '消息内容
            strSQL = strSQL & IIf(Val(int病人来源) = 1, "1000", "0100") & "','ZLHIS_BLOOD_006',"     ' 提醒场合, 类型编码
            strSQL = strSQL & "'" & mlng相关ID & ":" & mlng收发ID & "',"                      '业务标识（相关id:收发id）
            strSQL = strSQL & "1,0,NULL,'" & str提醒部门 & "',NULL)"                                                   '优先程度，是否已阅，登记时间,提醒部门
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    Else    '无输血反应，则查询是否存在有反应消息，若有，设为已读。
        If rsTmp.RecordCount > 0 Then
            strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊id & ",'ZLHIS_BLOOD_006',"
            strSQL = strSQL & "3,'" & UserInfo.姓名 & "'," & lng病区ID & ",NULL,"
            strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    End If
    
    rsTmp.Filter = "类型编码 = 'ZLHIS_BLOOD_007' And 业务标识 = '" & mlng相关ID & ":" & mlng医嘱ID & ":" & mlng收发ID & "'"
    If IsDate(txt执行时间(1).Text) Then
        If rsTmp.RecordCount = 0 Then
            strSQL = "Zl_业务消息清单_Insert(" & lng病人ID & "," & lng就诊id & ","  '病人id 就诊id
            strSQL = strSQL & Val(lng科室ID) & ","      '就诊科室id
            strSQL = strSQL & Val(lng病区ID) & ","      '就诊病区id
            strSQL = strSQL & int病人来源 & ","                                      '病人来源
            strSQL = strSQL & "'输血完成，请在24小时内收回血袋。','"                         '消息内容
            strSQL = strSQL & IIf(Val(int病人来源) = 1, "0001", "0010") & "','ZLHIS_BLOOD_007',"     ' 提醒场合, 类型编码
            strSQL = strSQL & "'" & mlng相关ID & ":" & mlng医嘱ID & ":" & mlng收发ID & "',"                      '业务标识（相关id:收发id）
            strSQL = strSQL & "1,0,NULL,'" & IIf(Val(int病人来源) = 1, lng执行科室ID, lng病区ID)  & "',NULL)"                                                   '优先程度，是否已阅，登记时间,提醒部门                                                      '
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    Else    '无血袋需要回收，则查询是否存在血袋消息，若有，设为已读。
        If rsTmp.RecordCount > 0 Then
            strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊id & ",'ZLHIS_BLOOD_007',"
            strSQL = strSQL & IIf(Val(int病人来源) = 1, 4, 3) & ",'" & UserInfo.姓名 & "'," & lng病区ID & ",NULL,"
            strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    End If
End Function
Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String, intIndex As Integer
    Dim objControl As Object
    
    intIndex = Val(dtpDate.Tag)
    If intIndex = 0 Then
        Set objControl = txt执行时间(0)
    ElseIf intIndex = 1 Then
        Set objControl = txt执行时间(1)
    ElseIf intIndex = 2 Then
        Set objControl = txt反应时间(0)
    ElseIf intIndex = 3 Then
        Set objControl = txt反应时间(1)
    ElseIf intIndex = 4 Then
        Set objControl = txt反应时间(2)
    Else
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Exit Sub
    End If
    
    '取值
    If IsDate(objControl.Text) Then
        strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(objControl.Text, "yyyy-MM-dd HH:mm"), 12, 5)
    Else
        strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
    End If
    objControl.Text = strDate
        
    If intIndex = 0 Or intIndex = 1 Then
        Call txt执行时间_Validate(intIndex, False)
    ElseIf intIndex = 2 Or intIndex = 3 Or intIndex = 4 Then
        Set objControl = txt执行时间(1)
        Call txt反应时间_Validate(intIndex - 2, False)
    End If
    dtpDate.Tag = ""
    dtpDate.Visible = False
    Call gobjControl.ControlSetFocus(objControl)
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    Dim intIndex As Integer
    Dim objControl As Object
    If KeyAscii = vbKeyEscape Then
        intIndex = Val(dtpDate.Tag)
        If intIndex = 0 Then
            Set objControl = txt执行时间(0)
        ElseIf intIndex = 1 Then
            Set objControl = txt执行时间(1)
        ElseIf intIndex = 2 Then
            Set objControl = txt反应时间(0)
        ElseIf intIndex = 3 Then
            Set objControl = txt反应时间(1)
        Else
            dtpDate.Tag = ""
            dtpDate.Visible = False
            Exit Sub
        End If
        
        Call gobjControl.ControlSetFocus(objControl)
        dtpDate.Tag = ""
        dtpDate.Visible = False
    End If
End Sub

Private Sub Form_Activate()
    Dim lngScrH As Long
    If mblnAcTive = False Then Exit Sub
    mblnAcTive = False
    '显示窗体
    lngScrH = GetSystemMetrics(17) * 15 '屏幕可用高度
    If Me.Top + Me.Height > lngScrH Then
        Me.Top = lngScrH - Me.Height
    End If
    If mblnOnlyRead = True Then
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.name <> "cbo执行人" And Me.ActiveControl.name <> "UCPatiVS" Then
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String
    Dim i As Integer
    Dim lngFindEndIndex As Long, lngFindCenterIndex As Long
    Dim lng最大已销 As Long
    
    On Error GoTo ErrHand
    mintTimerCount = 0
    mblnAcTive = True
    Call ClearExeInfo
    strsql = "Select 接收时间,接收状态 From 血液发送记录 where 收发ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng收发ID)
    If rsTmp.EOF Then
        MsgBox "血液还未发送,不能进行执行登记！", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    If gbln接收后才能执行 = True And mblnOnlyRead = False Then
        If Not (Val("" & rsTmp!接收状态) = 1 Or Val("" & rsTmp!接收状态) = 3) Then
            MsgBox "血液还未接收,不能进行执行登记！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    If IsDate("" & rsTmp!接收时间) Then
        mstr接收时间 = Format("" & rsTmp!接收时间, "YYYY-MM-DD HH:mm")
    Else
        mstr接收时间 = ""
    End If
    strsql = _
        " Select 收发id, 开始执行人, 开始时间, 前15分钟滴速, 输注前输血反应, 输注前反应时间, 后15分钟滴速,输注中执行人,输注中输血反应, 输注中反应时间,  " & vbNewLine & _
                "结束执行人, 结束时间, 执行科室id, 输注后输血反应, 输注后反应时间,输血部位渗漏,核查者, 复查者, 核对时间, 登记人,登记时间, 摘要" & vbNewLine & _
        " From 血液执行记录 where 收发ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng收发ID)
    If rsTmp.RecordCount > 0 Then
        With mExeInfo
            .收发ID = Val("" & rsTmp("收发ID"))
            .开始执行人 = "" & rsTmp("开始执行人")
            .开始时间 = "" & rsTmp("开始时间")
            .前15分钟滴速 = Val("" & rsTmp("前15分钟滴速"))
            .输注前输血反应 = "" & rsTmp("输注前输血反应")
            .输注前反应时间 = "" & rsTmp("输注前反应时间")
            .输注中执行人 = "" & rsTmp("输注中执行人")
            .后15分钟滴速 = Val("" & rsTmp("后15分钟滴速"))
            .输注中输血反应 = "" & rsTmp("输注中输血反应")
            .输注中反应时间 = "" & rsTmp("输注中反应时间")
            .结束执行人 = "" & rsTmp("结束执行人")
            .结束时间 = "" & rsTmp("结束时间")
            .输注后输血反应 = "" & rsTmp("输注后输血反应")
            .输注后反应时间 = "" & rsTmp("输注后反应时间")
            .执行科室ID = Val("" & rsTmp("执行科室ID"))
            .核查者 = "" & rsTmp("核查者")
            .复查者 = "" & rsTmp("复查者")
            .核对时间 = "" & rsTmp("核对时间")
            .输血部位渗漏 = "" & rsTmp("输血部位渗漏")
            .登记人 = "" & rsTmp("登记人")
            .登记时间 = "" & rsTmp("登记时间")
            .摘要 = "" & rsTmp("摘要")
        End With
    End If
    mstr缺省输血反应 = ""
    
    '滴数
    With cbo滴数(0)
        .Clear
        .AddItem 15: .ItemData(.NewIndex) = 15
        .AddItem 30: .ItemData(.NewIndex) = 30
        .AddItem "快速": .ItemData(.NewIndex) = -1
        .AddItem "加压": .ItemData(.NewIndex) = -2
        .ListIndex = 0
    End With
    With cbo滴数(1)
        .Clear
        .AddItem 15: .ItemData(.NewIndex) = 15
        .AddItem 30: .ItemData(.NewIndex) = 30
        .AddItem "快速": .ItemData(.NewIndex) = -1
        .AddItem "加压": .ItemData(.NewIndex) = -2
    End With
    '输血反应读取
    strsql = "Select 名称,缺省标志 From 输血反应"
    Call gobjDatabase.OpenRecordset(rsTmp, strsql, "输血反应")
    cboReaction(0).Clear
    cboReaction(1).Clear
    cboReaction(2).Clear
    Do While Not rsTmp.EOF
        cboReaction(0).AddItem "" & rsTmp!名称
        cboReaction(1).AddItem "" & rsTmp!名称
        cboReaction(2).AddItem "" & rsTmp!名称
        If Val(rsTmp!缺省标志) = 1 Then mstr缺省输血反应 = "" & rsTmp!名称
    rsTmp.MoveNext
    Loop
    '执行人读取
    cbo执行人(0).Clear: cbo执行人(0).Tag = -1
    cbo执行人(1).Clear: cbo执行人(1).Tag = -1
    cbo执行人(2).Clear: cbo执行人(2).Tag = -1
    cbo执行人(2).AddItem ""
    cbo执行人(2).ItemData(cbo执行人(2).NewIndex) = 0
    '读取执行人(本科人员)
    Set rsTmp = GetDataToPersons
    For i = 1 To rsTmp.RecordCount
        cbo执行人(0).AddItem rsTmp!编号 & "-" & rsTmp!姓名
        cbo执行人(0).ItemData(cbo执行人(0).NewIndex) = Val("" & rsTmp!id)
        cbo执行人(1).AddItem rsTmp!编号 & "-" & rsTmp!姓名
        cbo执行人(1).ItemData(cbo执行人(1).NewIndex) = Val("" & rsTmp!id)
        cbo执行人(2).AddItem rsTmp!编号 & "-" & rsTmp!姓名
        cbo执行人(2).ItemData(cbo执行人(2).NewIndex) = Val("" & rsTmp!id)
        If mExeInfo.开始时间 = "" Then
            If rsTmp!id = UserInfo.id Then
                cbo执行人(0).ListIndex = cbo执行人(0).NewIndex
            End If
        Else
            If rsTmp!姓名 = mExeInfo.开始执行人 Then
                cbo执行人(0).ListIndex = cbo执行人(0).NewIndex
            End If
        End If
        If mExeInfo.输注中执行人 = "" Then
            If rsTmp!id = UserInfo.id Then
                lngFindCenterIndex = cbo执行人(2).NewIndex
            End If
        Else
            If rsTmp!姓名 = mExeInfo.输注中执行人 Then
                cbo执行人(2).ListIndex = cbo执行人(2).NewIndex
            End If
        End If
        
        If mExeInfo.结束时间 = "" Then
            If rsTmp!id = UserInfo.id Then
                lngFindEndIndex = cbo执行人(1).NewIndex
            End If
        Else
            If rsTmp!姓名 = mExeInfo.结束执行人 Then
                cbo执行人(1).ListIndex = cbo执行人(1).NewIndex
            End If
        End If
        rsTmp.MoveNext
    Next
    
    If mlngModul = p医技工作站 Then
        If Val(gobjDatabase.GetPara(51, 100)) = 1 Then
            Me.cbo执行人(0).Enabled = False
            If cbo执行人(1).ListCount > 0 And cbo执行人(1).ListIndex = -1 Then cbo执行人(1).ListIndex = lngFindEndIndex
            Me.cbo执行人(1).Enabled = False
            If cbo执行人(2).ListCount > 0 And cbo执行人(2).ListIndex = -1 Then cbo执行人(2).ListIndex = lngFindCenterIndex
            Me.cbo执行人(2).Enabled = False
        End If
    End If
    
    If mExeInfo.核对时间 <> "" Then
        txtCheck(0).Text = mExeInfo.核查者
        txtCheck(1).Text = mExeInfo.复查者
        txtCheck(2).Text = Format(mExeInfo.核对时间, "YYYY-MM-DD HH:mm")
        lblLink.Enabled = False
        lblLink.Tag = "已核对"
    Else
        txtCheck(0).Text = ""
        txtCheck(1).Text = ""
        txtCheck(2).Text = ""
        lblLink.Enabled = True
        lblLink.Tag = ""
    End If
    '执行情况读取
    mint血袋数 = 0
    mint已执行血袋数 = 0
    txt发送数次.Tag = ""
    txt执行时间(0).Text = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    With mExeInfo
        If mExeInfo.开始时间 = "" Then '新增
            '获取上次执行信息
            strsql = _
                " Select Max(执行时间) as LastDate,Sum(本次数次) as curNum" & _
                " From 病人医嘱执行" & _
                " Where 医嘱ID=[1] And 发送号=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng医嘱ID, mlng发送号)
            If Not rsTmp.EOF Then
                txt执行时间(0).Tag = Format(Nvl(rsTmp!LastDate), "yyyy-MM-dd HH:mm:ss") '上次实际执行时间
                txt发送数次.Tag = Nvl(rsTmp!curNum, 0) '血液医嘱的执行次数总和，每次执行一袋血，执行次数为1
            End If
            
            '计算本次执行应该的要求时间
            strsql = "Select A.发送数次,Nvl(B.相关id, B.ID) 组ID,C.计算单位,A.首次时间,A.末次时间,Decode(B.病人来源, 2, Decode(A.记录性质, 1, 1, Decode(A.门诊记帐, 1, 1, 2)), 1) 费用性质," & _
                " B.开始执行时间,Decode(B.医嘱期效,0,B.执行终止时间,null) as 执行终止时间,B.上次执行时间,B.执行时间方案," & _
                " B.执行频次,B.频率次数,B.频率间隔,B.间隔单位,B.病人ID,b    .主页ID,c.类别,c.操作类型,c.执行分类,C.计算方式,B.医嘱期效,Nvl(b.总给予量, 1) as 总给予量,NVL(B.单次用量,1) AS 单次用量,A.NO " & _
                " From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C" & _
                " Where A.医嘱ID=B.ID And B.诊疗项目ID=C.ID" & _
                " And A.医嘱ID=[1] And A.发送号=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng医嘱ID, mlng发送号)
            dtp要求时间.Value = rsTmp!开始执行时间 '输血医嘱都为一次性执行的临嘱
            txt发送数次.Text = Val(rsTmp!发送数次 & "")
            mlng相关ID = rsTmp!组ID
            '查询单据中最大的已经退费或销帐的医嘱执行次数
            lng最大已销 = Get最大已销(mlng医嘱ID, rsTmp!NO & "", rsTmp!类别 & "", Val(rsTmp!费用性质 & ""))
            mint血袋数 = GetBloodNum
            mint已执行血袋数 = gobjComlib.FormatEx(Val(txt发送数次.Tag) * mint血袋数 / Val(txt发送数次.Text), 0) '上次执行的血袋数。已经有5位小数，用四舍五入就可以满足
            If lng最大已销 = 1 Then '对于输血医嘱 这个 mlng最大已销 的值只能是 0 或 1 因为只能销一次。
                MsgBox "该医嘱相关单据已经退费或销帐，不能再执行。", vbInformation, gstrSysName
                Unload Me: Exit Sub
            ElseIf lng最大已销 = 0 Then
                If mint已执行血袋数 >= mint血袋数 Then
                    MsgBox "该医嘱本次发送允许执行 " & mint血袋数 & "袋，当前已经执行了 " & mint已执行血袋数 & " 袋，不能再执行。", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
            txt本次数次.Text = 1 '每次执行默认为一袋
        Else '修改
            '上次已执行到的一些数据(不算本次)
            strsql = "Select " & _
                " Max(执行时间) as LastDate," & _
                " Sum(本次数次) as curNum" & _
                " From 病人医嘱执行" & _
                " Where 执行时间<[3] And 医嘱ID=[1] And 发送号=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng医嘱ID, mlng发送号, CDate(mExeInfo.开始时间))
            If Not rsTmp.EOF Then
                txt发送数次.Tag = Nvl(rsTmp!curNum, 0) '上次为止实际已执行的数次总量
                txt执行时间(0).Tag = Format(Nvl(rsTmp!LastDate), "yyyy-MM-dd HH:mm:ss")   '上次实际执行时间
            End If
            
            strsql = "Select A.要求时间,Nvl(C.相关id, C.ID) 组ID,A.执行时间,A.本次数次,A.执行摘要,A.执行结果,A.执行人,B.发送数次,Decode(C.病人来源, 2, Decode(B.记录性质, 1, 1, Decode(B.门诊记帐, 1, 1, 2)), 1) 费用性质,D.计算单位,Decode(c.医嘱期效,0,c.执行终止时间,null) as 执行终止时间 ,d.类别,d.操作类型,d.执行分类,c.病人ID,c.主页ID,B.NO" & _
                " From 病人医嘱执行 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 D" & _
                " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And B.医嘱ID=C.ID And C.诊疗项目ID=D.ID" & _
                " And A.医嘱ID=[1] And A.发送号=[2] And A.执行时间=[3]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng医嘱ID, mlng发送号, CDate(mExeInfo.开始时间))
            dtp要求时间.Value = rsTmp!要求时间
            txt本次数次.Text = gobjComlib.FormatEx(Nvl(rsTmp!本次数次), 5)
            txt执行摘要.Text = "" & rsTmp!执行摘要
            txt执行时间(0).Text = Format(rsTmp!执行时间, "YYYY-MM-DD HH:mm")
            txt发送数次.Text = Val(rsTmp!发送数次 & "")
            mlng相关ID = rsTmp!组ID
            gobjComlib.cbo.SetText cbo执行人(0), rsTmp!执行人
            
            mint血袋数 = GetBloodNum
            mint已执行血袋数 = gobjComlib.FormatEx(Val(txt发送数次.Tag) * mint血袋数 / Val(txt发送数次.Text), 0) '本次的执行的血袋数
            txt本次数次.Text = gobjComlib.FormatEx(Val("" & rsTmp!本次数次) * mint血袋数, 0)
            
            Select Case mExeInfo.前15分钟滴速
                Case 0
                    cbo滴数(0).Text = ""
                Case -1
                    cbo滴数(0).ListIndex = 2
                Case -2
                    cbo滴数(0).ListIndex = 3
                Case Else
                    cbo滴数(0).Text = mExeInfo.前15分钟滴速
            End Select
            If mExeInfo.输注前输血反应 <> "" Then
                optBegin(1).Value = True
                Call gobjControl.CboLocate(cboReaction(0), mExeInfo.输注前输血反应)
                txt反应时间(0).Text = Format(mExeInfo.输注前反应时间, "YYYY-MM-DD HH:mm")
            End If
            
            Select Case mExeInfo.后15分钟滴速
                Case 0
                    cbo滴数(1).Text = ""
                Case -1
                    cbo滴数(1).ListIndex = 2
                Case -2
                    cbo滴数(1).ListIndex = 3
                Case Else
                    cbo滴数(1).Text = mExeInfo.后15分钟滴速
            End Select
            gobjComlib.cbo.SetText cbo执行人(2), mExeInfo.输注中执行人
            If Val(Mid(.输血部位渗漏, 1, 1)) = 1 Then optLeakage(1).Value = True
            If mExeInfo.输注中输血反应 <> "" Then
                optBegin(5).Value = True
                Call gobjControl.CboLocate(cboReaction(2), mExeInfo.输注中输血反应)
                txt反应时间(2).Text = Format(mExeInfo.输注中反应时间, "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(Format(mExeInfo.结束时间, "YYYY-MM-DD HH:mm")) Then
                txt执行时间(1).Text = Format(mExeInfo.结束时间, "YYYY-MM-DD HH:mm")
                gobjComlib.cbo.SetText cbo执行人(1), mExeInfo.结束执行人
                If mExeInfo.输注后输血反应 <> "" Then
                    optBegin(3).Value = True
                    Call gobjControl.CboLocate(cboReaction(1), mExeInfo.输注后输血反应)
                    txt反应时间(1).Text = Format(mExeInfo.输注后反应时间, "YYYY-MM-DD HH:mm")
                End If
                If Val(Mid(.输血部位渗漏, 2, 1)) = 1 Then optLeakage(3).Value = True
            End If
        End If
    End With
    cmdDate(0).Tag = txt执行时间(0).Text
    If cbo执行人(0).ListIndex <> -1 Then cbo执行人(0).Tag = cbo执行人(0).ListIndex
    If cbo执行人(1).ListIndex <> -1 Then cbo执行人(1).Tag = cbo执行人(1).ListIndex
    If cbo执行人(2).ListIndex <> -1 Then cbo执行人(2).Tag = cbo执行人(2).ListIndex
    
    Call UCPatiVS(0).LoadPatiVitalSigns(mlng收发ID, 1)  '输血前生命体征
    Call UCPatiVS(1).LoadPatiVitalSigns(mlng收发ID, 2)  '输血中生命体征
    Call UCPatiVS(2).LoadPatiVitalSigns(mlng收发ID, 3)  '输血后生命体征
    
    Call SetFaceEnabledFalse
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub SetFaceEnabledFalse()
'功能：设置控件的可用性
    Dim objControl As Object
    If mblnOnlyRead = True Then
        Err.Clear
        On Error Resume Next
        For Each objControl In Me.Controls
            Select Case TypeName(objControl)
                Case "TextBox", "ComboBox"
                    objControl.locked = True
                    objControl.TabStop = False
                    objControl.BackColor = vbButtonFace
                    If objControl.Text <> "" Then
                        objControl.SelStart = 0: objControl.SelLength = 0
                    End If
                Case "CheckBox", "CommandButton", "OptionButton"
                    objControl.Enabled = False
                Case "MaskEdBox"
                    objControl.Enabled = False
                    objControl.BackColor = vbButtonFace
                    objControl.SelStart = 0: objControl.SelLength = 0
            End Select
        Next
        UCPatiVS(0).ControlLock = True
        UCPatiVS(1).ControlLock = True
        UCPatiVS(2).ControlLock = True
        lblLink.Enabled = False
        cmdCancel.Enabled = True
        If Err <> 0 Then Err.Clear
    End If
End Sub

Private Function GetBloodNum() As Integer
'获取本次医嘱发送的数量
    Dim rsTemp As New ADODB.Recordset
    Dim strsql As String
    On Error GoTo ErrHand
    strsql = "Select Count(收发id)  数量 From 血液发送记录 a, 血液配血记录 b Where a.配发id = b.Id And b.申请id = [1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng相关ID)
    GetBloodNum = rsTemp!数量
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub lblLink_Click()
    Dim blnOk As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strsql As String
    Dim strCheckOper As String, strCheckTime As String, strCheckResult As String
    If mExeInfo.核对时间 <> "" Then
        MsgBox "该袋血液已经核对，不允许再次核对！", vbInformation, gstrSysName
        Exit Sub
    End If
    blnOk = frmUserCheck.ShowMe(Me, mlngModul, mlng科室ID, mlng科室ID, mstr接收时间, "", True, 执行核对)
    If blnOk = True Then
        strCheckOper = frmUserCheck.SendAndTakeOper
        strCheckTime = frmUserCheck.SendTime
        strCheckResult = frmUserCheck.CheckResult
        
        txtCheck(0).Text = Split(strCheckOper, "'")(0)
        txtCheck(1).Text = Split(strCheckOper, "'")(1)
        txtCheck(2).Text = strCheckTime
        imgMore.Tag = strCheckResult
        lblLink.Tag = "已核对"
        If IsDate(txt执行时间(0).Text) Then
            If Format(txt执行时间(0).Text, "YYYY-MM-DD HH:mm") < Format(strCheckTime, "YYYY-MM-DD HH:mm") Then
                txt执行时间(0).Text = Format(strCheckTime, "YYYY-MM-DD HH:mm")
            End If
        End If
    End If
End Sub

Private Sub optBegin_Click(Index As Integer)
    Dim intIndex As Integer
    Dim blnEable As Boolean
    
    intIndex = (Index + 1) \ 3
    If optBegin(Index).Value = True And Index Mod 2 = 0 Then
        blnEable = False
    Else
        blnEable = True
    End If
    If blnEable = False Then
        cboReaction(intIndex).ListIndex = -1
    Else
       If cboReaction(intIndex).ListIndex < 0 Then Call gobjControl.CboLocate(cboReaction(intIndex), mstr缺省输血反应)
    End If
    cboReaction(intIndex).Enabled = blnEable
    txt反应时间(intIndex).Enabled = blnEable
    cmdDate(intIndex + 2).Enabled = blnEable
End Sub

Private Sub cbo滴数_GotFocus(Index As Integer)
    Call gobjControl.TxtSelAll(cbo滴数(Index))
End Sub

Private Sub cbo滴数_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    If Chr(KeyAscii) = 0 Then
        If cbo滴数(Index).Text = "" Or cbo滴数(Index).SelLength = Len(cbo滴数(Index).Text) Or cbo滴数(Index).SelStart = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If KeyAscii <> vbKeyReturn And KeyAscii <> 8 Then
        If cbo滴数(Index).SelLength = 0 And Len(cbo滴数(Index).Text) > 2 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub TimeFlash_Timer()
    mintTimerCount = mintTimerCount + 1
    
    If mintTimerCount Mod 2 = 0 Then
        lblTitle.ForeColor = 0
    Else
        lblTitle.ForeColor = 255
    End If
    
    If mintTimerCount = 10 Then mintTimerCount = 0
End Sub

Private Sub txt反应时间_GotFocus(Index As Integer)
    Call gobjComlib.os.OpenImeByName
'    Call gobjControl.TxtSelAll(txt反应时间(Index))
    txt反应时间(Index).SelStart = 0
    txt反应时间(Index).SelLength = Len(txt反应时间(Index).Text)
End Sub

Private Sub txt反应时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If Index = 0 Then
            Call cmdDate_Click(2)
        ElseIf Index = 1 Then
            Call cmdDate_Click(3)
        End If
    End If
End Sub

Private Sub txt反应时间_Validate(Index As Integer, Cancel As Boolean)
    Dim intIndex As Integer
    If Not IsDate(txt反应时间(Index).Text) Then
        If txt反应时间(Index).Text <> "____-__-__ __:__" Then
            Cancel = True
            Call txt反应时间_GotFocus(Index)
            Exit Sub
        Else
            If IsDate(cmdDate(Index + 2).Tag) Then
                txt反应时间(Index).Text = Format(cmdDate(Index + 2).Tag, "YYYY-MM-DD HH:mm")
            End If
        End If
    Else
        txt反应时间(Index).Text = Format(txt反应时间(Index).Text, "YYYY-MM-DD HH:mm")
        cmdDate(Index + 2).Tag = txt反应时间(Index).Text
    End If
End Sub

Private Sub txt执行时间_GotFocus(Index As Integer)
    Call gobjComlib.os.OpenImeByName
    txt执行时间(Index).SelStart = 0: txt执行时间(Index).SelLength = Len(txt执行时间(Index).Text)
End Sub

Private Sub txt执行时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If Index = 0 Then
            Call cmdDate_Click(0)
        ElseIf Index = 1 Then
            Call cmdDate_Click(1)
        End If
    End If
End Sub

Private Sub txt执行时间_Validate(Index As Integer, Cancel As Boolean)
    Dim strDate As String
    If Not IsDate(txt执行时间(Index).Text) Then
        If Index = 0 Then
            If IsDate(cmdDate(0).Tag) Then
                strDate = Format(cmdDate(0).Tag, "YYYY-MM-DD HH:mm")
            Else
                strDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
            End If
        ElseIf Index = 1 Then
            If IsDate(cmdDate(1).Tag) Then
                strDate = Format(cmdDate(1).Tag, "YYYY-MM-DD HH:mm")
            End If
        End If
        If txt执行时间(Index).Text <> "____-__-__ __:__" Then
            Cancel = True
            If IsDate(strDate) Then txt执行时间(Index).Text = strDate
            Call txt执行时间_GotFocus(Index)
            Exit Sub
        Else
            If Index = 0 And IsDate(strDate) Then
                 If IsDate(strDate) Then txt执行时间(Index).Text = strDate
            End If
        End If
    Else
        txt执行时间(Index).Text = Format(txt执行时间(Index).Text, "YYYY-MM-DD HH:mm")
        If Index = 0 Then
            cmdDate(0).Tag = txt执行时间(Index).Text
        ElseIf Index = 1 Then
            cmdDate(1).Tag = txt执行时间(Index).Text
        End If
    End If
End Sub

Private Sub txt执行摘要_GotFocus()
    Call gobjControl.TxtSelAll(txt执行摘要)
End Sub

Private Sub txt执行摘要_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function Get最大已销(ByVal lng医嘱ID As Long, ByVal strNO As String, ByVal str诊疗类别 As String, ByVal int费用性质 As Integer) As Long
'功能：获取某条医嘱，或某组医嘱的最大已销帐的医嘱执行次数
'       lng医嘱ID 该条医嘱ID
'       strNo  费用NO
'       str诊疗类别 该医嘱的诊疗类别
'       int费用性质 1-门诊费用，2-住院费用
    Dim rsTmp As ADODB.Recordset, strsql As String, strTable As String
    strTable = IIf(int费用性质 = 1, "门诊费用记录", "住院费用记录")
    On Error GoTo errH
    
    strsql = "Select -1 * Sum(Nvl(a.付数, 1) * a.数次 / b.数量) As 最大已销数" & vbNewLine & _
            "From " & strTable & " A, 病人医嘱计价 B" & vbNewLine & _
            "Where a.医嘱序号 = [1] And A.NO=[3] And b.医嘱id = a.医嘱序号 And b.收费细目id = a.收费细目id And Nvl(B.费用性质,0)=0 And a.记录状态 = 2 And a.记录性质 in(1,2,11) And a.价格父号 Is Null And" & vbNewLine & _
            "      a.收费类别 Not In ('5', '6', '7') And Not Exists" & vbNewLine & _
            " (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1)"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, lng医嘱ID, str诊疗类别, strNO)
    If rsTmp.RecordCount <> 0 Then
        Get最大已销 = Val(rsTmp!最大已销数 & "")
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function GetDataToPersons(Optional ByVal strIn As String = "") As ADODB.Recordset
'功能相应科室的医护人员信息
    Dim strsql As String, strNewSQL As String, strWhere As String
    Dim blnYn As Boolean
    
    
    On Error GoTo ErrHand
    If strIn <> "" Then blnYn = True
    
    '医技术站，只有执行他可项目才能看到其他科室的医嘱，临床护士站可能会具有全员病人的权限，需要加上操作员本人(能在病区看到病人，要么操作就是该病区的，要么就是具有全员病区权限)
    If InStr(mstrPrivs, "执行他科项目") > 0 Or Not (mlngModul = p医技工作站) Then
        strNewSQL = " Union " & vbNewLine & _
                            " Select " & UserInfo.id & " id,'" & UserInfo.编号 & "' 编号,'" & UserInfo.姓名 & "' 姓名,'" & UserInfo.简码 & "' 简码 From Dual "
    End If
        
    If Not mlngModul = p医技工作站 Then
        strWhere = "  Exists (Select 1 From 人员性质说明 Where 人员id = a.Id And Instr(',医生,护士,', ',' || 人员性质 || ',', 1) <> 0)"
    End If
    
    '当前登录操作员优先显示在前面
    If strNewSQL = "" Then
        strsql = "Select a.Id, a.编号, a.姓名, a.简码" & vbNewLine & _
            " From 人员表 a, 部门人员 b" & vbNewLine & _
            " Where a.Id = b.人员id And b.部门id = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And" & vbNewLine & _
            "      (a.站点 = ' & gstrNodeNo & ' Or a.站点 Is Null) " & vbNewLine & _
            IIf(blnYn, " And (A.编号 Like [2] Or A.简码 Like [3] Or A.姓名 Like [3])", "") & vbNewLine & _
            IIf(strWhere = "", "", " And " & strWhere) & " Order by Decode(a.id," & IIf(blnYn = True, "[4]", "[2]") & ",0,1),a.编号"
    Else
        strsql = "Select a.Id, a.编号, a.姓名, a.简码" & vbNewLine & _
            " From 人员表 a, 部门人员 b" & vbNewLine & _
            " Where a.Id = b.人员id And b.部门id = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And" & vbNewLine & _
            "      (a.站点 = ' & gstrNodeNo & ' Or a.站点 Is Null)" & IIf(strWhere = "", "", " And " & strWhere) & vbNewLine & _
            strNewSQL
        If blnYn Then
            strsql = " Select a.Id, a.编号, a.姓名, a.简码 From (" & strsql & ") a" & vbNewLine & _
                " Where (A.编号 Like [2] Or A.简码 Like [3] Or A.姓名 Like [3])  Order by Decode(a.id,[4],0,1),a.编号"
        Else
            strsql = " Select a.Id, a.编号, a.姓名, a.简码 From (" & strsql & ") a Order by Decode(a.id,[2],0,1),a.编号"
        End If
    End If
    If blnYn = True Then
        Set GetDataToPersons = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng科室ID, UCase(strIn) & "%", gstrLike & UCase(strIn) & "%", UserInfo.id)
    Else
        Set GetDataToPersons = gobjDatabase.OpenSQLRecord(strsql, Me.Caption, mlng科室ID, UserInfo.id)
    End If
    
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

