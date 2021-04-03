VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmBalance 
   AutoRedraw      =   -1  'True
   Caption         =   "病人结帐单"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   Icon            =   "frmBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Picture         =   "frmBalance.frx":08CA
   ScaleHeight     =   8130
   ScaleWidth      =   11790
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10305
      TabIndex        =   26
      ToolTipText     =   "热键:Esc"
      Top             =   7275
      Width           =   1410
   End
   Begin VB.CommandButton cmd结算卡 
      Caption         =   "结算卡(&V)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7335
      TabIndex        =   74
      ToolTipText     =   "热键：F5"
      Top             =   7275
      Width           =   1410
   End
   Begin VB.CommandButton cmdYB 
      Caption         =   "门诊验证(&Y)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   71
      ToolTipText     =   "医保病人身份验证,热键F6"
      Top             =   520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Frame fraTitle 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   0
      TabIndex        =   29
      Top             =   -120
      Width           =   12165
      Begin MSCommLib.MSComm com 
         Left            =   8880
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.PictureBox pic状态 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   3225
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   653
         Visible         =   0   'False
         Width           =   3255
         Begin VB.Label lbl付款方式 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1200
            TabIndex        =   70
            Top             =   30
            Width           =   1920
         End
         Begin VB.Label lbl状态 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   75
            TabIndex        =   51
            Top             =   30
            Width           =   960
         End
      End
      Begin VB.ComboBox cboNO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   10050
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   645
         Width           =   1515
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "废"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   11595
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "热键：F8"
         Top             =   630
         Width           =   465
      End
      Begin VB.TextBox txtInvoice 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7680
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   645
         Width           =   1425
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据格式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   210
         Left            =   10920
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6900
         TabIndex        =   27
         Top             =   705
         Width           =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   30
         X2              =   18000
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   25000
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "废"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   10935
         TabIndex        =   42
         Top             =   660
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9285
         TabIndex        =   32
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人结帐单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   31
         Top             =   180
         Width           =   1875
      End
   End
   Begin VB.Frame fraPatient 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   0
      TabIndex        =   30
      Top             =   825
      Width           =   12165
      Begin zlIDKind.IDKindNew IDKIND 
         Height          =   345
         Left            =   570
         TabIndex        =   75
         Top             =   195
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   609
         Appearance      =   2
         IDKindStr       =   $"frmBalance.frx":0C0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;F2;CTRL+F4;F6;F8;F9;F11;F12;ESC"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txt费别 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
      End
      Begin VB.TextBox txt科室 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10350
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.TextBox txtBed 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txt标识号 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtOld 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4590
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   585
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   600
      End
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1250
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "热键：F11"
         Top             =   180
         Width           =   1250
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5235
         TabIndex        =   52
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9850
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblBed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8760
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lbl标识号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6720
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4095
         TabIndex        =   36
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2930
         TabIndex        =   35
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   80
         TabIndex        =   34
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame fraDate 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7305
      TabIndex        =   56
      Top             =   1305
      Width           =   4860
      Begin VB.Frame fra费用期间 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   105
         TabIndex        =   64
         Top             =   615
         Width           =   4665
         Begin MSMask.MaskEdBox txtEnd 
            Height          =   360
            Left            =   3050
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   14737632
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtBegin 
            Height          =   360
            Left            =   645
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   14737632
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl费用 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费用"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   68
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lbl至 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   2400
            TabIndex        =   67
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.Frame fra结帐时间 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   105
         TabIndex        =   60
         Top             =   1395
         Width           =   4620
         Begin VB.TextBox txt天数 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   0
            Width           =   645
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   645
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   0
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   14737632
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd hh:mm:ss"
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结帐"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   63
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lbl天 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "天"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3870
            TabIndex        =   62
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.OptionButton opt中途 
         Caption         =   "中途结帐"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1620
         TabIndex        =   14
         Top             =   255
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton opt出院 
         Caption         =   "出院结帐"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdPar 
         Caption         =   "结帐设置(&S)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   75
         TabIndex        =   13
         ToolTipText     =   "热键：F9"
         Top             =   180
         Width           =   1365
      End
      Begin VB.Frame fra住院期间 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   105
         TabIndex        =   57
         Top             =   1005
         Width           =   4665
         Begin MSMask.MaskEdBox txtPatiEnd 
            Height          =   360
            Left            =   3050
            TabIndex        =   17
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPatiBegin 
            Height          =   360
            Left            =   645
            TabIndex        =   16
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl至 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   2400
            TabIndex        =   59
            Top             =   60
            Width           =   240
         End
         Begin VB.Label lbl住院 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   58
            Top             =   60
            Width           =   480
         End
      End
   End
   Begin VB.Frame fraBalance 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   7290
      TabIndex        =   39
      Top             =   3000
      Width           =   4860
      Begin VB.CommandButton cmdReturnCash 
         Caption         =   "退现"
         Height          =   330
         Left            =   60
         TabIndex        =   84
         Top             =   2880
         Visible         =   0   'False
         Width           =   560
      End
      Begin VB.PictureBox picThreeBalance 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   30
         ScaleHeight     =   675
         ScaleWidth      =   4455
         TabIndex        =   78
         Top             =   135
         Visible         =   0   'False
         Width           =   4455
         Begin VB.PictureBox picDelDeposit 
            Height          =   420
            Left            =   1275
            ScaleHeight     =   360
            ScaleWidth      =   3060
            TabIndex        =   79
            Top             =   90
            Width           =   3120
            Begin VB.Label lblDepositMoney 
               Caption         =   "2222"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   75
               TabIndex        =   80
               Top             =   15
               Width           =   3090
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDelBalance 
            Height          =   870
            Left            =   180
            TabIndex        =   81
            Top             =   945
            Width           =   4215
            _cx             =   7435
            _cy             =   1535
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
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
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483633
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmBalance.frx":0CA2
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
            ExplorerBar     =   0
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
         End
         Begin VB.Label lblDelDeposit 
            AutoSize        =   -1  'True
            Caption         =   "返回预交款"
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
            Left            =   90
            TabIndex        =   83
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label lbl误差 
            AutoSize        =   -1  'True
            Caption         =   "误差:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   60
            TabIndex        =   82
            Top             =   585
            Visible         =   0   'False
            Width           =   570
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid mshDeposit 
         Height          =   1065
         Left            =   30
         TabIndex        =   19
         Top             =   450
         Width           =   4455
         _cx             =   7858
         _cy             =   1879
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         ExplorerBar     =   0
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
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
         Height          =   870
         Left            =   45
         TabIndex        =   76
         Top             =   1935
         Width           =   4785
         _cx             =   8440
         _cy             =   1535
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         ExplorerBar     =   0
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
      End
      Begin VB.TextBox txtOwe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3720
         Width           =   1560
      End
      Begin VB.Label lblTicketCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交款收据:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2385
         TabIndex        =   69
         Top             =   3780
         Width           =   2400
      End
      Begin VB.Label lbl个人帐户 
         AutoSize        =   -1  'True
         Caption         =   "帐户余额:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2160
         TabIndex        =   47
         Top             =   1665
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lbl医保基金 
         AutoSize        =   -1  'True
         Caption         =   "统筹支付:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   60
         TabIndex        =   46
         Top             =   1665
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblDeposit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "冲预交:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   105
         TabIndex        =   45
         Top             =   165
         Width           =   840
      End
      Begin VB.Label lblSpare 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交余额:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2130
         TabIndex        =   44
         Top             =   165
         Width           =   1080
      End
      Begin VB.Label lblOwe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "差额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   40
         Top             =   3780
         Width           =   480
      End
      Begin VB.Label lblDelMoney 
         AutoSize        =   -1  'True
         Caption         =   "退支付宝:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   675
         TabIndex        =   77
         Top             =   2940
         Visible         =   0   'False
         Width           =   4305
      End
   End
   Begin VB.TextBox txtMoney 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1770
      MaxLength       =   10
      TabIndex        =   41
      Top             =   2220
      Visible         =   0   'False
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   4680
      Left            =   30
      TabIndex        =   10
      Top             =   1875
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   8255
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   7770
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2884
            MinWidth        =   882
            Picture         =   "frmBalance.frx":0D46
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10901
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "病人余额"
            Object.ToolTipText     =   "病人余额"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "设置"
            TextSave        =   "设置"
            Key             =   "LocalParSet"
            Object.ToolTipText     =   "本地参数设置"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.TabStrip tabCard 
      Height          =   5205
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   9181
      TabFixedWidth   =   1409
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      TabMinWidth     =   1411
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "结帐表"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "明细表"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "项目明细"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "分类表"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "分月表"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "费目表"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "逐日单据"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "逐日费目"
            ImageVarType    =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshQuery 
      Height          =   4770
      Left            =   30
      TabIndex        =   11
      Top             =   1815
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   8414
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483631
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      MergeCells      =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fra备注 
      Height          =   555
      Left            =   30
      TabIndex        =   73
      Top             =   6600
      Width           =   7260
      Begin VB.TextBox txt备注 
         Height          =   350
         Left            =   660
         MaxLength       =   50
         TabIndex        =   21
         Top             =   150
         Width           =   6540
      End
      Begin VB.Label lbl备注 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   75
         TabIndex        =   20
         Top             =   210
         Width           =   420
      End
   End
   Begin VB.Frame fraAppend 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   15
      TabIndex        =   48
      Top             =   7155
      Width           =   7290
      Begin VB.Frame fra找补 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   2880
         TabIndex        =   53
         Top             =   120
         Width           =   4410
         Begin VB.TextBox txt缴款 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   690
            MaxLength       =   12
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   0
            Width           =   1470
         End
         Begin VB.TextBox txt找补 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   2940
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lbl缴款 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "缴款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   0
            TabIndex        =   55
            Top             =   45
            Width           =   690
         End
         Begin VB.Label lbl找补 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   2235
            TabIndex        =   54
            Top             =   45
            Width           =   690
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   12
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   60
         TabIndex        =   49
         Top             =   165
         Width           =   690
      End
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
      Height          =   420
      Left            =   8895
      TabIndex        =   25
      ToolTipText     =   "热键：F2"
      Top             =   7260
      Width           =   1410
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Visible         =   0   'False
      Begin VB.Menu mnuFileZero 
         Caption         =   "显示零费用(&Z)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuDepositClear 
         Caption         =   "清除冲预交(&C)"
      End
      Begin VB.Menu mnuPopuDepositAll 
         Caption         =   "使用所有预交款(&A)"
      End
      Begin VB.Menu mnuPopuDepositBalance 
         Caption         =   "按结帐金额使用预交(&J)"
      End
      Begin VB.Menu mnuPopSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColsVisible 
         Caption         =   "显示列选择(&S)"
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "单据号(&N)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "票据号(&R)"
            Index           =   1
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "日期(&D)"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "结算方式(&T)"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "余额(&B)"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "冲预交(&P)"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "金额(&J)"
            Index           =   6
         End
      End
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'入口参数：
Public mlngPatientID As Long        '当前要结帐的病人ID
Public mbytInState As Byte          '0=结帐状态(默认新增,作废),1=浏览状态
Public mbytFunc As Byte              '0-门诊结帐;1-住院结帐
Public mblnViewCancel As Boolean    '是否查看已作废单据
Public mstrInNO As String           '要浏览或作废的单据号
Public mblnNOMoved As Boolean       '操作的单据是否在后备数据表中
Public mlngBillID As Long           '要游览单据的ID
Public mstrPrivs As String
Public mlngModul As Long
Public mstr主页Id As String   '结某次费用:0-结门诊;1-结住院第几次费用;空为不处理
Public mbln门诊转住院 As Boolean 'true:门诊转住院调用接口;False为其他
Public mstrPepositDate As String '指定特点的预交日期(主要是应用于门诊转住院费用时,使用转入的预交进行结帐)
'------------------------------------------------------------
Private mrsInfo As ADODB.Recordset '病人信息(病人ID,姓名,性别,年龄,住院号,床号,在院标志)
Private mrsBalance As ADODB.Recordset '病人未结病人明细
Private mrsDeposit As ADODB.Recordset '病人剩余预交明细
Private mcurSpare As Currency '病人费用余额
Private mlng领用ID As Long
Private mstrStyle As String
Private mstr结帐主页IDs As String
Private mstr录入住院号 As String
Private mblnThreeDepositAfter As Boolean, mblnThreeDepositAll As Boolean
Private mstr住院次数 As String
Private mcurTotal As Currency
Private mblnOutUse As Boolean
Private mstrDepositNo As String
Private mbln打印预交收据 As Boolean
Private mcur误差金额 As Currency
Private mblnPrint As Boolean '根据参数和操作选择决定是否打印票据
Private mstrDec As String   '本次结帐的费用最大小数位数,缺省为gstrDec
Private mblnNOCancel As Boolean '弹出结帐条件窗体时禁止取消
Private mintPatientRange As Integer '按姓名查找时,是否只显示未结费用的病人,0-含已结清,1-未结清,2-体检未结清,3-住院未结清
Private mblnSetPar As Boolean '本次结帐是否进行了结帐条件设置
Private mint结帐类型 As Integer '结帐类型
Private mblnOneCard As Boolean      '是否启用了一卡通接口
Private mrsOneCard As ADODB.Recordset
Private mstrOneCard As String       '读卡时所选择的一卡通接口对应的结算方式
Private mstr本次住院日期 As String
Private mblnShowThree As Boolean
Private mstrCardPrivs As String, mstrForceNote As String
Private mblnNotClearBill As Boolean '未清除单据
Private mblnNotClick As Boolean, mstrInvoice As String
'医保变量--------------------
Private mrs结算方式 As ADODB.Recordset
Private mstr缺省结算 As String '缺省结算方式
Private mstrBalance As String '医保返回的各种结算金额:"结算方式;金额;是否允许修改|..."
Private mbln合约单位 As Boolean
Private mbln个帐结算 As Boolean '本次是否返回了个帐结算
Private mcur个帐余额 As Currency '个人帐户余额
Private mcur个帐限额 As Currency '个人帐户最大限额
Private mcur个帐透支 As Currency '个人帐户允许透支金额
Private mstrYBPati As String    '医保病人身份信息
Private mintInsure As Integer   '作废时,读取的单据中的险类,用来判断是否退现金,算误差等
Private mbln医保作废全退 As Boolean     '是否有不支持的作废结算方式
Private mbytMCMode As Byte '医保病人身份证验模式,包括1-门诊,2-住院两种模式,0-表示非医保
Private mblnMC_TwoMode As Boolean '是否支持门诊和住院医保病人身份证验两种模式
Private mblnUnload As Boolean
'每个病人开始时初始(用于显示在设置窗体)
Private mstrAllTime As String '病人所有未结帐住院次数
Private mstrUnAuditTime As String '病人所有未审核住院次数
Private mstrAllUnit As String '病人所有未结帐科室
Private mstrALLItem As String '病人所有未结收据费目
Private mstrAllClass As String '病人所有未结费用类型
Private mstrALLChargeType As String '病人所有未结的收费类别 '34260
Private mstrAllDiagnose As String
Private mMinDate As Date, mMaxDate As Date
Private mblnDateMoved As Boolean '病人的登记时间是否在转出数据之前

'每个病人结完后初始(作为结帐参数)
Private mstrTime As String  '病人结帐次数(初始="",可以为"0,1,2,3...",0表示主页ID为空)
Private mDateBegin As Date  '病人结帐的开始时间,初始为'1900-01-01'
Private mDateEnd As Date    '病人结帐的结束时间,初始为'3000-01-01'
Private mstrUnit As String '病人结帐科室ID串(初始="",可以为"0,1,2,3...",0表示开单部门ID为空)
Private mstrClass As String  '费用类型=""-所有费用(含未设置),"'公费','比例',..."
Private mstrChargeType As String '收费类别 '34260
Private mbytBaby As Byte '是否仅结算婴儿费用(0-所有费用,1-病人费用,2及以上-第mbytbaby-1个婴儿费用)
Private mstrItem As String '要结的收据费目
Private mbytKind As Byte  '0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
Private mstrDiagnose As String

Private Const COL_标志 = 0
Private Const COL_住院 = 1
Private Const COL_科室 = 2
Private Const COL_时间 = 3
Private Const COL_单据号 = 4
Private Const COL_项目 = 5
Private Const COL_费目 = 6
Private Const COL_婴儿费 = 7
Private Const COL_ID = 8
Private Const COL_序号 = 9
Private Const COL_记录性质 = 10
Private Const COL_记录状态 = 11
Private Const COL_执行状态 = 12
Private Const COL_主页ID = 13
Private Const COL_开单部门ID = 14
Private Const COL_登记时间 = 15
Private Const COL_未结金额 = 16
Private Const COL_结帐金额 = 17
Private Const COL_类型 = 18
Private mcllThreeCard As Collection

'预交清单列标题,结帐时
'Private Const mstrDepositHeader = "ID|0|1,单据号|920|1,票据号|920|1,日期|940|6,结算方式|640|1,余额|980|7,冲预交|980|7"
'预交清单列标题,查看时
'Private Const mstrDepositRHeader = "ID|0|1,单据号|920|1,票据号|920|1,日期|1160|6,结算方式|940|1,金额|980|7"

Private Enum COLMoney
    C0名称 = 0
    C1金额 = 1
    C2号码 = 2
    C3性质 = 3
    C4缺省 = 4  '读取时才有该列
End Enum

'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    '1.门诊，住院结算共用的参数
    分币处理 As Boolean
    
    '2.门诊结算用的参数
    门诊病人结算作废 As Boolean
    门诊必须传递明细 As Boolean
    门诊预结算 As Boolean
    门诊结算_结帐设置 As Boolean
    
    '3.住院结算用的参数
    未结清出院 As Boolean
    结算使用个人帐户 As Boolean
    出院结算必须出院 As Boolean
    出院病人结算作废 As Boolean
    中途结算仅处理已上传部分 As Boolean
    结帐设置后调用接口 As Boolean
    结帐作废后打印回单 As Boolean
    允许结多次住院费用 As Boolean
    
    
End Type

Private MCPAR As TYPE_MedicarePAR
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private Type Ty_ModulePara
    int退款票据 As Integer  '0-不打印,1-提示打印,2-不提示打印;'刘兴洪 问题:27776 日期:2010-02-04 16:49:03
    int预交票据 As Integer
    bln结帐后不清信息 As Boolean    ''刘兴洪 问题:27776 日期:2010-02-04 16:49:03
    bln结帐检查病历接收 As Boolean '30036
    byt缴款输入控制 As Byte  '
    bytMzDeposit As Byte    '门诊预交缺省使用方式:0-缺省不使用交;1-按结帐金额使用预交;2-使用所有预交
    bln结帐退款方式 As Boolean 'True-结帐退款默认按预交结算方式 False-结帐退款默认现金
End Type
Private mty_ModulePara As Ty_ModulePara

'关于消费卡的处理变量
Private Type Ty_SquareCard
    blnExistsObjects As Boolean     '安装了消费卡的
    rsSquare As ADODB.Recordset
    dbl刷卡总额 As Double
    bln卡结算 As Boolean '当前读取的单据是卡结算
    str刷卡结算 As String   '刷卡结算方式;金额;是否允许修改|..."
End Type
Private mtySquareCard As Ty_SquareCard
Private mFactProperty As Ty_FactProperty
Private mobjInPatient As Object
Private mblnFirst As Boolean
'票据相关
Private mlngShareUseID As Long '共享领用批次ID
Private mstrUseType As String '使用类别
Private mintInvoiceFormat As Integer  '打印的发票格式,发票格式序号
Private mintInvoiceMode As Integer '0-不打印;1-自动打印;2-选择打印
Private mblnStartFactUseType As Boolean  '是否启用了多种使用类型票据
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mint预交类别 As Integer  '0-门诊和住院;1-门诊;2-住院
Private mlngCardTypeID As Long '当前刷卡类型56615

Private mrsCardType As ADODB.Recordset '医疗卡类别
Private mobjPayCards As Cards

Private Type TY_BrushCard    '刷卡类型
    str卡号 As String
    str密码 As String
    str交易流水号 As String    '交易流水号
    str交易说明  As String     '交易信息
    str扩展信息 As String    '交易的扩展信息
    dbl帐户余额 As Double
    dblMoney As Double     '当前退款或刷卡金额
End Type
Private mCurBrushCard As TY_BrushCard   '当前的刷卡信息



Private Sub InitDepositGridHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格
    '编制:刘兴洪
    '日期:2015-04-27 17:32:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, varData As Variant, blnView As Boolean
    Dim j As Long
    On Error GoTo errHandle
    blnView = mstrInNO <> "" Or cboNO.Text <> "" Or chkCancel.Value = 1
    With mshDeposit
        .Clear
        .Cols = 9: .Rows = 2: i = 0
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .TextMatrix(0, i) = "ID": .ColWidth(i) = 0: .ColHidden(i) = True: .ColData(i) = "-1||1": i = i + 1
        .TextMatrix(0, i) = "单据号": .ColWidth(i) = 920: .ColHidden(i) = False: .ColData(i) = "1||0": i = i + 1
        .TextMatrix(0, i) = "票据号": .ColWidth(i) = 920: .ColHidden(i) = False: .ColData(i) = "0||0": i = i + 1
        .TextMatrix(0, i) = "日期": .ColWidth(i) = IIf(blnView, 1160, 940): .ColHidden(i) = False: .ColData(i) = "0||0": i = i + 1
        .TextMatrix(0, i) = "结算方式": .ColWidth(i) = IIf(blnView, 640, 940): .ColHidden(i) = False: .ColData(i) = "0||0": i = i + 1
        .TextMatrix(0, i) = "余额": .ColWidth(i) = 980: .ColHidden(i) = IIf(blnView, True, False): .ColData(i) = "0||0": i = i + 1
        .TextMatrix(0, i) = "冲预交": .ColWidth(i) = 980: .ColHidden(i) = IIf(blnView, True, False): .ColData(i) = "1||0": i = i + 1
        .TextMatrix(0, i) = "金额": .ColWidth(i) = 980: .ColHidden(i) = IIf(Not blnView, True, False): .ColData(i) = "1||0": i = i + 1
        .TextMatrix(0, i) = "预交ID": .ColWidth(i) = 0: .ColHidden(i) = True: .ColData(i) = "-1||1": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            Select Case .ColKey(i)
            Case "余额", "冲预交", "金额"
                .ColAlignment(i) = flexAlignRightCenter
            Case Else
                .ColAlignment(i) = flexAlignLeftCenter
            End Select
        Next
    End With
    zl_vsGrid_Para_Restore mlngModul, mshDeposit, Me.Caption, "预交列表" & IIf(blnView, "1", "0")
    With mshDeposit
        For j = 0 To mnuViewToolCols.UBound
            For i = 0 To .Cols - 1
              If mnuViewToolCols(j).Caption Like .ColKey(i) & "*" Then
                  mnuViewToolCols(j).Checked = .ColHidden(i) = False
                  Select Case .ColKey(i)
                  Case "余额", "冲预交"
                    mnuViewToolCols(j).Visible = Not blnView
                  Case "金额"
                    mnuViewToolCols(j).Visible = blnView
                  End Select
                  Exit For
              End If
            Next
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlInitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2010-02-04 16:50:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mty_ModulePara
        '0-不打印,1-提示打印,2-不提示打印;'刘兴洪 问题:27776 日期:2010-02-04 16:49:03
        .int退款票据 = Val(zlDatabase.GetPara("退款收据打印", glngSys, mlngModul))
        .int预交票据 = Val(zlDatabase.GetPara("预交票据打印方式", glngSys, mlngModul))
        .bln结帐后不清信息 = IIf(Val(zlDatabase.GetPara("结帐后不清除信息", glngSys, mlngModul)) = 1, True, False)
        .bln结帐检查病历接收 = IIf(Val(zlDatabase.GetPara("结帐检查病历接收", glngSys, mlngModul)) = 1, True, False) '30036
        '问题:43153::0-不进行控制;1-存在收取现金时,必须输入缴款.
        .byt缴款输入控制 = Val(zlDatabase.GetPara("结帐缴款输入控制", glngSys, mlngModul, 0))
        .bytMzDeposit = Val(zlDatabase.GetPara("门诊预交缺省使用方式", glngSys, mlngModul, 2))
        .bln结帐退款方式 = IIf(Val(zlDatabase.GetPara("结帐退款缺省方式", glngSys, mlngModul)) = 1, True, False)
    End With
    
End Sub

Private Sub cmdReturnCash_Click()
    Dim dblMoney As Double, lngRow As Long
    Dim str操作员姓名 As String, strDBUser As String
    Dim strPrivs As String
    Dim intCount As Integer, intNotCashCount As Integer
    
    
    If mstrForceNote <> "" Then Exit Sub
    Call GetDelThreeCardDepositInfor(mblnThreeDepositAll, intCount, intNotCashCount, mblnThreeDepositAfter, mstrStyle)
    
    If mstrStyle = "" Then Exit Sub
    
    
    If InStr(";" & mstrCardPrivs & ";", ";三方退款强制退现;") = 0 And intNotCashCount > 0 Then 'intNotCashCount:不允许退现的结算方式
        str操作员姓名 = zlDatabase.UserIdentifyByUser(Me, "强制退现验证", glngSys, 1151, "三方退款强制退现")
        If str操作员姓名 = "" Then
            MsgBox "录入的操作员验证失败或者录入的操作员不具备强制退现权限，不能强制退现！", vbInformation, gstrSysName
            Exit Sub
        End If
        mstrForceNote = str操作员姓名 & "强制退现:" & lblDelMoney.Caption
    Else
        If intNotCashCount <> 0 Then
            If MsgBox("选择的结算卡不支持退现,是否强制将其退现？", _
                                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
        End If
        mstrForceNote = mstrForceNote & IIf(mstrForceNote = "", UserInfo.姓名 & "强制退现:", ";") & lblDelMoney.Caption
    End If
    
    mblnThreeDepositAfter = False
    Call ShowMoney(False, , mty_ModulePara.bytMzDeposit, True)
End Sub

Private Sub cmd结算卡_Click()
    Dim dblTotal As Double, rsFeeList As ADODB.Recordset
    Err = 0: On Error GoTo ErrHand:

    If mtySquareCard.blnExistsObjects = False Then Exit Sub
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    If Not IsNull(mrsInfo!险类) Then
        ShowMsgbox "目前结算卡不支持医保结算,请检查"
        Exit Sub
    End If

    '结算卡的一些相关处理
    dblTotal = Get可刷金额
    If dblTotal <= 0 Then
         Call MsgBox("没有可刷结算卡的金额,不必刷卡!", vbInformation + vbDefaultButton1, gstrSysName)
         Exit Sub
    End If

    Screen.MousePointer = 11
    If zlSquareCardFeeList(rsFeeList) = False Then Exit Sub

    '调用接口
    'Public Function zlBrushCardSquare(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal rsFeeList As ADODB.Recordset, ByVal dbl最大消费 As Double, ByRef rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: zlBrushCardSquare (刷卡结算接口)
    '入参:frmMain:HIS传入 调用的主窗体
    '     intCallType : HIS传入 0-  门诊费用调用 1-  住院结帐调用
    '     rsFeeList: HIS传入 如果是门诊多单据,则所有单据的明细,如果是住院结帐 , 则是本次结帐的所有明细
    '     dbl最大消费 :  HIS传入 表示刷卡不能超过此金额
    '
    '出参:rsSquare : 接口返回    本地记录集:接口传入空结构(接口返回相关的数据) , 结构如下:
    '                接口编号 , 消费卡ID, 结算方式, 结算金额, 卡号卡名称, 交易流水号, 交易时间, 备注
    '     rsSquare说明:主要是解决同一单据,刷多张卡消费的情况.,如果本次刷多张卡 , 则传入接口中已经刷过的卡信息
    '     rs分摊情况:单据序号 消费卡ID,卡号,结算方式,分摊额
    '返回:true:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:18:38
    '说明:
    '    1.  在门诊收费界面时,HIS在点"结算卡"时,调用本接口
    '    2.  在住院结帐界面时,HIS在点"结算卡"时,调用本接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlBrushCardSquare(Me, mlngModul, mstrPrivs, rsFeeList, dblTotal, mtySquareCard.rsSquare) = False Then
        GoTo goRestoreMouse:
    End If
    
    If mtySquareCard.rsSquare Is Nothing Then GoTo goRestoreMouse:
    If mtySquareCard.rsSquare.State <> 1 Then GoTo goRestoreMouse:
    '需要根据返回结果,重新计算单据
    If mtySquareCard.rsSquare.RecordCount = 0 Then
        Set mtySquareCard.rsSquare = Nothing: GoTo goRestoreMouse:
    End If
    If 住院刷结算卡() = False Then GoTo goRestoreMouse:


goRestoreMouse:
    Screen.MousePointer = 0
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
End Sub

 

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call FindPati(objCard, True, txtPatient.Text)
        End If
        Exit Sub
    End If
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
    
End Sub
 
Private Sub SetOneCardBalance()
'功能: 设置一卡通结算方式
    Dim curOneCard As Currency, strName As String
    
    If mblnOneCard And Not mobjICCard Is Nothing Then
        curOneCard = mobjICCard.GetSpare(strName)
        If curOneCard <> 0 Then
           mrsOneCard.Filter = "名称='" & strName & "'"
           If mrsOneCard.RecordCount > 0 Then mstrOneCard = mrsOneCard!结算方式
        End If
        sta.Panels(2).Text = "卡余额:" & Format(curOneCard, "0.00") & "元"
    End If
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
 

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    Call FindPati(objCard, True, objPatiInfor.卡号)
End Sub
Private Sub mnuPopuDepositAll_Click()
    '预交款全冲，多余的退给病人
    Call ShowMoney(True, , 2)
End Sub

Private Sub mnuPopuDepositBalance_Click()
    '按结帐金额冲预交
     Call ShowMoney(True, , 1)
End Sub

Private Sub mnuPopuDepositClear_Click()
    '清除预交金额
     Call ShowMoney(True, , 0)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim objCard As Card
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Set objCard = IDKIND.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    Call FindPati(objCard, True, strCardNo)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim objCard As Card
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Set objCard = IDKIND.GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    Call FindPati(objCard, True, strID)
End Sub



Private Sub SetDisibleColor(Optional bln As Boolean = False)
    If Not bln Then
        txtPatient.BackColor = &HE0E0E0
        txtPatiBegin.BackColor = &HE0E0E0
        txtPatiEnd.BackColor = &HE0E0E0
        txtTotal.BackColor = &HE0E0E0
        txtInvoice.BackColor = &HE0E0E0
    Else
        txtPatient.BackColor = &HFFFFFF
        txtPatiBegin.BackColor = &HFFFFFF
        txtPatiEnd.BackColor = &HFFFFFF
        txtTotal.BackColor = &HFFFFFF
        txtInvoice.BackColor = &HFFFFFF
    End If
End Sub

Private Sub InitPatiVariable()
'初始化每个病人结帐数据相关的变量
    mstrTime = "":  mstrUnit = "": mstrClass = "": mbytBaby = 0: mstrItem = "": mbytKind = 0
    mstrChargeType = ""
    mstrDiagnose = ""
    mDateBegin = CDate("0:00:00"): mDateEnd = CDate("0:00:00")
End Sub

Private Sub InitBalanceCondition()
'初始化每个病人结帐条件相关的变量
    mstrAllTime = "":  mstrAllUnit = "": mstrALLItem = "": mstrAllClass = "": mstrUnAuditTime = ""
    mstrALLChargeType = ""  '34260
    mstrAllDiagnose = ""
    mMinDate = #1/1/1900#: mMaxDate = #1/1/1900#
    mblnSetPar = False
End Sub

Private Sub chkCancel_Click()
    Dim i As Long, blnNew As Boolean
            
    blnNew = (chkCancel.Value = 0)
    IDKIND.Enabled = blnNew
    If blnNew Then cboNO.Text = "": mstrInNO = ""
    
    Call NewBill    '其中的InitBalanceSet设置了一些控件状态
    
    txtInvoice.Locked = Not blnNew
    cboNO.Locked = blnNew
    
    fraPatient.Enabled = blnNew
    If mbytFunc = 0 Then
        cmdYB.Visible = blnNew
    Else
        cmdYB.Visible = False
    End If
    cmdPar.Visible = blnNew
    opt出院.Visible = blnNew
    opt中途.Visible = blnNew
    fra住院期间.Enabled = blnNew
    txt备注.Enabled = blnNew: lbl备注.Enabled = blnNew
    fra找补.Visible = blnNew
    lblSpare.Visible = blnNew
    txtTotal.Locked = (Not blnNew) Or (InStr(mstrPrivs, ";结帐设置;") = 0)
    cmd结算卡.Visible = False ' blnNew And mtySquareCard.blnExistsObjects

    Call SetDisibleColor(blnNew)
        
    If Not blnNew Then
        For i = tabCard.Tabs.Count To 2 Step -1
            tabCard.Tabs.Remove i
        Next
        tabCard.SelectedItem = tabCard.Tabs(1)
        Call tabCard_Click
                
        chkCancel.ForeColor = &HFF&
        txtInvoice.Text = ""
        If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
    Else
        tabCard.Tabs.Add 2, , "明细表"
        tabCard.Tabs.Add 3, , "项目明细"
        tabCard.Tabs.Add 4, , "分类表"
        tabCard.Tabs.Add 5, , "分月表"
        tabCard.Tabs.Add 6, , "费目表"
        tabCard.Tabs.Add 7, , "逐日单据"
        tabCard.Tabs.Add 8, , "逐日费目"
        
        chkCancel.ForeColor = 0
        Call ReInitPatiInvoice
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    txtMoney.Visible = False
    
    If mbytInState = 0 Then
        '问题:
        If mty_ModulePara.bln结帐后不清信息 And mblnNotClearBill Then
            If mrsInfo Is Nothing Then
                Call NewBill
                mblnNotClearBill = False
                Exit Sub
            ElseIf mrsInfo.State <> 1 Then
                Call NewBill
                 mblnNotClearBill = False
                Exit Sub
            End If
        End If
        
        If chkCancel.Value = Checked And txtPatient.Text <> "" Then
            If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If mbytMCMode = 1 Then
                If MsgBox("确实要取消当前病人身份证验吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    If YBIdentifyCancel Then Call NewBill
                    Exit Sub
                    '不退出窗体,以便选择其它病人进行身份验证
                End If
            Else
                If Val(txtTotal.Text) <> 0 And mrsInfo.State = adStateOpen Then
                    If MsgBox("该病人尚未确定结帐,确实取消吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    Else
                        Call NewBill
                        Exit Sub
                    End If
                ElseIf txtPatient.Text <> "" Then
                    If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        End If
    End If
    Unload Me
End Sub


Private Function YBIdentifyCancel() As Boolean
'功能：取消医保病人身份验证
'返回：返回假时不退出界面或清除操作
    Dim lng病人ID As Long
    YBIdentifyCancel = True
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    If mstrYBPati <> "" Then
        If UBound(Split(mstrYBPati, ";")) >= 8 Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
        If lng病人ID <> 0 Then YBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng病人ID, mrsInfo!险类)
    End If
End Function

Private Function GetPatientState() As Integer
'功能:获取病人状态
'返回:0-出院,1-在院,2-预出院,-1-访问数据库出错
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    GetPatientState = -1
    On Error GoTo errH
    strSql = "Select a.当前科室id, a.主页id As 最后主页id, b.主页id, b.状态" & vbNewLine & _
            "From 病人信息 A, 病案主页 B" & vbNewLine & _
            "Where a.病人id = b.病人id And Nvl(b.主页id, 0) = [2] And b.病人id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(mrsInfo!病人ID), Val("" & mrsInfo!主页ID))
    
    If rsTmp.RecordCount > 0 Then
        If Val(Nvl(rsTmp!最后主页ID)) > Val(Nvl(rsTmp!主页ID)) Then
            GetPatientState = 0
        Else
            If Not IsNull(rsTmp!当前科室id) Then
                If Val("" & rsTmp!状态) = 3 Then
                    GetPatientState = 2
                Else
                    GetPatientState = 1
                End If
            Else
                GetPatientState = 0
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DelBalance()
    Dim blnTrans As Boolean, blnTransMC As Boolean
    Dim strSql As String, i As Long, lng结帐ID As Long, str误差NO As String, strBalance As String, strAdvance As String
    Dim curDeposit As Currency, blnAdded As Boolean, intCashRow As Long, curRetuCash As Currency
    Dim rsOneCard As ADODB.Recordset, objICCard As Object, strCardNo As String
    Dim strNo As String, lng病人ID As Long, lng冲销ID As Long
    Dim lngCardTypeID As Long, bytInvoiceKind As Byte
    Dim dblReturnDeposit As Double  '返回预交
    Dim cllPro As Collection, strDate As String, lng预交ID As Long
    Dim strDepositNo As String, tyBrushCard As TY_BrushCard
    Dim rsBalance As ADODB.Recordset, dblDelMoney As Double
    
    If mbytFunc = 0 Then
        bytInvoiceKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, "0"))
    Else
        bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, "0"))
    End If
        
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("性质"))) = 1 Then intCashRow = i: Exit For
        Next
    End With
    
    If InStr(1, mstrPrivs, ";预交退现金;") > 0 And Not picThreeBalance.Visible Then
        curDeposit = Val(lblDeposit.Tag)
        If curDeposit <> 0 Then
            If intCashRow > 0 Then
                curRetuCash = CentMoney(curDeposit)
                If curRetuCash <> 0 Then
                    If MsgBox("你要将结帐时冲减的预交款" & curRetuCash & "元退为现金吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        curDeposit = 0
                    Else
                        If curRetuCash <> curDeposit Then
                            '之前mcur误差金额记录的误差是医保不支持回退退现金产生的
                            mcur误差金额 = mcur误差金额 + (curRetuCash - curDeposit)
                            curDeposit = curRetuCash
                        End If
                    End If
                Else
                    curDeposit = 0
                End If
            Else
                curDeposit = 0
            End If
        End If
    End If
    dblReturnDeposit = 0
    If picThreeBalance.Visible Then
        '存在三方卡(处理方式:先退款,然后存预交充值,以现金为准)
        '结算方式|结算金额|结算号码||......
        If Val(lblDepositMoney.Tag) <> 0 Then
            dblReturnDeposit = Val(lblDepositMoney.Tag)
            strBalance = picDelDeposit.Tag & "|" & Val(lblDepositMoney.Tag) & "| "
            With vsDelBalance
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("金额"))) <> 0 Then
                        strBalance = strBalance & "||" & .TextMatrix(i, .ColIndex("结算方式"))
                        strBalance = strBalance & "|" & Val(.TextMatrix(i, .ColIndex("金额")))
                        strBalance = strBalance & "|" & IIf(.TextMatrix(i, .ColIndex("结算号码")) = "", " ", .TextMatrix(i, .ColIndex("结算号码")))
                    End If
                Next
            End With
        End If
    Else
        If mintInsure > 0 Or curDeposit <> 0 Then
            '收集退款方式及金额
            If Not mbln医保作废全退 Or curDeposit <> 0 Then
                With vsfMoney
                    For i = 1 To .Rows - 1
                        If Val(.TextMatrix(i, .ColIndex("金额"))) <> 0 Then  '结算方式|结算金额|结算号码||......  结算号码为空时,以空格分开,以便区分|和||,
                           If Val(.TextMatrix(i, .ColIndex("性质"))) = 1 Then blnAdded = True
                           strBalance = strBalance & "||" & .TextMatrix(i, .ColIndex("结算方式")) & "|" & Val(.TextMatrix(i, .ColIndex("金额"))) + IIf(.TextMatrix(i, .ColIndex("性质")) = 1, curDeposit, 0) & "|" & _
                                    IIf(.TextMatrix(i, .ColIndex("结算号码")) = "", " ", .TextMatrix(i, .ColIndex("结算号码")))
                        End If
                    Next
                    If Not blnAdded And curDeposit <> 0 Then
                        strBalance = strBalance & "||" & .TextMatrix(intCashRow, .ColIndex("结算方式")) & "|" & curDeposit & "| "
                    End If
                End With
            End If
        End If
    End If
    
    If Left(strBalance, 2) = "||" Then
        strBalance = Mid(strBalance, 3)
    End If
    
    strNo = cboNO.Text
    lng结帐ID = GetBalanceID(cboNO.Text, lng病人ID)
    If mblnOneCard Then
        Set rsOneCard = GetOneCardBalance(lng结帐ID)
        If rsOneCard.RecordCount > 0 Then
            On Error Resume Next
            Set objICCard = CreateObject("zlICCard.clsICCard")
            On Error GoTo 0
            If objICCard Is Nothing Then
                MsgBox "一卡通接口创建失败,不能进行退费!请检查接口文件.", vbInformation, gstrSysName
                Exit Sub
            End If
            strCardNo = objICCard.Read_Card(Me)
            If strCardNo = "" Then Exit Sub
            If strCardNo <> rsOneCard!单位帐号 Then
                MsgBox "当前卡号与扣款卡号不一致!不能进行退费.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    Set cllPro = New Collection
    
    lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
   
    If CheckThreeSwapDelValied(lng结帐ID, rsBalance, dblDelMoney) = False Then Exit Sub

    
    If dblReturnDeposit <> 0 Then
        '返回预交款
        If GetDepositSaveSql(lng结帐ID, lng冲销ID, lng病人ID, dblReturnDeposit, strDate, cllPro, strDepositNo, lng预交ID) = False Then Exit Sub
    End If
    
    'Zl_病人结帐记录_Delete
    strSql = "Zl_病人结帐记录_Delete("
    '  No_In           病人结帐记录.No%Type,
    strSql = strSql & "'" & cboNO.Text & "',"
    '  操作员编号_In   病人结帐记录.操作员编号%Type,
    strSql = strSql & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In   病人结帐记录.操作员姓名%Type,
    strSql = strSql & "'" & UserInfo.姓名 & "',"
    '  误差金额_In     病人预交记录.冲预交%Type := 0, --医保或预交退现金产生的误差
    strSql = strSql & "" & mcur误差金额 & ","
    '  结帐作废结算_In Varchar2 := Null, --结算方式|结算金额|结算号码||......
    strSql = strSql & "'" & strBalance & "',"
    '  预交退现金_In   Number := 0, --当预交款退现金时，结算方式及金额通过参数结帐作废结算_In传入
    strSql = strSql & " " & IIf(curDeposit <> 0 Or dblReturnDeposit <> 0, "1", "0") & ","
    '  冲销id_In       病人预交记录.结帐id%Type := Null,
    strSql = strSql & "" & lng冲销ID & ","
    '  冲销时间_In     Date := Null,
    strSql = strSql & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
     '  缴预交id_In     病人预交记录.Id%Type := Null --在作废时将相关的金额充值到预交款时填写
    strSql = strSql & "" & IIf(lng预交ID = 0, "NULL", lng预交ID) & ","
    '  票据号_In        病人结帐记录.实际票号%Type,
    strSql = strSql & "'" & mstrInvoice & "',"
    '  领用ID_In        票据领用记录.ID%Type,
    strSql = strSql & mlng领用ID & ","
    '  票种_In          票据使用明细.票种%Type,
    strSql = strSql & IIf(bytInvoiceKind = 0, 3, 1) & ")"
    zlAddArray cllPro, strSql
       
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill("", lng结帐ID) = False Then Exit Sub
    End If
    
    cmdOK.Enabled = False   '防止医保延时
    On Error GoTo errH
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    '保险接口
    blnTransMC = False
    If mintInsure <> 0 Then
        If mbytMCMode = 1 Then
            If MCPAR.门诊病人结算作废 Then
                strAdvance = lng冲销ID & "|0"
                If Not gclsInsure.ClinicDelSwap(lng结帐ID, , mintInsure, strAdvance) Then
                    gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Sub
                Else
                    blnTransMC = True
                End If
            End If
        Else
            If Not gclsInsure.SettleDelSwap(lng结帐ID, mintInsure) Then
                gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Sub
            Else
                blnTransMC = True
            End If
        End If
    ElseIf Not rsOneCard Is Nothing Then
        If rsOneCard.RecordCount > 0 Then
            If Not objICCard.ReturnSwap(rsOneCard!单位帐号, rsOneCard!医院编码, "" & rsOneCard!结算号码, rsOneCard!金额) Then
                gcnOracle.RollbackTrans
                MsgBox "一卡通退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
                cmdOK.Enabled = True: Exit Sub
            End If
        End If
    End If
            
    If ExecuteThreeSwapDelInterface(lng结帐ID, lng冲销ID, rsBalance, dblDelMoney) = False Then Exit Sub
    gcnOracle.CommitTrans: blnTrans = False
    If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mbytMCMode = 1, 交易Enum.Busi_ClinicDelSwap, 交易Enum.Busi_SettleDelSwap), True, mintInsure)
    cmdOK.Enabled = True   '防止医保延时
    
    strSql = "Zl_病人自动记帐_Restore('" & cboNO.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If Not gobjTax Is Nothing And gblnTax Then
        gstrTax = gobjTax.zlTaxInErase(gcnOracle, lng结帐ID)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    End If
    
    '问题:35554
    If mintInsure <> 0 Then
        If MCPAR.结帐作废后打印回单 And InStr(1, mstrPrivs, ";病人退费回单;") > 0 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me, "结帐ID=" & zlGet结帐冲销ID(lng结帐ID), 2)
        End If
    ElseIf InStr(1, mstrPrivs, ";病人退费回单;") > 0 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me, "结帐ID=" & zlGet结帐冲销ID(lng结帐ID), 2)
    End If
    
    If mbln打印预交收据 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mstrDepositNo, 2)
        mbln打印预交收据 = False
    End If
    
    Call WriteZYInforToCard(lng病人ID, lng冲销ID, True)
    
    If mblnPrint Then
        If ReportOpen(gcnOracle, glngSys, IIf(bytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6"), Me, "结帐ID=" & lng冲销ID, "病人ID=" & lng病人ID, "NO=" & mstrDepositNo, 2) = False Then
            MsgBox "结帐作废红票打印出现错误，请检查打印机是否正确!", vbInformation, gstrSysName
        End If
    End If
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mbytMCMode = 1, 交易Enum.Busi_ClinicDelSwap, 交易Enum.Busi_SettleDelSwap), False, mintInsure)
    End If
    Call SaveErrLog
End Sub

Private Function GetOneCardMoney() As Currency
'功能：获取一卡通结算金额
    Dim i As Long
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, vsfMoney.ColIndex("性质"))) = 7 And Val(vsfMoney.TextMatrix(i, .ColIndex("金额"))) <> 0 Then
                mrsOneCard.Filter = "结算方式='" & vsfMoney.TextMatrix(i, .ColIndex("结算方式")) & "'"
                GetOneCardMoney = Val(vsfMoney.TextMatrix(i, .ColIndex("金额")))
                Exit For
            End If
        Next
    End With
End Function

Private Function GetOneCardCount() As Long
    '功能：获取一共使用了几种一卡通结算方式
    Dim i As Long
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("性质"))) = 7 And Val(.TextMatrix(i, .ColIndex("金额"))) <> 0 Then
                GetOneCardCount = GetOneCardCount + 1
            End If
        Next
    End With
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lngSaveID As Long, i As Long, strNo As String, curDate As Date, curDeposit As Currency, cur消费金额 As Currency, curOneCard As Currency
    Dim blnOut As Boolean, intState As Integer, strInfo As String, strTmp As String, strTime As String
    Dim bln打印退款收据 As Boolean, str病历原因 As String, bln打印预交收据 As Boolean
    Dim bln打印费用明细 As Boolean
    Dim blnPrintBillEmpty As Boolean   '55052
    Dim objCard As Card
    
    If chkCancel.Value = 1 Then '作废结帐单
        If mintInsure > 0 And Not MCPAR.出院病人结算作废 And mbytMCMode <> 1 Then
            If Not isYBPati(CLng(txtPatient.Tag), True) Then
                MsgBox "该参保病人已经出院，不能作废该结帐单！", vbInformation, gstrSysName: Exit Sub
            End If
        End If
        
        If MsgBox("确实要将单据[" & cboNO.Text & "]作废吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        mblnPrint = True
        Select Case mintInvoiceMode
        Case 0: mblnPrint = False '不打印
        Case 2  '自动打印
            If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                mblnPrint = False
            End If
        End Select
        
        If mblnPrint Then
            Dim bytInvoiceKind As Byte
            If mbytFunc = 0 Then
                bytInvoiceKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, "0"))
            Else
                bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, "0"))
            End If
            If gblnStrictCtrl Then   '严格票据管理
                If Trim(mstrInvoice) = "" Then
                    MsgBox "红票票据号无效,请重试！", vbInformation, gstrSysName
                    Exit Sub
                End If
                mlng领用ID = GetInvoiceGroupID(IIf(bytInvoiceKind = 0, 3, 1), 1, mlng领用ID, mlngShareUseID, mstrInvoice, mstrUseType)
                If mlng领用ID <= 0 Then
                    Select Case mlng领用ID
                        Case 0 '操作失败
                        Case -1
                            MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                        Case -2
                            MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                        Case -3
                            MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入", vbInformation, gstrSysName
                    End Select
                    Exit Sub
                End If
            Else
                If Len(mstrInvoice) <> gbytFactLength And mstrInvoice <> "" Then
                    MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Else
            mstrInvoice = ""
        End If
        
        '刘兴洪:28947
        If mintInsure <> 0 Then
            If gclsInsure.CheckInsureValid(mintInsure) = False Then
                Exit Sub
            End If
        End If
        Call DelBalance
        chkCancel.Value = 0 '(并激活事件)
    Else '新单存盘
        txtMoney.Visible = False
        
        '1.数据逻辑检查
        If mrsInfo.State = 0 Then
            MsgBox "没有确定结帐病人,不能存盘！", vbExclamation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
        End If
        
        
        '病人住院时间有效性判断
        If txtPatiBegin.Text <> "____-__-__" And Not IsDate(txtPatiBegin.Text) Then
            MsgBox "请输入一个有效的开始时间！", vbInformation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
        End If
        If txtPatiEnd.Text <> "____-__-__" And Not IsDate(txtPatiEnd.Text) Then
            MsgBox "请输入一个有效的结束时间！", vbInformation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
        End If
        If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
            If txtPatiEnd < txtPatiBegin.Text Then
                MsgBox "结束时间不能小于开始时间！", vbInformation, gstrSysName
                If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
            End If
        End If
        If IsDate(txtPatiBegin.Text) And Not IsDate(txtPatiEnd.Text) Then
            MsgBox "请一并输入有效的结束时间！", vbInformation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
        End If
        If Not IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
            MsgBox "请一并输入有效的开始时间！", vbInformation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
        End If
            
        If mshDetail.Rows = 2 And mshDetail.TextMatrix(1, 0) = "" Then
            MsgBox "该设置下病人没有需要结帐的费用！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CCur(txtOwe.Text) <> 0 Then
            If InStr(1, mstrPrivs, ";仅用预交结帐;") = 0 Then
                If CCur(txtOwe.Text) > 0 Then
                    MsgBox "病人缴款不足,请按所显示的差额补款！", vbExclamation, gstrSysName
                    If vsfMoney.Enabled And vsfMoney.Visible Then vsfMoney.SetFocus
                    Exit Sub
                Else
                    MsgBox "病人缴款过多,请按所显示的差额补退病人！", vbExclamation, gstrSysName
                    If vsfMoney.Enabled And vsfMoney.Visible Then vsfMoney.SetFocus
                    Exit Sub
                End If
            Else
                If CCur(txtOwe.Text) > 0 Then
                    MsgBox "病人冲预交金额不足,请按所显示的差额调整冲预交金额！", vbExclamation, gstrSysName
                    If vsfMoney.Enabled And vsfMoney.Visible Then vsfMoney.SetFocus
                    Exit Sub
                Else
                    MsgBox "病人冲预交金额过多,请按所显示的差额调整冲预交金额！", vbExclamation, gstrSysName
                    If vsfMoney.Enabled And vsfMoney.Visible Then vsfMoney.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
            For i = 1 To vsfMoney.Rows - 1
                If vsfMoney.RowData(i) = 999 Then
                    If Val(vsfMoney.TextMatrix(i, 1)) < 0 Then
                        MsgBox "仅用预交结帐情况下，结帐不支持退款！", vbExclamation, gstrSysName
                        Exit Sub
                    End If
                End If
            Next i
        End If
        
        '43153
        '缴款控制:0-不进行控制;1-存在收取现金时,必须输入缴款.
        If mty_ModulePara.byt缴款输入控制 <> 0 And Val(txt找补.Tag) < 0 And Val(txt缴款.Text) = 0 Then
            MsgBox "你还未输入缴款金额,不能继续", vbExclamation, gstrSysName
            If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
            zlControl.TxtSelAll txt缴款: Exit Sub
        End If
        '刘兴洪:问题:25596
        If zlCommFun.StrIsValid(txt备注.Text, 50, txt备注.hWnd, "备注") = False Then Exit Sub
        
        '2.业务规则检查
        If mbytMCMode <> 1 And (mbytFunc = 1 And mbytInState = 0) Then
            intState = GetPatientState
            If Not IsNull(mrsInfo!险类) And opt出院.Value Then
                If MCPAR.出院结算必须出院 And intState <> 0 Then
                    If IsNull(mrsInfo!当前科室) Then
                        MsgBox "病人在结帐期间被撤销出院,医保病人出院结帐前必须先出院！", vbInformation, gstrSysName
                    Else
                        MsgBox "医保病人出院结帐前必须先出院！", vbInformation, gstrSysName
                    End If
                    Exit Sub
                End If
            End If
            
            '是否在院
            If gbln在院不准结帐 And opt出院.Value And (intState = 1 Or intState = 2) Then '  ' 30572:预出院也是在院.
                MsgBox "当前病人在院，不允许出院结帐。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '检查是否还有代收费用未退还病人
            If opt出院.Value = True Then
                If PatiHaveStorage(mrsInfo!病人ID) Then
                    Exit Sub
                End If
            End If
            
            'gbytAuditing:0-不检查,1-检查并提示,2-检查并禁止
            '问题:37369:中途结帐不检查
            If gbytAuditing <> 0 And opt出院.Value Then
                '61345:刘尔旋,2014-02-11,只检查需要结帐的住院次数的费用
'                strHosTimes = ""
'                For i = 0 To frmSetBalance.lstTime.ListCount - 1
'                    If frmSetBalance.lstTime.Selected(i) = True Then strHosTimes = strHosTimes & "," & frmSetBalance.lstTime.ItemData(i)
'                Next i
'                strHosTimes = Mid(strHosTimes, 2)
'                If strHosTimes = "0" Then strHosTimes = ""
                If HaveNOAuditing(mrsInfo!病人ID, mstrTime) Then
                    If gbytAuditing = 1 Then
                        '在读取病人信息时,已经检查了
                    ElseIf gbytAuditing = 2 Then
                         Call MsgBox("该病人还存在未审核的记帐费用,禁止结帐!", vbInformation + vbOKOnly, gstrSysName)
                         Exit Sub
                    End If
                End If
            End If
                        
            '需要再次检查,以防结帐期间已审核的病人被取消审核
            If (InStr(mstrPrivs, ";未审核病人中途结帐;") = 0 And opt中途.Value _
                Or InStr(mstrPrivs, ";未审核病人出院结帐;") = 0 And opt出院.Value) Then
                
                strTime = IIf(mstrTime = "", mstrAllTime, mstrTime)
                If strTime <> "" Then
                    For i = 0 To UBound(Split(strTime, ","))
                        strTmp = Split(strTime, ",")(i)
                        If Val(strTmp) <> 0 Then
                            If Not Chk病人审核(mrsInfo!病人ID, Val(strTmp)) Then
                                MsgBox "待结帐费用中包含病人第" & strTmp & "次住院未审核的费用记录。" & vbCrLf & _
                                    "你不能对未审核的费用进行结帐！", vbInformation, gstrSysName
                                If cmdPar.Visible And cmdPar.Enabled Then cmdPar.SetFocus: Exit Sub
                            End If
                        End If
                    Next
                End If
            End If
        End If
                      
         
         '检查病人是否有未执行完成的诊疗项目及未发药品
        If opt出院.Value And mbytFunc = 1 Then
            'mbytFunc :0-门诊结帐;1-住院结帐
            '只有出院结帐才检查 Or Not opt出院.Enabled
            '问题:45312
            If gbyt检查未执行 <> 0 Then
                strInfo = ExistWaitExe(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0))
                If strInfo <> "" Then
                    If gbyt检查未执行 = 1 Then
                        If MsgBox("发现病人" & mrsInfo!姓名 & "存在尚未执行完成的内容：" & _
                            vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                        End If
                    Else
                        MsgBox "发现病人" & mrsInfo!姓名 & "存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许出院结帐.", vbInformation, gstrSysName
                        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                    End If
                End If
            End If
            
            '问题:33048
            If gbyt检查未发药 <> 0 Then
                    strInfo = ExistWaitDrug(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0), 1)
                    If strInfo <> "" Then
                        If gbyt检查未发药 = 1 Then
                            If MsgBox("发现病人" & mrsInfo!姓名 & strInfo & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                            End If
                        Else
                            MsgBox "发现病人" & mrsInfo!姓名 & strInfo & vbCrLf & vbCrLf & "不允许出院结帐。", vbInformation, gstrSysName
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                        End If
                    End If
            End If
        End If
        
        If mbytFunc = 0 Then
            'mbytFunc :0-门诊结帐;1-住院结帐
            '只有门诊结帐才检查 Or Not opt出院.Enabled
            '问题:45312
            If gbyt检查门诊未执行 <> 0 Then
                strInfo = ExistWaitExe(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0), True)
                If strInfo <> "" Then
                    If gbyt检查门诊未执行 = 1 Then
                        If MsgBox("发现病人" & mrsInfo!姓名 & "存在尚未执行完成的内容：" & _
                            vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                        End If
                    Else
                        MsgBox "发现病人" & mrsInfo!姓名 & "存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许门诊结帐.", vbInformation, gstrSysName
                        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                    End If
                End If
            End If
            
            '问题:33048
            If gbyt检查门诊未发药 <> 0 Then
                    strInfo = ExistWaitDrug(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0), 1, True)
                    If strInfo <> "" Then
                        If gbyt检查门诊未发药 = 1 Then
                            If MsgBox("发现病人" & mrsInfo!姓名 & strInfo & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                            End If
                        Else
                            MsgBox "发现病人" & mrsInfo!姓名 & strInfo & vbCrLf & vbCrLf & "不允许门诊结帐。", vbInformation, gstrSysName
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                        End If
                    End If
            End If
        End If
        
        If gblnAutoOut And Not IsNull(mrsInfo!当前科室id) And opt出院.Value And mbytMCMode <> 1 _
            And (mbytFunc = 1 And mbytInState = 0) Then
            If GetUnAuditReFee(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0)) Then
                If MsgBox("病人" & txtPatient.Text & "存在已申请退费但未审核的记录,确定要进行出院结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        

        If Val(txtTotal.Text) <= 0 Then
            If MsgBox("病人实际没有可结费用,要结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
            End If
        ElseIf MsgBox("你确认要对该病人进行结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
        End If
        
        If gbln消费验证 Then
            curDeposit = 0
            With mshDeposit
                For i = 1 To .Rows - 1
                    curDeposit = curDeposit + Val(.TextMatrix(i, .ColIndex("冲预交")))
                Next
            End With
            strTime = IIf(mstrTime = "", mstrAllTime, mstrTime)
            If strTime = "0" And curDeposit <> 0 Then
                If Not zlDatabase.PatiIdentify(Me, glngSys, mrsInfo!病人ID, curDeposit) Then Exit Sub
            End If
        End If
        '30036
        If mty_ModulePara.bln结帐检查病历接收 And opt出院.Value = True _
            And (mbytFunc = 1 And mbytInState = 0) Then
            If IsCheck病历已接收(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = False Then
                If MsgBox("发现病人" & mrsInfo!姓名 & "没有进行病历审核," & _
                    vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                End If
                str病历原因 = ""
                If frmInputBox.InputBox(Me, "病历未接原因", "请输入病历未接原因信息:", 100, 3, True, False, str病历原因) = False Then
                    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                End If
            End If
        End If
        If mblnOneCard Then
            If GetOneCardCount > 1 Then
                MsgBox "不支持一次使用多种一卡通支付！", vbInformation, gstrSysName
                Exit Sub
            End If
            cur消费金额 = GetOneCardMoney
            If cur消费金额 <> 0 Then
                If mstrYBPati <> "" Then
                    MsgBox "不支持医保病人使用一卡通支付！", vbInformation, gstrSysName
                    Exit Sub
                End If
                If mobjICCard Is Nothing Or IsNull(mrsInfo!IC卡号) Then
                    MsgBox "使用一卡通支付必须先读卡！", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                curOneCard = mobjICCard.GetSpare
                If curOneCard < cur消费金额 Then
                    MsgBox "卡上余额" & Format(curOneCard, "0.00") & ",本次要求支付金额" & Format(cur消费金额, "0.00"), vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        bln打印退款收据 = False
        If mty_ModulePara.int退款票据 <> 0 And InStr(1, mstrPrivs, ";病人结帐退款收据;") > 0 Then
            '0-不打印,1-提示打印,2-不提示打印;'刘兴洪 问题:27776 日期:2010-02-04 16:49:03
            If mty_ModulePara.int退款票据 = 1 Then
               If MsgBox("你是否要打印“病人结帐退款收据”？" & vbCrLf & _
                       "   『是』：打印病人结帐退款收据" & vbCrLf & _
                       "   『否』：不打印病人结帐退款收据", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    bln打印退款收据 = True
                End If
            Else
                bln打印退款收据 = True
            End If
        End If
        
        
         '检查死亡情况:如果死亡则提示
'        '34681
'        If opt出院.Value Then
'            If zlCheckPatiIsDeath(Val(Nvl(mrsInfo!病人ID))) = True Then
'                If MsgBox("注意:" & vbCrLf & "    该病人已经死亡,是否继续结帐?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'            End If
'        End If

        '3.票据相关检查
        '问题:27559
        mblnPrint = True
        '保险病人根据使用类别来进行确认了
        Select Case mintInvoiceMode
        Case 0: mblnPrint = False '不打印
        Case 2  '自动打印
            If MsgBox("是否打印票据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                mblnPrint = False
            End If
        End Select
        bln打印费用明细 = False
        Select Case gbytFeePrintSet
        Case 1  '打印.
            If MsgBox("你是否要打印“病人结帐费用明细”？" & vbCrLf & _
                    "   『是』：打印病人结帐费用明细" & vbCrLf & _
                    "   『否』：不打印病人结帐费用明细", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    bln打印费用明细 = True
            End If
        Case 0  '不打印
        Case 2  '打印.但不提示
            bln打印费用明细 = True
        End Select
        
        '票据号码检查
        If mblnPrint Then
            If mbytFunc = 0 Then
                bytInvoiceKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, "0"))
            Else
                bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, "0"))
            End If
            If gblnStrictCtrl Then   '严格票据管理
                If Trim(txtInvoice.Text) = "" Then
                    MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
                mlng领用ID = GetInvoiceGroupID(IIf(bytInvoiceKind = 0, 3, 1), 1, mlng领用ID, mlngShareUseID, txtInvoice.Text, mstrUseType)
                If mlng领用ID <= 0 Then
                    Select Case mlng领用ID
                        Case 0 '操作失败
                        Case -1
                            MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                        Case -2
                            MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                        Case -3
                            MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入", vbInformation, gstrSysName
                            txtInvoice.SetFocus
                    End Select
                    Exit Sub
                End If
            Else
                If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                    MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            End If
        End If
        
        
        '三方卡检查
        If CheckThreePayDepositValied(objCard) = False Then
            If vsfMoney.Enabled And vsfMoney.Visible Then vsfMoney.SetFocus
            Exit Sub
        End If
        
        '4.存盘
        '-------------------------------------------------------------------------------------
        cmdOK.Enabled = False   '防止设置打印机弹出的非模态窗体,以及医保延时
        lngSaveID = SaveBalance(objCard, strNo, curDate, str病历原因)
        If lngSaveID = 0 Then cmdOK.Enabled = True: Exit Sub
        If mblnThreeDepositAfter And mblnThreeDepositAll Then
            frmBalanceDeposit.ShowMe Me, mlngModul, lngSaveID, Val(mrsInfo!病人ID), mblnThreeDepositAll, mblnDateMoved, mstr住院次数, mstrPepositDate, mint预交类别
        End If
        
        If bln打印退款收据 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_3", Me, "结帐ID=" & lngSaveID, 2)
        End If
        
        '票据打印
        If mblnPrint Then
       '问题:44332
RePrint:
            Dim strNotValiedNos As String
            Call frmPrint.ReportPrint(1, strNo, lngSaveID, mlng领用ID, mlngShareUseID, mstrUseType, txtInvoice.Text, curDate, txt缴款.Text, txt找补.Text, , mintInvoiceFormat, blnPrintBillEmpty, bytInvoiceKind + 1)
            Dim bytKind As Byte
            If mbytFunc = 0 Then
                bytKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, "0"))
            Else
                bytKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, "0"))
            End If
            If gblnStrictCtrl And blnPrintBillEmpty = False And _
                ((bytKind = 0 And InStr(1, mstrPrivs, ";收据打印;") > 0) _
                   Or (bytKind <> 0 And InStr(1, mstrPrivs, ";打印门诊收费票据;") > 0)) Then    'blnPrintBillEmpty:55052
                   '60155
                    If zlIsNotSucceedPrintBill(3, strNo, strNotValiedNos) = True Then
                            If MsgBox("结帐单据为[" & strNotValiedNos & "]的结帐票据打印未成功,是否重新打印结帐票据?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                    End If
            End If
        End If
        If bln打印费用明细 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me, "病人ID=" & Val(Nvl(mrsInfo!病人ID)), "结帐ID=" & lngSaveID, 2)
        End If
        
        '81697:李南春,2016/6/6,评价器
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.InPatiCashierAfter(Val(mrsInfo!病人ID), lngSaveID)
            Err.Clear
        End If
        
        '自动出院(出院结帐)
        If gblnAutoOut And Not IsNull(mrsInfo!当前科室id) And opt出院.Value And mbytMCMode <> 1 Then
            blnOut = True
            If Not IsNull(mrsInfo!险类) And Not MCPAR.未结清出院 Then
                Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, , , 2)
                If Not rsTmp Is Nothing Then
                    If Nvl(rsTmp!费用余额, 0) <> 0 Then blnOut = False
                End If
            End If
            
            If gbln医生允许才能出院 And blnOut Then
                If Not check医生下达出院医嘱(mrsInfo!病人ID, mrsInfo!主页ID) Then blnOut = False
            End If
            
            If blnOut Then
                frmAutoOut.mlng病人ID = mrsInfo!病人ID
                frmAutoOut.mlng主页ID = mrsInfo!主页ID
                frmAutoOut.mlngDepID = Val("" & mrsInfo!当前科室id)
                frmAutoOut.mint险类 = Nvl(mrsInfo!险类, 0)
                frmAutoOut.mstr性别 = Nvl(mrsInfo!性别)
                frmAutoOut.Show 1, Me
            End If
        End If
        
        '住院信息写卡:56615
        Call WriteZYInforToCard(Val(Nvl(mrsInfo!病人ID)), lngSaveID)
        If IsNull(mrsInfo!当前科室id) Then
            zlDatabase.SetPara "默认出院结帐", IIf(opt出院.Value = True, "1", "0"), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
        End If
        cmdOK.Enabled = True
    End If
    
    gblnOK = True
    
    
    '刘兴洪:
    cmdOK.Enabled = False
    cboNO.Text = ""
     
    '刘兴洪:27503
    If mty_ModulePara.bln结帐后不清信息 Then
        Set mrsInfo = New ADODB.Recordset
        If txtInvoice.Tag <> "" And txtInvoice.Text <> txtInvoice.Tag Then txtInvoice.Text = txtInvoice.Tag '主要是要保留信息,在确定后需要减判刑断
         Dim strTemp As String
         strTemp = txtInvoice.Text
        Call ReInitPatiInvoice
        txtInvoice.Text = strTemp   '主要是不要清空上次的发票,新的发票放在.tag中,在改变病人时,直接从这个地方读取
        mblnNotClearBill = True
        If mbytFunc = 0 Then
            txtPatient.Enabled = True: cmdYB.Enabled = True
        End If
    Else
        Call NewBill
        Call ReInitPatiInvoice(Not mblnStartFactUseType)
    End If
    sta.Panels(2) = "操作完毕，请输入其它病人标识！"
    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub InitBalanceSet(bln As Boolean)
'功能:门诊医保结帐状态时,相关控件控制
    chkCancel.Enabled = bln
    cmdYB.Enabled = bln
    txtPatient.Enabled = bln
    cmdPar.Enabled = bln
    txtPatiBegin.Enabled = bln
    txtPatiEnd.Enabled = bln
    
    If bln Then
        opt中途.Enabled = bln
        opt出院.Enabled = bln: opt出院.Caption = "出院结帐"
        txtTotal.Locked = (InStr(mstrPrivs, ";结帐设置;") = 0)
    Else
        opt中途.Enabled = bln
        opt出院.Enabled = Not bln: opt出院.Caption = "门诊结帐": opt出院.Value = True
        txtTotal.Locked = Not bln
        If MCPAR.门诊结算_结帐设置 Then cmdPar.Enabled = True
    End If
End Sub


Private Sub NewBill()
'功能:初始化结帐界面
    If mstrInNO = "" And mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    Set mrsInfo = New ADODB.Recordset '清除病人信息
    Set mtySquareCard.rsSquare = Nothing
    picThreeBalance.Visible = False
    txtOwe.Visible = True
    lblOwe.Visible = True
    mshDeposit.Visible = True
    vsfMoney.Visible = True
    

    mstrYBPati = "": mbytMCMode = 0
    mstrOneCard = ""
'''    Call zlClear结算卡
    Call ClearDetail
    Call AdjustBalance
    Call AdjustDeposit
    Call HideMoneyInfo
    Call InitBalanceCondition
    Call InitPatiVariable
    Call InitBalanceSet(True)
    
    pic状态.Visible = False: lbl状态.Caption = "":  lbl付款方式.Caption = ""
    mstr本次住院日期 = ""
    txtPatient.Text = "":    txtSex.Text = "":      txtOld.Text = ""
    txt费别.Text = "":       txt标识号.Text = "":   txtBed.Text = "": txt科室.Text = ""
    txtBegin.Text = "____-__-__": txtEnd.Text = "____-__-__"
    txtPatiBegin.Text = "____-__-__": txtPatiEnd.Text = "____-__-__":    txtPatiEnd.Tag = "____-__-__"
    txtDate.Text = "____-__-__ __:__:__": txt天数.Text = ""
    txt备注.Text = ""
    lblBed.Visible = False:     txtBed.Visible = False
    lbl标识号.Visible = False:  txt标识号.Visible = False
    lbl科室.Visible = False:    txt科室.Visible = False
    
    lblSpare.Caption = "预交余额:"
    lblSpare.Tag = ""
    mstr结帐主页IDs = ""
    mstr录入住院号 = ""
    sta.Panels(3).Text = ""
    lblDeposit.Caption = "冲预交:"
    lblDeposit.Tag = ""
    lblDelMoney.Caption = ""
    lblTicketCount.Caption = "预交款收据:"
    
    cmdOK.Enabled = True
    
    sta.Panels(2) = ""
    mstrForceNote = ""
End Sub
Private Sub cmdPar_Click()
    Dim blnAll As Boolean, i As Long
    If mrsInfo.State = 0 Then
        MsgBox "没有确定结帐病人,不能设置结帐参数！", vbExclamation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    End If
    
    With frmSetBalance
        .mstrUnAuditTime = mstrUnAuditTime
        .mblnNOCancel = mblnNOCancel
        .mlngInsure = IIf(MCPAR.允许结多次住院费用 Or Val(mstrTime) <> Val(Split(mstrAllTime, ",")(0)), 0, Val("" & mrsInfo!险类))
        .mlngPatient = mrsInfo!病人ID
        .mstrAllTime = mstrAllTime
        .mstrAllUnit = mstrAllUnit
        .mstrALLItem = mstrALLItem
        .mstrAllDiagnose = mstrAllDiagnose
        .mstrALLChargeType = mstrALLChargeType '34260
        .mstrAllClass = mstrAllClass
        .mMinDate = mMinDate
        .mMaxDate = mMaxDate
        .mbytKind = mbytKind
        .mstrTime = mstrTime
        .mbln门诊记帐结帐 = mbytMCMode = 1
        .mbytFunc = mbytFunc
        .Show 1, Me
    
        Me.Refresh
        If .mblnOk Then
            mblnSetPar = True
            '取参数处理
            Call InitPatiVariable
            '费用类型
            mstrClass = ""
            If Not .lstClass.Selected(0) Then
                For i = 1 To .lstClass.ListCount - 1
                    If .lstClass.Selected(i) Then
                        mstrClass = mstrClass & ",'" & .lstClass.List(i) & "'"
                    End If
                Next
            End If
            
            '收费类别:34260
            mstrChargeType = ""
            Dim objList As ListItem
            With .lvwChargeType
                If .ListItems("ALL").Checked = False Then
                    For Each objList In .ListItems
                        If objList.Key <> "ALL" And objList.Checked Then
                            mstrChargeType = mstrChargeType & ",'" & Mid(objList.Key, 2) & "'"
                        End If
                    Next
                End If
                
            End With
            
            mstrDiagnose = ""
            If .cboDiagnose.ListIndex <> 0 Then
                mstrDiagnose = mstrDiagnose & ",'" & .cboDiagnose.Text & "'"
            End If
            
            '婴儿费
            mbytBaby = .cboBabyFee.ListIndex
            
            '体检费
            mbytKind = 0
            If .chkKind(0).Value = 1 And .chkKind(1).Value = 1 Then
                mbytKind = 2
            Else
                If .chkKind(1).Value = 1 Then mbytKind = 1
            End If
            If mbytFunc = 0 Then
                mstrTime = ",0"
            Else
                If .lstTime.ListCount > 0 Then
                    blnAll = True
                    For i = 0 To .lstTime.ListCount - 1
                        If .lstTime.Selected(i) Then
                            mstrTime = mstrTime & "," & .lstTime.ItemData(i)
                        Else
                            blnAll = False
                        End If
                    Next
                    If blnAll And Not gbln仅用指定预交款 Then mstrTime = ""
                End If
             End If
            If .lstUnit.ListCount > 0 Then
                blnAll = True
                For i = 0 To .lstUnit.ListCount - 1
                    If .lstUnit.Selected(i) Then
                        mstrUnit = mstrUnit & "," & .lstUnit.ItemData(i)
                    Else
                        blnAll = False
                    End If
                Next
                If blnAll Then mstrUnit = ""
            End If
            If .lstItem.ListCount > 0 Then
                blnAll = True
                For i = 0 To .lstItem.ListCount - 1
                    If .lstItem.Selected(i) Then
                        mstrItem = mstrItem & ",'" & .lstItem.List(i) & "'"
                    Else
                        blnAll = False
                    End If
                Next
                If blnAll Then mstrItem = ""
            End If
            
            '用登记时间查询,发生时间显示
            '仅结体检费用时,不管期间
            If .chkKind(0).Value = 0 And .chkKind(1).Value = 1 Then
                mDateBegin = CDate("0:00:00")
                mDateEnd = CDate("0:00:00")
            Else
                mDateBegin = CDate(Format(.dtpBegin.Value, "yyyy-MM-dd 00:00:00"))
                mDateEnd = CDate(Format(.dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
            End If
                
            '显示结帐时间
            txtEnd.Text = Format(.dtpEnd.Value, txtEnd.Format)
            txtBegin.Text = Format(.dtpBegin.Value, txtBegin.Format)
            
            mstrTime = Mid(mstrTime, 2)
            mstrUnit = Mid(mstrUnit, 2)
            mstrItem = Mid(mstrItem, 2)
            mstrClass = Mid(mstrClass, 2)
            If mstrChargeType <> "" Then mstrChargeType = Mid(mstrChargeType, 2)   '34260
            If mstrDiagnose <> "" Then mstrDiagnose = Mid(mstrDiagnose, 2)
            
            '如果病人有多次住院费用未结，但只选择结某次住院费用，则根据该次住院信息来决定病人是否是医保病人
            If mstrTime <> "" And InStr(1, mstrTime, ",") = 0 And mrsInfo!主页ID <> mstrTime And InStr(1, mstrAllTime, ",") > 0 Then
                IDKIND.IDKIND = IDKIND.GetKindIndex("姓名")
                txtPatient.Text = "-" & mrsInfo!病人ID
                Call LoadPatientInfo(IDKIND.GetCurCard, False, 0, Val(mstrTime))
            End If
            
            If Not ShowBalance() Then
                cmdOK.Enabled = False
                MsgBox "该设置下病人没有需要结帐的费用！", vbInformation, gstrSysName
                If cmdPar.Visible And cmdPar.Enabled Then cmdPar.SetFocus
            Else
                If vsfMoney.Visible And vsfMoney.Enabled Then vsfMoney.SetFocus
            End If
        Else
            If mblnSetPar = False And Not IsNull(mrsInfo!险类) And MCPAR.结帐设置后调用接口 Then
                cmdOK.Enabled = False
            End If
        End If
    End With
End Sub

Private Sub OutputList(ByVal bytStyle As Byte)
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, lngRow As Long
    
    If mshDetail.TextMatrix(1, 0) = "" Then
        MsgBox "没有数据！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    objOut.Title.Text = "病人" & tabCard.SelectedItem.Caption
    If tabCard.SelectedItem.Index = 1 Then
        Set objOut.Title.Font = tabCard.Font
        Set objOut.Body = mshDetail
        
        lngRow = mshDetail.Row
    Else
        Set objOut.Title.Font = tabCard.Font
        Set objOut.Body = mshQuery
        
        lngRow = mshQuery.Row
        mshQuery_LeaveCell
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "单据号:" & cboNO.Text
    objRow.Add "实际号:" & txtInvoice.Text
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "病人:" & txtPatient.Text
    objRow.Add "住院号:" & txt标识号.Text
    objRow.Add "合计:" & txtTotal.Text
    objOut.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "打印时间:" & Format(zlDatabase.Currentdate, "YYYY-MM-DD hh:mm:ss")
    objRow.Add "结帐时间:" & txtDate.Text
    objOut.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    If mbytInState = 0 Then
        objRow.Add "备注:未保存"
    ElseIf mbytInState = 1 Then
        If mblnViewCancel Then
            objRow.Add "备注:作废单"
        Else
            objRow.Add "备注:"
        End If
    End If
    objOut.BelowAppRows.Add objRow
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
        
    If tabCard.SelectedItem.Index = 1 Then
        mshDetail.Row = lngRow
    Else
        mshQuery.Row = lngRow
        mshQuery_EnterCell
    End If
End Sub

Private Sub Form_Activate()
    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    If mblnUnload = True Then Unload Me: Exit Sub
    
    mblnFirst = False
    If mstrInNO = "" And mbytInState = 0 And gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    If mbytInState = 1 Then
        If cmdCancel.Visible And cmdCancel.Enabled Then cmdCancel.SetFocus
    ElseIf mstrInNO <> "" Then
        '作废时
        If txtPatient.Text = "" Then Unload Me: Exit Sub
        If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If txtMoney.Visible Then
                txtMoney.Visible = False
                If txtMoney.Left < fraBalance.Left Then
                    mshDetail.SetFocus
                Else
                    mshDeposit.SetFocus
                End If
            Else
                '取消按钮
                If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus: Call cmdCancel_Click
            End If
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus: Call cmdOK_Click
        Case vbKeyF4
            If Shift = vbCtrlMask Then
                If IDKIND.Enabled Then
                    Dim intIndex As Integer
                    intIndex = IDKIND.GetKindIndex("IC卡号")
                    If intIndex <= 0 Then Exit Sub
                    IDKIND.IDKIND = intIndex: Call IDKind_Click(IDKIND.GetCurCard)
                End If
            ElseIf Me.ActiveControl Is txtPatient Then
                If IDKIND.Enabled Then
                    If Shift = vbShiftMask Then
                        IDKIND.IDKIND = IIf(IDKIND.IDKIND = 0, UBound(Split(IDKIND.IDKindStr, ";")), IDKIND.IDKIND - 1)
                    Else
                        IDKIND.IDKIND = IIf(IDKIND.IDKIND = UBound(Split(IDKIND.IDKindStr, ";")), 0, IDKIND.IDKIND + 1)
                    End If
                End If
            End If
        Case vbKeyF6
            If cmdYB.Enabled And cmdYB.Visible Then cmdYB.SetFocus: Call cmdYB_Click
        Case vbKeyF8 '退号快捷
            chkCancel.Value = IIf(chkCancel.Value = 1, 0, 1)
        Case vbKeyF9 '结帐设置
            If cmdPar.Enabled And cmdPar.Visible Then cmdPar.SetFocus: Call cmdPar_Click
        Case vbKeyF11 '定位到病人输入框
            If Not txtPatient.Locked And txtPatient.Enabled Then txtPatient.SetFocus
        Case vbKeyF12 '定位到单号框
            If Not cboNO.Locked And cboNO.Enabled Then cboNO.SetFocus
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strTmp As String
    mblnFirst = True
    mblnUnload = False
    glngFormW = 11565: glngFormH = 8535
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
       
    mint预交类别 = 2 '缺省为住院预交
    Call RestoreWinState(Me, App.ProductName)
    gblnOK = False
    
    If mbytInState = 0 Then
        Set mrsOneCard = GetOneCard
        mblnOneCard = mrsOneCard.RecordCount > 0
    End If
    If InStr(1, mstrPrivs, ";费用打折结算;") = 0 Then
        strTmp = "1,2,3,4,5,9"    '7,8:问题:48810
    Else
        strTmp = "1,2,3,4,5,6,9"  '7,8:问题:48810
    End If
    
    If InStr(1, mstrPrivs, ";仅用预交结帐;") = 0 Then
        Set mrs结算方式 = Get结算方式("结帐", strTmp)
        If mrs结算方式.RecordCount = 0 Then
            MsgBox "未设置结帐场合可用的结算方式。", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        Set mrs结算方式 = Get结算方式("结帐", "3,4", True)
    End If
    
    mstrCardPrivs = GetPrivFunc(glngSys, 1151)
    
    Call InitFace
    
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
    End If
 

    
    If mbytInState = 1 Then                 '查看
        If Not ReadBalance(mstrInNO) Then mblnUnload = True: Exit Sub
    ElseIf mstrInNO <> "" Then        '作废
        mblnOutUse = True
        chkCancel.Value = 1     '调用Click事件
        cboNO.Text = mstrInNO
        cboNO_KeyPress (13)
        mblnOutUse = False
    Else '执行结帐
'        If Not CheckErrorItem Then
'            MsgBox "系统中尚未设置有效的误差处理项目，请先到基础参数设置中设置。", vbInformation, gstrSysName
'            mblnUnload = True:  Exit Sub
'        End If
        
        mintPatientRange = Val(zlDatabase.GetPara("显示结清病人", glngSys, mlngModul, 0))
        If mlngPatientID <> 0 Then
            txtPatient.Text = "-" & mlngPatientID
            mstrTime = mstr主页Id
            Call txtPatient_KeyPress(vbKeyReturn)
            If Val(mstr主页Id) = "0" Then cmdYB.Enabled = True
            If mrsInfo.State = 0 Then mblnUnload = True: Exit Sub
        End If
    End If
    
    '问题:47798
    If mbytInState = 0 Then
        Call GetRegisterItem(g私有模块, Me.Name, "idkind", strTmp)
        Err = 0: On Error Resume Next
        mblnNotClick = True
        IDKIND.IDKIND = Val(strTmp)
        mblnNotClick = False
        Err = 0: On Error GoTo 0
    End If
End Sub

Private Sub RefreshFact(Optional blnRed As Boolean)
    Dim bytInvoiceKind As Byte
    '功能：刷新收费票据号
    If mintInvoiceMode = 0 Then Exit Sub
    
    If mbytFunc = 0 Then
        bytInvoiceKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, "0"))
    Else
        bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, "0"))
    End If
    
    If gblnStrictCtrl Then
        mlng领用ID = CheckUsedBill(IIf(bytInvoiceKind = 0, 3, 1), IIf(mlng领用ID > 0, mlng领用ID, mlngShareUseID), , mstrUseType)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            If blnRed Then
                mstrInvoice = ""
            Else
                txtInvoice.Text = ""
                txtInvoice.Tag = ""
            End If
        Else
            '严格：取下一个号码
            If blnRed Then
                mstrInvoice = GetNextBill(mlng领用ID)
            Else
                txtInvoice.Text = GetNextBill(mlng领用ID)
                txtInvoice.Tag = txtInvoice.Text
            End If
        End If
    Else
        '松散：取下一个号码
        If blnRed Then
            mstrInvoice = IncStr(UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, "")))
        Else
            txtInvoice.Text = IncStr(UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, "")))
            txtInvoice.Tag = txtInvoice.Text
        End If
    End If
End Sub

Private Sub InitFace()
    Dim i As Long, bytInvoiceKind As Byte
    
    If mbytInState = 1 Then
         lblTitle.Caption = gstrUnitName & "病人结帐单"
    Else
         lblTitle.Caption = gstrUnitName & IIf(mbytFunc = 0, "门诊病人结帐单", "住院病人结帐单")
    End If
    
    sta.Panels("LocalParSet").Visible = mlngPatientID <> 0  '病人费用查询中调用时,提供本地参数设置
    
    Call zlInitModulePara
    Call initCardSquareData
    
    If mbytFunc = 0 Then
        bytInvoiceKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, "0"))
    Else
        bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, "0"))
    End If
    
    mblnStartFactUseType = zlStartFactUseType(IIf(bytInvoiceKind = 0, 3, 1))
    
    If Not (mbytInState = 0 And mstrInNO <> "") Then Call NewBill    '作废时在chkCancel.Value = 1时调用
    chkCancel.Visible = (mbytInState = 0 And (InStr(";" & mstrPrivs, ";结帐作废;") > 0))
         
    txtPatient.Width = txtPatient.Width + 400
    
    IDKIND.Enabled = (mbytInState = 0 And mstrInNO = "")
    If mbytInState = 0 And mstrInNO = "" Then
        Call ReInitPatiInvoice(Not mblnStartFactUseType)
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
        If InStr(mstrPrivs, ";保险结算;") > 0 Then
            cmdYB.Visible = True
            
            '成都老版医保支持门诊和住院两种身份验证模式
            mblnMC_TwoMode = InStr("," & GetSetting("ZLSOFT", "公共全局", "本地支持的医保", "") & ",", ",20,") > 0
            If mblnMC_TwoMode Then
                cmdYB.Caption = "刷"
                txtPatient.Width = txtPatient.Width - 400
                cmdYB.Left = txtPatient.Left + txtPatient.Width + 10
                cmdYB.Top = fraPatient.Top + 180
                cmdYB.Width = 400
                pic状态.Left = txtPatient.Left
            ElseIf InStr(mstrPrivs, ";门诊费用结帐;") = 0 Or mbytFunc = 1 Then    'mbytFunc=1:住院结算
                cmdYB.Visible = False
                pic状态.Left = txtPatient.Left
            End If
        Else
            cmdYB.Visible = False
            pic状态.Left = txtPatient.Left
        End If
    
        If InStr(mstrPrivs, ";结帐设置;") = 0 Then
            cmdPar.Visible = False
            txtTotal.Locked = True
            opt中途.Left = opt中途.Left - cmdPar.Width / 2
            opt出院.Left = opt出院.Left - cmdPar.Width / 2
        End If
        cboNO.Text = ""
        opt出院.Visible = True
        opt中途.Visible = True
        cmd结算卡.Visible = False
        Call Init预交类别
    ElseIf mbytInState = 1 Then
        If mblnViewCancel Then lblFlag.Visible = True
        cmdOK.Visible = False
        cmdCancel.Caption = "退出(&X)"
        txtPatient.Locked = True
        txtTotal.Locked = True
        
        fra找补.Visible = False
        txt备注.Enabled = False: lbl备注.Enabled = False
        cmdPar.Visible = False
        opt出院.Visible = False
        opt中途.Visible = False
        
        fra费用期间.Top = fra费用期间.Top - cmdPar.Height
        fra住院期间.Top = fra住院期间.Top - cmdPar.Height
        fra结帐时间.Top = fra结帐时间.Top - cmdPar.Height
        fraDate.Height = fraDate.Height - cmdPar.Height
        fraBalance.Top = fraBalance.Top - cmdPar.Height
        
        fraTitle.Enabled = False
        fra住院期间.Enabled = False
        Call SetDisibleColor
        cmd结算卡.Visible = False
    End If

End Sub
Private Sub SetSortMoneyData(ByVal BytType As Byte, ByVal blnHaveMoeny As Boolean, ByVal bytEdit As Byte, _
    ByRef k As Integer, ByRef ArrSort() As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据参数,设置结算方式显示顺序数据
    '入参:bytType-类型(0-非医保;1-医保)
    '       blnHaveMoeny-true:有金额;False;无金额
    '       bytEdit-0-不区分编辑;1允许编辑;2不可编辑
    '出参:K-返回最后一次顺序编号
    '       ArrSort-返回排序数据
    '返回:
    '编制:刘兴洪
    '日期:2010-09-26 15:03:35
    '问题:32322
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, bytTemp As Byte   '0非医保;>1医保
    Dim blnTempMoney As Boolean, bytTempEdit As Byte
    With vsfMoney
        For i = 1 To .Rows - 1
            bytTemp = IIf(InStr(1, ",3,4,", Trim(.TextMatrix(i, .ColIndex("性质")))) = 0, 0, 1)
            blnTempMoney = Val(.TextMatrix(i, .ColIndex("金额"))) <> 0
            bytTempEdit = IIf(bytEdit = 0, 0, IIf(.RowData(i) = 0, 1, 2))
            If bytTemp = BytType And blnHaveMoeny = blnTempMoney And bytTempEdit = bytEdit And Val(.TextMatrix(i, .ColIndex("性质"))) <> 9 Then
                '满足条件
                For j = 0 To .Cols - 1
                    ArrSort(k, j) = .TextMatrix(i, j)
                Next
                '附加数据
                ArrSort(k, .Cols) = .RowData(i)
                .Row = i:  .Col = 0
                ArrSort(k, .Cols + 1) = IIf(.CellFontBold, 1, 0)
                ArrSort(k, .Cols + 2) = .CellForeColor
                k = k + 1
            End If
        Next
    End With
End Sub

Private Sub SetSortMoneyError(ByRef k As Integer, ByRef ArrSort() As String)
    Dim i As Long, j As Long
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("性质"))) = 9 Then
                '满足条件
                For j = 0 To .Cols - 1
                    ArrSort(k, j) = .TextMatrix(i, j)
                Next
                '附加数据
                ArrSort(k, .Cols) = .RowData(i)
                .Row = i:  .Col = 0
                ArrSort(k, .Cols + 1) = IIf(.CellFontBold, 1, 0)
                ArrSort(k, .Cols + 2) = .CellForeColor
                k = k + 1
            End If
        Next
    End With
End Sub

Private Sub SortMoney()
'功能：调整结算方式表列表,使有金额的排在前面
'说明：同类中原有顺序不变
    Dim arrCell() As String, blnRedraw As Boolean
    Dim i As Long, j As Integer, k As Integer
    Dim lngRow As Long, lngCol As Long
    Dim varData As Variant
    Dim ArrTemp() As String
    
    ReDim ArrTemp(0 To vsfMoney.Cols + 2)
    ReDim arrCell(1 To vsfMoney.Rows, 0 To vsfMoney.Cols + 2)
    lngRow = vsfMoney.Row: lngCol = vsfMoney.Col
    blnRedraw = vsfMoney.Redraw
    vsfMoney.Redraw = False
    '问题:32322

    k = 1
    Call SetSortMoneyError(k, arrCell)
    varData = Split(gstr结算方式显示顺序, ";")
    '非医保结算-有金额;非医保结算-无金额;医保结算-有金额且允许修改;医保结算-无金额且允许修改;医保结算-有金额且不允许修改;医保结算-无金额且不允许修改
    For i = 0 To UBound(varData)
        Select Case varData(i)
        Case "非医保结算-有金额"
            Call SetSortMoneyData(0, True, 0, k, arrCell)
        Case "非医保结算-无金额"
            Call SetSortMoneyData(0, False, 0, k, arrCell)
        Case "医保结算-有金额且允许修改"
            Call SetSortMoneyData(1, True, 1, k, arrCell)
        Case "医保结算-无金额且允许修改"
            Call SetSortMoneyData(1, False, 1, k, arrCell)
        Case "医保结算-有金额且不允许修改"
            Call SetSortMoneyData(1, True, 2, k, arrCell)
        Case "医保结算-无金额且不允许修改"
            Call SetSortMoneyData(1, False, 2, k, arrCell)
        Case Else
        End Select
    Next
    '预防某些结算方式不加载,需进行数据修正
    Dim blnFind As Boolean
    With vsfMoney
        For i = 1 To .Rows - 1
            blnFind = False
            For j = 1 To UBound(arrCell)
                If .TextMatrix(i, .ColIndex("结算方式")) = arrCell(j, .ColIndex("结算方式")) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If blnFind = False Then
                '未找到数据,需要重新加载上去
                For j = 0 To vsfMoney.Cols - 1
                    arrCell(k, j) = vsfMoney.TextMatrix(i, j)
                Next
                '附加数据
                arrCell(k, vsfMoney.Cols) = vsfMoney.RowData(i)
                vsfMoney.Row = i: vsfMoney.Col = 0
                arrCell(k, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
                arrCell(k, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
                k = k + 1
            End If
        Next
    End With
  
    '重新填写表格
    For i = 1 To vsfMoney.Rows - 1
        For j = 0 To vsfMoney.Cols - 1
            vsfMoney.TextMatrix(i, j) = arrCell(i, j)
        Next
        
        '附加数据
        vsfMoney.RowData(i) = Val(arrCell(i, vsfMoney.Cols))
        vsfMoney.Row = i: vsfMoney.Col = 0
        vsfMoney.CellFontBold = IIf(Val(arrCell(i, vsfMoney.Cols + 1)) = 1, True, False)
        vsfMoney.CellForeColor = Val(arrCell(i, vsfMoney.Cols + 2))
    Next
    vsfMoney.Row = lngRow: vsfMoney.Col = lngCol
    vsfMoney.Redraw = blnRedraw
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("金额"))) = 0 And Val(.RowData(i)) = 1 Then
                .RowHidden(i) = True
            Else
                .RowHidden(i) = False
            End If
        Next i
        .Refresh
    End With
End Sub

Private Sub AdjustBalance()
'功能：调整结算项目列表
    Dim strSql As String, i As Long
    Dim intDef As Integer, lngW As Long, blnTmp As Boolean
            
    mbln个帐结算 = False
    mcur个帐余额 = 0
    mcur个帐限额 = 0
    mcur个帐透支 = 0
    mstr缺省结算 = ""
    mstrBalance = ""
    
    mrs结算方式.Filter = ""
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!险类) And mbytMCMode <> 1 Then
            If Not MCPAR.结算使用个人帐户 Then mrs结算方式.Filter = "性质<>3"
        End If
    End If
    
    Call InitBalanceGrid(vsfMoney)
    
    With vsfMoney
        blnTmp = .Redraw
        .Redraw = False
        '设置可用结算方式
        If Not mrs结算方式.EOF Then
            .Rows = IIf(mrs结算方式.RecordCount = 0, 1, mrs结算方式.RecordCount) + 1
            For i = 1 To mrs结算方式.RecordCount
                .TextMatrix(i, .ColIndex("结算方式")) = mrs结算方式!名称
                .TextMatrix(i, .ColIndex("性质")) = mrs结算方式!性质
                
                .Row = i: .Col = 0
                .CellForeColor = vbBlack
                '缺省方式粗体显示
                If mrs结算方式!缺省 = 1 Then
                    mstr缺省结算 = mrs结算方式!名称
                    .Row = i: .Col = 0
                    .CellFontBold = True
                    intDef = .Row
                ElseIf InStr(",3,4,", mrs结算方式!性质) > 0 Then
                    .Row = i: .Col = 0: .CellForeColor = vbBlue
                ElseIf InStr(",9,", mrs结算方式!性质) > 0 Then
                    .Row = i: .Col = 0
                    .CellForeColor = vbRed
                End If
                mrs结算方式.MoveNext
            Next
        End If
        lngW = .Width - 75
        If .Rows > .Height \ .RowHeight(0) Then lngW = lngW - 250
        .ColWidth(.ColIndex("结算方式")) = lngW * 0.3
        .ColWidth(.ColIndex("金额")) = lngW * 0.3
        .ColWidth(.ColIndex("结算号码")) = lngW * 0.4
        
        .Col = 1
        For i = 1 To .Rows - 1
            .Row = i
            .CellBackColor = txtMoney.BackColor
            If InStr(",3,4,", Val(.TextMatrix(i, .ColIndex("性质")))) > 0 Then
                .RowData(i) = 1 '医保结算缺省为不可编辑
            ElseIf Val(.TextMatrix(i, .ColIndex("性质"))) = 8 Then
                .RowData(i) = 1 '消费卡不可编辑
            ElseIf Val(.TextMatrix(i, .ColIndex("性质"))) = 9 Then
                .RowData(i) = 1 '误差费不可编辑
            ElseIf Val(.TextMatrix(i, .ColIndex("性质"))) = 999 Then
                .RowData(i) = 999 '冲预交结算方式
            Else
                .RowData(i) = 0 '普通结算缺省为可以编辑
            End If
            .TextMatrix(i, .ColIndex("金额")) = "0.00"
            .TextMatrix(i, .ColIndex("结算号码")) = ""
        Next
        If intDef > 0 Then .Row = intDef
        
        txtOwe.Text = "0.00"
        
        .Redraw = blnTmp
    End With
End Sub

Private Sub ClearDetail(Optional blnSetPatiForeColor As Boolean = True)
    Dim i As Long, j As Long
    With mshDetail
        .Redraw = False
        .Clear
        .ClearStructure
        .Rows = 2: .Cols = 2
        .ColWidth(0) = 1000: .ColWidth(1) = 1000
        .Row = 1: .Col = 0
        .Redraw = True
    End With
    txt缴款.Text = "0.00"
    txt找补.Text = "0.00"
    txtTotal.Text = gstrDec
    txtTotal.Tag = gstrDec
    mstrDec = gstrDec
    mcurTotal = 0: mcur误差金额 = 0
    If blnSetPatiForeColor Then txtPatient.ForeColor = Me.ForeColor
    With mshQuery
        .Tag = ""
        .Redraw = False
        .Clear
        .ClearStructure
        .Rows = 2
        .Cols = 2
        .Row = 1: .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Form_Resize()
    Dim lngCancelW As Long
    Dim lngInsureH As Long
    Dim sngTemp As Single
    
    On Error Resume Next
    If WindowState = 1 Then Exit Sub
    
    If chkCancel.Visible Or lblFlag.Visible Then lngCancelW = chkCancel.Width
    fraTitle.Width = Me.ScaleWidth - fraTitle.Left
    chkCancel.Left = fraTitle.Width - chkCancel.Width - 60
    lblFlag.Left = chkCancel.Left + (chkCancel.Width - lblFlag.Width) / 2
    
    cboNO.Left = fraTitle.Width - lngCancelW - 60 - cboNO.Width - 30
    lblNO.Left = cboNO.Left - lblNO.Width - 45
    txtInvoice.Left = lblNO.Left - txtInvoice.Width - 200
    lblFact.Left = txtInvoice.Left - lblFact.Width - 45
    
    fraPatient.Width = fraTitle.Width
    
    fraDate.Left = Me.ScaleWidth - fraDate.Width
    fraBalance.Left = fraDate.Left
    
    cmdCancel.Left = fraDate.Left + fraDate.Width - cmdCancel.Width - 50
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    
    tabCard.Width = Me.ScaleWidth - fraDate.Width - tabCard.Left - 30
    
    mshQuery.Width = tabCard.Width - mshQuery.Left - 60
    mshDetail.Width = tabCard.Width - mshDetail.Left - 60
    tabCard.Height = Me.ScaleHeight - tabCard.Top - fraAppend.Height - sta.Height - (fra备注.Height - 50)
    With fra备注
        .Width = tabCard.Width
        .Top = tabCard.Top + tabCard.Height - 50
        fraAppend.Top = .Top + .Height - 50
        txt备注.Width = .Width - txt备注.Left - .Left - 50
        fraBalance.Height = .Top + .Height - fraBalance.Top
    End With
    
    'fraAppend.Top = tabCard.Top + tabCard.Height
    mshDetail.Height = tabCard.Height - 480
    mshQuery.Height = tabCard.Height - 480
    
    'fraBalance.Height = tabCard.Top + tabCard.Height - fraBalance.Top
    
    cmdOK.Top = fraAppend.Top + (fraAppend.Height - cmdOK.Height) / 2
    cmdCancel.Top = cmdOK.Top
    cmd结算卡.Top = cmdOK.Top
    lngInsureH = IIf(lbl医保基金.Visible, lbl医保基金.Height + 30, 30)
    
    mshDeposit.Height = (fraBalance.Height - lblDeposit.Height - txtOwe.Height - 240) * 0.45

    lbl医保基金.Top = mshDeposit.Top + mshDeposit.Height + 15
    lbl个人帐户.Top = lbl医保基金.Top
    vsfMoney.Top = mshDeposit.Top + mshDeposit.Height + lngInsureH
    
    sngTemp = fraBalance.Height - lblDeposit.Height - txtOwe.Height _
                - IIf(lblDelMoney.Visible, lblDelMoney.Height + 240 * 2, 240) - 240
    
    vsfMoney.Height = sngTemp * 0.55 - lngInsureH
    mshDeposit.Width = vsfMoney.Width
    lblDelMoney.Top = vsfMoney.Top + vsfMoney.Height + 100
    cmdReturnCash.Top = lblDelMoney.Top - 60
    
    txtOwe.Top = IIf(lblDelMoney.Visible, lblDelMoney.Top + lblDelMoney.Height + 50, vsfMoney.Top + vsfMoney.Height) + 15
    lblOwe.Top = txtOwe.Top + (txtOwe.Height - lblOwe.Height) / 2
    lblTicketCount.Top = lblOwe.Top
    
    
    fraAppend.Width = fra找补.Width + lblTotal.Width + txtTotal.Width + 200
    fra找补.Left = fraAppend.Width - fra找补.Width
    
    picThreeBalance.Height = fraBalance.Height - 200
    picThreeBalance.Width = vsfMoney.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnView As Boolean
    If mbytInState = 0 And mstrYBPati <> "" And mstrInNO = "" Then
        If MsgBox("当前正在对医保病人结帐，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
        If YBIdentifyCancel = False Then        '取消医保病人身份验证,返回假时不退出
            Cancel = 1: Exit Sub
        End If
    End If
    
    blnView = mstrInNO <> "" Or chkCancel.Value = 1 Or cboNO.Text <> ""
    zl_vsGrid_Para_Save mlngModul, mshDeposit, Me.Caption, "预交列表" & IIf(blnView, "1", "0")
    
    '清除入口参数
    mlngPatientID = 0
    mbytInState = 0
    mblnViewCancel = False
    mstrInNO = ""
    mblnNOMoved = False
    mlngBillID = 0
    mstrPrivs = ""
    mstrForceNote = ""
    
    mstr缺省结算 = "": mstrBalance = ""
    mstrYBPati = "":   mbytMCMode = 0:    mintInsure = 0
    mlng领用ID = 0:    mcurTotal = 0:     mcur误差金额 = 0
    mcur个帐余额 = 0:  mcur个帐限额 = 0:  mcur个帐透支 = 0
    mbln门诊转住院 = False: mstr主页Id = "": mstrPepositDate = ""
    Call InitBalanceCondition
    Call InitPatiVariable
        
    Set mrsBalance = Nothing
    Set mrsDeposit = Nothing
    Set mrsInfo = New ADODB.Recordset
    
    Unload frmSetBalance
    
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    
    Call SaveWinState(Me, App.ProductName)
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    '问题:47798
    If mbytInState = 0 Then
        Call SaveRegisterItem(g私有模块, Me.Name, "idkind", IDKIND.IDKIND)
    End If

End Sub

Private Sub mshDeposit_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim blnView As Boolean
    blnView = mstrInNO <> "" Or cboNO.Text <> "" Or chkCancel.Value = 1
    zl_vsGrid_Para_Save mlngModul, mshDeposit, Me.Caption, "预交列表" & IIf(blnView, "1", "0")
End Sub

Private Sub mshDeposit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button <> 2 Then Exit Sub
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    
    If mbytFunc = 0 And mbytInState <> 1 Then
        Me.PopupMenu Me.mnuPopu, 0
    Else
        Me.PopupMenu Me.mnuColsVisible, 0
    End If
End Sub

Private Sub mnuViewToolCols_Click(Index As Integer)
    Dim i As Integer, j As Integer
    Dim strKey As String, blnShow As Boolean
    Dim blnView As Boolean
    
    mnuViewToolCols(Index).Checked = Not mnuViewToolCols(Index).Checked
    
    For i = 0 To mnuViewToolCols.UBound
        If mnuViewToolCols(i).Visible And mnuViewToolCols(i).Checked Then j = j + 1
    Next
    If j < 2 Then
        sta.Panels(2).Text = "要求至少保留两列显示!"
        mnuViewToolCols(Index).Checked = True
    End If
    
    strKey = Split(mnuViewToolCols(Index).Caption & "(", "(")(0)
    blnShow = mnuViewToolCols(Index).Checked
    With mshDeposit
        If .ColIndex(strKey) >= 0 Then
            .ColHidden(.ColIndex(strKey)) = Not blnShow
            
        End If
    End With
    blnView = mstrInNO <> "" Or cboNO.Text <> "" Or chkCancel.Value = 1
    zl_vsGrid_Para_Save mlngModul, mshDeposit, Me.Caption, "预交列表" & IIf(blnView, "1", "0")
End Sub

Private Sub mnuFileExcel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFilePrintSetup_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileZero_Click()
    mnuFileZero.Checked = Not mnuFileZero.Checked
    Call LoadCardData
End Sub
Private Sub picThreeBalance_Resize()
    Dim sngTop As Single
    Err = 0: On Error Resume Next
    With picThreeBalance
        picDelDeposit.Left = lblDelDeposit.Left + lblDelDeposit.Width + 15
        picDelDeposit.Width = .ScaleWidth - picDelDeposit.Left
        lbl误差.Top = picDelDeposit.Top + picDelDeposit.Height + 50
        sngTop = picDelDeposit.Top + picDelDeposit.Height + 50
        If lbl误差.Visible Then sngTop = lbl误差.Top + lbl误差.Height + 50
        vsDelBalance.Top = sngTop
        vsDelBalance.Left = .ScaleLeft
        vsDelBalance.Width = .ScaleWidth - vsDelBalance.Left
        vsDelBalance.Height = .ScaleHeight - vsDelBalance.Top
    End With
End Sub

Private Sub vsfMoney_DblClick()
    If Not txtMoney.Visible And vsfMoney.Row >= 1 And vsfMoney.Col > 0 _
        And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
                        
        '不可修改的结算方式
        If vsfMoney.RowData(vsfMoney.Row) = 1 Or vsfMoney.RowData(vsfMoney.Row) = 999 Then Exit Sub

        With txtMoney
            .MaxLength = IIf(vsfMoney.Col = 2, 30, 10)
            .Left = fraBalance.Left + vsfMoney.Left + vsfMoney.CellLeft + 30
            .Top = fraBalance.Top + vsfMoney.Top + vsfMoney.CellTop + (vsfMoney.CellHeight - txtMoney.Height) / 2 + 30
            .Width = vsfMoney.CellWidth - 60
            .ForeColor = vsfMoney.ForeColor
            .BackColor = vsfMoney.BackColor
            .Alignment = IIf(vsfMoney.Col = 2, 0, 1)
            .Text = vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub vsfMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsfMoney
        If .Row >= 1 Then
            If .Col < .Cols - 2 Then
                .Col = .Col + 1
            Else
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    .Col = 1
                    If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                        .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
                    End If
                Else
                    If txt备注.Visible And txt备注.Enabled Then
                        txt备注.SetFocus
                    ElseIf Get应缴 > 0 And txt缴款.Visible Then
                        txt缴款.SetFocus
                    ElseIf cmdOK.Visible And cmdOK.Enabled Then
                        cmdOK.SetFocus
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfMoney_KeyPress(KeyAscii As Integer)
    If Not txtMoney.Visible And vsfMoney.Row >= 1 And vsfMoney.Col > 0 _
        And KeyAscii <> 13 And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
        
        '不可修改的结算方式
        If vsfMoney.RowData(vsfMoney.Row) = 1 Or vsfMoney.RowData(vsfMoney.Row) = 999 Then Exit Sub
        
        '结算号码没限制
        If vsfMoney.Col = 1 Then If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
        With txtMoney
            .MaxLength = IIf(vsfMoney.Col = 2, 30, 10)
            .Left = fraBalance.Left + vsfMoney.Left + vsfMoney.CellLeft + 30
            .Top = fraBalance.Top + vsfMoney.Top + vsfMoney.CellTop + (vsfMoney.CellHeight - txtMoney.Height) / 2 + 30
            .Width = vsfMoney.CellWidth - 60
            .ForeColor = vsfMoney.ForeColor
            .BackColor = vsfMoney.BackColor
            .Alignment = IIf(vsfMoney.Col = 2, 0, 1)
            .Text = Chr(KeyAscii)
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshDetail_DblClick()
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    If InStr(mstrPrivs, ";结帐设置;") = 0 Then Exit Sub
    If mshDetail.Col <> GetColNum("结帐金额") Then Exit Sub
     
    If Not txtMoney.Visible And mshDetail.Row >= 1 _
        And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
        If IsNull(mrsInfo!医保号) Then
            With txtMoney
                .Left = mshDetail.Left + mshDetail.CellLeft + 15
                .Top = mshDetail.Top + mshDetail.CellTop + (mshDetail.CellHeight - txtMoney.Height) / 2 - 15
                .Width = mshDetail.CellWidth - 60
                .ForeColor = mshDetail.CellForeColor
                .BackColor = mshDetail.CellBackColor
                .Alignment = 1
                .Text = mshDetail.TextMatrix(mshDetail.Row, mshDetail.Col)
                .SelStart = 0: .SelLength = Len(.Text)
                .ZOrder: .Visible = True
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If mshDetail.Row >= 1 Then
            If mshDetail.Col = GetColNum("结帐金额") Then
                If mshDetail.Row < mshDetail.Rows - 1 Then
                    mshDetail.Row = mshDetail.Row + 1
                    If mshDetail.Row - (mshDetail.Height \ mshDetail.RowHeight(0) - 2) > 1 Then
                        mshDetail.TopRow = mshDetail.Row - (mshDetail.Height \ mshDetail.RowHeight(1) - 2)
                    End If
                Else
                    mshDeposit.SetFocus
                End If
            Else
                mshDetail.Col = mshDetail.Col + 1
            End If
        End If
    End If
End Sub

Private Sub mshDetail_KeyPress(KeyAscii As Integer)
    If InStr(mstrPrivs, ";结帐设置;") = 0 Then Exit Sub
    If mshDetail.Col <> GetColNum("结帐金额") Then Exit Sub
    
    If Not txtMoney.Visible And mshDetail.Row >= 1 _
        And KeyAscii <> 13 And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
        If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        If IsNull(mrsInfo!医保号) Then
            With txtMoney
                .Left = mshDetail.Left + mshDetail.CellLeft + 15
                .Top = mshDetail.Top + mshDetail.CellTop + (mshDetail.CellHeight - txtMoney.Height) / 2 - 15
                .Width = mshDetail.CellWidth - 60
                .ForeColor = mshDetail.CellForeColor
                .BackColor = mshDetail.CellBackColor
                .Alignment = 1
                .Text = Chr(KeyAscii)
                .SelStart = 1
                .ZOrder: .Visible = True
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub mshDetail_LeaveCell()
    txtMoney.Visible = False
End Sub

Private Sub mshDetail_Scroll()
    txtMoney.Visible = False
End Sub

Private Sub mshQuery_EnterCell()
    Dim i As Long, blnPre As Boolean
    Dim intRow As Long, intCol As Integer
    
    blnPre = mshQuery.Redraw
    intRow = mshQuery.Row: intCol = mshQuery.Col
    mshQuery.Redraw = False
    
    For i = 0 To mshQuery.Cols - 1
        mshQuery.Col = i
        mshQuery.CellBackColor = mshQuery.BackColorSel
        mshQuery.CellForeColor = mshQuery.ForeColorSel
    Next
    
    mshQuery.Row = intRow:  mshQuery.Col = intCol
    mshQuery.Redraw = blnPre
End Sub

Private Sub mshQuery_LeaveCell()
    Dim i As Long, blnPre As Boolean
    
    blnPre = mshQuery.Redraw
    mshQuery.Redraw = False
    
    For i = 0 To mshQuery.Cols - 1
        mshQuery.Col = i
        mshQuery.CellBackColor = mshQuery.BackColor
        mshQuery.CellForeColor = mshQuery.ForeColor
    Next
    
    mshQuery.Redraw = blnPre
End Sub

Private Sub mshQuery_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mnuFileZero.Visible = InStr(",2,4,7,", tabCard.SelectedItem.Index) > 0
        mnuFile_1.Visible = InStr(",2,4,7,", tabCard.SelectedItem.Index) > 0
        PopupMenu mnuFile, 2
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mnuFileZero.Visible = False
        mnuFile_1.Visible = False
        PopupMenu mnuFile, 2
    End If
End Sub

Private Sub opt出院_Click()
    
    Call zlChangeDefaultTime
    If mshDetail.TextMatrix(1, 0) <> "" Then
        If Not IsNull(mrsInfo!险类) And mbytMCMode <> 1 Then Call ShowBalance   '医保重新预结算
        Call ShowMoney(True, , mty_ModulePara.bytMzDeposit)
    End If
End Sub

Private Sub opt中途_Click()
    Call zlChangeDefaultTime
    If mshDetail.TextMatrix(1, 0) <> "" Then
        If Not IsNull(mrsInfo!险类) And mbytMCMode <> 1 Then Call ShowBalance '医保重新预结算
        Call ShowMoney(True, , mty_ModulePara.bytMzDeposit)
    End If
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "LocalParSet" Then
        frmSetExpence.mstrPrivs = mstrPrivs
        frmSetExpence.mbytInFun = 1
        frmSetExpence.Show 1, Me
    End If
End Sub

Private Sub tabCard_Click()
    If tabCard.SelectedItem.Index = 1 Then
        mshDetail.ZOrder
        txtMoney.ZOrder
        
        mshDetail.Visible = True
        mshQuery.Visible = False
        
        mshDetail.TopRow = 1
        mshDetail.Row = 1
        mshDetail.Col = GetColNum("结帐金额") ' mshDetail.Cols - 1
        If mshDetail.Visible Then mshDetail.SetFocus
    Else
        mshQuery.ZOrder
        mshQuery.Visible = True
        
        mshDetail.Visible = False
        
        '没有读取或清单类型时读取
        If (mshQuery.TextMatrix(1, 0) = "" And mshQuery.Rows = 2) _
            Or Val(mshQuery.Tag) <> tabCard.SelectedItem.Index Then
            Call LoadCardData
        End If
                
        If mshQuery.Visible And mshQuery.Enabled Then mshQuery.SetFocus
    End If
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0: txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then mshDeposit.SetFocus
End Sub

Private Sub txtInvoice_Change()
    lblFact.Tag = ""
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtInvoice.Text) = txtInvoice.MaxLength And KeyAscii <> 8 And txtInvoice.SelLength <> Len(txtInvoice) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtInvoice_GotFocus()
    SelAll txtInvoice
End Sub


Private Sub txtMoney_LostFocus()
    txtMoney.Visible = False
End Sub

Private Sub cboNO_GotFocus()
    If Not cboNO.Locked Then cboNO.SelStart = 0: cboNO.SelLength = Len(cboNO.Text)
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date, bytFlag As Byte
    Dim lng病人ID  As Long
    '转换成大写(汉字不可处理)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 15)
        
        '是否已转入后备数据表中
        If zlDatabase.NOMoved("病人结帐记录", cboNO.Text, , , Me.Caption) Then
            If Not ReturnMovedExes(cboNO.Text, 7, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
    
        '单据权限
        If Not ReadBillInfo(2, cboNO.Text, -1, strOper, vDate) Then
            cboNO.Text = "": If cboNO.Visible Then cboNO.SetFocus
            Exit Sub
        End If
        If Not BillOperCheck(7, strOper, vDate, "作废") Then
            cboNO.Text = "": If cboNO.Visible Then cboNO.SetFocus
            Exit Sub
        End If
        'lng病人ID:49084
        mintInsure = BalanceExistInsure(cboNO.Text, bytFlag, lng病人ID)
        mbytMCMode = bytFlag
        If mintInsure <> 0 Then
            '保险结算权限判断
            If InStr(mstrPrivs, ";保险结算;") = 0 Then
                MsgBox "你没有权限作废保险病人的结帐单据。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            MCPAR.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, mintInsure)
            If mbytMCMode = 1 Then
                MCPAR.门诊病人结算作废 = gclsInsure.GetCapability(support门诊结算作废, lng病人ID, mintInsure)
            Else
                MCPAR.出院病人结算作废 = gclsInsure.GetCapability(support出院病人结算作废, lng病人ID, mintInsure)
            End If
            MCPAR.结帐作废后打印回单 = gclsInsure.GetCapability(support结帐作废后打印回单, lng病人ID, mintInsure)
        Else
            If InStr(mstrPrivs, ";普通病人结算;") = 0 Then
                MsgBox "你没有权限作废普通病人的结帐单据。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If CheckExistsGathering(cboNO.Text) Then
            MsgBox "该结帐单据存在已缴款的应收款记录，请退款后再执行作废。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CheckBillBeforIN(cboNO.Text) Then
            If MsgBox("该结帐单是本次住院之前发生的，你确定要作废该单据吗?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        '读取要作废的结帐单
        If Not ReadBalance(cboNO.Text) Then
            Call NewBill
            cboNO.Text = "": If cboNO.Visible Then cboNO.SetFocus
        Else
            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        End If
    Else
           If InStr(mstrPrivs, ";普通病人结算;") = 0 Then
                MsgBox "你没有权限作废非保险病人的结帐单据。", vbInformation, gstrSysName
                Exit Sub
           End If
    End If
End Sub

Private Function CheckOutBalance(strNo As String) As Boolean
'功能：检查指定的结帐单对应的费用是否全是门诊记帐费用
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From 住院费用记录 A, 病人结帐记录 B" & vbNewLine & _
            "Where A.结帐id = B.ID And B.NO = [1] And A.门诊标志 = 2 And Rownum < 2"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    
    CheckOutBalance = rsTmp.RecordCount = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtMoney_Validate(Cancel As Boolean)
    If txtMoney.Visible Then Call txtMoney_KeyPress(13)
End Sub

Private Sub txtOwe_Change()
    If IsNumeric(txtOwe.Text) Then
        If CCur(txtOwe.Text) > 0 Then
            txtOwe.ForeColor = vbBlue
        ElseIf CCur(txtOwe.Text) < 0 Then
            txtOwe.ForeColor = vbRed
        Else
            txtOwe.ForeColor = vbBlack
        End If
    End If
End Sub

Private Sub txtPatiBegin_Change()
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        txt天数.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text)
        If Val(txt天数.Text) = 0 Then txt天数.Text = 1
    Else
        txt天数.Text = ""
    End If
End Sub

Private Sub txtPatiBegin_GotFocus()
    SelAll txtPatiBegin
End Sub

Private Sub txtPatiBegin_Validate(Cancel As Boolean)
    If txtPatiBegin.Text <> "____-__-__" And Not IsDate(txtPatiBegin.Text) Then
        Cancel = True
   End If
End Sub

Private Sub txtPatiEnd_Change()
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        txt天数.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text)
        If Val(txt天数.Text) = 0 Then txt天数.Text = 1
    Else
        txt天数.Text = ""
    End If
End Sub

Private Sub txtPatiEnd_GotFocus()
    SelAll txtPatiEnd
End Sub

Private Sub txtPatiEnd_Validate(Cancel As Boolean)
    If txtPatiEnd.Text <> "____-__-__" And Not IsDate(txtPatiEnd.Text) Then
        Cancel = True
   End If
End Sub

Private Sub txtPatient_Change()
    If Not Me.ActiveControl Is txtPatient Or txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKIND.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    SelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKIND.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub LoadPatientInfo(ByVal objCard As Card, ByVal blnCard As Boolean, _
    Optional ByVal intInsure As Integer, _
    Optional ByVal lng主页ID As Long)
    '功能:读取病人信息
    '       lng主页ID=读取指定住院次数的病人信息
    Dim strTmp As String, i As Long, strSql As String
    Dim blnICCard As Boolean, curDue As Currency, blnIDCard As Boolean
        
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset

    txtPatient.ForeColor = Me.ForeColor
    
    If objCard.名称 Like "IC卡*" And objCard.系统 = True Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.名称 Like "*身份证*" And objCard.系统 = True Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    sta.Panels(2).Text = ""
    If Not GetPatient(objCard, Trim(txtPatient.Text), blnCard, lng主页ID) Then
        If txtPatient.Text = "" Then MsgBox "没有找到该病人,请检查输入内容是否正确！", vbInformation, gstrSysName
        txtPatient.PasswordChar = "": txtPatient.Text = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        mstr本次住院日期 = ""
        Call ReInitPatiInvoice
        Exit Sub
    Else
        Unload frmSetBalance
        mstr本次住院日期 = ""
        '就诊卡密码检查
        If (objCard.名称 Like "IC卡*" Or objCard.名称 Like "*身份证*") And objCard.系统 = True And blnCard Then blnCard = False
        If Mid(gstrCardPass, 7, 1) = "1" And (blnCard Or ((blnICCard Or blnIDCard) And mstrPassWord <> "")) Then
            If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
                GoTo ExitHandle
            End If
        End If
        
        '102236,调用外挂部件接口
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            'PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, _
            '    ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, _
            '    ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
            ''功能：检查当前病人是否是指定的特殊病人
            ''返回：true时允许继续操作，False时不允许操作
            ''参数：
            ''      lngSys,lngModual=当前调用接口的主程序系统号及模块号
            ''      lngType 操作类型：1－门诊挂号，2－住院入院，3－门诊收费，4－住院结帐，5－门诊结帐。
            ''      lngPatiID-病人ID: 新建档的，为0,否则传入建档病人ID
            ''      lngPageID-主页ID: 新建档的，为0,否则传入建档主页ID(住院传入主页ID) 特殊说明：仅 lngType=4 时才传入 lngPageID，其它均传0
            ''      strPatiInforXML-病人信息:针对未建档病人传入，"姓名，性别，年龄，出生日期，医保号，身份证号"，出生日期 格式:2016-11-11 12:12:12
            ''                      固定格式：<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH>
            ''      strReserve=保留参数,用于扩展使用
            Dim blnChecked As Boolean
            blnChecked = gobjPlugIn.PatiValiedCheck(glngSys, mlngModul, IIf(mbytFunc = 0, 5, 4), Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), "")
            If Err <> 0 Then
                Call zlPlugInErrH(Err, "PatiValiedCheck"): Err.Clear
            Else
                If blnChecked = False Then GoTo ExitHandle
            End If
            On Error GoTo errHandle
        End If
        
        '问题:27690
        If Val(Nvl(mrsInfo!险类)) = 0 Then
                If InStr(1, mstrPrivs, ";普通病人结算;") = 0 Then
                    MsgBox "你没有权限对非保险病人进行结算。", vbInformation, gstrSysName
                    GoTo ExitHandle
                End If
        End If
        
        '医保相关判断
        If Not IsNull(mrsInfo!险类) Then
            If InStr(mstrPrivs, ";保险结算;") = 0 Then
                MsgBox "你没有权限对保险病人进行结算。", vbInformation, gstrSysName
                GoTo ExitHandle
            End If
            
            If mstrYBPati <> "" And intInsure <> mrsInfo!险类 Then
                MsgBox "病人登记的险类与医保身份验证的险类不符。", vbInformation, gstrSysName
                GoTo ExitHandle
            End If
            
            If mbytMCMode = 1 And Not IsNull(mrsInfo!当前科室id) Then
                MsgBox "在院病人不能进行门诊医保身份验证。", vbInformation, gstrSysName
                GoTo ExitHandle
            End If
            
            MCPAR.分币处理 = gclsInsure.GetCapability(support分币处理, mrsInfo!病人ID, mrsInfo!险类)
            If mbytMCMode = 1 Then
                MCPAR.门诊预结算 = gclsInsure.GetCapability(support门诊预算, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.门诊必须传递明细 = gclsInsure.GetCapability(support门诊必须传递明细, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.门诊结算_结帐设置 = gclsInsure.GetCapability(support门诊结帐_结帐设置后调用接口, mrsInfo!病人ID, mrsInfo!险类)
            Else
                MCPAR.未结清出院 = gclsInsure.GetCapability(support未结清出院, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.结算使用个人帐户 = gclsInsure.GetCapability(support结算使用个人帐户, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.出院结算必须出院 = gclsInsure.GetCapability(support出院结算必须出院, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.中途结算仅处理已上传部分 = gclsInsure.GetCapability(support中途结算仅处理已上传部分, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.结帐设置后调用接口 = gclsInsure.GetCapability(support结帐_结帐设置后调用接口, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.允许结多次住院费用 = gclsInsure.GetCapability(support允许一次结多次住院费用, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.门诊结算_结帐设置 = False
            End If
        ElseIf mstrYBPati <> "" Then
            MsgBox "病人身份验证成功,但病人登记的险类为空！", vbInformation, gstrSysName
                GoTo ExitHandle
        End If
        
        '问题:34763 检查病人是否存在备注信息
        
        If zlCheckPatiIsMemo(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = True Then
            Call zlCallPatiMemoWriteAndRead(Me, mlngModul, mstrPrivs, Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), mobjInPatient)
        End If
        
        If lng主页ID = 0 Then
            If mbytMCMode <> 1 Then
                If mrsInfo!主页ID <> 0 And (mbytFunc = 1 And mbytInState = 0) Then
                    '问题:30027:现在缺省的中途规则
                    '       1.出院病人,默认为出院结帐 或者:没有"中途结帐"权限的,也默认为出院结帐
                    '       2.在院病人-普通病人(根据上次出院病人的选择的为准)
                    '              默认出院结(即上次选择的中途结帐或住院结帐)参数为true,默认为出院结帐,否则默认为中途结帐
                    '       3.在院病人-医保病人(不处理)
                    '           由于医保这边不好确定,因此,暂与原来的功能一样,不根据上次出院病人的选择的为准!
                    If InStr(mstrPrivs, ";中途结帐;") = 0 Then
                        opt出院.Value = True: opt中途.Enabled = False
                    ElseIf Not IsNull(mrsInfo!当前科室id) And Nvl(mrsInfo!状态, 0) <> 3 Then  '在院病人()
                            If IsNull(mrsInfo!险类) Then
                                '医保病人需要支持中途结帐时只处理已上传部份,所以不管
                                If zlDatabase.GetPara("默认出院结帐", glngSys, mlngModul, "1") <> "0" Then
                                    opt出院.Value = True
                                Else
                                    opt中途.Value = True
                                End If
                            End If
                    Else
                        '出院病人(包含预出院的病人)
                         opt出院.Value = True
                    End If
                    opt出院.Enabled = True
                    
                    '在院病人不允许出院结帐(预出院病人可以)
                    If gbln在院不准结帐 And Not IsNull(mrsInfo!当前科室id) Then         'And Nvl(mrsInfo!状态, 0) <> 3:30572:预出院也是在院.
                        If Not opt中途.Enabled Then
                            MsgBox "在院病人不允许出院结帐,并且你没有中途结帐的权限,所以不能对该病人结帐!", vbInformation, gstrSysName
                            GoTo ExitHandle
                        End If
                        If mblnFirst And mlngPatientID <> 0 Then
                            '第一次自动读取病人结帐时,不去检查和提提
                            '38537:如果是在院病人,肯定需要设置为中途结帐
                            opt中途.Value = True: opt出院.Value = False: opt出院.Enabled = False
                        Else
                            If opt中途.Value Then
                                opt出院.Value = False: opt出院.Enabled = False
                            Else
                                If MsgBox("当前病人在院，不允许出院结帐。" & vbCrLf & "如果是出院结帐，请先将病人出院。" & _
                                    vbCrLf & "需要对该病人进行中途结帐吗?", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbYes Then
                                    opt出院.Value = False: opt出院.Enabled = False
                                    opt中途.Value = True
                                Else
                                    GoTo ExitHandle
                                End If
                            End If
                        End If
                    End If
                Else
                    '问题:47430
                    opt出院.Value = True: opt出院.Enabled = False
                    opt中途.Enabled = False
                End If
            End If
            
            
            '黑名单提醒
            strTmp = inBlackList(mrsInfo!病人ID)
            If strTmp <> "" Then
                If MsgBox("病人""" & mrsInfo!姓名 & """在特殊病人名单中。" & vbCrLf & vbCrLf & "原因：" & vbCrLf & vbCrLf & "　　" & strTmp & vbCrLf & vbCrLf & "要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    GoTo ExitHandle
                End If
            End If
                                                                                        
            'gbytAuditing:0-不检查,1-检查并提示,2-检查并禁止
            '问题:37369:中途结帐不检查
            If gbytAuditing <> 0 And (mbytFunc = 1 And mbytInState = 0) Then
                If HaveNOAuditing(mrsInfo!病人ID) Then
                    If gbytAuditing = 1 Then
                        If MsgBox("该病人还存在未审核的记帐费用，要结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            GoTo ExitHandle
                        End If
                    ElseIf gbytAuditing = 2 Then
                         If MsgBox("该病人还存在未审核的记帐费用，要结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                GoTo ExitHandle
                         End If
                          If opt中途.Enabled Then opt中途.Value = True
                    End If
                End If
            End If
            
            '自动计算病人的床位费用和护级费用
            If mrsInfo!主页ID <> 0 And mbytMCMode <> 1 Then
                strSql = "ZL1_AUTOCPTPATI(" & mrsInfo!病人ID & "," & mrsInfo!主页ID & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            Call Init预交类别
            '获取病人费用余额
            If mint预交类别 = 0 Then
                strSql = "Select Sum(预交余额) As 预交余额,Sum(费用余额) As 费用余额 From 病人余额 Where 病人ID= [1] And 性质=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsInfo!病人ID)))
            Else
                strSql = "Select 预交余额,费用余额 From 病人余额 Where 病人ID= [1] And 性质=1 And 类型= [2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsInfo!病人ID)), mint预交类别)
            End If
            mcurSpare = Get病人余额(mrsInfo!病人ID, 0, mint预交类别)
            lblSpare.Tag = Get病人余额(mrsInfo!病人ID, 1, mint预交类别)  'ShowBalance中LED显示会用到此金额
            lblSpare.Caption = "预交余额:" & Format(lblSpare.Tag, "0.00")
            '60615,刘尔旋,2013-12-20,状态栏显示预交余额、费用金额和剩余余额
            If rsTmp.RecordCount <> 0 Then
                sta.Panels(3).Text = "预交:" & Format(Nvl(rsTmp!预交余额), "0.00") & _
                                     "/费用:" & Format(Nvl(rsTmp!费用余额), "0.00") & _
                                     "/剩余:" & Format(Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额)), "0.00")
            End If
            
            If InStr(mstrPrivs, ";应收款管理;") > 0 Then
                curDue = GetPatientDue(Val(mrsInfo!病人ID))
                If curDue <> 0 Then
                    MsgBox mrsInfo!姓名 & ",应收款余额:" & Format(curDue, "0.00") & "元", vbInformation, gstrSysName
                    sta.Panels(2).Text = "病人应收款余额:" & Format(curDue, "0.00") & "元"
                End If
            End If
            
            '88786,结帐不处理历史数据
            mblnDateMoved = False
        Else
            If IsNull(mrsInfo!当前科室id) And Nvl(mrsInfo!状态, 0) <> 3 Then
                opt出院.Value = True: opt出院.Visible = True: opt出院.Enabled = True
            End If
        End If
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        
        txtPatient.IMEMode = 0
        txtPatient.Text = mrsInfo!姓名: txtSex.Text = Nvl(mrsInfo!性别): txtOld.Text = Nvl(mrsInfo!年龄)
        '显示病人险类
        '62906
        '挂号时,病人未进行医保验证时,门诊允许输入病人后,重新验证医保
        cmdYB.Enabled = IIf(mbytFunc = 0, True, False)
        If Not IsNull(mrsInfo!险类) Then
            sta.Panels(2).Text = sta.Panels(2).Text & "  险类：" & GetInsureName(mrsInfo!险类)
            If mbytMCMode = 1 Then Call InitBalanceSet(False)
            cmdOK.Enabled = False
        End If
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!病人类型))
        
        lbl状态.Caption = GetPatiState(mrsInfo!病人ID)
        lbl付款方式.Left = lbl状态.Left + lbl状态.Width + 100
        lbl付款方式.Caption = "" & mrsInfo!医疗付款方式
        pic状态.Width = lbl状态.Width + lbl付款方式.Width + 300
        pic状态.Visible = True
        
        txt费别.Text = Nvl(mrsInfo!费别)
        
        '问题65105,刘尔旋:门诊结帐时显示门诊号
        If mbytFunc = 1 Then
            If mstr录入住院号 <> "" Then
                txt标识号.Text = mstr录入住院号
                lbl标识号.Visible = True: txt标识号.Visible = True
                lbl标识号.Caption = "住院号"
            Else
                If Not IsNull(mrsInfo!住院号) Then
                    txt标识号.Text = mrsInfo!住院号
                    lbl标识号.Visible = True: txt标识号.Visible = True
                    lbl标识号.Caption = "住院号"
                End If
            End If
            If Not IsNull(mrsInfo!当前科室) Then
                txtBed.Text = "" & mrsInfo!当前床号
                txt科室.Text = mrsInfo!当前科室
                lblBed.Visible = True: txtBed.Visible = True
                lbl科室.Visible = True: txt科室.Visible = True
            ElseIf Not IsNull(mrsInfo!出院科室) Then
                txtBed.Text = Nvl(mrsInfo!出院病床)
                txt科室.Text = mrsInfo!出院科室
                lblBed.Visible = True: txtBed.Visible = True
                lbl科室.Visible = True: txt科室.Visible = True
            End If
        ElseIf mbytFunc = 0 Then
            If Not IsNull(mrsInfo!门诊号) Then
                txt标识号.Text = mrsInfo!门诊号
                lbl标识号.Visible = True: txt标识号.Visible = True
                lbl标识号.Caption = "门诊号"
            End If
        End If
        
        '显示病人要结帐内容,并初始化结算金额
        '-------------------------------------------------------------------------------------------
        If lng主页ID = 0 Then
            strTmp = ""
            If Not ShowBalance(True, strTmp) Then
                MsgBox strTmp, vbInformation, gstrSysName
                GoTo ExitHandle
            End If
                    
            Call Led欢迎信息
        End If
        
        If vsfMoney.Visible And vsfMoney.Enabled Then vsfMoney.SetFocus
    End If
    
    Call ReInitPatiInvoice
    Call Calc找补
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
ExitHandle:
    mcurSpare = 0
    Call NewBill
    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
    Exit Sub
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKIND.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, strInput As String
    If txtPatient.Locked Then Exit Sub
    '病人选择器
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        With frmPatiSelect
            .mstrPrivs = mstrPrivs
            .mbytUseType = 3
            Set .mfrmParent = Me
            .Show 1, Me
            mintPatientRange = Val(zlDatabase.GetPara("显示结清病人", glngSys, mlngModul, 0))
        End With
    Else
        If IDKIND.GetCurCard.名称 Like "姓名*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKIND.ShowPassText)
        ElseIf IDKIND.GetCurCard.名称 = "门诊号" Or IDKIND.GetCurCard.名称 = "住院号" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
            End If
        Else
            If IDKIND.GetCurCard.接口序号 <> 0 Then
                blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKIND.ShowPassText)
            End If
            txtPatient.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
    Me.Refresh
    
    '刷卡完毕或输入号码后回车
    If blnCard And Len(txtPatient.Text) = IDKIND.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        strInput = txtPatient.Text
        Call FindPati(IDKIND.GetCurCard, blnCard, strInput)
    End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call NewBill
    txtPatient.Text = strInput
    '刘兴洪:27503
    If mty_ModulePara.bln结帐后不清信息 Then
        If txtInvoice.Tag <> "" And txtInvoice.Text <> txtInvoice.Tag Then txtInvoice.Text = txtInvoice.Tag '主要是要保留信息,在确定后需要减判刑断
    End If
    If mblnFirst Then mstrTime = mstr主页Id
    If mblnOneCard And Not mobjICCard Is Nothing And objCard.名称 Like "IC卡*" And objCard.系统 Then
        Call SetOneCardBalance  '显示一卡通余额
    End If
    Call LoadPatientInfo(objCard, blnCard)
End Sub

Private Sub vsfMoney_Scroll()
    txtMoney.Visible = False
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人信息
    '入参:blnCard=是否就诊卡刷卡,lng主页ID=读取指定住院次数的病人信息
    '出参:
    '返回:是否读取成功,成功时mrsInfo中包含病人信息,失败时mrsInfo=Close,strInput返回是用来判断是否已提示过,避免再次提示没有找到病人
    '编制:刘兴洪
    '日期:2011-08-03 16:56:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strWhere As String, strField As String, bytMzMode As Byte
    Dim strSql As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, strRange As String
    Dim vRect As RECT, rsTemp As ADODB.Recordset, strTemp As String
    
    mstrPassWord = ""
    mlngCardTypeID = 0
    On Error GoTo errH
    strField = ",A.当前科室ID"
    bytMzMode = mbytMCMode
    mstr结帐主页IDs = ""
    mstr录入住院号 = ""
    
    If mlngPatientID <> 0 And mblnFirst Then
        '第一次取数时
        lng主页ID = Val(mstr主页Id)
        If Val(mstr主页Id) = 0 Then '门诊
            strWhere = strWhere & " And B.主页ID(+)=-100"
            bytMzMode = IIf(bytMzMode = 0, 0, 1): strField = " ,NULL as 当前科室ID"
            If mbytFunc = 1 Then bytMzMode = 2  '住院的:44022
        Else    '指定次数
            strWhere = strWhere & "  And B.主页ID=[3]"
            bytMzMode = 2   '住院的
        End If
    Else
        If mbytFunc = 0 Then    '门诊
            strWhere = strWhere & " And   A.主页ID=B.主页ID(+)"
            '问题:43730
            bytMzMode = IIf(bytMzMode = 0, 0, 1): strField = " ,NULL as 当前科室ID"
        Else
            '指定次数
            '76451,冉俊明,2014-8-19
            If lng主页ID <> 0 Then strField = ",Decode(A.主页ID,[3],A.当前科室ID,NULL) as 当前科室ID"
            strWhere = IIf(lng主页ID = 0, " And A.主页ID=B.主页ID(+)", " And B.主页ID=[3]")
            bytMzMode = 2
        End If
    End If
    strSql = _
        "Select A.病人ID,Nvl(B.主页ID,0) as 主页ID,A.门诊号,nvl(B.住院号,A.住院号) as 住院号,A.当前床号,B.出院病床," & _
        "       nvl(B.姓名,A.姓名) as 姓名, nvl(B.性别,Nvl(A.性别,'未知')) as  性别,Nvl(B.年龄,A.年龄) as 年龄,A.IC卡号,A.就诊卡号,A.卡验证码," & _
        "       Nvl(B.费别,A.费别) as 费别,C.名称 as 当前科室" & strField & ",D.名称 as 出院科室,B.出院科室ID," & _
                IIf(bytMzMode = 0, "NULL", IIf(bytMzMode = 1, "A.险类", "B.险类")) & " as 险类,E.卡号,E.医保号,E.密码," & _
        " A.登记时间,Nvl(B.状态,0) as 状态,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Nvl(B.审核标志,0) as 审核标志,B.入院日期,B.出院日期,B.病人性质,B.病人类型" & _
        " From 病人信息 A,病案主页 B,部门表 C,部门表 D,医保病人档案 E,医保病人关联表 F" & _
        " Where A.停用时间 is NULL And A.病人ID=B.病人ID(+)   " & strWhere & _
        " And A.病人ID=F.病人ID(+) And F.标志(+)=1 And F.医保号=E.医保号(+) And F.险类=E.险类(+) And F.中心 = E.中心(+)" & _
        " And A.当前科室ID=C.ID(+) And B.出院科室ID=D.ID(+)"
        
    If blnCard = True And objCard.名称 Like "姓名*" Then    '刷卡
        If IDKIND.Cards.按缺省卡查找 And Not IDKIND.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKIND.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        mlngCardTypeID = lng卡类别ID
        strSql = strSql & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSql = strSql & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSql = strSql & " And A.门诊号=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
    
        strSql = Replace(strSql, "And A.主页ID=B.主页ID(+)", "") & " And  B.住院号=[2] " ' & " And A.病人ID=(Select nvl(Max(病人ID),0) as 病人ID From 病案主页   Where  住院号=[2])"
        strInput = Mid(strInput, 2)
        strTemp = "Select Distinct(主页ID) As 主页ID From 病案主页 Where 住院号 = [1] And 病人性质 = 0 Order By 主页ID DESC"
        Set rsTemp = zlDatabase.OpenSQLRecord(strTemp, Me.Caption, strInput)
        Do While Not rsTemp.EOF
            If Val(Nvl(rsTemp!主页ID)) <> 0 Then
                mstr结帐主页IDs = mstr结帐主页IDs & "," & rsTemp!主页ID
            End If
            rsTemp.MoveNext
        Loop
        If mstr结帐主页IDs <> "" Then mstr结帐主页IDs = Mid(mstr结帐主页IDs, 2): mstr录入住院号 = strInput
    Else '当作姓名
        mlngCardTypeID = objCard.接口序号
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If mrsInfo.State = 1 Then
                    If mrsInfo!姓名 = Trim(txtPatient.Text) Then
                        GetPatient = True
                        Exit Function
                    End If
                End If
                
                If mintPatientRange > 0 Then
                    Select Case mintPatientRange
                        Case 1  '任何费用未结清病人
                            strRange = ""
                        Case 2  '体检未结清的病人
                            strRange = " And C.来源途径 = 4"
                        Case 3  '住院未结清的病人
                            strRange = " And C.来源途径 = 2"
                        Case 4  '门诊未结清的病人
                            strRange = " And C.来源途径 = 1"
                    End Select
                    strPati = " And Exists(Select 1 From 病人未结费用 C Where C.病人id=A.病人ID And Nvl(C.主页ID,0)=A.主页ID" & strRange & ")"
                End If
                
                 '通过姓名查找
                strPati = "" & _
                " Select A.病人ID as ID,A.病人ID,A.住院号, A.门诊号, nvl(B.性别,Nvl(A.性别,'未知')) as  性别, A.年龄, A.住院次数, A.家庭地址, A.工作单位," & vbNewLine & _
                "   To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,  To_Char(B.入院日期,'YYYY-MM-DD') as 入院日期, To_Char(B.出院日期,'YYYY-MM-DD') as 出院日期" & vbNewLine & _
                " From 病人信息 A, 病案主页 B" & vbNewLine & _
                " Where A.病人id = B.病人id(+) And A.主页ID = B.主页id(+) And A.停用时间 Is Null And A.姓名 = [1] " & vbNewLine & strPati & vbNewLine & _
                " Order By Decode(住院号, Null, 1, 0), 入院日期 Desc"
                        
                vRect = GetControlRect(txtPatient.hWnd)
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If Not mrsInfo Is Nothing Then
                    strInput = Val(mrsInfo!病人ID)
                    strSql = strSql & " And A.病人ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
                End If
            Case "医保号"
                strInput = UCase(strInput)
                strSql = strSql & " And A.医保号=[2]"
            Case "身份证号", "二代身份证", "身份证"
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strInput = "-" & lng病人ID
                blnHavePassWord = True
                strSql = strSql & " And A.病人ID=[1] "
            Case "IC卡号"
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strInput = "-" & lng病人ID
                blnHavePassWord = True
                strSql = strSql & " And A.病人ID=[1] "
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.病人ID= (Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号 = [2])"
                strTemp = "Select Distinct(主页ID) As 主页ID From 病案主页 Where 住院号 = [1] And 病人性质 = 0 Order By 主页ID DESC"
                Set rsTemp = zlDatabase.OpenSQLRecord(strTemp, Me.Caption, strInput)
                Do While Not rsTemp.EOF
                    If Val(Nvl(rsTemp!主页ID)) <> 0 Then
                        mstr结帐主页IDs = mstr结帐主页IDs & "," & rsTemp!主页ID
                    End If
                    rsTemp.MoveNext
                Loop
                If mstr结帐主页IDs <> "" Then mstr结帐主页IDs = Mid(mstr结帐主页IDs, 2): mstr录入住院号 = strInput
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strSql = strSql & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    strSql = strSql & " Order by 主页ID Desc"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Mid(strInput, 2)), strInput, lng主页ID)
    If mrsInfo.EOF Then Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
    
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then
        mstrPassWord = Nvl(mrsInfo!卡验证码)
    End If
    
    '检查死亡情况:如果死亡则提示
    '34681:35686
    If zlCheckPatiIsDeath(Val(Nvl(mrsInfo!病人ID))) = True Then
        If MsgBox("注意:" & vbCrLf & "    该病人已经死亡,是否继续结帐?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
        End If
    End If
    
    '需要再次检查,以防结帐期间已审核的病人被取消审核
    '36209
    If (InStr(mstrPrivs, ";未审核病人中途结帐;") = 0 And opt中途.Value _
        Or InStr(mstrPrivs, ";未审核病人出院结帐;") = 0 And opt出院.Value) _
        And (mbytFunc = 1 And mbytInState = 0) Then
        
        If Not Chk病人审核(mrsInfo!病人ID, Val(Nvl(mrsInfo!主页ID))) Then
            If MsgBox("待结帐费用中包含病人第" & Val(Nvl(mrsInfo!主页ID)) & "次住院未审核的费用记录。" & vbCrLf & _
                " 是否继续结帐?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
            End If
        End If
    End If
    GetPatient = True
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Function ShowBillFormat()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前登录的收费操作员显示它所使用收费票据格式
    '编制:刘兴洪
    '日期:2011-01-02 09:47:25
    '问题:35142
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intFormat As Integer, strRptName As String
    Dim bln医保病人 As Boolean, bytInvoiceKind As Byte
    
    lblFormat.Caption = "": bln医保病人 = False
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then bln医保病人 = Not IsNull(mrsInfo!险类)
    End If
    
    If mbytFunc = 0 Then
        bytInvoiceKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, "0"))
    Else
        bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, "0"))
    End If
    
    'bytInvoiceKind:结帐票据类型,0-住院票据;1-门诊票据
    strRptName = IIf(bytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    intFormat = mintInvoiceFormat
    
    If intFormat = 0 Then   '以缺省票据格式显示
        intFormat = Val(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\zl9Report\LocalSet\" & strRptName, "Format", 1))
    End If
    
    strSql = "Select B.说明 From zlReports A,zlRptFmts B" & _
        " Where A.ID=B.报表ID And A.编号=[1] And B.序号=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName, intFormat)
    If Not rsTmp.EOF Then
        lblFormat.Caption = "票据:" & Nvl(rsTmp!说明)
        lblFormat.Visible = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowRedFormat()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前登录的收费操作员显示它所使用收费票据格式
    '编制:刘兴洪
    '日期:2011-01-02 09:47:25
    '问题:35142
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intFormat As Integer, strRptName As String
    Dim bln医保病人 As Boolean
    
    lblFormat.Caption = "": bln医保病人 = False
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then bln医保病人 = Not IsNull(mrsInfo!险类)
    End If
    
    'gbytInvoiceKind:结帐票据类型,0-住院票据;1-门诊票据
    strRptName = IIf(gbytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6")
    intFormat = mintInvoiceFormat
    If intFormat = 0 Then   '以缺省票据格式显示
        intFormat = Val(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\zl9Report\LocalSet\" & strRptName, "Format", 1))
    End If
    
    strSql = "Select B.说明 From zlReports A,zlRptFmts B" & _
        " Where A.ID=B.报表ID And A.编号=[1] And B.序号=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName, intFormat)
    If Not rsTmp.EOF Then
        lblFormat.Caption = "票据:" & Nvl(rsTmp!说明)
        lblFormat.Visible = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowBalance(Optional ByVal blnFirst As Boolean, Optional ByRef strMessage As String) As Boolean
'功能：根据设置,显示病人要结帐内容,并初始化结算金额
'参数：blnFirst-病人身份确定时调用，strMessage-返回提示信息
'说明：该功能可能是上一个病人结帐完成后进行,也可能是当一个病人在结帐时另一病人中途进行
    On Error GoTo errHandle
    Dim i As Long, j As Long, cur统筹支付 As Currency, cur个人帐户 As Currency, curTmp As Currency, lngMaxLength As Long, lngP As Long
    Dim rsDetail As New ADODB.Recordset, intCount As Integer
    Dim strMoney As String, strInfo As String, strTime As String
    Dim blnUpload As Boolean, blnZero As Boolean, blnAll As Boolean
    Dim dBegin As Date, dEnd As Date, DatTmp As Date
    Dim dblMoney As Double, str住院次数 As String, strStyle As String
    
    Call ClearDetail(False)
    Call AdjustBalance
    Call AdjustDeposit
    
    If mrsInfo.State <> 1 Then Exit Function
    Screen.MousePointer = 11
    Me.Refresh
    
    blnZero = gblnZero
    
    If Not IsNull(mrsInfo!险类) And mbytMCMode <> 1 Then
        If opt中途.Value And MCPAR.中途结算仅处理已上传部分 Then blnUpload = True
    End If
    
    Set mrsBalance = GetBalance(mbytFunc, mstrPrivs, mrsInfo!病人ID, IIf(mbytFunc = 0, "0", mstrTime), mstrUnit, mstrClass, mDateBegin, mDateEnd, mbytBaby, mstrItem, blnUpload, blnZero, mblnDateMoved, mbytMCMode = 1, mbytKind, mtySquareCard.blnExistsObjects, mstrChargeType, mstrDiagnose)
    If mrsBalance Is Nothing Then Screen.MousePointer = 0: Exit Function
    
    If blnFirst And mrsBalance.RecordCount = 0 And mbytFunc = 0 Then
        mbytKind = 1 '缺省只取普通费用，如果没有再检查只有体检费用这种情况
        Set mrsBalance = GetBalance(mbytFunc, mstrPrivs, mrsInfo!病人ID, IIf(mbytFunc = 0, "0", mstrTime), mstrUnit, mstrClass, mDateBegin, mDateEnd, mbytBaby, mstrItem, blnUpload, blnZero, mblnDateMoved, mbytMCMode = 1, mbytKind, mtySquareCard.blnExistsObjects, mstrChargeType, mstrDiagnose)
        If mrsBalance Is Nothing Then
            Screen.MousePointer = 0: Exit Function
        ElseIf mrsBalance.RecordCount > 0 Then
            If MsgBox("该病人普通费用已结清,要对体检费用进行结帐吗?", vbInformation, Me.Caption) = vbNo Then
                Set mrsBalance = Nothing
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    If mrsBalance.RecordCount = 0 Then
        If blnFirst Then strMessage = "该病人没有需要结帐的费用！"
        Screen.MousePointer = 0: Exit Function
    End If
    
    If blnFirst Then
        Call GetStateIF
        If InStr(mstrPrivs, ";未审核病人中途结帐;") = 0 And InStr(mstrPrivs, ";未审核病人出院结帐;") = 0 _
            And (mbytFunc = 1 And mbytInState = 0) Then
            
            If CStr(mrsInfo!主页ID) = mstrAllTime Then
                If mrsInfo!审核标志 = 0 And mrsInfo!主页ID <> 0 Then
                    strMessage = "当前病人未审核，你不能对未审核的病人进行结帐。"
                    Screen.MousePointer = 0: Exit Function
                End If
            Else
                blnAll = True
                For i = 0 To UBound(Split(mstrAllTime, ","))
                    strTime = Split(mstrAllTime, ",")(i)
                    If Val(strTime) <> 0 Then
                        If Not Chk病人审核(mrsInfo!病人ID, Val(strTime)) Then
                            mstrUnAuditTime = IIf(mstrUnAuditTime = "", strTime, mstrUnAuditTime & "," & strTime)
                        Else
                            blnAll = False
                        End If
                    Else
                        blnAll = False
                    End If
                Next
                If blnAll Then
                    strMessage = "该病人所有住院费用都没有审核，不能进行结帐！"
                    Screen.MousePointer = 0: Exit Function
                End If
            End If
        End If
        If cmdPar.Enabled Then
            If (gbln多次住院弹出结帐设置 And InStr(1, mstrAllTime, ",") > 0 Or Not IsNull(mrsInfo!险类) And MCPAR.结帐设置后调用接口) Or MCPAR.门诊结算_结帐设置 Then
                '---------------------------------------------------------------------------------------
                '34260:输血费检查
                If gbyt结帐时输血费检查 = 1 And InStr(1, "," & mstrALLChargeType & ",", ",'K',") > 0 Then     '0:不检查;1-检查并提示
                    Call MsgBox("注意:" & vbCrLf & "    该病人未结费用中包含了输血费,请注意对输血费进行结帐!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
                End If
                Screen.MousePointer = 0
                mblnNOCancel = True
                Call cmdPar_Click
                mblnNOCancel = False
                ShowBalance = True  '结帐设置条件如果没有待结费用，仍返回成功，允许再次选择。
                Exit Function
            End If
        End If
        '---------------------------------------------------------------------------------------
        '34260:输血费检查
        If gbyt结帐时输血费检查 = 1 Then '0:不检查;1-检查并提示
            If InStr(1, "," & mstrALLChargeType & ",", ",'K',") > 0 Then  '34260
                If MsgBox("注意:" & vbCrLf & "    该病人未结费用中包含了输血费,本次是否只结输血费?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mstrChargeType = "'K'"
                     If ShowBalance(False) Then
                        ShowBalance = True
                     End If
                    Exit Function
                End If
            End If
        End If
        '---------------------------------------------------------------------------------------
    End If
    '78317:医保病人默认只读取最后一次住院的数据
    If Val(Nvl(mrsInfo!险类)) <> 0 And mstrTime = "" Then
        '117298,门诊验证农合病人信息确定后弹出错误提示
        If MCPAR.允许结多次住院费用 = False And mbytFunc <> 0 Then '114915
            mstrTime = Split(GetValiedTimes & ",", ",")(0)
            If Val(mstrTime) <> Val(Nvl(mrsInfo!主页ID)) Then
                '需要重新加载病人，以便进行相关医保判断)
                IDKIND.IDKIND = IDKIND.GetKindIndex("姓名")
                txtPatient.Text = "-" & mrsInfo!病人ID
                Call LoadPatientInfo(IDKIND.GetCurCard, False, 0, Val(mstrTime))
                
            End If
        End If
        Set mrsBalance = GetBalance(mbytFunc, mstrPrivs, mrsInfo!病人ID, IIf(mbytFunc = 0, "0", mstrTime), mstrUnit, mstrClass, mDateBegin, mDateEnd, mbytBaby, mstrItem, blnUpload, blnZero, mblnDateMoved, mbytMCMode = 1, mbytKind, mtySquareCard.blnExistsObjects, mstrChargeType)
        If blnFirst And mstr结帐主页IDs <> "" Then Call ReCaculateTime
    Else
        If blnFirst And mstr结帐主页IDs <> "" Then
            mstrTime = GetValiedTimes
            Set mrsBalance = GetBalance(mbytFunc, mstrPrivs, mrsInfo!病人ID, IIf(mbytFunc = 0, "0", mstrTime), mstrUnit, mstrClass, mDateBegin, mDateEnd, mbytBaby, mstrItem, blnUpload, blnZero, mblnDateMoved, mbytMCMode = 1, mbytKind, mtySquareCard.blnExistsObjects, mstrChargeType)
            If mrsBalance.RecordCount = 0 Then
                mstrTime = ""
                Set mrsBalance = GetBalance(mbytFunc, mstrPrivs, mrsInfo!病人ID, IIf(mbytFunc = 0, "0", mstrTime), mstrUnit, mstrClass, mDateBegin, mDateEnd, mbytBaby, mstrItem, blnUpload, blnZero, mblnDateMoved, mbytMCMode = 1, mbytKind, mtySquareCard.blnExistsObjects, mstrChargeType)
            End If
            Call ReCaculateTime
        End If
    End If
    
    '绑定显示费用明细
    '标志,住院,科室,时间,[单据号],项目,费目,婴儿费,[ID],[序号],[记录性质],[记录状态],[执行状态],[主页ID],[开单部门ID],[登记时间],未结金额,结帐金额
    
    With mshDetail
        .Redraw = False
        Set .DataSource = mrsBalance
        .Cols = 18 '  .Cols - 1 '不显示费用类型
        .ToolTipText = "共" & mrsBalance.RecordCount & "条明细记录!"
        
        '调整明细格式
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4
            Select Case .TextMatrix(0, i)
                Case "住院", "婴儿费", "单据号"
                    .ColAlignment(i) = 4
                Case "科室", "项目", "费目", "时间"
                    .ColAlignment(i) = 1
                Case "未结金额", "结帐金额"
                    .ColAlignment(i) = 7
            End Select
            Select Case .TextMatrix(0, i)
                Case "ID", "标志", "序号", "记录性质", "主页ID", "开单部门ID", "记录状态", "执行状态", "科室", "住院", "登记时间", _
                     "费别", "执行部门ID", "收费类别", "开单人", "数量", "价格", "统筹金额", "保险大类ID", "收费细目ID", "计算单位"
                    .ColWidth(i) = 0
                Case "婴儿费"
                    .ColWidth(i) = 520
                    .TextMatrix(0, i) = "婴儿"
                Case "费目"
                    .ColWidth(i) = 800
                Case "单据号"
                    .ColWidth(i) = 950
                Case "未结金额", "结帐金额"
                    .ColWidth(i) = 930
                Case "时间"
                    .ColWidth(i) = 1130
                Case "项目"
                    .ColWidth(i) = 1500
            End Select
            .ColData(i) = .ColWidth(i)
        Next
        
        lngMaxLength = Len(Mid(gstrDec, 3))
        If mrsBalance.RecordCount > 0 Then
            For i = 1 To mrsBalance.RecordCount
                lngP = InStr(1, CStr(mrsBalance!结帐金额), ".")
                If lngP > 0 Then
                    lngP = Len(Mid(CStr(mrsBalance!结帐金额), lngP + 1))
                    If lngP > lngMaxLength Then lngMaxLength = lngP
                End If
                mrsBalance.MoveNext
            Next
            mrsBalance.MoveFirst
        End If
        mstrDec = "0." & String(lngMaxLength, "0")
        
        For i = 1 To .Rows - 1
            .Row = i
            .Col = .Cols - 1
            If mbytFunc = 0 Then
                .CellBackColor = 12900351
            Else
                .CellBackColor = txtMoney.BackColor
            End If
            .Col = .Cols - 2
            .CellBackColor = 12900351
            .TextMatrix(i, COL_未结金额) = LTrim(Format(.TextMatrix(i, COL_未结金额), mstrDec))
            .TextMatrix(i, COL_结帐金额) = LTrim(Format(.TextMatrix(i, COL_结帐金额), mstrDec))
        Next
        .Redraw = True
    End With
    '医保预结算之前先显示结帐金额合计
    txtTotal.Text = Format(GetBalanceSum, mstrDec)
    txtTotal.Tag = txtTotal.Text
    dblMoney = Val(txtTotal.Text)
    '显示预交明细
    'mbln门诊转住院:36984
    str住院次数 = ""
    If mbytFunc <> 0 Then
        str住院次数 = IIf(gbln仅用指定预交款 And mbln门诊转住院 = False, IIf(mstrTime = "", mstrAllTime, mstrTime), "")
    End If
    
    Set mrsDeposit = GetDeposit(mrsInfo!病人ID, mblnDateMoved, str住院次数, mbln门诊转住院, mstrPepositDate, mint预交类别)
    mstr住院次数 = str住院次数
    intCount = 0
    mstrStyle = ""
    
    If Not mrsDeposit.EOF Then
        With mshDeposit
            .Redraw = False
            .Rows = mrsDeposit.RecordCount + 1
            For i = 1 To mrsDeposit.RecordCount
                .Row = i
                .Col = .ColIndex("冲预交"): .CellBackColor = txtMoney.BackColor
                .Col = .ColIndex("余额"): .CellBackColor = 12900351
                
                .RowData(i) = Val(Nvl(mrsDeposit!记录状态))
                .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(mrsDeposit!ID))
                .Cell(flexcpData, i, .ColIndex("ID")) = Nvl(mrsDeposit!卡类别ID) & "||" & Nvl(mrsDeposit!代扣) & "||" & Nvl(mrsDeposit!退现) & "||" & Nvl(mrsDeposit!缺省退现)
                
                If Val(Nvl(mrsDeposit!卡类别ID)) <> 0 Then
                    If Val(Nvl(mrsDeposit!缺省退现)) = 0 Then
                        intCount = intCount + 1
                        If InStr(mstrStyle, mrsDeposit!三方卡名称) = 0 Then
                            mstrStyle = mstrStyle & "," & mrsDeposit!三方卡名称
                        End If
                    End If
                End If
                
                .TextMatrix(i, .ColIndex("单据号")) = mrsDeposit!NO
                .TextMatrix(i, .ColIndex("票据号")) = "" & mrsDeposit!票据号
                .TextMatrix(i, .ColIndex("日期")) = Format(mrsDeposit!日期, "yyyy-MM-dd")
                .TextMatrix(i, .ColIndex("结算方式")) = IIf(IsNull(mrsDeposit!结算方式), "", mrsDeposit!结算方式)
                .TextMatrix(i, .ColIndex("余额")) = Format(mrsDeposit!金额, "0.00")
                .TextMatrix(i, .ColIndex("预交ID")) = Val(Nvl(mrsDeposit!预交ID))
                
                If Val(Nvl(mrsDeposit!金额)) <= dblMoney Then
                    .TextMatrix(i, .ColIndex("冲预交")) = Format(mrsDeposit!金额, "0.00")
                    dblMoney = dblMoney - FormatEx(Val(Nvl(mrsDeposit!金额)), 2)
                ElseIf dblMoney <> 0 Then
                    .TextMatrix(i, .ColIndex("冲预交")) = Format(dblMoney, "0.00")
                    dblMoney = 0
                End If
                
                mrsDeposit.MoveNext
            Next
            .Row = 1: .Col = .Cols - 1
            .Redraw = True
        End With
        
        lblTicketCount.Caption = "预交款收据:" & mrsDeposit.RecordCount & "张"
        mblnThreeDepositAfter = False
        If intCount > 1 And InStr(1, mstrPrivs, ";仅用预交结帐;") = 0 Then
            mblnThreeDepositAfter = True
        End If
        Call Form_Resize
    End If
    If mstrStyle <> "" Then mstrStyle = Mid(mstrStyle, 2)
    
    '刘兴洪:30043
    If IIf(mstrTime = "", mstrAllTime, mstrTime) <> "" Then
        Call zlSetDefaultTime(Val(Nvl(mrsInfo!病人ID)))
    End If
        
    Call GetPatiDate(dBegin, dEnd)
    
    txtPatiBegin.Text = Format(dBegin, txtPatiBegin.Format)
    txtPatiEnd.Text = Format(dEnd, txtPatiEnd.Format)
    txtPatiEnd.Tag = Format(dEnd, txtPatiEnd.Format)
    Call zlChangeDefaultTime
    '医保预结算
    If Not IsNull(mrsInfo!险类) And (Not MCPAR.结帐设置后调用接口 Or MCPAR.结帐设置后调用接口 And mblnSetPar) Then
        '获取费用明细
        Set rsDetail = GetVBalance(mbytFunc, mstrPrivs, mrsInfo!险类, mrsInfo!病人ID, IIf(mbytFunc = 0, "0", mstrTime), mDateBegin, mDateEnd, blnUpload, mblnDateMoved, mbytBaby, mbytMCMode = 1, mbytKind, mstrItem, mstrUnit, mstrClass, mstrChargeType, mstrDiagnose)
        
        '医保接口:返回各种报销金额
        If mbytMCMode = 1 Then
            If MCPAR.门诊预结算 Then
                If rsDetail.RecordCount = 0 Then
                    MsgBox "读取医保预结算数据失败!", vbInformation, gstrSysName
                    Screen.MousePointer = 0: Exit Function
                End If
            
                mstrBalance = ""
                If Not gclsInsure.ClinicPreSwap(rsDetail, mstrBalance, mrsInfo!险类, "1|1") Then
                    MsgBox "门诊医保预结算失败!", vbInformation, gstrSysName
                    Screen.MousePointer = 0: Exit Function
                End If
            End If
        Else
            mstrBalance = gclsInsure.WipeoffMoney(rsDetail, mrsInfo!病人ID, "" & mrsInfo!医保号, "1", mrsInfo!险类, "|" & IIf(opt中途.Value, 0, 1))
        End If
        
        '显示各类统筹报销总额
        cur统筹支付 = 0: cur个人帐户 = 0
        For i = 0 To UBound(Split(mstrBalance, "|"))
            strMoney = Split(mstrBalance, "|")(i)
            j = GetBalanceNature(Split(strMoney, ";")(0))
            If j = 3 Then
                cur个人帐户 = cur个人帐户 + Val(Split(strMoney, ";")(1))
            ElseIf j = 4 Then
                cur统筹支付 = cur统筹支付 + Val(Split(strMoney, ";")(1))
            End If
        Next
        lbl医保基金.Caption = "统筹支付:" & Format(cur统筹支付, "0.00")
        lbl医保基金.Visible = True
        
        '显示个帐余额
        mcur个帐余额 = gclsInsure.SelfBalance(mrsInfo!病人ID, "" & mrsInfo!医保号, IIf(mbytMCMode = 1, 10, 40), mcur个帐透支, mrsInfo!险类)
        lbl个人帐户.Caption = "帐户余额:" & Format(mcur个帐余额, "0.00")
        lbl个人帐户.Visible = True
        
        Call Form_Resize
        txtTotal.Enabled = False
        cmdOK.Enabled = mstrBalance <> "" Or (mbytMCMode = 1 And Not MCPAR.门诊预结算)
        
        If gblnLED Then
            zl9LedVoice.DisplayBank "医保结算:", "帐户余额" & Format(mcur个帐余额, "0.00"), "帐户支付" & Format(cur个人帐户, "0.00"), "统筹支付" & Format(cur统筹支付, "0.00")
            DatTmp = Time
            Do While Time < DateAdd("s", 4, DatTmp)
            Loop
        End If
    Else
        Call HideMoneyInfo
        txtTotal.Enabled = True
        cmdOK.Enabled = True
    End If
    
    strInfo = ShowMoney(True, , mty_ModulePara.bytMzDeposit)
    Call SortMoney
    
    mcurTotal = Val(txtTotal.Text) '本次设置的最大金额
    txtDate.Text = Format(zlDatabase.Currentdate, txtDate.Format)
    
    If tabCard.SelectedItem.Index <> 1 Then Call LoadCardData
    Screen.MousePointer = 0
        
    '提示未设置的结算方式
    If strInfo <> "" Then
        Me.Refresh
        MsgBox strInfo, vbInformation, gstrSysName
    End If
    
    ShowBalance = True
        Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetBalanceNature(ByVal strName As String) As Integer
'功能:根据指定的结算方式名称,返回结算性质,没有找到时,返回0
    Dim i As Long
    With vsfMoney
        For i = 1 To .Rows - 1
            If vsfMoney.TextMatrix(i, .ColIndex("结算方式")) = strName Then
                GetBalanceNature = Val(vsfMoney.TextMatrix(i, .ColIndex("性质")))
                Exit For
            End If
        Next
    End With
End Function

Private Sub ReCaculateTime()
    Dim dateThis As Date, DatTmp As Date
    Dim strArray() As String
    Dim lngMin As Long, lngMax As Long
    Dim i As Integer, strBookInDate As String
    
    mrsBalance.MoveFirst
    For i = 1 To mrsBalance.RecordCount
        '比较取最大最小值
        If gint费用时间 = 0 Then
            dateThis = mrsBalance!登记时间
        Else
            dateThis = mrsBalance!时间
        End If
        If i = 1 Then
            mMinDate = dateThis
            mMaxDate = dateThis
        Else
            If dateThis < mMinDate Then mMinDate = dateThis
            If dateThis > mMaxDate Then mMaxDate = dateThis
        End If
        
        mrsBalance.MoveNext
    Next
    
    If mstr结帐主页IDs <> "" Then
        strArray = Split(GetValiedTimes, ",")
        For i = 0 To UBound(strArray)
            If Val(strArray(i)) < lngMin Or lngMin = 0 Then lngMin = Val(strArray(i))
            If Val(strArray(i)) > lngMax Or lngMax = 0 Then lngMax = Val(strArray(i))
        Next i
    End If
    
    If lngMin <> 0 Then
        DatTmp = GetInOutDate(CLng(mrsInfo!病人ID), lngMin, 0)
        If DatTmp <> CDate("0:00:00") Then
            '获取登记时间,登记时间比入院时间要小,以登记时间为准:107022
            strBookInDate = GetBookInDate(CLng(mrsInfo!病人ID), lngMin, 1)
            If strBookInDate <> "" Then
                 If Format(DatTmp, "yyyy-mm-dd HH:MM:SS") > strBookInDate Then
                    DatTmp = CDate(strBookInDate)
                 End If
            End If
            mMinDate = DatTmp
        Else
            mMinDate = zlDatabase.Currentdate
        End If
    End If
    
    If lngMax <> 0 Then
        DatTmp = GetInOutDate(CLng(mrsInfo!病人ID), lngMax, 1)
        strBookInDate = GetBookInDate(CLng(mrsInfo!病人ID), lngMax, 0)
        If DatTmp <> CDate("0:00:00") Then
            '获取登记时间,登记时间比出院时间要大,以登记时间为准:63594
            If strBookInDate <> "" Then
                 If Format(DatTmp, "yyyy-mm-dd,HH:MM:SS") < strBookInDate Then
                    DatTmp = CDate(strBookInDate)
                 End If
            End If
            mMaxDate = DatTmp
        Else
            If strBookInDate <> "" Then
                 mMaxDate = CDate(strBookInDate)
            End If
            If mMinDate > mMaxDate Then
                mMaxDate = zlDatabase.Currentdate
            End If
        End If
    End If
    
    txtEnd.Text = Format(mMaxDate, txtEnd.Format)
    txtBegin.Text = Format(mMinDate, txtBegin.Format)
    mrsBalance.MoveFirst
End Sub

Private Function GetInOutDate(lngPati As Long, lngTimes As Long, bytMode As Byte) As Date
'功能:获取病人某次住院的入院或出院时间
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 入院日期,出院日期 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPati, lngTimes)
    If rsTmp.RecordCount > 0 Then
        If bytMode = 0 Then
            GetInOutDate = rsTmp!入院日期
        Else
            GetInOutDate = IIf(IsNull(rsTmp!出院日期), CDate("0:00:00"), rsTmp!出院日期)
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetBookInDate(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional bytMode As Byte) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的登记时间
    '返回:登记时间,格式:yyyy-mm-dd HH:MM:SS
    '编制:刘兴洪
    '日期:2013-10-22 17:16:47
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
        
    On Error GoTo errHandle
    
    strSql = " Select " & IIf(bytMode = 0, "Max", "Min") & IIf(gint费用时间 = 0, "(登记时间)", "(发生时间)")
    strSql = strSql & " As 登记时间 From 住院费用记录 Where Mod(记录性质,10) In (2,3) And 病人ID=[1] And 主页ID=[2]"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng主页ID)
    GetBookInDate = Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub GetStateIF()
'功能：获取病人的住院次数，费用科室,收入项目,费用类型,最小和最大费用时间
    Dim i As Long, dateThis As Date
    
    Call InitBalanceCondition
    
    mrsBalance.MoveFirst
    For i = 1 To mrsBalance.RecordCount
                
        '如果为空,则表示门诊记帐
        If InStr("," & mstrAllTime & ",", "," & Nvl(mrsBalance!主页ID, 0) & ",") = 0 Then
            mstrAllTime = mstrAllTime & "," & Nvl(mrsBalance!主页ID, 0)
        End If
        
        If Trim(Nvl(mrsBalance!诊断, "")) <> "" Then
            If InStr("," & mstrAllDiagnose & ",", "," & Nvl(mrsBalance!诊断) & ",") = 0 Then
                mstrAllDiagnose = mstrAllDiagnose & "," & Nvl(mrsBalance!诊断)
            End If
        End If
        
        If Trim(Nvl(mrsBalance!开单部门ID, "")) <> "" Then
            If Not IsNull(mrsBalance!科室) Then
                If InStr("," & mstrAllUnit & ",", "," & mrsBalance!开单部门ID & ":" & mrsBalance!科室 & ",") = 0 Then
                    mstrAllUnit = mstrAllUnit & "," & mrsBalance!开单部门ID & ":" & mrsBalance!科室
                End If
            End If
        End If
        
        If Trim(Nvl(mrsBalance!费目, "")) <> "" Then
            If InStr("," & mstrALLItem & ",", ",'" & mrsBalance!费目 & "',") = 0 Then
                mstrALLItem = mstrALLItem & ",'" & mrsBalance!费目 & "'"
            End If
        End If
        If Trim(Nvl(mrsBalance!收费类别)) <> "" Then '34260
            If InStr("," & mstrALLChargeType & ",", ",'" & mrsBalance!收费类别 & "',") = 0 Then
                mstrALLChargeType = mstrALLChargeType & ",'" & mrsBalance!收费类别 & "'"
            End If
        End If
        '如果为空,指没有设置费用类型
        If InStr("," & mstrAllClass & ",", ",'" & Nvl(mrsBalance!类型, "无") & "',") = 0 Then
            mstrAllClass = mstrAllClass & ",'" & Nvl(mrsBalance!类型, "无") & "'"
        End If
        
        '比较取最大最小值
        If gint费用时间 = 0 Then
            dateThis = mrsBalance!登记时间
        Else
            dateThis = mrsBalance!时间
        End If
        If i = 1 Then
            mMinDate = dateThis
            mMaxDate = dateThis
        Else
            If dateThis < mMinDate Then mMinDate = dateThis
            If dateThis > mMaxDate Then mMaxDate = dateThis
        End If
        
        mrsBalance.MoveNext
    Next
    mstrAllTime = Mid(mstrAllTime, 2)
    mstrAllUnit = Mid(mstrAllUnit, 2)
    mstrALLItem = Mid(mstrALLItem, 2)
    If mstrALLChargeType <> "" Then mstrALLChargeType = Mid(mstrALLChargeType, 2) '34260
    mstrAllClass = Mid(mstrAllClass, 2)
    If mstrAllDiagnose <> "" Then mstrAllDiagnose = Mid(mstrAllDiagnose, 2)
    
    '显示结帐时间
    txtEnd.Text = Format(mMaxDate, txtEnd.Format)
    txtBegin.Text = Format(mMinDate, txtBegin.Format)
    mrsBalance.MoveFirst
End Sub
Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKIND.SetAutoReadCard (False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mrsInfo.State = 1 Then
        If txtPatient.Text <> mrsInfo!姓名 Then txtPatient.Text = mrsInfo!姓名
    End If
End Sub

Private Sub txtTotal_GotFocus()
    SelAll txtTotal
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    Dim curMoney As Currency, i As Long
    
    If txtTotal.Locked Then Exit Sub
    If mrsInfo.State = 0 Then KeyAscii = 0: Exit Sub
    If mshDetail.TextMatrix(1, 0) = "" Then KeyAscii = 0: Exit Sub

    If KeyAscii = 13 Then
        If Not IsNumeric(txtTotal.Text) Then
            sta.Panels(2) = "输入错误！": Beep
            txtTotal.Text = txtTotal.Tag
            SelAll txtTotal
        ElseIf Val(txtTotal.Text) <> 0 And Val(txtTotal.Text) > mcurTotal Then
            sta.Panels(2) = "输入金额不能大于本次结帐的金额:" & Format(mcurTotal, mstrDec): Beep
            txtTotal.Text = txtTotal.Tag
            SelAll txtTotal
        Else
            '自动处理合计分配
            sta.Panels(2) = ""
            curMoney = Format(txtTotal.Text, mstrDec)
            mshDetail.Redraw = False
            For i = mshDetail.Rows - 1 To 1 Step -1
                If curMoney = 0 Then
                    mshDetail.TextMatrix(i, COL_结帐金额) = mstrDec
                Else
                    If Val(mshDetail.TextMatrix(i, COL_未结金额)) >= curMoney Then
                        mshDetail.TextMatrix(i, COL_结帐金额) = Format(curMoney, mstrDec)
                    ElseIf Val(mshDetail.TextMatrix(i, COL_未结金额)) < curMoney Then
                        mshDetail.TextMatrix(i, COL_结帐金额) = Format(mshDetail.TextMatrix(i, COL_未结金额), mstrDec)
                    End If
                    curMoney = curMoney - Val(mshDetail.TextMatrix(i, COL_结帐金额))
                End If
            Next
            If curMoney <> 0 Then
                mshDetail.TextMatrix(1, COL_结帐金额) = Format(Val(mshDetail.TextMatrix(1, COL_结帐金额)) + curMoney, mstrDec)
            End If
            Call ShowMoney(True, , mty_ModulePara.bytMzDeposit)
            
            mshDetail.Redraw = True
            mshDeposit.SetFocus
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtTotal_LostFocus()
    Dim blnView As Boolean
    If mbytInState = 1 Then Exit Sub
    blnView = mstrInNO <> "" Or chkCancel.Value = 1 Or cboNO.Text <> ""
    If blnView Then Exit Sub

    If Not IsNumeric(txtTotal.Text) Then
        txtTotal.SetFocus
    ElseIf CCur(txtTotal.Tag) <> CCur(txtTotal.Text) Then
        txtTotal.Text = Format(txtTotal.Tag, mstrDec)
    End If
End Sub

Private Sub AdjustDeposit()
    '功能:初始化预交款列表
    Dim i As Integer
    Call InitDepositGridHead
    With mshDeposit
        .FixedAlignment(.ColIndex("结算方式")) = 1   '考虑到800*600下有滚动条时显不下,左对齐
    End With
End Sub

Private Sub mshDeposit_DblClick()
    Dim lngCardTypeID As Long
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> adStateOpen Then Exit Sub
    If mbytInState <> 0 Or chkCancel.Value <> Unchecked Then Exit Sub
    If txtMoney.Visible Or mshDeposit.Row < 1 Or mshDeposit.Col <> mshDeposit.ColIndex("冲预交") Then Exit Sub
    
    '刘兴洪:增加权限控制，如果仅用预交结帐，结帐数据不正确，之前刘尔旋更改时，为什么能更改，暂时不知道原因，先限制，避免结帐数据出错
    If InStr(mstrPrivs, ";仅用预交结帐;") > 0 Then Exit Sub
    
    '卡类别ID||代扣||是否退现||缺省退现
    lngCardTypeID = Val(Split(mshDeposit.Cell(flexcpData, mshDeposit.Row, mshDeposit.ColIndex("ID")) & "||", "||")(0))
    
    If mblnThreeDepositAfter And lngCardTypeID <> 0 Then Exit Sub
 
    With txtMoney
        .Left = fraBalance.Left + mshDeposit.Left + mshDeposit.CellLeft + 15
        .Top = fraBalance.Top + mshDeposit.Top + mshDeposit.CellTop + (mshDeposit.CellHeight - txtMoney.Height) / 2 - 15
        .Width = mshDeposit.CellWidth - 60
        .ForeColor = mshDeposit.CellForeColor
        .BackColor = mshDeposit.CellBackColor
        .Alignment = 1
        .Text = mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.Col)
        .SelStart = 0: .SelLength = Len(.Text)
        .ZOrder: .Visible = True
        .SetFocus
    End With
 
End Sub

Private Sub mshDeposit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mshDeposit.Row >= 1 Then
            If mshDeposit.Row < mshDeposit.Rows - 1 Then
                mshDeposit.Row = mshDeposit.Row + 1
                mshDeposit.Col = mshDeposit.Cols - 1
                If mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(0) - 2) > 1 Then
                    mshDeposit.TopRow = mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(1) - 2)
                End If
            Else
                vsfMoney.SetFocus
            End If
        End If
    End If
End Sub

Private Sub mshDeposit_KeyPress(KeyAscii As Integer)
   '刘兴洪:增加权限控制，如果仅用预交结帐，结帐数据不正确，之前刘尔旋更改时，为什么能更改，暂时不知道原因，先限制，避免结帐数据出错
    If InStr(mstrPrivs, ";仅用预交结帐;") > 0 Then Exit Sub
    
    
    If Not txtMoney.Visible And mshDeposit.Row >= 1 And mshDeposit.Col = mshDeposit.ColIndex("冲预交") _
        And KeyAscii <> 13 And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
        If mblnThreeDepositAfter And Val(Split(mshDeposit.Cell(flexcpData, mshDeposit.Row, mshDeposit.ColIndex("ID")) & "||", "||")(0)) <> 0 Then
            Exit Sub
        End If
        If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        With txtMoney
            .Left = fraBalance.Left + mshDeposit.Left + mshDeposit.CellLeft + 15
            .Top = fraBalance.Top + mshDeposit.Top + mshDeposit.CellTop + (mshDeposit.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshDeposit.CellWidth - 60
            .ForeColor = mshDeposit.CellForeColor
            .BackColor = mshDeposit.CellBackColor
            .Alignment = 1
            .Text = Chr(KeyAscii)
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshDeposit_LeaveCell()
    txtMoney.Visible = False
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    Dim blnCent As Boolean, i As Long
    
    If KeyAscii <> 13 Then
        '输入限制
        If Not (txtMoney.Left > fraBalance.Left And txtMoney.Top > vsfMoney.Top + fraBalance.Top And vsfMoney.Col = 2) Then
            If InStr(txtMoney.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0: Beep: Exit Sub
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        '结算号码
        Else
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = 0
        sta.Panels(2) = ""
        If Not (txtMoney.Left > fraBalance.Left And txtMoney.Top > vsfMoney.Top + fraBalance.Top And vsfMoney.Col = 2) Then
            If Trim(txtMoney.Text) = "" Then
                sta.Panels(2) = "必须输入金额！"
                SelAll txtMoney: Call Beep: Exit Sub
            ElseIf Not IsNumeric(Trim(txtMoney.Text)) Then
                sta.Panels(2) = "输入了非法金额！"
                SelAll txtMoney: Call Beep: Exit Sub
            End If
        Else '结算号码防拷贝特殊字符
            If InStr(txtMoney.Text, "'") > 0 Or InStr(txtMoney.Text, "|") > 0 Or InStr(txtMoney.Text, ",") > 0 Then
                Call Beep: Exit Sub
            End If
        End If
        If txtMoney.Left < fraBalance.Left Then
            '在费用明细列表内:根据系统参数定小数输入位数
            txtMoney.Text = Format(Val(txtMoney.Text), mstrDec)
            
            '修改不能超过上限
            If Val(txtMoney.Text) > Val(mshDetail.TextMatrix(mshDetail.Row, COL_未结金额)) Then
                txtMoney.Text = Val(mshDetail.TextMatrix(mshDetail.Row, COL_未结金额))
            End If
            
            mshDetail.TextMatrix(mshDetail.Row, mshDetail.Col) = Format(Val(txtMoney.Text), mstrDec)
            
            txtMoney.Visible = False
            
            Call ShowMoney(True, , mty_ModulePara.bytMzDeposit)
            
            If mshDetail.Row = mshDetail.Rows - 1 Then
                '下一控件处理
                mshDeposit.SetFocus
            Else
                '下一行处理
                mshDetail.Row = mshDetail.Row + 1
                If mshDetail.Row - (mshDetail.Height \ mshDetail.RowHeight(0) - 2) > 1 Then
                    mshDetail.TopRow = mshDetail.Row - (mshDetail.Height \ mshDetail.RowHeight(1) - 2)
                End If
                mshDetail.Col = GetColNum("结帐金额") ' mshDetail.Cols - 1
                mshDetail.SetFocus
            End If
        ElseIf txtMoney.Top > fraBalance.Top + vsfMoney.Top Then
            '在结算金额列表内
            If vsfMoney.Col <> 1 Then
                '输入结算号
                vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col) = Trim(txtMoney.Text)
                Call Calc找补
            Else
                '输入结算金额:最多输到0.00
                txtMoney.Text = Format(Val(txtMoney.Text), "0.00")
                
                If Val(txtMoney.Text) <> 0 Then
                    If Val(vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.ColIndex("性质"))) = 1 Then
                        '如果是在现金栏内输入,则如果要处理分币则只准输到0.0
                        blnCent = True
                        If gBytMoney = 0 Then blnCent = False
                        If blnCent And Not IsNull(mrsInfo!险类) Then
                            If Not MCPAR.分币处理 Then blnCent = False
                        End If
                        If blnCent Then txtMoney.Text = Format(CentMoney(Val(txtMoney.Text)), "0.00")
                    ElseIf Val(vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.ColIndex("性质"))) = 3 Then
                        '个人帐户检查
                        If Val(txtMoney.Text) < 0 Then
                            MsgBox "个人帐户结算金额不允许为负数。", vbInformation, gstrSysName
                            Call zlControl.TxtSelAll(txtMoney):  Exit Sub
                        End If
                        '不允许超过返回的原始个帐限额(个人帐户允许透支时再判断)
                        If Val(txtMoney.Text) > mcur个帐限额 And mcur个帐限额 <> 0 And mcur个帐透支 = 0 And mbln个帐结算 Then
                            MsgBox "输入的金额大于了病人可支付的个人帐户限额:" & Format(mcur个帐限额, "0.00") & "。", vbInformation, gstrSysName
                            Call zlControl.TxtSelAll(txtMoney):  Exit Sub
                        End If
                        '不允许超过允许透支金额
                        If mcur个帐余额 - Val(txtMoney.Text) < -1 * mcur个帐透支 Then
                            MsgBox "帐户余额:" & Format(mcur个帐余额, "0.00") & _
                                IIf(mcur个帐透支 = 0, "", "(" & "允许透支:" & Format(mcur个帐透支, "0.00") & ")") & _
                                "不足要结算的金额。", vbInformation, gstrSysName
                            Call zlControl.TxtSelAll(txtMoney):  Exit Sub
                        End If
                    End If
                End If
            
                vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col) = Format(Val(txtMoney.Text), "0.00")
                Call ShowMoney(False, GetDefaultRow <> vsfMoney.Row, mty_ModulePara.bytMzDeposit)   '修改后自动补平,除非当前行是缺省结算方式行
                
                
            End If
            
            txtMoney.Visible = False
            
            If vsfMoney.Col < vsfMoney.Cols - 2 Then
                vsfMoney.Col = vsfMoney.Col + 1
                vsfMoney.SetFocus
            Else
                If vsfMoney.Row = vsfMoney.Rows - 1 Then
                    '下一控件处理
                    If Get应缴 > 0 And txt缴款.Visible Then
                        txt缴款.SetFocus
                    ElseIf cmdOK.Visible And cmdOK.Enabled Then
                        cmdOK.SetFocus
                    End If
                Else
                    '下一行处理
                    vsfMoney.Row = vsfMoney.Row + 1
                    vsfMoney.Col = 1
                    If vsfMoney.Row - (vsfMoney.Height \ vsfMoney.RowHeight(0) - 2) > 1 Then
                        vsfMoney.TopRow = vsfMoney.Row - (vsfMoney.Height \ vsfMoney.RowHeight(1) - 2)
                    End If
                    vsfMoney.SetFocus
                End If
            End If
        Else
            '在冲预交列表内:最多输到0.00
            txtMoney.Text = Format(Val(txtMoney.Text), "0.00")
            
            '修改不能超过上限
            If Val(txtMoney.Text) > Val(mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.ColIndex("余额"))) Then
                txtMoney.Text = Val(mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.ColIndex("余额")))
            End If
            mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.Col) = Format(Val(txtMoney.Text), "0.00")
            
            txtMoney.Visible = False
            Call ShowMoney(False, , mty_ModulePara.bytMzDeposit)
            
            If mshDeposit.Row = mshDeposit.Rows - 1 Then
                '下一控件处理
                If vsfMoney.Enabled And vsfMoney.Visible Then vsfMoney.SetFocus
            Else
                '下一行处理
                mshDeposit.Row = mshDeposit.Row + 1
                If mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(0) - 2) > 1 Then
                    mshDeposit.TopRow = mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(1) - 2)
                End If
                mshDeposit.Col = mshDeposit.Cols - 1
                mshDeposit.SetFocus
            End If
        End If
    End If
End Sub

Private Function ReadBalance(strNo As String) As Boolean
    '功能：查看或作废时,读取并显示结帐单
    '参数：strfullno=单据号
    '返回：
    '     -1:成功
    '      0:失败
    '      1:该单据不存在
    '      2:该单据已作废(mblnViewCancel=True时有效)
    '      3:单据内容不完整
    Dim rsTmp As ADODB.Recordset, rsFee As ADODB.Recordset, rsUnit As ADODB.Recordset
    Dim strFullNO As String, strTable As String, strSql As String
    Dim lngID As Long, lng病人ID As Long
    Dim dMax As Date, dMin As Date
    Dim curDeposit As Currency
    
    On Error GoTo errH
    
    '单据主体
    strFullNO = GetFullNO(strNo, 15)
    
    strTable = IIf(mblnNOMoved, "H", "") & "病人结帐记录"
    strSql = _
    "Select A.ID,A.实际票号 as 票据号,B.病人ID,B.门诊号,B.住院号,Nvl(D.出院病床,B.当前床号) as 当前床号, " & _
    "       Nvl(E.名称,C.名称) as 当前科室," & _
    "       Nvl(D.费别,B.费别) as 费别,nvl(D.姓名,B.姓名) as 姓名,nvl(D.性别,B.性别) as 性别,B.年龄,A.收费时间,A.开始日期,A.结束日期,A.备注,A.原因,A.结帐类型" & _
    " From " & strTable & " A,病人信息 B,部门表 C,病案主页 D,部门表 E" & _
    " Where A.病人ID=B.病人ID(+) And B.当前科室ID=C.ID(+) And D.出院科室ID=E.ID(+)" & _
    "       And B.病人ID=D.病人ID(+) And Nvl(B.主页ID,0)=D.主页ID(+) " & _
    "       And A.NO=[1] And A.记录状态 " & IIf(mblnViewCancel, "= 2", "In(1,3)")
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strFullNO)
    If rsTmp.EOF Then
        MsgBox "没有发现该结帐单据,可能已经作废！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not GetMinMaxDate(rsTmp!ID, dMin, dMax, mblnNOMoved) Then
        MsgBox "该结帐单据内容不正确，没有发现结帐的费用明细！", vbInformation, gstrSysName
        Exit Function
    End If
    mbln合约单位 = False
    cboNO.Text = strFullNO
    txtInvoice.Text = Nvl(rsTmp!票据号)
    
    lng病人ID = Val(Nvl(rsTmp!病人ID))
    If Val(Nvl(rsTmp!结帐类型)) = 0 Then
        lblTitle.Caption = gstrUnitName & "病人结帐单"
    ElseIf Val(Nvl(rsTmp!结帐类型)) = 1 Then
        lblTitle.Caption = gstrUnitName & "门诊病人结帐单"
    Else
        lblTitle.Caption = gstrUnitName & "住院病人结帐单"
    End If
    mint结帐类型 = Val(Nvl(rsTmp!结帐类型))
    
    
    '获取病人余额
    If Val(Nvl(rsTmp!结帐类型)) = 0 Then
        strSql = "Select Sum(预交余额) As 预交余额,Sum(费用余额) As 费用余额 From 病人余额 Where 病人ID= [1] And 性质=1"
        Set rsFee = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID)
    Else
        strSql = "Select 预交余额,费用余额 From 病人余额 Where 病人ID= [1] And 性质=1 And 类型= [2]"
        Set rsFee = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, Val(Nvl(rsTmp!结帐类型)))
    End If
    '60615,刘尔旋,2013-12-20,状态栏显示预交余额、费用金额和剩余余额
    If rsFee.RecordCount <> 0 Then
        sta.Panels(3).Text = "预交:" & Format(Nvl(rsFee!预交余额), "0.00") & _
                             "/费用:" & Format(Nvl(rsFee!费用余额), "0.00") & _
                             "/剩余:" & Format(Val(Nvl(rsFee!预交余额)) - Val(Nvl(rsFee!费用余额)), "0.00")
    End If
    
    '检查是否合约单位结帐:问题:35090
    If Val(Nvl(rsTmp!病人ID)) = 0 Then
        If Nvl(rsTmp!原因) <> "" Then
            txtPatient.Text = Nvl(rsTmp!原因)
        Else
            strSql = "" & _
            "   Select  D.名称 " & _
            "   From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A, 病人信息 C, 合约单位 D " & _
            "   Where A.结帐ID=[1]  And A.病人ID=C.病人ID And C.合同单位id = D.ID(+) and Rownum=1 " & _
            "    Union ALL " & _
            "   Select  D.名称 " & _
            "   From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 A, 病人信息 C, 合约单位 D " & _
            "   Where A.结帐ID=[1] And C.合同单位id = D.ID(+) and Rownum=1 " & _
            "   "
            Set rsUnit = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(rsTmp!ID)))
            If Not rsUnit.EOF Then
                txtPatient.Text = Nvl(rsUnit!名称)
            Else
                txtPatient.Text = "未找到合约单位"
            End If
        End If
        mbln合约单位 = True
        txtPatient.Tag = "合约单位"
    Else
        txtPatient.Text = Nvl(rsTmp!姓名)
        txtPatient.Tag = Val(Nvl(rsTmp!病人ID))
    End If
    
    txtSex.Text = Nvl(rsTmp!性别)
    txtOld.Text = Nvl(rsTmp!年龄)
    txt费别.Text = Nvl(rsTmp!费别)
    txtDate.Text = Format(rsTmp!收费时间, "yyyy-MM-dd HH:mm:ss")
    
    '问题65105,刘尔旋:结账查阅中新增门诊号码的显示
    Select Case Val(Nvl(rsTmp!结帐类型))
        '10.29以前的类型，不做处理
        Case 0
        Case 1
            txt标识号.Text = Nvl(rsTmp!门诊号)
            txt标识号.Visible = True
            lbl标识号.Visible = True
            lbl标识号.Caption = "门诊号"
        Case 2
            txt标识号.Text = Nvl(rsTmp!住院号)
            txt标识号.Visible = True
            lbl标识号.Visible = True
            lbl标识号.Caption = "住院号"
            
            If Not IsNull(rsTmp!当前床号) Then
                txtBed.Text = rsTmp!当前床号
                txtBed.Visible = True
                lblBed.Visible = True
            End If
            
            If Not IsNull(rsTmp!当前科室) Then
                txt科室.Text = rsTmp!当前科室
                txt科室.Visible = True
                lbl科室.Visible = True
            End If
    End Select
    
    txtBegin.Text = Format(dMin, txtBegin.Format)
    txtEnd.Text = Format(dMax, txtEnd.Format)
    txt备注.Text = Nvl(rsTmp!备注)
    If Not IsNull(rsTmp!开始日期) Then
        txtPatiBegin.Text = Format(rsTmp!开始日期, "yyyy-MM-dd")
    End If
    
    If Not IsNull(rsTmp!结束日期) Then
        txtPatiEnd.Text = Format(rsTmp!结束日期, "yyyy-MM-dd")
    End If
    
    lngID = rsTmp!ID
    
    '结帐明细
    '住院费用记录：[住院],[科室],时间,[单据号],项目,费目,[婴儿费],结帐金额
    '------------------------------------------------------------------------------------
    If loadBalanceFeeList(lngID) = False Then Exit Function
    
    '冲预交清单
    Call LoadBalanceDeposit(lngID, curDeposit)
    
    '结帐补款清单,未用的结算方式也列出,以便作废时,不允许的医保结算退现金
    '---------------------------------------------------------------------------------------------------------------------
    If LoadBalanceInfor(lngID, curDeposit, lng病人ID) = False Then Exit Function
    
    If chkCancel.Value = 1 Then
        Call InitPatiRedInvoice
    End If
    
    ReadBalance = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function loadBalanceFeeList(ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结费明细明细
    '入参:lng结帐ID-结帐ID
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-05 14:16:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim i As Long, curTmp As Currency, lngMaxLength As Long
    Dim lngP As Long
    
    On Error GoTo errHandle
    '结帐明细
    '住院费用记录：[住院],[科室],时间,[单据号],项目,费目,[婴儿费],结帐金额
    '------------------------------------------------------------------------------------
    strSql = "" & _
    "   Select  '门诊' as 住院,A.发生时间,A.NO,A.序号,A.收费细目ID,A.收据费目,A.婴儿费,A.结帐金额,A.开单部门ID " & _
    "   From 门诊费用记录 A" & _
    "   Where A.结帐ID=[1]" & _
    "    Union ALL " & _
    "   Select  Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次') as 住院,A.发生时间,A.NO,A.序号,A.收费细目ID,A.收据费目,A.婴儿费,A.结帐金额,A.开单部门ID " & _
    "   From 住院费用记录 A" & _
    "   Where A.结帐ID=[1] " & _
    "   "
    strSql = _
    "  Select   A.住院," & _
    "            Nvl(B.名称,'未知') as 科室,To_Char(A.发生时间,'YYYY-MM-DD') as 时间," & _
    "            A.NO as 单据号,Nvl(E.名称,D.名称) as 项目,A.收据费目 as 费目," & _
    "            Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.结帐金额" & _
    " From (" & strSql & ") A,部门表 B,收费项目目录 D,收费项目别名 E" & _
    " Where A.开单部门ID=B.ID(+) And A.收费细目ID=D.ID" & _
    "           And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    " Order by 住院 Desc,时间 Desc,单据号 Desc,序号"
    If mblnNOMoved Then strSql = Replace(Replace(strSql, "门诊费用记录", "H门诊费用记录"), "住院费用记录", "H住院费用记录")
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
    If rsTmp.EOF Then Exit Function
    With mshDetail
        .Redraw = False
        Call ClearDetail
        If Not rsTmp.EOF Then Set .DataSource = rsTmp
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4
            If i <= 4 Then .MergeCol(i) = True
            Select Case .TextMatrix(0, i)
                Case "住院", "婴儿费", "单据号"
                    .ColAlignment(i) = 4
                Case "科室", "项目", "费目", "时间"
                    .ColAlignment(i) = 1
                Case "结帐金额"
                    .ColAlignment(i) = 7
            End Select
            
            Select Case .TextMatrix(0, i)
                Case "科室", "住院"
                    .ColWidth(i) = 0
                Case "婴儿费"
                    .ColWidth(i) = 750
                Case "费目"
                    .ColWidth(i) = 800
                Case "结帐金额", "单据号"
                    .ColWidth(i) = 950
                Case "时间"
                    .ColWidth(i) = 1130
                Case "项目"
                    .ColWidth(i) = 2300
            End Select
            .ColData(i) = .ColWidth(i)
        Next
        
        lngMaxLength = Len(Mid(gstrDec, 3))
        If rsTmp.RecordCount > 0 Then
            For i = 1 To rsTmp.RecordCount
                lngP = InStr(1, CStr(rsTmp!结帐金额), ".")
                If lngP > 0 Then
                    lngP = Len(Mid(CStr(rsTmp!结帐金额), lngP + 1))
                    If lngP > lngMaxLength Then lngMaxLength = lngP
                End If
                rsTmp.MoveNext
            Next
            rsTmp.MoveFirst
        End If
        mstrDec = "0." & String(lngMaxLength, "0")
        
        curTmp = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, .Cols - 1) = Format(.TextMatrix(i, .Cols - 1), mstrDec)
            curTmp = curTmp + Val(.TextMatrix(i, .Cols - 1))
        Next
        .Redraw = True
    End With
    
    txtTotal.Text = Format(curTmp, mstrDec)
    curTmp = Val(txtTotal.Text)

    loadBalanceFeeList = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If


End Function
Private Function GetDefaultRow() As Long
'功能：获取当前缺省结算方式行号
    Dim i As Long, lngDefaultRow As Long, curBalance As Currency, curDeposit As Currency
    Dim str住院次数 As String, strSql As String, rsTmp As ADODB.Recordset
    
    If mblnOneCard And mstrOneCard <> "" Then
        For i = 1 To vsfMoney.Rows - 1
            If vsfMoney.TextMatrix(i, vsfMoney.ColIndex("结算方式")) = mstrOneCard Then
                lngDefaultRow = i: Exit For
            End If
        Next
    Else
        If mstr缺省结算 <> "" Then
            For i = 1 To vsfMoney.Rows - 1
                If vsfMoney.TextMatrix(i, vsfMoney.ColIndex("结算方式")) = mstr缺省结算 Then
                    lngDefaultRow = i: Exit For
                End If
            Next
        Else
            '78882:结账退款缺省按预交缴款结算方式退款：如果没有选择这个参数，缺省按现金退款
            '如果预交缴款有多种结算方式，按下列顺序处理
            '        1.银行卡(手工处理的银行卡,性质为2并且非支票的结算方式)
            '        2.现金
            '        3.支票
            '        4.其他结算方式
            If mbytFunc = 1 Then
                curBalance = GetBalanceSum
                For i = 1 To mshDeposit.Rows - 1
                    curDeposit = curDeposit + Val(mshDeposit.TextMatrix(i, mshDeposit.ColIndex("冲预交")))
                Next i
                If curDeposit > curBalance Then
                    If mty_ModulePara.bln结帐退款方式 = False Then
                        '缺省退现金
                        For i = 1 To vsfMoney.Rows - 1
                            If Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("性质"))) = 1 Then  '没有指定缺省时以现金为缺省行
                                 lngDefaultRow = i
                                 GetDefaultRow = lngDefaultRow
                                 Exit Function
                            End If
                        Next
                    Else
                        '缺省退预交缴款结算方式
                        str住院次数 = ""
                        If mbytFunc = 1 Then
                            str住院次数 = IIf(gbln仅用指定预交款 And mbln门诊转住院 = False, IIf(mstrTime = "", mstrAllTime, mstrTime), "")
                        End If
                        
                        strSql = " Select a.结算方式, Decode(Nvl(b.性质,0), 7, 1, 2, Decode(a.结算方式,'支票',4,2), 1, 3, 5) As 顺序 From 病人预交记录 A,结算方式 B " & _
                                 " Where a.记录性质 = 1 And a.病人id = [1] And a.预交类别 = 2 And a.结算方式 = b.名称(+) " & _
                                 IIf(str住院次数 = "", "", " And a.主页ID In (Select Column_Value From Table(f_str2list([2]))) ") & _
                                 " Order By 顺序 "
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsInfo!病人ID)), str住院次数)
                        If rsTmp.RecordCount <> 0 Then
                            For i = 1 To vsfMoney.Rows - 1
                                If vsfMoney.TextMatrix(i, vsfMoney.ColIndex("结算方式")) = Nvl(rsTmp!结算方式) Then
                                     lngDefaultRow = i
                                     GetDefaultRow = lngDefaultRow
                                     Exit Function
                                End If
                            Next
                        End If
                    End If
                End If
            End If
            For i = 1 To vsfMoney.Rows - 1
                If Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("性质"))) = 1 Then  '没有指定缺省时以现金为缺省行
                     lngDefaultRow = i: Exit For
                End If
            Next
        End If
    End If
    
    GetDefaultRow = lngDefaultRow
End Function

Private Function GetBalanceSum() As Currency
    Dim i As Long, cur结帐合计 As Currency
    Dim lngCol As Long
    lngCol = GetColNum("结帐金额")
    
    If lngCol <> COL_结帐金额 Then Exit Function
    
    For i = 1 To mshDetail.Rows - 1
        cur结帐合计 = cur结帐合计 + Val(mshDetail.TextMatrix(i, lngCol))
    Next
    GetBalanceSum = cur结帐合计
End Function

Private Function ShowMoney(blnFirst As Boolean, _
    Optional blnAutoCalc As Boolean = True, Optional bytMzDeposit As Byte = 2, _
    Optional blnRecalDeposit As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置和显示界面的各种金额
    '入参:blnFirst=是否重新处理结帐明细,冲预交额,医保结算部份,就象第一次调用本函数一样
    '     blnAutoCalc=根据差额自动补平并计算
    '     bytMzDeposit-针对门诊结帐有效,0-表示全清;1-代表根据结帐金额来分摊预交;2-预交款全冲
    '出参:
    '返回:医保可报销结算部分未被设置提示串
    '编制:刘兴洪
    '日期:2014-05-23 16:11:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim lng缺省Row As Long, bln缺省现金 As Boolean, i As Long, j As Long, lng误差 As Long
    Dim cur结帐合计 As Currency, curMoney As Currency, curTemp As Currency
    Dim strMoney As String, strNone As String, strHave As String
    Dim blnCent As Boolean, curOwn As Currency, curTmp As Currency
    Dim varData As Variant
     
    
    '判断缺省结算方式是否现金，是现金则自动补平时处理分币，否则仅处理误差
    '如果没有设置缺省结算方式，则将现金作为缺省的补平结算方式(如果有)
    '-----------------------------------------------------------------------------------------------------
    lng缺省Row = GetDefaultRow
    For i = 1 To vsfMoney.Rows - 1
        If Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("性质"))) = 9 Then
            vsfMoney.TextMatrix(i, vsfMoney.ColIndex("金额")) = 0
            lng误差 = i: Exit For
        End If
    Next i
    If lng缺省Row > 0 Then bln缺省现金 = (Val(vsfMoney.TextMatrix(lng缺省Row, vsfMoney.ColIndex("性质"))) = 1)
    
    '判断是否应该进行分币处理
    blnCent = True
    If gBytMoney = 0 Then blnCent = False
    If Not IsNull(mrsInfo!险类) And Not MCPAR.分币处理 Then blnCent = False
    
    '显示结帐合计及设置冲预交和各种结算金额
    '-----------------------------------------------------------------------------------------------------
    If blnFirst Then
        '统计并显示结帐金额合计
        cur结帐合计 = GetBalanceSum
        txtTotal.Text = Format(cur结帐合计, mstrDec)
        txtTotal.Tag = txtTotal.Text
            
        '设置医保结算部分金额
        For i = 0 To UBound(Split(mstrBalance, "|"))
            strMoney = Split(mstrBalance, "|")(i)
            For j = 1 To vsfMoney.Rows - 1
                If vsfMoney.TextMatrix(j, vsfMoney.ColIndex("结算方式")) = CStr(Split(strMoney, ";")(0)) _
                    And InStr(",3,4,", Val(vsfMoney.TextMatrix(j, vsfMoney.ColIndex("性质")))) > 0 Then
                    '个人帐户不超过余额
                    If Val(vsfMoney.TextMatrix(j, vsfMoney.ColIndex("性质"))) = 3 Then
                        '个人帐户最大支付金额
                        mbln个帐结算 = True
                        mcur个帐限额 = CCur(Split(strMoney, ";")(1))
                        
                        '缺省不能超过个人帐户余额或允许透支金额
                        If mcur个帐余额 - CCur(Split(strMoney, ";")(1)) >= -1 * mcur个帐透支 Then
                            vsfMoney.TextMatrix(j, vsfMoney.ColIndex("金额")) = Format(CCur(Split(strMoney, ";")(1)), "0.00") '在允许透支范围内足够(允许透支0为特例)
                        Else
                            vsfMoney.TextMatrix(j, vsfMoney.ColIndex("金额")) = "0.00"
                            MsgBox "个人帐户余额不足或未更新,不允许医保结算!", vbInformation, Me.Caption
                            cmdOK.Enabled = False
                        End If
                    Else
                        vsfMoney.TextMatrix(j, vsfMoney.ColIndex("金额")) = Format(CCur(Split(strMoney, ";")(1)), "0.00")
                    End If
                    
                    If Val(Split(strMoney, ";")(2)) = 0 Then
                        vsfMoney.RowData(j) = 1 '该结算金额不可更改
                    Else
                        vsfMoney.RowData(j) = 0 '该结算金额可以更改
                    End If
                    
                    '加入医保已处理的结算
                    cur结帐合计 = cur结帐合计 - Format(Val(vsfMoney.TextMatrix(j, vsfMoney.ColIndex("金额"))), "0.00")
                    strHave = strHave & ";" & CStr(Split(strMoney, ";")(0))
                    Exit For
                End If
            Next
            '未包含医保可报销结算方式
            If j = vsfMoney.Rows Then
                strNone = strNone & vbCrLf & vbTab & CStr(Split(strMoney, ";")(0)) & ":" & Format(CCur(Split(strMoney, ";")(1)), "0.00")
            End If
        Next
        
        '刘兴洪:针对结算卡进行处理
        Call zlReCalcRequare(cur结帐合计, strNone)
        
         
        '设置冲预交(结帐合计 - 保险合计)
        If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
            For i = 1 To vsfMoney.Rows - 1
                If vsfMoney.RowData(i) = 999 Then
                    vsfMoney.TextMatrix(i, 1) = Format(cur结帐合计, "0.00")
                End If
            Next i
        End If
        
        Call RecalDepsit(mbytFunc, bytMzDeposit, cur结帐合计, mblnShowThree, mblnThreeDepositAll)
         
        cur结帐合计 = FormatEx(cur结帐合计 - Val(lblDelMoney.Tag), 6)
                    
        '剩余应缴部份尝试设置到缺省结算方式
        If lng缺省Row <> 0 Then
        
            If bln缺省现金 And blnCent Then '现金时要进行分币处理
                vsfMoney.TextMatrix(lng缺省Row, vsfMoney.ColIndex("金额")) = Format(CentMoney(cur结帐合计), "0.00")
            Else
                vsfMoney.TextMatrix(lng缺省Row, vsfMoney.ColIndex("金额")) = Format(cur结帐合计, "0.00")
            End If
            cur结帐合计 = 0
        End If
    Else
        If blnRecalDeposit Then
            '设置冲预交(结帐合计 - 保险合计)
            If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
                For i = 1 To vsfMoney.Rows - 1
                    If vsfMoney.RowData(i) = 999 Then
                        vsfMoney.TextMatrix(i, 1) = Format(cur结帐合计, "0.00")
                    End If
                Next i
            End If
            Call RecalDepsit(mbytFunc, bytMzDeposit, cur结帐合计, mblnShowThree, mblnThreeDepositAll)
        Else
            '显示三方帐户退款金额
            Call ShowDelThreeSwap
        End If
    End If
    
    
    
    '显示当前冲预交额及差额
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetPaySum
    
    '这里是差额,不一定用现金,所以不处理分币,lblDelMoney.Tag本次退到三方帐户的金额
    curOwn = Val(txtTotal.Text) - curMoney
    txtOwe.Text = Format(curOwn, "0.00")
    
    '根据差额自动补平并计算'剩余部份尝试设置到缺省结算方式上
    '-----------------------------------------------------------------------------------------------------
    If blnAutoCalc And Val(txtOwe.Text) <> 0 And lng缺省Row <> 0 Then
        curTmp = Val(vsfMoney.TextMatrix(lng缺省Row, vsfMoney.ColIndex("金额"))) + curOwn
        If Abs(curTmp) >= 0.01 Then
            If bln缺省现金 And blnCent Then
                vsfMoney.TextMatrix(lng缺省Row, vsfMoney.ColIndex("金额")) = Format(CentMoney(curTmp), "0.00")
            Else
                vsfMoney.TextMatrix(lng缺省Row, vsfMoney.ColIndex("金额")) = Format(curTmp, "0.00")
            End If
        Else
            vsfMoney.TextMatrix(lng缺省Row, vsfMoney.ColIndex("金额")) = "0.00"
        End If
        txtOwe.Text = "0.00"
    End If
    
    '计算误差金额(结算金额-结帐金额)
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetPaySum

    '有可能应补差额正好是处理分币的误差部份,就不显示了
    If Val(txtOwe.Text) <> 0 And lng缺省Row <> 0 And bln缺省现金 And blnCent Then
        If Abs(Val(txtOwe.Text)) < 0.1 Or gBytMoney = 5 And Abs(Val(txtOwe.Text)) < 0.3 Then
            If CentMoney(Val(vsfMoney.TextMatrix(lng缺省Row, vsfMoney.ColIndex("金额"))) + Val(txtOwe.Text)) = Val(vsfMoney.TextMatrix(lng缺省Row, vsfMoney.ColIndex("金额"))) Then
                txtOwe.Text = "0.00"
            End If
        End If
    End If
    
    '可能应补部份是小数点的正常误差部份,如果四舍五入小于1分,就不显示了
    If Val(txtOwe.Text) <> 0 And mcur误差金额 + curOwn = 0 And Abs(curOwn) <= 0.005 Then
        txtOwe.Text = "0.00"
    End If
    'txtOwe.ToolTipText = "误差金额:" & Format(mcur误差金额, mstrDec)
        
    If lng误差 <> 0 And Val(txtOwe.Text) = 0 Then
        vsfMoney.TextMatrix(lng误差, vsfMoney.ColIndex("金额")) = Format(Val(txtTotal.Text) - curMoney, mstrDec)
        If Val(txtTotal.Text) - curMoney <> 0 Then
            vsfMoney.RowHidden(lng误差) = False
        Else
            vsfMoney.RowHidden(lng误差) = True
        End If
    Else
        mcur误差金额 = Format(curMoney - Val(txtTotal.Text), mstrDec)
        'vsfMoney.TextMatrix(lng误差, vsfMoney.ColIndex("金额")) = Format(vsfMoney.TextMatrix(lng误差, vsfMoney.ColIndex("金额")), mstrDec)
    End If
    
    
    
    curMoney = 0
    If mshDeposit.TextMatrix(1, mshDeposit.ColIndex("ID")) <> "" Then
        For i = 1 To mshDeposit.Rows - 1
            curMoney = curMoney + Val(mshDeposit.TextMatrix(i, mshDeposit.ColIndex("冲预交")))
        Next
    End If
    lblDeposit.Caption = "冲预交:" & Format(curMoney, "0.00")

    Call Calc找补
    If gblnLED Then
        curTmp = Get应缴
        zl9LedVoice.DisplayBank "总费用" & Format(txtTotal.Text, "0.00"), "预交款" & Format(lblSpare.Tag, "0.00"), _
                "冲预交" & Format(curMoney, "0.00"), IIf(curTmp < 0, "找补", "应缴") & Format(Abs(curTmp), "0.00")
    End If
    
    '返回提示
    '-----------------------------------------------------------------------------------------------------
    If strNone <> "" Then
        ShowMoney = "结帐场合的保险结算方式未设置完全,该病人还有以下保险结算方式可以报销：" & _
            vbCrLf & strNone & vbCrLf & vbCrLf & "您可以到费用基础项目\结算方式管理中去设置这些结算方式！"
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPaySum() As Currency
'功能：获取付款合计，包括冲预交和输入的各种付款方式金额
    Dim i As Long, curMoney As Currency
    
    If mshDeposit.TextMatrix(1, mshDeposit.ColIndex("ID")) <> "" Then
        For i = 1 To mshDeposit.Rows - 1
            curMoney = curMoney + Val(mshDeposit.TextMatrix(i, mshDeposit.ColIndex("冲预交")))
        Next
    End If
    
    For i = 1 To vsfMoney.Rows - 1
        If IsNumeric(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("金额"))) And vsfMoney.RowData(i) <> "999" Then
            curMoney = curMoney + Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("金额")))
        End If
    Next
    curMoney = curMoney - Val(lblDelMoney.Tag)
    GetPaySum = curMoney
End Function
Public Function Zl病人费用来源() As Byte
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取病人费来源信息
    '返回：0-权门诊;1-仅住院;2-门诊和住院(暂不能无此数据)
    '编制：刘兴洪
    '日期：2010-03-09 17:39:26
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim BytType As Byte
    '获取费用获取范围类型:'bytKind: 0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
    If mbytFunc = 0 Then BytType = 0
    If mbytFunc = 1 Then BytType = 1
    '刘兴洪:现在只分门诊和住院结帐;因此,取消以下判断
'''    If mbytKind = 1 Then '仅体检费用
'''        BytType = 0
'''    ElseIf (InStr(mstrPrivs, "住院费用结帐") = 0 Or mbytMCMode = 1) Then  '门诊部分的处理
'''            If InStr(mstrPrivs, "门诊费用结帐") = 0 Then
'''                '无权限,又处理门诊结帐数据的:
'''                ' a: 3-其他(就诊卡等额外的收费);4-体检
'''                BytType = IIf(mbytKind = 0, 1, 0) '如果是就诊卡,就读住院费用记录,否则读门诊费用记录
'''            Else
'''                '有门诊结算权限
'''                'a: 1-门诊,3-其他(就诊卡等额外的收费);4-体检
'''                BytType = IIf(mbytKind = 0, 2, 0)
'''            End If
'''    ElseIf InStr(mstrPrivs, "门诊费用结帐") = 0 Then    '住院结算,但不能结帐门诊的
'''        '2-住院;3-其他(就诊卡等额外的收费);4-体检
'''        BytType = IIf(mbytKind = 0, 1, 2)
'''    Else  '门诊和住院
'''        '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
'''        BytType = 2
'''    End If
    Zl病人费用来源 = BytType
End Function
Private Function Is门诊留观(ByVal lng病人ID As Long, ByRef lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前费用是否在门诊留观病人费用期间
    '入参:lng病人ID
    '出参:lng主页ID-返回当前病人ID(第几次留观的)
    '返回:
    '编制:刘兴洪
    '日期:2012-01-10 12:07:52
    '问题:45302
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, dtStartDate As Date, dtEndDate As Date
    Dim str时间 As String, strCond As String, rsTemp As ADODB.Recordset
    str时间 = IIf(gint费用时间 = 0, "A.登记时间", "A.发生时间")
    strCond = "": dtStartDate = CDate("1901-01-01"): dtEndDate = dtStartDate
    If Not mDateBegin = CDate("0:00:00") Then
        strCond = " " & str时间 & " Between [2] And [3]"
        dtStartDate = CDate(Format(mDateBegin, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(mDateEnd, "yyyy-MM-dd 23:59:59"))
    End If
    gstrSQL = "" & _
    "Select A.主页id " & _
    "   From 病案主页 A, " & _
    "        (Select Min(" & str时间 & ") As 最小费用时间, Max(" & str时间 & " ) 最大费用时间 " & _
    "          From 门诊费用记录 A " & _
    "          Where  病人id = [1] " & strCond & ") B " & _
    "   Where A.病人id = [1] And A.病人性质 = 1  " & _
    "       And (B.最小费用时间 Between A.入院日期 And Nvl(A.出院日期, Sysdate) Or " & _
    "                B.最大费用时间 Between A.入院日期 And Nvl(A.出院日期, Sysdate) Or " & _
    "                A.入院日期 Between B.最小费用时间 And B.最大费用时间 Or " & _
    "                Nvl(A.出院日期, Sysdate) Between B.最小费用时间 And B.最大费用时间)" & _
    "   Order by 主页ID Desc"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, dtStartDate, dtEndDate)
    If rsTemp.EOF Then rsTemp.Close: Set rsTemp = Nothing: Exit Function
    lng主页ID = Val(Nvl(rsTemp!主页ID))
    rsTemp.Close: Set rsTemp = Nothing
    Is门诊留观 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveBalance(objCard As Card, ByRef strNo As String, ByRef curDate As Date, str病历原因 As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对当前结帐单存盘处理
    '入参:objCard－卡三方卡
    '出参:strNo-返回结帐单号
    '     Curdate-当前结帐时间
    '返回:结帐ID
    '编制:刘兴洪
    '日期:2015-04-27 10:52:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrSQL() As Variant
    Dim lng结帐ID As Long, i As Long, j As Long, lngTmp As Long, intInsure As Integer
    Dim str费用IDs As String, str费用ID As String, str误差NO As String, strTmp As String
    Dim cur结帐金额合计 As Currency, str保险结算 As String, str保险信息 As String, strAdvance As String
    Dim bln医保结算校对 As Boolean, blnTrans As Boolean, blnTransMC As Boolean
    Dim cur个人帐户 As Currency, cur医保基金 As Currency, intMaxTime As Integer
    Dim cur缴款 As Currency, cur找补 As Currency, cur预交余额 As Currency, cur冲预交 As Currency, cur预交余额合计 As Currency, cur冲预交合计 As Currency
    Dim lng主页ID As Long, rsTmp As ADODB.Recordset
    Dim curOneCard As Currency, dblOneCardBalance As Double
    Dim strCardNo  As String, intCardType As Integer, strTransFlow As String
    Dim BytType As Byte, str住院次数 As String, strSql As String
    Dim blnMustCommit As Boolean
    Dim rsDeposit As ADODB.Recordset
    
    Screen.MousePointer = 11
    On Error GoTo ErrHand:
    arrSQL = Array()
    strNo = zlDatabase.GetNextNo(15)
    lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
    curDate = zlDatabase.Currentdate
    intInsure = Nvl(mrsInfo!险类, 0)
    If intInsure <> 0 Then str保险信息 = Nvl(mrsInfo!险类, " ") & "," & Nvl(mrsInfo!密码, " ") & "," & Nvl(mrsInfo!医保号, " ")
    intMaxTime = GetMinMaxTime(1)
    cur缴款 = Val(txt缴款.Text)
    cur找补 = Val(txt找补.Text)
    
    '0-仅门诊;1-仅住院;2-门诊和住院
    BytType = zlGetPatiSource
    blnMustCommit = False
    '1.病人结帐记录
    '问题:25596
    ' Zl_病人结帐记录_Insert
    strTmp = "zl_病人结帐记录_Insert("
    '  Id_In           病人结帐记录.ID%Type,
    strTmp = strTmp & "" & lng结帐ID & ","
    '  单据号_In       病人结帐记录.NO%Type,
    strTmp = strTmp & "'" & strNo & "',"
    '  病人id_In       病人结帐记录.病人id%Type,
    strTmp = strTmp & "" & Val(Nvl(mrsInfo!病人ID)) & ","
    '  收费时间_In     病人结帐记录.收费时间%Type,
    strTmp = strTmp & "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  开始日期_In     病人结帐记录.开始日期%Type,
    strTmp = strTmp & "" & IIf(IsDate(txtPatiBegin.Text), "To_Date('" & txtPatiBegin.Text & "','YYYY-MM-DD')", "NULL") & ","
    '  结束日期_In     病人结帐记录.结束日期%Type,
    strTmp = strTmp & "" & IIf(IsDate(txtPatiEnd.Text), "To_Date('" & txtPatiEnd.Text & "','YYYY-MM-DD')", "NULL") & ","
    '  中途结帐_In     病人结帐记录.中途结帐%Type := 0,
    strTmp = strTmp & "" & IIf(opt中途.Value, 1, 0) & ","
    '  多病人结帐_In   Number := 0,
    strTmp = strTmp & "" & 0 & ","
    '  最大结帐次数_In Number := 0,
    strTmp = strTmp & "" & intMaxTime & ","
    '  备注_In         病人结帐记录.备注%Type := Null
    strTmp = strTmp & "" & IIf(Trim(txt备注.Text) = "", "NULL", "'" & Trim(txt备注.Text) & "'") & ","
    '   来源_In         Number := 1,1-门诊;2-住院
    strTmp = strTmp & "" & BytType & ","
    '  原因_In         病人结帐记录.原因%Type := Null
    strTmp = strTmp & "" & IIf(Trim(str病历原因) = "", "NULL", "'" & Trim(str病历原因) & "'") & ","
    '    结帐类型_In     病人结帐记录.结帐类型%type:=2
    strTmp = strTmp & "" & IIf(mbytFunc = 0, 1, 2) & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strTmp: strTmp = ""
       
    '2.病人预交记录-冲预交：[ID],[NO],日期,结算方式,金额,金额
    With mshDeposit
        '
        If .TextMatrix(1, .ColIndex("ID")) <> "" Then
            '重读可用预交,并发操作判断
            Set rsDeposit = GetDeposit(mrsInfo!病人ID, mblnDateMoved, IIf(gbln仅用指定预交款, IIf(mstrTime = "", mstrAllTime, mstrTime), ""), , , mint预交类别)
            For i = 1 To .Rows - 1
                cur预交余额 = Val(.TextMatrix(i, .ColIndex("余额")))
                cur冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
                If cur冲预交 <> 0 Then
                    rsDeposit.Filter = "ID=" & CLng(.TextMatrix(i, .ColIndex("ID"))) & " And NO='" & .TextMatrix(i, .ColIndex("单据号")) & "' And 记录状态=" & .RowData(i) & " And 金额=" & cur预交余额
                    If rsDeposit.RecordCount = 0 Then
                        Call MsgBox("由于并发操作,病人预交款已发生变化,请重新提取病人结帐!", vbInformation, gstrSysName)
                        Screen.MousePointer = 0
                        Exit Function
                    End If
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "zl_结帐预交记录_Insert(" & CLng(.TextMatrix(i, .ColIndex("ID"))) & "," & _
                        "'" & .TextMatrix(i, .ColIndex("单据号")) & "'," & .RowData(i) & "," & _
                        cur冲预交 & "," & lng结帐ID & "," & mrsInfo!病人ID & ")"
                    cur冲预交合计 = cur冲预交合计 + cur冲预交
                End If
                cur预交余额合计 = cur预交余额合计 + cur预交余额
            Next
            '结帐冲过的预交单据在预交款管理中被作废后,会出现负的预交余额单据
            If cur冲预交合计 > cur预交余额合计 And cur冲预交合计 <> 0 Then
                Call MsgBox("可用预交余额不足冲款金额!", vbInformation, gstrSysName)
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
    End With
    
    lng主页ID = Val(Nvl(mrsInfo!主页ID))
    If lng主页ID = 0 Or mbytMCMode = 1 Or mbytFunc = 0 Then
        '门诊留观,需要保存主页ID
        '问题:45302
        If Nvl(mrsInfo!病人性质, 0) <> 1 And lng主页ID <> 0 Then
            '当前病人不是留观
              If Not Is门诊留观(mrsInfo!病人ID, lng主页ID) Then
                    lng主页ID = 0
              End If
        End If
    End If
    
    '3.病人预交记录-结帐补：结算方式,金额,结算号码
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("金额"))) <> 0 And .RowData(i) <> "999" Then
                '医保存储:缴款单位=保险类别,单位开户行=密码,单位帐号=医保号
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                
                arrSQL(UBound(arrSQL)) = _
                    "zl_结帐缴款记录_Insert('" & strNo & "'," & mrsInfo!病人ID & "," & lng主页ID & "," & _
                    IIf(IsNull(mrsInfo!当前科室id), 0, mrsInfo!当前科室id) & "," & _
                    "'" & .TextMatrix(i, .ColIndex("结算方式")) & "','" & .TextMatrix(i, .ColIndex("结算号码")) & "'," & _
                    CCur(.TextMatrix(i, .ColIndex("金额"))) & "," & lng结帐ID & ",'" & UserInfo.编号 & "'," & _
                    "'" & UserInfo.姓名 & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    IIf(InStr(",3,4,", Val(.TextMatrix(i, .ColIndex("性质")))) > 0, IIf(IsNull(mrsInfo!险类), "NULL", mrsInfo!险类), "NULL") & "," & _
                    IIf(InStr(",3,4,", Val(.TextMatrix(i, .ColIndex("性质")))) > 0, "'" & IIf(IsNull(mrsInfo!医保号), "", mrsInfo!医保号) & "'", "NULL") & "," & _
                    IIf(InStr(",3,4,", Val(.TextMatrix(i, .ColIndex("性质")))) > 0, "'" & IIf(IsNull(mrsInfo!密码), "", mrsInfo!密码) & "'", "NULL") & _
                    IIf(cur缴款 <> 0, "," & cur缴款 & "," & cur找补, ",Null,Null") & ",Null,Null,Null,'" & IIf(Val(.TextMatrix(i, .ColIndex("性质"))) = 1, mstrForceNote, "") & "')"
                    
                    cur缴款 = 0
                If intInsure <> 0 Then
                    '"结算方式|结算金额||....."
                    If InStr(",3,4,", Val(.TextMatrix(i, .ColIndex("性质")))) > 0 Then str保险结算 = str保险结算 & "||" & .TextMatrix(i, .ColIndex("结算方式")) & "|" & Val(.TextMatrix(i, .ColIndex("金额")))
                    If Val(.TextMatrix(i, .ColIndex("性质"))) = 3 Then cur个人帐户 = cur个人帐户 + Val(.TextMatrix(i, .ColIndex("金额")))
                    If Val(.TextMatrix(i, .ColIndex("性质"))) = 4 Then cur医保基金 = cur医保基金 + Val(.TextMatrix(i, .ColIndex("金额")))
                End If
                
                If mblnOneCard And Not mobjICCard Is Nothing Then
                    If .TextMatrix(i, .ColIndex("结算方式")) = mrsOneCard!结算方式 Then '在保存之前检查,只能使用一种一卡通结算方式
                        curOneCard = CCur(.TextMatrix(i, .ColIndex("金额")))
                    End If
                End If
            End If
        Next
        '退三方交易
       
        If Not objCard Is Nothing Then
            strTmp = "zl_结帐缴款记录_Insert("
            '    No_In         病人结帐记录.No%Type,
            strTmp = strTmp & "'" & strNo & "',"
            '    病人id_In     病人预交记录.病人id%Type,
            strTmp = strTmp & "" & mrsInfo!病人ID & ","
            '    主页id_In     病人预交记录.主页id%Type,
            strTmp = strTmp & "" & lng主页ID & ","
            '    科室id_In     病人预交记录.科室id%Type,
            strTmp = strTmp & "" & IIf(IsNull(mrsInfo!当前科室id), 0, mrsInfo!当前科室id) & ","
            '    结算方式_In   病人预交记录.结算方式%Type,
            strTmp = strTmp & "'" & objCard.结算方式 & "',"
            '    结算号码_In   病人预交记录.结算号码%Type,
            strTmp = strTmp & "NULL,"
            '    金额_In       病人预交记录.冲预交%Type,
            strTmp = strTmp & "" & -1 * mCurBrushCard.dblMoney & ","
            '    结帐id_In     病人预交记录.结帐id%Type,
            strTmp = strTmp & "" & lng结帐ID & ","
            '    操作员编号_In 病人预交记录.操作员编号%Type,
            strTmp = strTmp & "'" & UserInfo.编号 & "',"
            '    操作员姓名_In 病人预交记录.操作员姓名%Type,
            strTmp = strTmp & "'" & UserInfo.姓名 & "',"
            '    收费时间_In   病人预交记录.收款时间%Type,
            strTmp = strTmp & "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '    保险类别_In   保险帐户.险类%Type,
            strTmp = strTmp & "NULL,"
            '    保险帐号_In   保险帐户.医保号%Type,
            strTmp = strTmp & "NULL,"
            '    保险密码_In   保险帐户.密码%Type,
            strTmp = strTmp & "NULL,"
            '    缴款_In       病人预交记录.缴款%Type := Null,
            strTmp = strTmp & "" & IIf(cur缴款 <> 0, cur缴款, "NULL") & ","
            '    找补_In       病人预交记录.找补%Type := Null,
            strTmp = strTmp & "" & IIf(cur找补 <> 0, cur找补, "NULL") & ","
            '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strTmp = strTmp & "" & objCard.接口序号 & ","
            '    卡号_In       病人预交记录.卡号%Type := Null,
            strTmp = strTmp & "'" & mCurBrushCard.str卡号 & "',"
            '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strTmp = strTmp & "'" & mCurBrushCard.str交易流水号 & "',"
            '    交易说明_In   病人预交记录.交易说明%Type := Null
            strTmp = strTmp & "'" & mCurBrushCard.str交易说明 & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strTmp
        End If
    End With
    If str保险结算 <> "" Then str保险结算 = Mid(str保险结算, 3)
    
    '4.住院费用记录：住院,期间,科室,日期,[单据号],项目,费目,婴儿费,[ID],[序号],[记录性质],[记录状态],[执行状态],[A.主页ID],[A.开单部门ID],未结金额,结帐金额
    With mshDetail
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_结帐金额)) <> 0 Or Val(.TextMatrix(i, COL_未结金额)) = 0 Then
                'a.结剩余帐,或首次结帐但部分结
                If Val(.TextMatrix(i, COL_ID)) = 0 Or Val(.TextMatrix(i, COL_未结金额)) <> Val(.TextMatrix(i, COL_结帐金额)) Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "zl_结帐费用记录_Insert(" & .TextMatrix(i, COL_ID) & "," & _
                        "'" & .TextMatrix(i, COL_单据号) & "'," & .TextMatrix(i, COL_记录性质) & "," & _
                         .TextMatrix(i, COL_记录状态) & "," & Val(.TextMatrix(i, COL_执行状态)) & "," & _
                         .TextMatrix(i, COL_序号) & "," & CCur(.TextMatrix(i, COL_结帐金额)) & "," & _
                         lng结帐ID & ")"
                Else
                'b.首次结帐并且全结
                    str费用IDs = str费用IDs & .TextMatrix(i, COL_ID) & ","
                End If
                If intInsure <> 0 Then cur结帐金额合计 = cur结帐金额合计 + CCur(.TextMatrix(i, COL_结帐金额))
            End If
        Next
                
        While str费用IDs <> ""
            If Len(str费用IDs) > 3998 Then
                lngTmp = InStrRev(Mid(str费用IDs, 1, 3998), ",")
                str费用ID = Mid(str费用IDs, 1, lngTmp - 1)
                str费用IDs = Mid(str费用IDs, lngTmp + 1)
            Else
                str费用ID = Mid(str费用IDs, 1, Len(str费用IDs) - 1)
                str费用IDs = ""
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_结帐费用记录_Batch('" & str费用ID & "'," & mrsInfo!病人ID & "," & lng结帐ID & ")"
        Wend
    End With
    
    '5.填写开始票据号
    If mblnPrint And Trim(txtInvoice.Text) <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_票据起始号_Update('" & strNo & "','" & Trim(txtInvoice.Text) & "',3)"
    End If
        
    '最后执行前并发操作判断
    '------------------------------------------------------------------------------
    '6.检查结帐操作期间,病人费用余额是否发生变化.
    If opt出院.Value Then
        If mcurSpare <> Get病人余额(mrsInfo!病人ID, 0, mint预交类别) Then
        '刘兴洪 问题:问题:34244    日期:2010-11-19 15:06:09
        Call MsgBox("病人要结帐的费用余额与实际的费用余额不一致!" & vbCrLf & _
        "可能是结帐过程中,输入了病人信息后,病区修改了病人费用!" & vbCrLf & _
        "点击『确定』后,系统将强制重新读取病人费用!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
            If mDateBegin = CDate("0:00:00") Then
                txtPatient_KeyPress (13)  '不会因txt中是名字而出现重名的问题,因为mrsInfo是打开的,不会重读病人信息
            Else
                Call ShowBalance
            End If
            Screen.MousePointer = 0
            Exit Function
        End If
    End If
    
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        blnTransMC = False
        If intInsure <> 0 Then
            If mbytMCMode = 1 Then  '门诊医保结算
                If cur个人帐户 <> 0 Or cur医保基金 <> 0 Or MCPAR.门诊必须传递明细 Then
                    If Not gclsInsure.ClinicSwap(lng结帐ID, cur个人帐户, cur医保基金, 0, 0, intInsure, strAdvance) Then
                        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                    Else
                        blnTransMC = True
                    End If
                End If
            Else                    '住院医保结算
                If Not gclsInsure.SettleSwap(lng结帐ID, intInsure, strAdvance) Then
                    gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                Else
                    blnTransMC = True
                End If
            End If
            blnMustCommit = True
        Else
            '一卡通结算
            If mblnOneCard And Not mobjICCard Is Nothing Then
                If curOneCard <> 0 Then
                    If Not mobjICCard.PaymentSwap(curOneCard, dblOneCardBalance, intCardType, Val("" & mrsOneCard!医院编码), strCardNo, strTransFlow, lng结帐ID, mrsInfo!病人ID) Then
                        gcnOracle.RollbackTrans
                        MsgBox "一卡通结算失败", vbInformation, gstrSysName
                        Screen.MousePointer = 0
                        Exit Function
                    Else
                        gstrSQL = "zl_一卡通结算_Update(" & lng结帐ID & ",'" & mrsOneCard!结算方式 & "','" & strCardNo & "','" & intCardType & "','" & strTransFlow & "'," & dblOneCardBalance & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    End If
                End If
                blnMustCommit = True
            End If
        End If
    
    bln医保结算校对 = False
    If intInsure <> 0 Then
        bln医保结算校对 = CheckYbBalance(str保险结算, strAdvance)
    End If
    
    If Not bln医保结算校对 And Not mblnThreeDepositAfter Then
        '执行三方结算交易(过程中提交)
        If Not ExecuteThreeSwapPayInterface(objCard, lng结帐ID, mCurBrushCard.dblMoney, blnTransMC) Then Screen.MousePointer = 0: Exit Function
    Else
       gcnOracle.CommitTrans
    End If
    blnTrans = False
    If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mbytMCMode = 1, 交易Enum.Busi_ClinicSwap, 交易Enum.Busi_SettleSwap), True, intInsure)
    Screen.MousePointer = 0
    
    '正式结算前后,结算方式和结算金额未发生变化时不校对
    If bln医保结算校对 Then
        cur缴款 = Val(txt缴款.Text)
        str住院次数 = ""
        If mbytFunc <> 0 Then
            str住院次数 = IIf(gbln仅用指定预交款 And mbln门诊转住院 = False, IIf(mstrTime = "", mstrAllTime, mstrTime), "")
        End If

        bln医保结算校对 = frmMedicareReckoning.ShowMe(Me, _
            lng结帐ID, mrsInfo!病人ID, opt中途.Value, cur结帐金额合计, strAdvance, str保险信息, _
            intInsure, mstrDec, gBytMoney, cur缴款, "" & mrsInfo!医保号, mbytMCMode, str住院次数, mint预交类别, mblnThreeDepositAfter, IIf(mblnShowThree, lblDelMoney.Caption, ""), mrsCardType, mobjPayCards, objCard, mstrPrivs, mbytFunc = 0)
        If Not bln医保结算校对 Then
            MsgBox "单据[" & strNo & "]进行医保结算校对失败,结帐金额可能不正确!" & _
                vbCrLf & vbCrLf & "将不打印票据,请到[保险结算管理]或[病人结帐管理]中重新校对后再打印!", vbInformation, gstrSysName
            mblnPrint = False
        End If
    End If
    '加入单据历史记录(所有类型单据)
    strTmp = strNo
    For i = 0 To cboNO.ListCount - 1
        strTmp = strTmp & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strTmp, ","))
        cboNO.AddItem Split(strTmp, ",")(i)
        If i = 9 Then Exit For '只显示10个
    Next
    
    If (mbytFunc = 1 And mbytInState = 0) And opt出院.Value Then
        '出院结帐,检查是否结清
        Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, , , 2)
        If Not rsTmp Is Nothing Then
            '结清,调整自动记帐标志
            If Nvl(rsTmp!费用余额, 0) = 0 Then
                str住院次数 = ""
                str住院次数 = IIf(mstrTime = "", mstrAllTime, mstrTime)
                If str住院次数 <> "" Then
                    strSql = "zl_病人自动记帐_Stop(" & mrsInfo!病人ID & ",'" & str住院次数 & "')"
                    zlDatabase.ExecuteProcedure strSql, Me.Caption
                End If
            End If
        End If
    End If
    
    Set mtySquareCard.rsSquare = Nothing
    SaveBalance = lng结帐ID
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mbytMCMode = 1, 交易Enum.Busi_ClinicSwap, 交易Enum.Busi_SettleSwap), False, intInsure)
    End If
    
    Screen.MousePointer = 0
    Call SaveErrLog
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Screen.MousePointer = 99
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function ExecuteSquareUpdate(ByVal rsSquare As ADODB.Recordset, ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '入参:rsSquare-刷卡结算数据
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-01-09 22:47:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strTemp As String
    
     With rsSquare
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            'Zl_病人卡结算记录_Insert
            strSql = "Zl_病人卡结算记录_Insert("
            '  接口编号_In   In 病人卡结算记录.接口编号%Type,
            strSql = strSql & "" & Val(Nvl(!接口编号)) & ","
            '  消费卡id_In   In 病人卡结算记录.消费卡id%Type,
            strSql = strSql & "" & IIf(Val(Nvl(!消费卡ID)) = 0, "NULL", Val(Nvl(!消费卡ID))) & ","
            '  结算方式_In   In 病人卡结算记录.结算方式%Type,
            strSql = strSql & "'" & Trim(Nvl(!结算方式)) & "',"
            '  结算金额_In   In 病人卡结算记录.结算金额%Type,
            strSql = strSql & "" & Val(Nvl(!结算金额)) & ","
            '  卡号_In       In 病人卡结算记录.卡号%Type,
            strSql = strSql & "'" & Trim(Nvl(!卡号)) & "',"
            '  交易流水号_In In 病人卡结算记录.交易流水号%Type,
            
            strSql = strSql & "'" & Trim(Nvl(!交易流水号)) & "',"
            '  交易时间_In   In 病人卡结算记录.交易时间%Type,
            strTemp = Format(!交易时间, "yyyy-mm-dd HH:MM:SS")
            If strTemp = "" Then
                strSql = strSql & "NULL,"
            Else
                strSql = strSql & "to_date('" & strTemp & "','yyyy-mm-dd hh24:mi:ss'),"
            End If
            '  备注_In       In 病人卡结算记录.备注%Type,
            strSql = strSql & "'" & Trim(Nvl(!备注)) & "',"
            '  结帐id_In     In Varchar2
            strSql = strSql & "'" & lng结帐ID & "')"
            
            zlDatabase.ExecuteProcedure strSql, Me.Caption
            .MoveNext
        Loop
     End With
    ExecuteSquareUpdate = True
End Function

Private Function zlSequareBlance(ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡结算
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsSquare As ADODB.Recordset
    If mbytInState <> 0 Then GoTo GoEnd:

    '刘兴洪:
    If Not mtySquareCard.blnExistsObjects Then GoTo GoEnd:
    If gobjSquare.objSquareCard Is Nothing Then GoTo GoEnd:
    If mtySquareCard.rsSquare Is Nothing Then GoTo GoEnd:
    If mtySquareCard.rsSquare.State <> 1 Then GoTo GoEnd:
    If mtySquareCard.rsSquare.RecordCount = 0 Then GoTo GoEnd:

    Set rsSquare = zlDatabase.CopyNewRec(mtySquareCard.rsSquare)
    If rsSquare Is Nothing Then GoTo GoEnd:
    If rsSquare.State <> 1 Then GoTo GoEnd:
    If ExecuteSquareUpdate(rsSquare, lng结帐ID) = False Then Exit Function

    '调用相应的结算接口
    '调用接口
    'Public Function zlSquareFee(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal str结帐ID_IN As String, ByVal rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: zlSquareFee (结算接口)
    '入参:frmMain:HIS传入 调用的主窗体
    '     intCallType : HIS传入 0-  门诊费用调用 1-  住院结帐调用
    '     str结帐ID_IN: HIS传入 本次结帐的结帐ID集
    '     rsSquare :  本次应刷卡的交易
    '出参:
    '返回:true:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:18:38
    '说明:
    '    1. 在"门诊收费"界面点"确定"时,调用本接口
    '    2. 在"住院结帐"界面点"确定"时,调用本接口
    '注:
    '  此接口由于是在HIS事务中 , 因此不能在此接口存在与用户交互的操作
    '---------------------------------------------------------------------------------------------------------------------------------------------
     If gobjSquare.objSquareCard.zlSquareFee(Me, mlngModul, mstrPrivs, lng结帐ID, mtySquareCard.rsSquare) = False Then
          Exit Function
     End If
GoEnd:
    zlSequareBlance = True
    Exit Function
End Function

Private Function LoadCardData() As Boolean
'功能：根据当前选择的病人费用项目卡片，读取并设置费用清单
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim strInfo As String, strPre As String
    Dim strMoney As String, strTmp As String, strTmpSql As String
    Dim arrTotal() As Currency
    Dim strCond As String, BytType As Byte '0-门诊;1-住院;2-门诊和住院
    Dim DateBegin As Date, DateEnd As Date
    Dim strTable As String
    
    On Error GoTo errH
    
    If mbytInState = 0 And mrsInfo.State = 0 Then Exit Function
    
    strPre = sta.Panels(2).Text
    sta.Panels(2).Text = "正在读取数据,请稍候 ……"
    Screen.MousePointer = 11
    mshQuery.Redraw = False
    Me.Refresh
    
    If mbytInState = 0 Then
        strCond = ""
        strCond = strCond & IIf(mstrTime = "", "", " And Instr([2],','||Nvl(A.主页ID,0)||',')>0")
        If mDateBegin <> CDate("0:00:00") Then
            strCond = strCond & " And " & IIf(gint费用时间 = 0, "A.登记时间", "A.发生时间") & " Between [3] And [4]"
            DateBegin = CDate(Format(mDateBegin, "yyyy-MM-dd 00:00:00"))
            DateEnd = CDate(Format(mDateEnd, "yyyy-MM-dd 23:59:59"))
        End If
        strCond = strCond & IIf(mstrUnit = "", "", " And Instr([5],','||A.开单部门ID||',')>0")
        strCond = strCond & IIf(mbytBaby = 0, "", IIf(mbytBaby = 1, " And Nvl(A.婴儿费,0)=0", " And A.婴儿费=[6]"))
        strCond = strCond & IIf(mstrItem = "", "", " And Instr([7],','''||A.收据费目||''',')>0")
        
        If mbytKind = 1 Then
            strCond = strCond & " And A.门诊标志=4"
        Else
            If InStr(mstrPrivs, ";住院费用结帐;") = 0 Or mbytMCMode = 1 Then strCond = strCond & " And A.门诊标志<>2"
            If InStr(mstrPrivs, ";门诊费用结帐;") = 0 Then strCond = strCond & " And A.门诊标志<>1"
            If mbytKind = 0 Then strCond = strCond & " And A.门诊标志<>4"
        End If
        
        BytType = Zl病人费用来源
    
        '不用记录状态,只取有未结金额的单据(未明细到序号,要显示部份退费行)
        If Not mnuFileZero.Checked Then
            strSql = _
            " Select NO,Mod(记录性质,10) as 记录性质, Nvl(Sum(实收金额),0) as 实收金额,Nvl(Sum(结帐金额),0) as 结帐金额" & _
            " From 住院费用记录 A" & _
            " Where 记录状态<>0 And 记帐费用=1 " & strCond & _
            "       And 病人ID=[1]" & _
            " Group by NO,Mod(记录性质,10) " & _
            " Having Nvl(Sum(实收金额),0)-Nvl(Sum(结帐金额),0)<>0"
            
            strSql = _
                " Select Mod(A.记录性质,10) as 记录性质,A.发生时间,A.登记时间,A.NO,A.收费类别,A.收费细目ID,A.收据费目,A.开单部门ID,A.计算单位," & _
                "        A.数次,Nvl(A.付数,1) as 付数,A.标准单价,Sum(A.实收金额) As 实收金额,Sum(A.结帐金额) As 结帐金额,A.操作员姓名,A.费用类型" & _
                " From 住院费用记录 A,(" & strSql & ") B" & _
                " Where A.NO=B.NO And Mod(A.记录性质,10)=B.记录性质" & _
                "       And A.记录状态<>0 And A.记帐费用=1 And Nvl(A.实收金额,0)<>Nvl(A.结帐金额,0)" & _
                "       And A.病人ID+0=[1] " & strCond & _
                " Having Nvl(Sum(A.实收金额),0)-Nvl(Sum(A.结帐金额),0)<>0" & _
                " Group by Mod(A.记录性质,10),A.发生时间,A.登记时间,A.NO,A.收费类别,Nvl(A.价格父号,A.序号),A.收费细目ID," & _
                "       A.收据费目,A.开单部门ID,A.计算单位,A.数次,Nvl(A.付数,1),A.标准单价,A.操作员姓名,A.费用类型 "
            
            If mblnDateMoved Then
                strSql = strSql & " Union All " & Replace(strSql, "住院费用记录", "H住院费用记录")
            End If
        Else
            strSql = _
                " Select Mod(A.记录性质,10) as 记录性质,A.发生时间,A.登记时间,A.NO,A.收费类别,A.收费细目ID,A.收据费目,A.开单部门ID,A.计算单位," & _
                "       A.数次,Nvl(A.付数,1) as 付数,A.标准单价,Sum(A.实收金额) As 实收金额,Sum(A.结帐金额) As 结帐金额,A.操作员姓名,A.费用类型" & _
                " From " & IIf(mblnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & "" & _
                " Where A.记录状态<>0 And A.记帐费用=1  And A.病人ID=[1] " & strCond & _
                "       And (Nvl(A.实收金额,0)<>Nvl(A.结帐金额,0) Or Nvl(A.实收金额,0)=0 And A.结帐ID is Null)" & _
                " Having Nvl(Sum(A.实收金额),0)-Nvl(Sum(A.结帐金额),0)<>0 Or Sum(Nvl(A.实收金额,0))=0 And Sum(A.结帐金额) is Null" & _
               "  Group by Mod(A.记录性质,10),A.发生时间,A.登记时间,A.NO,A.收费类别,Nvl(A.价格父号,A.序号),A.收费细目ID,A.收据费目,A.开单部门ID,A.计算单位,A.数次,Nvl(A.付数,1),A.标准单价,A.操作员姓名,A.费用类型 "
        End If
        
        Select Case BytType
        Case 0 '门诊
            strSql = Replace(Replace(strSql, "住院费用记录", "门诊费用记录"), " And Instr([2],','||Nvl(A.主页ID,0)||',')>0", "")
            If Not mnuFileZero.Checked Then
                strTmpSql = _
                " Select NO,Mod(记录性质,10) as 记录性质, Nvl(Sum(实收金额),0) as 实收金额,Nvl(Sum(结帐金额),0) as 结帐金额" & _
                " From 住院费用记录 A" & _
                " Where 记录状态<>0 And 记帐费用=1 And Mod(记录性质,10)=5 And 主页ID Is Null " & strCond & _
                "       And 病人ID=[1]" & _
                " Group by NO,Mod(记录性质,10) " & _
                " Having Nvl(Sum(实收金额),0)-Nvl(Sum(结帐金额),0)<>0"
                
                strTmpSql = _
                " Select Mod(A.记录性质,10) as 记录性质,A.发生时间,A.登记时间,A.NO,A.收费类别,A.收费细目ID,A.收据费目,A.开单部门ID,A.计算单位," & _
                "        A.数次,Nvl(A.付数,1) as 付数,A.标准单价,Sum(A.实收金额) As 实收金额,Sum(A.结帐金额) As 结帐金额,A.操作员姓名,A.费用类型" & _
                " From 住院费用记录 A,(" & strTmpSql & ") B" & _
                " Where A.NO=B.NO And Mod(A.记录性质,10)=B.记录性质" & _
                "       And A.记录状态<>0 And A.记帐费用=1 And Mod(A.记录性质,10)=5 And A.主页ID Is Null And Nvl(A.实收金额,0)<>Nvl(A.结帐金额,0)" & _
                "       And A.病人ID+0=[1] " & strCond & _
                " Having Nvl(Sum(A.实收金额),0)-Nvl(Sum(A.结帐金额),0)<>0" & _
                " Group by Mod(A.记录性质,10),A.发生时间,A.登记时间,A.NO,A.收费类别,Nvl(A.价格父号,A.序号),A.收费细目ID," & _
                "       A.收据费目,A.开单部门ID,A.计算单位,A.数次,Nvl(A.付数,1),A.标准单价,A.操作员姓名,A.费用类型 "
                If mblnDateMoved Then
                    strTmpSql = strTmpSql & " Union All " & Replace(strTmpSql, "住院费用记录", "H住院费用记录")
                End If
            Else
                strTmpSql = _
                " Select Mod(A.记录性质,10) as 记录性质,A.发生时间,A.登记时间,A.NO,A.收费类别,A.收费细目ID,A.收据费目,A.开单部门ID,A.计算单位," & _
                "       A.数次,Nvl(A.付数,1) as 付数,A.标准单价,Sum(A.实收金额) As 实收金额,Sum(A.结帐金额) As 结帐金额,A.操作员姓名,A.费用类型" & _
                " From " & IIf(mblnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & "" & _
                " Where A.记录状态<>0 And A.记帐费用=1 And  Mod(A.记录性质,10)=5 And A.主页ID Is Null And A.病人ID=[1] " & strCond & _
                "       And (Nvl(A.实收金额,0)<>Nvl(A.结帐金额,0) Or Nvl(A.实收金额,0)=0 And A.结帐ID is Null)" & _
                " Having Nvl(Sum(A.实收金额),0)-Nvl(Sum(A.结帐金额),0)<>0 Or Sum(Nvl(A.实收金额,0))=0 And Sum(A.结帐金额) is Null" & _
               "  Group by Mod(A.记录性质,10),A.发生时间,A.登记时间,A.NO,A.收费类别,Nvl(A.价格父号,A.序号),A.收费细目ID,A.收据费目,A.开单部门ID,A.计算单位,A.数次,Nvl(A.付数,1),A.标准单价,A.操作员姓名,A.费用类型 "
            End If
            strTmpSql = Replace(strTmpSql, " And Instr([2],','||Nvl(A.主页ID,0)||',')>0", "")
            strSql = strSql & " Union All " & strTmpSql
        Case 1 '住院
        Case Else
            '门诊和住院
             strSql = strSql & " Union All " & Replace(Replace(strSql, "住院费用记录", "门诊费用记录"), " And Instr([2],','||Nvl(A.主页ID,0)||',')>0", "")
        End Select
        strTable = "(" & strSql & ") "
        
            
        '未结费用清单
        Select Case tabCard.SelectedItem.Index
            Case 2 '明细清单
                strSql = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 科室,Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||A.数次||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Nvl(A.实收金额,0)-Nvl(A.结帐金额,0),'999999999" & mstrDec & "')) as 未结金额,A.操作员姓名 as 操作员" & _
                " FROM " & strTable & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID " & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Order by 发生日期,单据号,费目"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 3 '分项目明细
                strSql = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 开单科室,Nvl(D.名称,C.名称) as 项目,Nvl(C.规格,' ') 规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||A.数次||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Nvl(A.实收金额,0)-Nvl(A.结帐金额,0),'999999999" & mstrDec & "')) as 未结金额," & _
                "       Nvl(A.费用类型,C.费用类型) as 类型,A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间" & _
                " FROM " & strTable & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1)
                
               strSql = strSql & _
                " Union All" & _
                " SELECT NULL as 发生日期,NULL as 单据号,NULL as 开单科室," & _
                "       Nvl(D.名称,C.名称) as 项目,Nvl(C.规格,' ')||'ZZZZZ' as 规格,NULL,to_char(sum(Nvl(A.数次,1)*Nvl(A.付数,1)))||' '||A.计算单位 as 数量,NULL as 标准单价," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额," & _
                "       NULL as 类型,NULL as 操作员,NULL as 登记时间" & _
                " FROM " & strTable & " A,收费项目目录 C,收费项目别名 D" & _
                " Where A.收费细目ID=C.ID And A.收费细目ID=D.收费细目ID(+)" & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                "              And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Group by Nvl(D.名称,C.名称),C.规格,A.计算单位" & _
                " Order by 项目,规格,发生日期,单据号"
                
                strMoney = "4,4,1,1,1,1,1,7,7,7,1,1,1"
            Case 4 '分类明细
                strSql = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 科室,Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||A.数次||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Nvl(A.实收金额,0)-Nvl(A.结帐金额,0),'999999999" & mstrDec & "')) as 未结金额,A.操作员姓名 as 操作员 " & _
                " FROM " & strTable & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID " & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Union All" & _
                " SELECT NULL as 发生日期,NULL as 单据号,NULL as 科室,NULL as 项目,Null as 规格,A.收据费目||'ZZZZZ' as 费目," & _
                "        NULL as 数量,NULL as 标准单价," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额,NULL as 操作员" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by A.收据费目||'ZZZZZ'" & _
                " Order by 费目,发生日期,单据号"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 5 '分月清单
                strSql = _
                " SELECT B.期间,A.收据费目 as 费目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额" & _
                "        FROM " & strTable & " A,期间表 B,收费项目目录 C" & _
                " Where A.登记时间 Between Trunc(B.开始日期) and Trunc(B.终止日期)+1-1/24/60/60 " & _
                "       And A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by B.期间,A.收据费目" & _
                " Union All" & _
                " SELECT B.期间||'ZZZZZ',NULL as 费目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额" & _
                " FROM " & strTable & " A,期间表 B,收费项目目录 C" & _
                " Where A.登记时间 Between Trunc(B.开始日期) and Trunc(B.终止日期)+1-1/24/60/60 " & _
                "       And A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by B.期间||'ZZZZZ'" & _
                " Order by 期间,费目"
                strMoney = "4,4,7,7"
                
            Case 6 '分类清单
                strSql = _
                " SELECT A.收据费目 as 费目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by A.收据费目 Order by 费目"
                strMoney = "4,7,7"
            Case 7 '逐日费用
                strSql = _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号,A.收据费目 as 费用项目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额," & _
                "        A.操作员姓名 as 操作员,A.记录性质" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by A.记录性质,TO_Char(A.发生时间,'YYYY-MM-DD'),A.NO,A.收据费目,A.操作员姓名"
                strSql = strSql & _
                " Union All" & _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO||'ZZZZZ' as 单据号,NULL as 费用项目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额,NULL as 操作员,A.记录性质" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Having Nvl(Sum(A.实收金额), 0) - Nvl(Sum(A.结帐金额), 0) <> 0" & _
                " Group by A.记录性质,TO_Char(A.发生时间,'YYYY-MM-DD'),A.NO" & _
                " Union All" & _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD')||'ZZZZZ' as 发生日期,NULL as 单据号,NULL as 费用项目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额,NULL as 操作员,-1" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Having Nvl(Sum(A.实收金额), 0) - Nvl(Sum(A.结帐金额), 0) <> 0" & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD')" & _
                " Order by 发生日期,记录性质 desc,单据号,费用项目"
                
                strMoney = "4,4,4,7,7,1"
            Case 8 '逐日费目
                strSql = _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.收据费目 as 费用项目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD'),A.收据费目" & _
                " Union All" & _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD')||'ZZZZZ' as 发生日期,NULL as 费用项目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 未结金额" & _
                " FROM " & strTable & " A,收费项目目录 C" & _
                " Where A.收费细目ID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.收费类别,Nvl(C.类别,'无'))||''',')>0") & _
                " Having Nvl(Sum(A.实收金额), 0) - Nvl(Sum(A.结帐金额), 0) <> 0" & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD')" & _
                " Order by 发生日期,费用项目"
                strMoney = "4,4,7,7"
        End Select
                
        mshQuery.MergeCells = flexMergeFree
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(mrsInfo!病人ID), "," & mstrTime & ",", DateBegin, DateEnd, _
                    "," & mstrUnit & ",", mbytBaby - 1, "," & mstrItem & ",", "," & mstrClass & ",", "," & mstrChargeType & ",")
        If rsTmp.RecordCount > 0 Then
            Set mshQuery.DataSource = rsTmp
        Else
            Call BandRectoGrid(mshQuery, rsTmp)
        End If
        
        
        mshQuery.Tag = tabCard.SelectedItem.Index
        For i = 0 To mshQuery.Cols - 1
            mshQuery.MergeCol(i) = False
        Next
        
        '求合计(小计)
        Select Case tabCard.SelectedItem.Index
            Case 2, 4  '明细清单、分类明细
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 5) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 8))
                            If IsNumeric(mshQuery.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 9))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            strTmp = mshQuery.TextMatrix(i, 5)
                            For j = 0 To 7
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "小 计:" & Left(strTmp, Len(strTmp) - 5)
                            Next
                            For j = 8 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 7
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 8) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 9) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 3 '分项目明细
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 4) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 8))
                            If IsNumeric(mshQuery.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 9))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            strTmp = mshQuery.TextMatrix(i, 3)
                            For j = 0 To 5
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "小 计:" & strTmp
                            Next
                            mshQuery.TextMatrix(i, 7) = " " '单价列
                            For j = 8 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 7
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 8) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 9) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 5 '分月清单
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 2))
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 3))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            For j = 0 To 1
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "小计:" & mshQuery.TextMatrix(i - 1, 0)
                            Next
                            For j = 2 To mshQuery.Cols - 1
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 1
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 6 '分类清单
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If IsNumeric(mshQuery.TextMatrix(i, 1)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 1))
                        If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 2))
                        mshQuery.MergeRow(i) = False
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.Col = 0: mshQuery.CellAlignment = 4
                    mshQuery.TextMatrix(mshQuery.Row, 0) = "合 计"
                    mshQuery.TextMatrix(mshQuery.Row, 1) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 7 '逐日单据
                If rsTmp.RecordCount > 0 Then
                    mshQuery.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 1) Like "*ZZZZZ") And Not (mshQuery.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 3))
                            If IsNumeric(mshQuery.TextMatrix(i, 4)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 4))
                            mshQuery.MergeRow(i) = False
                        Else
                            If mshQuery.TextMatrix(i, 1) Like "*ZZZZZ" Then
                                mshQuery.Row = i
                                mshQuery.MergeRow(i) = True
                                For j = 1 To 2
                                    mshQuery.Col = j: mshQuery.CellAlignment = 4
                                    mshQuery.TextMatrix(i, j) = "小计:" & mshQuery.TextMatrix(i - 1, 1)
                                Next
                                For j = 3 To mshQuery.Cols - 2
                                    mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                                Next
                            Else
                                mshQuery.Row = i
                                mshQuery.MergeRow(i) = True
                                For j = 0 To 2
                                    mshQuery.Col = j: mshQuery.CellAlignment = 4
                                    mshQuery.TextMatrix(i, j) = "小计:" & mshQuery.TextMatrix(i - 1, 0)
                                Next
                                For j = 3 To mshQuery.Cols - 2
                                    mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                                Next
                            End If
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 2
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 4) = Format(arrTotal(1), " " & mstrDec)
                    
                    '删除只有一行单据的小计行
                    j = 0
                    For i = 1 To mshQuery.Rows - 1
                        If mshQuery.TextMatrix(i, 1) Like "*小计*" Then
                            If j = 1 Then mshQuery.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
            Case 8 '逐日费目
                If rsTmp.RecordCount > 0 Then
                    mshQuery.MergeCol(0) = True
                    ReDim arrTotal(1)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 2))
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 3))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.MergeRow(i) = True
                            mshQuery.Row = i
                            mshQuery.Col = 1: mshQuery.CellAlignment = 4
                            mshQuery.TextMatrix(i, 0) = "小计:" & mshQuery.TextMatrix(i - 1, 0)
                            mshQuery.TextMatrix(i, 1) = mshQuery.TextMatrix(i, 0)
                            For j = 2 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 1
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(1), " " & mstrDec)
                
                    '删除只有一行费用的小计行
                    j = 0
                    For i = 1 To mshQuery.Rows - 1
                        If mshQuery.TextMatrix(i, 1) Like "*小计*" Then
                            If j = 1 Then mshQuery.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
        End Select
    Else
        strSql = "Select 发生时间,登记时间,NO,收据费目,费用类型,付数,数次,计算单位,标准单价,结帐金额,操作员姓名,开单部门ID,收费细目ID,结帐ID From 住院费用记录  where 结帐ID= [1]  Union ALL " & _
                 "Select 发生时间,登记时间,NO,收据费目,费用类型,付数,数次,计算单位,标准单价,结帐金额,操作员姓名,开单部门ID,收费细目ID,结帐ID From 门诊费用记录  where 结帐ID= [1]"
        
        If mblnNOMoved Then
            strSql = Replace(Replace(strSql, "住院费用记录", "H住院费用记录"), "门诊费用记录", "H门诊费用记录")
        End If
        strSql = "(" & strSql & ")"
        
        '读取结帐单时,点结帐分类明细
        Select Case tabCard.SelectedItem.Index
            Case 2 '明细
                '发生日期,单据号,科室,项目,费目,数量,单价,标准金额,结帐金额,操作员
                strSql = _
                " Select To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       Nvl(B.名称,'未知') as 科室,Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||A.数次||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(A.标准单价,'99999" & gstrFeePrecisionFmt & "')) as 单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(A.结帐金额,'999999999" & mstrDec & "')) as 结帐金额,A.操作员姓名 as 操作员" & _
                " From " & strSql & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID = B.ID(+) And A.收费细目ID=C.ID" & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Order by 发生日期,单据号,费目"
                
                '清单格式控制
               strMoney = "4,4,1,1,1,4,1,7,7,7,1"
            Case 3 '分项目明细
                '发生日期,单据号,科室,项目,规格,费目,数量,单价,标准金额,结帐金额,类型,操作员
                strSql = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 开单科室,Nvl(D.名称,C.名称) as 项目,Nvl(C.规格,' ') as 规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||A.数次||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Nvl(A.结帐金额,0),'999999999" & mstrDec & "')) as 结帐金额," & _
                "       Nvl(A.费用类型,C.费用类型) as 类型,A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间" & _
                " FROM " & strSql & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID" & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Union All" & _
                " SELECT NULL as 发生日期,NULL as 单据号,NULL as 开单科室,Nvl(D.名称,C.名称) as 项目,Nvl(C.规格,' ')||'ZZZZZ' as 规格," & _
                "        NULL as 费目,to_char(sum(Nvl(A.数次,1)*Nvl(A.付数,1)))||' '||A.计算单位 as 数量,NULL as 标准单价," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额,NULL as 类型,NULL as 操作员,NULL as 登记时间" & _
                " FROM " & strSql & " A,收费项目目录 C,收费项目别名 D" & _
                " Where A.收费细目ID=C.ID " & _
                "       And A.收费细目ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Group by Nvl(D.名称,C.名称),C.规格,A.计算单位" & _
                " Order by 项目,规格,发生日期,单据号"
                strMoney = "4,4,1,1,1,4,1,7,7,7,1,1,1"
            Case 4 '分类明细
                strSql = _
                " SELECT To_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号," & _
                "       B.名称 as 科室,Nvl(D.名称,C.名称) as 项目,C.规格,A.收据费目 as 费目," & _
                "       Decode(Nvl(A.付数,1),1,'',0,'',A.付数||' 付 × ')||A.数次||' '||A.计算单位 as 数量," & _
                "       Ltrim(To_Char(Nvl(A.标准单价,0),'999999999" & gstrFeePrecisionFmt & "')) as 标准单价," & _
                "       Ltrim(To_Char(Round(A.标准单价*A.数次*Nvl(A.付数,1),5),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Nvl(A.结帐金额,0),'999999999" & mstrDec & "')) as 结帐金额,A.操作员姓名 as 操作员 " & _
                " FROM " & strSql & " A,部门表 B,收费项目目录 C,收费项目别名 D" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID" & _
                "       And A.收费细目ID=D.收费细目ID(+) And 码类(+)=1 And D.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                " Union All" & _
                " SELECT NULL as 发生日期,NULL as 单据号,NULL as 科室,NULL as 项目,Null as 规格,A.收据费目||'ZZZZZ' as 费目," & _
                "       NULL as 数量,NULL as 标准单价," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额,NULL as 操作员" & _
                " FROM " & strSql & " A,部门表 B,收费项目目录 C" & _
                " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID " & _
                " Group by A.收据费目||'ZZZZZ' " & _
                " Order by 费目,发生日期,单据号"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 5 '分月清单
                strSql = _
                " SELECT B.期间,A.收据费目 as 费目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额" & _
                " FROM " & strSql & " A,期间表 B" & _
                " Where A.登记时间 Between Trunc(B.开始日期) and Trunc(B.终止日期)+1-1/24/60/60 " & _
                " Group by B.期间,A.收据费目" & _
                " Union All" & _
                " SELECT B.期间||'ZZZZZ',NULL as 费目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额" & _
                " FROM " & strSql & " A,期间表 B" & _
                " Where A.登记时间 Between Trunc(B.开始日期) and Trunc(B.终止日期)+1-1/24/60/60 " & _
                " Group by B.期间||'ZZZZZ'" & _
                " Order by 期间,费目"
                strMoney = "4,4,7,7"
            Case 6 '分类清单
                strSql = _
                " SELECT A.收据费目 as 费目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额" & _
                " FROM " & strSql & " A" & _
                " Group by A.收据费目 Order by 费目"
                strMoney = "4,7,7"
            Case 7 '逐日单据
                strSql = _
                    " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO as 单据号,A.收据费目 as 费用项目," & _
                    "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额,A.操作员姓名 as 操作员 " & _
                    " FROM " & strSql & " A" & _
                    " Group by TO_Char(A.发生时间,'YYYY-MM-DD'),A.NO,A.收据费目,A.操作员姓名" & _
                    " Union All" & _
                    " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.NO||'ZZZZZ' as 单据号,NULL as 费用项目," & _
                    "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额, NULL as 操作员  " & _
                    " FROM " & strSql & " A" & _
                    " Group by TO_Char(A.发生时间,'YYYY-MM-DD'),A.NO" & vbCrLf & _
                    " Union All" & _
                    " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,'ZZZZZAAAAA' as 单据号,NULL as 费用项目," & _
                    "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额,NULL as 操作员 " & _
                    " FROM  " & strSql & " A" & _
                    " Group by TO_Char(A.发生时间,'YYYY-MM-DD')" & _
                    " Order by 发生日期,单据号,费用项目"
                strMoney = "4,4,4,7,7,1"
            Case 8 '逐日费目
                strSql = _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD') as 发生日期,A.收据费目 as 费用项目," & _
                "       Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "       Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额" & _
                " FROM " & strSql & " A " & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD'),A.收据费目" & _
                " Union All" & _
                " SELECT TO_Char(A.发生时间,'YYYY-MM-DD')||'ZZZZZ' as 发生日期,NULL as 费用项目," & _
                "        Ltrim(To_Char(Sum(Round(A.标准单价*A.数次*Nvl(A.付数,1),5)),'999999999" & mstrDec & "')) as 标准金额," & _
                "        Ltrim(To_Char(Sum(Nvl(A.结帐金额,0)),'999999999" & mstrDec & "')) as 结帐金额" & _
                " FROM " & strSql & " A" & _
                " Group by TO_Char(A.发生时间,'YYYY-MM-DD')" & _
                " Order by 发生日期,费用项目"
                strMoney = "4,4,7,7"
        End Select
        
        mshQuery.MergeCells = flexMergeFree
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngBillID)
        If rsTmp.RecordCount > 0 Then
            Set mshQuery.DataSource = rsTmp
        Else
            Call BandRectoGrid(mshQuery, rsTmp)
        End If

        mshQuery.Tag = tabCard.SelectedItem.Index
        For i = 0 To mshQuery.Cols - 1
            mshQuery.MergeCol(i) = False
        Next

        Select Case tabCard.SelectedItem.Index
            Case 2, 4  '明细清单、分类明细
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 5) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 8))
                            If IsNumeric(mshQuery.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 9))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            strTmp = mshQuery.TextMatrix(i, 5)
                            For j = 0 To 7
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "小 计:" & Left(strTmp, Len(strTmp) - 5)
                            Next
                            For j = 8 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 7
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 8) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 9) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 3 '分项目明细
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 4) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 8))
                            If IsNumeric(mshQuery.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 9))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            strTmp = mshQuery.TextMatrix(i, 3)
                            For j = 0 To 5
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "小 计:" & strTmp
                            Next
                            mshQuery.TextMatrix(i, 7) = " " '单价列
                            For j = 8 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 7
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 8) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 9) = Format(arrTotal(1), " " & mstrDec)
                End If
             Case 5 '分月清单
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 2))
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 3))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            For j = 0 To 1
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "小计:" & mshQuery.TextMatrix(i - 1, 0)
                            Next
                            For j = 2 To mshQuery.Cols - 1
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 1
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(1), " " & mstrDec)
                End If
             Case 6 '分类清单
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If IsNumeric(mshQuery.TextMatrix(i, 1)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 1))
                        If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 2))
                        mshQuery.MergeRow(i) = False
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.Col = 0: mshQuery.CellAlignment = 4
                    mshQuery.TextMatrix(mshQuery.Row, 0) = "合 计"
                    mshQuery.TextMatrix(mshQuery.Row, 1) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 7
                For i = 0 To mshQuery.Cols - 1
                    mshQuery.MergeCol(i) = False
                Next
                If rsTmp.RecordCount > 0 Then
                    mshQuery.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 1) Like "*ZZZZZ") And Not (mshQuery.TextMatrix(i, 1) Like "*AAAAA") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 3))
                            If IsNumeric(mshQuery.TextMatrix(i, 4)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 4))
                            mshQuery.MergeRow(i) = False
                        Else
                            If Not (mshQuery.TextMatrix(i, 1) Like "*AAAAA") Then
                                mshQuery.Row = i
                                mshQuery.MergeRow(i) = True
                                For j = 1 To 2
                                    mshQuery.Col = j: mshQuery.CellAlignment = 4
                                    mshQuery.TextMatrix(i, j) = "单据小计:" & mshQuery.TextMatrix(i - 1, 1)
                                Next
                                For j = 3 To mshQuery.Cols - 2
                                    mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                                Next
                            Else
                                mshQuery.Row = i
                                mshQuery.MergeRow(i) = True
                                For j = 1 To 2
                                    mshQuery.Col = j: mshQuery.CellAlignment = 4
                                    mshQuery.TextMatrix(i, j) = "日小计"
                                Next
                                For j = 3 To mshQuery.Cols - 2
                                    mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                                Next
                            End If
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 2
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 4) = Format(arrTotal(1), " " & mstrDec)
                    
                    '删除只有一行单据的小计行
                    j = 0
                    For i = 1 To mshQuery.Rows - 1
                        If mshQuery.TextMatrix(i, 1) Like "*小计*" Then
                            If j = 1 Then mshQuery.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
            Case 8
                For i = 0 To mshQuery.Cols - 1
                    mshQuery.MergeCol(i) = False
                Next
                If rsTmp.RecordCount > 0 Then
                    mshQuery.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not mshQuery.TextMatrix(i, 0) Like "*ZZZZZ" Then
                            If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 2))
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 3))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = False
                            mshQuery.Col = 0: mshQuery.CellAlignment = 4
                            mshQuery.TextMatrix(i, 0) = Left(mshQuery.TextMatrix(i, 0), Len(mshQuery.TextMatrix(i, 0)) - 5)
                            mshQuery.TextMatrix(i, 1) = "日小计"
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 1
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "合 计"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(1), " " & mstrDec)
                    
                    '删除只有一行单据的小计行
                    j = 0
                    For i = 1 To mshQuery.Rows - 1
                        If mshQuery.TextMatrix(i, 1) Like "*小计*" Then
                            If j = 1 Then mshQuery.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
        End Select
    End If
    
    '总的格式控制
    If mshQuery.Rows = 1 Then mshQuery.Rows = 2
    
    For i = 0 To mshQuery.Cols - 1
        mshQuery.FixedAlignment(i) = 4
    Next
    
    '如果取了,由于没有设置初始列宽,打印会异常
    Call SetGridWidth(mshQuery, Me)
    
    '有个记录性质列
    If tabCard.SelectedItem.Index = 7 And mbytInState = 0 Then
        mshQuery.ColWidth(mshQuery.Cols - 1) = 0
    End If
    
    For i = 0 To UBound(Split(strMoney, ","))
        mshQuery.ColAlignment(i) = Split(strMoney, ",")(i)
    Next
    
    mshQuery.Row = 1: mshQuery.Col = 0: mshQuery_EnterCell
    
    sta.Panels(2).Text = strPre
    
    mshQuery.Redraw = True
    mshQuery.Refresh
    Screen.MousePointer = 0
    LoadCardData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    mshQuery.Redraw = True
    If ErrCenter() = 1 Then
        mshQuery.Redraw = False
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
    sta.Panels(2).Text = strPre
End Function

Private Function GetMinMaxTime(ByVal bytMode As Byte) As Integer
'功能:取未结费用中的最小或最大的住院次数,可能返回0
'参数:bytMode,0-最小次数,1-最大次数
    Dim strTime As String, arrTmp As Variant
    Dim i As Long, intTime As Integer
    
    strTime = IIf(mstrTime = "", mstrAllTime, mstrTime)
    arrTmp = Split(strTime, ",")
    For i = 0 To UBound(arrTmp)
        If i = 0 Then intTime = Val(arrTmp(i))
        If bytMode = 0 Then
            If intTime > Val(arrTmp(i)) Then intTime = Val(arrTmp(i))
        Else
            If intTime < Val(arrTmp(i)) Then intTime = Val(arrTmp(i))
        End If
    Next
    
    GetMinMaxTime = intTime
End Function

Private Sub GetFeeDate(dBegin As Date, dEnd As Date)
'功能：获取病人的最小和最大费用时间
    Dim i As Long, dateThis As Date
    
    mrsBalance.MoveFirst
    For i = 1 To mrsBalance.RecordCount
        If gint费用时间 = 0 Then
            dateThis = mrsBalance!登记时间
        Else
            dateThis = mrsBalance!时间
        End If
        If i = 1 Then
            dBegin = dateThis
            dEnd = dateThis
        Else
            If dateThis < dBegin Then dBegin = dateThis
            If dateThis > dEnd Then dEnd = dateThis
        End If
        
        mrsBalance.MoveNext
    Next
    mrsBalance.MoveFirst
End Sub

Private Function GetPatiDate(dBegin As Date, dEnd As Date) As Boolean
'功能：获取病人的入出院时间,门诊病人取最大和最小费用时间
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lng主页ID As Long

    Call GetFeeDate(dBegin, dEnd)
    If mrsInfo!主页ID <> 0 Then
        lng主页ID = GetMinMaxTime(0)
        If lng主页ID > 0 Then
            If lng主页ID = mrsInfo!主页ID Then
                dBegin = mrsInfo!入院日期
                If IsDate(mstr本次住院日期) Then    '问题:30043
                    If Format(dBegin, "yyyy-mm-dd") < mstr本次住院日期 Then dBegin = CDate(mstr本次住院日期)
                End If
                If Not IsNull(mrsInfo!出院日期) Then
                    dEnd = mrsInfo!出院日期
                Else
                    dEnd = zlDatabase.Currentdate
                End If
            Else
                If CStr(lng主页ID) = IIf(mstrTime = "", mstrAllTime, mstrTime) Then '可能是结以前某次住院的帐
                    On Error GoTo errH
                    strSql = "Select 入院日期,Nvl(出院日期,Sysdate) as 出院日期 From 病案主页" & _
                            " Where 病人ID=[1] And 主页ID=[2]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(mrsInfo!病人ID), lng主页ID)
                    dBegin = rsTmp!入院日期
                    If IsDate(mstr本次住院日期) Then
                        If Format(dBegin, "yyyy-mm-dd") < mstr本次住院日期 Then dBegin = CDate(mstr本次住院日期)
                    End If
                    dEnd = rsTmp!出院日期
                End If
            End If
        End If
    End If
    
    GetPatiDate = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshDetail.Cols - 1
        If mshDetail.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub cmdYB_Click()
'功能：门诊病人结帐前的身份验证(成都医保还支持住院病人医保身份验证)
    Dim lng病人ID As Long, bytMode As Byte
    Dim strMessage As String, intInsure As Integer
    lng病人ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    Call NewBill
    bytMode = 0
    If mblnMC_TwoMode Then
        If InStr(mstrPrivs, ";门诊费用结帐;") = 0 Then
            bytMode = 4
        Else
            If zlCommFun.ShowMsgbox("医保身证验证", "请选择病人身份验证模式。", "!住院医保(&Z),门诊医保(&M)", Me, vbInformation) = "住院医保" Then
                bytMode = 4
            End If
        End If
    End If
        
    '刘兴洪:门诊转住院费用时加入
    mstrYBPati = gclsInsure.Identify(bytMode, lng病人ID, intInsure)
    If mstrYBPati = "" Then GoTo ExceptionHand
    cmdOK.Enabled = False   '问题:43776
    
    mbytMCMode = IIf(bytMode = 0, 1, 2) '必须在LoadPatientInfo之前
    If mbytMCMode = 1 Then
        '        'lng病人ID:49084
        If Not gclsInsure.GetCapability(support门诊结帐, lng病人ID, intInsure) Then
            strMessage = "病人当前险类不支持门诊医保结帐。": GoTo ExceptionHand
        End If
    End If
    
    'New:空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
    If lng病人ID <> 0 Then
        txtPatient.Text = "-" & lng病人ID
        Call LoadPatientInfo(IDKIND.GetCurCard, False, intInsure)
        If mrsInfo.State = 0 Then GoTo ExceptionHand
    Else
        strMessage = "病人身份验证成功,但未发现病人的帐户信息!" & vbCrLf & "可能是病人入院时没有进行验证,不能进行保险结算！"
        GoTo ExceptionHand
    End If
    Exit Sub
ExceptionHand:
    If strMessage <> "" Then Call MsgBox(strMessage, vbInformation, gstrSysName)
    Set mrsInfo = New ADODB.Recordset
    mstrYBPati = "": mbytMCMode = 0
    txtPatient.Text = "": txtPatient.SetFocus
    cmdOK.Enabled = True
End Sub

Private Sub HideMoneyInfo()
    lbl医保基金.Caption = "统筹支付:"
    lbl医保基金.Visible = False
    lbl个人帐户.Caption = "帐户余额:"
    lbl个人帐户.Visible = False
    Form_Resize
End Sub

Private Sub txtTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtTotal.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtTotal.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtTotal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtTotal.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Function GetPatiState(lng病人ID As Long) As String
'功能：返回病人状态说明
'普通在院,留观在院,医保在院;普通出院,留观出院,医保出院;门诊普通,门诊留观,门诊医保
    Dim lng主页ID As Long
    If mrsInfo!主页ID = 0 Or mbytMCMode = 1 Then
        If IsNull(mrsInfo!险类) Then
            GetPatiState = "门诊普通"
        Else
            GetPatiState = "门诊医保"
        End If
    Else
        If Nvl(mrsInfo!病人性质, 0) = 1 Then
            GetPatiState = "门诊留观"
        Else
            If Not IsNull(mrsInfo!险类) Then
                GetPatiState = "医保"
            ElseIf Nvl(mrsInfo!病人性质, 0) = 2 Then
                GetPatiState = "留观"
            Else
                GetPatiState = "普通"
            End If
            If mbytFunc = 0 Then
                If Is门诊留观(mrsInfo!病人ID, lng主页ID) Then
                     GetPatiState = "门诊留观"
                Else
                    GetPatiState = "门诊" & GetPatiState
                End If
            Else
                If IsNull(mrsInfo!出院日期) Then
                    GetPatiState = GetPatiState & "在院"
                Else
                    GetPatiState = GetPatiState & "出院"
                End If
            End If
        End If
        If Nvl(mrsInfo!状态, 0) = 3 Then
            GetPatiState = GetPatiState & "(预出院)"
        End If
    End If
End Function

Private Function Get应缴() As Currency
    Dim i As Long
    
    For i = 1 To vsfMoney.Rows - 1
        If Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("性质"))) = 1 Then
            Get应缴 = Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("金额")))
            Exit Function
        End If
    Next
End Function

Private Sub txt备注_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt备注
End Sub

Private Sub txt备注_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Get应缴 > 0 And txt缴款.Visible Then
        txt缴款.SetFocus
    ElseIf cmdOK.Visible And cmdOK.Enabled Then
        cmdOK.SetFocus
    End If
End Sub
Private Sub txt备注_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt备注, KeyAscii, m文本式
End Sub
Private Sub txt备注_LostFocus()
   zlCommFun.OpenIme False
End Sub

Private Sub txt缴款_Change()
    
    If Val(txt缴款.Text) = 0 Then txt找补.Text = "0.00"
    Call Calc找补
    
'    txt找补.Text = Format(Val(txt缴款.Text) - Get应缴, "0.00")
End Sub

Private Sub txt缴款_GotFocus()
    '#21 1234.56   --请您付款一千二百三十四点五六元  J
    '#22 1234.56   --预收一千二百三十四点五六元 Y
    '#23 1234.56   --找零一千二百三十四点五六元 Z
    Dim curTotal As Currency
    
    Call zlControl.TxtSelAll(txt缴款)
    If gblnLED Then
        zl9LedVoice.DisplayBank (" ")
        curTotal = Get应缴
        If curTotal >= 0 Then
            zl9LedVoice.Speak "#21 " & curTotal
        Else
            zl9LedVoice.Speak "#23 " & Abs(curTotal)
        End If
    End If
End Sub

Private Sub Led欢迎信息()
    'LED初始化
    If mbytInState = 0 And gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.编号 & "号 为您服务", mlngModul, gcnOracle
        End If
        
        zl9LedVoice.DisplayPatient txtPatient.Text & " " & txtSex.Text & " " & txtOld.Text, Val("" & mrsInfo!病人ID)
    End If
End Sub

Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If KeyAscii = Asc(".") And InStr(txt缴款.Text, ".") > 0 Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt缴款_LostFocus()
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub

Private Sub txt缴款_Validate(Cancel As Boolean)
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
    
    If Val(txt缴款.Text) <> 0 Then
        If CSng(txt找补.Tag) < 0 Then
            MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
            Call SelAll(txt缴款): txt缴款.SetFocus
            Cancel = True: Exit Sub
        End If
    End If
    
    If gblnLED Then
        zl9LedVoice.DispCharge Format(Get应缴, "0.00"), txt缴款.Text, txt找补.Text
        zl9LedVoice.Speak "#22 " & txt缴款.Text
        zl9LedVoice.Speak "#23 " & CSng(txt找补.Tag)
        zl9LedVoice.Speak "#3"                  '#3  --请当面点清, 谢谢!
    End If
End Sub

Private Sub txt天数_KeyPress(KeyAscii As Integer)
    '此控件获得焦点,仅为了使前一控件:结帐时间输完后,不跳到预交款输入处,避免输入错误导致预交款被退.
    If KeyAscii = vbKeyReturn Then Call SendKeys("{Tab}")
End Sub
Private Sub Calc找补()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算找补
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-01-12 17:41:47
    '问题:27360
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl找补 As Double
    If Val(txt缴款.Text) = 0 Then txt找补.Text = "0.00"
    dbl找补 = FormatEx(Val(txt缴款.Text) - Get应缴, 2)
    txt找补.Text = Format(Abs(dbl找补), "0.00")
    txt找补.Tag = dbl找补
    If dbl找补 <= 0 Then
        lbl找补.Caption = "收款"
        lbl找补.ForeColor = &H0&
    Else
        lbl找补.Caption = "找补"
        lbl找补.ForeColor = vbRed   '35830
    End If
    txt找补.ForeColor = lbl找补.ForeColor
End Sub
Private Sub txt找补_Change()
    txt找补.Tag = ""
End Sub

Private Function Get可刷金额() As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡的可刷金额
    '返回:
    '编制:刘兴洪
    '日期:2010-02-08 13:49:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, intCol As Integer
    Dim dbl可刷金额 As Double, dbl冲预交 As Double
    Dim dbl总额 As Double
    
    dbl总额 = GetBalanceSum
    dbl可刷金额 = 0
    For i = 1 To vsfMoney.Rows - 1
        If InStr(1, ";8;1;", ";" & vsfMoney.TextMatrix(i, vsfMoney.ColIndex("性质")) & ";") = 0 And Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("金额"))) <> 0 Then
            dbl可刷金额 = dbl可刷金额 + Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("金额")))
        End If
    Next
    
    dbl冲预交 = 0
    For i = 1 To mshDeposit.Rows - 1
        dbl冲预交 = dbl冲预交 + Val(mshDeposit.TextMatrix(i, mshDeposit.ColIndex("冲预交")))
    Next
            
    dbl可刷金额 = dbl总额 - dbl冲预交 - dbl可刷金额
    If dbl可刷金额 < 0 Then dbl可刷金额 = 0
    Get可刷金额 = Format(dbl可刷金额, gstrDec)
End Function

Private Function zlSquareCardFeeList(ByRef rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡明细信息
    '入参:
    '出参:rsFreeList-返回明细数据
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-05 16:02:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As ADODB.Recordset, strDate As String, strInvoice As String
    Dim i As Long
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsBalance Is Nothing Then Exit Function
    
    If zlCreateFeeListStruc(rsFeeList) = False Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    Set rsTemp = mrsBalance  'GetVBalance(mstrPrivs, mrsInfo!险类, mrsInfo!病人ID, mstrTime, mDateBegin, mDateEnd, False, mblnDateMoved, mbytBaby, mbytMCMode = 1, mbytKind, mstrItem, mstrUnit, mstrClass)
    rsTemp.Filter = 0
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
          rsFeeList.AddNew
          rsFeeList!单据序号 = 1
          rsFeeList!费别 = Nvl(rsTemp!费别)
          rsFeeList!NO = Nvl(rsTemp!单据号)
          rsFeeList!实际票号 = txtInvoice.Text
          rsFeeList!结算时间 = CDate(strDate)
          rsFeeList!病人ID = Val(Nvl(mrsInfo!病人ID))
          rsFeeList!主页ID = Val(Nvl(rsTemp!主页ID))
          rsFeeList!收费类别 = Nvl(rsTemp!收费类别)
          If Nvl(rsTemp!费目) <> "" Then
              rsFeeList!收据费目 = Nvl(rsTemp!费目)
          Else
              rsFeeList!收据费目 = Null
          End If
          rsFeeList!开单人 = Nvl(rsTemp!开单人)
          rsFeeList!收费细目ID = Val(Nvl(rsTemp!收费细目ID))
          rsFeeList!计算单位 = Nvl(rsTemp!计算单位)
          rsFeeList!数量 = Val(Nvl(rsTemp!数量))
          rsFeeList!单价 = Format(Val(Nvl(rsTemp!价格)), gstrFeePrecisionFmt)
          rsFeeList!实收金额 = Format(Val(Nvl(rsTemp!未结金额)), gstrDec)
          rsFeeList!统筹金额 = Format(Val(Nvl(rsTemp!统筹金额)), gstrDec)
          rsFeeList!保险支付大类ID = IIf(Val(Nvl(rsTemp!保险大类ID)) = 0, Null, Val(Nvl(rsTemp!保险大类ID)))
          rsFeeList!是否医保 = 0 ' Val(Nvl(rsTemp!是否医保))
          rsFeeList!保险编码 = Null ' Nvl(rsTemp!保险编码)
          rsFeeList!摘要 = Null ' Nvl(rsTemp!摘要)
          rsFeeList!是否急诊 = 0 ' Val(Nvl(rsTemp!是否急诊))
          rsFeeList!开单部门ID = Val(Nvl(rsTemp!开单部门ID))
          rsFeeList!执行部门ID = Val(Nvl(rsTemp!执行部门ID))
          rsFeeList!本次结算 = 0
          rsFeeList.Update
          rsTemp.MoveNext
    Loop
     If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    zlSquareCardFeeList = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function 住院刷结算卡() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:住院刷结算卡
     '返回:计算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-06 09:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String, cll结算合计 As Collection, strTemp As String, strNone As String
    Dim dblTemp As Double
    Dim arrPage As Variant, arrBalance() As String, strBalance As String
    Dim cur个帐合计 As Currency, cur个帐 As Currency, cur结算金额 As Currency, cur可分配额 As Currency
    Dim i As Integer, j As Integer, k As Integer, P As Integer
    Dim strDate As String, strAdvance As String, strInvoice As String, str结算方式 As String
                
    strInvoice = Trim(txtInvoice.Text)
    
    On Error GoTo errH
    strTemp = "": strNone = ""
    mtySquareCard.str刷卡结算 = ""
    Set cll结算合计 = New Collection
    '
    '结算方式;金额;是否允许修改|..."
    '先检查各种结算方式是否存在?
    ''"接口编号" "消费卡ID",  "卡号", "结算方式", "卡名称",   "余额",  "结算金额"  "交易时间",  "备注",  "结算标志"
    With mtySquareCard.rsSquare
        .Filter = 0: If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '必须已设置该结算方式,且为结算卡的结算方式
            str结算方式 = Nvl(!结算方式)
            mrs结算方式.Filter = "名称='" & str结算方式 & "' And 性质=8"
            If mrs结算方式.EOF Then
               If InStr(strNone & ",", "," & str结算方式 & ",") = 0 Then
                   strNone = strNone & "," & str结算方式
               End If
            End If
            If InStr(1, strTemp & ",", "," & str结算方式 & ",") > 0 Then
                dblTemp = Val(cll结算合计("K" & str结算方式)(0)) + Val(Nvl(!结算金额))
                cll结算合计.Remove "K" & str结算方式
            Else
                dblTemp = Val(Nvl(!结算金额))
            End If
            cll结算合计.Add Array(dblTemp, str结算方式), "K" & str结算方式
            strTemp = strTemp & "," & str结算方式
            .MoveNext
        Loop
    End With
    
    If strNone <> "" Then
        strNone = Mid(strNone, 2)
        MsgBox "当前结算卡的结算使用的结算方式" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
        "在结帐未设置，请先到结算方式管理中设置这些结算方式！", vbInformation, gstrSysName
        Exit Function
    End If
    
    str结算方式 = ""
    For i = 1 To cll结算合计.Count
        str结算方式 = cll结算合计(i)(1)
        If InStr(1, mtySquareCard.str刷卡结算, ";" & str结算方式 & ";") = 0 Then
            dblTemp = 0
            For j = 1 To cll结算合计.Count
                If str结算方式 = cll结算合计(j)(1) Then
                    dblTemp = dblTemp + Val(cll结算合计(i)(0))
                End If
            Next
            mtySquareCard.str刷卡结算 = ";" & str结算方式 & ";" & dblTemp & ";0|"
        End If
    Next
    If mtySquareCard.str刷卡结算 <> "" Then
        mtySquareCard.str刷卡结算 = Mid(mtySquareCard.str刷卡结算, 2)
        mtySquareCard.str刷卡结算 = Mid(mtySquareCard.str刷卡结算, 1, Len(mtySquareCard.str刷卡结算) - 1)
    End If
    ShowMoney True, , mty_ModulePara.bytMzDeposit
    Screen.MousePointer = 0
    住院刷结算卡 = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlReCalcRequare(ByRef cur结帐余额 As Currency, ByRef strNotBlance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置结帐卡部分金额
    '入参:
    '出参:cur结帐余额-返回当前计算后的结帐余额
    '     strNotBlance-返回未设置结算的信息
    '返回:计算成功能,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2010-02-08 14:27:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varBalace As Variant, i As Long, j As Long
    Dim varItem As Variant, strMoney As String
    
    If mtySquareCard.str刷卡结算 = "" Then zlReCalcRequare = True: Exit Function
    '结算方式;金额;是否允许修改|..."
    varBalace = Split(mtySquareCard.str刷卡结算, "|")
    With vsfMoney
        '设置结帐卡部分金额
        For i = 0 To UBound(varBalace)
            strMoney = varBalace(i) '结算方式;金额;是否允许修改|....
            varItem = Split(strMoney, ";")  '结算方式;金额;是否允许修改
            For j = 1 To .Rows - 1
                If .TextMatrix(j, .ColIndex("结算方式")) = CStr(varItem(0)) And InStr(",8,", Val(vsfMoney.TextMatrix(j, .ColIndex("性质")))) > 0 Then
                     .TextMatrix(j, .ColIndex("金额")) = Format(CCur(varItem(1)), "0.00")
                    If Val(varItem(2)) = 0 Then
                        vsfMoney.RowData(j) = 1 '该结算金额不可更改
                    Else
                        vsfMoney.RowData(j) = 0 '该结算金额可以更改
                    End If
                    '加入结算卡已处理的结算
                    cur结帐余额 = cur结帐余额 - Format(Val(vsfMoney.TextMatrix(j, .ColIndex("金额"))), "0.00")
                    Exit For
                End If
            Next
            '未包含医保可报销结算方式
            If j = vsfMoney.Rows Then
                mrs结算方式.Filter = "结算方式='" & varItem(0) & "'"
                If mrs结算方式.EOF Then
                    strNotBlance = strNotBlance & vbCrLf & vbTab & CStr(Split(strMoney, ";")(0)) & ":" & Format(CCur(Split(strMoney, ";")(1)), "0.00")
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(CCur(varItem(1)), "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = varItem(0)
                    .TextMatrix(.Rows - 1, .ColIndex("性质")) = Nvl(mrs结算方式!性质)
                    If Val(varItem(2)) = 0 Then
                        vsfMoney.RowData(.Rows - 1) = 1 '该结算金额不可更改
                    Else
                        vsfMoney.RowData(.Rows - 1) = 0 '该结算金额可以更改
                    End If
                    '加入结算卡已处理的结算
                    cur结帐余额 = cur结帐余额 - Format(Val(vsfMoney.TextMatrix(.Rows - 1, .ColIndex("金额"))), "0.00")
                End If
            End If
        Next
    End With
End Function


Private Function zlCallSquare_DelFree(ByVal str结帐ID_In As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行卡结算退费
    '入参:str结帐ID_In－原结帐ID
    '出参:
    '返回:如果调用成功,返回true,否则返回False,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-12 14:19:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Err = 0: On Error GoTo ErrHand:
    '该张单据不存在卡结算,退出
    If Not mtySquareCard.bln卡结算 Then zlCallSquare_DelFree = True: Exit Function

    'Zl_病人卡结算记录_Strike(结帐id_In In Varchar2)
    strSql = "Zl_病人卡结算记录_Strike(" & str结帐ID_In & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption

    'Public Function zlDelSquareFee(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal str结帐ID_IN As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能: zlSquareFee (结算接口)
    '    '入参:frmMain:HIS传入 调用的主窗体
    '    '     intCallType : HIS传入 0-  门诊费用调用 1-  住院结帐调用
    '    '     str结帐ID_IN: HIS传入 本次结帐的结帐ID集
    '    '出参:
    '    '返回:true:调用成功,False:调用失败
    '    '编制:刘兴洪
    '    '日期:2009-12-15 15:18:38
    '    '说明:
    '    '    1. "门诊收费管理"和"住院结帐管理"中作废时,调用此接口
    '    '注:
    '    '  此接口由于是在HIS事务中 , 因此不能在此接口存在与用户交互的操作
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlDelSquareFee(Me, mlngModul, mstrPrivs, str结帐ID_In) = False Then
        zlCallSquare_DelFree = False
        gcnOracle.RollbackTrans
    Else
        zlCallSquare_DelFree = True
    End If
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
End Function
Private Function zlIsCheckCanelFee(ByVal str结帐ID_In As String, ByVal bln部分退费 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费是否合法,合法，返回true,否则返回False
    '入参:str结帐ID_IN-结帐ID_IN
    '出参:
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-14 09:45:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    If mtySquareCard.bln卡结算 = False Then zlIsCheckCanelFee = True: Exit Function
    '是退费,则需要检查结算卡是否安装部件
    If gobjSquare.objSquareCard Is Nothing Then
        ShowMsgbox ("注意:" & vbCrLf & "    当前没有安装卡结算部件，不能进行退费,请检查！")
        Exit Function
    End If
    If bln部分退费 Then
        ShowMsgbox ("注意:" & vbCrLf & "    刷卡时的费用单，不能进行部分退费,请检查！")
        Exit Function
    End If
    If str结帐ID_In = "" Then
        ShowMsgbox ("注意:" & vbCrLf & "    未选择退费的单据，不能进行退费,请检查！")
        Exit Function
    End If

    'Public Function zlCheckDelSquareValied(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal str结帐ID_IN As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:在执行退费时,检查相关的接口部件是否正常
    '    '入参:
    '    '出参:
    '    '返回:正常,返回true,否则返回False
    '    '编制:刘兴洪
    '    '日期:2009-12-31 16:39:47
    '    '说明;
    '    '     在退费时，需要进行相关的检查
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlCheckDelSquareValied(Me, mlngModul, mstrPrivs, str结帐ID_In) = False Then
        Exit Function
    End If
    zlIsCheckCanelFee = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub zlClear结算卡()
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:清除结算卡的相关信息
        '编制:刘兴洪
        '日期:2010-01-11 11:26:20
        '---------------------------------------------------------------------------------------------------------------------------------------------
        Dim j As Long
        If cmd结算卡.Visible = False Then Exit Sub
        cmd结算卡.TabStop = True
        '需要重新刷卡处理
        Set mtySquareCard.rsSquare = Nothing
        mtySquareCard.str刷卡结算 = ""
        '需要清除表格中的刷卡金额部分
        With vsfMoney
            '设置结帐卡部分金额
            For j = 1 To .Rows - 1
                If InStr(",8,", Val(vsfMoney.TextMatrix(j, .ColIndex("性质")))) > 0 Then
                     .TextMatrix(j, .ColIndex("金额")) = "0.00"
                End If
            Next
        End With
    End Sub
Private Function IsCheck病历已接收(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病历是否已经接收
    '入参:
    '出参:
    '返回:已接收返回True,否则返回False
    '编制:刘兴洪
    '日期:2010-05-24 16:39:47
    '说明;
    '     问题:30036
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "select nvl(信息值,0) as 病历接收 from 病案主页从表 where 病人id=[1] and 主页id=[2] and 信息名='病历接收'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
            IsCheck病历已接收 = Val(Nvl(rsTemp!病历接收)) = 1
    Else
            IsCheck病历已接收 = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub zlSetDefaultTime(ByVal lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的住院日期
    '入参:lng病人ID-病人ID
    '       lng主页ID-主页ID
    '出参:
    '编制:刘兴洪
    '日期:2010-05-24 16:39:47
    '说明;
    '     问题:30043
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strDate As String
    
    strSql = "" & _
    "   Select to_char( Max(结束日期)+1,'yyyy-mm-dd') as 结束日期 " & _
    "   From 病人结帐记录 " & _
    "   Where  记录状态=1  And 病人iD=[1] and nvl(中途结帐,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID)
    If Not rsTemp.EOF Then
        strDate = Nvl(rsTemp!结束日期)
    Else
        strDate = ""
    End If
    mstr本次住院日期 = strDate
End Sub

Private Sub zlChangeDefaultTime()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：改变缺省日期
    '编制：刘兴洪
    '日期：2010-05-25 10:25:18
    '说明：30043
    '------------------------------------------------------------------------------------------------------------------------
    If opt出院.Value Then
        txtPatiEnd.Text = txtPatiEnd.Tag
    Else
        txtPatiEnd.Text = Format(zlDatabase.Currentdate - 1, "yyyy-mm-dd")
        If txtPatiEnd.Text < txtPatiBegin.Text Then
            txtPatiEnd.Text = txtPatiEnd.Tag
        End If
    End If
End Sub
Private Function zlGetPatiSource() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人来源(主要应用于是否存放位置)
    '返回:1-门诊;2-住院
    '编制:刘兴洪
    '日期:2011-03-14 18:01:36
    '问题号:36121
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str费用IDs As String, rsTemp As ADODB.Recordset
    Dim bln门诊 As Boolean, bln住院 As Boolean
    Dim strTable As String, strSql As String
    Dim BytType As Byte
    Dim i As Long
    
    Err = 0: On Error GoTo ErrHand:
    '0-权门诊;1-仅住院;2-门诊和住院
    BytType = Zl病人费用来源
    '误差费存放规则:
    '如果只结门诊的,放在门诊费用记录中;
    '如果包含了住院结帐的,则放在住院费用记录中;
    If BytType <> 2 Then
        '直接确定得了的,则返回
        zlGetPatiSource = IIf(BytType = 0, 1, 2): Exit Function
    End If
    '如果区分不出来的,则需要检查费用在那边的,
    '如果费用在住院(或即在门诊也在住院的),则误差放在住院;
    '如果费用仅在门诊的,则放在门诊费用
    With mshDetail
        For i = 1 To .Rows - 1
            If bln住院 Then
                zlGetPatiSource = 2: Exit Function
            End If
            If Val(.TextMatrix(i, COL_标志)) = 1 Then
                bln门诊 = True
            Else
                bln住院 = True
            End If
        Next
    End With
    If bln门诊 And bln住院 = False Then
        zlGetPatiSource = 1
    Else
        zlGetPatiSource = 2
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim intInsure As Integer
    intInsure = mintInsure
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng病人ID = Val(Nvl(mrsInfo!病人ID)): lng主页ID = Val(Nvl(mrsInfo!主页ID))
            intInsure = Val(Nvl(mrsInfo!险类))
        End If
    End If
    If mblnStartFactUseType Then mlng领用ID = 0
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng病人ID, lng主页ID, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType, IIf(mbytFunc = 1, "2", "1"))
    mintInvoiceMode = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    If blnFact Then Call RefreshFact
    Call ShowBillFormat
    mFactProperty = zl_GetInvoicePreperty(mlngModul, 2, IIf(mbytFunc = 1, "2", "1"))
End Sub

Private Sub InitPatiRedInvoice(Optional blnFact As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim intInsure As Integer
    intInsure = mintInsure
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng病人ID = Val(Nvl(mrsInfo!病人ID)): lng主页ID = Val(Nvl(mrsInfo!主页ID))
            intInsure = Val(Nvl(mrsInfo!险类))
        End If
    End If
    If mblnStartFactUseType Then mlng领用ID = 0
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng病人ID, lng主页ID, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoiceRedFormat(mlngModul, mstrUseType)
    mintInvoiceMode = zl_GetInvoiceRedMode(mlngModul, mstrUseType)
    If blnFact Then Call RefreshFact(True)
    Call ShowRedFormat
'    mFactProperty = zl_GetInvoicePreperty(mlngModul, 2, IIf(mbytFunc = 1, "2", "1"))
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng领用ID = GetInvoiceGroupID(1, intNum, lng领用ID, mlngShareUseID, strInvoiceNO, mstrUseType)
    If lng领用ID <= 0 Then
        Select Case lng领用ID
            Case 0 '操作失败
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的『" & mstrUseType & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据的『" & mstrUseType & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
                If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡对象的相关信息
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytInState = 1 Then Exit Sub
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then
        Call CreateSquareCardObject(Me, mlngModul)
        If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    End If
    
    Set mrsCardType = gobjSquare.objSquareCard.zlGetYLCards
    
    '所有启用的三方帐户
    Set mobjPayCards = gobjSquare.objSquareCard.zlGetCards(3)
    
    Call IDKIND.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKIND.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKIND.Cards.按缺省卡查找
    mtySquareCard.blnExistsObjects = isExistsThreeSwap
    
End Sub
Private Sub Init预交类别()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化预交类别
    '编制:刘兴洪
    '日期:2011-09-05 01:53:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int类别 As Integer, varPage As Variant
    Dim i As Integer
    mint预交类别 = IIf(mbytFunc = 0, 1, 2)
End Sub
Private Function isExistsThreeSwap() As Boolean
    Dim strPayType As String, varData As Variant, varTemp As Variant
    Dim i As Long, j As Long
    If gobjSquare Is Nothing Then Exit Function
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    varData = Split(strPayType, ";")
    For i = 0 To UBound(varData)
        If InStr(1, varData(i), "|") <> 0 Then
            varTemp = Split(varData(i), "|")
            If Val(varTemp(5)) = 1 Then
                '目前只针对消费卡
                isExistsThreeSwap = True: Exit Function
            End If
            j = j + 1
        End If
    Next
End Function
Private Sub WriteZYInforToCard(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, Optional blnDelete As Boolean = False)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将住院信息写入卡中
    '入参:blnDelete-是否退费
    '编制:刘兴洪
    '日期:2012-12-14 17:06:27
    '说明:
    '问题:56615
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strExpend As String
    '未确定刷卡类别,直接退出
    If mbytFunc = 0 Then
        If InStr(1, mstrPrivs, ";门诊信息写卡;") = 0 Then Exit Sub
    Else
        If InStr(1, mstrPrivs, ";住院信息写卡;") = 0 Then Exit Sub
    End If
    If lng病人ID = 0 Then Exit Sub
    If mlngCardTypeID = 0 Then
        If blnDelete Then GoTo goDelete:
        Exit Sub
    End If
    Dim objCard As Card
    If IDKIND.GetCurCard.接口序号 = mlngCardTypeID Then
        Set objCard = IDKIND.GetCurCard
    Else
        Set objCard = IDKIND.GetIDKindCard(mlngCardTypeID, CardTypeID)
    End If
    If objCard Is Nothing Then Exit Sub
    If objCard.是否写卡 = False Or objCard.接口序号 <= 0 Then Exit Sub '不准写卡的,不调用接口
    lngCardTypeID = objCard.接口序号
goDelete:
    If mbytFunc = 0 Then
        Call gobjSquare.objSquareCard.zlMzInforWriteToCard(Me, mlngModul, lngCardTypeID, _
        lng病人ID, lng结帐ID, strExpend)
    Else
        Call gobjSquare.objSquareCard.zlzyInforWriteToCard(Me, mlngModul, lngCardTypeID, _
        lng病人ID, lng结帐ID, strExpend)
    End If
End Sub

Private Function GetDelBalanceID(ByVal strNo As String, ByRef lng病人ID As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取作废的结帐ID
    '出参:lng病人ID-返回病人ID
    '返回:返回作废的结帐ID
    '编制:刘兴洪
    '日期:2012-12-14 18:52:31
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSql = "Select ID,病人ID From 病人结帐记录 Where  NO=[1] and 记录状态=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    If rsTemp.EOF Then Exit Function
    lng病人ID = Val(Nvl(rsTemp!病人ID))
    GetDelBalanceID = Val(Nvl(rsTemp!ID))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub setCardTypeRec()
    
End Sub

Private Function GetThreePayDepositData(ByRef rsTemp As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取三方交易信息
    '出参:rsTemp-返回交易信息(卡类别ID,卡类别名称,结算方式,是否启用,余额,冲预交,剩余款,退预交)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-04-27 09:44:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl冲预交 As Double, dblMoney As Double, dbl余额 As Double
    Dim dblTotal As Double, dblTemp As Double, lngCardTypeID As Long
    Dim varData As Variant
    
    On Error GoTo errHandle
    
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Fields.Append "卡类别ID", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "卡类别名称", adVarChar, 200, adFldIsNullable
    rsTemp.Fields.Append "结算方式", adVarChar, 100, adFldIsNullable
    rsTemp.Fields.Append "是否启用", adBigInt, , adFldIsNullable
    rsTemp.Fields.Append "余额", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "冲预交", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "剩余款", adDouble, , adFldIsNullable
    rsTemp.Fields.Append "退预交", adDouble, , adFldIsNullable
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    If mrsCardType Is Nothing Then
        Call initCardSquareData
    ElseIf mrsCardType.State <> 1 Then
        Call initCardSquareData
    End If
    dblTotal = Val(txtTotal.Text) - GetYBMoney
    With mshDeposit
        dblMoney = 0: dbl冲预交 = 0: dbl余额 = 0: lngCardTypeID = 0
        For i = 1 To .Rows - 1
            '卡类别ID||代扣||是否退现||缺省退现
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||||", "||")
            If Val(varData(0)) <> 0 And Val(varData(3)) = 0 Then
                lngCardTypeID = Val(varData(0))
                dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
                rsTemp.Find "卡类别ID=" & lngCardTypeID
                mrsCardType.Filter = "ID=" & lngCardTypeID
                If rsTemp.EOF Then
                    rsTemp.AddNew
                    rsTemp!卡类别ID = lngCardTypeID
                    If Not mrsCardType.EOF Then
                       rsTemp!卡类别名称 = mrsCardType!名称
                       rsTemp!是否启用 = Val(Nvl(mrsCardType!是否启用))
                    Else
                       rsTemp!卡类别名称 = .TextMatrix(i, .ColIndex("结算方式"))
                       rsTemp!是否启用 = 0
                    End If
                    rsTemp!结算方式 = .TextMatrix(i, .ColIndex("结算方式"))
                    rsTemp!退预交 = 0
                End If
                rsTemp!余额 = FormatEx(Val(Nvl(rsTemp!余额)) + Val(.TextMatrix(i, .ColIndex("余额"))), 5)
                rsTemp!冲预交 = FormatEx(Val(Nvl(rsTemp!冲预交)) + dbl冲预交, 5)
                rsTemp!剩余款 = FormatEx(Val(Nvl(rsTemp!余额)) - Val(Nvl(rsTemp!冲预交)), 5)
                'rsTemp!退预交 = Val(Nvl(rsTemp!剩余款))
                If FormatEx(dblTotal - dbl冲预交, 6) < 0 Then
                    If dblTotal >= 0 Then
                        dblTemp = dbl冲预交 - dblTotal
                        rsTemp!退预交 = FormatEx(Val(Nvl(rsTemp!退预交)) + dblTemp, 5)
                    Else
                        rsTemp!退预交 = FormatEx(Val(Nvl(rsTemp!退预交)) + dbl冲预交, 5)
                    End If
                    dblTotal = 0
                Else
                    dblTotal = FormatEx(dblTotal - dbl冲预交, 6)
                End If
                rsTemp.Update
            End If
        Next
    End With
    GetThreePayDepositData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowDelThreeSwap()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示退支付宝信息
    '编制:刘兴洪
    '日期:2015-04-27 11:09:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strTittle As String
    Dim intCount  As Integer, intNotCashCount  As Integer
    On Error GoTo errHandle
    
    Call GetDelThreeCardDepositInfor(mblnThreeDepositAll, intCount, intNotCashCount, mblnThreeDepositAfter, mstrStyle)
    
    
    lblDelMoney.Visible = False
    cmdReturnCash.Visible = False
    lblDelMoney.Tag = "0"
    
    If mstrForceNote <> "" Then
        mblnThreeDepositAfter = False
        GoTo BrushWin
    End If
    
     If mblnThreeDepositAfter Then
        lblDelMoney.Caption = IIf(mstrStyle = "", "", "退三方卡:" & mstrStyle)
        lblDelMoney.Visible = True
        mblnShowThree = True
        cmdReturnCash.Visible = lblDelMoney.Visible And lblDelMoney.Caption <> ""
        Call Form_Resize
        Exit Sub

     End If

    If GetThreePayDepositData(rsTemp) = False Then GoTo BrushWin

    '无记录时,表示不存在三方交易,直接返回true
    If rsTemp.RecordCount = 0 Then GoTo BrushWin
    rsTemp.Filter = "退预交<>0"
    If rsTemp.RecordCount = 0 Then GoTo BrushWin
    strTittle = ""
    Do While Not rsTemp.EOF
         strTittle = strTittle & IIf(strTittle = "", "", "  ") & "退" & Nvl(rsTemp!卡类别名称) & ":" & Format(Val(Nvl(rsTemp!退预交)), "0.00")
         lblDelMoney.Tag = FormatEx(Val(lblDelMoney.Tag) + Val(Nvl(rsTemp!退预交)), 6)
         rsTemp.MoveNext
    Loop
    lblDelMoney.Caption = strTittle
    lblDelMoney.Visible = True
    cmdReturnCash.Visible = lblDelMoney.Visible
BrushWin:
    Call Form_Resize
    Exit Sub
errHandle:
    Call Form_Resize
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function CheckThreePayDepositValied(ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查三方卡缴预交的合法性
    '出参:返回支付对象(三方卡)
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-04-23 17:32:37
    '规则:
    '     1)目前只支持三方帐户中存在(转帐交易接口的)
    '     2)不能同时存在2种以上的三方帐户交易的,存在的话返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strMsg As String
    Dim dblTotal As Double, dblMoney As Double
    Set objCard = Nothing
    If mblnThreeDepositAfter Or mstrForceNote <> "" Then
        CheckThreePayDepositValied = True
        Exit Function
    End If
    mCurBrushCard.dblMoney = 0
    If GetThreePayDepositData(rsTemp) = False Then Exit Function
    '无记录时,表示不存在三方交易,直接返回true
    If rsTemp.RecordCount = 0 Then CheckThreePayDepositValied = True: Exit Function
    rsTemp.Filter = "退预交<>0"
    If rsTemp.RecordCount = 0 Then CheckThreePayDepositValied = True: Exit Function
    
    If rsTemp.RecordCount >= 2 Then
        mblnThreeDepositAfter = True
        CheckThreePayDepositValied = True
        Exit Function
    End If
    If Val(Nvl(rsTemp!是否启用)) = 0 Then
       MsgBox Nvl(rsTemp!卡类别名称) & "未启用，不允许退款!" & _
              "", vbInformation + vbOKOnly, gstrSysName
       Exit Function
    End If
    If Not GetCurCard(Val(Nvl(rsTemp!卡类别ID)), objCard) Then
       MsgBox Nvl(rsTemp!卡类别名称) & "未启用或读取失败，不允许退款!", vbInformation + vbOKOnly, gstrSysName
       Exit Function
    End If

    
    dblMoney = FormatEx(Val(Nvl(rsTemp!退预交)), 6)
    mCurBrushCard.dblMoney = dblMoney
    
    If dblMoney <> FormatEx(Val(lblDelMoney.Tag), 6) Then
       If MsgBox(Nvl(rsTemp!卡类别名称) & "中界面显未金额(" & lblDelMoney.Tag & ")与当前退款金额(" & dblMoney & ") 不一致!" & vbCrLf & "，是否重新刷新界面的退款金额!", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
           Call ShowDelThreeSwap
       End If
       Exit Function
    End If
    
    If CheckTreeSwapBalaceIsValied(objCard, dblMoney) = False Then Exit Function
    CheckThreePayDepositValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCurCard(ByVal lngCardTypeID As Long, ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前卡对象
    '入参:lngCardTypeID-当前卡类别ID
    '出参:objCard-返回当前退款或缴款的卡对象
    '返回:成功,返回卡对象
    '编制:刘兴洪
    '日期:2015-04-27 10:32:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Card
    On Error GoTo errHandle
    Set objCard = Nothing
    For Each objTemp In mobjPayCards
        If objTemp.接口序号 = lngCardTypeID And Not objTemp.消费卡 Then
            Set objCard = objTemp
            GetCurCard = True: Exit Function
        End If
    Next
    GetCurCard = False
    Exit Function
errHandle:
End Function

Private Function CheckTreeSwapBalaceIsValied(ByVal objCard As Card, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡验证
    '入参:objCard-当前卡
    '     dblMoney-退卡金额
    '返回:刷卡成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-18 15:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String, dbl帐户余额 As Double
    Dim cllSquareBalance As New Collection
    On Error GoTo errHandle
    
    If objCard.接口序号 <= 0 Then CheckTreeSwapBalaceIsValied = True: Exit Function
     
    '   zlBrushCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal rsClassMoney As ADODB.Recordset, _
    ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByRef dbl金额 As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean = False, _
    Optional ByVal bln余额不足禁止 As Boolean = True, _
    Optional ByRef varSquareBalance As Variant) As Boolean
    Dim strCardNo As String, strPassWord As String, strXmlIn As String
    If objCard.是否转帐及代扣 Then
        
        strXmlIn = "<IN><CZLX>0</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
            objCard.接口序号, False, _
        txtPatient.Text, txtSex.Text, txtOld.Text, dblMoney, strCardNo, strPassWord, _
        False, True, False, False, cllSquareBalance, False, strXmlIn) = False Then Exit Function
        mCurBrushCard.str卡号 = strCardNo
        mCurBrushCard.str密码 = strPassWord
    End If
    '调用转帐接口
    '    7.1.    zltransferAccountsCheck(转帐检查接口)
    'zlTransferAccountsCheck 转帐检查接口
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  HIS调用模块号
    'lngCardTypeID   Long    In  卡类别ID
    'strCardNo   String  In  卡号
    'dblMoney    Double  In  转帐金额(代扣时为负数)
    'strBalanceIDs   String  In  结帐IDs，多个用逗号分离，表示本次对哪此收费项目进行重新医保补结算
    'strXMLExpend String In   XML串:
    '                            <IN>
    '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-退费业务；2-结帐业务;3-结帐退费业务
    '                            </IN>
    '                    Out  XML串:
    '                            <OUT>
    '                               <ERRMSG>错误信息</ERRMSG >
    '                            </OUT>
    '    Boolean 函数返回    检查的数据合法,返回True:否则返回False
    '说明:
    '１. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '构造XML串
    If objCard.是否转帐及代扣 Then
        zlXML.ClearXmlText
        zlXML.AppendNode "IN"
            zlXML.appendData "CZLX", "2"
        zlXML.AppendNode "IN", True
        strXMLExpend = zlXML.XmlText
        zlXML.ClearXmlText
        If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, objCard.接口序号, _
            mCurBrushCard.str卡号, dblMoney, "", strXMLExpend) = False Then
            Call ShowErrMsg(0, strXMLExpend)
            Exit Function
        End If
    End If
                    
'    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
'    ByVal strCardTypeID As Long, _
'    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '出参:dblMoney-返回帐户余额
    Dim strExpand As String
    Call gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.接口序号, _
          mCurBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡)
    mCurBrushCard.dbl帐户余额 = FormatEx(dbl帐户余额, 2)
    
    CheckTreeSwapBalaceIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ExecuteThreeSwapPayInterface(objCard As Card, ByVal lng结帐ID As Long, _
      ByVal dblMoney As Double, Optional ByVal blnMustCommit As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:objCard-当前三方对象
    '     lng结帐ID-结帐ID
    '     dblMoney-本次支付金额
    '     blnMustCommit-必须提交(主要是医保接口)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-04-27 10:45:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String, strXMLExpend As String
    Dim cllPro As Collection, blnTrans As Boolean, rsTmp As ADODB.Recordset
    Dim i As Long, strSql As String, lngID As Long, varData As Variant, strExpend As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection, strInXML As String, strOutXML As String
    Dim objXml As New clsXML, dblCheck As Double, dbl冲预交 As Double, lngRow As Long, strValue As String
    Dim strCardNo As String, strBalanceIDs As String, strExtra As String
    dblCheck = dblMoney
    blnTrans = True
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    If objCard Is Nothing Then
        gcnOracle.CommitTrans
        ExecuteThreeSwapPayInterface = True: Exit Function
    End If
        
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Then gcnOracle.CommitTrans: ExecuteThreeSwapPayInterface = True: Exit Function
    If objCard.是否转帐及代扣 Then
        'zlTransferAccountsMoney
        '参数名  参数类型    入/出   备注
        'frmMain Object  In  调用的主窗体
        'lngModule   Long    In  HIS调用模块号
        'lngCardTypeID   Long    In  卡类别ID
        'strCardNo   String  In  卡号
        'strBalanceID    String  In  结算ID
        'dblMoney    Double  In  转帐金额
        'strSwapGlideNO  String  Out 交易流水号
        'strSwapMemo String  Out 交易说明
        'strSwapExtendInfor  String  Out 交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
        'strXMLExpend String In   XML串:
        '                            <IN>
        '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-退费业务；2-结帐业务;3-结帐退费业务
        '                            </IN>
        '                    Out  XML串:
        '                            <OUT>
        '                               <ERRMSG>错误信息</ERRMSG >
        '                            </OUT>
        '    Boolean 函数返回    True:调用成功,False:调用失败
        '说明:
        '１. 在医保补充结算时进行的三方转帐时调用。
        '２. 一般来说，成功转帐后，都应该打印相关的结算票据，可以放在此接口进行处理.
        '３. 在转帐成功后，返回交易流水号和相关交易说明；如果存在其他交易信息，可以放在扩展信息中返回.
        '构造XML串
        zlXML.ClearXmlText
        zlXML.AppendNode "IN"
            zlXML.appendData "CZLX", "2"
        zlXML.AppendNode "IN", True
        strXMLExpend = zlXML.XmlText
        zlXML.ClearXmlText
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, objCard.接口序号, mCurBrushCard.str卡号, _
            lng结帐ID, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
            If Not blnMustCommit Then   '医保必须提交，必须根据病人预交记录中的校对标志来确定
                gcnOracle.RollbackTrans:
            Else
                gcnOracle.CommitTrans
                ExecuteThreeSwapPayInterface = True
            End If
            Call ShowErrMsg(1, strXMLExpend)
            blnTrans = False
            Exit Function
        End If
        
        mCurBrushCard.str交易流水号 = strSwapGlideNO
        mCurBrushCard.str交易说明 = strSwapMemo
        Call zlAddUpdateSwapSQL(False, lng结帐ID, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
        Call zlAddThreeSwapSQLToCollection(False, lng结帐ID, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapExtendInfor, cllThreeSwap)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        gcnOracle.CommitTrans
    Else
        objXml.ClearXmlText
        strBalanceIDs = ""
        With mshDeposit
            Call objXml.AppendNode("JSLIST")
            For i = .Rows - 1 To 1 Step -1
                '卡类别ID||代扣||是否退现||缺省退现
                varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||", "||")
                If Val(varData(0)) = objCard.接口序号 And dblCheck > 0 Then
                    dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
                    If dblCheck >= dbl冲预交 Then
                        lngID = .TextMatrix(i, .ColIndex("预交ID"))
                        strSql = "Select ID,卡号,交易流水号,交易说明 From 病人预交记录 Where ID = [1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
                        If Not rsTmp.EOF Then
                            Call objXml.AppendNode("JS")
                                Call objXml.appendData("KH", Nvl(rsTmp!卡号))
                                Call objXml.appendData("JYLSH", Nvl(rsTmp!交易流水号))
                                Call objXml.appendData("JYSM", Nvl(rsTmp!交易说明))
                                Call objXml.appendData("ZFJE", dbl冲预交)
                                Call objXml.appendData("JSLX", 1)
                                Call objXml.appendData("ID", Nvl(rsTmp!ID))
                            Call objXml.AppendNode("JS", True)
                            strSql = "Zl_三方退款信息_Insert("
                            strSql = strSql & lng结帐ID & ","
                            strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                            strSql = strSql & dbl冲预交 & ",'"
                            strSql = strSql & Nvl(rsTmp!卡号) & "','"
                            strSql = strSql & Nvl(rsTmp!交易流水号) & "','"
                            strSql = strSql & Nvl(rsTmp!交易说明) & "')"
                            zlAddArray cllThreeSwap, strSql
                        End If
                        dblCheck = dblCheck - dbl冲预交
                    Else
                        lngID = .TextMatrix(i, .ColIndex("预交ID"))
                        strSql = "Select ID,卡号,交易流水号,交易说明 From 病人预交记录 Where ID = [1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
                        If Not rsTmp.EOF Then
                            Call objXml.AppendNode("JS")
                                Call objXml.appendData("KH", Nvl(rsTmp!卡号))
                                Call objXml.appendData("JYLSH", Nvl(rsTmp!交易流水号))
                                Call objXml.appendData("JYSM", Nvl(rsTmp!交易说明))
                                Call objXml.appendData("ZFJE", dblCheck)
                                Call objXml.appendData("JSLX", 1)
                                Call objXml.appendData("ID", Nvl(rsTmp!ID))
                            Call objXml.AppendNode("JS", True)
                            strSql = "Zl_三方退款信息_Insert("
                            strSql = strSql & lng结帐ID & ","
                            strSql = strSql & Val(Nvl(rsTmp!ID)) & ","
                            strSql = strSql & dblCheck & ",'"
                            strSql = strSql & Nvl(rsTmp!卡号) & "','"
                            strSql = strSql & Nvl(rsTmp!交易流水号) & "','"
                            strSql = strSql & Nvl(rsTmp!交易说明) & "')"
                            zlAddArray cllThreeSwap, strSql
                        End If
                        dblCheck = 0
                    End If
                    strBalanceIDs = strBalanceIDs & "," & lngID
                End If
            Next i
            Call objXml.AppendNode("JSLIST", True)
        End With
    
        strInXML = objXml.XmlText
        strExtra = objXml.XmlText
        If strBalanceIDs <> "" Then strBalanceIDs = "1|" & Mid(strBalanceIDs, 2)
        
        If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, objCard.接口序号, objCard.消费卡, "", strBalanceIDs, _
         dblMoney, "", "", strExtra) = False Then gcnOracle.RollbackTrans: Exit Function
        
        If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModul, objCard.接口序号, objCard.消费卡, strInXML, _
             lng结帐ID, strOutXML, strExpend) = False Then
            If Not blnMustCommit Then   '医保必须提交，必须根据病人预交记录中的校对标志来确定
                gcnOracle.RollbackTrans:
            Else
                gcnOracle.CommitTrans
                ExecuteThreeSwapPayInterface = True
            End If
            Call ShowErrMsg(1, strXMLExpend)
            blnTrans = False
            Exit Function
        End If
             
        If strOutXML <> "" Then
            If zlXML_Init = False Then Exit Function
            If zlXML_LoadXMLToDOMDocument(strOutXML, False) = False Then Exit Function
            Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
            For i = 0 To lngRow - 1
                Call zlXML_GetNodeValue("ID", i, strValue)
                strSql = "Zl_三方退款信息_Insert("
                strSql = strSql & lng结帐ID & ","
                strSql = strSql & Val(strValue) & ","
                strSql = strSql & 0 & ",'"
                Call zlXML_GetNodeValue("KH", i, strValue)
                strSql = strSql & strValue & "','"
                Call zlXML_GetNodeValue("TKLSH", i, strValue)
                strSql = strSql & strValue & "','"
                Call zlXML_GetNodeValue("TKSM", i, strValue)
                strSql = strSql & strValue & "',"
                strSql = strSql & 1 & ")"
                zlAddArray cllThreeSwap, strSql
            Next i
        End If
        
        If strExpend <> "" Then
            strSwapExtendInfor = ""
            If zlXML_LoadXMLToDOMDocument(strExpend, False) = False Then Exit Function
            Call zlXML_GetChildRows("EXPENDS", "EXPEND", lngRow)
            For i = 0 To lngRow - 1
                Call zlXML_GetNodeValue("XMMC", i, strValue)
                strSwapExtendInfor = strSwapExtendInfor & "||" & strValue
                Call zlXML_GetNodeValue("XMNR", i, strValue)
                strSwapExtendInfor = strSwapExtendInfor & "|" & strValue
            Next i
        End If
        If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
        strSql = "Select 卡号 From 病人预交记录 Where 结帐ID= [1] And 卡类别ID= [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID, objCard.接口序号)
        If Not rsTmp.EOF Then
            strCardNo = Nvl(rsTmp!卡号)
        End If
        Call zlAddUpdateSwapSQL(False, lng结帐ID, objCard.接口序号, objCard.消费卡, strCardNo, "", "", cllUpdate, 0)
        Call zlAddThreeSwapSQLToCollection(False, lng结帐ID, objCard.接口序号, objCard.消费卡, strCardNo, strSwapExtendInfor, cllThreeSwap)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        gcnOracle.CommitTrans
    End If
    Err = 0: On Error GoTo ErrOtherHand:
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    ExecuteThreeSwapPayInterface = True
    Exit Function
ErrOtherHand:
    ExecuteThreeSwapPayInterface = True
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub ShowErrMsg(ByVal BytType As Byte, ByVal strXMLErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方转账检查与代扣业务出错提示
    '编制:冉俊明
    '时间:2014-12-2
    '参数:
    '   bytType:0-转账检查,1-转账交易
    '   strXMLErrMsg:格式如下
    '            <OUT>
    '               <ERRMSG>错误信息</ERRMSG >
    '            </OUT>
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
    '解析错误信息
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '提示错误信息
    If Trim(strValue) = "" Then
        If BytType = 0 Then
            strValue = vbCrLf & "交易检查失败！"
        Else
            strValue = vbCrLf & "交易失败！"
        End If
    End If
    MsgBox strValue, vbExclamation + vbOKOnly, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Function LoadBalanceDeposit(ByVal lngID As Long, ByRef curDeposit As Currency) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结帐的冲预交
    '入参:lngID-结帐ID
    '出参:curDeposit-返回冲预交金额
    '编制:刘兴洪
    '日期:2015-04-29 10:12:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long
    On Error GoTo errHandle
    Me.lblSpare.Visible = False
        
    Call InitDepositGridHead '初始化列头
    Set rsTemp = GetBalanceDeposit(lngID, mblnNOMoved)
    curDeposit = 0
    With mshDeposit
        .Redraw = flexRDNone
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(rsTemp!ID))
            .Cell(flexcpData, i, .ColIndex("ID")) = Nvl(rsTemp!卡类别ID) & "||" & Nvl(rsTemp!代扣)
            .TextMatrix(i, .ColIndex("单据号")) = Nvl(rsTemp!单据号)
            .TextMatrix(i, .ColIndex("票据号")) = "" & rsTemp!票据号
            .TextMatrix(i, .ColIndex("日期")) = Format(rsTemp!日期, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("结算方式")) = Nvl(rsTemp!结算方式)
            .TextMatrix(i, .ColIndex("金额")) = Format(rsTemp!金额, "0.00")
            rsTemp.MoveNext
        Next
        .Row = 1: .Col = .Cols - 1
        .Redraw = flexRDBuffered
        curDeposit = 0
        For i = 1 To .Rows - 1
            curDeposit = curDeposit + Val(.TextMatrix(i, .ColIndex("金额")))
        Next
    End With
    
    lblDeposit.Caption = "冲预交:" & Format(curDeposit, "0.00")
    lblDeposit.Tag = curDeposit
    lblTicketCount.Caption = "预交款收据:" & rsTemp.RecordCount & "张"
    LoadBalanceDeposit = True
    Exit Function
errHandle:
    mshDeposit.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetThreeDelDepositMoney(ByVal lngCardTypeID As Long, _
    ByVal strNotCardTypeIDs As String, ByRef dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取三方交易的冲预交(退费部分)
    '入参:lngCardTypeID-卡类别Id(=0时,代表所有预交款)
    '     strNotCardTypeIDs-不包含的卡类别ID(lngCardTypeID=0时有效)
    '出参:dblMoney-返回冲预交金额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-04-29 10:42:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, varData As Variant
    On Error GoTo errHandle
    dblMoney = 0
    With mshDeposit
        For i = 1 To .Rows - 1
            '卡类别ID||代扣||是否退现||缺省退现
           varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "|||", "|")
           If Val(varData(0)) = lngCardTypeID And lngCardTypeID <> 0 Then
                dblMoney = FormatEx(dblMoney + Val(.TextMatrix(i, .ColIndex("金额"))), 6)
           ElseIf lngCardTypeID = 0 Then
                If InStr(1, "," & strNotCardTypeIDs & ",", "," & varData(0) & ",") = 0 Then
                    dblMoney = FormatEx(dblMoney + Val(.TextMatrix(i, .ColIndex("金额"))), 6)
                End If
           End If
        Next
    End With
    GetThreeDelDepositMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitBalanceGrid(ByRef vsGrid As VSFlexGrid)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算列头信息
    '编制:刘兴洪
    '日期:2015-05-04 17:33:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsGrid
        .Redraw = flexRDNone
        .Clear
        .Rows = 2: .Cols = 6: i = 0
        .TextMatrix(0, i) = "结算方式": i = i + 1
        .TextMatrix(0, i) = "金额": i = i + 1
        .TextMatrix(0, i) = "结算号码": i = i + 1
        .TextMatrix(0, i) = "性质": i = i + 1
        .TextMatrix(0, i) = "缺省": i = i + 1
        .TextMatrix(0, i) = "卡类别ID": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "金额" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            If .ColKey(i) = "性质" Or .ColKey(i) = "缺省" Or .ColKey(i) = "卡类别ID" Then
                .ColHidden(i) = True: .ColWidth(0) = 0
            End If
        Next
        .ColWidth(.ColIndex("结算方式")) = 1200
        .ColWidth(.ColIndex("金额")) = 1100
        .ColWidth(.ColIndex("结算号码")) = 1450
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function LoadBalanceInfor(ByVal lng结帐ID As Long, ByVal dblDepositMoney As Double, _
    ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结算信息
    '入参:lng结帐ID-结帐ID
    '     dblDepositMoney-预交金额
    '返回:加载成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-04 17:32:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngDefault As Long, lngCashRow As Long, lng误差Row As Long
    Dim dblBalance As Double, dblYbMoney As Double, dblTemp As Double, dblMoney As Double, dbl误差 As Double
    Dim blnThree As Boolean, blnCentMoney As Boolean, blnUndo As Boolean, blnFind As Boolean
    Dim strBalance As String
    Dim curMoney As Currency
    Dim rsTmp As ADODB.Recordset
    Dim blnVisible As Boolean
    
    On Error GoTo errHandle
    
    '结帐补款清单,未用的结算方式也列出,以便作废时,不允许的医保结算退现金
    '---------------------------------------------------------------------------------------------------------------------
    mrs结算方式.Filter = ""
    Call InitBalanceGrid(vsfMoney)
    Call InitBalanceGrid(vsDelBalance) '清除结算信息
    j = 1
 
    With vsfMoney
        .Redraw = False
        .Rows = mrs结算方式.RecordCount + 1
        lngCashRow = -1
        If mrs结算方式.RecordCount <> 0 Then mrs结算方式.MoveFirst
        For i = 1 To mrs结算方式.RecordCount
            .TextMatrix(i, .ColIndex("结算方式")) = mrs结算方式!名称
            .TextMatrix(i, .ColIndex("性质")) = mrs结算方式!性质
            .TextMatrix(i, .ColIndex("缺省")) = mrs结算方式!缺省
            mrs结算方式.MoveNext
        Next
        .Redraw = True
        
        '结算清单
        Me.lblSpare.Visible = False
        Set rsTmp = GetBalancePay(lng结帐ID, mblnNOMoved)
        If mbytInState <> 1 And rsTmp.RecordCount <> 0 And InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
            rsTmp.Filter = "性质 <> 3 And 性质 <> 4"
            If rsTmp.RecordCount <> 0 Then
                MsgBox "授予<仅用预交结帐>权限时,不能操作包含非预交结算的单据!", vbInformation, gstrSysName
                If mblnOutUse Then
                    mblnUnload = True
                Else
                    Call NewBill
                End If
                Exit Function
            End If
            rsTmp.Filter = ""
        End If
        blnFind = False
        For i = 1 To rsTmp.RecordCount
            blnFind = False
            For j = 1 To .Rows - 1
                If rsTmp!结算方式 = .TextMatrix(j, .ColIndex("结算方式")) Then
                    If Val(Nvl(rsTmp!性质)) = 9 Then
                        .TextMatrix(j, .ColIndex("金额")) = FormatEx(rsTmp!金额, 5)
                    Else
                        .TextMatrix(j, .ColIndex("金额")) = Format(rsTmp!金额, "0.00")
                    End If
                    .TextMatrix(j, .ColIndex("结算号码")) = "" & rsTmp!结算号码
                    .TextMatrix(j, .ColIndex("卡类别ID")) = Val(Nvl(rsTmp!卡类别ID))
                     blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                .Rows = .Rows + 1: j = .Rows - 1
                .TextMatrix(j, .ColIndex("结算方式")) = Nvl(rsTmp!结算方式)
                .TextMatrix(j, .ColIndex("金额")) = Format(rsTmp!金额, "0.00")
                .TextMatrix(j, .ColIndex("结算号码")) = "" & rsTmp!结算号码
                .TextMatrix(j, .ColIndex("卡类别ID")) = Val(Nvl(rsTmp!卡类别ID))
                If Val(Nvl(rsTmp!卡类别ID)) <> 0 Then blnThree = True
            End If
            rsTmp.MoveNext
        Next
        
        lngDefault = 0: lngCashRow = 0: lng误差Row = 0
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("性质"))) = 9 Then
                .Row = i: .Col = 0
                .Cell(flexcpForeColor, i, 0, i, 0) = vbRed
                lng误差Row = i
            End If
            If Nvl(.TextMatrix(i, .ColIndex("缺省"))) = 1 Then lngDefault = i: Exit For
            If Nvl(.TextMatrix(i, .ColIndex("性质"))) = 1 Then lngCashRow = i: Exit For
        Next
        
        If mbytInState = 0 And InStr(1, mstrPrivs, ";仅用预交结帐;") = 0 Then
            If mintInsure <> 0 Then
                If lngDefault = 0 And lngCashRow = 0 Then MsgBox "没有设置缺省结算方式,结帐场合也没有现金结算方式可用,无法进行医保结帐作废!", vbInformation, gstrSysName: Exit Function
            ElseIf blnThree Then
                If lngDefault = 0 And lngCashRow = 0 Then MsgBox "没有设置缺省结算方式,结帐场合也没有现金结算方式可用,无法返回预交款帐户!", vbInformation, gstrSysName: Exit Function
            End If
        End If
        '仅医保结帐作废时,将不支持回退的医保结算移到缺省结算方式上
        mbln医保作废全退 = True
        If mbytInState = 0 And mintInsure <> 0 Then        '
            .Row = lngDefault: .Col = 0
            .CellFontBold = True
            j = 1
            '医保不支持作废的结算方式退为缺省结算
            For i = 1 To .Rows - 1
                If InStr(",3,4,", "," & .TextMatrix(i, .ColIndex("性质")) & ",") > 0 _
                    And Val(.TextMatrix(i, .ColIndex("金额"))) <> 0 Then
                    '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                    If mbytMCMode = 1 And Not MCPAR.门诊病人结算作废 Then
                        blnUndo = Val(.TextMatrix(i, .ColIndex("性质"))) = 3
                    Else
                       'lng病人ID:49084
                        blnUndo = Not gclsInsure.GetCapability(IIf(mbytMCMode = 1, support门诊结算作废, support住院结算作废), lng病人ID, mintInsure, .TextMatrix(i, .ColIndex("结算方式")))
                        If blnUndo Then
                            blnUndo = Val(.TextMatrix(i, .ColIndex("性质"))) = 3
                        End If
                    End If
                     
                    If blnUndo Then
                        .TextMatrix(IIf(lngDefault <= 0, lngCashRow, lngDefault), .ColIndex("金额")) = Format(Val(.TextMatrix(IIf(lngDefault <= 0, lngCashRow, lngDefault), .ColIndex("金额"))) + Val(.TextMatrix(i, .ColIndex("金额"))), "0.00")
                        .TextMatrix(i, .ColIndex("金额")) = ""
                        mbln医保作废全退 = False
                    Else
                        .Row = i: .Col = 0: .CellBackColor = txtMoney.BackColor
                        .Col = 1: .CellBackColor = txtMoney.BackColor
                        .Col = 2: .CellBackColor = txtMoney.BackColor
                        vsDelBalance.TextMatrix(j, vsDelBalance.ColIndex("结算方式")) = .TextMatrix(i, .ColIndex("结算方式"))
                        vsDelBalance.TextMatrix(j, vsDelBalance.ColIndex("性质")) = .TextMatrix(i, .ColIndex("性质"))
                        vsDelBalance.TextMatrix(j, vsDelBalance.ColIndex("金额")) = Format(Val(.TextMatrix(i, .ColIndex("金额"))), "0.00")
                        vsDelBalance.TextMatrix(j, vsDelBalance.ColIndex("结算号码")) = .TextMatrix(i, .ColIndex("结算号码"))
                        vsDelBalance.TextMatrix(j, vsDelBalance.ColIndex("卡类别ID")) = .TextMatrix(i, .ColIndex("卡类别ID"))
                        vsDelBalance.Rows = vsDelBalance.Rows + 1
                        j = j + 1
                    End If
                End If
            Next
            
            If Not mbln医保作废全退 Then
                If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
                    MsgBox "医保不支持结帐作废,无法在仅用预交结帐时进行结帐作废!", vbInformation, gstrSysName: Exit Function
                End If
                '如果是现金,进行分币处理
                If Val(.TextMatrix(IIf(lngDefault <= 0, lngCashRow, lngDefault), .ColIndex("性质"))) = 1 _
                    And Val(.TextMatrix(IIf(lngDefault <= 0, lngCashRow, lngDefault), .ColIndex("金额"))) <> 0 _
                    And MCPAR.分币处理 Then
                    .TextMatrix(IIf(lngDefault <= 0, lngCashRow, lngDefault), .ColIndex("金额")) = Format(CentMoney(Val(.TextMatrix(IIf(lngDefault <= 0, lngCashRow, lngDefault), .ColIndex("金额")))), "0.00")
                End If
                For i = 1 To .Rows - 1
                    curMoney = curMoney + Val(.TextMatrix(i, .ColIndex("金额")))
                Next
            End If
        End If
    End With
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("金额"))) = 0 Then
                .RowHidden(i) = True
            Else
                .RowHidden(i) = False
            End If
            If InStr(",3,4,", "," & .TextMatrix(i, .ColIndex("性质")) & ",") > 0 Then
                dblYbMoney = dblYbMoney + Val(.TextMatrix(i, .ColIndex("金额")))
            Else
                dblBalance = dblBalance + Val(.TextMatrix(i, .ColIndex("金额")))
            End If
        Next i
        .Refresh
    End With
    
    blnVisible = blnThree And mbytInState = 0 And mbln合约单位 = False
    
    picThreeBalance.Visible = blnVisible
    mshDeposit.Visible = Not blnVisible
    vsfMoney.Visible = Not blnVisible
    txtOwe.Visible = Not blnVisible
    lblOwe.Visible = Not blnVisible
    
    Call picThreeBalance_Resize
    lblDepositMoney.Tag = "": picDelDeposit.Tag = ""
     
    mcur误差金额 = 0
    If mbytInState = 1 Then LoadBalanceInfor = True: Exit Function
    
    
    If Not blnThree Then
        If mintInsure <> 0 And Not mbln医保作废全退 Then
            '误差处理
            mcur误差金额 = dblDepositMoney + curMoney - Val(txtTotal.Text)
            vsfMoney.ToolTipText = "结帐作废,误差金额:" & FormatEx(mcur误差金额, 6)
        End If
        LoadBalanceInfor = True: Exit Function
    End If
    
    dblTemp = FormatEx(dblDepositMoney + dblBalance, 6) '- dblYbMoney
    dblMoney = dblTemp
    
    blnCentMoney = False '不需要分币处理
    If lngCashRow <> 0 Then
        blnCentMoney = True '需要分币处理
        If mintInsure <> 0 Then
            blnCentMoney = MCPAR.分币处理
        Else
            If gBytMoney = 0 Then blnCentMoney = False
        End If
    End If
    
    i = lngCashRow
    If i = 0 Then i = lngDefault
    If i <> 0 Then
        With vsfMoney
            strBalance = .TextMatrix(i, .ColIndex("结算方式"))
        End With
    End If
    If blnCentMoney Then
        dblMoney = CentMoney(CCur(dblMoney))
    Else
        dblMoney = FormatEx(dblMoney, 2)
    End If
    lblDepositMoney.Caption = Format(dblMoney, "0.00")
    lblDepositMoney.Tag = dblMoney
    picDelDeposit.Tag = strBalance
    
    mcur误差金额 = FormatEx(dblMoney - dblTemp, 6)
    lbl误差.Visible = mcur误差金额 <> 0
    lbl误差.Caption = IIf(mcur误差金额 <> 0, "误差:" & FormatEx(mcur误差金额, 6), "")
    lbl误差.Tag = mcur误差金额
    
    LoadBalanceInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function GetDepositSaveSql(ByVal lng原结帐ID As Long, ByVal lng冲销ID As Long, _
    ByVal lng病人ID As Long, ByVal dblMoney As Double, _
    ByVal strDate As String, ByRef cllPro As Collection, ByRef strNo As String, ByRef lng预交ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存相关的预交数据
    '入参:lng结帐ID-结帐ID
    '     dblMoney-本次充值金额
    '     strDate-充值日期(格式:yyyy-mm-dd hh24:mi:ss")
    '出参:cllPro-将SQL传入集合中
    '     strNO-预交单据号
    '     lng预交ID-预交ID
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-05 17:14:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int预交类型 As Integer
    Dim lng主页ID As Long, strSql As String
    Dim lng领用ID As Long, strFact As String
    
    '1-门诊;2-住院;0或NULL:10.29.0以前的结帐数据
    int预交类型 = 1
    If mint结帐类型 = 2 Or mint结帐类型 = 0 Then
        strSql = "" & _
        "   Select Nvl(A.主页ID,B.主页ID) as 主页ID From (Select  病人ID, Max(主页ID) as 主页ID From 病人预交记录  where 结帐ID=[1] Group by  病人ID) A" & _
        "        ,病人信息 B  " & _
        "   Where A.病人ID=B.病人ID(+) "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng原结帐ID)
        lng主页ID = Val(Nvl(rsTemp!主页ID))
        int预交类型 = 2
    End If
    strNo = zlDatabase.GetNextNo(11)
    mstrDepositNo = strNo
    lng预交ID = zlDatabase.GetNextId("病人预交记录")
    
    mbln打印预交收据 = False
    If mty_ModulePara.int预交票据 <> 0 And InStr(1, mstrPrivs, ";预交款收据打印;") > 0 Then
        '0-不打印,1-提示打印,2-不提示打印;'刘兴洪 问题:27776 日期:2010-02-04 16:49:03
        If mty_ModulePara.int预交票据 = 2 Then
           If MsgBox("你是否要打印“病人预交款收据”？" & vbCrLf & _
                   "   『是』：打印病人预交款收据" & vbCrLf & _
                   "   『否』：不打印病人预交款收据", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                mbln打印预交收据 = True
            End If
        Else
            mbln打印预交收据 = True
        End If
    End If
    
    If mbln打印预交收据 Then
        '严格控制
        If Mid(zlDatabase.GetPara(24, glngSys, , "00000"), 2, 1) = "1" Then
            lng领用ID = CheckUsedBill(2, IIf(lng领用ID > 0, lng领用ID, mFactProperty.lngShareUseID), "", mFactProperty.strUseType)
            If lng领用ID <= 0 Then
                Select Case lng领用ID
                    Case 0 '操作失败
                    Case -1
                        MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    Case -2
                        MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End Select
                lng领用ID = 0
                mbln打印预交收据 = False
            End If
            '严格：取下一个号码
            strFact = GetNextBill(lng领用ID)
        Else
            '松散：取下一个号码
            strFact = IncStr(UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngModul, "")))
        End If
    End If
    
    
    'Zl_病人预交记录_Insert
    strSql = "Zl_病人预交记录_Insert("
    '  Id_In         病人预交记录.ID%Type,
    strSql = strSql & "" & lng预交ID & ","
    '  单据号_In     病人预交记录.NO%Type,
    strSql = strSql & "'" & strNo & "',"
    '  票据号_In     票据使用明细.号码%Type,
    strSql = strSql & "'" & strFact & "',"
    '  病人id_In     病人预交记录.病人id%Type,
    strSql = strSql & "" & lng病人ID & ","
    '  主页id_In     病人预交记录.主页id%Type,
    strSql = strSql & "" & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
    '  科室id_In     病人预交记录.科室id%Type,
    strSql = strSql & "NULL,"
    '  金额_In       病人预交记录.金额%Type,
    strSql = strSql & "" & dblMoney & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSql = strSql & "'" & picDelDeposit.Tag & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSql = strSql & "NULL,"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    strSql = strSql & "NULL,"
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    strSql = strSql & "NULL,"
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    strSql = strSql & "NULL,"
    '  摘要_In       病人预交记录.摘要%Type,
    strSql = strSql & "'结帐作废返回预交款,结帐单号:" & cboNO.Text & "',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSql = strSql & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSql = strSql & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strSql = strSql & lng领用ID & ","
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strSql = strSql & "" & int预交类型 & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSql = strSql & "NULL,"
   '  结算卡序号_in 病人预交记录.结算卡序号%type:=NULL,
    strSql = strSql & "NULL,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSql = strSql & "NULL,"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSql = strSql & "NULL,"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSql = strSql & "NULL,"
    '  合作单位_In   病人预交记录.合作单位%Type := Null,
    strSql = strSql & "NULL,"
    '  收款时间_In   病人预交记录.收款时间%Type := Null
    strSql = strSql & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '   操作类型_In Integer:=0 :0-正常缴预交;1-存为划价单
    strSql = strSql & "0 ,"
    '   结算性质_In  病人预交记录.结算性质%Type := Null
    strSql = strSql & "12 )"
    
    zlAddArray cllPro, strSql
    GetDepositSaveSql = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckYbBalance(ByVal str虚拟结算 As String, ByVal str正式结算 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查虚拟结算和正式结算是否需要校对
    '入参:
    '出参:
    '返回:需要效对,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-05-06 11:57:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln医保结算校对 As Boolean, varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    Dim i As Long, j As Long
    
    If str正式结算 = "" Then Exit Function
    '医保结算校对
    If str虚拟结算 = str正式结算 Then Exit Function
    
    
    bln医保结算校对 = True
    varData = Split(str虚拟结算, "||")
    varData1 = Split(str正式结算, "||")
    If UBound(varData) <> UBound(varData1) Then CheckYbBalance = True: Exit Function
    
    bln医保结算校对 = False
    For i = 0 To UBound(varData)
        bln医保结算校对 = True
        varTemp = Split(varData(i), "|")
        For j = 0 To UBound(varData1)
            varTemp1 = Split(varData1(j), "|")
            If varTemp(0) = varTemp1(0) Then
                If Val(varTemp(1)) = Val(varTemp1(1)) Then
                    bln医保结算校对 = False
                End If
            End If
        Next
        If bln医保结算校对 Then Exit For
    Next
    CheckYbBalance = bln医保结算校对: Exit Function
    
End Function

Private Function GetYBMoney() As Currency
   '获取医保预结算金额
   Dim curMoney As Currency, j As Long
   curMoney = 0
    With vsfMoney
        For j = 1 To vsfMoney.Rows - 1
            If InStr(",3,4,", Val(.TextMatrix(j, .ColIndex("性质")))) > 0 Then
                curMoney = curMoney + Format(Val(.TextMatrix(j, .ColIndex("金额"))), "0.00")
            End If
        Next
    End With
    GetYBMoney = curMoney
End Function

Private Function CheckThreeSwapDelValied(ByVal lng原结帐ID As Long, _
    ByRef rsBalance As ADODB.Recordset, ByRef dblDelMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方交易验证退费检查
    '出参:rsBalance-当前三方结算信息
    '     dblDelMoney-退款金额
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 18:00:34
    '说明:同步验证了接口和刷卡接口的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTemp As Double, cllSquareBalance As Collection
    Dim strXMLExpend As String, bln退现 As Boolean
    Dim dbl帐户余额 As Double, dbl未付金额 As Double
    Dim strExpand As String, strSql As String
    Dim strBalanceIDs As String
    Dim intMousePointer As Integer
    Dim blnCurInput As Boolean, dblMoney As Double
    Dim tyBrushCard As TY_BrushCard
    Dim lng卡类别ID As Long, bln消费卡 As Boolean
    Dim i As Long, strXmlIn As String
    
    intMousePointer = Screen.MousePointer
    
    If Not mbln合约单位 Then CheckThreeSwapDelValied = True: Exit Function
    
    strSql = "" & _
    "   Select a.卡类别ID,A.结算方式,A.卡号,A.交易流水号,A.交易说明, A.冲预交, B.是否退款验卡 As 退款验卡 " & vbNewLine & _
    "   From 病人预交记录 A,医疗卡类别 B" & vbNewLine & _
    "Where A.结帐id = [1] and A.卡类别ID=B.ID and nvl(A.卡类别ID,0)<>0 and  mod(记录性质,10)<>1 and nvl(A.冲预交,0)<>0 "
    
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng原结帐ID)
    If rsBalance.RecordCount = 0 Then CheckThreeSwapDelValied = True: Exit Function
    
    If rsBalance.RecordCount > 1 Then
        MsgBox "该单据同时存在多种三方卡结算，当前版本不支持多种退费!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng卡类别ID = Val(Nvl(rsBalance!卡类别ID))
    
    dblTemp = 0
    With rsBalance
        Do While Not .EOF
            dblTemp = dblTemp + Val(Nvl(!冲预交))
            .MoveNext
        Loop
        rsBalance.MoveFirst
        dblTemp = FormatEx(dblTemp, 5)
    End With
    dblMoney = dblTemp
    dblDelMoney = dblMoney
    dblTemp = 0
    For i = 1 To vsfMoney.Rows - 1
        If Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("卡类别ID"))) = lng卡类别ID Then
            dblTemp = dblTemp + Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("金额")))
        End If
    Next
    If dblDelMoney > FormatEx(dblTemp, 5) Then
        MsgBox rsBalance!结算方式 & " 的退款金额超过了原始结算金额!" & vbCrLf & "原结算金额:" & Format(dblTemp, "0.00") & vbCrLf & "现退款金额:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:帐户回退交易前的检查
        '入参:frmMain-调用的主窗体
        '       lngModule-调用的模块号
        '       lngCardTypeID-卡类别ID
        '       strCardNo-卡号
        '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        '       dblMoney-退款金额
        '       strSwapNo-交易流水号(退款时检查)
        '       strSwapMemo-交易说明(退款时传入)
        '       strXMLExpend    XML IN  可选参数:异常单据重新退费(1)
        '返回:退款合法,返回true,否则返回Flase
        
    strXMLExpend = ""
    tyBrushCard.str卡号 = Nvl(rsBalance!卡号)
    tyBrushCard.str交易流水号 = Nvl(rsBalance!交易流水号)
    tyBrushCard.str交易说明 = Nvl(rsBalance!交易说明)

    strBalanceIDs = "2|" & lng原结帐ID
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, lng卡类别ID, bln消费卡, tyBrushCard.str卡号, _
        strBalanceIDs, dblMoney, tyBrushCard.str交易流水号, tyBrushCard.str交易说明, strXMLExpend) = False Then Exit Function
    CheckThreeSwapDelValied = True: Exit Function
    
    If Val(Nvl(rsBalance!退款验证)) = 1 Then
       '弹出刷卡界面
        'zlBrushCard(frmMain As Object, _
        'ByVal lngModule As Long, _
        'ByVal rsClassMoney As ADODB.Recordset, _
        'ByVal lngCardTypeID As Long, _
        'ByVal bln消费卡 As Boolean, _
        'ByVal strPatiName As String, ByVal strSex As String, _
        'ByVal strOld As String, ByVal dbl金额 As Double, _
        'Optional ByRef strCardNo As String, _
        'Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean) As Boolean
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, lng卡类别ID, _
            bln消费卡, txtPatient.Text, txtSex.Text, _
            txtOld.Text, dblMoney, tyBrushCard.str卡号, tyBrushCard.str密码, _
            True, True, bln退现, , , , strXmlIn) = False Then Exit Function
    End If
    CheckThreeSwapDelValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapDelInterface(ByVal lng原结帐ID As Long, ByVal lng冲销ID As Long, _
    ByVal rsBalance As ADODB.Recordset, _
    ByVal dblDelMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:rsBalance-当前需要退费记录集
    '     dblDelMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim i As Long, dblTemp As Double
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, dblMoney As Double, str结算方式  As String
    Dim lng卡类别ID As Long
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    If rsBalance Is Nothing Or Not mbln合约单位 Or dblDelMoney = 0 Then ExecuteThreeSwapDelInterface = True: Exit Function
    If rsBalance.State <> 1 Then ExecuteThreeSwapDelInterface = True: Exit Function
    If rsBalance.RecordCount = 0 Then ExecuteThreeSwapDelInterface = True: Exit Function
    
    rsBalance.MoveFirst
    lng卡类别ID = Val(Nvl(rsBalance!卡类别ID))
    dblTemp = 0
    With rsBalance
        Do While Not .EOF
            dblTemp = dblTemp + Val(Nvl(!冲预交))
            .MoveNext
        Loop
        rsBalance.MoveFirst
        dblTemp = FormatEx(dblTemp, 5)
    End With
    
    
    
    Err = 0: On Error GoTo ErrHand:
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    If FormatEx(dblTemp, 6) = 0 Then Exit Function
    If dblDelMoney > dblTemp Then
        MsgBox rsBalance!结算方式 & " 的退款金额超过了原始结算金额!" & vbCrLf & _
            "原结算金额:" & Format(dblTemp, "0.00") & vbCrLf & _
            "现退款金额:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    With rsBalance
        strCardNo = Nvl(!卡号)
        strSwapNO = Nvl(!交易流水号)
        strSwapMemo = Nvl(!交易说明)
        str结算方式 = Nvl(!结算方式)
    End With
    
    On Error GoTo ErrRoll:
      
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(扣款时的交易流水号)
    '       strSwapMemo-交易说明(扣款时的交易说明)
    '       strSwapExtendInfor-交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, lng卡类别ID, False, strCardNo, _
        "2|" & lng原结帐ID, dblDelMoney, strSwapNO, strSwapMemo, strSwapExtendInfor) = False Then gcnOracle.RollbackTrans: Exit Function
    'Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, strCardNO, strSwapNO, strSwapMemo, cllUpdate, 2)
    Call zlAddThreeSwapSQLToCollection(False, lng冲销ID, lng卡类别ID, False, strCardNo, strSwapExtendInfor, cllThreeSwap)
    
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    Err = 0: On Error GoTo GoEnd:
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption, True, True
GoEnd:
    ExecuteThreeSwapDelInterface = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function GetValiedTimes() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取实际的未结费用住院次数
    '返回:成功，返回有效的实际未结费用次数(主页ID,...)
    '编制:刘兴洪
    '日期:2017-10-18 17:03:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, strTemp As String
    Dim i As Long, lng主页ID As Long
    
    If mstr结帐主页IDs = "" Then GetValiedTimes = mstrAllTime: Exit Function
    
    varData = Split(mstr结帐主页IDs & ",", ",")
    strTemp = ""
    For i = 0 To UBound(varData)
        lng主页ID = Val(varData(i))
        If lng主页ID <> 0 Then
           If InStr("," & mstrAllTime & ",", "," & lng主页ID & ",") > 0 Then strTemp = strTemp & "," & lng主页ID
        End If
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    GetValiedTimes = IIf(strTemp = "", mstrAllTime, strTemp)
End Function


Private Function GetDelThreeCardDepositInfor(ByVal blnThreeDepositAll As Boolean, ByRef intThreeCount As Integer, ByRef intNotDelCashCount As Integer, _
    ByRef blnThreeDepositAfter As Boolean, ByRef strDelThreeNames As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取退三帐户的预交相关信息
    '入参:blnThreeDepositAll-是否全冲
    '出参:intNotDelCashCount-返回不允许退现的个数
    '     intThreeCount-三方帐户个数
    '     blnThreeDepositAfter-三方帐户余额存退(true:存在余额退款,False-不存在余额退款)
    '     strDelThreeNames-发生三方帐户余额退款的名称串，比如：招行,建行
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-10-25 11:59:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTotal As Double, varData As Variant, dbl冲预交 As Double
    Dim strStyle As String, i As Long
    Dim blnNotDeposit As Boolean '是否无冲预交
    
    
    On Error GoTo errHandle
     
    
    blnThreeDepositAfter = False
    dblTotal = RoundEx(Val(txtTotal.Text) - GetYBMoney, 2)
    
   
    If mrsCardType Is Nothing Then
        Call initCardSquareData
    ElseIf mrsCardType.State <> 1 Then
        Call initCardSquareData
    End If
    intThreeCount = 0
    intNotDelCashCount = 0
    blnNotDeposit = True
    With mshDeposit
        For i = 1 To .Rows - 1
            dbl冲预交 = Val(.TextMatrix(i, .ColIndex("冲预交")))
            If dbl冲预交 <> 0 Then blnNotDeposit = False
            
            ' 卡类别ID ||代扣||退现||缺省退现
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||||", "||")
            If Val(varData(0)) <> 0 Then
                If Val(varData(3)) = 0 And ((dblTotal - dbl冲预交) <= 0 Or dbl冲预交 = 0) Then   '非缺省退现
                    mrsCardType.Filter = "ID=" & Val(varData(0))
                    intThreeCount = intThreeCount + 1
                    If Not mrsCardType.EOF Then
                        If InStr(strStyle & ",", "," & Nvl(mrsCardType!名称) & ",") = 0 Then
                            strStyle = strStyle & "," & mrsCardType!名称
                            If Val(varData(2)) = 0 Then
                               intNotDelCashCount = intNotDelCashCount + 1
                            End If
                        End If
                    End If
                End If
            End If
            
            If FormatEx(dblTotal - dbl冲预交, 6) <= 0 Then
                dblTotal = 0
            Else
                dblTotal = FormatEx(dblTotal - dbl冲预交, 6)
            End If
        Next
    End With
    If blnThreeDepositAll = False Then
        strStyle = "":: strDelThreeNames = ""
    End If
    If (blnNotDeposit And mblnThreeDepositAll = False) Or intThreeCount = 0 Then
        strStyle = "": blnThreeDepositAfter = False: strDelThreeNames = ""
        GetDelThreeCardDepositInfor = True
        Exit Function
    End If
    
    If intThreeCount >= 1 And InStr(1, mstrPrivs, ";仅用预交结帐;") = 0 Then blnThreeDepositAfter = True
    
    If strStyle <> "" Then strStyle = Mid(strStyle, 2)
    strDelThreeNames = strStyle

    GetDelThreeCardDepositInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetIsHaveThirdDeposit(ByRef strDelThirds As String, Optional ByRef intCount As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取是否有三方卡预结结算
    '入参:
    '出参:strDelThirds-退的三方卡信息
    '     intCount-退的个数
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-09 10:43:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim varData As Variant
    On Error GoTo errHandle
    strDelThirds = ""
    With mshDeposit
        .Redraw = False
        For i = 1 To .Rows - 1
            '卡类别ID||代扣||退现
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||", "||")
            If Val(varData(0)) <> 0 And Val(varData(2)) = 0 Then
                intCount = intCount + 1
                mrsCardType.Filter = "ID=" & Val(varData(0))
                If Not mrsCardType.EOF Then
                    If InStr(strDelThirds & ",", "," & mrsCardType!名称 & ",") = 0 Then
                        strDelThirds = strDelThirds & "," & mrsCardType!名称
                    End If
                End If
            End If
        Next
        .Redraw = True
    End With
    GetIsHaveThirdDeposit = intCount >= 1 And InStr(1, mstrPrivs, ";仅用预交结帐;") = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 



Private Function RecalDepsit(ByVal bytFunc As Byte, ByVal bytMzDeposit As Byte, _
    ByRef cur结帐合计 As Currency, Optional ByRef blnShowThird_Out As Boolean, Optional blnThreeDepositAll_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算冲预交
    '入参:bytFun-0-门诊结帐;1-住院结帐
    '     bytMzDeposit-针对门诊结帐有效,0-表示全清;1-代表根据结帐金额来分摊预交;2-预交款全冲
    '出参:cur结帐合计-返回结帐金客
    '     blnShowThird_Out-是否显示三方卡信息
    '     blnThreeDepositAll_Out -是否所有三方预交
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-09 11:39:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln负数 As Boolean, i As Long
    Dim varData As Variant
    
    '设置冲预交(结帐合计 - 保险合计)
    
    On Error GoTo errHandle
        
    blnShowThird_Out = False
    blnThreeDepositAll_Out = False
    
    bln负数 = cur结帐合计 < 0
    If InStr(1, mstrPrivs, ";仅用预交结帐;") > 0 Then
        With mshDeposit
            For i = 1 To mshDeposit.Rows - 1
                If cur结帐合计 = 0 Then
                    .TextMatrix(i, .ColIndex("冲预交")) = "0.00"
                Else
                    If Val(.TextMatrix(i, .ColIndex("余额"))) <= Format(cur结帐合计, "0.00") Then
                         .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                    Else
                         .TextMatrix(i, .ColIndex("冲预交")) = Format(cur结帐合计, "0.00")
                    End If
                    cur结帐合计 = cur结帐合计 - Val(.TextMatrix(i, .ColIndex("冲预交")))
                End If
            Next i
        End With
        RecalDepsit = True
        Exit Function
    End If
    
    blnThreeDepositAll_Out = False
    If (bytFunc <> 0 And (opt出院.Value Or gbln中途结帐退预交)) _
        Or (bytFunc = 0 And bytMzDeposit = 2) Then
        blnThreeDepositAll_Out = True
    End If
     
            
    blnShowThird_Out = False
    With mshDeposit
        For i = 1 To .Rows - 1
            '卡类别ID||代扣||是否退现||缺省退现
            varData = Split(.Cell(flexcpData, i, .ColIndex("ID")) & "||||||", "||")
            
            If (mbytFunc = 0 And bytMzDeposit = 0) Or (Val(varData(0)) <> 0 And bln负数) Then
                '清空
                .TextMatrix(i, .ColIndex("冲预交")) = "0.00"
                mstrForceNote = ""
            Else
                If Val(varData(0)) <> 0 Then
                    '卡类别ID||代扣||是否退现||缺省退现
                    If Val(.TextMatrix(i, .ColIndex("余额"))) <= Format(cur结帐合计, "0.00") Or Val(varData(3)) = 1 Or mstrForceNote <> "" Then
                         .TextMatrix(i, mshDeposit.ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                        cur结帐合计 = cur结帐合计 - Val(mshDeposit.TextMatrix(i, .ColIndex("冲预交")))
                    Else
                         .TextMatrix(i, .ColIndex("冲预交")) = Format(cur结帐合计, "0.00")
                        cur结帐合计 = 0
                    End If
                    If Not blnShowThird_Out Then blnShowThird_Out = Val(.TextMatrix(i, .ColIndex("余额"))) <> Val(.TextMatrix(i, .ColIndex("冲预交")))
                Else
                    If (bytFunc <> 0 And (opt出院.Value Or gbln中途结帐退预交)) _
                        Or (bytFunc = 0 And bytMzDeposit = 2) Or Val(.TextMatrix(i, .ColIndex("余额"))) <= Format(cur结帐合计, "0.00") Then
                        '全冲
                        .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(.TextMatrix(i, .ColIndex("余额"))), "0.00")
                        cur结帐合计 = cur结帐合计 - Val(.TextMatrix(i, .ColIndex("冲预交")))
                    Else '根据结帐金额分摊
                        .TextMatrix(i, .ColIndex("冲预交")) = Format(cur结帐合计, "0.00")
                        cur结帐合计 = 0
                    End If
                End If
            End If
        Next i
    End With
    '显示三方帐户退款金额
    Call ShowDelThreeSwap
    RecalDepsit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



