VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "*\A..\ZLIDKIND\zlIDKind.vbp"
Begin VB.Form frmOutDoctorStation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "门诊医生工作站"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13650
   Icon            =   "frmOutDoctorStation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleMode       =   0  'User
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl rptNotify 
      Height          =   180
      Left            =   1305
      TabIndex        =   113
      Top             =   30
      Width           =   255
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   317
      _StockProps     =   0
   End
   Begin VB.PictureBox picYZ 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   840
      ScaleHeight     =   2895
      ScaleWidth      =   3855
      TabIndex        =   108
      Top             =   3000
      Width           =   3855
      Begin VB.CommandButton cmdOtherFilter 
         Caption         =   "更多条件"
         Height          =   300
         Left            =   2400
         TabIndex        =   112
         Top             =   0
         Width           =   1100
      End
      Begin VB.ComboBox cboSelectTime 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   15
         Width           =   1230
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   2325
         Index           =   2
         Left            =   0
         TabIndex        =   109
         Top             =   360
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgPati"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "挂号单"
            Object.Width           =   1905
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "门诊号"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "姓名"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "性别"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "年龄"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "急"
            Object.Width           =   635
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "复"
            Object.Width           =   635
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Key             =   "_社区"
            Text            =   "社区"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "时间"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "就诊医生"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Key             =   "就诊卡号"
            Text            =   "就诊卡号"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "病人类型"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Key             =   "西医诊断"
            Text            =   "西医诊断"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Key             =   "中医诊断"
            Text            =   "中医诊断"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl就诊时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊时间"
         Height          =   180
         Left            =   0
         TabIndex        =   111
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox picTmphwnd 
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   4800
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   106
      Top             =   120
      Width           =   15
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2040
      ScaleHeight     =   270
      ScaleWidth      =   495
      TabIndex        =   104
      Top             =   875
      Width           =   495
      Begin VB.Label lblFind 
         Caption         =   "查找:"
         Height          =   255
         Left            =   40
         TabIndex        =   105
         Top             =   40
         Width           =   500
      End
   End
   Begin zlIDKind.PatiIdentify PatiIdentify 
      Height          =   270
      Left            =   2520
      TabIndex        =   25
      Top             =   870
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmOutDoctorStation.frx":058A
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      IDKindAppearance=   0
      ShowPropertySet =   -1  'True
      DefaultCardType =   "就诊卡"
      IDKindWidth     =   555
      FindPatiShowName=   0   'False
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   99
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ListView lvwPatiHZ 
      Height          =   2205
      Left            =   270
      TabIndex        =   95
      Top             =   5565
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgPati"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "挂号单"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "门诊号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "姓名"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "性别"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "年龄"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "急"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "复"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   "_社区"
         Text            =   "社区"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "时间"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "就诊医生"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "就诊卡号"
         Text            =   "就诊卡号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "病人类型"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "传染病"
         Object.Width           =   2540
      EndProperty
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   375
      Left            =   13080
      TabIndex        =   93
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.PictureBox picPatiInput 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EDEDED&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   5535
      Left            =   3720
      ScaleHeight     =   5535
      ScaleWidth      =   9255
      TabIndex        =   41
      Top             =   360
      Width           =   9255
      Begin VB.PictureBox PicOutDoc 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   2385
         Left            =   0
         ScaleHeight     =   2385
         ScaleWidth      =   9195
         TabIndex        =   64
         Top             =   2520
         Width           =   9200
         Begin VB.CommandButton cmdImportEPRDemo 
            Caption         =   "导入范文(&I)"
            Height          =   350
            Left            =   6480
            TabIndex        =   102
            Top             =   2040
            Width           =   1200
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "全文编辑(&U)"
            Height          =   350
            Left            =   7800
            TabIndex        =   98
            Top             =   2040
            Width           =   1200
         End
         Begin VB.PictureBox picPrompt 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4545
            Picture         =   "frmOutDoctorStation.frx":0627
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   79
            ToolTipText     =   "查看所有过敏信息"
            Top             =   2078
            Width           =   260
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "取消签名(&Q)"
            Height          =   350
            Left            =   7820
            TabIndex        =   32
            Top             =   1677
            Width           =   1200
         End
         Begin VB.PictureBox picSentence 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   720
            ScaleHeight     =   240
            ScaleWidth      =   1155
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1185
            Begin VB.TextBox txtSentence 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Left            =   15
               TabIndex        =   66
               Top             =   30
               Width           =   930
            End
            Begin VB.Image imgSentence 
               Height          =   210
               Left            =   960
               Picture         =   "frmOutDoctorStation.frx":0A28
               ToolTipText     =   "请按 * 号键选择"
               Top             =   15
               Width           =   180
            End
         End
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   795
            Index           =   3
            Left            =   4935
            TabIndex        =   30
            Top             =   780
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmOutDoctorStation.frx":0F52
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   795
            Index           =   4
            Left            =   375
            TabIndex        =   31
            Top             =   1500
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmOutDoctorStation.frx":0FEF
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   795
            Index           =   0
            Left            =   380
            TabIndex        =   27
            Top             =   0
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmOutDoctorStation.frx":108C
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   795
            Index           =   2
            Left            =   375
            TabIndex        =   29
            Top             =   780
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmOutDoctorStation.frx":1129
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   795
            Index           =   1
            Left            =   4935
            TabIndex        =   28
            Top             =   0
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmOutDoctorStation.frx":11C6
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
         Begin VB.Label lbl病历名称 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "(门诊病历)"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4920
            TabIndex        =   75
            Top             =   1780
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lbl提示 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输入病历时按 ~ 键可提取或选择词句示范."
            ForeColor       =   &H00404040&
            Height          =   180
            Left            =   4920
            TabIndex        =   74
            Top             =   2115
            Width           =   3420
         End
         Begin VB.Label lbl医生 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "刘力红"
            ForeColor       =   &H00404040&
            Height          =   180
            Index           =   1
            Left            =   7215
            TabIndex        =   73
            Top             =   1785
            Width           =   540
         End
         Begin VB.Label lblDoc 
            BackStyle       =   0  'Transparent
            Caption         =   "主  诉"
            Height          =   540
            Index           =   0
            Left            =   120
            TabIndex        =   72
            Top             =   0
            Width           =   180
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "家族史"
            Height          =   540
            Index           =   3
            Left            =   4680
            TabIndex        =   71
            Top             =   907
            Width           =   180
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDoc 
            BackStyle       =   0  'Transparent
            Caption         =   "查  体"
            Height          =   540
            Index           =   4
            Left            =   120
            TabIndex        =   70
            Top             =   1627
            Width           =   180
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "过去史"
            Height          =   540
            Index           =   2
            Left            =   120
            TabIndex        =   69
            Top             =   907
            Width           =   180
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbl医生 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医生:"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   0
            Left            =   6705
            TabIndex        =   68
            Top             =   1785
            Width           =   450
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "现病史"
            Height          =   540
            Index           =   1
            Left            =   4680
            TabIndex        =   67
            Top             =   0
            Width           =   180
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox PicPatiInfo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EDEDED&
         BorderStyle     =   0  'None
         ForeColor       =   &H00808080&
         Height          =   645
         Left            =   0
         ScaleHeight     =   645
         ScaleWidth      =   9255
         TabIndex        =   62
         Top             =   4920
         Width           =   9255
         Begin VB.Label lblTitle号类 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "号类:"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   3455
            TabIndex        =   91
            Top             =   80
            Width           =   450
         End
         Begin VB.Label lblShow 
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   3935
            TabIndex        =   90
            Top             =   80
            Width           =   570
         End
         Begin VB.Label lblTitle社区号 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "社区号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   7125
            TabIndex        =   89
            Top             =   80
            Width           =   630
         End
         Begin VB.Label lblShow 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   7770
            TabIndex        =   88
            Top             =   80
            Width           =   1245
         End
         Begin VB.Label lblShow 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   5295
            TabIndex        =   87
            Top             =   80
            Width           =   1785
         End
         Begin VB.Label lblTitle医保号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医保号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   4560
            TabIndex        =   86
            Top             =   80
            Width           =   630
         End
         Begin VB.Label lblShow 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   1900
            TabIndex        =   85
            Top             =   80
            Width           =   1530
         End
         Begin VB.Label lblTitle付款 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付款:"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   1450
            TabIndex        =   84
            Top             =   80
            Width           =   450
         End
         Begin VB.Label lblShow 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   510
            TabIndex        =   83
            Top             =   80
            Width           =   930
         End
         Begin VB.Label lblTitle费别 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费别:"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   80
            TabIndex        =   82
            Top             =   80
            Width           =   450
         End
         Begin VB.Label lblDiag 
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   510
            TabIndex        =   76
            Top             =   380
            Width           =   8610
         End
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "诊断:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   80
            TabIndex        =   63
            Top             =   380
            Width           =   450
         End
      End
      Begin VB.PictureBox PicBasis 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         ForeColor       =   &H00808080&
         Height          =   2445
         Left            =   0
         ScaleHeight     =   2445
         ScaleWidth      =   9195
         TabIndex        =   42
         Top             =   0
         Width           =   9195
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   3690
            MaxLength       =   20
            TabIndex        =   23
            Text            =   "#"
            Top             =   1800
            Width           =   1425
         End
         Begin VB.PictureBox picPatient 
            Height          =   780
            Left            =   4800
            ScaleHeight     =   720
            ScaleWidth      =   990
            TabIndex        =   107
            Top             =   30
            Visible         =   0   'False
            Width           =   1050
            Begin VB.Image imgPatient 
               Height          =   705
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   975
            End
         End
         Begin VB.Frame fraLine 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   7200
            TabIndex        =   100
            Top             =   1155
            Width           =   1740
            Begin VB.ComboBox cboEdit 
               BackColor       =   &H00FDFDFD&
               Height          =   300
               Index           =   6
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   -30
               Width           =   1695
            End
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   10
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   8
            Text            =   "#"
            Top             =   885
            Width           =   435
         End
         Begin VB.Frame fraLine 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   3500
            TabIndex        =   97
            Top             =   855
            Width           =   800
            Begin VB.ComboBox cboEdit 
               BackColor       =   &H00FDFDFD&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   5
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   -30
               Width           =   900
            End
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   5370
            MaxLength       =   20
            TabIndex        =   20
            Text            =   "67071232,13320235008"
            Top             =   1470
            Width           =   1780
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H00FDFDFD&
            Caption         =   "复诊"
            Height          =   255
            Index           =   1
            Left            =   5955
            TabIndex        =   16
            Top             =   1155
            Width           =   675
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H00FDFDFD&
            Caption         =   "初诊"
            Height          =   255
            Index           =   0
            Left            =   5280
            TabIndex        =   15
            Top             =   1155
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   9
            Left            =   6720
            MaxLength       =   100
            TabIndex        =   12
            Text            =   "#"
            Top             =   885
            Width           =   1920
         End
         Begin MSComctlLib.ImageList ilexpand 
            Left            =   8760
            Top             =   1080
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorStation.frx":1263
                  Key             =   "展开"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorStation.frx":15FD
                  Key             =   "折叠"
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox picExpand 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            ForeColor       =   &H00808080&
            Height          =   260
            Left            =   8400
            Picture         =   "frmOutDoctorStation.frx":1997
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   94
            Top             =   0
            Width           =   260
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   200
            Index           =   8
            Left            =   7695
            MaxLength       =   64
            TabIndex        =   21
            Text            =   "#"
            Top             =   1470
            Width           =   6060
         End
         Begin VB.CommandButton cmdAller 
            Caption         =   "历史"
            Height          =   300
            Left            =   5685
            Style           =   1  'Graphical
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   1710
            Width           =   480
         End
         Begin VB.Frame fraLine 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   525
            TabIndex        =   60
            Top             =   855
            Width           =   1510
            Begin VB.ComboBox cboEdit 
               BackColor       =   &H00FDFDFD&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   -25
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   -30
               Width           =   1590
            End
         End
         Begin VB.Frame fraLine 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   3885
            TabIndex        =   58
            Top             =   80
            Width           =   580
            Begin VB.ComboBox cboEdit 
               BackColor       =   &H00FDFDFD&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   0
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   -30
               Width           =   660
            End
         End
         Begin VB.Frame fraLine 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   3495
            TabIndex        =   59
            Top             =   555
            Width           =   450
            Begin VB.ComboBox cboEdit 
               BackColor       =   &H00FDFDFD&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   -30
               Width           =   530
            End
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   1
            Left            =   2835
            MaxLength       =   10
            TabIndex        =   5
            Text            =   "#"
            Top             =   585
            Width           =   600
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   750
            MaxLength       =   18
            TabIndex        =   22
            Text            =   "51023219780124511x"
            Top             =   1710
            Width           =   1610
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   6
            Left            =   3680
            MaxLength       =   64
            TabIndex        =   19
            Text            =   "#"
            Top             =   1470
            Width           =   750
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   200
            Index           =   0
            Left            =   2805
            MaxLength       =   64
            TabIndex        =   1
            Text            =   "#"
            Top             =   110
            Width           =   980
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "…"
            Height          =   220
            Index           =   0
            Left            =   2720
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   1170
            Width           =   240
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "…"
            Height          =   220
            Index           =   1
            Left            =   2720
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   1470
            Width           =   240
         End
         Begin MSMask.MaskEdBox txt出生日期 
            Height          =   225
            Left            =   930
            TabIndex        =   3
            Top             =   465
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            BackColor       =   16645629
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt发病日期 
            Height          =   225
            Left            =   4410
            TabIndex        =   10
            Top             =   885
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            BackColor       =   16645629
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt发病时间 
            Height          =   225
            Left            =   5355
            TabIndex        =   11
            Top             =   885
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            BackColor       =   16645629
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   5
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt出生时间 
            Height          =   225
            Left            =   1875
            TabIndex        =   4
            Top             =   585
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            BackColor       =   16645629
            AutoTab         =   -1  'True
            MaxLength       =   5
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   3
            Left            =   520
            MaxLength       =   100
            TabIndex        =   13
            Text            =   "#"
            Top             =   1170
            Width           =   2400
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   5
            Left            =   520
            MaxLength       =   100
            TabIndex        =   18
            Text            =   "#"
            Top             =   1470
            Width           =   2400
         End
         Begin VB.Frame fraRegistInput 
            BackColor       =   &H00FDFDFD&
            Height          =   435
            Left            =   80
            TabIndex        =   45
            Top             =   -80
            Width           =   2040
            Begin VB.Frame fraLine 
               BackColor       =   &H00FDFDFD&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   450
               TabIndex        =   61
               Top             =   135
               Width           =   1520
               Begin VB.ComboBox cboRegist 
                  ForeColor       =   &H00C00000&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Left            =   -60
                  Style           =   2  'Dropdown List
                  TabIndex        =   0
                  Top             =   -30
                  Width           =   1620
               End
            End
            Begin VB.Label lblRegistInput 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "就诊"
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   60
               TabIndex        =   46
               Top             =   160
               Width           =   350
            End
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   3810
            MaxLength       =   20
            TabIndex        =   14
            Text            =   "#"
            Top             =   1170
            Width           =   1425
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAller 
            Height          =   255
            Left            =   6150
            TabIndex        =   24
            Top             =   1725
            Width           =   5730
            _cx             =   10107
            _cy             =   450
            Appearance      =   2
            BorderStyle     =   0
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
            BackColor       =   16579836
            ForeColor       =   -2147483640
            BackColorFixed  =   16579836
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   16579836
            BackColorAlternate=   16579836
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   14737632
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   0
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmOutDoctorStation.frx":1D21
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   115
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
            Editable        =   2
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
         Begin zl9CISJob.UCPatiVitalSigns ucPatiVitalSigns 
            Height          =   285
            Left            =   120
            TabIndex        =   103
            Top             =   2040
            Width           =   10530
            _ExtentX        =   17251
            _ExtentY        =   503
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
            ForeColor       =   0
            ShowMode        =   0
            Style           =   1
            XDis            =   200
            YDis            =   0
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "手机号:"
            Height          =   180
            Index           =   11
            Left            =   3000
            TabIndex        =   114
            Top             =   1815
            Width           =   630
         End
         Begin VB.Label lblRec 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "记"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   390
            Left            =   8685
            TabIndex        =   101
            Top             =   210
            Width           =   405
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "去向:"
            Height          =   180
            Index           =   15
            Left            =   6720
            TabIndex        =   96
            Top             =   1185
            Width           =   450
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发病地址:"
            Height          =   180
            Index           =   9
            Left            =   5880
            TabIndex        =   26
            Top             =   900
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "摘要:"
            Height          =   180
            Index           =   8
            Left            =   7245
            TabIndex        =   92
            Top             =   1605
            Width           =   450
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "过敏:"
            Height          =   180
            Index           =   14
            Left            =   5235
            TabIndex        =   81
            Top             =   1770
            Width           =   450
         End
         Begin VB.Label lbl多科就诊 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "当日多科就诊"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   8415
            TabIndex        =   78
            Top             =   1185
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lbl急 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "急"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   8760
            TabIndex        =   77
            Top             =   0
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住址:"
            Height          =   180
            Index           =   5
            Left            =   75
            TabIndex        =   57
            Top             =   1470
            Width           =   450
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生日期:"
            Height          =   180
            Index           =   24
            Left            =   75
            TabIndex        =   56
            Top             =   600
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄:"
            Height          =   180
            Index           =   1
            Left            =   2385
            TabIndex        =   55
            Top             =   600
            Width           =   450
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名:"
            Height          =   180
            Index           =   0
            Left            =   2385
            TabIndex        =   54
            Top             =   120
            Width           =   450
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "职业:"
            Height          =   180
            Index           =   2
            Left            =   75
            TabIndex        =   53
            Top             =   900
            Width           =   450
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份证:"
            Height          =   180
            Index           =   20
            Left            =   60
            TabIndex        =   52
            Top             =   1725
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位:"
            Height          =   180
            Index           =   3
            Left            =   75
            TabIndex        =   51
            Top             =   1185
            Width           =   450
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位电话:"
            Height          =   180
            Index           =   4
            Left            =   3000
            TabIndex        =   50
            Top             =   1185
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "家庭电话:"
            Height          =   180
            Index           =   7
            Left            =   4560
            TabIndex        =   49
            Top             =   1485
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "监护人:"
            Height          =   180
            Index           =   6
            Left            =   3000
            TabIndex        =   48
            Top             =   1485
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发病时间:"
            Height          =   180
            Index           =   21
            Left            =   2160
            TabIndex        =   47
            Top             =   900
            Width           =   810
         End
      End
   End
   Begin VB.Timer timRefresh 
      Interval        =   1000
      Left            =   3000
      Top             =   75
   End
   Begin VB.Frame fraRoom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   9210
      TabIndex        =   39
      Top             =   7545
      Width           =   300
      Begin VB.Label lblRoom 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   300
      End
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   2205
      Index           =   0
      Left            =   -15
      TabIndex        =   33
      Top             =   3720
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgPati"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "挂号单"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "门诊号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "姓名"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "性别"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "年龄"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "急"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "复"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   "_社区"
         Text            =   "社区"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "就诊诊室"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "就诊医生"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "号序"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "分诊时间"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "挂号时间"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Key             =   "就诊卡号"
         Text            =   "就诊卡号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "病人类型"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "转诊状态"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   38
      Top             =   7920
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOutDoctorStation.frx":1D87
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19288
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   1843
            MinWidth        =   1843
            Text            =   "诊室闲"
            TextSave        =   "诊室闲"
            Object.ToolTipText     =   "诊室状态(鼠标点击可设置)"
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   1875
      Left            =   3720
      TabIndex        =   37
      Top             =   6000
      Width           =   9210
      _Version        =   589884
      _ExtentX        =   16245
      _ExtentY        =   3307
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   3480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":2619
            Key             =   "候诊"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":2BB3
            Key             =   "就诊"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":314D
            Key             =   "已诊"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":36E7
            Key             =   "转诊"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":3C81
            Key             =   "拒绝"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":421B
            Key             =   "暂停"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":47B5
            Key             =   "消息"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   2205
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   2400
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgPati"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "挂号单"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "门诊号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "姓名"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "性别"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "年龄"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "急"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "复"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   "_社区"
         Text            =   "社区"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "就诊时间"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "就诊卡号"
         Text            =   "就诊卡号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "病人类型"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "转诊状态"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "传染病"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwIncept 
      Height          =   2205
      Left            =   360
      TabIndex        =   35
      Top             =   1080
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgPati"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "挂号单"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "门诊号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "姓名"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "性别"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "年龄"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "急"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "复"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   "_社区"
         Text            =   "社区"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "挂号时间"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "就诊卡号"
         Text            =   "就诊卡号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "病人类型"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "转诊状态"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView lvwReserve 
      Height          =   2205
      Left            =   495
      TabIndex        =   36
      Top             =   315
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgPati"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "挂号单"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "门诊号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "姓名"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "性别"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "年龄"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "急"
         Object.Width           =   635
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "预约医生"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "预约时间"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "身份证号"
         Object.Width           =   2541
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "就诊卡号"
         Text            =   "就诊卡号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "病人类型"
         Object.Width           =   2540
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmOutDoctorStation.frx":4B07
      Left            =   960
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
      ScaleMode       =   1
   End
End
Attribute VB_Name = "frmOutDoctorStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COLOR_FREE As Long = &HC000&
Private Const COLOR_BUSY As Long = &HFF&
Private Const DColor = &HEEEEEE, EColor = &HFDFDFD, HColor = &HFFDFDF

Private Enum PatiType
    pt候诊 = 0
    pt就诊 = 1
    pt已诊 = 2
    pt转诊 = 3
    pt预约 = 4
    pt回诊 = 5
    pt排队叫号 = 6
End Enum

Private Enum OPT_ENUM
    opt初诊 = 0
    opt复诊 = 1
End Enum

Private Enum cboEnum
    cbo性别 = 0
    cbo年龄 = 1
    cbo职业 = 2
    cbo血压单位 = 4
    cbo发病时间 = 5
    cbo去向 = 6
End Enum

Private Enum lineDoc
    lineY1 = 0
    lineY2 = 1
    lineX1 = 2
    lineX2 = 3
End Enum

Private Enum txtEnum    '一定要连续编号
    txt姓名 = 0
    txt年龄 = 1
    txt身份证号 = 2
    txt单位名称 = 3
    txt单位电话 = 4
    txt家庭地址 = 5
    txt监护人 = 6
    txt家庭电话 = 7
    txt就诊摘要 = 8
    txt发病地址 = 9
    txt发病 = 10
    txt手机号 = 11
End Enum

Private Enum cmdEnum
    cmd单位名称 = 0
    cmd家庭地址 = 1
End Enum

Private Enum rtfEnum
    txt主诉 = 0
    txt家族史 = 3
    txt现病史 = 1
    txt查体 = 4
    txt过去史 = 2
End Enum

Private Enum lblEditEnum
    lbl姓名 = 0
    lbl单位 = 3
    lbl单位电话 = 4
    lbl家庭电话 = 7
    lbl摘要 = 8
    lbl过敏 = 14
    lbl身份证 = 20
    lbl发病时间 = 21
    lbl去向 = 15
    lbl出生日期 = 24
    lbl体温 = 13
    lbl手机号 = 11
End Enum

Private Enum lblShowEnum
    lbl费别 = 0
    lbl付款 = 1
    lbl号类 = 2
    lbl医保号 = 3
    lbl社区号 = 4
End Enum

Private Enum Msg_Type '消息提醒类别
    m危机值 = 1
    m传染病 = 2
    m处方审查 = 3
    m备血完成 = 4
    m用血审核 = 5
    m输血反应 = 6
End Enum
 
Private Enum NOTIFYREPORT_COLUMN
    c_图标 = 0
    C_病人ID = 1
    C_No = 2
    c_姓名 = 3
    c_门诊号 = 4
    C_就诊时间 = 5
    C_状态 = 6
    '隐藏列
    C_消息 = 7
    C_序号 = 8
    C_日期 = 9
    C_业务 = 10
    C_挂号ID = 11
    C_ID = 12
End Enum

Private Type PatiInfo
    类型 As PatiType
    门诊号 As String
    挂号ID As Long
    挂号单 As String
    科室ID As Long
    诊室 As String
    社区 As Integer
    社区号 As String
    挂号时间 As Date
    数据转出 As Boolean
    病历文件id As Long
    病历id As Long
    保存人 As String
    是否签名 As Boolean
    性别 As String
    婚姻状况 As String
    民族 As String
    国籍 As String
    区域 As String
    出生地点 As String
    传染病上传 As Long
    家庭地址邮编 As String
    单位邮编 As String
    其他证件 As String
    户口地址 As String
    户口地址邮编 As String
    籍贯   As String
    Email As String
    QQ As String
    病人ID As Long
End Type

Private Type ty_Queue
    strQueuePrivs As String '排队叫号虚拟模块权限
    str呼叫站点 As String     '呼叫的站点:空为本站点;否则为其他站点
    byt排队叫号模式 As Byte '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    int呼叫人数 As Integer  '0-不限制,>0表示限制人数
    bln呼叫含回诊 As Boolean   '呼叫是否含回诊人数
    bln医生主动呼叫 As Boolean  'true:表示医生主动呼叫;False-医生非主动呼叫
    strCurrQueueName As String '当前队列名称
    lngcurr挂号ID As Long '当前挂号ID
End Type
Private mty_Queue As ty_Queue

'已诊过滤条件
Private Type COND_FILTER
    Begin As Date
    End As Date
    科室ID As Long
    医生 As String
    挂号单 As String
    门诊号 As String
    就诊卡 As String
    姓名 As String
End Type
Private mvCondFilter As COND_FILTER

'子窗体对象定义
Private mclsEMR As Object  '新版病历zlRichEMR.clsDockEMR
Private WithEvents mclsAdvices As zlPublicAdvice.clsDockOutAdvices
Attribute mclsAdvices.VB_VarHelpID = -1
Private WithEvents mclsEPRs As zlRichEPR.cDockOutEPRs
Attribute mclsEPRs.VB_VarHelpID = -1
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private WithEvents mobjEPRDoc As zlRichEPR.cEPRDocument
Attribute mobjEPRDoc.VB_VarHelpID = -1
Private WithEvents mclsInOutMedRec As zlMedRecPage.clsInOutMedRec
Attribute mclsInOutMedRec.VB_VarHelpID = -1
Private WithEvents mobjQueue As zlQueueManage.clsQueueManage
Attribute mobjQueue.VB_VarHelpID = -1
Private WithEvents mclsDisease As zl9Disease.clsDisease
Attribute mclsDisease.VB_VarHelpID = -1
Private mclsDisDoc As zl9Disease.cDockDisease
Private mcolSubForm As Collection
Private mfrmActive As Form
Private mblnShowLeavePati As Boolean
Private mclsZip As zlRichEPR.cZip
Private mclsUnZip As zlRichEPR.cUnzip
Private mobjKernel As zlPublicAdvice.clsPublicAdvice          '临床核心部件

'参数设置变量
Private mint接诊范围 As Integer '1-本人,2-本诊室,3-本科室
Private mlng接诊科室ID As Long
Private mstr接诊诊室 As String
Private mstr接诊医生 As String
Private mbln要求分诊 As Boolean
Private mintRefresh As Integer '候诊病人刷新间隔(s)
Private mbln自动接诊 As Boolean
Private mlng自动进行 As Long
Private mbln呼叫后接诊 As Boolean
Private mArrDate As Variant

Private mblnDocInput As Boolean    '显示病历快捷输入
Private mblnPatiDetail As Boolean  '显示病人详细信息
Private mblnPatiChange As Boolean '病人信息相关内容改变
Private mblnPatiEditable As Boolean '是否允许修改病人信息

Private mlng接诊控制 As Long '0-不控制 1-禁止 2-提示 问题号:57566
Private mlng提前接收时间 As Long  '当需要对预约号接收进行控制时,该值表明预约号可以提前接收的分钟数 问题号:57566
Private mblnUseTYT As Boolean '使用太元通接口
Private mint过敏输入来源 As Integer '医生站的过敏输入来源
Private mintOutPreTime As Integer

'其它窗体变量
Private mrsAller As ADODB.Recordset '病人过敏记录
Private mstrIDCard As String '最近自动刷出来的身份证号
Private WithEvents mobjIDCard As clsIDCard '身份证对象
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object 'IC卡对象
Private mblnUnRefresh As Boolean
Private mstrPrivs As String
Private mlngModul As Long, mstrVirutalPrivs As String
Private mintActive As PatiType '当前选择病人所在的列表索引
Private mPatiInfo As PatiInfo '历史就诊记录中的,不一定为当前的
Private mlng病人ID As Long, mstr挂号单 As String, mlng科室ID As Long '病人清单中的
Private mlng挂号ID As Long
Private mintFindType As Integer '0-按就诊卡,1-门诊号,2-挂号单,3-姓名查找,4-身份证,5-IC卡
Private mstrFindType As String '存储当前查找类型的名称
Private mblnIsInit As Boolean 'idkind控件初始化标志
Private mblnFindTypeEnabled As Boolean
Private mobjPatient As Object
Private mbln危急值 As Boolean '处危急值的权限

'医疗卡
Private mobjSquareCard As Object      '卡结算对象
Private mstrCardKind As String        '卡结算对象返回的可用的医疗卡
Private Enum CardProperty
    CP短名 = 0
    CP全名 = 1
    CP可读卡 = 2
    CP卡类别ID = 3
    CP卡号长度 = 4
    CP缺省类别 = 5
    CP存在帐户 = 6
    CP卡号密文显示 = 7
End Enum

Private mstrPrePati As String
Private mintPreTime As Integer
Private mblnMouseDown As Boolean
Private mlngCommunityID As Long '自动执行的社区功能
Private mbytSize As Byte '字体 0-小字体（9号字体），1-大字体（12号字体）
Private mblnTabTmp As Boolean
Private mblnSizeTmp As Boolean

Private mblnMsgOk As Boolean '是否有消息来过
Private mblnFirstMsg As Boolean 'mblnFirstMsg=false 表示打开医生站后的第一条消息
Private mintNotify As Integer '医嘱提醒自动刷新间隔(分钟)
Private mintNotifyDay As Integer '提醒多少天内的医嘱
Private mstrNotifyAdvice As String '提醒的医嘱类型
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln消息语音 As Boolean
Private mstrPreNotify As String
Private mblnMaskID As Boolean     '是否身份证掩码显示

Private Sub cboEdit_Click(Index As Integer)
    Dim datCur As Date, datRes As Date
    '编辑状态
    If cboEdit(Index).List(cboEdit(Index).ListIndex) <> cboEdit(Index).Tag Then
        Call SetPermitEscape(False)
    End If
    If Index = cbo发病时间 Then
        If cboEdit(cbo发病时间).ListIndex <= 0 Then Exit Sub
        If Trim(txtEdit(txt发病).Text) = "" Then Exit Sub
        datCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        Select Case cboEdit(cbo发病时间).ListIndex
                Case 1 '小时
                    datRes = DateAdd("n", -1 * Val(txtEdit(txt发病).Text) * 60, CDate(datCur))
                Case 2 '天
                    datRes = DateAdd("h", -1 * Val(txtEdit(txt发病).Text) * 24, CDate(datCur))
                Case 3 '周
                    datRes = DateAdd("d", -1 * 7 * Val(txtEdit(txt发病).Text), CDate(datCur))
                Case 4 '月
                    datRes = DateAdd("M", -1 * Int(Val(txtEdit(txt发病).Text)), CDate(datCur))
                    datRes = DateAdd("d", -1 * (Val(txtEdit(txt发病).Text) - Int(Val(txtEdit(txt发病).Text))) * 30, datRes)
                Case 5 '年
                    If Val(txtEdit(txt发病).Text) < 100 Then
                        datRes = DateAdd("yyyy", -1 * Int(Val(txtEdit(txt发病).Text)), CDate(datCur))
                        datRes = DateAdd("d", -1 * (Val(txtEdit(txt发病).Text) - Int(Val(txtEdit(txt发病).Text))) * 365, datRes)
                    Else
                        MsgBox "发病时间推算不能超过100年。", vbInformation, gstrSysName
                        txtEdit(txt发病).SetFocus: Exit Sub
                    End If
        End Select
        txt发病日期.Text = Format(CDate(datRes), "YYYY-MM-DD")
        If cboEdit(cbo发病时间).ListIndex < 3 Then
            txt发病时间.Text = Format(CDate(datRes), "HH:mm")
        End If
    End If
End Sub

Private Sub cboEdit_GotFocus(Index As Integer)
    If cboEdit(Index).Style = 0 Then
        Call zlControl.TxtSelAll(cboEdit(Index))
    End If
End Sub

Private Sub cboEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngidx As Long
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        Select Case Index
            Case cbo职业
                If SendMessage(cboEdit(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
                lngidx = MatchIndex(cboEdit(Index).hwnd, KeyAscii)
                If lngidx <> -2 Then cboEdit(Index).ListIndex = lngidx
            Case cbo去向
                lngidx = zlControl.CboMatchIndex(cboEdit(Index).hwnd, KeyAscii)
                If lngidx = -1 And cboEdit(Index).ListCount > 0 Then lngidx = 0
                cboEdit(Index).ListIndex = lngidx
        End Select
    End If
End Sub

Private Sub cboRegist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboSelectTime_Click()
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If Me.Visible Then
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mvCondFilter.Begin, mvCondFilter.End, cboSelectTime) Then
                '取消时恢复原来的选择
                Call zlControl.CboSetIndex(cboSelectTime.hwnd, mintOutPreTime)
                Exit Sub
            End If
        ElseIf intDateCount = 0 Then
            '今天  86114
            mvCondFilter.Begin = Format(datCurr, "yyyy-MM-dd 00:00:00")
            mvCondFilter.End = Format(datCurr, "yyyy-MM-dd 23:59:59")
        Else
            mvCondFilter.End = datCurr
            mvCondFilter.Begin = mvCondFilter.End - intDateCount
        End If
    End If
    '选择了时间之后，清除挂号单条件
    mvCondFilter.挂号单 = ""
    mvCondFilter.就诊卡 = ""
    mvCondFilter.门诊号 = ""
    mvCondFilter.姓名 = ""
    '保存参数，保证每个地方提取的出院病人都是在同一时间范围内（72783）
    Call zlDatabase.SetPara("已诊病人结束间隔", DateDiff("d", datCurr, mvCondFilter.End), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
    Call zlDatabase.SetPara("已诊病人开始间隔", DateDiff("d", mvCondFilter.Begin, datCurr), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
    cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
    lbl就诊时间.ToolTipText = cboSelectTime.ToolTipText
    mintOutPreTime = cboSelectTime.ListIndex
    
    Call LoadPatients("0010")
End Sub

Private Sub cmdOtherFilter_Click()
    Dim datCurr As Date
    
    With mvCondFilter
        .科室ID = IIf(.科室ID = 0, mlng接诊科室ID, .科室ID)
        If frmPatiFilter.ShowMe(Me, .Begin, .End, .科室ID, .医生, .挂号单, .门诊号, .就诊卡, .姓名, mstrPrivs) Then
            datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            Call Cbo.SetIndex(cboSelectTime.hwnd, 5)
            '保存参数，保证每个地方提取的出院病人都是在同一时间范围内（72783）
            Call zlDatabase.SetPara("已诊病人结束间隔", DateDiff("d", datCurr, mvCondFilter.End), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            Call zlDatabase.SetPara("已诊病人开始间隔", DateDiff("d", mvCondFilter.Begin, datCurr), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
            lbl就诊时间.ToolTipText = cboSelectTime.ToolTipText
            mintOutPreTime = cboSelectTime.ListIndex
            Call LoadPatients("0010")
        End If
    End With
End Sub

Private Sub cmdAller_Click()
    Dim i As Long, strTmp As String
    Dim objBar As CommandBar
    
    mrsAller.Filter = "病人ID=" & mPatiInfo.病人ID & " and 挂号单<>'" & mPatiInfo.挂号单 & "'"
    If mrsAller.RecordCount > 0 Then
        Set objBar = cbsMain.Add("过敏记录", xtpBarPopup)
        With mrsAller
            For i = 1 To .RecordCount
                If Not IsNull(!挂号时间) Then
                    strTmp = Format(!过敏时间, "yyyy-MM-dd HH:mm") & ",门诊就诊:" & Nvl(!挂号科室) & "," & Nvl(!药物名)
                Else
                    strTmp = Format(!过敏时间, "yyyy-MM-dd HH:mm") & ",第" & !主页ID & "次住院:" & Nvl(!住院科室) & "," & Nvl(!药物名)
                End If
                
                objBar.Controls.Add xtpControlButton, conMenu_Manage_ShowAller * 10 + i, strTmp, -1, False
                .MoveNext
            Next
        End With
        If Not objBar Is Nothing Then objBar.ShowPopup
    End If
End Sub

Private Sub ExecutePaitCancel()
    Dim rsTmp As ADODB.Recordset
    
    If MsgBox("你确实要放弃已改变的内容，重新读取该病人的信息吗？", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    mintPreTime = -1
    Call cboRegist_Click
    Call LoadAllerInfo(mrsAller)
            
    Call SetPermitEscape(True)
End Sub

Private Sub ExecuteOK()
    Dim arrSQL As Variant, i As Long, blnTrans As Boolean
    Dim lng病历ID As Long, blnDoc As Boolean
    Dim objLvw As ListView
    arrSQL = Array()
    
    If InStr(mstrPrivs, "门诊首页") > 0 Then
        If Not CheckOutMediRec Then Exit Sub
    
        Call GetSQLOutMediRec(arrSQL)
    End If
    
    If mblnDocInput And PicOutDoc.Tag = "2" Then
        blnDoc = True
        If mPatiInfo.病历id = 0 Then lng病历ID = zlDatabase.GetNextId("电子病历记录")
        Call GetSQLOutDoc(arrSQL, lng病历ID)
    End If
    
    
    '提交数据
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        
        If blnDoc Then
            If ReadRTFData(lng病历ID) = False Then GoTo errH
            If SaveRTFData(lng病历ID) = False Then GoTo errH
        End If
        
        '社区档案同步
        If Not gobjCommunity Is Nothing And mPatiInfo.社区号 <> "" Then
            If Not gobjCommunity.UpdateInfo(glngSys, p门诊医生站, mPatiInfo.社区, mPatiInfo.社区号, mPatiInfo.病人ID, mPatiInfo.挂号ID) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
    gcnOracle.CommitTrans: blnTrans = False
    If HaveRIS Then
        If gobjRis.HISModPati(1, mlng病人ID, mlng挂号ID) <> 1 Then
            MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
        End If
    ElseIf gbln启用影像信息系统接口 = True Then
        MsgBox "当前启用了影像信息系统接口，但于由RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
    End If
    
    If mintActive = pt回诊 Then
        '回诊:
        Set objLvw = lvwPatiHZ
    ElseIf mintActive = pt转诊 Then
        Set objLvw = lvwIncept
    Else
        Set objLvw = lvwPati(mintActive)
    End If
    
    '更新列表中的姓名，年龄，性别
    If txtEdit(txt姓名).Text <> objLvw.SelectedItem.SubItems(2) Then
         objLvw.SelectedItem.SubItems(2) = txtEdit(txt姓名).Text
    End If
    If cboEdit(cbo性别).Text <> objLvw.SelectedItem.SubItems(3) Then
        objLvw.SelectedItem.SubItems(3) = cboEdit(cbo性别).Text
    End If
    If txtEdit(txt年龄).Text & cboEdit(cbo年龄).Text <> objLvw.SelectedItem.SubItems(4) Then
        objLvw.SelectedItem.SubItems(4) = txtEdit(txt年龄).Text & cboEdit(cbo年龄).Text
    End If
    If mintActive <> pt预约 Then
        If IIf(optState(opt复诊).Value, "复", "") <> objLvw.SelectedItem.SubItems(6) Then
            objLvw.SelectedItem.SubItems(6) = IIf(optState(opt复诊).Value, "复", "")
        End If
    End If
   
    
    '刷新mPatiInfo对象，以及子窗体内容(病历)
    mintPreTime = -1
    Call cboRegist_Click        '由于病人ID没变，SubWinRefreshData中没有刷新病历清单
    If blnDoc Then
        With mPatiInfo
            Call mclsEPRs.zlRefresh(.病人ID, .挂号ID, mlng科室ID, mlng科室ID = .科室ID And (.类型 = pt就诊 Or .类型 = pt回诊) And mlng病人ID <> 0, .数据转出, True)
        End With
    End If
    
    Call ShowAller '重新读取过敏信息
    Call SetPermitEscape(True)
 
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ReadRTFData(ByVal lng病历ID As Long) As Boolean
'功能：读取病历文件的RTF数据到editor控件中
    Dim strZipFile As String, strTempFile As String
    Dim lngRecID As Long
    
    If mPatiInfo.病历id = 0 Then
        lngRecID = lng病历ID
    Else
        lngRecID = mPatiInfo.病历id
    End If
    
    On Error GoTo errH
    strZipFile = zlBlobRead(5, lngRecID)
    strTempFile = zlFileUnzip(strZipFile)
    edtEditor.OpenDoc strTempFile
    
     '删除临时文件
    Kill strTempFile
    Kill strZipFile
   
    ReadRTFData = True
    Exit Function
errH:
    ReadRTFData = False
End Function

Private Function SaveRTFData(ByVal lng病历ID As Long, Optional blnSign As Boolean) As Boolean
'功能：保存病人病历格式RTF数据
'参数：
    Dim strZipFile As String, strTempFile As String, i As Long
    Dim bFinded As Boolean, lngStartPos As Long, lngEndPos As Long, arrTmp As Variant
    Dim strContent As String, lngRecID As Long
    
    If mPatiInfo.病历id = 0 Then
        lngRecID = lng病历ID
    Else
        lngRecID = mPatiInfo.病历id
    End If
    
    If blnSign = False Then
        '替换提纲内容
        edtEditor.Freeze
        edtEditor.ForceEdit = True
        
        For i = 0 To lblDoc.UBound
            bFinded = FindOutLinePosition(edtEditor, CStr(lblDoc(i).Tag), lngStartPos, lngEndPos)
            If bFinded Then
                strContent = rtfEdit(i).Text    '去掉尾部的回车或换行
                Do While Len(strContent) > 2
                    If Mid(strContent, Len(strContent) - 1) = vbLf Or Mid(strContent, Len(strContent) - 1) = vbCr Then
                        strContent = Mid(strContent, Len(strContent) - 1)
                    Else
                        Exit Do
                    End If
                Loop
                edtEditor.Range(lngStartPos, lngEndPos).Text = strContent
            End If
        Next
        
        edtEditor.UnFreeze
        edtEditor.ForceEdit = False
        '要素内容更新
        If mPatiInfo.病历id = 0 Then Call ElementsUpdate(lngRecID)
    End If
    
    On Error GoTo errH
    strTempFile = App.Path & "\TMP.rtf"
    If Dir(strTempFile) <> "" Then Kill strTempFile
    edtEditor.SaveDoc strTempFile
    '压缩文件
    strZipFile = zlFileZip(strTempFile)
    '保存格式
    zlBlobSave 5, lngRecID, strZipFile
    
    '删除临时文件
    Kill strTempFile
    Kill strZipFile

    
    SaveRTFData = True
    Exit Function
errH:
    SaveRTFData = False
End Function

Private Sub EleToString(ByRef edtThis As Object, Ele As cEPRElement)
    Dim sKeyType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bNeeded As Boolean, bBeteenKeys As Boolean
    Dim bForce As Boolean, strOldTag As String
    
    bBeteenKeys = FindNextKey(edtThis, 0, "E", Ele.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bBeteenKeys Then
        Dim lngLen As Long, str内容 As String
        str内容 = Ele.内容文本
        lngLen = Len(str内容)
        With edtThis
            .Freeze
            strOldTag = .Tag
            .Tag = "EleToString"
            bForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = str内容
            .Range(lKSS, lKSS + lngLen).Font.Protected = False
            .Range(lKSS, lKSS + lngLen).Font.Hidden = False
            .Range(lKSS, lKSS + lngLen).Font.BackColor = tomAutoColor
            .Range(lKSS, lKSS + lngLen).Font.Underline = cprNone
            .ForceEdit = bForce
            .UnFreeze
            .Tag = strOldTag
        End With
    End If
End Sub

Private Function GetReplaceEleValue(ByVal ElementName As String, _
    ByVal sPatientID As String, _
    ByVal sPageID As String, _
    ByVal iPatientType As PatiFromEnum, _
    ByVal lng医嘱ID As Long) As String

    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5]) From Dual"
    err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取替换项", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lng医嘱ID)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value))
    End If
    Exit Function

DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

Private Function ElementsUpdate(ByVal lng病历ID As Long) As Boolean
'功能：更新Editor控件中的替换要素内容，以便保存为RTF文件
    Dim ThisElements As New zlRichEPR.cEPRElements
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, lngKey As Long
    Dim bFinded As Boolean, bNeeded As Boolean, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long

    strSQL = "Select 对象标记,ID From 电子病历内容 Where 文件ID= [1] And 对象类型 = 4 And 终止版=0 and 保留对象 =0 And 替换域 =1 order by 对象标记 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病历ID)
    For i = 1 To rsTmp.RecordCount
        lngKey = ThisElements.Add(Nvl(rsTmp("对象标记"), 0))
        ThisElements("K" & lngKey).GetElementFromDB cprET_单病历编辑, rsTmp("ID"), True
        rsTmp.MoveNext
    Next

     For i = 1 To ThisElements.Count
        If ThisElements(i).替换域 = 1 Then
            ThisElements(i).内容文本 = GetReplaceEleValue(ThisElements(i).要素名称, mPatiInfo.病人ID, mPatiInfo.挂号ID, 1, 0)
            bFinded = FindNextKey(edtEditor, 0, "E", ThisElements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            ThisElements(i).Refresh edtEditor
        End If
        If ThisElements(i).替换域 = 1 And ThisElements(i).自动转文本 Then
            EleToString edtEditor, ThisElements(i)     '自动转化为纯文本（暂时不删除该要素）
        End If
    Next
    Set ThisElements = Nothing
End Function


Public Function FindOutLinePosition(ByRef edtThis As Object, ByVal strOName As String, ByRef lngS As Long, lngE As Long) As Boolean
'功能：根据指定的提纲名称，返回提纲内容文本的起止位置
    Dim blnFindedNext As Boolean, lngCur As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, bFinded As Boolean, bNeeded As Boolean
    Dim strTmp As String
    
    bFinded = True
    While bFinded
        bFinded = FindNextKey(edtThis, lngCur, "O", 0, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            strTmp = edtThis.Range(lKEE, lKEE + Len(strOName))
            If strOName = strTmp Then
                lngS = lKEE + Len(strOName)
                blnFindedNext = FindNextAnyKey(edtThis, lngS, strTmp, lKSS, lKSE, lKES, lKEE, 0, bNeeded)
                If blnFindedNext Then
                    lngE = lKSS
                Else
                    lngE = Len(edtThis.Text)
                End If
                Do While lngE > lngS + 1    '去掉尾部的回车或换行
                    If edtThis.Range(lngE - 1, lngE) = vbLf Or edtThis.Range(lngE - 1, lngE) = vbCr Then
                        lngE = lngE - 1
                    Else
                        Exit Do
                    End If
                Loop
                FindOutLinePosition = True
                Exit Function
            Else
                lngCur = lKEE
            End If
        End If
    Wend
End Function


Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset, strSQL As String
    
    err = 0: On Error GoTo ErrHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    strSQL = "Select Zl_Lob_Read([1],[2],[3]" & IIf(blnMoved, ",1", "") & ") as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(strSQL, "zlBlobRead", Action, KeyWord, lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

ErrHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function

Public Function zlBlobSave(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String, strSQL As String
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    err = 0: On Error GoTo ErrHand
    
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        strSQL = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "zlBlobSave")
    Next
    Close lngFileNum
    zlBlobSave = True
    Exit Function

ErrHand:
    Close lngFileNum
    zlBlobSave = False
End Function

'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If strFile = "" Then Exit Function
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    Set mclsZip = New zlRichEPR.cZip
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
    Set mclsZip = Nothing
End Function

Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim objFSO As New Scripting.FileSystemObject    'FSO对象
    
    Dim strZipPath As String
    If strZipFile = "" Then Exit Function
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If objFSO.FileExists(strZipPath & "TMP.RTF") Then objFSO.DeleteFile strZipPath & "TMP.RTF"
    
    Set mclsUnZip = New zlRichEPR.cUnzip
    With mclsUnZip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
    Set mclsUnZip = Nothing
End Function

Public Function FindNextKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindNextKey = True
            End If
        End If
    End With
End Function

Public Function FindNextAnyKey(ByRef edtThis As Object, _
    ByRef lngCurPosition As Long, _
    ByRef strKeyType As String, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef lngKey As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = "S("
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = .TOM.TextDocument.Range(i - 2, i - 1)
                lngKSS = i - 2 '转换为0开始的坐标位置。
                lngKSE = i + 14
                lngKES = j - 2
                lngKEE = j + 14
                lngKey = Val(.TOM.TextDocument.Range(i + 1, i + 9))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 10, i + 11))
                FindNextAnyKey = True
            End If
        End If
    End With
End Function


Private Function CheckOutMediRec() As Boolean
'功能：检查首页输入数据合法性
'返回：
    Dim objTmp As Object, curDate As Date
    Dim arrInfo() As Variant, arrName As Variant
    Dim str身份证 As String, str出生日期 As String, lng性别 As Long
    Dim str年龄 As String, i As Long, j As Long
    
    
    '项目输入的长度检查
    '-----------------------------------------------------------------------------------------
    For Each objTmp In txtEdit
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "输入内容过长，请检查。(该项目最多允许 " & objTmp.MaxLength & " 个字符或 " & objTmp.MaxLength \ 2 & " 个汉字)")
                Exit Function
            End If
        End If
    Next
    For Each objTmp In rtfEdit
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "输入内容过长，请检查。(该项目最多允许 " & objTmp.MaxLength & " 个字符或 " & objTmp.MaxLength \ 2 & " 个汉字)")
                Exit Function
            End If
        End If
    Next
    
    '输入内容的有效性检查
    '-----------------------------------------------------------------------------------------
    
    curDate = zlDatabase.Currentdate
    
            
    '身份证号码检查
    '对身份证号进行验证
    If mblnMaskID Then
        str身份证 = txtEdit(txt身份证号).Tag
    Else
        str身份证 = txtEdit(txt身份证号).Text
    End If
    If str身份证 <> "" And lblEdit(20).Tag <> str身份证 Then
        If Len(str身份证) <> 15 And Len(str身份证) <> 18 Then
            Call ShowMessage(txtEdit(txt身份证号), "身份证号码的长度不正确，应为15位或18位。")
            Exit Function
        End If

        If Len(str身份证) = 15 Then
            str出生日期 = Mid(str身份证, 7, 6)
            str出生日期 = Format(GetFullDate(str出生日期), "yyyy-MM-dd")
            lng性别 = Val(Right(str身份证, 1))
        Else
            str出生日期 = Mid(str身份证, 7, 8)
            str出生日期 = Format(GetFullDate(str出生日期), "yyyy-MM-dd")
            lng性别 = Val(Mid(str身份证, 17, 1))
        End If
        If Not IsDate(str出生日期) Then
            If ShowMessage(txtEdit(txt身份证号), "身份证号码中的出生日期信息不正确，是否继续？", True) = vbNo Then Exit Function
        ElseIf IsDate(txt出生日期.Text) Then
            If Format(str出生日期, "yyyy-MM-dd") <> Format(txt出生日期.Text, "yyyy-MM-dd") Then
                If ShowMessage(txtEdit(txt身份证号), "身份证号码中的出生日期信息与病人的出生日期不符，是否继续？", True) = vbNo Then Exit Function
            End If
        End If
        If (lng性别 Mod 2 = 1 And InStr(cboEdit(cbo性别).Text, "女") > 0) Or (lng性别 Mod 2 = 0 And InStr(cboEdit(cbo性别).Text, "男") > 0) Then
            If ShowMessage(txtEdit(txt身份证号), "身份证号码中的性别信息与病人的性别不符，是否继续？", True) = vbNo Then Exit Function
        End If
    End If
    
    '过敏药物表格检查
    With vsAller
        For i = 0 To .Cols - 1
            If Trim(.TextMatrix(0, i)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(0, i)) > 60 Then
                    .Col = i
                    Call ShowMessage(vsAller, "过敏药物名太长，只允许60个字符或30个汉字。")
                    Exit Function
                End If
                For j = i + 1 To .Cols - 1
                    If Trim(.TextMatrix(0, j)) <> "" Then
                        If .TextMatrix(0, j) = .TextMatrix(0, i) Then
                            .Col = i
                            Call ShowMessage(vsAller, "发现存在两行相同的过敏药物。")
                            Exit Function
                        ElseIf Val(.ColData(i)) <> 0 And Val(.ColData(j)) = Val(.ColData(i)) Then
                            .Col = i
                            Call ShowMessage(vsAller, "发现存在两行相同的过敏药物。")
                            Exit Function
                        ElseIf .TextMatrix(1, i) <> "" And .TextMatrix(1, i) = .TextMatrix(1, j) Then
                            .Col = i
                            Call ShowMessage(vsAller, "发现存在两行相同的过敏药物。")
                            Exit Function
                        End If
                    End If
                Next
            End If
        Next
    End With
    
    '发病时间检查
    If txt发病日期.Text <> "____-__-__" Then
        If Not IsDate(txt发病日期.Text) Then
            Call ShowMessage(txt发病日期, "请输入正确的发病日期。")
            Exit Function
        Else
            If txt发病时间.Text <> "__:__" Then
                If Not IsDate(txt发病时间.Text) Then
                    Call ShowMessage(txt发病时间, "请输入正确的发病时间。")
                    Exit Function
                End If
            End If
            
            If txt发病日期.Text & IIf(txt发病时间.Text = "__:__", "", " " & txt发病时间.Text) _
                >= Format(curDate, txt发病日期.Format & IIf(txt发病时间.Text = "__:__", "", " " & txt发病时间.Format)) Then
                Call ShowMessage(txt发病日期, "发病时间应该早于当前时间。")
                Exit Function
            End If
        End If
    End If
    
    CheckOutMediRec = True
End Function

Private Sub GetSQLOutMediRec(ByRef arrSQL As Variant)
    '功能：保存门诊首页的各种信息
    Dim i As Integer, lngCnt As Long, blnExist As Boolean, curDate As Date
    Dim str生日 As String, str发病 As String
    Dim lng单位ID As Long
    Dim strTmpSQL As String
    
    curDate = zlDatabase.Currentdate
    
    If IsDate(txt出生日期.Text) Then
        If IsDate(txt出生时间.Text) Then
            str生日 = "To_Date('" & Format(txt出生日期.Text & " " & txt出生时间.Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Else
            str生日 = "To_Date('" & Format(txt出生日期.Text, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        End If
    Else
       str生日 = "NULL"
    End If
    If Trim(txtEdit(txt单位名称).Text) <> "" Then
        lng单位ID = Val(txtEdit(txt单位名称).Tag)
    End If
    
    '病人信息
    str发病 = "NULL"
    If IsDate(txt发病日期.Text) Then
        If IsDate(txt发病时间.Text) Then
            str发病 = "To_Date('" & txt发病日期.Text & " " & txt发病时间.Text & "','YYYY-MM-DD HH24:MI')"
        Else
            str发病 = "To_Date('" & txt发病日期.Text & "','YYYY-MM-DD')"
        End If
    End If
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    With mPatiInfo
    arrSQL(UBound(arrSQL)) = "ZL_病人信息_首页整理(" & _
        mPatiInfo.病人ID & ",'" & mPatiInfo.门诊号 & "','" & txtEdit(txt姓名).Text & "'," & _
        "'" & NeedName(cboEdit(cbo性别).Text) & "','" & txtEdit(txt年龄).Text & cboEdit(cbo年龄).Text & "'," & _
        "'" & .民族 & "','" & .国籍 & "','" & .区域 & "','" & .籍贯 & "','" & NeedName(cboEdit(cbo职业).Text) & "'," & _
        str生日 & ",'" & .出生地点 & "','" & txtEdit(txt身份证号).Tag & "','" & .其他证件 & "','" & .婚姻状况 & "'," & _
        "'" & lblShow(lbl付款).Caption & "','" & txtEdit(txt家庭地址).Text & "','" & txtEdit(txt家庭电话).Text & "'," & _
        "'" & .家庭地址邮编 & "','" & .户口地址 & "','" & .户口地址邮编 & "'," & ZVal(lng单位ID) & "," & _
        "'" & txtEdit(txt单位名称).Text & "','" & txtEdit(txt单位电话).Text & "','" & .单位邮编 & "'," & _
        "Null,Null,Null,Null,'" & .Email & "','" & .QQ & "','" & txtEdit(txt监护人).Text & "','" & mstr挂号单 & "'," & _
        IIf(optState(opt复诊).Value, 1, 0) & ",'" & txtEdit(txt就诊摘要).Text & "'," & .传染病上传 & "," & str发病 & "," & _
        "'" & txtEdit(txt发病地址).Text & "')"
    End With
    
    '更新手机号
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病人信息_更新信息(" & mPatiInfo.病人ID & ",'手机号','" & txtEdit(txt手机号).Text & "')"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & mPatiInfo.病人ID & ",'去向','" & cboEdit(cbo去向).Text & "'," & cboRegist.ItemData(cboRegist.ListIndex) & ")"
    strTmpSQL = ucPatiVitalSigns.GetSaveSQL(mlng病人ID, cboRegist.ItemData(cboRegist.ListIndex))
    If strTmpSQL <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strTmpSQL
    End If
    
    
            
    '过敏药物
    With vsAller
        '界面数据未变化时，不调删除过程
        lngCnt = 0: blnExist = False
        For i = 0 To .Cols - 1
            If CStr(.Cell(flexcpData, 0, i)) <> "" Then
                blnExist = True '如果只有一空列，则不调用删除
                If .Cell(flexcpData, 0, i) = .TextMatrix(0, i) Then    '删除的列是清空的
                    lngCnt = lngCnt + 1
                End If
            End If
        Next
        If blnExist And lngCnt <> .Cols - 1 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_病人过敏记录_Delete(" & mPatiInfo.病人ID & "," & mPatiInfo.挂号ID & ",3)"
        End If
        
        If blnExist = False Or blnExist And lngCnt <> .Cols - 1 Then
        For i = 0 To .Cols - 1
            If Trim(.TextMatrix(0, i)) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_病人过敏记录_Insert(" & mPatiInfo.病人ID & "," & mPatiInfo.挂号ID & "," & _
                    "3," & ZVal(.ColData(i)) & ",'" & .TextMatrix(0, i) & "',1," & _
                    "To_Date('" & Format(IIf(.Cell(flexcpData, 1, i) & "" = "", curDate, .Cell(flexcpData, 1, i)), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI:SS'),Null,'" & .TextMatrix(1, i) & "')"
            End If
        Next
        End If
    End With
End Sub


Private Sub GetSQLOutDoc(ByRef arrSQL As Variant, ByVal lng病历ID As Long)
'功能：组织快捷病历的数据保存SQL
'参数：lng病历ID-新增时传入新取的病历ID
    Dim i As Long, k As Long
    Dim strTmp(5) As String
    
    If mPatiInfo.病历id = 0 Then
        For i = 0 To rtfEdit.UBound
            If Trim(rtfEdit(i).Text) <> "" Then Exit For
        Next
        If i > rtfEdit.UBound Then Exit Sub     '新增时，如果没有填内容，则不保存
    End If
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    With mPatiInfo
        If .病历id = 0 Then
            If rtfEdit(txt主诉).Locked Then Exit Sub
            arrSQL(UBound(arrSQL)) = "Zl_简单门诊病历_Update(1," & mPatiInfo.病人ID & "," & _
                .挂号ID & "," & mlng科室ID & "," & .病历文件id & "," & lng病历ID & ",'" & UserInfo.姓名 & "','" & _
                Trim(rtfEdit(txt主诉).Text) & "','" & Trim(rtfEdit(txt家族史).Text) & "','" & Trim(rtfEdit(txt现病史).Text) & "','" & _
                Trim(rtfEdit(txt查体).Text) & "','" & Trim(rtfEdit(txt过去史).Text) & "')"
        Else
            k = 0
            For i = 0 To rtfEdit.UBound
                If rtfEdit(i).Locked = False Then
                    strTmp(i) = rtfEdit(i).Tag & "|" & Trim(rtfEdit(i).Text)
                    k = k + 1
                End If
            Next
            If k = 0 Then Exit Sub
            
            arrSQL(UBound(arrSQL)) = "Zl_简单门诊病历_Update(2," & mPatiInfo.病人ID & "," & _
                .挂号ID & "," & mlng科室ID & ",0," & .病历id & ",'" & UserInfo.姓名 & "','" & _
                strTmp(0) & "','" & strTmp(3) & "','" & strTmp(1) & "','" & strTmp(4) & "','" & strTmp(2) & "')"
        End If
    End With
End Sub

Private Function GetEPRDoc() As zlRichEPR.cEPRDocument
'功能：读取病历文件的RTF数据到editor控件中，并返回文档对象
    Dim objDoc As New zlRichEPR.cEPRDocument
   
    objDoc.InitEPRDoc cprEM_修改, cprET_单病历编辑, mPatiInfo.病历id, cprPF_门诊, mPatiInfo.病人ID, mPatiInfo.挂号ID, , mPatiInfo.科室ID
    If objDoc.ReadFileStructure(edtEditor) = True Then
        Set GetEPRDoc = objDoc
    End If
End Function

Private Sub cmdImportEPRDemo_Click()
    Dim objImportEPRDemo As New frmImportEPRDemo
    Dim rsDemo As New Recordset
    
    If mPatiInfo.病历id <> 0 Then
        MsgBox "该病人已经产生了病历文件，不能再导入范文。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If objImportEPRDemo.ShowMe(Me, mPatiInfo.病历文件id, mPatiInfo.病人ID, mPatiInfo.挂号ID, rsDemo) > 0 Then
        Call SetDocData(rsDemo, 1)
        Call SetPermitEscape(False)
        Call SetRTFEditFontSize
    End If
End Sub

Private Sub cmdSign_Click()
    Dim i As Long, str对象属性 As String, strSource As String, strSQL As String
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim patiSign As cEPRSign, objEPRDoc As cEPRDocument
    
    If mPatiInfo.是否签名 = False Then
        For i = 0 To rtfEdit.UBound
            If Trim(rtfEdit(i).Text) <> "" Then Exit For
        Next
        If i > rtfEdit.UBound Then
            MsgBox "请先输入病历信息后再进行签名。", vbInformation, gstrSysName
            Exit Sub    '新增时，如果没有填内容，则不保存
        End If
                       
        If mblnPatiChange Then
            Call ExecuteOK
            If mblnPatiChange Then Exit Sub    '保存失败则不再继续
        End If
                     
        If edtEditor.Text = "" Then
            If ReadRTFData(mPatiInfo.病历id) = False Then Exit Sub
        End If
        
        strSource = edtEditor.Text
        '76491,未知BUG,不得到焦点，在未保存情况下报错。
        If cmdSign.Visible And cmdSign.Enabled Then cmdSign.SetFocus
        Set patiSign = frmOutDocterSign.ShowMe(Me, strSource, mPatiInfo.病人ID, mPatiInfo.挂号ID)
        If patiSign Is Nothing Then Exit Sub
        With patiSign
            .Key = "1"
            str对象属性 = .签名方式 & ";" & .签名规则 & ";" & .证书ID & ";" & IIf(.显示手签, 1, 0) & ";" & _
                    Format(.签名时间, "yyyy-mm-dd hh:mm:ss") & ";" & .显示时间 & ";" & .签名要素
                    
            strSQL = "Zl_简单门诊病历_签名(1," & mPatiInfo.病历id & ",'" & str对象属性 & "','" & UserInfo.姓名 & "','" & _
                    .前置文字 & "','" & .时间戳 & "','" & .签名级别 & "','" & .签名信息 & "')"
        End With
        
        Set objEPRDoc = GetEPRDoc()
        If objEPRDoc Is Nothing Then Exit Sub
        Call patiSign.InsertIntoEditor(edtEditor, Len(edtEditor.Text), , objEPRDoc)
        Set objEPRDoc = Nothing
        Set patiSign = Nothing
    Else
        If MsgBox("你确定要取消签名吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        Set patiSign = GetSign(mPatiInfo.病历id)
        If patiSign Is Nothing Then Exit Sub
        
        Set objEPRDoc = GetEPRDoc()
        If objEPRDoc Is Nothing Then Exit Sub
        Call patiSign.DeleteFromEditor(edtEditor, objEPRDoc)
        Set objEPRDoc = Nothing
        Set patiSign = Nothing
        
        strSQL = "Zl_简单门诊病历_签名(0," & mPatiInfo.病历id & ")"
    End If
    
   
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        If SaveRTFData(mPatiInfo.病历id, True) = False Then GoTo errH
    gcnOracle.CommitTrans: blnTrans = False
    
        
    Call LoadDocData
    Call SetPermitEdit
    Call SetPermitEscape(True)
    Call PicBasis_Resize
    With mPatiInfo
        Call mclsEPRs.zlRefresh(mPatiInfo.病人ID, .挂号ID, mlng科室ID, mlng科室ID = .科室ID And (.类型 = pt就诊 Or .类型 = pt回诊) And mlng病人ID <> 0, .数据转出, True)
    End With
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Set objEPRDoc = Nothing: Set patiSign = Nothing
    Call SaveErrLog
End Sub
Private Function GetSign(ByVal lng病历ID As Long) As cEPRSign
'功能：获取当前用户的签名对象
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim OneSign As New cEPRSign, intSign As Integer, strUserName As String
    
    strUserName = UserInfo.姓名
    intSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0)
    If intSign = 1 Then
        strSQL = "Select 签名 From 人员表 Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
        If rsTemp.RecordCount > 0 Then
            If Not IsNull(rsTemp!签名) Then strUserName = rsTemp!签名
        End If
    End If
    strSQL = "Select Id,对象标记 From 电子病历内容 Where 文件id= [1] And 对象类型=8 And Instr(';'||内容文本||';',[2])>0 Order By 对象标记"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病历ID, ";" & strUserName & ";")
    If rsTemp.RecordCount > 0 Then
        OneSign.Key = Nvl(rsTemp!对象标记, 0)
        If OneSign.GetSignFromDB(rsTemp!ID) = True Then Set GetSign = OneSign
    End If
End Function

Private Sub cmdUpdate_Click()
    Dim blnDoc As Boolean
    
    If InStr(";" & GetPrivFunc(glngSys, p门诊病历管理) & ";", ";病历书写;") > 0 Then
        blnDoc = mlng科室ID <> 0 And mlng科室ID = mPatiInfo.科室ID And _
                 (mPatiInfo.病历id = 0 And mPatiInfo.病历文件id <> 0 Or mPatiInfo.病历id <> 0) And (mintActive = pt就诊 Or mintActive = pt回诊)
        If blnDoc And mPatiInfo.病历id <> 0 And lbl医生(1).Tag = "0" Then   '没有修改他人病历的权限
            blnDoc = mPatiInfo.保存人 = UserInfo.姓名
        End If
        
        If blnDoc Then
            If mobjEPRDoc Is Nothing Then
                Set mobjEPRDoc = New zlRichEPR.cEPRDocument
            End If
            If mPatiInfo.病历id = 0 And mPatiInfo.病历文件id <> 0 Then '如果没有新建则新建
                Call mobjEPRDoc.InitEPRDoc(0, 2, mPatiInfo.病历文件id, 1, mPatiInfo.病人ID, mPatiInfo.挂号ID, , mPatiInfo.科室ID, , False)
            Else
                Call mobjEPRDoc.InitEPRDoc(1, 2, mPatiInfo.病历id, 1, mPatiInfo.病人ID, mPatiInfo.挂号ID, , mPatiInfo.科室ID, , False)
            End If
            Call mobjEPRDoc.ShowEPREditor(Me)
        Else
            MsgBox "当前病历不能修改。", vbInformation, Me.Caption
        End If
    Else
        MsgBox "您没有病历书写的权限。", vbInformation, Me.Caption
    End If
End Sub



Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    '如果 Pane正在关闭，且是当前活动的病人列表，则禁止关闭
    If Action = PaneActionCollapsing And mintActive > -1 Then
        Select Case mintActive
            Case pt候诊, pt就诊, pt已诊
                If Pane.ID = mintActive + 1 Then
                    Cancel = True
                End If
            Case pt转诊
                If Pane.ID = 4 Then
                    Cancel = True
                End If
            Case pt预约
                If Pane.ID = 5 Then
                    Cancel = True
                End If
            Case pt回诊
                If Pane.ID = 7 Then
                    Cancel = True
                End If
            Case pt排队叫号
                If Pane.ID = pt排队叫号 Then
                    Cancel = True
                End If
        End Select
    End If
End Sub

Private Sub Form_Activate()
    If Check排队叫号 Then
        DoEvents
        mobjQueue.SetFocus
    End If
 
    '进入时不选中任何病人
    'If lvwPati(pt就诊).Visible And lvwPati(pt就诊).Enabled Then Call lvwPati(pt就诊).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        If picSentence.Visible Then
            Call HideWordInput   '隐藏词句输入
        ElseIf mblnPatiEditable And mblnPatiChange Then
            Call ExecutePaitCancel
            Call PicBasis_Resize
        End If
    End If
    '读卡
    PatiIdentify.ActiveFastKey
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("[|']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If

    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And Not Me.ActiveControl Is PatiIdentify And mstrFindType = "就诊卡" And PatiIdentify.Enabled And PatiIdentify.Visible Then
        If Not (ActiveControl.Container Is PicBasis Or ActiveControl.Container Is PicOutDoc Or ActiveControl.Container Is picSentence) Then
            PatiIdentify.Text = UCase(Chr(KeyAscii))
            PatiIdentify.NotAutoSel = True
            PatiIdentify.SetFocus
        End If
    End If
End Sub

Private Sub mclsAdvices_VSKeyPress(KeyAscii As Integer)
    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And mstrFindType = "就诊卡" And PatiIdentify.Enabled And PatiIdentify.Visible Then
        picFind.SetFocus
        PatiIdentify.Text = UCase(Chr(KeyAscii))
        PatiIdentify.NotAutoSel = True
        PatiIdentify.SetFocus
    End If
End Sub

Private Sub InitQueuePara(Optional blnOnlyRefresh医生主动呼叫 As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始化排队叫号参数
    '编制：刘兴洪
    '日期：2010-06-07 16:23:31
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    '刘兴洪:'排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    Dim bytType As Byte
    If blnOnlyRefresh医生主动呼叫 Then GoTo RefreshDoctor:
    
    mty_Queue.strQueuePrivs = ";" & GetPrivFunc(glngSys, p排队叫号虚拟模块) & ";"
    mty_Queue.byt排队叫号模式 = Val(zlDatabase.GetPara("排队叫号模式", glngSys, p门诊分诊管理))
    mty_Queue.str呼叫站点 = zlDatabase.GetPara("远端呼叫站点", glngSys, p排队叫号虚拟模块)
    
RefreshDoctor:
    If mty_Queue.byt排队叫号模式 = 1 Then
   
        mty_Queue.bln医生主动呼叫 = Val(zlGetLocaleComputerNamePara("排队呼叫站点", glngSys, p门诊分诊管理, "0", mty_Queue.str呼叫站点)) = 1
    Else
         mty_Queue.bln医生主动呼叫 = False
    End If
    If mty_Queue.bln医生主动呼叫 Then
        mty_Queue.int呼叫人数 = Val(zlDatabase.GetPara("医生就诊人数", glngSys, p门诊医生站))
    Else
        mty_Queue.int呼叫人数 = 0
    End If
    mty_Queue.bln呼叫含回诊 = Val(zlDatabase.GetPara("就诊人数含回诊", glngSys, p门诊医生站, "1")) = 1
End Sub

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer
    Dim intType As Integer, blnHave As Boolean, blnTmp As Boolean
    Dim i As Integer, arrType() As String
    Dim objControl As CommandBarControl
    Dim arrTmp As Variant, strTmp As String

    mstrPrivs = ";" & gstrPrivs & ";"
    mlngModul = glngModul
    mblnPatiChange = False
    mblnPatiEditable = False
    mblnShowLeavePati = False
    
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, p门诊医生站, GetInsidePrivs(p门诊医生站))
    Call AddMipModule(mclsMipModule)
    
    Set mclsDisease = New zl9Disease.clsDisease
    Call mclsDisease.InitDisease(gcnOracle, Me, glngSys, glngModul, mstrPrivs, mclsMipModule)
    Set mclsDisDoc = New zl9Disease.cDockDisease
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    
    Call InitQueuePara
    Call InitRegist
    '医嘱提醒刷新设置
    mstrNotifyAdvice = zlDatabase.GetPara("自动刷新内容", glngSys, p门诊医生站, "0")
    mintNotifyDay = Val(zlDatabase.GetPara("自动刷新病历审阅天数", glngSys, p门诊医生站, 1))
    mintNotify = Val(zlDatabase.GetPara("自动刷新病历审阅间隔", glngSys, p门诊医生站))
    mbln消息语音 = Val(zlDatabase.GetPara("启用语音提示", glngSys, p门诊医生站)) = 1
    mblnMaskID = Val(zlDatabase.GetPara("身份证加密显示", glngSys)) = 1
    mbln危急值 = InStr(GetInsidePrivs(p门诊医生站), ";危急值处理;") > 0
    
    
    '先读参数，菜单定义中需要判断
    blnTmp = InStr(GetInsidePrivs(p门诊病历管理), "病历书写") > 0
    If blnTmp Then
        lbl医生(1).Tag = IIf(InStr(GetInsidePrivs(p门诊病历管理), "他人病历") > 0, 1, 0)
        
        mblnDocInput = Val(zlDatabase.GetPara("显示病历快捷输入", glngSys, p门诊医生站, 0, , , intType)) = 1
        blnHave = IIf(InStr(GetInsidePrivs(1070), "签名权") > 0, True, False)
        lbl医生(1).Caption = UserInfo.姓名
    Else
        mblnDocInput = False
        blnHave = False
    End If
    mblnPatiDetail = Val(zlDatabase.GetPara("显示病人详细信息", glngSys, p门诊医生站, 0, , , intType)) = 1
    If mblnPatiDetail Then
        Set picExpand.Picture = ilexpand.ListImages("折叠").Picture
    Else
        Set picExpand.Picture = ilexpand.ListImages("展开").Picture
    End If
    
    cmdSign.Visible = blnHave
    lbl医生(0).Visible = blnHave
    lbl医生(1).Visible = blnHave
    PicOutDoc.Visible = mblnDocInput
    mArrDate = Array(txt出生日期, txt出生时间, txt发病日期, txt发病时间)
    
    '一卡通部件初始，须在tbcSub_SelectedChanged之前，以便传递给医嘱部件
     'zlGetIDKindStr中会自动补齐为至少8位属性
    mstrCardKind = "就|就诊卡|0|0|8|0|0|0;门|标识号|0|0|0|0|0|0;挂|挂号单|0|0|0|0|0|0;姓|姓名|0|0|0|0|0|0;身|二代身份证|0|0|0|0|0|0;ＩＣ|ＩＣ卡|1|0|0|0|0|0"
    If Check排队叫号 = True Then mstrCardKind = mstrCardKind & ";排|排队号|0|0|0|0|0|0;医|医保号|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    err.Clear: On Error GoTo 0
    If Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
        Else
            mstrCardKind = mobjSquareCard.zlGetIDKindStr(mstrCardKind)
        End If
    End If
    Call PatiIdentify.zlInit(Me, glngSys, p门诊医生站, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
    PatiIdentify.objIDKind.AllowAutoICCard = True
    PatiIdentify.objIDKind.AllowAutoIDCard = True
    mblnIsInit = True

    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 380, 550, DockLeftOf, Nothing)
    objPane.Title = "候诊病人"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(2, 280, 190, DockBottomOf, objPane)
    objPane.Title = "就诊病人"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPane = Me.dkpMain.CreatePane(7, 280, 250, DockBottomOf, objPane)
    objPane.Title = "回诊病人"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(3, 280, 550, DockBottomOf, objPane)
    objPane.Title = "已诊病人"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPane = Me.dkpMain.CreatePane(8, 280, 180, DockBottomOf, objPane)
    objPane.Title = "消息提醒"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPane = Me.dkpMain.CreatePane(4, 380, 350, DockTopOf, dkpMain.Panes(1))
    objPane.Title = "转诊病人"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    objPane.AttachTo dkpMain.Panes(1)
    Set objPane = Me.dkpMain.CreatePane(5, 380, 550, DockTopOf, dkpMain.Panes(1))
    objPane.Title = "预约病人"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    objPane.AttachTo dkpMain.Panes(1)
     
     
    'TabControl
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
    If GetInsidePrivs(p新版门诊病历, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "电子病历")
        If Not mclsEMR Is Nothing Then
            If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                Set mclsEMR = Nothing
            Else
                
            End If
        End If
    End If
    Set mclsAdvices = New zlPublicAdvice.clsDockOutAdvices
    Set mclsEPRs = New zlRichEPR.cDockOutEPRs
    
    Set mcolSubForm = New Collection
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_新病历"
    End If
    mcolSubForm.Add mclsAdvices.zlGetForm, "_医嘱"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_病历"
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        If GetInsidePrivs(p门诊医嘱下达) <> "" Then
            '先加载医嘱的原因:在启用美康接口，但客户端没有美康部件时。如果先加载排队叫号后加载医嘱的时候，
            '从“病历信息”切换到“医嘱信息”会因弹出Msgbox报错 问题号:67995
            .InsertItem(intIdx, "医嘱信息", mcolSubForm("_医嘱").hwnd, 0).Tag = "医嘱": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p门诊病历管理) <> "" Then
            .InsertItem(intIdx, "病历信息", picTmp.hwnd, 0).Tag = "病历": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p新版门诊病历, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "电子病历", picTmp.hwnd, 0).Tag = "新病历": intIdx = intIdx + 1
        End If
        '外挂提供的卡片
        Call CreatePlugInOK(p门诊医生站)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, p门诊医生站)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, p门诊医生站, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        If .ItemCount = 0 Then
            MsgBox "你没有使用门诊医生工作站的权限。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '恢复上次选择的卡片
        strTab = zlDatabase.GetPara("医护功能", glngSys, p门诊医生站)
        
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '避免激活事件
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            .Item(0).Selected = True '新建时就自动选中了这个,不会再激活事件
        End If
        '只加载选择的子窗体
        Call tbcSub_SelectedChanged(.Selected)
    End With
            
    '读取界面数据
    '-----------------------------------------------------
    mblnUnRefresh = True
    mstrPrePati = ""
    mintPreTime = -1
    mintActive = -1
    
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    On Error Resume Next
    Set mobjICCard = CreateObject("zlICCard.clsICCard")
    If Not mobjICCard Is Nothing Then
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    err.Clear: On Error GoTo 0
    
    Call GetLocalSetting '本地参数
    
    Call InitCondFilter '已诊病人过滤条件
    Call InitReportColumn
    Call InitPatiData
    Call LoadPatients '显示数据
    Call LoadNotify '消息提醒
    
    dkpMain(4).Hidden = True     '缺省不显示已诊病人区域
    
    '放到医嘱后面，是因为医嘱中合理用药部件如果不存在，弹了提示框后，会导致排队叫号窗口禁用。
    If Check排队叫号 = True Then
        '检查是否存在排队叫号
        Set objPane = Me.dkpMain.CreatePane(6, 380, 550, DockTopOf, dkpMain.Panes(1))
        objPane.Title = "排队叫号"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable
        objPane.AttachTo dkpMain.Panes(1)
        'mobjQueue.zlSetToolIcon 24, True
    End If
    
    '界面恢复:放在最后执行
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        '会恢复Panne的标题,Tag被清除
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    End If
    
    '上面一句会清空句柄，所以在这句之后再绑定句柄
    For i = 1 To Me.dkpMain.PanesCount
        If Me.dkpMain.Panes(i).ID = 4 Then
            Me.dkpMain.Panes(i).Handle = lvwIncept.hwnd '进入时没有激活AttachPane事情,所以强行赋值
        ElseIf Me.dkpMain.Panes(i).ID = 5 Then
            Me.dkpMain.Panes(i).Handle = lvwReserve.hwnd '进入时没有激活AttachPane事情,所以强行赋值
        ElseIf Me.dkpMain.Panes(i).ID = pt排队叫号 Then
            Me.dkpMain.Panes(i).Handle = mobjQueue.zlGetForm.hwnd '进入时没有激活AttachPane事情,所以强行赋值
        ElseIf Me.dkpMain.Panes(i).ID = 7 Then  '回诊
            Me.dkpMain.Panes(i).Handle = lvwPatiHZ.hwnd
        ElseIf Me.dkpMain.Panes(i).ID = 3 Then  '已诊
            Me.dkpMain.Panes(i).Handle = picYZ.hwnd
        ElseIf Me.dkpMain.Panes(i).ID = 8 Then
            Me.dkpMain.Panes(i).Handle = rptNotify.hwnd
        Else
            Me.dkpMain.Panes(i).Handle = lvwPati(Me.dkpMain.Panes(i).ID - 1).hwnd
        End If
    Next
    
    dkpMain.Panes(1).Select
    
    '设置缺省查找方式
    arrType = Split(mstrCardKind, ";")
    For i = 1 To UBound(arrType) + 1
        If i = mintFindType Then
            PatiIdentify.objIDKind.IDKind = i
            Exit For
        End If
    Next
    
    
    '其它界面设置
    picPatiInput.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    Call zlControl.CboSetWidth(cboRegist.hwnd, cboRegist.Width * 1.1)
    Call zlControl.CboSetWidth(cboEdit(cbo职业).hwnd, cboRegist.Width * 3)
    
    Call RestoreWinState(Me, App.ProductName, , True)
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
    Call RefreshTitle
    If Check排队叫号 = True Then
       '检查是否存在排队叫号
       Call ReshDataQueue
       For i = 1 To dkpMain.PanesCount
           If dkpMain.Panes(i).Title Like "*排队叫号*" Then
               dkpMain.Panes(i).Select: Exit For
           End If
       Next
    End If
    Call RefreshPass

    If ISPassShowCard Then Call Hide就诊卡号列
    ucPatiVitalSigns.LabToTxt = -20
    ucPatiVitalSigns.XDis = 100
    mblnUnRefresh = False
End Sub

Private Sub RefreshPass()
    '是否调用太元通接口部件
    mblnUseTYT = False
    If gbytPass = 3 Then
        If gint过敏输入来源 = 0 Then
            mint过敏输入来源 = Val(zlDatabase.GetPara("过敏输入来源", glngSys, p门诊医生站, "0"))
        End If
        mblnUseTYT = gint过敏输入来源 = 0 And mint过敏输入来源 = 1 Or gint过敏输入来源 = 2
    End If
    '创建太元通接口对象，创建失败，则不启用太元通
    If gbytPass = 3 Then
        On Error Resume Next
    
        If gobjPass Is Nothing Then
            Set gobjPass = CreateObject("Midlayer.ComInterface")
        End If
        If err.Number <> 0 Then err.Clear: gbytPass = 0
        If gobjPass Is Nothing Then gbytPass = 0
        
        On Error GoTo 0
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim str挂号单 As String, strCardNO As String
    Dim rsTmp As Recordset
    Dim str疾病ID As String, str诊断ID As String
    Dim intFindTypeTmp As Integer
    Dim strPictureFile As String
    
    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
 
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                    objControl.Style = xtpButtonIcon
                Else
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_FontSize_S '小字体
        If mbytSize <> 0 Then
            mbytSize = 0
            Call zlDatabase.SetPara("字体", mbytSize, glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '大字体
        If mbytSize <> 1 Then
            mbytSize = 1
            Call zlDatabase.SetPara("字体", mbytSize, glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_Find '查找
        If Me.ActiveControl Is PatiIdentify Then
            PatiIdentify.SetFocus '有时需要定位一下
            If PatiIdentify.Text <> "" Then
                Call ExecuteFindPati
            End If
        Else
            PatiIdentify.SetFocus
        End If
    Case conMenu_View_FindNext '查找下一个
        If PatiIdentify.Text = "" And mstrIDCard = "" Then
            PatiIdentify.SetFocus
        Else
            Call ExecuteFindPati(True, IIf(PatiIdentify.Text = "", mstrIDCard, ""))
        End If
    Case 3564 '预约登记
        Call gobjVitualExpense.zlExecuteCommandBars(Me, Control, str挂号单, mlng病人ID)
    Case conMenu_Edit_AppRequest
        Call gobjVitualExpense.zlExecuteCommandBars(Me, Control, str挂号单, mlng病人ID)
    Case conMenu_Edit_OpenArrangement
        Call gobjVitualExpense.zlOpenStopedPlanBySN(Me, p门诊医生站, , , UserInfo.ID)
    Case conMenu_View_PatInfor  '显示病人详细信息
        mblnPatiDetail = Not mblnPatiDetail
        Call zlDatabase.SetPara("显示病人详细信息", IIf(mblnPatiDetail, 1, 0), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
        Call picExpand_Click
        
    Case conMenu_View_PatiInput  '显示快捷输入面板
        mblnDocInput = Not mblnDocInput
        Call zlDatabase.SetPara("显示病历快捷输入", IIf(mblnDocInput, 1, 0), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
        
        PicOutDoc.Visible = mblnDocInput
        Call cbsMain_Resize
        
        If mblnDocInput Then Call LoadDocData   '不显示时没有读取
        Call SetPermitEdit
        
        Call PicBasis_Resize
        Call PicPatiInfo_Resize
                
    Case conMenu_View_Busy '诊室状态
        Call SetRoomState(lblRoom.BackColor = COLOR_FREE)
    Case conMenu_View_Refresh '刷新
        Call LoadPatients("110111")
        Call LoadNotify
    Case conMenu_View_Jump '跳转
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_File_Parameter '参数设置
        frmOutStationSetup.mstrPrivs = mstrPrivs
        frmOutStationSetup.Show 1, Me

        If gblnOK Then
            intFindTypeTmp = mintFindType
            Call GetLocalSetting
            mintFindType = intFindTypeTmp
            Call LoadPatients
            Call InitQueuePara
        End If
        If Check排队叫号 Then
            Call ReshDataQueue
        End If
    Case conMenu_Tool_KssAudit '抗菌用药审核
        On Error Resume Next
        Call frmKSSExamine.Show(0, Me)
    Case conMenu_Tool_CISMed  '临床自管药
        Call Set临床自管药(Me)
     Case conMenu_Tool_TransAudit '输血审核管理
        On Error Resume Next
        Call frmTransfuseExamine.ShowMe(Me, 1)
    Case conMenu_Tool_Archive '电子病案查阅
        Call frmArchiveView.ShowArchive(Me, mPatiInfo.病人ID, mPatiInfo.挂号ID)
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '诊疗措施参考
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    Case conMenu_Manage_FeeItemSet  '诊疗项目费用设置
        Call Set诊疗项目费用设置
        
    Case conMenu_Tool_Community * 100# + 1 '社区身份验证
        Call ExecuteCommunityIdentify
    Case conMenu_Tool_Community * 100# + 2 To conMenu_Tool_Community * 100# + 99 '社区其他功能
        If Not gobjCommunity Is Nothing And mPatiInfo.社区 <> 0 And mPatiInfo.挂号ID <> 0 Then
            If gobjCommunity.CommunityFunc(glngSys, mlngModul, Val(Control.Parameter), mPatiInfo.社区, mPatiInfo.社区号, mPatiInfo.病人ID, mPatiInfo.挂号ID) Then
                Call LoadPatients
            End If
        End If
    Case conMenu_Tool_MedRec '门诊首页
        If mclsInOutMedRec Is Nothing Then
            Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
            Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, p门诊医生站, mclsMipModule, gobjCommunity, gclsInsure)
        End If
        If mclsInOutMedRec.ShowOutMedRecEdit(Me, mPatiInfo.挂号单, mstrPrivs, IIf(mstr挂号单 = mPatiInfo.挂号单 And (mPatiInfo.类型 = pt就诊 Or mPatiInfo.类型 = pt回诊), 0, 1), strPictureFile) Then
'            If strPictureFile <> "" And strPictureFile <> "0" Then
'                Call ReadPatPricture(mlng病人ID, imgPatient, strPictureFile)
'                picPatient.Visible = True
'            ElseIf strPictureFile = "" Then
'                picPatient.Visible = False
'            End If
'            Call LoadPatients("110")
        End If
        Call RefreshPass
    Case conMenu_File_MedRecSetup '首页打印设置
        'Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1260_2", Me)
    Case conMenu_File_MedRecPreview '首页预览
        'Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1260_2", Me, "病人ID=" & mlng病人ID, "NO=" & mPatiInfo.挂号单, 1)
    Case conMenu_File_MedRecPrint '首页打印
        'Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1260_2", Me, "病人ID=" & mlng病人ID, "NO=" & mPatiInfo.挂号单, 2)
    Case conMenu_Manage_Regist '病人挂号
        Control.Enabled = False
        Call ExecuteRegist
        Control.Enabled = True
    Case conMenu_Manage_Bespeak '预约挂号
        Control.Enabled = False
        Call ExecuteBespeak
        Control.Enabled = True
    Case conMenu_File_Print_Bespeak '重打预约挂号单
        Control.Enabled = False
        Call ExecuteBespeakPrint
        Control.Enabled = True
    Case conMenu_Manage_Transfer_Send '病人转诊
        Call ExecuteTransferSend
    Case conMenu_Manage_Transfer_Cancel '取消转诊
        Call ExecuteTransferCancel
    Case conMenu_Manage_Transfer_Incept '接收转诊
        Call ExecuteTransferIncept
    Case conMenu_Manage_Transfer_Refuse '转诊拒绝
        Call ExecuteTransferRefuse
    Case conMenu_Manage_Transfer_Force '强制续诊
        str挂号单 = frmForceGet.ShowMe(Me, mstrPrivs, mlng接诊科室ID, mobjSquareCard)
        If str挂号单 <> "" Then
            If lvwPati(pt就诊).Visible Then
                Call LoadPatients("110011", pt就诊, str挂号单)
                lvwPati(pt就诊).SetFocus
            Else
                Call LoadPatients("110011")
            End If
        End If
    Case conMenu_Manage_Receive '病人接诊
        Call ExecuteReceive
    Case conMenu_Manage_Cancel '取消接诊
        Call ExecuteCancel
    Case conMenu_Manage_Finish '完成接诊
        Call ExecuteFinish
    Case conMenu_Manage_Redo '恢复接诊
        Call ExecuteRedo
    Case conMenu_Manage_ReBack '暂停就诊
          Call ExecuteStopAndReuse(False)
    Case conMenu_Manage_ReBackCancel '恢复暂停就诊
          Call ExecuteStopAndReuse(True)
    Case conMenu_Edit_Transf_Save   '保存病人信息
        Call ExecuteOK
        Call HideWordInput
        Call PicBasis_Resize
    Case conMenu_Edit_Transf_Cancle '取消病人信息
        Call ExecutePaitCancel
        Call HideWordInput
        Call PicBasis_Resize
   Case conmenu_View_Leave  '显示不就诊病人
         mblnShowLeavePati = Not mblnShowLeavePati
         Control.Checked = mblnShowLeavePati
        Call LoadPatients("10000")
    Case conmenu_Edit_Leave     '病人不就诊
        If Set病人挂号状态(-1) Then
            Call LoadPatients("10000")
            Call ReshDataQueue
        End If
    Case conmenu_Edit_Wait      '病人就诊
        If Set病人挂号状态(0) Then
            Call LoadPatients("10000")
            Call ReshDataQueue
        End If
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '退出
        Unload Me
    Case conMenu_Tool_HealthCard  '居民健康卡
        If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.zlHealthArchivesShow(Me, p门诊医生站, mlng病人ID, "")
        End If
    Case conMenu_Edit_TraReactionRecord '输血反应
        Call FuncTraReactionRecord(Me, 0, p门诊医嘱下达)
    Case conMenu_Tool_Positive '阳性结果查看
        i = GetOne阳性结果
        If i <> 0 Then Call mclsDisease.ShowDisRegist(Me, 1, i, mlng病人ID, 0, mstr挂号单)
    Case conMenu_Tool_Critical '危急值查看处理
        Call ExecuteCritical
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            With mPatiInfo
                If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_3" Then
                    If mlng接诊科室ID = 0 Then
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
                    Else
                        Set rsTmp = zlDatabase.OpenSQLRecord("Select 名称 From 部门表 Where ID=[1]", Me.Caption, mlng接诊科室ID)
                        If rsTmp.EOF Then Exit Sub
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "开嘱科室=" & rsTmp!名称 & "|=" & mlng接诊科室ID)
                    End If
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                        "病人ID=" & .病人ID, "门诊号=" & .门诊号, "挂号单=" & .挂号单, "诊室=" & .诊室)
                End If
            End With
        Else
            If Check排队叫号 = True Then
                mobjQueue.zlExecuteCommandBars Control
            End If
            Select Case Me.tbcSub.Selected.Tag
            Case "医嘱"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "病历"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "新病历"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, p门诊医生站, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mPatiInfo.病人ID, 0, mPatiInfo.挂号单)
                    Call zlPlugInErrH(err, "ExeButtomClick")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End If
    End Select
End Sub


Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl
    Dim strFunc As String, arrFunc As Variant
    Dim i As Long
    Dim arrKind() As String
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID

    Case conMenu_Manage_Transfer
        With CommandBar.Controls
            If .Count = 0 Then
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Send, "转诊病人(&S)", -1, False)
                objControl.IconId = conMenu_Manage_Transfer
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Cancel, "取消转诊(&C)", -1, False)
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Incept, "转诊接收(&I)", -1, False)
                objControl.IconId = conMenu_Manage_Receive
                objControl.BeginGroup = True
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Refuse, "转诊拒绝(&R)", -1, False)
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "强制续诊(&F)", -1, False)
                objControl.BeginGroup = True
            End If
        End With
    Case conMenu_Tool_Community '社区功能
        mlngCommunityID = 0
        With CommandBar.Controls
            .DeleteAll
            If Not gobjCommunity Is Nothing Then
                '补充验证
                If mPatiInfo.社区 = 0 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Tool_Community * 100# + 1, "身份验证(&V)")
                End If
                
                '其他功能
                If mPatiInfo.社区 <> 0 Then
                    strFunc = gobjCommunity.GetCommunityFunc(glngSys, p门诊医生站, mPatiInfo.社区)
                    If strFunc <> "" Then
                        arrFunc = Split(strFunc, ";")
                        For i = 0 To UBound(arrFunc)
                            Set objControl = .Add(xtpControlButton, conMenu_Tool_Community * 100# + i + 2, Split(arrFunc(i), ",")(1))
                            If i < 9 Then objControl.Caption = objControl.Caption & "(&" & i + 1 & ")"
                            
                            If UCase(arrFunc(i)) Like UCase("Auto:*") Then
                                objControl.Parameter = Mid(Split(arrFunc(i), ",")(0), 6)
                                mlngCommunityID = objControl.ID
                            Else
                                objControl.Parameter = Split(arrFunc(i), ",")(0)
                            End If
                            objControl.ToolTipText = Split(arrFunc(i), ",")(2)
                        Next
                    End If
                End If
            End If
        End With
    Case Else
       Select Case tbcSub.Selected.Tag
       Case "医嘱"
           Call mclsAdvices.zlPopupCommandBars(CommandBar)
       Case "病历"
       End Select
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    If Control.Enabled = False And mblnPatiEditable And mblnPatiChange = False Then
        If Not (Control.ID = conMenu_Edit_Transf_Save Or Control.ID = conMenu_Edit_Transf_Cancle Or Control.ID \ 10 = conMenu_Manage_ShowAller) Then Control.Enabled = True
    End If
        
    Select Case Control.ID
    Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
        Control.Enabled = mblnPatiEditable And mblnPatiChange
        Control.Visible = mblnPatiEditable And mblnPatiChange
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case 3564
        Control.Enabled = InStr(mstrVirutalPrivs, ";预约登记;") > 0
    Case conMenu_Edit_AppRequest
        Control.Enabled = InStr(mstrVirutalPrivs, ";预约登记;") > 0
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_FontSize_S '小字体
        Control.Checked = Not (mbytSize = 1)
    Case conMenu_View_FontSize_L '大字体
        Control.Checked = (mbytSize = 1)
    Case conMenu_View_PatiInput '显示快捷输入面板
        Control.Checked = mblnDocInput
    Case conMenu_View_PatInfor      '显示病人详细信息
        Control.Checked = mblnPatiDetail
        
    Case conMenu_View_Busy '诊室状态
        Control.Checked = lblRoom.BackColor = COLOR_BUSY
    Case conMenu_Tool_KssAudit  '抗菌用药审核
        If GetInsidePrivs(p抗菌用药审核) = "" Then
            Control.Visible = False
        End If
    Case conMenu_Tool_TransAudit '输血分级管理
        If GetInsidePrivs(p输血审核管理) = "" Or Not gbln输血分级管理 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_CISMed  '临床自管药
        If InStr(GetInsidePrivs(p门诊医生站), ";临床自管药;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_Archive '电子病案查阅
        If GetInsidePrivs(p电子病案查阅) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng病人ID <> 0
        End If
    Case conMenu_Tool_HealthCard  '居民健康卡
        Control.Enabled = mlng病人ID <> 0
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        If GetInsidePrivs(p疾病诊断参考) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 '药品及诊疗参考
        If GetInsidePrivs(p药品诊疗参考) = "" Then Control.Visible = False
    Case conMenu_Tool_Community '社区菜单
        If gobjCommunity Is Nothing Then
            Control.Visible = False
        End If
    Case conMenu_Edit_TraReactionRecord '输血反应
        Control.Visible = InStr(1, GetInsidePrivs(9005, , 2200), "输血反应登记") <> 0
        Control.Enabled = Control.Visible And gbln血库系统
    Case conMenu_Manage_FeeItemSet '诊疗项目费用设置,没有权限时可查看
                
    Case conMenu_Tool_Community * 100# + 1 '社区身份验证
        Control.Enabled = mlng病人ID <> 0 And mPatiInfo.社区 = 0 And (mPatiInfo.类型 = pt就诊 Or mPatiInfo.类型 = pt回诊) And InStr(mstrPrivs, "病人接诊") > 0
    Case conMenu_Tool_Community * 100# + 2 To conMenu_Tool_Community * 100# + 99 '社区其他功能
        Control.Enabled = mlng病人ID <> 0 And mPatiInfo.社区 <> 0
    Case conMenu_Tool_MedRec '门诊首页
        If InStr(mstrPrivs, "门诊首页") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng病人ID <> 0
        End If
    Case conMenu_File_MedRec '首页打印
        If InStr(mstrPrivs, "打印首页") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng病人ID <> 0
        End If
    
    Case conMenu_ManagePopup '“接诊”菜单
        If InStr(mstrPrivs, ";病人接诊;") = 0 Then Control.Visible = False
    Case conMenu_Manage_Regist '病人挂号
        If InStr(mstrPrivs, ";病人挂号;") = 0 Then Control.Visible = False
    Case conMenu_Manage_Bespeak '预约挂号
        If InStr(mstrPrivs, ";预约挂号;") = 0 Then Control.Visible = False
    Case conMenu_Edit_OpenArrangement
        If InStr(mstrPrivs, ";预约挂号;") = 0 And InStr(mstrPrivs, ";病人挂号;") = 0 And InStr(mstrVirutalPrivs, ";预约登记;") = 0 Then Control.Visible = False
    Case conMenu_File_Print_Bespeak
      Control.Visible = InStr(mstrPrivs, ";预约挂号单;") > 0 And lvwReserve.Visible     '56274
      Control.Enabled = lvwReserve.Visible And Not lvwReserve.SelectedItem Is Nothing
    Case conMenu_Manage_Transfer '转诊处理
        If InStr(mstrPrivs, "病人接诊") = 0 _
            And InStr(mstrPrivs, "病人转诊") = 0 _
                And InStr(mstrPrivs, "续诊病人") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Manage_Transfer_Send '病人转诊
        If InStr(mstrPrivs, "病人转诊") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = pt候诊 Or mintActive = pt就诊)
            If blnEnabled Then
                blnEnabled = Not lvwPati(mintActive).SelectedItem Is Nothing And lvwPati(mintActive).Visible
                If blnEnabled Then
                    '目前处于"无转诊/或已接收"状态
                    With lvwPati(mintActive).SelectedItem.ListSubItems(5)
                        blnEnabled = .Tag = "" Or Val(.Tag) = 1
                    End With
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Manage_Transfer_Cancel '取消转诊
        If InStr(mstrPrivs, "病人转诊") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = pt候诊 Or mintActive = pt就诊)
            If blnEnabled Then
                blnEnabled = Not lvwPati(mintActive).SelectedItem Is Nothing And lvwPati(mintActive).Visible
                If blnEnabled Then
                    '目前处于转诊"待接收/已拒绝"状态
                    With lvwPati(mintActive).SelectedItem.ListSubItems(5)
                        blnEnabled = Val(.Tag) = 0 And .Tag <> "" Or Val(.Tag) = -1
                    End With
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conmenu_View_Leave  '显示不就诊病人
            Control.Checked = mblnShowLeavePati
            'Control.Enabled = (mintActive = pt候诊)
    Case conmenu_Edit_Leave
            blnEnabled = (mintActive = pt候诊)
            If blnEnabled Then
                blnEnabled = Not lvwPati(mintActive).SelectedItem Is Nothing And lvwPati(mintActive).Visible
                If blnEnabled Then
                    '只有正常的候诊病人才可以取消就诊
                    With lvwPati(mintActive).SelectedItem.ListSubItems(9)
                        blnEnabled = Val(.Tag) = 0
                    End With
                End If
            End If
            Control.Enabled = blnEnabled
    Case conmenu_Edit_Wait
        blnEnabled = mintActive = pt候诊
        If blnEnabled Then
            blnEnabled = Not lvwPati(mintActive).SelectedItem Is Nothing And lvwPati(mintActive).Visible
            If blnEnabled Then
                '目前处于转诊"待接收/已拒绝"状态
                With lvwPati(mintActive).SelectedItem.ListSubItems(9)
                    blnEnabled = Val(.Tag) = -1
                End With
            End If
        End If
        Control.Enabled = blnEnabled
        
    Case conMenu_Manage_Transfer_Incept, conMenu_Manage_Transfer_Refuse '转诊接收,转诊拒绝
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            '转诊列表有病人，且处于"待接收"状态
            blnEnabled = Not lvwIncept.SelectedItem Is Nothing And lvwIncept.Visible
            Control.Enabled = blnEnabled
        End If
        
    Case conMenu_Manage_Transfer_Force '强制续诊
        If InStr(mstrPrivs, "病人接诊") = 0 Or InStr(mstrPrivs, "续诊病人") = 0 Then Control.Visible = False
    Case conMenu_Manage_ReBack '暂停就诊:需回诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            If Not lvwPati(pt就诊).SelectedItem Is Nothing And mintActive = pt就诊 And lvwPati(pt就诊).Visible Then
                '0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
                Control.Enabled = Val(lvwPati(pt就诊).SelectedItem.ListSubItems(8).Tag) < 2
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Manage_ReBackCancel '恢复暂停就诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            If Not lvwPatiHZ.SelectedItem Is Nothing And lvwPatiHZ.Visible And mintActive = pt回诊 Then
                ' 0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
                Control.Enabled = Val(lvwPatiHZ.SelectedItem.ListSubItems(8).Tag) = 2
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Manage_Receive '病人接诊
        If InStr(mstrPrivs, "病人接诊") = 0 Or (mty_Queue.bln医生主动呼叫 And mbln呼叫后接诊) Then
            Control.Enabled = False
            Control.Visible = False
        Else
            Control.Visible = True
            '候诊，预约挂号病人可以直接接诊，转诊病人不通过这个功能
            blnEnabled = False
            If lvwPati(pt候诊).Visible And lvwReserve.Visible Then
                blnEnabled = mintActive = pt候诊 And Not lvwPati(pt候诊).SelectedItem Is Nothing And Me.ActiveControl Is lvwPati(pt候诊) _
                    Or Not lvwReserve.SelectedItem Is Nothing And Me.ActiveControl Is lvwReserve
            ElseIf lvwPati(pt候诊).Visible Then
                blnEnabled = mintActive = pt候诊 And Not lvwPati(pt候诊).SelectedItem Is Nothing
            ElseIf lvwReserve.Visible Then
                blnEnabled = mintActive = pt预约 And Not lvwReserve.SelectedItem Is Nothing
            End If
            Control.Enabled = blnEnabled    '不用再判断当前是否为转诊病人列表，因为如果是转诊列表的话，blnEnabled已经是False
             
        End If
    Case conMenu_Manage_Cancel '取消接诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mintActive = pt就诊 And Not lvwPati(pt就诊).SelectedItem Is Nothing And lvwPati(pt就诊).Visible
        End If
    Case conMenu_Manage_Finish '完成就诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        ElseIf mintActive = pt就诊 Then
            Control.Enabled = Not lvwPati(pt就诊).SelectedItem Is Nothing And lvwPati(pt就诊).Visible
        ElseIf mintActive = pt回诊 Then
            Control.Enabled = Not lvwPatiHZ.SelectedItem Is Nothing And lvwPatiHZ.Visible
        Else
            Control.Enabled = False
        End If
    Case conMenu_Manage_Redo '恢复接诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = mintActive = pt已诊 And Not lvwPati(pt已诊).SelectedItem Is Nothing And lvwPati(pt已诊).Visible
            If blnEnabled Then '只能恢复接诊自已的病人(否则有权限可用强制续诊)
                blnEnabled = lvwPati(pt已诊).SelectedItem.ListSubItems(4).Tag = UserInfo.姓名
            End If
            Control.Enabled = blnEnabled
        End If
    Case Else
        '60075:刘鹏飞,2013-04-03,将外部对医嘱打印、预览菜单的处理，移植到此处,以前的方式导致无法调用虚拟模块的更新事件
        If (Control.ID = conMenu_File_Print Or Control.ID = conMenu_File_Preview Or Control.ID = conMenu_Help_Help) Then
            If tbcSub.Selected.Tag = "医嘱" Then
                Control.Visible = False
                Exit Sub
            Else
                Control.Visible = True
            End If
        End If
        If Check排队叫号 Then mobjQueue.zlUpdateCommandBars Control
        Select Case tbcSub.Selected.Tag
        Case "医嘱"
            Call mclsAdvices.zlUpdateCommandBars(Control)
        Case "病历"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "新病历"
            Call mclsEMR.zlUpdateCommandBars(Control)
        End Select
        '抗菌用药报表
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_3" Then
                If gblnKSSStrict Then
                    Control.Visible = True
                Else
                    Control.Visible = False
                End If
            End If
        End If
    End Select
    
    If Control.Enabled And mblnPatiEditable And mblnPatiChange Then
        If Not (Control.ID = conMenu_Edit_Transf_Save Or Control.ID = conMenu_Edit_Transf_Cancle Or Control.ID \ 10 = conMenu_Manage_ShowAller) Then Control.Enabled = False
    End If
End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'功能：刷新子窗体菜单及工具条
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String
    
    '记录现有菜单样式
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        idx = GetFirstCommandBar(cbsMain(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsMain(2).Visible
            bytStyle = cbsMain(2).Controls(idx).Style
        End If
    End If
    
    '刷新子窗口菜单
    Call LockWindowUpdate(Me.hwnd)
        
    Me.Caption = "门诊医生工作站 - " & objItem.Caption & "(当前用户：" & UserInfo.姓名 & ")"
    
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '主窗口重新加入
    Call MainDefCommandBar
'    If Check排队叫号 And dkpMain.Panes(pt排队叫号).Selected Then
'        mobjQueue.zlDefCommandBars cbsMain
'    End If

    '子窗口重新加入
    Select Case objItem.Tag
    Case "医嘱"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 0, gobjPlugIn, mobjSquareCard)
    Case "病历"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain)
    Case "新病历"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, p门诊医生站, mcolSubForm("_" & objItem.Tag), objItem.Tag)
            Call zlPlugInErrH(err, "GetButtomName")
            '构建菜单
            If strName <> "" Then Call PlugInInSideBar(cbsMain, strName)
            err.Clear: On Error GoTo 0
        End If
    End Select
    
    
    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                objControl.Style = xtpButtonIcon
            Else
                objControl.Style = bytStyle
            End If
        Next
        cbsMain(lngCount).Visible = blnShowBar
    Next
    
    '如果用了RecalcLayout反而不正常
    Call LockWindowUpdate(0)
    
    Set mfrmActive = mcolSubForm("_" & objItem.Tag)
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'功能：刷新子窗体数据及状态
    If mlng病人ID = 0 Or (mintActive = pt候诊 And mPatiInfo.挂号单 = mstr挂号单) Then
        '候诊和预约病人，本次就诊没有医嘱和病历数据
        '要求子窗体按无数据处理界面
        Select Case objItem.Tag
        Case "医嘱"
            Call mclsAdvices.zlRefresh(0, "", False)
        Case "病历"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "新病历"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 1)
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, p门诊医生站, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        With mPatiInfo
            Select Case objItem.Tag
            Case "医嘱"
                Call mclsAdvices.zlRefresh(.病人ID, .挂号单, mstr挂号单 = .挂号单 And (.类型 = pt就诊 Or .类型 = pt回诊) And mlng病人ID <> 0, .数据转出, , , mclsMipModule)
            Case "病历"
                Call mclsEPRs.zlRefresh(.病人ID, .挂号ID, mlng科室ID, mlng科室ID = .科室ID And (.类型 = pt就诊 Or .类型 = pt回诊) And mlng病人ID <> 0, .数据转出, True)
            Case "新病历"
                Call mclsEMR.zlRefresh(.病人ID, .挂号ID, mlng科室ID, .类型, 1)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, p门诊医生站, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng病人ID, mstr挂号单, 0, .数据转出, 0, 0)
                    Call zlPlugInErrH(err, "RefreshForm")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End With
    End If
    Call SetFontSize(Not Me.Visible)
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim strFunName As String

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…") '固有
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_File_MedRec, "首页打印(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_File_MedRecSetup, "打印设置(&S)", -1, False
            .Add xtpControlButton, conMenu_File_MedRecPreview, "打印预览(&V)", -1, False
            .Add xtpControlButton, conMenu_File_MedRecPrint, "打印首页(&P)", -1, False
        End With
        '56274
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_Bespeak, "重打预约挂号单(&P)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "接诊(&C)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Regist, "病人挂号(&H)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "预约挂号(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AppRequest, "预约登记(&A)")
        Set objControl = .Add(xtpControlButton, 3564, "预约登记管理(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OpenArrangement, "开放停诊安排(&P)")
        Set objControl = .Add(xtpControlButton, conmenu_Edit_Leave, "病人不就诊(&L)", -1, False): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conmenu_Edit_Wait, "病人待诊(&W)", -1, False)
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Transfer, "转诊处理(&C)"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Receive, "病人接诊(&Z)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Cancel, "取消接诊(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Finish, "完成接诊(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Redo, "恢复接诊(&R)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBack, "需回诊(&S)"): objControl.BeginGroup = True
        objControl.IconId = conMenu_Edit_Pause
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBackCancel, "取消回诊(&R)")
        objControl.IconId = conMenu_Edit_Reuse
        
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "字体大小(&N)") '固有
        With objPopup.CommandBar.Controls
             .Add xtpControlButton, conMenu_View_FontSize_S, "小字体(&S)", -1, False '固有(小字体对应小卡片，大字体对应大卡片)
             .Add xtpControlButton, conMenu_View_FontSize_L, "大字体(&L)", -1, False '固有
        End With
        objPopup.BeginGroup = True

        If InStr(GetInsidePrivs(p门诊病历管理), "病历书写") > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_View_PatiInput, "显示病历快捷输入(&I)")
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_View_PatInfor, "显示病人详细信息(&D)")
        Set objControl = .Add(xtpControlButton, conmenu_View_Leave, "显示不就诊病人(&4)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Busy, "诊室忙(&M)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "窗格跳转(&J)")
        
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Community, "社区功能(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_KssAudit, "抗菌用药审核(&K)")
        objControl.IconId = 3551
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "输血审核管理(&M)")
        objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, conMenu_Tool_CISMed, "临床自管药(&J)")
        objControl.IconId = 3901
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)"): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "资料参考(&R)"): objPopup.BeginGroup = True
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "疾病诊断参考(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "诊疗措施参考(&C)", -1, False
        End With
        
        If gbln血库系统 = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReactionRecord, "输血反应记录"): objControl.BeginGroup = True
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Positive, "阳性结果")
            objControl.IconId = 3551
        If mbln危急值 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Critical, "危急值")
                objControl.IconId = 4113
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "诊疗项目费用设置(&C)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "门诊首页(&M)")
        objControl.BeginGroup = True
        objControl.ToolTipText = "编辑或查看首页信息"
        On Error Resume Next
            If mobjSquareCard.zlHealthArchiveIsSHow(Me, p门诊医生站, strFunName, "") Then
                If err.Number = 0 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Tool_HealthCard, strFunName)
                    objControl.BeginGroup = True
                    objControl.IconId = 3208
                Else
                    strFunName = ""
                End If
            Else
                strFunName = ""
            End If
        On Error GoTo 0
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)") '固有
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Regist, "挂号"): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlPopup, conMenu_Manage_Transfer, "转诊")
        objPopup.ID = conMenu_Manage_Transfer
        objPopup.IconId = conMenu_Manage_Transfer
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Receive, "接诊")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Finish, "完成"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBack, "需回诊")
        objControl.IconId = conMenu_Edit_Pause
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBackCancel, "取消")
        objControl.IconId = conMenu_Edit_Reuse
        objControl.ToolTipText = "取消回诊"
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "首页")
        objControl.BeginGroup = True
        objControl.ToolTipText = "编辑或查看首页信息"
        Set objPopup = .Add(xtpControlPopup, conMenu_Tool_Community, "社区")
        objPopup.ID = conMenu_Tool_Community
        objPopup.IconId = conMenu_Tool_Community
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "输血审核")
        objControl.IconId = 3551
                If strFunName <> "" Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_HealthCard, strFunName)
                objControl.ToolTipText = strFunName
                objControl.IconId = 3208
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存")
        objControl.BeginGroup = True
        objControl.Enabled = False
        objControl.ToolTipText = "保存病人相关信息的修改"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "放弃")
        objControl.Enabled = False
        objControl.ToolTipText = "放弃病人相关信息的修改（ESC）"
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助") '固有
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
    End With

    '查找项特殊处理
    '-----------------------------------------------------
    '主菜单右侧的查找
    With cbsMain.ActiveMenuBar.Controls
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = picTmphwnd.hwnd
    End With
    
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyH, conMenu_Manage_Regist '挂号
        .Add 0, vbKeyF7, conMenu_Manage_Receive '接诊
        .Add 0, vbKeyF8, conMenu_Manage_Finish '完成就诊
        .Add FCONTROL, vbKeyB, conMenu_View_Busy '诊室状态
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找病人
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        .Add 0, vbKeyF12, conMenu_File_Parameter '参数设置
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF6, conMenu_View_Jump '跳转
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save '保存
    End With
    
    '设置一些公共的不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '打印设置
'        .AddHiddenCommand conMenu_File_Excel '输出到Excel
'        .AddHiddenCommand conMenu_View_Jump '跳转
    End With
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1260_2")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = mblnPatiChange
End Sub

Private Sub InitRegist()
    '初始化挂号
    Set gobjVitualExpense = New clsRegist
    gobjVitualExpense.zlInitCommon glngSys, gcnOracle, gstrDBUser
    gobjVitualExpense.zlInitData 1
    mstrVirutalPrivs = GetPrivFunc(glngSys, 9000)
End Sub

Private Sub lbl医生_Click(Index As Integer)
    lbl病历名称.Visible = Not lbl病历名称.Visible
End Sub

Private Sub lvwIncept_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwIncept, ColumnHeader.Index)
End Sub
Private Sub lvwIncept_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LvwItemClick(pt转诊, Item)
End Sub

Private Sub lvwReserve_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwReserve, ColumnHeader.Index)
End Sub

Private Sub lvwReserve_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LvwItemClick(pt预约, Item)
End Sub

Private Sub mclsAdvices_Activate()
    mblnUnRefresh = False
End Sub

Private Sub mclsAdvices_CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String, ByRef blnNo As Boolean)
'功能：根据诊断与疾病编码 得到病历编辑器
'      blnOnChek    是否只进行传染病报告卡书写检查
'      str疾病ID    疾病ID
'      str诊断ID   诊断ID
'blnNO 是否要填写传染病报告卡
    Call OpenEPRDoc(mobjEPRDoc, Me, mPatiInfo.病人ID, mPatiInfo.挂号ID, mPatiInfo.科室ID, str疾病ID, str诊断ID, 1, , False, blnOnChek, blnNo)
End Sub

Private Sub mclsAdvices_EditDiagnose(ParentForm As Object, ByVal 挂号单 As String, Succeed As Boolean)
'功能：要求输入门诊诊断
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, p门诊医生站, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    If Not mclsInOutMedRec.ShowOutMedRecEdit(ParentForm, 挂号单, mstrPrivs) Then
        Succeed = False
    Else
        Succeed = mclsInOutMedRec.IsDiagInput
    End If
End Sub

Private Sub mclsAdvices_RequestRefresh()
'功能：医嘱子窗体要求刷新
    Call LoadPatients
End Sub

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'功能：医嘱子窗体要求更新状态栏
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsAdvices_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
End Sub

Private Sub mclsAdvices_PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean)
'功能：按编辑格式打印报告
    Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr诊疗报告, 报告ID, Not Preview, True)
End Sub

Private Sub mclsAdvices_ViewPACSImage(ByVal 医嘱ID As Long)
'功能：PACS观片处理
    If CreateObjectPacs(gobjPublicPacs) Then
        Call gobjPublicPacs.ShowImage(医嘱ID, Me, mPatiInfo.数据转出)
    End If
End Sub

Private Sub SetPermitEscape(ByVal blnOk As Boolean)
    Dim i As Long
        
    If blnOk Then
        'Call SetPermitEscape(true)之前必须先调SetPermitEdit
        mblnPatiChange = False
        mblnUnRefresh = False
        
        For i = 0 To lvwPati.Count - 1
            lvwPati(i).Enabled = True: lvwPati(i).BackColor = EColor
        Next
        lvwReserve.Enabled = True: lvwReserve.BackColor = EColor
        lvwIncept.Enabled = True: lvwIncept.BackColor = EColor
        lvwPatiHZ.Enabled = True: lvwPatiHZ.BackColor = EColor
        If Not mobjQueue Is Nothing Then mobjQueue.Enable = True
        
        PatiIdentify.Enabled = True
        cboRegist.Enabled = True
    Else
        If Visible And cboRegist.Tag = "" And mblnPatiChange = False Then
            mblnUnRefresh = True
            mblnPatiChange = True
            
            For i = 0 To lvwPati.Count - 1
                lvwPati(i).Enabled = False: lvwPati(i).BackColor = DColor
            Next
            lvwReserve.Enabled = False: lvwReserve.BackColor = DColor
            lvwIncept.Enabled = False: lvwIncept.BackColor = DColor
            lvwPatiHZ.Enabled = False: lvwPatiHZ.BackColor = DColor
            If Not mobjQueue Is Nothing Then mobjQueue.Enable = False
            
            PicBasis.BackColor = HColor
            picExpand.BackColor = HColor
            PicOutDoc.BackColor = HColor
            picPrompt.BackColor = HColor
                       
            If ucPatiVitalSigns.Enabled And ucPatiVitalSigns.ControlLock = False Then
                ucPatiVitalSigns.BackColor = HColor
                ucPatiVitalSigns.TextBackColor = HColor
                ucPatiVitalSigns.LblBackColor = HColor
            End If
            
            For i = 0 To txtEdit.UBound
                If txtEdit(i).Enabled And txtEdit(i).Locked = False Then
                    txtEdit(i).BackColor = HColor
                End If
            Next
            
            For i = 0 To cboEdit.UBound
                If i = 3 Then i = i + 2
                If cboEdit(i).Enabled Then
                    cboEdit(i).BackColor = HColor
                    fraLine(i).BackColor = HColor
                End If
            Next
            For i = 0 To optState.Count - 1
                If optState(i).Enabled Then optState(i).BackColor = HColor
            Next
            For i = 0 To UBound(mArrDate)
                mArrDate(i).BackColor = HColor
            Next
            vsAller.BackColor = HColor
            vsAller.BackColorBkg = HColor
            vsAller.CellBackColor = HColor
            vsAller.BackColorSel = EColor
            
            For i = 0 To rtfEdit.Count - 1
                If rtfEdit(i).Locked = False And rtfEdit(i).Visible Then
                    rtfEdit(i).BackColor = HColor
                End If
            Next
            
            PatiIdentify.Enabled = False
            cboRegist.Enabled = False
            Call PicBasis_Resize
        End If
    End If
End Sub

Private Sub SetPermitEdit()
    Dim i As Long, blnDo As Boolean, blnBasis As Boolean, blnDoc As Boolean
    Dim k As Long
            
    blnDo = mlng科室ID <> 0 And mlng科室ID = mPatiInfo.科室ID And InStr(mstrPrivs, "门诊首页") > 0 And (mintActive = pt就诊 Or mintActive = pt回诊)
    blnBasis = InStr(mstrPrivs, "修改基本信息") > 0
        
    ucPatiVitalSigns.ControlLock = Not blnDo
    
    If ucPatiVitalSigns.ControlLock = False Then
        ucPatiVitalSigns.BackColor = EColor
        ucPatiVitalSigns.TextBackColor = EColor
        ucPatiVitalSigns.LblBackColor = EColor
    Else
        ucPatiVitalSigns.BackColor = DColor
        ucPatiVitalSigns.TextBackColor = DColor
        ucPatiVitalSigns.LblBackColor = DColor
    End If
    
    For i = 0 To txtEdit.Count - 1
        If i = txt单位名称 Then
            txtEdit(i).Locked = Not blnDo
            If blnDo And Val("" & txtEdit(txt单位名称).Tag) <> 0 Then
                txtEdit(i).Locked = InStr(GetInsidePrivs(p门诊医生站), "合约病人登记") = 0
            End If
        Else
            txtEdit(i).Locked = Not blnDo
        End If
        If txtEdit(i).Locked = False Then
            txtEdit(i).BackColor = EColor
        Else
            txtEdit(i).BackColor = DColor
        End If
    Next
    
    For i = 0 To cboEdit.UBound
        If i = 3 Then i = i + 2
        cboEdit(i).Enabled = blnDo
        If blnDo Then
            cboEdit(i).BackColor = EColor
            fraLine(i).BackColor = EColor
        Else
            cboEdit(i).BackColor = DColor
            fraLine(i).BackColor = DColor
        End If
    Next
        
    For i = 0 To optState.Count - 1
        optState(i).Enabled = blnDo
        If blnDo Then
            optState(i).BackColor = EColor
        Else
            optState(i).BackColor = DColor
        End If
    Next
    
    For i = 0 To cmdEdit.Count - 1
        cmdEdit(i).Enabled = blnDo
        If i = cmd单位名称 Then
            cmdEdit(i).Enabled = Not txtEdit(txt单位名称).Locked
        End If
    Next
    
    If blnDo Then
        PicBasis.BackColor = EColor
        picExpand.BackColor = EColor
        vsAller.Editable = flexEDKbdMouse
        vsAller.BackColor = EColor
        vsAller.BackColorBkg = EColor
        vsAller.CellBackColor = EColor
        vsAller.BackColorSel = HColor
    Else
        PicBasis.BackColor = DColor
        picExpand.BackColor = DColor
        vsAller.Editable = flexEDNone
        vsAller.BackColor = DColor
        vsAller.BackColorBkg = DColor
        vsAller.CellBackColor = DColor
    End If
    
    For i = 0 To UBound(mArrDate)
        mArrDate(i).Enabled = blnDo
        mArrDate(i).BackColor = IIf(blnDo, EColor, DColor)
    Next
        
    If mblnDocInput Then
        blnDoc = mlng科室ID <> 0 And mlng科室ID = mPatiInfo.科室ID And _
                 (mPatiInfo.病历id = 0 And mPatiInfo.病历文件id <> 0 Or mPatiInfo.病历id <> 0 And mPatiInfo.是否签名 = False) And (mintActive = pt就诊 Or mintActive = pt回诊)
        If blnDoc And mPatiInfo.病历id <> 0 And lbl医生(1).Tag = "0" Then   '没有修改他人病历的权限
            blnDoc = mPatiInfo.保存人 = UserInfo.姓名
        End If
       
        k = 0
        For i = 0 To rtfEdit.Count - 1
            rtfEdit(i).Locked = Not blnDoc Or InStr(rtfEdit(i).Tag, ",") > 0   '存在多行内容时(进行了全文编辑保存)，不允许再修改
            If rtfEdit(i).Locked = False Then
                rtfEdit(i).BackColor = EColor
                k = k + 1
            Else
                rtfEdit(i).BackColor = DColor
            End If
        Next
        If k > 0 Then
            PicOutDoc.BackColor = EColor
            picPrompt.BackColor = EColor
        Else
            PicOutDoc.BackColor = DColor
            picPrompt.BackColor = DColor
        End If

        If mPatiInfo.病历id = 0 Or mPatiInfo.病历id <> 0 And mPatiInfo.是否签名 = False Then
            cmdSign.Caption = "签名(&S)"
        Else
            cmdSign.Caption = "取消签名(&S)"
        End If
        cmdSign.Enabled = mlng科室ID <> 0 And mlng科室ID = mPatiInfo.科室ID And (mPatiInfo.病历id = 0 And mPatiInfo.病历文件id <> 0 Or mPatiInfo.病历id <> 0) And (mintActive = pt就诊 Or mintActive = pt回诊)
        
        If cmdSign.Enabled And mPatiInfo.病历id <> 0 And lbl医生(1).Tag = "0" Then   '没有修改他人病历的权限
            cmdSign.Enabled = mPatiInfo.保存人 = UserInfo.姓名
        End If
        cmdUpdate.Enabled = cmdSign.Enabled
    End If
                
    mblnPatiEditable = blnDo Or blnDoc
    
    '病人基本信息：姓名，性别，年龄，出生日期不允许修改
    txtEdit(txt姓名).BackColor = &H8000000F: txtEdit(txt姓名).Locked = True: txtEdit(txt姓名).TabStop = False
    cboEdit(cbo性别).BackColor = &H8000000F: cboEdit(cbo性别).Locked = True: cboEdit(cbo性别).TabStop = False
    txt出生日期.BackColor = &H8000000F: txt出生日期.Enabled = False
    txt出生时间.BackColor = &H8000000F: txt出生时间.Enabled = False
    txtEdit(txt年龄).BackColor = &H8000000F: txtEdit(txt年龄).Locked = True: txtEdit(txt年龄).TabStop = False
    cboEdit(cbo年龄).BackColor = &H8000000F: cboEdit(cbo年龄).Locked = True: cboEdit(cbo年龄).TabStop = False

End Sub

Private Sub mclsEPRs_RequestRefresh()
    If mblnDocInput Then
        Call LoadDocData
        Call SetPermitEdit
        Call PicBasis_Resize
    End If
End Sub


Private Function CheckIsAskNextQueue(Optional str业务ID As String = "") As Boolean
   '------------------------------------------------------------------------------------------------------------------------
    '功能：检查医生是否允许呼叫下一个队列
    '编制：刘兴洪
    '返回:允许,返回true,否则返回False
    '日期：2010-06-09 16:48:30
    '说明：检查标准:以实际已呼叫为准(只有完成后，才能再叫)(问题:37442)
    '   取掉:候诊人数(不包含不就诊的)+已接诊的+转的<呼叫人数
    '------------------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem, lngCount As Long, rsTemp As ADODB.Recordset
    Dim strSQL As String, strLimit As String, strResult As String, arrCheck As Variant
    
    If Val(str业务ID) <> 0 Then
           strSQL = "Select Zl_QueuedateCheck([1]) as Chk From Dual"
           Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str业务ID))
           strResult = Nvl(rsTemp!chk) & "|"
           arrCheck = Split(strResult, "|")
           If Val(arrCheck(0)) <> 0 Then
              If Val(arrCheck(0)) = 1 Then
                If MsgBox(CStr(arrCheck(1)) & vbCrLf & "是否继续?", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                    Exit Function
                End If
              Else
                 MsgBox CStr(arrCheck(1)), vbCritical, Me.Caption
                 Exit Function
              End If
              
           End If
    End If
    
    
    If mty_Queue.bln医生主动呼叫 = False Or mty_Queue.int呼叫人数 <= 0 Then
        CheckIsAskNextQueue = True: Exit Function
    End If
    '0:排队中，1:呼叫中，2：已弃号，3：暂停，4：完成就诊，6：回诊，7：已呼叫
    'mty_Queue.bln呼叫含回诊
    
    '问题:44250
    strLimit = ",0,4," & IIf(mty_Queue.bln呼叫含回诊, "", ",6,")
    strSQL = "" & _
    "   Select Count(distinct B.ID) as Count From 病人挂号记录 B ,排队叫号队列 A" & _
    "   Where A.业务ID=B.ID And A.业务类型=0  " & _
    "               And instr([4],','||A.排队状态||',')=0   And B.记录性质=1 And B.记录状态=1" & _
    "               And A.医生姓名||''=[1]   " & IIf(mty_Queue.bln呼叫含回诊, " And nvl(A.回诊序号,0) = 0", "") & _
    "               And (  (nvl(B.急诊,0)=1  and B.发生时间>=Sysdate-[3] ) or   (nvl(B.急诊,0)<>1  and B.发生时间>=Sysdate-[2] )) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数), IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数), strLimit)
    lngCount = Val(Nvl(rsTemp!Count))

    If lngCount >= mty_Queue.int呼叫人数 Then
            MsgBox "最多只能有" & mty_Queue.int呼叫人数 & "个候诊病人,不能再进行呼叫！", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
    End If
    CheckIsAskNextQueue = True
End Function
 
Private Sub mclsInOutMedRec_Closed(ByVal blnEditCancel As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String, ByVal strTag As String)
'功能：调用事件
' strTag=附加信息，现在存储门诊病人照片文件的路径，以后扩展时，以|分割
    Dim strPictureFile As String, blnNo As Boolean
    If Not blnEditCancel Then
        If InStr(";" & GetPrivFunc(glngSys, p门诊病历管理) & ";", ";病历书写;") > 0 Then
            If mobjEPRDoc Is Nothing Then
                Set mobjEPRDoc = New zlRichEPR.cEPRDocument
            End If
            Call OpenEPRDoc(mobjEPRDoc, Me, mPatiInfo.病人ID, mPatiInfo.挂号ID, mPatiInfo.科室ID, str疾病ID, str诊断ID, 1, , False, , blnNo)
            If blnNo Then
                Call mclsDisease.EditNotFillReason(Me, mPatiInfo.病人ID, mPatiInfo.挂号ID, 1)
            End If
        End If
        strPictureFile = Trim(Split(strTag & "|", "|")(0))
         '由于门诊修改为非模显示，因此不能通过方法返回值来刷新界面，通过Closed时间刷新
        If strPictureFile <> "" And strPictureFile <> "0" Then
            Call ReadPatPricture(mPatiInfo.病人ID, imgPatient, strPictureFile)
            picPatient.Visible = True
        ElseIf strPictureFile = "" Then
            picPatient.Visible = False
        End If
        Call LoadPatients("110")
    End If
    Call RefreshPass
End Sub

Private Sub mobjEPRDoc_AfterSaved(lngRecordId As Long)
    Call LoadDocData
    With mPatiInfo
        Call mclsEPRs.zlRefresh(.病人ID, .挂号ID, mlng科室ID, mlng科室ID = .科室ID And (.类型 = pt就诊 Or .类型 = pt回诊) And mlng病人ID <> 0, .数据转出, True)
        Call SetPermitEdit
    End With
End Sub

Private Sub mobjQueue_OnQueueExecuteAfter(ByVal str业务ID As String, ByVal byt操作类型 As Byte)
    '------------------------------------------------------------------------------------------------------------------------
    '入参：byt操作类型-0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
    '------------------------------------------------------------------------------------------------------------------------
    If mty_Queue.bln医生主动呼叫 = False Then Exit Sub
    If byt操作类型 <> 1 Then Exit Sub

    
    '重新刷新病人信息
    Call LoadPatients("1000")
End Sub

Private Sub mobjQueue_OnQueueExecuteBefore(ByVal str业务ID As String, ByVal byt操作类型 As Byte, blnCancel As Boolean, strNewQueueName As String)
    Dim strSQL As String, rsTemp As ADODB.Recordset
   ' byt操作类型 -0 - 复诊, 1 - 直呼, 2 - 弃号, 3 - 暂停, 4 - 完成就诊, 5 - 广播
   
    If InStr(1, "15", byt操作类型) = 0 Then Exit Sub
    If CheckIsAskNextQueue(str业务ID) = False Then blnCancel = True: Exit Sub
    
    strSQL = "SELECT a.ID,a.No,a.病人ID,a.执行部门ID,A.执行状态 From 病人挂号记录 A,排队叫号队列 B  " & _
        "  where  a.ID=b.业务id and b.业务类型=0 and a.ID=[1] and nvl(b.排队状态,0)=0 And a.记录性质 in(1,2) And a.记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str业务ID))
    If rsTemp.EOF Then Exit Sub
    
    '68736:刘尔旋,2014-02-18,转诊病人没有诊室信息
    If byt操作类型 = 1 Then
        If Is转诊病人(str业务ID) Then
            If CheckTransferDetail(str业务ID) = False Then
                strSQL = "ZL_病人挂号记录_更新诊室 ('" & Nvl(rsTemp!NO) & "'," & Val(Nvl(rsTemp!病人ID)) & ",'" & mstr接诊诊室 & "','" & UserInfo.姓名 & "',to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
            Exit Sub
        End If
    End If
    
    If InStr(1, "12", Val(Nvl(rsTemp!执行状态))) > 0 Then
        '1-完成就诊,2-正在就诊:主要是第二次呼叫
        '应用于:如果已经分诊后,医生接诊后,叫病人去检查后,再复诊来呼叫
        Exit Sub
    End If
    
    '更新诊室_In Integer := 1
    strSQL = "ZL_病人挂号记录_更新诊室 ('" & Nvl(rsTemp!NO) & "'," & Val(Nvl(rsTemp!病人ID)) & ",'" & mstr接诊诊室 & "','" & UserInfo.姓名 & "',to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),0)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
End Sub

Private Function CheckTransferDetail(strID As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------
'功能:检查该转诊病人是否有诊室信息
'入参:strID-str业务ID
'返回:True 代表转诊病人有诊室信息 False 代表转诊病人无诊室信息
'编制:刘尔旋
'日期:2014-02-18
'备注:
'-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    
    strSQL = "Select 诊室 From 排队叫号队列 Where 业务Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID)
    '排队叫号队列没有记录,不更新
    If rsTemp.EOF Then CheckTransferDetail = True: Exit Function
    If Nvl(rsTemp!诊室) = "" Then CheckTransferDetail = False: Exit Function
    CheckTransferDetail = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Is转诊病人(str业务ID As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能:检查该病人是否是转诊病人并且未接收
    '入参:str业务ID
    '返回:True 代表为转诊病人 False 代表为普通病人
    '编制:王吉
    '编制日期:2012-9-14
    '问题号:51514
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    strSQL = _
    "   Select Count(ID) as 是否为转诊病人 From 病人挂号记录 Where ID=[1] And Nvl(转诊科室ID,0) <> 0 And Nvl(转诊状态,0)=0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str业务ID)
    If rsTemp.EOF Then Is转诊病人 = False
    Is转诊病人 = rsTemp!是否为转诊病人 > 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mobjQueue_OnRecevieDiagnose(ByVal str业务ID As String, ByVal lng业务类型 As Long)
    '接诊:
    Dim objControl As CommandBarControl
    Dim strNO As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim bln回诊 As Boolean, arrCheck As Variant, strResult As String
    Dim bln转诊病人 As Boolean '问题号:51514
    Dim datCurr As Date
    
    If lng业务类型 <> 0 Then Exit Sub
    On Error GoTo errH
     If Val(str业务ID) <> 0 Then
           strSQL = "Select Zl_QueuedateCheck([1]) as Chk From Dual"
           Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str业务ID))
           strResult = Nvl(rsTmp!chk) & "|"
           arrCheck = Split(strResult, "|")
           If Val(arrCheck(0)) <> 0 Then
              If Val(arrCheck(0)) = 1 Then
                If MsgBox(CStr(arrCheck(1)) & vbCrLf & "是否继续?", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                    Exit Sub
                End If
              Else
                 MsgBox CStr(arrCheck(1)), vbCritical, Me.Caption
                 Exit Sub
              End If
              
           End If
    End If
    strSQL = "Select 病人ID,执行人,NO,记录标志,执行状态,记录性质,姓名,门诊号,id as 挂号id,复诊,急诊 From 病人挂号记录 Where  ID=[1]  "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str业务ID)
    If rsTmp.EOF Then
        MsgBox "该病人没有挂号记录不能接诊。", vbInformation, gstrSysName
        Call LoadPatients("100001"): Exit Sub
    End If
    
    '问题号:57566
    If Check接诊控制("接诊", rsTmp!NO) = False Then Exit Sub
    
    '0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
    If Val(rsTmp!执行状态) = 1 Then
        MsgBox "该病人已经完成就诊,不能再进行就诊操作。", vbInformation, gstrSysName
        Call LoadPatients("100001"): Exit Sub
    ElseIf Val(rsTmp!执行状态) = -1 Then
        MsgBox "该病人已经标记为不就诊,不能再进行就诊操作。", vbInformation, gstrSysName
        Call LoadPatients("100001"): Exit Sub
    End If
    strNO = Nvl(rsTmp!NO)
    
    '转诊接收 问题号:51514
    bln转诊病人 = Is转诊病人(str业务ID)
    If bln转诊病人 Then
        strSQL = "Zl_病人挂号记录_转诊('" & strNO & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        '刷新并定位病人
        If lvwPati(pt就诊).Visible Then
            Call LoadPatients("11011", pt就诊, strNO)
        Else
            Call LoadPatients("11011")
        End If
    End If
    
    '接收预约挂号单
    datCurr = zlDatabase.Currentdate
    If Val("" & rsTmp!记录性质) = 2 Then
        If Val(zlDatabase.GetPara("允许挂号划价单", glngSys, p门诊医生站, 1)) <> 1 And Not mobjSquareCard Is Nothing Then
            If Not mobjSquareCard.zlRegisterIncept(Me, mlngModul, strNO, mstr接诊诊室, 0, "") Then Exit Sub
        Else
            strSQL = "Zl_病人预约挂号_接收('" & strNO & "','" & mstr接诊诊室 & "',NULL,NULL,NULL,NULL,NULL,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Else
        If Val(Nvl(rsTmp!执行状态)) = 0 Then
            '正常挂号接诊
            strSQL = "zl_病人接诊(" & Val(Nvl(rsTmp!病人ID)) & ",'" & strNO & "',Null,'" & UserInfo.姓名 & "','" & mstr接诊诊室 & "',0,0,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Else
            'Zl_病人接诊
            strSQL = "Zl_病人接诊("
            '  病人id_In     病人信息.病人id%Type,
            strSQL = strSQL & "" & Val(Nvl(rsTmp!病人ID)) & ","
            '  No_In         病人挂号记录.NO%Type,
            strSQL = strSQL & "'" & strNO & "',"
            '  执行部门id_In 病人挂号记录.执行部门id%Type,
            strSQL = strSQL & "" & IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID) & ","
            '  执行人_In     病人挂号记录.执行人%Type,
            strSQL = strSQL & "'" & IIf(mstr接诊医生 = "", UserInfo.姓名, mstr接诊医生) & "',"
            '  诊室_In       病人挂号记录.诊室%Type := Null,
            strSQL = strSQL & "'" & mstr接诊诊室 & "',"
            '  标记急诊_In   病人挂号记录.急诊%Type := 0,
            strSQL = strSQL & "0,"
            '  回诊_In Integer:=0
            strSQL = strSQL & "1,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            bln回诊 = True
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
        
    mstr挂号单 = strNO
    mlng病人ID = Val(Nvl(rsTmp!病人ID))
    
    '门诊患者接诊消息发送
    Call ZLHIS_CIS_009(mclsMipModule, mlng病人ID, Nvl(rsTmp!姓名), Nvl(rsTmp!门诊号), 0, 0, Nvl(rsTmp!挂号ID), Nvl(rsTmp!复诊, 0), Nvl(rsTmp!急诊, 0), datCurr, mlng接诊科室ID, , mstr接诊诊室, UserInfo.姓名)
    
    '刷新并定位病人
    On Error GoTo 0
    If lvwPati(pt就诊).Visible Then
        Call LoadPatients("110001", pt就诊, strNO)
        lvwPati(pt就诊).SetFocus
    Else
        Call LoadPatients("110001")
    End If
    '社区病人自动调用功能
    If Not gobjCommunity Is Nothing And mlngCommunityID <> 0 And mlng病人ID <> 0 And mPatiInfo.社区 <> 0 Then
        Set objControl = cbsMain.FindControl(, mlngCommunityID, , True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
    
    Call CreatePlugInOK(p门诊医生站)
    '接诊后调用外挂接口
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicReceive(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID)
        Call zlPlugInErrH(err, "ClinicReceive")
        err.Clear: On Error GoTo errH
    End If
    
    '接诊之后自动进行医嘱下达状态
    If mlng自动进行 = 1 And bln回诊 = False Then
        If tbcSub.Selected.Tag <> "医嘱" Then tbcSub.Item(0).Selected = True
        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    ElseIf mlng自动进行 = 2 And bln回诊 = False Then
        If tbcSub.Selected.Tag <> "病历" Then tbcSub.Item(1).Selected = True
        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
        mblnUnRefresh = True
        Call mclsEPRs.zlOpenDefaultEPR(mstr挂号单)
    End If
    '处理排队叫号队列(重新刷新)
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mobjQueue_OnSelectionChanged(ByVal blnIsCallingList As Boolean, objReportRow As Object, cbrMain As Object)
    If mty_Queue.bln医生主动呼叫 Then
        mobjQueue.zlCommandBarSet 7, blnIsCallingList Or Not mbln呼叫后接诊
    End If
     
End Sub

Private Sub optState_Click(Index As Integer)
    Call SetPermitEscape(False)
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.病人ID
    End If
    
    Call ExecuteFindPati(False, , blnCard, lngPatiID)
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit = True Then mintFindType = Index: mstrFindType = objCard.名称
End Sub

Private Sub PicBasis_Resize()
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    PicBasis.Cls
    For i = 0 To txtEdit.UBound
        x1 = txtEdit(i).Left
        y1 = txtEdit(i).Top + txtEdit(i).Height
        x2 = txtEdit(i).Left + txtEdit(i).Width
        y2 = y1
        PicBasis.Line (x1, y1)-(x2, y2)
    Next
    For i = 0 To cboEdit.UBound
        If i = 3 Then i = i + 2
        x1 = fraLine(i).Left
        y1 = fraLine(i).Top + fraLine(i).Height
        x2 = fraLine(i).Left + fraLine(i).Width
        y2 = y1
        PicBasis.Line (x1, y1)-(x2, y2)
    Next
    
    For i = 0 To UBound(mArrDate)
        If mArrDate(i).Text <> "____-__-__" And mArrDate(i).Text <> "__:__" Then
            x1 = mArrDate(i).Left
            y1 = mArrDate(i).Top + mArrDate(i).Height
            x2 = mArrDate(i).Left + mArrDate(i).Width
            y2 = y1
            PicBasis.Line (x1, y1)-(x2, y2)
        End If
    Next
        
    x1 = vsAller.Left
    y1 = vsAller.Top + vsAller.Height + IIf(mbytSize = 0, 0, 75)
    x2 = vsAller.Left + vsAller.Width
    y2 = y1
    PicBasis.Line (x1, y1)-(x2, y2)
End Sub


Private Sub picExpand_Click()
    If picExpand.Picture Is ilexpand.ListImages("展开").Picture Then
        Set picExpand.Picture = ilexpand.ListImages("折叠").Picture
        mblnPatiDetail = True
    Else
        Set picExpand.Picture = ilexpand.ListImages("展开").Picture
        mblnPatiDetail = False
    End If
    Call zlDatabase.SetPara("显示病人详细信息", IIf(mblnPatiDetail, 1, 0), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
    Call cbsMain_Resize
End Sub

Private Sub PicPatiInfo_Resize()
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    PicPatiInfo.Cls
    For i = 0 To lblShow.UBound
        If lblShow(i).Visible Then
            x1 = lblShow(i).Left
            y1 = lblShow(i).Top + lblShow(i).Height
            x2 = lblShow(i).Left + lblShow(i).Width
            y2 = y1
            PicPatiInfo.Line (x1, y1)-(x2, y2)
        End If
    Next
    
    x1 = lblDiag(1).Left
    y1 = lblDiag(1).Top + lblDiag(1).Height
    x2 = lblDiag(1).Left + lblDiag(1).Width
    y2 = y1
    PicPatiInfo.Line (x1, y1)-(x2, y2)
End Sub



Private Sub picYZ_Resize()
    On Error Resume Next
        lbl就诊时间.Left = 100
    cboSelectTime.Left = lbl就诊时间.Left + lbl就诊时间.Width + 15
    cmdOtherFilter.Left = cboSelectTime.Left + cboSelectTime.Width + 50
    lvwPati(pt已诊).Top = cboSelectTime.Top + cboSelectTime.Height + 30
    lvwPati(pt已诊).Width = picYZ.Width
    lvwPati(pt已诊).Height = picYZ.Height - lvwPati(pt已诊).Top
End Sub

Private Sub rtfEdit_Change(Index As Integer)
    If mblnSizeTmp = True Then Exit Sub
    If picPatiInput.Tag = "" Then
        Call SetPermitEscape(False)
        
        If cboRegist.Tag = "" And PicOutDoc.Tag <> "2" Then PicOutDoc.Tag = "2"
    Else
        picPatiInput.Tag = ""
    End If
End Sub

Private Sub rtfEdit_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub rtfEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not mblnPatiChange Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '连续两次回车光标跳转
            With rtfEdit(Index)
                If Trim(.Text) = "" Then
                    KeyAscii = 0
                    Call zlCommFun.PressKey(vbKeyTab)
                ElseIf .SelStart - 1 > 0 Then
                    If Mid(.Text, .SelStart - 1, 2) = vbCrLf Then
                        KeyAscii = 0
                        Call zlCommFun.PressKey(vbKeyBack)
                        Call zlCommFun.PressKey(vbKeyTab)
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub rtfEdit_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme
End Sub

Private Sub rtfEdit_SelChange(Index As Integer)
    With rtfEdit(Index)
        If .SelLength = 0 And .SelStart > 0 And picPatiInput.Tag = "" Then
            If Mid(.Text, .SelStart, 1) = "`" Or Mid(.Text, .SelStart, 1) = "·" Then
                picPatiInput.Tag = "UnChange"
                .SelStart = .SelStart - 1
                .SelLength = 1
                .SelText = ""
                Call ShowWordInput(rtfEdit(Index))
                picPatiInput.Tag = ""
            End If
        End If
    End With
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme
End Sub

Private Sub txtSentence_GotFocus()
    Call zlCommFun.OpenIme(True)
    Call zlControl.TxtSelAll(txtSentence)
End Sub

Private Sub txtSentence_KeyPress(KeyAscii As Integer)
    Dim strSentence As String, blnCancel As Boolean, strType As String
       
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        Select Case Val(picSentence.Tag)
        Case txt主诉
            strType = "病人主诉"
        Case txt家族史
            strType = "家族史"
        Case txt现病史
            strType = "现病史"
        Case txt查体
            strType = "体格一般检查"
        Case txt过去史
            strType = "既往史"
        End Select
                
        strSentence = frmSentenceSel.ShowMe(Me, mPatiInfo.病历文件id, mPatiInfo.性别, mPatiInfo.婚姻状况, strType, txtSentence.Text, picSentence.hwnd, blnCancel)
        If strSentence <> "" Then
            rtfEdit(Val(picSentence.Tag)).SelText = strSentence
            Call HideWordInput
        Else
            If Not blnCancel Then
                MsgBox "没有找到匹配的词句。", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txtSentence)
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call imgSentence_Click
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
        Call HideWordInput
    End If
End Sub


Private Sub imgSentence_Click()
    Dim strSentence As String, strType As String
    
    Select Case Val(picSentence.Tag)
    Case txt主诉
        strType = "病人主诉"
    Case txt家族史
        strType = "家族史"
    Case txt现病史
        strType = "现病史"
    Case txt查体
        strType = "体格一般检查"
    Case txt过去史
        strType = "既往史"
    End Select
    
    strSentence = frmSentenceSel.ShowMe(Me, mPatiInfo.病历文件id, mPatiInfo.性别, mPatiInfo.婚姻状况, strType)
    If strSentence <> "" Then
        rtfEdit(Val(picSentence.Tag)).SelText = strSentence
        Call HideWordInput
    End If
End Sub

Private Sub txtSentence_LostFocus()
    If Not frmSentenceSel.mblnShow Then
        Call HideWordInput   '隐藏词句输入
    End If
End Sub


Private Sub ShowWordInput(ByRef txtThis As RichTextBox)
'功能：显示词句输入
    Dim vPos As POINTAPI
    
    If txtThis.Visible And txtThis.Enabled And Not txtThis.Locked Then
        picSentence.Tag = txtThis.Index '记下以便隐藏返回后定位
        
        If txtThis.Text = "" Then picPatiInput.Tag = "UnChange": txtThis.Text = " " '必须要有一个空字符才能返回其坐标
        vPos = GetCaretPos(txtThis.hwnd)
        If txtThis.Text = " " Then picPatiInput.Tag = "UnChange": txtThis.Text = ""
        
        If vPos.x <> -1 And vPos.y <> -1 Then
            If txtThis.Left + vPos.x + Screen.TwipsPerPixelX * 2 < txtThis.Left + txtThis.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX Then
                picSentence.Left = txtThis.Left + vPos.x + Screen.TwipsPerPixelX * 2
            Else
                picSentence.Left = txtThis.Left + txtThis.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX
            End If
            picSentence.Top = txtThis.Top + vPos.y + Screen.TwipsPerPixelY
            txtSentence.Text = ""
            picSentence.Visible = True
            txtSentence.SetFocus
        End If
    End If
End Sub


Private Sub HideWordInput()
'功能：隐藏词句输入
    Dim idx As Long
    
    If picSentence.Visible Then
        picSentence.Visible = False
        txtSentence.Text = ""
        
        idx = Val(picSentence.Tag)
        picSentence.Tag = ""
        
        If rtfEdit(idx).Visible And rtfEdit(idx).Enabled And Not rtfEdit(idx).Locked Then
            rtfEdit(idx).SetFocus
        End If
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    Dim lngPos As Long, lngLen As Long
    If Index = txt身份证号 Then
        If txtEdit(txt身份证号).Tag <> "不触发Change事件" Then
            lngPos = InStr(txtEdit(txt身份证号).Text, "*")
            lngLen = Len(Mid(txtEdit(txt身份证号).Text, 13, 2))
            Select Case lngPos
                Case 0
                    txtEdit(txt身份证号).Tag = txtEdit(txt身份证号).Text
                Case Else
                    txtEdit(txt身份证号).Tag = Mid(txtEdit(txt身份证号).Text, 1, lngPos - 1)
                    txtEdit(txt身份证号).Text = txtEdit(txt身份证号).Tag
                    txtEdit(txt身份证号).SelStart = Len(txtEdit(txt身份证号).Text)
            End Select
        End If
    End If
    Call SetPermitEscape(False)
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtEdit(Index))
    
    Select Case Index
        Case txt单位名称, txt家庭地址, txt监护人, txt就诊摘要
            Call zlCommFun.OpenIme(True)
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = txt家庭地址) And txtEdit(Index).Text <> "" Then
            '输入地区数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 地区 " & _
                " Where (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "地区", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!名称
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt单位名称 And txtEdit(Index).Text <> "" Then
            '输入工作单位
            strSQL = "Select ID,编码,名称,简码,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " And (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtEdit(Index).Container.hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "工作单位", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", gstrLike & UCase(txtEdit(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                If InStr(GetInsidePrivs(p门诊医生站), "合约病人登记") > 0 Then txtEdit(Index).Tag = Val(rsTmp!ID)
                If txtEdit(txt单位电话).Text = "" Then
                    txtEdit(txt单位电话).Text = Nvl(rsTmp!电话)
                End If
            Else
                txtEdit(Index).Tag = ""
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        '非控制按键
        
        '选择快捷键
        If KeyAscii = Asc("*") Then
            '注意界面上要求CMD和对应TXT的Index相同
            If Index = txt家庭地址 Then
                KeyAscii = 0
                Call cmdEdit_Click(cmd家庭地址)
                Exit Sub
            ElseIf Index = txt单位名称 Then
                KeyAscii = 0
                Call cmdEdit_Click(cmd单位名称)
                Exit Sub
            End If
        End If
        
        '限制输入长度
        If txtEdit(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtEdit(Index).Text) > txtEdit(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '限制输入内容
        Select Case Index
'            Case txt年龄 '允许自由录入了
'                strMask = "1234567890"
            'Case txt出生日期 'MaskEdit限制了
                'strMask = "1234567890-"
            Case txt身份证号
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Case txt家庭电话, txt单位电话
                strMask = "1234567890-()"
            Case txt手机号
                strMask = "1234567890"
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Dim datCur As Date, datRes As Date
    Select Case Index
    
        Case txt发病
            If Trim(txtEdit(txt发病).Text) <> "" Then
                If IsNumeric(txtEdit(txt发病).Text) Then
                    If Val(txtEdit(txt发病).Text) <= 0 Then
                        MsgBox "发病时间推算值必须为正数。", vbInformation, gstrSysName
                        txtEdit(txt发病).SetFocus: Exit Sub
                    End If
                Else
                    MsgBox "发病时间推算值必须为数字。", vbInformation, gstrSysName
                    txtEdit(txt发病).Text = "": txtEdit(txt发病).SetFocus: Exit Sub
                End If
            Else
                 Exit Sub
            End If
            If cboEdit(cbo发病时间).ListIndex <= 0 Then Exit Sub
            datCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            Select Case cboEdit(cbo发病时间).ListIndex
                Case 1 '小时
                    datRes = DateAdd("n", -1 * Val(txtEdit(txt发病).Text) * 60, CDate(datCur))
                Case 2 '天
                    datRes = DateAdd("h", -1 * Val(txtEdit(txt发病).Text) * 24, CDate(datCur))
                Case 3 '周
                    datRes = DateAdd("d", -1 * 7 * Val(txtEdit(txt发病).Text), CDate(datCur))
                Case 4 '月
                    datRes = DateAdd("M", -1 * Int(Val(txtEdit(txt发病).Text)), CDate(datCur))
                    datRes = DateAdd("d", -1 * (Val(txtEdit(txt发病).Text) - Int(Val(txtEdit(txt发病).Text))) * 30, datRes)
                Case 5 '年
                    If Val(txtEdit(txt发病).Text) < 100 Then
                        datRes = DateAdd("yyyy", -1 * Int(Val(txtEdit(txt发病).Text)), CDate(datCur))
                        datRes = DateAdd("d", -1 * (Val(txtEdit(txt发病).Text) - Int(Val(txtEdit(txt发病).Text))) * 365, datRes)
                    Else
                        MsgBox "发病时间推算不能超过100年。", vbInformation, gstrSysName
                        txtEdit(txt发病).SetFocus: Exit Sub
                    End If
            End Select
            txt发病日期.Text = Format(CDate(datRes), "YYYY-MM-DD")
            If cboEdit(cbo发病时间).ListIndex < 3 Then
                txt发病时间.Text = Format(CDate(datRes), "HH:mm")
            End If
     Case txt手机号
        If Not IsNumeric(Trim(txtEdit(txt手机号).Text)) And txtEdit(txt手机号).Text <> "" Then
            MsgBox "当前录入的手机号格式不正确，请重新录入!", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End Select
End Sub

Private Sub txt发病日期_Change()
    Call SetPermitEscape(False)
    
    If IsDate(txt发病日期.Text) Then
        txt发病时间.Enabled = True
    Else
        txt发病时间.Enabled = False
        txt发病时间.Text = "__:__"
    End If
End Sub

Private Sub txt发病日期_GotFocus()
    Call zlControl.TxtSelAll(txt发病日期)
End Sub

Private Sub txt发病日期_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt发病日期_Validate(Cancel As Boolean)
    If txt发病日期.Text <> "____-__-__" And Not IsDate(txt发病日期.Text) Then
        txt发病日期.Text = "____-__-__": Cancel = True
    ElseIf txt发病日期.Text = "____-__-__" Then
        If txt发病时间.Text <> "__:__" Then
            txt发病时间.Text = "__:__"
        End If
    End If
    Call PicBasis_Resize
End Sub

Private Sub txt发病时间_Change()
    Call SetPermitEscape(False)
End Sub

Private Sub txt发病时间_GotFocus()
    Call zlControl.TxtSelAll(txt发病时间)
End Sub

Private Sub txt发病时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt发病时间_Validate(Cancel As Boolean)
    If txt发病时间.Text <> "__:__" And Not IsDate(txt发病时间.Text) Then
        txt发病时间.Text = "__:__": Cancel = True
    End If
    Call PicBasis_Resize
End Sub

Private Sub cmdEdit_Click(Index As Integer)
'说明：注意界面上要求CMD和对应TXT的Index相同
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
        
    Select Case Index
        Case cmd家庭地址
            '选择地区数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 地区 Order by 编码"
            vPoint = GetCoordPos(txtEdit(txt家庭地址).Container.hwnd, txtEdit(txt家庭地址).Left, txtEdit(txt家庭地址).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.x, vPoint.y, txtEdit(txt家庭地址).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""地区""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtEdit(txt家庭地址).SetFocus
            Else
                txtEdit(txt家庭地址).Text = rsTmp!名称
                txtEdit(txt家庭地址).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case cmd单位名称
            '选择单位信息
            strSQL = "Select ID,上级ID,末级,编码,名称,简码,地址,电话,开户银行,帐号,联系人" & _
                " From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            vPoint = GetCoordPos(txtEdit(txt单位名称).Container.hwnd, txtEdit(txt单位名称).Left, txtEdit(txt单位名称).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "合约单位", , , , , True, True, vPoint.x, vPoint.y, txtEdit(txt单位名称).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""合约单位""数据，请先到合约单位管理中设置。", vbInformation, gstrSysName
                End If
                txtEdit(txt单位名称).Tag = ""
                txtEdit(txt单位名称).SetFocus
            Else
                txtEdit(txt单位名称).Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                If InStr(GetInsidePrivs(p门诊医生站), "合约病人登记") > 0 Then txtEdit(txt单位名称).Tag = Val(rsTmp!ID)
                If txtEdit(txt单位电话).Text = "" Then
                    txtEdit(txt单位电话).Text = Nvl(rsTmp!电话)
                End If
                txtEdit(txt单位名称).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
    End Select
End Sub

Private Sub cboRegist_Click()
'功能：选择某次历史就诊记录时，读取相关的病人信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    If cboRegist.ListIndex = -1 Then
        '按当前列表无数据刷新子窗体
        Call ClearPatiInfo
        
        cboRegist.Tag = ""
        Call SetPermitEdit
        Call PicBasis_Resize
        '刷新子窗体数据
        Call SubWinRefreshData(tbcSub.Selected)
        
        '读取简单病历内容
        If mblnDocInput Then Call LoadDocData
        Exit Sub
    End If
    If cboRegist.ListIndex = mintPreTime Then Exit Sub
    mintPreTime = cboRegist.ListIndex
       
    cboRegist.Tag = "Loading"   '用于对编辑控件数据是否改变的判断时排开加载时的初次改变
    mblnPatiChange = False              '用于记录是否修改过数据，判断是否需要保存
    
    On Error GoTo errH
    strSQL = "Select E.号类,B.Id,B.NO,B.门诊号,B.姓名,B.性别,B.年龄,A.出生日期,B.医疗付款方式,A.职业," & _
        "   A.费别,A.险类,A.医保号,B.急诊,A.结算模式,B.发生时间,B.执行人,B.执行状态,B.执行时间," & _
        "   B.执行部门ID as 科室ID,B.诊室,B.社区,D.社区号,C.名称 as 科室,B.复诊,B.摘要," & _
        "   A.身份证号,A.监护人,A.家庭地址,A.家庭电话,A.工作单位,A.合同单位id,A.单位电话,B.发病时间,B.发病地址," & _
        "   A.民族,A.国籍,A.区域,A.婚姻状况,A.家庭地址邮编,A.单位邮编,A.出生地点,B.传染病上传,A.其他证件,a.户口地址,a.户口地址邮编,a.籍贯,a.email,a.qq,A.病人类型,A.病人ID,A.手机号" & _
        " From 病人信息 A,病人挂号记录 B,部门表 C,病人社区信息 D,挂号安排 E" & _
        " Where A.病人ID=B.病人ID And B.ID=[1] And B.执行部门ID=C.ID" & _
        " And B.病人ID=D.病人ID(+) And B.社区=D.社区(+) And B.号别=E.号码(+)"
        '按ID读取挂号记录，不用加记录性质、状态的条件
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboRegist.ItemData(cboRegist.ListIndex))
    With rsTmp
        txtEdit(txt姓名).Text = "" & !姓名
        txtEdit(txt姓名).Tag = "" & !姓名
        '显示病人颜色
        If Not IsNull(!险类) And Nvl(rsTmp!病人类型) = "" Then
            txtEdit(txt姓名).ForeColor = &HC0&
        Else
            txtEdit(txt姓名).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!病人类型))
        End If
        
        Call zlControl.CboLocate(cboEdit(cbo性别), "" & !性别)
                
        If Not IsNull(!出生日期) Then
            txt出生日期.Text = Format(!出生日期, "yyyy-MM-dd")
            If Format(!出生日期, "HH:mm") <> "00:00" Then txt出生时间.Text = Format(!出生日期, "HH:mm")
        End If

        Call LoadOldData("" & !年龄, txtEdit(txt年龄), cboEdit(cbo年龄))
        
        Call zlControl.CboLocate(cboEdit(cbo职业), "" & !职业)
        If Not IsNull(!发病时间) Then
            txt发病日期.Text = Format(!发病时间, "yyyy-MM-dd")
            txt发病时间.Text = Format(!发病时间, "HH:mm")
            If txt发病时间.Text = "00:00" Then txt发病时间.Text = "__:__"
        Else
            txt发病日期.Text = "____-__-__": txt发病时间.Text = "__:__"
        End If
        txtEdit(txt发病地址).Text = Nvl(rsTmp!发病地址)
        lbl急.Visible = Nvl(!急诊, 0) <> 0
        lblRec.Visible = Nvl(!结算模式, 0) <> 0
                Call picPatiInput_Resize
        
        txtEdit(txt身份证号).Tag = "不触发Change事件"
        txtEdit(txt身份证号).Text = "" & !身份证号
        If zlCommFun.ActualLen(txtEdit(txt身份证号).Tag) > 12 And mblnMaskID Then   '生成身份证号掩码
            txtEdit(txt身份证号).Text = Mid(txtEdit(txt身份证号).Text, 1, 12) & String(Len(Mid(txtEdit(txt身份证号).Text, 13, 2)), "*") & Mid(txtEdit(txt身份证号).Text, 15)
        End If
        txtEdit(txt身份证号).Tag = "" & !身份证号
                lblEdit(20).Tag = "" & !身份证号  '备份数据保存检查判断时用到
        
        If Val("" & !复诊) = 1 Then
            optState(opt复诊).Value = True
        Else
            optState(opt初诊).Value = True
        End If
        txtEdit(txt单位名称).Text = "" & !工作单位
        txtEdit(txt单位名称).Tag = Val("" & !合同单位id)
                
        txtEdit(txt单位电话).Text = "" & !单位电话
        txtEdit(txt家庭地址).Text = "" & !家庭地址
        txtEdit(txt监护人).Text = "" & !监护人
        txtEdit(txt家庭电话).Text = "" & !家庭电话
        txtEdit(txt手机号).Text = "" & !手机号
        txtEdit(txt就诊摘要).Text = "" & !摘要
                        
        lblShow(lbl费别).Caption = "" & !费别
        lblShow(lbl付款).Caption = "" & !医疗付款方式
        lblShow(lbl号类).Caption = "" & !号类
        lblShow(lbl医保号).Caption = "" & !医保号
        If IsNull(!社区号) Then
            lblShow(lbl社区号).Caption = ""
            lblShow(lbl社区号).Visible = False
            lblTitle社区号.Visible = False
        Else
            lblShow(lbl社区号).Caption = "" & !社区号
            lblShow(lbl社区号).Visible = True
            lblTitle社区号.Visible = True
        End If
                
        '诊断
        lblDiag(1).Caption = GetPatiDiagnose(Val(rsTmp!病人ID & ""), cboRegist.ItemData(cboRegist.ListIndex), 1)
        
        '病人信息
        If mintActive = pt转诊 Then
            mPatiInfo.类型 = pt转诊
        Else
            mPatiInfo.类型 = Decode(Nvl(!执行状态, 0), 0, 0, 2, 1, 1, 2)
        End If
        mPatiInfo.门诊号 = Nvl(!门诊号)
        mPatiInfo.挂号ID = !ID
        mPatiInfo.病人ID = !病人ID
        mPatiInfo.挂号单 = !NO
        mPatiInfo.科室ID = !科室ID
        mPatiInfo.诊室 = Nvl(!诊室)
        mPatiInfo.社区 = Nvl(!社区, 0)
        mPatiInfo.社区号 = Nvl(!社区号)
        mPatiInfo.挂号时间 = !发生时间
        mPatiInfo.性别 = "" & !性别
        mPatiInfo.婚姻状况 = "" & !婚姻状况
        
        mPatiInfo.民族 = "" & !民族
        mPatiInfo.国籍 = "" & !国籍
        mPatiInfo.区域 = "" & !区域
        mPatiInfo.出生地点 = "" & !出生地点
        mPatiInfo.传染病上传 = Val("" & !传染病上传)
        mPatiInfo.家庭地址邮编 = "" & !家庭地址邮编
        mPatiInfo.单位邮编 = "" & !单位邮编
        mPatiInfo.其他证件 = "" & !其他证件
        mPatiInfo.户口地址 = "" & !户口地址
        mPatiInfo.户口地址邮编 = "" & !户口地址邮编
        mPatiInfo.籍贯 = "" & !籍贯
        mPatiInfo.Email = "" & !Email
        mPatiInfo.QQ = "" & !QQ
        
        If mPatiInfo.类型 = pt已诊 Then
            mPatiInfo.数据转出 = zlDatabase.NOMoved("病人挂号记录", !NO)
        Else
            mPatiInfo.数据转出 = False
        End If
        picPatient.Visible = ReadPatPricture(mPatiInfo.病人ID, imgPatient)
    End With
    '附加信息
    Call ucPatiVitalSigns.LoadPatiVitalSigns(mPatiInfo.病人ID, cboRegist.ItemData(cboRegist.ListIndex))
    strSQL = "Select 信息名,信息值 From 病人信息从表 Where 病人ID=[1] And (就诊ID=[2] Or 就诊ID is Null) Order by Nvl(就诊ID,999999999)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.病人ID, cboRegist.ItemData(cboRegist.ListIndex))
    rsTmp.Filter = "信息名='去向'"
    '删除自由添加的
    If cboEdit(cbo去向).ListCount <> 0 Then
        If cboEdit(cbo去向).ItemData(cboEdit(cbo去向).ListCount - 1) = -1 Then
            cboEdit(cbo去向).RemoveItem (cboEdit(cbo去向).ListCount - 1)
        End If
    End If
    If Not rsTmp.EOF Then
        If Not zlControl.CboLocate(cboEdit(cbo去向), Nvl(rsTmp!信息值)) Then
            cboEdit(cbo去向).AddItem Nvl(rsTmp!信息值)
            cboEdit(cbo去向).ItemData(cboEdit(cbo去向).NewIndex) = -1
        End If
        cboEdit(cbo去向).Text = Nvl(rsTmp!信息值)
    Else
        cboEdit(cbo去向).ListIndex = 0
    End If
    cboEdit(cbo发病时间).ListIndex = 0
    
    For i = 0 To cboEdit.UBound
        If i = 3 Then i = i + 2
        If cboEdit(i).ListIndex <> -1 Then
            cboEdit(i).Tag = cboEdit(i).List(cboEdit(i).ListIndex)
        Else
            cboEdit(i).Tag = ""
        End If
    Next
    txtEdit(txt发病).Text = ""
    txtEdit(txt发病).Tag = ""
    cboEdit(cbo发病时间).ListIndex = 0
    cboEdit(cbo发病时间).Tag = ""
    
    Call ShowAller
    
    '刷新子窗体数据
    Call SubWinRefreshData(tbcSub.Selected)
    
    '读取简单病历内容
    If mblnDocInput Then Call LoadDocData
    
    cboRegist.Tag = ""
    Call SetPermitEdit

    Call PicBasis_Resize
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cboRegist.Tag = ""
End Sub

Private Sub SetDocData(ByVal rsTmp As Recordset, ByVal intType As Integer)
'功能：设置快捷面板的内容
'参数：intType=0病历读取，intType=1范文导入，不清空病历ID
    Dim i As Long, j As Long, arrTmp As Variant
    Dim strContent As String
    
    With rsTmp
        If .RecordCount > 0 Then
            arrTmp = Split("-10,2,3,5,6", ",") '病人主诉,现病史,既往史,家族史,体格检查
            For i = 0 To UBound(arrTmp)
                .Filter = "预制提纲id=" & arrTmp(i)
                rtfEdit(i).Text = ""
                If intType = 1 Then
                    '导入范文后所有都可用。
                    rtfEdit(i).Locked = False
                    rtfEdit(i).BackColor = HColor
                End If
                For j = 1 To .RecordCount
                    If j = 1 Then
                        strContent = "" & !内容文本
                        If InStr(strContent, lblDoc(i).Tag) = 1 Then strContent = Mid(strContent, Len(lblDoc(i).Tag) + 1)
                        rtfEdit(i).Text = strContent
                        If intType = 0 Then rtfEdit(i).Tag = !ID
                    Else
                        rtfEdit(i).Text = rtfEdit(i).Text & vbCrLf & !内容文本
                        If intType = 0 Then rtfEdit(i).Tag = rtfEdit(i).Tag & "," & !ID
                    End If
                    .MoveNext
                Next
            Next
        End If
    End With
End Sub

Private Sub LoadDocData()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, j As Long, blnLoading As Boolean
    
    
    blnLoading = cboRegist.Tag = "Loading"
    If Not blnLoading Then cboRegist.Tag = "Loading"    '避免text赋值时调用text_Change事件中改变其它控件的可用状态
    
    For i = 0 To rtfEdit.UBound
        rtfEdit(i).Text = ""
        rtfEdit(i).Tag = ""
    Next
    lbl病历名称.Caption = ""    '仅用于技术人员查原因，例如：选择提纲词句时，没有列出预计的词句，可根据病历文件名称查是否设置了提纲词句对应
    
    '只显示简单病历模式下产生的文件
    strSQL = "Select id,文件id,签名级别,病历名称,保存人 From 电子病历记录 A Where 病人id = [1] And 主页id = [2] And 病历种类 = 1" & vbNewLine & _
            " And Exists(Select 1 From 病历文件列表 B Where A.文件ID = B.ID And B.保留 = '3')"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.病人ID, mPatiInfo.挂号ID)
    If rsTmp.RecordCount > 0 Then
        If SetCompendsTag(Val("" & rsTmp!文件ID)) Then
            mPatiInfo.病历文件id = Val("" & rsTmp!文件ID)
            mPatiInfo.是否签名 = IIf(Val("" & rsTmp!签名级别) > 0, True, False)
            mPatiInfo.病历id = rsTmp!ID
            mPatiInfo.保存人 = "" & rsTmp!保存人
            lbl病历名称.Caption = "" & rsTmp!病历名称
                                            
            '读取提纲下的段落文本,对象属性为-1表示提纲标题文本
            strSQL = "Select A.预制提纲id, B.内容文本, B.ID" & vbNewLine & _
                    "From 电子病历内容 A, 电子病历内容 B" & vbNewLine & _
                    "Where A.文件id = [1] And A.对象类型 = 1 And A.预制提纲id+0 In(-10,5,2,6,3)" & vbNewLine & _
                    "      And B.父id = A.ID And B.对象类型 = 2 Order By A.预制提纲id, B.对象序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.病历id)
            Call SetDocData(rsTmp, 0)
        End If
    Else
        If lbl急.Visible Then
            strSQL = " And (R.事件 = '急诊'  OR R.事件 IS NUll)"
        Else
            If optState(opt复诊).Value Then
                strSQL = " And (R.事件 = '门诊' Or R.事件 = '复诊'  OR R.事件 IS NUll)"
            Else
                strSQL = " And (R.事件 = '门诊' Or R.事件 = '初诊'  OR R.事件 IS NUll )"
            End If
        End If
        '系统定义了门(急)诊病历且对当前病人适用，具有5个固定预制提纲,才显示病历录入面板.
        strSQL = "Select F.ID, F.名称 as 病历名称" & vbNewLine & _
                "From (Select F.ID, F.通用, A.科室id, F.名称,Decode(R.事件,Null,2,1) 事件" & vbNewLine & _
                "       From 病历文件列表 F, 病历应用科室 A, 病历时限要求 R" & vbNewLine & _
                "       Where F.ID = A.文件id(+) And F.ID = R.文件id(+) And F.种类 = 1 And F.保留= '3'" & strSQL & ") F" & vbNewLine & _
                "Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [2]" & vbNewLine & _
                "Order By F.事件,F.通用 Desc,F.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.挂号ID, mPatiInfo.科室ID)
        If rsTmp.RecordCount > 0 Then
            mPatiInfo.病历文件id = rsTmp!ID
            lbl病历名称.Caption = "" & rsTmp!病历名称
            If SetCompendsTag(mPatiInfo.病历文件id) = False Then
                mPatiInfo.病历文件id = 0: lbl病历名称.Caption = ""
            End If
        Else
            mPatiInfo.病历文件id = 0: lbl病历名称.Caption = ""
        End If
        
        mPatiInfo.病历id = 0
        mPatiInfo.是否签名 = False
    End If
     '设置字体
     Call SetRTFEditFontSize
     
    PicOutDoc.Tag = ""
    If Not blnLoading Then cboRegist.Tag = ""
    cmdImportEPRDemo.Visible = mPatiInfo.病历文件id <> 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function SetCompendsTag(ByVal lng病历文件id As Long) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    
    strSQL = "Select Decode(A.预制提纲id, -10, 0, 5, 3, 2, 1, 6, 4, 3, 2) As 序号, B.内容文本" & vbNewLine & _
            "From 病历文件结构 A, 病历文件结构 B" & vbNewLine & _
            "Where A.文件id = [1] And A.预制提纲id+0 In (-10,5,2,6,3) And A.Id = B.父id And B.对象类型 = 2" & vbNewLine & _
            "Order By 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病历文件id)
    If rsTmp.RecordCount > 0 And rsTmp.RecordCount <= 5 Then
        If rsTmp!序号 & "" = "0" Then  '必须包含主诉
            For i = 0 To rsTmp.RecordCount - 1
                lblDoc(Val(rsTmp!序号 & "")).Tag = rsTmp!内容文本       '用于保存Rtf文件替换内容时定位
                rsTmp.MoveNext
            Next
            For i = 1 To lblDoc.Count - 1
                If lblDoc(i).Tag = "" Then
                    lblDoc(i).Visible = False
                    rtfEdit(i).Visible = False
                Else
                    lblDoc(i).Visible = True
                    rtfEdit(i).Visible = True
                End If
            Next
            picPatiInput_Resize
            SetCompendsTag = True
        End If
    End If
End Function


Private Sub mclsEPRs_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    Call gobjKernel.ShowDiagHelp(Modal, Me, DiagnosisID)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
'功能：身份证识别成功后激活
    mstrIDCard = strID
    If mstrFindType = "二代身份证" Then
        PatiIdentify.Text = mstrIDCard
    Else
        PatiIdentify.Text = "" '否则清除(目前是在已清除情况下才能激活)。
    End If
    Call ExecuteFindPati(False, mstrIDCard)
End Sub

Private Function CheckHaveAdvice(ByVal lng病人ID As Long, ByVal str挂号单 As String) As Boolean
'功能：判断病人是否开了医嘱
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "select 1 from 病人医嘱记录 where 病人ID=[1] and 挂号单=[2] and rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, str挂号单)
    CheckHaveAdvice = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：刷新子窗体界面及数据
'说明：仅在人为切换界面卡片激活
    Dim objControl As CommandBarControl
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    
    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值
     
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "病历"
                Set objItem = tbcSub.InsertItem(Index, "病历信息", mcolSubForm("_病历").hwnd, 0)
                objItem.Tag = "病历"
            Case "新病历"
                Set objItem = tbcSub.InsertItem(Index, "电子病历", mcolSubForm("_新病历").hwnd, 0)
                objItem.Tag = "新病历"
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    End If
     
    '刷新子窗体对应的CommandBar
    Call SubWinDefCommandBar(Item)
    
    '刷新子窗体数据
    Call SubWinRefreshData(Item)
    
    If Visible Then mfrmActive.SetFocus
    
    '自动新增一份门诊/急诊/复诊病历/如果是医嘱，则新增医嘱，先判断没有医嘱再新增
    If Item.Tag = "病历" And mlng自动进行 = 1 Then
        mblnUnRefresh = True
        Call mclsEPRs.zlOpenDefaultEPR(mstr挂号单)
        '因为执行命令的是非模态窗体，所以在mclsAdvices和mclsEPRs的active中设置 mblnUnRefresh = False
    ElseIf Item.Tag = "医嘱" And mlng自动进行 = 2 Then
        If CheckHaveAdvice(mlng病人ID, mstr挂号单) = False Then
            cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
            Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 4 Then
        Item.Handle = lvwIncept.hwnd
    ElseIf Item.ID = 5 Then
        Item.Handle = lvwReserve.hwnd
    ElseIf Item.ID = pt排队叫号 Then
        Item.Handle = mobjQueue.zlGetForm.hwnd
    ElseIf Item.ID = 7 Then  '回诊
        Item.Handle = lvwPatiHZ.hwnd
    ElseIf Item.ID = 3 Then '已诊
        Item.Handle = picYZ.hwnd
    ElseIf Item.ID = 8 Then
        Item.Handle = rptNotify.hwnd
    Else
        Item.Handle = lvwPati(Item.ID - 1).hwnd
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long, lngTopPanelHeight As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    If mblnPatiDetail = False Then
        PicBasis.Height = fraLine(cbo年龄).Height + fraLine(cbo年龄).Top
        PicPatiInfo.Height = 0
    Else
        PicBasis.Height = ucPatiVitalSigns.Top + ucPatiVitalSigns.Height
        PicPatiInfo.Height = IIf(mbytSize = 0, 800, 950)
    End If
    
    With Me.picPatiInput
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = .ScaleTop + PicBasis.Height + IIf(mblnDocInput, PicOutDoc.Height, 0) + PicPatiInfo.Height
        lngTopPanelHeight = .Height 'PicBasis.ScaleHeight + IIf(mblnDocInput, PicOutDoc.ScaleHeight, 0) + PicPatiInfo.ScaleHeight
    End With
    
    With Me.tbcSub
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = lngTop + lngTopPanelHeight: .Height = lngBottom - lngTop - lngTopPanelHeight
    End With
    With Me.fraRoom
        .Visible = Me.stbThis.Visible
        .Left = Me.stbThis.Panels(3).Left + 60: .Top = Me.stbThis.Top + 60
    End With
    
    PatiIdentify.Width = lngLeft - PatiIdentify.Left - 500
    picFind.Top = lngTop
    PatiIdentify.Top = picFind.Top
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If mblnDocInput Then Call HideWordInput
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim blnSetup As Boolean
    
    mblnMsgOk = False: mblnFirstMsg = False
    mblnIsInit = False
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    Call zlDatabase.SetPara("病人查找方式", mintFindType, glngSys, p门诊医生站, blnSetup)

    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("医护功能", tbcSub.Selected.Tag, glngSys, p门诊医生站, blnSetup)
    End If
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
    End If

    '公共部件固定按第一个控件的样式保存，工作站部件如果第一个是打印，则固定是图标样式,所以需恢复为其它按钮的样式
    If Me.Visible Then  'Form_load中退出时不处理
        cbsMain(2).Controls(1).Style = cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style
        Call SaveWinState(Me, App.ProductName)
    End If

    mstrIDCard = ""
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjSquareCard = Nothing

    '--关闭所有排队的窗体
    If Not mobjQueue Is Nothing Then
        Call mobjQueue.CloseWindows
        Set mobjQueue = Nothing
    End If
    If Not mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = Nothing
    End If
    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mclsAdvices = Nothing
    Set mclsEMR = Nothing
    Set mclsEPRs = Nothing
    Set mrsAller = Nothing
    Set mobjEPRDoc = Nothing
    Set mfrmActive = Nothing
    Set gobjPublicPacs = Nothing
    Set mobjKernel = Nothing
    
    '问题号:57566
    mlng接诊控制 = 0
    mlng提前接收时间 = 0
    
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Set mclsMsg = Nothing
    Set mrsMsg = Nothing
End Sub

Private Sub lblRoom_Click()
    Call SetRoomState(lblRoom.BackColor = COLOR_FREE)
End Sub

Private Sub lvwPati_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati(Index), ColumnHeader.Index)
End Sub

Private Sub lvwPati_GotFocus(Index As Integer)
    'MouseDown先于GotFocus执行
    If Not mblnMouseDown And Not lvwPati(Index).SelectedItem Is Nothing Then
        Call lvwPati_ItemClick(Index, lvwPati(Index).SelectedItem)
    End If
End Sub

Private Sub lvwPati_DblClick(Index As Integer)
'功能：双击自动接诊或完成接诊
    Dim objControl As CommandBarControl
    Dim objItem As ListItem
    Dim vPoint As POINTAPI
    
    Call GetCursorPos(vPoint)
    Call ScreenToClient(lvwPati(Index).hwnd, vPoint)
    Set objItem = lvwPati(Index).HitTest(vPoint.x * Screen.TwipsPerPixelX, vPoint.y * Screen.TwipsPerPixelY)
    If Not objItem Is Nothing And InStr(mstrPrivs, "病人接诊") > 0 Then
        If Index = pt候诊 Then
            Set objControl = cbsMain.FindControl(, conMenu_Manage_Receive, True, True)
        ElseIf Index = pt就诊 Then
            Set objControl = cbsMain.FindControl(, conMenu_Manage_Finish, True, True)
        End If
        If Not objControl Is Nothing Then
            If objControl.Enabled Then Call cbsMain_Update(objControl) '首次执行，没有显示菜单前，事件没有执行
            If objControl.Enabled Then objControl.Execute
        End If
    End If
End Sub

Private Sub LvwItemClick(ByVal Index As Integer, ByVal Item As MSComctlLib.ListItem)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据处理
    '入参:index-5回诊;0-候诊;1-就诊;2-完成就诊
    '编制:刘兴洪
    '日期:2011-01-17 10:59:25
    '主要是加入了回诊后,需要在点击相关的列表时,处理相关的数据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, strTmp As String
    Dim intCount As Integer
    Dim objLvw As ListView
    Dim str身份证号 As String
    Dim str病人IDs As String
    
    'Index:5-回诊
    If Index = pt回诊 Then
        Set objLvw = lvwPatiHZ
    ElseIf Index = pt预约 Then
        Set objLvw = lvwReserve
    ElseIf Index = pt转诊 Then
        Set objLvw = lvwIncept
    Else
        Set objLvw = lvwPati(Index)
    End If
    
    If objLvw.SelectedItem Is Nothing Then Exit Sub '非正常情况
    With objLvw.SelectedItem
        '当前活动列表
        mintActive = Index
        If .Key = mstrPrePati Then Exit Sub
        mstrPrePati = .Key
        '当前选择病人的列表中才可以看见选择项,以便区分
        lvwPatiHZ.HideSelection = Index <> pt回诊
        For i = 0 To lvwPati.UBound
            lvwPati(i).HideSelection = i <> Index
        Next
        mstr挂号单 = .Text
        mlng病人ID = Val("" & .Tag) '预约病人可能未建档
        mlng科室ID = Val(.ListSubItems(3).Tag)
        str身份证号 = .ListSubItems(6).Tag
        
        LockWindowUpdate Me.hwnd
        
        '验证身份证号
        If str身份证号 <> "" Then
            If mobjPatient Is Nothing Then
                On Error Resume Next
                Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                err.Clear: On Error GoTo 0
                If mobjPatient Is Nothing Then
                    MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, Me.Caption
                Else
                    Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.用户名)
                End If
            End If
            strTmp = ""
            If Not mobjPatient Is Nothing Then
                If mobjPatient.CheckPatiIdcard(str身份证号) Then
                    strTmp = str身份证号
                End If
            End If
            str身份证号 = strTmp
        End If
        
        On Error GoTo errH
        
        If str身份证号 <> "" Then
            strSQL = "select a.病人id from 病人信息 a where a.病人id<>[1] and a.身份证号=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, str身份证号)
            Do While Not rsTmp.EOF
                str病人IDs = str病人IDs & "," & rsTmp!病人ID
                rsTmp.MoveNext
            Loop
            If str病人IDs <> "" Then
                str病人IDs = mlng病人ID & str病人IDs
            End If
        End If
        
        
        If str病人IDs = "" Then
            '读取"历史的"就诊记录
            strSQL = "Select A.ID,A.NO,A.发生时间 as 时间,B.名称 as 科室 From 病人挂号记录 A,部门表 B" & _
                " Where A.执行部门ID=B.ID And A.病人ID=[1] And A.发生时间<=[2] And A.记录性质=1 And A.记录状态=1 Order by A.发生时间 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, CDate(Item.ListSubItems(2).Tag))
        Else
            strSQL = "Select A.ID,A.NO,A.发生时间 as 时间,B.名称 as 科室 From 病人挂号记录 A,部门表 B" & _
                " Where A.执行部门ID=B.ID And A.病人ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)" & _
                " And A.发生时间<=[2] And A.记录性质=1 And A.记录状态=1 Order by A.发生时间 Desc,a.接收时间 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str病人IDs, CDate(Item.ListSubItems(2).Tag))
        End If
        
        cboRegist.Clear
        Do While Not rsTmp.EOF
            cboRegist.AddItem Format(rsTmp!时间, "YYMMdd") & rsTmp!科室
            cboRegist.ItemData(cboRegist.NewIndex) = rsTmp!ID
            If rsTmp!NO = mstr挂号单 Then
                mlng挂号ID = rsTmp!ID
                Call zlControl.CboSetIndex(cboRegist.hwnd, cboRegist.NewIndex)
            End If
            
            '当日多科就诊
            If Format(rsTmp!时间, "yyyy-MM-dd") = Format(CDate(Item.ListSubItems(2).Tag), "yyyy-MM-dd") Then
                intCount = intCount + 1
            End If
            
            rsTmp.MoveNext
        Loop
        If cboRegist.ListIndex = -1 Then
            Call zlControl.CboSetIndex(cboRegist.hwnd, 0)
        End If
        
        lbl多科就诊.Visible = intCount > 1 And mintActive = pt就诊
        
        mintPreTime = -1
        If mblnDocInput Then edtEditor.Text = ""
        
        Call cboRegist_Click
        
        LockWindowUpdate 0
    End With
    Exit Sub
errH:
    LockWindowUpdate 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub
Private Sub lvwPati_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Call LvwItemClick(Index, Item)
End Sub

Private Sub ShowAller()
    Set mrsAller = New ADODB.Recordset
    Call LoadPatiAllergy(mPatiInfo.病人ID, , mrsAller)
    mrsAller.Filter = "病人ID=" & mPatiInfo.病人ID & " and 挂号单<>'" & mPatiInfo.挂号单 & "'"
    cmdAller.Enabled = mrsAller.RecordCount > 0
    
    Call LoadAllerInfo(mrsAller)
End Sub

Private Sub LoadAllerInfo(ByRef rsTmp As ADODB.Recordset)
    Dim i As Long
    Dim lngRow As Long
    
    With vsAller
        .Clear
        .Cols = 0   '清除所有列及列宽
        .Cols = 1
        .Rows = 2
        .RowHidden(1) = True
        '显示本次挂号的过敏记录用于修改
        If rsTmp.State = 1 Then
            rsTmp.Filter = "挂号单='" & mstr挂号单 & "'"
            If rsTmp.RecordCount > 0 Then
                .Cols = rsTmp.RecordCount + 1
                For i = 0 To rsTmp.RecordCount - 1
                    '其它来源的可能有重复
                    lngRow = -1
                    If Not IsNull(rsTmp!药物ID) Then
                        lngRow = .FindRow(CLng(rsTmp!药物ID))
                    ElseIf Not IsNull(rsTmp!药物名) Then
                        lngRow = .FindRow(CStr(rsTmp!药物名), 0)
                    End If
                    If lngRow = -1 Then
                        .TextMatrix(0, i) = "" & rsTmp!药物名
                        .Cell(flexcpData, 0, i) = .TextMatrix(0, i)   '用于判断是否被修改
                        .Cell(flexcpData, 1, i) = rsTmp!过敏时间 & ""
                        .ColData(i) = rsTmp!药物ID & ""
                        .TextMatrix(1, i) = rsTmp!过敏源编码 & ""
                        .ColAlignment(i) = flexAlignLeftCenter
                        .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + 300
                        rsTmp.MoveNext
                    End If
                Next
            End If
        End If
        .ColAlignment(.Cols - 1) = flexAlignLeftCenter
        .ColWidth(.Cols - 1) = 1200
        .Select 0, .Cols - 1
        If .Cols = 1 Then vsAller.ComboList = "..."
        
    End With
End Sub

Private Sub lvwPati_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMouseDown = True
End Sub

Private Sub lvwPati_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim objItem As ListItem
    
    mblnMouseDown = False
    
    If Button = 2 And InStr(mstrPrivs, "病人接诊") > 0 Then
        Set objItem = lvwPati(Index).HitTest(x, y)
        If Not objItem Is Nothing Then
            Set objPopup = cbsMain.ActiveMenuBar.Controls(2)
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub picPatiInput_Resize()
    Dim lngLeft As Long, i As Long, lngCount As Long
    Dim lngTmp As Long
    
    On Error Resume Next
    lngLeft = picPatiInput.ScaleLeft + picPatiInput.ScaleWidth
        
    If mblnDocInput Then
        PicOutDoc.Top = PicBasis.Top + PicBasis.Height + 60
        PicPatiInfo.Top = PicOutDoc.Top + PicOutDoc.Height + 60
    Else
        PicPatiInfo.Top = PicBasis.Top + PicBasis.Height + 60
    End If
    PicBasis.Width = picPatiInput.ScaleWidth
    If mblnDocInput Then PicOutDoc.Width = PicBasis.ScaleWidth
    PicPatiInfo.Width = PicBasis.ScaleWidth
        
    cboEdit(cbo去向).Width = IIf(mbytSize = 1, 1440, 1150)
    fraLine(cbo去向).Width = cboEdit(cbo去向).Width - 30
        
    lbl多科就诊.Left = fraLine(cbo去向).Left + fraLine(cbo去向).Width + 30
    lbl多科就诊.Top = optState(opt复诊).Top + 20
        
    lbl急.Top = fraLine(cbo性别).Top - 60
    lbl急.Left = lngLeft - lbl急.Width - 30
    
    lblRec.Top = lbl急.Top
    lblRec.Left = lbl急.Left - lblRec.Width - 10
    
    If lbl急.Visible = False Then lblRec.Left = lbl急.Left
    
    lngTmp = lngLeft
    If lbl急.Visible Then
        lngTmp = lbl急.Left
    End If
    If lblRec.Visible Then
        lngTmp = lblRec.Left
    End If
    picExpand.Left = lngTmp - picExpand.Width - 150
        
    txtEdit(txt就诊摘要).Width = PicBasis.Width - txtEdit(txt就诊摘要).Left - 30
    txtEdit(txt发病地址).Width = PicBasis.Width - txtEdit(txt发病地址).Left - 30
    vsAller.Width = txtEdit(txt就诊摘要).Width - picPrompt.Width - 30 + txtEdit(txt家庭电话).Left - txtEdit(txt监护人).Left + txtEdit(txt家庭电话).Width
    lblDiag(1).Width = PicPatiInfo.Width - lblDiag(1).Left
           
    
    '针对1024*786及以上特殊处理
    '---------------------------------------------------------------------------------------------------------------------
    If mblnDocInput Then
        If lngLeft - 200 > fraLine(cbo年龄).Left Then
            rtfEdit(txt主诉).Width = (picPatiInput.ScaleWidth - lblDoc(txt主诉).Width * 4 - 100) / 2
        Else
            rtfEdit(txt主诉).Width = fraLine(cbo性别).Left + fraLine(cbo性别).Width - rtfEdit(txt主诉).Left
        End If
        
        lngCount = 1
        rtfEdit(txt主诉).Left = lblDoc(txt主诉).Left + lblDoc(txt主诉).Width + 70
        lblDoc(txt主诉).Top = 0
        For i = 1 To rtfEdit.Count - 1
            If rtfEdit(i).Visible Then
                lngCount = lngCount + 1
                lblDoc(i).Left = IIf(lngCount Mod 2 = 1, lblDoc(txt主诉).Left, rtfEdit(txt主诉).Left + rtfEdit(txt主诉).Width + 100)
                rtfEdit(i).Left = lblDoc(i).Left + lblDoc(i).Width + 70
                rtfEdit(i).Width = IIf(lngCount Mod 2 = 1, rtfEdit(txt主诉).Width, lngLeft - rtfEdit(i).Left - 100)
                lblDoc(i).Top = ((lngCount - 1) \ 2) * rtfEdit(i).Height - (((lngCount - 1) \ 2) * 15)
                rtfEdit(i).Top = ((lngCount - 1) \ 2) * rtfEdit(i).Height - (((lngCount - 1) \ 2) * 15)
            End If
        Next
        
        
        cmdSign.Left = lngLeft - 100 - cmdSign.Width
        cmdSign.Top = ((lngCount) \ 2) * rtfEdit(txt主诉).Height + 50
        cmdUpdate.Left = cmdSign.Left
        cmdUpdate.Top = cmdSign.Top + cmdSign.Height + 20
        cmdImportEPRDemo.Left = cmdUpdate.Left - cmdImportEPRDemo.Width - 50
        cmdImportEPRDemo.Top = cmdUpdate.Top
        
        lbl医生(1).Left = cmdSign.Left - lbl医生(1).Width - 100
        lbl医生(0).Left = lbl医生(1).Left - lbl医生(0).Width - 20
        lbl医生(0).Top = ((lngCount) \ 2) * rtfEdit(txt主诉).Height + 150
        lbl医生(1).Top = lbl医生(0).Top
        
        picPrompt.Left = rtfEdit(txt主诉).Left + rtfEdit(txt主诉).Width + 120
        picPrompt.Top = ((lngCount) \ 2) * rtfEdit(txt主诉).Height + 550
        lbl提示.Left = picPrompt.Left + picPrompt.Width + 60
        lbl提示.Top = ((lngCount) \ 2) * rtfEdit(txt主诉).Height + 550
        
        lbl病历名称.Left = rtfEdit(txt主诉).Left + rtfEdit(txt主诉).Width + 400
        lbl病历名称.Top = ((lngCount) \ 2) * rtfEdit(txt主诉).Height + 250
        PicOutDoc.Height = cmdUpdate.Top + cmdUpdate.Height + 50
        Call Form_Resize
    End If
    '--------------------------------------------------------------------------
    
    Call PicBasis_Resize
    Call PicPatiInfo_Resize
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 3 Then
        Call SetRoomState(lblRoom.BackColor = COLOR_FREE)
    End If
End Sub

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Sub InitPatiData()
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 名称,编码 From 性别"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Call zlControl.CboAddData(cboEdit(cbo性别), rsTmp, True)
               
    Call SetCboFromList(Array("岁", "月", "天"), cboEdit(cbo年龄), 0)
    Call SetCboFromList(Array(" ", "小时前", "日前", "周前", "月前", "年前"), cboEdit(cbo发病时间), 0)
    
    strSQL = "Select 名称,编码 From 职业"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Call zlControl.CboAddData(cboEdit(cbo职业), rsTmp, True)
    
    strSQL = "Select 名称, 编码 From 病人去向"
    cboEdit(cbo去向).Clear
    cboEdit(cbo去向).AddItem ("")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Call zlControl.CboAddData(cboEdit(cbo去向), rsTmp, False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitCondFilter()
    Dim curDate As Date, intDay As Long
    Dim intStart As Long
    
    cboSelectTime.Clear
    
    With cboSelectTime
        .AddItem "今天"
        .ItemData(.NewIndex) = 0
        .AddItem "一周内"
        .ItemData(.NewIndex) = 7
        .AddItem "15天内"
        .ItemData(.NewIndex) = 15
        .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "60天内"
        .ItemData(.NewIndex) = 60
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    
    '已诊病人时间范围
    curDate = zlDatabase.Currentdate
    
    intStart = Val(zlDatabase.GetPara("已诊病人结束间隔", glngSys, p门诊医生站, "0", Array(lbl就诊时间, cboSelectTime), InStr(";" & mstrPrivs & ";", ";参数设置;") > 0))
    If lbl就诊时间.ForeColor <> vbBlue Then
        '私有参数
        mvCondFilter.End = Format(curDate, "yyyy-MM-dd 23:59:59")
        mvCondFilter.Begin = Format(mvCondFilter.End, "yyyy-MM-dd 00:00:00")
                If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
    Else
        '系统参数(恢复成管理员设置的值，防止通方)
        mvCondFilter.End = Format(curDate + intStart, "yyyy-MM-dd 23:59:59")
        intDay = Val(zlDatabase.GetPara("已诊病人开始间隔", glngSys, p门诊医生站, "7", Array(lbl就诊时间, cboSelectTime), InStr(";" & mstrPrivs & ";", ";参数设置;") > 0))
        If intDay > 7 Then intDay = 7
        mvCondFilter.Begin = Format(mvCondFilter.End - intDay, "yyyy-MM-dd 00:00:00")
        cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
        lbl就诊时间.ToolTipText = cboSelectTime.ToolTipText
        If intDay = 7 And intStart = 0 Then
            cboSelectTime.ListIndex = 1
                ElseIf intDay = 0 And intStart = 0 Then
                        cboSelectTime.ListIndex = 0
        Else
            cboSelectTime.ListIndex = 4
        End If
    End If
    
    '缺省医生本人
    mvCondFilter.医生 = UserInfo.姓名
    
    '其他不缺省
    mvCondFilter.挂号单 = ""
    mvCondFilter.就诊卡 = ""
    mvCondFilter.科室ID = 0
    mvCondFilter.门诊号 = ""
    mvCondFilter.姓名 = ""
    
End Sub

Private Sub GetLocalSetting()
'功能：从注册表读取出院病人的时间范围
    '接诊范围：1=挂本人号的病人,2=本诊室病人,3=本科室病人
    Dim strSQL As String, rsTmp As Recordset, intType As Integer
    Dim str病人接诊控制 As String '问题号:57566
    
    mint接诊范围 = Val(zlDatabase.GetPara("接诊范围", glngSys, p门诊医生站, "2"))
    mstr接诊诊室 = zlDatabase.GetPara("本地诊室", glngSys, p门诊医生站)
    mlng接诊科室ID = Val(zlDatabase.GetPara("接诊科室", glngSys, p门诊医生站))
    On Error GoTo errH
    strSQL = "Select Distinct B.ID,B.编码,B.名称,A.缺省" & _
        " From 部门人员 A,部门表 B,部门性质说明 C" & _
        " Where A.部门ID=B.ID And B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And A.人员ID=[1] And b.ID=[2]" & _
        " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, mlng接诊科室ID)
    If rsTmp.RecordCount = 0 Then mlng接诊科室ID = 0
    mbln要求分诊 = Val(zlDatabase.GetPara("只接收已经分诊的病人", glngSys, p门诊医生站)) <> 0
    
    '续诊病人
    If InStr(mstrPrivs, "续诊病人") > 0 Then
        mstr接诊医生 = zlDatabase.GetPara("接诊医生", glngSys, p门诊医生站, UserInfo.姓名)
    Else
        mstr接诊医生 = UserInfo.姓名
    End If
    
    '自动化参数
    mbln自动接诊 = Val(zlDatabase.GetPara("找到病人后自动接诊", glngSys, p门诊医生站)) <> 0
    mlng自动进行 = Val(zlDatabase.GetPara("接诊后自动进行", glngSys, p门诊医生站))
    
    '医生主动呼叫后才允许接诊
    mbln呼叫后接诊 = Val(zlDatabase.GetPara("医生主动呼叫后才允许接诊", glngSys, p门诊医生站)) <> 0
    '字体设置
    mbytSize = zlDatabase.GetPara("字体", glngSys, p门诊医生站, "0")


    mintFindType = Val(zlDatabase.GetPara("病人查找方式", glngSys, p门诊医生站, "1", , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0)
    
    '问题号:57566
    str病人接诊控制 = CStr(zlDatabase.GetPara("病人接诊控制", glngSys, p门诊医生站))
    If str病人接诊控制 <> "" Then
        mlng接诊控制 = Val(Left(str病人接诊控制, 1))
        If UBound(Split(str病人接诊控制, "|")) >= 1 Then
            mlng提前接收时间 = Val(Split(str病人接诊控制, "|")(1))
        End If
    End If
    
    '设置自动刷新
    Call SetTimer
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Function LoadPatients(Optional ByVal strRefesh As String = "111111", _
    Optional ByVal intActive As PatiType = -1, Optional ByVal strActNO As String) As Boolean
'功能：读取病人列表
'参数：intActive,strActNO=刷新后想要定位的列表索引和病人挂号单(如果有)
'      注意其中如果指定了intActive,则必须要包含strRefesh刷新列表中
'      strRefesh=分别是否刷新指定的列表，分别为"候诊，就诊，已诊，转诊，预约,回诊"
    Dim rsPati As New ADODB.Recordset
    Dim objItem As ListItem, intIdx As PatiType
    Dim strKeep As String, strPrePati As String
    Dim strSQL As String, i As Long, j As Long
    Dim strTime As String, blnRefresh As Boolean
    Dim objLvw As ListView
    Dim lngColor As Long, lngPatiTypeIdx As Long
    Dim rs传染病状态 As ADODB.Recordset
    Dim blnDo传染病状态 As Boolean
    Dim bln中医 As Boolean
    
    strPrePati = mstrPrePati '因为要破坏,因此临时记录
    
    Screen.MousePointer = 11
    On Error GoTo errH
    mblnUnRefresh = True
    strSQL = "select  m.病人id,m.id,m.no,max(m.记录) as 记录,max(m.填写) as 填写,max(m.状态) as 状态 from" & vbNewLine & _
        "(select a.病人id,a.id, a.no,1 as 记录,0 as 填写,0 as 状态 from 病人挂号记录 a,疾病阳性记录 b" & vbNewLine & _
        "where a.no=b.挂号单 and a.执行状态=2 And a.执行人||''=[1] And a.记录性质=1 And a.记录状态=1" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.病人id,a.id, a.no,0 as 记录,1 as 填写,0 as 状态" & vbNewLine & _
        "From 病人挂号记录 A, 电子病历记录 C, 病历文件列表 D" & vbNewLine & _
        "Where c.文件id = d.Id And d.种类 = 5 And a.病人id = c.病人id And a.id = c.主页id and a.执行状态=2 And a.执行人||''=[1] And a.记录性质=1 And a.记录状态=1" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.病人id,a.id, a.no,0 as 记录,1 as 填写,e.处理状态 as 状态" & vbNewLine & _
        "From 病人挂号记录 A,电子病历记录 C,病历文件列表 D,疾病申报记录 E" & vbNewLine & _
        "Where a.病人id = c.病人id And a.id = c.主页id and c.id=e.文件id and d.种类=5 and e.文件id =d.id and a.执行状态=2 And a.执行人||''=[1] And a.记录性质=1 And a.记录状态=1) M" & vbNewLine & _
        "group by m.病人id,m.id,m.no"
    Set rs传染病状态 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr接诊医生)
    If rs传染病状态.RecordCount > 0 Then blnDo传染病状态 = True
    
    For intIdx = 0 To lvwPati.UBound + 3
        If Mid(strRefesh, intIdx + 1, 1) = "1" Then
            If intIdx = pt候诊 Then    '候诊病人
                '接诊范围
                If mint接诊范围 = 1 Then
                    strSQL = " And B.执行人||''=[2]" '挂本人号
                    If mbln要求分诊 Then strSQL = strSQL & " And B.诊室 is Not NULL"
                ElseIf mint接诊范围 = 2 Then
                    '本诊室
                    If mlng接诊科室ID <> 0 Then
                        strSQL = " And B.诊室=[3] And b.执行部门id+0 =[4] And (B.执行人||''=[2] Or B.执行人 Is Null) "
                    Else    '10.28以前选诊室时没有定科室
                        strSQL = " And B.诊室=[3] And (B.执行人||''=[2] Or B.执行人 Is Null) " & _
                            "And Exists (Select 科室id" & vbNewLine & _
                            " From 挂号安排 F, 部门人员 D" & vbNewLine & _
                            " Where D.人员id = [6] And F.科室id = D.部门id And b.执行部门id = F.科室id)"
                    End If
                ElseIf mint接诊范围 = 3 Then
                    strSQL = " And B.执行部门ID+0=[4] And (B.执行人||''=[2] Or B.执行人 Is Null)" '本科室
                    If mbln要求分诊 Then strSQL = strSQL & " And B.诊室 is Not NULL"
                End If
                
                strSQL = _
                    " Select /*+ Rule*/B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区," & _
                    "       B.发生时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类," & _
                    "       B.号序,B.诊室,B.分诊时间,B.发生时间,B.执行部门ID,B.执行人," & _
                    "       B.转诊状态,C.名称 as 转诊科室,B.转诊诊室,B.转诊医生,B.执行状态,B.记录标志,A.病人类型" & _
                    " From 病人信息 A,病人挂号记录 B,部门表 C" & _
                    " Where B.病人ID=A.病人ID And (Nvl(B.执行状态,0)=0 or nvl(B.执行状态,0)=[5]) And B.转诊科室ID=C.ID(+) And B.记录性质=1 And B.记录状态=1" & _
                    "       And B.执行时间 is Null And B.发生时间 <= Trunc(Sysdate)+1-1/24/60/60 " & strSQL & _
                    IIf(gint普通挂号天数 = gint急诊挂号天数, " And B.发生时间>=Sysdate-" & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数), _
                    " And B.发生时间 >= Sysdate-" & IIf(gint普通挂号天数 > gint急诊挂号天数, gint普通挂号天数, gint急诊挂号天数) & " And B.发生时间>=Sysdate-Decode(B.急诊,1," & IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数) & "," & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数) & ")") & _
                    " Order By Decode(B.分诊时间,NULL,2,1),B.分诊时间,B.NO"
                '"Sysdate-Decode(B.急诊"导致索引失效，所以加了额外的条件
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "未用", UserInfo.姓名, mstr接诊诊室, IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID), IIf(mblnShowLeavePati, -1, 0), UserInfo.ID)
            ElseIf intIdx = pt就诊 Then '就诊病人
                strSQL = _
                    " Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区," & _
                    " B.执行时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类,B.发生时间,B.执行部门ID,B.执行人," & _
                    " B.转诊状态,C.名称 as 转诊科室,B.转诊诊室,B.转诊医生,B.执行状态,B.记录标志,A.病人类型" & _
                    " From 病人信息 A,病人挂号记录 B,部门表 C" & _
                    " Where B.病人ID=A.病人ID And B.转诊科室ID=C.ID(+)" & _
                    " And B.执行状态=2 and nvl(B.记录标志,0)<=1 And B.执行人||''=[1] And B.记录性质=1 And B.记录状态=1" & _
                    " Order By B.NO"
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr接诊医生)
            ElseIf intIdx = pt已诊 Then '已诊病人
                strSQL = "Select /*+ Rule*/" & vbNewLine & _
                    " Distinct(b.No), b.病人id, b.门诊号, b.姓名, b.性别, b.年龄, b.复诊, b.急诊, b.社区, b.执行时间 As 时间, a.就诊卡号, a.身份证号, a.Ic卡号, a.险类, b.发生时间, b.执行部门id," & vbNewLine & _
                    " b.执行人, b.执行状态, b.记录标志, a.病人类型," & vbNewLine & _
                    "First_Value(Decode(Sign(h.诊断类型 - 10), -1, h.诊断描述, '')) Over(Partition By h.病人id, h.主页id Order By Sign(h.诊断类型 - 10), Decode(h.记录来源, 4, 0, h.记录来源) Desc, Decode(h.诊断类型, 1, 1, 0) Desc, h.诊断次序) As 西医诊断," & vbNewLine & _
                    "First_Value(Decode(Sign(h.诊断类型 - 10), 1, h.诊断描述, '')) Over(Partition By h.病人id, h.主页id Order By -Sign(h.诊断类型 - 10), Decode(h.记录来源, 4, 0, h.记录来源) Desc, Decode(h.诊断类型, 11,11, 0) Desc, h.诊断次序) As 中医诊断" & vbNewLine & _
                    "From 病人信息 A, 病人挂号记录 B, 病人诊断记录 H" & IIf(mvCondFilter.就诊卡 <> "", ",病人医疗卡信息 C, 医疗卡类别 D", "") & vbNewLine & _
                    "Where b.病人id = a.病人id And h.病人id(+) = b.病人id And h.主页id(+) = b.id And b.执行状态 + 0 = 1 And b.记录性质 = 1 And b.记录状态 = 1" & _
                     IIf(mvCondFilter.就诊卡 <> "", " And c.病人id = a.病人id And c.卡类别id = d.Id And d.是否固定 = 1 And d.名称 = '就诊卡' ", "")
              
                If mvCondFilter.挂号单 <> "" Then
                    strSQL = strSQL & " And B.NO=[5]"
                ElseIf mvCondFilter.门诊号 <> "" Then
                    strSQL = strSQL & " And A.门诊号=[6]"
                ElseIf mvCondFilter.就诊卡 <> "" Then
                    strSQL = strSQL & " And C.卡号=[7]"
                
                Else
                    strSQL = strSQL & " And B.执行时间 Between To_Date('" & Format(mvCondFilter.Begin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mvCondFilter.End, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    strSQL = strSQL & IIf(mvCondFilter.医生 = "", "", " And B.执行人||''=[3]")
                    If mvCondFilter.科室ID <> 0 Then strSQL = strSQL & " And B.执行部门ID+0=[4]"
                                        If mvCondFilter.姓名 <> "" Then strSQL = strSQL & " And A.姓名=[8]"
                End If
                
                If zlDatabase.DateMoved(mvCondFilter.Begin) Then
                    strSQL = strSQL & " Union ALL " & Replace(strSQL, "病人挂号记录", "H病人挂号记录")
                End If

                strSQL = strSQL & " Order By NO Desc"
                
                With mvCondFilter
                    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "未用", "未用", .医生, .科室ID, .挂号单, .门诊号, .就诊卡, .姓名)
                End With
            ElseIf intIdx = pt回诊 Then    '回诊病人
                strSQL = _
                    " Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区," & _
                    " B.执行时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类,B.发生时间,B.执行部门ID,B.执行人," & _
                    " B.转诊状态,C.名称 as 转诊科室,B.转诊诊室,B.转诊医生,B.执行状态,B.记录标志,A.病人类型" & _
                    " From 病人信息 A,病人挂号记录 B,部门表 C" & _
                    " Where B.病人ID=A.病人ID And B.转诊科室ID=C.ID(+) And B.记录性质=1 And B.记录状态=1" & _
                    " And B.执行状态=2 and nvl(B.记录标志,0) in (2,3) And B.执行人||''=[1]" & _
                    " Order By B.NO"
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr接诊医生)
            ElseIf intIdx = lvwPati.UBound + 1 Then   '转诊病人
                '接诊范围
                If mint接诊范围 = 1 Then
                    strSQL = " And B.转诊医生=[2]" '转本人号
                ElseIf mint接诊范围 = 2 Then
                    '转本诊室：不是自已转的，接收医生是自已或者未指定接收医生
                    strSQL = " And B.转诊诊室=[3] And B.转诊科室ID=[4] And Nvl(B.执行人,'无')<>[2] And (B.转诊医生=[2] Or B.转诊医生 Is NULL)"
                ElseIf mint接诊范围 = 3 Then
                    '转本科室：不是自已转的，接收医生是自已或者未指定接收医生
                    strSQL = " And B.转诊科室ID=[4] And Nvl(B.执行人,'无')<>[2] And (B.转诊医生=[2] Or B.转诊医生 Is NULL)"
                End If
                strSQL = _
                    " Select /*+ Rule*/B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区,B.执行人," & _
                    " B.发生时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类,B.发生时间,B.转诊科室ID as 执行部门ID," & _
                    " B.转诊状态,C.名称 as 转诊科室,B.诊室 as 转诊诊室,B.执行人 as 转诊医生,B.执行状态,B.记录标志,A.病人类型" & _
                    " From 病人信息 A,病人挂号记录 B,部门表 C" & _
                    " Where B.病人ID=A.病人ID And B.转诊状态=0 And B.执行部门ID=C.ID And B.记录性质=1 And B.记录状态=1" & strSQL & _
                    IIf(gint普通挂号天数 = gint急诊挂号天数, " And B.发生时间>=Sysdate-" & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数), _
                    " And B.发生时间 >= Sysdate-" & IIf(gint普通挂号天数 > gint急诊挂号天数, gint普通挂号天数, gint急诊挂号天数) & " And B.发生时间>=Sysdate-Decode(B.急诊,1," & IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数) & "," & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数) & ")") & _
                    " Order By B.NO"
                '"Sysdate-Decode(B.急诊"导致索引失效，所以加了额外的条件
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "未用", UserInfo.姓名, mstr接诊诊室, IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID), 0, 0)
            ElseIf intIdx = lvwPati.UBound + 2 Then   '预约病人
                '接诊范围
                If mint接诊范围 = 1 Then
                    strSQL = " And A.执行人||''=[1]" '挂本人号
                                                            
                ElseIf mint接诊范围 = 2 Or mint接诊范围 = 3 Then '本诊室（预约挂号的发药窗口填的是预号，没有诊室），本科室
                    strSQL = " And A.执行部门ID+0=[2] And (A.执行人||''=[1] Or A.执行人 Is Null)"
                End If


                '现在现在的时间段：用表联接的方式会很慢。
                strTime = _
                    "Select 时间段 From 时间段 Where" & _
                    " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                    " Between" & _
                    " Decode(Sign(开始时间-终止时间),1,'3000-01-09 '||To_Char(开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(开始时间,'HH24:MI:SS'))" & _
                    " And" & _
                    " '3000-01-10 '||To_Char(终止时间,'HH24:MI:SS'))" & _
                    " Or" & _
                    " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                    " Between" & _
                    " '3000-01-10 '||To_Char(开始时间,'HH24:MI:SS')" & _
                    " And" & _
                    " Decode(Sign(开始时间-终止时间),1,'3000-01-11 '||To_Char(终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(终止时间,'HH24:MI:SS')))"
                
                '取现在的星期数对应安排的时间段
                strTime = " And Decode(To_Char(SysDate,'D'),'1',B.周日,'2',B.周一,'3',B.周二,'4',B.周三,'5',B.周四,'6',B.周五,'7',B.周六,NULL) IN(" & strTime & ")"
                strSQL = "Select A.NO,A.病人ID,A.标识号 as 门诊号,A.姓名,A.性别,A.年龄,A.加班标志 as 急诊,A.执行人," & _
                    " A.发生时间 as 时间,C.就诊卡号,C.身份证号,C.IC卡号,C.险类,A.发生时间,A.执行部门ID,0 as 执行状态,0 as 记录标志,C.病人类型" & _
                    " From 门诊费用记录 A,挂号安排 B,病人信息 C" & _
                    " Where A.计算单位=B.号码 And A.病人ID=C.病人ID(+) And A.序号=1" & _
                    " And A.记录性质=4 And A.记录状态=0 " & strTime & strSQL & _
                    " And A.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60"
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID))
            End If
            
            '记录列表现在选中的病人
            strKeep = ""
            If intIdx = lvwPati.UBound + 1 Then '转诊病人
                If Not lvwIncept.SelectedItem Is Nothing Then
                    strKeep = lvwIncept.SelectedItem.Key
                End If
                lvwIncept.ListItems.Clear
            ElseIf intIdx = lvwPati.UBound + 2 Then '预约病人
                If Not lvwReserve.SelectedItem Is Nothing Then
                    strKeep = lvwReserve.SelectedItem.Key
                End If
                lvwReserve.ListItems.Clear
            ElseIf intIdx = pt回诊 Then   '回诊病人
                If Not lvwPatiHZ.SelectedItem Is Nothing Then
                    strKeep = lvwPatiHZ.SelectedItem.Key
                End If
                lvwPatiHZ.ListItems.Clear
            Else
                If Not lvwPati(intIdx).SelectedItem Is Nothing Then
                    strKeep = lvwPati(intIdx).SelectedItem.Key
                End If
                lvwPati(intIdx).ListItems.Clear
            End If
            For i = 1 To rsPati.RecordCount
                If intIdx = lvwPati.UBound + 1 Then '转诊病人
                    Set objItem = lvwIncept.ListItems.Add(, "_" & rsPati!NO, rsPati!NO, , intIdx + 1)
                ElseIf intIdx = lvwPati.UBound + 2 Then '预约病人
                    Set objItem = lvwReserve.ListItems.Add(, "_" & rsPati!NO, rsPati!NO, , 1)
                ElseIf intIdx = 5 Then  'pt回诊
                    Set objItem = lvwPatiHZ.ListItems.Add(, "_" & rsPati!NO, rsPati!NO, , 1)
                Else
                    Set objItem = lvwPati(intIdx).ListItems.Add(, "_" & rsPati!NO, rsPati!NO, , intIdx + 1)
                End If
                objItem.SubItems(1) = Nvl(rsPati!门诊号)
                objItem.SubItems(2) = Nvl(rsPati!姓名)
                objItem.SubItems(3) = Nvl(rsPati!性别)
                objItem.SubItems(4) = Nvl(rsPati!年龄)
                objItem.SubItems(5) = IIf(Nvl(rsPati!急诊, 0) <> 0, "急", "")
                
                If intIdx = lvwPati.UBound + 2 Then
                    '预约病人
                    objItem.SubItems(6) = Nvl(rsPati!执行人)
                    objItem.SubItems(7) = Format(rsPati!时间, "yyyy-MM-dd HH:mm")
                    objItem.SubItems(8) = Nvl(rsPati!身份证号)
                    objItem.SubItems(9) = Nvl(rsPati!就诊卡号)
                    objItem.SubItems(10) = Nvl(rsPati!病人类型)
                    lngPatiTypeIdx = 10
                Else
                    objItem.SubItems(6) = IIf(Nvl(rsPati!复诊, 0) <> 0, "复", "")
                    objItem.SubItems(7) = IIf(Nvl(rsPati!社区, 0) <> 0, "√", "")
                    If intIdx = pt候诊 Then
                        objItem.SubItems(8) = Nvl(rsPati!诊室)
                        objItem.SubItems(9) = Nvl(rsPati!执行人)
                        objItem.ListSubItems(9).Tag = Nvl(rsPati!执行状态)
                        objItem.SubItems(10) = LPAD(Nvl(rsPati!号序), 5, " ")
                        objItem.SubItems(11) = Format(rsPati!分诊时间, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(12) = Format(rsPati!时间, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(13) = Nvl(rsPati!就诊卡号)
                        objItem.SubItems(14) = Nvl(rsPati!病人类型)
                        lngPatiTypeIdx = 14
                    ElseIf intIdx = pt已诊 Then
                        objItem.SubItems(8) = Format(rsPati!时间, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(9) = Nvl(rsPati!执行人)
                        objItem.SubItems(10) = Nvl(rsPati!就诊卡号)
                        objItem.SubItems(11) = Nvl(rsPati!病人类型)
                        objItem.SubItems(12) = Nvl(rsPati!西医诊断)
                        objItem.SubItems(13) = Nvl(rsPati!中医诊断)
                        If rsPati!中医诊断 & "" <> "" Then bln中医 = True
                        lngPatiTypeIdx = 13
                    ElseIf intIdx = pt回诊 Then  '回诊
                        If Nvl(rsPati!记录标志, "0") = 2 Then
                                objItem.SmallIcon = "暂停"
                        End If
                        objItem.SubItems(8) = Format(rsPati!时间, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(10) = Nvl(rsPati!就诊卡号)
                        objItem.SubItems(11) = Nvl(rsPati!病人类型)
                        lngPatiTypeIdx = 11
                        '添加传染病状态
                        strSQL = ""
                        If blnDo传染病状态 Then
                            rs传染病状态.Filter = "no='" & rsPati!NO & "'"
                            If Not rs传染病状态.EOF Then strSQL = Get传染病状态(Val(rs传染病状态!记录 & ""), Val(rs传染病状态!填写 & ""), Val(rs传染病状态!状态 & ""))
                        End If
                        objItem.SubItems(12) = strSQL
                    Else
                        objItem.SubItems(8) = Format(rsPati!时间, "yyyy-MM-dd HH:mm")
                        objItem.SubItems(9) = Nvl(rsPati!就诊卡号)
                        objItem.SubItems(10) = Nvl(rsPati!病人类型)
                        lngPatiTypeIdx = 10
                        If intIdx = pt就诊 Then
                            '添加传染病状态
                            strSQL = ""
                            If blnDo传染病状态 Then
                                rs传染病状态.Filter = "no='" & rsPati!NO & "'"
                                If Not rs传染病状态.EOF Then strSQL = Get传染病状态(Val(rs传染病状态!记录 & ""), Val(rs传染病状态!填写 & ""), Val(rs传染病状态!状态 & ""))
                            End If
                            objItem.SubItems(12) = strSQL
                        End If
                    End If
                End If
                objItem.ListSubItems(1).Tag = Nvl(rsPati!就诊卡号)
                objItem.ListSubItems(2).Tag = Format(rsPati!发生时间, "yyyy-MM-dd HH:mm:ss")
                objItem.ListSubItems(3).Tag = Nvl(rsPati!执行部门ID, 0)
                objItem.ListSubItems(4).Tag = Nvl(rsPati!执行人)
                objItem.ListSubItems(5).Tag = "" '后面记录转诊状态
                objItem.ListSubItems(6).Tag = Nvl(rsPati!身份证号)
                objItem.ListSubItems(7).Tag = Nvl(rsPati!IC卡号)
                objItem.ListSubItems(8).Tag = Val(Nvl(rsPati!记录标志))
                objItem.Tag = rsPati!病人ID
                
                '转诊状态:显示在最后一列
                If intIdx = pt候诊 Or intIdx = pt就诊 Then
                    If intIdx = pt就诊 Then
                        j = lvwPati(intIdx).ColumnHeaders.Count - 2
                    Else
                        j = lvwPati(intIdx).ColumnHeaders.Count - 1
                    End If
                
                    objItem.ListSubItems(5).Tag = Nvl(rsPati!转诊状态) 'Null和0不同
                    If Not IsNull(rsPati!转诊状态) Then
                        If rsPati!转诊状态 = 0 Then
                            '已经转诊
                            objItem.SmallIcon = "转诊"
                            objItem.SubItems(j) = "待对方接收,科室:" & rsPati!转诊科室 & _
                                IIf(Not IsNull(rsPati!转诊诊室), ",诊室:" & Nvl(rsPati!转诊诊室), "") & _
                                IIf(Not IsNull(rsPati!转诊医生), ",医生:" & Nvl(rsPati!转诊医生), "")
                        ElseIf rsPati!转诊状态 = -1 Then
                            '已拒绝转诊
                            objItem.SmallIcon = "拒绝"
                            objItem.SubItems(j) = "对方已拒绝,科室:" & rsPati!转诊科室 & _
                                IIf(Not IsNull(rsPati!转诊诊室), ",诊室:" & Nvl(rsPati!转诊诊室), "") & _
                                IIf(Not IsNull(rsPati!转诊医生), ",医生:" & Nvl(rsPati!转诊医生), "")
                        ElseIf rsPati!转诊状态 = 1 Then
                            '已接收转诊
                        End If
                    End If
                ElseIf intIdx = lvwPati.UBound + 1 Then
                    '转诊病人
                    objItem.SmallIcon = "候诊"
                    objItem.SubItems(lvwIncept.ColumnHeaders.Count - 1) = "待接收转诊,科室:" & rsPati!转诊科室 & _
                        IIf(Not IsNull(rsPati!转诊诊室), ",诊室:" & Nvl(rsPati!转诊诊室), "") & _
                        IIf(Not IsNull(rsPati!转诊医生), ",医生:" & Nvl(rsPati!转诊医生), "")
                End If
                
                '显示病人颜色
                lngColor = zlDatabase.GetPatiColor(Nvl(rsPati!病人类型))
                objItem.ListSubItems(1).ForeColor = lngColor
                objItem.ListSubItems(lngPatiTypeIdx).ForeColor = lngColor
                
                '保险病人用红色显示
                If Not IsNull(rsPati!险类) And rsPati!病人类型 & "" = "" Then
                    objItem.ListSubItems(1).ForeColor = &HC0&
                    objItem.ListSubItems(lngPatiTypeIdx).ForeColor = &HC0&
                End If
                
                '急诊标志红色突出显示
                If Nvl(rsPati!急诊, 0) <> 0 Then
                    objItem.ListSubItems(5).ForeColor = vbRed
                End If
                
                '定位到指定病人
                If objItem.Key = "_" & strActNO Then
                    mstrPrePati = "_" & strActNO '避免激活事件
                    objItem.Selected = True
                    strKeep = ""
                End If
                '定位到原先病人
                If objItem.Key = strKeep And Me.Visible Then
                    mstrPrePati = strKeep '避免激活事件
                    objItem.Selected = True
                End If
                  If intIdx = pt候诊 Then
                        If Val(objItem.ListSubItems(9).Tag) = -1 Then
                            objItem.ForeColor = &H808080
                            For j = 1 To objItem.ListSubItems.Count
                                objItem.ListSubItems(j).ForeColor = &H808080
                            Next
                        End If
                  End If
                rsPati.MoveNext
            Next
            
            '中医诊断为空时隐藏
            If intIdx = pt已诊 Then
                lvwPati(pt已诊).ColumnHeaders(14).Width = IIf(bln中医, 3000, 0)
            End If
            
            '刷新诊室忙闲状态
            If intIdx = pt就诊 Then
                Call SetRoomState(lvwPati(intIdx).ListItems.Count > 0)
            End If
        End If
    Next
    
    '更新列表标题
    Call RefreshTitle
    
    '如果当前活动列表未要求刷新且有数据,则不重复处理刷新
    blnRefresh = True
    If mintActive <> -1 Then
        If mintActive = pt回诊 Then
            '回诊
            Set objLvw = lvwPatiHZ
        ElseIf mintActive = pt预约 Then
            Set objLvw = lvwReserve
        ElseIf mintActive = pt转诊 Then
            Set objLvw = lvwIncept
        Else
            Set objLvw = lvwPati(mintActive)
        End If
        If Mid(strRefesh, mintActive + 1, 1) = "0" _
            And Not objLvw.SelectedItem Is Nothing Then
            If objLvw.SelectedItem.Key = strPrePati Then
                blnRefresh = False
                mstrPrePati = strPrePati
            End If
        End If
    End If
        
    '确定刷新后活动病人列表并刷新子数据:缺省不为已诊病人列表
    If blnRefresh Then
        If intActive = -1 Then intActive = mintActive
        If intActive = pt回诊 Then
            '回诊
            Set objLvw = lvwPatiHZ
        ElseIf intActive = pt预约 Then
            Set objLvw = lvwReserve
        ElseIf intActive = pt转诊 Then
            Set objLvw = lvwIncept
        ElseIf intActive <> -1 Then
            Set objLvw = lvwPati(intActive)
        End If
            
        If intActive = -1 Then
            If lvwPati(pt候诊).ListItems.Count > 0 Then
                intActive = 0
            ElseIf lvwPati(pt就诊).ListItems.Count > 0 Then
                intActive = 1
            End If
        ElseIf objLvw.ListItems.Count = 0 Then
            If lvwPati(pt候诊).ListItems.Count > 0 Then
                intActive = 0
            ElseIf lvwPati(pt就诊).ListItems.Count > 0 Then
                intActive = 1
            Else
                intActive = -1
            End If
        End If
        
        If intActive = pt回诊 Then
            '回诊
            Set objLvw = lvwPatiHZ
        ElseIf intActive = pt预约 Then
            Set objLvw = lvwReserve
        ElseIf intActive = pt转诊 Then
            Set objLvw = lvwIncept
        ElseIf intActive <> -1 Then
            Set objLvw = lvwPati(intActive)
        End If
        
        '刷新病人的相关数据
        mintActive = -1
        '默认设置一个值
        If Not Me.Visible Then mintActive = 1
        mstrPrePati = ""
        If intActive <> -1 And Me.Visible Then
            objLvw.SelectedItem.EnsureVisible
            Call LvwItemClick(CInt(intActive), objLvw.SelectedItem)
            'Call lvwPati_ItemClick(CInt(intActive), lvwPati(intActive).SelectedItem)
        Else
            
            '按当前列表无数据刷新子窗体
            Call ClearPatiInfo
            Call SubWinRefreshData(tbcSub.Selected)
        End If
    End If
    Screen.MousePointer = 0
    LoadPatients = True
    mblnUnRefresh = False
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnUnRefresh = False
End Function

Private Sub RefreshTitle()
'功能：更新列表标题
    Dim i As Integer
    
    For i = 1 To dkpMain.PanesCount
        If dkpMain.Panes(i).Title Like "候诊病人*" Then
            dkpMain.Panes(i).Title = "候诊病人" & IIf(lvwPati(pt候诊).ListItems.Count = 0, "", ":" & lvwPati(pt候诊).ListItems.Count & "人")
        ElseIf dkpMain.Panes(i).Title Like "*就诊病人*" Then
            If mstr接诊医生 <> UserInfo.姓名 Then
                dkpMain.Panes(i).Title = mstr接诊医生 & "的就诊病人" & IIf(lvwPati(pt就诊).ListItems.Count = 0, "", ":" & lvwPati(pt就诊).ListItems.Count & "人")
            Else
                dkpMain.Panes(i).Title = UserInfo.姓名 & "的就诊病人" & IIf(lvwPati(pt就诊).ListItems.Count = 0, "", ":" & lvwPati(pt就诊).ListItems.Count & "人")
            End If
        ElseIf dkpMain.Panes(i).Title Like "已诊病人*" Then
            dkpMain.Panes(i).Title = "已诊病人" & IIf(lvwPati(pt已诊).ListItems.Count = 0, "", ":" & lvwPati(pt已诊).ListItems.Count & "人")
        ElseIf dkpMain.Panes(i).Title Like "转诊病人*" Then
            dkpMain.Panes(i).Title = "转诊病人" & IIf(lvwIncept.ListItems.Count = 0, "", ":" & lvwIncept.ListItems.Count & "人")
        ElseIf dkpMain.Panes(i).Title Like "预约病人*" Then
            dkpMain.Panes(i).Title = "预约病人" & IIf(lvwReserve.ListItems.Count = 0, "", ":" & lvwReserve.ListItems.Count & "人")
        ElseIf dkpMain.Panes(i).Title Like "回诊病人*" Then
            dkpMain.Panes(i).Title = "回诊病人" & IIf(lvwPatiHZ.ListItems.Count = 0, "", ":" & lvwPatiHZ.ListItems.Count & "人")
        End If
    Next
End Sub

Private Sub ClearPatiInfo()
'功能：清除单个病人相关的显示信息
    Dim i As Long
    
    cboRegist.Tag = "Loading"
    mlng病人ID = 0
    mstr挂号单 = ""
    mlng科室ID = 0
    mlng挂号ID = 0
    mPatiInfo.类型 = 0
    mPatiInfo.门诊号 = ""
    mPatiInfo.挂号单 = ""
    mPatiInfo.挂号ID = 0
    mPatiInfo.科室ID = 0
    mPatiInfo.诊室 = ""
    mPatiInfo.社区 = 0
    mPatiInfo.社区号 = ""
    mPatiInfo.挂号时间 = CDate(0)
    mPatiInfo.数据转出 = False
    mPatiInfo.病历文件id = 0
    mPatiInfo.病历id = 0
    mPatiInfo.是否签名 = False
    mPatiInfo.保存人 = ""
    mPatiInfo.婚姻状况 = ""
    mPatiInfo.性别 = ""
    mPatiInfo.民族 = ""
    mPatiInfo.国籍 = ""
    mPatiInfo.区域 = ""
    mPatiInfo.出生地点 = ""
    mPatiInfo.传染病上传 = 0
    mPatiInfo.家庭地址邮编 = ""
    mPatiInfo.单位邮编 = ""
    mPatiInfo.其他证件 = ""
    mPatiInfo.病人ID = 0
        
    cboRegist.Clear
    lbl多科就诊.Visible = False
    lbl急.Visible = False
    lblRec.Visible = False
            
    For i = 0 To txtEdit.Count - 1
        txtEdit(i).Text = ""
    Next
    If mblnDocInput Then
        For i = 0 To rtfEdit.Count - 1
            rtfEdit(i).Text = ""
        Next
    End If
    txt出生日期.Text = "____-__-__"
    txt发病日期.Text = "____-__-__"
    txt出生时间.Text = "__:__"
    txt发病时间.Text = "__:__"
            
    For i = 0 To lblShow.Count - 1
        lblShow(i).Caption = ""
    Next
    
    vsAller.Clear: vsAller.Cols = 1
    cmdAller.Enabled = False
    lblDiag(1).Caption = ""
    
    Call SetPermitEdit
    cboRegist.Tag = ""
End Sub

Private Sub ExecuteRegist()
'功能：病人挂号
    Dim strCommon As String, intAtom As Integer
    Dim strNO As String, blnPrice As Boolean
    Dim objControl As CommandBarControl
    
    blnPrice = Val(zlDatabase.GetPara("允许挂号划价单", glngSys, p门诊医生站, 1)) = 1
    
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If gobjRegist Is Nothing Then Exit Sub
    End If
    err.Clear: On Error GoTo 0
    
    mblnUnRefresh = True
    '部件调用(处理合法性设置)
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    strNO = gobjRegist.StationRegister(Me, gcnOracle, glngSys, mstr接诊诊室, InStr(mstrPrivs, "挂号费别打折") = 0, blnPrice, , gstrDBUser)
    Call GlobalDeleteAtom(intAtom)
        
    '刷新并定位到刚挂号的病人上
    If strNO <> "" And lvwPati(pt就诊).Visible Then
        Call LoadPatients("11000", pt就诊, strNO)
        lvwPati(pt就诊).SetFocus
        
        '接诊之后自动进行医嘱下达状态
        If mlng自动进行 = 1 Then
            If tbcSub.Selected.Tag <> "医嘱" Then tbcSub.Item(0).Selected = True
            cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
            Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        ElseIf mlng自动进行 = 2 Then
            If tbcSub.Selected.Tag <> "病历" Then tbcSub.Item(1).Selected = True
            cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
            mblnUnRefresh = True
            Call mclsEPRs.zlOpenDefaultEPR(mstr挂号单)
        End If
    Else
        Call LoadPatients("11000")
    End If
    mblnUnRefresh = False
End Sub

Private Sub ExecuteBespeak()
'功能：预约挂号
    Dim strCommon As String, intAtom As Integer, strNO As String
            
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If gobjRegist Is Nothing Then Exit Sub
    End If
    err.Clear: On Error GoTo 0
    
    mblnUnRefresh = True
    '部件调用(处理合法性设置)
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    strNO = gobjRegist.StationBespeak(Me, gcnOracle, glngSys, "", InStr(mstrPrivs, "挂号费别打折") = 0, mlng病人ID, gstrDBUser)
    Call GlobalDeleteAtom(intAtom)
    mblnUnRefresh = False
End Sub

Private Sub ExecuteBespeakPrint()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打预约挂号单
    '编制:刘兴洪
    '日期:2012-12-24 10:55:39
    '说明:
    '问题:56274
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommon As String, intAtom As Integer, strNO As String
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If err <> 0 Then
            err = 0: On Error GoTo 0
        End If
        If gobjRegist Is Nothing Then Exit Sub
    End If
    On Error GoTo errHandle
    With lvwReserve
        strNO = Trim(.SelectedItem.Text)
    End With
    If strNO = "" Then Exit Sub
    '部件调用(处理合法性设置)
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    'zlPrintBespeak(ByVal frmMain As Object, ByVal cnMain As ADODB.Connection, _
    ByVal lngSys As Long, ByVal strDbUser As String, ByVal strPrivs As String, ByVal strNO As String)
    strNO = gobjRegist.zlPrintBespeak(Me, gcnOracle, glngSys, gstrDBUser, mstrPrivs, strNO)
    Call GlobalDeleteAtom(intAtom)
    mblnUnRefresh = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecuteTransferSend()
'功能：病人转诊
    Dim rsTmp As New ADODB.Recordset
    Dim lng科室ID As Long, str诊室 As String
    Dim str医生 As String, lng医生ID As Long
    Dim strSQL As String, objLvw As ListView
    
    If mintActive = pt回诊 Then
        Set objLvw = lvwPatiHZ
    Else
        Set objLvw = lvwPati(mintActive)
    End If
    
    With objLvw.SelectedItem
        If mstr挂号单 = "" Then
            MsgBox "请先选择病人。", vbInformation, gstrSysName
            Exit Sub
        End If
        If mintActive = pt已诊 Then
            If zlDatabase.NOMoved("门诊费用记录", mstr挂号单, "记录性质=", "4") Then
                MsgBox "该病人的挂号费用已经转出到后备数据库，不允许操作。" & vbCrLf & _
                    "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '检查挂号单时限
        If BillExpend(mstr挂号单) Then
            MsgBox "该病人挂号已超过有效天数，不能再进行转诊。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        On Error GoTo errH
        
        '对正在就诊的病人的检查
        If mintActive = pt就诊 Or mintActive = pt回诊 Then
            If InStr(GetInsidePrivs(p门诊医生站), "已下医嘱转诊") > 0 Then
                '检查是否还有未发送的医嘱
                strSQL = "Select ID From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[2] And 医嘱状态=1 And Nvl(执行性质,0)<>0 And Rownum = 1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
                If Not rsTmp.EOF Then
                    MsgBox "该病人还有未发送医嘱，只有将所有医嘱发送后才能进行转诊。", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else    '只要下过医嘱(不含已作废的)，说明就诊行为已发生，不允许转诊，须重新挂号
                strSQL = "Select ID From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[2] And 医嘱状态 <> 4 And Rownum = 1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
                If Not rsTmp.EOF Then
                    MsgBox "已经对该病人下过医嘱，不允许转诊，请删除或作废医嘱后再进行，或者重新挂号。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        If Not frmRegistPlan.ShowMe(Me, mstr挂号单, lng科室ID, str诊室, str医生, lng医生ID) Then mblnUnRefresh = False: Exit Sub
        
        '执行转诊
        strSQL = "Zl_病人挂号记录_转诊('" & mstr挂号单 & "',0," & lng科室ID & ",'" & str诊室 & "','" & str医生 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        '门诊患者转诊消息发送
        Call ZLHIS_CIS_007(mclsMipModule, mlng病人ID, Trim(txtEdit(txt姓名).Text), mPatiInfo.门诊号, mPatiInfo.挂号ID, mlng接诊科室ID, , lng科室ID, , lng医生ID, str医生, str诊室, UserInfo.姓名)
        
        Call zlShowQuence(mstr挂号单)
        '刷新界面
        Call LoadPatients("11011")
    End With
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlShowQuence(ByVal strNO As String)
    '功能:显示排队叫号队列的新号
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If Check排队叫号 = False Then Exit Sub
    strSQL = "Select 排队号码 From 排队叫号队列 Where 业务类型=0 and 业务ID in (Select ID From 病人挂号记录 where NO=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    MsgBox "注意:" & vbCrLf & "    该病人重新进行了排队处理,队号为:[ " & Nvl(rsTemp!排队号码) & " ]", vbInformation + vbOKOnly, gstrSysName
End Sub

Private Sub ExecuteTransferRefuse()
'功能：转诊拒绝
    Dim strSQL As String
        
    On Error GoTo errH
    
    With lvwIncept.SelectedItem
        If MsgBox("确实要拒绝该转诊病人""" & .SubItems(2) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        strSQL = "Zl_病人挂号记录_转诊('" & .Text & "',-1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End With
    '刷新界面
    Call LoadPatients("11011")
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteTransferCancel(Optional ByVal blnMsg As Boolean = True)
'功能：取消转诊
    Dim strSQL As String
    Dim objLvw As ListView
    On Error GoTo errH
    If mintActive = pt回诊 Then
        Set objLvw = lvwPatiHZ
    Else
        Set objLvw = lvwPati(mintActive)
    End If
    With objLvw.SelectedItem
        If blnMsg Then
            If MsgBox("确实要取消病人""" & .SubItems(2) & """的转诊吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        End If
        strSQL = "Zl_病人挂号记录_转诊('" & .Text & "',Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End With
    
    '刷新界面
    Call LoadPatients("11011")
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteTransferIncept()
'功能：接收转诊
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lvwIncept.SelectedItem Is Nothing Then Exit Sub
    
    With lvwIncept.SelectedItem
        If MsgBox(.SubItems(lvwIncept.ColumnHeaders.Count - 1) & vbCrLf & vbCrLf & "确认接收该转诊病人吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
 
        strSQL = "Zl_病人挂号记录_转诊('" & .Text & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        If HaveRIS Then
            If gobjRis.HISModPati(1, mlng病人ID, mlng挂号ID) <> 1 Then
                MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
            End If
        ElseIf gbln启用影像信息系统接口 = True Then
            MsgBox "当前启用了影像信息系统接口，但于由RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
        End If
        Call mclsAdvices.zlRefresh(0, "", False)
        '刷新并定位病人
        If lvwPati(pt就诊).Visible Then
            Call LoadPatients("11011", pt就诊, .Text)
            lvwPati(pt就诊).SetFocus
        Else
            Call LoadPatients("11011")
        End If
    End With
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteReceive(Optional ByVal blnIsCard As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人接诊
    '参数:blnIsCard-是否是刷卡调用接收预约病人
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
    Dim strSQL As String, strNO As String
    Dim blnReserve As Boolean
    Dim datCurr As Date
   
    On Error GoTo errH
    If lvwPati(pt候诊).Visible And lvwReserve.Visible Then
        blnReserve = Me.ActiveControl Is lvwReserve
    Else
        blnReserve = lvwReserve.Visible
    End If
    datCurr = zlDatabase.Currentdate
    If blnReserve Then
        '对预约挂号病人进行接诊
        If lvwReserve.SelectedItem Is Nothing Then Exit Sub
        
        '问题号:57566
        If Check接诊控制("接诊", Mid(lvwReserve.SelectedItem.Key, 2)) = False Then Exit Sub
        
        '门诊医生站预约接收时调用挂号部件的接收接口进行扣费的功能
        If Val(zlDatabase.GetPara("允许挂号划价单", glngSys, p门诊医生站, 1)) <> 1 And Not mobjSquareCard Is Nothing Then
            If Not mobjSquareCard.zlRegisterIncept(Me, mlngModul, Mid(lvwReserve.SelectedItem.Key, 2), mstr接诊诊室, PatiIdentify.objIDKind.GetCurCard.接口序号, PatiIdentify.Text) Then Exit Sub
        Else
            With lvwReserve.SelectedItem
                strNO = Mid(.Key, 2)
                strSQL = "Zl_病人预约挂号_接收('" & strNO & "','" & mstr接诊诊室 & "',NULL,NULL,NULL,NULL,NULL,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End With
        End If
        
        '刷新并定位病人
        On Error GoTo 0
        If lvwPati(pt就诊).Visible Then
            Call LoadPatients("11001", pt就诊, strNO)
            lvwPati(pt就诊).SetFocus
        Else
            Call LoadPatients("11001")
        End If
    Else
        '问题号:57566
        If Check接诊控制("接诊", mstr挂号单) = False Then Exit Sub
        '对正常挂号病人进行接诊
        strSQL = "Select 执行人 From 病人挂号记录 Where 病人ID+0=[1] And NO=[2] And Nvl(执行状态,0)<>0 And 记录性质=1 And 记录状态=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
        If Not rsTmp.EOF Then
            MsgBox "该病人已由" & IIf(IsNull(rsTmp!执行人), "其他医生", "医生：" & rsTmp!执行人 & " ") & "接诊。", vbInformation, gstrSysName
            Call LoadPatients("100"): Exit Sub
        End If
        
        strSQL = "Select 执行人 From 病人挂号记录 Where 病人ID+0=[1] And NO=[2] And Nvl(执行状态,0)=0 And 记录性质=1 And 记录状态=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
        If rsTmp.EOF Then
            MsgBox "该病人已退号，不能接诊。", vbInformation, gstrSysName
            Call LoadPatients("100"): Exit Sub
        End If
        
        strSQL = "zl_病人接诊(" & mlng病人ID & ",'" & mstr挂号单 & "',Null,'" & UserInfo.姓名 & "','" & mstr接诊诊室 & "',0,0,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        '刷新并定位病人
        On Error GoTo 0
        If lvwPati(pt就诊).Visible Then
            Call LoadPatients("110", pt就诊, mstr挂号单)
            lvwPati(pt就诊).SetFocus
        Else
            Call LoadPatients("110")
        End If
    End If
    '门诊患者接诊消息发送
    Call ZLHIS_CIS_009(mclsMipModule, mlng病人ID, Trim(txtEdit(txt姓名).Text), mPatiInfo.门诊号, Val(ucPatiVitalSigns.value身高), Val(ucPatiVitalSigns.value体重), mlng挂号ID, IIf(optState(opt复诊).Value, 1, 0), IIf(lbl急.Visible, 1, 0), datCurr, mlng接诊科室ID, , mstr接诊诊室, UserInfo.姓名)

    '社区病人自动调用功能
    If Not gobjCommunity Is Nothing And mlngCommunityID <> 0 And mlng病人ID <> 0 And mPatiInfo.社区 <> 0 Then
        Set objControl = cbsMain.FindControl(, mlngCommunityID, , True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
    Call CreatePlugInOK(p门诊医生站)
    '接诊后调用外挂接口
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicReceive(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID)
        Call zlPlugInErrH(err, "ClinicReceive")
        err.Clear: On Error GoTo errH
    End If
    
    '接诊之后自动进行医嘱下达状态
    If mlng自动进行 = 1 Then
        If tbcSub.Selected.Tag <> "医嘱" Then tbcSub.Item(0).Selected = True
        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    ElseIf mlng自动进行 = 2 Then
        If tbcSub.Selected.Tag <> "病历" Then tbcSub.Item(1).Selected = True
        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
        mblnUnRefresh = True
        Call mclsEPRs.zlOpenDefaultEPR(mstr挂号单)
    End If
    '处理排队叫号队列(重新刷新)
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteCancel()
'功能：取消接诊
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If BillExpend(mstr挂号单) Then
        MsgBox "该病人挂号已超过有效天数，不允许再取消接诊。", vbInformation, gstrSysName
        Exit Sub
    End If
        
    On Error GoTo errH
    
    '只能取消自己接诊的病人
    strSQL = "Select 执行人 From 病人挂号记录 Where id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.挂号ID)
    If rsTmp!执行人 <> UserInfo.姓名 Then
        MsgBox "只能取消自己接诊的病人。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ToDo:取消接诊时病历数据的检查
    '医嘱数据的检查
    strSQL = "Select Count(*) as 医嘱 From 病人医嘱记录 Where 医嘱状态 IN(1,8) And 病人ID+0=[1] And 挂号单=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
    If Nvl(rsTmp!医嘱, 0) > 0 Then
        MsgBox "该病人已有新开或已发送的医嘱，不能取消接诊。" & vbCrLf & _
            "如果确实要取消接诊，请先将这些医嘱删除或作废。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSQL = "Zl_病人接诊_Cancel(" & mlng病人ID & ",'" & mstr挂号单 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '刷新并定位病人
    If lvwPati(pt候诊).Visible Then
        Call LoadPatients("110", pt候诊, mstr挂号单)
        lvwPati(pt候诊).SetFocus
    Else
        Call LoadPatients("110")
    End If
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteFinish()
'功能：完成接诊
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnTran As Boolean
    Dim str疾病IDs As String, str诊断IDs As String
    Dim lng挂号id As Long
    Dim objLvw As ListView
    
    On Error GoTo errH
    If mintActive = pt回诊 Then
        Set objLvw = lvwPatiHZ
    Else
        Set objLvw = lvwPati(pt就诊)
    End If
    
    If objLvw.SelectedItem Is Nothing Then Exit Sub
    '如果列表长时间不刷新并发操作检查
    strSQL = "select 1 from 病人挂号记录 where no=[1] and 执行人=[2] And 执行状态=2 And 记录性质=1 And 记录状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单, mstr接诊医生)
    If rsTmp.EOF Then
        MsgBox """" & objLvw.SelectedItem.SubItems(2) & """可能被其他医生强制续诊接收，请重试。", vbInformation, gstrSysName
        Call LoadPatients
        Call ReshDataQueue
        Exit Sub
    End If
    'ToDo:完成接诊时病历数据的检查
    
    If objLvw.SelectedItem.ListSubItems(5).Tag = "0" Then
        If MsgBox("当前病人""" & objLvw.SelectedItem.SubItems(2) & """已经转诊，是否要取消转诊后再完成接诊？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            Call ExecuteTransferCancel(False)
            Call ExecuteFinish
            Exit Sub
        End If
    End If
    '检查是否存在有效医嘱
    strSQL = "Select Count(*) as 医嘱 From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[2] And 医嘱状态<>4"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
    If Nvl(rsTmp!医嘱, 0) = 0 Then
        If MsgBox("未对""" & objLvw.SelectedItem.SubItems(2) & """下达任何有效的医嘱，确实要完成接诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    '检查是否存在未发送的医嘱
    strSQL = "Select Count(*) as 医嘱 From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[2] And 医嘱状态=1 And Nvl(执行性质,0)<>0 And Nvl(皮试结果,'无')<>'免试'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
    If Nvl(rsTmp!医嘱, 0) > 0 Then
        MsgBox """" & objLvw.SelectedItem.SubItems(2) & """还有未发送的医嘱，不能完成接诊。", vbInformation, gstrSysName
        Exit Sub
    End If
    '检查未填写的疾病证明报告
    strSQL = "Select 主页ID,疾病ID,诊断ID From 病人诊断记录 Where 取消时间 is Null And 病人ID=[1] And 主页ID=(Select ID From 病人挂号记录 Where NO=[2] And 记录性质=1 And 记录状态=1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
    Do While Not rsTmp.EOF
        If lng挂号id = 0 Then lng挂号id = rsTmp!主页ID
        If Not IsNull(rsTmp!疾病id) Then str疾病IDs = str疾病IDs & "," & rsTmp!疾病id
        If Not IsNull(rsTmp!诊断id) Then str诊断IDs = str诊断IDs & "," & rsTmp!诊断id
        rsTmp.MoveNext
    Loop
    If str疾病IDs <> "" Or str诊断IDs <> "" Then
        If InStr(";" & GetPrivFunc(glngSys, p门诊病历管理) & ";", ";病历书写;") > 0 Then
            If Not CheckDiseaseFile(Me, mlng病人ID, lng挂号id, mlng接诊科室ID, Mid(str疾病IDs, 2), Mid(str诊断IDs, 2), , True) Then Exit Sub
        End If
    End If
    
    '读取必要的信息供社区接口调用:以左边就诊病人本次就诊为准,右边可能当前选择的历史就诊
    strSQL = "Select A.ID,A.社区,B.社区号 From 病人挂号记录 A,病人社区信息 B Where A.病人ID=B.病人ID(+) And A.记录性质=1 And A.记录状态=1 And A.社区=B.社区(+) And A.NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
    
    '执行过程
    '-----------------------------------
    gcnOracle.BeginTrans: blnTran = True
    
    strSQL = "Zl_病人接诊完成(" & mlng病人ID & ",'" & mstr挂号单 & "','" & mstr接诊诊室 & "','" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
         
    If Not gobjCommunity Is Nothing And Nvl(rsTmp!社区, 0) <> 0 Then
        '调用社区病人信息提交
        If Not gobjCommunity.ClinicSubmit(glngSys, mlngModul, rsTmp!社区, Nvl(rsTmp!社区号), mlng病人ID, rsTmp!ID) Then
            gcnOracle.RollbackTrans: blnTran = False: Exit Sub
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False

    '接诊后调用外挂接口
    Call CreatePlugInOK(p门诊医生站)
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicFinish(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID)
        Call zlPlugInErrH(err, "ClinicFinish")
        err.Clear: On Error GoTo errH
    End If
    
    '一卡通数据上传
    If Not mobjICCard Is Nothing Then
        strSQL = "Select 1 From 一卡通目录 Where 启用=2 And Rownum=1"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            mobjICCard.UploadSwap mlng病人ID, ""
        End If
    End If
    '刷新:不定位到已诊列表
    Call LoadPatients
    Call ReshDataQueue
    
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteRedo()
'恢复接诊
    Dim strSQL As String
    
    '只检查在线数据表中的
    If BillExpend(mstr挂号单) Then
        MsgBox "该病人挂号已超过有效天数，不允许再恢复接诊。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mintActive = pt已诊 Then
        If zlDatabase.NOMoved("病人挂号记录", mstr挂号单) Then
            MsgBox "该挂号记录已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '当前医生完成的病人才可以直接恢复(否则有权限可用强制续诊)
    With lvwPati(pt已诊).SelectedItem
        If .ListSubItems(4).Tag <> UserInfo.姓名 Then
            MsgBox "该病人不是由你完成就诊的，不能直接恢复接诊。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    On Error GoTo errH
    strSQL = "zl_病人接诊完成_Cancel(" & mlng病人ID & ",'" & mstr挂号单 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    '刷新并定位病人
    If lvwPati(pt就诊).Visible Then
        Call LoadPatients("011001", pt就诊, mstr挂号单)
        lvwPati(pt就诊).SetFocus
    Else
        Call LoadPatients("011001")
    End If
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteCommunityIdentify()
'功能：补充社区身份验证
    Dim arrSQL As Variant, i As Long
    Dim colInfo As New Collection
    Dim int社区 As Integer, str社区号 As String
    Dim str出生日期 As String
        
    If gobjCommunity Is Nothing Or mlng病人ID = 0 Or mPatiInfo.挂号ID = 0 Or mPatiInfo.社区 <> 0 Then Exit Sub
    
    If Not gobjCommunity.Identify(glngSys, p门诊医生站, int社区, str社区号, colInfo, mPatiInfo.病人ID, mPatiInfo.挂号ID) Then Exit Sub
    
    arrSQL = Array()
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病人社区信息_Insert(" & mPatiInfo.病人ID & "," & int社区 & ",'" & str社区号 & "',1,Sysdate)"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    str出生日期 = GetColItem(colInfo, "出生日期")
    If IsDate(str出生日期) Then
        str出生日期 = "To_Date('" & Format(str出生日期, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
    Else
        str出生日期 = "Null"
    End If
    arrSQL(UBound(arrSQL)) = "Zl_病人挂号记录_社区验证(" & mPatiInfo.病人ID & "," & mPatiInfo.挂号ID & "," & int社区 & "," & _
        "'" & GetColItem(colInfo, "姓名") & "','" & GetColItem(colInfo, "性别") & "','" & GetColItem(colInfo, "年龄") & "'," & _
        str出生日期 & ",'" & GetColItem(colInfo, "出生地点") & "','" & GetColItem(colInfo, "身份证号") & "'," & _
        "'" & GetColItem(colInfo, "民族") & "','" & GetColItem(colInfo, "国籍") & "','" & GetColItem(colInfo, "婚姻状况") & "'," & _
        "'" & GetColItem(colInfo, "职业") & "','" & GetColItem(colInfo, "家庭地址") & "','" & GetColItem(colInfo, "家庭电话") & "'," & _
        "'" & GetColItem(colInfo, "家庭地址邮编") & "','" & GetColItem(colInfo, "工作单位") & "','" & GetColItem(colInfo, "单位电话") & "'," & _
        "'" & GetColItem(colInfo, "单位邮编") & "','" & GetColItem(colInfo, "联系人姓名") & "','" & GetColItem(colInfo, "联系人关系") & "'," & _
        "'" & GetColItem(colInfo, "联系人电话") & "','" & GetColItem(colInfo, "联系人地址") & "','" & GetColItem(colInfo, "户口地址") & "','" & GetColItem(colInfo, "户口地址邮编") & "')"
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "ExecuteCommunityIdentify"
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    Call LoadPatients("110")
    Call ReshDataQueue
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetColItem(colInfo As Collection, strItem As String) As String
    If colInfo Is Nothing Then Exit Function
    
    err.Clear: On Error Resume Next
    GetColItem = colInfo("_" & strItem)
    err.Clear: On Error GoTo 0
End Function

Private Sub SetRoomState(ByVal blnBusy As Boolean)
'功能：设置诊室忙闲状态
    On Error GoTo DBError
    gcnOracle.Execute "Update 门诊诊室 Set 缺省标志=" & IIf(blnBusy, 1, 0) & " Where 名称='" & mstr接诊诊室 & "' And 缺省标志<>" & IIf(blnBusy, 1, 0)
    On Error GoTo 0
    
    Me.stbThis.Panels(3).Text = "诊室" + IIf(blnBusy, "忙", "闲")
    Me.lblRoom.BackColor = IIf(blnBusy, COLOR_BUSY, COLOR_FREE)
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetTimer()
    mintRefresh = Val(zlDatabase.GetPara("候诊刷新间隔", glngSys, p门诊医生站, 180))
    If mintRefresh <> 0 And mintRefresh < 30 Then mintRefresh = 30
    If mintRefresh = 0 Then
        timRefresh.Enabled = False
    Else
        timRefresh.Interval = 1000 '固定为1秒钟
        timRefresh.Enabled = True
    End If
End Sub
Private Sub timRefresh_Timer()
    Static lngSecond As Long
    Static strPreTime1 As String
    Dim curTime As Date
    
    If mbln消息语音 Then
        If Not mrsMsg Is Nothing Then
            If mrsMsg.RecordCount > 0 Then
                timRefresh.Enabled = False
                Call mclsMsg.PlayMsgSound(mrsMsg)
                Set mrsMsg = Nothing
                timRefresh.Enabled = True
            End If
        End If
    End If
    
    If Not mclsMipModule Is Nothing Then
        If mclsMipModule.IsConnect Then '使用了消息平台用新的刷新策略
            lngSecond = lngSecond + 1
            If lngSecond Mod 180 = 0 Then
                lngSecond = 0
                Call RefeshByMsg
            End If
            Exit Sub
        End If
    End If
    
    curTime = Now
    If mintNotify > 0 And rptNotify.Visible Then
        If strPreTime1 = "" Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime1), curTime) > mintNotify * CLng(60) Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            Call LoadNotify
        End If
    End If
    
    If mintRefresh = 0 Or mblnUnRefresh Or Me.hwnd <> GetForegroundWindow Then Exit Sub
    lngSecond = lngSecond + 1 '秒数
    If lngSecond Mod mintRefresh = 0 Then
        lngSecond = 0
        Call LoadPatients("100111")
        Call ReshDataQueue
    End If
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal strIDCard As String, Optional ByVal blnIsCard As Boolean _
                            , Optional ByVal lngPatiID As Long)
'功能：查找(下一个)病人
'参数：blnNext=是否查找下一个
'      strIDCard=当有值时，表示固定按身份证号查找
'      blnIsCard=是否是刷卡调用接收预约病人
    Static blnReStart As Boolean
    Dim intIdx As PatiType, i As Long
    Dim objControl As CommandBarControl
    Dim objLvw As ListView
    Dim lngReserve As Long  '从头找的时候先找预约的
    Dim blnQueueFind As Boolean
    
    If mintActive = -1 Or mintActive = pt转诊 Then
        PatiIdentify.Text = "": Exit Sub
    End If
    lngReserve = 1
    
    '按其他方式查找后，自动刷身份证的继续查找则取消
    If strIDCard = "" And PatiIdentify.Text <> "" Then mstrIDCard = ""
    
    If Not blnNext And mstrFindType = "挂号单" Then
        PatiIdentify.Text = GetFullNO(PatiIdentify.Text, 12)
    End If
    PatiIdentify.SetFocus
    If mintActive = pt回诊 Then
        Set objLvw = lvwPatiHZ
    ElseIf mintActive = pt预约 Then
        Set objLvw = lvwReserve
    Else
        Set objLvw = lvwPati(mintActive)
    End If
    
    '开始查找行
    If Not blnNext Or blnReStart Or objLvw.SelectedItem Is Nothing Then
        intIdx = pt候诊 - lngReserve: i = 1
    Else
        intIdx = mintActive
        '=3为会诊
        If intIdx = pt回诊 Then intIdx = pt回诊 - 2
        i = objLvw.SelectedItem.Index + 1
    End If
    
     '查找病人
    If lngPatiID = 0 And Not mobjSquareCard Is Nothing And mstrFindType <> "就诊卡" And mstrFindType <> "标识号" And mstrFindType <> "挂号单" And mstrFindType <> "姓名" And mstrFindType <> "二代身份证" Then
        If mstrFindType = "IC卡" Then
            Call mobjSquareCard.zlGetPatiID("IC卡", PatiIdentify.Text, , lngPatiID)
        Else
            Call mobjSquareCard.zlGetPatiID(Val(PatiIdentify.objIDKind.GetCurCard.接口序号), PatiIdentify.Text, , lngPatiID)
        End If
    End If
    
    '查找病人
    If Check排队叫号 = True Then
        blnQueueFind = mobjQueue.FindQueue(IIf(PatiIdentify.objIDKind.GetCurCard.接口序号 > 0, _
                            PatiIdentify.objIDKind.GetCurCard.接口序号, _
                            IIf(PatiIdentify.objIDKind.GetCurCard.名称 = "标识号", "门诊号", PatiIdentify.objIDKind.GetCurCard.名称)), _
                            PatiIdentify.Text)
    End If
    If blnQueueFind = False Then
        For intIdx = intIdx To lvwPati.UBound + 2
            If intIdx = lvwPati.UBound + 1 Then
                Set objLvw = lvwPatiHZ
            ElseIf intIdx = pt候诊 - lngReserve Or intIdx = lvwPati.UBound + 2 Then
                Set objLvw = lvwReserve
            Else
                Set objLvw = lvwPati(intIdx)
            End If
            For i = i To objLvw.ListItems.Count
                With objLvw.ListItems(i)
                    If strIDCard <> "" Then '身份证自动识别强制优先
                        If UCase(.ListSubItems(6).Tag) = UCase(strIDCard) Then Exit For
                    Else
                        If Val(.Tag) = lngPatiID And lngPatiID <> 0 Then Exit For
                        Select Case mstrFindType
                            Case "就诊卡"
                                If .ListSubItems(1).Tag = PatiIdentify.Text Then Exit For
                            Case "标识号"
                                If .SubItems(1) = PatiIdentify.Text Then Exit For '门诊号
                            Case "挂号单"
                                If UCase(.Text) = UCase(PatiIdentify.Text) Then Exit For '单据号
                            Case "姓名"
                                If .SubItems(2) Like "*" & PatiIdentify.Text & "*" Then Exit For
                            Case "二代身份证"
                                If UCase(.ListSubItems(6).Tag) = UCase(PatiIdentify.Text) Then Exit For
                            Case "IC卡"
                                If UCase(.ListSubItems(7).Tag) = UCase(PatiIdentify.Text) Then Exit For
                            Case Else
                                If Val(.Tag) = lngPatiID Then Exit For
                        End Select
                    End If
                End With
            Next
            If i <= objLvw.ListItems.Count Then Exit For
            i = 1
        Next
    
        If intIdx <= lvwPati.UBound + 2 Then
            blnReStart = False
            If intIdx = lvwPati.UBound + 1 Then
                Set objLvw = lvwPatiHZ
            ElseIf intIdx = pt候诊 - lngReserve Or intIdx = lvwPati.UBound + 2 Then
                Set objLvw = lvwReserve
                If intIdx = pt候诊 - lngReserve Then intIdx = pt预约
            Else
                Set objLvw = lvwPati(intIdx)
            End If
            mstrPrePati = objLvw.ListItems(i).Key
            objLvw.ListItems(i).Selected = True
            objLvw.SelectedItem.EnsureVisible
            
            mstrPrePati = ""
            If intIdx = lvwPati.UBound + 1 Then
                Call LvwItemClick(pt回诊, objLvw.SelectedItem)
            ElseIf intIdx <> lvwPati.UBound + 2 Then
                Call lvwPati_ItemClick(CInt(intIdx), objLvw.SelectedItem)
            ElseIf intIdx = pt预约 Then
                Call LvwItemClick(pt预约, objLvw.SelectedItem)
            End If
            
            If Not objLvw.Visible Then
                For i = 1 To dkpMain.PanesCount
                    If dkpMain.Panes(i).Handle = objLvw.hwnd Then
                        dkpMain.Panes(i).Select
                    End If
                Next
            End If
            If objLvw.Visible Then objLvw.SetFocus
            
            '找到后自动进行接诊,预约病人自动接收
            If (mbln自动接诊 And intIdx = pt候诊) Or intIdx = pt预约 Then
                cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                If intIdx = pt预约 Then
                    If mstrFindType = "标识号" Or mstrFindType = "挂号单" Or mstrFindType = "姓名" Or mstrFindType = "二代身份证" Then Exit Sub
                    Call ExecuteReceive(blnIsCard)
                Else
                    Set objControl = cbsMain.FindControl(, conMenu_Manage_Receive, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then Call cbsMain_Update(objControl) '首次执行，没有显示菜单前，事件没有执行
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
        Else
            blnReStart = True
            MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的病人。", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub ucPatiVitalSigns_Change(ByVal int序号 As Integer)
    Call SetPermitEscape(False)
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsAller.ComboList = "..."
    vsAller.FocusRect = flexFocusSolid
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int性别 As Integer
    
    With vsAller
        Call RefreshPass
        If mblnUseTYT Then
            strSQL = gobjPass.inputAllergy()
            If strSQL <> "" Then
                Call SetAllerInput(Col, , strSQL)
                Call AllerEnterNextCell
            End If
        Else
            If cboEdit(cbo性别).Text Like "*男*" Then
                int性别 = 1
            ElseIf cboEdit(cbo性别).Text Like "*女*" Then
                int性别 = 2
            End If
            
            strSQL = _
                " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                " Select ID,Nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
                " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试" & _
                " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                " Union All" & _
                " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码,A.名称," & _
                " A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                " From 诊疗项目目录 A,药品特性 B" & _
                " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
                IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[1])", "") & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "过敏药物", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int性别)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有药品数据可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call SetAllerInput(Col, rsTmp)
                Call AllerEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub vsAller_Click()
'    If vsAller.TextMatrix(0, vsAller.Col) = "" Then
'        Call vsAller_DblClick
'    End If
End Sub

Private Sub vsAller_DblClick()
    If vsAller.Editable = flexEDKbdMouse Then
        With vsAller
            .ComboList = ""
            .EditText = .TextMatrix(.Row, .Col)
            .EditCell
        End With
    End If
End Sub

Private Sub vsAller_GotFocus()
    vsAller.BackColorSel = IIf(mblnPatiChange, EColor, HColor)
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, blnDo As Boolean
    
    With vsAller
        If KeyCode = vbKeyDelete Then
            If .TextMatrix(0, .Col) <> "" Then
                If MsgBox("确实要清除该项过敏药物吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    blnDo = True
                End If
            Else
                blnDo = .Col <> .Cols - 1
            End If
            
            If blnDo Then
                If CStr(.Cell(flexcpData, 0, .Col)) <> "" Then  '删除以前保存过的数据
                    .ColWidth(.Col) = 0
                    '.ColHidden(.Col) = True    '不使用hidden，因为隐藏了列后，最后一列不等于.col-1
                    .TextMatrix(0, .Col) = ""
                    
                    If .Col = .Cols - 1 Then
                        Call AllerAddCol
                    Else
                        .Col = .Cols - 1
                        .ShowCell 0, .Col
                    End If
                    
                    Call SetPermitEscape(False)
                Else   '删除一列时，左移后续列
                    For i = .Col + 1 To .Cols - 1
                        .TextMatrix(0, i - 1) = .TextMatrix(0, i)
                        .ColData(i - 1) = .ColData(i)
                        .ColWidth(i - 1) = Me.TextWidth(.TextMatrix(0, i - 1)) + 260
                        .Cell(flexcpData, 1, i - 1) = .Cell(flexcpData, 1, i)
                    Next
                    .Cols = .Cols - 1
                    .ColWidth(.Cols - 1) = 1200
                End If
            End If
            .SetFocus
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsAller_KeyPress(KeyCode)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
            Call .ShowCell(0, .Col)
        End If
    End With
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    With vsAller
        If KeyAscii = vbKeySpace Then  'Space
            If mblnUseTYT Then KeyAscii = 0: Exit Sub
        End If
        If KeyAscii = 13 Then
            KeyAscii = 0
            vsAller.Tag = "KeyPress"
            Call AllerEnterNextCell
            vsAller.Tag = ""
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsAller_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsAller_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mblnUseTYT Then KeyAscii = 0
    End If
End Sub

Private Sub vsAller_LostFocus()
    vsAller.BackColorSel = vsAller.BackColor
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAller.EditSelStart = 0
    vsAller.EditSelLength = zlCommFun.ActualLen(vsAller.EditText)
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim int性别  As Integer
    
    With vsAller
        If .EditText = "" Then
            .Cell(flexcpData, Row, Col) = ""
            If vsAller.Tag = "KeyPress" Then Call AllerEnterNextCell
        ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
            If vsAller.Tag = "KeyPress" Then Call AllerEnterNextCell
        Else
            strInput = UCase(.EditText)
            If cboEdit(cbo性别).Text Like "*男*" Then
                int性别 = 1
            ElseIf cboEdit(cbo性别).Text Like "*女*" Then
                int性别 = 2
            End If
            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位," & _
                " B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
                " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.ID=C.诊疗项目ID" & _
                " And (A.编码 Like [1] Or A.名称 Like [2] Or C.名称 Like [2] Or C.简码 Like [2])" & _
                IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[3])", "") & _
                Decode(gint简码, 0, " And C.码类=[4]", 1, " And C.码类=[4]", "") & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " Order by A.编码"
            
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "过敏药物", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                strInput & "%", gstrLike & strInput & "%", int性别, gint简码 + 1)
            If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                Cancel = True
            Else
                Call SetAllerInput(Col, rsTmp): .EditText = .Text
                If vsAller.Tag = "KeyPress" Or vsAller.Col = vsAller.Cols - 1 Then Call AllerEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub SetAllerInput(ByVal lngCol As Long, Optional rsInput As ADODB.Recordset, Optional ByVal strTYTInput As String)
'功能：处理过敏药物的输入
    Dim strAllerOld As String, strAllerNew As String
    Dim arrTmp As Variant
    With vsAller
        strAllerOld = .TextMatrix(0, lngCol) & ";" & .TextMatrix(1, lngCol)
        If mblnUseTYT Then
            arrTmp = Split(strTYTInput, ";")
            If UBound(arrTmp) < 1 Then Exit Sub
            If strAllerOld <> strTYTInput Or Val(.ColData(lngCol) & "") <> 0 Then
                .TextMatrix(0, lngCol) = arrTmp(1)
                .TextMatrix(1, lngCol) = arrTmp(0)
                .ColData(lngCol) = 0
                .Cell(flexcpData, 1, lngCol) = ""
            End If
        Else
            If Not rsInput Is Nothing Then
                If Val(.ColData(lngCol) & "") = Val(rsInput!ID) Then Exit Sub
                .Cell(flexcpData, 1, lngCol) = ""
                .ColData(lngCol) = Val(rsInput!ID)
                .TextMatrix(0, lngCol) = "" & rsInput!名称
            Else    '当成自由录入
                
                If .TextMatrix(0, lngCol) = .EditText Then Exit Sub
                .Cell(flexcpData, 1, lngCol) = ""
                .ColData(lngCol) = 0
                .TextMatrix(0, lngCol) = .EditText
            End If
            If .TextMatrix(1, lngCol) <> "" Then .TextMatrix(1, lngCol) = ""
            strAllerNew = strAllerOld = .TextMatrix(0, lngCol) & ";" & .TextMatrix(1, lngCol)
            If strAllerNew <> strAllerOld Or Val(.ColData(lngCol) & "") <> 0 Then
                .TextMatrix(1, lngCol) = ""
            End If
        End If
        .AutoSize 0, lngCol
        
        Call SetPermitEscape(False)
    End With
End Sub

Private Sub AllerEnterNextCell()
    With vsAller
        If Trim(.TextMatrix(0, .Col)) <> "" Then
            If .Col = .Cols - 1 Then
                Call AllerAddCol
            Else
                .ShowCell 0, .Col + 1
            End If
            .Col = .Col + 1
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub AllerAddCol()
'功能：增加一空列用于输入
    With vsAller
        .Cols = .Cols + 1
        .ShowCell 0, .Cols - 1
        
        .ColWidth(.Cols - 1) = 1200
        .ColAlignment(.Cols - 1) = flexAlignLeftCenter
    End With
End Sub
Private Function Check排队叫号() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查和创建排队叫号功能
    '返回：排队叫号功能所有的都合法,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-06 10:19:43
    '说明：需检查: 权限合法检查;启用了排队叫号的;创建排队叫号成功!
    '------------------------------------------------------------------------------------------------------------------------
    '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    If mty_Queue.byt排队叫号模式 = 0 Then GoTo GOEND:
    If Not (InStr(mty_Queue.strQueuePrivs, ";基本;") > 0) Then GoTo GOEND:
    If mty_Queue.bln医生主动呼叫 = False And mty_Queue.byt排队叫号模式 = 1 Then GoTo GOEND:
    
    err = 0: On Error GoTo GOEND:
    If mobjQueue Is Nothing Then
        Set mobjQueue = CreateObject("zlQueueManage.clsQueueManage")
        err = 0: On Error GoTo ErrHand:
        mobjQueue.zlInitVar gcnOracle, glngSys, 0, IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数), mty_Queue.strQueuePrivs, "", False
        mobjQueue.zlSetToolIcon 24, True
        mobjQueue.IsShowFindTools = False
    End If
    Check排队叫号 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
GOEND:
    If Not mobjQueue Is Nothing Then mobjQueue.CloseWindows
    Set mobjQueue = Nothing

End Function
Private Sub ReshDataQueue()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：刷新排队叫号数据
    '编制：刘兴洪
    '日期：2010-06-07 15:27:57
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim varQueue() As String, strTemp As String, rsTemp As ADODB.Recordset, strSQL As String
    Dim str诊室 As String, str医生 As String, str科室 As String
    Dim intType As Integer
    
    If mobjQueue Is Nothing Then Exit Sub
    If Check排队叫号 = False Then Exit Sub
    '获取相关的队列名称
    '接诊范围：1=挂本人号的病人,2=本诊室病人,3=本科室病人
    mint接诊范围 = Val(zlDatabase.GetPara("接诊范围", glngSys, p门诊医生站, "2"))
    Dim strQueue() As String
    
    ReDim Preserve strQueue(1 To 1) As String
    str科室 = IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID)
    strQueue(1) = str科室
    str医生 = IIf(mstr接诊医生 = "", UserInfo.姓名, mstr接诊医生)
    str诊室 = mstr接诊诊室
    intType = 1
    Select Case mint接诊范围
    Case 1   '1=挂本人号的病人
        If Not mty_Queue.bln医生主动呼叫 Then
           str医生 = UserInfo.姓名  '64696,刘尔旋,2014-01-08,用登录人员的姓名过滤排队叫号队列
        End If
        If mlng接诊科室ID = 0 Then strQueue(1) = ""
        intType = 3
    Case 2  '2=本诊室病人
        If Not mty_Queue.bln医生主动呼叫 Then
           str诊室 = mstr接诊诊室
        End If
        If mlng接诊科室ID = 0 Then strQueue(1) = ""
        intType = 2
    Case 3  '3=本科室病人
    End Select
    
    '需要排队没有建档的病人
    strSQL = "" & _
    "   Select distinct  /*+ Rule*/  c.业务ID From 病人挂号记录 A ,排队叫号队列  C" & _
    "   Where A.id=C.业务ID and C.队列名称=[1]  and nvl(C.业务类型,0)=0 and nvl(A.病人ID,0) =0 And a.记录性质=1 And a.记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str科室)
    With rsTemp
        strTemp = ""
        Do While Not .EOF
            strTemp = strTemp & "," & Val(Nvl(rsTemp!业务id))
            .MoveNext
        Loop
        If strTemp <> "" Then strTemp = "0|" & Mid(strTemp, 2)
    End With
    Call mobjQueue.zlRefresh(strQueue, mty_Queue.strCurrQueueName, mty_Queue.lngcurr挂号ID, str诊室, str医生, strTemp, intType)
End Sub
 
Private Sub zlQueueStartus(intType As Integer, strNO As String, lng病人ID As Long)
  '------------------------------------------------------------------------------------------------------------------------
    '功能：功能操作后,
    '入参：2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-病人取消就诊
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-06-03 14:15:46
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strQueueName As String, lngID As Long
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim i As Byte
    If Check排队叫号 = False Then Exit Sub
    
    strSQL = "SELECT ID,执行部门ID,诊室,执行人 From 病人挂号记录 where NO=[1] And 记录性质=1 And 记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    
    strQueueName = Nvl(rsTemp!执行部门ID)
    If Nvl(rsTemp!执行人) <> "" Then
        strQueueName = strQueueName & ":" & Nvl(rsTemp!执行人)
    ElseIf Nvl(rsTemp!诊室) <> "" Then
        strQueueName = strQueueName & ":" & Nvl(rsTemp!诊室)
    End If
    
    lngID = Val(Nvl(rsTemp!ID))
    Select Case intType
    Case 3   ' 病人不就诊;
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 3
    Case 4, 6   '病人待诊,'病人取消就诊
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 0
    Case 5  '病人完成就诊
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 4
    End Select
End Sub

Private Function Set病人挂号状态(ByVal lngState As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置病人挂号状态
    '入参：lngState : -1- 病人不就诊
    '                         0-病人待诊
    '出参：
    '返回：是否设置成功，病人不就诊时可以删除划价单据，当再次设置待诊时会设置不成功 返回False ,其他情况返回True
    '编制：刘兴洪
    '日期：2010-06-03 15:24:48
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str划价NO As String
    
    If mstr挂号单 = "" Then Exit Function
    
    On Error GoTo errH
    
    If lngState = -1 Then
        '检查病人是否存在有效的医嘱
        strSQL = "Select 1 From 病人医嘱记录 Where 病人id = [1] And 挂号单 = [2]  And 医嘱状态 <> -1 And 医嘱状态 <> 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
        If Not rsTmp.EOF Then
            MsgBox "该病人存在有效医嘱,不能设置为不就诊!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If

    '获取挂号划价单信息
    strSQL = "Select 摘要 From 门诊费用记录 Where NO = [1] And 记录性质 = 4 And 记录状态 = 1 And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
    If Not rsTmp.EOF Then
        If rsTmp!摘要 & "" <> "" And InStr(rsTmp!摘要 & "", "划价:") <> 0 Then
            '获取挂号划价单信息,判断挂号划价单是否存在，不存在，则不允许将病人状态设置为待诊
            str划价NO = Mid(rsTmp!摘要 & "", Len("划价:") + 1)
            strSQL = "Select 1 From 门诊费用记录 Where NO = [1] And Mod(记录性质,10) = 1 And 记录状态 = 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str划价NO)
            If rsTmp.EOF Then
                If lngState = 0 Then '设置为待诊
                    MsgBox "该挂号单的划价费用不存在，请退号后重新挂号!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
            Else
                If lngState = -1 Then '设置为不就诊
                    If MsgBox("该病人存在挂号单的划价费用，设置为不就诊时将删除该挂号单的划价费用，" & vbCrLf & "并且不能再恢复为待诊,是否继续?。", vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    
    gcnOracle.BeginTrans
        strSQL = "Zl_病人挂号记录_状态 ('" & mstr挂号单 & "'," & lngState & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Call zlQueueStartus(IIf(lngState = -1, 3, 4), mstr挂号单, mlng病人ID)
        'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
    gcnOracle.CommitTrans
    MsgBox "操作成功!", vbInformation, gstrSysName
    
    Set病人挂号状态 = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub ExecuteStopAndReuse(ByVal bln启用 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对就诊病人进行暂停就诊或启用诊断
    '入参:bln启用-true:启用已经停用的就诊病人
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-08 20:26:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, bln暂停 As Boolean
    Dim strNO As String, rsTemp As ADODB.Recordset
    Dim objLvw As ListView
    If Not bln启用 Then
        Set objLvw = lvwPati(pt就诊)
    Else
        Set objLvw = lvwPatiHZ
    End If
    With objLvw
        If .SelectedItem Is Nothing Then Exit Sub
        bln暂停 = .SelectedItem.SmallIcon = "暂停"
        If bln启用 Then
            If bln暂停 = False Then
                MsgBox "注意:" & vbCrLf & "    该病人还未暂停就诊,不能进行恢复暂停就诊!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        Else
            If bln暂停 Then
                MsgBox "注意:" & vbCrLf & "    该病人还启用暂停就诊,不能进行暂停就诊!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        End If
        strNO = .SelectedItem.Text
        strSQL = "Select ID From 病人挂号记录 where NO=[1] And 记录性质=1 And 记录状态=1"
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTemp.EOF Then
            Exit Sub
        End If
    End With
    If Not bln启用 Then
        'Zl_病人挂号记录_回诊
        strSQL = "Zl_病人挂号记录_回诊("
        '  Id_In         病人挂号记录.ID%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!ID)) & ","
        '  新执行科室_In 病人挂号记录.执行部门id%Type,
        strSQL = strSQL & "NULL,"
        '  新诊室_In     病人挂号记录.诊室%Type,
        strSQL = strSQL & "NULL,"
        '  新医生_In     病人挂号记录.执行人%Type,
        strSQL = strSQL & "NULL,"
        '  需回诊_In Integer:=0
        strSQL = strSQL & "1)"
        '--需回诊_In :0-回诊操作;1-标记为需要回诊
    Else
        'Zl_病人挂号记录_取消回诊
        strSQL = "Zl_病人挂号记录_取消回诊("
        '  Id_In         病人挂号记录.ID%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!ID)) & ","
        '  需回诊_In Integer:=0
        strSQL = strSQL & "1)"
    End If
    On Error GoTo errHandle
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    '刷新:不定位到已诊列表
    Call LoadPatients
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub lvwPatiHZ_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call zlControl.LvwSortColumn(lvwPatiHZ, ColumnHeader.Index)
End Sub

Private Sub lvwPatiHZ_GotFocus()
    'MouseDown先于GotFocus执行
    If Not mblnMouseDown And Not lvwPatiHZ.SelectedItem Is Nothing Then
        Call lvwPatiHZ_ItemClick(lvwPatiHZ.SelectedItem)
    End If
End Sub

Private Sub lvwPatiHZ_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LvwItemClick(pt回诊, Item)
End Sub

Private Sub lvwPatiHZ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMouseDown = True
End Sub
Private Sub lvwPatiHZ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim objItem As ListItem
    mblnMouseDown = False
    If Button = 2 And InStr(mstrPrivs, "病人接诊") > 0 Then
        Set objItem = lvwPatiHZ.HitTest(x, y)
        If Not objItem Is Nothing Then
            Set objPopup = cbsMain.ActiveMenuBar.Controls(2)
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub


Private Sub Set诊疗项目费用设置()
    Dim lng科室ID As Long
    
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "诊疗基础部件(ZLCISBase)没有正确安装，该功能无法执行。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    If mlng科室ID = 0 Then
        lng科室ID = mPatiInfo.科室ID
    Else
        lng科室ID = mlng科室ID
    End If
    If lng科室ID = 0 Then
        lng科室ID = UserInfo.部门ID
    End If
        
    Call gobjCISBase.CallSetClinicCharge(lng科室ID, 1, Me, gcnOracle, glngSys, gstrDBUser, E门诊调用, InStr(GetInsidePrivs(p门诊医生站), ";诊疗项目费用设置;") = 0)
End Sub
Private Sub SetFontSize(ByVal blnSetMainFont As Boolean)
'功能: 进行界面字体的统一设置
'参数: blnSetMainFont 是否设置主界面字体(用以区分子界面切换)
    If blnSetMainFont Then
        Call SetPublicFontSize(Me, mbytSize)
        Call zlControl.VSFSetFontSize(vsAller, IIf(mbytSize = 0, 9, 12))
        vsAller.Height = vsAller.RowHeightMin + IIf(mbytSize = 0, 15, 30)

        Call SetPatiInfoPosition
        Call SetPicBasisFontSizeAndPosition
        Call SetPicOutDocFontSizeAndPosition

        Call picPatiInput_Resize
        Call picYZ_Resize
    End If

    Select Case tbcSub.Selected.Tag
        Case "医嘱"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "病历"
            Call mclsEPRs.SetFontSize(mbytSize)
                Case "新病历"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
    End Select
End Sub

Private Sub SetPatiInfoPosition()
'功能：设置病人信息的控件位置
    Dim lngDistance1 As Long
    Dim lngDistance2 As Long
    
    lngDistance1 = 30
    lngDistance2 = 120
    Call SetCtrlPosOnLine(False, 0, lblTitle费别, lngDistance1, lblShow(lbl费别), lngDistance2, lblTitle付款, lngDistance1, lblShow(lbl付款), lngDistance2, lblTitle号类, _
        lngDistance1, lblShow(lbl号类), lngDistance2, lblTitle医保号, lngDistance1, lblShow(lbl医保号), lngDistance2, lblTitle社区号, lngDistance1, lblShow(lbl社区号))
    
    lblDiag(0).Top = lblTitle费别.Top + lblTitle费别.Height + 90
    Call SetCtrlPosOnLine(False, 0, lblDiag(0), lngDistance1, lblDiag(1))
    picFind.Width = IIf(mbytSize = 1, 670, 475)
    lblFind.Width = picFind.Width
    Call SetCtrlPosOnLine(False, 0, picFind, 10, PatiIdentify)
        picFind.Height = PatiIdentify.Height
End Sub

Private Sub SetPicBasisFontSizeAndPosition()
'功能：设置病人信息的字体
    Dim lngDistance1 As Long
    Dim lngDistance2 As Long
    Dim objFont As Font
    Dim i As Long
    Dim lngFontSize As Long
    
    lngFontSize = IIf(mbytSize = 0, 9, 12)
    
    Set objFont = ucPatiVitalSigns.Font
    objFont.Size = lngFontSize
    Set ucPatiVitalSigns.Font = objFont
    
    lngDistance1 = 30
    lngDistance2 = 120
    On Error Resume Next
    lbl急.FontName = "黑体"
    lbl急.FontSize = IIf(mbytSize = 0, 14, 18)
    
    lblRec.FontName = "黑体"
    lblRec.FontSize = IIf(mbytSize = 0, 14, 18)
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
    
    '就诊科室控件的大小设置
    fraLine(3).Width = cboRegist.Width
    fraLine(3).Height = cboRegist.Height
    fraLine(3).Left = lblRegistInput.Left + lblRegistInput.Width + 20
    fraLine(3).Top = IIf(mbytSize = 0, 130, 160)
    lblRegistInput.Top = fraLine(3).Top + 30
    fraRegistInput.Width = fraLine(3).Left + fraLine(3).Width - 60
    fraRegistInput.Height = lblRegistInput.Top + lblRegistInput.Height + 110
    
    '性别年龄下拉列表以及相关控件的大小设置
    For i = 0 To cboEdit.UBound
        If i = 3 Then i = i + 2
        cboEdit(i).Width = Me.TextWidth("设置字")
        fraLine(i).Width = cboEdit(i).Width
        fraLine(i).Height = cboEdit(i).Height + cboEdit(i).Top - 30
        fraLine(i).Top = lblEdit(lbl姓名).Top + lblEdit(lbl姓名).Height - fraLine(i).Height - 30
    Next
    cmdAller.Height = Me.TextHeight("字") * IIf(mbytSize = 0, 2, 1.5)
    cmdAller.Width = Me.TextWidth(cmdAller.Caption & "字")
    Call SetCtrlPosOnLine(True, -1, fraRegistInput, 100, lblEdit(lbl出生日期), 60, lblEdit(cbo职业), 60, lblEdit(lbl单位), 60, lblEdit(txt家庭地址), 180, lblEdit(lbl身份证), 180, ucPatiVitalSigns)
    Call SetCtrlPosOnLine(False, 0, fraRegistInput, lngDistance2, lblEdit(txt姓名), lngDistance1, txtEdit(txt姓名), lngDistance2, fraLine(cbo性别))
    Call SetCtrlPosOnLine(False, 0, lblEdit(lbl出生日期), lngDistance1, txt出生日期, lngDistance1 * 0.5, txt出生时间, lngDistance2, lblEdit(txt年龄), lngDistance1, txtEdit(txt年龄), lngDistance1, fraLine(cbo年龄))

    cboEdit(cbo职业).Width = fraRegistInput.Width + fraRegistInput.Left - fraLine(cbo职业).Left
    fraLine(cbo职业).Width = cboEdit(cbo职业).Width
    fraLine(cbo发病时间).Width = cboEdit(cbo发病时间).Width
    Call SetPatiPictureSize '控制照片大小
    Call SetCtrlPosOnLine(False, 0, lblEdit(cbo职业), lngDistance1, fraLine(cbo职业), lngDistance2, lblEdit(lbl发病时间), lngDistance1, txtEdit(txt发病), lngDistance1 * 0.5, fraLine(cbo发病时间), lngDistance1, txt发病日期, lngDistance1, txt发病时间, lngDistance2, lblEdit(txt发病地址), lngDistance1, txtEdit(txt发病地址))
    Call SetCtrlPosOnLine(False, 0, lblEdit(lbl单位), lngDistance1, txtEdit(txt单位名称), -1 * cmdEdit(cmd单位名称).Width, cmdEdit(cmd单位名称), lngDistance2, lblEdit(lbl单位电话), lngDistance1, txtEdit(txt单位电话), lngDistance2, optState(opt初诊), lngDistance1, optState(opt复诊), lngDistance2, lblEdit(lbl去向), lngDistance1, fraLine(cbo去向), lngDistance2, lbl多科就诊)
    
    Call SetCtrlPosOnLine(False, 0, lblEdit(txt家庭地址), lngDistance1, txtEdit(txt家庭地址), -1 * cmdEdit(cmd单位名称).Width, cmdEdit(cmd家庭地址), lngDistance2, lblEdit(txt监护人), lngDistance1, txtEdit(txt监护人), lngDistance2, lblEdit(lbl家庭电话), lngDistance1, txtEdit(txt家庭电话), lngDistance2, lblEdit(lbl摘要), lngDistance1, txtEdit(txt就诊摘要))  'lblEdit(txt收缩压), lngDistance1, txtEdit(txt收缩压), lngDistance1, lblEdit(txt舒张压), lngDistance1, txtEdit(txt舒张压), lngDistance1, fraLine(cbo血压单位))
    Call SetCtrlPosOnLine(False, 0, lblEdit(lbl身份证), lngDistance1, txtEdit(txt身份证号), lngDistance2, lblEdit(lbl手机号), lngDistance1, txtEdit(txt手机号), lngDistance2, lblEdit(lbl过敏), lngDistance1, cmdAller, -30, vsAller)

End Sub

Private Sub SetPatiPictureSize()
    picPatient.Left = fraLine(cbo性别).Left + fraLine(cbo性别).Width
    If picPatient.Left < fraLine(cbo年龄).Left + fraLine(cbo年龄).Width Then
        picPatient.Left = fraLine(cbo年龄).Left + fraLine(cbo年龄).Width
    End If
    picPatient.Left = picPatient.Left + 100
    
    '设置照片大小
    picPatient.Height = lblEdit(lbl出生日期).Top + lblEdit(lbl出生日期).Height - picPatient.Top
    picPatient.Width = picPatient.Height * 1.25
    imgPatient.Height = picPatient.Height - 75
    imgPatient.Width = picPatient.Width - 75
End Sub

Private Sub SetPicOutDocFontSizeAndPosition()
'功能：设置病人病史信息块的字体及控件大小
    Dim i As Long
    
    For i = 0 To lblDoc.UBound
        lblDoc(i).Width = Me.TextWidth("字")
        lblDoc(i).Height = Me.TextHeight("字") * 3
    Next
    lbl提示.Left = rtfEdit(txt查体).Left
    lbl病历名称.Left = rtfEdit(txt查体).Left
    cmdSign.Left = rtfEdit(txt查体).Left + rtfEdit(txt查体).Width - cmdSign.Width
    
    '设置Fontsize时会触发change事件
    mblnSizeTmp = True
    Call SetRTFEditFontSize
    mblnSizeTmp = False
End Sub

Private Sub SetRTFEditFontSize()
'功能：设置病人病史信息的输入框字体
    Dim i As Long
    
    For i = 0 To rtfEdit.UBound
        Call SetPublicRTFFont(rtfEdit(i), IIf(mbytSize = 0, 9, 12))
    Next
End Sub

Private Function Check接诊控制(str操作 As String, strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人接诊控制
    '入参:str操作 -当前操作 strNo - 挂号单据号
    '出参:
    '返回:
    '编制:王吉
    '日期:2013-1-17 20:26:59
    '问题号:57566
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHanl:
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim rs预约时间 As Recordset
    Dim strMsg As String
    
    If mlng接诊控制 = 0 Then Check接诊控制 = True: Exit Function
    
    strSQL = "" & _
    "   Select  Nvl(A.预约时间,nvl(发生时间,sysdate)) - " & mlng提前接收时间 & "/24/60 as 挂号时间  " & _
    "   From 病人挂号记录 A " & _
    "   Where No=[1] And Nvl(A.预约时间,nvl(发生时间,sysdate))- " & mlng提前接收时间 & "*1/24/60>sysdate"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Check接诊控制 = True: Exit Function
    strMsg = "该病人需要在" & Format(rsTemp!挂号时间, "yyyy-mm-dd HH:MM:SS") & "后才允许进行" & str操作
    If mlng接诊控制 = 2 Then
        Check接诊控制 = (MsgBox(strMsg & ",您确定要进行" & str操作 & "吗？", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes)
    Else
        MsgBox strMsg & ",不允许" & str操作, vbInformation, gstrSysName
    End If
    Exit Function
ErrHanl:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
'功能：处理门诊医生站接收到的消息
    Dim objXML As zl9ComLib.clsXML
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim byt刷新 As Byte  '刷新方式：1-候诊列表，2－转诊列表
    Dim bln刷新 As Boolean
    
    On Error GoTo errH
    
    Set objXML = New zl9ComLib.clsXML
    Call objXML.OpenXMLDocument(strMsgContent)
    Select Case strMsgItemIdentity '获取挂号记录id
        Case "ZLHIS_REGIST_001", "ZLHIS_REGIST_002" '门诊患者挂号，门诊分诊通知。采取一分钟刷新一次方式，如果是第一条消息则立即刷新。
            byt刷新 = 1
            Call objXML.GetSingleNodeValue("register_id", strTmp)
        Case "ZLHIS_CIS_007" '门诊患者转诊。即时刷新，消息到来的时候就刷新，只刷新转诊列表
            byt刷新 = 2
            Call objXML.GetSingleNodeValue("clinic_id", strTmp)
    End Select
    
    If strTmp = "" Then Exit Sub
    
    strSQL = "Select 执行人,诊室,执行部门id,转诊医生,转诊诊室,转诊科室id From 病人挂号记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strTmp))
    
    If byt刷新 = 1 Then
        If mint接诊范围 = 1 And rsTmp!执行人 & "" = UserInfo.姓名 And (Not mbln要求分诊 Or mbln要求分诊 And rsTmp!诊室 & "" <> "") Then
            bln刷新 = True
        Else
            If (mint接诊范围 = 2 And rsTmp!诊室 & "" = mstr接诊诊室 Or mint接诊范围 = 3 And (Not mbln要求分诊 Or mbln要求分诊 And rsTmp!诊室 & "" <> "")) And _
                Val(rsTmp!执行部门ID & "") = IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID) And _
                (rsTmp!执行人 & "" = "" Or rsTmp!执行人 & "" = UserInfo.姓名) Then
                
                bln刷新 = True
            End If
        End If
        
        If bln刷新 Then
            mblnMsgOk = True
            If Not mblnFirstMsg Then     '是第一条消息
                mblnFirstMsg = True
                Call RefeshByMsg
            End If
        End If
    ElseIf byt刷新 = 2 Then
        If mint接诊范围 = 1 And rsTmp!转诊医生 & "" = UserInfo.姓名 Then
            bln刷新 = True
        Else
            If (mint接诊范围 = 2 And rsTmp!转诊诊室 & "" = mstr接诊诊室 Or mint接诊范围 = 3) And _
                Val(rsTmp!转诊科室ID & "") = IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID) And _
                UserInfo.姓名 <> IIf("" = rsTmp!执行人 & "", "无", rsTmp!执行人) And _
                (rsTmp!转诊医生 & "" = "" Or rsTmp!转诊医生 & "" = UserInfo.姓名) Then
                
                bln刷新 = True
            End If
        End If
        
        If bln刷新 Then
            Call LoadPatients("000100")
            Exit Sub
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefeshByMsg()
'功能：启用消息平台后使用的刷新方式
    Dim strTmp As String
    
    If Not mblnMsgOk Then Exit Sub
    '如果预约列表可见则一并刷新
    strTmp = "1000" & IIf(lvwReserve.Visible, 1, 0)
    Call LoadPatients(strTmp)
    Call ReshDataQueue
    mblnMsgOk = False
End Sub

Private Sub Hide就诊卡号列()
'功能：将界面的各个病人列表中的   就诊卡号  列  设置为隐藏
    Dim lngIndex As Long
    Dim strTmp As String
        
    strTmp = "就诊卡号"
    
    lngIndex = GetLvwColIndex(lvwPatiHZ, strTmp)
    lvwPatiHZ.ColumnHeaders(lngIndex).Width = 0
    
    lngIndex = GetLvwColIndex(lvwPati(0), strTmp)
    lvwPati(0).ColumnHeaders(lngIndex).Width = 0
    
    lngIndex = GetLvwColIndex(lvwPati(1), strTmp)
    lvwPati(1).ColumnHeaders(lngIndex).Width = 0
    
    lngIndex = GetLvwColIndex(lvwPati(2), strTmp)
    lvwPati(2).ColumnHeaders(lngIndex).Width = 0
    
    lngIndex = GetLvwColIndex(lvwReserve, strTmp)
    lvwReserve.ColumnHeaders(lngIndex).Width = 0
    
    lngIndex = GetLvwColIndex(lvwIncept, strTmp)
    lvwIncept.ColumnHeaders(lngIndex).Width = 0
    
End Sub

Private Function GetLvwColIndex(ByRef objLvw As ListView, ByVal strColName As String) As Long
'功能：反回 ListView 列表中的指定列的列索引值
    Dim i As Integer
    For i = 1 To objLvw.ColumnHeaders.Count
        If objLvw.ColumnHeaders(i).Text = strColName Then
            GetLvwColIndex = i
            Exit Function
        End If
    Next
End Function

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long
    
    With rptNotify
        Set objCol = .Columns.Add(c_图标, "", 18, True): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_病人ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_No, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(c_门诊号, "门诊号", 62, True)
        Set objCol = .Columns.Add(C_就诊时间, "就诊时间", 60, True)
        Set objCol = .Columns.Add(C_状态, "状态", 150, True)
         
        Set objCol = .Columns.Add(C_消息, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_序号, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_日期, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_业务, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_挂号ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_ID, "", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            If objCol.Index <> C_序号 Or objCol.Index <> C_日期 Then objCol.Sortable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有提醒内容..."
        End With
        .PreviewMode = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '排序 降序
        .SortOrder.Add .Columns(C_序号)
        .SortOrder(0).SortAscending = False
        .SortOrder.Add .Columns(C_日期)
        .SortOrder(1).SortAscending = False
    End With
    
End Sub

Private Function LoadNotify() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strSQL As String
    Dim strTmp As String
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    Dim strTag As String
    
    mstrPreNotify = ""
    rptNotify.Records.DeleteAll
    If Mid(mstrNotifyAdvice, m危机值, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_003,ZLHIS_PACS_005"
    If Mid(mstrNotifyAdvice, m传染病, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_032,ZLHIS_CIS_033"
    If Mid(mstrNotifyAdvice, m处方审查, 1) = "1" Then strTmp = strTmp & ",ZLHIS_RECIPEAUDIT_001"
    If Mid(mstrNotifyAdvice, m备血完成, 1) = "1" And gbln血库系统 Then strTmp = strTmp & ",ZLHIS_BLOOD_001"   '启用新血库流程才有此消息和参数
    If Mid(mstrNotifyAdvice, m用血审核, 1) = "1" And gbln血库系统 Then strTmp = strTmp & ",ZLHIS_BLOOD_004"   '启用新血库流程才有此消息和参数
    If Mid(mstrNotifyAdvice, m输血反应, 1) = "1" And gbln血库系统 Then strTmp = strTmp & ",ZLHIS_BLOOD_006"  '启用血库才有此消息和参数
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then LoadNotify = True: Exit Function
       
    strSQL = "Select b.id,a.病人id,a.NO,a.id as 挂号ID,a.门诊号,a.姓名,a.执行时间 as 就诊时间,b.消息内容,b.类型编码, b.业务标识, b.优先程度, b.登记时间,a.险类,b.病人来源" & _
        " From 业务消息清单 B, 病人挂号记录 A" & _
        " Where b.就诊id=a.Id And a.执行人||''=[1]  And b.登记时间>=Trunc(Sysdate-" & (mintNotifyDay - 1) & ")" & _
        " And Nvl(b.是否已阅,0)=0 And instr(','||[2]||',',','||b.类型编码||',')>0 AND substr(b.提醒场合,1,1)='1' " & _
        " Order By b.优先程度 Desc, b.登记时间 Desc"
    
    Screen.MousePointer = 11

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstr接诊医生, strTmp)
    
    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!类型编码
        Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
            strTag = strTag & "<TB>" & rsTmp!类型编码 & "," & rsTmp!ID
            blnDo = True
        Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!业务标识 & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!业务标识
                blnDo = True
            End If
        Case "ZLHIS_BLOOD_006"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!类型编码 & ":" & rsTmp!病人ID & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!类型编码 & ":" & rsTmp!病人ID
                blnDo = True
            End If
        Case Else
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!病人ID & "," & rsTmp!挂号ID & "," & rsTmp!类型编码 & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!病人ID & "," & rsTmp!挂号ID & "," & rsTmp!类型编码
                blnDo = True
            End If
        End Select
        
        If blnDo Then
            Call AddReportRow(rsTmp!病人ID & "," & rsTmp!挂号ID, rsTmp!病人ID, rsTmp!NO, Nvl(rsTmp!姓名), Nvl(rsTmp!门诊号), Format(rsTmp!就诊时间 & "", "yyyy-MM-dd HH:mm"), _
                 Nvl(rsTmp!消息内容), rsTmp!类型编码 & "", rsTmp!优先程度 & "", Format(rsTmp!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), Nvl(rsTmp!业务标识), rsTmp!病人来源 & "", _
                 Nvl(rsTmp!险类, 0), rsTmp!挂号ID, rsTmp!ID)
            blnDo = False
        End If
        rsTmp.MoveNext
    Next
    rptNotify.Populate '缺省不选中任何行
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln消息语音 Then
        If mclsMsg Is Nothing Then
            Set mclsMsg = New clsCISMsg
            Call mclsMsg.InitCISMsg(0)
        End If
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Set mrsMsg = rsTmp
        End If
    End If
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AddReportRow(ParamArray arrInput() As Variant)
'功能：向消息提配列表中增加一行
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objItemIcon As ReportRecordItem
    Dim strRowID As String '提醒列表行的唯一标识，"病人id,主页id,消息编码"
    Dim strNO As String
    Dim str业务 As String
    Dim str病人来源 As String
    Dim int优先级 As Integer
    Dim int险类 As Integer
    Dim Index As Integer
    
    On Error GoTo errH
     
    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tag值 病人ID,挂号ID
    Set objItem = objRecord.AddItem(""): objItem.Icon = 6
    Set objItemIcon = objItem
    
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '病人id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  'NO
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '姓名
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '门诊号
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '就诊时间
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '状态，内容
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                            '消息编号
    objRecord.AddItem strNO: Index = Index + 1
    
    int优先级 = Val(arrInput(Index))                     '优先级
    objRecord.AddItem int优先级: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '登记日期
    
    str业务 = arrInput(Index): Index = Index + 1              '业务标识
    str病人来源 = arrInput(Index): Index = Index + 1          '病人来源
    
    int险类 = arrInput(Index): Index = Index + 1
    objRecord.AddItem str业务
    
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1   '挂号ID
    objRecord.AddItem Val(arrInput(Index)) '消息ID：业务消息清单.ID
    
    If int优先级 > 1 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            If int优先级 = 3 Then
                objRecord.Item(Index).ForeColor = &HC0&
            End If
            objRecord.Item(Index).Bold = True
        Next
    End If
    '保险病人用红色显示
    If int险类 > 0 And int优先级 <> 3 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            objRecord.Item(Index).ForeColor = &HC0&
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'功能：将接收到的消息加入提醒列表中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "select a.NO,a.姓名,a.执行人,a.门诊号,a.执行时间,a.险类 from 病人挂号记录 a where a.id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!就诊ID & ""))

    If mstr接诊医生 = rsTmp!执行人 & "" Then
        '判断列表是否已经有这类消息了
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_消息).Value = rsMsg!类型编码 And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!病人ID & "," & rsMsg!就诊ID) Then
                    Exit Sub
                End If
            End If
        Next
        
        Call AddReportRow(rsMsg!病人ID & "," & rsMsg!就诊ID, rsMsg!病人ID, rsMsg!NO, rsTmp!姓名, Nvl(rsTmp!门诊号), Format(rsTmp!执行时间 & "", "yyyy-MM-dd HH:mm"), Nvl(rsMsg!消息内容), _
             rsMsg!类型编码 & "", rsMsg!优先程度 & "", Format(rsMsg!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!业务标识 & "", rsMsg!病人来源 & "", Nvl(rsTmp!险类, 0), rsMsg!就诊ID, 0)
        
        rptNotify.Populate
         
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'功能：自动进入医嘱校对、确认停止的执行界面
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng病人ID As Long
    Dim lng医嘱ID As Long, lng挂号id As Long, lng消息ID As Long
    Dim str业务 As String, blnOk As Boolean
    Dim blnFinded As Boolean
    Dim strTmp As String
    Dim strNO As String
    Dim str挂号单 As String
    Dim str消息内容 As String
    Dim i As Long
    Dim strPatis As String
    Dim blnOnePati As Boolean
    Dim blnTmp As Boolean
    
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strNO = .Item(C_消息).Value
                str业务 = .Item(C_业务).Value
                str挂号单 = .Item(C_No).Value
                str消息内容 = .Item(C_状态).Value
                lng病人ID = Val(.Item(C_病人ID).Value)
                lng挂号id = Val(.Item(C_挂号ID).Value)
                lng消息ID = Val(.Item(C_ID).Value)
                lngIndex = .Index
            End With
    
            blnTmp = True
            
            If str挂号单 <> mstr挂号单 Then blnTmp = LocatePati(str挂号单)
            
            If strNO = "ZLHIS_RECIPEAUDIT_001" Then
                '先将卡片切换到医嘱卡片方便查找菜单
                Call LocatedCard("医嘱")
                cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                If str消息内容 = "处方审查合格。" Then
                    '弹出消息发送窗体
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_Send, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                Else
                    '医嘱编辑窗体
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_CIS_032" Then
                Call mclsDisease.ShowDisRegist(Me, 1, Val(str业务), lng病人ID, 0, str挂号单)
            End If
            
            If strNO = "ZLHIS_BLOOD_006" Then
                If gobjPublicBlood Is Nothing And gbln血库系统 Then InitObjBlood
                blnOk = gobjPublicBlood.zlIsBloodMessageDone(2, lng病人ID, lng挂号id, 1, mlng科室ID)
                If blnOk Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                Else
                    If FuncTraReaction(Val(str业务), mlngModul, False, IIf(InStr(1, str业务, ":") > 0, Val(Split(str业务, ":")(1)), 0)) Then
                        If gobjPublicBlood.zlIsBloodMessageDone(2, lng病人ID, lng挂号id, 1, mlng科室ID) Then
                            Call rptNotify.Records.RemoveAt(lngIndex)
                        End If
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_CIS_033" Then
            '传染病报告反修改消息阅读
                blnOk = ReadMsgCIS033(lng病人ID, lng挂号id, str业务, lng消息ID)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            
            If strNO <> "ZLHIS_CIS_033" And strNO <> "ZLHIS_BLOOD_006" Then
                blnOk = ReadMsg(lng病人ID, lng挂号id, strNO, str业务, lng消息ID, str挂号单)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            Call rptNotify.Populate
        End If
    End If
End Sub

Private Sub rptNotify_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptNotify_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub rptNotify_SelectionChanged()
    Dim str挂号单 As String
    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '非正常情况
    
    str挂号单 = rptNotify.SelectedRows(0).Record.Item(C_No).Value
 
    If str挂号单 <> mstr挂号单 Then Call LocatePati(str挂号单)
    
End Sub

Private Function ReadMsg(ByVal lng病人ID As Long, ByVal lng挂号id As Long, ByVal strNO As String, ByVal str业务 As String, ByVal lng消息ID As Long, ByVal str挂号单 As String) As Boolean
'功能：阅读消息
'说明：消息阅读方式目前有3种：按消息编译码阅读，消息ID阅读，按业务标识阅读
    Dim strSQL As String
    Dim lng科室ID As Long
    Dim str医嘱ID As String
    Dim blnDo As Boolean
    Dim lng危急值ID As Long  '本次处理的危急值记录ID
    Dim strSQLReadMsg As String
    Dim blnHis危急值 As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objControl As CommandBarControl
    
    If mlng接诊科室ID = 0 Then
        lng科室ID = UserInfo.部门ID
    Else
        lng科室ID = mlng接诊科室ID
    End If
    blnDo = True
    
    On Error GoTo errH
    
    strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng挂号id & ",'" & strNO & "',1,'" & UserInfo.姓名 & "'," & lng科室ID
    Select Case strNO
    Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
        strSQL = strSQL & ",null,null,'" & str业务 & "'"
    Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
        strSQL = strSQL & ",null," & lng消息ID
    End Select
    strSQL = strSQL & ")"
    
    strSQLReadMsg = strSQL
    
    If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
        If mbln危急值 Then
            '危急值消息相关处理
            Call mobjKernel.ShowDealCritical(Me, lng病人ID, 0, str挂号单, lng危急值ID)
            
            If lng危急值ID <> 0 Then
                '将消息设置为已阅
                Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
                '如果是LIS危急值调用LIS接口
                If strNO = "ZLHIS_LIS_003" Then
                    Call InitObjLis(p门诊医生站)
                    If Not gobjLIS Is Nothing Then
                        strSQL = "select a.标本id,a.处理情况,a.确认人 from 病人危急值记录 a where a.id=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng危急值ID)
                        If Not rsTmp.EOF Then
                            Call gobjLIS.WriteNotifyToLis(Val(rsTmp!标本ID & ""), rsTmp!确认人 & "", rsTmp!处理情况 & "")
                        End If
                    End If
                End If
            End If
            Call SetCriticalAdvice(lng危急值ID)
            blnHis危急值 = True
        End If
    End If
    
    If Not blnHis危急值 Then
        If strNO = "ZLHIS_LIS_003" Then
            If str业务 <> "" Then
                str医嘱ID = str业务
                Call InitObjLis(p门诊医生站)
                If Not gobjLIS Is Nothing Then
                    blnDo = gobjLIS.GetReadNotify(Me, str医嘱ID, UserInfo.姓名)
                End If
            End If
        End If
        If strNO = "ZLHIS_BLOOD_004" Then
            '用血审核消息的阅读状态设置在血库部件内部，临床不用执行阅读消息过程
            strSQL = "select 1 from 病人医嘱记录 a where a.挂号单=[1] and a.医嘱状态=1 and a.诊疗类别='K' and a.检查方法='1' and a.审核状态=1 and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str挂号单)
            If Not rsTmp.EOF Then
                '如果有数据，则弹出医嘱修改界面，本过程中不执行消息阅读SQL语句
                '先将卡片切换到医嘱卡片方便查找菜单
                Call LocatedCard("医嘱")
                cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                '医嘱编辑窗体
                Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
                If Not objControl Is Nothing Then
                    If objControl.Enabled Then objControl.Execute
                End If
                ReadMsg = True
                Exit Function
            End If
        End If
        If blnDo Then
            Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
        End If
    End If
    
    ReadMsg = blnDo
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LocatePati(ByVal strTag As String) As Boolean
'功能：通过挂号单定位，当前可以见的列表，在就诊列表和回诊列表中找。
    Dim blnTmp As Boolean
    Dim objLvw As ListView
    Dim i As Long
    Dim objItem As MSComctlLib.ListItem
    Dim lngIndex As Long
    
    Set objLvw = lvwPati(pt就诊)
    lngIndex = pt就诊
    For i = 1 To objLvw.ListItems.Count
        With objLvw.ListItems(i)
            If UCase(.Text) = UCase(strTag) Then
                objLvw.ListItems(i).Selected = True
                Set objItem = objLvw.ListItems(i)
                objLvw.SelectedItem.EnsureVisible
                blnTmp = True
                Exit For
            End If
        End With
    Next
    
    If Not blnTmp Then
        Set objLvw = lvwPatiHZ
        lngIndex = pt回诊
        For i = 1 To objLvw.ListItems.Count
            With objLvw.ListItems(i)
                If UCase(.Text) = UCase(strTag) Then
                    objLvw.ListItems(i).Selected = True
                    Set objItem = objLvw.ListItems(i)
                    objLvw.SelectedItem.EnsureVisible
                    blnTmp = True
                    Exit For
                End If
            End With
        Next
    End If
    
    If blnTmp Then
        If Not objLvw.Visible Then
            For i = 1 To dkpMain.PanesCount
                If dkpMain.Panes(i).Handle = objLvw.hwnd Then
                    dkpMain.Panes(i).Select
                End If
            Next
        End If
        If objLvw.Visible Then objLvw.SetFocus
    End If
    If blnTmp Then
        Call LvwItemClick(lngIndex, objItem)
    End If
    LocatePati = blnTmp
End Function

Private Sub mclsDisease_PatiTransfer(ByVal lng病人ID As Long, ByVal str挂号No As String)
'功能：传染病阳性界面触发事件转诊。
    Call ExecuteTransferSend
End Sub

Private Function GetOne阳性结果() As Long
'功能：获取一条指定的阳性结果
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    If mlng病人ID = 0 Then
        MsgBox "请选择一个病人。", vbInformation, gstrSysName
        Exit Function
    End If

   strSQL = "Select A.ID, '门诊' As 来源,a.记录状态, a.病人id,  b.姓名,  b.性别,  b.年龄, e.名称 As 科室, " & vbNewLine & _
        "b.门诊号 As 标识号, a.送检时间, a.送检医生, f.名称 As 登记科室, a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, a.登记人, a.登记时间, a.处理人, " & vbNewLine & _
        "a.处理时间, a.处理情况说明 " & vbNewLine & _
        "From 疾病阳性记录 A, 病人挂号记录 B,  部门表 E, 部门表 F " & vbNewLine & _
        "Where  a.病人id = b.病人id And a.挂号单 = b.No  And " & vbNewLine & _
        "a.登记科室ID = f.Id(+) And b.执行部门id = e.Id(+) And a.挂号单=[1] order by a.送检时间 desc"

   Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)

    If rsTmp.RecordCount = 0 Then
        MsgBox "该病人无阳性结果记录。", vbInformation, gstrSysName
        Exit Function
    ElseIf rsTmp.RecordCount = 1 Then
        GetOne阳性结果 = Val(rsTmp!ID & "")
        Exit Function
    End If
    
    GetOne阳性结果 = mclsDisease.ShowPatiDis(rsTmp, Me)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadMsgCIS033(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str标识 As String, ByVal lng消息ID As Long) As Boolean
'功能：传染病报告反修改消息阅读
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lng文件ID As Long
    Dim lng科室ID As Long
    Dim objControl As CommandBarControl
    
    On Error GoTo errH
 
    lng文件ID = Val(Split(str标识, ",")(0))
    
    strSQL = "Select 1 From 疾病申报记录 where 文件ID=[1] and 处理状态=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng文件ID, 4)
    If rsTmp.RecordCount = 0 Then
    '把消息标记为已读
        If mlng接诊科室ID = 0 Then
            lng科室ID = UserInfo.部门ID
        Else
            lng科室ID = mlng接诊科室ID
        End If
        strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.姓名 & "'," & lng科室ID & ",null," & lng消息ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        ReadMsgCIS033 = True
        Exit Function
    End If
    
    '弹出来修改报告
    Call mclsDisDoc.ModifyDiseaseDoc(Me, lng文件ID, mlng病人ID, mlng挂号ID, 1, mlng科室ID)
    
    strSQL = "Select 1 From 疾病申报记录 where 文件ID=[1] and 处理状态=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng文件ID, 4)
    If rsTmp.RecordCount = 0 Then
    '把消息标记为已读
        If mlng接诊科室ID = 0 Then
            lng科室ID = UserInfo.部门ID
        Else
            lng科室ID = mlng接诊科室ID
        End If
        strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.姓名 & "'," & lng科室ID & ",null," & lng消息ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        ReadMsgCIS033 = True
        Exit Function
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LocatedCard(ByVal strTag As String)
'功能：定位到指定的页签卡片，内部页签
    Dim i As Long
 
    If tbcSub.Selected.Tag <> strTag Then
        For i = 0 To tbcSub.ItemCount - 1
            If tbcSub.Item(i).Visible Then
                If tbcSub.Item(i).Tag = strTag Then
                    tbcSub.Item(i).Selected = True
                    Exit For
                End If
            End If
        Next
    End If
End Sub

Private Sub SetCriticalAdvice(ByVal lng记录ID As Long)
'功能：确认是危急值后弹出医嘱下达界面，刚才当前保存的医嘱与本次的记录进关联
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim objControl As Object
    
    On Error GoTo errH
    If lng记录ID = 0 Then Exit Sub
    strSQL = "select 1 from 病人危急值记录 a where a.id=[1] and a.是否危急值=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    
    If Not rsTmp.EOF Then
        '弹出下达医嘱的窗口
        If tbcSub.Tag <> "医嘱" Then
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub.Item(i).Visible Then
                    If tbcSub.Item(i).Tag = "医嘱" Then
                        tbcSub.Item(i).Selected = True
                        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                        Exit For
                    End If
                End If
            Next
        End If
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then
                objControl.Parameter = lng记录ID
                objControl.Execute
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExecuteCritical()
'功能：危急值相关处理
    Dim lng危急值ID As Long  '本次处理的危急值记录ID
    
    Call mobjKernel.ShowDealCritical(Me, mlng病人ID, 0, mstr挂号单, lng危急值ID)
    
    Call SetCriticalAdvice(lng危急值ID)
End Sub
