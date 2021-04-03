VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSublimeInNurseStation 
   Caption         =   "新版住院护士工作站"
   ClientHeight    =   8985
   ClientLeft      =   225
   ClientTop       =   255
   ClientWidth     =   13410
   Icon            =   "frmSublimeInNurseStation.frx":0000
   LinkTopic       =   "frmSublimeInNurseStation"
   ScaleHeight     =   8985
   ScaleWidth      =   13410
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgRPT 
      Left            =   11775
      Top             =   6555
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":18F2
            Key             =   "Pati"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":1E8C
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2426
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":29C0
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2F5A
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":34F4
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3F06
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4918
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4EB2
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":58C4
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":62D6
            Key             =   "审查归档"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":CB38
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":D0D2
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":D66C
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":E07E
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":E618
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":EBB2
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":F14C
            Key             =   "单病种"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":159AE
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":15F48
            Key             =   "紧急"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":164E2
            Key             =   "男人"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":1CD44
            Key             =   "女人"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picHLDJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4140
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   67
      Top             =   1995
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Timer timKey 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6120
      Top             =   30
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   11355
      MaxLength       =   100
      TabIndex        =   66
      Top             =   510
      Width           =   1000
   End
   Begin VB.PictureBox pic床位状况 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9510
      ScaleHeight     =   345
      ScaleWidth      =   1365
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Width           =   1365
      Begin VB.ComboBox cbo床位状况 
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   30
         Width           =   1365
      End
   End
   Begin MSComctlLib.ImageList imgHLDJ 
      Index           =   999
      Left            =   3360
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fra审查 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   9390
      TabIndex        =   27
      Top             =   6120
      Width           =   3360
      Begin VB.Label lbl审查 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "共有 XXX 条未处理的病案审查反馈..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   450
         MouseIcon       =   "frmSublimeInNurseStation.frx":235A6
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   75
         Width           =   3060
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   105
         Picture         =   "frmSublimeInNurseStation.frx":236F8
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.PictureBox pic主题过滤 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   4290
      ScaleHeight     =   345
      ScaleWidth      =   3855
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   510
      Width           =   3855
      Begin VB.ComboBox cbo内容 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   30
         Width           =   1365
      End
      Begin VB.ComboBox cbo主题 
         Height          =   300
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   30
         Width           =   1365
      End
      Begin VB.Label lbl内容 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "内容"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2040
         TabIndex        =   9
         Top             =   90
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标记"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   90
         Width           =   360
      End
   End
   Begin VB.CheckBox chk包含空床 
      Appearance      =   0  'Flat
      Caption         =   "包含空床"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8400
      TabIndex        =   11
      ToolTipText     =   "Ctrl+勾选：单独选择"
      Top             =   570
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.PictureBox pic病况 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2310
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   525
      Width           =   1755
      Begin VB.CheckBox chk病况条件 
         Appearance      =   0  'Flat
         Caption         =   "一般"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   4
         ToolTipText     =   "Ctrl+勾选：单独选择"
         Top             =   75
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.CheckBox chk病况条件 
         Appearance      =   0  'Flat
         Caption         =   "危"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   690
         TabIndex        =   5
         ToolTipText     =   "Ctrl+勾选：单独选择"
         Top             =   75
         Value           =   1  'Checked
         Width           =   465
      End
      Begin VB.CheckBox chk病况条件 
         Appearance      =   0  'Flat
         Caption         =   "重"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "Ctrl+勾选：单独选择"
         Top             =   75
         Value           =   1  'Checked
         Width           =   480
      End
   End
   Begin VB.PictureBox pic护理等级 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   30
      ScaleHeight     =   345
      ScaleWidth      =   2175
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   510
      Width           =   2175
      Begin VB.CommandButton cmd护理条件 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   1860
         Picture         =   "frmSublimeInNurseStation.frx":23C82
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   60
         Width           =   270
      End
      Begin VB.TextBox txt护理条件 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   30
         Width           =   2160
      End
   End
   Begin MSComctlLib.ImageList Img标记 
      Index           =   999
      Left            =   3360
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   44
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":23D78
            Key             =   "监护仪"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":240CA
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2441C
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2476E
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":24AC0
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":24E12
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":25164
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":254B6
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":25808
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":25B5A
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":25EAC
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":261FE
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":26550
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":268A2
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":26BF4
            Key             =   "预转科"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":26F46
            Key             =   "预出院"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":27298
            Key             =   "刀"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":275EA
            Key             =   "男孩"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2793C
            Key             =   "女孩"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":27C8E
            Key             =   "男人"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":27FE0
            Key             =   "女人"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":28332
            Key             =   "药"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":28684
            Key             =   "针"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":289D6
            Key             =   "盾牌"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":28D28
            Key             =   "铅笔"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2907A
            Key             =   "曲别针"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":293CC
            Key             =   "体温计"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2971E
            Key             =   "准备"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":29A70
            Key             =   "停止"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":29DC2
            Key             =   "正确"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2A114
            Key             =   "PDA"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2A466
            Key             =   "灯泡"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2A7B8
            Key             =   "提醒"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2AB0A
            Key             =   "红旗"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2AE5C
            Key             =   "禁止"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2B1AE
            Key             =   "手机"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2B500
            Key             =   "刷子"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2B852
            Key             =   "锁"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2BBA4
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2BEF6
            Key             =   "疑问"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2C248
            Key             =   "五角星"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2C59A
            Key             =   "胸花"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2C8EC
            Key             =   "病床"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2CC3E
            Key             =   "单病种"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img标记 
      Index           =   0
      Left            =   2790
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   44
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":334A0
            Key             =   "监护仪"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":33BB2
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":33F04
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":34256
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":345A8
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":348FA
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":34C4C
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":34F9E
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":352F0
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":35642
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":35994
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":35CE6
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":36038
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3638A
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":366DC
            Key             =   "预转科"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":36DEE
            Key             =   "预出院"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":37500
            Key             =   "刀"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":37C12
            Key             =   "男孩"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":38324
            Key             =   "女孩"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":38A36
            Key             =   "男人"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":39148
            Key             =   "女人"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3985A
            Key             =   "药"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":39F6C
            Key             =   "针"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3A67E
            Key             =   "盾牌"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3AD90
            Key             =   "铅笔"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3B4A2
            Key             =   "曲别针"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3BBB4
            Key             =   "体温计"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3C2C6
            Key             =   "准备"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3C9D8
            Key             =   "停止"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3D0EA
            Key             =   "正确"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3D7FC
            Key             =   "PDA"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3DF0E
            Key             =   "灯泡"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3E620
            Key             =   "提醒"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3ED32
            Key             =   "红旗"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3F444
            Key             =   "禁止"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3FB56
            Key             =   "手机"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":40268
            Key             =   "刷子"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4097A
            Key             =   "锁"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4108C
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4179E
            Key             =   "疑问"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":41EB0
            Key             =   "五角星"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":425C2
            Key             =   "胸花"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":42CD4
            Key             =   "病床"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":433E6
            Key             =   "单病种"
         EndProperty
      EndProperty
   End
   Begin VB.VScrollBar HScr 
      Height          =   5745
      LargeChange     =   25
      Left            =   13140
      Max             =   100
      SmallChange     =   5
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1920
      Picture         =   "frmSublimeInNurseStation.frx":49C48
      ScaleHeight     =   285
      ScaleWidth      =   1815
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   1845
   End
   Begin MSComctlLib.ImageList imgHLDJ 
      Index           =   0
      Left            =   2790
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer timNotify 
      Interval        =   500
      Left            =   480
      Top             =   15
   End
   Begin VB.PictureBox pic护理条件 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   30
      ScaleHeight     =   1695
      ScaleWidth      =   2115
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   855
      Visible         =   0   'False
      Width           =   2145
      Begin VB.ListBox lst护理条件 
         Appearance      =   0  'Flat
         Height          =   1080
         Left            =   -15
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   -15
         Width           =   2145
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   1530
         Picture         =   "frmSublimeInNurseStation.frx":4B78E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "取消"
         Top             =   1320
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   990
         Picture         =   "frmSublimeInNurseStation.frx":4BD18
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "确认"
         Top             =   1320
         Width           =   450
      End
   End
   Begin VB.Timer timeRefreshCard 
      Interval        =   100
      Left            =   30
      Top             =   15
   End
   Begin VB.ComboBox cboUnit 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "cboUnit"
      Top             =   195
      Width           =   1905
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   13215
      TabIndex        =   17
      Top             =   840
      Width           =   13215
      Begin VB.Label lblInpatientArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病区基本信息:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   90
         Width           =   11475
      End
   End
   Begin VB.PictureBox PicDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7515
      ScaleWidth      =   13335
      TabIndex        =   23
      Top             =   1200
      Width           =   13395
      Begin VB.PictureBox picPati 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3240
         Index           =   0
         Left            =   180
         Picture         =   "frmSublimeInNurseStation.frx":4C2A2
         ScaleHeight     =   3240
         ScaleWidth      =   2640
         TabIndex        =   47
         Top             =   1275
         Visible         =   0   'False
         Width           =   2640
         Begin VB.Label lblMedPay 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "城镇职工基本医疗保险"
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
            Index           =   0
            Left            =   60
            TabIndex        =   50
            Top             =   2250
            Width           =   840
         End
         Begin VB.Label lbl病情 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   0
            Left            =   2130
            TabIndex        =   49
            Top             =   1920
            Width           =   105
         End
         Begin VB.Label lblCardNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1000123456"
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
            Index           =   0
            Left            =   1305
            TabIndex        =   53
            Top             =   2250
            Width           =   1050
         End
         Begin VB.Image img单病种 
            Height          =   360
            Index           =   0
            Left            =   2190
            Picture         =   "frmSublimeInNurseStation.frx":69510
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   360
         End
         Begin VB.Image img护理等级 
            Appearance      =   0  'Flat
            Height          =   360
            Index           =   0
            Left            =   2170
            Picture         =   "frmSublimeInNurseStation.frx":6FD62
            Stretch         =   -1  'True
            Top             =   38
            Width           =   345
         End
         Begin VB.Label lbl结余 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "欠款金额"
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
            Index           =   0
            Left            =   60
            TabIndex        =   63
            Top             =   2835
            Width           =   840
         End
         Begin VB.Label lbl床号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "09123"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   30
            TabIndex        =   62
            Top             =   450
            Width           =   825
         End
         Begin VB.Label lblSplit 
            BackColor       =   &H0000FF00&
            Height          =   60
            Index           =   0
            Left            =   30
            TabIndex        =   61
            Top             =   750
            Width           =   2475
         End
         Begin VB.Label lbl住院号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "027647132"
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
            Index           =   0
            Left            =   60
            TabIndex        =   60
            Top             =   930
            Width           =   945
         End
         Begin VB.Label lbl性别 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "男"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   0
            Left            =   1260
            TabIndex        =   59
            Top             =   945
            Width           =   195
         End
         Begin VB.Label lbl年龄 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "33"
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
            Index           =   0
            Left            =   1650
            TabIndex        =   58
            Top             =   930
            Width           =   210
         End
         Begin VB.Label lbl医师 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "医护:徐文举/李泽霞"
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
            Height          =   215
            Index           =   0
            Left            =   60
            TabIndex        =   57
            Top             =   1590
            Width           =   2415
         End
         Begin VB.Label lbl入院日期 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2010-06-09"
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
            Index           =   0
            Left            =   60
            TabIndex        =   56
            Top             =   2535
            Width           =   1050
         End
         Begin VB.Label lbl诊断 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎"
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
            Index           =   0
            Left            =   60
            TabIndex        =   55
            Top             =   1260
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.Label lbl结余总额 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "34998.48"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   1320
            TabIndex        =   54
            Top             =   2835
            Width           =   1020
         End
         Begin VB.Label lbl护士 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "费别:自费"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   52
            Top             =   1920
            Width           =   945
         End
         Begin VB.Label lbl住院天数 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "25天"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   0
            Left            =   1920
            TabIndex        =   51
            Top             =   2535
            Width           =   420
         End
         Begin VB.Image img个性标记2 
            Height          =   360
            Index           =   0
            Left            =   1440
            Picture         =   "frmSublimeInNurseStation.frx":70464
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img个性标记1 
            Height          =   360
            Index           =   0
            Left            =   1095
            Picture         =   "frmSublimeInNurseStation.frx":70B66
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img出院 
            Height          =   360
            Index           =   0
            Left            =   735
            Picture         =   "frmSublimeInNurseStation.frx":71268
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img临床路径 
            Height          =   360
            Index           =   0
            Left            =   375
            Picture         =   "frmSublimeInNurseStation.frx":7196A
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img病案审查 
            Height          =   360
            Index           =   0
            Left            =   30
            Picture         =   "frmSublimeInNurseStation.frx":7206C
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img个性标记3 
            Height          =   360
            Index           =   0
            Left            =   1800
            Picture         =   "frmSublimeInNurseStation.frx":7276E
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Label lbl姓名 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "李四王麻中华人民共和国"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   1020
            TabIndex        =   48
            Top             =   450
            Width           =   1500
         End
         Begin VB.Label lblSelect 
            BackColor       =   &H00FFC0C0&
            Height          =   330
            Index           =   0
            Left            =   30
            TabIndex        =   64
            Top             =   420
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.Label lbl房间号 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   180
            Index           =   0
            Left            =   2160
            TabIndex        =   69
            Top             =   960
            Visible         =   0   'False
            Width           =   90
         End
      End
      Begin VB.PictureBox picPati 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2820
         Index           =   999
         Left            =   2895
         Picture         =   "frmSublimeInNurseStation.frx":72E70
         ScaleHeight     =   2820
         ScaleWidth      =   2235
         TabIndex        =   29
         Top             =   1665
         Visible         =   0   'False
         Width           =   2235
         Begin VB.Label lblMedPay 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "城镇职工基本医疗保险"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   35
            Top             =   1935
            Width           =   720
         End
         Begin VB.Label lbl病情 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   999
            Left            =   1740
            TabIndex        =   30
            Top             =   1620
            Width           =   105
         End
         Begin VB.Label lblCardNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1000123456"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   1080
            TabIndex        =   36
            Top             =   1935
            Width           =   900
         End
         Begin VB.Image img单病种 
            Height          =   240
            Index           =   999
            Left            =   1860
            Picture         =   "frmSublimeInNurseStation.frx":87EB2
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label lbl房间号 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Height          =   180
            Index           =   999
            Left            =   1800
            TabIndex        =   70
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lbl床号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "09123"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   999
            Left            =   30
            TabIndex        =   45
            Top             =   375
            Width           =   675
         End
         Begin VB.Label lblSplit 
            BackColor       =   &H008080FF&
            Height          =   60
            Index           =   999
            Left            =   30
            TabIndex        =   44
            Top             =   630
            Width           =   2040
         End
         Begin VB.Image img护理等级 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   999
            Left            =   1850
            Picture         =   "frmSublimeInNurseStation.frx":8E704
            Stretch         =   -1  'True
            Top             =   30
            Width           =   240
         End
         Begin VB.Label lbl住院号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "027647132"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   43
            Top             =   840
            Width           =   810
         End
         Begin VB.Label lbl性别 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "男"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   999
            Left            =   1065
            TabIndex        =   42
            Top             =   840
            Width           =   180
         End
         Begin VB.Label lbl年龄 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "33"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   1365
            TabIndex        =   41
            Top             =   840
            Width           =   180
         End
         Begin VB.Label lbl医师 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "医护:徐文举/李泽霞"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   40
            Top             =   1380
            Width           =   1995
         End
         Begin VB.Label lbl入院日期 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2010-06-09"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   39
            Top             =   2205
            Width           =   900
         End
         Begin VB.Label lbl诊断 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎慢性支气管炎"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   38
            Top             =   1110
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.Label lbl结余 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "欠款金额"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   37
            Top             =   2475
            Width           =   720
         End
         Begin VB.Label lbl结余总额 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "34998.48"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   999
            Left            =   960
            TabIndex        =   34
            Top             =   2475
            Width           =   1020
         End
         Begin VB.Label lbl护士 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "费别:自费"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   33
            Top             =   1650
            Width           =   810
         End
         Begin VB.Label lbl姓名 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "李四王麻中华人民共和国"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   999
            Left            =   840
            TabIndex        =   32
            Top             =   375
            Width           =   1275
         End
         Begin VB.Label lbl住院天数 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "25天"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   999
            Left            =   1605
            TabIndex        =   31
            Top             =   2205
            Width           =   360
         End
         Begin VB.Image img个性标记2 
            Height          =   240
            Index           =   999
            Left            =   1260
            Picture         =   "frmSublimeInNurseStation.frx":8EA46
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img个性标记1 
            Height          =   240
            Index           =   999
            Left            =   960
            Picture         =   "frmSublimeInNurseStation.frx":8ED88
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img出院 
            Height          =   240
            Index           =   999
            Left            =   660
            Picture         =   "frmSublimeInNurseStation.frx":8F0CA
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img临床路径 
            Height          =   240
            Index           =   999
            Left            =   360
            Picture         =   "frmSublimeInNurseStation.frx":8F40C
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img病案审查 
            Height          =   240
            Index           =   999
            Left            =   60
            Picture         =   "frmSublimeInNurseStation.frx":8F74E
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img个性标记3 
            Height          =   240
            Index           =   999
            Left            =   1560
            Picture         =   "frmSublimeInNurseStation.frx":8FA90
            Top             =   60
            Width           =   240
         End
         Begin VB.Label lblSelect 
            BackColor       =   &H00FFC0C0&
            Height          =   330
            Index           =   999
            Left            =   30
            TabIndex        =   46
            Top             =   330
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   2325
         Index           =   0
         Left            =   90
         TabIndex        =   71
         Top             =   435
         Width           =   5610
         _Version        =   589884
         _ExtentX        =   9895
         _ExtentY        =   4101
         _StockProps     =   0
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   2325
         Index           =   3
         Left            =   75
         TabIndex        =   72
         Top             =   390
         Width           =   5610
         _Version        =   589884
         _ExtentX        =   9895
         _ExtentY        =   4101
         _StockProps     =   0
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picPara 
         BorderStyle     =   0  'None
         Height          =   2715
         Index           =   1
         Left            =   3810
         ScaleHeight     =   2715
         ScaleWidth      =   5970
         TabIndex        =   76
         Top             =   15
         Width           =   5970
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   2325
            Index           =   2
            Left            =   -15
            TabIndex        =   89
            Top             =   390
            Width           =   5610
            _Version        =   589884
            _ExtentX        =   9895
            _ExtentY        =   4101
            _StockProps     =   0
            BorderStyle     =   1
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.CheckBox chkSettle 
            Caption         =   "已结清"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   2400
            TabIndex        =   86
            Top             =   90
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.CheckBox chkSettle 
            Caption         =   "未结清"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   3405
            TabIndex        =   87
            Top             =   90
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.PictureBox picPara 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   2
            Left            =   30
            ScaleHeight     =   345
            ScaleWidth      =   2250
            TabIndex        =   84
            Top             =   15
            Visible         =   0   'False
            Width           =   2250
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Left            =   795
               Style           =   2  'Dropdown List
               TabIndex        =   85
               Top             =   10
               Width           =   1440
            End
            Begin VB.Label lbl出院时间 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院时间"
               Height          =   180
               Left            =   0
               TabIndex        =   88
               Top             =   60
               Width           =   720
            End
         End
      End
      Begin VB.PictureBox picPara 
         BorderStyle     =   0  'None
         Height          =   2715
         Index           =   0
         Left            =   120
         ScaleHeight     =   2715
         ScaleWidth      =   5625
         TabIndex        =   77
         Top             =   405
         Width           =   5625
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   2325
            Index           =   1
            Left            =   75
            TabIndex        =   83
            Top             =   -1170
            Width           =   5610
            _Version        =   589884
            _ExtentX        =   9895
            _ExtentY        =   4101
            _StockProps     =   0
            BorderStyle     =   1
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.PictureBox picPara 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   320
            Index           =   3
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   78
            Top             =   45
            Visible         =   0   'False
            Width           =   3855
            Begin VB.TextBox txtChange 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Left            =   780
               MaxLength       =   3
               TabIndex        =   81
               Text            =   "7"
               Top             =   0
               Width           =   285
            End
            Begin VB.Frame fraChange 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Left            =   750
               TabIndex        =   80
               Top             =   210
               Width           =   300
            End
            Begin VB.CommandButton cmdRef 
               Caption         =   "刷新"
               Height          =   255
               Left            =   2520
               TabIndex        =   79
               Top             =   0
               Width           =   615
            End
            Begin VB.Label lbl转出 
               AutoSize        =   -1  'True
               Caption         =   "显示最近    天的转出病人"
               Height          =   180
               Left            =   15
               TabIndex        =   82
               Top             =   30
               Width           =   2160
            End
         End
      End
      Begin VB.PictureBox pic出院查找 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2880
         ScaleHeight     =   315
         ScaleWidth      =   2325
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   4485
         Width           =   2325
         Begin VB.TextBox txt住院号 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00C0C0C0&
            Height          =   300
            Left            =   825
            MaxLength       =   100
            TabIndex        =   74
            ToolTipText     =   "根据住院号定位病人"
            Top             =   0
            Width           =   1485
         End
         Begin VB.Label lblPatiInputType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院号↓"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   90
            TabIndex        =   90
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.Frame fraPatiUD 
         BorderStyle     =   0  'None
         Height          =   45
         Left            =   2640
         MousePointer    =   7  'Size N S
         TabIndex        =   68
         Top             =   6000
         Width           =   6120
      End
      Begin VB.PictureBox picList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2625
         Left            =   240
         ScaleHeight     =   2625
         ScaleWidth      =   12315
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4860
         Width           =   12315
      End
      Begin XtremeSuiteControls.TabControl PatiPage 
         Height          =   2565
         Left            =   60
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   15
         Width           =   4755
         _Version        =   589884
         _ExtentX        =   8387
         _ExtentY        =   4524
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   8625
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSublimeInNurseStation.frx":8FDD2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19129
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人颜色"
            TextSave        =   "病人颜色"
            Key             =   "病人颜色"
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
   Begin MSComctlLib.ImageList imgIcon 
      Index           =   0
      Left            =   120
      Top             =   7830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":90664
            Key             =   "监护仪"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":90DDE
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":91558
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":91CD2
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9244C
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":92BC6
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":93340
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":93ABA
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":94234
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":949AE
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":95128
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":958A2
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9601C
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":96796
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":96F10
            Key             =   "手术刀"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9768A
            Key             =   "床"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":97E04
            Key             =   "男孩"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9857E
            Key             =   "女孩"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":98CF8
            Key             =   "男人"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":99472
            Key             =   "女人"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":99BEC
            Key             =   "药"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9A366
            Key             =   "针"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9AAE0
            Key             =   "盾牌"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9B25A
            Key             =   "铅笔"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9B9D4
            Key             =   "曲别针"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9C14E
            Key             =   "体温计"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9C8C8
            Key             =   "准备"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9D042
            Key             =   "停止"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9D7BC
            Key             =   "完成"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcon 
      Index           =   999
      Left            =   690
      Top             =   7830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9DF36
            Key             =   "监护仪"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9E2D0
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9E66A
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9EA04
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9ED9E
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9F138
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9F4D2
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9F86C
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9FC06
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9FFA0
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A033A
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A06D4
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A0A6E
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A0E08
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A11A2
            Key             =   "手术刀"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A153C
            Key             =   "床"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A18D6
            Key             =   "男孩"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A1C70
            Key             =   "女孩"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A200A
            Key             =   "男人"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A23A4
            Key             =   "女人"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A273E
            Key             =   "药"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A2AD8
            Key             =   "针"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A2E72
            Key             =   "盾牌"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A320C
            Key             =   "铅笔"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A35A6
            Key             =   "曲别针"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A3940
            Key             =   "体温计"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A3CDA
            Key             =   "准备"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A4074
            Key             =   "停止"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A440E
            Key             =   "完成"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic卡片背景 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   5340
      ScaleHeight     =   4245
      ScaleWidth      =   7515
      TabIndex        =   65
      Top             =   1350
      Visible         =   0   'False
      Width           =   7515
      Begin VB.Image img卡片背景 
         Height          =   2880
         Index           =   4
         Left            =   3300
         Picture         =   "frmSublimeInNurseStation.frx":A47A8
         Top             =   30
         Width           =   2235
      End
      Begin VB.Image img卡片背景 
         Height          =   3315
         Index           =   5
         Left            =   4740
         Picture         =   "frmSublimeInNurseStation.frx":B97EA
         Top             =   45
         Width           =   2685
      End
      Begin VB.Image img卡片背景 
         Height          =   945
         Index           =   3
         Left            =   2910
         Picture         =   "frmSublimeInNurseStation.frx":D6A58
         Top             =   3210
         Width           =   2685
      End
      Begin VB.Image img卡片背景 
         Height          =   840
         Index           =   2
         Left            =   0
         Picture         =   "frmSublimeInNurseStation.frx":DEF7E
         Top             =   3210
         Width           =   2235
      End
      Begin VB.Image img卡片背景 
         Height          =   2985
         Index           =   1
         Left            =   645
         Picture         =   "frmSublimeInNurseStation.frx":E51C0
         Top             =   0
         Width           =   2685
      End
      Begin VB.Image img卡片背景 
         Height          =   2595
         Index           =   0
         Left            =   0
         Picture         =   "frmSublimeInNurseStation.frx":FF5C6
         Top             =   0
         Width           =   2235
      End
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   2325
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmSublimeInNurseStation.frx":1124C8
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1920
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSublimeInNurseStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PATI_TYPE
    pt入院待入住 = 0
    pt转科待入住 = 1
    pt转病区待入住 = 2
    pt在院 = 3
    pt家庭病床 = 3.1
    pt预转科 = 3.2
'    pt转病区 = 3.3
    pt预出 = 4
    pt出院 = 5
    pt死亡 = 6
    pt最近转出 = 7
End Enum
Private Enum EFun
    E入住 = 0
    E转科 = 1
    E换床 = 2
    E包房 = 3
    E出院 = 4
    E转为住院 = 5
    E更改床位等级 = 6
    E调整病人信息 = 7
    E新生儿登记 = 8
    E重算费用 = 9
    E医保病种选择 = 10
    E撤销 = 11
    E修改出院时间 = 12
    E床位对换 = 13
    E转医疗小组 = 14
    E转病区 = 15
    E转病区入住 = 16
    E病人备注编辑 = 17
End Enum
Private Enum PATI_COLUMN
    c_类型 = 0
    c_审查 = 1
    c_图标 = 2
    c_路径状态 = 3
    C_病人ID = 4
    C_主页ID = 5
    c_姓名 = 6
    c_住院号 = 7
    c_床号 = 8
    c_性别 = 9
    c_年龄 = 10
    c_费别 = 11
    c_付款方式 = 12
    c_医生 = 13
    c_入院日期 = 14
    c_出院日期 = 15
    c_病人类型 = 16
    c_就诊卡号 = 17
End Enum

Private Const mstrColWidth As String = "0,16,18,18,0,0,80,80,50,50,50,120,120,70,130,130,100,100"
        
Private Enum EFun_医嘱提醒
    E发送 = 0
    E校对 = 1
    E停止 = 2
    E查看 = 3
End Enum

Private Const clngX = 100

Private Const 卡片背景_标准卡片 As Integer = 0
Private Const 卡片背景_大卡片 As Integer = 1
Private Const 卡片背景_标准卡片_折叠 As Integer = 2
Private Const 卡片背景_大卡片_折叠 As Integer = 3
Private Const 卡片背景_标准卡片_就诊卡 As Integer = 4
Private Const 卡片背景_大卡片_就诊卡 As Integer = 5

Private Const clngBaseHeight_Normal = 2595  '标准卡片未折叠时的高度
Private Const clngBigHeight_Normal = 2985   '大卡片未折叠时的高度
Private Const clngBaseCardHeight_Normal = 2880  '标准卡片未折叠时的高度（显示就诊卡）
Private Const clngBigCardHeight_Normal = 3315   '大卡片未折叠时的高度（显示就诊卡）
'用色带的颜色来表示病人类型时
Private Const clngBaseHeight_Collapse = 825 '标准卡片折叠时的高度
Private Const clngBigHeight_Collapse = 920  '大卡片折叠时的高度

'todo:执行以下功能时,弹出病人事务处理模块,最多50个自定义模块
Private Const conMenu_病人事务处理 = 990000
Private Const conMenu_查看医嘱 = 990001
Private Const conMenu_查看费用 = 990002
Private Const conMenu_查看病历 = 990003
Private Const conMenu_查看体温单 = 990004
Private Const conMenu_查看护理记录 = 990005
Private Const conMenu_查看护理病历 = 990006

Private Const conMenu_图标 = 990050                     '标注所使用的图标ID从990050开始,最多150个图标
Private Const conMenu_标注1 = 990200
Private Const conMenu_标注2 = 990300
Private Const conMenu_标注3 = 990400
Private Const conMenu_标注结束 = 990500
Private Const conMenu_Manage_BedExchange = 2613         '*床位对换
Private Const conMenu_Edit_AnimalHeat = 3035            '*批量录入体温单
Private Const conMenu_Edit_NurseLogFile = 3036          '*批量录入记录单
Private Const conMenu_ProveCollect = 3037               '检验采集工作站
Private Const conMenu_Edit_BatExecute = 3098            '*医嘱批量执行

Private mPatiInfo As PatiInfo

'子窗体对象定义
Private mclsAdvices As zlPublicAdvice.clsDockInAdvices
Private mclsTends As zl9TendFile.clsTendFile
Private WithEvents mclsFeeQuery As zl9InExse.clsFeeQuery
Attribute mclsFeeQuery.VB_VarHelpID = -1
Private WithEvents mfrmResponse As frmAuditResponse '审查反馈窗口
Attribute mfrmResponse.VB_VarHelpID = -1
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mfrmNoticeBoard As frmNoticeBoard  '病人公告栏窗口
Attribute mfrmNoticeBoard.VB_VarHelpID = -1
Private mclsInPatient As zl9InPatient.clsInPatient
Private mclsWardMonitor As clsWardMonitor     '监护仪接口
Private mcolSubForm As Collection
Private mobjProveCollect As Object
Private mobjPlugIn As Object
Private mlngPlugInID As Long
Private mrsPlugInBar As ADODB.Recordset '菜单样式结构见 zlPlugIn/mdlPlugIn/ 中 GetBarInfo 方法
'54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
Private mclsInOutMedRec As zlMedRecPage.clsInOutMedRec

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

'参数设置变量
Private blnUnload As Boolean
Private mstrPrivs As String
Private mstrPrivs_检验采集 As String
Private mlngModul As Long
Private mstrUnits As String
Private mstrScope As String
Private mintFindType As Integer
Private mintPatiInputType As Integer  '出院病人查找
Private mintChange As Integer
Private mintPage As Integer             '最小一个有效的页面
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mintOutPreTime As Integer
Private mintNotify As Integer           '医嘱提醒自动刷新间隔(分钟)
Private mintNotifyDay As Integer        '提醒多少天内的医嘱
Private mstrNotifyAdvice As String      '提醒的医嘱类型
Private mstrCardInfo As String          '卡片显示内容
Private mblnCardBalance As Boolean      '卡片余额是否包含担保金额
Private mblnCardOrder As Boolean         '卡片排序是否按照床位号排序
Private mblnCollateAutoFind As Boolean  '医嘱处理后自动定位到医嘱页面
Public mintREPORTSEL As Integer        '当前选择非在床清单索引
Private mstrNoteItems As String         '所有个性主题的内容,如:准备手术,开始手术,手术结束|男孩,女孩

Private mblnMonitor As Boolean          '监护仪程序是否存在
Private mstrMonitor As String           '监护仪程序路径
Private mstrBoardKeys As String         '病区公告栏返回的重新组装的信息

'以下两个变量只记录在床病人的信息
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlngPre病人ID As Long
Private mlngPre主页ID As Long
Private mblnReturn As Boolean           '病区按钮
'控制变量
Private mintCards As Integer            '显示的床位卡片数
Public mblnRoutine As Boolean           '是否加载病人事务管理模块
Private mstrSQL As String
Private mintPreDept As Integer          '上一病区
Private mblnShow As Boolean             '决定是否显示完整的卡片内容
Private mblnRefresh As Boolean          '决定是否刷新病区床位一览表
Private mlngSelect As Long              '当前选择的卡片索引
Private mlngSource As Long              '记录当前是标准卡片还是大卡片
Private mbytFontSize As Byte             '字体信息9号字体12号字体
Private mblnStart As Boolean            '程序是否正常启动
Private mblnCardCollapse As Boolean     '卡片是否折叠
Private mdblScaleHeight As Double       '床位卡区域实际高度
Private mblnHScroll As Boolean          '纵向滚动条是否显示
Private mblnOutDept As Boolean          '是否仅服务于门诊的科室（门诊留观病人显示门诊号）
Private mblnShowCard As Boolean         '是否显示就诊卡号
Private mblnHavePath As Boolean          '当前病区是否具有可查看的临床路径

Private mobjPopup As CommandBarPopup    '右键弹出菜单\病人入出
Private mobjPopupBatch As CommandBarPopup    '右键弹出菜单\病区批量工作
Private mobjTheme As CommandBarControl  '主题过滤
Private mobjFilter As CommandBar

'病区基本信息
Private mlng空床 As Long
Private mlng在床 As Long
Private mlng入院 As Long
Private mlng转入 As Long
Private mlng家床 As Long
Private mlng出院 As Long
Private mlng预出院 As Long
Private mlng转出 As Long
Private mlng死亡 As Long
Private mlng手术 As Long
Private mlng危 As Long
Private mlng重 As Long

'内部记录集及相关变量
Private mstrFields As String
Private mstrValues As String
Private mrsBedInfo As New ADODB.Recordset   '当前病区床位信息
Private mrsPatiColor As New ADODB.Recordset '病人类型设置
Public mrsPatiInfo As New ADODB.Recordset  '病人记录集保留
Private mrsNotes As New ADODB.Recordset     '病区自已设定的标记内容
Private mrsPatiNotes As New ADODB.Recordset '病区所有病人的标记清单
Private mintMecStandard As Integer  '病案首页格式 0-卫生部标准，1-四川省标准，2-云南省标准
Private mlngMedRedDay As Long     '病案审查反馈天数

Dim mstrBriefCode As String
Dim mblnSupport As Boolean

Private Enum 页面
    待入科
    转科
    出院
    家庭病床
End Enum


'加载护理等级颜色
Private Const ALTERNATE = 1
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" _
    (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreatePen Lib "gdi32" _
    (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Polyline Lib "gdi32" _
    (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'设定一个窗体捕获鼠标，即所有鼠标输入消息都发往该窗体
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'取消鼠标捕获
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private mlngColor As Long
Private mintIndex As Long
Private mobjFileSys As New FileSystemObject

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReMoveCtrol()
    Dim objCtrl As Object
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objFilter As CommandBar
    
    
    '设置条件大小
    lst护理条件.Height = lst护理条件.ListCount * 210 + 30
    pic护理条件.Height = lst护理条件.Height + cmdFilterOK.Height + 120
    pic护理条件.Visible = False
    
    pic病况.Height = TextHeight("刘") + 60
    chk病况条件(0).Left = 0
    chk病况条件(0).Top = (pic病况.Height - chk病况条件(0).Height) \ 2
    If chk病况条件(0).Top < 0 Then chk病况条件(0).Top = 0
    chk病况条件(1).Left = chk病况条件(0).Left + chk病况条件(0).Width
    chk病况条件(1).Top = chk病况条件(0).Top
    chk病况条件(2).Left = chk病况条件(1).Left + chk病况条件(1).Width
    chk病况条件(2).Top = chk病况条件(0).Top
    pic病况.Width = chk病况条件(2).Left + chk病况条件(2).Width
    
    Label1.Top = cbo主题.Top + (cbo主题.Height - Label1.Height) \ 2
    cbo主题.Left = Label1.Left + Label1.Width + 50
    lbl内容.Left = cbo主题.Left + cbo主题.Width + TextWidth("刘") / 2
    lbl内容.Top = Label1.Top
    cbo内容.Left = lbl内容.Left + lbl内容.Width + 50
    cbo内容.Top = cbo主题.Top
    pic主题过滤.Width = cbo内容.Left + cbo内容.Width + 30
    chk包含空床.Width = TextWidth("刘鹏" & chk包含空床.Caption) - TextWidth("刘") / 3
    txtFind.Width = 6 * TextWidth("刘")
    
    '重新绑定下控件
    Set objFilter = cbsMain.Add("过滤工具栏", xtpBarTop)   '固有
    objFilter.EnableDocking xtpFlagStretched
    objFilter.ContextMenuPresent = False
    With objFilter.Controls
        Set objControl = .Add(xtpControlLabel, 1, "护理等级")
        Set objCustom = .Add(xtpControlCustom, 2, "")
        objCustom.Handle = pic护理等级.hwnd
        Set objControl = .Add(xtpControlLabel, 3, "床位状况"): objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, 4, "")
        objCustom.Handle = pic床位状况.hwnd
        Set objControl = .Add(xtpControlLabel, 5, "当前病况"): objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, 6, "")
        objCustom.Handle = pic病况.hwnd
        Set objCustom = .Add(xtpControlCustom, 7, ""): objCustom.BeginGroup = True
        objCustom.Handle = pic主题过滤.hwnd
        Set objCustom = .Add(xtpControlCustom, 8, "")
        objCustom.Handle = chk包含空床.hwnd: objCustom.BeginGroup = True

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "↓按床号查找")
        objPopup.Caption = "↓按床号查找"
        objPopup.ID = conMenu_View_FindType
        objPopup.Style = xtpButtonCaption
        objPopup.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, 10, "")
        objCustom.Handle = txtFind.hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With

    For Each objCtrl In mobjFilter.Controls
        objCtrl.Delete
    Next
    mobjFilter.Delete
    Set mobjFilter = objFilter
    '页面转科
    fraChange.Left = lbl转出.Left + TextWidth("页面转科")
    fraChange.Top = lbl转出.Height + lbl转出.Top
    fraChange.Width = TextWidth("转科")
    txtChange.Width = TextWidth("999")
    txtChange.Left = fraChange.Left + (fraChange.Width - txtChange.Width) / 2
    txtChange.Height = TextHeight("刘")
    txtChange.Top = fraChange.Top - txtChange.Height
    cmdRef.Left = lbl转出.Left + lbl转出.Width + 100
    cmdRef.Height = TextHeight("刘") + 100
    cmdRef.Width = TextWidth(" 刷新 ")
    cmdRef.Top = lbl转出.Top - (cmdRef.Height - lbl转出.Height) \ 2
    
    '出院查询
    cboSelectTime.Left = lbl出院时间.Left + lbl出院时间.Width + TextWidth("刘") / 2
    picPara(2).Width = cboSelectTime.Left + cboSelectTime.Width + TextWidth("刘")
    picPara(2).Height = (cboSelectTime.Top * 2) + cboSelectTime.Height
    chkSettle(0).Left = picPara(2).Width + 100
    If (picPara(2).Height - TextWidth("刘")) \ 2 >= 0 Then
        chkSettle(0).Top = (picPara(2).Height - TextWidth("刘")) \ 2
    End If
    chkSettle(1).Left = chkSettle(0).Left + chkSettle(0).Width + 100
    chkSettle(1).Top = chkSettle(0).Top
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    Dim lngCol As Long, lngIndex As Long, arrWidth() As String
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Call frmNotify.SetFontSize(bytSize)
    
    Me.FontSize = mbytFontSize
    Me.FontName = "宋体"
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            Select Case UCase(objCtrl.Name)
                Case UCase("Label1"), UCase("lbl内容"), UCase("lblInpatientArea"), UCase("lbl出院时间"), UCase("lbl审查"), UCase("lbl转出"), UCase("Label2"), _
                    UCase("lbl转出"), UCase("lblPatiInputType")
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("刘") + 20
            End Select
        Case UCase("ListBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("VsFlexGrid")
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("刘鹏" & objCtrl.Caption) - TextWidth("刘") / 3
        Case UCase("CheckBox")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("刘鹏" & objCtrl.Caption) - TextWidth("刘") / 3
        Case UCase("DTPicker")
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01") + 400
            objCtrl.Height = TextHeight("刘") * 1.5
        Case UCase("TextBox")
            objCtrl.FontSize = mbytFontSize
            If bytSize = 0 Then
                objCtrl.Height = 300
            End If
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = mbytFontSize
        End Select
    Next
    
    '病人列表列宽设置
    arrWidth = Split(mstrColWidth, ",")
    For lngIndex = 0 To rptPati.UBound
        For lngCol = c_图标 To rptPati(lngIndex).Columns.Count - 1
            rptPati(lngIndex).Columns.Column(lngCol).Width = Val(arrWidth(lngCol)) + (Val(arrWidth(lngCol)) * IIf(bytSize = 0, 0, 1)) \ 3
        Next lngCol
        rptPati(lngIndex).Redraw
    Next lngIndex
    
    Call Form_Resize
    Call ReMoveCtrol
End Sub

Private Sub InitSelectTime()
    
    mdtOutEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtOutBegin = mdtOutEnd
    
    cboSelectTime.Clear '出院
    With cboSelectTime
        .AddItem "今天内"
        .ItemData(.NewIndex) = 0
        .AddItem "昨天内"
        .ItemData(.NewIndex) = 1
        .AddItem "前天内"
        .ItemData(.NewIndex) = 2
        .AddItem "一周内"
        .ItemData(.NewIndex) = 7
        .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "60天内"
        .ItemData(.NewIndex) = 60
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
End Sub

Private Sub cboSelectTime_Click()
'功能：当时间范围是指定是，弹出时间选择窗体
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If cboSelectTime.ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboSelectTime) Then
            '取消时恢复原来的选择
            Call zlControl.CboSetIndex(cboSelectTime.hwnd, mintOutPreTime)
            Exit Sub
        End If
    Else
        mdtOutEnd = datCurr
        mdtOutBegin = mdtOutEnd - intDateCount
    End If
    If mdtOutBegin = CDate(0) Or mdtOutEnd = CDate(0) Then
        cboSelectTime.ToolTipText = ""
    Else
        cboSelectTime.ToolTipText = "范围：" & Format(mdtOutBegin, "yyyy-MM-dd") & " 至 " & Format(mdtOutEnd, "yyyy-MM-dd")
    End If
    '保存参数，保证每个地方提取的出院病人都是在同一时间范围内（72783）
    Call zlDatabase.SetPara("出院病人结束间隔", DateDiff("d", datCurr, mdtOutEnd), glngSys, p住院护士站)
    Call zlDatabase.SetPara("出院病人开始间隔", DateDiff("d", mdtOutBegin, datCurr), glngSys, p住院护士站)
    mintOutPreTime = cboSelectTime.ListIndex
    rptPati(PatiPage.Selected.Index).Tag = ""
    rptPati(PatiPage.Selected.Index).Records.DeleteAll
    If rptPati(PatiPage.Selected.Index).Columns.Count > c_审查 Then rptPati(PatiPage.Selected.Index).Columns(c_审查).Visible = False
    Call PatiPage_SelectedChanged(PatiPage.Selected)
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
'    Dim lngidx As Long
    Dim rsTmp As New ADODB.Recordset
    
    On Error Resume Next
    '先关闭所有计时器,再打开按键延时记时器(不关就无法输入匹配)
    If KeyAscii <> 13 Then
        timKey.Enabled = False
        timNotify.Enabled = False
        timeRefreshCard.Enabled = False
        timKey.Interval = 1000
        timKey.Enabled = True
    End If

    mblnReturn = False
    If cboUnit.ListIndex <> -1 Then mintPreDept = cboUnit.ListIndex
    If KeyAscii = 13 Then
        mblnReturn = True
        KeyAscii = 0
        If cboUnit.Text <> "" Then
            Set rsTmp = GetDataToUnits(cboUnit.Text)
            If Not rsTmp.EOF Then
                Call FindCboIndex(cboUnit, rsTmp!ID)
            Else
                cboUnit.ListIndex = mintPreDept
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            timKey.Tag = cboUnit.ListIndex
        Else
            cboUnit.ListIndex = mintPreDept
            timKey.Tag = mintPreDept
        End If
    End If
End Sub

Private Sub cboUnit_Validate(Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call zlControl.CboSetIndex(cboUnit.hwnd, mintPreDept)
    End If
End Sub

Private Sub cbo床位状况_Click()
    If mblnHScroll Then
        If HScr.Value <> 0 Then
            mstrBoardKeys = ""
            HScr.Value = 0
        Else
            Call AdjustCard
        End If
    Else
        Call AdjustCard
    End If
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim rsPatiLog As ADODB.Recordset
    Dim i As Long, j As Long, strPrivs As String
    Dim objControl As CommandBarControl
    
    If CommandBar.Parent Is Nothing Then Exit Sub
        
    'Call CommandBar.Controls.DeleteAll
        
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "床  号(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "住院号(&2)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "就诊卡(&3)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 4, "姓  名(&4)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 5, "简  码(&5)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 9, "清除"
            End If
        End With
    Case conMenu_File_MedRecPrint
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 1, "正面(&1)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 2, "反面(&2)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 3, "附页1(&3)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 4, "附页2(&4)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 5, "正面+附页1(&5)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 6, "反面+附页2(&6)"
            End If
        End With
    Case conMenu_File_MedRecPreview
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 1, "正面(&1)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 2, "反面(&2)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 3, "附页1(&3)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 4, "附页2(&4)"
            End If
        End With
    Case conMenu_Manage_Change_Undo
        With CommandBar.Controls
            .DeleteAll
            If Not LocatePatiRecord Then Exit Sub
            
            Set rsPatiLog = GetPatiLog(mrsPatiInfo!病人ID, mrsPatiInfo!主页ID)
            If rsPatiLog.RecordCount > 0 Then '动态子菜单,扩1位
                
                strPrivs = GetInsidePrivs(Enum_Inside_Program.p病人入出)
                rsPatiLog.MoveFirst
                For i = 1 To rsPatiLog.RecordCount
                    If Not IsNull(rsPatiLog!终止时间) And rsPatiLog!终止原因 = 1 Then
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Undo * 10 + i, "出院")
                        j = j + 1
                        If InStr(";" & strPrivs & ";", ";撤消出院;") = 0 Or j > 1 Then objControl.Enabled = False
                    Else
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Undo * 10 + i, rsPatiLog!操作)
                        If rsPatiLog.RecordCount > 1 And rsPatiLog!开始原因 = 1 Then objControl.Visible = False
                        j = j + 1
                        If j > 1 Then
                            objControl.Enabled = False
                        Else
                            If (objControl.Caption Like "*入住" Or objControl.Caption = "转病区入住") Then
                                If InStr(strPrivs, "撤消入科") = 0 Then objControl.Enabled = False
                            End If
                            If objControl.Caption = "转为住院病人" Then
                                If InStr(strPrivs, "住院留观转住院") = 0 Then objControl.Enabled = False
                            ElseIf objControl.Caption = "预出院" Then
                                If InStr(strPrivs, "病人预出院") = 0 Then objControl.Enabled = False
                                
                            ElseIf objControl.Caption = "换床" Then
                                If InStr(strPrivs, "换床") = 0 Then objControl.Enabled = False
                            End If
                        End If
                    End If
                    objControl.Category = "撤销"
                    If i <> 1 Then objControl.Enabled = False
                    rsPatiLog.MoveNext
                Next
            End If
        End With
    Case conMenu_Tool_PlugInPop
        If Not mrsPlugInBar Is Nothing Then
            mrsPlugInBar.Filter = "IsInTool=0 and BarType=3"
            If mrsPlugInBar.RecordCount > 0 Then
                With CommandBar.Controls
                    .DeleteAll
                    For i = 1 To mrsPlugInBar.RecordCount
                        Set objControl = .Add(xtpControlButton, mrsPlugInBar!功能ID, mrsPlugInBar!菜单名)
                            objControl.IconId = mrsPlugInBar!图标ID
                            objControl.Parameter = mrsPlugInBar!功能名
                            objControl.Style = xtpButtonIconAndCaption
                        If Val(mrsPlugInBar!IsGroup) = 1 Then
                            objControl.BeginGroup = True
                        End If
                        mrsPlugInBar.MoveNext
                    Next
                End With
            End If
            mrsPlugInBar.Filter = 0
        End If
    End Select
End Sub

Private Sub chkSettle_Click(Index As Integer)
    '68259:刘鹏飞,2012-02-11,出院病人查找添加未结清已结清功能
    If chkSettle(0).Value = 0 And chkSettle(1).Value = 0 Then
        chkSettle((Index + 1) Mod 2).Value = 1
    End If
    rptPati(PatiPage.Selected.Index).Tag = ""
    rptPati(PatiPage.Selected.Index).Records.DeleteAll
    If rptPati(PatiPage.Selected.Index).Columns.Count > c_审查 Then rptPati(PatiPage.Selected.Index).Columns(c_审查).Visible = False
    Call PatiPage_SelectedChanged(PatiPage.Selected)
End Sub

Private Sub chk病况条件_GotFocus(Index As Integer)
    mintREPORTSEL = -1
End Sub

Private Sub cmdRef_Click()
'54436:刘鹏飞,2012-10-10
    Call txtChange_KeyPress(vbKeyReturn)
End Sub

Private Sub fraPatiUD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And picList.Visible = True Then
        fraPatiUD.Tag = 0
    End If
End Sub

Private Sub fraPatiUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And picList.Visible = True Then
        If fraPatiUD.Top + y < picPati(mlngSource).Height + 10 Or picList.Height - y < 2000 Then Exit Sub
        fraPatiUD.Top = fraPatiUD.Top + y
        picList.Top = fraPatiUD.Top
        picList.Height = PicDraw.Height - picList.Top
        PatiPage.Height = picList.Height - 60
        Me.Refresh
        fraPatiUD.Tag = 1
        Call Form_Resize
    Else
        fraPatiUD.Tag = 0
    End If
End Sub

Private Sub fraPatiUD_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And picList.Visible = True Then
        If Val(fraPatiUD.Tag) = 1 Then
            Call HScr_Change
            fraPatiUD.Tag = 0
        End If
    End If
End Sub

'61824:刘鹏飞,2013-05-23,显示单病种标志
Private Sub img单病种_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img单病种(Index).Left + x, img单病种(Index).Top + y)
End Sub

Private Sub img单病种_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img单病种(Index).Tag, True
End Sub

Private Sub img单病种_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lblCardNo_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblCardNo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblCardNo(Index).Left + x, lblCardNo(Index).Top + y)
End Sub

Private Sub lblCardNo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, "就诊卡号：" & lblCardNo(Index).Caption, True
End Sub

Private Sub lblCardNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lblInpatientArea_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picInfo.hwnd, lblInpatientArea.Caption, True
End Sub

Private Sub lblMedPay_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblMedPay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblMedPay(Index).Left + x, lblMedPay(Index).Top + y)
End Sub

Private Sub lblMedPay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, "医疗付款方式：" & lblMedPay(Index).Caption, True
End Sub

Private Sub lblMedPay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '49752,刘鹏飞,2012-09-05,出院病人提供多钟查找方式(床号、住院号、就诊卡、姓名)
    If Button = vbRightButton Then Exit Sub
   
    '弹出菜单
    Dim intType As Integer
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Set cbrPopupBar = Me.cbsMain.Add("弹出菜单", xtpBarPopup)
    intType = mintPatiInputType
    '床号、住院号、就诊卡、姓名、简码
    With cbrPopupBar
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_View_FindType * 100# + 11, "床  号(&1)")
        If intType = 10 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_View_FindType * 100# + 12, "住院号(&2)")
        If intType = 11 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_View_FindType * 100# + 13, "就诊卡(&3)")
        If intType = 12 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_View_FindType * 100# + 14, "姓  名(&4)")
        If intType >= 13 Then cbrPopupItem.Checked = True
    End With
    cbrPopupBar.ShowPopup
End Sub

Private Sub lbl结余总额_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, Trim(lbl结余总额(Index).Caption), True
End Sub

Private Sub lbl审查_Click()
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    '非模态显示审查反馈窗体
    If mfrmResponse Is Nothing Then
        Set mfrmResponse = New frmAuditResponse
    End If
    
    Call mfrmResponse.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex), 1, False, 1, mstrPrivs)
End Sub

Private Sub cboUnit_Click()

    mblnReturn = True
    If cboUnit.ListIndex = mintPreDept Then Exit Sub
    mintPreDept = cboUnit.ListIndex
    
    mlngSelect = -1
    mblnRefresh = True
    mintREPORTSEL = -1
    
    '关闭业务窗体
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
    End If
    
    '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
    If Not mclsInOutMedRec Is Nothing Then
        Call mclsInOutMedRec.FormUnLoad
    End If

    Call Have部门性质(cboUnit.ItemData(cboUnit.ListIndex), "护理", mblnOutDept)
    With frmNotify
        .mintNotify = mintNotify
        .mintNotifyDay = mintNotifyDay
        .mstrNotifyAdvice = mstrNotifyAdvice
        .mdtOutBegin = mdtOutBegin
        .mdtOutEnd = mdtOutEnd
        .mlng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
    End With
    frmNotify.mblnFirst = True
End Sub

Private Sub cbo内容_Click()
    Dim strInfo As String
    
    mintREPORTSEL = -1
    If Not mblnStart Then Exit Sub
    '更新条件
    strInfo = "所有主题"
    If Me.cbo主题.Text <> "所有" Then
        strInfo = cbo主题.Text
        
        If Me.cbo内容.Text <> "所有" Then
            strInfo = strInfo & "\" & Me.cbo内容.Text
        End If
    End If
    
    '刷新病区床位一览表
    If mblnHScroll Then
        If HScr.Value <> 0 Then
            mstrBoardKeys = ""
            HScr.Value = 0
        Else
            Call AdjustCard
        End If
    Else
        Call AdjustCard
    End If
End Sub

Private Sub cbo主题_Click()
    Dim arrData
    Dim strData As String
    Dim i As Integer, j As Integer
    
    mintREPORTSEL = -1
    Me.cbo内容.Clear
    Me.cbo内容.AddItem "所有"
    If Me.cbo主题.Text <> "所有" Then
        strData = Split(Me.cbo主题.Tag, "|")(Me.cbo主题.ListIndex - 1)
        If strData <> "" Then
            arrData = Split(strData, ",")
            j = UBound(arrData)
            For i = 0 To j
                '个性标记内容存储的是说明'标记序号
                If InStr(1, arrData(i), "'") <> 0 Then
                    Me.cbo内容.AddItem Split(arrData(i), "'")(0)
                    Me.cbo内容.ItemData(cbo内容.NewIndex) = Val(Split(arrData(i), "'")(1))
                Else
                    Me.cbo内容.AddItem arrData(i)
                End If
            Next
        End If
    End If
    Me.cbo内容.ListIndex = 0
    Me.cbo内容.Enabled = (Me.cbo内容.ListCount > 1)
    Me.cbo内容.BackColor = IIf(Me.cbo内容.Enabled, &H80000005, &HC0C0C0)
End Sub

Private Function LocatePatiRecord() As Boolean
    Dim intIndex As Integer
    Dim strTag As String
    Dim blnTrue As Boolean
    '根据当前的活动控件来定位病人
    
    '122993
    If mrsPatiInfo.State = adStateClosed Then Exit Function
    If mintREPORTSEL = -1 Then
        If mlng病人ID = 0 Then Exit Function
        mrsPatiInfo.Filter = "病人ID=" & mlng病人ID & " And 主页ID=" & mlng主页ID ' & " And (排序 >=3 and 排序<=3)"
        blnTrue = mrsPatiInfo.RecordCount
    Else
        intIndex = mintREPORTSEL
        If rptPati(intIndex).SelectedRows.Count = 0 Then GoTo ErrNext
        If rptPati(intIndex).SelectedRows(0).Record Is Nothing Then GoTo ErrNext
        If rptPati(intIndex).SelectedRows(0).Childs.Count > 0 Then GoTo ErrNext
        strTag = rptPati(intIndex).SelectedRows(0).Record.Tag
        mrsPatiInfo.Filter = "病人ID=" & Split(strTag, "|")(0) & " And 主页ID=" & Split(strTag, "|")(1)
        blnTrue = mrsPatiInfo.RecordCount
    End If
    '53740:刘鹏飞,2012-09-19,如果选择的不是病人卡片或者没有选中任何病人，取消卡片的选中
ErrNext:
    If mintREPORTSEL <> -1 Or blnTrue = False Then
        If mlngSelect >= 0 Then
            '包床也一并取消选中
            With mrsBedInfo
                .Filter = "卡片索引=" & mlngSelect
                If !病人ID <> 0 Then
                    If PicDraw.Enabled And PicDraw.Visible Then PicDraw.SetFocus
                    .Filter = "病人ID=" & !病人ID
                    Do While Not .EOF
                        '将选择状态清除,同时将卡片大小还原(有可能在折叠模式下)
                        picPati(!卡片索引).ZOrder 1
                        lblSelect(!卡片索引).Visible = False
                        If mblnCardCollapse Then
                            picPati(!卡片索引).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
                            picPati(!卡片索引).Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_折叠, 卡片背景_标准卡片_折叠)).Picture
                        End If
                        
                        .MoveNext
                    Loop
                End If
                .Filter = 0
            End With
            picPati(mlngSelect).ZOrder 0
            mlngSelect = -1
            mlng病人ID = 0: mlng主页ID = 0
        End If
    End If
    
    LocatePatiRecord = blnTrue
End Function

Private Sub InNurseRoutine(Optional ByVal strPage As String = "医嘱")
    '54408:刘鹏飞,2012-10-10,传入病人信息记录集
    Call frmInNurseRoutine.zlInitMip(mclsMipModule)
    Call frmInNurseRoutine.NurseRoutine(Me, mstrPrivs, Me.cboUnit.ItemData(Me.cboUnit.ListIndex), _
         Val(mrsPatiInfo.Fields("病人ID").Value), mdtOutBegin, mdtOutEnd, mintChange, mstrScope, mPatiInfo, strPage, mrsPatiInfo, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
End Sub

Private Sub RefreshPatiList_Rountine()
    If Not mblnRoutine Then Exit Sub
    Call frmInNurseRoutine.RefreshPatiList(mrsPatiInfo)
End Sub

Private Sub OrientTabPage_Rountine(Optional ByVal strPage As String = "医嘱", Optional ByVal strID As String = "")
    '-------------------------------------------------------------
    '功能:定位到病人事物中指定的页面,以及对应页面指定的文件或医嘱等
    '-------------------------------------------------------------
    '55430:刘鹏飞,2013-02-27,双击作废医嘱定位到病人事物的医嘱页面
    If Not mblnRoutine Then Exit Sub
    Call frmInNurseRoutine.OrientTabPage(strPage, strID)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, byt入住方式 As Byte, str床号 As String, int床位索引 As Integer
    Dim strPrivs_病人入出 As String, strPrivs_护理 As String, strParentTitle As String, strTmp As String
    Dim blnExecuted As Boolean              '已执行则退出
    Dim blnHotKey As Boolean
    Dim objControl As Object
    Dim strKey As String, arrTag, strNote As String
    Dim arrSQL
    On Error GoTo ErrHand
    '功能说明:只有打印床头卡功能是和卡片选择相关,其他功能有可能是在床病人,也可能是不在床病人
    
    If Control.ID = conMenu_File_Exit Then
        Unload Me
        Exit Sub
    End If
    
    '如果是标注菜单,执行完即退出
    If Control.ID > conMenu_标注1 And Control.ID < conMenu_标注结束 Then
        If Not LocatePatiRecord Then Exit Sub
        mrsBedInfo.Filter = "病人ID=" & mrsPatiInfo!病人ID & " And 包床=0"
        If mrsBedInfo.RecordCount = 0 Then
            mrsBedInfo.Filter = ""
            Exit Sub
        End If
        arrTag = Split(Control.Category, "|")
        str床号 = mrsBedInfo!床号
        int床位索引 = mrsBedInfo!卡片索引
        strKey = ""
        If Val(arrTag(0)) = 1 And Nvl(mrsBedInfo!个性标注1) <> "" Then
            strKey = Split(mrsBedInfo!个性标注1, ",")(0) & "," & Split(mrsBedInfo!个性标注1, ",")(1)
        ElseIf Val(arrTag(0)) = 2 And Nvl(mrsBedInfo!个性标注2) <> "" Then
            strKey = Split(mrsBedInfo!个性标注2, ",")(0) & "," & Split(mrsBedInfo!个性标注2, ",")(1)
        Else
            If Nvl(mrsBedInfo!个性标注3) <> "" Then
                strKey = Split(mrsBedInfo!个性标注3, ",")(0) & "," & Split(mrsBedInfo!个性标注3, ",")(1)
            End If
        End If
        mrsBedInfo.Filter = ""
        '保存数据
        arrSQL = Array()
        If arrTag(3) <> 0 And strKey <> "" Then
            '更新主题图标则先删除原有的设置,可能设置的组发生变化
            If strKey <> arrTag(1) & "," & arrTag(2) Then
                mstrSQL = "ZL_病区标记记录_UPDATE(" & Me.cboUnit.ItemData(Me.cboUnit.ListIndex) & "," & Val(mrsPatiInfo.Fields("病人ID").Value) & "," & _
                    Val(mrsPatiInfo.Fields("主页ID").Value) & "," & Split(strKey, ",")(1) & "," & 0 & "," & arrTag(0) & IIf(Val(Split(strKey, ",")(0)) = 0, "", "," & Split(strKey, ",")(0)) & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mstrSQL
            End If
        End If
        mstrSQL = "ZL_病区标记记录_UPDATE(" & Me.cboUnit.ItemData(Me.cboUnit.ListIndex) & "," & Val(mrsPatiInfo.Fields("病人ID").Value) & "," & _
                Val(mrsPatiInfo.Fields("主页ID").Value) & "," & arrTag(2) & "," & arrTag(3) & "," & arrTag(0) & IIf(Val(arrTag(1)) = 0, "", "," & arrTag(1)) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = mstrSQL
        
        For i = 0 To UBound(arrSQL)
            If CStr(arrSQL(i)) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "更新病人标记")
        Next
        
        strKey = arrTag(1) & "," & arrTag(2) & "," & arrTag(3) & "," & arrTag(4)
        strNote = arrTag(5)
        '更新内部记录集
        If Val(arrTag(0)) = 1 Then
            Call Record_Update(mrsBedInfo, "个性标注1|个性标注1名称", strKey & "|" & strNote, "床号|" & Trim(str床号))
        ElseIf Val(arrTag(0)) = 2 Then
            Call Record_Update(mrsBedInfo, "个性标注2|个性标注2名称", strKey & "|" & strNote, "床号|" & Trim(str床号))
        Else
            Call Record_Update(mrsBedInfo, "个性标注3|个性标注3名称", strKey & "|" & strNote, "床号|" & Trim(str床号))
        End If
        '更新卡片
        Call SetCardLabel(int床位索引)
        
        Exit Sub
    End If
    
    strPrivs_病人入出 = GetInsidePrivs(Enum_Inside_Program.p病人入出)
    strPrivs_护理 = GetInsidePrivs(Enum_Inside_Program.p护理记录管理)
    
    '快捷键方式调入,父对象为空(只考虑病区批量工作下的功能菜单)
    If Control.Parent Is Nothing Then
        Select Case Control.ID
        '61762:刘鹏飞,2013-05-20,增加发送输液药品医嘱的功能
        Case conMenu_Edit_PreBalance, conMenu_Edit_Audit, conMenu_Edit_Send, conMenu_Edit_SendInfusion, conMenu_Report_Reports, conMenu_Report_DrugQuery, conMenu_Edit_SendBack, _
             conMenu_File_PrintMultiBill, conMenu_Edit_BatExecute, conMenu_Edit_AnimalHeat, conMenu_Edit_NurseLogFile
             strParentTitle = "病区批量工作"
        End Select
    Else
        strParentTitle = Control.Parent.Title
    End If
    If strParentTitle = "右键菜单" Then
        Select Case Control.ID
        Case conMenu_Edit_ReStop, conMenu_Manage_ReportLisView
            strParentTitle = "医嘱业务"
        Case conMenu_Edit_Billing, conMenu_Edit_ReBillingApply
            strParentTitle = "费用业务"
        End Select
    End If
    
    '外挂菜单
    If Control.ID > conMenu_Tool_PlugIn_Item And Control.ID < conMenu_Tool_PlugIn_Item + 100 Then '外挂功能执行
        If Not mobjPlugIn Is Nothing Then
            If Not LocatePatiRecord Then
                Call mobjPlugIn.ExecuteFunc(glngSys, P新版护士站, Control.Parameter, 0, 0, 0, , 1)
            Else
                Call mobjPlugIn.ExecuteFunc(glngSys, P新版护士站, Control.Parameter, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, , 1)
            End If
        End If
    End If
    
    '批量事务菜单
    If strParentTitle <> "" Then
        '按快捷键执行功能时，传入的按钮对象应该是控件自动创建的，没有父对象
        
        If strParentTitle = "病区批量工作" Then
            '54409:刘鹏飞,2012-09-25,病区批量工作没有选择病人也可以使用(除病人事务处理外)
            Select Case Control.ID
            Case conMenu_Edit_PreBalance                '预结算
                If LocatePatiRecord Then
                    Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("科室ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
                Else
                    Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, 0, 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
                End If
            Case conMenu_File_PrintMultiBill            '催款管理（新）
                Call mclsFeeQuery.zlPatiPressMoney(Me, gcnOracle, glngSys, mlngModul, gstrDBUser, mstrPrivs, cboUnit.ItemData(cboUnit.ListIndex), Split(cboUnit.Text, "-")(1))
            Case conMenu_Edit_BatExecute, conMenu_Manage_ThingAudit '执行登记（新）、执行核对
                If Not LocatePatiRecord Then mrsPatiInfo.Filter = ""
                If mrsPatiInfo.RecordCount > 0 Then
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("科室ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                Else
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, 0, 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                End If
            Case conMenu_Edit_AnimalHeat                '批量录入体温单（新）
                On Error Resume Next
                Dim strDLL As String
                Dim strSQL As String
                Dim objChart As Object
                Dim rsTemp As New ADODB.Recordset
                
                strSQL = " Select 新部件 From 体温部件 Where Nvl(启用,0)=1"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取体温部件")
                If err <> 0 Then
                    strDLL = "zl9TemperatureChart"
                Else
                    If rsTemp.RecordCount = 0 Then
                        strDLL = "zl9TemperatureChart"
                    Else
                        strDLL = Nvl(rsTemp!新部件, "zl9TemperatureChart")
                    End If
                End If
                
                err = 0
                strDLL = strDLL & ".clsBodyEditor"
                Set objChart = CreateObject(strDLL)
                If err <> 0 Then
                    MsgBox "    创建体温部件失败！" & vbCrLf & "    程序将创建标准的体温部件进行数据展现，请检查指定的体温部件是否存在或已损坏！" & vbCrLf & "    详细错误：" & err.Description, vbInformation, gstrSysName
                    
                    '如果创建指定的体温部件出错则创建标准的体温部件，因为这里不处理的话，后面可能存在直接使用体温部件中的对象，从而导致程序崩溃
                    strDLL = "zl9TemperatureChart.clsBodyEditor"
                    Set objChart = CreateObject(strDLL)
                End If
                
                On Error GoTo ErrHand
                Call objChart.InitBodyEditor(glngSys, gcnOracle)
                Call objChart.BodyMutilEditor(Me, cboUnit.ItemData(cboUnit.ListIndex), strPrivs_护理, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
            Case conMenu_Edit_NurseLogFile              '批量录入记录单（新）
                Call mclsTends.TendFileMutilEditor(Me, cboUnit.ItemData(cboUnit.ListIndex), strPrivs_护理, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
            Case conMenu_病人事务处理                   '病人事务处理（新）
                Call InNurseRoutine
            Case conMenu_ProveCollect                   '检验采集工作站
                If mobjProveCollect Is Nothing Then
                    On Error Resume Next
                    Set mobjProveCollect = CreateObject("zl9LisWork.clsLisWork")
                    If err <> 0 Then Exit Sub
                End If
                On Error GoTo ErrHand
                Call mobjProveCollect.CodeMan(glngSys, 1211, gcnOracle, Me, gstrDBUser)
            Case conMenu_Edit_BatUnPack '批量打包
                mclsAdvices.zlCompoundUnpack Me, cboUnit.ItemData(cboUnit.ListIndex), mlng病人ID, cboUnit.ItemData(cboUnit.ListIndex)
            Case conMenu_Tool_RisPrintBat '批量打印预约单
                mclsAdvices.AdviceRisReport Me, cboUnit.ItemData(cboUnit.ListIndex)
            Case Else   '医嘱校对、医嘱发送、医嘱暂停、医嘱启用、医嘱确认停止、病区常用报表（打印执行单）、超期收回(conMenu_Edit_Audit, conMenu_Edit_Send,conMenu_Edit_Pause,conMenu_Edit_Reus,conMenu_Edit_ReStop, conMenu_Report_Reports, conMenu_Report_DrugQuery, conMenu_Edit_SendBack)
                If Not LocatePatiRecord Then mrsPatiInfo.Filter = ""
                Call mclsAdvices.SetFontSize(IIf(mbytFontSize = 12, 1, 0))
                If mrsPatiInfo.RecordCount = 0 Then
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, 0, 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                Else
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("科室ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                End If
            End Select
            blnExecuted = True
        ElseIf strParentTitle = "医嘱业务" Then
            If Control.ID = conMenu_View_Notify Then
                With frmNotify
                    .mintNotify = mintNotify
                    .mintNotifyDay = mintNotifyDay
                    .mstrNotifyAdvice = mstrNotifyAdvice
                End With
                frmNotify.mblnFirst = True
            Else
                If Not LocatePatiRecord Then Exit Sub
                If Control.ID = conMenu_查看医嘱 Then
                    Call InNurseRoutine
                Else
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("科室ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                End If
            End If
            blnExecuted = True
        ElseIf strParentTitle = "费用业务" Then
            If Control.ID <> conMenu_Manage_Change_ReCalcFee Then
                If Not LocatePatiRecord Then Exit Sub
                If Control.ID = conMenu_查看费用 Then
                    Call InNurseRoutine("费用")
                Else
                    Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("科室ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
                End If
                blnExecuted = True
            End If
        ElseIf strParentTitle = "护理业务" Or strParentTitle = "病历业务" Then
            Call InNurseRoutine(Mid(strParentTitle, 1, 2))
            blnExecuted = True
        End If
    End If
    If blnExecuted Then Exit Sub
    
    Select Case Control.ID
    '---------------------------------------------------------------
    '管理菜单，病人入出转
    Case conMenu_Manage_Change_In
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        If mrsPatiInfo!排序 = pt转病区待入住 Then
            mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转病区入住, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), "", 0)
        ElseIf mrsPatiInfo!排序 = pt转科待入住 Then
            mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E入住, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), "", _
                    Val(mrsPatiInfo.Fields("科室ID").Value), 1)
        Else
            mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E入住, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), "", _
                    Val(mrsPatiInfo.Fields("科室ID").Value), 0)
        End If
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_Turn
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转科, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_TurnUnit
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转病区, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_TurnTeam
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转医疗小组, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_Bed
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E换床, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, "", "")
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_TransposeBed
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E床位对换, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), Nvl(mrsPatiInfo.Fields("床号").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_House
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E换床, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 1, "", "")
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_Out
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E出院, Me, strPrivs_病人入出, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_InPati
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E转为住院, Me, strPrivs_病人入出, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), _
        Val(mrsPatiInfo.Fields("住院号").Value), CStr(mrsPatiInfo.Fields("姓名").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_BedGrid
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E更改床位等级, Me, strPrivs_病人入出, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), _
        Trim(CStr(Nvl(mrsPatiInfo.Fields("床号").Value))))
    Case conMenu_Manage_Change_PatiInfo
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E调整病人信息, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
    Case conMenu_Manage_Change_PaitNote
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E病人备注编辑, Me, strPrivs_病人入出, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
    Case conMenu_Manage_Change_Baby
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E新生儿登记, Me, strPrivs_病人入出, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
    Case conMenu_Manage_Change_ReCalcFee
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E重算费用, Me, strPrivs_病人入出, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), _
        CStr(mrsPatiInfo.Fields("姓名").Value))
    Case conMenu_Manage_Change_InsureSel
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E医保病种选择, Me, strPrivs_病人入出, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), Val(mrsPatiInfo.Fields("险类").Value))
    Case conMenu_Manage_Change_Undo * 10 + 1
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E撤销, Me, strPrivs_病人入出, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), Val(mrsPatiInfo.Fields("险类").Value), Control.Caption)
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Monitor '监护仪
        Call InNurseRoutine("监护")
    '---------------------------------------------------------------
    
    '其他功能
    Case conMenu_Tool_Archive '电子病案查阅
        If Not LocatePatiRecord Then Exit Sub
        Call frmArchiveView.ShowArchive(Me, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
    Case conMenu_View_Warrant '担保信息查阅
        If Not LocatePatiRecord Then Exit Sub
        Call frmPatiSurety.ShowMe(Me, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '诊疗措施参考
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    Case conMenu_Manage_FeeItemSet  '诊疗项目费用设置
        Call Set诊疗项目费用设置
    Case conMenu_Tool_MedRecAuditResponse '审查反馈
        '都可以调用，至少可以查看(当前或历史)
        Call lbl审查_Click
'    Case conMenu_Tool_UnitSubject '病区标记设置
'         Call frmUnitSubjectSet.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex), mstrPrivs)
'         If gblnOK Then mblnRefresh = True
    Case conMenu_Tool_UnitNBoard
        If frmNoticeBoardSet.ShowMe(Me, mstrPrivs, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
            If Not mfrmNoticeBoard Is Nothing Then
                If mfrmNoticeBoard.mblnShow = True Then Call mfrmNoticeBoard.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex))
            End If
        End If
    '基础功能
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For Each objControl In Me.cbsMain(2).Controls
            If objControl.ID <> conMenu_View_Find Then
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            End If
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_FontSize_S      '标准卡片 小字体
        mlngSource = 999
        lbl床号(mlngSource).Tag = lbl床号(0).Tag
        Call SetSourceCardH
        mblnRefresh = True
        Call SetFontSize(0)
    Case conMenu_View_FontSize_L      '大卡片 大字体
        mlngSource = 0
        lbl床号(mlngSource).Tag = lbl床号(999).Tag
        Call SetSourceCardH
        mblnRefresh = True
        Call SetFontSize(1)
    Case conMenu_View_Expend_AllCollapse    '卡片折叠
        mblnCardCollapse = mblnCardCollapse Xor True
        Call SetSourceCardH
        mblnRefresh = True
    Case conMenu_View_Expend_CurCollapse      '非在床病人
        picList.Visible = picList.Visible Xor True
        PatiPage.Visible = picList.Visible
        Call picPatiIn_Resize
        If picList.Visible Then
            fra审查.Left = picList.Width - fra审查.Width
            fra审查.Top = PicDraw.Top + picList.Top + 50
        Else
            fra审查.Left = stbThis.Width - fra审查.Width - 1500
            fra审查.Top = stbThis.Top + 50
        End If
        fraPatiUD.Visible = picList.Visible
        mblnHScroll = (mdblScaleHeight > PicDraw.Height - IIf(picList.Visible, picList.Height, 0))
        With HScr
            .Value = 0
            .Top = PicDraw.Top
            .Left = PicDraw.Width - .Width
            .Height = PicDraw.Height
            .Visible = mblnHScroll
        End With
    Case conMenu_View_Append '显示房间号
        lbl床号(mlngSource).Tag = Val(lbl床号(mlngSource).Tag) Xor 1
        With mrsBedInfo
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                If ISShowCard Then
                    lbl床号(!卡片索引).Caption = IIf(Val(lbl床号(mlngSource).Tag) = 1, IIf(Trim(Nvl(!房间号)) = "", "", Trim(!房间号)) & IIf(IsNumeric(Trim(!房间号)), "_", ""), "") & Trim(!床号)
                    lbl房间号(!卡片索引).Caption = lbl床号(!卡片索引).Caption
                    lbl床号(!卡片索引).Caption = Mid(lbl床号(!卡片索引).Caption, 1, 7)
                End If
                .MoveNext
            Loop
        End With
    Case conMenu_View_NoticBoard
        If cboUnit.ListIndex = -1 Then Exit Sub
        '非模态显示公告栏窗体
        If mfrmNoticeBoard Is Nothing Then
            Set mfrmNoticeBoard = New frmNoticeBoard
        End If
        
        Call mfrmNoticeBoard.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex))
    Case conMenu_View_Notify '医嘱提醒
            With frmNotify
                .mintNotify = mintNotify
                .mintNotifyDay = mintNotifyDay
                .mstrNotifyAdvice = mstrNotifyAdvice
            End With
            frmNotify.mblnFirst = True
    Case conMenu_View_Refresh '刷新
        mblnRefresh = True
        '刷新医嘱提醒
        With frmNotify
            .mintNotify = mintNotify
            .mintNotifyDay = mintNotifyDay
            .mstrNotifyAdvice = mstrNotifyAdvice
            .mblnFirst = True
        End With
    Case conMenu_File_Parameter '参数设置
        frmSublimeStationSetup.mstrPrivs = mstrPrivs
        Call frmSublimeStationSetup.ShowMe
        If gblnOK Then
            Call GetLocalSetting
            mblnRefresh = True
            '刷新医嘱提醒
             With frmNotify
                .mintNotify = mintNotify
                .mintNotifyDay = mintNotifyDay
                .mstrNotifyAdvice = mstrNotifyAdvice
                .mblnFirst = True
            End With
        End If
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '退出
        Unload Me
    Case conMenu_File_PrintBedCard          '打印床头卡
        If Not LocatePatiRecord Then Exit Sub
        Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("科室ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
    Case conMenu_Manage_Print_Label '打印腕带
        If Not LocatePatiRecord Then Exit Sub
        If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me, "病人ID=" & Val(mrsPatiInfo.Fields("病人ID").Value), "主页ID=" & Val(mrsPatiInfo.Fields("主页ID").Value), 2)
        End If
    Case conMenu_File_PrintDayDetail        '一日清单
        If Not LocatePatiRecord Then Exit Sub
        Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("科室ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
    Case conMenu_File_PrintPageSet          '打印帐页设置
        If Not LocatePatiRecord Then Exit Sub
        Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("科室ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
    Case conMenu_File_MedRecSetup '首页打印设置
        Call PrintInMedRec(mclsInOutMedRec, 0, IIf(mlng病人ID = 0, -1, 0), mlng主页ID, mobjReport, Val(mrsPatiInfo.Fields("科室ID").Value), Me)
    Case conMenu_File_MedRecPreview '首页预览
        If Not LocatePatiRecord Then Exit Sub
        Call PrintInMedRec(mclsInOutMedRec, 1, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), mobjReport, Val(mrsPatiInfo.Fields("科室ID").Value), Me)
    Case conMenu_File_MedRecPrint '首页打印
        If Not LocatePatiRecord Then Exit Sub
        Call PrintInMedRec(mclsInOutMedRec, 2, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value), mobjReport, Val(mrsPatiInfo.Fields("科室ID").Value), Me)
    '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
    Case conMenu_Tool_MedRec '首页整理
        If Not LocatePatiRecord Then Exit Sub
        Call ExecuteEditMediRec
'    Case conMenu_View_FindNext '查找下一个
'        If txtFind.Text = "" Then
'            txtFind.SetFocus
'        Else
'            Call ExecuteFindPati(True)
'        End If
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 5 '查找方式
        mintFindType = Val(Right(Control.ID, 2)) - 1
        cbsMain.RecalcLayout
        txtFind.Text = ""
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
    Case conMenu_View_FindType * 100# + 9
        mintFindType = Val(Right(Control.ID, 2)) - 1
        cbsMain.RecalcLayout
        txtFind.Text = ""
        Call ExecuteFindPati
    Case conMenu_View_FindType * 100# + 11 To conMenu_View_FindType * 100# + 14 '查找方式
        mintPatiInputType = Val(Right(Control.ID, 2)) - 1
        cbsMain.RecalcLayout
        txt住院号.Text = ""
        If pic出院查找.Enabled And pic出院查找.Visible Then pic出院查找.SetFocus
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            strTmp = Split(Control.Parameter, ",")(1)
            If strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1132" Then '住院科室日报
                If Not LocatePatiRecord Then Exit Sub
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                         "病区=" & cboUnit.ItemData(cboUnit.ListIndex), "病人ID=" & Val(mrsPatiInfo.Fields("病人ID").Value), "主页ID=" & Val(mrsPatiInfo.Fields("主页ID").Value))
            ElseIf strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Or strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then    '病人帐页和催款表
                Call mclsFeeQuery.zlExecuteCommandBars(Control)
            Else
                If Not LocatePatiRecord Then Exit Sub
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                    "病人ID=" & Val(mrsPatiInfo.Fields("病人ID").Value), "主页ID=" & Val(mrsPatiInfo.Fields("主页ID").Value), "住院号=" & CStr(mrsPatiInfo.Fields("住院号").Value), "病人病区=" & cboUnit.ItemData(cboUnit.ListIndex), _
                    "病人科室=" & Val(mrsPatiInfo.Fields("科室ID").Value), "床号=" & Nvl(mrsPatiInfo.Fields("床号").Value))
            End If
        ElseIf Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 1, conMenu_File_MedRecPreview * 100# + 4) Then
            Call PrintInMedRec(mclsInOutMedRec, IIf(Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6), 2, 1), mlng病人ID, mlng主页ID, mobjReport, mPatiInfo.科室ID, Me, Val(Mid(Control.ID & "", Len(Control.ID & ""))))
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetSourceCardH()
'    If mblnCardCollapse Then
'        picPati(mlngSource).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
'    ElseIf mblnShowCard = True Then
'        picPati(mlngSource).Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
'    Else
'        picPati(mlngSource).Height = IIf(mlngSource = 0, clngBigHeight_Normal, clngBaseHeight_Normal)
'    End If
    If mblnCardCollapse Then
        picPati(mlngSource).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
    Else
        picPati(mlngSource).Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
    End If
    picList.ZOrder 0
    PatiPage.ZOrder 0
    fraPatiUD.ZOrder 0
    picPara(2).ZOrder 0
    picPara(3).ZOrder 0
    pic出院查找.ZOrder 0
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置病人相关的菜单和工具栏的可见状态
    Dim blnVisible As Boolean, strPrivs As String

    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Parameter = "已判断" Then Exit Sub

    blnVisible = True
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p病人入出)
    
    Select Case Control.ID
        Case conMenu_Manage_Change_In
            blnVisible = strPrivs <> ""
        Case conMenu_Manage_Change_Out
            blnVisible = InStr(strPrivs, "病人出院") > 0
        Case conMenu_Manage_Change_Turn
            blnVisible = InStr(strPrivs, "病人转科") > 0
        Case conMenu_Manage_Change_Bed, conMenu_Manage_Change_TransposeBed, conMenu_Manage_Change_House
            blnVisible = InStr(strPrivs, "换床") > 0
        Case conMenu_Manage_Change_TurnUnit
            blnVisible = InStr(strPrivs, "转病区") > 0
        Case conMenu_Manage_Change_PatiInfo
            blnVisible = InStr(strPrivs, "调整病人信息") > 0
        Case conMenu_Manage_Change_Baby
            blnVisible = InStr(strPrivs, "新生儿登记") > 0
        Case conMenu_Manage_Change_ReCalcFee
            blnVisible = InStr(strPrivs, "重算费用") > 0
        Case conMenu_Manage_Change_BedGrid
            blnVisible = InStr(strPrivs, "调整床位等级") > 0
        Case conMenu_Manage_Change_InPati
            blnVisible = InStr(strPrivs, "住院留观转住院") > 0
    End Select

    Control.Visible = blnVisible
    Control.Parameter = "已判断"
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean, blnSelect As Boolean, blnWaitIn As Boolean, blnWriteMedRec As Boolean
    Dim blnOut As Boolean, blnPreOut As Boolean, blnOutTo As Boolean, lngType As Long, strPrivs As String
    Dim strCustom As String
    
    If Not mblnStart Then Exit Sub
    If blnUnload Then Exit Sub

    blnSelect = LocatePatiRecord
    If blnSelect Then
        lngType = Val(mrsPatiInfo.Fields("排序").Value)
        blnWaitIn = lngType = pt转科待入住 Or lngType = pt入院待入住 Or lngType = pt转病区待入住
        blnOut = lngType = pt出院
        blnPreOut = lngType = pt预出
        '85200:控制最近转出页面的病人不允许进行相关操作，如：撤销操作
        blnOutTo = lngType = pt最近转出
    End If
    
    '首页报表
    If Between(Control.ID, conMenu_File_MedRecPrint * 100# + 3, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 3, conMenu_File_MedRecPreview * 100# + 4) Then
        If mintMecStandard = 0 Or mintMecStandard = 3 Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    End If
    
    If Control.Category = "撤销" Then
        Exit Sub    '在cbsMain_InitCommandsPopup已设置,退出避免子窗体设置其可见性
    ElseIf Control.Category = "病人" Then
        
        Call SetControlVisible(Control)
        If Not Control.Visible Then Exit Sub
        
        strPrivs = GetInsidePrivs(Enum_Inside_Program.p病人入出)
        If InStr(strPrivs, "所有病区") = 0 Then
            If InStr("," & mstrUnits & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Control.Enabled = False: Exit Sub
        End If
    End If
    
    '由子程序根据权限设置菜单功能的状态
    strCustom = ""
    If Not Control.Parent Is Nothing Then
        strCustom = Control.Parent.Title
    End If
    If strCustom <> "" Then
        If strCustom = "右键菜单" Then
            Select Case Control.ID
            Case conMenu_Edit_ReStop, conMenu_Manage_ReportLisView
                strCustom = "医嘱业务"
            Case conMenu_Edit_Billing, conMenu_Edit_ReBillingApply
                strCustom = "费用业务"
            End Select
        End If
        If strCustom = "医嘱业务" Then
            If Control.ID = conMenu_View_Notify Then
                Control.Enabled = True
            Else
                Call mclsAdvices.zlCheckPrivs(Control, 1)
                Control.Enabled = Control.Visible And blnSelect
                '50906:刘鹏飞,2012-09-18,入院待入住病人根据参数"允许给待入住病人下达医嘱"决定是否可以新开医嘱
                If Control.ID = conMenu_Edit_NewItem And Control.Enabled = True And lngType = pt入院待入住 Then
                    Control.Enabled = (Val(zlDatabase.GetPara("允许给待入住病人下达医嘱", glngSys, p住院医嘱下达, 1)) = 1)
                End If
            End If
            Exit Sub
        ElseIf strCustom = "费用业务" Then
            Call mclsFeeQuery.zlCheckPrivs(Control)
            Control.Enabled = Control.Visible And blnSelect
            
            If Control.ID = conMenu_Edit_PreBalance And Control.Enabled = True Then
                Control.Enabled = blnSelect And Nvl(mrsPatiInfo.Fields("险类").Value, 0) <> 0
            ElseIf Control.ID = conMenu_Manage_Change_ReCalcFee And Control.Enabled = True Then
                Control.Enabled = blnSelect And Nvl(mrsPatiInfo.Fields("险类").Value, 0) = 0
            End If
            Exit Sub
        ElseIf strCustom = "护理业务" Then
            Control.Enabled = (GetInsidePrivs(p护理记录管理, True) <> "") And blnSelect
        ElseIf strCustom = "病历业务" Then
            Control.Enabled = (GetInsidePrivs(p住院病历管理, True) <> "") And blnSelect
        ElseIf strCustom = "病区批量工作" Then
            '54409:刘鹏飞,2012-09-25,病区批量工作没有选择病人也可以使用(除病人事务处理外)
            Select Case Control.ID
            Case conMenu_Edit_PreBalance                '预结算
                Control.Enabled = True ' blnSelect
            '61762:刘鹏飞,2013-05-20,增加发送输液药品医嘱的功能
            Case conMenu_Edit_Audit, conMenu_Edit_Send, conMenu_Edit_SendInfusion, conMenu_Edit_Pause, conMenu_Edit_Reuse, conMenu_Edit_ReStop '医嘱校对、医嘱发送、发送输液药品医嘱、医嘱暂停、医嘱启用、医嘱确认停止
                Call mclsAdvices.zlCheckPrivs(Control, 1)
                 'Control.Enabled = Control.Visible And blnSelect
                If Not mrsPatiInfo Is Nothing Then
                    If mrsPatiInfo.State = adStateOpen Then
                        If blnSelect = False Then mrsPatiInfo.Filter = ""
                        Control.Enabled = Control.Visible And (mrsPatiInfo.RecordCount > 0)
                    End If
                End If
            Case conMenu_File_PrintMultiBill            '催款管理（新）
                Control.Visible = InStr(1, ";" & mstrPrivs & ";", ";病人催款表;")
                Control.Enabled = Control.Visible
            Case conMenu_Edit_BatExecute                   '执行登记（新）
                '60781:刘鹏飞,2013-07-15
                'Call mclsAdvices.zlCheckPrivs(Control, 1)
                Control.Visible = (InStr(GetInsidePrivs(p住院医嘱发送), ";批量执行登记;") > 0)
                Control.Enabled = Control.Visible
            Case conMenu_Edit_AnimalHeat                '批量录入体温单（新）
                Control.Visible = InStr(1, GetInsidePrivs(p护理记录管理, True), ";体温单作图;")
                Control.Enabled = Control.Visible
            Case conMenu_Edit_NurseLogFile              '批量录入记录单（新）
                Control.Visible = InStr(1, GetInsidePrivs(p护理记录管理, True), ";护理记录登记;")
                Control.Enabled = Control.Visible
            Case conMenu_Manage_ThingAudit, conMenu_Report_DrugQuery, conMenu_Edit_Surplus, conMenu_Report_Reports, conMenu_Edit_SendBack                '摆药查询,留存登记,打印执行单,超期收回
                Call mclsAdvices.zlCheckPrivs(Control, 1)
                Control.Enabled = Control.Visible
            Case conMenu_ProveCollect
                Control.Visible = mstrPrivs_检验采集 <> ""
                Control.Enabled = Control.Visible
            Case conMenu_病人事务处理                   '病人事务处理（新）
                Control.Enabled = blnSelect
            End Select
            Exit Sub
        End If
    End If
    
    Select Case Control.ID
    Case conMenu_Manage_Change_Undo
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOutTo
        If Control.Enabled = True Then
            Control.Enabled = Val(Nvl(mrsPatiInfo.Fields("主页ID").Value, 0)) = Val(Nvl(mrsPatiInfo.Fields("最大主页Id").Value, 0))
        End If
    Case conMenu_Manage_Change_In
        Control.Enabled = blnWaitIn
    Case conMenu_Manage_Change_InPati
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.性质 = 2
        End If
    '转科，换床，包房，调整病人信息，重算费用,转病区，转小组,床位对换
    Case conMenu_Manage_Change_Turn, conMenu_Manage_Change_Bed, conMenu_Manage_Change_House, _
         conMenu_Manage_Change_PatiInfo, conMenu_Manage_Change_ReCalcFee, conMenu_Manage_Change_TurnUnit, _
         conMenu_Manage_Change_TurnTeam, conMenu_Manage_Change_TransposeBed
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = mrsPatiInfo.Fields("状态").Value <> 2
            
            If Control.ID = conMenu_Manage_Change_TransposeBed Then '床位对换
                Control.Enabled = Trim(CStr(mrsPatiInfo.Fields("床号").Value)) <> ""
            ElseIf Control.ID = conMenu_Manage_Change_ReCalcFee Then
                Control.Enabled = Nvl(mrsPatiInfo.Fields("险类").Value, 0) = 0
            End If
        End If
    Case conMenu_Manage_Change_InsureSel
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = Nvl(mrsPatiInfo.Fields("险类").Value, 0) <> 0
        End If
    Case conMenu_Manage_Change_BedGrid
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = Trim(Nvl(mrsPatiInfo.Fields("床号").Value)) <> "" And mrsPatiInfo.Fields("状态").Value <> 2
        End If
    Case conMenu_Manage_Change_Out
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = (InStr(1, "," & pt在院 & ",3.1,", mrsPatiInfo.Fields("排序").Value) <> 0 Or blnPreOut) And mrsPatiInfo.Fields("状态").Value <> 2
        End If
    Case conMenu_Manage_Change_Baby
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.产科
        End If
    Case conMenu_Manage_Change_PaitNote
            Control.Enabled = Not blnOutTo
    Case conMenu_Manage_Monitor '监护仪
        Control.Visible = mblnMonitor And (InStr(GetInsidePrivs(p住院护士站), "护理监护") > 0)
        Control.Enabled = False
        If blnSelect Then
            mrsBedInfo.Filter = "床号='" & mrsPatiInfo!床号 & "'"
            If mrsBedInfo.RecordCount <> 0 Then
                Control.Enabled = Nvl(mrsBedInfo!监护仪, 0) > 0
            End If
            mrsBedInfo.Filter = ""
        End If
    Case conMenu_Tool_Archive '电子病案查阅
        Control.Visible = GetInsidePrivs(p电子病案查阅) <> ""
        Control.Enabled = Control.Visible And blnSelect
    Case conMenu_View_Warrant '担保信息查阅
        Control.Enabled = blnSelect
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        If GetInsidePrivs(p疾病诊断参考) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 '药品及诊疗参考
        If GetInsidePrivs(p药品诊疗参考) = "" Then Control.Visible = False
    Case conMenu_Tool_MedRecAuditResponse '审查反馈
        '都可以调用，至少可以查看(当前或历史)
        Control.Enabled = blnSelect
    Case conMenu_Manage_Print_Label '打印腕带
        Control.Visible = InStr(mstrPrivs, ";腕带打印;")
        If blnSelect = True Then
            Control.Enabled = mintREPORTSEL <> 页面.出院
        End If
        
    Case conMenu_File_MedRec '首页打印
        Control.Visible = InStr(mstrPrivs, "打印首页")
    '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
    Case conMenu_Tool_MedRec '首页整理
        blnWriteMedRec = Val(zlDatabase.GetPara("医生和护士分别填写病案首页", glngSys, p住院医生站, "0")) = 1
        Control.Enabled = blnSelect And blnWriteMedRec
        Control.Visible = blnWriteMedRec
    Case conMenu_File_Parameter '参数设置
        'If InStr(mstrPrivs, "参数设置") = 0 Then Control.Visible = False
'    Case conMenu_Tool_UnitSubject '病区标记设置
'        Control.Visible = InStr(1, ";" & mstrPrivs & ";", ";病区标记设置;")
'        Control.Enabled = Control.Visible
    Case conMenu_Tool_UnitNBoard
        Control.Visible = InStr(1, ";" & mstrPrivs & ";", ";病区公告栏设置;")
        Control.Enabled = Control.Visible
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
        Call cbsMain_Resize
    Case conMenu_View_FontSize_S      '标准卡片 小字体
        Control.Checked = (mlngSource = 999)
    Case conMenu_View_FontSize_L      '大卡片 大字体
        Control.Checked = (mlngSource = 0)
    Case conMenu_View_Expend_AllCollapse    '卡片折叠
        Control.Checked = mblnCardCollapse
    Case conMenu_View_Expend_CurCollapse
        Control.Checked = picList.Visible
    Case conMenu_View_Append
        Control.Checked = (Val(lbl床号(mlngSource).Tag) = 1)
    Case conMenu_View_FindType '查找方式
        Control.Enabled = True
        Control.Caption = "↓按" & Decode(mintFindType, 0, "床号", 1, "住院号", 2, "就诊卡", 3, "姓名", 4, "简码", 8, "床号") & "查找"
        txtFind.PasswordChar = IIf(mintFindType = 2 And gblnCardHide, "*", "")
        
        '出院病人查找方式
        lblPatiInputType.Caption = Decode(mintPatiInputType, 10, "床 号", 11, "住院号", 12, "就诊卡", 13, "姓 名", "姓 名") & "↓"
        txt住院号.PasswordChar = IIf(mintPatiInputType = 2 And gblnCardHide, "*", "")
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '查找方式
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
    Case conMenu_View_FindType * 100# + 11 To conMenu_View_FindType * 100# + 14 '出院病人查找方式
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintPatiInputType
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '外挂功能执行
        Control.Enabled = True 'blnSelect
    End Select
    
End Sub

Private Sub GetLocalSetting()
'功能：从注册表读取出院病人的时间范围
    Dim curDate As Date, intDay As Integer

    '病人显示范围
    mstrScope = "11111"
    mintChange = Val(zlDatabase.GetPara("最近转出天数", glngSys, p住院护士站, 7))
    '转出病人天数
    txtChange.Text = mintChange
    
    '出院病人时间范围
'    curDate = zlDatabase.Currentdate
'    intDay = Val(zlDatabase.GetPara("出院病人结束间隔", glngSys, p住院护士站, 0))
'    mdtOutEnd = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
'    intDay = Val(zlDatabase.GetPara("出院病人开始间隔", glngSys, p住院护士站, 0))
'    mdtOutBegin = Format(curDate - intDay, "yyyy-MM-dd 00:00:00")
    
    '医嘱提醒刷新设置
    mstrNotifyAdvice = zlDatabase.GetPara("自动刷新医嘱类型", glngSys, p住院护士站, "0000000")
    mintNotifyDay = Val(zlDatabase.GetPara("自动刷新医嘱天数", glngSys, p住院护士站, 1))
    mintNotify = Val(zlDatabase.GetPara("自动刷新医嘱间隔", glngSys, p住院护士站))
    
    '卡片显示内容(诊断,余额)
    mstrCardInfo = zlDatabase.GetPara("卡片显示内容", glngSys, p住院护士站, "11")
    
    '病案审查反馈天数
    mlngMedRedDay = Val(zlDatabase.GetPara("病案审查反馈天数", glngSys, p住院护士站))
    
    '病案首页标准
    mintMecStandard = Val(zlDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0"))
    
    mblnCardBalance = (Val(zlDatabase.GetPara("卡片余额含担保金额", glngSys, 1265, 0)) = 1)
    '92852:刘鹏飞,2016-01-20,床位卡片的排序方式,0-床号排序,1-床位编制编号+床号排序
    mblnCardOrder = (Val(zlDatabase.GetPara("床位卡片排序方式", glngSys, 1265, 0)) = 0)
    '54370:刘鹏飞,2013-05-02,添加参数"医嘱校对后自动定位到医嘱页面"
    mblnCollateAutoFind = (Val(zlDatabase.GetPara("医嘱处理后自动定位到医嘱页面", glngSys, 1265, 0)) = 1)
    '设置页面控件的状态
    PatiPage.Item(页面.待入科).Visible = True
    PatiPage.Item(页面.转科).Visible = True
    PatiPage.Item(页面.出院).Visible = True
    
    '获取最小有效的页面序号
    If PatiPage.Item(页面.待入科).Visible Then
        mintPage = 页面.待入科
    ElseIf PatiPage.Item(页面.转科).Visible Then
        mintPage = 页面.转科
    ElseIf PatiPage.Item(页面.出院).Visible Then
        mintPage = 页面.出院
    Else
        mintPage = 页面.家庭病床
    End If
    Call InitColor
End Sub

Private Sub RefreshData()
    Dim rsPati As New ADODB.Recordset
    
    '输入匹配时，当页面下拉框清空，F5刷新，应该恢复上一个的值
    If cboUnit.ListIndex = -1 Then Call zlControl.CboSetIndex(cboUnit.hwnd, mintPreDept)
    mblnHavePath = HavePath(cboUnit.ItemData(cboUnit.ListIndex))
    Call init非在床清单
    mstrBoardKeys = ""
    mblnShow = False        '避免激活选择事件，导致卡片在最上面显示
    mintREPORTSEL = -1
    mlng病人ID = 0:    mlng主页ID = 0: mlngPre病人ID = 0: mlngPre主页ID = 0
    mlng空床 = 0: mlng在床 = 0: mlng入院 = 0: mlng转入 = 0: mlng出院 = 0: mlng预出院 = 0
    mlng转出 = 0: mlng死亡 = 0: mlng手术 = 0: mlng危 = 0: mlng重 = 0: mlng家床 = 0
    
    '1初始化内存记录集
    '61824:刘鹏飞,2013-05-23,显示单病种标志
    Set mrsBedInfo = New ADODB.Recordset
    mstrFields = "卡片索引," & adDouble & ",18|床号," & adLongVarChar & ",10|住院号," & adDouble & ",18|病人ID," & adDouble & ",18|" & _
                 "主页ID," & adDouble & ",18|病况," & adLongVarChar & ",10|监护仪," & adDouble & ",18|病案审查," & adDouble & ",18|" & _
                 "临床路径," & adDouble & ",18|个性标注1," & adLongVarChar & ",100|病人状态," & adDouble & ",18|个性标注2," & adLongVarChar & ",100|个性标注3," & adLongVarChar & ",100|" & _
                 "监护仪名称," & adLongVarChar & ",20|病案审查名称," & adLongVarChar & ",20|临床路径名称," & adLongVarChar & ",20|" & _
                 "个性标注1名称," & adLongVarChar & ",20|病人状态名称," & adLongVarChar & ",20|个性标注2名称," & adLongVarChar & ",20|个性标注3名称," & adLongVarChar & ",20|" & _
                 "护理等级," & adDouble & ",18|护理等级名称," & adLongVarChar & ",20|病人类型," & adLongVarChar & ",20|" & _
                 "包床," & adDouble & ",2|姓名," & adLongVarChar & ",100|简码," & adLongVarChar & ",200|床位编制," & adLongVarChar & ",50|房间号," & adLongVarChar & ",20|单病种," & adLongVarChar & ",10|住院天数," & adLongVarChar & ",10"
    Call Record_Init(mrsBedInfo, mstrFields)
    
    '提取病区标记内容
    Call LoadNotes
    
    '2装载本病区的所有床位
    Call ShowGuage("装载本病区的所有床位", 10)
    'debug.print "装载本病区的所有床位,Start:" & Now
    If Not LoadBeds And Not mblnStart Then
        Unload Me
        Exit Sub
    End If
    
    '3提取本病区所有病人清单
    Call ShowGuage("提取本病区所有病人清单", 20)
    'debug.print "提取本病区所有病人清单,Start:" & Now
    Call LoadPatients(rsPati)
    
    '4更新在床病人数据
    Call ShowGuage("更新在床病人数据", 30)
    'debug.print "更新在床病人数据,Start:" & Now
    Call UpgradeBeds(rsPati)
    
    '5装载不在床病人(家庭病床，如果勾选了待入科则加载待入科病人；已出院与最近转出的页面点击才加载)
    Call ShowGuage("装载不在床病人清单", 90)
    'debug.print "装载不在床病人,Start:" & Now
    
    Dim strField As String, strValue As String
    strField = "排序," & adDouble & ",2|排序2," & adDouble & ",2|类型," & adLongVarChar & ",50|病人ID," & adDouble & ",18|主页ID," & adDouble & ",18|" & _
               "住院号," & adDouble & ",18|姓名," & adLongVarChar & ",20|简码," & adLongVarChar & ",200|性别," & adLongVarChar & ",10|年龄," & adLongVarChar & ",20|科室," & adLongVarChar & ",50|" & _
               "科室ID," & adDouble & ",18|住院医师," & adLongVarChar & ",20|责任护士," & adLongVarChar & ",20|病案状态," & adLongVarChar & ",20|" & _
               "床号," & adLongVarChar & ",20|护理等级," & adLongVarChar & ",50|费别," & adLongVarChar & ",50|医疗付款方式," & adLongVarChar & ",50|当前病况," & adLongVarChar & ",50|" & _
               "入院日期," & adLongVarChar & ",20|出院日期," & adLongVarChar & ",20|住院天数," & adLongVarChar & ",20|出院方式," & adLongVarChar & ",20|" & _
               "病人类型," & adLongVarChar & ",50|状态," & adLongVarChar & ",10|险类," & adDouble & ",18|就诊卡号," & adLongVarChar & ",20|路径状态," & adLongVarChar & ",20|" & _
               "颜色," & adDouble & ",18|单病种," & adLongVarChar & ",10|婴儿科室ID," & adDouble & ",18|婴儿病区ID," & adDouble & ",18|最大主页Id," & adDouble & ",18"
    Call Record_Init(mrsPatiInfo, strField)
    
    Call UpgradeList(rsPati)
    '激活当前非在床页面的点击事件
    If PatiPage.Selected Is Nothing Then
        PatiPage.Item(mintPage).Selected = True
    Else
        If PatiPage.Selected.Visible = False Then
            PatiPage.Item(mintPage).Selected = True
        End If
    End If
    Call PatiPage_SelectedChanged(PatiPage.Selected)
    '更改页面内容
    If GetPatiCount(页面.待入科) <> 0 Then PatiPage.Item(页面.待入科).Caption = "待入科" & GetPatiCount(页面.待入科) & "人"
    If GetPatiCount(页面.转科) <> 0 Then PatiPage.Item(页面.转科).Caption = "最近转科" & GetPatiCount(页面.转科) & "人"
    If GetPatiCount(页面.出院) <> 0 Then PatiPage.Item(页面.出院).Caption = "最近出院" & GetPatiCount(页面.出院) & "人"
    If GetPatiCount(页面.家庭病床) <> 0 Then PatiPage.Item(页面.家庭病床).Caption = "家庭病床" & GetPatiCount(页面.家庭病床) & "人"

    Call ShowGuage("数据读取结束", 100)
    'debug.print "结束,OVER:" & Now
    Call GetInpatientAreaInfo
    
    '6再根据设定的条件显示或隐藏相应的卡片
    Call ShowSelect                 '人为的调一下，避免卡片没有人为点击却显示在最上面
    Call AdjustCard
    
    Call CopyReocrd(rsPati)
    
    Call AddSendCommandBar
End Sub

Private Sub LoadNotes()
    Dim strPatientFilter As String
    Dim blnNext As Boolean, strItems As String
    Dim i As Integer, strKey As String
    On Error GoTo ErrHand
    
    With Me.cbo主题
        .Clear
        .AddItem "所有"
        .AddItem "病案审查"
        .AddItem "临床路径"
        .AddItem "病人状态"
        '提取当前病区设定的标注主题
        mstrSQL = "Select nvl(病区ID,0) 病区ID,主题序号, 标记序号, Replace(说明, '|', '') 说明, 图形索引, 有效天数" & vbNewLine & _
            " From 病区标记内容" & vbNewLine & _
            " Where 病区id Is Null Or 病区id = [1]" & vbNewLine & _
            " Order By Nvl(病区id, 0), 主题序号, 标记序号"
        Set mrsNotes = zlDatabase.OpenSQLRecord(mstrSQL, "提取病区标记内容", Me.cboUnit.ItemData(Me.cboUnit.ListIndex))
        strItems = "": strKey = ""
        Do While Not mrsNotes.EOF
            If Val("" & mrsNotes!标记序号) = 0 Then
                blnNext = True
                strKey = mrsNotes!病区ID & "-" & mrsNotes!主题序号
                .AddItem mrsNotes!说明 & ""
                .ItemData(.NewIndex) = Val(mrsNotes!病区ID) + Val(mrsNotes!主题序号)
                strItems = strItems & "|"
            Else
                If strKey = mrsNotes!病区ID & "-" & mrsNotes!主题序号 Then
                    strItems = strItems & IIf(blnNext, "", ",") & mrsNotes!说明 & "'" & mrsNotes!标记序号
                    blnNext = False
                End If
            End If
            mrsNotes.MoveNext
        Loop
        If mrsNotes.RecordCount <> 0 Then mrsNotes.MoveFirst
        If strItems <> "" Then strItems = Mid(strItems, 2)
        mstrNoteItems = strItems
        
        strPatientFilter = zlDatabase.GetPara("入科天数", glngSys, 1265, "3")
        .Tag = "等待审查,拒绝审查,正在抽查,正在审查,抽查反馈,审查反馈,抽查整改,审查整改|未导入,执行中,不符合,正常结束,变异结束|预转科,预出院" & IIf(Val(strPatientFilter) = 0, "", ",入科" & strPatientFilter & "天内") & "|" & strItems
        .ListIndex = 0
    End With
    
    '提取当前病区的标注记录
    'LPF,2014-10-21,性能优化:添加在院病人表
    mstrSQL = "" & _
            " Select a.病人id, a.主页id,nvl(a.主题病区ID,0) 主题病区ID, a.主题序号, a.标记序号,a.标记顺序, a.日期, Replace(b.说明, '|', '') 说明, b.图形索引, b.有效天数, Floor(Sysdate - a.日期) As 实际天数" & vbNewLine & _
            " From 病区标记记录 a, 病区标记内容 b, 病人信息 c, 在院病人 e" & vbNewLine & _
            " Where a.主题序号 = b.主题序号 And a.标记序号 = b.标记序号 And nvl(a.主题病区ID,0) = nvl(b.病区id,0) And a.病人id = c.病人id And a.主页id = c.主页id And " & vbNewLine & _
            "      a.病区id = c.当前病区id And c.病人id = e.病人id And c.当前病区id = e.病区id And e.病区id = [1] And " & vbNewLine & _
            "      (b.有效天数 = 0 Or (b.有效天数 > Floor(Sysdate - a.日期)))" & vbNewLine & _
            " Order By a.病人id, a.主页id,a.标记顺序,a.主题序号"

    Set mrsPatiNotes = zlDatabase.OpenSQLRecord(mstrSQL, "提取指定病区的有效标注记录", Me.cboUnit.ItemData(Me.cboUnit.ListIndex))
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CopyReocrd(ByVal rsPati As ADODB.Recordset)
    Dim strField As String, strValue As String
    '61824:刘鹏飞,2013-05-23,显示单病种标志
    rsPati.Filter = 0
    If rsPati.RecordCount <> 0 Then rsPati.MoveFirst
    strField = "排序|排序2|类型|病人ID|主页ID|住院号|姓名|简码|性别|年龄|科室|科室ID|住院医师|责任护士|病案状态|床号|护理等级|费别|医疗付款方式|当前病况|入院日期|出院日期|住院天数|出院方式|病人类型|状态|险类|就诊卡号|路径状态|颜色|单病种|婴儿科室ID|婴儿病区ID|最大主页Id"
    Do While Not rsPati.EOF
        strValue = rsPati!排序 & "|" & rsPati!排序2 & "|" & rsPati!类型 & "|" & rsPati!病人ID & "|" & rsPati!主页ID & "|" & Nvl(rsPati!住院号, 0) & "|" & rsPati!姓名 & "|" & Nvl(rsPati!简码) & "|" & rsPati!性别 & "|" & _
                  rsPati!年龄 & "|" & Nvl(rsPati!科室) & "|" & Nvl(rsPati!科室ID, 0) & "|" & Nvl(rsPati!住院医师) & "|" & Nvl(rsPati!责任护士) & "|" & Nvl(rsPati!病案状态, 0) & "|" & Nvl(rsPati!床号) & "|" & _
                  Nvl(rsPati!护理等级, "三级") & "|" & Nvl(rsPati!费别) & "|" & Nvl(rsPati!医疗付款方式) & "|" & Nvl(rsPati!当前病况, "一般") & "|" & Format(rsPati!入院日期, "yyyy-MM-dd") & "|" & Format(rsPati!出院日期, "yyyy-MM-dd") & "|" & rsPati!住院天数 & "|" & rsPati!出院方式 & "|" & _
                  Nvl(rsPati!病人类型, "普通病人") & "|" & rsPati!状态 & "|" & Nvl(rsPati!险类, 0) & "|" & Nvl(rsPati!就诊卡号) & "|" & Nvl(rsPati!路径状态, 0) & "|" & Nvl(rsPati!颜色, 0) & "|" & Nvl(rsPati!单病种) & "|" & Nvl(rsPati!婴儿科室ID, 0) & "|" & Nvl(rsPati!婴儿病区ID, 0) & "|" & Nvl(rsPati!最大主页ID, 0)
        Call Record_Add(mrsPatiInfo, strField, strValue)
        rsPati.MoveNext
    Loop
End Sub

Private Sub chk包含空床_Click()
    If Not mblnStart Then Exit Sub
    mintREPORTSEL = -1
    If mblnHScroll Then
        If HScr.Value <> 0 Then
            mstrBoardKeys = ""
            HScr.Value = 0
        Else
            Call AdjustCard
        End If
    Else
        Call AdjustCard
    End If
End Sub

Private Sub HScr_Change()
    Dim lngMove As Long
    Dim lngY As Long
    
    '计算单步步长
    lngMove = CLng((mdblScaleHeight - (PicDraw.Height - IIf(picList.Visible, picList.Height, 0))) / 100)
    If lngMove < 0 Then lngMove = 0
    lngY = -1 * HScr.Value * lngMove
    If lngY >= 0 And lngY < 100 Then lngY = 100
    Call AdjustCard(lngY, mstrBoardKeys)
End Sub

Private Sub lbl姓名_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, lbl姓名(Index).Caption, True
End Sub

Private Sub lbl床号_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, lbl房间号(Index).Caption, True
End Sub

Private Sub lbl医师_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, lbl医师(Index).Caption, True
End Sub

Private Sub lbl诊断_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, "诊断：" & lbl诊断(Index).Caption, True
End Sub

Private Sub picPatiIn_Resize()
    Dim i As Long, y As Long, dblTop As Double
    On Error Resume Next
    If picList.Visible = False Then
        picPara(2).Visible = False
        picPara(3).Visible = False
        pic出院查找.Visible = False
        Exit Sub
    Else
        For i = 2 To 3
            picPara(i).Visible = False
        Next
        If PatiPage.Selected.Index = 页面.待入科 Then
            pic出院查找.Tag = 页面.待入科
        ElseIf PatiPage.Selected.Index = 页面.出院 Then
            picPara(2).Visible = True
            pic出院查找.Tag = 页面.出院
        ElseIf PatiPage.Selected.Index = 页面.转科 Then
            picPara(3).Visible = True
            pic出院查找.Tag = 页面.转科
        ElseIf PatiPage.Selected.Index = 页面.家庭病床 Then
            pic出院查找.Tag = 页面.家庭病床
        End If
    End If
    
    picPara(2).Top = 20
    For i = 3 To 3
        picPara(i).Top = picPara(i - 1).Top + IIf(picPara(i - 1).Visible, picPara(i - 1).Height, 0)
        If picPara(i - 1).Visible Then y = y + 1
    Next
    
    If picPara(i - 1).Visible Then y = y + 1
    rptPati(PatiPage.Selected.Index).Top = IIf(y > 0, 20 + TextWidth("刘") - 180, 350) + y * (310 + TextWidth("刘") - 180)
    rptPati(PatiPage.Selected.Index).Left = 0
    rptPati(PatiPage.Selected.Index).Width = picList.Width
    rptPati(PatiPage.Selected.Index).Height = picList.Height - rptPati(PatiPage.Selected.Index).Top - IIf(y > 0, 350, 0)
    
    picPara(2).ZOrder 0
    picPara(3).ZOrder 0
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    Dim strValue As String
    Dim lngDept As Long, lngUnit As Long, lngCurrUnit As Long, lngCurrDept As Long
    Dim lngPatID As Long, lngPageID As Long, strName As String, strBed As String, strOutWay As String
    Dim strSQL As String, rsTmp As New ADODB.Recordset, rsBed As New ADODB.Recordset
    Dim blnFresh As Boolean
    Dim intCardIndex As Integer, i As Long
    Dim strKey As String
    Dim arrCardIndex As Variant
    
    On Error GoTo ErrHand
    
    Select Case UCase(strMsgItemIdentity)
        Case "ZLHIS_PATIENT_001" '入院待入科
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            '提取病人ID、主页ID、姓名
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            '检查病区
            strValue = "": Call mclsXML.GetSingleNodeValue("in_dept_id", strValue, xsNumber): lngDept = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("in_area_id", strValue, xsNumber)
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Or lngDept = 0 Then Exit Sub
            
            If Val(strValue) = 0 Then
                strValue = ""
                strSQL = "Select 病区ID From 病区科室对应 where 科室ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病区信息", lngDept)
                Do While Not rsTmp.EOF
                    strValue = strValue & "," & rsTmp!病区ID
                rsTmp.MoveNext
                Loop
                strValue = Mid(strValue, 2)
            End If
            If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
            If FreshPatiCard("新增待入科病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新代办入住的病人:" & strName, "待办入住提醒")
            End If
        Case "ZLHIS_PATIENT_002" '入住
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '提取病人ID、主页ID、姓名...
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("in_bed", strValue, xsString): strBed = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("in_area_id", strValue, xsNumber): lngUnit = Val(strValue)
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            If InStr(1, "," & lngUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
            '检查病人是否正常在院
            strSQL = "Select 病人ID From 病案主页 Where 病人ID=[1] And 主页ID=[2] And NVL(状态,0)=0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否正常在院", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
            If Not mrsPatiInfo.EOF Then
                '检查病人存在入院待入住列表中
                If mrsPatiInfo!排序 = 0 Then
                    mrsPatiInfo.Delete: mrsPatiInfo.Filter = ""
                    strKey = ""
                    If mintREPORTSEL = 页面.待入科 Then
                        If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                            If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                            End If
                        End If
                    End If
                    rptPati(页面.待入科).Records.DeleteAll
                    Call UpgradeList(mrsPatiInfo, 页面.待入科)
                    PatiPage.Item(页面.待入科).Caption = "待入科" & GetPatiCount(页面.待入科) & "人"
                    If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                Else
                    Exit Sub
                End If
            End If
            If FreshPatiCard("新增在院病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新入住的病人:" & strName & "   床号:" & strBed, "入住提醒")
            End If
            
        Case "ZLHIS_PATIENT_003" '转出
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '提取病人ID、主页ID、姓名
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("current_bed", strValue, xsString): strBed = strValue
            
            '1、转入科室待入科列表刷新
            strValue = "": Call mclsXML.GetSingleNodeValue("current_area_id", strValue, xsNumber): lngCurrUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("change_dept_id", strValue, xsNumber): lngDept = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("change_area_id", strValue, xsNumber): lngUnit = Val(strValue)
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            If Not (lngUnit = 0 And lngDept = 0) Then
                If lngUnit = 0 Then
                    strValue = ""
                    strSQL = "Select 病区ID From 病区科室对应 where 科室ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病区信息", lngDept)
                    Do While Not rsTmp.EOF
                        strValue = strValue & "," & rsTmp!病区ID
                    rsTmp.MoveNext
                    Loop
                    strValue = Mid(strValue, 2)
                Else
                    strValue = lngUnit
                End If
                If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    If FreshPatiCard("新增转出待入科病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
                        If strName <> "" Then
                            Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新转入的病人:" & strName, "待办入住提醒")
                        End If
                    End If
                End If
            End If
            '2、转出科室在院病人列表刷新
            If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
            '处理在床病人图标
            strSQL = "Select 病人ID From 病案主页 Where 病人ID=[1] And 主页ID=[2] And NVL(状态,0)=2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否处于预转科状态", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            
            mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
            blnFresh = False
            Do While Not mrsPatiInfo.EOF
                If InStr(1, ",3,3.1,4,", "," & mrsPatiInfo!排序 & ",") <> 0 Then
                    blnFresh = True
                    If mrsPatiInfo!排序 = 3.1 Then
                        mrsPatiInfo!状态 = 2
                    Else
                        mrsPatiInfo!排序 = 3.2: mrsPatiInfo!类型 = "预转科病人": mrsPatiInfo!状态 = 2
                    End If
                    mrsPatiInfo.Update
                End If
            mrsPatiInfo.MoveNext
            Loop
            If blnFresh = False Then Exit Sub
            
            mrsBedInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID & " And 包床<>1"
            If Not mrsBedInfo.EOF Then
                intCardIndex = mrsBedInfo!卡片索引
                mrsBedInfo!病人状态 = Img标记(mlngSource).ListImages("预转科").Index
                mrsBedInfo!病人状态名称 = "预转科"
                mrsBedInfo.Update
                Call SetCardLabel(intCardIndex)
            End If
            mrsBedInfo.Filter = 0
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新待转出的病人:" & strName & "   床号:" & strBed, "待转出提醒")
            End If
        Case "ZLHIS_PATIENT_009" '预出院
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '提取病人ID、主页ID、姓名
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("out_area_id", strValue, xsNumber): lngCurrUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("out_bed", strValue, xsNumber): strBed = strValue
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
            '处理在床病人图标
            strSQL = "Select 病人ID From 病案主页 Where 病人ID=[1] And 主页ID=[2] And NVL(状态,0)=3"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否处于预出院状态", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            
            mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
            blnFresh = False
            Do While Not mrsPatiInfo.EOF
                If InStr(1, ",3,3.1,3.2,", "," & mrsPatiInfo!排序 & ",") <> 0 Then
                    blnFresh = True
                    mrsPatiInfo!排序 = 4: mrsPatiInfo!类型 = "预出院病人": mrsPatiInfo!状态 = 3
                    mrsPatiInfo.Update
                End If
            mrsPatiInfo.MoveNext
            Loop
            If blnFresh = False Then Exit Sub
            
            mrsBedInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID & " And 包床<>1"
            If Not mrsBedInfo.EOF Then
                intCardIndex = mrsBedInfo!卡片索引
                mrsBedInfo!病人状态 = Img标记(mlngSource).ListImages("预出院").Index
                mrsBedInfo!病人状态名称 = "预出院"
                mrsBedInfo.Update
                Call SetCardLabel(intCardIndex)
            End If
            mrsPatiInfo.Filter = "类型='预出院病人'"
            mlng预出院 = mrsPatiInfo.RecordCount
            mrsPatiInfo.Filter = 0
            mrsBedInfo.Filter = 0
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新预出院的病人:" & strName & "   床号:" & strBed, "预出院提醒")
            End If
            
        Case "ZLHIS_PATIENT_010" '出院
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '提取病人ID、主页ID、姓名
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("out_area_id", strValue, xsNumber): lngCurrUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("out_bed", strValue, xsNumber): strBed = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("out_way", strValue, xsNumber): strOutWay = strValue
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            If lngCurrUnit <> cboUnit.ItemData(cboUnit.ListIndex) Then Exit Sub
            
            strSQL = "Select 病人ID From 病案主页 Where 病人ID=[1] And 主页ID=[2] And 出院日期 IS NOT NULL"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否处于出院状态", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            
            '处理病人
            If FreshPatiCard("删除在院病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
            strKey = ""
            If mintREPORTSEL = 页面.出院 Then
                If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                    If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                        strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                    End If
                End If
            End If

            rptPati(页面.出院).Tag = "": rptPati(页面.出院).Records.DeleteAll
            If rptPati(页面.出院).Columns.Count > c_审查 Then rptPati(页面.出院).Columns(c_审查).Visible = False
            If PatiPage.Selected.Index = 页面.出院 Then Call PatiPage_SelectedChanged(PatiPage.Selected)
            If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
            
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新出院的病人:" & strName & "   床号:" & strBed & "   出院方式:" & strOutWay, "预出院提醒")
            End If
                
        Case "ZLHIS_PATIENT_012" '转入科室
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '提取病人ID、主页ID、姓名...
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("in_bed", strValue, xsString): strBed = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("out_area_id", strValue, xsNumber): lngCurrUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("in_area_id", strValue, xsNumber): lngUnit = Val(strValue)
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            '检查病人是否正常在院
            strSQL = "Select 病人ID From 病案主页 Where 病人ID=[1] And 主页ID=[2] And NVL(状态,0)=0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否正常在院", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            'a)转出病区病人清单的刷新(一定要放在转入病区之前,转科可能存在入住病区和转出病区相同的情况)
            If lngCurrUnit = cboUnit.ItemData(cboUnit.ListIndex) Then
                '处理病人
                If FreshPatiCard("删除在院病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
                    rptPati(页面.转科).Tag = "": rptPati(页面.转科).Records.DeleteAll
                    If rptPati(页面.转科).Columns.Count > c_审查 Then rptPati(页面.转科).Columns(c_审查).Visible = False
                    If PatiPage.Selected.Index = 页面.转科 Then Call PatiPage_SelectedChanged(PatiPage.Selected)
                    
                    If strName <> "" Then
                        Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新已转出的病人:" & strName & "   床号:" & strBed, "已转出提醒")
                    End If
                End If
            End If
            'b)转入病区病人清单的刷新
            If lngUnit = cboUnit.ItemData(cboUnit.ListIndex) Then
                mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID & " And 排序<>7"
                If Not mrsPatiInfo.EOF Then
                    '检查病人存在入院待入住列表中
                    If mrsPatiInfo!排序 = 1 Or mrsPatiInfo!排序 = 2 Then
                        mrsPatiInfo.Delete: mrsPatiInfo.Filter = ""
                        strKey = ""
                        If mintREPORTSEL = 页面.待入科 Then
                            If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                                If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                    strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                                End If
                            End If
                        End If

                        rptPati(页面.待入科).Records.DeleteAll
                        If rptPati(页面.待入科).Columns.Count > c_审查 Then rptPati(页面.待入科).Columns(c_审查).Visible = False
                        Call UpgradeList(mrsPatiInfo, 页面.待入科)
                        PatiPage.Item(页面.待入科).Caption = "待入科" & GetPatiCount(页面.待入科) & "人"
                        If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                    Else
                        Exit Sub
                    End If
                End If
                If FreshPatiCard("新增在院病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                If strName <> "" Then
                    Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新转入已入住的病人:" & strName & "   床号:" & strBed, "入住提醒")
                End If
            End If
        Case "ZLHIS_PATIENT_006" '撤销变动
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("cancel_kind", strValue, xsString): strOutWay = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("before_area_id", strValue, xsNumber): lngCurrUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("before_dept_id", strValue, xsNumber): lngCurrDept = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("after_area_id", strValue, xsNumber): lngUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("after_dept_id", strValue, xsNumber): lngDept = Val(strValue)
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            
            Select Case strOutWay
            Case "出院"
                If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
                
                strSQL = "Select 出院病床 From 病案主页 Where 病人ID=[1] And 主页ID=[2] And　出院日期 IS NULL"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否正常在院", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!出院病床)
                mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
                If Not mrsPatiInfo.EOF Then
                    '检查病人存在出院列表中
                    If mrsPatiInfo!排序 = 5 Or mrsPatiInfo!排序 = 6 Then
                        strKey = ""
                        If mintREPORTSEL = 页面.出院 Then
                            If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                                If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                    strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                                End If
                            End If
                        End If
                        rptPati(页面.出院).Tag = "": rptPati(页面.出院).Records.DeleteAll
                        If rptPati(页面.出院).Columns.Count > c_审查 Then rptPati(页面.出院).Columns(c_审查).Visible = False
                        If PatiPage.Selected.Index = 页面.出院 Then Call PatiPage_SelectedChanged(PatiPage.Selected)
                        If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                    Else
                        Exit Sub
                    End If
                End If
                If FreshPatiCard("新增在院病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                If strName <> "" Then
                    Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新撤销出院的病人:" & strName & "   床号:" & strBed, "撤销出院提醒")
                End If
            Case "预出院"
                If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
                '处理在床病人图标
                strSQL = "Select 出院病床 From 病案主页 Where 病人ID=[1] And 主页ID=[2] And NVL(状态,0)=0"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否正常在院", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!出院病床)
                mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
                blnFresh = False
                Do While Not mrsPatiInfo.EOF
                    If InStr(1, ",4,3.2,", "," & mrsPatiInfo!排序 & ",") <> 0 Then
                        blnFresh = True
                        If strBed = "" Then
                            mrsPatiInfo!排序 = 3.1: mrsPatiInfo!类型 = "家庭病床": mrsPatiInfo!状态 = 0
                        Else
                            mrsPatiInfo!排序 = 3: mrsPatiInfo!类型 = "在院病人": mrsPatiInfo!状态 = 0
                        End If
                        mrsPatiInfo.Update
                    End If
                mrsPatiInfo.MoveNext
                Loop
                If blnFresh = False Then Exit Sub
            
                mrsBedInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID & " And 包床<>1"
                If Not mrsBedInfo.EOF Then
                    intCardIndex = mrsBedInfo!卡片索引
                    mrsBedInfo!病人状态 = 0
                    mrsBedInfo!病人状态名称 = ""
                    mrsBedInfo.Update
                    Call SetCardLabel(intCardIndex)
                End If
                mrsPatiInfo.Filter = "类型='预出院病人'"
                mlng预出院 = mrsPatiInfo.RecordCount
                mrsPatiInfo.Filter = 0
            
                mrsBedInfo.Filter = 0
                If strName <> "" Then
                    Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新撤销预出院的病人:" & strName & "   床号:" & strBed, "撤销预出院提醒")
                End If
            Case "转病区入住", "转科入住"
                '病人状态和刷新病区检查
                strSQL = "Select 出院病床,当前病区ID From 病案主页 Where 病人ID=[1] And 主页ID=[2] And NVL(状态,0)=2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否正常在院", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!出院病床)
                'a)  入住病区在院
                If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    If FreshPatiCard("删除在院病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                    If strName <> "" Then
                        Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新撤销入住的病人:" & strName, "撤销入住提醒")
                    End If
                End If
                
                'b)  转出病区在院列表/转出列表刷新
                If InStr(1, "," & Nvl(rsTmp!当前病区ID, 0) & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
                    If Not mrsPatiInfo.EOF Then
                        '检查病人存在最近转出列表中
                        If mrsPatiInfo!排序 = 7 Then
                            strKey = ""
                            If mintREPORTSEL = 页面.转科 Then
                                If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                                    If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                        strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                                    End If
                                End If
                            End If

                            rptPati(页面.转科).Tag = "": rptPati(页面.转科).Records.DeleteAll
                            If rptPati(页面.转科).Columns.Count > c_审查 Then rptPati(页面.转科).Columns(c_审查).Visible = False
                            If PatiPage.Selected.Index = 页面.转科 Then Call PatiPage_SelectedChanged(PatiPage.Selected)
                            If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                        Else
                            Exit Sub
                        End If
                    End If
                    
                    If FreshPatiCard("新增在院病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
                        If strName <> "" Then
                            Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新撤销朱的病人:" & strName & "   床号:" & strBed, "撤销出院提醒")
                        End If
                    End If
                End If
                
                'c)完成待入科列表刷新
                If lngUnit = 0 Then
                    strValue = ""
                    strSQL = "Select 病区ID From 病区科室对应 where 科室ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病区信息", lngDept)
                    Do While Not rsTmp.EOF
                        strValue = strValue & "," & rsTmp!病区ID
                    rsTmp.MoveNext
                    Loop
                    strValue = Mid(strValue, 2)
                Else
                    strValue = lngUnit
                End If
                If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    If FreshPatiCard("新增转出待入科病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
                        If strName <> "" Then
                            Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新转入的病人:" & strName, "待办入住提醒")
                        End If
                    End If
                End If
            Case "转病区", "转科"
                '病人状态和刷新病区检查
                strSQL = "Select 出院病床 From 病案主页 Where 病人ID=[1] And 主页ID=[2] And NVL(状态,0)=0"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否正常在院", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!出院病床)
                'a)转入病区待入科列表更新
                If lngCurrUnit <> 0 Or lngCurrDept <> 0 Then
                    If lngCurrUnit = 0 Then
                        strValue = ""
                        strSQL = "Select 病区ID From 病区科室对应 where 科室ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病区信息", lngCurrDept)
                        Do While Not rsTmp.EOF
                            strValue = strValue & "," & rsTmp!病区ID
                        rsTmp.MoveNext
                        Loop
                        strValue = Mid(strValue, 2)
                    Else
                        strValue = lngCurrUnit
                    End If
                    If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                        mrsPatiInfo.Filter = "(病人ID=" & lngPatID & " And 主页ID=" & lngPageID & " And 排序=1) OR (病人ID=" & lngPatID & " And 主页ID=" & lngPageID & " And 排序=2)"
                        If Not mrsPatiInfo.EOF Then
                            mrsPatiInfo.Delete
                            strKey = ""
                            If mintREPORTSEL = 页面.待入科 Then
                                If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                                    If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                        strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                                    End If
                                End If
                            End If
                            rptPati(页面.待入科).Records.DeleteAll
                            If rptPati(页面.待入科).Columns.Count > c_审查 Then rptPati(页面.待入科).Columns(c_审查).Visible = False
                            Call UpgradeList(mrsPatiInfo, 页面.待入科)
                            PatiPage.Item(页面.待入科).Caption = "待入科" & GetPatiCount(页面.待入科) & "人"
                            If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                            
                            If strName <> "" Then
                                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新撤销转入的病人:" & strName, "撤销转入提醒")
                            End If
                        End If
                    End If
                End If
                'b)转出病区在院列表更新
                If InStr(1, "," & lngUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
                    blnFresh = False
                    Do While Not mrsPatiInfo.EOF
                        If InStr(1, ",4,3.2,3.1,", "," & mrsPatiInfo!排序 & ",") <> 0 Then
                            blnFresh = True
                            If mrsPatiInfo!排序 = 3.1 Then
                                mrsPatiInfo!状态 = 0
                            Else
                                mrsPatiInfo!排序 = 3: mrsPatiInfo!类型 = "在院病人": mrsPatiInfo!状态 = 0
                            End If
                            mrsPatiInfo.Update
                        End If
                    mrsPatiInfo.MoveNext
                    Loop
                    If blnFresh = False Then Exit Sub
                    mrsBedInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID & " And 包床<>1"
                    If Not mrsBedInfo.EOF Then
                        intCardIndex = mrsBedInfo!卡片索引
                        mrsBedInfo!病人状态 = 0
                        mrsBedInfo!病人状态名称 = ""
                        mrsBedInfo.Update
                        Call SetCardLabel(intCardIndex)
                    End If
                    mrsBedInfo.Filter = 0
                    If strName <> "" Then
                        Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新撤销转出的病人:" & strName & "   床号:" & strBed, "撤销转出提醒")
                    End If
                End If
            Case "入住", "入院入住"
                '病人状态和刷新病区检查
                strSQL = "Select 出院病床 From 病案主页 Where 病人ID=[1] And 主页ID=[2] And NVL(状态,0)=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否正常在院", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!出院病床)
                'a) 入住病区在院病人列表刷新
                If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    If FreshPatiCard("删除在院病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                    If strName <> "" Then
                        Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新撤销入住的病人:" & strName, "撤销入住提醒")
                    End If
                End If
                'b)  待入住病区待入科列表刷新
                If lngUnit = 0 Then
                    strValue = ""
                    strSQL = "Select 病区ID From 病区科室对应 where 科室ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病区信息", lngDept)
                    Do While Not rsTmp.EOF
                        strValue = strValue & "," & rsTmp!病区ID
                    rsTmp.MoveNext
                    Loop
                    strValue = Mid(strValue, 2)
                Else
                    strValue = lngUnit
                End If
                If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
                If FreshPatiCard("新增待入科病人", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                If strName <> "" Then
                    Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新代办入住的病人:" & strName, "待办入住提醒")
                End If
            End Select
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function FreshPatiCard(ByVal strType As String, ByVal lngPatID As Long, ByVal lngPageID As Long, ByVal lngUnit As Long) As Boolean
    Dim strSQL As String, strFields As String, strValues As String, strKey As String
    Dim rsTmp As New ADODB.Recordset, rsBed As New ADODB.Recordset
    Dim blnFresh As Boolean
    Dim intCardIndex As Integer, i As Long
    Dim arrCardIndex As Variant
    
    On Error GoTo ErrHand
    
    FreshPatiCard = False
    Select Case strType
    Case "新增在院病人"
        mrsBedInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
        If mrsBedInfo.RecordCount > 0 Then mrsBedInfo.Filter = "": Exit Function
        mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
        Do While Not mrsPatiInfo.EOF
            If mrsPatiInfo!排序 = 3.1 Or (mrsPatiInfo!排序 = 4 And Trim(Nvl(mrsPatiInfo!床号)) = "") Then
                Exit Function
            End If
        mrsPatiInfo.MoveNext
        Loop
        '提取病人信息
        strSQL = "Select /*+ RULE */ Decode(B.状态,3,4,DECODE(B.出院病床, NULL, 3.1,DECODE(B.状态,2,3.2,3))) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.状态,3,'预出院病人',DECODE(B.出院病床, NULL, '家庭病床',DECODE(B.状态,2,'预转科病人', '在院病人'))) as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名" & mstrBriefCode & ",NVL(b.性别,a.性别) 性别,NVL(b.年龄,a.年龄) 年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
            " B.出院病床 as 床号,E.名称 as 护理等级,B.费别,B.医疗付款方式,B.当前病况,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID,A.主页Id 最大主页Id" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z,在院病人 R" & _
            " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And A.主页ID=B.主页ID And Nvl(B.状态,0)<>1" & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And R.病区ID=[3] And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And a.病人ID=R.病人ID And A.当前病区ID=R.病区ID And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL" & _
            " And B.病人id =[1] And B.主页id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病人信息", lngPatID, lngPageID, lngUnit)
        If rsTmp.EOF Then Exit Function
        If rsTmp!排序 = 3.1 Or (rsTmp!排序 = 4 And Trim(Nvl(rsTmp!床号)) = "") Then '家庭病床
            Call UpgradeList(rsTmp)
            Call CopyReocrd(rsTmp)
            PatiPage.Item(页面.家庭病床).Caption = "家庭病床" & GetPatiCount(页面.家庭病床) & "人"
        Else
            strSQL = " Select Lpad(d.床号, 10, ' ') As 床号, Lpad(d.房间号, 10, ' ') 房间号, d.床位编制, Nvl(b.姓名, a.姓名) 姓名" & mstrBriefCode & ", b.住院号, b.病人id, b.主页id" & vbNewLine & _
                " From 病人信息 a, 病案主页 b, 在院病人 c, 床位状况记录 d" & vbNewLine & _
                " Where a.病人id = b.病人id And a.主页Id = b.主页id And a.病人id = c.病人id And a.病人id = d.病人id And a.当前病区id = c.病区id And" & vbNewLine & _
                "      a.当前病区id = d.病区id And b.病人id = [1] And b.主页id = [2] And c.病区id = [3]" & vbNewLine & _
                " Order By Lpad(d.床号, 10, ' ')"
            Set rsBed = zlDatabase.OpenSQLRecord(strSQL, "提取病人床位信息", lngPatID, lngPageID, lngUnit)
            If rsBed.EOF Then Exit Function
            Do While Not rsBed.EOF
                mrsBedInfo.Filter = "床号='" & Trim(Nvl(rsBed!床号, "ZYB")) & "'"
                If mrsBedInfo.RecordCount <> 0 Then
                    strFields = "床位编制|床号|住院号|姓名|简码|病人ID|主页ID|监护仪|病案审查|临床路径|个性标注1|病人状态|个性标注2|个性标注3|护理等级|病人类型|房间号|单病种"
                    strValues = Trim(rsBed!床位编制) & "|" & Trim(rsBed!床号) & "|" & Nvl(rsBed!住院号, 0) & "|" & rsBed!姓名 & "|" & Nvl(rsBed!简码) & "|" & Nvl(rsBed!病人ID, 0) & "|" & Nvl(rsBed!主页ID, 0) & "|0|0|0||0|||0|0|" & Trim(Nvl(rsBed!房间号)) & "|"
                    Call Record_Update(mrsBedInfo, strFields, strValues, "卡片索引|" & mrsBedInfo!卡片索引)
                    mlng空床 = mlng空床 - 1
                    mlng在床 = mlng在床 + 1
                End If
            rsBed.MoveNext
            Loop
            mrsBedInfo.Filter = ""
            Call UpgradeBeds(rsTmp)
            Call ShowGuage("数据读取结束", 100)
            Call AdjustCard
            Call CopyReocrd(rsTmp)
        End If
        FreshPatiCard = True
    Case "新增待入科病人"
        '开始加载病人信息
        mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
        If Not mrsPatiInfo.EOF Then Exit Function
        strSQL = "Select 0 排序, Decode(Nvl(b.病案状态, 0), 0, 999, b.病案状态) As 排序2, '入院待入住病人' As 类型, a.病人id, b.主页id, a.门诊号, b.住院号," & vbNewLine & _
            "       Nvl(b.姓名, a.姓名) 姓名" & mstrBriefCode & ", Nvl(b.性别, a.性别) 性别, Nvl(b.年龄, a.年龄) 年龄, d.名称 As 科室, c.科室id, c.经治医师 As 住院医师, b.责任护士, b.病案状态," & vbNewLine & _
            "       c.床号, e.名称 As 护理等级, b.费别,b.医疗付款方式, b.当前病况, Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期, b.出院日期, b.出院方式, b.病人类型, b.状态, b.险类, a.就诊卡号, -1 As 路径状态," & vbNewLine & _
            "       Trunc(Sysdate) - Trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数, z.颜色, b.单病种, b.婴儿科室id, b.婴儿病区id,A.主页Id 最大主页Id" & vbNewLine & _
            " From 病人信息 a, 病案主页 b, 病人变动记录 c, 部门表 d, 收费项目目录 e, 病人类型 z" & vbNewLine & _
            " Where a.在院 = 1 And b.病人类型 = z.名称(+) And a.病人id = b.病人id And Nvl(b.主页id, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id And" & vbNewLine & _
            "      (c.病区id = [3] Or c.病区id Is Null) And c.科室id = d.Id And (d.站点 = '" & gstrNodeNo & "' Or d.站点 Is Null) And b.护理等级id = e.Id(+) And" & vbNewLine & _
            "      Nvl(c.附加床位, 0) = 0 And c.终止时间 Is Null And c.开始原因 = 1 And b.状态 = 1 And Exists" & vbNewLine & _
            " (Select 1 From 病区科室对应 h Where c.科室id = h.科室id And h.病区id = [3]) And b.病人id = [1] And b.主页id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取待入科病人信息", lngPatID, lngPageID, lngUnit)
        If Not rsTmp.EOF Then
            Call UpgradeList(rsTmp)
            Call CopyReocrd(rsTmp)
            PatiPage.Item(页面.待入科).Caption = "待入科" & GetPatiCount(页面.待入科) & "人"
            FreshPatiCard = True
        End If
    Case "新增转出待入科病人"
        blnFresh = True
        mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
        Do While Not mrsPatiInfo.EOF
            If mrsPatiInfo!排序 = 1 Or mrsPatiInfo!排序 = 2 Then blnFresh = False: Exit Do
            mrsPatiInfo.MoveNext
        Loop
        If blnFresh = True Then
            '开始加载病人信息
            strSQL = " Select Decode(c.开始原因, 3, 1, 2) As 排序, Decode(Nvl(b.病案状态, 0), 0, 999, b.病案状态) As 排序2," & vbNewLine & _
                "       Decode(c.开始原因, 3, '转科待入住病人', '转病区待入住病人') As 类型, a.病人id, b.主页Id, a.门诊号, b.住院号," & vbNewLine & _
                "       Nvl(b.姓名, a.姓名) 姓名" & mstrBriefCode & ", Nvl(b.性别, a.性别) 性别, Nvl(b.年龄, a.年龄) 年龄, d.名称 As 科室, c.科室id," & vbNewLine & _
                "       c.经治医师 As 住院医师, b.责任护士, b.病案状态, c.床号, e.名称 As 护理等级, b.费别,b.医疗付款方式, b.当前病况, Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期, b.出院日期, b.出院方式, b.病人类型, b.状态, b.险类," & vbNewLine & _
                "       a.就诊卡号, -1 As 路径状态, Trunc(Sysdate) - Trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数, z.颜色, b.单病种, b.婴儿科室id, b.婴儿病区id,A.主页Id 最大主页Id" & vbNewLine & _
                " From 病人信息 a, 病案主页 b, 病人变动记录 c, 部门表 d, 收费项目目录 e, 病人类型 z" & vbNewLine & _
                " Where a.在院 = 1 And b.病人类型 = z.名称(+) And a.病人id = b.病人id And Nvl(b.主页id, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id And" & vbNewLine & _
                "      (c.病区id = [3] Or c.病区id Is Null) And c.科室id = d.Id And (d.站点 = '"" & gstrNodeNo & ""' Or d.站点 Is Null) And" & vbNewLine & _
                "      b.护理等级id = e.Id(+) And Nvl(c.附加床位, 0) = 0 And c.终止时间 Is Null And" & vbNewLine & _
                "      (c.开始原因 = 3 And Exists (Select 1 From 病区科室对应 h Where c.科室id = h.科室id And h.病区id = [3]) Or c.开始原因 = 15 And c.病区id = [3]) And" & vbNewLine & _
                "      (c.开始原因 In (3, 15) And c.开始时间 Is Null And b.状态 = 2) And b. 病人id = [1] And b.主页id = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取待入科病人信息", lngPatID, lngPageID, lngUnit)
            If Not rsTmp.EOF Then
                Call UpgradeList(rsTmp)
                Call CopyReocrd(rsTmp)
                PatiPage.Item(页面.待入科).Caption = "待入科" & GetPatiCount(页面.待入科) & "人"
                FreshPatiCard = True
            End If
        End If
    Case "删除在院病人"
        mrsBedInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
        If Not mrsBedInfo.EOF Then '在床病人
            blnFresh = False
            mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
            Do While Not mrsPatiInfo.EOF
                If InStr(1, ",3,3.2,", "," & mrsPatiInfo!排序 & ",") <> 0 Or (mrsPatiInfo!排序 = 4 And Trim(Nvl(mrsPatiInfo!床号)) <> "") Then
                    blnFresh = True
                    mrsPatiInfo.Delete
                End If
                mrsPatiInfo.MoveNext
            Loop
            If blnFresh = False Then mrsBedInfo.Filter = 0: Exit Function
            arrCardIndex = Array()
            Do While Not mrsBedInfo.EOF
                intCardIndex = mrsBedInfo!卡片索引
                ReDim Preserve arrCardIndex(UBound(arrCardIndex) + 1)
                arrCardIndex(UBound(arrCardIndex)) = intCardIndex
                '住院号,姓名,性别,年龄,诊断,医/护,费别,医疗付款方式,病况,入院日期,住院天数,余额,病人颜色,护理等级,就诊卡号）
                Call SetCardInfo(intCardIndex, Array("", "", "", "", "", "", "", "", "", "", "", "", &HFFFFFF, "", ""))
                mrsBedInfo.MoveNext
            Loop
            For i = 0 To UBound(arrCardIndex)
                strFields = "住院号|姓名|简码|病人ID|主页ID|监护仪|病案审查|临床路径|个性标注1|病人状态|个性标注2|个性标注3|护理等级|病人类型|单病种"
                strValues = "0|||0|0|0|0|0||0|||0|0|"
                Call Record_Update(mrsBedInfo, strFields, strValues, "卡片索引|" & Val(arrCardIndex(i)))
                
                picPati(Val(arrCardIndex(i))).ZOrder 1
                lblSelect(Val(arrCardIndex(i))).Visible = False
                If mblnCardCollapse Then
                    picPati(Val(arrCardIndex(i))).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
                    picPati(Val(arrCardIndex(i))).Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_折叠, 卡片背景_标准卡片_折叠)).Picture
                End If
                
                mlng空床 = mlng空床 + 1
                mlng在床 = mlng在床 - 1
            Next i
            mrsPatiInfo.Filter = ""
            mrsBedInfo.Filter = ""
            Call AdjustCard
        Else '非在床病人,就是家庭病床病人，如果为其他不做处理
            mrsBedInfo.Filter = 0
            mrsPatiInfo.Filter = "病人ID=" & lngPatID & " And 主页ID=" & lngPageID
            blnFresh = False
            Do While Not mrsPatiInfo.EOF
                If mrsPatiInfo!排序 = 3.1 Or (mrsPatiInfo!排序 = 4 And Trim(Nvl(mrsPatiInfo!床号)) = "") Then
                    blnFresh = True
                    mrsPatiInfo.Delete
                End If
                mrsPatiInfo.MoveNext
            Loop
            If blnFresh = False Then Exit Function
            
            strKey = ""
            If mintREPORTSEL = 页面.家庭病床 Then
                If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                    If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                        strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                    End If
                End If
            End If
            rptPati(页面.家庭病床).Records.DeleteAll
            If rptPati(页面.家庭病床).Columns.Count > c_审查 Then rptPati(页面.家庭病床).Columns(c_审查).Visible = False
            mlng家床 = 0: Call UpgradeList(mrsPatiInfo, 页面.家庭病床)
            PatiPage.Item(页面.家庭病床).Caption = "家庭病床" & GetPatiCount(页面.家庭病床) & "人"
            If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
        End If
        FreshPatiCard = True
    End Select
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub mfrmNoticeBoard_ItemClick(ByVal strBeds As String)
    Dim strKeys As String
    If strBeds = "" Then Exit Sub
    '根据床号获取病人ID(因为此处获取的床号为主床号)
    mrsBedInfo.Filter = ""
    Do While Not mrsBedInfo.EOF
        If InStr(1, "," & strBeds & ",", "," & Nvl(mrsBedInfo!床号) & ",") <> 0 Then
            strKeys = strKeys & "," & mrsBedInfo!病人ID
        End If
    mrsBedInfo.MoveNext
    Loop
    strKeys = Mid(strKeys, 2)
    If mblnHScroll Then
        If HScr.Value <> 0 Then
            mstrBoardKeys = strKeys
            HScr.Value = 0
        Else
            Call AdjustCard(, strKeys)
        End If
    Else
        Call AdjustCard(, strKeys)
    End If
End Sub

Private Sub mfrmResponse_Closed(ByVal DataChange As Boolean)
    If DataChange Then Call LoadResponse
End Sub

Private Sub mfrmResponse_OpenObject(ByVal PatiID As Long, ByVal PageID As Long, ByVal ObjectType As Integer, ByVal ObjectID As String)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngDept As Long
    Dim objRow As ReportRow
    Dim blnEnabled As Boolean, blnSeek As Boolean
    Dim strTab As String, strPrivs As String
    Dim objDoc As cEPRDocument
    Dim objEmr As Object, strReturn As String, strDocID As String, strSubdocID As String, rsEmr As ADODB.Recordset
        
    '当前病人为当前要定位的病人
    blnSeek = False
    
    mrsPatiInfo.Filter = "病人ID=" & PatiID & " and 主页ID=" & PageID
    blnSeek = mrsPatiInfo.RecordCount > 0
    If blnSeek = True Then
        lngDept = Val(mrsPatiInfo.Fields("科室ID").Value)
        mrsBedInfo.Filter = "病人ID=" & PatiID & " And 包床=0"
        If mrsBedInfo.RecordCount > 0 Then strTab = Nvl(mrsBedInfo.Fields("床号").Value)
        mrsBedInfo.Filter = ""
    End If
    mrsPatiInfo.Filter = 0
    If Not blnSeek Then
        MsgBox "当前病区病人清单中没有找到该病人。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call SelPatiCard(strTab, PatiID & "|" & PageID)
    If Not LocatePatiRecord Then
        MsgBox "定位病人失败,请在当前病区病人清单中核实病人是否存在。", vbInformation, gstrSysName
        Exit Sub
    End If

    '定位到对应的数据页面
    strTab = Decode(ObjectType, 1, "医嘱", 2, "病历", 3, "护理病历", 4, "护理", 5, "", 6, "医嘱", 7, "病历", 8, "病历")
    
    If ObjectType = 1 Or ObjectType = 4 Or ObjectType = 6 Then
        '判断权限
        blnSeek = False
        If ObjectType = 4 Then
            If GetInsidePrivs(p护理记录管理, True) <> "" Then
                blnSeek = True
            Else
                strTab = "护理"
            End If
        Else
            If GetInsidePrivs(p住院医嘱下达, True) <> "" Or GetInsidePrivs(p住院医嘱发送, True) <> "" Then
                blnSeek = True
            Else
                strTab = "医嘱"
            End If
        End If
        If blnSeek = False Then
            MsgBox "不能打开" & strTab & "页面,可能是您没有相应的权限。", vbInformation, gstrSysName
        Else
            Call InNurseRoutine(strTab)
            Call OrientTabPage_Rountine(strTab, ObjectID)
        End If
        Exit Sub
    End If
    
    '打开对应的对象
    Select Case ObjectType
    Case 1 '住院医嘱
    Case 2, 3, 7, 8 '住院病历,护理病历,疾病证明,知情文件
        If ObjectID = "0" Or ObjectID = "" Then Exit Sub
        If IsNumeric(ObjectID) Then
            Call gobjRichEPR.EditDocument(P新版护士站, Me, cboUnit.ItemData(cboUnit.ListIndex), ObjectID)
        Else '新版病历
            If gobjEmr Is Nothing Then Exit Sub
            If InStr(ObjectID, "|") = 0 Then
                strDocID = ObjectID
                strSubdocID = ""
            Else
                strDocID = Split(ObjectID, "|")(0)
                strSubdocID = Split(ObjectID, "|")(1)
            End If
            strSQL = "Select Hextoraw(c.Master_Id) Masterid, Hextoraw(c.Id) Actlogid, Hextoraw(c.Basiclog_Id) Basiclogid," & vbNewLine & _
                        "       Hextoraw(c.Action_Id) Actionid, Hextoraw(b.Id) Taskid, Hextoraw(b.Antetype_Id) Antetypeid, d.Type Doctype," & vbNewLine & _
                        "       Hextoraw(a.Id) Docid, 2 Occasion, a.Sealed Besealed, e.Code Docsecret, b.Subdoc_Id Subdocid,b.completor" & vbNewLine & _
                        "From Bz_Doc_Log A, Bz_Doc_Tasks B, Bz_Act_Log C, Antetype_List D, Secret_Grades E" & vbNewLine & _
                        "Where a.Actlog_Id = c.Id And a.Id = Hextoraw(:docid) And a.Id = b.Real_Doc_Id And " & IIf(strSubdocID = "", "", "b.Subdoc_Id = Hextoraw(:subdocid) And") & vbNewLine & _
                        "      b.Antetype_Id = d.Id And Decode(b.Subdoc_Id, Null, b.Antetype_Id, a.Antetype_Id) = a.Antetype_Id And" & vbNewLine & _
                        "      a.Secret = e.Code(+) And Rownum=1"
            strReturn = gobjEmr.OpenSQLRecordset(strSQL, strDocID & "^16^docid" & IIf(strSubdocID = "", "", "|" & strSubdocID & "^16^subdocid"), rsEmr)
            If strReturn <> "" Then Exit Sub
            
            strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, p电子病历管理) & ";"
            If Nvl(rsEmr!completor) = "" Then
                If InStr(strPrivs, ";文档书写;") > 0 Then '有书写权限
                    Call gobjEmr.OpenFormForModifyDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, rsEmr!basiclogid, rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, CInt(rsEmr!Occasion), CInt(rsEmr!besealed), CInt(rsEmr!docsecret), Nvl(rsEmr!subdocid), strPrivs)
                Else '无权限只能查看
                    Set objEmr = DynamicCreate("zlRichEMR.clsDockContent", "显示病历", True)
                    If Not objEmr Is Nothing Then
                        Call objEmr.Init(gobjEmr, gcnOracle, glngSys, 0)
                        Call objEmr.zlShowDoc(strDocID, strSubdocID)
                        Call objEmr.zlViewDoc(Me, "查阅病历", strSubdocID)
                    End If
                End If
            Else
                If InStr(strPrivs, ";文档审订;") > 0 Then '有书写权限
                    Call gobjEmr.OpenFormForAuditDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, rsEmr!basiclogid, rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, CInt(rsEmr!Occasion), CInt(rsEmr!besealed), CInt(rsEmr!docsecret), Nvl(rsEmr!subdocid), strPrivs)
                Else '无权限只能查看
                    Set objEmr = DynamicCreate("zlRichEMR.clsDockContent", "显示病历", True)
                    If Not objEmr Is Nothing Then
                        Call objEmr.Init(gobjEmr, gcnOracle, glngSys, 0)
                        Call objEmr.zlShowDoc(strDocID, strSubdocID)
                        Call objEmr.zlViewDoc(Me, "查阅病历", strSubdocID)
                    End If
                End If
            End If
        End If
    Case 4 '护理记录
    Case 5 '首页记录
        Call PrintInMedRec(mclsInOutMedRec, 1, PatiID, PageID, mobjReport, lngDept, Me)
    Case 6 '医嘱报告
        
    End Select
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub PatiPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim strSQL As String
    Dim strField As String, strValue As String
    Dim rsPati As New ADODB.Recordset
    Dim intSettle As Integer
    
    If Not mblnStart Then Exit Sub
    '修改此SQL的条件,病人事务管理模块也需要调整
    Dim i As Long
    
    Call picPatiIn_Resize
    Me.MousePointer = 11
    '61824:刘鹏飞,2013-05-23,显示单病种标志
    mintREPORTSEL = Item.Index
    strField = "排序|排序2|类型|病人ID|主页ID|住院号|姓名|简码|性别|年龄|科室|科室ID|住院医师|责任护士|病案状态|床号|护理等级|费别|医疗付款方式|当前病况|入院日期|出院日期|住院天数|出院方式|病人类型|状态|险类|就诊卡号|路径状态|颜色|单病种|婴儿科室ID|婴儿病区ID|最大主页Id"
    If rptPati(Item.Index).Tag = "" Then
        If Item.Index = 页面.出院 Or Item.Index = 页面.转科 Then
            If Item.Index = 页面.出院 Then
                '68259:刘鹏飞,2012-02-11,出院病人查找添加未结清已结清功能
                If chkSettle(0).Value = 1 And chkSettle(1).Value = 1 Then
                    intSettle = 0              '都显示
                ElseIf chkSettle(0).Value = 0 And chkSettle(1).Value = 1 Then
                    intSettle = 1               '只显示未结清的
                ElseIf chkSettle(0).Value = 1 And chkSettle(1).Value = 0 Then
                    intSettle = 2              '只显示已结清的
                End If
    
                '出院病人:出院病人可能已有多次住院
                strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
                    "Select /*+ RULE */ Decode(B.出院方式,'死亡',6,5) as 排序," & _
                    " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
                    " Decode(B.出院方式,'死亡','死亡病人','出院病人') as 类型," & _
                    " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名" & mstrBriefCode & ",NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
                    " B.出院病床 AS 床号,E.名称 as 护理等级,B.费别,B.医疗付款方式,B.当前病况,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期,B.出院日期,B.出院方式,B.病人类型," & _
                    " B.状态,B.险类,A.就诊卡号,Nvl(b.路径状态,-1) 路径状态,trunc(b.出院日期)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID,A.主页Id 最大主页Id" & _
                    " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z" & _
                    " Where A.病人ID=B.病人ID And B.病人类型=Z.名称(+) And Nvl(B.主页ID,0)<>0 And B.状态=0" & _
                    " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And B.当前病区ID+0=[1] And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
                    " And B.出院日期 Between [2] And [3] And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL" & _
                    IIf(intSettle = 0, "", " And " & IIf(intSettle = 1, "", "Not") & " Exists(Select 1 From 病人未结费用 Where B.病人id = 病人id  And B.主页id = 主页id and 来源途径=2 Having Nvl(Sum(金额), 0) <> 0)")
            Else
                '转出病人:在院,医生和床号显示本科转出前的
                strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
                    "Select /*+ RULE */ Distinct 7 as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,'转出病人' as 类型," & _
                    " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名" & mstrBriefCode & ",NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,D.名称 as 科室,C.科室ID,C.经治医师 as 住院医师,B.责任护士,B.病案状态," & _
                    " C.床号,E.名称 as 护理等级,B.费别,B.医疗付款方式,B.当前病况,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期,B.出院日期,B.出院方式,B.病人类型," & _
                    " B.状态,B.险类,A.就诊卡号,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID,A.主页Id 最大主页Id" & _
                    " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D,收费项目目录 E,病人类型 Z" & _
                    " Where A.病人ID=B.病人ID And B.病人类型=Z.名称(+) And Nvl(B.主页ID,0)<>0 And B.护理等级ID=E.ID(+)" & _
                    " And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
                    " And B.当前病区ID<>[1] And C.病区ID+0=[1] And C.科室ID=D.ID" & _
                    " And Nvl(C.附加床位,0)=0 And C.终止原因 In(3,15) And C.终止时间 Between Sysdate-[4] And Sysdate" & _
                    " And Nvl(B.状态,0)<>2 And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
            End If
            strSQL = strSQL & " Order by 排序,排序2,床号,主页ID Desc"

            On Error GoTo ErrHand
            Set rsPati = New ADODB.Recordset
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), _
                CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")), mintChange)
            
            Call UpgradeList(rsPati)
            
            '先删除原有记录集
            If Item.Index = 页面.出院 Then
                mrsPatiInfo.Filter = "排序=5 or 排序=6"
            Else
                mrsPatiInfo.Filter = "排序=7"
            End If
            For i = 1 To mrsPatiInfo.RecordCount
                mrsPatiInfo.Delete
                mrsPatiInfo.MoveNext
            Next
            
            '追加记录集
            mrsPatiInfo.Filter = 0
            If rsPati.RecordCount <> 0 Then rsPati.MoveFirst
            Do While Not rsPati.EOF
                strValue = rsPati!排序 & "|" & Nvl(rsPati!排序2, 0) & "|" & Nvl(rsPati!类型) & "|" & Nvl(rsPati!病人ID, 0) & "|" & Nvl(rsPati!主页ID, 0) & "|" & Nvl(rsPati!住院号, 0) & "|" & Nvl(rsPati!姓名) & "|" & Nvl(rsPati!简码) & "|" & Nvl(rsPati!性别) & "|" & _
                          Nvl(rsPati!年龄) & "|" & Nvl(rsPati!科室) & "|" & Nvl(rsPati!科室ID, 0) & "|" & Nvl(rsPati!住院医师) & "|" & Nvl(rsPati!责任护士) & "|" & Nvl(rsPati!病案状态, 0) & "|" & Nvl(rsPati!床号) & "|" & _
                          Nvl(rsPati!护理等级, "三级") & "|" & Nvl(rsPati!费别) & "|" & Nvl(rsPati!医疗付款方式) & "|" & Nvl(rsPati!当前病况, "一般") & "|" & Nvl(rsPati!入院日期) & "|" & Nvl(rsPati!出院日期) & "|" & Nvl(rsPati!住院天数) & "|" & Nvl(rsPati!出院方式) & "|" & _
                          Nvl(rsPati!病人类型, "普通病人") & "|" & Nvl(rsPati!状态, 0) & "|" & Nvl(rsPati!险类, 0) & "|" & Nvl(rsPati!就诊卡号) & "|" & Nvl(rsPati!路径状态, 0) & "|" & Nvl(rsPati!颜色, 0) & "|" & Nvl(rsPati!单病种) & "|" & Nvl(rsPati!婴儿科室ID, 0) & "|" & Nvl(rsPati!婴儿病区ID, 0) & "|" & Nvl(rsPati!最大主页ID, 0)
                Call Record_Add(mrsPatiInfo, strField, strValue)
                rsPati.MoveNext
            Loop
            
            rptPati(Item.Index).Tag = "OK"
            If GetPatiCount(Item.Index) <> 0 Then
                PatiPage.Item(Item.Index).Caption = IIf(Item.Index = 页面.出院, "最近出院", "最近转科") & GetPatiCount(Item.Index) & "人"
            End If
        End If
    End If
    pic出院查找.Visible = True
    pic出院查找.ZOrder 0

    If Item.Index = 页面.出院 Then
        '将当前页面的过滤条件显示在状态栏中
        Me.stbThis.Panels(2).Text = Format(mdtOutBegin, "yyyy-MM-dd") & "到" & Format(mdtOutEnd, "yyyy-MM-dd") & "之间" & IIf(intSettle = 0, "", IIf(intSettle = 1, "未结清", "已结清")) & "的出院病人"
    ElseIf Item.Index = 页面.转科 Then
        '将当前页面的过滤条件显示在状态栏中
        Me.stbThis.Panels(2).Text = "最近" & mintChange & "天内的转出病人"
    Else
        Me.stbThis.Panels(2).Text = ""
    End If
    
    Call GetPatiOtherInfo
    Me.MousePointer = 0
    
    On Error Resume Next
    If picList.Visible = True And rptPati(Item.Index).Visible = True Then rptPati(Item.Index).SetFocus
    If err <> 0 Then err.Clear
    
    Exit Sub
ErrHand:
    Me.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picPati_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, ""
End Sub

Private Sub pic出院查找_GotFocus()
    If txt住院号.Enabled And txt住院号.Visible Then txt住院号.SetFocus
End Sub

Private Sub rptPati_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object, i As Long
    Dim blnEnabled As Boolean, blnSelect As Boolean, blnWaitIn As Boolean
    Dim blnOut As Boolean, blnPreOut As Boolean, blnOutTo As Boolean, lngType As Long, strPrivs As String
    
    DoEvents
    mintREPORTSEL = Index
    If Button <> 2 Then Exit Sub

    '取病人基本信息
    blnSelect = LocatePatiRecord
    If blnSelect Then
        lngType = Val(mrsPatiInfo.Fields("排序").Value)
        blnWaitIn = lngType = pt转科待入住 Or lngType = pt入院待入住
        blnOut = lngType = pt出院
        blnPreOut = lngType = pt预出
        '85200:控制最近转出页面的病人不允许进行相关操作，如：撤销操作
        blnOutTo = lngType = pt最近转出
    Else
        Exit Sub
    End If
    '设置按钮状态
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p病人入出)
    If InStr(strPrivs, "所有病区") = 0 Then
        If InStr("," & mstrUnits & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
    End If

    '组装右键菜单
    Set cbrMenuBar = mobjPopup
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.IconId = cbrControl.IconId
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        cbrPopupItem.Visible = cbrControl.Visible

        Call SetControlVisible(cbrPopupItem)

        '设置按钮的状态
        Select Case cbrControl.ID
        Case conMenu_Manage_Change_Undo
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOutTo
            If cbrPopupItem.Enabled = True Then
                cbrPopupItem.Enabled = Val(Nvl(mrsPatiInfo.Fields("主页ID").Value, 0)) = Val(Nvl(mrsPatiInfo.Fields("最大主页Id").Value, 0))
            End If
            Call cbsMain_InitCommandsPopup(cbrMenuBar.CommandBar)
        Case conMenu_Manage_Change_In
            cbrPopupItem.Enabled = blnWaitIn
        Case conMenu_Manage_Change_InPati
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = mPatiInfo.性质 = 2
            End If
        Case conMenu_Manage_Change_Turn, conMenu_Manage_Change_Bed, conMenu_Manage_Change_House, _
             conMenu_Manage_Change_PatiInfo, conMenu_Manage_Change_ReCalcFee, conMenu_Manage_BedExchange
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = mrsPatiInfo.Fields("状态").Value <> 2
            End If
            If cbrPopupItem.ID = conMenu_Manage_Change_ReCalcFee And cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = Nvl(mrsPatiInfo.Fields("险类").Value, 0) = 0
            End If
        Case conMenu_Manage_Change_InsureSel
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = Nvl(mrsPatiInfo.Fields("险类").Value, 0) <> 0
            End If
        Case conMenu_Manage_Change_BedGrid
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = Trim(Nvl(mrsPatiInfo.Fields("床号").Value)) <> "" And mrsPatiInfo.Fields("状态").Value <> 2
            End If
        Case conMenu_Manage_Change_Out
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = (InStr(1, "," & pt在院 & ",3.1,", mrsPatiInfo.Fields("排序").Value) <> 0 Or blnPreOut) And mrsPatiInfo.Fields("状态").Value <> 2
            End If
        Case conMenu_Manage_Change_Baby
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = mPatiInfo.产科
            End If
        Case conMenu_Manage_Change_PaitNote
            cbrPopupItem.Enabled = Not blnOutTo
        Case conMenu_Manage_Monitor '监护仪
            cbrPopupItem.Visible = mblnMonitor And (InStr(GetInsidePrivs(p住院护士站), "护理监护") > 0)
        End Select
    Next
    If Not mrsPlugInBar Is Nothing Then
        mrsPlugInBar.Filter = "IsInTool=1 and BarType=3"
        For i = 1 To mrsPlugInBar.RecordCount
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mrsPlugInBar!功能ID, mrsPlugInBar!功能名)
                cbrPopupItem.IconId = mrsPlugInBar!图标ID
                cbrPopupItem.Parameter = mrsPlugInBar!功能名
                If Val(mrsPlugInBar!IsGroup) = 1 Then cbrPopupItem.BeginGroup = True
            mrsPlugInBar.MoveNext
        Next
        mrsPlugInBar.Filter = "IsInTool=0 and BarType=3"
        If mrsPlugInBar.RecordCount > 0 Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugInPop, "扩展功能"): cbrPopupItem.BeginGroup = True
                cbrPopupItem.IconId = conMenu_Tool_PlugIn
        End If
        mrsPlugInBar.Filter = 0
    End If
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptPati_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mintREPORTSEL = Index
    If Not LocatePatiRecord Then Exit Sub
    Call InNurseRoutine
End Sub


Private Sub rptPati_RowDblClick(Index As Integer, ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Childs.Count > 0 Then
        Row.Expanded = Not Row.Expanded
        Exit Sub
    End If
    mintREPORTSEL = Index
    If Not LocatePatiRecord Then Exit Sub
    Call InNurseRoutine
End Sub

Private Sub rptPati_SelectionChanged(Index As Integer)
    '53740:刘鹏飞,2012-09-19
    mintREPORTSEL = Index
    If Not LocatePatiRecord Then Exit Sub
    Call AutoExecutePlugIn(cbsMain)
    On Error Resume Next
    If picList.Visible = True And rptPati(Index).Visible = True Then rptPati(Index).SetFocus
    If err <> 0 Then err.Clear
End Sub

Private Sub rptPati_SortOrderChanged(Index As Integer)
    Dim objCol As ReportColumn
    Dim objRecord As ReportRecord, objParent As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsTemp As ADODB.Recordset, strFields As String, strValues As String, strKey As String
    Dim i As Long, j As Long, lngColor As Long
    Dim blnAsc As Boolean, lngIndex As Long
    '排序时，强行先按审查状态排序
    '子项排序功能无效，它随主项一起排序
    On Error GoTo ErrHand
    
    If rptPati(Index).SortOrder.Count > 0 Then
        If rptPati(Index).SortOrder(0).Index <> c_审查 Then
            Set objCol = rptPati(Index).SortOrder(0)
            rptPati(Index).SortOrder.DeleteAll
            rptPati(Index).SortOrder.Add rptPati(Index).Columns(c_审查)
            rptPati(Index).SortOrder.Add objCol
        Else
            '此判断为了以防万一，只有在点击审查列的时候COUNT=1，而审查列不可见。所以正常情况COUNT=2
            If rptPati(Index).SortOrder.Count > 1 Then
                Set objCol = rptPati(Index).SortOrder(1)
            Else
                Exit Sub
            End If
        End If
        blnAsc = objCol.SortAscending
        lngIndex = objCol.Index
        
        If lngIndex = c_图标 Then Exit Sub
        '86154:刘鹏飞,2015-07-02,ReportControl不支持子类排序处理
        For i = 0 To rptPati(Index).Records.Count - 1
            Set objParent = rptPati(Index).Records(i)
            If objParent.Childs.Count > 0 Then
                '初始化记录集
                strFields = "主键," & adVarChar & ",50|病人ID," & adDouble & ",20|主页ID," & adDouble & ",20|类型," & adVarChar & ",20|病案状态," & adDouble & ",10|" & _
                    "病案理由," & adLongVarChar & ",500|单病种," & adVarChar & ",10|路径状态," & adDouble & ",10|姓名," & adLongVarChar & ",100|" & _
                    "住院号," & adVarChar & ",20|床号," & adVarChar & ",20|性别," & adVarChar & ",20|年龄," & adVarChar & ",50|费别," & adVarChar & ",20|" & _
                    "付款方式," & adVarChar & ",30|住院医师," & adLongVarChar & ",100|入院日期," & adVarChar & ",20|出院日期," & adVarChar & ",20|" & _
                    "病人类型," & adVarChar & ",50|就诊卡号," & adVarChar & ",50"
                Call Record_Init(rsTemp, strFields)
                strFields = "主键|病人ID|主页ID|类型|病案状态|病案理由|单病种|路径状态|姓名|住院号|床号|性别|年龄|费别|付款方式|住院医师|入院日期|出院日期|病人类型|就诊卡号"
                For j = 0 To objParent.Childs.Count - 1
                    strKey = objParent.Childs(j).Item(C_病人ID).Value & "-" & objParent.Childs(j).Item(C_主页ID).Value
                    strValues = strKey & "'" & objParent.Childs(j).Item(C_病人ID).Value & "'" & objParent.Childs(j).Item(C_主页ID).Value & "'" & objParent.Childs(j).Item(c_类型).Value & "'" & _
                        objParent.Childs(j).Item(c_审查).Value & "'" & objParent.Childs(j).PreviewText & "'" & objParent.Childs(j).Item(c_图标).Value & "'" & _
                        objParent.Childs(j).Item(c_路径状态).Value & "'" & objParent.Childs(j).Item(c_姓名).Value & "'" & objParent.Childs(j).Item(c_住院号).Value & "'" & _
                        objParent.Childs(j).Item(c_床号).Value & "'" & objParent.Childs(j).Item(c_性别).Value & "'" & objParent.Childs(j).Item(c_年龄).Value & "'" & _
                        objParent.Childs(j).Item(c_费别).Value & "'" & objParent.Childs(j).Item(c_付款方式).Value & "'" & objParent.Childs(j).Item(c_医生).Value & "'" & _
                        objParent.Childs(j).Item(c_入院日期).Value & "'" & objParent.Childs(j).Item(c_出院日期).Value & "'" & objParent.Childs(j).Item(c_病人类型).Value & "'" & _
                        objParent.Childs(j).Item(c_就诊卡号).Value
                    Call Record_Update(rsTemp, strFields, strValues, "主键|" & strKey, , "'")
                Next j
                objParent.Childs.DeleteAll
                '根据选择的列排序
                strKey = ""
                Select Case lngIndex
                    Case c_类型
                        strKey = "类型"
                    Case c_审查
                        strKey = "病案状态"
                    Case c_图标
                        strKey = ""
                    Case c_路径状态
                        strKey = "路径状态"
                    Case C_病人ID
                        strKey = "病人ID"
                    Case C_主页ID
                        strKey = "主页ID"
                    Case c_姓名
                        strKey = "姓名"
                    Case c_住院号
                        strKey = "住院号"
                    Case c_床号
                        strKey = "床号"
                    Case c_性别
                        strKey = "性别"
                    Case c_年龄
                        strKey = "年龄"
                    Case c_费别
                        strKey = "费别"
                    Case c_付款方式
                        strKey = "付款方式"
                    Case c_医生
                        strKey = "住院医师"
                    Case c_入院日期
                        strKey = "入院日期"
                    Case c_出院日期
                        strKey = "出院日期"
                    Case c_病人类型
                        strKey = "病人类型"
                    Case c_就诊卡号
                        strKey = "就诊卡号"
                End Select
                
                rsTemp.Filter = ""
                If strKey <> "" Then rsTemp.Sort = strKey & IIf(blnAsc, "", " DESC")
                '排序后重新添加子类
                With rsTemp
                    Do While Not .EOF
                        Set objRecord = objParent.Childs.Add
                        objRecord.Tag = CStr(!病人ID & "|" & !主页ID)
                        Set objItem = objRecord.AddItem(CStr("" & !类型))
                        objItem.Caption = CStr("" & !类型)
                        
                        Set objItem = objRecord.AddItem(Val(Decode(Nvl(!病案状态, 0), 0, 999, Nvl(!病案状态, 0))))
                        objItem.Caption = " "
                        If Val(Nvl(!病案状态, 0)) = 2 Then
                            objRecord.PreviewText = "" & !病案理由
                        End If
                        
                        Set objItem = objRecord.AddItem(Nvl(!单病种))
                        objItem.Caption = " "
                        '81308:刘鹏飞,2015-01-19,显示病案审查标志
                        '61824:刘鹏飞,2013-05-23,显示单病种标志
                        If Nvl(!病案状态, 0) <> 0 Then
                            objItem.Icon = Get病案图标序号(!病案状态, False) - 1
                        ElseIf Nvl(!单病种) <> "" Then
                            objItem.Icon = imgRPT.ListImages("单病种").Index - 1
                        Else
                            objItem.Icon = Val(IIf(!性别 = "女", imgRPT.ListImages("女人").Index, imgRPT.ListImages("男人").Index)) - 1
                        End If
                        
                        '路径状态
                        Set objItem = objRecord.AddItem(Val("" & !路径状态))
                        objItem.Caption = " "
                        objItem.Icon = Get临床路径序号(Val("" & !路径状态) + 2, False) - 1
                        
                        objRecord.AddItem Val(!病人ID)
                        objRecord.AddItem Val(!主页ID)
                        objRecord.AddItem CStr(Nvl(!姓名))
                        Set objItem = objRecord.AddItem(CStr(Nvl(!住院号)))
                        objItem.Caption = Nvl(!住院号, " ")
                        Set objItem = objRecord.AddItem(Nvl(!床号))
                        objItem.Caption = CStr(Nvl(!床号, " "))
                        Set objItem = objRecord.AddItem(CStr(Nvl(!性别, "男")))
                        objItem.Caption = CStr(Nvl(!性别, "男"))
                        Set objItem = objRecord.AddItem(Nvl(!年龄, "0"))
                        objItem.Caption = Nvl(!年龄, "0")
                        Set objItem = objRecord.AddItem(Nvl(!费别, ""))
                        objItem.Caption = CStr(Nvl(!费别, ""))
                        Set objItem = objRecord.AddItem(Nvl(!付款方式, ""))
                        objItem.Caption = CStr(Nvl(!付款方式, ""))
                        Set objItem = objRecord.AddItem(Nvl(!住院医师, ""))
                        objItem.Caption = CStr(Nvl(!住院医师, ""))
                        Set objItem = objRecord.AddItem(CStr(Format(!入院日期, "yyyy-MM-dd HH:mm:ss")))
                        objItem.Caption = CStr(Format(!入院日期, "yyyy-MM-dd HH:mm:ss"))
                        Set objItem = objRecord.AddItem(CStr(Format(!出院日期, "yyyy-MM-dd HH:mm:ss")))
                        objItem.Caption = CStr(Format(!出院日期, "yyyy-MM-dd HH:mm:ss"))
                        Set objItem = objRecord.AddItem(Nvl(!病人类型, "普通病人"))
                        objItem.Caption = CStr(Nvl(!病人类型, "普通病人"))
                        Set objItem = objRecord.AddItem(CStr(Nvl(!就诊卡号)))
                        objItem.Caption = Nvl(!就诊卡号, "")
                        '提取病人类型的颜色
                        lngColor = 0
                        mrsPatiColor.Filter = "名称='" & Nvl(!病人类型, "普通病人") & "'"
                        If mrsPatiColor.RecordCount <> 0 Then
                            lngColor = Nvl(mrsPatiColor!颜色, 0)
                        End If
                        If lngColor <> 0 Then
                            objRecord.Item(c_姓名).ForeColor = lngColor
                        End If
                        
                    .MoveNext
                    Loop
                End With
                rptPati(Index).Populate
            End If
        Next i
    End If
    Exit Sub
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "病人颜色" Then
        Call zlDatabase.ShowPatiColorTip(Me)
    End If
End Sub

Private Sub timKey_Timer()
    Static strPreTime As String
    Dim curTime As Date
    Dim blnRefresh As Boolean
    
    If timNotify.Enabled = False Then timNotify.Enabled = True
    If timeRefreshCard.Enabled = False Then timeRefreshCard.Enabled = True
    If cboUnit.ListIndex <> -1 Then
        timKey.Enabled = False
        strPreTime = ""
        Exit Sub
    End If
    
    curTime = Now
    If Me.ActiveControl.Name <> "cboUnit" Then
        blnRefresh = True
    Else
        If strPreTime = "" Then strPreTime = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        '30s输入不做任何响应则自动还原
        If DateDiff("s", CDate(strPreTime), curTime) > CLng(30) Then
            blnRefresh = True
        End If
    End If
    If IsNumeric(timKey.Tag) And blnRefresh Then
        cboUnit.ListIndex = Val(timKey.Tag)
        timKey.Enabled = False
        strPreTime = ""
    End If
End Sub

Private Sub timNotify_Timer()
    Static strPreTime1 As String
    Static strPreTime2 As String
    Dim curTime As Date
    
    If blnUnload Then Exit Sub
    If mblnRefresh Then Exit Sub
    curTime = Now
    
    '刷新病案审阅反馈：每5分钟
    If strPreTime2 = "" Then strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
    If DateDiff("s", CDate(strPreTime2), curTime) > 5 * CLng(60) Then
        strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        Call LoadResponse
    End If
End Sub

Public Sub SelPatiCard(ByVal strBed As String, ByVal strKey As String)
    Dim intIndex As Integer
    Dim intPage As Integer
    Dim blnFind As Boolean
    On Error GoTo ErrHand
    '提供给外部程序的接口,选中指定病人的床位卡
    
    If strBed <> "" Then
        mrsBedInfo.Filter = "床号='" & strBed & "'"
        If mrsBedInfo.RecordCount <> 0 Then intIndex = mrsBedInfo!卡片索引
        mrsBedInfo.Filter = 0
    End If
    
    If intIndex > 0 Then
        '取消上次选择
        Call picPati_MouseDown(intIndex, 1, 0, 0, 0)
        '选择指定卡片
        mblnShow = False            '必须加,不然又会触发ShowSelect
        Call ShowSelect
        '避免卡片显示于最上面
        Call picPati_MouseUp(intIndex, 1, 0, 0, 0)
        '将选中卡片显示在可视区域内
        Call ShowEfficiency
    Else
        '非在床病人
        intPage = -1
        mrsPatiInfo.Filter = "病人ID=" & Split(strKey, "|")(0) & " And 主页ID=" & Split(strKey, "|")(1)
        If mrsPatiInfo.RecordCount <> 0 Then
            If mrsPatiInfo!排序 = 0 Or mrsPatiInfo!排序 = 1 Or mrsPatiInfo!排序 = 2 Then
                intPage = 0
            ElseIf mrsPatiInfo!排序 = 7 Then
                intPage = 1
            ElseIf mrsPatiInfo!排序 = 6 Or mrsPatiInfo!排序 = 5 Then
                intPage = 2
            ElseIf mrsPatiInfo!排序 = 3.1 Then '家庭病床
                intPage = 3
            End If
        End If
        mrsPatiInfo.Filter = 0
        
        If intPage <> -1 Then
            PatiPage(intPage).Selected = True
            mintREPORTSEL = intPage
            
            '查找定位病人
            Dim objRptRow As ReportRow
            For Each objRptRow In rptPati(intPage).Rows
                If Not objRptRow.Record Is Nothing Then
                    If objRptRow.Record.Childs.Count = 0 Then
                        If Val(objRptRow.Record.Item(C_病人ID).Value) = Val(Split(strKey, "|")(0)) And _
                            Val(objRptRow.Record.Item(C_主页ID).Value) = Val(Split(strKey, "|")(1)) Then
                            rptPati(intPage).Rows(objRptRow.Index).Selected = True
                            blnFind = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowEfficiency()
'点击医嘱提醒，将选中的病人显示在有效区域内
    Dim lngTop As Long, lngY As Long
    Dim lngMove As Long
    
    lngMove = CLng((mdblScaleHeight - (PicDraw.Height - IIf(picList.Visible, picList.Height, 0))) / 100) '获取步长
    If lngMove <= 0 Then lngMove = 1
    lngY = clngX + picPati(mlngSelect).Height
    lngTop = picPati(mlngSelect).Top - (-1 * HScr.Value * lngMove)  '获取原始卡片的Top
    If (lngTop - lngY) / lngMove > HScr.Max Then
        HScr.Value = HScr.Max
    ElseIf (lngTop - lngY) / lngMove < HScr.Min Then
        HScr.Value = HScr.Min
    Else
        HScr.Value = (lngTop - lngY) / lngMove
    End If
    Call HScr_Change
End Sub

Public Sub ExecFuncs(ByVal intFunc As Integer)
    Dim lngKey As Long
    Dim blnSel As Boolean
    Dim objControl As CommandBarControl, objParent As CommandBarPopup
    On Error GoTo ErrHand
    '54370:刘鹏飞,2013-05-02,添加参数"医嘱处理后自动定位到医嘱页面"
    '提供给医嘱提醒的专用接口,非病区性批量功能
BeginFunc:
    Select Case intFunc
    Case E发送
        lngKey = conMenu_Edit_Send
    Case E校对
        lngKey = conMenu_Edit_Audit
    Case E停止
        lngKey = conMenu_Edit_ReStop
    '55430:刘鹏飞,2013-02-27,双击作废医嘱定位到病人事物的医嘱页面
    Case E查看
        lngKey = conMenu_病人事务处理
    Case 11 '输液审核未通过
        lngKey = conMenu_病人事务处理
    Case 12 '费用销帐申请
        lngKey = conMenu_Edit_ReBillingApply
    End Select
    Select Case intFunc
    Case E查看
        Set objParent = cbsMain.Item(1).Controls.Item(3)        '病区批量工作
    Case E发送, E校对, E停止
        Set objParent = cbsMain.Item(1).Controls.Item(4)        '医嘱业务
    Case 11 '输液审核未通过
        Set objParent = cbsMain.Item(1).Controls.Item(3)        '病区批量工作
    Case 12 '费用销帐申请
        Set objParent = cbsMain.Item(1).Controls.Item(5)        '费用业务
    End Select
    For Each objControl In objParent.CommandBar.Controls
        If objControl.ID = lngKey Then blnSel = True: Exit For
    Next
    If blnSel Then
        objControl.Execute
        If intFunc = E查看 Or intFunc = 11 Then
            Call OrientTabPage_Rountine
        ElseIf intFunc = E发送 Or intFunc = E校对 Or intFunc = E停止 Then
            If mblnCollateAutoFind = True Then intFunc = E查看: GoTo BeginFunc
        End If
    End If
    frmNotify.mblnFirst = True
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function LoadResponse() As Boolean
'功能：读取病案审查反馈
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngCount As Long
    Dim curDate As Date
    
    If cboUnit.ListIndex = -1 Then
        fra审查.Visible = False: LoadResponse = True: Exit Function
    End If

    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    Screen.MousePointer = 11

    '读取当前病区的在院、出院病人，以"病案反馈记录"为准全部扫描
    strSQL = "Select Count(*) as 数量 From 病案主页 B,病案反馈记录 A" & _
        " Where A.病人ID=B.病人ID and A.主页ID=B.主页ID And A.记录状态=1" & _
        " And A.反馈对象 IN(3,4) And B.当前病区ID + 0 =[1]" & _
        " And a.反馈时间 Between [2] And [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LoadResponse", cboUnit.ItemData(cboUnit.ListIndex), CDate(Format(curDate - mlngMedRedDay, "yyyy-MM-dd")), CDate(Format(curDate, "yyyy-MM-dd HH:mm:ss")))
    If Not rsTmp.EOF Then lngCount = Nvl(rsTmp!数量, 0)
    
    lbl审查.Caption = mlngMedRedDay & "天内共有 " & lngCount & " 条未处理的病案审查反馈..."
    fra审查.Visible = lngCount > 0
    lbl审查.Tag = lngCount

    Screen.MousePointer = 0
    LoadResponse = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub init非在床清单()
    Dim objCol As ReportColumn
    '初始化非在床病人清单
    PatiPage.Item(页面.待入科).Caption = "待入科"
    PatiPage.Item(页面.转科).Caption = "最近转科"
    PatiPage.Item(页面.出院).Caption = "最近出院"
    PatiPage.Item(页面.家庭病床).Caption = "家庭病床"

    rptPati(页面.待入科).Tag = ""       '此标记用来判断数据是否需要刷新
    rptPati(页面.转科).Tag = ""
    rptPati(页面.出院).Tag = ""
    rptPati(页面.家庭病床).Tag = ""

    rptPati(页面.待入科).Records.DeleteAll
    rptPati(页面.转科).Records.DeleteAll
    rptPati(页面.出院).Records.DeleteAll
    rptPati(页面.家庭病床).Records.DeleteAll
    
    Call InitReportControl(页面.待入科)
    Call InitReportControl(页面.转科)
    Call InitReportControl(页面.出院)
    Call InitReportControl(页面.家庭病床)
End Sub

Private Sub InitReportControl(ByVal intIndex As Integer)
    Dim arrWith() As String
    Dim objCol As ReportColumn
    
    arrWith = Split(mstrColWidth, ",")
    With rptPati(intIndex)
        .Columns.DeleteAll
        Set objCol = .Columns.Add(c_类型, "类型", Val(arrWith(c_类型)), True): objCol.Groupable = True ': objCol.Visible = IIf(intIndex = 页面.待入科, True, IIf(intIndex = 页面.出院, True, False))
        Set objCol = .Columns.Add(c_审查, "", Val(arrWith(c_审查)), False): objCol.TreeColumn = True: objCol.Visible = False: objCol.Sortable = False: objCol.AllowDrag = False
       Set objCol = .Columns.Add(c_图标, "", Val(arrWith(c_图标)), False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_路径状态, "路径状态", Val(arrWith(c_路径状态)), True): objCol.Visible = mblnHavePath
        Set objCol = .Columns.Add(C_病人ID, "病人ID", Val(arrWith(C_病人ID)), False): objCol.Visible = False
        Set objCol = .Columns.Add(C_主页ID, "主页ID", Val(arrWith(C_主页ID)), False): objCol.Visible = False
        Set objCol = .Columns.Add(c_姓名, "姓名", Val(arrWith(c_姓名)), True)
        Set objCol = .Columns.Add(c_住院号, "住院号", Val(arrWith(c_住院号)), True)
        Set objCol = .Columns.Add(c_床号, "床号", Val(arrWith(c_床号)), True)
        Set objCol = .Columns.Add(c_性别, "性别", Val(arrWith(c_性别)), True)
        Set objCol = .Columns.Add(c_年龄, "年龄", Val(arrWith(c_年龄)), True)
        Set objCol = .Columns.Add(c_费别, "费别", Val(arrWith(c_费别)), True)
        Set objCol = .Columns.Add(c_付款方式, "医疗付款方式", Val(arrWith(c_付款方式)), True)
        Set objCol = .Columns.Add(c_医生, "医生", Val(arrWith(c_医生)), True)
        Set objCol = .Columns.Add(c_入院日期, "入院日期", Val(arrWith(c_入院日期)), True): objCol.SortAscending = False
        Set objCol = .Columns.Add(c_出院日期, "出院日期", Val(arrWith(c_出院日期)), True): objCol.Visible = IIf(intIndex = 页面.出院, True, False)
        Set objCol = .Columns.Add(c_病人类型, "病人类型", Val(arrWith(c_病人类型)), True)
        Set objCol = .Columns.Add(c_就诊卡号, "就诊卡号", Val(arrWith(c_就诊卡号)), True): objCol.Visible = mblnShowCard
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Sortable = True
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有病人..."
        End With
        .TabStop = True
        .PreviewMode = True
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgRPT
    
        .GroupsOrder.DeleteAll
        If intIndex = 页面.待入科 Or intIndex = 页面.出院 Then
            .GroupsOrder.Add .Columns.Find(c_类型)
            .GroupsOrder(0).SortAscending = True
        End If
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns.Find(c_审查)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(c_入院日期)
    End With
End Sub

Private Function InitBedlevel() As Boolean
'功能：初始化床位状况数据
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    cbo床位状况.Clear
    cbo床位状况.AddItem "全部"
    On Error GoTo errH
    strSQL = _
        " Select 名称 from 床位编制分类 Order by DECODE(缺省标志,1,1,2)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InitNurselevel")
    Do While Not rsTmp.EOF
        cbo床位状况.AddItem rsTmp!名称
        rsTmp.MoveNext
    Loop
    cbo床位状况.ListIndex = 0

    InitBedlevel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitNurselevel() As Boolean
'功能：初始化住院护理等级条件
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSel As String
    Dim blnSelAll As Boolean

    txt护理条件.Text = ""
    txt护理条件.Tag = ""

    lst护理条件.AddItem "全部"
    strSel = zlDatabase.GetPara("护理等级过滤", glngSys, p住院护士站, "", Array(txt护理条件, cmd护理条件), InStr(mstrPrivs, "参数设置") > 0)
    blnSelAll = True
    On Error GoTo errH
    strSQL = _
        " Select ID,编码,名称 From 收费项目目录 Where 类别='H' And 项目特性>=1" & _
        " And (撤档时间 is NULL Or Trunc(撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " And (站点='" & gstrNodeNo & "' Or 站点 is Null)" & _
        " Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InitNurselevel")
    Do While Not rsTmp.EOF
        lst护理条件.AddItem rsTmp!名称
        lst护理条件.ItemData(lst护理条件.NewIndex) = rsTmp!ID
        If strSel = "" Or InStr("," & strSel & ",", "," & rsTmp!ID & ",") > 0 Then
            txt护理条件.Text = txt护理条件.Text & "," & rsTmp!名称
            txt护理条件.Tag = txt护理条件.Tag & "," & rsTmp!ID
        Else
            blnSelAll = False
        End If
        rsTmp.MoveNext
    Loop

    If blnSelAll Then
        txt护理条件.Text = "全部"
        txt护理条件.Tag = ""
    Else
        txt护理条件.Text = Mid(txt护理条件.Text, 2)
        txt护理条件.Tag = Mid(txt护理条件.Tag, 2)
    End If

    InitNurselevel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院护理病区
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long

    On Error GoTo errH
    mstrUnits = GetUser病区IDs

    '包含门观察室
    If InStr(mstrPrivs, "全院病人") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=A.科室ID)" & _
            " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=A.科室ID)" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If

    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If InStr(mstrPrivs, "全院病人") > 0 Then
                If rsTmp!ID = UserInfo.部门ID Then '直接所属优先
                    Call zlControl.CboSetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
                If InStr("," & mstrUnits & ",", "," & rsTmp!ID & ",") > 0 And cboUnit.ListIndex = -1 Then
                    Call zlControl.CboSetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            Else '所属缺省病区包含的可能有多个
                If rsTmp!缺省 = 1 And cboUnit.ListIndex = -1 Then
                    Call zlControl.CboSetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboUnit.hwnd, 0)
    End If
    mintPreDept = cboUnit.ListIndex
    InitUnits = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetDataToUnits(Optional ByVal strIn As String = "") As ADODB.Recordset
'功能：获取科室列表数据记录集
'参数：strIn 过滤条件
    Dim strSQL As String
    Dim blnYN As Boolean
    
    If strIn <> "" Then blnYN = True
    If InStr(mstrPrivs, "全院病人") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            IIf(blnYN, " And (A.编码 Like [2] Or A.简码 Like [3] Or A.名称 Like [3])", "") & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=A.科室ID)" & _
            " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=A.科室ID)" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            IIf(blnYN, " And (C.编码 Like [2] Or C.简码 Like [3] Or C.名称 Like [3])", "") & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If
    
    On Error GoTo errH
    If blnYN Then
        Set GetDataToUnits = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, UCase(strIn) & "%", gstrLike & UCase(strIn) & "%")
    Else
        Set GetDataToUnits = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadBeds() As Boolean
    '装载当前病区的床位
    Dim lngX As Long, lngY As Long, lngRowCount As Long
    Dim rsBeds As New ADODB.Recordset
    Dim strBriefCode As String
    
    On Error GoTo ErrHand
    
    lngX = clngX
    lngY = clngX
    lngRowCount = (PicDraw.Width - HScr.Width - 50) \ (picPati(mlngSource).Width + 15)
    Call UnloadControls
    PicDraw.Refresh
    'debug.print "卸载床位卡片:" & Now
    
    If mblnSupport Then
        strBriefCode = ",zlpinyincode(c.姓名,0,0,',',1) AS 简码 "
    Else
        strBriefCode = ",zlspellcode(c.姓名) AS 简码"
    End If
    
    '60800:刘鹏飞,2013-07-09,不显示修缮的床位
    '提取本病区的床位
'    mstrSQL = " SELECT Lpad(b.床号, 10, ' ') AS 床号, Lpad(b.房间号, 10, ' ') 房间号, b.床位编制, a.姓名" & strBriefCode & ", a.住院号, a.病人id, a.主页id" & vbNewLine & _
'            " FROM 床位状况记录 b," & vbNewLine & _
'            "     (SELECT NVL(c.姓名,b.姓名) || Decode(c.婴儿病区id, NULL, '', '之子') 姓名, b.住院号, b.病人id, b.主页id" & vbNewLine & _
'            "       FROM 病人信息 b, 病案主页 c, 在院病人 r" & vbNewLine & _
'            "       WHERE b.病人id = r.病人id AND c.病人id = b.病人id AND b.主页id = c.主页id AND b.当前病区id = r.病区id AND" & vbNewLine & _
'            "             (r.病区id = [1] OR c.婴儿病区id = [1])) a" & vbNewLine & _
'            " WHERE b.病人id = a.病人id(+) AND b.病区id = [1] And NOT (b.状态='修缮' And b.病人ID IS NULL)" & vbNewLine & _
'            " ORDER BY Lpad(b.床号, 10, ' ')"
    '74214:刘鹏飞,2013-06-20,性能优化
    '性能优化
    '78761:刘鹏飞,2014-11-03,床号按床位编制编码排序
    mstrSQL = " Select LPad(b.床号, 10, ' ') As 床号, LPad(b.房间号, 10, ' ') 房间号, b.床位编制, c.姓名" & strBriefCode & ",c.住院号," & vbNewLine & _
            "       c.病人id, c.主页id,trunc(sysdate)-trunc(DECODE(C.入科时间,NULL,C.入院日期,C.入科时间)) as 住院天数" & vbNewLine & _
            " From 床位状况记录 B, 病案主页 C, 床位编制分类 D" & vbNewLine & _
            " Where b.病区id =[1] And (c.当前病区id = b.病区id Or c.婴儿病区id = b.病区id Or b.病人ID is NULL)" & vbNewLine & _
            "      And b.病人id = c.病人id(+) And c.出院日期(+) is Null And B.床位编制=D.名称(+) " & vbNewLine & _
            "      And Not (b.状态 = '修缮' And b.病人id Is Null)"
    If mblnCardOrder = True Then
        mstrSQL = mstrSQL & vbNewLine & " Order By LPad(b.床号, 10, ' ')"
    Else
        mstrSQL = mstrSQL & vbNewLine & " Order By D.编码,LPad(b.床号, 10, ' ')"
    End If
    
    Set rsBeds = zlDatabase.OpenSQLRecord(mstrSQL, "装载当前病区的床位", cboUnit.ItemData(cboUnit.ListIndex))
    With rsBeds
        If .RecordCount = 0 Then
            MsgBox "当前病区还没有床位，请在病区床位管理中进行添加！", vbInformation, gstrSysName
            Exit Function
        End If
        
        Do While Not .EOF
            '更新内存映射记录集
            mstrFields = "卡片索引|床位编制|床号|住院号|姓名|简码|病人ID|主页ID|监护仪|病案审查|临床路径|个性标注1|病人状态|个性标注2|个性标注3|护理等级|病人类型|房间号|单病种|住院天数"
            mstrValues = .AbsolutePosition & "|" & Trim(!床位编制) & "|" & Trim(!床号) & "|" & Nvl(!住院号, 0) & "|" & !姓名 & "|" & Nvl(!简码) & "|" & Nvl(!病人ID, 0) & "|" & Nvl(!主页ID, 0) & "|0|0|0||0|||0|0|" & Trim(Nvl(!房间号)) & "||" & IIf(IsNull(!住院天数), "NULL", IIf(Val("" & !住院天数) = 0, 1, Val("" & !住院天数)))
            Call Record_Add(mrsBedInfo, mstrFields, mstrValues)
            '添加空白卡片
            Call LoadPatiCard(.AbsolutePosition, IIf(Val(lbl床号(mlngSource).Tag) = 1, IIf(Trim(Nvl(!房间号)) = "", "", Trim(!房间号) & IIf(IsNumeric(Trim(!房间号)), "_", "")), "") & Trim(!床号), lngX, lngY)
            
            If Nvl(!病人ID, 0) = 0 Then
                mlng空床 = mlng空床 + 1
            Else
                mlng在床 = mlng在床 + 1
            End If
            
            '计算下一张卡片的坐标
            lngX = lngX + picPati(mlngSource).Width '+ 30
            If .AbsolutePosition Mod lngRowCount = 0 Then
                lngX = clngX
                lngY = lngY + picPati(mlngSource).Height '+ 30
            End If
            
            .MoveNext
        Loop
    End With
    
    picList.ZOrder 0
    PatiPage.ZOrder 0
    fraPatiUD.ZOrder 0
    picPara(2).ZOrder 0
    picPara(3).ZOrder 0
    pic出院查找.ZOrder 0
    
    'debug.print "完成床位卡片装载:" & Now
    LoadBeds = True
    
    mdblScaleHeight = (lngY + picPati(mlngSource).Height) ' + 30)
    mblnHScroll = (mdblScaleHeight > PicDraw.Height - IIf(picList.Visible, picList.Height, 0))
    With HScr
        .Value = 0
        .Top = PicDraw.Top
        .Left = PicDraw.Width - .Width
        .Height = PicDraw.Height
        .Visible = mblnHScroll
    End With
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UpgradeList(ByVal rsPati As ADODB.Recordset, Optional ByVal intCurPage As Integer = -1) As Boolean
    '装载不在床的病人清单
    Dim j As Integer
    Dim str类型 As String
    Dim intPage As Integer
    Dim lngColor As Long
    Dim objItem As ReportRecordItem
    Dim objRecord As ReportRecord
    Dim objRpt As ReportControl
    Dim objParent As ReportRecord
    
    On Error GoTo ErrHand
    
    With rsPati
        '排序:0-转科待入科;1-入院待入科;2.2-家庭病床;4-出院;5-死亡;6-转出
        .Filter = "类型 <>'在院病人' " ' AND 类型 <>'预出院病人' " ' AND 类型 <>'转科待入住病人' AND 类型 <>'转病区待入住病人' AND 类型 <>'入院待入住病人'"
        '.Sort = " 入院日期 desc "
        .Sort = "排序,排序2,床号,主页ID Desc"
        Do While Not .EOF
            intPage = -1
            If !排序 = 0 Or !排序 = 1 Or !排序 = 2 Then
                intPage = 0
            ElseIf !排序 = 7 Then
                intPage = 1
            ElseIf !排序 = 6 Or !排序 = 5 Then
                intPage = 2
            ElseIf !排序 = 3.1 Or (!排序 = 4 And Nvl(!床号) = "") Then '家庭病床
                intPage = 3
                mlng家床 = mlng家床 + 1
            End If
            
            If intPage > -1 And IIf(intCurPage = -1, True, intPage = intCurPage) Then
                Select Case Nvl(!排序)
                Case 0
                    str类型 = "入院"
                Case 1
                    str类型 = "转科"
                Case 2
                    str类型 = "转病区"
                Case 5
                    str类型 = "出院"
                Case 6
                    str类型 = "死亡"
                End Select
                '根据提交审查情况添加父行
                If Nvl(!病案状态, 0) <> 0 Then
                    rptPati(intPage).Columns(c_审查).Visible = True
                    If objParent Is Nothing Then
                        Set objParent = Me.rptPati(intPage).Records.Add()
                    ElseIf objParent.Tag <> CStr(!病案状态) Then
                        Set objParent = Me.rptPati(intPage).Records.Add()
                    End If
                    If objParent.Tag <> CStr(!病案状态) Then
                        objParent.Tag = CStr(!病案状态)
                        objParent.Expanded = True
                        For j = 0 To rptPati(intPage).Columns.Count - 1
                            If j = c_类型 Then
                                Set objItem = objParent.AddItem(Val(!排序))
                                objItem.Caption = str类型
                            ElseIf j = c_审查 Then
                                Set objItem = objParent.AddItem(Val(Decode(Nvl(!病案状态, 0), 0, 999, !病案状态)))
                                objItem.Caption = " "
                            ElseIf j = c_姓名 Then
                                Set objItem = objParent.AddItem(Get病案主题(!病案状态))
                                objItem.ForeColor = rptPati(intPage).PaintManager.GroupForeColor
                            Else
                                Set objItem = objParent.AddItem("")
                                If j = c_图标 Then objItem.Icon = Get病案图标序号(!病案状态, False) - 1
                            End If
                            objItem.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
                        Next
                    End If
                Else
                    Set objParent = Nothing
                End If
                
                '添加具体的病人数据行(子行或独立行)
                If Not objParent Is Nothing Then
                    Set objRecord = objParent.Childs.Add()
                Else
                    Set objRecord = Me.rptPati(intPage).Records.Add()
                End If
                
                objRecord.Tag = CStr(!病人ID & "|" & !主页ID)
                
                Set objItem = objRecord.AddItem(str类型)
                objItem.Caption = str类型
                
                Set objItem = objRecord.AddItem(Val(Decode(Nvl(!病案状态, 0), 0, 999, !病案状态)))
                objItem.Caption = " "
                If Nvl(rsPati!病案状态, 0) = 2 Then
                    objRecord.PreviewText = "  理由:" & GetRefuseReason(Val(!病人ID), Val(!主页ID))
                End If
                
                Set objItem = objRecord.AddItem(Nvl(!单病种))
                objItem.Caption = " "
                '81308:刘鹏飞,2015-01-19,显示病案审查标志
                '61824:刘鹏飞,2013-05-23,显示单病种标志
                If Nvl(!病案状态, 0) <> 0 Then
                    objItem.Icon = Get病案图标序号(!病案状态, False) - 1
                ElseIf Nvl(!单病种) <> "" Then
                    objItem.Icon = imgRPT.ListImages("单病种").Index - 1
                Else
                    objItem.Icon = Val(IIf(!性别 = "女", imgRPT.ListImages("女人").Index, imgRPT.ListImages("男人").Index)) - 1
                End If
                
                '路径状态
                Set objItem = objRecord.AddItem(Val("" & !路径状态))
                objItem.Caption = " "
                objItem.Icon = Get临床路径序号(Val("" & !路径状态) + 2, False) - 1
                
                objRecord.AddItem Val(!病人ID)
                objRecord.AddItem Val(!主页ID)
                objRecord.AddItem CStr(Nvl(!姓名))
                Set objItem = objRecord.AddItem(CStr(Nvl(!住院号)))
                objItem.Caption = Nvl(!住院号, " ")
                Set objItem = objRecord.AddItem(LPAD(Nvl(!床号), 10, " "))
                objItem.Caption = CStr(Nvl(!床号, " "))
                Set objItem = objRecord.AddItem(CStr(Nvl(!性别, "男")))
                objItem.Caption = CStr(Nvl(!性别, "男"))
                Set objItem = objRecord.AddItem(Nvl(!年龄, "0"))
                objItem.Caption = Nvl(!年龄, "0")
                Set objItem = objRecord.AddItem(Nvl(!费别, ""))
                objItem.Caption = CStr(Nvl(!费别, ""))
                Set objItem = objRecord.AddItem(Nvl(!医疗付款方式, ""))
                objItem.Caption = CStr(Nvl(!医疗付款方式, ""))
                Set objItem = objRecord.AddItem(Nvl(!住院医师, ""))
                objItem.Caption = CStr(Nvl(!住院医师, ""))
                Set objItem = objRecord.AddItem(CStr(Format(!入院日期, "yyyy-MM-dd HH:mm:ss")))
                objItem.Caption = CStr(Format(!入院日期, "yyyy-MM-dd HH:mm:ss"))
                Set objItem = objRecord.AddItem(CStr(Format(!出院日期, "yyyy-MM-dd HH:mm:ss")))
                objItem.Caption = CStr(Format(!出院日期, "yyyy-MM-dd HH:mm:ss"))
                Set objItem = objRecord.AddItem(Nvl(!病人类型, "普通病人"))
                objItem.Caption = CStr(Nvl(!病人类型, "普通病人"))
                Set objItem = objRecord.AddItem(CStr(Nvl(!就诊卡号)))
                objItem.Caption = Nvl(!就诊卡号, "")
                '提取病人类型的颜色
                lngColor = 0
                mrsPatiColor.Filter = "名称='" & Nvl(!病人类型, "普通病人") & "'"
                If mrsPatiColor.RecordCount <> 0 Then
                    lngColor = Nvl(mrsPatiColor!颜色, 0)
                End If
                If lngColor <> 0 Then
                    objRecord.Item(c_姓名).ForeColor = lngColor
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    
    If intCurPage = 页面.待入科 Or intCurPage = -1 Then rptPati(页面.待入科).Populate '缺省不选中任何行
    If intCurPage = 页面.转科 Or intCurPage = -1 Then rptPati(页面.转科).Populate  '缺省不选中任何行
    If intCurPage = 页面.出院 Or intCurPage = -1 Then rptPati(页面.出院).Populate  '缺省不选中任何行
    If intCurPage = 页面.家庭病床 Or intCurPage = -1 Then rptPati(页面.家庭病床).Populate  '缺省不选中任何行
    
    UpgradeList = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UpgradeBeds(ByVal rsPati As ADODB.Recordset) As Boolean
    '更新在院病人的床位数据并显示
    Dim arrBeds
    Dim i As Integer, j As Integer, lngCardIndex As Integer
    Dim lngPatiColor As Long
    Dim strDiag As String
    Dim strBeds As String, strAmountSQL As String, strDurationSQL As String
    Dim strMonitor As String
    Dim strBalance As String, strNotes As String
    Dim rsBalance As New ADODB.Recordset
    Dim rsDiagnosis As New ADODB.Recordset
    '49535,刘鹏飞,2012-08-14,病人信息由字符串连接，变更为数组
    Dim ArrPatiInfo As Variant
    On Error GoTo ErrHand
    
    '提取监护仪涉及到的住院号清单
    If mclsWardMonitor.Enabled And InStr(GetInsidePrivs(p住院护士站), "护理监护") > 0 Then
        strMonitor = mclsWardMonitor.GetListPati
    End If
    
    '显示所有床位卡片(考虑到大数据量及并发,先将卡片显示出来)
    j = picPati.Count - 2
    For i = 1 To j
        picPati(i).Visible = True
    Next
    
    If Mid(mstrCardInfo, 2, 1) = "1" Then
        '提取本病区所有病人的实际余额
        '56960:刘鹏飞,2013-01-17,病人余额添加包含担保额
        If mblnCardBalance = True Then
            strAmountSQL = "(SELECT  NVL(SUM(NVL(担保额 ,0)),0)" & vbNewLine & _
                "   FROM 病人担保记录" & vbNewLine & _
                "   WHERE 病人ID = C.病人ID AND 主页ID =C.主页ID AND" & vbNewLine & _
                "   (到期时间 IS NULL OR 到期时间 > SYSDATE) AND 删除标志 = 1)+"
            
            strDurationSQL = ",(SELECT 1" & vbNewLine & _
                " FROM 病人担保记录" & vbNewLine & _
                " WHERE 病人ID = C.病人ID AND 主页ID = C.主页ID AND (到期时间 IS NOT NULL And 到期时间 > SYSDATE)" & vbNewLine & _
                " And 担保额 = 999999999 AND 删除标志 = 1 And RowNum < 2) 不限担保额"
        Else
            strAmountSQL = ""
            strDurationSQL = ",NULL 不限担保额"
        End If
        mstrSQL = "  Select D.病人ID,D.主页ID,D.住院号," & strAmountSQL & "NVL(A.预交余额,0)+NVL(B.医保报销,0)-NVL(A.费用余额,0) AS 余额" & strDurationSQL & _
                   " From 病人余额 A," & _
                   "      (Select B.病人ID,B.主页ID,SUM(B.金额) AS 医保报销" & _
                   "      From 保险模拟结算 B,结算方式 D,病人信息 A,在院病人 R" & _
                   "      Where B.结算方式=D.名称 And D.性质 IN (3,4) And B.病人ID=A.病人ID And B.主页ID=A.主页id And a.病人ID=R.病人ID And A.当前病区ID=R.病区ID  And R.病区ID=[1]" & _
                   "      GROUP BY B.病人ID,B.主页ID) B," & _
                   "      病案主页 C,病人信息 D,在院病人 R" & _
                   " Where A.病人ID(+) =C.病人ID AND A.性质(+)=1 AND A.类型(+)=2" & _
                   " And B.病人ID(+)=C.病人ID And B.主页ID(+)=C.主页ID" & _
                   " And D.病人ID=R.病人ID And D.病人ID=C.病人ID And D.主页id=C.主页ID And D.当前病区ID=R.病区ID And R.病区ID=[1]"
        Set rsBalance = zlDatabase.OpenSQLRecord(mstrSQL, "提取本病区所有病人的实际余额", cboUnit.ItemData(cboUnit.ListIndex))
    End If
    Call ShowGuage("提取本病区所有病人的实际余额", 50)
    'debug.print "...提取本病区所有病人的实际余额:" & Now
    
    If Mid(mstrCardInfo, 1, 1) = "1" Then
        '提取本病区所有病人的诊断主要诊断
        '诊断类型:
        '    1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;
        '    10-并发症;11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
        '记录来源:
        '    1-病历；2-入院登记；3-首页整理;4-病案
'        mstrSQL = " Select A.病人ID,A.主页ID,A.诊断类型,A.记录来源,A.诊断次序,A.疾病ID,A.诊断ID,A.诊断描述,A.是否未治,A.是否疑诊,A.备注" & _
'                  " From 病人诊断记录 A,病案主页 B,病人信息 C,在院病人 R" & _
'                  " Where a.诊断类型 In (1, 2, 3, 11, 12, 13) And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.病人ID=C.病人ID And C.主页id=B.主页ID And C.病人ID=R.病人ID And C.当前病区ID=R.病区ID " & _
'                  " And 诊断次序=1 And (R.病区ID=[1] Or b.婴儿病区ID=[1])" & _
'                  " Order by A.病人ID asc,A.记录来源 desc,A.诊断类型 desc"
'        Set rsDiagnosis = zlDatabase.OpenSQLRecord(mstrSQL, "提取本病区所有病人的诊断", cboUnit.ItemData(cboUnit.ListIndex))
        Set rsDiagnosis = GetPatiDiagnoseByDept(cboUnit.ItemData(cboUnit.ListIndex), 1)
    End If
    Call ShowGuage("提取本病区所有病人的主要诊断", 60)
    'debug.print "...提取本病区所有病人的主要诊断:" & Now
    
    '更新内存映射记录集
    mstrFields = "病况|护理等级|护理等级名称|病人类型|监护仪|病案审查|临床路径|个性标注1|病人状态|个性标注2|个性标注3|监护仪名称|病案审查名称|临床路径名称|个性标注1名称|病人状态名称|个性标注2名称|个性标注3名称|单病种"
    With rsPati
        .Filter = "类型 ='在院病人' Or 类型 ='预出院病人' Or 类型 ='预转科病人' Or 类型='转病区病人'"
        Do While Not .EOF
            '找到该病人的床号
            
            '82383:此处过滤主要是为了修正之前，同时开两个ZLHIS，将不同病人换到一张床的情况(保持和病人事物定位相同的病人)
            lngCardIndex = -1
            mrsBedInfo.Filter = "床号='" & Trim(Nvl(!床号, "ZYB")) & "'"
            If mrsBedInfo.RecordCount <> 0 Then
                If mrsBedInfo!病人ID = 0 Or mrsBedInfo!病人ID = !病人ID Then
                    lngCardIndex = mrsBedInfo!卡片索引
                End If
            End If
            If lngCardIndex = -1 Then
                mrsBedInfo.Filter = "病人ID=" & !病人ID
                If mrsBedInfo.RecordCount <> 0 Then
                    mrsBedInfo.Filter = "床号='" & Trim(Nvl(mrsBedInfo!床号, "ZYB")) & "'"
                    If mrsBedInfo.RecordCount <> 0 Then lngCardIndex = mrsBedInfo!卡片索引
                End If
            End If
            mrsBedInfo.Filter = 0
            
            If lngCardIndex <> -1 Then
                '准备更新病人卡片信息区域
                strBalance = ""
                If Mid(mstrCardInfo, 2, 1) = "1" Then
                    rsBalance.Filter = "病人ID=" & !病人ID
                    If rsBalance.RecordCount <> 0 Then
                        strBalance = Format(Nvl(rsBalance!余额, 0), "#0.00;-#0.00; ;")
                        If Val(Nvl(rsBalance!不限担保额, 0)) = 1 Then
                            strBalance = "不限额度担保"
                        End If
                    End If
                    rsBalance.Filter = 0
                End If
                
                '住院号,姓名,性别,年龄,诊断,医/护,费别,医疗付款方式,病况,入院日期,住院天数,余额,病人颜色,护理等级,就诊卡号）
                '56958:刘鹏飞,2013-01-16,医生和护士显示一行
                If Trim(Nvl(!住院医师)) = "" And Trim(Nvl(!责任护士)) = "" Then
                    strDiag = ""
                Else
                    strDiag = Trim(Nvl(!住院医师)) & "/" & Trim(Nvl(!责任护士))
                End If
                ArrPatiInfo = Array(IIf(mblnOutDept, Nvl(!门诊号), Nvl(!住院号)), Nvl(!姓名), Nvl(!性别), Nvl(!年龄), "[诊断]", strDiag, Nvl(!费别), Nvl(!医疗付款方式), _
                             IIf(Nvl(!当前病况) = "一般", "", Nvl(!当前病况)), Format(!入院日期, "yyyy-MM-dd"), Nvl(!住院天数), strBalance, 0, "", Nvl(!就诊卡号))
                '提取诊断(中医科中医诊断优先，然后诊断类型反序优先，然后诊断来源反序优先)
                strDiag = ""
                If Mid(mstrCardInfo, 1, 1) = "1" Then
                    rsDiagnosis.Filter = "病人ID=" & !病人ID
                    If rsDiagnosis.RecordCount <> 0 Then
                        strDiag = Nvl(rsDiagnosis!诊断描述)
                    End If
                    rsDiagnosis.Filter = 0
                End If
                ArrPatiInfo(4) = Replace(CStr(ArrPatiInfo(4)), "[诊断]", strDiag)
                '提取病人类型的颜色(为了避免颜色多了分散操作员注意力,黑色缺省不显示)
                mrsPatiColor.Filter = "名称='" & Nvl(!病人类型, "普通病人") & "'"
                If mrsPatiColor.RecordCount <> 0 Then
                    lngPatiColor = IIf(Nvl(!病人类型, "普通病人") = "普通病人", &HFFFFFF, Nvl(mrsPatiColor!颜色, 0))
                Else
                    lngPatiColor = &HFFFFFF
                End If
                mrsPatiColor.Filter = 0
                ArrPatiInfo(12) = lngPatiColor
                ArrPatiInfo(13) = Nvl(!护理等级, "三级护理")
                
                '1、更新卡片上的信息区域
                Call SetCardInfo(lngCardIndex, ArrPatiInfo)
                mstrValues = Nvl(!当前病况) & "|" & Get护理等级(Nvl(!护理等级, "三级护理")) & "|" & Nvl(!护理等级, "三级护理") & "|" & Nvl(!病人类型, "普通病人")
                
                '提取主题
                '2、更新卡片上的标注区域（监护仪|病案审查|临床路径|个性标注1|病人状态|个性标注2|个性标注3|护理等级）
                strNotes = UpgradeNotes(rsPati, strMonitor)
                mstrValues = mstrValues & strNotes
                Call Record_Update(mrsBedInfo, mstrFields & "|包床", mstrValues & "|0", "卡片索引|" & lngCardIndex)
                Call SetCardLabel(lngCardIndex)
                
                '3、更新包床
                strBeds = ""
                mrsBedInfo.Filter = "病人ID=" & !病人ID
                With mrsBedInfo
                    Do While Not .EOF
                        strBeds = strBeds & "," & !卡片索引 & "|" & !床号
                        .MoveNext
                    Loop
                End With
                mrsBedInfo.Filter = 0
                If strBeds <> "" Then strBeds = Mid(strBeds, 2)
                arrBeds = Split(strBeds, ",")
                j = UBound(arrBeds)
                For i = 0 To j
                    If Split(arrBeds(i), "|")(0) <> lngCardIndex Then
                        '住院号,姓名,性别,年龄,诊断,医/护,费别,医疗付款方式,病况,入院日期,住院天数,余额,病人颜色,护理等级,就诊卡号）
                        ArrPatiInfo = Array("", Nvl(rsPati!姓名), "包床", "", "", "", "", "", "", "", "", "", lngPatiColor, "", "")
                        Call SetCardInfo(Split(arrBeds(i), "|")(0), ArrPatiInfo)
                        
                        '更新包床的信息
                        Call Record_Update(mrsBedInfo, mstrFields & "|包床", mstrValues & "|1", "卡片索引|" & Split(arrBeds(i), "|")(0))
                    End If
                Next
            End If
            
            .MoveNext
        Loop
        rsPati.Filter = 0
    End With
    
    Call ShowGuage("完成病区床位卡内容更新", 80)
    'debug.print "...完成卡片内容更新:" & Now
    
    '同步刷新审查反馈信息
    Call LoadResponse
    UpgradeBeds = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    rsPati.Filter = 0
End Function

Private Function UpgradeNotes(ByVal rsPati As ADODB.Recordset, ByVal strMonitor As String) As String
    Dim int病案审查 As Integer, int临床路径 As Integer, int病人状态 As Integer, int监护仪 As Integer, str标注1 As String, str标注2 As String, str标注3 As String
    Dim str病人状态 As String, str个性标注1 As String, str个性标注2 As String, str个性标注3 As String, str单病种 As String
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    '获取当前病人的标注图形索引
    '61824:刘鹏飞,2013-05-23,显示单病种标志
    str单病种 = Nvl(rsPati!单病种)
    int病案审查 = Nvl(rsPati!病案状态, 0)
    int临床路径 = rsPati!路径状态 + 2
    If rsPati!排序 = "3.2" Or rsPati!排序 = "3.3" Then     '预转科
        str病人状态 = "预转科"
        int病人状态 = Img标记(mlngSource).ListImages("预转科").Index
    ElseIf rsPati!排序 = pt预出 Then     '预出院
        str病人状态 = "预出院"
        int病人状态 = Img标记(mlngSource).ListImages("预出院").Index
    End If
    If strMonitor <> "" And Not IsNull(rsPati!住院号) Then
        If InStr("," & strMonitor & ",", "," & rsPati!住院号 & ",") > 0 Then
            int监护仪 = 1
        End If
    End If
    
    '图形索引+1是因为标注程序是从0开始
    mrsPatiNotes.Filter = "病人ID=" & rsPati!病人ID & " And 主页ID=" & rsPati!主页ID
    mrsPatiNotes.Sort = "标记顺序"
    Do While Not mrsPatiNotes.EOF
        i = Val("" & mrsPatiNotes!标记顺序)
        If i = 1 Then
            str标注1 = mrsPatiNotes!主题病区ID & "," & mrsPatiNotes!主题序号 & "," & mrsPatiNotes!标记序号 & "," & mrsPatiNotes!图形索引 + 1
        ElseIf i = 2 Then
            str标注2 = mrsPatiNotes!主题病区ID & "," & mrsPatiNotes!主题序号 & "," & mrsPatiNotes!标记序号 & "," & mrsPatiNotes!图形索引 + 1
        ElseIf i = 3 Then
            str标注3 = mrsPatiNotes!主题病区ID & "," & mrsPatiNotes!主题序号 & "," & mrsPatiNotes!标记序号 & "," & mrsPatiNotes!图形索引 + 1
        End If
        mrsNotes.Filter = "病区ID=" & mrsPatiNotes!主题病区ID & " And 主题序号=" & mrsPatiNotes!主题序号 & " And 标记序号=" & mrsPatiNotes!标记序号
        If mrsNotes.RecordCount <> 0 Then
            str个性标注1 = mrsNotes!说明
            If i = 1 Then
                str个性标注1 = mrsNotes!说明
            ElseIf i = 2 Then
                str个性标注2 = mrsNotes!说明
            ElseIf i = 3 Then
                str个性标注3 = mrsNotes!说明
            End If
        End If
        mrsPatiNotes.MoveNext
    Loop

    mrsPatiNotes.Filter = ""
    mrsNotes.Filter = ""

    UpgradeNotes = "|" & int监护仪 & "|" & int病案审查 & "|" & int临床路径 & "|" & str标注1 & "|" & int病人状态 & "|" & str标注2 & "|" & str标注3 & "|" & _
                   IIf(int监护仪 > 0, "监护仪", "") & "|" & Get病案主题(int病案审查) & "|" & Get临床路径主题(int临床路径) & "|" & str个性标注1 & "|" & str病人状态 & "|" & str个性标注2 & "|" & str个性标注3 & "|" & str单病种
End Function

Private Function Get临床路径序号(ByVal lng状态 As Long, Optional ByVal blnCard As Boolean = True) As Long
    Dim imgList As ImageList
    If blnCard = True Then
        Set imgList = Img标记(mlngSource)
    Else
        Set imgList = imgRPT
    End If
    Get临床路径序号 = Choose(lng状态, imgList.ListImages("未导入").Index, imgList.ListImages("不符合").Index, _
            imgList.ListImages("执行中").Index, imgList.ListImages("正常结束").Index, imgList.ListImages("变异结束").Index)
End Function

Private Function Get临床路径主题(ByVal lng状态 As Long) As String
    Get临床路径主题 = Choose(lng状态, "未导入", "不符合", "执行中", "正常结束", "变异结束")
End Function

Private Function Get病案图标序号(ByVal lng状态 As Long, Optional ByVal blnCard As Boolean = True) As Long
    Dim i As Long
    Dim imgList As ImageList
    
    If blnCard = True Then
        Set imgList = Img标记(mlngSource)
    Else
        Set imgList = imgRPT
    End If
    Select Case lng状态
        Case 1
            i = imgList.ListImages("等待审查").Index
        Case 2
            i = imgList.ListImages("拒绝审查").Index
        Case 13
            i = imgList.ListImages("正在抽查").Index
        Case 3
            i = imgList.ListImages("正在审查").Index
        Case 14
            i = imgList.ListImages("抽查反馈").Index
        Case 4
            i = imgList.ListImages("审查反馈").Index
        Case 16
            i = imgList.ListImages("抽查整改").Index
        Case 6
            i = imgList.ListImages("审查整改").Index
        Case 5
            i = imgList.ListImages("审查归档").Index
        Case 10
            i = imgList.ListImages("等待审查").Index
    End Select
    Get病案图标序号 = i
End Function

Private Function Get病案主题(ByVal lng状态 As Long) As String
    Dim i As Long
    
    Select Case lng状态
        Case 1
            Get病案主题 = "等待审查" '提交接收
        Case 2
            Get病案主题 = "拒绝审查" '拒绝接受
        Case 13
            Get病案主题 = "正在抽查"
        Case 3
            Get病案主题 = "正在审查"
        Case 14
            Get病案主题 = "抽查反馈"
        Case 4
            Get病案主题 = "审查反馈"
        Case 16
            Get病案主题 = "抽查整改"
        Case 6
            Get病案主题 = "审查整改"
        Case 10
            Get病案主题 = "接收待审"
        Case 5
            Get病案主题 = "审查归档"
    End Select
End Function

Private Function GetVersion() As String
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    strSQL = " select 版本号 from zlsystems where 编号=100"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取标准版本号")
    GetVersion = rsTemp!版本号
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadPatients(ByRef rsPati As ADODB.Recordset) As Boolean
'功能：读取病人列表
    Dim strSQL As String
    Dim int入院天数 As Integer, strPatiFileter As String
    '修改此SQL的条件,病人事务管理模块也需要调整
    '61824:刘鹏飞,2013-05-23,显示单病种标志
    '转科待入科病人
    
    '当页面下拉框清空，F5刷新，应该恢复上一个的值
    If cboUnit.ListIndex = -1 Then Call zlControl.CboSetIndex(cboUnit.hwnd, mintPreDept)
    
    '111016:入院待入科病人过滤,为0表示不控制
    int入院天数 = zlDatabase.GetPara("入院天数", glngSys, mlngModul, 0)
    If int入院天数 > 0 Then
        strPatiFileter = " And B.入院日期>=Sysdate-[2]"
    End If
    
    If Val(Mid(mstrScope, 5, 1)) <> 0 Then
        '84938:刘鹏飞,待入科病人提取优化(添加A.主页ID=B.主页ID)
        strSQL = _
            "Select /*+ RULE */Distinct" & vbNewLine & _
            " Decode(B.状态,1,0,Decode(c.开始原因,3,1,2)) As 排序, Decode(Nvl(b.病案状态, 0), 0, 999, b.病案状态) As 排序2," & _
            " Decode(B.状态,1,'入院待入住病人',Decode(c.开始原因,3,'转科待入住病人','转病区待入住病人')) As 类型," & _
            " a.病人id, b.主页id, A.门诊号,B.住院号, NVL(B.姓名,A.姓名) 姓名" & mstrBriefCode & ", NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄," & vbNewLine & _
            " d.名称 As 科室, c.科室id, c.经治医师 As 住院医师,b.责任护士, b.病案状态, c.床号," & _
            " e.名称 As 护理等级, b.费别,B.医疗付款方式,b.当前病况, Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期, b.出院日期,B.出院方式, b.病人类型, b.状态, b.险类, a.就诊卡号," & vbNewLine & _
            " -1 As 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID,A.主页Id 最大主页Id" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病人变动记录 C, 部门表 D, 收费项目目录 E,病人类型 Z" & vbNewLine & _
            "Where a.在院 = 1 And B.病人类型=Z.名称(+) And a.病人id = b.病人id And A.主页ID=B.主页ID And Nvl(b.主页id, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id " & vbNewLine & _
            "      And (C.病区ID=[1] or C.病区ID is null) And c.科室id = d.Id" & vbNewLine & _
            "      And (d.站点='" & gstrNodeNo & "' Or d.站点 is Null)" & vbNewLine & _
            "      And b.护理等级id = e.Id(+) And Nvl(c.附加床位, 0) = 0 And c.终止时间 Is Null" & vbNewLine & _
            "      And (c.开始原因 in(1,3) And Exists(Select 1 From 病区科室对应 H Where c.科室id = h.科室id And h.病区id = [1]) or c.开始原因=15 And c.病区id = [1])" & vbNewLine & _
            "      And ((c.开始原因 = 1 And b.状态 = 1 " & strPatiFileter & ") Or (c.开始原因 in (3,15) And c.开始时间 Is Null And b.状态 = 2)) "
    
    End If
    '在院病人（床位一览表的模式，必须显示在院病人）
    strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
        "Select /*+ RULE */ Decode(B.状态,3,4,DECODE(B.出院病床, NULL, 3.1,DECODE(B.状态,2,3.2,3))) as 排序," & _
        " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
        " Decode(B.状态,3,'预出院病人',DECODE(B.出院病床, NULL, '家庭病床',DECODE(B.状态,2,'预转科病人', '在院病人'))) as 类型," & _
        " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名" & mstrBriefCode & ",NVL(b.性别,a.性别) 性别,NVL(b.年龄,a.年龄) 年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
        " B.出院病床 as 床号,E.名称 as 护理等级,B.费别,B.医疗付款方式,B.当前病况,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期,B.出院日期,B.出院方式,B.病人类型," & _
        " B.状态,B.险类,A.就诊卡号,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID,A.主页Id 最大主页Id" & _
        " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z,在院病人 R" & _
        " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And A.主页ID=B.主页ID And Nvl(B.状态,0)<>1" & _
        " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And R.病区ID=[1] And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
        " And a.病人ID=R.病人ID And A.当前病区ID=R.病区ID And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
    strSQL = strSQL & " Order by 排序,排序2,床号,主页ID Desc"
    
    On Error GoTo errH
    Set rsPati = New ADODB.Recordset
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), int入院天数)
    
    rsPati.Filter = "类型='预出院病人'"
    mlng预出院 = rsPati.RecordCount
    rsPati.Filter = 0
    
    LoadPatients = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdjustCard(Optional ByVal lngY As Long = clngX, Optional ByVal strKeys As String = "")
    'strKeys不为空则直接根据病人过滤，说明是公告栏过滤
    Dim i As Integer, j As Integer
    Dim blnAdjust As Boolean
    Dim lngX As Long, lngRowCount As Long, lngShowed As Long
    Dim lng病人ID As Long, lngIndex As Long
    Dim blnShowCard As Boolean
    '只有切换病区的时候才重新读取数据,病区内的条件变化,只是将卡片隐藏后重新设置位置即可
    
    '刷新子窗口菜单
    Call LockWindowUpdate(Me.hwnd)
    
    '隐藏所有床位卡片
    mintCards = 0
    lng病人ID = mlng病人ID
    mlng病人ID = 0
    mstrBoardKeys = strKeys
    j = picPati.Count - 2
    For i = 1 To j
        picPati(i).Visible = False
    Next
    
    If j = 0 Then Exit Sub
    blnAdjust = (lngY = clngX)
    '重新进行坐标计算
    lngX = clngX
    lngRowCount = (PicDraw.Width - HScr.Width - 50) \ (picPati(mlngSource).Width + 15)
    PicDraw.Refresh
    
    lngIndex = -1
    With mrsBedInfo
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If strKeys = "" Then
                blnShowCard = ISShowCard
            Else
                blnShowCard = (InStr(1, "," & strKeys & ",", "," & Nvl(mrsBedInfo!病人ID) & ",") <> 0)
            End If
            If blnShowCard Then
                If !病人ID = lng病人ID And lng病人ID <> 0 Then
                    lngIndex = !卡片索引
                End If
                lngShowed = lngShowed + 1
                With picPati(!卡片索引)
                    .Left = lngX
                    .Top = lngY
                    .Width = picPati(mlngSource).Width
'                    If mblnCardCollapse Then
'                        .Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
'                    ElseIf mblnShowCard = True Then
'                        .Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
'                    Else
'                        .Height = IIf(mlngSource = 0, clngBigHeight_Normal, clngBaseHeight_Normal)
'                    End If
                    If mblnCardCollapse Then
                        .Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
                    Else
                        .Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
                    End If
                    .Visible = True
                    '.ZOrder 0
                End With
                
                '计算下一张卡片的坐标
                lngX = lngX + picPati(mlngSource).Width ' + 30
                If lngShowed Mod lngRowCount = 0 Then
                    lngX = clngX
                    lngY = lngY + picPati(mlngSource).Height ' + 30
                End If
            End If
            .MoveNext
        Loop
    End With
    picList.ZOrder 0
    PatiPage.ZOrder 0
    fraPatiUD.ZOrder 0
    picPara(2).ZOrder 0
    picPara(3).ZOrder 0
    pic出院查找.ZOrder 0
    
    If blnAdjust Then
        mdblScaleHeight = (lngY + picPati(mlngSource).Height) ' + 30)
        mblnHScroll = (mdblScaleHeight > PicDraw.Height - IIf(picList.Visible, picList.Height, 0))
        With HScr
            .Value = 0
            .Top = PicDraw.Top
            .Left = PicDraw.Width - .Width
            .Height = PicDraw.Height
            .Visible = mblnHScroll
        End With
    End If
    
    If lngIndex <> -1 Then
        If mlngSelect <> lngIndex Then
            mlngSelect = lngIndex
            Call ShowSelect
        Else
            mlng病人ID = lng病人ID
        End If
    End If

    '刷新子窗口菜单
    Call LockWindowUpdate(0)
End Sub

Private Function ISShowCard() As Boolean
    Dim arr护理
    Dim strInfo As String, int入科天数 As Integer
    Dim i As Integer, j As Integer
    Dim arrSignNotes(0 To 2) As String, arrNote(0 To 2) As String
    
    '判断当前卡片是否符合条件
    int入科天数 = zlDatabase.GetPara("入科天数", glngSys, mlngModul, 0)
    ISShowCard = (chk包含空床.Value = 1 Or Not (chk包含空床.Value = 0 And Nvl(mrsBedInfo!病人ID, 0) = 0))
    If ISShowCard Then
        '病况过滤
        Select Case Nvl(mrsBedInfo!病况)
        Case "危"
            ISShowCard = (chk病况条件(1).Value = 1)
        Case "重"
            ISShowCard = (chk病况条件(2).Value = 1)
        Case Else
            ISShowCard = (chk病况条件(0).Value = 1)
        End Select
    End If
    If ISShowCard And cbo床位状况.Text <> "全部" Then
        '根据护理等级的名称来判断
        ISShowCard = (mrsBedInfo!床位编制 = cbo床位状况.Text)
    End If
    If ISShowCard And txt护理条件.Text <> "全部" Then
        '根据护理等级的名称来判断
        ISShowCard = (InStr(1, "," & txt护理条件.Text & ",", "," & mrsBedInfo!护理等级名称 & ",") <> 0)
    End If
    If ISShowCard Then
        '主题过滤
        If Me.cbo内容.Text <> "所有" Then strInfo = cbo内容.Text
        If Me.cbo主题.Text <> "所有" Then
            Select Case Me.cbo主题.ListIndex
            Case 1
                If Me.cbo内容.Text = "所有" Then
                    ISShowCard = (mrsBedInfo!病案审查 <> 0)
                Else
                    ISShowCard = (Nvl(mrsBedInfo!病案审查名称) = strInfo)
                End If
            Case 2
                If Me.cbo内容.Text = "所有" Then
                    ISShowCard = (mrsBedInfo!临床路径 <> 0)
                Else
                    ISShowCard = (Nvl(mrsBedInfo!临床路径名称) = strInfo)
                End If
            Case 3
                On Error GoTo errH
                If Me.cbo内容.Text = "所有" Then
                    ISShowCard = (mrsBedInfo!病人状态 <> 0)
                    If Not ISShowCard Then
                        If mrsBedInfo!病人ID <> 0 Then
                            If Not IsNull(mrsBedInfo!住院天数) Then
                                ISShowCard = (Val(mrsBedInfo!住院天数) <= int入科天数)
                            Else
                                ISShowCard = False
                            End If
                        Else
                            ISShowCard = False
                        End If
                    End If
                ElseIf Me.cbo内容.Text Like "入科*天内" Then
                    If mrsBedInfo!病人ID <> 0 Then
                        If Not IsNull(mrsBedInfo!住院天数) Then
                            ISShowCard = (Val(mrsBedInfo!住院天数) <= int入科天数)
                        Else
                            ISShowCard = False
                        End If
                    Else
                        ISShowCard = False
                    End If
                Else
                    ISShowCard = (Nvl(mrsBedInfo!病人状态名称) = strInfo)
                End If
            Case Is > 3 '个性标注
                ISShowCard = False
                If Nvl(mrsBedInfo!个性标注1) <> "" Then
                    arrSignNotes(0) = Split(mrsBedInfo!个性标注1, ",")(0) & "," & Split(mrsBedInfo!个性标注1, ",")(1)
                    arrNote(0) = Split(mrsBedInfo!个性标注1, ",")(2)
                End If
                If Nvl(mrsBedInfo!个性标注2) <> "" Then
                    arrSignNotes(1) = Split(mrsBedInfo!个性标注2, ",")(0) & "," & Split(mrsBedInfo!个性标注2, ",")(1)
                    arrNote(1) = Split(mrsBedInfo!个性标注2, ",")(2)
                End If
                If Nvl(mrsBedInfo!个性标注3) <> "" Then
                    arrSignNotes(2) = Split(mrsBedInfo!个性标注3, ",")(0) & "," & Split(mrsBedInfo!个性标注3, ",")(1)
                    arrNote(2) = Split(mrsBedInfo!个性标注3, ",")(2)
                End If
                
                If Me.cbo内容.Text = "所有" Then
                    mrsNotes.Filter = "标记序号=0"
                Else
                    mrsNotes.Filter = "标记序号>0"
                End If
                mrsNotes.Sort = "病区ID,主题序号"
                Do While Not mrsNotes.EOF
                    If Val(mrsNotes!病区ID) + Val(mrsNotes!主题序号) = Val(cbo主题.ItemData(cbo主题.ListIndex)) Then
                        For i = 0 To UBound(arrSignNotes)
                            If arrSignNotes(i) = mrsNotes!病区ID & "," & mrsNotes!主题序号 Then
                                If Me.cbo内容.Text = "所有" Then
                                    ISShowCard = True
                                Else
                                    If Val(arrNote(i)) = Val(cbo内容.ItemData(cbo内容.ListIndex)) Then
                                        ISShowCard = True
                                    End If
                                End If
                                Exit For
                            End If
                        Next
                        Exit Do
                    End If
                mrsNotes.MoveNext
                Loop
                
            End Select
        End If
    End If
    
    If ISShowCard Then mintCards = mintCards + 1
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitComponent()
    Set mclsAdvices = New zlPublicAdvice.clsDockInAdvices
    If Not mobjPlugIn Is Nothing Then Call mclsAdvices.zlInitPlugIn(mobjPlugIn)
    
    Set mclsFeeQuery = New zl9InExse.clsFeeQuery
    Call mclsFeeQuery.InitCallByNurse(gfrmMain, gcnOracle, gstrDBUser, glngSys)
        
    Set mclsInPatient = New zl9InPatient.clsInPatient
    Call mclsInPatient.InitCallByNurse(gfrmMain, gcnOracle, gstrDBUser, glngSys)
    
    Set mclsTends = New zl9TendFile.clsTendFile
    Call mclsTends.InitTendFile(gcnOracle, glngSys)
    Set mclsWardMonitor = New clsWardMonitor

    '保存各对象窗体
    Set mcolSubForm = New Collection
    mcolSubForm.Add mclsAdvices.zlGetForm, "_医嘱"
    mcolSubForm.Add mclsFeeQuery.zlGetForm, "_费用"
    If mclsWardMonitor.Enabled Then
        mcolSubForm.Add mclsWardMonitor.zlGetForm, "_监护"
    End If
End Sub

Private Sub AddSendCommandBar()
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim strPrivs As String, strPara As String
    Dim strUnit As String
    Dim i As Long
    '61762:刘鹏飞,2013-05-20,增加发送输液药品医嘱的功能
    If gstr输液配置中心 <> "" Then
        strUnit = cboUnit.ItemData(cboUnit.ListIndex)
        strPrivs = GetInsidePrivs(p住院医嘱发送)
        If InStr(";" & strPrivs & ";", ";发送药疗临嘱;") = 0 Or InStr(";" & strPrivs & ";", ";发送药疗长嘱;") = 0 Then
            strPrivs = ""
        End If
    End If
    
    strPara = zlDatabase.GetPara("在输液配置中心发药的病人科室", glngSys, p住院医嘱下达)
    If strPara = "*" Then strUnit = "*"
    '一、病区批量工作医嘱发送菜单添加
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(3)
    '删除发送医嘱按钮
    For i = cbrMenuBar.CommandBar.Controls.Count To 1 Step -1
        If cbrMenuBar.CommandBar.Controls(i).ID = conMenu_Edit_Send Then
            cbrMenuBar.CommandBar.Controls(i).Delete
        End If
    Next i
    '添加医嘱按钮
    With cbrMenuBar.CommandBar.Controls
        '先找到发送之前的校对按钮
        Set cbrControl = .Find(, conMenu_Edit_Audit)
        If Not cbrControl Is Nothing Then
            If strPrivs <> "" Then
                Set cbrMenuBar = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "医嘱发送(&4)", cbrControl.Index + 1)
                cbrMenuBar.CommandBar.Title = "病区批量工作"
                cbrMenuBar.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Send, "发送所有医嘱(&S)"
                If InStr(1, "," & strPara & ",", "," & strUnit & ",") > 0 Then
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送输液药品(&I)")
                Else
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送静脉营养药品(&I)")
                End If
                cbrControl.IconId = conMenu_Edit_Send
            Else
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "医嘱发送(&4)", cbrControl.Index + 1): cbrControl.ToolTipText = ""
            End If
        End If
    End With
    
    '二、单个病人医嘱业务发送按钮添加
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(4)
    '删除发送医嘱按钮
    For i = cbrMenuBar.CommandBar.Controls.Count To 1 Step -1
        If cbrMenuBar.CommandBar.Controls(i).ID = conMenu_Edit_Send Then
            cbrMenuBar.CommandBar.Controls(i).Delete
        End If
    Next i
    '添加医嘱发送按钮
    With cbrMenuBar.CommandBar.Controls
        '先找到发送之前的校对按钮
        Set cbrControl = .Find(, conMenu_Edit_Price)
        If Not cbrControl Is Nothing Then
            If strPrivs <> "" Then
                Set cbrMenuBar = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "发送(&G)", cbrControl.Index + 1)
                cbrMenuBar.CommandBar.Title = "医嘱业务"
                cbrMenuBar.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Send, "发送所有医嘱(&1)"
                If InStr(1, "," & strPara & ",", "," & strUnit & ",") > 0 Then
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送输液药品(&2)")
                Else
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送静脉营养药品(&2)")
                End If
                cbrControl.IconId = conMenu_Edit_Send
            Else
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送(&G)", cbrControl.Index + 1)
            End If
        End If
    End With
    '三、工具栏医嘱发送菜单添加
    '删除发送医嘱按钮
    For i = cbsMain(2).Controls.Count To 1 Step -1
        If cbsMain(2).Controls(i).ID = conMenu_Edit_Send Then
            cbsMain(2).Controls(i).Delete
        End If
    Next i
    
    '添加医嘱发送按钮
    With cbsMain(2).Controls
        '先找到发送之前的校对按钮
        Set cbrControl = .Find(, conMenu_Edit_Audit)
        If Not cbrControl Is Nothing Then
            If strPrivs <> "" Then
                Set cbrMenuBar = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "发送", cbrControl.Index + 1): cbrMenuBar.Style = xtpButtonIconAndCaption
                cbrMenuBar.CommandBar.Title = "病区批量工作"
                cbrMenuBar.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Send, "发送所有医嘱(&S)"
                If InStr(1, "," & strPara & ",", "," & strUnit & ",") > 0 Then
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送输液药品(&I)")
                Else
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送静脉营养药品(&I)")
                End If
                cbrControl.IconId = conMenu_Edit_Send
            Else
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送", cbrControl.Index + 1): cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "医嘱发送"
            End If
        End If
    End With
    
    cbsMain.RecalcLayout
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup, objFile As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    
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
    cbsMain.Icons = imgPublic.Icons
    
    '菜单定义
    '-----------------------------------------------------
    Set objFile = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False) '固有
    objFile.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objFile.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintBedCard, "打印床头卡(&K)…")  '打印床头卡
        '49854:刘鹏飞,2013-10-31,病人腕带打印
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print_Label, "打印腕带(&W)…")  '打印腕带
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "打印一日清单(&D)…", 1)
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintPageSet, "打印帐页(&Z)…", 1)
        objControl.Parameter = "100,ZL1_INSIDE_1139_2"
        objControl.IconId = conMenu_ReportPopup * 100#      '取第一个菜单项的图标
        Set objControl = .Add(xtpControlButton, conMenu_ReportPopup * 100# + 91, "住院科室日报(&R)…", 1)
        objControl.Parameter = "100,ZL1_INSIDE_1132"
        objControl.IconId = conMenu_ReportPopup * 100#      '取第一个菜单项的图标

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_File_MedRec, "首页打印(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_File_MedRecSetup, "打印设置(&S)", -1, False
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "打印预览(&V)", -1, False
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "打印首页(&P)", -1, False
        End With

        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set mobjPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "病人入出(&P)", -1, False) '固有
    mobjPopup.ID = conMenu_ManagePopup '对xtpControlPopup类型的命令ID需重新赋值
    mobjPopup.CommandBar.Title = "病人入出"
    With mobjPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_In, "入住(&I)"): objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Turn, "转科(&C)"): objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TurnUnit, "转病区(&D)"): objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TurnTeam, "转小组(&T)"): objControl.Category = "病人"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Bed, "换床(&B)"): objControl.BeginGroup = True: objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TransposeBed, "床位对换(&Q)"): objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_House, "包房(&H)"): objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_BedGrid, "更改床位等级(&G)"): objControl.Category = "病人"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_PatiInfo, "调整住院信息(&P)"): objControl.BeginGroup = True: objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_PaitNote, "病人备注信息(&F)"): objControl.Category = "病人"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Out, "出院(&O)"): objControl.BeginGroup = True: objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InPati, "转为住院病人(&Z)"): objControl.Category = "病人"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Baby, "新生儿登记(&N)"): objControl.BeginGroup = True: objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_ReCalcFee, "按费别重算费用(&R)"): objControl.Category = "病人"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InsureSel, "医保病种选择(&M)"): objControl.Category = "病人"
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Change_Undo, "撤销(&U)"): objPopup.BeginGroup = True: objPopup.Category = "病人"
        objPopup.IconId = conMenu_Edit_Untread
        
        '监护仪
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Monitor, "监护仪(&N)")
        objControl.BeginGroup = True
        objControl.Category = "病人"
    End With

    Set mobjPopupBatch = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "病区批量工作(&B)", -1, False)  '固有
    mobjPopupBatch.ID = conMenu_ManagePopup '对xtpControlPopup类型的命令ID需重新赋值
    mobjPopupBatch.CommandBar.Title = "病区批量工作"
    With mobjPopupBatch.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "预结算(&1)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintMultiBill, "催款(&2)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "医嘱校对(&3)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "医嘱发送(&4)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "医嘱暂停(&5)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "医嘱启用(&6)"): objControl.ToolTipText = ""
        '67386:刘鹏飞,2013-12-20,添加批量医嘱确认停止、医嘱批量核对功能
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "确认停止(&7)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Report_Reports, "打印执行单(&8)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Report_DrugQuery, "摆药查询(&9)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "药品留存登记(&J)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "超期发送收回(&S)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "医嘱批量执行(&B)"): objControl.IconId = 3587: objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "医嘱批量核对(&T)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AnimalHeat, "批量录入体温单(&A)"): objControl.BeginGroup = True: objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NurseLogFile, "批量录入记录单(&L)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_病人事务处理, "病人事务处理(&0)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_ProveCollect, "检验采集工作站(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatUnPack, "批量打包(&U)"): objControl.BeginGroup = True: objControl.IconId = 3051
        If gbln启用影像信息系统预约 = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrintBat, "批量打印预约单(&R)"): objControl.BeginGroup = True: objControl.IconId = 103
        End If
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "医嘱业务(&A)", -1, False)   '固有：医嘱A，费用F，病历E，护理L
    objMenu.ID = conMenu_EditPopup '对xtpControlPopup类型的命令ID需重新赋值
    objMenu.CommandBar.Title = "医嘱业务"
    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_查看医嘱, "查看医嘱(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新开(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "校对(&J)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Price, "计价(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送(&G)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停止(&S)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "确认停止(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "作废(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "暂停(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "浏览检验结果(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Notify, "刷新提醒(&N)"): objControl.BeginGroup = True
    End With
    
    '63608:刘鹏飞,2013-07-22,修改费用业务的快捷键为C
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "费用业务(&C)", -1, False) '固有：医嘱A，费用C，病历E，护理L
    objMenu.ID = conMenu_EditPopup '对xtpControlPopup类型的命令ID需重新赋值
    objMenu.CommandBar.Title = "费用业务"
    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_查看费用, "查看费用(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Billing, "记帐(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Balance, "结帐(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBillingApply, "销帐申请(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBillingAudit, "销帐审核(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "预结算(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_ReCalcFee, "按费别重算费用(&R)")
    End With
'
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "护理业务(&L)", -1, False) '固有：医嘱A，费用F，病历E，护理L
'    objMenu.ID = conMenu_EditPopup '对xtpControlPopup类型的命令ID需重新赋值
'    objMenu.CommandBar.Title = "护理业务"
'    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_查看体温单, "查看体温单(&T)")
'        Set objControl = .Add(xtpControlButton, conMenu_查看护理记录, "查看护理记录单(&H)")
'        Set objControl = .Add(xtpControlButton, conMenu_查看护理病历, "查看护理病历(&B)")
'    End With
'
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "病历业务(&E)", -1, False) '固有：医嘱A，费用F，病历E，护理L
'    objMenu.ID = conMenu_EditPopup '对xtpControlPopup类型的命令ID需重新赋值
'    objMenu.CommandBar.Title = "病历业务"
'    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_查看病历, "查看病历(&E)")
'    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)  '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "字体大小(&N)") '固有
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_FontSize_S, "小字体(&S)", -1, False '固有(小字体对应小卡片，大字体对应大卡片)
            .Add xtpControlButton, conMenu_View_FontSize_L, "大字体(&L)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "卡片折叠(&C)") '固有

        Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "非在床病人"): objControl.BeginGroup = True '固有
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "显示房间号"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_View_NoticBoard, "公告栏"): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "查找方式(&Y)"): objPopup.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
        '53132:刘鹏飞,2013-11-08,添加病人担保信息查看
        Set objControl = .Add(xtpControlButton, conMenu_View_Warrant, "担保信息查阅(&W)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "资料参考(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "疾病诊断参考(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "诊疗措施参考(&C)", -1, False
        End With
        '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "首页整理(&M)")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditResponse, "审查反馈(&S)")
            objControl.BeginGroup = True
            objControl.ToolTipText = "处理或查看病案审查反馈"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "诊疗项目费用设置(&C)")
            objControl.BeginGroup = True
        'Set objControl = .Add(xtpControlButton, conMenu_Tool_UnitSubject, "病区标记设置(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_UnitNBoard, "病区公告栏设置(&B)")
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
    cbsMain(1).EnableDocking xtpFlagHideWrap

    '工具栏定义:病区批量性工作
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("病区事务工具栏", xtpBarTop)      '固有
    objBar.Title = "病区批量工作"
    objBar.EnableDocking xtpFlagStretched
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "预结"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "批量预结"
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintMultiBill, "催款"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "病区催款"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "校对"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "医嘱校对": objControl.BeginGroup = True
        '59098:刘鹏飞,2013-04-18,医嘱发送、暂停、启用提示信息错误和菜单ID错误
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "医嘱发送"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "暂停"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "医嘱暂停": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "医嘱启用"
        '67386:刘鹏飞,2013-12-20,添加批量医嘱确认停止、医嘱批量核对功能
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "确认停止"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "确认停止": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Report_Reports, "执行单"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "打印执行单": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Report_DrugQuery, "摆药"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "摆药查询"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "留存"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "留存登记"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "超期收回"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "超期发送收回"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "执行登记"): objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 3587: objControl.ToolTipText = "医嘱批量执行登记"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "核对"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "医嘱批量执行核对"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AnimalHeat, "体温单"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "批量录入体温单": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NurseLogFile, "记录单"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "批量录入记录单"
        Set objControl = .Add(xtpControlButton, conMenu_病人事务处理, "病人事务"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "病人事务处理": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_ProveCollect, "检验采集"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "检验采集工作站": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatUnPack, "打包"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "批量打包": objControl.BeginGroup = True: objControl.IconId = 3051
        If gbln启用影像信息系统预约 = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrintBat, "预约单"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "批量打印预约单": objControl.BeginGroup = True: objControl.IconId = 103
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "退出": objControl.BeginGroup = True
    End With
    '特殊处理
    '-----------------------------------------------------
    '工具栏右侧病区下拉框选择
    With objBar.Controls
        Set objControl = .Add(xtpControlLabel, 99999901, "病区")
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, 99999901, "病区")
        objCustom.Handle = Me.cboUnit.hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    
    '工具栏定义:过滤条件
    '-----------------------------------------------------
    Set mobjFilter = cbsMain.Add("过滤工具栏", xtpBarTop)   '固有
    mobjFilter.EnableDocking xtpFlagStretched
    mobjFilter.ContextMenuPresent = False
    With mobjFilter.Controls
        Set objControl = .Add(xtpControlLabel, 1, "护理等级")
        Set objCustom = .Add(xtpControlCustom, 2, "")
        objCustom.Handle = pic护理等级.hwnd
        Set objControl = .Add(xtpControlLabel, 3, "床位状况"): objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, 4, "")
        objCustom.Handle = pic床位状况.hwnd
        Set objControl = .Add(xtpControlLabel, 5, "当前病况"): objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, 6, "")
        objCustom.Handle = pic病况.hwnd
        Set objCustom = .Add(xtpControlCustom, 7, ""): objCustom.BeginGroup = True
        objCustom.Handle = pic主题过滤.hwnd
        Set objCustom = .Add(xtpControlCustom, 8, "")
        objCustom.Handle = chk包含空床.hwnd: objCustom.BeginGroup = True
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "↓按床号查找")
        objPopup.Caption = "↓按床号查找"
        objPopup.ID = conMenu_View_FindType
        objPopup.Style = xtpButtonCaption
        objPopup.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, 10, "")
        objCustom.Handle = txtFind.hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
'        .Add 0, vbKeyF1, conMenu_Edit_Audit         '医嘱校对
'        .Add 0, vbKeyF2, conMenu_Edit_Send          '医嘱发送
'        .Add 0, vbKeyF3, conMenu_Report_Reports     '打印执行单
'        .Add 0, vbKeyF4, conMenu_Report_DrugQuery   '摆药查询
'        .Add 0, vbKeyF6, conMenu_Edit_PreBalance    '预结算
'        .Add 0, vbKeyF7, conMenu_File_PrintMultiBill '催款
'        .Add 0, vbKeyF8, conMenu_Edit_BatExecute       '执行登记
'        .Add 0, vbKeyF9, conMenu_Edit_AnimalHeat    '批量录入体温单
'        .Add 0, vbKeyF10, conMenu_Edit_NurseLogFile '批量录入记录单
        
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找病人
'        .Add 0, vbKeyF3, conMenu_View_FindNext      '查找下一个
        .Add 0, vbKeyF10, conMenu_View_Notify       '医嘱提醒
        .Add 0, vbKeyF5, conMenu_View_Refresh       '刷新
        .Add 0, vbKeyF4, conMenu_View_NoticBoard    '公告栏
        .Add 0, vbKeyF12, conMenu_File_Parameter    '参数设置
    End With

    '读取发布到该模块的报表(不含虚拟模块的,病人帐页、住院科室日报、催款单、催款表都不显示,后面手工加到文件菜单下)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1261_5", "ZL1_INSIDE_1261_4", "ZL1_INSIDE_1261_6", "ZL1_INSIDE_1132", "ZL1_INSIDE_1139_1", "ZL1_INSIDE_1139_2", "ZL1_INSIDE_1139_3", "ZL1_INSIDE_1261_7", "ZL1_INSIDE_1261_8")
    
    '再处理分页控件
    With PatiPage
        With .PaintManager
            .Color = xtpTabColorOffice2003
            .Appearance = xtpTabAppearanceVisualStudio
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(页面.待入科, "待入科", rptPati(页面.待入科).hwnd, 0).Tag = "待入科"
        .InsertItem(页面.转科, "最近转科", picPara(0).hwnd, 0).Tag = "最近转科"
        .InsertItem(页面.出院, "最近出院", picPara(1).hwnd, 0).Tag = "最近出院"
        .InsertItem(页面.家庭病床, "家庭病床", rptPati(页面.家庭病床).hwnd, 0).Tag = "家庭病床"
    End With
    
    '53740:刘鹏飞,2012-09-19,加载外挂功能菜单
    Call DefCommandPlugIn(cbsMain)
End Sub

Private Sub DefCommandPlugIn(ByRef cbsMain As Object)
'功能：外挂部件菜单接入，菜单栏和工具栏
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim i As Long
    Dim lngTmp As Long
    Dim blnGroup As Boolean
    Dim rsBar As ADODB.Recordset
    Dim strFunc As String
    Dim strFuncXML As String
    
    Set mrsPlugInBar = Nothing
    
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        err.Clear: On Error GoTo 0
    End If

    If mobjPlugIn Is Nothing Then Exit Sub
    Call mobjPlugIn.Initialize(gcnOracle, glngSys, P新版护士站, 1)
    strFunc = mobjPlugIn.GetFuncNames(glngSys, P新版护士站, 1, strFuncXML)
    If strFunc = "" Then Exit Sub
        Call MakePlugInBar(strFunc, strFuncXML, rsBar)
    Set mrsPlugInBar = zlDatabase.CopyNewRec(rsBar)
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    On Error GoTo errH
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    '菜单栏
    If Not objMenu Is Nothing Then
        rsBar.Filter = "IsInTool=1 and BarType=1"
        If Not rsBar.EOF Then
            rsBar.Sort = "序号"
            With objMenu.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!功能名)
                        objControl.IconId = rsBar!图标ID
                        objControl.Parameter = rsBar!功能名
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                        blnGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
        
        rsBar.Filter = "IsInTool=0 and BarType=1"
        If Not rsBar.EOF Then
            rsBar.Sort = "序号"
            Set objPopup = objMenu.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能", , False)
                objPopup.BeginGroup = True
                
            With objPopup.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                        blnGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '工具栏按钮
    Set objBar = cbsMain(2)
    Set objControl = objBar.FindControl(, conMenu_File_Exit)
    If Not objControl Is Nothing Then
        objControl.BeginGroup = True
        lngTmp = objControl.Index - 1
    Else
        lngTmp = -1
    End If
    
    rsBar.Filter = "IsInTool=1 and BarType=2"
    If Not rsBar.EOF Then
        With objBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!功能名, lngTmp + 1)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                lngTmp = objControl.Index
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                rsBar.MoveNext
            Next
            objControl.BeginGroup = True
        End With
    End If
    
    rsBar.Filter = "IsInTool=0 and BarType=2"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        Set objPopup = objBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "扩展功能", lngTmp + 1, False)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.IconId = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        lngTmp = objPopup.Index
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名, lngTmp + 1)
                objControl.IconId = rsBar!图标ID
                objControl.Parameter = rsBar!功能名
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                lngTmp = objPopup.Index
                rsBar.MoveNext
            Next
        End With
    End If
    
    '自动执行的功能
    rsBar.Filter = "IsAuto=1"
    If Not rsBar.EOF Then mlngPlugInID = rsBar!功能ID
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd护理条件_Click()
    Dim i As Integer
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    mintREPORTSEL = -1
    Call mobjFilter.GetWindowRect(lngLeft, lngTop, lngRight, lngBottom)
    For i = 0 To lst护理条件.ListCount - 1
        If txt护理条件.Tag = "" Then
            lst护理条件.Selected(i) = True
        ElseIf InStr("," & txt护理条件.Tag & ",", "," & lst护理条件.ItemData(i) & ",") > 0 Then
            lst护理条件.Selected(i) = True
        Else
            lst护理条件.Selected(i) = False
        End If
    Next
    lst护理条件.ListIndex = 0
    pic护理条件.Top = lngTop + IIf(mobjFilter.Position = 4, 350, 0) '成为移动工具条后,需要加上标题栏的高度
    pic护理条件.Left = lngLeft + pic护理等级.Left
    pic护理条件.Width = txt护理条件.Width
    pic护理条件.Visible = True
    pic护理条件.ZOrder
    If lst护理条件.Visible And lst护理条件.Enabled Then lst护理条件.SetFocus
End Sub

Private Sub lst护理条件_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 Then
        For i = 1 To lst护理条件.ListCount - 1
            lst护理条件.Selected(i) = lst护理条件.Selected(0)
        Next
    ElseIf Not lst护理条件.Selected(Item) Then
        lst护理条件.Selected(0) = False
    ElseIf lst护理条件.SelCount = lst护理条件.ListCount - 1 Then
        lst护理条件.Selected(0) = True
    End If
End Sub

Private Sub lst护理条件_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst护理条件 _
        And Not Me.ActiveControl Is pic护理条件 Then pic护理条件.Visible = False
End Sub

Private Sub pic护理条件_GotFocus()
    If lst护理条件.Visible And lst护理条件.Enabled Then lst护理条件.SetFocus
End Sub

Private Sub pic护理条件_Resize()
    On Error Resume Next
    
    lst护理条件.Left = -15
    lst护理条件.Top = -15
    lst护理条件.Width = pic护理条件.Width
    
    cmdFilterCancel.Left = pic护理条件.ScaleWidth - cmdFilterCancel.Width - 100
    cmdFilterOK.Left = cmdFilterCancel.Left - cmdFilterOK.Width - 60
    
    cmdFilterOK.Top = lst护理条件.Height + (pic护理条件.ScaleHeight - lst护理条件.Height - cmdFilterOK.Height) / 2
    cmdFilterCancel.Top = cmdFilterOK.Top
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strTmp As String
    Dim rsPati As New ADODB.Recordset
    On Error GoTo ErrHand
    
    blnUnload = False
    mblnStart = False
    mlngSelect = -1
    mintPreDept = -1
    mbytFontSize = IIf(Val(zlDatabase.GetPara("显示字体大小", glngSys, 1265, 0, True)) = 0, 9, 12)
    mlngSource = IIf(mbytFontSize = 9, 999, 0)
    mintIndex = 0
    mblnRefresh = False
    mblnCardCollapse = False
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    mstrPrivs_检验采集 = GetPrivFunc(glngSys, 1211)
    mintPatiInputType = 11
    '74410:就诊卡为密则不显示
    mblnShowCard = Not ISPassShowCard
    Me.Caption = "新版住院护士工作站"
    
    'Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    If gblnDo = True Then
        lbl床号(mlngSource).Tag = IIf(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(cbsMain), cbsMain.Name & "ShowHomeNo", "0")) <> 0, 1, 0)
    Else
        lbl床号(mlngSource).Tag = 0
    End If
        Call HaveRIS
    '初始化菜单
    Call MainDefCommandBar
    cbsMain.RecalcLayout
    '监护仪
    mstrMonitor = ""
    mblnMonitor = Dir(App.Path & "\..\gdhs\AC2005.exe") <> ""
    If mblnMonitor Then mstrMonitor = App.Path & "\..\gdhs\AC2005.exe"
'    mblnMonitor = True '测试时使用(zlWardMonitor工程已经手工修改为测试用)
    Call InitComponent
    
    mintOutPreTime = -1
    Call InitSelectTime
    Call GetLocalSetting '本地参数
    
    '提取病人类型
    mstrSQL = " Select 名称,颜色 From 病人类型"
    Set mrsPatiColor = zlDatabase.OpenSQLRecord(mstrSQL, "提取病人类型设置信息")
    
    mblnSupport = (Val(Split(GetVersion, ".")(1)) >= 32)
    If mblnSupport Then
        mstrBriefCode = ",zlpinyincode(NVL(B.姓名,a.姓名),0,0,',',1) AS 简码 "
    Else
        mstrBriefCode = ",zlspellcode(NVL(B.姓名,a.姓名)) AS 简码"
    End If
    
    '初始化病人过滤条件
    strTmp = zlDatabase.GetPara("当前病况过滤", glngSys, p住院护士站, "111", _
        Array(chk病况条件(0), chk病况条件(1), chk病况条件(2)), InStr(mstrPrivs, "参数设置") > 0)
    For i = 0 To chk病况条件.UBound
        chk病况条件(i).Value = IIf(Mid(strTmp, i + 1, 1) = "1", 1, 0)
    Next
    If Not InitBedlevel Then Unload Me: Exit Sub
    If Not InitNurselevel Then Unload Me: Exit Sub
    If Not InitUnits Then Unload Me: Exit Sub
    If cboUnit.ListIndex = -1 Then
        If InStr(mstrPrivs, "全院病人") > 0 Then
            MsgBox "没有发现住院病区信息,请先到部门管理中设置！", vbInformation, gstrSysName
        Else
            MsgBox "没有发现你所属病区,不能使用住院护士工作站！", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If
    Call zlControl.CboSetWidth(cboUnit.hwnd, 3500)
    
    '正常启动结束
    Call RestoreWinState(Me, App.ProductName)
    
    '55928:刘鹏飞,2012-11-20,读取卡片是否折叠
    If gblnDo = True Then
        mblnCardCollapse = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(cbsMain), cbsMain.Name & "Collapse", "0")) <> 0
        Call SetSourceCardH
    End If
    
    Call RaisEffect(picInfo, 2)
    mblnRefresh = True
    mblnStart = True
    
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1265, mstrPrivs, Me.hwnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    Call mclsAdvices.zlInitMip(mclsMipModule)
    
    Call Hook(Me)
    
    '加载医嘱提醒内容
    With frmNotify
        .mblnNormal = False
        .mintNotify = mintNotify
        .mintNotifyDay = mintNotifyDay
        .mstrNotifyAdvice = mstrNotifyAdvice
        .mdtOutBegin = mdtOutBegin
        .mdtOutEnd = mdtOutEnd
        .mstrScope = mstrScope
        .mlng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
        .Show , Me
    End With
    
    Call ReSetFontSize
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    Me.WindowState = 2
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    picInfo.Top = lngTop
    picInfo.Width = Me.ScaleWidth
    
    PicDraw.Top = picInfo.Top + picInfo.Height
    PicDraw.Width = Me.ScaleWidth - 30
    PicDraw.Height = Me.ScaleHeight - PicDraw.Top - IIf(stbThis.Visible, stbThis.Height, 0)
    HScr.Top = PicDraw.Top
    HScr.Height = PicDraw.Height
    '下部控件
    picList.Left = 0
    picList.Top = fraPatiUD.Top
    picList.Height = PicDraw.Height - picList.Top
    picList.Width = PicDraw.Width - 255
    PatiPage.Left = 0
    PatiPage.Top = picList.Top
    PatiPage.Width = picList.Width
    PatiPage.Height = picList.Height - 60
    
    Call picPatiIn_Resize
    
    fraPatiUD.Left = picList.Left
    fraPatiUD.Width = picList.Width
    
    If picList.Visible Then
        fra审查.Left = picList.Width - fra审查.Width
        fra审查.Top = PicDraw.Top + picList.Top + 50
    Else
        fra审查.Left = stbThis.Width - fra审查.Width - 1500
        fra审查.Top = stbThis.Top + 50
    End If
    fraPatiUD.Visible = picList.Visible
    
    lblPatiInputType.Left = 120
    txt住院号.Left = lblPatiInputType.Left + lblPatiInputType.Width + 50
    pic出院查找.Top = picList.Top + 50
    pic出院查找.Left = 5000 + (TextWidth("刘") - 180) * 15
    pic出院查找.Width = txt住院号.Left + txt住院号.Width
    pic出院查找.Height = txt住院号.Height + txt住院号.Top
    
    picPara(2).Left = 80
    picPara(3).Left = 80
    
    Call RaisEffect(picInfo, 2)
    picInfo.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTmp As String
    Dim blnSetup As Boolean
    
    blnUnload = True
    timNotify.Enabled = False
    timeRefreshCard.Enabled = False

    '需要关闭所有子窗体（非模态）
    Unload frmNotify
    
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
        Set mfrmResponse = Nothing
    End If
    
    If Not mfrmNoticeBoard Is Nothing Then
        Unload mfrmNoticeBoard
        Set mfrmNoticeBoard = Nothing
    End If
    
    '54621:刘鹏飞,2013-02-28,护士站添加首页整理功能
    If Not mclsInOutMedRec Is Nothing Then
        Call mclsInOutMedRec.FormUnLoad
        Set mclsInOutMedRec = Nothing
    End If
    
    '卸载消息对象
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
    
    DoEvents
    
    Call UnHook(Me)
    Call UnloadControls
    
    strTmp = ""
    For i = 0 To chk病况条件.UBound
        strTmp = strTmp & IIf(chk病况条件(i).Value = 1, "1", "0")
    Next
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    Call zlDatabase.SetPara("当前病况过滤", strTmp, glngSys, p住院护士站, blnSetup)
    Call zlDatabase.SetPara("护理等级过滤", txt护理条件.Tag, glngSys, p住院护士站, blnSetup)
    '病人范围
    Dim curDate As Date
    curDate = zlDatabase.Currentdate
    '54436:刘鹏飞,2012-10-10,修改相应参数的模块号为p住院护士站
    Call zlDatabase.SetPara("最近转出天数", Val(txtChange.Text), glngSys, p住院护士站, blnSetup)
    Call zlDatabase.SetPara("显示字体大小", IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)), glngSys, mlngModul, blnSetup)

    '55928:刘鹏飞,2012-11-20,设置卡片是否折叠
    If gblnDo = True Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(cbsMain), cbsMain.Name & "Collapse", IIf(mblnCardCollapse = True, 1, 0)
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(cbsMain), cbsMain.Name & "ShowHomeNo", Val(lbl床号(mlngSource).Tag)
    End If

    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjPlugIn Is Nothing Then
        Call mobjPlugIn.Terminate(glngSys, P新版护士站, 1)
    End If
    
    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mclsAdvices = Nothing
    Set mclsTends = Nothing
    Set mclsFeeQuery = Nothing
    Set mclsInPatient = Nothing
    Set mclsWardMonitor = Nothing
    Set mobjProveCollect = Nothing
    Set mobjReport = Nothing
    Set mobjPlugIn = Nothing

    Call DeleteFile
End Sub

Private Sub chk病况条件_Click(Index As Integer)
    Dim i As Integer, k As Integer
    '至少选择一个
    For i = 0 To chk病况条件.UBound
        If chk病况条件(i).Value = 1 Then k = k + 1
    Next
    If k = 0 Then chk病况条件(Index).Value = 1
    
    If mblnHScroll Then
        If HScr.Value <> 0 Then
            mstrBoardKeys = ""
            HScr.Value = 0
        Else
            Call AdjustCard
        End If
    Else
        Call AdjustCard
    End If
End Sub

Private Sub cmdFilterCancel_Click()
    If txt护理条件.Enabled And txt护理条件.Visible Then txt护理条件.SetFocus
    pic护理条件.Visible = False
End Sub

Private Sub cmdFilterCancel_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst护理条件 _
        And Not Me.ActiveControl Is pic护理条件 Then pic护理条件.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim i As Integer
    
    If lst护理条件.SelCount = 0 Then
        MsgBox "请至少选择一种护理等级。", vbInformation, gstrSysName
        If lst护理条件.Enabled And lst护理条件.Visible Then lst护理条件.SetFocus
    End If
    
    If lst护理条件.Selected(0) Then
        txt护理条件.Text = "全部"
        txt护理条件.Tag = ""
    Else
        txt护理条件.Text = ""
        txt护理条件.Tag = ""
        For i = 1 To lst护理条件.ListCount - 1
            If lst护理条件.Selected(i) Then
                txt护理条件.Text = txt护理条件.Text & "," & lst护理条件.List(i)
                txt护理条件.Tag = txt护理条件.Tag & "," & lst护理条件.ItemData(i)
            End If
        Next
        txt护理条件.Text = Mid(txt护理条件.Text, 2)
        txt护理条件.Tag = Mid(txt护理条件.Tag, 2)
    End If
    
    If txt护理条件.Enabled And txt护理条件.Visible Then txt护理条件.SetFocus
    pic护理条件.Visible = False
    
    If mblnHScroll Then
        If HScr.Value <> 0 Then
            mstrBoardKeys = ""
            HScr.Value = 0
        Else
            Call AdjustCard
        End If
    Else
        Call AdjustCard
    End If
End Sub

Private Function Get护理等级(ByVal str护理等级 As String) As Integer
    '三级或无等级时,返回3
    If InStr(1, str护理等级, "特") <> 0 Or InStr(1, str护理等级, "重") <> 0 Then
        Get护理等级 = 0
    ElseIf InStr(1, str护理等级, "III") <> 0 Then
        Get护理等级 = 3
    ElseIf InStr(1, str护理等级, "二") <> 0 Or InStr(1, str护理等级, "2") <> 0 Or InStr(1, str护理等级, "Ⅱ") <> 0 Or InStr(1, str护理等级, "II") <> 0 Then
        Get护理等级 = 2
    ElseIf InStr(1, str护理等级, "一") <> 0 Or InStr(1, str护理等级, "1") <> 0 Or InStr(1, str护理等级, "Ⅰ") <> 0 Or InStr(1, str护理等级, "I") <> 0 Then
        Get护理等级 = 1
    Else
        Get护理等级 = 3
    End If
End Function

Private Sub cmdFilterOK_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst护理条件 _
        And Not Me.ActiveControl Is pic护理条件 Then pic护理条件.Visible = False
End Sub

Private Sub picPati_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim blnValid As Boolean
    
    mintREPORTSEL = -1
    '显示卡片选择标记
    If mlngSelect >= 0 Then
        '包床也一并取消选中
        With mrsBedInfo
            .Filter = "卡片索引=" & mlngSelect
            If !病人ID <> 0 Then
                If PicDraw.Enabled And PicDraw.Visible Then PicDraw.SetFocus
                .Filter = "病人ID=" & !病人ID
                Do While Not .EOF
                    '将选择状态清除,同时将卡片大小还原(有可能在折叠模式下)
                    picPati(!卡片索引).ZOrder 1
                    lblSelect(!卡片索引).Visible = False
                    If mblnCardCollapse Then
                        picPati(!卡片索引).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
                        picPati(!卡片索引).Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_折叠, 卡片背景_标准卡片_折叠)).Picture
                    End If
                    
                    .MoveNext
                Loop
            End If
            .Filter = 0
        End With
    End If
    
    mlngSelect = Index
    '53740:刘鹏飞,2012-09-19,先执行插件自动执行，在弹出菜单(以前方式导致无法正常显示右键菜单)
    'mblnShow = True
    mblnShow = False: Call ShowSelect
    If Button = 2 Then
        Dim cbrPopupBar As CommandBar
        Dim cbrPopupItem As CommandBarControl
        Dim cbrMenuBar As CommandBarControl
        Dim cbrPopup As CommandBarPopup
        Dim cbrControl As Object
        Dim intIndex As Integer, int卡片索引 As Integer
        Dim str个性标注 As String, strKey As String, blnDeleteMunu As Boolean, strDeployKey As String
        Dim rsCopyNotes As New ADODB.Recordset
        
        If y < Me.lblSelect(Index).Top Then     '点击的标注区域
            '显示出所有标注主题并提供选择
            If mrsNotes.RecordCount = 0 Then Exit Sub
            If Not LocatePatiRecord Then Exit Sub
            mrsBedInfo.Filter = "病人ID=" & mrsPatiInfo!病人ID & " And 包床=0"
            If mrsBedInfo.RecordCount = 0 Then
                mrsBedInfo.Filter = ""
                Exit Sub
            End If
            str个性标注 = mrsBedInfo!个性标注1 & "'" & mrsBedInfo!个性标注2 & "'" & mrsBedInfo!个性标注3
            
            int卡片索引 = mrsBedInfo!卡片索引
            
            intIndex = 0
            If x > img个性标记1(mlngSource).Left And x < img个性标记1(mlngSource).Left + img个性标记1(mlngSource).Width Then
                intIndex = 1
            ElseIf x > img个性标记2(mlngSource).Left And x < img个性标记2(mlngSource).Left + img个性标记2(mlngSource).Width Then
                If mrsBedInfo!个性标注1 = "" And mrsBedInfo!个性标注2 = "" Then
                    intIndex = 1
                Else
                    intIndex = 2
                End If
            ElseIf x > img个性标记3(mlngSource).Left And x < img个性标记3(mlngSource).Left + img个性标记3(mlngSource).Width Then
                If mrsBedInfo!个性标注1 = "" And mrsBedInfo!个性标注2 = "" And mrsBedInfo!个性标注3 = "" Then
                    intIndex = 1
                ElseIf mrsBedInfo!个性标注2 = "" And mrsBedInfo!个性标注3 = "" Then
                    intIndex = 2
                Else
                    intIndex = 3
                End If
            Else
                If mrsBedInfo!个性标注1 <> "" And mrsBedInfo!个性标注2 <> "" And mrsBedInfo!个性标注3 <> "" Then
                    Exit Sub
                ElseIf mrsBedInfo!个性标注1 = "" Then
                    intIndex = 1
                ElseIf mrsBedInfo!个性标注2 = "" Then
                    intIndex = 2
                Else
                    intIndex = 3
                End If
            End If
            '根据要更新显示的图标组号，排开已经存在的组
            strDeployKey = ""
            If intIndex = 1 Then
                If mrsBedInfo!个性标注2 <> "" Then
                    strDeployKey = Split(mrsBedInfo!个性标注2, ",")(0) & "," & Split(mrsBedInfo!个性标注2, ",")(1)
                End If
                If mrsBedInfo!个性标注3 <> "" Then
                    strDeployKey = IIf(strDeployKey = "", "", strDeployKey & "'") & Split(mrsBedInfo!个性标注3, ",")(0) & "," & Split(mrsBedInfo!个性标注3, ",")(1)
                End If
            ElseIf intIndex = 2 Then
                If mrsBedInfo!个性标注1 <> "" Then
                    strDeployKey = Split(mrsBedInfo!个性标注1, ",")(0) & "," & Split(mrsBedInfo!个性标注1, ",")(1)
                End If
                If mrsBedInfo!个性标注3 <> "" Then
                    strDeployKey = IIf(strDeployKey = "", "", strDeployKey & "'") & Split(mrsBedInfo!个性标注3, ",")(0) & "," & Split(mrsBedInfo!个性标注3, ",")(1)
                End If
            Else
                If mrsBedInfo!个性标注1 <> "" Then
                    strDeployKey = Split(mrsBedInfo!个性标注1, ",")(0) & "," & Split(mrsBedInfo!个性标注1, ",")(1)
                End If
                If mrsBedInfo!个性标注2 <> "" Then
                    strDeployKey = IIf(strDeployKey = "", "", strDeployKey & "'") & Split(mrsBedInfo!个性标注2, ",")(0) & "," & Split(mrsBedInfo!个性标注2, ",")(1)
                End If
            End If
            
            mrsBedInfo.Filter = ""
            If int卡片索引 <> Index Then Exit Sub
            
            Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
            cbrPopupBar.Title = "标注设定"
            If mlngSource = 999 Then
                Call cbrPopupBar.SetIconSize(16, 16)
            Else
                Call cbrPopupBar.SetIconSize(24, 24)
            End If
            
            mrsNotes.Filter = ""
            Set rsCopyNotes = zlDatabase.CopyNewRec(mrsNotes)
            mrsNotes.Filter = "标记序号 = 0"
            mrsNotes.Sort = "病区id,主题序号,标记序号"
            Do While Not mrsNotes.EOF
                '排开对应的主题
                If InStr(1, "'" & strDeployKey & "'", "'" & mrsNotes!病区ID & "," & mrsNotes!主题序号 & "'") = 0 Then
                    Set cbrPopup = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_标注1, mrsNotes!说明)
                    If mlngSource = 999 Then
                        Call cbrPopup.CommandBar.SetIconSize(16, 16)
                    Else
                        Call cbrPopup.CommandBar.SetIconSize(24, 24)
                    End If
                    blnDeleteMunu = False
                    rsCopyNotes.Filter = "病区id=" & mrsNotes!病区ID & " And 主题序号=" & mrsNotes!主题序号 & " And 标记序号>0"
                    Do While Not rsCopyNotes.EOF
                        strKey = rsCopyNotes!病区ID & "," & rsCopyNotes!主题序号 & "," & rsCopyNotes!标记序号 & "," & rsCopyNotes!图形索引 + 1
                        Set cbrPopupItem = cbrPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_标注1 + rsCopyNotes.RecordCount, rsCopyNotes!说明)
                        cbrPopupItem.IconId = conMenu_图标 + (rsCopyNotes!图形索引)
                        cbrPopupItem.Category = intIndex & "|" & rsCopyNotes!病区ID & "|" & rsCopyNotes!主题序号 & "|" & rsCopyNotes!标记序号 & "|" & rsCopyNotes!图形索引 + 1 & "|" & rsCopyNotes!说明
                        If InStr(1, "'" & str个性标注 & "'", "'" & strKey & "'") <> 0 Then
                            cbrPopupItem.Checked = True
                            blnDeleteMunu = True
                        End If
                        rsCopyNotes.MoveNext
                    Loop
                    If blnDeleteMunu = True Then
                        Set cbrPopupItem = cbrPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_标注1 + mrsNotes.RecordCount + 1, "清除标注")
                        cbrPopupItem.Category = intIndex & "|" & mrsNotes!病区ID & "|" & mrsNotes!主题序号 & "|0|0|"
                    End If
                End If
                mrsNotes.MoveNext
            Loop
            mrsNotes.Filter = 0
            cbrPopupBar.ShowPopup
        Else
            mrsBedInfo.Filter = "卡片索引=" & mlngSelect
            blnValid = (mrsBedInfo!病人ID > 0)
            mrsBedInfo.Filter = 0
            
            If blnValid Then
                '组装右键菜单(常用功能才加进来)
                Set cbrMenuBar = mobjPopupBatch
                Set cbrPopupBar = cbsMain.Add("右键菜单", xtpBarPopup)
                cbrPopupBar.Title = "右键菜单"
                
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_TurnUnit, "转病区(&D)"): cbrPopupItem.Category = "病人"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_TurnTeam, "转小组(&T)"):  cbrPopupItem.Category = "病人"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_Turn, "转科(&C)"): cbrPopupItem.Category = "病人": cbrPopupItem.BeginGroup = True
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_Bed, "换床(&B)"):  cbrPopupItem.Category = "病人"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_Out, "出院(&O)"):  cbrPopupItem.Category = "病人"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_Manage_Change_Undo, "撤销(&U)"): cbrPopupItem.BeginGroup = True: cbrPopupItem.Category = "撤销"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ReStop, "确认停止(&C)"): cbrPopupItem.BeginGroup = True: cbrPopupItem.Category = "医嘱业务"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_ReportLisView, "浏览检验结果(&R)"): cbrPopupItem.Category = "医嘱业务"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Billing, "记帐(&C)"): cbrPopupItem.BeginGroup = True: cbrPopupItem.Category = "费用业务"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ReBillingApply, "销帐申请(&L)"): cbrPopupItem.Category = "费用业务"
                
                If Not mrsPlugInBar Is Nothing Then
                    mrsPlugInBar.Filter = "IsInTool=1 and BarType=3"
                    For intIndex = 1 To mrsPlugInBar.RecordCount
                        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mrsPlugInBar!功能ID, mrsPlugInBar!功能名)
                            cbrPopupItem.IconId = mrsPlugInBar!图标ID
                            cbrPopupItem.Parameter = mrsPlugInBar!功能名
                            If Val(mrsPlugInBar!IsGroup) = 1 Then cbrPopupItem.BeginGroup = True
                        mrsPlugInBar.MoveNext
                    Next
                    mrsPlugInBar.Filter = "IsInTool=0 and BarType=3"
                    If mrsPlugInBar.RecordCount > 0 Then
                        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugInPop, "扩展功能"): cbrPopupItem.BeginGroup = True
                            cbrPopupItem.IconId = conMenu_Tool_PlugIn
                    End If
                    mrsPlugInBar.Filter = 0
                End If
                cbrPopupBar.ShowPopup
            End If
        End If
    Else
        '如果是左键,且是简洁模式
        If Button = 1 And mblnCardCollapse Then
'            If mblnShowCard = True Then
'                picPati(Index).Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
'                picPati(Index).Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_就诊卡, 卡片背景_标准卡片_就诊卡)).Picture
'            Else
'                picPati(Index).Height = IIf(mlngSource = 0, clngBigHeight_Normal, clngBaseHeight_Normal)
'                picPati(Index).Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片, 卡片背景_标准卡片)).Picture
'            End If
            picPati(Index).Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
            picPati(Index).Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_就诊卡, 卡片背景_标准卡片_就诊卡)).Picture
        End If
    End If
End Sub

Private Sub picPati_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And mblnCardCollapse Then
        picPati(Index).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
        picPati(Index).Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_折叠, 卡片背景_标准卡片_折叠)).Picture
    End If
    
    picList.ZOrder 0
    PatiPage.ZOrder 0
    fraPatiUD.ZOrder 0
    picPara(2).ZOrder 0
    picPara(3).ZOrder 0
    pic出院查找.ZOrder 0
End Sub

'-------------------------------------------------------------------------------
'以下是基础代码
'-------------------------------------------------------------------------------
Private Sub LoadPatiCard(ByVal intIndex As Integer, ByVal str床号 As String, ByVal lngX As Long, ByVal lngY As Long, Optional ByVal blnVisible As Boolean = False)
    '加载床位卡片
    '1、卡片上部
    '2、卡片主体
    
    Load picPati(intIndex)
    With picPati(intIndex)
        .Left = lngX
        .Top = lngY
        .Width = picPati(mlngSource).Width
'        If mblnCardCollapse Then
'            .Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
'            .Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_折叠, 卡片背景_标准卡片_折叠)).Picture
'        ElseIf mblnShowCard = True Then
'            .Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
'            .Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_就诊卡, 卡片背景_标准卡片_就诊卡)).Picture
'        Else
'            .Height = IIf(mlngSource = 0, clngBigHeight_Normal, clngBaseHeight_Normal)
'            .Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片, 卡片背景_标准卡片)).Picture
'        End If
        If mblnCardCollapse Then
            .Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
            .Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_折叠, 卡片背景_标准卡片_折叠)).Picture
        Else
            .Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
            .Picture = img卡片背景(IIf(mlngSource = 0, 卡片背景_大卡片_就诊卡, 卡片背景_标准卡片_就诊卡)).Picture
        End If
        .Visible = blnVisible
        .ZOrder 0
    End With
    
    '1、卡片上部
    Load img护理等级(intIndex)
    img护理等级(intIndex).Visible = True
    Set img护理等级(intIndex).Container = picPati(intIndex)
    Set img护理等级(intIndex).Picture = Nothing
    img护理等级(intIndex).Left = img护理等级(mlngSource).Left
    img护理等级(intIndex).Top = img护理等级(mlngSource).Top
    img护理等级(intIndex).Height = img护理等级(mlngSource).Height
    img护理等级(intIndex).Width = img护理等级(mlngSource).Width
    img护理等级(intIndex).ZOrder 1
    
    Load lblSelect(intIndex)
    Set lblSelect(intIndex).Container = picPati(intIndex)
    lblSelect(intIndex).Left = lblSelect(mlngSource).Left
    lblSelect(intIndex).Top = lblSelect(mlngSource).Top
    lblSelect(intIndex).Width = lblSelect(mlngSource).Width
    
    Load lbl床号(intIndex)
    Set lbl床号(intIndex).Container = picPati(intIndex)
    lbl床号(intIndex).Visible = True
    lbl床号(intIndex).FontSize = lbl床号(mlngSource).FontSize
    lbl床号(intIndex).Left = lbl床号(mlngSource).Left
    lbl床号(intIndex).Top = lbl床号(mlngSource).Top
    lbl床号(intIndex).Width = lbl床号(mlngSource).Width
    lbl床号(intIndex).Height = lbl床号(mlngSource).Height
    lbl床号(intIndex).Caption = Mid(str床号, 1, 7)
    
    Load lbl房间号(intIndex)
    Set lbl房间号(intIndex).Container = picPati(intIndex)
    lbl房间号(intIndex).Caption = str床号
    lbl房间号(intIndex).Visible = False
    
    Load lbl姓名(intIndex)
    Set lbl姓名(intIndex).Container = picPati(intIndex)
    lbl姓名(intIndex).Visible = True
    lbl姓名(intIndex).FontSize = lbl姓名(mlngSource).FontSize
    lbl姓名(intIndex).Left = lbl姓名(mlngSource).Left
    lbl姓名(intIndex).Top = lbl姓名(mlngSource).Top
    lbl姓名(intIndex).Width = lbl姓名(mlngSource).Width
    lbl姓名(intIndex).Height = lbl姓名(mlngSource).Height
    lbl姓名(intIndex).Caption = ""
    
    Load lblSplit(intIndex)
    Set lblSplit(intIndex).Container = picPati(intIndex)
    lblSplit(intIndex).Visible = True
    lblSplit(intIndex).Left = lblSplit(mlngSource).Left
    lblSplit(intIndex).Top = lblSplit(mlngSource).Top
    lblSplit(intIndex).Width = lblSplit(mlngSource).Width
    lblSplit(intIndex).BackColor = &HFFFFFF
    
    Load img个性标记2(intIndex)
    Set img个性标记2(intIndex).Container = picPati(intIndex)
    img个性标记2(intIndex).Picture = img个性标记2(mlngSource).Picture
    img个性标记2(intIndex).Stretch = img个性标记2(mlngSource).Stretch
    img个性标记2(intIndex).Top = img个性标记2(mlngSource).Top
    img个性标记2(intIndex).Left = img个性标记2(mlngSource).Left
    img个性标记2(intIndex).Width = img个性标记2(mlngSource).Width
    img个性标记2(intIndex).Height = img个性标记2(mlngSource).Height
    
    Load img个性标记3(intIndex)
    Set img个性标记3(intIndex).Container = picPati(intIndex)
    img个性标记3(intIndex).Picture = img个性标记3(mlngSource).Picture
    img个性标记3(intIndex).Stretch = img个性标记3(mlngSource).Stretch
    img个性标记3(intIndex).Top = img个性标记3(mlngSource).Top
    img个性标记3(intIndex).Left = img个性标记3(mlngSource).Left
    img个性标记3(intIndex).Width = img个性标记3(mlngSource).Width
    img个性标记3(intIndex).Height = img个性标记3(mlngSource).Height
    
    Load img临床路径(intIndex)
    Set img临床路径(intIndex).Container = picPati(intIndex)
    img临床路径(intIndex).Picture = img临床路径(mlngSource).Picture
    img临床路径(intIndex).Stretch = img临床路径(mlngSource).Stretch
    img临床路径(intIndex).Top = img临床路径(mlngSource).Top
    img临床路径(intIndex).Left = img临床路径(mlngSource).Left
    img临床路径(intIndex).Width = img临床路径(mlngSource).Width
    img临床路径(intIndex).Height = img临床路径(mlngSource).Height
    
    Load img病案审查(intIndex)
    Set img病案审查(intIndex).Container = picPati(intIndex)
    img病案审查(intIndex).Picture = img病案审查(mlngSource).Picture
    img病案审查(intIndex).Stretch = img病案审查(mlngSource).Stretch
    img病案审查(intIndex).Top = img病案审查(mlngSource).Top
    img病案审查(intIndex).Left = img病案审查(mlngSource).Left
    img病案审查(intIndex).Width = img病案审查(mlngSource).Width
    img病案审查(intIndex).Height = img病案审查(mlngSource).Height
    
    Load img个性标记1(intIndex)
    Set img个性标记1(intIndex).Container = picPati(intIndex)
    img个性标记1(intIndex).Picture = img个性标记1(mlngSource).Picture
    img个性标记1(intIndex).Stretch = img个性标记1(mlngSource).Stretch
    img个性标记1(intIndex).Top = img个性标记1(mlngSource).Top
    img个性标记1(intIndex).Left = img个性标记1(mlngSource).Left
    img个性标记1(intIndex).Width = img个性标记1(mlngSource).Width
    img个性标记1(intIndex).Height = img个性标记1(mlngSource).Height
    
    Load img出院(intIndex)
    Set img出院(intIndex).Container = picPati(intIndex)
    img出院(intIndex).Picture = img出院(mlngSource).Picture
    img出院(intIndex).Stretch = img出院(mlngSource).Stretch
    img出院(intIndex).Top = img出院(mlngSource).Top
    img出院(intIndex).Left = img出院(mlngSource).Left
    img出院(intIndex).Width = img出院(mlngSource).Width
    img出院(intIndex).Height = img出院(mlngSource).Height
    
    '2、卡片主体
    Load lbl住院号(intIndex)
    Set lbl住院号(intIndex).Container = picPati(intIndex)
    lbl住院号(intIndex).Visible = True
    lbl住院号(intIndex).FontSize = lbl住院号(mlngSource).FontSize
    lbl住院号(intIndex).Left = lbl住院号(mlngSource).Left
    lbl住院号(intIndex).Top = lbl住院号(mlngSource).Top
    lbl住院号(intIndex).Width = lbl住院号(mlngSource).Width
    lbl住院号(intIndex).Height = lbl住院号(mlngSource).Height
    lbl住院号(intIndex).Caption = ""
    
    Load lbl性别(intIndex)
    Set lbl性别(intIndex).Container = picPati(intIndex)
    lbl性别(intIndex).Visible = True
    lbl性别(intIndex).FontSize = lbl性别(mlngSource).FontSize
    lbl性别(intIndex).Left = lbl性别(mlngSource).Left
    lbl性别(intIndex).Top = lbl性别(mlngSource).Top
    lbl性别(intIndex).Width = lbl性别(mlngSource).Width
    lbl性别(intIndex).Height = lbl性别(mlngSource).Height
    lbl性别(intIndex).Caption = ""
    
    Load lbl年龄(intIndex)
    Set lbl年龄(intIndex).Container = picPati(intIndex)
    lbl年龄(intIndex).Visible = True
    lbl年龄(intIndex).FontSize = lbl年龄(mlngSource).FontSize
    lbl年龄(intIndex).Left = lbl年龄(mlngSource).Left
    lbl年龄(intIndex).Top = lbl年龄(mlngSource).Top
    lbl年龄(intIndex).Width = lbl年龄(mlngSource).Width
    lbl年龄(intIndex).Height = lbl年龄(mlngSource).Height
    lbl年龄(intIndex).Caption = ""
    
    Load lbl医师(intIndex)
    Set lbl医师(intIndex).Container = picPati(intIndex)
    lbl医师(intIndex).Visible = True
    lbl医师(intIndex).FontSize = lbl医师(mlngSource).FontSize
    lbl医师(intIndex).Left = lbl医师(mlngSource).Left
    lbl医师(intIndex).Top = lbl医师(mlngSource).Top
    lbl医师(intIndex).Width = lbl医师(mlngSource).Width
    lbl医师(intIndex).Height = lbl医师(mlngSource).Height
    lbl医师(intIndex).Caption = ""
    
    Load lbl护士(intIndex)
    Set lbl护士(intIndex).Container = picPati(intIndex)
    lbl护士(intIndex).Visible = True
    lbl护士(intIndex).FontSize = lbl护士(mlngSource).FontSize
    lbl护士(intIndex).Left = lbl护士(mlngSource).Left
    lbl护士(intIndex).Top = lbl护士(mlngSource).Top
    lbl护士(intIndex).Width = lbl护士(mlngSource).Width
    lbl护士(intIndex).Height = lbl护士(mlngSource).Height
    lbl护士(intIndex).Caption = ""
    
    Load lbl病情(intIndex)
    Set lbl病情(intIndex).Container = picPati(intIndex)
    lbl病情(intIndex).Visible = True
    lbl病情(intIndex).FontSize = lbl病情(mlngSource).FontSize
    lbl病情(intIndex).Left = lbl病情(mlngSource).Left
    lbl病情(intIndex).Top = lbl病情(mlngSource).Top
    lbl病情(intIndex).Width = lbl病情(mlngSource).Width
    lbl病情(intIndex).Height = lbl病情(mlngSource).Height
    lbl病情(intIndex).Caption = ""
    
    Load lbl入院日期(intIndex)
    Set lbl入院日期(intIndex).Container = picPati(intIndex)
    lbl入院日期(intIndex).Visible = True
    lbl入院日期(intIndex).FontSize = lbl入院日期(mlngSource).FontSize
    lbl入院日期(intIndex).Left = lbl入院日期(mlngSource).Left
    lbl入院日期(intIndex).Top = lbl入院日期(mlngSource).Top
    lbl入院日期(intIndex).Width = lbl入院日期(mlngSource).Width
    lbl入院日期(intIndex).Height = lbl入院日期(mlngSource).Height
    lbl入院日期(intIndex).Caption = ""
    
    Load lbl住院天数(intIndex)
    Set lbl住院天数(intIndex).Container = picPati(intIndex)
    lbl住院天数(intIndex).Visible = True
    lbl住院天数(intIndex).FontSize = lbl住院天数(mlngSource).FontSize
    lbl住院天数(intIndex).Left = lbl住院天数(mlngSource).Left
    lbl住院天数(intIndex).Top = lbl住院天数(mlngSource).Top
    lbl住院天数(intIndex).Width = lbl住院天数(mlngSource).Width
    lbl住院天数(intIndex).Height = lbl住院天数(mlngSource).Height
    lbl住院天数(intIndex).Caption = ""
    
    Load lbl诊断(intIndex)
    Set lbl诊断(intIndex).Container = picPati(intIndex)
    lbl诊断(intIndex).FontSize = lbl诊断(mlngSource).FontSize
    lbl诊断(intIndex).Visible = True
    lbl诊断(intIndex).Left = lbl诊断(mlngSource).Left
    lbl诊断(intIndex).Top = lbl诊断(mlngSource).Top
    lbl诊断(intIndex).Width = lbl诊断(mlngSource).Width
    lbl诊断(intIndex).Height = lbl诊断(mlngSource).Height
    lbl诊断(intIndex).Caption = ""
    
    '61824:刘鹏飞,2013-05-23,显示单病种标志
    Load img单病种(intIndex)
    Set img单病种(intIndex).Container = picPati(intIndex)
    img单病种(intIndex).Picture = img单病种(mlngSource).Picture
    img单病种(intIndex).Stretch = img单病种(mlngSource).Stretch
    img单病种(intIndex).Top = img单病种(mlngSource).Top
    img单病种(intIndex).Left = img单病种(mlngSource).Left
    img单病种(intIndex).Width = img单病种(mlngSource).Width
    img单病种(intIndex).Height = img单病种(mlngSource).Height
    
    Load lbl结余(intIndex)
    Set lbl结余(intIndex).Container = picPati(intIndex)
    lbl结余(intIndex).Visible = True
    lbl结余(intIndex).FontSize = lbl结余(mlngSource).FontSize
    lbl结余(intIndex).Left = lbl结余(mlngSource).Left
    lbl结余(intIndex).Top = lbl结余(mlngSource).Top
    lbl结余(intIndex).Width = lbl结余(mlngSource).Width
    lbl结余(intIndex).Height = lbl结余(mlngSource).Height
    lbl结余(intIndex).Caption = ""
    
    Load lbl结余总额(intIndex)
    Set lbl结余总额(intIndex).Container = picPati(intIndex)
    lbl结余总额(intIndex).Visible = True
    lbl结余总额(intIndex).FontSize = lbl结余总额(mlngSource).FontSize
    lbl结余总额(intIndex).Left = lbl结余总额(mlngSource).Left
    lbl结余总额(intIndex).Top = lbl结余总额(mlngSource).Top
    lbl结余总额(intIndex).Width = lbl结余总额(mlngSource).Width
    lbl结余总额(intIndex).Height = lbl结余总额(mlngSource).Height
    lbl结余总额(intIndex).Caption = ""
    
    '74410:卡片显示就诊卡号
    Load lblCardNo(intIndex)
    Set lblCardNo(intIndex).Container = picPati(intIndex)
    lblCardNo(intIndex).Visible = mblnShowCard
    lblCardNo(intIndex).FontSize = lblCardNo(mlngSource).FontSize
    lblCardNo(intIndex).Left = lblCardNo(mlngSource).Left
    lblCardNo(intIndex).Top = lblCardNo(mlngSource).Top
    lblCardNo(intIndex).Width = lblCardNo(mlngSource).Width
    lblCardNo(intIndex).Height = lblCardNo(mlngSource).Height
    lblCardNo(intIndex).Caption = ""
    
    '66618:显示医疗付款方式
    Load lblMedPay(intIndex)
    Set lblMedPay(intIndex).Container = picPati(intIndex)
    lblMedPay(intIndex).Visible = True
    lblMedPay(intIndex).FontSize = lblMedPay(mlngSource).FontSize
    lblMedPay(intIndex).Left = lblMedPay(mlngSource).Left
    lblMedPay(intIndex).Top = lblMedPay(mlngSource).Top
    lblMedPay(intIndex).Width = IIf(mblnShowCard = True, lblMedPay(mlngSource).Width, lbl医师(mlngSource).Width)
    lblMedPay(intIndex).Height = lblMedPay(mlngSource).Height
    lblMedPay(intIndex).Caption = ""
    
'    If mblnShowCard = False Then
'        lbl结余(intIndex).Top = lbl入院日期(intIndex).Top
'        lbl结余总额(intIndex).Top = lbl结余(intIndex).Top
'        lbl入院日期(intIndex).Top = lblCardNo(intIndex).Top
'        lbl住院天数(intIndex).Top = lbl入院日期(intIndex).Top
'    End If
End Sub

Private Sub SetCardInfo(ByVal intIndex As Integer, ByVal ArrPatiInfo As Variant)
    Dim imgManIcon As ImageManagerIcons
    Dim int护理等级 As Integer
    
    '住院号,姓名,性别,年龄,诊断,医/护,费别,医疗付款方式,病况,入院日期,住院天数,余额,病人颜色,护理等级,就诊卡号
    lbl住院号(intIndex).Caption = CStr(ArrPatiInfo(0))
    lbl姓名(intIndex).Caption = CStr(ArrPatiInfo(1))
    lbl姓名(intIndex).Alignment = 1
    lbl性别(intIndex).Caption = CStr(ArrPatiInfo(2))
    If lbl性别(intIndex).Caption = "包床" Then lbl性别(intIndex).Visible = False
    lbl年龄(intIndex).Caption = CStr(ArrPatiInfo(3))
    If IsNumeric(lbl年龄(intIndex).Caption) Then lbl年龄(intIndex) = lbl年龄(intIndex) & "岁"
    lbl医师(intIndex).Caption = "医护:" & CStr(ArrPatiInfo(5))
    lbl护士(intIndex).Caption = "费别:" & CStr(ArrPatiInfo(6))
    lblMedPay(intIndex).Caption = CStr(ArrPatiInfo(7))
    lblCardNo(intIndex).Caption = CStr(ArrPatiInfo(14))
    lbl病情(intIndex).Caption = CStr(ArrPatiInfo(8))
    lbl入院日期(intIndex).Caption = CStr(ArrPatiInfo(9))
    lbl住院天数(intIndex).Caption = IIf(Val(ArrPatiInfo(10) & "") = 0, 1, ArrPatiInfo(10)) & "天"
    lbl诊断(intIndex).Caption = CStr(ArrPatiInfo(4))
    lbl结余总额(intIndex).Caption = CStr(ArrPatiInfo(11))
    lblSplit(intIndex).BackColor = Val(CStr(ArrPatiInfo(12)))
    
    '设置护理等级(特级紫,一级红,二级蓝,三级无)
    int护理等级 = Get护理等级(CStr(ArrPatiInfo(13)))
    Set img护理等级(intIndex).Picture = imgHLDJ(mlngSource).ListImages(int护理等级 + 1).Picture
    
    If lbl结余总额(intIndex).Caption <> "" Then
        If lbl结余总额(intIndex).Caption = "不限额度担保" Then
            lbl结余总额(intIndex).Caption = ""
            lbl结余(intIndex).Caption = "不限额度担保"
            lbl结余(intIndex).ForeColor = &HFF0000
            lbl结余(intIndex).ZOrder 0
        Else
            If Val(lbl结余总额(intIndex).Caption) < 0 Then
                lbl结余(intIndex).Caption = "欠费"
                lbl结余(intIndex).ForeColor = &HFF&
                lbl结余总额(intIndex).ForeColor = &HFF&
            Else
                lbl结余(intIndex).Caption = "余额"
            End If
        End If
    Else
        lbl结余(intIndex) = ""
        lbl结余总额(intIndex).Caption = ""
        lbl医师(intIndex).Caption = ""
        lbl护士(intIndex).Caption = ""
        lblMedPay(intIndex).Caption = ""
        lblCardNo(intIndex).Caption = ""
        lbl住院天数(intIndex).Caption = ""
        Set img个性标记2(intIndex).Picture = Nothing
        Set img病案审查(intIndex).Picture = Nothing
        Set img临床路径(intIndex).Picture = Nothing
        Set img个性标记3(intIndex).Picture = Nothing
        Set img个性标记1(intIndex).Picture = Nothing
        Set img出院(intIndex).Picture = Nothing
        Set img护理等级(intIndex).Picture = Nothing
        Set img单病种(intIndex).Picture = Nothing
    End If
    
    If mblnShowCard = True Then
        If Trim(lblCardNo(intIndex).Caption) = "" Then
            lblMedPay(intIndex).Width = lbl医师(mlngSource).Width
        Else
            lblMedPay(intIndex).Width = lblMedPay(mlngSource).Width
        End If
    End If
End Sub

Private Sub SetCardLabel(ByVal intIndex As Integer)
    Dim intTar As Integer
    Dim intSignIndex As Integer
    On Error GoTo ErrHand
    
    '设置卡片标注区域
    mrsBedInfo.Filter = "卡片索引=" & intIndex
    If mrsBedInfo.RecordCount <> 0 Then
        If mrsBedInfo!病案审查 <> 0 Then
            Set img病案审查(intIndex).Picture = Img标记(mlngSource).ListImages(Get病案图标序号(mrsBedInfo!病案审查)).Picture
        End If
        img病案审查(intIndex).Visible = mrsBedInfo!病案审查
        img病案审查(intIndex).Tag = "" & mrsBedInfo!病案审查名称
        
        If mrsBedInfo!临床路径 <> 0 Then
            Set img临床路径(intIndex).Picture = Img标记(mlngSource).ListImages(Get临床路径序号(mrsBedInfo!临床路径)).Picture
        End If
        img临床路径(intIndex).Visible = mrsBedInfo!临床路径
        img临床路径(intIndex).Tag = "" & mrsBedInfo!临床路径名称
        img临床路径(intIndex).Visible = mblnHavePath
        
        intSignIndex = 0
        If mrsBedInfo!个性标注1 <> "" Then
            intSignIndex = Split(mrsBedInfo!个性标注1, ",")(3)
            If intSignIndex > 0 And intSignIndex <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
                Set img个性标记1(intIndex).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(intSignIndex).Picture
            Else
                intSignIndex = 0
            End If
        End If
        img个性标记1(intIndex).Visible = intSignIndex > 0
        img个性标记1(intIndex).Tag = "" & mrsBedInfo!个性标注1名称
        
        If mrsBedInfo!病人状态 <> 0 Then
            Set img出院(intIndex).Picture = Img标记(mlngSource).ListImages(CLng(mrsBedInfo!病人状态)).Picture
        End If
        img出院(intIndex).Visible = mrsBedInfo!病人状态
        img出院(intIndex).Tag = "" & mrsBedInfo!病人状态名称
        
        intSignIndex = 0
        If mrsBedInfo!个性标注2 <> "" Then
            intSignIndex = Split(mrsBedInfo!个性标注2, ",")(3)
            If intSignIndex > 0 And intSignIndex <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
                Set img个性标记2(intIndex).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(intSignIndex).Picture
            Else
                intSignIndex = 0
            End If
        End If
        img个性标记2(intIndex).Visible = intSignIndex > 0
        img个性标记2(intIndex).Tag = "" & mrsBedInfo!个性标注2名称
        
        intSignIndex = 0
        If mrsBedInfo!个性标注3 <> "" Then
            intSignIndex = Split(mrsBedInfo!个性标注3, ",")(3)
            If intSignIndex > 0 And intSignIndex <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
                Set img个性标记3(intIndex).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(intSignIndex).Picture
            Else
                intSignIndex = 0
            End If
        End If
        img个性标记3(intIndex).Visible = intSignIndex > 0
        img个性标记3(intIndex).Tag = "" & mrsBedInfo!个性标注3名称
        
        '61824:刘鹏飞,2013-05-23,显示单病种标志
        If Nvl(mrsBedInfo!单病种) <> "" Then
            Set img单病种(intIndex).Picture = Img标记(mlngSource).ListImages("单病种").Picture
        End If
        img单病种(intIndex).Visible = Nvl(mrsBedInfo!单病种) <> ""
        img单病种(intIndex).Tag = Nvl(mrsBedInfo!单病种)
    End If
    
    mrsBedInfo.Filter = 0
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub UnloadControls()
    Dim i As Integer, j As Integer
    Dim strOut As String

    strOut = "删除控件开始时间: " & Now
    For j = picPati.Count - 2 To 1 Step -1
        Unload lblSplit(j)
        Unload lblSelect(j)
        Unload lbl床号(j)
        Unload lbl房间号(j)
        Unload lbl住院号(j)
        Unload lbl姓名(j)
        Unload lbl性别(j)
        Unload lbl年龄(j)
        Unload lbl医师(j)
        Unload lbl护士(j)
        Unload lbl病情(j)
        Unload lbl入院日期(j)
        Unload lbl住院天数(j)
        Unload lbl诊断(j)
        Unload lbl结余(j)
        Unload lbl结余总额(j)
        Unload lblCardNo(j)
        Unload lblMedPay(j)

        Unload img个性标记2(j)
        Unload img个性标记3(j)
        Unload img临床路径(j)
        Unload img病案审查(j)
        Unload img个性标记1(j)
        Unload img出院(j)
        '61824:刘鹏飞,2013-05-23,显示单病种标志
        Unload img单病种(j)
        Unload img护理等级(j)
        Unload picPati(j)
    Next
    strOut = strOut & vbCrLf & "删除控件开始时间: " & Now
    'MsgBox strOut
End Sub

Private Sub timeRefreshCard_Timer()
    Dim lngIndex As Long
    '如果选中了某个项目,进行闪烁处理
    If blnUnload Then Exit Sub
    If mblnShow Then Call ShowSelect: mblnShow = False
    If Not mblnRefresh Then Exit Sub
    
    lngIndex = cboUnit.ListIndex
    timeRefreshCard.Enabled = False
    mblnShow = True
    Call RefreshData
    mblnRefresh = False
    timeRefreshCard.Enabled = True
    If lngIndex >= 0 And cboUnit.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboUnit.hwnd, lngIndex)
    End If

    If mblnShow Then Call ShowSelect: mblnShow = False
    
    '刷新公告栏
    If Not mfrmNoticeBoard Is Nothing And cboUnit.ListIndex <> -1 Then
        If mfrmNoticeBoard.mblnShow = True Then Call mfrmNoticeBoard.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex))
    End If
End Sub

Private Sub ShowSelect()
    Dim rsTmp As New ADODB.Recordset
    '显示当前选择的项目
    
    If mlngSelect < 0 Then Exit Sub
    '包床也一并选中
    With mrsBedInfo
        .Filter = "卡片索引=" & mlngSelect
        If !病人ID <> 0 Then
            mlng病人ID = !病人ID
            mlng主页ID = !主页ID
            
            .Filter = "病人ID=" & !病人ID
            Do While Not .EOF
                lblSelect(!卡片索引).Visible = True
                lblSelect(!卡片索引).ZOrder 1
                img护理等级(!卡片索引).ZOrder 1
                .MoveNext
            Loop
        Else
            mlng病人ID = 0
            mlng主页ID = 0
        End If
        .Filter = 0
    End With

    picPati(mlngSelect).ZOrder 0
    If picPati(mlngSelect).Visible And picPati(mlngSelect).Enabled Then picPati(mlngSelect).SetFocus
    
    Call GetPatiOtherInfo
End Sub

Private Sub GetPatiOtherInfo()
    '不管是在床病人还是非在床病人,均需提取其住院信息,在按钮状态变化时需使用
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '以下信息取当前住院次数的
    If Not LocatePatiRecord Then Exit Sub
    
    mPatiInfo.排序 = CStr(mrsPatiInfo!排序)
    mPatiInfo.病案状态 = Nvl(mrsPatiInfo!病案状态, 0)
    mPatiInfo.路径状态 = mrsPatiInfo!路径状态
    
    '取其它信息
    mstrSQL = "Select B.主页ID,B.状态,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期,b.出院日期,B.住院号,b.出院病床,B.病人性质,B.数据转出,B.险类,b.当前病区id,B.出院科室ID,B.当前病区ID,Decode(Nvl(X.费用余额, 0), 0, '√', '') As 结清" & _
        " From 病案主页 B,病人余额 X" & _
        " Where B.病人ID=[1] And B.主页ID=[2] And B.病人ID = X.病人ID(+) And X.性质(+) = 1 And X.类型(+)=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(mrsPatiInfo.Fields("病人ID").Value), Val(mrsPatiInfo.Fields("主页ID").Value))
    With rsTmp
        mPatiInfo.状态 = Nvl(!状态, 0)
        mPatiInfo.主页ID = Nvl(!主页ID, 0)
        mPatiInfo.住院号 = Nvl(!住院号)
        mPatiInfo.床号 = Nvl(!出院病床)
        mPatiInfo.病区ID = Nvl(!当前病区ID, 0)
        mPatiInfo.科室ID = Nvl(!出院科室ID, 0)
        mPatiInfo.入院日期 = !入院日期
        If Not IsNull(!出院日期) Then
            mPatiInfo.出院日期 = !出院日期
        Else
            mPatiInfo.出院日期 = CDate(0)
        End If
        mPatiInfo.险类 = Val("" & !险类)
        mPatiInfo.结清 = Not IsNull(!结清)
        mPatiInfo.性质 = Nvl(!病人性质, 0)
        mPatiInfo.产科 = is产科(Val(!出院科室ID & ""))
        mPatiInfo.数据转出 = Nvl(!数据转出, 0) <> 0
    End With
    '53740:刘鹏飞,2012-09-19,切换病人自动执行外挂功能
    Call AutoExecutePlugIn(cbsMain)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AutoExecutePlugIn(ByVal cbsMain As Object)
    Dim objControl As CommandBarControl
    Dim lng病人ID As Long, lng主页ID As Long
    
    If mrsPatiInfo.RecordCount = 0 Then
        lng病人ID = 0
        lng主页ID = 0
    Else
        lng病人ID = Val(mrsPatiInfo.Fields("病人ID").Value)
        lng主页ID = Val(mrsPatiInfo.Fields("病人ID").Value)
    End If
    '执行自动插件功能
    If mlngPlugInID <> 0 And (mlngPre病人ID <> lng病人ID Or (mlngPre病人ID = lng病人ID And mlngPre主页ID <> lng主页ID)) Then
        mlngPre病人ID = lng病人ID: mlngPre主页ID = lng主页ID
        Set objControl = cbsMain.FindControl(, mlngPlugInID, , True)
        If Not objControl Is Nothing Then objControl.Execute
    End If
End Sub

Private Sub GetInpatientAreaInfo()
    Dim strAdvance As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    '手术病人在标注手术主题时记录,在进行相关操作时更新,刷新的时候才从数据库里读取
    '53907:刘鹏飞,2012-09-28
'    mstrSQL = "" & _
'            " SELECT SUM(入院) AS 入院,SUM(入科) AS 入科,SUM(转出) AS 转出,SUM(死亡) AS 死亡,SUM(出院) AS 出院,SUM(危) AS 危,SUM(重) AS 重" & _
'            " FROM (" & _
'            "     SELECT SUM(DECODE(开始原因,2,1,0)) AS 入院,SUM(DECODE(开始原因,3,1,0)) AS 入科,0 AS 转出,0 AS 死亡,0 AS 出院,0 AS 危,0 AS 重" & _
'            "     From 病人变动记录" & _
'            "     Where 病区ID = [1]" & _
'            "     AND 开始时间 BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS 入院,0 AS 入科,SUM(DECODE(终止原因,3,1,0)) AS 转出,0 AS 死亡,0 AS 出院,0 AS 危,0 AS 重" & _
'            "     From 病人变动记录" & _
'            "     Where 病区ID = [1]" & _
'            "     AND 终止时间 BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS 入院,0 AS 入科,0 AS 转出,SUM(DECODE(出院方式,'死亡',1,0)) AS 死亡,SUM(DECODE(出院方式,'死亡',0,1)) AS 出院,0 AS 危,0 AS 重" & _
'            "     From 病案主页 A,病人信息 B" & _
'            "     Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.在院=1 And A.当前病区ID = [1]" & _
'            "     AND 出院日期 BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS 入院,0 AS 入科,0 AS 转出,0 AS 死亡,0 AS 出院,SUM(DECODE(当前病况,'危',1,0)) AS 危,SUM(DECODE(当前病况,'重',1,0)) AS 重" & _
'            "     From 病案主页 A,病人信息 B" & _
'            "     Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.在院=1 And A.当前病区ID = [1]" & _
'            "     AND 出院日期 IS NULL" & _
'            ")"
    mstrSQL = "" & _
            " SELECT SUM(入院) AS 入院,SUM(入科) AS 入科,SUM(转出) AS 转出,SUM(死亡) AS 死亡,SUM(出院) AS 出院,SUM(危) AS 危,SUM(重) AS 重" & _
            " FROM (" & _
            "     SELECT SUM(DECODE(开始原因,2,1,0)) AS 入院,SUM(DECODE(开始原因,3,1,15,1,0)) AS 入科,0 AS 转出,0 AS 死亡,0 AS 出院,0 AS 危,0 AS 重" & _
            "     From 病人变动记录" & _
            "     Where 病区ID = [1] And NVL(附加床位,0)=0" & _
            "     AND 开始时间 BETWEEN [2] AND SYSDATE" & _
            "     Union" & _
            "     Select SUM(1) as 入院,0 AS 入科,0 AS 转出,0 AS 死亡,0 AS 出院,0 AS 危,0 AS 重" & _
            "     From 病人变动记录 a, 病案主页 b" & _
            "     Where a.病人id = b.病人id And a.主页id = b.主页id And A.病区ID=[1] And A.开始时间 Between [2] And Sysdate And a.开始原因 = 1 And Nvl(a.附加床位, 0) = 0 And" & _
            "       Nvl(b.状态, 0) <> 1 And Not Exists" & _
            "       (Select 1 From 病人变动记录 Where 病人id = a.病人id And 主页id = b.主页id And 开始原因 = 2)"
    mstrSQL = mstrSQL & _
            "     Union" & _
            "     SELECT 0 AS 入院,0 AS 入科,SUM(DECODE(终止原因,3,1,15,1,0)) AS 转出,0 AS 死亡,0 AS 出院,0 AS 危,0 AS 重" & _
            "     From 病人变动记录" & _
            "     Where 病区ID = [1] And NVL(附加床位,0)=0" & _
            "     AND 终止时间 BETWEEN [2] AND SYSDATE" & _
            "     Union" & _
            "     SELECT 0 AS 入院,0 AS 入科,0 AS 转出,SUM(DECODE(出院方式,'死亡',1,0)) AS 死亡,SUM(DECODE(出院方式,'死亡',0,1)) AS 出院,0 AS 危,0 AS 重" & _
            "     From 病案主页 A,病人信息 B" & _
            "     Where A.病人ID=B.病人ID  And A.当前病区ID = [1]" & _
            "     AND 出院日期 BETWEEN [2] AND SYSDATE" & _
            "     Union" & _
            "     SELECT 0 AS 入院,0 AS 入科,0 AS 转出,0 AS 死亡,0 AS 出院,SUM(DECODE(当前病况,'危',1,0)) AS 危,SUM(DECODE(当前病况,'重',1,0)) AS 重" & _
            "     From 病案主页 A,病人信息 B,在院病人 C" & _
            "     Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And NVL(A.状态,0)<>1 And Nvl(A.病案状态,0)<>5 And A.封存时间 is NULL And B.病人ID=C.病人ID " & _
            "       And B.当前病区ID=C.病区ID And C.病区ID=[1]" & _
            ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "提取病区基本信息", cboUnit.ItemData(cboUnit.ListIndex), CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")))
    mlng入院 = Nvl(rsTemp!入院, 0)
    mlng转入 = Nvl(rsTemp!入科, 0)
    mlng出院 = Nvl(rsTemp!出院, 0)
    mlng转出 = Nvl(rsTemp!转出, 0)
    mlng死亡 = Nvl(rsTemp!死亡, 0)
    mlng危 = Nvl(rsTemp!危, 0)
    mlng重 = Nvl(rsTemp!重, 0)
    
    'LPF,2014-10-21,性能优化:添加在院病人表
'    mstrSQL = "" & _
'        " Select B.ID,B.名称,count(*) AS 人数" & vbNewLine & _
'        " From 病案主页 A,收费项目目录 B" & vbNewLine & _
'        " Where A.护理等级ID=B.ID And A.出院日期 IS Null And NVL(A.状态,0)<>1 And Nvl(A.病案状态,0)<>5 And A.封存时间 is NULL And A.当前病区ID=[1]" & vbNewLine & _
'        " Group by B.ID,B.名称"
    mstrSQL = "" & _
        " Select b.Id, b.名称, Count(*) As 人数" & vbNewLine & _
        " From 收费项目目录 b, 病人信息 c, 病案主页 a, 在院病人 e" & vbNewLine & _
        " Where b.Id = a.护理等级id And a.出院日期 Is Null And Nvl(a.状态, 0) <> 1 And Nvl(a.病案状态, 0) <> 5 And a.封存时间 Is Null And" & vbNewLine & _
        "      c.病人id = a.病人id And c.主页id = a.主页id And c.病人id = e.病人id And c.当前病区id = e.病区id And e.病区id = [1]" & vbNewLine & _
        " Group By b.Id, b.名称"

    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "提取病区基本信息", cboUnit.ItemData(cboUnit.ListIndex))
    Do While Not rsTemp.EOF
        strAdvance = strAdvance & "，" & rsTemp!名称 & "：" & rsTemp!人数 & "人"
        rsTemp.MoveNext
    Loop
    If strAdvance <> "" Then
        strAdvance = Mid(strAdvance, 2)
        strAdvance = "；" & strAdvance
    End If
    
    Call ShowInpatientAreaInfo(strAdvance)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowGuage(ByVal strInfo As String, ByVal dblPer As Double)
    Dim dblLength As Double     '进度条的当前宽度
    
    picInfo.Height = 315
    picInfo.BorderStyle = 1
    
    '显示进度条
    lblInpatientArea.Top = 60
    lblInpatientArea.AutoSize = False
    lblInpatientArea.Width = 3000
    lblInpatientArea.Caption = strInfo
    
    dblLength = picInfo.Width - lblInpatientArea.Width - 50
    '作图
    picInfo.Cls
    On Error Resume Next
    If Format(dblPer, "#0.00;-#0.00;0") <> "0" Then
        picInfo.PaintPicture picSource.Picture, lblInpatientArea.Width, 0, dblLength * dblPer / 100
    End If
    If err <> 0 Then err.Clear
    picInfo.Refresh
End Sub

Private Sub ShowInpatientAreaInfo(Optional ByVal strAdvance As String = "")
    Dim lng在院人数 As Long, lng总床位 As Long
    '10张空床，现有52人，入院_人，转入4人，转病区3人，出院5人，转出_人，死亡_人，危/重：1/_，手术5人
    
    '71923
    mrsBedInfo.Filter = "包床=0"
    lng在院人数 = mrsBedInfo.RecordCount + mlng家床 '- mlng预出院
    mrsBedInfo.Filter = "病人ID=0"
    mlng空床 = mrsBedInfo.RecordCount
    mrsBedInfo.Filter = 0
    lng总床位 = mrsBedInfo.RecordCount
    mlng在床 = lng总床位 - mlng空床
    
    picInfo.Cls
    picInfo.Height = 345
    
    lblInpatientArea.Top = 80
    lblInpatientArea.AutoSize = True
    lblInpatientArea.Caption = cboUnit.Text & "【基本情况】：共有" & lng总床位 & "张可用床位，" & mlng空床 & "张空床，现有" & lng在院人数 & "人(其中家床：" & mlng家床 & "人)；危/重：" & mlng危 & "/" & mlng重 & strAdvance
    lblInpatientArea.Caption = lblInpatientArea.Caption & "【当天情况】：入院" & mlng入院 & "人，转入" & mlng转入 & "人，出院" & mlng出院 & "人，转出" & mlng转出 & "人，死亡" & mlng死亡 & "人"
    
    Call RaisEffect(picInfo, 2)
End Sub

Private Sub Set诊疗项目费用设置()
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "诊疗基础部件(ZLCISBase)没有正确安装，该功能无法执行。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallSetClinicCharge(Me.cboUnit.ItemData(Me.cboUnit.ListIndex), 1, Me, gcnOracle, glngSys, gstrDBUser, E住院调用, InStr(GetInsidePrivs(mlngModul), ";诊疗项目费用设置;") = 0)
End Sub


'-------------------------------------------------------------------------------
'以下代码请忽略
'-------------------------------------------------------------------------------

Private Sub img病案审查_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img病案审查(Index).Tag, True
End Sub

Private Sub img出院_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img出院(Index).Tag, True
End Sub

Private Sub img个性标记1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img个性标记1(Index).Tag, True
End Sub

Private Sub img个性标记2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img个性标记2(Index).Tag, True
End Sub

Private Sub img个性标记3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img个性标记3(Index).Tag, True
End Sub

Private Sub img临床路径_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img临床路径(Index).Tag, True
End Sub

Private Sub img病案审查_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img出院_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img个性标记1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img个性标记2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img个性标记3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img临床路径_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img病案审查_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img病案审查(Index).Left + x, img病案审查(Index).Top + y)
End Sub

Private Sub img出院_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img出院(Index).Left + x, img出院(Index).Top + y)
End Sub

Private Sub img个性标记2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img个性标记2(Index).Left + x, img个性标记2(Index).Top + y)
End Sub

Private Sub img临床路径_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img临床路径(Index).Left + x, img临床路径(Index).Top + y)
End Sub

Private Sub img个性标记1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img个性标记1(Index).Left + x, img个性标记1(Index).Top + y)
End Sub

Private Sub img个性标记3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img个性标记3(Index).Left + x, img个性标记3(Index).Top + y)
End Sub

Private Sub lblSelect_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblSelect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblSelect(Index).Left + x, lblSelect(Index).Top + y)
End Sub

Private Sub lblSelect_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl床号_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl床号_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl床号(Index).Left + x, lbl床号(Index).Top + y)
End Sub

Private Sub lbl床号_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl护士_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl护士_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl护士(Index).Left + x, lbl护士(Index).Top + y)
End Sub

Private Sub lbl护士_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img护理等级_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub img护理等级_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img护理等级(Index).Left + x, img护理等级(Index).Top + y)
End Sub

Private Sub img护理等级_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl结余_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl结余_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl结余(Index).Left + x, lbl结余(Index).Top + y)
End Sub

Private Sub lbl结余_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl结余总额_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl结余总额_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl结余总额(Index).Left + x, lbl结余总额(Index).Top + y)
End Sub

Private Sub lbl结余总额_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl年龄_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl年龄_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl年龄(Index).Left + x, lbl年龄(Index).Top + y)
End Sub

Private Sub lbl年龄_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl入院日期_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl入院日期_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl入院日期(Index).Left + x, lbl入院日期(Index).Top + y)
End Sub

Private Sub lbl入院日期_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl性别_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl性别_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl性别(Index).Left + x, lbl性别(Index).Top + y)
End Sub

Private Sub lbl性别_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl姓名_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl姓名_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl姓名(Index).Left + x, lbl姓名(Index).Top + y)
End Sub

Private Sub lbl姓名_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl医师_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl医师_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl医师(Index).Left + x, lbl医师(Index).Top + y)
End Sub

Private Sub lbl医师_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl诊断_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl诊断_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl诊断(Index).Left + x, lbl诊断(Index).Top + y)
End Sub

Private Sub lbl诊断_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl住院号_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl住院号_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl住院号(Index).Left + x, lbl住院号(Index).Top + y)
End Sub

Private Sub lbl住院号_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl住院天数_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl住院天数_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl住院天数(Index).Left + x, lbl住院天数(Index).Top + y)
End Sub

Private Sub lbl住院天数_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lblSplit_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblSplit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblSplit(Index).Left + x, lblSplit(Index).Top + y)
End Sub

Private Sub lblSplit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub picPati_DblClick(Index As Integer)
    '弹出病人事务处理模块
    If Not LocatePatiRecord Then Exit Sub
    Call InNurseRoutine
End Sub

'54436:刘鹏飞,2012-10-10,修改了转科天数，过滤后，不能过滤出修改天数转科的病人
Private Sub txtChange_GotFocus()
    Call zlControl.TxtSelAll(txtChange)
End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mintChange = Val(txtChange.Text)
    txtChange.Text = mintChange
    
    rptPati(PatiPage.Selected.Index).Tag = ""
    rptPati(PatiPage.Selected.Index).Records.DeleteAll
    If rptPati(PatiPage.Selected.Index).Columns.Count > c_审查 Then rptPati(PatiPage.Selected.Index).Columns(c_审查).Visible = False
    Call PatiPage_SelectedChanged(PatiPage.Selected)
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Tag = "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
    txtFind.Tag = ""
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    
    '是否刷卡完成
    blnCard = mintFindType = 2 And KeyAscii <> 8 And Len(txtFind.Text) = gbytCardLen - 1 And txtFind.SelLength <> Len(txtFind.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txtFind.Text = txtFind.Text & Chr(KeyAscii)
            txtFind.SelStart = Len(txtFind.Text)
        End If
        KeyAscii = 0
        Call ExecuteFindPati
    Else
        If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
        Select Case mintFindType
            Case 0 '床号
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            Case 1 '住院号
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case 2 '就诊卡
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
            Case 3 '姓名
        End Select
    End If
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean)
    Dim blnRefresh As Boolean
    Dim str床号 As String, lng病人ID As Long, lng主页ID As Long, int排序 As Integer, intPage As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim objRptRow As ReportRow, strInput As String
    
    Call zlControl.TxtSelAll(txtFind)
    
    If Trim(txtFind.Text) = "" Then
        If mintFindType = 8 Then mintFindType = 0
        mrsBedInfo.Filter = ""
        Call AdjustCard
        Exit Sub
    End If
    
redo:
    '查找病人
    With mrsPatiInfo
        If mintFindType = 0 Then '床号
            .Filter = "床号='" & UCase(txtFind.Text) & "'"
        End If
        If mintFindType = 1 Then '住院号
            .Filter = "住院号=" & Val(txtFind.Text)
        End If
        If mintFindType = 2 Then '就诊卡
            .Filter = "就诊卡号='" & UCase(txtFind.Text) & "'"
        End If
        If mintFindType = 3 Then '姓名
            .Filter = "姓名 = '" & txtFind.Text & "'"
        End If
        If mintFindType = 4 Then '简码
            .Filter = "简码 Like '" & UCase(txtFind.Text) & "*'"
        End If
        If mintFindType = 4 Then
            mrsBedInfo.Filter = "简码 Like '" & UCase(txtFind.Text) & "*' OR 简码 Like '*," & UCase(txtFind.Text) & "*'"
            Call AdjustCard
            Exit Sub
        End If
        If .RecordCount = 0 Then
            .Filter = 0
            MsgBox "没有找到符合条件的记录！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        str床号 = !床号
        lng病人ID = !病人ID
        lng主页ID = !主页ID
        int排序 = !排序
        strInput = !住院号
        .Filter = 0
    End With
    On Error GoTo errH
    '检查搜索床的病人与数据库中是否相符,不符则重新提取床位卡
    'mstrSQL = " Select 当前床号 From 病人信息 Where 在院=1 And 病人ID=[1] And 当前病区ID=[2]"
    '53907:刘鹏飞,2012-09-28,应该加上病案主页，避免病人两次都在院
    mstrSQL = " Select B.出院病床 当前床号 From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID And B.病人ID=[1] And B.主页ID=[2] And B.当前病区ID=[3] And B.出院日期 IS NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "提取病人信息", lng病人ID, lng主页ID, CLng(Me.cboUnit.ItemData(Me.cboUnit.ListIndex)))
    If rsTemp.RecordCount <> 0 Then
        blnRefresh = (Nvl(rsTemp!当前床号, "") <> str床号)
        txtFind.Text = Nvl(rsTemp!当前床号, "")
        mintFindType = 0
    Else
        If int排序 = 5 Or int排序 = 6 Or int排序 = 7 Or int排序 = 1 Or int排序 = 0 Then
            blnRefresh = False
        Else
            blnRefresh = True
        End If
    End If
    If blnRefresh Then
        mblnRefresh = True
        Do While True
            DoEvents
            If mblnRefresh = False Then Exit Do
        Loop
        GoTo redo
    End If
    intPage = -1
    mrsBedInfo.Filter = "床号='" & str床号 & "'"
    If mrsBedInfo.RecordCount > 0 Then
        If Val(Nvl(mrsBedInfo!病人ID, 0)) = 0 Then
            mrsBedInfo.Filter = ""
            GoTo ErrNext
        End If
    Else
ErrNext:
        If int排序 = 0 Or int排序 = 1 Or int排序 = 2 Then
            intPage = 0
        ElseIf int排序 = 7 Then
            intPage = 1
        ElseIf int排序 = 6 Or int排序 = 5 Then
            intPage = 2
        ElseIf int排序 Like "3*" Or (int排序 = 4 And str床号 = "") Then '家庭病床
            intPage = 3
        End If
        PatiPage.Item(intPage).Selected = True
        
        For Each objRptRow In rptPati(intPage).Rows
            If Not objRptRow.Record Is Nothing Then
                If objRptRow.Record.Childs.Count = 0 Then
                    If IIf(Val(strInput) = 0, objRptRow.Record.Item(2).Value, objRptRow.Record.Item(5).Value) = IIf(Val(strInput) = 0, lng病人ID, strInput) Then
                        rptPati(intPage).Rows(objRptRow.Index).Selected = True
                        rptPati(intPage).SelectedRows(0).EnsureVisible
                        If rptPati(intPage).Visible Then rptPati(intPage).SetFocus
                        Exit For
                    End If
                End If
            End If
        Next
        mrsBedInfo.Filter = ""
    End If
'    If Not picPati(mrsBedInfo!卡片索引).Visible Then
'        mrsBedInfo.Filter = 0
'        MsgBox "已找到该病人，但由于该病人不符合过滤条件，请修改过滤条件后重新查找！", vbInformation, gstrSysName
'        Exit Sub
'    End If
    
    Call AdjustCard
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt护理条件_GotFocus()
    mintREPORTSEL = -1
End Sub

Private Sub txt住院号_GotFocus()
    txt住院号.ForeColor = &HFF0000
    Call zlControl.TxtSelAll(txt住院号)
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    Dim strValue As String, strField As String
    Dim strInput As String, strSQL As String
    Dim objRptRow As ReportRow
    Dim rsTemp As New ADODB.Recordset
    Dim blnCard As Boolean, blnOk As Boolean
    Dim strFilter As String
    Dim blnExit As Boolean
    On Error GoTo ErrHand
    
    '49752,刘鹏飞,2012-09-05,出院病人查找提供多种查找方式
    txt住院号.ForeColor = &HFF0000
    If KeyAscii = 39 Then KeyAscii = 0
    '是否刷卡完成
    blnCard = mintPatiInputType = 12 And KeyAscii <> 8 And Len(txt住院号.Text) = gbytCardLen - 1 And txt住院号.SelLength <> Len(txt住院号.Text)
    
    If KeyAscii = vbKeyReturn Or blnCard = True Then
        If KeyAscii <> 13 Then
            txt住院号.Text = txt住院号.Text & Chr(KeyAscii)
            txt住院号.SelStart = Len(txt住院号.Text)
        End If
        KeyAscii = 0
    Else
        Select Case mintPatiInputType
            Case 10 '床号
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            Case 11 '住院号
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case 12 '就诊卡
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
            Case 13 '姓名
        End Select
        Exit Sub
    End If
    
    strInput = Trim(txt住院号.Text)
    If strInput = "" Then Exit Sub
   
    '在出院页面中根据输入的住院号定位病人
    blnExit = False
FindPati:
    blnOk = False
    For Each objRptRow In rptPati(Val(pic出院查找.Tag)).Rows
        If Not objRptRow.Record Is Nothing Then
            If objRptRow.Record.Childs.Count = 0 Then
                Select Case mintPatiInputType
                    Case 10 '床号
                        If UCase(Trim(objRptRow.Record.Item(c_床号).Value)) = UCase(strInput) Then blnOk = True
                    Case 11 '住院号
                        If Val(objRptRow.Record.Item(c_住院号).Value) = Val(strInput) Then blnOk = True
                    Case 12 '就诊卡
                        If UCase(objRptRow.Record.Item(c_就诊卡号).Value) = UCase(strInput) Then blnOk = True
                    Case Else
                        If objRptRow.Record.Item(c_姓名).Value = strInput Then blnOk = True
                End Select
                If blnOk = True Then
                    rptPati(Val(pic出院查找.Tag)).Rows(objRptRow.Index).Selected = True
                    rptPati(Val(pic出院查找.Tag)).SelectedRows(0).EnsureVisible
                    If rptPati(Val(pic出院查找.Tag)).Visible Then rptPati(Val(pic出院查找.Tag)).SetFocus
                    Exit Sub
                End If
            End If
        End If
    Next
    
    '强制选择最后一个，避免错误导致死循环
    If blnExit = True And rptPati(Val(pic出院查找.Tag)).Rows.Count > 0 Then
        If Not rptPati(Val(pic出院查找.Tag)).Rows(rptPati(Val(pic出院查找.Tag)).Rows.Count - 1) Is Nothing Then
            rptPati(Val(pic出院查找.Tag)).Rows(rptPati(Val(pic出院查找.Tag)).Rows.Count - 1).Selected = True
            rptPati(Val(pic出院查找.Tag)).SelectedRows(0).EnsureVisible
            If rptPati(Val(pic出院查找.Tag)).Visible Then rptPati(Val(pic出院查找.Tag)).SetFocus
            Exit Sub
        End If
    End If
    If Val(pic出院查找.Tag) = 页面.家庭病床 Or Val(pic出院查找.Tag) = 页面.待入科 Then Exit Sub
    
    '如果找不到再从数据库中提取(出院病人页面提供此功能)
    '1--组织SQL条件
    strFilter = ""
    Select Case mintPatiInputType
        Case 10 '床号
            strFilter = " And B.出院病床=[2] "
        Case 11 '住院号
            strFilter = " And B.住院号=[2] "
        Case 12 '就诊卡
            strFilter = " And A.就诊卡号=[2] "
        Case Else
            strFilter = " And A.姓名=[2] "
    End Select
    '61824:刘鹏飞,2013-05-23,显示单病种标志
    '2--开始提取数据
    If pic出院查找.Tag = 页面.出院 Then
        strSQL = "" & _
            "Select /*+ RULE */ Decode(B.出院方式,'死亡',6,5) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.出院方式,'死亡','死亡病人','出院病人') as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(b.姓名,a.姓名) 姓名, NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
            " B.出院病床 AS 床号,E.名称 as 护理等级,B.费别,B.医疗付款方式,B.当前病况,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(b.路径状态,-1) 路径状态,trunc(b.出院日期)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID,A.主页Id 最大主页Id" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z" & _
            " Where A.病人ID=B.病人ID And B.病人类型=Z.名称(+) And Nvl(B.主页ID,0)<>0 And B.状态=0" & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And B.当前病区ID=[1] " & strFilter & " And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And B.出院日期 Is Not NULL And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
    ElseIf pic出院查找.Tag = 页面.转科 Then
         strSQL = "" & _
            "Select  Distinct 7 as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,'转出病人' as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,D.名称 as 科室,C.科室ID,C.经治医师 as 住院医师,B.责任护士,B.病案状态," & _
            " C.床号,E.名称 as 护理等级,B.费别,B.医疗付款方式,B.当前病况,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) As 入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间)) as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID,A.主页Id 最大主页Id" & _
            " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D,收费项目目录 E,病人类型 Z" & _
            " Where A.病人ID=B.病人ID And B.病人类型=Z.名称(+) And Nvl(B.主页ID,0)<>0 And B.护理等级ID=E.ID(+)" & _
            " And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
            " And B.当前病区ID<>[1] And C.病区ID+0=[1] And C.科室ID=D.ID " & strFilter & _
            " And Nvl(C.附加床位,0)=0 And C.终止原因 In(3,15) And C.终止时间 is Not Null " & _
            " And Nvl(B.状态,0)<>2 And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
    End If
    Set rsTemp = New ADODB.Recordset

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), strInput)
    Call UpgradeList(rsTemp)

    '追加记录集
    mrsPatiInfo.Filter = 0
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
        strField = "排序|排序2|类型|病人ID|主页ID|住院号|姓名|性别|年龄|科室|科室ID|住院医师|责任护士|病案状态|床号|护理等级|费别|医疗付款方式|当前病况|入院日期|出院日期|住院天数|出院方式|病人类型|状态|险类|就诊卡号|路径状态|颜色|单病种|婴儿科室ID|婴儿病区ID|最大主页Id"
        Do While Not rsTemp.EOF
            strValue = rsTemp!排序 & "|" & Nvl(rsTemp!排序2, 0) & "|" & Nvl(rsTemp!类型) & "|" & Nvl(rsTemp!病人ID, 0) & "|" & Nvl(rsTemp!主页ID, 0) & "|" & Nvl(rsTemp!住院号, 0) & "|" & Nvl(rsTemp!姓名) & "|" & Nvl(rsTemp!性别) & "|" & _
                      Nvl(rsTemp!年龄) & "|" & Nvl(rsTemp!科室) & "|" & Nvl(rsTemp!科室ID, 0) & "|" & Nvl(rsTemp!住院医师) & "|" & Nvl(rsTemp!责任护士) & "|" & Nvl(rsTemp!病案状态, 0) & "|" & Nvl(rsTemp!床号) & "|" & _
                      Nvl(rsTemp!护理等级, "三级") & "|" & Nvl(rsTemp!费别) & "|" & Nvl(rsTemp!医疗付款方式) & "|" & Nvl(rsTemp!当前病况, "一般") & "|" & Nvl(rsTemp!入院日期) & "|" & Nvl(rsTemp!出院日期) & "|" & Nvl(rsTemp!住院天数) & "|" & Nvl(rsTemp!出院方式) & "|" & _
                      Nvl(rsTemp!病人类型, "普通病人") & "|" & Nvl(rsTemp!状态, 0) & "|" & Nvl(rsTemp!险类, 0) & "|" & Nvl(rsTemp!就诊卡号) & "|" & Nvl(rsTemp!路径状态, 0) & "|" & Nvl(rsTemp!颜色, 0) & "|" & Nvl(rsTemp!单病种) & "|" & Nvl(rsTemp!婴儿科室ID, 0) & "|" & Nvl(rsTemp!婴儿病区ID, 0) & "|" & Nvl(rsTemp!最大主页ID, 0)
            Call Record_Add(mrsPatiInfo, strField, strValue)
            rsTemp.MoveNext
        Loop
        blnExit = True
        GoTo FindPati
    Else
        MsgBox "没有找到符合条件的记录！", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txt住院号_LostFocus()
    txt住院号.Text = ""
    txt住院号.ForeColor = &HC0C0C0
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'功能：结束打印事件，写入首页打印数据
    Dim strSQL As String
    
    strSQL = _
            "Zl_电子病历打印_Insert(Null,9," & mlng病人ID & "," & mPatiInfo.主页ID & ",'" & UserInfo.姓名 & "')"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitColor()
    Dim strValue As String
    Dim lng特级 As Long, lng一级 As Long, lng二级 As Long, lng三级 As Long
    Const c紫色 As Long = 8388736
    Const c红色 As Long = 255
    Const c兰色 As Long = 16711680
    Const c白色 As Long = 16777215
    
    Call DeleteFile
    mintIndex = 0
    imgHLDJ(0).ListImages.Clear
    imgHLDJ(999).ListImages.Clear
    '读取护理等级现有设置(无则取缺省数据)
    strValue = zlDatabase.GetPara("特级护理颜色", glngSys, 1265, "")
    lng特级 = IIf(strValue = "", c紫色, Val(strValue))
    strValue = zlDatabase.GetPara("一级护理颜色", glngSys, 1265, "")
    lng一级 = IIf(strValue = "", c红色, Val(strValue))
    strValue = zlDatabase.GetPara("二级护理颜色", glngSys, 1265, "")
    lng二级 = IIf(strValue = "", c兰色, Val(strValue))
    strValue = zlDatabase.GetPara("三级护理颜色", glngSys, 1265, "")
    lng三级 = IIf(strValue = "", c白色, Val(strValue))
    
    '绘图
    mlngColor = lng特级
    Call DrawPoly
    mlngColor = lng一级
    Call DrawPoly
    mlngColor = lng二级
    Call DrawPoly
    mlngColor = lng三级
    Call DrawPoly
End Sub

Private Sub AddColor()
    Dim strFile As String
    mintIndex = mintIndex + 1
    '不保存为文件,当创建多个图片时,加入到imagelist里的始终只有最后一个,应该是由于image中保存的是图片ID造成
    
    strFile = App.Path & "\HLDJTMP" & mintIndex & ".BMP"
    SavePicture picHLDJ.Image, strFile
    picHLDJ.Picture = LoadPicture(strFile)
    imgHLDJ(0).ListImages.Add , "K_" & mintIndex, picHLDJ.Picture
    imgHLDJ(999).ListImages.Add , "K_" & mintIndex, picHLDJ.Picture
End Sub

Private Sub DrawPoly()
    Dim lngRgn As Long, lngBrush As Long
    Dim lngPen As Long, lngOldPen As Long
    Dim PtInPoly() As POINTAPI

    '填充区域并划边线
    ReDim PtInPoly(4) As POINTAPI
    PtInPoly(1).x = 0
    PtInPoly(1).y = 0
    PtInPoly(2).x = picHLDJ.ScaleWidth
    PtInPoly(2).y = 0
    PtInPoly(3).x = picHLDJ.ScaleWidth
    PtInPoly(3).y = picHLDJ.ScaleHeight
    PtInPoly(4).x = PtInPoly(1).x
    PtInPoly(4).y = PtInPoly(1).y
    
    '创建系统刷子
    picHLDJ.Cls
    lngBrush = CreateSolidBrush(mlngColor)

    '如果创建刷子成功,才选入
    If lngBrush <> 0 Then
        lngRgn = CreatePolygonRgn(PtInPoly(1), UBound(PtInPoly), ALTERNATE)
        FillRgn picHLDJ.hdc, lngRgn, lngBrush
        Call DeleteObject(lngRgn)
        Call DeleteObject(lngBrush)
    End If
    picHLDJ.Refresh
    
    Call AddColor
End Sub

Private Sub DeleteFile()
    Dim objFile As File
    For Each objFile In mobjFileSys.GetFolder(App.Path).Files
        If Left(objFile.Name, 7) = "HLDJTMP" Then
            mobjFileSys.DeleteFile objFile.Path, True
        End If
    Next
End Sub

Private Sub ExecuteEditMediRec(Optional ByVal blnEditable As Boolean)
'功能：进行病案首页整理
'参数：blnEditable=是否允许编辑(有权限及签名允许的情况下)
    Dim blnReadOnly As Boolean
    
    If mPatiInfo.数据转出 Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '病案编目之后不可以整理
    If Not (CheckMecRed(mrsPatiInfo.Fields("病人ID").Value, mrsPatiInfo.Fields("主页ID").Value, Me.Caption) Or blnEditable) Then
        blnReadOnly = True
    End If
    
    '非模态显示首页整理
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, P新版护士站, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    If Not mclsInOutMedRec.IsOpen Then
        Call mclsInOutMedRec.ShowInMedRecEdit(Me, mrsPatiInfo.Fields("病人ID").Value, mrsPatiInfo.Fields("主页ID").Value, mrsPatiInfo.Fields("科室ID").Value, mrsPatiInfo.Fields("路径状态").Value, , mstrPrivs, IIf(blnReadOnly, 1, 0), False)
    End If
End Sub


Private Function CheckBabyInOut() As Boolean
'功能：检查婴儿和母亲是否分离，切当前在婴儿科室
    If Val(Nvl(mrsPatiInfo.Fields("婴儿病区ID").Value)) <> 0 Then
        If Val(Nvl(mrsPatiInfo.Fields("婴儿病区ID").Value)) = cboUnit.ItemData(cboUnit.ListIndex) And mintREPORTSEL = -1 Then
            MsgBox "该病人已经转出本科室了，只有婴儿留在本科室，不允许操作病人。", vbInformation, Me.Caption
            CheckBabyInOut = True
        End If
    End If
End Function

Private Function GetPatiCount(ByVal Index As Integer) As Long
'功能:获取非在床病人数目(由于病人列表进行了分组Records.Count统计出来不包含子项目,此处需要重新统计)
    Dim i As Long, lngCount As Long
    Dim objRecord As ReportRecord
    
    For i = 0 To rptPati(Index).Records.Count - 1
         If rptPati(Index).Records(i).Childs.Count > 0 Then
            lngCount = lngCount + rptPati(Index).Records(i).Childs.Count
         Else
            lngCount = lngCount + 1
         End If
    Next i
    
    GetPatiCount = lngCount
End Function

Private Sub MakePlugInBar(ByVal strFunc As String, ByVal strXML As String, rsBar As ADODB.Recordset)
'功能：组织菜单到本地记录集中，注意对老版本的兼容处理
'参数：strFunc 老版本功能列串，strXML含配置信息的功能串
    Dim strM As String
    Dim strB As String
    Dim strP As String
    Dim strTag As String
    Dim i As Long
    Dim strTmp As String
    Dim lngS As Long, lngE As Long
    
    If strXML = "" And strFunc <> "" Then
        '兼容以前老版本的方式
        Call InitPlugInRsBar(rsBar)
        Call AddPlugInBarRs(rsBar, strFunc, 1)
        Call AddPlugInBarRs(rsBar, strFunc, 2)
        Call AddPlugInBarRs(rsBar, strFunc, 3)
        Call SetPlugInBar(rsBar, 1)
        Exit Sub
    End If
    
    On Error GoTo errH
    strXML = Trim(strXML)
    '暂定为200个扩展功能插件，防止死循环
    For i = 0 To 200
        lngS = InStr(strXML, "<")
        lngE = InStr(strXML, ">")
        strTag = Mid(strXML, lngS + 1, lngE - lngS - 1)
        If strTag = "menubar" Then
            lngS = lngE
            lngE = InStr(strXML, "</menubar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strM = strM & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "toolbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</toolbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strB = strB & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "popbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</popbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strP = strP & "," & strTmp
            strXML = Mid(strXML, lngE + 9)
        End If
        If strXML = "" Then
            Exit For
        End If
    Next
    If strM = "" Then Exit Sub
    strM = Mid(strM, 2)
    strB = Mid(strB, 2)
    strP = Mid(strP, 2)

    Call InitPlugInRsBar(rsBar)
    Call AddPlugInBarRs(rsBar, strM, 1)
    Call AddPlugInBarRs(rsBar, strB, 2)
    Call AddPlugInBarRs(rsBar, strP, 3)
    Call SetPlugInBar(rsBar, 2)
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddPlugInBarRs(ByRef rsBar As ADODB.Recordset, ByVal strFunc As String, ByVal intType As Integer)
'功能：将功能串转换为记录集方式
'参数：strFunc 功能串，intType 功能按钮属于那一栏 1-菜单栏，2-工具栏，3-左键栏
    Dim varFunc As Variant
    Dim i As Long
    Dim strFuncName As String
    Dim blnFirstTool As Boolean
    If strFunc = "" Then Exit Sub
    varFunc = Split(strFunc, ",")
    With rsBar
        For i = 0 To UBound(varFunc)
            strFuncName = varFunc(i)
            .AddNew
            !BarType = intType
            If InStr(strFuncName, "Auto:") > 0 Then
                !IsAuto = 1
                strFuncName = Replace(strFuncName, "Auto:", "")
            Else
                !IsAuto = 0
            End If
            
            If InStr(strFuncName, "InTool:") > 0 Then
                !IsInTool = 1
                strFuncName = Replace(strFuncName, "InTool:", "")
            Else
                !IsInTool = 0
            End If
            If InStr(strFuncName, "|:") > 0 Then
                !IsGroup = 1
                strFuncName = Replace(strFuncName, "|:", "")
            Else
                !IsGroup = 0
                If Not blnFirstTool And !IsInTool = 1 Then
                    '第一个独立按钮显示分割线
                    blnFirstTool = True
                    !IsGroup = 1
                End If
            End If
            !功能名 = strFuncName
            !菜单名 = strFuncName
            .Update
        Next
    End With
End Sub

Private Function SetPlugInBar(ByRef rsBar As ADODB.Recordset, ByVal lngV As Long) As String
'功能：分配功能ID，加菜单快键
'参数：lngV 版本，1-老版，2-新版
'返回：字符串，以前低版本方式的功能串
    Dim i As Long
    '分配功能ID，图标ID
    With rsBar
        .Filter = 0
        If .EOF Then Exit Function
        .MoveFirst
        For i = 1 To .RecordCount
            !序号 = i
            !功能ID = conMenu_Tool_PlugIn_Item + i
            !图标ID = conMenu_Tool_PlugIn_Item
            If lngV = 1 Then
                !IsInTool = 0
                !IsGroup = 0
            End If
            .Update
            .MoveNext
        Next
    End With
    Call SetPlugInBarKey(rsBar, 1, lngV)
    Call SetPlugInBarKey(rsBar, 2, lngV)
    Call SetPlugInBarKey(rsBar, 3, lngV)
    rsBar.Filter = 0
End Function

Private Sub SetPlugInBarKey(rsBar As ADODB.Recordset, ByVal intType As Integer, ByVal lngV As Long)
'功能：设定快键
'参数：lngV 版本，1-老版，2-新版 intType 功能按钮属于那一栏 1-菜单栏，2-工具栏，3-左键栏
    Dim i As Long
    With rsBar
        .Filter = "IsInTool=0 and BarType=" & intType
        If .RecordCount = 1 And lngV = 2 Then
            '如果只有一个，也归为独立按钮
            !IsInTool = 1
            .Update
        Else
            For i = 1 To .RecordCount
                If i <= 35 Then
                    If i <= 9 Then
                        !菜单名 = !菜单名 & "(&" & i & ")"
                    Else
                        !菜单名 = !菜单名 & "(&" & Chr(55 + i) & ")"
                    End If
                    .Update
                    .MoveNext
                Else
                    Exit For
                End If
            Next
        End If
        
        .Filter = "IsInTool=1 and BarType=" & intType
        For i = 1 To .RecordCount
            If i <= 35 Then
                If i <= 9 Then
                    !菜单名 = !菜单名 & "(&" & i & ")"
                Else
                    !菜单名 = !菜单名 & "(&" & Chr(55 + i) & ")"
                End If
                .Update
                .MoveNext
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub InitPlugInRsBar(rsBar As ADODB.Recordset)
    Set rsBar = New ADODB.Recordset
    rsBar.Fields.Append "序号", adBigInt '用于排序
    rsBar.Fields.Append "功能ID", adBigInt '菜单按钮 Control.ID
    rsBar.Fields.Append "图标ID", adBigInt
    rsBar.Fields.Append "功能名", adVarChar, 1000 '去掉关键字之后的 名称 即工具栏上的按钮名称
    rsBar.Fields.Append "菜单名", adVarChar, 1000 '菜单栏/右键菜单 名称
    rsBar.Fields.Append "IsAuto", adInteger '是否自动执行功能
    rsBar.Fields.Append "IsGroup", adInteger '是否分割线
    rsBar.Fields.Append "IsInTool", adInteger '是否独立显示
    rsBar.Fields.Append "BarType", adInteger '1-菜单栏，2－工具栏，3－弹出栏
    rsBar.CursorLocation = adUseClient
    rsBar.LockType = adLockOptimistic
    rsBar.CursorType = adOpenStatic
    rsBar.Open
End Sub
