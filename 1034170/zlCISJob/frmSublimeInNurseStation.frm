VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSublimeInNurseStation 
   Caption         =   "�°�סԺ��ʿ����վ"
   ClientHeight    =   8985
   ClientLeft      =   225
   ClientTop       =   255
   ClientWidth     =   13410
   Icon            =   "frmSublimeInNurseStation.frx":0000
   LinkTopic       =   "frmSublimeInNurseStation"
   ScaleHeight     =   8985
   ScaleWidth      =   13410
   StartUpPosition =   2  '��Ļ����
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
            Key             =   "�ȴ����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":29C0
            Key             =   "�ܾ����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2F5A
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":34F4
            Key             =   "���ڳ��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3F06
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4918
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4EB2
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":58C4
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":62D6
            Key             =   "���鵵"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":CB38
            Key             =   "δ����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":D0D2
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":D66C
            Key             =   "������"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":E07E
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":E618
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":EBB2
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":F14C
            Key             =   "������"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":159AE
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":15F48
            Key             =   "����"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":164E2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":1CD44
            Key             =   "Ů��"
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
   Begin VB.PictureBox pic��λ״�� 
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
      Begin VB.ComboBox cbo��λ״�� 
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
   Begin VB.Frame fra��� 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   9390
      TabIndex        =   27
      Top             =   6120
      Width           =   3360
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "���� XXX ��δ����Ĳ�����鷴��..."
         BeginProperty Font 
            Name            =   "����"
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
   Begin VB.PictureBox pic������� 
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
      Begin VB.ComboBox cbo���� 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   30
         Width           =   1365
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   30
         Width           =   1365
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Caption         =   "���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   90
         Width           =   360
      End
   End
   Begin VB.CheckBox chk�����մ� 
      Appearance      =   0  'Flat
      Caption         =   "�����մ�"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8400
      TabIndex        =   11
      ToolTipText     =   "Ctrl+��ѡ������ѡ��"
      Top             =   570
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.PictureBox pic���� 
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
      Begin VB.CheckBox chk�������� 
         Appearance      =   0  'Flat
         Caption         =   "һ��"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   4
         ToolTipText     =   "Ctrl+��ѡ������ѡ��"
         Top             =   75
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.CheckBox chk�������� 
         Appearance      =   0  'Flat
         Caption         =   "Σ"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   690
         TabIndex        =   5
         ToolTipText     =   "Ctrl+��ѡ������ѡ��"
         Top             =   75
         Value           =   1  'Checked
         Width           =   465
      End
      Begin VB.CheckBox chk�������� 
         Appearance      =   0  'Flat
         Caption         =   "��"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "Ctrl+��ѡ������ѡ��"
         Top             =   75
         Value           =   1  'Checked
         Width           =   480
      End
   End
   Begin VB.PictureBox pic����ȼ� 
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
      Begin VB.CommandButton cmd�������� 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   1860
         Picture         =   "frmSublimeInNurseStation.frx":23C82
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(F4)"
         Top             =   60
         Width           =   270
      End
      Begin VB.TextBox txt�������� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   30
         Width           =   2160
      End
   End
   Begin MSComctlLib.ImageList Img��� 
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
            Key             =   "�໤��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":240CA
            Key             =   "�ȴ����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2441C
            Key             =   "�ܾ����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2476E
            Key             =   "���ڳ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":24AC0
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":24E12
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":25164
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":254B6
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":25808
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":25B5A
            Key             =   "δ����"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":25EAC
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":261FE
            Key             =   "������"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":26550
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":268A2
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":26BF4
            Key             =   "Ԥת��"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":26F46
            Key             =   "Ԥ��Ժ"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":27298
            Key             =   "��"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":275EA
            Key             =   "�к�"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2793C
            Key             =   "Ů��"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":27C8E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":27FE0
            Key             =   "Ů��"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":28332
            Key             =   "ҩ"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":28684
            Key             =   "��"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":289D6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":28D28
            Key             =   "Ǧ��"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2907A
            Key             =   "������"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":293CC
            Key             =   "���¼�"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2971E
            Key             =   "׼��"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":29A70
            Key             =   "ֹͣ"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":29DC2
            Key             =   "��ȷ"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2A114
            Key             =   "PDA"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2A466
            Key             =   "����"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2A7B8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2AB0A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2AE5C
            Key             =   "��ֹ"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2B1AE
            Key             =   "�ֻ�"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2B500
            Key             =   "ˢ��"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2B852
            Key             =   "��"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2BBA4
            Key             =   "ȷ��"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2BEF6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2C248
            Key             =   "�����"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2C59A
            Key             =   "�ػ�"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2C8EC
            Key             =   "����"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":2CC3E
            Key             =   "������"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img��� 
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
            Key             =   "�໤��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":33BB2
            Key             =   "�ȴ����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":33F04
            Key             =   "�ܾ����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":34256
            Key             =   "���ڳ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":345A8
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":348FA
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":34C4C
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":34F9E
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":352F0
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":35642
            Key             =   "δ����"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":35994
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":35CE6
            Key             =   "������"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":36038
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3638A
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":366DC
            Key             =   "Ԥת��"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":36DEE
            Key             =   "Ԥ��Ժ"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":37500
            Key             =   "��"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":37C12
            Key             =   "�к�"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":38324
            Key             =   "Ů��"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":38A36
            Key             =   "����"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":39148
            Key             =   "Ů��"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3985A
            Key             =   "ҩ"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":39F6C
            Key             =   "��"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3A67E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3AD90
            Key             =   "Ǧ��"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3B4A2
            Key             =   "������"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3BBB4
            Key             =   "���¼�"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3C2C6
            Key             =   "׼��"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3C9D8
            Key             =   "ֹͣ"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3D0EA
            Key             =   "��ȷ"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3D7FC
            Key             =   "PDA"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3DF0E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3E620
            Key             =   "����"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3ED32
            Key             =   "����"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3F444
            Key             =   "��ֹ"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":3FB56
            Key             =   "�ֻ�"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":40268
            Key             =   "ˢ��"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4097A
            Key             =   "��"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4108C
            Key             =   "ȷ��"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":4179E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":41EB0
            Key             =   "�����"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":425C2
            Key             =   "�ػ�"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":42CD4
            Key             =   "����"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":433E6
            Key             =   "������"
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
   Begin VB.PictureBox pic�������� 
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
      Begin VB.ListBox lst�������� 
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
         ToolTipText     =   "ȡ��"
         Top             =   1320
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   990
         Picture         =   "frmSublimeInNurseStation.frx":4BD18
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "ȷ��"
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
         Caption         =   "����������Ϣ:"
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
            Caption         =   "����ְ������ҽ�Ʊ���"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
         Begin VB.Image img������ 
            Height          =   360
            Index           =   0
            Left            =   2190
            Picture         =   "frmSublimeInNurseStation.frx":69510
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   360
         End
         Begin VB.Image img����ȼ� 
            Appearance      =   0  'Flat
            Height          =   360
            Index           =   0
            Left            =   2170
            Picture         =   "frmSublimeInNurseStation.frx":6FD62
            Stretch         =   -1  'True
            Top             =   38
            Width           =   345
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ƿ����"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "09123"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lblסԺ�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "027647132"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl�Ա� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "33"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lblҽʦ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ��:���ľ�/����ϼ"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl��Ժ���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2010-06-09"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl��� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����֧����������֧����������֧����������֧����������֧����������֧����������֧����������֧������"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl�����ܶ� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "34998.48"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl��ʿ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "�ѱ�:�Է�"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lblסԺ���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "25��"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Image img���Ա��2 
            Height          =   360
            Index           =   0
            Left            =   1440
            Picture         =   "frmSublimeInNurseStation.frx":70464
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img���Ա��1 
            Height          =   360
            Index           =   0
            Left            =   1095
            Picture         =   "frmSublimeInNurseStation.frx":70B66
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img��Ժ 
            Height          =   360
            Index           =   0
            Left            =   735
            Picture         =   "frmSublimeInNurseStation.frx":71268
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img�ٴ�·�� 
            Height          =   360
            Index           =   0
            Left            =   375
            Picture         =   "frmSublimeInNurseStation.frx":7196A
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img������� 
            Height          =   360
            Index           =   0
            Left            =   30
            Picture         =   "frmSublimeInNurseStation.frx":7206C
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Image img���Ա��3 
            Height          =   360
            Index           =   0
            Left            =   1800
            Picture         =   "frmSublimeInNurseStation.frx":7276E
            Stretch         =   -1  'True
            Top             =   60
            Width           =   360
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���������л����񹲺͹�"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl����� 
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
            Caption         =   "����ְ������ҽ�Ʊ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   35
            Top             =   1935
            Width           =   720
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Image img������ 
            Height          =   240
            Index           =   999
            Left            =   1860
            Picture         =   "frmSublimeInNurseStation.frx":87EB2
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label lbl����� 
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
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "09123"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Image img����ȼ� 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   999
            Left            =   1850
            Picture         =   "frmSublimeInNurseStation.frx":8E704
            Stretch         =   -1  'True
            Top             =   30
            Width           =   240
         End
         Begin VB.Label lblסԺ�� 
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
         Begin VB.Label lbl�Ա� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   999
            Left            =   1065
            TabIndex        =   42
            Top             =   840
            Width           =   180
         End
         Begin VB.Label lbl���� 
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
         Begin VB.Label lblҽʦ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ��:���ľ�/����ϼ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   40
            Top             =   1380
            Width           =   1995
         End
         Begin VB.Label lbl��Ժ���� 
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
         Begin VB.Label lbl��� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����֧����������֧����������֧����������֧����������֧����������֧����������֧����������֧����������֧������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   38
            Top             =   1110
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ƿ����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   37
            Top             =   2475
            Width           =   720
         End
         Begin VB.Label lbl�����ܶ� 
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
         Begin VB.Label lbl��ʿ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "�ѱ�:�Է�"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   999
            Left            =   60
            TabIndex        =   33
            Top             =   1650
            Width           =   810
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���������л����񹲺͹�"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lblסԺ���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "25��"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   999
            Left            =   1605
            TabIndex        =   31
            Top             =   2205
            Width           =   360
         End
         Begin VB.Image img���Ա��2 
            Height          =   240
            Index           =   999
            Left            =   1260
            Picture         =   "frmSublimeInNurseStation.frx":8EA46
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img���Ա��1 
            Height          =   240
            Index           =   999
            Left            =   960
            Picture         =   "frmSublimeInNurseStation.frx":8ED88
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img��Ժ 
            Height          =   240
            Index           =   999
            Left            =   660
            Picture         =   "frmSublimeInNurseStation.frx":8F0CA
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img�ٴ�·�� 
            Height          =   240
            Index           =   999
            Left            =   360
            Picture         =   "frmSublimeInNurseStation.frx":8F40C
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img������� 
            Height          =   240
            Index           =   999
            Left            =   60
            Picture         =   "frmSublimeInNurseStation.frx":8F74E
            Top             =   60
            Width           =   240
         End
         Begin VB.Image img���Ա��3 
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
            Caption         =   "�ѽ���"
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
            Caption         =   "δ����"
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
            Begin VB.Label lbl��Ժʱ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժʱ��"
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
               Caption         =   "ˢ��"
               Height          =   255
               Left            =   2520
               TabIndex        =   79
               Top             =   0
               Width           =   615
            End
            Begin VB.Label lblת�� 
               AutoSize        =   -1  'True
               Caption         =   "��ʾ���    ���ת������"
               Height          =   180
               Left            =   15
               TabIndex        =   82
               Top             =   30
               Width           =   2160
            End
         End
      End
      Begin VB.PictureBox pic��Ժ���� 
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
         Begin VB.TextBox txtסԺ�� 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00C0C0C0&
            Height          =   300
            Left            =   825
            MaxLength       =   100
            TabIndex        =   74
            ToolTipText     =   "����סԺ�Ŷ�λ����"
            Top             =   0
            Width           =   1485
         End
         Begin VB.Label lblPatiInputType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ�š�"
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
            Text            =   "�������"
            TextSave        =   "�������"
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
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
            Key             =   "������ɫ"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
            Key             =   "�໤��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":90DDE
            Key             =   "�ȴ����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":91558
            Key             =   "�ܾ����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":91CD2
            Key             =   "���ڳ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9244C
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":92BC6
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":93340
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":93ABA
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":94234
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":949AE
            Key             =   "δ����"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":95128
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":958A2
            Key             =   "������"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9601C
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":96796
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":96F10
            Key             =   "������"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9768A
            Key             =   "��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":97E04
            Key             =   "�к�"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9857E
            Key             =   "Ů��"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":98CF8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":99472
            Key             =   "Ů��"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":99BEC
            Key             =   "ҩ"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9A366
            Key             =   "��"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9AAE0
            Key             =   "����"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9B25A
            Key             =   "Ǧ��"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9B9D4
            Key             =   "������"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9C14E
            Key             =   "���¼�"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9C8C8
            Key             =   "׼��"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9D042
            Key             =   "ֹͣ"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9D7BC
            Key             =   "���"
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
            Key             =   "�໤��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9E2D0
            Key             =   "�ȴ����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9E66A
            Key             =   "�ܾ����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9EA04
            Key             =   "���ڳ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9ED9E
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9F138
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9F4D2
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9F86C
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9FC06
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":9FFA0
            Key             =   "δ����"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A033A
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A06D4
            Key             =   "������"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A0A6E
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A0E08
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A11A2
            Key             =   "������"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A153C
            Key             =   "��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A18D6
            Key             =   "�к�"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A1C70
            Key             =   "Ů��"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A200A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A23A4
            Key             =   "Ů��"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A273E
            Key             =   "ҩ"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A2AD8
            Key             =   "��"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A2E72
            Key             =   "����"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A320C
            Key             =   "Ǧ��"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A35A6
            Key             =   "������"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A3940
            Key             =   "���¼�"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A3CDA
            Key             =   "׼��"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A4074
            Key             =   "ֹͣ"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSublimeInNurseStation.frx":A440E
            Key             =   "���"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic��Ƭ���� 
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
      Begin VB.Image img��Ƭ���� 
         Height          =   2880
         Index           =   4
         Left            =   3300
         Picture         =   "frmSublimeInNurseStation.frx":A47A8
         Top             =   30
         Width           =   2235
      End
      Begin VB.Image img��Ƭ���� 
         Height          =   3315
         Index           =   5
         Left            =   4740
         Picture         =   "frmSublimeInNurseStation.frx":B97EA
         Top             =   45
         Width           =   2685
      End
      Begin VB.Image img��Ƭ���� 
         Height          =   945
         Index           =   3
         Left            =   2910
         Picture         =   "frmSublimeInNurseStation.frx":D6A58
         Top             =   3210
         Width           =   2685
      End
      Begin VB.Image img��Ƭ���� 
         Height          =   840
         Index           =   2
         Left            =   0
         Picture         =   "frmSublimeInNurseStation.frx":DEF7E
         Top             =   3210
         Width           =   2235
      End
      Begin VB.Image img��Ƭ���� 
         Height          =   2985
         Index           =   1
         Left            =   645
         Picture         =   "frmSublimeInNurseStation.frx":E51C0
         Top             =   0
         Width           =   2685
      End
      Begin VB.Image img��Ƭ���� 
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
    pt��Ժ����ס = 0
    ptת�ƴ���ס = 1
    ptת��������ס = 2
    pt��Ժ = 3
    pt��ͥ���� = 3.1
    ptԤת�� = 3.2
'    ptת���� = 3.3
    ptԤ�� = 4
    pt��Ժ = 5
    pt���� = 6
    pt���ת�� = 7
End Enum
Private Enum EFun
    E��ס = 0
    Eת�� = 1
    E���� = 2
    E���� = 3
    E��Ժ = 4
    EתΪסԺ = 5
    E���Ĵ�λ�ȼ� = 6
    E����������Ϣ = 7
    E�������Ǽ� = 8
    E������� = 9
    Eҽ������ѡ�� = 10
    E���� = 11
    E�޸ĳ�Ժʱ�� = 12
    E��λ�Ի� = 13
    Eתҽ��С�� = 14
    Eת���� = 15
    Eת������ס = 16
    E���˱�ע�༭ = 17
End Enum
Private Enum PATI_COLUMN
    c_���� = 0
    c_��� = 1
    c_ͼ�� = 2
    c_·��״̬ = 3
    C_����ID = 4
    C_��ҳID = 5
    c_���� = 6
    c_סԺ�� = 7
    c_���� = 8
    c_�Ա� = 9
    c_���� = 10
    c_�ѱ� = 11
    c_���ʽ = 12
    c_ҽ�� = 13
    c_��Ժ���� = 14
    c_��Ժ���� = 15
    c_�������� = 16
    c_���￨�� = 17
End Enum

Private Const mstrColWidth As String = "0,16,18,18,0,0,80,80,50,50,50,120,120,70,130,130,100,100"
        
Private Enum EFun_ҽ������
    E���� = 0
    EУ�� = 1
    Eֹͣ = 2
    E�鿴 = 3
End Enum

Private Const clngX = 100

Private Const ��Ƭ����_��׼��Ƭ As Integer = 0
Private Const ��Ƭ����_��Ƭ As Integer = 1
Private Const ��Ƭ����_��׼��Ƭ_�۵� As Integer = 2
Private Const ��Ƭ����_��Ƭ_�۵� As Integer = 3
Private Const ��Ƭ����_��׼��Ƭ_���￨ As Integer = 4
Private Const ��Ƭ����_��Ƭ_���￨ As Integer = 5

Private Const clngBaseHeight_Normal = 2595  '��׼��Ƭδ�۵�ʱ�ĸ߶�
Private Const clngBigHeight_Normal = 2985   '��Ƭδ�۵�ʱ�ĸ߶�
Private Const clngBaseCardHeight_Normal = 2880  '��׼��Ƭδ�۵�ʱ�ĸ߶ȣ���ʾ���￨��
Private Const clngBigCardHeight_Normal = 3315   '��Ƭδ�۵�ʱ�ĸ߶ȣ���ʾ���￨��
'��ɫ������ɫ����ʾ��������ʱ
Private Const clngBaseHeight_Collapse = 825 '��׼��Ƭ�۵�ʱ�ĸ߶�
Private Const clngBigHeight_Collapse = 920  '��Ƭ�۵�ʱ�ĸ߶�

'todo:ִ�����¹���ʱ,��������������ģ��,���50���Զ���ģ��
Private Const conMenu_���������� = 990000
Private Const conMenu_�鿴ҽ�� = 990001
Private Const conMenu_�鿴���� = 990002
Private Const conMenu_�鿴���� = 990003
Private Const conMenu_�鿴���µ� = 990004
Private Const conMenu_�鿴�����¼ = 990005
Private Const conMenu_�鿴������ = 990006

Private Const conMenu_ͼ�� = 990050                     '��ע��ʹ�õ�ͼ��ID��990050��ʼ,���150��ͼ��
Private Const conMenu_��ע1 = 990200
Private Const conMenu_��ע2 = 990300
Private Const conMenu_��ע3 = 990400
Private Const conMenu_��ע���� = 990500
Private Const conMenu_Manage_BedExchange = 2613         '*��λ�Ի�
Private Const conMenu_Edit_AnimalHeat = 3035            '*����¼�����µ�
Private Const conMenu_Edit_NurseLogFile = 3036          '*����¼���¼��
Private Const conMenu_ProveCollect = 3037               '����ɼ�����վ
Private Const conMenu_Edit_BatExecute = 3098            '*ҽ������ִ��

Private mPatiInfo As PatiInfo

'�Ӵ��������
Private mclsAdvices As zlPublicAdvice.clsDockInAdvices
Private mclsTends As zl9TendFile.clsTendFile
Private WithEvents mclsFeeQuery As zl9InExse.clsFeeQuery
Attribute mclsFeeQuery.VB_VarHelpID = -1
Private WithEvents mfrmResponse As frmAuditResponse '��鷴������
Attribute mfrmResponse.VB_VarHelpID = -1
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mfrmNoticeBoard As frmNoticeBoard  '���˹���������
Attribute mfrmNoticeBoard.VB_VarHelpID = -1
Private mclsInPatient As zl9InPatient.clsInPatient
Private mclsWardMonitor As clsWardMonitor     '�໤�ǽӿ�
Private mcolSubForm As Collection
Private mobjProveCollect As Object
Private mobjPlugIn As Object
Private mlngPlugInID As Long
Private mrsPlugInBar As ADODB.Recordset '�˵���ʽ�ṹ�� zlPlugIn/mdlPlugIn/ �� GetBarInfo ����
'54621:������,2013-02-28,��ʿվ�����ҳ������
Private mclsInOutMedRec As zlMedRecPage.clsInOutMedRec

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

'�������ñ���
Private blnUnload As Boolean
Private mstrPrivs As String
Private mstrPrivs_����ɼ� As String
Private mlngModul As Long
Private mstrUnits As String
Private mstrScope As String
Private mintFindType As Integer
Private mintPatiInputType As Integer  '��Ժ���˲���
Private mintChange As Integer
Private mintPage As Integer             '��Сһ����Ч��ҳ��
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mintOutPreTime As Integer
Private mintNotify As Integer           'ҽ�������Զ�ˢ�¼��(����)
Private mintNotifyDay As Integer        '���Ѷ������ڵ�ҽ��
Private mstrNotifyAdvice As String      '���ѵ�ҽ������
Private mstrCardInfo As String          '��Ƭ��ʾ����
Private mblnCardBalance As Boolean      '��Ƭ����Ƿ�����������
Private mblnCardOrder As Boolean         '��Ƭ�����Ƿ��մ�λ������
Private mblnCollateAutoFind As Boolean  'ҽ��������Զ���λ��ҽ��ҳ��
Public mintREPORTSEL As Integer        '��ǰѡ����ڴ��嵥����
Private mstrNoteItems As String         '���и������������,��:׼������,��ʼ����,��������|�к�,Ů��

Private mblnMonitor As Boolean          '�໤�ǳ����Ƿ����
Private mstrMonitor As String           '�໤�ǳ���·��
Private mstrBoardKeys As String         '�������������ص�������װ����Ϣ

'������������ֻ��¼�ڴ����˵���Ϣ
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlngPre����ID As Long
Private mlngPre��ҳID As Long
Private mblnReturn As Boolean           '������ť
'���Ʊ���
Private mintCards As Integer            '��ʾ�Ĵ�λ��Ƭ��
Public mblnRoutine As Boolean           '�Ƿ���ز����������ģ��
Private mstrSQL As String
Private mintPreDept As Integer          '��һ����
Private mblnShow As Boolean             '�����Ƿ���ʾ�����Ŀ�Ƭ����
Private mblnRefresh As Boolean          '�����Ƿ�ˢ�²�����λһ����
Private mlngSelect As Long              '��ǰѡ��Ŀ�Ƭ����
Private mlngSource As Long              '��¼��ǰ�Ǳ�׼��Ƭ���Ǵ�Ƭ
Private mbytFontSize As Byte             '������Ϣ9������12������
Private mblnStart As Boolean            '�����Ƿ���������
Private mblnCardCollapse As Boolean     '��Ƭ�Ƿ��۵�
Private mdblScaleHeight As Double       '��λ������ʵ�ʸ߶�
Private mblnHScroll As Boolean          '����������Ƿ���ʾ
Private mblnOutDept As Boolean          '�Ƿ������������Ŀ��ң��������۲�����ʾ����ţ�
Private mblnShowCard As Boolean         '�Ƿ���ʾ���￨��
Private mblnHavePath As Boolean          '��ǰ�����Ƿ���пɲ鿴���ٴ�·��

Private mobjPopup As CommandBarPopup    '�Ҽ������˵�\�������
Private mobjPopupBatch As CommandBarPopup    '�Ҽ������˵�\������������
Private mobjTheme As CommandBarControl  '�������
Private mobjFilter As CommandBar

'����������Ϣ
Private mlng�մ� As Long
Private mlng�ڴ� As Long
Private mlng��Ժ As Long
Private mlngת�� As Long
Private mlng�Ҵ� As Long
Private mlng��Ժ As Long
Private mlngԤ��Ժ As Long
Private mlngת�� As Long
Private mlng���� As Long
Private mlng���� As Long
Private mlngΣ As Long
Private mlng�� As Long

'�ڲ���¼������ر���
Private mstrFields As String
Private mstrValues As String
Private mrsBedInfo As New ADODB.Recordset   '��ǰ������λ��Ϣ
Private mrsPatiColor As New ADODB.Recordset '������������
Public mrsPatiInfo As New ADODB.Recordset  '���˼�¼������
Private mrsNotes As New ADODB.Recordset     '���������趨�ı������
Private mrsPatiNotes As New ADODB.Recordset '�������в��˵ı���嵥
Private mintMecStandard As Integer  '������ҳ��ʽ 0-��������׼��1-�Ĵ�ʡ��׼��2-����ʡ��׼
Private mlngMedRedDay As Long     '������鷴������

Dim mstrBriefCode As String
Dim mblnSupport As Boolean

Private Enum ҳ��
    �����
    ת��
    ��Ժ
    ��ͥ����
End Enum


'���ػ���ȼ���ɫ
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
'�趨һ�����岶����꣬���������������Ϣ�������ô���
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'ȡ����겶��
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private mlngColor As Long
Private mintIndex As Long
Private mobjFileSys As New FileSystemObject

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
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
    
    
    '����������С
    lst��������.Height = lst��������.ListCount * 210 + 30
    pic��������.Height = lst��������.Height + cmdFilterOK.Height + 120
    pic��������.Visible = False
    
    pic����.Height = TextHeight("��") + 60
    chk��������(0).Left = 0
    chk��������(0).Top = (pic����.Height - chk��������(0).Height) \ 2
    If chk��������(0).Top < 0 Then chk��������(0).Top = 0
    chk��������(1).Left = chk��������(0).Left + chk��������(0).Width
    chk��������(1).Top = chk��������(0).Top
    chk��������(2).Left = chk��������(1).Left + chk��������(1).Width
    chk��������(2).Top = chk��������(0).Top
    pic����.Width = chk��������(2).Left + chk��������(2).Width
    
    Label1.Top = cbo����.Top + (cbo����.Height - Label1.Height) \ 2
    cbo����.Left = Label1.Left + Label1.Width + 50
    lbl����.Left = cbo����.Left + cbo����.Width + TextWidth("��") / 2
    lbl����.Top = Label1.Top
    cbo����.Left = lbl����.Left + lbl����.Width + 50
    cbo����.Top = cbo����.Top
    pic�������.Width = cbo����.Left + cbo����.Width + 30
    chk�����մ�.Width = TextWidth("����" & chk�����մ�.Caption) - TextWidth("��") / 3
    txtFind.Width = 6 * TextWidth("��")
    
    '���°��¿ؼ�
    Set objFilter = cbsMain.Add("���˹�����", xtpBarTop)   '����
    objFilter.EnableDocking xtpFlagStretched
    objFilter.ContextMenuPresent = False
    With objFilter.Controls
        Set objControl = .Add(xtpControlLabel, 1, "����ȼ�")
        Set objCustom = .Add(xtpControlCustom, 2, "")
        objCustom.Handle = pic����ȼ�.hwnd
        Set objControl = .Add(xtpControlLabel, 3, "��λ״��"): objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, 4, "")
        objCustom.Handle = pic��λ״��.hwnd
        Set objControl = .Add(xtpControlLabel, 5, "��ǰ����"): objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, 6, "")
        objCustom.Handle = pic����.hwnd
        Set objCustom = .Add(xtpControlCustom, 7, ""): objCustom.BeginGroup = True
        objCustom.Handle = pic�������.hwnd
        Set objCustom = .Add(xtpControlCustom, 8, "")
        objCustom.Handle = chk�����մ�.hwnd: objCustom.BeginGroup = True

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "�������Ų���")
        objPopup.Caption = "�������Ų���"
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
    'ҳ��ת��
    fraChange.Left = lblת��.Left + TextWidth("ҳ��ת��")
    fraChange.Top = lblת��.Height + lblת��.Top
    fraChange.Width = TextWidth("ת��")
    txtChange.Width = TextWidth("999")
    txtChange.Left = fraChange.Left + (fraChange.Width - txtChange.Width) / 2
    txtChange.Height = TextHeight("��")
    txtChange.Top = fraChange.Top - txtChange.Height
    cmdRef.Left = lblת��.Left + lblת��.Width + 100
    cmdRef.Height = TextHeight("��") + 100
    cmdRef.Width = TextWidth(" ˢ�� ")
    cmdRef.Top = lblת��.Top - (cmdRef.Height - lblת��.Height) \ 2
    
    '��Ժ��ѯ
    cboSelectTime.Left = lbl��Ժʱ��.Left + lbl��Ժʱ��.Width + TextWidth("��") / 2
    picPara(2).Width = cboSelectTime.Left + cboSelectTime.Width + TextWidth("��")
    picPara(2).Height = (cboSelectTime.Top * 2) + cboSelectTime.Height
    chkSettle(0).Left = picPara(2).Width + 100
    If (picPara(2).Height - TextWidth("��")) \ 2 >= 0 Then
        chkSettle(0).Top = (picPara(2).Height - TextWidth("��")) \ 2
    End If
    chkSettle(1).Left = chkSettle(0).Left + chkSettle(0).Width + 100
    chkSettle(1).Top = chkSettle(0).Top
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    Dim lngCol As Long, lngIndex As Long, arrWidth() As String
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Call frmNotify.SetFontSize(bytSize)
    
    Me.FontSize = mbytFontSize
    Me.FontName = "����"
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            Select Case UCase(objCtrl.Name)
                Case UCase("Label1"), UCase("lbl����"), UCase("lblInpatientArea"), UCase("lbl��Ժʱ��"), UCase("lbl���"), UCase("lblת��"), UCase("Label2"), _
                    UCase("lblת��"), UCase("lblPatiInputType")
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("��") + 20
            End Select
        Case UCase("ListBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("VsFlexGrid")
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption) - TextWidth("��") / 3
        Case UCase("CheckBox")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption) - TextWidth("��") / 3
        Case UCase("DTPicker")
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01") + 400
            objCtrl.Height = TextHeight("��") * 1.5
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
    
    '�����б��п�����
    arrWidth = Split(mstrColWidth, ",")
    For lngIndex = 0 To rptPati.UBound
        For lngCol = c_ͼ�� To rptPati(lngIndex).Columns.Count - 1
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
    
    cboSelectTime.Clear '��Ժ
    With cboSelectTime
        .AddItem "������"
        .ItemData(.NewIndex) = 0
        .AddItem "������"
        .ItemData(.NewIndex) = 1
        .AddItem "ǰ����"
        .ItemData(.NewIndex) = 2
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "60����"
        .ItemData(.NewIndex) = 60
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
End Sub

Private Sub cboSelectTime_Click()
'���ܣ���ʱ�䷶Χ��ָ���ǣ�����ʱ��ѡ����
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If cboSelectTime.ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboSelectTime) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
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
        cboSelectTime.ToolTipText = "��Χ��" & Format(mdtOutBegin, "yyyy-MM-dd") & " �� " & Format(mdtOutEnd, "yyyy-MM-dd")
    End If
    '�����������֤ÿ���ط���ȡ�ĳ�Ժ���˶�����ͬһʱ�䷶Χ�ڣ�72783��
    Call zlDatabase.SetPara("��Ժ���˽������", DateDiff("d", datCurr, mdtOutEnd), glngSys, pסԺ��ʿվ)
    Call zlDatabase.SetPara("��Ժ���˿�ʼ���", DateDiff("d", mdtOutBegin, datCurr), glngSys, pסԺ��ʿվ)
    mintOutPreTime = cboSelectTime.ListIndex
    rptPati(PatiPage.Selected.Index).Tag = ""
    rptPati(PatiPage.Selected.Index).Records.DeleteAll
    If rptPati(PatiPage.Selected.Index).Columns.Count > c_��� Then rptPati(PatiPage.Selected.Index).Columns(c_���).Visible = False
    Call PatiPage_SelectedChanged(PatiPage.Selected)
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
'    Dim lngidx As Long
    Dim rsTmp As New ADODB.Recordset
    
    On Error Resume Next
    '�ȹر����м�ʱ��,�ٴ򿪰�����ʱ��ʱ��(���ؾ��޷�����ƥ��)
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

Private Sub cbo��λ״��_Click()
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
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "��  ��(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "סԺ��(&2)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "���￨(&3)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 4, "��  ��(&4)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 5, "��  ��(&5)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 9, "���"
            End If
        End With
    Case conMenu_File_MedRecPrint
        With CommandBar.Controls
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 1, "����(&1)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 2, "����(&2)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 3, "��ҳ1(&3)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 4, "��ҳ2(&4)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 5, "����+��ҳ1(&5)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 6, "����+��ҳ2(&6)"
            End If
        End With
    Case conMenu_File_MedRecPreview
        With CommandBar.Controls
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 1, "����(&1)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 2, "����(&2)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 3, "��ҳ1(&3)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 4, "��ҳ2(&4)"
            End If
        End With
    Case conMenu_Manage_Change_Undo
        With CommandBar.Controls
            .DeleteAll
            If Not LocatePatiRecord Then Exit Sub
            
            Set rsPatiLog = GetPatiLog(mrsPatiInfo!����ID, mrsPatiInfo!��ҳID)
            If rsPatiLog.RecordCount > 0 Then '��̬�Ӳ˵�,��1λ
                
                strPrivs = GetInsidePrivs(Enum_Inside_Program.p�������)
                rsPatiLog.MoveFirst
                For i = 1 To rsPatiLog.RecordCount
                    If Not IsNull(rsPatiLog!��ֹʱ��) And rsPatiLog!��ֹԭ�� = 1 Then
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Undo * 10 + i, "��Ժ")
                        j = j + 1
                        If InStr(";" & strPrivs & ";", ";������Ժ;") = 0 Or j > 1 Then objControl.Enabled = False
                    Else
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Undo * 10 + i, rsPatiLog!����)
                        If rsPatiLog.RecordCount > 1 And rsPatiLog!��ʼԭ�� = 1 Then objControl.Visible = False
                        j = j + 1
                        If j > 1 Then
                            objControl.Enabled = False
                        Else
                            If (objControl.Caption Like "*��ס" Or objControl.Caption = "ת������ס") Then
                                If InStr(strPrivs, "�������") = 0 Then objControl.Enabled = False
                            End If
                            If objControl.Caption = "תΪסԺ����" Then
                                If InStr(strPrivs, "סԺ����תסԺ") = 0 Then objControl.Enabled = False
                            ElseIf objControl.Caption = "Ԥ��Ժ" Then
                                If InStr(strPrivs, "����Ԥ��Ժ") = 0 Then objControl.Enabled = False
                                
                            ElseIf objControl.Caption = "����" Then
                                If InStr(strPrivs, "����") = 0 Then objControl.Enabled = False
                            End If
                        End If
                    End If
                    objControl.Category = "����"
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
                        Set objControl = .Add(xtpControlButton, mrsPlugInBar!����ID, mrsPlugInBar!�˵���)
                            objControl.IconId = mrsPlugInBar!ͼ��ID
                            objControl.Parameter = mrsPlugInBar!������
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
    '68259:������,2012-02-11,��Ժ���˲������δ�����ѽ��幦��
    If chkSettle(0).Value = 0 And chkSettle(1).Value = 0 Then
        chkSettle((Index + 1) Mod 2).Value = 1
    End If
    rptPati(PatiPage.Selected.Index).Tag = ""
    rptPati(PatiPage.Selected.Index).Records.DeleteAll
    If rptPati(PatiPage.Selected.Index).Columns.Count > c_��� Then rptPati(PatiPage.Selected.Index).Columns(c_���).Visible = False
    Call PatiPage_SelectedChanged(PatiPage.Selected)
End Sub

Private Sub chk��������_GotFocus(Index As Integer)
    mintREPORTSEL = -1
End Sub

Private Sub cmdRef_Click()
'54436:������,2012-10-10
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

'61824:������,2013-05-23,��ʾ�����ֱ�־
Private Sub img������_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img������(Index).Left + x, img������(Index).Top + y)
End Sub

Private Sub img������_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img������(Index).Tag, True
End Sub

Private Sub img������_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lblCardNo_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblCardNo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblCardNo(Index).Left + x, lblCardNo(Index).Top + y)
End Sub

Private Sub lblCardNo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, "���￨�ţ�" & lblCardNo(Index).Caption, True
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
    zlCommFun.ShowTipInfo picPati(Index).hwnd, "ҽ�Ƹ��ʽ��" & lblMedPay(Index).Caption, True
End Sub

Private Sub lblMedPay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '49752,������,2012-09-05,��Ժ�����ṩ���Ӳ��ҷ�ʽ(���š�סԺ�š����￨������)
    If Button = vbRightButton Then Exit Sub
   
    '�����˵�
    Dim intType As Integer
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Set cbrPopupBar = Me.cbsMain.Add("�����˵�", xtpBarPopup)
    intType = mintPatiInputType
    '���š�סԺ�š����￨������������
    With cbrPopupBar
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_View_FindType * 100# + 11, "��  ��(&1)")
        If intType = 10 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_View_FindType * 100# + 12, "סԺ��(&2)")
        If intType = 11 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_View_FindType * 100# + 13, "���￨(&3)")
        If intType = 12 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_View_FindType * 100# + 14, "��  ��(&4)")
        If intType >= 13 Then cbrPopupItem.Checked = True
    End With
    cbrPopupBar.ShowPopup
End Sub

Private Sub lbl�����ܶ�_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, Trim(lbl�����ܶ�(Index).Caption), True
End Sub

Private Sub lbl���_Click()
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    '��ģ̬��ʾ��鷴������
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
    
    '�ر�ҵ����
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
    End If
    
    '54621:������,2013-02-28,��ʿվ�����ҳ������
    If Not mclsInOutMedRec Is Nothing Then
        Call mclsInOutMedRec.FormUnLoad
    End If

    Call Have��������(cboUnit.ItemData(cboUnit.ListIndex), "����", mblnOutDept)
    With frmNotify
        .mintNotify = mintNotify
        .mintNotifyDay = mintNotifyDay
        .mstrNotifyAdvice = mstrNotifyAdvice
        .mdtOutBegin = mdtOutBegin
        .mdtOutEnd = mdtOutEnd
        .mlng����ID = cboUnit.ItemData(cboUnit.ListIndex)
    End With
    frmNotify.mblnFirst = True
End Sub

Private Sub cbo����_Click()
    Dim strInfo As String
    
    mintREPORTSEL = -1
    If Not mblnStart Then Exit Sub
    '��������
    strInfo = "��������"
    If Me.cbo����.Text <> "����" Then
        strInfo = cbo����.Text
        
        If Me.cbo����.Text <> "����" Then
            strInfo = strInfo & "\" & Me.cbo����.Text
        End If
    End If
    
    'ˢ�²�����λһ����
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

Private Sub cbo����_Click()
    Dim arrData
    Dim strData As String
    Dim i As Integer, j As Integer
    
    mintREPORTSEL = -1
    Me.cbo����.Clear
    Me.cbo����.AddItem "����"
    If Me.cbo����.Text <> "����" Then
        strData = Split(Me.cbo����.Tag, "|")(Me.cbo����.ListIndex - 1)
        If strData <> "" Then
            arrData = Split(strData, ",")
            j = UBound(arrData)
            For i = 0 To j
                '���Ա�����ݴ洢����˵��'������
                If InStr(1, arrData(i), "'") <> 0 Then
                    Me.cbo����.AddItem Split(arrData(i), "'")(0)
                    Me.cbo����.ItemData(cbo����.NewIndex) = Val(Split(arrData(i), "'")(1))
                Else
                    Me.cbo����.AddItem arrData(i)
                End If
            Next
        End If
    End If
    Me.cbo����.ListIndex = 0
    Me.cbo����.Enabled = (Me.cbo����.ListCount > 1)
    Me.cbo����.BackColor = IIf(Me.cbo����.Enabled, &H80000005, &HC0C0C0)
End Sub

Private Function LocatePatiRecord() As Boolean
    Dim intIndex As Integer
    Dim strTag As String
    Dim blnTrue As Boolean
    '���ݵ�ǰ�Ļ�ؼ�����λ����
    
    '122993
    If mrsPatiInfo.State = adStateClosed Then Exit Function
    If mintREPORTSEL = -1 Then
        If mlng����ID = 0 Then Exit Function
        mrsPatiInfo.Filter = "����ID=" & mlng����ID & " And ��ҳID=" & mlng��ҳID ' & " And (���� >=3 and ����<=3)"
        blnTrue = mrsPatiInfo.RecordCount
    Else
        intIndex = mintREPORTSEL
        If rptPati(intIndex).SelectedRows.Count = 0 Then GoTo ErrNext
        If rptPati(intIndex).SelectedRows(0).Record Is Nothing Then GoTo ErrNext
        If rptPati(intIndex).SelectedRows(0).Childs.Count > 0 Then GoTo ErrNext
        strTag = rptPati(intIndex).SelectedRows(0).Record.Tag
        mrsPatiInfo.Filter = "����ID=" & Split(strTag, "|")(0) & " And ��ҳID=" & Split(strTag, "|")(1)
        blnTrue = mrsPatiInfo.RecordCount
    End If
    '53740:������,2012-09-19,���ѡ��Ĳ��ǲ��˿�Ƭ����û��ѡ���κβ��ˣ�ȡ����Ƭ��ѡ��
ErrNext:
    If mintREPORTSEL <> -1 Or blnTrue = False Then
        If mlngSelect >= 0 Then
            '����Ҳһ��ȡ��ѡ��
            With mrsBedInfo
                .Filter = "��Ƭ����=" & mlngSelect
                If !����ID <> 0 Then
                    If PicDraw.Enabled And PicDraw.Visible Then PicDraw.SetFocus
                    .Filter = "����ID=" & !����ID
                    Do While Not .EOF
                        '��ѡ��״̬���,ͬʱ����Ƭ��С��ԭ(�п������۵�ģʽ��)
                        picPati(!��Ƭ����).ZOrder 1
                        lblSelect(!��Ƭ����).Visible = False
                        If mblnCardCollapse Then
                            picPati(!��Ƭ����).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
                            picPati(!��Ƭ����).Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_�۵�, ��Ƭ����_��׼��Ƭ_�۵�)).Picture
                        End If
                        
                        .MoveNext
                    Loop
                End If
                .Filter = 0
            End With
            picPati(mlngSelect).ZOrder 0
            mlngSelect = -1
            mlng����ID = 0: mlng��ҳID = 0
        End If
    End If
    
    LocatePatiRecord = blnTrue
End Function

Private Sub InNurseRoutine(Optional ByVal strPage As String = "ҽ��")
    '54408:������,2012-10-10,���벡����Ϣ��¼��
    Call frmInNurseRoutine.zlInitMip(mclsMipModule)
    Call frmInNurseRoutine.NurseRoutine(Me, mstrPrivs, Me.cboUnit.ItemData(Me.cboUnit.ListIndex), _
         Val(mrsPatiInfo.Fields("����ID").Value), mdtOutBegin, mdtOutEnd, mintChange, mstrScope, mPatiInfo, strPage, mrsPatiInfo, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
End Sub

Private Sub RefreshPatiList_Rountine()
    If Not mblnRoutine Then Exit Sub
    Call frmInNurseRoutine.RefreshPatiList(mrsPatiInfo)
End Sub

Private Sub OrientTabPage_Rountine(Optional ByVal strPage As String = "ҽ��", Optional ByVal strID As String = "")
    '-------------------------------------------------------------
    '����:��λ������������ָ����ҳ��,�Լ���Ӧҳ��ָ�����ļ���ҽ����
    '-------------------------------------------------------------
    '55430:������,2013-02-27,˫������ҽ����λ�����������ҽ��ҳ��
    If Not mblnRoutine Then Exit Sub
    Call frmInNurseRoutine.OrientTabPage(strPage, strID)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, byt��ס��ʽ As Byte, str���� As String, int��λ���� As Integer
    Dim strPrivs_������� As String, strPrivs_���� As String, strParentTitle As String, strTmp As String
    Dim blnExecuted As Boolean              '��ִ�����˳�
    Dim blnHotKey As Boolean
    Dim objControl As Object
    Dim strKey As String, arrTag, strNote As String
    Dim arrSQL
    On Error GoTo ErrHand
    '����˵��:ֻ�д�ӡ��ͷ�������ǺͿ�Ƭѡ�����,���������п������ڴ�����,Ҳ�����ǲ��ڴ�����
    
    If Control.ID = conMenu_File_Exit Then
        Unload Me
        Exit Sub
    End If
    
    '����Ǳ�ע�˵�,ִ���꼴�˳�
    If Control.ID > conMenu_��ע1 And Control.ID < conMenu_��ע���� Then
        If Not LocatePatiRecord Then Exit Sub
        mrsBedInfo.Filter = "����ID=" & mrsPatiInfo!����ID & " And ����=0"
        If mrsBedInfo.RecordCount = 0 Then
            mrsBedInfo.Filter = ""
            Exit Sub
        End If
        arrTag = Split(Control.Category, "|")
        str���� = mrsBedInfo!����
        int��λ���� = mrsBedInfo!��Ƭ����
        strKey = ""
        If Val(arrTag(0)) = 1 And Nvl(mrsBedInfo!���Ա�ע1) <> "" Then
            strKey = Split(mrsBedInfo!���Ա�ע1, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע1, ",")(1)
        ElseIf Val(arrTag(0)) = 2 And Nvl(mrsBedInfo!���Ա�ע2) <> "" Then
            strKey = Split(mrsBedInfo!���Ա�ע2, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע2, ",")(1)
        Else
            If Nvl(mrsBedInfo!���Ա�ע3) <> "" Then
                strKey = Split(mrsBedInfo!���Ա�ע3, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע3, ",")(1)
            End If
        End If
        mrsBedInfo.Filter = ""
        '��������
        arrSQL = Array()
        If arrTag(3) <> 0 And strKey <> "" Then
            '��������ͼ������ɾ��ԭ�е�����,�������õ��鷢���仯
            If strKey <> arrTag(1) & "," & arrTag(2) Then
                mstrSQL = "ZL_������Ǽ�¼_UPDATE(" & Me.cboUnit.ItemData(Me.cboUnit.ListIndex) & "," & Val(mrsPatiInfo.Fields("����ID").Value) & "," & _
                    Val(mrsPatiInfo.Fields("��ҳID").Value) & "," & Split(strKey, ",")(1) & "," & 0 & "," & arrTag(0) & IIf(Val(Split(strKey, ",")(0)) = 0, "", "," & Split(strKey, ",")(0)) & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mstrSQL
            End If
        End If
        mstrSQL = "ZL_������Ǽ�¼_UPDATE(" & Me.cboUnit.ItemData(Me.cboUnit.ListIndex) & "," & Val(mrsPatiInfo.Fields("����ID").Value) & "," & _
                Val(mrsPatiInfo.Fields("��ҳID").Value) & "," & arrTag(2) & "," & arrTag(3) & "," & arrTag(0) & IIf(Val(arrTag(1)) = 0, "", "," & arrTag(1)) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = mstrSQL
        
        For i = 0 To UBound(arrSQL)
            If CStr(arrSQL(i)) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "���²��˱��")
        Next
        
        strKey = arrTag(1) & "," & arrTag(2) & "," & arrTag(3) & "," & arrTag(4)
        strNote = arrTag(5)
        '�����ڲ���¼��
        If Val(arrTag(0)) = 1 Then
            Call Record_Update(mrsBedInfo, "���Ա�ע1|���Ա�ע1����", strKey & "|" & strNote, "����|" & Trim(str����))
        ElseIf Val(arrTag(0)) = 2 Then
            Call Record_Update(mrsBedInfo, "���Ա�ע2|���Ա�ע2����", strKey & "|" & strNote, "����|" & Trim(str����))
        Else
            Call Record_Update(mrsBedInfo, "���Ա�ע3|���Ա�ע3����", strKey & "|" & strNote, "����|" & Trim(str����))
        End If
        '���¿�Ƭ
        Call SetCardLabel(int��λ����)
        
        Exit Sub
    End If
    
    strPrivs_������� = GetInsidePrivs(Enum_Inside_Program.p�������)
    strPrivs_���� = GetInsidePrivs(Enum_Inside_Program.p�����¼����)
    
    '��ݼ���ʽ����,������Ϊ��(ֻ���ǲ������������µĹ��ܲ˵�)
    If Control.Parent Is Nothing Then
        Select Case Control.ID
        '61762:������,2013-05-20,���ӷ�����ҺҩƷҽ���Ĺ���
        Case conMenu_Edit_PreBalance, conMenu_Edit_Audit, conMenu_Edit_Send, conMenu_Edit_SendInfusion, conMenu_Report_Reports, conMenu_Report_DrugQuery, conMenu_Edit_SendBack, _
             conMenu_File_PrintMultiBill, conMenu_Edit_BatExecute, conMenu_Edit_AnimalHeat, conMenu_Edit_NurseLogFile
             strParentTitle = "������������"
        End Select
    Else
        strParentTitle = Control.Parent.Title
    End If
    If strParentTitle = "�Ҽ��˵�" Then
        Select Case Control.ID
        Case conMenu_Edit_ReStop, conMenu_Manage_ReportLisView
            strParentTitle = "ҽ��ҵ��"
        Case conMenu_Edit_Billing, conMenu_Edit_ReBillingApply
            strParentTitle = "����ҵ��"
        End Select
    End If
    
    '��Ҳ˵�
    If Control.ID > conMenu_Tool_PlugIn_Item And Control.ID < conMenu_Tool_PlugIn_Item + 100 Then '��ҹ���ִ��
        If Not mobjPlugIn Is Nothing Then
            If Not LocatePatiRecord Then
                Call mobjPlugIn.ExecuteFunc(glngSys, P�°滤ʿվ, Control.Parameter, 0, 0, 0, , 1)
            Else
                Call mobjPlugIn.ExecuteFunc(glngSys, P�°滤ʿվ, Control.Parameter, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, , 1)
            End If
        End If
    End If
    
    '��������˵�
    If strParentTitle <> "" Then
        '����ݼ�ִ�й���ʱ������İ�ť����Ӧ���ǿؼ��Զ������ģ�û�и�����
        
        If strParentTitle = "������������" Then
            '54409:������,2012-09-25,������������û��ѡ����Ҳ����ʹ��(��������������)
            Select Case Control.ID
            Case conMenu_Edit_PreBalance                'Ԥ����
                If LocatePatiRecord Then
                    Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
                Else
                    Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, 0, 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
                End If
            Case conMenu_File_PrintMultiBill            '�߿�����£�
                Call mclsFeeQuery.zlPatiPressMoney(Me, gcnOracle, glngSys, mlngModul, gstrDBUser, mstrPrivs, cboUnit.ItemData(cboUnit.ListIndex), Split(cboUnit.Text, "-")(1))
            Case conMenu_Edit_BatExecute, conMenu_Manage_ThingAudit 'ִ�еǼǣ��£���ִ�к˶�
                If Not LocatePatiRecord Then mrsPatiInfo.Filter = ""
                If mrsPatiInfo.RecordCount > 0 Then
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                Else
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, 0, 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                End If
            Case conMenu_Edit_AnimalHeat                '����¼�����µ����£�
                On Error Resume Next
                Dim strDLL As String
                Dim strSQL As String
                Dim objChart As Object
                Dim rsTemp As New ADODB.Recordset
                
                strSQL = " Select �²��� From ���²��� Where Nvl(����,0)=1"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���²���")
                If err <> 0 Then
                    strDLL = "zl9TemperatureChart"
                Else
                    If rsTemp.RecordCount = 0 Then
                        strDLL = "zl9TemperatureChart"
                    Else
                        strDLL = Nvl(rsTemp!�²���, "zl9TemperatureChart")
                    End If
                End If
                
                err = 0
                strDLL = strDLL & ".clsBodyEditor"
                Set objChart = CreateObject(strDLL)
                If err <> 0 Then
                    MsgBox "    �������²���ʧ�ܣ�" & vbCrLf & "    ���򽫴�����׼�����²�����������չ�֣�����ָ�������²����Ƿ���ڻ����𻵣�" & vbCrLf & "    ��ϸ����" & err.Description, vbInformation, gstrSysName
                    
                    '�������ָ�������²��������򴴽���׼�����²�������Ϊ���ﲻ����Ļ���������ܴ���ֱ��ʹ�����²����еĶ��󣬴Ӷ����³������
                    strDLL = "zl9TemperatureChart.clsBodyEditor"
                    Set objChart = CreateObject(strDLL)
                End If
                
                On Error GoTo ErrHand
                Call objChart.InitBodyEditor(glngSys, gcnOracle)
                Call objChart.BodyMutilEditor(Me, cboUnit.ItemData(cboUnit.ListIndex), strPrivs_����, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
            Case conMenu_Edit_NurseLogFile              '����¼���¼�����£�
                Call mclsTends.TendFileMutilEditor(Me, cboUnit.ItemData(cboUnit.ListIndex), strPrivs_����, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
            Case conMenu_����������                   '�����������£�
                Call InNurseRoutine
            Case conMenu_ProveCollect                   '����ɼ�����վ
                If mobjProveCollect Is Nothing Then
                    On Error Resume Next
                    Set mobjProveCollect = CreateObject("zl9LisWork.clsLisWork")
                    If err <> 0 Then Exit Sub
                End If
                On Error GoTo ErrHand
                Call mobjProveCollect.CodeMan(glngSys, 1211, gcnOracle, Me, gstrDBUser)
            Case conMenu_Edit_BatUnPack '�������
                mclsAdvices.zlCompoundUnpack Me, cboUnit.ItemData(cboUnit.ListIndex), mlng����ID, cboUnit.ItemData(cboUnit.ListIndex)
            Case conMenu_Tool_RisPrintBat '������ӡԤԼ��
                mclsAdvices.AdviceRisReport Me, cboUnit.ItemData(cboUnit.ListIndex)
            Case Else   'ҽ��У�ԡ�ҽ�����͡�ҽ����ͣ��ҽ�����á�ҽ��ȷ��ֹͣ���������ñ�����ӡִ�е����������ջ�(conMenu_Edit_Audit, conMenu_Edit_Send,conMenu_Edit_Pause,conMenu_Edit_Reus,conMenu_Edit_ReStop, conMenu_Report_Reports, conMenu_Report_DrugQuery, conMenu_Edit_SendBack)
                If Not LocatePatiRecord Then mrsPatiInfo.Filter = ""
                Call mclsAdvices.SetFontSize(IIf(mbytFontSize = 12, 1, 0))
                If mrsPatiInfo.RecordCount = 0 Then
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, 0, 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 0, 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                Else
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, True, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                End If
            End Select
            blnExecuted = True
        ElseIf strParentTitle = "ҽ��ҵ��" Then
            If Control.ID = conMenu_View_Notify Then
                With frmNotify
                    .mintNotify = mintNotify
                    .mintNotifyDay = mintNotifyDay
                    .mstrNotifyAdvice = mstrNotifyAdvice
                End With
                frmNotify.mblnFirst = True
            Else
                If Not LocatePatiRecord Then Exit Sub
                If Control.ID = conMenu_�鿴ҽ�� Then
                    Call InNurseRoutine
                Else
                    Call mclsAdvices.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1)
                End If
            End If
            blnExecuted = True
        ElseIf strParentTitle = "����ҵ��" Then
            If Control.ID <> conMenu_Manage_Change_ReCalcFee Then
                If Not LocatePatiRecord Then Exit Sub
                If Control.ID = conMenu_�鿴���� Then
                    Call InNurseRoutine("����")
                Else
                    Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
                End If
                blnExecuted = True
            End If
        ElseIf strParentTitle = "����ҵ��" Or strParentTitle = "����ҵ��" Then
            Call InNurseRoutine(Mid(strParentTitle, 1, 2))
            blnExecuted = True
        End If
    End If
    If blnExecuted Then Exit Sub
    
    Select Case Control.ID
    '---------------------------------------------------------------
    '����˵����������ת
    Case conMenu_Manage_Change_In
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        If mrsPatiInfo!���� = ptת��������ס Then
            mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.Eת������ס, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), "", 0)
        ElseIf mrsPatiInfo!���� = ptת�ƴ���ס Then
            mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E��ס, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), "", _
                    Val(mrsPatiInfo.Fields("����ID").Value), 1)
        Else
            mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E��ס, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), "", _
                    Val(mrsPatiInfo.Fields("����ID").Value), 0)
        End If
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_Turn
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.Eת��, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_TurnUnit
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.Eת����, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_TurnTeam
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.Eתҽ��С��, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_Bed
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E����, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, "", "")
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_TransposeBed
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E��λ�Ի�, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), Nvl(mrsPatiInfo.Fields("����").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_House
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E����, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 1, "", "")
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_Out
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E��Ժ, Me, strPrivs_�������, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_InPati
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.EתΪסԺ, Me, strPrivs_�������, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), _
        Val(mrsPatiInfo.Fields("סԺ��").Value), CStr(mrsPatiInfo.Fields("����").Value))
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Change_BedGrid
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E���Ĵ�λ�ȼ�, Me, strPrivs_�������, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), _
        Trim(CStr(Nvl(mrsPatiInfo.Fields("����").Value))))
    Case conMenu_Manage_Change_PatiInfo
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E����������Ϣ, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
    Case conMenu_Manage_Change_PaitNote
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E���˱�ע�༭, Me, strPrivs_�������, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
    Case conMenu_Manage_Change_Baby
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E�������Ǽ�, Me, strPrivs_�������, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
    Case conMenu_Manage_Change_ReCalcFee
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E�������, Me, strPrivs_�������, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), _
        CStr(mrsPatiInfo.Fields("����").Value))
    Case conMenu_Manage_Change_InsureSel
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.Eҽ������ѡ��, Me, strPrivs_�������, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), Val(mrsPatiInfo.Fields("����").Value))
    Case conMenu_Manage_Change_Undo * 10 + 1
        If Not LocatePatiRecord Then Exit Sub
        If CheckBabyInOut Then Exit Sub
        mblnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E����, Me, strPrivs_�������, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), Val(mrsPatiInfo.Fields("����").Value), Control.Caption)
        Call RefreshPatiList_Rountine
    Case conMenu_Manage_Monitor '�໤��
        Call InNurseRoutine("�໤")
    '---------------------------------------------------------------
    
    '��������
    Case conMenu_Tool_Archive '���Ӳ�������
        If Not LocatePatiRecord Then Exit Sub
        Call frmArchiveView.ShowArchive(Me, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
    Case conMenu_View_Warrant '������Ϣ����
        If Not LocatePatiRecord Then Exit Sub
        Call frmPatiSurety.ShowMe(Me, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
    Case conMenu_Tool_Reference_1 '������ϲο�
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '���ƴ�ʩ�ο�
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    Case conMenu_Manage_FeeItemSet  '������Ŀ��������
        Call Set������Ŀ��������
    Case conMenu_Tool_MedRecAuditResponse '��鷴��
        '�����Ե��ã����ٿ��Բ鿴(��ǰ����ʷ)
        Call lbl���_Click
'    Case conMenu_Tool_UnitSubject '�����������
'         Call frmUnitSubjectSet.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex), mstrPrivs)
'         If gblnOK Then mblnRefresh = True
    Case conMenu_Tool_UnitNBoard
        If frmNoticeBoardSet.ShowMe(Me, mstrPrivs, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
            If Not mfrmNoticeBoard Is Nothing Then
                If mfrmNoticeBoard.mblnShow = True Then Call mfrmNoticeBoard.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex))
            End If
        End If
    '��������
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For Each objControl In Me.cbsMain(2).Controls
            If objControl.ID <> conMenu_View_Find Then
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            End If
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_FontSize_S      '��׼��Ƭ С����
        mlngSource = 999
        lbl����(mlngSource).Tag = lbl����(0).Tag
        Call SetSourceCardH
        mblnRefresh = True
        Call SetFontSize(0)
    Case conMenu_View_FontSize_L      '��Ƭ ������
        mlngSource = 0
        lbl����(mlngSource).Tag = lbl����(999).Tag
        Call SetSourceCardH
        mblnRefresh = True
        Call SetFontSize(1)
    Case conMenu_View_Expend_AllCollapse    '��Ƭ�۵�
        mblnCardCollapse = mblnCardCollapse Xor True
        Call SetSourceCardH
        mblnRefresh = True
    Case conMenu_View_Expend_CurCollapse      '���ڴ�����
        picList.Visible = picList.Visible Xor True
        PatiPage.Visible = picList.Visible
        Call picPatiIn_Resize
        If picList.Visible Then
            fra���.Left = picList.Width - fra���.Width
            fra���.Top = PicDraw.Top + picList.Top + 50
        Else
            fra���.Left = stbThis.Width - fra���.Width - 1500
            fra���.Top = stbThis.Top + 50
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
    Case conMenu_View_Append '��ʾ�����
        lbl����(mlngSource).Tag = Val(lbl����(mlngSource).Tag) Xor 1
        With mrsBedInfo
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                If ISShowCard Then
                    lbl����(!��Ƭ����).Caption = IIf(Val(lbl����(mlngSource).Tag) = 1, IIf(Trim(Nvl(!�����)) = "", "", Trim(!�����)) & IIf(IsNumeric(Trim(!�����)), "_", ""), "") & Trim(!����)
                    lbl�����(!��Ƭ����).Caption = lbl����(!��Ƭ����).Caption
                    lbl����(!��Ƭ����).Caption = Mid(lbl����(!��Ƭ����).Caption, 1, 7)
                End If
                .MoveNext
            Loop
        End With
    Case conMenu_View_NoticBoard
        If cboUnit.ListIndex = -1 Then Exit Sub
        '��ģ̬��ʾ����������
        If mfrmNoticeBoard Is Nothing Then
            Set mfrmNoticeBoard = New frmNoticeBoard
        End If
        
        Call mfrmNoticeBoard.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex))
    Case conMenu_View_Notify 'ҽ������
            With frmNotify
                .mintNotify = mintNotify
                .mintNotifyDay = mintNotifyDay
                .mstrNotifyAdvice = mstrNotifyAdvice
            End With
            frmNotify.mblnFirst = True
    Case conMenu_View_Refresh 'ˢ��
        mblnRefresh = True
        'ˢ��ҽ������
        With frmNotify
            .mintNotify = mintNotify
            .mintNotifyDay = mintNotifyDay
            .mstrNotifyAdvice = mstrNotifyAdvice
            .mblnFirst = True
        End With
    Case conMenu_File_Parameter '��������
        frmSublimeStationSetup.mstrPrivs = mstrPrivs
        Call frmSublimeStationSetup.ShowMe
        If gblnOK Then
            Call GetLocalSetting
            mblnRefresh = True
            'ˢ��ҽ������
             With frmNotify
                .mintNotify = mintNotify
                .mintNotifyDay = mintNotifyDay
                .mstrNotifyAdvice = mstrNotifyAdvice
                .mblnFirst = True
            End With
        End If
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case conMenu_File_PrintBedCard          '��ӡ��ͷ��
        If Not LocatePatiRecord Then Exit Sub
        Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
    Case conMenu_Manage_Print_Label '��ӡ���
        If Not LocatePatiRecord Then Exit Sub
        If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me, "����ID=" & Val(mrsPatiInfo.Fields("����ID").Value), "��ҳID=" & Val(mrsPatiInfo.Fields("��ҳID").Value), 2)
        End If
    Case conMenu_File_PrintDayDetail        'һ���嵥
        If Not LocatePatiRecord Then Exit Sub
        Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
    Case conMenu_File_PrintPageSet          '��ӡ��ҳ����
        If Not LocatePatiRecord Then Exit Sub
        Call mclsFeeQuery.zlExecuteCommandBarsDirect(Control, Me, mstrPrivs, False, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), Val(mrsPatiInfo.Fields("����ID").Value), 0, cboUnit.ItemData(cboUnit.ListIndex), 1, False)
    Case conMenu_File_MedRecSetup '��ҳ��ӡ����
        Call PrintInMedRec(mclsInOutMedRec, 0, IIf(mlng����ID = 0, -1, 0), mlng��ҳID, mobjReport, Val(mrsPatiInfo.Fields("����ID").Value), Me)
    Case conMenu_File_MedRecPreview '��ҳԤ��
        If Not LocatePatiRecord Then Exit Sub
        Call PrintInMedRec(mclsInOutMedRec, 1, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), mobjReport, Val(mrsPatiInfo.Fields("����ID").Value), Me)
    Case conMenu_File_MedRecPrint '��ҳ��ӡ
        If Not LocatePatiRecord Then Exit Sub
        Call PrintInMedRec(mclsInOutMedRec, 2, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value), mobjReport, Val(mrsPatiInfo.Fields("����ID").Value), Me)
    '54621:������,2013-02-28,��ʿվ�����ҳ������
    Case conMenu_Tool_MedRec '��ҳ����
        If Not LocatePatiRecord Then Exit Sub
        Call ExecuteEditMediRec
'    Case conMenu_View_FindNext '������һ��
'        If txtFind.Text = "" Then
'            txtFind.SetFocus
'        Else
'            Call ExecuteFindPati(True)
'        End If
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 5 '���ҷ�ʽ
        mintFindType = Val(Right(Control.ID, 2)) - 1
        cbsMain.RecalcLayout
        txtFind.Text = ""
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
    Case conMenu_View_FindType * 100# + 9
        mintFindType = Val(Right(Control.ID, 2)) - 1
        cbsMain.RecalcLayout
        txtFind.Text = ""
        Call ExecuteFindPati
    Case conMenu_View_FindType * 100# + 11 To conMenu_View_FindType * 100# + 14 '���ҷ�ʽ
        mintPatiInputType = Val(Right(Control.ID, 2)) - 1
        cbsMain.RecalcLayout
        txtסԺ��.Text = ""
        If pic��Ժ����.Enabled And pic��Ժ����.Visible Then pic��Ժ����.SetFocus
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            strTmp = Split(Control.Parameter, ",")(1)
            If strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1132" Then 'סԺ�����ձ�
                If Not LocatePatiRecord Then Exit Sub
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                         "����=" & cboUnit.ItemData(cboUnit.ListIndex), "����ID=" & Val(mrsPatiInfo.Fields("����ID").Value), "��ҳID=" & Val(mrsPatiInfo.Fields("��ҳID").Value))
            ElseIf strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Or strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then    '������ҳ�ʹ߿��
                Call mclsFeeQuery.zlExecuteCommandBars(Control)
            Else
                If Not LocatePatiRecord Then Exit Sub
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                    "����ID=" & Val(mrsPatiInfo.Fields("����ID").Value), "��ҳID=" & Val(mrsPatiInfo.Fields("��ҳID").Value), "סԺ��=" & CStr(mrsPatiInfo.Fields("סԺ��").Value), "���˲���=" & cboUnit.ItemData(cboUnit.ListIndex), _
                    "���˿���=" & Val(mrsPatiInfo.Fields("����ID").Value), "����=" & Nvl(mrsPatiInfo.Fields("����").Value))
            End If
        ElseIf Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 1, conMenu_File_MedRecPreview * 100# + 4) Then
            Call PrintInMedRec(mclsInOutMedRec, IIf(Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6), 2, 1), mlng����ID, mlng��ҳID, mobjReport, mPatiInfo.����ID, Me, Val(Mid(Control.ID & "", Len(Control.ID & ""))))
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
    pic��Ժ����.ZOrder 0
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò�����صĲ˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean, strPrivs As String

    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Parameter = "���ж�" Then Exit Sub

    blnVisible = True
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p�������)
    
    Select Case Control.ID
        Case conMenu_Manage_Change_In
            blnVisible = strPrivs <> ""
        Case conMenu_Manage_Change_Out
            blnVisible = InStr(strPrivs, "���˳�Ժ") > 0
        Case conMenu_Manage_Change_Turn
            blnVisible = InStr(strPrivs, "����ת��") > 0
        Case conMenu_Manage_Change_Bed, conMenu_Manage_Change_TransposeBed, conMenu_Manage_Change_House
            blnVisible = InStr(strPrivs, "����") > 0
        Case conMenu_Manage_Change_TurnUnit
            blnVisible = InStr(strPrivs, "ת����") > 0
        Case conMenu_Manage_Change_PatiInfo
            blnVisible = InStr(strPrivs, "����������Ϣ") > 0
        Case conMenu_Manage_Change_Baby
            blnVisible = InStr(strPrivs, "�������Ǽ�") > 0
        Case conMenu_Manage_Change_ReCalcFee
            blnVisible = InStr(strPrivs, "�������") > 0
        Case conMenu_Manage_Change_BedGrid
            blnVisible = InStr(strPrivs, "������λ�ȼ�") > 0
        Case conMenu_Manage_Change_InPati
            blnVisible = InStr(strPrivs, "סԺ����תסԺ") > 0
    End Select

    Control.Visible = blnVisible
    Control.Parameter = "���ж�"
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean, blnSelect As Boolean, blnWaitIn As Boolean, blnWriteMedRec As Boolean
    Dim blnOut As Boolean, blnPreOut As Boolean, blnOutTo As Boolean, lngType As Long, strPrivs As String
    Dim strCustom As String
    
    If Not mblnStart Then Exit Sub
    If blnUnload Then Exit Sub

    blnSelect = LocatePatiRecord
    If blnSelect Then
        lngType = Val(mrsPatiInfo.Fields("����").Value)
        blnWaitIn = lngType = ptת�ƴ���ס Or lngType = pt��Ժ����ס Or lngType = ptת��������ס
        blnOut = lngType = pt��Ժ
        blnPreOut = lngType = ptԤ��
        '85200:�������ת��ҳ��Ĳ��˲����������ز������磺��������
        blnOutTo = lngType = pt���ת��
    End If
    
    '��ҳ����
    If Between(Control.ID, conMenu_File_MedRecPrint * 100# + 3, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 3, conMenu_File_MedRecPreview * 100# + 4) Then
        If mintMecStandard = 0 Or mintMecStandard = 3 Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    End If
    
    If Control.Category = "����" Then
        Exit Sub    '��cbsMain_InitCommandsPopup������,�˳������Ӵ���������ɼ���
    ElseIf Control.Category = "����" Then
        
        Call SetControlVisible(Control)
        If Not Control.Visible Then Exit Sub
        
        strPrivs = GetInsidePrivs(Enum_Inside_Program.p�������)
        If InStr(strPrivs, "���в���") = 0 Then
            If InStr("," & mstrUnits & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Control.Enabled = False: Exit Sub
        End If
    End If
    
    '���ӳ������Ȩ�����ò˵����ܵ�״̬
    strCustom = ""
    If Not Control.Parent Is Nothing Then
        strCustom = Control.Parent.Title
    End If
    If strCustom <> "" Then
        If strCustom = "�Ҽ��˵�" Then
            Select Case Control.ID
            Case conMenu_Edit_ReStop, conMenu_Manage_ReportLisView
                strCustom = "ҽ��ҵ��"
            Case conMenu_Edit_Billing, conMenu_Edit_ReBillingApply
                strCustom = "����ҵ��"
            End Select
        End If
        If strCustom = "ҽ��ҵ��" Then
            If Control.ID = conMenu_View_Notify Then
                Control.Enabled = True
            Else
                Call mclsAdvices.zlCheckPrivs(Control, 1)
                Control.Enabled = Control.Visible And blnSelect
                '50906:������,2012-09-18,��Ժ����ס���˸��ݲ���"���������ס�����´�ҽ��"�����Ƿ�����¿�ҽ��
                If Control.ID = conMenu_Edit_NewItem And Control.Enabled = True And lngType = pt��Ժ����ס Then
                    Control.Enabled = (Val(zlDatabase.GetPara("���������ס�����´�ҽ��", glngSys, pסԺҽ���´�, 1)) = 1)
                End If
            End If
            Exit Sub
        ElseIf strCustom = "����ҵ��" Then
            Call mclsFeeQuery.zlCheckPrivs(Control)
            Control.Enabled = Control.Visible And blnSelect
            
            If Control.ID = conMenu_Edit_PreBalance And Control.Enabled = True Then
                Control.Enabled = blnSelect And Nvl(mrsPatiInfo.Fields("����").Value, 0) <> 0
            ElseIf Control.ID = conMenu_Manage_Change_ReCalcFee And Control.Enabled = True Then
                Control.Enabled = blnSelect And Nvl(mrsPatiInfo.Fields("����").Value, 0) = 0
            End If
            Exit Sub
        ElseIf strCustom = "����ҵ��" Then
            Control.Enabled = (GetInsidePrivs(p�����¼����, True) <> "") And blnSelect
        ElseIf strCustom = "����ҵ��" Then
            Control.Enabled = (GetInsidePrivs(pסԺ��������, True) <> "") And blnSelect
        ElseIf strCustom = "������������" Then
            '54409:������,2012-09-25,������������û��ѡ����Ҳ����ʹ��(��������������)
            Select Case Control.ID
            Case conMenu_Edit_PreBalance                'Ԥ����
                Control.Enabled = True ' blnSelect
            '61762:������,2013-05-20,���ӷ�����ҺҩƷҽ���Ĺ���
            Case conMenu_Edit_Audit, conMenu_Edit_Send, conMenu_Edit_SendInfusion, conMenu_Edit_Pause, conMenu_Edit_Reuse, conMenu_Edit_ReStop 'ҽ��У�ԡ�ҽ�����͡�������ҺҩƷҽ����ҽ����ͣ��ҽ�����á�ҽ��ȷ��ֹͣ
                Call mclsAdvices.zlCheckPrivs(Control, 1)
                 'Control.Enabled = Control.Visible And blnSelect
                If Not mrsPatiInfo Is Nothing Then
                    If mrsPatiInfo.State = adStateOpen Then
                        If blnSelect = False Then mrsPatiInfo.Filter = ""
                        Control.Enabled = Control.Visible And (mrsPatiInfo.RecordCount > 0)
                    End If
                End If
            Case conMenu_File_PrintMultiBill            '�߿�����£�
                Control.Visible = InStr(1, ";" & mstrPrivs & ";", ";���˴߿��;")
                Control.Enabled = Control.Visible
            Case conMenu_Edit_BatExecute                   'ִ�еǼǣ��£�
                '60781:������,2013-07-15
                'Call mclsAdvices.zlCheckPrivs(Control, 1)
                Control.Visible = (InStr(GetInsidePrivs(pסԺҽ������), ";����ִ�еǼ�;") > 0)
                Control.Enabled = Control.Visible
            Case conMenu_Edit_AnimalHeat                '����¼�����µ����£�
                Control.Visible = InStr(1, GetInsidePrivs(p�����¼����, True), ";���µ���ͼ;")
                Control.Enabled = Control.Visible
            Case conMenu_Edit_NurseLogFile              '����¼���¼�����£�
                Control.Visible = InStr(1, GetInsidePrivs(p�����¼����, True), ";�����¼�Ǽ�;")
                Control.Enabled = Control.Visible
            Case conMenu_Manage_ThingAudit, conMenu_Report_DrugQuery, conMenu_Edit_Surplus, conMenu_Report_Reports, conMenu_Edit_SendBack                '��ҩ��ѯ,����Ǽ�,��ӡִ�е�,�����ջ�
                Call mclsAdvices.zlCheckPrivs(Control, 1)
                Control.Enabled = Control.Visible
            Case conMenu_ProveCollect
                Control.Visible = mstrPrivs_����ɼ� <> ""
                Control.Enabled = Control.Visible
            Case conMenu_����������                   '�����������£�
                Control.Enabled = blnSelect
            End Select
            Exit Sub
        End If
    End If
    
    Select Case Control.ID
    Case conMenu_Manage_Change_Undo
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOutTo
        If Control.Enabled = True Then
            Control.Enabled = Val(Nvl(mrsPatiInfo.Fields("��ҳID").Value, 0)) = Val(Nvl(mrsPatiInfo.Fields("�����ҳId").Value, 0))
        End If
    Case conMenu_Manage_Change_In
        Control.Enabled = blnWaitIn
    Case conMenu_Manage_Change_InPati
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.���� = 2
        End If
    'ת�ƣ�����������������������Ϣ���������,ת������תС��,��λ�Ի�
    Case conMenu_Manage_Change_Turn, conMenu_Manage_Change_Bed, conMenu_Manage_Change_House, _
         conMenu_Manage_Change_PatiInfo, conMenu_Manage_Change_ReCalcFee, conMenu_Manage_Change_TurnUnit, _
         conMenu_Manage_Change_TurnTeam, conMenu_Manage_Change_TransposeBed
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = mrsPatiInfo.Fields("״̬").Value <> 2
            
            If Control.ID = conMenu_Manage_Change_TransposeBed Then '��λ�Ի�
                Control.Enabled = Trim(CStr(mrsPatiInfo.Fields("����").Value)) <> ""
            ElseIf Control.ID = conMenu_Manage_Change_ReCalcFee Then
                Control.Enabled = Nvl(mrsPatiInfo.Fields("����").Value, 0) = 0
            End If
        End If
    Case conMenu_Manage_Change_InsureSel
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = Nvl(mrsPatiInfo.Fields("����").Value, 0) <> 0
        End If
    Case conMenu_Manage_Change_BedGrid
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = Trim(Nvl(mrsPatiInfo.Fields("����").Value)) <> "" And mrsPatiInfo.Fields("״̬").Value <> 2
        End If
    Case conMenu_Manage_Change_Out
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = (InStr(1, "," & pt��Ժ & ",3.1,", mrsPatiInfo.Fields("����").Value) <> 0 Or blnPreOut) And mrsPatiInfo.Fields("״̬").Value <> 2
        End If
    Case conMenu_Manage_Change_Baby
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.����
        End If
    Case conMenu_Manage_Change_PaitNote
            Control.Enabled = Not blnOutTo
    Case conMenu_Manage_Monitor '�໤��
        Control.Visible = mblnMonitor And (InStr(GetInsidePrivs(pסԺ��ʿվ), "����໤") > 0)
        Control.Enabled = False
        If blnSelect Then
            mrsBedInfo.Filter = "����='" & mrsPatiInfo!���� & "'"
            If mrsBedInfo.RecordCount <> 0 Then
                Control.Enabled = Nvl(mrsBedInfo!�໤��, 0) > 0
            End If
            mrsBedInfo.Filter = ""
        End If
    Case conMenu_Tool_Archive '���Ӳ�������
        Control.Visible = GetInsidePrivs(p���Ӳ�������) <> ""
        Control.Enabled = Control.Visible And blnSelect
    Case conMenu_View_Warrant '������Ϣ����
        Control.Enabled = blnSelect
    Case conMenu_Tool_Reference_1 '������ϲο�
        If GetInsidePrivs(p������ϲο�) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 'ҩƷ�����Ʋο�
        If GetInsidePrivs(pҩƷ���Ʋο�) = "" Then Control.Visible = False
    Case conMenu_Tool_MedRecAuditResponse '��鷴��
        '�����Ե��ã����ٿ��Բ鿴(��ǰ����ʷ)
        Control.Enabled = blnSelect
    Case conMenu_Manage_Print_Label '��ӡ���
        Control.Visible = InStr(mstrPrivs, ";�����ӡ;")
        If blnSelect = True Then
            Control.Enabled = mintREPORTSEL <> ҳ��.��Ժ
        End If
        
    Case conMenu_File_MedRec '��ҳ��ӡ
        Control.Visible = InStr(mstrPrivs, "��ӡ��ҳ")
    '54621:������,2013-02-28,��ʿվ�����ҳ������
    Case conMenu_Tool_MedRec '��ҳ����
        blnWriteMedRec = Val(zlDatabase.GetPara("ҽ���ͻ�ʿ�ֱ���д������ҳ", glngSys, pסԺҽ��վ, "0")) = 1
        Control.Enabled = blnSelect And blnWriteMedRec
        Control.Visible = blnWriteMedRec
    Case conMenu_File_Parameter '��������
        'If InStr(mstrPrivs, "��������") = 0 Then Control.Visible = False
'    Case conMenu_Tool_UnitSubject '�����������
'        Control.Visible = InStr(1, ";" & mstrPrivs & ";", ";�����������;")
'        Control.Enabled = Control.Visible
    Case conMenu_Tool_UnitNBoard
        Control.Visible = InStr(1, ";" & mstrPrivs & ";", ";��������������;")
        Control.Enabled = Control.Visible
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
        Call cbsMain_Resize
    Case conMenu_View_FontSize_S      '��׼��Ƭ С����
        Control.Checked = (mlngSource = 999)
    Case conMenu_View_FontSize_L      '��Ƭ ������
        Control.Checked = (mlngSource = 0)
    Case conMenu_View_Expend_AllCollapse    '��Ƭ�۵�
        Control.Checked = mblnCardCollapse
    Case conMenu_View_Expend_CurCollapse
        Control.Checked = picList.Visible
    Case conMenu_View_Append
        Control.Checked = (Val(lbl����(mlngSource).Tag) = 1)
    Case conMenu_View_FindType '���ҷ�ʽ
        Control.Enabled = True
        Control.Caption = "����" & Decode(mintFindType, 0, "����", 1, "סԺ��", 2, "���￨", 3, "����", 4, "����", 8, "����") & "����"
        txtFind.PasswordChar = IIf(mintFindType = 2 And gblnCardHide, "*", "")
        
        '��Ժ���˲��ҷ�ʽ
        lblPatiInputType.Caption = Decode(mintPatiInputType, 10, "�� ��", 11, "סԺ��", 12, "���￨", 13, "�� ��", "�� ��") & "��"
        txtסԺ��.PasswordChar = IIf(mintPatiInputType = 2 And gblnCardHide, "*", "")
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '���ҷ�ʽ
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
    Case conMenu_View_FindType * 100# + 11 To conMenu_View_FindType * 100# + 14 '��Ժ���˲��ҷ�ʽ
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintPatiInputType
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '��ҹ���ִ��
        Control.Enabled = True 'blnSelect
    End Select
    
End Sub

Private Sub GetLocalSetting()
'���ܣ���ע����ȡ��Ժ���˵�ʱ�䷶Χ
    Dim curDate As Date, intDay As Integer

    '������ʾ��Χ
    mstrScope = "11111"
    mintChange = Val(zlDatabase.GetPara("���ת������", glngSys, pסԺ��ʿվ, 7))
    'ת����������
    txtChange.Text = mintChange
    
    '��Ժ����ʱ�䷶Χ
'    curDate = zlDatabase.Currentdate
'    intDay = Val(zlDatabase.GetPara("��Ժ���˽������", glngSys, pסԺ��ʿվ, 0))
'    mdtOutEnd = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
'    intDay = Val(zlDatabase.GetPara("��Ժ���˿�ʼ���", glngSys, pסԺ��ʿվ, 0))
'    mdtOutBegin = Format(curDate - intDay, "yyyy-MM-dd 00:00:00")
    
    'ҽ������ˢ������
    mstrNotifyAdvice = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, pסԺ��ʿվ, "0000000")
    mintNotifyDay = Val(zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, pסԺ��ʿվ, 1))
    mintNotify = Val(zlDatabase.GetPara("�Զ�ˢ��ҽ�����", glngSys, pסԺ��ʿվ))
    
    '��Ƭ��ʾ����(���,���)
    mstrCardInfo = zlDatabase.GetPara("��Ƭ��ʾ����", glngSys, pסԺ��ʿվ, "11")
    
    '������鷴������
    mlngMedRedDay = Val(zlDatabase.GetPara("������鷴������", glngSys, pסԺ��ʿվ))
    
    '������ҳ��׼
    mintMecStandard = Val(zlDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0"))
    
    mblnCardBalance = (Val(zlDatabase.GetPara("��Ƭ���������", glngSys, 1265, 0)) = 1)
    '92852:������,2016-01-20,��λ��Ƭ������ʽ,0-��������,1-��λ���Ʊ��+��������
    mblnCardOrder = (Val(zlDatabase.GetPara("��λ��Ƭ����ʽ", glngSys, 1265, 0)) = 0)
    '54370:������,2013-05-02,��Ӳ���"ҽ��У�Ժ��Զ���λ��ҽ��ҳ��"
    mblnCollateAutoFind = (Val(zlDatabase.GetPara("ҽ��������Զ���λ��ҽ��ҳ��", glngSys, 1265, 0)) = 1)
    '����ҳ��ؼ���״̬
    PatiPage.Item(ҳ��.�����).Visible = True
    PatiPage.Item(ҳ��.ת��).Visible = True
    PatiPage.Item(ҳ��.��Ժ).Visible = True
    
    '��ȡ��С��Ч��ҳ�����
    If PatiPage.Item(ҳ��.�����).Visible Then
        mintPage = ҳ��.�����
    ElseIf PatiPage.Item(ҳ��.ת��).Visible Then
        mintPage = ҳ��.ת��
    ElseIf PatiPage.Item(ҳ��.��Ժ).Visible Then
        mintPage = ҳ��.��Ժ
    Else
        mintPage = ҳ��.��ͥ����
    End If
    Call InitColor
End Sub

Private Sub RefreshData()
    Dim rsPati As New ADODB.Recordset
    
    '����ƥ��ʱ����ҳ����������գ�F5ˢ�£�Ӧ�ûָ���һ����ֵ
    If cboUnit.ListIndex = -1 Then Call zlControl.CboSetIndex(cboUnit.hwnd, mintPreDept)
    mblnHavePath = HavePath(cboUnit.ItemData(cboUnit.ListIndex))
    Call init���ڴ��嵥
    mstrBoardKeys = ""
    mblnShow = False        '���⼤��ѡ���¼������¿�Ƭ����������ʾ
    mintREPORTSEL = -1
    mlng����ID = 0:    mlng��ҳID = 0: mlngPre����ID = 0: mlngPre��ҳID = 0
    mlng�մ� = 0: mlng�ڴ� = 0: mlng��Ժ = 0: mlngת�� = 0: mlng��Ժ = 0: mlngԤ��Ժ = 0
    mlngת�� = 0: mlng���� = 0: mlng���� = 0: mlngΣ = 0: mlng�� = 0: mlng�Ҵ� = 0
    
    '1��ʼ���ڴ��¼��
    '61824:������,2013-05-23,��ʾ�����ֱ�־
    Set mrsBedInfo = New ADODB.Recordset
    mstrFields = "��Ƭ����," & adDouble & ",18|����," & adLongVarChar & ",10|סԺ��," & adDouble & ",18|����ID," & adDouble & ",18|" & _
                 "��ҳID," & adDouble & ",18|����," & adLongVarChar & ",10|�໤��," & adDouble & ",18|�������," & adDouble & ",18|" & _
                 "�ٴ�·��," & adDouble & ",18|���Ա�ע1," & adLongVarChar & ",100|����״̬," & adDouble & ",18|���Ա�ע2," & adLongVarChar & ",100|���Ա�ע3," & adLongVarChar & ",100|" & _
                 "�໤������," & adLongVarChar & ",20|�����������," & adLongVarChar & ",20|�ٴ�·������," & adLongVarChar & ",20|" & _
                 "���Ա�ע1����," & adLongVarChar & ",20|����״̬����," & adLongVarChar & ",20|���Ա�ע2����," & adLongVarChar & ",20|���Ա�ע3����," & adLongVarChar & ",20|" & _
                 "����ȼ�," & adDouble & ",18|����ȼ�����," & adLongVarChar & ",20|��������," & adLongVarChar & ",20|" & _
                 "����," & adDouble & ",2|����," & adLongVarChar & ",100|����," & adLongVarChar & ",200|��λ����," & adLongVarChar & ",50|�����," & adLongVarChar & ",20|������," & adLongVarChar & ",10|סԺ����," & adLongVarChar & ",10"
    Call Record_Init(mrsBedInfo, mstrFields)
    
    '��ȡ�����������
    Call LoadNotes
    
    '2װ�ر����������д�λ
    Call ShowGuage("װ�ر����������д�λ", 10)
    'debug.print "װ�ر����������д�λ,Start:" & Now
    If Not LoadBeds And Not mblnStart Then
        Unload Me
        Exit Sub
    End If
    
    '3��ȡ���������в����嵥
    Call ShowGuage("��ȡ���������в����嵥", 20)
    'debug.print "��ȡ���������в����嵥,Start:" & Now
    Call LoadPatients(rsPati)
    
    '4�����ڴ���������
    Call ShowGuage("�����ڴ���������", 30)
    'debug.print "�����ڴ���������,Start:" & Now
    Call UpgradeBeds(rsPati)
    
    '5װ�ز��ڴ�����(��ͥ�����������ѡ�˴��������ش���Ʋ��ˣ��ѳ�Ժ�����ת����ҳ�����ż���)
    Call ShowGuage("װ�ز��ڴ������嵥", 90)
    'debug.print "װ�ز��ڴ�����,Start:" & Now
    
    Dim strField As String, strValue As String
    strField = "����," & adDouble & ",2|����2," & adDouble & ",2|����," & adLongVarChar & ",50|����ID," & adDouble & ",18|��ҳID," & adDouble & ",18|" & _
               "סԺ��," & adDouble & ",18|����," & adLongVarChar & ",20|����," & adLongVarChar & ",200|�Ա�," & adLongVarChar & ",10|����," & adLongVarChar & ",20|����," & adLongVarChar & ",50|" & _
               "����ID," & adDouble & ",18|סԺҽʦ," & adLongVarChar & ",20|���λ�ʿ," & adLongVarChar & ",20|����״̬," & adLongVarChar & ",20|" & _
               "����," & adLongVarChar & ",20|����ȼ�," & adLongVarChar & ",50|�ѱ�," & adLongVarChar & ",50|ҽ�Ƹ��ʽ," & adLongVarChar & ",50|��ǰ����," & adLongVarChar & ",50|" & _
               "��Ժ����," & adLongVarChar & ",20|��Ժ����," & adLongVarChar & ",20|סԺ����," & adLongVarChar & ",20|��Ժ��ʽ," & adLongVarChar & ",20|" & _
               "��������," & adLongVarChar & ",50|״̬," & adLongVarChar & ",10|����," & adDouble & ",18|���￨��," & adLongVarChar & ",20|·��״̬," & adLongVarChar & ",20|" & _
               "��ɫ," & adDouble & ",18|������," & adLongVarChar & ",10|Ӥ������ID," & adDouble & ",18|Ӥ������ID," & adDouble & ",18|�����ҳId," & adDouble & ",18"
    Call Record_Init(mrsPatiInfo, strField)
    
    Call UpgradeList(rsPati)
    '���ǰ���ڴ�ҳ��ĵ���¼�
    If PatiPage.Selected Is Nothing Then
        PatiPage.Item(mintPage).Selected = True
    Else
        If PatiPage.Selected.Visible = False Then
            PatiPage.Item(mintPage).Selected = True
        End If
    End If
    Call PatiPage_SelectedChanged(PatiPage.Selected)
    '����ҳ������
    If GetPatiCount(ҳ��.�����) <> 0 Then PatiPage.Item(ҳ��.�����).Caption = "�����" & GetPatiCount(ҳ��.�����) & "��"
    If GetPatiCount(ҳ��.ת��) <> 0 Then PatiPage.Item(ҳ��.ת��).Caption = "���ת��" & GetPatiCount(ҳ��.ת��) & "��"
    If GetPatiCount(ҳ��.��Ժ) <> 0 Then PatiPage.Item(ҳ��.��Ժ).Caption = "�����Ժ" & GetPatiCount(ҳ��.��Ժ) & "��"
    If GetPatiCount(ҳ��.��ͥ����) <> 0 Then PatiPage.Item(ҳ��.��ͥ����).Caption = "��ͥ����" & GetPatiCount(ҳ��.��ͥ����) & "��"

    Call ShowGuage("���ݶ�ȡ����", 100)
    'debug.print "����,OVER:" & Now
    Call GetInpatientAreaInfo
    
    '6�ٸ����趨��������ʾ��������Ӧ�Ŀ�Ƭ
    Call ShowSelect                 '��Ϊ�ĵ�һ�£����⿨Ƭû����Ϊ���ȴ��ʾ��������
    Call AdjustCard
    
    Call CopyReocrd(rsPati)
    
    Call AddSendCommandBar
End Sub

Private Sub LoadNotes()
    Dim strPatientFilter As String
    Dim blnNext As Boolean, strItems As String
    Dim i As Integer, strKey As String
    On Error GoTo ErrHand
    
    With Me.cbo����
        .Clear
        .AddItem "����"
        .AddItem "�������"
        .AddItem "�ٴ�·��"
        .AddItem "����״̬"
        '��ȡ��ǰ�����趨�ı�ע����
        mstrSQL = "Select nvl(����ID,0) ����ID,�������, ������, Replace(˵��, '|', '') ˵��, ͼ������, ��Ч����" & vbNewLine & _
            " From �����������" & vbNewLine & _
            " Where ����id Is Null Or ����id = [1]" & vbNewLine & _
            " Order By Nvl(����id, 0), �������, ������"
        Set mrsNotes = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ�����������", Me.cboUnit.ItemData(Me.cboUnit.ListIndex))
        strItems = "": strKey = ""
        Do While Not mrsNotes.EOF
            If Val("" & mrsNotes!������) = 0 Then
                blnNext = True
                strKey = mrsNotes!����ID & "-" & mrsNotes!�������
                .AddItem mrsNotes!˵�� & ""
                .ItemData(.NewIndex) = Val(mrsNotes!����ID) + Val(mrsNotes!�������)
                strItems = strItems & "|"
            Else
                If strKey = mrsNotes!����ID & "-" & mrsNotes!������� Then
                    strItems = strItems & IIf(blnNext, "", ",") & mrsNotes!˵�� & "'" & mrsNotes!������
                    blnNext = False
                End If
            End If
            mrsNotes.MoveNext
        Loop
        If mrsNotes.RecordCount <> 0 Then mrsNotes.MoveFirst
        If strItems <> "" Then strItems = Mid(strItems, 2)
        mstrNoteItems = strItems
        
        strPatientFilter = zlDatabase.GetPara("�������", glngSys, 1265, "3")
        .Tag = "�ȴ����,�ܾ����,���ڳ��,�������,��鷴��,��鷴��,�������,�������|δ����,ִ����,������,��������,�������|Ԥת��,Ԥ��Ժ" & IIf(Val(strPatientFilter) = 0, "", ",���" & strPatientFilter & "����") & "|" & strItems
        .ListIndex = 0
    End With
    
    '��ȡ��ǰ�����ı�ע��¼
    'LPF,2014-10-21,�����Ż�:�����Ժ���˱�
    mstrSQL = "" & _
            " Select a.����id, a.��ҳid,nvl(a.���ⲡ��ID,0) ���ⲡ��ID, a.�������, a.������,a.���˳��, a.����, Replace(b.˵��, '|', '') ˵��, b.ͼ������, b.��Ч����, Floor(Sysdate - a.����) As ʵ������" & vbNewLine & _
            " From ������Ǽ�¼ a, ����������� b, ������Ϣ c, ��Ժ���� e" & vbNewLine & _
            " Where a.������� = b.������� And a.������ = b.������ And nvl(a.���ⲡ��ID,0) = nvl(b.����id,0) And a.����id = c.����id And a.��ҳid = c.��ҳid And " & vbNewLine & _
            "      a.����id = c.��ǰ����id And c.����id = e.����id And c.��ǰ����id = e.����id And e.����id = [1] And " & vbNewLine & _
            "      (b.��Ч���� = 0 Or (b.��Ч���� > Floor(Sysdate - a.����)))" & vbNewLine & _
            " Order By a.����id, a.��ҳid,a.���˳��,a.�������"

    Set mrsPatiNotes = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡָ����������Ч��ע��¼", Me.cboUnit.ItemData(Me.cboUnit.ListIndex))
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CopyReocrd(ByVal rsPati As ADODB.Recordset)
    Dim strField As String, strValue As String
    '61824:������,2013-05-23,��ʾ�����ֱ�־
    rsPati.Filter = 0
    If rsPati.RecordCount <> 0 Then rsPati.MoveFirst
    strField = "����|����2|����|����ID|��ҳID|סԺ��|����|����|�Ա�|����|����|����ID|סԺҽʦ|���λ�ʿ|����״̬|����|����ȼ�|�ѱ�|ҽ�Ƹ��ʽ|��ǰ����|��Ժ����|��Ժ����|סԺ����|��Ժ��ʽ|��������|״̬|����|���￨��|·��״̬|��ɫ|������|Ӥ������ID|Ӥ������ID|�����ҳId"
    Do While Not rsPati.EOF
        strValue = rsPati!���� & "|" & rsPati!����2 & "|" & rsPati!���� & "|" & rsPati!����ID & "|" & rsPati!��ҳID & "|" & Nvl(rsPati!סԺ��, 0) & "|" & rsPati!���� & "|" & Nvl(rsPati!����) & "|" & rsPati!�Ա� & "|" & _
                  rsPati!���� & "|" & Nvl(rsPati!����) & "|" & Nvl(rsPati!����ID, 0) & "|" & Nvl(rsPati!סԺҽʦ) & "|" & Nvl(rsPati!���λ�ʿ) & "|" & Nvl(rsPati!����״̬, 0) & "|" & Nvl(rsPati!����) & "|" & _
                  Nvl(rsPati!����ȼ�, "����") & "|" & Nvl(rsPati!�ѱ�) & "|" & Nvl(rsPati!ҽ�Ƹ��ʽ) & "|" & Nvl(rsPati!��ǰ����, "һ��") & "|" & Format(rsPati!��Ժ����, "yyyy-MM-dd") & "|" & Format(rsPati!��Ժ����, "yyyy-MM-dd") & "|" & rsPati!סԺ���� & "|" & rsPati!��Ժ��ʽ & "|" & _
                  Nvl(rsPati!��������, "��ͨ����") & "|" & rsPati!״̬ & "|" & Nvl(rsPati!����, 0) & "|" & Nvl(rsPati!���￨��) & "|" & Nvl(rsPati!·��״̬, 0) & "|" & Nvl(rsPati!��ɫ, 0) & "|" & Nvl(rsPati!������) & "|" & Nvl(rsPati!Ӥ������ID, 0) & "|" & Nvl(rsPati!Ӥ������ID, 0) & "|" & Nvl(rsPati!�����ҳID, 0)
        Call Record_Add(mrsPatiInfo, strField, strValue)
        rsPati.MoveNext
    Loop
End Sub

Private Sub chk�����մ�_Click()
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
    
    '���㵥������
    lngMove = CLng((mdblScaleHeight - (PicDraw.Height - IIf(picList.Visible, picList.Height, 0))) / 100)
    If lngMove < 0 Then lngMove = 0
    lngY = -1 * HScr.Value * lngMove
    If lngY >= 0 And lngY < 100 Then lngY = 100
    Call AdjustCard(lngY, mstrBoardKeys)
End Sub

Private Sub lbl����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, lbl����(Index).Caption, True
End Sub

Private Sub lbl����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, lbl�����(Index).Caption, True
End Sub

Private Sub lblҽʦ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, lblҽʦ(Index).Caption, True
End Sub

Private Sub lbl���_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, "��ϣ�" & lbl���(Index).Caption, True
End Sub

Private Sub picPatiIn_Resize()
    Dim i As Long, y As Long, dblTop As Double
    On Error Resume Next
    If picList.Visible = False Then
        picPara(2).Visible = False
        picPara(3).Visible = False
        pic��Ժ����.Visible = False
        Exit Sub
    Else
        For i = 2 To 3
            picPara(i).Visible = False
        Next
        If PatiPage.Selected.Index = ҳ��.����� Then
            pic��Ժ����.Tag = ҳ��.�����
        ElseIf PatiPage.Selected.Index = ҳ��.��Ժ Then
            picPara(2).Visible = True
            pic��Ժ����.Tag = ҳ��.��Ժ
        ElseIf PatiPage.Selected.Index = ҳ��.ת�� Then
            picPara(3).Visible = True
            pic��Ժ����.Tag = ҳ��.ת��
        ElseIf PatiPage.Selected.Index = ҳ��.��ͥ���� Then
            pic��Ժ����.Tag = ҳ��.��ͥ����
        End If
    End If
    
    picPara(2).Top = 20
    For i = 3 To 3
        picPara(i).Top = picPara(i - 1).Top + IIf(picPara(i - 1).Visible, picPara(i - 1).Height, 0)
        If picPara(i - 1).Visible Then y = y + 1
    Next
    
    If picPara(i - 1).Visible Then y = y + 1
    rptPati(PatiPage.Selected.Index).Top = IIf(y > 0, 20 + TextWidth("��") - 180, 350) + y * (310 + TextWidth("��") - 180)
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
        Case "ZLHIS_PATIENT_001" '��Ժ�����
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            '��ȡ����ID����ҳID������
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            '��鲡��
            strValue = "": Call mclsXML.GetSingleNodeValue("in_dept_id", strValue, xsNumber): lngDept = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("in_area_id", strValue, xsNumber)
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Or lngDept = 0 Then Exit Sub
            
            If Val(strValue) = 0 Then
                strValue = ""
                strSQL = "Select ����ID From �������Ҷ�Ӧ where ����ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lngDept)
                Do While Not rsTmp.EOF
                    strValue = strValue & "," & rsTmp!����ID
                rsTmp.MoveNext
                Loop
                strValue = Mid(strValue, 2)
            End If
            If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
            If FreshPatiCard("��������Ʋ���", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���´�����ס�Ĳ���:" & strName, "������ס����")
            End If
        Case "ZLHIS_PATIENT_002" '��ס
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '��ȡ����ID����ҳID������...
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("in_bed", strValue, xsString): strBed = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("in_area_id", strValue, xsNumber): lngUnit = Val(strValue)
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            If InStr(1, "," & lngUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
            '��鲡���Ƿ�������Ժ
            strSQL = "Select ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2] And NVL(״̬,0)=0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ�������Ժ", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
            If Not mrsPatiInfo.EOF Then
                '��鲡�˴�����Ժ����ס�б���
                If mrsPatiInfo!���� = 0 Then
                    mrsPatiInfo.Delete: mrsPatiInfo.Filter = ""
                    strKey = ""
                    If mintREPORTSEL = ҳ��.����� Then
                        If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                            If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                            End If
                        End If
                    End If
                    rptPati(ҳ��.�����).Records.DeleteAll
                    Call UpgradeList(mrsPatiInfo, ҳ��.�����)
                    PatiPage.Item(ҳ��.�����).Caption = "�����" & GetPatiCount(ҳ��.�����) & "��"
                    If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                Else
                    Exit Sub
                End If
            End If
            If FreshPatiCard("������Ժ����", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "������ס�Ĳ���:" & strName & "   ����:" & strBed, "��ס����")
            End If
            
        Case "ZLHIS_PATIENT_003" 'ת��
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '��ȡ����ID����ҳID������
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("current_bed", strValue, xsString): strBed = strValue
            
            '1��ת����Ҵ�����б�ˢ��
            strValue = "": Call mclsXML.GetSingleNodeValue("current_area_id", strValue, xsNumber): lngCurrUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("change_dept_id", strValue, xsNumber): lngDept = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("change_area_id", strValue, xsNumber): lngUnit = Val(strValue)
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            If Not (lngUnit = 0 And lngDept = 0) Then
                If lngUnit = 0 Then
                    strValue = ""
                    strSQL = "Select ����ID From �������Ҷ�Ӧ where ����ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lngDept)
                    Do While Not rsTmp.EOF
                        strValue = strValue & "," & rsTmp!����ID
                    rsTmp.MoveNext
                    Loop
                    strValue = Mid(strValue, 2)
                Else
                    strValue = lngUnit
                End If
                If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    If FreshPatiCard("����ת������Ʋ���", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
                        If strName <> "" Then
                            Call mclsMipModule.ShowMessage(strMsgItemIdentity, "����ת��Ĳ���:" & strName, "������ס����")
                        End If
                    End If
                End If
            End If
            '2��ת��������Ժ�����б�ˢ��
            If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
            '�����ڴ�����ͼ��
            strSQL = "Select ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2] And NVL(״̬,0)=2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ���Ԥת��״̬", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            
            mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
            blnFresh = False
            Do While Not mrsPatiInfo.EOF
                If InStr(1, ",3,3.1,4,", "," & mrsPatiInfo!���� & ",") <> 0 Then
                    blnFresh = True
                    If mrsPatiInfo!���� = 3.1 Then
                        mrsPatiInfo!״̬ = 2
                    Else
                        mrsPatiInfo!���� = 3.2: mrsPatiInfo!���� = "Ԥת�Ʋ���": mrsPatiInfo!״̬ = 2
                    End If
                    mrsPatiInfo.Update
                End If
            mrsPatiInfo.MoveNext
            Loop
            If blnFresh = False Then Exit Sub
            
            mrsBedInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID & " And ����<>1"
            If Not mrsBedInfo.EOF Then
                intCardIndex = mrsBedInfo!��Ƭ����
                mrsBedInfo!����״̬ = Img���(mlngSource).ListImages("Ԥת��").Index
                mrsBedInfo!����״̬���� = "Ԥת��"
                mrsBedInfo.Update
                Call SetCardLabel(intCardIndex)
            End If
            mrsBedInfo.Filter = 0
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���´�ת���Ĳ���:" & strName & "   ����:" & strBed, "��ת������")
            End If
        Case "ZLHIS_PATIENT_009" 'Ԥ��Ժ
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '��ȡ����ID����ҳID������
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("out_area_id", strValue, xsNumber): lngCurrUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("out_bed", strValue, xsNumber): strBed = strValue
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
            '�����ڴ�����ͼ��
            strSQL = "Select ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2] And NVL(״̬,0)=3"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ���Ԥ��Ժ״̬", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            
            mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
            blnFresh = False
            Do While Not mrsPatiInfo.EOF
                If InStr(1, ",3,3.1,3.2,", "," & mrsPatiInfo!���� & ",") <> 0 Then
                    blnFresh = True
                    mrsPatiInfo!���� = 4: mrsPatiInfo!���� = "Ԥ��Ժ����": mrsPatiInfo!״̬ = 3
                    mrsPatiInfo.Update
                End If
            mrsPatiInfo.MoveNext
            Loop
            If blnFresh = False Then Exit Sub
            
            mrsBedInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID & " And ����<>1"
            If Not mrsBedInfo.EOF Then
                intCardIndex = mrsBedInfo!��Ƭ����
                mrsBedInfo!����״̬ = Img���(mlngSource).ListImages("Ԥ��Ժ").Index
                mrsBedInfo!����״̬���� = "Ԥ��Ժ"
                mrsBedInfo.Update
                Call SetCardLabel(intCardIndex)
            End If
            mrsPatiInfo.Filter = "����='Ԥ��Ժ����'"
            mlngԤ��Ժ = mrsPatiInfo.RecordCount
            mrsPatiInfo.Filter = 0
            mrsBedInfo.Filter = 0
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "����Ԥ��Ժ�Ĳ���:" & strName & "   ����:" & strBed, "Ԥ��Ժ����")
            End If
            
        Case "ZLHIS_PATIENT_010" '��Ժ
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '��ȡ����ID����ҳID������
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("out_area_id", strValue, xsNumber): lngCurrUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("out_bed", strValue, xsNumber): strBed = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("out_way", strValue, xsNumber): strOutWay = strValue
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            If lngCurrUnit <> cboUnit.ItemData(cboUnit.ListIndex) Then Exit Sub
            
            strSQL = "Select ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2] And ��Ժ���� IS NOT NULL"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ��ڳ�Ժ״̬", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            
            '������
            If FreshPatiCard("ɾ����Ժ����", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
            strKey = ""
            If mintREPORTSEL = ҳ��.��Ժ Then
                If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                    If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                        strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                    End If
                End If
            End If

            rptPati(ҳ��.��Ժ).Tag = "": rptPati(ҳ��.��Ժ).Records.DeleteAll
            If rptPati(ҳ��.��Ժ).Columns.Count > c_��� Then rptPati(ҳ��.��Ժ).Columns(c_���).Visible = False
            If PatiPage.Selected.Index = ҳ��.��Ժ Then Call PatiPage_SelectedChanged(PatiPage.Selected)
            If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
            
            If strName <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���³�Ժ�Ĳ���:" & strName & "   ����:" & strBed & "   ��Ժ��ʽ:" & strOutWay, "Ԥ��Ժ����")
            End If
                
        Case "ZLHIS_PATIENT_012" 'ת�����
            If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
            strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
            If strValue <> "" And Val(strValue) = Me.hwnd Then mclsXML.CloseXMLDocument: Exit Sub
            '��ȡ����ID����ҳID������...
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsString): strName = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("in_bed", strValue, xsString): strBed = strValue
            strValue = "": Call mclsXML.GetSingleNodeValue("out_area_id", strValue, xsNumber): lngCurrUnit = Val(strValue)
            strValue = "": Call mclsXML.GetSingleNodeValue("in_area_id", strValue, xsNumber): lngUnit = Val(strValue)
            mclsXML.CloseXMLDocument
            If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
            '��鲡���Ƿ�������Ժ
            strSQL = "Select ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2] And NVL(״̬,0)=0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ�������Ժ", lngPatID, lngPageID)
            If rsTmp.EOF Then Exit Sub
            'a)ת�����������嵥��ˢ��(һ��Ҫ����ת�벡��֮ǰ,ת�ƿ��ܴ�����ס������ת��������ͬ�����)
            If lngCurrUnit = cboUnit.ItemData(cboUnit.ListIndex) Then
                '������
                If FreshPatiCard("ɾ����Ժ����", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
                    rptPati(ҳ��.ת��).Tag = "": rptPati(ҳ��.ת��).Records.DeleteAll
                    If rptPati(ҳ��.ת��).Columns.Count > c_��� Then rptPati(ҳ��.ת��).Columns(c_���).Visible = False
                    If PatiPage.Selected.Index = ҳ��.ת�� Then Call PatiPage_SelectedChanged(PatiPage.Selected)
                    
                    If strName <> "" Then
                        Call mclsMipModule.ShowMessage(strMsgItemIdentity, "������ת���Ĳ���:" & strName & "   ����:" & strBed, "��ת������")
                    End If
                End If
            End If
            'b)ת�벡�������嵥��ˢ��
            If lngUnit = cboUnit.ItemData(cboUnit.ListIndex) Then
                mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID & " And ����<>7"
                If Not mrsPatiInfo.EOF Then
                    '��鲡�˴�����Ժ����ס�б���
                    If mrsPatiInfo!���� = 1 Or mrsPatiInfo!���� = 2 Then
                        mrsPatiInfo.Delete: mrsPatiInfo.Filter = ""
                        strKey = ""
                        If mintREPORTSEL = ҳ��.����� Then
                            If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                                If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                    strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                                End If
                            End If
                        End If

                        rptPati(ҳ��.�����).Records.DeleteAll
                        If rptPati(ҳ��.�����).Columns.Count > c_��� Then rptPati(ҳ��.�����).Columns(c_���).Visible = False
                        Call UpgradeList(mrsPatiInfo, ҳ��.�����)
                        PatiPage.Item(ҳ��.�����).Caption = "�����" & GetPatiCount(ҳ��.�����) & "��"
                        If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                    Else
                        Exit Sub
                    End If
                End If
                If FreshPatiCard("������Ժ����", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                If strName <> "" Then
                    Call mclsMipModule.ShowMessage(strMsgItemIdentity, "����ת������ס�Ĳ���:" & strName & "   ����:" & strBed, "��ס����")
                End If
            End If
        Case "ZLHIS_PATIENT_006" '�����䶯
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
            Case "��Ժ"
                If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
                
                strSQL = "Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2] And����Ժ���� IS NULL"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ�������Ժ", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!��Ժ����)
                mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
                If Not mrsPatiInfo.EOF Then
                    '��鲡�˴��ڳ�Ժ�б���
                    If mrsPatiInfo!���� = 5 Or mrsPatiInfo!���� = 6 Then
                        strKey = ""
                        If mintREPORTSEL = ҳ��.��Ժ Then
                            If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                                If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                    strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                                End If
                            End If
                        End If
                        rptPati(ҳ��.��Ժ).Tag = "": rptPati(ҳ��.��Ժ).Records.DeleteAll
                        If rptPati(ҳ��.��Ժ).Columns.Count > c_��� Then rptPati(ҳ��.��Ժ).Columns(c_���).Visible = False
                        If PatiPage.Selected.Index = ҳ��.��Ժ Then Call PatiPage_SelectedChanged(PatiPage.Selected)
                        If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                    Else
                        Exit Sub
                    End If
                End If
                If FreshPatiCard("������Ժ����", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                If strName <> "" Then
                    Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���³�����Ժ�Ĳ���:" & strName & "   ����:" & strBed, "������Ժ����")
                End If
            Case "Ԥ��Ժ"
                If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
                '�����ڴ�����ͼ��
                strSQL = "Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2] And NVL(״̬,0)=0"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ�������Ժ", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!��Ժ����)
                mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
                blnFresh = False
                Do While Not mrsPatiInfo.EOF
                    If InStr(1, ",4,3.2,", "," & mrsPatiInfo!���� & ",") <> 0 Then
                        blnFresh = True
                        If strBed = "" Then
                            mrsPatiInfo!���� = 3.1: mrsPatiInfo!���� = "��ͥ����": mrsPatiInfo!״̬ = 0
                        Else
                            mrsPatiInfo!���� = 3: mrsPatiInfo!���� = "��Ժ����": mrsPatiInfo!״̬ = 0
                        End If
                        mrsPatiInfo.Update
                    End If
                mrsPatiInfo.MoveNext
                Loop
                If blnFresh = False Then Exit Sub
            
                mrsBedInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID & " And ����<>1"
                If Not mrsBedInfo.EOF Then
                    intCardIndex = mrsBedInfo!��Ƭ����
                    mrsBedInfo!����״̬ = 0
                    mrsBedInfo!����״̬���� = ""
                    mrsBedInfo.Update
                    Call SetCardLabel(intCardIndex)
                End If
                mrsPatiInfo.Filter = "����='Ԥ��Ժ����'"
                mlngԤ��Ժ = mrsPatiInfo.RecordCount
                mrsPatiInfo.Filter = 0
            
                mrsBedInfo.Filter = 0
                If strName <> "" Then
                    Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���³���Ԥ��Ժ�Ĳ���:" & strName & "   ����:" & strBed, "����Ԥ��Ժ����")
                End If
            Case "ת������ס", "ת����ס"
                '����״̬��ˢ�²������
                strSQL = "Select ��Ժ����,��ǰ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2] And NVL(״̬,0)=2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ�������Ժ", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!��Ժ����)
                'a)  ��ס������Ժ
                If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    If FreshPatiCard("ɾ����Ժ����", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                    If strName <> "" Then
                        Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���³�����ס�Ĳ���:" & strName, "������ס����")
                    End If
                End If
                
                'b)  ת��������Ժ�б�/ת���б�ˢ��
                If InStr(1, "," & Nvl(rsTmp!��ǰ����ID, 0) & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
                    If Not mrsPatiInfo.EOF Then
                        '��鲡�˴������ת���б���
                        If mrsPatiInfo!���� = 7 Then
                            strKey = ""
                            If mintREPORTSEL = ҳ��.ת�� Then
                                If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                                    If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                        strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                                    End If
                                End If
                            End If

                            rptPati(ҳ��.ת��).Tag = "": rptPati(ҳ��.ת��).Records.DeleteAll
                            If rptPati(ҳ��.ת��).Columns.Count > c_��� Then rptPati(ҳ��.ת��).Columns(c_���).Visible = False
                            If PatiPage.Selected.Index = ҳ��.ת�� Then Call PatiPage_SelectedChanged(PatiPage.Selected)
                            If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                        Else
                            Exit Sub
                        End If
                    End If
                    
                    If FreshPatiCard("������Ժ����", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
                        If strName <> "" Then
                            Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���³�����Ĳ���:" & strName & "   ����:" & strBed, "������Ժ����")
                        End If
                    End If
                End If
                
                'c)��ɴ�����б�ˢ��
                If lngUnit = 0 Then
                    strValue = ""
                    strSQL = "Select ����ID From �������Ҷ�Ӧ where ����ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lngDept)
                    Do While Not rsTmp.EOF
                        strValue = strValue & "," & rsTmp!����ID
                    rsTmp.MoveNext
                    Loop
                    strValue = Mid(strValue, 2)
                Else
                    strValue = lngUnit
                End If
                If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    If FreshPatiCard("����ת������Ʋ���", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = True Then
                        If strName <> "" Then
                            Call mclsMipModule.ShowMessage(strMsgItemIdentity, "����ת��Ĳ���:" & strName, "������ס����")
                        End If
                    End If
                End If
            Case "ת����", "ת��"
                '����״̬��ˢ�²������
                strSQL = "Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2] And NVL(״̬,0)=0"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ�������Ժ", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!��Ժ����)
                'a)ת�벡��������б����
                If lngCurrUnit <> 0 Or lngCurrDept <> 0 Then
                    If lngCurrUnit = 0 Then
                        strValue = ""
                        strSQL = "Select ����ID From �������Ҷ�Ӧ where ����ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lngCurrDept)
                        Do While Not rsTmp.EOF
                            strValue = strValue & "," & rsTmp!����ID
                        rsTmp.MoveNext
                        Loop
                        strValue = Mid(strValue, 2)
                    Else
                        strValue = lngCurrUnit
                    End If
                    If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                        mrsPatiInfo.Filter = "(����ID=" & lngPatID & " And ��ҳID=" & lngPageID & " And ����=1) OR (����ID=" & lngPatID & " And ��ҳID=" & lngPageID & " And ����=2)"
                        If Not mrsPatiInfo.EOF Then
                            mrsPatiInfo.Delete
                            strKey = ""
                            If mintREPORTSEL = ҳ��.����� Then
                                If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                                    If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                                        strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                                    End If
                                End If
                            End If
                            rptPati(ҳ��.�����).Records.DeleteAll
                            If rptPati(ҳ��.�����).Columns.Count > c_��� Then rptPati(ҳ��.�����).Columns(c_���).Visible = False
                            Call UpgradeList(mrsPatiInfo, ҳ��.�����)
                            PatiPage.Item(ҳ��.�����).Caption = "�����" & GetPatiCount(ҳ��.�����) & "��"
                            If InStr(1, strKey, "|") > 0 Then Call SelPatiCard("", strKey)
                            
                            If strName <> "" Then
                                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���³���ת��Ĳ���:" & strName, "����ת������")
                            End If
                        End If
                    End If
                End If
                'b)ת��������Ժ�б����
                If InStr(1, "," & lngUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
                    blnFresh = False
                    Do While Not mrsPatiInfo.EOF
                        If InStr(1, ",4,3.2,3.1,", "," & mrsPatiInfo!���� & ",") <> 0 Then
                            blnFresh = True
                            If mrsPatiInfo!���� = 3.1 Then
                                mrsPatiInfo!״̬ = 0
                            Else
                                mrsPatiInfo!���� = 3: mrsPatiInfo!���� = "��Ժ����": mrsPatiInfo!״̬ = 0
                            End If
                            mrsPatiInfo.Update
                        End If
                    mrsPatiInfo.MoveNext
                    Loop
                    If blnFresh = False Then Exit Sub
                    mrsBedInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID & " And ����<>1"
                    If Not mrsBedInfo.EOF Then
                        intCardIndex = mrsBedInfo!��Ƭ����
                        mrsBedInfo!����״̬ = 0
                        mrsBedInfo!����״̬���� = ""
                        mrsBedInfo.Update
                        Call SetCardLabel(intCardIndex)
                    End If
                    mrsBedInfo.Filter = 0
                    If strName <> "" Then
                        Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���³���ת���Ĳ���:" & strName & "   ����:" & strBed, "����ת������")
                    End If
                End If
            Case "��ס", "��Ժ��ס"
                '����״̬��ˢ�²������
                strSQL = "Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2] And NVL(״̬,0)=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ�������Ժ", lngPatID, lngPageID)
                If rsTmp.EOF Then Exit Sub
                strBed = Nvl(rsTmp!��Ժ����)
                'a) ��ס������Ժ�����б�ˢ��
                If InStr(1, "," & lngCurrUnit & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Then
                    If FreshPatiCard("ɾ����Ժ����", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                    If strName <> "" Then
                        Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���³�����ס�Ĳ���:" & strName, "������ס����")
                    End If
                End If
                'b)  ����ס����������б�ˢ��
                If lngUnit = 0 Then
                    strValue = ""
                    strSQL = "Select ����ID From �������Ҷ�Ӧ where ����ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lngDept)
                    Do While Not rsTmp.EOF
                        strValue = strValue & "," & rsTmp!����ID
                    rsTmp.MoveNext
                    Loop
                    strValue = Mid(strValue, 2)
                Else
                    strValue = lngUnit
                End If
                If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
                If FreshPatiCard("��������Ʋ���", lngPatID, lngPageID, cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
                If strName <> "" Then
                    Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���´�����ס�Ĳ���:" & strName, "������ס����")
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
    Case "������Ժ����"
        mrsBedInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
        If mrsBedInfo.RecordCount > 0 Then mrsBedInfo.Filter = "": Exit Function
        mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
        Do While Not mrsPatiInfo.EOF
            If mrsPatiInfo!���� = 3.1 Or (mrsPatiInfo!���� = 4 And Trim(Nvl(mrsPatiInfo!����)) = "") Then
                Exit Function
            End If
        mrsPatiInfo.MoveNext
        Loop
        '��ȡ������Ϣ
        strSQL = "Select /*+ RULE */ Decode(B.״̬,3,4,DECODE(B.��Ժ����, NULL, 3.1,DECODE(B.״̬,2,3.2,3))) as ����," & _
            " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
            " Decode(B.״̬,3,'Ԥ��Ժ����',DECODE(B.��Ժ����, NULL, '��ͥ����',DECODE(B.״̬,2,'Ԥת�Ʋ���', '��Ժ����'))) as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(B.����,A.����) ����" & mstrBriefCode & ",NVL(b.�Ա�,a.�Ա�) �Ա�,NVL(b.����,a.����) ����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " B.��Ժ���� as ����,E.���� as ����ȼ�,B.�ѱ�,B.ҽ�Ƹ��ʽ,B.��ǰ����,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID,A.��ҳId �����ҳId" & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z,��Ժ���� R" & _
            " Where B.��������=Z.����(+) And A.����ID=B.����ID And A.��ҳID=B.��ҳID And Nvl(B.״̬,0)<>1" & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And R.����ID=[3] And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And a.����ID=R.����ID And A.��ǰ����ID=R.����ID And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL" & _
            " And B.����id =[1] And B.��ҳid = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lngPatID, lngPageID, lngUnit)
        If rsTmp.EOF Then Exit Function
        If rsTmp!���� = 3.1 Or (rsTmp!���� = 4 And Trim(Nvl(rsTmp!����)) = "") Then '��ͥ����
            Call UpgradeList(rsTmp)
            Call CopyReocrd(rsTmp)
            PatiPage.Item(ҳ��.��ͥ����).Caption = "��ͥ����" & GetPatiCount(ҳ��.��ͥ����) & "��"
        Else
            strSQL = " Select Lpad(d.����, 10, ' ') As ����, Lpad(d.�����, 10, ' ') �����, d.��λ����, Nvl(b.����, a.����) ����" & mstrBriefCode & ", b.סԺ��, b.����id, b.��ҳid" & vbNewLine & _
                " From ������Ϣ a, ������ҳ b, ��Ժ���� c, ��λ״����¼ d" & vbNewLine & _
                " Where a.����id = b.����id And a.��ҳId = b.��ҳid And a.����id = c.����id And a.����id = d.����id And a.��ǰ����id = c.����id And" & vbNewLine & _
                "      a.��ǰ����id = d.����id And b.����id = [1] And b.��ҳid = [2] And c.����id = [3]" & vbNewLine & _
                " Order By Lpad(d.����, 10, ' ')"
            Set rsBed = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˴�λ��Ϣ", lngPatID, lngPageID, lngUnit)
            If rsBed.EOF Then Exit Function
            Do While Not rsBed.EOF
                mrsBedInfo.Filter = "����='" & Trim(Nvl(rsBed!����, "ZYB")) & "'"
                If mrsBedInfo.RecordCount <> 0 Then
                    strFields = "��λ����|����|סԺ��|����|����|����ID|��ҳID|�໤��|�������|�ٴ�·��|���Ա�ע1|����״̬|���Ա�ע2|���Ա�ע3|����ȼ�|��������|�����|������"
                    strValues = Trim(rsBed!��λ����) & "|" & Trim(rsBed!����) & "|" & Nvl(rsBed!סԺ��, 0) & "|" & rsBed!���� & "|" & Nvl(rsBed!����) & "|" & Nvl(rsBed!����ID, 0) & "|" & Nvl(rsBed!��ҳID, 0) & "|0|0|0||0|||0|0|" & Trim(Nvl(rsBed!�����)) & "|"
                    Call Record_Update(mrsBedInfo, strFields, strValues, "��Ƭ����|" & mrsBedInfo!��Ƭ����)
                    mlng�մ� = mlng�մ� - 1
                    mlng�ڴ� = mlng�ڴ� + 1
                End If
            rsBed.MoveNext
            Loop
            mrsBedInfo.Filter = ""
            Call UpgradeBeds(rsTmp)
            Call ShowGuage("���ݶ�ȡ����", 100)
            Call AdjustCard
            Call CopyReocrd(rsTmp)
        End If
        FreshPatiCard = True
    Case "��������Ʋ���"
        '��ʼ���ز�����Ϣ
        mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
        If Not mrsPatiInfo.EOF Then Exit Function
        strSQL = "Select 0 ����, Decode(Nvl(b.����״̬, 0), 0, 999, b.����״̬) As ����2, '��Ժ����ס����' As ����, a.����id, b.��ҳid, a.�����, b.סԺ��," & vbNewLine & _
            "       Nvl(b.����, a.����) ����" & mstrBriefCode & ", Nvl(b.�Ա�, a.�Ա�) �Ա�, Nvl(b.����, a.����) ����, d.���� As ����, c.����id, c.����ҽʦ As סԺҽʦ, b.���λ�ʿ, b.����״̬," & vbNewLine & _
            "       c.����, e.���� As ����ȼ�, b.�ѱ�,b.ҽ�Ƹ��ʽ, b.��ǰ����, Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����, b.��Ժ����, b.��Ժ��ʽ, b.��������, b.״̬, b.����, a.���￨��, -1 As ·��״̬," & vbNewLine & _
            "       Trunc(Sysdate) - Trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����, z.��ɫ, b.������, b.Ӥ������id, b.Ӥ������id,A.��ҳId �����ҳId" & vbNewLine & _
            " From ������Ϣ a, ������ҳ b, ���˱䶯��¼ c, ���ű� d, �շ���ĿĿ¼ e, �������� z" & vbNewLine & _
            " Where a.��Ժ = 1 And b.�������� = z.����(+) And a.����id = b.����id And Nvl(b.��ҳid, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid And" & vbNewLine & _
            "      (c.����id = [3] Or c.����id Is Null) And c.����id = d.Id And (d.վ�� = '" & gstrNodeNo & "' Or d.վ�� Is Null) And b.����ȼ�id = e.Id(+) And" & vbNewLine & _
            "      Nvl(c.���Ӵ�λ, 0) = 0 And c.��ֹʱ�� Is Null And c.��ʼԭ�� = 1 And b.״̬ = 1 And Exists" & vbNewLine & _
            " (Select 1 From �������Ҷ�Ӧ h Where c.����id = h.����id And h.����id = [3]) And b.����id = [1] And b.��ҳid = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ʋ�����Ϣ", lngPatID, lngPageID, lngUnit)
        If Not rsTmp.EOF Then
            Call UpgradeList(rsTmp)
            Call CopyReocrd(rsTmp)
            PatiPage.Item(ҳ��.�����).Caption = "�����" & GetPatiCount(ҳ��.�����) & "��"
            FreshPatiCard = True
        End If
    Case "����ת������Ʋ���"
        blnFresh = True
        mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
        Do While Not mrsPatiInfo.EOF
            If mrsPatiInfo!���� = 1 Or mrsPatiInfo!���� = 2 Then blnFresh = False: Exit Do
            mrsPatiInfo.MoveNext
        Loop
        If blnFresh = True Then
            '��ʼ���ز�����Ϣ
            strSQL = " Select Decode(c.��ʼԭ��, 3, 1, 2) As ����, Decode(Nvl(b.����״̬, 0), 0, 999, b.����״̬) As ����2," & vbNewLine & _
                "       Decode(c.��ʼԭ��, 3, 'ת�ƴ���ס����', 'ת��������ס����') As ����, a.����id, b.��ҳId, a.�����, b.סԺ��," & vbNewLine & _
                "       Nvl(b.����, a.����) ����" & mstrBriefCode & ", Nvl(b.�Ա�, a.�Ա�) �Ա�, Nvl(b.����, a.����) ����, d.���� As ����, c.����id," & vbNewLine & _
                "       c.����ҽʦ As סԺҽʦ, b.���λ�ʿ, b.����״̬, c.����, e.���� As ����ȼ�, b.�ѱ�,b.ҽ�Ƹ��ʽ, b.��ǰ����, Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����, b.��Ժ����, b.��Ժ��ʽ, b.��������, b.״̬, b.����," & vbNewLine & _
                "       a.���￨��, -1 As ·��״̬, Trunc(Sysdate) - Trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����, z.��ɫ, b.������, b.Ӥ������id, b.Ӥ������id,A.��ҳId �����ҳId" & vbNewLine & _
                " From ������Ϣ a, ������ҳ b, ���˱䶯��¼ c, ���ű� d, �շ���ĿĿ¼ e, �������� z" & vbNewLine & _
                " Where a.��Ժ = 1 And b.�������� = z.����(+) And a.����id = b.����id And Nvl(b.��ҳid, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid And" & vbNewLine & _
                "      (c.����id = [3] Or c.����id Is Null) And c.����id = d.Id And (d.վ�� = '"" & gstrNodeNo & ""' Or d.վ�� Is Null) And" & vbNewLine & _
                "      b.����ȼ�id = e.Id(+) And Nvl(c.���Ӵ�λ, 0) = 0 And c.��ֹʱ�� Is Null And" & vbNewLine & _
                "      (c.��ʼԭ�� = 3 And Exists (Select 1 From �������Ҷ�Ӧ h Where c.����id = h.����id And h.����id = [3]) Or c.��ʼԭ�� = 15 And c.����id = [3]) And" & vbNewLine & _
                "      (c.��ʼԭ�� In (3, 15) And c.��ʼʱ�� Is Null And b.״̬ = 2) And b. ����id = [1] And b.��ҳid = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ʋ�����Ϣ", lngPatID, lngPageID, lngUnit)
            If Not rsTmp.EOF Then
                Call UpgradeList(rsTmp)
                Call CopyReocrd(rsTmp)
                PatiPage.Item(ҳ��.�����).Caption = "�����" & GetPatiCount(ҳ��.�����) & "��"
                FreshPatiCard = True
            End If
        End If
    Case "ɾ����Ժ����"
        mrsBedInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
        If Not mrsBedInfo.EOF Then '�ڴ�����
            blnFresh = False
            mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
            Do While Not mrsPatiInfo.EOF
                If InStr(1, ",3,3.2,", "," & mrsPatiInfo!���� & ",") <> 0 Or (mrsPatiInfo!���� = 4 And Trim(Nvl(mrsPatiInfo!����)) <> "") Then
                    blnFresh = True
                    mrsPatiInfo.Delete
                End If
                mrsPatiInfo.MoveNext
            Loop
            If blnFresh = False Then mrsBedInfo.Filter = 0: Exit Function
            arrCardIndex = Array()
            Do While Not mrsBedInfo.EOF
                intCardIndex = mrsBedInfo!��Ƭ����
                ReDim Preserve arrCardIndex(UBound(arrCardIndex) + 1)
                arrCardIndex(UBound(arrCardIndex)) = intCardIndex
                'סԺ��,����,�Ա�,����,���,ҽ/��,�ѱ�,ҽ�Ƹ��ʽ,����,��Ժ����,סԺ����,���,������ɫ,����ȼ�,���￨�ţ�
                Call SetCardInfo(intCardIndex, Array("", "", "", "", "", "", "", "", "", "", "", "", &HFFFFFF, "", ""))
                mrsBedInfo.MoveNext
            Loop
            For i = 0 To UBound(arrCardIndex)
                strFields = "סԺ��|����|����|����ID|��ҳID|�໤��|�������|�ٴ�·��|���Ա�ע1|����״̬|���Ա�ע2|���Ա�ע3|����ȼ�|��������|������"
                strValues = "0|||0|0|0|0|0||0|||0|0|"
                Call Record_Update(mrsBedInfo, strFields, strValues, "��Ƭ����|" & Val(arrCardIndex(i)))
                
                picPati(Val(arrCardIndex(i))).ZOrder 1
                lblSelect(Val(arrCardIndex(i))).Visible = False
                If mblnCardCollapse Then
                    picPati(Val(arrCardIndex(i))).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
                    picPati(Val(arrCardIndex(i))).Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_�۵�, ��Ƭ����_��׼��Ƭ_�۵�)).Picture
                End If
                
                mlng�մ� = mlng�մ� + 1
                mlng�ڴ� = mlng�ڴ� - 1
            Next i
            mrsPatiInfo.Filter = ""
            mrsBedInfo.Filter = ""
            Call AdjustCard
        Else '���ڴ�����,���Ǽ�ͥ�������ˣ����Ϊ������������
            mrsBedInfo.Filter = 0
            mrsPatiInfo.Filter = "����ID=" & lngPatID & " And ��ҳID=" & lngPageID
            blnFresh = False
            Do While Not mrsPatiInfo.EOF
                If mrsPatiInfo!���� = 3.1 Or (mrsPatiInfo!���� = 4 And Trim(Nvl(mrsPatiInfo!����)) = "") Then
                    blnFresh = True
                    mrsPatiInfo.Delete
                End If
                mrsPatiInfo.MoveNext
            Loop
            If blnFresh = False Then Exit Function
            
            strKey = ""
            If mintREPORTSEL = ҳ��.��ͥ���� Then
                If rptPati(mintREPORTSEL).SelectedRows.Count <> 0 Then
                    If Not rptPati(mintREPORTSEL).SelectedRows(0).Record Is Nothing Then
                        strKey = rptPati(mintREPORTSEL).SelectedRows(0).Record.Tag
                    End If
                End If
            End If
            rptPati(ҳ��.��ͥ����).Records.DeleteAll
            If rptPati(ҳ��.��ͥ����).Columns.Count > c_��� Then rptPati(ҳ��.��ͥ����).Columns(c_���).Visible = False
            mlng�Ҵ� = 0: Call UpgradeList(mrsPatiInfo, ҳ��.��ͥ����)
            PatiPage.Item(ҳ��.��ͥ����).Caption = "��ͥ����" & GetPatiCount(ҳ��.��ͥ����) & "��"
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
    '���ݴ��Ż�ȡ����ID(��Ϊ�˴���ȡ�Ĵ���Ϊ������)
    mrsBedInfo.Filter = ""
    Do While Not mrsBedInfo.EOF
        If InStr(1, "," & strBeds & ",", "," & Nvl(mrsBedInfo!����) & ",") <> 0 Then
            strKeys = strKeys & "," & mrsBedInfo!����ID
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
        
    '��ǰ����Ϊ��ǰҪ��λ�Ĳ���
    blnSeek = False
    
    mrsPatiInfo.Filter = "����ID=" & PatiID & " and ��ҳID=" & PageID
    blnSeek = mrsPatiInfo.RecordCount > 0
    If blnSeek = True Then
        lngDept = Val(mrsPatiInfo.Fields("����ID").Value)
        mrsBedInfo.Filter = "����ID=" & PatiID & " And ����=0"
        If mrsBedInfo.RecordCount > 0 Then strTab = Nvl(mrsBedInfo.Fields("����").Value)
        mrsBedInfo.Filter = ""
    End If
    mrsPatiInfo.Filter = 0
    If Not blnSeek Then
        MsgBox "��ǰ���������嵥��û���ҵ��ò��ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call SelPatiCard(strTab, PatiID & "|" & PageID)
    If Not LocatePatiRecord Then
        MsgBox "��λ����ʧ��,���ڵ�ǰ���������嵥�к�ʵ�����Ƿ���ڡ�", vbInformation, gstrSysName
        Exit Sub
    End If

    '��λ����Ӧ������ҳ��
    strTab = Decode(ObjectType, 1, "ҽ��", 2, "����", 3, "������", 4, "����", 5, "", 6, "ҽ��", 7, "����", 8, "����")
    
    If ObjectType = 1 Or ObjectType = 4 Or ObjectType = 6 Then
        '�ж�Ȩ��
        blnSeek = False
        If ObjectType = 4 Then
            If GetInsidePrivs(p�����¼����, True) <> "" Then
                blnSeek = True
            Else
                strTab = "����"
            End If
        Else
            If GetInsidePrivs(pסԺҽ���´�, True) <> "" Or GetInsidePrivs(pסԺҽ������, True) <> "" Then
                blnSeek = True
            Else
                strTab = "ҽ��"
            End If
        End If
        If blnSeek = False Then
            MsgBox "���ܴ�" & strTab & "ҳ��,��������û����Ӧ��Ȩ�ޡ�", vbInformation, gstrSysName
        Else
            Call InNurseRoutine(strTab)
            Call OrientTabPage_Rountine(strTab, ObjectID)
        End If
        Exit Sub
    End If
    
    '�򿪶�Ӧ�Ķ���
    Select Case ObjectType
    Case 1 'סԺҽ��
    Case 2, 3, 7, 8 'סԺ����,������,����֤��,֪���ļ�
        If ObjectID = "0" Or ObjectID = "" Then Exit Sub
        If IsNumeric(ObjectID) Then
            Call gobjRichEPR.EditDocument(P�°滤ʿվ, Me, cboUnit.ItemData(cboUnit.ListIndex), ObjectID)
        Else '�°没��
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
            
            strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, p���Ӳ�������) & ";"
            If Nvl(rsEmr!completor) = "" Then
                If InStr(strPrivs, ";�ĵ���д;") > 0 Then '����дȨ��
                    Call gobjEmr.OpenFormForModifyDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, rsEmr!basiclogid, rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, CInt(rsEmr!Occasion), CInt(rsEmr!besealed), CInt(rsEmr!docsecret), Nvl(rsEmr!subdocid), strPrivs)
                Else '��Ȩ��ֻ�ܲ鿴
                    Set objEmr = DynamicCreate("zlRichEMR.clsDockContent", "��ʾ����", True)
                    If Not objEmr Is Nothing Then
                        Call objEmr.Init(gobjEmr, gcnOracle, glngSys, 0)
                        Call objEmr.zlShowDoc(strDocID, strSubdocID)
                        Call objEmr.zlViewDoc(Me, "���Ĳ���", strSubdocID)
                    End If
                End If
            Else
                If InStr(strPrivs, ";�ĵ���;") > 0 Then '����дȨ��
                    Call gobjEmr.OpenFormForAuditDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, rsEmr!basiclogid, rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, CInt(rsEmr!Occasion), CInt(rsEmr!besealed), CInt(rsEmr!docsecret), Nvl(rsEmr!subdocid), strPrivs)
                Else '��Ȩ��ֻ�ܲ鿴
                    Set objEmr = DynamicCreate("zlRichEMR.clsDockContent", "��ʾ����", True)
                    If Not objEmr Is Nothing Then
                        Call objEmr.Init(gobjEmr, gcnOracle, glngSys, 0)
                        Call objEmr.zlShowDoc(strDocID, strSubdocID)
                        Call objEmr.zlViewDoc(Me, "���Ĳ���", strSubdocID)
                    End If
                End If
            End If
        End If
    Case 4 '�����¼
    Case 5 '��ҳ��¼
        Call PrintInMedRec(mclsInOutMedRec, 1, PatiID, PageID, mobjReport, lngDept, Me)
    Case 6 'ҽ������
        
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
    '�޸Ĵ�SQL������,�����������ģ��Ҳ��Ҫ����
    Dim i As Long
    
    Call picPatiIn_Resize
    Me.MousePointer = 11
    '61824:������,2013-05-23,��ʾ�����ֱ�־
    mintREPORTSEL = Item.Index
    strField = "����|����2|����|����ID|��ҳID|סԺ��|����|����|�Ա�|����|����|����ID|סԺҽʦ|���λ�ʿ|����״̬|����|����ȼ�|�ѱ�|ҽ�Ƹ��ʽ|��ǰ����|��Ժ����|��Ժ����|סԺ����|��Ժ��ʽ|��������|״̬|����|���￨��|·��״̬|��ɫ|������|Ӥ������ID|Ӥ������ID|�����ҳId"
    If rptPati(Item.Index).Tag = "" Then
        If Item.Index = ҳ��.��Ժ Or Item.Index = ҳ��.ת�� Then
            If Item.Index = ҳ��.��Ժ Then
                '68259:������,2012-02-11,��Ժ���˲������δ�����ѽ��幦��
                If chkSettle(0).Value = 1 And chkSettle(1).Value = 1 Then
                    intSettle = 0              '����ʾ
                ElseIf chkSettle(0).Value = 0 And chkSettle(1).Value = 1 Then
                    intSettle = 1               'ֻ��ʾδ�����
                ElseIf chkSettle(0).Value = 1 And chkSettle(1).Value = 0 Then
                    intSettle = 2              'ֻ��ʾ�ѽ����
                End If
    
                '��Ժ����:��Ժ���˿������ж��סԺ
                strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
                    "Select /*+ RULE */ Decode(B.��Ժ��ʽ,'����',6,5) as ����," & _
                    " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
                    " Decode(B.��Ժ��ʽ,'����','��������','��Ժ����') as ����," & _
                    " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(B.����,A.����) ����" & mstrBriefCode & ",NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
                    " B.��Ժ���� AS ����,E.���� as ����ȼ�,B.�ѱ�,B.ҽ�Ƹ��ʽ,B.��ǰ����,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
                    " B.״̬,B.����,A.���￨��,Nvl(b.·��״̬,-1) ·��״̬,trunc(b.��Ժ����)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID,A.��ҳId �����ҳId" & _
                    " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z" & _
                    " Where A.����ID=B.����ID And B.��������=Z.����(+) And Nvl(B.��ҳID,0)<>0 And B.״̬=0" & _
                    " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And B.��ǰ����ID+0=[1] And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
                    " And B.��Ժ���� Between [2] And [3] And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL" & _
                    IIf(intSettle = 0, "", " And " & IIf(intSettle = 1, "", "Not") & " Exists(Select 1 From ����δ����� Where B.����id = ����id  And B.��ҳid = ��ҳid and ��Դ;��=2 Having Nvl(Sum(���), 0) <> 0)")
            Else
                'ת������:��Ժ,ҽ���ʹ�����ʾ����ת��ǰ��
                strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
                    "Select /*+ RULE */ Distinct 7 as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,'ת������' as ����," & _
                    " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(B.����,A.����) ����" & mstrBriefCode & ",NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,D.���� as ����,C.����ID,C.����ҽʦ as סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
                    " C.����,E.���� as ����ȼ�,B.�ѱ�,B.ҽ�Ƹ��ʽ,B.��ǰ����,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
                    " B.״̬,B.����,A.���￨��,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID,A.��ҳId �����ҳId" & _
                    " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D,�շ���ĿĿ¼ E,�������� Z" & _
                    " Where A.����ID=B.����ID And B.��������=Z.����(+) And Nvl(B.��ҳID,0)<>0 And B.����ȼ�ID=E.ID(+)" & _
                    " And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
                    " And B.��ǰ����ID<>[1] And C.����ID+0=[1] And C.����ID=D.ID" & _
                    " And Nvl(C.���Ӵ�λ,0)=0 And C.��ֹԭ�� In(3,15) And C.��ֹʱ�� Between Sysdate-[4] And Sysdate" & _
                    " And Nvl(B.״̬,0)<>2 And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
            End If
            strSQL = strSQL & " Order by ����,����2,����,��ҳID Desc"

            On Error GoTo ErrHand
            Set rsPati = New ADODB.Recordset
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), _
                CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")), mintChange)
            
            Call UpgradeList(rsPati)
            
            '��ɾ��ԭ�м�¼��
            If Item.Index = ҳ��.��Ժ Then
                mrsPatiInfo.Filter = "����=5 or ����=6"
            Else
                mrsPatiInfo.Filter = "����=7"
            End If
            For i = 1 To mrsPatiInfo.RecordCount
                mrsPatiInfo.Delete
                mrsPatiInfo.MoveNext
            Next
            
            '׷�Ӽ�¼��
            mrsPatiInfo.Filter = 0
            If rsPati.RecordCount <> 0 Then rsPati.MoveFirst
            Do While Not rsPati.EOF
                strValue = rsPati!���� & "|" & Nvl(rsPati!����2, 0) & "|" & Nvl(rsPati!����) & "|" & Nvl(rsPati!����ID, 0) & "|" & Nvl(rsPati!��ҳID, 0) & "|" & Nvl(rsPati!סԺ��, 0) & "|" & Nvl(rsPati!����) & "|" & Nvl(rsPati!����) & "|" & Nvl(rsPati!�Ա�) & "|" & _
                          Nvl(rsPati!����) & "|" & Nvl(rsPati!����) & "|" & Nvl(rsPati!����ID, 0) & "|" & Nvl(rsPati!סԺҽʦ) & "|" & Nvl(rsPati!���λ�ʿ) & "|" & Nvl(rsPati!����״̬, 0) & "|" & Nvl(rsPati!����) & "|" & _
                          Nvl(rsPati!����ȼ�, "����") & "|" & Nvl(rsPati!�ѱ�) & "|" & Nvl(rsPati!ҽ�Ƹ��ʽ) & "|" & Nvl(rsPati!��ǰ����, "һ��") & "|" & Nvl(rsPati!��Ժ����) & "|" & Nvl(rsPati!��Ժ����) & "|" & Nvl(rsPati!סԺ����) & "|" & Nvl(rsPati!��Ժ��ʽ) & "|" & _
                          Nvl(rsPati!��������, "��ͨ����") & "|" & Nvl(rsPati!״̬, 0) & "|" & Nvl(rsPati!����, 0) & "|" & Nvl(rsPati!���￨��) & "|" & Nvl(rsPati!·��״̬, 0) & "|" & Nvl(rsPati!��ɫ, 0) & "|" & Nvl(rsPati!������) & "|" & Nvl(rsPati!Ӥ������ID, 0) & "|" & Nvl(rsPati!Ӥ������ID, 0) & "|" & Nvl(rsPati!�����ҳID, 0)
                Call Record_Add(mrsPatiInfo, strField, strValue)
                rsPati.MoveNext
            Loop
            
            rptPati(Item.Index).Tag = "OK"
            If GetPatiCount(Item.Index) <> 0 Then
                PatiPage.Item(Item.Index).Caption = IIf(Item.Index = ҳ��.��Ժ, "�����Ժ", "���ת��") & GetPatiCount(Item.Index) & "��"
            End If
        End If
    End If
    pic��Ժ����.Visible = True
    pic��Ժ����.ZOrder 0

    If Item.Index = ҳ��.��Ժ Then
        '����ǰҳ��Ĺ���������ʾ��״̬����
        Me.stbThis.Panels(2).Text = Format(mdtOutBegin, "yyyy-MM-dd") & "��" & Format(mdtOutEnd, "yyyy-MM-dd") & "֮��" & IIf(intSettle = 0, "", IIf(intSettle = 1, "δ����", "�ѽ���")) & "�ĳ�Ժ����"
    ElseIf Item.Index = ҳ��.ת�� Then
        '����ǰҳ��Ĺ���������ʾ��״̬����
        Me.stbThis.Panels(2).Text = "���" & mintChange & "���ڵ�ת������"
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

Private Sub pic��Ժ����_GotFocus()
    If txtסԺ��.Enabled And txtסԺ��.Visible Then txtסԺ��.SetFocus
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

    'ȡ���˻�����Ϣ
    blnSelect = LocatePatiRecord
    If blnSelect Then
        lngType = Val(mrsPatiInfo.Fields("����").Value)
        blnWaitIn = lngType = ptת�ƴ���ס Or lngType = pt��Ժ����ס
        blnOut = lngType = pt��Ժ
        blnPreOut = lngType = ptԤ��
        '85200:�������ת��ҳ��Ĳ��˲����������ز������磺��������
        blnOutTo = lngType = pt���ת��
    Else
        Exit Sub
    End If
    '���ð�ť״̬
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p�������)
    If InStr(strPrivs, "���в���") = 0 Then
        If InStr("," & mstrUnits & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
    End If

    '��װ�Ҽ��˵�
    Set cbrMenuBar = mobjPopup
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.IconId = cbrControl.IconId
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        cbrPopupItem.Visible = cbrControl.Visible

        Call SetControlVisible(cbrPopupItem)

        '���ð�ť��״̬
        Select Case cbrControl.ID
        Case conMenu_Manage_Change_Undo
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOutTo
            If cbrPopupItem.Enabled = True Then
                cbrPopupItem.Enabled = Val(Nvl(mrsPatiInfo.Fields("��ҳID").Value, 0)) = Val(Nvl(mrsPatiInfo.Fields("�����ҳId").Value, 0))
            End If
            Call cbsMain_InitCommandsPopup(cbrMenuBar.CommandBar)
        Case conMenu_Manage_Change_In
            cbrPopupItem.Enabled = blnWaitIn
        Case conMenu_Manage_Change_InPati
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = mPatiInfo.���� = 2
            End If
        Case conMenu_Manage_Change_Turn, conMenu_Manage_Change_Bed, conMenu_Manage_Change_House, _
             conMenu_Manage_Change_PatiInfo, conMenu_Manage_Change_ReCalcFee, conMenu_Manage_BedExchange
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = mrsPatiInfo.Fields("״̬").Value <> 2
            End If
            If cbrPopupItem.ID = conMenu_Manage_Change_ReCalcFee And cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = Nvl(mrsPatiInfo.Fields("����").Value, 0) = 0
            End If
        Case conMenu_Manage_Change_InsureSel
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = Nvl(mrsPatiInfo.Fields("����").Value, 0) <> 0
            End If
        Case conMenu_Manage_Change_BedGrid
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = Trim(Nvl(mrsPatiInfo.Fields("����").Value)) <> "" And mrsPatiInfo.Fields("״̬").Value <> 2
            End If
        Case conMenu_Manage_Change_Out
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = (InStr(1, "," & pt��Ժ & ",3.1,", mrsPatiInfo.Fields("����").Value) <> 0 Or blnPreOut) And mrsPatiInfo.Fields("״̬").Value <> 2
            End If
        Case conMenu_Manage_Change_Baby
            cbrPopupItem.Enabled = blnSelect And Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
            If cbrPopupItem.Enabled Then
                cbrPopupItem.Enabled = mPatiInfo.����
            End If
        Case conMenu_Manage_Change_PaitNote
            cbrPopupItem.Enabled = Not blnOutTo
        Case conMenu_Manage_Monitor '�໤��
            cbrPopupItem.Visible = mblnMonitor And (InStr(GetInsidePrivs(pסԺ��ʿվ), "����໤") > 0)
        End Select
    Next
    If Not mrsPlugInBar Is Nothing Then
        mrsPlugInBar.Filter = "IsInTool=1 and BarType=3"
        For i = 1 To mrsPlugInBar.RecordCount
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mrsPlugInBar!����ID, mrsPlugInBar!������)
                cbrPopupItem.IconId = mrsPlugInBar!ͼ��ID
                cbrPopupItem.Parameter = mrsPlugInBar!������
                If Val(mrsPlugInBar!IsGroup) = 1 Then cbrPopupItem.BeginGroup = True
            mrsPlugInBar.MoveNext
        Next
        mrsPlugInBar.Filter = "IsInTool=0 and BarType=3"
        If mrsPlugInBar.RecordCount > 0 Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugInPop, "��չ����"): cbrPopupItem.BeginGroup = True
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
    '53740:������,2012-09-19
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
    '����ʱ��ǿ���Ȱ����״̬����
    '������������Ч����������һ������
    On Error GoTo ErrHand
    
    If rptPati(Index).SortOrder.Count > 0 Then
        If rptPati(Index).SortOrder(0).Index <> c_��� Then
            Set objCol = rptPati(Index).SortOrder(0)
            rptPati(Index).SortOrder.DeleteAll
            rptPati(Index).SortOrder.Add rptPati(Index).Columns(c_���)
            rptPati(Index).SortOrder.Add objCol
        Else
            '���ж�Ϊ���Է���һ��ֻ���ڵ������е�ʱ��COUNT=1��������в��ɼ��������������COUNT=2
            If rptPati(Index).SortOrder.Count > 1 Then
                Set objCol = rptPati(Index).SortOrder(1)
            Else
                Exit Sub
            End If
        End If
        blnAsc = objCol.SortAscending
        lngIndex = objCol.Index
        
        If lngIndex = c_ͼ�� Then Exit Sub
        '86154:������,2015-07-02,ReportControl��֧������������
        For i = 0 To rptPati(Index).Records.Count - 1
            Set objParent = rptPati(Index).Records(i)
            If objParent.Childs.Count > 0 Then
                '��ʼ����¼��
                strFields = "����," & adVarChar & ",50|����ID," & adDouble & ",20|��ҳID," & adDouble & ",20|����," & adVarChar & ",20|����״̬," & adDouble & ",10|" & _
                    "��������," & adLongVarChar & ",500|������," & adVarChar & ",10|·��״̬," & adDouble & ",10|����," & adLongVarChar & ",100|" & _
                    "סԺ��," & adVarChar & ",20|����," & adVarChar & ",20|�Ա�," & adVarChar & ",20|����," & adVarChar & ",50|�ѱ�," & adVarChar & ",20|" & _
                    "���ʽ," & adVarChar & ",30|סԺҽʦ," & adLongVarChar & ",100|��Ժ����," & adVarChar & ",20|��Ժ����," & adVarChar & ",20|" & _
                    "��������," & adVarChar & ",50|���￨��," & adVarChar & ",50"
                Call Record_Init(rsTemp, strFields)
                strFields = "����|����ID|��ҳID|����|����״̬|��������|������|·��״̬|����|סԺ��|����|�Ա�|����|�ѱ�|���ʽ|סԺҽʦ|��Ժ����|��Ժ����|��������|���￨��"
                For j = 0 To objParent.Childs.Count - 1
                    strKey = objParent.Childs(j).Item(C_����ID).Value & "-" & objParent.Childs(j).Item(C_��ҳID).Value
                    strValues = strKey & "'" & objParent.Childs(j).Item(C_����ID).Value & "'" & objParent.Childs(j).Item(C_��ҳID).Value & "'" & objParent.Childs(j).Item(c_����).Value & "'" & _
                        objParent.Childs(j).Item(c_���).Value & "'" & objParent.Childs(j).PreviewText & "'" & objParent.Childs(j).Item(c_ͼ��).Value & "'" & _
                        objParent.Childs(j).Item(c_·��״̬).Value & "'" & objParent.Childs(j).Item(c_����).Value & "'" & objParent.Childs(j).Item(c_סԺ��).Value & "'" & _
                        objParent.Childs(j).Item(c_����).Value & "'" & objParent.Childs(j).Item(c_�Ա�).Value & "'" & objParent.Childs(j).Item(c_����).Value & "'" & _
                        objParent.Childs(j).Item(c_�ѱ�).Value & "'" & objParent.Childs(j).Item(c_���ʽ).Value & "'" & objParent.Childs(j).Item(c_ҽ��).Value & "'" & _
                        objParent.Childs(j).Item(c_��Ժ����).Value & "'" & objParent.Childs(j).Item(c_��Ժ����).Value & "'" & objParent.Childs(j).Item(c_��������).Value & "'" & _
                        objParent.Childs(j).Item(c_���￨��).Value
                    Call Record_Update(rsTemp, strFields, strValues, "����|" & strKey, , "'")
                Next j
                objParent.Childs.DeleteAll
                '����ѡ���������
                strKey = ""
                Select Case lngIndex
                    Case c_����
                        strKey = "����"
                    Case c_���
                        strKey = "����״̬"
                    Case c_ͼ��
                        strKey = ""
                    Case c_·��״̬
                        strKey = "·��״̬"
                    Case C_����ID
                        strKey = "����ID"
                    Case C_��ҳID
                        strKey = "��ҳID"
                    Case c_����
                        strKey = "����"
                    Case c_סԺ��
                        strKey = "סԺ��"
                    Case c_����
                        strKey = "����"
                    Case c_�Ա�
                        strKey = "�Ա�"
                    Case c_����
                        strKey = "����"
                    Case c_�ѱ�
                        strKey = "�ѱ�"
                    Case c_���ʽ
                        strKey = "���ʽ"
                    Case c_ҽ��
                        strKey = "סԺҽʦ"
                    Case c_��Ժ����
                        strKey = "��Ժ����"
                    Case c_��Ժ����
                        strKey = "��Ժ����"
                    Case c_��������
                        strKey = "��������"
                    Case c_���￨��
                        strKey = "���￨��"
                End Select
                
                rsTemp.Filter = ""
                If strKey <> "" Then rsTemp.Sort = strKey & IIf(blnAsc, "", " DESC")
                '����������������
                With rsTemp
                    Do While Not .EOF
                        Set objRecord = objParent.Childs.Add
                        objRecord.Tag = CStr(!����ID & "|" & !��ҳID)
                        Set objItem = objRecord.AddItem(CStr("" & !����))
                        objItem.Caption = CStr("" & !����)
                        
                        Set objItem = objRecord.AddItem(Val(Decode(Nvl(!����״̬, 0), 0, 999, Nvl(!����״̬, 0))))
                        objItem.Caption = " "
                        If Val(Nvl(!����״̬, 0)) = 2 Then
                            objRecord.PreviewText = "" & !��������
                        End If
                        
                        Set objItem = objRecord.AddItem(Nvl(!������))
                        objItem.Caption = " "
                        '81308:������,2015-01-19,��ʾ��������־
                        '61824:������,2013-05-23,��ʾ�����ֱ�־
                        If Nvl(!����״̬, 0) <> 0 Then
                            objItem.Icon = Get����ͼ�����(!����״̬, False) - 1
                        ElseIf Nvl(!������) <> "" Then
                            objItem.Icon = imgRPT.ListImages("������").Index - 1
                        Else
                            objItem.Icon = Val(IIf(!�Ա� = "Ů", imgRPT.ListImages("Ů��").Index, imgRPT.ListImages("����").Index)) - 1
                        End If
                        
                        '·��״̬
                        Set objItem = objRecord.AddItem(Val("" & !·��״̬))
                        objItem.Caption = " "
                        objItem.Icon = Get�ٴ�·�����(Val("" & !·��״̬) + 2, False) - 1
                        
                        objRecord.AddItem Val(!����ID)
                        objRecord.AddItem Val(!��ҳID)
                        objRecord.AddItem CStr(Nvl(!����))
                        Set objItem = objRecord.AddItem(CStr(Nvl(!סԺ��)))
                        objItem.Caption = Nvl(!סԺ��, " ")
                        Set objItem = objRecord.AddItem(Nvl(!����))
                        objItem.Caption = CStr(Nvl(!����, " "))
                        Set objItem = objRecord.AddItem(CStr(Nvl(!�Ա�, "��")))
                        objItem.Caption = CStr(Nvl(!�Ա�, "��"))
                        Set objItem = objRecord.AddItem(Nvl(!����, "0"))
                        objItem.Caption = Nvl(!����, "0")
                        Set objItem = objRecord.AddItem(Nvl(!�ѱ�, ""))
                        objItem.Caption = CStr(Nvl(!�ѱ�, ""))
                        Set objItem = objRecord.AddItem(Nvl(!���ʽ, ""))
                        objItem.Caption = CStr(Nvl(!���ʽ, ""))
                        Set objItem = objRecord.AddItem(Nvl(!סԺҽʦ, ""))
                        objItem.Caption = CStr(Nvl(!סԺҽʦ, ""))
                        Set objItem = objRecord.AddItem(CStr(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss")))
                        objItem.Caption = CStr(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                        Set objItem = objRecord.AddItem(CStr(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss")))
                        objItem.Caption = CStr(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                        Set objItem = objRecord.AddItem(Nvl(!��������, "��ͨ����"))
                        objItem.Caption = CStr(Nvl(!��������, "��ͨ����"))
                        Set objItem = objRecord.AddItem(CStr(Nvl(!���￨��)))
                        objItem.Caption = Nvl(!���￨��, "")
                        '��ȡ�������͵���ɫ
                        lngColor = 0
                        mrsPatiColor.Filter = "����='" & Nvl(!��������, "��ͨ����") & "'"
                        If mrsPatiColor.RecordCount <> 0 Then
                            lngColor = Nvl(mrsPatiColor!��ɫ, 0)
                        End If
                        If lngColor <> 0 Then
                            objRecord.Item(c_����).ForeColor = lngColor
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
    If Panel.Key = "������ɫ" Then
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
        '30s���벻���κ���Ӧ���Զ���ԭ
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
    
    'ˢ�²������ķ�����ÿ5����
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
    '�ṩ���ⲿ����Ľӿ�,ѡ��ָ�����˵Ĵ�λ��
    
    If strBed <> "" Then
        mrsBedInfo.Filter = "����='" & strBed & "'"
        If mrsBedInfo.RecordCount <> 0 Then intIndex = mrsBedInfo!��Ƭ����
        mrsBedInfo.Filter = 0
    End If
    
    If intIndex > 0 Then
        'ȡ���ϴ�ѡ��
        Call picPati_MouseDown(intIndex, 1, 0, 0, 0)
        'ѡ��ָ����Ƭ
        mblnShow = False            '�����,��Ȼ�ֻᴥ��ShowSelect
        Call ShowSelect
        '���⿨Ƭ��ʾ��������
        Call picPati_MouseUp(intIndex, 1, 0, 0, 0)
        '��ѡ�п�Ƭ��ʾ�ڿ���������
        Call ShowEfficiency
    Else
        '���ڴ�����
        intPage = -1
        mrsPatiInfo.Filter = "����ID=" & Split(strKey, "|")(0) & " And ��ҳID=" & Split(strKey, "|")(1)
        If mrsPatiInfo.RecordCount <> 0 Then
            If mrsPatiInfo!���� = 0 Or mrsPatiInfo!���� = 1 Or mrsPatiInfo!���� = 2 Then
                intPage = 0
            ElseIf mrsPatiInfo!���� = 7 Then
                intPage = 1
            ElseIf mrsPatiInfo!���� = 6 Or mrsPatiInfo!���� = 5 Then
                intPage = 2
            ElseIf mrsPatiInfo!���� = 3.1 Then '��ͥ����
                intPage = 3
            End If
        End If
        mrsPatiInfo.Filter = 0
        
        If intPage <> -1 Then
            PatiPage(intPage).Selected = True
            mintREPORTSEL = intPage
            
            '���Ҷ�λ����
            Dim objRptRow As ReportRow
            For Each objRptRow In rptPati(intPage).Rows
                If Not objRptRow.Record Is Nothing Then
                    If objRptRow.Record.Childs.Count = 0 Then
                        If Val(objRptRow.Record.Item(C_����ID).Value) = Val(Split(strKey, "|")(0)) And _
                            Val(objRptRow.Record.Item(C_��ҳID).Value) = Val(Split(strKey, "|")(1)) Then
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
'���ҽ�����ѣ���ѡ�еĲ�����ʾ����Ч������
    Dim lngTop As Long, lngY As Long
    Dim lngMove As Long
    
    lngMove = CLng((mdblScaleHeight - (PicDraw.Height - IIf(picList.Visible, picList.Height, 0))) / 100) '��ȡ����
    If lngMove <= 0 Then lngMove = 1
    lngY = clngX + picPati(mlngSelect).Height
    lngTop = picPati(mlngSelect).Top - (-1 * HScr.Value * lngMove)  '��ȡԭʼ��Ƭ��Top
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
    '54370:������,2013-05-02,��Ӳ���"ҽ��������Զ���λ��ҽ��ҳ��"
    '�ṩ��ҽ�����ѵ�ר�ýӿ�,�ǲ�������������
BeginFunc:
    Select Case intFunc
    Case E����
        lngKey = conMenu_Edit_Send
    Case EУ��
        lngKey = conMenu_Edit_Audit
    Case Eֹͣ
        lngKey = conMenu_Edit_ReStop
    '55430:������,2013-02-27,˫������ҽ����λ�����������ҽ��ҳ��
    Case E�鿴
        lngKey = conMenu_����������
    Case 11 '��Һ���δͨ��
        lngKey = conMenu_����������
    Case 12 '������������
        lngKey = conMenu_Edit_ReBillingApply
    End Select
    Select Case intFunc
    Case E�鿴
        Set objParent = cbsMain.Item(1).Controls.Item(3)        '������������
    Case E����, EУ��, Eֹͣ
        Set objParent = cbsMain.Item(1).Controls.Item(4)        'ҽ��ҵ��
    Case 11 '��Һ���δͨ��
        Set objParent = cbsMain.Item(1).Controls.Item(3)        '������������
    Case 12 '������������
        Set objParent = cbsMain.Item(1).Controls.Item(5)        '����ҵ��
    End Select
    For Each objControl In objParent.CommandBar.Controls
        If objControl.ID = lngKey Then blnSel = True: Exit For
    Next
    If blnSel Then
        objControl.Execute
        If intFunc = E�鿴 Or intFunc = 11 Then
            Call OrientTabPage_Rountine
        ElseIf intFunc = E���� Or intFunc = EУ�� Or intFunc = Eֹͣ Then
            If mblnCollateAutoFind = True Then intFunc = E�鿴: GoTo BeginFunc
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
'���ܣ���ȡ������鷴��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngCount As Long
    Dim curDate As Date
    
    If cboUnit.ListIndex = -1 Then
        fra���.Visible = False: LoadResponse = True: Exit Function
    End If

    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    Screen.MousePointer = 11

    '��ȡ��ǰ��������Ժ����Ժ���ˣ���"����������¼"Ϊ׼ȫ��ɨ��
    strSQL = "Select Count(*) as ���� From ������ҳ B,����������¼ A" & _
        " Where A.����ID=B.����ID and A.��ҳID=B.��ҳID And A.��¼״̬=1" & _
        " And A.�������� IN(3,4) And B.��ǰ����ID + 0 =[1]" & _
        " And a.����ʱ�� Between [2] And [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LoadResponse", cboUnit.ItemData(cboUnit.ListIndex), CDate(Format(curDate - mlngMedRedDay, "yyyy-MM-dd")), CDate(Format(curDate, "yyyy-MM-dd HH:mm:ss")))
    If Not rsTmp.EOF Then lngCount = Nvl(rsTmp!����, 0)
    
    lbl���.Caption = mlngMedRedDay & "���ڹ��� " & lngCount & " ��δ����Ĳ�����鷴��..."
    fra���.Visible = lngCount > 0
    lbl���.Tag = lngCount

    Screen.MousePointer = 0
    LoadResponse = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub init���ڴ��嵥()
    Dim objCol As ReportColumn
    '��ʼ�����ڴ������嵥
    PatiPage.Item(ҳ��.�����).Caption = "�����"
    PatiPage.Item(ҳ��.ת��).Caption = "���ת��"
    PatiPage.Item(ҳ��.��Ժ).Caption = "�����Ժ"
    PatiPage.Item(ҳ��.��ͥ����).Caption = "��ͥ����"

    rptPati(ҳ��.�����).Tag = ""       '�˱�������ж������Ƿ���Ҫˢ��
    rptPati(ҳ��.ת��).Tag = ""
    rptPati(ҳ��.��Ժ).Tag = ""
    rptPati(ҳ��.��ͥ����).Tag = ""

    rptPati(ҳ��.�����).Records.DeleteAll
    rptPati(ҳ��.ת��).Records.DeleteAll
    rptPati(ҳ��.��Ժ).Records.DeleteAll
    rptPati(ҳ��.��ͥ����).Records.DeleteAll
    
    Call InitReportControl(ҳ��.�����)
    Call InitReportControl(ҳ��.ת��)
    Call InitReportControl(ҳ��.��Ժ)
    Call InitReportControl(ҳ��.��ͥ����)
End Sub

Private Sub InitReportControl(ByVal intIndex As Integer)
    Dim arrWith() As String
    Dim objCol As ReportColumn
    
    arrWith = Split(mstrColWidth, ",")
    With rptPati(intIndex)
        .Columns.DeleteAll
        Set objCol = .Columns.Add(c_����, "����", Val(arrWith(c_����)), True): objCol.Groupable = True ': objCol.Visible = IIf(intIndex = ҳ��.�����, True, IIf(intIndex = ҳ��.��Ժ, True, False))
        Set objCol = .Columns.Add(c_���, "", Val(arrWith(c_���)), False): objCol.TreeColumn = True: objCol.Visible = False: objCol.Sortable = False: objCol.AllowDrag = False
       Set objCol = .Columns.Add(c_ͼ��, "", Val(arrWith(c_ͼ��)), False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_·��״̬, "·��״̬", Val(arrWith(c_·��״̬)), True): objCol.Visible = mblnHavePath
        Set objCol = .Columns.Add(C_����ID, "����ID", Val(arrWith(C_����ID)), False): objCol.Visible = False
        Set objCol = .Columns.Add(C_��ҳID, "��ҳID", Val(arrWith(C_��ҳID)), False): objCol.Visible = False
        Set objCol = .Columns.Add(c_����, "����", Val(arrWith(c_����)), True)
        Set objCol = .Columns.Add(c_סԺ��, "סԺ��", Val(arrWith(c_סԺ��)), True)
        Set objCol = .Columns.Add(c_����, "����", Val(arrWith(c_����)), True)
        Set objCol = .Columns.Add(c_�Ա�, "�Ա�", Val(arrWith(c_�Ա�)), True)
        Set objCol = .Columns.Add(c_����, "����", Val(arrWith(c_����)), True)
        Set objCol = .Columns.Add(c_�ѱ�, "�ѱ�", Val(arrWith(c_�ѱ�)), True)
        Set objCol = .Columns.Add(c_���ʽ, "ҽ�Ƹ��ʽ", Val(arrWith(c_���ʽ)), True)
        Set objCol = .Columns.Add(c_ҽ��, "ҽ��", Val(arrWith(c_ҽ��)), True)
        Set objCol = .Columns.Add(c_��Ժ����, "��Ժ����", Val(arrWith(c_��Ժ����)), True): objCol.SortAscending = False
        Set objCol = .Columns.Add(c_��Ժ����, "��Ժ����", Val(arrWith(c_��Ժ����)), True): objCol.Visible = IIf(intIndex = ҳ��.��Ժ, True, False)
        Set objCol = .Columns.Add(c_��������, "��������", Val(arrWith(c_��������)), True)
        Set objCol = .Columns.Add(c_���￨��, "���￨��", Val(arrWith(c_���￨��)), True): objCol.Visible = mblnShowCard
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Sortable = True
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�в���..."
        End With
        .TabStop = True
        .PreviewMode = True
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgRPT
    
        .GroupsOrder.DeleteAll
        If intIndex = ҳ��.����� Or intIndex = ҳ��.��Ժ Then
            .GroupsOrder.Add .Columns.Find(c_����)
            .GroupsOrder(0).SortAscending = True
        End If
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns.Find(c_���)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(c_��Ժ����)
    End With
End Sub

Private Function InitBedlevel() As Boolean
'���ܣ���ʼ����λ״������
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    cbo��λ״��.Clear
    cbo��λ״��.AddItem "ȫ��"
    On Error GoTo errH
    strSQL = _
        " Select ���� from ��λ���Ʒ��� Order by DECODE(ȱʡ��־,1,1,2)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InitNurselevel")
    Do While Not rsTmp.EOF
        cbo��λ״��.AddItem rsTmp!����
        rsTmp.MoveNext
    Loop
    cbo��λ״��.ListIndex = 0

    InitBedlevel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitNurselevel() As Boolean
'���ܣ���ʼ��סԺ����ȼ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSel As String
    Dim blnSelAll As Boolean

    txt��������.Text = ""
    txt��������.Tag = ""

    lst��������.AddItem "ȫ��"
    strSel = zlDatabase.GetPara("����ȼ�����", glngSys, pסԺ��ʿվ, "", Array(txt��������, cmd��������), InStr(mstrPrivs, "��������") > 0)
    blnSelAll = True
    On Error GoTo errH
    strSQL = _
        " Select ID,����,���� From �շ���ĿĿ¼ Where ���='H' And ��Ŀ����>=1" & _
        " And (����ʱ�� is NULL Or Trunc(����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " And (վ��='" & gstrNodeNo & "' Or վ�� is Null)" & _
        " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InitNurselevel")
    Do While Not rsTmp.EOF
        lst��������.AddItem rsTmp!����
        lst��������.ItemData(lst��������.NewIndex) = rsTmp!ID
        If strSel = "" Or InStr("," & strSel & ",", "," & rsTmp!ID & ",") > 0 Then
            txt��������.Text = txt��������.Text & "," & rsTmp!����
            txt��������.Tag = txt��������.Tag & "," & rsTmp!ID
        Else
            blnSelAll = False
        End If
        rsTmp.MoveNext
    Loop

    If blnSelAll Then
        txt��������.Text = "ȫ��"
        txt��������.Tag = ""
    Else
        txt��������.Text = Mid(txt��������.Text, 2)
        txt��������.Tag = Mid(txt��������.Tag, 2)
    End If

    InitNurselevel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long

    On Error GoTo errH
    mstrUnits = GetUser����IDs

    '�����Ź۲���
    If InStr(mstrPrivs, "ȫԺ����") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=A.����ID)" & _
            " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=A.����ID)" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
    End If

    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If InStr(mstrPrivs, "ȫԺ����") > 0 Then
                If rsTmp!ID = UserInfo.����ID Then 'ֱ����������
                    Call zlControl.CboSetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
                If InStr("," & mstrUnits & ",", "," & rsTmp!ID & ",") > 0 And cboUnit.ListIndex = -1 Then
                    Call zlControl.CboSetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            Else '����ȱʡ���������Ŀ����ж��
                If rsTmp!ȱʡ = 1 And cboUnit.ListIndex = -1 Then
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
'���ܣ���ȡ�����б����ݼ�¼��
'������strIn ��������
    Dim strSQL As String
    Dim blnYN As Boolean
    
    If strIn <> "" Then blnYN = True
    If InStr(mstrPrivs, "ȫԺ����") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=A.����ID)" & _
            " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=A.����ID)" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            IIf(blnYN, " And (C.���� Like [2] Or C.���� Like [3] Or C.���� Like [3])", "") & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
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
    'װ�ص�ǰ�����Ĵ�λ
    Dim lngX As Long, lngY As Long, lngRowCount As Long
    Dim rsBeds As New ADODB.Recordset
    Dim strBriefCode As String
    
    On Error GoTo ErrHand
    
    lngX = clngX
    lngY = clngX
    lngRowCount = (PicDraw.Width - HScr.Width - 50) \ (picPati(mlngSource).Width + 15)
    Call UnloadControls
    PicDraw.Refresh
    'debug.print "ж�ش�λ��Ƭ:" & Now
    
    If mblnSupport Then
        strBriefCode = ",zlpinyincode(c.����,0,0,',',1) AS ���� "
    Else
        strBriefCode = ",zlspellcode(c.����) AS ����"
    End If
    
    '60800:������,2013-07-09,����ʾ���ɵĴ�λ
    '��ȡ�������Ĵ�λ
'    mstrSQL = " SELECT Lpad(b.����, 10, ' ') AS ����, Lpad(b.�����, 10, ' ') �����, b.��λ����, a.����" & strBriefCode & ", a.סԺ��, a.����id, a.��ҳid" & vbNewLine & _
'            " FROM ��λ״����¼ b," & vbNewLine & _
'            "     (SELECT NVL(c.����,b.����) || Decode(c.Ӥ������id, NULL, '', '֮��') ����, b.סԺ��, b.����id, b.��ҳid" & vbNewLine & _
'            "       FROM ������Ϣ b, ������ҳ c, ��Ժ���� r" & vbNewLine & _
'            "       WHERE b.����id = r.����id AND c.����id = b.����id AND b.��ҳid = c.��ҳid AND b.��ǰ����id = r.����id AND" & vbNewLine & _
'            "             (r.����id = [1] OR c.Ӥ������id = [1])) a" & vbNewLine & _
'            " WHERE b.����id = a.����id(+) AND b.����id = [1] And NOT (b.״̬='����' And b.����ID IS NULL)" & vbNewLine & _
'            " ORDER BY Lpad(b.����, 10, ' ')"
    '74214:������,2013-06-20,�����Ż�
    '�����Ż�
    '78761:������,2014-11-03,���Ű���λ���Ʊ�������
    mstrSQL = " Select LPad(b.����, 10, ' ') As ����, LPad(b.�����, 10, ' ') �����, b.��λ����, c.����" & strBriefCode & ",c.סԺ��," & vbNewLine & _
            "       c.����id, c.��ҳid,trunc(sysdate)-trunc(DECODE(C.���ʱ��,NULL,C.��Ժ����,C.���ʱ��)) as סԺ����" & vbNewLine & _
            " From ��λ״����¼ B, ������ҳ C, ��λ���Ʒ��� D" & vbNewLine & _
            " Where b.����id =[1] And (c.��ǰ����id = b.����id Or c.Ӥ������id = b.����id Or b.����ID is NULL)" & vbNewLine & _
            "      And b.����id = c.����id(+) And c.��Ժ����(+) is Null And B.��λ����=D.����(+) " & vbNewLine & _
            "      And Not (b.״̬ = '����' And b.����id Is Null)"
    If mblnCardOrder = True Then
        mstrSQL = mstrSQL & vbNewLine & " Order By LPad(b.����, 10, ' ')"
    Else
        mstrSQL = mstrSQL & vbNewLine & " Order By D.����,LPad(b.����, 10, ' ')"
    End If
    
    Set rsBeds = zlDatabase.OpenSQLRecord(mstrSQL, "װ�ص�ǰ�����Ĵ�λ", cboUnit.ItemData(cboUnit.ListIndex))
    With rsBeds
        If .RecordCount = 0 Then
            MsgBox "��ǰ������û�д�λ�����ڲ�����λ�����н�����ӣ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        Do While Not .EOF
            '�����ڴ�ӳ���¼��
            mstrFields = "��Ƭ����|��λ����|����|סԺ��|����|����|����ID|��ҳID|�໤��|�������|�ٴ�·��|���Ա�ע1|����״̬|���Ա�ע2|���Ա�ע3|����ȼ�|��������|�����|������|סԺ����"
            mstrValues = .AbsolutePosition & "|" & Trim(!��λ����) & "|" & Trim(!����) & "|" & Nvl(!סԺ��, 0) & "|" & !���� & "|" & Nvl(!����) & "|" & Nvl(!����ID, 0) & "|" & Nvl(!��ҳID, 0) & "|0|0|0||0|||0|0|" & Trim(Nvl(!�����)) & "||" & IIf(IsNull(!סԺ����), "NULL", IIf(Val("" & !סԺ����) = 0, 1, Val("" & !סԺ����)))
            Call Record_Add(mrsBedInfo, mstrFields, mstrValues)
            '��ӿհ׿�Ƭ
            Call LoadPatiCard(.AbsolutePosition, IIf(Val(lbl����(mlngSource).Tag) = 1, IIf(Trim(Nvl(!�����)) = "", "", Trim(!�����) & IIf(IsNumeric(Trim(!�����)), "_", "")), "") & Trim(!����), lngX, lngY)
            
            If Nvl(!����ID, 0) = 0 Then
                mlng�մ� = mlng�մ� + 1
            Else
                mlng�ڴ� = mlng�ڴ� + 1
            End If
            
            '������һ�ſ�Ƭ������
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
    pic��Ժ����.ZOrder 0
    
    'debug.print "��ɴ�λ��Ƭװ��:" & Now
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
    'װ�ز��ڴ��Ĳ����嵥
    Dim j As Integer
    Dim str���� As String
    Dim intPage As Integer
    Dim lngColor As Long
    Dim objItem As ReportRecordItem
    Dim objRecord As ReportRecord
    Dim objRpt As ReportControl
    Dim objParent As ReportRecord
    
    On Error GoTo ErrHand
    
    With rsPati
        '����:0-ת�ƴ����;1-��Ժ�����;2.2-��ͥ����;4-��Ժ;5-����;6-ת��
        .Filter = "���� <>'��Ժ����' " ' AND ���� <>'Ԥ��Ժ����' " ' AND ���� <>'ת�ƴ���ס����' AND ���� <>'ת��������ס����' AND ���� <>'��Ժ����ס����'"
        '.Sort = " ��Ժ���� desc "
        .Sort = "����,����2,����,��ҳID Desc"
        Do While Not .EOF
            intPage = -1
            If !���� = 0 Or !���� = 1 Or !���� = 2 Then
                intPage = 0
            ElseIf !���� = 7 Then
                intPage = 1
            ElseIf !���� = 6 Or !���� = 5 Then
                intPage = 2
            ElseIf !���� = 3.1 Or (!���� = 4 And Nvl(!����) = "") Then '��ͥ����
                intPage = 3
                mlng�Ҵ� = mlng�Ҵ� + 1
            End If
            
            If intPage > -1 And IIf(intCurPage = -1, True, intPage = intCurPage) Then
                Select Case Nvl(!����)
                Case 0
                    str���� = "��Ժ"
                Case 1
                    str���� = "ת��"
                Case 2
                    str���� = "ת����"
                Case 5
                    str���� = "��Ժ"
                Case 6
                    str���� = "����"
                End Select
                '�����ύ��������Ӹ���
                If Nvl(!����״̬, 0) <> 0 Then
                    rptPati(intPage).Columns(c_���).Visible = True
                    If objParent Is Nothing Then
                        Set objParent = Me.rptPati(intPage).Records.Add()
                    ElseIf objParent.Tag <> CStr(!����״̬) Then
                        Set objParent = Me.rptPati(intPage).Records.Add()
                    End If
                    If objParent.Tag <> CStr(!����״̬) Then
                        objParent.Tag = CStr(!����״̬)
                        objParent.Expanded = True
                        For j = 0 To rptPati(intPage).Columns.Count - 1
                            If j = c_���� Then
                                Set objItem = objParent.AddItem(Val(!����))
                                objItem.Caption = str����
                            ElseIf j = c_��� Then
                                Set objItem = objParent.AddItem(Val(Decode(Nvl(!����״̬, 0), 0, 999, !����״̬)))
                                objItem.Caption = " "
                            ElseIf j = c_���� Then
                                Set objItem = objParent.AddItem(Get��������(!����״̬))
                                objItem.ForeColor = rptPati(intPage).PaintManager.GroupForeColor
                            Else
                                Set objItem = objParent.AddItem("")
                                If j = c_ͼ�� Then objItem.Icon = Get����ͼ�����(!����״̬, False) - 1
                            End If
                            objItem.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
                        Next
                    End If
                Else
                    Set objParent = Nothing
                End If
                
                '��Ӿ���Ĳ���������(���л������)
                If Not objParent Is Nothing Then
                    Set objRecord = objParent.Childs.Add()
                Else
                    Set objRecord = Me.rptPati(intPage).Records.Add()
                End If
                
                objRecord.Tag = CStr(!����ID & "|" & !��ҳID)
                
                Set objItem = objRecord.AddItem(str����)
                objItem.Caption = str����
                
                Set objItem = objRecord.AddItem(Val(Decode(Nvl(!����״̬, 0), 0, 999, !����״̬)))
                objItem.Caption = " "
                If Nvl(rsPati!����״̬, 0) = 2 Then
                    objRecord.PreviewText = "  ����:" & GetRefuseReason(Val(!����ID), Val(!��ҳID))
                End If
                
                Set objItem = objRecord.AddItem(Nvl(!������))
                objItem.Caption = " "
                '81308:������,2015-01-19,��ʾ��������־
                '61824:������,2013-05-23,��ʾ�����ֱ�־
                If Nvl(!����״̬, 0) <> 0 Then
                    objItem.Icon = Get����ͼ�����(!����״̬, False) - 1
                ElseIf Nvl(!������) <> "" Then
                    objItem.Icon = imgRPT.ListImages("������").Index - 1
                Else
                    objItem.Icon = Val(IIf(!�Ա� = "Ů", imgRPT.ListImages("Ů��").Index, imgRPT.ListImages("����").Index)) - 1
                End If
                
                '·��״̬
                Set objItem = objRecord.AddItem(Val("" & !·��״̬))
                objItem.Caption = " "
                objItem.Icon = Get�ٴ�·�����(Val("" & !·��״̬) + 2, False) - 1
                
                objRecord.AddItem Val(!����ID)
                objRecord.AddItem Val(!��ҳID)
                objRecord.AddItem CStr(Nvl(!����))
                Set objItem = objRecord.AddItem(CStr(Nvl(!סԺ��)))
                objItem.Caption = Nvl(!סԺ��, " ")
                Set objItem = objRecord.AddItem(LPAD(Nvl(!����), 10, " "))
                objItem.Caption = CStr(Nvl(!����, " "))
                Set objItem = objRecord.AddItem(CStr(Nvl(!�Ա�, "��")))
                objItem.Caption = CStr(Nvl(!�Ա�, "��"))
                Set objItem = objRecord.AddItem(Nvl(!����, "0"))
                objItem.Caption = Nvl(!����, "0")
                Set objItem = objRecord.AddItem(Nvl(!�ѱ�, ""))
                objItem.Caption = CStr(Nvl(!�ѱ�, ""))
                Set objItem = objRecord.AddItem(Nvl(!ҽ�Ƹ��ʽ, ""))
                objItem.Caption = CStr(Nvl(!ҽ�Ƹ��ʽ, ""))
                Set objItem = objRecord.AddItem(Nvl(!סԺҽʦ, ""))
                objItem.Caption = CStr(Nvl(!סԺҽʦ, ""))
                Set objItem = objRecord.AddItem(CStr(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss")))
                objItem.Caption = CStr(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                Set objItem = objRecord.AddItem(CStr(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss")))
                objItem.Caption = CStr(Format(!��Ժ����, "yyyy-MM-dd HH:mm:ss"))
                Set objItem = objRecord.AddItem(Nvl(!��������, "��ͨ����"))
                objItem.Caption = CStr(Nvl(!��������, "��ͨ����"))
                Set objItem = objRecord.AddItem(CStr(Nvl(!���￨��)))
                objItem.Caption = Nvl(!���￨��, "")
                '��ȡ�������͵���ɫ
                lngColor = 0
                mrsPatiColor.Filter = "����='" & Nvl(!��������, "��ͨ����") & "'"
                If mrsPatiColor.RecordCount <> 0 Then
                    lngColor = Nvl(mrsPatiColor!��ɫ, 0)
                End If
                If lngColor <> 0 Then
                    objRecord.Item(c_����).ForeColor = lngColor
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    
    If intCurPage = ҳ��.����� Or intCurPage = -1 Then rptPati(ҳ��.�����).Populate 'ȱʡ��ѡ���κ���
    If intCurPage = ҳ��.ת�� Or intCurPage = -1 Then rptPati(ҳ��.ת��).Populate  'ȱʡ��ѡ���κ���
    If intCurPage = ҳ��.��Ժ Or intCurPage = -1 Then rptPati(ҳ��.��Ժ).Populate  'ȱʡ��ѡ���κ���
    If intCurPage = ҳ��.��ͥ���� Or intCurPage = -1 Then rptPati(ҳ��.��ͥ����).Populate  'ȱʡ��ѡ���κ���
    
    UpgradeList = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UpgradeBeds(ByVal rsPati As ADODB.Recordset) As Boolean
    '������Ժ���˵Ĵ�λ���ݲ���ʾ
    Dim arrBeds
    Dim i As Integer, j As Integer, lngCardIndex As Integer
    Dim lngPatiColor As Long
    Dim strDiag As String
    Dim strBeds As String, strAmountSQL As String, strDurationSQL As String
    Dim strMonitor As String
    Dim strBalance As String, strNotes As String
    Dim rsBalance As New ADODB.Recordset
    Dim rsDiagnosis As New ADODB.Recordset
    '49535,������,2012-08-14,������Ϣ���ַ������ӣ����Ϊ����
    Dim ArrPatiInfo As Variant
    On Error GoTo ErrHand
    
    '��ȡ�໤���漰����סԺ���嵥
    If mclsWardMonitor.Enabled And InStr(GetInsidePrivs(pסԺ��ʿվ), "����໤") > 0 Then
        strMonitor = mclsWardMonitor.GetListPati
    End If
    
    '��ʾ���д�λ��Ƭ(���ǵ���������������,�Ƚ���Ƭ��ʾ����)
    j = picPati.Count - 2
    For i = 1 To j
        picPati(i).Visible = True
    Next
    
    If Mid(mstrCardInfo, 2, 1) = "1" Then
        '��ȡ���������в��˵�ʵ�����
        '56960:������,2013-01-17,���������Ӱ���������
        If mblnCardBalance = True Then
            strAmountSQL = "(SELECT  NVL(SUM(NVL(������ ,0)),0)" & vbNewLine & _
                "   FROM ���˵�����¼" & vbNewLine & _
                "   WHERE ����ID = C.����ID AND ��ҳID =C.��ҳID AND" & vbNewLine & _
                "   (����ʱ�� IS NULL OR ����ʱ�� > SYSDATE) AND ɾ����־ = 1)+"
            
            strDurationSQL = ",(SELECT 1" & vbNewLine & _
                " FROM ���˵�����¼" & vbNewLine & _
                " WHERE ����ID = C.����ID AND ��ҳID = C.��ҳID AND (����ʱ�� IS NOT NULL And ����ʱ�� > SYSDATE)" & vbNewLine & _
                " And ������ = 999999999 AND ɾ����־ = 1 And RowNum < 2) ���޵�����"
        Else
            strAmountSQL = ""
            strDurationSQL = ",NULL ���޵�����"
        End If
        mstrSQL = "  Select D.����ID,D.��ҳID,D.סԺ��," & strAmountSQL & "NVL(A.Ԥ�����,0)+NVL(B.ҽ������,0)-NVL(A.�������,0) AS ���" & strDurationSQL & _
                   " From ������� A," & _
                   "      (Select B.����ID,B.��ҳID,SUM(B.���) AS ҽ������" & _
                   "      From ����ģ����� B,���㷽ʽ D,������Ϣ A,��Ժ���� R" & _
                   "      Where B.���㷽ʽ=D.���� And D.���� IN (3,4) And B.����ID=A.����ID And B.��ҳID=A.��ҳid And a.����ID=R.����ID And A.��ǰ����ID=R.����ID  And R.����ID=[1]" & _
                   "      GROUP BY B.����ID,B.��ҳID) B," & _
                   "      ������ҳ C,������Ϣ D,��Ժ���� R" & _
                   " Where A.����ID(+) =C.����ID AND A.����(+)=1 AND A.����(+)=2" & _
                   " And B.����ID(+)=C.����ID And B.��ҳID(+)=C.��ҳID" & _
                   " And D.����ID=R.����ID And D.����ID=C.����ID And D.��ҳid=C.��ҳID And D.��ǰ����ID=R.����ID And R.����ID=[1]"
        Set rsBalance = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���������в��˵�ʵ�����", cboUnit.ItemData(cboUnit.ListIndex))
    End If
    Call ShowGuage("��ȡ���������в��˵�ʵ�����", 50)
    'debug.print "...��ȡ���������в��˵�ʵ�����:" & Now
    
    If Mid(mstrCardInfo, 1, 1) = "1" Then
        '��ȡ���������в��˵������Ҫ���
        '�������:
        '    1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;
        '    10-����֢;11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
        '��¼��Դ:
        '    1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����
'        mstrSQL = " Select A.����ID,A.��ҳID,A.�������,A.��¼��Դ,A.��ϴ���,A.����ID,A.���ID,A.�������,A.�Ƿ�δ��,A.�Ƿ�����,A.��ע" & _
'                  " From ������ϼ�¼ A,������ҳ B,������Ϣ C,��Ժ���� R" & _
'                  " Where a.������� In (1, 2, 3, 11, 12, 13) And A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.����ID=C.����ID And C.��ҳid=B.��ҳID And C.����ID=R.����ID And C.��ǰ����ID=R.����ID " & _
'                  " And ��ϴ���=1 And (R.����ID=[1] Or b.Ӥ������ID=[1])" & _
'                  " Order by A.����ID asc,A.��¼��Դ desc,A.������� desc"
'        Set rsDiagnosis = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���������в��˵����", cboUnit.ItemData(cboUnit.ListIndex))
        Set rsDiagnosis = GetPatiDiagnoseByDept(cboUnit.ItemData(cboUnit.ListIndex), 1)
    End If
    Call ShowGuage("��ȡ���������в��˵���Ҫ���", 60)
    'debug.print "...��ȡ���������в��˵���Ҫ���:" & Now
    
    '�����ڴ�ӳ���¼��
    mstrFields = "����|����ȼ�|����ȼ�����|��������|�໤��|�������|�ٴ�·��|���Ա�ע1|����״̬|���Ա�ע2|���Ա�ע3|�໤������|�����������|�ٴ�·������|���Ա�ע1����|����״̬����|���Ա�ע2����|���Ա�ע3����|������"
    With rsPati
        .Filter = "���� ='��Ժ����' Or ���� ='Ԥ��Ժ����' Or ���� ='Ԥת�Ʋ���' Or ����='ת��������'"
        Do While Not .EOF
            '�ҵ��ò��˵Ĵ���
            
            '82383:�˴�������Ҫ��Ϊ������֮ǰ��ͬʱ������ZLHIS������ͬ���˻���һ�Ŵ������(���ֺͲ������ﶨλ��ͬ�Ĳ���)
            lngCardIndex = -1
            mrsBedInfo.Filter = "����='" & Trim(Nvl(!����, "ZYB")) & "'"
            If mrsBedInfo.RecordCount <> 0 Then
                If mrsBedInfo!����ID = 0 Or mrsBedInfo!����ID = !����ID Then
                    lngCardIndex = mrsBedInfo!��Ƭ����
                End If
            End If
            If lngCardIndex = -1 Then
                mrsBedInfo.Filter = "����ID=" & !����ID
                If mrsBedInfo.RecordCount <> 0 Then
                    mrsBedInfo.Filter = "����='" & Trim(Nvl(mrsBedInfo!����, "ZYB")) & "'"
                    If mrsBedInfo.RecordCount <> 0 Then lngCardIndex = mrsBedInfo!��Ƭ����
                End If
            End If
            mrsBedInfo.Filter = 0
            
            If lngCardIndex <> -1 Then
                '׼�����²��˿�Ƭ��Ϣ����
                strBalance = ""
                If Mid(mstrCardInfo, 2, 1) = "1" Then
                    rsBalance.Filter = "����ID=" & !����ID
                    If rsBalance.RecordCount <> 0 Then
                        strBalance = Format(Nvl(rsBalance!���, 0), "#0.00;-#0.00; ;")
                        If Val(Nvl(rsBalance!���޵�����, 0)) = 1 Then
                            strBalance = "���޶�ȵ���"
                        End If
                    End If
                    rsBalance.Filter = 0
                End If
                
                'סԺ��,����,�Ա�,����,���,ҽ/��,�ѱ�,ҽ�Ƹ��ʽ,����,��Ժ����,סԺ����,���,������ɫ,����ȼ�,���￨�ţ�
                '56958:������,2013-01-16,ҽ���ͻ�ʿ��ʾһ��
                If Trim(Nvl(!סԺҽʦ)) = "" And Trim(Nvl(!���λ�ʿ)) = "" Then
                    strDiag = ""
                Else
                    strDiag = Trim(Nvl(!סԺҽʦ)) & "/" & Trim(Nvl(!���λ�ʿ))
                End If
                ArrPatiInfo = Array(IIf(mblnOutDept, Nvl(!�����), Nvl(!סԺ��)), Nvl(!����), Nvl(!�Ա�), Nvl(!����), "[���]", strDiag, Nvl(!�ѱ�), Nvl(!ҽ�Ƹ��ʽ), _
                             IIf(Nvl(!��ǰ����) = "һ��", "", Nvl(!��ǰ����)), Format(!��Ժ����, "yyyy-MM-dd"), Nvl(!סԺ����), strBalance, 0, "", Nvl(!���￨��))
                '��ȡ���(��ҽ����ҽ������ȣ�Ȼ��������ͷ������ȣ�Ȼ�������Դ��������)
                strDiag = ""
                If Mid(mstrCardInfo, 1, 1) = "1" Then
                    rsDiagnosis.Filter = "����ID=" & !����ID
                    If rsDiagnosis.RecordCount <> 0 Then
                        strDiag = Nvl(rsDiagnosis!�������)
                    End If
                    rsDiagnosis.Filter = 0
                End If
                ArrPatiInfo(4) = Replace(CStr(ArrPatiInfo(4)), "[���]", strDiag)
                '��ȡ�������͵���ɫ(Ϊ�˱�����ɫ���˷�ɢ����Աע����,��ɫȱʡ����ʾ)
                mrsPatiColor.Filter = "����='" & Nvl(!��������, "��ͨ����") & "'"
                If mrsPatiColor.RecordCount <> 0 Then
                    lngPatiColor = IIf(Nvl(!��������, "��ͨ����") = "��ͨ����", &HFFFFFF, Nvl(mrsPatiColor!��ɫ, 0))
                Else
                    lngPatiColor = &HFFFFFF
                End If
                mrsPatiColor.Filter = 0
                ArrPatiInfo(12) = lngPatiColor
                ArrPatiInfo(13) = Nvl(!����ȼ�, "��������")
                
                '1�����¿�Ƭ�ϵ���Ϣ����
                Call SetCardInfo(lngCardIndex, ArrPatiInfo)
                mstrValues = Nvl(!��ǰ����) & "|" & Get����ȼ�(Nvl(!����ȼ�, "��������")) & "|" & Nvl(!����ȼ�, "��������") & "|" & Nvl(!��������, "��ͨ����")
                
                '��ȡ����
                '2�����¿�Ƭ�ϵı�ע���򣨼໤��|�������|�ٴ�·��|���Ա�ע1|����״̬|���Ա�ע2|���Ա�ע3|����ȼ���
                strNotes = UpgradeNotes(rsPati, strMonitor)
                mstrValues = mstrValues & strNotes
                Call Record_Update(mrsBedInfo, mstrFields & "|����", mstrValues & "|0", "��Ƭ����|" & lngCardIndex)
                Call SetCardLabel(lngCardIndex)
                
                '3�����°���
                strBeds = ""
                mrsBedInfo.Filter = "����ID=" & !����ID
                With mrsBedInfo
                    Do While Not .EOF
                        strBeds = strBeds & "," & !��Ƭ���� & "|" & !����
                        .MoveNext
                    Loop
                End With
                mrsBedInfo.Filter = 0
                If strBeds <> "" Then strBeds = Mid(strBeds, 2)
                arrBeds = Split(strBeds, ",")
                j = UBound(arrBeds)
                For i = 0 To j
                    If Split(arrBeds(i), "|")(0) <> lngCardIndex Then
                        'סԺ��,����,�Ա�,����,���,ҽ/��,�ѱ�,ҽ�Ƹ��ʽ,����,��Ժ����,סԺ����,���,������ɫ,����ȼ�,���￨�ţ�
                        ArrPatiInfo = Array("", Nvl(rsPati!����), "����", "", "", "", "", "", "", "", "", "", lngPatiColor, "", "")
                        Call SetCardInfo(Split(arrBeds(i), "|")(0), ArrPatiInfo)
                        
                        '���°�������Ϣ
                        Call Record_Update(mrsBedInfo, mstrFields & "|����", mstrValues & "|1", "��Ƭ����|" & Split(arrBeds(i), "|")(0))
                    End If
                Next
            End If
            
            .MoveNext
        Loop
        rsPati.Filter = 0
    End With
    
    Call ShowGuage("��ɲ�����λ�����ݸ���", 80)
    'debug.print "...��ɿ�Ƭ���ݸ���:" & Now
    
    'ͬ��ˢ����鷴����Ϣ
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
    Dim int������� As Integer, int�ٴ�·�� As Integer, int����״̬ As Integer, int�໤�� As Integer, str��ע1 As String, str��ע2 As String, str��ע3 As String
    Dim str����״̬ As String, str���Ա�ע1 As String, str���Ա�ע2 As String, str���Ա�ע3 As String, str������ As String
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    '��ȡ��ǰ���˵ı�עͼ������
    '61824:������,2013-05-23,��ʾ�����ֱ�־
    str������ = Nvl(rsPati!������)
    int������� = Nvl(rsPati!����״̬, 0)
    int�ٴ�·�� = rsPati!·��״̬ + 2
    If rsPati!���� = "3.2" Or rsPati!���� = "3.3" Then     'Ԥת��
        str����״̬ = "Ԥת��"
        int����״̬ = Img���(mlngSource).ListImages("Ԥת��").Index
    ElseIf rsPati!���� = ptԤ�� Then     'Ԥ��Ժ
        str����״̬ = "Ԥ��Ժ"
        int����״̬ = Img���(mlngSource).ListImages("Ԥ��Ժ").Index
    End If
    If strMonitor <> "" And Not IsNull(rsPati!סԺ��) Then
        If InStr("," & strMonitor & ",", "," & rsPati!סԺ�� & ",") > 0 Then
            int�໤�� = 1
        End If
    End If
    
    'ͼ������+1����Ϊ��ע�����Ǵ�0��ʼ
    mrsPatiNotes.Filter = "����ID=" & rsPati!����ID & " And ��ҳID=" & rsPati!��ҳID
    mrsPatiNotes.Sort = "���˳��"
    Do While Not mrsPatiNotes.EOF
        i = Val("" & mrsPatiNotes!���˳��)
        If i = 1 Then
            str��ע1 = mrsPatiNotes!���ⲡ��ID & "," & mrsPatiNotes!������� & "," & mrsPatiNotes!������ & "," & mrsPatiNotes!ͼ������ + 1
        ElseIf i = 2 Then
            str��ע2 = mrsPatiNotes!���ⲡ��ID & "," & mrsPatiNotes!������� & "," & mrsPatiNotes!������ & "," & mrsPatiNotes!ͼ������ + 1
        ElseIf i = 3 Then
            str��ע3 = mrsPatiNotes!���ⲡ��ID & "," & mrsPatiNotes!������� & "," & mrsPatiNotes!������ & "," & mrsPatiNotes!ͼ������ + 1
        End If
        mrsNotes.Filter = "����ID=" & mrsPatiNotes!���ⲡ��ID & " And �������=" & mrsPatiNotes!������� & " And ������=" & mrsPatiNotes!������
        If mrsNotes.RecordCount <> 0 Then
            str���Ա�ע1 = mrsNotes!˵��
            If i = 1 Then
                str���Ա�ע1 = mrsNotes!˵��
            ElseIf i = 2 Then
                str���Ա�ע2 = mrsNotes!˵��
            ElseIf i = 3 Then
                str���Ա�ע3 = mrsNotes!˵��
            End If
        End If
        mrsPatiNotes.MoveNext
    Loop

    mrsPatiNotes.Filter = ""
    mrsNotes.Filter = ""

    UpgradeNotes = "|" & int�໤�� & "|" & int������� & "|" & int�ٴ�·�� & "|" & str��ע1 & "|" & int����״̬ & "|" & str��ע2 & "|" & str��ע3 & "|" & _
                   IIf(int�໤�� > 0, "�໤��", "") & "|" & Get��������(int�������) & "|" & Get�ٴ�·������(int�ٴ�·��) & "|" & str���Ա�ע1 & "|" & str����״̬ & "|" & str���Ա�ע2 & "|" & str���Ա�ע3 & "|" & str������
End Function

Private Function Get�ٴ�·�����(ByVal lng״̬ As Long, Optional ByVal blnCard As Boolean = True) As Long
    Dim imgList As ImageList
    If blnCard = True Then
        Set imgList = Img���(mlngSource)
    Else
        Set imgList = imgRPT
    End If
    Get�ٴ�·����� = Choose(lng״̬, imgList.ListImages("δ����").Index, imgList.ListImages("������").Index, _
            imgList.ListImages("ִ����").Index, imgList.ListImages("��������").Index, imgList.ListImages("�������").Index)
End Function

Private Function Get�ٴ�·������(ByVal lng״̬ As Long) As String
    Get�ٴ�·������ = Choose(lng״̬, "δ����", "������", "ִ����", "��������", "�������")
End Function

Private Function Get����ͼ�����(ByVal lng״̬ As Long, Optional ByVal blnCard As Boolean = True) As Long
    Dim i As Long
    Dim imgList As ImageList
    
    If blnCard = True Then
        Set imgList = Img���(mlngSource)
    Else
        Set imgList = imgRPT
    End If
    Select Case lng״̬
        Case 1
            i = imgList.ListImages("�ȴ����").Index
        Case 2
            i = imgList.ListImages("�ܾ����").Index
        Case 13
            i = imgList.ListImages("���ڳ��").Index
        Case 3
            i = imgList.ListImages("�������").Index
        Case 14
            i = imgList.ListImages("��鷴��").Index
        Case 4
            i = imgList.ListImages("��鷴��").Index
        Case 16
            i = imgList.ListImages("�������").Index
        Case 6
            i = imgList.ListImages("�������").Index
        Case 5
            i = imgList.ListImages("���鵵").Index
        Case 10
            i = imgList.ListImages("�ȴ����").Index
    End Select
    Get����ͼ����� = i
End Function

Private Function Get��������(ByVal lng״̬ As Long) As String
    Dim i As Long
    
    Select Case lng״̬
        Case 1
            Get�������� = "�ȴ����" '�ύ����
        Case 2
            Get�������� = "�ܾ����" '�ܾ�����
        Case 13
            Get�������� = "���ڳ��"
        Case 3
            Get�������� = "�������"
        Case 14
            Get�������� = "��鷴��"
        Case 4
            Get�������� = "��鷴��"
        Case 16
            Get�������� = "�������"
        Case 6
            Get�������� = "�������"
        Case 10
            Get�������� = "���մ���"
        Case 5
            Get�������� = "���鵵"
    End Select
End Function

Private Function GetVersion() As String
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    strSQL = " select �汾�� from zlsystems where ���=100"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��׼�汾��")
    GetVersion = rsTemp!�汾��
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadPatients(ByRef rsPati As ADODB.Recordset) As Boolean
'���ܣ���ȡ�����б�
    Dim strSQL As String
    Dim int��Ժ���� As Integer, strPatiFileter As String
    '�޸Ĵ�SQL������,�����������ģ��Ҳ��Ҫ����
    '61824:������,2013-05-23,��ʾ�����ֱ�־
    'ת�ƴ���Ʋ���
    
    '��ҳ����������գ�F5ˢ�£�Ӧ�ûָ���һ����ֵ
    If cboUnit.ListIndex = -1 Then Call zlControl.CboSetIndex(cboUnit.hwnd, mintPreDept)
    
    '111016:��Ժ����Ʋ��˹���,Ϊ0��ʾ������
    int��Ժ���� = zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, 0)
    If int��Ժ���� > 0 Then
        strPatiFileter = " And B.��Ժ����>=Sysdate-[2]"
    End If
    
    If Val(Mid(mstrScope, 5, 1)) <> 0 Then
        '84938:������,����Ʋ�����ȡ�Ż�(���A.��ҳID=B.��ҳID)
        strSQL = _
            "Select /*+ RULE */Distinct" & vbNewLine & _
            " Decode(B.״̬,1,0,Decode(c.��ʼԭ��,3,1,2)) As ����, Decode(Nvl(b.����״̬, 0), 0, 999, b.����״̬) As ����2," & _
            " Decode(B.״̬,1,'��Ժ����ס����',Decode(c.��ʼԭ��,3,'ת�ƴ���ס����','ת��������ס����')) As ����," & _
            " a.����id, b.��ҳid, A.�����,B.סԺ��, NVL(B.����,A.����) ����" & mstrBriefCode & ", NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����," & vbNewLine & _
            " d.���� As ����, c.����id, c.����ҽʦ As סԺҽʦ,b.���λ�ʿ, b.����״̬, c.����," & _
            " e.���� As ����ȼ�, b.�ѱ�,B.ҽ�Ƹ��ʽ,b.��ǰ����, Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����, b.��Ժ����,B.��Ժ��ʽ, b.��������, b.״̬, b.����, a.���￨��," & vbNewLine & _
            " -1 As ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID,A.��ҳId �����ҳId" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, ���ű� D, �շ���ĿĿ¼ E,�������� Z" & vbNewLine & _
            "Where a.��Ժ = 1 And B.��������=Z.����(+) And a.����id = b.����id And A.��ҳID=B.��ҳID And Nvl(b.��ҳid, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid " & vbNewLine & _
            "      And (C.����ID=[1] or C.����ID is null) And c.����id = d.Id" & vbNewLine & _
            "      And (d.վ��='" & gstrNodeNo & "' Or d.վ�� is Null)" & vbNewLine & _
            "      And b.����ȼ�id = e.Id(+) And Nvl(c.���Ӵ�λ, 0) = 0 And c.��ֹʱ�� Is Null" & vbNewLine & _
            "      And (c.��ʼԭ�� in(1,3) And Exists(Select 1 From �������Ҷ�Ӧ H Where c.����id = h.����id And h.����id = [1]) or c.��ʼԭ��=15 And c.����id = [1])" & vbNewLine & _
            "      And ((c.��ʼԭ�� = 1 And b.״̬ = 1 " & strPatiFileter & ") Or (c.��ʼԭ�� in (3,15) And c.��ʼʱ�� Is Null And b.״̬ = 2)) "
    
    End If
    '��Ժ���ˣ���λһ�����ģʽ��������ʾ��Ժ���ˣ�
    strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
        "Select /*+ RULE */ Decode(B.״̬,3,4,DECODE(B.��Ժ����, NULL, 3.1,DECODE(B.״̬,2,3.2,3))) as ����," & _
        " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
        " Decode(B.״̬,3,'Ԥ��Ժ����',DECODE(B.��Ժ����, NULL, '��ͥ����',DECODE(B.״̬,2,'Ԥת�Ʋ���', '��Ժ����'))) as ����," & _
        " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(B.����,A.����) ����" & mstrBriefCode & ",NVL(b.�Ա�,a.�Ա�) �Ա�,NVL(b.����,a.����) ����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
        " B.��Ժ���� as ����,E.���� as ����ȼ�,B.�ѱ�,B.ҽ�Ƹ��ʽ,B.��ǰ����,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
        " B.״̬,B.����,A.���￨��,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID,A.��ҳId �����ҳId" & _
        " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z,��Ժ���� R" & _
        " Where B.��������=Z.����(+) And A.����ID=B.����ID And A.��ҳID=B.��ҳID And Nvl(B.״̬,0)<>1" & _
        " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And R.����ID=[1] And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
        " And a.����ID=R.����ID And A.��ǰ����ID=R.����ID And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
    strSQL = strSQL & " Order by ����,����2,����,��ҳID Desc"
    
    On Error GoTo errH
    Set rsPati = New ADODB.Recordset
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), int��Ժ����)
    
    rsPati.Filter = "����='Ԥ��Ժ����'"
    mlngԤ��Ժ = rsPati.RecordCount
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
    'strKeys��Ϊ����ֱ�Ӹ��ݲ��˹��ˣ�˵���ǹ���������
    Dim i As Integer, j As Integer
    Dim blnAdjust As Boolean
    Dim lngX As Long, lngRowCount As Long, lngShowed As Long
    Dim lng����ID As Long, lngIndex As Long
    Dim blnShowCard As Boolean
    'ֻ���л�������ʱ������¶�ȡ����,�����ڵ������仯,ֻ�ǽ���Ƭ���غ���������λ�ü���
    
    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hwnd)
    
    '�������д�λ��Ƭ
    mintCards = 0
    lng����ID = mlng����ID
    mlng����ID = 0
    mstrBoardKeys = strKeys
    j = picPati.Count - 2
    For i = 1 To j
        picPati(i).Visible = False
    Next
    
    If j = 0 Then Exit Sub
    blnAdjust = (lngY = clngX)
    '���½����������
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
                blnShowCard = (InStr(1, "," & strKeys & ",", "," & Nvl(mrsBedInfo!����ID) & ",") <> 0)
            End If
            If blnShowCard Then
                If !����ID = lng����ID And lng����ID <> 0 Then
                    lngIndex = !��Ƭ����
                End If
                lngShowed = lngShowed + 1
                With picPati(!��Ƭ����)
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
                
                '������һ�ſ�Ƭ������
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
    pic��Ժ����.ZOrder 0
    
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
            mlng����ID = lng����ID
        End If
    End If

    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(0)
End Sub

Private Function ISShowCard() As Boolean
    Dim arr����
    Dim strInfo As String, int������� As Integer
    Dim i As Integer, j As Integer
    Dim arrSignNotes(0 To 2) As String, arrNote(0 To 2) As String
    
    '�жϵ�ǰ��Ƭ�Ƿ��������
    int������� = zlDatabase.GetPara("�������", glngSys, mlngModul, 0)
    ISShowCard = (chk�����մ�.Value = 1 Or Not (chk�����մ�.Value = 0 And Nvl(mrsBedInfo!����ID, 0) = 0))
    If ISShowCard Then
        '��������
        Select Case Nvl(mrsBedInfo!����)
        Case "Σ"
            ISShowCard = (chk��������(1).Value = 1)
        Case "��"
            ISShowCard = (chk��������(2).Value = 1)
        Case Else
            ISShowCard = (chk��������(0).Value = 1)
        End Select
    End If
    If ISShowCard And cbo��λ״��.Text <> "ȫ��" Then
        '���ݻ���ȼ����������ж�
        ISShowCard = (mrsBedInfo!��λ���� = cbo��λ״��.Text)
    End If
    If ISShowCard And txt��������.Text <> "ȫ��" Then
        '���ݻ���ȼ����������ж�
        ISShowCard = (InStr(1, "," & txt��������.Text & ",", "," & mrsBedInfo!����ȼ����� & ",") <> 0)
    End If
    If ISShowCard Then
        '�������
        If Me.cbo����.Text <> "����" Then strInfo = cbo����.Text
        If Me.cbo����.Text <> "����" Then
            Select Case Me.cbo����.ListIndex
            Case 1
                If Me.cbo����.Text = "����" Then
                    ISShowCard = (mrsBedInfo!������� <> 0)
                Else
                    ISShowCard = (Nvl(mrsBedInfo!�����������) = strInfo)
                End If
            Case 2
                If Me.cbo����.Text = "����" Then
                    ISShowCard = (mrsBedInfo!�ٴ�·�� <> 0)
                Else
                    ISShowCard = (Nvl(mrsBedInfo!�ٴ�·������) = strInfo)
                End If
            Case 3
                On Error GoTo errH
                If Me.cbo����.Text = "����" Then
                    ISShowCard = (mrsBedInfo!����״̬ <> 0)
                    If Not ISShowCard Then
                        If mrsBedInfo!����ID <> 0 Then
                            If Not IsNull(mrsBedInfo!סԺ����) Then
                                ISShowCard = (Val(mrsBedInfo!סԺ����) <= int�������)
                            Else
                                ISShowCard = False
                            End If
                        Else
                            ISShowCard = False
                        End If
                    End If
                ElseIf Me.cbo����.Text Like "���*����" Then
                    If mrsBedInfo!����ID <> 0 Then
                        If Not IsNull(mrsBedInfo!סԺ����) Then
                            ISShowCard = (Val(mrsBedInfo!סԺ����) <= int�������)
                        Else
                            ISShowCard = False
                        End If
                    Else
                        ISShowCard = False
                    End If
                Else
                    ISShowCard = (Nvl(mrsBedInfo!����״̬����) = strInfo)
                End If
            Case Is > 3 '���Ա�ע
                ISShowCard = False
                If Nvl(mrsBedInfo!���Ա�ע1) <> "" Then
                    arrSignNotes(0) = Split(mrsBedInfo!���Ա�ע1, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע1, ",")(1)
                    arrNote(0) = Split(mrsBedInfo!���Ա�ע1, ",")(2)
                End If
                If Nvl(mrsBedInfo!���Ա�ע2) <> "" Then
                    arrSignNotes(1) = Split(mrsBedInfo!���Ա�ע2, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע2, ",")(1)
                    arrNote(1) = Split(mrsBedInfo!���Ա�ע2, ",")(2)
                End If
                If Nvl(mrsBedInfo!���Ա�ע3) <> "" Then
                    arrSignNotes(2) = Split(mrsBedInfo!���Ա�ע3, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע3, ",")(1)
                    arrNote(2) = Split(mrsBedInfo!���Ա�ע3, ",")(2)
                End If
                
                If Me.cbo����.Text = "����" Then
                    mrsNotes.Filter = "������=0"
                Else
                    mrsNotes.Filter = "������>0"
                End If
                mrsNotes.Sort = "����ID,�������"
                Do While Not mrsNotes.EOF
                    If Val(mrsNotes!����ID) + Val(mrsNotes!�������) = Val(cbo����.ItemData(cbo����.ListIndex)) Then
                        For i = 0 To UBound(arrSignNotes)
                            If arrSignNotes(i) = mrsNotes!����ID & "," & mrsNotes!������� Then
                                If Me.cbo����.Text = "����" Then
                                    ISShowCard = True
                                Else
                                    If Val(arrNote(i)) = Val(cbo����.ItemData(cbo����.ListIndex)) Then
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

    '�����������
    Set mcolSubForm = New Collection
    mcolSubForm.Add mclsAdvices.zlGetForm, "_ҽ��"
    mcolSubForm.Add mclsFeeQuery.zlGetForm, "_����"
    If mclsWardMonitor.Enabled Then
        mcolSubForm.Add mclsWardMonitor.zlGetForm, "_�໤"
    End If
End Sub

Private Sub AddSendCommandBar()
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim strPrivs As String, strPara As String
    Dim strUnit As String
    Dim i As Long
    '61762:������,2013-05-20,���ӷ�����ҺҩƷҽ���Ĺ���
    If gstr��Һ�������� <> "" Then
        strUnit = cboUnit.ItemData(cboUnit.ListIndex)
        strPrivs = GetInsidePrivs(pסԺҽ������)
        If InStr(";" & strPrivs & ";", ";����ҩ������;") = 0 Or InStr(";" & strPrivs & ";", ";����ҩ�Ƴ���;") = 0 Then
            strPrivs = ""
        End If
    End If
    
    strPara = zlDatabase.GetPara("����Һ�������ķ�ҩ�Ĳ��˿���", glngSys, pסԺҽ���´�)
    If strPara = "*" Then strUnit = "*"
    'һ��������������ҽ�����Ͳ˵����
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(3)
    'ɾ������ҽ����ť
    For i = cbrMenuBar.CommandBar.Controls.Count To 1 Step -1
        If cbrMenuBar.CommandBar.Controls(i).ID = conMenu_Edit_Send Then
            cbrMenuBar.CommandBar.Controls(i).Delete
        End If
    Next i
    '���ҽ����ť
    With cbrMenuBar.CommandBar.Controls
        '���ҵ�����֮ǰ��У�԰�ť
        Set cbrControl = .Find(, conMenu_Edit_Audit)
        If Not cbrControl Is Nothing Then
            If strPrivs <> "" Then
                Set cbrMenuBar = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "ҽ������(&4)", cbrControl.Index + 1)
                cbrMenuBar.CommandBar.Title = "������������"
                cbrMenuBar.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Send, "��������ҽ��(&S)"
                If InStr(1, "," & strPara & ",", "," & strUnit & ",") > 0 Then
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "������ҺҩƷ(&I)")
                Else
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "���;���Ӫ��ҩƷ(&I)")
                End If
                cbrControl.IconId = conMenu_Edit_Send
            Else
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "ҽ������(&4)", cbrControl.Index + 1): cbrControl.ToolTipText = ""
            End If
        End If
    End With
    
    '������������ҽ��ҵ���Ͱ�ť���
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(4)
    'ɾ������ҽ����ť
    For i = cbrMenuBar.CommandBar.Controls.Count To 1 Step -1
        If cbrMenuBar.CommandBar.Controls(i).ID = conMenu_Edit_Send Then
            cbrMenuBar.CommandBar.Controls(i).Delete
        End If
    Next i
    '���ҽ�����Ͱ�ť
    With cbrMenuBar.CommandBar.Controls
        '���ҵ�����֮ǰ��У�԰�ť
        Set cbrControl = .Find(, conMenu_Edit_Price)
        If Not cbrControl Is Nothing Then
            If strPrivs <> "" Then
                Set cbrMenuBar = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "����(&G)", cbrControl.Index + 1)
                cbrMenuBar.CommandBar.Title = "ҽ��ҵ��"
                cbrMenuBar.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Send, "��������ҽ��(&1)"
                If InStr(1, "," & strPara & ",", "," & strUnit & ",") > 0 Then
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "������ҺҩƷ(&2)")
                Else
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "���;���Ӫ��ҩƷ(&2)")
                End If
                cbrControl.IconId = conMenu_Edit_Send
            Else
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "����(&G)", cbrControl.Index + 1)
            End If
        End If
    End With
    '����������ҽ�����Ͳ˵����
    'ɾ������ҽ����ť
    For i = cbsMain(2).Controls.Count To 1 Step -1
        If cbsMain(2).Controls(i).ID = conMenu_Edit_Send Then
            cbsMain(2).Controls(i).Delete
        End If
    Next i
    
    '���ҽ�����Ͱ�ť
    With cbsMain(2).Controls
        '���ҵ�����֮ǰ��У�԰�ť
        Set cbrControl = .Find(, conMenu_Edit_Audit)
        If Not cbrControl Is Nothing Then
            If strPrivs <> "" Then
                Set cbrMenuBar = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "����", cbrControl.Index + 1): cbrMenuBar.Style = xtpButtonIconAndCaption
                cbrMenuBar.CommandBar.Title = "������������"
                cbrMenuBar.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Send, "��������ҽ��(&S)"
                If InStr(1, "," & strPara & ",", "," & strUnit & ",") > 0 Then
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "������ҺҩƷ(&I)")
                Else
                    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "���;���Ӫ��ҩƷ(&I)")
                End If
                cbrControl.IconId = conMenu_Edit_Send
            Else
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "����", cbrControl.Index + 1): cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "ҽ������"
            End If
        End If
    End With
    
    cbsMain.RecalcLayout
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
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
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = imgPublic.Icons
    
    '�˵�����
    '-----------------------------------------------------
    Set objFile = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False) '����
    objFile.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objFile.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintBedCard, "��ӡ��ͷ��(&K)��")  '��ӡ��ͷ��
        '49854:������,2013-10-31,���������ӡ
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print_Label, "��ӡ���(&W)��")  '��ӡ���
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "��ӡһ���嵥(&D)��", 1)
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintPageSet, "��ӡ��ҳ(&Z)��", 1)
        objControl.Parameter = "100,ZL1_INSIDE_1139_2"
        objControl.IconId = conMenu_ReportPopup * 100#      'ȡ��һ���˵����ͼ��
        Set objControl = .Add(xtpControlButton, conMenu_ReportPopup * 100# + 91, "סԺ�����ձ�(&R)��", 1)
        objControl.Parameter = "100,ZL1_INSIDE_1132"
        objControl.IconId = conMenu_ReportPopup * 100#      'ȡ��һ���˵����ͼ��

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_File_MedRec, "��ҳ��ӡ(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_File_MedRecSetup, "��ӡ����(&S)", -1, False
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "��ӡԤ��(&V)", -1, False
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "��ӡ��ҳ(&P)", -1, False
        End With

        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set mobjPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "�������(&P)", -1, False) '����
    mobjPopup.ID = conMenu_ManagePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    mobjPopup.CommandBar.Title = "�������"
    With mobjPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_In, "��ס(&I)"): objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Turn, "ת��(&C)"): objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TurnUnit, "ת����(&D)"): objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TurnTeam, "תС��(&T)"): objControl.Category = "����"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Bed, "����(&B)"): objControl.BeginGroup = True: objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TransposeBed, "��λ�Ի�(&Q)"): objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_House, "����(&H)"): objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_BedGrid, "���Ĵ�λ�ȼ�(&G)"): objControl.Category = "����"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_PatiInfo, "����סԺ��Ϣ(&P)"): objControl.BeginGroup = True: objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_PaitNote, "���˱�ע��Ϣ(&F)"): objControl.Category = "����"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Out, "��Ժ(&O)"): objControl.BeginGroup = True: objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InPati, "תΪסԺ����(&Z)"): objControl.Category = "����"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Baby, "�������Ǽ�(&N)"): objControl.BeginGroup = True: objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_ReCalcFee, "���ѱ��������(&R)"): objControl.Category = "����"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InsureSel, "ҽ������ѡ��(&M)"): objControl.Category = "����"
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Change_Undo, "����(&U)"): objPopup.BeginGroup = True: objPopup.Category = "����"
        objPopup.IconId = conMenu_Edit_Untread
        
        '�໤��
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Monitor, "�໤��(&N)")
        objControl.BeginGroup = True
        objControl.Category = "����"
    End With

    Set mobjPopupBatch = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "������������(&B)", -1, False)  '����
    mobjPopupBatch.ID = conMenu_ManagePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    mobjPopupBatch.CommandBar.Title = "������������"
    With mobjPopupBatch.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "Ԥ����(&1)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintMultiBill, "�߿�(&2)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "ҽ��У��(&3)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "ҽ������(&4)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "ҽ����ͣ(&5)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ҽ������(&6)"): objControl.ToolTipText = ""
        '67386:������,2013-12-20,�������ҽ��ȷ��ֹͣ��ҽ�������˶Թ���
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "ȷ��ֹͣ(&7)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Report_Reports, "��ӡִ�е�(&8)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Report_DrugQuery, "��ҩ��ѯ(&9)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "ҩƷ����Ǽ�(&J)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "���ڷ����ջ�(&S)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "ҽ������ִ��(&B)"): objControl.IconId = 3587: objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "ҽ�������˶�(&T)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AnimalHeat, "����¼�����µ�(&A)"): objControl.BeginGroup = True: objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NurseLogFile, "����¼���¼��(&L)"): objControl.ToolTipText = ""
        Set objControl = .Add(xtpControlButton, conMenu_����������, "����������(&0)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_ProveCollect, "����ɼ�����վ(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatUnPack, "�������(&U)"): objControl.BeginGroup = True: objControl.IconId = 3051
        If gbln����Ӱ����ϢϵͳԤԼ = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrintBat, "������ӡԤԼ��(&R)"): objControl.BeginGroup = True: objControl.IconId = 103
        End If
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҽ��ҵ��(&A)", -1, False)   '���У�ҽ��A������F������E������L
    objMenu.ID = conMenu_EditPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    objMenu.CommandBar.Title = "ҽ��ҵ��"
    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_�鿴ҽ��, "�鿴ҽ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�¿�(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "У��(&J)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Price, "�Ƽ�(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����(&G)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ֹͣ(&S)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "ȷ��ֹͣ(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "����(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "��ͣ(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "���������(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Notify, "ˢ������(&N)"): objControl.BeginGroup = True
    End With
    
    '63608:������,2013-07-22,�޸ķ���ҵ��Ŀ�ݼ�ΪC
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����ҵ��(&C)", -1, False) '���У�ҽ��A������C������E������L
    objMenu.ID = conMenu_EditPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    objMenu.CommandBar.Title = "����ҵ��"
    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_�鿴����, "�鿴����(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Billing, "����(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Balance, "����(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBillingApply, "��������(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReBillingAudit, "�������(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "Ԥ����(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_ReCalcFee, "���ѱ��������(&R)")
    End With
'
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����ҵ��(&L)", -1, False) '���У�ҽ��A������F������E������L
'    objMenu.ID = conMenu_EditPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
'    objMenu.CommandBar.Title = "����ҵ��"
'    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_�鿴���µ�, "�鿴���µ�(&T)")
'        Set objControl = .Add(xtpControlButton, conMenu_�鿴�����¼, "�鿴�����¼��(&H)")
'        Set objControl = .Add(xtpControlButton, conMenu_�鿴������, "�鿴������(&B)")
'    End With
'
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����ҵ��(&E)", -1, False) '���У�ҽ��A������F������E������L
'    objMenu.ID = conMenu_EditPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
'    objMenu.CommandBar.Title = "����ҵ��"
'    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_�鿴����, "�鿴����(&E)")
'    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)  '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "�����С(&N)") '����
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_FontSize_S, "С����(&S)", -1, False '����(С�����ӦС��Ƭ���������Ӧ��Ƭ)
            .Add xtpControlButton, conMenu_View_FontSize_L, "������(&L)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "��Ƭ�۵�(&C)") '����

        Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "���ڴ�����"): objControl.BeginGroup = True '����
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "��ʾ�����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_View_NoticBoard, "������"): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "���ҷ�ʽ(&Y)"): objPopup.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&N)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
        '53132:������,2013-11-08,��Ӳ��˵�����Ϣ�鿴
        Set objControl = .Add(xtpControlButton, conMenu_View_Warrant, "������Ϣ����(&W)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "���ϲο�(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "������ϲο�(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "���ƴ�ʩ�ο�(&C)", -1, False
        End With
        '54621:������,2013-02-28,��ʿվ�����ҳ������
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "��ҳ����(&M)")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditResponse, "��鷴��(&S)")
            objControl.BeginGroup = True
            objControl.ToolTipText = "�����鿴������鷴��"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "������Ŀ��������(&C)")
            objControl.BeginGroup = True
        'Set objControl = .Add(xtpControlButton, conMenu_Tool_UnitSubject, "�����������(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_UnitNBoard, "��������������(&B)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False) '����
    objMenu.ID = conMenu_HelpPopup
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)") '����
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName) '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True '����
    End With
    cbsMain(1).EnableDocking xtpFlagHideWrap

    '����������:���������Թ���
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("�������񹤾���", xtpBarTop)      '����
    objBar.Title = "������������"
    objBar.EnableDocking xtpFlagStretched
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_PreBalance, "Ԥ��"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "����Ԥ��"
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintMultiBill, "�߿�"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "�����߿�"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "У��"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ҽ��У��": objControl.BeginGroup = True
        '59098:������,2013-04-18,ҽ�����͡���ͣ��������ʾ��Ϣ����Ͳ˵�ID����
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ҽ������"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "��ͣ"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ҽ����ͣ": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ҽ������"
        '67386:������,2013-12-20,�������ҽ��ȷ��ֹͣ��ҽ�������˶Թ���
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "ȷ��ֹͣ"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ȷ��ֹͣ": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Report_Reports, "ִ�е�"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "��ӡִ�е�": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Report_DrugQuery, "��ҩ"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "��ҩ��ѯ"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "����"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "����Ǽ�"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "�����ջ�"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "���ڷ����ջ�"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "ִ�еǼ�"): objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 3587: objControl.ToolTipText = "ҽ������ִ�еǼ�"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "�˶�"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ҽ������ִ�к˶�"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AnimalHeat, "���µ�"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "����¼�����µ�": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NurseLogFile, "��¼��"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "����¼���¼��"
        Set objControl = .Add(xtpControlButton, conMenu_����������, "��������"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "����������": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_ProveCollect, "����ɼ�"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "����ɼ�����վ": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatUnPack, "���"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "�������": objControl.BeginGroup = True: objControl.IconId = 3051
        If gbln����Ӱ����ϢϵͳԤԼ = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrintBat, "ԤԼ��"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "������ӡԤԼ��": objControl.BeginGroup = True: objControl.IconId = 103
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "�˳�": objControl.BeginGroup = True
    End With
    '���⴦��
    '-----------------------------------------------------
    '�������Ҳಡ��������ѡ��
    With objBar.Controls
        Set objControl = .Add(xtpControlLabel, 99999901, "����")
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, 99999901, "����")
        objCustom.Handle = Me.cboUnit.hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    
    '����������:��������
    '-----------------------------------------------------
    Set mobjFilter = cbsMain.Add("���˹�����", xtpBarTop)   '����
    mobjFilter.EnableDocking xtpFlagStretched
    mobjFilter.ContextMenuPresent = False
    With mobjFilter.Controls
        Set objControl = .Add(xtpControlLabel, 1, "����ȼ�")
        Set objCustom = .Add(xtpControlCustom, 2, "")
        objCustom.Handle = pic����ȼ�.hwnd
        Set objControl = .Add(xtpControlLabel, 3, "��λ״��"): objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, 4, "")
        objCustom.Handle = pic��λ״��.hwnd
        Set objControl = .Add(xtpControlLabel, 5, "��ǰ����"): objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, 6, "")
        objCustom.Handle = pic����.hwnd
        Set objCustom = .Add(xtpControlCustom, 7, ""): objCustom.BeginGroup = True
        objCustom.Handle = pic�������.hwnd
        Set objCustom = .Add(xtpControlCustom, 8, "")
        objCustom.Handle = chk�����մ�.hwnd: objCustom.BeginGroup = True
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "�������Ų���")
        objPopup.Caption = "�������Ų���"
        objPopup.ID = conMenu_View_FindType
        objPopup.Style = xtpButtonCaption
        objPopup.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, 10, "")
        objCustom.Handle = txtFind.hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
'        .Add 0, vbKeyF1, conMenu_Edit_Audit         'ҽ��У��
'        .Add 0, vbKeyF2, conMenu_Edit_Send          'ҽ������
'        .Add 0, vbKeyF3, conMenu_Report_Reports     '��ӡִ�е�
'        .Add 0, vbKeyF4, conMenu_Report_DrugQuery   '��ҩ��ѯ
'        .Add 0, vbKeyF6, conMenu_Edit_PreBalance    'Ԥ����
'        .Add 0, vbKeyF7, conMenu_File_PrintMultiBill '�߿�
'        .Add 0, vbKeyF8, conMenu_Edit_BatExecute       'ִ�еǼ�
'        .Add 0, vbKeyF9, conMenu_Edit_AnimalHeat    '����¼�����µ�
'        .Add 0, vbKeyF10, conMenu_Edit_NurseLogFile '����¼���¼��
        
        .Add FCONTROL, vbKeyF, conMenu_View_Find '���Ҳ���
'        .Add 0, vbKeyF3, conMenu_View_FindNext      '������һ��
        .Add 0, vbKeyF10, conMenu_View_Notify       'ҽ������
        .Add 0, vbKeyF5, conMenu_View_Refresh       'ˢ��
        .Add 0, vbKeyF4, conMenu_View_NoticBoard    '������
        .Add 0, vbKeyF12, conMenu_File_Parameter    '��������
    End With

    '��ȡ��������ģ��ı���(��������ģ���,������ҳ��סԺ�����ձ����߿���߿������ʾ,�����ֹ��ӵ��ļ��˵���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1261_5", "ZL1_INSIDE_1261_4", "ZL1_INSIDE_1261_6", "ZL1_INSIDE_1132", "ZL1_INSIDE_1139_1", "ZL1_INSIDE_1139_2", "ZL1_INSIDE_1139_3", "ZL1_INSIDE_1261_7", "ZL1_INSIDE_1261_8")
    
    '�ٴ����ҳ�ؼ�
    With PatiPage
        With .PaintManager
            .Color = xtpTabColorOffice2003
            .Appearance = xtpTabAppearanceVisualStudio
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(ҳ��.�����, "�����", rptPati(ҳ��.�����).hwnd, 0).Tag = "�����"
        .InsertItem(ҳ��.ת��, "���ת��", picPara(0).hwnd, 0).Tag = "���ת��"
        .InsertItem(ҳ��.��Ժ, "�����Ժ", picPara(1).hwnd, 0).Tag = "�����Ժ"
        .InsertItem(ҳ��.��ͥ����, "��ͥ����", rptPati(ҳ��.��ͥ����).hwnd, 0).Tag = "��ͥ����"
    End With
    
    '53740:������,2012-09-19,������ҹ��ܲ˵�
    Call DefCommandPlugIn(cbsMain)
End Sub

Private Sub DefCommandPlugIn(ByRef cbsMain As Object)
'���ܣ���Ҳ����˵����룬�˵����͹�����
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
    Call mobjPlugIn.Initialize(gcnOracle, glngSys, P�°滤ʿվ, 1)
    strFunc = mobjPlugIn.GetFuncNames(glngSys, P�°滤ʿվ, 1, strFuncXML)
    If strFunc = "" Then Exit Sub
        Call MakePlugInBar(strFunc, strFuncXML, rsBar)
    Set mrsPlugInBar = zlDatabase.CopyNewRec(rsBar)
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    On Error GoTo errH
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    '�˵���
    If Not objMenu Is Nothing Then
        rsBar.Filter = "IsInTool=1 and BarType=1"
        If Not rsBar.EOF Then
            rsBar.Sort = "���"
            With objMenu.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!������)
                        objControl.IconId = rsBar!ͼ��ID
                        objControl.Parameter = rsBar!������
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
            rsBar.Sort = "���"
            Set objPopup = objMenu.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����", , False)
                objPopup.BeginGroup = True
                
            With objPopup.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                    objControl.IconId = rsBar!ͼ��ID
                    objControl.Parameter = rsBar!������
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
    
    '��������ť
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
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!������, lngTmp + 1)
                    objControl.IconId = rsBar!ͼ��ID
                    objControl.Parameter = rsBar!������
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
        rsBar.Sort = "���"
        Set objPopup = objBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "��չ����", lngTmp + 1, False)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.IconId = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        lngTmp = objPopup.Index
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���, lngTmp + 1)
                objControl.IconId = rsBar!ͼ��ID
                objControl.Parameter = rsBar!������
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                lngTmp = objPopup.Index
                rsBar.MoveNext
            Next
        End With
    End If
    
    '�Զ�ִ�еĹ���
    rsBar.Filter = "IsAuto=1"
    If Not rsBar.EOF Then mlngPlugInID = rsBar!����ID
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd��������_Click()
    Dim i As Integer
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    mintREPORTSEL = -1
    Call mobjFilter.GetWindowRect(lngLeft, lngTop, lngRight, lngBottom)
    For i = 0 To lst��������.ListCount - 1
        If txt��������.Tag = "" Then
            lst��������.Selected(i) = True
        ElseIf InStr("," & txt��������.Tag & ",", "," & lst��������.ItemData(i) & ",") > 0 Then
            lst��������.Selected(i) = True
        Else
            lst��������.Selected(i) = False
        End If
    Next
    lst��������.ListIndex = 0
    pic��������.Top = lngTop + IIf(mobjFilter.Position = 4, 350, 0) '��Ϊ�ƶ���������,��Ҫ���ϱ������ĸ߶�
    pic��������.Left = lngLeft + pic����ȼ�.Left
    pic��������.Width = txt��������.Width
    pic��������.Visible = True
    pic��������.ZOrder
    If lst��������.Visible And lst��������.Enabled Then lst��������.SetFocus
End Sub

Private Sub lst��������_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 Then
        For i = 1 To lst��������.ListCount - 1
            lst��������.Selected(i) = lst��������.Selected(0)
        Next
    ElseIf Not lst��������.Selected(Item) Then
        lst��������.Selected(0) = False
    ElseIf lst��������.SelCount = lst��������.ListCount - 1 Then
        lst��������.Selected(0) = True
    End If
End Sub

Private Sub lst��������_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst�������� _
        And Not Me.ActiveControl Is pic�������� Then pic��������.Visible = False
End Sub

Private Sub pic��������_GotFocus()
    If lst��������.Visible And lst��������.Enabled Then lst��������.SetFocus
End Sub

Private Sub pic��������_Resize()
    On Error Resume Next
    
    lst��������.Left = -15
    lst��������.Top = -15
    lst��������.Width = pic��������.Width
    
    cmdFilterCancel.Left = pic��������.ScaleWidth - cmdFilterCancel.Width - 100
    cmdFilterOK.Left = cmdFilterCancel.Left - cmdFilterOK.Width - 60
    
    cmdFilterOK.Top = lst��������.Height + (pic��������.ScaleHeight - lst��������.Height - cmdFilterOK.Height) / 2
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
    mbytFontSize = IIf(Val(zlDatabase.GetPara("��ʾ�����С", glngSys, 1265, 0, True)) = 0, 9, 12)
    mlngSource = IIf(mbytFontSize = 9, 999, 0)
    mintIndex = 0
    mblnRefresh = False
    mblnCardCollapse = False
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    mstrPrivs_����ɼ� = GetPrivFunc(glngSys, 1211)
    mintPatiInputType = 11
    '74410:���￨Ϊ������ʾ
    mblnShowCard = Not ISPassShowCard
    Me.Caption = "�°�סԺ��ʿ����վ"
    
    'Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    If gblnDo = True Then
        lbl����(mlngSource).Tag = IIf(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(cbsMain), cbsMain.Name & "ShowHomeNo", "0")) <> 0, 1, 0)
    Else
        lbl����(mlngSource).Tag = 0
    End If
        Call HaveRIS
    '��ʼ���˵�
    Call MainDefCommandBar
    cbsMain.RecalcLayout
    '�໤��
    mstrMonitor = ""
    mblnMonitor = Dir(App.Path & "\..\gdhs\AC2005.exe") <> ""
    If mblnMonitor Then mstrMonitor = App.Path & "\..\gdhs\AC2005.exe"
'    mblnMonitor = True '����ʱʹ��(zlWardMonitor�����Ѿ��ֹ��޸�Ϊ������)
    Call InitComponent
    
    mintOutPreTime = -1
    Call InitSelectTime
    Call GetLocalSetting '���ز���
    
    '��ȡ��������
    mstrSQL = " Select ����,��ɫ From ��������"
    Set mrsPatiColor = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ��������������Ϣ")
    
    mblnSupport = (Val(Split(GetVersion, ".")(1)) >= 32)
    If mblnSupport Then
        mstrBriefCode = ",zlpinyincode(NVL(B.����,a.����),0,0,',',1) AS ���� "
    Else
        mstrBriefCode = ",zlspellcode(NVL(B.����,a.����)) AS ����"
    End If
    
    '��ʼ�����˹�������
    strTmp = zlDatabase.GetPara("��ǰ��������", glngSys, pסԺ��ʿվ, "111", _
        Array(chk��������(0), chk��������(1), chk��������(2)), InStr(mstrPrivs, "��������") > 0)
    For i = 0 To chk��������.UBound
        chk��������(i).Value = IIf(Mid(strTmp, i + 1, 1) = "1", 1, 0)
    Next
    If Not InitBedlevel Then Unload Me: Exit Sub
    If Not InitNurselevel Then Unload Me: Exit Sub
    If Not InitUnits Then Unload Me: Exit Sub
    If cboUnit.ListIndex = -1 Then
        If InStr(mstrPrivs, "ȫԺ����") > 0 Then
            MsgBox "û�з���סԺ������Ϣ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Else
            MsgBox "û�з�������������,����ʹ��סԺ��ʿ����վ��", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If
    Call zlControl.CboSetWidth(cboUnit.hwnd, 3500)
    
    '������������
    Call RestoreWinState(Me, App.ProductName)
    
    '55928:������,2012-11-20,��ȡ��Ƭ�Ƿ��۵�
    If gblnDo = True Then
        mblnCardCollapse = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(cbsMain), cbsMain.Name & "Collapse", "0")) <> 0
        Call SetSourceCardH
    End If
    
    Call RaisEffect(picInfo, 2)
    mblnRefresh = True
    mblnStart = True
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1265, mstrPrivs, Me.hwnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    Call mclsAdvices.zlInitMip(mclsMipModule)
    
    Call Hook(Me)
    
    '����ҽ����������
    With frmNotify
        .mblnNormal = False
        .mintNotify = mintNotify
        .mintNotifyDay = mintNotifyDay
        .mstrNotifyAdvice = mstrNotifyAdvice
        .mdtOutBegin = mdtOutBegin
        .mdtOutEnd = mdtOutEnd
        .mstrScope = mstrScope
        .mlng����ID = cboUnit.ItemData(cboUnit.ListIndex)
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
    '�²��ؼ�
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
        fra���.Left = picList.Width - fra���.Width
        fra���.Top = PicDraw.Top + picList.Top + 50
    Else
        fra���.Left = stbThis.Width - fra���.Width - 1500
        fra���.Top = stbThis.Top + 50
    End If
    fraPatiUD.Visible = picList.Visible
    
    lblPatiInputType.Left = 120
    txtסԺ��.Left = lblPatiInputType.Left + lblPatiInputType.Width + 50
    pic��Ժ����.Top = picList.Top + 50
    pic��Ժ����.Left = 5000 + (TextWidth("��") - 180) * 15
    pic��Ժ����.Width = txtסԺ��.Left + txtסԺ��.Width
    pic��Ժ����.Height = txtסԺ��.Height + txtסԺ��.Top
    
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

    '��Ҫ�ر������Ӵ��壨��ģ̬��
    Unload frmNotify
    
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
        Set mfrmResponse = Nothing
    End If
    
    If Not mfrmNoticeBoard Is Nothing Then
        Unload mfrmNoticeBoard
        Set mfrmNoticeBoard = Nothing
    End If
    
    '54621:������,2013-02-28,��ʿվ�����ҳ������
    If Not mclsInOutMedRec Is Nothing Then
        Call mclsInOutMedRec.FormUnLoad
        Set mclsInOutMedRec = Nothing
    End If
    
    'ж����Ϣ����
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
    For i = 0 To chk��������.UBound
        strTmp = strTmp & IIf(chk��������(i).Value = 1, "1", "0")
    Next
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zlDatabase.SetPara("��ǰ��������", strTmp, glngSys, pסԺ��ʿվ, blnSetup)
    Call zlDatabase.SetPara("����ȼ�����", txt��������.Tag, glngSys, pסԺ��ʿվ, blnSetup)
    '���˷�Χ
    Dim curDate As Date
    curDate = zlDatabase.Currentdate
    '54436:������,2012-10-10,�޸���Ӧ������ģ���ΪpסԺ��ʿվ
    Call zlDatabase.SetPara("���ת������", Val(txtChange.Text), glngSys, pסԺ��ʿվ, blnSetup)
    Call zlDatabase.SetPara("��ʾ�����С", IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)), glngSys, mlngModul, blnSetup)

    '55928:������,2012-11-20,���ÿ�Ƭ�Ƿ��۵�
    If gblnDo = True Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(cbsMain), cbsMain.Name & "Collapse", IIf(mblnCardCollapse = True, 1, 0)
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(cbsMain), cbsMain.Name & "ShowHomeNo", Val(lbl����(mlngSource).Tag)
    End If

    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjPlugIn Is Nothing Then
        Call mobjPlugIn.Terminate(glngSys, P�°滤ʿվ, 1)
    End If
    
    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
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

Private Sub chk��������_Click(Index As Integer)
    Dim i As Integer, k As Integer
    '����ѡ��һ��
    For i = 0 To chk��������.UBound
        If chk��������(i).Value = 1 Then k = k + 1
    Next
    If k = 0 Then chk��������(Index).Value = 1
    
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
    If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
    pic��������.Visible = False
End Sub

Private Sub cmdFilterCancel_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst�������� _
        And Not Me.ActiveControl Is pic�������� Then pic��������.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim i As Integer
    
    If lst��������.SelCount = 0 Then
        MsgBox "������ѡ��һ�ֻ���ȼ���", vbInformation, gstrSysName
        If lst��������.Enabled And lst��������.Visible Then lst��������.SetFocus
    End If
    
    If lst��������.Selected(0) Then
        txt��������.Text = "ȫ��"
        txt��������.Tag = ""
    Else
        txt��������.Text = ""
        txt��������.Tag = ""
        For i = 1 To lst��������.ListCount - 1
            If lst��������.Selected(i) Then
                txt��������.Text = txt��������.Text & "," & lst��������.List(i)
                txt��������.Tag = txt��������.Tag & "," & lst��������.ItemData(i)
            End If
        Next
        txt��������.Text = Mid(txt��������.Text, 2)
        txt��������.Tag = Mid(txt��������.Tag, 2)
    End If
    
    If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
    pic��������.Visible = False
    
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

Private Function Get����ȼ�(ByVal str����ȼ� As String) As Integer
    '�������޵ȼ�ʱ,����3
    If InStr(1, str����ȼ�, "��") <> 0 Or InStr(1, str����ȼ�, "��") <> 0 Then
        Get����ȼ� = 0
    ElseIf InStr(1, str����ȼ�, "III") <> 0 Then
        Get����ȼ� = 3
    ElseIf InStr(1, str����ȼ�, "��") <> 0 Or InStr(1, str����ȼ�, "2") <> 0 Or InStr(1, str����ȼ�, "��") <> 0 Or InStr(1, str����ȼ�, "II") <> 0 Then
        Get����ȼ� = 2
    ElseIf InStr(1, str����ȼ�, "һ") <> 0 Or InStr(1, str����ȼ�, "1") <> 0 Or InStr(1, str����ȼ�, "��") <> 0 Or InStr(1, str����ȼ�, "I") <> 0 Then
        Get����ȼ� = 1
    Else
        Get����ȼ� = 3
    End If
End Function

Private Sub cmdFilterOK_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst�������� _
        And Not Me.ActiveControl Is pic�������� Then pic��������.Visible = False
End Sub

Private Sub picPati_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim blnValid As Boolean
    
    mintREPORTSEL = -1
    '��ʾ��Ƭѡ����
    If mlngSelect >= 0 Then
        '����Ҳһ��ȡ��ѡ��
        With mrsBedInfo
            .Filter = "��Ƭ����=" & mlngSelect
            If !����ID <> 0 Then
                If PicDraw.Enabled And PicDraw.Visible Then PicDraw.SetFocus
                .Filter = "����ID=" & !����ID
                Do While Not .EOF
                    '��ѡ��״̬���,ͬʱ����Ƭ��С��ԭ(�п������۵�ģʽ��)
                    picPati(!��Ƭ����).ZOrder 1
                    lblSelect(!��Ƭ����).Visible = False
                    If mblnCardCollapse Then
                        picPati(!��Ƭ����).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
                        picPati(!��Ƭ����).Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_�۵�, ��Ƭ����_��׼��Ƭ_�۵�)).Picture
                    End If
                    
                    .MoveNext
                Loop
            End If
            .Filter = 0
        End With
    End If
    
    mlngSelect = Index
    '53740:������,2012-09-19,��ִ�в���Զ�ִ�У��ڵ����˵�(��ǰ��ʽ�����޷�������ʾ�Ҽ��˵�)
    'mblnShow = True
    mblnShow = False: Call ShowSelect
    If Button = 2 Then
        Dim cbrPopupBar As CommandBar
        Dim cbrPopupItem As CommandBarControl
        Dim cbrMenuBar As CommandBarControl
        Dim cbrPopup As CommandBarPopup
        Dim cbrControl As Object
        Dim intIndex As Integer, int��Ƭ���� As Integer
        Dim str���Ա�ע As String, strKey As String, blnDeleteMunu As Boolean, strDeployKey As String
        Dim rsCopyNotes As New ADODB.Recordset
        
        If y < Me.lblSelect(Index).Top Then     '����ı�ע����
            '��ʾ�����б�ע���Ⲣ�ṩѡ��
            If mrsNotes.RecordCount = 0 Then Exit Sub
            If Not LocatePatiRecord Then Exit Sub
            mrsBedInfo.Filter = "����ID=" & mrsPatiInfo!����ID & " And ����=0"
            If mrsBedInfo.RecordCount = 0 Then
                mrsBedInfo.Filter = ""
                Exit Sub
            End If
            str���Ա�ע = mrsBedInfo!���Ա�ע1 & "'" & mrsBedInfo!���Ա�ע2 & "'" & mrsBedInfo!���Ա�ע3
            
            int��Ƭ���� = mrsBedInfo!��Ƭ����
            
            intIndex = 0
            If x > img���Ա��1(mlngSource).Left And x < img���Ա��1(mlngSource).Left + img���Ա��1(mlngSource).Width Then
                intIndex = 1
            ElseIf x > img���Ա��2(mlngSource).Left And x < img���Ա��2(mlngSource).Left + img���Ա��2(mlngSource).Width Then
                If mrsBedInfo!���Ա�ע1 = "" And mrsBedInfo!���Ա�ע2 = "" Then
                    intIndex = 1
                Else
                    intIndex = 2
                End If
            ElseIf x > img���Ա��3(mlngSource).Left And x < img���Ա��3(mlngSource).Left + img���Ա��3(mlngSource).Width Then
                If mrsBedInfo!���Ա�ע1 = "" And mrsBedInfo!���Ա�ע2 = "" And mrsBedInfo!���Ա�ע3 = "" Then
                    intIndex = 1
                ElseIf mrsBedInfo!���Ա�ע2 = "" And mrsBedInfo!���Ա�ע3 = "" Then
                    intIndex = 2
                Else
                    intIndex = 3
                End If
            Else
                If mrsBedInfo!���Ա�ע1 <> "" And mrsBedInfo!���Ա�ע2 <> "" And mrsBedInfo!���Ա�ע3 <> "" Then
                    Exit Sub
                ElseIf mrsBedInfo!���Ա�ע1 = "" Then
                    intIndex = 1
                ElseIf mrsBedInfo!���Ա�ע2 = "" Then
                    intIndex = 2
                Else
                    intIndex = 3
                End If
            End If
            '����Ҫ������ʾ��ͼ����ţ��ſ��Ѿ����ڵ���
            strDeployKey = ""
            If intIndex = 1 Then
                If mrsBedInfo!���Ա�ע2 <> "" Then
                    strDeployKey = Split(mrsBedInfo!���Ա�ע2, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע2, ",")(1)
                End If
                If mrsBedInfo!���Ա�ע3 <> "" Then
                    strDeployKey = IIf(strDeployKey = "", "", strDeployKey & "'") & Split(mrsBedInfo!���Ա�ע3, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע3, ",")(1)
                End If
            ElseIf intIndex = 2 Then
                If mrsBedInfo!���Ա�ע1 <> "" Then
                    strDeployKey = Split(mrsBedInfo!���Ա�ע1, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע1, ",")(1)
                End If
                If mrsBedInfo!���Ա�ע3 <> "" Then
                    strDeployKey = IIf(strDeployKey = "", "", strDeployKey & "'") & Split(mrsBedInfo!���Ա�ע3, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע3, ",")(1)
                End If
            Else
                If mrsBedInfo!���Ա�ע1 <> "" Then
                    strDeployKey = Split(mrsBedInfo!���Ա�ע1, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע1, ",")(1)
                End If
                If mrsBedInfo!���Ա�ע2 <> "" Then
                    strDeployKey = IIf(strDeployKey = "", "", strDeployKey & "'") & Split(mrsBedInfo!���Ա�ע2, ",")(0) & "," & Split(mrsBedInfo!���Ա�ע2, ",")(1)
                End If
            End If
            
            mrsBedInfo.Filter = ""
            If int��Ƭ���� <> Index Then Exit Sub
            
            Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
            cbrPopupBar.Title = "��ע�趨"
            If mlngSource = 999 Then
                Call cbrPopupBar.SetIconSize(16, 16)
            Else
                Call cbrPopupBar.SetIconSize(24, 24)
            End If
            
            mrsNotes.Filter = ""
            Set rsCopyNotes = zlDatabase.CopyNewRec(mrsNotes)
            mrsNotes.Filter = "������ = 0"
            mrsNotes.Sort = "����id,�������,������"
            Do While Not mrsNotes.EOF
                '�ſ���Ӧ������
                If InStr(1, "'" & strDeployKey & "'", "'" & mrsNotes!����ID & "," & mrsNotes!������� & "'") = 0 Then
                    Set cbrPopup = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_��ע1, mrsNotes!˵��)
                    If mlngSource = 999 Then
                        Call cbrPopup.CommandBar.SetIconSize(16, 16)
                    Else
                        Call cbrPopup.CommandBar.SetIconSize(24, 24)
                    End If
                    blnDeleteMunu = False
                    rsCopyNotes.Filter = "����id=" & mrsNotes!����ID & " And �������=" & mrsNotes!������� & " And ������>0"
                    Do While Not rsCopyNotes.EOF
                        strKey = rsCopyNotes!����ID & "," & rsCopyNotes!������� & "," & rsCopyNotes!������ & "," & rsCopyNotes!ͼ������ + 1
                        Set cbrPopupItem = cbrPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_��ע1 + rsCopyNotes.RecordCount, rsCopyNotes!˵��)
                        cbrPopupItem.IconId = conMenu_ͼ�� + (rsCopyNotes!ͼ������)
                        cbrPopupItem.Category = intIndex & "|" & rsCopyNotes!����ID & "|" & rsCopyNotes!������� & "|" & rsCopyNotes!������ & "|" & rsCopyNotes!ͼ������ + 1 & "|" & rsCopyNotes!˵��
                        If InStr(1, "'" & str���Ա�ע & "'", "'" & strKey & "'") <> 0 Then
                            cbrPopupItem.Checked = True
                            blnDeleteMunu = True
                        End If
                        rsCopyNotes.MoveNext
                    Loop
                    If blnDeleteMunu = True Then
                        Set cbrPopupItem = cbrPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_��ע1 + mrsNotes.RecordCount + 1, "�����ע")
                        cbrPopupItem.Category = intIndex & "|" & mrsNotes!����ID & "|" & mrsNotes!������� & "|0|0|"
                    End If
                End If
                mrsNotes.MoveNext
            Loop
            mrsNotes.Filter = 0
            cbrPopupBar.ShowPopup
        Else
            mrsBedInfo.Filter = "��Ƭ����=" & mlngSelect
            blnValid = (mrsBedInfo!����ID > 0)
            mrsBedInfo.Filter = 0
            
            If blnValid Then
                '��װ�Ҽ��˵�(���ù��ܲżӽ���)
                Set cbrMenuBar = mobjPopupBatch
                Set cbrPopupBar = cbsMain.Add("�Ҽ��˵�", xtpBarPopup)
                cbrPopupBar.Title = "�Ҽ��˵�"
                
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_TurnUnit, "ת����(&D)"): cbrPopupItem.Category = "����"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_TurnTeam, "תС��(&T)"):  cbrPopupItem.Category = "����"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_Turn, "ת��(&C)"): cbrPopupItem.Category = "����": cbrPopupItem.BeginGroup = True
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_Bed, "����(&B)"):  cbrPopupItem.Category = "����"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Change_Out, "��Ժ(&O)"):  cbrPopupItem.Category = "����"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_Manage_Change_Undo, "����(&U)"): cbrPopupItem.BeginGroup = True: cbrPopupItem.Category = "����"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ReStop, "ȷ��ֹͣ(&C)"): cbrPopupItem.BeginGroup = True: cbrPopupItem.Category = "ҽ��ҵ��"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_ReportLisView, "���������(&R)"): cbrPopupItem.Category = "ҽ��ҵ��"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Billing, "����(&C)"): cbrPopupItem.BeginGroup = True: cbrPopupItem.Category = "����ҵ��"
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ReBillingApply, "��������(&L)"): cbrPopupItem.Category = "����ҵ��"
                
                If Not mrsPlugInBar Is Nothing Then
                    mrsPlugInBar.Filter = "IsInTool=1 and BarType=3"
                    For intIndex = 1 To mrsPlugInBar.RecordCount
                        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mrsPlugInBar!����ID, mrsPlugInBar!������)
                            cbrPopupItem.IconId = mrsPlugInBar!ͼ��ID
                            cbrPopupItem.Parameter = mrsPlugInBar!������
                            If Val(mrsPlugInBar!IsGroup) = 1 Then cbrPopupItem.BeginGroup = True
                        mrsPlugInBar.MoveNext
                    Next
                    mrsPlugInBar.Filter = "IsInTool=0 and BarType=3"
                    If mrsPlugInBar.RecordCount > 0 Then
                        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugInPop, "��չ����"): cbrPopupItem.BeginGroup = True
                            cbrPopupItem.IconId = conMenu_Tool_PlugIn
                    End If
                    mrsPlugInBar.Filter = 0
                End If
                cbrPopupBar.ShowPopup
            End If
        End If
    Else
        '��������,���Ǽ��ģʽ
        If Button = 1 And mblnCardCollapse Then
'            If mblnShowCard = True Then
'                picPati(Index).Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
'                picPati(Index).Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_���￨, ��Ƭ����_��׼��Ƭ_���￨)).Picture
'            Else
'                picPati(Index).Height = IIf(mlngSource = 0, clngBigHeight_Normal, clngBaseHeight_Normal)
'                picPati(Index).Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ, ��Ƭ����_��׼��Ƭ)).Picture
'            End If
            picPati(Index).Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
            picPati(Index).Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_���￨, ��Ƭ����_��׼��Ƭ_���￨)).Picture
        End If
    End If
End Sub

Private Sub picPati_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And mblnCardCollapse Then
        picPati(Index).Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
        picPati(Index).Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_�۵�, ��Ƭ����_��׼��Ƭ_�۵�)).Picture
    End If
    
    picList.ZOrder 0
    PatiPage.ZOrder 0
    fraPatiUD.ZOrder 0
    picPara(2).ZOrder 0
    picPara(3).ZOrder 0
    pic��Ժ����.ZOrder 0
End Sub

'-------------------------------------------------------------------------------
'�����ǻ�������
'-------------------------------------------------------------------------------
Private Sub LoadPatiCard(ByVal intIndex As Integer, ByVal str���� As String, ByVal lngX As Long, ByVal lngY As Long, Optional ByVal blnVisible As Boolean = False)
    '���ش�λ��Ƭ
    '1����Ƭ�ϲ�
    '2����Ƭ����
    
    Load picPati(intIndex)
    With picPati(intIndex)
        .Left = lngX
        .Top = lngY
        .Width = picPati(mlngSource).Width
'        If mblnCardCollapse Then
'            .Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
'            .Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_�۵�, ��Ƭ����_��׼��Ƭ_�۵�)).Picture
'        ElseIf mblnShowCard = True Then
'            .Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
'            .Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_���￨, ��Ƭ����_��׼��Ƭ_���￨)).Picture
'        Else
'            .Height = IIf(mlngSource = 0, clngBigHeight_Normal, clngBaseHeight_Normal)
'            .Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ, ��Ƭ����_��׼��Ƭ)).Picture
'        End If
        If mblnCardCollapse Then
            .Height = IIf(mlngSource = 0, clngBigHeight_Collapse, clngBaseHeight_Collapse)
            .Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_�۵�, ��Ƭ����_��׼��Ƭ_�۵�)).Picture
        Else
            .Height = IIf(mlngSource = 0, clngBigCardHeight_Normal, clngBaseCardHeight_Normal)
            .Picture = img��Ƭ����(IIf(mlngSource = 0, ��Ƭ����_��Ƭ_���￨, ��Ƭ����_��׼��Ƭ_���￨)).Picture
        End If
        .Visible = blnVisible
        .ZOrder 0
    End With
    
    '1����Ƭ�ϲ�
    Load img����ȼ�(intIndex)
    img����ȼ�(intIndex).Visible = True
    Set img����ȼ�(intIndex).Container = picPati(intIndex)
    Set img����ȼ�(intIndex).Picture = Nothing
    img����ȼ�(intIndex).Left = img����ȼ�(mlngSource).Left
    img����ȼ�(intIndex).Top = img����ȼ�(mlngSource).Top
    img����ȼ�(intIndex).Height = img����ȼ�(mlngSource).Height
    img����ȼ�(intIndex).Width = img����ȼ�(mlngSource).Width
    img����ȼ�(intIndex).ZOrder 1
    
    Load lblSelect(intIndex)
    Set lblSelect(intIndex).Container = picPati(intIndex)
    lblSelect(intIndex).Left = lblSelect(mlngSource).Left
    lblSelect(intIndex).Top = lblSelect(mlngSource).Top
    lblSelect(intIndex).Width = lblSelect(mlngSource).Width
    
    Load lbl����(intIndex)
    Set lbl����(intIndex).Container = picPati(intIndex)
    lbl����(intIndex).Visible = True
    lbl����(intIndex).FontSize = lbl����(mlngSource).FontSize
    lbl����(intIndex).Left = lbl����(mlngSource).Left
    lbl����(intIndex).Top = lbl����(mlngSource).Top
    lbl����(intIndex).Width = lbl����(mlngSource).Width
    lbl����(intIndex).Height = lbl����(mlngSource).Height
    lbl����(intIndex).Caption = Mid(str����, 1, 7)
    
    Load lbl�����(intIndex)
    Set lbl�����(intIndex).Container = picPati(intIndex)
    lbl�����(intIndex).Caption = str����
    lbl�����(intIndex).Visible = False
    
    Load lbl����(intIndex)
    Set lbl����(intIndex).Container = picPati(intIndex)
    lbl����(intIndex).Visible = True
    lbl����(intIndex).FontSize = lbl����(mlngSource).FontSize
    lbl����(intIndex).Left = lbl����(mlngSource).Left
    lbl����(intIndex).Top = lbl����(mlngSource).Top
    lbl����(intIndex).Width = lbl����(mlngSource).Width
    lbl����(intIndex).Height = lbl����(mlngSource).Height
    lbl����(intIndex).Caption = ""
    
    Load lblSplit(intIndex)
    Set lblSplit(intIndex).Container = picPati(intIndex)
    lblSplit(intIndex).Visible = True
    lblSplit(intIndex).Left = lblSplit(mlngSource).Left
    lblSplit(intIndex).Top = lblSplit(mlngSource).Top
    lblSplit(intIndex).Width = lblSplit(mlngSource).Width
    lblSplit(intIndex).BackColor = &HFFFFFF
    
    Load img���Ա��2(intIndex)
    Set img���Ա��2(intIndex).Container = picPati(intIndex)
    img���Ա��2(intIndex).Picture = img���Ա��2(mlngSource).Picture
    img���Ա��2(intIndex).Stretch = img���Ա��2(mlngSource).Stretch
    img���Ա��2(intIndex).Top = img���Ա��2(mlngSource).Top
    img���Ա��2(intIndex).Left = img���Ա��2(mlngSource).Left
    img���Ա��2(intIndex).Width = img���Ա��2(mlngSource).Width
    img���Ա��2(intIndex).Height = img���Ա��2(mlngSource).Height
    
    Load img���Ա��3(intIndex)
    Set img���Ա��3(intIndex).Container = picPati(intIndex)
    img���Ա��3(intIndex).Picture = img���Ա��3(mlngSource).Picture
    img���Ա��3(intIndex).Stretch = img���Ա��3(mlngSource).Stretch
    img���Ա��3(intIndex).Top = img���Ա��3(mlngSource).Top
    img���Ա��3(intIndex).Left = img���Ա��3(mlngSource).Left
    img���Ա��3(intIndex).Width = img���Ա��3(mlngSource).Width
    img���Ա��3(intIndex).Height = img���Ա��3(mlngSource).Height
    
    Load img�ٴ�·��(intIndex)
    Set img�ٴ�·��(intIndex).Container = picPati(intIndex)
    img�ٴ�·��(intIndex).Picture = img�ٴ�·��(mlngSource).Picture
    img�ٴ�·��(intIndex).Stretch = img�ٴ�·��(mlngSource).Stretch
    img�ٴ�·��(intIndex).Top = img�ٴ�·��(mlngSource).Top
    img�ٴ�·��(intIndex).Left = img�ٴ�·��(mlngSource).Left
    img�ٴ�·��(intIndex).Width = img�ٴ�·��(mlngSource).Width
    img�ٴ�·��(intIndex).Height = img�ٴ�·��(mlngSource).Height
    
    Load img�������(intIndex)
    Set img�������(intIndex).Container = picPati(intIndex)
    img�������(intIndex).Picture = img�������(mlngSource).Picture
    img�������(intIndex).Stretch = img�������(mlngSource).Stretch
    img�������(intIndex).Top = img�������(mlngSource).Top
    img�������(intIndex).Left = img�������(mlngSource).Left
    img�������(intIndex).Width = img�������(mlngSource).Width
    img�������(intIndex).Height = img�������(mlngSource).Height
    
    Load img���Ա��1(intIndex)
    Set img���Ա��1(intIndex).Container = picPati(intIndex)
    img���Ա��1(intIndex).Picture = img���Ա��1(mlngSource).Picture
    img���Ա��1(intIndex).Stretch = img���Ա��1(mlngSource).Stretch
    img���Ա��1(intIndex).Top = img���Ա��1(mlngSource).Top
    img���Ա��1(intIndex).Left = img���Ա��1(mlngSource).Left
    img���Ա��1(intIndex).Width = img���Ա��1(mlngSource).Width
    img���Ա��1(intIndex).Height = img���Ա��1(mlngSource).Height
    
    Load img��Ժ(intIndex)
    Set img��Ժ(intIndex).Container = picPati(intIndex)
    img��Ժ(intIndex).Picture = img��Ժ(mlngSource).Picture
    img��Ժ(intIndex).Stretch = img��Ժ(mlngSource).Stretch
    img��Ժ(intIndex).Top = img��Ժ(mlngSource).Top
    img��Ժ(intIndex).Left = img��Ժ(mlngSource).Left
    img��Ժ(intIndex).Width = img��Ժ(mlngSource).Width
    img��Ժ(intIndex).Height = img��Ժ(mlngSource).Height
    
    '2����Ƭ����
    Load lblסԺ��(intIndex)
    Set lblסԺ��(intIndex).Container = picPati(intIndex)
    lblסԺ��(intIndex).Visible = True
    lblסԺ��(intIndex).FontSize = lblסԺ��(mlngSource).FontSize
    lblסԺ��(intIndex).Left = lblסԺ��(mlngSource).Left
    lblסԺ��(intIndex).Top = lblסԺ��(mlngSource).Top
    lblסԺ��(intIndex).Width = lblסԺ��(mlngSource).Width
    lblסԺ��(intIndex).Height = lblסԺ��(mlngSource).Height
    lblסԺ��(intIndex).Caption = ""
    
    Load lbl�Ա�(intIndex)
    Set lbl�Ա�(intIndex).Container = picPati(intIndex)
    lbl�Ա�(intIndex).Visible = True
    lbl�Ա�(intIndex).FontSize = lbl�Ա�(mlngSource).FontSize
    lbl�Ա�(intIndex).Left = lbl�Ա�(mlngSource).Left
    lbl�Ա�(intIndex).Top = lbl�Ա�(mlngSource).Top
    lbl�Ա�(intIndex).Width = lbl�Ա�(mlngSource).Width
    lbl�Ա�(intIndex).Height = lbl�Ա�(mlngSource).Height
    lbl�Ա�(intIndex).Caption = ""
    
    Load lbl����(intIndex)
    Set lbl����(intIndex).Container = picPati(intIndex)
    lbl����(intIndex).Visible = True
    lbl����(intIndex).FontSize = lbl����(mlngSource).FontSize
    lbl����(intIndex).Left = lbl����(mlngSource).Left
    lbl����(intIndex).Top = lbl����(mlngSource).Top
    lbl����(intIndex).Width = lbl����(mlngSource).Width
    lbl����(intIndex).Height = lbl����(mlngSource).Height
    lbl����(intIndex).Caption = ""
    
    Load lblҽʦ(intIndex)
    Set lblҽʦ(intIndex).Container = picPati(intIndex)
    lblҽʦ(intIndex).Visible = True
    lblҽʦ(intIndex).FontSize = lblҽʦ(mlngSource).FontSize
    lblҽʦ(intIndex).Left = lblҽʦ(mlngSource).Left
    lblҽʦ(intIndex).Top = lblҽʦ(mlngSource).Top
    lblҽʦ(intIndex).Width = lblҽʦ(mlngSource).Width
    lblҽʦ(intIndex).Height = lblҽʦ(mlngSource).Height
    lblҽʦ(intIndex).Caption = ""
    
    Load lbl��ʿ(intIndex)
    Set lbl��ʿ(intIndex).Container = picPati(intIndex)
    lbl��ʿ(intIndex).Visible = True
    lbl��ʿ(intIndex).FontSize = lbl��ʿ(mlngSource).FontSize
    lbl��ʿ(intIndex).Left = lbl��ʿ(mlngSource).Left
    lbl��ʿ(intIndex).Top = lbl��ʿ(mlngSource).Top
    lbl��ʿ(intIndex).Width = lbl��ʿ(mlngSource).Width
    lbl��ʿ(intIndex).Height = lbl��ʿ(mlngSource).Height
    lbl��ʿ(intIndex).Caption = ""
    
    Load lbl����(intIndex)
    Set lbl����(intIndex).Container = picPati(intIndex)
    lbl����(intIndex).Visible = True
    lbl����(intIndex).FontSize = lbl����(mlngSource).FontSize
    lbl����(intIndex).Left = lbl����(mlngSource).Left
    lbl����(intIndex).Top = lbl����(mlngSource).Top
    lbl����(intIndex).Width = lbl����(mlngSource).Width
    lbl����(intIndex).Height = lbl����(mlngSource).Height
    lbl����(intIndex).Caption = ""
    
    Load lbl��Ժ����(intIndex)
    Set lbl��Ժ����(intIndex).Container = picPati(intIndex)
    lbl��Ժ����(intIndex).Visible = True
    lbl��Ժ����(intIndex).FontSize = lbl��Ժ����(mlngSource).FontSize
    lbl��Ժ����(intIndex).Left = lbl��Ժ����(mlngSource).Left
    lbl��Ժ����(intIndex).Top = lbl��Ժ����(mlngSource).Top
    lbl��Ժ����(intIndex).Width = lbl��Ժ����(mlngSource).Width
    lbl��Ժ����(intIndex).Height = lbl��Ժ����(mlngSource).Height
    lbl��Ժ����(intIndex).Caption = ""
    
    Load lblסԺ����(intIndex)
    Set lblסԺ����(intIndex).Container = picPati(intIndex)
    lblסԺ����(intIndex).Visible = True
    lblסԺ����(intIndex).FontSize = lblסԺ����(mlngSource).FontSize
    lblסԺ����(intIndex).Left = lblסԺ����(mlngSource).Left
    lblסԺ����(intIndex).Top = lblסԺ����(mlngSource).Top
    lblסԺ����(intIndex).Width = lblסԺ����(mlngSource).Width
    lblסԺ����(intIndex).Height = lblסԺ����(mlngSource).Height
    lblסԺ����(intIndex).Caption = ""
    
    Load lbl���(intIndex)
    Set lbl���(intIndex).Container = picPati(intIndex)
    lbl���(intIndex).FontSize = lbl���(mlngSource).FontSize
    lbl���(intIndex).Visible = True
    lbl���(intIndex).Left = lbl���(mlngSource).Left
    lbl���(intIndex).Top = lbl���(mlngSource).Top
    lbl���(intIndex).Width = lbl���(mlngSource).Width
    lbl���(intIndex).Height = lbl���(mlngSource).Height
    lbl���(intIndex).Caption = ""
    
    '61824:������,2013-05-23,��ʾ�����ֱ�־
    Load img������(intIndex)
    Set img������(intIndex).Container = picPati(intIndex)
    img������(intIndex).Picture = img������(mlngSource).Picture
    img������(intIndex).Stretch = img������(mlngSource).Stretch
    img������(intIndex).Top = img������(mlngSource).Top
    img������(intIndex).Left = img������(mlngSource).Left
    img������(intIndex).Width = img������(mlngSource).Width
    img������(intIndex).Height = img������(mlngSource).Height
    
    Load lbl����(intIndex)
    Set lbl����(intIndex).Container = picPati(intIndex)
    lbl����(intIndex).Visible = True
    lbl����(intIndex).FontSize = lbl����(mlngSource).FontSize
    lbl����(intIndex).Left = lbl����(mlngSource).Left
    lbl����(intIndex).Top = lbl����(mlngSource).Top
    lbl����(intIndex).Width = lbl����(mlngSource).Width
    lbl����(intIndex).Height = lbl����(mlngSource).Height
    lbl����(intIndex).Caption = ""
    
    Load lbl�����ܶ�(intIndex)
    Set lbl�����ܶ�(intIndex).Container = picPati(intIndex)
    lbl�����ܶ�(intIndex).Visible = True
    lbl�����ܶ�(intIndex).FontSize = lbl�����ܶ�(mlngSource).FontSize
    lbl�����ܶ�(intIndex).Left = lbl�����ܶ�(mlngSource).Left
    lbl�����ܶ�(intIndex).Top = lbl�����ܶ�(mlngSource).Top
    lbl�����ܶ�(intIndex).Width = lbl�����ܶ�(mlngSource).Width
    lbl�����ܶ�(intIndex).Height = lbl�����ܶ�(mlngSource).Height
    lbl�����ܶ�(intIndex).Caption = ""
    
    '74410:��Ƭ��ʾ���￨��
    Load lblCardNo(intIndex)
    Set lblCardNo(intIndex).Container = picPati(intIndex)
    lblCardNo(intIndex).Visible = mblnShowCard
    lblCardNo(intIndex).FontSize = lblCardNo(mlngSource).FontSize
    lblCardNo(intIndex).Left = lblCardNo(mlngSource).Left
    lblCardNo(intIndex).Top = lblCardNo(mlngSource).Top
    lblCardNo(intIndex).Width = lblCardNo(mlngSource).Width
    lblCardNo(intIndex).Height = lblCardNo(mlngSource).Height
    lblCardNo(intIndex).Caption = ""
    
    '66618:��ʾҽ�Ƹ��ʽ
    Load lblMedPay(intIndex)
    Set lblMedPay(intIndex).Container = picPati(intIndex)
    lblMedPay(intIndex).Visible = True
    lblMedPay(intIndex).FontSize = lblMedPay(mlngSource).FontSize
    lblMedPay(intIndex).Left = lblMedPay(mlngSource).Left
    lblMedPay(intIndex).Top = lblMedPay(mlngSource).Top
    lblMedPay(intIndex).Width = IIf(mblnShowCard = True, lblMedPay(mlngSource).Width, lblҽʦ(mlngSource).Width)
    lblMedPay(intIndex).Height = lblMedPay(mlngSource).Height
    lblMedPay(intIndex).Caption = ""
    
'    If mblnShowCard = False Then
'        lbl����(intIndex).Top = lbl��Ժ����(intIndex).Top
'        lbl�����ܶ�(intIndex).Top = lbl����(intIndex).Top
'        lbl��Ժ����(intIndex).Top = lblCardNo(intIndex).Top
'        lblסԺ����(intIndex).Top = lbl��Ժ����(intIndex).Top
'    End If
End Sub

Private Sub SetCardInfo(ByVal intIndex As Integer, ByVal ArrPatiInfo As Variant)
    Dim imgManIcon As ImageManagerIcons
    Dim int����ȼ� As Integer
    
    'סԺ��,����,�Ա�,����,���,ҽ/��,�ѱ�,ҽ�Ƹ��ʽ,����,��Ժ����,סԺ����,���,������ɫ,����ȼ�,���￨��
    lblסԺ��(intIndex).Caption = CStr(ArrPatiInfo(0))
    lbl����(intIndex).Caption = CStr(ArrPatiInfo(1))
    lbl����(intIndex).Alignment = 1
    lbl�Ա�(intIndex).Caption = CStr(ArrPatiInfo(2))
    If lbl�Ա�(intIndex).Caption = "����" Then lbl�Ա�(intIndex).Visible = False
    lbl����(intIndex).Caption = CStr(ArrPatiInfo(3))
    If IsNumeric(lbl����(intIndex).Caption) Then lbl����(intIndex) = lbl����(intIndex) & "��"
    lblҽʦ(intIndex).Caption = "ҽ��:" & CStr(ArrPatiInfo(5))
    lbl��ʿ(intIndex).Caption = "�ѱ�:" & CStr(ArrPatiInfo(6))
    lblMedPay(intIndex).Caption = CStr(ArrPatiInfo(7))
    lblCardNo(intIndex).Caption = CStr(ArrPatiInfo(14))
    lbl����(intIndex).Caption = CStr(ArrPatiInfo(8))
    lbl��Ժ����(intIndex).Caption = CStr(ArrPatiInfo(9))
    lblסԺ����(intIndex).Caption = IIf(Val(ArrPatiInfo(10) & "") = 0, 1, ArrPatiInfo(10)) & "��"
    lbl���(intIndex).Caption = CStr(ArrPatiInfo(4))
    lbl�����ܶ�(intIndex).Caption = CStr(ArrPatiInfo(11))
    lblSplit(intIndex).BackColor = Val(CStr(ArrPatiInfo(12)))
    
    '���û���ȼ�(�ؼ���,һ����,������,������)
    int����ȼ� = Get����ȼ�(CStr(ArrPatiInfo(13)))
    Set img����ȼ�(intIndex).Picture = imgHLDJ(mlngSource).ListImages(int����ȼ� + 1).Picture
    
    If lbl�����ܶ�(intIndex).Caption <> "" Then
        If lbl�����ܶ�(intIndex).Caption = "���޶�ȵ���" Then
            lbl�����ܶ�(intIndex).Caption = ""
            lbl����(intIndex).Caption = "���޶�ȵ���"
            lbl����(intIndex).ForeColor = &HFF0000
            lbl����(intIndex).ZOrder 0
        Else
            If Val(lbl�����ܶ�(intIndex).Caption) < 0 Then
                lbl����(intIndex).Caption = "Ƿ��"
                lbl����(intIndex).ForeColor = &HFF&
                lbl�����ܶ�(intIndex).ForeColor = &HFF&
            Else
                lbl����(intIndex).Caption = "���"
            End If
        End If
    Else
        lbl����(intIndex) = ""
        lbl�����ܶ�(intIndex).Caption = ""
        lblҽʦ(intIndex).Caption = ""
        lbl��ʿ(intIndex).Caption = ""
        lblMedPay(intIndex).Caption = ""
        lblCardNo(intIndex).Caption = ""
        lblסԺ����(intIndex).Caption = ""
        Set img���Ա��2(intIndex).Picture = Nothing
        Set img�������(intIndex).Picture = Nothing
        Set img�ٴ�·��(intIndex).Picture = Nothing
        Set img���Ա��3(intIndex).Picture = Nothing
        Set img���Ա��1(intIndex).Picture = Nothing
        Set img��Ժ(intIndex).Picture = Nothing
        Set img����ȼ�(intIndex).Picture = Nothing
        Set img������(intIndex).Picture = Nothing
    End If
    
    If mblnShowCard = True Then
        If Trim(lblCardNo(intIndex).Caption) = "" Then
            lblMedPay(intIndex).Width = lblҽʦ(mlngSource).Width
        Else
            lblMedPay(intIndex).Width = lblMedPay(mlngSource).Width
        End If
    End If
End Sub

Private Sub SetCardLabel(ByVal intIndex As Integer)
    Dim intTar As Integer
    Dim intSignIndex As Integer
    On Error GoTo ErrHand
    
    '���ÿ�Ƭ��ע����
    mrsBedInfo.Filter = "��Ƭ����=" & intIndex
    If mrsBedInfo.RecordCount <> 0 Then
        If mrsBedInfo!������� <> 0 Then
            Set img�������(intIndex).Picture = Img���(mlngSource).ListImages(Get����ͼ�����(mrsBedInfo!�������)).Picture
        End If
        img�������(intIndex).Visible = mrsBedInfo!�������
        img�������(intIndex).Tag = "" & mrsBedInfo!�����������
        
        If mrsBedInfo!�ٴ�·�� <> 0 Then
            Set img�ٴ�·��(intIndex).Picture = Img���(mlngSource).ListImages(Get�ٴ�·�����(mrsBedInfo!�ٴ�·��)).Picture
        End If
        img�ٴ�·��(intIndex).Visible = mrsBedInfo!�ٴ�·��
        img�ٴ�·��(intIndex).Tag = "" & mrsBedInfo!�ٴ�·������
        img�ٴ�·��(intIndex).Visible = mblnHavePath
        
        intSignIndex = 0
        If mrsBedInfo!���Ա�ע1 <> "" Then
            intSignIndex = Split(mrsBedInfo!���Ա�ע1, ",")(3)
            If intSignIndex > 0 And intSignIndex <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
                Set img���Ա��1(intIndex).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(intSignIndex).Picture
            Else
                intSignIndex = 0
            End If
        End If
        img���Ա��1(intIndex).Visible = intSignIndex > 0
        img���Ա��1(intIndex).Tag = "" & mrsBedInfo!���Ա�ע1����
        
        If mrsBedInfo!����״̬ <> 0 Then
            Set img��Ժ(intIndex).Picture = Img���(mlngSource).ListImages(CLng(mrsBedInfo!����״̬)).Picture
        End If
        img��Ժ(intIndex).Visible = mrsBedInfo!����״̬
        img��Ժ(intIndex).Tag = "" & mrsBedInfo!����״̬����
        
        intSignIndex = 0
        If mrsBedInfo!���Ա�ע2 <> "" Then
            intSignIndex = Split(mrsBedInfo!���Ա�ע2, ",")(3)
            If intSignIndex > 0 And intSignIndex <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
                Set img���Ա��2(intIndex).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(intSignIndex).Picture
            Else
                intSignIndex = 0
            End If
        End If
        img���Ա��2(intIndex).Visible = intSignIndex > 0
        img���Ա��2(intIndex).Tag = "" & mrsBedInfo!���Ա�ע2����
        
        intSignIndex = 0
        If mrsBedInfo!���Ա�ע3 <> "" Then
            intSignIndex = Split(mrsBedInfo!���Ա�ע3, ",")(3)
            If intSignIndex > 0 And intSignIndex <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
                Set img���Ա��3(intIndex).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(intSignIndex).Picture
            Else
                intSignIndex = 0
            End If
        End If
        img���Ա��3(intIndex).Visible = intSignIndex > 0
        img���Ա��3(intIndex).Tag = "" & mrsBedInfo!���Ա�ע3����
        
        '61824:������,2013-05-23,��ʾ�����ֱ�־
        If Nvl(mrsBedInfo!������) <> "" Then
            Set img������(intIndex).Picture = Img���(mlngSource).ListImages("������").Picture
        End If
        img������(intIndex).Visible = Nvl(mrsBedInfo!������) <> ""
        img������(intIndex).Tag = Nvl(mrsBedInfo!������)
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

    strOut = "ɾ���ؼ���ʼʱ��: " & Now
    For j = picPati.Count - 2 To 1 Step -1
        Unload lblSplit(j)
        Unload lblSelect(j)
        Unload lbl����(j)
        Unload lbl�����(j)
        Unload lblסԺ��(j)
        Unload lbl����(j)
        Unload lbl�Ա�(j)
        Unload lbl����(j)
        Unload lblҽʦ(j)
        Unload lbl��ʿ(j)
        Unload lbl����(j)
        Unload lbl��Ժ����(j)
        Unload lblסԺ����(j)
        Unload lbl���(j)
        Unload lbl����(j)
        Unload lbl�����ܶ�(j)
        Unload lblCardNo(j)
        Unload lblMedPay(j)

        Unload img���Ա��2(j)
        Unload img���Ա��3(j)
        Unload img�ٴ�·��(j)
        Unload img�������(j)
        Unload img���Ա��1(j)
        Unload img��Ժ(j)
        '61824:������,2013-05-23,��ʾ�����ֱ�־
        Unload img������(j)
        Unload img����ȼ�(j)
        Unload picPati(j)
    Next
    strOut = strOut & vbCrLf & "ɾ���ؼ���ʼʱ��: " & Now
    'MsgBox strOut
End Sub

Private Sub timeRefreshCard_Timer()
    Dim lngIndex As Long
    '���ѡ����ĳ����Ŀ,������˸����
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
    
    'ˢ�¹�����
    If Not mfrmNoticeBoard Is Nothing And cboUnit.ListIndex <> -1 Then
        If mfrmNoticeBoard.mblnShow = True Then Call mfrmNoticeBoard.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex))
    End If
End Sub

Private Sub ShowSelect()
    Dim rsTmp As New ADODB.Recordset
    '��ʾ��ǰѡ�����Ŀ
    
    If mlngSelect < 0 Then Exit Sub
    '����Ҳһ��ѡ��
    With mrsBedInfo
        .Filter = "��Ƭ����=" & mlngSelect
        If !����ID <> 0 Then
            mlng����ID = !����ID
            mlng��ҳID = !��ҳID
            
            .Filter = "����ID=" & !����ID
            Do While Not .EOF
                lblSelect(!��Ƭ����).Visible = True
                lblSelect(!��Ƭ����).ZOrder 1
                img����ȼ�(!��Ƭ����).ZOrder 1
                .MoveNext
            Loop
        Else
            mlng����ID = 0
            mlng��ҳID = 0
        End If
        .Filter = 0
    End With

    picPati(mlngSelect).ZOrder 0
    If picPati(mlngSelect).Visible And picPati(mlngSelect).Enabled Then picPati(mlngSelect).SetFocus
    
    Call GetPatiOtherInfo
End Sub

Private Sub GetPatiOtherInfo()
    '�������ڴ����˻��Ƿ��ڴ�����,������ȡ��סԺ��Ϣ,�ڰ�ť״̬�仯ʱ��ʹ��
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '������Ϣȡ��ǰסԺ������
    If Not LocatePatiRecord Then Exit Sub
    
    mPatiInfo.���� = CStr(mrsPatiInfo!����)
    mPatiInfo.����״̬ = Nvl(mrsPatiInfo!����״̬, 0)
    mPatiInfo.·��״̬ = mrsPatiInfo!·��״̬
    
    'ȡ������Ϣ
    mstrSQL = "Select B.��ҳID,B.״̬,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����,b.��Ժ����,B.סԺ��,b.��Ժ����,B.��������,B.����ת��,B.����,b.��ǰ����id,B.��Ժ����ID,B.��ǰ����ID,Decode(Nvl(X.�������, 0), 0, '��', '') As ����" & _
        " From ������ҳ B,������� X" & _
        " Where B.����ID=[1] And B.��ҳID=[2] And B.����ID = X.����ID(+) And X.����(+) = 1 And X.����(+)=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(mrsPatiInfo.Fields("����ID").Value), Val(mrsPatiInfo.Fields("��ҳID").Value))
    With rsTmp
        mPatiInfo.״̬ = Nvl(!״̬, 0)
        mPatiInfo.��ҳID = Nvl(!��ҳID, 0)
        mPatiInfo.סԺ�� = Nvl(!סԺ��)
        mPatiInfo.���� = Nvl(!��Ժ����)
        mPatiInfo.����ID = Nvl(!��ǰ����ID, 0)
        mPatiInfo.����ID = Nvl(!��Ժ����ID, 0)
        mPatiInfo.��Ժ���� = !��Ժ����
        If Not IsNull(!��Ժ����) Then
            mPatiInfo.��Ժ���� = !��Ժ����
        Else
            mPatiInfo.��Ժ���� = CDate(0)
        End If
        mPatiInfo.���� = Val("" & !����)
        mPatiInfo.���� = Not IsNull(!����)
        mPatiInfo.���� = Nvl(!��������, 0)
        mPatiInfo.���� = is����(Val(!��Ժ����ID & ""))
        mPatiInfo.����ת�� = Nvl(!����ת��, 0) <> 0
    End With
    '53740:������,2012-09-19,�л������Զ�ִ����ҹ���
    Call AutoExecutePlugIn(cbsMain)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AutoExecutePlugIn(ByVal cbsMain As Object)
    Dim objControl As CommandBarControl
    Dim lng����ID As Long, lng��ҳID As Long
    
    If mrsPatiInfo.RecordCount = 0 Then
        lng����ID = 0
        lng��ҳID = 0
    Else
        lng����ID = Val(mrsPatiInfo.Fields("����ID").Value)
        lng��ҳID = Val(mrsPatiInfo.Fields("����ID").Value)
    End If
    'ִ���Զ��������
    If mlngPlugInID <> 0 And (mlngPre����ID <> lng����ID Or (mlngPre����ID = lng����ID And mlngPre��ҳID <> lng��ҳID)) Then
        mlngPre����ID = lng����ID: mlngPre��ҳID = lng��ҳID
        Set objControl = cbsMain.FindControl(, mlngPlugInID, , True)
        If Not objControl Is Nothing Then objControl.Execute
    End If
End Sub

Private Sub GetInpatientAreaInfo()
    Dim strAdvance As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    '���������ڱ�ע��������ʱ��¼,�ڽ�����ز���ʱ����,ˢ�µ�ʱ��Ŵ����ݿ����ȡ
    '53907:������,2012-09-28
'    mstrSQL = "" & _
'            " SELECT SUM(��Ժ) AS ��Ժ,SUM(���) AS ���,SUM(ת��) AS ת��,SUM(����) AS ����,SUM(��Ժ) AS ��Ժ,SUM(Σ) AS Σ,SUM(��) AS ��" & _
'            " FROM (" & _
'            "     SELECT SUM(DECODE(��ʼԭ��,2,1,0)) AS ��Ժ,SUM(DECODE(��ʼԭ��,3,1,0)) AS ���,0 AS ת��,0 AS ����,0 AS ��Ժ,0 AS Σ,0 AS ��" & _
'            "     From ���˱䶯��¼" & _
'            "     Where ����ID = [1]" & _
'            "     AND ��ʼʱ�� BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS ��Ժ,0 AS ���,SUM(DECODE(��ֹԭ��,3,1,0)) AS ת��,0 AS ����,0 AS ��Ժ,0 AS Σ,0 AS ��" & _
'            "     From ���˱䶯��¼" & _
'            "     Where ����ID = [1]" & _
'            "     AND ��ֹʱ�� BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS ��Ժ,0 AS ���,0 AS ת��,SUM(DECODE(��Ժ��ʽ,'����',1,0)) AS ����,SUM(DECODE(��Ժ��ʽ,'����',0,1)) AS ��Ժ,0 AS Σ,0 AS ��" & _
'            "     From ������ҳ A,������Ϣ B" & _
'            "     Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.��Ժ=1 And A.��ǰ����ID = [1]" & _
'            "     AND ��Ժ���� BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS ��Ժ,0 AS ���,0 AS ת��,0 AS ����,0 AS ��Ժ,SUM(DECODE(��ǰ����,'Σ',1,0)) AS Σ,SUM(DECODE(��ǰ����,'��',1,0)) AS ��" & _
'            "     From ������ҳ A,������Ϣ B" & _
'            "     Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.��Ժ=1 And A.��ǰ����ID = [1]" & _
'            "     AND ��Ժ���� IS NULL" & _
'            ")"
    mstrSQL = "" & _
            " SELECT SUM(��Ժ) AS ��Ժ,SUM(���) AS ���,SUM(ת��) AS ת��,SUM(����) AS ����,SUM(��Ժ) AS ��Ժ,SUM(Σ) AS Σ,SUM(��) AS ��" & _
            " FROM (" & _
            "     SELECT SUM(DECODE(��ʼԭ��,2,1,0)) AS ��Ժ,SUM(DECODE(��ʼԭ��,3,1,15,1,0)) AS ���,0 AS ת��,0 AS ����,0 AS ��Ժ,0 AS Σ,0 AS ��" & _
            "     From ���˱䶯��¼" & _
            "     Where ����ID = [1] And NVL(���Ӵ�λ,0)=0" & _
            "     AND ��ʼʱ�� BETWEEN [2] AND SYSDATE" & _
            "     Union" & _
            "     Select SUM(1) as ��Ժ,0 AS ���,0 AS ת��,0 AS ����,0 AS ��Ժ,0 AS Σ,0 AS ��" & _
            "     From ���˱䶯��¼ a, ������ҳ b" & _
            "     Where a.����id = b.����id And a.��ҳid = b.��ҳid And A.����ID=[1] And A.��ʼʱ�� Between [2] And Sysdate And a.��ʼԭ�� = 1 And Nvl(a.���Ӵ�λ, 0) = 0 And" & _
            "       Nvl(b.״̬, 0) <> 1 And Not Exists" & _
            "       (Select 1 From ���˱䶯��¼ Where ����id = a.����id And ��ҳid = b.��ҳid And ��ʼԭ�� = 2)"
    mstrSQL = mstrSQL & _
            "     Union" & _
            "     SELECT 0 AS ��Ժ,0 AS ���,SUM(DECODE(��ֹԭ��,3,1,15,1,0)) AS ת��,0 AS ����,0 AS ��Ժ,0 AS Σ,0 AS ��" & _
            "     From ���˱䶯��¼" & _
            "     Where ����ID = [1] And NVL(���Ӵ�λ,0)=0" & _
            "     AND ��ֹʱ�� BETWEEN [2] AND SYSDATE" & _
            "     Union" & _
            "     SELECT 0 AS ��Ժ,0 AS ���,0 AS ת��,SUM(DECODE(��Ժ��ʽ,'����',1,0)) AS ����,SUM(DECODE(��Ժ��ʽ,'����',0,1)) AS ��Ժ,0 AS Σ,0 AS ��" & _
            "     From ������ҳ A,������Ϣ B" & _
            "     Where A.����ID=B.����ID  And A.��ǰ����ID = [1]" & _
            "     AND ��Ժ���� BETWEEN [2] AND SYSDATE" & _
            "     Union" & _
            "     SELECT 0 AS ��Ժ,0 AS ���,0 AS ת��,0 AS ����,0 AS ��Ժ,SUM(DECODE(��ǰ����,'Σ',1,0)) AS Σ,SUM(DECODE(��ǰ����,'��',1,0)) AS ��" & _
            "     From ������ҳ A,������Ϣ B,��Ժ���� C" & _
            "     Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And NVL(A.״̬,0)<>1 And Nvl(A.����״̬,0)<>5 And A.���ʱ�� is NULL And B.����ID=C.����ID " & _
            "       And B.��ǰ����ID=C.����ID And C.����ID=[1]" & _
            ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ����������Ϣ", cboUnit.ItemData(cboUnit.ListIndex), CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")))
    mlng��Ժ = Nvl(rsTemp!��Ժ, 0)
    mlngת�� = Nvl(rsTemp!���, 0)
    mlng��Ժ = Nvl(rsTemp!��Ժ, 0)
    mlngת�� = Nvl(rsTemp!ת��, 0)
    mlng���� = Nvl(rsTemp!����, 0)
    mlngΣ = Nvl(rsTemp!Σ, 0)
    mlng�� = Nvl(rsTemp!��, 0)
    
    'LPF,2014-10-21,�����Ż�:�����Ժ���˱�
'    mstrSQL = "" & _
'        " Select B.ID,B.����,count(*) AS ����" & vbNewLine & _
'        " From ������ҳ A,�շ���ĿĿ¼ B" & vbNewLine & _
'        " Where A.����ȼ�ID=B.ID And A.��Ժ���� IS Null And NVL(A.״̬,0)<>1 And Nvl(A.����״̬,0)<>5 And A.���ʱ�� is NULL And A.��ǰ����ID=[1]" & vbNewLine & _
'        " Group by B.ID,B.����"
    mstrSQL = "" & _
        " Select b.Id, b.����, Count(*) As ����" & vbNewLine & _
        " From �շ���ĿĿ¼ b, ������Ϣ c, ������ҳ a, ��Ժ���� e" & vbNewLine & _
        " Where b.Id = a.����ȼ�id And a.��Ժ���� Is Null And Nvl(a.״̬, 0) <> 1 And Nvl(a.����״̬, 0) <> 5 And a.���ʱ�� Is Null And" & vbNewLine & _
        "      c.����id = a.����id And c.��ҳid = a.��ҳid And c.����id = e.����id And c.��ǰ����id = e.����id And e.����id = [1]" & vbNewLine & _
        " Group By b.Id, b.����"

    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ����������Ϣ", cboUnit.ItemData(cboUnit.ListIndex))
    Do While Not rsTemp.EOF
        strAdvance = strAdvance & "��" & rsTemp!���� & "��" & rsTemp!���� & "��"
        rsTemp.MoveNext
    Loop
    If strAdvance <> "" Then
        strAdvance = Mid(strAdvance, 2)
        strAdvance = "��" & strAdvance
    End If
    
    Call ShowInpatientAreaInfo(strAdvance)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowGuage(ByVal strInfo As String, ByVal dblPer As Double)
    Dim dblLength As Double     '�������ĵ�ǰ���
    
    picInfo.Height = 315
    picInfo.BorderStyle = 1
    
    '��ʾ������
    lblInpatientArea.Top = 60
    lblInpatientArea.AutoSize = False
    lblInpatientArea.Width = 3000
    lblInpatientArea.Caption = strInfo
    
    dblLength = picInfo.Width - lblInpatientArea.Width - 50
    '��ͼ
    picInfo.Cls
    On Error Resume Next
    If Format(dblPer, "#0.00;-#0.00;0") <> "0" Then
        picInfo.PaintPicture picSource.Picture, lblInpatientArea.Width, 0, dblLength * dblPer / 100
    End If
    If err <> 0 Then err.Clear
    picInfo.Refresh
End Sub

Private Sub ShowInpatientAreaInfo(Optional ByVal strAdvance As String = "")
    Dim lng��Ժ���� As Long, lng�ܴ�λ As Long
    '10�ſմ�������52�ˣ���Ժ_�ˣ�ת��4�ˣ�ת����3�ˣ���Ժ5�ˣ�ת��_�ˣ�����_�ˣ�Σ/�أ�1/_������5��
    
    '71923
    mrsBedInfo.Filter = "����=0"
    lng��Ժ���� = mrsBedInfo.RecordCount + mlng�Ҵ� '- mlngԤ��Ժ
    mrsBedInfo.Filter = "����ID=0"
    mlng�մ� = mrsBedInfo.RecordCount
    mrsBedInfo.Filter = 0
    lng�ܴ�λ = mrsBedInfo.RecordCount
    mlng�ڴ� = lng�ܴ�λ - mlng�մ�
    
    picInfo.Cls
    picInfo.Height = 345
    
    lblInpatientArea.Top = 80
    lblInpatientArea.AutoSize = True
    lblInpatientArea.Caption = cboUnit.Text & "�����������������" & lng�ܴ�λ & "�ſ��ô�λ��" & mlng�մ� & "�ſմ�������" & lng��Ժ���� & "��(���мҴ���" & mlng�Ҵ� & "��)��Σ/�أ�" & mlngΣ & "/" & mlng�� & strAdvance
    lblInpatientArea.Caption = lblInpatientArea.Caption & "���������������Ժ" & mlng��Ժ & "�ˣ�ת��" & mlngת�� & "�ˣ���Ժ" & mlng��Ժ & "�ˣ�ת��" & mlngת�� & "�ˣ�����" & mlng���� & "��"
    
    Call RaisEffect(picInfo, 2)
End Sub

Private Sub Set������Ŀ��������()
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "���ƻ�������(ZLCISBase)û����ȷ��װ���ù����޷�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallSetClinicCharge(Me.cboUnit.ItemData(Me.cboUnit.ListIndex), 1, Me, gcnOracle, glngSys, gstrDBUser, EסԺ����, InStr(GetInsidePrivs(mlngModul), ";������Ŀ��������;") = 0)
End Sub


'-------------------------------------------------------------------------------
'���´��������
'-------------------------------------------------------------------------------

Private Sub img�������_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img�������(Index).Tag, True
End Sub

Private Sub img��Ժ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img��Ժ(Index).Tag, True
End Sub

Private Sub img���Ա��1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img���Ա��1(Index).Tag, True
End Sub

Private Sub img���Ա��2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img���Ա��2(Index).Tag, True
End Sub

Private Sub img���Ա��3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img���Ա��3(Index).Tag, True
End Sub

Private Sub img�ٴ�·��_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    zlCommFun.ShowTipInfo picPati(Index).hwnd, img�ٴ�·��(Index).Tag, True
End Sub

Private Sub img�������_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img��Ժ_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img���Ա��1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img���Ա��2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img���Ա��3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img�ٴ�·��_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img�������_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img�������(Index).Left + x, img�������(Index).Top + y)
End Sub

Private Sub img��Ժ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img��Ժ(Index).Left + x, img��Ժ(Index).Top + y)
End Sub

Private Sub img���Ա��2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img���Ա��2(Index).Left + x, img���Ա��2(Index).Top + y)
End Sub

Private Sub img�ٴ�·��_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img�ٴ�·��(Index).Left + x, img�ٴ�·��(Index).Top + y)
End Sub

Private Sub img���Ա��1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img���Ա��1(Index).Left + x, img���Ա��1(Index).Top + y)
End Sub

Private Sub img���Ա��3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img���Ա��3(Index).Left + x, img���Ա��3(Index).Top + y)
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

Private Sub lbl����_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl����(Index).Left + x, lbl����(Index).Top + y)
End Sub

Private Sub lbl����_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl��ʿ_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl��ʿ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl��ʿ(Index).Left + x, lbl��ʿ(Index).Top + y)
End Sub

Private Sub lbl��ʿ_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub img����ȼ�_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub img����ȼ�_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, img����ȼ�(Index).Left + x, img����ȼ�(Index).Top + y)
End Sub

Private Sub img����ȼ�_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl����_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl����(Index).Left + x, lbl����(Index).Top + y)
End Sub

Private Sub lbl����_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl�����ܶ�_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl�����ܶ�_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl�����ܶ�(Index).Left + x, lbl�����ܶ�(Index).Top + y)
End Sub

Private Sub lbl�����ܶ�_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl����_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl����(Index).Left + x, lbl����(Index).Top + y)
End Sub

Private Sub lbl����_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl��Ժ����_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl��Ժ����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl��Ժ����(Index).Left + x, lbl��Ժ����(Index).Top + y)
End Sub

Private Sub lbl��Ժ����_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl�Ա�_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl�Ա�_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl�Ա�(Index).Left + x, lbl�Ա�(Index).Top + y)
End Sub

Private Sub lbl�Ա�_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl����_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl����(Index).Left + x, lbl����(Index).Top + y)
End Sub

Private Sub lbl����_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lblҽʦ_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblҽʦ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblҽʦ(Index).Left + x, lblҽʦ(Index).Top + y)
End Sub

Private Sub lblҽʦ_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lbl���_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lbl���_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lbl���(Index).Left + x, lbl���(Index).Top + y)
End Sub

Private Sub lbl���_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lblסԺ��_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblסԺ��_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblסԺ��(Index).Left + x, lblסԺ��(Index).Top + y)
End Sub

Private Sub lblסԺ��_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub lblסԺ����_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblסԺ����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblסԺ����(Index).Left + x, lblסԺ����(Index).Top + y)
End Sub

Private Sub lblסԺ����_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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
    '��������������ģ��
    If Not LocatePatiRecord Then Exit Sub
    Call InNurseRoutine
End Sub

'54436:������,2012-10-10,�޸���ת�����������˺󣬲��ܹ��˳��޸�����ת�ƵĲ���
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
    If rptPati(PatiPage.Selected.Index).Columns.Count > c_��� Then rptPati(PatiPage.Selected.Index).Columns(c_���).Visible = False
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
    
    '�Ƿ�ˢ�����
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
            Case 0 '����
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            Case 1 'סԺ��
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case 2 '���￨
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
            Case 3 '����
        End Select
    End If
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean)
    Dim blnRefresh As Boolean
    Dim str���� As String, lng����ID As Long, lng��ҳID As Long, int���� As Integer, intPage As Integer
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
    '���Ҳ���
    With mrsPatiInfo
        If mintFindType = 0 Then '����
            .Filter = "����='" & UCase(txtFind.Text) & "'"
        End If
        If mintFindType = 1 Then 'סԺ��
            .Filter = "סԺ��=" & Val(txtFind.Text)
        End If
        If mintFindType = 2 Then '���￨
            .Filter = "���￨��='" & UCase(txtFind.Text) & "'"
        End If
        If mintFindType = 3 Then '����
            .Filter = "���� = '" & txtFind.Text & "'"
        End If
        If mintFindType = 4 Then '����
            .Filter = "���� Like '" & UCase(txtFind.Text) & "*'"
        End If
        If mintFindType = 4 Then
            mrsBedInfo.Filter = "���� Like '" & UCase(txtFind.Text) & "*' OR ���� Like '*," & UCase(txtFind.Text) & "*'"
            Call AdjustCard
            Exit Sub
        End If
        If .RecordCount = 0 Then
            .Filter = 0
            MsgBox "û���ҵ����������ļ�¼��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        str���� = !����
        lng����ID = !����ID
        lng��ҳID = !��ҳID
        int���� = !����
        strInput = !סԺ��
        .Filter = 0
    End With
    On Error GoTo errH
    '����������Ĳ��������ݿ����Ƿ����,������������ȡ��λ��
    'mstrSQL = " Select ��ǰ���� From ������Ϣ Where ��Ժ=1 And ����ID=[1] And ��ǰ����ID=[2]"
    '53907:������,2012-09-28,Ӧ�ü��ϲ�����ҳ�����ⲡ�����ζ���Ժ
    mstrSQL = " Select B.��Ժ���� ��ǰ���� From ������Ϣ A,������ҳ B Where A.����ID=B.����ID And B.����ID=[1] And B.��ҳID=[2] And B.��ǰ����ID=[3] And B.��Ժ���� IS NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ������Ϣ", lng����ID, lng��ҳID, CLng(Me.cboUnit.ItemData(Me.cboUnit.ListIndex)))
    If rsTemp.RecordCount <> 0 Then
        blnRefresh = (Nvl(rsTemp!��ǰ����, "") <> str����)
        txtFind.Text = Nvl(rsTemp!��ǰ����, "")
        mintFindType = 0
    Else
        If int���� = 5 Or int���� = 6 Or int���� = 7 Or int���� = 1 Or int���� = 0 Then
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
    mrsBedInfo.Filter = "����='" & str���� & "'"
    If mrsBedInfo.RecordCount > 0 Then
        If Val(Nvl(mrsBedInfo!����ID, 0)) = 0 Then
            mrsBedInfo.Filter = ""
            GoTo ErrNext
        End If
    Else
ErrNext:
        If int���� = 0 Or int���� = 1 Or int���� = 2 Then
            intPage = 0
        ElseIf int���� = 7 Then
            intPage = 1
        ElseIf int���� = 6 Or int���� = 5 Then
            intPage = 2
        ElseIf int���� Like "3*" Or (int���� = 4 And str���� = "") Then '��ͥ����
            intPage = 3
        End If
        PatiPage.Item(intPage).Selected = True
        
        For Each objRptRow In rptPati(intPage).Rows
            If Not objRptRow.Record Is Nothing Then
                If objRptRow.Record.Childs.Count = 0 Then
                    If IIf(Val(strInput) = 0, objRptRow.Record.Item(2).Value, objRptRow.Record.Item(5).Value) = IIf(Val(strInput) = 0, lng����ID, strInput) Then
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
'    If Not picPati(mrsBedInfo!��Ƭ����).Visible Then
'        mrsBedInfo.Filter = 0
'        MsgBox "���ҵ��ò��ˣ������ڸò��˲����Ϲ������������޸Ĺ������������²��ң�", vbInformation, gstrSysName
'        Exit Sub
'    End If
    
    Call AdjustCard
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt��������_GotFocus()
    mintREPORTSEL = -1
End Sub

Private Sub txtסԺ��_GotFocus()
    txtסԺ��.ForeColor = &HFF0000
    Call zlControl.TxtSelAll(txtסԺ��)
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    Dim strValue As String, strField As String
    Dim strInput As String, strSQL As String
    Dim objRptRow As ReportRow
    Dim rsTemp As New ADODB.Recordset
    Dim blnCard As Boolean, blnOk As Boolean
    Dim strFilter As String
    Dim blnExit As Boolean
    On Error GoTo ErrHand
    
    '49752,������,2012-09-05,��Ժ���˲����ṩ���ֲ��ҷ�ʽ
    txtסԺ��.ForeColor = &HFF0000
    If KeyAscii = 39 Then KeyAscii = 0
    '�Ƿ�ˢ�����
    blnCard = mintPatiInputType = 12 And KeyAscii <> 8 And Len(txtסԺ��.Text) = gbytCardLen - 1 And txtסԺ��.SelLength <> Len(txtסԺ��.Text)
    
    If KeyAscii = vbKeyReturn Or blnCard = True Then
        If KeyAscii <> 13 Then
            txtסԺ��.Text = txtסԺ��.Text & Chr(KeyAscii)
            txtסԺ��.SelStart = Len(txtסԺ��.Text)
        End If
        KeyAscii = 0
    Else
        Select Case mintPatiInputType
            Case 10 '����
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            Case 11 'סԺ��
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case 12 '���￨
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
            Case 13 '����
        End Select
        Exit Sub
    End If
    
    strInput = Trim(txtסԺ��.Text)
    If strInput = "" Then Exit Sub
   
    '�ڳ�Ժҳ���и��������סԺ�Ŷ�λ����
    blnExit = False
FindPati:
    blnOk = False
    For Each objRptRow In rptPati(Val(pic��Ժ����.Tag)).Rows
        If Not objRptRow.Record Is Nothing Then
            If objRptRow.Record.Childs.Count = 0 Then
                Select Case mintPatiInputType
                    Case 10 '����
                        If UCase(Trim(objRptRow.Record.Item(c_����).Value)) = UCase(strInput) Then blnOk = True
                    Case 11 'סԺ��
                        If Val(objRptRow.Record.Item(c_סԺ��).Value) = Val(strInput) Then blnOk = True
                    Case 12 '���￨
                        If UCase(objRptRow.Record.Item(c_���￨��).Value) = UCase(strInput) Then blnOk = True
                    Case Else
                        If objRptRow.Record.Item(c_����).Value = strInput Then blnOk = True
                End Select
                If blnOk = True Then
                    rptPati(Val(pic��Ժ����.Tag)).Rows(objRptRow.Index).Selected = True
                    rptPati(Val(pic��Ժ����.Tag)).SelectedRows(0).EnsureVisible
                    If rptPati(Val(pic��Ժ����.Tag)).Visible Then rptPati(Val(pic��Ժ����.Tag)).SetFocus
                    Exit Sub
                End If
            End If
        End If
    Next
    
    'ǿ��ѡ�����һ���������������ѭ��
    If blnExit = True And rptPati(Val(pic��Ժ����.Tag)).Rows.Count > 0 Then
        If Not rptPati(Val(pic��Ժ����.Tag)).Rows(rptPati(Val(pic��Ժ����.Tag)).Rows.Count - 1) Is Nothing Then
            rptPati(Val(pic��Ժ����.Tag)).Rows(rptPati(Val(pic��Ժ����.Tag)).Rows.Count - 1).Selected = True
            rptPati(Val(pic��Ժ����.Tag)).SelectedRows(0).EnsureVisible
            If rptPati(Val(pic��Ժ����.Tag)).Visible Then rptPati(Val(pic��Ժ����.Tag)).SetFocus
            Exit Sub
        End If
    End If
    If Val(pic��Ժ����.Tag) = ҳ��.��ͥ���� Or Val(pic��Ժ����.Tag) = ҳ��.����� Then Exit Sub
    
    '����Ҳ����ٴ����ݿ�����ȡ(��Ժ����ҳ���ṩ�˹���)
    '1--��֯SQL����
    strFilter = ""
    Select Case mintPatiInputType
        Case 10 '����
            strFilter = " And B.��Ժ����=[2] "
        Case 11 'סԺ��
            strFilter = " And B.סԺ��=[2] "
        Case 12 '���￨
            strFilter = " And A.���￨��=[2] "
        Case Else
            strFilter = " And A.����=[2] "
    End Select
    '61824:������,2013-05-23,��ʾ�����ֱ�־
    '2--��ʼ��ȡ����
    If pic��Ժ����.Tag = ҳ��.��Ժ Then
        strSQL = "" & _
            "Select /*+ RULE */ Decode(B.��Ժ��ʽ,'����',6,5) as ����," & _
            " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
            " Decode(B.��Ժ��ʽ,'����','��������','��Ժ����') as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(b.����,a.����) ����, NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " B.��Ժ���� AS ����,E.���� as ����ȼ�,B.�ѱ�,B.ҽ�Ƹ��ʽ,B.��ǰ����,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(b.·��״̬,-1) ·��״̬,trunc(b.��Ժ����)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID,A.��ҳId �����ҳId" & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z" & _
            " Where A.����ID=B.����ID And B.��������=Z.����(+) And Nvl(B.��ҳID,0)<>0 And B.״̬=0" & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And B.��ǰ����ID=[1] " & strFilter & " And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And B.��Ժ���� Is Not NULL And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
    ElseIf pic��Ժ����.Tag = ҳ��.ת�� Then
         strSQL = "" & _
            "Select  Distinct 7 as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,'ת������' as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,D.���� as ����,C.����ID,C.����ҽʦ as סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " C.����,E.���� as ����ȼ�,B.�ѱ�,B.ҽ�Ƹ��ʽ,B.��ǰ����,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) As ��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID,A.��ҳId �����ҳId" & _
            " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D,�շ���ĿĿ¼ E,�������� Z" & _
            " Where A.����ID=B.����ID And B.��������=Z.����(+) And Nvl(B.��ҳID,0)<>0 And B.����ȼ�ID=E.ID(+)" & _
            " And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
            " And B.��ǰ����ID<>[1] And C.����ID+0=[1] And C.����ID=D.ID " & strFilter & _
            " And Nvl(C.���Ӵ�λ,0)=0 And C.��ֹԭ�� In(3,15) And C.��ֹʱ�� is Not Null " & _
            " And Nvl(B.״̬,0)<>2 And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
    End If
    Set rsTemp = New ADODB.Recordset

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), strInput)
    Call UpgradeList(rsTemp)

    '׷�Ӽ�¼��
    mrsPatiInfo.Filter = 0
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
        strField = "����|����2|����|����ID|��ҳID|סԺ��|����|�Ա�|����|����|����ID|סԺҽʦ|���λ�ʿ|����״̬|����|����ȼ�|�ѱ�|ҽ�Ƹ��ʽ|��ǰ����|��Ժ����|��Ժ����|סԺ����|��Ժ��ʽ|��������|״̬|����|���￨��|·��״̬|��ɫ|������|Ӥ������ID|Ӥ������ID|�����ҳId"
        Do While Not rsTemp.EOF
            strValue = rsTemp!���� & "|" & Nvl(rsTemp!����2, 0) & "|" & Nvl(rsTemp!����) & "|" & Nvl(rsTemp!����ID, 0) & "|" & Nvl(rsTemp!��ҳID, 0) & "|" & Nvl(rsTemp!סԺ��, 0) & "|" & Nvl(rsTemp!����) & "|" & Nvl(rsTemp!�Ա�) & "|" & _
                      Nvl(rsTemp!����) & "|" & Nvl(rsTemp!����) & "|" & Nvl(rsTemp!����ID, 0) & "|" & Nvl(rsTemp!סԺҽʦ) & "|" & Nvl(rsTemp!���λ�ʿ) & "|" & Nvl(rsTemp!����״̬, 0) & "|" & Nvl(rsTemp!����) & "|" & _
                      Nvl(rsTemp!����ȼ�, "����") & "|" & Nvl(rsTemp!�ѱ�) & "|" & Nvl(rsTemp!ҽ�Ƹ��ʽ) & "|" & Nvl(rsTemp!��ǰ����, "һ��") & "|" & Nvl(rsTemp!��Ժ����) & "|" & Nvl(rsTemp!��Ժ����) & "|" & Nvl(rsTemp!סԺ����) & "|" & Nvl(rsTemp!��Ժ��ʽ) & "|" & _
                      Nvl(rsTemp!��������, "��ͨ����") & "|" & Nvl(rsTemp!״̬, 0) & "|" & Nvl(rsTemp!����, 0) & "|" & Nvl(rsTemp!���￨��) & "|" & Nvl(rsTemp!·��״̬, 0) & "|" & Nvl(rsTemp!��ɫ, 0) & "|" & Nvl(rsTemp!������) & "|" & Nvl(rsTemp!Ӥ������ID, 0) & "|" & Nvl(rsTemp!Ӥ������ID, 0) & "|" & Nvl(rsTemp!�����ҳID, 0)
            Call Record_Add(mrsPatiInfo, strField, strValue)
            rsTemp.MoveNext
        Loop
        blnExit = True
        GoTo FindPati
    Else
        MsgBox "û���ҵ����������ļ�¼��", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txtסԺ��_LostFocus()
    txtסԺ��.Text = ""
    txtסԺ��.ForeColor = &HC0C0C0
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'���ܣ�������ӡ�¼���д����ҳ��ӡ����
    Dim strSQL As String
    
    strSQL = _
            "Zl_���Ӳ�����ӡ_Insert(Null,9," & mlng����ID & "," & mPatiInfo.��ҳID & ",'" & UserInfo.���� & "')"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitColor()
    Dim strValue As String
    Dim lng�ؼ� As Long, lngһ�� As Long, lng���� As Long, lng���� As Long
    Const c��ɫ As Long = 8388736
    Const c��ɫ As Long = 255
    Const c��ɫ As Long = 16711680
    Const c��ɫ As Long = 16777215
    
    Call DeleteFile
    mintIndex = 0
    imgHLDJ(0).ListImages.Clear
    imgHLDJ(999).ListImages.Clear
    '��ȡ����ȼ���������(����ȡȱʡ����)
    strValue = zlDatabase.GetPara("�ؼ�������ɫ", glngSys, 1265, "")
    lng�ؼ� = IIf(strValue = "", c��ɫ, Val(strValue))
    strValue = zlDatabase.GetPara("һ��������ɫ", glngSys, 1265, "")
    lngһ�� = IIf(strValue = "", c��ɫ, Val(strValue))
    strValue = zlDatabase.GetPara("����������ɫ", glngSys, 1265, "")
    lng���� = IIf(strValue = "", c��ɫ, Val(strValue))
    strValue = zlDatabase.GetPara("����������ɫ", glngSys, 1265, "")
    lng���� = IIf(strValue = "", c��ɫ, Val(strValue))
    
    '��ͼ
    mlngColor = lng�ؼ�
    Call DrawPoly
    mlngColor = lngһ��
    Call DrawPoly
    mlngColor = lng����
    Call DrawPoly
    mlngColor = lng����
    Call DrawPoly
End Sub

Private Sub AddColor()
    Dim strFile As String
    mintIndex = mintIndex + 1
    '������Ϊ�ļ�,���������ͼƬʱ,���뵽imagelist���ʼ��ֻ�����һ��,Ӧ��������image�б������ͼƬID���
    
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

    '������򲢻�����
    ReDim PtInPoly(4) As POINTAPI
    PtInPoly(1).x = 0
    PtInPoly(1).y = 0
    PtInPoly(2).x = picHLDJ.ScaleWidth
    PtInPoly(2).y = 0
    PtInPoly(3).x = picHLDJ.ScaleWidth
    PtInPoly(3).y = picHLDJ.ScaleHeight
    PtInPoly(4).x = PtInPoly(1).x
    PtInPoly(4).y = PtInPoly(1).y
    
    '����ϵͳˢ��
    picHLDJ.Cls
    lngBrush = CreateSolidBrush(mlngColor)

    '�������ˢ�ӳɹ�,��ѡ��
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
'���ܣ����в�����ҳ����
'������blnEditable=�Ƿ�����༭(��Ȩ�޼�ǩ������������)
    Dim blnReadOnly As Boolean
    
    If mPatiInfo.����ת�� Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '������Ŀ֮�󲻿�������
    If Not (CheckMecRed(mrsPatiInfo.Fields("����ID").Value, mrsPatiInfo.Fields("��ҳID").Value, Me.Caption) Or blnEditable) Then
        blnReadOnly = True
    End If
    
    '��ģ̬��ʾ��ҳ����
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, P�°滤ʿվ, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    If Not mclsInOutMedRec.IsOpen Then
        Call mclsInOutMedRec.ShowInMedRecEdit(Me, mrsPatiInfo.Fields("����ID").Value, mrsPatiInfo.Fields("��ҳID").Value, mrsPatiInfo.Fields("����ID").Value, mrsPatiInfo.Fields("·��״̬").Value, , mstrPrivs, IIf(blnReadOnly, 1, 0), False)
    End If
End Sub


Private Function CheckBabyInOut() As Boolean
'���ܣ����Ӥ����ĸ���Ƿ���룬�е�ǰ��Ӥ������
    If Val(Nvl(mrsPatiInfo.Fields("Ӥ������ID").Value)) <> 0 Then
        If Val(Nvl(mrsPatiInfo.Fields("Ӥ������ID").Value)) = cboUnit.ItemData(cboUnit.ListIndex) And mintREPORTSEL = -1 Then
            MsgBox "�ò����Ѿ�ת���������ˣ�ֻ��Ӥ�����ڱ����ң�������������ˡ�", vbInformation, Me.Caption
            CheckBabyInOut = True
        End If
    End If
End Function

Private Function GetPatiCount(ByVal Index As Integer) As Long
'����:��ȡ���ڴ�������Ŀ(���ڲ����б�����˷���Records.Countͳ�Ƴ�������������Ŀ,�˴���Ҫ����ͳ��)
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
'���ܣ���֯�˵������ؼ�¼���У�ע����ϰ汾�ļ��ݴ���
'������strFunc �ϰ汾�����д���strXML��������Ϣ�Ĺ��ܴ�
    Dim strM As String
    Dim strB As String
    Dim strP As String
    Dim strTag As String
    Dim i As Long
    Dim strTmp As String
    Dim lngS As Long, lngE As Long
    
    If strXML = "" And strFunc <> "" Then
        '������ǰ�ϰ汾�ķ�ʽ
        Call InitPlugInRsBar(rsBar)
        Call AddPlugInBarRs(rsBar, strFunc, 1)
        Call AddPlugInBarRs(rsBar, strFunc, 2)
        Call AddPlugInBarRs(rsBar, strFunc, 3)
        Call SetPlugInBar(rsBar, 1)
        Exit Sub
    End If
    
    On Error GoTo errH
    strXML = Trim(strXML)
    '�ݶ�Ϊ200����չ���ܲ������ֹ��ѭ��
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
'���ܣ������ܴ�ת��Ϊ��¼����ʽ
'������strFunc ���ܴ���intType ���ܰ�ť������һ�� 1-�˵�����2-��������3-�����
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
                    '��һ��������ť��ʾ�ָ���
                    blnFirstTool = True
                    !IsGroup = 1
                End If
            End If
            !������ = strFuncName
            !�˵��� = strFuncName
            .Update
        Next
    End With
End Sub

Private Function SetPlugInBar(ByRef rsBar As ADODB.Recordset, ByVal lngV As Long) As String
'���ܣ����书��ID���Ӳ˵����
'������lngV �汾��1-�ϰ棬2-�°�
'���أ��ַ�������ǰ�Ͱ汾��ʽ�Ĺ��ܴ�
    Dim i As Long
    '���书��ID��ͼ��ID
    With rsBar
        .Filter = 0
        If .EOF Then Exit Function
        .MoveFirst
        For i = 1 To .RecordCount
            !��� = i
            !����ID = conMenu_Tool_PlugIn_Item + i
            !ͼ��ID = conMenu_Tool_PlugIn_Item
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
'���ܣ��趨���
'������lngV �汾��1-�ϰ棬2-�°� intType ���ܰ�ť������һ�� 1-�˵�����2-��������3-�����
    Dim i As Long
    With rsBar
        .Filter = "IsInTool=0 and BarType=" & intType
        If .RecordCount = 1 And lngV = 2 Then
            '���ֻ��һ����Ҳ��Ϊ������ť
            !IsInTool = 1
            .Update
        Else
            For i = 1 To .RecordCount
                If i <= 35 Then
                    If i <= 9 Then
                        !�˵��� = !�˵��� & "(&" & i & ")"
                    Else
                        !�˵��� = !�˵��� & "(&" & Chr(55 + i) & ")"
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
                    !�˵��� = !�˵��� & "(&" & i & ")"
                Else
                    !�˵��� = !�˵��� & "(&" & Chr(55 + i) & ")"
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
    rsBar.Fields.Append "���", adBigInt '��������
    rsBar.Fields.Append "����ID", adBigInt '�˵���ť Control.ID
    rsBar.Fields.Append "ͼ��ID", adBigInt
    rsBar.Fields.Append "������", adVarChar, 1000 'ȥ���ؼ���֮��� ���� ���������ϵİ�ť����
    rsBar.Fields.Append "�˵���", adVarChar, 1000 '�˵���/�Ҽ��˵� ����
    rsBar.Fields.Append "IsAuto", adInteger '�Ƿ��Զ�ִ�й���
    rsBar.Fields.Append "IsGroup", adInteger '�Ƿ�ָ���
    rsBar.Fields.Append "IsInTool", adInteger '�Ƿ������ʾ
    rsBar.Fields.Append "BarType", adInteger '1-�˵�����2����������3��������
    rsBar.CursorLocation = adUseClient
    rsBar.LockType = adLockOptimistic
    rsBar.CursorType = adOpenStatic
    rsBar.Open
End Sub
